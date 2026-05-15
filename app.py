import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import base64
import json
import os
import re
import time
from io import StringIO
import requests

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]
CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
ZOOM_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "zoom_config.json")
ZOOM_TOKEN_URL = "https://zoom.us/oauth/token"
ZOOM_API_BASE = "https://api.zoom.us/v2"

st.set_page_config(page_title="Zoom Attendance Marker", layout="wide", page_icon="📋")

st.markdown("""
<style>
    .stApp { max-width: 1000px; margin: 0 auto; }
    div[data-testid="stMetric"] {
        background: #f0f2f6; border-radius: 10px; padding: 15px;
    }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# GOOGLE SHEETS CONNECTION
# ──────────────────────────────────────────────

@st.cache_resource
def get_gspread_client():
    creds_dict = None

    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
    except (KeyError, FileNotFoundError):
        pass

    if creds_dict is None and os.path.exists(CREDS_FILE):
        with open(CREDS_FILE, "r") as f:
            creds_dict = json.load(f)

    if creds_dict is None:
        raise ValueError(
            "No credentials found! "
            "Place your service account JSON as 'credentials.json' next to app.py."
        )

    credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(credentials)


# ──────────────────────────────────────────────
# ZOOM (Server-to-Server OAuth + Reports API)
# ──────────────────────────────────────────────

def load_zoom_config() -> dict | None:
    cfg = None
    try:
        z = dict(st.secrets["zoom"])
        if z.get("account_id") and z.get("client_id") and z.get("client_secret"):
            cfg = {
                "account_id": str(z["account_id"]).strip(),
                "client_id": str(z["client_id"]).strip(),
                "client_secret": str(z["client_secret"]).strip(),
            }
    except (KeyError, FileNotFoundError, TypeError):
        pass

    if cfg is None and os.path.exists(ZOOM_CONFIG_FILE):
        with open(ZOOM_CONFIG_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        if raw.get("account_id") and raw.get("client_id") and raw.get("client_secret"):
            cfg = {
                "account_id": str(raw["account_id"]).strip(),
                "client_id": str(raw["client_id"]).strip(),
                "client_secret": str(raw["client_secret"]).strip(),
            }

    if cfg is None:
        aid = os.environ.get("ZOOM_ACCOUNT_ID", "").strip()
        cid = os.environ.get("ZOOM_CLIENT_ID", "").strip()
        csec = os.environ.get("ZOOM_CLIENT_SECRET", "").strip()
        if aid and cid and csec:
            cfg = {"account_id": aid, "client_id": cid, "client_secret": csec}

    return cfg


def get_zoom_access_token(config: dict) -> str:
    now = time.time()
    exp = float(st.session_state.get("_zoom_token_expires") or 0)
    tok = st.session_state.get("_zoom_token")
    if tok and exp > now + 60:
        return tok

    basic = base64.b64encode(
        f"{config['client_id']}:{config['client_secret']}".encode()
    ).decode()
    resp = requests.post(
        ZOOM_TOKEN_URL,
        headers={
            "Authorization": f"Basic {basic}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data={
            "grant_type": "account_credentials",
            "account_id": config["account_id"],
        },
        timeout=45,
    )
    try:
        resp.raise_for_status()
    except requests.HTTPError as e:
        detail = ""
        try:
            detail = resp.text[:500]
        except Exception:
            pass
        raise RuntimeError(f"Zoom OAuth failed ({resp.status_code}): {detail}") from e

    js = resp.json()
    st.session_state["_zoom_token"] = js["access_token"]
    st.session_state["_zoom_token_expires"] = now + float(js.get("expires_in", 3600))
    return st.session_state["_zoom_token"]


def fetch_meeting_participants_report(meeting_id: str, token: str) -> pd.DataFrame:
    """Uses Reports API — requires `report:read:admin` (or equivalent) on the S2S app."""
    rows: list[dict] = []
    next_token = ""
    meeting_id = meeting_id.strip()
    while True:
        params: dict = {"page_size": 300}
        if next_token:
            params["next_page_token"] = next_token
        url = f"{ZOOM_API_BASE}/report/meetings/{meeting_id}/participants"
        r = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=60,
        )
        if r.status_code == 401:
            st.session_state.pop("_zoom_token", None)
            st.session_state.pop("_zoom_token_expires", None)
        r.raise_for_status()
        data = r.json()
        for p in data.get("participants") or []:
            email = (p.get("user_email") or p.get("email") or "").strip()
            rows.append(
                {
                    "Name": (p.get("name") or "").strip(),
                    "User Email": email,
                    "Duration (minutes)": p.get("duration"),
                    "Join Time": p.get("join_time"),
                    "Leave Time": p.get("leave_time"),
                }
            )
        next_token = (data.get("next_page_token") or "").strip()
        if not next_token:
            break
    return pd.DataFrame(rows)


def test_zoom_api_connection(config: dict) -> None:
    """Validates Server-to-Server OAuth (client credentials + token exchange)."""
    get_zoom_access_token(config)


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def parse_zoom_report(uploaded_file) -> pd.DataFrame:
    content = uploaded_file.getvalue().decode("utf-8", errors="replace")
    lines = content.strip().split("\n")

    header_keywords = ["name", "email", "user email", "participant", "join time", "duration"]
    header_idx = 0
    for i, line in enumerate(lines):
        lower = line.lower()
        if sum(1 for kw in header_keywords if kw in lower) >= 2:
            header_idx = i
            break

    csv_text = "\n".join(lines[header_idx:])
    df = pd.read_csv(StringIO(csv_text))
    df.columns = df.columns.str.strip()
    return df


def extract_zoom_emails(zoom_df: pd.DataFrame) -> set[str]:
    for col in zoom_df.columns:
        if "email" in col.lower():
            emails = zoom_df[col].dropna().astype(str).str.strip().str.lower()
            return set(emails) - {""}
    return set()


def extract_zoom_names(zoom_df: pd.DataFrame) -> set[str]:
    for col in zoom_df.columns:
        if "name" in col.lower():
            names = zoom_df[col].dropna().astype(str).str.strip().str.lower()
            return set(names) - {""}
    return set()


def normalize(val) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"\s+", " ", val.strip().lower())


def match_with_fallback(
    sheet_data: pd.DataFrame,
    email_col: str,
    name_col: str,
    zoom_emails: set[str],
    zoom_names: set[str],
) -> pd.DataFrame:
    """
    Match logic:
      1. Try email match first
      2. For unmatched rows, try name match as fallback
      3. Tag each row with match method or 'Unmatched'
    """
    results = []

    for _, row in sheet_data.iterrows():
        email = normalize(str(row.get(email_col, "")))
        name = normalize(str(row.get(name_col, "")))

        if email and email in zoom_emails:
            results.append("✅ Matched by Email")
        elif name and name in zoom_names:
            results.append("⚠️ Matched by Name (fallback)")
        else:
            results.append("❌ Unmatched")

    sheet_data = sheet_data.copy()
    sheet_data["_match_status"] = results
    return sheet_data


# ══════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════

st.title("📋 Zoom Attendance → Google Sheet Marker")

zoom_config = load_zoom_config()

# ── Connect: Google Sheets (optional — Zoom still loads if this fails) ──
gc = None
try:
    gc = get_gspread_client()
except Exception as e:
    st.error(f"❌ Google Sheets not available: {e}")
    st.info(
        "Fix **`credentials.json`**: download the **full** service account JSON from Google Cloud "
        "(IAM → Service Accounts → **Keys** → **Add key** → JSON). "
        "Empty or truncated `private_key` causes PEM / MalformedFraming errors. **Section 1 Zoom still works.**"
    )

sheet_ok = "✅ Google Sheets connected" if gc is not None else "⚠️ Google Sheets offline"
zoom_status = "✅ Zoom API configured" if zoom_config else "⚠️ Zoom API not configured (CSV only)"
st.caption(
    f"{sheet_ok}  •  {zoom_status}  •  Load participants from Zoom or upload CSV → match → mark attendance"
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 1: Zoom (API or CSV)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.header("1️⃣ Zoom participants")

with st.expander("Zoom API (Server-to-Server OAuth)", expanded=bool(zoom_config)):
    if not zoom_config:
        st.markdown(
            "Add **Server-to-Server OAuth** credentials so this app can call Zoom directly:\n\n"
            "1. [Zoom Marketplace](https://marketplace.zoom.us/) → **Develop** → **Build App** → **Server-to-Server OAuth**.\n"
            "2. Copy **Account ID**, **Client ID**, **Client Secret**.\n"
            "3. Under **Scopes**, add **report:read:admin** (needed for participant reports by meeting ID).\n"
            "4. Put them in `.streamlit/secrets.toml` under `[zoom]`, or create `zoom_config.json` next to `app.py`, "
            "or set env vars `ZOOM_ACCOUNT_ID`, `ZOOM_CLIENT_ID`, `ZOOM_CLIENT_SECRET`.\n\n"
            "`[zoom]` example:\n```toml\n[zoom]\naccount_id = \"...\"\nclient_id = \"...\"\nclient_secret = \"...\"\n```"
        )
    else:
        st.success("Zoom credentials found — use **Test connection**, then **Meeting ID** + **Load from Zoom**.")
        tc1, tc2 = st.columns([1, 2])
        with tc1:
            if st.button("Test Zoom connection", key="zoom_test"):
                try:
                    with st.spinner("Calling Zoom API..."):
                        test_zoom_api_connection(zoom_config)
                    st.session_state["_zoom_api_ok"] = True
                    st.success("Zoom API responded OK.")
                except Exception as e:
                    st.session_state["_zoom_api_ok"] = False
                    st.error(f"Connection failed: {e}")
        with tc2:
            if st.session_state.get("_zoom_api_ok"):
                st.caption("Last test succeeded this session.")

        zc1, zc2, zc3 = st.columns([2, 1, 1])
        with zc1:
            meeting_id_input = st.text_input(
                "Meeting ID (numeric, from Zoom)",
                placeholder="e.g. 1234567890",
                key="zoom_meeting_id",
                help="Past meetings: use the same ID shown in Zoom reports / meeting details.",
            )
        with zc2:
            load_api = st.button("Load from Zoom", type="primary", key="zoom_load_api")
        with zc3:
            if st.button("Clear API load", key="zoom_clear_api"):
                st.session_state.pop("zoom_api_df", None)
                st.session_state.pop("zoom_api_meeting_id", None)
                st.rerun()

        if load_api and meeting_id_input.strip():
            try:
                with st.spinner("Fetching participant report from Zoom..."):
                    token = get_zoom_access_token(zoom_config)
                    api_df = fetch_meeting_participants_report(meeting_id_input.strip(), token)
                st.session_state["zoom_api_df"] = api_df
                st.session_state["zoom_api_meeting_id"] = meeting_id_input.strip()
                st.success(f"Loaded **{len(api_df)}** participant row(s) for meeting `{meeting_id_input.strip()}`.")
            except Exception as e:
                st.error(f"Zoom API error: {e}")
                st.caption(
                    "Typical fixes: correct **Meeting ID**, meeting already ended and report available, "
                    "and S2S app has **report:read:admin** scope."
                )

        if st.session_state.get("zoom_api_df") is not None and not st.session_state["zoom_api_df"].empty:
            mid = st.session_state.get("zoom_api_meeting_id", "")
            st.info(f"Using Zoom API data for meeting **{mid}** — merged with any CSV you upload below.")
            with st.expander("Preview Zoom API participants"):
                st.dataframe(st.session_state["zoom_api_df"].head(30), use_container_width=True)

st.subheader("Or upload Zoom participant CSV")
zoom_files = st.file_uploader(
    "Upload Zoom CSV(s)", type=["csv"], accept_multiple_files=True, key="zoom"
)

zoom_df_parts: list[pd.DataFrame] = []
if st.session_state.get("zoom_api_df") is not None and not st.session_state["zoom_api_df"].empty:
    zoom_df_parts.append(st.session_state["zoom_api_df"].copy())

if zoom_files:
    for f in zoom_files:
        try:
            zdf = parse_zoom_report(f)
            zoom_df_parts.append(zdf)
            st.success(f"✅ `{f.name}` — {len(zdf)} participants")
        except Exception as e:
            st.error(f"❌ Failed to parse `{f.name}`: {e}")

zoom_emails: set[str] = set()
zoom_names: set[str] = set()
zoom_df = pd.DataFrame()

if zoom_df_parts:
    zoom_df = pd.concat(zoom_df_parts, ignore_index=True)
    zoom_emails = extract_zoom_emails(zoom_df)
    zoom_names = extract_zoom_names(zoom_df)

    if zoom_emails or zoom_names:
        st.info(
            f"📧 Extracted **{len(zoom_emails)}** unique emails and **{len(zoom_names)}** unique names "
            f"from **{len(zoom_df)}** participant row(s)."
        )
        with st.expander("Preview combined Zoom data"):
            st.dataframe(zoom_df.head(20), use_container_width=True)
    else:
        st.error("❌ No email or name columns found in the Zoom data (API or CSV).")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 2: Google Sheet Details
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if zoom_emails or zoom_names:
    st.header("2️⃣ Google Sheet Details")

    if gc is None:
        st.warning(
            "**Google Sheets is not connected** — Step 2–3 are hidden until `credentials.json` is valid. "
            "In Google Cloud: Service account → Keys → Add key → JSON (full file). **Zoom in section 1 still works.**"
        )
    else:
        sheet_url = st.text_input(
            "🔗 Google Sheet URL",
            placeholder="https://docs.google.com/spreadsheets/d/xxxxx/edit",
        )
    
        col1, col2 = st.columns(2)
        with col1:
            tab_name = st.text_input("📄 Tab name", value="Sheet1")
        with col2:
            email_col_name = st.text_input("📧 Email column name (primary match)", value="Email")
    
        col3, col4, col5 = st.columns(3)
        with col3:
            name_col_name = st.text_input("👤 Name column name (fallback match)", value="Name")
        with col4:
            attendance_col_name = st.text_input("✅ Attendance column (to update)", value="Attendance")
        with col5:
            update_value = st.text_input("📝 Value to mark", value="Present")
    
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # STEP 3: Preview & Mark
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        ws = None
        sheet_data = None
    
        if sheet_url and tab_name and email_col_name and name_col_name and attendance_col_name and update_value:
            try:
                if "docs.google.com" in sheet_url:
                    spreadsheet = gc.open_by_url(sheet_url)
                else:
                    spreadsheet = gc.open_by_key(sheet_url.strip())
    
                ws = spreadsheet.worksheet(tab_name)
    
                # Use get_all_values to handle duplicate/empty headers
                all_values = ws.get_all_values()
                if len(all_values) < 2:
                    st.warning("⚠️ Worksheet is empty or has no data rows.")
                    sheet_data = pd.DataFrame()
                else:
                    headers = all_values[0]
                    # Make headers unique: append _2, _3 etc. for duplicates, name empty cols
                    seen = {}
                    unique_headers = []
                    for i, h in enumerate(headers):
                        h = h.strip()
                        if not h:
                            h = f"_unnamed_{i}"
                        if h in seen:
                            seen[h] += 1
                            unique_headers.append(f"{h}_{seen[h]}")
                        else:
                            seen[h] = 1
                            unique_headers.append(h)
                    sheet_data = pd.DataFrame(all_values[1:], columns=unique_headers)
    
                if sheet_data.empty:
                    st.warning("⚠️ Worksheet is empty or has no header row.")
                else:
                    sheet_cols = sheet_data.columns.tolist()
                    valid = True
    
                    if email_col_name not in sheet_cols:
                        st.error(f"❌ Email column `{email_col_name}` not found. Available: `{'`, `'.join(sheet_cols)}`")
                        valid = False
                    if name_col_name not in sheet_cols:
                        st.error(f"❌ Name column `{name_col_name}` not found. Available: `{'`, `'.join(sheet_cols)}`")
                        valid = False
    
                    if not valid:
                        sheet_data = None
                    else:
                        st.success(f"✅ Loaded **{len(sheet_data)}** rows from `{tab_name}`")
                        if attendance_col_name not in sheet_cols:
                            st.warning(f"⚠️ Column `{attendance_col_name}` doesn't exist — will be created automatically.")
    
            except gspread.exceptions.WorksheetNotFound:
                st.error(f"❌ Tab `{tab_name}` not found.")
            except Exception as e:
                st.error(f"❌ Could not load sheet: {e}")
    
        if sheet_data is not None and not sheet_data.empty and ws:
            st.header("3️⃣ Preview & Mark Attendance")
    
            # Run matching with fallback
            matched_data = match_with_fallback(
                sheet_data, email_col_name, name_col_name, zoom_emails, zoom_names
            )
    
            email_matched = (matched_data["_match_status"] == "✅ Matched by Email").sum()
            name_matched = (matched_data["_match_status"] == "⚠️ Matched by Name (fallback)").sum()
            unmatched = (matched_data["_match_status"] == "❌ Unmatched").sum()
            total = len(matched_data)
            total_matched = email_matched + name_matched
    
            # ── Count Zoom attendees not in sheet ──
            sheet_emails_set_quick = set(sheet_data[email_col_name].astype(str).apply(normalize).tolist()) - {""}
            sheet_names_set_quick = set(sheet_data[name_col_name].astype(str).apply(normalize).tolist()) - {""}
            zoom_only_count = 0
            for _, row in zoom_df.iterrows():
                z_email = None
                z_name = None
                for col in zoom_df.columns:
                    if "email" in col.lower():
                        z_email = normalize(str(row[col]))
                    if "name" in col.lower():
                        z_name = normalize(str(row[col]))
                if not (z_email and z_email in sheet_emails_set_quick) and not (z_name and z_name in sheet_names_set_quick):
                    zoom_only_count += 1
    
            # ── Metrics ──
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total in Sheet", total)
            m2.metric("Matched by Email", email_matched)
            m3.metric("Matched by Name", name_matched)
            m4.metric("❌ Unmatched", unmatched)
            m5.metric("🆕 In Zoom only", zoom_only_count)
    
            # ── Full Preview ──
            display_cols = [name_col_name, email_col_name, "_match_status"]
            preview = matched_data[display_cols].rename(columns={"_match_status": "Match Status"})
    
            st.subheader("Full Match Report")
            st.dataframe(
                preview.style.map(
                    lambda v: (
                        "background-color: #d4edda" if "Email" in str(v)
                        else "background-color: #fff3cd" if "Name" in str(v)
                        else "background-color: #f8d7da" if "Unmatched" in str(v)
                        else ""
                    ),
                    subset=["Match Status"],
                ),
                use_container_width=True,
                height=400,
            )
    
            # ── Unmatched Report ──
            if unmatched > 0:
                st.subheader(f"⚠️ Unmatched Participants ({unmatched})")
                st.caption("These participants could NOT be matched by email or name.")
                unmatched_df = matched_data[matched_data["_match_status"] == "❌ Unmatched"]
                st.dataframe(
                    unmatched_df[[name_col_name, email_col_name]],
                    use_container_width=True,
                )
    
            # ── Zoom attendees NOT in sheet ──
            sheet_emails_set = set(sheet_data[email_col_name].astype(str).apply(normalize).tolist()) - {""}
            sheet_names_set = set(sheet_data[name_col_name].astype(str).apply(normalize).tolist()) - {""}
    
            zoom_not_in_sheet = []
            for _, row in zoom_df.iterrows():
                z_email = None
                z_name = None
                for col in zoom_df.columns:
                    if "email" in col.lower():
                        z_email = normalize(str(row[col]))
                    if "name" in col.lower():
                        z_name = normalize(str(row[col]))
    
                email_found = z_email and z_email in sheet_emails_set
                name_found = z_name and z_name in sheet_names_set
    
                if not email_found and not name_found:
                    zoom_not_in_sheet.append({
                        "Name": row.get(next((c for c in zoom_df.columns if "name" in c.lower()), ""), ""),
                        "Email": row.get(next((c for c in zoom_df.columns if "email" in c.lower()), ""), ""),
                    })
    
            if zoom_not_in_sheet:
                not_in_sheet_df = pd.DataFrame(zoom_not_in_sheet).drop_duplicates()
                st.subheader(f"🆕 In Zoom Meeting but NOT in Sheet ({len(not_in_sheet_df)})")
                st.caption("These people attended the Zoom meeting but don't exist in your Google Sheet.")
                st.dataframe(not_in_sheet_df, use_container_width=True)
    
            # ── Name-fallback detail ──
            if name_matched > 0:
                st.subheader(f"ℹ️ Matched by Name — Verify ({name_matched})")
                st.caption("These matched by name only (email didn't match). Please verify they're correct.")
                name_df = matched_data[matched_data["_match_status"] == "⚠️ Matched by Name (fallback)"]
                st.dataframe(
                    name_df[[name_col_name, email_col_name]],
                    use_container_width=True,
                )
    
            st.divider()
    
            # ── Mark Button ──
            if total_matched > 0:
                if st.button(
                    f"✅ Mark {total_matched} participants as `{update_value}` ({email_matched} by email + {name_matched} by name)",
                    type="primary",
                    use_container_width=True,
                ):
                    with st.spinner("Updating Google Sheet..."):
                        try:
                            header_row = ws.row_values(1)
    
                            if attendance_col_name in header_row:
                                att_col_idx = header_row.index(attendance_col_name) + 1
                            else:
                                att_col_idx = len(header_row) + 1
                                ws.update_cell(1, att_col_idx, attendance_col_name)
                                st.info(f"📌 Created column `{attendance_col_name}`.")
    
                            cells_to_update = []
                            for i, status in enumerate(matched_data["_match_status"]):
                                if "Matched" in status:
                                    cells_to_update.append(gspread.Cell(i + 2, att_col_idx, update_value))
    
                            if cells_to_update:
                                ws.update_cells(cells_to_update)
    
                            st.success(
                                f"✅ Done! Marked **{total_matched}** participants as `{update_value}` "
                                f"({email_matched} by email, {name_matched} by name)."
                            )
                            if unmatched > 0:
                                st.warning(f"⚠️ {unmatched} participants remain unmatched — see report above.")
                            st.balloons()
    
                        except Exception as e:
                            st.error(f"❌ Failed: {e}")
                            st.exception(e)
            else:
                st.error("❌ No matches found. Check column names and data.")

# ── Footer ──
st.divider()
st.caption("Built for Iron Lady 🔧")
