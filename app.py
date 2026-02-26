import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import os
import re
from io import StringIO

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]
CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")

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

# ── Connect ──
try:
    gc = get_gspread_client()
except Exception as e:
    st.error(f"❌ Google Sheets connection failed: {e}")
    st.info("Place your service account JSON as `credentials.json` next to `app.py`.")
    st.stop()

st.caption("✅ Google Sheets connected  •  Upload Zoom reports → match emails → mark attendance")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 1: Upload Zoom Reports
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.header("1️⃣ Upload Zoom Participant Report(s)")

zoom_files = st.file_uploader(
    "Upload Zoom CSV(s)", type=["csv"], accept_multiple_files=True, key="zoom"
)

zoom_emails = set()
zoom_names = set()

if zoom_files:
    all_zoom_dfs = []
    for f in zoom_files:
        try:
            zdf = parse_zoom_report(f)
            all_zoom_dfs.append(zdf)
            st.success(f"✅ `{f.name}` — {len(zdf)} participants")
        except Exception as e:
            st.error(f"❌ Failed to parse `{f.name}`: {e}")

    if all_zoom_dfs:
        zoom_df = pd.concat(all_zoom_dfs, ignore_index=True)
        zoom_emails = extract_zoom_emails(zoom_df)
        zoom_names = extract_zoom_names(zoom_df)

        if zoom_emails or zoom_names:
            st.info(f"📧 Extracted **{len(zoom_emails)}** emails and **{len(zoom_names)}** names from Zoom report(s).")
            with st.expander("Preview Zoom data"):
                st.dataframe(zoom_df.head(20), use_container_width=True)
        else:
            st.error("❌ No email or name columns found in the Zoom CSV.")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 2: Google Sheet Details
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if zoom_emails or zoom_names:
    st.header("2️⃣ Google Sheet Details")

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
