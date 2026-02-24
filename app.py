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

# Path to your service account JSON file (keep it in the same folder as app.py)
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
    """
    Connect to Google Sheets.
    Priority: st.secrets (Streamlit Cloud) → credentials.json file (local)
    """
    creds_dict = None

    # 1. Try Streamlit Cloud secrets
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
    except (KeyError, FileNotFoundError):
        pass

    # 2. Fall back to credentials.json file
    if creds_dict is None and os.path.exists(CREDS_FILE):
        with open(CREDS_FILE, "r") as f:
            creds_dict = json.load(f)

    if creds_dict is None:
        raise ValueError(
            "No credentials found!\n"
            "Place your service account JSON file as 'credentials.json' in the same folder as app.py."
        )

    credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(credentials)


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def parse_zoom_report(uploaded_file) -> pd.DataFrame:
    """Parse Zoom participant CSV, auto-detecting the header row."""
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
    """Find the email column in Zoom data and return normalized email set."""
    for col in zoom_df.columns:
        if "email" in col.lower():
            emails = zoom_df[col].dropna().astype(str).str.strip().str.lower()
            return set(emails) - {""}
    return set()


def normalize_email(val) -> str:
    if not isinstance(val, str):
        return ""
    return val.strip().lower()


# ══════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════

st.title("📋 Zoom Attendance → Google Sheet Marker")

# ── Connect ──
try:
    gc = get_gspread_client()
except Exception as e:
    st.error(f"❌ Google Sheets connection failed: {e}")
    st.info(
        "**Fix:** Place your Google service account JSON file as `credentials.json` "
        "in the same folder as `app.py`."
    )
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

        if zoom_emails:
            st.info(f"📧 Extracted **{len(zoom_emails)}** unique emails from Zoom report(s).")
            with st.expander("Preview Zoom data"):
                st.dataframe(zoom_df.head(20), use_container_width=True)
        else:
            st.error("❌ No email column found in the Zoom CSV.")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STEP 2: Google Sheet Details
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if zoom_emails:
    st.header("2️⃣ Google Sheet Details")

    sheet_url = st.text_input(
        "🔗 Google Sheet URL",
        placeholder="https://docs.google.com/spreadsheets/d/xxxxx/edit",
    )

    col1, col2 = st.columns(2)
    with col1:
        tab_name = st.text_input("📄 Tab name", value="Sheet1")
    with col2:
        email_col_name = st.text_input("📧 Email column name (for matching)", value="Email")

    col3, col4 = st.columns(2)
    with col3:
        attendance_col_name = st.text_input("✅ Attendance column name (to update)", value="Attendance")
    with col4:
        update_value = st.text_input("📝 Value to mark", value="Present")

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # STEP 3: Preview & Mark
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    ws = None
    sheet_data = None

    if sheet_url and tab_name and email_col_name and attendance_col_name and update_value:
        try:
            if "docs.google.com" in sheet_url:
                spreadsheet = gc.open_by_url(sheet_url)
            else:
                spreadsheet = gc.open_by_key(sheet_url.strip())

            ws = spreadsheet.worksheet(tab_name)
            sheet_data = pd.DataFrame(ws.get_all_records())

            if sheet_data.empty:
                st.warning("⚠️ Worksheet is empty or has no header row.")
            else:
                sheet_cols = sheet_data.columns.tolist()
                if email_col_name not in sheet_cols:
                    st.error(f"❌ Column `{email_col_name}` not found. Available: `{'`, `'.join(sheet_cols)}`")
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

        # Match emails
        sheet_emails = sheet_data[email_col_name].astype(str).apply(normalize_email).tolist()
        matches = [e in zoom_emails for e in sheet_emails]

        matched_count = sum(matches)
        total = len(matches)

        m1, m2, m3 = st.columns(3)
        m1.metric("Total in Sheet", total)
        m2.metric("Matched ✅", matched_count)
        m3.metric("Not Matched", total - matched_count)

        # Preview
        preview = sheet_data.copy()
        preview["Status"] = ["✅ " + update_value if m else "—" for m in matches]

        name_col = None
        for c in sheet_data.columns:
            if "name" in c.lower():
                name_col = c
                break

        display_cols = []
        if name_col:
            display_cols.append(name_col)
        display_cols += [email_col_name, "Status"]
        st.dataframe(preview[display_cols], use_container_width=True, height=400)

        st.divider()

        if st.button(f"✅ Mark {matched_count} participants as `{update_value}`", type="primary", use_container_width=True):
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
                    for i, matched in enumerate(matches):
                        if matched:
                            cells_to_update.append(gspread.Cell(i + 2, att_col_idx, update_value))

                    if cells_to_update:
                        ws.update_cells(cells_to_update)

                    st.success(f"✅ Done! Marked **{matched_count}** participants as `{update_value}`.")
                    st.balloons()

                except Exception as e:
                    st.error(f"❌ Failed: {e}")
                    st.exception(e)

# ── Footer ──
st.divider()
st.caption("Built for Iron Lady 🔧")