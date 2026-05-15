"""
Shared attendance logic (no Streamlit). Used by app.py and auto_sync_attendance.py.
"""
from __future__ import annotations

import base64
import json
import os
import re
import time
import gspread
import pandas as pd
import requests
from google.oauth2.service_account import Credentials

ROOT = os.path.dirname(os.path.abspath(__file__))
CREDS_FILE = os.path.join(ROOT, "credentials.json")
ZOOM_CONFIG_FILE = os.path.join(ROOT, "zoom_config.json")
ZOOM_TOKEN_URL = "https://zoom.us/oauth/token"
ZOOM_API_BASE = "https://api.zoom.us/v2"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# In-memory Zoom token cache for CLI / non-Session use
_zoom_cli_cache: dict = {}


def load_zoom_config_from_disk() -> dict | None:
    cfg = None
    if os.path.exists(ZOOM_CONFIG_FILE):
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


def get_gspread_client_from_file() -> gspread.Client:
    if not os.path.exists(CREDS_FILE):
        raise FileNotFoundError(f"Missing {CREDS_FILE}")
    with open(CREDS_FILE, "r", encoding="utf-8") as f:
        creds_dict = json.load(f)
    credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(credentials)


def get_zoom_access_token_cli(config: dict) -> str:
    now = time.time()
    if (
        _zoom_cli_cache.get("token")
        and float(_zoom_cli_cache.get("expires", 0)) > now + 60
    ):
        return _zoom_cli_cache["token"]

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
    resp.raise_for_status()
    js = resp.json()
    _zoom_cli_cache["token"] = js["access_token"]
    _zoom_cli_cache["expires"] = now + float(js.get("expires_in", 3600))
    return _zoom_cli_cache["token"]


def fetch_meeting_participants_report(meeting_id: str, token: str) -> pd.DataFrame:
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


def worksheet_to_dataframe(ws: gspread.Worksheet) -> tuple[pd.DataFrame, list[str]]:
    all_values = ws.get_all_values()
    if len(all_values) < 2:
        return pd.DataFrame(), []
    headers = all_values[0]
    seen: dict[str, int] = {}
    unique_headers: list[str] = []
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
    df = pd.DataFrame(all_values[1:], columns=unique_headers)
    return df, headers


def normalize(val) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"\s+", " ", val.strip().lower())


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


def match_with_fallback(
    sheet_data: pd.DataFrame,
    email_col: str,
    name_col: str,
    zoom_emails: set[str],
    zoom_names: set[str],
) -> pd.DataFrame:
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
    out = sheet_data.copy()
    out["_match_status"] = results
    return out


def apply_attendance_to_sheet(
    ws: gspread.Worksheet,
    matched_data: pd.DataFrame,
    attendance_col_name: str,
    update_value: str,
) -> int:
    header_row = ws.row_values(1)
    if attendance_col_name in header_row:
        att_col_idx = header_row.index(attendance_col_name) + 1
    else:
        att_col_idx = len(header_row) + 1
        ws.update_cell(1, att_col_idx, attendance_col_name)

    cells = []
    for i, status in enumerate(matched_data["_match_status"]):
        if "Matched" in status:
            cells.append(gspread.Cell(i + 2, att_col_idx, update_value))
    if cells:
        ws.update_cells(cells)
    return len(cells)
