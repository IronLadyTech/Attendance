#!/usr/bin/env python3
"""
Scheduled / automated Zoom → Google Sheet attendance sync.

Usage:
  python auto_sync_attendance.py --meeting-id 1234567890 --sheet-url "https://docs.google.com/..." \\
    --tab Sheet1 --email-col Email --name-col Name --attendance-col "Day 1" --value Yes

Optional wait before fetching Zoom report (e.g. end of class + buffer):
  python auto_sync_attendance.py --delay-minutes 30 ...

Dry run (no sheet writes):
  python auto_sync_attendance.py --dry-run ...

Schedule on Windows: Task Scheduler → Create Basic Task → Action: Start a program
  Program: python
  Arguments: "D:\\Attendence\\Attendance\\auto_sync_attendance.py" --meeting-id ... (full args)
  Start in: D:\\Attendence\\Attendance

Day 2 next week: same command with --attendance-col "Day 2" (add columns in the sheet first).
"""
from __future__ import annotations

import argparse
import sys
import time

import attendance_core as ac


def main() -> int:
    p = argparse.ArgumentParser(description="Sync Zoom participant report to Google Sheet attendance column.")
    p.add_argument("--meeting-id", required=True, help="Zoom numeric meeting ID")
    p.add_argument("--sheet-url", required=True, help="Google Sheet URL or spreadsheet ID")
    p.add_argument("--tab", default="Sheet1", help="Worksheet tab name")
    p.add_argument("--email-col", default="Email", help="Sheet column for email (must match header)")
    p.add_argument("--name-col", default="Name", help="Sheet column for name")
    p.add_argument(
        "--attendance-col",
        required=True,
        help='Header to write (e.g. "Day 1", "Day 2", "Attendance") — created if missing',
    )
    p.add_argument("--value", default="Yes", help="Cell value for matched rows")
    p.add_argument(
        "--delay-minutes",
        type=float,
        default=0,
        help="Wait this many minutes before calling Zoom (0 = immediately)",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Fetch Zoom + match only; do not update the sheet",
    )
    args = p.parse_args()

    zcfg = ac.load_zoom_config_from_disk()
    if not zcfg:
        print("ERROR: No Zoom config. Set zoom_config.json or ZOOM_* env vars.", file=sys.stderr)
        return 1

    if args.delay_minutes > 0:
        print(f"Waiting {args.delay_minutes} minutes before Zoom fetch...")
        time.sleep(args.delay_minutes * 60.0)

    print("Fetching Zoom participant report...")
    token = ac.get_zoom_access_token_cli(zcfg)
    zoom_df = ac.fetch_meeting_participants_report(args.meeting_id, token)
    if zoom_df.empty:
        print("WARNING: Zoom returned 0 participants. Meeting may still be live or report not ready.")
    zoom_emails = ac.extract_zoom_emails(zoom_df)
    zoom_names = ac.extract_zoom_names(zoom_df)
    print(f"Zoom: {len(zoom_df)} row(s), {len(zoom_emails)} emails, {len(zoom_names)} names")

    if args.dry_run:
        print("Dry run: skipping Google Sheets.")
        return 0

    print("Opening Google Sheet...")
    gc = ac.get_gspread_client_from_file()
    if "docs.google.com" in args.sheet_url:
        spreadsheet = gc.open_by_url(args.sheet_url)
    else:
        spreadsheet = gc.open_by_key(args.sheet_url.strip())
    ws = spreadsheet.worksheet(args.tab)
    sheet_data, _ = ac.worksheet_to_dataframe(ws)
    if sheet_data.empty:
        print("ERROR: Sheet has no data rows.", file=sys.stderr)
        return 1

    for col in (args.email_col, args.name_col):
        if col not in sheet_data.columns:
            print(f"ERROR: Column '{col}' not found. Available: {list(sheet_data.columns)}", file=sys.stderr)
            return 1

    matched = ac.match_with_fallback(
        sheet_data, args.email_col, args.name_col, zoom_emails, zoom_names
    )
    email_n = (matched["_match_status"] == "✅ Matched by Email").sum()
    name_n = (matched["_match_status"] == "⚠️ Matched by Name (fallback)").sum()
    bad_n = (matched["_match_status"] == "❌ Unmatched").sum()
    total_ok = int(email_n + name_n)
    print(f"Match: {email_n} by email, {name_n} by name, {bad_n} unmatched (sheet rows: {len(matched)})")

    if total_ok == 0:
        print("Nothing to update.")
        return 0

    n = ac.apply_attendance_to_sheet(ws, matched, args.attendance_col, args.value)
    print(f"Updated {n} cell(s) in column '{args.attendance_col}' with value '{args.value}'.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
