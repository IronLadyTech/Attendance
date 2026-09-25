"""
Backfill MC Day 1 / Day 2 attendance from a Zoom participant CSV export.

Uses the same final-check rule as the live bridge: present at T+30 from the
session anchor → Yes; joined but not at T+30 → No; batch lead never joined → Absent.

  python tools/mc_import_zoom_report.py participants.csv \\
      --session-day "Day 2" --batch-date 2026-08-25 --session-date 2026-08-26 --anchor-ist 18:30

  python tools/mc_import_zoom_report.py participants.csv ... --apply \\
      --forward-url "https://flow.zoho.in/..."

  python tools/mc_import_zoom_report.py participants.csv ... --crm-apply
"""
from __future__ import annotations

import argparse
import csv
import glob
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import requests

_TOOLS = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_TOOLS)
if _TOOLS not in sys.path:
    sys.path.insert(0, _TOOLS)

_IST = timezone(timedelta(hours=5, minutes=30))
_FINAL_OFFSET_MIN = 30

_NAME_COLS = ("name (original name)", "name(original name)", "name", "user name")
_MAIL_COLS = ("user email", "email", "user e-mail")
_JOIN_COLS = ("join time", "join time (ist)", "join time (local)")
_LEAVE_COLS = ("leave time", "leave time (ist)", "leave time (local)")
_MINS_COLS = ("duration (minutes)", "duration(minutes)", "duration in minutes", "duration")

_STAFF_EMAIL_SUFFIXES = ("@iamironlady.com",)


def _load_dotenv() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    for c in (
        os.path.join(_ROOT, ".env"),
        os.path.join(_TOOLS, ".env"),
        os.path.join(_TOOLS, "zoho_flow", ".env"),
    ):
        if os.path.isfile(c):
            load_dotenv(c, override=False)


def _pick(header: list[str], wanted: tuple[str, ...]) -> str | None:
    low = {h.strip().lower(): h for h in header}
    for w in wanted:
        if w in low:
            return low[w]
    for key, orig in low.items():
        for w in wanted:
            if w in key:
                return orig
    return None


def normalize_name(s: str) -> str:
    s = (s or "").lower().strip()
    for t in ("ms. ", "mr. ", "mrs. ", "dr. ", "prof. ", "miss ", "rev. "):
        s = s.replace(t, "")
    return re.sub(r"\s+", " ", s).strip()


def _parse_zoom_local(ts: str) -> datetime | None:
    ts = (ts or "").strip().strip('"')
    if not ts:
        return None
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M:%S"):
        try:
            return datetime.strptime(ts, fmt).replace(tzinfo=_IST)
        except ValueError:
            continue
    return None


def read_report(path: str) -> list[dict]:
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as fh:
        rows = list(csv.reader(fh))
    header = None
    body: list[list[str]] = []
    for i, row in enumerate(rows):
        if not row:
            continue
        name_c = _pick(row, _NAME_COLS)
        join_c = _pick(row, _JOIN_COLS)
        if name_c and join_c:
            header = row
            body = rows[i + 1 :]
            break
    if not header:
        raise SystemExit(f"{path}: no participant table found")

    name_c = _pick(header, _NAME_COLS)
    mail_c = _pick(header, _MAIL_COLS)
    join_c = _pick(header, _JOIN_COLS)
    leave_c = _pick(header, _LEAVE_COLS)
    mins_c = _pick(header, _MINS_COLS)
    idx = {h: n for n, h in enumerate(header)}

    out: list[dict] = []
    for row in body:
        if not row or len(row) < len(header):
            continue
        rec = {h: row[idx[h]] for h in header if idx[h] < len(row)}
        name = (rec.get(name_c or "") or "").strip()
        if not name:
            continue
        join_dt = _parse_zoom_local(rec.get(join_c or "", ""))
        leave_dt = _parse_zoom_local(rec.get(leave_c or "", "")) if leave_c else None
        raw_mins = (rec.get(mins_c or "") or "0").strip() if mins_c else "0"
        try:
            mins = float(re.sub(r"[^0-9.]", "", raw_mins) or 0)
        except ValueError:
            mins = 0.0
        out.append(
            {
                "name": name,
                "email": (rec.get(mail_c or "") or "").strip().lower() if mail_c else "",
                "join": join_dt,
                "leave": leave_dt,
                "minutes": mins,
            }
        )
    return out


def _is_staff(email: str, name: str) -> bool:
    em = (email or "").lower()
    if any(em.endswith(s) for s in _STAFF_EMAIL_SUFFIXES):
        return True
    if "admin iron lady" in (name or "").lower():
        return True
    return False


def aggregate_mc(rows: list[dict], anchor: datetime, final_at: datetime) -> dict[str, dict]:
    """Classify each person: Yes / No / (not in cohort handling for Absent)."""
    people: dict[str, dict] = {}
    for r in rows:
        if _is_staff(r["email"], r["name"]):
            continue
        key = f"email:{r['email']}" if r["email"] else f"name:{normalize_name(r['name'])}"
        if not key.split(":", 1)[1]:
            continue
        p = people.setdefault(
            key,
            {
                "name": r["name"],
                "email": r["email"],
                "minutes": 0.0,
                "ever_joined": False,
                "present_at_final": False,
                "sessions": 0,
            },
        )
        p["minutes"] += r["minutes"]
        p["sessions"] += 1
        p["ever_joined"] = True
        if len(r["name"]) > len(p["name"]):
            p["name"] = r["name"]
        if r["email"] and not p["email"]:
            p["email"] = r["email"]
        join_dt = r["join"]
        leave_dt = r["leave"]
        if join_dt and leave_dt and join_dt <= final_at < leave_dt:
            p["present_at_final"] = True
        elif join_dt and join_dt <= final_at and leave_dt is None:
            p["present_at_final"] = True

    for p in people.values():
        if p["present_at_final"]:
            p["status"] = "Yes"
        elif p["ever_joined"]:
            p["status"] = "No"
        else:
            p["status"] = "No"
    return people


def _topic_for_day(session_day: str) -> str:
    if session_day == "Day 2":
        return "Art of War MC Day 2"
    return "BHAG Breakthrough Actions Day 1"


def _post_to_zoho(forward_url: str, payload: dict, label: str) -> None:
    try:
        r = requests.post(forward_url, json=payload, timeout=60)
        print(f"  [{label}] Zoho {r.status_code}")
    except requests.RequestException as e:
        print(f"  [{label}] ERROR {e}")


def _crm_apply(
    people: dict[str, dict],
    *,
    batch_date: str,
    session_day: str,
    dry_run: bool,
) -> tuple[int, int, int]:
    from zoho_crm_mlm import _access_token, _api_base, _crm_configured

    if not _crm_configured():
        raise SystemExit("CRM OAuth not configured (ZOHO_CRM_CLIENT_ID/SECRET/REFRESH_TOKEN)")

    att_field = "Day_2_Attendance" if session_day == "Day 2" else "Day_1_Attendance"
    batch_dt = datetime.strptime(batch_date, "%Y-%m-%d")
    day_before = (batch_dt - timedelta(days=1)).strftime("%Y-%m-%d")
    day_after = (batch_dt + timedelta(days=1)).strftime("%Y-%m-%d")
    day_start = f"{day_before}T00:00:00+05:30"
    day_end = f"{day_after}T23:59:59+05:30"

    token = _access_token()
    headers = {"Authorization": f"Zoho-oauthtoken {token}", "Content-Type": "application/json"}
    all_leads: list[dict] = []
    for offset in range(0, 2000, 200):
        q = (
            "select id, Email, Full_Name, Day_1_Attendance, Day_2_Attendance, "
            f"MC_Start_Date_Time from Leads where Payment_Status = 'Completed' "
            f"and (MC_Start_Date_Time between '{day_start}' and '{day_end}') "
            f"limit {offset}, 200"
        )
        resp = requests.post(
            f"{_api_base()}/coql",
            headers=headers,
            json={"select_query": q},
            timeout=60,
        )
        if resp.status_code == 401:
            try:
                detail = resp.json()
            except Exception:
                detail = resp.text
            raise SystemExit(
                "Zoho CRM 401 Unauthorized on COQL.\n"
                f"Response: {detail}\n\n"
                "Fix:\n"
                "  1. Use Zoho IN credentials (accounts.zoho.in / zohoapis.in)\n"
                "  2. Regenerate refresh token with scopes including:\n"
                "     ZohoCRM.modules.leads.ALL  (or READ+UPDATE)\n"
                "     ZohoCRM.coql.READ\n"
                "  3. Update ZOHO_CRM_REFRESH_TOKEN in tools/zoho_flow/.env"
            )
        resp.raise_for_status()
        data = resp.json()
        if data.get("status") == "error":
            raise SystemExit(f"COQL error: {data}")
        page = data.get("data") or []
        if not page:
            break
        for lead in page:
            mc_dt = lead.get("MC_Start_Date_Time")
            if mc_dt is None:
                continue
            ymd = str(mc_dt)[:10]
            if ymd == batch_date or batch_date in str(mc_dt):
                all_leads.append(lead)
        if len(page) < 200:
            break

    print(f"Batch leads in CRM: {len(all_leads)}")

    by_email = {p["email"]: p for p in people.values() if p.get("email")}
    by_name = {normalize_name(p["name"]): p for p in people.values()}

    yes_n = no_n = absent_n = 0
    updates: list[dict] = []

    for lead in all_leads:
        lid = lead["id"]
        email = (lead.get("Email") or "").strip().lower()
        name = lead.get("Full_Name") or ""
        person = by_email.get(email)
        if person is None and name:
            person = by_name.get(normalize_name(name))

        if person:
            status = person["status"]
        else:
            status = "Absent"

        cur = lead.get(att_field)
        if cur is not None and str(cur).strip() in ("Yes", "No", "Absent"):
            if str(cur).strip() == status:
                continue

        updates.append({"id": lid, att_field: status, "email": email, "name": name})
        if status == "Yes":
            yes_n += 1
        elif status == "No":
            no_n += 1
        else:
            absent_n += 1

    print(f"To update: Yes={yes_n} No={no_n} Absent={absent_n} (total={len(updates)})")
    if dry_run:
        for u in updates[:20]:
            print(f"  {u['email'] or u['name']}: -> {u[att_field]}")
        if len(updates) > 20:
            print(f"  ... and {len(updates) - 20} more")
        return yes_n, no_n, absent_n

    for i in range(0, len(updates), 100):
        chunk = updates[i : i + 100]
        body = {
            "data": [{"id": u["id"], att_field: u[att_field]} for u in chunk],
        }
        resp = requests.put(
            f"{_api_base()}/Leads",
            headers=headers,
            json=body,
            timeout=120,
        )
        resp.raise_for_status()
        result = resp.json()
        print(f"  CRM batch {i // 100 + 1}: {result.get('data', result)}")

    return yes_n, no_n, absent_n


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("reports", nargs="+", help="Zoom participant CSV path(s)")
    ap.add_argument("--session-day", default="Day 1", choices=["Day 1", "Day 2"])
    ap.add_argument("--batch-date", required=True, help="MC batch start (Day 1) yyyy-MM-dd")
    ap.add_argument(
        "--session-date",
        default="",
        help="Zoom session calendar date yyyy-MM-dd (Day 2 default: batch-date + 1 day)",
    )
    ap.add_argument(
        "--anchor-ist",
        default="18:30",
        help="Session anchor HH:MM IST on session-date (default 18:30)",
    )
    ap.add_argument("--forward-url", default="", help="Zoho Flow MC webhook URL")
    ap.add_argument("--csv-out", metavar="PATH", help="Write classification CSV")
    ap.add_argument("--apply", action="store_true", help="Post mark_yes + final_check to Zoho Flow")
    ap.add_argument("--crm-apply", action="store_true", help="Update Zoho CRM directly via OAuth API")
    ap.add_argument(
        "--crm-preview",
        action="store_true",
        help="Show CRM updates only (OAuth required; no writes)",
    )
    args = ap.parse_args()

    paths: list[str] = []
    for pat in args.reports:
        hits = glob.glob(pat)
        if not hits and os.path.isfile(pat):
            hits = [pat]
        if not hits:
            raise SystemExit(f"no file matched: {pat}")
        paths.extend(hits)

    batch_date = args.batch_date.strip()
    session_date = args.session_date.strip()
    if not session_date:
        batch_dt = datetime.strptime(batch_date, "%Y-%m-%d")
        if args.session_day == "Day 2":
            session_date = (batch_dt + timedelta(days=1)).strftime("%Y-%m-%d")
        else:
            session_date = batch_date
    hh, mm = args.anchor_ist.split(":")
    sy, smo, sd = map(int, session_date.split("-"))
    anchor = datetime(sy, smo, sd, int(hh), int(mm), 0, tzinfo=_IST)
    final_at = anchor + timedelta(minutes=_FINAL_OFFSET_MIN)

    print(f"batch={batch_date}  session={session_date}  field={'Day_2_Attendance' if args.session_day == 'Day 2' else 'Day_1_Attendance'}")

    rows: list[dict] = []
    for p in sorted(set(paths)):
        got = read_report(p)
        print(f"  {os.path.basename(p):<50} {len(got):>4} join rows")
        rows.extend(got)

    people = aggregate_mc(rows, anchor, final_at)
    yes = {k: v for k, v in people.items() if v["status"] == "Yes"}
    no = {k: v for k, v in people.items() if v["status"] == "No"}

    print(
        f"\nanchor={anchor.isoformat()}  final_check={final_at.isoformat()}  "
        f"people={len(people)}  Yes={len(yes)}  No={len(no)}\n"
    )
    for title, group in (("YES (at T+30)", yes), ("NO (joined, left before T+30)", no)):
        print(f"--- {title} ---")
        for v in sorted(group.values(), key=lambda x: (-x["minutes"], x["name"].lower())):
            print(f"  {v['name'][:36]:<36} {v['minutes']:6.0f} min  {v['email']}")
        print()

    if args.csv_out:
        with open(args.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["name", "email", "total_minutes", "sessions", "status"])
            for v in sorted(people.values(), key=lambda x: x["name"].lower()):
                w.writerow([v["name"], v["email"], round(v["minutes"], 1), v["sessions"], v["status"]])
        print(f"wrote {os.path.abspath(args.csv_out)}\n")

    if args.crm_preview or args.crm_apply:
        _load_dotenv()
        from zoho_crm_mlm import _crm_configured

        if not _crm_configured():
            raise SystemExit(
                "CRM OAuth not configured. Add to .env in repo root:\n"
                "  ZOHO_CRM_CLIENT_ID=...\n"
                "  ZOHO_CRM_CLIENT_SECRET=...\n"
                "  ZOHO_CRM_REFRESH_TOKEN=...\n"
                "  ZOHO_CRM_ACCOUNTS_URL=https://accounts.zoho.in\n"
                "  ZOHO_CRM_API_URL=https://www.zohoapis.in/crm/v7"
            )
        _crm_apply(
            people,
            batch_date=batch_date,
            session_day=args.session_day,
            dry_run=not args.crm_apply,
        )
        if args.crm_preview:
            print("\n(preview only — re-run with --crm-apply to write to CRM)")
        else:
            print("\nCRM update complete.")
        return 0

    if args.apply:
        _load_dotenv()
        forward = args.forward_url or os.environ.get("ZOHO_WEBHOOK_FORWARD_URL", "").strip()
        if not forward:
            raise SystemExit("Set --forward-url or ZOHO_WEBHOOK_FORWARD_URL")

        topic = _topic_for_day(args.session_day)
        base = {
            "meeting_id": "85212528515",
            "meeting_topic": topic,
            "topic": topic,
            "start_time": anchor.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "session_date": session_date,
            "session_day": args.session_day,
            "batch_date": batch_date,
            "program": "MC",
        }
        for v in sorted(yes.values(), key=lambda x: x["name"].lower()):
            p = dict(base)
            p.update(
                {
                    "event": "attendance.mark_yes",
                    "participant_email": v["email"],
                    "participant_name": v["name"],
                }
            )
            _post_to_zoho(forward, p, "mc-yes")

        present_emails = ",".join(v["email"] for v in yes.values() if v["email"])
        ever_emails = ",".join(v["email"] for v in people.values() if v["email"])
        b = dict(base)
        b.update(
            {
                "event": "attendance.final_check",
                "ever_joined_emails": ever_emails,
                "present_emails": present_emails,
            }
        )
        _post_to_zoho(forward, b, "mc-final")
        print(f"\nPosted {len(yes)} mark_yes + 1 final_check to Flow")
        return 0

    if not args.csv_out and not args.crm_preview:
        print("(preview — use --crm-preview, --crm-apply, --apply, or --csv-out)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
