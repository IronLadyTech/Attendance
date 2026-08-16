"""
Backfill LEP attendance from Zoom's own participant reports.

Use when a session ran without checkpoints (no QStash schedule), so Redis holds
only "ever joined" and cannot tell a full attendee from a 30-second visitor.
Zoom's usage report has real join/leave times, so a duration threshold gives a
defensible Present/Absent comparable to the checkpoint standard.

Export from Zoom: Reports -> Usage -> pick the date -> click the participant
count for each room -> Export. One CSV per room; pass them all.

  python tools/lep_import_zoom_report.py day1_*.csv \
      --batch-date 2026-08-15 --session-day "Day 1"                 # preview
  python tools/lep_import_zoom_report.py day1_*.csv \
      --batch-date 2026-08-15 --session-day "Day 1" --apply         # post to Zoho

Rows are aggregated per person across every join/leave and every room, so
someone who drops and rejoins is credited with their total time.
"""
from __future__ import annotations

import argparse
import csv
import glob
import os
import re
import sys

_TOOLS = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_TOOLS)
if _TOOLS not in sys.path:
    sys.path.insert(0, _TOOLS)


def _load_dotenv() -> str:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return ""
    for c in (os.path.join(_ROOT, ".env"), os.path.join(_TOOLS, ".env"),
              os.path.join(_TOOLS, "zoho_flow", ".env")):
        if os.path.isfile(c):
            load_dotenv(c, override=False)
            return c
    return ""


# Zoom changes these headers between plans and exports.
_NAME_COLS = ("name (original name)", "name(original name)", "name", "user name")
_MAIL_COLS = ("user email", "email", "user e-mail")
_MINS_COLS = ("duration (minutes)", "duration(minutes)", "duration in minutes",
              "duration")


def _pick(header: list[str], wanted: tuple[str, ...]) -> str | None:
    low = {h.strip().lower(): h for h in header}
    for w in wanted:
        if w in low:
            return low[w]
    for key, orig in low.items():          # substring fallback
        for w in wanted:
            if w in key:
                return orig
    return None


def normalize_name(s: str) -> str:
    s = (s or "").lower().strip()
    for t in ("ms. ", "mr. ", "mrs. ", "dr. ", "prof. ", "miss ", "rev. "):
        s = s.replace(t, "")
    return re.sub(r"\s+", " ", s).strip()


def read_report(path: str) -> list[dict]:
    """One row per join session, as Zoom exports it."""
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as fh:
        rows = list(csv.reader(fh))
    # Zoom prepends a meeting-summary block; the participant table is the first
    # row that carries both a name and a duration column.
    for i, row in enumerate(rows):
        if not row:
            continue
        name_c = _pick(row, _NAME_COLS)
        mins_c = _pick(row, _MINS_COLS)
        if name_c and mins_c:
            header = row
            body = rows[i + 1:]
            break
    else:
        raise SystemExit(f"{path}: no participant table found (need a name and "
                         f"duration column)")
    name_c, mail_c, mins_c = (_pick(header, _NAME_COLS),
                              _pick(header, _MAIL_COLS),
                              _pick(header, _MINS_COLS))
    idx = {h: n for n, h in enumerate(header)}
    out = []
    for row in body:
        if not row or len(row) < len(header):
            continue
        rec = {h: row[idx[h]] for h in header if idx[h] < len(row)}
        name = (rec.get(name_c) or "").strip()
        if not name:
            continue
        raw = (rec.get(mins_c) or "0").strip()
        try:
            mins = float(re.sub(r"[^0-9.]", "", raw) or 0)
        except ValueError:
            mins = 0.0
        out.append({
            "name": name,
            "email": (rec.get(mail_c) or "").strip().lower() if mail_c else "",
            "minutes": mins,
            "source": os.path.basename(path),
        })
    return out


def aggregate(rows: list[dict]) -> dict[str, dict]:
    """Total minutes per person; email wins over name as the identity."""
    people: dict[str, dict] = {}
    for r in rows:
        key = f"email:{r['email']}" if r["email"] else f"name:{normalize_name(r['name'])}"
        if not key.split(":", 1)[1]:
            continue
        p = people.setdefault(key, {"name": r["name"], "email": r["email"],
                                    "minutes": 0.0, "rooms": set(), "sessions": 0})
        p["minutes"] += r["minutes"]
        p["rooms"].add(r["source"])
        p["sessions"] += 1
        if len(r["name"]) > len(p["name"]):     # prefer the fuller display name
            p["name"] = r["name"]
        if r["email"] and not p["email"]:
            p["email"] = r["email"]
    return people


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("reports", nargs="+", help="Zoom participant CSV(s); globs ok")
    ap.add_argument("--batch-date", required=True, help="cohort LEP Start Date, YYYY-MM-DD")
    ap.add_argument("--session-day", default="Day 1", choices=["Day 1", "Day 2"])
    ap.add_argument("--threshold", type=float, default=60.0,
                    help="minutes required for Present (default 60)")
    ap.add_argument("--topic", default="IronLady LEP Day 1 & 2 session",
                    help="must contain an accepted LEP keyword")
    ap.add_argument("--forward-url", default="",
                    help="Zoho Flow webhook; default reads it from the Redis session")
    ap.add_argument("--csv", metavar="PATH", help="write the computed table here")
    ap.add_argument("--apply", action="store_true", help="actually post to Zoho")
    args = ap.parse_args()

    paths: list[str] = []
    for pat in args.reports:
        hits = glob.glob(pat)
        if not hits:
            raise SystemExit(f"no file matched: {pat}")
        paths.extend(hits)

    rows: list[dict] = []
    for p in sorted(set(paths)):
        got = read_report(p)
        print(f"  {os.path.basename(p):<42} {len(got):>4} join rows")
        rows.extend(got)
    people = aggregate(rows)

    present = {k: v for k, v in people.items() if v["minutes"] >= args.threshold}
    short = {k: v for k, v in people.items() if v["minutes"] < args.threshold}
    print(f"\nfiles={len(set(paths))} join_rows={len(rows)} people={len(people)}")
    print(f"threshold={args.threshold:g} min  ->  Present={len(present)}  "
          f"below={len(short)}\n")

    for title, group in (("PRESENT", present), ("BELOW THRESHOLD", short)):
        print(f"--- {title} ---")
        for k, v in sorted(group.items(), key=lambda kv: -kv[1]["minutes"]):
            print(f"  {v['name'][:34]:<34} {v['minutes']:7.0f} min  "
                  f"{v['sessions']:>2} join(s)  {v['email']}")
        print()

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["name", "email", "identity", "total_minutes",
                        "join_sessions", "rooms", "result"])
            for k, v in sorted(people.items(), key=lambda kv: -kv[1]["minutes"]):
                w.writerow([v["name"], v["email"], k, round(v["minutes"], 1),
                            v["sessions"], len(v["rooms"]),
                            "Present" if v["minutes"] >= args.threshold else "Absent"])
        print(f"wrote {os.path.abspath(args.csv)}\n")

    if not args.apply:
        print("(preview only — re-run with --apply to post to Zoho)")
        return 0

    _load_dotenv()
    import lep_redis as lr
    from zoom_webhook_bridge import _post_to_zoho

    forward = args.forward_url
    if not forward:
        sk = lr.session_key_for(args.batch_date, args.session_day)
        forward = lr.get_session_meta(sk).get("forward_url", "")
    if not forward:
        raise SystemExit("no forward_url — pass --forward-url")

    base = {
        "meeting_id": "", "meeting_topic": args.topic, "topic": args.topic,
        "start_time": "", "session_date": args.batch_date,
        "session_day": args.session_day, "batch_date": args.batch_date,
        "program": "LEP",
    }
    for v in sorted(present.values(), key=lambda x: x["name"].lower()):
        p = dict(base)
        p.update({"event": "attendance.lep_final",
                  "participant_email": v["email"], "participant_name": v["name"],
                  "check_number": "final", "attendance_result": "Present",
                  "attendance_scope": "GLOBAL"})
        _post_to_zoho(forward, p, "lep-zoomreport")

    emails = ",".join(v["email"] for v in present.values() if v["email"])
    names = ",".join(v["name"] for v in present.values() if v["name"])
    b = dict(base)
    b.update({"event": "attendance.lep_final_batch", "check_number": "final",
              "present_emails": emails, "present_names": names,
              "ever_joined_emails": emails, "ever_joined_names": names,
              "attendance_scope": "GLOBAL"})
    _post_to_zoho(forward, b, "lep-zoomreport-batch")
    print(f"\nposted {len(present)} Present + 1 batch sweep for "
          f"{args.session_day} cohort {args.batch_date}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
