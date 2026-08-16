"""
Read-only inspector for LEP checkpoint truth in Redis.

Shows exactly what execute_final will see: the C1/C2/C3 snapshots, who was
in each, and the majority each participant resolves to. Writes nothing.

  python tools/lep_inspect.py --session lep:2026-08-16:day1
  python tools/lep_inspect.py --date 2026-08-16 --day 1
  python tools/lep_inspect.py --date 2026-08-16 --day 1 --who namita

Needs UPSTASH_REDIS_REST_URL + UPSTASH_REDIS_REST_TOKEN in the environment
(same values as Render). LEP_DURABLE is not required — this only reads.
"""
from __future__ import annotations

import argparse
import csv
import os
import sys

_TOOLS = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_TOOLS)
if _TOOLS not in sys.path:
    sys.path.insert(0, _TOOLS)


def _load_dotenv() -> None:
    """Pick up UPSTASH_* from a local .env so nothing has to be exported."""
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    for candidate in (os.path.join(_ROOT, ".env"), os.path.join(_TOOLS, ".env")):
        if os.path.isfile(candidate):
            load_dotenv(candidate, override=False)


_load_dotenv()

import lep_redis as lr  # noqa: E402
from lep_checkpoint import _alt_identities  # noqa: E402  (match final's matching exactly)
from lep_identity import build_lep_session_key, calculate_final  # noqa: E402


def _label(session_key: str, ident: str) -> str:
    disp = lr.load_identity_display(session_key, ident)
    name = (disp.get("display_name") or "").strip()
    email = (disp.get("email") or "").strip()
    if name and email:
        return f"{name} <{email}>"
    if name:
        return name
    if email:
        return email
    return ident


def _checkpoint_rows(session_key: str) -> list[tuple[str, str, str, str, str, str]]:
    ever = lr.build_global_ever_joined(session_key)
    rows = []
    for ident in sorted(ever):
        alts = _alt_identities(session_key, ident)
        c1 = lr.checkpoint_result_for_identity(session_key, 1, ident, alts)
        c2 = lr.checkpoint_result_for_identity(session_key, 2, ident, alts)
        c3 = lr.checkpoint_result_for_identity(session_key, 3, ident, alts)
        rows.append((_label(session_key, ident), ident, c1, c2, c3,
                     calculate_final(c1, c2, c3)))
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="Inspect LEP Redis checkpoints (read-only)")
    ap.add_argument("--session", help="e.g. lep:2026-08-16:day1")
    ap.add_argument("--date", help="session date YYYY-MM-DD (with --day)")
    ap.add_argument("--day", choices=["1", "2"], help="1 or 2 (with --date)")
    ap.add_argument("--who", help="only rows whose name/email/identity contains this")
    ap.add_argument("--csv", metavar="PATH", help="also write the table to a CSV file")
    args = ap.parse_args()

    if args.session:
        session_key = args.session.strip()
    elif args.date and args.day:
        session_key = build_lep_session_key(args.date.strip(), f"Day {args.day}")
    else:
        ap.error("give --session, or both --date and --day")

    if not lr.redis_configured():
        sys.stderr.write(
            "UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN not set.\n\n"
            "Get them from the Upstash console: open the database, scroll to\n"
            "'REST API', copy UPSTASH_REDIS_REST_URL and UPSTASH_REDIS_REST_TOKEN\n"
            "(the https:// one, NOT the redis:// connection string).\n\n"
            f"Then put them in {os.path.join(_ROOT, '.env')} as:\n"
            "  UPSTASH_REDIS_REST_URL=https://....upstash.io\n"
            "  UPSTASH_REDIS_REST_TOKEN=...\n"
        )
        return 2
    if lr.get_redis() is None:
        sys.stderr.write("Could not build a Redis client (check the REST URL/token).\n")
        return 2

    print(f"session   : {session_key}")
    meta = lr.get_session_meta(session_key)
    if not meta:
        print("\n!! No session meta at this key — nothing was ever recorded for it.")
        print("   Day 2 uses the SATURDAY batch date, not Sunday's date.")
        return 1
    print(f"topic     : {meta.get('topic', '')}")
    print(f"batch date: {meta.get('batch_date', '')}   day: {meta.get('session_day', '')}")

    r = lr.get_redis()
    rooms = lr.list_meetings(session_key)
    print(f"\nrooms seen: {len(rooms)}")
    for zoom_account, meeting_id in sorted(rooms):
        cur = r.smembers(
            lr.meeting_current_key(session_key, zoom_account, meeting_id)
        ) or set()
        ever = r.smembers(
            lr.meeting_ever_joined_key(session_key, zoom_account, meeting_id)
        ) or set()
        print(f"  {zoom_account:6} meeting={meeting_id:14} in_room_now={len(cur):3} "
              f"ever_joined={len(ever):3}")

    print("\ncheckpoints")
    for check_no in (1, 2, 3):
        state = lr.checkpoint_state(session_key, check_no) or "not run"
        status = r.hgetall(lr.checkpoint_status_key(session_key, check_no)) or {}
        present = lr.load_checkpoint_present(session_key, check_no)
        stamp = status.get("completed_at", "")
        seen = status.get("rooms_seen", "?")
        print(f"  C{check_no}: {state:10} present={len(present):3} rooms={seen:>2}  {stamp}")

    final_state = lr.final_state(session_key) or "not run"
    print(f"  final: {final_state}")

    rows = _checkpoint_rows(session_key)
    if args.who:
        needle = args.who.strip().lower()
        rows = [r for r in rows if needle in r[0].lower() or needle in r[1].lower()]

    print(f"\nparticipants who ever joined: {len(rows)}")
    print(f"  {'NAME':<38} {'C1':<8} {'C2':<8} {'C3':<8} FINAL")
    print("  " + "-" * 78)
    for label, _ident, c1, c2, c3, final in rows:
        print(f"  {label[:38]:<38} {c1:<8} {c2:<8} {c3:<8} {final}")

    if args.csv:
        # utf-8-sig so Excel on Windows renders names like "Namita’s" correctly.
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["name", "email", "identity", "check_1", "check_2",
                        "check_3", "final", "session"])
            for label, ident, c1, c2, c3, final in rows:
                disp = lr.load_identity_display(session_key, ident)
                w.writerow([
                    disp.get("display_name", ""), disp.get("email", ""), ident,
                    c1, c2, c3, final, session_key,
                ])
        print(f"\nwrote {len(rows)} rows to {os.path.abspath(args.csv)}")

    tally: dict[str, int] = {}
    for *_rest, final in rows:
        tally[final] = tally.get(final, 0) + 1
    print("\n  " + "   ".join(f"{k}={v}" for k, v in sorted(tally.items())))
    if tally.get("Unresolved"):
        print("  Unresolved = a checkpoint never ran for them; final refuses to "
              "guess Absent.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
