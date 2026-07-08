#!/usr/bin/env python3
"""
Zoom webhook bridge for Iron Lady MC and 100BM checkpoint attendance.

MC route (POST /):
  - meeting.started (MC topic) → schedule T+15 and T+30 checkpoint sweeps
  - meeting.participant_joined / left → maintain in-meeting roster (no per-person timer)
  - T+15 → mark_yes for everyone in roster + attendance.first_check to Zoho
  - T+30 → mark_yes for everyone in roster + attendance.final_check to Zoho
  - T+60 → mark_yes for everyone in roster + attendance.hour_check (late Yes upgrade only)
  - meeting.ended (MC topic) → meeting.ended to Zoho (MC Completed only; no attendance)

100BM route (POST /100bm) — same T+15 / T+30 / T+60 checkpoint model as MC:
  - meeting.started (100BM topic) → schedule remaining checkpoint sweeps from start_time
  - meeting.participant_joined / left → maintain roster; after redeploy, recover timers from start_time
  - T+15 → mark_yes (in room) + lookup (joined-left) + attendance.first_check
  - T+30 → mark_yes (in room) + mark_no (joined-left dropout) + attendance.final_check
  - T+60 → mark_yes (in room) + attendance.hour_check (upgrade No/Absent → Yes only)

Environment variables required:
  ZOOM_WEBHOOK_SECRET_TOKEN, ZOHO_WEBHOOK_FORWARD_URL

Optional:
  ZOOM_WEBHOOK_SECRET_TOKEN_100BM, ZOHO_WEBHOOK_FORWARD_URL_100BM
  MC_CHECKPOINT_1_SECONDS (default 900 = 15 min)
  MC_CHECKPOINT_2_SECONDS (default 1800 = 30 min)
  MC_CHECKPOINT_3_SECONDS (default 3600 = 60 min)
  BM100_CHECKPOINT_1_SECONDS (default 900 = 15 min, 100BM route)
  BM100_CHECKPOINT_2_SECONDS (default 1800 = 30 min, 100BM route)
  BM100_CHECKPOINT_3_SECONDS (default 3600 = 60 min, 100BM route)
"""
from __future__ import annotations

import argparse
import hashlib
import hmac
import json
import os
import sys
import threading
import time
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from socketserver import ThreadingMixIn

import requests

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# MC checkpoint delays from actual Zoom meeting start
_MC_CHECKPOINT_1 = int(os.environ.get("MC_CHECKPOINT_1_SECONDS", "900"))   # T+15
_MC_CHECKPOINT_2 = int(os.environ.get("MC_CHECKPOINT_2_SECONDS", "1800"))  # T+30
_MC_CHECKPOINT_3 = int(os.environ.get("MC_CHECKPOINT_3_SECONDS", "3600"))  # T+60

# 100BM checkpoint delays (same defaults as MC)
_BM100_CHECKPOINT_1 = int(os.environ.get("BM100_CHECKPOINT_1_SECONDS", "900"))   # T+15
_BM100_CHECKPOINT_2 = int(os.environ.get("BM100_CHECKPOINT_2_SECONDS", "1800"))  # T+30
_BM100_CHECKPOINT_3 = int(os.environ.get("BM100_CHECKPOINT_3_SECONDS", "3600"))  # T+60

_DAY1_KEYWORDS = ["bhag", "breakthrough actions"]
_DAY2_KEYWORDS = ["art of war", "shameless pitching"]
_MC_MEETING_KEYWORDS = _DAY1_KEYWORDS + _DAY2_KEYWORDS

_100BM_KEYWORDS = ["orientation session", "fast track your leadership growth"]

_IST = timezone(timedelta(hours=5, minutes=30))

# MC per-meeting roster + checkpoint timers
_mc_meetings: dict[str, dict] = {}
_mc_lock = threading.Lock()

# 100BM per-meeting roster + checkpoint timers
_100bm_meetings: dict[str, dict] = {}
_100bm_lock = threading.Lock()


def _load_dotenv() -> None:
    try:
        from dotenv import load_dotenv
        load_dotenv(os.path.join(_REPO_ROOT, ".env"))
    except ImportError:
        pass


def _is_mc_topic(topic: str) -> bool:
    topic_l = topic.lower()
    return any(kw in topic_l for kw in _MC_MEETING_KEYWORDS)


def _is_100bm_topic(topic: str) -> bool:
    topic_l = topic.lower()
    return any(kw in topic_l for kw in _100BM_KEYWORDS)


def _mc_session_day(topic: str) -> str:
    topic_l = topic.lower()
    if any(kw in topic_l for kw in _DAY2_KEYWORDS):
        return "Day 2"
    return "Day 1"


def _parse_zoom_datetime(iso_ts: str) -> datetime:
    if not iso_ts or iso_ts.strip() in ("", "null"):
        return datetime.now(timezone.utc)
    s = iso_ts.strip()
    if s.endswith("Z"):
        try:
            return datetime.strptime(s, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    try:
        parsed = datetime.fromisoformat(s)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except ValueError:
        pass
    return datetime.now(timezone.utc)


def _session_date_ist(iso_ts: str) -> str:
    return _parse_zoom_datetime(iso_ts).astimezone(_IST).strftime("%Y-%m-%d")


def _batch_date(session_date: str, session_day: str) -> str:
    """Day 1 → batch = session date. Day 2 → batch = session date minus 1 day."""
    if session_day == "Day 1":
        return session_date
    dt = datetime.strptime(session_date, "%Y-%m-%d").date()
    return (dt - timedelta(days=1)).strftime("%Y-%m-%d")


def _validation_response(body: dict, secret: str) -> bytes:
    pt = body.get("payload", {}).get("plainToken")
    if not pt or not secret:
        raise ValueError("missing plainToken or ZOOM_WEBHOOK_SECRET_TOKEN")
    enc = hmac.new(secret.encode("utf-8"), pt.encode("utf-8"), hashlib.sha256).hexdigest()
    return json.dumps({"plainToken": pt, "encryptedToken": enc}).encode("utf-8")


def _post_to_zoho(forward_url: str, payload: dict, label: str) -> None:
    try:
        r = requests.post(forward_url, json=payload, timeout=30)
        sys.stderr.write(f"[{label}] Zoho response: {r.status_code}\n")
    except requests.RequestException as e:
        sys.stderr.write(f"[{label}] ERROR posting to Zoho: {e}\n")


def _mc_roster_key(email: str, name: str) -> str:
    if email:
        return f"email:{email.lower().strip()}"
    if name:
        return f"name:{name.lower().strip()}"
    return ""


def _mc_base_payload(state: dict) -> dict:
    return {
        "meeting_id": state["meeting_id"],
        "meeting_topic": state["topic"],
        "topic": state["topic"],
        "start_time": state["start_time"],
        "session_date": state["session_date"],
        "batch_date": state["batch_date"],
        "session_day": state["session_day"],
        "program": "MC",
    }


def _mc_mark_no(forward_url: str, state: dict, email: str, name: str, join_time: str) -> None:
    """T+30 only: joined at some point but not in room at final checkpoint → No (Yes→No allowed)."""
    payload = _mc_base_payload(state)
    payload.update({
        "event": "attendance.mark_no",
        "participant_email": email,
        "participant_name": name,
        "join_time": join_time,
    })
    _post_to_zoho(forward_url, payload, "mc-no")


def _mc_mark_yes(forward_url: str, state: dict, email: str, name: str, join_time: str) -> None:
    payload = _mc_base_payload(state)
    payload.update({
        "event": "attendance.mark_yes",
        "participant_email": email,
        "participant_name": name,
        "join_time": join_time,
    })
    _post_to_zoho(forward_url, payload, "mc-yes")


def _mc_lookup_participant(forward_url: str, state: dict, email: str, name: str, join_time: str) -> None:
    """Joined but left before checkpoint — verify CRM match or log to unmatched sheet."""
    payload = _mc_base_payload(state)
    payload.update({
        "event": "attendance.lookup",
        "participant_email": email,
        "participant_name": name,
        "join_time": join_time,
    })
    _post_to_zoho(forward_url, payload, "mc-lookup")


def _ever_joined_email_csv(ever_joined: dict) -> str:
    emails: list[str] = []
    for info in ever_joined.values():
        e = (info.get("email") or "").strip().lower()
        if e:
            emails.append(e)
    return ",".join(emails)


def _roster_email_csv(roster: dict) -> str:
    """Emails in the meeting room at this checkpoint (T+30 present list)."""
    emails: list[str] = []
    for info in roster.values():
        e = (info.get("email") or "").strip().lower()
        if e:
            emails.append(e)
    return ",".join(emails)


def _checkpoint_event(sweep: int) -> str:
    if sweep == 1:
        return "attendance.first_check"
    if sweep == 2:
        return "attendance.final_check"
    return "attendance.hour_check"


def _mc_sweep(meeting_id: str, sweep: int) -> None:
    with _mc_lock:
        state = _mc_meetings.get(meeting_id)
        if not state:
            sys.stderr.write(f"[sweep{sweep}] No state for meeting={meeting_id}\n")
            return
        roster = dict(state["roster"])
        ever_joined = dict(state.get("ever_joined", {}))
        forward_url = state["forward_url"]

    sys.stderr.write(
        f"[sweep{sweep}] meeting={meeting_id} roster={len(roster)} "
        f"ever_joined={len(ever_joined)} topic={state['topic']!r} batch={state['batch_date']}\n"
    )

    # Present at checkpoint → mark Yes (CRM) or unmatched sheet if no lead
    for info in roster.values():
        email = info.get("email", "")
        name = info.get("name", "")
        if not email and not name:
            continue
        _mc_mark_yes(forward_url, state, email, name, info.get("join_time", ""))

    # T+15 only: joined earlier but left before check → verify match or sheet
    if sweep == 1:
        roster_keys = set(roster.keys())
        for rkey, info in ever_joined.items():
            if rkey in roster_keys:
                continue
            email = info.get("email", "")
            name = info.get("name", "")
            if not email and not name:
                continue
            _mc_lookup_participant(forward_url, state, email, name, info.get("join_time", ""))

    # T+30 only: joined earlier but not in room now → No (downgrades Yes from T+15 if they left)
    if sweep == 2:
        roster_keys = set(roster.keys())
        for rkey, info in ever_joined.items():
            if rkey in roster_keys:
                continue
            email = info.get("email", "")
            name = info.get("name", "")
            if not email and not name:
                continue
            _mc_mark_no(forward_url, state, email, name, info.get("join_time", ""))

    # T+60: no lookup, no mark_no — only mark_yes (above) + hour_check batch upgrade

    event = _checkpoint_event(sweep)
    payload = _mc_base_payload(state)
    payload["event"] = event
    payload["ever_joined_emails"] = _ever_joined_email_csv(ever_joined)
    payload["present_emails"] = _roster_email_csv(roster)
    _post_to_zoho(forward_url, payload, f"mc-sweep{sweep}")


def _mc_cancel_timers(state: dict) -> None:
    for t in state.get("timers", []):
        t.cancel()


def _seconds_since_meeting_start(start_time: str) -> float:
    """UTC seconds elapsed since Zoom start_time (0 if missing/unparseable)."""
    if not start_time or str(start_time).strip() in ("", "null"):
        return 0.0
    started = _parse_zoom_datetime(str(start_time))
    return max(0.0, (datetime.now(timezone.utc) - started).total_seconds())


def _schedule_checkpoint_timers(
    meeting_id: str,
    state: dict,
    *,
    sweep_fn,
    delays: list[tuple[int, int]],
    log_prefix: str,
) -> None:
    """
    Schedule only checkpoints still in the future, based on meeting start_time.

    Survives Render/Railway redeploys: after restart, the next join recreates state and
    this recomputes remaining delay (e.g. only T+60 left). Past first/final are skipped
    so we do not re-fire blank→No / Absent. If the hour check is already due (or slightly
    overdue), run sweep 3 soon so late Yes upgrades are not lost.
    """
    for t in state.get("timers", []):
        t.cancel()

    elapsed = _seconds_since_meeting_start(state.get("start_time", ""))
    timers: list[threading.Timer] = []
    planned: list[str] = []
    for sweep, target_sec in delays:
        remaining = float(target_sec) - elapsed
        if remaining <= 0:
            # Past due: only catch up the hour sweep (late Yes). Skip earlier sweeps.
            if sweep == 3 and remaining > -7200:
                overdue = -remaining
                remaining = 2.0
                planned.append(f"sweep{sweep}=immediate(overdue by {overdue:.0f}s)")
            else:
                sys.stderr.write(
                    f"[{log_prefix}] Skip past checkpoint sweep{sweep} "
                    f"(elapsed={elapsed:.0f}s target={target_sec}s) meeting={meeting_id}\n"
                )
                continue
        else:
            planned.append(f"sweep{sweep}=T+{remaining:.0f}s")

        timer = threading.Timer(remaining, sweep_fn, args=[meeting_id, sweep])
        timer.daemon = True
        timer.start()
        timers.append(timer)

    state["timers"] = timers
    if timers:
        sys.stderr.write(
            f"[{log_prefix}] Checkpoints scheduled (elapsed={elapsed:.0f}s): "
            f"{', '.join(planned)} meeting={meeting_id}\n"
        )
    else:
        sys.stderr.write(
            f"[{log_prefix}] No checkpoint timers left (elapsed={elapsed:.0f}s) "
            f"meeting={meeting_id}\n"
        )


def _mc_schedule_checkpoints(meeting_id: str, state: dict) -> None:
    _schedule_checkpoint_timers(
        meeting_id,
        state,
        sweep_fn=_mc_sweep,
        delays=[
            (1, _MC_CHECKPOINT_1),
            (2, _MC_CHECKPOINT_2),
            (3, _MC_CHECKPOINT_3),
        ],
        log_prefix="started",
    )


def _handle_meeting_started(body: dict, forward_url: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    topic = obj.get("topic", "")
    if not _is_mc_topic(topic):
        sys.stderr.write(f"[started] Not an MC topic (topic={topic!r}) — skipping\n")
        return

    meeting_id = str(obj.get("id", ""))
    start_time = obj.get("start_time", "")
    session_day = _mc_session_day(topic)
    session_date = _session_date_ist(start_time)
    batch = _batch_date(session_date, session_day)

    with _mc_lock:
        existing = _mc_meetings.get(meeting_id)
        roster: dict = {}
        ever_joined: dict = {}
        if existing:
            roster = dict(existing.get("roster", {}))
            ever_joined = dict(existing.get("ever_joined", {}))
            _mc_cancel_timers(existing)

        state = {
            "meeting_id": meeting_id,
            "topic": topic,
            "start_time": start_time,
            "session_date": session_date,
            "batch_date": batch,
            "session_day": session_day,
            "forward_url": forward_url,
            "roster": roster,
            "ever_joined": ever_joined,
            "timers": [],
        }
        _mc_schedule_checkpoints(meeting_id, state)
        _mc_meetings[meeting_id] = state

    sys.stderr.write(
        f"[started] MC {session_day} meeting={meeting_id} "
        f"session={session_date} batch={batch} roster={len(roster)}\n"
    )


def _handle_mc_participant_joined(body: dict, forward_url: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = participant.get("email", "").strip()
    name = participant.get("user_name", "").strip()
    join_time = participant.get("join_time", "")
    topic = obj.get("topic", "")

    if not _is_mc_topic(topic):
        sys.stderr.write(f"[join/mc] Not MC topic (topic={topic!r}) — skipping\n")
        return

    rkey = _mc_roster_key(email, name)
    if not rkey:
        sys.stderr.write("[join/mc] Missing email and name — skipping roster\n")
        return

    with _mc_lock:
        state = _mc_meetings.get(meeting_id)
        if state is None:
            session_day = _mc_session_day(topic)
            session_date = _session_date_ist(join_time or obj.get("start_time", ""))
            state = {
                "meeting_id": meeting_id,
                "topic": topic,
                "start_time": obj.get("start_time", join_time),
                "session_date": session_date,
                "batch_date": _batch_date(session_date, session_day),
                "session_day": session_day,
                "forward_url": forward_url,
                "roster": {},
                "ever_joined": {},
                "timers": [],
            }
            _mc_meetings[meeting_id] = state
            _mc_schedule_checkpoints(meeting_id, state)
            sys.stderr.write(
                f"[join/mc] Roster + checkpoints recovered (no prior state) "
                f"meeting={meeting_id}\n"
            )
        elif not state.get("timers"):
            _mc_schedule_checkpoints(meeting_id, state)
            sys.stderr.write(
                f"[join/mc] Rescheduled empty timers meeting={meeting_id}\n"
            )

        pinfo = {
            "email": email,
            "name": name,
            "join_time": join_time,
        }
        state["roster"][rkey] = pinfo
        state.setdefault("ever_joined", {})[rkey] = pinfo

    sys.stderr.write(f"[join/mc] Roster +1 {email or name} meeting={meeting_id}\n")


def _handle_mc_participant_left(body: dict) -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = participant.get("email", "").strip()
    name = participant.get("user_name", "").strip()
    topic = obj.get("topic", "")

    if not _is_mc_topic(topic):
        return

    rkey = _mc_roster_key(email, name)
    if not rkey:
        return

    with _mc_lock:
        state = _mc_meetings.get(meeting_id)
        if state and rkey in state.get("roster", {}):
            state["roster"].pop(rkey, None)
            sys.stderr.write(
                f"[left/mc] Roster -1 {email or name} meeting={meeting_id} "
                "(no downgrade; checkpoint decides)\n"
            )


def _handle_meeting_ended(body: dict, forward_url: str, program: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    topic = obj.get("topic", "")
    start_time = obj.get("start_time", "")
    meeting_id = str(obj.get("id", ""))

    if program == "100BM":
        if not _is_100bm_topic(topic):
            sys.stderr.write(f"[ended] Not a 100BM session (topic={topic!r}) — skipping\n")
            return
        with _100bm_lock:
            state = _100bm_meetings.pop(meeting_id, None)
            if state:
                _100bm_cancel_timers(state)
        sys.stderr.write(f"[ended] 100BM session ended — posting meeting.ended (meeting={meeting_id})\n")
    else:
        if not _is_mc_topic(topic):
            sys.stderr.write(f"[ended] Not an MC meeting (topic={topic!r}) — skipping\n")
            return
        day_label = _mc_session_day(topic)
        with _mc_lock:
            state = _mc_meetings.pop(meeting_id, None)
            if state:
                _mc_cancel_timers(state)
        sys.stderr.write(f"[ended] MC {day_label} ended — posting meeting.ended (meeting={meeting_id})\n")

    _post_to_zoho(forward_url, {
        "event": "meeting.ended",
        "meeting_id": meeting_id,
        "start_time": start_time,
        "topic": topic,
        "meeting_topic": topic,
        "program": program or "MC",
        "session_date": _session_date_ist(start_time),
    }, "ended")


# --- 100BM checkpoint model (same T+15 / T+30 as MC) ---

def _100bm_base_payload(state: dict) -> dict:
    return {
        "meeting_id": state["meeting_id"],
        "meeting_topic": state["topic"],
        "topic": state["topic"],
        "start_time": state["start_time"],
        "session_date": state["session_date"],
        "batch_date": state["session_date"],
        "program": "100BM",
    }


def _100bm_mark_yes(forward_url: str, state: dict, email: str, name: str, join_time: str) -> None:
    payload = _100bm_base_payload(state)
    payload.update({
        "event": "attendance.mark_yes",
        "participant_email": email,
        "participant_name": name,
        "join_time": join_time,
    })
    _post_to_zoho(forward_url, payload, "100bm-yes")


def _100bm_mark_no(forward_url: str, state: dict, email: str, name: str, join_time: str) -> None:
    """T+30 only: joined at some point but not in room at final checkpoint → No (Yes→No allowed)."""
    payload = _100bm_base_payload(state)
    payload.update({
        "event": "attendance.mark_no",
        "participant_email": email,
        "participant_name": name,
        "join_time": join_time,
    })
    _post_to_zoho(forward_url, payload, "100bm-no")


def _100bm_lookup_participant(forward_url: str, state: dict, email: str, name: str, join_time: str) -> None:
    payload = _100bm_base_payload(state)
    payload.update({
        "event": "attendance.lookup",
        "participant_email": email,
        "participant_name": name,
        "join_time": join_time,
    })
    _post_to_zoho(forward_url, payload, "100bm-lookup")


def _100bm_sweep(meeting_id: str, sweep: int) -> None:
    with _100bm_lock:
        state = _100bm_meetings.get(meeting_id)
        if not state:
            sys.stderr.write(f"[100bm/sweep{sweep}] No state for meeting={meeting_id}\n")
            return
        roster = dict(state["roster"])
        ever_joined = dict(state.get("ever_joined", {}))
        forward_url = state["forward_url"]

    sys.stderr.write(
        f"[100bm/sweep{sweep}] meeting={meeting_id} roster={len(roster)} "
        f"ever_joined={len(ever_joined)} topic={state['topic']!r} session={state['session_date']}\n"
    )

    for info in roster.values():
        email = info.get("email", "")
        name = info.get("name", "")
        if not email and not name:
            continue
        _100bm_mark_yes(forward_url, state, email, name, info.get("join_time", ""))

    if sweep == 1:
        roster_keys = set(roster.keys())
        for rkey, info in ever_joined.items():
            if rkey in roster_keys:
                continue
            email = info.get("email", "")
            name = info.get("name", "")
            if not email and not name:
                continue
            _100bm_lookup_participant(forward_url, state, email, name, info.get("join_time", ""))

    # T+30 only: joined earlier but not in room now → No (downgrades Yes from T+15 if they left)
    if sweep == 2:
        roster_keys = set(roster.keys())
        for rkey, info in ever_joined.items():
            if rkey in roster_keys:
                continue
            email = info.get("email", "")
            name = info.get("name", "")
            if not email and not name:
                continue
            _100bm_mark_no(forward_url, state, email, name, info.get("join_time", ""))

    event = _checkpoint_event(sweep)
    payload = _100bm_base_payload(state)
    payload["event"] = event
    payload["ever_joined_emails"] = _ever_joined_email_csv(ever_joined)
    payload["present_emails"] = _roster_email_csv(roster)
    _post_to_zoho(forward_url, payload, f"100bm-sweep{sweep}")


def _100bm_cancel_timers(state: dict) -> None:
    for t in state.get("timers", []):
        t.cancel()


def _100bm_schedule_checkpoints(meeting_id: str, state: dict) -> None:
    _schedule_checkpoint_timers(
        meeting_id,
        state,
        sweep_fn=_100bm_sweep,
        delays=[
            (1, _BM100_CHECKPOINT_1),
            (2, _BM100_CHECKPOINT_2),
            (3, _BM100_CHECKPOINT_3),
        ],
        log_prefix="100bm/started",
    )


def _handle_100bm_meeting_started(body: dict, forward_url: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    topic = obj.get("topic", "")
    if not _is_100bm_topic(topic):
        sys.stderr.write(f"[100bm/started] Not a 100BM topic (topic={topic!r}) — skipping\n")
        return

    meeting_id = str(obj.get("id", ""))
    start_time = obj.get("start_time", "")
    session_date = _session_date_ist(start_time)

    with _100bm_lock:
        existing = _100bm_meetings.get(meeting_id)
        roster: dict = {}
        ever_joined: dict = {}
        if existing:
            roster = dict(existing.get("roster", {}))
            ever_joined = dict(existing.get("ever_joined", {}))
            _100bm_cancel_timers(existing)

        state = {
            "meeting_id": meeting_id,
            "topic": topic,
            "start_time": start_time,
            "session_date": session_date,
            "forward_url": forward_url,
            "roster": roster,
            "ever_joined": ever_joined,
            "timers": [],
        }
        _100bm_schedule_checkpoints(meeting_id, state)
        _100bm_meetings[meeting_id] = state

    sys.stderr.write(
        f"[100bm/started] meeting={meeting_id} session={session_date} roster={len(roster)}\n"
    )


def _handle_100bm_participant_joined(body: dict, forward_url: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = participant.get("email", "").strip()
    name = participant.get("user_name", "").strip()
    join_time = participant.get("join_time", "")
    topic = obj.get("topic", "")

    if not _is_100bm_topic(topic):
        sys.stderr.write(f"[100bm/join] Not a 100BM topic (topic={topic!r}) — skipping\n")
        return

    rkey = _mc_roster_key(email, name)
    if not rkey:
        sys.stderr.write("[100bm/join] Missing email and name — skipping roster\n")
        return

    with _100bm_lock:
        state = _100bm_meetings.get(meeting_id)
        if state is None:
            # After Render/Railway redeploy, memory is empty — rebuild state and
            # reschedule any checkpoints still in the future (usually T+60).
            session_date = _session_date_ist(join_time or obj.get("start_time", ""))
            state = {
                "meeting_id": meeting_id,
                "topic": topic,
                "start_time": obj.get("start_time", join_time),
                "session_date": session_date,
                "forward_url": forward_url,
                "roster": {},
                "ever_joined": {},
                "timers": [],
            }
            _100bm_meetings[meeting_id] = state
            _100bm_schedule_checkpoints(meeting_id, state)
            sys.stderr.write(
                f"[100bm/join] Roster + checkpoints recovered (no prior state) "
                f"meeting={meeting_id}\n"
            )
        elif not state.get("timers"):
            _100bm_schedule_checkpoints(meeting_id, state)
            sys.stderr.write(
                f"[100bm/join] Rescheduled empty timers meeting={meeting_id}\n"
            )

        pinfo = {
            "email": email,
            "name": name,
            "join_time": join_time,
        }
        state["roster"][rkey] = pinfo
        state.setdefault("ever_joined", {})[rkey] = pinfo

    sys.stderr.write(f"[100bm/join] Roster +1 {email or name} meeting={meeting_id}\n")


def _handle_100bm_participant_left(body: dict) -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = participant.get("email", "").strip()
    name = participant.get("user_name", "").strip()
    topic = obj.get("topic", "")

    if not _is_100bm_topic(topic):
        return

    rkey = _mc_roster_key(email, name)
    if not rkey:
        return

    with _100bm_lock:
        state = _100bm_meetings.get(meeting_id)
        if state and rkey in state.get("roster", {}):
            state["roster"].pop(rkey, None)
            sys.stderr.write(
                f"[100bm/left] Roster -1 {email or name} meeting={meeting_id} "
                "(no downgrade; checkpoint decides)\n"
            )


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


def make_handler(secret: str, forward_url: str, secret_100bm: str = "", forward_url_100bm: str = ""):
    class H(BaseHTTPRequestHandler):
        _secret = secret
        _secret_100bm = secret_100bm
        _forward_url = forward_url
        _forward_url_100bm = forward_url_100bm or forward_url

        def log_message(self, fmt: str, *args) -> None:
            sys.stderr.write(fmt % args + "\n")

        def _route(self) -> tuple[str, str, str] | None:
            path = self.path.split("?", 1)[0].rstrip("/") or "/"
            if path == "/100bm":
                if not self._secret_100bm:
                    return None
                return self._secret_100bm, "100BM", self._forward_url_100bm
            return self._secret, "", self._forward_url

        def do_POST(self) -> None:
            length = int(self.headers.get("Content-Length", "0") or "0")
            raw = self.rfile.read(length) if length else b""

            route = self._route()
            if route is None:
                self.send_response(404)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"error":"/100bm disabled: set ZOOM_WEBHOOK_SECRET_TOKEN_100BM"}')
                return
            route_secret, program, route_forward = route

            try:
                body = json.loads(raw.decode("utf-8"))
            except json.JSONDecodeError:
                self.send_response(400)
                self.end_headers()
                return

            event = body.get("event", "")

            if event == "endpoint.url_validation":
                try:
                    payload = _validation_response(body, route_secret)
                except ValueError as e:
                    self.send_response(500)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"error": str(e)}).encode())
                    return
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return

            if program == "100BM":
                if event == "meeting.started":
                    _handle_100bm_meeting_started(body, route_forward)
                elif event == "meeting.participant_joined":
                    _handle_100bm_participant_joined(body, route_forward)
                elif event == "meeting.participant_left":
                    _handle_100bm_participant_left(body)
                elif event == "meeting.ended":
                    _handle_meeting_ended(body, route_forward, program)
                else:
                    self._forward(raw, route_forward)
            else:
                if event == "meeting.started":
                    _handle_meeting_started(body, route_forward)
                elif event == "meeting.participant_joined":
                    _handle_mc_participant_joined(body, route_forward)
                elif event == "meeting.participant_left":
                    _handle_mc_participant_left(body)
                elif event == "meeting.ended":
                    _handle_meeting_ended(body, route_forward, program)
                else:
                    self._forward(raw, route_forward)

            self._ok()

        def _ok(self) -> None:
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')

        def _forward(self, raw: bytes, forward_url: str) -> None:
            try:
                r = requests.post(
                    forward_url,
                    data=raw,
                    headers={"Content-Type": "application/json"},
                    timeout=60,
                )
                self.send_response(r.status_code if r.status_code < 500 else 502)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"forwarded":true}')
            except requests.RequestException as e:
                self.send_response(502)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())

        def do_HEAD(self) -> None:
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()

        def do_GET(self) -> None:
            with _mc_lock:
                active_mc = len(_mc_meetings)
            with _100bm_lock:
                active_100bm = len(_100bm_meetings)
            resp = json.dumps({
                "ok": True,
                "service": "zoom_webhook_bridge",
                "mc_checkpoint_1_seconds": _MC_CHECKPOINT_1,
                "mc_checkpoint_2_seconds": _MC_CHECKPOINT_2,
                "mc_checkpoint_3_seconds": _MC_CHECKPOINT_3,
                "100bm_checkpoint_1_seconds": _BM100_CHECKPOINT_1,
                "100bm_checkpoint_2_seconds": _BM100_CHECKPOINT_2,
                "100bm_checkpoint_3_seconds": _BM100_CHECKPOINT_3,
                "active_mc_meetings": active_mc,
                "active_100bm_meetings": active_100bm,
                "route_100bm": bool(self._secret_100bm),
            }).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(resp)

    return H


def main() -> None:
    _load_dotenv()
    secret = os.environ.get("ZOOM_WEBHOOK_SECRET_TOKEN", "").strip()
    secret_100bm = os.environ.get("ZOOM_WEBHOOK_SECRET_TOKEN_100BM", "").strip()
    forward = os.environ.get("ZOHO_WEBHOOK_FORWARD_URL", "").strip()
    forward_100bm = os.environ.get("ZOHO_WEBHOOK_FORWARD_URL_100BM", "").strip()

    if not secret:
        print("ERROR: Set ZOOM_WEBHOOK_SECRET_TOKEN", file=sys.stderr)
        sys.exit(1)
    if not forward:
        print("ERROR: Set ZOHO_WEBHOOK_FORWARD_URL", file=sys.stderr)
        sys.exit(1)

    p = argparse.ArgumentParser()
    p.add_argument("--host", default="0.0.0.0")
    default_port = int(os.environ.get("PORT", "8080"))
    p.add_argument("--port", type=int, default=default_port)
    args = p.parse_args()

    handler = make_handler(secret, forward, secret_100bm, forward_100bm)
    server = ThreadingHTTPServer((args.host, args.port), handler)
    route_100bm = "/100bm enabled" if secret_100bm else "/100bm disabled"
    print(
        f"Bridge listening http://{args.host}:{args.port}/\n"
        f"MC checkpoints     : T+{_MC_CHECKPOINT_1}s (first), T+{_MC_CHECKPOINT_2}s (final), T+{_MC_CHECKPOINT_3}s (hour)\n"
        f"100BM checkpoints: T+{_BM100_CHECKPOINT_1}s (first), T+{_BM100_CHECKPOINT_2}s (final), T+{_BM100_CHECKPOINT_3}s (hour)\n"
        f"Day 1 keywords     : {_DAY1_KEYWORDS}\n"
        f"Day 2 keywords     : {_DAY2_KEYWORDS}\n"
        f"100BM keywords     : {_100BM_KEYWORDS}\n"
        f"100BM route        : {route_100bm}\n"
        f"Forward URL        : {forward}\n"
        "MC /               : meeting.started → roster → T+15/T+30/T+60 sweeps\n"
        "100BM /100bm       : meeting.started → roster → T+15/T+30/T+60 sweeps\n"
        "meeting.ended      : MC Completed trigger only (no attendance at end)",
        flush=True,
    )
    server.serve_forever()


if __name__ == "__main__":
    main()
