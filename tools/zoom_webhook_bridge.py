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

LEP route (POST /lep|/lep2|/lep3|/lep4) — IL LEP Sessions (9:00 AM IST anchor):
  Durable mode (UPSTASH_REDIS_REST_URL set):
  - join/leave → Redis per-room roster (four rooms UNIONed at checkpoint)
  - QStash schedules one Check1/2/3 + Final per session_key=lep:{batch}:{day}
  - POST /internal/lep/checkpoint|final → global Present → Zoho (never per-room batch Absent)
  - Final majority from Redis C1/C2/C3; Missing != Absent
  Legacy mode (no Redis): in-memory timers + per-meeting sweeps (deprecated)

Environment variables required:
  ZOOM_WEBHOOK_SECRET_TOKEN, ZOHO_WEBHOOK_FORWARD_URL

Optional:
  ZOOM_WEBHOOK_SECRET_TOKEN_100BM, ZOHO_WEBHOOK_FORWARD_URL_100BM
  MC_CHECKPOINT_1_SECONDS (default 900 = 15 min)
  MC_CHECKPOINT_2_SECONDS (default 1800 = 30 min)
  MC_CHECKPOINT_3_SECONDS (default 3600 = 60 min)
  MC_USE_MLM_PLANNED_START (default 1) — schedule checkpoints from Meeting Link Manager
  ZOHO_CRM_CLIENT_ID / ZOHO_CRM_CLIENT_SECRET / ZOHO_CRM_REFRESH_TOKEN (for MLM lookup)
  BM100_CHECKPOINT_1_SECONDS (default 900 = 15 min, 100BM route)
  BM100_CHECKPOINT_2_SECONDS (default 1800 = 30 min, 100BM route)
  BM100_CHECKPOINT_3_SECONDS (default 3600 = 60 min, 100BM route)
  ZOOM_WEBHOOK_SECRET_TOKEN_LEP, ZOHO_WEBHOOK_FORWARD_URL_LEP (LEP route /lep)
  Extra LEP Zoom accounts: ZOOM_WEBHOOK_SECRET_TOKEN_LEP_2 → /lep2, _LEP_3 → /lep3, _LEP_4 → /lep4
  LEP day offsets are fixed from 9:00 AM IST anchor (see _LEP_DELAYS_DAY1 / _LEP_DELAYS_DAY2)
  BRIDGE_STATE_PATH (default /tmp/bridge_meetings.json) — optional MC/100BM cache only when LEP durable
  BRIDGE_STATE_PERSIST=0 to disable file persistence
  UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN — durable LEP Redis
  QSTASH_TOKEN / QSTASH_CURRENT_SIGNING_KEY / QSTASH_NEXT_SIGNING_KEY — durable LEP timers
  PUBLIC_BASE_URL — Render public URL for QStash callbacks
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
_TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
if _TOOLS_DIR not in sys.path:
    sys.path.insert(0, _TOOLS_DIR)

from zoho_crm_mlm import resolve_checkpoint_anchor  # noqa: E402
from bridge_state_persist import load_meetings, persist_enabled, persist_path, save_meetings  # noqa: E402
import lep_redis as _lep_redis  # noqa: E402
import lep_qstash as _lep_qstash  # noqa: E402
from lep_checkpoint import execute_checkpoint, execute_final  # noqa: E402
from lep_identity import (  # noqa: E402
    extract_zoom_participant_id,
    participant_identity,
    zoom_account_from_path,
)

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

# LEP meeting topics (any match → LEP route). Override via LEP_TOPIC_KEYWORDS=a,b,c
_LEP_TOPIC_KEYWORDS_DEFAULT = [
    "il lep sessions",
    "il lep",
    "ironlady lep",
    "lep day",
    "2 day session",
    "2-day session",
]


def _lep_topic_keywords() -> list[str]:
    raw = os.environ.get("LEP_TOPIC_KEYWORDS", "").strip()
    if raw:
        return [k.strip().lower() for k in raw.split(",") if k.strip()]
    return list(_LEP_TOPIC_KEYWORDS_DEFAULT)

# LEP checkpoint offsets from 9:00 AM IST on session_date (seconds)
# Day 1: 9:15, 15:30, 18:15, final 18:30
_LEP_DELAYS_DAY1 = [(1, 900), (2, 23400), (3, 33300), (4, 34200)]
# Day 2: 9:15, 12:30, 16:15, final 16:30
_LEP_DELAYS_DAY2 = [(1, 900), (2, 12600), (3, 26100), (4, 27000)]

_IST = timezone(timedelta(hours=5, minutes=30))

# MC per-meeting roster + checkpoint timers
_mc_meetings: dict[str, dict] = {}
_mc_lock = threading.Lock()

# 100BM per-meeting roster + checkpoint timers
_100bm_meetings: dict[str, dict] = {}
_100bm_lock = threading.Lock()

# LEP per-meeting roster + checkpoint timers + check history
_lep_meetings: dict[str, dict] = {}
_lep_lock = threading.Lock()


def _bridge_persist() -> None:
    """Write MC / 100BM / LEP meeting state to disk (survives Render restarts)."""
    if not persist_enabled():
        return
    snap_mc: dict[str, dict] = {}
    snap_100bm: dict[str, dict] = {}
    snap_lep: dict[str, dict] = {}
    with _mc_lock:
        snap_mc = dict(_mc_meetings)
    with _100bm_lock:
        snap_100bm = dict(_100bm_meetings)
    with _lep_lock:
        snap_lep = dict(_lep_meetings)
    save_meetings(snap_mc, snap_100bm, snap_lep)


def _bridge_restore_and_reschedule(
    forward_url: str,
    forward_url_100bm: str,
    forward_url_lep: str,
) -> None:
    """Load persisted meetings and reschedule remaining checkpoint timers."""
    global _mc_meetings, _100bm_meetings, _lep_meetings
    loaded_mc, loaded_100bm, loaded_lep = load_meetings()
    if not loaded_mc and not loaded_100bm and not loaded_lep:
        return

    with _mc_lock:
        for mid, state in loaded_mc.items():
            state["forward_url"] = forward_url
            state["timers"] = []
            _mc_meetings[mid] = state
            _mc_schedule_checkpoints(mid, state)
            sys.stderr.write(
                f"[persist] Restored MC meeting={mid} roster={len(state.get('roster', {}))} "
                f"ever_joined={len(state.get('ever_joined', {}))}\n"
            )

    with _100bm_lock:
        for mid, state in loaded_100bm.items():
            state["forward_url"] = forward_url_100bm or forward_url
            state["timers"] = []
            _100bm_meetings[mid] = state
            _100bm_schedule_checkpoints(mid, state)
            sys.stderr.write(
                f"[persist] Restored 100BM meeting={mid} roster={len(state.get('roster', {}))} "
                f"ever_joined={len(state.get('ever_joined', {}))}\n"
            )

    with _lep_lock:
        for mid, state in loaded_lep.items():
            state["forward_url"] = forward_url_lep or forward_url
            state["timers"] = []
            _lep_meetings[mid] = state
            _lep_schedule_checkpoints(mid, state)
            checks = state.get("check_results", {})
            sys.stderr.write(
                f"[persist] Restored LEP meeting={mid} roster={len(state.get('roster', {}))} "
                f"ever_joined={len(state.get('ever_joined', {}))} check_results={len(checks)}\n"
            )


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


def _is_lep_topic(topic: str) -> bool:
    if not topic:
        return False
    topic_l = topic.lower()
    return any(kw in topic_l for kw in _lep_topic_keywords())


def _lep_session_day(topic: str, session_date: str = "") -> str:
    """Day 1 = Saturday; Day 2 = Sunday. Topic hint overrides weekday."""
    topic_l = topic.lower()
    if "day 2" in topic_l or "day2" in topic_l or "session 2" in topic_l:
        return "Day 2"
    if "day 1" in topic_l or "day1" in topic_l or "session 1" in topic_l:
        return "Day 1"
    if session_date:
        try:
            wd = datetime.strptime(session_date, "%Y-%m-%d").date().weekday()
            if wd == 5:
                return "Day 1"
            if wd == 6:
                return "Day 2"
        except ValueError:
            pass
    return "Day 1"


def _lep_delays(session_day: str) -> list[tuple[int, int]]:
    return _LEP_DELAYS_DAY2 if session_day == "Day 2" else _LEP_DELAYS_DAY1


def _lep_checkpoint_anchor(session_date: str) -> datetime:
    """Planned LEP start = 9:00 AM IST on session_date."""
    day = datetime.strptime(session_date, "%Y-%m-%d").date()
    anchor = datetime(day.year, day.month, day.day, 9, 0, 0, tzinfo=_IST)
    return anchor.astimezone(timezone.utc)


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


def _participant_email(participant: dict) -> str:
    """Zoom may omit email for guests; try known fields when present."""
    if not isinstance(participant, dict):
        return ""
    for key in ("email", "user_email", "participant_email"):
        val = participant.get(key)
        if val is not None and str(val).strip() and str(val).strip().lower() != "null":
            return str(val).strip()
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


def _ever_joined_name_csv(ever_joined: dict) -> str:
    names: list[str] = []
    for info in ever_joined.values():
        n = (info.get("name") or "").strip()
        if n:
            names.append(n)
    return ",".join(names)


def _roster_email_csv(roster: dict) -> str:
    """Emails in the meeting room at this checkpoint (T+30 present list)."""
    emails: list[str] = []
    for info in roster.values():
        e = (info.get("email") or "").strip().lower()
        if e:
            emails.append(e)
    return ",".join(emails)


def _roster_name_csv(roster: dict) -> str:
    names: list[str] = []
    for info in roster.values():
        n = (info.get("name") or "").strip()
        if n:
            names.append(n)
    return ",".join(names)


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


def _apply_checkpoint_anchor(state: dict) -> None:
    """Set checkpoint_anchor from Meeting Link Manager planned time (fallback: Zoom start)."""
    session_day = state.get("session_day", "Day 1")
    session_date = state.get("session_date", "")
    zoom_start = state.get("start_time", "")
    anchor, source = resolve_checkpoint_anchor(session_day, session_date, zoom_start)
    state["planned_start_time"] = anchor
    state["checkpoint_anchor"] = anchor
    state["checkpoint_anchor_source"] = source


def _mc_cancel_timers(state: dict) -> None:
    for t in state.get("timers", []):
        t.cancel()


def _schedule_checkpoint_timers(
    meeting_id: str,
    state: dict,
    *,
    sweep_fn,
    delays: list[tuple[int, int]],
    log_prefix: str,
    use_mlm: bool = True,
) -> None:
    """
    Schedule checkpoints at anchor + offset (T+15 / T+30 / T+60).

    Anchor = Meeting Link Manager planned start when CRM OAuth is configured
    (MC only — MLM has no 100BM records), otherwise Zoom meeting.started.
    Survives redeploys; past first/final are skipped.
    """
    for t in state.get("timers", []):
        t.cancel()

    if not state.get("checkpoint_anchor"):
        if use_mlm:
            _apply_checkpoint_anchor(state)
        else:
            state["checkpoint_anchor"] = state.get("start_time", "")
            state["checkpoint_anchor_source"] = "zoom"

    anchor = _parse_zoom_datetime(state.get("checkpoint_anchor") or state.get("start_time", ""))
    now = datetime.now(timezone.utc)
    source = state.get("checkpoint_anchor_source", "?")
    timers: list[threading.Timer] = []
    planned: list[str] = []

    sys.stderr.write(
        f"[{log_prefix}] checkpoint anchor={anchor.astimezone(_IST).isoformat()} "
        f"source={source} zoom_start={state.get('start_time', '')!r} meeting={meeting_id}\n"
    )

    for sweep, offset_sec in delays:
        fire_at = anchor + timedelta(seconds=float(offset_sec))
        remaining = (fire_at - now).total_seconds()
        if remaining <= 0:
            if sweep >= 3 and remaining > -7200:
                overdue = -remaining
                remaining = 2.0
                planned.append(f"sweep{sweep}=immediate(overdue by {overdue:.0f}s)")
            else:
                sys.stderr.write(
                    f"[{log_prefix}] Skip past checkpoint sweep{sweep} "
                    f"(fire_at={fire_at.astimezone(_IST).isoformat()} "
                    f"offset={offset_sec}s) meeting={meeting_id}\n"
                )
                continue
        else:
            planned.append(f"sweep{sweep}=in {remaining:.0f}s (at +{offset_sec}s from anchor)")

        timer = threading.Timer(remaining, sweep_fn, args=[meeting_id, sweep])
        timer.daemon = True
        timer.start()
        timers.append(timer)

    state["timers"] = timers
    if timers:
        sys.stderr.write(
            f"[{log_prefix}] Checkpoints scheduled: {', '.join(planned)} meeting={meeting_id}\n"
        )
    else:
        sys.stderr.write(
            f"[{log_prefix}] No checkpoint timers left meeting={meeting_id}\n"
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
        _apply_checkpoint_anchor(state)
        _mc_schedule_checkpoints(meeting_id, state)
        _mc_meetings[meeting_id] = state

    sys.stderr.write(
        f"[started] MC {session_day} meeting={meeting_id} "
        f"session={session_date} batch={batch} roster={len(roster)} "
        f"anchor={state.get('checkpoint_anchor_source', '?')}\n"
    )
    _bridge_persist()


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
            _apply_checkpoint_anchor(state)
            _mc_schedule_checkpoints(meeting_id, state)
            sys.stderr.write(
                f"[join/mc] Roster + checkpoints recovered (no prior state) "
                f"meeting={meeting_id}\n"
            )
        elif not state.get("timers"):
            _apply_checkpoint_anchor(state)
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
    _bridge_persist()


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
    _bridge_persist()


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
    _bridge_persist()


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
        use_mlm=False,
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
    _bridge_persist()


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
    _bridge_persist()


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
    _bridge_persist()


# --- LEP checkpoint model (3 samples + majority final) ---

def _lep_majority(present_flags: list[bool]) -> str:
    """2+ Present → Present; tie or 2+ Absent → Absent."""
    while len(present_flags) < 3:
        present_flags.append(False)
    present_count = sum(1 for p in present_flags[:3] if p)
    absent_count = 3 - present_count
    if present_count > absent_count:
        return "Present"
    return "Absent"


def _lep_base_payload(state: dict) -> dict:
    return {
        "meeting_id": state["meeting_id"],
        "meeting_topic": state["topic"],
        "topic": state["topic"],
        "start_time": state["start_time"],
        "session_date": state["session_date"],
        "session_day": state["session_day"],
        "batch_date": state.get("batch_date") or _batch_date(
            state.get("session_date", ""),
            state.get("session_day", "Day 1"),
        ),
        "program": "LEP",
    }


def _lep_mark_check(
    forward_url: str,
    state: dict,
    email: str,
    name: str,
    check_number: int,
    result: str,
) -> None:
    payload = _lep_base_payload(state)
    payload.update({
        "event": "attendance.lep_check",
        "participant_email": email,
        "participant_name": name,
        "check_number": str(check_number),
        "attendance_result": result,
    })
    _post_to_zoho(forward_url, payload, f"lep-check{check_number}")


def _lep_mark_batch_check(
    forward_url: str,
    state: dict,
    check_number: int,
    roster: dict,
    ever_joined: dict,
) -> None:
    """Cohort not in room at this check → Absent (sales call list) + report payload."""
    payload = _lep_base_payload(state)
    payload.update({
        "event": "attendance.lep_batch_check",
        "check_number": str(check_number),
        "present_emails": _roster_email_csv(roster),
        "present_names": _roster_name_csv(roster),
        "ever_joined_emails": _ever_joined_email_csv(ever_joined),
        "ever_joined_names": _ever_joined_name_csv(ever_joined),
    })
    _post_to_zoho(forward_url, payload, f"lep-batch-check{check_number}")


def _lep_mark_final(
    forward_url: str,
    state: dict,
    email: str,
    name: str,
    result: str,
    checks: list[bool],
) -> None:
    payload = _lep_base_payload(state)
    payload.update({
        "event": "attendance.lep_final",
        "participant_email": email,
        "participant_name": name,
        "check_number": "final",
        "attendance_result": result,
        "check_1": "Present" if len(checks) > 0 and checks[0] else "Absent",
        "check_2": "Present" if len(checks) > 1 and checks[1] else "Absent",
        "check_3": "Present" if len(checks) > 2 and checks[2] else "Absent",
    })
    _post_to_zoho(forward_url, payload, "lep-final")


def _lep_mark_final_report(
    forward_url: str,
    state: dict,
    majority_present_emails: list[str],
    ever_joined: dict,
) -> None:
    """One webhook after majority finals so Flow can email a cohort report."""
    emails = sorted({e.strip().lower() for e in majority_present_emails if e and e.strip()})
    payload = _lep_base_payload(state)
    payload.update({
        "event": "attendance.lep_final_report",
        "check_number": "final",
        "present_emails": ",".join(emails),
        "ever_joined_emails": _ever_joined_email_csv(ever_joined),
        "ever_joined_names": _ever_joined_name_csv(ever_joined),
    })
    _post_to_zoho(forward_url, payload, "lep-final-report")


def _lep_mark_final_batch(
    forward_url: str,
    state: dict,
    ever_joined: dict,
) -> None:
    """After majority: mark cohort members who never joined Zoom as Absent."""
    payload = _lep_base_payload(state)
    payload.update({
        "event": "attendance.lep_final_batch",
        "check_number": "final",
        # Treat anyone who ever joined as "present" for skip — do not overwrite to Absent
        "present_emails": _ever_joined_email_csv(ever_joined),
        "present_names": _ever_joined_name_csv(ever_joined),
        "ever_joined_emails": _ever_joined_email_csv(ever_joined),
        "ever_joined_names": _ever_joined_name_csv(ever_joined),
    })
    _post_to_zoho(forward_url, payload, "lep-final-batch")


def _lep_sweep(meeting_id: str, sweep: int) -> None:
    # Durable Redis+QStash mode: per-meeting timers must NOT update Zoho
    # (would mark Absent from a single room). Checkpoints run via /internal/lep/*.
    if _lep_redis.durable_lep_enabled():
        sys.stderr.write(
            f"[lep/sweep{sweep}] durable mode — skip per-room Zoho "
            f"meeting={meeting_id}\n"
        )
        return

    with _lep_lock:
        state = _lep_meetings.get(meeting_id)
        if not state:
            sys.stderr.write(f"[lep/sweep{sweep}] No state for meeting={meeting_id}\n")
            return
        roster = dict(state["roster"])
        ever_joined = dict(state.get("ever_joined", {}))
        check_results: dict = dict(state.get("check_results", {}))
        forward_url = state["forward_url"]

    sys.stderr.write(
        f"[lep/sweep{sweep}] meeting={meeting_id} roster={len(roster)} "
        f"ever_joined={len(ever_joined)} day={state['session_day']} session={state['session_date']}\n"
    )

    if sweep in (1, 2, 3):
        roster_keys = set(roster.keys())
        with _lep_lock:
            st = _lep_meetings.get(meeting_id)
            if not st:
                return
            for rkey, info in ever_joined.items():
                present = rkey in roster_keys
                history = st.setdefault("check_results", {}).setdefault(rkey, [])
                history.append(present)
                email = info.get("email", "")
                name = info.get("name", "")
                result = "Present" if present else "Absent"
                sys.stderr.write(
                    f"[lep/sweep{sweep}] check {email or rkey}: {result}\n"
                )
                if email or name:
                    _lep_mark_check(forward_url, st, email, name, sweep, result)
        _lep_mark_batch_check(forward_url, state, sweep, roster, ever_joined)
        _bridge_persist()
        return

    # sweep 4 — final majority (legacy only)
    with _lep_lock:
        st = _lep_meetings.get(meeting_id)
        if not st:
            return
        ever_joined = dict(st.get("ever_joined", {}))
        check_results = dict(st.get("check_results", {}))
        forward_url = st["forward_url"]
        state = dict(st)

    final_results: list[tuple[str, str, str, list[bool]]] = []
    majority_present_emails: list[str] = []
    for rkey, info in ever_joined.items():
        email = info.get("email", "")
        name = info.get("name", "")
        if not email and not name:
            continue
        checks = list(check_results.get(rkey, []))
        while len(checks) < 3:
            checks.append(False)
        result = _lep_majority(checks)
        final_results.append((email, name, result, checks[:3]))
        if result == "Present" and email:
            majority_present_emails.append(email)
        _lep_mark_final(forward_url, state, email, name, result, checks[:3])

    _lep_mark_final_report(forward_url, state, majority_present_emails, ever_joined)
    _lep_mark_final_batch(forward_url, state, ever_joined)

    sys.stderr.write(
        f"[lep/sweep4] final updates={len(final_results)} "
        f"ever_joined={len(ever_joined)} meeting={meeting_id}\n"
    )
    _bridge_persist()


def _lep_cancel_timers(state: dict) -> None:
    for t in state.get("timers", []):
        t.cancel()


def _lep_schedule_checkpoints(meeting_id: str, state: dict) -> None:
    if _lep_redis.durable_lep_enabled():
        # QStash owns timing — do not start in-process Timers for LEP.
        state["timers"] = []
        sys.stderr.write(
            f"[lep/started] durable mode — skipping in-process timers "
            f"meeting={meeting_id}\n"
        )
        return
    session_day = state.get("session_day", "Day 1")
    session_date = state.get("session_date", "")
    anchor = _lep_checkpoint_anchor(session_date)
    state["checkpoint_anchor"] = anchor.astimezone(timezone.utc).isoformat()
    state["checkpoint_anchor_source"] = "lep_9am"
    _schedule_checkpoint_timers(
        meeting_id,
        state,
        sweep_fn=_lep_sweep,
        delays=_lep_delays(session_day),
        log_prefix="lep/started",
        use_mlm=False,
    )


def _lep_sync_durable_started(
    *,
    zoom_account: str,
    meeting_id: str,
    topic: str,
    start_time: str,
    session_date: str,
    session_day: str,
    batch_date: str,
    forward_url: str,
) -> None:
    if not _lep_redis.durable_lep_enabled():
        return
    sk = _lep_redis.session_key_for(batch_date, session_day)
    _lep_redis.ensure_session_meta(
        sk,
        batch_date=batch_date,
        session_day=session_day,
        session_date=session_date,
        topic=topic,
        forward_url=forward_url,
    )
    _lep_redis.register_meeting(
        sk,
        zoom_account,
        meeting_id,
        topic=topic,
        start_time=start_time,
    )
    sys.stderr.write(f"[lep/session] key={sk}\n")
    try:
        _lep_qstash.ensure_lep_schedule(
            sk,
            batch_date=batch_date,
            session_day=session_day,
            session_date=session_date,
        )
    except Exception as e:
        sys.stderr.write(f"[lep/qstash] ensure schedule error: {e}\n")


def _lep_sync_durable_join(
    *,
    zoom_account: str,
    meeting_id: str,
    email: str,
    name: str,
    participant: dict,
    session_date: str,
    session_day: str,
    batch_date: str,
    forward_url: str,
    topic: str,
    start_time: str,
) -> None:
    if not _lep_redis.durable_lep_enabled():
        return
    sk = _lep_redis.session_key_for(batch_date, session_day)
    _lep_redis.ensure_session_meta(
        sk,
        batch_date=batch_date,
        session_day=session_day,
        session_date=session_date,
        topic=topic,
        forward_url=forward_url,
    )
    _lep_redis.register_meeting(
        sk, zoom_account, meeting_id, topic=topic, start_time=start_time
    )

    zoom_pid = extract_zoom_participant_id(participant)
    mapped_crm = ""
    if zoom_pid:
        mapped_crm = _lep_redis.get_participant_mapping(sk, meeting_id, zoom_pid)

    if mapped_crm:
        identity = f"crm:{mapped_crm}"
        sys.stderr.write(
            f"[lep/identity] zoom_pid={zoom_pid} reuse CRM={mapped_crm}\n"
        )
    else:
        identity = participant_identity(name, email)
        if not identity:
            sys.stderr.write("[lep/identity] no email/name — skip Redis roster\n")
            return

    _lep_redis.roster_join(sk, zoom_account, meeting_id, identity)
    _lep_redis.store_identity_display(sk, identity, name, email)
    try:
        _lep_qstash.ensure_lep_schedule(
            sk,
            batch_date=batch_date,
            session_day=session_day,
            session_date=session_date,
        )
    except Exception as e:
        sys.stderr.write(f"[lep/qstash] ensure schedule error: {e}\n")


def _handle_lep_meeting_started(
    body: dict, forward_url: str, zoom_account: str = "zoom1"
) -> None:
    obj = body.get("payload", {}).get("object", {})
    topic = obj.get("topic", "")
    if not _is_lep_topic(topic):
        sys.stderr.write(f"[lep/started] Not an LEP topic (topic={topic!r}) — skipping\n")
        return

    meeting_id = str(obj.get("id", ""))
    start_time = obj.get("start_time", "")
    session_date = _session_date_ist(start_time)
    session_day = _lep_session_day(topic, session_date)
    batch = _batch_date(session_date, session_day)

    with _lep_lock:
        existing = _lep_meetings.get(meeting_id)
        roster: dict = {}
        ever_joined: dict = {}
        check_results: dict = {}
        if existing:
            roster = dict(existing.get("roster", {}))
            ever_joined = dict(existing.get("ever_joined", {}))
            check_results = dict(existing.get("check_results", {}))
            _lep_cancel_timers(existing)

        state = {
            "meeting_id": meeting_id,
            "topic": topic,
            "start_time": start_time,
            "session_date": session_date,
            "session_day": session_day,
            "batch_date": batch,
            "forward_url": forward_url,
            "zoom_account": zoom_account,
            "roster": roster,
            "ever_joined": ever_joined,
            "check_results": check_results,
            "timers": [],
        }
        _lep_schedule_checkpoints(meeting_id, state)
        _lep_meetings[meeting_id] = state

    _lep_sync_durable_started(
        zoom_account=zoom_account,
        meeting_id=meeting_id,
        topic=topic,
        start_time=start_time,
        session_date=session_date,
        session_day=session_day,
        batch_date=batch,
        forward_url=forward_url,
    )

    sys.stderr.write(
        f"[lep/started] {session_day} meeting={meeting_id} session={session_date} "
        f"batch={batch} zoom={zoom_account} roster={len(roster)} anchor=9:00IST\n"
    )
    _bridge_persist()


def _handle_lep_participant_joined(
    body: dict, forward_url: str, zoom_account: str = "zoom1"
) -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = _participant_email(participant)
    name = participant.get("user_name", "").strip()
    join_time = participant.get("join_time", "")
    topic = obj.get("topic", "")

    if not _is_lep_topic(topic):
        sys.stderr.write(f"[lep/join] Not an LEP topic (topic={topic!r}) — skipping\n")
        return

    rkey = _mc_roster_key(email, name)
    if not rkey:
        sys.stderr.write("[lep/join] Missing email and name — skipping roster\n")
        return

    if not email:
        sys.stderr.write(
            f"[lep/join] Zoom sent no email for {name!r} "
            "(guest / not logged in / not registered) — will match CRM by name\n"
        )

    with _lep_lock:
        state = _lep_meetings.get(meeting_id)
        if state is None:
            session_date = _session_date_ist(join_time or obj.get("start_time", ""))
            session_day = _lep_session_day(topic, session_date)
            state = {
                "meeting_id": meeting_id,
                "topic": topic,
                "start_time": obj.get("start_time", join_time),
                "session_date": session_date,
                "session_day": session_day,
                "batch_date": _batch_date(session_date, session_day),
                "forward_url": forward_url,
                "zoom_account": zoom_account,
                "roster": {},
                "ever_joined": {},
                "check_results": {},
                "timers": [],
            }
            _lep_meetings[meeting_id] = state
            _lep_schedule_checkpoints(meeting_id, state)
            # Durable mode: empty RAM is OK — Redis holds history. Legacy: warn.
            if _lep_redis.durable_lep_enabled():
                sys.stderr.write(
                    f"[lep/join] no prior RAM state meeting={meeting_id} "
                    f"(durable Redis mode — OK)\n"
                )
            else:
                sys.stderr.write(
                    f"[lep/join] Roster + checkpoints recovered (no prior state) "
                    f"meeting={meeting_id}\n"
                )
        elif not state.get("timers") and not _lep_redis.durable_lep_enabled():
            _lep_schedule_checkpoints(meeting_id, state)
            sys.stderr.write(f"[lep/join] Rescheduled empty timers meeting={meeting_id}\n")

        pinfo = {"email": email, "name": name, "join_time": join_time}
        state["roster"][rkey] = pinfo
        state.setdefault("ever_joined", {})[rkey] = pinfo
        state["zoom_account"] = zoom_account

        session_date = state["session_date"]
        session_day = state["session_day"]
        batch_date = state["batch_date"]
        start_time = state.get("start_time", "")

    _lep_sync_durable_join(
        zoom_account=zoom_account,
        meeting_id=meeting_id,
        email=email,
        name=name,
        participant=participant if isinstance(participant, dict) else {},
        session_date=session_date,
        session_day=session_day,
        batch_date=batch_date,
        forward_url=forward_url,
        topic=topic,
        start_time=start_time,
    )

    sys.stderr.write(
        f"[lep/join] Roster +1 {email or name} meeting={meeting_id} "
        f"zoom={zoom_account}\n"
    )
    _bridge_persist()


def _handle_lep_participant_left(body: dict, zoom_account: str = "zoom1") -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = _participant_email(participant)
    name = participant.get("user_name", "").strip()
    topic = obj.get("topic", "")

    if not _is_lep_topic(topic):
        return

    rkey = _mc_roster_key(email, name)
    if not rkey:
        return

    batch_date = ""
    session_day = ""
    with _lep_lock:
        state = _lep_meetings.get(meeting_id)
        if state and rkey in state.get("roster", {}):
            state["roster"].pop(rkey, None)
            sys.stderr.write(
                f"[lep/left] Roster -1 {email or name} meeting={meeting_id}\n"
            )
        if state:
            batch_date = state.get("batch_date", "")
            session_day = state.get("session_day", "")
            zoom_account = state.get("zoom_account") or zoom_account

    if _lep_redis.durable_lep_enabled() and batch_date:
        sk = _lep_redis.session_key_for(batch_date, session_day or "Day 1")
        zoom_pid = extract_zoom_participant_id(
            participant if isinstance(participant, dict) else {}
        )
        identity = ""
        if zoom_pid:
            crm = _lep_redis.get_participant_mapping(sk, meeting_id, zoom_pid)
            if crm:
                identity = f"crm:{crm}"
        if not identity:
            identity = participant_identity(name, email) or ""
        if identity:
            _lep_redis.roster_leave(sk, zoom_account, meeting_id, identity)

    _bridge_persist()


def _handle_lep_meeting_ended(body: dict, zoom_account: str = "zoom1") -> None:
    obj = body.get("payload", {}).get("object", {})
    topic = obj.get("topic", "")
    meeting_id = str(obj.get("id", ""))

    if not _is_lep_topic(topic):
        sys.stderr.write(f"[lep/ended] Not an LEP meeting (topic={topic!r}) — skipping\n")
        return

    batch_date = ""
    session_day = ""
    with _lep_lock:
        state = _lep_meetings.pop(meeting_id, None)
        if state:
            _lep_cancel_timers(state)
            batch_date = state.get("batch_date", "")
            session_day = state.get("session_day", "")
            zoom_account = state.get("zoom_account") or zoom_account

    if _lep_redis.durable_lep_enabled() and batch_date:
        sk = _lep_redis.session_key_for(batch_date, session_day or "Day 1")
        # Clear live roster only — checkpoint snapshots stay until TTL
        _lep_redis.mark_meeting_inactive(sk, zoom_account, meeting_id)

    sys.stderr.write(
        f"[lep/ended] {session_day or '?'} ended — timers cancelled "
        f"(no CRM completion) meeting={meeting_id}\n"
    )
    _bridge_persist()


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


def make_handler(
    secret: str,
    forward_url: str,
    secret_100bm: str = "",
    forward_url_100bm: str = "",
    secret_lep: str = "",
    forward_url_lep: str = "",
    lep_secrets: dict[str, str] | None = None,
):
    """lep_secrets maps path → Zoom secret, e.g. {"/lep": "...", "/lep2": "..."}."""
    _lep_by_path = dict(lep_secrets or {})
    if secret_lep and "/lep" not in _lep_by_path:
        _lep_by_path["/lep"] = secret_lep

    class H(BaseHTTPRequestHandler):
        _secret = secret
        _secret_100bm = secret_100bm
        _lep_secrets = _lep_by_path
        _forward_url = forward_url
        _forward_url_100bm = forward_url_100bm or forward_url
        _forward_url_lep = forward_url_lep or forward_url

        def log_message(self, fmt: str, *args) -> None:
            sys.stderr.write(fmt % args + "\n")

        def _route(self) -> tuple[str, str, str] | None:
            path = self.path.split("?", 1)[0].rstrip("/") or "/"
            if path in self._lep_secrets:
                return self._lep_secrets[path], "LEP", self._forward_url_lep
            if path.startswith("/lep"):
                return None
            if path == "/100bm":
                if not self._secret_100bm:
                    return None
                return self._secret_100bm, "100BM", self._forward_url_100bm
            return self._secret, "", self._forward_url

        def _handle_internal_lep(self, raw: bytes) -> None:
            """QStash callbacks: /internal/lep/checkpoint and /internal/lep/final."""
            path = self.path.split("?", 1)[0].rstrip("/")
            signature = self.headers.get("Upstash-Signature", "") or self.headers.get(
                "upstash-signature", ""
            )
            base = _lep_qstash.public_base_url()
            full_url = f"{base}{path}" if base else path
            if not _lep_qstash.verify_qstash_request(signature, raw, full_url):
                self.send_response(401)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"error":"invalid qstash signature"}')
                return
            try:
                body = json.loads(raw.decode("utf-8") or "{}")
            except json.JSONDecodeError:
                self.send_response(400)
                self.end_headers()
                return

            session_key = (body.get("session_key") or "").strip()
            if not session_key:
                self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"error":"session_key required"}')
                return

            if path.endswith("/checkpoint"):
                try:
                    check_no = int(body.get("check", 0))
                except (TypeError, ValueError):
                    check_no = 0
                result = execute_checkpoint(
                    session_key, check_no, post_to_zoho=_post_to_zoho
                )
            elif path.endswith("/final"):
                result = execute_final(session_key, post_to_zoho=_post_to_zoho)
            else:
                self.send_response(404)
                self.end_headers()
                return

            payload = json.dumps(result).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def do_POST(self) -> None:
            length = int(self.headers.get("Content-Length", "0") or "0")
            raw = self.rfile.read(length) if length else b""

            path = self.path.split("?", 1)[0].rstrip("/") or "/"
            if path in ("/internal/lep/checkpoint", "/internal/lep/final"):
                self._handle_internal_lep(raw)
                return

            route = self._route()
            if route is None:
                if path == "/100bm":
                    err = '{"error":"/100bm disabled: set ZOOM_WEBHOOK_SECRET_TOKEN_100BM"}'
                elif path.startswith("/lep"):
                    err = '{"error":"LEP route disabled: set ZOOM_WEBHOOK_SECRET_TOKEN_LEP (and _LEP_2/_3/_4 for extra accounts)"}'
                else:
                    err = '{"error":"unknown route"}'
                self.send_response(404)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(err.encode())
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

            if program == "LEP":
                zoom_account = zoom_account_from_path(path)
                if event == "meeting.started":
                    _handle_lep_meeting_started(body, route_forward, zoom_account)
                elif event == "meeting.participant_joined":
                    _handle_lep_participant_joined(body, route_forward, zoom_account)
                elif event == "meeting.participant_left":
                    _handle_lep_participant_left(body, zoom_account)
                elif event == "meeting.ended":
                    _handle_lep_meeting_ended(body, zoom_account)
                else:
                    self._forward(raw, route_forward)
            elif program == "100BM":
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
            with _lep_lock:
                active_lep = len(_lep_meetings)
            resp = json.dumps({
                "ok": True,
                "service": "zoom_webhook_bridge",
                "mc_checkpoint_1_seconds": _MC_CHECKPOINT_1,
                "mc_checkpoint_2_seconds": _MC_CHECKPOINT_2,
                "mc_checkpoint_3_seconds": _MC_CHECKPOINT_3,
                "100bm_checkpoint_1_seconds": _BM100_CHECKPOINT_1,
                "100bm_checkpoint_2_seconds": _BM100_CHECKPOINT_2,
                "100bm_checkpoint_3_seconds": _BM100_CHECKPOINT_3,
                "lep_delays_day1": _LEP_DELAYS_DAY1,
                "lep_delays_day2": _LEP_DELAYS_DAY2,
                "active_mc_meetings": active_mc,
                "active_100bm_meetings": active_100bm,
                "active_lep_meetings": active_lep,
                "route_100bm": bool(self._secret_100bm),
                "route_lep": bool(self._lep_secrets),
                "lep_routes": sorted(self._lep_secrets.keys()),
                "lep_durable_redis": _lep_redis.durable_lep_enabled(),
                "lep_qstash_configured": _lep_qstash.qstash_configured(),
                "public_base_url": _lep_qstash.public_base_url() or None,
                "bridge_state_persist": persist_enabled(),
                "bridge_state_path": persist_path(),
                "mlm_planned_start_enabled": os.environ.get("MC_USE_MLM_PLANNED_START", "1") not in ("0", "false", "no"),
                "mlm_crm_configured": bool(
                    os.environ.get("ZOHO_CRM_CLIENT_ID", "").strip()
                    and os.environ.get("ZOHO_CRM_CLIENT_SECRET", "").strip()
                    and os.environ.get("ZOHO_CRM_REFRESH_TOKEN", "").strip()
                ),
            }).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(resp)

    return H


def _load_lep_secrets() -> dict[str, str]:
    """Map /lep, /lep2, /lep3, /lep4 → Zoom Secret Token for each LEP Zoom account."""
    mapping = [
        ("ZOOM_WEBHOOK_SECRET_TOKEN_LEP", "/lep"),
        ("ZOOM_WEBHOOK_SECRET_TOKEN_LEP_2", "/lep2"),
        ("ZOOM_WEBHOOK_SECRET_TOKEN_LEP_3", "/lep3"),
        ("ZOOM_WEBHOOK_SECRET_TOKEN_LEP_4", "/lep4"),
    ]
    out: dict[str, str] = {}
    for env_key, path in mapping:
        val = os.environ.get(env_key, "").strip()
        if val:
            out[path] = val
    return out


def main() -> None:
    _load_dotenv()
    # Temporary Redis connectivity check (does not change LEP attendance).
    _lep_redis.run_redis_connectivity_test()
    secret = os.environ.get("ZOOM_WEBHOOK_SECRET_TOKEN", "").strip()
    secret_100bm = os.environ.get("ZOOM_WEBHOOK_SECRET_TOKEN_100BM", "").strip()
    secret_lep = os.environ.get("ZOOM_WEBHOOK_SECRET_TOKEN_LEP", "").strip()
    lep_secrets = _load_lep_secrets()
    forward = os.environ.get("ZOHO_WEBHOOK_FORWARD_URL", "").strip()
    forward_100bm = os.environ.get("ZOHO_WEBHOOK_FORWARD_URL_100BM", "").strip()
    forward_lep = os.environ.get("ZOHO_WEBHOOK_FORWARD_URL_LEP", "").strip()

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

    _bridge_restore_and_reschedule(forward, forward_100bm, forward_lep)

    handler = make_handler(
        secret,
        forward,
        secret_100bm,
        forward_100bm,
        secret_lep,
        forward_lep,
        lep_secrets=lep_secrets,
    )
    server = ThreadingHTTPServer((args.host, args.port), handler)
    route_100bm = "/100bm enabled" if secret_100bm else "/100bm disabled"
    route_lep = (
        f"enabled {sorted(lep_secrets.keys())}" if lep_secrets else "disabled"
    )
    print(
        f"Bridge listening http://{args.host}:{args.port}/\n"
        f"MC checkpoints     : T+{_MC_CHECKPOINT_1}s (first), T+{_MC_CHECKPOINT_2}s (final), T+{_MC_CHECKPOINT_3}s (hour)\n"
        f"100BM checkpoints: T+{_BM100_CHECKPOINT_1}s (first), T+{_BM100_CHECKPOINT_2}s (final), T+{_BM100_CHECKPOINT_3}s (hour)\n"
        f"LEP checkpoints  : Day1 {_LEP_DELAYS_DAY1} / Day2 {_LEP_DELAYS_DAY2} (9:00 AM IST anchor)\n"
        f"Day 1 keywords     : {_DAY1_KEYWORDS}\n"
        f"Day 2 keywords     : {_DAY2_KEYWORDS}\n"
        f"100BM keywords     : {_100BM_KEYWORDS}\n"
        f"LEP topic keywords : {_lep_topic_keywords()}\n"
        f"100BM route        : {route_100bm}\n"
        f"LEP routes         : {route_lep}\n"
        f"LEP durable        : redis={'on' if _lep_redis.durable_lep_enabled() else 'off'} "
        f"qstash={'on' if _lep_qstash.qstash_configured() else 'off'}\n"
        f"Public base URL    : {_lep_qstash.public_base_url() or '(unset)'}\n"
        f"State persistence  : {'on → ' + persist_path() if persist_enabled() else 'off'}\n"
        f"Forward URL        : {forward}\n"
        "MC /               : meeting.started → roster → T+15/T+30/T+60 sweeps\n"
        "100BM /100bm       : meeting.started → roster → T+15/T+30/T+60 sweeps\n"
        "LEP /lep[/2/3/4]   : legacy in-memory timers (set LEP_DURABLE=1 for Redis)\n"
        "meeting.ended      : MC Completed trigger only (no attendance at end)",
        flush=True,
    )
    server.serve_forever()


if __name__ == "__main__":
    main()
