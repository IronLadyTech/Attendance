"""
Persist MC / 100BM / LEP in-memory meeting state to disk so Render restarts
do not lose roster, ever_joined, or LEP check_results.

Set BRIDGE_STATE_PATH (default /tmp/bridge_meetings.json).
Set BRIDGE_STATE_PERSIST=0 to disable.
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timedelta, timezone

_IST = timezone(timedelta(hours=5, minutes=30))

_PERSIST_ENABLED = os.environ.get("BRIDGE_STATE_PERSIST", "1").strip().lower() not in ("0", "false", "no")
_STATE_PATH = os.environ.get("BRIDGE_STATE_PATH", "/tmp/bridge_meetings.json").strip()


def persist_path() -> str:
    return _STATE_PATH


def persist_enabled() -> bool:
    return _PERSIST_ENABLED


def _is_stale_session(session_date: str) -> bool:
    if not session_date:
        return True
    try:
        sd = datetime.strptime(session_date[:10], "%Y-%m-%d").date()
        today = datetime.now(_IST).date()
        return (today - sd).days > 1
    except ValueError:
        return True


def serialize_meeting_state(state: dict) -> dict:
    """Copy meeting state without threading.Timer objects."""
    out: dict = {}
    for key, val in state.items():
        if key == "timers":
            continue
        if key == "checkpoint_anchor" and val is not None and hasattr(val, "isoformat"):
            out[key] = val.isoformat()
        else:
            out[key] = val
    return out


def save_meetings(
    mc: dict[str, dict],
    bm100: dict[str, dict],
    lep: dict[str, dict],
) -> None:
    if not _PERSIST_ENABLED:
        return
    payload = {
        "version": 1,
        "saved_at": datetime.now(timezone.utc).isoformat(),
        "mc": {mid: serialize_meeting_state(st) for mid, st in mc.items()},
        "100bm": {mid: serialize_meeting_state(st) for mid, st in bm100.items()},
        "lep": {mid: serialize_meeting_state(st) for mid, st in lep.items()},
    }
    path = _STATE_PATH
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, separators=(",", ":"))
        os.replace(tmp, path)
    except OSError as e:
        sys.stderr.write(f"[persist] ERROR writing {path}: {e}\n")


def load_meetings() -> tuple[dict[str, dict], dict[str, dict], dict[str, dict]]:
    if not _PERSIST_ENABLED or not os.path.isfile(_STATE_PATH):
        return {}, {}, {}
    try:
        with open(_STATE_PATH, encoding="utf-8") as f:
            payload = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        sys.stderr.write(f"[persist] ERROR reading {_STATE_PATH}: {e}\n")
        return {}, {}, {}

    mc: dict[str, dict] = {}
    bm100: dict[str, dict] = {}
    lep: dict[str, dict] = {}

    for mid, st in (payload.get("mc") or {}).items():
        if _is_stale_session(st.get("session_date", "")):
            continue
        st["timers"] = []
        mc[str(mid)] = st

    for mid, st in (payload.get("100bm") or {}).items():
        if _is_stale_session(st.get("session_date", "")):
            continue
        st["timers"] = []
        bm100[str(mid)] = st

    for mid, st in (payload.get("lep") or {}).items():
        if _is_stale_session(st.get("session_date", "")):
            continue
        st["timers"] = []
        st.setdefault("check_results", {})
        st.setdefault("ever_joined", {})
        st.setdefault("roster", {})
        lep[str(mid)] = st

    sys.stderr.write(
        f"[persist] Loaded mc={len(mc)} 100bm={len(bm100)} lep={len(lep)} from {_STATE_PATH}\n"
    )
    return mc, bm100, lep
