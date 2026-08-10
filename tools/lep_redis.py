"""
Upstash Redis helpers for durable LEP session / roster / checkpoint state.

Source of truth for LEP checkpoints. Render RAM and /tmp are not authoritative.
"""
from __future__ import annotations

import json
import os
import sys
from typing import Any

from lep_identity import build_lep_session_key

LEP_TTL_SECONDS = int(os.environ.get("LEP_REDIS_TTL_SECONDS", str(7 * 24 * 60 * 60)))

_redis = None
_redis_init_attempted = False


def redis_configured() -> bool:
    return bool(
        os.environ.get("UPSTASH_REDIS_REST_URL", "").strip()
        and os.environ.get("UPSTASH_REDIS_REST_TOKEN", "").strip()
    )


def get_redis():
    """Reusable Upstash Redis client from env (no hard-coded credentials)."""
    global _redis, _redis_init_attempted
    if _redis is not None:
        return _redis
    if _redis_init_attempted:
        return None
    _redis_init_attempted = True
    if not redis_configured():
        return None
    try:
        from upstash_redis import Redis
    except ImportError:
        sys.stderr.write(
            "[lep/redis] upstash_redis not installed — Redis client unavailable\n"
        )
        return None
    # Official: reads UPSTASH_REDIS_REST_URL + UPSTASH_REDIS_REST_TOKEN
    _redis = Redis.from_env()
    return _redis


def durable_lep_enabled() -> bool:
    """
    Attendance durable mode (Redis rosters / QStash) — opt-in.

    Redis env alone only enables connectivity tests; set LEP_DURABLE=1 to
    switch LEP attendance off legacy in-memory timers.
    """
    flag = os.environ.get("LEP_DURABLE", "").strip().lower()
    if flag not in ("1", "true", "yes"):
        return False
    return get_redis() is not None


def run_redis_connectivity_test() -> bool:
    """
    Temporary safe connectivity check.

    SET lep:test:connection = working (EX 60), GET it back.
    On success logs only: [lep/redis-test] working
    Never logs URL or token.
    """
    if not redis_configured():
        sys.stderr.write(
            "[lep/redis-test] skipped — UPSTASH_REDIS_REST_URL/TOKEN not set\n"
        )
        return False
    try:
        r = get_redis()
        if r is None:
            sys.stderr.write("[lep/redis-test] failed — client unavailable\n")
            return False
        key = "lep:test:connection"
        r.set(key, "working", ex=60)
        value = r.get(key)
        if value == "working":
            sys.stderr.write("[lep/redis-test] working\n")
            return True
        sys.stderr.write("[lep/redis-test] failed — unexpected value\n")
        return False
    except Exception as e:
        # Do not include env secrets; exception text from Upstash is usually safe.
        sys.stderr.write(f"[lep/redis-test] failed — {type(e).__name__}\n")
        return False


def _touch(r, *keys: str) -> None:
    for key in keys:
        if key:
            try:
                r.expire(key, LEP_TTL_SECONDS)
            except Exception as e:
                sys.stderr.write(f"[lep/redis] expire error key={key}: {e}\n")


def session_meta_key(session_key: str) -> str:
    return f"{session_key}:meta"


def meetings_key(session_key: str) -> str:
    return f"{session_key}:meetings"


def meeting_current_key(session_key: str, zoom_account: str, meeting_id: str) -> str:
    return f"{session_key}:meeting:{zoom_account}:{meeting_id}:current"


def meeting_ever_joined_key(session_key: str, zoom_account: str, meeting_id: str) -> str:
    return f"{session_key}:meeting:{zoom_account}:{meeting_id}:ever_joined"


def meeting_meta_key(session_key: str, zoom_account: str, meeting_id: str) -> str:
    return f"{session_key}:meeting:{zoom_account}:{meeting_id}:meta"


def participant_map_key(session_key: str, meeting_id: str, zoom_pid: str) -> str:
    return f"{session_key}:meeting:{meeting_id}:pmap:{zoom_pid}"


def checkpoint_present_key(session_key: str, check_no: int) -> str:
    return f"{session_key}:check:{check_no}:present"


def checkpoint_status_key(session_key: str, check_no: int) -> str:
    return f"{session_key}:check:{check_no}:status"


def schedule_created_key(session_key: str) -> str:
    return f"{session_key}:schedule_created"


def final_status_key(session_key: str) -> str:
    return f"{session_key}:final:status"


def lock_key(session_key: str, name: str) -> str:
    return f"{session_key}:lock:{name}"


def ensure_session_meta(
    session_key: str,
    *,
    batch_date: str,
    session_day: str,
    session_date: str,
    topic: str,
    forward_url: str,
) -> None:
    r = get_redis()
    if not r:
        return
    key = session_meta_key(session_key)
    r.hset(
        key,
        values={
            "batch_date": batch_date,
            "session_day": session_day,
            "session_date": session_date,
            "topic": topic or "",
            "forward_url": forward_url or "",
        },
    )
    _touch(r, key)


def get_session_meta(session_key: str) -> dict[str, str]:
    r = get_redis()
    if not r:
        return {}
    raw = r.hgetall(session_meta_key(session_key)) or {}
    return {str(k): str(v) for k, v in raw.items()}


def register_meeting(
    session_key: str,
    zoom_account: str,
    meeting_id: str,
    *,
    topic: str,
    start_time: str,
) -> None:
    r = get_redis()
    if not r:
        return
    mk = meetings_key(session_key)
    member = f"{zoom_account}:{meeting_id}"
    r.sadd(mk, member)
    meta = meeting_meta_key(session_key, zoom_account, meeting_id)
    r.hset(
        meta,
        values={
            "topic": topic or "",
            "start_time": start_time or "",
            "zoom_account": zoom_account,
            "meeting_id": meeting_id,
            "active": "1",
        },
    )
    _touch(r, mk, meta)
    sys.stderr.write(
        f"[lep/redis] register meeting session={session_key} "
        f"zoom={zoom_account} meeting={meeting_id}\n"
    )


def mark_meeting_inactive(session_key: str, zoom_account: str, meeting_id: str) -> None:
    r = get_redis()
    if not r:
        return
    meta = meeting_meta_key(session_key, zoom_account, meeting_id)
    cur = meeting_current_key(session_key, zoom_account, meeting_id)
    r.hset(meta, values={"active": "0"})
    r.delete(cur)
    _touch(r, meta)
    sys.stderr.write(
        f"[lep/redis] meeting inactive session={session_key} "
        f"zoom={zoom_account} meeting={meeting_id}\n"
    )


def roster_join(
    session_key: str,
    zoom_account: str,
    meeting_id: str,
    identity: str,
) -> None:
    r = get_redis()
    if not r or not identity:
        return
    cur = meeting_current_key(session_key, zoom_account, meeting_id)
    ever = meeting_ever_joined_key(session_key, zoom_account, meeting_id)
    r.sadd(cur, identity)
    r.sadd(ever, identity)
    _touch(r, cur, ever)
    sys.stderr.write(
        f"[lep/redis] join session={session_key} zoom={zoom_account} "
        f"meeting={meeting_id} identity={identity}\n"
    )


def roster_leave(
    session_key: str,
    zoom_account: str,
    meeting_id: str,
    identity: str,
) -> None:
    r = get_redis()
    if not r or not identity:
        return
    cur = meeting_current_key(session_key, zoom_account, meeting_id)
    r.srem(cur, identity)
    _touch(r, cur)
    sys.stderr.write(
        f"[lep/redis] leave session={session_key} zoom={zoom_account} "
        f"meeting={meeting_id} identity={identity}\n"
    )


def set_participant_mapping(
    session_key: str,
    meeting_id: str,
    zoom_pid: str,
    crm_record_id: str,
) -> None:
    r = get_redis()
    if not r or not zoom_pid or not crm_record_id:
        return
    key = participant_map_key(session_key, meeting_id, zoom_pid)
    r.set(key, crm_record_id, ex=LEP_TTL_SECONDS)


def get_participant_mapping(
    session_key: str,
    meeting_id: str,
    zoom_pid: str,
) -> str:
    r = get_redis()
    if not r or not zoom_pid:
        return ""
    val = r.get(participant_map_key(session_key, meeting_id, zoom_pid))
    return str(val) if val else ""


def list_meetings(session_key: str) -> list[tuple[str, str]]:
    r = get_redis()
    if not r:
        return []
    members = r.smembers(meetings_key(session_key)) or set()
    out: list[tuple[str, str]] = []
    for m in members:
        s = str(m)
        if ":" not in s:
            continue
        zoom, mid = s.split(":", 1)
        out.append((zoom, mid))
    return out


def build_global_present(session_key: str) -> set[str]:
    """UNION of current rosters across all registered Zoom rooms for this session."""
    r = get_redis()
    if not r:
        return set()
    present: set[str] = set()
    rooms = list_meetings(session_key)
    for zoom, mid in rooms:
        members = r.smembers(meeting_current_key(session_key, zoom, mid)) or set()
        for m in members:
            present.add(str(m))
    sys.stderr.write(
        f"[lep/redis] global_present session={session_key} "
        f"rooms={len(rooms)} count={len(present)}\n"
    )
    return present


def build_global_ever_joined(session_key: str) -> set[str]:
    r = get_redis()
    if not r:
        return set()
    ever: set[str] = set()
    for zoom, mid in list_meetings(session_key):
        members = r.smembers(meeting_ever_joined_key(session_key, zoom, mid)) or set()
        for m in members:
            ever.add(str(m))
    return ever


def checkpoint_state(session_key: str, check_no: int) -> str:
    r = get_redis()
    if not r:
        return ""
    val = r.hget(checkpoint_status_key(session_key, check_no), "state")
    return str(val) if val else ""


def save_checkpoint_snapshot(
    session_key: str,
    check_no: int,
    present: set[str],
    *,
    rooms_seen: int,
) -> None:
    r = get_redis()
    if not r:
        return
    snap = checkpoint_present_key(session_key, check_no)
    status = checkpoint_status_key(session_key, check_no)
    r.delete(snap)
    if present:
        r.sadd(snap, *sorted(present))
    r.hset(
        status,
        values={
            "state": "completed",
            "completed_at": _utc_now(),
            "present_count": str(len(present)),
            "rooms_seen": str(rooms_seen),
        },
    )
    _touch(r, snap, status)
    sys.stderr.write(
        f"[lep/check{check_no}] redis snapshot stored session={session_key} "
        f"count={len(present)} rooms={rooms_seen}\n"
    )


def load_checkpoint_present(session_key: str, check_no: int) -> set[str]:
    r = get_redis()
    if not r:
        return set()
    members = r.smembers(checkpoint_present_key(session_key, check_no)) or set()
    return {str(m) for m in members}


def checkpoint_result_for_identity(
    session_key: str,
    check_no: int,
    identity: str,
    alt_identities: list[str] | None = None,
) -> str:
    """Present | Absent | Missing for one participant at one checkpoint."""
    state = checkpoint_state(session_key, check_no)
    if state != "completed":
        return "Missing"
    present = load_checkpoint_present(session_key, check_no)
    candidates = [identity] + (alt_identities or [])
    for c in candidates:
        if c and c in present:
            return "Present"
    return "Absent"


def try_acquire_lock(session_key: str, name: str, ttl_seconds: int = 120) -> bool:
    r = get_redis()
    if not r:
        return False
    key = lock_key(session_key, name)
    ok = r.set(key, "1", nx=True, ex=ttl_seconds)
    return bool(ok)


def release_lock(session_key: str, name: str) -> None:
    r = get_redis()
    if not r:
        return
    r.delete(lock_key(session_key, name))


def mark_final_status(session_key: str, state: str, extra: dict[str, str] | None = None) -> None:
    r = get_redis()
    if not r:
        return
    key = final_status_key(session_key)
    values = {"state": state, "updated_at": _utc_now()}
    if extra:
        values.update(extra)
    r.hset(key, values=values)
    _touch(r, key)


def final_state(session_key: str) -> str:
    r = get_redis()
    if not r:
        return ""
    val = r.hget(final_status_key(session_key), "state")
    return str(val) if val else ""


def try_claim_schedule(session_key: str) -> bool:
    """SET NX schedule_created=creating. True if this process should create QStash jobs."""
    r = get_redis()
    if not r:
        return False
    key = schedule_created_key(session_key)
    ok = r.set(key, "creating", nx=True, ex=LEP_TTL_SECONDS)
    return bool(ok)


def mark_schedule_created(session_key: str) -> None:
    r = get_redis()
    if not r:
        return
    r.set(schedule_created_key(session_key), "created", ex=LEP_TTL_SECONDS)


def clear_schedule_claim(session_key: str) -> None:
    r = get_redis()
    if not r:
        return
    r.delete(schedule_created_key(session_key))


def schedule_status(session_key: str) -> str:
    r = get_redis()
    if not r:
        return ""
    val = r.get(schedule_created_key(session_key))
    return str(val) if val else ""


def identities_to_present_csvs(identities: set[str]) -> tuple[str, str]:
    """Split crm/email/name identities into email CSV and name CSV for Zoho batch."""
    emails: list[str] = []
    names: list[str] = []
    for ident in sorted(identities):
        if ident.startswith("email:"):
            emails.append(ident[6:])
        elif ident.startswith("name:"):
            names.append(ident[5:])
        # crm:* alone cannot go to Deluge present lists without a CRM→name map;
        # checkpoint Zoho path should also send display names when available.
    return ",".join(emails), ",".join(names)


def store_identity_display(session_key: str, identity: str, display_name: str, email: str) -> None:
    """Remember display name/email for an identity so Zoho batch can use names."""
    r = get_redis()
    if not r or not identity:
        return
    key = f"{session_key}:idmeta:{identity}"
    r.hset(
        key,
        values={
            "display_name": display_name or "",
            "email": email or "",
        },
    )
    _touch(r, key)


def load_identity_display(session_key: str, identity: str) -> dict[str, str]:
    r = get_redis()
    if not r or not identity:
        return {}
    raw = r.hgetall(f"{session_key}:idmeta:{identity}") or {}
    return {str(k): str(v) for k, v in raw.items()}


def present_lists_for_zoho(session_key: str, present: set[str]) -> tuple[str, str]:
    """Build present_emails + present_names CSVs from snapshot identities + idmeta."""
    emails: list[str] = []
    names: list[str] = []
    for ident in sorted(present):
        meta = load_identity_display(session_key, ident)
        email = meta.get("email", "")
        name = meta.get("display_name", "")
        if ident.startswith("email:") and not email:
            email = ident[6:]
        if ident.startswith("name:") and not name:
            name = ident[5:]
        if email:
            emails.append(email)
        if name:
            names.append(name)
    return ",".join(emails), ",".join(names)


def _utc_now() -> str:
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).isoformat()


def session_key_for(batch_date: str, session_day: str) -> str:
    return build_lep_session_key(batch_date, session_day)
