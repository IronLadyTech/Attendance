"""
QStash scheduling and signature verification for LEP checkpoints.

Replaces in-memory threading.Timer as the authoritative LEP clock.
"""
from __future__ import annotations

import os
import sys
from datetime import datetime, timedelta, timezone

_IST = timezone(timedelta(hours=5, minutes=30))

# Matches existing bridge offsets from 9:00 AM IST
_LEP_DELAYS_DAY1 = [(1, 900), (2, 23400), (3, 33300), (4, 34200)]
_LEP_DELAYS_DAY2 = [(1, 900), (2, 12600), (3, 26100), (4, 27000)]


def qstash_configured() -> bool:
    return bool(os.environ.get("QSTASH_TOKEN", "").strip())


def public_base_url() -> str:
    return os.environ.get("PUBLIC_BASE_URL", "").strip().rstrip("/")


def _delays(session_day: str) -> list[tuple[int, int]]:
    return _LEP_DELAYS_DAY2 if session_day == "Day 2" else _LEP_DELAYS_DAY1


def _anchor_utc(session_date: str) -> datetime:
    day = datetime.strptime(session_date, "%Y-%m-%d").date()
    anchor = datetime(day.year, day.month, day.day, 9, 0, 0, tzinfo=_IST)
    return anchor.astimezone(timezone.utc)


def ensure_lep_schedule(
    session_key: str,
    *,
    batch_date: str,
    session_day: str,
    session_date: str,
) -> bool:
    """
    Create at most one QStash schedule per LEP session (four Zoom starts → one schedule).

    Returns True if this process created the schedule.
    """
    import lep_redis as lr

    if not lr.durable_lep_enabled():
        return False
    if not qstash_configured():
        sys.stderr.write(
            f"[lep/qstash] QSTASH_TOKEN missing — cannot schedule session={session_key}\n"
        )
        return False
    base = public_base_url()
    if not base:
        sys.stderr.write(
            f"[lep/qstash] PUBLIC_BASE_URL missing — cannot schedule session={session_key}\n"
        )
        return False

    if not lr.try_claim_schedule(session_key):
        sys.stderr.write(
            f"[lep/qstash] schedule already exists session={session_key}\n"
        )
        return False

    try:
        from qstash import QStash
    except ImportError:
        sys.stderr.write("[lep/qstash] qstash package not installed\n")
        lr.clear_schedule_claim(session_key)
        return False

    token = os.environ["QSTASH_TOKEN"].strip()
    client = QStash(token)
    anchor = _anchor_utc(session_date)
    now = datetime.now(timezone.utc)

    try:
        for sweep, offset_sec in _delays(session_day):
            fire_at = anchor + timedelta(seconds=float(offset_sec))
            # If already past, fire soon (idempotent handlers)
            not_before = int(fire_at.timestamp())
            if fire_at <= now:
                not_before = int((now + timedelta(seconds=5)).timestamp())

            if sweep == 4:
                url = f"{base}/internal/lep/final"
                body = {
                    "session_key": session_key,
                    "batch_date": batch_date,
                    "session_day": session_day,
                    "session_date": session_date,
                }
                label = "final"
            else:
                url = f"{base}/internal/lep/checkpoint"
                body = {
                    "session_key": session_key,
                    "check": sweep,
                    "batch_date": batch_date,
                    "session_day": session_day,
                    "session_date": session_date,
                }
                label = f"check{sweep}"

            res = client.message.publish_json(
                url=url,
                body=body,
                not_before=not_before,
            )
            msg_id = getattr(res, "message_id", res)
            sys.stderr.write(
                f"[lep/qstash] scheduled {label} session={session_key} "
                f"not_before={fire_at.astimezone(_IST).isoformat()} id={msg_id}\n"
            )

        lr.mark_schedule_created(session_key)
        sys.stderr.write(f"[lep/qstash] schedule created session={session_key}\n")
        return True
    except Exception as e:
        sys.stderr.write(f"[lep/qstash] schedule FAILED session={session_key}: {e}\n")
        lr.clear_schedule_claim(session_key)
        raise


def verify_qstash_request(signature: str, body: bytes, url: str) -> bool:
    """Verify Upstash-Signature. Returns True if valid or if verification not configured (dev)."""
    current = os.environ.get("QSTASH_CURRENT_SIGNING_KEY", "").strip()
    next_key = os.environ.get("QSTASH_NEXT_SIGNING_KEY", "").strip()
    if not current:
        # Allow local/manual testing without keys; production must set keys.
        if os.environ.get("QSTASH_SKIP_VERIFY", "").strip() in ("1", "true", "yes"):
            sys.stderr.write("[lep/qstash] SKIP verify (QSTASH_SKIP_VERIFY=1)\n")
            return True
        sys.stderr.write("[lep/qstash] signing keys not set — rejecting\n")
        return False
    try:
        from qstash import Receiver
    except ImportError:
        sys.stderr.write("[lep/qstash] qstash package missing — rejecting\n")
        return False

    receiver = Receiver(
        current_signing_key=current,
        next_signing_key=next_key or current,
    )
    try:
        body_text = body.decode("utf-8") if isinstance(body, (bytes, bytearray)) else str(body)
        receiver.verify(
            body=body_text,
            signature=signature,
            url=url,
        )
        return True
    except Exception as e:
        sys.stderr.write(f"[lep/qstash] signature verify failed: {e}\n")
        return False
