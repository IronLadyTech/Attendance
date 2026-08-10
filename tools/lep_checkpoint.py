"""
LEP consolidated checkpoint + final majority (Redis-backed).

Four Zoom rooms are UNIONed before any Zoho update.
Final reads Redis C1/C2/C3 only — never Zoho field history, never RAM.
Missing checkpoint != Absent.
"""
from __future__ import annotations

import sys
from typing import Callable

import lep_redis as lr
from lep_identity import calculate_final

PostFn = Callable[[str, dict, str], None]


def execute_checkpoint(
    session_key: str,
    check_no: int,
    *,
    post_to_zoho: PostFn,
) -> dict:
    if check_no not in (1, 2, 3):
        return {"status": "error", "error": "invalid check"}
    if not lr.durable_lep_enabled():
        return {"status": "error", "error": "redis not configured"}

    if lr.checkpoint_state(session_key, check_no) == "completed":
        sys.stderr.write(
            f"[lep/check{check_no}] duplicate ignored session={session_key}\n"
        )
        return {"status": "already_completed"}

    if not lr.try_acquire_lock(session_key, f"check:{check_no}", ttl_seconds=180):
        if lr.checkpoint_state(session_key, check_no) == "completed":
            return {"status": "already_completed"}
        return {"status": "locked"}

    try:
        if lr.checkpoint_state(session_key, check_no) == "completed":
            return {"status": "already_completed"}

        meta = lr.get_session_meta(session_key)
        rooms = lr.list_meetings(session_key)
        global_present = lr.build_global_present(session_key)

        # Snapshot FIRST — Zoho failure must not lose checkpoint truth
        lr.save_checkpoint_snapshot(
            session_key,
            check_no,
            global_present,
            rooms_seen=len(rooms),
        )

        forward_url = meta.get("forward_url", "")
        if forward_url:
            _post_zoho_checkpoint(
                session_key,
                check_no,
                global_present,
                meta,
                post_to_zoho,
                forward_url,
            )
        else:
            sys.stderr.write(
                f"[lep/check{check_no}] no forward_url — Redis kept, Zoho skipped\n"
            )

        sys.stderr.write(
            f"[lep/check{check_no}] session={session_key} rooms={len(rooms)} "
            f"global_present={len(global_present)}\n"
        )
        return {
            "status": "ok",
            "present_count": len(global_present),
            "rooms_seen": len(rooms),
        }
    finally:
        lr.release_lock(session_key, f"check:{check_no}")


def execute_final(
    session_key: str,
    *,
    post_to_zoho: PostFn,
) -> dict:
    if not lr.durable_lep_enabled():
        return {"status": "error", "error": "redis not configured"}

    if lr.final_state(session_key) == "completed":
        sys.stderr.write(f"[lep/final] duplicate ignored session={session_key}\n")
        return {"status": "already_completed"}

    if not lr.try_acquire_lock(session_key, "final", ttl_seconds=300):
        if lr.final_state(session_key) == "completed":
            return {"status": "already_completed"}
        return {"status": "locked"}

    try:
        if lr.final_state(session_key) == "completed":
            return {"status": "already_completed"}

        meta = lr.get_session_meta(session_key)
        forward_url = meta.get("forward_url", "")
        ever = lr.build_global_ever_joined(session_key)

        majority_present: set[str] = set()
        unresolved: list[str] = []
        present_n = absent_n = unresolved_n = 0

        for ident in sorted(ever):
            alts = _alt_identities(session_key, ident)
            c1 = lr.checkpoint_result_for_identity(session_key, 1, ident, alts)
            c2 = lr.checkpoint_result_for_identity(session_key, 2, ident, alts)
            c3 = lr.checkpoint_result_for_identity(session_key, 3, ident, alts)
            final = calculate_final(c1, c2, c3)

            disp = lr.load_identity_display(session_key, ident)
            label = disp.get("display_name") or ident
            sys.stderr.write(
                f"[lep/final] participant={label!r} c1={c1} c2={c2} c3={c3} "
                f"final={final}\n"
            )

            if final == "Unresolved":
                unresolved_n += 1
                unresolved.append(label)
                sys.stderr.write(
                    f"[lep/final] refusing unsafe Absent update due to missing "
                    f"checkpoint participant={label!r}\n"
                )
                continue

            if final == "Present":
                present_n += 1
                majority_present.add(ident)
            else:
                absent_n += 1

            if forward_url:
                _post_individual_final(
                    session_key,
                    ident,
                    final,
                    c1,
                    c2,
                    c3,
                    meta,
                    post_to_zoho,
                    forward_url,
                )

        if forward_url:
            emails, names = lr.present_lists_for_zoho(session_key, ever)
            batch = _base_payload(meta)
            batch.update({
                "event": "attendance.lep_final_batch",
                "check_number": "final",
                "present_emails": emails,
                "present_names": names,
                "ever_joined_emails": emails,
                "ever_joined_names": names,
                "attendance_scope": "GLOBAL",
            })
            post_to_zoho(forward_url, batch, "lep-final-batch")

            maj_emails, maj_names = lr.present_lists_for_zoho(
                session_key, majority_present
            )
            report = _base_payload(meta)
            report.update({
                "event": "attendance.lep_final_report",
                "check_number": "final",
                "present_emails": maj_emails,
                "present_names": maj_names,
                "ever_joined_emails": emails,
                "ever_joined_names": names,
            })
            post_to_zoho(forward_url, report, "lep-final-report")

        lr.mark_final_status(
            session_key,
            "completed",
            {
                "present": str(present_n),
                "absent": str(absent_n),
                "unresolved": str(unresolved_n),
            },
        )
        return {
            "status": "ok",
            "present": present_n,
            "absent": absent_n,
            "unresolved": unresolved_n,
            "unresolved_names": unresolved,
        }
    finally:
        lr.release_lock(session_key, "final")


def _base_payload(meta: dict) -> dict:
    return {
        "meeting_id": "",
        "meeting_topic": meta.get("topic", ""),
        "topic": meta.get("topic", ""),
        "start_time": "",
        "session_date": meta.get("session_date", ""),
        "session_day": meta.get("session_day", "Day 1"),
        "batch_date": meta.get("batch_date", ""),
        "program": "LEP",
    }


def _alt_identities(session_key: str, ident: str) -> list[str]:
    """Allow matching Present set if display was stored under name/email variants."""
    disp = lr.load_identity_display(session_key, ident)
    alts: list[str] = []
    email = (disp.get("email") or "").strip().lower()
    name = (disp.get("display_name") or "").strip()
    if email:
        alts.append(f"email:{email}")
    if name:
        from lep_identity import normalize_name
        nn = normalize_name(name)
        if nn:
            alts.append(f"name:{nn}")
    return alts


def _post_zoho_checkpoint(
    session_key: str,
    check_no: int,
    global_present: set[str],
    meta: dict,
    post_to_zoho: PostFn,
    forward_url: str,
) -> None:
    emails, names = lr.present_lists_for_zoho(session_key, global_present)
    ever = lr.build_global_ever_joined(session_key)
    ever_emails, ever_names = lr.present_lists_for_zoho(session_key, ever)

    for ident in sorted(global_present):
        disp = lr.load_identity_display(session_key, ident)
        email = disp.get("email", "")
        name = disp.get("display_name", "")
        if ident.startswith("email:") and not email:
            email = ident[6:]
        if ident.startswith("name:") and not name:
            name = ident[5:]
        if not email and not name:
            continue
        payload = _base_payload(meta)
        payload.update({
            "event": "attendance.lep_check",
            "participant_email": email,
            "participant_name": name,
            "check_number": str(check_no),
            "attendance_result": "Present",
            "attendance_scope": "GLOBAL",
        })
        post_to_zoho(forward_url, payload, f"lep-check{check_no}")

    batch = _base_payload(meta)
    batch.update({
        "event": "attendance.lep_batch_check",
        "check_number": str(check_no),
        "present_emails": emails,
        "present_names": names,
        "ever_joined_emails": ever_emails,
        "ever_joined_names": ever_names,
        "attendance_scope": "GLOBAL",
    })
    post_to_zoho(forward_url, batch, f"lep-batch-check{check_no}")
    sys.stderr.write(
        f"[lep/check{check_no}] zoho present={len(global_present)}\n"
    )


def _post_individual_final(
    session_key: str,
    ident: str,
    result: str,
    c1: str,
    c2: str,
    c3: str,
    meta: dict,
    post_to_zoho: PostFn,
    forward_url: str,
) -> None:
    disp = lr.load_identity_display(session_key, ident)
    email = disp.get("email", "")
    name = disp.get("display_name", "")
    if ident.startswith("email:") and not email:
        email = ident[6:]
    if ident.startswith("name:") and not name:
        name = ident[5:]
    if not email and not name:
        return

    def _zoho_check_label(v: str) -> str:
        # Payload fields are informational; do not invent Absent for Missing in majority.
        if v == "Present":
            return "Present"
        if v == "Absent":
            return "Absent"
        return "Missing"

    payload = _base_payload(meta)
    payload.update({
        "event": "attendance.lep_final",
        "participant_email": email,
        "participant_name": name,
        "check_number": "final",
        "attendance_result": result,
        "check_1": _zoho_check_label(c1),
        "check_2": _zoho_check_label(c2),
        "check_3": _zoho_check_label(c3),
        "attendance_scope": "GLOBAL",
    })
    post_to_zoho(forward_url, payload, "lep-final")
