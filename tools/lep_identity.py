"""
LEP participant name normalization and identity helpers.

Email is optional (LEP Zoom links usually have no registration).
First-name fallback is allowed only when exactly one cohort candidate matches.
"""
from __future__ import annotations

import re
from typing import Any

_TITLE_RE = re.compile(
    r"^(mr|mrs|ms|miss|dr|prof|rev)\.?\s+",
    re.IGNORECASE,
)


def normalize_name(value: str | None) -> str:
    if not value:
        return ""
    name = value.strip().lower()
    previous = None
    while previous != name:
        previous = name
        name = _TITLE_RE.sub("", name).strip()
    name = re.sub(r"\s+", " ", name)
    name = name.strip(" .,-")
    return name


def build_lep_session_key(batch_date: str, session_day: str) -> str:
    day = "day2" if "day 2" in (session_day or "").lower() else "day1"
    return f"lep:{batch_date}:{day}"


def zoom_account_from_path(path: str) -> str:
    """Map /lep → zoom1, /lep2 → zoom2, …"""
    p = (path or "").rstrip("/") or "/lep"
    if p == "/lep":
        return "zoom1"
    if p == "/lep2":
        return "zoom2"
    if p == "/lep3":
        return "zoom3"
    if p == "/lep4":
        return "zoom4"
    return "zoom1"


def participant_identity(participant_name: str | None, participant_email: str | None) -> str | None:
    """Interim Zoom-side identity before CRM resolve."""
    if participant_email:
        email = participant_email.strip().lower()
        if email and email != "null":
            return f"email:{email}"
    name = normalize_name(participant_name)
    if name:
        return f"name:{name}"
    return None


def crm_identity(record_id: str) -> str:
    return f"crm:{str(record_id).strip()}"


def first_name(normalized: str) -> str:
    if not normalized:
        return ""
    if " " in normalized:
        return normalized.split(" ", 1)[0]
    return normalized


def safe_first_name_match(
    normalized_zoom_name: str,
    cohort_records: list[dict[str, Any]],
) -> dict[str, Any]:
    """
    First-name fallback only when exactly one cohort Name shares that first name.

    cohort_records items need at least: id, name (raw display name).
    """
    zoom_first = first_name(normalized_zoom_name)
    if len(zoom_first) < 4:
        return {"status": "unmatched"}

    candidates: list[dict[str, Any]] = []
    for record in cohort_records:
        crm_name = normalize_name(record.get("name") or "")
        if not crm_name:
            continue
        if first_name(crm_name) == zoom_first:
            candidates.append(record)

    if len(candidates) == 1:
        return {
            "status": "matched",
            "record": candidates[0],
            "method": "first-name-unique",
        }
    if len(candidates) > 1:
        return {
            "status": "ambiguous",
            "candidates": candidates,
            "method": "first-name-ambiguous",
        }
    return {"status": "unmatched"}


def calculate_final(c1: str, c2: str, c3: str) -> str:
    """
    2-of-3 majority. Missing checkpoint is NOT Absent.

    Returns: Present | Absent | Unresolved
    """
    values = [c1, c2, c3]
    present_count = sum(1 for v in values if v == "Present")
    absent_count = sum(1 for v in values if v == "Absent")
    if present_count >= 2:
        return "Present"
    if absent_count >= 2:
        return "Absent"
    return "Unresolved"


def extract_zoom_participant_id(participant: dict) -> str:
    if not isinstance(participant, dict):
        return ""
    for key in (
        "participant_uuid",
        "user_id",
        "participant_user_id",
        "id",
    ):
        val = participant.get(key)
        if val is not None and str(val).strip() and str(val).strip().lower() != "null":
            return str(val).strip()
    return ""
