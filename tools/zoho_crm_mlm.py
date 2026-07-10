"""Zoho CRM Meeting Link Manager lookup for MC planned session start times."""
from __future__ import annotations

import os
import sys
import threading
import time
from datetime import datetime, timezone
from typing import Any
from zoneinfo import ZoneInfo

import requests

_IST = ZoneInfo("Asia/Kolkata")
_MLM_MODULE = "Meeting_Link_Manager"
_MLM_COURSE = "Iron Lady Leadership Masterclass"
_FIELD_DAY1 = "Day_1_Date_and_Time"
_FIELD_DAY2 = "Day_2_Date_and_Time"

_token_lock = threading.Lock()
_token_cache: dict[str, Any] = {"access_token": "", "expires_at": 0.0}


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def _crm_configured() -> bool:
    return bool(
        _env("ZOHO_CRM_CLIENT_ID")
        and _env("ZOHO_CRM_CLIENT_SECRET")
        and _env("ZOHO_CRM_REFRESH_TOKEN")
    )


def _accounts_url() -> str:
    return _env("ZOHO_CRM_ACCOUNTS_URL", "https://accounts.zoho.in")


def _api_base() -> str:
    return _env("ZOHO_CRM_API_URL", "https://www.zohoapis.in/crm/v7").rstrip("/")


def _access_token() -> str:
    with _token_lock:
        now = time.time()
        if _token_cache["access_token"] and now < _token_cache["expires_at"] - 60:
            return _token_cache["access_token"]

        resp = requests.post(
            f"{_accounts_url()}/oauth/v2/token",
            params={
                "refresh_token": _env("ZOHO_CRM_REFRESH_TOKEN"),
                "client_id": _env("ZOHO_CRM_CLIENT_ID"),
                "client_secret": _env("ZOHO_CRM_CLIENT_SECRET"),
                "grant_type": "refresh_token",
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        token = data.get("access_token", "")
        if not token:
            raise RuntimeError(f"Zoho token refresh failed: {data}")

        expires_in = int(data.get("expires_in", 3600))
        _token_cache["access_token"] = token
        _token_cache["expires_at"] = now + expires_in
        return token


def _crm_get(path: str, params: dict | None = None) -> dict:
    headers = {"Authorization": f"Zoho-oauthtoken {_access_token()}"}
    resp = requests.get(f"{_api_base()}{path}", headers=headers, params=params or {}, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _parse_crm_datetime(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        dt = value
    else:
        s = str(value).strip()
        if not s or s in ("null", "-None-"):
            return None
        if s.endswith("Z"):
            try:
                dt = datetime.strptime(s, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
            except ValueError:
                return None
        else:
            try:
                dt = datetime.fromisoformat(s)
            except ValueError:
                return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=_IST)
    return dt.astimezone(timezone.utc)


def _date_matches_session(dt: datetime, session_date: str) -> bool:
    """session_date is yyyy-MM-dd in IST."""
    ist = dt.astimezone(_IST)
    ymd = ist.strftime("%Y-%m-%d")
    dmy = ist.strftime("%d-%m-%Y")
    ddmmyyyy = ist.strftime("%d/%m/%y")
    return session_date in (ymd, dmy) or session_date.replace("-", "/") in ddmmyyyy


def _iso_utc(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def fetch_mc_planned_start(session_day: str, session_date: str) -> str | None:
    """
    Return planned session start (UTC ISO Z) from Meeting Link Manager.

    session_day: "Day 1" or "Day 2"
    session_date: yyyy-MM-dd (IST calendar date of this Zoom session)
    """
    if not _crm_configured():
        return None

    field = _FIELD_DAY1 if session_day == "Day 1" else _FIELD_DAY2
    criteria = f"(Course_Name:equals:{_MLM_COURSE})"

    try:
        for page in range(1, 11):
            data = _crm_get(
                f"/{_MLM_MODULE}/search",
                {"criteria": criteria, "page": page, "per_page": 200},
            )
            records = data.get("data") or []
            if not records:
                break

            for rec in records:
                dt = _parse_crm_datetime(rec.get(field))
                if dt and _date_matches_session(dt, session_date):
                    name = rec.get("Name") or rec.get("Meeting_Link_Manager_Name") or "?"
                    sys.stderr.write(
                        f"[mlm] matched record={name!r} {field}={rec.get(field)!r} "
                        f"session={session_day} date={session_date} → {_iso_utc(dt)}\n"
                    )
                    return _iso_utc(dt)

            if len(records) < 200:
                break

        sys.stderr.write(
            f"[mlm] no Meeting Link Manager row for {session_day} date={session_date}\n"
        )
        return None
    except requests.RequestException as exc:
        sys.stderr.write(f"[mlm] CRM lookup failed: {exc}\n")
        return None


def resolve_checkpoint_anchor(
    session_day: str,
    session_date: str,
    zoom_start_time: str,
) -> tuple[str, str]:
    """
    Returns (checkpoint_anchor_iso, source) where source is 'mlm' or 'zoom'.
    Anchor is used for T+15 / T+30 / T+60 scheduling (not admin early-start).
    """
    if _env("MC_USE_MLM_PLANNED_START", "1") not in ("1", "true", "yes"):
        return zoom_start_time, "zoom"

    planned = fetch_mc_planned_start(session_day, session_date)
    if planned:
        return planned, "mlm"
    return zoom_start_time, "zoom"
