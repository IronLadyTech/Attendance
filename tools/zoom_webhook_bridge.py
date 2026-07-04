#!/usr/bin/env python3
"""
Zoom webhook bridge with 30-minute presence timer and Day 2 completion handler.

All outgoing calls go to the single ZOHO_WEBHOOK_FORWARD_URL.
The Zoho Flow branches on the "event" field (and "program" for 100BM):
  - "attendance.mark_yes"       → call mark_attendance_yes
  - "attendance.mark_no"        → call mark_attendance_no
  - "attendance.update_duration"→ optional (no Flow branch required)
  - "meeting.ended"             → set_blank + mark_mc_completed (Day 1 / Day 2)

How it works:
  - meeting.participant_joined  → start a 30-min timer for that person
  - meeting.participant_left    → cancel timer + POST {"event":"attendance.mark_no", ...} to Zoho
  - timer fires (still present) → POST {"event":"attendance.mark_yes", ...} to Zoho
  - meeting.ended (Day 1 or Day 2 topic) → POST {"event":"meeting.ended", ...} to Zoho
  - endpoint.url_validation     → answer Zoom's HMAC challenge
  - all other Zoom events       → forward as-is to Zoho

Multi-account routing:
  The MC Zoom account's app posts to "/" (unchanged). A second Zoom account
  (100BM weekly sessions) posts to "/100bm" — same handlers, but validated
  with its own secret token, filtered to topics matching _100BM_KEYWORDS,
  and every outgoing payload gets "program":"100BM" plus "session_date"
  (IST, YYYY-MM-DD) so the Zoho Flow can match records by invite date.

Environment variables required:
  ZOOM_WEBHOOK_SECRET_TOKEN  — Zoom app Secret Token (webhooks section)
  ZOHO_WEBHOOK_FORWARD_URL   — single Zoho Flow webhook URL (handles all events)

Optional:
  ZOOM_WEBHOOK_SECRET_TOKEN_100BM — Secret Token of the 100BM account's Zoom app;
                               enables the /100bm endpoint (404 when unset)
  PRESENCE_SECONDS           — seconds before marking Yes (default: 1800 = 30 min)
  PORT                       — listen port (default: 8080, Render sets this)
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

# Seconds a participant must be present before attendance is marked Yes
_PRESENCE_SECONDS = int(os.environ.get("PRESENCE_SECONDS", "1800"))

# Topic keywords that identify MC sessions (case-insensitive)
_DAY1_KEYWORDS = ["bhag", "breakthrough actions"]
_DAY2_KEYWORDS = ["art of war", "shameless pitching"]
_MC_MEETING_KEYWORDS = _DAY1_KEYWORDS + _DAY2_KEYWORDS

# Topic keywords that identify 100BM Wednesday sessions on the /100bm route
# (e.g. "Orientation Session- Fast track your Leadership Growth")
_100BM_KEYWORDS = ["orientation session", "fast track your leadership growth"]

_IST = timezone(timedelta(hours=5, minutes=30))


def _is_100bm_topic(topic: str) -> bool:
    topic_l = topic.lower()
    return any(kw in topic_l for kw in _100BM_KEYWORDS)


def _session_date_ist(iso_ts: str) -> str:
    """Zoom timestamps are UTC; Zoho invite dates are IST — convert before matching."""
    try:
        dt = datetime.strptime(iso_ts, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
    except (ValueError, TypeError):
        dt = datetime.now(timezone.utc)
    return dt.astimezone(_IST).strftime("%Y-%m-%d")

# Thread-safe registry: timer_key -> threading.Timer
_timers: dict[str, threading.Timer] = {}
# Tracks wall-clock join time so we can calculate duration on leave
_join_timestamps: dict[str, float] = {}
_timers_lock = threading.Lock()


def _load_dotenv() -> None:
    try:
        from dotenv import load_dotenv
        load_dotenv(os.path.join(_REPO_ROOT, ".env"))
    except ImportError:
        pass


def _validation_response(body: dict, secret: str) -> bytes:
    pt = body.get("payload", {}).get("plainToken")
    if not pt or not secret:
        raise ValueError("missing plainToken or ZOOM_WEBHOOK_SECRET_TOKEN")
    enc = hmac.new(secret.encode("utf-8"), pt.encode("utf-8"), hashlib.sha256).hexdigest()
    return json.dumps({"plainToken": pt, "encryptedToken": enc}).encode("utf-8")


def _timer_key(meeting_id: str, email: str) -> str:
    return f"{meeting_id}:{email.lower().strip()}"


def _post_to_zoho(forward_url: str, payload: dict, label: str) -> None:
    try:
        r = requests.post(forward_url, json=payload, timeout=30)
        sys.stderr.write(f"[{label}] Zoho response: {r.status_code}\n")
    except requests.RequestException as e:
        sys.stderr.write(f"[{label}] ERROR posting to Zoho: {e}\n")


def _mark_attendance(forward_url: str, meeting_id: str, email: str, name: str, topic: str, join_time: str, program: str) -> None:
    """Called by timer thread when participant has been present for PRESENCE_SECONDS."""
    key = _timer_key(meeting_id, email)
    with _timers_lock:
        _timers.pop(key, None)

    sys.stderr.write(f"[timer] {_PRESENCE_SECONDS}s elapsed — marking Yes: email={email} meeting={meeting_id}\n")
    _post_to_zoho(forward_url, {
        "event": "attendance.mark_yes",
        "meeting_id": meeting_id,
        "participant_email": email,
        "participant_name": name,
        "meeting_topic": topic,
        "join_time": join_time,
        "program": program,
        "session_date": _session_date_ist(join_time),
    }, "timer")


def _handle_participant_joined(body: dict, forward_url: str, program: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = participant.get("email", "").strip()
    name = participant.get("user_name", "").strip()
    join_time = participant.get("join_time", "")
    topic = obj.get("topic", "")

    if not email or not meeting_id:
        sys.stderr.write("[join] Missing email or meeting_id — skipping timer\n")
        return

    if program == "100BM" and not _is_100bm_topic(topic):
        sys.stderr.write(f"[join] Not a 100BM session (topic={topic!r}) — skipping timer\n")
        return

    key = _timer_key(meeting_id, email)
    with _timers_lock:
        existing = _timers.pop(key, None)
        if existing:
            existing.cancel()
            sys.stderr.write(f"[join] Cancelled previous timer for {email} (rejoin)\n")

        _join_timestamps[key] = time.time()

        t = threading.Timer(
            _PRESENCE_SECONDS,
            _mark_attendance,
            args=[forward_url, meeting_id, email, name, topic, join_time, program],
        )
        t.daemon = True
        t.start()
        _timers[key] = t

    sys.stderr.write(f"[join] Timer started — {email} must stay {_PRESENCE_SECONDS}s (meeting={meeting_id})\n")


def _handle_participant_left(body: dict, forward_url: str, program: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    participant = obj.get("participant", {})
    meeting_id = str(obj.get("id", ""))
    email = participant.get("email", "").strip()
    name = participant.get("user_name", "").strip()
    topic = obj.get("topic", "")
    leave_time = participant.get("leave_time", "")

    if not email or not meeting_id:
        return

    if program == "100BM" and not _is_100bm_topic(topic):
        sys.stderr.write(f"[left] Not a 100BM session (topic={topic!r}) — skipping\n")
        return

    key = _timer_key(meeting_id, email)
    with _timers_lock:
        timer = _timers.pop(key, None)
        joined_ts = _join_timestamps.pop(key, None)

    duration_seconds = int(time.time() - joined_ts) if joined_ts else 0

    if timer:
        timer.cancel()
        sys.stderr.write(f"[left] Timer cancelled — {email} left before {_PRESENCE_SECONDS}s ({duration_seconds}s) → marking No\n")
        _post_to_zoho(forward_url, {
            "event": "attendance.mark_no",
            "meeting_id": meeting_id,
            "participant_email": email,
            "participant_name": name,
            "meeting_topic": topic,
            "duration_seconds": duration_seconds,
            "program": program,
            "session_date": _session_date_ist(leave_time),
        }, "left-early")
    elif duration_seconds > 0:
        sys.stderr.write(f"[left] Already marked Yes — {email} total {duration_seconds}s → updating duration\n")
        _post_to_zoho(forward_url, {
            "event": "attendance.update_duration",
            "meeting_id": meeting_id,
            "participant_email": email,
            "participant_name": name,
            "meeting_topic": topic,
            "duration_seconds": duration_seconds,
            "program": program,
            "session_date": _session_date_ist(leave_time),
        }, "left-update-duration")
    else:
        sys.stderr.write(f"[left] No active timer for {email} (already marked or never joined)\n")


def _handle_meeting_ended(body: dict, forward_url: str, program: str) -> None:
    obj = body.get("payload", {}).get("object", {})
    topic = obj.get("topic", "")
    start_time = obj.get("start_time", "")
    meeting_id = str(obj.get("id", ""))

    topic_l = topic.lower()
    if program != "100BM":
        is_mc = any(kw in topic_l for kw in _MC_MEETING_KEYWORDS)
        if not is_mc:
            sys.stderr.write(f"[ended] Not an MC meeting (topic={topic!r}) — skipping\n")
            return
        day_label = "Day 2" if any(kw in topic_l for kw in _DAY2_KEYWORDS) else "Day 1"
        sys.stderr.write(f"[ended] {day_label} ended — posting meeting.ended (meeting={meeting_id})\n")
    else:
        if not _is_100bm_topic(topic):
            sys.stderr.write(f"[ended] Not a 100BM session (topic={topic!r}) — skipping\n")
            return
        sys.stderr.write(f"[ended] 100BM session ended — posting meeting.ended (meeting={meeting_id})\n")

    _post_to_zoho(forward_url, {
        "event": "meeting.ended",
        "meeting_id": meeting_id,
        "start_time": start_time,
        "topic": topic,
        "meeting_topic": topic,
        "program": program,
        "session_date": _session_date_ist(start_time),
    }, "ended")


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    """Handles each request in a separate thread so timers can fire concurrently."""
    daemon_threads = True


def make_handler(secret: str, forward_url: str, secret_100bm: str = ""):
    class H(BaseHTTPRequestHandler):
        _secret = secret
        _secret_100bm = secret_100bm
        _forward_url = forward_url

        def log_message(self, fmt: str, *args) -> None:
            sys.stderr.write(fmt % args + "\n")

        def _route(self) -> tuple[str, str] | None:
            """Returns (secret, program) for this request path, or None for 404."""
            path = self.path.split("?", 1)[0].rstrip("/") or "/"
            if path == "/100bm":
                if not self._secret_100bm:
                    return None
                return self._secret_100bm, "100BM"
            return self._secret, ""

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
            route_secret, program = route

            try:
                body = json.loads(raw.decode("utf-8"))
            except json.JSONDecodeError:
                self.send_response(400)
                self.end_headers()
                return

            event = body.get("event", "")

            # Zoom URL validation challenge
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

            # Participant joined → start presence timer
            if event == "meeting.participant_joined":
                _handle_participant_joined(body, self._forward_url, program)
                self._ok()
                return

            # Participant left → cancel timer + mark No
            if event == "meeting.participant_left":
                _handle_participant_left(body, self._forward_url, program)
                self._ok()
                return

            # Meeting ended → MC topics on "/", every meeting on "/100bm"
            if event == "meeting.ended":
                _handle_meeting_ended(body, self._forward_url, program)
                self._ok()
                return

            # All other events → forward as-is to Zoho
            self._forward(raw)

        def _ok(self) -> None:
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')

        def _forward(self, raw: bytes) -> None:
            try:
                r = requests.post(
                    self._forward_url,
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

        def do_GET(self) -> None:
            with _timers_lock:
                active = len(_timers)
            resp = json.dumps({
                "ok": True,
                "service": "zoom_webhook_bridge",
                "active_timers": active,
                "presence_seconds": _PRESENCE_SECONDS,
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

    handler = make_handler(secret, forward, secret_100bm)
    server = ThreadingHTTPServer((args.host, args.port), handler)
    route_100bm = "/100bm enabled (program=100BM)" if secret_100bm else "/100bm disabled (set ZOOM_WEBHOOK_SECRET_TOKEN_100BM)"
    print(
        f"Bridge listening http://{args.host}:{args.port}/\n"
        f"Presence threshold : {_PRESENCE_SECONDS}s ({_PRESENCE_SECONDS // 60}m)\n"
        f"Day 1 keywords     : {_DAY1_KEYWORDS}\n"
        f"Day 2 keywords     : {_DAY2_KEYWORDS}\n"
        f"100BM keywords     : {_100BM_KEYWORDS}\n"
        f"100BM route        : {route_100bm}\n"
        f"Forward URL        : {forward}\n"
        "participant_joined → timer started\n"
        "participant_left   → timer cancelled + POST attendance.mark_no to Zoho\n"
        "timer fires        → POST attendance.mark_yes to Zoho\n"
        "meeting.ended (D1/D2) → POST meeting.ended to Zoho\n"
        "other events       → forwarded as-is to Zoho",
        flush=True,
    )
    server.serve_forever()


if __name__ == "__main__":
    main()
