#!/usr/bin/env python3
"""
Zoom webhook bridge: answers Zoom URL validation, then forwards real events to Zoho Flow.

Why: Zoom validates with event "endpoint.url_validation" and expects encryptedToken (HMAC).
    A plain Zoho Flow webhook URL often cannot return that, so Zoom validation fails.

Setup:
  1. In Zoom app → Feature → copy **Secret Token** for Event subscriptions.
  2. Set environment variables:
       ZOOM_WEBHOOK_SECRET_TOKEN = <that secret>
       ZOHO_WEBHOOK_FORWARD_URL   = full Zoho incoming webhook URL (with zapikey)
  3. Run this server on a **public HTTPS** URL (deploy to Render/Railway/Azure, or test with ngrok).
  4. Paste **this server's HTTPS URL** into Zoom "Event notification endpoint URL", not Zoho's URL.

Example (local test with ngrok):
  Copy env.zoom-zoho.example to .env and fill values, or:
  set ZOOM_WEBHOOK_SECRET_TOKEN=xxx
  set ZOHO_WEBHOOK_FORWARD_URL=https://flow.zoho.in/.../incoming?zapikey=...
  pip install python-dotenv
  python tools/zoom_webhook_bridge.py --port 8080
  ngrok http 8080
  Use https://xxxx.ngrok-free.app in Zoom (must be HTTPS).

Requires: requests (pip install requests)
"""
from __future__ import annotations

import argparse
import hashlib
import hmac
import json
import os
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer

import requests

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


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
    out = {"plainToken": pt, "encryptedToken": enc}
    return json.dumps(out).encode("utf-8")


def make_handler(secret: str, forward_url: str):
    class H(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args) -> None:
            sys.stderr.write(fmt % args + "\n")

        def do_POST(self) -> None:
            length = int(self.headers.get("Content-Length", "0") or "0")
            raw = self.rfile.read(length) if length else b""

            try:
                body = json.loads(raw.decode("utf-8"))
            except json.JSONDecodeError:
                self.send_response(400)
                self.end_headers()
                return

            if body.get("event") == "endpoint.url_validation":
                try:
                    payload = _validation_response(body, secret)
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

            # Forward real Zoom events to Zoho Flow
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

        def do_GET(self) -> None:
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"ok":true,"service":"zoom_webhook_bridge"}')

    return H


def main() -> None:
    _load_dotenv()
    secret = os.environ.get("ZOOM_WEBHOOK_SECRET_TOKEN", "").strip()
    forward = os.environ.get("ZOHO_WEBHOOK_FORWARD_URL", "").strip()
    if not secret:
        print("ERROR: Set ZOOM_WEBHOOK_SECRET_TOKEN (Zoom app Secret Token for webhooks).", file=sys.stderr)
        sys.exit(1)
    if not forward:
        print("ERROR: Set ZOHO_WEBHOOK_FORWARD_URL (full Zoho Flow incoming webhook URL).", file=sys.stderr)
        sys.exit(1)

    p = argparse.ArgumentParser()
    p.add_argument("--host", default="0.0.0.0")
    default_port = int(os.environ.get("PORT", "8080"))
    p.add_argument("--port", type=int, default=default_port)
    args = p.parse_args()

    handler = make_handler(secret, forward)
    server = HTTPServer((args.host, args.port), handler)
    print(
        f"Bridge listening http://{args.host}:{args.port}/\n"
        "Put your PUBLIC HTTPS URL (not this raw HTTP in prod) in Zoom.\n"
        "Zoom validation -> encryptedToken OK; other events -> forwarded to Zoho.",
        flush=True,
    )
    server.serve_forever()


if __name__ == "__main__":
    main()
