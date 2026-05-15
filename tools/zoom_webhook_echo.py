#!/usr/bin/env python3
"""
Local helper to inspect Zoom webhook HTTP requests before you point Zoom at Zoho Flow.

Usage:
  python tools/zoom_webhook_echo.py
  python tools/zoom_webhook_echo.py --port 8765

Then expose with ngrok (optional):  ngrok http 8765
Use the ngrok HTTPS URL + path /zoom in Zoom's webhook config for testing only.

Security: do not run this on a public server without auth — debug tool only.
"""
from __future__ import annotations

import argparse
import json
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write(fmt % args + "\n")

    def _send(self, code: int, body: bytes = b"") -> None:
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        print(f"[GET] path={self.path}", flush=True)
        self._send(200, b"{\"ok\":true}")

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length) if length else b""
        print(f"[POST] path={self.path}", flush=True)
        print(self.headers, flush=True)
        try:
            print(json.dumps(json.loads(raw.decode("utf-8", errors="replace")), indent=2), flush=True)
        except Exception:
            print(raw[:8000], flush=True)
        self._send(200, b"{}")


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=8765)
    args = p.parse_args()
    server = HTTPServer((args.host, args.port), Handler)
    print(f"Listening http://{args.host}:{args.port}/ — POST Zoom payloads here for debugging.", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
