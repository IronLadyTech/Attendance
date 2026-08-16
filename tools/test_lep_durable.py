"""Unit tests for LEP durable identity + majority (no Redis required)."""
from __future__ import annotations

import io
import json
import os
import socket
import sys
import threading
import time
import unittest

_TOOLS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools")
if _TOOLS not in sys.path:
    sys.path.insert(0, _TOOLS)

from lep_identity import (  # noqa: E402
    build_lep_session_key,
    calculate_final,
    normalize_name,
    safe_first_name_match,
    zoom_account_from_path,
)


class TestNormalizeName(unittest.TestCase):
    def test_titles_and_spaces(self):
        self.assertEqual(normalize_name(" Ms. Piyali   Karmakar "), "piyali karmakar")
        self.assertEqual(normalize_name("PIYALI Karmakar"), "piyali karmakar")
        self.assertEqual(normalize_name("Dr Sruthi Rao"), "sruthi rao")


class TestSessionKey(unittest.TestCase):
    def test_day_collision(self):
        d1 = build_lep_session_key("2026-08-09", "Day 1")
        d2 = build_lep_session_key("2026-08-08", "Day 2")
        self.assertEqual(d1, "lep:2026-08-09:day1")
        self.assertEqual(d2, "lep:2026-08-08:day2")
        self.assertNotEqual(d1, d2)


class TestZoomAccount(unittest.TestCase):
    def test_paths(self):
        self.assertEqual(zoom_account_from_path("/lep"), "zoom1")
        self.assertEqual(zoom_account_from_path("/lep4"), "zoom4")


class TestMajority(unittest.TestCase):
    def test_present_patterns(self):
        self.assertEqual(calculate_final("Present", "Present", "Present"), "Present")
        self.assertEqual(calculate_final("Present", "Present", "Absent"), "Present")
        self.assertEqual(calculate_final("Present", "Absent", "Present"), "Present")
        self.assertEqual(calculate_final("Absent", "Present", "Present"), "Present")

    def test_absent_patterns(self):
        self.assertEqual(calculate_final("Absent", "Absent", "Present"), "Absent")
        self.assertEqual(calculate_final("Present", "Absent", "Absent"), "Absent")
        self.assertEqual(calculate_final("Absent", "Present", "Absent"), "Absent")
        self.assertEqual(calculate_final("Absent", "Absent", "Absent"), "Absent")

    def test_missing_not_absent(self):
        self.assertEqual(calculate_final("Present", "Present", "Missing"), "Present")
        self.assertEqual(calculate_final("Absent", "Absent", "Missing"), "Absent")
        self.assertEqual(calculate_final("Present", "Absent", "Missing"), "Unresolved")
        self.assertEqual(calculate_final("Absent", "Present", "Missing"), "Unresolved")


class TestFirstNameUnique(unittest.TestCase):
    def test_unique(self):
        cohort = [
            {"id": "1", "name": "Piyali Karmakar"},
            {"id": "2", "name": "Manisha Kunwar"},
            {"id": "3", "name": "Sruthi Rao"},
        ]
        r = safe_first_name_match("piyali", cohort)
        self.assertEqual(r["status"], "matched")
        self.assertEqual(r["record"]["id"], "1")
        self.assertEqual(r["method"], "first-name-unique")

    def test_ambiguous(self):
        cohort = [
            {"id": "1", "name": "Piyali Karmakar"},
            {"id": "2", "name": "Piyali Sharma"},
        ]
        r = safe_first_name_match("piyali", cohort)
        self.assertEqual(r["status"], "ambiguous")
        self.assertEqual(len(r["candidates"]), 2)


class TestPiyaliRestartSemantics(unittest.TestCase):
    """C1+C2 Present survive conceptual restart; C3 Present → Final Present."""

    def test_majority_from_persisted_checks(self):
        # Simulate Redis surviving restart: only these three values matter
        c1, c2 = "Present", "Present"
        # RAM cleared — irrelevant
        c3 = "Present"
        self.assertEqual(calculate_final(c1, c2, c3), "Present")

    def test_pad_false_was_the_bug(self):
        # Old bug: missing checks treated as Absent (False)
        # P P + Missing padded as Absent → still Present (2P)
        # But P + Missing+Missing padded → Absent wrongly if only C3 ran after restart
        self.assertEqual(calculate_final("Missing", "Missing", "Present"), "Unresolved")
        # Must NOT become Absent
        self.assertNotEqual(calculate_final("Missing", "Missing", "Present"), "Absent")


class _FakePipe:
    """Records queued commands; exec() counts as one HTTP round-trip."""

    def __init__(self, trips):
        self.trips = trips
        self.cmds: list[str] = []

    def __getattr__(self, name):
        def queue(*a, **k):
            self.cmds.append(name)
            return self
        return queue

    def exec(self):
        self.trips.append(("PIPELINE", list(self.cmds)))
        self.cmds = []
        return []


class _FakeRedis:
    def __init__(self, trips, broken=False):
        self.trips = trips
        self.broken = broken

    def pipeline(self):
        pipe = _FakePipe(self.trips)
        if self.broken:
            pipe.exec = self._boom
        return pipe

    @staticmethod
    def _boom():
        raise RuntimeError("upstash unavailable")

    def get(self, key):
        self.trips.append(("get", key))
        return None

    def __getattr__(self, name):
        def call(*a, **k):
            self.trips.append((name, a[0] if a else ""))
            return None
        return call


class TestJoinRoundTrips(unittest.TestCase):
    """Zoom allows ~3s per webhook; a join must not cost a dozen REST calls."""

    def setUp(self):
        os.environ["LEP_DURABLE"] = "1"
        os.environ["UPSTASH_REDIS_REST_URL"] = "https://fake"
        os.environ["UPSTASH_REDIS_REST_TOKEN"] = "fake"
        os.environ["BRIDGE_STATE_PERSIST"] = "0"
        os.environ["QSTASH_TOKEN"] = ""
        import lep_redis as lr
        self.lr = lr
        self.trips: list[tuple] = []
        lr._redis = _FakeRedis(self.trips)
        lr._redis_init_attempted = True

    def _join(self, bridge, meeting_id="86352877237"):
        bridge._lep_sync_durable_join(
            zoom_account="zoom2",
            meeting_id=meeting_id,
            email="",
            name="Namita Dutt",
            participant={"user_name": "Namita Dutt", "participant_uuid": "pid-1"},
            session_date="2026-08-16",
            session_day="Day 1",
            batch_date="2026-08-16",
            forward_url="https://flow.example/hook",
            topic="IL LEP Sessions - 1 & 2 Days",
            start_time="2026-08-16T03:30:00Z",
        )

    def test_join_is_one_batch(self):
        import zoom_webhook_bridge as bridge
        self.assertTrue(self.lr.durable_lep_enabled())
        self._join(bridge)
        pipelines = [t for t in self.trips if t[0] == "PIPELINE"]
        self.assertEqual(len(pipelines), 1, "all join writes must share one round-trip")
        cmds = pipelines[0][1]
        # session meta + meetings set + meeting meta + current + ever + idmeta
        self.assertGreaterEqual(cmds.count("sadd"), 3, cmds)
        self.assertGreaterEqual(cmds.count("hset"), 3, cmds)
        self.assertLessEqual(len(self.trips), 3, f"join should be <=3 trips: {self.trips}")

    def test_failed_batch_raises_so_zoom_retries(self):
        """A dropped write must surface as 500 — never a silent missed join."""
        import zoom_webhook_bridge as bridge
        self.lr._redis = _FakeRedis(self.trips, broken=True)
        with self.assertRaises(RuntimeError):
            self._join(bridge, meeting_id="999")


class TestWebhookAckResilience(unittest.TestCase):
    """Caller hanging up must not lose work or spam tracebacks."""

    def setUp(self):
        os.environ["BRIDGE_STATE_PERSIST"] = "0"
        os.environ["LEP_DURABLE"] = ""
        import zoom_webhook_bridge as bridge
        self.bridge = bridge
        self.processed: list[str] = []
        self._real_join = bridge._handle_lep_participant_joined
        self._real_left = bridge._handle_lep_participant_left

        def slow_join(body, forward_url, zoom_account="zoom1"):
            time.sleep(0.25)
            self.processed.append(body["payload"]["object"]["id"])

        def failing_left(body, zoom_account="zoom1"):
            raise RuntimeError("simulated redis outage")

        bridge._handle_lep_participant_joined = slow_join
        bridge._handle_lep_participant_left = failing_left

        self.err = io.StringIO()
        self._real_stderr = sys.stderr
        sys.stderr = self.err

        handler = bridge.make_handler("mc", "http://unused", lep_secrets={"/lep2": "s"})
        self.srv = bridge.ThreadingHTTPServer(("127.0.0.1", 0), handler)
        self.port = self.srv.server_address[1]
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        sys.stderr = self._real_stderr
        self.bridge._handle_lep_participant_joined = self._real_join
        self.bridge._handle_lep_participant_left = self._real_left

    def _post(self, event, meeting_id, hang_up=False):
        body = json.dumps({
            "event": event,
            "payload": {"object": {
                "id": meeting_id,
                "topic": "IL LEP Sessions",
                "participant": {"user_name": "Namita Dutt"},
            }},
        }).encode()
        s = socket.create_connection(("127.0.0.1", self.port), timeout=5)
        s.sendall(
            b"POST /lep2 HTTP/1.1\r\nHost: x\r\nContent-Type: application/json\r\n"
            b"Content-Length: " + str(len(body)).encode() + b"\r\n\r\n" + body
        )
        if hang_up:
            s.close()
            return ""
        data = s.recv(4096)
        s.close()
        return data.split(b"\r\n", 1)[0].decode()

    def test_client_hangup_keeps_attendance_and_stays_quiet(self):
        self._post("meeting.participant_joined", "111", hang_up=True)
        time.sleep(0.8)
        log = self.err.getvalue()
        self.assertIn("111", self.processed, "join must commit even if caller hangs up")
        self.assertNotIn("Exception occurred during processing", log)
        self.assertNotIn("socketserver.py", log)
        self.assertIn("hung up before ack", log)

    def test_handler_failure_returns_500_for_zoom_retry(self):
        status = self._post("meeting.participant_left", "333")
        self.assertIn("500", status)
        self.assertIn("handler FAILED", self.err.getvalue())

    def test_normal_request_acks_200(self):
        self.assertIn("200", self._post("meeting.participant_joined", "222"))


if __name__ == "__main__":
    unittest.main()
