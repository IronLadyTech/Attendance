"""Unit tests for LEP durable identity + majority (no Redis required)."""
from __future__ import annotations

import os
import sys
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


if __name__ == "__main__":
    unittest.main()
