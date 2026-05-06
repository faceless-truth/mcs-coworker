r"""
Tests for archive-title sanitisation (Fix B3).

Pins fix_b3_archive_title.md's contract for `_sanitise_filename_part`:
  - Each Windows-unsafe character (/ \ : * ? " < > |) is replaced with `-`.
  - Multiple whitespace runs collapse to a single space.
  - Leading and trailing whitespace is trimmed.
  - Strings longer than 80 characters truncate to 80.
  - An already-clean title passes through unchanged.

The Haiku call inside `generate_archive_title` is intentionally NOT mocked
here — that test would only verify the mock. End-of-Phase-B smoke test
covers the live call.

Same DB-override dance as the other api_server-importing tests so the
Flask app boots against a temp DB instead of the real one.
"""
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg  # noqa: E402

_OUR_DB_PATH = Path(tempfile.mktemp(suffix=".db"))
cfg.DB_PATH = _OUR_DB_PATH
cfg.init_db()

from api_server import _sanitise_filename_part  # noqa: E402


class SanitiseFilenamePartTests(unittest.TestCase):
    def test_each_windows_unsafe_character_replaced_with_dash(self):
        for bad in ('/', '\\', ':', '*', '?', '"', '<', '>', '|'):
            self.assertEqual(
                _sanitise_filename_part(f"a{bad}b"),
                "a-b",
                msg=f"character {bad!r} should be replaced with '-'",
            )

    def test_multiple_whitespace_runs_collapse_to_single_space(self):
        self.assertEqual(
            _sanitise_filename_part("foo    bar\t\tbaz\n\nqux"),
            "foo bar baz qux",
        )

    def test_leading_and_trailing_whitespace_trimmed(self):
        self.assertEqual(
            _sanitise_filename_part("   hello world   "),
            "hello world",
        )

    def test_truncates_to_80_chars(self):
        long = "a" * 200
        result = _sanitise_filename_part(long)
        self.assertEqual(len(result), 80)
        self.assertEqual(result, "a" * 80)

    def test_truncation_boundary_exactly_80(self):
        eighty = "a" * 80
        self.assertEqual(_sanitise_filename_part(eighty), eighty)

    def test_clean_title_passes_through_unchanged(self):
        clean = "GST Treatment Of Going Concern Sale"
        self.assertEqual(_sanitise_filename_part(clean), clean)

    def test_combined_transformation(self):
        # Realistic worst-case: punctuation, double spaces, leading space,
        # and over-length all in one shot.
        messy = (
            '   Re: "GST Treatment" / Going Concern\\Sale   '
            + "x" * 100
        )
        result = _sanitise_filename_part(messy)
        self.assertEqual(len(result), 80)
        self.assertFalse(result.startswith(" "))
        for bad in ('/', '\\', ':', '*', '?', '"', '<', '>', '|'):
            self.assertNotIn(bad, result)


if __name__ == "__main__":
    unittest.main()
