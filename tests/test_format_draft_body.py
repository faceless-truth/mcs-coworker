"""
Tests for graph_client.format_draft_body
=========================================
Covers the central draft-body wrapper used by every plugin draft path.
"""
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

# Use a temp DB so tests never touch the real user database.
import config as cfg  # noqa: E402

_tmp_db = tempfile.mktemp(suffix=".db")
cfg.DB_PATH = Path(_tmp_db)
cfg.init_db()

from graph_client import (  # noqa: E402
    DEFAULT_DRAFT_FONT_FAMILY,
    DEFAULT_DRAFT_FONT_SIZE,
    format_draft_body,
)
from draft_sanitizer import sanitize_draft_body  # noqa: E402


class TestFormatDraftBody(unittest.TestCase):

    def setUp(self):
        # Restore defaults at the start of each test so cross-test order
        # doesn't matter.
        cfg.set_setting("draft_font_family", DEFAULT_DRAFT_FONT_FAMILY)
        cfg.set_setting("draft_font_size", DEFAULT_DRAFT_FONT_SIZE)
        cfg.set_setting("draft_font_color", "#000000")

    def test_html_body_gets_wrapped_with_inline_style(self):
        out = format_draft_body("<p>Hello</p>")
        self.assertIn('font-family: Aptos, Calibri, sans-serif', out)
        self.assertIn('font-size: 11pt', out)
        self.assertIn('color: #000000', out)
        self.assertIn('data-mcs-style="1"', out)
        self.assertIn('<p>Hello</p>', out)
        # Inline style only — never a <style> block.
        self.assertNotIn('<style', out)

    def test_plain_text_newlines_converted_then_wrapped(self):
        out = format_draft_body("Line one\nLine two")
        self.assertIn('Line one<br>', out)
        self.assertIn('Line two', out)
        self.assertIn('data-mcs-style="1"', out)

    def test_already_wrapped_body_does_not_double_wrap(self):
        once = format_draft_body("<p>Hello</p>")
        twice = format_draft_body(once)
        # Exactly one wrapping div.
        self.assertEqual(twice.count('data-mcs-style="1"'), 1)
        self.assertEqual(once, twice)

    def test_empty_body_returns_empty(self):
        self.assertEqual(format_draft_body(""), "")
        self.assertIsNone(format_draft_body(None))

    def test_full_html_document_is_not_wrapped(self):
        # Wrapping a full <html> doc inside a div produces invalid output.
        body = "<html><body><p>Hi</p></body></html>"
        out = format_draft_body(body)
        self.assertNotIn('data-mcs-style="1"', out)
        self.assertIn('<html>', out)

    def test_custom_font_family_setting_is_honoured(self):
        cfg.set_setting("draft_font_family", "Calibri")
        out = format_draft_body("<p>Hi</p>")
        self.assertIn('font-family: Calibri;', out)
        # Should not still carry the Aptos default.
        self.assertNotIn('Aptos', out)

    def test_invalid_font_size_falls_back_to_default(self):
        cfg.set_setting("draft_font_size", "not-a-size")
        out = format_draft_body("<p>Hi</p>")
        self.assertIn('font-size: 11pt', out)

    def test_valid_px_font_size_is_honoured(self):
        cfg.set_setting("draft_font_size", "14px")
        out = format_draft_body("<p>Hi</p>")
        self.assertIn('font-size: 14px', out)

    def test_format_draft_body_no_longer_strips_dashes_itself(self):
        # Dash handling moved to draft_sanitizer; format_draft_body is now a
        # pure wrapper. Bodies passed in with dashes still in place are
        # wrapped as-is. Production callers always run sanitize_draft_body
        # first (see test_pipeline_replaces_em_dash_with_comma below).
        out = format_draft_body("<p>One — two — three</p>")
        self.assertIn("—", out)
        self.assertIn("data-mcs-style=\"1\"", out)

    def test_pipeline_replaces_em_dash_with_comma(self):
        # End-to-end pipeline as used by GraphClient.create_draft / send_email:
        # sanitiser first (markdown + dash policy), then format_draft_body wraps.
        # Behaviour change vs v2.3 pre-sanitiser: em-dash now becomes ", "
        # (was "-"). En-dash between digits is preserved as a hyphen.
        sanitised = sanitize_draft_body("One — two — three")
        out = format_draft_body(sanitised)
        self.assertNotIn("—", out)
        self.assertIn("One, two, three", out)
        self.assertIn("data-mcs-style=\"1\"", out)

    def test_pipeline_preserves_numeric_range_as_hyphen(self):
        sanitised = sanitize_draft_body("FY2024–2025 carried-forward losses.")
        out = format_draft_body(sanitised)
        self.assertIn("FY2024-2025", out)
        self.assertNotIn("–", out)


if __name__ == "__main__":
    unittest.main()
