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

    def test_typographic_punctuation_still_stripped(self):
        # Existing safety net behaviour must survive the rename.
        out = format_draft_body("<p>One — two — three</p>")
        self.assertNotIn("—", out)
        self.assertIn("One - two - three", out)


if __name__ == "__main__":
    unittest.main()
