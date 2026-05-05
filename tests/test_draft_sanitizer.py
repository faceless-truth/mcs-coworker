"""
Tests for draft_sanitizer
=========================
Covers the markdown→HTML and dash/quote normalisation that runs on every
outgoing draft body before it is wrapped by graph_client.format_draft_body.
"""
import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

from draft_sanitizer import (  # noqa: E402
    looks_like_already_html,
    sanitize_draft_body,
    strip_typographic_punctuation,
)


class TestStripTypographicPunctuation(unittest.TestCase):

    def test_em_dash_with_spaces_becomes_comma(self):
        out = strip_typographic_punctuation(
            "The redraw balance — your own funds — cannot be deducted."
        )
        self.assertNotIn("—", out)
        self.assertIn(
            "redraw balance, your own funds, cannot be deducted", out
        )

    def test_em_dash_no_spaces_becomes_comma(self):
        out = strip_typographic_punctuation(
            "Sale price—less cost base—equals capital gain."
        )
        self.assertNotIn("—", out)
        self.assertIn("Sale price, less cost base, equals capital gain", out)

    def test_en_dash_in_numeric_range_becomes_hyphen(self):
        out = strip_typographic_punctuation("FY2024–2025 carried-forward losses.")
        self.assertIn("FY2024-2025", out)
        self.assertNotIn("–", out)

    def test_en_dash_in_prose_becomes_comma(self):
        out = strip_typographic_punctuation("Quick note – see attached.")
        self.assertNotIn("–", out)
        self.assertIn("Quick note, see attached", out)

    def test_html_entity_dashes_handled(self):
        out = strip_typographic_punctuation("Foo &mdash; bar &ndash; baz.")
        self.assertNotIn("&mdash;", out)
        self.assertNotIn("&ndash;", out)
        self.assertIn("Foo, bar, baz", out)

    def test_smart_quotes_normalised(self):
        out = strip_typographic_punctuation("“Hello” she said ‘hi’.")
        self.assertEqual(out, '"Hello" she said \'hi\'.')

    def test_doubled_punctuation_tidied(self):
        # "word — ." should not leave a stranded ", ."
        out = strip_typographic_punctuation("End of sentence — .")
        self.assertNotIn(", .", out)
        self.assertTrue(out.endswith("."))

    def test_idempotent(self):
        once = strip_typographic_punctuation(
            "FY2024–2025 — see “report”."
        )
        twice = strip_typographic_punctuation(once)
        self.assertEqual(once, twice)

    def test_empty_input_returned_unchanged(self):
        self.assertEqual(strip_typographic_punctuation(""), "")
        self.assertIsNone(strip_typographic_punctuation(None))


class TestLooksLikeAlreadyHtml(unittest.TestCase):

    def test_paragraph_tag_detected(self):
        self.assertTrue(looks_like_already_html("<p>Hello</p>"))

    def test_strong_tag_detected(self):
        self.assertTrue(looks_like_already_html("Already <strong>bold</strong>."))

    def test_plain_markdown_not_detected(self):
        self.assertFalse(looks_like_already_html("**bold** and *italic*"))

    def test_empty_returns_false(self):
        self.assertFalse(looks_like_already_html(""))
        self.assertFalse(looks_like_already_html(None))


class TestSanitizeDraftBody(unittest.TestCase):

    def test_bold_conversion(self):
        out = sanitize_draft_body("Thanks for the **quick reply** today.")
        self.assertIn("<strong>quick reply</strong>", out)
        self.assertNotIn("**", out)

    def test_italic_conversion(self):
        out = sanitize_draft_body("Please review *carefully* before signing.")
        self.assertIn("<em>carefully</em>", out)

    def test_em_dash_replaced_with_comma(self):
        out = sanitize_draft_body(
            "The redraw balance — your own funds — cannot be deducted."
        )
        self.assertNotIn("—", out)
        self.assertIn(
            "redraw balance, your own funds, cannot be deducted", out
        )

    def test_em_dash_no_spaces(self):
        out = sanitize_draft_body(
            "Sale price—less cost base—equals capital gain."
        )
        self.assertNotIn("—", out)

    def test_en_dash_in_range_becomes_hyphen(self):
        out = sanitize_draft_body("FY2024–2025 carried-forward losses.")
        self.assertIn("FY2024-2025", out)

    def test_bullets_become_list(self):
        raw = (
            "Cost base includes:\n"
            "- Original purchase price\n"
            "- Stamp duty\n"
            "- Selling costs\n"
            "\n"
            "Let me know if you have questions."
        )
        out = sanitize_draft_body(raw)
        self.assertIn("<ul>", out)
        self.assertIn("</ul>", out)
        self.assertEqual(out.count("<li>"), 3)
        self.assertIn("Original purchase price", out)

    def test_paragraphs_wrapped(self):
        out = sanitize_draft_body("First paragraph.\n\nSecond paragraph.")
        self.assertEqual(out.count("<p>"), 2)

    def test_idempotent_on_html_input(self):
        # If the body is already HTML, we should not wrap again.
        html_input = "<p>Already <strong>formatted</strong>.</p>"
        out = sanitize_draft_body(html_input)
        self.assertEqual(out.count("<p>"), 1)
        self.assertEqual(out.count("<strong>"), 1)

    def test_html_input_still_has_dashes_stripped(self):
        # Even when we skip the markdown pipeline, dashes must be cleaned.
        html_input = "<p>The redraw — your funds — cannot be deducted.</p>"
        out = sanitize_draft_body(html_input)
        self.assertNotIn("—", out)
        self.assertIn("redraw, your funds, cannot be deducted", out)

    def test_empty_input_returns_empty(self):
        self.assertEqual(sanitize_draft_body(""), "")
        self.assertEqual(sanitize_draft_body(None), "")

    def test_full_screenshot_case(self):
        """Reproduces the failing draft from the bug report."""
        raw = (
            "Thank you for your questions.\n"
            "\n"
            "**Calculating the capital gain on the property sale**\n"
            "\n"
            "The profit for tax purposes is not simply the net cash you received "
            "after settlement. The capital gain is calculated as:\n"
            "\n"
            "**Sale proceeds** (the contract price) **less** your **cost base**, "
            "which includes:\n"
            "- The original purchase price\n"
            "- Acquisition costs (stamp duty, legal fees, etc.)\n"
            "- Capital improvement costs incurred during ownership\n"
            "- Selling costs (agent commissions, legal fees, advertising, etc.)\n"
            "\n"
            "**Regarding the redraw account funds**\n"
            "\n"
            "The money in the redraw account is your own funds — it is "
            "essentially a reduction of your outstanding loan balance."
        )
        out = sanitize_draft_body(raw)
        self.assertNotIn("**", out)
        self.assertNotIn("—", out)
        self.assertIn(
            "<strong>Calculating the capital gain on the property sale</strong>",
            out,
        )
        self.assertIn("<ul>", out)
        self.assertEqual(out.count("<li>"), 4)


if __name__ == "__main__":
    unittest.main()
