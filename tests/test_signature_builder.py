"""
Tests for signature_builder.build_signature_html
================================================
Covers the dynamic per-accountant signature path: mode routing, staff lookup
by M365 email, firm-constant rendering, and graceful fallback when assets or
rows are missing. Plus a small integration test confirming graph_client's
_append_signature actually delegates to signature_builder.

Each test runs against a fresh temp SQLite DB seeded by config.init_db() so
real user data is never touched and pre-seeded staff rows are present.
"""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg  # noqa: E402

# Point config at a temp DB BEFORE init_db runs anywhere.
_tmp_db = tempfile.mktemp(suffix=".db")
cfg.DB_PATH = Path(_tmp_db)
cfg.init_db()

import signature_builder as sb  # noqa: E402


def _legacy() -> str:
    """Sentinel returned by the legacy fallback so tests can detect routing."""
    return "<LEGACY-SIGNATURE-SENTINEL>"


class TestSignatureBuilder(unittest.TestCase):

    def setUp(self):
        # Reset modes + caches before each test so order doesn't matter.
        cfg.set_setting("signature_mode", "dynamic")
        cfg.set_setting("signature_company", "MC&S Pty Ltd")
        cfg.set_setting("signature_phone", "(03) 9794 0000")
        cfg.set_setting("signature_website_display", "mcands.com.au")
        cfg.set_setting("signature_website_url", "https://www.mcands.com.au")
        cfg.set_setting("signature_address_line1", "23 Timor Circuit, Keysborough, Vic 3173")
        cfg.set_setting("signature_address_line2", "PO BOX 4440, Dandenong South, VIC, 3164")
        cfg.set_setting("signature_instagram_url", "https://www.instagram.com/mcsaccounting")
        cfg.set_setting("signature_facebook_url", "https://www.facebook.com/mcandsaccounting")
        cfg.set_setting("signature_linkedin_url", "")
        cfg.set_setting("signature_google_review_url", "https://example.com/review")
        cfg.set_setting("signature_privacy_text", "Confidential test text.")
        sb.reset_image_cache()

    # ── Mode routing ────────────────────────────────────────────────────

    def test_mode_disabled_returns_empty(self):
        cfg.set_setting("signature_mode", "disabled")
        self.assertEqual(sb.build_signature_html("elio@mcands.com.au", _legacy), "")

    def test_mode_legacy_image_bypasses_dynamic_path(self):
        cfg.set_setting("signature_mode", "legacy_image")
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertEqual(out, "<LEGACY-SIGNATURE-SENTINEL>")

    def test_unknown_mode_defaults_to_dynamic(self):
        cfg.set_setting("signature_mode", "garbage")
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        # Garbage mode treated as dynamic, finds elio, renders signature.
        self.assertIn("Elio Scarton", out)

    # ── Staff lookup ────────────────────────────────────────────────────

    def test_match_by_email_returns_name_and_title(self):
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertIn("Elio Scarton", out)
        self.assertIn("CPA, Tax Agent", out)
        self.assertIn("MC&amp;S Pty Ltd", out)  # HTML-escaped & in company name

    def test_email_match_is_case_insensitive(self):
        out = sb.build_signature_html("ELIO@MCANDS.COM.AU", _legacy)
        self.assertIn("Elio Scarton", out)

    def test_title_null_omits_title_div(self):
        # Vince has title=NULL in seed data
        out = sb.build_signature_html("vince@mcands.com.au", _legacy)
        self.assertIn("Vince Mercuri", out)
        # Title div uses the distinctive #444444 colour — its absence proves
        # the title block was skipped, not just empty.
        self.assertNotIn("color:#444444", out)

    def test_unknown_email_falls_back_to_legacy(self):
        out = sb.build_signature_html("nobody@example.com", _legacy)
        self.assertEqual(out, "<LEGACY-SIGNATURE-SENTINEL>")

    def test_disabled_row_falls_back_to_legacy(self):
        # Disable Elio's row, then look him up.
        elio = cfg.get_staff_signature_by_email("elio@mcands.com.au")
        cfg.delete_staff_signature(elio["id"], soft=True)
        sb.reset_image_cache()  # also clears warned-once cache
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertEqual(out, "<LEGACY-SIGNATURE-SENTINEL>")
        # Restore for other tests
        cfg.save_staff_signature({**elio, "enabled": 1})

    def test_no_email_falls_back_to_legacy(self):
        out = sb.build_signature_html(None, _legacy)
        self.assertEqual(out, "<LEGACY-SIGNATURE-SENTINEL>")

    # ── Per-user include_signature toggle ──────────────────────────────

    def test_include_signature_default_on_renders_signature(self):
        # Existing seeded rows default to include_signature=1, so behaviour is
        # unchanged from before the column was added.
        elio = cfg.get_staff_signature_by_email("elio@mcands.com.au")
        self.assertEqual(elio["include_signature"], 1)
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertIn("Elio Scarton", out)

    def test_include_signature_off_returns_empty_string(self):
        # Toggle Elio's flag off and confirm the builder returns an empty
        # string — not the legacy fallback. Otherwise the toggle would do
        # nothing for users with a legacy image configured.
        elio = cfg.get_staff_signature_by_email("elio@mcands.com.au")
        cfg.save_staff_signature({**elio, "include_signature": 0})
        try:
            out = sb.build_signature_html("elio@mcands.com.au", _legacy)
            self.assertEqual(out, "")
        finally:
            cfg.save_staff_signature({**elio, "include_signature": 1})

    def test_include_signature_toggle_round_trips_via_save_helper(self):
        elio = cfg.get_staff_signature_by_email("elio@mcands.com.au")
        cfg.save_staff_signature({**elio, "include_signature": 0})
        try:
            self.assertEqual(
                cfg.get_staff_signature_by_email("elio@mcands.com.au")["include_signature"],
                0,
            )
        finally:
            cfg.save_staff_signature({**elio, "include_signature": 1})
        self.assertEqual(
            cfg.get_staff_signature_by_email("elio@mcands.com.au")["include_signature"],
            1,
        )

    # ── Firm constants + LinkedIn slot ─────────────────────────────────

    def test_linkedin_url_unset_omits_linkedin(self):
        # LinkedIn intentionally never renders in v1, so this asserts the URL
        # being unset doesn't break the rest of the signature.
        cfg.set_setting("signature_linkedin_url", "")
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertNotIn("linkedin", out.lower())

    def test_google_review_url_unset_omits_review_line(self):
        cfg.set_setting("signature_google_review_url", "")
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertNotIn("Leave us a Google review", out)

    def test_privacy_text_renders_with_italic_small_font(self):
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertIn("font-style:italic", out)
        self.assertIn("font-size:8pt", out)
        self.assertIn("Confidential test text.", out)

    def test_phone_change_reflected_immediately(self):
        cfg.set_setting("signature_phone", "(03) 0000 9999")
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        self.assertIn("(03) 0000 9999", out)

    # ── Image data URIs ────────────────────────────────────────────────

    def test_image_data_uris_load_from_bundled_assets(self):
        # logo.png + the two social PNGs ship in assets/signature/. They were
        # copied as the first commit of feat/dynamic-signatures; if they go
        # missing this test catches it.
        out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        # Logo: 80x80 in the template
        self.assertIn('data:image/png;base64,', out)
        self.assertIn('width="80" height="80"', out)
        # Social icons: 20x20 in the template
        self.assertIn('width="20" height="20"', out)

    def test_missing_logo_skips_logo_cell(self):
        # Patch the asset path to something that doesn't exist; signature
        # should still render with name/title/etc but no logo img.
        with patch.object(sb, "_LOGO_FILE", Path("/no/such/logo.png")):
            sb.reset_image_cache()
            out = sb.build_signature_html("elio@mcands.com.au", _legacy)
        # Logo's distinctive 80x80 marker is gone
        self.assertNotIn('width="80" height="80"', out)
        # But the rest of the signature is intact
        self.assertIn("Elio Scarton", out)
        self.assertIn("CPA, Tax Agent", out)

    # ── Pre-seed idempotency ───────────────────────────────────────────

    def test_pre_seed_inserts_nine_rows_and_is_idempotent(self):
        rows = cfg.get_staff_signatures(include_disabled=True)
        # We may have soft-deleted in earlier tests; count enabled+disabled,
        # then re-init to confirm no duplicates.
        original = len(rows)
        cfg.init_db()
        rows2 = cfg.get_staff_signatures(include_disabled=True)
        self.assertEqual(len(rows2), original)
        emails = [r["email"] for r in rows2]
        self.assertEqual(len(emails), len(set(emails)), "duplicate emails after re-seed")
        # The seed list defines exactly 9 distinct emails
        seed_emails = {
            "elio@mcands.com.au", "vince@mcands.com.au", "angelo@mcands.com.au",
            "ross@mcands.com.au", "reception@mcands.com.au", "brooke@mcands.com.au",
            "harry@mcands.com.au", "lyn@mcands.com.au", "louise@mcands.com.au",
        }
        self.assertTrue(seed_emails.issubset(set(emails)))


class TestGraphClientIntegration(unittest.TestCase):
    """Confirm graph_client._append_signature delegates to signature_builder
    when signature_mode == 'dynamic'."""

    def setUp(self):
        cfg.set_setting("signature_mode", "dynamic")
        cfg.set_setting("ms_account_email", "elio@mcands.com.au")

    def test_append_signature_calls_build_signature_html(self):
        from graph_client import GraphClient
        gc = GraphClient.__new__(GraphClient)  # bypass MSAL init for unit test

        with patch("signature_builder.build_signature_html",
                   return_value="<DYNAMIC-OK>") as mock_build:
            result = gc._append_signature("<p>Hello</p>")

        mock_build.assert_called_once()
        # Body should now end with our mocked signature
        self.assertTrue(result.endswith("<DYNAMIC-OK>"))
        # Email arg passed in is the cached session user (lowercased)
        args, kwargs = mock_build.call_args
        passed_email = args[0] if args else kwargs.get("user_email")
        self.assertEqual(passed_email, "elio@mcands.com.au")

    def test_append_signature_legacy_mode_skips_dynamic_builder(self):
        cfg.set_setting("signature_mode", "legacy_image")
        from graph_client import GraphClient
        gc = GraphClient.__new__(GraphClient)
        with patch("signature_builder.build_signature_html",
                   return_value="<DYNAMIC-OK>") as mock_build, \
             patch.object(GraphClient, "get_signature_html",
                          return_value="<LEGACY-IMG>"):
            result = gc._append_signature("<p>Hello</p>")
        mock_build.assert_not_called()
        self.assertIn("<LEGACY-IMG>", result)

    def test_append_signature_disabled_mode_appends_nothing(self):
        cfg.set_setting("signature_mode", "disabled")
        from graph_client import GraphClient
        gc = GraphClient.__new__(GraphClient)
        result = gc._append_signature("<p>Hello</p>")
        # Body untouched (linkify is a no-op on this content)
        self.assertEqual(result, "<p>Hello</p>")


if __name__ == "__main__":
    unittest.main()
