"""Tests for the morning brief's SharePoint retry pass + issues section.

Covers C3 of Stream B:
  - Empty queue → no Graph calls in the retry pass; empty render output.
  - Missing-folder row that stays missing → row remains, "Folder not found"
    rendering.
  - Ambiguous-folder row that stays ambiguous → row remains, all candidate
    names rendered verbatim.
  - Folder situation resolved between retries → retry succeeds, row deleted,
    issues section disappears.
  - One row >7 days old → renders only under the "Stale" sub-heading, not
    the main one.
  - Mix of fresh and stale → both sub-sections render, no row duplicated.

Mock target is graph._ensure_client_folder_exists, not upload_to_sharepoint —
the retry is a probe, not an upload re-attempt.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(REPO_ROOT / "plugins"))

import config as cfg  # noqa: E402

_TEST_DB = Path(tempfile.mktemp(suffix="_morning_brief_sharepoint.db"))
cfg.DB_PATH = _TEST_DB
cfg.init_db()

from graph_client import (  # noqa: E402
    SHAREPOINT_CLIENT_BASE,
    SharePointFolderAmbiguous,
    SharePointFolderMissing,
)


def _make_graph():
    """Minimal Graph mock — site/drive resolution succeeds, the folder
    probe is configurable per test via side_effect."""
    g = MagicMock()
    g.get_sharepoint_site_id.return_value = "SITE_ID"
    g.get_sharepoint_drive_id.return_value = "DRIVE_ID"
    return g


def _backdate_row_created_at(message_id: str, days_ago: int) -> None:
    """Rewrite a pending row's created_at to N days in the past so the
    staleness threshold can be exercised without sleeping."""
    when = (datetime.utcnow() - timedelta(days=days_ago)).isoformat(
        sep=" ", timespec="seconds"
    )
    conn = cfg.get_db()
    conn.execute(
        "UPDATE sharepoint_upload_pending SET created_at = ? "
        "WHERE message_id = ?",
        (when, message_id),
    )
    conn.commit()
    conn.close()


class TestRetryPendingSharePointUploads(unittest.TestCase):

    def setUp(self):
        conn = cfg.get_db()
        conn.execute("DELETE FROM sharepoint_upload_pending")
        conn.commit()
        conn.close()
        from plugin_morning_briefing import MorningBriefingPlugin
        self.plugin = MorningBriefingPlugin()

    def test_empty_queue_is_a_noop_with_no_graph_calls(self):
        graph = _make_graph()
        scanned, resolved = self.plugin._retry_pending_sharepoint_uploads(graph)
        self.assertEqual((scanned, resolved), (0, 0))
        graph.get_sharepoint_site_id.assert_not_called()
        graph.get_sharepoint_drive_id.assert_not_called()
        graph._ensure_client_folder_exists.assert_not_called()

    def test_still_missing_folder_leaves_row_in_place(self):
        cfg.record_sharepoint_upload_pending("MSG-1", "Korkie, Gordon")
        graph = _make_graph()
        graph._ensure_client_folder_exists.side_effect = (
            SharePointFolderMissing("still missing")
        )
        scanned, resolved = self.plugin._retry_pending_sharepoint_uploads(graph)
        self.assertEqual((scanned, resolved), (1, 0))
        self.assertEqual(len(cfg.list_sharepoint_upload_pending()), 1)

    def test_still_ambiguous_folder_leaves_row_in_place(self):
        cfg.record_sharepoint_upload_pending(
            "MSG-2", "Beta Holdings",
            candidate_names=["Beta Holdings", "Beta  Holdings"],
        )
        graph = _make_graph()
        graph._ensure_client_folder_exists.side_effect = (
            SharePointFolderAmbiguous("still ambiguous", candidate_names=["A", "B"])
        )
        scanned, resolved = self.plugin._retry_pending_sharepoint_uploads(graph)
        self.assertEqual((scanned, resolved), (1, 0))
        self.assertEqual(len(cfg.list_sharepoint_upload_pending()), 1)

    def test_resolved_folder_deletes_the_row(self):
        cfg.record_sharepoint_upload_pending("MSG-3", "Korkie, Gordon")
        graph = _make_graph()
        graph._ensure_client_folder_exists.return_value = "Korkie, Gordon"
        scanned, resolved = self.plugin._retry_pending_sharepoint_uploads(graph)
        self.assertEqual((scanned, resolved), (1, 1))
        self.assertEqual(cfg.list_sharepoint_upload_pending(), [])

    def test_unexpected_exception_keeps_row_and_continues(self):
        cfg.record_sharepoint_upload_pending("MSG-4a", "Alpha")
        cfg.record_sharepoint_upload_pending("MSG-4b", "Beta")
        graph = _make_graph()
        # Alpha errors transiently; Beta resolves cleanly. Beta's row should
        # be deleted, Alpha's should remain.
        def _probe(*, folder_name, **_):
            if folder_name == "Alpha":
                raise RuntimeError("graph timeout")
            return folder_name
        graph._ensure_client_folder_exists.side_effect = _probe
        scanned, resolved = self.plugin._retry_pending_sharepoint_uploads(graph)
        self.assertEqual((scanned, resolved), (2, 1))
        remaining = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(remaining), 1)
        self.assertEqual(remaining[0]["client_name"], "Alpha")

    def test_skips_when_drive_unreachable(self):
        cfg.record_sharepoint_upload_pending("MSG-5", "X")
        graph = _make_graph()
        graph.get_sharepoint_site_id.return_value = None
        scanned, resolved = self.plugin._retry_pending_sharepoint_uploads(graph)
        self.assertEqual((scanned, resolved), (1, 0))
        graph._ensure_client_folder_exists.assert_not_called()
        # Row still there.
        self.assertEqual(len(cfg.list_sharepoint_upload_pending()), 1)


class TestRenderSharePointIssuesSection(unittest.TestCase):

    def setUp(self):
        conn = cfg.get_db()
        conn.execute("DELETE FROM sharepoint_upload_pending")
        conn.commit()
        conn.close()
        from plugin_morning_briefing import MorningBriefingPlugin
        self.plugin = MorningBriefingPlugin()

    # 1. Empty queue → empty render output.
    def test_empty_queue_renders_nothing(self):
        self.assertEqual(self.plugin._render_sharepoint_issues_section(), "")

    # 2. Missing-folder fresh row → main section, "Folder not found" wording.
    def test_missing_folder_fresh_row_renders_under_main_heading(self):
        cfg.record_sharepoint_upload_pending("MSG-A", "Korkie, Gordon")
        out = self.plugin._render_sharepoint_issues_section()
        self.assertIn("SharePoint folder issues", out)
        self.assertNotIn("Stale SharePoint issues", out)
        self.assertIn("• Korkie, Gordon", out)
        self.assertIn("Folder not found", out)
        self.assertIn(SHAREPOINT_CLIENT_BASE, out)

    # 3. Ambiguous-folder fresh row → "Multiple folders match" + verbatim.
    def test_ambiguous_folder_fresh_row_renders_all_candidates_verbatim(self):
        cfg.record_sharepoint_upload_pending(
            "MSG-B", "Beta Holdings",
            candidate_names=["Beta Holdings", "Beta  Holdings"],
        )
        out = self.plugin._render_sharepoint_issues_section()
        self.assertIn("• Beta Holdings", out)
        self.assertIn("Multiple folders match this client name:", out)
        # Both verbatim names — including the one with the double-space.
        self.assertIn(f"{SHAREPOINT_CLIENT_BASE}/Beta Holdings", out)
        self.assertIn(f"{SHAREPOINT_CLIENT_BASE}/Beta  Holdings", out)
        self.assertIn("Please merge these in SharePoint.", out)

    # 4. Stale row only → renders ONLY under the stale sub-heading.
    def test_stale_only_row_renders_under_stale_heading_only(self):
        cfg.record_sharepoint_upload_pending("MSG-C", "Old Client")
        _backdate_row_created_at("MSG-C", days_ago=8)
        out = self.plugin._render_sharepoint_issues_section()
        # Stale heading present; main heading should NOT be — there are no
        # fresh rows to put under it.
        self.assertIn("Stale SharePoint issues", out)
        self.assertNotIn("SharePoint folder issues\n", out.split("Stale")[0])
        self.assertIn("• Old Client", out)
        # The client only appears once in total.
        self.assertEqual(out.count("• Old Client"), 1)

    # 5. Mix of fresh and stale rows → both sub-sections, no duplication.
    def test_mix_of_fresh_and_stale_renders_both_sections(self):
        cfg.record_sharepoint_upload_pending("MSG-FRESH", "Fresh Client")
        cfg.record_sharepoint_upload_pending("MSG-STALE", "Stale Client")
        _backdate_row_created_at("MSG-STALE", days_ago=8)
        out = self.plugin._render_sharepoint_issues_section()
        self.assertIn("SharePoint folder issues", out)
        self.assertIn("Stale SharePoint issues", out)
        self.assertIn("• Fresh Client", out)
        self.assertIn("• Stale Client", out)
        # No duplication — each client appears in exactly one section.
        self.assertEqual(out.count("• Fresh Client"), 1)
        self.assertEqual(out.count("• Stale Client"), 1)
        # Stale heading appears AFTER the fresh content (sections are
        # rendered in fresh-then-stale order).
        self.assertLess(out.index("Fresh Client"), out.index("Stale SharePoint"))


if __name__ == "__main__":
    unittest.main()
