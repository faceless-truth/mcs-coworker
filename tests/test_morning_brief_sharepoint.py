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
from unittest.mock import MagicMock, patch

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


class TestQuietDaySuppressionWithSharePointPending(unittest.TestCase):
    """Pin the contract that unresolved SharePoint folder issues count
    toward the quiet-day actionable tally. A quiet day with accumulated
    SP issues must NOT suppress the brief — otherwise the punch list
    silently grows."""

    def setUp(self):
        # Wipe both pending queue and any previous user_email so two tests
        # in the same process don't leak settings into each other.
        conn = cfg.get_db()
        conn.execute("DELETE FROM sharepoint_upload_pending")
        conn.execute(
            "DELETE FROM settings WHERE key IN ('user_email')"
        )
        conn.commit()
        conn.close()

        from plugin_morning_briefing import MorningBriefingPlugin
        self.plugin = MorningBriefingPlugin()

        from plugin_base import PluginContext
        self.ctx = MagicMock(spec=PluginContext)
        self.ctx.draft_mode = True
        self.ctx.graph = MagicMock()
        self.ctx.memory = None
        self.ctx.gateway = None
        self.ctx.event_bus = None
        self.ctx.approval_queue = None
        self.ctx.log = MagicMock()
        self.ctx.claude_reason = MagicMock()
        self.ctx.claude_fast = MagicMock()
        self.ctx.claude = self.ctx.claude_reason

    def _run_at_8am_monday(self, actionable_count: int):
        """Drive run() through the time/business-day gates with a known
        Monday 08:00 'now', a stubbed compile that returns the desired
        actionable count, and a no-op retry pass so seeded SP rows survive
        into the bump-and-suppression-check window.

        Patches `datetime` surgically: `.now()` returns the fixed Monday
        morning so the time gate passes, but `.utcnow()` and
        `.fromisoformat()` delegate to the real class so the issues-section
        renderer's staleness arithmetic works.
        """
        # 2025-05-12 was a Monday; 08:00 matches the default briefing_hour.
        mocked_now = datetime(2025, 5, 12, 8, 0)
        with patch("plugin_morning_briefing.datetime") as mock_dt, \
             patch.object(self.plugin, "_compile_accountant_briefing",
                          return_value=("Brief body", actionable_count)), \
             patch.object(self.plugin, "_retry_pending_sharepoint_uploads",
                          return_value=(0, 0)):
            mock_dt.now.return_value = mocked_now
            mock_dt.utcnow.side_effect = datetime.utcnow
            mock_dt.fromisoformat.side_effect = datetime.fromisoformat
            return self.plugin.run(self.ctx)

    # Regression-pin existing behaviour: empty queue → suppression still
    # fires when the day is otherwise quiet.
    def test_quiet_day_with_empty_queue_still_suppresses(self):
        result = self._run_at_8am_monday(actionable_count=0)
        self.assertIn("suppressed", result.summary.lower())
        self.ctx.graph.create_draft.assert_not_called()
        self.ctx.graph.send_email.assert_not_called()

    # New behaviour: empty otherwise-actionable but unresolved SP rows
    # bump the tally past the quiet threshold, brief proceeds, issues
    # section ends up in the email body.
    def test_quiet_day_with_sp_rows_does_not_suppress_and_renders_section(self):
        cfg.record_sharepoint_upload_pending("MSG-Q-1", "Korkie, Gordon")
        cfg.record_sharepoint_upload_pending(
            "MSG-Q-2", "Beta Holdings",
            candidate_names=["Beta Holdings", "Beta  Holdings"],
        )
        cfg.record_sharepoint_upload_pending("MSG-Q-3", "Gamma Inc")
        # Configure a recipient so create_draft actually fires; otherwise
        # the brief code path skips delivery on empty recipients and we
        # can't inspect the body that would have been sent.
        cfg.set_setting("user_email", "elio@mcands.com.au")

        result = self._run_at_8am_monday(actionable_count=0)

        self.assertNotIn("suppressed", result.summary.lower())
        self.ctx.graph.create_draft.assert_called_once()
        # Body argument carries the issues section verbatim.
        positional = self.ctx.graph.create_draft.call_args.args
        body_html = positional[2] if len(positional) >= 3 else ""
        self.assertIn("SharePoint folder issues", body_html)
        self.assertIn("Korkie, Gordon", body_html)
        # Verbatim ambiguous candidates including the double-space.
        self.assertIn(f"{SHAREPOINT_CLIENT_BASE}/Beta  Holdings", body_html)


if __name__ == "__main__":
    unittest.main()
