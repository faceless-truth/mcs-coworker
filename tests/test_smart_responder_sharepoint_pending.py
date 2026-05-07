"""Tests for smart_responder's catch-and-record behaviour around the
SharePoint folder governance exceptions.

Pins the contract introduced by Stream B's prevention fix:

  - SharePointFolderMissing → row in sharepoint_upload_pending,
    smart_responder_processed STILL has the draft row (idempotency
    intact), no exception propagates.
  - SharePointFolderAmbiguous → row in sharepoint_upload_pending with
    candidate_names populated.
  - Generic Exception during upload → no upload_pending row, the existing
    best-effort warning behaviour is preserved.
  - Re-running smart_responder against an unresolved message doesn't
    duplicate the upload_pending row (INSERT OR IGNORE).
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(REPO_ROOT / "plugins"))

import config as cfg  # noqa: E402

_TEST_DB = Path(tempfile.mktemp(suffix="_smart_responder_sharepoint_pending.db"))
cfg.DB_PATH = _TEST_DB
cfg.init_db()


from graph_client import (  # noqa: E402
    SharePointFolderAmbiguous,
    SharePointFolderMissing,
)


class TestSmartResponderSharePointPending(unittest.TestCase):
    """Pin the catch-and-record behaviour added in Stream B."""

    def setUp(self):
        # Wipe both tables between tests so prior runs don't bleed in.
        conn = cfg.get_db()
        conn.execute("DELETE FROM smart_responder_processed")
        conn.execute("DELETE FROM sharepoint_upload_pending")
        conn.commit()
        conn.close()

        from plugin_smart_responder import SmartEmailResponderPlugin
        self.plugin = SmartEmailResponderPlugin()

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

        self.message_id = "AAMkAGI2N-sp-pending-1"
        self.email = {
            "id": self.message_id,
            "subject": "Quote for Korkie",
            "from": {"emailAddress": {
                "address": "gordon.korkie@example.com",
                "name": "Gordon Korkie",
            }},
            "body": {"content": "Body"},
            "bodyPreview": "Body",
            "hasAttachments": False,
            "toRecipients": [],
        }
        self.ctx.graph.fetch_unread_emails.return_value = [self.email]
        self.ctx.graph.create_draft.return_value = "DRAFT_ID_SP_1"

    def _patch_claude(self):
        return patch.object(
            self.plugin, "_ask_claude",
            return_value="<p>Quote attached.</p>",
        )

    def _processed_row_exists(self) -> bool:
        conn = cfg.get_db()
        row = conn.execute(
            "SELECT 1 FROM smart_responder_processed WHERE message_id = ?",
            (self.message_id,),
        ).fetchone()
        conn.close()
        return row is not None

    # ── 1. Missing folder → upload_pending row, draft preserved ──────────────
    def test_missing_folder_records_pending_and_preserves_draft(self):
        self.ctx.graph.upload_to_sharepoint.side_effect = (
            SharePointFolderMissing("Korkie, Gordon not found")
        )
        with self._patch_claude(), \
             self.assertLogs("plugin_smart_responder", level="INFO") as cap:
            result = self.plugin.run(self.ctx)

        self.assertTrue(result.success)
        # Draft idempotency intact.
        self.assertTrue(self._processed_row_exists(),
                        "smart_responder_processed row must be present")
        # Pending row recorded with no candidate_names.
        rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["message_id"], self.message_id)
        self.assertEqual(rows[0]["client_name"], "Korkie, Gordon")
        self.assertIsNone(rows[0]["candidate_names"])
        # Logged at INFO, not WARNING — this is expected operational state.
        self.assertTrue(
            any("SharePoint folder missing" in line for line in cap.output),
            f"expected info log, got: {cap.output}",
        )

    # ── 2. Ambiguous folder → candidate_names populated ──────────────────────
    def test_ambiguous_folder_records_candidates(self):
        self.ctx.graph.upload_to_sharepoint.side_effect = (
            SharePointFolderAmbiguous(
                "Multiple matches",
                candidate_names=["Beta Holdings", "Beta  Holdings"],
            )
        )
        with self._patch_claude(), \
             self.assertLogs("plugin_smart_responder", level="INFO"):
            result = self.plugin.run(self.ctx)

        self.assertTrue(result.success)
        self.assertTrue(self._processed_row_exists())
        rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(rows), 1)
        self.assertEqual(
            rows[0]["candidate_names"],
            ["Beta Holdings", "Beta  Holdings"],
        )

    # ── 3. Generic Exception → no pending row, behaviour preserved ───────────
    def test_generic_exception_during_upload_does_not_record_pending(self):
        self.ctx.graph.upload_to_sharepoint.side_effect = (
            RuntimeError("transient network error")
        )
        with self._patch_claude(), \
             self.assertLogs("plugin_smart_responder", level="WARNING") as cap:
            result = self.plugin.run(self.ctx)

        self.assertTrue(result.success)
        self.assertTrue(self._processed_row_exists())
        # No pending row — the regression boundary.
        self.assertEqual(cfg.list_sharepoint_upload_pending(), [])
        # Existing warning log is preserved.
        self.assertTrue(
            any("SharePoint auto-file failed" in line for line in cap.output),
            f"expected warning log, got: {cap.output}",
        )

    # ── 4. INSERT OR IGNORE: re-run on same message doesn't duplicate ────────
    def test_rerun_on_same_message_does_not_duplicate_pending_row(self):
        # First run: missing folder → records the pending row.
        self.ctx.graph.upload_to_sharepoint.side_effect = (
            SharePointFolderMissing("not found")
        )
        with self._patch_claude(), self.assertLogs("plugin_smart_responder",
                                                   level="INFO"):
            self.plugin.run(self.ctx)
        first_rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(first_rows), 1)

        # Simulate a second run by clearing the processed-table guard so the
        # plugin re-enters the upload path for the same message_id.
        # (In production, this happens if the processed-table is dropped or
        # the message_id collides — the safety here is purely INSERT OR
        # IGNORE on the pending queue.)
        conn = cfg.get_db()
        conn.execute(
            "DELETE FROM smart_responder_processed WHERE message_id = ?",
            (self.message_id,),
        )
        conn.commit()
        conn.close()

        with self._patch_claude(), self.assertLogs("plugin_smart_responder",
                                                   level="INFO"):
            self.plugin.run(self.ctx)
        second_rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(second_rows), 1,
                         "second run must not duplicate the pending row")
        self.assertEqual(second_rows[0]["id"], first_rows[0]["id"])


if __name__ == "__main__":
    unittest.main()
