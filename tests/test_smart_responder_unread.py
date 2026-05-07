"""Tests for the post-draft mark-as-unread flip in SmartEmailResponderPlugin.

Verifies the contract introduced by the Stream A fix: after a draft reply is
created, the original email is restored to unread so the accountant sees a
bold envelope, while idempotency is owned by the smart_responder_processed
SQLite table (not the Outlook isRead flag).
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

# Pin the DB to a per-process temp file so init_db creates the schema there
# and the real _is_already_processed / _mark_as_processed read/write it.
_TEST_DB = Path(tempfile.mktemp(suffix="_smart_responder_unread.db"))
cfg.DB_PATH = _TEST_DB
cfg.init_db()


class TestSmartResponderUnreadFlip(unittest.TestCase):
    """Pin the contract: draft → mark_as_unread, with the processed-table
    as the sole loop guard."""

    def setUp(self):
        from config import get_db
        # Wipe the processed-table between tests so prior runs don't bleed in.
        conn = get_db()
        conn.execute("DELETE FROM smart_responder_processed")
        conn.commit()
        conn.close()

        from plugin_smart_responder import SmartEmailResponderPlugin
        self.plugin = SmartEmailResponderPlugin()

        from plugin_base import PluginContext
        self.ctx = MagicMock(spec=PluginContext)
        self.ctx.draft_mode = True
        self.ctx.graph = MagicMock()
        # Memory and gateway off — the plugin's auto-file / memory writes are
        # all wrapped in try/except, so leaving them as None keeps the test
        # focused on the read/unread contract.
        self.ctx.memory = None
        self.ctx.gateway = None
        self.ctx.event_bus = None
        self.ctx.approval_queue = None
        self.ctx.log = MagicMock()
        self.ctx.claude_reason = MagicMock()
        self.ctx.claude_fast = MagicMock()
        self.ctx.claude = self.ctx.claude_reason  # legacy alias

        self.message_id = "AAMkAGI2N-test-msg-1"
        self.email = {
            "id": self.message_id,
            "subject": "Quick question",
            "from": {"emailAddress": {"address": "client@example.com",
                                      "name": "Test Client"}},
            "body": {"content": "Body text"},
            "bodyPreview": "Body text",
            "hasAttachments": False,
            "toRecipients": [],
        }
        self.ctx.graph.fetch_unread_emails.return_value = [self.email]
        self.ctx.graph.create_draft.return_value = "DRAFT_ID_123"
        self.ctx.graph.upload_to_sharepoint.return_value = (
            "https://sharepoint.example/file"
        )

    def _patch_claude(self):
        """Bypass the Anthropic call by short-circuiting _ask_claude."""
        return patch.object(
            self.plugin, "_ask_claude",
            return_value="<p>Hello, here's a reply.</p>",
        )

    # ── 1. Happy path ────────────────────────────────────────────────────────
    def test_happy_path_marks_unread_and_persists_processed_row(self):
        with self._patch_claude():
            result = self.plugin.run(self.ctx)

        self.assertTrue(result.success)
        self.ctx.graph.create_draft.assert_called_once()
        self.ctx.graph.mark_as_unread.assert_called_once_with(self.message_id)
        # The drafted-path must NOT call mark_as_read — that was the
        # behaviour we are explicitly inverting.
        self.assertFalse(
            self.ctx.graph.mark_as_read.called,
            "mark_as_read should not be called after a successful draft",
        )

        from config import get_db
        conn = get_db()
        row = conn.execute(
            "SELECT message_id, draft_id, action "
            "FROM smart_responder_processed WHERE message_id = ?",
            (self.message_id,),
        ).fetchone()
        conn.close()
        self.assertIsNotNone(row, "processed-table row should exist")
        self.assertEqual(row[0], self.message_id)
        self.assertEqual(row[1], "DRAFT_ID_123")
        self.assertEqual(row[2], "drafted")

    # ── 2. mark_as_unread failure must be non-fatal ──────────────────────────
    def test_mark_as_unread_failure_does_not_unwind_the_draft(self):
        self.ctx.graph.mark_as_unread.side_effect = RuntimeError("graph timeout")

        with self._patch_claude(), \
             self.assertLogs("plugin_smart_responder", level="WARNING") as captured:
            result = self.plugin.run(self.ctx)

        self.assertTrue(result.success)
        self.ctx.graph.create_draft.assert_called_once()
        self.ctx.graph.mark_as_unread.assert_called_once_with(self.message_id)

        from config import get_db
        conn = get_db()
        row = conn.execute(
            "SELECT 1 FROM smart_responder_processed WHERE message_id = ?",
            (self.message_id,),
        ).fetchone()
        conn.close()
        self.assertIsNotNone(
            row, "processed-table row must persist even if mark_as_unread fails"
        )
        self.assertTrue(
            any("Failed to mark message as unread after draft" in line
                for line in captured.output),
            f"expected warning log, got: {captured.output}",
        )

    # ── 3. Already-processed email is skipped ────────────────────────────────
    def test_already_processed_email_is_skipped(self):
        from config import get_db
        conn = get_db()
        conn.execute(
            "INSERT INTO smart_responder_processed "
            "(message_id, draft_id, action, processed_at) "
            "VALUES (?, ?, ?, CURRENT_TIMESTAMP)",
            (self.message_id, "EXISTING_DRAFT_ID", "drafted"),
        )
        conn.commit()
        conn.close()

        with self._patch_claude() as mock_claude:
            result = self.plugin.run(self.ctx)

        # Idempotency guard prevents Claude, draft, and any state-changing
        # Graph call from firing.
        mock_claude.assert_not_called()
        self.ctx.graph.create_draft.assert_not_called()
        self.ctx.graph.mark_as_unread.assert_not_called()
        self.ctx.graph.mark_as_read.assert_not_called()
        self.assertTrue(result.success)

    # ── 4. Order of operations: create_draft → mark_as_unread ───────────────
    def test_create_draft_is_called_before_mark_as_unread(self):
        order_log: list[str] = []

        def _record_create_draft(*args, **kwargs):
            order_log.append("create_draft")
            return "DRAFT_ID_123"

        def _record_mark_as_unread(*args, **kwargs):
            order_log.append("mark_as_unread")

        self.ctx.graph.create_draft.side_effect = _record_create_draft
        self.ctx.graph.mark_as_unread.side_effect = _record_mark_as_unread

        with self._patch_claude():
            self.plugin.run(self.ctx)

        self.assertIn("create_draft", order_log)
        self.assertIn("mark_as_unread", order_log)
        self.assertLess(
            order_log.index("create_draft"),
            order_log.index("mark_as_unread"),
            f"create_draft must precede mark_as_unread; got: {order_log}",
        )


if __name__ == "__main__":
    unittest.main()
