"""Tests for the sharepoint_upload_pending SQLite queue and its CRUD helpers
in config.py.

The queue is recorded into by smart_responder when an upload hits a folder
governance error, and drained by the morning brief's retry pass. These tests
pin the contract: INSERT OR IGNORE on (message_id, action), JSON round-trip
for candidate_names, and resolution-via-delete.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg  # noqa: E402

# Per-process temp DB so init_db creates the new sharepoint_upload_pending
# table there and the queue helpers operate on it.
_TEST_DB = Path(tempfile.mktemp(suffix="_sharepoint_pending.db"))
cfg.DB_PATH = _TEST_DB
cfg.init_db()


class TestSharePointUploadPendingQueue(unittest.TestCase):

    def setUp(self):
        # Wipe between tests so prior runs don't bleed in.
        conn = cfg.get_db()
        conn.execute("DELETE FROM sharepoint_upload_pending")
        conn.commit()
        conn.close()

    # ── Insert + read-back ───────────────────────────────────────────────────
    def test_record_inserts_a_row_with_no_candidates(self):
        inserted = cfg.record_sharepoint_upload_pending(
            message_id="MSG-1",
            client_name="Korkie, Gordon",
        )
        self.assertTrue(inserted)
        rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["message_id"], "MSG-1")
        self.assertEqual(rows[0]["client_name"], "Korkie, Gordon")
        self.assertEqual(rows[0]["action"], "upload_pending")
        self.assertIsNone(rows[0]["candidate_names"])
        self.assertIsNone(rows[0]["resolved_at"])

    def test_record_with_candidate_names_round_trips_as_list(self):
        cfg.record_sharepoint_upload_pending(
            message_id="MSG-2",
            client_name="Beta Holdings",
            candidate_names=["Beta Holdings", "Beta  Holdings"],
        )
        rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(rows), 1)
        self.assertEqual(
            rows[0]["candidate_names"],
            ["Beta Holdings", "Beta  Holdings"],
        )
        # Underlying storage is JSON — sanity check the column directly.
        conn = cfg.get_db()
        raw = conn.execute(
            "SELECT candidate_names_json FROM sharepoint_upload_pending "
            "WHERE message_id = ?", ("MSG-2",),
        ).fetchone()[0]
        conn.close()
        self.assertEqual(json.loads(raw), ["Beta Holdings", "Beta  Holdings"])

    # ── INSERT OR IGNORE ─────────────────────────────────────────────────────
    def test_duplicate_message_id_is_ignored_not_duplicated(self):
        first = cfg.record_sharepoint_upload_pending(
            message_id="MSG-3", client_name="Same Client",
        )
        second = cfg.record_sharepoint_upload_pending(
            message_id="MSG-3", client_name="Same Client",
            candidate_names=["A", "B"],  # different payload — still ignored
        )
        self.assertTrue(first)
        self.assertFalse(second)
        rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(rows), 1)
        # The first record's payload (None candidates) wins.
        self.assertIsNone(rows[0]["candidate_names"])

    # ── Filtering ────────────────────────────────────────────────────────────
    def test_list_only_unresolved_excludes_resolved_rows(self):
        cfg.record_sharepoint_upload_pending(
            message_id="MSG-A", client_name="A",
        )
        cfg.record_sharepoint_upload_pending(
            message_id="MSG-B", client_name="B",
        )
        # Mark MSG-A as resolved by hand — production code uses delete, but
        # the resolved_at column is in the schema for forward-compat.
        conn = cfg.get_db()
        conn.execute(
            "UPDATE sharepoint_upload_pending SET resolved_at = "
            "CURRENT_TIMESTAMP WHERE message_id = 'MSG-A'"
        )
        conn.commit()
        conn.close()

        unresolved = cfg.list_sharepoint_upload_pending(only_unresolved=True)
        self.assertEqual(len(unresolved), 1)
        self.assertEqual(unresolved[0]["message_id"], "MSG-B")

        all_rows = cfg.list_sharepoint_upload_pending(only_unresolved=False)
        self.assertEqual(len(all_rows), 2)

    # ── Delete ───────────────────────────────────────────────────────────────
    def test_delete_removes_the_row(self):
        cfg.record_sharepoint_upload_pending(
            message_id="MSG-D", client_name="DeleteMe",
        )
        rows = cfg.list_sharepoint_upload_pending()
        self.assertEqual(len(rows), 1)
        row_id = rows[0]["id"]
        cfg.delete_sharepoint_upload_pending(row_id)
        self.assertEqual(cfg.list_sharepoint_upload_pending(), [])

    # ── Edge cases ───────────────────────────────────────────────────────────
    def test_empty_message_id_or_client_name_inserts_nothing(self):
        self.assertFalse(cfg.record_sharepoint_upload_pending("", "client"))
        self.assertFalse(cfg.record_sharepoint_upload_pending("msg", ""))
        self.assertEqual(cfg.list_sharepoint_upload_pending(), [])


if __name__ == "__main__":
    unittest.main()
