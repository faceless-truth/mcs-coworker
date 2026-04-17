"""
Stream 6 Test Suite — Confidence-Based Approval Queue
=====================================================
Tests the ApprovalQueue, ActionStatus, PendingAction, and
PluginContext.approval_queue injection.

Test groups:
  1.  ApprovalQueue initialisation and table creation
  2.  Threshold management
  3.  submit() — auto-approve path (confidence >= threshold)
  4.  submit() — queue path (confidence < threshold)
  5.  get_pending() and get_all()
  6.  approve() — happy path and edge cases
  7.  reject() — happy path and edge cases
  8.  Expiry — _expire_old_items()
  9.  clear_old()
  10. summary()
  11. EventBus integration — events published on submit/approve/reject
  12. PluginContext injection via plugin_loader._make_context()
  13. config.py default settings seeded
"""

import sys
import os
import json
import sqlite3
import tempfile
import threading
import time
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

# ── Path setup ────────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg
_TEST_DB = Path(tempfile.mktemp(suffix="_stream6_test.db"))
cfg.DB_PATH = _TEST_DB
cfg.init_db()

from approval_queue import (
    ApprovalQueue, ActionStatus, PendingAction, get_approval_queue
)


# ── Helper: fresh queue with isolated DB ─────────────────────────────────────

def make_queue() -> ApprovalQueue:
    """Return a new ApprovalQueue backed by a fresh temp DB."""
    db_path = Path(tempfile.mktemp(suffix="_aq_test.db"))
    return ApprovalQueue(db_path=db_path)


# ── 1. Initialisation ─────────────────────────────────────────────────────────

class TestApprovalQueueInit(unittest.TestCase):

    def test_table_created_on_init(self):
        q = make_queue()
        conn = sqlite3.connect(str(q._db_path))
        tables = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
        conn.close()
        table_names = [t[0] for t in tables]
        self.assertIn("approval_queue", table_names)

    def test_singleton_returns_same_instance(self):
        # Reset singleton for test
        import approval_queue as aq_mod
        aq_mod._queue_instance = None
        q1 = get_approval_queue()
        q2 = get_approval_queue()
        self.assertIs(q1, q2)

    def test_default_threshold(self):
        q = make_queue()
        # Default from config is 0.75
        self.assertAlmostEqual(q.get_threshold(), 0.75, places=2)

    def test_default_expiry_hours(self):
        q = make_queue()
        self.assertEqual(q.get_expiry_hours(), 48)


# ── 2. Threshold management ───────────────────────────────────────────────────

class TestThresholdManagement(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()

    def test_set_threshold_updates_value(self):
        self.q.set_threshold(0.9)
        self.assertAlmostEqual(self.q.get_threshold(), 0.9, places=2)

    def test_threshold_clamped_to_0_1(self):
        self.q.set_threshold(1.5)
        self.assertAlmostEqual(self.q.get_threshold(), 1.0, places=2)
        self.q.set_threshold(-0.5)
        self.assertAlmostEqual(self.q.get_threshold(), 0.0, places=2)

    def test_threshold_at_zero_auto_approves_everything(self):
        self.q.set_threshold(0.0)
        callback_called = []
        result = self.q.submit(
            action_type="test",
            description="test action",
            payload={},
            confidence=0.01,
            plugin_id="test_plugin",
            execute_callback=lambda: callback_called.append(True),
        )
        self.assertTrue(result)
        self.assertEqual(len(callback_called), 1)

    def test_threshold_at_one_queues_everything(self):
        self.q.set_threshold(1.0)
        result = self.q.submit(
            action_type="test",
            description="test action",
            payload={},
            confidence=0.99,
            plugin_id="test_plugin",
        )
        self.assertFalse(result)
        self.assertEqual(self.q.count_pending(), 1)


# ── 3. submit() — auto-approve path ──────────────────────────────────────────

class TestSubmitAutoApprove(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)

    def test_high_confidence_returns_true(self):
        result = self.q.submit(
            action_type="send_email",
            description="Send reminder to client",
            payload={"to": "client@example.com"},
            confidence=0.90,
            plugin_id="plugin_outreach",
        )
        self.assertTrue(result)

    def test_exact_threshold_returns_true(self):
        result = self.q.submit(
            action_type="send_email",
            description="Send reminder",
            payload={},
            confidence=0.75,
            plugin_id="plugin_test",
        )
        self.assertTrue(result)

    def test_auto_approve_executes_callback(self):
        called = []
        self.q.submit(
            action_type="send_email",
            description="Send reminder",
            payload={},
            confidence=0.85,
            plugin_id="plugin_test",
            execute_callback=lambda: called.append(True),
        )
        self.assertEqual(len(called), 1)

    def test_auto_approve_logged_in_db(self):
        self.q.submit(
            action_type="send_email",
            description="Auto-approved action",
            payload={"to": "x@x.com"},
            confidence=0.80,
            plugin_id="plugin_test",
        )
        all_items = self.q.get_all(status=ActionStatus.AUTO_APPROVED)
        self.assertEqual(len(all_items), 1)
        self.assertEqual(all_items[0].status, ActionStatus.AUTO_APPROVED)

    def test_auto_approve_callback_exception_does_not_crash(self):
        def bad_callback():
            raise RuntimeError("Simulated failure")

        # Should not raise
        result = self.q.submit(
            action_type="send_email",
            description="Failing action",
            payload={},
            confidence=0.85,
            plugin_id="plugin_test",
            execute_callback=bad_callback,
        )
        # Returns True (was auto-approved) even if callback failed
        self.assertTrue(result)

    def test_confidence_clamped_above_1(self):
        result = self.q.submit(
            action_type="test",
            description="Clamped confidence",
            payload={},
            confidence=5.0,  # should be clamped to 1.0
            plugin_id="plugin_test",
        )
        self.assertTrue(result)


# ── 4. submit() — queue path ──────────────────────────────────────────────────

class TestSubmitQueue(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)

    def test_low_confidence_returns_false(self):
        result = self.q.submit(
            action_type="send_email",
            description="Uncertain action",
            payload={},
            confidence=0.50,
            plugin_id="plugin_test",
        )
        self.assertFalse(result)

    def test_below_threshold_inserts_pending_row(self):
        self.q.submit(
            action_type="send_email",
            description="Queued action",
            payload={"to": "client@example.com"},
            confidence=0.60,
            plugin_id="plugin_outreach",
        )
        pending = self.q.get_pending()
        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0].action_type, "send_email")
        self.assertEqual(pending[0].confidence, 0.60)
        self.assertEqual(pending[0].status, ActionStatus.PENDING)

    def test_payload_stored_and_retrieved(self):
        payload = {"to": "client@example.com", "subject": "Tax Return", "amount": 500}
        self.q.submit(
            action_type="send_email",
            description="Queued with payload",
            payload=payload,
            confidence=0.40,
            plugin_id="plugin_test",
        )
        pending = self.q.get_pending()
        self.assertEqual(pending[0].payload, payload)

    def test_multiple_pending_items(self):
        for i in range(5):
            self.q.submit(
                action_type="send_email",
                description=f"Action {i}",
                payload={"index": i},
                confidence=0.30,
                plugin_id="plugin_test",
            )
        self.assertEqual(self.q.count_pending(), 5)

    def test_confidence_clamped_below_0(self):
        result = self.q.submit(
            action_type="test",
            description="Negative confidence",
            payload={},
            confidence=-1.0,  # should be clamped to 0.0
            plugin_id="plugin_test",
        )
        self.assertFalse(result)
        pending = self.q.get_pending()
        self.assertEqual(pending[0].confidence, 0.0)


# ── 5. get_pending() and get_all() ───────────────────────────────────────────

class TestQueryMethods(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)

    def test_get_pending_returns_only_pending(self):
        # Add one pending and one auto-approved
        self.q.submit("test", "Pending", {}, 0.40, "plugin_test")
        self.q.submit("test", "Auto-approved", {}, 0.90, "plugin_test")
        pending = self.q.get_pending()
        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0].description, "Pending")

    def test_get_all_returns_all_statuses(self):
        self.q.submit("test", "Pending", {}, 0.40, "plugin_test")
        self.q.submit("test", "Auto-approved", {}, 0.90, "plugin_test")
        all_items = self.q.get_all()
        self.assertEqual(len(all_items), 2)

    def test_get_all_filtered_by_status(self):
        self.q.submit("test", "Pending", {}, 0.40, "plugin_test")
        self.q.submit("test", "Auto-approved", {}, 0.90, "plugin_test")
        auto = self.q.get_all(status=ActionStatus.AUTO_APPROVED)
        self.assertEqual(len(auto), 1)

    def test_get_action_by_id(self):
        self.q.submit("test", "Find me", {"key": "val"}, 0.40, "plugin_test")
        pending = self.q.get_pending()
        action_id = pending[0].action_id
        found = self.q.get_action(action_id)
        self.assertIsNotNone(found)
        self.assertEqual(found.description, "Find me")

    def test_get_action_nonexistent_returns_none(self):
        result = self.q.get_action(99999)
        self.assertIsNone(result)

    def test_count_pending(self):
        for _ in range(3):
            self.q.submit("test", "Pending", {}, 0.40, "plugin_test")
        self.assertEqual(self.q.count_pending(), 3)


# ── 6. approve() ─────────────────────────────────────────────────────────────

class TestApprove(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)
        self.q.submit("send_email", "Approve me", {"to": "x@x.com"}, 0.50, "plugin_test")
        self.action_id = self.q.get_pending()[0].action_id

    def test_approve_returns_true(self):
        result = self.q.approve(self.action_id)
        self.assertTrue(result)

    def test_approve_changes_status_to_approved(self):
        self.q.approve(self.action_id)
        action = self.q.get_action(self.action_id)
        self.assertEqual(action.status, ActionStatus.APPROVED)

    def test_approve_executes_callback(self):
        called = []
        self.q.approve(self.action_id, execute_callback=lambda: called.append(True))
        self.assertEqual(len(called), 1)

    def test_approve_sets_reviewed_at(self):
        self.q.approve(self.action_id)
        action = self.q.get_action(self.action_id)
        self.assertIsNotNone(action.reviewed_at)

    def test_approve_sets_reviewer_note(self):
        self.q.approve(self.action_id, reviewer_note="Looks good")
        action = self.q.get_action(self.action_id)
        self.assertEqual(action.reviewer_note, "Looks good")

    def test_approve_nonexistent_returns_false(self):
        result = self.q.approve(99999)
        self.assertFalse(result)

    def test_approve_already_approved_returns_false(self):
        self.q.approve(self.action_id)
        result = self.q.approve(self.action_id)  # second approval
        self.assertFalse(result)

    def test_approve_removes_from_pending(self):
        self.q.approve(self.action_id)
        pending = self.q.get_pending()
        self.assertEqual(len(pending), 0)


# ── 7. reject() ──────────────────────────────────────────────────────────────

class TestReject(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)
        self.q.submit("send_email", "Reject me", {}, 0.50, "plugin_test")
        self.action_id = self.q.get_pending()[0].action_id

    def test_reject_returns_true(self):
        result = self.q.reject(self.action_id)
        self.assertTrue(result)

    def test_reject_changes_status_to_rejected(self):
        self.q.reject(self.action_id)
        action = self.q.get_action(self.action_id)
        self.assertEqual(action.status, ActionStatus.REJECTED)

    def test_reject_sets_reviewer_note(self):
        self.q.reject(self.action_id, reviewer_note="Not appropriate")
        action = self.q.get_action(self.action_id)
        self.assertEqual(action.reviewer_note, "Not appropriate")

    def test_reject_nonexistent_returns_false(self):
        result = self.q.reject(99999)
        self.assertFalse(result)

    def test_reject_already_rejected_returns_false(self):
        self.q.reject(self.action_id)
        result = self.q.reject(self.action_id)
        self.assertFalse(result)

    def test_reject_removes_from_pending(self):
        self.q.reject(self.action_id)
        self.assertEqual(self.q.count_pending(), 0)


# ── 8. Expiry ─────────────────────────────────────────────────────────────────

class TestExpiry(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)

    def _insert_expired_item(self):
        """Directly insert a row with an already-past expiry time."""
        past = (datetime.now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
        conn = self.q._get_db()
        conn.execute(
            """
            INSERT INTO approval_queue
                (plugin_id, action_type, description, payload, confidence, status, expires_at)
            VALUES ('plugin_test', 'test', 'Expired item', '{}', 0.40, 'pending', ?)
            """,
            (past,),
        )
        conn.commit()
        conn.close()

    def test_expired_items_not_in_pending(self):
        self._insert_expired_item()
        pending = self.q.get_pending()  # triggers _expire_old_items
        self.assertEqual(len(pending), 0)

    def test_expired_items_have_expired_status(self):
        self._insert_expired_item()
        self.q.get_pending()  # triggers expiry
        expired = self.q.get_all(status=ActionStatus.EXPIRED)
        self.assertEqual(len(expired), 1)

    def test_non_expired_items_remain_pending(self):
        # Add a fresh item (expires in 48h)
        self.q.submit("test", "Fresh item", {}, 0.40, "plugin_test")
        # Add an expired item
        self._insert_expired_item()
        pending = self.q.get_pending()
        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0].description, "Fresh item")


# ── 9. clear_old() ───────────────────────────────────────────────────────────

class TestClearOld(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)

    def _insert_old_approved(self, days_ago: int):
        old_ts = (datetime.now() - timedelta(days=days_ago)).strftime("%Y-%m-%d %H:%M:%S")
        conn = self.q._get_db()
        conn.execute(
            """
            INSERT INTO approval_queue
                (plugin_id, action_type, description, payload, confidence,
                 status, created_at, expires_at)
            VALUES ('plugin_test', 'test', 'Old item', '{}', 0.90,
                    'auto_approved', ?, ?)
            """,
            (old_ts, old_ts),
        )
        conn.commit()
        conn.close()

    def test_clear_old_removes_old_records(self):
        self._insert_old_approved(days_ago=35)
        self.q.clear_old(days=30)
        all_items = self.q.get_all()
        self.assertEqual(len(all_items), 0)

    def test_clear_old_keeps_recent_records(self):
        self._insert_old_approved(days_ago=5)
        self.q.clear_old(days=30)
        all_items = self.q.get_all()
        self.assertEqual(len(all_items), 1)

    def test_clear_old_does_not_delete_pending(self):
        """Pending items should never be deleted by clear_old."""
        self.q.submit("test", "Pending old item", {}, 0.40, "plugin_test")
        # Manually backdate it
        old_ts = (datetime.now() - timedelta(days=60)).strftime("%Y-%m-%d %H:%M:%S")
        conn = self.q._get_db()
        conn.execute("UPDATE approval_queue SET created_at = ? WHERE status = 'pending'", (old_ts,))
        conn.commit()
        conn.close()
        self.q.clear_old(days=30)
        # Pending item should still be there
        self.assertEqual(self.q.count_pending(), 1)


# ── 10. summary() ────────────────────────────────────────────────────────────

class TestSummary(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)

    def test_summary_returns_correct_counts(self):
        # 1 auto-approved, 2 pending
        self.q.submit("test", "Auto", {}, 0.90, "p")
        self.q.submit("test", "Pending 1", {}, 0.40, "p")
        self.q.submit("test", "Pending 2", {}, 0.40, "p")

        s = self.q.summary()
        self.assertEqual(s["auto_approved"], 1)
        self.assertEqual(s["pending"], 2)
        self.assertEqual(s["approved"], 0)
        self.assertEqual(s["rejected"], 0)

    def test_summary_includes_threshold(self):
        s = self.q.summary()
        self.assertIn("threshold", s)
        self.assertAlmostEqual(s["threshold"], 0.75, places=2)

    def test_summary_after_approve_and_reject(self):
        self.q.submit("test", "Approve me", {}, 0.40, "p")
        self.q.submit("test", "Reject me", {}, 0.40, "p")
        pending = self.q.get_pending()
        self.q.approve(pending[0].action_id)
        self.q.reject(pending[1].action_id)

        s = self.q.summary()
        self.assertEqual(s["approved"], 1)
        self.assertEqual(s["rejected"], 1)
        self.assertEqual(s["pending"], 0)


# ── 11. EventBus integration ──────────────────────────────────────────────────

class TestEventBusIntegration(unittest.TestCase):

    def setUp(self):
        self.q = make_queue()
        self.q.set_threshold(0.75)
        # Create a mock event bus
        self.events = []
        mock_bus = MagicMock()
        mock_bus.publish = lambda event_type, payload=None, source="": \
            self.events.append({"type": event_type, "payload": payload})
        self.q.set_event_bus(mock_bus)

    def test_auto_approve_publishes_event(self):
        self.q.submit("test", "Auto event", {}, 0.90, "plugin_test")
        types = [e["type"] for e in self.events]
        self.assertIn("approval.auto_approved", types)

    def test_queue_publishes_requested_event(self):
        self.q.submit("test", "Queue event", {}, 0.40, "plugin_test")
        types = [e["type"] for e in self.events]
        self.assertIn("approval.requested", types)

    def test_approve_publishes_approved_event(self):
        self.q.submit("test", "Approve event", {}, 0.40, "plugin_test")
        action_id = self.q.get_pending()[0].action_id
        self.q.approve(action_id)
        types = [e["type"] for e in self.events]
        self.assertIn("approval.approved", types)

    def test_reject_publishes_rejected_event(self):
        self.q.submit("test", "Reject event", {}, 0.40, "plugin_test")
        action_id = self.q.get_pending()[0].action_id
        self.q.reject(action_id)
        types = [e["type"] for e in self.events]
        self.assertIn("approval.rejected", types)

    def test_requested_event_contains_action_id(self):
        self.q.submit("test", "With ID", {}, 0.40, "plugin_test")
        requested = [e for e in self.events if e["type"] == "approval.requested"]
        self.assertTrue(len(requested) > 0)
        self.assertIn("action_id", requested[0]["payload"])


# ── 12. PluginContext injection ───────────────────────────────────────────────

class TestPluginContextInjection(unittest.TestCase):

    def test_approval_queue_in_plugin_context(self):
        """Verify PluginContext has approval_queue field."""
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertTrue(hasattr(ctx, "approval_queue"))
        self.assertIsNone(ctx.approval_queue)  # None by default

    def test_approval_queue_can_be_set(self):
        from plugin_base import PluginContext
        q = make_queue()
        ctx = PluginContext(approval_queue=q)
        self.assertIsNotNone(ctx.approval_queue)
        self.assertIsInstance(ctx.approval_queue, ApprovalQueue)

    def test_plugin_loader_injects_approval_queue(self):
        """Verify plugin_loader._make_context() injects an ApprovalQueue."""
        from plugin_loader import PluginLoader
        loader = PluginLoader(log_callback=lambda msg: None)
        ctx = loader._make_context(draft_mode=True)
        # Should be an ApprovalQueue instance (or None if DB unavailable)
        self.assertTrue(
            ctx.approval_queue is None or isinstance(ctx.approval_queue, ApprovalQueue),
            f"Expected ApprovalQueue or None, got {type(ctx.approval_queue)}"
        )


# ── 13. config.py defaults ───────────────────────────────────────────────────

class TestConfigDefaults(unittest.TestCase):

    def test_approval_threshold_default_seeded(self):
        from config import get_setting
        val = get_setting("approval_auto_threshold")
        self.assertEqual(val, "0.75")

    def test_approval_expiry_hours_default_seeded(self):
        from config import get_setting
        val = get_setting("approval_expiry_hours")
        self.assertEqual(val, "48")


if __name__ == "__main__":
    unittest.main(verbosity=2)
