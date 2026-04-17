"""
MC & S Desktop Agent — Confidence-Based Approval Queue (Stream 6)
==================================================================
Implements configurable autonomy: plugins can submit proposed actions
with a confidence score. Actions above the auto-approve threshold execute
immediately; actions below it are queued for human review.

DESIGN
------
  ApprovalQueue  — singleton that manages the SQLite-backed queue
  PendingAction  — dataclass representing a queued action
  ActionStatus   — enum: PENDING / APPROVED / REJECTED / EXPIRED / AUTO_APPROVED

WORKFLOW
--------
  1. Plugin calls: queue.submit(action_type, description, payload,
                                confidence, plugin_id, context)
  2. If confidence >= auto_approve_threshold → execute immediately,
     log as AUTO_APPROVED, return True.
  3. If confidence < threshold → insert into approval_queue table,
     publish "approval.requested" event, return False.
  4. UI polls queue.get_pending() and shows items for review.
  5. User clicks Approve → queue.approve(action_id) → execute callback,
     mark APPROVED.
  6. User clicks Reject  → queue.reject(action_id)  → mark REJECTED.
  7. Items older than expiry_hours are marked EXPIRED on next poll.

CONFIDENCE SCALE
----------------
  0.0 – 0.49  → Always requires approval (low confidence)
  0.50 – 0.74 → Requires approval unless threshold is set lower
  0.75 – 0.89 → Auto-approved at default threshold (0.75)
  0.90 – 1.0  → Always auto-approved

DEFAULT THRESHOLD
-----------------
  Stored in settings as "approval_auto_threshold" (default: "0.75").
  Configurable per-plugin via plugin settings.
"""

import sqlite3
import json
import threading
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Optional


# ── Status enum ───────────────────────────────────────────────────────────────

class ActionStatus(str, Enum):
    PENDING       = "pending"
    APPROVED      = "approved"
    REJECTED      = "rejected"
    EXPIRED       = "expired"
    AUTO_APPROVED = "auto_approved"


# ── PendingAction dataclass ───────────────────────────────────────────────────

@dataclass
class PendingAction:
    """Represents a proposed action awaiting human review."""
    action_id: int
    plugin_id: str
    action_type: str
    description: str
    payload: dict
    confidence: float
    status: ActionStatus
    created_at: str
    expires_at: str
    reviewed_at: Optional[str] = None
    reviewer_note: Optional[str] = None
    # Callback to execute when approved (not persisted to DB)
    _execute_callback: Optional[Callable] = field(default=None, repr=False, compare=False)


# ── ApprovalQueue ─────────────────────────────────────────────────────────────

class ApprovalQueue:
    """
    Singleton approval queue backed by SQLite.

    Usage:
        from approval_queue import ApprovalQueue
        queue = ApprovalQueue()

        # In a plugin's run() method:
        approved = queue.submit(
            action_type="send_email",
            description="Send tax return reminder to john@example.com",
            payload={"to": "john@example.com", "subject": "Tax Return Reminder"},
            confidence=0.82,
            plugin_id="plugin_client_outreach",
            execute_callback=lambda: context.graph.send_email(...),
        )
        if approved:
            result.actions_taken += 1
        else:
            result.items_skipped += 1  # queued for review
    """

    DEFAULT_THRESHOLD = 0.75
    DEFAULT_EXPIRY_HOURS = 48

    def __init__(self, db_path: Optional[Path] = None):
        if db_path is None:
            from config import DB_PATH
            db_path = DB_PATH
        self._db_path = db_path
        self._lock = threading.Lock()
        self._event_bus = None  # injected after init
        self._ensure_table()

    def set_event_bus(self, event_bus):
        """Inject the EventBus so the queue can publish events."""
        self._event_bus = event_bus

    # ── DB setup ──────────────────────────────────────────────────────────────

    def _get_db(self):
        Path(self._db_path).parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(self._db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _ensure_table(self):
        """Create the approval_queue table if it doesn't exist."""
        conn = self._get_db()
        conn.execute("""
            CREATE TABLE IF NOT EXISTS approval_queue (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                plugin_id     TEXT NOT NULL,
                action_type   TEXT NOT NULL,
                description   TEXT NOT NULL,
                payload       TEXT NOT NULL DEFAULT '{}',
                confidence    REAL NOT NULL DEFAULT 0.0,
                status        TEXT NOT NULL DEFAULT 'pending',
                created_at    TEXT DEFAULT (datetime('now','localtime')),
                expires_at    TEXT,
                reviewed_at   TEXT,
                reviewer_note TEXT
            )
        """)
        conn.commit()
        conn.close()

    # ── Threshold helpers ─────────────────────────────────────────────────────

    def get_threshold(self) -> float:
        """Return the current auto-approve confidence threshold."""
        try:
            from config import get_setting
            val = get_setting("approval_auto_threshold", str(self.DEFAULT_THRESHOLD))
            return float(val)
        except Exception:
            return self.DEFAULT_THRESHOLD

    def set_threshold(self, threshold: float):
        """Update the auto-approve confidence threshold."""
        threshold = max(0.0, min(1.0, threshold))
        try:
            from config import set_setting
            set_setting("approval_auto_threshold", str(threshold))
        except Exception:
            pass

    def get_expiry_hours(self) -> int:
        """Return how many hours before pending items expire."""
        try:
            from config import get_setting
            val = get_setting("approval_expiry_hours", str(self.DEFAULT_EXPIRY_HOURS))
            return int(val)
        except Exception:
            return self.DEFAULT_EXPIRY_HOURS

    # ── Submit ────────────────────────────────────────────────────────────────

    def submit(
        self,
        action_type: str,
        description: str,
        payload: dict,
        confidence: float,
        plugin_id: str = "unknown",
        execute_callback: Optional[Callable] = None,
    ) -> bool:
        """
        Submit a proposed action for approval or auto-execution.

        Returns:
            True  — action was auto-approved and executed (or confidence >= threshold)
            False — action was queued for human review
        """
        confidence = max(0.0, min(1.0, float(confidence)))
        threshold = self.get_threshold()

        if confidence >= threshold:
            # Auto-approve: execute immediately and log
            self._log_auto_approved(
                plugin_id, action_type, description, payload, confidence
            )
            if execute_callback:
                try:
                    execute_callback()
                except Exception as e:
                    # Log execution failure but don't re-queue
                    self._log_execution_error(plugin_id, action_type, description, str(e))
            if self._event_bus:
                self._event_bus.publish(
                    "approval.auto_approved",
                    {
                        "plugin_id": plugin_id,
                        "action_type": action_type,
                        "description": description,
                        "confidence": confidence,
                    },
                    source="approval_queue",
                )
            return True
        else:
            # Queue for human review
            action_id = self._insert_pending(
                plugin_id, action_type, description, payload, confidence
            )
            if self._event_bus:
                self._event_bus.publish(
                    "approval.requested",
                    {
                        "action_id": action_id,
                        "plugin_id": plugin_id,
                        "action_type": action_type,
                        "description": description,
                        "confidence": confidence,
                    },
                    source="approval_queue",
                )
            return False

    def _insert_pending(
        self,
        plugin_id: str,
        action_type: str,
        description: str,
        payload: dict,
        confidence: float,
    ) -> int:
        expiry = (
            datetime.now() + timedelta(hours=self.get_expiry_hours())
        ).strftime("%Y-%m-%d %H:%M:%S")

        with self._lock:
            conn = self._get_db()
            cur = conn.execute(
                """
                INSERT INTO approval_queue
                    (plugin_id, action_type, description, payload, confidence,
                     status, expires_at)
                VALUES (?, ?, ?, ?, ?, 'pending', ?)
                """,
                (
                    plugin_id,
                    action_type,
                    description,
                    json.dumps(payload),
                    confidence,
                    expiry,
                ),
            )
            action_id = cur.lastrowid
            conn.commit()
            conn.close()
        return action_id

    def _log_auto_approved(
        self,
        plugin_id: str,
        action_type: str,
        description: str,
        payload: dict,
        confidence: float,
    ):
        expiry = (
            datetime.now() + timedelta(hours=self.get_expiry_hours())
        ).strftime("%Y-%m-%d %H:%M:%S")
        reviewed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with self._lock:
            conn = self._get_db()
            conn.execute(
                """
                INSERT INTO approval_queue
                    (plugin_id, action_type, description, payload, confidence,
                     status, expires_at, reviewed_at, reviewer_note)
                VALUES (?, ?, ?, ?, ?, 'auto_approved', ?, ?, 'Auto-approved by confidence threshold')
                """,
                (
                    plugin_id,
                    action_type,
                    description,
                    json.dumps(payload),
                    confidence,
                    expiry,
                    reviewed_at,
                ),
            )
            conn.commit()
            conn.close()

    def _log_execution_error(
        self,
        plugin_id: str,
        action_type: str,
        description: str,
        error: str,
    ):
        with self._lock:
            conn = self._get_db()
            conn.execute(
                """
                INSERT INTO approval_queue
                    (plugin_id, action_type, description, payload, confidence,
                     status, reviewer_note)
                VALUES (?, ?, ?, '{}', 0.0, 'rejected', ?)
                """,
                (plugin_id, action_type, description, f"Execution error: {error}"),
            )
            conn.commit()
            conn.close()

    # ── Query ─────────────────────────────────────────────────────────────────

    # Alias so API server can call either name
    def list_pending(self) -> list[PendingAction]:
        """Alias for get_pending() — used by the API server."""
        return self.get_pending()

    def get_pending(self) -> list[PendingAction]:
        """Return all pending (not yet reviewed, not expired) actions."""
        self._expire_old_items()
        conn = self._get_db()
        rows = conn.execute(
            """
            SELECT * FROM approval_queue
            WHERE status = 'pending'
            ORDER BY created_at DESC
            """
        ).fetchall()
        conn.close()
        return [self._row_to_action(r) for r in rows]

    def edit_payload(self, action_id: int, updated_payload: dict) -> bool:
        """
        Update the payload of a pending action before approving (edit-before-approve).
        Returns True on success, False if not found or not pending.
        """
        with self._lock:
            conn = self._get_db()
            row = conn.execute(
                "SELECT id FROM approval_queue WHERE id = ? AND status = 'pending'",
                (action_id,),
            ).fetchone()
            if not row:
                conn.close()
                return False
            conn.execute(
                "UPDATE approval_queue SET payload = ? WHERE id = ?",
                (json.dumps(updated_payload), action_id),
            )
            conn.commit()
            conn.close()
        return True

    def get_all(
        self,
        status: Optional[ActionStatus] = None,
        limit: int = 100,
    ) -> list[PendingAction]:
        """Return all actions, optionally filtered by status."""
        self._expire_old_items()
        conn = self._get_db()
        if status:
            rows = conn.execute(
                """
                SELECT * FROM approval_queue
                WHERE status = ?
                ORDER BY created_at DESC
                LIMIT ?
                """,
                (status.value, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT * FROM approval_queue
                ORDER BY created_at DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        conn.close()
        return [self._row_to_action(r) for r in rows]

    def get_action(self, action_id: int) -> Optional[PendingAction]:
        """Return a single action by ID."""
        conn = self._get_db()
        row = conn.execute(
            "SELECT * FROM approval_queue WHERE id = ?", (action_id,)
        ).fetchone()
        conn.close()
        return self._row_to_action(row) if row else None

    def count_pending(self) -> int:
        """Return the number of pending actions."""
        conn = self._get_db()
        count = conn.execute(
            "SELECT COUNT(*) FROM approval_queue WHERE status = 'pending'"
        ).fetchone()[0]
        conn.close()
        return count

    # ── Review ────────────────────────────────────────────────────────────────

    def approve(
        self,
        action_id: int,
        reviewer_note: str = "",
        execute_callback: Optional[Callable] = None,
    ) -> bool:
        """
        Approve a pending action. Executes the callback if provided.
        Returns True on success, False if action not found or not pending.
        """
        with self._lock:
            conn = self._get_db()
            row = conn.execute(
                "SELECT * FROM approval_queue WHERE id = ? AND status = 'pending'",
                (action_id,),
            ).fetchone()
            if not row:
                conn.close()
                return False

            reviewed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                """
                UPDATE approval_queue
                SET status = 'approved', reviewed_at = ?, reviewer_note = ?
                WHERE id = ?
                """,
                (reviewed_at, reviewer_note or "Manually approved", action_id),
            )
            conn.commit()
            conn.close()

        if execute_callback:
            try:
                execute_callback()
            except Exception as e:
                self._mark_execution_failed(action_id, str(e))
                return False

        # Fix 17: store reviewer notes as lessons so AI learns from approvals
        if reviewer_note and reviewer_note not in ("Manually approved",):
            try:
                from memory_store import MemoryStore
                MemoryStore.store_lesson(
                    lesson=reviewer_note,
                    source=f"approval_queue:approved:{action_id}",
                )
            except Exception:
                pass  # Lesson storage is best-effort

        if self._event_bus:
            self._event_bus.publish(
                "approval.approved",
                {"action_id": action_id, "reviewer_note": reviewer_note},
                source="approval_queue",
            )
        return True

    def reject(self, action_id: int, reviewer_note: str = "") -> bool:
        """
        Reject a pending action.
        Returns True on success, False if action not found or not pending.
        """
        with self._lock:
            conn = self._get_db()
            row = conn.execute(
                "SELECT * FROM approval_queue WHERE id = ? AND status = 'pending'",
                (action_id,),
            ).fetchone()
            if not row:
                conn.close()
                return False

            reviewed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                """
                UPDATE approval_queue
                SET status = 'rejected', reviewed_at = ?, reviewer_note = ?
                WHERE id = ?
                """,
                (reviewed_at, reviewer_note or "Manually rejected", action_id),
            )
            conn.commit()
            conn.close()

        # Fix 17: store rejection reason as a lesson so AI learns from rejections
        if reviewer_note and reviewer_note not in ("Manually rejected",):
            try:
                from memory_store import MemoryStore
                MemoryStore.store_lesson(
                    lesson=reviewer_note,
                    source=f"approval_queue:rejected:{action_id}",
                )
            except Exception:
                pass  # Lesson storage is best-effort

        if self._event_bus:
            self._event_bus.publish(
                "approval.rejected",
                {"action_id": action_id, "reviewer_note": reviewer_note},
                source="approval_queue",
            )
        return True

    def _mark_execution_failed(self, action_id: int, error: str):
        with self._lock:
            conn = self._get_db()
            conn.execute(
                """
                UPDATE approval_queue
                SET status = 'rejected', reviewer_note = ?
                WHERE id = ?
                """,
                (f"Execution failed: {error}", action_id),
            )
            conn.commit()
            conn.close()

    # ── Expiry ────────────────────────────────────────────────────────────────

    def _expire_old_items(self):
        """Mark pending items that have passed their expiry time as expired."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._lock:
            conn = self._get_db()
            conn.execute(
                """
                UPDATE approval_queue
                SET status = 'expired'
                WHERE status = 'pending' AND expires_at IS NOT NULL AND expires_at < ?
                """,
                (now,),
            )
            conn.commit()
            conn.close()

    def clear_old(self, days: int = 30):
        """Delete records older than N days to keep the DB tidy."""
        cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
        with self._lock:
            conn = self._get_db()
            conn.execute(
                "DELETE FROM approval_queue WHERE created_at < ? AND status != 'pending'",
                (cutoff,),
            )
            conn.commit()
            conn.close()

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _row_to_action(self, row) -> PendingAction:
        return PendingAction(
            action_id=row["id"],
            plugin_id=row["plugin_id"],
            action_type=row["action_type"],
            description=row["description"],
            payload=json.loads(row["payload"] or "{}"),
            confidence=float(row["confidence"]),
            status=ActionStatus(row["status"]),
            created_at=row["created_at"] or "",
            expires_at=row["expires_at"] or "",
            reviewed_at=row["reviewed_at"],
            reviewer_note=row["reviewer_note"],
        )

    def summary(self) -> dict:
        """Return a summary dict of queue statistics."""
        conn = self._get_db()
        rows = conn.execute(
            """
            SELECT status, COUNT(*) as cnt
            FROM approval_queue
            GROUP BY status
            """
        ).fetchall()
        conn.close()
        counts = {r["status"]: r["cnt"] for r in rows}
        return {
            "pending":       counts.get("pending", 0),
            "approved":      counts.get("approved", 0),
            "rejected":      counts.get("rejected", 0),
            "expired":       counts.get("expired", 0),
            "auto_approved": counts.get("auto_approved", 0),
            "threshold":     self.get_threshold(),
        }


# ── Module-level singleton ────────────────────────────────────────────────────

_queue_instance: Optional[ApprovalQueue] = None
_queue_lock = threading.Lock()


def get_approval_queue() -> ApprovalQueue:
    """Return the module-level ApprovalQueue singleton."""
    global _queue_instance
    if _queue_instance is None:
        with _queue_lock:
            if _queue_instance is None:
                _queue_instance = ApprovalQueue()
    return _queue_instance
