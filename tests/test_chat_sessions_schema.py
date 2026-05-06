"""
Tests for the chat_sessions / chat_messages schema (Fix A1).

Phase A of the chat persistence + SharePoint archive task. Pins:
  - both tables and their indexes get created by init_db()
  - inserting a session + N messages works
  - cascade delete removes child messages when a session is deleted
    (this exercises the new PRAGMA foreign_keys=ON in apply_wal_pragmas)
  - UNIQUE (user_id, specialist_key, status) blocks two ACTIVE sessions for
    the same user+specialist, but allows an archived sibling to coexist
"""
import sqlite3
import sys
import tempfile
import unittest
import uuid
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

# Match the convention used by test_schedule.py: override DB_PATH on the
# already-loaded config module rather than reloading sys.modules. Reloading
# config under a different MCS_DATA_DIR contaminates downstream tests in the
# same process — they end up reading from a tempdir that's been cleaned up.
import config as cfg  # noqa: E402

_tmp_db = tempfile.mktemp(suffix=".db")
cfg.DB_PATH = Path(_tmp_db)
cfg.init_db()


class ChatSessionsSchemaTests(unittest.TestCase):
    # Each test starts from a clean chat_sessions / chat_messages pair so
    # UNIQUE-constraint cases don't bleed into one another. We only wipe the
    # two new tables — the rest of the schema is shared with init_db()'s
    # seeded settings/links/etc. and doesn't need resetting.
    def setUp(self):
        conn = cfg.get_db()
        try:
            conn.execute("DELETE FROM chat_messages")
            conn.execute("DELETE FROM chat_sessions")
            conn.commit()
        finally:
            conn.close()

    # ── helpers ────────────────────────────────────────────────────────────
    def _insert_session(self, conn, user_id, specialist_key, status="active"):
        sid = str(uuid.uuid4())
        conn.execute(
            "INSERT INTO chat_sessions (id, user_id, specialist_key, status) "
            "VALUES (?, ?, ?, ?)",
            (sid, user_id, specialist_key, status),
        )
        return sid

    def _insert_message(self, conn, session_id, role, content):
        mid = str(uuid.uuid4())
        conn.execute(
            "INSERT INTO chat_messages (id, session_id, role, content) "
            "VALUES (?, ?, ?, ?)",
            (mid, session_id, role, content),
        )
        return mid

    # ── tests ──────────────────────────────────────────────────────────────
    def test_tables_and_indexes_exist(self):
        conn = cfg.get_db()
        try:
            tables = {
                r["name"]
                for r in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
            self.assertIn("chat_sessions", tables)
            self.assertIn("chat_messages", tables)

            indexes = {
                r["name"]
                for r in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='index'"
                ).fetchall()
            }
            self.assertIn("ix_chat_sessions_user_specialist", indexes)
            self.assertIn("ix_chat_messages_session_created", indexes)
        finally:
            conn.close()

    def test_foreign_keys_pragma_is_on(self):
        # Sanity-check the apply_wal_pragmas change: every connection
        # handed out by config.get_db() must have FK enforcement enabled,
        # otherwise the cascade-delete contract below is silently a no-op.
        conn = cfg.get_db()
        try:
            (fk_on,) = conn.execute("PRAGMA foreign_keys").fetchone()
            self.assertEqual(fk_on, 1)
        finally:
            conn.close()

    def test_insert_session_and_three_messages(self):
        conn = cfg.get_db()
        try:
            sid = self._insert_session(conn, "elio@mcands.com.au", "gst")
            for role, body in [
                ("user", "Is a new residential property subject to GST?"),
                ("assistant", "Generally yes — section 40-65 outlines..."),
                ("user", "What about a substantial renovation?"),
            ]:
                self._insert_message(conn, sid, role, body)
            conn.commit()

            count = conn.execute(
                "SELECT COUNT(*) FROM chat_messages WHERE session_id=?",
                (sid,),
            ).fetchone()[0]
            self.assertEqual(count, 3)

            # Confirm role + content round-trip intact. Ordering is a
            # retrieval concern handled by the API layer (Fix A2) where it
            # uses microsecond-precision timestamps; SQLite's
            # CURRENT_TIMESTAMP only has 1-second resolution so three
            # back-to-back inserts collide and ordering by created_at alone
            # is non-deterministic at this layer.
            rows = conn.execute(
                "SELECT role, content FROM chat_messages WHERE session_id=?",
                (sid,),
            ).fetchall()
            self.assertEqual(
                sorted((r["role"], r["content"]) for r in rows),
                sorted([
                    ("user", "Is a new residential property subject to GST?"),
                    ("assistant", "Generally yes — section 40-65 outlines..."),
                    ("user", "What about a substantial renovation?"),
                ]),
            )
        finally:
            conn.close()

    def test_cascade_delete_removes_messages(self):
        conn = cfg.get_db()
        try:
            sid = self._insert_session(conn, "elio@mcands.com.au", "smsf")
            for i in range(3):
                self._insert_message(conn, sid, "user", f"msg {i}")
            conn.commit()

            self.assertEqual(
                conn.execute(
                    "SELECT COUNT(*) FROM chat_messages WHERE session_id=?",
                    (sid,),
                ).fetchone()[0],
                3,
            )

            conn.execute("DELETE FROM chat_sessions WHERE id=?", (sid,))
            conn.commit()

            self.assertEqual(
                conn.execute(
                    "SELECT COUNT(*) FROM chat_messages WHERE session_id=?",
                    (sid,),
                ).fetchone()[0],
                0,
                "ON DELETE CASCADE should have removed child messages — "
                "if this is non-zero, PRAGMA foreign_keys=ON is not in "
                "effect for this connection.",
            )
        finally:
            conn.close()

    def test_unique_active_session_per_user_and_specialist(self):
        conn = cfg.get_db()
        try:
            self._insert_session(conn, "elio@mcands.com.au", "gst")
            conn.commit()

            with self.assertRaises(sqlite3.IntegrityError):
                self._insert_session(conn, "elio@mcands.com.au", "gst")
                conn.commit()
        finally:
            conn.close()

    def test_archived_sibling_does_not_block_new_active(self):
        # Once a session is archived its (user, specialist, status) tuple
        # changes from 'active' to 'archived', so the user can start a fresh
        # active conversation in the same specialist.
        conn = cfg.get_db()
        try:
            sid_old = self._insert_session(conn, "elio@mcands.com.au", "gst")
            conn.execute(
                "UPDATE chat_sessions SET status='archived' WHERE id=?",
                (sid_old,),
            )
            sid_new = self._insert_session(conn, "elio@mcands.com.au", "gst")
            conn.commit()
            self.assertNotEqual(sid_old, sid_new)

            statuses = sorted(
                r["status"]
                for r in conn.execute(
                    "SELECT status FROM chat_sessions WHERE user_id=? "
                    "AND specialist_key=?",
                    ("elio@mcands.com.au", "gst"),
                ).fetchall()
            )
            self.assertEqual(statuses, ["active", "archived"])
        finally:
            conn.close()


if __name__ == "__main__":
    unittest.main()
