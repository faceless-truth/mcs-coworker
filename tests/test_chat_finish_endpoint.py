"""
Tests for POST /api/chat/<agent_id>/finish (Fix B4).

Pins the contract from fix_b4_finish_endpoint.md:

  1. plugin_builder -> 400, exact message.
  2. Unknown agent_id -> 400, message contains the offending id.
  3. No active session -> 404, exact message.
  4. Empty session (row exists, no messages) -> 400, exact message.
  5. Happy path -> 200; chat_sessions row flipped to archived with the
     SharePoint webUrl persisted; filename matches YYYY-MM-DD <Short> -
     <Title>.docx pattern.
  6. SharePointFolderMissing from upload -> 500; session stays active.
  7. Generic upload exception -> 500 with "SharePoint upload failed:"
     prefix; session stays active.
  8. Title-fallback path: when generate_archive_title returns "",
     filename ends with " conversation.docx".
  9. Post-archive, GET /active returns the empty payload.

The actual Graph PUT is never exercised here (manual smoke test handles
the live call). _get_graph and generate_archive_title / upload_chat_archive
are patched in-process via unittest.mock.
"""
import re
import sys
import tempfile
import unittest
import uuid
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg  # noqa: E402

# Same DB-pinning dance as the other api_server-importing tests so the
# Flask app boots against a temp DB and not the real coworker.db.
_OUR_DB_PATH = Path(tempfile.mktemp(suffix=".db"))
cfg.DB_PATH = _OUR_DB_PATH
cfg.init_db()

cfg.set_setting("local_api_token", "test-token")
cfg.set_setting("ms_account_email", "elio@mcands.com.au")
cfg.set_setting("anthropic_api_key", "sk-test-not-real")

import api_server  # noqa: E402

from graph_client import SharePointFolderMissing  # noqa: E402


HEADERS = {
    "Authorization": "Bearer test-token",
    "Content-Type": "application/json",
}


def _seed_session_with_messages(agent_id: str, contents: list[tuple[str, str]]) -> str:
    """Create an active chat_sessions row + interleaved chat_messages.

    Returns the session id. Bypasses the /api/chat persistence path so
    these tests stay decoupled from the chat handler's implementation.
    """
    session_id = str(uuid.uuid4())
    user_id = api_server._resolve_chat_user_id()
    conn = cfg.get_db()
    try:
        conn.execute(
            "INSERT INTO chat_sessions (id, user_id, specialist_key) "
            "VALUES (?, ?, ?)",
            (session_id, user_id, agent_id),
        )
        for role, content in contents:
            conn.execute(
                "INSERT INTO chat_messages (id, session_id, role, content, created_at) "
                "VALUES (?, ?, ?, ?, ?)",
                (
                    str(uuid.uuid4()),
                    session_id,
                    role,
                    content,
                    datetime.now().isoformat(timespec="microseconds"),
                ),
            )
        conn.commit()
    finally:
        conn.close()
    return session_id


def _seed_empty_session(agent_id: str) -> str:
    """Create a chat_sessions row with no messages — for the empty-session case."""
    session_id = str(uuid.uuid4())
    user_id = api_server._resolve_chat_user_id()
    conn = cfg.get_db()
    try:
        conn.execute(
            "INSERT INTO chat_sessions (id, user_id, specialist_key) "
            "VALUES (?, ?, ?)",
            (session_id, user_id, agent_id),
        )
        conn.commit()
    finally:
        conn.close()
    return session_id


def _row(session_id: str) -> dict | None:
    conn = cfg.get_db()
    try:
        r = conn.execute(
            "SELECT id, status, archived_at, archived_to_url "
            "FROM chat_sessions WHERE id=?",
            (session_id,),
        ).fetchone()
        return dict(r) if r else None
    finally:
        conn.close()


class FinishChatEndpointTests(unittest.TestCase):
    def setUp(self):
        cfg.DB_PATH = _OUR_DB_PATH
        self.client = api_server.app.test_client()
        conn = cfg.get_db()
        try:
            conn.execute("DELETE FROM chat_messages")
            conn.execute("DELETE FROM chat_sessions")
            conn.commit()
        finally:
            conn.close()
        api_server._chat_timestamps.clear()

    # ── Guard cases ─────────────────────────────────────────────────────
    def test_plugin_builder_returns_400_with_exact_message(self):
        resp = self.client.post(
            "/api/chat/plugin_builder/finish", headers=HEADERS,
        )
        self.assertEqual(resp.status_code, 400)
        self.assertEqual(
            resp.get_json(),
            {"error": "Plugin Builder sessions are not archived."},
        )

    def test_unknown_agent_id_returns_400_with_offending_id_in_message(self):
        resp = self.client.post(
            "/api/chat/not_a_real_agent/finish", headers=HEADERS,
        )
        self.assertEqual(resp.status_code, 400)
        body = resp.get_json()
        self.assertIn("not_a_real_agent", body["error"])
        self.assertIn("Unknown agent_id", body["error"])

    def test_no_active_session_returns_404(self):
        resp = self.client.post("/api/chat/gst/finish", headers=HEADERS)
        self.assertEqual(resp.status_code, 404)
        self.assertEqual(
            resp.get_json(),
            {"error": "No active session to archive"},
        )

    def test_empty_session_returns_400(self):
        _seed_empty_session("gst")
        resp = self.client.post("/api/chat/gst/finish", headers=HEADERS)
        self.assertEqual(resp.status_code, 400)
        self.assertEqual(
            resp.get_json(),
            {"error": "Cannot archive an empty session"},
        )

    # ── Happy path ──────────────────────────────────────────────────────
    def test_happy_path_archives_session_and_returns_metadata(self):
        session_id = _seed_session_with_messages("gst", [
            ("user", "What is the GST treatment of a going concern sale?"),
            ("assistant", "Going concern sales can be GST-free if conditions met."),
            ("user", "What conditions specifically?"),
        ])

        fake_gc = MagicMock()
        fake_gc.upload_chat_archive.return_value = {
            "web_url": "https://mcandscomau.sharepoint.com/sites/MCS354/foo.docx",
            "item_id": "abc-123",
        }

        with patch.object(api_server, "_get_graph", return_value=fake_gc), \
             patch.object(api_server, "generate_archive_title",
                          return_value="GST Going Concern Sale"):
            resp = self.client.post("/api/chat/gst/finish", headers=HEADERS)

        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertTrue(body["archived"])
        self.assertEqual(
            body["web_url"],
            "https://mcandscomau.sharepoint.com/sites/MCS354/foo.docx",
        )
        self.assertRegex(
            body["filename"],
            r"^\d{4}-\d{2}-\d{2} [^/]+ - .+\.docx$",
        )
        # The short name is "GST" for agent_id "gst" — pin the prefix shape
        # so a future rename of SPECIALIST_SHORT_NAMES surfaces here.
        self.assertIn(" GST - ", body["filename"])
        self.assertTrue(body["filename"].endswith(".docx"))

        # upload_chat_archive must have been called with the expected kwargs.
        kwargs = fake_gc.upload_chat_archive.call_args.kwargs
        self.assertEqual(kwargs["agent_id"], "gst")
        self.assertEqual(kwargs["filename"], body["filename"])
        self.assertIsInstance(kwargs["file_bytes"], (bytes, bytearray))
        self.assertGreater(len(kwargs["file_bytes"]), 0)

        # DB state: session flipped to archived, web_url persisted.
        row = _row(session_id)
        self.assertIsNotNone(row)
        self.assertEqual(row["status"], "archived")
        self.assertIsNotNone(row["archived_at"])
        self.assertEqual(
            row["archived_to_url"],
            "https://mcandscomau.sharepoint.com/sites/MCS354/foo.docx",
        )

    # ── Failure paths leave the session active ──────────────────────────
    def test_sharepoint_folder_missing_returns_500_and_session_stays_active(self):
        session_id = _seed_session_with_messages("gst", [
            ("user", "Q1"),
            ("assistant", "A1"),
        ])
        fake_gc = MagicMock()
        fake_gc.upload_chat_archive.side_effect = SharePointFolderMissing(
            "folder X not found"
        )
        with patch.object(api_server, "_get_graph", return_value=fake_gc), \
             patch.object(api_server, "generate_archive_title",
                          return_value="Some Title"):
            resp = self.client.post("/api/chat/gst/finish", headers=HEADERS)

        self.assertEqual(resp.status_code, 500)
        self.assertEqual(resp.get_json(), {"error": "folder X not found"})

        row = _row(session_id)
        self.assertEqual(row["status"], "active")
        self.assertIsNone(row["archived_at"])
        self.assertIsNone(row["archived_to_url"])

    def test_generic_upload_exception_returns_500_and_session_stays_active(self):
        session_id = _seed_session_with_messages("gst", [
            ("user", "Q1"),
            ("assistant", "A1"),
        ])
        fake_gc = MagicMock()
        fake_gc.upload_chat_archive.side_effect = RuntimeError("graph timeout")
        with patch.object(api_server, "_get_graph", return_value=fake_gc), \
             patch.object(api_server, "generate_archive_title",
                          return_value="Some Title"):
            resp = self.client.post("/api/chat/gst/finish", headers=HEADERS)

        self.assertEqual(resp.status_code, 500)
        body = resp.get_json()
        self.assertTrue(body["error"].startswith("SharePoint upload failed:"))
        self.assertIn("graph timeout", body["error"])

        row = _row(session_id)
        self.assertEqual(row["status"], "active")
        self.assertIsNone(row["archived_at"])

    # ── Title fallback ──────────────────────────────────────────────────
    def test_title_fallback_to_timestamp_conversation(self):
        _seed_session_with_messages("gst", [
            ("user", "Q1"),
            ("assistant", "A1"),
        ])
        fake_gc = MagicMock()
        fake_gc.upload_chat_archive.return_value = {
            "web_url": "https://example/foo.docx",
            "item_id": "xyz",
        }
        with patch.object(api_server, "_get_graph", return_value=fake_gc), \
             patch.object(api_server, "generate_archive_title", return_value=""):
            resp = self.client.post("/api/chat/gst/finish", headers=HEADERS)

        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertTrue(body["filename"].endswith(" conversation.docx"))
        # And HH-MM should appear immediately before " conversation.docx".
        self.assertRegex(body["filename"], r" \d{2}-\d{2} conversation\.docx$")

    # ── Post-archive GET semantics ──────────────────────────────────────
    def test_after_archive_get_active_returns_empty_payload(self):
        _seed_session_with_messages("gst", [
            ("user", "Q1"),
            ("assistant", "A1"),
        ])
        fake_gc = MagicMock()
        fake_gc.upload_chat_archive.return_value = {
            "web_url": "https://example/foo.docx",
            "item_id": "z",
        }
        with patch.object(api_server, "_get_graph", return_value=fake_gc), \
             patch.object(api_server, "generate_archive_title",
                          return_value="Title"):
            finish = self.client.post("/api/chat/gst/finish", headers=HEADERS)
        self.assertEqual(finish.status_code, 200)

        get_resp = self.client.get("/api/chat/gst/active", headers=HEADERS)
        self.assertEqual(get_resp.status_code, 200)
        body = get_resp.get_json()
        # The /active endpoint wraps in the ok() envelope.
        self.assertTrue(body["ok"])
        self.assertIsNone(body["data"]["session_id"])
        self.assertEqual(body["data"]["messages"], [])


if __name__ == "__main__":
    unittest.main()
