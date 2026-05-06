"""
Tests for the chat persistence endpoints (Fix A2).

Pins the contract laid out in task_chat_persistence_and_archive.md:
  - GET /api/chat/<agent_id>/active returns an empty payload before any turn
  - GET returns the persisted turns once a session exists
  - Two POST /api/chat calls in the same (user, agent) share one session
    and produce 4 messages (user/assistant interleaved)
  - When the Anthropic call raises, NO chat_sessions or chat_messages rows
    are written — persistence happens only after a successful model call
  - Different agents land in different sessions with no cross-talk

The Anthropic SDK is mocked at the module level (api_server.chat() does
``import anthropic as anthropic_lib`` inside the function, then calls
``anthropic_lib.Anthropic(api_key=...)``). Patching ``anthropic.Anthropic``
replaces the class lookup that import sees.
"""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

# Same DB-override pattern as test_schedule.py / test_chat_sessions_schema.py:
# point config.DB_PATH at a temp file before anything else imports config or
# api_server, so the Flask app boots against a clean DB.
import config as cfg  # noqa: E402

_OUR_DB_PATH = Path(tempfile.mktemp(suffix=".db"))
cfg.DB_PATH = _OUR_DB_PATH
cfg.init_db()

# Seed the auth token + identity so the Flask before_request gate lets
# requests through and _resolve_chat_user_id() returns a stable value.
cfg.set_setting("local_api_token", "test-token")
cfg.set_setting("ms_account_email", "elio@mcands.com.au")
# Non-empty so the chat handler doesn't short-circuit with "API key not
# configured" — the actual key value is irrelevant since we mock the client.
cfg.set_setting("anthropic_api_key", "sk-test-not-real")

import api_server  # noqa: E402


def _make_fake_anthropic_client(reply_text: str) -> MagicMock:
    """Build a MagicMock that mimics the bits of anthropic.Anthropic the
    chat handler touches: messages.create returns an object with a
    .content list of blocks (each with .text) plus a .usage object."""
    block = MagicMock()
    block.text = reply_text
    response = MagicMock()
    response.content = [block]
    response.usage.input_tokens = 10
    response.usage.output_tokens = 20
    client = MagicMock()
    client.messages.create.return_value = response
    return client


class ChatPersistenceEndpointTests(unittest.TestCase):
    HEADERS = {
        "Authorization": "Bearer test-token",
        "Content-Type": "application/json",
    }

    def setUp(self):
        # Other test modules also override config.DB_PATH at their own
        # module-import time. Pin ours back at the start of every test so
        # the Flask app and our setUp/asserts both consistently read from
        # the DB we seeded with the auth token + identity.
        cfg.DB_PATH = _OUR_DB_PATH
        self.client = api_server.app.test_client()
        # Wipe just the chat tables so UNIQUE constraints and counts start
        # clean per test, without touching seeded settings/links/etc.
        conn = cfg.get_db()
        try:
            conn.execute("DELETE FROM chat_messages")
            conn.execute("DELETE FROM chat_sessions")
            conn.commit()
        finally:
            conn.close()
        # Reset the in-process rate limit window — back-to-back tests would
        # otherwise eventually hit the 30-call ceiling.
        api_server._chat_timestamps.clear()

    # ── helpers ────────────────────────────────────────────────────────────
    def _post_chat(self, agent_id: str, user_text: str, reply_text: str):
        fake = _make_fake_anthropic_client(reply_text)
        with patch("anthropic.Anthropic", return_value=fake):
            return self.client.post(
                "/api/chat",
                headers=self.HEADERS,
                json={
                    "messages": [{"role": "user", "content": user_text}],
                    "agent_id": agent_id,
                },
            )

    def _all_sessions(self):
        conn = cfg.get_db()
        try:
            return [
                dict(r)
                for r in conn.execute(
                    "SELECT id, user_id, specialist_key, status FROM chat_sessions"
                ).fetchall()
            ]
        finally:
            conn.close()

    def _messages_for(self, session_id: str):
        conn = cfg.get_db()
        try:
            return [
                dict(r)
                for r in conn.execute(
                    "SELECT role, content FROM chat_messages WHERE session_id=? "
                    "ORDER BY created_at ASC, rowid ASC",
                    (session_id,),
                ).fetchall()
            ]
        finally:
            conn.close()

    # ── tests ──────────────────────────────────────────────────────────────
    def test_get_active_returns_empty_when_none_exists(self):
        resp = self.client.get("/api/chat/gst/active", headers=self.HEADERS)
        self.assertEqual(resp.status_code, 200)
        body = resp.get_json()
        self.assertTrue(body["ok"])
        self.assertIsNone(body["data"]["session_id"])
        self.assertEqual(body["data"]["messages"], [])

    def test_get_active_returns_persisted_messages_in_order(self):
        post_resp = self._post_chat("gst", "Q1", "A1")
        self.assertEqual(post_resp.status_code, 200)
        self.assertTrue(post_resp.get_json()["ok"])

        resp = self.client.get("/api/chat/gst/active", headers=self.HEADERS)
        body = resp.get_json()
        msgs = body["data"]["messages"]
        self.assertEqual(len(msgs), 2)
        self.assertEqual(msgs[0]["role"], "user")
        self.assertEqual(msgs[0]["content"], "Q1")
        self.assertEqual(msgs[1]["role"], "assistant")
        self.assertEqual(msgs[1]["content"], "A1")

    def test_two_consecutive_calls_share_session_with_four_messages(self):
        self._post_chat("gst", "Q1", "A1")
        self._post_chat("gst", "Q2", "A2")

        sessions = self._all_sessions()
        self.assertEqual(len(sessions), 1, "Both calls should share one session")
        msgs = self._messages_for(sessions[0]["id"])
        self.assertEqual(len(msgs), 4)
        self.assertEqual(
            [m["role"] for m in msgs],
            ["user", "assistant", "user", "assistant"],
        )
        self.assertEqual(
            [m["content"] for m in msgs],
            ["Q1", "A1", "Q2", "A2"],
        )

    def test_anthropic_failure_persists_no_rows(self):
        # Patch anthropic.Anthropic so messages.create raises. The chat
        # handler's outer try/except returns an error to the caller, and
        # _persist_chat_turn is never reached — so the DB must stay empty.
        fake = MagicMock()
        fake.messages.create.side_effect = RuntimeError("boom")
        with patch("anthropic.Anthropic", return_value=fake):
            resp = self.client.post(
                "/api/chat",
                headers=self.HEADERS,
                json={
                    "messages": [{"role": "user", "content": "Q"}],
                    "agent_id": "gst",
                },
            )

        body = resp.get_json()
        self.assertFalse(body["ok"], "Handler should report failure to caller")

        conn = cfg.get_db()
        try:
            session_count = conn.execute(
                "SELECT COUNT(*) FROM chat_sessions"
            ).fetchone()[0]
            message_count = conn.execute(
                "SELECT COUNT(*) FROM chat_messages"
            ).fetchone()[0]
        finally:
            conn.close()
        self.assertEqual(session_count, 0)
        self.assertEqual(message_count, 0)

    def test_different_agents_get_different_sessions_with_no_crosstalk(self):
        self._post_chat("gst", "GQ", "GA")
        self._post_chat("smsf", "SQ", "SA")

        sessions = sorted(self._all_sessions(), key=lambda s: s["specialist_key"])
        self.assertEqual(len(sessions), 2)
        self.assertEqual(
            [s["specialist_key"] for s in sessions],
            ["gst", "smsf"],
        )
        self.assertNotEqual(sessions[0]["id"], sessions[1]["id"])

        for s in sessions:
            msgs = self._messages_for(s["id"])
            if s["specialist_key"] == "gst":
                self.assertEqual([m["content"] for m in msgs], ["GQ", "GA"])
            else:
                self.assertEqual([m["content"] for m in msgs], ["SQ", "SA"])


if __name__ == "__main__":
    unittest.main()
