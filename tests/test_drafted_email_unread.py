"""
Regression tests for the post-draft "mark unread, don't flag" behaviour.

Pins the contract from task_unread_instead_of_flag.md:

  * After CoWorker drafts a reply, the original email must be marked
    unread (isRead = False) — that bold-row state is the cue the
    accountant uses to know a draft is waiting.
  * The Drafted category stays — the visible text label that names
    *what* is waiting.
  * The red flag has been retired. The post-draft PATCH payload must
    NOT contain a `flag` key.

These tests mock `requests.patch` directly so we are pinning the wire
shape that hits Graph, not just our Python wrappers.
"""
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

from graph_client import GraphClient  # noqa: E402


def _make_client() -> GraphClient:
    """Build a GraphClient whose _headers() short-circuits to a fake token
    so mark_as_unread() never tries to refresh against MSAL."""
    client = GraphClient.__new__(GraphClient)
    client._headers = MagicMock(return_value={
        "Authorization": "Bearer test-token",
        "Content-Type": "application/json",
    })
    return client


class MarkAsUnreadPatchPayloadTests(unittest.TestCase):
    """The PATCH that fires after a draft is created."""

    def test_payload_marks_isread_false(self):
        client = _make_client()
        with patch("graph_client.requests.patch") as mock_patch:
            mock_patch.return_value.raise_for_status = MagicMock()
            client.mark_as_unread("AAMkA-test-message-id")

        self.assertEqual(mock_patch.call_count, 1)
        payload = mock_patch.call_args.kwargs["json"]
        self.assertEqual(payload, {"isRead": False})

    def test_payload_has_no_flag_key(self):
        """Regression: drafted emails must not be flagged anymore. The
        red flag was redundant with the Drafted category and visually
        noisy in the inbox."""
        client = _make_client()
        with patch("graph_client.requests.patch") as mock_patch:
            mock_patch.return_value.raise_for_status = MagicMock()
            client.mark_as_unread("AAMkA-test-message-id")

        payload = mock_patch.call_args.kwargs["json"]
        self.assertNotIn(
            "flag", payload,
            "Drafted emails must not carry a flag — unread is the cue.",
        )
        self.assertNotIn("flagStatus", str(payload))


class PostDraftCallSitesUseUnreadNotFlagTests(unittest.TestCase):
    """Source-level guard: the three plugins that draft replies must not
    call flag_email() on the post-draft path. If anyone reintroduces
    flag_email() into a drafted-reply flow these assertions fail."""

    PLUGINS_DIR = REPO_ROOT / "plugins"

    def _read(self, name: str) -> str:
        return (self.PLUGINS_DIR / name).read_text(encoding="utf-8")

    def test_smart_responder_does_not_flag_after_drafting(self):
        src = self._read("plugin_smart_responder.py")
        self.assertNotIn(
            "graph.flag_email", src,
            "smart_responder must mark unread, not flag, after drafting.",
        )
        self.assertIn("graph.mark_as_unread", src)
        self.assertIn('graph.add_category(message_id, "Drafted")', src)

    def test_engagement_letter_does_not_flag_after_drafting(self):
        src = self._read("plugin_engagement_letter.py")
        self.assertNotIn(
            "context.graph.flag_email", src,
            "engagement_letter must mark unread, not flag, after drafting.",
        )
        self.assertIn("context.graph.mark_as_unread", src)

    def test_noa_processor_does_not_flag_drafted_emails(self):
        src = self._read("plugin_noa_processor.py")
        # NOA still flags on extraction-failure / missing-email paths
        # (those are NOT drafted emails). The post-draft call site must
        # use mark_as_unread now. Pin that the post-draft block contains
        # mark_as_unread, not flag_email.
        post_draft_marker = "draft has already been created"
        self.assertIn(post_draft_marker, src)
        idx = src.index(post_draft_marker)
        post_draft_block = src[idx:idx + 400]
        self.assertIn("mark_as_unread", post_draft_block)
        self.assertNotIn("flag_email", post_draft_block)


if __name__ == "__main__":
    unittest.main()
