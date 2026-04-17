"""
Stream 4 Test Suite — Gateway Client (XPM, FuseSign, Teams)
=============================================================
All tests use mocked HTTP so no real API credentials are needed.

Test groups:
  1.  GatewayClient facade — load, available_integrations, status_summary
  2.  XPMClient — is_configured, list_clients, get_client, get_client_by_email,
                  list_jobs, update_job_status, add_client_note, error handling
  3.  FuseSignClient — is_configured, create_envelope, get_envelope_status,
                       list_envelopes, void_envelope, resend_envelope,
                       list_templates, error handling
  4.  TeamsClient — is_configured, send_message, send_alert, error handling
  5.  PluginContext.gateway field
  6.  plugin_loader._make_context() injects GatewayClient
  7.  config.py — gateway credential defaults seeded
  8.  Graceful degradation — gateway=None when import fails
"""

import sys
import os
import json
import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch, MagicMock
from io import BytesIO

# ── Path setup ────────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg
cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
cfg.init_db()


# ── HTTP mock helpers ─────────────────────────────────────────────────────────

class MockHTTPResponse:
    """Minimal mock for urllib.request.urlopen response."""
    def __init__(self, data: dict | list, status: int = 200):
        self._data = json.dumps(data).encode()
        self.status = status

    def read(self):
        return self._data

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass


def mock_urlopen(response_data, status=200):
    """Return a context manager that yields a MockHTTPResponse."""
    return patch(
        "urllib.request.urlopen",
        return_value=MockHTTPResponse(response_data, status),
    )


# ── 1. GatewayClient facade ───────────────────────────────────────────────────

class TestGatewayClientFacade(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_gateway_client_instantiates(self):
        from gateway_client import GatewayClient
        gw = GatewayClient()
        self.assertIsNotNone(gw.xpm)
        self.assertIsNotNone(gw.fusesign)
        self.assertIsNotNone(gw.teams)

    def test_available_integrations_empty_when_not_configured(self):
        from gateway_client import GatewayClient
        gw = GatewayClient()
        available = gw.available_integrations()
        self.assertEqual(available, [])

    def test_available_integrations_returns_configured_ones(self):
        from gateway_client import GatewayClient
        cfg.set_setting("xpm_api_key", "test-xpm-key")
        cfg.set_setting("teams_webhook_url", "https://outlook.office.com/webhook/test")
        gw = GatewayClient()
        available = gw.available_integrations()
        self.assertIn("xpm", available)
        self.assertIn("teams", available)
        self.assertNotIn("fusesign", available)
        # Cleanup
        cfg.set_setting("xpm_api_key", "")
        cfg.set_setting("teams_webhook_url", "")

    def test_is_available(self):
        from gateway_client import GatewayClient
        cfg.set_setting("fusesign_api_key", "test-key")
        gw = GatewayClient()
        self.assertTrue(gw.is_available("fusesign"))
        self.assertFalse(gw.is_available("xpm"))
        cfg.set_setting("fusesign_api_key", "")

    def test_status_summary_returns_string(self):
        from gateway_client import GatewayClient
        gw = GatewayClient()
        summary = gw.status_summary()
        self.assertIsInstance(summary, str)
        self.assertIn("XPM", summary)
        self.assertIn("FuseSign", summary)
        self.assertIn("Teams", summary)

    def test_load_refreshes_credentials(self):
        from gateway_client import GatewayClient
        gw = GatewayClient()
        self.assertFalse(gw.xpm.is_configured)
        cfg.set_setting("xpm_api_key", "new-key")
        gw.load()
        self.assertTrue(gw.xpm.is_configured)
        cfg.set_setting("xpm_api_key", "")


# ── 2. XPMClient ──────────────────────────────────────────────────────────────

class TestXPMClient(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()
        cfg.set_setting("xpm_api_key", "test-xpm-key-123")
        from gateway_client import XPMClient
        self.xpm = XPMClient()

    def tearDown(self):
        cfg.set_setting("xpm_api_key", "")
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_is_configured_true_when_key_set(self):
        self.assertTrue(self.xpm.is_configured)

    def test_is_configured_false_when_no_key(self):
        from gateway_client import XPMClient
        # Clear the DB key so get_setting() also returns empty
        cfg.set_setting("xpm_api_key", "")
        xpm = XPMClient(api_key="")
        self.assertFalse(xpm.is_configured)

    def test_list_clients_returns_list(self):
        mock_data = {"items": [
            {"id": "1", "name": "John Smith", "email": "john@example.com"},
            {"id": "2", "name": "Jane Doe", "email": "jane@example.com"},
        ]}
        with mock_urlopen(mock_data):
            clients = self.xpm.list_clients()
        self.assertEqual(len(clients), 2)
        self.assertEqual(clients[0]["name"], "John Smith")

    def test_list_clients_with_search(self):
        mock_data = {"items": [
            {"id": "1", "name": "John Smith", "email": "john@example.com"},
        ]}
        with mock_urlopen(mock_data):
            clients = self.xpm.list_clients(search="Smith")
        self.assertEqual(len(clients), 1)

    def test_get_client_returns_dict(self):
        mock_data = {"id": "1", "name": "John Smith", "email": "john@example.com"}
        with mock_urlopen(mock_data):
            client = self.xpm.get_client("1")
        self.assertEqual(client["id"], "1")
        self.assertEqual(client["name"], "John Smith")

    def test_get_client_by_email_found(self):
        mock_data = {"items": [
            {"id": "1", "name": "John Smith", "email": "john@example.com"},
        ]}
        with mock_urlopen(mock_data):
            client = self.xpm.get_client_by_email("john@example.com")
        self.assertIsNotNone(client)
        self.assertEqual(client["id"], "1")

    def test_get_client_by_email_not_found(self):
        mock_data = {"items": []}
        with mock_urlopen(mock_data):
            client = self.xpm.get_client_by_email("nobody@example.com")
        self.assertIsNone(client)

    def test_list_jobs_returns_list(self):
        mock_data = {"items": [
            {"id": "j1", "name": "Tax Return FY2024", "status": "In Progress"},
        ]}
        with mock_urlopen(mock_data):
            jobs = self.xpm.list_jobs(client_id="1")
        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["status"], "In Progress")

    def test_update_job_status(self):
        mock_data = {"id": "j1", "status": "Completed"}
        with mock_urlopen(mock_data):
            result = self.xpm.update_job_status("j1", "Completed")
        self.assertEqual(result["status"], "Completed")

    def test_add_client_note(self):
        mock_data = {"id": "n1", "content": "Called client about refund"}
        with mock_urlopen(mock_data):
            result = self.xpm.add_client_note("1", "Called client about refund")
        self.assertIn("id", result)

    def test_raises_xpm_error_when_not_configured(self):
        from gateway_client import XPMClient, XPMError
        xpm = XPMClient(api_key="")
        with self.assertRaises(XPMError):
            xpm.list_clients()

    def test_raises_xpm_error_on_http_error(self):
        from gateway_client import XPMError
        import urllib.error
        with patch("urllib.request.urlopen",
                   side_effect=urllib.error.HTTPError(
                       url="", code=401, msg="Unauthorized", hdrs=None, fp=BytesIO(b"Unauthorized")
                   )):
            with self.assertRaises(XPMError):
                self.xpm.list_clients()

    def test_raises_xpm_error_on_network_failure(self):
        from gateway_client import XPMError
        with patch("urllib.request.urlopen",
                   side_effect=ConnectionError("Network unreachable")):
            with self.assertRaises(XPMError):
                self.xpm.list_clients()


# ── 3. FuseSignClient ─────────────────────────────────────────────────────────

class TestFuseSignClient(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()
        cfg.set_setting("fusesign_api_key", "test-fusesign-key-456")
        from gateway_client import FuseSignClient
        self.fs = FuseSignClient()

    def tearDown(self):
        cfg.set_setting("fusesign_api_key", "")
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_is_configured_true_when_key_set(self):
        self.assertTrue(self.fs.is_configured)

    def test_is_configured_false_when_no_key(self):
        from gateway_client import FuseSignClient
        # Clear the DB key so get_setting() also returns empty
        cfg.set_setting("fusesign_api_key", "")
        fs = FuseSignClient(api_key="")
        self.assertFalse(fs.is_configured)

    def test_create_envelope_returns_dict_with_id(self):
        mock_data = {
            "id": "env_abc123",
            "status": "pending",
            "signing_url": "https://sign.fusesign.com/env_abc123",
        }
        with mock_urlopen(mock_data):
            envelope = self.fs.create_envelope(
                template_id="tpl_tax_return",
                recipients=[{"name": "John Smith", "email": "john@example.com"}],
                subject="Tax Return FY2024 — Please Sign",
            )
        self.assertEqual(envelope["id"], "env_abc123")
        self.assertIn("signing_url", envelope)

    def test_get_envelope_status(self):
        mock_data = {
            "id": "env_abc123",
            "status": "completed",
            "completed_at": "2024-07-01T10:00:00Z",
        }
        with mock_urlopen(mock_data):
            status = self.fs.get_envelope_status("env_abc123")
        self.assertEqual(status["status"], "completed")

    def test_list_envelopes_returns_list(self):
        mock_data = {"items": [
            {"id": "env_1", "status": "pending"},
            {"id": "env_2", "status": "completed"},
        ]}
        with mock_urlopen(mock_data):
            envelopes = self.fs.list_envelopes()
        self.assertEqual(len(envelopes), 2)

    def test_list_envelopes_filtered_by_status(self):
        mock_data = {"items": [
            {"id": "env_1", "status": "pending"},
        ]}
        with mock_urlopen(mock_data):
            envelopes = self.fs.list_envelopes(status="pending")
        self.assertEqual(len(envelopes), 1)

    def test_void_envelope(self):
        mock_data = {"id": "env_abc123", "status": "voided"}
        with mock_urlopen(mock_data):
            result = self.fs.void_envelope("env_abc123", reason="Client requested")
        self.assertEqual(result["status"], "voided")

    def test_resend_envelope(self):
        mock_data = {"id": "env_abc123", "status": "sent"}
        with mock_urlopen(mock_data):
            result = self.fs.resend_envelope("env_abc123")
        self.assertIn("id", result)

    def test_list_templates(self):
        mock_data = {"items": [
            {"id": "tpl_1", "name": "Tax Return"},
            {"id": "tpl_2", "name": "Engagement Letter"},
        ]}
        with mock_urlopen(mock_data):
            templates = self.fs.list_templates()
        self.assertEqual(len(templates), 2)

    def test_raises_fusesign_error_when_not_configured(self):
        from gateway_client import FuseSignClient, FuseSignError
        fs = FuseSignClient(api_key="")
        with self.assertRaises(FuseSignError):
            fs.list_envelopes()

    def test_raises_fusesign_error_on_http_error(self):
        from gateway_client import FuseSignError
        import urllib.error
        with patch("urllib.request.urlopen",
                   side_effect=urllib.error.HTTPError(
                       url="", code=403, msg="Forbidden", hdrs=None, fp=BytesIO(b"Forbidden")
                   )):
            with self.assertRaises(FuseSignError):
                self.fs.list_envelopes()


# ── 4. TeamsClient ────────────────────────────────────────────────────────────

class TestTeamsClient(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()
        cfg.set_setting("teams_webhook_url",
                        "https://outlook.office.com/webhook/test-webhook-url")
        from gateway_client import TeamsClient
        self.teams = TeamsClient()

    def tearDown(self):
        cfg.set_setting("teams_webhook_url", "")
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_is_configured_true_when_webhook_set(self):
        self.assertTrue(self.teams.is_configured)

    def test_is_configured_false_when_no_credentials(self):
        from gateway_client import TeamsClient
        # Clear the DB keys so get_setting() also returns empty
        cfg.set_setting("teams_webhook_url", "")
        cfg.set_setting("teams_graph_token", "")
        teams = TeamsClient(webhook_url="", graph_token="")
        self.assertFalse(teams.is_configured)

    def test_webhook_configured_property(self):
        self.assertTrue(self.teams.webhook_configured)

    def test_send_message_returns_true_on_success(self):
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = self.teams.send_message("Test message from CoWorker")
        self.assertTrue(result)

    def test_send_message_with_title(self):
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = self.teams.send_message(
                text="NOA processed for John Smith",
                title="NOA Alert",
            )
        self.assertTrue(result)

    def test_send_alert_urgent(self):
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = self.teams.send_alert(
                title="Urgent: Client Action Required",
                body="John Smith has an outstanding balance",
                urgent=True,
            )
        self.assertTrue(result)

    def test_send_message_raises_when_not_configured(self):
        from gateway_client import TeamsClient, TeamsError
        # Clear DB keys so get_setting() also returns empty
        cfg.set_setting("teams_webhook_url", "")
        cfg.set_setting("teams_graph_token", "")
        teams = TeamsClient(webhook_url="", graph_token="")
        with self.assertRaises(TeamsError):
            teams.send_message("Test")

    def test_send_message_raises_on_http_error(self):
        from gateway_client import TeamsError
        import urllib.error
        with patch("urllib.request.urlopen",
                   side_effect=urllib.error.HTTPError(
                       url="", code=400, msg="Bad Request", hdrs=None, fp=BytesIO(b"Bad Request")
                   )):
            with self.assertRaises(TeamsError):
                self.teams.send_message("Test")

    def test_send_channel_message_via_graph(self):
        from gateway_client import TeamsClient
        cfg.set_setting("teams_graph_token", "test-graph-token")
        teams = TeamsClient()
        mock_data = {"id": "msg_123", "body": {"content": "Hello Teams"}}
        with mock_urlopen(mock_data):
            result = teams.send_channel_message(
                team_id="team_abc",
                channel_id="channel_xyz",
                content="Hello Teams",
            )
        self.assertEqual(result["id"], "msg_123")
        cfg.set_setting("teams_graph_token", "")

    def test_send_channel_message_raises_when_no_graph_token(self):
        from gateway_client import TeamsClient, TeamsError
        teams = TeamsClient(graph_token="")
        with self.assertRaises(TeamsError):
            teams.send_channel_message("team", "channel", "Hello")


# ── 5. PluginContext.gateway field ────────────────────────────────────────────

class TestPluginContextGatewayField(unittest.TestCase):

    def test_context_has_gateway_field(self):
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertTrue(hasattr(ctx, "gateway"),
                        "PluginContext must have a 'gateway' field")

    def test_context_gateway_defaults_to_none(self):
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertIsNone(ctx.gateway)

    def test_context_accepts_gateway_client(self):
        from plugin_base import PluginContext
        from gateway_client import GatewayClient
        # Ensure DB is initialised before GatewayClient reads settings
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()
        gw = GatewayClient()
        ctx = PluginContext(gateway=gw)
        self.assertIs(ctx.gateway, gw)


# ── 6. plugin_loader._make_context() injects GatewayClient ───────────────────

class TestPluginLoaderGatewayInjection(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_make_context_injects_gateway(self):
        from plugin_loader import PluginLoader
        from gateway_client import GatewayClient
        loader = PluginLoader()
        ctx = loader._make_context(draft_mode=True)
        self.assertIsNotNone(ctx.gateway,
                             "_make_context() should inject GatewayClient into context.gateway")
        self.assertIsInstance(ctx.gateway, GatewayClient)

    def test_make_context_gateway_has_correct_attributes(self):
        from plugin_loader import PluginLoader
        loader = PluginLoader()
        ctx = loader._make_context(draft_mode=True)
        if ctx.gateway is not None:
            self.assertTrue(hasattr(ctx.gateway, "xpm"))
            self.assertTrue(hasattr(ctx.gateway, "fusesign"))
            self.assertTrue(hasattr(ctx.gateway, "teams"))
            self.assertTrue(hasattr(ctx.gateway, "available_integrations"))


# ── 7. config.py — gateway credential defaults ───────────────────────────────

class TestConfigGatewayDefaults(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_xpm_api_key_default_empty(self):
        self.assertEqual(cfg.get_setting("xpm_api_key"), "")

    def test_xpm_base_url_default(self):
        self.assertEqual(cfg.get_setting("xpm_base_url"), "https://api.xpm.xero.com")

    def test_fusesign_api_key_default_empty(self):
        self.assertEqual(cfg.get_setting("fusesign_api_key"), "")

    def test_fusesign_base_url_default(self):
        self.assertEqual(cfg.get_setting("fusesign_base_url"), "https://api.fusesign.com/v1")

    def test_teams_webhook_url_default_empty(self):
        self.assertEqual(cfg.get_setting("teams_webhook_url"), "")

    def test_teams_graph_token_default_empty(self):
        self.assertEqual(cfg.get_setting("teams_graph_token"), "")

    def test_heartbeat_interval_default(self):
        self.assertEqual(cfg.get_setting("heartbeat_interval_seconds"), "60")


# ── 8. Graceful degradation ───────────────────────────────────────────────────

class TestGatewayGracefulDegradation(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_loader_graceful_when_gateway_client_missing(self):
        """If gateway_client import fails, _make_context() sets gateway=None."""
        from plugin_loader import PluginLoader
        loader = PluginLoader()
        with patch.dict("sys.modules", {"gateway_client": None}):
            # Force import failure
            import importlib
            with patch("builtins.__import__",
                       side_effect=lambda name, *a, **kw: (_ for _ in ()).throw(
                           ImportError("No module named 'gateway_client'")
                       ) if name == "gateway_client" else importlib.import_module(name)):
                ctx = loader._make_context(draft_mode=True)
                self.assertIsNone(ctx.gateway)


if __name__ == "__main__":
    unittest.main(verbosity=2)
