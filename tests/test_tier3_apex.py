"""
Tests for Tier 3 APEX capabilities:
  - event_wiring (multi-agent event wiring)
  - token_meter  (Claude usage tracking)
  - kpi_monitor  (proactive KPI monitoring)
  - auto_updater (GitHub release checker)
"""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch, call
from datetime import datetime, timedelta

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg
_TEST_DB = Path(tempfile.mktemp(suffix="_tier3_test.db"))
cfg.DB_PATH = _TEST_DB
cfg.init_db()


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def make_context(with_gateway=False, with_memory=False):
    ctx = MagicMock()
    ctx.graph = None
    ctx.claude = None
    ctx.claude_fast = None
    ctx.claude_reason = None
    ctx.memory = MagicMock() if with_memory else None
    ctx.event_bus = MagicMock()
    ctx.notify = MagicMock()
    ctx.log = MagicMock()
    ctx.draft_mode = True
    ctx.settings = {}
    if with_gateway:
        ctx.gateway = MagicMock()
        ctx.gateway.is_available.return_value = True
        ctx.gateway.xpm = MagicMock()
        ctx.gateway.fusesign = MagicMock()
        ctx.gateway.teams = MagicMock()
    else:
        ctx.gateway = None
    return ctx


# ─────────────────────────────────────────────────────────────────────────────
# event_wiring tests
# ─────────────────────────────────────────────────────────────────────────────

class TestEventWiring(unittest.TestCase):

    def test_events_constants_exist(self):
        from event_wiring import Events
        self.assertEqual(Events.EMAIL_TRIAGE_COMPLETE, "email.triage.complete")
        self.assertEqual(Events.NOA_PROCESSED, "noa.processed")
        self.assertEqual(Events.FUSESIGN_COMPLETED, "fusesign.completed")
        self.assertEqual(Events.HEARTBEAT_TICK, "heartbeat.tick")

    def test_wire_all_subscribes_handlers(self):
        from event_wiring import wire_all
        mock_bus = MagicMock()
        mock_bus.subscribe = MagicMock()
        plugins = {}
        wire_all(mock_bus, plugins)
        # Should have subscribed at least 4 handlers
        self.assertGreaterEqual(mock_bus.subscribe.call_count, 4)

    def test_patch_noa_processor_plugin(self):
        from event_wiring import patch_noa_processor_plugin
        mock_bus = MagicMock()
        plugin = MagicMock()
        result = MagicMock()
        result.actions_taken = 1
        plugin.run.return_value = result

        patch_noa_processor_plugin(plugin, mock_bus)
        ctx = make_context()
        plugin.run(ctx)

        mock_bus.publish.assert_called_once()
        args = mock_bus.publish.call_args
        self.assertIn("noa.processed", args[0])

    def test_apply_all_patches_handles_missing_plugins(self):
        from event_wiring import apply_all_patches
        mock_bus = MagicMock()
        # Should not raise even if plugins are missing
        apply_all_patches({}, mock_bus)

    def test_triage_router_routes_documents_received(self):
        """Test that DOCUMENTS_RECEIVED category triggers NOA processor notification."""
        from event_wiring import _wire_triage_to_processors
        mock_bus = MagicMock()
        noa_plugin = MagicMock()
        noa_plugin._pending_msg_ids = set()
        plugins = {"plugin_noa_processor": noa_plugin}

        _wire_triage_to_processors(mock_bus, plugins)

        # Get the handler that was subscribed
        subscribe_call = mock_bus.subscribe.call_args
        handler = subscribe_call[0][1]

        # Simulate a triage complete event
        handler({"payload": {"category": "DOCUMENTS_RECEIVED", "msg_id": "msg123"}})

        # NOA processor should have been notified
        self.assertIn("msg123", noa_plugin._pending_msg_ids)


# ─────────────────────────────────────────────────────────────────────────────
# token_meter tests
# ─────────────────────────────────────────────────────────────────────────────

class TestTokenMeter(unittest.TestCase):

    def setUp(self):
        self.db = Path(tempfile.mktemp(suffix="_token_test.db"))
        from token_meter import init_meter_db
        init_meter_db(self.db)

    def tearDown(self):
        if self.db.exists():
            self.db.unlink()

    def test_calculate_cost_haiku(self):
        from token_meter import calculate_cost_usd
        # 1000 input + 500 output at Haiku rates
        cost = calculate_cost_usd("claude-haiku-4-5-20251001", 1000, 500)
        expected = (1000 * 0.80 + 500 * 4.00) / 1_000_000
        self.assertAlmostEqual(cost, expected, places=8)

    def test_calculate_cost_sonnet(self):
        from token_meter import calculate_cost_usd
        cost = calculate_cost_usd("claude-sonnet-4-6", 2000, 1000)
        expected = (2000 * 3.00 + 1000 * 15.00) / 1_000_000
        self.assertAlmostEqual(cost, expected, places=8)

    def test_calculate_cost_unknown_model(self):
        from token_meter import calculate_cost_usd
        # Unknown model should use default pricing
        cost = calculate_cost_usd("claude-unknown-model", 1000, 500)
        self.assertGreater(cost, 0)

    def test_usd_to_aud(self):
        from token_meter import usd_to_aud
        aud = usd_to_aud(1.0)
        self.assertGreater(aud, 1.0)  # AUD should be more than USD

    def test_log_usage_and_retrieve(self):
        from token_meter import log_usage, get_usage_summary
        log_usage("plugin_test", "claude-haiku-4-5-20251001",
                  1000, 500, tier="fast",
                  prompt_summary="test prompt", db_path=self.db)
        summary = get_usage_summary(days=30, db_path=self.db)
        self.assertEqual(summary["total_calls"], 1)
        self.assertEqual(summary["total_tokens"], 1500)
        self.assertGreater(summary["total_cost_aud"], 0)

    def test_get_usage_summary_empty(self):
        from token_meter import get_usage_summary
        summary = get_usage_summary(days=30, db_path=self.db)
        self.assertEqual(summary["total_calls"], 0)
        self.assertEqual(summary["by_plugin"], [])

    def test_usage_summary_by_plugin(self):
        from token_meter import log_usage, get_usage_summary
        log_usage("plugin_a", "claude-haiku-4-5-20251001", 100, 50, db_path=self.db)
        log_usage("plugin_b", "claude-sonnet-4-6", 200, 100, db_path=self.db)
        log_usage("plugin_a", "claude-haiku-4-5-20251001", 150, 75, db_path=self.db)
        summary = get_usage_summary(days=30, db_path=self.db)
        plugin_ids = [p["plugin_id"] for p in summary["by_plugin"]]
        self.assertIn("plugin_a", plugin_ids)
        self.assertIn("plugin_b", plugin_ids)

    def test_claude_usage_wrapper_intercepts_calls(self):
        from token_meter import ClaudeUsageWrapper, init_meter_db, get_usage_summary
        init_meter_db(self.db)

        # Mock Anthropic client
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.usage.input_tokens = 100
        mock_response.usage.output_tokens = 50
        mock_client.messages.create.return_value = mock_response

        wrapper = ClaudeUsageWrapper(mock_client, plugin_id="test_plugin",
                                      tier="fast", db_path=self.db)
        result = wrapper.messages.create(
            model="claude-haiku-4-5-20251001",
            messages=[{"role": "user", "content": "Hello"}]
        )

        # Original call was made
        mock_client.messages.create.assert_called_once()
        # Result is the original response
        self.assertEqual(result, mock_response)
        # Usage was logged
        summary = get_usage_summary(days=1, db_path=self.db)
        self.assertEqual(summary["total_calls"], 1)
        self.assertEqual(summary["total_tokens"], 150)

    def test_claude_usage_wrapper_handles_api_error(self):
        from token_meter import ClaudeUsageWrapper, get_usage_summary
        mock_client = MagicMock()
        mock_client.messages.create.side_effect = Exception("API error")

        wrapper = ClaudeUsageWrapper(mock_client, plugin_id="test_plugin",
                                      tier="fast", db_path=self.db)
        with self.assertRaises(Exception):
            wrapper.messages.create(
                model="claude-haiku-4-5-20251001",
                messages=[{"role": "user", "content": "Hello"}]
            )
        # Failed call was still logged
        summary = get_usage_summary(days=1, db_path=self.db)
        self.assertEqual(summary["total_calls"], 1)

    def test_wrap_context_claude(self):
        from token_meter import ClaudeUsageWrapper, wrap_context_claude, init_meter_db
        init_meter_db(self.db)

        ctx = MagicMock()
        mock_client = MagicMock()
        ctx.claude = mock_client
        ctx.claude_fast = mock_client
        ctx.claude_reason = mock_client

        with patch("token_meter._get_meter_db_path", return_value=self.db):
            wrap_context_claude(ctx, "test_plugin")

        self.assertIsInstance(ctx.claude, ClaudeUsageWrapper)
        self.assertIsInstance(ctx.claude_fast, ClaudeUsageWrapper)
        self.assertIsInstance(ctx.claude_reason, ClaudeUsageWrapper)

    def test_wrap_context_claude_none_clients(self):
        from token_meter import wrap_context_claude, init_meter_db
        init_meter_db(self.db)

        ctx = MagicMock()
        ctx.claude = None
        ctx.claude_fast = None
        ctx.claude_reason = None

        with patch("token_meter._get_meter_db_path", return_value=self.db):
            wrap_context_claude(ctx, "test_plugin")

        # None clients should remain None
        self.assertIsNone(ctx.claude)
        self.assertIsNone(ctx.claude_fast)
        self.assertIsNone(ctx.claude_reason)


# ─────────────────────────────────────────────────────────────────────────────
# kpi_monitor tests
# ─────────────────────────────────────────────────────────────────────────────

class TestKPIMonitor(unittest.TestCase):

    def setUp(self):
        self.db = Path(tempfile.mktemp(suffix="_kpi_test.db"))
        from kpi_monitor import init_kpi_db, seed_default_kpis
        init_kpi_db(self.db)
        seed_default_kpis(self.db)

    def tearDown(self):
        if self.db.exists():
            self.db.unlink()

    def test_default_kpis_seeded(self):
        from kpi_monitor import get_kpi_config
        kpis = get_kpi_config(self.db)
        kpi_ids = [k["kpi_id"] for k in kpis]
        self.assertIn("wip_ageing", kpi_ids)
        self.assertIn("debtor_days", kpi_ids)
        self.assertIn("unsigned_docs", kpi_ids)
        self.assertIn("inbox_backlog", kpi_ids)
        self.assertIn("plugin_failures", kpi_ids)
        self.assertIn("ai_cost_spike", kpi_ids)

    def test_update_kpi_threshold(self):
        from kpi_monitor import update_kpi_threshold, get_kpi_config
        update_kpi_threshold("wip_ageing", 120.0, enabled=True, db_path=self.db)
        kpis = {k["kpi_id"]: k for k in get_kpi_config(self.db)}
        self.assertEqual(kpis["wip_ageing"]["threshold"], 120.0)

    def test_cooldown_prevents_repeated_alerts(self):
        from kpi_monitor import _was_recently_alerted, _record_alert
        _record_alert("wip_ageing", "warning", 5.0, 90.0, "test", self.db)
        self.assertTrue(_was_recently_alerted("wip_ageing", 4.0, self.db))

    def test_no_cooldown_for_different_kpi(self):
        from kpi_monitor import _was_recently_alerted, _record_alert
        _record_alert("wip_ageing", "warning", 5.0, 90.0, "test", self.db)
        self.assertFalse(_was_recently_alerted("debtor_days", 4.0, self.db))

    def test_ai_cost_spike_check_triggers(self):
        from kpi_monitor import _check_ai_cost_spike
        config = {"threshold": 1.00}  # $1 AUD threshold
        with patch("token_meter.get_usage_summary",
                   return_value={"today_cost_aud": 5.00}):
            result = _check_ai_cost_spike(make_context(), config)
        self.assertIsNotNone(result)
        self.assertEqual(result[0], 5.00)

    def test_ai_cost_spike_no_trigger_below_threshold(self):
        from kpi_monitor import _check_ai_cost_spike
        config = {"threshold": 10.00}
        with patch("token_meter.get_usage_summary",
                   return_value={"today_cost_aud": 0.50}):
            result = _check_ai_cost_spike(make_context(), config)
        self.assertIsNone(result)

    def test_plugin_failures_check_triggers(self):
        from kpi_monitor import _check_plugin_failures
        config = {"threshold": 2}
        recent_time = datetime.now().timestamp()
        mock_history = [
            {"timestamp": recent_time, "payload": {"plugin_id": "plugin_a"}},
            {"timestamp": recent_time, "payload": {"plugin_id": "plugin_b"}},
            {"timestamp": recent_time, "payload": {"plugin_id": "plugin_c"}},
        ]
        with patch("event_bus.EventBus") as mock_bus:
            mock_bus.get_history.return_value = mock_history
            result = _check_plugin_failures(make_context(), config)
        self.assertIsNotNone(result)
        self.assertEqual(result[0], 3.0)

    def test_wip_ageing_no_gateway(self):
        from kpi_monitor import _check_wip_ageing
        ctx = make_context(with_gateway=False)
        result = _check_wip_ageing(ctx, {"threshold": 90})
        self.assertIsNone(result)

    def test_run_checks_returns_list(self):
        from kpi_monitor import _KPIMonitor
        monitor = _KPIMonitor()
        monitor._db_path = self.db
        ctx = make_context()
        results = monitor.run_checks(ctx)
        self.assertIsInstance(results, list)

    def test_get_recent_alerts_empty(self):
        from kpi_monitor import get_recent_alerts
        alerts = get_recent_alerts(db_path=self.db)
        self.assertEqual(alerts, [])

    def test_dispatch_alert_calls_teams(self):
        from kpi_monitor import _dispatch_alert
        ctx = make_context(with_gateway=True)
        _dispatch_alert(ctx, "wip_ageing", "warning", 5.0, 90.0, "Test alert")
        ctx.gateway.teams.send_alert.assert_called_once()


# ─────────────────────────────────────────────────────────────────────────────
# auto_updater tests
# ─────────────────────────────────────────────────────────────────────────────

class TestAutoUpdater(unittest.TestCase):

    def setUp(self):
        self.db = Path(tempfile.mktemp(suffix="_updater_test.db"))
        from auto_updater import init_updater_db
        init_updater_db(self.db)

    def tearDown(self):
        if self.db.exists():
            self.db.unlink()

    def test_get_current_version(self):
        from auto_updater import get_current_version
        version = get_current_version()
        self.assertRegex(version, r"\d+\.\d+\.\d+")

    def test_parse_version(self):
        from auto_updater import parse_version
        self.assertEqual(parse_version("1.2.3"), (1, 2, 3))
        self.assertEqual(parse_version("v2.0.0"), (2, 0, 0))
        self.assertEqual(parse_version("1.0.0-beta"), (1, 0, 0))

    def test_is_newer(self):
        from auto_updater import is_newer
        self.assertTrue(is_newer("1.1.0", "1.0.0"))
        self.assertFalse(is_newer("1.0.0", "1.0.0"))
        self.assertFalse(is_newer("0.9.0", "1.0.0"))

    def test_check_for_update_network_error(self):
        from auto_updater import check_for_update
        from urllib.error import URLError
        with patch("auto_updater.urlopen", side_effect=URLError("no network")):
            result = check_for_update(self.db)
        self.assertIsNone(result)

    def test_check_for_update_up_to_date(self):
        from auto_updater import check_for_update
        import json
        mock_data = json.dumps({
            "tag_name": "v0.0.1",  # older than current 1.0.0
            "html_url": "https://github.com/test",
            "body": "Old release",
            "assets": []
        }).encode()
        mock_resp = MagicMock()
        mock_resp.read.return_value = mock_data
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        with patch("auto_updater.urlopen", return_value=mock_resp):
            result = check_for_update(self.db)
        self.assertIsNone(result)

    def test_check_for_update_new_version(self):
        from auto_updater import check_for_update
        import json
        mock_data = json.dumps({
            "tag_name": "v99.0.0",  # much newer
            "html_url": "https://github.com/test/releases/v99.0.0",
            "body": "Major update",
            "assets": [{"name": "mcs-coworker.zip",
                         "browser_download_url": "https://example.com/update.zip"}]
        }).encode()
        mock_resp = MagicMock()
        mock_resp.read.return_value = mock_data
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        with patch("auto_updater.urlopen", return_value=mock_resp):
            result = check_for_update(self.db)
        self.assertIsNotNone(result)
        self.assertEqual(result["latest_version"], "99.0.0")
        self.assertEqual(result["zip_url"], "https://example.com/update.zip")

    def test_get_update_history_empty(self):
        from auto_updater import get_update_history
        history = get_update_history(db_path=self.db)
        self.assertEqual(history, [])

    def test_set_auto_update(self):
        from auto_updater import _AutoUpdater, _get_config
        updater = _AutoUpdater()
        updater._db_path = self.db
        updater.set_auto_update(True)
        self.assertEqual(_get_config("auto_update", "0", self.db), "1")
        updater.set_auto_update(False)
        self.assertEqual(_get_config("auto_update", "0", self.db), "0")

    def test_get_status(self):
        from auto_updater import _AutoUpdater
        updater = _AutoUpdater()
        updater._db_path = self.db
        status = updater.get_status()
        self.assertIn("current_version", status)
        self.assertIn("auto_update", status)
        self.assertIn("history", status)

    def test_notify_calls_teams(self):
        from auto_updater import _AutoUpdater
        updater = _AutoUpdater()
        updater._db_path = self.db
        ctx = make_context(with_gateway=True)
        updater._context = ctx
        updater._notify({
            "current_version": "1.0.0",
            "latest_version": "2.0.0",
            "release_notes": "New features",
        })
        ctx.gateway.teams.send_alert.assert_called_once()


if __name__ == "__main__":
    unittest.main()
