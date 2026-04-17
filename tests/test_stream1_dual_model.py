"""
Stream 1 Test Suite — Two-Tier Claude Model System
====================================================
Tests every aspect of the dual-model implementation:
  1. config.py  — new model getters, update_claude_models(), backward compat
  2. plugin_base.py — PluginContext has claude_fast / claude_reason fields
  3. plugin_loader.py — set_claude() populates all three client slots
  4. Backward compat — existing plugins using context.claude still work
  5. Live API test — update_claude_models() actually hits Anthropic and
     returns real model IDs (requires ANTHROPIC_API_KEY env var or stored key)
"""

import sys
import os
import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path

# ── Path setup ────────────────────────────────────────────────────────────────
# Run from repo root: python -m pytest tests/test_stream1_dual_model.py -v
REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

# Use a temp DB so tests never touch the real user database
import tempfile
_tmp_dir = tempfile.mkdtemp()
os.environ["MCS_TEST_DB"] = os.path.join(_tmp_dir, "test_config.db")

# Monkey-patch DB_PATH before importing config
import config as cfg
cfg.DB_PATH = Path(os.environ["MCS_TEST_DB"])


# ── 1. config.py tests ────────────────────────────────────────────────────────

class TestConfigDualModel(unittest.TestCase):

    def setUp(self):
        """Fresh DB for every test."""
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_defaults_seeded(self):
        """init_db() seeds both model slots."""
        fast = cfg.get_setting("claude_model_fast")
        reasoning = cfg.get_setting("claude_model_reasoning")
        self.assertIn("haiku", fast.lower(),
                      f"claude_model_fast should be a Haiku model, got: {fast}")
        self.assertIn("sonnet", reasoning.lower(),
                      f"claude_model_reasoning should be a Sonnet model, got: {reasoning}")

    def test_get_claude_model_fast(self):
        """get_claude_model_fast() returns the fast model."""
        cfg.set_setting("claude_model_fast", "claude-haiku-test-1")
        self.assertEqual(cfg.get_claude_model_fast(), "claude-haiku-test-1")

    def test_get_claude_model_reasoning(self):
        """get_claude_model_reasoning() returns the reasoning model."""
        cfg.set_setting("claude_model_reasoning", "claude-sonnet-test-1")
        self.assertEqual(cfg.get_claude_model_reasoning(), "claude-sonnet-test-1")

    def test_get_claude_model_backward_compat(self):
        """get_claude_model() still works and returns the fast model."""
        cfg.set_setting("claude_model_fast", "claude-haiku-compat-test")
        self.assertEqual(cfg.get_claude_model(), "claude-haiku-compat-test")

    def test_update_claude_model_backward_compat(self):
        """Legacy update_claude_model() returns the fast model ID."""
        mock_response_data = {
            "data": [
                {"id": "claude-haiku-new", "type": "model", "created_at": "2025-01-02"},
                {"id": "claude-sonnet-new", "type": "model", "created_at": "2025-01-01"},
            ]
        }
        with patch("urllib.request.urlopen") as mock_urlopen:
            mock_cm = MagicMock()
            mock_cm.__enter__ = MagicMock(return_value=mock_cm)
            mock_cm.__exit__ = MagicMock(return_value=False)
            mock_cm.read.return_value = str(mock_response_data).replace("'", '"').encode()
            import json
            mock_cm.read.return_value = json.dumps(mock_response_data).encode()
            mock_urlopen.return_value = mock_cm
            result = cfg.update_claude_model("fake-api-key")
        self.assertEqual(result, "claude-haiku-new")

    def test_update_claude_models_returns_both(self):
        """update_claude_models() returns dict with 'fast' and 'reasoning' keys."""
        mock_response_data = {
            "data": [
                {"id": "claude-haiku-4-5-latest", "type": "model", "created_at": "2025-06-01"},
                {"id": "claude-haiku-4-5-old", "type": "model", "created_at": "2025-01-01"},
                {"id": "claude-sonnet-4-5-latest", "type": "model", "created_at": "2025-06-01"},
                {"id": "claude-sonnet-4-5-old", "type": "model", "created_at": "2025-01-01"},
                {"id": "claude-opus-3", "type": "model", "created_at": "2024-01-01"},
            ]
        }
        import json
        with patch("urllib.request.urlopen") as mock_urlopen:
            mock_cm = MagicMock()
            mock_cm.__enter__ = MagicMock(return_value=mock_cm)
            mock_cm.__exit__ = MagicMock(return_value=False)
            mock_cm.read.return_value = json.dumps(mock_response_data).encode()
            mock_urlopen.return_value = mock_cm
            result = cfg.update_claude_models("fake-api-key")

        self.assertIn("fast", result)
        self.assertIn("reasoning", result)
        self.assertEqual(result["fast"], "claude-haiku-4-5-latest")
        self.assertEqual(result["reasoning"], "claude-sonnet-4-5-latest")

        # Verify settings were persisted
        self.assertEqual(cfg.get_setting("claude_model_fast"), "claude-haiku-4-5-latest")
        self.assertEqual(cfg.get_setting("claude_model_reasoning"), "claude-sonnet-4-5-latest")
        # Legacy key also updated
        self.assertEqual(cfg.get_setting("claude_model"), "claude-haiku-4-5-latest")

    def test_update_claude_models_graceful_on_error(self):
        """update_claude_models() returns current values if API call fails."""
        cfg.set_setting("claude_model_fast", "claude-haiku-fallback")
        cfg.set_setting("claude_model_reasoning", "claude-sonnet-fallback")
        with patch("urllib.request.urlopen", side_effect=Exception("network error")):
            result = cfg.update_claude_models("fake-api-key")
        self.assertEqual(result["fast"], "claude-haiku-fallback")
        self.assertEqual(result["reasoning"], "claude-sonnet-fallback")


# ── 2. plugin_base.py tests ───────────────────────────────────────────────────

class TestPluginContextDualModel(unittest.TestCase):

    def test_context_has_claude_fast_field(self):
        """PluginContext has a claude_fast field."""
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertTrue(hasattr(ctx, "claude_fast"),
                        "PluginContext must have a 'claude_fast' field")

    def test_context_has_claude_reason_field(self):
        """PluginContext has a claude_reason field."""
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertTrue(hasattr(ctx, "claude_reason"),
                        "PluginContext must have a 'claude_reason' field")

    def test_context_legacy_claude_field_still_exists(self):
        """PluginContext still has the legacy 'claude' field for backward compat."""
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertTrue(hasattr(ctx, "claude"),
                        "PluginContext must still have legacy 'claude' field")

    def test_context_defaults_are_none(self):
        """All three client fields default to None."""
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertIsNone(ctx.claude)
        self.assertIsNone(ctx.claude_fast)
        self.assertIsNone(ctx.claude_reason)

    def test_context_accepts_separate_clients(self):
        """PluginContext can hold different objects for fast and reason clients."""
        from plugin_base import PluginContext
        fast_mock = MagicMock(name="fast_client")
        reason_mock = MagicMock(name="reason_client")
        ctx = PluginContext(claude=fast_mock, claude_fast=fast_mock, claude_reason=reason_mock)
        self.assertIs(ctx.claude_fast, fast_mock)
        self.assertIs(ctx.claude_reason, reason_mock)
        self.assertIs(ctx.claude, fast_mock)  # legacy alias

    def test_agent_plugin_has_model_helpers(self):
        """AgentPlugin subclass has get_claude_model_fast() and get_claude_model_reasoning()."""
        from plugin_base import AgentPlugin, PluginContext, PluginResult

        class DummyPlugin(AgentPlugin):
            name = "Dummy"
            def run(self, context: PluginContext) -> PluginResult:
                return PluginResult(success=True)

        p = DummyPlugin()
        self.assertTrue(hasattr(p, "get_claude_model_fast"),
                        "AgentPlugin must have get_claude_model_fast()")
        self.assertTrue(hasattr(p, "get_claude_model_reasoning"),
                        "AgentPlugin must have get_claude_model_reasoning()")
        self.assertTrue(hasattr(p, "get_claude_model"),
                        "AgentPlugin must still have legacy get_claude_model()")

    def test_agent_plugin_model_helpers_return_strings(self):
        """Model helper methods return non-empty strings."""
        from plugin_base import AgentPlugin, PluginContext, PluginResult

        # Use a temp DB
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

        class DummyPlugin(AgentPlugin):
            name = "Dummy"
            def run(self, context: PluginContext) -> PluginResult:
                return PluginResult(success=True)

        p = DummyPlugin()
        fast = p.get_claude_model_fast()
        reasoning = p.get_claude_model_reasoning()
        legacy = p.get_claude_model()

        self.assertIsInstance(fast, str)
        self.assertGreater(len(fast), 0)
        self.assertIsInstance(reasoning, str)
        self.assertGreater(len(reasoning), 0)
        # Legacy should equal fast
        self.assertEqual(legacy, fast)

        os.unlink(str(cfg.DB_PATH))


# ── 3. plugin_loader.py tests ─────────────────────────────────────────────────

class TestPluginLoaderDualModel(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_loader_has_three_client_slots(self):
        """PluginLoader.__init__ creates _claude, _claude_fast, _claude_reason."""
        from plugin_loader import PluginLoader
        loader = PluginLoader()
        self.assertTrue(hasattr(loader, "_claude"))
        self.assertTrue(hasattr(loader, "_claude_fast"))
        self.assertTrue(hasattr(loader, "_claude_reason"))

    def test_set_claude_without_api_key(self):
        """set_claude() with no API key sets all three slots to None."""
        from plugin_loader import PluginLoader
        cfg.set_setting("anthropic_api_key", "")
        loader = PluginLoader()
        loader.set_claude()
        self.assertIsNone(loader._claude)
        self.assertIsNone(loader._claude_fast)
        self.assertIsNone(loader._claude_reason)

    def test_set_claude_with_api_key(self):
        """set_claude() with an API key populates all three slots."""
        from plugin_loader import PluginLoader
        cfg.set_setting("anthropic_api_key", "sk-ant-test-key-12345")

        mock_client = MagicMock(name="anthropic_client")
        with patch("anthropic.Anthropic", return_value=mock_client):
            loader = PluginLoader()
            loader.set_claude()

        self.assertIsNotNone(loader._claude)
        self.assertIsNotNone(loader._claude_fast)
        self.assertIsNotNone(loader._claude_reason)
        # All three should be the same client object
        self.assertIs(loader._claude, loader._claude_fast)
        self.assertIs(loader._claude, loader._claude_reason)

    def test_make_context_includes_all_claude_fields(self):
        """_make_context() returns a PluginContext with all three claude fields."""
        from plugin_loader import PluginLoader
        from plugin_base import PluginContext
        cfg.set_setting("anthropic_api_key", "sk-ant-test-key-12345")

        mock_client = MagicMock(name="anthropic_client")
        with patch("anthropic.Anthropic", return_value=mock_client):
            loader = PluginLoader()
            loader.set_claude()
            ctx = loader._make_context(draft_mode=True)

        self.assertIsInstance(ctx, PluginContext)
        self.assertIsNotNone(ctx.claude)
        self.assertIsNotNone(ctx.claude_fast)
        self.assertIsNotNone(ctx.claude_reason)
        # Legacy alias matches fast
        self.assertIs(ctx.claude, ctx.claude_fast)

    def test_make_context_claude_none_when_no_key(self):
        """_make_context() returns None claude fields when no API key."""
        from plugin_loader import PluginLoader
        cfg.set_setting("anthropic_api_key", "")
        loader = PluginLoader()
        loader.set_claude()
        ctx = loader._make_context(draft_mode=True)
        self.assertIsNone(ctx.claude)
        self.assertIsNone(ctx.claude_fast)
        self.assertIsNone(ctx.claude_reason)


# ── 4. Backward compatibility test ───────────────────────────────────────────

class TestBackwardCompatibility(unittest.TestCase):
    """Ensure existing plugins that use context.claude still work unchanged."""

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_existing_plugin_using_context_claude(self):
        """A plugin that uses context.claude.messages.create() still works."""
        from plugin_base import AgentPlugin, PluginContext, PluginResult

        class LegacyPlugin(AgentPlugin):
            name = "Legacy"
            requires_claude = True

            def run(self, context: PluginContext) -> PluginResult:
                # Simulate old-style usage
                if not context.claude:
                    return PluginResult(success=False, error="No claude")
                response = context.claude.messages.create(
                    model=self.get_claude_model(),
                    max_tokens=10,
                    messages=[{"role": "user", "content": "test"}]
                )
                return PluginResult(success=True, summary="legacy worked")

        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text="ok")]
        mock_client.messages.create.return_value = mock_response

        ctx = PluginContext(
            claude=mock_client,
            claude_fast=mock_client,
            claude_reason=mock_client,
        )
        plugin = LegacyPlugin()
        result = plugin.run(ctx)

        self.assertTrue(result.success)
        self.assertEqual(result.summary, "legacy worked")
        mock_client.messages.create.assert_called_once()

    def test_new_plugin_using_context_claude_reason(self):
        """A new plugin that uses context.claude_reason works correctly."""
        from plugin_base import AgentPlugin, PluginContext, PluginResult

        class NewPlugin(AgentPlugin):
            name = "New"
            requires_claude = True

            def run(self, context: PluginContext) -> PluginResult:
                if not context.claude_reason:
                    return PluginResult(success=False, error="No reasoning client")
                response = context.claude_reason.messages.create(
                    model=self.get_claude_model_reasoning(),
                    max_tokens=100,
                    messages=[{"role": "user", "content": "analyse this"}]
                )
                return PluginResult(success=True, summary="reasoning worked")

        fast_mock = MagicMock(name="fast")
        reason_mock = MagicMock(name="reason")
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text="analysis result")]
        reason_mock.messages.create.return_value = mock_response

        ctx = PluginContext(
            claude=fast_mock,
            claude_fast=fast_mock,
            claude_reason=reason_mock,
        )
        plugin = NewPlugin()
        result = plugin.run(ctx)

        self.assertTrue(result.success)
        self.assertEqual(result.summary, "reasoning worked")
        reason_mock.messages.create.assert_called_once()
        fast_mock.messages.create.assert_not_called()


# ── 5. Live API test (optional — skipped if no key available) ─────────────────

class TestLiveAPI(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_live_update_claude_models(self):
        """Live test: update_claude_models() returns real model IDs from Anthropic."""
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            self.skipTest("ANTHROPIC_API_KEY not set — skipping live API test")

        result = cfg.update_claude_models(api_key)
        self.assertIn("fast", result)
        self.assertIn("reasoning", result)
        self.assertIn("haiku", result["fast"].lower(),
                      f"Fast model should be Haiku, got: {result['fast']}")
        self.assertIn("sonnet", result["reasoning"].lower(),
                      f"Reasoning model should be Sonnet, got: {result['reasoning']}")
        print(f"\n  Live models → fast: {result['fast']} | reasoning: {result['reasoning']}")

    def test_live_fast_model_inference(self):
        """Live test: fast model can actually respond to a simple message."""
        import anthropic as ant
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            self.skipTest("ANTHROPIC_API_KEY not set — skipping live inference test")

        cfg.set_setting("anthropic_api_key", api_key)
        models = cfg.update_claude_models(api_key)
        fast_model = models["fast"]

        client = ant.Anthropic(api_key=api_key)
        response = client.messages.create(
            model=fast_model,
            max_tokens=20,
            messages=[{"role": "user", "content": "Say 'ok' and nothing else."}]
        )
        text = response.content[0].text.strip().lower()
        self.assertIn("ok", text, f"Expected 'ok' in response, got: {text}")
        print(f"\n  Fast model ({fast_model}) responded: {text}")

    def test_live_reasoning_model_inference(self):
        """Live test: reasoning model can actually respond to a simple message."""
        import anthropic as ant
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            self.skipTest("ANTHROPIC_API_KEY not set — skipping live inference test")

        cfg.set_setting("anthropic_api_key", api_key)
        models = cfg.update_claude_models(api_key)
        reasoning_model = models["reasoning"]

        client = ant.Anthropic(api_key=api_key)
        response = client.messages.create(
            model=reasoning_model,
            max_tokens=20,
            messages=[{"role": "user", "content": "Say 'ok' and nothing else."}]
        )
        text = response.content[0].text.strip().lower()
        self.assertIn("ok", text, f"Expected 'ok' in response, got: {text}")
        print(f"\n  Reasoning model ({reasoning_model}) responded: {text}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
