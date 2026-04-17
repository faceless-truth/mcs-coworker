"""
Stream 3 Test Suite — Event Bus & Heartbeat Plugin
====================================================
Tests every aspect of the EventBus and HeartbeatPlugin implementation:
  1. EventBus core: subscribe, publish, unsubscribe
  2. Wildcard subscriptions
  3. Async dispatch
  4. Error isolation (bad handler doesn't break others)
  5. Event history
  6. HeartbeatPlugin: start, tick, stop, interval change
  7. plugin_base.py — PluginContext.event_bus field
  8. plugin_loader.py — EventBus integration:
       - subscribe to heartbeat.tick in __init__
       - publish plugin.run.complete on run
       - publish plugin.run.failed on exception
       - start/stop heartbeat with scheduler
  9. Plugin-to-plugin communication via EventBus
"""

import sys
import os
import time
import threading
import unittest
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

# ── Path setup ────────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

import config as cfg
cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
cfg.init_db()


# ── Helper: fresh EventBus for each test ─────────────────────────────────────

def fresh_bus():
    """Return a new _EventBus instance (not the module singleton)."""
    from event_bus import _EventBus
    return _EventBus()


# ── 1. Core subscribe / publish / unsubscribe ─────────────────────────────────

class TestEventBusCore(unittest.TestCase):

    def setUp(self):
        self.bus = fresh_bus()

    def test_subscribe_and_publish(self):
        received = []
        self.bus.subscribe("test.event", lambda e: received.append(e))
        self.bus.publish(__import__("event_bus").Event("test.event", {"val": 42}))
        self.assertEqual(len(received), 1)
        self.assertEqual(received[0].payload["val"], 42)

    def test_emit_convenience(self):
        received = []
        self.bus.subscribe("test.emit", lambda e: received.append(e))
        self.bus.emit("test.emit", {"key": "value"})
        self.assertEqual(len(received), 1)
        self.assertEqual(received[0].payload["key"], "value")

    def test_publish_returns_handler_count(self):
        self.bus.subscribe("count.test", lambda e: None)
        self.bus.subscribe("count.test", lambda e: None)
        from event_bus import Event
        count = self.bus.publish(Event("count.test"))
        self.assertEqual(count, 2)

    def test_unsubscribe_removes_handler(self):
        received = []
        handler = lambda e: received.append(e)
        self.bus.subscribe("unsub.test", handler)
        self.bus.unsubscribe("unsub.test", handler)
        self.bus.emit("unsub.test")
        self.assertEqual(len(received), 0)

    def test_unsubscribe_all(self):
        received = []
        handler = lambda e: received.append(e)
        self.bus.subscribe("type.a", handler)
        self.bus.subscribe("type.b", handler)
        self.bus.unsubscribe_all(handler)
        self.bus.emit("type.a")
        self.bus.emit("type.b")
        self.assertEqual(len(received), 0)

    def test_no_handlers_returns_zero(self):
        from event_bus import Event
        count = self.bus.publish(Event("no.handlers"))
        self.assertEqual(count, 0)

    def test_duplicate_subscribe_ignored(self):
        """Subscribing the same handler twice should only call it once."""
        received = []
        handler = lambda e: received.append(e)
        self.bus.subscribe("dup.test", handler)
        self.bus.subscribe("dup.test", handler)
        self.bus.emit("dup.test")
        self.assertEqual(len(received), 1)

    def test_multiple_handlers_all_called(self):
        results = []
        self.bus.subscribe("multi.test", lambda e: results.append("a"))
        self.bus.subscribe("multi.test", lambda e: results.append("b"))
        self.bus.subscribe("multi.test", lambda e: results.append("c"))
        self.bus.emit("multi.test")
        self.assertIn("a", results)
        self.assertIn("b", results)
        self.assertIn("c", results)

    def test_event_type_stored_correctly(self):
        received = []
        self.bus.subscribe("type.check", lambda e: received.append(e.type))
        self.bus.emit("type.check")
        self.assertEqual(received[0], "type.check")

    def test_event_source_stored_correctly(self):
        received = []
        self.bus.subscribe("src.check", lambda e: received.append(e.source))
        self.bus.emit("src.check", source="TestSource")
        self.assertEqual(received[0], "TestSource")

    def test_event_timestamp_is_float(self):
        received = []
        self.bus.subscribe("ts.check", lambda e: received.append(e.timestamp))
        self.bus.emit("ts.check")
        self.assertIsInstance(received[0], float)
        self.assertGreater(received[0], 0)


# ── 2. Wildcard subscriptions ─────────────────────────────────────────────────

class TestWildcardSubscriptions(unittest.TestCase):

    def setUp(self):
        self.bus = fresh_bus()

    def test_wildcard_receives_all_events(self):
        received_types = []
        self.bus.subscribe("*", lambda e: received_types.append(e.type))
        self.bus.emit("event.one")
        self.bus.emit("event.two")
        self.bus.emit("event.three")
        self.assertIn("event.one", received_types)
        self.assertIn("event.two", received_types)
        self.assertIn("event.three", received_types)

    def test_wildcard_and_specific_both_called(self):
        results = []
        self.bus.subscribe("*", lambda e: results.append("wildcard"))
        self.bus.subscribe("specific.event", lambda e: results.append("specific"))
        self.bus.emit("specific.event")
        self.assertIn("wildcard", results)
        self.assertIn("specific", results)


# ── 3. Async dispatch ─────────────────────────────────────────────────────────

class TestAsyncDispatch(unittest.TestCase):

    def setUp(self):
        self.bus = fresh_bus()

    def test_async_handler_called_eventually(self):
        received = []
        done = threading.Event()

        def async_handler(e):
            received.append(e)
            done.set()

        self.bus.subscribe("async.test", async_handler, async_dispatch=True)
        self.bus.emit("async.test", {"data": "async"})
        # Wait up to 2 seconds for the async handler to complete
        done.wait(timeout=2.0)
        self.assertEqual(len(received), 1)
        self.assertEqual(received[0].payload["data"], "async")

    def test_sync_handler_blocks_publisher(self):
        """Sync handler should complete before publish() returns."""
        completed = []
        self.bus.subscribe("sync.test", lambda e: completed.append(True))
        self.bus.emit("sync.test")
        self.assertEqual(len(completed), 1)


# ── 4. Error isolation ────────────────────────────────────────────────────────

class TestErrorIsolation(unittest.TestCase):

    def setUp(self):
        self.bus = fresh_bus()

    def test_bad_handler_doesnt_break_others(self):
        results = []

        def bad_handler(e):
            raise ValueError("I am a bad handler")

        def good_handler(e):
            results.append("good")

        self.bus.subscribe("error.test", bad_handler)
        self.bus.subscribe("error.test", good_handler)
        # Should not raise
        try:
            self.bus.emit("error.test")
        except Exception:
            self.fail("EventBus.emit() raised an exception from a handler")
        self.assertIn("good", results)


# ── 5. Event history ──────────────────────────────────────────────────────────

class TestEventHistory(unittest.TestCase):

    def setUp(self):
        self.bus = fresh_bus()

    def test_history_records_events(self):
        self.bus.emit("hist.a")
        self.bus.emit("hist.b")
        history = self.bus.get_history()
        types = [e.type for e in history]
        self.assertIn("hist.a", types)
        self.assertIn("hist.b", types)

    def test_history_filtered_by_type(self):
        self.bus.emit("filter.yes")
        self.bus.emit("filter.no")
        history = self.bus.get_history(event_type="filter.yes")
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0].type, "filter.yes")

    def test_history_limit(self):
        for i in range(10):
            self.bus.emit(f"limit.test.{i}")
        history = self.bus.get_history(limit=3)
        self.assertLessEqual(len(history), 3)

    def test_clear_history(self):
        self.bus.emit("clear.test")
        self.bus.clear_history()
        history = self.bus.get_history()
        self.assertEqual(len(history), 0)

    def test_subscriber_count(self):
        self.bus.subscribe("count.a", lambda e: None)
        self.bus.subscribe("count.a", lambda e: None)
        self.bus.subscribe("count.b", lambda e: None)
        self.assertEqual(self.bus.subscriber_count("count.a"), 2)
        self.assertEqual(self.bus.subscriber_count("count.b"), 1)
        self.assertEqual(self.bus.subscriber_count(), 3)


# ── 6. HeartbeatPlugin ────────────────────────────────────────────────────────

class TestHeartbeatPlugin(unittest.TestCase):

    def setUp(self):
        # Always stop the heartbeat before each test
        from event_bus import HeartbeatPlugin
        HeartbeatPlugin.stop()
        time.sleep(0.1)

    def tearDown(self):
        from event_bus import HeartbeatPlugin
        HeartbeatPlugin.stop()
        time.sleep(0.1)

    def test_heartbeat_starts_and_stops(self):
        from event_bus import HeartbeatPlugin
        HeartbeatPlugin.start(interval_seconds=5)
        self.assertTrue(HeartbeatPlugin.is_running())
        HeartbeatPlugin.stop()
        time.sleep(0.1)
        self.assertFalse(HeartbeatPlugin.is_running())

    def test_heartbeat_fires_tick_event(self):
        from event_bus import EventBus, HeartbeatPlugin
        ticks = []
        done = threading.Event()

        def on_tick(e):
            ticks.append(e)
            done.set()

        # Subscribe on the module-level EventBus BEFORE starting
        EventBus.subscribe("heartbeat.tick", on_tick)
        try:
            HeartbeatPlugin.start(interval_seconds=1)  # 1s for testing
            done.wait(timeout=4.0)  # wait up to 4s for first tick
        finally:
            HeartbeatPlugin.stop()
            EventBus.unsubscribe("heartbeat.tick", on_tick)

        self.assertGreater(len(ticks), 0,
                           "HeartbeatPlugin should have fired at least one tick")
        tick = ticks[0]
        self.assertEqual(tick.type, "heartbeat.tick")
        self.assertIn("tick_count", tick.payload)
        self.assertIn("interval_seconds", tick.payload)
        self.assertEqual(tick.source, "HeartbeatPlugin")

    def test_heartbeat_tick_count_increments(self):
        from event_bus import EventBus, HeartbeatPlugin
        tick_counts = []
        done = threading.Event()

        def on_tick(e):
            tick_counts.append(e.payload["tick_count"])
            if len(tick_counts) >= 2:
                done.set()

        EventBus.subscribe("heartbeat.tick", on_tick)
        HeartbeatPlugin.start(interval_seconds=1)
        done.wait(timeout=5.0)
        HeartbeatPlugin.stop()
        EventBus.unsubscribe("heartbeat.tick", on_tick)

        if len(tick_counts) >= 2:
            self.assertGreater(tick_counts[-1], tick_counts[0])

    def test_heartbeat_safe_to_start_twice(self):
        from event_bus import HeartbeatPlugin
        HeartbeatPlugin.start(interval_seconds=60)
        HeartbeatPlugin.start(interval_seconds=60)  # should not raise or create second thread
        self.assertTrue(HeartbeatPlugin.is_running())
        HeartbeatPlugin.stop()


# ── 7. PluginContext.event_bus field ─────────────────────────────────────────

class TestPluginContextEventBus(unittest.TestCase):

    def test_context_has_event_bus_field(self):
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertTrue(hasattr(ctx, "event_bus"),
                        "PluginContext must have an 'event_bus' field")

    def test_context_event_bus_defaults_to_none(self):
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertIsNone(ctx.event_bus)

    def test_context_accepts_event_bus(self):
        from plugin_base import PluginContext
        from event_bus import EventBus
        ctx = PluginContext(event_bus=EventBus)
        self.assertIs(ctx.event_bus, EventBus)


# ── 8. plugin_loader.py EventBus integration ─────────────────────────────────

class TestPluginLoaderEventBusIntegration(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_make_context_injects_event_bus(self):
        from plugin_loader import PluginLoader
        from event_bus import EventBus
        loader = PluginLoader()
        ctx = loader._make_context(draft_mode=True)
        self.assertIsNotNone(ctx.event_bus)
        self.assertIs(ctx.event_bus, EventBus)

    def test_loader_subscribes_to_heartbeat_on_init(self):
        from plugin_loader import PluginLoader
        from event_bus import EventBus
        initial_count = EventBus.subscriber_count("heartbeat.tick")
        loader = PluginLoader()
        new_count = EventBus.subscriber_count("heartbeat.tick")
        self.assertGreater(new_count, initial_count,
                           "PluginLoader should subscribe to heartbeat.tick in __init__")

    def test_run_plugin_publishes_complete_event(self):
        from plugin_loader import PluginLoader
        from event_bus import EventBus
        from plugin_base import AgentPlugin, PluginContext, PluginResult

        # Create a minimal plugin that always succeeds
        class DummyPlugin(AgentPlugin):
            name = "Dummy"
            def run(self, context: PluginContext) -> PluginResult:
                return PluginResult(success=True, summary="dummy ran")

        complete_events = []
        EventBus.subscribe("plugin.run.complete",
                           lambda e: complete_events.append(e))

        loader = PluginLoader()
        # Manually inject the plugin
        from plugin_loader import LoadedPlugin
        lp = LoadedPlugin.__new__(LoadedPlugin)
        lp.plugin_id = "plugin_dummy"
        lp.plugin_cls = DummyPlugin
        lp.instance = DummyPlugin()
        lp.is_ready = True
        lp.enabled = True
        lp.draft_mode = True
        lp.schedule_seconds = 0
        lp.display_name = None
        lp.last_run = None
        lp.last_result = "—"
        lp.last_summary = ""
        lp._next_run_at = 0.0
        loader._plugins["plugin_dummy"] = lp

        loader.run_plugin("plugin_dummy", manual=True)

        EventBus.unsubscribe("plugin.run.complete",
                             complete_events.append)  # cleanup

        self.assertGreater(len(complete_events), 0,
                           "run_plugin() should publish plugin.run.complete")
        evt = complete_events[0]
        self.assertEqual(evt.payload["plugin_id"], "plugin_dummy")
        self.assertTrue(evt.payload["success"])

    def test_run_plugin_publishes_failed_event_on_exception(self):
        from plugin_loader import PluginLoader
        from event_bus import EventBus
        from plugin_base import AgentPlugin, PluginContext, PluginResult

        class CrashPlugin(AgentPlugin):
            name = "Crash"
            def run(self, context: PluginContext) -> PluginResult:
                raise RuntimeError("intentional crash")

        failed_events = []
        EventBus.subscribe("plugin.run.failed",
                           lambda e: failed_events.append(e))

        loader = PluginLoader()
        from plugin_loader import LoadedPlugin
        lp = LoadedPlugin.__new__(LoadedPlugin)
        lp.plugin_id = "plugin_crash"
        lp.plugin_cls = CrashPlugin
        lp.instance = CrashPlugin()
        lp.is_ready = True
        lp.enabled = True
        lp.draft_mode = True
        lp.schedule_seconds = 0
        lp.display_name = None
        lp.last_run = None
        lp.last_result = "—"
        lp.last_summary = ""
        lp._next_run_at = 0.0
        loader._plugins["plugin_crash"] = lp

        loader.run_plugin("plugin_crash", manual=True)

        EventBus.unsubscribe("plugin.run.failed",
                             failed_events.append)

        self.assertGreater(len(failed_events), 0,
                           "run_plugin() should publish plugin.run.failed on exception")
        evt = failed_events[0]
        self.assertEqual(evt.payload["plugin_id"], "plugin_crash")
        self.assertIn("intentional crash", evt.payload["error"])

    def test_start_scheduler_starts_heartbeat(self):
        from plugin_loader import PluginLoader
        from event_bus import HeartbeatPlugin
        HeartbeatPlugin.stop()
        time.sleep(0.1)
        loader = PluginLoader()
        loader.start_scheduler()
        time.sleep(0.2)
        self.assertTrue(HeartbeatPlugin.is_running())
        loader.stop_scheduler()
        time.sleep(0.2)

    def test_stop_scheduler_stops_heartbeat(self):
        from plugin_loader import PluginLoader
        from event_bus import HeartbeatPlugin
        loader = PluginLoader()
        loader.start_scheduler()
        time.sleep(0.2)
        loader.stop_scheduler()
        time.sleep(0.2)
        self.assertFalse(HeartbeatPlugin.is_running())


# ── 9. Plugin-to-plugin communication ────────────────────────────────────────

class TestPluginToPluginCommunication(unittest.TestCase):
    """Verify that plugins can communicate via EventBus through context.event_bus."""

    def test_plugin_a_publishes_plugin_b_receives(self):
        from plugin_base import AgentPlugin, PluginContext, PluginResult
        from event_bus import EventBus, Event

        received_by_b = []

        class PluginA(AgentPlugin):
            name = "PluginA"
            def run(self, context: PluginContext) -> PluginResult:
                context.event_bus.emit(
                    "noa.processed",
                    {"client_email": "test@example.com", "outcome": "REFUND"},
                    source="PluginA",
                )
                return PluginResult(success=True)

        class PluginB(AgentPlugin):
            name = "PluginB"
            def load(self, context: PluginContext) -> bool:
                context.event_bus.subscribe("noa.processed", self._on_noa)
                return True
            def _on_noa(self, event: Event):
                received_by_b.append(event)
            def run(self, context: PluginContext) -> PluginResult:
                return PluginResult(success=True)

        # Simulate context with real EventBus
        ctx = PluginContext(event_bus=EventBus)

        plugin_b = PluginB()
        plugin_b.load(ctx)

        plugin_a = PluginA()
        plugin_a.run(ctx)

        self.assertEqual(len(received_by_b), 1)
        self.assertEqual(received_by_b[0].payload["outcome"], "REFUND")
        self.assertEqual(received_by_b[0].source, "PluginA")

        # Cleanup
        EventBus.unsubscribe("noa.processed", plugin_b._on_noa)


if __name__ == "__main__":
    unittest.main(verbosity=2)
