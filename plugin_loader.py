"""
MC & S Desktop Agent — Plugin Loader & Scheduler
==================================================
Discovers all plugins in the plugins/ folder, manages their lifecycle,
and runs them on schedule in background threads.
"""

import importlib
import importlib.util
import inspect
import calendar
import os
import sys
import threading
import time
from datetime import datetime, date
from pathlib import Path
from typing import Callable

try:
    import anthropic
except ImportError:
    anthropic = None

try:
    import pytz
except ImportError:
    pytz = None

from plugin_base import AgentPlugin, PluginContext, PluginResult
from config import (
    get_setting, get_plugin_state, save_plugin_state, get_all_plugin_states
)
from event_bus import EventBus, HeartbeatPlugin
from event_wiring import wire_all, apply_all_patches, Events

if getattr(sys, 'frozen', False):
    # Running as PyInstaller bundle — plugins sit next to the .exe
    _base = os.path.dirname(sys.executable)
    _internal = os.path.join(_base, '_internal')
    PLUGINS_DIR = Path(os.path.join(_base, 'plugins'))
    # Add paths so plugin imports (plugin_base, config, graph_client) resolve
    for _p in [_base, _internal, str(PLUGINS_DIR)]:
        if _p not in sys.path:
            sys.path.insert(0, _p)
else:
    # Running from source
    PLUGINS_DIR = Path(__file__).parent / "plugins"

# Plugin IDs that are shown as templates / not run automatically
# Only plugin_email_triage is active — everything else is hidden
TEMPLATE_PLUGIN_IDS = {
    "plugin_template",
    "plugin_noa_workflow",
    "plugin_fusesign_monitor",
    "plugin_meeting_prep",
    "plugin_monthly_invoicing",
    "plugin_debtor_followup",
    "plugin_asic_reminder",
    "plugin_client_checkin",
    "plugin_client_outreach",
}


class LoadedPlugin:
    """Wraps a plugin instance with runtime state."""

    def __init__(self, plugin_cls: type, plugin_id: str):
        self.plugin_id = plugin_id
        self.plugin_cls = plugin_cls
        self.instance = plugin_cls()
        self.is_ready = False

        # Load persisted state from DB (or defaults)
        state = get_plugin_state(plugin_id)
        self.enabled = bool(state.get("enabled", 1))
        self.draft_mode = bool(state.get("draft_mode", 1))
        self.display_name = state.get("display_name") or None

        # Use DB schedule if set, otherwise fall back to plugin's default
        db_sched = state.get("schedule_seconds", 0)
        if db_sched and db_sched > 0:
            self.schedule_seconds = db_sched
        else:
            self.schedule_seconds = self.instance.default_schedule.interval_seconds

        self.last_run: datetime | None = None
        self.last_result: str = "—"
        self.last_summary: str = ""
        self._next_run_at: float = 0.0  # unix timestamp

    @property
    def name(self) -> str:
        return self.display_name or self.instance.name

    @property
    def description(self) -> str:
        return self.instance.description

    @property
    def detail(self) -> str:
        return self.instance.detail

    @property
    def icon(self) -> str:
        return self.instance.icon

    @property
    def version(self) -> str:
        return self.instance.version

    @property
    def is_template(self) -> bool:
        return self.plugin_id in TEMPLATE_PLUGIN_IDS

    @property
    def schedule_label(self) -> str:
        sched = self.instance.default_schedule
        if sched.is_calendar_based():
            return sched.label
        if self.schedule_seconds <= 0:
            return "Manual only"
        if self.schedule_seconds < 3600:
            mins = self.schedule_seconds // 60
            return f"Every {mins} min" if mins > 1 else "Every 1 min"
        hours = self.schedule_seconds // 3600
        return f"Every {hours} hr" if hours > 1 else "Every 1 hr"

    def persist(self):
        save_plugin_state(
            self.plugin_id,
            enabled=int(self.enabled),
            draft_mode=int(self.draft_mode),
            schedule_seconds=self.schedule_seconds,
            last_run=self.last_run.isoformat() if self.last_run else None,
            last_result=self.last_result,
            last_summary=self.last_summary,
        )

    def _next_calendar_run(self) -> float:
        """Return the unix timestamp of the next calendar-based run."""
        sched = self.instance.default_schedule
        now = datetime.now()
        day = sched.day_of_month
        months_step = sched.months_interval

        # Try this month first, then advance by months_step until we find a future date
        year, month = now.year, now.month
        for _ in range(24):  # safety: max 2 years look-ahead
            # Clamp day to last valid day of month
            max_day = calendar.monthrange(year, month)[1]
            run_day = min(day, max_day)
            candidate = datetime(year, month, run_day, 8, 0, 0)  # run at 08:00
            if candidate > now:
                return candidate.timestamp()
            # Advance by months_interval
            month += months_step
            while month > 12:
                month -= 12
                year += 1
        return 0.0

    def schedule_next(self):
        sched = self.instance.default_schedule
        if sched.is_calendar_based():
            self._next_run_at = self._next_calendar_run()
        elif self.schedule_seconds > 0:
            self._next_run_at = time.time() + self.schedule_seconds
        else:
            self._next_run_at = 0.0

    def is_due(self) -> bool:
        sched = self.instance.default_schedule
        is_scheduled = sched.is_calendar_based() or self.schedule_seconds > 0
        return (
            self.enabled
            and self.is_ready
            and not self.is_template
            and is_scheduled
            and self._next_run_at > 0
            and time.time() >= self._next_run_at
        )


class PluginLoader:
    """
    Discovers, loads, and schedules all plugins.
    Provides the interface used by the main app UI.
    """

    def __init__(self, log_callback: Callable[[str], None] = print):
        self._log = log_callback
        self._plugins: dict[str, LoadedPlugin] = {}
        self._graph = None
        self._claude = None        # legacy — kept for backward compat, mirrors _claude_fast
        self._claude_fast = None   # Haiku: fast, cheap tasks
        self._claude_reason = None # Sonnet: complex reasoning & analysis
        self._scheduler_thread: threading.Thread | None = None
        self._running = False
        self._on_run_complete: Callable | None = None  # called after each plugin run
        self._on_plugin_registered: Callable | None = None  # called when new plugin discovered
        # Subscribe to heartbeat ticks so the scheduler wakes up on each tick
        EventBus.subscribe("heartbeat.tick", self._on_heartbeat_tick, async_dispatch=False)

    # ── Setup ─────────────────────────────────────────────────────────────────

    def set_graph(self, graph_client):
        self._graph = graph_client

    def set_claude(self):
        """Initialise both Claude clients from the stored API key.

        - _claude_fast   uses the fast (Haiku) model slot
        - _claude_reason uses the reasoning (Sonnet) model slot
        - _claude        is kept as a legacy alias pointing to the same client
          as _claude_fast so that existing plugins using context.claude continue
          to work without any changes.
        """
        if anthropic is None:
            self._claude = None
            self._claude_fast = None
            self._claude_reason = None
            return
        api_key = get_setting("anthropic_api_key")
        if api_key:
            # Both tiers share the same Anthropic client object — the model
            # is selected per-call via get_claude_model_fast() /
            # get_claude_model_reasoning() on the plugin instance.
            client = anthropic.Anthropic(api_key=api_key)
            self._claude_fast = client
            self._claude_reason = client
            self._claude = client  # backward compat alias
        else:
            self._claude = None
            self._claude_fast = None
            self._claude_reason = None

    def on_run_complete(self, callback: Callable):
        """Register a callback to be called after any plugin finishes running."""
        self._on_run_complete = callback

    def on_plugin_registered(self, callback: Callable):
        """Register a callback to be called when a new plugin is discovered."""
        self._on_plugin_registered = callback

    # ── Discovery ─────────────────────────────────────────────────────────────

    def discover(self) -> list[str]:
        """
        Scan the plugins/ directory, import all plugin_*.py files,
        and register any AgentPlugin subclasses found.
        Returns a list of newly discovered plugin IDs.
        """
        discovered = []
        if not PLUGINS_DIR.exists():
            return discovered

        for path in sorted(PLUGINS_DIR.glob("plugin_*.py")):
            module_name = path.stem
            plugin_id = module_name  # e.g. "plugin_email_triage"

            if plugin_id in self._plugins:
                continue  # already loaded

            try:
                spec = importlib.util.spec_from_file_location(module_name, path)
                module = importlib.util.module_from_spec(spec)
                sys.modules[module_name] = module
                spec.loader.exec_module(module)

                # Find the first concrete AgentPlugin subclass in the module
                for _, obj in inspect.getmembers(module, inspect.isclass):
                    if (
                        issubclass(obj, AgentPlugin)
                        and obj is not AgentPlugin
                        and not inspect.isabstract(obj)
                    ):
                        lp = LoadedPlugin(obj, plugin_id)
                        self._plugins[plugin_id] = lp
                        discovered.append(plugin_id)
                        self._log(f"  Loaded plugin: {lp.name} (v{lp.version})")
                        if self._on_plugin_registered:
                            self._on_plugin_registered(plugin_id, lp)
                        break

            except Exception as e:
                import traceback
                self._log(f"⚠ Failed to load plugin {module_name}: {e}")
                self._log(traceback.format_exc())

        return discovered

    def load_all(self):
        """Discover plugins and call .load() on each one."""
        self.discover()
        ctx = self._make_context(draft_mode=True)

        for pid, lp in self._plugins.items():
            if lp.is_template:
                lp.is_ready = False
                continue
            try:
                lp.is_ready = lp.instance.load(ctx)
                lp.schedule_next()
            except Exception as e:
                self._log(f"⚠ Plugin {lp.name} failed to load: {e}")
                lp.is_ready = False

    # ── Running ───────────────────────────────────────────────────────────────

    def run_plugin(self, plugin_id: str, manual: bool = False) -> PluginResult:
        """Run a specific plugin immediately."""
        lp = self._plugins.get(plugin_id)
        if not lp:
            return PluginResult(
                success=False, error=f"Plugin '{plugin_id}' not found."
            )

        if not lp.is_ready and not lp.is_template:
            return PluginResult(
                success=False,
                error=f"Plugin '{lp.name}' is not ready (check configuration).",
            )

        # Re-init claude in case API key was just set
        self.set_claude()

        ctx = self._make_context(draft_mode=lp.draft_mode)

        self._log(f"\n{'─' * 50}")
        self._log(f"{lp.icon} Running: {lp.name}")

        try:
            result = lp.instance.run(ctx)
        except Exception as e:
            result = PluginResult(success=False, error=str(e))
            # Publish failure event
            EventBus.emit(
                "plugin.run.failed",
                payload={"plugin_id": plugin_id, "error": str(e)},
                source="PluginLoader",
            )

        lp.last_run = datetime.now()
        lp.last_result = "✅ Success" if result.success else f"❌ {result.error}"
        lp.last_summary = result.summary
        lp.persist()

        if not manual:
            lp.schedule_next()

        self._log(f"{lp.icon} Done: {lp.last_summary or lp.last_result}")
        self._log(f"{'─' * 50}\n")

        # Publish completion event so other plugins/components can react
        EventBus.emit(
            "plugin.run.complete",
            payload={
                "plugin_id": plugin_id,
                "success": result.success,
                "summary": result.summary,
                "actions_taken": result.actions_taken,
                "drafts_created": result.drafts_created,
            },
            source="PluginLoader",
        )

        if self._on_run_complete:
            self._on_run_complete(plugin_id, result)

        return result

    def run_all_due(self):
        """Check all plugins and run any that are due."""
        for pid, lp in self._plugins.items():
            if lp.is_due():
                self.run_plugin(pid)

    # ── Scheduler ─────────────────────────────────────────────────────────────

    def start_scheduler(self):
        """Start the background scheduler thread and the heartbeat."""
        if self._running:
            return
        self._running = True
        self._scheduler_thread = threading.Thread(
            target=self._scheduler_loop, daemon=True
        )
        self._scheduler_thread.start()
        # Start the heartbeat — fires every 60 s by default
        heartbeat_interval = int(get_setting("heartbeat_interval_seconds", "60"))
        HeartbeatPlugin.start(interval_seconds=heartbeat_interval)
        # Wire cross-plugin event handlers (Tier 3A — APEX multi-agent wiring)
        plugin_instances = {
            pid: lp.instance for pid, lp in self._plugins.items()
            if lp.instance is not None}
        wire_all(EventBus, plugin_instances)
        apply_all_patches(plugin_instances, EventBus)
        EventBus.emit("app.started", source="PluginLoader")
        self._log("⏱ Scheduler started with event wiring.")

    def stop_scheduler(self):
        """Stop the background scheduler and heartbeat."""
        self._running = False
        HeartbeatPlugin.stop()
        EventBus.emit("app.stopping", source="PluginLoader")
        self._log("⏱ Scheduler stopped.")

    def _on_heartbeat_tick(self, event) -> None:
        """Called on every heartbeat tick — runs any due plugins."""
        if self._running:
            try:
                self.run_all_due()
            except Exception as e:
                self._log(f"⚠ Heartbeat tick error: {e}")

    def _is_within_business_hours(self) -> bool:
        """Check if current Melbourne time is within configured business hours."""
        if get_setting("business_hours_enabled", "1") != "1":
            return True  # business hours check disabled

        if pytz is None:
            return True  # can't check without pytz — allow all

        try:
            melb_tz = pytz.timezone("Australia/Melbourne")
            now = datetime.now(melb_tz)

            # Check day of week (isoweekday: 1=Mon … 7=Sun)
            business_days_str = get_setting("business_days", "1,2,3,4,5")
            business_days = [
                int(d.strip()) for d in business_days_str.split(",") if d.strip()
            ]
            if now.isoweekday() not in business_days:
                return False

            start_hour = int(get_setting("business_hours_start", "8"))
            end_hour = int(get_setting("business_hours_end", "18"))
            if not (start_hour <= now.hour < end_hour):
                return False

            return True
        except Exception:
            return True  # on error, allow execution

    def _scheduler_loop(self):
        _outside_hours_logged = False
        while self._running:
            try:
                if not self._is_within_business_hours():
                    if not _outside_hours_logged:
                        self._log("⏱ Outside business hours — scheduler paused.")
                        _outside_hours_logged = True
                    time.sleep(10)
                    continue
                _outside_hours_logged = False
                self.run_all_due()
            except Exception as e:
                self._log(f"⚠ Scheduler error: {e}")
            time.sleep(10)  # check every 10s

    # ── Plugin management (called from UI) ───────────────────────────────────

    def get_plugins(self) -> list[LoadedPlugin]:
        return list(self._plugins.values())

    def get_plugin(self, plugin_id: str) -> LoadedPlugin | None:
        return self._plugins.get(plugin_id)

    def set_plugin_enabled(self, plugin_id: str, enabled: bool):
        lp = self._plugins.get(plugin_id)
        if lp:
            lp.enabled = enabled
            if enabled:
                lp.schedule_next()
            lp.persist()

    def set_plugin_draft_mode(self, plugin_id: str, draft_mode: bool):
        lp = self._plugins.get(plugin_id)
        if lp:
            lp.draft_mode = draft_mode
            lp.persist()

    def set_plugin_schedule(self, plugin_id: str, seconds: int):
        lp = self._plugins.get(plugin_id)
        if lp:
            lp.schedule_seconds = seconds
            lp.schedule_next()
            lp.persist()

    def reload_plugin(self, plugin_id: str):
        """Re-call load() on a plugin (e.g. after settings change)."""
        lp = self._plugins.get(plugin_id)
        if lp:
            ctx = self._make_context(draft_mode=lp.draft_mode)
            try:
                lp.is_ready = lp.instance.load(ctx)
            except Exception as e:
                self._log(f"⚠ Reload failed for {lp.name}: {e}")
                lp.is_ready = False

    def reload_plugins(self) -> list[str]:
        """Re-scan plugins/ directory and load any newly added plugins.
        Also removes plugins whose files have been deleted from disk."""
        # Remove plugins whose source files no longer exist
        existing_files = set()
        if PLUGINS_DIR.exists():
            existing_files = {p.stem for p in PLUGINS_DIR.glob("plugin_*.py")}

        removed_ids = [
            pid for pid in list(self._plugins.keys())
            if pid not in existing_files
        ]
        for pid in removed_ids:
            lp = self._plugins.pop(pid)
            self._log(f"  Unregistered deleted plugin: {lp.name} ({pid})")
            # Also remove from sys.modules so it can be cleanly re-imported if re-added
            if pid in sys.modules:
                del sys.modules[pid]

        new_ids = self.discover()
        ctx = self._make_context(draft_mode=True)
        for pid in new_ids:
            lp = self._plugins.get(pid)
            if lp and not lp.is_template:
                try:
                    lp.is_ready = lp.instance.load(ctx)
                    lp.schedule_next()
                except Exception as e:
                    self._log(f"⚠ Plugin {lp.name} failed to load: {e}")
                    lp.is_ready = False
        return new_ids

    # ── Context factory ────────────────────────────────────────────────────────────

    def _make_context(self, draft_mode: bool) -> PluginContext:
        from config import get_all_settings

        def notify(subject: str, body: str, to: str = None):
            user_email = get_setting("user_email")
            recipients = [to] if to else ([user_email] if user_email else [])
            for email_addr in recipients:
                try:
                    if self._graph and email_addr:
                        self._graph.send_email(email_addr, subject, body)
                except Exception as e:
                    self._log(f"⚠ Notify failed: {e}")

        # Lazy-load MemoryStore — gracefully degrades if chromadb not installed
        memory = None
        try:
            from memory_store import MemoryStore
            memory = MemoryStore
        except Exception as e:
            self._log(f"⚠ MemoryStore unavailable: {e}")

        # Lazy-load GatewayClient — gracefully degrades if credentials not set
        gateway = None
        try:
            from gateway_client import GatewayClient
            gateway = GatewayClient()
        except Exception as e:
            self._log(f"⚠ GatewayClient unavailable: {e}")

        # Lazy-load ApprovalQueue — gracefully degrades if DB unavailable
        approval_queue = None
        try:
            from approval_queue import get_approval_queue
            approval_queue = get_approval_queue()
            # Wire up the EventBus so the queue can publish events
            approval_queue.set_event_bus(EventBus)
        except Exception as e:
            self._log(f"⚠ ApprovalQueue unavailable: {e}")

        ctx = PluginContext(
            graph=self._graph,
            claude=self._claude,           # legacy alias — same as claude_fast
            claude_fast=self._claude_fast,
            claude_reason=self._claude_reason,
            memory=memory,
            event_bus=EventBus,
            gateway=gateway,
            approval_queue=approval_queue,
            log=self._log,
            notify=notify,
            settings=get_all_settings(),
            draft_mode=draft_mode,
        )
        # Wrap Claude clients with token metering (Tier 3B)
        try:
            from token_meter import wrap_context_claude
            wrap_context_claude(ctx, plugin_id="plugin_loader")
        except Exception as e:
            self._log(f"⚠ TokenMeter unavailable: {e}")
        return ctx