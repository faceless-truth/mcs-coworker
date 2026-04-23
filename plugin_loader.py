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
import logging
import os
import sys
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, date
from pathlib import Path
from typing import Callable

logger = logging.getLogger(__name__)

try:
    import anthropic
except ImportError:
    anthropic = None

try:
    import pytz
except ImportError:
    pytz = None

from plugin_base import AgentPlugin, PluginContext, PluginResult


def _safe_log(callback: "Callable[[str], None]") -> "Callable[[str], None]":
    """Wrap a log callback so it never raises UnicodeEncodeError.
    On Windows the default stdout uses cp1252 which can't handle emoji.
    We encode to ascii with xmlcharrefreplace so the message is still readable.
    """
    def _log(msg: str) -> None:
        try:
            callback(msg)
        except (UnicodeEncodeError, UnicodeDecodeError):
            safe = msg.encode("ascii", errors="replace").decode("ascii")
            try:
                callback(safe)
            except Exception:
                pass
        except Exception:
            pass
    return _log


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
TEMPLATE_PLUGIN_IDS = {
    "plugin_template",
    "plugin_noa_workflow",
    "plugin_monthly_invoicing",
    "plugin_client_checkin",
}


def _is_plugin_allowed_in_mode(plugin_instance) -> bool:
    """Return True if a plugin's category matches the current reception_mode.

    - Universal plugins run everywhere.
    - Reception plugins run only when reception_mode = "1".
    - Accountant plugins run only when reception_mode = "0".
    """
    category = getattr(plugin_instance, "category", "universal")
    reception_mode = get_setting("reception_mode", "0") == "1"
    if category == "reception" and not reception_mode:
        return False
    if category == "accountant" and reception_mode:
        return False
    return True


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
        if not (self.enabled and self.is_ready and not self.is_template
                and is_scheduled and self._next_run_at > 0
                and time.time() >= self._next_run_at):
            return False
        # Mode gate: category must match the current reception_mode
        return _is_plugin_allowed_in_mode(self.instance)


class PluginLoader:
    """
    Discovers, loads, and schedules all plugins.
    Provides the interface used by the main app UI.
    """

    def __init__(self, log_callback: Callable[[str], None] = print):
        self._log = _safe_log(log_callback)
        self._plugins: dict[str, LoadedPlugin] = {}
        self._graph = None
        self._claude = None        # legacy — kept for backward compat, mirrors _claude_fast
        self._claude_fast = None   # Haiku: fast, cheap tasks
        self._claude_reason = None # Sonnet: complex reasoning & analysis
        self._scheduler_thread: threading.Thread | None = None
        self._running = False
        self._on_run_complete: Callable | None = None  # called after each plugin run
        self._on_plugin_registered: Callable | None = None  # called when new plugin discovered
        # Shared lock between the heartbeat-tick path and the scheduler-loop
        # path. Both call run_all_due(); without this lock a plugin can fire
        # twice because the second path sees its _next_run_at before the first
        # has updated it.
        self._run_lock = threading.Lock()
        # Pool for concurrent plugin execution. Without this, a slow plugin
        # (e.g. Smart Email Responder working a busy inbox) blocks every other
        # plugin behind it on the scheduler thread.
        self._plugin_executor = ThreadPoolExecutor(
            max_workers=4, thread_name_prefix="plugin-worker"
        )
        # Tracks plugins currently mid-run so we don't resubmit before the
        # previous run completes (cheap guard in addition to _run_lock).
        self._inflight: set[str] = set()
        self._inflight_lock = threading.Lock()
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
            # Log the full traceback via the logging framework (visible in
            # stdout) AND append it to a dedicated file so remote accountants'
            # crashes are diagnosable without shelling in.
            logger.error(
                f"Plugin {lp.name} failed: {e}", exc_info=True
            )
            try:
                from config import DATA_DIR as _data_dir
                error_log_path = _data_dir / "plugin_errors.log"
                error_log_path.parent.mkdir(parents=True, exist_ok=True)
                with open(error_log_path, "a", encoding="utf-8") as f:
                    f.write(f"\n{'=' * 60}\n")
                    f.write(f"[{datetime.now().isoformat()}] Plugin: {lp.name}\n")
                    traceback.print_exc(file=f)
            except Exception:
                pass
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
        """Submit every due plugin to the thread pool so slow plugins don't
        block fast ones. Each run is fire-and-forget; errors are caught inside
        run_plugin() and logged by _on_plugin_complete.
        """
        for pid, lp in list(self._plugins.items()):
            if not lp.is_due():
                continue
            # Skip if we already submitted this plugin and it's still running.
            # Prevents a long-running plugin from stacking duplicate submissions
            # if subsequent scheduler ticks see it as "due" before completion.
            with self._inflight_lock:
                if pid in self._inflight:
                    continue
                self._inflight.add(pid)
            future = self._plugin_executor.submit(self._run_plugin_inflight, pid)
            future.add_done_callback(
                lambda f, _pid=pid: self._on_plugin_complete(f, _pid)
            )

    def _run_plugin_inflight(self, plugin_id: str) -> PluginResult:
        """Wrapper around run_plugin that always clears the in-flight marker."""
        try:
            return self.run_plugin(plugin_id)
        finally:
            with self._inflight_lock:
                self._inflight.discard(plugin_id)

    def _on_plugin_complete(self, future, plugin_id: str) -> None:
        """Log any exception that escaped run_plugin's own try/except."""
        try:
            exc = future.exception()
            if exc:
                self._log(f"⚠ Plugin {plugin_id} raised uncaught exception: {exc}")
        except Exception:
            pass

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
        # Don't wait for in-flight plugins — if the user is shutting down we
        # want the window to close immediately. Uncaught errors are already
        # logged by _on_plugin_complete.
        self._plugin_executor.shutdown(wait=False)
        EventBus.emit("app.stopping", source="PluginLoader")
        self._log("⏱ Scheduler stopped.")

    def _on_heartbeat_tick(self, event) -> None:
        """Called on every heartbeat tick — runs any due plugins."""
        if not self._running:
            return
        # Respect configured business hours — without this the heartbeat will
        # fire plugins at midnight on weekends regardless of the setting.
        if not self._is_within_business_hours():
            return
        # Skip if the scheduler loop is already running a batch; we don't want
        # both paths to double-fire plugins whose _next_run_at hasn't been
        # updated yet.
        if not self._run_lock.acquire(blocking=False):
            self._log("⏱ Heartbeat skipped — run_all_due already in progress")
            return
        try:
            self.run_all_due()
        except Exception as e:
            self._log(f"⚠ Heartbeat tick error: {e}")
        finally:
            self._run_lock.release()

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

            # Check Australian public holidays (skip if disabled in settings)
            if get_setting("skip_public_holidays", "1") == "1":
                if self._is_australian_public_holiday(now.date()):
                    return False

            start_hour = int(get_setting("business_hours_start", "8"))
            end_hour = int(get_setting("business_hours_end", "18"))
            if not (start_hour <= now.hour < end_hour):
                return False

            return True
        except Exception:
            return True  # on error, allow execution

    @staticmethod
    def _is_australian_public_holiday(check_date) -> bool:
        """
        Returns True if check_date is an Australian national or Victorian public holiday.
        Covers: New Year's Day, Australia Day, Good Friday, Easter Saturday/Monday,
        ANZAC Day, King's Birthday (VIC), AFL Grand Final Friday (VIC),
        Melbourne Cup (VIC), Christmas Day, Boxing Day.
        State can be overridden via settings key 'public_holiday_state' (default: VIC).
        """
        from datetime import date, timedelta
        y = check_date.year
        state = get_setting("public_holiday_state", "VIC").upper()

        # ── Fixed national holidays ───────────────────────────────────────────
        fixed = [
            date(y, 1, 1),   # New Year's Day
            date(y, 1, 26),  # Australia Day (observed)
            date(y, 4, 25),  # ANZAC Day
            date(y, 12, 25), # Christmas Day
            date(y, 12, 26), # Boxing Day
        ]
        # Substitute rule: if fixed holiday falls on weekend, observed on Monday
        observed = set()
        for h in fixed:
            if h.weekday() == 5:   # Saturday → Monday
                observed.add(h + timedelta(days=2))
            elif h.weekday() == 6: # Sunday → Monday
                observed.add(h + timedelta(days=1))
            else:
                observed.add(h)

        # ── Easter (Good Friday, Easter Saturday, Easter Monday) ──────────────
        # Gauss algorithm for Easter Sunday
        a = y % 19
        b, c = divmod(y, 100)
        d, e = divmod(b, 4)
        f = (b + 8) // 25
        g = (b - f + 1) // 3
        h = (19 * a + b - d - g + 15) % 30
        i, k = divmod(c, 4)
        l = (32 + 2 * e + 2 * i - h - k) % 7
        m = (a + 11 * h + 22 * l) // 451
        month = (h + l - 7 * m + 114) // 31
        day   = ((h + l - 7 * m + 114) % 31) + 1
        easter_sunday = date(y, month, day)
        observed.add(easter_sunday - timedelta(days=2))  # Good Friday
        observed.add(easter_sunday - timedelta(days=1))  # Easter Saturday
        observed.add(easter_sunday + timedelta(days=1))  # Easter Monday

        # ── Victorian-specific holidays ───────────────────────────────────────
        if state == "VIC":
            # King's Birthday: 2nd Monday of June
            first_monday_june = date(y, 6, 1)
            while first_monday_june.weekday() != 0:
                first_monday_june += timedelta(days=1)
            observed.add(first_monday_june + timedelta(weeks=1))
            # Melbourne Cup: 1st Tuesday of November
            first_tue_nov = date(y, 11, 1)
            while first_tue_nov.weekday() != 1:
                first_tue_nov += timedelta(days=1)
            observed.add(first_tue_nov)

        return check_date in observed

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
                # Skip if the heartbeat tick path is already running a batch.
                if self._run_lock.acquire(blocking=False):
                    try:
                        self.run_all_due()
                    finally:
                        self._run_lock.release()
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