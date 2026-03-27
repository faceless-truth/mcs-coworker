"""
MC & S Desktop Agent — Plugin Base Class
=========================================
Every plugin must inherit from AgentPlugin and implement the required methods.

HOW TO CREATE A NEW PLUGIN
---------------------------
1. Create a new file in the plugins/ folder:  plugins/plugin_your_name.py
2. Import and subclass AgentPlugin
3. Implement all @abstractmethod methods
4. The app will auto-discover and load it on next launch

WHAT A PLUGIN CAN DO
---------------------
Each plugin receives a shared "context" object when it runs, giving it access to:

  - context.graph    → Microsoft Graph API (email, calendar, OneDrive)
  - context.claude   → Anthropic Claude API client
  - context.log(msg) → Write to the live dashboard log
  - context.notify(…)→ Send a staff notification email
  - context.settings → Read any app-wide setting

Plugins decide their own:
  - Schedule (how often to run, or manual-only)
  - Draft mode (create drafts vs act automatically)
  - Whether they need Microsoft 365, Claude, or neither

PLUGIN LIFECYCLE
----------------
  load() → Called once when app starts. Set up state, validate config.
  run()  → Called on schedule or manually. Do the work.
  stop() → Called when app closes or plugin is disabled mid-run.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any


# ── Schedule helpers ──────────────────────────────────────────────────────────

class Schedule:
    """
    Describe when a plugin should run.

    Examples:
        Schedule.every_minutes(5)
        Schedule.every_hours(4)
        Schedule.daily_at(hour=8)
        Schedule.manual_only()
    """

    def __init__(self, interval_seconds: int = 0, label: str = "Manual only"):
        self.interval_seconds = interval_seconds
        self.label = label

    @classmethod
    def every_seconds(cls, n: int) -> "Schedule":
        return cls(interval_seconds=n, label=f"Every {n}s")

    @classmethod
    def every_minutes(cls, n: int) -> "Schedule":
        return cls(interval_seconds=n * 60, label=f"Every {n} min")

    @classmethod
    def every_hours(cls, n: int) -> "Schedule":
        return cls(interval_seconds=n * 3600, label=f"Every {n} hr")

    @classmethod
    def daily_at(cls, hour: int) -> "Schedule":
        return cls(interval_seconds=86400, label=f"Daily at {hour:02d}:00")

    @classmethod
    def manual_only(cls) -> "Schedule":
        return cls(interval_seconds=0, label="Manual only")

    def is_scheduled(self) -> bool:
        return self.interval_seconds > 0


# ── Run result ─────────────────────────────────────────────────────────────────

@dataclass
class PluginResult:
    """
    Returned by plugin.run(). Tells the app what happened.

    Examples:
        return PluginResult(success=True, summary="3 emails processed, 2 drafted")
        return PluginResult(success=False, error="API connection failed")
        return PluginResult(success=True, actions_taken=5, drafts_created=2)
    """
    success: bool = True
    summary: str = ""
    error: str = ""
    actions_taken: int = 0
    drafts_created: int = 0
    items_skipped: int = 0
    extra: dict = field(default_factory=dict)


# ── Plugin context ────────────────────────────────────────────────────────────

@dataclass
class PluginContext:
    """
    Passed to plugin.run(). Provides shared services.
    All fields may be None if not configured — plugins should check before use.
    """
    graph: Any = None       # graph_client.GraphClient instance
    claude: Any = None      # anthropic.Anthropic client instance
    log: Any = None         # callable: log(message: str)
    notify: Any = None      # callable: notify(subject, body, to=None)
    settings: dict = field(default_factory=dict)  # all app settings as dict
    draft_mode: bool = True  # plugin's individual draft mode setting


# ── Base class ────────────────────────────────────────────────────────────────

class AgentPlugin(ABC):
    """
    Base class for all MC & S desktop agent plugins.
    Subclass this and implement the abstract methods.
    """

    # ── Identity (override these in your plugin) ──────────────────────────────

    #: Human-readable plugin name shown in the UI
    name: str = "Unnamed Plugin"

    #: One-line description shown in the Plugins tab
    description: str = "No description provided."

    #: Longer explanation of what this plugin does (shown in info panel)
    detail: str = ""

    #: Semantic version string
    version: str = "1.0.0"

    #: Emoji icon shown next to the plugin name in the UI
    icon: str = "🔌"

    #: Author name (optional)
    author: str = "MC & S"

    #: Whether this plugin needs Microsoft Graph (email/calendar access)
    requires_graph: bool = False

    #: Whether this plugin needs the Claude/Anthropic API
    requires_claude: bool = False

    # ── Schedule ──────────────────────────────────────────────────────────────

    #: Default schedule. Override in subclass.
    default_schedule: Schedule = Schedule.manual_only()

    # ── Settings schema ───────────────────────────────────────────────────────

    @classmethod
    def email_templates_schema(cls) -> list[dict]:
        """
        Declare editable email templates this plugin exposes in the UI.

        Return a list of template definitions. Example:
            return [
                {"key": "draft_prompt", "label": "Draft Email Prompt",
                 "default": "You are a professional assistant...",
                 "type": "prompt"},
                {"key": "email_closing", "label": "Sign-off Text",
                 "default": "Kind regards,", "type": "textarea"},
            ]

        Supported types: "text", "textarea", "prompt"
        """
        return []

    @classmethod
    def settings_schema(cls) -> list[dict]:
        """
        Declare plugin-specific settings that appear in the Plugins tab.

        Return a list of field definitions. Example:
            return [
                {"key": "folder_to_watch", "label": "Folder to Watch",
                 "default": "Inbox", "type": "text"},
                {"key": "min_days_overdue", "label": "Min Days Overdue",
                 "default": "45", "type": "number"},
                {"key": "send_to_all_staff", "label": "Notify All Staff",
                 "default": "1", "type": "bool"},
            ]

        Supported types: "text", "password", "number", "bool", "textarea"
        These are rendered automatically in the plugin settings panel.
        """
        return []

    # ── Lifecycle (implement these) ───────────────────────────────────────────

    def load(self, context: PluginContext) -> bool:
        """
        Called once when the app starts (or when plugin is enabled).
        Use this to validate config, set up state, connect to APIs.
        Return True if ready to run, False if misconfigured
        (the plugin will show as "Not Ready" in the UI).
        """
        return True

    @abstractmethod
    def run(self, context: PluginContext) -> PluginResult:
        """
        The main work method. Called on schedule or when "Run Now" is clicked.

        - Use context.log() to write progress to the dashboard
        - Check context.draft_mode before taking irreversible actions
        - Always return a PluginResult
        - Catch your own exceptions and return PluginResult(success=False, error=str(e))
        """
        ...

    def stop(self):
        """
        Called when plugin is disabled or app is closing mid-run.
        Clean up any open connections or state here.
        """
        pass

    # ── Helpers available to subclasses ──────────────────────────────────────

    def get_plugin_setting(self, key: str, default: str = "") -> str:
        """Read a plugin-specific setting from the database."""
        from config import get_setting
        return get_setting(f"plugin_{self.__class__.__name__}_{key}", default)

    def set_plugin_setting(self, key: str, value: str):
        """Write a plugin-specific setting to the database."""
        from config import set_setting
        set_setting(f"plugin_{self.__class__.__name__}_{key}", value)

    def get_email_template(self, key: str, default: str = "") -> str:
        """Read an editable email template from the database."""
        from config import get_plugin_template
        plugin_id = "plugin_" + self.__class__.__name__.replace("Plugin", "").lower()
        # Try the class-based ID first, then fall back to module-style
        result = get_plugin_template(plugin_id, key, None)
        if result is None:
            # Also try with the full class name style used by plugin_loader
            result = get_plugin_template(
                f"plugin_{self.__class__.__module__}", key, None
            )
        return result if result is not None else default

    def get_claude_model(self) -> str:
        """Return the current Claude model string from settings."""
        from config import get_claude_model
        return get_claude_model()

    def log_activity(self, source: str, subject: str, category: str,
                     action: str, draft_created: int = 0,
                     notification_sent: int = 0):
        """Write an entry to the shared activity log."""
        from config import log_activity
        log_activity(source, subject, category, action,
                     draft_created, notification_sent)
