# MC & S CoWorker — Claude Code Context

## Project Overview
Windows desktop automation agent for **MC & S Pty Ltd**, an accounting practice in
Keysborough, Victoria, Australia. Runs in the background and automates routine
tasks (email triage, NOA processing, debtor follow-up, FuseSign nudges, meeting
prep, etc.) on configurable schedules — without staff needing to be present.

---

## Tech Stack
| Component | Detail |
|-----------|--------|
| Frontend | React + Vite (built into `frontend/dist/public`) |
| Native window | pywebview — **EdgeChromium backend** (WebView2) |
| API server | Flask + flask-cors on `127.0.0.1:7842` |
| Backend | Python 3.11 (embeddable runtime in production) |
| Database | SQLite — production path resolved via `MCS_DATA_DIR` env var |
| Auth | Microsoft MSAL (single-tenant, OAuth2) |
| AI | Anthropic API — `claude-haiku-4-5-20251001` |
| Email | Microsoft Graph API — `Mail.ReadWrite` + `Mail.Send` |
| System tray | pystray + Pillow |
| Timezone | pytz — Melbourne (AUS Eastern Standard Time) |
| Architecture | Plugin-based — each automation is a self-contained module |

---

## Distribution Model (read this before changing anything)
Accountants are non-technical. They never run pip, cmd, or python. The delivery is:

1. Dev (Elio) runs `build_installer.bat` → produces `installer_output\MCSCoWorker_Setup.exe`
2. Installer is uploaded to SharePoint
3. Accountants download and double-click the .exe
4. Inno Setup installs to `%LOCALAPPDATA%\Programs\MCS CoWorker\` (no admin needed)
5. Bundled embeddable Python 3.11 + all packages from `requirements.txt` are pre-installed
6. On every launch, `auto_updater.py` does `git pull` + `pip install -r requirements.txt`
   — so code/package updates flow without rebuilding the installer
7. Only when adding a brand-new BINARY dependency (rare) is an installer rebuild required

`requirements.txt` is the **single source of truth**. Both `build_installer.bat` and
`auto_updater.py` read it. Never hardcode a package list anywhere else.

---

## Launch Chain (Production)
```
Desktop shortcut
  → wscript.exe MCSCoWorker.vbs           (no console window)
    → pythonw.exe app\launcher.py         (auto_updater: git pull + pip install)
      → pythonw.exe app\main.py           (Flask + pywebview + tray icon)
```

Install layout:
```
%LOCALAPPDATA%\Programs\MCS CoWorker\
  MCSCoWorker.vbs
  python\          ← embeddable Python 3.11 + all packages pre-installed
  app\             ← git clone of this repo (auto-updates via git pull)
    main.py
    api_server.py
    plugins\
    frontend_dist\ ← built React app
  data\            ← SQLite DB, logs, startup_error.log (never touched by updates)
  assets\
```

---

## Launch Chain (Dev — Elio's machine only)
```
cd C:\Users\Elio\mcs-coworker
venv\Scripts\activate
set PYWEBVIEW_GUI=edgechromium
python main.py
```

---

## File Structure
```
mcs-coworker/
├── main.py                           # Entry point: Flask + pywebview + tray
├── launcher.py                       # Pre-launch: runs auto_updater, then main.py
├── auto_updater.py                   # git pull + pip install -r requirements.txt
├── api_server.py                     # Flask routes — all UI <-> backend traffic
├── config.py                         # SQLite schema + CRUD
├── plugin_base.py                    # AgentPlugin base + Schedule + PluginResult + PluginContext
├── plugin_loader.py                  # Plugin discovery, scheduling, execution
├── graph_client.py                   # Microsoft Graph API wrapper (tenant/client IDs hardcoded)
├── gateway_client.py                 # Outbound gateway for downstream services
├── event_bus.py                      # In-process pub/sub
├── event_wiring.py                   # Wires plugins to event_bus topics
├── token_meter.py                    # Anthropic token usage tracker
├── memory_store.py                   # Vector / KB style memory for plugins
├── approval_queue.py                 # Human-in-loop approvals (SQLite-backed)
├── kpi_monitor.py                    # KPI snapshot/aggregation singleton
├── xero_oauth.py                     # Xero OAuth2 flow
├── requirements.txt                  # SINGLE SOURCE OF TRUTH for packages
├── build_installer.bat               # Dev only — builds MCSCoWorker_Setup.exe
├── installer.iss                     # Inno Setup script
├── update.bat                        # Dev convenience — git pull + pip + frontend build + copy
├── launch.bat / launch_silent.vbs    # Local dev launchers
├── frontend/                         # React + Vite source
├── plugins/                          # ~18 plugins (see below)
├── assets/                           # Icons, images, templates
└── tests/
```

---

## Plugins (in `plugins/`)
| File | Purpose |
|------|---------|
| `plugin_smart_responder.py` | Unified email classification, triage, and auto-reply (replaces the retired triage/ross/elio/reply plugins) |
| `plugin_noa_processor.py` | Notice of Assessment detection + cover email |
| `plugin_tax_return_processor.py` | Tax-return workflow automation |
| `plugin_asic_returns.py` | Parse ASIC reminder emails, log/calendar |
| `plugin_bas_reminder.py` | BAS reminders |
| `plugin_debtor_followup.py` | Aged-debtor progressive follow-up |
| `plugin_fusesign_monitor.py` | Nudge clients on unsigned FuseSign bundles |
| `plugin_meeting_prep.py` | Pre-meeting client brief |
| `plugin_morning_briefing.py` | Daily morning summary |
| `plugin_engagement_letter.py` | Engagement letter generator |
| `plugin_annual_review.py` | Annual review workflow |
| `plugin_client_outreach.py` | Periodic client check-ins |
| `plugin_correspondence_logger.py` | Correspondence logging |
| `plugin_template.py` | TEMPLATE — listed in `TEMPLATE_PLUGIN_IDS`, never runs |

---

## Plugin Architecture (`plugin_base.py`)
- `class Schedule` — `every_minutes(n)`, `every_hours(n)`, `daily_at(hour)`, `manual_only()`
- `dataclass PluginResult` — `success`, `summary`, `error`, `actions_taken`, `drafts_created`, `items_skipped`, `extra`
- `dataclass PluginContext` — `graph`, `claude`, `log`, `notify`, `settings`, `draft_mode`
- `abstract class AgentPlugin` — class attributes: `name`, `description`, `detail`, `version`, `icon`, `author`, `requires_graph`, `requires_claude`, `default_schedule`
  - `classmethod settings_schema() -> list[dict]`
  - `def load(context) -> bool`
  - `abstract def run(context) -> PluginResult`
  - helpers: `get_plugin_setting(key)`, `set_plugin_setting(key, value)`, `log_activity(...)`

`plugin_loader.py` scans `plugins/plugin_*.py` on startup, registers any
`AgentPlugin` subclass, persists state to `plugin_registry`, and runs a 10-second
scheduler tick in a background thread. Templates listed in `TEMPLATE_PLUGIN_IDS`
are visible in the UI but never execute.

---

## Adding a New Plugin
**Only create `plugins/plugin_{name}.py` — do not modify any other files.**
The loader auto-discovers it on next launch.

```python
from plugin_base import AgentPlugin, Schedule, PluginResult, PluginContext

class MyPlugin(AgentPlugin):
    name = "My Plugin Name"
    description = "One sentence description"
    icon = "🔧"
    version = "1.0.0"
    requires_graph = True
    requires_claude = False
    default_schedule = Schedule.every_hours(4)

    @classmethod
    def settings_schema(cls):
        return []

    def load(self, context: PluginContext) -> bool:
        return True

    def run(self, context: PluginContext) -> PluginResult:
        return PluginResult(success=True, summary="Done", actions_taken=0)
```

---

## Microsoft Graph Client (`graph_client.py`)
- `PublicClientApplication` with `SerializableTokenCache`
- Scopes: `Mail.ReadWrite`, `Mail.Send`, `offline_access`
- Tenant ID and Client ID are **hardcoded** in this file — intentional for
  single-tenant MC & S deployment. Do not parameterise.
- Methods: `authenticate(callback)`, `is_authenticated()`, `get_user_info()`,
  `fetch_unread_emails(folder, max_count)`, `mark_as_read(message_id)`,
  `send_email(to, subject, body_html, reply_to_id)`, `create_draft(...)`,
  `flag_email(message_id)`, `add_category(message_id, category)`

---

## Draft Mode (applies to ALL plugins)
- **ON** → create Outlook draft + send HTML notification email to all staff with `receives_drafts=1`
- **OFF** → send/action automatically, no notification
- **Default is always ON** — auto-send requires a conscious flip

---

## Business Hours Logic
- Runs before every plugin execution cycle
- Convert UTC → Melbourne time (pytz)
- Check weekday (Mon–Fri) and hour (`start_hour <= hour < end_hour`)
- If outside hours: log message, skip the cycle entirely
- Configurable from the Settings tab in the UI

---

## Non-Negotiable Code Rules
- **`requirements.txt` is the single source of truth.** Never hardcode package
  lists in `build_installer.bat`, `update.bat`, or anywhere else.
- **pywebview backend must be edgechromium** — set via
  `os.environ.setdefault("PYWEBVIEW_GUI", "edgechromium")` in `main.py`. The
  default winforms/.NET backend crashes on embeddable Python (cffi init failure).
- **Never block the main thread** — all network/API calls go in background threads.
- **All DB writes** use parameterised queries — no string formatting in SQL.
- **Plugin settings** namespaced as `plugin_{ClassName}_{key}` in the settings table.
- **`init_db()` must be idempotent** — safe to run multiple times.
- **Plugin loader** handles import errors gracefully — log, skip, continue.
- **Claude classification prompt** requests ONLY valid JSON, strip markdown
  fences before parsing.
- **`main.py` has a top-of-file `_log_fatal` excepthook** that writes any uncaught
  exception to `data\startup_error.log`. Do not move it lower or import anything
  before it.

---

## Known Issues / Gotchas
- **`config.py` DB path:** the legacy default is `~/.mcs_email_automation/config.db`,
  which does NOT match the installer's data path (`%INSTALL_DIR%\data\`).
  `launcher.py` sets `MCS_DATA_DIR` but `config.py` currently ignores it. Resolve
  by reading `os.environ["MCS_DATA_DIR"]` in `config.py` before falling back.
- **Frontend fallback:** if `frontend_dist\` is missing, `main.py` falls back to
  `http://127.0.0.1:3000/` (Vite dev server). On accountant machines that's a
  blank window — `_get_frontend_url()` now logs a warning when this happens.

---

## Cost Reference
- Anthropic API (Claude Haiku, ~500 emails/month): ~$1–3 AUD/month
- Microsoft 365: already subscribed, no additional cost
