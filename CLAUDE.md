# MC & S Desktop Agent — Claude Code Context

## Project Overview
Windows desktop automation agent for **MC & S Pty Ltd**, an accounting practice in
Keysborough, Victoria, Australia. The app runs in the background and automates
routine tasks (starting with email triage) on a configurable schedule — without
staff needing to be present.

---

## Tech Stack
| Component | Detail |
|-----------|--------|
| UI | Python + CustomTkinter (modern tkinter wrapper) |
| Database | SQLite at `~/.mcs_email_automation/config.db` |
| Auth | Microsoft MSAL — OAuth2 device code flow, tokens cached |
| AI | Anthropic API — `claude-haiku-4-5-20251001` |
| Email | Microsoft Graph API — `Mail.ReadWrite` + `Mail.Send` |
| Timezone | pytz — Melbourne (AUS Eastern Standard Time) |
| Architecture | Plugin-based — each automation is a self-contained module |

---

## File Structure
```
mcs-coworker/
├── app.py                          # Main UI — CustomTkinter, 6 nav tabs
├── config.py                       # SQLite DB manager + all CRUD functions
├── graph_client.py                 # Microsoft Graph API wrapper
├── plugin_base.py                  # AgentPlugin base class + Schedule + PluginResult + PluginContext
├── plugin_loader.py                # Plugin discovery, scheduling, execution engine
├── requirements.txt                # pip dependencies
├── launch.bat                      # Windows launcher (creates venv on first run)
├── CLAUDE.md                       # This file
└── plugins/
    ├── __init__.py
    ├── plugin_email_triage.py      # ACTIVE: email classification + auto-response
    └── plugin_template.py          # Template only — never runs (in TEMPLATE_PLUGIN_IDS)
```

---

## UI Layout (app.py)
**Colour scheme:**
- `BRAND_BLUE` = `#1565C0`
- `BRAND_DARK` = `#0D47A1`
- `ACCENT_GREEN` = `#2E7D32`
- `ACCENT_AMBER` = `#E65100`
- `BG_LIGHT` = `#F5F7FA`
- `CARD_BG` = `#FFFFFF`

**Layout:** Dark left nav (210px) + content area. Header bar (64px) shows app title,
scheduler status pill, and auth status.

**Nav tabs (in order):**
1. Dashboard — scheduler start/stop, live log, stat counters
2. Plugins — per-plugin cards with enable/disable, draft mode, schedule, Run Now
3. Email Rules — CRUD editor for categories, keywords, response templates
4. Staff & Notify — staff who receive draft notification emails
5. Settings — MS365 credentials, Anthropic API key, business hours
6. Activity Log — timestamped audit trail

---

## Plugin Architecture

### plugin_base.py defines:
- `class Schedule` — classmethods: `every_minutes(n)`, `every_hours(n)`, `daily_at(hour)`, `manual_only()`
- `dataclass PluginResult` — `success`, `summary`, `error`, `actions_taken`, `drafts_created`, `items_skipped`, `extra`
- `dataclass PluginContext` — `graph`, `claude`, `log` callable, `notify` callable, `settings` dict, `draft_mode` bool
- `abstract class AgentPlugin` — class attributes: `name`, `description`, `detail`, `version`, `icon`, `author`, `requires_graph`, `requires_claude`, `default_schedule`
  - `classmethod settings_schema() -> list[dict]`
  - `def load(context) -> bool`
  - `abstract def run(context) -> PluginResult`
  - `def stop()`
  - helpers: `get_plugin_setting(key)`, `set_plugin_setting(key, value)`, `log_activity(...)`

### plugin_loader.py:
- Scans `plugins/` for `plugin_*.py` files on startup
- Auto-discovers and registers any `AgentPlugin` subclass
- `LoadedPlugin` wrapper per plugin: `plugin_id`, `enabled`, `draft_mode`, `schedule_seconds`, `last_run`, `last_result`, `last_summary`, `is_ready`, `_next_run_at`
- Persists all plugin state to `plugin_registry` table in SQLite
- `TEMPLATE_PLUGIN_IDS` set — templates shown in UI but NEVER run
- Scheduler loop runs every 10 seconds in a background thread

---

## Database Schema (config.py)
| Table | Key columns |
|-------|-------------|
| `settings` | key, value — all app-wide config |
| `plugin_registry` | plugin_id, enabled, draft_mode, schedule_seconds, last_run, last_result |
| `email_rules` | id, category, keywords, subject_template, body_template, enabled, sort_order |
| `staff_notifications` | id, name, email, receives_drafts, enabled |
| `activity_log` | id, timestamp, from_email, subject, classification, action, draft_created, notification_sent |

**Default settings seeded on first run:**
`draft_mode=1`, `business_hours_enabled=1`, `business_hours_start=8`,
`business_hours_end=18`, `business_days=1,2,3,4,5`, `polling_interval=60`,
`ms_tenant_id=''`, `ms_client_id=''`, `anthropic_api_key=''`,
`ms_account_email=''`, `monitor_folder=Inbox`, `practice_name=MC & S`,
`timezone=AUS Eastern Standard Time`

---

## Email Triage Plugin (plugins/plugin_email_triage.py)
- `name = 'Email Triage & Auto-Response'`
- `default_schedule = Schedule.every_minutes(1)`
- `requires_graph = True`, `requires_claude = True`
- Settings schema: `folder` (default: Inbox), `max_per_run` (default: 25)

**Run logic:**
1. Fetch unread emails from configured folder
2. Strip HTML from body
3. Call Claude (haiku) — returns JSON: `{category, confidence, reasoning, sender_name}`
4. `OTHER` → leave in inbox, log no_action
5. Matched → apply template (replace `{client_name}`, `{date}`, `{subject}`)
6. `draft_mode ON` → `create_draft()` + send staff notification email
7. `draft_mode OFF` → `send_email()` directly
8. `DOCUMENTS_RECEIVED` → also `flag_email()` for follow-up
9. `mark_as_read()` on processed emails
10. Track processed message IDs in a set to avoid reprocessing

**Seeded email rules:**
- `PRICING_ENQUIRY` — keywords: how much, price, cost, fee, rates, quote
- `CHECKLIST_REQUEST` — keywords: what do i need, checklist, what to bring
- `DOCUMENTS_RECEIVED` — keywords: please find attached, here are my documents

---

## Draft Mode Behaviour (applies to ALL plugins)
- **ON** → create Outlook draft + send HTML notification email to all staff with `receives_drafts=1`
- **OFF** → send/action automatically, no notification
- **Default is always ON** — auto-send requires a conscious decision to flip

**Staff notification email format:**
- Blue header: "Draft Email Ready for Review"
- Table rows: From, Subject, Category (blue pill badge)
- Amber box: "Action Required — open Outlook Drafts, review, then send"
- Practice name footer

---

## Business Hours Logic
- Runs before every plugin execution cycle
- Convert UTC → Melbourne time (pytz)
- Check weekday (Mon–Fri) and hour (`start_hour <= hour < end_hour`)
- If outside hours: log message, skip cycle entirely
- Configurable from Settings tab

---

## Scheduler Behaviour
- **Start** requires: MS365 auth + Anthropic API key — shows warning if missing
- Calls `loader.set_graph()`, `loader.set_claude()`, `loader.load_all()`, `loader.start_scheduler()`
- Loop checks every 10s: if `time.time() >= _next_run_at` → `run_plugin()`
- Next run scheduled as: `_next_run_at = time.time() + schedule_seconds`

---

## Non-Negotiable Code Rules
- **Never block the main thread** — all network/API calls go in background threads
- **All UI updates from threads** use `self.after(0, ...)` — never touch widgets directly from threads
- **All DB writes** use parameterised queries — no string formatting in SQL
- **Plugin settings** namespaced as `plugin_{ClassName}_{key}` in the settings table
- **`init_db()` must be idempotent** — safe to run multiple times
- **Plugin loader** handles import errors gracefully — log, skip, continue
- **Claude classification prompt** requests ONLY valid JSON, strip markdown code fences before parsing
- **Graph client `authenticate()`** uses a daemon thread for the callback wait

---

## Microsoft Graph Client (graph_client.py)
- `PublicClientApplication` with `SerializableTokenCache`
- Scopes: `Mail.ReadWrite`, `Mail.Send`, `offline_access`
- Redirect URI: `http://localhost:8765` (captured by local HTTPServer)
- Required methods: `authenticate(callback)`, `is_authenticated()`, `get_user_info()`,
  `fetch_unread_emails(folder, max_count)`, `mark_as_read(message_id)`,
  `send_email(to, subject, body_html, reply_to_id)`, `create_draft(to, subject, body_html, reply_to_id)`,
  `flag_email(message_id)`, `add_category(message_id, category)`

---

## Adding a New Plugin
**Only create `plugins/plugin_{name}.py` — do not modify any other files.**

Template pattern:
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
        return []  # Add field defs if needed

    def load(self, context: PluginContext) -> bool:
        return True  # Return False if not ready

    def run(self, context: PluginContext) -> PluginResult:
        # Your automation logic here
        return PluginResult(success=True, summary="Done", actions_taken=0)
```

---

## Planned Plugins (backlog)
| Plugin | Schedule | Description |
|--------|----------|-------------|
| `plugin_noa_workflow.py` | Daily 8am | Detect Notice of Assessment emails, extract refund/payable amounts, draft personalised cover email |
| `plugin_fusesign_monitor.py` | Every 4hrs | Check FuseSign bundles unsigned >X days, draft polite nudge to client |
| `plugin_meeting_prep.py` | Scheduled | 30min before calendar appointment: pull last invoice + recent emails, compile one-page brief |
| `plugin_monthly_invoicing.py` | 1st of month | Pull retainer clients from Xero, generate invoice list, email to billing staff |
| `plugin_debtor_followup.py` | Daily | Pull debtors aged >45 days, draft progressive follow-up sequence |
| `plugin_asic_reminder.py` | Daily | Parse ASIC reminder emails, extract company + due date, create calendar reminder |
| `plugin_client_checkin.py` | Weekly | Surface companies/trusts not contacted in 6+ months, draft check-in email |

---

## Cost Reference
- Anthropic API (Claude Haiku, ~500 emails/month): ~$1–3 AUD/month
- Microsoft 365: already subscribed, no additional cost
