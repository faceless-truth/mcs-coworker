# MC & S Desktop Agent

A plugin-based desktop automation agent for **MC & S Accountants**. Monitors your Outlook inbox, classifies emails using Claude AI, and drafts or sends tailored responses — all from a modern desktop GUI.

---

## Features

- **Plugin Architecture** — Modular, extensible system. Add new automations by dropping a Python file into the `plugins/` folder.
- **Email Triage & Auto-Response** — Classifies incoming emails (pricing enquiries, checklist requests, documents received) and generates context-aware replies.
- **Draft Mode** — Human-in-the-loop safety. Creates drafts in Outlook for staff review before sending.
- **Staff Notifications** — Notifies team members via email when drafts are ready for review.
- **Configurable Rules** — Define email categories, keywords, and HTML response templates from the GUI.
- **Scheduler** — Runs plugins on configurable intervals (1 min to 24 hrs) or manually on demand.
- **Activity Log** — Full audit trail of every email classified and action taken.
- **Business Hours** — Optionally restrict plugin execution to working hours only.

---

## Prerequisites

| Requirement | Detail |
|:---|:---|
| **Python** | 3.10 or higher |
| **OS** | Windows 10/11 (designed for Windows desktop) |
| **Microsoft 365** | An Entra ID (Azure AD) app registration with `Mail.ReadWrite` and `Mail.Send` permissions |
| **Anthropic API Key** | From [console.anthropic.com](https://console.anthropic.com) |

---

## Quick Start

### 1. Clone the repository

```bash
git clone https://github.com/faceless-truth/forge-one-app.git
cd forge-one-app
```

### 2. Launch (Windows)

Double-click **`launch.bat`** — it will:
- Create a virtual environment (first run only)
- Install all dependencies
- Start the application

### 3. Configure

1. Open the **Settings** tab
2. Enter your **Entra ID Tenant ID** and **Client ID**
3. Enter your **Anthropic API Key**
4. Click **Save All Settings**
5. Click **Sign in to Microsoft 365** (opens browser for OAuth)

### 4. Run

- Go to the **Dashboard** tab
- Click **▶ Start Scheduler** to begin automated monitoring
- Or go to **Plugins** → **▶ Run Now** to test manually

---

## Project Structure

```
mcs-desktop-agent/
├── app.py                          # Main GUI application (CustomTkinter)
├── config.py                       # SQLite database & configuration manager
├── graph_client.py                 # Microsoft Graph API client (MSAL OAuth2)
├── plugin_base.py                  # Abstract base class for all plugins
├── plugin_loader.py                # Plugin discovery, lifecycle & scheduler
├── plugins/
│   ├── __init__.py
│   ├── plugin_email_triage.py      # Email classification & auto-response
│   └── plugin_template.py          # Documented template for new plugins
├── requirements.txt
├── launch.bat                      # Windows launcher script
└── README.md
```

---

## Entra ID App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Entra ID** → **App registrations** → **New registration**
2. Name: `MC & S Desktop Agent`
3. Supported account types: **Single tenant**
4. Redirect URI: **Public client/native** → `http://localhost:8765`
5. Under **API permissions**, add:
   - `Microsoft Graph` → `Mail.ReadWrite` (Delegated)
   - `Microsoft Graph` → `Mail.Send` (Delegated)
6. Grant admin consent
7. Copy the **Tenant ID** and **Application (client) ID** into the app's Settings tab

---

## Building New Plugins

1. Copy `plugins/plugin_template.py` → `plugins/plugin_your_name.py`
2. Rename the class and fill in identity fields
3. Set `requires_graph` / `requires_claude` as needed
4. Set a `default_schedule`
5. Implement `run()` with your automation logic
6. Restart the app — it auto-discovers new plugins

### Plugin API

| Capability | Usage |
|:---|:---|
| Read emails | `context.graph.fetch_unread_emails(folder, max_count)` |
| Send email | `context.graph.send_email(to, subject, body)` |
| Create draft | `context.graph.create_draft(to, subject, body)` |
| Flag email | `context.graph.flag_email(message_id)` |
| AI reasoning | `context.claude.messages.create(...)` |
| Log to dashboard | `context.log("message")` |
| Check mode | `context.draft_mode` → `True` / `False` |
| Plugin settings | `self.get_plugin_setting(key)` / `self.set_plugin_setting(key, val)` |

### Plugin Ideas

- **NOA Workflow** — Detect ATO assessment notices, extract amounts, draft client emails
- **FuseSign Monitor** — Nudge clients with unsigned documents after X days
- **Meeting Prep Brief** — Compile client summary before calendar appointments
- **Debtor Follow-Up** — Automated overdue invoice reminders (1st, 2nd, 3rd notice)
- **ASIC Reminders** — Parse ASIC annual review notices, create calendar entries
- **Monthly Invoicing** — Surface retainer clients on the 1st of each month
- **Client Check-Ins** — Prompt 6-monthly outreach to companies/trusts

---

## Security Notes

- API keys and tokens are stored locally in `~/.mcs_email_automation/config.db` (SQLite)
- The MSAL token cache is stored at `~/.mcs_email_automation/.msal_cache.bin`
- Neither file is committed to Git (see `.gitignore`)
- OAuth uses the authorization code flow with a local redirect server on port 8765
- All email operations use delegated permissions (user context, not application-level)

---

## License

Private — MC & S Accountants. All rights reserved.
