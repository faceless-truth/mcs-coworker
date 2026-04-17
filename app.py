"""
MC & S Desktop Agent — Main Application
"""
import customtkinter as ctk
import threading
import time
import sys
import os
import json
import re
import shutil
from datetime import datetime
from tkinter import messagebox, filedialog
import tkinter as tk

try:
    import pystray
    from PIL import Image as PILImage
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

try:
    import anthropic as anthropic_lib
except ImportError:
    anthropic_lib = None

import config
from config import (
    init_db, get_setting, set_setting, get_rules, save_rule, delete_rule,
    get_recent_activity,
    get_links, save_link, delete_link,
    get_staff, save_staff, delete_staff,
    update_claude_model,
    update_claude_models,
)
from graph_client import GraphClient, MCS_TENANT_ID, MCS_CLIENT_ID
from plugin_loader import PluginLoader

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BRAND_BLUE    = "#1565C0"
BRAND_DARK    = "#0D47A1"
ACCENT_GREEN  = "#2E7D32"
ACCENT_AMBER  = "#E65100"
BG_LIGHT      = "#F5F7FA"
CARD_BG       = "#FFFFFF"
TEXT_PRIMARY   = "#1A1A2E"
TEXT_SECONDARY = "#4B5563"
TEXT_MUTED     = "#6B7280"
SUCCESS_FG     = "#2E7D32"
DRAFT_BG       = "#E3F2FD"
DRAFT_FG       = "#1565C0"
CHAT_DARK      = "#2D2D3F"

SCHEDULE_OPTIONS = [
    ("Manual", 0),
    ("1 min", 60),
    ("5 min", 300),
    ("15 min", 900),
    ("30 min", 1800),
    ("1 hr", 3600),
    ("4 hr", 14400),
    ("24 hr", 86400),
]

CHAT_SYSTEM_PROMPT = """\
You are an autonomous AI automation engineer built into MC & S CoWorker.
You operate in TWO tiers:

  TIER 1 — Template Builder (fast, safe, for common patterns)
  TIER 2 — Custom Plugin Writer (full Python, for anything complex)

Always try Tier 1 first. Escalate to Tier 2 only when the request
cannot be satisfied by any of the four templates.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TIER 1 — TEMPLATE BUILDER
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TYPE 1: SENDER_AUTO_REPLY
Use when: accountant wants a fixed automatic reply to emails \
from a specific email address.

TYPE 2: SENDER_AI_REPLY
Use when: accountant wants Claude to read emails from a specific \
sender and draft an intelligent contextual reply.

TYPE 3: KEYWORD_AUTO_REPLY
Use when: accountant wants to reply to emails containing \
certain words or phrases.

TYPE 4: FORWARD_AND_FILE
Use when: accountant wants emails from a specific sender \
forwarded to a staff member and saved in an Outlook folder.

Tier 1 workflow:
  1. Identify the type. Ask one clarifying question if unclear.
  2. Gather required fields one at a time.
  3. Confirm in plain English: "Does this sound right?"
  4. Build it — output ONLY the JSON tool call.
  5. Explain what was built and where to find it.

Tier 1 JSON formats:

SENDER_AUTO_REPLY:
{
  "tool": "create_plugin_from_template",
  "template_type": "SENDER_AUTO_REPLY",
  "plugin_name": "Auto Reply to Tony",
  "plugin_filename": "plugin_auto_reply_tony.py",
  "description": "Automatically replies to emails from tony@client.com",
  "sender_email": "tony@client.com",
  "reply_body_html": "<p>Thank you for your email...</p>",
  "draft_mode": true,
  "schedule_minutes": 5
}

SENDER_AI_REPLY:
{
  "tool": "create_plugin_from_template",
  "template_type": "SENDER_AI_REPLY",
  "plugin_name": "AI Reply to Elio",
  "plugin_filename": "plugin_ai_reply_elio.py",
  "description": "AI-drafted replies to emails from elio@example.com",
  "sender_email": "elio@example.com",
  "ai_instructions": "Always be brief and professional.",
  "draft_mode": true,
  "schedule_minutes": 5
}

KEYWORD_AUTO_REPLY:
{
  "tool": "create_plugin_from_template",
  "template_type": "KEYWORD_AUTO_REPLY",
  "plugin_name": "Pricing Enquiry Reply",
  "plugin_filename": "plugin_pricing_reply.py",
  "description": "Replies to emails asking about pricing",
  "keywords": ["how much", "price", "cost", "fee"],
  "reply_body_html": "<p>Thank you for your enquiry...</p>",
  "draft_mode": true,
  "schedule_minutes": 5
}

FORWARD_AND_FILE:
{
  "tool": "create_plugin_from_template",
  "template_type": "FORWARD_AND_FILE",
  "plugin_name": "Forward Tony Emails",
  "plugin_filename": "plugin_forward_tony.py",
  "description": "Forwards emails from tony@client.com to Harry",
  "sender_email": "tony@client.com",
  "forward_to": "harry@mcands.com.au",
  "forward_note": "FYI - forwarded automatically by CoWorker",
  "folder_name": "Tony Phillips",
  "schedule_minutes": 5
}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TIER 2 — CUSTOM PLUGIN WRITER
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Use Tier 2 when the request requires:
  - Multi-step workflows (e.g. fetch email → look up XPM → send Teams alert)
  - Semantic memory (remembering past client interactions)
  - XPM, FuseSign, or Teams integration
  - Custom scheduling logic or event-driven triggers
  - Database tables, file processing, or complex business logic
  - Anything that cannot be expressed as one of the four Tier 1 templates

Tier 2 workflow:
  1. Confirm the request needs Tier 2 (explain why briefly).
  2. Ask clarifying questions until you have enough to write the plugin.
  3. Confirm the plan: "Here is what I will build — does this sound right?"
  4. Write the complete plugin as a single Python file.
  5. Output the JSON tool call below.
  6. Explain what was built, what triggers it, and how to enable it.

Tier 2 JSON format:
{
  "tool": "create_plugin",
  "filename": "plugin_example.py",
  "plugin_name": "Example Plugin",
  "description": "One-sentence description",
  "code": "<complete Python plugin code as a string>"
}

=== PLUGIN ARCHITECTURE (TIER 2 REFERENCE) ===

Every plugin must follow this exact structure:

```python
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule


class MyPlugin(AgentPlugin):
    name = "My Plugin"          # Display name in the UI
    description = "What it does"  # Short description
    detail = "Longer detail"    # Shown in plugin card
    version = "1.0.0"
    icon = "\U0001f916"          # Emoji icon
    author = "CoWorker AI"
    requires_graph = True       # Set False if no email/Outlook needed
    requires_claude = True      # Set False if no AI needed
    default_schedule = Schedule.every_minutes(5)  # or .every_hours(N), .daily_at(H), .manual_only()

    def load(self, context: PluginContext) -> bool:
        # Called once on startup. Return True if ready, False to disable.
        return bool(context.graph)  # and bool(context.claude) if needed

    def run(self, context: PluginContext) -> PluginResult:
        # Main logic. Must return PluginResult.
        # ...
        return PluginResult(
            success=True,
            summary="What happened",
            actions_taken=0,
            drafts_created=0,
            items_skipped=0,
        )
```

=== CONTEXT SERVICES AVAILABLE IN TIER 2 ===

context.graph          — Microsoft 365 / Outlook
  .fetch_unread_emails(folder, limit)     → list of email dicts
  .send_email(to, subject, body, reply_to_id=None)
  .create_draft(to, subject, body, reply_to_id=None)
  .create_draft_with_inline_image(to, subject, body, img_path, cid, reply_to_id)
  .send_email_with_inline_image(to, subject, body, img_path, cid, reply_to_id)
  .mark_as_read(message_id)
  .move_email(message_id, folder_name)
  .flag_email(message_id)
  .get_signature_image_path() → str | None

context.claude         — Haiku (fast) — for triage, drafting, classification
context.claude_fast    — same as context.claude
context.claude_reason  — Sonnet (powerful) — for complex analysis, decisions
  Usage: context.claude.messages.create(
      model=self.get_claude_model(),   # always use this, never hardcode
      max_tokens=1000,
      messages=[{"role": "user", "content": prompt}]
  )
  For complex reasoning: context.claude_reason.messages.create(
      model=self.get_claude_model_reasoning(),
      ...
  )

context.memory         — Semantic vector memory (ChromaDB)
  .store(collection, doc_id, text, metadata={})  — store a memory
  .search(collection, query, n_results=5)         — semantic search
  .store_client_interaction(client_email, interaction_type, summary, metadata={})
  .get_client_context(client_email, query, n_results=3) → str
  .store_lesson(lesson, source)                   — store a learned lesson
  .store_document(doc_id, text, metadata={})      — store a processed document
  Always check: if context.memory is None: (graceful degradation)

context.event_bus      — Publish/subscribe event system
  .publish(event_type, payload={}, source="")     — fire an event
  .subscribe(event_type, handler)                 — listen for events
  .emit(event_type, payload={}, source="")        — alias for publish
  Standard event types: "email.received", "email.sent", "noa.processed",
  "outreach.drafted", "plugin.run.complete", "plugin.run.failed", "heartbeat.tick"

context.gateway        — External platform integrations
  .xpm.list_clients(search="")                    — search XPM clients
  .xpm.get_client_by_email(email)                 — find client by email
  .xpm.list_jobs(client_id, status="In Progress") — list jobs
  .xpm.update_job_status(job_id, status)          — update job
  .xpm.add_client_note(client_id, note)           — add note to client
  .fusesign.create_envelope(template_id, recipients, subject, message)
  .fusesign.get_envelope_status(envelope_id)
  .fusesign.list_envelopes(status="pending")
  .teams.send_message(text, title="")             — send Teams notification
  .teams.send_alert(title, body, urgent=False)    — send alert card
  Always check: if context.gateway.is_available("xpm"): (graceful degradation)

context.log(message)   — Write to the activity log
context.notify(subject, body)  — Send notification email to staff
context.settings       — dict of all app settings
context.draft_mode     — bool: True = create drafts, False = send immediately

=== TIER 2 CODING RULES ===
1. Always use self.get_claude_model() — never hardcode a model string.
2. Use self.get_claude_model_reasoning() for complex analysis.
3. Always check context.memory is not None before using memory.
4. Always check context.gateway.is_available("xpm") before using XPM.
5. Always check context.gateway.is_available("teams") before Teams.
6. Always check context.gateway.is_available("fusesign") before FuseSign.
7. Always include a load() method that returns bool.
8. Always return a PluginResult — never return None.
9. Use context.draft_mode to decide send vs draft — never hardcode.
10. Use get_signature_image_path() when sending/drafting emails.
11. plugin_filename must start with plugin_ and end with .py.
12. Use default_schedule = Schedule.every_minutes(N) for scheduling.
13. For plugins that should only run manually: Schedule.manual_only()
14. Wrap all external calls in try/except and log errors with context.log().
15. Store meaningful outcomes in context.memory for future context.
16. Publish events to context.event_bus after significant actions.

=== SCHEDULE OPTIONS ===
Schedule.every_minutes(1)    # every 1 minute
Schedule.every_minutes(5)    # every 5 minutes
Schedule.every_minutes(15)   # every 15 minutes
Schedule.every_minutes(30)   # every 30 minutes
Schedule.every_hours(1)      # every hour
Schedule.every_hours(4)      # every 4 hours
Schedule.daily_at(8)         # once daily at 8am
Schedule.manual_only()       # only runs when manually triggered

=== PLUGINRESULT FIELDS ===
PluginResult(
    success=True,
    summary="Human-readable summary",
    actions_taken=0,    # emails sent, records updated, etc.
    drafts_created=0,   # email drafts created
    items_skipped=0,    # items that were skipped
)

=== UNIVERSAL TOOLS ===

create_email_rule (simple keyword auto-responses only):
{
  "tool": "create_email_rule",
  "category": "CATEGORY_NAME",
  "keywords": "comma,separated,keywords",
  "subject_template": "Re: Subject here",
  "body_template": "<html>...</html>",
  "enabled": 1
}

update_setting:
{
  "tool": "update_setting",
  "key": "setting_key",
  "value": "value"
}

clarify:
{
  "tool": "clarify",
  "question": "..."
}

=== UNIVERSAL RULES ===
1. Never suggest Outlook rules or external tools.
2. Always confirm before building.
3. Always use draft_mode=true unless accountant explicitly asks for automatic sending.
4. Use clarify tool ONLY during requirements gathering.
5. plugin_filename must always start with plugin_ and end with .py.
6. plugin_name should be descriptive and derived from the request.
7. If a request is completely outside what CoWorker can do, say so clearly.
"""

# ────────────────────────────────────────────────────────────────────────
# Plugin Templates — pre-validated, working code filled by parameters
# ────────────────────────────────────────────────────────────────────────

SENDER_AUTO_REPLY_TEMPLATE = '''\
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule


class {class_name}(AgentPlugin):
    name = "{plugin_name}"
    description = "{description}"
    detail = "{description}"
    version = "1.0.0"
    icon = "\\U0001f4e7"
    author = "CoWorker AI"
    requires_graph = True
    requires_claude = False
    default_schedule = Schedule.every_minutes({schedule_minutes})

    _processed_ids: set

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        return bool(context.graph)

    @classmethod
    def email_templates_schema(cls):
        return [
            {{
                "key": "reply_body",
                "label": "Reply Email Body (HTML)",
                "default": {reply_body_repr},
                "type": "textarea"
            }}
        ]

    def run(self, context: PluginContext) -> PluginResult:
        emails = context.graph.fetch_unread_emails("Inbox", 25)
        target = [
            e for e in emails
            if e.get('from', {{}}).get('emailAddress', {{}})
            .get('address', '').lower() == '{sender_email}'
            and e.get('id') not in self._processed_ids
        ]

        if not target:
            return PluginResult(success=True,
                summary="No emails from {sender_email}.",
                items_skipped=0)

        reply_body = self.get_email_template(
            "reply_body", {reply_body_repr}
        )
        sig_image_path = context.graph.get_signature_image_path()
        actions = 0
        drafts = 0

        for email in target:
            msg_id = email.get('id')
            sender = email.get('from', {{}}).get(
                'emailAddress', {{}}).get('address', '')
            subject = email.get('subject', 'No Subject')
            reply_subject = f"Re: {{subject}}"

            try:
                if {draft_mode_bool}:
                    if sig_image_path:
                        context.graph.create_draft_with_inline_image(
                            sender, reply_subject, reply_body,
                            sig_image_path, "signature_image", msg_id
                        )
                    else:
                        context.graph.create_draft(
                            sender, reply_subject, reply_body, msg_id
                        )
                    context.log(f"Draft created for {{sender}}")
                    drafts += 1
                else:
                    if sig_image_path:
                        context.graph.send_email_with_inline_image(
                            sender, reply_subject, reply_body,
                            sig_image_path, "signature_image", msg_id
                        )
                    else:
                        context.graph.send_email(
                            sender, reply_subject, reply_body, msg_id
                        )
                    context.log(f"Reply sent to {{sender}}")
                    actions += 1

                context.graph.mark_as_read(msg_id)
                self._processed_ids.add(msg_id)

            except Exception as e:
                context.log(f"Error processing email: {{e}}")

        return PluginResult(
            success=True,
            summary=f"{{actions}} sent, {{drafts}} drafted.",
            actions_taken=actions,
            drafts_created=drafts
        )
'''

SENDER_AI_REPLY_TEMPLATE = '''\
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule


class {class_name}(AgentPlugin):
    name = "{plugin_name}"
    description = "{description}"
    detail = "{description}"
    version = "1.0.0"
    icon = "\\U0001f916"
    author = "CoWorker AI"
    requires_graph = True
    requires_claude = True
    default_schedule = Schedule.every_minutes({schedule_minutes})

    _processed_ids: set

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        return bool(context.graph) and bool(context.claude)

    @classmethod
    def email_templates_schema(cls):
        return [
            {{
                "key": "ai_instructions",
                "label": "AI Drafting Instructions",
                "default": {ai_instructions_repr},
                "type": "prompt"
            }}
        ]

    def run(self, context: PluginContext) -> PluginResult:
        emails = context.graph.fetch_unread_emails("Inbox", 25)
        target = [
            e for e in emails
            if e.get('from', {{}}).get('emailAddress', {{}})
            .get('address', '').lower() == '{sender_email}'
            and e.get('id') not in self._processed_ids
        ]

        if not target:
            return PluginResult(success=True,
                summary="No emails from {sender_email}.",
                items_skipped=0)

        ai_instructions = self.get_email_template(
            "ai_instructions", {ai_instructions_repr}
        )
        sig_image_path = context.graph.get_signature_image_path()
        drafts = 0

        for email in target:
            msg_id = email.get('id')
            sender = email.get('from', {{}}).get(
                'emailAddress', {{}}).get('address', '')
            subject = email.get('subject', 'No Subject')
            body = email.get('body', {{}}).get('content', '')

            import re
            body_plain = re.sub(r'<[^>]+>', ' ', body)
            body_plain = re.sub(r'\\s+', ' ', body_plain).strip()

            prompt = (
                "You are a professional accountant's assistant "
                "at MC & S Pty Ltd.\\n\\n"
                + ai_instructions + "\\n\\n"
                "Draft a professional email reply to this email:\\n"
                f"From: {{sender}}\\n"
                f"Subject: {{subject}}\\n"
                f"Body: {{body_plain[:1500]}}\\n\\n"
                "Reply in HTML format only. Be professional, helpful "
                "and concise. Start with \\"Hi,\\" and end with "
                "\\"Kind regards,\\"\\n"
                "Do not include a subject line in your response."
            )

            try:
                response = context.claude.messages.create(
                    model=self.get_claude_model(),
                    max_tokens=1000,
                    messages=[{{"role": "user", "content": prompt}}]
                )
                reply_body = response.content[0].text.strip()
                reply_subject = f"Re: {{subject}}"

                if sig_image_path:
                    context.graph.create_draft_with_inline_image(
                        sender, reply_subject, reply_body,
                        sig_image_path, "signature_image", msg_id
                    )
                else:
                    context.graph.create_draft(
                        sender, reply_subject, reply_body, msg_id
                    )

                context.log(f"AI draft created for {{sender}}")
                context.graph.mark_as_read(msg_id)
                self._processed_ids.add(msg_id)
                drafts += 1

            except Exception as e:
                context.log(f"Error processing email: {{e}}")

        return PluginResult(
            success=True,
            summary=f"{{drafts}} AI draft(s) created.",
            drafts_created=drafts
        )
'''

FORWARD_AND_FILE_TEMPLATE = '''\
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule


class {class_name}(AgentPlugin):
    name = "{plugin_name}"
    description = "{description}"
    detail = "{description}"
    version = "1.0.0"
    icon = "\\U0001f4c1"
    author = "CoWorker AI"
    requires_graph = True
    requires_claude = False
    default_schedule = Schedule.every_minutes({schedule_minutes})

    _processed_ids: set

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        return bool(context.graph)

    def run(self, context: PluginContext) -> PluginResult:
        emails = context.graph.fetch_unread_emails("Inbox", 25)
        target = [
            e for e in emails
            if e.get('from', {{}}).get('emailAddress', {{}})
            .get('address', '').lower() == '{sender_email}'
            and e.get('id') not in self._processed_ids
        ]

        if not target:
            return PluginResult(success=True,
                summary="No emails from {sender_email}.",
                items_skipped=0)

        actions = 0
        for email in target:
            msg_id = email.get('id')
            sender = email.get('from', {{}}).get(
                'emailAddress', {{}}).get('address', '')
            subject = email.get('subject', 'No Subject')
            body = email.get('body', {{}}).get('content', '')

            try:
                forward_body = (
                    "<p>{forward_note}</p>"
                    "<hr>"
                    f"<p><strong>From:</strong> {{sender}}<br>"
                    f"<strong>Subject:</strong> {{subject}}</p>"
                    f"{{body}}"
                )
                context.graph.send_email(
                    '{forward_to}',
                    f'Fwd: {{subject}}',
                    forward_body
                )
                context.graph.move_email(msg_id, '{folder_name}')
                context.graph.mark_as_read(msg_id)
                self._processed_ids.add(msg_id)
                context.log(
                    f"Forwarded {{sender}} to {forward_to} "
                    f"and moved to {folder_name}"
                )
                actions += 1
            except Exception as e:
                context.log(f"Error: {{e}}")

        return PluginResult(
            success=True,
            summary=f"{{actions}} email(s) forwarded and filed.",
            actions_taken=actions
        )
'''

KEYWORD_AUTO_REPLY_TEMPLATE = '''\
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule


class {class_name}(AgentPlugin):
    name = "{plugin_name}"
    description = "{description}"
    detail = "{description}"
    version = "1.0.0"
    icon = "\\U0001f511"
    author = "CoWorker AI"
    requires_graph = True
    requires_claude = False
    default_schedule = Schedule.every_minutes({schedule_minutes})

    _processed_ids: set
    KEYWORDS = {keywords_list}

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        return bool(context.graph)

    @classmethod
    def email_templates_schema(cls):
        return [
            {{
                "key": "reply_body",
                "label": "Reply Email Body (HTML)",
                "default": {reply_body_repr},
                "type": "textarea"
            }}
        ]

    def run(self, context: PluginContext) -> PluginResult:
        emails = context.graph.fetch_unread_emails("Inbox", 25)

        def has_keyword(email):
            subject = email.get('subject', '').lower()
            body = email.get('body', {{}}).get('content', '').lower()
            return any(kw.lower() in subject or kw.lower() in body
                      for kw in self.KEYWORDS)

        target = [
            e for e in emails
            if has_keyword(e)
            and e.get('id') not in self._processed_ids
        ]

        if not target:
            return PluginResult(success=True,
                summary="No matching emails found.",
                items_skipped=0)

        reply_body = self.get_email_template(
            "reply_body", {reply_body_repr}
        )
        sig_image_path = context.graph.get_signature_image_path()
        actions = 0
        drafts = 0

        for email in target:
            msg_id = email.get('id')
            sender = email.get('from', {{}}).get(
                'emailAddress', {{}}).get('address', '')
            subject = email.get('subject', 'No Subject')

            try:
                if {draft_mode_bool}:
                    if sig_image_path:
                        context.graph.create_draft_with_inline_image(
                            sender, f"Re: {{subject}}", reply_body,
                            sig_image_path, "signature_image", msg_id
                        )
                    else:
                        context.graph.create_draft(
                            sender, f"Re: {{subject}}",
                            reply_body, msg_id
                        )
                    drafts += 1
                else:
                    if sig_image_path:
                        context.graph.send_email_with_inline_image(
                            sender, f"Re: {{subject}}", reply_body,
                            sig_image_path, "signature_image", msg_id
                        )
                    else:
                        context.graph.send_email(
                            sender, f"Re: {{subject}}",
                            reply_body, msg_id
                        )
                    actions += 1

                context.graph.mark_as_read(msg_id)
                self._processed_ids.add(msg_id)

            except Exception as e:
                context.log(f"Error: {{e}}")

        return PluginResult(
            success=True,
            summary=f"{{actions}} sent, {{drafts}} drafted.",
            actions_taken=actions,
            drafts_created=drafts
        )
'''


def _seconds_to_schedule_label(seconds: int) -> str:
    """Convert schedule seconds to a dropdown label."""
    for label, secs in SCHEDULE_OPTIONS:
        if secs == seconds:
            return label
    if seconds <= 0:
        return "Manual"
    if seconds < 3600:
        return f"{seconds // 60} min"
    return f"{seconds // 3600} hr"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        init_db()
        self.title("MC & S — CoWorker")
        self.geometry("1220x760")
        self.minsize(1000, 660)
        self.configure(fg_color=BG_LIGHT)

        self._loader = PluginLoader(log_callback=self._log)
        self._loader.on_run_complete(self._on_plugin_run_complete)
        self._graph: GraphClient | None = None
        self._session_actions = 0
        self._session_drafts  = 0

        # Check if first-run setup is needed
        if get_setting("user_setup_complete") == "1":
            self._launch_main_ui()
        else:
            self._show_setup_wizard()

    def _launch_main_ui(self):
        """Build and show the main application interface."""
        # Destroy wizard frame if it exists
        if hasattr(self, '_wizard_frame'):
            self._wizard_frame.destroy()

        self._tray_icon = None
        self._tray_thread = None

        self._build_header()
        self._build_nav()
        self._build_content_area()
        self._try_restore_session()
        self.after(200, self._initialise_plugins)
        self._show_page("dashboard")

        # Auto-start scheduler after a short delay to allow plugin init
        self.after(1500, self._auto_start_scheduler)

        # Auto-detect latest Claude model in background
        self.after(3000, lambda: threading.Thread(
            target=self._auto_update_claude_model,
            daemon=True).start()
        )

        # Override window close to minimise to tray instead of quitting
        self.protocol("WM_DELETE_WINDOW", self._on_close_requested)

    # ────────────────────────────────────────────────────────────────────────
    # First-Run Onboarding Wizard (3 steps)
    # ────────────────────────────────────────────────────────────────────────

    def _show_setup_wizard(self):
        """Display a 3-step onboarding wizard for first-time users."""
        self._wizard_step = 0
        self._wizard_frame = ctk.CTkFrame(self, fg_color=BG_LIGHT, corner_radius=0)
        self._wizard_frame.pack(fill="both", expand=True)

        # Header bar
        hdr = ctk.CTkFrame(self._wizard_frame, height=64, fg_color=BRAND_BLUE, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="MC & S  CoWorker",
                     font=ctk.CTkFont(family="Arial", size=20, weight="bold"),
                     text_color="white").pack(side="left", padx=20, pady=16)

        # Progress indicator
        self._wizard_progress_frame = ctk.CTkFrame(self._wizard_frame, fg_color=BG_LIGHT, height=50)
        self._wizard_progress_frame.pack(fill="x", padx=60, pady=(20, 0))
        self._wizard_progress_frame.pack_propagate(False)
        self._wizard_step_labels = []
        steps = ["Welcome", "Your Details", "Connect Email"]
        for i, step_name in enumerate(steps):
            lbl = ctk.CTkLabel(self._wizard_progress_frame,
                               text=f"  Step {i+1}: {step_name}  ",
                               font=ctk.CTkFont(size=13, weight="bold" if i == 0 else "normal"),
                               text_color=BRAND_BLUE if i == 0 else TEXT_MUTED)
            lbl.pack(side="left", padx=12)
            self._wizard_step_labels.append(lbl)

        # Content area for wizard steps
        self._wizard_content = ctk.CTkFrame(self._wizard_frame, fg_color=BG_LIGHT)
        self._wizard_content.pack(fill="both", expand=True, padx=60, pady=(10, 20))

        # Button row
        self._wizard_btn_frame = ctk.CTkFrame(self._wizard_frame, fg_color=BG_LIGHT, height=60)
        self._wizard_btn_frame.pack(fill="x", padx=60, pady=(0, 30))
        self._wizard_btn_frame.pack_propagate(False)

        self._wizard_back_btn = ctk.CTkButton(
            self._wizard_btn_frame, text="Back", width=120, height=42,
            fg_color="transparent", hover_color="#E3F2FD",
            text_color=BRAND_BLUE, border_width=1, border_color=BRAND_BLUE,
            font=ctk.CTkFont(size=14), command=self._wizard_back)
        self._wizard_back_btn.pack(side="left")

        self._wizard_next_btn = ctk.CTkButton(
            self._wizard_btn_frame, text="Get Started", width=180, height=42,
            fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
            font=ctk.CTkFont(size=14, weight="bold"), command=self._wizard_next)
        self._wizard_next_btn.pack(side="right")

        self._wizard_show_step(0)

    def _wizard_update_progress(self, step):
        for i, lbl in enumerate(self._wizard_step_labels):
            if i < step:
                lbl.configure(text_color=ACCENT_GREEN,
                              font=ctk.CTkFont(size=13, weight="normal"))
            elif i == step:
                lbl.configure(text_color=BRAND_BLUE,
                              font=ctk.CTkFont(size=13, weight="bold"))
            else:
                lbl.configure(text_color=TEXT_MUTED,
                              font=ctk.CTkFont(size=13, weight="normal"))

    def _wizard_show_step(self, step):
        self._wizard_step = step
        self._wizard_update_progress(step)

        # Clear content area
        for w in self._wizard_content.winfo_children():
            w.destroy()

        # Update button states
        if step == 0:
            self._wizard_back_btn.configure(state="disabled")
            self._wizard_next_btn.configure(text="Get Started", state="normal")
        elif step == 1:
            self._wizard_back_btn.configure(state="normal")
            self._wizard_next_btn.configure(text="Next", state="normal")
        elif step == 2:
            self._wizard_back_btn.configure(state="normal")
            self._wizard_next_btn.configure(text="Finish Setup", state="disabled")

        if step == 0:
            self._wizard_step_welcome()
        elif step == 1:
            self._wizard_step_details()
        elif step == 2:
            self._wizard_step_connect()

    def _wizard_step_welcome(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=10)

        ctk.CTkLabel(card, text="Welcome to MC&S CoWorker",
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(50, 8))
        ctk.CTkLabel(card, text="Let's get you set up in 3 easy steps.",
                     font=ctk.CTkFont(size=14), text_color=TEXT_MUTED).pack(pady=(0, 30))

        features = [
            ("📨", "Email Triage", "Automatically classifies incoming emails and drafts replies"),
            ("✏️", "Draft Mode", "Emails are drafted for your review before sending"),
            ("⚡", "Always On", "Runs in the background while you work"),
        ]
        for icon, title, desc in features:
            row = ctk.CTkFrame(card, fg_color="#F0F4FF", corner_radius=10)
            row.pack(fill="x", padx=80, pady=3)
            ctk.CTkLabel(row, text=icon, font=ctk.CTkFont(size=20)).pack(side="left", padx=(16, 10), pady=10)
            text_frame = ctk.CTkFrame(row, fg_color="transparent")
            text_frame.pack(side="left", fill="x", expand=True, pady=6)
            ctk.CTkLabel(text_frame, text=title,
                         font=ctk.CTkFont(size=14, weight="bold"),
                         text_color=TEXT_PRIMARY, anchor="w").pack(anchor="w")
            ctk.CTkLabel(text_frame, text=desc,
                         font=ctk.CTkFont(size=12),
                         text_color=TEXT_MUTED, anchor="w").pack(anchor="w")

    def _wizard_step_details(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=10)

        ctk.CTkLabel(card, text="Your Details",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(40, 4))
        ctk.CTkLabel(card, text="Tell us a bit about you so we can personalise your experience.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED).pack(pady=(0, 24))

        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(padx=100, fill="x")

        ctk.CTkLabel(form, text="Your Name",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_name = ctk.CTkEntry(form, height=40,
                                         font=ctk.CTkFont(size=14),
                                         placeholder_text="e.g. Sarah Chen")
        saved_name = get_setting("user_name")
        if saved_name:
            self._wizard_name.insert(0, saved_name)
        self._wizard_name.pack(fill="x", pady=(0, 16))

        ctk.CTkLabel(form, text="Your Firm Name",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_firm = ctk.CTkEntry(form, height=40,
                                         font=ctk.CTkFont(size=14),
                                         placeholder_text="e.g. Smith & Associates Accounting")
        saved_firm = get_setting("user_firm")
        if saved_firm:
            self._wizard_firm.insert(0, saved_firm)
        self._wizard_firm.pack(fill="x", pady=(0, 16))

        ctk.CTkLabel(form, text="Your Email Address",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))
        self._wizard_email = ctk.CTkEntry(form, height=40,
                                          font=ctk.CTkFont(size=14),
                                          placeholder_text="e.g. sarah@smithaccounting.com.au")
        saved_email = get_setting("user_email")
        if saved_email:
            self._wizard_email.insert(0, saved_email)
        self._wizard_email.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(form, text="This is the mailbox CoWorker will monitor for incoming emails.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(anchor="w")

    def _wizard_step_connect(self):
        card = ctk.CTkFrame(self._wizard_content, fg_color=CARD_BG, corner_radius=16)
        card.pack(fill="both", expand=True, padx=40, pady=10)

        ctk.CTkLabel(card, text="Connect your Microsoft 365 account",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(pady=(50, 8))
        ctk.CTkLabel(card, text="Click below to sign in. We'll only access your inbox to help\nyou respond to emails.",
                     font=ctk.CTkFont(size=13), text_color=TEXT_MUTED,
                     wraplength=500, justify="center").pack(pady=(0, 30))

        self._wizard_signin_btn = ctk.CTkButton(
            card, text="Sign in to Microsoft 365",
            width=300, height=48, fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._wizard_do_signin)
        self._wizard_signin_btn.pack()

        # Spinner / status area
        self._wizard_signin_status = ctk.CTkLabel(
            card, text="", font=ctk.CTkFont(size=13), text_color=TEXT_MUTED)
        self._wizard_signin_status.pack(pady=(12, 0))

        # Success area (hidden initially)
        self._wizard_success_frame = ctk.CTkFrame(card, fg_color="#E8F5E9", corner_radius=10)

    def _wizard_do_signin(self):
        self._graph = GraphClient()
        self._wizard_signin_status.configure(text="Opening browser... please sign in.",
                                             text_color=TEXT_MUTED)
        self._wizard_signin_btn.configure(state="disabled")

        def callback(success, error):
            if success:
                self.after(0, self._wizard_signin_success)
            else:
                self.after(0, lambda: self._wizard_signin_fail(str(error)))

        self._graph.authenticate(callback=callback)

    def _wizard_signin_success(self):
        # Get user info
        email = get_setting("user_email") or ""
        try:
            info = self._graph.get_user_info()
            email = info.get("mail", info.get("userPrincipalName", email))
        except Exception:
            pass

        self._wizard_signin_status.configure(text="", text_color=TEXT_MUTED)
        self._wizard_signin_btn.configure(
            text="Connected", fg_color=ACCENT_GREEN, state="disabled")

        # Show success indicator
        self._wizard_success_frame.pack(fill="x", padx=80, pady=(16, 0))
        ctk.CTkLabel(self._wizard_success_frame,
                     text=f"Connected as {email}",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=ACCENT_GREEN).pack(padx=20, pady=12)

        # Enable the Finish Setup button
        self._wizard_next_btn.configure(state="normal")

    def _wizard_signin_fail(self, error):
        self._wizard_signin_status.configure(
            text=f"Sign in failed: {error}", text_color="#C62828")
        self._wizard_signin_btn.configure(state="normal")

    def _wizard_next(self):
        if self._wizard_step == 0:
            self._wizard_show_step(1)

        elif self._wizard_step == 1:
            name = self._wizard_name.get().strip()
            firm = self._wizard_firm.get().strip()
            email = self._wizard_email.get().strip()
            if not name:
                messagebox.showerror("Required", "Please enter your name.")
                return
            if not firm:
                messagebox.showerror("Required", "Please enter your firm name.")
                return
            if not email:
                messagebox.showerror("Required", "Please enter your email address.")
                return
            # Save details
            set_setting("user_name", name)
            set_setting("user_firm", firm)
            set_setting("user_email", email)
            set_setting("ms_account_email", email)
            set_setting("practice_name", firm)
            self._wizard_show_step(2)

        elif self._wizard_step == 2:
            # Finish — mark setup complete and launch
            set_setting("user_setup_complete", "1")
            self._launch_main_ui()

    def _wizard_back(self):
        if self._wizard_step > 0:
            self._wizard_show_step(self._wizard_step - 1)

    def _initialise_plugins(self):
        self._loader.discover()
        self._loader.load_all()

    # ────────────────────────────────────────────────────────────────────────
    # Layout
    # ────────────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self, height=64, fg_color=BRAND_BLUE, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="MC & S  CoWorker",
                     font=ctk.CTkFont(family="Arial", size=20, weight="bold"),
                     text_color="white").pack(side="left", padx=20, pady=16)

        # Auth status
        self.auth_status_label = ctk.CTkLabel(hdr, text="",
                                              text_color="#66BB6A", font=ctk.CTkFont(size=12))
        self.auth_status_label.pack(side="right", padx=20)

        # User name
        user_name = get_setting("user_name")
        if user_name:
            ctk.CTkLabel(hdr, text=user_name,
                         text_color="white", font=ctk.CTkFont(size=13)).pack(side="right", padx=8)

        # Scheduler status
        self.scheduler_label = ctk.CTkLabel(hdr, text="",
                                            text_color="#CFD8DC", font=ctk.CTkFont(size=12))
        self.scheduler_label.pack(side="right", padx=12)

    def _build_nav(self):
        nav = ctk.CTkFrame(self, width=210, fg_color=BRAND_DARK, corner_radius=0)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        self._nav_btns = {}
        pages = [
            ("dashboard", "  Dashboard"),
            ("plugins",   "  Plugins"),
            ("approvals", "  Approvals"),
            ("rules",     "  Email Rules"),
            ("staff",     "  Staff & Notify"),
            ("settings",  "  Settings"),
            ("chat",      "  Chat"),
            ("activity",  "  Activity Log"),
            ("usage",     "  AI Usage"),
        ]
        ctk.CTkLabel(nav, text="", height=10).pack()
        for key, label in pages:
            btn = ctk.CTkButton(nav, text=label, width=200, height=42,
                                fg_color="transparent", hover_color="#1565C0",
                                text_color="white", anchor="w", font=ctk.CTkFont(size=13),
                                command=lambda k=key: self._show_page(k))
            btn.pack(pady=2, padx=6)
            self._nav_btns[key] = btn
        # Badge label for pending approvals count
        self._approvals_badge = ctk.CTkLabel(nav, text="", font=ctk.CTkFont(size=11),
                                              text_color="#FF5252", anchor="w")
        self._approvals_badge.pack(pady=(0, 4), padx=(70, 6))

    def _build_content_area(self):
        self.content = ctk.CTkFrame(self, fg_color=BG_LIGHT, corner_radius=0)
        self.content.pack(side="left", fill="both", expand=True)
        self._pages = {}
        self._build_dashboard()
        self._build_plugins_page()
        self._build_approvals_page()
        self._build_rules_page()
        self._build_staff_page()
        self._build_settings_page()
        self._build_chat_page()
        self._build_activity_page()
        self._build_usage_page()

    # ────────────────────────────────────────────────────────────────────────
    # Dashboard
    # ────────────────────────────────────────────────────────────────────────

    def _build_dashboard(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["dashboard"] = page

        ctk.CTkLabel(page, text="Dashboard",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(24, 2))
        ctk.CTkLabel(page, text="Your MC & S CoWorker — automating email triage while you work.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(0, 16))

        card_row = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        card_row.pack(fill="x", padx=28)
        self._stat_labels = {}
        for key, title, default, icon in [
            ("scheduler", "Status",           "Starting...", "⏱"),
            ("actions",   "Emails Processed",  "0",           "⚡"),
            ("drafts",    "Drafts Created",    "0",           "📝"),
        ]:
            f = ctk.CTkFrame(card_row, fg_color=CARD_BG, corner_radius=12)
            f.pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkLabel(f, text=icon, font=ctk.CTkFont(size=28)).pack(pady=(16, 4))
            lbl = ctk.CTkLabel(f, text=default,
                               font=ctk.CTkFont(size=22, weight="bold"), text_color=TEXT_PRIMARY)
            lbl.pack()
            ctk.CTkLabel(f, text=title, text_color=TEXT_MUTED, font=ctk.CTkFont(size=12)).pack(pady=(0, 16))
            self._stat_labels[key] = lbl

        ctk.CTkLabel(page, text="Live Log", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(16, 0))
        self.log_box = ctk.CTkTextbox(page, height=300,
                                      font=ctk.CTkFont(family="Courier", size=12),
                                      fg_color="#1A1A2E", text_color="#E0E0E0", corner_radius=8)
        self.log_box.pack(fill="both", expand=True, padx=28, pady=(6, 20))
        self.log_box.configure(state="disabled")

    # ────────────────────────────────────────────────────────────────────────
    # Scheduler (auto-start)
    # ────────────────────────────────────────────────────────────────────────

    def _auto_start_scheduler(self):
        """Auto-start the scheduler silently after setup is complete."""
        if not self._graph or not self._graph.is_authenticated():
            self._log("Waiting for Microsoft 365 connection...")
            self.after(3000, self._auto_start_scheduler)
            return

        self._loader.set_graph(self._graph)
        self._loader.set_claude()
        self._loader.load_all()
        self._loader.start_scheduler()

        self._stat_labels["scheduler"].configure(text="Running")
        self.scheduler_label.configure(text="Running", text_color="#A5D6A7")
        self._log("Scheduler started. Email triage running automatically.")

    def _auto_update_claude_model(self):
        """Query Anthropic API and cache both the fast (Haiku) and reasoning (Sonnet) models."""
        api_key = get_setting("anthropic_api_key")
        if not api_key:
            return
        try:
            models = update_claude_models(api_key)
            fast = models.get("fast", "unknown")
            reasoning = models.get("reasoning", "unknown")
            self.after(0, lambda: self._log(
                f"[Claude] Fast model: {fast}  |  Reasoning model: {reasoning}"
            ))
        except Exception:
            pass

    # ────────────────────────────────────────────────────────────────────────
    # Approvals page
    # ────────────────────────────────────────────────────────────────────────

    def _build_approvals_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["approvals"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Approvals",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="Refresh", width=100, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._refresh_approvals).pack(side="right")

        ctk.CTkLabel(page,
                     text="Actions queued for your review. Approve to execute, reject to discard.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(
                         anchor="w", padx=28, pady=(4, 4))

        # Summary bar
        self._approvals_summary_label = ctk.CTkLabel(
            page, text="", font=ctk.CTkFont(size=12), text_color=TEXT_MUTED)
        self._approvals_summary_label.pack(anchor="w", padx=28, pady=(0, 8))

        # Filter row
        filter_row = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        filter_row.pack(fill="x", padx=28, pady=(0, 8))
        ctk.CTkLabel(filter_row, text="Show:",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(side="left")
        self._approvals_filter = ctk.StringVar(value="pending")
        for label, val in [("Pending", "pending"), ("All", "all"),
                           ("Approved", "approved"), ("Rejected", "rejected")]:
            ctk.CTkRadioButton(filter_row, text=label, variable=self._approvals_filter,
                               value=val, font=ctk.CTkFont(size=12),
                               command=self._refresh_approvals).pack(
                                   side="left", padx=8)

        self._approvals_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self._approvals_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))

        self.after(600, self._refresh_approvals)
        # Poll for new pending items every 30 seconds
        self._schedule_approvals_poll()

    def _schedule_approvals_poll(self):
        """Refresh the approvals badge every 30 seconds."""
        self._refresh_approvals_badge()
        self.after(30000, self._schedule_approvals_poll)

    def _refresh_approvals_badge(self):
        """Update the sidebar badge with pending count."""
        try:
            from approval_queue import get_approval_queue
            count = get_approval_queue().count_pending()
            if count > 0:
                self._approvals_badge.configure(text=f"  {count} pending")
            else:
                self._approvals_badge.configure(text="")
        except Exception:
            pass

    def _refresh_approvals(self):
        """Rebuild the approvals list from the queue."""
        for w in self._approvals_scroll.winfo_children():
            w.destroy()

        try:
            from approval_queue import get_approval_queue, ActionStatus
            aq = get_approval_queue()
            summary = aq.summary()
            self._approvals_summary_label.configure(
                text=f"Pending: {summary.get('pending', 0)}  •  "
                     f"Auto-approved: {summary.get('auto_approved', 0)}  •  "
                     f"Approved: {summary.get('approved', 0)}  •  "
                     f"Rejected: {summary.get('rejected', 0)}  •  "
                     f"Threshold: {summary.get('threshold', 0.75):.0%}")

            filt = self._approvals_filter.get()
            if filt == "pending":
                items = aq.get_pending()
            else:
                all_items = aq.get_all(limit=100)
                if filt == "all":
                    items = all_items
                else:
                    items = [i for i in all_items if i.status.value == filt]

            if not items:
                empty = ctk.CTkFrame(self._approvals_scroll, fg_color=CARD_BG,
                                     corner_radius=12)
                empty.pack(fill="x", pady=20)
                msg = "No pending approvals." if filt == "pending" else "No items found."
                ctk.CTkLabel(empty, text=msg,
                             font=ctk.CTkFont(size=15, weight="bold"),
                             text_color=TEXT_PRIMARY).pack(pady=(30, 4))
                ctk.CTkLabel(empty,
                             text="Actions above the confidence threshold execute automatically.",
                             font=ctk.CTkFont(size=12),
                             text_color=TEXT_MUTED).pack(pady=(0, 30))
            else:
                for item in items:
                    self._approval_card(self._approvals_scroll, item, aq)

            self._refresh_approvals_badge()
        except Exception as e:
            ctk.CTkLabel(self._approvals_scroll,
                         text=f"Error loading approvals: {e}",
                         text_color="#C62828",
                         font=ctk.CTkFont(size=12)).pack(anchor="w", pady=8)

    def _approval_card(self, parent, item, aq):
        """Render a single approval queue item as a card."""
        from approval_queue import ActionStatus
        status_colors = {
            ActionStatus.PENDING: ACCENT_AMBER,
            ActionStatus.APPROVED: ACCENT_GREEN,
            ActionStatus.AUTO_APPROVED: ACCENT_GREEN,
            ActionStatus.REJECTED: "#C62828",
            ActionStatus.EXPIRED: TEXT_MUTED,
        }
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=12)
        card.pack(fill="x", pady=6)

        # Header row
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(hdr, text=item.action_type.replace("_", " ").title(),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        status_text = item.status.value.replace("_", " ").upper()
        ctk.CTkLabel(hdr, text=status_text,
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=status_colors.get(item.status, TEXT_MUTED)).pack(
                         side="right")

        # Details
        details = [
            ("Plugin", item.plugin_id),
            ("Recipient", item.recipient or "—"),
            ("Subject", item.subject or "—"),
            ("Confidence", f"{item.confidence:.0%}" if item.confidence else "—"),
            ("Queued", item.timestamp),
        ]
        for label, val in details:
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=1)
            ctk.CTkLabel(row, text=f"{label}:",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=TEXT_MUTED, width=90, anchor="w").pack(side="left")
            ctk.CTkLabel(row, text=str(val)[:120],
                         font=ctk.CTkFont(size=11),
                         text_color=TEXT_PRIMARY, anchor="w").pack(side="left")

        # Body preview
        if item.body_preview:
            ctk.CTkLabel(card, text="Preview:",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=TEXT_MUTED, anchor="w").pack(
                             anchor="w", padx=16, pady=(6, 2))
            preview_box = ctk.CTkTextbox(card, height=60,
                                         font=ctk.CTkFont(size=11),
                                         fg_color="#1A1A2E", text_color="#E0E0E0",
                                         corner_radius=6)
            preview_box.pack(fill="x", padx=16, pady=(0, 6))
            preview_box.insert("1.0", item.body_preview[:500])
            preview_box.configure(state="disabled")

        # Action buttons (only for PENDING items)
        from approval_queue import ActionStatus
        if item.status == ActionStatus.PENDING:
            btn_row = ctk.CTkFrame(card, fg_color="transparent")
            btn_row.pack(fill="x", padx=16, pady=(4, 12))

            def _approve(item_id=item.id):
                try:
                    aq.approve(item_id, approved_by="user")
                    self._log(f"[Approvals] Approved action #{item_id}")
                    self.after(100, self._refresh_approvals)
                except Exception as e:
                    self._log(f"[Approvals] Error approving #{item_id}: {e}")

            def _reject(item_id=item.id):
                try:
                    aq.reject(item_id, approved_by="user")
                    self._log(f"[Approvals] Rejected action #{item_id}")
                    self.after(100, self._refresh_approvals)
                except Exception as e:
                    self._log(f"[Approvals] Error rejecting #{item_id}: {e}")

            ctk.CTkButton(btn_row, text="Approve", width=100, height=32,
                          fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                          command=_approve).pack(side="left", padx=(0, 8))
            ctk.CTkButton(btn_row, text="Reject", width=100, height=32,
                          fg_color="#C62828", hover_color="#7F0000",
                          command=_reject).pack(side="left")
        else:
            ctk.CTkLabel(card, text="", height=8).pack()

    # ────────────────────────────────────────────────────────────────────────
    # Email Rules page
    # ────────────────────────────────────────────────────────────────────────

    def _build_rules_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["rules"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Email Rules",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="+ Add Rule", width=110, height=34,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      command=self._add_rule).pack(side="right")

        ctk.CTkLabel(page,
                     text="Define categories, keywords, and response templates for automatic email replies.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(4, 12))

        self.rules_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self.rules_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        self._refresh_rules_list()

    def _refresh_rules_list(self):
        for w in self.rules_scroll.winfo_children():
            w.destroy()

        rules = get_rules()
        if not rules:
            # Empty state
            empty = ctk.CTkFrame(self.rules_scroll, fg_color=CARD_BG, corner_radius=12)
            empty.pack(fill="x", pady=20)
            ctk.CTkLabel(empty, text="No rules yet.",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(pady=(30, 4))
            ctk.CTkLabel(empty, text="Click 'Add Rule' to create your first auto-response.",
                         font=ctk.CTkFont(size=13),
                         text_color=TEXT_MUTED).pack(pady=(0, 30))
            return

        for rule in rules:
            self._rule_card(self.rules_scroll, rule)

    def _rule_card(self, parent, rule):
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10)
        card.pack(fill="x", pady=6)

        top_row = ctk.CTkFrame(card, fg_color=CARD_BG)
        top_row.pack(fill="x", padx=16, pady=(12, 4))

        en_var = ctk.BooleanVar(value=bool(rule.get("enabled", 1)))
        ctk.CTkCheckBox(top_row, text="", variable=en_var, width=24,
                        command=lambda r=rule, v=en_var: self._toggle_rule_enabled(r, v)).pack(side="left")

        ctk.CTkLabel(top_row, text=rule["category"],
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=BRAND_BLUE).pack(side="left", padx=8)

        ctk.CTkButton(top_row, text="Edit", width=60, height=28,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=lambda r=rule: self._edit_rule(r)).pack(side="right")
        ctk.CTkButton(top_row, text="Delete", width=60, height=28,
                      fg_color="#C62828", hover_color="#7F0000",
                      command=lambda r=rule: self._delete_rule(r)).pack(side="right", padx=4)

        ctk.CTkLabel(card, text=f"Keywords: {rule['keywords']}",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=12),
                     wraplength=700, anchor="w").pack(anchor="w", padx=16, pady=(0, 4))
        ctk.CTkLabel(card, text=f"Subject: {rule.get('subject_template','')[:80]}",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=12),
                     anchor="w").pack(anchor="w", padx=16, pady=(0, 10))

    def _toggle_rule_enabled(self, rule, var):
        rule["enabled"] = 1 if var.get() else 0
        save_rule(rule)

    def _add_rule(self):
        self._rule_dialog(None)

    def _edit_rule(self, rule):
        self._rule_dialog(rule)

    def _delete_rule(self, rule):
        if messagebox.askyesno("Delete Rule", f"Delete rule '{rule['category']}'?"):
            delete_rule(rule["id"])
            self._refresh_rules_list()

    def _rule_dialog(self, rule=None):
        win = ctk.CTkToplevel(self)
        win.title("Edit Rule" if rule else "Add Rule")
        win.geometry("680x680")
        win.grab_set()

        fields = {}

        def row(label, default="", height=None):
            ctk.CTkLabel(win, text=label, font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(anchor="w", padx=24, pady=(12, 2))
            if height:
                w = ctk.CTkTextbox(win, height=height, font=ctk.CTkFont(size=12))
                w.insert("1.0", default)
            else:
                w = ctk.CTkEntry(win, height=36, font=ctk.CTkFont(size=13))
                w.insert(0, default)
            w.pack(fill="x", padx=24)
            return w

        fields["category"] = row("Category Name", rule["category"] if rule else "")
        fields["keywords"] = row("Keywords (comma separated)", rule["keywords"] if rule else "")
        fields["subject_template"] = row("Reply Subject Template",
                                         rule.get("subject_template", "") if rule else "")
        fields["body_template"] = row("Reply Body (HTML supported)",
                                      rule.get("body_template", "") if rule else "", height=240)

        ctk.CTkLabel(win, text="Use {client_name}, {date}, {subject} as placeholders.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=24)

        def save():
            bw = fields["body_template"]
            bval = bw.get("1.0", "end-1c") if isinstance(bw, ctk.CTkTextbox) else bw.get()
            r = {
                "id": rule["id"] if rule else None,
                "category": fields["category"].get().strip().upper().replace(" ", "_"),
                "keywords": fields["keywords"].get().strip(),
                "subject_template": fields["subject_template"].get().strip(),
                "body_template": bval,
                "enabled": rule.get("enabled", 1) if rule else 1,
                "sort_order": rule.get("sort_order", 99) if rule else 99,
            }
            save_rule(r)
            self._refresh_rules_list()
            win.destroy()

        ctk.CTkButton(win, text="Save Rule", height=42,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=save).pack(fill="x", padx=24, pady=16)

    # ────────────────────────────────────────────────────────────────────────
    # Activity log page
    # ────────────────────────────────────────────────────────────────────────

    def _build_activity_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["activity"] = page

        ctk.CTkLabel(page, text="Activity Log",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(24, 8))

        # Sub-tab bar
        tab_bar = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        tab_bar.pack(fill="x", padx=28, pady=(0, 8))
        self._activity_sub_panels = {}
        self._activity_sub_btns = {}

        def _show_sub(key):
            for k, p in self._activity_sub_panels.items():
                p.pack_forget()
            self._activity_sub_panels[key].pack(fill="both", expand=True,
                                                 padx=28, pady=(0, 20))
            for k, b in self._activity_sub_btns.items():
                b.configure(fg_color=BRAND_BLUE if k == key else "transparent")

        for key, label in [("email", "Email Activity"),
                           ("memory", "Memory Browser"),
                           ("events", "Event Log")]:
            btn = ctk.CTkButton(tab_bar, text=label, width=140, height=32,
                                fg_color=BRAND_BLUE if key == "email" else "transparent",
                                hover_color="#1565C0", text_color="white",
                                font=ctk.CTkFont(size=12),
                                command=lambda k=key: _show_sub(k))
            btn.pack(side="left", padx=(0, 6))
            self._activity_sub_btns[key] = btn

        # ── Email Activity sub-panel ──
        email_panel = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        self._activity_sub_panels["email"] = email_panel
        top = ctk.CTkFrame(email_panel, fg_color=BG_LIGHT)
        top.pack(fill="x", pady=(0, 6))
        ctk.CTkButton(top, text="Refresh", width=100, height=30,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._refresh_activity).pack(side="right")
        self.activity_box = ctk.CTkTextbox(email_panel,
                                           font=ctk.CTkFont(family="Courier", size=12),
                                           fg_color=CARD_BG, text_color=TEXT_PRIMARY,
                                           corner_radius=8)
        self.activity_box.pack(fill="both", expand=True)

        # ── Memory Browser sub-panel ──
        mem_panel = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        self._activity_sub_panels["memory"] = mem_panel
        mem_top = ctk.CTkFrame(mem_panel, fg_color=BG_LIGHT)
        mem_top.pack(fill="x", pady=(0, 6))
        self._mem_search_entry = ctk.CTkEntry(mem_top, placeholder_text="Search memories...",
                                              height=32, font=ctk.CTkFont(size=12))
        self._mem_search_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(mem_top, text="Search", width=90, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._search_memory).pack(side="left", padx=(0, 6))
        ctk.CTkButton(mem_top, text="Show All", width=90, height=32,
                      fg_color="transparent", hover_color="#E3F2FD",
                      text_color=BRAND_BLUE, border_width=1, border_color=BRAND_BLUE,
                      command=self._show_all_memories).pack(side="left")
        self._mem_results_box = ctk.CTkTextbox(mem_panel,
                                               font=ctk.CTkFont(family="Courier", size=11),
                                               fg_color=CARD_BG, text_color=TEXT_PRIMARY,
                                               corner_radius=8)
        self._mem_results_box.pack(fill="both", expand=True)
        self._mem_results_box.insert("1.0", "Enter a search term above to browse semantic memory.\n")
        self._mem_results_box.configure(state="disabled")

        # ── Event Log sub-panel ──
        evt_panel = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        self._activity_sub_panels["events"] = evt_panel
        evt_top = ctk.CTkFrame(evt_panel, fg_color=BG_LIGHT)
        evt_top.pack(fill="x", pady=(0, 6))
        ctk.CTkButton(evt_top, text="Refresh", width=100, height=30,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._refresh_event_log).pack(side="right")
        ctk.CTkButton(evt_top, text="Clear", width=80, height=30,
                      fg_color="#C62828", hover_color="#7F0000",
                      command=self._clear_event_log).pack(side="right", padx=(0, 8))
        self._event_log_box = ctk.CTkTextbox(evt_panel,
                                             font=ctk.CTkFont(family="Courier", size=11),
                                             fg_color="#1A1A2E", text_color="#E0E0E0",
                                             corner_radius=8)
        self._event_log_box.pack(fill="both", expand=True)

        # Show email panel by default
        email_panel.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        self._refresh_activity()
        self._refresh_event_log()

    def _refresh_activity(self):
        self.activity_box.configure(state="normal")
        self.activity_box.delete("1.0", "end")
        records = get_recent_activity(100)
        if not records:
            self.activity_box.insert("end", "No activity recorded yet.\n")
        else:
            self.activity_box.insert("end",
                f"{'Timestamp':<22} {'From':<35} {'Classification':<22} {'Action':<16} Draft\n")
            self.activity_box.insert("end", "-" * 110 + "\n")
            for r in records:
                self.activity_box.insert("end",
                    f"{r['timestamp']:<22} {str(r.get('from_email',''))[:33]:<35} "
                    f"{str(r.get('classification',''))[:20]:<22} "
                    f"{str(r.get('action',''))[:14]:<16} "
                    f"{'Yes' if r.get('draft_created') else 'No'}\n")
        self.activity_box.configure(state="disabled")

    def _search_memory(self):
        query = self._mem_search_entry.get().strip()
        if not query:
            return
        self._mem_results_box.configure(state="normal")
        self._mem_results_box.delete("1.0", "end")
        self._mem_results_box.insert("1.0", f"Searching for: {query}...\n")
        self._mem_results_box.configure(state="disabled")
        def _do_search():
            try:
                from memory_store import MemoryStore
                results = MemoryStore.search(query, n_results=20)
                lines = [f"Found {len(results)} result(s) for '{query}':\n", "-" * 80 + "\n"]
                for i, r in enumerate(results, 1):
                    meta = r.get("metadata", {})
                    lines.append(f"[{i}] Collection: {meta.get('collection','?')}  "
                                 f"Score: {r.get('distance', 0):.3f}\n")
                    lines.append(f"    Client: {meta.get('client_name', meta.get('client','?'))}  "
                                 f"Type: {meta.get('type','?')}  "
                                 f"Date: {meta.get('date', meta.get('timestamp','?'))}\n")
                    lines.append(f"    {str(r.get('document',''))[:200]}\n\n")
                self.after(0, lambda: self._update_mem_box("".join(lines)))
            except Exception as e:
                self.after(0, lambda err=str(e): self._update_mem_box(f"Error: {err}\n"))
        threading.Thread(target=_do_search, daemon=True).start()

    def _show_all_memories(self):
        self._mem_results_box.configure(state="normal")
        self._mem_results_box.delete("1.0", "end")
        self._mem_results_box.insert("1.0", "Loading all memories...\n")
        self._mem_results_box.configure(state="disabled")
        def _do_load():
            try:
                from memory_store import MemoryStore
                import chromadb
                from pathlib import Path
                CHROMA_PATH = str(Path.home() / ".mcs_email_automation" / "chroma_db")
                client = chromadb.PersistentClient(path=CHROMA_PATH)
                lines = []
                for col_name in ["client_interactions", "lessons", "documents", "general"]:
                    try:
                        col = client.get_collection(col_name)
                        count = col.count()
                        lines.append(f"\n=== {col_name} ({count} items) ===\n")
                        if count > 0:
                            data = col.get(limit=50, include=["documents", "metadatas"])
                            for doc, meta in zip(data["documents"], data["metadatas"]):
                                lines.append(f"  [{meta.get('type','?')}] "
                                             f"{meta.get('client_name', meta.get('client','?'))} "
                                             f"({meta.get('date', meta.get('timestamp','?'))})\n")
                                lines.append(f"    {str(doc)[:150]}\n")
                    except Exception:
                        lines.append(f"\n=== {col_name} (empty or not created) ===\n")
                self.after(0, lambda: self._update_mem_box("".join(lines) or "No memories stored yet.\n"))
            except Exception as e:
                self.after(0, lambda err=str(e): self._update_mem_box(f"Error: {err}\n"))
        threading.Thread(target=_do_load, daemon=True).start()

    def _update_mem_box(self, text):
        self._mem_results_box.configure(state="normal")
        self._mem_results_box.delete("1.0", "end")
        self._mem_results_box.insert("1.0", text)
        self._mem_results_box.configure(state="disabled")

    def _refresh_event_log(self):
        self._event_log_box.configure(state="normal")
        self._event_log_box.delete("1.0", "end")
        try:
            from event_bus import EventBus
            history = EventBus.get_history(limit=200)
            if not history:
                self._event_log_box.insert("1.0", "No events recorded yet.\n")
            else:
                self._event_log_box.insert("1.0",
                    f"{'Timestamp':<22} {'Event Type':<35} {'Source':<25} Payload\n")
                self._event_log_box.insert("end", "-" * 110 + "\n")
                for evt in reversed(history):
                    payload_str = str(evt.payload)[:60] if evt.payload else ""
                    self._event_log_box.insert("end",
                        f"{str(evt.timestamp)[:21]:<22} "
                        f"{evt.event_type[:33]:<35} "
                        f"{str(evt.source or '')[:23]:<25} "
                        f"{payload_str}\n")
        except Exception as e:
            self._event_log_box.insert("1.0", f"Error loading events: {e}\n")
        self._event_log_box.configure(state="disabled")

    def _clear_event_log(self):
        try:
            from event_bus import EventBus
            EventBus.clear_history()
            self._refresh_event_log()
            self._log("Event log cleared.")
        except Exception as e:
            self._log(f"Error clearing event log: {e}")

    # ────────────────────────────────────────────────────────────────────────
    # AI Usage page (Token Metering — Tier 3B)
    # ────────────────────────────────────────────────────────────────────────

    def _build_usage_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["usage"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="AI Usage & Cost",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="Refresh", width=90,
                      command=self._refresh_usage_page).pack(side="right")

        ctk.CTkLabel(page,
                     text="Tracks every Claude API call — tokens used and estimated AUD cost.",
                     font=ctk.CTkFont(size=12), text_color=TEXT_SECONDARY
                     ).pack(anchor="w", padx=28, pady=(4, 12))

        # Summary cards row
        cards_frame = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        cards_frame.pack(fill="x", padx=28, pady=(0, 12))
        self._usage_today_lbl    = ctk.CTkLabel(cards_frame, text="Today: $0.00",
                                                font=ctk.CTkFont(size=14, weight="bold"),
                                                text_color="#4CAF50")
        self._usage_today_lbl.pack(side="left", padx=(0, 24))
        self._usage_month_lbl   = ctk.CTkLabel(cards_frame, text="This Month: $0.00",
                                                font=ctk.CTkFont(size=14, weight="bold"),
                                                text_color="#2196F3")
        self._usage_month_lbl.pack(side="left", padx=(0, 24))
        self._usage_total_lbl   = ctk.CTkLabel(cards_frame, text="30-Day Total: $0.00",
                                                font=ctk.CTkFont(size=14, weight="bold"),
                                                text_color=TEXT_PRIMARY)
        self._usage_total_lbl.pack(side="left", padx=(0, 24))
        self._usage_calls_lbl   = ctk.CTkLabel(cards_frame, text="Calls: 0",
                                                font=ctk.CTkFont(size=14),
                                                text_color=TEXT_SECONDARY)
        self._usage_calls_lbl.pack(side="left")

        # Breakdown table
        ctk.CTkLabel(page, text="By Plugin (last 30 days)",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(8, 4))

        tbl_frame = ctk.CTkFrame(page, fg_color=BG_CARD, corner_radius=8)
        tbl_frame.pack(fill="x", padx=28, pady=(0, 12))
        self._usage_table_frame = tbl_frame

        # By model breakdown
        ctk.CTkLabel(page, text="By Model",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(8, 4))
        model_frame = ctk.CTkFrame(page, fg_color=BG_CARD, corner_radius=8)
        model_frame.pack(fill="x", padx=28, pady=(0, 12))
        self._usage_model_frame = model_frame

        # Daily log
        ctk.CTkLabel(page, text="Daily Log (last 30 days)",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(8, 4))
        log_frame = ctk.CTkScrollableFrame(page, fg_color=BG_CARD,
                                           corner_radius=8, height=180)
        log_frame.pack(fill="x", padx=28, pady=(0, 16))
        self._usage_log_frame = log_frame

        self._refresh_usage_page()

    def _refresh_usage_page(self):
        """Reload token usage data from the meter DB and update the UI."""
        try:
            from token_meter import get_usage_summary
            summary = get_usage_summary(days=30)
        except Exception:
            summary = {"total_calls": 0, "total_tokens": 0,
                       "total_cost_aud": 0.0, "today_cost_aud": 0.0,
                       "this_month_cost_aud": 0.0,
                       "by_plugin": [], "by_model": [], "by_day": []}

        # Update summary cards
        self._usage_today_lbl.configure(
            text=f"Today: A${summary['today_cost_aud']:.4f}")
        self._usage_month_lbl.configure(
            text=f"This Month: A${summary['this_month_cost_aud']:.4f}")
        self._usage_total_lbl.configure(
            text=f"30-Day Total: A${summary['total_cost_aud']:.4f}")
        self._usage_calls_lbl.configure(
            text=f"Calls: {summary['total_calls']:,}  •  "
                 f"Tokens: {summary['total_tokens']:,}")

        # Rebuild plugin table
        for w in self._usage_table_frame.winfo_children():
            w.destroy()
        headers = ["Plugin", "Calls", "Tokens", "Cost (AUD)"]
        for col, h in enumerate(headers):
            ctk.CTkLabel(self._usage_table_frame, text=h,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=TEXT_SECONDARY
                         ).grid(row=0, column=col, padx=12, pady=6, sticky="w")
        for row_i, row in enumerate(summary["by_plugin"], start=1):
            vals = [row["plugin_id"].replace("plugin_", ""),
                    str(row["calls"]),
                    f"{row['tokens']:,}",
                    f"A${row['cost_aud']:.4f}"]
            for col, v in enumerate(vals):
                ctk.CTkLabel(self._usage_table_frame, text=v,
                             font=ctk.CTkFont(size=11),
                             text_color=TEXT_PRIMARY
                             ).grid(row=row_i, column=col, padx=12, pady=3, sticky="w")
        if not summary["by_plugin"]:
            ctk.CTkLabel(self._usage_table_frame,
                         text="No usage recorded yet. Usage appears here after the first Claude API call.",
                         font=ctk.CTkFont(size=11), text_color=TEXT_SECONDARY
                         ).grid(row=1, column=0, columnspan=4, padx=12, pady=8, sticky="w")

        # Rebuild model table
        for w in self._usage_model_frame.winfo_children():
            w.destroy()
        for col, h in enumerate(["Model", "Tier", "Calls", "Tokens", "Cost (AUD)"]):
            ctk.CTkLabel(self._usage_model_frame, text=h,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=TEXT_SECONDARY
                         ).grid(row=0, column=col, padx=12, pady=6, sticky="w")
        for row_i, row in enumerate(summary["by_model"], start=1):
            vals = [row["model"], row.get("tier", ""),
                    str(row["calls"]),
                    f"{row['tokens']:,}",
                    f"A${row['cost_aud']:.4f}"]
            for col, v in enumerate(vals):
                ctk.CTkLabel(self._usage_model_frame, text=v,
                             font=ctk.CTkFont(size=11),
                             text_color=TEXT_PRIMARY
                             ).grid(row=row_i, column=col, padx=12, pady=3, sticky="w")
        if not summary["by_model"]:
            ctk.CTkLabel(self._usage_model_frame,
                         text="No model data yet.",
                         font=ctk.CTkFont(size=11), text_color=TEXT_SECONDARY
                         ).grid(row=1, column=0, columnspan=5, padx=12, pady=8, sticky="w")

        # Rebuild daily log
        for w in self._usage_log_frame.winfo_children():
            w.destroy()
        for row in summary["by_day"]:
            line = (f"{row['date']}   "
                    f"{row['calls']} calls   "
                    f"{row['tokens']:,} tokens   "
                    f"A${row['cost_aud']:.4f}")
            ctk.CTkLabel(self._usage_log_frame, text=line,
                         font=ctk.CTkFont(size=11), text_color=TEXT_PRIMARY,
                         anchor="w").pack(fill="x", padx=8, pady=1)
        if not summary["by_day"]:
            ctk.CTkLabel(self._usage_log_frame,
                         text="No daily data yet.",
                         font=ctk.CTkFont(size=11), text_color=TEXT_SECONDARY
                         ).pack(padx=8, pady=8)

    # ────────────────────────────────────────────────────────────────────────
    # Plugins page
    # ────────────────────────────────────────────────────────────────────────

    def _build_plugins_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["plugins"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Plugins",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="Refresh", width=100, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._refresh_plugins_list).pack(side="right")

        ctk.CTkLabel(page,
                     text="Manage your automation plugins — enable, configure schedules, and run on demand.",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(4, 12))

        self._plugins_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self._plugins_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))

        self.after(500, self._refresh_plugins_list)

    def _refresh_plugins_list(self):
        for w in self._plugins_scroll.winfo_children():
            w.destroy()

        plugins = self._loader.get_plugins()
        active_plugins = [p for p in plugins if not p.is_template]
        template_plugins = [p for p in plugins if p.is_template]

        if not active_plugins and not template_plugins:
            empty = ctk.CTkFrame(self._plugins_scroll, fg_color=CARD_BG, corner_radius=12)
            empty.pack(fill="x", pady=20)
            ctk.CTkLabel(empty, text="No plugins loaded yet.",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(pady=(30, 4))
            ctk.CTkLabel(empty, text="Plugins are discovered automatically from the plugins/ folder.",
                         font=ctk.CTkFont(size=13),
                         text_color=TEXT_MUTED).pack(pady=(0, 30))
            return

        for lp in active_plugins:
            self._plugin_card(self._plugins_scroll, lp)

        if template_plugins:
            ctk.CTkLabel(self._plugins_scroll,
                         text="Available Templates — not active",
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT_MUTED).pack(anchor="w", pady=(20, 6))
            for lp in template_plugins:
                self._plugin_card(self._plugins_scroll, lp, is_template=True)

    def _plugin_card(self, parent, lp, is_template=False):
        CORE_PLUGIN_IDS = {"plugin_email_triage", "plugin_noa_processor",
                           "plugin_asic_returns", "plugin_correspondence_logger"}
        card_fg = "#ECECEC" if is_template else CARD_BG
        text_col = TEXT_MUTED if is_template else TEXT_PRIMARY
        card = ctk.CTkFrame(parent, fg_color=card_fg, corner_radius=12)
        card.pack(fill="x", pady=6)

        # ── Row 1: identity ──
        row1 = ctk.CTkFrame(card, fg_color="transparent")
        row1.pack(fill="x", padx=16, pady=(12, 2))
        ctk.CTkLabel(row1, text=f"{lp.icon}  {lp.name}",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=text_col).pack(side="left")
        ctk.CTkLabel(row1, text=f"v{lp.version}",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left", padx=8)

        # Rename button for non-core, non-template plugins
        base_id = lp.plugin_id.replace("plugins.", "").split(".")[-1]
        if not is_template and base_id not in CORE_PLUGIN_IDS:
            def _rename_plugin(pid=lp.plugin_id, loaded_plugin=lp):
                dlg = ctk.CTkInputDialog(
                    text="Enter a new name for this plugin:",
                    title="Rename Plugin")
                new_name = dlg.get_input()
                if new_name and new_name.strip():
                    new_name = new_name.strip()
                    loaded_plugin.display_name = new_name
                    config.save_plugin_state(pid, display_name=new_name)
                    self._refresh_plugins_list()
            ctk.CTkButton(row1, text="Rename", width=60, height=24,
                          fg_color="transparent", hover_color="#E3F2FD",
                          text_color=BRAND_BLUE, border_width=1,
                          border_color=BRAND_BLUE,
                          font=ctk.CTkFont(size=11),
                          command=_rename_plugin).pack(side="left", padx=8)

        ctk.CTkLabel(card, text=lp.description,
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED,
                     wraplength=700, anchor="w").pack(anchor="w", padx=16, pady=(0, 6))

        if is_template:
            ctk.CTkLabel(card, text="Template only — copy and customise to activate.",
                         font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(anchor="w", padx=16, pady=(0, 10))
            return

        # ── Row 2: controls ──
        row2 = ctk.CTkFrame(card, fg_color="transparent")
        row2.pack(fill="x", padx=16, pady=(0, 4))

        en_var = ctk.BooleanVar(value=lp.enabled)
        ctk.CTkSwitch(row2, text="Enabled", variable=en_var, width=50,
                       onvalue=True, offvalue=False,
                       command=lambda pid=lp.plugin_id, v=en_var: (
                           self._loader.set_plugin_enabled(pid, v.get())
                       )).pack(side="left", padx=(0, 16))

        dm_var = ctk.BooleanVar(value=lp.draft_mode)
        dm_label = ctk.CTkLabel(row2,
                                text="Draft Mode: ON" if lp.draft_mode else "Draft Mode: OFF",
                                font=ctk.CTkFont(size=12),
                                text_color=DRAFT_FG if lp.draft_mode else ACCENT_AMBER)
        def _toggle_draft(pid=lp.plugin_id, v=dm_var, lbl=dm_label):
            val = v.get()
            self._loader.set_plugin_draft_mode(pid, val)
            lbl.configure(text="Draft Mode: ON" if val else "Draft Mode: OFF",
                          text_color=DRAFT_FG if val else ACCENT_AMBER)
        ctk.CTkSwitch(row2, text="", variable=dm_var, width=50,
                       onvalue=True, offvalue=False,
                       command=_toggle_draft).pack(side="left")
        dm_label.pack(side="left", padx=(4, 16))

        # Schedule dropdown
        ctk.CTkLabel(row2, text="Schedule:", font=ctk.CTkFont(size=12),
                     text_color=TEXT_PRIMARY).pack(side="left", padx=(8, 4))
        sched_labels = [s[0] for s in SCHEDULE_OPTIONS]
        current_label = _seconds_to_schedule_label(lp.schedule_seconds)
        sched_var = ctk.StringVar(value=current_label)
        def _on_schedule(choice, pid=lp.plugin_id):
            secs = dict(SCHEDULE_OPTIONS).get(choice, 0)
            self._loader.set_plugin_schedule(pid, secs)
        ctk.CTkOptionMenu(row2, variable=sched_var, values=sched_labels,
                          width=110, height=30, command=_on_schedule,
                          fg_color=BRAND_BLUE, button_color=BRAND_DARK
                          ).pack(side="left", padx=(0, 16))

        # Run Now button
        def _run_now(pid=lp.plugin_id):
            def do_run():
                self._loader.run_plugin(pid, manual=True)
                self.after(0, self._refresh_plugins_list)
            threading.Thread(target=do_run, daemon=True).start()
        ctk.CTkButton(row2, text="Run Now", width=80, height=30,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      command=_run_now).pack(side="right")

        # Delete Plugin button
        def _delete_plugin(pid=lp.plugin_id, pname=lp.name):
            # Block deletion of core plugins
            base_id = pid.replace("plugins.", "").split(".")[-1]
            if base_id in CORE_PLUGIN_IDS:
                dlg = ctk.CTkToplevel(self)
                dlg.title("Cannot Delete")
                dlg.geometry("420x160")
                dlg.grab_set()
                ctk.CTkLabel(dlg, text="This is a core plugin and cannot be deleted.\n"
                             "You can disable it instead.",
                             font=ctk.CTkFont(size=13), text_color=TEXT_PRIMARY,
                             wraplength=380, justify="center").pack(pady=(28, 16))
                ctk.CTkButton(dlg, text="OK", width=100, height=34,
                              fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                              command=dlg.destroy).pack()
                return

            # Confirmation dialog
            dlg = ctk.CTkToplevel(self)
            dlg.title("Delete Plugin")
            dlg.geometry("460x180")
            dlg.grab_set()
            ctk.CTkLabel(dlg, text=f"Are you sure you want to delete {pname}?\n"
                         "This will permanently remove the plugin file.",
                         font=ctk.CTkFont(size=13), text_color=TEXT_PRIMARY,
                         wraplength=420, justify="center").pack(pady=(24, 16))

            btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
            btn_row.pack()

            def _do_delete():
                dlg.destroy()
                self._loader.set_plugin_enabled(pid, False)
                if getattr(sys, 'frozen', False):
                    plugins_dir = os.path.join(os.path.dirname(sys.executable), 'plugins')
                else:
                    plugins_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plugins')
                # Derive filename from plugin_id
                fname = base_id + ".py"
                fpath = os.path.join(plugins_dir, fname)
                try:
                    if os.path.exists(fpath):
                        os.remove(fpath)
                except Exception as e:
                    self._log(f"Error deleting plugin file: {e}")
                self._loader.reload_plugins()
                self._log(f"Plugin {pname} deleted.")
                self.after(300, self._refresh_plugins_list)

            ctk.CTkButton(btn_row, text="Delete", width=100, height=34,
                          fg_color="#C62828", hover_color="#7F0000",
                          text_color="white",
                          command=_do_delete).pack(side="left", padx=8)
            ctk.CTkButton(btn_row, text="Cancel", width=100, height=34,
                          fg_color="transparent", hover_color="#E3F2FD",
                          text_color=TEXT_PRIMARY, border_width=1, border_color=TEXT_MUTED,
                          command=dlg.destroy).pack(side="left", padx=8)

        ctk.CTkButton(row2, text="Delete Plugin", width=100, height=30,
                      fg_color="#C62828", hover_color="#7F0000",
                      text_color="white",
                      command=_delete_plugin).pack(side="right", padx=(0, 8))

        # ── Row 3: status strip ──
        row3 = ctk.CTkFrame(card, fg_color="#F0F4FF", corner_radius=6)
        row3.pack(fill="x", padx=16, pady=(4, 4))

        ready_text = "Ready" if lp.is_ready else "Not Ready"
        ready_color = ACCENT_GREEN if lp.is_ready else ACCENT_AMBER
        ctk.CTkLabel(row3, text=ready_text, font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=ready_color).pack(side="left", padx=8, pady=4)
        ctk.CTkLabel(row3, text=f"|  {lp.schedule_label}",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left", padx=4)

        last_run_str = lp.last_run.strftime("%H:%M:%S") if lp.last_run else "—"
        ctk.CTkLabel(row3, text=f"|  Last: {last_run_str}",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left", padx=4)
        ctk.CTkLabel(row3, text=f"|  {lp.last_result}",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left", padx=4)

        if lp.last_summary:
            ctk.CTkLabel(card, text=f"Summary: {lp.last_summary}",
                         font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
                         wraplength=700, anchor="w").pack(anchor="w", padx=16, pady=(0, 2))

        # ── Plugin settings (if schema exists) ──
        schema = lp.instance.settings_schema()
        if schema:
            settings_frame = ctk.CTkFrame(card, fg_color="transparent")
            settings_frame.pack(fill="x", padx=16, pady=(4, 10))

            ctk.CTkLabel(settings_frame, text="Plugin Settings",
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))

            field_widgets = {}
            for field in schema:
                key = field["key"]
                label = field.get("label", key)
                default = field.get("default", "")
                ftype = field.get("type", "text")
                current_val = lp.instance.get_plugin_setting(key, default)

                frow = ctk.CTkFrame(settings_frame, fg_color="transparent")
                frow.pack(fill="x", pady=2)
                ctk.CTkLabel(frow, text=label, font=ctk.CTkFont(size=12),
                             text_color=TEXT_PRIMARY, width=200, anchor="w").pack(side="left")

                if ftype == "bool":
                    bvar = ctk.BooleanVar(value=current_val == "1")
                    ctk.CTkSwitch(frow, text="", variable=bvar, width=50,
                                  onvalue=True, offvalue=False).pack(side="left")
                    field_widgets[key] = ("bool", bvar)
                elif ftype == "textarea":
                    tb = ctk.CTkTextbox(frow, height=60, font=ctk.CTkFont(size=12), width=400)
                    tb.insert("1.0", current_val)
                    tb.pack(side="left", fill="x", expand=True)
                    field_widgets[key] = ("textarea", tb)
                elif ftype == "password":
                    ent = ctk.CTkEntry(frow, height=30, font=ctk.CTkFont(size=12),
                                       show="*", width=400)
                    ent.insert(0, current_val)
                    ent.pack(side="left", fill="x", expand=True)
                    field_widgets[key] = ("entry", ent)
                else:
                    ent = ctk.CTkEntry(frow, height=30, font=ctk.CTkFont(size=12), width=400)
                    ent.insert(0, current_val)
                    ent.pack(side="left", fill="x", expand=True)
                    field_widgets[key] = ("entry", ent)

                help_text = field.get("help", "")
                if help_text:
                    ctk.CTkLabel(settings_frame, text=help_text,
                                 font=ctk.CTkFont(size=10), text_color=TEXT_MUTED,
                                 wraplength=600, anchor="w").pack(anchor="w", padx=200)

            def _save_plugin_settings(pid=lp.plugin_id, inst=lp.instance, fw=field_widgets):
                for fkey, (fkind, widget) in fw.items():
                    if fkind == "bool":
                        inst.set_plugin_setting(fkey, "1" if widget.get() else "0")
                    elif fkind == "textarea":
                        inst.set_plugin_setting(fkey, widget.get("1.0", "end-1c"))
                    else:
                        inst.set_plugin_setting(fkey, widget.get())
                self._loader.reload_plugin(pid)
                self._log(f"Plugin settings saved for {inst.name}.")

            ctk.CTkButton(settings_frame, text="Save Plugin Settings",
                          width=160, height=30, fg_color=ACCENT_GREEN,
                          hover_color="#1B5E20", command=_save_plugin_settings
                          ).pack(anchor="e", pady=(6, 0))

        # ── Email Templates (if schema exists) ──
        try:
            tmpl_schema = lp.instance.email_templates_schema()
        except Exception as _tmpl_err:
            self._log(f"⚠ email_templates_schema() failed for {lp.name}: {_tmpl_err}")
            tmpl_schema = []
        tmpl_base_id = lp.plugin_id.replace("plugins.", "").split(".")[-1]
        self._log(f"[Templates] {lp.name} ({tmpl_base_id}): email_templates_schema() returned {len(tmpl_schema)} field(s)")
        if not tmpl_schema:
            # For chat-created (non-core) plugins, show a default draft_prompt template
            if tmpl_base_id not in {"plugin_email_triage", "plugin_noa_processor",
                               "plugin_asic_returns", "plugin_correspondence_logger",
                               "plugin_template"}:
                tmpl_schema = [
                    {
                        "key": "draft_prompt",
                        "label": "Draft Email Prompt",
                        "default": "You are a professional accountant's assistant. "
                                   "Draft a helpful, professional email reply.",
                        "type": "prompt",
                    },
                ]
                self._log(f"[Templates] {lp.name}: injected default Draft Email Prompt template")

        if tmpl_schema:
            from config import get_plugin_template, save_plugin_template

            tmpl_frame = ctk.CTkFrame(card, fg_color="transparent")
            tmpl_frame.pack(fill="x", padx=16, pady=(4, 10))

            ctk.CTkLabel(tmpl_frame, text="Email Templates",
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=BRAND_BLUE).pack(anchor="w", pady=(0, 2))
            ctk.CTkLabel(tmpl_frame,
                         text="Customise how this plugin structures its emails",
                         font=ctk.CTkFont(size=10), text_color=TEXT_MUTED
                         ).pack(anchor="w", pady=(0, 6))

            tmpl_widgets = {}
            for tdef in tmpl_schema:
                tkey = tdef["key"]
                tlabel = tdef.get("label", tkey)
                tdefault = tdef.get("default", "")
                ttype = tdef.get("type", "textarea")

                current_val = get_plugin_template(lp.plugin_id, tkey, tdefault)

                trow = ctk.CTkFrame(tmpl_frame, fg_color="transparent")
                trow.pack(fill="x", pady=2)
                ctk.CTkLabel(trow, text=tlabel, font=ctk.CTkFont(size=12),
                             text_color=TEXT_PRIMARY, width=200, anchor="w").pack(side="left")

                if ttype == "text":
                    ent = ctk.CTkEntry(trow, height=30, font=ctk.CTkFont(size=12), width=400)
                    ent.insert(0, current_val or "")
                    ent.pack(side="left", fill="x", expand=True)
                    tmpl_widgets[tkey] = ("entry", ent)
                else:
                    # textarea or prompt
                    h = 150 if ttype == "prompt" else 100
                    tb = ctk.CTkTextbox(trow, height=h, font=ctk.CTkFont(size=12), width=400)
                    tb.insert("1.0", current_val or "")
                    tb.pack(side="left", fill="x", expand=True)
                    tmpl_widgets[tkey] = ("textarea", tb)

            def _save_templates(pid=lp.plugin_id, tw=tmpl_widgets, frame=tmpl_frame):
                for tkey, (tkind, widget) in tw.items():
                    if tkind == "entry":
                        save_plugin_template(pid, tkey, widget.get())
                    else:
                        save_plugin_template(pid, tkey, widget.get("1.0", "end-1c"))
                # Show confirmation
                confirm = ctk.CTkLabel(frame, text="Templates saved",
                                       font=ctk.CTkFont(size=11, weight="bold"),
                                       text_color=SUCCESS_FG)
                confirm.pack(anchor="e")
                confirm.after(2000, confirm.destroy)

            ctk.CTkButton(tmpl_frame, text="Save Email Templates",
                          width=160, height=30, fg_color=ACCENT_GREEN,
                          hover_color="#1B5E20", command=_save_templates
                          ).pack(anchor="e", pady=(6, 0))

    # ────────────────────────────────────────────────────────────────────────
    # Staff & Notify page
    # ────────────────────────────────────────────────────────────────────────

    def _build_staff_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["staff"] = page

        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Staff & Notifications",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="+ Add Staff", width=120, height=34,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      command=self._add_staff_dialog).pack(side="right")

        info = ctk.CTkFrame(page, fg_color=DRAFT_BG, corner_radius=8)
        info.pack(fill="x", padx=28, pady=(12, 12))
        ctk.CTkLabel(info,
                     text="Staff listed here receive email notifications when a draft is created in Draft Mode.",
                     font=ctk.CTkFont(size=13), text_color=DRAFT_FG).pack(padx=16, pady=10)

        self._staff_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        self._staff_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        self._refresh_staff_list()

    def _refresh_staff_list(self):
        for w in self._staff_scroll.winfo_children():
            w.destroy()

        staff = get_staff()
        if not staff:
            empty = ctk.CTkFrame(self._staff_scroll, fg_color=CARD_BG, corner_radius=12)
            empty.pack(fill="x", pady=20)
            ctk.CTkLabel(empty, text="No staff added yet.",
                         font=ctk.CTkFont(size=16, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(pady=(30, 4))
            ctk.CTkLabel(empty, text="Add staff members who should receive draft notification emails.",
                         font=ctk.CTkFont(size=13),
                         text_color=TEXT_MUTED).pack(pady=(0, 30))
            return

        for s in staff:
            card = ctk.CTkFrame(self._staff_scroll, fg_color=CARD_BG, corner_radius=10)
            card.pack(fill="x", pady=4)

            row = ctk.CTkFrame(card, fg_color=CARD_BG)
            row.pack(fill="x", padx=16, pady=10)

            ctk.CTkLabel(row, text=s["name"],
                         font=ctk.CTkFont(size=14, weight="bold"),
                         text_color=TEXT_PRIMARY).pack(side="left")
            ctk.CTkLabel(row, text=s["email"],
                         font=ctk.CTkFont(size=13),
                         text_color=TEXT_MUTED).pack(side="left", padx=12)

            ctk.CTkButton(row, text="Delete", width=60, height=28,
                          fg_color="#C62828", hover_color="#7F0000",
                          command=lambda sid=s["id"]: self._delete_staff_member(sid)
                          ).pack(side="right")

            drafts_var = ctk.BooleanVar(value=bool(s.get("receives_drafts", 1)))
            def _toggle_drafts(staff_dict=s, v=drafts_var):
                staff_dict["receives_drafts"] = 1 if v.get() else 0
                save_staff(staff_dict)
            ctk.CTkCheckBox(row, text="Receives draft notifications",
                            variable=drafts_var, font=ctk.CTkFont(size=12),
                            command=_toggle_drafts).pack(side="right", padx=12)

    def _delete_staff_member(self, staff_id):
        if messagebox.askyesno("Delete Staff", "Remove this staff member?"):
            delete_staff(staff_id)
            self._refresh_staff_list()

    def _add_staff_dialog(self):
        win = ctk.CTkToplevel(self)
        win.title("Add Staff Member")
        win.geometry("460x280")
        win.grab_set()

        ctk.CTkLabel(win, text="Name", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=24, pady=(20, 2))
        name_entry = ctk.CTkEntry(win, height=36, font=ctk.CTkFont(size=13),
                                  placeholder_text="e.g. Sarah Chen")
        name_entry.pack(fill="x", padx=24)

        ctk.CTkLabel(win, text="Email", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=24, pady=(12, 2))
        email_entry = ctk.CTkEntry(win, height=36, font=ctk.CTkFont(size=13),
                                   placeholder_text="e.g. sarah@firm.com.au")
        email_entry.pack(fill="x", padx=24)

        def save():
            n = name_entry.get().strip()
            e = email_entry.get().strip()
            if not n or not e:
                messagebox.showerror("Required", "Both name and email are required.")
                return
            save_staff({"name": n, "email": e, "receives_drafts": 1, "enabled": 1})
            self._refresh_staff_list()
            win.destroy()

        ctk.CTkButton(win, text="Save", height=42,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=save).pack(fill="x", padx=24, pady=20)

    # ────────────────────────────────────────────────────────────────────────
    # Settings page
    # ────────────────────────────────────────────────────────────────────────

    def _build_settings_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["settings"] = page

        ctk.CTkLabel(page, text="Settings",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", padx=28, pady=(24, 12))

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_LIGHT)
        scroll.pack(fill="both", expand=True, padx=28, pady=(0, 20))

        self._settings_fields = {}

        # ── Microsoft 365 ─────────────────────────────────────────────────
        self._settings_section(scroll, "Microsoft 365")

        self._settings_field(scroll, "ms_tenant_id", "Tenant ID",
                             get_setting("ms_tenant_id", MCS_TENANT_ID))
        self._settings_field(scroll, "ms_client_id", "Client ID",
                             get_setting("ms_client_id", MCS_CLIENT_ID))
        self._settings_field(scroll, "ms_account_email", "Mailbox to Monitor",
                             get_setting("ms_account_email"))

        signin_row = ctk.CTkFrame(scroll, fg_color="transparent")
        signin_row.pack(fill="x", pady=(4, 8))
        ctk.CTkButton(signin_row, text="Sign in to Microsoft 365",
                      width=220, height=36, fg_color=BRAND_BLUE,
                      hover_color=BRAND_DARK,
                      command=self._settings_do_signin).pack(side="left")
        self._settings_auth_label = ctk.CTkLabel(
            signin_row, text="", font=ctk.CTkFont(size=12), text_color=TEXT_MUTED)
        self._settings_auth_label.pack(side="left", padx=12)

        if self._graph and self._graph.is_authenticated():
            self._settings_auth_label.configure(text="Connected", text_color=ACCENT_GREEN)

        # ── Claude AI ─────────────────────────────────────────────────────
        self._settings_section(scroll, "Claude AI")
        self._settings_field(scroll, "anthropic_api_key", "Anthropic API Key",
                             get_setting("anthropic_api_key"), show="*")

        # Model display row
        model_row = ctk.CTkFrame(scroll, fg_color="transparent")
        model_row.pack(fill="x", pady=(4, 0))
        fast_model = get_setting("claude_model_fast", "Not detected yet")
        reason_model = get_setting("claude_model_reasoning", "Not detected yet")
        ctk.CTkLabel(model_row, text="Fast model (Haiku):",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED,
                     width=200, anchor="w").pack(side="left")
        self._fast_model_label = ctk.CTkLabel(model_row, text=fast_model,
                     font=ctk.CTkFont(size=12), text_color=ACCENT_GREEN)
        self._fast_model_label.pack(side="left")
        model_row2 = ctk.CTkFrame(scroll, fg_color="transparent")
        model_row2.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(model_row2, text="Reasoning model (Sonnet):",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED,
                     width=200, anchor="w").pack(side="left")
        self._reason_model_label = ctk.CTkLabel(model_row2, text=reason_model,
                     font=ctk.CTkFont(size=12), text_color=ACCENT_GREEN)
        self._reason_model_label.pack(side="left")

        # Monthly token cap
        self._settings_field(scroll, "monthly_token_cap", "Monthly Token Cap",
                             get_setting("monthly_token_cap", "500000"))
        ctk.CTkLabel(scroll,
                     text="Maximum tokens per month across all plugins. Default: 500,000 (~$3-30 AUD/month).",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
                     wraplength=500, anchor="w", justify="left").pack(anchor="w", pady=(0, 4))

        # Heartbeat interval
        self._settings_field(scroll, "heartbeat_interval_seconds", "Heartbeat Interval (seconds)",
                             get_setting("heartbeat_interval_seconds", "60"))
        ctk.CTkLabel(scroll,
                     text="How often the background heartbeat fires (minimum 30s). Default: 60s.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
                     wraplength=500, anchor="w", justify="left").pack(anchor="w", pady=(0, 4))

        # ── Practice Details ──────────────────────────────────────────────
        self._settings_section(scroll, "Practice Details")
        self._settings_field(scroll, "practice_name", "Practice Name",
                             get_setting("practice_name", "MC & S"))
        self._settings_field(scroll, "monitor_folder", "Default Folder to Watch",
                             get_setting("monitor_folder", "Inbox"))

        # ── Business Hours ────────────────────────────────────────────────
        self._settings_section(scroll, "Business Hours")

        bh_row = ctk.CTkFrame(scroll, fg_color="transparent")
        bh_row.pack(fill="x", pady=2)
        bh_var = ctk.BooleanVar(value=get_setting("business_hours_enabled", "1") == "1")
        ctk.CTkCheckBox(bh_row, text="Only run plugins during business hours",
                        variable=bh_var, font=ctk.CTkFont(size=13)).pack(side="left")
        self._settings_fields["business_hours_enabled"] = ("bool", bh_var)

        hours_row = ctk.CTkFrame(scroll, fg_color="transparent")
        hours_row.pack(fill="x", pady=4)
        ctk.CTkLabel(hours_row, text="Start Hour (0-23):", font=ctk.CTkFont(size=12),
                     text_color=TEXT_PRIMARY).pack(side="left")
        start_entry = ctk.CTkEntry(hours_row, width=60, height=30)
        start_entry.insert(0, get_setting("business_hours_start", "8"))
        start_entry.pack(side="left", padx=(4, 20))
        self._settings_fields["business_hours_start"] = ("entry", start_entry)

        ctk.CTkLabel(hours_row, text="End Hour (0-23):", font=ctk.CTkFont(size=12),
                     text_color=TEXT_PRIMARY).pack(side="left")
        end_entry = ctk.CTkEntry(hours_row, width=60, height=30)
        end_entry.insert(0, get_setting("business_hours_end", "18"))
        end_entry.pack(side="left", padx=4)
        self._settings_fields["business_hours_end"] = ("entry", end_entry)

        # Business days checkboxes
        days_row = ctk.CTkFrame(scroll, fg_color="transparent")
        days_row.pack(fill="x", pady=4)
        ctk.CTkLabel(days_row, text="Business Days:", font=ctk.CTkFont(size=12),
                     text_color=TEXT_PRIMARY).pack(side="left", padx=(0, 8))
        active_days = get_setting("business_days", "1,2,3,4,5").split(",")
        day_names = [("Mon", "1"), ("Tue", "2"), ("Wed", "3"), ("Thu", "4"),
                     ("Fri", "5"), ("Sat", "6"), ("Sun", "7")]
        self._bh_day_vars = {}
        for name, num in day_names:
            var = ctk.BooleanVar(value=num in active_days)
            ctk.CTkCheckBox(days_row, text=name, variable=var, width=50,
                            font=ctk.CTkFont(size=12)).pack(side="left", padx=4)
            self._bh_day_vars[num] = var

        # ── Integrations: XPM (Xero OAuth 2.0) ───────────────────────────────────────────
        self._settings_section(scroll, "XPM / Xero Practice Manager")
        ctk.CTkLabel(
            scroll,
            text="Connect to Xero XPM via OAuth to enable WIP summaries, job lookups, client notes, and meeting prep.",
            font=ctk.CTkFont(size=12), text_color=TEXT_MUTED,
            wraplength=500, anchor="w", justify="left",
        ).pack(anchor="w", pady=(0, 6))

        # Client ID display row
        cid_row = ctk.CTkFrame(scroll, fg_color="transparent")
        cid_row.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(cid_row, text="App (Client ID):", width=130,
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(side="left")
        ctk.CTkLabel(cid_row,
                     text=get_setting("xero_client_id", "Not configured"),
                     font=ctk.CTkFont(size=12, family="Courier"),
                     text_color=TEXT_PRIMARY).pack(side="left", padx=4)

        # OAuth status + action buttons row
        xpm_row = ctk.CTkFrame(scroll, fg_color="transparent")
        xpm_row.pack(fill="x", pady=(6, 8))

        try:
            from xero_oauth import is_authorised as _xero_authed
            _xero_connected = _xero_authed()
        except Exception:
            _xero_connected = False

        self._xpm_status_label = ctk.CTkLabel(
            xpm_row,
            text="✅ Connected to Xero XPM" if _xero_connected else "⚠ Not yet authorised",
            font=ctk.CTkFont(size=12),
            text_color=ACCENT_GREEN if _xero_connected else ACCENT_AMBER,
        )
        self._xpm_status_label.pack(side="left", padx=(0, 14))

        ctk.CTkButton(
            xpm_row, text="Connect Xero", width=140, height=32,
            fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
            command=self._connect_xero,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            xpm_row, text="Test Connection", width=140, height=32,
            fg_color="#37474F", hover_color="#263238",
            command=self._test_xpm_connection,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            xpm_row, text="Disconnect", width=110, height=32,
            fg_color="#C62828", hover_color="#B71C1C",
            command=self._disconnect_xero,
        ).pack(side="left")

        # ── Integrations: FuseSign ────────────────────────────────────────────
        self._settings_section(scroll, "FuseSign")
        ctk.CTkLabel(scroll,
                     text="Connect to FuseSign to monitor signing bundles and generate engagement letters.",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED,
                     wraplength=500, anchor="w", justify="left").pack(anchor="w", pady=(0, 6))
        self._settings_field(scroll, "fusesign_api_key", "FuseSign API Key",
                             get_setting("fusesign_api_key", ""), show="*")
        self._settings_field(scroll, "fusesign_base_url", "FuseSign Base URL",
                             get_setting("fusesign_base_url", "https://api.fusesign.com/v1"))

        fusesign_row = ctk.CTkFrame(scroll, fg_color="transparent")
        fusesign_row.pack(fill="x", pady=(4, 8))
        self._fusesign_status_label = ctk.CTkLabel(fusesign_row, text="",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED)
        ctk.CTkButton(fusesign_row, text="Test FuseSign Connection", width=200, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._test_fusesign_connection).pack(side="left")
        self._fusesign_status_label.pack(side="left", padx=12)

        # ── Integrations: Microsoft Teams ─────────────────────────────────────
        self._settings_section(scroll, "Microsoft Teams")
        ctk.CTkLabel(scroll,
                     text="Connect Teams to receive automated alerts and briefings in your team channel.",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED,
                     wraplength=500, anchor="w", justify="left").pack(anchor="w", pady=(0, 6))
        self._settings_field(scroll, "teams_webhook_url", "Incoming Webhook URL",
                             get_setting("teams_webhook_url", ""))
        ctk.CTkLabel(scroll,
                     text="Create an Incoming Webhook connector in your Teams channel and paste the URL above.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
                     wraplength=500, anchor="w", justify="left").pack(anchor="w", pady=(0, 4))

        teams_row = ctk.CTkFrame(scroll, fg_color="transparent")
        teams_row.pack(fill="x", pady=(4, 8))
        self._teams_status_label = ctk.CTkLabel(teams_row, text="",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED)
        ctk.CTkButton(teams_row, text="Send Test Message", width=180, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._test_teams_connection).pack(side="left")
        self._teams_status_label.pack(side="left", padx=12)

        # ── Autonomy ──────────────────────────────────────────────────────────
        self._settings_section(scroll, "Autonomy & Approval")
        ctk.CTkLabel(scroll,
                     text="Actions above the confidence threshold execute automatically. "
                          "Below it, they queue for your approval in the Approvals tab.",
                     font=ctk.CTkFont(size=12), text_color=TEXT_MUTED,
                     wraplength=500, anchor="w", justify="left").pack(anchor="w", pady=(0, 6))

        thresh_row = ctk.CTkFrame(scroll, fg_color="transparent")
        thresh_row.pack(fill="x", pady=4)
        ctk.CTkLabel(thresh_row, text="Auto-approve threshold:",
                     font=ctk.CTkFont(size=13), text_color=TEXT_PRIMARY,
                     width=200, anchor="w").pack(side="left")
        thresh_val = float(get_setting("approval_auto_threshold", "0.75"))
        self._thresh_label = ctk.CTkLabel(thresh_row,
                     text=f"{thresh_val:.0%}",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=ACCENT_GREEN)
        self._thresh_label.pack(side="right", padx=8)
        self._thresh_slider = ctk.CTkSlider(thresh_row, from_=0.5, to=1.0,
                     number_of_steps=10, width=200)
        self._thresh_slider.set(thresh_val)
        self._thresh_slider.configure(command=self._on_threshold_change)
        self._thresh_slider.pack(side="left", padx=8)

        expiry_row = ctk.CTkFrame(scroll, fg_color="transparent")
        expiry_row.pack(fill="x", pady=4)
        ctk.CTkLabel(expiry_row, text="Pending action expiry (hours):",
                     font=ctk.CTkFont(size=13), text_color=TEXT_PRIMARY,
                     width=200, anchor="w").pack(side="left")
        expiry_entry = ctk.CTkEntry(expiry_row, width=80, height=30)
        expiry_entry.insert(0, get_setting("approval_expiry_hours", "48"))
        expiry_entry.pack(side="left", padx=4)
        self._settings_fields["approval_expiry_hours"] = ("entry", expiry_entry)

        # ── Email Signature ───────────────────────────────────────────────
        self._settings_section(scroll, "Email Signature")
        ctk.CTkLabel(scroll, text="Your email signature",
                     font=ctk.CTkFont(size=13),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(0, 4))

        # Preview box
        self._sig_preview_frame = ctk.CTkFrame(
            scroll, width=300, height=120, fg_color="#F0F0F0",
            corner_radius=10, border_width=2, border_color="#555555")
        self._sig_preview_frame.pack(anchor="w", pady=(0, 8))
        self._sig_preview_frame.pack_propagate(False)

        self._sig_preview_label = ctk.CTkLabel(
            self._sig_preview_frame, text="No signature image uploaded",
            font=ctk.CTkFont(size=12), text_color=TEXT_MUTED)
        self._sig_preview_label.pack(expand=True)

        # Keep a reference to prevent garbage collection of the CTkImage
        self._sig_ctk_image = None

        # Load existing signature image into preview
        self._refresh_signature_preview()

        # Buttons row
        sig_btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        sig_btn_row.pack(anchor="w", pady=(0, 4))
        ctk.CTkButton(sig_btn_row, text="Upload Signature Image",
                      width=180, height=34, fg_color=BRAND_BLUE,
                      hover_color=BRAND_DARK,
                      command=self._upload_signature_image).pack(side="left", padx=(0, 8))
        ctk.CTkButton(sig_btn_row, text="Remove Signature",
                      width=140, height=34, fg_color="#C62828",
                      hover_color="#7F0000",
                      command=self._remove_signature_image).pack(side="left")

        # Status label for upload feedback
        self._sig_status_label = ctk.CTkLabel(
            scroll, text="", font=ctk.CTkFont(size=12), text_color=ACCENT_GREEN)
        self._sig_status_label.pack(anchor="w")

        ctk.CTkLabel(scroll,
                     text="Take a screenshot of your Outlook signature and upload it here.\n"
                          "It will be appended to all automated email replies.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
                     wraplength=500, anchor="w", justify="left").pack(anchor="w", pady=(2, 0))

        # ── Save button ──────────────────────────────────────────────────
        ctk.CTkButton(scroll, text="Save All Settings", height=42,
                      fg_color=ACCENT_GREEN, hover_color="#1B5E20",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._save_all_settings).pack(fill="x", pady=(20, 10))

    def _settings_section(self, parent, title):
        ctk.CTkLabel(parent, text=title,
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=BRAND_BLUE).pack(anchor="w", pady=(16, 6))

    def _settings_field(self, parent, key, label, value, show=None):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=3)
        ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=13),
                     text_color=TEXT_PRIMARY, width=200, anchor="w").pack(side="left")
        kwargs = {"height": 34, "font": ctk.CTkFont(size=13)}
        if show:
            kwargs["show"] = show
        entry = ctk.CTkEntry(row, **kwargs)
        entry.insert(0, value or "")
        entry.pack(side="left", fill="x", expand=True)
        self._settings_fields[key] = ("entry", entry)

    def _on_threshold_change(self, value):
        """Live-update the threshold label as the slider moves."""
        pct = f"{float(value):.0%}"
        self._thresh_label.configure(text=pct)
        set_setting("approval_auto_threshold", str(round(float(value), 2)))
        from approval_queue import get_approval_queue
        get_approval_queue().set_threshold(float(value))

    def _connect_xero(self):
        """Launch the Xero OAuth flow in a background thread."""
        self._xpm_status_label.configure(text="Opening Xero login...", text_color=TEXT_MUTED)
        def _do_connect():
            try:
                from xero_oauth import start_auth_flow
                token_data = start_auth_flow(timeout=300)
                self.after(0, lambda: self._xpm_status_label.configure(
                    text="✅ Connected to Xero XPM", text_color=ACCENT_GREEN))
            except Exception as e:
                self.after(0, lambda err=str(e): self._xpm_status_label.configure(
                    text=f"Failed: {err[:80]}", text_color="#C62828"))
        threading.Thread(target=_do_connect, daemon=True).start()

    def _disconnect_xero(self):
        """Revoke the Xero refresh token and update the status label."""
        try:
            from xero_oauth import revoke_token
            revoke_token()
            self._xpm_status_label.configure(
                text="⚠ Disconnected", text_color=ACCENT_AMBER)
        except Exception as e:
            messagebox.showerror("Disconnect Error", str(e))

    def _test_xpm_connection(self):
        """Test the XPM connection using the stored OAuth token."""
        self._xpm_status_label.configure(text="Testing...", text_color=TEXT_MUTED)
        def _do_test():
            try:
                from gateway_client import XPMClient
                client = XPMClient()
                clients = client.list_clients(page=1, page_size=1)
                count = len(clients) if isinstance(clients, list) else 1
                self.after(0, lambda: self._xpm_status_label.configure(
                    text=f"✅ Connected ({count} client(s) found)",
                    text_color=ACCENT_GREEN))
            except Exception as e:
                self.after(0, lambda err=str(e): self._xpm_status_label.configure(
                    text=f"Failed: {err[:80]}", text_color="#C62828"))
        threading.Thread(target=_do_test, daemon=True).start()

    def _test_fusesign_connection(self):
        """Save current FuseSign credentials then test the connection."""
        api_key = self._settings_fields.get("fusesign_api_key", (None, None))[1]
        base_url = self._settings_fields.get("fusesign_base_url", (None, None))[1]
        if api_key:
            set_setting("fusesign_api_key", api_key.get().strip())
        if base_url:
            set_setting("fusesign_base_url", base_url.get().strip())
        self._fusesign_status_label.configure(text="Testing...", text_color=TEXT_MUTED)
        def _do_test():
            try:
                from gateway_client import FuseSignClient
                client = FuseSignClient()
                envelopes = client.list_envelopes(status="pending", limit=1)
                self.after(0, lambda: self._fusesign_status_label.configure(
                    text=f"Connected ({len(envelopes)} pending envelope(s))",
                    text_color=ACCENT_GREEN))
            except Exception as e:
                self.after(0, lambda err=str(e): self._fusesign_status_label.configure(
                    text=f"Failed: {err[:60]}", text_color="#C62828"))
        threading.Thread(target=_do_test, daemon=True).start()

    def _test_teams_connection(self):
        """Save current Teams webhook then send a test message."""
        webhook = self._settings_fields.get("teams_webhook_url", (None, None))[1]
        if webhook:
            set_setting("teams_webhook_url", webhook.get().strip())
        self._teams_status_label.configure(text="Sending...", text_color=TEXT_MUTED)
        def _do_test():
            try:
                from gateway_client import TeamsClient
                client = TeamsClient()
                client.send_alert(
                    title="CoWorker Test",
                    message="Teams integration is working correctly.",
                    color="00C853")
                self.after(0, lambda: self._teams_status_label.configure(
                    text="Test message sent successfully", text_color=ACCENT_GREEN))
            except Exception as e:
                self.after(0, lambda err=str(e): self._teams_status_label.configure(
                    text=f"Failed: {err[:60]}", text_color="#C62828"))
        threading.Thread(target=_do_test, daemon=True).start()

    def _save_all_settings(self):
        for key, (kind, widget) in self._settings_fields.items():
            if kind == "bool":
                set_setting(key, "1" if widget.get() else "0")
            elif kind == "entry":
                set_setting(key, widget.get().strip())

        # Save business days
        active = [num for num, var in self._bh_day_vars.items() if var.get()]
        set_setting("business_days", ",".join(sorted(active)))

        # Save threshold slider value
        if hasattr(self, "_thresh_slider"):
            set_setting("approval_auto_threshold",
                        str(round(float(self._thresh_slider.get()), 2)))

        # Re-init Claude client with new API key
        self._loader.set_claude()

        # Refresh model labels
        if hasattr(self, "_fast_model_label"):
            self._fast_model_label.configure(
                text=get_setting("claude_model_fast", "Not detected yet"))
        if hasattr(self, "_reason_model_label"):
            self._reason_model_label.configure(
                text=get_setting("claude_model_reasoning", "Not detected yet"))

        self._log("Settings saved.")
        messagebox.showinfo("Settings", "All settings saved successfully.")

    def _settings_do_signin(self):
        tid = self._settings_fields["ms_tenant_id"][1].get().strip() or None
        cid = self._settings_fields["ms_client_id"][1].get().strip() or None
        self._graph = GraphClient(tenant_id=tid, client_id=cid)
        self._settings_auth_label.configure(text="Opening browser...", text_color=TEXT_MUTED)

        def callback(success, error):
            if success:
                self.after(0, self._settings_signin_success)
            else:
                self.after(0, lambda: self._settings_auth_label.configure(
                    text=f"Failed: {error}", text_color="#C62828"))

        self._graph.authenticate(callback=callback)

    def _settings_signin_success(self):
        self._settings_auth_label.configure(text="Connected", text_color=ACCENT_GREEN)
        self.auth_status_label.configure(text="Connected", text_color="#66BB6A")
        self._loader.set_graph(self._graph)
        self._log("Connected to Microsoft 365 (from Settings).")

    # ── Signature image helpers ───────────────────────────────────────────

    def _get_signature_dest_path(self) -> str:
        """Return the canonical path for the saved signature image."""
        from pathlib import Path
        return str(Path.home() / ".mcs_email_automation" / "signature.png")

    def _refresh_signature_preview(self):
        """Load the saved signature image into the preview box, or show placeholder."""
        sig_path = get_setting("signature_image_path", "")
        if sig_path and os.path.isfile(sig_path):
            try:
                from PIL import Image as PILImageLib
                img = PILImageLib.open(sig_path)
                # Scale to fit 300x120 box while maintaining aspect ratio
                img.thumbnail((296, 116), PILImageLib.LANCZOS)
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img,
                                       size=(img.width, img.height))
                self._sig_ctk_image = ctk_img  # prevent GC
                self._sig_preview_label.configure(image=ctk_img, text="")
            except ImportError:
                self._sig_preview_label.configure(
                    image=None, text="Pillow not installed.\nRun: pip install Pillow")
            except Exception as e:
                self._sig_preview_label.configure(
                    image=None, text=f"Could not load image:\n{e}")
        else:
            self._sig_ctk_image = None
            self._sig_preview_label.configure(
                image=None if not hasattr(self._sig_preview_label, '_image') else None,
                text="No signature image uploaded")

    def _upload_signature_image(self):
        """Open a file dialog to select a signature image and save it."""
        filepath = filedialog.askopenfilename(
            title="Select Signature Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif"),
                ("PNG", "*.png"),
                ("JPEG", "*.jpg *.jpeg"),
                ("GIF", "*.gif"),
            ],
        )
        if not filepath:
            return  # User cancelled

        def do_upload():
            try:
                from PIL import Image as PILImageLib
            except ImportError:
                self.after(0, lambda: self._sig_status_label.configure(
                    text="Please run: pip install Pillow", text_color="#C62828"))
                return

            try:
                dest = self._get_signature_dest_path()
                os.makedirs(os.path.dirname(dest), exist_ok=True)

                # Convert to PNG regardless of input format
                img = PILImageLib.open(filepath)
                img.save(dest, "PNG")

                set_setting("signature_image_path", dest)

                # Clear graph signature cache so the new image is picked up
                if self._graph:
                    self._graph.clear_signature_cache()

                self.after(0, self._refresh_signature_preview)
                self.after(0, lambda: self._sig_status_label.configure(
                    text="Signature image saved", text_color=ACCENT_GREEN))
            except Exception as e:
                self.after(0, lambda: self._sig_status_label.configure(
                    text=f"Error: {e}", text_color="#C62828"))

        threading.Thread(target=do_upload, daemon=True).start()

    def _remove_signature_image(self):
        """Delete the signature image and clear the setting."""
        try:
            sig_path = get_setting("signature_image_path", "")
            if sig_path and os.path.isfile(sig_path):
                os.remove(sig_path)
        except Exception:
            pass

        set_setting("signature_image_path", "")

        if self._graph:
            self._graph.clear_signature_cache()

        self._refresh_signature_preview()
        self._sig_status_label.configure(text="Signature removed.", text_color=TEXT_MUTED)

    # ────────────────────────────────────────────────────────────────────────
    # Chat page — AI Automation Builder
    # ────────────────────────────────────────────────────────────────────────

    def _build_chat_page(self):
        page = ctk.CTkFrame(self.content, fg_color=BG_LIGHT, corner_radius=0)
        self._pages["chat"] = page

        self._chat_messages = []  # list of {"role": ..., "content": ...}

        # Header row
        top = ctk.CTkFrame(page, fg_color=BG_LIGHT)
        top.pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(top, text="Chat — Automation Builder",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(side="left")
        ctk.CTkButton(top, text="What can I build?", width=140, height=32,
                      fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
                      command=self._chat_show_examples).pack(side="right")

        ctk.CTkLabel(page,
                     text="Describe any automation in plain English. "
                          "Simple requests use templates (Tier 1). "
                          "Complex workflows with XPM, Teams, memory, or custom logic "
                          "are built as full Python plugins (Tier 2).",
                     text_color=TEXT_MUTED, font=ctk.CTkFont(size=13),
                     wraplength=900, justify="left").pack(anchor="w", padx=28, pady=(4, 8))

        # Chat history area
        self._chat_scroll = ctk.CTkScrollableFrame(
            page, fg_color="#1A1A2E", corner_radius=8)
        self._chat_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 8))

        # Typing indicator (hidden by default)
        self._chat_typing_label = ctk.CTkLabel(
            self._chat_scroll, text="", font=ctk.CTkFont(size=12),
            text_color="#888888")

        # Input area
        input_row = ctk.CTkFrame(page, fg_color=BG_LIGHT, height=50)
        input_row.pack(fill="x", padx=28, pady=(0, 20))
        input_row.pack_propagate(False)

        self._chat_input = ctk.CTkEntry(
            input_row, height=42, font=ctk.CTkFont(size=14),
            placeholder_text="Describe the automation you want...")
        self._chat_input.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._chat_input.bind("<Return>", lambda e: self._chat_send())

        self._chat_send_btn = ctk.CTkButton(
            input_row, text="Send", width=80, height=42,
            fg_color=BRAND_BLUE, hover_color=BRAND_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._chat_send)
        self._chat_send_btn.pack(side="right")

    def _chat_add_bubble(self, role, text):
        """Add a message bubble to the chat display."""
        msg_frame = ctk.CTkFrame(self._chat_scroll, fg_color="transparent")
        msg_frame.pack(fill="x", padx=8, pady=4)

        if role == "user":
            bubble = ctk.CTkFrame(msg_frame, fg_color=BRAND_BLUE, corner_radius=12)
            bubble.pack(side="right", padx=(60, 0))
            ctk.CTkLabel(bubble, text=text, text_color="white",
                         wraplength=500, justify="left",
                         font=ctk.CTkFont(size=13)).pack(padx=14, pady=10)
        elif role == "assistant":
            bubble = ctk.CTkFrame(msg_frame, fg_color=CHAT_DARK, corner_radius=12)
            bubble.pack(side="left", padx=(0, 60))
            ctk.CTkLabel(bubble, text=text, text_color="#E0E0E0",
                         wraplength=500, justify="left",
                         font=ctk.CTkFont(size=13)).pack(padx=14, pady=10)
        elif role == "system":
            bubble = ctk.CTkFrame(msg_frame, fg_color="#1B3A1B", corner_radius=8)
            bubble.pack(anchor="center")
            ctk.CTkLabel(bubble, text=text, text_color="#A5D6A7",
                         wraplength=600, justify="left",
                         font=ctk.CTkFont(size=12)).pack(padx=12, pady=6)

        self.after(50, lambda: self._chat_scroll_to_bottom())

    def _chat_scroll_to_bottom(self):
        self._chat_scroll._parent_canvas.update_idletasks()
        self._chat_scroll._parent_canvas.yview_moveto(1.0)

    def _chat_show_typing(self, show=True):
        if show:
            self._chat_typing_label.configure(text="Thinking...")
            self._chat_typing_label.pack(anchor="w", padx=16, pady=4)
        else:
            self._chat_typing_label.pack_forget()

    def _chat_send(self):
        text = self._chat_input.get().strip()
        if not text:
            return

        api_key = get_setting("anthropic_api_key")
        if not api_key or anthropic_lib is None:
            self._chat_add_bubble("system",
                                  "Please add your Anthropic API key in Settings first.")
            return

        self._chat_input.delete(0, "end")
        self._chat_add_bubble("user", text)
        self._chat_messages.append({"role": "user", "content": text})

        self._chat_send_btn.configure(state="disabled")
        self._chat_show_typing(True)

        def do_call():
            try:
                client = anthropic_lib.Anthropic(api_key=api_key)
                from config import get_claude_model, get_claude_model_reasoning
                # Tier 2 detection: use Sonnet (reasoning) if the conversation
                # contains keywords suggesting a complex custom plugin is needed.
                last_user_msg = text.lower()
                tier2_signals = [
                    "xpm", "fusesign", "teams", "memory", "remember",
                    "database", "sqlite", "complex", "multi-step", "workflow",
                    "custom plugin", "write a plugin", "build a plugin",
                    "event", "heartbeat", "schedule", "deadline",
                    "correspondence", "log", "track", "report",
                ]
                is_tier2 = any(sig in last_user_msg for sig in tier2_signals)
                # Also check if conversation already has a Tier 2 tool call in progress
                for msg in self._chat_messages:
                    if msg["role"] == "assistant" and "create_plugin" in msg.get("content", ""):
                        is_tier2 = True
                        break

                model = get_claude_model_reasoning() if is_tier2 else get_claude_model()
                max_tokens = 8192 if is_tier2 else 4096

                response = client.messages.create(
                    model=model,
                    max_tokens=max_tokens,
                    system=CHAT_SYSTEM_PROMPT,
                    messages=self._chat_messages,
                )
                reply_text = response.content[0].text
                self._chat_messages.append({"role": "assistant", "content": reply_text})
                self.after(0, lambda: self._chat_handle_response(reply_text))
            except Exception as e:
                self.after(0, lambda: self._chat_add_bubble("system", f"Error: {e}"))
            finally:
                self.after(0, lambda: self._chat_send_btn.configure(state="normal"))
                self.after(0, lambda: self._chat_show_typing(False))

        threading.Thread(target=do_call, daemon=True).start()

    def _chat_handle_response(self, text):
        """Parse response for tool calls, execute them, and display results."""
        tools, clean_text = self._extract_tool_calls(text)

        if clean_text.strip():
            self._chat_add_bubble("assistant", clean_text.strip())

        for tool in tools:
            self._chat_execute_tool(tool)

    def _extract_tool_calls(self, text):
        """Find and parse JSON tool-call blocks from the assistant response."""
        tools = []
        # Strip markdown code fences for easier parsing
        stripped = re.sub(r'```json\s*', '', text)
        stripped = re.sub(r'```\s*', '', stripped)
        clean = stripped

        positions = []  # (start, end) of tool JSON in `stripped`
        i = 0
        while i < len(stripped):
            if stripped[i] == '{':
                depth = 0
                start = i
                for j in range(i, len(stripped)):
                    if stripped[j] == '{':
                        depth += 1
                    elif stripped[j] == '}':
                        depth -= 1
                        if depth == 0:
                            candidate = stripped[start:j + 1]
                            try:
                                obj = json.loads(candidate)
                                if isinstance(obj, dict) and "tool" in obj:
                                    tools.append(obj)
                                    positions.append((start, j + 1))
                            except (json.JSONDecodeError, ValueError):
                                pass
                            i = j + 1
                            break
                else:
                    i += 1
            else:
                i += 1

        # Remove tool JSON blocks from display text (reverse order to preserve offsets)
        for start, end in reversed(positions):
            clean = clean[:start] + clean[end:]

        clean = re.sub(r'\n{3,}', '\n\n', clean).strip()
        return tools, clean

    def _validate_plugin_code(self, code, filename):
        """
        Validate and auto-fix common Chat-generated plugin mistakes.
        Returns (is_valid, fixed_code).
        """
        fixed = code
        issues = []

        # Fix 1: wrong schedule attribute name
        if re.search(r'^\s+schedule\s*=\s*Schedule\.', fixed, re.MULTILINE):
            fixed = re.sub(
                r'(\s+)schedule(\s*=\s*Schedule\.)',
                r'\1default_schedule\2',
                fixed
            )
            issues.append("Fixed: 'schedule' renamed to 'default_schedule'")

        # Fix 2: Schedule constant style
        schedule_constants = {
            'Schedule.EVERY_1_MINUTE': 'Schedule.every_minutes(1)',
            'Schedule.EVERY_5_MINUTES': 'Schedule.every_minutes(5)',
            'Schedule.EVERY_15_MINUTES': 'Schedule.every_minutes(15)',
            'Schedule.EVERY_30_MINUTES': 'Schedule.every_minutes(30)',
            'Schedule.EVERY_HOUR': 'Schedule.every_hours(1)',
            'Schedule.EVERY_4_HOURS': 'Schedule.every_hours(4)',
            'Schedule.EVERY_24_HOURS': 'Schedule.daily_at(8)',
            'Schedule.MANUAL': 'Schedule.manual_only()',
            'Schedule.MANUAL_ONLY': 'Schedule.manual_only()',
        }
        for wrong, right in schedule_constants.items():
            if wrong in fixed:
                fixed = fixed.replace(wrong, right)
                issues.append(f"Fixed: '{wrong}' replaced with '{right}'")

        # Fix 3: keyword arguments on graph methods
        graph_method_fixes = [
            (r'\.send_email\(\s*to\s*=\s*', '.send_email('),
            (r'\.send_email\(\s*recipient\s*=\s*', '.send_email('),
            (r'\.create_draft\(\s*to\s*=\s*', '.create_draft('),
            (r'\.create_draft\(\s*recipient\s*=\s*', '.create_draft('),
            (r'\.fetch_unread_emails\(\s*folder\s*=\s*', '.fetch_unread_emails('),
            (r'\.mark_as_read\(\s*message_id\s*=\s*', '.mark_as_read('),
            (r'\.move_email\(\s*message_id\s*=\s*', '.move_email('),
            (r'\.flag_email\(\s*message_id\s*=\s*', '.flag_email('),
        ]
        for pattern, replacement in graph_method_fixes:
            if re.search(pattern, fixed):
                fixed = re.sub(pattern, replacement, fixed)
                issues.append("Fixed: removed keyword argument from graph method")

        # Fix 4: invalid PluginResult fields
        invalid_fields = [
            'message=', 'result=', 'output=',
            'emails_processed=', 'count=', 'data='
        ]
        for field in invalid_fields:
            if field in fixed:
                fixed = re.sub(
                    rf',?\s*{re.escape(field)}[^,\)]*', '', fixed
                )
                issues.append(
                    f"Fixed: removed invalid PluginResult field '{field}'"
                )

        # Fix 5: missing load() method
        if 'def load(' not in fixed:
            fixed = fixed.replace(
                '    def run(self, context',
                '    def load(self, context: PluginContext) -> bool:\n'
                '        return bool(context.graph)\n\n'
                '    def run(self, context'
            )
            issues.append("Fixed: added missing load() method")

        # Fix 6: missing requires_graph attribute
        if 'requires_graph' not in fixed:
            fixed = fixed.replace(
                '    default_schedule',
                '    requires_graph = True\n    default_schedule'
            )
            issues.append("Fixed: added missing requires_graph attribute")

        # Fix 7: missing signature image check
        if ('send_email(' in fixed or 'create_draft(' in fixed):
            if 'get_signature_image_path' not in fixed:
                issues.append(
                    "Warning: plugin sends/drafts emails but does not "
                    "include signature image check. Consider adding "
                    "get_signature_image_path() logic."
                )

        # Fix 8: hardcoded model string — replace with get_claude_model()
        model_pattern = r'model\s*=\s*["\']claude-[^"\']+["\']'
        if re.search(model_pattern, fixed):
            fixed = re.sub(
                model_pattern,
                'model=self.get_claude_model()',
                fixed
            )
            issues.append(
                "Fixed: replaced hardcoded model string with "
                "self.get_claude_model()"
            )

        # Fix 9: context.memory used without None check
        if 'context.memory.' in fixed and 'context.memory is None' not in fixed:
            issues.append(
                "Warning: context.memory is used but no None check found. "
                "Add: if context.memory is None: return PluginResult(success=True, "
                "summary='Memory unavailable', items_skipped=1)"
            )

        # Fix 10: context.gateway used without availability check
        if 'context.gateway.' in fixed:
            for svc in ['xpm', 'fusesign', 'teams']:
                if f'context.gateway.{svc}.' in fixed:
                    if f'is_available("{svc}")' not in fixed and f"is_available('{svc}')" not in fixed:
                        issues.append(
                            f"Warning: context.gateway.{svc} is used but no "
                            f'is_available("{svc}") check found.'
                        )

        # Fix 11: context.claude_reason used but model not set to reasoning
        if 'context.claude_reason' in fixed and 'get_claude_model_reasoning' not in fixed:
            fixed = fixed.replace(
                'context.claude_reason.messages.create(',
                'context.claude_reason.messages.create(  # uses Sonnet reasoning model\n'
            )
            issues.append(
                "Warning: context.claude_reason used but get_claude_model_reasoning() "
                "not called. Use model=self.get_claude_model_reasoning()."
            )

        # Fix 12: ensure PluginResult is imported
        if 'PluginResult' in fixed and 'from plugin_base import' in fixed:
            if 'PluginResult' not in fixed.split('from plugin_base import')[1].split('\n')[0]:
                fixed = fixed.replace(
                    'from plugin_base import AgentPlugin, PluginContext',
                    'from plugin_base import AgentPlugin, PluginContext, PluginResult'
                )
                issues.append("Fixed: added PluginResult to plugin_base import")

        if issues:
            summary = (
                f"[Plugin Validator] Auto-fixed {len(issues)} issue(s) "
                f"in {filename}:\n" +
                "\n".join(f"  \u2022 {i}" for i in issues)
            )
            self.after(0, lambda s=summary: self._log(s))

        return True, fixed

    def _build_plugin_from_template(self, template_type, params):
        """Fill template parameters and return complete plugin code."""
        class_name = ''.join(
            word.capitalize()
            for word in params.get('plugin_name', 'Custom').split()
        ) + 'Plugin'

        if template_type == "SENDER_AUTO_REPLY":
            return SENDER_AUTO_REPLY_TEMPLATE.format(
                class_name=class_name,
                plugin_name=params['plugin_name'],
                description=params.get('description', ''),
                sender_email=params['sender_email'].lower(),
                reply_body_repr=repr(params.get('reply_body_html',
                    '<p>Thank you for your email. We will be in touch shortly.</p>')),
                draft_mode_bool=str(params.get('draft_mode', True)),
                schedule_minutes=params.get('schedule_minutes', 5),
            )

        elif template_type == "SENDER_AI_REPLY":
            return SENDER_AI_REPLY_TEMPLATE.format(
                class_name=class_name,
                plugin_name=params['plugin_name'],
                description=params.get('description', ''),
                sender_email=params['sender_email'].lower(),
                ai_instructions_repr=repr(params.get('ai_instructions',
                    'Draft a professional, helpful reply.')),
                schedule_minutes=params.get('schedule_minutes', 5),
            )

        elif template_type == "FORWARD_AND_FILE":
            return FORWARD_AND_FILE_TEMPLATE.format(
                class_name=class_name,
                plugin_name=params['plugin_name'],
                description=params.get('description', ''),
                sender_email=params['sender_email'].lower(),
                forward_to=params['forward_to'],
                forward_note=params.get('forward_note', ''),
                folder_name=params['folder_name'],
                schedule_minutes=params.get('schedule_minutes', 5),
            )

        elif template_type == "KEYWORD_AUTO_REPLY":
            keywords_list = repr(params.get('keywords', []))
            return KEYWORD_AUTO_REPLY_TEMPLATE.format(
                class_name=class_name,
                plugin_name=params['plugin_name'],
                description=params.get('description', ''),
                keywords_list=keywords_list,
                reply_body_repr=repr(params.get('reply_body_html',
                    '<p>Thank you for your email.</p>')),
                draft_mode_bool=str(params.get('draft_mode', True)),
                schedule_minutes=params.get('schedule_minutes', 5),
            )

        raise ValueError(f"Unknown template type: {template_type}")

    def _write_plugin_file(self, filename, code):
        """Write plugin code to the plugins directory (and source dir if frozen)."""
        if not filename.startswith("plugin_"):
            filename = f"plugin_{filename}"
        if not filename.endswith(".py"):
            filename += ".py"

        if getattr(sys, 'frozen', False):
            plugins_dir = os.path.join(os.path.dirname(sys.executable), 'plugins')
        else:
            plugins_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plugins')
        filepath = os.path.join(plugins_dir, filename)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(code)
        # Also write to source plugins folder so the file survives rebuilds
        if getattr(sys, 'frozen', False):
            source_plugins_dir = os.path.normpath(
                os.path.join(os.path.dirname(sys.executable), '..', '..', 'plugins'))
            if os.path.exists(source_plugins_dir):
                source_path = os.path.join(source_plugins_dir, filename)
                with open(source_path, "w", encoding="utf-8") as f:
                    f.write(code)
        return filename

    def _chat_execute_tool(self, tool):
        """Execute a parsed tool call and show result in chat."""
        tool_name = tool.get("tool", "")

        if tool_name == "create_email_rule":
            try:
                rule = {
                    "category": tool.get("category", "UNKNOWN").upper().replace(" ", "_"),
                    "keywords": tool.get("keywords", ""),
                    "subject_template": tool.get("subject_template", ""),
                    "body_template": tool.get("body_template", ""),
                    "enabled": tool.get("enabled", 1),
                    "sort_order": 99,
                }
                save_rule(rule)
                self._refresh_rules_list()
                self._chat_add_bubble("system",
                    f"Created email rule: {rule['category']}\n"
                    f"Keywords: {rule['keywords']}\n"
                    f"Find it in the Email Rules tab.")
            except Exception as e:
                self._chat_add_bubble("system", f"Error creating rule: {e}")

        elif tool_name == "create_plugin_from_template":
            try:
                template_type = tool.get("template_type", "")
                filename = tool.get("plugin_filename", "")
                if not template_type or not filename:
                    self._chat_add_bubble("system",
                        "Error: missing template_type or plugin_filename.")
                    return

                code = self._build_plugin_from_template(template_type, tool)
                self._write_plugin_file(filename, code)
                self._loader.reload_plugins()
                self.after(500, self._refresh_plugins_list)
                plugin_name = tool.get('plugin_name', filename)
                self._chat_add_bubble("system",
                    f"Plugin '{plugin_name}' created from "
                    f"{template_type} template.\n"
                    f"File: {filename}\n"
                    f"Find it in the Plugins tab.")
                self.after(500, self._switch_to_plugins)
            except Exception as e:
                self._chat_add_bubble("system",
                    f"Error creating plugin from template: {e}")

        elif tool_name == "create_plugin":
            # Legacy fallback — free-form code generation
            try:
                filename = tool.get("filename", "")
                code = tool.get("code", "")
                if not filename or not code:
                    self._chat_add_bubble("system", "Error: missing filename or code.")
                    return

                # Run auto-validator before writing
                is_valid, fixed_code = self._validate_plugin_code(code, filename)
                code = fixed_code

                self._write_plugin_file(filename, code)
                self._loader.reload_plugins()
                self.after(500, self._refresh_plugins_list)
                self._chat_add_bubble("system",
                    f"Plugin created and loaded. Find it in the Plugins tab.")
                self.after(500, self._switch_to_plugins)
            except Exception as e:
                self._chat_add_bubble("system", f"Error creating plugin: {e}")

        elif tool_name == "update_setting":
            try:
                key = tool.get("key", "")
                value = tool.get("value", "")
                set_setting(key, value)
                if key == "anthropic_api_key":
                    self._loader.set_claude()
                self._chat_add_bubble("system",
                    f"Updated setting: {key} = {value[:30]}{'...' if len(value) > 30 else ''}")
            except Exception as e:
                self._chat_add_bubble("system", f"Error updating setting: {e}")

        elif tool_name == "clarify":
            question = tool.get("question", "Could you provide more detail?")
            self._chat_add_bubble("assistant", question)

    def _chat_show_examples(self):
        examples = (
            "Example prompts you can try:\n\n"
            "— TIER 1 (Simple Templates) —\n"
            "1. \"When I get an email from tony@client.com, "
            "reply with a standard acknowledgement.\"\n"
            "2. \"When anyone asks about tax return fees, "
            "send our pricing email.\"\n"
            "3. \"When Elio gets an email from the ATO, "
            "draft a smart reply based on what they're asking.\"\n"
            "4. \"Forward all emails from tony@client.com to "
            "harry@mcands.com.au and save in the Tony Phillips folder.\"\n\n"
            "— TIER 2 (Custom Plugins — Full Power) —\n"
            "5. \"When a client emails asking about their tax return, "
            "look them up in XPM and send a Teams alert to the team.\"\n"
            "6. \"Every morning, check for overdue jobs in XPM and "
            "send a summary email to the team.\"\n"
            "7. \"When I process an NOA, remember the outcome and "
            "use it next time I email that client.\"\n"
            "8. \"Build a plugin that tracks all correspondence with "
            "clients and logs it to a database.\""
        )
        self._chat_add_bubble("assistant", examples)

    def _switch_to_plugins(self):
        """Programmatically navigate to the Plugins tab."""
        self._show_page("plugins")

    # ────────────────────────────────────────────────────────────────────────
    # Navigation / Auth / Log
    # ────────────────────────────────────────────────────────────────────────

    def _show_page(self, key):
        for page in self._pages.values():
            page.pack_forget()
        self._pages[key].pack(fill="both", expand=True)
        for k, btn in self._nav_btns.items():
            btn.configure(fg_color="#1565C0" if k == key else "transparent")

    def _try_restore_session(self):
        # If already authenticated (e.g. from setup wizard), just update the UI
        if self._graph and self._graph.is_authenticated():
            self._on_auth_success()
            return
        # Otherwise try to restore from hardcoded credentials
        self._graph = GraphClient()
        if self._graph.is_authenticated():
            self._on_auth_success()

    def _on_auth_success(self):
        self.auth_status_label.configure(text="Connected", text_color="#66BB6A")
        self._loader.set_graph(self._graph)
        self._log("Connected to Microsoft 365.")

    def _on_plugin_run_complete(self, plugin_id, result):
        self._session_actions += result.actions_taken
        self._session_drafts  += result.drafts_created
        self.after(0, self._update_dashboard_stats)

        # Send tray notification if drafts were created and window is hidden
        if result.drafts_created > 0 and not self.winfo_viewable():
            self.send_tray_notification(
                "MC & S CoWorker",
                f"{result.drafts_created} new draft(s) ready for review in Outlook."
            )

    def _update_dashboard_stats(self):
        self._stat_labels["actions"].configure(text=str(self._session_actions))
        self._stat_labels["drafts"].configure(text=str(self._session_drafts))

    # ────────────────────────────────────────────────────────────────────────
    # System Tray
    # ────────────────────────────────────────────────────────────────────────

    def _on_close_requested(self):
        """Called when the user clicks the X button. Minimise to tray if available."""
        if HAS_TRAY:
            self._minimise_to_tray()
        else:
            if messagebox.askyesno(
                "Close MC & S CoWorker",
                "Closing the window will stop the scheduler.\n\n"
                "Are you sure you want to quit?"
            ):
                self._quit_app()

    def _minimise_to_tray(self):
        """Hide the window and show a system tray icon."""
        self.withdraw()  # Hide the window

        if self._tray_icon is not None:
            return  # Already in tray

        # Load the app icon for the tray
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon.png")
        if os.path.exists(icon_path):
            tray_image = PILImage.open(icon_path)
        else:
            # Fallback: create a simple blue square icon
            tray_image = PILImage.new("RGB", (64, 64), BRAND_BLUE)

        menu = pystray.Menu(
            pystray.MenuItem("Open MC & S CoWorker", self._restore_from_tray, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Scheduler Running", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit_from_tray),
        )

        self._tray_icon = pystray.Icon(
            "mcs_coworker",
            tray_image,
            "MC & S CoWorker",
            menu,
        )

        self._tray_thread = threading.Thread(target=self._tray_icon.run, daemon=True)
        self._tray_thread.start()

    def _restore_from_tray(self, icon=None, item=None):
        """Restore the window from the system tray."""
        if self._tray_icon is not None:
            self._tray_icon.stop()
            self._tray_icon = None
            self._tray_thread = None

        self.after(0, self._do_restore)

    def _do_restore(self):
        """Restore window on the main thread."""
        self.deiconify()  # Show the window
        self.lift()       # Bring to front
        self.focus_force()

    def _quit_from_tray(self, icon=None, item=None):
        """Quit the app from the tray menu."""
        if self._tray_icon is not None:
            self._tray_icon.stop()
            self._tray_icon = None
        self.after(0, self._quit_app)

    def _quit_app(self):
        """Fully shut down the application."""
        try:
            if hasattr(self, '_loader') and self._loader:
                self._loader.stop_scheduler()
        except Exception:
            pass
        self.destroy()

    def send_tray_notification(self, title: str, message: str):
        """Show a notification bubble from the system tray icon."""
        if self._tray_icon is not None and HAS_TRAY:
            try:
                self._tray_icon.notify(message, title)
            except Exception:
                pass  # Not all platforms support notifications

    def _log(self, message: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {message}\n"

        def _write():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", line)
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        self.after(0, _write)


if __name__ == "__main__":
    app = App()
    app.mainloop()
