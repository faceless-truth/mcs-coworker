"""
MC & S Plugin: [YOUR PLUGIN NAME HERE]
=======================================
Plugin ID  : plugin_template
Version    : 1.0.0
Status     : TEMPLATE — Not active until renamed and registered

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
HOW TO BUILD A NEW PLUGIN FROM THIS TEMPLATE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Step 1: Copy this file
        cp plugins/plugin_template.py plugins/plugin_your_name.py

Step 2: Rename the class
        class YourPluginName(AgentPlugin):

Step 3: Fill in the identity fields (name, description, icon etc.)

Step 4: Declare what your plugin needs
        requires_graph = True   # set True if you need to read/send emails or calendar
        requires_claude = True  # set True if you need AI classification or drafting

Step 5: Set a default schedule
        default_schedule = Schedule.every_hours(4)    # runs every 4 hours
        default_schedule = Schedule.daily_at(hour=8)  # runs at 8am daily
        default_schedule = Schedule.manual_only()      # only when user clicks "Run Now"

Step 6: Declare any plugin-specific settings (shown in the Plugins tab UI)
        See settings_schema() below for examples.

Step 7: Implement run() — this is where your automation logic goes.
        - Use context.graph to read emails, calendar events, etc.
        - Use context.claude to call Claude AI
        - Use context.log() to write progress to the dashboard
        - Check context.draft_mode before any irreversible actions
        - Return a PluginResult with a summary of what happened

Step 8: Restart the app — it will auto-discover your plugin.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PLUGIN IDEAS FOR MC & S (future candidates)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  NOA Workflow
    Detect Notice of Assessment emails forwarded by clients.
    Extract the tax payable/refund amount and due date.
    Draft a personalised cover email to the client explaining the outcome.
    Flag in Xero or log to SharePoint.

  FuseSign Monitor
    Poll FuseSign API for bundles that have been sitting unsigned > X days.
    Draft a polite nudge email to the client.
    Notify the responsible staff member.

  Meeting Prep Brief
    30 minutes before a calendar appointment, pull:
      - Client's last invoice from Xero
      - Last 3 email threads from the client
      - Any open WIP items
    Compile into a one-page brief and email it to Elio.

  Monthly Invoicing Reminder
    On the 1st of each month, pull all clients on monthly retainers from Xero.
    Generate a list of invoices that need to be raised.
    Email the list to Elio or the nominated billing staff member.

  Debtor Follow-Up
    Pull debtors aged > 45 days from Xero.
    For each, draft a progressively firmer follow-up email (1st, 2nd, 3rd notice).
    Escalate to Elio if > 90 days.

  ASIC Reminder Detector
    Scan emails from ASIC for annual review reminders.
    Extract company name and due date.
    Create a calendar reminder and draft a client notification.

  Client Check-In Prompts
    Every 6 months, surface company/trust clients who haven't been contacted recently.
    Draft a check-in email for Elio to review and send.
"""

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule


class TemplatePlugin(AgentPlugin):
    """
    Template plugin — copy and rename this file to create a new automation.
    This plugin does nothing on its own; it's a documented starting point.
    """

    # ── Identity ──────────────────────────────────────────────────────────────

    name        = "New Plugin (Template)"
    description = "Copy this file to build your next automation."
    detail      = (
        "This is a template. Duplicate this file, rename the class, implement run(), "
        "and the app will auto-discover it. See the file header for full instructions."
    )
    version = "1.0.0"
    icon    = "📋"
    author  = "MC & S"

    # ── Requirements ──────────────────────────────────────────────────────────
    # Set to True if your plugin needs these services
    requires_graph  = False   # Microsoft 365: email, calendar, OneDrive
    requires_claude = False   # Anthropic Claude API: AI reasoning and drafting

    # ── Schedule ──────────────────────────────────────────────────────────────
    # How often should this plugin run automatically?
    # Options: Schedule.every_minutes(n), every_hours(n), daily_at(hour), manual_only()
    default_schedule = Schedule.manual_only()

    # ── Plugin-specific settings ──────────────────────────────────────────────

    @classmethod
    def settings_schema(cls) -> list[dict]:
        """
        Declare any settings your plugin needs.
        These are rendered automatically in the Plugins tab settings panel.
        Remove or replace these with your own fields.
        """
        return [
            {
                "key": "example_text_setting",
                "label": "Example Text Field",
                "default": "some default value",
                "type": "text",
                "help": "Description shown under the field in the UI."
            },
            {
                "key": "example_number_setting",
                "label": "Example Number (e.g. days)",
                "default": "7",
                "type": "number",
                "help": "Numeric value used in your plugin logic."
            },
            {
                "key": "example_toggle",
                "label": "Example Toggle",
                "default": "1",
                "type": "bool",
                "help": "A true/false setting."
            },
        ]

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def load(self, context: PluginContext) -> bool:
        """
        Called when the app starts. Return True if ready, False if misconfigured.
        Example: check required settings exist, test API connections.
        """
        context.log(f"📋 {self.name}: Template plugin loaded (does nothing).")
        return True

    def run(self, context: PluginContext) -> PluginResult:
        """
        Main work method. Replace the body with your automation logic.

        Tips:
          - Use context.log() to write progress to the dashboard
          - Check context.draft_mode before sending emails or taking actions
          - Return PluginResult with a summary of what happened
          - Catch exceptions and return PluginResult(success=False, error=str(e))
        """
        context.log(f"📋 {self.name}: Run triggered (template — no action taken).")

        # ── EXAMPLE: Read emails ──────────────────────────────────────────────
        # if context.graph:
        #     emails = context.graph.fetch_unread_emails(folder="Inbox", max_count=10)
        #     context.log(f"  Found {len(emails)} unread emails.")

        # ── EXAMPLE: Call Claude ──────────────────────────────────────────────
        # if context.claude:
        #     response = context.claude.messages.create(
        #         model="claude-haiku-4-5-20251001",
        #         max_tokens=200,
        #         messages=[{"role": "user", "content": "Summarise this in one sentence: ..."}]
        #     )
        #     text = response.content[0].text
        #     context.log(f"  Claude says: {text}")

        # ── EXAMPLE: Draft mode check before sending ──────────────────────────
        # if context.draft_mode:
        #     context.graph.create_draft(to, subject, body)
        #     context.log("  Draft created.")
        # else:
        #     context.graph.send_email(to, subject, body)
        #     context.log("  Email sent.")

        # ── EXAMPLE: Save plugin-specific state ───────────────────────────────
        # self.set_plugin_setting("last_run_count", str(count))

        return PluginResult(
            success=True,
            summary="Template plugin ran — no actions taken.",
            actions_taken=0,
            drafts_created=0,
        )

    def stop(self):
        """
        Called when plugin is disabled or app closes.
        Close any open connections or clean up state here.
        """
        pass
