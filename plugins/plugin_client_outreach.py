"""
MC & S Plugin: Client Outreach (StatementHub Integration)
=========================================================
Plugin ID  : plugin_client_outreach
Version    : 1.0.0
Author     : MC & S

Connects to the StatementHub platform API to retrieve a queue of
entities/clients that need follow-up, then uses Claude AI to draft
personalised outreach emails based on the reason and context.

Outreach reasons include:
  - Stale financial years (no progress in 30+ days)
  - Documents needed (FY created but no records uploaded)
  - ASIC returns overdue or upcoming
  - Debtor follow-up (45+ days overdue)
  - General check-in (no activity in 6+ months)

All emails are created as Drafts in Outlook for human review.
"""

import json
import traceback
from datetime import datetime

import anthropic
import requests

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule


# ── Prompt templates for Claude ──────────────────────────────────────────────

OUTREACH_SYSTEM_PROMPT = """\
You are a professional email drafting assistant for MC & S Accountants, \
a Melbourne-based accounting firm. You draft client outreach emails on \
behalf of the assigned accountant.

RULES:
- Write in a warm, professional Australian tone
- Keep emails concise (3-5 short paragraphs max)
- Never use overly formal language — be approachable
- Do NOT include a sign-off or signature — the system appends the user's \
  real Outlook signature automatically
- Address the client by their first name if possible, otherwise use the entity name
- Reference specific details (financial year, due dates, amounts) naturally
- Always include a clear call to action
- Never mention "StatementHub" or internal systems to the client
- For debtor emails, be firm but respectful — never threatening
- For ASIC emails, emphasise the deadline and potential late fees
- For check-in emails, be genuinely warm and offer to help
"""

OUTREACH_USER_TEMPLATE = """\
Draft an outreach email for the following situation:

ENTITY: {entity_name} ({entity_type})
CONTACT EMAIL: {contact_email}
ASSIGNED ACCOUNTANT: {accountant_name}
REASON: {outreach_reason}
DETAIL: {reason_detail}
PRIORITY: {priority}

ADDITIONAL CONTEXT:
{context_json}

{memory_instructions}

Write ONLY the email body in HTML format (no subject line, no signature). \
Use <p> tags for paragraphs. Keep it professional and concise.
"""

SUBJECT_TEMPLATES = {
    "stale_financial_year": "Checking in on your {fy} — MC & S Accounting",
    "documents_needed": "Documents needed for {fy} — MC & S Accounting",
    "asic_overdue": "Urgent: ASIC {return_type} overdue — {entity_name}",
    "asic_upcoming": "Reminder: ASIC {return_type} due soon — {entity_name}",
    "debtor_followup": "Outstanding account — MC & S Accounting",
    "debtor_escalation": "Urgent: Outstanding account requires attention — MC & S Accounting",
    "general_checkin": "Checking in — MC & S Accounting",
}


class ClientOutreachPlugin(AgentPlugin):
    """
    Pulls outreach queue from StatementHub and drafts personalised
    client emails using Claude AI.
    """

    # ── Identity ──────────────────────────────────────────────────────────────

    name        = "Client Outreach"
    description = "Drafts personalised client outreach emails using data from StatementHub."
    detail      = (
        "Connects to the StatementHub platform to identify clients who need "
        "follow-up — stale financial years, missing documents, overdue ASIC "
        "returns, outstanding invoices, or just a general check-in. Uses AI "
        "to draft a personalised email for each, saved to your Drafts folder."
    )
    version = "1.0.0"
    icon    = "📬"
    author  = "MC & S"

    # ── Requirements ──────────────────────────────────────────────────────────

    requires_graph  = True
    requires_claude = True

    # ── Schedule ──────────────────────────────────────────────────────────────

    default_schedule = Schedule.daily_at(hour=7)

    # ── Plugin-specific settings ──────────────────────────────────────────────

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {
                "key": "statementhub_url",
                "label": "StatementHub URL",
                "default": "https://statementhub.com.au",
                "type": "text",
                "help": "Base URL of the StatementHub platform (no trailing slash).",
            },
            {
                "key": "statementhub_api_key",
                "label": "StatementHub API Key",
                "default": "",
                "type": "password",
                "help": "Bearer token for the StatementHub Coworker API. Ask your admin for this.",
            },
            {
                "key": "outreach_reasons",
                "label": "Outreach Reasons (comma-separated)",
                "default": "stale_financial_year,documents_needed,asic_overdue,asic_upcoming,debtor_followup,debtor_escalation,general_checkin",
                "type": "text",
                "help": "Which outreach reasons to include. Remove any you don't want.",
            },
            {
                "key": "max_drafts_per_run",
                "label": "Max Drafts Per Run",
                "default": "10",
                "type": "number",
                "help": "Maximum number of outreach drafts to create in a single run.",
            },
            {
                "key": "filter_by_my_clients",
                "label": "Only My Clients",
                "default": "1",
                "type": "bool",
                "help": "If enabled, only fetch outreach items for entities assigned to you.",
            },
        ]

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def load(self, context: PluginContext) -> bool:
        api_key = self.get_plugin_setting("statementhub_api_key")
        if not api_key:
            context.log(
                f"📬 {self.name}: No StatementHub API key configured. "
                f"Go to Plugins → Client Outreach → Settings to add it."
            )
            return False
        context.log(f"📬 {self.name}: Loaded and ready.")
        return True

    def run(self, context: PluginContext) -> PluginResult:
        try:
            return self._do_run(context)
        except Exception as e:
            context.log(f"📬 {self.name}: Error — {e}")
            traceback.print_exc()
            return PluginResult(success=False, error=str(e))

    def _do_run(self, context: PluginContext) -> PluginResult:
        # ── 1. Read settings ─────────────────────────────────────────────────
        base_url = self.get_plugin_setting("statementhub_url", "https://statementhub.com.au").rstrip("/")
        api_key = self.get_plugin_setting("statementhub_api_key")
        reasons = self.get_plugin_setting(
            "outreach_reasons",
            "stale_financial_year,documents_needed,asic_overdue,asic_upcoming,debtor_followup,debtor_escalation,general_checkin",
        )
        max_drafts = int(self.get_plugin_setting("max_drafts_per_run", "10"))
        filter_mine = self.get_plugin_setting("filter_by_my_clients", "1") == "1"

        if not api_key:
            context.log(f"📬 {self.name}: No API key — skipping.")
            return PluginResult(success=False, error="No StatementHub API key configured.")

        # ── 2. Get current user's email for filtering ────────────────────────
        accountant_email = None
        if filter_mine and context.graph:
            try:
                user_info = context.graph.get_user_info()
                accountant_email = user_info.get("mail") or user_info.get("userPrincipalName")
            except Exception:
                pass

        # ── 3. Call StatementHub API ─────────────────────────────────────────
        context.log(f"📬 {self.name}: Fetching outreach queue from StatementHub...")

        params = {
            "reasons": reasons,
            "limit": str(max_drafts * 2),  # fetch extra in case some are skipped
        }
        if accountant_email:
            params["accountant_email"] = accountant_email

        headers = {
            "Authorization": f"Bearer {api_key}",
            "Accept": "application/json",
        }

        try:
            resp = requests.get(
                f"{base_url}/api/coworker/outreach-queue/",
                params=params,
                headers=headers,
                timeout=30,
            )
            resp.raise_for_status()
        except requests.exceptions.ConnectionError:
            msg = f"Cannot connect to StatementHub at {base_url}"
            context.log(f"📬 {self.name}: {msg}")
            return PluginResult(success=False, error=msg)
        except requests.exceptions.HTTPError as e:
            msg = f"StatementHub API error: {resp.status_code} — {resp.text[:200]}"
            context.log(f"📬 {self.name}: {msg}")
            return PluginResult(success=False, error=msg)

        data = resp.json()
        items = data.get("items", [])

        if not items:
            context.log(f"📬 {self.name}: No outreach items — all clients are up to date!")
            return PluginResult(success=True, summary="No outreach needed.", actions_taken=0)

        context.log(f"📬 {self.name}: Found {len(items)} outreach items.")

        # ── 4. Check for already-drafted items (avoid duplicates) ────────────
        already_drafted = set()
        drafted_key = f"_drafted_{datetime.now().strftime('%Y-%m-%d')}"
        try:
            existing = self.get_plugin_setting(drafted_key, "")
            if existing:
                already_drafted = set(existing.split(","))
        except Exception:
            pass

        # ── 5. Load memory/style preferences for personalisation ─────────────
        from config import get_style_preferences, get_active_lessons
        style = get_style_preferences()
        lessons = get_active_lessons()

        memory_instructions = ""
        if style:
            memory_instructions += f"\nSTYLE PREFERENCES:\n{style}\n"
        if lessons:
            memory_instructions += "\nLEARNED LESSONS:\n"
            for lesson in lessons:
                memory_instructions += f"- {lesson['lesson']}\n"

        # ── 6. Draft emails ──────────────────────────────────────────────────
        drafts_created = 0
        skipped = 0

        for item in items:
            if drafts_created >= max_drafts:
                break

            entity_id = item.get("entity_id", "")
            reason = item.get("outreach_reason", "")
            dedup_key = f"{entity_id}:{reason}"

            if dedup_key in already_drafted:
                skipped += 1
                continue

            contact_email = item.get("contact_email", "")
            if not contact_email:
                skipped += 1
                continue

            context.log(
                f"📬 Drafting {reason} email for {item.get('entity_name', 'Unknown')}..."
            )

            try:
                # Generate email body with Claude
                body_html = self._generate_email_body(
                    context.claude, item, memory_instructions
                )

                # Generate subject line
                subject = self._generate_subject(item)

                # Create draft in Outlook
                if context.graph:
                    draft_result = context.graph.create_draft(
                        to_address=contact_email,
                        subject=subject,
                        body_html=body_html,
                    )

                    if draft_result:
                        drafts_created += 1
                        already_drafted.add(dedup_key)

                        self.log_activity(
                            source=self.name,
                            subject=f"{item.get('entity_name', '')} — {reason}",
                            category="client_outreach",
                            action=f"Draft created: {subject}",
                            draft_created=1,
                        )

                        context.log(
                            f"  ✅ Draft created for {item.get('entity_name', '')} "
                            f"({item.get('priority', '')} priority)"
                        )

            except Exception as e:
                context.log(f"  ❌ Failed to draft for {item.get('entity_name', '')}: {e}")
                skipped += 1

        # ── 7. Save dedup state ──────────────────────────────────────────────
        self.set_plugin_setting(drafted_key, ",".join(already_drafted))

        # ── 8. Send staff notification ───────────────────────────────────────
        if drafts_created > 0 and context.notify:
            context.notify(
                subject=f"📬 {drafts_created} outreach draft(s) ready for review",
                body=(
                    f"<p>The Client Outreach plugin has created "
                    f"<strong>{drafts_created}</strong> draft email(s) in your "
                    f"Outlook Drafts folder.</p>"
                    f"<p>Please review and send them at your convenience.</p>"
                ),
            )

        summary = (
            f"Created {drafts_created} outreach draft(s), "
            f"skipped {skipped} item(s)."
        )
        context.log(f"📬 {self.name}: Done — {summary}")

        return PluginResult(
            success=True,
            summary=summary,
            actions_taken=drafts_created + skipped,
            drafts_created=drafts_created,
            items_skipped=skipped,
        )

    # ── Private helpers ──────────────────────────────────────────────────────

    def _generate_email_body(
        self,
        claude_client: anthropic.Anthropic,
        item: dict,
        memory_instructions: str,
    ) -> str:
        """Use Claude to draft a personalised outreach email."""
        prompt = OUTREACH_USER_TEMPLATE.format(
            entity_name=item.get("entity_name", "Client"),
            entity_type=item.get("entity_type", "entity"),
            contact_email=item.get("contact_email", ""),
            accountant_name=item.get("assigned_accountant_name", "the team at MC & S"),
            outreach_reason=item.get("outreach_reason", ""),
            reason_detail=item.get("reason_detail", ""),
            priority=item.get("priority", "medium"),
            context_json=json.dumps(item.get("context", {}), indent=2),
            memory_instructions=memory_instructions,
        )

        response = claude_client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=800,
            system=OUTREACH_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )

        return response.content[0].text

    def _generate_subject(self, item: dict) -> str:
        """Generate a subject line from the template, with fallback."""
        reason = item.get("outreach_reason", "general_checkin")
        template = SUBJECT_TEMPLATES.get(reason, "Following up — MC & S Accounting")

        ctx = item.get("context", {})
        try:
            return template.format(
                entity_name=item.get("entity_name", ""),
                fy=ctx.get("financial_year", "your financial year"),
                return_type=ctx.get("return_type_display", "Annual Return"),
            )
        except (KeyError, IndexError):
            return template

    def stop(self):
        pass
