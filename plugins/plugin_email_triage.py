"""
MC & S Plugin: Email Triage & Auto-Response
=============================================
Plugin ID  : plugin_email_triage
Version    : 1.0.0

WHAT IT DOES
------------
Monitors a nominated Outlook inbox folder for unread client emails.
Uses Claude to classify each email into a category (Pricing Enquiry,
Checklist Request, Documents Received, or Other). For matched categories
it either:
  - Draft mode ON  → Creates a draft reply + notifies staff to review
  - Draft mode OFF → Sends the reply automatically

Emails classified as OTHER are left untouched in the inbox.
Documents Received emails are also flagged for follow-up.

CONFIGURATION
-------------
Uses the Email Rules tab in the main app to define categories,
keywords, and response templates. No additional plugin settings needed.

SCHEDULE
--------
Default: every 60 seconds (matches the polling_interval setting).
Adjustable from the Plugins tab.
"""

import json
import re
from datetime import datetime

import anthropic

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_rules, get_staff, get_setting, log_activity, get_style_preferences, get_active_lessons, get_links_as_dict


class EmailTriagePlugin(AgentPlugin):

    name        = "Email Triage & Auto-Response"
    description = "Monitors inbox, classifies emails with Claude, drafts or sends replies."
    detail      = (
        "Reads unread emails from your nominated Outlook folder. Claude classifies each "
        "email against your configured keyword rules. Matching emails get a tailored "
        "draft or auto-send response. Unmatched emails (classified as OTHER) are left "
        "in your inbox for staff to handle manually."
    )
    version = "1.0.0"
    icon    = "📧"

    requires_graph  = True
    requires_claude = True

    default_schedule = Schedule.every_minutes(1)

    # Track IDs we've already processed this session to avoid double-handling
    _processed_ids: set

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()

        if not context.graph:
            context.log("📧 Email Triage: Microsoft 365 not connected.")
            return False
        if not context.claude:
            context.log("📧 Email Triage: Anthropic API key not configured.")
            return False

        return True

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {
                "key": "folder",
                "label": "Folder to Monitor",
                "default": "Inbox",
                "type": "text",
                "help": "The Outlook folder name to watch for unread emails."
            },
            {
                "key": "max_per_run",
                "label": "Max Emails Per Run",
                "default": "25",
                "type": "number",
                "help": "Maximum emails to process each time the plugin runs."
            },
        ]

    def run(self, context: PluginContext) -> PluginResult:
        graph      = context.graph
        claude     = context.claude
        log        = context.log
        draft_mode = context.draft_mode

        folder    = self.get_plugin_setting("folder", get_setting("monitor_folder", "Inbox"))
        max_count = int(self.get_plugin_setting("max_per_run", "25"))

        rules         = get_rules()
        enabled_rules = [r for r in rules if r.get("enabled")]

        if not enabled_rules:
            log("  No enabled email rules — nothing to do.")
            return PluginResult(
                success=True, summary="No enabled rules.", items_skipped=0
            )

        try:
            emails = graph.fetch_unread_emails(folder=folder, max_count=max_count)
        except Exception as e:
            return PluginResult(success=False, error=f"Could not fetch emails: {e}")

        log(f"  {len(emails)} unread email(s) in {folder}.")

        result = PluginResult(success=True)

        for email in emails:
            msg_id     = email["id"]
            if msg_id in self._processed_ids:
                continue

            subject    = email.get("subject", "(No Subject)")
            from_email = (
                email.get("from", {}).get("emailAddress", {}).get("address", "")
            )
            body_text  = email.get("body", {}).get("content", "")
            body_plain = re.sub(r"<[^>]+>", " ", body_text)
            body_plain = re.sub(r"\s+", " ", body_plain).strip()

            log(f'  Classifying: "{subject}" from {from_email}')

            try:
                classification = self._classify(
                    claude, subject, body_plain, enabled_rules
                )
                category    = classification.get("category", "OTHER")
                sender_name = classification.get("sender_name", "there")
                confidence  = classification.get("confidence", "medium")

                log(f"    ↳ {category} ({confidence})")

                if category == "OTHER":
                    log("    ↳ Left in inbox — no rule matched.")
                    log_activity(from_email, subject, category, "no_action")
                    self._processed_ids.add(msg_id)
                    result.items_skipped += 1
                    continue

                matching_rule = next(
                    (r for r in enabled_rules if r["category"] == category), None
                )
                if not matching_rule:
                    continue

                reply_subject = self._apply_template(
                    matching_rule["subject_template"], sender_name, subject
                )
                reply_body = self._apply_template(
                    matching_rule["body_template"], sender_name, subject
                )

                if draft_mode:
                    graph.create_draft(
                        from_email, reply_subject, reply_body, msg_id
                    )
                    log("    ↳ Draft created in Drafts folder.")
                    self._send_staff_notification(
                        context, from_email, subject, category
                    )
                    log_activity(
                        from_email, subject, category, "draft_created",
                        draft_created=1, notification_sent=1,
                    )
                    result.drafts_created += 1
                else:
                    graph.send_email(
                        from_email, reply_subject, reply_body, msg_id
                    )
                    log("    ↳ Reply sent.")
                    log_activity(from_email, subject, category, "auto_sent")

                if category == "DOCUMENTS_RECEIVED":
                    graph.flag_email(msg_id)
                    log("    ↳ Flagged for follow-up.")

                graph.mark_as_read(msg_id)
                self._processed_ids.add(msg_id)
                result.actions_taken += 1

            except json.JSONDecodeError as e:
                log(f"    ↳ ⚠ Unexpected Claude response format: {e}")
            except Exception as e:
                log(f"    ↳ Error: {e}")

        result.summary = (
            f"{result.actions_taken} processed, "
            f"{result.drafts_created} drafted, "
            f"{result.items_skipped} skipped."
        )
        return result

    # ── Private helpers ───────────────────────────────────────────────────────

    def _classify(self, claude_client: anthropic.Anthropic,
                  subject: str, body: str, rules: list) -> dict:

        categories_desc = "\n".join([
            f"- {r['category']}: Keywords include: {r['keywords']}"
            for r in rules if r.get("enabled")
        ])

        # ── Inject memory context ──
        memory_block = ""
        style_prefs = get_style_preferences()
        if style_prefs:
            memory_block += f"\n\nIMPORTANT — TONE & STYLE INSTRUCTIONS FROM THE USER:\n{style_prefs}\n"

        lessons = get_active_lessons()
        if lessons:
            memory_block += "\nLEARNED PREFERENCES (apply these to all responses):\n"
            memory_block += "\n".join(f"- {l['lesson']}" for l in lessons)
            memory_block += "\n"

        prompt = f"""You are an email classifier for MC & S, an accounting firm in Keysborough, Victoria.

Classify the email below into one of these categories, or OTHER if none fit:

{categories_desc}
- OTHER: Anything not listed above (meeting requests, complaints, ATO notices, etc.)

Also extract the sender's first name from any sign-off in the body. If not found, return "there".
{memory_block}
Subject: {subject}
Body: {body[:1500]}

Respond ONLY with valid JSON:
{{
  "category": "PRICING_ENQUIRY",
  "confidence": "high",
  "reasoning": "Asks about fees.",
  "sender_name": "John"
}}"""

        response = claude_client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )

        text = response.content[0].text.strip()
        text = re.sub(r"```json\s*|```", "", text).strip()
        return json.loads(text)

    def _apply_template(self, template: str, sender_name: str,
                        subject: str) -> str:
        result = template.replace("{client_name}", sender_name or "there")
        result = result.replace("{subject}", subject or "")
        result = result.replace("{date}", datetime.now().strftime("%d %B %Y"))
        # Inject all dynamic links from Links & Forms manager
        links = get_links_as_dict()
        for tag, url in links.items():
            result = result.replace(f"{{{tag}}}", url)
        return result

    def _send_staff_notification(self, context: PluginContext,
                                 client_email: str, original_subject: str,
                                 category: str):
        staff      = get_staff()
        notifiable = [s for s in staff if s.get("receives_drafts")]

        if not notifiable:
            context.log("    ↳ ⚠ No staff configured for draft notifications.")
            return

        practice = get_setting("practice_name", "MC & S")

        cat_friendly = {
            "PRICING_ENQUIRY":    "Pricing Enquiry",
            "CHECKLIST_REQUEST":  "Checklist Request",
            "DOCUMENTS_RECEIVED": "Documents Received",
        }.get(category, category.replace("_", " ").title())

        subject = f"[DRAFT READY] {cat_friendly} — Response awaiting your review"

        body = f"""
<div style="font-family:Arial,sans-serif;max-width:600px">
  <div style="background:#1565C0;color:white;padding:16px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;font-size:18px">📝 Draft Email Ready for Review</h2>
  </div>
  <div style="background:#f5f5f5;padding:24px;border-radius:0 0 8px 8px;border:1px solid #ddd">
    <p>A draft response has been prepared and is waiting in your <strong>Drafts folder</strong> in Outlook.</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0">
      <tr>
        <td style="padding:8px;color:#555;width:120px"><strong>From:</strong></td>
        <td style="padding:8px">{client_email}</td>
      </tr>
      <tr style="background:#fff">
        <td style="padding:8px;color:#555"><strong>Subject:</strong></td>
        <td style="padding:8px">{original_subject}</td>
      </tr>
      <tr>
        <td style="padding:8px;color:#555"><strong>Category:</strong></td>
        <td style="padding:8px">
          <span style="background:#E3F2FD;color:#1565C0;padding:2px 8px;
                       border-radius:12px;font-size:13px">{cat_friendly}</span>
        </td>
      </tr>
    </table>
    <div style="background:#FFF9C4;border-left:4px solid #F9A825;padding:12px 16px;
                margin:16px 0;border-radius:4px">
      <strong>Action Required:</strong> Open Outlook → Drafts, review, personalise if needed, then send.
    </div>
    <p style="color:#888;font-size:12px;margin-top:24px">
      — {practice} Desktop Agent · Email Triage Plugin
    </p>
  </div>
</div>"""

        for s in notifiable:
            try:
                context.graph.send_email(s["email"], subject, body)
                context.log(f"    ↳ Notified {s['name']}")
            except Exception as e:
                context.log(f"    ↳ ⚠ Could not notify {s['name']}: {e}")
