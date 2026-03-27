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

import requests

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_rules, get_setting, log_activity, get_links_as_dict

# Proxy server URL — classification requests are routed here
PROXY_URL = "http://134.199.150.35:8000"


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
    requires_claude = False

    default_schedule = Schedule.every_minutes(1)

    # Track IDs we've already processed this session to avoid double-handling
    _processed_ids: set

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()

        if not context.graph:
            context.log("📧 Email Triage: Microsoft 365 not connected.")
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
                    subject, body_plain, enabled_rules
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

                # Append the user's Outlook signature
                signature = graph.get_signature_html()
                if signature:
                    reply_body = reply_body + "<br>" + signature

                # Check if a signature image is uploaded for inline embedding
                sig_image_path = graph.get_signature_image_path()

                if draft_mode:
                    if sig_image_path:
                        graph.create_draft_with_inline_image(
                            from_email, reply_subject, reply_body,
                            sig_image_path, "signature_image", msg_id
                        )
                    else:
                        graph.create_draft(
                            from_email, reply_subject, reply_body, msg_id
                        )
                    log("    ↳ Draft created in Drafts folder.")
                    log_activity(
                        from_email, subject, category, "draft_created",
                        draft_created=1,
                    )
                    result.drafts_created += 1
                else:
                    if sig_image_path:
                        graph.send_email_with_inline_image(
                            from_email, reply_subject, reply_body,
                            sig_image_path, "signature_image", msg_id
                        )
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

    def _classify(self, subject: str, body: str, rules: list) -> dict:
        """Classify an email via the proxy server."""
        payload = {
            "email_subject": subject,
            "email_body": body[:1500],
            "rules": [
                {"category": r["category"], "keywords": r["keywords"]}
                for r in rules if r.get("enabled")
            ],
        }

        resp = requests.post(
            f"{PROXY_URL}/classify",
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

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

