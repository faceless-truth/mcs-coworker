"""
MC & S Plugin: Smart Email Responder
=====================================
Plugin ID  : plugin_smart_responder
Version    : 1.0.0

WHAT IT DOES
------------
Reads unread emails and asks Claude to draft a contextual reply, using the
knowledge base (pricing, checklists, procedures, firm policies) as grounding.

Claude returns either:
  - a draft email body → saved as an Outlook draft reply
  - the literal token "NO_REPLY" → email is marked as read with no reply

Always runs in draft mode — never auto-sends.

REPLACES
--------
- plugin_email_triage
- plugin_auto_reply_ross
- plugin_auto_response_elio_claude
- plugin_email_reply
- plugin_elio_draft_replies

SCHEDULE
--------
Default: every 1 minute.
"""

from __future__ import annotations

import html
import re
from typing import Iterable

from plugin_base import (
    AgentPlugin, PluginContext, PluginResult, Schedule, PluginCategory,
)
from config import (
    get_setting, log_activity, get_knowledge_entries,
)
from prompt_utils import wrap_untrusted_content, UNTRUSTED_CONTENT_SYSTEM_PROMPT


NO_REPLY_TOKEN = "NO_REPLY"


class SmartEmailResponderPlugin(AgentPlugin):
    """AI-powered email responder grounded in the practice knowledge base."""

    name        = "Smart Email Responder"
    description = (
        "AI-powered email responder — uses your knowledge base to draft "
        "contextual replies. Newsletters and receipts are skipped."
    )
    detail      = (
        "Fetches unread emails, asks Claude to classify each one, and drafts "
        "a professional reply grounded in the knowledge base (pricing, "
        "checklists, procedures, firm policies). If Claude decides a reply "
        "is not appropriate, the email is marked as read without a draft. "
        "Always creates drafts — never auto-sends."
    )
    version = "1.0.0"
    icon    = "🧠"
    author  = "MC & S"

    requires_graph  = True
    requires_claude = True
    category        = PluginCategory.UNIVERSAL

    default_schedule = Schedule.every_minutes(1)

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {"key": "max_emails_per_run", "label": "Max emails per run",
             "default": "10", "type": "number"},
            {"key": "monitor_folder", "label": "Folder to monitor",
             "default": "Inbox", "type": "text"},
        ]

    def load(self, context: PluginContext) -> bool:
        if not context.graph:
            context.log("🧠 Smart Email Responder: Microsoft 365 not connected.")
            return False
        if not (context.claude or context.claude_fast):
            context.log("🧠 Smart Email Responder: Claude not configured.")
            return False
        return True

    # ── Main entry ────────────────────────────────────────────────────────────

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()
        graph = context.graph
        claude = context.claude_fast or context.claude

        if graph is None or claude is None:
            result.success = False
            result.error = "Microsoft 365 or Claude not configured."
            return result

        folder = self.get_plugin_setting("monitor_folder", "Inbox")
        try:
            max_emails = int(self.get_plugin_setting("max_emails_per_run", "10"))
        except ValueError:
            max_emails = 10

        try:
            emails = graph.fetch_unread_emails(folder=folder, max_count=max_emails)
        except Exception as e:
            result.success = False
            result.error = f"Failed to fetch emails: {e}"
            return result

        if not emails:
            result.summary = "No unread emails."
            return result

        kb_block = self._knowledge_base_block()
        practice_name = get_setting("practice_name", "MC & S")
        user_name     = get_setting("user_name", "")
        user_firm     = get_setting("user_firm", practice_name)
        staff_profile = get_setting("staff_profile", "")

        drafted = 0
        skipped = 0
        errors  = 0
        model = self.get_claude_model_fast()

        for email in emails:
            subject = email.get("subject", "(no subject)")
            sender  = self._sender_address(email)
            message_id = email.get("id")
            body_preview = self._plain_text(
                (email.get("body") or {}).get("content", "")
                or email.get("bodyPreview", "")
            )

            try:
                reply_body = self._ask_claude(
                    claude=claude,
                    model=model,
                    subject=subject,
                    sender=sender,
                    body=body_preview,
                    kb_block=kb_block,
                    practice_name=practice_name,
                    user_name=user_name,
                    user_firm=user_firm,
                    staff_profile=staff_profile,
                )
            except Exception as e:
                context.log(f"🧠 Claude error for '{subject}': {e}")
                log_activity(sender, subject, "smart_responder",
                             f"Claude error: {e}", 0, 0)
                errors += 1
                continue

            if reply_body is None or reply_body.strip().upper() == NO_REPLY_TOKEN:
                # No reply needed — mark as read so it drops out of the queue
                try:
                    graph.mark_as_read(message_id)
                except Exception as e:
                    context.log(f"🧠 Couldn't mark read '{subject}': {e}")
                log_activity(sender, subject, "smart_responder",
                             "no_reply", 0, 0)
                skipped += 1
                continue

            reply_subject = subject if subject.lower().startswith("re:") \
                else f"Re: {subject}"
            try:
                graph.create_draft(
                    to_address=sender,
                    subject=reply_subject,
                    body_html=self._ensure_html(reply_body),
                    reply_to_id=message_id,
                )
                drafted += 1
                log_activity(sender, subject, "smart_responder",
                             "drafted", 1, 0)
            except Exception as e:
                context.log(f"🧠 Draft creation failed for '{subject}': {e}")
                log_activity(sender, subject, "smart_responder",
                             f"draft_failed: {e}", 0, 0)
                errors += 1

        result.actions_taken  = drafted + skipped
        result.drafts_created = drafted
        result.items_skipped  = skipped
        result.summary = (
            f"{drafted} drafts, {skipped} skipped, {errors} errors "
            f"across {len(emails)} emails"
        )
        if errors and not drafted and not skipped:
            result.success = False
            result.error = result.summary
        return result

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _sender_address(email: dict) -> str:
        sender = email.get("from") or {}
        addr = (sender.get("emailAddress") or {}).get("address", "")
        return addr or "unknown@unknown"

    @staticmethod
    def _plain_text(body: str) -> str:
        if not body:
            return ""
        text = re.sub(r"<[^>]+>", " ", body)
        text = html.unescape(text)
        text = re.sub(r"\s+", " ", text).strip()
        return text[:4000]

    @staticmethod
    def _ensure_html(body: str) -> str:
        """Wrap plaintext replies in simple HTML paragraphs."""
        stripped = body.strip()
        if "<" in stripped and ">" in stripped:
            return stripped
        paragraphs = [p.strip() for p in stripped.split("\n\n") if p.strip()]
        return "".join(f"<p>{html.escape(p).replace(chr(10), '<br>')}</p>"
                       for p in paragraphs)

    @staticmethod
    def _knowledge_base_block() -> str:
        """Render all enabled knowledge base entries as a structured block."""
        entries: Iterable[dict] = get_knowledge_entries(only_enabled=True)
        lines: list[str] = []
        for e in entries:
            cat  = (e.get("category") or "").strip()
            title = (e.get("title") or "").strip()
            content = (e.get("content") or "").strip()
            if not content:
                continue
            header = f"[{cat}] {title}" if cat else title
            lines.append(f"### {header}\n{content}")
        if not lines:
            return "(knowledge base is empty)"
        return "\n\n".join(lines)

    def _ask_claude(
        self,
        *,
        claude,
        model: str,
        subject: str,
        sender: str,
        body: str,
        kb_block: str,
        practice_name: str,
        user_name: str,
        user_firm: str,
        staff_profile: str,
    ) -> str | None:
        system = (
            f"You are an assistant at {practice_name}, an Australian accounting "
            f"firm. Your job is to draft a professional, helpful reply to the "
            f"email below.\n\n"
            f"Use ONLY the knowledge base below to answer questions about "
            f"pricing, checklists, procedures, and firm policies. If the "
            f"knowledge base does not contain the answer, say so politely and "
            f"offer to get back to the client after checking internally — "
            f"do NOT invent facts.\n\n"
            f"If the email does not require a reply (newsletter, "
            f"notification, receipt, automated calendar invite, delivery "
            f"notification, bounce, unsubscribe confirmation, etc.), respond "
            f"with exactly the single token: {NO_REPLY_TOKEN}\n\n"
            f"Otherwise respond with ONLY the reply body text (no subject line, "
            f"no commentary). Keep it concise, professional, Australian "
            f"English.\n\n"
            f"Do NOT include any sign-off, closing, or signature in your "
            f"draft. No Kind regards, no Best regards, no name, no company "
            f"name. The email signature is appended automatically — just end "
            f"with your last sentence of actual content.\n\n"
            f"KNOWLEDGE BASE\n"
            f"==============\n{kb_block}"
        )
        if staff_profile:
            system += f"\n\nStaff profile hint: {staff_profile}"

        # The email body is attacker-controlled — wrap it in a tag and add a
        # system-prompt addendum that marks the tag as DATA, not instructions.
        system += "\n\n" + UNTRUSTED_CONTENT_SYSTEM_PROMPT
        wrapped_body = wrap_untrusted_content(body or "(empty body)", "email_body")

        user_prompt = (
            f"From: {sender}\n"
            f"Subject: {subject}\n\n"
            f"Body:\n{wrapped_body}"
        )

        response = self.call_claude_with_retry(
            claude,
            model=model,
            max_tokens=1024,
            system=system,
            messages=[{"role": "user", "content": user_prompt}],
        )
        if response is None:
            return None
        text_parts = []
        for block in response.content:
            text = getattr(block, "text", None)
            if text:
                text_parts.append(text)
        return ("".join(text_parts)).strip() or None
