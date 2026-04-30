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
import logging
import re
from typing import Iterable

logger = logging.getLogger(__name__)

from plugin_base import (
    AgentPlugin, PluginContext, PluginResult, Schedule, PluginCategory,
    email_identity_prompt,
)
from config import (
    get_setting, log_activity, get_knowledge_entries, get_db,
)
from prompt_utils import wrap_untrusted_content, UNTRUSTED_CONTENT_SYSTEM_PROMPT
from bas_dates import format_bas_dates_for_prompt
from client_utils import normalise_client_name


NO_REPLY_TOKEN = "NO_REPLY"

# Trigger phrases that mean we should give Claude the calculated BAS dates as
# extra context. Matched case-insensitively against subject + body.
BAS_KEYWORDS = (
    "bas",
    "business activity statement",
    "quarterly",
    "lodgement",
    "lodgment",
    "activity statement",
    "gst return",
)


def _is_bas_email(subject: str, body: str) -> bool:
    haystack = f"{subject or ''}\n{body or ''}".lower()
    return any(kw in haystack for kw in BAS_KEYWORDS)


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

    MODEL_TIER = "sonnet"

    default_schedule = Schedule.every_seconds(60)

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
        if not (context.claude_reason or context.claude):
            context.log("🧠 Smart Email Responder: Claude not configured.")
            return False
        return True

    # ── Main entry ────────────────────────────────────────────────────────────

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()
        graph = context.graph
        claude = context.claude_reason or context.claude

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

        # Drop processed-tracking rows older than 30 days so the table doesn't
        # grow indefinitely.
        try:
            self._cleanup_old_processed(days=30)
        except Exception as e:
            context.log(f"🧠 processed-table cleanup failed: {e}")

        kb_block = self._knowledge_base_block()
        practice_name = get_setting("practice_name", "MC & S")
        user_name     = get_setting("user_name", "")
        user_firm     = get_setting("user_firm", practice_name)
        staff_profile = get_setting("staff_profile", "")

        drafted = 0
        skipped = 0
        errors  = 0
        model = self.get_claude_model_reasoning()

        for email in emails:
            subject = email.get("subject", "(no subject)")
            sender  = self._sender_address(email)
            message_id = email.get("id")
            body_preview = self._plain_text(
                (email.get("body") or {}).get("content", "")
                or email.get("bodyPreview", "")
            )

            # Belt-and-braces: even if mark_as_read fails or the email comes
            # back as unread, never re-process the same Graph message id.
            if message_id and self._is_already_processed(message_id):
                skipped += 1
                continue

            attachment_info = self._build_attachment_info(graph, email, body_preview)

            try:
                reply_body = self._ask_claude(
                    claude=claude,
                    model=model,
                    subject=subject,
                    sender=sender,
                    body=body_preview,
                    attachment_info=attachment_info,
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
                self._mark_as_processed(message_id, draft_id=None, action="no_reply")
                log_activity(sender, subject, "smart_responder",
                             "no_reply", 0, 0)
                skipped += 1
                continue

            reply_subject = subject if subject.lower().startswith("re:") \
                else f"Re: {subject}"
            try:
                draft_id = graph.create_draft(
                    to_address=sender,
                    subject=reply_subject,
                    body_html=self._ensure_html(reply_body),
                    reply_to_id=message_id,
                )
                # IMMEDIATELY mark the original as read so the next 60-second
                # tick doesn't see it as unread and draft another reply. The
                # processed-tracking row below is the second layer of defence
                # in case mark_as_read fails (transient Graph error).
                try:
                    graph.mark_as_read(message_id)
                except Exception as e:
                    context.log(f"🧠 Couldn't mark read '{subject}' after drafting: {e}")
                self._mark_as_processed(message_id, draft_id=draft_id, action="drafted")
                drafted += 1
                # Visible notification path for the accountant — shows up
                # in the Activity Log / Dashboard SSE feed so they know a
                # draft is awaiting review without having to poll Outlook.
                from_block = email.get("from") or {}
                sender_name = (
                    (from_block.get("emailAddress") or {}).get("name")
                    or sender
                )
                logger.info(
                    "Draft reply created for %s — RE: %s",
                    sender_name, subject,
                )
                context.log(
                    f"📝 Draft reply created for {sender_name} — RE: {subject}"
                )
                log_activity(sender, subject, "smart_responder",
                             "drafted", 1, 0)
                # Persist a memory entry under the normalised sender name so
                # other plugins (and future AI Chat sessions) can recall it.
                client_name = None
                try:
                    memory = getattr(context, "memory", None)
                    from_block = email.get("from") or {}
                    from_name = ((from_block.get("emailAddress") or {})
                                 .get("name") or "")
                    client_name = normalise_client_name(from_name or sender)
                    if memory is not None:
                        draft_preview = self._plain_text(reply_body)[:200]
                        memory.store_client_interaction(
                            client_name=client_name,
                            client_email=sender,
                            interaction_type="email_draft",
                            summary=f"Draft reply RE: {subject} — {draft_preview}",
                            metadata={
                                "subject": subject[:200],
                                "draft_id": draft_id or "",
                            },
                        )
                except Exception as e:
                    context.log(f"🧠 memory write skipped: {e}")

                # Best-effort copy of the draft to the client's SharePoint
                # folder so the correspondence file is visible alongside
                # other client documents. Skip silently if SharePoint is
                # not configured or the upload fails — the draft itself
                # has already been created successfully.
                try:
                    if client_name:
                        from datetime import datetime as _dt
                        ts = _dt.now().strftime("%Y-%m-%d_%H%M")
                        outlook_email = get_setting("outlook_email", "elio@mcands.com.au")
                        body_text = self._plain_text(reply_body)
                        content = (
                            f"Subject: {subject}\n"
                            f"Date: {_dt.now().strftime('%d %B %Y %H:%M')}\n"
                            f"From: {outlook_email}\n"
                            f"To: {sender}\n"
                            f"Status: Draft created by CoWorker\n"
                            f"\n---\n\n"
                            f"{body_text}\n"
                        )
                        graph.upload_to_sharepoint(
                            file_content=content.encode("utf-8"),
                            filename=f"correspondence_{ts}.txt",
                            client_name=client_name,
                            subfolder="CoWorker Correspondence",
                        )
                except Exception as e:
                    logger.warning("SharePoint auto-file failed (non-fatal): %s", e)
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

    # ── Processed-message tracking ────────────────────────────────────────────

    @staticmethod
    def _is_already_processed(message_id: str) -> bool:
        """Return True if we have a row in ``smart_responder_processed`` for
        this Graph message id. Prevents duplicate drafts even if mark_as_read
        previously failed."""
        if not message_id:
            return False
        try:
            conn = get_db()
            row = conn.execute(
                "SELECT 1 FROM smart_responder_processed WHERE message_id = ?",
                (message_id,),
            ).fetchone()
            conn.close()
            return row is not None
        except Exception:
            return False

    @staticmethod
    def _mark_as_processed(message_id: str, draft_id: str | None,
                            action: str) -> None:
        """Record that we have handled this Graph message id."""
        if not message_id:
            return
        try:
            conn = get_db()
            conn.execute(
                "INSERT OR REPLACE INTO smart_responder_processed "
                "(message_id, draft_id, action, processed_at) "
                "VALUES (?, ?, ?, CURRENT_TIMESTAMP)",
                (message_id, draft_id or "", action),
            )
            conn.commit()
            conn.close()
        except Exception:
            pass

    @staticmethod
    def _cleanup_old_processed(days: int = 30) -> None:
        """Drop processed-tracking rows older than N days."""
        try:
            conn = get_db()
            conn.execute(
                "DELETE FROM smart_responder_processed "
                "WHERE processed_at < datetime('now', ?)",
                (f"-{int(days)} days",),
            )
            conn.commit()
            conn.close()
        except Exception:
            pass

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

    @staticmethod
    def _build_attachment_info(graph, email: dict, body_preview: str) -> str:
        """Return a string describing the email's attachments for Claude.

        Three cases:
          1. Email has attachments → list filename(s) and sizes, tell Claude
             they are confirmed present so it doesn't ask the sender to re-send.
          2. Body mentions attachments but message has none → tell Claude it
             may politely ask the sender to re-send. (This is the ONLY case
             where mentioning missing attachments is allowed.)
          3. Neither → empty string, Claude shouldn't mention attachments.
        """
        message_id = email.get("id")
        has_attachments = bool(email.get("hasAttachments", False))

        if has_attachments and message_id:
            try:
                attachments = graph.get_attachment_metadata(message_id)
            except Exception:
                attachments = []
            if attachments:
                att_list = ", ".join(
                    f"{att.get('name', 'unknown')} ({(att.get('size', 0) or 0) // 1024} KB)"
                    for att in attachments
                )
                return (
                    f"\n\nATTACHMENT INFO: This email has {len(attachments)} "
                    f"attachment(s): {att_list}. The attachments are confirmed "
                    f"present in the email — do NOT tell the sender their "
                    f"attachments are missing or ask them to re-send. When "
                    f"acknowledging, reference them naturally (e.g. \"Thank "
                    f"you for sending through the [filename]\")."
                )

        body_lc = (body_preview or "").lower()
        mentions_attachment = any(
            kw in body_lc
            for kw in ("attached", "attachment", "enclosed", "find attached")
        )
        if not has_attachments and mentions_attachment:
            return (
                "\n\nATTACHMENT INFO: The sender mentions attachments in their "
                "email, but no attachments were found on this message. It may "
                "be appropriate to politely ask the sender to re-send with the "
                "files attached."
            )

        return ""

    def _ask_claude(
        self,
        *,
        claude,
        model: str,
        subject: str,
        sender: str,
        body: str,
        attachment_info: str,
        kb_block: str,
        practice_name: str,
        user_name: str,
        user_firm: str,
        staff_profile: str,
    ) -> str | None:
        system = (
            email_identity_prompt()
            + "\n"
            f"You work at {practice_name}, an Australian accounting "
            f"firm. Your job is to draft a professional, helpful reply to the "
            f"email below — written AS the accountant identified above, in "
            f"the first person.\n\n"
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
            f"When referencing URLs, present them as clickable links with "
            f"context — e.g. \"You can access our Tax Return Checklist here: "
            f"<URL>\" rather than pasting a raw URL on its own line. The "
            f"system automatically converts plain-text URLs to clickable "
            f"hyperlinks, so you do not need to add HTML anchor tags.\n\n"
            f"ATTACHMENT HANDLING\n"
            f"-------------------\n"
            f"- If the email metadata confirms attachments are present, "
            f"acknowledge receipt of the files by name.\n"
            f"- Never tell the sender their attachments are missing when the "
            f"system confirms they are attached.\n"
            f"- If the email mentions attachments but no attachment metadata "
            f"is provided, you may politely ask the sender to confirm.\n"
            f"- When acknowledging attachments, reference them naturally — "
            f"e.g. \"Thank you for sending through the [filename] documents.\"\n\n"
            f"KNOWLEDGE BASE\n"
            f"==============\n{kb_block}"
        )
        if staff_profile:
            system += f"\n\nStaff profile hint: {staff_profile}"

        if _is_bas_email(subject, body):
            try:
                system += (
                    "\n\nBAS LODGEMENT DATES (calculated, current)\n"
                    "==========================================\n"
                    + format_bas_dates_for_prompt()
                    + "\n\nWhen the client asks about BAS due dates, use the "
                    "agent due dates above (we lodge under the tax agent program)."
                )
            except Exception:
                pass

        # The email body is attacker-controlled — wrap it in a tag and add a
        # system-prompt addendum that marks the tag as DATA, not instructions.
        system += "\n\n" + UNTRUSTED_CONTENT_SYSTEM_PROMPT
        wrapped_body = wrap_untrusted_content(body or "(empty body)", "email_body")

        # ATTACHMENT INFO is system-derived (Graph metadata), not user content,
        # so it sits OUTSIDE the untrusted_content wrapping. The model should
        # treat it as a trusted statement of fact about the message.
        user_prompt = (
            f"From: {sender}\n"
            f"Subject: {subject}\n\n"
            f"Body:\n{wrapped_body}"
            f"{attachment_info or ''}"
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
