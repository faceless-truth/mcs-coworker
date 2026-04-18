"""
MC & S Plugin: Email Reply (Unified)
======================================
Plugin ID  : plugin_email_reply
Version    : 1.0.0

WHAT IT DOES
------------
Unified email reply plugin that replaces:
  - plugin_email_triage (template-based classification and auto-reply)
  - plugin_elio_draft_replies (Claude deep-reasoning drafts)
  - plugin_auto_reply_ross (rule-based auto-reply)

Each installation is configured with a STAFF PROFILE that controls how
the plugin behaves for that person's inbox.

STAFF PROFILES
--------------
  reception  — Template-based. Classifies using keyword rules, drafts or
               sends template replies. Tier 1 categories can auto-send.
               Mirrors the original plugin_email_triage behaviour.

  elio       — Deep mode. Uses Claude Sonnet + client memory to write a
               personalised draft reply for every email that needs a
               response. Always creates a draft (never auto-sends).

  ross       — Rule-based. Matches sender/subject patterns, sends a
               configurable auto-reply. Respects draft_mode.

  harry      — Deep mode (same as elio). Always drafts.

  brooke     — Deep mode (same as elio). Always drafts.

  louise     — Deep mode (same as elio). Always drafts.

  lyn        — Deep mode (same as elio). Always drafts.

CONFIGURATION
-------------
  staff_profile   — Set in main config (reception | elio | ross | harry |
                    brooke | louise | lyn). Default: elio.
  folder          — Outlook folder to monitor. Default: Inbox.
  polling_minutes — How often to poll. Default: 1.
  auto_reply_senders — (ross profile only) Comma-separated list of sender
                       email addresses to auto-reply to.
  auto_reply_body — (ross profile only) HTML body for the auto-reply.

SCHEDULE
--------
Default: every 1 minute. Adjustable from the Plugins tab.
"""

import json
import re
from datetime import datetime
import requests
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import (
    get_rules, get_setting, log_activity, get_links_as_dict,
    get_active_lessons,
)

# Proxy server URL — classification requests are routed here (reception profile)
PROXY_URL = "http://134.199.150.35:8000"

# Profiles that use Claude deep-reasoning drafts
DEEP_PROFILES = {"elio", "harry", "brooke", "louise", "lyn"}

# No-reply address patterns — skip these for all profiles
NO_REPLY_PATTERNS = [
    "noreply@", "no-reply@", "donotreply@", "do-not-reply@",
    "unsubscribe@", "notification@", "alerts@",
    "mailer-daemon@", "postmaster@",
]

# Tier 1 categories — high-confidence template responses that can auto-send
TIER_1_CATEGORIES = {"CHECKLIST_REQUEST", "PRICING_ENQUIRY"}

# EVA processing tag applied while an email is being handled
EVA_CATEGORY_TAG = "EVA Processing"


class EmailReplyPlugin(AgentPlugin):
    """Unified email reply plugin with named staff profiles."""

    PLUGIN_ID   = "plugin_email_reply"
    NAME        = "Email Reply"
    DESCRIPTION = (
        "Unified email reply plugin. Monitors the inbox and drafts or sends replies "
        "based on the configured staff profile (reception, elio, ross, harry, brooke, "
        "louise, lyn). Reception uses template classification; accountant profiles use "
        "Claude deep-reasoning drafts."
    )
    VERSION     = "1.0.0"
    ICON        = "📧"
    requires_graph  = True
    requires_claude = True
    default_schedule = Schedule.every_minutes(1)

    # Track IDs processed this session to avoid double-handling
    _processed_ids: set

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        if not context.graph:
            context.log("[EmailReply] Microsoft 365 not connected.")
            return False
        return True

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {
                "key":     "folder",
                "label":   "Folder to Monitor",
                "default": "Inbox",
                "type":    "text",
                "help":    "Outlook folder name to watch for unread emails.",
            },
            {
                "key":     "polling_minutes",
                "label":   "Polling Interval (minutes)",
                "default": "1",
                "type":    "number",
                "help":    "How often to check for new emails.",
            },
            {
                "key":     "auto_reply_senders",
                "label":   "Auto-Reply Senders (ross profile)",
                "default": "ross@mcands.com.au",
                "type":    "text",
                "help":    "Comma-separated sender addresses to auto-reply to (ross profile only).",
            },
            {
                "key":     "auto_reply_body",
                "label":   "Auto-Reply Body (ross profile)",
                "default": "<p>Thank you for your email. I will get back to you shortly.</p>",
                "type":    "textarea",
                "help":    "HTML body for the auto-reply (ross profile only).",
            },
            {
                "key":     "deep_reply_max_tokens",
                "label":   "Deep Reply Max Tokens",
                "default": "400",
                "type":    "number",
                "help":    "Max tokens for Claude deep-reasoning draft replies.",
            },
        ]

    # ── Main entry point ──────────────────────────────────────────────────────

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()
        profile = get_setting("staff_profile", "elio").lower().strip()
        folder  = self.get_plugin_setting("folder", "Inbox")
        graph   = context.graph

        try:
            emails = graph.fetch_unread_emails(folder, max_count=25)
        except Exception as e:
            result.summary = f"Could not fetch emails: {e}"
            return result

        if not emails:
            result.summary = "No unread emails."
            return result

        context.log(f"[EmailReply] Profile: {profile}. {len(emails)} unread email(s).")

        if profile == "reception":
            return self._run_reception(context, result, emails)
        elif profile == "ross":
            return self._run_ross(context, result, emails)
        elif profile in DEEP_PROFILES:
            return self._run_deep(context, result, emails, profile)
        else:
            result.summary = f"Unknown staff_profile '{profile}'. Check config."
            return result

    # ── Reception profile (template classification) ───────────────────────────

    def _run_reception(self, context: PluginContext, result: PluginResult,
                       emails: list) -> PluginResult:
        """
        Template-based classification and reply.
        Mirrors the original plugin_email_triage behaviour.
        """
        rules = get_rules()
        if not rules:
            result.summary = "No email rules configured. Add rules in the Email Rules tab."
            return result

        graph      = context.graph
        draft_mode = context.draft_mode

        for email in emails:
            msg_id  = email.get("id")
            subject = email.get("subject", "(no subject)")
            sender  = email.get("from", {}).get("emailAddress", {})
            from_email   = sender.get("address", "")
            sender_name  = sender.get("name", from_email.split("@")[0])
            body_preview = email.get("bodyPreview", "")

            if msg_id in self._processed_ids:
                result.items_skipped += 1
                continue
            if self._is_no_reply(from_email):
                result.items_skipped += 1
                continue

            context.log(f"[EmailReply] Reception: {from_email} — {subject}")

            try:
                graph.add_category(msg_id, EVA_CATEGORY_TAG)
            except Exception:
                pass

            try:
                classification = self._classify(subject, body_preview, rules)
                category    = classification.get("category", "OTHER")
                confidence  = float(classification.get("confidence", 0.0))
            except Exception as e:
                context.log(f"[EmailReply] Classification failed: {e}")
                result.items_skipped += 1
                continue

            if category == "OTHER":
                context.log(f"    ↳ Classified as OTHER — leaving in inbox.")
                result.items_skipped += 1
                try:
                    graph.remove_category(msg_id, EVA_CATEGORY_TAG)
                except Exception:
                    pass
                continue

            matching_rule = next(
                (r for r in rules if r.get("category") == category and r.get("enabled")),
                None,
            )
            if not matching_rule:
                result.items_skipped += 1
                continue

            reply_subject = f"Re: {subject}"
            reply_body    = self._apply_template(
                matching_rule["body_template"], sender_name, subject
            )
            signature = graph.get_signature_html()
            if signature:
                reply_body += "<br>" + signature

            sig_image_path = graph.get_signature_image_path()

            tier1_auto_send = (
                not draft_mode
                or (
                    category in TIER_1_CATEGORIES
                    and context.approval_queue is not None
                    and context.approval_queue.get_threshold() <= confidence
                )
            )

            if not tier1_auto_send:
                if sig_image_path:
                    graph.create_draft_with_inline_image(
                        from_email, reply_subject, reply_body,
                        sig_image_path, "signature_image", msg_id,
                    )
                else:
                    graph.create_draft(from_email, reply_subject, reply_body, msg_id)
                context.log("    ↳ Draft created.")
                log_activity(from_email, subject, category, "draft_created", draft_created=1)
                result.drafts_created += 1
            else:
                if sig_image_path:
                    graph.send_email_with_inline_image(
                        from_email, reply_subject, reply_body,
                        sig_image_path, "signature_image", msg_id,
                    )
                else:
                    graph.send_email(from_email, reply_subject, reply_body, msg_id)
                context.log("    ↳ Reply sent.")
                log_activity(from_email, subject, category, "auto_sent")
                result.actions_taken += 1

            if category == "DOCUMENTS_RECEIVED":
                try:
                    graph.flag_email(msg_id)
                except Exception:
                    pass

            try:
                graph.remove_category(msg_id, EVA_CATEGORY_TAG)
                graph.add_category(msg_id, category.replace("_", " ").title())
            except Exception:
                pass

            graph.mark_as_read(msg_id)
            self._processed_ids.add(msg_id)

        result.summary = (
            f"{result.actions_taken} sent, "
            f"{result.drafts_created} drafted, "
            f"{result.items_skipped} skipped."
        )
        return result

    # ── Ross profile (rule-based auto-reply) ──────────────────────────────────

    def _run_ross(self, context: PluginContext, result: PluginResult,
                  emails: list) -> PluginResult:
        """
        Rule-based auto-reply for a specific set of sender addresses.
        Respects draft_mode.
        """
        graph      = context.graph
        draft_mode = context.draft_mode

        senders_setting = self.get_plugin_setting(
            "auto_reply_senders", "ross@mcands.com.au"
        )
        target_senders = {
            s.strip().lower()
            for s in senders_setting.split(",")
            if s.strip()
        }
        reply_body = self.get_plugin_setting(
            "auto_reply_body",
            "<p>Thank you for your email. I will get back to you shortly.</p>",
        )

        for email in emails:
            msg_id  = email.get("id")
            subject = email.get("subject", "(no subject)")
            sender  = email.get("from", {}).get("emailAddress", {})
            from_email = sender.get("address", "").lower()

            if msg_id in self._processed_ids:
                result.items_skipped += 1
                continue
            if from_email not in target_senders:
                result.items_skipped += 1
                continue
            if self._is_no_reply(from_email):
                result.items_skipped += 1
                continue

            context.log(f"[EmailReply] Ross: matched {from_email} — {subject}")

            if draft_mode:
                graph.create_draft(from_email, f"Re: {subject}", reply_body, msg_id)
                context.log("    ↳ Draft created (draft_mode=True).")
                result.drafts_created += 1
            else:
                graph.send_email(from_email, f"Re: {subject}", reply_body, msg_id)
                context.log("    ↳ Auto-reply sent.")
                result.actions_taken += 1

            graph.mark_as_read(msg_id)
            self._processed_ids.add(msg_id)

        result.summary = (
            f"{result.actions_taken} sent, "
            f"{result.drafts_created} drafted, "
            f"{result.items_skipped} skipped."
        )
        return result

    # ── Deep profile (Claude reasoning drafts) ────────────────────────────────

    def _run_deep(self, context: PluginContext, result: PluginResult,
                  emails: list, profile: str) -> PluginResult:
        """
        Claude Sonnet deep-reasoning draft for every email that needs a reply.
        Always creates a draft — never auto-sends.
        """
        graph = context.graph

        for email in emails:
            msg_id  = email.get("id")
            subject = email.get("subject", "(no subject)")
            sender  = email.get("from", {}).get("emailAddress", {})
            from_email  = sender.get("address", "")
            sender_name = sender.get("name", from_email.split("@")[0])

            if msg_id in self._processed_ids:
                result.items_skipped += 1
                continue
            if self._is_no_reply(from_email):
                result.items_skipped += 1
                continue

            context.log(f"[EmailReply] {profile.title()}: {from_email} — {subject}")

            # Fetch full email body
            try:
                full_email = self._fetch_full_email(context, msg_id)
                body_text  = self._extract_body_text(full_email)
            except Exception as e:
                context.log(f"    ↳ Could not fetch full email: {e}")
                body_text = email.get("bodyPreview", "")

            # Draft reply with Claude
            try:
                draft_html = self._draft_with_claude(
                    context, profile, sender_name, from_email, subject, body_text
                )
            except Exception as e:
                context.log(f"    ↳ Claude draft failed: {e}")
                draft_html = self._fallback_draft(profile, sender_name, subject)

            # Append Outlook signature
            try:
                signature = graph.get_signature_html()
                if signature:
                    draft_html += "<br>" + signature
            except Exception:
                pass

            # Create draft
            try:
                graph.create_draft(
                    to=from_email,
                    subject=f"Re: {subject}",
                    body_html=draft_html,
                    reply_to_id=msg_id,
                )
                context.log(f"    ↳ Draft created.")
                log_activity(from_email, subject, "deep_draft", "draft_created", draft_created=1)
                result.drafts_created += 1
            except Exception as e:
                context.log(f"    ↳ Draft creation failed: {e}")

            graph.mark_as_read(msg_id)
            self._processed_ids.add(msg_id)

        result.summary = (
            f"{result.drafts_created} draft(s) created, "
            f"{result.items_skipped} skipped."
        )
        return result

    # ── Claude helpers ────────────────────────────────────────────────────────

    def _draft_with_claude(self, context: PluginContext, profile: str,
                           sender_name: str, sender_email: str,
                           subject: str, body: str) -> str:
        """Use Claude Sonnet to draft a personalised reply."""
        practice_name = get_setting("practice_name", "MC & S Accounting")
        user_name     = get_setting("user_name", profile.title())
        lessons       = get_active_lessons()
        lessons_text  = ""
        if lessons:
            lessons_text = "\nLearned preferences:\n" + \
                           "\n".join(f"- {l['lesson']}" for l in lessons[:5])

        # Pull client history from memory if available
        client_context = ""
        if context.memory:
            try:
                results = context.memory.search(
                    query=f"{sender_name} {sender_email} {subject}",
                    n_results=3,
                    collection="client_interactions",
                )
                if results:
                    snippets = [r.get("document", "")[:200] for r in results]
                    client_context = "\nClient history:\n" + "\n".join(snippets)
            except Exception:
                pass

        max_tokens = int(self.get_plugin_setting("deep_reply_max_tokens", "400"))

        prompt = f"""You are drafting a professional email reply on behalf of {user_name} \
at {practice_name}, an Australian accounting firm.

Sender: {sender_name} <{sender_email}>
Subject: {subject}
Email:
{body[:2500]}
{client_context}
{lessons_text}

Draft a professional, concise reply that:
1. Addresses the sender by first name
2. Directly responds to their question or request
3. Provides clear next steps where relevant
4. Uses a professional but warm tone appropriate for an accounting practice
5. Is 3–6 sentences — no padding, no filler

Respond with ONLY the reply body text. No subject line. No salutation. No sign-off."""

        if not context.claude_reason:
            raise RuntimeError("Claude not available")

        response = context.claude_reason.messages.create(
            model=self.get_claude_model_reasoning(),
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
        )
        reply_text = response.content[0].text.strip()

        # Wrap in minimal HTML
        paragraphs = reply_text.split("\n\n")
        html_paras = "".join(f"<p>{p.replace(chr(10), '<br>')}</p>" for p in paragraphs)
        return f"<html><body>{html_paras}</body></html>"

    def _fallback_draft(self, profile: str, sender_name: str, subject: str) -> str:
        """Generic fallback draft when Claude is unavailable."""
        first_name = sender_name.split()[0] if sender_name else "there"
        user_name  = get_setting("user_name", profile.title())
        return (
            f"<html><body>"
            f"<p>Hi {first_name},</p>"
            f"<p>Thank you for your email regarding {subject}. "
            f"I will review this and get back to you shortly.</p>"
            f"<p>Kind regards,<br>{user_name}</p>"
            f"</body></html>"
        )

    # ── Classification (reception profile) ───────────────────────────────────

    def _classify(self, subject: str, body: str, rules: list) -> dict:
        """Classify an email via the proxy server."""
        payload = {
            "email_subject": subject,
            "email_body":    body[:1500],
            "rules": [
                {"category": r["category"], "keywords": r["keywords"]}
                for r in rules if r.get("enabled")
            ],
        }
        resp = requests.post(f"{PROXY_URL}/classify", json=payload, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def _apply_template(self, template: str, sender_name: str, subject: str) -> str:
        """Fill placeholders in a template string."""
        result = template.replace("{client_name}", sender_name or "there")
        result = result.replace("{subject}", subject or "")
        result = result.replace("{date}", datetime.now().strftime("%d %B %Y"))
        links = get_links_as_dict()
        for tag, url in links.items():
            result = result.replace(f"{{{tag}}}", url)
        return result

    # ── Utility ───────────────────────────────────────────────────────────────

    def _is_no_reply(self, address: str) -> bool:
        addr = address.lower()
        return any(p in addr for p in NO_REPLY_PATTERNS)

    def _fetch_full_email(self, context: PluginContext, msg_id: str) -> dict:
        endpoint = f"/me/messages/{msg_id}"
        return context.graph._make_request("GET", endpoint)

    def _extract_body_text(self, email_data: dict) -> str:
        """Extract plain text from an email dict."""
        if not email_data:
            return ""
        body = email_data.get("body", {})
        content = body.get("content", "")
        if body.get("contentType", "").lower() == "html":
            content = re.sub(r"<[^>]+>", " ", content)
            content = re.sub(r"\s+", " ", content).strip()
        return content[:3000]
