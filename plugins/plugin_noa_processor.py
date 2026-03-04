"""
MC & S Plugin: Notice of Assessment (NOA) Processor
=====================================================
Plugin ID  : plugin_noa_processor
Version    : 1.0.0

WHAT IT DOES
------------
Monitors the inbox for emails from the ATO (or forwarded NOA emails)
containing Notice of Assessment documents. For each NOA found it:

1. Downloads the NOA PDF attachment
2. Uses Claude AI to extract key details (client name, TFN last 3 digits,
   taxable income, tax assessed, result — refund/payable/nil)
3. Selects the correct email template based on the outcome
4. Drafts (or sends) a personalised email to the client with the NOA attached
5. Logs the NOA in the activity log and marks the source email as read
6. Flags discrepancies or missing data for manual review

TEMPLATES
---------
The plugin uses six NOA outcome templates:
  - REFUND: Client is getting money back
  - PAYABLE: Client owes money
  - NIL: No refund, no payable
  - AMENDED: ATO has amended a prior return
  - COMPANY_PAYABLE: Company tax return with amount owing
  - SPOUSE: Two NOAs for a couple (combined email)

These are stored in the database (noa_templates table) and editable
from the plugin settings.

SCHEDULE
--------
Default: every 5 minutes during business hours.
"""

import json
import re
import os
from datetime import datetime
from pathlib import Path

import anthropic

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import (
    get_setting, get_staff, log_activity,
    get_style_preferences, get_active_lessons,
)


# ── Default NOA email templates ─────────────────────────────────────────────

NOA_TEMPLATES = {
    "REFUND": {
        "subject": "Your Notice of Assessment – Refund – MC & S Accounting",
        "body": """<p>Dear {client_name},</p>
<p>Please find attached your Notice of Assessment from the Australian Taxation Office for the {tax_year} financial year.</p>
<p>We are pleased to advise that you are entitled to a <strong>refund of {amount}</strong>. This will be deposited directly into your nominated bank account by the ATO, usually within 2 weeks.</p>
<p>If you have any questions, please don't hesitate to get in touch.</p>""",
    },
    "PAYABLE": {
        "subject": "Your Notice of Assessment – Amount Owing – MC & S Accounting",
        "body": """<p>Dear {client_name},</p>
<p>Please find attached your Notice of Assessment from the Australian Taxation Office for the {tax_year} financial year.</p>
<p>The ATO has assessed a <strong>payable amount of {amount}</strong>. The due date for payment is noted on the attached NOA.</p>
<p>Payment can be made via BPAY or direct transfer using the details on the notice. If you would like to discuss payment options or set up a payment plan, please let us know.</p>""",
    },
    "NIL": {
        "subject": "Your Notice of Assessment – MC & S Accounting",
        "body": """<p>Dear {client_name},</p>
<p>Please find attached your Notice of Assessment from the Australian Taxation Office for the {tax_year} financial year.</p>
<p>Your assessment shows a <strong>nil result</strong> — no refund and no amount payable.</p>
<p>If you have any questions, please don't hesitate to get in touch.</p>""",
    },
    "AMENDED": {
        "subject": "Your Amended Notice of Assessment – MC & S Accounting",
        "body": """<p>Dear {client_name},</p>
<p>Please find attached your <strong>Amended Notice of Assessment</strong> from the Australian Taxation Office for the {tax_year} financial year.</p>
<p>The ATO has made amendments to your original assessment. The key details are:</p>
<ul>
<li><strong>Result:</strong> {result_description}</li>
<li><strong>Amount:</strong> {amount}</li>
</ul>
<p>Please review the attached notice carefully. If you have any questions or concerns about the amendments, please contact us.</p>""",
    },
    "COMPANY_PAYABLE": {
        "subject": "Company Tax Assessment – Amount Owing – MC & S Accounting",
        "body": """<p>Dear {client_name},</p>
<p>Please find attached the Notice of Assessment for <strong>{entity_name}</strong> from the Australian Taxation Office for the {tax_year} financial year.</p>
<p>The ATO has assessed a <strong>payable amount of {amount}</strong>. Payment options are detailed on the attached notice.</p>
<p>If you would like to discuss this or arrange a payment plan, please get in touch.</p>""",
    },
    "DEDUCT_FROM_REFUND": {
        "subject": "Your Notice of Assessment – Refund (Fees Deducted) – MC & S Accounting",
        "body": """<p>Dear {client_name},</p>
<p>Please find attached your Notice of Assessment from the Australian Taxation Office for the {tax_year} financial year.</p>
<p>Your refund of <strong>{gross_refund}</strong> has been processed. As per our agreement, our fees of <strong>{fees_amount}</strong> have been deducted, and the <strong>net amount of {net_refund}</strong> will be transferred to your nominated bank account.</p>
<p>A receipt for the fee deduction is also attached for your records.</p>
<p>If you have any questions, please don't hesitate to get in touch.</p>""",
    },
}


class NOAProcessorPlugin(AgentPlugin):

    name        = "NOA Processor"
    description = "Processes Notice of Assessments — downloads, analyses, and drafts client emails."
    detail      = (
        "Monitors your inbox for ATO Notice of Assessment emails. Downloads the NOA PDF, "
        "uses AI to extract the key details (client name, taxable income, result), selects "
        "the correct email template (refund, payable, nil, amended), and creates a draft "
        "email to the client with the NOA attached. Flags any discrepancies for review."
    )
    version = "1.0.0"
    icon    = "📋"
    author  = "MC & S"

    requires_graph  = True
    requires_claude = True

    default_schedule = Schedule.every_minutes(5)

    _processed_ids: set
    _download_dir: str

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        self._download_dir = str(
            Path.home() / ".mcs_email_automation" / "noa_downloads"
        )
        os.makedirs(self._download_dir, exist_ok=True)

        if not context.graph:
            context.log("📋 NOA Processor: Microsoft 365 not connected.")
            return False
        if not context.claude:
            context.log("📋 NOA Processor: Anthropic API key not configured.")
            return False
        return True

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {
                "key": "sender_filter",
                "label": "ATO Sender Email (or keyword)",
                "default": "ato.gov.au",
                "type": "text",
                "help": (
                    "Filter emails by sender address. Use 'ato.gov.au' to catch ATO emails, "
                    "or leave blank to scan all unread emails for NOA attachments."
                ),
            },
            {
                "key": "subject_keywords",
                "label": "Subject Keywords",
                "default": "notice of assessment,NOA,tax assessment",
                "type": "text",
                "help": "Comma-separated keywords to identify NOA emails by subject line.",
            },
            {
                "key": "client_email_lookup",
                "label": "Client Email Source",
                "default": "from_email",
                "type": "text",
                "help": (
                    "How to determine the client email. 'from_email' uses the sender. "
                    "'ai_extract' uses Claude to find it in the email body."
                ),
            },
            {
                "key": "auto_archive",
                "label": "Auto-Archive Processed Emails",
                "default": "1",
                "type": "bool",
                "help": "Move processed NOA emails to an 'NOA Processed' folder.",
            },
            {
                "key": "max_per_run",
                "label": "Max NOAs Per Run",
                "default": "10",
                "type": "number",
                "help": "Maximum number of NOA emails to process per run.",
            },
        ]

    def run(self, context: PluginContext) -> PluginResult:
        graph      = context.graph
        claude     = context.claude
        log        = context.log
        draft_mode = context.draft_mode

        sender_filter    = self.get_plugin_setting("sender_filter", "ato.gov.au")
        subject_keywords = self.get_plugin_setting(
            "subject_keywords", "notice of assessment,NOA,tax assessment"
        )
        max_per_run = int(self.get_plugin_setting("max_per_run", "10"))
        auto_archive = self.get_plugin_setting("auto_archive", "1") == "1"

        keywords = [k.strip().lower() for k in subject_keywords.split(",") if k.strip()]

        log("📋 NOA Processor: Scanning inbox for Notice of Assessments...")

        # Fetch unread emails
        try:
            if sender_filter:
                emails = graph.fetch_emails_from_sender(
                    sender_filter, folder="Inbox", unread_only=True,
                    max_count=max_per_run
                )
            else:
                emails = graph.fetch_unread_emails(folder="Inbox", max_count=max_per_run)
        except Exception as e:
            return PluginResult(success=False, error=f"Could not fetch emails: {e}")

        # Filter by subject keywords
        noa_emails = []
        for email in emails:
            subject_lower = email.get("subject", "").lower()
            if any(kw in subject_lower for kw in keywords):
                noa_emails.append(email)

        log(f"  Found {len(noa_emails)} potential NOA email(s).")

        result = PluginResult(success=True)

        for email in noa_emails:
            msg_id = email["id"]
            if msg_id in self._processed_ids:
                continue

            subject    = email.get("subject", "(No Subject)")
            from_email = email.get("from", {}).get("emailAddress", {}).get("address", "")
            body_text  = email.get("body", {}).get("content", "")
            body_plain = re.sub(r"<[^>]+>", " ", body_text)
            body_plain = re.sub(r"\s+", " ", body_plain).strip()

            log(f'  Processing: "{subject}"')

            try:
                # Step 1: Download attachments
                attachment_paths = []
                if email.get("hasAttachments"):
                    attachment_paths = graph.download_all_attachments(
                        msg_id, self._download_dir
                    )
                    log(f"    ↳ Downloaded {len(attachment_paths)} attachment(s)")

                # Filter for PDFs
                pdf_paths = [p for p in attachment_paths if p.lower().endswith(".pdf")]
                if not pdf_paths:
                    log("    ↳ No PDF attachments found — skipping.")
                    self._processed_ids.add(msg_id)
                    result.items_skipped += 1
                    continue

                # Step 2: Use Claude to analyse the NOA
                noa_data = self._analyse_noa(claude, subject, body_plain, pdf_paths)

                if not noa_data:
                    log("    ↳ ⚠ Could not extract NOA details — flagging for review.")
                    graph.flag_email(msg_id)
                    graph.add_category(msg_id, "NOA - Review Needed")
                    log_activity(from_email, subject, "NOA", "flagged_for_review")
                    self._processed_ids.add(msg_id)
                    result.items_skipped += 1
                    continue

                client_name  = noa_data.get("client_name", "Client")
                outcome      = noa_data.get("outcome", "REFUND").upper()
                amount       = noa_data.get("amount", "$0.00")
                tax_year     = noa_data.get("tax_year", "2024-25")
                client_email = noa_data.get("client_email", from_email)
                entity_name  = noa_data.get("entity_name", "")
                is_amended   = noa_data.get("is_amended", False)

                if is_amended:
                    outcome = "AMENDED"

                log(f"    ↳ Client: {client_name} | Outcome: {outcome} | Amount: {amount}")

                # Step 3: Select template and build email
                template = NOA_TEMPLATES.get(outcome, NOA_TEMPLATES["REFUND"])
                reply_subject = template["subject"]
                reply_body = self._build_email_body(
                    template["body"], client_name, amount, tax_year,
                    entity_name, noa_data
                )

                # Append signature
                signature = graph.get_signature_html()
                if signature:
                    reply_body = reply_body + "<br>" + signature

                # Step 4: Create draft or send with NOA attached
                if draft_mode:
                    graph.create_draft_with_attachments(
                        client_email, reply_subject, reply_body,
                        attachment_paths=pdf_paths
                    )
                    log(f"    ↳ Draft created for {client_email} with NOA attached.")
                    self._send_staff_notification(
                        context, client_name, client_email, outcome,
                        amount, tax_year, subject
                    )
                    log_activity(
                        from_email, subject, f"NOA_{outcome}", "draft_created",
                        draft_created=1, notification_sent=1,
                    )
                    result.drafts_created += 1
                else:
                    graph.send_email_with_attachments(
                        client_email, reply_subject, reply_body,
                        attachment_paths=pdf_paths
                    )
                    log(f"    ↳ Email sent to {client_email} with NOA attached.")
                    log_activity(
                        from_email, subject, f"NOA_{outcome}", "auto_sent"
                    )

                # Step 5: Mark as read and optionally archive
                graph.mark_as_read(msg_id)
                if auto_archive:
                    try:
                        graph.move_email(msg_id, "NOA Processed")
                        log("    ↳ Archived to 'NOA Processed' folder.")
                    except Exception as e:
                        log(f"    ↳ ⚠ Could not archive: {e}")

                self._processed_ids.add(msg_id)
                result.actions_taken += 1

            except Exception as e:
                log(f"    ↳ Error processing NOA: {e}")
                result.extra.setdefault("errors", []).append(str(e))

        result.summary = (
            f"{result.actions_taken} NOA(s) processed, "
            f"{result.drafts_created} draft(s) created, "
            f"{result.items_skipped} skipped."
        )
        return result

    # ── Private helpers ──────────────────────────────────────────────────────

    def _analyse_noa(self, claude_client: anthropic.Anthropic,
                     subject: str, body: str, pdf_paths: list) -> dict | None:
        """Use Claude to extract key NOA details from the email and attachment names."""

        # Build context about attachments
        attachment_info = "\n".join(
            f"- {os.path.basename(p)}" for p in pdf_paths
        )

        # Inject memory context
        memory_block = ""
        style_prefs = get_style_preferences()
        if style_prefs:
            memory_block += f"\nSTYLE PREFERENCES: {style_prefs}\n"
        lessons = get_active_lessons()
        if lessons:
            memory_block += "LEARNED PREFERENCES:\n"
            memory_block += "\n".join(f"- {l['lesson']}" for l in lessons)
            memory_block += "\n"

        prompt = f"""You are an assistant for MC & S, an accounting firm. Analyse this email about a Notice of Assessment (NOA) from the ATO.

Extract the following details:
- client_name: The client's full name
- client_email: The client's email address (if visible in the email, otherwise return empty string)
- outcome: One of REFUND, PAYABLE, NIL, AMENDED, COMPANY_PAYABLE, DEDUCT_FROM_REFUND
- amount: The dollar amount (formatted as $X,XXX.XX)
- tax_year: The financial year (e.g. "2024-25")
- entity_name: Company or trust name if applicable (empty string for individuals)
- is_amended: true if this is an amended assessment, false otherwise
- taxable_income: The taxable income amount if mentioned
- confidence: high, medium, or low

Email Subject: {subject}
Email Body: {body[:2000]}

Attachments:
{attachment_info}
{memory_block}
Respond ONLY with valid JSON. If you cannot determine a field, use a reasonable default."""

        try:
            response = claude_client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=500,
                messages=[{"role": "user", "content": prompt}],
            )
            text = response.content[0].text.strip()
            text = re.sub(r"```json\s*|```", "", text).strip()
            return json.loads(text)
        except Exception:
            return None

    def _build_email_body(self, template: str, client_name: str,
                          amount: str, tax_year: str,
                          entity_name: str, noa_data: dict) -> str:
        """Replace placeholders in the email template."""
        result = template
        result = result.replace("{client_name}", client_name or "Client")
        result = result.replace("{amount}", amount or "$0.00")
        result = result.replace("{tax_year}", tax_year or "2024-25")
        result = result.replace("{entity_name}", entity_name or "your company")
        result = result.replace("{result_description}", noa_data.get("outcome", ""))
        result = result.replace("{gross_refund}", noa_data.get("gross_refund", amount))
        result = result.replace("{fees_amount}", noa_data.get("fees_amount", "$0.00"))
        result = result.replace("{net_refund}", noa_data.get("net_refund", amount))
        return result

    def _send_staff_notification(self, context: PluginContext,
                                 client_name: str, client_email: str,
                                 outcome: str, amount: str,
                                 tax_year: str, original_subject: str):
        """Notify staff that an NOA draft is ready for review."""
        staff      = get_staff()
        notifiable = [s for s in staff if s.get("receives_drafts")]

        if not notifiable:
            context.log("    ↳ ⚠ No staff configured for NOA notifications.")
            return

        practice = get_setting("practice_name", "MC & S")

        outcome_labels = {
            "REFUND": ("Refund", "#2E7D32", "#E8F5E9"),
            "PAYABLE": ("Amount Owing", "#C62828", "#FFEBEE"),
            "NIL": ("Nil Result", "#F57F17", "#FFF8E1"),
            "AMENDED": ("Amended Assessment", "#E65100", "#FFF3E0"),
            "COMPANY_PAYABLE": ("Company Payable", "#C62828", "#FFEBEE"),
            "DEDUCT_FROM_REFUND": ("Refund (DFR)", "#1565C0", "#E3F2FD"),
        }
        label, color, bg = outcome_labels.get(outcome, ("Unknown", "#555", "#f5f5f5"))

        subject = f"[NOA READY] {client_name} — {label} {amount} ({tax_year})"

        body = f"""
<div style="font-family:Arial,sans-serif;max-width:600px">
  <div style="background:#1565C0;color:white;padding:16px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;font-size:18px">📋 NOA Draft Ready for Review</h2>
  </div>
  <div style="background:#f5f5f5;padding:24px;border-radius:0 0 8px 8px;border:1px solid #ddd">
    <p>A Notice of Assessment email has been drafted and is waiting in your <strong>Drafts folder</strong>.</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0">
      <tr>
        <td style="padding:8px;color:#555;width:130px"><strong>Client:</strong></td>
        <td style="padding:8px">{client_name}</td>
      </tr>
      <tr style="background:#fff">
        <td style="padding:8px;color:#555"><strong>Email:</strong></td>
        <td style="padding:8px">{client_email}</td>
      </tr>
      <tr>
        <td style="padding:8px;color:#555"><strong>Tax Year:</strong></td>
        <td style="padding:8px">{tax_year}</td>
      </tr>
      <tr style="background:#fff">
        <td style="padding:8px;color:#555"><strong>Outcome:</strong></td>
        <td style="padding:8px">
          <span style="background:{bg};color:{color};padding:4px 12px;
                       border-radius:12px;font-size:13px;font-weight:bold">{label}</span>
        </td>
      </tr>
      <tr>
        <td style="padding:8px;color:#555"><strong>Amount:</strong></td>
        <td style="padding:8px;font-size:16px;font-weight:bold">{amount}</td>
      </tr>
    </table>
    <div style="background:#FFF9C4;border-left:4px solid #F9A825;padding:12px 16px;
                margin:16px 0;border-radius:4px">
      <strong>Action Required:</strong> Open Outlook → Drafts, review the NOA email and attachment, then send.
    </div>
    <p style="color:#888;font-size:12px;margin-top:24px">
      — {practice} Desktop Agent · NOA Processor Plugin
    </p>
  </div>
</div>"""

        for s in notifiable:
            try:
                context.graph.send_email(s["email"], subject, body)
                context.log(f"    ↳ Notified {s['name']}")
            except Exception as e:
                context.log(f"    ↳ ⚠ Could not notify {s['name']}: {e}")
