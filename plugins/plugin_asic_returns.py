"""
MC & S Plugin: ASIC Annual Return Handler
==========================================
Plugin ID  : plugin_asic_returns
Version    : 1.0.0

WHAT IT DOES
------------
Monitors the inbox for ASIC Annual Return emails from Nowinfinity.
For each email it:

1. Detects the Nowinfinity ASIC annual return notification
2. Downloads the 3 PDF attachments:
   - Cover letter
   - Solvency statement (needs client signature)
   - Company statement / ASIC invoice
3. Uses Claude AI to extract company name, ACN, due date, and ASIC fee
4. Saves the PDFs to the client's folder (if configured)
5. Drafts an email to the client with:
   - All 3 PDFs attached
   - The MC & S ASIC fee invoice amount
   - Instructions to sign the solvency statement and pay the ASIC fee
6. Logs the ASIC return in a tracking table for follow-up
7. Supports a weekly "Burning" check for overdue returns

REPLACES
--------
- Manual download and filing of Nowinfinity emails
- Manual creation of Xero invoices (future: Xero API integration)
- Manual email drafting using Outlook templates
- Manual tracking via ASIC Return Details spreadsheet

SCHEDULE
--------
Default: every 5 minutes during business hours.
"""

import json
import re
import os
from datetime import datetime, timedelta
from pathlib import Path

import anthropic

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity, get_db, get_style_preferences, get_active_lessons


# ── Database setup for ASIC tracking ────────────────────────────────────────

def _ensure_asic_tables():
    """Create the ASIC tracking tables if they don't exist."""
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS asic_returns (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp       TEXT DEFAULT (datetime('now','localtime')),
            company_name    TEXT,
            acn             TEXT,
            client_name     TEXT,
            client_email    TEXT,
            asic_fee        TEXT,
            mcs_fee         TEXT,
            due_date        TEXT,
            status          TEXT DEFAULT 'pending',
            solvency_signed INTEGER DEFAULT 0,
            asic_paid       INTEGER DEFAULT 0,
            mcs_invoiced    INTEGER DEFAULT 0,
            email_sent      INTEGER DEFAULT 0,
            reminder_count  INTEGER DEFAULT 0,
            last_reminder   TEXT,
            source_email_id TEXT,
            notes           TEXT
        );

        CREATE TABLE IF NOT EXISTS asic_settings (
            id              INTEGER PRIMARY KEY CHECK (id = 1),
            default_mcs_fee TEXT DEFAULT '66.00',
            reminder_days   INTEGER DEFAULT 14,
            save_path       TEXT DEFAULT ''
        );
    """)
    # Ensure settings row exists
    existing = conn.execute("SELECT id FROM asic_settings WHERE id=1").fetchone()
    if not existing:
        conn.execute(
            "INSERT INTO asic_settings (id, default_mcs_fee, reminder_days) VALUES (1, '66.00', 14)"
        )
    conn.commit()
    conn.close()


def get_asic_returns(status: str = None, limit: int = 100) -> list[dict]:
    """Query ASIC returns with optional status filter."""
    conn = get_db()
    query = "SELECT * FROM asic_returns WHERE 1=1"
    params = []
    if status:
        query += " AND status=?"
        params.append(status)
    query += " ORDER BY id DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_overdue_asic_returns(days: int = 14) -> list[dict]:
    """Get ASIC returns that are overdue for follow-up."""
    conn = get_db()
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    rows = conn.execute(
        """SELECT * FROM asic_returns
           WHERE status IN ('pending', 'awaiting_solvency', 'awaiting_payment')
           AND timestamp < ?
           ORDER BY timestamp ASC""",
        (cutoff,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_asic_return(return_id: int, **kwargs):
    """Update fields on an ASIC return record."""
    conn = get_db()
    updates = []
    params = []
    for key, value in kwargs.items():
        updates.append(f"{key}=?")
        params.append(value)
    params.append(return_id)
    conn.execute(
        f"UPDATE asic_returns SET {', '.join(updates)} WHERE id=?",
        params,
    )
    conn.commit()
    conn.close()


# ── Default ASIC email template ─────────────────────────────────────────────

ASIC_EMAIL_TEMPLATE = """<p>Dear {client_name},</p>

<p>We have received the <strong>ASIC Annual Return</strong> for <strong>{company_name}</strong> (ACN: {acn}).</p>

<p>Please find attached:</p>
<ol>
  <li><strong>Cover Letter</strong> — summary of the annual return</li>
  <li><strong>Solvency Resolution</strong> — this <u>must be signed</u> and returned to us</li>
  <li><strong>Company Statement</strong> — includes the ASIC annual review fee of <strong>{asic_fee}</strong></li>
</ol>

<p><strong>Action required:</strong></p>
<ul>
  <li>Sign the Solvency Resolution and return it to us (scan/photo is fine)</li>
  <li>Pay the ASIC fee of <strong>{asic_fee}</strong> using the payment details on the Company Statement</li>
  <li>Our administration fee of <strong>{mcs_fee}</strong> (inc. GST) will be invoiced separately</li>
</ul>

<p>The ASIC fee is due by <strong>{due_date}</strong>. Late payment may incur penalties from ASIC.</p>

<p>If you have any questions, please don't hesitate to get in touch.</p>"""

ASIC_REMINDER_TEMPLATE = """<p>Dear {client_name},</p>

<p>This is a friendly reminder regarding the <strong>ASIC Annual Return</strong> for <strong>{company_name}</strong> (ACN: {acn}).</p>

<p>Our records show the following items are still outstanding:</p>
<ul>
{outstanding_items}
</ul>

<p>The ASIC fee was due on <strong>{due_date}</strong>. To avoid late penalties, please action these items at your earliest convenience.</p>

<p>If you have already attended to this, please disregard this email and let us know so we can update our records.</p>

<p>If you need any assistance, please don't hesitate to contact us.</p>"""

ASIC_OVERDUE_TEMPLATE = """<p>Dear {client_name},</p>

<p><strong>URGENT — ASIC Annual Return Overdue</strong></p>

<p>We are writing to advise that the ASIC Annual Return for <strong>{company_name}</strong> (ACN: {acn}) is now <strong>overdue</strong>.</p>

<p>The following items remain outstanding:</p>
<ul>
{outstanding_items}
</ul>

<p>ASIC may impose late fees or commence deregistration proceedings if this is not attended to promptly.</p>

<p>Please action these items as a matter of urgency. If you have any questions or need assistance, please contact us immediately.</p>"""


class ASICReturnPlugin(AgentPlugin):

    name        = "ASIC Annual Return Handler"
    description = "Processes ASIC annual returns from Nowinfinity — downloads, tracks, and drafts client emails."
    detail      = (
        "Monitors your inbox for ASIC Annual Return emails from Nowinfinity. Downloads the "
        "3 PDF attachments (cover letter, solvency statement, company statement), extracts "
        "company details using AI, drafts a client email with all attachments, and tracks "
        "the return status. Supports automated reminders for overdue returns."
    )
    version = "1.0.0"
    icon    = "🏢"
    author  = "MC & S"

    requires_graph  = True
    requires_claude = True

    default_schedule = Schedule.every_minutes(5)

    _processed_ids: set
    _download_dir: str

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        self._download_dir = str(
            Path.home() / ".mcs_email_automation" / "asic_downloads"
        )
        os.makedirs(self._download_dir, exist_ok=True)
        _ensure_asic_tables()

        if not context.graph:
            context.log("🏢 ASIC Handler: Microsoft 365 not connected.")
            return False
        if not context.claude:
            context.log("🏢 ASIC Handler: Anthropic API key not configured.")
            return False
        return True

    @classmethod
    def email_templates_schema(cls) -> list[dict]:
        return [
            {
                "key": "draft_prompt",
                "label": "ASIC Analysis Prompt",
                "default": (
                    "You are an assistant for MC & S, an accounting firm. Analyse this email "
                    "about an ASIC Annual Return from Nowinfinity.\n\n"
                    "Extract: company_name, acn, client_name, client_email, asic_fee, "
                    "due_date, confidence.\n\n"
                    "Respond ONLY with valid JSON."
                ),
                "type": "prompt",
            },
            {
                "key": "email_closing",
                "label": "Sign-off Text",
                "default": "If you have any questions, please don't hesitate to get in touch.",
                "type": "textarea",
            },
        ]

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {
                "key": "nowinfinity_sender",
                "label": "Nowinfinity Sender Email",
                "default": "nowinfinity",
                "type": "text",
                "help": "Keyword to match Nowinfinity sender emails (partial match).",
            },
            {
                "key": "subject_keywords",
                "label": "Subject Keywords",
                "default": "annual return,annual statement,ASIC,solvency",
                "type": "text",
                "help": "Comma-separated keywords to identify ASIC emails.",
            },
            {
                "key": "default_mcs_fee",
                "label": "MC & S Admin Fee (inc. GST)",
                "default": "$66.00",
                "type": "text",
                "help": "Default MC & S administration fee for ASIC annual returns.",
            },
            {
                "key": "reminder_days",
                "label": "Days Before First Reminder",
                "default": "14",
                "type": "number",
                "help": "Number of days after initial email before sending a reminder.",
            },
            {
                "key": "auto_archive",
                "label": "Auto-Archive Processed Emails",
                "default": "1",
                "type": "bool",
                "help": "Move processed Nowinfinity emails to an 'ASIC Processed' folder.",
            },
            {
                "key": "check_burning",
                "label": "Check Burning (Overdue) Returns",
                "default": "1",
                "type": "bool",
                "help": "Automatically check for overdue ASIC returns and draft reminders.",
            },
            {
                "key": "max_per_run",
                "label": "Max Emails Per Run",
                "default": "10",
                "type": "number",
                "help": "Maximum ASIC emails to process per run.",
            },
        ]

    def run(self, context: PluginContext) -> PluginResult:
        graph      = context.graph
        claude     = context.claude
        log        = context.log
        draft_mode = context.draft_mode

        sender_filter    = self.get_plugin_setting("nowinfinity_sender", "nowinfinity")
        subject_keywords = self.get_plugin_setting(
            "subject_keywords", "annual return,annual statement,ASIC,solvency"
        )
        max_per_run  = int(self.get_plugin_setting("max_per_run", "10"))
        auto_archive = self.get_plugin_setting("auto_archive", "1") == "1"
        check_burning = self.get_plugin_setting("check_burning", "1") == "1"
        mcs_fee      = self.get_plugin_setting("default_mcs_fee", "$66.00")

        keywords = [k.strip().lower() for k in subject_keywords.split(",") if k.strip()]

        log("🏢 ASIC Handler: Scanning inbox for annual return emails...")

        result = PluginResult(success=True)

        # ── Process new ASIC emails ──────────────────────────────────────────
        try:
            emails = graph.fetch_emails_from_sender(
                sender_filter, folder="Inbox", unread_only=True,
                max_count=max_per_run
            )
        except Exception as e:
            return PluginResult(success=False, error=f"Could not fetch emails: {e}")

        # Filter by subject keywords
        asic_emails = []
        for email in emails:
            subject_lower = email.get("subject", "").lower()
            if any(kw in subject_lower for kw in keywords):
                asic_emails.append(email)

        log(f"  Found {len(asic_emails)} potential ASIC email(s).")

        for email in asic_emails:
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

                pdf_paths = [p for p in attachment_paths if p.lower().endswith(".pdf")]

                # Step 2: Use Claude to extract company details
                asic_data = self._analyse_asic_email(
                    claude, subject, body_plain, pdf_paths
                )

                if not asic_data:
                    log("    ↳ ⚠ Could not extract ASIC details — flagging for review.")
                    graph.flag_email(msg_id)
                    self._processed_ids.add(msg_id)
                    result.items_skipped += 1
                    continue

                company_name = asic_data.get("company_name", "Unknown Company")
                acn          = asic_data.get("acn", "")
                client_name  = asic_data.get("client_name", "Director")
                client_email = asic_data.get("client_email", "")
                asic_fee     = asic_data.get("asic_fee", "$290.00")
                due_date     = asic_data.get("due_date", "")

                log(f"    ↳ Company: {company_name} | ACN: {acn} | Fee: {asic_fee}")

                # Step 3: Record in tracking database
                conn = get_db()
                conn.execute(
                    """INSERT INTO asic_returns
                       (company_name, acn, client_name, client_email, asic_fee,
                        mcs_fee, due_date, status, source_email_id)
                       VALUES (?, ?, ?, ?, ?, ?, ?, 'pending', ?)""",
                    (company_name, acn, client_name, client_email, asic_fee,
                     mcs_fee, due_date, msg_id),
                )
                conn.commit()
                conn.close()

                # Step 4: Draft email to client (if we have their email)
                if client_email:
                    reply_subject = f"ASIC Annual Return — {company_name} — MC & S Accounting"
                    reply_body = ASIC_EMAIL_TEMPLATE.format(
                        client_name=client_name,
                        company_name=company_name,
                        acn=acn,
                        asic_fee=asic_fee,
                        mcs_fee=mcs_fee,
                        due_date=due_date or "as per the attached statement",
                    )

                    # Append signature
                    signature = graph.get_signature_html()
                    if signature:
                        reply_body = reply_body + "<br>" + signature

                    if draft_mode:
                        graph.create_draft_with_attachments(
                            client_email, reply_subject, reply_body,
                            attachment_paths=pdf_paths
                        )
                        log(f"    ↳ Draft created in Drafts folder.")
                        log_activity(
                            from_email, subject, "ASIC_ANNUAL_RETURN", "draft_created",
                            draft_created=1,
                        )
                        result.drafts_created += 1
                    else:
                        graph.send_email_with_attachments(
                            client_email, reply_subject, reply_body,
                            attachment_paths=pdf_paths
                        )
                        log(f"    ↳ Email sent to {client_email}.")
                        log_activity(
                            from_email, subject, "ASIC_ANNUAL_RETURN", "auto_sent"
                        )

                    # Update tracking
                    conn = get_db()
                    conn.execute(
                        "UPDATE asic_returns SET email_sent=1, status='awaiting_solvency' WHERE source_email_id=?",
                        (msg_id,)
                    )
                    conn.commit()
                    conn.close()
                else:
                    log("    ↳ ⚠ No client email found — flagged for manual handling.")
                    conn = get_db()
                    conn.execute(
                        "UPDATE asic_returns SET status='no_email', notes='Client email not found' WHERE source_email_id=?",
                        (msg_id,)
                    )
                    conn.commit()
                    conn.close()

                # Step 5: Mark as read and archive
                graph.mark_as_read(msg_id)
                if auto_archive:
                    try:
                        graph.move_email(msg_id, "ASIC Processed")
                        log("    ↳ Archived to 'ASIC Processed' folder.")
                    except Exception as e:
                        log(f"    ↳ ⚠ Could not archive: {e}")

                self._processed_ids.add(msg_id)
                result.actions_taken += 1

            except Exception as e:
                log(f"    ↳ Error processing ASIC email: {e}")
                result.extra.setdefault("errors", []).append(str(e))

        # ── Check for burning (overdue) returns ──────────────────────────────
        if check_burning:
            reminder_days = int(self.get_plugin_setting("reminder_days", "14"))
            overdue = get_overdue_asic_returns(days=reminder_days)

            if overdue:
                log(f"  🔥 Found {len(overdue)} overdue ASIC return(s) needing reminders.")
                for ret in overdue:
                    # Only send one reminder per run, max 3 total
                    if ret["reminder_count"] >= 3:
                        continue

                    self._send_reminder(context, ret, draft_mode)
                    result.actions_taken += 1

        result.summary = (
            f"{result.actions_taken} ASIC return(s) processed, "
            f"{result.drafts_created} draft(s) created, "
            f"{result.items_skipped} skipped."
        )
        return result

    # ── Private helpers ──────────────────────────────────────────────────────

    def _analyse_asic_email(self, claude_client: anthropic.Anthropic,
                            subject: str, body: str, pdf_paths: list) -> dict | None:
        """Use Claude to extract ASIC annual return details."""
        attachment_info = "\n".join(
            f"- {os.path.basename(p)}" for p in pdf_paths
        )

        prompt = f"""You are an assistant for MC & S, an accounting firm. Analyse this email about an ASIC Annual Return from Nowinfinity.

Extract the following details:
- company_name: The company's registered name
- acn: The Australian Company Number (9 digits, formatted as XXX XXX XXX)
- client_name: The director or contact person's name
- client_email: The client's email address (if visible, otherwise empty string)
- asic_fee: The ASIC annual review fee amount (formatted as $XXX.XX)
- due_date: The due date for the ASIC fee (formatted as DD/MM/YYYY)
- confidence: high, medium, or low

Email Subject: {subject}
Email Body: {body[:2000]}

Attachments:
{attachment_info}

Respond ONLY with valid JSON. If you cannot determine a field, use a reasonable default."""

        try:
            response = claude_client.messages.create(
                model=self.get_claude_model(),
                max_tokens=400,
                messages=[{"role": "user", "content": prompt}],
            )
            text = response.content[0].text.strip()
            text = re.sub(r"```json\s*|```", "", text).strip()
            return json.loads(text)
        except Exception:
            return None

    def _send_reminder(self, context: PluginContext, ret: dict, draft_mode: bool):
        """Send a reminder email for an overdue ASIC return."""
        client_email = ret.get("client_email", "")
        if not client_email:
            context.log(f"    ↳ ⚠ No email for {ret['company_name']} — skipping reminder.")
            return

        # Build outstanding items list
        items = []
        if not ret.get("solvency_signed"):
            items.append("<li>Signed Solvency Resolution — please sign and return to us</li>")
        if not ret.get("asic_paid"):
            items.append(f"<li>ASIC annual review fee of <strong>{ret.get('asic_fee', '$290')}</strong></li>")

        if not items:
            return

        outstanding_items = "\n".join(items)

        # Choose template based on reminder count
        reminder_count = ret.get("reminder_count", 0)
        if reminder_count >= 2:
            template = ASIC_OVERDUE_TEMPLATE
            subject_prefix = "URGENT — "
        else:
            template = ASIC_REMINDER_TEMPLATE
            subject_prefix = "Reminder — "

        reply_subject = f"{subject_prefix}ASIC Annual Return — {ret['company_name']}"
        reply_body = template.format(
            client_name=ret.get("client_name", "Director"),
            company_name=ret.get("company_name", ""),
            acn=ret.get("acn", ""),
            asic_fee=ret.get("asic_fee", "$290"),
            due_date=ret.get("due_date", "as per original notice"),
            outstanding_items=outstanding_items,
        )

        # Append signature
        signature = context.graph.get_signature_html()
        if signature:
            reply_body = reply_body + "<br>" + signature

        try:
            if draft_mode:
                context.graph.create_draft(
                    client_email, reply_subject, reply_body
                )
                context.log(f"    ↳ Reminder draft created for {ret['company_name']} ({client_email})")
            else:
                context.graph.send_email(
                    client_email, reply_subject, reply_body
                )
                context.log(f"    ↳ Reminder sent to {client_email}")

            # Update tracking
            update_asic_return(
                ret["id"],
                reminder_count=reminder_count + 1,
                last_reminder=datetime.now().strftime("%Y-%m-%d %H:%M"),
            )
            log_activity(
                client_email, reply_subject, "ASIC_REMINDER",
                "draft_created" if draft_mode else "auto_sent",
            )

        except Exception as e:
            context.log(f"    ↳ ⚠ Error sending reminder: {e}")

