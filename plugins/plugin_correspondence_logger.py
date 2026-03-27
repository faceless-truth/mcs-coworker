"""
MC & S Plugin: Correspondence Logger
======================================
Plugin ID  : plugin_correspondence_logger
Version    : 1.0.0

WHAT IT DOES
------------
Replaces the manual Excel spreadsheets on the Z Drive for tracking
incoming and outgoing correspondence. The plugin:

1. Automatically logs all emails processed by the agent (sent, drafted, received)
2. Provides a searchable correspondence register stored in a local SQLite database
3. Allows manual logging of physical mail (incoming and outgoing) via the UI
4. Tracks status: pending, actioned, awaiting reply, complete
5. Exports the register to CSV/Excel on demand
6. Runs a daily summary showing outstanding items needing follow-up

REPLACES
--------
- Z Drive > MC&S > Admin > Correspondence spreadsheet (outgoing)
- Z Drive > MC&S > Admin > Clients Documents In spreadsheet (incoming)

SCHEDULE
--------
Default: every 10 minutes (checks for new sent/received items to log).
Also runs a daily summary at 8:00 AM.
"""

import json
import re
import csv
import os
from datetime import datetime, timedelta
from pathlib import Path

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity, get_db


# ── Database setup for correspondence ────────────────────────────────────────

def _ensure_correspondence_table():
    """Create the correspondence table if it doesn't exist."""
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS correspondence_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp       TEXT DEFAULT (datetime('now','localtime')),
            direction       TEXT NOT NULL,
            type            TEXT DEFAULT 'email',
            client_name     TEXT,
            client_email    TEXT,
            subject         TEXT,
            description     TEXT,
            status          TEXT DEFAULT 'logged',
            tracking_number TEXT,
            actioned_by     TEXT,
            actioned_date   TEXT,
            notes           TEXT,
            source_plugin   TEXT,
            message_id      TEXT
        );

        CREATE TABLE IF NOT EXISTS correspondence_last_check (
            id              INTEGER PRIMARY KEY CHECK (id = 1),
            last_sent_check TEXT,
            last_recv_check TEXT
        );
    """)
    # Ensure the single-row tracking record exists
    existing = conn.execute("SELECT id FROM correspondence_last_check WHERE id=1").fetchone()
    if not existing:
        conn.execute(
            "INSERT INTO correspondence_last_check (id, last_sent_check, last_recv_check) VALUES (1, ?, ?)",
            (datetime.now().isoformat(), datetime.now().isoformat()),
        )
    conn.commit()
    conn.close()


def log_correspondence(direction: str, client_name: str = "",
                       client_email: str = "", subject: str = "",
                       description: str = "", type_: str = "email",
                       status: str = "logged", tracking_number: str = "",
                       source_plugin: str = "", message_id: str = "",
                       notes: str = ""):
    """Add an entry to the correspondence log."""
    conn = get_db()
    conn.execute(
        """INSERT INTO correspondence_log
           (direction, type, client_name, client_email, subject,
            description, status, tracking_number, source_plugin, message_id, notes)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (direction, type_, client_name, client_email, subject,
         description, status, tracking_number, source_plugin, message_id, notes),
    )
    conn.commit()
    conn.close()


def get_correspondence(limit: int = 200, direction: str = None,
                       status: str = None, search: str = None) -> list[dict]:
    """Query the correspondence log with optional filters."""
    conn = get_db()
    query = "SELECT * FROM correspondence_log WHERE 1=1"
    params = []

    if direction:
        query += " AND direction=?"
        params.append(direction)
    if status:
        query += " AND status=?"
        params.append(status)
    if search:
        query += " AND (client_name LIKE ? OR subject LIKE ? OR description LIKE ?)"
        params.extend([f"%{search}%"] * 3)

    query += " ORDER BY id DESC LIMIT ?"
    params.append(limit)

    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_correspondence_status(entry_id: int, status: str,
                                 actioned_by: str = "", notes: str = ""):
    """Update the status of a correspondence entry."""
    conn = get_db()
    updates = ["status=?"]
    params = [status]

    if actioned_by:
        updates.append("actioned_by=?")
        params.append(actioned_by)
    if status in ("actioned", "complete"):
        updates.append("actioned_date=?")
        params.append(datetime.now().strftime("%Y-%m-%d %H:%M"))
    if notes:
        updates.append("notes=?")
        params.append(notes)

    params.append(entry_id)
    conn.execute(
        f"UPDATE correspondence_log SET {', '.join(updates)} WHERE id=?",
        params,
    )
    conn.commit()
    conn.close()


def get_outstanding_correspondence() -> list[dict]:
    """Return all correspondence items that need follow-up."""
    conn = get_db()
    rows = conn.execute(
        """SELECT * FROM correspondence_log
           WHERE status IN ('logged', 'awaiting_reply', 'pending')
           ORDER BY timestamp ASC"""
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def export_correspondence_csv(filepath: str, direction: str = None):
    """Export the correspondence log to a CSV file."""
    entries = get_correspondence(limit=10000, direction=direction)
    if not entries:
        return 0

    fieldnames = [
        "id", "timestamp", "direction", "type", "client_name",
        "client_email", "subject", "description", "status",
        "tracking_number", "actioned_by", "actioned_date", "notes",
    ]
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(entries)

    return len(entries)


class CorrespondenceLoggerPlugin(AgentPlugin):

    name        = "Correspondence Logger"
    description = "Automatically logs all email correspondence and replaces the manual spreadsheets."
    detail      = (
        "Replaces the manual Excel spreadsheets on the Z Drive for tracking incoming and "
        "outgoing correspondence. Automatically logs all emails processed by the agent, "
        "allows manual logging of physical mail, tracks status (pending, actioned, complete), "
        "and provides a searchable register with CSV export."
    )
    version = "1.0.0"
    icon    = "📝"
    author  = "MC & S"

    requires_graph  = True
    requires_claude = False

    default_schedule = Schedule.every_minutes(10)

    def load(self, context: PluginContext) -> bool:
        _ensure_correspondence_table()

        if not context.graph:
            context.log("📝 Correspondence Logger: Microsoft 365 not connected.")
            return False
        return True

    @classmethod
    def email_templates_schema(cls) -> list[dict]:
        return [
            {
                "key": "summary_prompt",
                "label": "Daily Summary Prompt",
                "default": (
                    "Summarise the outstanding correspondence items for today. "
                    "Group by incoming (needs action) and outgoing (awaiting reply)."
                ),
                "type": "prompt",
            },
            {
                "key": "summary_subject",
                "label": "Daily Summary Subject Line",
                "default": "[DAILY SUMMARY] Outstanding Correspondence Items",
                "type": "text",
            },
        ]

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {
                "key": "log_sent_emails",
                "label": "Log Sent Emails",
                "default": "1",
                "type": "bool",
                "help": "Automatically log emails sent from your account.",
            },
            {
                "key": "log_received_emails",
                "label": "Log Received Emails",
                "default": "1",
                "type": "bool",
                "help": "Automatically log emails received in your inbox.",
            },
            {
                "key": "ignore_senders",
                "label": "Ignore Senders (comma-separated)",
                "default": "noreply@,no-reply@,notifications@github.com,mailer-daemon@",
                "type": "textarea",
                "help": (
                    "Skip emails from these senders. Use partial matches. "
                    "E.g. 'noreply@' will skip all noreply addresses."
                ),
            },
            {
                "key": "daily_summary_hour",
                "label": "Daily Summary Hour (0-23)",
                "default": "8",
                "type": "number",
                "help": "Hour of the day to generate the outstanding items summary.",
            },
            {
                "key": "export_dir",
                "label": "CSV Export Directory",
                "default": "",
                "type": "text",
                "help": "Directory for CSV exports. Leave blank for default (Documents).",
            },
        ]

    def run(self, context: PluginContext) -> PluginResult:
        graph = context.graph
        log   = context.log

        log_sent     = self.get_plugin_setting("log_sent_emails", "1") == "1"
        log_received = self.get_plugin_setting("log_received_emails", "1") == "1"
        ignore_list  = [
            s.strip().lower()
            for s in self.get_plugin_setting(
                "ignore_senders",
                "noreply@,no-reply@,notifications@github.com,mailer-daemon@"
            ).split(",")
            if s.strip()
        ]

        log("📝 Correspondence Logger: Scanning for new correspondence...")

        result = PluginResult(success=True)
        logged_count = 0

        # Get last check timestamps
        conn = get_db()
        check_row = conn.execute(
            "SELECT * FROM correspondence_last_check WHERE id=1"
        ).fetchone()
        conn.close()

        last_sent_check = check_row["last_sent_check"] if check_row else None
        last_recv_check = check_row["last_recv_check"] if check_row else None

        # Log sent emails
        if log_sent:
            try:
                sent_since = None
                if last_sent_check:
                    sent_since = last_sent_check.replace(" ", "T")
                    if not sent_since.endswith("Z"):
                        sent_since += "Z"

                sent_emails = graph.fetch_recent_emails(
                    folder="SentItems", max_count=50,
                    since_datetime=sent_since
                )

                for email in sent_emails:
                    sender = email.get("from", {}).get("emailAddress", {}).get("address", "")
                    to_list = email.get("toRecipients", [])
                    to_email = to_list[0]["emailAddress"]["address"] if to_list else ""
                    subject = email.get("subject", "")
                    received = email.get("receivedDateTime", "")
                    msg_id = email.get("id", "")

                    # Check if already logged
                    conn = get_db()
                    existing = conn.execute(
                        "SELECT id FROM correspondence_log WHERE message_id=?",
                        (msg_id,)
                    ).fetchone()
                    conn.close()

                    if existing:
                        continue

                    # Extract client name from subject or recipient
                    client_name = self._extract_name_from_email(to_email)

                    log_correspondence(
                        direction="outgoing",
                        client_name=client_name,
                        client_email=to_email,
                        subject=subject,
                        description=f"Email sent to {to_email}",
                        type_="email",
                        status="complete",
                        source_plugin="correspondence_logger",
                        message_id=msg_id,
                    )
                    logged_count += 1

            except Exception as e:
                log(f"  ⚠ Error logging sent emails: {e}")

        # Log received emails
        if log_received:
            try:
                recv_since = None
                if last_recv_check:
                    recv_since = last_recv_check.replace(" ", "T")
                    if not recv_since.endswith("Z"):
                        recv_since += "Z"

                received_emails = graph.fetch_recent_emails(
                    folder="Inbox", max_count=50,
                    since_datetime=recv_since
                )

                for email in received_emails:
                    from_addr = email.get("from", {}).get("emailAddress", {}).get("address", "")
                    from_name = email.get("from", {}).get("emailAddress", {}).get("name", "")
                    subject = email.get("subject", "")
                    msg_id = email.get("id", "")

                    # Skip ignored senders
                    if any(ign in from_addr.lower() for ign in ignore_list):
                        continue

                    # Check if already logged
                    conn = get_db()
                    existing = conn.execute(
                        "SELECT id FROM correspondence_log WHERE message_id=?",
                        (msg_id,)
                    ).fetchone()
                    conn.close()

                    if existing:
                        continue

                    log_correspondence(
                        direction="incoming",
                        client_name=from_name or self._extract_name_from_email(from_addr),
                        client_email=from_addr,
                        subject=subject,
                        description=f"Email received from {from_name or from_addr}",
                        type_="email",
                        status="logged",
                        source_plugin="correspondence_logger",
                        message_id=msg_id,
                    )
                    logged_count += 1

            except Exception as e:
                log(f"  ⚠ Error logging received emails: {e}")

        # Update last check timestamps
        now = datetime.now().isoformat()
        conn = get_db()
        conn.execute(
            "UPDATE correspondence_last_check SET last_sent_check=?, last_recv_check=? WHERE id=1",
            (now, now),
        )
        conn.commit()
        conn.close()

        log(f"  Logged {logged_count} new correspondence item(s).")

        # Check if it's time for the daily summary
        summary_hour = int(self.get_plugin_setting("daily_summary_hour", "8"))
        current_hour = datetime.now().hour
        last_summary = self.get_plugin_setting("last_summary_date", "")
        today_str = datetime.now().strftime("%Y-%m-%d")

        if current_hour == summary_hour and last_summary != today_str:
            self._send_daily_summary(context)
            self.set_plugin_setting("last_summary_date", today_str)

        result.actions_taken = logged_count
        result.summary = f"{logged_count} correspondence item(s) logged."
        return result

    # ── Private helpers ──────────────────────────────────────────────────────

    def _extract_name_from_email(self, email_addr: str) -> str:
        """Try to extract a human name from an email address."""
        if not email_addr:
            return ""
        local = email_addr.split("@")[0]
        # Replace dots, underscores, hyphens with spaces
        name = re.sub(r"[._\-]", " ", local)
        return name.title()

    def _send_daily_summary(self, context: PluginContext):
        """Log a daily summary of outstanding correspondence items."""
        outstanding = get_outstanding_correspondence()
        if not outstanding:
            context.log("  📝 Daily summary: No outstanding items.")
            return

        incoming = [o for o in outstanding if o["direction"] == "incoming"]
        outgoing = [o for o in outstanding if o["direction"] == "outgoing"]
        context.log(
            f"  📝 Daily summary: {len(outstanding)} outstanding items "
            f"({len(incoming)} incoming, {len(outgoing)} outgoing awaiting reply)."
        )
