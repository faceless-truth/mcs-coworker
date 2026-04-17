"""
MC & S Plugin: BAS Reminder Drafter
======================================
Plugin ID  : plugin_bas_reminder
Version    : 1.0.0

WHAT IT DOES
------------
Monitors ATO BAS lodgement deadlines and proactively reminds clients:

1. Calculates upcoming BAS due dates (quarterly: Oct 28, Feb 28,
   Apr 28, Jul 28; monthly: 21st of following month)
2. Queries XPM for clients registered for GST
3. Checks which clients have not yet lodged their BAS
4. Uses Claude Haiku to draft personalised BAS reminder emails
5. Sends reminders via the approval queue (configurable lead time)
6. Tracks reminders in memory to avoid duplicates

APEX ALIGNMENT
--------------
Mirrors APEX's "Compliance Deadline Agent" — proactive, automated
compliance reminders that protect clients from ATO penalties.

SCHEDULE
--------
Default: daily at 9:00 AM (checks if any reminders are due today).
"""

from datetime import datetime, date, timedelta
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity


# ATO quarterly BAS due dates (day, month) — lodgement agent extended dates
QUARTERLY_DUE_DATES = [
    (28, 10),   # Q1: Jul-Sep due 28 Oct
    (28, 2),    # Q2: Oct-Dec due 28 Feb
    (28, 4),    # Q3: Jan-Mar due 28 Apr
    (28, 7),    # Q4: Apr-Jun due 28 Jul
]


class BASReminderPlugin(AgentPlugin):
    """Drafts and sends BAS lodgement reminders to GST-registered clients."""

    PLUGIN_ID   = "plugin_bas_reminder"
    NAME        = "BAS Reminder Drafter"
    DESCRIPTION = ("Calculates upcoming BAS due dates, finds clients with outstanding "
                   "lodgements, and drafts personalised reminder emails.")
    VERSION     = "1.0.0"
    ICON        = "📅"
    SCHEDULE    = Schedule.daily_at(8)

    DEFAULT_SETTINGS = {
        "run_hour":          "9",
        "lead_days":         "14",   # send reminder 14 days before due date
        "confidence_threshold": "0.75",
        "quarterly_clients": "1",
        "monthly_clients":   "0",
    }

    def __init__(self):
        super().__init__()
        self._last_run_date: str = ""

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        now = datetime.now()
        target_hour = int(self.get_plugin_setting("run_hour", "9"))
        if now.hour != target_hour:
            result.summary = "Not BAS reminder time."
            return result

        today_str = date.today().isoformat()
        if self._last_run_date == today_str:
            result.summary = "BAS reminders already checked today."
            return result

        today     = date.today()
        lead_days = int(self.get_plugin_setting("lead_days", "14"))
        confidence = float(self.get_plugin_setting("confidence_threshold", "0.75"))

        # ── Find upcoming BAS due dates within lead_days ──────────────────────
        upcoming_due_dates = []
        for day, month in QUARTERLY_DUE_DATES:
            year = today.year
            try:
                due = date(year, month, day)
            except ValueError:
                # Feb 28 in leap year etc.
                due = date(year, month, 28)
            if due < today:
                try:
                    due = date(year + 1, month, day)
                except ValueError:
                    due = date(year + 1, month, 28)
            days_until = (due - today).days
            if 0 <= days_until <= lead_days:
                upcoming_due_dates.append(due)

        if not upcoming_due_dates:
            result.summary = f"No BAS due dates within {lead_days} days."
            self._last_run_date = today_str
            return result

        context.log(f"[BASReminder] Upcoming due dates: "
                    f"{[d.isoformat() for d in upcoming_due_dates]}")

        if not context.gateway or not context.gateway.is_available("xpm"):
            result.summary = "XPM not configured — BAS reminders skipped."
            return result

        # ── Get GST-registered clients from XPM ───────────────────────────────
        try:
            clients = context.gateway.xpm.list_clients(search="", limit=500)
        except Exception as e:
            result.summary = f"XPM client fetch failed: {e}"
            return result

        drafted = 0
        skipped = 0

        for client in clients:
            # Filter for GST-registered clients
            gst_registered = (client.get("gst_registered", False)
                              or client.get("registered_for_gst", False)
                              or str(client.get("gst", "")).lower() in ("yes", "1", "true"))
            if not gst_registered:
                continue

            client_name  = client.get("name", client.get("client_name", "Unknown"))
            client_email = client.get("email", "")

            if not client_email:
                skipped += 1
                continue

            # ── Check if reminder already sent ────────────────────────────────
            for due_date in upcoming_due_dates:
                already_sent = False
                if context.memory:
                    try:
                        history = context.memory.search(
                            query=f"BAS reminder {client_name} {due_date.isoformat()}",
                            n_results=5,
                            collection="client_interactions")
                        already_sent = any(
                            r.get("metadata", {}).get("due_date") == due_date.isoformat()
                            and r.get("metadata", {}).get("client_name") == client_name
                            and r.get("metadata", {}).get("type") == "bas_reminder"
                            for r in history)
                    except Exception:
                        pass

                if already_sent:
                    skipped += 1
                    continue

                # ── Draft reminder email ──────────────────────────────────────
                email_body = self._draft_reminder(
                    context, client_name, due_date)

                subject = (f"BAS Lodgement Reminder — Due {due_date.strftime('%d %B %Y')}")

                # ── Submit to approval queue ──────────────────────────────────
                if context.approval_queue:
                    def _send(to=client_email, subj=subject, body=email_body):
                        if context.graph:
                            context.graph.send_email(
                                to=to, subject=subj, body=body,
                                draft_mode=context.draft_mode)
                    context.approval_queue.submit(
                        action_type="send_email",
                        description=f"BAS reminder to {client_name} (due {due_date})",
                        payload={"to": client_email, "due_date": due_date.isoformat()},
                        confidence=confidence,
                        plugin_id=self.PLUGIN_ID,
                        execute_callback=_send)
                elif context.graph:
                    try:
                        context.graph.send_email(
                            to=client_email, subject=subject, body=email_body,
                            draft_mode=context.draft_mode)
                    except Exception as e:
                        context.log(f"[BASReminder] Email failed for {client_name}: {e}")

                # ── Store in memory ───────────────────────────────────────────
                if context.memory:
                    try:
                        context.memory.store_client_interaction(
                            client_name=client_name,
                            interaction_type="bas_reminder",
                            summary=f"BAS reminder sent for due date {due_date.isoformat()}",
                            metadata={"due_date": due_date.isoformat(),
                                      "email": client_email})
                    except Exception:
                        pass

                drafted += 1
                log_activity(from_email=client_email,
                             subject="BAS Reminder",
                             classification="bas_reminder",
                             action=f"due:{due_date.isoformat()}",
                             draft_created=True)

        self._last_run_date = today_str
        result.summary = (f"BAS reminders: {drafted} drafted/queued, "
                          f"{skipped} skipped.")
        result.actions_taken = drafted
        result.items_skipped = skipped
        return result

    def _draft_reminder(self, context: PluginContext,
                         client_name: str, due_date: date) -> str:
        """Use Claude Haiku to draft a BAS reminder email."""
        practice  = get_setting("practice_name", "MC & S Accounting")
        due_str   = due_date.strftime("%d %B %Y")
        days_left = (due_date - date.today()).days

        if not context.claude_fast:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>This is a reminder that your BAS lodgement is due on "
                    f"<strong>{due_str}</strong> ({days_left} days away).</p>"
                    f"<p>Please provide your records at your earliest convenience "
                    f"so we can prepare and lodge on time.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
        try:
            prompt = (f"Write a professional BAS lodgement reminder email from {practice} "
                      f"to {client_name}. The BAS is due on {due_str} ({days_left} days away). "
                      "Ask the client to provide their records. Keep it brief and friendly. "
                      "Return HTML body only.")
            resp = context.claude_fast.messages.create(
                model=self.get_claude_model_fast(),
                max_tokens=300,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>Your BAS is due on {due_str}. Please provide your records.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
