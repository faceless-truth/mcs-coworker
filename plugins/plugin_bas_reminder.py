"""
MC & S Plugin: BAS Reminder Drafter
======================================
Plugin ID  : plugin_bas_reminder
Version    : 2.0.0

WHAT IT DOES
------------
Monitors ATO BAS lodgement deadlines and proactively reminds clients
using ADAPTIVE scheduling — each client is reminded on the correct
cadence based on their lodgement frequency (monthly or quarterly),
which is learned from XPM and stored in vector memory.

1. Runs on the 21st of every month (catches both monthly and quarterly
   windows in a single pass)
2. For each GST-registered client, determines their lodgement frequency:
     - Checks XPM client record for "bas_frequency" field
     - Falls back to vector memory for previously learned frequency
     - Falls back to "quarterly" as the safe default
3. Calculates the correct next due date for that client's frequency
4. Sends a reminder if the due date is within the configured lead window
5. Stores the learned frequency in memory so it improves over time
6. Uses Claude Haiku to draft personalised, frequency-aware emails

ADAPTIVE LEARNING
-----------------
Each time the plugin runs, it stores the detected frequency per client
in vector memory under the key "bas_frequency_learned". On subsequent
runs, if XPM does not return a frequency field, the plugin uses the
stored value — creating a self-improving knowledge base.

ATO DUE DATES
-------------
Quarterly lodgers (extended lodgement agent dates):
  Q1 Jul-Sep  → 28 Oct
  Q2 Oct-Dec  → 28 Feb
  Q3 Jan-Mar  → 28 Apr
  Q4 Apr-Jun  → 28 Jul

Monthly lodgers:
  21st of the following month (e.g., Jan BAS due 21 Feb)

SCHEDULE
--------
Default: 21st of every month at 08:00.
Covers both monthly (due 21st) and quarterly windows in a single pass.
"""

from datetime import datetime, date, timedelta
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule, PluginCategory
from config import get_setting, log_activity


# ATO quarterly BAS due dates (day, month) — lodgement agent extended dates
QUARTERLY_DUE_DATES = [
    (28, 10),   # Q1: Jul-Sep due 28 Oct
    (28, 2),    # Q2: Oct-Dec due 28 Feb
    (28, 4),    # Q3: Jan-Mar due 28 Apr
    (28, 7),    # Q4: Apr-Jun due 28 Jul
]

# XPM field names that may indicate BAS frequency
XPM_FREQUENCY_FIELDS = [
    "bas_frequency", "gst_frequency", "lodgement_frequency",
    "reporting_frequency", "bas_reporting_period",
]

# Values that indicate monthly lodgement
MONTHLY_INDICATORS = {"monthly", "month", "m", "1", "1m", "monthly lodger"}


class BASReminderPlugin(AgentPlugin):
    """Drafts BAS reminders with adaptive monthly/quarterly scheduling per client."""

    PLUGIN_ID   = "plugin_bas_reminder"
    name        = "BAS Reminder Drafter"
    description = ("Calculates BAS due dates per client based on their lodgement "
                   "frequency (monthly or quarterly), learned from XPM and memory.")
    version     = "2.0.0"
    icon        = "📅"
    # Run on the 21st of every month — catches both monthly (due 21st) and
    # quarterly windows in a single pass
    default_schedule = Schedule.monthly_on_day(21)
    category    = PluginCategory.RECEPTION

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()
        # Calendar scheduler ensures this runs on the 21st of each month at 08:00

        today      = date.today()
        lead_days  = int(self.get_plugin_setting("lead_days", "14"))
        confidence = float(self.get_plugin_setting("confidence_threshold", "0.75"))
        default_freq = self.get_plugin_setting("default_frequency", "quarterly")
        learn      = self.get_plugin_setting("learn_from_xpm", "1") == "1"

        if not context.gateway or not context.gateway.is_available("xpm"):
            result.summary = "XPM not configured — BAS reminders skipped."
            return result

        # ── Get GST-registered clients from XPM ───────────────────────────────
        try:
            clients = context.gateway.xpm.list_clients(search="", limit=500)
        except Exception as e:
            result.summary = f"XPM client fetch failed: {e}"
            return result

        drafted  = 0
        skipped  = 0
        learned  = 0

        for client in clients:
            # Filter for GST-registered clients only
            gst_registered = (
                client.get("gst_registered", False)
                or client.get("registered_for_gst", False)
                or str(client.get("gst", "")).lower() in ("yes", "1", "true")
            )
            if not gst_registered:
                continue

            client_name  = client.get("name", client.get("client_name", "Unknown"))
            client_email = client.get("email", "")
            if not client_email:
                skipped += 1
                continue

            # ── Determine lodgement frequency ─────────────────────────────────
            frequency = self._detect_frequency(context, client, client_name, default_freq)

            # ── Store learned frequency in memory ─────────────────────────────
            if learn and context.memory:
                try:
                    context.memory.store_lesson(
                        lesson=(f"{client_name} BAS lodgement frequency: {frequency}"),
                        category="bas_frequency_learned",
                        metadata={"client_name": client_name,
                                  "frequency": frequency,
                                  "detected_date": today.isoformat()})
                    learned += 1
                except Exception:
                    pass

            # ── Calculate upcoming due dates for this client ───────────────────
            upcoming_due_dates = self._get_upcoming_due_dates(frequency, today, lead_days)

            if not upcoming_due_dates:
                continue

            # ── Draft and queue a reminder for each upcoming due date ──────────
            for due_date in upcoming_due_dates:
                # Check if reminder already sent for this client + due date
                already_sent = self._check_already_sent(context, client_name, due_date)
                if already_sent:
                    skipped += 1
                    continue

                email_body = self._draft_reminder(
                    context, client_name, due_date, frequency)
                subject = (f"BAS Lodgement Reminder — Due "
                           f"{due_date.strftime('%d %B %Y')}")

                # ── Submit to approval queue ──────────────────────────────────
                draft_mode = context.draft_mode
                if context.approval_queue:
                    def _send(to=client_email, subj=subject, body=email_body, dm=draft_mode):
                        if context.graph:
                            if dm:
                                context.graph.create_draft(to, subj, body)
                            else:
                                context.graph.send_email(to, subj, body)
                    context.approval_queue.submit(
                        action_type="send_email",
                        description=(f"BAS reminder ({frequency}) to "
                                     f"{client_name} (due {due_date})"),
                        payload={"to": client_email,
                                 "due_date": due_date.isoformat(),
                                 "frequency": frequency},
                        confidence=confidence,
                        plugin_id=self.PLUGIN_ID,
                        execute_callback=_send)
                elif context.graph:
                    try:
                        if draft_mode:
                            context.graph.create_draft(client_email, subject, email_body)
                        else:
                            context.graph.send_email(client_email, subject, email_body)
                    except Exception as e:
                        context.log(
                            f"[BASReminder] Email failed for {client_name}: {e}")

                # ── Store in memory to prevent duplicates ─────────────────────
                if context.memory:
                    try:
                        context.memory.store_client_interaction(
                            client_name=client_name,
                            interaction_type="bas_reminder",
                            summary=(f"BAS reminder ({frequency}) sent for "
                                     f"due date {due_date.isoformat()}"),
                            metadata={"due_date": due_date.isoformat(),
                                      "email": client_email,
                                      "frequency": frequency})
                    except Exception:
                        pass

                drafted += 1
                log_activity(
                    from_email=client_email,
                    subject="BAS Reminder",
                    classification="bas_reminder",
                    action=f"due:{due_date.isoformat()}_freq:{frequency}",
                    draft_created=True)

        result.summary = (
            f"BAS reminders: {drafted} drafted/queued, {skipped} skipped. "
            f"Frequency learned/updated for {learned} clients.")
        result.actions_taken = drafted
        result.items_skipped = skipped
        return result

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _detect_frequency(self, context: PluginContext, client: dict,
                           client_name: str, default_freq: str) -> str:
        """
        Determine a client's BAS lodgement frequency.
        Priority: XPM field → vector memory lesson → default setting.
        Returns 'monthly' or 'quarterly'.
        """
        # 1. Check XPM client record fields
        for field in XPM_FREQUENCY_FIELDS:
            val = str(client.get(field, "")).strip().lower()
            if val:
                if val in MONTHLY_INDICATORS:
                    return "monthly"
                if val in {"quarterly", "quarter", "q", "3", "3m"}:
                    return "quarterly"

        # 2. Check vector memory for previously learned frequency
        if context is not None and context.memory:
            try:
                results = context.memory.search(
                    query=f"{client_name} BAS lodgement frequency",
                    n_results=3,
                    collection="lessons")
                for r in results:
                    meta = r.get("metadata", {})
                    if (meta.get("client_name") == client_name
                            and meta.get("category") == "bas_frequency_learned"):
                        freq = meta.get("frequency", "").lower()
                        if freq in ("monthly", "quarterly"):
                            return freq
            except Exception:
                pass

        # 3. Fall back to configured default
        return default_freq

    def _get_upcoming_due_dates(self, frequency: str, today: date,
                                 lead_days: int) -> list:
        """Return a list of due dates within the lead window for this frequency."""
        upcoming = []

        if frequency == "monthly":
            # Monthly: 21st of the following month
            # Check this month's 21st and next month's 21st
            for offset in range(3):
                month = today.month + offset
                year  = today.year
                while month > 12:
                    month -= 12
                    year  += 1
                try:
                    due = date(year, month, 21)
                except ValueError:
                    continue
                days_until = (due - today).days
                if 0 <= days_until <= lead_days:
                    upcoming.append(due)
        else:
            # Quarterly: ATO extended lodgement dates
            for day, month in QUARTERLY_DUE_DATES:
                year = today.year
                try:
                    due = date(year, month, day)
                except ValueError:
                    due = date(year, month, 28)
                if due < today:
                    try:
                        due = date(year + 1, month, day)
                    except ValueError:
                        due = date(year + 1, month, 28)
                days_until = (due - today).days
                if 0 <= days_until <= lead_days:
                    upcoming.append(due)

        return upcoming

    def _check_already_sent(self, context: PluginContext,
                             client_name: str, due_date: date) -> bool:
        """Return True if a reminder for this client + due date is already in memory."""
        if not context.memory:
            return False
        try:
            history = context.memory.search(
                query=f"BAS reminder {client_name} {due_date.isoformat()}",
                n_results=5,
                collection="client_interactions")
            return any(
                r.get("metadata", {}).get("due_date") == due_date.isoformat()
                and r.get("metadata", {}).get("client_name") == client_name
                and r.get("metadata", {}).get("type") == "bas_reminder"
                for r in history)
        except Exception:
            return False

    def _draft_reminder(self, context: PluginContext, client_name: str,
                         due_date: date, frequency: str) -> str:
        """Use Claude Haiku to draft a frequency-aware BAS reminder email."""
        practice  = get_setting("practice_name", "MC & S Accounting")
        due_str   = due_date.strftime("%d %B %Y")
        days_left = (due_date - date.today()).days
        freq_label = "monthly" if frequency == "monthly" else "quarterly"

        if not context.claude_fast:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>This is a reminder that your {freq_label} BAS lodgement "
                    f"is due on <strong>{due_str}</strong> ({days_left} days away).</p>"
                    f"<p>Please provide your records at your earliest convenience "
                    f"so we can prepare and lodge on time.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
        try:
            prompt = (
                f"Write a professional {freq_label} BAS lodgement reminder email "
                f"from {practice} to {client_name}. "
                f"The BAS is due on {due_str} ({days_left} days away). "
                f"Mention that this is their {freq_label} BAS. "
                "Ask the client to provide their records. Keep it brief and friendly. "
                "Return HTML body only.")
            resp = context.claude_fast.messages.create(
                model=self.get_claude_model_fast(),
                max_tokens=350,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>Your {freq_label} BAS is due on {due_str}. "
                    f"Please provide your records.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
