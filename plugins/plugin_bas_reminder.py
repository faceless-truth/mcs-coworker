"""
MC & S Plugin: BAS Reminder Drafter
======================================
Plugin ID  : plugin_bas_reminder
Version    : 3.0.0

WHAT IT DOES
------------
Drives BAS reminders from the accountant-managed BAS Client List (managed in
Settings → BAS Client Tracking). Each row carries its own lodgement frequency,
contact email, and workflow status. The plugin:

1. Loads every BAS client from the ``bas_clients`` SQLite table.
2. Reads the Knowledge Base entry "BAS Lodgement Policy" to include the firm's
   own deadline text in the draft email (ATO due dates are also calculated
   algorithmically so the logic still works even if the KB entry is empty).
3. For each client:
     - Checks semantic memory for a recent inbound email referencing BAS /
       records / figures. When one is found the client's
       ``last_data_received`` is updated and status is bumped to
       ``Data Received``.
     - If status is ``Data Received`` or ``Lodged``, no reminder is sent.
     - If status is ``Awaiting Data`` and a deadline falls within the
       configured lead window, drafts a reminder via Claude Haiku and
       stamps ``last_reminder_sent``.

STATUS WORKFLOW
---------------
``Awaiting Data``  →  ``Data Received``  →  ``Lodged``
(Staff flip Lodged manually in Settings once they've submitted the BAS.)
"""

from datetime import datetime, date, timedelta
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule, PluginCategory
from config import (
    get_setting, log_activity,
    get_bas_clients, update_bas_client,
    get_knowledge_entries,
)


# ATO quarterly BAS due dates (day, month) — lodgement agent extended dates
QUARTERLY_DUE_DATES = [
    (28, 10),   # Q1: Jul-Sep due 28 Oct
    (28, 2),    # Q2: Oct-Dec due 28 Feb
    (28, 4),    # Q3: Jan-Mar due 28 Apr
    (28, 7),    # Q4: Apr-Jun due 28 Jul
]

# Annual BAS (for eligible small businesses lodging an Annual GST Return)
ANNUAL_DUE_DATE = (28, 2)  # 28 February of the following year

BAS_DATA_KEYWORDS = (
    "bas", "records", "figures", "bookkeeping", "payroll summary",
    "gst", "quarterly figures",
)


class BASReminderPlugin(AgentPlugin):
    """Drafts BAS reminders driven by the local BAS Client Tracking list."""

    PLUGIN_ID   = "plugin_bas_reminder"
    name        = "BAS Reminder Drafter"
    description = ("Reminds clients on the BAS Client List that their BAS is "
                   "coming due, tracks when data arrives, and updates status.")
    version     = "3.0.0"
    icon        = "📅"
    # Daily tick — the lead-window gate does the actual 'should I send now' work
    # so the plugin can react to status changes mid-cycle.
    default_schedule = Schedule.daily_at(8)
    category    = PluginCategory.RECEPTION

    @classmethod
    def settings_schema(cls):
        return [
            {"key": "lead_days", "label": "Days before due date to remind",
             "type": "number", "default": "14"},
            {"key": "confidence_threshold", "label": "Approval confidence",
             "type": "number", "default": "0.75"},
            {"key": "memory_lookback_days",
             "label": "Treat inbound BAS emails within N days as 'data received'",
             "type": "number", "default": "45"},
        ]

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()
        today      = date.today()
        lead_days  = int(self.get_plugin_setting("lead_days", "14"))
        confidence = float(self.get_plugin_setting("confidence_threshold", "0.75"))
        lookback   = int(self.get_plugin_setting("memory_lookback_days", "45"))

        clients = get_bas_clients()
        if not clients:
            result.summary = ("No BAS clients configured. Add them in "
                              "Settings → BAS Client Tracking.")
            return result

        policy_text = self._get_policy_text()

        drafted  = 0
        skipped  = 0
        updated  = 0

        for client in clients:
            client_name  = (client.get("client_name") or "").strip()
            client_email = (client.get("client_email") or "").strip()
            frequency    = (client.get("frequency") or "Quarterly").strip()
            status       = (client.get("status") or "Awaiting Data").strip()

            if not client_name or not client_email:
                skipped += 1
                continue

            # ── Cross-reference memory for inbound BAS data ──────────────────
            received_ts = self._detect_received(context, client, lookback)
            if received_ts:
                changes: dict = {"last_data_received": received_ts.date().isoformat()}
                if status == "Awaiting Data":
                    changes["status"] = "Data Received"
                    status = "Data Received"
                try:
                    update_bas_client(client["id"], changes)
                    updated += 1
                except Exception as e:
                    context.log(f"[BASReminder] status update failed for {client_name}: {e}")

            # ── Skip clients that are no longer waiting for data ─────────────
            if status in ("Data Received", "Lodged"):
                continue

            # ── Calculate upcoming due dates for this client ─────────────────
            upcoming = self._get_upcoming_due_dates(frequency, today, lead_days)
            if not upcoming:
                continue

            for due_date in upcoming:
                # Don't hammer the client — skip if we already reminded them
                # for this cycle.
                if self._already_reminded(client, due_date):
                    skipped += 1
                    continue

                subject = (f"BAS Lodgement Reminder — Due "
                           f"{due_date.strftime('%d %B %Y')}")
                body    = self._draft_reminder(
                    context, client_name, due_date, frequency, policy_text)
                draft_mode = context.draft_mode

                if context.approval_queue:
                    def _send(to=client_email, subj=subject, b=body, dm=draft_mode):
                        if context.graph:
                            if dm:
                                context.graph.create_draft(to, subj, b)
                            else:
                                context.graph.send_email(to, subj, b)
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
                            context.graph.create_draft(client_email, subject, body)
                        else:
                            context.graph.send_email(client_email, subject, body)
                    except Exception as e:
                        context.log(
                            f"[BASReminder] Email failed for {client_name}: {e}")

                try:
                    update_bas_client(client["id"],
                                       {"last_reminder_sent": today.isoformat()})
                except Exception:
                    pass

                if context.memory:
                    try:
                        context.memory.store_client_interaction(
                            client_name=client_name,
                            client_email=client_email,
                            interaction_type="bas_reminder",
                            summary=(f"BAS reminder ({frequency}) sent for "
                                     f"due date {due_date.isoformat()}"),
                            metadata={"due_date": due_date.isoformat(),
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
            f"Status updated for {updated} clients.")
        result.actions_taken = drafted
        result.items_skipped = skipped
        return result

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _get_policy_text(self) -> str:
        """Return the firm's BAS Lodgement Policy content from the KB, if any."""
        try:
            for entry in get_knowledge_entries(only_enabled=True):
                title = (entry.get("title") or "").strip().lower()
                if "bas" in title and "lodgement" in title:
                    return (entry.get("content") or "").strip()
        except Exception:
            pass
        return ""

    def _detect_received(self, context: PluginContext, client: dict,
                          lookback_days: int):
        """Return the timestamp of a recent inbound BAS email from this client,
        or None if nothing qualifies.
        """
        if not context.memory:
            return None
        client_email = (client.get("client_email") or "").strip().lower()
        if not client_email:
            return None

        cutoff = datetime.now() - timedelta(days=lookback_days)
        last_received = client.get("last_data_received") or ""

        try:
            hits = context.memory.search(
                query=f"BAS records figures {client.get('client_name', '')}",
                n_results=10,
                collection="interactions",
            )
        except Exception:
            return None

        newest = None
        for hit in hits:
            meta = hit.get("metadata", {}) or {}
            hit_email = str(meta.get("client_email", "")).lower()
            if hit_email and hit_email != client_email:
                continue
            ts_str = str(meta.get("timestamp", ""))
            try:
                ts = datetime.fromisoformat(ts_str)
            except Exception:
                continue
            if ts < cutoff:
                continue
            text = (hit.get("content") or "").lower()
            if not any(k in text for k in BAS_DATA_KEYWORDS):
                continue
            # Only interactions coming IN from the client count as 'data
            # received' — our own outbound reminders should be ignored.
            itype = str(meta.get("interaction_type", "")).lower()
            if itype in ("bas_reminder", "outbound", "sent"):
                continue
            if newest is None or ts > newest:
                newest = ts

        if newest is None:
            return None
        # Don't downgrade an already-stamped later date.
        if last_received:
            try:
                prev = datetime.fromisoformat(last_received)
                if prev >= newest:
                    return None
            except Exception:
                pass
        return newest

    def _already_reminded(self, client: dict, due_date: date) -> bool:
        """Only send one reminder per due-date cycle. Uses the stamped date to
        approximate: if a reminder was sent within the lead window of the
        upcoming due date, skip."""
        last = client.get("last_reminder_sent") or ""
        if not last:
            return False
        try:
            sent = date.fromisoformat(last)
        except Exception:
            return False
        return 0 <= (due_date - sent).days <= 30

    def _get_upcoming_due_dates(self, frequency: str, today: date,
                                 lead_days: int) -> list:
        """Return due dates within the lead window for the given frequency."""
        upcoming: list = []
        freq = (frequency or "").strip().lower()

        if freq == "monthly":
            # Monthly lodgers: BAS due 21st of the following month.
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

        elif freq == "annual":
            day, month = ANNUAL_DUE_DATE
            for year_offset in (0, 1):
                year = today.year + year_offset
                try:
                    due = date(year, month, day)
                except ValueError:
                    due = date(year, month, 28)
                if due < today:
                    continue
                days_until = (due - today).days
                if 0 <= days_until <= lead_days:
                    upcoming.append(due)
                    break

        else:
            # Default / quarterly.
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

    def _draft_reminder(self, context: PluginContext, client_name: str,
                         due_date: date, frequency: str, policy_text: str) -> str:
        """Use Claude Haiku (if available) to draft the reminder email."""
        practice  = get_setting("practice_name", "MC & S Accounting")
        due_str   = due_date.strftime("%d %B %Y")
        days_left = (due_date - date.today()).days
        freq_label = (frequency or "quarterly").lower()

        if not context.claude_fast:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>This is a reminder that your {freq_label} BAS lodgement "
                    f"is due on <strong>{due_str}</strong> ({days_left} days away).</p>"
                    f"<p>Please provide your records at your earliest convenience "
                    f"so we can prepare and lodge on time.</p>")
        try:
            policy_block = (
                f"\n\nFirm BAS policy (for reference, do not quote verbatim):\n{policy_text}"
                if policy_text else "")
            prompt = (
                f"Write a professional {freq_label} BAS lodgement reminder email "
                f"from {practice} to {client_name}. "
                f"The BAS is due on {due_str} ({days_left} days away). "
                f"Mention that this is their {freq_label} BAS. "
                f"Ask the client to provide their records. Keep it brief and friendly. "
                f"Return HTML body only.\n\n"
                f"Do NOT include any sign-off, closing, or signature in your "
                f"draft. No Kind regards, no Best regards, no name, no company "
                f"name. The email signature is appended automatically — just end "
                f"with your last sentence of actual content."
                f"{policy_block}")
            resp = context.claude_fast.messages.create(
                model=self.get_claude_model_fast(),
                max_tokens=350,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>Your {freq_label} BAS is due on {due_str}. "
                    f"Please provide your records.</p>")
