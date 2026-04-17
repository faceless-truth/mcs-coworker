"""
MC & S Plugin: Debtor Follow-Up
=================================
Plugin ID  : plugin_debtor_followup
Version    : 1.0.0

WHAT IT DOES
------------
Queries XPM for clients with overdue invoices and automates
the follow-up process:

1. Retrieves a list of clients with outstanding balances from XPM
2. Checks vector memory for previous follow-up history per client
3. Uses Claude Haiku to draft a personalised, polite follow-up email
4. Submits the email to the approval queue (confidence-based)
5. Logs all follow-up attempts in memory for escalation tracking
6. Escalates to Teams alert if a client has been overdue 60+ days

APEX ALIGNMENT
--------------
Mirrors APEX's "Debtor Management Agent" — automated, tiered
follow-up that respects client relationships while protecting
the practice's cash flow.

SCHEDULE
--------
Default: every Wednesday at 10:00 AM.
"""

from datetime import datetime, date
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity


class DebtorFollowUpPlugin(AgentPlugin):
    """Automated debtor follow-up with tiered escalation."""

    PLUGIN_ID   = "plugin_debtor_followup"
    NAME        = "Debtor Follow-Up"
    DESCRIPTION = ("Finds overdue invoices in XPM, drafts personalised follow-up "
                   "emails using AI, and escalates persistent debtors to Teams.")
    VERSION     = "1.0.0"
    ICON        = "💰"
    SCHEDULE    = Schedule.every_hours(168)

    DEFAULT_SETTINGS = {
        "run_day":               "3",   # Wednesday
        "run_hour":              "10",
        "overdue_threshold_days": "14", # start following up after 14 days
        "escalate_days":         "60",  # escalate to Teams after 60 days
        "max_follow_ups":        "3",   # max automated follow-ups before manual
        "confidence_threshold":  "0.7", # approval queue threshold
    }

    def __init__(self):
        super().__init__()
        self._last_run_week: str = ""

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        # ── Day/hour gate ─────────────────────────────────────────────────────
        now = datetime.now()
        target_day  = int(self.get_plugin_setting("run_day",  "3"))
        target_hour = int(self.get_plugin_setting("run_hour", "10"))
        if (now.weekday() + 1) != target_day or now.hour != target_hour:
            result.summary = "Not debtor follow-up time."
            return result

        week_str = f"{date.today().isocalendar()[0]}-W{date.today().isocalendar()[1]}"
        if self._last_run_week == week_str:
            result.summary = "Debtor follow-up already run this week."
            return result

        if not context.gateway or not context.gateway.is_available("xpm"):
            result.summary = "XPM not configured — debtor follow-up skipped."
            return result

        context.log("[DebtorFollowUp] Fetching overdue clients from XPM...")

        overdue_days = int(self.get_plugin_setting("overdue_threshold_days", "14"))
        escalate_days = int(self.get_plugin_setting("escalate_days", "60"))
        max_followups = int(self.get_plugin_setting("max_follow_ups", "3"))
        confidence = float(self.get_plugin_setting("confidence_threshold", "0.7"))

        # ── Fetch overdue clients ─────────────────────────────────────────────
        try:
            clients = context.gateway.xpm.list_clients(search="", limit=200)
        except Exception as e:
            result.summary = f"XPM client fetch failed: {e}"
            return result

        drafted = 0
        escalated = 0
        skipped = 0

        for client in clients:
            client_name  = client.get("name", client.get("client_name", "Unknown"))
            client_email = client.get("email", "")
            balance      = float(client.get("outstanding_balance",
                                            client.get("balance", 0)) or 0)
            days_overdue = int(client.get("days_overdue", 0) or 0)

            if balance <= 0 or days_overdue < overdue_days:
                continue

            # ── Check follow-up history ───────────────────────────────────────
            prior_followups = 0
            if context.memory:
                try:
                    history = context.memory.search(
                        query=f"debtor follow-up {client_name}",
                        n_results=10,
                        collection="client_interactions")
                    prior_followups = sum(
                        1 for r in history
                        if r.get("metadata", {}).get("type") == "debtor_followup"
                        and r.get("metadata", {}).get("client_name") == client_name)
                except Exception:
                    pass

            if prior_followups >= max_followups:
                context.log(f"[DebtorFollowUp] {client_name}: max follow-ups reached, skipping.")
                skipped += 1
                continue

            # ── Escalate to Teams if very overdue ─────────────────────────────
            if days_overdue >= escalate_days:
                if context.gateway and context.gateway.is_available("teams"):
                    try:
                        context.gateway.teams.send_alert(
                            title=f"⚠️ Debtor Escalation: {client_name}",
                            message=(f"{client_name} has an outstanding balance of "
                                     f"${balance:,.2f} and is {days_overdue} days overdue. "
                                     f"Manual follow-up required."),
                            color="FF0000")
                        escalated += 1
                    except Exception as e:
                        context.log(f"[DebtorFollowUp] Teams escalation failed: {e}")
                continue

            if not client_email:
                context.log(f"[DebtorFollowUp] {client_name}: no email address, skipping.")
                skipped += 1
                continue

            # ── Draft follow-up email ─────────────────────────────────────────
            email_body = self._draft_followup_email(
                context, client_name, client_email, balance, days_overdue, prior_followups)

            # ── Submit to approval queue ──────────────────────────────────────
            subject = f"Outstanding Invoice Reminder — {client_name}"
            if context.approval_queue:
                def _send(to=client_email, subj=subject, body=email_body):
                    if context.graph:
                        context.graph.send_email(
                            to=to, subject=subj, body=body,
                            draft_mode=context.draft_mode)
                approved = context.approval_queue.submit(
                    action_type="send_email",
                    description=f"Debtor follow-up to {client_name} ({client_email})",
                    payload={"to": client_email, "subject": subject,
                             "balance": balance, "days_overdue": days_overdue},
                    confidence=confidence,
                    plugin_id=self.PLUGIN_ID,
                    execute_callback=_send,
                    recipient=client_email,
                    subject=subject,
                    body_preview=email_body[:300])
                if approved:
                    drafted += 1
                else:
                    drafted += 1  # queued counts as drafted
            elif context.graph:
                try:
                    context.graph.send_email(
                        to=client_email, subject=subject, body=email_body,
                        draft_mode=context.draft_mode)
                    drafted += 1
                except Exception as e:
                    context.log(f"[DebtorFollowUp] Email failed for {client_name}: {e}")

            # ── Store in memory ───────────────────────────────────────────────
            if context.memory:
                try:
                    context.memory.store_client_interaction(
                        client_name=client_name,
                        interaction_type="debtor_followup",
                        summary=f"Follow-up #{prior_followups + 1} sent. "
                                f"Balance: ${balance:,.2f}, {days_overdue} days overdue.",
                        metadata={"balance": str(balance),
                                  "days_overdue": str(days_overdue),
                                  "followup_number": str(prior_followups + 1)})
                except Exception:
                    pass

        self._last_run_week = week_str
        result.summary = (f"Debtor follow-up complete. Drafted/queued: {drafted}, "
                          f"Escalated: {escalated}, Skipped: {skipped}.")
        result.actions_taken = drafted + escalated
        result.items_skipped = skipped
        log_activity(from_email="system", subject="Debtor Follow-Up", classification="debtor_followup",
                     action=f"drafted:{drafted}_escalated:{escalated}",
                     draft_created=drafted > 0)
        return result

    def _draft_followup_email(self, context: PluginContext, client_name: str,
                               client_email: str, balance: float,
                               days_overdue: int, prior_followups: int) -> str:
        """Use Claude Haiku to draft a personalised follow-up email."""
        practice = get_setting("practice_name", "MC & S Accounting")
        tone = "gentle" if prior_followups == 0 else ("firm" if prior_followups == 1 else "urgent")

        if not context.claude_fast:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>We wish to draw your attention to an outstanding invoice of "
                    f"${balance:,.2f} which is now {days_overdue} days overdue.</p>"
                    f"<p>Please arrange payment at your earliest convenience.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
        try:
            prompt = (f"Write a {tone} email from {practice} to {client_name} "
                      f"({client_email}) regarding an outstanding invoice of "
                      f"${balance:,.2f} that is {days_overdue} days overdue. "
                      f"This is follow-up number {prior_followups + 1}. "
                      "Keep it professional, concise, and relationship-preserving. "
                      "Include a clear call to action. Return HTML body only.")
            resp = context.claude_fast.messages.create(
                model=self.get_claude_model_fast(),
                max_tokens=400,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>We wish to draw your attention to an outstanding invoice of "
                    f"${balance:,.2f} which is now {days_overdue} days overdue.</p>"
                    f"<p>Please arrange payment at your earliest convenience.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
