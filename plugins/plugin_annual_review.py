"""
MC & S Plugin: Annual Review Prompt
======================================
Plugin ID  : plugin_annual_review
Version    : 1.0.0

WHAT IT DOES
------------
Identifies clients who are due for an annual review and proactively
reaches out to book a meeting:

1. Queries XPM for clients whose last review/meeting was 11+ months ago
2. Checks vector memory for the last interaction date per client
3. Uses Claude Haiku to draft a personalised annual review invitation
4. Submits the email to the approval queue
5. Stores the outreach in memory to prevent duplicate contact
6. Publishes an event so other plugins can react (e.g., Meeting Prep)

APEX ALIGNMENT
--------------
Mirrors APEX's "Annual Review Agent" — proactive relationship
management that drives recurring revenue and client retention.

SCHEDULE
--------
Default: every Monday at 8:30 AM.
"""

from datetime import datetime, date, timedelta
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity


class AnnualReviewPlugin(AgentPlugin):
    """Identifies clients due for annual review and sends personalised invitations."""

    PLUGIN_ID   = "plugin_annual_review"
    NAME        = "Annual Review Prompt"
    DESCRIPTION = ("Finds clients due for their annual review, drafts personalised "
                   "invitations using AI, and queues them for approval.")
    VERSION     = "1.0.0"
    ICON        = "🔄"
    SCHEDULE    = Schedule.every_hours(168)

    DEFAULT_SETTINGS = {
        "run_day":              "1",     # Monday
        "run_hour":             "8",
        "run_minute":           "30",
        "review_interval_days": "335",   # ~11 months
        "confidence_threshold": "0.75",
        "max_per_run":          "20",
    }

    def __init__(self):
        super().__init__()
        self._last_run_week: str = ""

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        now = datetime.now()
        target_day  = int(self.get_plugin_setting("run_day",  "1"))
        target_hour = int(self.get_plugin_setting("run_hour", "8"))
        if (now.weekday() + 1) != target_day or now.hour != target_hour:
            result.summary = "Not annual review time."
            return result

        week_str = f"{date.today().isocalendar()[0]}-W{date.today().isocalendar()[1]}"
        if self._last_run_week == week_str:
            result.summary = "Annual review prompts already sent this week."
            return result

        if not context.gateway or not context.gateway.is_available("xpm"):
            result.summary = "XPM not configured — annual review skipped."
            return result

        review_interval = int(self.get_plugin_setting("review_interval_days", "335"))
        confidence      = float(self.get_plugin_setting("confidence_threshold", "0.75"))
        max_per_run     = int(self.get_plugin_setting("max_per_run", "20"))
        today           = date.today()
        cutoff          = today - timedelta(days=review_interval)

        context.log("[AnnualReview] Fetching clients from XPM...")

        try:
            clients = context.gateway.xpm.list_clients(search="", limit=500)
        except Exception as e:
            result.summary = f"XPM client fetch failed: {e}"
            return result

        due_for_review = []
        for client in clients:
            client_name  = client.get("name", client.get("client_name", "Unknown"))
            client_email = client.get("email", "")
            if not client_email:
                continue

            # Check last interaction in memory
            last_contact = None
            if context.memory:
                try:
                    history = context.memory.search(
                        query=f"client interaction {client_name}",
                        n_results=5,
                        collection="client_interactions")
                    if history:
                        dates = []
                        for r in history:
                            ts = r.get("metadata", {}).get("timestamp", "")
                            if ts:
                                try:
                                    dates.append(datetime.fromisoformat(ts[:10]).date())
                                except Exception:
                                    pass
                        if dates:
                            last_contact = max(dates)
                except Exception:
                    pass

            # Also check XPM last activity
            if not last_contact:
                xpm_last = client.get("last_activity", client.get("last_contact", ""))
                if xpm_last:
                    try:
                        last_contact = datetime.strptime(
                            xpm_last[:10], "%Y-%m-%d").date()
                    except Exception:
                        pass

            if last_contact and last_contact > cutoff:
                continue  # recently contacted

            # Check if already prompted this cycle
            already_prompted = False
            if context.memory:
                try:
                    history = context.memory.search(
                        query=f"annual review {client_name}",
                        n_results=3,
                        collection="client_interactions")
                    for r in history:
                        meta = r.get("metadata", {})
                        if meta.get("type") == "annual_review_prompt":
                            ts = meta.get("timestamp", "")
                            if ts:
                                try:
                                    sent_date = datetime.fromisoformat(ts[:10]).date()
                                    if (today - sent_date).days < review_interval:
                                        already_prompted = True
                                        break
                                except Exception:
                                    pass
                except Exception:
                    pass

            if not already_prompted:
                due_for_review.append(client)

        context.log(f"[AnnualReview] {len(due_for_review)} clients due for review.")

        due_for_review = due_for_review[:max_per_run]
        drafted = 0

        for client in due_for_review:
            client_name  = client.get("name", client.get("client_name", "Unknown"))
            client_email = client.get("email", "")

            email_body = self._draft_invitation(context, client_name)
            subject    = f"Annual Review — {get_setting('practice_name', 'MC & S Accounting')}"

            if context.approval_queue:
                def _send(to=client_email, subj=subject, body=email_body):
                    if context.graph:
                        context.graph.send_email(
                            to=to, subject=subj, body=body,
                            draft_mode=context.draft_mode)
                context.approval_queue.submit(
                    action_type="send_email",
                    description=f"Annual review invitation to {client_name}",
                    payload={"to": client_email, "client": client_name},
                    confidence=confidence,
                    plugin_id=self.PLUGIN_ID,
                    execute_callback=_send)
            elif context.graph:
                try:
                    context.graph.send_email(
                        to=client_email, subject=subject, body=email_body,
                        draft_mode=context.draft_mode)
                except Exception as e:
                    context.log(f"[AnnualReview] Email failed for {client_name}: {e}")

            # Store in memory
            if context.memory:
                try:
                    context.memory.store_client_interaction(
                        client_name=client_name,
                        interaction_type="annual_review_prompt",
                        summary="Annual review invitation sent.",
                        metadata={"email": client_email,
                                  "timestamp": today.isoformat()})
                except Exception:
                    pass

            # Publish event
            if context.event_bus:
                context.event_bus.publish(
                    "annual_review.prompted",
                    source=self.PLUGIN_ID,
                    payload={"client_name": client_name, "email": client_email})

            drafted += 1
            log_activity(from_email=client_email,
                         subject="Annual Review Invitation",
                         classification="annual_review",
                         action="invitation_sent",
                         draft_created=True)

        self._last_run_week = week_str
        result.summary = (f"Annual review: {drafted} invitations sent/queued "
                          f"from {len(due_for_review)} clients due.")
        result.actions_taken = drafted
        return result

    def _draft_invitation(self, context: PluginContext, client_name: str) -> str:
        """Use Claude Haiku to draft a personalised annual review invitation."""
        practice = get_setting("practice_name", "MC & S Accounting")
        if not context.claude_fast:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>It has been a while since we last connected and we would love "
                    f"to schedule your annual review meeting.</p>"
                    f"<p>Please reply to arrange a convenient time.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
        try:
            prompt = (f"Write a warm, professional annual review invitation email from "
                      f"{practice} to {client_name}. Mention it has been about a year, "
                      "invite them to book a meeting, and highlight the value of the review "
                      "(tax planning, business health, goals). Keep it under 150 words. "
                      "Return HTML body only.")
            resp = context.claude_fast.messages.create(
                model=self.get_claude_model_fast(),
                max_tokens=300,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>We would like to invite you to your annual review meeting.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
