"""
MC & S Plugin: FuseSign Monitor
==================================
Plugin ID  : plugin_fusesign_monitor
Version    : 1.0.0

WHAT IT DOES
------------
Monitors FuseSign for signing bundle status changes and:

1. Checks all pending envelopes for status updates
2. When a bundle is fully signed — notifies the team via Teams
   and logs the completion in XPM as a job note
3. When a bundle has been pending for too long — sends a reminder
   email to the client and a Teams alert
4. Tracks all signing activity in vector memory
5. Publishes events for other plugins to react to (e.g., trigger
   next workflow step after signing)

APEX ALIGNMENT
--------------
Mirrors APEX's "Document Signing Agent" — closes the loop on
document workflows automatically without manual checking.

SCHEDULE
--------
Default: every 15 minutes.
"""

from datetime import datetime, date, timedelta
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule, PluginCategory
from config import get_setting, log_activity
from client_utils import normalise_client_name


class FuseSignMonitorPlugin(AgentPlugin):
    """Monitors FuseSign envelopes and automates signing follow-up."""

    PLUGIN_ID   = "plugin_fusesign_monitor"
    name        = "FuseSign Monitor"
    description = ("Monitors FuseSign for signing completions and overdue bundles, "
                   "notifies Teams, and logs completions in XPM.")
    version     = "1.0.0"
    icon        = "✍️"
    default_schedule = Schedule.every_minutes(15)
    category    = PluginCategory.RECEPTION

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        if not context.gateway or not context.gateway.is_available("fusesign"):
            result.summary = "FuseSign not configured."
            return result

        context.log("[FuseSignMonitor] Checking envelope statuses...")

        overdue_days = int(self.get_plugin_setting("overdue_days", "7"))
        today = date.today()
        completed = 0
        overdue_count = 0
        reminders_sent = 0

        try:
            # ── Check pending envelopes ───────────────────────────────────────
            pending = context.gateway.fusesign.list_envelopes(
                status="pending", limit=50)

            for env in pending:
                env_id      = env.get("id", env.get("envelope_id", ""))
                env_name    = env.get("name", env.get("subject", "Unknown"))
                client_name = env.get("client_name", env.get("recipient_name", "Unknown"))
                client_email = env.get("recipient_email", env.get("email", ""))
                created_str = env.get("created_at", env.get("created", ""))

                # Check if overdue
                try:
                    created = datetime.strptime(created_str[:10], "%Y-%m-%d").date()
                    days_pending = (today - created).days
                except Exception:
                    days_pending = 0

                if days_pending >= overdue_days:
                    overdue_count += 1

                    # Check reminder history
                    prior_reminders = 0
                    if context.memory:
                        try:
                            history = context.memory.search(
                                query=f"fusesign reminder {env_id}",
                                n_results=5,
                                collection="client_interactions")
                            prior_reminders = sum(
                                1 for r in history
                                if r.get("metadata", {}).get("envelope_id") == env_id
                                and r.get("metadata", {}).get("type") == "signing_reminder")
                        except Exception:
                            pass

                    max_reminders = int(self.get_plugin_setting("max_reminders", "2"))
                    if prior_reminders < max_reminders:
                        # Send Teams alert
                        if (self.get_plugin_setting("send_to_teams", "1") == "1"
                                and context.gateway.is_available("teams")):
                            try:
                                context.gateway.teams.send_alert(
                                    title=f"⏰ Signing Overdue: {env_name}",
                                    message=(f"{client_name} has not signed '{env_name}' "
                                             f"({days_pending} days pending). "
                                             f"Reminder #{prior_reminders + 1} sent."),
                                    color="FF9800")
                            except Exception as e:
                                context.log(f"[FuseSignMonitor] Teams alert failed: {e}")

                        # Send reminder email
                        if (self.get_plugin_setting("send_reminder", "1") == "1"
                                and client_email and context.graph):
                            try:
                                body = self._draft_reminder(
                                    context, client_name, env_name, days_pending)
                                subj = f"Action Required: Please Sign '{env_name}'"
                                if context.draft_mode:
                                    context.graph.create_draft(client_email, subj, body)
                                else:
                                    context.graph.send_email(client_email, subj, body)
                                reminders_sent += 1
                            except Exception as e:
                                context.log(f"[FuseSignMonitor] Reminder email failed: {e}")

                        # Store reminder in memory
                        if context.memory:
                            try:
                                context.memory.store_client_interaction(
                                    client_name=normalise_client_name(client_name),
                                    interaction_type="signing_reminder",
                                    summary=f"Signing reminder #{prior_reminders + 1} for '{env_name}'",
                                    metadata={"envelope_id": env_id,
                                              "days_pending": str(days_pending)})
                            except Exception:
                                pass

            # ── Check recently completed envelopes ────────────────────────────
            completed_envs = context.gateway.fusesign.list_envelopes(
                status="completed", limit=20)

            for env in completed_envs:
                env_id      = env.get("id", env.get("envelope_id", ""))
                env_name    = env.get("name", env.get("subject", "Unknown"))
                client_name = env.get("client_name", env.get("recipient_name", "Unknown"))
                completed_at = env.get("completed_at", env.get("signed_at", ""))

                # Only process completions from the last 30 minutes
                already_processed = False
                if context.memory:
                    try:
                        history = context.memory.search(
                            query=f"fusesign completed {env_id}",
                            n_results=3,
                            collection="client_interactions")
                        already_processed = any(
                            r.get("metadata", {}).get("envelope_id") == env_id
                            and r.get("metadata", {}).get("type") == "signing_completed"
                            for r in history)
                    except Exception:
                        pass

                if already_processed:
                    continue

                completed += 1

                # Notify Teams
                if (self.get_plugin_setting("send_to_teams", "1") == "1"
                        and context.gateway.is_available("teams")):
                    try:
                        context.gateway.teams.send_alert(
                            title=f"✅ Signing Complete: {env_name}",
                            message=f"{client_name} has signed '{env_name}'.",
                            color="00C853")
                    except Exception as e:
                        context.log(f"[FuseSignMonitor] Teams notification failed: {e}")

                # Log to XPM
                if (self.get_plugin_setting("log_to_xpm", "1") == "1"
                        and context.gateway.is_available("xpm")):
                    try:
                        xpm_clients = context.gateway.xpm.list_clients(
                            search=client_name, limit=1)
                        if xpm_clients:
                            client_id = xpm_clients[0].get("id", "")
                            context.gateway.xpm.add_client_note(
                                client_id=client_id,
                                note=f"FuseSign: '{env_name}' signed by {client_name} "
                                     f"on {completed_at[:10] if completed_at else 'today'}.")
                    except Exception as e:
                        context.log(f"[FuseSignMonitor] XPM note failed: {e}")

                # Store in memory
                if context.memory:
                    try:
                        context.memory.store_client_interaction(
                            client_name=normalise_client_name(client_name),
                            interaction_type="signing_completed",
                            summary=f"'{env_name}' signed and completed.",
                            metadata={"envelope_id": env_id,
                                      "completed_at": completed_at})
                    except Exception:
                        pass

                # Publish event
                if context.event_bus:
                    context.event_bus.publish(
                        "fusesign.envelope.completed",
                        source=self.PLUGIN_ID,
                        payload={"envelope_id": env_id, "envelope_name": env_name,
                                 "client_name": client_name})

                log_activity(from_email="fusesign", subject="Document Signed", classification="signing_completed",
                             action=f"envelope:{env_name}", draft_created=False)

        except Exception as e:
            result.summary = f"FuseSign monitor error: {e}"
            return result

        result.summary = (f"FuseSign monitor: {completed} completed, "
                          f"{overdue_count} overdue, {reminders_sent} reminders sent.")
        result.actions_taken = completed + reminders_sent
        return result

    def _draft_reminder(self, context: PluginContext, client_name: str,
                         env_name: str, days_pending: int) -> str:
        """Draft a signing reminder email."""
        practice = get_setting("practice_name", "MC & S Accounting")
        if not context.claude_fast:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>We wanted to follow up on the document <strong>'{env_name}'</strong> "
                    f"that was sent to you for signing {days_pending} days ago.</p>"
                    f"<p>Please sign at your earliest convenience.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
        try:
            prompt = (f"Write a brief, polite reminder email from {practice} to {client_name} "
                      f"asking them to sign the document '{env_name}' which has been pending "
                      f"for {days_pending} days. Return HTML body only.")
            resp = context.claude_fast.messages.create(
                model=self.get_claude_model_fast(),
                max_tokens=300,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception:
            return (f"<p>Dear {client_name},</p>"
                    f"<p>This is a reminder to sign '{env_name}'.</p>"
                    f"<p>Kind regards,<br>{practice}</p>")
