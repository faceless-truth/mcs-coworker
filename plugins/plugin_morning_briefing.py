"""
MC & S Plugin: Morning Briefing
================================
Plugin ID  : plugin_morning_briefing
Version    : 1.0.0

WHAT IT DOES
------------
Runs once each business morning (default 8:00 AM) and compiles a
personalised daily briefing for the practice. It:

1. Queries XPM for today's jobs due, overdue WIP, and upcoming deadlines
2. Scans the inbox for unread emails flagged as high-priority
3. Retrieves recent client interaction history from vector memory
4. Uses Claude Sonnet to synthesise a concise morning briefing
5. Sends the briefing as a Microsoft Teams Adaptive Card
6. Optionally emails the briefing to nominated staff

APEX ALIGNMENT
--------------
This plugin directly mirrors APEX's "Morning Briefing Agent" — the
proactive daily summary that gives the team situational awareness
before the day begins. It uses:
  - context.claude_reason  (Sonnet) for synthesis
  - context.gateway.xpm    for job/deadline data
  - context.gateway.teams  for Teams delivery
  - context.memory         for client history context
  - Event Bus heartbeat     for schedule triggering

SCHEDULE
--------
Default: once daily at 8:00 AM (business days only).
"""

import json
from datetime import datetime, date
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity, get_active_lessons


class MorningBriefingPlugin(AgentPlugin):
    """Compiles and delivers a daily morning briefing to the team."""

    PLUGIN_ID   = "plugin_morning_briefing"
    NAME        = "Morning Briefing"
    DESCRIPTION = ("Compiles a daily morning briefing from XPM jobs, inbox, and client "
                   "history, then delivers it to Teams and/or email at 8 AM.")
    VERSION     = "1.0.0"
    ICON        = "🌅"
    SCHEDULE    = Schedule.daily_at(8)    # once per day; business-hours check inside

    # ── Plugin settings ───────────────────────────────────────────────────────

    DEFAULT_SETTINGS = {
        "briefing_hour":         "8",     # hour to fire (24h)
        "briefing_minute":       "0",
        "send_to_teams":         "1",
        "send_email":            "0",
        "email_recipients":      "",      # comma-separated
        "max_jobs_shown":        "10",
        "include_inbox_summary": "1",
        "include_memory_context":"1",
    }

    def __init__(self):
        super().__init__()
        self._last_run_date: str = ""     # prevent double-firing on same day

    # ── Main entry point ──────────────────────────────────────────────────────

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        # ── Time gate: only fire once per day at the configured hour ──────────
        now = datetime.now()
        target_hour   = int(self.get_plugin_setting("briefing_hour",   "8"))
        target_minute = int(self.get_plugin_setting("briefing_minute", "0"))

        if now.hour != target_hour:
            result.summary = f"Not briefing time yet (target {target_hour:02d}:{target_minute:02d})."
            return result

        today_str = date.today().isoformat()
        if self._last_run_date == today_str:
            result.summary = "Morning briefing already sent today."
            return result

        # ── Business hours check ──────────────────────────────────────────────
        if not self._is_business_day(now):
            result.summary = "Not a business day — skipping morning briefing."
            return result

        context.log(f"[MorningBriefing] Compiling briefing for {today_str}...")

        # ── 1. Gather XPM data ────────────────────────────────────────────────
        xpm_section = self._gather_xpm_data(context)

        # ── 2. Gather inbox summary ───────────────────────────────────────────
        inbox_section = ""
        if self.get_plugin_setting("include_inbox_summary", "1") == "1":
            inbox_section = self._gather_inbox_summary(context)

        # ── 3. Gather memory context ──────────────────────────────────────────
        memory_section = ""
        if self.get_plugin_setting("include_memory_context", "1") == "1":
            memory_section = self._gather_memory_context(context, today_str)

        # ── 4. Synthesise with Claude Sonnet ──────────────────────────────────
        briefing_text = self._synthesise_briefing(
            context, today_str, xpm_section, inbox_section, memory_section)

        # ── 5. Deliver ────────────────────────────────────────────────────────
        delivered = []

        if self.get_plugin_setting("send_to_teams", "1") == "1":
            if context.gateway and context.gateway.is_available("teams"):
                try:
                    context.gateway.teams.send_alert(
                        title=f"☀️ Morning Briefing — {today_str}",
                        message=briefing_text[:3000],
                        color="0078D4")
                    delivered.append("Teams")
                    context.log("[MorningBriefing] Sent to Teams.")
                except Exception as e:
                    context.log(f"[MorningBriefing] Teams delivery failed: {e}")
            else:
                context.log("[MorningBriefing] Teams not configured — skipping.")

        if self.get_plugin_setting("send_email", "0") == "1":
            recipients = self.get_plugin_setting("email_recipients", "")
            if recipients and context.graph:
                for email in [r.strip() for r in recipients.split(",") if r.strip()]:
                    try:
                        context.graph.send_email(
                            to=email,
                            subject=f"Morning Briefing — {today_str}",
                            body=f"<pre>{briefing_text}</pre>",
                            draft_mode=context.draft_mode)
                        delivered.append(f"email:{email}")
                    except Exception as e:
                        context.log(f"[MorningBriefing] Email to {email} failed: {e}")

        # ── 6. Store in memory ────────────────────────────────────────────────
        if context.memory:
            try:
                context.memory.store(
                    document=briefing_text,
                    metadata={"type": "morning_briefing", "date": today_str,
                              "delivered_to": ", ".join(delivered)},
                    collection="general",
                    doc_id=f"briefing_{today_str}")
            except Exception:
                pass

        # ── 7. Publish event ──────────────────────────────────────────────────
        if context.event_bus:
            context.event_bus.publish("morning_briefing.sent",
                                      source=self.PLUGIN_ID,
                                      payload={"date": today_str,
                                               "delivered_to": delivered})

        self._last_run_date = today_str
        result.summary = (f"Morning briefing sent for {today_str}. "
                          f"Delivered via: {', '.join(delivered) or 'none configured'}.")
        result.actions_taken = 1
        log_activity(from_email="system", subject="Morning Briefing", classification="morning_briefing",
                     action="briefing_sent", draft_created=False)
        return result

    # ── Data gathering helpers ────────────────────────────────────────────────

    def _gather_xpm_data(self, context: PluginContext) -> str:
        """Pull today's jobs and overdue WIP from XPM."""
        if not context.gateway or not context.gateway.is_available("xpm"):
            return "XPM not connected — job data unavailable."
        try:
            max_jobs = int(self.get_plugin_setting("max_jobs_shown", "10"))
            jobs = context.gateway.xpm.list_jobs(
                status="inprogress", limit=max_jobs)
            if not jobs:
                return "No active jobs found in XPM."
            lines = [f"Active Jobs ({len(jobs)} shown):"]
            for j in jobs:
                due = j.get("due_date", "no due date")
                client = j.get("client_name", j.get("client", "Unknown"))
                name = j.get("name", j.get("job_name", "Unnamed job"))
                status = j.get("status", "")
                lines.append(f"  • {client} — {name} (due: {due}, status: {status})")
            return "\n".join(lines)
        except Exception as e:
            return f"XPM data unavailable: {e}"

    def _gather_inbox_summary(self, context: PluginContext) -> str:
        """Summarise unread high-priority emails from the inbox."""
        if not context.graph:
            return "Inbox not connected."
        try:
            messages = context.graph.get_unread_emails(max_results=20)
            if not messages:
                return "Inbox: No unread emails."
            lines = [f"Unread Emails ({len(messages)}):"]
            for m in messages[:10]:
                sender = m.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
                subject = m.get("subject", "(no subject)")[:80]
                lines.append(f"  • From: {sender} — {subject}")
            if len(messages) > 10:
                lines.append(f"  ... and {len(messages) - 10} more.")
            return "\n".join(lines)
        except Exception as e:
            return f"Inbox summary unavailable: {e}"

    def _gather_memory_context(self, context: PluginContext, today: str) -> str:
        """Retrieve relevant recent interactions from vector memory."""
        if not context.memory:
            return ""
        try:
            results = context.memory.search(
                query=f"client interaction {today}",
                n_results=5,
                collection="client_interactions")
            if not results:
                return ""
            lines = ["Recent Client Activity (from memory):"]
            for r in results:
                meta = r.get("metadata", {})
                client = meta.get("client_name", meta.get("client", "Unknown"))
                doc = str(r.get("document", ""))[:120]
                lines.append(f"  • {client}: {doc}")
            return "\n".join(lines)
        except Exception:
            return ""

    def _synthesise_briefing(self, context: PluginContext, today: str,
                              xpm_section: str, inbox_section: str,
                              memory_section: str) -> str:
        """Use Claude Sonnet to write a concise, actionable morning briefing."""
        if not context.claude_reason:
            # Fallback: plain text assembly
            parts = [f"Morning Briefing — {today}", "=" * 50]
            if xpm_section:
                parts += ["", "JOBS & DEADLINES", xpm_section]
            if inbox_section:
                parts += ["", "INBOX", inbox_section]
            if memory_section:
                parts += ["", "RECENT CLIENT ACTIVITY", memory_section]
            return "\n".join(parts)

        practice_name = get_setting("practice_name", "MC & S Accounting")
        lessons = get_active_lessons()
        lessons_section = ""
        if lessons:
            lessons_section = "\n--- LEARNED PREFERENCES ---\n" + "\n".join(f"- {l['lesson']}" for l in lessons) + "\n"
        prompt = f"""You are the AI assistant for {practice_name}.
Today is {today}. Compile a concise, professional morning briefing for the team.

Use the data below to write a briefing that:
1. Opens with a one-sentence overview of the day
2. Highlights the top 3-5 priorities (jobs due today or overdue)
3. Notes any urgent inbox items
4. Mentions any relevant recent client activity
5. Closes with a brief motivational line

Keep it under 400 words. Use plain text (no markdown).

--- XPM DATA ---
{xpm_section}

--- INBOX ---
{inbox_section}

--- RECENT CLIENT ACTIVITY ---
{memory_section}
{lessons_section}"""
        try:
            response = context.claude_reason.messages.create(
                model=self.get_claude_model_reasoning(),
                max_tokens=600,
                messages=[{"role": "user", "content": prompt}])
            return response.content[0].text.strip()
        except Exception as e:
            return (f"Morning Briefing — {today}\n{'=' * 50}\n"
                    f"{xpm_section}\n\n{inbox_section}\n\n{memory_section}\n"
                    f"\n[AI synthesis unavailable: {e}]")

    # ── Utility ───────────────────────────────────────────────────────────────

    def _is_business_day(self, dt: datetime) -> bool:
        """Return True if dt falls on a configured business day."""
        business_days_str = get_setting("business_days", "1,2,3,4,5")
        try:
            business_days = [int(d) for d in business_days_str.split(",") if d.strip()]
        except ValueError:
            business_days = [1, 2, 3, 4, 5]
        # weekday(): Monday=0 … Sunday=6; our setting: Mon=1 … Sun=7
        return (dt.weekday() + 1) in business_days
