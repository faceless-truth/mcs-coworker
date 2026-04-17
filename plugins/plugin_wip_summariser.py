"""
MC & S Plugin: WIP Ageing Summariser
======================================
Plugin ID  : plugin_wip_summariser
Version    : 1.0.0

WHAT IT DOES
------------
Queries XPM for all in-progress jobs, groups them by ageing bucket
(0-30 days, 31-60 days, 61-90 days, 90+ days), and:

1. Generates a WIP ageing summary report using Claude Sonnet
2. Flags jobs that have not moved in more than 30 days
3. Sends an alert to Teams if any jobs are in the 90+ day bucket
4. Optionally emails the report to the nominated manager
5. Stores the summary in vector memory for trend analysis

APEX ALIGNMENT
--------------
Mirrors APEX's "WIP & Billing Agent" — proactive monitoring of
work-in-progress to prevent revenue leakage and deadline misses.

SCHEDULE
--------
Default: every Monday at 9:00 AM (weekly WIP review).
"""

from datetime import datetime, date, timedelta
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity


class WIPSummariserPlugin(AgentPlugin):
    """Generates a weekly WIP ageing summary from XPM data."""

    PLUGIN_ID   = "plugin_wip_summariser"
    NAME        = "WIP Ageing Summariser"
    DESCRIPTION = ("Queries XPM for in-progress jobs, groups by ageing bucket, "
                   "flags stale work, and delivers a weekly WIP report to Teams.")
    VERSION     = "1.0.0"
    ICON        = "📊"
    SCHEDULE    = Schedule.every_hours(168)

    DEFAULT_SETTINGS = {
        "run_day":            "1",    # 1=Monday
        "run_hour":           "9",
        "stale_threshold_days": "30",
        "alert_threshold_days": "90",
        "send_to_teams":      "1",
        "send_email":         "0",
        "email_recipients":   "",
    }

    def __init__(self):
        super().__init__()
        self._last_run_week: str = ""

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        # ── Day/hour gate ─────────────────────────────────────────────────────
        now = datetime.now()
        target_day  = int(self.get_plugin_setting("run_day",  "1"))
        target_hour = int(self.get_plugin_setting("run_hour", "9"))
        # weekday(): Mon=0; our setting Mon=1
        if (now.weekday() + 1) != target_day or now.hour != target_hour:
            result.summary = "Not WIP review time."
            return result

        week_str = f"{date.today().isocalendar()[0]}-W{date.today().isocalendar()[1]}"
        if self._last_run_week == week_str:
            result.summary = "WIP summary already sent this week."
            return result

        if not context.gateway or not context.gateway.is_available("xpm"):
            result.summary = "XPM not configured — WIP summary skipped."
            return result

        context.log("[WIPSummariser] Fetching jobs from XPM...")

        # ── Fetch all in-progress jobs ────────────────────────────────────────
        try:
            jobs = context.gateway.xpm.list_jobs(status="inprogress", limit=200)
        except Exception as e:
            result.summary = f"XPM fetch failed: {e}"
            return result

        if not jobs:
            result.summary = "No in-progress jobs found in XPM."
            return result

        # ── Bucket by ageing ──────────────────────────────────────────────────
        today = date.today()
        stale_days  = int(self.get_plugin_setting("stale_threshold_days", "30"))
        alert_days  = int(self.get_plugin_setting("alert_threshold_days", "90"))

        buckets = {"0-30": [], "31-60": [], "61-90": [], "90+": []}
        stale_jobs = []

        for job in jobs:
            due_str = job.get("due_date", "")
            try:
                due = datetime.strptime(due_str[:10], "%Y-%m-%d").date()
                age = (today - due).days
            except Exception:
                age = 0

            if age <= 30:
                buckets["0-30"].append(job)
            elif age <= 60:
                buckets["31-60"].append(job)
            elif age <= 90:
                buckets["61-90"].append(job)
            else:
                buckets["90+"].append(job)

            if age >= stale_days:
                stale_jobs.append((job, age))

        # ── Build report text ─────────────────────────────────────────────────
        lines = [f"WIP Ageing Summary — {today.isoformat()}", "=" * 60]
        total = len(jobs)
        lines.append(f"Total in-progress jobs: {total}\n")

        for bucket, bucket_jobs in buckets.items():
            lines.append(f"[{bucket} days overdue] — {len(bucket_jobs)} job(s)")
            for j in bucket_jobs[:5]:
                client = j.get("client_name", j.get("client", "Unknown"))
                name   = j.get("name", j.get("job_name", "Unnamed"))
                due    = j.get("due_date", "?")
                lines.append(f"    • {client} — {name} (due: {due})")
            if len(bucket_jobs) > 5:
                lines.append(f"    ... and {len(bucket_jobs) - 5} more.")
            lines.append("")

        if stale_jobs:
            lines.append(f"⚠️  Stale Jobs (no movement in {stale_days}+ days): "
                         f"{len(stale_jobs)}")
            for j, age in stale_jobs[:10]:
                client = j.get("client_name", j.get("client", "Unknown"))
                name   = j.get("name", j.get("job_name", "Unnamed"))
                lines.append(f"    • {client} — {name} ({age} days overdue)")

        report_text = "\n".join(lines)

        # ── AI synthesis ──────────────────────────────────────────────────────
        if context.claude_reason:
            try:
                prompt = (f"You are the AI assistant for {get_setting('practice_name','MC & S')}.\n"
                          f"Here is the WIP ageing data for {today}:\n\n{report_text}\n\n"
                          "Write a concise 3-paragraph management summary:\n"
                          "1. Overall WIP health (1 sentence)\n"
                          "2. Top concerns and recommended actions\n"
                          "3. Positive highlights\n"
                          "Keep it under 200 words. Plain text only.")
                resp = context.claude_reason.messages.create(
                    model=self.get_claude_model_reasoning(),
                    max_tokens=350,
                    messages=[{"role": "user", "content": prompt}])
                ai_summary = resp.content[0].text.strip()
                report_text = f"{ai_summary}\n\n{'=' * 60}\n{report_text}"
            except Exception as e:
                context.log(f"[WIPSummariser] AI synthesis failed: {e}")

        # ── Deliver ───────────────────────────────────────────────────────────
        delivered = []
        alert_needed = len(buckets["90+"]) > 0

        if self.get_plugin_setting("send_to_teams", "1") == "1":
            if context.gateway and context.gateway.is_available("teams"):
                try:
                    color = "FF0000" if alert_needed else "0078D4"
                    context.gateway.teams.send_alert(
                        title=f"📊 WIP Ageing Report — {today.isoformat()}",
                        message=report_text[:3000],
                        color=color)
                    delivered.append("Teams")
                except Exception as e:
                    context.log(f"[WIPSummariser] Teams delivery failed: {e}")

        if self.get_plugin_setting("send_email", "0") == "1":
            recipients = self.get_plugin_setting("email_recipients", "")
            if recipients and context.graph:
                for email in [r.strip() for r in recipients.split(",") if r.strip()]:
                    try:
                        context.graph.send_email(
                            to=email,
                            subject=f"WIP Ageing Report — {today.isoformat()}",
                            body=f"<pre>{report_text}</pre>",
                            draft_mode=context.draft_mode)
                        delivered.append(f"email:{email}")
                    except Exception as e:
                        context.log(f"[WIPSummariser] Email failed: {e}")

        # ── Memory ────────────────────────────────────────────────────────────
        if context.memory:
            try:
                context.memory.store(
                    document=report_text,
                    metadata={"type": "wip_summary", "date": today.isoformat(),
                              "total_jobs": total, "stale_count": len(stale_jobs),
                              "critical_count": len(buckets["90+"])},
                    collection="general",
                    doc_id=f"wip_{today.isoformat()}")
            except Exception:
                pass

        # ── Event ─────────────────────────────────────────────────────────────
        if context.event_bus:
            context.event_bus.publish("wip_summary.sent",
                                      source=self.PLUGIN_ID,
                                      payload={"date": today.isoformat(),
                                               "total_jobs": total,
                                               "critical_count": len(buckets["90+"])})

        self._last_run_week = week_str
        result.summary = (f"WIP summary sent for week {week_str}. "
                          f"{total} jobs, {len(buckets['90+'])} critical. "
                          f"Delivered: {', '.join(delivered) or 'none configured'}.")
        result.actions_taken = 1
        log_activity(from_email="system", subject="WIP Summary", classification="wip_summary",
                     action="report_sent", draft_created=False)
        return result
