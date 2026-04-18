"""
MC & S Plugin: Morning Briefing
================================
Plugin ID  : plugin_morning_briefing
Version    : 2.0.0

WHAT IT DOES
------------
Runs once each business morning (default 8:00 AM) and compiles a
personalised daily briefing delivered by email.

Content varies by installation mode (reception_mode setting):

RECEPTION MODE (reception_mode = "1")
  - Unsigned FuseSign documents
  - ASIC returns overdue / awaiting solvency
  - Debtor summary (outstanding invoices, 90+ days)
  - NOA processing queue (unprocessed ATO emails)
  - Unread inbox summary
  - Pending approvals

ACCOUNTANT MODE (reception_mode = "0")
  - XPM jobs due today / overdue WIP
  - Pending approvals
  - Unread inbox summary
  - Compliance deadlines in the next 7 days
  - On Mondays: Weekly tax & accounting industry digest
    (ATO announcements, draft legislation, tax cases, super changes)

SCHEDULE
--------
Default: once daily at 8:00 AM (business days only).
"""

import re
import time
import requests
from datetime import datetime, date, timedelta
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, log_activity, get_active_lessons


# ── Industry digest sources ───────────────────────────────────────────────────
# Each entry: (label, url, css_hint) — css_hint is a rough text marker to
# help extract the relevant section if the page is large.
DIGEST_SOURCES = [
    ("ATO What's New",
     "https://www.ato.gov.au/about-ato/news-and-updates/",
     "news"),
    ("ATO Tax Professionals Newsroom",
     "https://www.ato.gov.au/tax-and-super-professionals/for-tax-professionals/newsroom/",
     "newsroom"),
    ("Treasury Consultations & Exposure Drafts",
     "https://treasury.gov.au/consultation",
     "consultation"),
    ("Tax Practitioners Board News",
     "https://www.tpb.gov.au/news",
     "news"),
    ("AAT Tax Decisions",
     "https://www.aat.gov.au/AAT/media/AAT/Files/Tax%20decisions/Tax-decisions.pdf",
     ""),
]

_DIGEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (compatible; MCSCoWorker/2.0; +https://mcands.com.au)"
    )
}


class MorningBriefingPlugin(AgentPlugin):
    """Compiles and delivers a daily morning briefing to the team."""

    PLUGIN_ID   = "plugin_morning_briefing"
    NAME        = "Morning Briefing"
    DESCRIPTION = ("Compiles a daily morning briefing by email. Reception mode covers "
                   "ASIC/NOA/debtors/FuseSign. Accountant mode covers XPM jobs and "
                   "compliance deadlines. Mondays add a weekly tax industry digest.")
    VERSION     = "2.0.0"
    ICON        = "🌅"
    SCHEDULE    = Schedule.daily_at(8)

    DEFAULT_SETTINGS = {
        "briefing_hour":         "8",
        "briefing_minute":       "0",
        "send_email":            "1",
        "email_recipients":      "",      # comma-separated; falls back to user_email
        "max_jobs_shown":        "15",
        "include_inbox_summary": "1",
        "include_memory_context":"1",
        # Quiet-day suppression
        "suppress_if_quiet":     "1",
        "quiet_threshold":       "3",
        # Monday industry digest
        "monday_digest":         "1",     # 1 = include weekly digest on Mondays
    }

    def __init__(self):
        super().__init__()
        self._last_run_date: str = ""

    # ── Main entry point ──────────────────────────────────────────────────────

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()
        now = datetime.now()
        target_hour = int(self.get_plugin_setting("briefing_hour", "8"))

        if now.hour != target_hour:
            result.summary = f"Not briefing time yet (target {target_hour:02d}:00)."
            return result

        today_str = date.today().isoformat()
        if self._last_run_date == today_str:
            result.summary = "Morning briefing already sent today."
            return result

        if not self._is_business_day(now):
            result.summary = "Not a business day — skipping morning briefing."
            return result

        context.log(f"[MorningBriefing] Compiling briefing for {today_str}...")

        reception_mode = get_setting("reception_mode", "0") == "1"
        is_monday = now.weekday() == 0  # Monday = 0

        # ── Gather data sections ──────────────────────────────────────────────
        if reception_mode:
            briefing_text, actionable = self._compile_reception_briefing(
                context, today_str)
        else:
            briefing_text, actionable = self._compile_accountant_briefing(
                context, today_str, is_monday)

        # ── Quiet-day suppression ─────────────────────────────────────────────
        if self.get_plugin_setting("suppress_if_quiet", "1") == "1":
            quiet_threshold = int(self.get_plugin_setting("quiet_threshold", "3"))
            if actionable < quiet_threshold:
                self._last_run_date = today_str
                result.summary = (
                    f"Morning briefing suppressed — only {actionable} actionable "
                    f"items (threshold: {quiet_threshold}). Quiet day."
                )
                context.log(
                    f"[MorningBriefing] Quiet day: {actionable} items < {quiet_threshold}."
                )
                return result

        # ── Deliver by email ──────────────────────────────────────────────────
        delivered = []
        recipients_setting = self.get_plugin_setting("email_recipients", "")
        if not recipients_setting:
            recipients_setting = get_setting("user_email", "")

        if self.get_plugin_setting("send_email", "1") == "1" and context.graph:
            if is_monday and not reception_mode:
                subject = f"Weekly Tax & Accounting Digest — week ending {today_str}"
            elif reception_mode:
                subject = f"Reception Briefing — {today_str}"
            else:
                subject = f"Morning Briefing — {today_str}"

            for email in [r.strip() for r in recipients_setting.split(",") if r.strip()]:
                try:
                    context.graph.send_email(
                        to=email,
                        subject=subject,
                        body=f"<pre style='font-family:Calibri,sans-serif;font-size:14px'>"
                             f"{briefing_text}</pre>",
                        draft_mode=context.draft_mode,
                    )
                    delivered.append(email)
                    context.log(f"[MorningBriefing] Sent to {email}.")
                except Exception as e:
                    context.log(f"[MorningBriefing] Email to {email} failed: {e}")

        # ── Store in memory ───────────────────────────────────────────────────
        if context.memory:
            try:
                context.memory.store(
                    document=briefing_text,
                    metadata={"type": "morning_briefing", "date": today_str,
                              "mode": "reception" if reception_mode else "accountant",
                              "delivered_to": ", ".join(delivered)},
                    collection="general",
                    doc_id=f"briefing_{today_str}",
                )
            except Exception:
                pass

        if context.event_bus:
            context.event_bus.publish(
                "morning_briefing.sent",
                source=self.PLUGIN_ID,
                payload={"date": today_str, "delivered_to": delivered},
            )

        self._last_run_date = today_str
        result.summary = (
            f"Morning briefing sent for {today_str}. "
            f"Delivered to: {', '.join(delivered) or 'none configured'}."
        )
        result.actions_taken = 1
        log_activity(
            from_email="system", subject="Morning Briefing",
            classification="morning_briefing", action="briefing_sent",
            draft_created=context.draft_mode,
        )
        return result

    # ── Reception briefing ────────────────────────────────────────────────────

    def _compile_reception_briefing(self, context: PluginContext,
                                    today: str) -> tuple[str, int]:
        """Compile a reception-mode briefing. Returns (text, actionable_count)."""
        sections = []
        actionable = 0

        # 1. FuseSign unsigned documents
        fusesign_section, fs_count = self._gather_fusesign(context)
        if fusesign_section:
            sections.append(fusesign_section)
            actionable += fs_count

        # 2. ASIC overdue / awaiting solvency
        asic_section, asic_count = self._gather_asic_status()
        if asic_section:
            sections.append(asic_section)
            actionable += asic_count

        # 3. Debtor summary
        debtor_section, debtor_count = self._gather_debtor_summary(context)
        if debtor_section:
            sections.append(debtor_section)
            actionable += debtor_count

        # 4. NOA queue
        noa_section, noa_count = self._gather_noa_queue()
        if noa_section:
            sections.append(noa_section)
            actionable += noa_count

        # 5. Inbox summary
        if self.get_plugin_setting("include_inbox_summary", "1") == "1":
            inbox_section = self._gather_inbox_summary(context)
            if inbox_section:
                sections.append(inbox_section)

        # 6. Pending approvals
        approvals_section, appr_count = self._gather_pending_approvals(context)
        if approvals_section:
            sections.append(approvals_section)
            actionable += appr_count

        if not sections:
            return f"Reception Briefing — {today}\n{'=' * 50}\nAll clear. Nothing requires attention today.", 0

        raw = "\n\n".join(sections)
        briefing_text = self._synthesise(
            context, today, raw,
            mode_instruction=(
                "You are compiling a morning briefing for the reception team at an "
                "Australian accounting firm. Focus on what needs to be actioned today: "
                "unsigned documents, overdue debtors, ASIC reminders, unprocessed NOAs, "
                "and inbox items. Be direct and specific. No motivational language. "
                "Use plain text, no markdown. Keep under 350 words."
            )
        )
        return briefing_text, actionable

    # ── Accountant briefing ───────────────────────────────────────────────────

    def _compile_accountant_briefing(self, context: PluginContext,
                                     today: str, is_monday: bool) -> tuple[str, int]:
        """Compile an accountant-mode briefing. Returns (text, actionable_count)."""
        sections = []
        actionable = 0

        # 1. XPM jobs
        xpm_section = self._gather_xpm_data(context)
        if xpm_section:
            sections.append(xpm_section)
            actionable += xpm_section.lower().count("overdue")
            actionable += xpm_section.lower().count("due today")

        # 2. Pending approvals
        approvals_section, appr_count = self._gather_pending_approvals(context)
        if approvals_section:
            sections.append(approvals_section)
            actionable += appr_count

        # 3. Inbox summary
        if self.get_plugin_setting("include_inbox_summary", "1") == "1":
            inbox_section = self._gather_inbox_summary(context)
            if inbox_section:
                sections.append(inbox_section)

        # 4. Memory context
        if self.get_plugin_setting("include_memory_context", "1") == "1":
            memory_section = self._gather_memory_context(context, today)
            if memory_section:
                sections.append(memory_section)

        # 5. Monday industry digest
        digest_section = ""
        if is_monday and self.get_plugin_setting("monday_digest", "1") == "1":
            digest_section = self._gather_industry_digest(context)

        raw = "\n\n".join(sections)
        briefing_text = self._synthesise(
            context, today, raw,
            mode_instruction=(
                "You are compiling a morning briefing for an accountant at an "
                "Australian accounting firm. Highlight the top priorities: jobs due "
                "today or overdue, items awaiting approval, and any urgent inbox items. "
                "Flag any compliance deadlines in the next 7 days. "
                "Be direct and professional. No motivational language. "
                "Use plain text, no markdown. Keep under 400 words."
            ),
            digest_section=digest_section,
        )
        return briefing_text, actionable

    # ── Data gathering helpers ────────────────────────────────────────────────

    def _gather_fusesign(self, context: PluginContext) -> tuple[str, int]:
        """Return FuseSign unsigned document summary and count."""
        try:
            if not context.gateway or not context.gateway.is_available("fusesign"):
                return "", 0
            envelopes = context.gateway.fusesign.list_pending(limit=50)
            if not envelopes:
                return "", 0
            lines = [f"FuseSign — Unsigned Documents ({len(envelopes)}):"]
            for e in envelopes[:10]:
                client = e.get("recipient_name", e.get("client_name", "Unknown"))
                doc = e.get("document_name", e.get("title", "Unnamed"))
                sent = e.get("sent_date", e.get("created_at", ""))[:10]
                lines.append(f"  • {client} — {doc} (sent {sent})")
            if len(envelopes) > 10:
                lines.append(f"  ... and {len(envelopes) - 10} more.")
            return "\n".join(lines), len(envelopes)
        except Exception as e:
            return f"FuseSign data unavailable: {e}", 0

    def _gather_asic_status(self) -> tuple[str, int]:
        """Return ASIC returns summary and count of items needing action."""
        try:
            from plugins.plugin_asic_returns import get_asic_returns
            all_returns = get_asic_returns(limit=200)
            open_returns = [r for r in all_returns
                            if r.get("status") not in ("completed", "cancelled")]
            if not open_returns:
                return "", 0
            overdue = [r for r in open_returns if r.get("status") == "overdue"]
            awaiting = [r for r in open_returns if r.get("status") == "awaiting_solvency"]
            pending = [r for r in open_returns
                       if r.get("status") not in ("overdue", "awaiting_solvency")]
            lines = [f"ASIC Annual Returns — {len(open_returns)} open:"]
            if overdue:
                lines.append(f"  OVERDUE ({len(overdue)}): "
                             + ", ".join(r.get("company_name", "?") for r in overdue[:5]))
            if awaiting:
                lines.append(f"  Awaiting solvency ({len(awaiting)}): "
                             + ", ".join(r.get("company_name", "?") for r in awaiting[:5]))
            if pending:
                lines.append(f"  Pending ({len(pending)}): "
                             + ", ".join(r.get("company_name", "?") for r in pending[:5]))
            return "\n".join(lines), len(overdue) + len(awaiting)
        except Exception:
            return "", 0

    def _gather_debtor_summary(self, context: PluginContext) -> tuple[str, int]:
        """Return debtor summary from XPM invoices and count of overdue clients."""
        try:
            if not context.gateway or not context.gateway.is_available("xpm"):
                return "", 0
            summary = context.gateway.xpm.get_debtor_summary()
            total = summary.get("total_outstanding", 0)
            count = summary.get("invoice_count", 0)
            over90 = summary.get("90_plus", 0)
            over60 = summary.get("61_90", 0)
            if count == 0:
                return "", 0
            lines = [
                f"Debtors — {count} outstanding invoices, total ${total:,.0f}:",
                f"  90+ days overdue: ${over90:,.0f}",
                f"  61–90 days: ${over60:,.0f}",
            ]
            # List top overdue clients
            top = summary.get("top_overdue", [])
            for client in top[:5]:
                name = client.get("client_name", "Unknown")
                amount = client.get("total_outstanding", 0)
                days = client.get("max_days_overdue", 0)
                lines.append(f"  • {name}: ${amount:,.0f} ({days} days)")
            actionable = len([c for c in top if c.get("max_days_overdue", 0) > 30])
            return "\n".join(lines), actionable
        except Exception as e:
            return f"Debtor data unavailable: {e}", 0

    def _gather_noa_queue(self) -> tuple[str, int]:
        """Return count of unprocessed NOA emails from the NOA processor DB."""
        try:
            from plugins.plugin_noa_processor import get_noa_queue_count
            count = get_noa_queue_count()
            if count == 0:
                return "", 0
            return f"NOA Queue — {count} unprocessed ATO notice(s) awaiting client email.", count
        except Exception:
            return "", 0

    def _gather_pending_approvals(self, context: PluginContext) -> tuple[str, int]:
        """Return pending approval queue summary and count."""
        try:
            if not context.approval_queue:
                return "", 0
            count = context.approval_queue.count_pending()
            if count == 0:
                return "", 0
            items = context.approval_queue.list_pending(limit=5)
            lines = [f"Approval Queue — {count} item(s) awaiting review:"]
            for item in items:
                desc = item.get("description", item.get("action_type", "?"))[:70]
                plugin = item.get("plugin_id", "?")
                lines.append(f"  • [{plugin}] {desc}")
            if count > 5:
                lines.append(f"  ... and {count - 5} more.")
            return "\n".join(lines), count
        except Exception:
            return "", 0

    def _gather_xpm_data(self, context: PluginContext) -> str:
        """Pull today's jobs and overdue WIP from XPM."""
        if not context.gateway or not context.gateway.is_available("xpm"):
            return "XPM not connected — job data unavailable."
        try:
            max_jobs = int(self.get_plugin_setting("max_jobs_shown", "15"))
            jobs = context.gateway.xpm.list_jobs(status="inprogress", limit=max_jobs)
            if not jobs:
                return "No active jobs found in XPM."
            today = date.today()
            overdue, due_today, upcoming = [], [], []
            for j in jobs:
                due_str = j.get("due_date", "")
                client = j.get("client_name", j.get("client", "Unknown"))
                name = j.get("name", j.get("job_name", "Unnamed job"))
                try:
                    due_date = date.fromisoformat(due_str[:10]) if due_str else None
                except ValueError:
                    due_date = None
                entry = f"  • {client} — {name} (due: {due_str or 'no date'})"
                if due_date:
                    if due_date < today:
                        overdue.append(entry)
                    elif due_date == today:
                        due_today.append(entry)
                    elif due_date <= today + timedelta(days=7):
                        upcoming.append(entry)
            lines = [f"XPM Jobs ({len(jobs)} active):"]
            if overdue:
                lines.append(f"  OVERDUE ({len(overdue)}):")
                lines.extend(overdue[:5])
            if due_today:
                lines.append(f"  Due today ({len(due_today)}):")
                lines.extend(due_today)
            if upcoming:
                lines.append(f"  Due this week ({len(upcoming)}):")
                lines.extend(upcoming[:5])
            return "\n".join(lines)
        except Exception as e:
            return f"XPM data unavailable: {e}"

    def _gather_inbox_summary(self, context: PluginContext) -> str:
        """Summarise unread emails from the inbox."""
        if not context.graph:
            return ""
        try:
            messages = context.graph.get_unread_emails(max_results=20)
            if not messages:
                return ""
            lines = [f"Unread Emails ({len(messages)}):"]
            for m in messages[:10]:
                sender = m.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
                subject = m.get("subject", "(no subject)")[:80]
                lines.append(f"  • {sender} — {subject}")
            if len(messages) > 10:
                lines.append(f"  ... and {len(messages) - 10} more.")
            return "\n".join(lines)
        except Exception:
            return ""

    def _gather_memory_context(self, context: PluginContext, today: str) -> str:
        """Retrieve recent client interactions from vector memory."""
        if not context.memory:
            return ""
        try:
            results = context.memory.search(
                query=f"client interaction {today}",
                n_results=5,
                collection="client_interactions",
            )
            if not results:
                return ""
            lines = ["Recent Client Activity (from memory):"]
            for r in results:
                meta = r.get("metadata", {})
                client = meta.get("client_name", meta.get("client", "Unknown"))
                note = r.get("document", "")[:100]
                lines.append(f"  • {client}: {note}")
            return "\n".join(lines)
        except Exception:
            return ""

    def _gather_industry_digest(self, context: PluginContext) -> str:
        """
        Fetch the past 7 days of Australian tax/super/legislation news
        from ATO, Treasury, TPB, and AAT. Returns a raw text block for
        Claude to summarise as the Monday weekly digest.
        """
        context.log("[MorningBriefing] Fetching Monday industry digest sources...")
        raw_chunks = []
        cutoff = datetime.now() - timedelta(days=7)

        for label, url, _ in DIGEST_SOURCES:
            try:
                resp = requests.get(url, headers=_DIGEST_HEADERS, timeout=15)
                if resp.status_code == 200:
                    # Strip HTML tags to plain text
                    text = re.sub(r"<[^>]+>", " ", resp.text)
                    text = re.sub(r"\s+", " ", text).strip()
                    # Take first 3000 chars — Claude will summarise
                    raw_chunks.append(f"--- {label} ---\n{text[:3000]}")
                else:
                    raw_chunks.append(f"--- {label} ---\n(Unavailable: HTTP {resp.status_code})")
            except Exception as e:
                raw_chunks.append(f"--- {label} ---\n(Fetch failed: {e})")

        if not any("Unavailable" not in c and "failed" not in c for c in raw_chunks):
            return ""

        return "\n\n".join(raw_chunks)

    # ── Claude synthesis ──────────────────────────────────────────────────────

    def _synthesise(self, context: PluginContext, today: str,
                    raw_data: str, mode_instruction: str,
                    digest_section: str = "") -> str:
        """Use Claude to write the final briefing from raw data sections."""
        practice_name = get_setting("practice_name", "MC & S Accounting")
        staff_profile = get_setting("staff_profile", "accountant")
        lessons = get_active_lessons()
        lessons_text = ""
        if lessons:
            lessons_text = "\n--- LEARNED PREFERENCES ---\n" + \
                           "\n".join(f"- {l['lesson']}" for l in lessons[:10])

        digest_instruction = ""
        if digest_section:
            digest_instruction = (
                "\n\nAfter the daily briefing, add a clearly separated section titled "
                "'WEEKLY TAX & ACCOUNTING DIGEST — WEEK ENDING " + today + "'. "
                "Summarise the most important developments from the sources below "
                "that are relevant to Australian tax practitioners: ATO announcements, "
                "new or amended legislation, significant tax cases, and super changes. "
                "Be specific — name the ruling, case, or announcement. Ignore "
                "irrelevant content. Keep this section under 300 words."
            )

        prompt = f"""You are the AI assistant for {practice_name}.
Today is {today}. Staff profile: {staff_profile}.

{mode_instruction}{digest_instruction}

--- DATA ---
{raw_data}
{lessons_text}
{f"--- INDUSTRY DIGEST SOURCES ---{chr(10)}{digest_section}" if digest_section else ""}"""

        if not context.claude_reason:
            # Fallback: plain text assembly
            parts = [f"Morning Briefing — {today}", "=" * 50, raw_data]
            if digest_section:
                parts += ["", f"WEEKLY TAX & ACCOUNTING DIGEST — WEEK ENDING {today}",
                          "(AI unavailable — raw source data below)", digest_section[:1000]]
            return "\n".join(parts)

        try:
            response = context.claude_reason.messages.create(
                model=self.get_claude_model_reasoning(),
                max_tokens=900,
                messages=[{"role": "user", "content": prompt}],
            )
            return response.content[0].text.strip()
        except Exception as e:
            return (
                f"Morning Briefing — {today}\n{'=' * 50}\n{raw_data}"
                f"\n\n[AI synthesis unavailable: {e}]"
            )

    # ── Utility ───────────────────────────────────────────────────────────────

    def _is_business_day(self, dt: datetime) -> bool:
        """Return True if dt falls on a configured business day."""
        business_days_str = get_setting("business_days", "1,2,3,4,5")
        try:
            business_days = [int(d) for d in business_days_str.split(",") if d.strip()]
        except ValueError:
            business_days = [1, 2, 3, 4, 5]
        return (dt.weekday() + 1) in business_days
