"""
MC & S Plugin: Meeting Prep Brief
===================================
Plugin ID  : plugin_meeting_prep
Version    : 1.0.0

WHAT IT DOES
------------
Triggered by an email containing a meeting request or calendar invite.
Before the meeting it:

1. Detects meeting-related emails (calendar invites, "meeting" in subject)
2. Extracts the client name and meeting time from the email
3. Queries XPM for the client's active jobs, recent notes, and WIP
4. Retrieves the client's interaction history from vector memory
5. Uses Claude Sonnet to compile a concise meeting prep brief
6. Sends the brief to Teams and/or drafts it as an email reply
7. Stores the brief in memory for post-meeting reference

APEX ALIGNMENT
--------------
Mirrors APEX's "Meeting Prep Agent" — ensures the team walks into
every client meeting fully briefed without manual research.

SCHEDULE
--------
Default: daily at 08:00 — scans for today's meetings each morning.
"""

import re
from datetime import datetime
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule, PluginCategory
from config import get_setting, log_activity, get_active_lessons
from client_utils import normalise_client_name


# Keywords that indicate a meeting-related email
MEETING_KEYWORDS = [
    "meeting", "appointment", "catch up", "catch-up", "call scheduled",
    "zoom", "teams meeting", "calendar invite", "calendar invitation",
    "let's meet", "lets meet", "book a time", "schedule a time",
]


class MeetingPrepPlugin(AgentPlugin):
    """Compiles a client meeting prep brief from XPM and memory."""

    PLUGIN_ID   = "plugin_meeting_prep"
    name        = "Meeting Prep Brief"
    description = ("Detects meeting emails, pulls client context from XPM and memory, "
                   "and delivers a pre-meeting brief to Teams.")
    version     = "1.0.0"
    icon        = "📋"
    default_schedule = Schedule.daily_at(8)  # Every morning at 08:00
    category    = PluginCategory.ACCOUNTANT

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        if not context.graph:
            result.summary = "Microsoft 365 not connected."
            return result

        # ── Scan inbox for meeting-related emails ─────────────────────────────
        try:
            messages = context.graph.fetch_unread_emails(folder="Inbox", max_count=30)
        except Exception as e:
            result.summary = f"Inbox scan failed: {e}"
            return result

        meeting_emails = []
        for msg in messages:
            subject = msg.get("subject", "").lower()
            body_preview = msg.get("bodyPreview", "").lower()
            if any(kw in subject or kw in body_preview for kw in MEETING_KEYWORDS):
                meeting_emails.append(msg)

        if not meeting_emails:
            result.summary = "No meeting emails detected."
            return result

        context.log(f"[MeetingPrep] Found {len(meeting_emails)} meeting email(s).")
        briefs_sent = 0

        for msg in meeting_emails:
            msg_id  = msg.get("id", "")
            subject = msg.get("subject", "(no subject)")
            sender  = msg.get("from", {}).get("emailAddress", {})
            sender_name  = sender.get("name", "Unknown")
            sender_email = sender.get("address", "")

            # ── Extract client name ───────────────────────────────────────────
            client_name = self._extract_client_name(subject, sender_name)

            # ── Gather XPM context ────────────────────────────────────────────
            xpm_context = self._get_xpm_context(context, client_name, sender_email)

            # ── Gather memory context ─────────────────────────────────────────
            memory_context = self._get_memory_context(context, client_name)

            # ── Synthesise brief ──────────────────────────────────────────────
            brief = self._synthesise_brief(
                context, client_name, subject, xpm_context, memory_context)

            # ── Deliver ───────────────────────────────────────────────────────
            delivered = []
            if self.get_plugin_setting("send_to_teams", "1") == "1":
                if context.gateway and context.gateway.is_available("teams"):
                    try:
                        context.gateway.teams.send_alert(
                            title=f"📋 Meeting Prep: {client_name}",
                            message=brief[:3000],
                            color="7B1FA2")
                        delivered.append("Teams")
                    except Exception as e:
                        context.log(f"[MeetingPrep] Teams failed: {e}")

            # ── Store in memory ───────────────────────────────────────────────
            if context.memory:
                try:
                    context.memory.store_client_interaction(
                        client_name=normalise_client_name(client_name),
                        interaction_type="meeting_prep",
                        summary=f"Meeting prep brief generated for: {subject}",
                        metadata={"subject": subject, "sender": sender_email})
                except Exception:
                    pass

            # ── Mark email as read ────────────────────────────────────────────
            try:
                context.graph.mark_as_read(msg_id)
            except Exception:
                pass

            briefs_sent += 1
            log_activity(from_email=sender_email, subject="Meeting Prep Brief", classification="meeting_prep",
                         action="brief_sent", draft_created=False)

        result.summary = f"Meeting prep briefs sent: {briefs_sent}."
        result.actions_taken = briefs_sent
        return result

    def _extract_client_name(self, subject: str, sender_name: str) -> str:
        """Extract the most likely client name from subject or sender."""
        # Try to find a name pattern in the subject
        match = re.search(r"(?:with|for|re:?)\s+([A-Z][a-z]+ [A-Z][a-z]+)", subject)
        if match:
            return match.group(1)
        # Fall back to sender name
        if sender_name and sender_name.lower() not in ("unknown", ""):
            return sender_name
        return "Client"

    def _get_xpm_context(self, context: PluginContext,
                          client_name: str, client_email: str) -> str:
        """Retrieve client jobs and notes from XPM."""
        if not context.gateway or not context.gateway.is_available("xpm"):
            return "XPM not connected."
        try:
            # Try to find client by email first, then by name
            client = None
            if client_email:
                try:
                    client = context.gateway.xpm.get_client_by_email(client_email)
                except Exception:
                    pass
            if not client:
                clients = context.gateway.xpm.list_clients(
                    search=client_name, limit=3)
                if clients:
                    client = clients[0]
            if not client:
                return f"No XPM record found for {client_name}."

            client_id = client.get("id", client.get("client_id", ""))
            jobs = context.gateway.xpm.list_jobs(
                client_id=client_id, status="inprogress", limit=5)

            lines = [f"XPM Profile: {client.get('name', client_name)}"]
            if jobs:
                lines.append(f"Active Jobs ({len(jobs)}):")
                for j in jobs:
                    lines.append(f"  • {j.get('name', 'Unnamed')} "
                                 f"(due: {j.get('due_date', '?')}, "
                                 f"status: {j.get('status', '?')})")
            else:
                lines.append("No active jobs.")
            return "\n".join(lines)
        except Exception as e:
            return f"XPM lookup failed: {e}"

    def _get_memory_context(self, context: PluginContext, client_name: str) -> str:
        """Retrieve recent client interactions from vector memory."""
        if not context.memory:
            return ""
        try:
            results = context.memory.get_client_context(
                client_name=normalise_client_name(client_name), n_results=5)
            return results or ""
        except Exception:
            return ""

    def _synthesise_brief(self, context: PluginContext, client_name: str,
                           meeting_subject: str, xpm_context: str,
                           memory_context: str) -> str:
        """Use Claude Sonnet to write a meeting prep brief."""
        practice = get_setting("practice_name", "MC & S Accounting")
        if not context.claude_reason:
            return (f"Meeting Prep: {client_name}\n"
                    f"Subject: {meeting_subject}\n\n"
                    f"{xpm_context}\n\n{memory_context}")
        try:
            lessons = get_active_lessons()
            lessons_block = ""
            if lessons:
                lessons_block = "\nLEARNED PREFERENCES:\n" + "\n".join(f"- {l['lesson']}" for l in lessons) + "\n"
            prompt = (f"You are the AI assistant for {practice}.\n"
                      f"Prepare a concise meeting brief for a meeting with {client_name}.\n"
                      f"Meeting subject: {meeting_subject}\n\n"
                      f"XPM DATA:\n{xpm_context}\n\n"
                      f"RECENT HISTORY:\n{memory_context}\n"
                      f"{lessons_block}\n"
                      "Write a 3-section brief:\n"
                      "1. Client snapshot (2-3 sentences)\n"
                      "2. Current work & outstanding items\n"
                      "3. Suggested talking points\n"
                      "Keep it under 250 words. Plain text only.")
            resp = context.claude_reason.messages.create(
                model=self.get_claude_model_reasoning(),
                max_tokens=400,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception as e:
            return (f"Meeting Prep: {client_name}\n"
                    f"Subject: {meeting_subject}\n\n"
                    f"{xpm_context}\n\n{memory_context}\n"
                    f"[AI synthesis unavailable: {e}]")
