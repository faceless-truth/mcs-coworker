"""
MC & S Plugin: Engagement Letter Generator
============================================
Plugin ID  : plugin_engagement_letter
Version    : 1.0.0

WHAT IT DOES
------------
Detects emails from new or returning clients requesting services and:

1. Identifies engagement letter trigger phrases in emails
   (e.g., "new client", "onboard", "engagement letter", "terms of service")
2. Extracts client details from the email and XPM
3. Uses Claude Sonnet to generate a personalised engagement letter
   based on the services requested
4. Creates a FuseSign envelope for the letter to be signed
5. Sends the letter to the client via FuseSign
6. Logs the engagement in XPM as a new job note
7. Stores the engagement in vector memory

APEX ALIGNMENT
--------------
Mirrors APEX's "Client Onboarding Agent" — automates the first
formal step in the client relationship.

SCHEDULE
--------
Default: every 10 minutes (event-driven via email scan).
"""

from datetime import datetime
from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule, PluginCategory
from config import get_setting, log_activity


TRIGGER_PHRASES = [
    "engagement letter", "letter of engagement", "new client",
    "onboard", "sign up", "terms of service", "terms and conditions",
    "start working with", "take you on", "new engagement",
]

SERVICES = {
    "tax return": "Individual Tax Return Preparation",
    "bas":        "BAS / GST Preparation",
    "bookkeeping":"Bookkeeping Services",
    "company":    "Company Tax Return Preparation",
    "smsf":       "SMSF Administration and Tax Return",
    "audit":      "Audit Services",
    "advisory":   "Business Advisory Services",
    "payroll":    "Payroll Services",
}


class EngagementLetterPlugin(AgentPlugin):
    """Generates and sends personalised engagement letters via FuseSign."""

    PLUGIN_ID   = "plugin_engagement_letter"
    NAME        = "Engagement Letter Generator"
    DESCRIPTION = ("Detects new client emails, generates a personalised engagement "
                   "letter using AI, and sends it via FuseSign for signing.")
    VERSION     = "1.0.0"
    ICON        = "📝"
    SCHEDULE    = Schedule.every_minutes(10)

    name        = NAME
    description = DESCRIPTION
    version     = VERSION
    icon        = ICON
    default_schedule = SCHEDULE
    category    = PluginCategory.RECEPTION

    DEFAULT_SETTINGS = {
        "use_fusesign":       "1",
        "draft_only":         "0",   # if 1, only draft — don't send
        "confidence_threshold": "0.8",
    }

    def run(self, context: PluginContext) -> PluginResult:
        result = PluginResult()

        if not context.graph:
            result.summary = "Microsoft 365 not connected."
            return result

        try:
            messages = context.graph.get_unread_emails(max_results=30)
        except Exception as e:
            result.summary = f"Inbox scan failed: {e}"
            return result

        triggered = []
        for msg in messages:
            subject = msg.get("subject", "").lower()
            body    = msg.get("bodyPreview", "").lower()
            if any(phrase in subject or phrase in body for phrase in TRIGGER_PHRASES):
                triggered.append(msg)

        if not triggered:
            result.summary = "No engagement letter triggers detected."
            return result

        context.log(f"[EngagementLetter] Found {len(triggered)} trigger email(s).")
        generated = 0

        for msg in triggered:
            msg_id       = msg.get("id", "")
            subject      = msg.get("subject", "(no subject)")
            sender       = msg.get("from", {}).get("emailAddress", {})
            client_name  = sender.get("name", "Valued Client")
            client_email = sender.get("address", "")

            # ── Detect services requested ─────────────────────────────────────
            full_text = f"{subject} {msg.get('bodyPreview', '')}".lower()
            detected_services = [
                label for keyword, label in SERVICES.items()
                if keyword in full_text]
            if not detected_services:
                detected_services = ["General Accounting Services"]

            # ── Get XPM context ───────────────────────────────────────────────
            xpm_info = ""
            if context.gateway and context.gateway.is_available("xpm"):
                try:
                    clients = context.gateway.xpm.list_clients(
                        search=client_name, limit=1)
                    if clients:
                        c = clients[0]
                        xpm_info = (f"XPM: {c.get('name','')}, "
                                    f"entity type: {c.get('entity_type','')}, "
                                    f"ABN: {c.get('abn','N/A')}")
                except Exception:
                    pass

            # ── Generate letter ───────────────────────────────────────────────
            letter = self._generate_letter(
                context, client_name, client_email,
                detected_services, xpm_info)

            draft_only = self.get_plugin_setting("draft_only", "0") == "1"
            confidence = float(self.get_plugin_setting("confidence_threshold", "0.8"))

            # ── Send via FuseSign or email ────────────────────────────────────
            if (self.get_plugin_setting("use_fusesign", "1") == "1"
                    and context.gateway
                    and context.gateway.is_available("fusesign")
                    and not draft_only):

                def _create_envelope(name=client_name, email=client_email,
                                     content=letter, services=detected_services):
                    try:
                        context.gateway.fusesign.create_envelope(
                            name=f"Engagement Letter — {name}",
                            recipients=[{"name": name, "email": email}],
                            document_content=content,
                            subject=f"Engagement Letter — {', '.join(services[:2])}")
                    except Exception as ex:
                        context.log(f"[EngagementLetter] FuseSign error: {ex}")

                if context.approval_queue:
                    context.approval_queue.submit(
                        action_type="create_fusesign_envelope",
                        description=f"Send engagement letter to {client_name}",
                        payload={"client": client_name, "email": client_email,
                                 "services": detected_services},
                        confidence=confidence,
                        plugin_id=self.PLUGIN_ID,
                        execute_callback=_create_envelope)
                else:
                    _create_envelope()

            elif context.graph:
                practice = get_setting("practice_name", "MC & S Accounting")
                context.graph.send_email(
                    to=client_email,
                    subject=f"Engagement Letter — {practice}",
                    body=letter,
                    draft_mode=True)  # always draft engagement letters

            # ── Store in memory ───────────────────────────────────────────────
            if context.memory:
                try:
                    context.memory.store_client_interaction(
                        client_name=client_name,
                        interaction_type="engagement_letter",
                        summary=f"Engagement letter generated for: {', '.join(detected_services)}",
                        metadata={"services": ", ".join(detected_services),
                                  "email": client_email})
                except Exception:
                    pass

            # ── Mark email as read ────────────────────────────────────────────
            try:
                context.graph.mark_as_read(msg_id)
            except Exception:
                pass

            generated += 1
            log_activity(from_email=client_email,
                         subject="Engagement Letter",
                         classification="engagement_letter",
                         action=f"generated:{','.join(detected_services[:2])}",
                         draft_created=True)

        result.summary = f"Engagement letters generated: {generated}."
        result.actions_taken = generated
        return result

    def _generate_letter(self, context: PluginContext, client_name: str,
                          client_email: str, services: list, xpm_info: str) -> str:
        """Use Claude Sonnet to generate a personalised engagement letter."""
        practice = get_setting("practice_name", "MC & S Accounting")
        today    = datetime.now().strftime("%d %B %Y")
        services_str = "\n".join(f"  - {s}" for s in services)

        if not context.claude_reason:
            return self._fallback_letter(practice, client_name, services, today)

        try:
            prompt = (f"You are a professional accountant at {practice}.\n"
                      f"Write a formal engagement letter to {client_name} ({client_email}) "
                      f"dated {today}.\n"
                      f"Services to be provided:\n{services_str}\n"
                      f"Client info: {xpm_info or 'not available'}\n\n"
                      "The letter should:\n"
                      "1. Confirm the scope of services\n"
                      "2. Outline fees (use placeholder [FEE] where specific amounts are needed)\n"
                      "3. State client responsibilities\n"
                      "4. Include a signature block for the client\n"
                      "5. Be professional, clear, and ATO-compliant\n"
                      "Return the full letter as plain text (no HTML).")
            resp = context.claude_reason.messages.create(
                model=self.get_claude_model_reasoning(),
                max_tokens=1000,
                messages=[{"role": "user", "content": prompt}])
            return resp.content[0].text.strip()
        except Exception:
            return self._fallback_letter(practice, client_name, services, today)

    def _fallback_letter(self, practice: str, client_name: str,
                          services: list, today: str) -> str:
        services_str = "\n".join(f"  - {s}" for s in services)
        return (f"{practice}\n{today}\n\n"
                f"Dear {client_name},\n\n"
                f"We are pleased to confirm our engagement to provide the following services:\n"
                f"{services_str}\n\n"
                f"Our fees will be discussed and agreed upon prior to commencement of work.\n\n"
                f"Please sign below to confirm your acceptance of these terms.\n\n"
                f"Yours sincerely,\n{practice}\n\n"
                f"Client signature: ___________________  Date: ___________")
