"""
MC & S Desktop Agent — Gateway Client (Stream 4 — APEX Upgrade)
================================================================
Provides a unified interface to external practice management and
communication platforms beyond Microsoft 365:

  - XPM (Xero Practice Manager) — client/job/task management via OAuth 2.0
  - FuseSign                    — document signing workflow
  - Microsoft Teams             — channel messages and notifications

DESIGN
------
Each integration is a separate class:
  XPMClient      — wraps the XPM REST API (OAuth 2.0 via xero_oauth.py)
  FuseSignClient — wraps the FuseSign REST API
  TeamsClient    — wraps the Microsoft Graph API (Teams endpoints)

GatewayClient is a facade that holds all three and exposes a single
`context.gateway` object to plugins.

CREDENTIALS
-----------
All credentials are stored in the existing SQLite settings table
via config.get_setting() / config.set_setting().  No plaintext secrets
are hard-coded here.

  Setting key              Description
  ─────────────────────    ─────────────────────────────────────────
  xero_client_id           MCS Mate Xero app Client ID
  xero_client_secret       MCS Mate Xero app Client Secret
  xero_refresh_token       Long-lived OAuth refresh token
  xero_access_token        Short-lived OAuth access token (auto-refreshed)
  xero_token_expiry        ISO timestamp of access token expiry
  xero_tenant_id           Xero organisation/tenant ID
  fusesign_api_key         FuseSign API key
  fusesign_base_url        FuseSign base URL (default: https://api.fusesign.com/v1)
  teams_webhook_url        Teams Incoming Webhook URL for a channel
  teams_graph_token        Microsoft Graph access token (for full Teams API)

USAGE IN PLUGINS
----------------
    from gateway_client import GatewayClient

    gw = GatewayClient()
    gw.load()  # reads credentials from settings

    # XPM
    clients = gw.xpm.list_clients(search="Smith")
    jobs    = gw.xpm.list_jobs(client_id="123", status="In Progress")

    # FuseSign
    envelope = gw.fusesign.create_envelope(
        template_id="tpl_abc",
        recipients=[{"name": "John Smith", "email": "john@example.com"}],
        subject="Tax Return FY2024 — Please Sign",
    )
    status = gw.fusesign.get_envelope_status(envelope["id"])

    # Teams
    gw.teams.send_message("New NOA processed for John Smith — REFUND $2,400")

    # Availability check
    print(gw.available_integrations())  # ["xpm", "fusesign", "teams"]
"""

import json
import urllib.request
import urllib.parse
import urllib.error
from typing import Any

from config import get_setting


# ── XPM Client ────────────────────────────────────────────────────────────────

class XPMClient:
    """
    Wraps the Xero Practice Manager (XPM) REST API.

    Authentication uses OAuth 2.0 Authorization Code flow via xero_oauth.py.
    Tokens are automatically refreshed before expiry.
    All methods return parsed JSON dicts/lists or raise XPMError on failure.
    """

    DEFAULT_BASE_URL = "https://api.xero.com/practicemanager/3.1"

    def __init__(self, base_url: str = ""):
        self.base_url = (base_url or get_setting("xpm_base_url", self.DEFAULT_BASE_URL)).rstrip("/")

    @property
    def is_configured(self) -> bool:
        """Return True if OAuth credentials are stored (client ID + secret)."""
        try:
            from xero_oauth import is_configured
            return is_configured()
        except Exception:
            return False

    @property
    def is_authorised(self) -> bool:
        """Return True if a refresh token is stored (user has completed OAuth)."""
        try:
            from xero_oauth import is_authorised
            return is_authorised()
        except Exception:
            return False

    def _get_token(self) -> str:
        """Return a valid access token, refreshing if needed."""
        try:
            from xero_oauth import get_valid_token
            return get_valid_token()
        except Exception as e:
            raise XPMError(f"Xero OAuth error: {e}") from e

    def _get_tenant_id(self) -> str:
        """Return the stored tenant ID."""
        tenant_id = get_setting("xero_tenant_id", "")
        if not tenant_id:
            raise XPMError(
                "Xero tenant ID not set. Please complete OAuth setup in Settings → Xero."
            )
        return tenant_id

    def _request(self, method: str, path: str, params: dict = None, body: dict = None) -> Any:
        """Make an authenticated request to the XPM API."""
        if not self.is_configured:
            raise XPMError("Xero not configured. Add Client ID and Secret in Settings → Xero.")
        if not self.is_authorised:
            raise XPMError("Xero not authorised. Complete OAuth setup in Settings → Xero.")

        access_token = self._get_token()
        tenant_id    = self._get_tenant_id()

        url = f"{self.base_url}{path}"
        if params:
            url += "?" + urllib.parse.urlencode(params)

        headers = {
            "Authorization":  f"Bearer {access_token}",
            "Xero-Tenant-Id": tenant_id,
            "Content-Type":   "application/json",
            "Accept":         "application/json",
        }

        data = json.dumps(body).encode() if body else None
        req = urllib.request.Request(url, data=data, headers=headers, method=method)

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            error_body = e.read().decode() if e.fp else ""
            raise XPMError(f"XPM API error {e.code}: {error_body}") from e
        except Exception as e:
            raise XPMError(f"XPM request failed: {e}") from e

    # ── Clients ───────────────────────────────────────────────────────────────

    def list_clients(self, search: str = "", page: int = 1, page_size: int = 50) -> list:
        """
        List XPM clients, optionally filtered by name/email search.

        Returns a list of client dicts with keys:
          id, name, email, phone, is_active, etc.
        """
        params = {"page": page, "pageSize": page_size}
        if search:
            params["search"] = search
        result = self._request("GET", "/clients", params=params)
        return result.get("items", result) if isinstance(result, dict) else result

    def list_all_clients(
        self,
        search: str = "",
        page_size: int = 50,
        staff_uuid: str = "",
    ) -> list:
        """
        Fetch ALL clients across all pages (Fix 13: XPM pagination).
        Optionally filter to clients assigned to a specific staff member
        (Fix 14: per-accountant filtering) by passing staff_uuid.
        """
        all_items: list = []
        page = 1
        while True:
            batch = self.list_clients(search=search, page=page, page_size=page_size)
            if not batch:
                break
            all_items.extend(batch)
            if len(batch) < page_size:
                break  # Last page
            page += 1
        if staff_uuid:
            all_items = [
                c for c in all_items
                if c.get("AccountManager", {}).get("UUID", "") == staff_uuid
                or c.get("account_manager_uuid", "") == staff_uuid
            ]
        return all_items

    def get_client(self, client_id: str) -> dict:
        """Return a single client by ID."""
        return self._request("GET", f"/clients/{client_id}")

    def get_client_by_email(self, email: str) -> dict | None:
        """Find a client by email address. Returns None if not found."""
        clients = self.list_clients(search=email, page_size=10)
        for c in clients:
            if c.get("email", "").lower() == email.lower():
                return c
        return None

    def get_client_by_name(self, name: str) -> dict | None:
        """
        Find a client by name (case-insensitive, partial match).
        Tries exact match first, then falls back to first partial match.
        Returns None if not found.
        """
        if not name:
            return None
        clients = self.list_clients(search=name, page_size=20)
        name_lower = name.lower()
        # Prefer exact match
        for c in clients:
            full_name = (c.get("Name") or c.get("name") or "").lower()
            if full_name == name_lower:
                return c
        # Fall back to first partial match
        for c in clients:
            full_name = (c.get("Name") or c.get("name") or "").lower()
            if name_lower in full_name or full_name in name_lower:
                return c
        return None

    # ── Jobs ──────────────────────────────────────────────────────────────────

    def list_jobs(
        self,
        client_id: str = "",
        status: str = "",
        job_type: str = "",
        page: int = 1,
        page_size: int = 50,
    ) -> list:
        """
        List XPM jobs, optionally filtered.

        status options: "In Progress", "Completed", "Not Started", "Cancelled"
        """
        params: dict = {"page": page, "pageSize": page_size}
        if client_id:
            params["clientId"] = client_id
        if status:
            params["status"] = status
        if job_type:
            params["jobType"] = job_type
        result = self._request("GET", "/jobs", params=params)
        return result.get("items", result) if isinstance(result, dict) else result

    def list_all_jobs(
        self,
        client_id: str = "",
        status: str = "",
        job_type: str = "",
        page_size: int = 50,
    ) -> list:
        """
        Fetch ALL jobs across all pages (Fix 13: XPM pagination).
        """
        all_items: list = []
        page = 1
        while True:
            batch = self.list_jobs(
                client_id=client_id, status=status,
                job_type=job_type, page=page, page_size=page_size
            )
            if not batch:
                break
            all_items.extend(batch)
            if len(batch) < page_size:
                break
            page += 1
        return all_items

    def get_job(self, job_id: str) -> dict:
        """Return a single job by ID."""
        return self._request("GET", f"/jobs/{job_id}")

    def update_job_status(self, job_id: str, status: str) -> dict:
        """Update the status of a job."""
        return self._request("PATCH", f"/jobs/{job_id}", body={"status": status})

    # ── Tasks ─────────────────────────────────────────────────────────────────

    def list_tasks(self, job_id: str, status: str = "") -> list:
        """List tasks for a job."""
        params: dict = {}
        if status:
            params["status"] = status
        result = self._request("GET", f"/jobs/{job_id}/tasks", params=params or None)
        return result.get("items", result) if isinstance(result, dict) else result

    def complete_task(self, job_id: str, task_id: str) -> dict:
        """Mark a task as complete."""
        return self._request("PATCH", f"/jobs/{job_id}/tasks/{task_id}",
                             body={"status": "Completed"})

    # ── Notes ─────────────────────────────────────────────────────────────────

    def add_client_note(self, client_id: str, note: str) -> dict:
        """Add a note to a client record."""
        return self._request("POST", f"/clients/{client_id}/notes",
                             body={"content": note})

    def add_job_note(self, job_id: str, note: str) -> dict:
        """Add a note to a job."""
        return self._request("POST", f"/jobs/{job_id}/notes",
                             body={"content": note})

    # ── Invoices ────────────────────────────────────────────────────────────────────────────

    def list_invoices(
        self,
        client_id: str = "",
        status: str = "",
        page: int = 1,
        page_size: int = 50,
    ) -> list:
        """
        List XPM invoices, optionally filtered by client or status.
        status options: 'Draft', 'Approved', 'Sent', 'Paid', 'Voided'
        """
        params: dict = {"page": page, "pageSize": page_size}
        if client_id:
            params["clientId"] = client_id
        if status:
            params["status"] = status
        result = self._request("GET", "/invoices", params=params)
        return result.get("items", result) if isinstance(result, dict) else result

    def list_all_invoices(
        self,
        client_id: str = "",
        status: str = "",
        page_size: int = 50,
    ) -> list:
        """Fetch ALL invoices across all pages."""
        all_items: list = []
        page = 1
        while True:
            batch = self.list_invoices(
                client_id=client_id, status=status,
                page=page, page_size=page_size
            )
            if not batch:
                break
            all_items.extend(batch)
            if len(batch) < page_size:
                break
            page += 1
        return all_items

    def get_debtor_summary(self) -> dict:
        """
        Return a summary of outstanding invoices grouped by ageing bucket.
        Returns: {"total_outstanding": float, "0_30": float, "31_60": float,
                  "61_90": float, "90_plus": float, "invoice_count": int}
        """
        from datetime import date
        today = date.today()
        invoices = self.list_all_invoices(status="Approved")  # Approved = sent but unpaid in XPM
        buckets = {"0_30": 0.0, "31_60": 0.0, "61_90": 0.0, "90_plus": 0.0}
        total = 0.0
        count = 0
        for inv in invoices:
            amount = float(inv.get("AmountDue") or inv.get("amount_due") or 0)
            if amount <= 0:
                continue
            due_str = inv.get("DueDate") or inv.get("due_date") or ""
            try:
                due_date = date.fromisoformat(due_str[:10])
                days_overdue = (today - due_date).days
            except Exception:
                days_overdue = 0
            total += amount
            count += 1
            if days_overdue <= 30:
                buckets["0_30"] += amount
            elif days_overdue <= 60:
                buckets["31_60"] += amount
            elif days_overdue <= 90:
                buckets["61_90"] += amount
            else:
                buckets["90_plus"] += amount
        return {"total_outstanding": round(total, 2), "invoice_count": count, **buckets}


class XPMError(Exception):
    """Raised when an XPM API call fails."""


# ── FuseSign Client ───────────────────────────────────────────────────────────

class FuseSignClient:
    """
    Wraps the FuseSign document signing API.

    FuseSign uses Bearer token authentication.
    """

    DEFAULT_BASE_URL = "https://api.fusesign.com/v1"

    def __init__(self, api_key: str = "", base_url: str = ""):
        self.api_key = api_key or get_setting("fusesign_api_key", "")
        self.base_url = (base_url or get_setting("fusesign_base_url", self.DEFAULT_BASE_URL)).rstrip("/")

    @property
    def is_configured(self) -> bool:
        return bool(self.api_key)

    def _request(self, method: str, path: str, params: dict = None, body: dict = None) -> Any:
        """Make an authenticated request to the FuseSign API."""
        if not self.is_configured:
            raise FuseSignError("FuseSign API key not configured. Add it in Settings.")

        url = f"{self.base_url}{path}"
        if params:
            url += "?" + urllib.parse.urlencode(params)

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        data = json.dumps(body).encode() if body else None
        req = urllib.request.Request(url, data=data, headers=headers, method=method)

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            error_body = e.read().decode() if e.fp else ""
            raise FuseSignError(f"FuseSign API error {e.code}: {error_body}") from e
        except Exception as e:
            raise FuseSignError(f"FuseSign request failed: {e}") from e

    # ── Envelopes ─────────────────────────────────────────────────────────────

    def create_envelope(
        self,
        template_id: str,
        recipients: list[dict],
        subject: str = "",
        message: str = "",
        metadata: dict = None,
    ) -> dict:
        """
        Create a signing envelope from a template.

        Parameters
        ----------
        template_id  : FuseSign template ID
        recipients   : List of dicts with 'name' and 'email' keys
        subject      : Email subject sent to signers
        message      : Optional message to signers
        metadata     : Optional dict of key/value pairs stored on the envelope

        Returns the created envelope dict including 'id' and 'signing_url'.
        """
        body: dict = {
            "templateId": template_id,
            "recipients": recipients,
        }
        if subject:
            body["subject"] = subject
        if message:
            body["message"] = message
        if metadata:
            body["metadata"] = metadata
        return self._request("POST", "/envelopes", body=body)

    def get_envelope_status(self, envelope_id: str) -> dict:
        """
        Get the current status of an envelope.

        Returns dict with keys: id, status, recipients (with signing status),
        created_at, completed_at, etc.

        Status values: "pending", "sent", "partially_signed",
                       "completed", "declined", "voided"
        """
        return self._request("GET", f"/envelopes/{envelope_id}")

    def list_envelopes(
        self,
        status: str = "",
        page: int = 1,
        page_size: int = 50,
    ) -> list:
        """List envelopes, optionally filtered by status."""
        params: dict = {"page": page, "pageSize": page_size}
        if status:
            params["status"] = status
        result = self._request("GET", "/envelopes", params=params)
        return result.get("items", result) if isinstance(result, dict) else result

    def void_envelope(self, envelope_id: str, reason: str = "") -> dict:
        """Void (cancel) an envelope."""
        body: dict = {}
        if reason:
            body["reason"] = reason
        return self._request("DELETE", f"/envelopes/{envelope_id}", body=body or None)

    def resend_envelope(self, envelope_id: str) -> dict:
        """Resend signing reminders to pending recipients."""
        return self._request("POST", f"/envelopes/{envelope_id}/resend")

    # ── Templates ─────────────────────────────────────────────────────────────

    def list_templates(self) -> list:
        """List all available FuseSign templates."""
        result = self._request("GET", "/templates")
        return result.get("items", result) if isinstance(result, dict) else result

    def get_template(self, template_id: str) -> dict:
        """Return a single template by ID."""
        return self._request("GET", f"/templates/{template_id}")


class FuseSignError(Exception):
    """Raised when a FuseSign API call fails."""


# ── Teams Client ──────────────────────────────────────────────────────────────

class TeamsClient:
    """
    Sends messages to Microsoft Teams.

    Supports two modes:
      1. Incoming Webhook (simple, no auth needed beyond the URL)
      2. Microsoft Graph API (full Teams API, requires access token)

    For most CoWorker use cases (notifications, alerts), the webhook
    mode is sufficient and much simpler to configure.
    """

    def __init__(self, webhook_url: str = "", graph_token: str = ""):
        self.webhook_url = webhook_url or get_setting("teams_webhook_url", "")
        self.graph_token = graph_token or get_setting("teams_graph_token", "")

    @property
    def is_configured(self) -> bool:
        return bool(self.webhook_url or self.graph_token)

    @property
    def webhook_configured(self) -> bool:
        return bool(self.webhook_url)

    @property
    def graph_configured(self) -> bool:
        return bool(self.graph_token)

    # ── Webhook mode ──────────────────────────────────────────────────────────

    def send_message(self, text: str, title: str = "", color: str = "") -> bool:
        """
        Send a simple text message to a Teams channel via Incoming Webhook.

        Parameters
        ----------
        text  : Message body (supports Markdown)
        title : Optional card title
        color : Optional accent colour hex (e.g. "0078D4" for blue)

        Returns True on success, False on failure.
        """
        if not self.webhook_url:
            raise TeamsError("Teams webhook URL not configured. Add it in Settings.")

        # Build an Adaptive Card payload
        card: dict = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.4",
                        "body": [],
                    },
                }
            ],
        }

        body_items = card["attachments"][0]["content"]["body"]

        if title:
            body_items.append({
                "type": "TextBlock",
                "text": title,
                "weight": "Bolder",
                "size": "Medium",
                "color": "Accent" if not color else "Default",
            })

        body_items.append({
            "type": "TextBlock",
            "text": text,
            "wrap": True,
        })

        if color:
            card["attachments"][0]["content"]["msteams"] = {
                "width": "Full",
            }

        payload = json.dumps(card).encode()
        req = urllib.request.Request(
            self.webhook_url,
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                return resp.status in (200, 202)
        except urllib.error.HTTPError as e:
            error_body = e.read().decode() if e.fp else ""
            raise TeamsError(f"Teams webhook error {e.code}: {error_body}") from e
        except Exception as e:
            raise TeamsError(f"Teams webhook request failed: {e}") from e

    def send_alert(self, title: str, body: str, urgent: bool = False) -> bool:
        """
        Send a formatted alert card to Teams.

        urgent=True uses a red accent colour to draw attention.
        """
        color = "FF0000" if urgent else "0078D4"
        return self.send_message(text=body, title=title, color=color)

    # ── Graph API mode ────────────────────────────────────────────────────────

    def send_channel_message(
        self,
        team_id: str,
        channel_id: str,
        content: str,
        content_type: str = "text",
    ) -> dict:
        """
        Post a message to a specific Teams channel via Microsoft Graph API.

        Requires a valid Graph access token with ChannelMessage.Send permission.
        """
        if not self.graph_token:
            raise TeamsError("Teams Graph token not configured. Add it in Settings.")

        url = (
            f"https://graph.microsoft.com/v1.0/teams/{team_id}"
            f"/channels/{channel_id}/messages"
        )
        body = {
            "body": {
                "contentType": content_type,
                "content": content,
            }
        }
        payload = json.dumps(body).encode()
        req = urllib.request.Request(
            url,
            data=payload,
            headers={
                "Authorization": f"Bearer {self.graph_token}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            error_body = e.read().decode() if e.fp else ""
            raise TeamsError(f"Teams Graph error {e.code}: {error_body}") from e
        except Exception as e:
            raise TeamsError(f"Teams Graph request failed: {e}") from e


class TeamsError(Exception):
    """Raised when a Teams API call fails."""


# ── GatewayClient facade ──────────────────────────────────────────────────────

class GatewayClient:
    """
    Unified facade for all external platform integrations.

    Plugins access this via context.gateway:

        context.gateway.xpm.list_clients(search="Smith")
        context.gateway.fusesign.create_envelope(...)
        context.gateway.teams.send_message("NOA processed for John Smith")
        context.gateway.available_integrations()  # ["xpm", "teams"]
    """

    def __init__(self):
        self.xpm = XPMClient()
        self.fusesign = FuseSignClient()
        self.teams = TeamsClient()

    def load(self) -> None:
        """Re-read credentials from settings (call after settings change)."""
        self.xpm = XPMClient()
        self.fusesign = FuseSignClient()
        self.teams = TeamsClient()

    def available_integrations(self) -> list[str]:
        """Return a list of integration names that are currently configured."""
        available = []
        if self.xpm.is_configured:
            available.append("xpm")
        if self.fusesign.is_configured:
            available.append("fusesign")
        if self.teams.is_configured:
            available.append("teams")
        return available

    def is_available(self, integration: str) -> bool:
        """Check if a specific integration is configured."""
        return integration in self.available_integrations()

    def status_summary(self) -> str:
        """Return a human-readable status string for all integrations."""
        lines = []
        if self.xpm.is_authorised:
            lines.append("XPM:      ✅ Connected (OAuth authorised)")
        elif self.xpm.is_configured:
            lines.append("XPM:      ⚠ Credentials set — OAuth not yet authorised")
        else:
            lines.append("XPM:      ⚠ Not configured")
        lines.append(f"FuseSign: {'✅ Configured' if self.fusesign.is_configured else '⚠ Not configured'}")
        lines.append(f"Teams:    {'✅ Configured' if self.teams.is_configured else '⚠ Not configured'}")
        return "\n".join(lines)
