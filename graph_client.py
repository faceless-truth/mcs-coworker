"""
MC & S Email Automation - Microsoft Graph API Client
Handles OAuth2 authentication and email operations.
"""
import threading
import webbrowser
import requests
import json
from datetime import datetime, timezone
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import msal

GRAPH_SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
REDIRECT_URI = "http://localhost:8765"


class GraphClient:

    def __init__(self, tenant_id: str, client_id: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self._token_cache = msal.SerializableTokenCache()
        self._app = None
        self._account = None
        self._access_token = None
        self._auth_code = None
        self._auth_event = threading.Event()
        self._setup_app()

    def _setup_app(self):
        if self.tenant_id and self.client_id:
            self._app = msal.PublicClientApplication(
                client_id=self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                token_cache=self._token_cache,
            )

    def is_configured(self):
        return bool(self.tenant_id and self.client_id)

    def is_authenticated(self):
        if not self._app:
            return False
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._access_token = result["access_token"]
                self._account = accounts[0]
                return True
        return False

    def authenticate(self, callback=None):
        """Open browser for OAuth2 login, capture code via local server."""
        if not self._app:
            raise ValueError(
                "Graph client not configured. Add Tenant ID and Client ID first."
            )

        # Start local server to capture redirect
        auth_code_container = {"code": None, "error": None}

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                parsed = urlparse(self.path)
                params = parse_qs(parsed.query)

                if "code" in params:
                    auth_code_container["code"] = params["code"][0]
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html")
                    self.end_headers()
                    self.wfile.write(b"""
<html><body style='font-family:Arial;text-align:center;padding:60px'>
<h2 style='color:#2E7D32'>&#10003; Authentication Successful</h2>
<p>You can close this window and return to MC&S Email Automation.</p>
</body></html>""")
                else:
                    auth_code_container["error"] = params.get(
                        "error", ["Unknown"]
                    )[0]
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(b"Authentication failed.")

            def log_message(self, format, *args):
                pass  # Suppress server logs

        server = HTTPServer(("localhost", 8765), Handler)
        server.timeout = 120  # 2 min timeout

        def run_server():
            server.handle_request()
            self._auth_event.set()

        threading.Thread(target=run_server, daemon=True).start()

        # Build auth URL and open browser
        auth_url = self._app.get_authorization_request_url(
            scopes=GRAPH_SCOPES,
            redirect_uri=REDIRECT_URI,
        )
        webbrowser.open(auth_url)

        def wait_and_complete():
            self._auth_event.wait(timeout=120)

            if auth_code_container["code"]:
                result = self._app.acquire_token_by_authorization_code(
                    code=auth_code_container["code"],
                    scopes=GRAPH_SCOPES,
                    redirect_uri=REDIRECT_URI,
                )
                if "access_token" in result:
                    self._access_token = result["access_token"]
                    accounts = self._app.get_accounts()
                    if accounts:
                        self._account = accounts[0]
                    if callback:
                        callback(success=True, error=None)
                else:
                    if callback:
                        callback(
                            success=False,
                            error=result.get(
                                "error_description", "Unknown error"
                            ),
                        )
            else:
                if callback:
                    callback(
                        success=False,
                        error=auth_code_container.get(
                            "error", "Timeout or cancelled"
                        ),
                    )

        threading.Thread(target=wait_and_complete, daemon=True).start()

    def _get_token(self):
        """Get a fresh access token."""
        if not self._app:
            raise ValueError("Graph client not configured.")

        accounts = self._app.get_accounts()
        if not accounts:
            raise ValueError("Not authenticated. Please sign in first.")

        result = self._app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

        raise ValueError("Token refresh failed. Please sign in again.")

    def _headers(self):
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    # ── API Methods ───────────────────────────────────────────────────────────

    def get_user_info(self):
        """Return the signed-in user's display name and email."""
        r = requests.get(f"{GRAPH_BASE}/me", headers=self._headers())
        r.raise_for_status()
        return r.json()

    def fetch_unread_emails(self, folder="Inbox", max_count=25):
        """Fetch unread emails from the specified folder."""
        params = {
            "$filter": "isRead eq false",
            "$top": max_count,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,body,bodyPreview,hasAttachments,toRecipients",
        }
        url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
        r = requests.get(url, headers=self._headers(), params=params)
        r.raise_for_status()
        return r.json().get("value", [])

    def mark_as_read(self, message_id: str):
        """Mark a message as read."""
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.patch(url, headers=self._headers(), json={"isRead": True})
        r.raise_for_status()

    def send_email(self, to_address: str, subject: str, body_html: str,
                   reply_to_id: str = None):
        """Send an email directly."""
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        }

        if reply_to_id:
            url = f"{GRAPH_BASE}/me/messages/{reply_to_id}/reply"
            r = requests.post(
                url, headers=self._headers(), json={"message": message}
            )
        else:
            url = f"{GRAPH_BASE}/me/sendMail"
            r = requests.post(
                url,
                headers=self._headers(),
                json={"message": message, "saveToSentItems": True},
            )

        r.raise_for_status()

    def create_draft(self, to_address: str, subject: str, body_html: str,
                     reply_to_id: str = None):
        """Save a draft reply without sending it. Returns the draft message ID."""
        if reply_to_id:
            # Create a reply draft
            url = f"{GRAPH_BASE}/me/messages/{reply_to_id}/createReply"
            r = requests.post(url, headers=self._headers(), json={})
            r.raise_for_status()
            draft = r.json()
            draft_id = draft["id"]

            # Update the draft with our content
            update_url = f"{GRAPH_BASE}/me/messages/{draft_id}"
            r2 = requests.patch(
                update_url,
                headers=self._headers(),
                json={
                    "subject": subject,
                    "body": {"contentType": "HTML", "content": body_html},
                },
            )
            r2.raise_for_status()
            return draft_id
        else:
            # Create a fresh draft
            url = f"{GRAPH_BASE}/me/messages"
            payload = {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body_html},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
                "isDraft": True,
            }
            r = requests.post(url, headers=self._headers(), json=payload)
            r.raise_for_status()
            return r.json()["id"]

    def get_draft_link(self, draft_id: str):
        """Get a deeplink to a draft in Outlook Web."""
        return f"https://outlook.office.com/mail/drafts"

    def flag_email(self, message_id: str):
        """Flag an email for follow-up."""
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.patch(
            url,
            headers=self._headers(),
            json={"flag": {"flagStatus": "flagged"}},
        )
        r.raise_for_status()

    def add_category(self, message_id: str, category: str):
        """Add an Outlook category/label to a message."""
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.patch(
            url, headers=self._headers(), json={"categories": [category]}
        )
        r.raise_for_status()

    # ── Signature Extraction ─────────────────────────────────────────────────

    _cached_signature: str = ""

    def get_signature_html(self) -> str:
        """Extract the user's email signature from a recent sent email.

        The Graph API does not expose signatures directly, so we look at
        the most recent sent email and extract everything after the last
        occurrence of common signature delimiters.
        """
        if self._cached_signature:
            return self._cached_signature

        try:
            url = f"{GRAPH_BASE}/me/mailFolders/SentItems/messages"
            params = {
                "$top": 5,
                "$orderby": "sentDateTime desc",
                "$select": "body",
            }
            r = requests.get(url, headers=self._headers(), params=params)
            r.raise_for_status()
            messages = r.json().get("value", [])

            for msg in messages:
                body = msg.get("body", {}).get("content", "")
                if not body:
                    continue

                # Look for common signature markers
                import re
                # Try to find signature after "Regards", "Kind regards",
                # "Warm regards", "Thanks", "Cheers", etc.
                patterns = [
                    r'(?i)((?:warm\s+|kind\s+|best\s+)?regards[,.]?\s*<br\s*/?>.*)',
                    r'(?i)(cheers[,.]?\s*<br\s*/?>.*)',
                    r'(?i)(thanks[,.]?\s*<br\s*/?>.*)',
                    r'(?i)(thank\s+you[,.]?\s*<br\s*/?>.*)',
                ]
                for pattern in patterns:
                    match = re.search(pattern, body, re.DOTALL)
                    if match:
                        sig = match.group(1).strip()
                        # Only use if it contains meaningful content
                        # (name, phone, image, etc.) beyond just the sign-off
                        if len(sig) > 100:
                            self._cached_signature = sig
                            return sig

                # Fallback: look for a <table> near the end (many signatures
                # use tables for layout)
                table_match = re.search(
                    r'(?i)((?:warm\s+|kind\s+|best\s+)?regards[\s\S]{0,50}<table[\s\S]*</table>)',
                    body, re.DOTALL
                )
                if table_match:
                    self._cached_signature = table_match.group(1).strip()
                    return self._cached_signature

        except Exception:
            pass  # Signature extraction is best-effort

        return ""

    def clear_signature_cache(self):
        """Force re-fetch of signature on next call."""
        self._cached_signature = ""

