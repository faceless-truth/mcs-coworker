"""
MC & S Email Automation - Microsoft Graph API Client
Handles OAuth2 authentication and email operations.
"""
import atexit
import html
import logging
import os
import re
import threading
import webbrowser
import requests
import json
from datetime import datetime, timezone
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import msal

logger = logging.getLogger(__name__)


def _get_cache_path() -> Path:
    """Return the on-disk location of the MSAL serialised token cache."""
    from config import DATA_DIR
    return DATA_DIR / ".msal_cache.bin"


def _load_token_cache() -> msal.SerializableTokenCache:
    """Load the token cache from disk, or return a fresh one on first run."""
    cache = msal.SerializableTokenCache()
    cache_path = _get_cache_path()
    logger.info(
        f"MSAL cache loaded from {cache_path}: "
        f"{'found' if cache_path.exists() else 'not found (fresh start)'}"
    )
    if cache_path.exists():
        try:
            cache.deserialize(cache_path.read_text(encoding="utf-8"))
        except Exception as e:
            logger.warning(f"Failed to load MSAL cache: {e}")
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache) -> None:
    """Persist the cache to disk after a token acquisition.

    Writes when MSAL marks the cache dirty OR when the file doesn't yet
    exist (first run). The fresh-start branch is needed because the dirty
    flag can be cleared between acquisition and the first save attempt,
    which previously left the cache file uncreated forever.
    """
    if cache is None:
        return
    cache_path = _get_cache_path()
    state_changed = bool(getattr(cache, "has_state_changed", False))
    logger.debug(
        f"_save_token_cache called: has_state_changed={state_changed}, "
        f"path={cache_path}, path_exists={cache_path.exists()}"
    )
    try:
        if state_changed or not cache_path.exists():
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            cache_path.write_text(cache.serialize(), encoding="utf-8")
            logger.info(
                f"MSAL token cache saved to {cache_path} "
                f"(state_changed={state_changed})"
            )
        else:
            logger.debug("MSAL token cache unchanged, skipping save")
    except Exception as e:
        logger.error(f"Failed to save MSAL cache: {e}")

GRAPH_SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "User.Read",
    "Files.Read.All",
    "Files.ReadWrite.All",
    "Sites.Read.All",
    "Sites.ReadWrite.All",
]

print(f"[Auth] Requesting scopes: {GRAPH_SCOPES}")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
REDIRECT_URI = "http://localhost:7842/auth/callback"

# Hardcoded MC&S Entra ID credentials — accountants don't need Azure setup
MCS_TENANT_ID = "88ea4eb1-1dce-414e-bbfe-ce0c51a5bd98"
MCS_CLIENT_ID = "0bad4828-f05f-4580-8e6f-ede96c098642"

# MC&S SharePoint — hardcoded, same for all accountants
SHAREPOINT_SITE_URL = "https://mcandscomau.sharepoint.com/sites/MCS354"
SHAREPOINT_LIBRARY = "Documents"
SHAREPOINT_CLIENT_BASE = "Server/Clients"

print(f"[SharePoint] GRAPH_SCOPES = {GRAPH_SCOPES}")
print(f"[SharePoint] Site URL = {SHAREPOINT_SITE_URL}")
print(f"[SharePoint] Library = {SHAREPOINT_LIBRARY}")
print(f"[SharePoint] Client base = {SHAREPOINT_CLIENT_BASE}")


_URL_LINKIFY_PATTERN = re.compile(
    r'(?<!href=")(?<!href=\')(?<!src=")(?<!src=\')(https?://[^\s<>"\']+)'
)


def linkify_urls(text: str) -> str:
    """Wrap plain-text http(s) URLs in HTML anchor tags.

    Skips URLs already inside an href="" or src="" attribute so existing
    anchors and inline images are not double-wrapped. Used by every outgoing
    email path so KB-sourced links (e.g. the Tax Return Checklist) arrive as
    clickable links in Outlook instead of raw text the client has to copy.
    """
    if not text:
        return text
    return _URL_LINKIFY_PATTERN.sub(r'<a href="\1">\1</a>', text)


def _odata_escape(value: str) -> str:
    """Escape a string for safe use inside an OData $filter single-quoted literal.

    In OData v4, a single quote inside a string literal is escaped by doubling
    it. Without this, a sender address like ``o'brien@example.com`` breaks the
    filter query, and a crafted value could manipulate the filter logic.
    """
    return value.replace("'", "''")


class GraphClient:

    def __init__(self, tenant_id: str = None, client_id: str = None):
        self.tenant_id = tenant_id or MCS_TENANT_ID
        self.client_id = client_id or MCS_CLIENT_ID
        # Load cached tokens from disk so the accountant isn't re-prompted to
        # sign in on every app restart. The refresh token is good for 90 days.
        self._token_cache = _load_token_cache()
        self._app = None
        self._account = None
        self._access_token = None
        self._auth_code = None
        self._auth_event = threading.Event()
        self._setup_app()
        # Belt-and-braces: flush the cache on interpreter shutdown. The normal
        # save points below cover every successful token acquisition.
        atexit.register(self._persist_cache)

    def _persist_cache(self) -> None:
        """Write the cache to disk if MSAL has marked it dirty."""
        _save_token_cache(self._token_cache)

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
            # Silent acquisition may have refreshed the token — flush to disk.
            self._persist_cache()
            if result and "access_token" in result:
                self._access_token = result["access_token"]
                self._account = accounts[0]
                return True
        return False

    def authenticate(self, callback=None):
        """Open browser for OAuth2 login.

        The redirect is handled by the Flask server's /auth/callback route
        (same port 7842 as the app), so the webview never navigates away.
        """
        if not self._app:
            raise ValueError(
                "Graph client not configured. Add Tenant ID and Client ID first."
            )

        # Reset state for this auth attempt
        self._auth_event.clear()
        self._pending_callback = callback
        self._pending_auth_code = None
        self._pending_auth_error = None

        # Build auth URL and open in the system browser
        auth_url = self._app.get_authorization_request_url(
            scopes=GRAPH_SCOPES,
            redirect_uri=REDIRECT_URI,
        )
        webbrowser.open(auth_url)

        def wait_and_complete():
            # Wait for /auth/callback to call receive_auth_code()
            # 60s — even if offline detection lets us through, we don't want
            # to block the backend init thread for 2 full minutes if the
            # Microsoft endpoint is unreachable mid-flow.
            self._auth_event.wait(timeout=60)

            if self._pending_auth_code:
                result = self._app.acquire_token_by_authorization_code(
                    code=self._pending_auth_code,
                    scopes=GRAPH_SCOPES,
                    redirect_uri=REDIRECT_URI,
                )
                # Persist regardless of outcome — MSAL may have written
                # diagnostic state even on failure.
                self._persist_cache()
                if "access_token" in result:
                    self._access_token = result["access_token"]
                    accounts = self._app.get_accounts()
                    if accounts:
                        self._account = accounts[0]
                    logger.info("MSAL token cache saved after authentication")
                    if callback:
                        callback(success=True, error=None)
                else:
                    if callback:
                        callback(
                            success=False,
                            error=result.get("error_description", "Unknown error"),
                        )
            else:
                if callback:
                    callback(
                        success=False,
                        error=self._pending_auth_error or "Timeout or cancelled",
                    )

        threading.Thread(target=wait_and_complete, daemon=True).start()

    def receive_auth_code(self, code: str = None, error: str = None):
        """Called by the Flask /auth/callback route when Microsoft redirects back."""
        self._pending_auth_code = code
        self._pending_auth_error = error
        self._auth_event.set()

    def _get_token(self):
        """Get a fresh access token."""
        if not self._app:
            raise ValueError("Graph client not configured.")

        accounts = self._app.get_accounts()
        if not accounts:
            raise ValueError("Not authenticated. Please sign in first.")

        result = self._app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        # Silent acquisition frequently refreshes the token — persist so the
        # updated refresh token survives the next restart.
        self._persist_cache()
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
        r = requests.get(f"{GRAPH_BASE}/me", headers=self._headers(), timeout=30)
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
        r = requests.get(url, headers=self._headers(), timeout=30, params=params)
        r.raise_for_status()
        return r.json().get("value", [])

    def mark_as_read(self, message_id: str):
        """Mark a message as read."""
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.patch(url, headers=self._headers(), timeout=30, json={"isRead": True})
        r.raise_for_status()

    def mark_as_unread(self, message_id: str):
        """Mark a message as unread so it re-highlights in the inbox."""
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.patch(url, headers=self._headers(), timeout=30, json={"isRead": False})
        r.raise_for_status()

    def _append_signature(self, body_html: str) -> str:
        """Append the user's configured signature to an outgoing HTML body.

        Centralised here so every draft/send goes through one place. Plugins
        do not need to call this directly — it is applied automatically by
        `send_email`, `create_draft`, and the `*_with_attachments` variants.
        """
        if not body_html:
            body_html = ""
        # Linkify before appending the signature: the body may contain raw
        # URLs from Claude/KB (e.g. checklist links), and we want every email
        # path to deliver them as clickable anchors. The signature is added
        # after, so its own anchors and data: image URIs are untouched.
        body_html = linkify_urls(body_html)
        try:
            sig = self.get_signature_html()
        except Exception:
            sig = ""
        if not sig:
            return body_html
        # Idempotency guard: if the caller has already embedded the signature
        # (e.g. legacy code paths), don't double it up.
        if sig and sig in body_html:
            return body_html
        return body_html + sig

    def send_email(self, to_address: str, subject: str, body_html: str,
                   reply_to_id: str = None):
        """Send an email directly."""
        body_html = self._append_signature(body_html)
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        }

        if reply_to_id:
            url = f"{GRAPH_BASE}/me/messages/{reply_to_id}/reply"
            r = requests.post(
                url, headers=self._headers(), timeout=30, json={"message": message}
            )
        else:
            url = f"{GRAPH_BASE}/me/sendMail"
            r = requests.post(
                url,
                headers=self._headers(), timeout=30,
                json={"message": message, "saveToSentItems": True},
            )

        r.raise_for_status()

    def create_draft(self, to_address: str, subject: str, body_html: str,
                     reply_to_id: str = None):
        """Save a draft reply without sending it. Returns the draft message ID."""
        body_html = self._append_signature(body_html)
        if reply_to_id:
            try:
                draft_id = self._create_threaded_reply_draft(reply_to_id, body_html)
                # Do NOT mark the original as unread — that creates an infinite
                # loop where the next plugin run sees it as unread again and
                # drafts a duplicate reply. Callers are responsible for
                # marking the original as read after a successful draft.
                return draft_id
            except Exception as e:
                logger.warning(
                    "createReply failed for %s: %s — falling back to standalone draft",
                    reply_to_id, e,
                )
                # Fall through to standalone draft below.

        url = f"{GRAPH_BASE}/me/messages"
        payload = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
            "isDraft": True,
        }
        r = requests.post(url, headers=self._headers(), timeout=30, json=payload)
        r.raise_for_status()
        return r.json()["id"]

    def _create_threaded_reply_draft(self, reply_to_id: str, body_html: str) -> str:
        """Create a reply draft linked to the original conversation.

        Uses Graph's `createReply` so the draft inherits the original's
        conversationId, To, "Re: " subject, and quoted body. We then PATCH our
        AI-generated body in *above* the quoted original so the accountant can
        see the conversation context. Returns the new draft's id.
        """
        url = f"{GRAPH_BASE}/me/messages/{reply_to_id}/createReply"
        r = requests.post(url, headers=self._headers(), timeout=30, json={})
        r.raise_for_status()
        draft = r.json()
        draft_id = draft["id"]

        quoted_body = (draft.get("body") or {}).get("content", "")
        combined = body_html + quoted_body

        # IMPORTANT: this PATCH must contain ONLY the `body` key. createReply
        # has already set the subject to "RE: <original subject>" — including
        # `subject` here (even as None or an empty string) blanks it out, and
        # the draft shows up in Outlook as "(No subject)". Likewise do not add
        # toRecipients here; createReply inherited those too. If you need to
        # update something else, do it in a separate PATCH call.
        patch_payload = {"body": {"contentType": "HTML", "content": combined}}
        update_url = f"{GRAPH_BASE}/me/messages/{draft_id}"
        r2 = requests.patch(
            update_url, headers=self._headers(), timeout=30,
            json=patch_payload,
        )
        r2.raise_for_status()
        return draft_id

    def _reopen_original(self, message_id: str) -> None:
        """Flip the original email back to unread so the accountant sees
        the draft reply highlighted in their inbox. Graph's `createReply`
        marks the parent as read — we undo that here. Non-fatal on failure."""
        try:
            self.mark_as_unread(message_id)
        except Exception as e:
            logger.warning("mark_as_unread(%s) failed: %s", message_id, e)

    def get_draft_link(self, draft_id: str):
        """Get a deeplink to a draft in Outlook Web."""
        return f"https://outlook.office.com/mail/drafts"

    def flag_email(self, message_id: str):
        """Flag an email for follow-up."""
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.patch(
            url,
            headers=self._headers(), timeout=30,
            json={"flag": {"flagStatus": "flagged"}},
        )
        r.raise_for_status()

    def add_category(self, message_id: str, category: str):
        """Add an Outlook category/label to a message."""
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.patch(
            url, headers=self._headers(), timeout=30, json={"categories": [category]}
        )
        r.raise_for_status()

    def remove_category(self, message_id: str, category: str):
        """Remove a specific Outlook category tag from a message."""
        # Fetch current categories first, then patch with the tag removed
        url = f"{GRAPH_BASE}/me/messages/{message_id}"
        r = requests.get(
            url, headers=self._headers(), timeout=30,
            params={"$select": "categories"}
        )
        r.raise_for_status()
        current = r.json().get("categories", [])
        updated = [c for c in current if c != category]
        r2 = requests.patch(
            url, headers=self._headers(), timeout=30, json={"categories": updated}
        )
        r2.raise_for_status()

    # ── Signature ─────────────────────────────────────────────────────────────

    _cached_signature: str = ""

    def get_signature_image_path(self) -> str | None:
        """Return the path to the uploaded signature image at DATA_DIR/signature.png."""
        try:
            from config import DATA_DIR
            path = DATA_DIR / "signature.png"
            if path.is_file():
                return str(path)
        except Exception:
            pass
        return None

    def get_signature_html(self) -> str:
        """Return an HTML signature block.

        Precedence:
          1. Uploaded signature image (DATA_DIR/signature.png) — embedded inline
             as base64. If `email_signature` setting has text, it is prepended
             as HTML paragraphs above the image (e.g. "Kind regards, \\n Name").
          2. Text-only signature from `email_signature` setting.
          3. Auto-detection from recent sent emails.
        """
        # Load the configurable text prefix (used above the image, or standalone).
        try:
            from config import get_setting
            user_sig_text = (get_setting("email_signature", "") or "").strip()
        except Exception:
            user_sig_text = ""

        # 1. Uploaded signature image — preferred when present.
        sig_path = self.get_signature_image_path()
        if sig_path:
            try:
                import base64 as _b64
                with open(sig_path, "rb") as f:
                    b64 = _b64.b64encode(f.read()).decode("ascii")
                parts = ["<br><br>"]
                if user_sig_text:
                    if "<" in user_sig_text and ">" in user_sig_text:
                        parts.append(user_sig_text)
                    else:
                        escaped = html.escape(user_sig_text).replace("\n", "<br>")
                        parts.append(f"<p>{escaped}</p>")
                parts.append(
                    f'<img src="data:image/png;base64,{b64}" '
                    f'alt="signature" style="max-width:400px;">'
                )
                return "".join(parts)
            except Exception as e:
                logger.warning("Failed to embed signature image: %s", e)
                # Fall through to text-only signature.

        # 2. Text-only signature from Settings.
        if user_sig_text:
            if "<" in user_sig_text and ">" in user_sig_text:
                rendered = user_sig_text
            else:
                rendered = html.escape(user_sig_text).replace("\n", "<br>")
            return "<br><br>" + rendered

        if self._cached_signature:
            return self._cached_signature

        try:
            url = f"{GRAPH_BASE}/me/mailFolders/SentItems/messages"
            params = {
                "$top": 5,
                "$orderby": "sentDateTime desc",
                "$select": "body",
            }
            r = requests.get(url, headers=self._headers(), timeout=30, params=params)
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

    # ── Attachment Methods ───────────────────────────────────────────────────

    def get_attachments(self, message_id: str) -> list[dict]:
        """Return a list of attachments for a given message.
        Each dict has: id, name, contentType, size, contentBytes (base64).
        """
        url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        return r.json().get("value", [])

    def get_attachment_metadata(self, message_id: str) -> list[dict]:
        """Return lightweight attachment metadata (name, size, content type) for
        a message — no contentBytes payload. Used by plugins that just need to
        know whether files are attached and what they are called, so Claude
        doesn't ask the sender to re-send files that are already there."""
        url = (
            f"{GRAPH_BASE}/me/messages/{message_id}/attachments"
            f"?$select=name,size,contentType"
        )
        try:
            r = requests.get(url, headers=self._headers(), timeout=30)
            r.raise_for_status()
            value = r.json().get("value", [])
        except Exception as e:
            logger.warning("get_attachment_metadata(%s) failed: %s", message_id, e)
            return []
        return [
            {
                "name": att.get("name", "unknown"),
                "size": att.get("size", 0),
                "content_type": att.get("contentType", ""),
            }
            for att in value
        ]

    def download_attachment(self, message_id: str, attachment_id: str,
                           save_dir: str) -> str:
        """Download a specific attachment and save to save_dir.
        Returns the full path to the saved file.
        """
        import base64
        url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments/{attachment_id}"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        att = r.json()
        filename = att.get("name", "attachment")
        content = base64.b64decode(att.get("contentBytes", ""))
        os.makedirs(save_dir, exist_ok=True)
        filepath = os.path.join(save_dir, filename)
        with open(filepath, "wb") as f:
            f.write(content)
        return filepath

    def download_all_attachments(self, message_id: str,
                                save_dir: str) -> list[str]:
        """Download all attachments for a message. Returns list of file paths."""
        attachments = self.get_attachments(message_id)
        paths = []
        for att in attachments:
            if att.get("@odata.type", "") == "#microsoft.graph.fileAttachment":
                import base64
                filename = att.get("name", "attachment")
                content = base64.b64decode(att.get("contentBytes", ""))
                os.makedirs(save_dir, exist_ok=True)
                filepath = os.path.join(save_dir, filename)
                with open(filepath, "wb") as f:
                    f.write(content)
                paths.append(filepath)
        return paths

    # ── Email Search & Filtering ─────────────────────────────────────────────

    def search_emails(self, query: str, folder: str = "Inbox",
                      max_count: int = 50) -> list[dict]:
        """Search emails using OData $search query.
        E.g. query='subject:Notice of Assessment' or 'from:noreply@nowinfinity.com.au'
        """
        url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
        params = {
            "$search": f'"{query}"',
            "$top": max_count,
            "$select": "id,subject,from,receivedDateTime,body,bodyPreview,hasAttachments,toRecipients",
        }
        r = requests.get(url, headers=self._headers(), timeout=30, params=params)
        r.raise_for_status()
        return r.json().get("value", [])

    def fetch_emails_from_sender(self, sender_email: str,
                                 folder: str = "Inbox",
                                 unread_only: bool = True,
                                 max_count: int = 50) -> list[dict]:
        """Fetch emails from a specific sender address."""
        filter_str = f"from/emailAddress/address eq '{_odata_escape(sender_email)}'"
        if unread_only:
            filter_str += " and isRead eq false"
        url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
        params = {
            "$filter": filter_str,
            "$top": max_count,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,body,bodyPreview,hasAttachments,toRecipients",
        }
        r = requests.get(url, headers=self._headers(), timeout=30, params=params)
        r.raise_for_status()
        return r.json().get("value", [])

    def fetch_recent_emails(self, folder: str = "Inbox",
                            max_count: int = 50,
                            since_datetime: str = None) -> list[dict]:
        """Fetch recent emails, optionally filtered by date.
        since_datetime should be ISO format e.g. '2025-01-01T00:00:00Z'
        """
        filter_str = ""
        if since_datetime:
            filter_str = f"receivedDateTime ge {since_datetime}"
        url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
        params = {
            "$top": max_count,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,body,bodyPreview,hasAttachments,toRecipients,categories",
        }
        if filter_str:
            params["$filter"] = filter_str
        r = requests.get(url, headers=self._headers(), timeout=30, params=params)
        r.raise_for_status()
        return r.json().get("value", [])

    # ── Email with Attachments ───────────────────────────────────────────────

    def send_email_with_attachments(self, to_address: str, subject: str,
                                    body_html: str,
                                    attachment_paths: list[str] = None,
                                    reply_to_id: str = None):
        """Send an email with file attachments."""
        import base64
        body_html = self._append_signature(body_html)
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        }
        if attachment_paths:
            attachments = []
            for path in attachment_paths:
                filename = os.path.basename(path)
                with open(path, "rb") as f:
                    content = base64.b64encode(f.read()).decode("utf-8")
                attachments.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentBytes": content,
                })
            message["attachments"] = attachments

        url = f"{GRAPH_BASE}/me/sendMail"
        r = requests.post(
            url, headers=self._headers(), timeout=30,
            json={"message": message, "saveToSentItems": True},
        )
        r.raise_for_status()

    def create_draft_with_attachments(self, to_address: str, subject: str,
                                      body_html: str,
                                      attachment_paths: list[str] = None,
                                      reply_to_id: str = None) -> str:
        """Create a draft email with file attachments. Returns draft ID."""
        import base64
        body_html = self._append_signature(body_html)

        threaded = False
        draft_id = None
        if reply_to_id:
            try:
                draft_id = self._create_threaded_reply_draft(reply_to_id, body_html)
                threaded = True
            except Exception as e:
                logger.warning(
                    "createReply failed for %s: %s — falling back to standalone draft",
                    reply_to_id, e,
                )

        if draft_id is None:
            url = f"{GRAPH_BASE}/me/messages"
            payload = {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body_html},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
                "isDraft": True,
            }
            r = requests.post(url, headers=self._headers(), timeout=30, json=payload)
            r.raise_for_status()
            draft_id = r.json()["id"]

        if attachment_paths:
            for path in attachment_paths:
                filename = os.path.basename(path)
                with open(path, "rb") as f:
                    content = base64.b64encode(f.read()).decode("utf-8")
                att_url = f"{GRAPH_BASE}/me/messages/{draft_id}/attachments"
                att_payload = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentBytes": content,
                }
                r = requests.post(att_url, headers=self._headers(), timeout=30, json=att_payload)
                r.raise_for_status()

        # Intentionally do NOT mark the original as unread here — the caller
        # must mark it read after drafting to avoid duplicate-draft loops.

        return draft_id

    # ── Email with Inline Images ─────────────────────────────────────────────

    def _add_inline_image_to_draft(self, draft_id: str, image_path: str,
                                    content_id: str = "signature_image"):
        """Attach an image as an inline (CID) attachment to an existing draft."""
        import base64
        filename = os.path.basename(image_path)
        ext = os.path.splitext(filename)[1].lower()
        content_type_map = {
            ".png": "image/png", ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg", ".gif": "image/gif",
        }
        content_type = content_type_map.get(ext, "image/png")

        with open(image_path, "rb") as f:
            content_bytes = base64.b64encode(f.read()).decode("utf-8")

        att_url = f"{GRAPH_BASE}/me/messages/{draft_id}/attachments"
        att_payload = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": filename,
            "contentType": content_type,
            "contentBytes": content_bytes,
            "contentId": content_id,
            "isInline": True,
        }
        r = requests.post(att_url, headers=self._headers(), timeout=30, json=att_payload)
        r.raise_for_status()

    def create_draft_with_inline_image(self, to_address: str, subject: str,
                                        body_html: str,
                                        image_path: str,
                                        content_id: str = "signature_image",
                                        reply_to_id: str = None) -> str:
        """Create a draft email with an inline image attachment. Returns draft ID."""
        # Create the draft first (reuse existing method)
        draft_id = self.create_draft(to_address, subject, body_html, reply_to_id)
        # Add the inline image
        self._add_inline_image_to_draft(draft_id, image_path, content_id)
        return draft_id

    def send_email_with_inline_image(self, to_address: str, subject: str,
                                      body_html: str,
                                      image_path: str,
                                      content_id: str = "signature_image",
                                      reply_to_id: str = None):
        """Send an email with an inline image. Creates draft, attaches, then sends."""
        import base64
        filename = os.path.basename(image_path)
        ext = os.path.splitext(filename)[1].lower()
        content_type_map = {
            ".png": "image/png", ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg", ".gif": "image/gif",
        }
        content_type = content_type_map.get(ext, "image/png")

        with open(image_path, "rb") as f:
            content_bytes = base64.b64encode(f.read()).decode("utf-8")

        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": filename,
                "contentType": content_type,
                "contentBytes": content_bytes,
                "contentId": content_id,
                "isInline": True,
            }],
        }

        if reply_to_id:
            url = f"{GRAPH_BASE}/me/messages/{reply_to_id}/reply"
            r = requests.post(
                url, headers=self._headers(), timeout=30, json={"message": message}
            )
        else:
            url = f"{GRAPH_BASE}/me/sendMail"
            r = requests.post(
                url, headers=self._headers(), timeout=30,
                json={"message": message, "saveToSentItems": True},
            )
        r.raise_for_status()

    # ── Folder Operations ────────────────────────────────────────────────────

    def move_email(self, message_id: str, destination_folder: str):
        """Move an email to a different folder by folder name or ID."""
        # First try to resolve folder name to ID
        folder_id = self._resolve_folder_id(destination_folder)
        url = f"{GRAPH_BASE}/me/messages/{message_id}/move"
        r = requests.post(
            url, headers=self._headers(), timeout=30,
            json={"destinationId": folder_id},
        )
        r.raise_for_status()
        return r.json()

    def _resolve_folder_id(self, folder_name: str) -> str:
        """Resolve a folder name to its Graph API ID."""
        # Well-known folder names work directly
        well_known = [
            "inbox", "drafts", "sentitems", "deleteditems",
            "archive", "junkemail", "outbox",
        ]
        if folder_name.lower().replace(" ", "") in well_known:
            return folder_name
        # Otherwise search for the folder
        url = f"{GRAPH_BASE}/me/mailFolders"
        params = {"$filter": f"displayName eq '{_odata_escape(folder_name)}'"}
        r = requests.get(url, headers=self._headers(), timeout=30, params=params)
        r.raise_for_status()
        folders = r.json().get("value", [])
        if folders:
            return folders[0]["id"]
        # If not found, create it
        r = requests.post(
            url, headers=self._headers(), timeout=30,
            json={"displayName": folder_name},
        )
        r.raise_for_status()
        return r.json()["id"]

    def create_folder(self, folder_name: str) -> str:
        """Create a mail folder if it doesn't exist. Returns folder ID."""
        return self._resolve_folder_id(folder_name)


    # ── SharePoint ───────────────────────────────────────────────────────────

    def _make_request(self, method: str, url: str, json=None,
                      data=None, headers: dict = None,
                      auth: bool = True, timeout: int = 60):
        """Generic Graph API helper used by SharePoint methods.

        Supports GET / POST / PUT / PATCH. Returns the parsed JSON body on
        success, or None on 404 / HTTP error so callers can degrade gracefully
        rather than raising. Existing Outlook methods keep their explicit
        raise_for_status() behaviour.

        Set ``auth=False`` for pre-authenticated URLs (e.g. SharePoint upload
        session chunks) where Microsoft's docs explicitly say not to add an
        Authorization header.
        """
        method_upper = method.upper()
        print(f"[Graph API] {method_upper} {url}")
        try:
            request_headers = self._headers() if auth else {}
            if headers:
                # Allow Content-Type override (e.g., octet-stream for uploads).
                request_headers.update(headers)
            kwargs = {"headers": request_headers, "timeout": timeout}
            if json is not None:
                kwargs["json"] = json
            if data is not None:
                kwargs["data"] = data
            r = requests.request(method_upper, url, **kwargs)
            print(f"[Graph API] {method_upper} {url} -> {r.status_code}")
            if r.status_code == 404:
                return None
            if r.status_code >= 400:
                logger.warning(
                    "Graph API %s %s returned %s: %s",
                    method_upper, url, r.status_code, r.text[:200],
                )
                print(f"[Graph API] error body: {r.text[:500]}")
                return None
            try:
                return r.json()
            except ValueError:
                return {}
        except Exception as e:
            logger.error("Graph API %s %s failed: %s", method, url, e)
            print(f"[Graph API] {method_upper} {url} EXCEPTION: {e}")
            return None

    def _get_sharepoint_config(self) -> dict:
        """Return the hardcoded MC&S SharePoint config."""
        return {
            "site_url": SHAREPOINT_SITE_URL,
            "library": SHAREPOINT_LIBRARY,
            "client_base": SHAREPOINT_CLIENT_BASE,
        }

    def get_sharepoint_site_id(self) -> str | None:
        """Resolve the SharePoint site ID via the same auth path as email.

        Strategy 1: GET /sites/{hostname}:{path} (exact path match).
        Strategy 2: GET /sites?search={site_name} (fallback).
        """
        config = self._get_sharepoint_config()
        parsed = urlparse(config["site_url"])
        hostname = parsed.hostname
        site_path = parsed.path or ""
        print(
            f"[SharePoint] site lookup: hostname={hostname} site_path={site_path}"
        )
        if not hostname:
            return None

        # Strategy 1 — direct hostname:/path lookup, via _make_request
        # (the SAME method that works for email).
        url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
        print(f"[SharePoint] Looking up site: {url}")
        result = self._make_request("GET", url)
        print(f"[SharePoint] Result: {result}")
        if result and "id" in result:
            print(f"[SharePoint] site resolved via direct lookup: {result['id']}")
            return result["id"]

        # Strategy 2 — search by site name (last path segment)
        site_name = site_path.rstrip("/").split("/")[-1] if site_path else ""
        if site_name:
            search_url = f"{GRAPH_BASE}/sites?search={site_name}"
            print(f"[SharePoint] direct lookup failed; trying search: {search_url}")
            search_result = self._make_request("GET", search_url)
            if search_result and "value" in search_result:
                target = config["site_url"].rstrip("/").lower()
                for site in search_result["value"]:
                    if site.get("webUrl", "").rstrip("/").lower() == target:
                        print(f"[SharePoint] site resolved via search (exact): {site['id']}")
                        return site["id"]
                if search_result["value"]:
                    sid = search_result["value"][0]["id"]
                    print(f"[SharePoint] site resolved via search (first match): {sid}")
                    return sid
        return None

    def get_sharepoint_drive_id(self, site_id: str) -> str | None:
        """Get the drive ID for the configured document library."""
        config = self._get_sharepoint_config()
        url = f"{GRAPH_BASE}/sites/{site_id}/drives"
        result = self._make_request("GET", url)
        if result and "value" in result:
            for drive in result["value"]:
                if drive.get("name", "").lower() == config["library"].lower():
                    return drive["id"]
            # Fallback: return the first drive
            if result["value"]:
                return result["value"][0]["id"]
        return None

    def sharepoint_folder_exists(self, client_name: str,
                                 entity_name: str = None) -> bool:
        """Check if a client folder exists in SharePoint."""
        site_id = self.get_sharepoint_site_id()
        if not site_id:
            return False
        drive_id = self.get_sharepoint_drive_id(site_id)
        if not drive_id:
            return False
        config = self._get_sharepoint_config()
        folder_path = f"{config['client_base']}/{client_name}"
        if entity_name:
            folder_path += f"/{entity_name}"
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}"
        result = self._make_request("GET", url)
        return result is not None and "id" in result

    def upload_to_sharepoint(
        self,
        file_content: bytes,
        filename: str,
        client_name: str,
        entity_name: str = None,
        subfolder: str = None,
    ) -> str | None:
        """Upload a file to a client's SharePoint folder.

        Returns the SharePoint webUrl if successful, None if failed.
        Path: /{library}/{client_base}/{client_name}/{entity_name}/{subfolder}/{filename}
        """
        site_id = self.get_sharepoint_site_id()
        if not site_id:
            logger.error("SharePoint site ID not found at %s", SHAREPOINT_SITE_URL)
            return None
        drive_id = self.get_sharepoint_drive_id(site_id)
        if not drive_id:
            logger.error("SharePoint drive ID not found for library '%s'", SHAREPOINT_LIBRARY)
            return None

        config = self._get_sharepoint_config()
        folder_path = config["client_base"]
        if client_name:
            folder_path += f"/{client_name}"
        if entity_name:
            folder_path += f"/{entity_name}"
        if subfolder:
            folder_path += f"/{subfolder}"

        file_path = f"{folder_path}/{filename}"
        result = None

        if len(file_content) < 4 * 1024 * 1024:  # < 4MB → simple upload
            url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file_path}:/content"
            headers = {"Content-Type": "application/octet-stream"}
            result = self._make_request("PUT", url, data=file_content, headers=headers)
        else:
            # Large files → upload session in 10MB chunks
            url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file_path}:/createUploadSession"
            session = self._make_request("POST", url, json={
                "item": {"@microsoft.graph.conflictBehavior": "rename"}
            })
            if not session or "uploadUrl" not in session:
                logger.error("Failed to create upload session")
                return None
            upload_url = session["uploadUrl"]
            chunk_size = 10 * 1024 * 1024
            total = len(file_content)
            for i in range(0, total, chunk_size):
                chunk = file_content[i:i + chunk_size]
                end = min(i + chunk_size, total)
                chunk_headers = {
                    "Content-Range": f"bytes {i}-{end - 1}/{total}",
                    "Content-Type": "application/octet-stream",
                }
                # Pre-authenticated upload session URL — must NOT carry the
                # Bearer token, per Microsoft Graph upload-session docs.
                chunk_result = self._make_request(
                    "PUT", upload_url, data=chunk,
                    headers=chunk_headers, auth=False, timeout=120,
                )
                if chunk_result is None:
                    logger.error("Upload chunk failed at offset %d", i)
                    return None
                # The final chunk's response is the created file's metadata
                # (contains webUrl); intermediate chunks return a JSON body
                # with nextExpectedRanges.
                if "webUrl" in chunk_result:
                    result = chunk_result

        if result and "webUrl" in result:
            logger.info("Uploaded to SharePoint: %s", file_path)
            return result["webUrl"]
        return None

    def list_sharepoint_folder(self, client_name: str,
                               entity_name: str = None) -> list:
        """List files in a client's SharePoint folder."""
        site_id = self.get_sharepoint_site_id()
        if not site_id:
            return []
        drive_id = self.get_sharepoint_drive_id(site_id)
        if not drive_id:
            return []
        config = self._get_sharepoint_config()
        folder_path = config["client_base"]
        if client_name:
            folder_path += f"/{client_name}"
        if entity_name:
            folder_path += f"/{entity_name}"
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}:/children"
        result = self._make_request("GET", url)
        if result and "value" in result:
            return [
                {
                    "name": item["name"],
                    "size": item.get("size", 0),
                    "modified": item.get("lastModifiedDateTime", ""),
                    "url": item.get("webUrl", ""),
                    "is_folder": "folder" in item,
                }
                for item in result["value"]
            ]
        return []

    def test_sharepoint_connection(self) -> dict:
        """Test the SharePoint connection using the SAME token path as email.

        Routes every Graph call through `_make_request`, which is the helper
        the working email methods rely on. No `is_authenticated()` gate — that
        check reads the on-disk MSAL cache and was returning false even when
        the in-memory access token was perfectly valid (and email was working
        fine through the same client). If `_make_request` can't authenticate,
        its console output shows the exact URL and HTTP status.
        """
        print("[SharePoint Test] Starting connection test...")

        # Gate: check M365 authentication before attempting any Graph calls.
        # If not authenticated, give a clear actionable error instead of a
        # cryptic 'Connection failed'.
        try:
            self._get_token()
        except ValueError as e:
            return {"ok": False, "error": f"Not signed in to Microsoft 365 — {e}"}

        try:
            # Step 1 — find the site
            site_id = self.get_sharepoint_site_id()
            if not site_id:
                return {"ok": False, "error": f"Could not find SharePoint site at {SHAREPOINT_SITE_URL}"}

            # Step 2 — find the drive
            drive_id = self.get_sharepoint_drive_id(site_id)
            if not drive_id:
                return {"ok": False, "error": f"Could not find document library '{SHAREPOINT_LIBRARY}'"}

            # Step 3 — check the client base folder
            config = self._get_sharepoint_config()
            url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{config['client_base']}:/children?$top=1"
            result = self._make_request("GET", url)
            if result and "value" in result:
                return {
                    "ok": True,
                    "message": f"Connected to SharePoint. Found client folders in /{config['client_base']}/",
                }

            return {
                "ok": False,
                "error": f"Connected to site but could not find /{config['client_base']}/ folder",
            }
        except Exception as e:
            logger.error("SharePoint connection test failed: %s", e)
            return {"ok": False, "error": f"SharePoint test error: {e}"}

    # ── Calendar ──────────────────────────────────────────────────────────────

    def get_calendar_events(self, days_ahead: int = 1) -> list[dict]:
        """
        Fetch calendar events for the next days_ahead days using the
        Microsoft Graph /me/calendarView endpoint.
        Fix 16: enables meeting_prep to pull real calendar data instead of
        relying solely on email keyword detection.

        Returns a list of event dicts with keys:
          id, subject, start, end, location, organizer, attendees, bodyPreview
        """
        from datetime import timezone, timedelta
        now = datetime.utcnow().replace(tzinfo=timezone.utc)
        end = now + timedelta(days=days_ahead)
        start_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
        end_str   = end.strftime("%Y-%m-%dT%H:%M:%SZ")
        params = {
            "startDateTime": start_str,
            "endDateTime":   end_str,
            "$select": "id,subject,start,end,location,organizer,attendees,bodyPreview",
            "$orderby": "start/dateTime asc",
            "$top": 50,
        }
        url = f"{GRAPH_BASE}/me/calendarView"
        r = requests.get(url, headers=self._headers(), params=params, timeout=30)
        r.raise_for_status()
        return r.json().get("value", [])
