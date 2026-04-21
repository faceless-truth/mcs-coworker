"""
xero_oauth.py — Xero OAuth 2.0 Authorization Code Flow for MCS CoWorker
========================================================================
Handles the full OAuth lifecycle for the MCS Mate Xero Web App:

  1. build_auth_url()       — Construct the Xero authorization URL
  2. exchange_code()        — Exchange auth code for access + refresh tokens
  3. refresh_access_token() — Use stored refresh token to get a new access token
  4. get_valid_token()      — Return a valid access token (auto-refreshes if needed)
  5. start_auth_flow()      — Open browser + start local callback server (one-time setup)
  6. revoke_token()         — Revoke the stored refresh token

Credentials stored in the CoWorker SQLite settings DB (via config.py):
  xero_client_id        — MCS Mate app Client ID
  xero_client_secret    — MCS Mate app Client Secret 2
  xero_refresh_token    — Long-lived refresh token (persisted after first auth)
  xero_access_token     — Short-lived access token (cached, auto-refreshed)
  xero_token_expiry     — ISO timestamp when access token expires
  xero_tenant_id        — Xero organisation/tenant ID (set after first auth)

XPM API base URL: https://api.xero.com/practicemanager/3.1/
OAuth endpoints:
  Authorise: https://login.xero.com/identity/connect/authorize
  Token:     https://identity.xero.com/connect/token
  Revoke:    https://identity.xero.com/connect/revocation
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import secrets
import threading
import time
import urllib.parse
import urllib.request
import urllib.error
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from typing import Optional

from config import get_setting, set_setting

# ── Constants ─────────────────────────────────────────────────────────────────

XERO_AUTH_URL   = "https://login.xero.com/identity/connect/authorize"
XERO_TOKEN_URL  = "https://identity.xero.com/connect/token"
XERO_REVOKE_URL = "https://identity.xero.com/connect/revocation"
XERO_CONNECTIONS_URL  = "https://api.xero.com/connections"
XERO_USERINFO_URL     = "https://identity.xero.com/connect/userinfo"
XPM_STAFF_LIST_URL    = "https://api.xero.com/practicemanager/3.0/staff.api/list"

REDIRECT_URI    = "http://localhost:7842/oauth/callback"
# New granular Practice Manager scopes (replaced the old single 'practicemanager' scope)
SCOPES          = (
    "openid profile email offline_access "
    "practicemanager.client.read practicemanager.client "
    "practicemanager.job.read practicemanager.job "
    "practicemanager.staff.read "
    "practicemanager.time.read"
)

# ── Credentials helpers ───────────────────────────────────────────────────────

def get_client_id() -> str:
    # Env var takes priority so ops can supply credentials without touching the DB
    return os.environ.get("XERO_CLIENT_ID") or get_setting("xero_client_id", "")

def get_client_secret() -> str:
    return os.environ.get("XERO_CLIENT_SECRET") or get_setting("xero_client_secret", "")

def is_configured() -> bool:
    """Return True if client ID and secret are stored."""
    return bool(get_client_id() and get_client_secret())

def is_authorised() -> bool:
    """Return True if a refresh token is stored (user has completed OAuth)."""
    return bool(get_setting("xero_refresh_token", ""))

def get_tenant_id() -> str:
    return get_setting("xero_tenant_id", "")

# ── PKCE helpers ──────────────────────────────────────────────────────────────

def _generate_pkce() -> tuple[str, str]:
    """Generate a PKCE code_verifier and code_challenge (S256)."""
    code_verifier = secrets.token_urlsafe(64)
    digest = hashlib.sha256(code_verifier.encode()).digest()
    code_challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode()
    return code_verifier, code_challenge

# ── Authorization URL ─────────────────────────────────────────────────────────

def build_auth_url(state: str = "", code_verifier: str = "") -> tuple[str, str, str]:
    """
    Build the Xero authorization URL.

    Returns:
        (auth_url, state, code_verifier)
    """
    if not state:
        state = secrets.token_urlsafe(16)
    if not code_verifier:
        code_verifier, code_challenge = _generate_pkce()
    else:
        digest = hashlib.sha256(code_verifier.encode()).digest()
        code_challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode()

    params = {
        "response_type":         "code",
        "client_id":             get_client_id(),
        "redirect_uri":          REDIRECT_URI,
        "scope":                 SCOPES,
        "state":                 state,
        "code_challenge":        code_challenge,
        "code_challenge_method": "S256",
    }
    url = XERO_AUTH_URL + "?" + urllib.parse.urlencode(params)
    return url, state, code_verifier

# ── Token exchange ────────────────────────────────────────────────────────────

def exchange_code(auth_code: str, code_verifier: str) -> dict:
    """
    Exchange an authorization code for access + refresh tokens.

    Returns the full token response dict:
      access_token, refresh_token, expires_in, token_type, scope
    """
    client_id     = get_client_id()
    client_secret = get_client_secret()

    body = urllib.parse.urlencode({
        "grant_type":    "authorization_code",
        "code":          auth_code,
        "redirect_uri":  REDIRECT_URI,
        "code_verifier": code_verifier,
    }).encode()

    credentials = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    req = urllib.request.Request(
        XERO_TOKEN_URL,
        data=body,
        headers={
            "Authorization": f"Basic {credentials}",
            "Content-Type":  "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            token_data = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        error_body = e.read().decode() if e.fp else ""
        raise XeroOAuthError(f"Token exchange failed ({e.code}): {error_body}") from e

    _store_tokens(token_data)
    return token_data

# ── Token refresh ─────────────────────────────────────────────────────────────

def refresh_access_token() -> str:
    """
    Use the stored refresh token to obtain a new access token.
    Stores the new tokens and returns the new access_token string.
    Raises XeroOAuthError if no refresh token is stored or refresh fails.
    """
    refresh_token = get_setting("xero_refresh_token", "")
    if not refresh_token:
        raise XeroOAuthError(
            "No Xero refresh token stored. Please complete the OAuth setup in Settings."
        )

    client_id     = get_client_id()
    client_secret = get_client_secret()

    body = urllib.parse.urlencode({
        "grant_type":    "refresh_token",
        "refresh_token": refresh_token,
    }).encode()

    credentials = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    req = urllib.request.Request(
        XERO_TOKEN_URL,
        data=body,
        headers={
            "Authorization": f"Basic {credentials}",
            "Content-Type":  "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            token_data = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        error_body = e.read().decode() if e.fp else ""
        raise XeroOAuthError(f"Token refresh failed ({e.code}): {error_body}") from e

    _store_tokens(token_data)
    return token_data["access_token"]

# ── Get valid token (auto-refresh) ────────────────────────────────────────────

def get_valid_token() -> str:
    """
    Return a valid Xero access token.
    If the stored token has expired (or expires within 60 seconds), refresh it.
    Raises XeroOAuthError if not authorised or refresh fails.
    """
    if not is_authorised():
        raise XeroOAuthError(
            "Xero not authorised. Please complete OAuth setup in Settings → Xero."
        )

    access_token = get_setting("xero_access_token", "")
    expiry_str   = get_setting("xero_token_expiry", "")

    if access_token and expiry_str:
        try:
            expiry = datetime.fromisoformat(expiry_str)
            # Refresh if less than 60 seconds remain
            if expiry > datetime.now(timezone.utc).replace(tzinfo=None) + \
               __import__("datetime").timedelta(seconds=60):
                return access_token
        except ValueError:
            pass  # Bad expiry format — fall through to refresh

    return refresh_access_token()

# ── Tenant discovery ──────────────────────────────────────────────────────────

def fetch_tenant_id(access_token: str) -> str:
    """
    Query the Xero connections endpoint to find the tenant ID for the
    practice management (XPM) organisation.

    Stores the tenant ID in settings and returns it.
    """
    req = urllib.request.Request(
        XERO_CONNECTIONS_URL,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept":        "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            connections = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        error_body = e.read().decode() if e.fp else ""
        raise XeroOAuthError(f"Failed to fetch connections ({e.code}): {error_body}") from e

    if not connections:
        raise XeroOAuthError("No Xero connections found. Ensure the app is connected to an organisation.")

    # Prefer a PRACTICE_MANAGER connection; fall back to first connection
    tenant_id = None
    for conn in connections:
        if conn.get("tenantType", "").upper() in ("PRACTICE_MANAGER", "ORGANISATION"):
            tenant_id = conn["tenantId"]
            break
    if not tenant_id:
        tenant_id = connections[0]["tenantId"]

    set_setting("xero_tenant_id", tenant_id)
    return tenant_id

# ── Staff identity (per-accountant scoping) ──────────────────────────────────

def fetch_staff_identity(access_token: str, tenant_id: str) -> dict:
    """
    Identify the logged-in accountant's XPM staff record.

    Steps:
      1. Call Xero /connect/userinfo to get the authenticated user's email.
      2. Call XPM /staff.api/list to find the matching staff member.
      3. Store xero_staff_uuid and xero_staff_name in settings.

    Returns a dict with keys: email, uuid, name.
    Raises XeroOAuthError if the staff member cannot be matched.
    """
    # Step 1: Get the authenticated user's email from Xero userinfo
    req = urllib.request.Request(
        XERO_USERINFO_URL,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept":        "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            userinfo = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        error_body = e.read().decode() if e.fp else ""
        raise XeroOAuthError(f"Failed to fetch userinfo ({e.code}): {error_body}") from e

    user_email = (userinfo.get("email") or "").lower().strip()
    if not user_email:
        raise XeroOAuthError("Could not determine authenticated user email from Xero userinfo.")

    # Step 2: Fetch all XPM staff and match by email
    req2 = urllib.request.Request(
        XPM_STAFF_LIST_URL,
        headers={
            "Authorization":  f"Bearer {access_token}",
            "Xero-Tenant-Id": tenant_id,
            "Accept":         "application/xml",
        },
    )
    try:
        with urllib.request.urlopen(req2, timeout=30) as resp:
            xml_bytes = resp.read()
    except urllib.error.HTTPError as e:
        error_body = e.read().decode() if e.fp else ""
        raise XeroOAuthError(f"Failed to fetch XPM staff list ({e.code}): {error_body}") from e

    # Parse XML — use simple string matching to avoid lxml dependency
    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        raise XeroOAuthError(f"Failed to parse XPM staff XML: {e}") from e

    # Find the staff member whose email matches the authenticated user
    matched_uuid = ""
    matched_name = ""
    for staff_el in root.iter("Staff"):
        email_el = staff_el.find("Email")
        uuid_el  = staff_el.find("UUID")
        name_el  = staff_el.find("Name")
        if email_el is not None and email_el.text:
            if email_el.text.lower().strip() == user_email:
                matched_uuid = (uuid_el.text or "").strip() if uuid_el is not None else ""
                matched_name = (name_el.text or "").strip() if name_el is not None else ""
                break

    if not matched_uuid:
        # Soft failure: store the email so the user can see what was tried
        set_setting("xero_staff_uuid", "")
        set_setting("xero_staff_name", f"(not matched — {user_email})")
        raise XeroOAuthError(
            f"Could not find XPM staff record for email '{user_email}'. "
            "Ensure the accountant's Xero login email matches their XPM staff email."
        )

    set_setting("xero_staff_uuid", matched_uuid)
    set_setting("xero_staff_name", matched_name)
    return {"email": user_email, "uuid": matched_uuid, "name": matched_name}


def get_staff_uuid() -> str:
    """Return the stored XPM staff UUID for the connected accountant."""
    return get_setting("xero_staff_uuid", "")


def get_staff_name() -> str:
    """Return the stored XPM staff name for the connected accountant."""
    return get_setting("xero_staff_name", "")


# ── Revoke ────────────────────────────────────────────────────────────────────

def revoke_token() -> None:
    """Revoke the stored refresh token and clear all Xero token settings."""
    refresh_token = get_setting("xero_refresh_token", "")
    if not refresh_token:
        return

    client_id     = get_client_id()
    client_secret = get_client_secret()
    credentials   = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()

    body = urllib.parse.urlencode({
        "token":           refresh_token,
        "token_type_hint": "refresh_token",
    }).encode()

    req = urllib.request.Request(
        XERO_REVOKE_URL,
        data=body,
        headers={
            "Authorization": f"Basic {credentials}",
            "Content-Type":  "application/x-www-form-urlencoded",
        },
        method="POST",
    )
    try:
        urllib.request.urlopen(req, timeout=15)
    except Exception:
        pass  # Best-effort revocation

    _clear_tokens()

# ── One-time OAuth flow (desktop) ─────────────────────────────────────────────

# Shared state for the callback server
_oauth_result: dict = {}
_oauth_lock   = threading.Event()

class _CallbackHandler(BaseHTTPRequestHandler):
    """Minimal HTTP handler that captures the OAuth callback."""

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/oauth/callback":
            params = dict(urllib.parse.parse_qsl(parsed.query))
            _oauth_result.update(params)
            _oauth_lock.set()

            # Serve a simple success page
            html = (
                "<!DOCTYPE html>"
                "<html><head><title>MCS CoWorker - Xero Connected</title>"
                "<style>body{font-family:sans-serif;text-align:center;padding:60px;background:#f0f4f8}"
                "h1{color:#13b5ea}p{color:#444}</style></head>"
                "<body><h1>&#10003; Xero Connected!</h1>"
                "<p>MCS CoWorker is now connected to Xero XPM.</p>"
                "<p>You can close this window.</p></body></html>"
            ).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.send_header("Content-Length", str(len(html)))
            self.end_headers()
            self.wfile.write(html)
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        pass  # Suppress server logs


def start_auth_flow(timeout: int = 300) -> dict:
    """
    Complete the full OAuth Authorization Code flow:
      1. Start a temporary HTTP server on port 7842 (path /oauth/callback)
      2. Build the authorization URL and open it in the user's browser
      3. Wait for the callback (up to `timeout` seconds)
      4. Exchange the auth code for tokens
      5. Fetch and store the tenant ID

    Returns the token response dict on success.
    Raises XeroOAuthError on failure or timeout.

    NOTE: This function is called from the Flask api_server.py route
    /api/xero/start-auth — it runs in a background thread so the Flask
    server can still serve the callback on the same port.
    """
    if not is_configured():
        raise XeroOAuthError(
            "Xero Client ID and Secret not configured. Add them in Settings → Xero."
        )

    # Reset shared state
    _oauth_result.clear()
    _oauth_lock.clear()

    # Generate PKCE and state
    code_verifier, code_challenge = _generate_pkce()
    state = secrets.token_urlsafe(16)

    # Store verifier for use in exchange_code after callback
    set_setting("_xero_pkce_verifier", code_verifier)
    set_setting("_xero_oauth_state",   state)

    # Build auth URL
    digest = hashlib.sha256(code_verifier.encode()).digest()
    code_challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode()

    params = {
        "response_type":         "code",
        "client_id":             get_client_id(),
        "redirect_uri":          REDIRECT_URI,
        "scope":                 SCOPES,
        "state":                 state,
        "code_challenge":        code_challenge,
        "code_challenge_method": "S256",
    }
    auth_url = XERO_AUTH_URL + "?" + urllib.parse.urlencode(params)

    # Open browser
    webbrowser.open(auth_url)

    # Wait for callback (the Flask server handles the actual HTTP request)
    if not _oauth_lock.wait(timeout=timeout):
        raise XeroOAuthError("OAuth flow timed out. Please try again.")

    # Validate state
    returned_state = _oauth_result.get("state", "")
    if returned_state != state:
        raise XeroOAuthError("OAuth state mismatch — possible CSRF attack.")

    error = _oauth_result.get("error")
    if error:
        raise XeroOAuthError(f"Xero OAuth error: {error} — {_oauth_result.get('error_description', '')}")

    auth_code = _oauth_result.get("code", "")
    if not auth_code:
        raise XeroOAuthError("No authorization code received from Xero.")

    # Exchange code for tokens
    token_data = exchange_code(auth_code, code_verifier)

    # Fetch tenant ID
    tenant_id = ""
    try:
        tenant_id = fetch_tenant_id(token_data["access_token"])
    except Exception:
        pass  # Non-fatal — tenant ID can be set manually

    # Resolve the accountant's XPM staff identity (for per-accountant scoping)
    if tenant_id:
        try:
            fetch_staff_identity(token_data["access_token"], tenant_id)
        except XeroOAuthError:
            pass  # Non-fatal — scoping will be disabled but connection still works
        except Exception:
            pass

    # Clean up temp settings
    set_setting("_xero_pkce_verifier", "")
    set_setting("_xero_oauth_state",   "")

    return token_data


def notify_oauth_callback(code: str, state: str, error: str = "") -> None:
    """
    Called by the Flask /oauth/callback route to pass the auth code
    back to the waiting start_auth_flow() thread.
    """
    _oauth_result["code"]  = code
    _oauth_result["state"] = state
    if error:
        _oauth_result["error"] = error
    _oauth_lock.set()

# ── Internal helpers ──────────────────────────────────────────────────────────

def _store_tokens(token_data: dict) -> None:
    """Persist access token, refresh token, and expiry to settings DB."""
    access_token  = token_data.get("access_token", "")
    refresh_token = token_data.get("refresh_token", "")
    expires_in    = int(token_data.get("expires_in", 1800))

    expiry = datetime.utcnow().replace(tzinfo=None) + \
             __import__("datetime").timedelta(seconds=expires_in)

    if access_token:
        set_setting("xero_access_token",  access_token)
    if refresh_token:
        set_setting("xero_refresh_token", refresh_token)
    set_setting("xero_token_expiry", expiry.isoformat())


def _clear_tokens() -> None:
    """Clear all stored Xero tokens and staff identity."""
    for key in ("xero_access_token", "xero_refresh_token",
                "xero_token_expiry", "xero_tenant_id",
                "xero_staff_uuid", "xero_staff_name"):
        set_setting(key, "")


# ── Exception ─────────────────────────────────────────────────────────────────

class XeroOAuthError(Exception):
    """Raised when the Xero OAuth flow fails."""
