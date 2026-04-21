"""Transparent encryption for sensitive settings using Windows DPAPI.

Stored values are prefixed with ``DPAPI:`` and the payload is base64-encoded
ciphertext. Values without the prefix are treated as legacy plaintext and
returned unchanged — they get re-encrypted the next time ``save_setting`` is
called for that key, so no migration is needed.

If ``win32crypt`` isn't available (non-Windows, missing pywin32), the helpers
fall through to plaintext storage and log a warning. That keeps the app
functional on dev machines without hiding the risk.
"""
from __future__ import annotations

import base64
import logging

logger = logging.getLogger(__name__)

SENSITIVE_KEYS = {
    "anthropic_api_key",
    "fusesign_api_key",
    "statementhub_api_key",
    "teams_webhook_url",
    "xero_client_secret",
    "xero_access_token",
    "xero_refresh_token",
    "local_api_token",
}

DPAPI_PREFIX = "DPAPI:"

try:
    import win32crypt  # type: ignore
    _HAS_DPAPI = True
except ImportError:
    _HAS_DPAPI = False
    logger.warning("win32crypt not available — sensitive settings will be stored in plaintext")


def encrypt_value(plaintext: str) -> str:
    if not _HAS_DPAPI or not plaintext:
        return plaintext
    try:
        encrypted = win32crypt.CryptProtectData(
            plaintext.encode("utf-8"), "MCSCoWorker", None, None, None, 0
        )
        return DPAPI_PREFIX + base64.b64encode(encrypted).decode("ascii")
    except Exception as e:
        logger.error(f"DPAPI encrypt failed: {e}")
        return plaintext


def decrypt_value(stored: str) -> str:
    if not stored or not stored.startswith(DPAPI_PREFIX):
        return stored  # legacy plaintext or empty
    if not _HAS_DPAPI:
        logger.error("Cannot decrypt DPAPI value — win32crypt not available")
        return ""
    try:
        raw = base64.b64decode(stored[len(DPAPI_PREFIX):])
        _, plaintext = win32crypt.CryptUnprotectData(raw, None, None, None, 0)
        return plaintext.decode("utf-8")
    except Exception as e:
        logger.error(f"DPAPI decrypt failed: {e}")
        return ""
