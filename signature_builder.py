"""
MC & S CoWorker — Dynamic Email Signature Builder
=================================================
Builds the per-accountant HTML signature appended to every plugin-created
draft. Resolves the signed-in M365 user against the `staff_signatures` table
to pick the right name + title; reads firm constants (logo, phone, address,
social URLs, privacy text) from the `settings` table; embeds the bundled
logo and social icons inline as base64 PNGs.

See docs/planning/dynamic_signatures_design.md for the design rationale.

Public API:
    build_signature_html(user_email: str | None) -> str
"""
from __future__ import annotations

import base64
import html as _html
import logging
import threading
from pathlib import Path
from typing import Callable, Optional

logger = logging.getLogger(__name__)

# Resolve the bundled image directory the same way main.py picks BASE_DIR — the
# embeddable-Python install lays out app/ at <install>/app, so assets/signature
# sits next to this file at app/assets/signature.
_THIS_DIR = Path(__file__).parent
_ASSET_DIR = _THIS_DIR / "assets" / "signature"

_LOGO_FILE = _ASSET_DIR / "logo.png"
_INSTAGRAM_FILE = _ASSET_DIR / "instagram.png"
_FACEBOOK_FILE = _ASSET_DIR / "facebook.png"

# In-memory image cache. Each entry is the full `data:image/png;base64,...`
# URI string ready to drop into an <img src=...>. Loaded on first call and
# kept for the process lifetime — accountants don't restart often, and the
# files don't change without a redeploy.
_image_cache: dict[str, str] = {}
_image_cache_lock = threading.Lock()
_image_cache_loaded = False

# Track which "fall back to legacy" warnings we've already logged this process
# so a busy plugin run doesn't spam the log on every draft.
_warned_keys: set[str] = set()


def _warn_once(key: str, message: str) -> None:
    if key in _warned_keys:
        return
    _warned_keys.add(key)
    logger.warning("Signature: %s", message)


def _load_image_data_uri(path: Path) -> str:
    """Read a PNG and return its data URI, or empty string on failure."""
    try:
        if not path.is_file():
            _warn_once(
                f"missing:{path.name}",
                f"asset {path} not found, signature will render without it",
            )
            return ""
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        return f"data:image/png;base64,{b64}"
    except Exception as e:
        _warn_once(
            f"err:{path.name}", f"failed to load {path}: {e}"
        )
        return ""


def _ensure_images_loaded() -> None:
    global _image_cache_loaded
    if _image_cache_loaded:
        return
    with _image_cache_lock:
        if _image_cache_loaded:
            return
        _image_cache["logo"] = _load_image_data_uri(_LOGO_FILE)
        _image_cache["instagram"] = _load_image_data_uri(_INSTAGRAM_FILE)
        _image_cache["facebook"] = _load_image_data_uri(_FACEBOOK_FILE)
        _image_cache_loaded = True


def reset_image_cache() -> None:
    """Force the next call to re-read PNGs from disk. Used by tests."""
    global _image_cache_loaded
    with _image_cache_lock:
        _image_cache.clear()
        _image_cache_loaded = False
        _warned_keys.clear()


# ── Settings + staff lookup ───────────────────────────────────────────────────

def _get_setting(key: str, default: str = "") -> str:
    # Imported lazily so signature_builder stays importable in tests that don't
    # initialise the DB until after import time.
    from config import get_setting
    return (get_setting(key, default) or "").strip()


def _lookup_staff(user_email: Optional[str]) -> Optional[dict]:
    if not user_email:
        return None
    from config import get_staff_signature_by_email
    return get_staff_signature_by_email(user_email)


# ── Legacy fallback ───────────────────────────────────────────────────────────

def _invoke_legacy(legacy_fallback: Optional[Callable[[], str]]) -> str:
    """Call the supplied legacy-signature callable safely. Used when dynamic
    mode can't resolve a staff row — we hand control back to graph_client's
    existing image / text / auto-detect path so behaviour matches pre-change."""
    if legacy_fallback is None:
        return ""
    try:
        return legacy_fallback() or ""
    except Exception as e:
        logger.warning("Legacy signature fallback failed: %s", e)
        return ""


# ── Template assembly ─────────────────────────────────────────────────────────

# Brand link colour (used for hyperlinks in the signature). Picked to match
# the navy in the firm logo without requiring a settings key — Elio can change
# it later by editing this constant and it'll flow through every signature.
_LINK_COLOUR = "#1A4A6E"


def _esc(text: str) -> str:
    """HTML-escape free text. Safe to call on settings values that may contain
    user-typed punctuation."""
    return _html.escape(text or "", quote=True)


def _build_html(staff: dict, settings: dict, images: dict) -> str:
    """Render the full HTML signature.

    Args:
        staff:    dict with name, title (may be empty), email
        settings: dict of resolved firm-constant strings (already escaped where
                  they're inserted as text; URLs and href values use the
                  attribute-quote-safe _esc as well)
        images:   dict of data-URI strings (logo, instagram, facebook)
    """
    name = _esc(staff["name"])
    title = (staff.get("title") or "").strip()

    company = _esc(settings["company"])
    phone = _esc(settings["phone"])
    website_display = _esc(settings["website_display"])
    website_url = _esc(settings["website_url"])
    address1 = _esc(settings["address1"])
    address2 = _esc(settings["address2"])
    instagram_url = _esc(settings["instagram_url"])
    facebook_url = _esc(settings["facebook_url"])
    linkedin_url = _esc(settings["linkedin_url"])
    google_review_url = _esc(settings["google_review_url"])
    privacy_text = _esc(settings["privacy_text"])

    # ── Logo cell (skipped if asset missing) ──────────────────────────────
    if images.get("logo"):
        logo_cell = (
            '<td style="padding-right:16px;vertical-align:top;">'
            f'<img src="{images["logo"]}" width="80" height="80" alt="MC&amp;S" '
            'style="display:block;border:0;" />'
            '</td>'
        )
    else:
        logo_cell = ""

    # ── Title line — only if non-empty ────────────────────────────────────
    if title:
        title_line = (
            f'<div style="font-size:9pt;color:#444444;">{_esc(title)}</div>'
        )
    else:
        title_line = ""

    # ── Address lines — skip blanks ───────────────────────────────────────
    address_lines = []
    if address1:
        address_lines.append(
            f'<div style="font-size:9pt;color:#555555;">{address1}</div>'
        )
    if address2:
        address_lines.append(
            f'<div style="font-size:9pt;color:#555555;">{address2}</div>'
        )
    address_block = "".join(address_lines)

    # ── Phone | Website line ──────────────────────────────────────────────
    phone_web_parts = []
    if phone:
        phone_web_parts.append(phone)
    if website_url and website_display:
        phone_web_parts.append(
            f'<a href="{website_url}" '
            f'style="color:{_LINK_COLOUR};text-decoration:none;">'
            f'{website_display}</a>'
        )
    phone_web_line = ""
    if phone_web_parts:
        phone_web_line = (
            '<div>'
            + '&nbsp;&nbsp;|&nbsp;&nbsp;'.join(phone_web_parts)
            + '</div>'
        )

    # ── Social row ────────────────────────────────────────────────────────
    social_icons: list[str] = []
    if instagram_url and images.get("instagram"):
        social_icons.append(
            f'<a href="{instagram_url}" style="text-decoration:none;">'
            f'<img src="{images["instagram"]}" width="20" height="20" '
            f'alt="Instagram" style="display:inline-block;border:0;'
            f'margin-right:4px;" /></a>'
        )
    if facebook_url and images.get("facebook"):
        social_icons.append(
            f'<a href="{facebook_url}" style="text-decoration:none;">'
            f'<img src="{images["facebook"]}" width="20" height="20" '
            f'alt="Facebook" style="display:inline-block;border:0;'
            f'margin-right:4px;" /></a>'
        )
    # LinkedIn intentionally URL-only (no bundled icon yet) — out of v1 per
    # design. If the URL is set we still skip rendering until an icon ships.

    if social_icons:
        social_row = (
            '<div style="padding-top:6px;">'
            + "".join(social_icons)
            + '</div>'
        )
    else:
        social_row = ""

    # ── Google review CTA ────────────────────────────────────────────────
    if google_review_url:
        review_row = (
            '<div style="padding-top:8px;font-size:10pt;">'
            f'Love what we do? <a href="{google_review_url}" '
            f'style="color:{_LINK_COLOUR};">Leave us a Google review here</a>'
            '</div>'
        )
    else:
        review_row = ""

    # ── Details cell ─────────────────────────────────────────────────────
    # All firm-constant strings were escaped at the top of this function;
    # don't re-escape here or "&amp;" turns into "&amp;amp;".
    details_cell = (
        '<td style="vertical-align:top;">'
        f'<div style="font-size:11pt;font-weight:bold;">{name}</div>'
        f'{title_line}'
        f'<div>{company}</div>'
        f'{phone_web_line}'
        f'{address_block}'
        f'{social_row}'
        f'{review_row}'
        '</td>'
    )

    # ── Privacy block ────────────────────────────────────────────────────
    if privacy_text:
        privacy_block = (
            '<div style="margin-top:12px;'
            'font-family:Aptos,Calibri,sans-serif;font-size:8pt;'
            'color:#666666;font-style:italic;line-height:1.4;">'
            f'{privacy_text}'
            '</div>'
        )
    else:
        privacy_block = ""

    return (
        '<br><br>'
        '<table cellpadding="0" cellspacing="0" border="0" '
        'style="font-family:Aptos,Calibri,sans-serif;font-size:10pt;'
        'color:#000000;border-collapse:collapse;">'
        f'<tr>{logo_cell}{details_cell}</tr>'
        '</table>'
        f'{privacy_block}'
    )


def _resolve_settings() -> dict:
    """Read all firm-constant settings into a dict ready for template insertion.

    Empty/missing settings just produce empty strings — the renderer skips
    blank rows.
    """
    return {
        "company":           _get_setting("signature_company", "MC&S Pty Ltd"),
        "phone":             _get_setting("signature_phone", ""),
        "website_display":   _get_setting("signature_website_display", ""),
        "website_url":       _get_setting("signature_website_url", ""),
        "address1":          _get_setting("signature_address_line1", ""),
        "address2":          _get_setting("signature_address_line2", ""),
        "instagram_url":     _get_setting("signature_instagram_url", ""),
        "facebook_url":      _get_setting("signature_facebook_url", ""),
        "linkedin_url":      _get_setting("signature_linkedin_url", ""),
        "google_review_url": _get_setting("signature_google_review_url", ""),
        "privacy_text":      _get_setting("signature_privacy_text", ""),
    }


# ── Public API ────────────────────────────────────────────────────────────────

def build_signature_html(
    user_email: Optional[str],
    legacy_fallback: Optional[Callable[[], str]] = None,
) -> str:
    """Build the HTML signature for the given M365 user email.

    Resolution order:
      1. signature_mode = 'disabled' -> return empty string (no signature).
      2. signature_mode = 'legacy_image' -> defer to legacy_fallback.
      3. signature_mode = 'dynamic' (default):
           - lookup staff row by email -> render dynamic HTML
           - no match / no email / disabled row -> legacy_fallback
      4. Any unexpected exception -> legacy_fallback (logged once).

    The `legacy_fallback` callable is supplied by graph_client and points at
    its existing get_signature_html() — kept as a callable so this module
    doesn't need to instantiate GraphClient (which would require an MSAL
    token + network access).
    """
    try:
        mode = _get_setting("signature_mode", "dynamic").lower() or "dynamic"

        if mode == "disabled":
            return ""
        if mode == "legacy_image":
            return _invoke_legacy(legacy_fallback)

        # Dynamic mode below.
        normalised_email = (user_email or "").strip().lower()
        staff = _lookup_staff(normalised_email)
        if not staff:
            _warn_once(
                f"nomatch:{normalised_email or 'no-email'}",
                f"M365 user {normalised_email or '(none)'} not found in "
                "staff_signatures, using legacy image.",
            )
            return _invoke_legacy(legacy_fallback)

        # Per-user opt-out. The accountant has explicitly turned the standard
        # signature off for their drafts — return empty rather than falling
        # back to the legacy image, otherwise the toggle would do nothing.
        if not staff.get("include_signature", 1):
            return ""

        _ensure_images_loaded()
        settings = _resolve_settings()
        return _build_html(staff, settings, _image_cache)

    except Exception as e:
        # Never let signature rendering break a draft — fall back silently
        # after a single warning.
        _warn_once("render-error", f"dynamic signature failed ({e}), using legacy")
        return _invoke_legacy(legacy_fallback)
