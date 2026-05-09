"""
MC & S Desktop Agent - Configuration & Database Manager
"""
import sqlite3
import json
import logging
import os
import secrets
from pathlib import Path

from crypto_utils import SENSITIVE_KEYS, encrypt_value, decrypt_value

logger = logging.getLogger(__name__)

# Columns that save_plugin_state is allowed to write. Callers that pass any
# other key get dropped before the SET clause is built — this prevents a
# crafted kwarg like ``"; DROP TABLE..."`` from ever reaching the SQL string.
ALLOWED_PLUGIN_COLUMNS = {
    "enabled", "draft_mode", "schedule_seconds",
    "last_run", "last_result", "last_summary", "display_name",
}

def _get_data_dir() -> Path:
    """Return the app data directory, respecting ``MCS_DATA_DIR`` if set.

    ``launcher.py`` points MCS_DATA_DIR at ``%INSTALL_DIR%\\data\\`` on
    production installs. Falling back to ``~/.mcs_email_automation`` keeps dev
    machines working, but OneDrive/roaming-profile syncing of ``~`` is the
    reason the env-var path is strongly preferred.
    """
    d = os.environ.get("MCS_DATA_DIR")
    p = Path(d) if d else Path.home() / ".mcs_email_automation"
    p.mkdir(parents=True, exist_ok=True)
    return p


DATA_DIR = _get_data_dir()
DB_PATH = DATA_DIR / "config.db"


def apply_wal_pragmas(conn: sqlite3.Connection) -> None:
    """Enable WAL mode + busy_timeout on a SQLite connection.

    Shared helper for every module that opens its own SQLite connection
    (auto_updater, approval_queue, kpi_monitor, token_meter, etc.) so each
    database file gets consistent concurrency behaviour.

    WAL lets readers and writers coexist without blocking each other.
    busy_timeout (5s) means concurrent waiters retry instead of immediately
    raising 'database is locked'.
    """
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=5000")
    except sqlite3.Error:
        # Some platforms / in-memory DBs may reject these — not fatal.
        pass


def get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    apply_wal_pragmas(conn)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create all tables and seed defaults if first run."""
    conn = get_db()
    c = conn.cursor()

    c.executescript("""
        CREATE TABLE IF NOT EXISTS settings (
            key   TEXT PRIMARY KEY,
            value TEXT
        );

        CREATE TABLE IF NOT EXISTS plugin_registry (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            plugin_id        TEXT UNIQUE NOT NULL,
            enabled          INTEGER DEFAULT 1,
            draft_mode       INTEGER DEFAULT 1,
            schedule_seconds INTEGER DEFAULT 0,
            last_run         TEXT,
            last_result      TEXT,
            last_summary     TEXT,
            display_name     TEXT
        );

        CREATE TABLE IF NOT EXISTS email_rules (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            category         TEXT NOT NULL,
            keywords         TEXT NOT NULL,
            subject_template TEXT,
            body_template    TEXT,
            enabled          INTEGER DEFAULT 1,
            sort_order       INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS staff_notifications (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            name            TEXT NOT NULL,
            email           TEXT NOT NULL,
            receives_drafts INTEGER DEFAULT 1,
            enabled         INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS activity_log (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp         TEXT DEFAULT (datetime('now','localtime')),
            from_email        TEXT,
            subject           TEXT,
            classification    TEXT,
            action            TEXT,
            draft_created     INTEGER DEFAULT 0,
            notification_sent INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS links_forms (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            name      TEXT NOT NULL,
            tag       TEXT NOT NULL UNIQUE,
            url       TEXT NOT NULL DEFAULT '',
            enabled   INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS plugin_templates (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            plugin_id       TEXT NOT NULL,
            template_key    TEXT NOT NULL,
            template_label  TEXT,
            template_value  TEXT,
            template_type   TEXT DEFAULT 'textarea',
            UNIQUE(plugin_id, template_key)
        );

        CREATE TABLE IF NOT EXISTS memory_style (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            content   TEXT NOT NULL,
            updated   TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS memory_feedback (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT DEFAULT (datetime('now','localtime')),
            role      TEXT NOT NULL,
            message   TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS memory_lessons (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT DEFAULT (datetime('now','localtime')),
            lesson    TEXT NOT NULL,
            source    TEXT,
            active    INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS knowledge_base (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            category   TEXT NOT NULL,
            title      TEXT NOT NULL,
            content    TEXT NOT NULL,
            enabled    INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS bas_clients (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name        TEXT NOT NULL,
            entity_name        TEXT,
            abn                TEXT,
            frequency          TEXT DEFAULT 'Quarterly',
            client_email       TEXT,
            last_data_received TEXT,
            last_reminder_sent TEXT,
            status             TEXT DEFAULT 'Awaiting Data',
            notes              TEXT,
            created_at         TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS smart_responder_processed (
            message_id   TEXT PRIMARY KEY,
            processed_at TEXT DEFAULT CURRENT_TIMESTAMP,
            draft_id     TEXT,
            action       TEXT
        );

        -- Pending SharePoint uploads that hit a folder governance error
        -- (missing or ambiguous client folder). Recorded by the
        -- smart_responder catch-and-record path; drained by the morning
        -- briefing's pre-brief retry pass.
        --
        -- candidate_names_json is a JSON-encoded list of the verbatim
        -- SharePoint folder names that collided (for ambiguous), or NULL
        -- (for missing).
        --
        -- UNIQUE(message_id, action) prevents queue bloat when the same
        -- email re-runs through smart_responder before resolution; second
        -- and subsequent INSERTs use OR IGNORE.
        CREATE TABLE IF NOT EXISTS sharepoint_upload_pending (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            message_id           TEXT NOT NULL,
            action               TEXT NOT NULL DEFAULT 'upload_pending',
            client_name          TEXT NOT NULL,
            candidate_names_json TEXT,
            created_at           TEXT DEFAULT CURRENT_TIMESTAMP,
            resolved_at          TEXT,
            UNIQUE(message_id, action)
        );

        CREATE TABLE IF NOT EXISTS staff_signatures (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            name              TEXT NOT NULL,
            title             TEXT,
            email             TEXT NOT NULL UNIQUE,
            enabled           INTEGER NOT NULL DEFAULT 1,
            include_signature INTEGER NOT NULL DEFAULT 1,
            created_at        REAL NOT NULL DEFAULT (strftime('%s','now')),
            updated_at        REAL NOT NULL DEFAULT (strftime('%s','now'))
        );

        CREATE INDEX IF NOT EXISTS idx_staff_signatures_email
            ON staff_signatures(LOWER(email));
    """)

    # Add display_name column if upgrading from older schema
    try:
        c.execute("ALTER TABLE plugin_registry ADD COLUMN display_name TEXT")
    except sqlite3.OperationalError:
        pass  # Column already exists

    # Per-staff signature toggle. Default 1 preserves the prior behaviour for
    # rows that existed before this column was introduced.
    try:
        c.execute(
            "ALTER TABLE staff_signatures "
            "ADD COLUMN include_signature INTEGER NOT NULL DEFAULT 1"
        )
    except sqlite3.OperationalError:
        pass  # Column already exists

    # Seed default settings
    defaults = {
        "claude_model":           "claude-haiku-4-5-20251001",
        "claude_model_fast":      "claude-haiku-4-5-20251001",
        "claude_model_reasoning": "claude-sonnet-4-5",
        "draft_mode":             "1",
        "business_hours_enabled": "1",
        "business_hours_start":   "8",
        "business_hours_end":     "18",
        "business_days":          "1,2,3,4,5",
        "polling_interval":       "60",
        "ms_account_email":       "",
        "monitor_folder":         "Inbox",
        "practice_name":          "MC & S",
        "practice_email":         "",
        "timezone":               "AUS Eastern Standard Time",
        "setup_complete":         "0",
        "user_name":              "",
        "user_firm":              "",
        "user_email":             "",
        "user_setup_complete":    "0",
        # Stream 4 — Gateway integrations (Xero OAuth)
        # Credentials are NOT hardcoded here — populate via environment
        # variables (XERO_CLIENT_ID / XERO_CLIENT_SECRET) or the Settings UI.
        "xero_client_id":         "",
        "xero_client_secret":     "",
        "xero_refresh_token":     "",
        "xero_access_token":      "",
        "xero_token_expiry":      "",
        "xero_tenant_id":         "",
        "xero_staff_uuid":        "",
        "xero_staff_name":        "",
        "xpm_base_url":           "https://api.xero.com/practicemanager/3.1",
        "fusesign_api_key":       "",
        "fusesign_base_url":      "https://api.fusesign.com/v1",
        "teams_webhook_url":      "",
        "teams_graph_token":      "",
        "statementhub_api_key":   "",
        "statementhub_base_url":  "https://api.statementhub.com.au/v1",
        # Business hours
        "skip_public_holidays":   "1",
        "public_holiday_state":   "VIC",
        # Installation mode
        "reception_mode":         "0",   # 1 = reception inbox (NOA/ASIC/Debtor/FuseSign active)
        "staff_profile":          "elio", # reception | elio | ross | harry | brooke | louise | lyn
        # Per-machine processing mode — 'active' runs plugins, 'passive' is monitor-only
        "processing_mode":        "active",
        # Stream 3 — Heartbeat interval
        "heartbeat_interval_seconds": "60",
        # Stream 6 — Approval queue
        "approval_auto_threshold": "0.75",
        "approval_expiry_hours":   "48",
        # Draft formatting — applied to every plugin-created Outlook draft body
        "draft_font_family":       "Aptos, Calibri, sans-serif",
        "draft_font_size":         "11pt",
        "draft_font_color":        "#000000",
        # Dynamic per-accountant signature — see docs/planning/dynamic_signatures_design.md
        "signature_mode":          "dynamic",
        "signature_company":       "MC&S Pty Ltd",
        "signature_phone":         "(03) 9794 0000",
        "signature_website_display": "mcands.com.au",
        "signature_website_url":   "https://www.mcands.com.au",
        "signature_address_line1": "23 Timor Circuit, Keysborough, Vic 3173",
        "signature_address_line2": "PO BOX 4440, Dandenong South, VIC, 3164",
        "signature_instagram_url": "https://www.instagram.com/mcsaccounting",
        "signature_facebook_url":  "https://www.facebook.com/mcandsaccounting",
        "signature_linkedin_url":  "",
        "signature_google_review_url": "https://www.google.com/maps/place//data=!4m3!3m2!1s0x6ad6140f023d542b:0x4be4c0f80d96d34b!12e1?source=g.page.m.ia._&laa=nmx-review-solicitation-ia2",
        "signature_privacy_text":  (
            "This email and any attachments are confidential and may be subject "
            "to copyright, legal or some other professional privilege. They are "
            "intended solely for the attention and use of the named addressee(s). "
            "They may only be copied, distributed or disclosed with the consent "
            "of the copyright owner. If you have received this email by mistake "
            "or by breach of the confidentiality clause, please notify the sender "
            "immediately by return email and delete or destroy all copies of the "
            "email. Any confidentiality, privilege or copyright is not waived or "
            "lost because this email has been sent to you by mistake. Liability "
            "limited by a scheme approved under Professional Standards Legislation."
        ),
        "signature_seeded":        "0",
    }
    for key, value in defaults.items():
        c.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            (key, value),
        )

    # No default email rules — accountant builds their own from scratch

    # Seed default links & forms
    existing_links = c.execute("SELECT COUNT(*) FROM links_forms").fetchone()[0]
    if existing_links == 0:
        default_links = [
            ("Tax Return Checklist", "checklist_form", ""),
        ]
        c.executemany(
            "INSERT INTO links_forms (name, tag, url) VALUES (?, ?, ?)",
            default_links,
        )

    # Seed starter knowledge base entries — accountant fills in the content
    existing_kb = c.execute("SELECT COUNT(*) FROM knowledge_base").fetchone()[0]
    if existing_kb == 0:
        starter_kb = [
            ("Pricing",    "Standard Fees",              ""),
            ("Checklists", "Tax Return Checklist Link",  ""),
            ("Procedures", "BAS Lodgement Policy",       ""),
        ]
        c.executemany(
            "INSERT INTO knowledge_base (category, title, content) VALUES (?, ?, ?)",
            starter_kb,
        )

    # Pre-seed staff_signatures on first start of dynamic mode. Gated on the
    # one-shot setting so reinstalls / db restores don't clobber edits the
    # accountant has made since.
    seeded_flag = c.execute(
        "SELECT value FROM settings WHERE key='signature_seeded'"
    ).fetchone()
    already_seeded = bool(seeded_flag and seeded_flag["value"] == "1")
    existing_sigs = c.execute(
        "SELECT COUNT(*) FROM staff_signatures"
    ).fetchone()[0]
    if not already_seeded and existing_sigs == 0:
        seed_rows = [
            ("Elio Scarton",   "CPA, Tax Agent",   "elio@mcands.com.au"),
            ("Vince Mercuri",  None,                "vince@mcands.com.au"),
            ("Angelo Covelli", None,                "angelo@mcands.com.au"),
            ("Ross Mercuri",   "CA",                "ross@mcands.com.au"),
            ("Eliza Lewis",    "Reception",         "reception@mcands.com.au"),
            ("Brooke Austin",  None,                "brooke@mcands.com.au"),
            ("Harry Gan",      "CPA, SMSF Auditor", "harry@mcands.com.au"),
            ("Lyn Karman",     None,                "lyn@mcands.com.au"),
            ("Louise Boyd",    None,                "louise@mcands.com.au"),
        ]
        c.executemany(
            "INSERT OR IGNORE INTO staff_signatures (name, title, email) "
            "VALUES (?, ?, ?)",
            seed_rows,
        )
        c.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES ('signature_seeded', '1')"
        )

    conn.commit()
    conn.close()


# ── Settings ──────────────────────────────────────────────────────────────────

def get_setting(key, default=""):
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    if not row:
        return default
    value = row["value"]
    # Transparently decrypt sensitive keys. Legacy plaintext values (no DPAPI:
    # prefix) pass through unchanged and get re-encrypted on next save.
    if key in SENSITIVE_KEYS and value is not None:
        value = decrypt_value(value)
    return value


def set_setting(key, value):
    stored = str(value)
    if key in SENSITIVE_KEYS:
        stored = encrypt_value(stored)
    conn = get_db()
    conn.execute(
        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
        (key, stored),
    )
    conn.commit()
    conn.close()


def save_setting(key, value):
    """Alias for set_setting — kept for naming consistency with get_setting."""
    set_setting(key, value)


def is_active_mode() -> bool:
    """Check if this machine is in active processing mode.

    Active machines run plugins and create drafts. Passive machines are
    monitor-only and skip all plugin execution so two installs sharing the
    same Outlook mailbox don't draft duplicate replies.
    """
    return get_setting("processing_mode", "active").lower() == "active"


def ensure_api_token() -> str:
    """Generate and persist a per-install API token on first run.

    The token gates all /api/* routes so browser tabs, other apps, or malware
    cannot silently drive the Flask server running on 127.0.0.1:7842.
    """
    token = get_setting("local_api_token", "")
    if not token:
        token = secrets.token_urlsafe(48)
        set_setting("local_api_token", token)
    return token


def get_claude_model() -> str:
    """Return the fast (Haiku) Claude model string from settings.
    Kept for backward compatibility — equivalent to get_claude_model_fast().
    """
    return get_setting("claude_model_fast",
                       get_setting("claude_model", "claude-haiku-4-5-20251001"))


def get_claude_model_fast() -> str:
    """Return the fast Claude model (Haiku) for triage, drafting, and classification."""
    return get_setting("claude_model_fast",
                       get_setting("claude_model", "claude-haiku-4-5-20251001"))


def get_claude_model_reasoning() -> str:
    """Return the reasoning Claude model (Sonnet) for complex analysis and decisions."""
    return get_setting("claude_model_reasoning", "claude-sonnet-4-5")


def update_claude_models(api_key: str) -> dict:
    """
    Query the Anthropic models API and update all three stored model slots:
      - claude_model_fast      → newest available Haiku model
      - claude_model_reasoning → newest available Sonnet model
      - opus_model             → newest available Opus model

    Also keeps the legacy 'claude_model' key in sync with the fast model
    for backward compatibility with older plugins.

    Returns a dict with keys 'fast', 'reasoning', and 'opus'.
    Falls back to currently stored values on any error.
    """
    import urllib.request
    fast_current = get_claude_model_fast()
    reasoning_current = get_claude_model_reasoning()
    opus_current = get_setting("opus_model", "claude-opus-4-6")
    result = {"fast": fast_current, "reasoning": reasoning_current, "opus": opus_current}
    try:
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/models",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode())

        all_models = [
            m for m in data.get("data", []) if m.get("type") == "model"
        ]

        # ── Fast model: newest Haiku ──────────────────────────────────────────
        haiku_models = [
            m for m in all_models if "haiku" in m.get("id", "").lower()
        ]
        if haiku_models:
            haiku_models.sort(key=lambda m: m.get("created_at", ""), reverse=True)
            newest_haiku = haiku_models[0]["id"]
            set_setting("claude_model_fast", newest_haiku)
            set_setting("claude_model", newest_haiku)  # backward compat
            result["fast"] = newest_haiku

        # ── Reasoning model: newest Sonnet ────────────────────────────────────
        sonnet_models = [
            m for m in all_models if "sonnet" in m.get("id", "").lower()
        ]
        if sonnet_models:
            sonnet_models.sort(key=lambda m: m.get("created_at", ""), reverse=True)
            newest_sonnet = sonnet_models[0]["id"]
            set_setting("claude_model_reasoning", newest_sonnet)
            result["reasoning"] = newest_sonnet

        # ── Opus model: newest Opus ───────────────────────────────────────────
        opus_models = [
            m for m in all_models if "opus" in m.get("id", "").lower()
        ]
        if opus_models:
            opus_models.sort(key=lambda m: m.get("created_at", ""), reverse=True)
            newest_opus = opus_models[0]["id"]
            set_setting("opus_model", newest_opus)
            result["opus"] = newest_opus

    except Exception:
        pass  # return current values on any error

    return result


def update_claude_model(api_key: str) -> str:
    """
    Legacy single-model update — kept for backward compatibility.
    Calls update_claude_models() and returns the fast model ID.
    """
    return update_claude_models(api_key)["fast"]


def get_all_settings():
    conn = get_db()
    rows = conn.execute("SELECT key, value FROM settings").fetchall()
    conn.close()
    result = {}
    for r in rows:
        k, v = r["key"], r["value"]
        if k in SENSITIVE_KEYS and v is not None:
            v = decrypt_value(v)
        result[k] = v
    return result


# ── Email Rules ───────────────────────────────────────────────────────────────

def get_rules():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM email_rules ORDER BY sort_order, id"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_rule(rule: dict):
    conn = get_db()
    if rule.get("id"):
        conn.execute(
            """UPDATE email_rules
               SET category=?, keywords=?, subject_template=?,
                   body_template=?, enabled=?, sort_order=?
               WHERE id=?""",
            (
                rule["category"], rule["keywords"], rule["subject_template"],
                rule["body_template"], rule["enabled"], rule["sort_order"],
                rule["id"],
            ),
        )
    else:
        conn.execute(
            """INSERT INTO email_rules
               (category, keywords, subject_template, body_template, enabled, sort_order)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                rule["category"], rule["keywords"], rule["subject_template"],
                rule["body_template"], rule.get("enabled", 1),
                rule.get("sort_order", 99),
            ),
        )
    conn.commit()
    conn.close()


def delete_rule(rule_id: int):
    conn = get_db()
    conn.execute("DELETE FROM email_rules WHERE id=?", (rule_id,))
    conn.commit()
    conn.close()


# ── Staff ─────────────────────────────────────────────────────────────────────

def get_staff():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM staff_notifications WHERE enabled=1"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_staff(staff: dict):
    conn = get_db()
    if staff.get("id"):
        conn.execute(
            """UPDATE staff_notifications
               SET name=?, email=?, receives_drafts=?, enabled=?
               WHERE id=?""",
            (
                staff["name"], staff["email"],
                staff["receives_drafts"], staff["enabled"],
                staff["id"],
            ),
        )
    else:
        conn.execute(
            """INSERT INTO staff_notifications (name, email, receives_drafts, enabled)
               VALUES (?, ?, ?, ?)""",
            (
                staff["name"], staff["email"],
                staff.get("receives_drafts", 1),
                staff.get("enabled", 1),
            ),
        )
    conn.commit()
    conn.close()


def delete_staff(staff_id: int):
    conn = get_db()
    conn.execute("DELETE FROM staff_notifications WHERE id=?", (staff_id,))
    conn.commit()
    conn.close()


# ── Staff Signatures (dynamic per-accountant signature lookup) ───────────────
# These rows feed signature_builder.build_signature_html — the accountant's
# M365 signed-in email is matched (case-insensitively) to the `email` column
# to pick up their name + title for every plugin draft.

STAFF_SIGNATURE_COLUMNS = {"name", "title", "email", "enabled", "include_signature"}


def _validate_signature_email(email: str) -> str:
    """Normalise + sanity-check an email. Returns the lowercased value or
    raises ValueError. Validation is intentionally cheap — the M365 directory
    is the source of truth, this just blocks obvious typos."""
    e = (email or "").strip().lower()
    if not e or "@" not in e or " " in e:
        raise ValueError(f"Invalid email: {email!r}")
    return e


def get_staff_signatures(include_disabled: bool = True) -> list[dict]:
    """Return all staff signature rows, ordered by name."""
    conn = get_db()
    sql = "SELECT * FROM staff_signatures"
    if not include_disabled:
        sql += " WHERE enabled=1"
    sql += " ORDER BY name COLLATE NOCASE"
    rows = conn.execute(sql).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_staff_signature_by_email(email: str) -> dict | None:
    """Lookup a single enabled staff row by email (case-insensitive). Returns
    None if no row matches or the row is disabled — callers fall back to the
    legacy signature in either case."""
    if not email:
        return None
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM staff_signatures "
        "WHERE LOWER(email)=LOWER(?) AND enabled=1 LIMIT 1",
        (email.strip(),),
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def save_staff_signature(data: dict) -> int:
    """Insert or update a staff signature row. Returns the row id.

    Required keys: name, email. Optional: title, enabled, id (for update).
    Email uniqueness is enforced by the table constraint; on collision we
    raise sqlite3.IntegrityError so the API layer can surface a 400.
    """
    name = (data.get("name") or "").strip()
    if not name:
        raise ValueError("name is required")
    email = _validate_signature_email(data.get("email", ""))
    title = (data.get("title") or "").strip() or None
    enabled = 1 if data.get("enabled", 1) else 0
    include_signature = 1 if data.get("include_signature", 1) else 0

    conn = get_db()
    try:
        if data.get("id"):
            conn.execute(
                "UPDATE staff_signatures "
                "SET name=?, title=?, email=?, enabled=?, include_signature=?, "
                "    updated_at=strftime('%s','now') "
                "WHERE id=?",
                (name, title, email, enabled, include_signature, int(data["id"])),
            )
            row_id = int(data["id"])
        else:
            cur = conn.execute(
                "INSERT INTO staff_signatures "
                "(name, title, email, enabled, include_signature) "
                "VALUES (?, ?, ?, ?, ?)",
                (name, title, email, enabled, include_signature),
            )
            row_id = cur.lastrowid
        conn.commit()
        return row_id
    finally:
        conn.close()


def delete_staff_signature(row_id: int, soft: bool = True) -> None:
    """Soft-delete (enabled=0) by default to preserve history; hard-delete if
    soft=False. Either way the row stops matching for new drafts."""
    conn = get_db()
    if soft:
        conn.execute(
            "UPDATE staff_signatures SET enabled=0, "
            "updated_at=strftime('%s','now') WHERE id=?",
            (row_id,),
        )
    else:
        conn.execute("DELETE FROM staff_signatures WHERE id=?", (row_id,))
    conn.commit()
    conn.close()


# ── Activity Log ──────────────────────────────────────────────────────────────

def log_activity(from_email, subject, classification, action,
                 draft_created=0, notification_sent=0):
    conn = get_db()
    conn.execute(
        """INSERT INTO activity_log
           (from_email, subject, classification, action, draft_created, notification_sent)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (from_email, subject, classification, action, draft_created, notification_sent),
    )
    conn.commit()
    conn.close()


def get_recent_activity(limit=100):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM activity_log ORDER BY id DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ── Plugin registry ───────────────────────────────────────────────────────────

def get_plugin_state(plugin_id: str) -> dict:
    """Return persisted state for a plugin (enabled, draft_mode, schedule etc.)"""
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM plugin_registry WHERE plugin_id=?", (plugin_id,)
    ).fetchone()
    conn.close()
    if row:
        return dict(row)
    # First time we've seen this plugin — return defaults
    return {
        "plugin_id":        plugin_id,
        "enabled":          1,
        "draft_mode":       1,
        "schedule_seconds": 0,
        "last_run":         None,
        "last_result":      None,
        "last_summary":     None,
        "display_name":     None,
    }


def save_plugin_state(plugin_id: str, **kwargs):
    """Upsert plugin state in the registry.

    Column names from kwargs flow into the SQL string, so any caller that
    supplies an unknown key has their update dropped — only values keyed by
    ALLOWED_PLUGIN_COLUMNS reach the query.
    """
    unsafe_keys = set(kwargs.keys()) - ALLOWED_PLUGIN_COLUMNS
    if unsafe_keys:
        logger.warning(
            f"save_plugin_state rejected disallowed columns for {plugin_id}: "
            f"{sorted(unsafe_keys)}"
        )
    safe_kwargs = {k: v for k, v in kwargs.items() if k in ALLOWED_PLUGIN_COLUMNS}
    if not safe_kwargs:
        return

    conn = get_db()
    existing = conn.execute(
        "SELECT id FROM plugin_registry WHERE plugin_id=?", (plugin_id,)
    ).fetchone()

    if existing:
        sets = ", ".join(f"{k}=?" for k in safe_kwargs)
        vals = list(safe_kwargs.values()) + [plugin_id]
        conn.execute(
            f"UPDATE plugin_registry SET {sets} WHERE plugin_id=?", vals
        )
    else:
        safe_kwargs["plugin_id"] = plugin_id
        cols = ", ".join(safe_kwargs.keys())
        placeholders = ", ".join("?" * len(safe_kwargs))
        conn.execute(
            f"INSERT INTO plugin_registry ({cols}) VALUES ({placeholders})",
            list(safe_kwargs.values()),
        )

    conn.commit()
    conn.close()


def get_all_plugin_states() -> dict:
    """Return {plugin_id: state_dict} for all registered plugins."""
    conn = get_db()
    rows = conn.execute("SELECT * FROM plugin_registry").fetchall()
    conn.close()
    return {r["plugin_id"]: dict(r) for r in rows}


# ── Links & Forms ────────────────────────────────────────────────────────────

def get_links() -> list[dict]:
    """Return all links/forms."""
    conn = get_db()
    rows = conn.execute("SELECT * FROM links_forms ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_links_as_dict() -> dict:
    """Return {tag: url} for all enabled links. Used for template injection."""
    conn = get_db()
    rows = conn.execute(
        "SELECT tag, url FROM links_forms WHERE enabled=1"
    ).fetchall()
    conn.close()
    return {r["tag"]: r["url"] for r in rows}


def save_link(link: dict):
    conn = get_db()
    if link.get("id"):
        conn.execute(
            """UPDATE links_forms SET name=?, tag=?, url=?, enabled=? WHERE id=?""",
            (link["name"], link["tag"], link["url"], link["enabled"], link["id"]),
        )
    else:
        conn.execute(
            """INSERT INTO links_forms (name, tag, url, enabled) VALUES (?, ?, ?, ?)""",
            (link["name"], link["tag"], link.get("url", ""), link.get("enabled", 1)),
        )
    conn.commit()
    conn.close()


def delete_link(link_id: int):
    conn = get_db()
    conn.execute("DELETE FROM links_forms WHERE id=?", (link_id,))
    conn.commit()
    conn.close()


# ── Memory: Style Preferences ────────────────────────────────────────────────

def get_style_preferences() -> str:
    """Return the global style/tone instructions, or empty string."""
    conn = get_db()
    row = conn.execute("SELECT content FROM memory_style ORDER BY id DESC LIMIT 1").fetchone()
    conn.close()
    return row["content"] if row else ""


def save_style_preferences(content: str):
    """Overwrite the global style/tone instructions."""
    conn = get_db()
    conn.execute("DELETE FROM memory_style")
    if content.strip():
        conn.execute("INSERT INTO memory_style (content) VALUES (?)", (content.strip(),))
    conn.commit()
    conn.close()


# ── Memory: Chat Feedback ────────────────────────────────────────────────────

def add_feedback_message(role: str, message: str):
    """Add a message to the feedback chat history. role = 'user' or 'agent'."""
    conn = get_db()
    conn.execute(
        "INSERT INTO memory_feedback (role, message) VALUES (?, ?)",
        (role, message),
    )
    conn.commit()
    conn.close()


def get_feedback_history(limit=200) -> list[dict]:
    """Return chat history ordered oldest-first."""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM memory_feedback ORDER BY id ASC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def clear_feedback_history():
    conn = get_db()
    conn.execute("DELETE FROM memory_feedback")
    conn.commit()
    conn.close()


# ── Memory: Extracted Lessons ────────────────────────────────────────────────

def add_lesson(lesson: str, source: str = ""):
    """Store an extracted lesson from user feedback."""
    conn = get_db()
    conn.execute(
        "INSERT INTO memory_lessons (lesson, source) VALUES (?, ?)",
        (lesson, source),
    )
    conn.commit()
    conn.close()


def get_active_lessons() -> list[dict]:
    """Return all active lessons for injection into prompts."""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM memory_lessons WHERE active=1 ORDER BY id ASC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_lesson(lesson_id: int):
    conn = get_db()
    conn.execute("DELETE FROM memory_lessons WHERE id=?", (lesson_id,))
    conn.commit()
    conn.close()


def toggle_lesson(lesson_id: int, active: bool):
    conn = get_db()
    conn.execute(
        "UPDATE memory_lessons SET active=? WHERE id=?",
        (1 if active else 0, lesson_id),
    )
    conn.commit()
    conn.close()


# ── Plugin Templates ────────────────────────────────────────────────────────

def get_plugin_templates(plugin_id: str) -> list[dict]:
    """Return all templates for a given plugin."""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM plugin_templates WHERE plugin_id=? ORDER BY id",
        (plugin_id,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_plugin_template(plugin_id: str, key: str, value: str):
    """Upsert a single template value for a plugin."""
    conn = get_db()
    conn.execute(
        """INSERT INTO plugin_templates (plugin_id, template_key, template_value)
           VALUES (?, ?, ?)
           ON CONFLICT(plugin_id, template_key)
           DO UPDATE SET template_value=excluded.template_value""",
        (plugin_id, key, value),
    )
    conn.commit()
    conn.close()


def get_plugin_template(plugin_id: str, key: str, default: str = None) -> str:
    """Return a single template value, or default if not set."""
    conn = get_db()
    row = conn.execute(
        "SELECT template_value FROM plugin_templates WHERE plugin_id=? AND template_key=?",
        (plugin_id, key),
    ).fetchone()
    conn.close()
    return row["template_value"] if row else default


# ── Knowledge Base ────────────────────────────────────────────────────────────

def get_knowledge_entries(only_enabled: bool = False) -> list[dict]:
    """Return every knowledge base entry, optionally filtered to enabled only."""
    conn = get_db()
    sql = "SELECT * FROM knowledge_base"
    if only_enabled:
        sql += " WHERE enabled=1"
    sql += " ORDER BY category, id"
    rows = conn.execute(sql).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_knowledge_entry(category: str, title: str, content: str,
                        enabled: int = 1) -> int:
    """Insert a new knowledge base entry and return its ID."""
    conn = get_db()
    cur = conn.execute(
        """INSERT INTO knowledge_base (category, title, content, enabled)
           VALUES (?, ?, ?, ?)""",
        (category, title, content, enabled),
    )
    entry_id = cur.lastrowid
    conn.commit()
    conn.close()
    return entry_id


def update_knowledge_entry(entry_id: int, category: str = None,
                           title: str = None, content: str = None,
                           enabled: int = None):
    """Update any subset of fields on a knowledge base entry."""
    fields, values = [], []
    if category is not None:
        fields.append("category=?"); values.append(category)
    if title is not None:
        fields.append("title=?"); values.append(title)
    if content is not None:
        fields.append("content=?"); values.append(content)
    if enabled is not None:
        fields.append("enabled=?"); values.append(int(enabled))
    if not fields:
        return
    fields.append("updated_at=datetime('now','localtime')")
    values.append(entry_id)
    conn = get_db()
    conn.execute(
        f"UPDATE knowledge_base SET {', '.join(fields)} WHERE id=?",
        values,
    )
    conn.commit()
    conn.close()


def delete_knowledge_entry(entry_id: int):
    """Delete a knowledge base entry by ID."""
    conn = get_db()
    conn.execute("DELETE FROM knowledge_base WHERE id=?", (entry_id,))
    conn.commit()
    conn.close()


# ── BAS Clients ───────────────────────────────────────────────────────────────

BAS_CLIENT_COLUMNS = {
    "client_name", "entity_name", "abn", "frequency", "client_email",
    "last_data_received", "last_reminder_sent", "status", "notes",
}


def get_bas_clients() -> list[dict]:
    """Return every BAS client row ordered by client_name."""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM bas_clients ORDER BY client_name COLLATE NOCASE"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_bas_client(client_id: int) -> dict | None:
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM bas_clients WHERE id=?", (client_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def add_bas_client(data: dict) -> int:
    """Insert a BAS client row. Only whitelisted columns are used."""
    safe = {k: (data.get(k) or None) for k in BAS_CLIENT_COLUMNS if k in data}
    if not safe.get("client_name"):
        raise ValueError("client_name is required")
    safe.setdefault("frequency", "Quarterly")
    safe.setdefault("status", "Awaiting Data")
    cols = list(safe.keys())
    placeholders = ", ".join("?" * len(cols))
    conn = get_db()
    cur = conn.execute(
        f"INSERT INTO bas_clients ({', '.join(cols)}) VALUES ({placeholders})",
        [safe[c] for c in cols],
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return new_id


def update_bas_client(client_id: int, data: dict):
    """Update any subset of columns on a BAS client row."""
    safe = {k: data.get(k) for k in BAS_CLIENT_COLUMNS if k in data}
    if not safe:
        return
    sets = ", ".join(f"{k}=?" for k in safe)
    values = list(safe.values()) + [client_id]
    conn = get_db()
    conn.execute(f"UPDATE bas_clients SET {sets} WHERE id=?", values)
    conn.commit()
    conn.close()


def delete_bas_client(client_id: int):
    conn = get_db()
    conn.execute("DELETE FROM bas_clients WHERE id=?", (client_id,))
    conn.commit()
    conn.close()


def bulk_replace_bas_clients(rows: list[dict]) -> int:
    """Replace the entire bas_clients table with the supplied rows.
    Returns the number of rows inserted.
    """
    conn = get_db()
    conn.execute("DELETE FROM bas_clients")
    count = 0
    for data in rows:
        safe = {k: (data.get(k) or None) for k in BAS_CLIENT_COLUMNS if k in data}
        if not safe.get("client_name"):
            continue
        safe.setdefault("frequency", "Quarterly")
        safe.setdefault("status", "Awaiting Data")
        cols = list(safe.keys())
        placeholders = ", ".join("?" * len(cols))
        conn.execute(
            f"INSERT INTO bas_clients ({', '.join(cols)}) VALUES ({placeholders})",
            [safe[c] for c in cols],
        )
        count += 1
    conn.commit()
    conn.close()
    return count


# ── SharePoint upload-pending queue ───────────────────────────────────────────
# Recorded when smart_responder's auto-file path hits a folder governance
# error (SharePointFolderMissing or SharePointFolderAmbiguous). Drained by
# the morning briefing's pre-brief retry pass.

def record_sharepoint_upload_pending(
    message_id: str,
    client_name: str,
    candidate_names: list[str] | None = None,
) -> bool:
    """Record a pending SharePoint upload. INSERT OR IGNORE on
    (message_id, action) so re-runs against the same unresolved email don't
    bloat the queue.

    Returns True if a new row was inserted, False if the (message_id, action)
    pair already existed and the INSERT was ignored.

    candidate_names: pass the verbatim SharePoint folder names for an
    ambiguous-folder situation; pass None for a missing-folder situation.
    """
    if not message_id or not client_name:
        return False
    payload = (
        json.dumps(list(candidate_names))
        if candidate_names
        else None
    )
    conn = get_db()
    cur = conn.execute(
        "INSERT OR IGNORE INTO sharepoint_upload_pending "
        "(message_id, action, client_name, candidate_names_json) "
        "VALUES (?, 'upload_pending', ?, ?)",
        (message_id, client_name, payload),
    )
    inserted = cur.rowcount > 0
    conn.commit()
    conn.close()
    return inserted


def list_sharepoint_upload_pending(only_unresolved: bool = True) -> list[dict]:
    """Return pending SharePoint upload rows, oldest first.

    Each row dict contains: id, message_id, action, client_name,
    candidate_names (parsed list or None), created_at, resolved_at.
    """
    conn = get_db()
    if only_unresolved:
        sql = (
            "SELECT * FROM sharepoint_upload_pending "
            "WHERE resolved_at IS NULL ORDER BY created_at ASC"
        )
    else:
        sql = "SELECT * FROM sharepoint_upload_pending ORDER BY created_at ASC"
    rows = conn.execute(sql).fetchall()
    conn.close()
    out = []
    for r in rows:
        d = dict(r)
        raw = d.pop("candidate_names_json", None)
        try:
            d["candidate_names"] = json.loads(raw) if raw else None
        except (TypeError, ValueError):
            d["candidate_names"] = None
        out.append(d)
    return out


def delete_sharepoint_upload_pending(row_id: int) -> None:
    """Remove a pending row — called by the retry path on a successful
    upload. resolution-via-delete is the contract; the resolved_at column
    exists for forward-compat (e.g. an external reconciler that wants to
    mark resolved without deleting), but is unused today."""
    conn = get_db()
    conn.execute(
        "DELETE FROM sharepoint_upload_pending WHERE id = ?",
        (row_id,),
    )
    conn.commit()
    conn.close()
