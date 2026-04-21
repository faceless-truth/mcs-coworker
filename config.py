"""
MC & S Desktop Agent - Configuration & Database Manager
"""
import sqlite3
import json
import os
from pathlib import Path

DB_PATH = Path.home() / ".mcs_email_automation" / "config.db"


def get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
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
    """)

    # Add display_name column if upgrading from older schema
    try:
        c.execute("ALTER TABLE plugin_registry ADD COLUMN display_name TEXT")
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
        # Stream 3 — Heartbeat interval
        "heartbeat_interval_seconds": "60",
        # Stream 6 — Approval queue
        "approval_auto_threshold": "0.75",
        "approval_expiry_hours":   "48",
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

    conn.commit()
    conn.close()


# ── Settings ──────────────────────────────────────────────────────────────────

def get_setting(key, default=""):
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key, value):
    conn = get_db()
    conn.execute(
        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
        (key, str(value)),
    )
    conn.commit()
    conn.close()


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
    Query the Anthropic models API and update both stored model slots:
      - claude_model_fast      → newest available Haiku model
      - claude_model_reasoning → newest available Sonnet model

    Also keeps the legacy 'claude_model' key in sync with the fast model
    for backward compatibility with older plugins.

    Returns a dict with keys 'fast' and 'reasoning'.
    Falls back to currently stored values on any error.
    """
    import urllib.request
    fast_current = get_claude_model_fast()
    reasoning_current = get_claude_model_reasoning()
    result = {"fast": fast_current, "reasoning": reasoning_current}
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
    return {r["key"]: r["value"] for r in rows}


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
    """Upsert plugin state in the registry."""
    conn = get_db()
    existing = conn.execute(
        "SELECT id FROM plugin_registry WHERE plugin_id=?", (plugin_id,)
    ).fetchone()

    if existing:
        sets = ", ".join(f"{k}=?" for k in kwargs)
        vals = list(kwargs.values()) + [plugin_id]
        conn.execute(
            f"UPDATE plugin_registry SET {sets} WHERE plugin_id=?", vals
        )
    else:
        kwargs["plugin_id"] = plugin_id
        cols = ", ".join(kwargs.keys())
        placeholders = ", ".join("?" * len(kwargs))
        conn.execute(
            f"INSERT INTO plugin_registry ({cols}) VALUES ({placeholders})",
            list(kwargs.values()),
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
