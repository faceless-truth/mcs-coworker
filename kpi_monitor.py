"""
MCS CoWorker — Proactive KPI Monitoring (Tier 3C)
==================================================
Monitors key practice metrics and escalates automatically when thresholds
are breached — without waiting for a human to notice.

APEX ALIGNMENT
--------------
APEX monitors KPIs and escalates autonomously. This module gives CoWorker
the same capability: define thresholds, check them on every heartbeat tick,
and alert via Teams/email when breached.

BUILT-IN KPIs
-------------
1. WIP Ageing       — jobs with no movement for N days
2. Debtor Days      — invoices overdue by N days
3. Unsigned Docs    — FuseSign envelopes pending for N days
4. Inbox Backlog    — unread emails exceeding N count
5. Plugin Failures  — plugin error rate exceeding N% in last hour
6. AI Cost Spike    — daily AI spend exceeding $N AUD

USAGE
-----
KPIMonitor is a singleton. Call KPIMonitor.start(context) once at startup.
It subscribes to the heartbeat and checks all enabled KPIs on each tick.
Alerts are rate-limited to avoid spam (default: once per 4 hours per KPI).
"""

from __future__ import annotations
import logging
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger(__name__)


# ── Database helpers ──────────────────────────────────────────────────────────

def _get_kpi_db_path() -> Path:
    try:
        import config as cfg
        return Path(cfg.DB_PATH).parent / "kpi_monitor.db"
    except Exception:
        return Path.home() / ".mcs_coworker" / "kpi_monitor.db"


def init_kpi_db(db_path: Optional[Path] = None) -> None:
    """Create the KPI alert history table."""
    path = db_path or _get_kpi_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS kpi_alerts (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp   TEXT    DEFAULT (datetime('now','localtime')),
            kpi_id      TEXT    NOT NULL,
            severity    TEXT    NOT NULL DEFAULT 'warning',
            value       REAL,
            threshold   REAL,
            message     TEXT,
            notified    INTEGER DEFAULT 0
        );
        CREATE INDEX IF NOT EXISTS idx_kpi_alerts_ts
            ON kpi_alerts(timestamp);
        CREATE INDEX IF NOT EXISTS idx_kpi_alerts_kpi
            ON kpi_alerts(kpi_id);

        CREATE TABLE IF NOT EXISTS kpi_config (
            kpi_id      TEXT    PRIMARY KEY,
            enabled     INTEGER DEFAULT 1,
            threshold   REAL    NOT NULL,
            cooldown_h  REAL    DEFAULT 4.0,
            severity    TEXT    DEFAULT 'warning'
        );
    """)
    conn.commit()
    conn.close()


def _get_kpi_conn(db_path: Optional[Path] = None) -> sqlite3.Connection:
    path = db_path or _get_kpi_db_path()
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    return conn


# ── Default KPI definitions ───────────────────────────────────────────────────

DEFAULT_KPIS: list[dict[str, Any]] = [
    {
        "kpi_id":     "wip_ageing",
        "label":      "WIP Ageing",
        "description":"Jobs with no movement for more than N days",
        "threshold":  90.0,   # days
        "cooldown_h": 8.0,
        "severity":   "warning",
        "unit":       "days",
    },
    {
        "kpi_id":     "debtor_days",
        "label":      "Debtor Days",
        "description":"Invoices overdue by more than N days",
        "threshold":  60.0,   # days
        "cooldown_h": 8.0,
        "severity":   "warning",
        "unit":       "days",
    },
    {
        "kpi_id":     "unsigned_docs",
        "label":      "Unsigned Documents",
        "description":"FuseSign envelopes pending for more than N days",
        "threshold":  14.0,   # days
        "cooldown_h": 4.0,
        "severity":   "warning",
        "unit":       "days",
    },
    {
        "kpi_id":     "inbox_backlog",
        "label":      "Inbox Backlog",
        "description":"Unread emails exceeding N count",
        "threshold":  50.0,   # count
        "cooldown_h": 2.0,
        "severity":   "info",
        "unit":       "emails",
    },
    {
        "kpi_id":     "plugin_failures",
        "label":      "Plugin Failure Rate",
        "description":"Plugin error count in the last hour exceeding N",
        "threshold":  3.0,    # count
        "cooldown_h": 1.0,
        "severity":   "critical",
        "unit":       "errors",
    },
    {
        "kpi_id":     "ai_cost_spike",
        "label":      "AI Cost Spike",
        "description":"Daily AI spend exceeding $N AUD",
        "threshold":  5.00,   # AUD
        "cooldown_h": 4.0,
        "severity":   "warning",
        "unit":       "AUD",
    },
]


def seed_default_kpis(db_path: Optional[Path] = None) -> None:
    """Insert default KPI config rows if they don't exist."""
    conn = _get_kpi_conn(db_path)
    for kpi in DEFAULT_KPIS:
        conn.execute(
            "INSERT OR IGNORE INTO kpi_config (kpi_id, enabled, threshold, cooldown_h, severity) "
            "VALUES (?,1,?,?,?)",
            (kpi["kpi_id"], kpi["threshold"], kpi["cooldown_h"], kpi["severity"])
        )
    conn.commit()
    conn.close()


def get_kpi_config(db_path: Optional[Path] = None) -> list[dict]:
    """Return all KPI config rows merged with DEFAULT_KPIS metadata."""
    conn = _get_kpi_conn(db_path)
    rows = {r["kpi_id"]: dict(r) for r in
            conn.execute("SELECT * FROM kpi_config").fetchall()}
    conn.close()

    result = []
    for kpi in DEFAULT_KPIS:
        merged = dict(kpi)
        if kpi["kpi_id"] in rows:
            merged.update(rows[kpi["kpi_id"]])
        result.append(merged)
    return result


def update_kpi_threshold(kpi_id: str, threshold: float,
                          enabled: bool = True,
                          db_path: Optional[Path] = None) -> None:
    """Update a KPI threshold and enabled flag."""
    conn = _get_kpi_conn(db_path)
    conn.execute(
        "INSERT OR REPLACE INTO kpi_config (kpi_id, enabled, threshold, cooldown_h, severity) "
        "VALUES (?, ?, ?, "
        "(SELECT cooldown_h FROM kpi_config WHERE kpi_id=?), "
        "(SELECT severity  FROM kpi_config WHERE kpi_id=?))",
        (kpi_id, int(enabled), threshold, kpi_id, kpi_id)
    )
    conn.commit()
    conn.close()


def get_recent_alerts(limit: int = 50, db_path: Optional[Path] = None) -> list[dict]:
    """Return the most recent KPI alerts."""
    conn = _get_kpi_conn(db_path)
    rows = conn.execute(
        "SELECT * FROM kpi_alerts ORDER BY timestamp DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def _was_recently_alerted(kpi_id: str, cooldown_h: float,
                           db_path: Optional[Path] = None) -> bool:
    """Return True if this KPI was alerted within the cooldown window."""
    conn = _get_kpi_conn(db_path)
    cutoff = (datetime.now() - timedelta(hours=cooldown_h)).strftime("%Y-%m-%d %H:%M:%S")
    row = conn.execute(
        "SELECT id FROM kpi_alerts WHERE kpi_id=? AND timestamp >= ? LIMIT 1",
        (kpi_id, cutoff)
    ).fetchone()
    conn.close()
    return row is not None


def _record_alert(kpi_id: str, severity: str, value: float,
                  threshold: float, message: str,
                  db_path: Optional[Path] = None) -> None:
    """Persist a KPI alert to the database."""
    conn = _get_kpi_conn(db_path)
    conn.execute(
        "INSERT INTO kpi_alerts (kpi_id, severity, value, threshold, message) "
        "VALUES (?,?,?,?,?)",
        (kpi_id, severity, value, threshold, message)
    )
    conn.commit()
    conn.close()


# ── KPI check functions ───────────────────────────────────────────────────────

def _check_wip_ageing(context, config: dict) -> Optional[tuple[float, str]]:
    """Return (value, message) if threshold breached, else None."""
    if not (context.gateway and context.gateway.is_available("xpm")):
        return None
    try:
        jobs = context.gateway.xpm.list_jobs(status="in_progress") or []
        threshold_days = config["threshold"]
        cutoff = datetime.now() - timedelta(days=threshold_days)
        stale = []
        for job in jobs:
            last_updated = job.get("last_updated") or job.get("updated_at", "")
            if last_updated:
                try:
                    dt = datetime.fromisoformat(last_updated.replace("Z", "+00:00"))
                    if dt.replace(tzinfo=None) < cutoff:
                        stale.append(job.get("name", job.get("id", "?")))
                except Exception:
                    pass
        if stale:
            msg = (f"{len(stale)} jobs have had no movement for "
                   f"{threshold_days:.0f}+ days: "
                   f"{', '.join(stale[:5])}"
                   f"{'...' if len(stale) > 5 else ''}")
            return (float(len(stale)), msg)
    except Exception as e:
        logger.debug(f"[KPIMonitor] WIP ageing check failed: {e}")
    return None


def _check_debtor_days(context, config: dict) -> Optional[tuple[float, str]]:
    """Return (value, message) if threshold breached, else None."""
    if not (context.gateway and context.gateway.is_available("xpm")):
        return None
    try:
        invoices = context.gateway.xpm.list_jobs(status="invoiced") or []
        threshold_days = config["threshold"]
        cutoff = datetime.now() - timedelta(days=threshold_days)
        overdue = []
        for inv in invoices:
            invoice_date = inv.get("invoice_date") or inv.get("created_at", "")
            if invoice_date:
                try:
                    dt = datetime.fromisoformat(invoice_date.replace("Z", "+00:00"))
                    if dt.replace(tzinfo=None) < cutoff:
                        overdue.append(inv.get("client_name", inv.get("id", "?")))
                except Exception:
                    pass
        if overdue:
            msg = (f"{len(overdue)} invoices overdue by {threshold_days:.0f}+ days: "
                   f"{', '.join(overdue[:5])}"
                   f"{'...' if len(overdue) > 5 else ''}")
            return (float(len(overdue)), msg)
    except Exception as e:
        logger.debug(f"[KPIMonitor] Debtor days check failed: {e}")
    return None


def _check_unsigned_docs(context, config: dict) -> Optional[tuple[float, str]]:
    """Return (value, message) if threshold breached, else None."""
    if not (context.gateway and context.gateway.is_available("fusesign")):
        return None
    try:
        envelopes = context.gateway.fusesign.list_envelopes(status="pending") or []
        threshold_days = config["threshold"]
        cutoff = datetime.now() - timedelta(days=threshold_days)
        overdue = []
        for env in envelopes:
            sent_at = env.get("sent_at") or env.get("created_at", "")
            if sent_at:
                try:
                    dt = datetime.fromisoformat(sent_at.replace("Z", "+00:00"))
                    if dt.replace(tzinfo=None) < cutoff:
                        overdue.append(env.get("name", env.get("id", "?")))
                except Exception:
                    pass
        if overdue:
            msg = (f"{len(overdue)} documents unsigned for {threshold_days:.0f}+ days: "
                   f"{', '.join(overdue[:5])}"
                   f"{'...' if len(overdue) > 5 else ''}")
            return (float(len(overdue)), msg)
    except Exception as e:
        logger.debug(f"[KPIMonitor] Unsigned docs check failed: {e}")
    return None


def _check_inbox_backlog(context, config: dict) -> Optional[tuple[float, str]]:
    """Return (value, message) if threshold breached, else None."""
    if not context.graph:
        return None
    try:
        # Use the graph client to count unread emails
        messages = context.graph.get_unread_emails(max_results=200) or []
        count = len(messages)
        threshold = config["threshold"]
        if count >= threshold:
            msg = f"Inbox has {count} unread emails (threshold: {threshold:.0f})"
            return (float(count), msg)
    except Exception as e:
        logger.debug(f"[KPIMonitor] Inbox backlog check failed: {e}")
    return None


def _check_plugin_failures(context, config: dict) -> Optional[tuple[float, str]]:
    """Return (value, message) if threshold breached, else None."""
    try:
        from event_bus import EventBus
        cutoff_ts = (datetime.now() - timedelta(hours=1)).timestamp()
        history = EventBus.get_history(event_type="plugin.run.failed") or []
        recent_failures = [
            e for e in history
            if e.get("timestamp", 0) >= cutoff_ts
        ]
        count = len(recent_failures)
        threshold = config["threshold"]
        if count >= threshold:
            failed_plugins = list({
                e.get("payload", {}).get("plugin_id", "?")
                for e in recent_failures
            })
            msg = (f"{count} plugin failures in the last hour: "
                   f"{', '.join(failed_plugins[:5])}")
            return (float(count), msg)
    except Exception as e:
        logger.debug(f"[KPIMonitor] Plugin failures check failed: {e}")
    return None


def _check_ai_cost_spike(context, config: dict) -> Optional[tuple[float, str]]:
    """Return (value, message) if threshold breached, else None."""
    try:
        from token_meter import get_usage_summary
        summary = get_usage_summary(days=1)
        today_cost = summary.get("today_cost_aud", 0.0)
        threshold = config["threshold"]
        if today_cost >= threshold:
            msg = (f"Daily AI spend is A${today_cost:.2f} "
                   f"(threshold: A${threshold:.2f})")
            return (today_cost, msg)
    except Exception as e:
        logger.debug(f"[KPIMonitor] AI cost spike check failed: {e}")
    return None


_KPI_CHECKERS = {
    "wip_ageing":      _check_wip_ageing,
    "debtor_days":     _check_debtor_days,
    "unsigned_docs":   _check_unsigned_docs,
    "inbox_backlog":   _check_inbox_backlog,
    "plugin_failures": _check_plugin_failures,
    "ai_cost_spike":   _check_ai_cost_spike,
}


# ── Alert dispatch ────────────────────────────────────────────────────────────

def _dispatch_alert(context, kpi_id: str, severity: str,
                    value: float, threshold: float, message: str) -> None:
    """Send the alert via Teams and/or email."""
    title_map = {
        "warning":  "⚠️ KPI Alert",
        "critical": "🚨 Critical Alert",
        "info":     "ℹ️ KPI Notice",
    }
    title = title_map.get(severity, "⚠️ KPI Alert")
    full_msg = f"[{kpi_id}] {message}"

    # Teams
    if context.gateway and context.gateway.is_available("teams"):
        try:
            context.gateway.teams.send_alert(
                title=title,
                message=full_msg,
                color=severity)
            logger.info(f"[KPIMonitor] Teams alert sent for {kpi_id}")
        except Exception as e:
            logger.warning(f"[KPIMonitor] Teams alert failed: {e}")

    # Email fallback
    if context.notify:
        try:
            context.notify(
                subject=f"{title}: {kpi_id}",
                body=full_msg)
        except Exception as e:
            logger.debug(f"[KPIMonitor] Email notify failed: {e}")


# ── KPIMonitor singleton ──────────────────────────────────────────────────────

class _KPIMonitor:
    """Singleton that runs KPI checks on every heartbeat tick."""

    def __init__(self):
        self._context = None
        self._db_path: Optional[Path] = None
        self._running = False

    def start(self, context, db_path: Optional[Path] = None) -> None:
        """Initialise the monitor and subscribe to heartbeat ticks."""
        self._context = context
        self._db_path = db_path
        init_kpi_db(db_path)
        seed_default_kpis(db_path)

        if not self._running:
            try:
                from event_bus import EventBus
                EventBus.subscribe("heartbeat.tick", self._on_tick,
                                   subscriber_id="kpi_monitor.tick")
                self._running = True
                logger.info("[KPIMonitor] Started — subscribed to heartbeat.tick")
            except Exception as e:
                logger.warning(f"[KPIMonitor] Could not subscribe to heartbeat: {e}")

    def stop(self) -> None:
        """Unsubscribe from the heartbeat."""
        try:
            from event_bus import EventBus
            EventBus.unsubscribe("heartbeat.tick", subscriber_id="kpi_monitor.tick")
        except Exception:
            pass
        self._running = False

    def update_context(self, context) -> None:
        """Update the context (called when a new context is built)."""
        self._context = context

    def run_checks(self, context=None) -> list[dict]:
        """
        Run all enabled KPI checks immediately.
        Returns a list of triggered alerts.
        """
        ctx = context or self._context
        if ctx is None:
            return []

        kpi_configs = get_kpi_config(self._db_path)
        triggered = []

        for kpi in kpi_configs:
            kpi_id  = kpi["kpi_id"]
            enabled = kpi.get("enabled", 1)
            if not enabled:
                continue

            checker = _KPI_CHECKERS.get(kpi_id)
            if not checker:
                continue

            try:
                result = checker(ctx, kpi)
                if result is None:
                    continue

                value, message = result
                cooldown_h = kpi.get("cooldown_h", 4.0)
                severity   = kpi.get("severity", "warning")

                if _was_recently_alerted(kpi_id, cooldown_h, self._db_path):
                    logger.debug(f"[KPIMonitor] {kpi_id} in cooldown — skipping")
                    continue

                _record_alert(kpi_id, severity, value,
                              kpi["threshold"], message, self._db_path)
                _dispatch_alert(ctx, kpi_id, severity,
                                value, kpi["threshold"], message)

                triggered.append({
                    "kpi_id":    kpi_id,
                    "severity":  severity,
                    "value":     value,
                    "threshold": kpi["threshold"],
                    "message":   message,
                })
                logger.info(f"[KPIMonitor] Alert triggered: {kpi_id} = {value}")

            except Exception as e:
                logger.warning(f"[KPIMonitor] Check failed for {kpi_id}: {e}")

        return triggered

    def _on_tick(self, event) -> None:
        """Called on every heartbeat tick — run KPI checks."""
        if self._context:
            try:
                self.run_checks()
            except Exception as e:
                logger.warning(f"[KPIMonitor] Tick error: {e}")


# Module-level singleton
KPIMonitor = _KPIMonitor()
