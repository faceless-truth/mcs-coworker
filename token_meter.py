"""
MCS CoWorker — Token Metering & Cost Tracking (Tier 3B)
========================================================
Wraps every Claude API call to capture token usage, calculate cost,
and persist records to SQLite for the dashboard.

APEX ALIGNMENT
--------------
APEX tracks AI spend per agent and per task. This module gives CoWorker
the same visibility — every Claude call is logged with model, plugin,
input/output tokens, and estimated AUD cost.

PRICING (as of April 2025, USD converted at 0.65 AUD/USD)
----------------------------------------------------------
Model                           Input ($/1M)   Output ($/1M)
claude-haiku-4-5-20251001           0.80            4.00
claude-sonnet-4-6                   3.00           15.00
claude-3-5-haiku-20241022           0.80            4.00
claude-3-5-sonnet-20241022          3.00           15.00
claude-3-opus-20240229             15.00           75.00
(default fallback for unknown)      3.00           15.00

All costs stored in USD; displayed in AUD in the UI.
"""

from __future__ import annotations
import logging
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger(__name__)

# ── Pricing table (USD per 1M tokens) ────────────────────────────────────────

_PRICING: dict[str, tuple[float, float]] = {
    # (input_per_1M_usd, output_per_1M_usd)
    "claude-haiku-4-5-20251001":    (0.80,  4.00),
    "claude-3-5-haiku-20241022":    (0.80,  4.00),
    "claude-3-haiku-20240307":      (0.25,  1.25),
    "claude-sonnet-4-6":            (3.00, 15.00),
    "claude-3-5-sonnet-20241022":   (3.00, 15.00),
    "claude-3-5-sonnet-20240620":   (3.00, 15.00),
    "claude-3-sonnet-20240229":     (3.00, 15.00),
    "claude-3-opus-20240229":      (15.00, 75.00),
}
_DEFAULT_PRICING = (3.00, 15.00)
_AUD_RATE = 1.55  # 1 USD ≈ 1.55 AUD (approximate)


def _get_pricing(model: str) -> tuple[float, float]:
    """Return (input_per_1M, output_per_1M) in USD for the given model."""
    # Exact match first
    if model in _PRICING:
        return _PRICING[model]
    # Prefix match (handles versioned names)
    for key, pricing in _PRICING.items():
        if model.startswith(key) or key.startswith(model):
            return pricing
    return _DEFAULT_PRICING


def calculate_cost_usd(model: str, input_tokens: int, output_tokens: int) -> float:
    """Calculate the USD cost for a single API call."""
    inp_rate, out_rate = _get_pricing(model)
    return (input_tokens * inp_rate + output_tokens * out_rate) / 1_000_000


def usd_to_aud(usd: float) -> float:
    """Convert USD to AUD using the approximate rate."""
    return usd * _AUD_RATE


# ── Database helpers ──────────────────────────────────────────────────────────

def _get_meter_db_path() -> Path:
    """Return the path to the token meter SQLite database."""
    try:
        import config as cfg
        return Path(cfg.DB_PATH).parent / "token_meter.db"
    except Exception:
        return Path.home() / ".mcs_coworker" / "token_meter.db"


def init_meter_db(db_path: Optional[Path] = None) -> None:
    """Create the token_usage table if it doesn't exist."""
    path = db_path or _get_meter_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS token_usage (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp       TEXT    DEFAULT (datetime('now','localtime')),
            plugin_id       TEXT    NOT NULL,
            model           TEXT    NOT NULL,
            tier            TEXT    NOT NULL DEFAULT 'unknown',
            input_tokens    INTEGER NOT NULL DEFAULT 0,
            output_tokens   INTEGER NOT NULL DEFAULT 0,
            total_tokens    INTEGER NOT NULL DEFAULT 0,
            cost_usd        REAL    NOT NULL DEFAULT 0.0,
            cost_aud        REAL    NOT NULL DEFAULT 0.0,
            prompt_summary  TEXT,
            success         INTEGER NOT NULL DEFAULT 1
        );
        CREATE INDEX IF NOT EXISTS idx_token_usage_ts
            ON token_usage(timestamp);
        CREATE INDEX IF NOT EXISTS idx_token_usage_plugin
            ON token_usage(plugin_id);
    """)
    conn.commit()
    conn.close()


def log_usage(
    plugin_id: str,
    model: str,
    input_tokens: int,
    output_tokens: int,
    tier: str = "unknown",
    prompt_summary: str = "",
    success: bool = True,
    db_path: Optional[Path] = None,
) -> float:
    """
    Log a Claude API call to the token meter database.
    Returns the USD cost of the call.
    """
    path = db_path or _get_meter_db_path()
    cost_usd = calculate_cost_usd(model, input_tokens, output_tokens)
    cost_aud = usd_to_aud(cost_usd)
    total    = input_tokens + output_tokens

    try:
        conn = sqlite3.connect(str(path))
        conn.execute(
            """INSERT INTO token_usage
               (plugin_id, model, tier, input_tokens, output_tokens,
                total_tokens, cost_usd, cost_aud, prompt_summary, success)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (plugin_id, model, tier, input_tokens, output_tokens,
             total, cost_usd, cost_aud, prompt_summary[:200], int(success))
        )
        conn.commit()
        conn.close()
    except Exception as e:
        logger.warning(f"[TokenMeter] Failed to log usage: {e}")

    return cost_usd


def get_usage_summary(
    days: int = 30,
    db_path: Optional[Path] = None,
) -> dict[str, Any]:
    """
    Return a summary dict for the dashboard:
    {
        total_calls, total_tokens, total_cost_usd, total_cost_aud,
        by_plugin: [{plugin_id, calls, tokens, cost_aud}],
        by_model:  [{model, calls, tokens, cost_aud}],
        by_day:    [{date, calls, tokens, cost_aud}],
        today_cost_aud, this_month_cost_aud
    }
    """
    path = db_path or _get_meter_db_path()
    if not path.exists():
        return _empty_summary()

    since = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
    today = datetime.now().strftime("%Y-%m-%d")
    month_start = datetime.now().strftime("%Y-%m-01")

    try:
        conn = sqlite3.connect(str(path))
        conn.row_factory = sqlite3.Row

        totals = conn.execute(
            "SELECT COUNT(*) as calls, SUM(total_tokens) as tokens, "
            "SUM(cost_usd) as cost_usd, SUM(cost_aud) as cost_aud "
            "FROM token_usage WHERE timestamp >= ?", (since,)
        ).fetchone()

        by_plugin = conn.execute(
            "SELECT plugin_id, COUNT(*) as calls, SUM(total_tokens) as tokens, "
            "SUM(cost_aud) as cost_aud "
            "FROM token_usage WHERE timestamp >= ? "
            "GROUP BY plugin_id ORDER BY cost_aud DESC LIMIT 20", (since,)
        ).fetchall()

        by_model = conn.execute(
            "SELECT model, tier, COUNT(*) as calls, SUM(total_tokens) as tokens, "
            "SUM(cost_aud) as cost_aud "
            "FROM token_usage WHERE timestamp >= ? "
            "GROUP BY model ORDER BY cost_aud DESC", (since,)
        ).fetchall()

        by_day = conn.execute(
            "SELECT substr(timestamp,1,10) as date, COUNT(*) as calls, "
            "SUM(total_tokens) as tokens, SUM(cost_aud) as cost_aud "
            "FROM token_usage WHERE timestamp >= ? "
            "GROUP BY date ORDER BY date DESC LIMIT 30", (since,)
        ).fetchall()

        today_cost = conn.execute(
            "SELECT COALESCE(SUM(cost_aud), 0) FROM token_usage "
            "WHERE timestamp >= ?", (today + " 00:00:00",)
        ).fetchone()[0]

        month_cost = conn.execute(
            "SELECT COALESCE(SUM(cost_aud), 0) FROM token_usage "
            "WHERE timestamp >= ?", (month_start + " 00:00:00",)
        ).fetchone()[0]

        conn.close()

        return {
            "total_calls":       totals["calls"] or 0,
            "total_tokens":      totals["tokens"] or 0,
            "total_cost_usd":    round(totals["cost_usd"] or 0, 4),
            "total_cost_aud":    round(totals["cost_aud"] or 0, 4),
            "today_cost_aud":    round(today_cost, 4),
            "this_month_cost_aud": round(month_cost, 4),
            "by_plugin": [dict(r) for r in by_plugin],
            "by_model":  [dict(r) for r in by_model],
            "by_day":    [dict(r) for r in by_day],
            "days":      days,
        }
    except Exception as e:
        logger.warning(f"[TokenMeter] Failed to get summary: {e}")
        return _empty_summary()


def _empty_summary() -> dict[str, Any]:
    return {
        "total_calls": 0, "total_tokens": 0,
        "total_cost_usd": 0.0, "total_cost_aud": 0.0,
        "today_cost_aud": 0.0, "this_month_cost_aud": 0.0,
        "by_plugin": [], "by_model": [], "by_day": [], "days": 30,
    }


# ── ClaudeUsageWrapper ────────────────────────────────────────────────────────

class ClaudeUsageWrapper:
    """
    Wraps an Anthropic client to intercept every messages.create() call,
    extract token usage from the response, and log it to the meter DB.

    Usage:
        wrapped = ClaudeUsageWrapper(client, plugin_id="plugin_email_triage", tier="fast")
        response = wrapped.messages.create(model=..., messages=...)
    """

    def __init__(
        self,
        client,
        plugin_id: str = "unknown",
        tier: str = "unknown",
        db_path: Optional[Path] = None,
    ):
        self._client    = client
        self._plugin_id = plugin_id
        self._tier      = tier
        self._db_path   = db_path
        self.messages   = _MessagesProxy(self)

    @property
    def plugin_id(self) -> str:
        return self._plugin_id

    @plugin_id.setter
    def plugin_id(self, value: str):
        self._plugin_id = value
        self.messages._wrapper = self

    def _record(
        self,
        model: str,
        response,
        prompt_summary: str = "",
        success: bool = True,
    ) -> None:
        """Extract usage from response and log it."""
        try:
            usage = getattr(response, "usage", None)
            if usage:
                input_tokens  = getattr(usage, "input_tokens",  0) or 0
                output_tokens = getattr(usage, "output_tokens", 0) or 0
            else:
                input_tokens = output_tokens = 0

            log_usage(
                plugin_id=self._plugin_id,
                model=model,
                input_tokens=input_tokens,
                output_tokens=output_tokens,
                tier=self._tier,
                prompt_summary=prompt_summary,
                success=success,
                db_path=self._db_path,
            )
        except Exception as e:
            logger.debug(f"[TokenMeter] Record error: {e}")


class _MessagesProxy:
    """Proxy for client.messages that intercepts create() calls."""

    def __init__(self, wrapper: ClaudeUsageWrapper):
        self._wrapper = wrapper

    def create(self, *args, **kwargs) -> Any:
        model          = kwargs.get("model", args[0] if args else "unknown")
        messages_list  = kwargs.get("messages", [])
        prompt_summary = ""
        if messages_list:
            first = messages_list[0]
            content = first.get("content", "")
            if isinstance(content, str):
                prompt_summary = content[:100]
            elif isinstance(content, list) and content:
                prompt_summary = str(content[0])[:100]

        success = True
        response = None
        try:
            response = self._wrapper._client.messages.create(*args, **kwargs)
        except Exception as e:
            success = False
            log_usage(
                plugin_id=self._wrapper._plugin_id,
                model=str(model),
                input_tokens=0,
                output_tokens=0,
                tier=self._wrapper._tier,
                prompt_summary=prompt_summary,
                success=False,
                db_path=self._wrapper._db_path,
            )
            raise

        self._wrapper._record(
            model=str(model),
            response=response,
            prompt_summary=prompt_summary,
            success=success,
        )
        return response

    def __getattr__(self, name: str):
        """Pass through any other messages.* attributes to the real client."""
        return getattr(self._wrapper._client.messages, name)


# ── Inject into PluginContext ─────────────────────────────────────────────────

def wrap_context_claude(context, plugin_id: str) -> None:
    """
    Replace context.claude / context.claude_fast / context.claude_reason
    with ClaudeUsageWrapper instances so all calls are metered.
    Called by plugin_loader._make_context() after building the context.
    """
    init_meter_db()

    if context.claude is not None:
        context.claude = ClaudeUsageWrapper(
            context.claude, plugin_id=plugin_id, tier="default")

    if context.claude_fast is not None:
        context.claude_fast = ClaudeUsageWrapper(
            context.claude_fast, plugin_id=plugin_id, tier="fast")

    if context.claude_reason is not None:
        context.claude_reason = ClaudeUsageWrapper(
            context.claude_reason, plugin_id=plugin_id, tier="reason")
