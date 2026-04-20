"""
MCS CoWorker — Flask API Server
Bridges the React/pywebview frontend to all Python backend modules.
Runs on localhost:7842 — not exposed externally.
"""
from __future__ import annotations

import json
import os
import sys
import threading
import traceback
from datetime import datetime
from functools import wraps
from typing import Any

import queue
import time
from flask import Flask, Response, jsonify, request, stream_with_context, send_from_directory
from flask_cors import CORS

# ── Backend imports ────────────────────────────────────────────────────────────
import config
from config import (
    init_db,
    get_setting, set_setting,
    get_all_settings,
    get_recent_activity,
    get_all_plugin_states,
    get_claude_model_fast,
    get_claude_model_reasoning,
    update_claude_models,
    get_rules, save_rule, delete_rule,
    get_staff, save_staff, delete_staff,
    get_links, save_link, delete_link,
    get_active_lessons, add_lesson, delete_lesson, toggle_lesson,
    get_style_preferences, save_style_preferences,
    add_feedback_message, get_feedback_history, clear_feedback_history,
)
from plugin_loader import PluginLoader
from approval_queue import ApprovalQueue
from token_meter import get_usage_summary
from event_bus import EventBus
from kpi_monitor import KPIMonitor

# ── App setup ──────────────────────────────────────────────────────────────────
app = Flask(__name__)
CORS(app, origins=[
    "http://localhost:7842", "http://127.0.0.1:7842",
    "http://localhost:3000", "http://127.0.0.1:3000",
    "http://localhost:5173", "http://127.0.0.1:5173",
], supports_credentials=True)
# Electron uses file:// or app:// — handle via wildcard on those routes separately

API_PORT = 7842

# Shared state — populated by main.py on startup
_loader: PluginLoader | None = None
_approval_queue: ApprovalQueue | None = None
_kpi_monitor: KPIMonitor | None = None
_graph_client = None  # set by main.py after GraphClient is created
_start_time: datetime = datetime.now()


def set_loader(loader: PluginLoader):
    global _loader
    _loader = loader


def set_approval_queue(aq: ApprovalQueue):
    global _approval_queue
    _approval_queue = aq


def set_kpi_monitor(km: KPIMonitor):
    global _kpi_monitor
    _kpi_monitor = km


def set_graph_client(gc):
    global _graph_client
    _graph_client = gc


# ── Helpers ────────────────────────────────────────────────────────────────────
def ok(data: Any = None, **kwargs):
    payload = {"ok": True}
    if data is not None:
        payload["data"] = data
    payload.update(kwargs)
    return jsonify(payload)


def err(message: str, status: int = 400):
    return jsonify({"ok": False, "error": message}), status


def require_loader(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if _loader is None:
            return err("Plugin loader not initialised", 503)
        return f(*args, **kwargs)
    return wrapper


# ── Health ─────────────────────────────────────────────────────────────────────
@app.route("/api/health")
def health():
    uptime_secs = int((datetime.now() - _start_time).total_seconds())
    h, rem = divmod(uptime_secs, 3600)
    m, s = divmod(rem, 60)
    return ok({
        "status": "ok",
        "uptime": f"{h}h {m}m {s}s",
        "version": get_setting("app_version", "2.4.1"),
        "fast_model": get_claude_model_fast(),
        "reasoning_model": get_claude_model_reasoning(),
    })


# ── Plugins ────────────────────────────────────────────────────────────────────
@app.route("/api/plugins")
@require_loader
def list_plugins():
    plugins = []
    states = get_all_plugin_states()
    for lp in _loader.get_plugins():
        inst = lp.instance
        state = states.get(lp.plugin_id, {})
        plugins.append({
            "id": lp.plugin_id,
            "name": inst.name,
            "description": getattr(inst, "description", ""),
            "enabled": lp.enabled,
            "status": "disabled" if not lp.enabled else "idle",
            "lastRun": _format_dt(lp.last_run),
            "nextRun": _format_next(lp),
            "schedule": _schedule_label(lp),
            "runsToday": state.get("runs_today", 0),
            "successRate": state.get("success_rate", 100),
            "model": getattr(inst, "MODEL_TIER", "haiku"),
            "category": getattr(inst, "CATEGORY", "Core"),
        })
    return ok(plugins)


@app.route("/api/plugins/<plugin_id>/enable", methods=["POST"])
@require_loader
def enable_plugin(plugin_id):
    body = request.get_json(silent=True) or {}
    enabled = body.get("enabled", True)
    _loader.set_plugin_enabled(plugin_id, enabled)
    return ok({"plugin_id": plugin_id, "enabled": enabled})


@app.route("/api/plugins/<plugin_id>/run", methods=["POST"])
@require_loader
def run_plugin(plugin_id):
    def _run():
        _loader.run_plugin(plugin_id, manual=True)
    threading.Thread(target=_run, daemon=True).start()
    return ok({"plugin_id": plugin_id, "triggered": True})


# ── Activity ───────────────────────────────────────────────────────────────────
@app.route("/api/activity")
def activity():
    limit = int(request.args.get("limit", 100))
    rows = get_recent_activity(limit)
    formatted = []
    for r in rows:
        formatted.append({
            "id": r.get("id"),
            "time": _format_time(r.get("timestamp", "")),
            "plugin": r.get("plugin_id", ""),
            "action": r.get("subject", r.get("body", "")),
            "status": r.get("status", "success"),
        })
    return ok(formatted)


# ── Activity SSE stream ────────────────────────────────────────────────────────
# Subscribers receive new activity rows as Server-Sent Events
_sse_subscribers: list[queue.Queue] = []
_sse_lock = threading.Lock()


def _broadcast_activity(entry: dict):
    """Push a new activity entry to all connected SSE clients."""
    with _sse_lock:
        dead = []
        for q in _sse_subscribers:
            try:
                q.put_nowait(entry)
            except queue.Full:
                dead.append(q)
        for q in dead:
            _sse_subscribers.remove(q)


# Wire into EventBus so every plugin.run.complete event triggers a broadcast
def _wire_sse_to_event_bus():
    from event_bus import EventBus
    def _on_plugin_complete(event_type: str, data: dict):
        entry = {
            "id": f"live-{int(time.time()*1000)}",
            "time": datetime.now().strftime("%H:%M:%S"),
            "plugin": data.get("plugin_id", "unknown"),
            "action": data.get("message", "Plugin run completed"),
            "status": "success" if data.get("success", True) else "error",
        }
        _broadcast_activity(entry)
    EventBus.subscribe("plugin.run.complete", _on_plugin_complete, subscriber_id="sse_bridge")
    EventBus.subscribe("plugin.run.failed", lambda et, d: _broadcast_activity({
        "id": f"live-{int(time.time()*1000)}",
        "time": datetime.now().strftime("%H:%M:%S"),
        "plugin": d.get("plugin_id", "unknown"),
        "action": d.get("error", "Plugin run failed"),
        "status": "error",
    }), subscriber_id="sse_bridge_fail")


@app.route("/api/activity/stream")
def activity_stream():
    """Server-Sent Events endpoint — streams new activity entries in real time."""
    client_q: queue.Queue = queue.Queue(maxsize=50)
    with _sse_lock:
        _sse_subscribers.append(client_q)

    def generate():
        # Send a ping immediately so the client knows the connection is live
        yield "event: ping\ndata: connected\n\n"
        try:
            while True:
                try:
                    entry = client_q.get(timeout=20)
                    yield f"data: {json.dumps(entry)}\n\n"
                except queue.Empty:
                    # Keepalive ping every 20s
                    yield "event: ping\ndata: keepalive\n\n"
        except GeneratorExit:
            pass
        finally:
            with _sse_lock:
                if client_q in _sse_subscribers:
                    _sse_subscribers.remove(client_q)

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Access-Control-Allow-Origin": "*",
        },
    )


# ── Approvals ──────────────────────────────────────────────────────────────────
@app.route("/api/approvals")
def list_approvals():
    if _approval_queue is None:
        return ok([])
    items = _approval_queue.list_pending()
    return ok(items)


@app.route("/api/approvals/<action_id>/approve", methods=["POST"])
def approve_action(action_id):
    if _approval_queue is None:
        return err("Approval queue not initialised", 503)
    body = request.get_json(silent=True) or {}
    reviewer_note = body.get("reviewer_note", "")
    _approval_queue.approve(int(action_id), reviewer_note=reviewer_note)
    return ok({"action_id": action_id, "approved": True})


@app.route("/api/approvals/<action_id>/reject", methods=["POST"])
def reject_action(action_id):
    if _approval_queue is None:
        return err("Approval queue not initialised", 503)
    body = request.get_json(silent=True) or {}
    reviewer_note = body.get("reviewer_note", "")
    _approval_queue.reject(int(action_id), reviewer_note=reviewer_note)
    return ok({"action_id": action_id, "rejected": True})


@app.route("/api/approvals/<action_id>/edit", methods=["POST"])
def edit_approval(action_id):
    """Edit the payload of a pending action before approving (edit-before-approve)."""
    if _approval_queue is None:
        return err("Approval queue not initialised", 503)
    body = request.get_json(silent=True) or {}
    updated_payload = body.get("payload")
    if updated_payload is None:
        return err("'payload' field is required")
    success = _approval_queue.edit_payload(int(action_id), updated_payload)
    if not success:
        return err("Action not found or not pending", 404)
    return ok({"action_id": action_id, "updated": True})


# ── Memory ─────────────────────────────────────────────────────────────────────
@app.route("/api/memory")
def list_memory():
    try:
        from memory_store import MemoryStore
        ms = MemoryStore()
        query = request.args.get("q", "recent client interactions")
        limit = int(request.args.get("limit", 50))
        results = ms.search(query, n_results=limit)
        return ok(results)
    except Exception as e:
        return ok([])  # graceful degradation if ChromaDB not ready


@app.route("/api/memory/<record_id>", methods=["DELETE"])
def delete_memory(record_id):
    try:
        from memory_store import MemoryStore
        ms = MemoryStore()
        ms.delete(record_id)
        return ok({"deleted": record_id})
    except Exception as e:
        return err(str(e))


# ── Events ─────────────────────────────────────────────────────────────────────
@app.route("/api/events")
def list_events():
    limit = int(request.args.get("limit", 50))
    history = EventBus.get_history(limit)
    formatted = []
    for evt in reversed(history):
        formatted.append({
            "id": evt.get("id", ""),
            "time": _format_time(evt.get("timestamp", "")),
            "type": evt.get("event_type", ""),
            "source": evt.get("source", ""),
            "payload": str(evt.get("data", "")),
        })
    return ok(formatted)


# ── KPI ────────────────────────────────────────────────────────────────────────
@app.route("/api/kpi")
def kpi():
    if _kpi_monitor is None:
        return ok([])
    try:
        metrics = _kpi_monitor.get_current_metrics()
        return ok(metrics)
    except Exception:
        return ok([])


# ── Usage ──────────────────────────────────────────────────────────────────────
@app.route("/api/usage")
def usage():
    try:
        s = get_usage_summary(days=30)
        return ok({
            "todayCost": s.get("today_cost_aud", 0),
            "monthlyCost": s.get("this_month_cost_aud", 0),
            "monthlyBudget": float(get_setting("monthly_ai_budget_aud", "100")),
            "totalCalls": s.get("total_calls", 0),
            "totalTokensIn": s.get("total_tokens", 0),
            "totalTokensOut": 0,
            "byPlugin": [
                {"plugin_id": r["plugin_id"], "calls": r["calls"],
                 "tokens": r["tokens"], "cost_aud": r["cost_aud"]}
                for r in (s.get("by_plugin") or [])
            ],
            "byDay": [
                {"date": r["date"], "calls": r["calls"],
                 "tokens": r["tokens"], "cost_aud": r["cost_aud"]}
                for r in (s.get("by_day") or [])
            ],
        })
    except Exception:
        return ok({"todayCost": 0, "monthlyCost": 0, "monthlyBudget": 100,
                   "totalCalls": 0, "totalTokensIn": 0, "totalTokensOut": 0,
                   "byPlugin": [], "byDay": []})


# ── Settings ───────────────────────────────────────────────────────────────────
@app.route("/api/settings")
def get_settings():
    s = get_all_settings()
    # Mask sensitive keys
    for key in ("anthropic_api_key", "fusesign_api_key", "teams_webhook_url",
                "statementhub_api_key",
                "xero_client_secret", "xero_access_token", "xero_refresh_token"):
        if s.get(key):
            s[key] = s[key][:4] + "••••••••••••••••••••"
    s["fast_model"] = get_claude_model_fast()
    s["reasoning_model"] = get_claude_model_reasoning()
    # Add Xero OAuth status
    try:
        from xero_oauth import is_configured as xero_is_configured, is_authorised as xero_is_authorised
        s["xero_configured"] = xero_is_configured()
        s["xero_authorised"] = xero_is_authorised()
    except Exception:
        s["xero_configured"] = False
        s["xero_authorised"] = False
    return ok(s)


@app.route("/api/settings", methods=["POST"])
def save_settings():
    body = request.get_json(silent=True) or {}
    # Only save non-masked values
    safe_keys = {
        "anthropic_api_key", "outlook_email",
        "fusesign_api_key", "teams_webhook_url",
        "statementhub_api_key", "statementhub_base_url",
        "confidence_threshold", "heartbeat_interval_seconds",
        "draft_mode", "auto_update_enabled",
        "fast_model", "reasoning_model",
        "monthly_ai_budget_aud",
        "skip_public_holidays", "public_holiday_state",
        "reception_mode", "staff_profile",
        # Xero OAuth credentials
        "xero_client_id", "xero_client_secret",
    }
    saved = []
    for key, value in body.items():
        if key in safe_keys and "••••" not in str(value):
            set_setting(key, str(value))
            saved.append(key)
    # Re-detect models if API key changed
    if "anthropic_api_key" in saved:
        try:
            update_claude_models()
        except Exception:
            pass
    return ok({"saved": saved})


@app.route("/api/settings/test/<service>", methods=["POST"])
def test_connection(service):
    try:
        from gateway_client import GatewayClient
        gw = GatewayClient()
        gw.load()
        if service == "xpm":
            result = gw.xpm.list_clients(page=1, page_size=1)
            return ok({"connected": True, "service": "xpm"})
        elif service == "fusesign":
            result = gw.fusesign.list_envelopes(limit=1)
            return ok({"connected": True, "service": "fusesign"})
        elif service == "teams":
            gw.teams.send_alert("CoWorker", "Connection test successful ✅")
            return ok({"connected": True, "service": "teams"})
        else:
            return err(f"Unknown service: {service}")
    except Exception as e:
        return ok({"connected": False, "service": service, "error": str(e)})


# ── Xero OAuth ─────────────────────────────────────────────────────────────────

# Background thread state for OAuth flow
_xero_auth_thread: threading.Thread | None = None
_xero_auth_status: dict = {"status": "idle", "message": ""}


@app.route("/api/xero/status")
def xero_status():
    """Return current Xero OAuth status."""
    try:
        from xero_oauth import is_configured, is_authorised, get_tenant_id
        return ok({
            "configured":  is_configured(),
            "authorised":  is_authorised(),
            "tenant_id":   get_tenant_id(),
            "client_id":   get_setting("xero_client_id", ""),
            "auth_status": _xero_auth_status,
        })
    except Exception as e:
        return err(str(e))


@app.route("/api/xero/start-auth", methods=["POST"])
def xero_start_auth():
    """
    Start the Xero OAuth Authorization Code flow in a background thread.
    Opens the user's browser to the Xero login page.
    The callback is handled by /oauth/callback below.
    """
    global _xero_auth_thread, _xero_auth_status

    if _xero_auth_thread and _xero_auth_thread.is_alive():
        return ok({"status": "in_progress", "message": "OAuth flow already in progress"})

    _xero_auth_status = {"status": "in_progress", "message": "Opening Xero login..."}

    def _run_auth():
        global _xero_auth_status
        try:
            from xero_oauth import start_auth_flow
            token_data = start_auth_flow(timeout=300)
            _xero_auth_status = {
                "status":  "success",
                "message": "Xero connected successfully!",
                "scope":   token_data.get("scope", ""),
            }
        except Exception as e:
            _xero_auth_status = {"status": "error", "message": str(e)}

    _xero_auth_thread = threading.Thread(target=_run_auth, daemon=True)
    _xero_auth_thread.start()

    return ok({"status": "in_progress", "message": "Xero login page opened in browser"})


@app.route("/oauth/callback")
def xero_oauth_callback():
    """
    Xero OAuth callback — receives the authorization code after user login.
    Passes the code to the waiting start_auth_flow() thread.
    """
    code  = request.args.get("code", "")
    state = request.args.get("state", "")
    error = request.args.get("error", "")

    try:
        from xero_oauth import notify_oauth_callback
        notify_oauth_callback(code=code, state=state, error=error)
    except Exception as e:
        return f"<h1>Error</h1><p>{e}</p>", 500

    if error:
        return (
            "<!DOCTYPE html><html><head><title>Xero — Error</title>"
            "<style>body{font-family:sans-serif;text-align:center;padding:60px}</style></head>"
            f"<body><h1>\u274c Xero Connection Failed</h1><p>{error}</p>"
            "<p>You can close this window.</p></body></html>"
        ), 400

    return (
        "<!DOCTYPE html><html><head><title>MCS CoWorker — Xero Connected</title>"
        "<style>body{font-family:sans-serif;text-align:center;padding:60px;background:#f0f4f8}"
        "h1{color:#13b5ea}p{color:#444}</style></head>"
        "<body><h1>&#10003; Xero Connected!</h1>"
        "<p>MCS CoWorker is now connected to Xero XPM.</p>"
        "<p>You can close this window and return to CoWorker.</p></body></html>"
    )


@app.route("/api/xero/disconnect", methods=["POST"])
def xero_disconnect():
    """Revoke the Xero refresh token and clear all stored tokens."""
    try:
        from xero_oauth import revoke_token
        revoke_token()
        return ok({"disconnected": True})
    except Exception as e:
        return err(str(e))


# ── Chat ───────────────────────────────────────────────────────────────────────
@app.route("/api/chat", methods=["POST"])
def chat():
    body = request.get_json(silent=True) or {}
    messages = body.get("messages", [])
    if not messages:
        return err("No messages provided")

    try:
        import anthropic as anthropic_lib
        api_key = get_setting("anthropic_api_key", "")
        if not api_key:
            return err("Anthropic API key not configured")

        client = anthropic_lib.Anthropic(api_key=api_key)

        # Detect tier
        last_msg = messages[-1].get("content", "").lower()
        tier2_keywords = ["xpm", "fusesign", "teams", "memory", "workflow",
                          "report", "wip", "debtor", "engagement", "onboard",
                          "gateway", "event", "heartbeat", "kpi"]
        is_tier2 = any(k in last_msg for k in tier2_keywords)
        model = get_claude_model_reasoning() if is_tier2 else get_claude_model_fast()

        # Build system prompt
        system = _build_chat_system_prompt()

        response = client.messages.create(
            model=model,
            max_tokens=4096 if is_tier2 else 2048,
            system=system,
            messages=messages,
        )
        return ok({
            "content": response.content[0].text,
            "model": model,
            "tier": 2 if is_tier2 else 1,
            "usage": {
                "input_tokens": response.usage.input_tokens,
                "output_tokens": response.usage.output_tokens,
            }
        })
    except Exception as e:
        return err(f"Chat error: {str(e)}")


# ── System ─────────────────────────────────────────────────────────────────────
@app.route("/api/system/status")
def system_status():
    uptime_secs = int((datetime.now() - _start_time).total_seconds())
    h, rem = divmod(uptime_secs, 3600)
    m, _ = divmod(rem, 60)
    try:
        from memory_store import MemoryStore
        mem_count = MemoryStore().count()
    except Exception:
        mem_count = 0
    try:
        from token_meter import get_usage_summary
        cost_today = f"${get_usage_summary().get('today_cost_aud', 0):.2f}"
    except Exception:
        cost_today = "$0.00"
    try:
        from event_bus import EventBus
        tick = len(EventBus.get_history(10000))
    except Exception:
        tick = 0

    return ok({
        "heartbeat": "Active" if _loader and _loader._running else "Stopped",
        "heartbeatTick": tick,
        "fastModel": get_claude_model_fast(),
        "reasoningModel": get_claude_model_reasoning(),
        "memoryRecords": mem_count,
        "costToday": cost_today,
        "uptime": f"{h}h {m}m",
        "version": get_setting("app_version", "2.4.1"),
        "updateAvailable": False,
    })


# ── Helpers ────────────────────────────────────────────────────────────────────
def _format_dt(dt) -> str:
    if dt is None:
        return "Never"
    if isinstance(dt, datetime):
        now = datetime.now()
        diff = now - dt
        if diff.total_seconds() < 60:
            return "Just now"
        if diff.total_seconds() < 3600:
            return f"{int(diff.total_seconds() // 60)} min ago"
        if diff.total_seconds() < 86400:
            return f"{int(diff.total_seconds() // 3600)} hr ago"
        return dt.strftime("%d %b")
    return str(dt)


def _format_time(ts: str) -> str:
    try:
        dt = datetime.fromisoformat(ts)
        return dt.strftime("%H:%M")
    except Exception:
        return ts[:5] if ts else ""


def _format_next(lp) -> str:
    if not lp.enabled:
        return "Disabled"
    if lp.schedule_seconds == 0:
        return "On event"
    if lp.next_run is None:
        return "Soon"
    diff = (lp.next_run - datetime.now()).total_seconds()
    if diff <= 0:
        return "Now"
    if diff < 60:
        return f"{int(diff)}s"
    if diff < 3600:
        return f"{int(diff // 60)} min"
    return f"{int(diff // 3600)} hr"


def _schedule_label(lp) -> str:
    s = lp.schedule_seconds
    if s == 0:
        return "Event-driven"
    if s < 3600:
        return f"Every {s // 60} minutes"
    if s < 86400:
        return f"Every {s // 3600} hours"
    return "Daily"


def _build_chat_system_prompt() -> str:
    style = config.get_style_preferences() if hasattr(config, "get_style_preferences") else ""

    # Inject live practice context so the assistant can answer operational questions
    context_lines = []
    try:
        pending_count = _approval_queue.count_pending() if _approval_queue else 0
        context_lines.append(f"Pending approvals in queue: {pending_count}")
        if _approval_queue and pending_count > 0:
            pending_items = _approval_queue.list_pending(limit=5)
            for item in pending_items:
                desc = item.get("description", item.get("action_type", "?"))
                plugin = item.get("plugin_id", "?")
                context_lines.append(f"  - [{plugin}] {desc[:80]}")
    except Exception:
        pass
    try:
        from plugins.plugin_asic_returns import get_asic_returns
        asic_open = [r for r in get_asic_returns(limit=200)
                     if r.get("status") not in ("completed", "cancelled")]
        context_lines.append(f"Open ASIC annual returns: {len(asic_open)}")
        awaiting_solvency = [r for r in asic_open if r.get("status") == "awaiting_solvency"]
        overdue_asic = [r for r in asic_open if r.get("status") == "overdue"]
        if awaiting_solvency:
            context_lines.append(f"  Awaiting solvency resolution: {len(awaiting_solvency)} "
                                  f"({', '.join(r.get('company_name','?') for r in awaiting_solvency[:3])})")
        if overdue_asic:
            context_lines.append(f"  Overdue ASIC returns: {len(overdue_asic)} "
                                  f"({', '.join(r.get('company_name','?') for r in overdue_asic[:3])})")
        elif asic_open:
            names = ", ".join(r.get("company_name", "?") for r in asic_open[:5])
            context_lines.append(f"  Companies: {names}{'...' if len(asic_open) > 5 else ''}")
    except Exception:
        pass
    try:
        lessons = get_active_lessons()
        if lessons:
            lesson_text = "; ".join(l["lesson"] for l in lessons[:10])
            context_lines.append(f"Learned preferences: {lesson_text}")
    except Exception:
        pass
    try:
        recent = get_recent_activity(limit=10)
        if recent:
            activity_text = "; ".join(
                f"{r.get('classification','?')}: {r.get('subject','?')[:40]}"
                for r in recent
            )
            context_lines.append(f"Recent activity (last 10): {activity_text}")
    except Exception:
        pass
    try:
        # Overdue debtors summary from XPM invoices
        from gateway_client import XPMClient
        xpm = XPMClient()
        if xpm.is_available():
            summary = xpm.get_debtor_summary()
            total = summary.get("total_outstanding", 0)
            count = summary.get("invoice_count", 0)
            over90 = summary.get("90_plus", 0)
            context_lines.append(
                f"Debtor summary: {count} outstanding invoices totalling ${total:,.0f} "
                f"(${over90:,.0f} is 90+ days overdue)"
            )
    except Exception:
        pass

    practice_context = ("\n".join(context_lines)) if context_lines else ""
    practice_name = get_setting("practice_name", "MC & S")

    return f"""You are the AI assistant built into {practice_name} CoWorker, an intelligent automation platform for an Australian accounting firm.
You can answer questions about the practice's current state, explain what plugins do, help build new automation, and advise on workflow improvements.

CURRENT PRACTICE STATE
{practice_context if practice_context else '(No live data available yet)'}

PLUGIN DEVELOPMENT
TIER 1 — Template Builder: for common email/auto-reply patterns
TIER 2 — Custom Plugin Writer: full Python plugins using:
  - context.claude_fast / context.claude_reason (dual Claude models)
  - context.memory (ChromaDB vector store)
  - context.event_bus (publish/subscribe events)
  - context.gateway.xpm (XPM practice management)
  - context.gateway.fusesign (document signing)
  - context.gateway.teams (Teams notifications)
  - context.approval_queue (confidence-based human review)

Always produce working Python code. Use PluginResult(success=True/False, message="...").
{f'Style preferences: {style}' if style else ''}"""


# ── ASIC Tracker ─────────────────────────────────────────────────────────────
@app.route("/api/asic")
def list_asic_returns():
    """List all ASIC annual return records."""
    try:
        from plugins.plugin_asic_returns import get_asic_returns
        status = request.args.get("status", None)
        limit  = int(request.args.get("limit", 100))
        rows = get_asic_returns(status=status, limit=limit)
        return ok(rows)
    except Exception as e:
        return err(str(e))


@app.route("/api/asic/<int:return_id>/mark-paid", methods=["POST"])
def asic_mark_paid(return_id):
    """Mark an ASIC return as paid and update status to 'completed' if solvency also signed."""
    try:
        from plugins.plugin_asic_returns import update_asic_return, get_asic_returns
        rows = get_asic_returns(limit=1000)
        record = next((r for r in rows if r["id"] == return_id), None)
        if not record:
            return err("ASIC return not found", 404)
        new_status = "completed" if record.get("solvency_signed") else "awaiting_solvency"
        update_asic_return(return_id, asic_paid=1, status=new_status)
        return ok({"id": return_id, "asic_paid": True, "status": new_status})
    except Exception as e:
        return err(str(e))


@app.route("/api/asic/<int:return_id>/mark-solvency-signed", methods=["POST"])
def asic_mark_solvency_signed(return_id):
    """Mark an ASIC return's solvency resolution as signed and update status."""
    try:
        from plugins.plugin_asic_returns import update_asic_return, get_asic_returns
        rows = get_asic_returns(limit=1000)
        record = next((r for r in rows if r["id"] == return_id), None)
        if not record:
            return err("ASIC return not found", 404)
        new_status = "completed" if record.get("asic_paid") else "awaiting_payment"
        update_asic_return(return_id, solvency_signed=1, status=new_status)
        return ok({"id": return_id, "solvency_signed": True, "status": new_status})
    except Exception as e:
        return err(str(e))


@app.route("/api/asic/<int:return_id>/notes", methods=["POST"])
def asic_update_notes(return_id):
    """Update the notes field on an ASIC return."""
    try:
        from plugins.plugin_asic_returns import update_asic_return
        body = request.get_json(silent=True) or {}
        notes = body.get("notes", "")
        update_asic_return(return_id, notes=notes)
        return ok({"id": return_id, "notes": notes})
    except Exception as e:
        return err(str(e))


# ── Email Rules ───────────────────────────────────────────────────────────────
@app.route("/api/rules")
def list_rules():
    return ok(get_rules())

@app.route("/api/rules", methods=["POST"])
def create_rule():
    data = request.get_json(force=True)
    save_rule(data)
    return ok(get_rules())

@app.route("/api/rules/<int:rule_id>", methods=["DELETE"])
def remove_rule(rule_id):
    delete_rule(rule_id)
    return ok()

# ── Staff ─────────────────────────────────────────────────────────────────────
@app.route("/api/staff")
def list_staff():
    return ok(get_staff())

@app.route("/api/staff", methods=["POST"])
def create_staff():
    data = request.get_json(force=True)
    save_staff(data)
    return ok(get_staff())

@app.route("/api/staff/<int:staff_id>", methods=["DELETE"])
def remove_staff(staff_id):
    delete_staff(staff_id)
    return ok()

# ── Links & Forms ─────────────────────────────────────────────────────────────
@app.route("/api/links")
def list_links():
    return ok(get_links())

@app.route("/api/links", methods=["POST"])
def create_link():
    data = request.get_json(force=True)
    save_link(data)
    return ok(get_links())

@app.route("/api/links/<int:link_id>", methods=["DELETE"])
def remove_link(link_id):
    delete_link(link_id)
    return ok()

# ── Plugin management extras ──────────────────────────────────────────────────
@app.route("/api/plugins/<plugin_id>/disable", methods=["POST"])
@require_loader
def disable_plugin(plugin_id):
    lp = _loader.get_plugin(plugin_id)
    if not lp:
        return err("Plugin not found", 404)
    lp.enabled = False
    from config import save_plugin_state
    save_plugin_state(plugin_id, enabled=False)
    return ok({"enabled": False})

@app.route("/api/plugins/<plugin_id>", methods=["DELETE"])
@require_loader
def delete_plugin(plugin_id):
    lp = _loader.get_plugin(plugin_id)
    if not lp:
        return err("Plugin not found", 404)
    import os
    try:
        os.remove(lp.path)
        _loader.unload_plugin(plugin_id)
    except Exception as e:
        return err(str(e))
    return ok()

# ── Lessons ───────────────────────────────────────────────────────────────────
@app.route("/api/lessons")
def list_lessons():
    return ok(get_active_lessons())

@app.route("/api/lessons", methods=["POST"])
def create_lesson():
    data = request.get_json(force=True)
    add_lesson(data.get("lesson", ""), data.get("source", ""))
    return ok(get_active_lessons())

@app.route("/api/lessons/<int:lesson_id>", methods=["DELETE"])
def remove_lesson(lesson_id):
    delete_lesson(lesson_id)
    return ok()

@app.route("/api/lessons/<int:lesson_id>/toggle", methods=["POST"])
def toggle_lesson_route(lesson_id):
    data = request.get_json(force=True)
    toggle_lesson(lesson_id, data.get("active", True))
    return ok()

# ── Style preferences ─────────────────────────────────────────────────────────
@app.route("/api/style")
def get_style():
    return ok({"content": get_style_preferences()})

@app.route("/api/style", methods=["POST"])
def save_style():
    data = request.get_json(force=True)
    save_style_preferences(data.get("content", ""))
    return ok()

# ── Chat history ──────────────────────────────────────────────────────────────
@app.route("/api/chat/history")
def chat_history():
    return ok(get_feedback_history(200))

@app.route("/api/chat/history", methods=["DELETE"])
def clear_chat_history():
    clear_feedback_history()
    return ok()

# ── Microsoft OAuth callback ─────────────────────────────────────────────────
@app.route("/auth/callback")
def auth_callback():
    """Receives the Microsoft OAuth2 redirect and passes the code to GraphClient."""
    code = request.args.get("code")
    error = request.args.get("error")
    if _graph_client is not None:
        _graph_client.receive_auth_code(code=code, error=error)
    if error:
        html = f"""<html><body style='font-family:Arial;text-align:center;padding:60px'>
<h2 style='color:#c0392b'>Authentication Failed</h2>
<p>{error}</p>
<p>You can close this tab.</p>
</body></html>"""
        return html, 400
    html = """<html><body style='font-family:Arial;text-align:center;padding:60px'>
<h2 style='color:#2E7D32'>&#10003; Authentication Successful</h2>
<p>You can close this tab and return to MC&amp;S CoWorker.</p>
<script>setTimeout(function(){window.close();},2000);</script>
</body></html>"""
    return html, 200


# ── Frontend static file serving ──────────────────────────────────────────────────────────────
# Serves the built React app (frontend_dist/) when running in installed mode.

def _frontend_dir() -> str | None:
    candidate = os.path.join(os.path.dirname(os.path.abspath(__file__)), "frontend_dist")
    return candidate if os.path.isdir(candidate) else None


@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_frontend(path):
    """Serve the React SPA. Non-API paths return index.html (SPA routing)."""
    frontend = _frontend_dir()
    if frontend is None:
        return jsonify({"error": "Frontend not built."}), 404
    # Serve real static assets (JS, CSS, images, etc.)
    if path and os.path.isfile(os.path.join(frontend, path)):
        return send_from_directory(frontend, path)
    # All other paths → index.html (React Router handles routing client-side)
    return send_from_directory(frontend, "index.html")


def run_server(host="127.0.0.1", port=API_PORT, debug=False):
    """Start the Flask server in a background thread."""
    app.run(host=host, port=port, debug=debug, use_reloader=False, threaded=True)


def start_in_thread(host="127.0.0.1", port=API_PORT):
    # Wire SSE broadcaster to EventBus before starting
    try:
        _wire_sse_to_event_bus()
    except Exception:
        pass  # EventBus may not be ready yet — wiring happens lazily
    t = threading.Thread(target=run_server, args=(host, port), daemon=True)
    t.start()
    return t
