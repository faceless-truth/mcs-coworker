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
from flask import Flask, Response, jsonify, request, stream_with_context
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
)
from plugin_loader import PluginLoader
from approval_queue import ApprovalQueue
from token_meter import get_usage_summary
from event_bus import EventBus
from kpi_monitor import KPIMonitor

# ── App setup ──────────────────────────────────────────────────────────────────
app = Flask(__name__)
CORS(app, origins=["http://localhost:7842", "http://127.0.0.1:7842",
                   "http://localhost:3000", "http://127.0.0.1:3000"])

API_PORT = 7842

# Shared state — populated by main.py on startup
_loader: PluginLoader | None = None
_approval_queue: ApprovalQueue | None = None
_kpi_monitor: KPIMonitor | None = None
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
            "name": inst.NAME,
            "description": getattr(inst, "DESCRIPTION", ""),
            "enabled": lp.enabled,
            "status": "running" if lp.running else ("disabled" if not lp.enabled else "idle"),
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
    _approval_queue.approve(action_id)
    return ok({"action_id": action_id, "approved": True})


@app.route("/api/approvals/<action_id>/reject", methods=["POST"])
def reject_action(action_id):
    if _approval_queue is None:
        return err("Approval queue not initialised", 503)
    _approval_queue.reject(action_id)
    return ok({"action_id": action_id, "rejected": True})


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
    for key in ("anthropic_api_key", "xpm_api_key", "fusesign_api_key", "teams_webhook_url"):
        if s.get(key):
            s[key] = s[key][:8] + "••••••••••••••••••••••••"
    s["fast_model"] = get_claude_model_fast()
    s["reasoning_model"] = get_claude_model_reasoning()
    return ok(s)


@app.route("/api/settings", methods=["POST"])
def save_settings():
    body = request.get_json(silent=True) or {}
    # Only save non-masked values
    safe_keys = {
        "anthropic_api_key", "outlook_email", "xpm_api_key",
        "fusesign_api_key", "teams_webhook_url",
        "confidence_threshold", "heartbeat_interval_seconds",
        "draft_mode", "auto_update_enabled",
        "fast_model", "reasoning_model",
        "monthly_ai_budget_aud",
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
            result = gw.xpm.list_clients(limit=1)
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
        "heartbeat": "Active" if _loader and _loader._scheduler_running else "Stopped",
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
    return f"""You are an autonomous AI automation engineer built into MC & S CoWorker.
You help build and manage automation plugins for an accounting firm.

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
{f"Style preferences: {style}" if style else ""}"""


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
