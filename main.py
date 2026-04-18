"""
MC & S CoWorker — Desktop Entry Point (pywebview)
Replaces the Tkinter app.py shell.
Starts Flask API on localhost:7842, then opens a native window via pywebview.
All Python backend modules (plugin_loader, config, graph_client, etc.) remain unchanged.
"""
from __future__ import annotations

# ── Force UTF-8 output on Windows (prevents UnicodeEncodeError with emoji) ───
import sys as _sys
if hasattr(_sys.stdout, "reconfigure"):
    try:
        _sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        _sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import os
import sys
import time
import threading
import logging
from pathlib import Path

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("coworker")

# ── Resolve paths ──────────────────────────────────────────────────────────────
# When bundled with PyInstaller, sys._MEIPASS is the temp extraction dir
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)
    APP_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent
    APP_DIR = BASE_DIR

FRONTEND_DIR = BASE_DIR / "frontend_dist"

# Ensure the app directory is on sys.path so sibling modules (config, graph_client,
# plugin_loader, etc.) can be imported regardless of how the process was started.
# This is required when launched via the VBScript launcher from the installer.
for _p in [str(BASE_DIR), str(APP_DIR)]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

# ── Initialise database ────────────────────────────────────────────────────────
import config
config.init_db()

# ── Initialise plugin loader ───────────────────────────────────────────────────
from graph_client import GraphClient
from plugin_loader import PluginLoader
from approval_queue import ApprovalQueue
from kpi_monitor import KPIMonitor
import api_server

def _init_backend():
    """Initialise all backend services and wire them into the API server."""
    log.info("Initialising backend services...")

    # Graph client (Microsoft 365)
    graph = GraphClient()
    try:
        graph.authenticate()
        log.info("Microsoft Graph authenticated")
    except Exception as e:
        log.warning(f"Graph auth failed (will retry): {e}")

    # Plugin loader
    loader = PluginLoader()
    loader.set_graph(graph)
    loader.set_claude()
    loader.load_all()
    loader.start_scheduler()
    log.info(f"Plugin loader started — {len(loader.get_plugins())} plugins loaded")

    # Approval queue
    aq = ApprovalQueue()
    log.info("Approval queue ready")

    # KPI monitor
    km = KPIMonitor  # already a singleton instance, not a class
    log.info("KPI monitor ready")

    # Wire into API server
    api_server.set_loader(loader)
    api_server.set_approval_queue(aq)
    api_server.set_kpi_monitor(km)

    return loader, aq, km


def _wait_for_server(host="127.0.0.1", port=7842, timeout=10.0):
    """Block until the Flask server is accepting connections."""
    import socket
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.1)
    return False


def _get_frontend_url() -> str:
    """Return the URL to load in the webview window."""
    # If bundled frontend exists, serve it via Flask static
    if FRONTEND_DIR.exists():
        return "http://127.0.0.1:7842/"
    # Dev mode: Vite dev server
    return "http://127.0.0.1:3000/"


def main():
    log.info("MC & S CoWorker starting...")

    # 1. Start Flask API server in background thread
    log.info("Starting API server on port 7842...")
    api_server.start_in_thread(host="127.0.0.1", port=7842)

    # 2. Initialise backend in background (don't block the window)
    backend_thread = threading.Thread(target=_init_backend, daemon=True)
    backend_thread.start()

    # 3. Wait for Flask to be ready
    if not _wait_for_server(timeout=10):
        log.error("API server failed to start within 10 seconds")
        sys.exit(1)
    log.info("API server ready")

    # 4. Open pywebview window
    try:
        import webview
    except ImportError:
        log.error("pywebview not installed. Run: pip install pywebview")
        sys.exit(1)

    url = _get_frontend_url()
    log.info(f"Opening window: {url}")

    window = webview.create_window(
        title="MC & S CoWorker",
        url=url,
        width=1280,
        height=800,
        min_size=(1024, 640),
        resizable=True,
        text_select=False,
        confirm_close=False,
    )

    # Inject a flag so the React app knows it's in desktop mode
    def on_loaded():
        try:
            window.evaluate_js("window.__pywebview__ = true;")
        except Exception:
            pass

    window.events.loaded += on_loaded

    # Intercept navigation: keep 127.0.0.1:7842 inside the webview,
    # open everything else (OAuth redirects, external links) in the real browser.
    def on_navigating(url: str):
        import webbrowser as _wb
        if "127.0.0.1:7842" in url or "localhost:7842" in url:
            return  # allow — this is our own app
        # External URL or OAuth redirect — open in default browser
        _wb.open(url)
        return False  # cancel navigation inside webview

    try:
        window.events.before_load += on_navigating
    except AttributeError:
        pass  # older pywebview versions may not have before_load

    # Start webview (blocks until window is closed)
    webview.start(debug=False, http_server=False)
    log.info("Window closed — shutting down")


if __name__ == "__main__":
    main()
