"""
MC & S CoWorker — Desktop Entry Point (pywebview)
Starts Flask API on localhost:7842, then opens a native window via pywebview.
Minimises to system tray when closed. Registers itself to start with Windows.
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
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)
    APP_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent
    APP_DIR = BASE_DIR

FRONTEND_DIR = BASE_DIR / "frontend_dist"
INSTALL_DIR = APP_DIR.parent  # one level up from app/

# Ensure the app directory is on sys.path
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

    graph = GraphClient()
    try:
        graph.authenticate()
        log.info("Microsoft Graph authenticated")
    except Exception as e:
        log.warning(f"Graph auth failed (will retry): {e}")

    loader = PluginLoader()
    loader.set_graph(graph)
    loader.set_claude()
    loader.load_all()
    loader.start_scheduler()
    log.info(f"Plugin loader started — {len(loader.get_plugins())} plugins loaded")

    aq = ApprovalQueue()
    log.info("Approval queue ready")

    km = KPIMonitor  # singleton instance
    log.info("KPI monitor ready")

    api_server.set_loader(loader)
    api_server.set_approval_queue(aq)
    api_server.set_kpi_monitor(km)
    api_server.set_graph_client(graph)

    return loader, aq, km


def _wait_for_server(host="127.0.0.1", port=7842, timeout=10.0):
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
    if FRONTEND_DIR.exists():
        return "http://127.0.0.1:7842/"
    return "http://127.0.0.1:3000/"


def _register_startup():
    """Add MCS CoWorker to Windows startup registry (HKCU — no admin needed)."""
    try:
        import winreg
        # Find the VBS launcher path
        vbs = INSTALL_DIR / "MCSCoWorker.vbs"
        if not vbs.exists():
            return
        cmd = f'wscript.exe "{vbs}"'
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(key, "MCSCoWorker", 0, winreg.REG_SZ, cmd)
        winreg.CloseKey(key)
        log.info("Registered in Windows startup")
    except Exception as e:
        log.warning(f"Could not register startup: {e}")


def _build_tray_icon():
    """Build a pystray Icon using the app's icon.ico, or a generated fallback."""
    try:
        from PIL import Image as PILImage
        icon_path = INSTALL_DIR / "icon.ico"
        if icon_path.exists():
            img = PILImage.open(str(icon_path))
            # pystray needs RGBA
            img = img.convert("RGBA")
        else:
            # Fallback: simple navy square with white "M"
            img = PILImage.new("RGBA", (64, 64), (15, 52, 96, 255))
        return img
    except Exception as e:
        log.warning(f"Could not load tray icon: {e}")
        return None


def main():
    log.info("MC & S CoWorker starting...")

    # Register in Windows startup (silent, HKCU only)
    _register_startup()

    # 1. Start Flask API server
    log.info("Starting API server on port 7842...")
    api_server.start_in_thread(host="127.0.0.1", port=7842)

    # 2. Initialise backend in background
    backend_thread = threading.Thread(target=_init_backend, daemon=True)
    backend_thread.start()

    # 3. Wait for Flask
    if not _wait_for_server(timeout=10):
        log.error("API server failed to start within 10 seconds")
        sys.exit(1)
    log.info("API server ready")

    # 4. Open pywebview window
    try:
        import webview
    except ImportError:
        log.error("pywebview not installed.")
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

    # Inject desktop flag for React app
    def on_loaded():
        try:
            window.evaluate_js("window.__pywebview__ = true;")
        except Exception:
            pass

    window.events.loaded += on_loaded

    # Open external links in real browser
    try:
        import webview as _wv
        _wv.settings['OPEN_EXTERNAL_LINKS_IN_BROWSER'] = True
    except Exception:
        pass

    # ── System tray setup ──────────────────────────────────────────────────────
    _quit_event = threading.Event()

    def _show_window(icon=None, item=None):
        """Restore and focus the window from the tray."""
        try:
            window.show()
        except Exception:
            pass

    def _quit_app(icon=None, item=None):
        """Quit the app entirely from the tray menu."""
        _quit_event.set()
        try:
            if icon:
                icon.stop()
        except Exception:
            pass
        try:
            webview.windows[0].destroy()
        except Exception:
            pass

    tray_icon = None
    try:
        import pystray
        from pystray import MenuItem as TrayItem, Menu as TrayMenu

        tray_img = _build_tray_icon()
        if tray_img:
            tray_icon = pystray.Icon(
                "MCSCoWorker",
                tray_img,
                "MC & S CoWorker",
                menu=TrayMenu(
                    TrayItem("Open MC & S CoWorker", _show_window, default=True),
                    TrayMenu.SEPARATOR,
                    TrayItem("Quit", _quit_app),
                )
            )

            # Run tray in its own thread
            tray_thread = threading.Thread(target=tray_icon.run, daemon=True)
            tray_thread.start()
            log.info("System tray icon active")

    except ImportError:
        log.warning("pystray not installed — tray icon unavailable")
    except Exception as e:
        log.warning(f"Tray icon failed: {e}")

    # Override close: hide to tray instead of quitting
    def on_closing():
        """Called when user clicks X — hide to tray instead of closing."""
        if tray_icon is not None:
            try:
                window.hide()
                return False  # prevent default close
            except Exception:
                pass
        return True  # no tray — allow close

    window.events.closing += on_closing

    # Start webview (blocks until window is destroyed)
    webview.start(debug=False, http_server=False)

    # If we get here and quit wasn't requested, it means window was destroyed
    # (e.g. via _quit_app). Stop tray if still running.
    if tray_icon is not None and not _quit_event.is_set():
        try:
            tray_icon.stop()
        except Exception:
            pass

    log.info("MC & S CoWorker shut down")


if __name__ == "__main__":
    main()
