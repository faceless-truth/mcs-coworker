"""
MC & S CoWorker — Desktop Entry Point (pywebview)
Starts Flask API on localhost:7842, then opens a native window via pywebview.
Minimises to system tray when closed. Registers itself to start with Windows.
"""
from __future__ import annotations

# ── Emergency crash catcher (must be first — before any imports can fail) ─────
import sys as _sys, traceback as _tb, os as _os
from pathlib import Path as _Path
try:
    _startup_log = _Path(__file__).parent.parent / "data" / "startup_error.log"
    _startup_log.parent.mkdir(parents=True, exist_ok=True)
except Exception:
    _startup_log = _Path(_os.environ.get("TEMP", ".")) / "mcs_startup_error.log"

def _log_fatal(exc_type, exc_val, exc_tb):
    try:
        from datetime import datetime
        with open(_startup_log, "a", encoding="utf-8") as f:
            f.write(f"\n=== FATAL {datetime.now().isoformat()} ===\n")
            _tb.print_exception(exc_type, exc_val, exc_tb, file=f)
    except Exception:
        pass

_sys.excepthook = _log_fatal
# ── End crash catcher ─────────────────────────────────────────────────────────

# ── Force UTF-8 output on Windows (prevents UnicodeEncodeError with emoji) ───
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

# Force EdgeChromium backend — avoids .NET/pythonnet init failures in embeddable Python
os.environ.setdefault("PYWEBVIEW_GUI", "edgechromium")

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

# Generate the per-install API token (idempotent — only writes on first run).
# The token gates all /api/* routes; it is injected into the webview below so
# the React frontend can authenticate without prompting the user.
API_TOKEN = config.ensure_api_token()

# ── Initialise plugin loader ───────────────────────────────────────────────────
from graph_client import GraphClient
from plugin_loader import PluginLoader
from approval_queue import ApprovalQueue
from kpi_monitor import KPIMonitor
import api_server


def _has_internet(timeout: float = 3.0) -> bool:
    """Quick connectivity probe used to skip blocking auth flows on offline
    first-launch (rural/temporary outages, train trips, fresh-imaged laptops
    sitting on captive Wi-Fi portals)."""
    import urllib.request
    try:
        urllib.request.urlopen("https://login.microsoftonline.com", timeout=timeout)
        return True
    except Exception:
        return False


def _init_backend():
    """Initialise all backend services and wire them into the API server."""
    log.info("Initialising backend services...")

    graph = GraphClient()
    if not _has_internet():
        log.warning("No internet connection detected — skipping Microsoft authentication")
        # Surfaced via /api/system/status so the frontend can show a banner.
        config.set_setting("_offline_mode", "true")
    else:
        config.set_setting("_offline_mode", "false")
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

    # Sanity check the SharePoint endpoints share the loader's authenticated
    # GraphClient. If these IDs ever diverge, plugins and SharePoint will
    # disagree about whether the user is signed in.
    log.info(
        f"GraphClient wiring: loader={id(loader._graph)} "
        f"api_server={id(graph)} same={loader._graph is graph}"
    )

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
    log.warning(
        "frontend_dist not found at %s — falling back to Vite dev server (port 3000). "
        "This will show a blank window on production installs.", FRONTEND_DIR
    )
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


# Stable identifier so Windows groups our windows under our own taskbar entry
# (and shows the icon set via WM_SETICON) instead of grouping under pythonw.exe.
APP_USER_MODEL_ID = "MCandS.CoWorker.Desktop.1"
WINDOW_TITLE = "MC & S CoWorker"


def _icon_path() -> Path | None:
    """Resolve the bundled icon.ico across dev and production layouts.

    Production: {install_dir}\\assets\\icon.ico (installer copies assets/* there).
    Dev:        {repo}\\assets\\icon.ico.
    """
    for cand in (
        INSTALL_DIR / "assets" / "icon.ico",
        BASE_DIR / "assets" / "icon.ico",
    ):
        if cand.exists():
            return cand
    return None


def _set_app_user_model_id():
    """Tell Windows we're a distinct app, not just another pythonw.exe instance.
    Must be called before any window is shown for the taskbar grouping to apply."""
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception as e:
        log.debug(f"Could not set AppUserModelID: {e}")


def _apply_window_icon():
    """Set the pywebview window's title-bar and taskbar icon via Win32 WM_SETICON.
    pywebview's edgechromium backend doesn't expose an icon parameter, so we
    locate the HWND by title after the page has loaded and push the icon in."""
    ico = _icon_path()
    if ico is None:
        log.warning("No icon.ico found — taskbar/title-bar will use default")
        return
    try:
        import ctypes
        from ctypes import wintypes

        WM_SETICON      = 0x0080
        ICON_SMALL      = 0
        ICON_BIG        = 1
        IMAGE_ICON      = 1
        LR_LOADFROMFILE = 0x0010

        user32 = ctypes.windll.user32
        user32.LoadImageW.restype = wintypes.HANDLE
        user32.FindWindowW.restype = wintypes.HWND
        user32.SendMessageW.restype = wintypes.LPARAM

        hwnd = user32.FindWindowW(None, WINDOW_TITLE)
        if not hwnd:
            log.debug("Window not found yet — skipping icon set")
            return

        hicon_small = user32.LoadImageW(0, str(ico), IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        hicon_big   = user32.LoadImageW(0, str(ico), IMAGE_ICON, 32, 32, LR_LOADFROMFILE)

        if hicon_small:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
        if hicon_big:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_big)
        log.info(f"Window icon applied from {ico}")
    except Exception as e:
        log.warning(f"Could not set window icon: {e}")


def _build_tray_icon():
    """Build a pystray Icon image using the app's icon.ico, or a generated fallback."""
    try:
        from PIL import Image as PILImage
        ico = _icon_path()
        if ico is not None:
            img = PILImage.open(str(ico))
            # pystray needs RGBA
            img = img.convert("RGBA")
        else:
            log.warning("No icon.ico found — using fallback tray square")
            # Fallback: simple navy square
            img = PILImage.new("RGBA", (64, 64), (15, 52, 96, 255))
        return img
    except Exception as e:
        log.warning(f"Could not load tray icon: {e}")
        return None


def main():
    log.info("MC & S CoWorker starting...")

    # Set our AppUserModelID before any window appears, otherwise Windows
    # groups our taskbar entry under pythonw.exe and ignores WM_SETICON.
    _set_app_user_model_id()

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
        title=WINDOW_TITLE,
        url=url,
        width=1280,
        height=800,
        min_size=(1024, 640),
        resizable=True,
        text_select=False,
        confirm_close=False,
    )

    # Inject desktop flag + API token for React app.
    # The token must be JSON-encoded to escape safely (it's a URL-safe base64
    # string, so no real risk, but belt-and-braces).
    def on_loaded():
        try:
            import json as _json
            token_js = _json.dumps(API_TOKEN)
            window.evaluate_js(
                f"window.__pywebview__ = true; window.__API_TOKEN__ = {token_js};"
            )
        except Exception:
            pass
        # Apply the taskbar/title-bar icon now that the native window exists.
        _apply_window_icon()

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
