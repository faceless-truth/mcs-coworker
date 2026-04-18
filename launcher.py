"""
EVA — Windows Launcher
================================
This is the entry point launched by EVA.vbs via wscript.exe.
It is a thin wrapper that:

  1. Shows a small splash/status window ("Checking for updates...")
  2. Runs auto_updater.apply_update() silently
  3. Launches main.py (the full app) in a subprocess using the bundled Python
  4. Exits once the main app is running

The launcher itself is kept tiny so it compiles fast and rarely needs updating.
The main app code (main.py, plugins, etc.) lives in the app folder and updates
via git pull — so a new installer is only needed when Python dependencies change.

INSTALL LAYOUT (created by installer.iss)
-----------------------------------------
C:\\Program Files\\EVA\\
    EVA.vbs              <- VBScript launcher (no compilation needed)
    python\\                 <- embedded Python 3.11 runtime
        python.exe
        pythonw.exe
        Lib\\
        ...
    app\\                    <- the git repo (mcs-coworker)
        main.py
        auto_updater.py
        plugins\\
        requirements.txt
        VERSION
        ...
    data\\                   <- user data (coworker.db, config)
        coworker.db
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
import threading
import tkinter as tk
from pathlib import Path


# ── Path resolution ────────────────────────────────────────────────────────────

def _install_dir() -> Path:
    """Root of the installation directory (where EVA.exe lives)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    # Dev mode: assume we are in the repo root
    return Path(__file__).parent


def _python_exe() -> Path:
    """Path to the bundled Python executable."""
    bundled = _install_dir() / "python" / "python.exe"
    if bundled.exists():
        return bundled
    # Fallback: use the same Python that is running this launcher
    return Path(sys.executable)


def _app_dir() -> Path:
    """Path to the app source code (the git repo)."""
    return _install_dir() / "app"


def _data_dir() -> Path:
    """Path to the user data directory."""
    d = _install_dir() / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


# ── Splash window ──────────────────────────────────────────────────────────────

class SplashWindow:
    """Minimal status window shown during update check."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("EVA")
        self.root.resizable(False, False)
        self.root.overrideredirect(True)   # borderless
        self.root.configure(bg="#1a1a2e")

        # Centre on screen
        w, h = 380, 110
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        tk.Label(
            self.root, text="EVA",
            font=("Segoe UI", 16, "bold"),
            fg="#e0e0e0", bg="#1a1a2e"
        ).pack(pady=(18, 4))

        self._status_var = tk.StringVar(value="Starting…")
        tk.Label(
            self.root, textvariable=self._status_var,
            font=("Segoe UI", 10),
            fg="#9ca3af", bg="#1a1a2e"
        ).pack()

        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.update()

    def set_status(self, msg: str) -> None:
        self._status_var.set(msg)
        self.root.update()

    def close(self) -> None:
        try:
            self.root.destroy()
        except Exception:
            pass


# ── Update step ────────────────────────────────────────────────────────────────

def _run_update(splash: SplashWindow) -> None:
    """Run the auto-updater from the app directory."""
    app_dir = _app_dir()
    if not app_dir.exists():
        splash.set_status("App folder not found — please reinstall.")
        time.sleep(4)
        return

    # Add app dir to sys.path so we can import auto_updater
    if str(app_dir) not in sys.path:
        sys.path.insert(0, str(app_dir))

    try:
        import auto_updater as au
        splash.set_status("Checking for updates…")
        info = au.check_for_update()
        if info and info.get("update_available"):
            splash.set_status(f"Updating to {info['latest']}…")
            result = au.apply_update(force=True)
            au.log_update_result(result)
            if result["success"]:
                splash.set_status(f"Updated to {result['version_after']} ✓")
            else:
                splash.set_status("Update failed — using current version")
            time.sleep(1)
        else:
            splash.set_status("Up to date ✓")
            time.sleep(0.5)
    except Exception as e:
        splash.set_status(f"Update check skipped ({e})")
        time.sleep(1)


# ── Launch main app ────────────────────────────────────────────────────────────

def _launch_app() -> subprocess.Popen:
    """Start main.py using the bundled Python runtime."""
    python  = _python_exe()
    main_py = _app_dir() / "main.py"
    data_d  = _data_dir()

    env = os.environ.copy()
    env["EVA_DATA_DIR"]    = str(data_d)
    env["EVA_INSTALL_DIR"] = str(_install_dir())

    # pythonw.exe suppresses the console window on Windows
    pythonw = python.parent / "pythonw.exe"
    runner  = pythonw if pythonw.exists() else python

    proc = subprocess.Popen(
        [str(runner), str(main_py)],
        cwd=str(_app_dir()),
        env=env,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )
    return proc


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    splash = SplashWindow()

    # Run update check in a thread so the splash window stays responsive
    update_done = threading.Event()

    def _update_thread():
        _run_update(splash)
        update_done.set()

    t = threading.Thread(target=_update_thread, daemon=True)
    t.start()

    # Keep splash alive while update runs
    while not update_done.is_set():
        splash.root.update()
        time.sleep(0.05)

    splash.set_status("Launching…")
    splash.root.update()

    try:
        proc = _launch_app()
    except Exception as e:
        splash.set_status(f"Failed to start: {e}")
        time.sleep(5)
        splash.close()
        sys.exit(1)

    # Wait briefly to confirm the app started, then close splash
    time.sleep(1.5)
    splash.close()

    # Wait for the main app to exit (so the launcher process stays alive
    # in Task Manager while the app is running — useful for auto-start)
    proc.wait()


if __name__ == "__main__":
    main()
