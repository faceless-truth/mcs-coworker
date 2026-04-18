"""
MCS CoWorker — Windows Launcher
================================
This is the entry point launched by MCSCoWorker.vbs via wscript.exe.
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
C:\\Program Files\\MCS CoWorker\\
    MCSCoWorker.vbs              <- VBScript launcher (no compilation needed)
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
import traceback
from pathlib import Path

# ── Logging to file (since there is no console window) ────────────────────────

def _log_dir() -> Path:
    """Return a writable log directory."""
    # Try next to the script first (install dir), fall back to TEMP
    try:
        d = Path(__file__).parent.parent / "data"
        d.mkdir(parents=True, exist_ok=True)
        return d
    except Exception:
        return Path(os.environ.get("TEMP", "."))

def _write_log(msg: str) -> None:
    """Append a timestamped line to launcher.log."""
    try:
        import datetime
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_path = _log_dir() / "launcher.log"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass

# ── Path resolution ────────────────────────────────────────────────────────────

def _install_dir() -> Path:
    """Root of the installation directory (parent of the 'app' folder)."""
    # launcher.py lives in: {install_dir}\app\launcher.py
    return Path(__file__).parent.parent


def _python_exe() -> Path:
    """Path to the bundled Python executable."""
    bundled = _install_dir() / "python" / "python.exe"
    if bundled.exists():
        return bundled
    # Fallback: use the same Python that is running this launcher
    return Path(sys.executable)


def _app_dir() -> Path:
    """Path to the app source code."""
    return Path(__file__).parent


def _data_dir() -> Path:
    """Path to the user data directory."""
    d = _install_dir() / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


# ── Splash window ──────────────────────────────────────────────────────────────
# tkinter is NOT available in the embeddable Python distribution.
# We use a ctypes MessageBox as a last-resort error dialog, and skip the
# splash entirely — the app opens fast enough that a splash is not needed.

def _show_error(title: str, msg: str) -> None:
    """Show a Windows error MessageBox without requiring tkinter."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, title, 0x10)  # MB_ICONERROR
    except Exception:
        pass


# ── Update step ────────────────────────────────────────────────────────────────

def _run_update() -> None:
    """Run the auto-updater from the app directory (best-effort, never fatal)."""
    app_dir = _app_dir()
    if not app_dir.exists():
        _write_log(f"App folder not found: {app_dir}")
        return

    if str(app_dir) not in sys.path:
        sys.path.insert(0, str(app_dir))

    try:
        import auto_updater as au
        _write_log("Checking for updates...")
        info = au.check_for_update()
        if info and info.get("update_available"):
            _write_log(f"Update available: {info.get('latest')} — applying...")
            result = au.apply_update(force=True)
            au.log_update_result(result)
            if result["success"]:
                _write_log(f"Updated to {result['version_after']}")
            else:
                _write_log("Update failed — using current version")
        else:
            _write_log("Already up to date.")
    except Exception as e:
        _write_log(f"Update check skipped: {e}")


# ── Launch main app ────────────────────────────────────────────────────────────

def _launch_app() -> subprocess.Popen:
    """Start main.py using the bundled Python runtime."""
    python  = _python_exe()
    main_py = _app_dir() / "main.py"
    data_d  = _data_dir()

    _write_log(f"python  = {python}")
    _write_log(f"main.py = {main_py}")
    _write_log(f"cwd     = {_app_dir()}")

    if not python.exists():
        raise FileNotFoundError(f"Python not found: {python}")
    if not main_py.exists():
        raise FileNotFoundError(f"main.py not found: {main_py}")

    env = os.environ.copy()
    env["MCS_DATA_DIR"]    = str(data_d)
    env["PYTHONIOENCODING"] = "utf-8"
    env["MCS_INSTALL_DIR"] = str(_install_dir())

    # pythonw.exe suppresses the console window on Windows
    pythonw = python.parent / "pythonw.exe"
    runner  = pythonw if pythonw.exists() else python

    _write_log(f"runner  = {runner}")

    proc = subprocess.Popen(
        [str(runner), str(main_py)],
        cwd=str(_app_dir()),
        env=env,
        # Redirect stdout/stderr to a log file so errors are visible
        stdout=open(_log_dir() / "app_stdout.log", "w", encoding="utf-8", errors="replace"),
        stderr=open(_log_dir() / "app_stderr.log", "w", encoding="utf-8", errors="replace"),
    )
    return proc


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    _write_log("=== MCS CoWorker launcher starting ===")
    _write_log(f"install_dir = {_install_dir()}")
    _write_log(f"app_dir     = {_app_dir()}")

    # Run update check (best-effort, non-blocking)
    try:
        _run_update()
    except Exception as e:
        _write_log(f"Update step error: {e}")

    # Launch the main app
    try:
        proc = _launch_app()
        _write_log(f"App launched, PID={proc.pid}")
    except Exception as e:
        _write_log(f"FATAL: Failed to launch app: {e}\n{traceback.format_exc()}")
        _show_error(
            "MCS CoWorker — Launch Failed",
            f"The app could not start.\n\n{e}\n\n"
            f"Check the log file at:\n{_log_dir() / 'launcher.log'}"
        )
        sys.exit(1)

    # Wait for the main app to exit (keeps the launcher alive in Task Manager
    # while the app is running — useful for the auto-start registry entry)
    proc.wait()
    _write_log("App exited. Launcher done.")


if __name__ == "__main__":
    main()
