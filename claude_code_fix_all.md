# MCS CoWorker — Full Codebase Fix & Harden

## What This App Is

A Windows desktop app for MC & S Accountants (Keysborough, Melbourne). It automates email triage, AI-drafted replies, and workflow tasks via a plugin architecture. Built with React + Vite frontend, pywebview native window, Flask API server on localhost:7842, and Python backend.

**Repo:** `C:\Users\Elio\mcs-coworker` (GitHub: `faceless-truth/mcs-coworker`)

## Distribution Model — This Is Critical

Accountants are non-technical. They will NEVER run pip, cmd, python, or any terminal command. The entire delivery is:

1. I run `build_installer.bat` on my dev machine → produces `MCSCoWorker_Setup.exe`
2. Accountants download the .exe from SharePoint and double-click to install
3. Future code updates happen automatically via `git pull` inside `auto_updater.py` on each launch
4. Future PACKAGE updates require me to rebuild the installer and re-distribute the .exe

The installer (`build_installer.bat` + `installer.iss`) already handles this correctly — it downloads embeddable Python 3.11, installs pip, installs all packages, bundles the frontend, and produces a single .exe via Inno Setup. The problem is that several files are broken or incomplete, which causes crashes after install.

## How It Runs (Production)

```
Desktop shortcut
  → wscript.exe MCSCoWorker.vbs (no console window)
    → pythonw.exe app\launcher.py (checks for git updates)
      → pythonw.exe app\main.py (starts Flask + pywebview + tray icon)
```

Install layout:
```
%LOCALAPPDATA%\Programs\MCS CoWorker\
  MCSCoWorker.vbs
  python\          ← embeddable Python 3.11 + all packages pre-installed
  app\             ← git clone of this repo (auto-updates via git pull)
    main.py
    api_server.py
    plugins\
    frontend_dist\ ← built React app
  data\            ← SQLite DB, logs (never touched by updates)
  assets\
```

## How It Runs (Dev — my machine only)

```
cd C:\Users\Elio\mcs-coworker
venv\Scripts\activate
set PYWEBVIEW_GUI=edgechromium
python main.py
```

---

## ALL Issues To Fix (work through in this order)

---

### 1. Fix `requirements.txt` — it's dangerously incomplete

This is the root cause of most problems. The current `requirements.txt` is missing ~10 packages that the app actually imports. Meanwhile it includes `chromadb` which drags in 40+ transitive dependencies (onnxruntime, grpc, kubernetes, opentelemetry) that aren't needed.

`build_installer.bat` (lines 124-129) has its own hardcoded package list that differs from `requirements.txt`. The auto-updater runs `pip install -r requirements.txt` on each launch — so if requirements.txt is wrong, every accountant's machine breaks.

**Replace the entire contents of `requirements.txt` with:**

```
# === MCS CoWorker Python Dependencies ===
# This file is the SINGLE SOURCE OF TRUTH for all packages.
# - build_installer.bat reads this file
# - auto_updater.py runs pip install -r this file on each update
# - Accountants never see this — packages are pre-bundled in the installer
#
# DO NOT add chromadb — it pulls 40+ transitive deps and is not used.
# If needed later, make it a separate optional install.

# Core app framework
flask>=3.0.0
flask-cors>=4.0.0
pywebview>=5.0
pythonnet>=3.0.3
cffi>=1.17

# AI & auth
anthropic>=0.34.0
msal>=1.26.0

# HTTP & networking
requests>=2.31.0

# Data processing & parsing
pdfminer.six>=20221105
openpyxl>=3.1.0
pandas>=2.0.0
beautifulsoup4>=4.12.0

# Scheduling & time
schedule>=1.2.0
python-dateutil>=2.8.0
holidays>=0.40
pytz>=2024.1

# Desktop UI & system tray
pystray>=0.19.5
Pillow>=10.0.0
customtkinter>=5.2.0
```

---

### 2. Update `build_installer.bat` to use `requirements.txt` instead of hardcoded packages

Currently `build_installer.bat` lines 124-129 have a hardcoded pip install command. This means requirements.txt and the installer can drift apart. Change it to read from requirements.txt.

**Find this block (around line 124):**
```bat
"%PYTHON_DIR%\python.exe" -m pip install --no-warn-script-location --quiet ^
    anthropic flask flask-cors requests ^
    pywebview pythonnet ^
    pdfminer.six pillow openpyxl pandas beautifulsoup4 ^
    schedule python-dateutil holidays ^
    msal
```

**Replace with:**
```bat
"%PYTHON_DIR%\python.exe" -m pip install --no-warn-script-location --quiet ^
    -r "%REPO_DIR%\requirements.txt"
```

This makes `requirements.txt` the single source of truth. When I add a new package, I update one file and both dev and installer pick it up.

---

### 3. Add `PYWEBVIEW_GUI=edgechromium` to `main.py`

On a clean install, pywebview defaults to the winforms/.NET backend via pythonnet. This crashes with `No module named '_cffi_backend'` because the cffi C extension doesn't initialise cleanly in embeddable Python. The EdgeChromium backend (WebView2) works on every Windows 10/11 machine — Edge is always present.

**In `main.py`, add this line after the logging setup (after line 32, before any other imports):**

```python
# Force EdgeChromium backend — avoids .NET/pythonnet init failures in embeddable Python
os.environ.setdefault("PYWEBVIEW_GUI", "edgechromium")
```

Use `setdefault` so it doesn't override an explicit env var.

---

### 4. Add startup crash catcher to `main.py`

When `main.py` crashes on import, `pythonw.exe` (no console) swallows the error completely. The accountant sees a flash and nothing else. No log, no error dialog, nothing.

**Add this at the VERY TOP of `main.py`, before ALL other code (before `from __future__ import annotations`):**

```python
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
```

Now if any future crash happens — even a missing package — it writes to `data\startup_error.log` and I can diagnose it remotely.

---

### 5. Add frontend fallback warning in `main.py`

In `_get_frontend_url()` (around line 106), if `frontend_dist` doesn't exist, it silently falls back to `http://127.0.0.1:3000/` (Vite dev server). On an accountant's machine nobody is running Vite, so they'd get a blank window with no explanation.

**Replace the `_get_frontend_url` function with:**

```python
def _get_frontend_url() -> str:
    if FRONTEND_DIR.exists():
        return "http://127.0.0.1:7842/"
    log.warning(
        "frontend_dist not found at %s — falling back to Vite dev server (port 3000). "
        "This will show a blank window on production installs.", FRONTEND_DIR
    )
    return "http://127.0.0.1:3000/"
```

---

### 6. Rewrite `update.bat` to include package installation

The current `update.bat` only copies Python source files and builds the frontend. It does NOT install packages. If an accountant somehow uninstalls and re-runs update.bat (or if the embeddable Python's packages get corrupted), the app crashes with no explanation.

**Replace the entire contents of `update.bat` with:**

```bat
@echo off
setlocal
set REPO_DIR=%~dp0
set INSTALL_DIR=%LOCALAPPDATA%\Programs\MCS CoWorker
set APP_DIR=%INSTALL_DIR%\app
set PYTHON=%INSTALL_DIR%\python\python.exe
set FRONTEND_DIR=%REPO_DIR%frontend

echo.
echo  ============================================================
echo   MCS CoWorker - Update
echo  ============================================================
echo.

:: ── Step 0: Verify install exists ─────────────────────────────
if not exist "%PYTHON%" (
    echo [ERROR] Python runtime not found at:
    echo         %PYTHON%
    echo.
    echo         The app is not installed. Run MCSCoWorker_Setup.exe first.
    echo         Download it from SharePoint.
    pause
    exit /b 1
)

echo [1/5] Pulling latest code from GitHub...
cd /d "%REPO_DIR%"
git pull origin main
if errorlevel 1 (
    echo [WARN] git pull failed — continuing with local code...
)
echo.

echo [2/5] Ensuring Python packages are up to date...
:: Enable site-packages in embeddable Python (idempotent — safe to run every time)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$pth = Get-ChildItem '%INSTALL_DIR%\python' -Filter 'python*._pth' | Select-Object -First 1; " ^
    "if ($pth) { (Get-Content $pth.FullName) -replace '#import site','import site' | Set-Content $pth.FullName }"

:: Ensure pip is installed
"%PYTHON%" -m pip --version >nul 2>&1
if errorlevel 1 (
    echo       Installing pip...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%TEMP%\get-pip.py' -UseBasicParsing"
    "%PYTHON%" "%TEMP%\get-pip.py" --no-warn-script-location --quiet
    del "%TEMP%\get-pip.py" >nul 2>&1
)

:: Install/upgrade from requirements.txt
"%PYTHON%" -m pip install --no-warn-script-location --quiet -r "%REPO_DIR%requirements.txt"
if errorlevel 1 (
    echo [WARN] Some packages may not have installed — continuing...
)
echo       Packages OK.
echo.

echo [3/5] Building frontend...
if not exist "%FRONTEND_DIR%\package.json" (
    echo       Frontend source not found — skipping build.
) else (
    cd /d "%FRONTEND_DIR%"
    call pnpm install --frozen-lockfile 2>nul || call pnpm install
    call pnpm build
    echo       Frontend built.
)
echo.

echo [4/5] Copying app files to install...
if not exist "%APP_DIR%" mkdir "%APP_DIR%"
if not exist "%APP_DIR%\plugins" mkdir "%APP_DIR%\plugins"
if not exist "%APP_DIR%\assets" mkdir "%APP_DIR%\assets"
if not exist "%APP_DIR%\frontend_dist" mkdir "%APP_DIR%\frontend_dist"
if not exist "%INSTALL_DIR%\data" mkdir "%INSTALL_DIR%\data"
xcopy /E /I /Y "%REPO_DIR%*.py" "%APP_DIR%\" >nul
xcopy /Y "%REPO_DIR%requirements.txt" "%APP_DIR%\" >nul
xcopy /E /I /Y "%REPO_DIR%plugins\*.py" "%APP_DIR%\plugins\" >nul
xcopy /E /I /Y "%REPO_DIR%assets\*" "%APP_DIR%\assets\" >nul
echo       Done.
echo.

echo [5/5] Deploying frontend...
if exist "%FRONTEND_DIR%\dist\public" (
    xcopy /E /I /Y "%FRONTEND_DIR%\dist\public\*" "%APP_DIR%\frontend_dist\" >nul
    echo       Done.
) else (
    echo       No frontend build output found — skipping.
)
echo.

echo  ============================================================
echo   Update complete!
echo   Launch MCS CoWorker from the Desktop shortcut.
echo  ============================================================
echo.
pause
```

---

### 7. Fix hardcoded usernames in all batch files

Several batch files hardcode `C:\Users\ElioScarton\...` which won't work on any other machine.

**`fix_launch.bat` line 8:**
Change `set INSTALL_APP=C:\Users\ElioScarton\AppData\Local\Programs\MCS CoWorker\app`
To: `set INSTALL_APP=%LOCALAPPDATA%\Programs\MCS CoWorker\app`

**`patch_install.bat` lines 8-11:**
Replace:
```bat
set INSTALL_DIR=C:\Users\ElioScarton\AppData\Local\Programs\MCS CoWorker
set PYTHON=%INSTALL_DIR%\python\python.exe
set APP_DIR=%INSTALL_DIR%\app
set REPO_DIR=C:\Users\ElioScarton\mcs-coworker
```
With:
```bat
set INSTALL_DIR=%LOCALAPPDATA%\Programs\MCS CoWorker
set PYTHON=%INSTALL_DIR%\python\python.exe
set APP_DIR=%INSTALL_DIR%\app
set REPO_DIR=%~dp0
```

**`build_installer.bat` line 13:**
Change `set REPO_DIR=C:\Users\ElioScarton\mcs-coworker`
To: `set REPO_DIR=%~dp0`

**`build_installer.bat` line 14:**
Change `set FRONTEND_DIR=C:\Users\ElioScarton\mcs-coworker-demo`
To: `set FRONTEND_DIR=%~dp0..\mcs-coworker-demo`
(And make sure the `find_frontend_dir` subroutine at the bottom also uses relative paths.)

---

### 8. Verify `auto_updater.py` uses `requirements.txt`

The auto-updater runs on every app launch and does `git pull` + `pip install -r requirements.txt` (line ~14 in the docstring, actual implementation further down). Since we're fixing requirements.txt in step 1, the auto-updater will now install the correct packages automatically.

**Verify:** Check that `auto_updater.py` actually runs `pip install -r requirements.txt` (not a hardcoded list). If it has a hardcoded package list anywhere, replace it with `-r requirements.txt`.

---

### 9. Update `CLAUDE.md`

The `CLAUDE.md` file is the context document for future Claude Code sessions. It's severely outdated — still describes CustomTkinter as the UI when the app has been React + pywebview for months.

**Rewrite `CLAUDE.md` to reflect the current state. Key updates:**

- **Tech Stack:** React + Vite frontend, pywebview (edgechromium backend) native window, Flask API on localhost:7842. NOT CustomTkinter.
- **Entry point:** `main.py` (pywebview + Flask), not `app.py` (old CustomTkinter UI)
- **File structure:** Add all current files: `api_server.py`, `gateway_client.py`, `event_bus.py`, `event_wiring.py`, `token_meter.py`, `memory_store.py`, `xero_oauth.py`, `auto_updater.py`, `main.py`. Note that `app.py` is the legacy CustomTkinter UI — still in the repo but not the current entry point.
- **Launch chain:** VBS → launcher.py → main.py
- **Distribution:** `build_installer.bat` → `MCSCoWorker_Setup.exe` via Inno Setup. Accountants download and double-click. No pip, no terminal, no Python knowledge.
- **Auto-updates:** `auto_updater.py` does `git pull` + `pip install -r requirements.txt` on every launch
- **Requirements:** `requirements.txt` is the single source of truth. `build_installer.bat` reads it. Don't hardcode package lists elsewhere.
- **Plugins:** ~18 plugins in `plugins/` folder now
- **Known issue:** `config.py` DB path (`~/.mcs_email_automation/config.db`) doesn't match installer data path (`INSTALL_DIR\data\`). The `MCS_DATA_DIR` env var is set by launcher.py but config.py ignores it.
- **pywebview backend:** Must use edgechromium. Set via `os.environ.setdefault("PYWEBVIEW_GUI", "edgechromium")` in main.py.
- **Azure credentials:** Tenant ID and Client ID are hardcoded in `graph_client.py` — intentional for single-tenant deployment.

---

## After All Fixes

1. Test locally:
```
venv\Scripts\activate
set PYWEBVIEW_GUI=edgechromium
python main.py
```
Confirm: all plugins load (no pydantic errors), window opens, tray icon works.

2. Commit and push:
```
git add -A
git commit -m "fix: complete requirements.txt, add crash logging, set edgechromium backend, fix hardcoded paths, harden update pipeline"
git push origin main
```

3. Build installer (when ready to distribute):
```
build_installer.bat
```
The output `installer_output\MCSCoWorker_Setup.exe` is all any accountant needs. They double-click it, the app installs, and it auto-updates forever after.
