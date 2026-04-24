@echo off
title MC ^& S Coworker
echo ============================================
echo   MC ^& S Coworker — Launcher
echo ============================================
echo.

:: Use bundled Python (shipped with the installer at ..\python\)
set "PYTHON_DIR=%~dp0..\python"
set "PYTHON_EXE=%PYTHON_DIR%\python.exe"
set "PYTHONW_EXE=%PYTHON_DIR%\pythonw.exe"
set "PIP_EXE=%PYTHON_DIR%\Scripts\pip.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Bundled Python not found at %PYTHON_EXE%
    echo         The installation appears to be corrupted. Please reinstall MC ^& S Coworker.
    pause
    exit /b
)

echo [SETUP] Installing/updating dependencies...
"%PIP_EXE%" install -r requirements.txt --quiet

echo.
echo [START] Launching MC ^& S Coworker...
echo         You can close this window — the app will keep running.
echo.

:: Force EdgeChromium webview backend (winforms/.NET backend crashes embeddable Python)
set PYWEBVIEW_GUI=edgechromium

:: Launch with pythonw (no terminal) if available, otherwise fall back to python
if exist "%PYTHONW_EXE%" (
    start "" "%PYTHONW_EXE%" main.py
) else (
    start /min "" "%PYTHON_EXE%" main.py
)

:: Give it a moment to start, then close the terminal
timeout /t 2 /nobreak >nul
exit
