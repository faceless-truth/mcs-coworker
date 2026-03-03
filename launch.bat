@echo off
title MC ^& S Coworker
echo ============================================
echo   MC ^& S Coworker — Launcher
echo ============================================
echo.

:: Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ from python.org
    echo         Make sure to tick "Add Python to PATH" during installation.
    pause
    exit /b
)

:: Create venv if it doesn't exist
if not exist "venv\" (
    echo [SETUP] First-time setup — creating virtual environment...
    python -m venv venv
)

:: Activate and install
echo [SETUP] Activating environment...
call venv\Scripts\activate.bat

echo [SETUP] Installing/updating dependencies...
pip install -r requirements.txt --quiet

:: Create desktop shortcut on first run
if not exist ".shortcut_created" (
    echo [SETUP] Creating desktop shortcut...
    pip install pywin32 winshell --quiet
    python create_shortcut.py
    echo done > .shortcut_created
    echo.
    echo ============================================
    echo   Desktop shortcut created!
    echo   You can now launch from the icon on
    echo   your desktop: "MC ^& S Coworker"
    echo ============================================
    echo.
)

echo.
echo [START] Launching MC ^& S Coworker...
echo.
python app.py

pause
