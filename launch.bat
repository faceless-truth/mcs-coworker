@echo off
title MC ^& S Desktop Agent
echo ============================================
echo   MC ^& S Desktop Agent — Launcher
echo ============================================
echo.

:: Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ from python.org
    pause
    exit /b
)

:: Create venv if it doesn't exist
if not exist "venv\" (
    echo [SETUP] Creating virtual environment...
    python -m venv venv
)

:: Activate and install
echo [SETUP] Activating environment...
call venv\Scripts\activate.bat

echo [SETUP] Installing/updating dependencies...
pip install -r requirements.txt --quiet

echo.
echo [START] Launching MC ^& S Desktop Agent...
echo.
python app.py

pause
