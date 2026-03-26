@echo off
title MC ^& S Desktop Agent — Build
echo ============================================
echo   MC ^& S Desktop Agent — Build Script
echo ============================================
echo.

:: Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found on PATH.
    echo         Install Python 3.10+ from python.org and tick "Add Python to PATH".
    pause
    exit /b 1
)

:: Create venv if it doesn't exist
if not exist "venv\" (
    echo [SETUP] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

:: Activate venv
echo [SETUP] Activating environment...
call venv\Scripts\activate.bat

:: Install dependencies
echo [SETUP] Installing dependencies...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install requirements.
    pause
    exit /b 1
)

:: Install PyInstaller if not already present
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [SETUP] Installing PyInstaller...
    pip install pyinstaller --quiet
    if errorlevel 1 (
        echo [ERROR] Failed to install PyInstaller.
        pause
        exit /b 1
    )
)

:: Run the build
echo.
echo [BUILD] Running PyInstaller...
echo.
pyinstaller build.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. See output above for details.
    pause
    exit /b 1
)

:: Copy plugins folder next to the exe (must be a real writable directory)
echo.
echo [BUILD] Copying plugins folder...
xcopy /E /I /Y plugins "dist\MCS Desktop Agent\plugins"
if errorlevel 1 (
    echo [WARN] Failed to copy plugins folder — check it exists.
)

echo.
echo ============================================
echo   Build complete. Find your app in the dist/ folder.
echo ============================================
echo.
pause
