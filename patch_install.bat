@echo off
:: ============================================================
:: MCS CoWorker — One-Shot Install Patcher
:: Run this once to fix the existing installation.
:: ============================================================
setlocal

set INSTALL_DIR=C:\Users\ElioScarton\AppData\Local\Programs\MCS CoWorker
set PYTHON=%INSTALL_DIR%\python\python.exe
set APP_DIR=%INSTALL_DIR%\app
set REPO_DIR=C:\Users\ElioScarton\mcs-coworker

echo ============================================================
echo  MCS CoWorker — Install Patcher
echo ============================================================
echo.

:: ── Step 1: Pull latest code ─────────────────────────────────
echo [1/4] Pulling latest code from GitHub...
cd /d "%REPO_DIR%"
git pull origin main
if errorlevel 1 (
    echo [ERROR] git pull failed. Check your internet connection.
    pause & exit /b 1
)
echo.

:: ── Step 2: Copy updated Python files ───────────────────────
echo [2/4] Copying updated files to install...
copy /Y "%REPO_DIR%\main.py"          "%APP_DIR%\main.py"
copy /Y "%REPO_DIR%\api_server.py"    "%APP_DIR%\api_server.py"
copy /Y "%REPO_DIR%\launcher.py"      "%APP_DIR%\launcher.py"
copy /Y "%REPO_DIR%\plugin_loader.py" "%APP_DIR%\plugin_loader.py"
echo.

:: ── Step 3: Install pythonnet (required for pywebview on Windows) ──
echo [3/4] Installing pythonnet for pywebview WebView2 backend...
echo       (This may take 1-2 minutes on first run)
"%PYTHON%" -m pip install --no-warn-script-location --quiet pythonnet
if errorlevel 1 (
    echo [WARN] pythonnet install failed — trying clr-loader separately...
    "%PYTHON%" -m pip install --no-warn-script-location --quiet clr-loader
)
echo.

:: ── Step 4: Copy frontend dist ───────────────────────────────
echo [4/4] Checking frontend...
if exist "%APP_DIR%\frontend_dist\index.html" (
    echo       Frontend already present — skipping.
) else (
    echo       Frontend not found in install.
    echo       Please download frontend_dist.zip from Manus and run:
    echo         powershell -Command "Expand-Archive -Path '%%USERPROFILE%%\Downloads\frontend_dist.zip' -DestinationPath '%APP_DIR%\frontend_dist' -Force"
)
echo.

echo ============================================================
echo  Done! Launch MCS CoWorker from the Start Menu or desktop.
echo ============================================================
echo.
pause
