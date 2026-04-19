@echo off
:: ============================================================
::  MCS CoWorker — Update Script
::  Run this to pull the latest code and rebuild the frontend.
::  Place this file in: C:\Users\ElioScarton\mcs-coworker\
:: ============================================================
setlocal

set "REPO_DIR=%~dp0"
set "INSTALL_DIR=%LOCALAPPDATA%\Programs\MCS CoWorker"
set "APP_DIR=%INSTALL_DIR%\app"
set "PYTHON=%INSTALL_DIR%\python\python.exe"
set "FRONTEND_DIR=%REPO_DIR%frontend"

echo.
echo  ============================================================
echo   MCS CoWorker — Update
echo  ============================================================
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

:: ── Step 2: Build frontend ────────────────────────────────────
echo [2/4] Building frontend...
cd /d "%FRONTEND_DIR%"

:: Install node_modules if missing
if not exist "node_modules" (
    echo       Installing npm packages (first time only)...
    where pnpm >nul 2>&1
    if errorlevel 1 (
        npm install -g pnpm
    )
    pnpm install --frozen-lockfile
)

pnpm build
if errorlevel 1 (
    echo [ERROR] Frontend build failed.
    pause & exit /b 1
)
echo       Frontend built.
echo.

:: ── Step 3: Copy updated Python files ───────────────────────
echo [3/4] Copying updated app files...
if not exist "%APP_DIR%" mkdir "%APP_DIR%"
if not exist "%APP_DIR%\plugins" mkdir "%APP_DIR%\plugins"

copy /Y "%REPO_DIR%main.py"          "%APP_DIR%\main.py"          >nul
copy /Y "%REPO_DIR%api_server.py"    "%APP_DIR%\api_server.py"    >nul
copy /Y "%REPO_DIR%launcher.py"      "%APP_DIR%\launcher.py"      >nul
copy /Y "%REPO_DIR%plugin_loader.py" "%APP_DIR%\plugin_loader.py" >nul
copy /Y "%REPO_DIR%plugin_base.py"   "%APP_DIR%\plugin_base.py"   >nul
copy /Y "%REPO_DIR%graph_client.py"  "%APP_DIR%\graph_client.py"  >nul
copy /Y "%REPO_DIR%config.py"        "%APP_DIR%\config.py"        >nul
copy /Y "%REPO_DIR%memory_store.py"  "%APP_DIR%\memory_store.py"  >nul
copy /Y "%REPO_DIR%token_meter.py"   "%APP_DIR%\token_meter.py"   >nul
copy /Y "%REPO_DIR%kpi_monitor.py"   "%APP_DIR%\kpi_monitor.py"   >nul
copy /Y "%REPO_DIR%event_bus.py"     "%APP_DIR%\event_bus.py"     >nul
copy /Y "%REPO_DIR%event_wiring.py"  "%APP_DIR%\event_wiring.py"  >nul
copy /Y "%REPO_DIR%approval_queue.py" "%APP_DIR%\approval_queue.py" >nul
copy /Y "%REPO_DIR%xero_oauth.py"    "%APP_DIR%\xero_oauth.py"    >nul
copy /Y "%REPO_DIR%gateway_client.py" "%APP_DIR%\gateway_client.py" >nul

:: Copy all plugins
xcopy /Y /Q "%REPO_DIR%plugins\*.py" "%APP_DIR%\plugins\" >nul
echo       Done.
echo.

:: ── Step 4: Copy built frontend ───────────────────────────────
echo [4/4] Deploying frontend...
if not exist "%APP_DIR%\frontend_dist" mkdir "%APP_DIR%\frontend_dist"
xcopy /E /I /Y /Q "%FRONTEND_DIR%\dist\public\*" "%APP_DIR%\frontend_dist\" >nul
echo       Done.
echo.

echo  ============================================================
echo   Update complete!
echo   Restart MCS CoWorker from the system tray to apply changes.
echo  ============================================================
echo.
pause
