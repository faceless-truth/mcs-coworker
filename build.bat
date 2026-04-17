@echo off
title MC ^& S CoWorker — Build
echo ============================================
echo   MC ^& S CoWorker — Build Script (pywebview)
echo ============================================
echo.

:: Check prerequisites
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found on PATH.
    pause
    exit /b 1
)
node --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js not found. Install from nodejs.org.
    pause
    exit /b 1
)

:: Step 1: Build React frontend
echo [1/4] Building React frontend...
set FRONTEND_DIR=..\mcs-coworker-demo
if not exist "%FRONTEND_DIR%\package.json" (
    echo [ERROR] Frontend not found at %FRONTEND_DIR%
    pause
    exit /b 1
)
pushd %FRONTEND_DIR%
call pnpm install --frozen-lockfile
call pnpm build
if errorlevel 1 (
    popd
    echo [ERROR] Frontend build failed.
    pause
    exit /b 1
)
popd

if exist "frontend_dist\" rmdir /s /q "frontend_dist"
xcopy /E /I /Y "%FRONTEND_DIR%\dist" "frontend_dist"
echo [1/4] Frontend ready.
echo.

:: Step 2: Python venv
echo [2/4] Setting up Python environment...
if not exist "venv\" python -m venv venv
call venv\Scripts\activate.bat

:: Step 3: Install dependencies
echo [3/4] Installing Python dependencies...
pip install -r requirements.txt --quiet
pip install pyinstaller pywebview flask flask-cors --quiet
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)
echo [3/4] Dependencies installed.
echo.

:: Step 4: PyInstaller
echo [4/4] Running PyInstaller...
pyinstaller build.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] Build failed. See output above.
    pause
    exit /b 1
)

:: Post-build: copy plugins and launch scripts
echo.
echo [POST] Copying plugins and launch scripts...
xcopy /E /I /Y plugins "dist\MCS CoWorker\plugins"
copy /Y launch.bat "dist\MCS CoWorker\launch.bat" >nul 2>&1
copy /Y launch_silent.vbs "dist\MCS CoWorker\launch_silent.vbs" >nul 2>&1

echo.
echo ============================================
echo   Build complete!
echo   App:  dist\MCS CoWorker\MCS CoWorker.exe
echo ============================================
echo.
pause
