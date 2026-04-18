@echo off
echo ===================================
echo  MCS CoWorker - Build Frontend
echo ===================================
echo.

:: Check Node is available
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Node.js not found. Install from https://nodejs.org/
    pause
    exit /b 1
)

:: Navigate to the frontend source (mcs-coworker-demo)
:: Assumes the frontend repo is cloned alongside mcs-coworker
set FRONTEND_DIR=%~dp0..\mcs-coworker-demo

if not exist "%FRONTEND_DIR%" (
    echo ERROR: Frontend directory not found at %FRONTEND_DIR%
    echo Please clone mcs-coworker-demo alongside mcs-coworker
    pause
    exit /b 1
)

echo Building React frontend...
cd /d "%FRONTEND_DIR%"
call npm install
call npm run build

:: Copy the build output to the electron/frontend/dist directory
set DEST=%~dp0frontend\dist
if not exist "%~dp0frontend" mkdir "%~dp0frontend"

echo Copying build to %DEST%...
xcopy /E /I /Y "%FRONTEND_DIR%\dist" "%DEST%"

echo.
echo ✓ Frontend built successfully
echo   Output: %DEST%
echo.
pause
