@echo off
setlocal
set REPO_DIR=%~dp0
set INSTALL_DIR=%LOCALAPPDATA%\Programs\MCS CoWorker
set APP_DIR=%INSTALL_DIR%\app
set FRONTEND_DIR=%REPO_DIR%frontend

echo.
echo  ============================================================
echo   MCS CoWorker - Update
echo  ============================================================
echo.

echo [1/4] Pulling latest code from GitHub...
cd /d "%REPO_DIR%"
git pull origin main
if errorlevel 1 ( echo [ERROR] git pull failed. & pause & exit /b 1 )
echo.

echo [2/4] Building frontend...
cd /d "%FRONTEND_DIR%"
if not exist "node_modules" (
    echo       Installing npm packages for the first time...
    where pnpm >nul 2>&1
    if errorlevel 1 ( npm install -g pnpm )
    pnpm install --frozen-lockfile
    if errorlevel 1 ( echo [ERROR] pnpm install failed. & pause & exit /b 1 )
)
echo       Running pnpm build...
pnpm build
if errorlevel 1 ( echo [ERROR] Frontend build failed. & pause & exit /b 1 )
echo       Frontend built OK.
echo.

echo [3/4] Copying app files...
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

echo [4/4] Deploying frontend...
xcopy /E /I /Y "%FRONTEND_DIR%\dist\public\*" "%APP_DIR%\frontend_dist\" >nul
echo       Done.
echo.

echo  ============================================================
echo   Update complete!
echo   Launch MCS CoWorker from the Desktop shortcut.
echo  ============================================================
echo.
pause
