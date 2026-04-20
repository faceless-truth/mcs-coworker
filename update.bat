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

:: -- Step 0: Verify install exists -----------------------------
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
    echo [WARN] git pull failed - continuing with local code...
)
echo.

echo [2/5] Ensuring Python packages are up to date...
:: Enable site-packages in embeddable Python (idempotent - safe to run every time)
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
    echo [WARN] Some packages may not have installed - continuing...
)
echo       Packages OK.
echo.

echo [3/5] Building frontend...
if not exist "%FRONTEND_DIR%\package.json" (
    echo       Frontend source not found - skipping build.
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
    echo       No frontend build output found - skipping.
)
echo.

echo  ============================================================
echo   Update complete!
echo   Launch MCS CoWorker from the Desktop shortcut.
echo  ============================================================
echo.
pause
