@echo off
setlocal EnableDelayedExpansion

echo.
echo  ============================================================
echo   MCS CoWorker - Windows Installer Builder
echo  ============================================================
echo.

:: ---------------------------------------------------------------------------
:: CONFIGURATION
:: ---------------------------------------------------------------------------
set REPO_DIR=%~dp0
:: Frontend lives inside the repo. The old sibling-folder path
:: (..\mcs-coworker-demo) came from an early scaffold and silently caused the
:: installer to skip the frontend build on clean checkouts.
set FRONTEND_DIR=%REPO_DIR%frontend
set BUILD_DIR=%REPO_DIR%\installer_build
set OUTPUT_DIR=%REPO_DIR%\installer_output
set PYTHON_EMBED_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip
set PYTHON_EMBED_ZIP=%BUILD_DIR%\python-embed.zip
set PYTHON_DIR=%BUILD_DIR%\python
set APP_DIR=%BUILD_DIR%\app
set INNO_SETUP_DIR=C:\Program Files (x86)\Inno Setup 6
set INNO_INSTALLER_URL=https://github.com/jrsoftware/issrc/releases/download/is-6_7_1/innosetup-6.7.1.exe
set INNO_INSTALLER_TMP=%TEMP%\innosetup-installer.exe

:: ---------------------------------------------------------------------------
:: STEP 0 - Auto-install missing build tools
:: ---------------------------------------------------------------------------
echo [0/6] Checking build tools...

:: Check Python (needed to run this build script itself)
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.11+ from https://python.org/downloads
    echo         Tick "Add Python to PATH" during install, then re-run this script.
    pause & exit /b 1
)

:: Check Git
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found. Install from https://git-scm.com then re-run.
    pause & exit /b 1
)

:: Auto-install Inno Setup if missing
if not exist "%INNO_SETUP_DIR%\iscc.exe" (
    echo [0/6] Inno Setup not found - downloading and installing automatically...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%INNO_INSTALLER_URL%' -OutFile '%INNO_INSTALLER_TMP%' -UseBasicParsing"
    if errorlevel 1 (
        echo [ERROR] Could not download Inno Setup. Check internet connection.
        pause & exit /b 1
    )
    "%INNO_INSTALLER_TMP%" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
    del "%INNO_INSTALLER_TMP%" >nul 2>&1
    ver >nul
    if not exist "%INNO_SETUP_DIR%\iscc.exe" (
        echo [ERROR] Inno Setup install failed.
        pause & exit /b 1
    )
    echo [0/6] Inno Setup installed.
) else (
    echo [0/6] Inno Setup already installed.
)
ver >nul

:: Add Inno Setup to PATH for this session
set "PATH=%PATH%;%INNO_SETUP_DIR%"

:: Auto-install pnpm if missing
where pnpm >nul 2>&1
if errorlevel 1 (
    echo [0/6] pnpm not found - installing via npm...
    npm install -g pnpm --silent >nul 2>&1
    ver >nul
    echo [0/6] pnpm installed.
)

echo [0/6] All build tools ready.
echo.

:: ---------------------------------------------------------------------------
:: STEP 1 - Clean previous build
:: ---------------------------------------------------------------------------
echo [1/6] Cleaning previous build...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%OUTPUT_DIR%" rmdir /s /q "%OUTPUT_DIR%"
mkdir "%BUILD_DIR%"
mkdir "%OUTPUT_DIR%"
mkdir "%PYTHON_DIR%"
mkdir "%APP_DIR%"
echo [1/6] Done.
echo.

:: ---------------------------------------------------------------------------
:: STEP 2 - Download and set up embeddable Python runtime
:: ---------------------------------------------------------------------------
echo [2/6] Setting up Python 3.11 runtime...

powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%PYTHON_EMBED_URL%' -OutFile '%PYTHON_EMBED_ZIP%' -UseBasicParsing"
if errorlevel 1 (
    echo [ERROR] Failed to download Python runtime. Check internet connection.
    pause & exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Path '%PYTHON_EMBED_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force"
del "%PYTHON_EMBED_ZIP%" >nul 2>&1

:: Enable site-packages (uncomment "import site" in the ._pth file)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$pth = Get-ChildItem '%PYTHON_DIR%' -Filter 'python*._pth' | Select-Object -First 1; " ^
    "if ($pth) { (Get-Content $pth.FullName) -replace '#import site','import site' | Set-Content $pth.FullName }"

:: Install pip
echo [2/6] Installing pip...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%BUILD_DIR%\get-pip.py' -UseBasicParsing"
"%PYTHON_DIR%\python.exe" "%BUILD_DIR%\get-pip.py" --no-warn-script-location --quiet
del "%BUILD_DIR%\get-pip.py" >nul 2>&1

:: Install all required packages
:: NOTE: pywebview is forced to the EdgeChromium (WebView2) backend in main.py —
::       pythonnet/cffi are NOT installed because the .NET backend crashes on the
::       embeddable Python bundled here.
echo [2/6] Installing Python packages (this takes a few minutes)...
"%PYTHON_DIR%\python.exe" -m pip install --no-warn-script-location --quiet ^
    -r "%REPO_DIR%\requirements.txt"
if errorlevel 1 (
    echo [ERROR] pip install failed. Check internet connection.
    pause & exit /b 1
)

:: pywebview on Windows needs the WebView2 runtime (Edge-based). Download the
:: Evergreen bootstrapper and bundle it INTO the installer — installer.iss
:: runs it conditionally on the target machine via IsWebView2Installed().
:: Do NOT run it on the build machine.
echo [2/6] Downloading Microsoft Edge WebView2 bootstrapper for installer...
set WV2_URL=https://go.microsoft.com/fwlink/p/?LinkId=2124703
set WV2_EXTRAS_DIR=%BUILD_DIR%\installer_extras
if not exist "%WV2_EXTRAS_DIR%" mkdir "%WV2_EXTRAS_DIR%"
set WV2_OUT=%WV2_EXTRAS_DIR%\MicrosoftEdgeWebview2Setup.exe
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%WV2_URL%' -OutFile '%WV2_OUT%' -UseBasicParsing" >nul 2>&1
if exist "%WV2_OUT%" (
    echo [2/6] WebView2 bootstrapper bundled.
) else (
    echo [ERROR] Could not download WebView2 bootstrapper. Check internet connection.
    pause ^& exit /b 1
)

echo [2/6] Python runtime ready.
echo.

:: ---------------------------------------------------------------------------
:: STEP 3 - Build React frontend
:: ---------------------------------------------------------------------------
echo [3/6] Building React frontend...
if not exist "%FRONTEND_DIR%\package.json" (
    echo [ERROR] Frontend not found at %FRONTEND_DIR%. Check FRONTEND_DIR.
    pause ^& exit /b 1
)
pushd "%FRONTEND_DIR%"
call pnpm install --frozen-lockfile --silent 2>nul
if errorlevel 1 call pnpm install --silent
call pnpm build --silent
if errorlevel 1 (
    popd
    echo [ERROR] Frontend build failed.
    pause ^& exit /b 1
)
popd

:: vite.config.ts puts the build output at frontend\dist\public\.
:: Fail loudly if it's missing — otherwise the installer ships with no frontend.
if not exist "%FRONTEND_DIR%\dist\public\index.html" (
    echo [ERROR] Frontend build output missing at %FRONTEND_DIR%\dist\public\index.html.
    echo         Run "cd frontend ^&^& pnpm build" manually to diagnose.
    pause ^& exit /b 1
)
xcopy /E /I /Y /Q "%FRONTEND_DIR%\dist\public" "%APP_DIR%\frontend_dist" >nul
echo [3/6] Frontend built and copied.
echo.

:: ---------------------------------------------------------------------------
:: STEP 4 - Clone app source into build dir (preserves .git for auto-updater)
:: ---------------------------------------------------------------------------
echo [4/6] Cloning app source (with .git for auto-updates)...

:: Remove the app dir created in STEP 1 so git clone can create it fresh
if exist "%APP_DIR%" rmdir /s /q "%APP_DIR%"

:: Clone the local repo — this is fast (local disk) and includes .git
git clone "%REPO_DIR%" "%APP_DIR%" --quiet
if errorlevel 1 (
    echo [ERROR] git clone failed.
    pause ^& exit /b 1
)

:: Copy the built frontend dist into the cloned app dir
:: Vite writes to dist\public per vite.config.ts outDir.
if exist "%FRONTEND_DIR%\dist\public\index.html" (
    xcopy /E /I /Y /Q "%FRONTEND_DIR%\dist\public" "%APP_DIR%\frontend_dist" >nul
)

:: Remove large/unnecessary folders from the cloned copy to keep installer small
if exist "%APP_DIR%\installer_build"  rmdir /s /q "%APP_DIR%\installer_build"
if exist "%APP_DIR%\installer_output" rmdir /s /q "%APP_DIR%\installer_output"
if exist "%APP_DIR%\node_modules"      rmdir /s /q "%APP_DIR%\node_modules"
if exist "%APP_DIR%\venv"              rmdir /s /q "%APP_DIR%\venv"
if exist "%APP_DIR%\dist"              rmdir /s /q "%APP_DIR%\dist"
if exist "%APP_DIR%\build"             rmdir /s /q "%APP_DIR%\build"

:: Write a VERSION file with the current git commit hash
for /f %%i in ('git -C "%REPO_DIR%" rev-parse --short HEAD 2^>nul') do set GIT_HASH=%%i
if defined GIT_HASH (
    echo %GIT_HASH%> "%APP_DIR%\VERSION"
    echo [4/6] App source copied. Version: %GIT_HASH%
) else (
    echo [4/6] App source copied.
)
echo.

:: ---------------------------------------------------------------------------
:: STEP 5 - Create VBScript launcher (no PyInstaller needed)
:: ---------------------------------------------------------------------------
echo [5/6] Creating launcher...

:: MCSCoWorker.vbs
:: - Uses pythonw.exe so no console window appears
:: - Passes the full path to launcher.py explicitly
:: - Sets working directory to the app folder
:: - Waits=False so the VBS exits immediately (app runs independently)
(
echo Set oShell = CreateObject^("WScript.Shell"^)
echo Set oFSO   = CreateObject^("Scripting.FileSystemObject"^)
echo.
echo scriptDir      = oFSO.GetParentFolderName^(WScript.ScriptFullName^)
echo appDir         = scriptDir ^& "\app"
echo pythonwExe     = scriptDir ^& "\python\pythonw.exe"
echo launcherScript = appDir    ^& "\launcher.py"
echo.
echo ' Sanity checks
echo If Not oFSO.FileExists^(pythonwExe^) Then
echo     MsgBox "MCS CoWorker: pythonw.exe not found at:" ^& vbCrLf ^& pythonwExe ^& vbCrLf ^& vbCrLf ^& _
echo            "Please reinstall from SharePoint.", vbCritical, "MCS CoWorker"
echo     WScript.Quit 1
echo End If
echo If Not oFSO.FileExists^(launcherScript^) Then
echo     MsgBox "MCS CoWorker: launcher.py not found at:" ^& vbCrLf ^& launcherScript ^& vbCrLf ^& vbCrLf ^& _
echo            "Please reinstall from SharePoint.", vbCritical, "MCS CoWorker"
echo     WScript.Quit 1
echo End If
echo.
echo ' Set working directory to app folder
echo oShell.CurrentDirectory = appDir
echo.
echo ' Run: pythonw.exe "C:\...\app\launcher.py"
echo ' WindowStyle 0 = hidden window, bWaitOnReturn False = fire and forget
echo cmd = Chr^(34^) ^& pythonwExe ^& Chr^(34^) ^& " " ^& Chr^(34^) ^& launcherScript ^& Chr^(34^)
echo oShell.Run cmd, 0, False
) > "%BUILD_DIR%\MCSCoWorker.vbs"

echo [5/6] Launcher created (MCSCoWorker.vbs).
echo.

:: ---------------------------------------------------------------------------
:: STEP 6 - Build installer with Inno Setup
:: ---------------------------------------------------------------------------
echo [6/6] Building installer with Inno Setup...
iscc "%REPO_DIR%\installer.iss" /DMyBuildDir="%BUILD_DIR%" /DMyRepoDir="%REPO_DIR%"
if errorlevel 1 (
    echo [ERROR] Inno Setup failed. See output above.
    pause & exit /b 1
)
echo [6/6] Installer built.
echo.

echo  ============================================================
echo   Build complete!
echo.
echo   Installer: %OUTPUT_DIR%\MCSCoWorker_Setup.exe
echo.
echo   Upload MCSCoWorker_Setup.exe to SharePoint.
echo   Accountants download and double-click to install.
echo   Future updates: just git push - no reinstall needed.
echo  ============================================================
echo.
pause
goto :eof
