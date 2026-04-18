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
set REPO_DIR=C:\Users\ElioScarton\mcs-coworker
set FRONTEND_DIR=C:\Users\ElioScarton\mcs-coworker-demo
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
echo [2/6] Installing Python packages (this takes a few minutes)...
"%PYTHON_DIR%\python.exe" -m pip install --no-warn-script-location --quiet ^
    anthropic flask flask-cors requests pywebview ^
    pdfminer.six pillow openpyxl pandas beautifulsoup4 ^
    schedule python-dateutil holidays
if errorlevel 1 (
    echo [ERROR] pip install failed. Check internet connection.
    pause & exit /b 1
)

echo [2/6] Python runtime ready.
echo.

:: ---------------------------------------------------------------------------
:: STEP 3 - Build React frontend
:: ---------------------------------------------------------------------------
echo [3/6] Building React frontend...
if not exist "%FRONTEND_DIR%\package.json" (
    echo [WARN] Frontend not found at %FRONTEND_DIR% - skipping.
) else (
    pushd "%FRONTEND_DIR%"
    call pnpm install --frozen-lockfile --silent 2>nul
    if errorlevel 1 call pnpm install --silent
    call pnpm build --silent
    if errorlevel 1 (
        popd
        echo [ERROR] Frontend build failed.
        pause & exit /b 1
    )
    popd
    xcopy /E /I /Y /Q "%FRONTEND_DIR%\dist" "%APP_DIR%\frontend_dist" >nul
    echo [3/6] Frontend built and copied.
)
echo.

:: ---------------------------------------------------------------------------
:: STEP 4 - Copy app source files
:: ---------------------------------------------------------------------------
echo [4/6] Copying app source...

robocopy "%REPO_DIR%" "%APP_DIR%" *.py /S ^
    /XD __pycache__ venv .git dist build installer_build installer_output node_modules ^
    /XF *.pyc *.pyo ^
    /NFL /NDL /NJH /NJS >nul

robocopy "%REPO_DIR%\plugins" "%APP_DIR%\plugins" *.py /S ^
    /XD __pycache__ ^
    /XF *.pyc *.pyo ^
    /NFL /NDL /NJH /NJS >nul

if exist "%REPO_DIR%\assets" (
    robocopy "%REPO_DIR%\assets" "%APP_DIR%\assets" /E /NFL /NDL /NJH /NJS >nul
)

if exist "%REPO_DIR%\requirements.txt" (
    copy /Y "%REPO_DIR%\requirements.txt" "%APP_DIR%\requirements.txt" >nul
)

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

:: MCSCoWorker.vbs - double-clickable launcher, no console window
:: It finds pythonw.exe relative to itself and runs launcher.py
(
echo Set oShell = CreateObject^("WScript.Shell"^)
echo Set oFSO = CreateObject^("Scripting.FileSystemObject"^)
echo.
echo scriptDir = oFSO.GetParentFolderName^(WScript.ScriptFullName^)
echo appDir    = scriptDir ^& "\app"
echo pythonExe = scriptDir ^& "\python\pythonw.exe"
echo launcherScript = appDir ^& "\launcher.py"
echo.
echo If Not oFSO.FileExists^(pythonExe^) Then
echo     MsgBox "MCS CoWorker installation appears damaged." ^& vbCrLf ^& _
echo            "Please reinstall from SharePoint.", vbCritical, "MCS CoWorker"
echo     WScript.Quit 1
echo End If
echo.
echo oShell.CurrentDirectory = appDir
echo oShell.Run Chr^(34^) ^& pythonExe ^& Chr^(34^) ^& " " ^& Chr^(34^) ^& launcherScript ^& Chr^(34^), 0, False
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
