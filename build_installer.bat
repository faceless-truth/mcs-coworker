@echo off
setlocal EnableDelayedExpansion
title MCS CoWorker - Build Installer

echo.
echo  ============================================================
echo   MCS CoWorker - Windows Installer Builder
echo  ============================================================
echo.

:: ---------------------------------------------------------------------------
:: CONFIGURATION - paths specific to this machine
:: ---------------------------------------------------------------------------
set REPO_DIR=C:\Users\ElioScarton\mcs-coworker\
set FRONTEND_DIR=C:\Users\ElioScarton\mcs-coworker-demo
set BUILD_DIR=%REPO_DIR%installer_build
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
echo [0/7] Checking build tools...

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo         Install Python 3.11+ from https://python.org/downloads
    echo         Tick "Add Python to PATH" during install, then re-run this script.
    pause & exit /b 1
)

:: Check Git
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found.
    echo         Install Git from https://git-scm.com then re-run this script.
    pause & exit /b 1
)

:: Auto-install Inno Setup if missing
if not exist "%INNO_SETUP_DIR%\iscc.exe" (
    echo [0/7] Inno Setup not found - downloading and installing automatically...
    powershell -Command "Invoke-WebRequest -Uri '%INNO_INSTALLER_URL%' -OutFile '%INNO_INSTALLER_TMP%' -UseBasicParsing"
    if errorlevel 1 (
        echo [ERROR] Could not download Inno Setup. Check internet connection.
        pause & exit /b 1
    )
    :: Silent install - /VERYSILENT suppresses all UI
    "%INNO_INSTALLER_TMP%" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
    del "%INNO_INSTALLER_TMP%" >nul 2>&1
    :: Reset errorlevel - Inno Setup installer returns non-zero even on success with /VERYSILENT
    ver >nul
    if not exist "%INNO_SETUP_DIR%\iscc.exe" (
        echo [ERROR] Inno Setup install failed.
        pause & exit /b 1
    )
    echo [0/7] Inno Setup installed.
) else (
    echo [0/7] Inno Setup already installed.
)

:: Add Inno Setup to PATH for this session
set "PATH=%PATH%;%INNO_SETUP_DIR%"

:: Auto-install PyInstaller if missing
where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [0/7] PyInstaller not found - installing...
    pip install pyinstaller --quiet
    ver >nul
    if errorlevel 1 (
        echo [ERROR] Could not install PyInstaller.
        pause & exit /b 1
    )
    echo [0/7] PyInstaller installed.
)

:: Auto-install pnpm if missing (needed for React frontend build)
where pnpm >nul 2>&1
if errorlevel 1 (
    echo [0/7] pnpm not found - installing via npm...
    npm install -g pnpm --silent >nul 2>&1
    ver >nul
    if errorlevel 1 (
        echo [WARN] Could not install pnpm. Frontend build will be skipped.
    ) else (
        echo [0/7] pnpm installed.
    )
)

echo [0/7] All build tools ready.
echo.

:: ---------------------------------------------------------------------------
:: STEP 1 - Clean previous build
:: ---------------------------------------------------------------------------
echo [1/7] Cleaning previous build...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
mkdir "%BUILD_DIR%"
mkdir "%PYTHON_DIR%"
mkdir "%APP_DIR%"
echo [1/7] Done.
echo.

:: ---------------------------------------------------------------------------
:: STEP 2 - Download embeddable Python runtime
:: ---------------------------------------------------------------------------
echo [2/7] Downloading Python 3.11 embeddable runtime...
powershell -Command "Invoke-WebRequest -Uri '%PYTHON_EMBED_URL%' -OutFile '%PYTHON_EMBED_ZIP%' -UseBasicParsing"
if errorlevel 1 (
    echo [ERROR] Failed to download Python runtime. Check internet connection.
    pause & exit /b 1
)
powershell -Command "Expand-Archive -Path '%PYTHON_EMBED_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force"
del "%PYTHON_EMBED_ZIP%"
echo [2/7] Python runtime ready.
echo.

:: ---------------------------------------------------------------------------
:: STEP 3 - Enable pip in embeddable Python
:: ---------------------------------------------------------------------------
echo [3/7] Enabling pip in embedded Python...
powershell -Command "(Get-Content '%PYTHON_DIR%\python311._pth') -replace '#import site','import site' | Set-Content '%PYTHON_DIR%\python311._pth'"
powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%BUILD_DIR%\get-pip.py' -UseBasicParsing"
"%PYTHON_DIR%\python.exe" "%BUILD_DIR%\get-pip.py" --quiet
del "%BUILD_DIR%\get-pip.py"
echo [3/7] pip enabled.
echo.

:: ---------------------------------------------------------------------------
:: STEP 4 - Install Python dependencies into embedded runtime
:: ---------------------------------------------------------------------------
echo [4/7] Installing Python dependencies (this takes a few minutes)...
"%PYTHON_DIR%\python.exe" -m pip install -r "%REPO_DIR%requirements.txt" --quiet --no-warn-script-location
"%PYTHON_DIR%\python.exe" -m pip install pywebview flask flask-cors pdfminer.six --quiet --no-warn-script-location
if errorlevel 1 (
    echo [ERROR] pip install failed. Check requirements.txt and internet connection.
    pause & exit /b 1
)
echo [4/7] Dependencies installed.
echo.

:: ---------------------------------------------------------------------------
:: STEP 5 - Build React frontend
:: ---------------------------------------------------------------------------
echo [5/7] Building React frontend...
if not exist "%FRONTEND_DIR%\package.json" (
    echo [WARN] Frontend not found at %FRONTEND_DIR% - skipping.
    echo        The app will use the dev server on first run.
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
    echo [5/7] Frontend built and copied.
)
echo.

:: ---------------------------------------------------------------------------
:: STEP 6 - Copy app source
:: ---------------------------------------------------------------------------
echo [6/7] Copying app source...
robocopy "%REPO_DIR%" "%APP_DIR%" *.py /S /XD __pycache__ venv .git dist build installer_build /XF *.pyc /NFL /NDL /NJH /NJS >nul
robocopy "%REPO_DIR%plugins" "%APP_DIR%\plugins" *.py /S /XD __pycache__ /XF *.pyc /NFL /NDL /NJH /NJS >nul
if exist "%REPO_DIR%assets" robocopy "%REPO_DIR%assets" "%APP_DIR%\assets" /E /NFL /NDL /NJH /NJS >nul
copy /Y "%REPO_DIR%requirements.txt" "%APP_DIR%\requirements.txt" >nul

for /f %%i in ('git -C "%REPO_DIR%" rev-parse --short HEAD 2^>nul') do set GIT_HASH=%%i
if defined GIT_HASH (
    echo %GIT_HASH%> "%APP_DIR%\VERSION"
    echo [6/7] App source copied. Version: %GIT_HASH%
) else (
    echo [6/7] App source copied.
)
echo.

:: ---------------------------------------------------------------------------
:: STEP 6b - Build launcher.exe with PyInstaller
:: ---------------------------------------------------------------------------
echo [6b] Compiling launcher.exe...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "MCSCoWorker" ^
    --distpath "%BUILD_DIR%\launcher_dist" ^
    --workpath "%BUILD_DIR%\launcher_work" ^
    --specpath "%BUILD_DIR%" ^
    "%REPO_DIR%launcher.py" >nul 2>&1
if errorlevel 1 (
    echo [WARN] PyInstaller failed without icon - trying without icon flag...
    pyinstaller ^
        --onefile ^
        --windowed ^
        --name "MCSCoWorker" ^
        --distpath "%BUILD_DIR%\launcher_dist" ^
        --workpath "%BUILD_DIR%\launcher_work" ^
        --specpath "%BUILD_DIR%" ^
        "%REPO_DIR%launcher.py"
    if errorlevel 1 (
        echo [ERROR] PyInstaller failed. See output above.
        pause & exit /b 1
    )
)
copy /Y "%BUILD_DIR%\launcher_dist\MCSCoWorker.exe" "%BUILD_DIR%\MCSCoWorker.exe" >nul
echo [6b] Launcher compiled.
echo.

:: ---------------------------------------------------------------------------
:: STEP 7 - Build installer with Inno Setup
:: ---------------------------------------------------------------------------
echo [7/7] Building installer with Inno Setup...
iscc "%REPO_DIR%installer.iss" /DMyBuildDir="%BUILD_DIR%"
if errorlevel 1 (
    echo [ERROR] Inno Setup failed. See output above.
    pause & exit /b 1
)
echo [7/7] Installer built.
echo.

:: ---------------------------------------------------------------------------
:: DONE
:: ---------------------------------------------------------------------------
echo  ============================================================
echo   Build complete!
echo.
echo   Installer: %REPO_DIR%installer_output\MCSCoWorker_Setup.exe
echo.
echo   Upload MCSCoWorker_Setup.exe to SharePoint.
echo   Accountants download and double-click to install.
echo   Future updates: just git push - no reinstall needed.
echo  ============================================================
echo.
pause
