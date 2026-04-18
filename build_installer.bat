@echo off
setlocal EnableDelayedExpansion
title MCS CoWorker — Build Installer

echo.
echo  ============================================================
echo   MCS CoWorker — Windows Installer Builder
echo  ============================================================
echo.

:: ── Prerequisites check ────────────────────────────────────────────────────
echo [CHECK] Verifying prerequisites...

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.11+ and add to PATH.
    pause & exit /b 1
)

git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found. Install Git for Windows.
    pause & exit /b 1
)

where iscc >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Inno Setup compiler (iscc) not found.
    echo         Download from https://jrsoftware.org/isdl.php
    echo         Then add "C:\Program Files (x86)\Inno Setup 6" to PATH.
    pause & exit /b 1
)

echo [CHECK] All prerequisites found.
echo.

:: ── Configuration ──────────────────────────────────────────────────────────
set REPO_DIR=%~dp0
set BUILD_DIR=%REPO_DIR%installer_build
set PYTHON_EMBED_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip
set PYTHON_EMBED_ZIP=%BUILD_DIR%\python-embed.zip
set PYTHON_DIR=%BUILD_DIR%\python
set APP_DIR=%BUILD_DIR%\app
set FRONTEND_DIR=%REPO_DIR%..\mcs-coworker-demo

:: ── Clean previous build ───────────────────────────────────────────────────
echo [1/7] Cleaning previous build...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
mkdir "%BUILD_DIR%"
mkdir "%PYTHON_DIR%"
mkdir "%APP_DIR%"
echo [1/7] Done.
echo.

:: ── Download embeddable Python ─────────────────────────────────────────────
echo [2/7] Downloading Python 3.11 embeddable runtime...
powershell -Command "Invoke-WebRequest -Uri '%PYTHON_EMBED_URL%' -OutFile '%PYTHON_EMBED_ZIP%' -UseBasicParsing"
if errorlevel 1 (
    echo [ERROR] Failed to download Python runtime.
    pause & exit /b 1
)
powershell -Command "Expand-Archive -Path '%PYTHON_EMBED_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force"
del "%PYTHON_EMBED_ZIP%"
echo [2/7] Python runtime ready.
echo.

:: ── Enable pip in embeddable Python ───────────────────────────────────────
echo [3/7] Enabling pip in embedded Python...
:: Uncomment the import site line in python311._pth
powershell -Command "(Get-Content '%PYTHON_DIR%\python311._pth') -replace '#import site','import site' | Set-Content '%PYTHON_DIR%\python311._pth'"
:: Download get-pip.py
powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%BUILD_DIR%\get-pip.py' -UseBasicParsing"
"%PYTHON_DIR%\python.exe" "%BUILD_DIR%\get-pip.py" --quiet
del "%BUILD_DIR%\get-pip.py"
echo [3/7] pip enabled.
echo.

:: ── Install Python dependencies into embedded runtime ─────────────────────
echo [4/7] Installing Python dependencies...
"%PYTHON_DIR%\python.exe" -m pip install -r "%REPO_DIR%requirements.txt" --quiet --no-warn-script-location
"%PYTHON_DIR%\python.exe" -m pip install pywebview flask flask-cors --quiet --no-warn-script-location
if errorlevel 1 (
    echo [ERROR] pip install failed. Check requirements.txt.
    pause & exit /b 1
)
echo [4/7] Dependencies installed.
echo.

:: ── Build React frontend ───────────────────────────────────────────────────
echo [5/7] Building React frontend...
if not exist "%FRONTEND_DIR%\package.json" (
    echo [WARN] Frontend not found at %FRONTEND_DIR% — skipping.
    echo        The app will use the dev server on first run.
) else (
    pushd "%FRONTEND_DIR%"
    call pnpm install --frozen-lockfile --silent
    call pnpm build --silent
    if errorlevel 1 (
        popd
        echo [ERROR] Frontend build failed.
        pause & exit /b 1
    )
    popd
    xcopy /E /I /Y /Q "%FRONTEND_DIR%\dist" "%APP_DIR%\frontend_dist"
    echo [5/7] Frontend built and copied.
)
echo.

:: ── Copy app source into build dir ────────────────────────────────────────
echo [6/7] Copying app source...
:: Copy all Python files and assets — exclude dev/cache artefacts
robocopy "%REPO_DIR%" "%APP_DIR%" *.py /S /XD __pycache__ venv .git dist build installer_build /XF *.pyc /NFL /NDL /NJH /NJS >nul
robocopy "%REPO_DIR%plugins" "%APP_DIR%\plugins" *.py /S /XD __pycache__ /XF *.pyc /NFL /NDL /NJH /NJS >nul
robocopy "%REPO_DIR%assets" "%APP_DIR%\assets" /E /NFL /NDL /NJH /NJS >nul 2>&1
copy /Y "%REPO_DIR%requirements.txt" "%APP_DIR%\requirements.txt" >nul

:: Write current version
for /f %%i in ('git -C "%REPO_DIR%" rev-parse --short HEAD 2^>nul') do set GIT_HASH=%%i
if defined GIT_HASH (
    echo %GIT_HASH%> "%APP_DIR%\VERSION"
    echo [6/7] App source copied. Version: %GIT_HASH%
) else (
    echo [6/7] App source copied. (Could not determine git version)
)
echo.

:: ── Build launcher.exe with PyInstaller ───────────────────────────────────
echo [6b] Building launcher.exe...
pip install pyinstaller --quiet >nul 2>&1
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "MCSCoWorker" ^
    --icon "%REPO_DIR%assets\icon.ico" ^
    --distpath "%BUILD_DIR%\launcher_dist" ^
    --workpath "%BUILD_DIR%\launcher_work" ^
    --specpath "%BUILD_DIR%" ^
    "%REPO_DIR%launcher.py"
if errorlevel 1 (
    echo [ERROR] PyInstaller failed.
    pause & exit /b 1
)
copy /Y "%BUILD_DIR%\launcher_dist\MCSCoWorker.exe" "%BUILD_DIR%\MCSCoWorker.exe" >nul
echo [6b] Launcher built.
echo.

:: ── Run Inno Setup ────────────────────────────────────────────────────────
echo [7/7] Building installer with Inno Setup...
iscc "%REPO_DIR%installer.iss" /DMyBuildDir="%BUILD_DIR%"
if errorlevel 1 (
    echo [ERROR] Inno Setup failed.
    pause & exit /b 1
)
echo [7/7] Installer built.
echo.

:: ── Done ──────────────────────────────────────────────────────────────────
echo  ============================================================
echo   Build complete!
echo.
echo   Installer: %REPO_DIR%installer_output\MCSCoWorker_Setup.exe
echo.
echo   Upload MCSCoWorker_Setup.exe to SharePoint.
echo   Accountants download and double-click to install.
echo   Future updates: just git push — no reinstall needed.
echo  ============================================================
echo.
pause
