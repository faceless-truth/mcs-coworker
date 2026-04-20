@echo off
:: ============================================================
:: MCS CoWorker - Emergency Launch Fix
:: Double-click this to patch the installed app.
:: ============================================================

set INSTALL_APP=%LOCALAPPDATA%\Programs\MCS CoWorker\app

echo Patching main.py...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/faceless-truth/mcs-coworker/main/main.py' -OutFile '%INSTALL_APP%\main.py' -UseBasicParsing"

if errorlevel 1 (
    echo.
    echo ERROR: Could not download the fix. Check your internet connection.
    pause
    exit /b 1
)

echo.
echo Done! main.py has been updated.
echo You can now launch MCS CoWorker from the Start Menu or desktop shortcut.
echo.
pause
