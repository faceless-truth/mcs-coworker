# EVA — Windows Setup Guide

## For Accountants: Installing the App

1. Download **EVA_Setup.exe** from the MC&S SharePoint.
2. Double-click the file and follow the installer prompts.
3. Choose whether to add a Desktop shortcut and whether to start automatically with Windows.
4. Click **Finish** — the app will launch immediately.

**That is all.** No Python, no Git, no technical knowledge required.

The app updates itself silently every time you open it. You will always have the latest version without doing anything.

---

## For Elio: Building the Installer

You only need to rebuild the installer when Python package dependencies change (rare — maybe once every few months). For normal bug fixes and new features, just `git push` and every machine picks it up automatically.

### Prerequisites (one-time setup on the build machine)

| Tool | Download | Notes |
|------|----------|-------|
| Python 3.11+ | https://python.org/downloads | Tick "Add to PATH" |
| Git for Windows | https://git-scm.com | Default options |
| Inno Setup 6 | https://jrsoftware.org/isdl.php | Add to PATH: `C:\Program Files (x86)\Inno Setup 6` |
| Node.js + pnpm | https://nodejs.org | For building the React frontend |

### Build steps

```bat
cd C:\Users\ElioScarton\mcs-coworker
build_installer.bat
```

The script will:
1. Download the Python 3.11 embeddable runtime automatically
2. Install all Python dependencies into it
3. Build the React frontend
4. Copy the app source
5. Compile `EVA.exe` (the launcher) with PyInstaller
6. Run Inno Setup to produce `installer_output\EVA_Setup.exe`

Upload `EVA_Setup.exe` to SharePoint. Done.

### Build time

Approximately 10–15 minutes on first run (downloads Python runtime and all packages). Subsequent builds are faster (~5 minutes) because pip caches packages.

---

## How Updates Work

### Day-to-day bug fixes and new features

```bash
# Fix the bug
git add -A
git commit -m "fix: description of fix"
git push origin main
```

Every accountant's machine picks up the fix automatically the next time they open the app. No action required from them.

### New Python package required

1. Add the package to `requirements.txt`
2. Add it to `hiddenimports` in `build.spec` if it uses dynamic imports
3. Run `build_installer.bat` to produce a new installer
4. Upload the new installer to SharePoint
5. Ask accountants to download and reinstall once

This should happen rarely — the existing package set covers almost all needs.

---

## What Gets Installed Where

```
C:\Program Files\EVA\
    EVA.exe          ← launcher (compiled, rarely changes)
    python\                  ← bundled Python 3.11 runtime
    app\                     ← the git repo (auto-updates from GitHub)
        main.py
        auto_updater.py
        plugins\
        requirements.txt
        VERSION              ← current commit hash
    data\                    ← user data (never touched by updates)
        coworker.db          ← all settings, memory, approvals, logs
    assets\
        icon.ico
```

The `data\` folder is never modified by updates — settings, memory, and history are always preserved.

---

## Configuring a New Installation

After installing, open the app and go to **Settings**:

| Setting | What to enter |
|---------|--------------|
| `user_name` | The accountant's name (e.g. `Elio`) |
| `user_email` | Their Microsoft 365 email address |
| `staff_profile` | Their profile: `elio`, `ross`, `harry`, `brooke`, `louise`, `lyn`, or `reception` |
| `reception_mode` | `1` for the reception inbox, `0` for everyone else |
| `practice_name` | `MC & S Accountants` |
| `xero_staff_uuid` | From XPM — filters jobs and clients to this person only |

Then connect Microsoft 365 (OAuth) and Xero Practice Manager from the Connections tab.

---

## Troubleshooting

**App opens but shows a blank screen**
The React frontend may not have built correctly. Check that `app\frontend_dist\` exists inside the install folder.

**"Could not clone the app from GitHub" during install**
The machine needs internet access to GitHub (github.com port 443). Check firewall rules.

**Update failed on launch**
The app will continue running with the previous version. Check the update history in Settings > About. If git pull is consistently failing, run `git status` in `C:\Program Files\EVA\app\` to check for conflicts.

**Uninstall**
Use Windows Settings > Apps > EVA > Uninstall. The `data\` folder (containing `coworker.db`) is preserved — delete it manually if you want a clean removal.
