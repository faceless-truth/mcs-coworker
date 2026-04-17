# MCS CoWorker — Windows Setup Guide

## Prerequisites

Install these once on each accountant's machine:

1. **Python 3.11+** — https://www.python.org/downloads/
   - During install: tick **"Add Python to PATH"**
2. **Node.js 20+** — https://nodejs.org/
3. **Git** — https://git-scm.com/download/win

---

## First-Time Setup

Open **Command Prompt** or **PowerShell** and run:

```bat
:: 1. Clone the repository
git clone https://github.com/faceless-truth/mcs-coworker.git C:\MCSCoWorker
cd C:\MCSCoWorker

:: 2. Install Python dependencies
pip install -r requirements.txt

:: 3. Install and build the React frontend
cd frontend
npm install
npm run build
cd ..

:: 4. Install Electron dependencies
cd electron
npm install
cd ..
```

---

## Running CoWorker

### Option A — Run directly (development / testing)
```bat
cd C:\MCSCoWorker\electron
npx electron .
```

### Option B — Build a proper Windows installer
```bat
cd C:\MCSCoWorker\electron
npm run build
```
This creates `dist-electron\MCS CoWorker Setup.exe` — install it like any Windows app.

---

## Updating to the Latest Version

```bat
cd C:\MCSCoWorker
git pull origin main
pip install -r requirements.txt
cd frontend && npm run build && cd ..
```

---

## Per-Accountant Configuration

After launching CoWorker for the first time:

1. Go to **Settings → XPM / Xero Practice Manager**
2. Click **Connect Xero** and log in with your Xero account
3. CoWorker will automatically scope all client data to your manager portfolio

Each accountant connects with their own Xero credentials — their CoWorker instance only sees their own clients.

---

## Auto-Start on Login

CoWorker configures itself to start automatically with Windows. To disable this:
- Open **Task Manager → Startup apps** and disable **MCS CoWorker**

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Python not found" | Reinstall Python with "Add to PATH" ticked |
| "Module not found" | Run `pip install -r requirements.txt` again |
| Blank white window | Wait 10 seconds for the Python server to start |
| Can't connect to Xero | Check Settings → XPM and click Connect Xero |
