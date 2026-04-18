/**
 * EVA — Electron Main Process
 * Launches the Python Flask API server, then opens the React UI in a frameless window.
 * Supports system tray, auto-start on login, and graceful shutdown.
 */

const { app, BrowserWindow, Tray, Menu, nativeImage, ipcMain, shell } = require("electron");
const path = require("path");
const { spawn } = require("child_process");
const http = require("http");
const fs = require("fs");

// ── Constants ──────────────────────────────────────────────────────────────────
const API_PORT = 7842;
const API_URL = `http://127.0.0.1:${API_PORT}`;
const REACT_BUILD = path.join(__dirname, "..", "frontend", "dist");
const PYTHON_ENTRY = path.join(__dirname, "..", "api_server_standalone.py");
const APP_NAME = "EVA";

let mainWindow = null;
let tray = null;
let pythonProcess = null;
let isQuitting = false;

// ── Python API Server ──────────────────────────────────────────────────────────
function startPythonServer() {
  return new Promise((resolve, reject) => {
    // Find python executable
    const pythonCmd = process.platform === "win32" ? "python" : "python3.11";
    const cwd = path.join(__dirname, "..");

    console.log(`[Electron] Starting Python API server: ${pythonCmd} ${PYTHON_ENTRY}`);

    pythonProcess = spawn(pythonCmd, [PYTHON_ENTRY], {
      cwd,
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    });

    pythonProcess.stdout.on("data", (data) => {
      console.log(`[Python] ${data.toString().trim()}`);
    });

    pythonProcess.stderr.on("data", (data) => {
      const msg = data.toString().trim();
      if (msg) console.error(`[Python ERR] ${msg}`);
    });

    pythonProcess.on("exit", (code) => {
      if (!isQuitting) {
        console.error(`[Python] Process exited unexpectedly with code ${code}`);
      }
    });

    // Poll until the API is ready
    let attempts = 0;
    const poll = setInterval(() => {
      attempts++;
      http.get(`${API_URL}/api/health`, (res) => {
        if (res.statusCode === 200) {
          clearInterval(poll);
          console.log(`[Electron] Python API ready after ${attempts} attempts`);
          resolve();
        }
      }).on("error", () => {
        if (attempts > 60) {
          clearInterval(poll);
          reject(new Error("Python API failed to start after 30 seconds"));
        }
      });
    }, 500);
  });
}

function stopPythonServer() {
  if (pythonProcess) {
    console.log("[Electron] Stopping Python API server...");
    if (process.platform === "win32") {
      spawn("taskkill", ["/pid", pythonProcess.pid.toString(), "/f", "/t"]);
    } else {
      pythonProcess.kill("SIGTERM");
    }
    pythonProcess = null;
  }
}

// ── Main Window ────────────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    frame: false,           // Custom title bar
    titleBarStyle: "hidden",
    backgroundColor: "#0f172a",
    show: false,
    icon: path.join(__dirname, "assets", "icon.png"),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      // Allow loading from file:// and calling http://127.0.0.1:7842
      webSecurity: false,
    },
  });

  // Load the React build
  const indexPath = path.join(REACT_BUILD, "index.html");
  if (fs.existsSync(indexPath)) {
    mainWindow.loadFile(indexPath);
  } else {
    // Fallback: load from dev server (for development)
    mainWindow.loadURL("http://localhost:3000");
  }

  // Show once ready to avoid white flash
  mainWindow.once("ready-to-show", () => {
    mainWindow.show();
    if (process.env.NODE_ENV === "development") {
      mainWindow.webContents.openDevTools({ mode: "detach" });
    }
  });

  // Minimise to tray instead of closing
  mainWindow.on("close", (e) => {
    if (!isQuitting) {
      e.preventDefault();
      mainWindow.hide();
      if (process.platform === "win32") {
        tray.displayBalloon({
          iconType: "info",
          title: APP_NAME,
          content: "CoWorker is still running in the background.",
        });
      }
    }
  });

  // Open external links in browser, not Electron
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });
}

// ── System Tray ────────────────────────────────────────────────────────────────
function createTray() {
  const iconPath = path.join(__dirname, "assets", "tray-icon.png");
  const icon = fs.existsSync(iconPath)
    ? nativeImage.createFromPath(iconPath).resize({ width: 16, height: 16 })
    : nativeImage.createEmpty();

  tray = new Tray(icon);
  tray.setToolTip(APP_NAME);

  const contextMenu = Menu.buildFromTemplate([
    {
      label: "Open CoWorker",
      click: () => {
        mainWindow.show();
        mainWindow.focus();
      },
    },
    { type: "separator" },
    {
      label: "Dashboard",
      click: () => {
        mainWindow.show();
        mainWindow.webContents.send("navigate", "dashboard");
      },
    },
    {
      label: "Approvals",
      click: () => {
        mainWindow.show();
        mainWindow.webContents.send("navigate", "approvals");
      },
    },
    { type: "separator" },
    {
      label: "Quit CoWorker",
      click: () => {
        isQuitting = true;
        app.quit();
      },
    },
  ]);

  tray.setContextMenu(contextMenu);
  tray.on("double-click", () => {
    mainWindow.show();
    mainWindow.focus();
  });
}

// ── IPC Handlers ───────────────────────────────────────────────────────────────
ipcMain.on("window-minimize", () => mainWindow?.minimize());
ipcMain.on("window-maximize", () => {
  if (mainWindow?.isMaximized()) {
    mainWindow.unmaximize();
  } else {
    mainWindow?.maximize();
  }
});
ipcMain.on("window-close", () => mainWindow?.hide());
ipcMain.on("window-quit", () => {
  isQuitting = true;
  app.quit();
});

// ── Auto-start on Login ────────────────────────────────────────────────────────
function configureAutoStart() {
  app.setLoginItemSettings({
    openAtLogin: true,
    openAsHidden: true,   // Start minimised to tray
    name: APP_NAME,
  });
}

// ── App Lifecycle ──────────────────────────────────────────────────────────────
app.whenReady().then(async () => {
  // Set app user model ID for Windows notifications
  if (process.platform === "win32") {
    app.setAppUserModelId("au.com.mcsaccountants.coworker");
  }

  // Inject __ELECTRON__ flag so the React app knows it's in desktop mode
  app.on("web-contents-created", (_, contents) => {
    contents.on("dom-ready", () => {
      contents.executeJavaScript("window.__ELECTRON__ = true;");
    });
  });

  createTray();
  configureAutoStart();

  try {
    await startPythonServer();
  } catch (e) {
    console.error("[Electron] Failed to start Python server:", e.message);
    // Still open the window — it will show mock data
  }

  createWindow();
});

app.on("window-all-closed", () => {
  // On macOS keep running; on Windows/Linux quit
  if (process.platform !== "darwin") {
    isQuitting = true;
    app.quit();
  }
});

app.on("activate", () => {
  if (mainWindow === null) {
    createWindow();
  } else {
    mainWindow.show();
  }
});

app.on("before-quit", () => {
  isQuitting = true;
  stopPythonServer();
});
