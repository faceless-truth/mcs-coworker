"""
MCS CoWorker — Auto-Update System (Tier 3D)
============================================
Checks GitHub for new releases and can apply updates automatically
or notify the team when a new version is available.

APEX ALIGNMENT
--------------
APEX is designed to be continuously improved and self-updating. This module
gives CoWorker the same capability — checking for new releases on GitHub
and applying them without manual intervention.

BEHAVIOUR
---------
1. On startup and every 6 hours, checks the GitHub releases API for the
   faceless-truth/mcs-coworker repository.
2. Compares the latest release tag against the current VERSION constant.
3. If a newer version is found:
   a. Notifies via Teams and/or the app log.
   b. If auto_update=True (default: False), downloads the release asset
      and applies it by replacing plugin files.
4. All update history is logged to SQLite.

VERSIONING
----------
Version format: MAJOR.MINOR.PATCH (e.g. "1.2.0")
The current version is read from VERSION in this file or from a
VERSION file in the repo root if present.
"""

from __future__ import annotations
import json
import logging
import os
import re
import shutil
import sqlite3
import tempfile
import threading
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from urllib.request import urlopen, Request
from urllib.error import URLError

logger = logging.getLogger(__name__)

# ── Version ───────────────────────────────────────────────────────────────────

VERSION = "1.0.0"  # Current version — update this on each release

GITHUB_REPO   = "faceless-truth/mcs-coworker"
GITHUB_API    = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
CHECK_INTERVAL_H = 6  # hours between checks


def get_current_version() -> str:
    """Return the current version string."""
    # Check for a VERSION file in the repo root (overrides the constant)
    try:
        version_file = Path(__file__).parent / "VERSION"
        if version_file.exists():
            return version_file.read_text().strip()
    except Exception:
        pass
    return VERSION


def parse_version(version_str: str) -> tuple[int, ...]:
    """Parse a version string into a comparable tuple."""
    try:
        # Strip leading 'v' if present
        clean = version_str.lstrip("v").strip()
        parts = re.split(r"[.\-]", clean)
        return tuple(int(p) for p in parts if p.isdigit())
    except Exception:
        return (0,)


def is_newer(latest: str, current: str) -> bool:
    """Return True if latest version is newer than current."""
    return parse_version(latest) > parse_version(current)


# ── Database helpers ──────────────────────────────────────────────────────────

def _get_updater_db_path() -> Path:
    try:
        import config as cfg
        return Path(cfg.DB_PATH).parent / "auto_updater.db"
    except Exception:
        return Path.home() / ".mcs_coworker" / "auto_updater.db"


def init_updater_db(db_path: Optional[Path] = None) -> None:
    """Create the update history table."""
    path = db_path or _get_updater_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS update_history (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp       TEXT    DEFAULT (datetime('now','localtime')),
            current_version TEXT,
            latest_version  TEXT,
            action          TEXT,
            success         INTEGER DEFAULT 1,
            notes           TEXT
        );
        CREATE TABLE IF NOT EXISTS update_config (
            key     TEXT PRIMARY KEY,
            value   TEXT
        );
    """)
    # Seed defaults
    conn.execute("INSERT OR IGNORE INTO update_config (key, value) VALUES ('auto_update', '0')")
    conn.execute("INSERT OR IGNORE INTO update_config (key, value) VALUES ('last_check', '')")
    conn.execute("INSERT OR IGNORE INTO update_config (key, value) VALUES ('last_notified_version', '')")
    conn.commit()
    conn.close()


def _get_config(key: str, default: str = "",
                db_path: Optional[Path] = None) -> str:
    path = db_path or _get_updater_db_path()
    if not path.exists():
        return default
    conn = sqlite3.connect(str(path))
    row = conn.execute(
        "SELECT value FROM update_config WHERE key=?", (key,)
    ).fetchone()
    conn.close()
    return row[0] if row else default


def _set_config(key: str, value: str, db_path: Optional[Path] = None) -> None:
    path = db_path or _get_updater_db_path()
    conn = sqlite3.connect(str(path))
    conn.execute(
        "INSERT OR REPLACE INTO update_config (key, value) VALUES (?,?)",
        (key, value)
    )
    conn.commit()
    conn.close()


def _log_update(current: str, latest: str, action: str,
                success: bool = True, notes: str = "",
                db_path: Optional[Path] = None) -> None:
    path = db_path or _get_updater_db_path()
    conn = sqlite3.connect(str(path))
    conn.execute(
        "INSERT INTO update_history (current_version, latest_version, action, success, notes) "
        "VALUES (?,?,?,?,?)",
        (current, latest, action, int(success), notes)
    )
    conn.commit()
    conn.close()


def get_update_history(limit: int = 20,
                       db_path: Optional[Path] = None) -> list[dict]:
    """Return recent update history."""
    path = db_path or _get_updater_db_path()
    if not path.exists():
        return []
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT * FROM update_history ORDER BY timestamp DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ── GitHub API ────────────────────────────────────────────────────────────────

def check_for_update(db_path: Optional[Path] = None) -> Optional[dict]:
    """
    Query the GitHub releases API for the latest release.
    Returns a dict with release info if a newer version is available,
    or None if already up to date or the check fails.
    """
    current = get_current_version()
    _set_config("last_check",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"), db_path)

    try:
        req = Request(GITHUB_API,
                      headers={"Accept": "application/vnd.github.v3+json",
                                "User-Agent": "MCSCoWorker-AutoUpdater"})
        with urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode())

        latest_tag  = data.get("tag_name", "").lstrip("v")
        release_url = data.get("html_url", "")
        body        = data.get("body", "")
        assets      = data.get("assets", [])
        zip_url     = next(
            (a["browser_download_url"] for a in assets
             if a["name"].endswith(".zip")), None)

        if not latest_tag:
            return None

        if is_newer(latest_tag, current):
            logger.info(f"[AutoUpdater] New version available: {latest_tag} "
                        f"(current: {current})")
            return {
                "current_version": current,
                "latest_version":  latest_tag,
                "release_url":     release_url,
                "release_notes":   body[:500],
                "zip_url":         zip_url,
            }
        else:
            logger.debug(f"[AutoUpdater] Up to date: {current}")
            return None

    except URLError as e:
        logger.debug(f"[AutoUpdater] Network error checking for updates: {e}")
        return None
    except Exception as e:
        logger.warning(f"[AutoUpdater] Update check failed: {e}")
        return None


# ── Plugin-only updater ───────────────────────────────────────────────────────

def apply_plugin_update(zip_url: str, repo_dir: Optional[Path] = None,
                        db_path: Optional[Path] = None) -> bool:
    """
    Download a release ZIP and replace plugin files only.
    Does NOT replace core files (app.py, config.py, etc.) for safety.
    Returns True on success.
    """
    if repo_dir is None:
        repo_dir = Path(__file__).parent

    plugins_dir = repo_dir / "plugins"
    current     = get_current_version()

    try:
        logger.info(f"[AutoUpdater] Downloading update from {zip_url}")
        req = Request(zip_url, headers={"User-Agent": "MCSCoWorker-AutoUpdater"})
        with tempfile.TemporaryDirectory() as tmp:
            zip_path = Path(tmp) / "update.zip"
            with urlopen(req, timeout=60) as resp:
                zip_path.write_bytes(resp.read())

            with zipfile.ZipFile(zip_path) as zf:
                # Only extract plugin_*.py files
                plugin_files = [
                    name for name in zf.namelist()
                    if re.search(r"plugins/plugin_[^/]+\.py$", name)
                ]
                if not plugin_files:
                    logger.warning("[AutoUpdater] No plugin files found in ZIP")
                    _log_update(current, "?", "download_no_plugins",
                                success=False, db_path=db_path)
                    return False

                # Backup existing plugins
                backup_dir = repo_dir / f"plugins_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                shutil.copytree(plugins_dir, backup_dir)
                logger.info(f"[AutoUpdater] Plugins backed up to {backup_dir}")

                # Extract new plugins
                for name in plugin_files:
                    filename = Path(name).name
                    dest = plugins_dir / filename
                    with zf.open(name) as src, open(dest, "wb") as dst:
                        dst.write(src.read())
                    logger.info(f"[AutoUpdater] Updated: {filename}")

        _log_update(current, "latest", "plugin_update_applied",
                    success=True,
                    notes=f"Updated {len(plugin_files)} plugin files",
                    db_path=db_path)
        return True

    except Exception as e:
        logger.error(f"[AutoUpdater] Update failed: {e}")
        _log_update(current, "?", "plugin_update_failed",
                    success=False, notes=str(e), db_path=db_path)
        return False


# ── AutoUpdater singleton ─────────────────────────────────────────────────────

class _AutoUpdater:
    """
    Singleton that periodically checks for updates and optionally applies them.
    """

    def __init__(self):
        self._db_path: Optional[Path] = None
        self._context = None
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._on_update_available = None  # callback(info: dict)

    def start(self, context=None,
              on_update_available=None,
              db_path: Optional[Path] = None) -> None:
        """Start the background update checker."""
        self._context  = context
        self._db_path  = db_path
        self._on_update_available = on_update_available
        init_updater_db(db_path)

        if not self._running:
            self._running = True
            self._thread = threading.Thread(
                target=self._loop, daemon=True, name="AutoUpdater")
            self._thread.start()
            logger.info("[AutoUpdater] Started.")

    def stop(self) -> None:
        self._running = False

    def check_now(self) -> Optional[dict]:
        """Trigger an immediate update check and return result."""
        return self._do_check()

    def _do_check(self) -> Optional[dict]:
        info = check_for_update(self._db_path)
        if info:
            last_notified = _get_config("last_notified_version", "", self._db_path)
            if info["latest_version"] != last_notified:
                _set_config("last_notified_version",
                            info["latest_version"], self._db_path)
                _log_update(info["current_version"], info["latest_version"],
                            "update_available", db_path=self._db_path)
                self._notify(info)
                if self._on_update_available:
                    try:
                        self._on_update_available(info)
                    except Exception as e:
                        logger.debug(f"[AutoUpdater] Callback error: {e}")

                # Auto-update plugins if enabled
                auto = _get_config("auto_update", "0", self._db_path)
                if auto == "1" and info.get("zip_url"):
                    logger.info("[AutoUpdater] Auto-update enabled — applying update.")
                    apply_plugin_update(
                        info["zip_url"], db_path=self._db_path)
        return info

    def _notify(self, info: dict) -> None:
        """Notify via Teams and app log."""
        msg = (f"New version {info['latest_version']} available "
               f"(current: {info['current_version']}). "
               f"Release notes: {info.get('release_notes', '')[:200]}")

        if self._context:
            # Teams notification
            if (self._context.gateway and
                    self._context.gateway.is_available("teams")):
                try:
                    self._context.gateway.teams.send_alert(
                        title="🆕 CoWorker Update Available",
                        message=msg,
                        color="info")
                except Exception as e:
                    logger.debug(f"[AutoUpdater] Teams notify failed: {e}")

            # Email notification
            if self._context.notify:
                try:
                    self._context.notify(
                        subject=f"CoWorker Update: v{info['latest_version']} available",
                        body=msg)
                except Exception as e:
                    logger.debug(f"[AutoUpdater] Email notify failed: {e}")

        logger.info(f"[AutoUpdater] {msg}")

    def _loop(self) -> None:
        """Background thread: check every CHECK_INTERVAL_H hours."""
        # Initial check after 30 seconds (let the app finish starting)
        import time
        time.sleep(30)
        while self._running:
            try:
                self._do_check()
            except Exception as e:
                logger.warning(f"[AutoUpdater] Loop error: {e}")
            # Sleep in 60s slices so we can stop quickly
            for _ in range(CHECK_INTERVAL_H * 60):
                if not self._running:
                    break
                time.sleep(60)

    def get_status(self) -> dict:
        """Return current updater status for the UI."""
        return {
            "current_version": get_current_version(),
            "last_check":      _get_config("last_check", "", self._db_path),
            "auto_update":     _get_config("auto_update", "0", self._db_path) == "1",
            "history":         get_update_history(10, self._db_path),
        }

    def set_auto_update(self, enabled: bool) -> None:
        """Enable or disable automatic plugin updates."""
        _set_config("auto_update", "1" if enabled else "0", self._db_path)


# Module-level singleton
AutoUpdater = _AutoUpdater()
