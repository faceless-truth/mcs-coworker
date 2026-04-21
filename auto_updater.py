"""
MCS CoWorker — Auto-Update System (Hybrid Installer Edition)
=============================================================
Performs a silent fetch + author-verified reset + pip install on every launch
so that pushing a fix to GitHub is all that is needed to update every
accountant's machine. No reinstall required.

BEHAVIOUR
---------
1. On launch, check if a newer commit exists on GitHub main branch.
2. If yes (or if forced):
   a. Back up the current app folder (full tree, excluding venv/.git/caches).
   b. `git fetch origin main`, verify commit author is in
      ALLOWED_UPDATE_AUTHORS, then `git reset --hard origin/main`.
   c. Post-update sanity check on critical files — if any are missing or
      empty, restore from backup immediately.
   d. Run `pip install -r requirements.txt --quiet`.
   e. Write the new commit hash to VERSION.
   f. Set restart_flag so the launcher can restart the app.
3. If fetch/verify/reset fails, restore the backup and continue with the
   existing version — the accountant always has a working app.
4. Every 6 hours while the app is running, re-check and apply
   updates silently in the background (takes effect on next restart).

VERSIONING
----------
Version is the short git commit hash (e.g. "d70588d").
Stored in the VERSION file in the repo root.
"""
from __future__ import annotations

import json
import logging
import os
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.request import urlopen, Request
from urllib.error import URLError

logger = logging.getLogger(__name__)

GITHUB_REPO      = "faceless-truth/mcs-coworker"
CHECK_INTERVAL_H = 6   # hours between background checks

# Only apply commits authored by these emails. If a pushed commit has a
# different author (e.g. a compromised GitHub account pushing via a PR merge),
# refuse the update so we don't roll malicious code out to every accountant
# machine in the next 6h cycle.
ALLOWED_UPDATE_AUTHORS = {"elio@mcands.com.au"}

# Post-update sanity check: if any of these are missing or empty after the
# reset, we assume the update corrupted the tree and roll back to the backup.
CRITICAL_FILES = ("main.py", "api_server.py", "config.py", "plugin_loader.py")

# Backup directory excludes — keep the backup small and skip anything that
# shouldn't be restored (vendored deps, build artefacts, caches).
_BACKUP_IGNORE = shutil.ignore_patterns(
    ".git", "venv", "__pycache__", "node_modules", "*.pyc",
    "dist", "build", "*.log", "*.db", "coworker.db",
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    """Return the root directory of the installed app."""
    return Path(__file__).parent.resolve()


def get_current_version() -> str:
    """Return the current version (short commit hash or fallback)."""
    version_file = _app_dir() / "VERSION"
    if version_file.exists():
        v = version_file.read_text().strip()
        if v:
            return v
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True, text=True, cwd=_app_dir(), timeout=5
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return "unknown"


def _write_version(commit_hash: str) -> None:
    (_app_dir() / "VERSION").write_text(commit_hash.strip())


def _get_remote_commit() -> Optional[str]:
    """Fetch the latest commit SHA on main from GitHub API."""
    try:
        req = Request(
            f"https://api.github.com/repos/{GITHUB_REPO}/commits/main",
            headers={"Accept": "application/vnd.github.v3+json",
                     "User-Agent": "MCS CoWorker-Updater/2.0"},
        )
        with urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
            return data["sha"][:7]
    except Exception as e:
        logger.debug(f"[AutoUpdater] Remote commit fetch failed: {e}")
        return None


def _is_git_repo() -> bool:
    return (_app_dir() / ".git").exists()


def _run(cmd: list, cwd: Path, timeout: int = 120):
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, cwd=cwd, timeout=timeout
        )
        output = (result.stdout + result.stderr).strip()
        return result.returncode == 0, output
    except subprocess.TimeoutExpired:
        return False, "Timed out"
    except Exception as e:
        return False, str(e)


def _python_exe() -> str:
    venv_python = _app_dir() / "venv" / "Scripts" / "python.exe"
    if venv_python.exists():
        return str(venv_python)
    return sys.executable


def _pip_install(cwd: Path):
    req_file = cwd / "requirements.txt"
    if not req_file.exists():
        return True, "No requirements.txt — skipped"
    return _run(
        [_python_exe(), "-m", "pip", "install", "-r", "requirements.txt",
         "--quiet", "--no-warn-script-location"],
        cwd=cwd, timeout=180,
    )


# ── Core update logic ─────────────────────────────────────────────────────────

def _fetch_and_verify(cwd: Path) -> tuple[bool, str]:
    """Fetch from origin/main, verify the latest commit's author, and fast-forward
    via reset --hard.

    Returns (applied, notes). ``applied`` is True only when a new commit was
    actually moved to. Returns (False, reason) when refused, already up to
    date, or on any failure.
    """
    # Step 1 — fetch the latest refs without touching the working tree
    ok, out = _run(["git", "fetch", "origin", "main"], cwd=cwd)
    if not ok:
        return False, f"git fetch failed: {out}"

    # Step 2 — check the author of origin/main's tip before touching anything
    ok, out = _run(
        ["git", "log", "--format=%ae", "origin/main", "-1"], cwd=cwd, timeout=15
    )
    if not ok:
        return False, f"git log failed: {out}"
    author_email = out.strip().lower()
    if author_email not in ALLOWED_UPDATE_AUTHORS:
        return False, (
            f"Update BLOCKED — commit author '{author_email}' not in "
            f"allowed list. Refusing to roll out."
        )

    # Step 3 — compare local HEAD against origin/main; skip if no-op
    ok, local_head = _run(["git", "rev-parse", "HEAD"], cwd=cwd, timeout=10)
    if not ok:
        return False, f"git rev-parse HEAD failed: {local_head}"
    ok, remote_head = _run(
        ["git", "rev-parse", "origin/main"], cwd=cwd, timeout=10
    )
    if not ok:
        return False, f"git rev-parse origin/main failed: {remote_head}"
    local_head = local_head.strip()
    remote_head = remote_head.strip()
    if local_head == remote_head:
        return False, "Already up to date"

    # Step 4 — atomic fast-forward to origin/main
    logger.info(f"[AutoUpdater] Updating {local_head[:8]} -> {remote_head[:8]}")
    ok, out = _run(["git", "reset", "--hard", "origin/main"], cwd=cwd)
    if not ok:
        return False, f"git reset failed: {out}"

    # Step 5 — sanity check the tree before declaring success
    for f in CRITICAL_FILES:
        p = cwd / f
        if not p.exists() or p.stat().st_size == 0:
            return False, (
                f"Post-update sanity check FAILED: {f} missing or empty "
                f"after reset"
            )

    return True, f"Updated {local_head[:8]} -> {remote_head[:8]}"


def check_for_update() -> Optional[dict]:
    """
    Check if a newer commit is available on GitHub main.
    Returns dict with current/latest/update_available, or None on failure.
    """
    current = get_current_version()
    latest  = _get_remote_commit()
    if latest is None:
        return None
    return {
        "current":          current,
        "latest":           latest,
        "update_available": current != latest and current != "unknown",
    }


def apply_update(force: bool = False) -> dict:
    """
    Pull the latest code from GitHub and reinstall dependencies.
    Returns a result dict: success, version_before, version_after, notes.
    """
    cwd            = _app_dir()
    version_before = get_current_version()
    result = {
        "success":        False,
        "version_before": version_before,
        "version_after":  version_before,
        "notes":          "",
    }

    if not _is_git_repo():
        result["notes"] = "Not a git repository — cannot auto-update."
        logger.warning(f"[AutoUpdater] {result['notes']}")
        return result

    if not force:
        info = check_for_update()
        if info is None:
            result["notes"]   = "Could not reach GitHub — skipping update."
            result["success"] = True
            logger.info(f"[AutoUpdater] {result['notes']}")
            return result
        if not info["update_available"]:
            result["notes"]   = f"Already up to date ({version_before})."
            result["success"] = True
            logger.info(f"[AutoUpdater] {result['notes']}")
            return result

    # Backup — copy the full app tree (excluding vendor/build dirs) so the
    # restore path can put the tree back even after a bad `git reset --hard`.
    backup_dir = cwd.parent / f"mcs-coworker-backup-{version_before}"
    try:
        if backup_dir.exists():
            shutil.rmtree(backup_dir)
        shutil.copytree(cwd, backup_dir, ignore=_BACKUP_IGNORE)
        logger.info(f"[AutoUpdater] Backup at {backup_dir}")
    except Exception as e:
        logger.warning(f"[AutoUpdater] Backup failed (continuing): {e}")

    # Fetch, verify author, and fast-forward atomically
    applied, notes = _fetch_and_verify(cwd)
    if not applied:
        # If the sanity check fired, the tree may already be damaged — try to
        # restore before reporting failure.
        if "sanity check FAILED" in notes:
            logger.critical(f"[AutoUpdater] {notes} — restoring backup")
            _restore_backup(cwd, backup_dir)
        logger.info(f"[AutoUpdater] {notes}")
        result["notes"]   = notes
        result["success"] = (notes == "Already up to date")
        return result
    logger.info(f"[AutoUpdater] {notes}")

    # pip install
    ok, out = _pip_install(cwd)
    if not ok:
        logger.warning(f"[AutoUpdater] pip install warning: {out}")
        result["notes"] += f" | pip warning: {out[:200]}"

    new_version = get_current_version()
    _write_version(new_version)

    try:
        shutil.rmtree(backup_dir, ignore_errors=True)
    except Exception:
        pass

    result["success"]       = True
    result["version_after"] = new_version
    result["notes"]         = (
        result["notes"] or f"Updated {version_before} to {new_version}."
    )
    logger.info(f"[AutoUpdater] {result['notes']}")
    return result


def _restore_backup(cwd: Path, backup_dir: Path) -> None:
    """Copy every file from backup_dir back into cwd, overwriting whatever is
    there. Anything under the backup-ignore patterns (venv, .git, caches, DBs)
    is left untouched in cwd since we never backed it up in the first place.
    """
    if not backup_dir.exists():
        return
    try:
        for src in backup_dir.rglob("*"):
            if src.is_dir():
                continue
            rel  = src.relative_to(backup_dir)
            dest = cwd / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
        logger.info("[AutoUpdater] Backup restored.")
    except Exception as e:
        logger.error(f"[AutoUpdater] Restore failed: {e}")


# ── SQLite history ────────────────────────────────────────────────────────────

def _get_db_path() -> Path:
    return _app_dir() / "coworker.db"


def log_update_result(result: dict) -> None:
    try:
        con = sqlite3.connect(_get_db_path())
        con.execute("""
            CREATE TABLE IF NOT EXISTS update_history (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp      TEXT NOT NULL,
                version_before TEXT,
                version_after  TEXT,
                success        INTEGER,
                notes          TEXT
            )
        """)
        con.execute(
            "INSERT INTO update_history "
            "(timestamp, version_before, version_after, success, notes) "
            "VALUES (?,?,?,?,?)",
            (datetime.utcnow().isoformat(),
             result.get("version_before", ""),
             result.get("version_after", ""),
             1 if result.get("success") else 0,
             result.get("notes", "")),
        )
        con.commit()
        con.close()
    except Exception as e:
        logger.debug(f"[AutoUpdater] DB log failed: {e}")


def get_update_history(limit: int = 10) -> list:
    try:
        db = _get_db_path()
        if not db.exists():
            return []
        con  = sqlite3.connect(db)
        rows = con.execute(
            "SELECT timestamp, version_before, version_after, success, notes "
            "FROM update_history ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
        con.close()
        return [{"timestamp": r[0], "version_before": r[1],
                 "version_after": r[2], "success": bool(r[3]),
                 "notes": r[4]} for r in rows]
    except Exception:
        return []


# ── Background checker / singleton ───────────────────────────────────────────

class _AutoUpdater:
    """
    Singleton that checks for updates every CHECK_INTERVAL_H hours
    while the app is running and applies them silently.
    Updates take effect on the next app restart.
    """

    def __init__(self):
        self._running      = False
        self._thread       = None
        self._restart_flag = False

    def start(self, **_kwargs) -> None:
        if not self._running:
            self._running = True
            self._thread  = threading.Thread(
                target=self._loop, daemon=True, name="AutoUpdater")
            self._thread.start()
            logger.info("[AutoUpdater] Background checker started.")

    def stop(self) -> None:
        self._running = False

    def check_now(self) -> Optional[dict]:
        return check_for_update()

    def apply_now(self, force: bool = False) -> dict:
        result = apply_update(force=force)
        log_update_result(result)
        if result["success"] and result["version_before"] != result["version_after"]:
            self._restart_flag = True
        return result

    def needs_restart(self) -> bool:
        return self._restart_flag

    def get_status(self) -> dict:
        return {
            "current_version": get_current_version(),
            "needs_restart":   self._restart_flag,
            "history":         get_update_history(10),
        }

    def set_auto_update(self, enabled: bool) -> None:
        pass   # always enabled in hybrid model

    def _loop(self) -> None:
        time.sleep(60)
        while self._running:
            try:
                result = apply_update(force=False)
                log_update_result(result)
                if (result["success"] and
                        result["version_before"] != result["version_after"]):
                    self._restart_flag = True
                    logger.info("[AutoUpdater] Update applied — restart to activate.")
            except Exception as e:
                logger.warning(f"[AutoUpdater] Loop error: {e}")
            for _ in range(CHECK_INTERVAL_H * 60):
                if not self._running:
                    break
                time.sleep(60)


# Module-level singleton (same interface as before)
AutoUpdater = _AutoUpdater()


# ── Legacy compatibility shims ────────────────────────────────────────────────

def init_updater_db(*args, **kwargs) -> None:
    pass


def get_update_history_legacy(limit: int = 10, db_path=None) -> list:
    return get_update_history(limit)
