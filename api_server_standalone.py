"""
MCS CoWorker — Standalone API Server Entry Point
Used by the Electron wrapper to start the Flask backend without the tkinter UI.
Initialises all modules and starts the Flask server on port 7842.
"""
from __future__ import annotations

import logging
import os
import sys
import time

# Ensure the CoWorker directory is on the path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("coworker")

def main():
    log.info("MCS CoWorker API Server starting...")

    # ── Initialise database ────────────────────────────────────────────────────
    from config import init_db
    init_db()
    log.info("Database initialised")

    # ── Initialise plugin loader ───────────────────────────────────────────────
    from plugin_loader import PluginLoader
    from approval_queue import ApprovalQueue
    from kpi_monitor import KPIMonitor
    import api_server

    loader = PluginLoader()
    loader.load_plugins()
    log.info(f"Loaded {len(loader.get_plugins())} plugins")

    approval_queue = ApprovalQueue()
    kpi_monitor = KPIMonitor()

    # Wire into API server
    api_server.set_loader(loader)
    api_server.set_approval_queue(approval_queue)
    api_server.set_kpi_monitor(kpi_monitor)

    # ── Start plugin scheduler in background ──────────────────────────────────
    try:
        from scheduler import Scheduler
        scheduler = Scheduler(loader)
        scheduler.start()
        log.info("Plugin scheduler started")
    except ImportError:
        log.warning("Scheduler module not found — plugins will not run automatically")

    # ── Start Flask server (blocking) ─────────────────────────────────────────
    log.info("Starting Flask API server on port 7842...")
    api_server.run_server(host="127.0.0.1", port=7842, debug=False)


if __name__ == "__main__":
    main()
