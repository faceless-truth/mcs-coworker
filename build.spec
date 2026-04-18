# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for MC & S CoWorker (pywebview edition)
# Entry point: main.py  (replaces app.py Tkinter shell)
# Frontend:    frontend_dist/  (React build output, copied by build.bat)

import os

block_cipher = None

# ── Locate webview package data ────────────────────────────────────────────────
try:
    import webview
    webview_path = os.path.dirname(webview.__file__)
except ImportError:
    webview_path = None

# ── Data files to bundle ───────────────────────────────────────────────────────
datas = []

# React frontend build — expected at frontend_dist/ next to this spec
local_frontend = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'frontend_dist')
if os.path.exists(local_frontend):
    datas.append((local_frontend, 'frontend_dist'))

# Assets (icon, etc.)
assets_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
if os.path.exists(assets_dir):
    datas.append((assets_dir, 'assets'))

# pywebview assets
if webview_path:
    datas.append((webview_path, 'webview'))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'flask', 'flask_cors', 'werkzeug', 'werkzeug.serving',
        'werkzeug.routing', 'werkzeug.exceptions', 'jinja2', 'click',
        'webview', 'webview.platforms.winforms', 'clr',
        'msal', 'anthropic', 'pytz', 'sqlite3', 'chromadb',
        'config', 'plugin_base', 'plugin_loader', 'graph_client',
        'api_server', 'approval_queue', 'event_bus', 'event_wiring',
        'gateway_client', 'memory_store', 'token_meter', 'kpi_monitor',
        'auto_updater', 'launcher',
        'pdfminer', 'pdfminer.high_level', 'pdfminer.layout', 'pdfminer.pdfpage',
        'pdfminer.pdfinterp', 'pdfminer.converter', 'pdfminer.pdfdocument',
        # Active plugins
        'plugins.plugin_email_reply',
        'plugins.plugin_correspondence_logger', 'plugins.plugin_client_outreach',
        'plugins.plugin_asic_returns', 'plugins.plugin_noa_processor',
        'plugins.plugin_morning_briefing', 'plugins.plugin_wip_summariser',
        'plugins.plugin_debtor_followup', 'plugins.plugin_meeting_prep',
        'plugins.plugin_fusesign_monitor', 'plugins.plugin_engagement_letter',
        'plugins.plugin_bas_reminder', 'plugins.plugin_annual_review',
        'plugins.plugin_tax_return_processor',
        # Retired (kept for reference, not loaded by plugin_loader)
        'plugins.plugin_email_triage', 'plugins.plugin_elio_draft_replies',
        'plugins.plugin_auto_reply_ross',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'customtkinter', 'pystray',
        'matplotlib', 'numpy', 'pandas', 'scipy', 'torch',
        'tensorflow', 'jupyter', 'notebook', 'IPython', 'pytest', 'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

icon_path = os.path.join('assets', 'icon.ico')
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EVA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=icon_path if os.path.exists(icon_path) else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EVA',
)
