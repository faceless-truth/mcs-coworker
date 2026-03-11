# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for MC & S Desktop Agent

import os
import glob
import customtkinter

block_cipher = None

# Locate CustomTkinter package data (themes, assets)
ctk_path = os.path.dirname(customtkinter.__file__)

# Collect all plugin files
plugin_files = []
plugins_dir = os.path.join(os.getcwd(), 'plugins')
for f in glob.glob(os.path.join(plugins_dir, '*.py')):
    plugin_files.append((f, 'plugins'))

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        (ctk_path, 'customtkinter'),
        (os.path.join('plugins', '__init__.py'), 'plugins'),
        (os.path.join('assets', '*'), 'assets'),
    ] + plugin_files,
    hiddenimports=[
        'customtkinter',
        'msal',
        'anthropic',
        'pytz',
        'sqlite3',
        'pystray',
        'PIL',
        'PIL._tkinter_finder',
        'plugin_base',
        'plugin_loader',
        'config',
        'graph_client',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'torch',
        'tensorflow',
        'jupyter',
        'notebook',
        'IPython',
        'pytest',
        'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='MCS Desktop Agent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=os.path.join('assets', 'icon.ico'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MCS Desktop Agent',
)
