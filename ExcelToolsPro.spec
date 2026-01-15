# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for ExcelToolsPro - Production Build
Version: 1.0.0
Build without console, optimized for end users
"""

import sys
from pathlib import Path

block_cipher = None

# Chemins
ROOT_DIR = Path(SPECPATH)
SRC_DIR = ROOT_DIR / 'src'
ICO_DIR = ROOT_DIR / 'ico'

# Données additionnelles à inclure
datas = [
    (str(ICO_DIR), 'ico'),
]

# Hidden imports pour customtkinter et dépendances
hiddenimports = [
    'customtkinter',
    'PIL._tkinter_finder',
    'openpyxl',
    'pandas',
    'numpy',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
]

a = Analysis(
    ['run.py'],
    pathex=[str(ROOT_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelToolsPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Pas de console pour la production
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICO_DIR / 'icone.ico'),
    version='version_info.txt' if (ROOT_DIR / 'version_info.txt').exists() else None,
)
