# -*- mode: python ; coding: utf-8 -*-
# macOS-specific PyInstaller spec file for Property PDF Builder

import os
from PyInstaller.utils.hooks import collect_data_files

# Collect all data files needed
datas = []

# Include logo.png if it exists
if os.path.exists('logo.png'):
    datas.append(('logo.png', '.'))

# Include sample_images directory if it exists
if os.path.exists('sample_images'):
    datas.append(('sample_images', 'sample_images'))

# Collect reportlab data files
try:
    reportlab_datas = collect_data_files('reportlab')
    datas.extend(reportlab_datas)
except:
    pass

a = Analysis(
    ['pdf_builder_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['reportlab', 'PIL', 'PIL._tkinter_finder', '_tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PropertyPDFBuilder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PropertyPDFBuilder',
)

app = APP(
    coll,
    name='PropertyPDFBuilder',
    icon=None,  # Can add .icns file here if you have one
    version=None,
)

