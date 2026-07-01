# -*- mode: python ; coding: utf-8 -*-

import re
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files

source_text = Path('gl_input_copy_gui.py').read_text(encoding='utf-8')
app_version = re.search(r'^APP_VERSION = "([^"]+)"', source_text, re.MULTILINE).group(1)
exe_name = f'GLInputCopyTool-v{app_version}'
tkinterdnd2_datas = collect_data_files('tkinterdnd2')

a = Analysis(
    ['gl_input_copy_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets')] + tkinterdnd2_datas,
    hiddenimports=['tkinterdnd2'],
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
    name=exe_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/app.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GLInputCopyTool',
)
