# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = [('absolute_v5/ui/assets', 'absolute_v5/ui/assets'), ('C:\\Users\\Rozhman\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tcl8.6', 'tcl/tcl8.6'), ('C:\\Users\\Rozhman\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tk8.6', 'tcl/tk8.6')]
datas += collect_data_files('tkinter')


a = Analysis(
    ['dorabotka.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['tkinter', '_tkinter'],
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
    a.binaries,
    a.datas,
    [],
    name='DataScope',
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
    icon=['absolute_v5\\ui\\assets\\icon_result.ico'],
)
