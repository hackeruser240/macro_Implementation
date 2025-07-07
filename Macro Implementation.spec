# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['CodeGUI.py'],
    pathex=[],
    binaries=[('D:\\ProgramFiles\\Anaconda\\pkgs\\tk-8.6.14-h0416ee5_0\\Library\\bin\\tcl86t.dll', '.'), ('D:\\ProgramFiles\\Anaconda\\pkgs\\tk-8.6.14-h0416ee5_0\\Library\\bin\\tk86t.dll', '.')],
    datas=[('TXT Files\\\\Macro_1.txt', 'TXT Files')],
    hiddenimports=['_tkinter'],
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
    name='Macro Implementation',
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
)
