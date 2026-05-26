# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['fai_dimension_numberer.py'],
    pathex=[],
    binaries=[],
    datas=[('dimension_detector.py', '.')],
    hiddenimports=[],
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
    name='FAI_Dimension_Numberer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,   # UPX compression is a common malware technique; disabling it
    upx_exclude=[],  # reduces false-positive AV flags on the built exe.
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
