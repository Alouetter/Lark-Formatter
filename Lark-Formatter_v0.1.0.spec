# -*- mode: python ; coding: utf-8 -*-
# Legacy release spec kept for local reference.
# Current release packaging uses: Lark-Formatter_v0.20_LTS.spec

datas = [
    ("src/scene/presets", "templates"),
    ("src/ui/themes", "src/ui/themes"),
    ("src/ui/icons", "src/ui/icons"),
]

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
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
    [],
    exclude_binaries=True,
    name="Lark-Formatter",
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
    icon="src\\ui\\icons\\app_icon.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Lark-Formatter_v0.2.0",
)
