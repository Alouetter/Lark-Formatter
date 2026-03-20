# -*- mode: python ; coding: utf-8 -*-
# Optional one-file packaging spec kept for manual/local use.
# Current standard release packaging uses: Lark-Formatter_v0.20_LTS.spec


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('src/scene/presets', 'templates'), ('src/ui/themes', 'src/ui/themes'), ('src/ui/icons', 'src/ui/icons')],
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
    name='Lark-Formatter-onefile',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['src\\ui\\icons\\app_icon.ico'],
)
