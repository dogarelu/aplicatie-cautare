# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('default.docx', '.'),    # Include default.docx template (optional, for custom styles)
    ],
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
    name='app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Disable UPX compression (can trigger antivirus false positives)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    onefile=False,  # Use onedir mode (less likely to trigger antivirus)
)

# BUNDLE is macOS-specific, only used on macOS
# On Windows, the EXE above is the final output
# app = BUNDLE(
#     exe,
#     name='app.app',
#     icon=None,
#     bundle_identifier=None,
# )
