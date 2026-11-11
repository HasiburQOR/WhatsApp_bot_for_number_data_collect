# -*- mode: python ; coding: utf-8 -*-
#
# Build (from this folder):
#   build.bat   (or: pyinstaller --clean --noconfirm WhatsApp_Number_Checker_Web.spec)
#
# NOTE: this spec was written during recovery (the original spec was lost). It
# produces a single console exe. Chromium is NOT bundled inside the exe; the
# portable edition keeps it in a "Browsers" folder next to the exe, which
# launcher_web.ensure_chromium() finds automatically.

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = ['checker_web']
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('qrcode')
hiddenimports += collect_submodules('playwright')

# Playwright ships its Node driver inside the package; it must travel with the exe.
datas = collect_data_files('playwright')

a = Analysis(
    ['launcher_web.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'pytest', 'numpy', 'PIL'],
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
    name='WhatsApp_Number_Checker_Web',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,   # must stay True: the program is console driven
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
