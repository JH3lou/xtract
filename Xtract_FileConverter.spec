# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app\\app.py'],
    pathex=[],
    binaries=[],
    datas=[('app', 'app'), ('assets', 'assets'), ('templates', 'templates'),('README.md', 'README.md'), ('C:\\\\Users\\\\jhelou\\\\AppData\\\\Local\\\\anaconda3\\\\envs\\\\venv1\\\\Lib\\\\site-packages\\\\customtkinter', 'customtkinter'), ('C:\\\\Users\\\\jhelou\\\\AppData\\\\Local\\\\anaconda3\\\\envs\\\\venv1\\\\Lib\\\\site-packages\\\\openpyxl', 'openpyxl')],
    hiddenimports=['openpyxl', 'openpyxl.cell._writer'],
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
    name='Xtract - File Converter',
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
    icon=['assets\\siteIcon__.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Xtract - File Converter',
)
