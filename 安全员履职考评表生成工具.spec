# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：安全员履职考评表生成工具"""

import os

BASE = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['generate_kaoping.py'],
    pathex=[BASE],
    binaries=[],
    datas=[
        ('kaoping_core.py', '.'),
        (os.path.join(BASE, '安全员安全生产责任制履职清单考评表（模板）.doc'), '.'),
    ],
    hiddenimports=[
        'tkinterdnd2',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'openpyxl',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
    ],
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
    name='安全员履职考评表生成工具',
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
