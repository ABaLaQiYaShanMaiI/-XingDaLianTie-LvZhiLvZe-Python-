# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：安全员履职考评表生成工具（兼容 Windows 7 SP1 ~ Windows 11）

- 运行时：Python 3.8.10（官方确认支持 Win7 的最后版本，同时兼容 Win8.1/10/11）
- 打包器：PyInstaller 5.13.2（支持 Python 3.8，spec 全部参数兼容）
- 目标位数：默认 x64；构建 32 位版本时设置环境变量：
    $env:PYI_BITS=32; $env:PYI_EXE_NAME="安全员履职考评表生成工具_win32"
  32 位版可同时运行在 32 位与 64 位 Windows 上（WoW64）。
- UCRT 已内置进 exe（binaries），未打 KB2999226 补丁的纯净 Win7 SP1 也可直接运行。
"""

import glob
import os

BASE = os.path.dirname(os.path.abspath(SPEC))

BITS = os.environ.get("PYI_BITS", "64")
EXE_NAME = os.environ.get("PYI_EXE_NAME", "安全员履职考评表生成工具")


def _ucrt_binaries():
    """把 Win7 运行时整套文件打进 exe，纯净（未打 KB2999226 补丁）Win7 SP1 也能直接运行：
    - ucrtbase.dll                    —— UCRT 本体
    - downlevel\\api-ms-win-*.dll      —— 全部 API-Set 转发器（api-ms-win-crt-* 与
      api-ms-win-core-*、api-ms-win-security-* 等），否则现代 ucrtbase.dll 在 Win7
      上会报「丢失 api-ms-win-core-sysinfo-l1-2-0.dll」（实测复现）。
    x64 构建取 System32(downlevel)，x86 构建取 SysWOW64(downlevel)。"""
    if BITS == "32":
        sysdir = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "SysWOW64")
    else:
        sysdir = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32")
    items = []
    for pat in ("ucrtbase.dll", os.path.join("downlevel", "api-ms-win-*.dll")):
        for f in glob.glob(os.path.join(sysdir, pat)):
            items.append((f, "."))
    return items


a = Analysis(
    ['generate_kaoping.py'],
    pathex=[BASE],
    binaries=_ucrt_binaries(),
    # 模板为外置文件：不内置进 exe，发布时与 exe 同目录分发
    # （定位逻辑见 kaoping_core.find_template，旧版 exe 内置模板仍兼容）
    datas=[],
    hiddenimports=[
        'tkinterdnd2',
        'win32com',
        'win32com.client',
        'win32com.client.gencache',
        'pythoncom',
        'pywintypes',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'openpyxl',
        'olefile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # upx=False：避免压缩触发杀毒软件误报，旧系统启动也更稳
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
