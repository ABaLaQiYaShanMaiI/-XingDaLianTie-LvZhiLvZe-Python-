# -*- coding: utf-8 -*-
"""临时构建脚本：用 PyInstaller 5.13.2 + Python 3.8.10 构建 Win7~Win11 兼容版 exe。

用法：
    默认构建 32 位版（可运行于 32/64 位 Windows）：
        python _build_win7.py
    构建 64 位版：
        set PYI_BITS=64
        set PYI_EXE_NAME=班组长履职考评表生成工具
        python _build_win7.py

注意：workpath 放在 %TEMP%（避开 OneDrive 目录的文件锁），dist 输出到项目 dist。
"""
import os
import sys
import tempfile

from PyInstaller.__main__ import run

PROJ = os.path.dirname(os.path.abspath(__file__))
SPEC = os.path.join(PROJ, "班组长履职考评表生成工具.spec")

os.environ.setdefault("PYI_BITS", "32")
os.environ.setdefault("PYI_EXE_NAME", "班组长履职考评表生成工具_win32")

exe_name = os.environ["PYI_EXE_NAME"]
workpath = os.environ.get("PYI_WORKPATH") or os.path.join(
    tempfile.gettempdir(), "pyi_work_" + exe_name)

print("PYI_BITS   =", os.environ["PYI_BITS"])
print("PYI_EXE_NAME =", exe_name)
print("workpath   =", workpath)
sys.stdout.flush()

run(["--clean", "-y", "--workpath", workpath, "--distpath",
     os.path.join(PROJ, "dist"), SPEC])

