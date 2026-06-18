# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files

datas = [("templates", "templates")]

hiddenimports = [
    "pandas.io.excel._openpyxl",
    "pandas.io.excel._xlrd",
    "packaging",
    "packaging.version",
    "packaging.specifiers",
]

# ── 收集 Python 运行时 DLL ───────────────────────────
# python39.dll 及其依赖的 VC++ 运行时必须打进文件夹
python_dir = os.path.dirname(sys.executable)
binaries = []
for dll_name in ["python39.dll", "vcruntime140.dll", "vcruntime140_1.dll"]:
    src = os.path.join(python_dir, dll_name)
    if os.path.exists(src):
        binaries.append((src, "."))

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter", "tcl", "Tkinter",
        "unittest", "pytest",
        "setuptools", "pip", "wheel", "distutils",
        "IPython", "jupyter", "notebook",
        "sphinx", "docutils",
        "Cython", "scipy",
        "PIL", "Pillow",
        "curses", "sqlite3",
        "matplotlib", "networkx",
        "psycopg2", "sqlalchemy",
        "MySQLdb", "pymysql", "cx_Oracle",
        "libpq",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# 过滤无用 DLL
cleaned_binaries = [
    t for t in a.binaries
    if "libpq" not in t[0].lower()
    and "libcrypto" not in os.path.basename(t[0]).lower()
    and "libssl" not in os.path.basename(t[0]).lower()
    and "ssleay" not in os.path.basename(t[0]).lower()
]

exe = EXE(
    pyz,
    a.scripts,
    cleaned_binaries,
    a.datas,
    [],
    name="Datafloat数据处理平台",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    cleaned_binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Datafloat数据处理平台",
)
