# -*- mode: python ; coding: utf-8 -*-

import os

from PyInstaller.utils.hooks import collect_data_files

# ── 数据文件 ──────────────────────────────────────────
datas = [
    ("templates", "templates"),
]
datas += collect_data_files("matplotlib", include_py_files=False)

# ── 隐藏导入 ─────────────────────────────────────────
hiddenimports = [
    "pandas.io.excel._openpyxl",
    "pandas.io.excel._xlrd",
    "matplotlib.backends.backend_qt5agg",
    "matplotlib.figure",
    "matplotlib.pyplot",
    "packaging",
    "packaging.version",
    "packaging.specifiers",
]

# ── 分析 ─────────────────────────────────────────────
a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
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
        "psycopg2", "sqlalchemy",
        "MySQLdb", "pymysql", "cx_Oracle",
        "libpq",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# ── 过滤无关 DLL ─────────────────────────────────────
cleaned_binaries = [
    t for t in a.binaries
    if "libpq" not in t[0].lower()
    and "libcrypto" not in os.path.basename(t[0]).lower()
    and "libssl" not in os.path.basename(t[0]).lower()
    and "ssleay" not in os.path.basename(t[0]).lower()
]

# ── 主程序 EXE ───────────────────────────────────────
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

# ── 收集为文件夹（onedir）—— 避免单文件解压崩溃 ──────
coll = COLLECT(
    exe,
    cleaned_binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Datafloat数据处理平台",
)
