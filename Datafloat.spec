# -*- mode: python ; coding: utf-8 -*-

import os

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ── 数据文件 ──────────────────────────────────────────
datas = [
    ("templates", "templates"),
]
# matplotlib 字体、配置文件
datas += collect_data_files("matplotlib", include_py_files=False)

# ── 隐藏导入（只列出 PyInstaller 自动检测可能遗漏的）──
hiddenimports = [
    # pandas I/O（Excel 引擎入口）
    "pandas.io.excel._openpyxl",
    "pandas.io.excel._xlrd",
    # matplotlib 可视化后端
    "matplotlib.backends.backend_qt5agg",
    "matplotlib.figure",
    "matplotlib.pyplot",
    # 包版本检测（matplotlib 内部依赖）
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
        # 排除用不到的大型库，缩小 exe 体积
        "tkinter",
        "tcl",
        "Tkinter",
        "unittest",
        "pytest",
        "setuptools",
        "pip",
        "wheel",
        "distutils",
        "IPython",
        "jupyter",
        "notebook",
        "sphinx",
        "docutils",
        "Cython",
        "scipy",
        "PIL",
        "Pillow",
        "curses",
        "sqlite3",
        # 数据库（Datafloat 不用）
        "psycopg2",
        "sqlalchemy",
        "MySQLdb",
        "pymysql",
        "cx_Oracle",
        "libpq",
        # 注意：不要排除 test，pandas 运行可能需要某些内部模块
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# ── 过滤无关二进制 ────────────────────────────────────
cleaned_binaries = [
    (src, dst) for (src, dst) in a.binaries
    if "libpq" not in src.lower()
    and "libcrypto" not in os.path.basename(src).lower()
    and "libssl" not in os.path.basename(src).lower()
    and "ssleay" not in os.path.basename(src).lower()
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
