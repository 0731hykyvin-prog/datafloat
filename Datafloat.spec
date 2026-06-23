# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [("templates", "templates")]

hiddenimports = [
    "pandas.io.excel._openpyxl",
    "pandas.io.excel._xlrd",
    "packaging",
    "packaging.version",
    "packaging.specifiers",
    # numpy 内部模块（pandas 依赖）
    "numpy",
    "numpy._core",
    "numpy.core",
    "numpy.core._dtype_ctypes",
    "numpy.core._internal",
    "numpy.core.multiarray",
    "numpy.core.numeric",
    "numpy.core.umath",
    "numpy.lib",
    "numpy.lib.format",
    "numpy.random",
    "numpy.random._common",
    "numpy.random._generator",
    "numpy.random._mt19937",
    "numpy.random._pcg64",
    "numpy.random._philox",
    "numpy.random._sfc64",
    "numpy.random.bit_generator",
    "numpy.random.mtrand",
    "numpy.linalg",
    "numpy.linalg.lapack_lite",
    # Python 标准库可能缺失的
    "secrets",
    "hashlib",
    "hmac",
    "binascii",
]

# ── 收集 Python 运行时 DLL ───────────────────────────
# python38.dll + VC++ 运行时 + UCRT（Win7 需要）
python_dir = os.path.dirname(sys.executable)
binaries = []

# 核心 VC++ 运行时（Python 安装目录自带）
for dll_name in ["python38.dll", "vcruntime140.dll", "vcruntime140_1.dll"]:
    src = os.path.join(python_dir, dll_name)
    if os.path.exists(src):
        binaries.append((src, "."))

# UCRT（Windows 7 默认没有，从 System32 收集）
system32 = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "System32")
ucrt_dlls = [
    "ucrtbase.dll",
    "api-ms-win-core-console-l1-1-0.dll",
    "api-ms-win-core-console-l1-2-0.dll",
    "api-ms-win-core-datetime-l1-1-0.dll",
    "api-ms-win-core-debug-l1-1-0.dll",
    "api-ms-win-core-errorhandling-l1-1-0.dll",
    "api-ms-win-core-file-l1-1-0.dll",
    "api-ms-win-core-file-l1-2-0.dll",
    "api-ms-win-core-file-l2-1-0.dll",
    "api-ms-win-core-handle-l1-1-0.dll",
    "api-ms-win-core-heap-l1-1-0.dll",
    "api-ms-win-core-interlocked-l1-1-0.dll",
    "api-ms-win-core-libraryloader-l1-1-0.dll",
    "api-ms-win-core-localization-l1-2-0.dll",
    "api-ms-win-core-memory-l1-1-0.dll",
    "api-ms-win-core-namedpipe-l1-1-0.dll",
    "api-ms-win-core-processenvironment-l1-1-0.dll",
    "api-ms-win-core-processthreads-l1-1-0.dll",
    "api-ms-win-core-processthreads-l1-1-1.dll",
    "api-ms-win-core-profile-l1-1-0.dll",
    "api-ms-win-core-rtlsupport-l1-1-0.dll",
    "api-ms-win-core-string-l1-1-0.dll",
    "api-ms-win-core-synch-l1-1-0.dll",
    "api-ms-win-core-synch-l1-2-0.dll",
    "api-ms-win-core-sysinfo-l1-1-0.dll",
    "api-ms-win-core-timezone-l1-1-0.dll",
    "api-ms-win-core-util-l1-1-0.dll",
    "api-ms-win-crt-conio-l1-1-0.dll",
    "api-ms-win-crt-convert-l1-1-0.dll",
    "api-ms-win-crt-environment-l1-1-0.dll",
    "api-ms-win-crt-filesystem-l1-1-0.dll",
    "api-ms-win-crt-heap-l1-1-0.dll",
    "api-ms-win-crt-locale-l1-1-0.dll",
    "api-ms-win-crt-math-l1-1-0.dll",
    "api-ms-win-crt-multibyte-l1-1-0.dll",
    "api-ms-win-crt-private-l1-1-0.dll",
    "api-ms-win-crt-process-l1-1-0.dll",
    "api-ms-win-crt-runtime-l1-1-0.dll",
    "api-ms-win-crt-stdio-l1-1-0.dll",
    "api-ms-win-crt-string-l1-1-0.dll",
    "api-ms-win-crt-time-l1-1-0.dll",
    "api-ms-win-crt-utility-l1-1-0.dll",
]
for dll_name in ucrt_dlls:
    src = os.path.join(system32, dll_name)
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
