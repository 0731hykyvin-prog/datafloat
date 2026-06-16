# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ── 数据文件 ──────────────────────────────────────────
datas = [
    ("templates", "templates"),
]
# matplotlib 字体、配置文件等
datas += collect_data_files("matplotlib", include_py_files=False)

# ── 隐藏导入（PyInstaller 自动检测不到的模块）─────────
hiddenimports = []

# pandas
hiddenimports += collect_submodules("pandas")
hiddenimports += collect_submodules("pandas._libs")
hiddenimports += collect_submodules("pandas.io")

# openpyxl（Excel 读写引擎）
hiddenimports += collect_submodules("openpyxl")

# xlrd（老格式 .xls）
hiddenimports += collect_submodules("xlrd")

# networkx（关系图谱）
hiddenimports += collect_submodules("networkx")
hiddenimports += collect_submodules("networkx.algorithms")
hiddenimports += collect_submodules("networkx.drawing")
hiddenimports += collect_submodules("networkx.generators")
hiddenimports += collect_submodules("networkx.readwrite")

# matplotlib 后端（PySide6 使用 QtAgg）
hiddenimports += collect_submodules("matplotlib.backends")
hiddenimports += [
    "matplotlib.backends.backend_qtagg",
    "matplotlib.backends.backend_qt5agg",
    "matplotlib.figure",
    "matplotlib.pyplot",
]
hiddenimports += collect_submodules("matplotlib")

# PySide6
hiddenimports += collect_submodules("PySide6")

# 其他可能缺失的依赖
hiddenimports += [
    "packaging",
    "packaging.version",
    "packaging.specifiers",
    "packaging.requirements",
    "packaging.markers",
    "packaging.tags",
    "packaging.utils",
    "xml",
    "xml.etree",
    "xml.etree.ElementTree",
    "lxml",
    "html",
    "html.parser",
    "json",
    "csv",
    "io",
    "os",
    "sys",
    "re",
    "datetime",
    "collections",
    "warnings",
    "tempfile",
    "shutil",
    "pathlib",
    "ctypes",
    "numbers",
    "decimal",
    "zoneinfo",
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
        # 排除不需要的大型库，减小体积
        "tkinter",
        "tcl",
        "Tkinter",
        "test",
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
        "numpy.tests",
        "scipy",
        "PIL",
        "Pillow",
        "curses",
        "sqlite3",
        "ssl",
    ],
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
    name="Datafloat数据处理平台",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # 关闭 UPX 避免找不到 upx 导致打包失败
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # 发布时不显示控制台窗口；调试时可改为 True
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,           # 可在此指定 .ico 图标路径
)
