@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo  Datafloat - 下载离线依赖包
echo ============================================
echo.

:: ── 检查 Python ──────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

python --version

:: ── 创建 wheels 目录 ─────────────────────────
if not exist "wheels" mkdir wheels

:: ── 升级 pip ─────────────────────────────────
echo.
echo [1/3] 升级 pip...
python -m pip install --upgrade pip

:: ── 下载所有依赖到 wheels ────────────────────
echo.
echo [2/3] 下载依赖包到 wheels 目录...
python -m pip download ^
    -r requirements.txt ^
    -d wheels

if %errorlevel% neq 0 (
    echo.
    echo [错误] 下载失败，请检查网络连接后重试。
    pause
    exit /b 1
)

:: ── 下载 PyInstaller ─────────────────────────
echo.
echo [3/3] 下载 PyInstaller...
python -m pip download pyinstaller -d wheels

:: ── 完成 ─────────────────────────────────────
echo.
echo ============================================
echo  下载完成！
echo  依赖包已保存到: wheels\
echo.
echo  将整个项目目录（包含 wheels 文件夹）
echo  拷贝到内网电脑，然后双击运行:
echo    _build_offline_exe.cmd
echo ============================================
pause
