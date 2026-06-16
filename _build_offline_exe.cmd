@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo  Datafloat - 离线打包 EXE
echo ============================================
echo.

:: ── 定位 Python ──────────────────────────────
set PYTHON_CMD=

:: 优先使用便携版 Python
if exist "portable_python\python.exe" (
    set "PYTHON_CMD=portable_python\python.exe"
    echo [信息] 使用便携版 Python: portable_python\
    goto :found_python
)

:: 其次使用系统 Python
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python"
    echo [信息] 使用系统 Python
    goto :found_python
)

echo.
echo [错误] 未找到 Python！
echo.
echo 方案一（推荐）：在项目目录下放置便携版 Python
echo   创建 portable_python 文件夹，放入 python.exe 及完整运行环境
echo   参考: 便携Python放置说明.txt
echo.
echo 方案二：安装系统 Python
echo   下载地址: https://www.python.org/downloads/
echo   安装时勾选 "Add Python to PATH"
pause
exit /b 1

:found_python
%PYTHON_CMD% --version
echo.

:: ── 检查 wheels 目录 ─────────────────────────
if not exist "wheels" (
    echo [错误] 未找到 wheels 目录
    echo 请先在联网电脑上运行: _download_dependencies.cmd
    echo 然后将包含 wheels 的整个项目目录拷贝到此电脑
    pause
    exit /b 1
)

:: ── 离线安装依赖 ─────────────────────────────
echo [1/3] 从 wheels 离线安装依赖...
%PYTHON_CMD% -m pip install ^
    --no-index ^
    --find-links=wheels ^
    -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo [错误] 依赖安装失败
    echo 请检查:
    echo   1. wheels 目录中的依赖包是否完整
    echo   2. Python 版本是否与 wheels 包匹配（如 cp311 需 Python 3.11）
    pause
    exit /b 1
)

:: ── 安装 PyInstaller ─────────────────────────
echo.
echo [2/3] 安装 PyInstaller...
%PYTHON_CMD% -m pip install --no-index --find-links=wheels pyinstaller

if %errorlevel% neq 0 (
    echo [警告] 离线安装 PyInstaller 失败，尝试联网安装...
    %PYTHON_CMD% -m pip install pyinstaller
)

:: ── 打包 ──────────────────────────────────────
echo.
echo [3/3] 开始打包...
echo 这可能需要几分钟，请耐心等待...
echo.

%PYTHON_CMD% -m PyInstaller Datafloat.spec

if %errorlevel% neq 0 (
    echo.
    echo ============================================
    echo  打包失败！
    echo.
    echo  排查建议:
    echo   1. 检查上方错误信息
    echo   2. 将 Datafloat.spec 中 console=False 改为 True
    echo      然后重新运行此脚本，查看详细错误
    echo ============================================
    pause
    exit /b 1
)

:: ── 完成 ──────────────────────────────────────
echo.
echo ============================================
echo  打包成功！
echo.
echo  输出位置: dist\Datafloat数据处理平台.exe
echo.
echo  将该 exe 拷贝到任意 Windows 电脑即可运行
echo  无需安装 Python 或其他依赖！
echo ============================================
pause
