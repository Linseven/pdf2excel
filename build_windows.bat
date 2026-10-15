@echo off
rem =====================================================================
rem PDF2Excel · Windows 一键构建脚本
rem 产出：dist\PDF2Excel.exe（单文件，双击即可运行）
rem
rem 前置条件：已安装 Python 3.10+ 并勾选 "Add python.exe to PATH"
rem 用法：双击本文件，或在命令行执行 build_windows.bat
rem =====================================================================
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 Python。请先安装 Python 3.10+ 并勾选 "Add python.exe to PATH"：
    echo        https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ==^> 使用 Python：
python --version

if not exist .venv (
    echo ==^> 创建虚拟环境 .venv
    python -m venv .venv
    if errorlevel 1 (
        echo [错误] 创建虚拟环境失败
        pause
        exit /b 1
    )
)

echo ==^> 安装依赖（首次较慢，请耐心等待）
.venv\Scripts\python.exe -m pip install --quiet --upgrade pip
.venv\Scripts\python.exe -m pip install --quiet -r requirements-dev.txt
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络
    pause
    exit /b 1
)

if not exist assets\icon.ico (
    echo ==^> 生成应用图标
    .venv\Scripts\python.exe tools\make_icons.py
)

echo ==^> PyInstaller 打包中（约 1-3 分钟）…
.venv\Scripts\pyinstaller.exe ^
    --noconfirm --clean --noconsole --onefile ^
    --name "PDF2Excel" ^
    --collect-all pywebview ^
    --add-data "ui;ui" ^
    --add-data "assets;assets" ^
    --icon assets\icon.ico ^
    main.py
if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo 构建完成：%cd%\dist\PDF2Excel.exe
echo 提示：Windows 10/11 自带 WebView2 运行时；如遇白屏请安装：
echo        https://developer.microsoft.com/microsoft-edge/webview2/
pause
