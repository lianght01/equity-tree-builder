@echo off
chcp 65001 >nul
title 股权树生成器 Windows 一键构建
echo ===================================================
echo    股权树生成器 — Windows 一键构建
echo ===================================================
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [步骤 1/4] 下载 Python...
    echo 正在从官网下载 Python 3.12（约25MB）...
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe' -OutFile '%TEMP%\python-installer.exe'"
    if %errorlevel% neq 0 (
        echo ❌ 下载失败，请手动安装 Python 3.12 后重试
        echo   下载地址: https://www.python.org/downloads/
        pause
        exit /b 1
    )
    echo 安装 Python（请勾选"Add Python to PATH"）...
    start /wait %TEMP%\python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
    echo ✅ Python 安装完成
) else (
    echo ✅ 已检测到 Python
    python --version
)

echo.
echo [步骤 2/4] 安装依赖包...
python -m pip install --upgrade pip -q
python -m pip install pyinstaller openpyxl -q
echo ✅ 依赖安装完成

echo.
echo [步骤 3/4] 准备模板文件...
if not exist templates (
    echo ⚠ 模板目录不存在，从程序内嵌模板构建...
)

echo.
echo [步骤 4/4] 打包中（约1-3分钟）...
cd /d "%~dp0"
pyinstaller --onefile --windowed ^
    --name "股权树生成器" ^
    --add-data "templates;templates" ^
    --distpath dist ^
    --workpath build ^
    --noconfirm ^
    equity_tree_gui.py

if %errorlevel% equ 0 (
    echo.
    echo ===================================================
    echo ✅ 打包成功!
    echo 输出: %~dp0dist\股权树生成器.exe
    echo.
    REM Create zip
    cd dist
    powershell -Command "Compress-Archive -Path '股权树生成器.exe' -DestinationPath '股权树生成器_Windows.zip' -Force"
    echo 压缩包: %~dp0dist\股权树生成器_Windows.zip
    echo.
    echo 双击 dist\股权树生成器.exe 即可运行
    echo 也可将压缩包发送给其他人，解压即用
    echo ===================================================
) else (
    echo ❌ 打包失败，请截图错误信息反馈
)

pause