@echo off
REM ============================================================
REM 股权树桌面版 — Windows 打包脚本
REM 运行环境: Windows, 已安装 Python + PyInstaller
REM 用法: 双击运行或在 cmd 中执行 build_windows.bat
REM 输出: dist\股权树桌面版.exe
REM ============================================================
cd /d "%~dp0"

set APP_NAME=股权树桌面版

echo === 股权树桌面版 Windows 打包 ===
echo.

REM 检查模板
if not exist templates (
    echo 错误: templates/ 目录不存在
    pause
    exit /b 1
)

REM 清理旧构建
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo 开始打包...
pyinstaller --onefile --windowed ^
    --name "%APP_NAME%" ^
    --add-data "templates;templates" ^
    --distpath dist ^
    --workpath build ^
    --noconfirm ^
    equity_tree_app.py

echo.
echo === 打包完成 ===
echo 输出: dist\%APP_NAME%.exe

REM 创建zip压缩包
cd dist
powershell -Command "Compress-Archive -Path '%APP_NAME%.exe' -DestinationPath '%APP_NAME%_Windows_x64.zip' -Force"
echo 压缩包: dist\%APP_NAME%_Windows_x64.zip

pause
