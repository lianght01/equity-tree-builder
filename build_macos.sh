#!/bin/bash
# ============================================================
# 股权树生成器 — macOS 打包脚本
# 运行环境: macOS, 已安装 Python3 + PyInstaller
# 输出: dist/股权树生成器.app
# ============================================================
set -e
cd "$(dirname "$0")"

APP_NAME="股权树生成器"
VERSION="1.0"

echo "=== 股权树生成器 macOS 打包 ==="
echo ""

# 检测架构
ARCH=$(uname -m)
echo "架构: $ARCH"

# 确保模板文件存在
if [ ! -d "templates" ]; then
    echo "错误: templates/ 目录不存在"
    exit 1
fi

# 清理旧构建
rm -rf build dist *.spec

echo "开始打包..."
pyinstaller --onefile --windowed \
    --name "$APP_NAME" \
    --icon icon.icns \
    --add-data "templates:templates" \
    --distpath dist \
    --workpath build \
    --noconfirm \
    equity_tree_gui.py

echo ""
echo "=== 打包完成 ==="
echo "输出: dist/${APP_NAME}.app"

# 显示文件大小
if [ -d "dist/${APP_NAME}.app" ]; then
    SIZE=$(du -sh "dist/${APP_NAME}.app" | cut -f1)
    echo "大小: $SIZE"
    
    # 创建zip压缩包（方便传输）
    cd dist
    zip -r "${APP_NAME}_macOS_${ARCH}.zip" "${APP_NAME}.app" > /dev/null
    echo "压缩包: dist/${APP_NAME}_macOS_${ARCH}.zip"
fi