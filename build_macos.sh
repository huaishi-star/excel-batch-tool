#!/usr/bin/env bash
# Lsy的Excel-Tool — macOS 本地打包脚本
# 用法：chmod +x build_macos.sh && ./build_macos.sh
# 产物：dist/Lsy的Excel-Tool.app 与 dist/Lsy的Excel-Tool-macos-<arch>.zip
set -euo pipefail
cd "$(dirname "$0")"

PYTHON="${PYTHON:-python3}"

echo "=== Lsy的Excel-Tool macOS 打包 ==="
"$PYTHON" --version
echo "架构：$(uname -m)"

# 使用独立虚拟环境，避免污染系统 Python
if [ ! -d .venv ]; then
    "$PYTHON" -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller==6.22.0

echo "[1/3] 运行核心测试..."
python test_make_samples.py
python test_core.py

echo "[2/3] 打包 .app..."
python -m PyInstaller --noconfirm --clean --windowed \
    --name "Lsy的Excel-Tool" \
    --osx-bundle-identifier "com.lsy.excel-tool" \
    --exclude-module unittest --exclude-module pydoc_data --exclude-module test \
    --exclude-module numpy --exclude-module PIL --exclude-module lxml \
    --exclude-module scipy --exclude-module pandas --exclude-module matplotlib \
    --exclude-module sklearn app.py

APP="dist/Lsy的Excel-Tool.app"
if [ ! -d "$APP" ]; then
    echo "未找到 $APP，打包失败" >&2
    exit 1
fi

echo "[3/3] 压缩为 zip..."
ARCH="$(uname -m)"
ZIP="dist/Lsy的Excel-Tool-macos-${ARCH}.zip"
rm -f "$ZIP"
# ditto 能保留 .app 的权限、符号链接和资源分叉
ditto -c -k --keepParent "$APP" "$ZIP"

echo
echo "完成：$ZIP"
echo "首次打开若被 Gatekeeper 拦截，可右键 → 打开，或执行："
echo "  xattr -dr com.apple.quarantine \"$APP\""
