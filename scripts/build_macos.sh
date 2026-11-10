#!/bin/zsh
set -euo pipefail
cd "$(dirname "$0")/.."
export PYTHONDONTWRITEBYTECODE=1
BUILD_PYTHON="${WORDSHIELD_PYTHON:-python3}"
"$BUILD_PYTHON" -B -c 'import tkinter, docx, cryptography; import platform; assert platform.system() == "Darwin", "请在 macOS 上构建"'
DEV_DIR="$($BUILD_PYTHON -B -c 'from wordshield.paths import development_directory; print(development_directory())')"
export PYINSTALLER_CONFIG_DIR="$DEV_DIR/pyinstaller-cache"
"$BUILD_PYTHON" -B -m pytest -p no:cacheprovider
"$BUILD_PYTHON" -B -m PyInstaller --noconfirm --clean --distpath "$DEV_DIR/dist" --workpath "$DEV_DIR/build" WordShield.spec
/usr/bin/codesign --force --deep --sign - "$DEV_DIR/dist/WordShield.app"
/usr/bin/codesign --verify --deep --strict "$DEV_DIR/dist/WordShield.app"
print "已生成 $DEV_DIR/dist/WordShield.app。本机构建为当前架构，分发前请在目标 Mac 验证。"
