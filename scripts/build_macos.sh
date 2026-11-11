#!/bin/zsh
set -euo pipefail
cd "$(dirname "$0")/.."
export PYTHONDONTWRITEBYTECODE=1
BUILD_PYTHON="${WORDSHIELD_PYTHON:-python3}"
"$BUILD_PYTHON" -B -c 'import tkinter, docx, cryptography; import platform; assert platform.system() == "Darwin", "请在 macOS 上构建"'
DEV_DIR="$($BUILD_PYTHON -B -c 'from wordshield.paths import development_directory; print(development_directory())')"
export PYINSTALLER_CONFIG_DIR="$DEV_DIR/pyinstaller-cache"
# 先检查原生扩展没有链接外部动态库，避免打出启动即崩溃的应用。
"$BUILD_PYTHON" -B scripts/check_native_deps.py
"$BUILD_PYTHON" -B -m pytest -p no:cacheprovider
"$BUILD_PYTHON" -B -m PyInstaller --noconfirm --clean --distpath "$DEV_DIR/dist" --workpath "$DEV_DIR/build" WordShield.spec
/usr/bin/codesign --force --deep --sign - "$DEV_DIR/dist/WordShield.app"
/usr/bin/codesign --verify --deep --strict "$DEV_DIR/dist/WordShield.app"
# 运行打包后的程序做导入与加解密自检，不打开界面。
WORDSHIELD_SELF_TEST=1 "$DEV_DIR/dist/WordShield.app/Contents/MacOS/WordShield"
print "已生成 $DEV_DIR/dist/WordShield.app。本机构建为当前架构，分发前请在目标 Mac 验证。"
