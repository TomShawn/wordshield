#!/bin/zsh
set -euo pipefail
cd "$(dirname "$0")/.."
export PYTHONDONTWRITEBYTECODE=1
BUILD_PYTHON="${WORDSHIELD_PYTHON:-python3}"
# 留空构建当前架构；设为 universal2 时同时包含 Intel 与 Apple Silicon。
export WORDSHIELD_TARGET_ARCH="${WORDSHIELD_TARGET_ARCH:-}"
"$BUILD_PYTHON" -B -c 'import tkinter, docx, cryptography; import platform; assert platform.system() == "Darwin", "请在 macOS 上构建"'
DEV_DIR="$($BUILD_PYTHON -B -c 'from wordshield.paths import development_directory; print(development_directory())')"
VERSION="$($BUILD_PYTHON -B -c 'import wordshield; print(wordshield.__version__)')"
ARCH_LABEL="${WORDSHIELD_TARGET_ARCH:-$(uname -m)}"
APP="$DEV_DIR/dist/WordShield.app"
export PYINSTALLER_CONFIG_DIR="$DEV_DIR/pyinstaller-cache"
# 先检查原生扩展没有链接外部动态库，避免打出启动即崩溃的应用。
"$BUILD_PYTHON" -B scripts/check_native_deps.py
"$BUILD_PYTHON" -B -m pytest -p no:cacheprovider
"$BUILD_PYTHON" -B -m PyInstaller --noconfirm --clean --distpath "$DEV_DIR/dist" --workpath "$DEV_DIR/build" WordShield.spec
/usr/bin/codesign --force --deep --sign - "$APP"
/usr/bin/codesign --verify --deep --strict "$APP"
# 运行打包后的程序做导入与内存中的脱敏回填自检，不打开界面。只能覆盖本机架构。
WORDSHIELD_SELF_TEST=1 "$APP/Contents/MacOS/WordShield"
print "已生成 $APP（$VERSION，$(/usr/bin/lipo -archs "$APP/Contents/MacOS/WordShield")）。"

# 只有从干净的工作树构建才生成发布用 zip，保证二进制能对应到某个提交。
if [[ -n "$(git status --porcelain 2>/dev/null)" ]]; then
  print '工作树有未提交改动，跳过发布 zip。提交后重新构建即可生成。' >&2
  exit 0
fi
ARCHIVE="$DEV_DIR/dist/WordShield-$VERSION-macOS-$ARCH_LABEL.zip"
rm -f "$ARCHIVE" "$ARCHIVE.sha256"
# ditto 保留 .app 内的符号链接、扩展属性与签名。
/usr/bin/ditto -c -k --keepParent "$APP" "$ARCHIVE"
(cd "$DEV_DIR/dist" && /usr/bin/shasum -a 256 "${ARCHIVE:t}" > "${ARCHIVE:t}.sha256")
print "发布包 $ARCHIVE"
print "构建提交 $(git rev-parse HEAD)"
cat "$ARCHIVE.sha256"
