#!/bin/zsh
# 统一开发入口。所有自动生成的文件都写到源码目录之外。
set -euo pipefail
PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$PROJECT_DIR"
export PYTHONDONTWRITEBYTECODE=1
BASE_PYTHON="${WORDSHIELD_PYTHON:-python3}"
DEV_DIR="$($BASE_PYTHON -B -c 'from wordshield.paths import development_directory; print(development_directory())')"
VENV_PYTHON="$DEV_DIR/venv/bin/python"
ACTION="${1:-run}"

if [[ "$ACTION" == "paths" ]]; then
  print "源码目录 $PROJECT_DIR"
  print "开发环境与构建目录 $DEV_DIR"
  exit 0
fi
if [[ "$ACTION" == "setup" ]]; then
  "$BASE_PYTHON" -B -c 'import tkinter; print("Tkinter 已就绪")'
  mkdir -p "$DEV_DIR"
  "$BASE_PYTHON" -B -m venv "$DEV_DIR/venv"
  DEPENDENCIES=("${(@f)$($BASE_PYTHON -B -c 'import tomllib; from pathlib import Path; p=tomllib.loads(Path("pyproject.toml").read_text())["project"]; print("\n".join(p["dependencies"] + p["optional-dependencies"]["dev"]))')}")
  # 优先使用官方二进制 wheel。源码编译的扩展可能链接 Homebrew 动态库，导致打包后的应用无法启动。
  "$VENV_PYTHON" -B -m pip install --prefer-binary "${DEPENDENCIES[@]}"
  print "开发依赖已安装到 $DEV_DIR/venv"
  exit 0
fi
if [[ ! -x "$VENV_PYTHON" ]]; then
  print '请先执行 scripts/dev.sh setup 安装外置开发环境。' >&2
  exit 1
fi
case "$ACTION" in
  run) exec "$VENV_PYTHON" -B -m wordshield ;;
  test) exec "$VENV_PYTHON" -B -m pytest -p no:cacheprovider ;;
  fixtures) exec "$VENV_PYTHON" -B scripts/generate_fixtures.py ;;
  build) WORDSHIELD_PYTHON="$VENV_PYTHON" exec scripts/build_macos.sh ;;
  *) print '用法 scripts/dev.sh [setup|run|test|fixtures|build|paths]' >&2; exit 2 ;;
esac
