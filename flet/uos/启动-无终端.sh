#!/usr/bin/env bash
# 供 .desktop 双击启动使用（无终端窗口）。
# 等价的终端命令：./启动Flet-UOS.sh  或  .venv/bin/python run.py --flet
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="$ROOT_DIR/.venv/bin/python"

if [ ! -x "$PYTHON_BIN" ]; then
  zenity --error --text "未找到 .venv，请先执行：./安装依赖Flet-UOS.sh" 2>/dev/null \
    || echo "未找到 .venv，请先执行：./安装依赖Flet-UOS.sh" >&2
  exit 1
fi

cd "$ROOT_DIR"
exec "$PYTHON_BIN" run.py --flet
