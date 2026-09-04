#!/usr/bin/env bash
# 启动统信原生 Flet 工作台。
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="$ROOT_DIR/.venv/bin/python"

if [ ! -x "$PYTHON_BIN" ]; then
  echo "未找到 .venv。请先执行：./安装依赖Flet-UOS.sh" >&2
  exit 1
fi

cd "$ROOT_DIR"
exec "$PYTHON_BIN" run.py --flet
