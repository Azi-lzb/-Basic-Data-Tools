#!/usr/bin/env bash
# 在目标 CPU 架构的统信 UOS 上构建 Flet Linux 发行包。
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="${PYTHON_BIN:-$ROOT_DIR/.venv/bin/python}"

if [ ! -x "$PYTHON_BIN" ]; then
  echo "未找到 .venv。请先执行：./安装依赖Flet-UOS.sh" >&2
  exit 1
fi

cd "$ROOT_DIR"
"$PYTHON_BIN" -m pip install -r requirements.txt -r requirements-build.txt
# flet-cli 由 requirements-build.txt 固定安装。Flet 会打包自己的 Flutter
# 桌面运行时；输出目录不包含模板、历史审核说明或用户 data。
"$ROOT_DIR/.venv/bin/flet" build linux . --yes \
  --output "$ROOT_DIR/dist" \
  --project "base-audit-uos" \
  --product "基础数据审核工具（统信版）" \
  --artifact "基础数据审核工具-统信" \
  --cleanup-app --cleanup-packages

echo "构建完成：$ROOT_DIR/dist"
