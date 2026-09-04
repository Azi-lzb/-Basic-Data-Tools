#!/bin/sh
# Flask 审核工具 UOS/麒麟 启动脚本
cd "$(dirname "$0")" || exit 1

PY=""
if command -v python3 >/dev/null 2>&1; then
    PY=python3
elif command -v python >/dev/null 2>&1; then
    PY=python
else
    echo "[错误] 未找到 python3，请先安装 Python 3 并加入 PATH。"
    exit 1
fi

if ! "$PY" -c "import flask" >/dev/null 2>&1; then
    echo "[错误] 缺少 flask，请先执行：pip3 install -r requirements.txt"
    exit 1
fi

exec "$PY" run.py "$@"
