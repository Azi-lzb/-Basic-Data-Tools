#!/usr/bin/env bash
# 在统信 UOS 中执行：创建本项目专用虚拟环境并安装 Python 依赖。
# LibreOffice、python3-uno 由系统软件中心或 apt 安装，不在本脚本中提权安装。
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

# system-site-packages 让统信通过 apt 安装的 python3-uno 能被本项目使用；
# 它只在执行“条件格式结果提取”时才是必需的。
python3 -m venv --system-site-packages .venv
"$ROOT_DIR/.venv/bin/python" -m pip install --upgrade pip
"$ROOT_DIR/.venv/bin/python" -m pip install -r requirements.txt

echo
echo "Python 依赖已安装。"
echo "如需复杂公式或条件格式提取，请确认系统已安装 libreoffice-calc 和 python3-uno。"
echo "启动命令：./启动Flet-UOS.sh"
