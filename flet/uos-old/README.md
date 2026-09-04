# 基础数据审核工具（统信 UOS / 麒麟 Flet 版）

这是 Linux 原生发行项目。界面使用 Flet，工作簿读写使用 `openpyxl`；公式重算和条件格式实际填充色读取统一通过目标系统安装的 LibreOffice Calc 完成。

它不包含也不依赖 `pywin32`、Excel/WPS COM、pywebview、formulas、IronCalc 或 excel-formulas-calculator。

## 开发运行

```bash
cd flet/uos
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
python run.py
```

运行前请确保系统已安装 LibreOffice Calc，且 `soffice` 可从命令行调用。执行 `python run.py --check-engine` 可检查计算能力。

## 打包

在 Linux/UOS 环境执行 `./打包Flet-UOS.sh`。发行包只打入 Python/Flet/openpyxl 应用依赖；LibreOffice 作为系统依赖，由安装文档或 DEB 依赖提供。

共享的模板协议、流程配置和业务字段请见根目录 [AGENTS.md](../../AGENTS.md)。
