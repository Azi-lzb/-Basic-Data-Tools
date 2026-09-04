# 基础数据审核工具（Windows Flet 版）

这是 Windows 的 Flet 发行项目。界面使用 Flet，工作簿读写使用 `openpyxl`；公式重算和条件格式实际填充色读取通过 Microsoft Excel COM 或 WPS 表格 COM 完成。

它不包含 LibreOffice、formulas、IronCalc、excel-formulas-calculator 及其研究型依赖；目标电脑需自行安装 Microsoft Excel 或 Windows 版 WPS 表格。

## 开发运行

```bat
cd flet\windows
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python run.py
```

也可双击 `启动Flet-windows.bat`。执行 `python run.py --check-engine` 可检查 Excel/WPS COM 是否可用。

## 打包

使用 `打包Flet-windows.bat` 构建 Windows Flet 成品。打包前会安装本目录 `requirements-windows.txt` 中的最小依赖。

共享的模板协议、流程配置和业务字段请见根目录 [AGENTS.md](../../AGENTS.md)。
