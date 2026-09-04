# 基础数据审核工具（Flask / 统信 UOS · 麒麟 版）

与 `Flask/windows` 共用 `Flask/shared` 的唯一业务核心（`base_audit` 后端 +
`frontend` 前端 + `tests` 测试）。本目录只存放 UOS 外壳差异。

## 当前可用功能（纯 openpyxl，无办公套件依赖）

- 模板自动识别、逐文件模板归类、数据期识别、机构说明文件清单
- 设置中心：功能、流程编排、主界面显示、历史审核配置路径
- 历史审核记录分页/筛选读取（只读）
- 历史本地校验结果配置读取与汇总三列回写（历史工作簿改名 `历史审核配置.xlsx`）

## 待完成（需要 LibreOffice Calc/UNO 引擎适配）

`base_audit/excel_com.py` 的 `ExcelSession` 是 Windows COM 实现；在 UOS 上
进入需要计算引擎的流程（审核前检查、汇总核查表校验、公式重算、条件格式
渲染）时会**明确失败**并提示适配进行中——不猜测、不静默回退。

移植来源：`flet/uos/src/uos_audit/` 的 `libreoffice.py`、`calculation.py`、
`openpyxl_workbook.py`、`conditional_format.py`。验收以 AGENTS.md 的
UOS 公式计算与条件格式策略为准（INDEX/MATCH/VLOOKUP、错误值、命名区域、
缓存一致性需真实模板基准）。

## 运行

```sh
./启动Flask-UOS.sh
```

或 `python3 run.py`（默认 `http://127.0.0.1:8750/`）。

## 打包

```sh
./打包Flask-UOS.sh
```

产出 `dist/基础数据审核工具_Flask_UOS_<日期>.zip`（无 zip 命令时为 .tar.gz）及
SHA256。包内自带 `shared/`（与 Windows 版同一份核心），`app.py` 自动识别打包
布局。不打包历史审核配置.xlsx（首次运行自动按标准表头创建）和 tests。

## 外壳差异（相对 Windows 外壳）

- 打开文件/目录用 `xdg-open`
- 计算引擎设置只显示 LibreOffice Calc（由 shared 按平台提供）
- 不依赖 pywin32 / pywebview；requirements 只有 flask + openpyxl

## 测试

共享测试位于 `Flask/shared/tests/`，在任一平台运行：

```sh
cd Flask/shared && PYTHONPATH=src python3 -m pytest tests/ -q
```

其中 COM 相关用例在 UOS 上会跳过或按平台断言（见各用例说明）。
