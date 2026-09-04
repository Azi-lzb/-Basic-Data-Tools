# 基础数据审核工具（Flask / Windows 外壳）

与 `Flask/uos` 共用 `Flask/shared` 的唯一业务核心（`base_audit` 后端 +
`frontend` 前端 + `tests` 测试）。本目录只存放 Windows 外壳与运行数据。

## 目录职责

- `app.py`：Flask 外壳。页面注入 `bridge.api` 桥接层（前端 HTML 与 UOS 版共用，一字不改）；
  继承 `WebApi`，文件选择对话框使用 tkinter，窗口控制按钮无操作。
- `run.py`、`启动Flask-windows.bat`（ASCII + CRLF）：本地服务入口，默认
  `http://127.0.0.1:8750/`。
- `历史审核配置.xlsx`、`data/`：本机运行数据。
- Excel/WPS COM 引擎在 `Flask/shared/src/base_audit/excel_com.py`（懒加载，
  引擎列表按操作系统提供：Windows 显示 Excel/WPS，UOS 显示 LibreOffice Calc）。

## 运行

```bat
启动Flask-windows.bat
```

或 `python run.py`（`--no-browser` 不自动开浏览器，`--port` 改端口）。

## 测试

```sh
cd Flask/shared && PYTHONPATH=src python -m pytest tests/ -q
```

## 与其他发行线的关系

- `Flask/shared` 是唯一活跃开发的业务核心；改功能只改这一处，两个外壳同时生效。
- `external/pywebview2/`（原 `pywebview2/`）是已验证的冻结归档基线，仅作对照，不再开发新功能。
