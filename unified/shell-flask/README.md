# 基础数据审核工具（Flask 外壳：Windows + 统信 UOS/麒麟）

`unified/shell-flask` 是跨平台 Flask 外壳，与 `unified/shell-pywebview`（仅
Windows）共用 `unified/core` 的唯一业务核心（`base_audit` 后端 +
`frontend` 前端 + `tests` 测试）。本目录只放外壳与平台启动/打包脚本。

## 架构

- **前后端唯一接口是桥接层**：页面注入脚本把 `bridge.api.方法名(...)`
  代理为 `POST /api/方法名`，前端 HTML 两个平台一字不改。
- **引擎按操作系统自动识别**（`core/base_audit/engines`）：
  - Windows：Microsoft Excel / WPS 表格 COM；
  - 统信 UOS/麒麟：LibreOffice Calc（`soffice --headless` 重算 +
    openpyxl 原生管线，条件格式走 OOXML 规则求值，与 WPS 兜底同族）。
- 平台差异收口：文件打开（`os.startfile` / `xdg-open`）、引擎列表、
  计算管线（COM / native）均在 core 的 engines 与 native 包，外壳不判断。

## Windows 运行与打包

```bat
启动Flask.bat                 :: 运行（默认 http://127.0.0.1:8750/）
打包Flask.bat                 :: PyInstaller 成品 EXE（现代 Windows）
打包Flask-Win7.bat            :: Win7 兼容版（Python 3.7 + 离线 wheels）
```

## 统信 UOS/麒麟 运行与打包

```sh
./启动Flask-UOS.sh            # 运行（python3 + flask + openpyxl）
./打包Flask-UOS.sh            # 源码发行包 zip/tar.gz + SHA256
```

UOS 环境要求：python3、`pip3 install -r requirements.txt`（pywin32 在
Linux 自动跳过）、系统安装 LibreOffice Calc（含玲珑商店版，引擎自动探测）。

**验收状态**：模板识别、配置中心、历史读取等 openpyxl 功能已可离线使用；
完整审核流程（soffice 重算 + 条件格式）须在真实 UOS 环境按
`unified/docs/TEST_UOS.md` 清单验收后方可正式使用。联合模板制作等
COM 专属功能暂仅支持 Windows 外壳。

## 测试

```sh
cd unified/core && PYTHONPATH=src python3 -m pytest tests/ -q
```

## 与其他发行线的关系

- 业务功能一律改 `unified/core`，两个外壳同时生效；归档线见
  `external/Flask`、`external/pywebview2`（各含归档说明）。
