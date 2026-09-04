# 基础数据审核工具开发约定

## 项目目标

本项目只实现“基础数据审核”，不迁移旧 VBA 中的通用汇总、台账拆分、文件转换等功能。

第一版必须完成：

1. 批量读取机构报送的 `.xlsx` 文件。
2. 根据 Excel 模板把校验公式及格式复制到审核副本。
3. 调用本机 Microsoft Excel 重新计算公式。
4. 提取结构化问题并生成问题汇总。
5. 使用 Excel 文件保存历史问题，提示新增、连续出现、再次出现和已整改。
6. 原始报送文件只读留存，所有修改写入审核副本。

## 目录规划

```text
.
├─ AGENTS.md                 # 开发约定和目录规划
├─ README.md                 # 项目说明与快速查阅
├─ 基础数据审核工具使用说明.docx # 工作台“小书”打开的正式使用说明
├─ run.py                    # 程序入口（CLI 与 web 工作台）
├─ frontend/
│  └─ web/                   # 本地 HTML 前端（由 Claude + DSV4Flash 维护）
├─ src/
│  └─ base_audit/
│     ├─ __init__.py
│     ├─ web_app.py          # pywebview 前后端接口和任务调度
│     ├─ models.py           # 审核任务和问题数据结构
│     ├─ template.py         # 模板规则读取与自检
│     ├─ discovery.py        # 模板推荐、数据期识别、机构说明文件清单
│     ├─ settings.py         # 常用目录及最近使用路径
│     ├─ single_instance.py  # 单实例锁
│     ├─ excel_com.py        # Excel COM 生命周期、命名区域调查与工作簿操作
│     ├─ history.py          # Excel 历史审核结果读取、匹配与保存
│     ├─ name_config.py      # 内置模块/流程与 流程配置.json 的自定义流程、按钮读取
│     ├─ preflight_xlsx.py   # openpyxl 只读的表结构比对与报告写入
│     ├─ region_summary.py   # 任意行/固定行/汇总表合并的区域汇总
│     ├─ feature_log.py      # 运行日志收集：每个功能一个 sheet，写“流程名_运行日志_时间戳.xlsx”
│     ├─ merge_org.py        # 按文件名第一段合并同一机构的多个报送工作簿
│     └─ service.py          # 批量审核与流程编排（含运行日志收集）
├─ 2026-07-31/
│  ├─ 模板文件/              # 用户维护的正式审核模板
│  ├─ 单位贷款核查表/        # 本期真实报送核查表（只读）
│  ├─ 单位贷款校验结果说明/  # 本期校验结果及报送说明 + 汇总信息
│  ├─ 单位贷款202607.xls     # 参考校验结果说明汇总（旧 VBA 输出）
│  └─ 报表工具-校验结果-！金融基础数据-*.xls  # 参考校验结果
├─ data/
│  ├─ 流程配置.json          # 功能、自定义流程、流程显示设置（由界面维护，Windows 下隐藏）
│  └─ 用户设置.json / 模板索引.json
├─ tests/
│  ├─ test_config_history.py # 配置初始化与流程读取测试
│  ├─ test_discovery.py      # 模板推荐/文件分类测试
│  ├─ test_external.py       # 外部工作表规划测试
│  ├─ test_feature_log.py    # 运行日志写入与命名区域调查测试
│  ├─ test_history.py        # 历史状态测试
│  ├─ test_preflight_xlsx.py # 表结构比对与结构报告测试
│  ├─ test_region_summary.py # 区域汇总测试
│  ├─ test_template.py       # 模板字段和规则校验测试
│  └─ test_web_app.py        # web 接口状态测试
├─ packaging/
│  ├─ requirements/          # Windows 运行与打包依赖
│  ├─ specs/                 # 标准版、Win7 x64/x86 PyInstaller 配置
│  ├─ hooks/                 # Win7 pywebview 打包钩子
│  └─ scripts/               # 构建及成品收尾脚本
├─ tools/
│  ├─ build_release.py       # 发布包与 SHA256
│  ├─ build_user_guide.py    # 生成使用说明 docx
│  ├─ inspect_formal_templates.py  # 只读检查正式模板协议
│  ├─ benchmark_calculation.py     # 公式计算性能基准
│  ├─ export_docx_pdf_word.py      # docx 转 PDF（Word COM）
│  └─ render_pdf_pages.py          # PDF 页面渲染为 PNG
└─ uos_audit/                # 当前 Flet/Linux 实现及跨平台统一迁移来源
```

### 当前项目目录（优先于上方历史目录图）

根目录的活跃发行线是 `unified/`（core 共享核心 + shell-flask / shell-pywebview 双外壳，含 engines 引擎适配层）；`external/pywebview2/`、`external/Flask/`、`flet/` 均为归档/冻结基线（各目录下有归档说明），只作对照，不再开发新功能。不得再在根目录创建 `src/`、`frontend/`、`packaging/`、`tests/`、`tools/` 或 `uos_audit/`：

```text
.
├─ unified/                  # 活跃发行线：core 共享核心 + 双平台外壳
│  ├─ core/                  # 唯一业务核心：src/base_audit（含 engines/ 引擎适配层：
│  │                         #   protocol.py、excel_wps.py、libreoffice*.py）+ frontend/web + tests/
│  ├─ shell-flask/           # Flask 外壳（本地服务 + 浏览器，面向 UOS/国产化）
│  └─ shell-pywebview/       # pywebview 外壳（Windows 桌面窗口）
├─ external/Flask/ → 已归档（2026-09-04，归档说明.md 在其目录下）
│  # 前一代 Flask 线；含确认弹窗/公式批量读性能修复/bridge.api 改名等待移植改进
├─ external/pywebview2/ → 已归档（2026-09-04，归档说明.md 在其目录下）
│  # Windows pywebview 冻结基线；唯一 Win7 发行链路；只读对照，不再开发
├─ flet/windows/             # 冻结基线：Windows Flet 项目；仅 Excel/WPS COM
│  ├─ main.py, run.py, src/, tests/, packaging/
│  ├─ data/, 历史审核说明.xlsx
│  └─ requirements-windows.txt、打包Flet-windows.bat、启动Flet-windows.bat
├─ flet/uos/                 # 冻结基线：UOS/麒麟 Flet 项目；LibreOffice UNO 实现的移植来源
│  ├─ main.py, run.py, src/, tests/, packaging/
│  ├─ data/, 历史审核说明.xlsx
│  └─ requirements.txt、打包Flet-UOS.sh、启动Flet-UOS.sh
├─ external/                 # Flutter 工具链归档、历史研究资产、旧产物与聊天记录
├─ 2026-07-31/、reference/   # 真实业务数据/模板与参考文件，只读留存
└─ AGENTS.md、CLAUDE.md 等    # 根目录共享协作与业务协议文件
```

### unified 线同步规则（活跃开发约定）

- 业务功能一律改 `unified/core/`（后端 `src/base_audit`、前端 `frontend/web/index.html`、测试 `tests/`），两个外壳 `unified/shell-flask/`、`unified/shell-pywebview/` 只放平台差异，不得各自复制业务代码。
- 前端接口写法为 `pywebview.api.方法名(...)`——这是 unified 现行的历史命名，指核心的 `web_app.py`，**不是** `external/pywebview2/` 归档项目；shell-flask 通过注入桥接层把它代理为 `POST /api/方法名`。改接口先改 core 的 `web_app.py`，外壳不写接口。
- 平台差异（Excel/WPS COM 与 LibreOffice UNO）封装在 `unified/core/src/base_audit/engines/`；业务流程层不得直接判断 COM、UNO 或操作系统。
- 改动后在 Windows 上运行 unified/core 的 pytest 全绿才可合入；UOS/LibreOffice 专属路径需在真实 UOS 环境验收。
- `external/Flask/` 中存在三项 unified 尚未包含的改进（执行前确认弹窗、外部公式批量读性能修复、bridge.api 改名），移植前不得删除该归档。

## 跨平台统一架构与发布

- 项目后续采用“一套业务核心、两个平台发行包”，不得继续维护两套会逐渐分叉的审核逻辑。模板协议、文件发现、流程编排、表结构比对、公式/格式复制、问题模型、历史记录、汇总输出、运行日志和 Flet 界面均应进入跨平台共享层。
- 不设置可由用户手动切换的“Windows / 统信 UOS”系统按钮。程序在启动时自动识别操作系统和已安装能力；左上角只显示当前环境与可用办公套件状态。用户只在设置中心选择当前环境真正可用的计算引擎，不可用选项应禁用并解释缺失依赖。
- 平台差异必须封装为适配器，业务流程不得直接判断 COM、UNO 或操作系统：
  - Windows 适配器：Microsoft Excel COM（`Excel.Application`）、WPS COM（`ket.Application` / `KET.Application`）；COM 依赖必须延迟导入。
  - Linux/UOS 适配器：仅 LibreOffice Calc/UNO；不得依赖或打包 `pywin32`、`formulas`、IronCalc 或 EFC。
- 条件格式属于“计算后的渲染结果读取”，不是 `openpyxl` 能完整替代的普通样式读取：
  - Windows Excel/WPS 使用实时计算后的 `DisplayFormat.Interior.Color` 与基础 `Interior.Color` 对比，只提取实际改变填充色的单元格。
  - UOS/Linux 使用 LibreOffice UNO 读取实时 `CellBackColor`，并与 OOXML 基础填充色对比。
  - 两个平台都只扫描模板命名区域 `条件格式区域`，读取批注作为问题描述，并结合 `表结构区域` 定位指标。
  - 无可用办公套件渲染引擎时，条件格式不得猜测颜色；必须明确失败并提示安装当前平台要求的办公套件。
- 公式计算和条件格式渲染可以由同一办公套件适配器完成，但在共享业务层中必须是两个独立能力：`recalculate_workbook` 与 `extract_rendered_conditional_formats`。Windows 用 Excel/WPS COM，UOS 用 LibreOffice Calc/UNO。
- 发布物保持分平台生成，但来自同一源码版本和同一测试基线：
  - Windows：生成 Windows EXE，包含 Flet 共享界面、共享业务核心及 Excel/WPS COM 适配器；不包含 LibreOffice 或纯 Python 公式引擎。
  - 统信 UOS/麒麟：生成 Linux 安装包，包含相同 Flet 界面、共享业务核心及 LibreOffice 适配层，不包含 `pywin32`、pywebview、Windows COM 或纯 Python 公式引擎。
- 两个平台必须读取同一份 `流程配置.json`、模板命名区域协议和 `历史审核说明.xlsx`，并生成相同字段、相同超链接语义和相同运行日志口径。平台适配器造成的结果差异必须通过真实模板基准测试记录，不得在业务层静默修正。
- 迁移期间以 Windows 已验证实现和 `uos_audit/` 已同步实现为来源，逐模块抽取共享核心；每抽取一个模块必须同时运行 Windows 与 UOS 测试，确认后再删除旧的重复实现。

### 条件格式提取的引擎差异（重要）

条件格式分「布尔触发型」和「渐变型」两类。本项目只关心“规则是否触发”（找被标红的可疑单元格），因此只正式支持布尔触发型：

- **布尔触发型（支持）**：`cellIs`（单元格值比较）与 `expression`（公式返回 TRUE/FALSE）。这是报送系统用来标红的常规类型。
- **渐变型（不支持/不判定）**：`colorScale`（色阶）、`dataBar`（数据条）、`iconSet`（图标集）。这三类给范围内每个单元格做连续渐变/分档，没有“触发/不触发”边界，不属于本项目语义。

各引擎/路径的行为差异：

- **Excel COM**：读实时 `DisplayFormat.Interior.Color` 与基础 `Interior.Color` 对比。对 colorScale/dataBar/iconSet 会把范围内每个单元格都误报为“触发”（因为每个格子的颜色都变了）；模板应避免使用这三类。
- **WPS COM**：WPS 在隐藏启动 + `ScreenUpdating=False` 下，`DisplayFormat.Interior.Color` 可能返回静态色或 `None`。必须先 `_prepare_conditional_format_sheet`（`ScreenUpdating=True → 激活工作表 → CalculateFullRebuild`）刷新；仍读不到时抛「实际显示颜色」错误。
- **OOXML 兜底**：WPS 读不到实际颜色时，不猜颜色、不报触发，改为直接用 openpyxl 读 `.xlsx` 里存的条件格式规则，用共享的 `evaluate_expression_formula` 解析 cellIs/expression 并判断触发（相对/绝对引用按 AppliesTo 锚点换算）。无法可靠解析的规则计入 `unsupported_count`，运行日志写明「WPS 条件规则暂不支持」。colorScale/dataBar/iconSet 在兜底路径下正确跳过（不误报）。
- **统一描述**：所有路径的条件格式「描述」统一为 `条件格式规则：<公式>`（公式按锚点换算到目标单元格，如 `AND(C26>0,C26>1)`），不再是通用的「条件格式填充已触发，请核实」。有单元格批注时批注优先。
- **合并单元格**：COM 路径读 `MergeArea` 父标签来定位行/列指标；OOXML 兜底是简化版、不处理合并单元格。表结构区域含合并单元格时，两者指标可能略有差异。

四路（Flet-Win / pywebview2-Win × Excel / WPS）在「cellIs / expression 规则 + 表结构区域无合并单元格」前提下，本期审核结果一致。

### 正式发行包依赖精简（强制）

- 纯 Python 公式计算引擎在真实模板上不具备可接受的可靠性。`formulas`、IronCalc、EFC 仅可保留为源码研究/历史基准资产，**不得进入任何正式 Windows、UOS/麒麟 DEB 或 `.run` 发行包**，也不得作为正式界面的可选计算引擎。
- Windows 正式包只保留当前业务所需的最小依赖：Flet 桌面运行时、共享业务核心、`openpyxl`（及其必要 XML 依赖）、`pywin32` 与 Microsoft Excel / WPS COM 适配器。不得打入 LibreOffice、UNO、`formulas`、IronCalc、EFC，或它们引入的 `numpy`、`scipy`、`numpy-financial`、`schedula`、`regex`、`tqdm` 等传递依赖；Excel/WPS 由目标机器自行安装。
- UOS/麒麟正式 DEB 或 `.run` 包只保留 Flet 桌面运行时、共享业务核心、`openpyxl`（及其必要 XML 依赖）和 LibreOffice 适配层。公式重算与条件格式渲染统一调用目标系统安装的 LibreOffice Calc/UNO；不得打入 Excel/WPS COM、`pywin32`、`formulas`、IronCalc、EFC 或其 `numpy`、`scipy`、`numpy-financial`、`schedula`、`regex`、`tqdm` 等依赖。安装说明或 DEB 依赖必须明确要求 LibreOffice Calc。
- 打包前必须以平台专用 requirements/锁定清单做依赖审计；新增库必须说明其运行时用途。仅用于开发、测试、基准或构建的依赖必须放入独立的 requirements 文件，不得随正式包分发。
- 如果源码中暂时保留可选引擎的导入，必须延迟导入；在精简发行包中缺失时，程序应明确显示“该诊断引擎未随发行包提供”，不能导致启动失败或静默回退。

## 文件与数据边界

- `src/` 中不得写死机构名称、数据期、绝对路径或具体校验公式。
- 校验公式必须保存在 Excel 模板中；Python 只负责复制、计算和提取。
- 根目录的 `历史审核说明.xlsx` 仅保存“历史审核结果”，与程序版本分离；升级程序不得覆盖用户历史。
- 初始模块和常用流程写在程序中；用户可在设置中心修改，保存后的有效配置、流程和按钮均存放在 `data/流程配置.json`，升级时一并保留；“恢复默认设置”可还原初始状态。
- `data/用户设置.json`、`data/模板索引.json` 为本机运行数据，升级程序时一并保留。
- 本项目不内置伪数据；真实报送文件（如 `2026-07-31/` 下的核查表、校验结果说明）只读留存。
- 默认输出到用户选择的目录；不得覆盖原始报送文件。
- EXE 只包含程序代码和运行依赖；正式模板、历史库及用户设置必须保留为 EXE 同级外部目录。
- 审核副本文件名统一为 `<原文件名>_审核版.xlsx`。

## 模板约定

新式模板可包含 `审核规则` 工作表，第一行为固定字段：

```text
规则编号,启用,报表代码,工作表,公式单元格,定位单元格,级别,问题说明
```

- 校验公式本体放在模板对应业务工作表的“公式单元格”中。
- `启用` 使用 `是/否`。
- `级别` 第一版支持 `错误/核实/提示`。
- 规则编号必须唯一且长期稳定，历史匹配不得依赖问题文案。
- 公式结果为空表示未触发；非空表示触发。

正式模板定位使用 Excel 原生命名区域；程序只认识内置模块和 `流程配置.json` 中已保存的自定义流程：

- `修改_公式校验复制`：默认命名区域为 `校验区域`，用于复制公式、重算并提取问题；可使用不连续区域。
- `核对_表结构比对`：默认命名区域为 `表结构区域`，仅核对固定表头，不应包含机构名称、机构代码、日期等动态数据。
- “模块化功能”固定列为 `功能名、执行模块、命名区域名、输入、输出、备注`，其中“输入 / 输出”用于说明模块通常读取和产出什么，不参与调度。执行模块统一使用 `检查_命名区域存在性、核对_表结构比对、修改_外部文件添加、修改_公式校验复制、修改_组合联合核查表、修改_组合工作表、汇总_公式校验结果提取、汇总_条件格式结果提取、汇总_审核结果输出、汇总_任意行汇总、汇总_固定行汇总、汇总_汇总表合并`。条件格式提取优先扫描命名区域 `条件格式区域` 内被条件格式实际改变填充色的单元格，不限定红色；读取批注为描述，并用 `表结构区域` 交叉定位行、列指标。命名区域名按“相同或前缀+下划线”匹配，可用逗号/顿号配置多个必需区域。执行流程固定列为 `流程名、顺序、功能名、启用、失败后处理、是否输出结果、输出文件名、输入、输出、备注`；失败后处理仅支持“停止/跳过”，每个流程至少有一个启用步骤输出结果。公式/条件格式结果提取填写相同“输出”会追加到同一最终工作表；`汇总_审核结果输出`负责落盘。自定义流程可在流程编排中勾选“显示在主界面”，入口文字固定为流程名；不再维护独立自定义按钮。恢复默认设置后回到初始功能、流程和显示设置。
- 不得将定位规则写死在 Python 代码中。

旧 VBA 批注协议仅用于未迁移模板的临时兼容：

- `本地校验区域1` 与 `本地校验区域#1` 确定一块复制区域；序号可递增。
- `联表校验区域1` 与 `联表校验区域#1` 确定一块复制区域；序号可递增。
- 程序整块复制校验区域，并从其中包含“错/错误/失败/硬错/硬性/核实/提示/警告”结果的公式提取问题。
- 旧模板没有规则编号时，使用“模板名称+工作表+公式单元格”生成稳定规则编号。
- 旧模板公式单元格批注可增加 `规则编号：XXX`；固定编号的问题身份不得依赖单元格位置。
- 新增公式建议统一返回 `级别|对比值|参考值|差值|详细说明`；结果表必须拆列并保留完整公式结果，同时兼容旧四段 `级别|指标名称|当前值|问题说明` 及更早格式。

## 实现约束

- 目标环境同时包括 Windows 与统信 UOS/麒麟；正式功能依赖目标平台已安装且经真实模板验证的办公套件或公式引擎。
- 当前发行版是 Windows EXE，公式计算依赖 Windows COM（优先
  `Excel.Application`，未安装 Excel 时尝试 WPS 表格
  `ket.Application` / `KET.Application`）。Windows 版 WPS 可作为兼容
  引擎使用，但应以真实模板验证公式结果。
- Win7 兼容版使用专用 `HZPBCwin7` 构建环境（Python 3.7、PyInstaller
  4.10、pywebview 5.4），通过“打包pywebview2-windows-win7x64.bat”生成至
  `dist/基础数据审核工具_Win7.exe`。该包仅包含 x64 WebView2
  桥接文件；目标机器仍需自行安装兼容的 WebView2 Runtime、.NET Framework
  4.6.2+，以及 Excel 或 Windows 版 WPS。Win7 已停止支持，须在实际目标
  环境验证后方可正式交付。
- 32 位 Win7 仅在目标机器为 32 位 Windows 时使用；它必须由专用
  `HZPBCwin7x86`（win-32 Python 3.7）环境构建，不能从 64 位 EXE 转换。
  通过“打包pywebview2-windows-win7x86.bat”生成至
  `dist/基础数据审核工具_Win7_x86.exe`。
- 统信 UOS、麒麟等 Linux 环境不能直接运行 Windows EXE，也不支持 `pywin32`
  / Windows COM。当前 Linux 原生实现位于 `flet/uos/`：Flet 是唯一图形入口，
  `openpyxl` 负责 OOXML 文件处理，LibreOffice UNO / `soffice --headless` 是唯一正式的公式计算与条件格式读取引擎；
  Linux 发行包不得导入或包含 Windows `excel_com.py`、`pywin32`、pywebview、WebKit、`formulas`、IronCalc 或 EFC。
  实际 UOS 环境和真实模板的公式缓存一致性仍须逐项验收后方可交付。
- Linux 版须先逐项验证 `INDEX`、`MATCH`、`VLOOKUP`、命名区域、公式错误
  `#N/A/#REF!` 及保存后的缓存结果与 Excel 一致。WPS Linux 的 JS 加载项
  可作为后续界面集成方案，但不是现有 Python COM 后端的直接替代。
- Excel 操作使用独立的隐藏 COM 实例，禁止连接或关闭用户已打开的 Excel。
- 每个工作簿必须在 `finally` 中关闭；Excel 实例必须可靠退出。
- 批量任务中模板只读打开一次。
- 复制公式前必须只读完成模板体检和报送文件结构匹配；不匹配文件不得生成审核副本。
- 表结构比对是批量报送热路径：模板命名区域必须通过 COM `.Value2` 整批读取且每次流程只读一次；报送文件必须使用 `openpyxl` 只读流式模式，不得为结构比对启动 Excel 或逐格 COM。
- 多个表结构区域必须先解析坐标并合并为尽量少的紧凑读取批次；比对循环不得使用正则或日期、符号、空白标准化，只跳过模板空白单元格。
- 每次流程运行必须生成运行日志 `流程名_运行日志_时间戳.xlsx`（一个功能一个 sheet，含模板体检）；不再单独生成“审核前检查”报告。
- 第一版优先简单、可追踪，不加入插件系统、复杂配置中心或通用报表功能。
- 工作台命令栏只能调用工具内置动作，禁止将用户输入直接交给 shell 执行。
- 选择外部辅助文件后，程序优先从模板公式识别工作表引用，识别不到才复制外部文件全部工作表。复制发生在每个审核副本中、模板公式复制之前；原始报送文件和外部文件都只读留存。
- 外部文件若与报送文件存在同名的待复制工作表，必须明确失败，不得覆盖原表或自动改名。
- 代码读取 `历史审核说明.xlsx` 的“历史审核记录”时必须允许用户增加“机构反馈”和“审核意见”，程序更新时不得覆盖这些人工字段。

### UOS 公式计算与条件格式策略

- UOS/麒麟正式流程统一使用 LibreOffice Calc/UNO 计算公式、保存缓存并读取条件格式实际渲染结果；不再以纯 Python 公式引擎作为正式或回退路径。
- 公式错误（如 `#REF!`、`#NAME?`、`#VALUE!`、`#N/A`）必须在 LibreOffice 计算后扫描、定位并写入运行日志；条件格式同样由 LibreOffice 的实际显示结果提取。
- `external/research/engine_trials/` 中的 `formulas`、IronCalc、EFC 小样本仅是历史研究材料。它们不构成正式功能、运行依赖或打包依据；如后续移除这些资产，不得影响 LibreOffice 正式链路。
- 模板仍优先使用固定范围的 `SUMIFS`、`IF`、`ABS`、`AND`、`OR`、`ISBLANK`、精确匹配 `INDEX/MATCH` 或 `VLOOKUP`；复杂或易失公式以 LibreOffice 与 Windows Excel 的真实模板基准验证为准。
- 接入新公式引擎或放宽函数白名单前，必须为该函数组合新增可重复测试，并验证：计算结果、错误值、命名区域/跨表引用，以及输出工作簿保留公式和可读取的缓存值。

## 前后端任务分配

- Claude + DSV4Flash 只维护 `frontend/web/` 内的本地 HTML、CSS、JavaScript 和界面交互；详细约定见 `CLAUDE.md`。
- Codex 只维护 Python 后端、Excel COM、命名区域、模板协议、历史库、汇总、自动测试及打包。
- 前端不得修改 `web_app.py` 中既有 `pywebview.api` 接口，也不得修改 `service.py`、`excel_com.py` 等审核后端。
- 后端变更若需要新增前端接口，应先在 `CLAUDE.md` 更新接口说明。

## 验证要求

- 单元测试覆盖四种历史状态：新增、连续出现、再次出现、已整改。
- 示例端到端测试至少包含两家机构、两期数据和三类规则。
- 验证原始文件哈希在审核前后保持不变。
- 验证审核副本保留公式，问题汇总与公式结果一致。
- 验证模板规则编号重复、工作表缺失、公式缺失时明确失败，不得当作审核通过。
