# 统信 UOS/麒麟 实机验收清单（Flask 外壳 + native 管线）

> 以下项目必须在真实统信 UOS/麒麟 环境、用真实报送模板逐项验收通过后，
> UOS 版审核结果才可作为正式结论。Windows 侧的 mock 测试无法替代。

## 0. 环境准备

- [ ] `python3 --version`（3.8+）；`pip3 install -r requirements.txt`
- [ ] LibreOffice Calc 已安装：
  - [ ] 直装版：`soffice --version` 可执行
  - [ ] 玲珑商店版：`ll-cli list` 含 `org.libreoffice.libreoffice`
- [ ] 启动：`./启动Flask-UOS.sh`，浏览器打开 `http://127.0.0.1:8750/`
- [ ] 左下角/设置中心引擎只显示「自动 / LibreOffice Calc」

## 1. 纯 openpyxl 功能（不依赖 soffice）

- [ ] 源数据目录选择 → 文件清单、模板自动识别（共识逻辑）正确
- [ ] 数据期识别（文件名/文件夹名）
- [ ] 机构说明文件卡片提醒
- [ ] 设置中心：功能/流程编排读写、恢复默认、执行前确认开关
- [ ] 历史审核配置选择 + 历史记录分页/筛选读取

## 2. 审核流程（soffice 重算 + native 管线）

用一期真实报送目录 + 正式模板（如「！10.基础数据-单位贷款」）：

- [ ] 表结构比对：通过与跳过名单与 Windows 版一致
- [ ] 公式复制：审核副本含模板公式；`fullCalcOnLoad` 已置位
- [ ] soffice 重算：每个审核副本生成缓存值；重算前后文件时间戳变化
- [ ] 公式错误扫描：`#N/A/#REF!/#NAME?/#VALUE!` 等进入结果并与 Windows 版对齐
- [ ] 问题提取条数与 Windows Excel/WPS 版一致（逐条核对差异并记录）
- [ ] 条件格式提取（OOXML 规则求值）：触发单元格与 Windows 版 DisplayFormat
      结果一致；colorScale/dataBar/iconSet 不误报，计入不支持数

## 3. 汇总流程

- [ ] 「汇总校验结果说明」：业务说明/任务说明无历史三列；本地校验结果
      末尾三列（历史触发条数/历史分类说明/审核结果）正确回写
- [ ] 「汇总表合并」按表头分组正确
- [ ] 运行日志 `流程名_运行日志_时间戳.xlsx` 每功能一个 sheet

## 4. 公式兼容性基准（AGENTS 强制项）

对真实模板逐函数验证 soffice 与 Excel 结果一致，不一致的记录到本文档：

- [ ] `INDEX/MATCH`（精确匹配）、`VLOOKUP`
- [ ] `SUMIFS`（固定范围）、`IF/AND/OR/ABS/ISBLANK`
- [ ] 命名区域（工作簿级 + 工作表级）、跨表引用
- [ ] 错误值传播（`#N/A` 等）与缓存值落盘一致

## 5. 安全边界

- [ ] 审核前后原始报送文件 SHA-256 不变（程序自动校验 + 人工抽查）
- [ ] 审核副本命名 `<原名>_审核版.xlsx`；soffice 私有 profile 不触碰
      用户正打开的 LibreOffice 窗口
- [ ] 断网/无桌面环境（headless）下流程可运行

## 6. 明确不支持（应报错而非静默）

- [ ] 「组合联合核查表」「组合工作表」→ 提示"暂仅支持 Windows 外壳"
- [ ] 制作联合模板（shell-flask UOS 菜单入口行为与 Windows 一致性待查）

## 差异记录

| 项目 | Windows (Excel/WPS) | UOS (soffice+OOXML) | 结论 |
| --- | --- | --- | --- |
| （验收时填写） |  |  |  |
