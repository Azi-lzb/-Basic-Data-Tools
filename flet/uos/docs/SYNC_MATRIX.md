# Windows / 统信原生版功能对照

| Windows 执行模块 | 统信原生实现 | 统信实机重点 |
|---|---|---|
| 检查_命名区域存在性 | `openpyxl_workbook.named_ranges` + `workflow` | 同名及“名称_后缀”匹配 |
| 核对_表结构比对 | `preflight_xlsx` | 合并单元格、表头静态值 |
| 修改_外部文件添加 | `openpyxl_workbook.add_external_sheets` | 隐藏表、重名表必须报错 |
| 修改_公式校验复制 | `copy_formula_ranges` + formulas + IronCalc 交叉验证 / 可选 LibreOffice | 公式结果、外链、保存缓存 |
| 汇总_公式校验结果提取 | `extract_issues` | 错误、软性、六段结果格式 |
| 汇总_条件格式结果提取 | `conditional_format` + UNO | 实际显示填充色 |
| 汇总_审核结果输出 | `audit._write_summary` | 与 Windows 表头、历史说明一致 |
| 汇总_任意行汇总 | `region_summary.run_region_summaries`（xlsx=openpyxl，xls=pandas/xlrd） | 多行表头、单元格区域 |
| 汇总_固定行汇总 | `region_summary.run_region_summaries` | 固定框选行 |
| 汇总_汇总表合并 | `merge_workbook_tables` | 不同表头不得混合 |
| 修改_组合联合核查表 | `merge_by_first_filename_part` | 样式、隐藏表、机构文件名规则 |
| 修改_组合工作表 | `combine_by_plan` | 正则与名称/关键字方案 |
| 联合模板制作 | `create_combined_template` CLI | 外部引用改为本工作簿引用 |

## 流程与工作台同步

- `workflow.run_flow` 按“执行流程”的顺序、输入、输出文件名、是否输出结果和
  “失败后处理”执行，不再把这些字段只当作界面说明；
- 外部文件添加/公式校验复制的审核副本可以被后续汇总步骤作为“输入”引用；
- Flet 工作台和 CLI 根据配置中的执行模块判断流程是否需要模板，重命名的组合流程
  不会再因流程名不同而要求选择模板；
- 所有源工作簿在审核处理后重新校验 SHA-256，原文件变化会明确失败。

所有统信实现位于 `uos_audit/`。该目录不导入 pywin32、Excel COM、Windows WPS COM、Wine、pywebview、GTK 或 WebKitGTK。
