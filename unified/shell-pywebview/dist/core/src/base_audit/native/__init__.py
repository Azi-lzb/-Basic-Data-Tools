"""统信 UOS/麒麟 原生审核管线（openpyxl + soffice 重算，无 COM/UNO 依赖）。

模块职责：
- ``openpyxl_workbook``：模板读取、命名区域、公式复制、外部表复制、问题提取；
- ``conditional_scan``：条件格式 OOXML 规则求值（与 Windows WPS 兜底同族，
  表达式求值复用 core 的 ``conditional_format.evaluate_expression_formula``）；
- ``history_io``：历史审核工作簿读写（对齐 core history.py 表名/表头）；
- ``audit_flow``：逐文件审核编排（复制→外部表→公式复制→soffice 重算→
  问题提取→条件格式→历史/汇总）；
- ``summary``：任意行/固定行汇总、汇总表合并、组合联合核查表/组合工作表。

安全边界与 Windows 版一致：原始报送文件只读（SHA-256 前后校验），全部修改
写入 ``_审核版.xlsx`` 副本。
"""
