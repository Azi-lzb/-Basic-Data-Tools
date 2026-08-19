from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from src.base_audit.service import (
    AuditService,
    CONFIG_HISTORY_SHEET,
    _output_prefix,
    _resolve_process_source,
)
from src.base_audit.name_config import (
    FLOW_HEADERS,
    FLOW_LEGACY_HEADERS,
    FLOW_PREVIOUS_HEADERS,
    FLOW_PREVIOUS_9_HEADERS,
    FlowStep,
    FIXED_ROW_SUMMARY_FUNCTION,
    FORMULA_COPY_FUNCTION,
    NAMED_RANGE_CHECK_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    USED_RANGE_SUMMARY_FUNCTION,
    WORKBOOK_TABLE_MERGE_FUNCTION,
    initialize_config,
    load_feature_mappings,
    load_flow_steps,
    reset_default_configuration,
    validate_config,
)
from src.base_audit.excel_com import _a1_address


class ConfigHistoryTests(unittest.TestCase):
    def test_config_workbook_is_the_history_destination(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            config = root / "config.xlsx"
            workbook = Workbook()
            workbook.active.title = "区域映射"
            workbook.save(config)
            destination, sheet, migration = AuditService(
                config_path=config
            )._history_storage(root / "历史审核库.xlsx")
        self.assertEqual(config.resolve(), destination)
        self.assertEqual(CONFIG_HISTORY_SHEET, sheet)
        self.assertEqual((root / "历史审核库.xlsx").resolve(), migration)

    def test_initialize_config_restores_default_region_mappings(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            mappings = load_feature_mappings(config, Path("任意模板.xlsx"))
        names = {mapping.feature_type: mapping.range_names for mapping in mappings}
        named_checks = {mapping.name: mapping.range_names for mapping in mappings if mapping.feature_type == NAMED_RANGE_CHECK_FUNCTION}
        self.assertEqual(("校验区域",), names[FORMULA_COPY_FUNCTION])
        self.assertEqual(("校验区域",), named_checks["检查校验区域"])
        self.assertEqual(("表结构区域",), named_checks["检查表结构区域"])
        self.assertEqual(("任意行汇总区域", "表头区域"), named_checks["检查汇总区域"])
        self.assertEqual(("表结构区域",), names[STRUCTURE_COMPARE_FUNCTION])
        self.assertEqual(("任意行汇总区域",), names[USED_RANGE_SUMMARY_FUNCTION])
        self.assertEqual(("固定行汇总区域",), names[FIXED_ROW_SUMMARY_FUNCTION])
        self.assertEqual((), names[WORKBOOK_TABLE_MERGE_FUNCTION])
        self.assertNotIn("区域汇总", names)

    def test_reset_defaults_keeps_history_and_custom_buttons(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            from openpyxl import load_workbook
            workbook = load_workbook(config)
            workbook["模块化功能"].append(("临时模块", FORMULA_COPY_FUNCTION, "临时区域", "否", "", "待删除"))
            workbook["执行流程"].append(("临时流程", 10, "临时模块", "是", "停止", "是", "", "待删除"))
            workbook["自定义按钮"].append(("我的按钮", "审核前检查", "是", "保留"))
            workbook["历史审核结果"].append(("文件.xlsx", "Sheet1", "A1", "错误", "指标", "说明", "1", "", "", "规则｜1", "历史", "意见"))
            workbook.save(config)
            workbook.close()
            reset_default_configuration(config)
            check = load_workbook(config, data_only=True)
            try:
                self.assertNotIn("临时模块", [row[0] for row in check["模块化功能"].iter_rows(min_row=2, values_only=True)])
                self.assertNotIn("临时流程", [row[0] for row in check["执行流程"].iter_rows(min_row=2, values_only=True)])
                self.assertEqual("我的按钮", check["自定义按钮"].cell(2, 1).value)
                self.assertEqual("文件.xlsx", check["历史审核结果"].cell(2, 1).value)
            finally:
                check.close()

    def test_reset_defaults_recovers_from_broken_module_sheet_headers(self):
        with TemporaryDirectory() as folder:
            config = Path(folder) / "config.xlsx"
            workbook = Workbook()
            workbook.active.title = "模块化功能"
            workbook.active.append(("被修改的表头",))
            workbook.save(config)
            workbook.close()
            reset_default_configuration(config)
            check = load_workbook(config, data_only=True)
            try:
                self.assertEqual(
                    ("功能名", "功能类型", "命名区域名", "是否限定工作簿", "工作簿关键字", "备注"),
                    tuple(cell.value for cell in check["模块化功能"][1]),
                )
                self.assertIn("执行流程", check.sheetnames)
            finally:
                check.close()

    def test_default_audit_flow_only_generates_final_module_result(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            steps = load_flow_steps(config, "汇总核查表校验")
        self.assertEqual(
            ("检查校验区域", "检查表结构区域", "表结构比对", "外部文件添加", "公式校验复制", "校验结果提取"),
            tuple(step.feature_name for step in steps),
        )
        self.assertEqual(
            (False, False, False, False, False, True),
            tuple(step.output_result for step in steps),
        )

    def test_default_explanation_flow_checks_structure_before_summary(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            steps = load_flow_steps(config, "汇总校验结果说明")
        self.assertEqual(
            ("检查表结构区域", "检查汇总区域", "表结构比对", "任意行汇总"),
            tuple(step.feature_name for step in steps),
        )

    def test_config_requires_an_output_step_per_flow(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            workbook = load_workbook(config)
            sheet = workbook["执行流程"]
            for row in range(2, sheet.max_row + 1):
                if sheet.cell(row, 1).value == "汇总核查表校验":
                    sheet.cell(row, 6, "否")
            workbook.save(config)
            workbook.close()
            errors = validate_config(config)
        self.assertIn("流程“汇总核查表校验”没有启用且“是否输出结果=是”的模块", errors)

    def test_default_flow_prefills_flow_name_as_output_name(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            audit_steps = load_flow_steps(config, "汇总核查表校验")
            summary_steps = load_flow_steps(config, "汇总校验结果说明")
        self.assertEqual("", audit_steps[0].output_name)
        self.assertEqual("汇总核查表校验", audit_steps[-1].output_name)
        self.assertEqual("汇总校验结果说明", summary_steps[-1].output_name)

    def test_initialize_config_migrates_legacy_seven_column_flow_sheet(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            workbook = load_workbook(config)
            # 构造 7 列旧配置：移除“输出文件名”“处理对象”两列。
            workbook["执行流程"].delete_cols(len(FLOW_LEGACY_HEADERS), 2)
            workbook["执行流程"].append(("我的自定义流程", 10, "检查校验区域", "是", "停止", "是", "自定义备注"))
            workbook.save(config)
            workbook.close()
            initialize_config(config)
            migrated = load_workbook(config, data_only=True)
            try:
                sheet = migrated["执行流程"]
                self.assertEqual(FLOW_HEADERS, tuple(cell.value for cell in sheet[1]))
                rows = {(row[0], row[2]): row for row in sheet.iter_rows(min_row=2, values_only=True)}
                self.assertIn(("我的自定义流程", "检查校验区域"), rows)
                self.assertEqual("汇总核查表校验", rows[("汇总核查表校验", "校验结果提取")][6])
                self.assertEqual("公式校验复制", rows[("汇总核查表校验", "校验结果提取")][7])
                self.assertEqual("输出正式本期审核结果", rows[("汇总核查表校验", "校验结果提取")][8])
                self.assertEqual("汇总校验结果说明", rows[("汇总校验结果说明", "任意行汇总")][6])
                self.assertEqual([], validate_config(config))
            finally:
                migrated.close()

    def test_initialize_config_migrates_previous_eight_column_flow_sheet(self):
        # 上一版 8 列配置（有“输出文件名”列、无“处理对象”列）也应增量迁移为 9 列。
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            workbook = load_workbook(config)
            workbook["执行流程"].delete_cols(len(FLOW_PREVIOUS_HEADERS))  # 移除“处理对象”列
            workbook.save(config)
            workbook.close()
            initialize_config(config)
            migrated = load_workbook(config, data_only=True)
            try:
                sheet = migrated["执行流程"]
                self.assertEqual(FLOW_HEADERS, tuple(cell.value for cell in sheet[1]))
                rows = {(row[0], row[2]): row for row in sheet.iter_rows(min_row=2, values_only=True)}
                self.assertEqual("公式校验复制", rows[("汇总核查表校验", "校验结果提取")][7])
                self.assertEqual("汇总校验结果说明", rows[("汇总校验结果说明", "任意行汇总")][6])
                self.assertEqual([], validate_config(config))
            finally:
                migrated.close()

    def test_initialize_config_renames_previous_nine_column_process_source_header(self):
        # 上一版 9 列配置（“处理对象”还叫“待处理对象”）应就地改名，
        # 并回填公式校验复制的默认处理对象“外部文件添加”。
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            workbook = load_workbook(config)
            workbook["执行流程"].cell(1, 8, "待处理对象")
            workbook["执行流程"].append(("我的自定义流程", 10, "检查校验区域", "是", "停止", "是", "", "", "自定义备注"))
            workbook.save(config)
            workbook.close()
            initialize_config(config)
            migrated = load_workbook(config, data_only=True)
            try:
                sheet = migrated["执行流程"]
                self.assertEqual(FLOW_HEADERS, tuple(cell.value for cell in sheet[1]))
                rows = {(row[0], row[2]): row for row in sheet.iter_rows(min_row=2, values_only=True)}
                self.assertIn(("我的自定义流程", "检查校验区域"), rows)
                self.assertEqual("外部文件添加", rows[("汇总核查表校验", "公式校验复制")][7])
                self.assertEqual("公式校验复制", rows[("汇总核查表校验", "校验结果提取")][7])
                self.assertEqual([], validate_config(config))
            finally:
                migrated.close()

    def test_legacy_seven_column_flow_sheet_loads_with_empty_output_name(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            workbook = load_workbook(config)
            workbook["执行流程"].delete_cols(len(FLOW_LEGACY_HEADERS))
            workbook.save(config)
            workbook.close()
            steps = load_flow_steps(config, "汇总核查表校验")
        self.assertTrue(steps)
        self.assertEqual("", steps[-1].output_name)

    def test_default_flow_prefills_process_source(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            audit_steps = load_flow_steps(config, "汇总核查表校验")
            summary_steps = load_flow_steps(config, "汇总校验结果说明")
        self.assertEqual("", audit_steps[0].process_source)
        self.assertEqual("外部文件添加", audit_steps[4].process_source)
        self.assertEqual("公式校验复制", audit_steps[-1].process_source)
        self.assertEqual("", summary_steps[-1].process_source)

    def test_validate_config_rejects_missing_process_source(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            workbook = load_workbook(config)
            sheet = workbook["执行流程"]
            for row in range(2, sheet.max_row + 1):
                if sheet.cell(row, 1).value == "汇总校验结果说明" and sheet.cell(row, 3).value == "任意行汇总":
                    sheet.cell(row, 8, "校验结果提取")
            workbook.save(config)
            workbook.close()
            errors = validate_config(config)
        self.assertTrue(any("处理对象" in error and "不存在" in error for error in errors))

    def test_output_prefix_uses_output_name_when_filled(self):
        self.assertEqual("汇总核查表校验", _output_prefix(60, "校验结果提取", "汇总核查表校验"))
        self.assertEqual("60_校验结果提取", _output_prefix(60, "校验结果提取", ""))
        self.assertEqual("40_任意行汇总", _output_prefix(40, "任意行汇总"))

    def test_validate_config_rejects_illegal_output_name(self):
        with TemporaryDirectory() as folder:
            config = initialize_config(Path(folder) / "config.xlsx")
            workbook = load_workbook(config)
            sheet = workbook["执行流程"]
            for row in range(2, sheet.max_row + 1):
                if sheet.cell(row, 1).value == "汇总核查表校验" and sheet.cell(row, 3).value == "校验结果提取":
                    sheet.cell(row, 7, "非法/名")
            workbook.save(config)
            workbook.close()
            errors = validate_config(config)
        self.assertTrue(any("包含非法字符" in error for error in errors))

    def test_structure_check_address_is_computed_without_excel(self):
        self.assertEqual("B3", _a1_address("B3:D5", 0, 0))
        self.assertEqual("D5", _a1_address("B3:D5", 2, 2))
        self.assertEqual("AA10", _a1_address("Z8:AA10", 2, 1))

    def _steps(self, *rows: tuple) -> list[FlowStep]:
        return [
            FlowStep(
                flow_name=row[0], order=row[1], feature_name=row[2], on_failure="停止",
                output_result=bool(row[3]), remark="", process_source=row[4] if len(row) > 4 else "",
            )
            for row in rows
        ]

    def test_resolve_process_source_empty_means_source(self):
        steps = self._steps(
            ("汇总核查表校验", 40, "外部文件添加", False),
            ("汇总核查表校验", 50, "公式校验复制", False, "外部文件添加"),
        )
        self.assertEqual((None, None), _resolve_process_source(steps, steps[0], {"外部文件添加", "公式校验复制"}))

    def test_resolve_process_source_matches_earlier_feature(self):
        steps = self._steps(
            ("汇总核查表校验", 40, "外部文件添加", False),
            ("汇总核查表校验", 50, "公式校验复制", False, "外部文件添加"),
            ("汇总核查表校验", 60, "校验结果提取", True, "公式校验复制"),
        )
        cp = {"外部文件添加", "公式校验复制"}
        self.assertEqual(("外部文件添加", None), _resolve_process_source(steps, steps[1], cp))
        self.assertEqual(("公式校验复制", None), _resolve_process_source(steps, steps[2], cp))

    def test_resolve_process_source_falls_back_to_nearest_previous(self):
        steps = self._steps(
            ("汇总核查表校验", 40, "外部文件添加", False),
            ("汇总核查表校验", 50, "公式校验复制", False, "不存在的功能"),
            ("汇总核查表校验", 60, "校验结果提取", True, "外部文件添加"),
        )
        effective, message = _resolve_process_source(steps, steps[1], {"外部文件添加", "公式校验复制"})
        self.assertEqual("外部文件添加", effective)
        self.assertIn("产副本", message)
        self.assertIn("外部文件添加", message)

    def test_resolve_process_source_self_reference_falls_back(self):
        # 引用自身：不算“排在前面的功能”，兜底到最近一个产副本功能。
        steps = self._steps(
            ("汇总核查表校验", 40, "外部文件添加", False),
            ("汇总核查表校验", 50, "公式校验复制", False),
            ("汇总核查表校验", 60, "校验结果提取", True, "校验结果提取"),
        )
        effective, _ = _resolve_process_source(steps, steps[2], {"外部文件添加", "公式校验复制"})
        self.assertEqual("公式校验复制", effective)

    def test_resolve_process_source_reference_to_non_copy_feature_falls_back(self):
        # 引用了排在前面的“表结构比对”（不产副本）也算不匹配，兜底到最近产副本功能。
        steps = self._steps(
            ("汇总核查表校验", 30, "表结构比对", False),
            ("汇总核查表校验", 40, "外部文件添加", False),
            ("汇总核查表校验", 60, "校验结果提取", True, "表结构比对"),
        )
        effective, message = _resolve_process_source(steps, steps[2], {"外部文件添加", "公式校验复制"})
        self.assertEqual("外部文件添加", effective)
        self.assertIn("产副本", message)

    def test_resolve_process_source_falls_back_to_source_when_no_copy_producer(self):
        # 前面只有检查/核对类（不产副本），兜底到源数据目录。
        steps = self._steps(
            ("汇总核查表校验", 10, "检查校验区域", False),
            ("汇总核查表校验", 30, "表结构比对", False),
            ("汇总核查表校验", 60, "校验结果提取", True, "不存在的功能"),
        )
        effective, message = _resolve_process_source(steps, steps[2], {"外部文件添加", "公式校验复制"})
        self.assertIsNone(effective)
        self.assertIn("源数据目录", message)

    def test_resolve_process_source_falls_back_to_source_when_first(self):
        steps = self._steps(
            ("检查表结构", 10, "表结构比对", True, "不存在的功能"),
        )
        effective, message = _resolve_process_source(steps, steps[0], {"外部文件添加", "公式校验复制"})
        self.assertIsNone(effective)
        self.assertIn("第一个功能", message)
        self.assertIn("源数据目录", message)
