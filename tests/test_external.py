from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from src.base_audit.external import make_external_sheet_plan, referenced_external_sheets
from src.base_audit.name_config import (
    FORMULA_COPY_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    initialize_config,
    load_feature_mappings,
)


class ExternalSheetPlanTests(unittest.TestCase):
    def test_detects_only_sheets_referenced_by_formula(self):
        names = referenced_external_sheets(
            ["=INDEX(集中系统数据!A:Z,MATCH(A1,参照表!A:A,0),2)"],
            ("集中系统数据", "参照表", "无需复制"),
        )
        self.assertEqual(("集中系统数据", "参照表"), names)

    def test_plan_uses_formula_detection(self):
        plan = make_external_sheet_plan(
            formulas=["=INDEX(集中系统数据!A:Z,MATCH(A1,参照表!A:A,0),2)"],
            available_sheets=("集中系统数据", "参照表", "无需复制"),
        )
        self.assertEqual("公式识别", plan.source)
        self.assertEqual(("集中系统数据", "参照表"), plan.sheet_names)

    def test_plan_falls_back_to_all_sheets_without_formula(self):
        plan = make_external_sheet_plan(
            formulas=["=SUM(A1:A2)"],
            available_sheets=("集中系统数据", "参照表", "无需复制"),
        )
        self.assertEqual("全部工作表", plan.source)
        self.assertEqual(("集中系统数据", "参照表", "无需复制"), plan.sheet_names)

    def test_region_mapping_does_not_depend_on_template_filename(self):
        with TemporaryDirectory() as folder:
            config_path = Path(folder) / "config.xlsx"
            initialize_config(config_path)
            workflow = config_path.with_name("流程配置.json")
            workflow.write_text(
                '{"version":1,"customModules":[{"功能名":"公式复制","执行模块":"修改_公式校验复制","命名区域名":"校验公式","输出":"审核副本","备注":""},{"功能名":"结构比对","执行模块":"核对_表结构比对","命名区域名":"固定表头","输出":"运行日志","备注":""}],"customFlows":[],"combineSheetsPlans":[{"id":"default","name":"默认","mode":"regex","pattern":"(?P<组合>.+)","groups":[]}],"activeCombineSheetsPlanId":"default"}',
                encoding="utf-8",
            )
            mappings = load_feature_mappings(config_path, Path("单位贷款.xlsx"))
        names = {mapping.name: mapping.range_names for mapping in mappings}
        self.assertEqual(("校验公式",), names["公式复制"])
        self.assertEqual(("固定表头",), names["结构比对"])
