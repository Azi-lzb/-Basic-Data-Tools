from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from src.base_audit.external import make_external_sheet_plan, referenced_external_sheets
from src.base_audit.name_config import (
    FORMULA_COPY_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
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

    def test_region_mapping_uses_named_sheet_and_template_override(self):
        with TemporaryDirectory() as folder:
            config_path = Path(folder) / "config.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "模块化功能"
            sheet.append(["功能名", "功能类型", "命名区域名", "是否限定工作簿", "工作簿关键字", "备注"])
            sheet.append(["公式复制", FORMULA_COPY_FUNCTION, "校验公式", "否", "", ""])
            sheet.append(["结构比对", STRUCTURE_COMPARE_FUNCTION, "固定表头", "否", "", ""])
            sheet.append(["单位公式复制", FORMULA_COPY_FUNCTION, "单位贷款校验", "是", "单位贷款", ""])
            workbook.create_sheet("历史审核记录")
            workbook.save(config_path)
            mappings = load_feature_mappings(config_path, Path("单位贷款.xlsx"))
        names = {mapping.name: mapping.range_names for mapping in mappings}
        self.assertEqual(("单位贷款校验",), names["单位公式复制"])
        self.assertEqual(("固定表头",), names["结构比对"])
