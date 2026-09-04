from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.models import AuditRule
from base_audit.template import (
    TemplateError,
    clean_rule_comment,
    normalize_template_name,
    parse_comment_rule_id,
    parse_formula_result,
    parse_rule_rows,
)
from base_audit.excel_com import ExcelSession
from base_audit.name_config import (
    FORMULA_COPY_FUNCTION, STRUCTURE_COMPARE_FUNCTION, FeatureMapping, matches_named_range,
)


class TemplateTests(unittest.TestCase):
    def test_parse_valid_rule(self) -> None:
        rules = parse_rule_rows(
            [
                {
                    "规则编号": "JC-A01-001",
                    "启用": "是",
                    "报表代码": "A01",
                    "工作表": "基础数据",
                    "公式单元格": "H5",
                    "定位单元格": "B5",
                    "级别": "错误",
                    "问题说明": "余额不能为负数",
                }
            ]
        )
        self.assertEqual(rules[0].rule_id, "JC-A01-001")
        self.assertTrue(rules[0].enabled)

    def test_duplicate_rule_id_fails(self) -> None:
        row = {
            "规则编号": "JC-A01-001",
            "启用": "是",
            "报表代码": "A01",
            "工作表": "基础数据",
            "公式单元格": "H5",
            "定位单元格": "B5",
            "级别": "错误",
            "问题说明": "测试",
        }
        with self.assertRaises(TemplateError):
            parse_rule_rows([row, row])

    def test_template_date_does_not_change_identity(self) -> None:
        self.assertEqual(
            normalize_template_name("！金融基础数据-单位贷款202605"),
            normalize_template_name("！金融基础数据-单位贷款2026-06"),
        )

    def test_formula_errors_are_recognized(self) -> None:
        for value in ("#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?"):
            self.assertEqual(ExcelSession._formula_error_text(value), value)
        self.assertEqual(ExcelSession._formula_error_text("########"), "")

    def test_auto_result_markers_are_recognized(self) -> None:
        for value in (
            '="错误|余额异常"',
            '="错|余额异常"',
            '="硬性|不等于子项之和"',
            '="软性|增幅较大"',
        ):
            self.assertTrue(ExcelSession._contains_result_keyword(value))
        self.assertFalse(ExcelSession._contains_result_keyword('="核实|增幅较大"'))
        self.assertFalse(ExcelSession._contains_result_keyword('="提示|请关注"'))
        self.assertFalse(ExcelSession._contains_result_keyword("=SUM(A1:A3)"))

    def test_standard_six_part_result_is_aligned(self) -> None:
        parsed = parse_formula_result("错|单位贷款|不应该为0|0|0|0")
        self.assertEqual(parsed.severity, "错误")
        self.assertEqual(parsed.indicator, "单位贷款")
        self.assertEqual(parsed.detail, "不应该为0")
        self.assertEqual(parsed.value, "0")
        self.assertEqual(parsed.comparison_value, "0")
        self.assertEqual(parsed.difference_value, "0")

    def test_missing_tail_segments_stay_empty(self) -> None:
        parsed = parse_formula_result("错|文本1|数字1|文本2|数字2")
        self.assertEqual(parsed.severity, "错误")
        self.assertEqual(parsed.indicator, "文本1")
        self.assertEqual(parsed.detail, "数字1")
        self.assertEqual(parsed.value, "文本2")
        self.assertEqual(parsed.comparison_value, "数字2")
        self.assertEqual(parsed.difference_value, "")

    def test_all_empty_segments_keep_blank_fields(self) -> None:
        parsed = parse_formula_result("错||||")
        self.assertEqual(parsed.severity, "错误")
        self.assertEqual(parsed.indicator, "")
        self.assertEqual(parsed.detail, "")
        self.assertEqual(parsed.value, "")
        self.assertEqual(parsed.comparison_value, "")
        self.assertEqual(parsed.difference_value, "")

    def test_three_part_result_uses_third_part_as_description(self) -> None:
        parsed = parse_formula_result("软性|注册资本|超过百亿，核实是否出现放大倍数的情况")
        self.assertEqual(parsed.indicator, "注册资本")
        self.assertEqual(parsed.detail, "超过百亿，核实是否出现放大倍数的情况")
        self.assertEqual(parsed.value, "")
        self.assertEqual(parsed.comparison_value, "")
        self.assertEqual(parsed.difference_value, "")

    def test_legacy_four_part_aligned_positionally(self) -> None:
        parsed = parse_formula_result("错|字段其他|4|不应有数，请核实")
        self.assertEqual(parsed.severity, "错误")
        self.assertEqual(parsed.indicator, "字段其他")
        self.assertEqual(parsed.detail, "4")
        self.assertEqual(parsed.value, "不应有数，请核实")
        self.assertEqual(parsed.comparison_value, "")
        self.assertEqual(parsed.difference_value, "")

    def test_new_six_part_result_places_description_before_values(self) -> None:
        parsed = parse_formula_result(
            "错误|贷款余额|基础数据与参考数据不一致|305.22|334.39|-29.17"
        )
        self.assertEqual(parsed.indicator, "贷款余额")
        self.assertEqual(parsed.detail, "基础数据与参考数据不一致")
        self.assertEqual(parsed.value, "305.22")
        self.assertEqual(parsed.comparison_value, "334.39")
        self.assertEqual(parsed.difference_value, "-29.17")

    def test_full_width_pipe_is_normalised(self) -> None:
        parsed = parse_formula_result("错｜文本｜描述｜1｜2｜3")
        self.assertEqual(parsed.severity, "错误")
        self.assertEqual(parsed.indicator, "文本")
        self.assertEqual(parsed.detail, "描述")
        self.assertEqual(parsed.value, "1")
        self.assertEqual(parsed.comparison_value, "2")
        self.assertEqual(parsed.difference_value, "3")

    def test_two_part_result_indicator_is_second_segment(self) -> None:
        parsed = parse_formula_result("错误|余额不能为负数")
        self.assertEqual(parsed.severity, "错误")
        self.assertEqual(parsed.indicator, "余额不能为负数")
        self.assertEqual(parsed.detail, "")
        self.assertEqual(parsed.value, "")
        self.assertEqual(parsed.comparison_value, "")
        self.assertEqual(parsed.difference_value, "")

    def test_comment_rule_id_and_message_are_separated(self) -> None:
        comment = "规则编号：DWCK-001\n本地校验区域1\n余额不能为负数"
        self.assertEqual(parse_comment_rule_id(comment), "DWCK-001")
        self.assertEqual(clean_rule_comment(comment), "余额不能为负数")

    def test_invalid_comment_rule_id_fails(self) -> None:
        with self.assertRaises(TemplateError):
            parse_comment_rule_id("规则编号：A/B")

    def test_issue_identity_is_readable_and_tracks_formula_location(self) -> None:
        first = AuditRule(
            "DWCK-001",
            True,
            "DWCK",
            "单位存款",
            "M10",
            "B10",
            "错误",
            "",
            value_from_result=True,
        )
        moved = AuditRule(
            "DWCK-001",
            True,
            "DWCK",
            "单位存款",
            "M11",
            "B11",
            "错误",
            "",
            value_from_result=True,
        )
        first_id = ExcelSession._issue_id("001", "甲银行_单位存款_2026-06.xlsx", first)
        moved_id = ExcelSession._issue_id("001", "甲银行_单位存款_2026-06.xlsx", moved)
        self.assertEqual(first_id, "甲银行_单位存款｜单位存款｜M10｜校验指标未识别")
        self.assertNotEqual(first_id, moved_id)

    def test_named_range_prefix_is_discovered(self) -> None:
        structure = FeatureMapping("结构", STRUCTURE_COMPARE_FUNCTION, ("表结构区域",), False, "", "")
        formula = FeatureMapping("公式", FORMULA_COPY_FUNCTION, ("校验区域",), False, "", "")
        self.assertTrue(matches_named_range(structure, "表结构区域_001"))
        self.assertTrue(matches_named_range(formula, "模板.xlsx!校验区域_005"))
        self.assertFalse(matches_named_range(structure, "校验区域_001"))


if __name__ == "__main__":
    unittest.main()
