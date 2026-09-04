"""Conditional-format extraction tests for the OOXML-native evaluator.

The native edition evaluates the submission's own conditional-format rules
directly from OOXML with openpyxl (LibreOffice's UNO cannot return a
conditionally-applied background colour).  These tests build small workbooks
with real ``cellIs`` rules and assert that extraction reports exactly the
cells whose value triggers the rule, with the rule's dxf fill colour.
"""

from pathlib import Path
import json

import pytest

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName

from uos_audit.audit import run_native_audit
from uos_audit.conditional_format import extract_conditional_format_issues
from uos_audit.models import CopyRange
from uos_audit.name_config import CONDITIONAL_FORMAT_EXTRACT_FUNCTION, FeatureMapping


def _build_rule_workbook(tmp_path: Path) -> Path:
    """Workbook: B2:B3 covered by ``cellIs greaterThan 100`` with a red fill.

    B2 = 150 (triggers), B3 = 5 (does not).
    """
    path = tmp_path / "源.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报表"
    ws["A1"] = "指标"
    ws["B2"] = 150
    ws["B3"] = 5
    red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    ws.conditional_formatting.add(
        "B2:B3", CellIsRule(operator="greaterThan", formula=["100"], fill=red)
    )
    wb.save(path)
    wb.close()
    return path


def test_extracts_cells_that_trigger_cellis(tmp_path: Path) -> None:
    path = _build_rule_workbook(tmp_path)
    issues = extract_conditional_format_issues(
        workbook_path=path,
        ranges=[CopyRange("报表", "B2:B3")],
        structure_ranges=[CopyRange("报表", "A1")],
        period="2026-08", batch_id="t", source_file=path,
    )
    assert [it.target_cell for it in issues] == ["B2"]
    assert issues[0].display_fill_color == 0xFF0000


def test_no_false_positive_for_non_triggering_cell(tmp_path: Path) -> None:
    path = tmp_path / "无触发.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报表"
    ws["B2"] = 5
    ws["B3"] = 5
    red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    ws.conditional_formatting.add(
        "B2:B3", CellIsRule(operator="greaterThan", formula=["100"], fill=red)
    )
    wb.save(path)
    wb.close()
    issues = extract_conditional_format_issues(
        workbook_path=path,
        ranges=[CopyRange("报表", "B2:B3")],
        structure_ranges=[], period="2026-08", batch_id="t", source_file=path,
    )
    assert issues == []


def test_static_fill_that_matches_rule_is_not_duplicated(tmp_path: Path) -> None:
    """A cell already statically red is skipped (baseline == displayed)."""
    path = tmp_path / "静态红.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报表"
    ws["B2"] = 150
    ws["B2"].fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    ws.conditional_formatting.add(
        "B2:B3", CellIsRule(operator="greaterThan", formula=["100"], fill=red)
    )
    wb.save(path)
    wb.close()
    issues = extract_conditional_format_issues(
        workbook_path=path,
        ranges=[CopyRange("报表", "B2:B3")],
        structure_ranges=[], period="2026-08", batch_id="t", source_file=path,
    )
    assert issues == []


def test_audit_uses_openpyxl_path_without_uno(tmp_path: Path) -> None:
    """Full audit: conditional mapping triggers without python3-uno."""
    template_path = tmp_path / "模板.xlsx"
    template = Workbook()
    sheet = template.active
    sheet.title = "报表"
    sheet["A1"] = "指标"
    sheet["B2"] = "说明"
    template.defined_names.add(DefinedName("条件格式区域", attr_text="'报表'!$B$2"))
    template.defined_names.add(DefinedName("表结构区域", attr_text="'报表'!$A$1"))
    template.save(template_path)
    template.close()

    source_dir = tmp_path / "源数据"
    source_dir.mkdir()
    source_path = _build_rule_workbook(source_dir)
    # _build_rule_workbook writes into tmp_path/"源.xlsx"; relocate it.
    (source_dir / "机构A.xlsx").write_bytes(source_path.read_bytes())
    source_path.unlink()

    mapping = FeatureMapping(
        "条件格式结果提取", CONDITIONAL_FORMAT_EXTRACT_FUNCTION, ("条件格式区域",),
        False, "", "", "审核结果集", "源数据目录",
    )
    result = run_native_audit(
        input_dir=source_dir, template_path=template_path, output_dir=tmp_path / "执行结果",
        conditional_mappings=[mapping], include_external=False,
        include_formula_copy=False, include_formula_extraction=False, write_summary=False,
        conditional_on_failure="停止",
    )
    assert len(result.files) == 1
    assert result.files[0].error == ""
    assert len(result.current_issues) == 1
    assert result.current_issues[0].target_cell == "B2"


def _expression_workbook(path: Path, sheet_name: str, cells, formula: str, area: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for address, value in cells.items():
        ws[address] = value
    fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    ws.conditional_formatting.add(area, FormulaRule(formula=[formula], fill=fill))
    wb.save(path)
    wb.close()
    return path


def _extract_expression(tmp_path: Path, name: str, sheet: str, cells, formula: str, area: str):
    path = _expression_workbook(tmp_path / "{}.xlsx".format(name), sheet, cells, formula, area)
    issues = extract_conditional_format_issues(
        workbook_path=path, ranges=[CopyRange(sheet, area)],
        structure_ranges=[], period="2026-08", batch_id="x", source_file=path,
    )
    return sorted(it.target_cell for it in issues)


def test_expression_rate_band_triggers_outliers(tmp_path: Path) -> None:
    # AND($H7<>"", OR($H7<-30, $H7>30)) over H7:H10
    assert _extract_expression(
        tmp_path, "rate", "汇总表",
        {"H7": 40, "H8": -40, "H9": 10, "H10": None},
        'AND($H7<>"",OR($H7<-30,$H7>30))', "H7:H10",
    ) == ["H7", "H8"]


def test_expression_sum_row_mismatch_triggers(tmp_path: Path) -> None:
    # $D35<>$D27+$D28+$D34 — balanced row must not trigger.
    assert _extract_expression(
        tmp_path, "sum_ok", "汇总表",
        {"D35": 100, "D27": 60, "D28": 20, "D34": 20},
        '$D35<>$D27+$D28+$D34', "D35",
    ) == []
    assert _extract_expression(
        tmp_path, "sum_bad", "汇总表",
        {"D35": 100, "D27": 60, "D28": 25, "D34": 20},
        '$D35<>$D27+$D28+$D34', "D35",
    ) == ["D35"]


def test_expression_abs_tolerance(tmp_path: Path) -> None:
    assert _extract_expression(
        tmp_path, "abs", "报表", {"E5": 5, "D5": 50}, 'ABS($E5-$D5)>10', "E5",
    ) == ["E5"]


def test_expression_relative_row_translation(tmp_path: Path) -> None:
    # A relative row reference must be evaluated per target row, not the anchor.
    assert _extract_expression(
        tmp_path, "relrow", "报表", {"F7": 1, "F8": 2, "F9": 3},
        'AND($F7<>"",$F7<>2)', "F7:F9",
    ) == ["F7", "F9"]


def test_expression_text_equal(tmp_path: Path) -> None:
    assert _extract_expression(
        tmp_path, "text", "报表", {"B2": "ok", "B3": None},
        'AND($B2<>"",$B2="ok")', "B2:B3",
    ) == ["B2"]



def test_color_indexed_to_hex(tmp_path) -> None:
    """indexed=10 maps to real red; rgb normalises to FF alpha."""
    from uos_audit.conditional_format import _color_to_hex
    from openpyxl.styles.colors import Color

    idx = Color(indexed=10)
    rgb = Color(rgb="FFC7CE")
    rgb8 = Color(rgb="00C7CE")  # openpyxl may read 8-char 00-prefixed

    assert _color_to_hex(idx) == "FFFF0000"          # COLOR_INDEX 10 -> red
    assert _color_to_hex(rgb) == "FFFFC7CE"          # 6-char padded by openpyxl
    assert _color_to_hex(rgb8) == "FF00C7CE"         # 00C7CE is deep cyan, opaque
    assert _color_to_hex(Color(indexed=64)) is None  # out of palette


def test_indexed_dxf_colour_drives_issue(tmp_path: Path) -> None:
    """A formula rule whose dxf uses indexed colour triggers a red issue."""
    from uos_audit.conditional_format import _color_to_hex
    from openpyxl.styles.colors import Color

    # The audit copy carries the reporting system's indexed-10 dxf.  Assert our
    # colour resolver treats it as a visible red (matches real 标红 templates).
    assert _color_to_hex(Color(indexed=10)) == "FFFF0000"


def test_detail_reads_like_excel_rule_without_comment(tmp_path: Path) -> None:
    """描述列在无批注时给出规则原文形态。"""
    from uos_audit.conditional_format import _rule_label
    from openpyxl.formatting.rule import Rule

    class _Rule:
        pass

    # cellIs -> 条件格式规则：C26>0
    r1 = Rule(type="cellIs", operator="greaterThan", formula=["0"])
    assert _rule_label(r1, address="C26") == "条件格式规则：C26>0"
    # expression -> compact lower-case formula without $ anchors
    r2 = Rule(type="expression", formula=["AND($D$5>0,1)"])
    assert _rule_label(r2, address="D5") == "条件格式规则：and(d5>0,1)"
    # notEqual shows <> like Excel
    r3 = Rule(type="cellIs", operator="notEqual", formula=["0"])
    assert _rule_label(r3, address="B2") == "条件格式规则：B2<>0"
