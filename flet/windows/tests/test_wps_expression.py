"""WPS conditional-format fallback: expression rule evaluation.

The Windows WPS adapter cannot always read a rendered ``DisplayFormat`` colour
(WPS 12.0 returns ``None``).  In that case it evaluates the submission's own
OOXML ``cellIs`` and ``expression`` rules directly.  These tests build small
workbooks around ``AND(C5>0,C5>1)`` and assert the relative/absolute reference
translation and trigger results for C5, D5, E5 and C26, plus the "don't guess"
behaviour for unparseable rules.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill

from uos_audit.calculation import WpsComCalculator
from uos_audit.conditional_format import evaluate_expression_formula
from uos_audit.models import CopyRange


def test_expression_and_relative_reference_conversion() -> None:
    """AND(C5>0,C5>1) reads each target cell, not a fixed C5.

    C5/D5/E5 share the anchor row and move one column at a time; C26 moves down
    21 rows from the C5 anchor.  Setting D5=0 and E5=0.5 makes any accidental
    "fixed on C5" evaluation report True where the real rule reports False.
    """
    wb = Workbook()
    ws = wb.active
    ws["C5"] = 2     # >1  -> True
    ws["D5"] = 0     # not >0 -> False (would be True if wrongly fixed on C5)
    ws["E5"] = 0.5   # >0 but not >1 -> False
    ws["C26"] = 3    # >1 -> True (proves the +21 row translation)
    formula = "AND(C5>0,C5>1)"

    results = {
        "C5": evaluate_expression_formula(formula, ws, 5, 3, 5, 3),
        "D5": evaluate_expression_formula(formula, ws, 5, 3, 5, 4),
        "E5": evaluate_expression_formula(formula, ws, 5, 3, 5, 5),
        "C26": evaluate_expression_formula(formula, ws, 5, 3, 26, 3),
    }
    wb.close()
    assert results == {"C5": True, "D5": False, "E5": False, "C26": True}


def _expression_workbook(path: Path, cells: dict[str, object], formula: str, area: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "报表"
    for address, value in cells.items():
        ws[address] = value
    fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    ws.conditional_formatting.add(area, FormulaRule(formula=[formula], fill=fill))
    wb.save(path)
    wb.close()
    return path


def test_wps_fallback_expression_relative_description_and_trigger(tmp_path: Path) -> None:
    """The fallback triggers C5/D5/E5/C26 and writes a destination-relative rule.

    Each issue's 描述 must translate ``AND(C5>0,C5>1)`` to the triggered cell
    (``AND(C26>0,C26>1)`` for C26) rather than keep the anchor's ``C5``.
    """
    path = _expression_workbook(
        tmp_path / "表达式.xlsx",
        {"C5": 2, "D5": 3, "E5": 4, "C26": 5},
        "AND(C5>0,C5>1)", "C5:E26",
    )
    issues = WpsComCalculator()._extract_rule_based_conditional_formats(
        workbook_path=path, ranges=[CopyRange("报表", "C5:E26")], structure_ranges=[],
        period="2026-08", batch_id="t", source_file=path,
    )
    by_cell = {item.target_cell: item.detail for item in issues}
    assert set(by_cell) == {"C5", "D5", "E5", "C26"}
    assert by_cell["C5"] == "条件格式规则：AND(C5>0,C5>1)"
    assert by_cell["D5"] == "条件格式规则：AND(D5>0,D5>1)"
    assert by_cell["E5"] == "条件格式规则：AND(E5>0,E5>1)"
    assert by_cell["C26"] == "条件格式规则：AND(C26>0,C26>1)"


def test_wps_fallback_unsupported_rule_does_not_trigger(tmp_path: Path) -> None:
    """A rule that cannot be parsed must not trigger or be guessed.

    ``INDIRECT`` is outside the supported expression subset, so no issue is
    reported and the run-log counter records the unsupported rule.
    """
    path = _expression_workbook(
        tmp_path / "不支持.xlsx",
        {"C5": 1, "D5": 2},
        'AND(C5>0,INDIRECT("A1"))', "C5:D5",
    )
    calculator = WpsComCalculator()
    issues = calculator._extract_rule_based_conditional_formats(
        workbook_path=path, ranges=[CopyRange("报表", "C5:D5")], structure_ranges=[],
        period="2026-08", batch_id="t", source_file=path,
    )
    assert issues == []
    assert calculator.last_conditional_metrics["unsupported_count"] == 1


def test_wps_fallback_keeps_cellis_handling(tmp_path: Path) -> None:
    """Regression: the existing ``cellIs`` path still triggers by threshold."""
    path = tmp_path / "cellIs.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "报表"
    ws["C5"] = 2
    ws["C6"] = 0
    fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    ws.conditional_formatting.add(
        "C5:C6", CellIsRule(operator="greaterThan", formula=["1"], fill=fill)
    )
    wb.save(path)
    wb.close()

    issues = WpsComCalculator()._extract_rule_based_conditional_formats(
        workbook_path=path, ranges=[CopyRange("报表", "C5:C6")], structure_ranges=[],
        period="2026-08", batch_id="t", source_file=path,
    )
    assert [item.target_cell for item in issues] == ["C5"]
    assert issues[0].detail == "条件格式规则：C5>1"
