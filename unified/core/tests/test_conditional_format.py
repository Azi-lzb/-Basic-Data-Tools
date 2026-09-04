from types import SimpleNamespace

from src.base_audit.excel_com import ExcelSession


class _Cell:
    def __init__(self, displayed, base):
        self.DisplayFormat = SimpleNamespace(Interior=SimpleNamespace(Color=displayed))
        self.Interior = SimpleNamespace(Color=base)


class _ExpressionCondition:
    Type = 2  # xlExpression
    Operator = 0
    Formula1 = "=AND(C5>0,C5>1)"
    Formula2 = ""


class _CellValueCondition:
    Type = 1  # xlCellValue
    Operator = 5  # xlGreater
    Formula1 = "=0"
    Formula2 = ""


def test_active_conditional_fill_accepts_any_changed_colour():
    assert ExcelSession._active_conditional_fill_color(_Cell(0x0000FF, 0xFFFFFF)) == 0x0000FF
    assert ExcelSession._active_conditional_fill_color(_Cell(0x00FF00, 0xFFFFFF)) == 0x00FF00
    assert ExcelSession._active_conditional_fill_color(_Cell(0x00FF00, 0x00FF00)) is None


def test_conditional_rule_text_expression_translates_to_destination():
    text = ExcelSession._conditional_rule_text(_ExpressionCondition(), origin="C5", destination="C26")
    assert text == "条件格式规则：AND(C26>0,C26>1)"


def test_conditional_rule_text_cell_value_includes_operator():
    text = ExcelSession._conditional_rule_text(_CellValueCondition(), origin="D2", destination="D26")
    assert text == "条件格式规则：D26>0"
