from types import SimpleNamespace

from src.base_audit.excel_com import ExcelSession


class _Cell:
    def __init__(self, displayed, base):
        self.DisplayFormat = SimpleNamespace(Interior=SimpleNamespace(Color=displayed))
        self.Interior = SimpleNamespace(Color=base)


def test_active_conditional_fill_accepts_any_changed_colour():
    assert ExcelSession._active_conditional_fill_color(_Cell(0x0000FF, 0xFFFFFF)) == 0x0000FF
    assert ExcelSession._active_conditional_fill_color(_Cell(0x00FF00, 0xFFFFFF)) == 0x00FF00
    assert ExcelSession._active_conditional_fill_color(_Cell(0x00FF00, 0x00FF00)) is None
