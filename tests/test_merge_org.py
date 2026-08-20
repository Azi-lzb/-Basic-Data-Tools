from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from src.base_audit.merge_org import (
    _organisation_from_source,
    _merge_one_org_openpyxl,
    _output_folder,
    _safe_sheet_name,
    _selected_sources,
    _source_period,
)


def _workbook(path: Path) -> None:
    book = Workbook()
    book.save(path)
    book.close()


def test_organisation_key_is_first_filename_segment() -> None:
    path = Path("TCL科技集团财务有限公司_金融基础数据-单位贷款_B00_2026-07-31_在线核查表.xlsx")
    assert _organisation_from_source(path) == "TCL科技集团财务有限公司"
    assert _source_period(path) == "2026-07-31"


def test_sheet_prefix_is_sanitized_and_deduplicated() -> None:
    existing: set[str] = set()
    assert _safe_sheet_name("单位贷款_明细/表", existing) == "单位贷款_明细_表"
    assert _safe_sheet_name("单位贷款_明细/表", existing) == "单位贷款_明细_表_2"
    assert len(_safe_sheet_name("A" * 40, existing)) == 31


def test_empty_selection_means_all_source_files() -> None:
    with TemporaryDirectory() as folder:
        root = Path(folder)
        first = root / "机构A_表一_B00_2026-07-31_在线核查表.xlsx"
        second = root / "机构B_表二_B00_2026-07-31_在线核查表.xlsx"
        _workbook(first)
        _workbook(second)
        assert _selected_sources(root, [], recursive=True) == [first, second]


def test_output_folder_uses_requested_prefix_and_rejects_source_root() -> None:
    with TemporaryDirectory() as folder:
        root = Path(folder)
        result = _output_folder(root, root / "审核结果", "合并同机构多表")
        assert result.parent == root / "审核结果"
        assert result.name.startswith("合并同机构多表_")
        try:
            _output_folder(root, root, "合并同机构多表")
        except ValueError as exc:
            assert "不能与源数据目录相同" in str(exc)
        else:
            raise AssertionError("源数据目录应被拒绝作为输出目录")


def test_openpyxl_merge_keeps_all_sheets_and_formulas() -> None:
    with TemporaryDirectory() as folder:
        root = Path(folder)
        first = root / "机构A_个人贷款_B00_2026-07-31_在线核查表.xlsx"
        second = root / "机构A_单位贷款_B00_2026-07-31_在线核查表.xlsx"
        for path, sheet_name, hidden in ((first, "个人明细", False), (second, "单位明细", True)):
            book = Workbook()
            sheet = book.active
            sheet.title = sheet_name
            sheet["A1"] = "指标"
            sheet["B1"] = "金额"
            sheet["B2"] = "=1+1"
            sheet.merge_cells("A4:B4")
            if hidden:
                book.create_sheet("封面")
                sheet.sheet_state = "hidden"
            book.save(path)
            book.close()
        output = root / "机构A_合并_2026-07-31.xlsx"
        assert _merge_one_org_openpyxl([first, second], output) == 3
        merged = load_workbook(output, read_only=False, data_only=False)
        try:
            assert len(merged.worksheets) == 4
            assert merged["个人贷款_个人明细"]["B2"].value == "=1+1"
            assert "A4:B4" in {str(item) for item in merged["个人贷款_个人明细"].merged_cells.ranges}
            assert merged["单位贷款_单位明细"].sheet_state == "hidden"
        finally:
            merged.close()
