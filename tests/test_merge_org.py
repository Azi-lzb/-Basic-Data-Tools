from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from src.base_audit.merge_org import (
    _organisation_from_source,
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
