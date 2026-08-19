from pathlib import Path
from unittest.mock import patch

from base_audit.models import CopyRange
from base_audit.region_summary import (
    _combine_header_rows,
    _header_signature,
    _metadata_plan,
    _output_filename_prefix,
    _output_sheet_name,
    _sort_extra_fields,
    _source_files,
)


def test_multirow_header_is_joined_top_to_bottom_with_underscore() -> None:
    assert _combine_header_rows([
        ["贷款", "贷款", ""],
        ["余额", "笔数", "备注"],
    ]) == ["贷款_余额", "贷款_笔数", "备注"]


def test_recursive_summary_prefers_institution_subdirectories() -> None:
    root = Path("C:/说明目录")
    root_file = root / "汇总信息.xlsx"
    nested = root / "机构A" / "报送说明.xlsx"

    with patch.object(Path, "rglob", return_value=[root_file, nested]):
        assert _source_files(root, None, recursive=True) == [nested]


def test_recursive_summary_uses_root_files_when_no_subdirectory_files() -> None:
    root = Path("C:/说明目录")
    source = root / "报送说明.xlsx"

    with patch.object(Path, "rglob", return_value=[source]):
        assert _source_files(root, None, recursive=True) == [source]


def test_non_recursive_summary_only_reads_current_folder() -> None:
    root = Path("C:/说明目录")
    root_file = root / "报送说明.xlsx"
    nested = root / "机构A" / "报送说明.xlsx"

    with patch.object(Path, "glob", return_value=[root_file]):
        assert _source_files(root, None, recursive=False) == [root_file]


def test_distinct_layout_uses_original_sheet_name() -> None:
    assert _output_sheet_name("说明汇总", ["本地校验结果"], set()) == "本地校验结果"


def test_matching_layout_uses_module_name() -> None:
    assert _output_sheet_name("说明汇总", ["本地校验结果", "在线抽验结果"], set()) == "本地校验结果"


def test_merge_headers_only_when_all_header_cells_match() -> None:
    assert _header_signature(["序号", "金额", None]) == _header_signature(["序号", "金额", ""])
    assert _header_signature(["序号", "金额"]) != _header_signature(["序号", "余额"])


def test_explanation_summary_uses_business_facing_output_name() -> None:
    assert _output_filename_prefix("汇总校验结果说明") == "校验结果与报送说明汇总"
    assert _output_filename_prefix("其他汇总流程") == "区域汇总"


def test_extra_fields_are_sorted_top_to_bottom_then_left_to_right() -> None:
    fields = [
        ("数据日期", CopyRange("说明", "F3")),
        ("机构名称", CopyRange("说明", "B3")),
        ("备注", CopyRange("说明", "A4")),
    ]
    assert [name for name, _ in _sort_extra_fields(fields)] == ["机构名称", "数据日期", "备注"]


def test_global_field_on_own_sheet_is_used_without_duplicate_output_header() -> None:
    global_fields = [("机构名称", CopyRange("任务说明", "B3")), ("数据日期", CopyRange("任务说明", "F3"))]
    output_fields, repeated = _metadata_plan(global_fields, [], ["序号", "机构名称"], "任务说明")
    assert [name for name, _ in output_fields] == ["机构名称", "数据日期"]
    assert repeated == ["机构名称"]
