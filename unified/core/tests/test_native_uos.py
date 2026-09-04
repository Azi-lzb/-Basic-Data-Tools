"""统信 UOS 原生管线（native 包）在 Windows 上可跑的 mock 测试。

覆盖：模板读取/公式复制（openpyxl_workbook）、条件格式 OOXML 求值
（conditional_scan，复用 unified 求值器）、soffice 重算引擎（mock
subprocess）、逐文件编排（audit_flow，mock LibreOfficeCalculator）、
汇总三列（summary）、历史读写（history_io）与平台分派（engines/service）。
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

import pytest
from openpyxl import Workbook, load_workbook

from base_audit import engines, settings as settings_mod
from base_audit.engines.libreoffice import LibreOfficeCalculator
from base_audit.engines.libreoffice import (
    find_calc_engine,
)  # noqa: F401  (re-exported module surface)
from base_audit.native import summary as native_summary
from base_audit.native.conditional_scan import extract_conditional_format_issues
from base_audit.native.history_io import read_history_xlsx, write_history_xlsx
from base_audit.native.openpyxl_workbook import copy_formula_ranges, read_template


# ---------------------------------------------------------------------------
# LibreOffice 引擎（engines/libreoffice.py，移植自 flet/uos）
# ---------------------------------------------------------------------------

def test_direct_engine_respects_configured_executable(monkeypatch) -> None:
    monkeypatch.setenv("BASE_AUDIT_ENGINE", "direct")
    monkeypatch.setenv("BASE_AUDIT_SOFFICE", sys.executable)
    engine = engines.libreoffice.find_calc_engine()
    assert engine is not None
    assert engine.source == "环境变量 BASE_AUDIT_SOFFICE"


def test_linglong_engine_uses_verified_command_prefix(monkeypatch) -> None:
    monkeypatch.setenv("BASE_AUDIT_ENGINE", "ll-cli")
    monkeypatch.setattr(
        engines.libreoffice.shutil, "which",
        lambda name: "/usr/bin/ll-cli" if name == "ll-cli" else None,
    )
    monkeypatch.setattr(
        engines.libreoffice.subprocess, "run",
        lambda *args, **kwargs: SimpleNamespace(stdout="org.libreoffice.libreoffice 25.8", stderr="", returncode=0),
    )
    engine = engines.libreoffice.find_calc_engine()
    assert engine is not None
    assert engine.command_prefix == ("/usr/bin/ll-cli", "run", "org.libreoffice.libreoffice", "--", "soffice")


def test_recalculate_stages_and_replaces_only_after_success(tmp_path: Path, monkeypatch) -> None:
    workbook_path = tmp_path / "审核版.xlsx"
    book = Workbook()
    book.active["A1"] = "=SUM(2,3)"
    book.save(workbook_path)
    book.close()
    commands: list[list[str]] = []

    def fake_run(command, **kwargs):
        commands.append(command)
        output_dir = Path(command[command.index("--outdir") + 1])
        staged_arg = command[-1]
        staged = (
            Path(url2pathname(unquote(urlparse(staged_arg).path)))
            if staged_arg.startswith("file:") else Path(staged_arg)
        )
        shutil.copy2(staged, output_dir / staged.name)
        return SimpleNamespace(stdout="", stderr="", returncode=0)

    monkeypatch.setattr(engines.libreoffice.subprocess, "run", fake_run)
    result = LibreOfficeCalculator(sys.executable if False else Path(sys.executable)).recalculate(workbook_path)

    assert result.workbook_path == workbook_path
    assert "-env:UserInstallation=" in " ".join(commands[0])
    restored = load_workbook(workbook_path, data_only=False)
    assert restored.active["A1"].value == "=SUM(2,3)"
    restored.close()


# ---------------------------------------------------------------------------
# 模板读取与公式复制（native/openpyxl_workbook.py）
# ---------------------------------------------------------------------------

def _make_template(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "报表"
    sheet["D1"] = "指标"
    sheet["D2"] = "=IF(C2>0,\"错误|C2|0|余额应为零\",\"\")"
    sheet["C2"] = 5
    book.create_sheet("审核规则").append(("规则编号", "启用", "报表代码", "工作表", "公式单元格", "定位单元格", "级别", "问题说明"))
    rules = book["审核规则"]
    rules.append(("R001", "是", "", "报表", "D2", "D2", "错误", "余额核验"))
    book.save(path)
    book.close()


def test_read_template_structured_rules(tmp_path: Path) -> None:
    template = tmp_path / "模板.xlsx"
    _make_template(template)
    definition = read_template(template)
    assert definition.structured
    assert [rule.rule_id for rule in definition.rules] == ["R001"]
    assert definition.rules[0].enabled


def test_copy_formula_ranges_writes_full_calc_flag(tmp_path: Path) -> None:
    template = tmp_path / "模板.xlsx"
    _make_template(template)
    audit = tmp_path / "机构A_审核版.xlsx"
    book = Workbook()
    book.active.title = "报表"
    book.active["C2"] = 5
    book.save(audit)
    book.close()
    definition = read_template(template)
    copy_formula_ranges(template, audit, definition)
    restored = load_workbook(audit, data_only=False)
    assert str(restored["报表"]["D2"].value).startswith("=IF")
    assert restored.calculation.fullCalcOnLoad
    restored.close()


# ---------------------------------------------------------------------------
# 条件格式 OOXML 求值（native/conditional_scan.py）
# ---------------------------------------------------------------------------

def _cellis_workbook(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet["B2"] = 10
    sheet["B3"] = 0
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill
    red = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    sheet.conditional_formatting.add("B2:B3", CellIsRule(operator="greaterThan", formula=["0"], fill=red))
    book.save(path)
    book.close()


def test_extracts_cells_that_trigger_cellis(tmp_path: Path) -> None:
    workbook = tmp_path / "机构A.xlsx"
    _cellis_workbook(workbook)
    from base_audit.models import CopyRange
    issues = extract_conditional_format_issues(
        workbook_path=workbook, ranges=[CopyRange("数据", "B2:B3")],
        structure_ranges=[], period="2026-08", batch_id="B1", source_file=workbook,
    )
    assert [item.target_cell for item in issues] == ["B2"]


def test_expression_relative_translation_triggers(tmp_path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet["C2"] = 5
    sheet["D2"] = 3
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    red = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    sheet.conditional_formatting.add("C2:C2", FormulaRule(formula=["AND($D$2>0,$C2>$D$2)"], fill=red))
    path = tmp_path / "机构B.xlsx"
    book.save(path)
    book.close()
    from base_audit.models import CopyRange
    issues = extract_conditional_format_issues(
        workbook_path=path, ranges=[CopyRange("数据", "C2:C2")],
        structure_ranges=[], period="2026-08", batch_id="B1", source_file=path,
    )
    assert len(issues) == 1


# ---------------------------------------------------------------------------
# 平台分派（engines + settings + service 门面）
# ---------------------------------------------------------------------------

def test_platform_dispatch_windows_com() -> None:
    if sys.platform == "win32":
        assert engines.pipeline_kind("自动") == "com"
    else:
        assert engines.pipeline_kind("自动") == "native"


def test_platform_dispatch_simulated_linux() -> None:
    with patch.object(sys, "platform", "linux"):
        assert engines.pipeline_kind("自动") == "native"
        assert [item["value"] for item in engines.available_engines()] == ["自动", "LibreOffice Calc"]


def test_settings_accept_libreoffice_engine() -> None:
    stored = settings_mod.UserSettings(calculation_engine="LibreOffice Calc")
    assert stored.calculation_engine == "LibreOffice Calc"


def test_excel_session_rejects_non_windows() -> None:
    from base_audit.excel_com import ExcelSession, ExcelUnavailableError
    with patch.object(sys, "platform", "linux"):
        with pytest.raises(ExcelUnavailableError):
            ExcelSession("自动").__enter__()


# ---------------------------------------------------------------------------
# 汇总三列（native/summary.py）
# ---------------------------------------------------------------------------

def test_region_summary_appends_history_columns_only_for_identity_sheets(tmp_path: Path) -> None:
    from base_audit.history import (
        HISTORY_AUDIT_SHEET, LOCAL_VALIDATION_HISTORY_FIELDS, LOCAL_VALIDATION_HISTORY_SHEET,
    )
    from base_audit.name_config import FeatureMapping, USED_RANGE_SUMMARY_FUNCTION
    from openpyxl.workbook.defined_name import DefinedName

    template = tmp_path / "汇总模板.xlsx"
    book = Workbook()
    local = book.active
    local.title = "本地校验结果"
    for offset, name in enumerate(["批次", "表单名称", "规则编号", "规则类型", "规则描述", "校验字段"], start=1):
        local.cell(1, offset, name)
    book.save(template)
    book.close()

    source = tmp_path / "机构A_2026-08.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "本地校验结果"
    sheet.append(("批次", "表单名称", "规则编号", "规则类型", "规则描述", "校验字段"))
    sheet.append(("", "", "R1", "", "余额核验", "余额"))
    book.save(source)
    book.close()

    # 历史配置：一条 R1 的历史记录 → 三列应出现 1/说明/待整改
    config = tmp_path / "历史审核配置.xlsx"
    book = Workbook()
    book.remove(book.active)
    lv = book.create_sheet(LOCAL_VALIDATION_HISTORY_SHEET)
    lv.append(["来源文件", "来源工作表", "批次", "表单名称", "规则编号", "规则类型", "规则描述", "校验字段", "当前值", "对比值", "差值", *LOCAL_VALIDATION_HISTORY_FIELDS])
    lv.append(["机构A", "本地校验结果", "", "", "R1", "", "余额核验", "余额", "", "", "", 1, "利率超限", "待整改"])
    book.create_sheet(HISTORY_AUDIT_SHEET).append(("占位",))
    book.save(config)
    book.close()

    # 命名区域：汇总区域=数据区（含表头），表头区域=表头行
    book = load_workbook(template)
    book.defined_names.add(DefinedName("汇总区域", attr_text="'本地校验结果'!$A$2:$F$2"))
    book.defined_names.add(DefinedName("表头区域", attr_text="'本地校验结果'!$A$1:$F$1"))
    book.save(template)
    book.close()

    feature = FeatureMapping("本地校验汇总", USED_RANGE_SUMMARY_FUNCTION, ("汇总区域",), False, "", "")
    path = native_summary.run_region_summaries(
        template_path=template, input_dir=tmp_path, output_dir=tmp_path / "汇总输出",
        features=[feature], recursive=False, output_name="汇总测试",
        history_config_path=config,
    )
    out = load_workbook(path)
    local_sheet = out["本地校验结果"]
    headers = [cell.value for cell in local_sheet[1]]
    assert headers[-3:] == list(LOCAL_VALIDATION_HISTORY_FIELDS)
    last_row = [cell.value for cell in local_sheet[2]]
    assert last_row[0] == "机构A"  # 来源文件去日期
    assert last_row[4] == "R1"    # 0来源文件 1来源工作表 2批次 3表单名称 4规则编号
    assert last_row[-3:] == ["1", "利率超限", "待整改"]  # join 后为字符串
    out.close()
