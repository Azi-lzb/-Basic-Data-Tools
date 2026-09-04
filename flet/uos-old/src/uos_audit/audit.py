"""Native, batch audit pipeline for UOS/Linux.

The pipeline has the same file-safety boundary as the Windows edition: source
workbooks are never modified; all workbook changes happen in ``_审核版.xlsx``
copies.  LibreOffice is used only after openpyxl has completed copying formula
ranges and external sheets.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import hashlib
import os
from pathlib import Path
from shutil import copy2
import time
from typing import Callable

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink

from .discovery import detect_period, source_workbooks
from .external import make_external_sheet_plan
from .history import classify_current_issues, merge_history, read_history_xlsx, write_history_xlsx
from .calculation import (
    ENGINE_DEFAULT,
    CalculationResult,
    FormulaEngineError,
    get_calculator,
)
from .models import FileAuditResult, Issue
from .name_config import (
    CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
    FORMULA_COPY_FUNCTION,
    ISSUE_EXTRACT_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    features_of_type,
    load_feature_mappings,
)
from .openpyxl_workbook import (
    add_external_sheets,
    copy_formula_ranges,
    extract_issues,
    read_template,
    template_structure_values,
)
from .conditional_format import (
    extract_conditional_format_issues,
    libreoffice_conditional_batch,
)
from .preflight_xlsx import validate_source_xlsx
from .template import TemplateError


@dataclass
class NativeAuditResult:
    output_dir: Path
    summary_path: Path | None
    history_path: Path | None = None
    files: list[FileAuditResult] = field(default_factory=list)
    current_issues: list[Issue] = field(default_factory=list)
    calculation_engine: str = ""
    calculation_reports: list[tuple[str, CalculationResult]] = field(default_factory=list)
    # 供配置流程后续步骤引用的、实际保留的副本目录。
    copies: dict[str, Path] = field(default_factory=dict)

    def summary_text(self) -> str:
        succeeded = sum(1 for item in self.files if not item.error)
        failed = len(self.files) - succeeded
        lines = ["统信原生审核完成：成功 {} 个文件，失败 {} 个文件。".format(succeeded, failed)]
        lines.append("本期审核结果 {} 条。".format(len(self.current_issues)))
        if self.summary_path:
            lines.append("问题汇总：{}".format(self.summary_path))
        if self.history_path:
            lines.append("历史审核说明：{}".format(self.history_path))
        return "\n".join(lines)


def _template_formulas(template_path: Path) -> list[object]:
    book = load_workbook(template_path, read_only=True, data_only=False, keep_links=False)
    try:
        return [cell.value for sheet in book.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")]
    finally:
        book.close()


def _summary_path(output_dir: Path, period: str, output_name: str = "") -> Path:
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    if output_name.strip():
        # 流程配置填的是展示名称，不要求用户再写扩展名。
        stem = Path(output_name.strip()).stem
        return output_dir / "{}_{}.xlsx".format(stem, stamp)
    suffix = period.replace("-", "_") if period else "本期"
    return output_dir / "基础数据审核结果_{}_{}.xlsx".format(suffix, stamp)


def _sha256(path: Path) -> str:
    """Return a stable source-file fingerprint without loading the workbook."""
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _write_summary(path: Path, issues: list[Issue]) -> None:
    def relative_target(value: str) -> str:
        """Use an OOXML external-link target, not a file:// URL.

        Excel/WPS frequently mishandle non-ASCII file URLs.  A path relative
        to the result workbook preserves Chinese names in OOXML and lets the
        spreadsheet application resolve it natively.
        """
        return os.path.relpath(Path(value).resolve(), start=path.parent.resolve()).replace("\\", "/")

    def attach_link(cell, workbook: str, location: str = "") -> None:
        if not workbook:
            return
        cell.hyperlink = Hyperlink(
            ref=cell.coordinate, target=relative_target(workbook), location=location,
            display=str(cell.value or ""),
        )
        cell.style = "Hyperlink"

    book = Workbook()
    sheet = book.active
    sheet.title = "本期审核结果"
    headers = ("工作簿名", "工作表名", "定位单元格", "错误类型", "校验指标", "当前值", "对比值", "差值", "描述", "规则编号", "历史校验说明", "审核意见")
    sheet.append(headers)
    for item in issues:
        sheet.append((
            Path(item.source_file).name, item.sheet_name, item.target_cell, item.severity,
            item.check_field, item.target_value, item.comparison_value, item.difference_value,
            item.detail or item.message, item.rule_id, item.institution_feedback, item.auditor_opinion,
        ))
        row = sheet.max_row
        target_book = item.audit_file if item.audit_file and Path(item.audit_file).is_file() else item.source_file
        if target_book and item.sheet_name:
            escaped_sheet = str(item.sheet_name).replace("'", "''")
            attach_link(
                sheet.cell(row, 3), target_book,
                "'{}'!{}".format(escaped_sheet, item.target_cell or "A1"),
            )
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in {"A": 38, "B": 24, "C": 14, "D": 16, "E": 28, "F": 18, "G": 18, "H": 18, "I": 42, "J": 38, "K": 34, "L": 24}.items():
        sheet.column_dimensions[column].width = width
    book.save(path)
    book.close()


def run_native_audit(
    *,
    input_dir: Path,
    template_path: Path,
    output_dir: Path,
    config_path: Path | None = None,
    external_path: Path | None = None,
    history_path: Path | None = None,
    recursive: bool = False,
    period: str = "",
    on_step: Callable[[str], None] | None = None,
    conditional_mappings=(),
    formula_mappings=None,
    structure_mappings=None,
    extraction_mappings=None,
    include_external: bool = True,
    include_formula_copy: bool = True,
    include_formula_extraction: bool = True,
    write_summary: bool = True,
    summary_output_name: str = "",
    update_history: bool = True,
    keep_audit_copies: bool = True,
    audit_copies_dir: Path | None = None,
    stop_on_file_failure: bool = False,
    selected_files: list[Path] | None = None,
    calculation_engine: str = ENGINE_DEFAULT,
) -> NativeAuditResult:
    """Run the enabled audit modules without ever changing source workbooks.

    The defaults preserve the complete audit pipeline used by the command-line
    entry point.  ``workflow.run_flow`` supplies the module switches so a
    condition-format-only flow does not create audit copies or calculate
    unrelated validation formulas.
    """
    say = on_step or (lambda _text: None)
    input_dir, template_path, output_dir = input_dir.resolve(), template_path.resolve(), output_dir.resolve()
    if not template_path.is_file():
        raise FileNotFoundError("模板文件不存在：{}".format(template_path))
    sources = source_workbooks(input_dir, recursive=recursive)
    if selected_files is not None:
        allowed = {path.resolve() for path in sources}
        sources = [
            Path(path).resolve()
            for path in selected_files
            if Path(path).resolve() in allowed
        ]
    if not sources:
        raise FileNotFoundError("源数据目录没有可审核的 .xlsx 文件：{}".format(input_dir))
    output_dir.mkdir(parents=True, exist_ok=True)
    audit_copies_dir = (audit_copies_dir or output_dir).resolve()
    config_path = config_path.resolve() if config_path else None
    mappings = load_feature_mappings(config_path, template_path)
    formula_mappings = list(formula_mappings) if formula_mappings is not None else features_of_type(mappings, FORMULA_COPY_FUNCTION)
    structure_mappings = (
        list(structure_mappings)
        if structure_mappings is not None
        else (features_of_type(mappings, STRUCTURE_COMPARE_FUNCTION) if include_formula_copy else [])
    )
    extraction_mappings = list(extraction_mappings) if extraction_mappings is not None else features_of_type(mappings, ISSUE_EXTRACT_FUNCTION)
    if not include_formula_copy:
        formula_mappings = []
    if not include_formula_extraction:
        extraction_mappings = []
    definition = read_template(
        template_path, formula_mappings=formula_mappings, structure_mappings=structure_mappings,
        extraction_mappings=extraction_mappings,
    )
    conditional_ranges = []
    if conditional_mappings:
        from .openpyxl_workbook import named_ranges
        template_book = load_workbook(template_path, read_only=True, data_only=False, keep_links=False)
        try:
            conditional_ranges = named_ranges(template_book, conditional_mappings)
        finally:
            template_book.close()
    expected_structure = template_structure_values(template_path, definition) if definition.structure_ranges else []
    external_plan = None
    if external_path and include_external:
        external = load_workbook(external_path, read_only=True, data_only=False, keep_links=False)
        try:
            external_plan = make_external_sheet_plan(
                formulas=_template_formulas(template_path), available_sheets=external.sheetnames,
            )
        finally:
            external.close()
        say("外部文件：{}，将复制 {} 个工作表（{}）".format(external_path.name, len(external_plan.sheet_names), external_plan.source))

    calculator = get_calculator(calculation_engine) if include_formula_copy else None
    produce_audit_copy = include_external or include_formula_copy
    if produce_audit_copy:
        audit_copies_dir.mkdir(parents=True, exist_ok=True)
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    result = NativeAuditResult(output_dir, None, calculation_engine=calculation_engine)
    source_copy_seconds = 0.0
    prepare_seconds = 0.0
    issue_extract_seconds = 0.0
    conditional_extract_seconds = 0.0
    com_batch_started = False
    if calculator is not None and hasattr(calculator, "begin_batch"):
        say("正在启动批量 {} 会话（模板和外部文件仅打开一次）".format(calculation_engine))
        calculator.begin_batch(
            template_path=template_path,
            external_path=external_path if include_external else None,
        )
        com_batch_started = True
    libreoffice_context = None
    com_conditional_renderer = calculator is not None and hasattr(
        calculator, "extract_rendered_conditional_formats"
    )
    if conditional_ranges and not com_conditional_renderer:
        try:
            libreoffice_context = libreoffice_conditional_batch()
            libreoffice_context.__enter__()
            say("已启动批量 LibreOffice 条件格式会话（本批仅启动一次）")
        except Exception as exc:
            raise FormulaEngineError(
                "条件格式提取必须使用 LibreOffice UNO 实际渲染，当前不可用：{}。"
                "请安装 python3-uno，并使用项目的“安装依赖Flet-UOS.sh”创建带系统 UNO 的环境。".format(exc)
            ) from exc
    try:
      for source in sources:
        say("正在检查：{}".format(source.name))
        if source.suffix.casefold() == ".xls":
            message = "旧版 .xls 仅支持读取和汇总；公式复制审核请先另存为 .xlsx"
            result.files.append(FileAuditResult(source, None, "", "", error=message))
            say("已跳过：{}（{}）".format(source.name, message))
            continue
        source_period = period or detect_period(source).period
        if expected_structure:
            match = validate_source_xlsx(source, definition, expected_structure, external_sheet_names=())
            if not match.matched:
                result.files.append(FileAuditResult(source, None, "", "", error="报送文件与模板不匹配：" + match.details))
                say("已跳过结构不匹配文件：{}".format(source.name))
                continue
        source_hash = _sha256(source)
        audit_path = audit_copies_dir / "{}_审核版.xlsx".format(source.stem)
        try:
            issues: list[Issue] = []
            if produce_audit_copy:
                started = time.monotonic()
                copy2(source, audit_path)
                source_copy_seconds += time.monotonic() - started
            if include_formula_copy:
                # Excel/WPS COM testing follows the Windows lifecycle from
                # copy through save.  Do not let openpyxl rewrite this audit
                # copy before COM has opened it.
                com_prepare = hasattr(calculator, "prepare_audit_copy")
                if include_external and external_path and external_plan:
                    say("正在添加外部表：{}".format(source.name))
                say("正在复制校验公式：{}".format(source.name))
                say("正在通过 {} 计算：{}".format(calculation_engine, source.name))
                started = time.monotonic()
                if com_prepare:
                    calculation: CalculationResult = calculator.prepare_audit_copy(
                        template_path=template_path,
                        audit_path=audit_path,
                        definition=definition,
                        external_path=external_path if include_external else None,
                        external_sheet_names=(
                            external_plan.sheet_names
                            if include_external and external_plan else ()
                        ),
                    )
                else:
                    if include_external and external_path and external_plan:
                        add_external_sheets(audit_path, external_path, external_plan.sheet_names)
                    copy_formula_ranges(template_path, audit_path, definition)
                    # 兼容性判断只针对本次复制到审核副本的校验公式区域。
                    # 工作簿中可能保留与当前流程无关的辅助公式，不应影响本流程。
                    calculation = calculator.recalculate(
                        audit_path, formula_ranges=definition.copy_ranges
                    )
                prepare_seconds += time.monotonic() - started
                result.calculation_engine = calculation.engine_name
                result.calculation_reports.append((source.name, calculation))
                if com_prepare:
                    # COM extraction below reads the configured formula cells
                    # in bulk.  Do not reopen the whole audit file merely to
                    # print an all-zero diagnostic count here.
                    say("计算完成：{}".format(calculation.engine_name))
                else:
                    say(
                        "计算诊断：{}；{}；公式错误单元格 {} 个".format(
                            calculation.engine_name,
                            calculation.capability.summary(),
                            len(calculation.error_cells),
                        )
                    )
                if calculation.diagnostic:
                    say("计算提示：" + calculation.diagnostic)
                if calculation.error_cells:
                    sample = "、".join(
                        "{}!{}={}".format(sheet_name, address, value)
                        for sheet_name, address, value in calculation.error_cells[:8]
                    )
                    say("校验区域公式错误：{}{}".format(
                        sample,
                        "；其余 {} 个".format(len(calculation.error_cells) - 8)
                        if len(calculation.error_cells) > 8 else "",
                    ))
            elif include_external and external_path and external_plan:
                say("正在添加外部表：{}".format(source.name))
                add_external_sheets(audit_path, external_path, external_plan.sheet_names)
            if include_formula_extraction:
                if not include_formula_copy:
                    raise TemplateError("公式校验结果提取需要在同一流程中先启用“公式校验复制”")
                started = time.monotonic()
                issues.extend(extract_issues(
                    audit_path, definition.rules, period=source_period, batch_id=batch_id,
                    source_file=source, extraction_ranges=definition.extraction_ranges,
                ))
                issue_extract_seconds += time.monotonic() - started
            if conditional_ranges:
                # This template's condition rules depend on formulas copied
                # from the audit template, so their rendered colours must be
                # read from the calculated audit copy.
                conditional_path = audit_path if include_formula_copy else source
                if com_conditional_renderer:
                    say("正在通过 {} 提取条件格式实际填充：{}".format(calculation_engine, source.name))
                    started = time.monotonic()
                    extracted = calculator.extract_rendered_conditional_formats(
                        workbook_path=conditional_path, ranges=conditional_ranges,
                        structure_ranges=definition.structure_ranges, period=source_period,
                        batch_id=batch_id, source_file=source,
                    )
                    conditional_extract_seconds += time.monotonic() - started
                    metrics = getattr(calculator, "last_conditional_metrics", None)
                    if metrics:
                        say(
                            "条件格式明细：打开副本 {open_seconds:.1f} 秒；区域批量读取 "
                            "{region_seconds:.1f} 秒；AppliesTo {scope_count} 段/"
                            "{scope_seconds:.1f} 秒；条件格式单元格 {candidate_count} 个；"
                            "显示色对比 {colour_seconds:.1f} 秒；空白未判定 "
                            "{skipped_blank_count} 个。".format(**metrics)
                        )
                    unsupported = 0
                elif libreoffice_context is not None:
                    say("正在通过 LibreOffice 提取条件格式实际填充：{}".format(source.name))
                    started = time.monotonic()
                    extracted = extract_conditional_format_issues(
                        workbook_path=conditional_path, ranges=conditional_ranges,
                        structure_ranges=definition.structure_ranges, period=source_period,
                        batch_id=batch_id, source_file=source,
                    )
                    conditional_extract_seconds += time.monotonic() - started
                issues.extend(extracted)
                say("条件格式实际渲染提取完成：{}，触发 {} 条".format(source.name, len(extracted)))
            retained_audit_path = audit_path if produce_audit_copy and keep_audit_copies else None
            if produce_audit_copy and not keep_audit_copies and audit_path.exists():
                # 公式复制仍在临时审核副本中完成；没有要求保留该修改步骤
                # 时，问题提取/条件格式提取完成后立即删除，避免输出目录堆积。
                audit_path.unlink()
            if _sha256(source) != source_hash:
                raise RuntimeError("源报送文件哈希发生变化，已停止：{}".format(source.name))
            result.files.append(FileAuditResult(source, retained_audit_path, "", "", issues=issues))
            result.current_issues.extend(issues)
            say("完成：{}，发现 {} 条问题".format(source.name, len(issues)))
        except Exception as exc:
            retained_audit_path = audit_path if audit_path.exists() and keep_audit_copies else None
            if audit_path.exists() and not keep_audit_copies:
                audit_path.unlink()
            # Even a failed step must never have changed the original source.
            if _sha256(source) != source_hash:
                raise RuntimeError("源报送文件哈希发生变化，已停止：{}".format(source.name)) from exc
            result.files.append(FileAuditResult(source, retained_audit_path, "", "", error=str(exc)))
            say("失败：{}：{}".format(source.name, exc))
            if stop_on_file_failure:
                raise RuntimeError("处理“{}”失败：{}".format(source.name, exc)) from exc
    finally:
      if com_batch_started:
        calculator.end_batch()
      if libreoffice_context is not None:
        libreoffice_context.__exit__(None, None, None)
    # 历史说明与 Windows 一样是独立于程序包的业务文件。只更新该工作表，
    # 人工填写的“历史校验说明/审核意见”会按规则身份带回到本期结果。
    if update_history and history_path is not None:
        old_history = read_history_xlsx(history_path)
        audited_period = period or (result.current_issues[0].period if result.current_issues else "")
        current, _resolved = classify_current_issues(audited_period, result.current_issues, old_history, batch_id=batch_id)
        result.current_issues = current
        write_history_xlsx(history_path, merge_history(old_history, current))
        result.history_path = history_path
    if write_summary:
        result.summary_path = _summary_path(output_dir, period, summary_output_name)
        _write_summary(result.summary_path, result.current_issues)
    if produce_audit_copy and keep_audit_copies:
        if include_external:
            result.copies["外部文件添加"] = audit_copies_dir
        if include_formula_copy:
            result.copies["公式校验复制"] = audit_copies_dir
    if include_formula_copy:
        say(
            "耗时分解：文件复制 {:.1f} 秒；Excel 副本复制与计算 {:.1f} 秒；"
            "公式问题提取 {:.1f} 秒；条件格式提取 {:.1f} 秒".format(
                source_copy_seconds, prepare_seconds, issue_extract_seconds, conditional_extract_seconds,
            )
        )
    return result
