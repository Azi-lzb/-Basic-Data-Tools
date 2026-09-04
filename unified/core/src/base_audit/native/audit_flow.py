"""逐文件原生审核编排（UOS/麒麟，openpyxl + soffice 重算）。

安全边界与 Windows COM 版一致：原始报送文件只读（SHA-256 前后校验），
所有修改写入 ``_审核版.xlsx`` 副本；LibreOffice 仅在 openpyxl 完成公式/
外部表复制之后对副本做整簿重算（``soffice --headless --convert-to``）。
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

from ..discovery import detect_period, source_workbooks
from ..external import make_external_sheet_plan
from ..history import classify_current_issues, merge_history
from ..models import AuditRunResult, FileAuditResult, Issue
from ..name_config import (
    CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
    FORMULA_COPY_FUNCTION,
    ISSUE_EXTRACT_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    features_of_type,
    load_feature_mappings,
)
from ..preflight_xlsx import validate_source_xlsx
from ..template import TemplateError
from ..engines.libreoffice import LibreOfficeCalculator
from ..engines.protocol import scan_formula_errors
from .conditional_scan import extract_conditional_format_issues
from .history_io import read_history_xlsx, write_history_xlsx
from .openpyxl_workbook import (
    add_external_sheets,
    copy_formula_ranges,
    extract_issues,
    named_ranges,
    read_template,
    template_structure_values,
)


@dataclass
class NativeAuditResult:
    """与 models.AuditRunResult 同构的原生管线结果对象。"""

    batch_id: str
    period: str
    output_dir: Path
    current_issues: list[Issue] = field(default_factory=list)
    resolved_issues: list[Issue] = field(default_factory=list)
    files: list[FileAuditResult] = field(default_factory=list)
    summary_path: Path | None = None
    history_path: Path | None = None
    log_path: Path | None = None
    calculation_engine: str = ""
    performance_lines: tuple[str, ...] = ()
    copies: dict[str, Path] = field(default_factory=dict)

    @property
    def successful_files(self) -> int:
        return sum(1 for item in self.files if not item.error)

    @property
    def failed_files(self) -> int:
        return sum(1 for item in self.files if item.error)

    def summary_text(self) -> str:
        parts = [
            f"审核完成：成功 {self.successful_files} 个文件，失败 {self.failed_files} 个文件。",
            f"本期审核结果 {len(self.current_issues)} 条。",
        ]
        if self.summary_path:
            parts.append(f"问题汇总：{self.summary_path}")
        if self.history_path:
            parts.append(f"历史库：{self.history_path}")
        if self.log_path:
            parts.append(f"运行日志：{self.log_path}")
        skipped = [item.source_path.name for item in self.files if item.error.startswith("报送文件与模板不匹配：")]
        if skipped:
            parts.append(f"已跳过 {len(skipped)} 个结构不匹配文件：" + "、".join(skipped))
        if self.performance_lines:
            parts.append("耗时分解：" + "；".join(self.performance_lines))
        return "\n".join(parts)


def _template_formulas(template_path: Path) -> list[object]:
    book = load_workbook(template_path, read_only=True, data_only=False, keep_links=False)
    try:
        return [
            cell.value for sheet in book.worksheets for row in sheet.iter_rows() for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        ]
    finally:
        book.close()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _write_summary(path: Path, issues: list[Issue]) -> None:
    """问题汇总输出；超链接语义与 Windows 版一致（相对路径定位到审核副本）。"""

    def relative_target(value: str) -> str:
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
    widths = {"A": 38, "B": 24, "C": 14, "D": 16, "E": 28, "F": 18, "G": 18, "H": 18, "I": 42, "J": 38, "K": 34, "L": 24}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    book.save(path)
    book.close()


def run_native_audit(
    *,
    template_path: Path,
    output_dir: Path,
    config_path: Path | None = None,
    input_dir: Path | None = None,
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
    summary_sheet_name: str = "本期审核结果",
    update_history: bool = True,
    keep_audit_copies: bool = True,
    audit_copies_dir: Path | None = None,
    stop_on_file_failure: bool = False,
    conditional_on_failure: str = "停止",
    selected_files: list[Path] | None = None,
    write_flow_log: Callable[[object], Path | None] | None = None,
) -> NativeAuditResult:
    """跑通启用模块的原生审核；从不修改原始报送文件。"""
    say = on_step or (lambda _text: None)
    template_path = template_path.resolve()
    if not template_path.is_file():
        raise FileNotFoundError("模板文件不存在：{}".format(template_path))
    if input_dir is not None:
        input_dir = input_dir.resolve()
        sources = source_workbooks(input_dir, recursive=recursive)
    else:
        sources = []
    if selected_files is not None:
        allowed = {path.resolve() for path in sources} if sources else set()
        picked = [Path(path).resolve() for path in selected_files]
        sources = [path for path in picked if not allowed or path in allowed]
    if not sources:
        raise FileNotFoundError("没有可审核的 .xlsx 文件")
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

    # 重算引擎：每工作簿一个短命 soffice 进程（私有 profile、成功才回写）。
    calculator = LibreOfficeCalculator() if include_formula_copy else None
    engine_name = "LibreOffice Calc"
    produce_audit_copy = include_external or include_formula_copy
    if produce_audit_copy:
        audit_copies_dir.mkdir(parents=True, exist_ok=True)
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    result = NativeAuditResult(
        batch_id=batch_id, period=period, output_dir=output_dir, calculation_engine=engine_name,
    )
    source_copy_seconds = 0.0
    prepare_seconds = 0.0
    issue_extract_seconds = 0.0
    conditional_extract_seconds = 0.0

    if conditional_ranges and conditional_on_failure != "跳过":
        # 条件格式为 OOXML 规则求值，不依赖 UNO；这里只校验重算引擎可用，
        # 使“失败后处理=停止”的语义与 Windows 版保持一致。
        try:
            calculator.require_available()
        except Exception as exc:
            raise TemplateError(
                "条件格式结果提取依赖 LibreOffice 重算引擎，当前不可用：{}".format(exc)
            ) from exc

    for source in sources:
        say("正在检查：{}".format(source.name))
        if source.suffix.casefold() == ".xls":
            message = "旧版 .xls 仅支持读取和汇总；公式复制审核请先另存为 .xlsx"
            result.files.append(FileAuditResult(source, None, "", "", error=message))
            say("已跳过：{}（{}）".format(source.name, message))
            continue
        source_period = period or detect_period(source.parent).period or ""
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
                if include_external and external_path and external_plan:
                    say("正在添加外部表：{}".format(source.name))
                    add_external_sheets(audit_path, external_path, external_plan.sheet_names)
                say("正在复制校验公式：{}".format(source.name))
                copy_formula_ranges(template_path, audit_path, definition)
                say("正在通过 {} 计算：{}".format(engine_name, source.name))
                started = time.monotonic()
                calculator.recalculate(audit_path)
                prepare_seconds += time.monotonic() - started
                # 公式错误扫描（缓存值），与 Excel 版“公式错误入结果”语义对齐。
                error_cells = scan_formula_errors(audit_path, formula_ranges=definition.copy_ranges)
                for sheet_name, address, value in error_cells[:8]:
                    detail = "校验公式计算异常：{}，请检查模板公式及引用".format(value)
                    identity = "｜".join((source.stem, sheet_name, address, "公式异常"))
                    issues.append(Issue(
                        issue_id=identity, period=source_period, batch_id=batch_id,
                        audit_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        triggered=True, status="", first_seen_period="", previous_seen_period="",
                        consecutive_count=1, org_code="", org_name="",
                        report_code="", sheet_name=sheet_name, rule_id=identity,
                        severity="错误", formula_cell=address, target_cell=address,
                        target_value="", formula_result=value, message=detail,
                        source_file=str(source.resolve()), audit_file=str(audit_path.resolve()),
                        detail=detail,
                    ))
                say("计算完成：{}".format(engine_name))
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
                # 条件规则依赖模板复制进副本的公式，必须读重算后的审核副本。
                conditional_path = audit_path if include_formula_copy else source
                say("正在提取条件格式实际填充：{}".format(source.name))
                started = time.monotonic()
                extracted = extract_conditional_format_issues(
                    workbook_path=conditional_path, ranges=conditional_ranges,
                    structure_ranges=definition.structure_ranges, period=source_period,
                    batch_id=batch_id, source_file=source,
                )
                conditional_extract_seconds += time.monotonic() - started
                issues.extend(extracted)
                say("条件格式提取完成：{}，触发 {} 条".format(source.name, len(extracted)))
            retained_audit_path = audit_path if produce_audit_copy and keep_audit_copies else None
            if produce_audit_copy and not keep_audit_copies and audit_path.exists():
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
            if _sha256(source) != source_hash:
                raise RuntimeError("源报送文件哈希发生变化，已停止：{}".format(source.name)) from exc
            result.files.append(FileAuditResult(source, retained_audit_path, "", "", error=str(exc)))
            say("失败：{}：{}".format(source.name, exc))
            if stop_on_file_failure:
                raise RuntimeError("处理“{}”失败：{}".format(source.name, exc)) from exc
    # 历史说明是独立于程序包的业务文件；人工填写的“历史校验说明/审核意见”
    # 按规则身份带回到本期结果，更新时永不覆盖。
    if update_history and history_path is not None:
        old_history = read_history_xlsx(history_path)
        audited_period = period or (result.current_issues[0].period if result.current_issues else "")
        current, _resolved = classify_current_issues(audited_period, result.current_issues, old_history, batch_id=batch_id)
        result.current_issues = current
        write_history_xlsx(history_path, merge_history(old_history, current))
        result.history_path = history_path
    if write_summary:
        stamp = datetime.now().strftime("%Y%m%d%H%M%S")
        stem = Path(summary_output_name.strip()).stem if summary_output_name.strip() else "基础数据审核结果"
        result.summary_path = output_dir / "{}_{}.xlsx".format(stem, stamp)
        _write_summary(result.summary_path, result.current_issues)
    if produce_audit_copy and keep_audit_copies:
        if include_external:
            result.copies["外部文件添加"] = audit_copies_dir
        if include_formula_copy:
            result.copies["公式校验复制"] = audit_copies_dir
    if include_formula_copy:
        result.performance_lines = (
            "文件复制 {:.1f} 秒".format(source_copy_seconds),
            "副本复制与重算 {:.1f} 秒".format(prepare_seconds),
            "公式问题提取 {:.1f} 秒".format(issue_extract_seconds),
            "条件格式提取 {:.1f} 秒".format(conditional_extract_seconds),
        )
        say(
            "耗时分解：文件复制 {:.1f} 秒；副本复制与重算 {:.1f} 秒；"
            "公式问题提取 {:.1f} 秒；条件格式提取 {:.1f} 秒".format(
                source_copy_seconds, prepare_seconds, issue_extract_seconds, conditional_extract_seconds,
            )
        )
    return result
