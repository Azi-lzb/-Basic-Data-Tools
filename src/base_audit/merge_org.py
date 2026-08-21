"""Merge source workbooks that belong to the same reporting institution.

The result is deliberately a standalone preparation artifact.  It does not
attempt to preserve or repair cross-workbook formulas: users create named
ranges on the merged workbook afterwards, then use the ordinary audit flow.
"""
from __future__ import annotations

import re
import shutil
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from .discovery import source_workbooks
from .excel_com import ExcelSession


_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_QUOTED_EXTERNAL_SHEET_REFERENCE = re.compile(r"'(?:[^']*\[[^\]]+\])([^']+)'!")
_BARE_EXTERNAL_SHEET_REFERENCE = re.compile(r"\[[^\]]+\]([^'!\[]+)!")


@dataclass(frozen=True)
class MergeOrgItem:
    organisation: str
    output_path: Path | None
    source_files: tuple[Path, ...]
    sheet_count: int
    status: str
    message: str = ""
    sheet_names: tuple[str, ...] = ()


@dataclass(frozen=True)
class MergeOrgResult:
    output_dir: Path
    items: tuple[MergeOrgItem, ...]
    total_files: int
    log_path: Path | None = None

    @property
    def successful_items(self) -> int:
        return sum(item.status == "成功" for item in self.items)

    @property
    def failed_items(self) -> int:
        return len(self.items) - self.successful_items

    def summary_text(self) -> str:
        text = (
            f"组合联合核查表完成：成功 {self.successful_items} 家机构，"
            f"失败 {self.failed_items} 家机构，来源文件 {self.total_files} 个。\n"
            f"合并结果目录：{self.output_dir}"
        )
        if self.log_path:
            text += f"\n运行日志：{self.log_path}"
        return text


@dataclass(frozen=True)
class TemplateMergeResult:
    """Result of the low-frequency, interactive template-composition action."""

    output_path: Path
    report_path: Path
    base_template: Path
    copied_sheets: tuple[tuple[str, str], ...]
    skipped_sheets: tuple[tuple[str, str], ...]
    formula_findings: tuple[tuple[str, str, str, str, str], ...]

    def summary_text(self) -> str:
        return (
            "联合模板制作完成："
            f"复制 {len(self.copied_sheets)} 张工作表，"
            f"跳过同名表 {len(self.skipped_sheets)} 张，"
            f"公式待核实 {len(self.formula_findings)} 处。\n"
            f"联合模板：{self.output_path}\n"
            f"检查报告：{self.report_path}"
        )


def _organisation_from_source(path: Path) -> str:
    """Institution key: the first underscore-delimited filename segment."""
    return path.stem.split("_", 1)[0].strip() or path.stem


def _source_period(path: Path) -> str:
    parts = [part.strip() for part in path.stem.split("_")]
    return parts[3] if len(parts) > 3 else ""


def _format_merged_sheet_sources(entries: list[str] | tuple[str, ...]) -> str:
    """Group ``工作簿｜工作表`` records for compact, traceable log output."""
    grouped: dict[str, list[str]] = {}
    for entry in entries:
        source, separator, sheet = str(entry).partition("｜")
        if not separator:
            source, sheet = "来源工作簿未识别", str(entry)
        grouped.setdefault(source, []).append(sheet)
    return "\n".join(
        f"{source}（{'、'.join(sheets)}）" for source, sheets in grouped.items()
    )


def _safe_sheet_name(value: str, existing: set[str]) -> str:
    """Create a valid unique Excel sheet title, retaining a numeric suffix."""
    base = _INVALID_SHEET_CHARS.sub("_", value).strip("'") or "工作表"
    base = base[:31]
    candidate = base
    index = 2
    while candidate.casefold() in existing:
        suffix = f"_{index}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1
    existing.add(candidate.casefold())
    return candidate


def _selected_sources(input_dir: Path, selected_files: list[Path] | None, recursive: bool) -> list[Path]:
    sources = source_workbooks(input_dir, recursive=recursive)
    if not selected_files:
        return sources
    allowed = {path.resolve() for path in sources}
    requested = {Path(path).resolve() for path in selected_files}
    invalid = requested - allowed
    if invalid:
        raise ValueError("选择的待合并文件不在源数据目录中：" + str(sorted(invalid)[0]))
    return [path for path in sources if path.resolve() in requested]


def _output_folder(input_dir: Path, output_dir: Path, output_name: str) -> Path:
    """Avoid placing recursive inputs in a folder that will be scanned again."""
    input_resolved = input_dir.resolve()
    output_resolved = output_dir.resolve()
    if output_resolved == input_resolved:
        raise ValueError("输出目录不能与源数据目录相同；请使用源数据目录下的“执行结果”或名称含“_skip”的其他目录")
    try:
        relative = output_resolved.relative_to(input_resolved)
    except ValueError:
        relative = None
    if relative is not None and not (
        "执行结果" in relative.parts
        or any("_skip" in part.casefold() for part in relative.parts)
    ):
        raise ValueError("输出目录位于源数据目录内，但既不是“执行结果”也不含“_skip”，递归时可能重复合并；请调整输出目录")
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    return output_resolved / f"{output_name}_{batch_id}"


def _template_report_label(path: Path) -> str:
    """Extract a concise report label from a formal template filename."""
    from .template import normalize_template_name

    label = normalize_template_name(path.stem)
    label = re.sub(r"^金融基础数据[-_]*", "", label)
    label = re.sub(r"(?:联合模板|模板)$", "", label).strip(" _.-")
    return label or "未识别报表"


def _template_merge_output(
    base_template: Path, source_templates: list[Path] | tuple[Path, ...] = ()
) -> tuple[Path, Path]:
    """Return descriptive artifact paths without ever writing to the selected base."""
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    output_dir = base_template.parent / "联合模板"
    output_dir.mkdir(parents=True, exist_ok=True)
    labels: list[str] = []
    seen: set[str] = set()
    for path in (base_template, *source_templates):
        label = _template_report_label(path)
        if label.casefold() not in seen:
            labels.append(label)
            seen.add(label.casefold())
    joined_labels = "+".join(labels)[:100]
    prefix = f"！联合模板_{joined_labels}"
    output = output_dir / f"{prefix}_{batch_id}{base_template.suffix}"
    report = output_dir / f"{prefix}_检查报告_{batch_id}.xlsx"
    return output, report


def _matrix(value: object) -> list[list[object]]:
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [[value]]
    if value and not isinstance(value[0], tuple):
        return [list(value)]
    return [list(row) for row in value]


def _column_name(column: int) -> str:
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _rebase_formula_to_local_sheets(formula: str) -> str:
    """Turn an external workbook reference into a local sheet reference.

    Every workbook-qualified reference becomes a local-sheet reference.  The
    combined template must not keep external workbook dependencies; if a sheet
    was omitted, Excel will expose that as ``#REF!`` and the final scan reports
    it.
    """
    def replace(match: re.Match[str]) -> str:
        sheet_name = match.group(1)
        return "'" + sheet_name.replace("'", "''") + "'!"

    formula = _QUOTED_EXTERNAL_SHEET_REFERENCE.sub(replace, formula)
    return _BARE_EXTERNAL_SHEET_REFERENCE.sub(replace, formula)


def _rebase_external_formula_references(
    workbook: object,
) -> int:
    """Replace only resolvable external references after all sheets are copied."""
    rewritten_count = 0
    for sheet_index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(sheet_index)
        used = sheet.UsedRange
        formulas = _matrix(used.Formula)
        for row_offset, formula_row in enumerate(formulas):
            for column_offset, formula_value in enumerate(formula_row):
                if not isinstance(formula_value, str) or not formula_value.startswith("="):
                    continue
                original = str(formula_value)
                rebased = _rebase_formula_to_local_sheets(original)
                if rebased == original:
                    continue
                row = int(used.Row) + row_offset
                column = int(used.Column) + column_offset
                sheet.Cells(row, column).Formula = rebased
                rewritten_count += 1
    return rewritten_count


def _formula_findings(workbook: object) -> list[tuple[str, str, str, str, str]]:
    """Find broken or still-external formulas in the completed template."""
    findings: list[tuple[str, str, str, str, str]] = []
    for sheet_index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(sheet_index)
        used = sheet.UsedRange
        formulas = _matrix(used.Formula)
        values = _matrix(used.Value2)
        for row_offset, formula_row in enumerate(formulas):
            for column_offset, formula_value in enumerate(formula_row):
                if not isinstance(formula_value, str) or not formula_value.startswith("="):
                    continue
                value = ""
                if row_offset < len(values) and column_offset < len(values[row_offset]):
                    value = values[row_offset][column_offset]
                address = f"{_column_name(int(used.Column) + column_offset)}{int(used.Row) + row_offset}"
                formula_text = str(formula_value)
                value_text = "" if value is None else str(value)
                if "#REF!" in formula_text or "#REF!" in value_text:
                    findings.append((sheet.Name, address, "#REF! 引用错误", formula_text, value_text))
                elif "#NAME?" in value_text:
                    findings.append((sheet.Name, address, "#NAME? 名称错误", formula_text, value_text))
                elif re.search(r"\[[^\]]+\]", formula_text):
                    findings.append((sheet.Name, address, "仍引用外部工作簿", formula_text, value_text))
    return findings


def _write_template_merge_report(
    report_path: Path,
    *,
    base_template: Path,
    base_sheets: list[str],
    sources: list[Path],
    copied: list[tuple[str, str]],
    skipped: list[tuple[str, str]],
    findings: list[tuple[str, str, str, str, str]],
) -> None:
    from openpyxl import Workbook

    book = Workbook()
    summary = book.active
    summary.title = "合并说明"
    summary.append(("项目", "内容"))
    summary.append(("底稿模板", str(base_template)))
    summary.append(("来源工作簿", "；".join(str(path) for path in sources)))
    summary.append(("复制工作表数", len(copied)))
    summary.append(("跳过同名工作表数", len(skipped)))
    summary.append(("公式待核实数", len(findings)))
    summary.append(("提示", "请将含集中系统数据、参照表等外部依赖工作表的模板选作底稿；原始底稿和来源文件均未修改。"))

    copied_sheet = book.create_sheet("工作表处理清单")
    copied_sheet.append(("来源工作簿", "工作表", "处理结果"))
    for sheet in base_sheets:
        copied_sheet.append((base_template.name, sheet, "保留为底稿"))
    for source, sheet in copied:
        copied_sheet.append((source, sheet, "已复制"))
    for source, sheet in skipped:
        copied_sheet.append((source, sheet, "同名已存在，保留底稿/先复制版本"))

    formula_sheet = book.create_sheet("公式检查")
    formula_sheet.append(("工作表", "定位单元格", "检查结果", "公式", "当前结果"))
    for item in findings:
        formula_sheet.append(item)
    if not findings:
        formula_sheet.append(("—", "—", "未发现 #REF!、#NAME? 或外部工作簿引用", "", ""))
    for sheet in book.worksheets:
        sheet.freeze_panes = "A2"
        for column in sheet.columns:
            letter = column[0].column_letter
            sheet.column_dimensions[letter].width = min(
                60, max(12, max(len(str(cell.value or "")) for cell in column) + 2)
            )
    book.save(report_path)
    book.close()


def _copy_sheet_contents(source_sheet: object, target_workbook: object) -> object:
    """Create a destination sheet and copy the source's actual used range.

    ``Worksheet.Copy`` is unreliable in some Excel COM installations: it can
    silently create a separate workbook even when ``After`` is supplied.  The
    explicit new-sheet + Range.Copy route keeps values, formulas, formatting,
    merged cells and column widths in the intended workbook.  Charts and
    workbook-level names are intentionally not carried over because the
    merged workbook is a new template-preparation artifact.
    """
    target_sheet = target_workbook.Worksheets(target_workbook.Worksheets.Count)
    copied = target_workbook.Worksheets.Add(After=target_sheet)
    used = source_sheet.UsedRange
    used.Copy(copied.Cells(used.Row, used.Column))
    # Range.Copy does not retain column widths.  Restrict the work to the
    # actual used columns rather than iterating the full worksheet.
    first_column = int(used.Column)
    last_column = first_column + int(used.Columns.Count) - 1
    for column in range(first_column, last_column + 1):
        try:
            copied.Columns(column).ColumnWidth = source_sheet.Columns(column).ColumnWidth
        except Exception:
            pass
    try:
        copied.Visible = source_sheet.Visible
    except Exception:
        pass
    return copied


def _copy_sheet_openpyxl(
    source_sheet, target_workbook, title: str, style_cache: dict[tuple[int, tuple], object]
) -> None:
    """Fast, format-preserving worksheet copy for source-workbook merging.

    Directly assigning ``_style`` across workbooks is invalid because its IDs
    point to the source workbook's style table.  Instead each distinct source
    style is registered once in the target workbook and the resulting style
    array is reused by every matching cell.  This preserves visible formatting
    without repeatedly copying font/fill/border objects cell by cell.
    """
    from openpyxl.cell.cell import Cell, MergedCell

    target = target_workbook.create_sheet(title)
    for row in source_sheet.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            copied = target.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                key = (id(source_sheet.parent), tuple(cell._style))
                target_style = style_cache.get(key)
                if target_style is None:
                    # A detached Cell safely registers copied components in the
                    # destination workbook without adding a spare visible cell.
                    prototype = Cell(target, row=1, column=1)
                    prototype.font = copy(cell.font)
                    prototype.fill = copy(cell.fill)
                    prototype.border = copy(cell.border)
                    prototype.alignment = copy(cell.alignment)
                    prototype.number_format = cell.number_format
                    prototype.protection = copy(cell.protection)
                    target_style = copy(prototype._style)
                    style_cache[key] = target_style
                copied._style = copy(target_style)
            if cell.comment is not None:
                copied.comment = copy(cell.comment)
            if cell.hyperlink is not None:
                copied._hyperlink = copy(cell.hyperlink)
    for merged_range in source_sheet.merged_cells.ranges:
        target.merge_cells(str(merged_range))
    for key, dimension in source_sheet.column_dimensions.items():
        target_dimension = target.column_dimensions[key]
        target_dimension.width = dimension.width
        target_dimension.hidden = dimension.hidden
        target_dimension.outline_level = dimension.outline_level
    for index, dimension in source_sheet.row_dimensions.items():
        target_dimension = target.row_dimensions[index]
        target_dimension.height = dimension.height
        target_dimension.hidden = dimension.hidden
        target_dimension.outline_level = dimension.outline_level
    target.freeze_panes = source_sheet.freeze_panes
    target.sheet_view.showGridLines = source_sheet.sheet_view.showGridLines
    target.sheet_properties.tabColor = source_sheet.sheet_properties.tabColor
    target.sheet_state = source_sheet.sheet_state


def _merge_one_org_openpyxl(
    files: list[Path], out_path: Path, merged_sheet_names: list[str] | None = None
) -> int:
    """Pure-Python merge used first for speed and cross-platform portability."""
    from openpyxl import Workbook, load_workbook

    merged = Workbook()
    placeholder = merged.active
    existing: set[str] = set()
    sheet_count = 0
    try:
        for source_path in sorted(files):
            source_book = load_workbook(source_path, read_only=False, data_only=False)
            try:
                style_cache: dict[tuple[int, tuple], object] = {}
                for source_sheet in source_book.worksheets:
                    if source_sheet.title.casefold() in existing:
                        continue
                    title = _safe_sheet_name(source_sheet.title, existing)
                    _copy_sheet_openpyxl(source_sheet, merged, title, style_cache)
                    if merged_sheet_names is not None:
                        # 日志保留来源，避免同一机构合并后只看到工作表名、
                        # 无法追溯该表来自哪一个报送工作簿。
                        merged_sheet_names.append(f"{source_path.name}｜{title}")
                    sheet_count += 1
            finally:
                source_book.close()
        copied_sheets = merged.worksheets[1:]
        if not copied_sheets:
            raise ValueError("来源工作簿没有可复制的工作表")
        if not any(sheet.sheet_state == "visible" for sheet in copied_sheets):
            copied_sheets[0].sheet_state = "visible"
        merged.remove(placeholder)
        merged.save(out_path)
        return sheet_count
    finally:
        merged.close()


def _merge_one_org_com(
    files: list[Path], out_path: Path, merged_sheet_names: list[str] | None = None
) -> int:
    """Compatibility fallback when a workbook cannot be handled by openpyxl."""
    source_sheet_count = 0
    with ExcelSession() as excel:
        merged = None
        try:
            merged = excel.excel.Workbooks.Add()
            while merged.Worksheets.Count > 1:
                merged.Worksheets(merged.Worksheets.Count).Delete()
            placeholder = merged.Worksheets(1)
            placeholder.Visible = -1
            existing: set[str] = set()
            for source_path in sorted(files):
                source_book = excel.open_workbook(source_path, read_only=True)
                try:
                    for index in range(1, source_book.Worksheets.Count + 1):
                        source_sheet = source_book.Worksheets(index)
                        if str(source_sheet.Name).casefold() in existing:
                            continue
                        copied = _copy_sheet_contents(source_sheet, merged)
                        copied.Name = _safe_sheet_name(source_sheet.Name, existing)
                        if merged_sheet_names is not None:
                            merged_sheet_names.append(
                                f"{source_path.name}｜{str(copied.Name)}"
                            )
                        source_sheet_count += 1
                finally:
                    excel.close_workbook(source_book)
            copied_sheets = [merged.Worksheets(i) for i in range(2, merged.Worksheets.Count + 1)]
            if not copied_sheets:
                raise ValueError("来源工作簿没有可复制的工作表")
            if not any(int(sheet.Visible) == -1 for sheet in copied_sheets):
                copied_sheets[0].Visible = -1
            placeholder.Delete()
            merged.SaveAs(str(out_path), FileFormat=51)
            return source_sheet_count
        finally:
            if merged is not None:
                excel.close_workbook(merged, save=False)


def run_template_merge(
    *,
    base_template: Path,
    source_templates: list[Path],
    on_step: Optional[Callable[[str], None]] = None,
) -> TemplateMergeResult:
    """Make a new combined template from a manually selected base template.

    This deliberately has no relationship with the ordinary source-directory
    merge.  The caller selects every file explicitly.  The base is copied at
    filesystem level first, then only source sheets whose names are absent in
    that copy are imported with Excel/WPS native range copying.  Therefore the
    base's names, styles and common dependency sheets stay authoritative.
    """
    base_template = Path(base_template).resolve()
    if not base_template.is_file():
        raise FileNotFoundError(f"底稿模板不存在：{base_template}")
    if base_template.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError("底稿模板必须是 Excel 文件")
    unique_sources: list[Path] = []
    seen = {base_template}
    for raw_path in source_templates:
        path = Path(raw_path).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"待合并工作簿不存在：{path}")
        if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            raise ValueError(f"待合并工作簿不是 Excel 文件：{path.name}")
        if path not in seen:
            unique_sources.append(path)
            seen.add(path)
    if not unique_sources:
        raise ValueError("请至少选择一个除底稿模板外的来源工作簿")

    output_path, report_path = _template_merge_output(base_template, unique_sources)
    copied: list[tuple[str, str]] = []
    skipped: list[tuple[str, str]] = []
    findings: list[tuple[str, str, str, str, str]] = []
    base_sheets: list[str] = []
    shutil.copy2(str(base_template), str(output_path))
    if on_step is not None:
        on_step("已创建底稿副本；原始模板不会修改")
        on_step("提示：请确认底稿模板已包含集中系统数据、参照表等外部依赖工作表")

    try:
        with ExcelSession() as excel:
            merged = excel.open_workbook(output_path, read_only=False)
            try:
                base_sheets = [
                    str(merged.Worksheets(index).Name)
                    for index in range(1, merged.Worksheets.Count + 1)
                ]
                existing = {
                    str(merged.Worksheets(index).Name).casefold()
                    for index in range(1, merged.Worksheets.Count + 1)
                }
                for source_path in unique_sources:
                    if on_step is not None:
                        on_step(f"正在复制工作簿：{source_path.name}")
                    source_book = excel.open_workbook(source_path, read_only=True)
                    try:
                        for index in range(1, source_book.Worksheets.Count + 1):
                            source_sheet = source_book.Worksheets(index)
                            sheet_name = str(source_sheet.Name)
                            if sheet_name.casefold() in existing:
                                skipped.append((source_path.name, sheet_name))
                                continue
                            copied_sheet = _copy_sheet_contents(source_sheet, merged)
                            copied_sheet.Name = sheet_name
                            existing.add(sheet_name.casefold())
                            copied.append((source_path.name, sheet_name))
                    finally:
                        excel.close_workbook(source_book)
                _rebase_external_formula_references(merged)
                try:
                    merged.Calculate()
                except Exception:
                    # Formula scanning still catches formula-text #REF! and
                    # remaining external links when this engine cannot recalc.
                    pass
                findings = _formula_findings(merged)
                merged.Save()
            finally:
                excel.close_workbook(merged, save=False)
    except Exception:
        # Do not leave a plausible-looking partial template after a failure.
        if output_path.exists():
            output_path.unlink()
        raise

    _write_template_merge_report(
        report_path,
        base_template=base_template,
        base_sheets=base_sheets,
        sources=unique_sources,
        copied=copied,
        skipped=skipped,
        findings=findings,
    )
    if on_step is not None:
        on_step(f"联合模板已生成：复制 {len(copied)} 张，跳过同名表 {len(skipped)} 张")
        on_step(f"公式检查完成：待核实 {len(findings)} 处")
    return TemplateMergeResult(
        output_path=output_path,
        report_path=report_path,
        base_template=base_template,
        copied_sheets=tuple(copied),
        skipped_sheets=tuple(skipped),
        formula_findings=tuple(findings),
    )


def run_merge_org(
    *,
    input_dir: Path,
    output_dir: Path,
    period: str = "",
    selected_files: list[Path] | None = None,
    flow_name: str | None = None,
    recursive: bool = True,
    on_step: Optional[Callable[[str], None]] = None,
    feature_log: object | None = None,
    output_name: str | None = None,
) -> MergeOrgResult:
    """Merge all selected workbooks by their filename's first segment."""
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()
    if not input_dir.is_dir():
        raise FileNotFoundError(f"源数据目录不存在：{input_dir}")
    sources = _selected_sources(input_dir, selected_files, recursive)
    if not sources:
        raise ValueError("源数据目录中没有可合并的 .xlsx 文件")

    output_prefix = output_name or "组合联合核查表"
    target_dir = _output_folder(input_dir, output_dir, output_prefix)
    target_dir.mkdir(parents=True, exist_ok=False)
    groups: dict[str, list[Path]] = {}
    for path in sources:
        groups.setdefault(_organisation_from_source(path), []).append(path)

    if on_step is not None:
        on_step("提示：机构识别使用文件名下划线第一段；文件名第一段必须是机构全称")
    items: list[MergeOrgItem] = []
    for organisation, files in sorted(groups.items()):
        dates = sorted({value for value in (_source_period(path) for path in files) if value})
        if len(dates) > 1 and on_step is not None:
            on_step(f"提示：机构“{organisation}”包含多个数据期（{'、'.join(dates)}），请确认")
        output_period = period.strip() or (dates[0] if dates else "未识别日期")
        out_path = target_dir / f"{organisation}_合并_{output_period}.xlsx"
        source_sheet_count = 0
        merged_sheet_names: list[str] = []
        engine = "openpyxl"
        try:
            if on_step is not None:
                on_step(f"正在合并：{organisation}（{len(files)} 个文件，openpyxl）")
            source_sheet_count = _merge_one_org_openpyxl(files, out_path, merged_sheet_names)
        except Exception as python_error:
            engine = "Excel/WPS COM"
            if out_path.exists():
                out_path.unlink()
            merged_sheet_names.clear()
            if on_step is not None:
                on_step(
                    f"提示：机构“{organisation}”无法使用 openpyxl 合并（{python_error}），"
                    "正在回退 Excel/WPS"
                )
            try:
                source_sheet_count = _merge_one_org_com(files, out_path, merged_sheet_names)
            except Exception as com_error:
                items.append(MergeOrgItem(
                    organisation, None, tuple(files), source_sheet_count, "失败",
                    f"openpyxl：{python_error}；Excel/WPS：{com_error}", tuple(merged_sheet_names),
                ))
                if on_step is not None:
                    on_step(f"合并失败：{organisation}（{com_error}），已继续其他机构")
                continue
        items.append(MergeOrgItem(
            organisation, out_path, tuple(files), source_sheet_count, "成功",
            "已合并工作表：" + _format_merged_sheet_sources(merged_sheet_names), tuple(merged_sheet_names),
        ))
        if on_step is not None:
            on_step(
                f"完成合并：{organisation}（{source_sheet_count} 张工作表）"
                f"；已合并：{_format_merged_sheet_sources(merged_sheet_names)}"
            )

    if feature_log is not None:
        feature_log.add_sheet(
            "组合联合核查表",
            ("机构", "输出文件", "来源文件数", "来源文件", "合并工作表数", "已合并工作表", "结果"),
            [
                (item.organisation, str(item.output_path or ""), len(item.source_files),
                 "\n".join(path.name for path in item.source_files), item.sheet_count,
                 _format_merged_sheet_sources(item.sheet_names),
                 item.status if item.status == "成功" else f"{item.status}：{item.message}")
                for item in items
            ],
        )
    return MergeOrgResult(target_dir, tuple(items), len(sources))
