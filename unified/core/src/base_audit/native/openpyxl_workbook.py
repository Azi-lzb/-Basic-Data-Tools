"""Native workbook operations for the UOS/Linux edition.

This module deliberately has no COM, Wine, or Windows WPS dependency.  It
implements the part of the audit pipeline that is safe to do with OOXML:
named-range discovery, copying template formulas to an audit copy, importing
auxiliary sheets and reading LibreOffice-saved formula caches.
"""

from __future__ import annotations

from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from ..models import AuditRule, CopyRange, Issue, TemplateDefinition
from ..name_config import FeatureMapping, matches_named_range
from ..template import (
    REQUIRED_HEADERS,
    TemplateError,
    clean_rule_comment,
    normalize_template_name,
    parse_comment_rule_id,
    parse_formula_result,
    parse_rule_rows,
)


RESULT_KEYWORDS = ("错误", "硬性", "错", "软性")
FORMULA_ERRORS = (
    "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!",
    "#SPILL!", "#CALC!", "#FIELD!", "#GETTING_DATA",
)


def _normalise_address(address: str) -> str:
    return address.replace("$", "").strip()


def _all_defined_names(workbook: Any) -> list[Any]:
    """Return workbook- and worksheet-scoped Excel names.

    Recent openpyxl versions keep local names on ``Worksheet.defined_names``
    instead of exposing all of them from ``Workbook.defined_names``.  Excel's
    name manager displays both, and the template protocol deliberately allows
    either scope, so resolving only the latter produces a false “missing named
    range” error for valid templates.
    """
    result = list(workbook.defined_names.values())
    for sheet in workbook.worksheets:
        result.extend(sheet.defined_names.values())
    return result


def _issue_id(source_file: Path | str, rule: AuditRule, indicator: str = "") -> str:
    source_name = normalize_template_name(Path(str(source_file)).stem)
    return "｜".join((
        source_name or "工作簿未识别", rule.sheet_name or "工作表未识别",
        rule.formula_cell or "公式单元格未识别", indicator or "校验指标未识别",
    ))


def _formula_error(value: object) -> str:
    text = str(value or "").strip().upper()
    return next((item for item in FORMULA_ERRORS if text.startswith(item)), "")


def _has_result_marker(value: object) -> bool:
    return str(value or "").strip().startswith(RESULT_KEYWORDS)


def _in_range(cell: str, address: str) -> bool:
    try:
        row, column = coordinate_to_tuple(cell.replace("$", ""))
        min_col, min_row, max_col, max_row = range_boundaries(_normalise_address(address))
    except ValueError:
        return False
    return min_row <= row <= max_row and min_col <= column <= max_col


def template_formulas(template_path: Path) -> list[object]:
    """收集模板全部公式文本（外部文件规划用；每个流程只读一次）。"""
    book = load_workbook(template_path, read_only=True, data_only=False, keep_links=False)
    try:
        return [
            cell.value for sheet in book.worksheets for row in sheet.iter_rows() for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        ]
    finally:
        book.close()


def named_ranges(workbook: Any, mappings: Iterable[FeatureMapping]) -> list[CopyRange]:
    """Resolve configured Excel names to individual areas.

    Both workbook- and worksheet-scoped names are supported.  Non-range
    compatibility names such as ``_xlfn.SUMIFS`` have no destinations and are
    intentionally ignored.
    """
    mappings = tuple(mappings)
    result: list[CopyRange] = []
    seen: set[tuple[str, str]] = set()
    for defined_name in _all_defined_names(workbook):
        if not any(matches_named_range(mapping, defined_name.name) for mapping in mappings):
            continue
        try:
            destinations = tuple(defined_name.destinations)
        except (AttributeError, ValueError):
            continue
        for sheet_name, address in destinations:
            address = _normalise_address(address)
            if sheet_name not in workbook.sheetnames:
                continue
            try:
                range_boundaries(address)
            except ValueError:
                continue
            item = CopyRange(sheet_name, address)
            identity = (item.sheet_name, item.address.upper())
            if identity not in seen:
                result.append(item)
                seen.add(identity)
    return result


def require_named_ranges(workbook: Any, mapping: FeatureMapping) -> list[CopyRange]:
    missing: list[str] = []
    resolved: list[CopyRange] = []
    for range_name in mapping.range_names:
        expected = FeatureMapping(
            mapping.name, mapping.feature_type, (range_name,), mapping.workbook_limited,
            mapping.workbook_keyword, mapping.remark, mapping.output, mapping.input_note,
        )
        areas = named_ranges(workbook, (expected,))
        if areas:
            resolved.extend(areas)
        else:
            missing.append(range_name)
    if missing:
        raise TemplateError("模块“{}”缺少命名区域：{}".format(mapping.name, "、".join(missing)))
    return resolved


def named_single_cells_with_prefix(workbook: Any, prefix: str) -> list[tuple[str, CopyRange]]:
    """Return one-cell names such as ``全局单元格区域.数据日期``.

    This is used by the two row-summary modules to bring reporting metadata
    into every emitted record.  Names that do not resolve to exactly one cell
    are rejected rather than guessed.
    """
    result: list[tuple[str, CopyRange]] = []
    for defined_name in _all_defined_names(workbook):
        simple = defined_name.name.rsplit("!", 1)[-1]
        if not simple.casefold().startswith(prefix.casefold()):
            continue
        field_name = simple[len(prefix):].strip()
        if not field_name:
            continue
        try:
            destinations = tuple(defined_name.destinations)
        except (AttributeError, ValueError):
            continue
        for sheet_name, address in destinations:
            address = _normalise_address(address)
            try:
                min_col, min_row, max_col, max_row = range_boundaries(address)
            except ValueError:
                continue
            if min_col != max_col or min_row != max_row:
                raise ValueError("命名区域“{}”必须只定位一个单元格".format(simple))
            result.append((field_name, CopyRange(sheet_name, address)))
    return result


def template_structure_values(template_path: Path, definition: TemplateDefinition) -> list[tuple[CopyRange, list[list[Any]]]]:
    """Read the configured fixed labels without starting LibreOffice."""
    # 旧模板把稳定规则编号写在公式单元格批注中。ReadOnlyCell 不公开
    # comment，因此这里必须以普通模式读取模板；模板每个流程仅读取一次，
    # 批量报送文件仍使用流式/批量读取，不进入这条路径。
    book = load_workbook(template_path, read_only=False, data_only=False, keep_links=False)
    try:
        result: list[tuple[CopyRange, list[list[Any]]]] = []
        for item in definition.structure_ranges:
            if item.sheet_name not in book.sheetnames:
                raise TemplateError("模板缺少表结构工作表：{}".format(item.sheet_name))
            sheet = book[item.sheet_name]
            min_col, min_row, max_col, max_row = range_boundaries(item.address)
            result.append((item, [list(row) for row in sheet.iter_rows(
                min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col,
                values_only=True,
            )]))
        return result
    finally:
        book.close()


def read_template(
    template_path: Path,
    *,
    formula_mappings: Iterable[FeatureMapping] = (),
    structure_mappings: Iterable[FeatureMapping] = (),
    extraction_mappings: Iterable[FeatureMapping] = (),
) -> TemplateDefinition:
    """Read a structured ``审核规则`` template or named-range template."""
    formula_mappings = tuple(formula_mappings)
    structure_mappings = tuple(structure_mappings)
    extraction_mappings = tuple(extraction_mappings)
    # 旧模板把稳定规则编号写在公式单元格批注中。ReadOnlyCell 不公开
    # comment，因此这里必须以普通模式读取模板；模板每个流程仅读取一次，
    # 批量报送文件仍使用流式/批量读取，不进入这条路径。
    book = load_workbook(template_path, read_only=False, data_only=False, keep_links=False)
    try:
        structure_ranges = named_ranges(book, structure_mappings)
        extraction_ranges = named_ranges(book, extraction_mappings)
        if "审核规则" in book.sheetnames:
            sheet = book["审核规则"]
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                raise TemplateError("模板“审核规则”工作表为空")
            headers = [str(value).strip() if value is not None else "" for value in rows[0]]
            missing = [name for name in REQUIRED_HEADERS if name not in headers]
            if missing:
                raise TemplateError("审核规则表缺少字段：" + "、".join(missing))
            records = [dict(zip(headers, values)) for values in rows[1:]]
            rules = parse_rule_rows(records)
            errors = []
            for rule in rules:
                if rule.enabled and (
                    rule.sheet_name not in book.sheetnames
                    or not str(book[rule.sheet_name][rule.formula_cell].value or "").startswith("=")
                ):
                    errors.append("{}：{}!{}没有公式".format(rule.rule_id, rule.sheet_name, rule.formula_cell))
            if errors:
                raise TemplateError("；".join(errors))
            return TemplateDefinition(
                rules=rules,
                copy_ranges=[CopyRange(rule.sheet_name, rule.copy_range) for rule in rules if rule.enabled],
                structured=True,
                structure_ranges=structure_ranges,
                extraction_ranges=extraction_ranges,
            )

        copy_ranges = named_ranges(book, formula_mappings)
        # 条件格式提取、表结构比对等流程只需要模板定位区域，并不要求
        # 模板同时具备“校验区域”或任何公式。
        if formula_mappings and not copy_ranges:
            raise TemplateError("公式校验功能对应的命名区域中没有可用区域")
        rules: list[AuditRule] = []
        seen: set[tuple[str, str]] = set()
        seen_rule_ids: dict[str, tuple[str, str]] = {}
        for item in copy_ranges:
            sheet = book[item.sheet_name]
            min_col, min_row, max_col, max_row = range_boundaries(item.address)
            for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    formula = str(cell.value or "")
                    if not formula.startswith("=") or (item.sheet_name, cell.coordinate) in seen:
                        continue
                    seen.add((item.sheet_name, cell.coordinate))
                    raw_comment = str(cell.comment.text or "") if cell.comment else ""
                    explicit_rule_id = parse_comment_rule_id(raw_comment)
                    rule_id = explicit_rule_id or "{}-{}".format(item.sheet_name, cell.coordinate)
                    identity = (item.sheet_name, cell.coordinate)
                    previous = seen_rule_ids.get(rule_id)
                    if previous and previous != identity:
                        raise TemplateError(
                            "规则编号“{}”重复：{}!{}、{}!{}".format(
                                rule_id, previous[0], previous[1], item.sheet_name, cell.coordinate,
                            )
                        )
                    seen_rule_ids[rule_id] = identity
                    rules.append(AuditRule(
                        rule_id=rule_id, enabled=True,
                        report_code=normalize_template_name(template_path.stem), sheet_name=item.sheet_name,
                        formula_cell=cell.coordinate, target_cell=cell.coordinate, severity="错误",
                        message=clean_rule_comment(raw_comment),
                        copy_range=item.address,
                        result_mode="keyword" if any(marker in formula for marker in RESULT_KEYWORDS) else "error_only",
                        value_from_result=True,
                    ))
        if formula_mappings and not rules:
            raise TemplateError("公式校验功能对应的命名区域中没有公式")
        return TemplateDefinition(rules, copy_ranges, False, structure_ranges, extraction_ranges)
    finally:
        book.close()


def copy_formula_ranges(template_path: Path, audit_path: Path, definition: TemplateDefinition) -> None:
    """Copy cell formulas and visible styles to a prepared audit copy."""
    template = load_workbook(template_path, data_only=False, keep_links=False)
    audit = load_workbook(audit_path, data_only=False, keep_links=False)
    try:
        missing = sorted({rule.sheet_name for rule in definition.rules if rule.enabled} - set(audit.sheetnames))
        if missing:
            raise TemplateError("报送文件缺少工作表：" + "、".join(missing))
        copied: set[tuple[str, str]] = set()
        for item in definition.copy_ranges:
            identity = (item.sheet_name, item.address.upper())
            if identity in copied:
                continue
            copied.add(identity)
            if item.sheet_name not in audit.sheetnames:
                raise TemplateError("报送文件缺少工作表：" + item.sheet_name)
            source_sheet, target_sheet = template[item.sheet_name], audit[item.sheet_name]
            min_col, min_row, max_col, max_row = range_boundaries(item.address)
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    source = source_sheet.cell(row, column)
                    target = target_sheet.cell(row, column)
                    target.value = source.value
                    if source.has_style:
                        target._style = copy(source._style)
                    if source.number_format:
                        target.number_format = source.number_format
                    if source.alignment:
                        target.alignment = copy(source.alignment)
                    if source.protection:
                        target.protection = copy(source.protection)
                    if source.comment:
                        target.comment = copy(source.comment)
        if definition.structured:
            if "审核规则" in audit.sheetnames:
                del audit["审核规则"]
            _copy_sheet(template["审核规则"], audit, "审核规则")
        audit.calculation.fullCalcOnLoad = True
        audit.calculation.forceFullCalc = True
        audit.save(audit_path)
    finally:
        template.close()
        audit.close()


def _copy_sheet(source: Any, target_book: Any, target_name: str) -> None:
    """Cross-workbook sheet copy retaining values, basic styles, merges and hidden state."""
    target = target_book.create_sheet(target_name)
    target.sheet_state = source.sheet_state
    target.freeze_panes = source.freeze_panes
    for row in source.iter_rows():
        for cell in row:
            destination = target[cell.coordinate]
            destination.value = cell.value
            if cell.has_style:
                destination._style = copy(cell._style)
            if cell.comment:
                destination.comment = copy(cell.comment)
            if cell.hyperlink:
                destination._hyperlink = copy(cell.hyperlink)
    for key, dimension in source.row_dimensions.items():
        target.row_dimensions[key].height = dimension.height
        target.row_dimensions[key].hidden = dimension.hidden
    for key, dimension in source.column_dimensions.items():
        target.column_dimensions[key].width = dimension.width
        target.column_dimensions[key].hidden = dimension.hidden
    for merged in source.merged_cells.ranges:
        target.merge_cells(str(merged))


def add_external_sheets(audit_path: Path, external_path: Path, sheet_names: Iterable[str]) -> None:
    """Copy selected external sheets; conflicts are errors instead of overwrites."""
    audit = load_workbook(audit_path, data_only=False, keep_links=False)
    external = load_workbook(external_path, data_only=False, keep_links=False)
    try:
        names = tuple(sheet_names)
        missing = [name for name in names if name not in external.sheetnames]
        conflicts = [name for name in names if name in audit.sheetnames]
        if missing:
            raise TemplateError("外部文件缺少工作表：" + "、".join(missing))
        if conflicts:
            raise TemplateError("外部文件工作表与报送文件重名，不能覆盖原表：" + "、".join(conflicts))
        for name in names:
            _copy_sheet(external[name], audit, name)
        audit.save(audit_path)
    finally:
        audit.close()
        external.close()


def extract_issues(
    audit_path: Path,
    rules: Iterable[AuditRule],
    *, period: str, batch_id: str, org_code: str = "", org_name: str = "", source_file: Path | None = None,
    extraction_ranges: Iterable[CopyRange] = (),
    cached_values: dict[tuple[str, str], Any] | None = None,
) -> list[Issue]:
    """Build issue objects from saved caches or a live spreadsheet value map.

    ``cached_values`` is supplied by the Windows COM adapter during a batch
    run.  It avoids opening the sizeable audit XLSX again only to read a small
    set of configured formula/target cells.
    """
    source_file = source_file or audit_path
    allowed = tuple(extraction_ranges)
    enabled = [rule for rule in rules if rule.enabled and (
        not allowed or any(item.sheet_name == rule.sheet_name and _in_range(rule.formula_cell, item.address) for item in allowed)
    )]
    book = None
    try:
        audit_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        result: list[Issue] = []
        if cached_values is None:
            cached_values = _read_issue_value_blocks(audit_path, enabled)
        for rule in enabled:
            formula_result = cached_values.get((rule.sheet_name, rule.formula_cell))
            error = _formula_error(formula_result)
            if error:
                detail = (
                    "校验公式引用无效（#REF!）。请确认外部文件添加步骤已执行，且工作表名称与模板公式一致。"
                    if error == "#REF!" else "校验公式计算异常：{}，请检查模板公式及引用数据".format(error)
                )
                identity = _issue_id(source_file, rule)
                result.append(Issue(
                    issue_id=identity, period=period, batch_id=batch_id, audit_time=audit_time,
                    triggered=True, status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                    org_code=org_code, org_name=org_name, report_code=rule.report_code, sheet_name=rule.sheet_name,
                    rule_id=identity, severity="错误", formula_cell=rule.formula_cell, target_cell=rule.formula_cell,
                    target_value="", formula_result=error, message=detail, source_file=str(source_file.resolve()),
                    audit_file=str(audit_path.resolve()), detail=detail,
                ))
                continue
            if rule.result_mode == "error_only" or formula_result in (None, False, 0) or not str(formula_result).strip():
                continue
            if rule.result_mode == "keyword" and not _has_result_marker(formula_result):
                continue
            parsed = parse_formula_result(str(formula_result), default_severity=rule.severity, default_message=rule.message)
            target_value = parsed.value if rule.value_from_result else cached_values.get(
                (rule.sheet_name, rule.target_cell)
            )
            identity = _issue_id(source_file, rule, parsed.indicator)
            result.append(Issue(
                issue_id=identity, period=period, batch_id=batch_id, audit_time=audit_time,
                triggered=True, status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                org_code=org_code, org_name=org_name, report_code=rule.report_code, sheet_name=rule.sheet_name,
                rule_id=identity, severity=parsed.severity, formula_cell=rule.formula_cell, target_cell=rule.target_cell,
                target_value=target_value or "", formula_result=formula_result, message=parsed.message,
                source_file=str(source_file.resolve()), audit_file=str(audit_path.resolve()), check_field=parsed.indicator,
                comparison_value=parsed.comparison_value, reference_value=parsed.reference_value,
                difference_value=parsed.difference_value, detail=parsed.detail,
            ))
        return result
    finally:
        if book is not None:
            book.close()


def _read_issue_value_blocks(
    audit_path: Path, rules: Iterable[AuditRule],
) -> dict[tuple[str, str], Any]:
    """Read each sheet's required cells as one compact value matrix.

    ``ReadOnlyWorksheet.__getitem__`` is a streaming lookup: repeating it for
    hundreds of scattered rules repeatedly walks the worksheet XML.  Grouping
    rule cells into one rectangle per sheet keeps extraction linear in the
    actual cell area, then all rule parsing happens in normal Python memory.
    """
    cells_by_sheet: dict[str, set[str]] = {}
    for rule in rules:
        cells_by_sheet.setdefault(rule.sheet_name, set()).add(rule.formula_cell)
        if not rule.value_from_result:
            cells_by_sheet.setdefault(rule.sheet_name, set()).add(rule.target_cell)
    values: dict[tuple[str, str], Any] = {}
    book = load_workbook(audit_path, read_only=True, data_only=True, keep_links=False)
    try:
        for sheet_name, addresses in cells_by_sheet.items():
            if sheet_name not in book.sheetnames:
                raise TemplateError("审核副本缺少工作表：" + sheet_name)
            positions = {
                address: coordinate_to_tuple(address.replace("$", ""))
                for address in addresses
            }
            rows = [row for row, _ in positions.values()]
            columns = [column for _, column in positions.values()]
            first_row, last_row = min(rows), max(rows)
            first_column, last_column = min(columns), max(columns)
            matrix = [
                tuple(row) for row in book[sheet_name].iter_rows(
                    min_row=first_row, max_row=last_row,
                    min_col=first_column, max_col=last_column,
                    values_only=True,
                )
            ]
            for address, (row, column) in positions.items():
                values[(sheet_name, address)] = matrix[row - first_row][column - first_column]
    finally:
        book.close()
    return values
