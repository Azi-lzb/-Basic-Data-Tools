from __future__ import annotations

import time
import re
from pathlib import Path
from typing import Any, Iterable

from .history import HISTORY_AUDIT_SHEET, HISTORY_HEADERS
from .name_config import (
    CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
    FORMULA_COPY_FUNCTION,
    FeatureMapping,
    ISSUE_EXTRACT_FUNCTION,
    NAMED_RANGE_CHECK_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    features_of_type,
    load_feature_mappings,
    matches_named_range,
)
from .models import (
    AuditRule,
    CopyRange,
    Issue,
    PreflightItem,
    SourceMatch,
    StructureCheck,
    TemplateDefinition,
)
from .preflight_xlsx import (
    STRUCTURE_MATCH_THRESHOLD,
    _display_value,
    _is_blank,
    _structure_values_equal,
)
from .template import (
    DATE_TOKEN_RE,
    REQUIRED_HEADERS,
    TemplateError,
    clean_rule_comment,
    normalize_template_name,
    parse_comment_rule_id,
    parse_formula_result,
    parse_rule_rows,
)
from .conditional_format import evaluate_expression_formula


# Legacy templates do not use one uniform flag. Keep discovery and extraction
# on the same marker list so a valid result formula cannot silently disappear.
# 自动扫描命名区域时只接受统一的结果起始标识：前三类为硬性错误，
# “软性”表示需人工确认。避免“请核实”等中间说明被误提取。
RESULT_KEYWORDS = ("错误", "硬性", "错", "软性")

XL_OPEN_XML_WORKBOOK = 51
XL_CALCULATION_DONE = 0
XL_CALCULATION_MANUAL = -4135


def _matrix(value: Any) -> list[list[Any]]:
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [[value]]
    if value and not isinstance(value[0], tuple):
        return [list(value)]
    return [list(row) for row in value]


def _range_matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
    """Normalise a COM Range.Value2 result without losing one-row/column data."""
    if rows <= 0 or columns <= 0:
        return []
    if rows == 1 and columns == 1:
        return [[value]]
    if rows == 1:
        if isinstance(value, tuple) and len(value) == 1 and isinstance(value[0], tuple):
            return [list(value[0])]
        return [list(value) if isinstance(value, tuple) else [value]]
    if columns == 1:
        if not isinstance(value, tuple):
            return [[value]]
        return [[row[0] if isinstance(row, tuple) else row] for row in value]
    return _matrix(value)


def _plain(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _cell_in_address(cell: str, address: str) -> bool:
    """Return whether A1 cell belongs to a simple A1:A2 named-range area."""
    def parts(value: str) -> tuple[int, int]:
        match = re.fullmatch(r"([A-Z]{1,3})(\d+)", value.replace("$", "").upper())
        if not match:
            return -1, -1
        letters, row = match.groups()
        column = 0
        for letter in letters:
            column = column * 26 + ord(letter) - 64
        return column, int(row)

    start, _, end = address.replace("$", "").split(":", 1)[0], ":", address.replace("$", "").split(":", 1)[-1]
    column, row = parts(cell)
    start_column, start_row = parts(start)
    end_column, end_row = parts(end)
    return start_column <= column <= end_column and start_row <= row <= end_row


def _a1_address(start_address: str, row_offset: int, column_offset: int) -> str:
    """Return an A1 address without another COM call for Cell.Address."""
    match = re.fullmatch(r"([A-Z]{1,3})(\d+)(?::[A-Z]{1,3}\d+)?", start_address.upper())
    if not match:
        return start_address
    letters, row_text = match.groups()
    column = 0
    for letter in letters:
        column = column * 26 + ord(letter) - 64
    column += column_offset
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return f"{result}{int(row_text) + row_offset}"


def _cell_position(address: str) -> tuple[int, int]:
    """Return (row, column) for one A1 address without a COM Cells call."""
    match = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", address.strip().upper())
    if not match:
        raise TemplateError(f"单元格地址无效：{address}")
    letters, row_text = match.groups()
    column = 0
    for letter in letters:
        column = column * 26 + ord(letter) - 64
    return int(row_text), column


def _column_letters(column: int) -> str:
    text = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        text = chr(65 + remainder) + text
    return text


def _rectangle_address(
    first_row: int, first_column: int, last_row: int, last_column: int
) -> str:
    return (
        f"{_column_letters(first_column)}{first_row}:"
        f"{_column_letters(last_column)}{last_row}"
    )


class ExcelUnavailableError(RuntimeError):
    pass


from .com_message_filter import (
    register_com_message_filter,
    revoke_com_message_filter,
)


class ExcelSession:
    """Owns an isolated, hidden Excel process."""

    def __init__(self, engine_preference: str = "自动") -> None:
        self.excel = None
        self._pythoncom = None
        self.engine_name = ""
        # 累计本次会话各类 COM 重操作耗时。由 service 写入运行日志，
        # 便于在真实模板上定位瓶颈，而不是凭感觉反复调整计算策略。
        self.copy_formula_seconds = 0.0
        self.calculate_seconds = 0.0
        self.external_copy_seconds = 0.0
        self.engine_preference = engine_preference if engine_preference in {"自动", "Microsoft Excel", "WPS 表格"} else "自动"

    def __enter__(self) -> "ExcelSession":
        import sys

        if sys.platform != "win32":
            raise ExcelUnavailableError(
                "当前操作系统不支持 Excel/WPS COM。统信 UOS / 麒麟版需使用 "
                "LibreOffice Calc/UNO 计算引擎（适配进行中），Windows 版请使用 "
                "Flask/windows 或 pywebview2/windows 发行包。"
            )
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ExcelUnavailableError(
                "缺少 pywin32，请先执行 pip install -r packaging/requirements/runtime-windows.txt"
            ) from exc

        self._pythoncom = pythoncom
        pythoncom.CoInitialize()
        # Win7/慢速机器上 Excel 重算或保存时忙，COM 会以
        # RPC_E_CALL_REJECTED（“被呼叫方拒绝接收呼叫”）拒绝呼入；
        # 注册 IMessageFilter 让被拒呼叫自动挂起重试，而不是直接失败。
        self._message_filter_registered = register_com_message_filter()
        last_error: Exception | None = None
        candidates = {
            "自动": (
                ("Excel.Application", "Microsoft Excel"),
                ("ket.Application", "WPS 表格"),
                ("KET.Application", "WPS 表格"),
            ),
            "Microsoft Excel": (("Excel.Application", "Microsoft Excel"),),
            "WPS 表格": (("ket.Application", "WPS 表格"), ("KET.Application", "WPS 表格")),
        }[self.engine_preference]
        for progid, engine_name in candidates:
            try:
                self.excel = win32com.client.DispatchEx(progid)
                self.engine_name = engine_name
                break
            except Exception as exc:
                last_error = exc
        if self.excel is None:
            pythoncom.CoUninitialize()
            self._pythoncom = None
            raise ExcelUnavailableError(
                f"无法启动{self.engine_preference if self.engine_preference != '自动' else ' Microsoft Excel 或 WPS 表格'}，请确认已安装表格软件并能正常打开工作簿"
            ) from last_error

        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        self.excel.EnableEvents = False
        self.excel.AskToUpdateLinks = False
        # 某些含数据模型/OLAP 连接的 Excel 实例禁止通过 COM 修改
        # Application.Calculation。该优化失败时必须降级，不能影响审核。
        try:
            self.excel.Calculation = XL_CALCULATION_MANUAL
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.excel is not None:
            try:
                while self.excel.Workbooks.Count:
                    self.excel.Workbooks(1).Close(False)
            except Exception:
                pass
            try:
                self.excel.Quit()
            except Exception:
                pass
        self.excel = None
        if getattr(self, "_message_filter_registered", False):
            revoke_com_message_filter()
            self._message_filter_registered = False
        if self._pythoncom is not None:
            self._pythoncom.CoUninitialize()
        self._pythoncom = None

    def open_workbook(self, path: Path, *, read_only: bool) -> Any:
        try:
            return self.excel.Workbooks.Open(
                str(path.resolve()),
                UpdateLinks=0,
                ReadOnly=read_only,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
        except Exception:
            # WPS 的 COM Open 参数集与 Excel 略有差异；使用双方共有的
            # 最小参数重试，避免因不支持 AddToMru 等可选项而无法打开文件。
            return self.excel.Workbooks.Open(
                str(path.resolve()), UpdateLinks=0, ReadOnly=read_only
            )

    @staticmethod
    def close_workbook(workbook: Any, *, save: bool = False) -> None:
        workbook.Close(SaveChanges=save)

    def read_template(
        self, template_workbook: Any, *, config_path: Path | None = None,
        features: list[FeatureMapping] | None = None,
    ) -> TemplateDefinition:
        features = features if features is not None else load_feature_mappings(config_path, Path(str(template_workbook.FullName)))
        try:
            sheet = template_workbook.Worksheets("审核规则")
        except Exception:
            return self._discover_named_template(template_workbook, features=features)

        rows = _matrix(sheet.UsedRange.Value2)
        if not rows:
            raise TemplateError("模板“审核规则”工作表为空")
        headers = [str(item).strip() if item is not None else "" for item in rows[0]]
        missing = [name for name in REQUIRED_HEADERS if name not in headers]
        if missing:
            raise TemplateError("审核规则表缺少字段：" + "、".join(missing))

        records = []
        for values in rows[1:]:
            records.append(
                {
                    header: values[index] if index < len(values) else None
                    for index, header in enumerate(headers)
                }
            )
        rules = parse_rule_rows(records)

        sheet_names = {
            template_workbook.Worksheets(index).Name
            for index in range(1, template_workbook.Worksheets.Count + 1)
        }
        errors: list[str] = []
        for rule in rules:
            if not rule.enabled:
                continue
            if rule.sheet_name not in sheet_names:
                errors.append(f"{rule.rule_id}：模板缺少工作表“{rule.sheet_name}”")
                continue
            formula = template_workbook.Worksheets(rule.sheet_name).Range(
                rule.formula_cell
            ).Formula
            if not isinstance(formula, str) or not formula.startswith("="):
                errors.append(
                    f"{rule.rule_id}：{rule.sheet_name}!{rule.formula_cell}没有公式"
                )
        if errors:
            raise TemplateError("；".join(errors))
        return TemplateDefinition(
            rules=rules,
            copy_ranges=[
                CopyRange(rule.sheet_name, rule.copy_range)
                for rule in rules
                if rule.enabled
            ],
            structured=True,
            structure_ranges=self._named_ranges_for_features(
                template_workbook, features_of_type(features, STRUCTURE_COMPARE_FUNCTION)
            ),
            extraction_ranges=self._named_ranges_for_features(
                template_workbook, features_of_type(features, ISSUE_EXTRACT_FUNCTION)
            ),
        )

    def _discover_named_template(
        self, workbook: Any, *, features: list[FeatureMapping]
    ) -> TemplateDefinition:
        copy_ranges = self._named_ranges_for_features(
            workbook, features_of_type(features, FORMULA_COPY_FUNCTION)
        )
        if copy_ranges:
            return self._definition_from_copy_ranges(
                workbook,
                copy_ranges,
                structure_ranges=self._named_ranges_for_features(
                    workbook, features_of_type(features, STRUCTURE_COMPARE_FUNCTION)
                ),
                extraction_ranges=self._named_ranges_for_features(
                    workbook, features_of_type(features, ISSUE_EXTRACT_FUNCTION)
                ),
            )
        # Temporary compatibility during migration of old templates. New and
        # migrated templates use Excel names exclusively.
        return self._discover_legacy_template(workbook)

    def _named_ranges_for_features(
        self, workbook: Any, mappings: list[FeatureMapping]
    ) -> list[CopyRange]:
        ranges: list[CopyRange] = []
        seen: set[tuple[str, str]] = set()
        try:
            names = workbook.Names
            count = int(names.Count)
        except Exception:
            return []
        for index in range(1, count + 1):
            name = names.Item(index)
            if not any(matches_named_range(mapping, str(name.Name)) for mapping in mappings):
                continue
            try:
                resolved = name.RefersToRange
                for area in resolved.Areas:
                    item = CopyRange(
                        str(area.Worksheet.Name), str(area.Address).replace("$", "")
                    )
                    identity = (item.sheet_name, item.address.upper())
                    if identity not in seen:
                        seen.add(identity)
                        ranges.append(item)
            except Exception:
                # Excel sometimes cannot materialize a large non-contiguous
                # name as one COM Range, although the name itself is valid.
                # Its RefersTo formula remains authoritative, so parse each
                # component and keep the areas separate for audit use.
                for item in self._ranges_from_refers_to(workbook, str(name.RefersTo)):
                    identity = (item.sheet_name, item.address.upper())
                    if identity not in seen:
                        seen.add(identity)
                        ranges.append(item)
        return ranges

    def require_named_ranges(self, workbook: Any, mapping: FeatureMapping) -> list[CopyRange]:
        """Require every configured name/prefix to resolve to at least one cell range."""
        missing: list[str] = []
        resolved: list[CopyRange] = []
        for range_name in mapping.range_names:
            expected = FeatureMapping(
                mapping.name, NAMED_RANGE_CHECK_FUNCTION, (range_name,),
                mapping.workbook_limited, mapping.workbook_keyword, mapping.remark,
            )
            areas = self._named_ranges_for_features(workbook, [expected])
            if not areas:
                missing.append(range_name)
            else:
                resolved.extend(areas)
        if missing:
            raise TemplateError(
                f"模块“{mapping.name}”缺少命名区域：" + "、".join(missing)
            )
        return resolved

    def survey_named_ranges(
        self, workbook: Any, mapping: FeatureMapping
    ) -> list[tuple[str, str, str, str]]:
        """Enumerate every configured name and the sheets/areas it covers.

        运行日志使用：列出“哪些表有这个命名区域、覆盖区域是什么”。缺失的
        区域也记一行“缺失”，不抛异常，让日志保留完整的检查记录。
        """
        rows: list[tuple[str, str, str, str]] = []
        for range_name in mapping.range_names:
            expected = FeatureMapping(
                mapping.name, NAMED_RANGE_CHECK_FUNCTION, (range_name,),
                mapping.workbook_limited, mapping.workbook_keyword, mapping.remark,
            )
            areas = self._named_ranges_for_features(workbook, [expected])
            if not areas:
                rows.append(("—", range_name, "—", "缺失"))
            else:
                for area in areas:
                    rows.append((area.sheet_name, range_name, area.address, "通过"))
        return rows

    def _named_ranges_with_prefix(
        self, workbook: Any, prefix: str
    ) -> list[tuple[str, CopyRange]]:
        """Return single-cell named areas whose local name starts with *prefix*."""
        items: list[tuple[str, CopyRange]] = []
        try:
            names = workbook.Names
            count = int(names.Count)
        except Exception:
            return items
        for index in range(1, count + 1):
            name = names.Item(index)
            local_name = str(name.Name).rsplit("!", 1)[-1].strip().strip("'")
            if not local_name.casefold().startswith(prefix.casefold()):
                continue
            field_name = local_name[len(prefix):].strip()
            if not field_name:
                continue
            try:
                resolved = name.RefersToRange
                for area in resolved.Areas:
                    item = CopyRange(
                        str(area.Worksheet.Name), str(area.Address).replace("$", "")
                    )
                    if int(area.Rows.Count) != 1 or int(area.Columns.Count) != 1:
                        raise ValueError(
                            f"命名区域“{local_name}”必须只定位一个单元格"
                        )
                    items.append((field_name, item))
            except ValueError:
                raise
            except Exception as exc:
                raise ValueError(f"无法读取命名区域“{local_name}”：{exc}") from exc
        return items

    @staticmethod
    def _ranges_from_refers_to(workbook: Any, refers_to: str) -> list[CopyRange]:
        text = str(refers_to or "").strip().lstrip("=").strip("()")
        items: list[CopyRange] = []
        for part in text.split(","):
            if "!" not in part:
                continue
            sheet_name, address = part.rsplit("!", 1)
            sheet_name = sheet_name.strip().strip("'").replace("''", "'")
            address = address.strip().replace("$", "")
            if not sheet_name or not re.fullmatch(r"[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?", address):
                continue
            try:
                workbook.Worksheets(sheet_name)
            except Exception:
                continue
            items.append(CopyRange(sheet_name, address))
        return items

    def _definition_from_copy_ranges(
        self,
        workbook: Any,
        copy_ranges: list[CopyRange],
        *,
        structure_ranges: list[CopyRange],
        extraction_ranges: list[CopyRange],
    ) -> TemplateDefinition:
        template_name = normalize_template_name(Path(str(workbook.FullName)).stem)
        # 新式命名区域模板不再生成哈希编号。位置本身就是最直观的规则来源。
        report_code = template_name
        rules: list[AuditRule] = []
        seen_rules: set[tuple[str, str]] = set()
        for item in copy_ranges:
            area = workbook.Worksheets(item.sheet_name).Range(item.address)
            # Read the formula matrix in one COM call.  A typical template has
            # more than a thousand formula cells; SpecialCells + Cells.Item
            # turned discovery into more than a thousand cross-process calls.
            for row_index, formula_row in enumerate(_matrix(area.Formula)):
                for column_index, formula_value in enumerate(formula_row):
                    formula = str(formula_value or "")
                    if not formula.startswith("="):
                        continue
                    cell_address = _a1_address(item.address, row_index, column_index)
                    identity = (item.sheet_name, cell_address)
                    if identity in seen_rules:
                        continue
                    seen_rules.add(identity)
                    rules.append(
                        AuditRule(
                            rule_id=f"{item.sheet_name}-{cell_address}",
                            enabled=True,
                            report_code=report_code,
                            sheet_name=item.sheet_name,
                            formula_cell=cell_address,
                            target_cell=cell_address,
                            severity="错误",
                            message="",
                            copy_range=item.address,
                            result_mode=(
                                "keyword"
                                if self._contains_result_keyword(formula)
                                else "error_only"
                            ),
                            value_from_result=True,
                        )
                    )
        if not rules:
            raise TemplateError("公式校验功能对应的命名区域中没有公式")
        return TemplateDefinition(
            rules=rules,
            copy_ranges=copy_ranges,
            structured=False,
            structure_ranges=structure_ranges,
            extraction_ranges=extraction_ranges,
        )

    def _discover_legacy_template(self, workbook: Any) -> TemplateDefinition:
        template_name = normalize_template_name(Path(str(workbook.FullName)).stem)
        report_code = template_name
        copy_ranges: list[CopyRange] = []
        rules: list[AuditRule] = []
        seen_rules: set[tuple[str, str]] = set()
        seen_rule_ids: dict[str, tuple[str, str]] = {}
        marker_re = re.compile(r"(本地校验区域|联表校验区域)(#?)(\d+)")

        for sheet in workbook.Worksheets:
            starts: dict[tuple[str, str], str] = {}
            ends: dict[tuple[str, str], str] = {}
            try:
                comments = list(sheet.Comments)
            except Exception:
                comments = []
            for comment in comments:
                text = str(comment.Text() or "")
                address = str(comment.Parent.Address).replace("$", "")
                for kind, end_mark, number in marker_re.findall(text):
                    key = (kind, number)
                    if end_mark == "#":
                        if key in ends and ends[key] != address:
                            raise TemplateError(
                                f"{sheet.Name}中的“校验区域#{number}”存在多个结束标记，"
                                "请在名称管理器中改用命名区域“校验区域”"
                            )
                        ends[key] = address
                    else:
                        if key in starts and starts[key] != address:
                            raise TemplateError(
                                f"{sheet.Name}中的“校验区域{number}”存在多个开始标记，"
                                "请在名称管理器中改用命名区域“校验区域”"
                            )
                        starts[key] = address

            for key in sorted(ends.keys() - starts.keys()):
                raise TemplateError(
                    f"{sheet.Name}中的“校验区域#{key[1]}”缺少开始标记“校验区域{key[1]}”，"
                    "请在名称管理器中改用命名区域“校验区域”"
                )

            for key, start in starts.items():
                end = ends.get(key)
                if not end:
                    raise TemplateError(
                        f"{sheet.Name}中的“校验区域{key[1]}”缺少结束标记“校验区域#{key[1]}”，"
                        "请在名称管理器中改用命名区域“校验区域”"
                    )
                area = sheet.Range(f"{start}:{end}")
                area_address = str(area.Address).replace("$", "")
                copy_ranges.append(CopyRange(sheet.Name, area_address))
                try:
                    formula_cells = area.SpecialCells(-4123)
                except Exception:
                    continue
                for cell_index in range(1, int(formula_cells.Cells.Count) + 1):
                    cell = formula_cells.Cells.Item(cell_index)
                    formula = str(cell.Formula or "")
                    cell_address = str(cell.Address).replace("$", "")
                    identity = (sheet.Name, cell_address)
                    if identity in seen_rules:
                        continue
                    seen_rules.add(identity)
                    raw_comment = ""
                    try:
                        if cell.Comment is not None:
                            raw_comment = str(cell.Comment.Text() or "")
                    except Exception:
                        pass
                    explicit_rule_id = parse_comment_rule_id(raw_comment)
                    rule_id = explicit_rule_id or f"{sheet.Name}-{cell_address}"
                    previous = seen_rule_ids.get(rule_id)
                    if previous and previous != identity:
                        raise TemplateError(
                            f"规则编号“{rule_id}”重复："
                            f"{previous[0]}!{previous[1]}、{sheet.Name}!{cell_address}"
                        )
                    seen_rule_ids[rule_id] = identity
                    rules.append(
                        AuditRule(
                            rule_id=rule_id,
                            enabled=True,
                            report_code=report_code,
                            sheet_name=sheet.Name,
                            formula_cell=cell_address,
                            target_cell=cell_address,
                            severity="错误",
                            message=clean_rule_comment(raw_comment),
                            copy_range=area_address,
                            result_mode=(
                                "keyword"
                                if self._contains_result_keyword(formula)
                                else "error_only"
                            ),
                            value_from_result=True,
                        )
                    )

        if not copy_ranges:
            raise TemplateError(
                "模板没有找到名为“校验区域”的命名区域，请在模板的名称管理器中添加"
            )
        if not rules:
            raise TemplateError("已识别校验区域，但区域内没有公式")
        return TemplateDefinition(rules=rules, copy_ranges=copy_ranges, structured=False)

    def inspect_template_health(
        self,
        workbook: Any,
        definition: TemplateDefinition,
    ) -> list[PreflightItem]:
        items: list[PreflightItem] = []
        enabled_rules = [rule for rule in definition.rules if rule.enabled]
        result_rules = [rule for rule in enabled_rules if rule.result_mode != "error_only"]

        if not enabled_rules:
            items.append(
                PreflightItem("规则", "错误", "需处理", "", "", "模板没有启用的校验规则")
            )
        if not result_rules:
            items.append(
                PreflightItem(
                    "规则",
                    "错误",
                    "需处理",
                    "",
                    "",
                    "没有识别到会返回问题标记的公式，审核结果可能始终为零",
                )
            )
        items.append(
            PreflightItem(
                "规则",
                "提示",
                "信息",
                "",
                "",
                f"启用公式 {len(enabled_rules)} 个，其中问题规则 {len(result_rules)} 个、"
                f"公式健康检查 {len(enabled_rules) - len(result_rules)} 个",
            )
        )

        structure_label_count = 0
        for item in definition.structure_ranges:
            try:
                structure_label_count += int(
                    workbook.Worksheets(item.sheet_name).Range(item.address).Cells.Count
                )
            except Exception:
                pass
        items.append(
            PreflightItem(
                "表结构区域",
                "提示" if structure_label_count else "错误",
                "通过" if structure_label_count else "需处理",
                "",
                "",
                (
                    f"已设置 {structure_label_count} 个“表结构区域”命名区域单元格"
                    if structure_label_count
                    else "模板没有“表结构区域”命名区域，无法精确核对报送文件结构"
                ),
            )
        )
        seen_ranges: set[tuple[str, str]] = set()
        range_objects: list[tuple[CopyRange, Any]] = []
        for item in definition.copy_ranges:
            identity = (item.sheet_name, item.address.upper())
            if identity in seen_ranges:
                items.append(
                    PreflightItem(
                        "校验区域", "警告", "关注", item.sheet_name, item.address, "校验区域重复"
                    )
                )
                continue
            seen_ranges.add(identity)
            try:
                area = workbook.Worksheets(item.sheet_name).Range(item.address)
                range_objects.append((item, area))
            except Exception:
                items.append(
                    PreflightItem(
                        "校验区域", "错误", "需处理", item.sheet_name, item.address, "区域地址无效"
                    )
                )
        for index, (left, left_area) in enumerate(range_objects):
            for right, right_area in range_objects[index + 1 :]:
                if left.sheet_name != right.sheet_name:
                    continue
                try:
                    overlap = self.excel.Intersect(left_area, right_area)
                except Exception:
                    overlap = None
                if overlap is not None:
                    items.append(
                        PreflightItem(
                            "校验区域",
                            "警告",
                            "关注",
                            left.sheet_name,
                            f"{left.address} / {right.address}",
                            "两个校验区域发生重叠，请确认是否为有意设置",
                        )
                    )

        # Cache formulas and stored results by range.  This retains the same
        # template health checks while avoiding a COM round trip per formula.
        cached_formulas: dict[tuple[str, str], str] = {}
        cached_values: dict[tuple[str, str], Any] = {}
        for item in definition.copy_ranges:
            area = workbook.Worksheets(item.sheet_name).Range(item.address)
            formula_values = _matrix(area.Formula)
            result_values = _matrix(area.Value2)
            for row_index, formula_row in enumerate(formula_values):
                result_row = result_values[row_index] if row_index < len(result_values) else []
                for column_index, formula_value in enumerate(formula_row):
                    address = _a1_address(item.address, row_index, column_index)
                    key = (item.sheet_name, address)
                    cached_formulas[key] = str(formula_value or "")
                    cached_values[key] = result_row[column_index] if column_index < len(result_row) else ""

        broken_refs: list[str] = []
        calculated_errors: list[str] = []
        for rule in enabled_rules:
            key = (rule.sheet_name, rule.formula_cell)
            formula = cached_formulas.get(key)
            saved_result = cached_values.get(key)
            if formula is None:
                cell = workbook.Worksheets(rule.sheet_name).Range(rule.formula_cell)
                formula = str(cell.Formula or "")
                try:
                    saved_result = cell.Value2
                except Exception:
                    saved_result = ""
            if "#REF!" in formula.upper():
                broken_refs.append(f"{rule.sheet_name}!{rule.formula_cell}")
            error_text = self._formula_error_text(str(saved_result or ""))
            if error_text:
                calculated_errors.append(
                    f"{rule.sheet_name}!{rule.formula_cell}={error_text}"
                )
        if broken_refs:
            items.append(
                PreflightItem(
                    "公式",
                    "警告",
                    "关注",
                    "",
                    "、".join(broken_refs[:10]),
                    f"发现 {len(broken_refs)} 个公式含 #REF! 引用；"
                    "本次允许继续审核，若实际计算触发错误会写入问题结果",
                )
            )
        if calculated_errors:
            items.append(
                PreflightItem(
                    "公式",
                    "警告",
                    "关注",
                    "",
                    "、".join(calculated_errors[:10]),
                    f"模板当前保存值中有 {len(calculated_errors)} 个公式错误；"
                    "可能由模板未填数据导致，审核时仍会逐条提示",
                )
            )

        try:
            links = workbook.LinkSources()
        except Exception:
            links = None
        if links:
            link_values = list(links) if isinstance(links, tuple) else [links]
            items.append(
                PreflightItem(
                    "外部链接",
                    "警告",
                    "关注",
                    "",
                    "",
                    f"模板包含 {len(link_values)} 个外部链接："
                    + "；".join(str(value) for value in link_values[:5]),
                )
            )

        has_error = any(item.level == "错误" for item in items)
        items.insert(
            0,
            PreflightItem(
                "总体",
                "错误" if has_error else "提示",
                "不通过" if has_error else "通过",
                "",
                "",
                "模板存在必须处理的问题" if has_error else "模板结构可用于审核",
            ),
        )
        return items

    def validate_source_workbook(
        self,
        template_workbook: Any,
        source_workbook: Any,
        definition: TemplateDefinition,
    ) -> SourceMatch:
        required_names = sorted({item.sheet_name for item in definition.copy_ranges})
        source_names = {
            source_workbook.Worksheets(index).Name
            for index in range(1, source_workbook.Worksheets.Count + 1)
        }
        missing = tuple(name for name in required_names if name not in source_names)
        if missing:
            return SourceMatch(
                False,
                0.0,
                0,
                0,
                missing,
                "缺少模板要求的工作表：" + "、".join(missing),
            )

        checked = 0
        matched = 0
        mismatches: list[str] = []
        checks: list[StructureCheck] = []
        if not definition.structure_ranges:
            return SourceMatch(
                False, 0.0, 0, 0, (), "模板没有可用的“表结构区域”命名区域，请先维护模板"
            )

        for sheet_name in sorted({item.sheet_name for item in definition.structure_ranges}):
            template_sheet = template_workbook.Worksheets(sheet_name)
            source_sheet = source_workbook.Worksheets(sheet_name)
            for item in (value for value in definition.structure_ranges if value.sheet_name == sheet_name):
                # Range.Value2 returns the whole area in one COM call.  The previous
                # cell-by-cell approach made a 200-cell structure check perform more
                # than 400 cross-process calls for every source workbook.
                expected_values = _matrix(template_sheet.Range(item.address).Value2)
                actual_values = _matrix(source_sheet.Range(item.address).Value2)
                for row_index, expected_row in enumerate(expected_values):
                    actual_row = actual_values[row_index] if row_index < len(actual_values) else []
                    for column_index, expected_value in enumerate(expected_row):
                        if _is_blank(expected_value):
                            continue
                        actual_value = (
                            actual_row[column_index]
                            if column_index < len(actual_row)
                            else None
                        )
                        checked += 1
                        address = _a1_address(item.address, row_index, column_index)
                        is_match = _structure_values_equal(expected_value, actual_value)
                        checks.append(
                            StructureCheck(
                                sheet_name, address, _display_value(expected_value),
                                _display_value(actual_value), is_match
                            )
                        )
                        if is_match:
                            matched += 1
                        elif len(mismatches) < 8:
                            mismatches.append(
                                f"{sheet_name}!{address}：应为“{_display_value(expected_value)}”，"
                                f"实际“{_display_value(actual_value)}”"
                            )

        if checked == 0:
            return SourceMatch(
                False,
                0.0,
                0,
                0,
                (),
                "模板没有可用的“表结构区域”命名区域，请先维护模板",
            )

        score = matched / checked
        enough_matches = matched >= 3 if checked >= 5 else matched >= 1 or checked == 0
        is_match = score >= STRUCTURE_MATCH_THRESHOLD and enough_matches
        details = (
            f"表结构区域精确匹配 {matched}/{checked}（{score:.0%}）"
            + (("；不一致示例：" + "；".join(mismatches)) if mismatches else "")
        )
        return SourceMatch(is_match, score, matched, checked, (), details, tuple(checks))

    def apply_rules(
        self,
        template_workbook: Any,
        audit_workbook: Any,
        definition: TemplateDefinition,
        *,
        calculate: bool = True,
    ) -> None:
        rules = definition.rules
        target_names = {
            audit_workbook.Worksheets(index).Name
            for index in range(1, audit_workbook.Worksheets.Count + 1)
        }
        required_names = {rule.sheet_name for rule in rules if rule.enabled}
        missing = sorted(required_names - target_names)
        if missing:
            raise TemplateError("报送文件缺少工作表：" + "、".join(missing))

        # 某些工作簿（例如含数据模型的模板）不允许设置
        # Application.Calculation=手动。此时 Excel 会在每次写公式后尝试
        # 自动重算，批量复制会非常慢。先在审核副本的工作表级别暂停计算，
        # 复制完再恢复并统一计算；不支持该属性的 Excel 自动降级为原行为。
        suspended_sheets = self._suspend_sheet_calculation(audit_workbook)
        started = time.monotonic()
        try:
            copied: set[tuple[str, str]] = set()
            for item in definition.copy_ranges:
                identity = (item.sheet_name, item.address)
                if identity in copied:
                    continue
                copied.add(identity)
                source = template_workbook.Worksheets(item.sheet_name).Range(item.address)
                target = audit_workbook.Worksheets(item.sheet_name).Range(item.address)
                # 与旧 VBA 的 CopyAndPaste 语义一致：保留格式后，重新写入
                # 公式，使跨表引用始终指向当前审核副本。
                source.Copy(Destination=target)
                target.Formula = source.Formula
            try:
                self.excel.CutCopyMode = False
            except Exception:
                pass
        finally:
            self._restore_sheet_calculation(suspended_sheets)
            self.copy_formula_seconds += time.monotonic() - started

        if definition.structured:
            self._copy_rule_sheet(template_workbook, audit_workbook)
        if calculate:
            self.calculate_pending_workbooks()

    def calculate_pending_workbooks(self) -> None:
        """Calculate all pending audit workbooks once for the current batch."""
        started = time.monotonic()
        try:
            self.excel.Calculate()
        except Exception:
            self.excel.CalculateFullRebuild()
        deadline = time.monotonic() + 120
        while self._calculation_is_pending():
            if time.monotonic() >= deadline:
                raise TimeoutError("Excel公式计算超过120秒")
            time.sleep(0.1)
        self.calculate_seconds += time.monotonic() - started

    def _calculation_is_pending(self) -> bool:
        """WPS may not expose CalculationState; Save still persists its result."""
        try:
            return self.excel.CalculationState != XL_CALCULATION_DONE
        except Exception:
            return False

    @staticmethod
    def _suspend_sheet_calculation(workbook: Any) -> list[tuple[Any, bool]]:
        """Best-effort per-sheet calculation suspension for automatic Excel mode."""
        suspended: list[tuple[Any, bool]] = []
        for index in range(1, workbook.Worksheets.Count + 1):
            sheet = workbook.Worksheets(index)
            try:
                was_enabled = bool(sheet.EnableCalculation)
                if was_enabled:
                    sheet.EnableCalculation = False
                suspended.append((sheet, was_enabled))
            except Exception:
                # 老版本 Excel / 特殊工作表不支持时，维持原有计算方式。
                continue
        return suspended

    @staticmethod
    def _restore_sheet_calculation(sheets: Iterable[tuple[Any, bool]]) -> None:
        for sheet, was_enabled in sheets:
            if not was_enabled:
                continue
            try:
                sheet.EnableCalculation = True
            except Exception:
                pass

    def inspect_external_workbook(
        self, workbook: Any, sheet_names: Iterable[str], *, plan_source: str
    ) -> PreflightItem:
        """Validate an auxiliary workbook before it is copied into audit copies."""
        available = self._worksheet_names(workbook)
        requested = tuple(sheet_names)
        missing = [name for name in requested if name not in available]
        if missing:
            return PreflightItem(
                "外部文件", "错误", "需处理", "", "",
                "外部文件缺少需要复制的工作表：" + "、".join(missing),
            )
        if not requested:
            return PreflightItem(
                "外部文件", "错误", "需处理", "", "", "没有识别到需要复制的外部工作表"
            )
        return PreflightItem(
            "外部文件",
            "提示",
            "通过",
            "",
            "",
            f"{plan_source}，审核时将复制 {len(requested)} 个工作表："
            + "、".join(requested[:8]),
        )

    def validate_external_sheet_names(
        self, source_workbook: Any, sheet_names: Iterable[str]
    ) -> tuple[str, ...]:
        source_names = set(self._worksheet_names(source_workbook))
        return tuple(sorted(source_names & set(sheet_names)))

    def copy_external_sheets(
        self, external_workbook: Any, audit_workbook: Any, sheet_names: Iterable[str]
    ) -> None:
        """Copy selected auxiliary sheets into the audit copy, never into source data."""
        names = tuple(sheet_names)
        conflicts = self.validate_external_sheet_names(audit_workbook, names)
        if conflicts:
            raise TemplateError(
                "外部文件工作表与报送文件重名，不能覆盖原表：" + "、".join(conflicts)
            )
        started = time.monotonic()
        try:
            for name in names:
                # Excel's COM binding does not reliably honour named optional
                # arguments for Worksheet.Copy.  Positional Before, After keeps
                # the copied tab in the audit workbook instead of a new workbook.
                external_workbook.Worksheets(name).Copy(
                    None, audit_workbook.Worksheets(audit_workbook.Worksheets.Count)
                )
        finally:
            self.external_copy_seconds += time.monotonic() - started

    @staticmethod
    def _worksheet_names(workbook: Any) -> tuple[str, ...]:
        return tuple(
            str(workbook.Worksheets(index).Name)
            for index in range(1, workbook.Worksheets.Count + 1)
        )

    def _copy_rule_sheet(self, template_workbook: Any, audit_workbook: Any) -> None:
        names = {
            audit_workbook.Worksheets(index).Name
            for index in range(1, audit_workbook.Worksheets.Count + 1)
        }
        if "审核规则" in names:
            audit_workbook.Worksheets("审核规则").Delete()
        template_workbook.Worksheets("审核规则").Copy(
            After=audit_workbook.Worksheets(audit_workbook.Worksheets.Count)
        )

    def extract_issues(
        self,
        audit_workbook: Any,
        rules: Iterable[AuditRule],
        *,
        period: str,
        batch_id: str,
        audit_time: str,
        org_code: str,
        org_name: str,
        source_file: Path,
        audit_file: Path,
        extraction_ranges: Iterable[CopyRange] | None = None,
        saved_workbook_path: Path | None = None,
    ) -> list[Issue]:
        issues: list[Issue] = []
        allowed = tuple(extraction_ranges or ())
        enabled_rules = [
            rule
            for rule in rules
            if rule.enabled
            and (
                not allowed
                or any(
                    item.sheet_name == rule.sheet_name
                    and _cell_in_address(rule.formula_cell, item.address)
                    for item in allowed
                )
            )
        ]

        # 不要逐格通过 COM 取 Value2/Text。对于本期 7 个工作簿、360 个
        # 公式，原实现产生数千次跨进程调用，耗时约 50 秒。每个工作表只取
        # 一次覆盖公式单元格与定位单元格的最小矩形，随后全部在 Python 内存
        # 中解析。这里读的是尚未保存的 Excel 内存结果，openpyxl/pandas
        # 无法替代这个阶段（它们只能读磁盘中的上一次缓存结果）。
        if saved_workbook_path is not None:
            # Excel 已计算并保存后，openpyxl 可稳定读取 xlsx 中的公式缓存，
            # 包括 COM 批量 Value2 会吞掉的 #N/A/#REF! 等错误值。
            values = self._read_issue_values_from_xlsx(saved_workbook_path, enabled_rules)
            error_texts: dict[tuple[str, str], str] = {}
        else:
            values, formulas = self._read_issue_values(audit_workbook, enabled_rules)
            error_texts = self._read_formula_error_texts(
                audit_workbook, enabled_rules, values, formulas
            )
        for rule in enabled_rules:
            result = self._cached_cell_value(values, rule.sheet_name, rule.formula_cell)
            formula_error = self._formula_error_text(str(result or "")) or error_texts.get(
                (rule.sheet_name, rule.formula_cell), ""
            )
            if formula_error:
                issue_id = self._issue_id(org_code, source_file, rule)
                if formula_error == "#REF!":
                    detail = (
                        "校验公式引用无效（#REF!）。如本公式依赖外部数据，请确认已选择外部文件，"
                        "且所需工作表已在“外部文件添加”步骤复制进审核副本；"
                        "若外部表已存在，请检查模板公式中的工作表名和单元格引用。"
                    )
                else:
                    detail = f"校验公式计算异常：{formula_error}，请检查模板公式及引用数据"
                issues.append(
                    Issue(
                        issue_id=issue_id,
                        period=period,
                        batch_id=batch_id,
                        audit_time=audit_time,
                        triggered=True,
                        status="",
                        first_seen_period="",
                        previous_seen_period="",
                        consecutive_count=1,
                        org_code=org_code,
                        org_name=org_name,
                        report_code=rule.report_code,
                        sheet_name=rule.sheet_name,
                        rule_id=rule.rule_id,
                        severity="错误",
                        formula_cell=rule.formula_cell,
                        target_cell=rule.formula_cell,
                        target_value="",
                        formula_result=formula_error,
                        message=detail,
                        source_file=str(source_file.resolve()),
                        audit_file=str(audit_file.resolve()),
                        detail=detail,
                    )
                )
                continue
            if rule.result_mode == "error_only":
                continue
            if result is None or result is False or result == 0:
                continue
            if isinstance(result, str) and not result.strip():
                continue
            result_text = str(result)
            if not self._has_result_marker(result_text):
                # Structured rules may deliberately return another non-empty flag.
                if rule.result_mode == "keyword":
                    continue
            parsed = parse_formula_result(
                result_text,
                default_severity=rule.severity,
                default_message=rule.message,
            )
            target_value = (
                parsed.value
                if rule.value_from_result
                else self._cached_cell_value(values, rule.sheet_name, rule.target_cell)
            )
            issue_id = self._issue_id(org_code, source_file, rule, parsed.indicator)
            issues.append(
                Issue(
                    issue_id=issue_id,
                    period=period,
                    batch_id=batch_id,
                    audit_time=audit_time,
                    triggered=True,
                    status="",
                    first_seen_period="",
                    previous_seen_period="",
                    consecutive_count=1,
                    org_code=org_code,
                    org_name=org_name,
                    report_code=rule.report_code,
                    sheet_name=rule.sheet_name,
                    rule_id=rule.rule_id,
                    severity=parsed.severity,
                    formula_cell=rule.formula_cell,
                    target_cell=rule.target_cell,
                    target_value=_plain(target_value),
                    formula_result=_plain(result),
                    message=parsed.message,
                    source_file=str(source_file.resolve()),
                    audit_file=str(audit_file.resolve()),
                    check_field=parsed.indicator,
                    comparison_value=parsed.comparison_value,
                    reference_value=parsed.reference_value,
                    difference_value=parsed.difference_value,
                    detail=parsed.detail,
                )
            )
        return issues

    def extract_conditional_format_issues(
        self,
        workbook: Any,
        *,
        mapping: FeatureMapping,
        structure_ranges: Iterable[CopyRange],
        period: str,
        batch_id: str,
        audit_time: str,
        org_code: str,
        org_name: str,
        source_file: Path,
        workbook_path: Path | None = None,
    ) -> list[Issue]:
        """Extract cells whose conditional formatting changes the fill colour.

        ``DisplayFormat`` is deliberately used instead of the stored cell fill:
        the latter describes the base style and cannot tell whether a conditional
        formatting rule has actually been triggered.  A named ``条件格式区域`` is
        preferred; without it only the sheets' conditional-format applies-to
        ranges are scanned.

        When the host engine cannot expose a rendered colour (WPS 12.0 returns
        ``None`` for ``DisplayFormat.Interior.Color``) and a ``workbook_path`` is
        available, fall back to evaluating the submission's own ``cellIs`` /
        ``expression`` rules from OOXML rather than failing the whole step.
        """
        self.last_conditional_unsupported_count = 0
        try:
            return self._extract_conditional_format_issues_com(
                workbook, mapping=mapping, structure_ranges=structure_ranges,
                period=period, batch_id=batch_id, audit_time=audit_time,
                org_code=org_code, org_name=org_name, source_file=source_file,
            )
        except RuntimeError as exc:
            if workbook_path is None or "实际显示颜色" not in str(exc):
                raise
            issues, unsupported = self._extract_conditional_format_issues_from_rules(
                workbook_path=workbook_path, mapping=mapping,
                structure_ranges=structure_ranges, period=period, batch_id=batch_id,
                audit_time=audit_time, org_code=org_code, org_name=org_name,
                source_file=source_file,
            )
            self.last_conditional_unsupported_count = unsupported
            return issues

    def _extract_conditional_format_issues_com(
        self,
        workbook: Any,
        *,
        mapping: FeatureMapping,
        structure_ranges: Iterable[CopyRange],
        period: str,
        batch_id: str,
        audit_time: str,
        org_code: str,
        org_name: str,
        source_file: Path,
    ) -> list[Issue]:
        """Read a rendered conditional-format fill colour through COM."""
        requested = self._named_ranges_for_features(workbook, [mapping])
        rule_scopes: dict[str, list[tuple[int, int, int, int, str, Any]]] = {}

        def scopes_for(sheet: Any) -> list[tuple[int, int, int, int, str, Any]] | None:
            name = str(sheet.Name)
            if name in rule_scopes:
                return rule_scopes[name]
            try:
                conditions = sheet.UsedRange.FormatConditions
                scopes: list[tuple[int, int, int, int, str, Any]] = []
                for condition_index in range(1, int(conditions.Count) + 1):
                    condition = conditions.Item(condition_index)
                    applies_to = condition.AppliesTo
                    for area_index in range(1, int(applies_to.Areas.Count) + 1):
                        applies_area = applies_to.Areas.Item(area_index)
                        row, column = int(applies_area.Row), int(applies_area.Column)
                        scopes.append((
                            column, row,
                            column + int(applies_area.Columns.Count) - 1,
                            row + int(applies_area.Rows.Count) - 1,
                            _column_letters(column) + str(row),
                            condition,
                        ))
                rule_scopes[name] = scopes
                return scopes
            except Exception:
                # An unavailable collection must not silently hide a named
                # conditional-format range; use the explicit range as fallback.
                return None

        def matching_rules(sheet_name: str, row: int, column: int) -> list[tuple[str, Any]] | None:
            scopes = rule_scopes.get(sheet_name)
            if scopes is None:
                return None
            return [
                (origin, condition)
                for left, top, right, bottom, origin, condition in scopes
                if left <= column <= right and top <= row <= bottom
            ]

        scan_areas: list[tuple[Any, Any, list[tuple[int, int, int, int, str, Any]] | None]] = []
        if requested:
            for item in requested:
                sheet = workbook.Worksheets(item.sheet_name)
                scan_areas.append((sheet, sheet.Range(item.address), scopes_for(sheet)))
        else:
            for index in range(1, workbook.Worksheets.Count + 1):
                sheet = workbook.Worksheets(index)
                scopes = scopes_for(sheet)
                if scopes is None:
                    continue
                for left, top, right, bottom, _origin, _condition in scopes:
                    scan_areas.append((sheet, sheet.Range(
                        "{}{}:{}{}".format(_column_letters(left), top, _column_letters(right), bottom)
                    ), scopes))

        issues: list[Issue] = []
        seen: set[tuple[str, int, int]] = set()
        for sheet, area, scopes in scan_areas:
            self._prepare_conditional_format_sheet(workbook, sheet)
            first_row, first_column = int(area.Row), int(area.Column)
            rows, columns = int(area.Rows.Count), int(area.Columns.Count)
            values = _range_matrix(area.Value2, rows, columns)
            for row_offset in range(rows):
                for column_offset in range(columns):
                    try:
                        value = values[row_offset][column_offset]
                    except IndexError:
                        value = None
                    if value in (None, ""):
                        continue
                    row, column = first_row + row_offset, first_column + column_offset
                    matched_rules = matching_rules(str(sheet.Name), row, column)
                    if matched_rules == []:
                        continue
                    identity = (str(sheet.Name), row, column)
                    if identity in seen:
                        continue
                    seen.add(identity)
                    cell = sheet.Cells(row, column)
                    color = self._active_conditional_fill_color(cell)
                    if color is None:
                        continue
                    target_cell = _column_letters(column) + str(row)
                    check_field = self._conditional_check_field(sheet, row, column, structure_ranges)
                    comment = self._cell_comment_text(cell)
                    detail = comment or self._conditional_rules_text(matched_rules, target_cell)
                    rule = AuditRule(
                        rule_id="条件格式填充", enabled=True, report_code="",
                        sheet_name=str(sheet.Name), formula_cell=target_cell, target_cell=target_cell,
                        severity="总行条件格式触发", message=detail,
                    )
                    issues.append(Issue(
                        issue_id=self._issue_id(org_code, source_file, rule, check_field),
                        period=period, batch_id=batch_id, audit_time=audit_time, triggered=True,
                        status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                        org_code=org_code, org_name=org_name, report_code="", sheet_name=str(sheet.Name),
                        rule_id="条件格式填充", severity="总行条件格式触发",
                        formula_cell=target_cell, target_cell=target_cell, target_value=_plain(value),
                        formula_result=_plain(value), message=detail, source_file=str(source_file.resolve()),
                        audit_file=str(source_file.resolve()), check_field=check_field, detail=detail,
                    ))
        return issues

    # —— WPS OOXML fallback：不猜、不报触发，无法解析的规则计入 unsupported ——

    def _named_ranges_via_openpyxl(self, book: Any, mapping: FeatureMapping) -> list[CopyRange]:
        """Resolve ``条件格式区域`` names to areas from an openpyxl workbook."""
        result: list[CopyRange] = []
        seen: set[tuple[str, str]] = set()
        names = list(book.defined_names.values())
        for sheet in book.worksheets:
            names.extend(sheet.defined_names.values())
        for defined_name in names:
            if not matches_named_range(mapping, defined_name.name):
                continue
            try:
                destinations = tuple(defined_name.destinations)
            except (AttributeError, ValueError):
                continue
            for sheet_name, address in destinations:
                if not sheet_name or sheet_name not in book.sheetnames:
                    continue
                item = CopyRange(sheet_name, address.replace("$", "").strip())
                identity = (item.sheet_name, item.address.upper())
                if identity not in seen:
                    seen.add(identity)
                    result.append(item)
        return result

    @staticmethod
    def _cellis_operand(sheet: Any, raw: object) -> object:
        """Resolve a simple OOXML ``cellIs`` threshold without guessing."""
        text = str(raw or "").strip().lstrip("=")
        if not text:
            return None
        if len(text) >= 2 and text[0] == text[-1] == '"':
            return text[1:-1]
        try:
            return float(text)
        except ValueError:
            pass
        match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", text)
        if match:
            return sheet["{}{}".format(match.group(1), match.group(2))].value
        return None

    @classmethod
    def _cellis_matches_rule(cls, rule: Any, value: object, sheet: Any) -> bool:
        if value in (None, ""):
            return False
        formulas = list(rule.formula or [])
        first = cls._cellis_operand(sheet, formulas[0] if formulas else None)
        second = cls._cellis_operand(sheet, formulas[1] if len(formulas) > 1 else None)
        if first is None:
            return False
        try:
            left, right = float(value), float(first)
            upper = float(second) if second is not None else None
        except (TypeError, ValueError):
            left, right, upper = str(value), str(first), str(second) if second is not None else None
        operator = str(rule.operator or "").casefold()
        return {
            "equal": left == right, "notequal": left != right,
            "greaterthan": left > right, "greaterthanorequal": left >= right,
            "lessthan": left < right, "lessthanorequal": left <= right,
            "between": upper is not None and right <= left <= upper,
            "notbetween": upper is not None and not (right <= left <= upper),
        }.get(operator, False)

    @staticmethod
    def _has_dxf_fill(book: Any, rule: Any) -> bool:
        try:
            style = book._differential_styles[rule.dxfId]
            fill = style.fill
            return bool(fill and (fill.patternType or fill.fgColor.type or fill.bgColor.type))
        except Exception:
            return False

    @staticmethod
    def _expression_rule_text(formula: str, origin: str, destination: str) -> str:
        """Return a destination-relative expression rule description.

        openpyxl stores the formula without a leading ``=``, but the Translator
        only rewrites relative references when it sees a formula.
        """
        from openpyxl.formula.translate import Translator

        text = str(formula or "").strip()
        if text and not text.startswith("="):
            text = "=" + text
        try:
            translated = Translator(text, origin=origin).translate_formula(destination)
        except Exception:
            translated = text
        return "条件格式规则：{}".format(str(translated).lstrip("=").replace("$", ""))

    @staticmethod
    def _conditional_check_field_xlsx(
        sheet: Any, row: int, column: int, structure_ranges: Iterable[CopyRange],
    ) -> str:
        """Build ``left labels｜top labels`` from openpyxl table-structure areas."""
        from openpyxl.utils.cell import range_boundaries

        areas = [item for item in structure_ranges if item.sheet_name == sheet.title]
        if not areas:
            return ""
        bounds: list[tuple[int, int, int, int]] = []
        for item in areas:
            try:
                min_col, min_row, max_col, max_row = range_boundaries(item.address)
            except ValueError:
                continue
            bounds.append((min_row, min_col, max_row, max_col))
        if not bounds:
            return ""

        def text_value(r: int, c: int) -> str:
            try:
                value = sheet.cell(row=r, column=c).value
            except Exception:
                return ""
            if value is None or (isinstance(value, (int, float)) and not isinstance(value, bool)):
                return ""
            text = str(value).strip()
            if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text.replace(",", "")):
                return ""
            return text

        min_row = min(item[0] for item in bounds)
        min_col = min(item[1] for item in bounds)
        left: list[str] = []
        for c in range(min_col, column):
            value = text_value(row, c)
            if value and value not in left:
                left.append(value)
        top: list[str] = []
        for r in range(min_row, row):
            value = text_value(r, column)
            if value and value not in top:
                top.append(value)
        parts = []
        if left:
            parts.append("_".join(left))
        if top:
            parts.append("_".join(top))
        return "｜".join(parts)

    def _extract_conditional_format_issues_from_rules(
        self,
        *,
        workbook_path: Path,
        mapping: FeatureMapping,
        structure_ranges: Iterable[CopyRange],
        period: str,
        batch_id: str,
        audit_time: str,
        org_code: str,
        org_name: str,
        source_file: Path,
    ) -> tuple[list[Issue], int]:
        """Evaluate ``cellIs``/``expression`` rules from OOXML for WPS.

        Returns ``(issues, unsupported_rule_count)``.  Unparseable rules are
        not guessed: they yield no trigger and are only counted so the run log
        can report “WPS 条件规则暂不支持”.
        """
        from openpyxl import load_workbook
        from openpyxl.utils.cell import range_boundaries

        book = load_workbook(workbook_path, data_only=True, keep_links=False)
        try:
            # 条件格式区域 named ranges limit scanning; when the source file
            # lacks them, fall back to every sheet's applies-to ranges (matching
            # the COM path).
            named_bounds: dict[str, list[tuple[int, int, int, int]]] = {}
            for area in self._named_ranges_via_openpyxl(book, mapping):
                if area.sheet_name not in book.sheetnames:
                    continue
                try:
                    left, top, right, bottom = range_boundaries(area.address)
                except ValueError:
                    continue
                named_bounds.setdefault(area.sheet_name, []).append((left, top, right, bottom))

            issues: list[Issue] = []
            seen: set[tuple[str, int, int]] = set()
            unsupported = 0
            symbols = {
                "equal": "=", "notequal": "<>", "greaterthan": ">",
                "greaterthanorequal": ">=", "lessthan": "<", "lessthanorequal": "<=",
                "between": "介于", "notbetween": "不介于",
            }
            for sheet in book.worksheets:
                if not sheet.conditional_formatting:
                    continue
                sheet_named = named_bounds.get(sheet.title)
                for conditional in sheet.conditional_formatting:
                    try:
                        applies_ranges = list(conditional.sqref.ranges)
                    except Exception:
                        applies_ranges = []
                    if not applies_ranges:
                        continue
                    # Excel anchors a rule's relative references to the top-left
                    # cell of its AppliesTo; a multi-area AppliesTo keeps one
                    # anchor (the top-left of the whole range).
                    anchor_col = min(applies.bounds[0] for applies in applies_ranges)
                    anchor_row = min(applies.bounds[1] for applies in applies_ranges)
                    anchor_address = _column_letters(anchor_col) + str(anchor_row)
                    for rule in conditional.rules:
                        rule_type = str(rule.type or "").casefold()
                        if rule_type not in ("cellis", "expression"):
                            unsupported += 1
                            continue
                        if not self._has_dxf_fill(book, rule):
                            continue
                        formula = None
                        if rule_type == "expression":
                            formula = str((rule.formula or [""])[0] or "").strip()
                            if formula.startswith("="):
                                formula = formula[1:].strip()
                            if not formula:
                                unsupported += 1
                                continue
                            try:
                                evaluate_expression_formula(
                                    formula, sheet, anchor_row, anchor_col, anchor_row, anchor_col
                                )
                            except (ValueError, NotImplementedError):
                                unsupported += 1
                                continue
                        for applies in applies_ranges:
                            min_col, min_row, max_col, max_row = applies.bounds
                            # Effective scan bounds: intersect the applies-to with
                            # each named area (when present), else scan the whole
                            # applies-to.
                            scan_bounds = [
                                (max(n_left, min_col), max(n_top, min_row),
                                 min(n_right, max_col), min(n_bottom, max_row))
                                for n_left, n_top, n_right, n_bottom in (sheet_named or [])
                            ] or [(min_col, min_row, max_col, max_row)]
                            for left, top, right, bottom in scan_bounds:
                                if left > right or top > bottom:
                                    continue
                                for row in range(top, bottom + 1):
                                    for column in range(left, right + 1):
                                        identity = (sheet.title, row, column)
                                        if identity in seen:
                                            continue
                                        value = sheet.cell(row=row, column=column).value
                                        if rule_type == "cellis":
                                            matched = self._cellis_matches_rule(rule, value, sheet)
                                        else:
                                            matched = evaluate_expression_formula(
                                                formula, sheet, anchor_row, anchor_col, row, column
                                            )
                                        if not matched:
                                            continue
                                        seen.add(identity)
                                        target_cell = _column_letters(column) + str(row)
                                        if rule_type == "cellis":
                                            threshold = str((rule.formula or [""])[0]).lstrip("=").replace("$", "")
                                            detail = "条件格式规则：{}{}{}".format(
                                                target_cell, symbols.get(str(rule.operator or "").casefold(), "?"), threshold
                                            )
                                        else:
                                            detail = self._expression_rule_text(formula, anchor_address, target_cell)
                                        comment = sheet.cell(row=row, column=column).comment
                                        if comment and comment.text and comment.text.strip():
                                            detail = comment.text.strip()
                                        check_field = self._conditional_check_field_xlsx(
                                            sheet, row, column, structure_ranges
                                        )
                                        audit_rule = AuditRule(
                                            rule_id="条件格式填充", enabled=True, report_code="",
                                            sheet_name=str(sheet.title), formula_cell=target_cell,
                                            target_cell=target_cell, severity="总行条件格式触发", message=detail,
                                        )
                                        issues.append(Issue(
                                            issue_id=self._issue_id(org_code, source_file, audit_rule, check_field),
                                            period=period, batch_id=batch_id, audit_time=audit_time, triggered=True,
                                            status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                                            org_code=org_code, org_name=org_name, report_code="", sheet_name=str(sheet.title),
                                            rule_id="条件格式填充", severity="总行条件格式触发",
                                            formula_cell=target_cell, target_cell=target_cell, target_value=_plain(value),
                                            formula_result=_plain(value), message=detail, source_file=str(source_file.resolve()),
                                            audit_file=str(source_file.resolve()), check_field=check_field, detail=detail,
                                        ))
            return issues, unsupported
        finally:
            book.close()

    def _prepare_conditional_format_sheet(self, workbook: Any, sheet: Any) -> None:
        """Refresh WPS's rendered conditional-format cache once per sheet.

        WPS may return the static ``Interior.Color`` from ``DisplayFormat`` when
        the private application is launched hidden with screen updates disabled.
        Activating the sheet and recalculating once is materially cheaper than
        doing it per cell and keeps the COM instance private to this task.
        """
        if self.engine_name != "WPS 表格":
            return
        app = self.excel
        if app is None:
            return
        try:
            app.ScreenUpdating = True
        except Exception:
            pass
        try:
            workbook.Activate()
        except Exception:
            pass
        try:
            sheet.Activate()
        except Exception:
            pass
        try:
            app.CalculateFullRebuild()
        except Exception:
            try:
                app.Calculate()
            except Exception:
                pass
        try:
            app.ScreenUpdating = False
        except Exception:
            pass

    @staticmethod
    def _active_conditional_fill_color(cell: Any) -> int | None:
        try:
            displayed = int(cell.DisplayFormat.Interior.Color)
            base = int(cell.Interior.Color)
        except Exception as exc:
            raise RuntimeError(
                "当前表格引擎无法读取条件格式的实际显示颜色；请改用 Microsoft Excel 或 Windows 版 WPS 后重试"
            ) from exc
        return displayed if displayed != base else None

    @staticmethod
    def _conditional_rule_text(condition: Any, *, origin: str, destination: str) -> str:
        """Return a readable, destination-relative Excel condition formula.

        ``FormatCondition.Formula1`` is stored relative to the first cell of
        ``AppliesTo``.  Translate it to the triggered cell so the summary shows
        the rule the reviewer can verify (``AND(C5>0,C5>1)``) instead of a
        generic "条件格式填充已触发" message.
        """
        from openpyxl.formula.translate import Translator

        formulas: list[str] = []
        for attribute in ("Formula1", "Formula2"):
            try:
                value = str(getattr(condition, attribute) or "").strip()
            except Exception:
                value = ""
            if not value:
                continue
            try:
                value = Translator(value, origin=origin).translate_formula(destination)
            except Exception:
                pass
            formulas.append(value.lstrip("=").replace("$", ""))

        try:
            condition_type = int(getattr(condition, "Type"))
        except Exception:
            condition_type = None
        try:
            operator = int(getattr(condition, "Operator"))
        except Exception:
            operator = None
        operator_symbols = {
            1: (">=", "<="),  # xlBetween
            2: ("<", ">"),    # xlNotBetween
            3: ("=",),         # xlEqual
            4: ("<>",),        # xlNotEqual
            5: (">",),         # xlGreater
            6: ("<",),         # xlLess
            7: (">=",),        # xlGreaterEqual
            8: ("<=",),        # xlLessEqual
        }
        if condition_type == 1 and formulas and operator in operator_symbols:
            symbols = operator_symbols[operator]
            if operator == 1 and len(formulas) >= 2:
                text = "{}{}{} 且 {}{}{}".format(
                    destination, symbols[0], formulas[0], destination, symbols[1], formulas[1]
                )
            elif operator == 2 and len(formulas) >= 2:
                text = "{}{}{} 或 {}{}{}".format(
                    destination, symbols[0], formulas[0], destination, symbols[1], formulas[1]
                )
            else:
                text = "{}{}{}".format(destination, symbols[0], formulas[0])
            return "条件格式规则：{}".format(text)
        if formulas:
            return "条件格式规则：{}".format("；".join(dict.fromkeys(formulas)))
        try:
            rule_type = str(getattr(condition, "Type"))
        except Exception:
            rule_type = "未提供公式"
        return "条件格式规则：类型 {}（未提供可显示的公式）".format(rule_type)

    def _conditional_rules_text(self, matched_rules, destination: str) -> str:
        """Join the destination-relative rule text of every matching condition."""
        if not matched_rules:
            return "条件格式填充已触发，请核实"
        descriptions = [
            self._conditional_rule_text(condition, origin=origin, destination=destination)
            for origin, condition in matched_rules
        ]
        return "；".join(dict.fromkeys(descriptions)) or "条件格式填充已触发，请核实"

    @staticmethod
    def _cell_comment_text(cell: Any) -> str:
        try:
            comment = cell.Comment
            if comment is not None:
                return str(comment.Text() or "").strip()
        except Exception:
            pass
        return ""

    @staticmethod
    def _conditional_check_field(
        sheet: Any, row: int, column: int, structure_ranges: Iterable[CopyRange],
    ) -> str:
        """Build ``left labels｜top labels`` from configured table structure.

        Reading a merged cell through ``MergeArea.Cells(1, 1)`` keeps parent
        labels available on child rows/columns and supports multi-level headers.
        """
        areas = [item for item in structure_ranges if item.sheet_name == str(sheet.Name)]
        if not areas:
            return ""
        bounds: list[tuple[int, int, int, int]] = []
        for item in areas:
            try:
                area = sheet.Range(item.address)
                bounds.append((int(area.Row), int(area.Column), int(area.Rows.Count), int(area.Columns.Count)))
            except Exception:
                continue
        if not bounds:
            return ""

        def display_value(r: int, c: int) -> str:
            try:
                cell = sheet.Cells(r, c)
                merged = cell.MergeArea
                cell = merged.Cells(1, 1)
                value = cell.Value2
                if value is None or isinstance(value, (int, float)) and not isinstance(value, bool):
                    return ""
                text = str(value).strip()
                # 表结构区域常覆盖数据列。只把文本型行、列表头拼成校验指标，
                # 不能把本期/上期金额等数值带入“借款人证件类型_合计”。
                if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text.replace(",", "")):
                    return ""
                return text
            except Exception:
                return ""

        min_row = min(item[0] for item in bounds)
        min_col = min(item[1] for item in bounds)
        left: list[str] = []
        for c in range(min_col, column):
            value = display_value(row, c)
            if value and value not in left:
                left.append(value)
        top: list[str] = []
        for r in range(min_row, row):
            value = display_value(r, column)
            if value and value not in top:
                top.append(value)
        parts = []
        if left:
            parts.append("_".join(left))
        if top:
            parts.append("_".join(top))
        return "｜".join(parts)

    @staticmethod
    def _read_issue_values(
        workbook: Any, rules: Iterable[AuditRule]
    ) -> tuple[
        dict[str, tuple[int, int, list[list[Any]]]],
        dict[str, tuple[int, int, list[list[Any]]]],
    ]:
        """Read all formula/target values for each sheet with two COM calls."""
        cells_by_sheet: dict[str, list[tuple[int, int]]] = {}
        for rule in rules:
            cells_by_sheet.setdefault(rule.sheet_name, []).append(
                _cell_position(rule.formula_cell)
            )
            if not rule.value_from_result:
                cells_by_sheet[rule.sheet_name].append(_cell_position(rule.target_cell))

        values: dict[str, tuple[int, int, list[list[Any]]]] = {}
        formulas: dict[str, tuple[int, int, list[list[Any]]]] = {}
        for sheet_name, cells in cells_by_sheet.items():
            rows = [item[0] for item in cells]
            columns = [item[1] for item in cells]
            first_row, last_row = min(rows), max(rows)
            first_column, last_column = min(columns), max(columns)
            address = _rectangle_address(first_row, first_column, last_row, last_column)
            area = workbook.Worksheets(sheet_name).Range(address)
            matrix = _matrix(area.Value2)
            values[sheet_name] = (first_row, first_column, matrix)
            formulas[sheet_name] = (first_row, first_column, _matrix(area.Formula))
        return values, formulas

    @staticmethod
    def _read_issue_values_from_xlsx(
        path: Path, rules: Iterable[AuditRule]
    ) -> dict[str, tuple[int, int, list[list[Any]]]]:
        """Read saved Excel calculation cache without per-cell COM calls."""
        from openpyxl import load_workbook

        cells_by_sheet: dict[str, list[tuple[int, int]]] = {}
        for rule in rules:
            cells_by_sheet.setdefault(rule.sheet_name, []).append(
                _cell_position(rule.formula_cell)
            )
            if not rule.value_from_result:
                cells_by_sheet[rule.sheet_name].append(_cell_position(rule.target_cell))

        result: dict[str, tuple[int, int, list[list[Any]]]] = {}
        workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            for sheet_name, cells in cells_by_sheet.items():
                if sheet_name not in workbook.sheetnames:
                    raise TemplateError(f"审核副本缺少工作表：{sheet_name}")
                rows = [item[0] for item in cells]
                columns = [item[1] for item in cells]
                first_row, last_row = min(rows), max(rows)
                first_column, last_column = min(columns), max(columns)
                sheet = workbook[sheet_name]
                matrix = [
                    list(row)
                    for row in sheet.iter_rows(
                        min_row=first_row,
                        max_row=last_row,
                        min_col=first_column,
                        max_col=last_column,
                        values_only=True,
                    )
                ]
                result[sheet_name] = (first_row, first_column, matrix)
        finally:
            workbook.close()
        return result

    def _read_formula_error_texts(
        self,
        workbook: Any,
        rules: Iterable[AuditRule],
        values: dict[str, tuple[int, int, list[list[Any]]]],
        formulas: dict[str, tuple[int, int, list[list[Any]]]],
    ) -> dict[tuple[str, str], str]:
        """Read display errors only for formula cells where bulk Value2 is blank.

        Excel COM occasionally represents errors in a multi-cell Value2 read as
        ``None``.  Group adjacent candidates into row ranges first; most #N/A
        blocks then require one Text call rather than one call per cell.
        """
        candidates: dict[str, list[tuple[int, int, str]]] = {}
        for rule in rules:
            result = self._cached_cell_value(values, rule.sheet_name, rule.formula_cell)
            formula = self._cached_cell_value(formulas, rule.sheet_name, rule.formula_cell)
            if result is None and str(formula or "").startswith("="):
                row, column = _cell_position(rule.formula_cell)
                candidates.setdefault(rule.sheet_name, []).append(
                    (row, column, rule.formula_cell)
                )

        errors: dict[tuple[str, str], str] = {}
        for sheet_name, cells in candidates.items():
            by_row: dict[int, list[tuple[int, str]]] = {}
            for row, column, address in cells:
                by_row.setdefault(row, []).append((column, address))
            sheet = workbook.Worksheets(sheet_name)
            for row, column_cells in by_row.items():
                column_cells.sort()
                start = 0
                while start < len(column_cells):
                    end = start
                    while (
                        end + 1 < len(column_cells)
                        and column_cells[end + 1][0] == column_cells[end][0] + 1
                    ):
                        end += 1
                    run = column_cells[start : end + 1]
                    address = _rectangle_address(row, run[0][0], row, run[-1][0])
                    try:
                        display = self._formula_error_text(str(sheet.Range(address).Text or ""))
                    except Exception:
                        display = ""
                    if display:
                        for _, cell_address in run:
                            errors[(sheet_name, cell_address)] = display
                    elif len(run) == 1:
                        # A non-error blank formula falls through harmlessly.
                        pass
                    else:
                        # Range.Text is Null when cells display different errors.
                        # This uncommon fallback still reads only the ambiguous run.
                        for _, cell_address in run:
                            try:
                                display = self._formula_error_text(
                                    str(sheet.Range(cell_address).Text or "")
                                )
                            except Exception:
                                display = ""
                            if display:
                                errors[(sheet_name, cell_address)] = display
                    start = end + 1
        return errors

    @staticmethod
    def _cached_cell_value(
        values: dict[str, tuple[int, int, list[list[Any]]]], sheet_name: str, cell: str
    ) -> Any:
        first_row, first_column, matrix = values[sheet_name]
        row, column = _cell_position(cell)
        row_index, column_index = row - first_row, column - first_column
        if row_index < 0 or column_index < 0:
            return None
        if row_index >= len(matrix) or column_index >= len(matrix[row_index]):
            return None
        return matrix[row_index][column_index]

    @staticmethod
    def _formula_error_text(value: str) -> str:
        text = value.strip().upper()
        known = (
            "#REF!",
            "#VALUE!",
            "#N/A",
            "#DIV/0!",
            "#NAME?",
            "#NUM!",
            "#NULL!",
            "#SPILL!",
            "#CALC!",
            "#FIELD!",
            "#GETTING_DATA",
        )
        return next((item for item in known if text.startswith(item)), "")

    @staticmethod
    def _contains_result_keyword(value: Any) -> bool:
        """Whether a formula definition contains a likely result marker."""
        text = "" if value is None else str(value)
        return any(word in text for word in RESULT_KEYWORDS)

    @staticmethod
    def _has_result_marker(value: Any) -> bool:
        """Only an output *starting* with a marker is an auto-extracted issue.

        A formula may generate a helper sentence such as “……请核实”, which is
        then consumed by a neighbouring result formula.  Matching a keyword
        anywhere in the sentence incorrectly turns that helper into an issue.
        Auto-discovered named-range formulas therefore use a strict output
        protocol: 错误/核实/提示等标识必须在结果开头.
        """
        text = "" if value is None else str(value).strip()
        return text.startswith(RESULT_KEYWORDS)

    @staticmethod
    def _issue_id(
        org_code: str,
        source_file: Path | str,
        rule: AuditRule,
        indicator: str = "",
    ) -> str:
        """Return a human-readable, location-based identity for history matching.

        The reporting date is stripped from the source workbook name so the
        same institution/report remains comparable across periods.
        """
        source_name = normalize_template_name(Path(str(source_file)).stem)
        parts = (
            source_name or "工作簿未识别",
            rule.sheet_name or "工作表未识别",
            rule.formula_cell or "公式单元格未识别",
            indicator or "校验指标未识别",
        )
        return "｜".join(parts)

    def write_navigation(self, workbook: Any, issues: list[Issue]) -> None:
        names = {
            workbook.Worksheets(index).Name
            for index in range(1, workbook.Worksheets.Count + 1)
        }
        if "审核导航" in names:
            workbook.Worksheets("审核导航").Delete()
        sheet = workbook.Worksheets.Add(Before=workbook.Worksheets(1))
        sheet.Name = "审核导航"
        headers = (
            "状态",
            "级别",
            "规则编号",
            "报表",
            "问题位置",
            "校验字段",
            "当前值",
            "对比值",
            "参考值",
            "差值",
            "详细说明",
            "连续期数",
            "首次出现期",
            "上次出现期",
            "问题说明",
            "公式结果",
            "机构反馈",
            "审核意见",
        )
        values = [headers]
        for item in issues:
            values.append(
                (
                    item.status,
                    item.severity,
                    item.rule_id,
                    item.report_code,
                    f"{item.sheet_name}!{item.target_cell}",
                    item.check_field,
                    _plain(item.target_value),
                    _plain(item.comparison_value),
                    _plain(item.reference_value),
                    _plain(item.difference_value),
                    item.detail,
                    item.consecutive_count,
                    item.first_seen_period,
                    item.previous_seen_period,
                    item.message,
                    _plain(item.formula_result),
                    item.institution_feedback,
                    item.auditor_opinion,
                )
            )
        if len(values) == 1:
            values.append(("本期未发现问题",) + ("",) * (len(headers) - 1))
        self._write_table(sheet, values, freeze=True)
        for row_index, item in enumerate(issues, start=2):
            sheet.Hyperlinks.Add(
                Anchor=sheet.Cells(row_index, 5),
                Address="",
                SubAddress=f"'{item.sheet_name}'!{item.target_cell}",
                TextToDisplay=f"{item.sheet_name}!{item.target_cell}",
            )

    def read_history(self, path: Path, *, sheet_name: str = HISTORY_AUDIT_SHEET) -> list[Issue]:
        if not path.exists():
            return []
        workbook = self.open_workbook(path, read_only=True)
        try:
            try:
                sheet = workbook.Worksheets(sheet_name)
            except Exception:
                # 旧 sheet 名（历史审核结果 / 历史审核记录）兼容读取。
                if sheet_name != HISTORY_AUDIT_SHEET:
                    return []
                sheet = None
                for legacy_name in ("历史审核结果", "历史审核记录"):
                    try:
                        sheet = workbook.Worksheets(legacy_name)
                        break
                    except Exception:
                        continue
                if sheet is None:
                    return []
            rows = _matrix(sheet.UsedRange.Value2)
            if len(rows) < 2:
                return []
            headers = [str(value).strip() if value is not None else "" for value in rows[0]]
            index = {header: i for i, header in enumerate(headers)}

            def get(row: list[Any], name: str) -> Any:
                i = index.get(name)
                return row[i] if i is not None and i < len(row) else ""

            result: list[Issue] = []
            for row in rows[1:]:
                source_file = str(get(row, "工作簿名") or get(row, "来源工作簿") or get(row, "源文件") or "")
                org_code = str(get(row, "机构代码") or "")
                sheet_name = str(get(row, "工作表名") or get(row, "工作表") or "")
                formula_cell = str(get(row, "公式单元格") or "")
                target_cell = str(get(row, "定位单元格") or "")
                legacy_identifier = str(
                    get(row, "规则编号") or get(row, "问题标识") or get(row, "问题ID") or ""
                ).strip()
                if not legacy_identifier and not any(
                    (source_file, org_code, sheet_name, formula_cell, target_cell)
                ):
                    continue
                indicator = str(get(row, "校验指标") or get(row, "校验字段") or "")
                issue_id = re.sub(r"｜\d+$", "", legacy_identifier) if legacy_identifier else self._issue_id(
                    org_code, source_file, AuditRule(
                        rule_id="", enabled=True, report_code="", sheet_name=sheet_name,
                        formula_cell=formula_cell or target_cell, target_cell=target_cell,
                        severity="错误", message="",
                    ), indicator,
                )
                formula_result = get(row, "公式结果")
                stored_message = str(get(row, "问题说明") or get(row, "描述") or "")
                stored_severity = str(get(row, "级别") or get(row, "错误类型") or "")
                parsed = parse_formula_result(
                    formula_result,
                    default_severity=stored_severity or "错误",
                    default_message=stored_message,
                )

                def stored_or_parsed(name: str, fallback: Any) -> Any:
                    stored = get(row, name)
                    return fallback if stored is None or stored == "" else stored

                detail = str(
                    stored_or_parsed(
                        "详细说明",
                        get(row, "描述") or get(row, "详细描述") or parsed.detail,
                    )
                    or ""
                )
                result.append(
                    Issue(
                        issue_id=issue_id,
                        period=str(get(row, "数据期") or ""),
                        batch_id=str(get(row, "批次号") or ""),
                        audit_time=str(get(row, "审核时间") or ""),
                        triggered=True,
                        status="",
                        first_seen_period="",
                        previous_seen_period="",
                        consecutive_count=1,
                        org_code=org_code,
                        org_name=str(get(row, "机构名称") or ""),
                        report_code=str(get(row, "报表代码") or ""),
                        sheet_name=sheet_name,
                        rule_id=str(get(row, "规则编号") or ""),
                        severity=stored_severity or str(get(row, "错误类型") or ""),
                        formula_cell=formula_cell or target_cell,
                        target_cell=target_cell,
                        target_value=get(row, "当前值"),
                        formula_result=formula_result,
                        message=stored_message or parsed.message,
                        source_file=source_file,
                        audit_file=str(get(row, "审核副本") or ""),
                        check_field=indicator or str(stored_or_parsed("校验字段", parsed.indicator) or ""),
                        comparison_value=stored_or_parsed(
                            "对比值", parsed.comparison_value
                        ),
                        reference_value=stored_or_parsed(
                            "参考值", parsed.reference_value
                        ),
                        difference_value=stored_or_parsed(
                            "差值", parsed.difference_value
                        ),
                        detail=detail,
                        institution_feedback=str(get(row, "历史校验说明") or get(row, "机构反馈") or ""),
                        auditor_opinion=str(get(row, "审核意见") or ""),
                    )
                )
            return result
        finally:
            self.close_workbook(workbook)

    def write_history(
        self, path: Path, issues: list[Issue], *, sheet_name: str = "问题历史"
    ) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        workbook = None
        try:
            if path.exists():
                workbook = self.open_workbook(path, read_only=False)
                names = set(self._worksheet_names(workbook))
                if sheet_name in names:
                    sheet = workbook.Worksheets(sheet_name)
                    sheet.Cells.Clear()
                else:
                    sheet = workbook.Worksheets.Add(
                        After=workbook.Worksheets(workbook.Worksheets.Count)
                    )
                    sheet.Name = sheet_name
            else:
                workbook = self.excel.Workbooks.Add()
                sheet = workbook.Worksheets(1)
                sheet.Name = sheet_name
            values = [HISTORY_HEADERS]
            values.extend(self._issue_history_row(item) for item in issues)
            self._write_table(sheet, values, freeze=True)
            if path.exists():
                workbook.Save()
            else:
                workbook.SaveAs(str(path.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
        finally:
            if workbook is not None:
                self.close_workbook(workbook)

    def write_preflight_report(
        self,
        path: Path,
        template_path: Path,
        template_items: list[PreflightItem],
        source_matches: list[tuple[Path, SourceMatch]],
    ) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        workbook = self.excel.Workbooks.Add()
        try:
            while workbook.Worksheets.Count < 2:
                workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
            while workbook.Worksheets.Count > 2:
                workbook.Worksheets(workbook.Worksheets.Count).Delete()
            health_sheet = workbook.Worksheets(1)
            health_sheet.Name = "模板体检"
            match_sheet = workbook.Worksheets(2)
            match_sheet.Name = "文件匹配"

            health_rows: list[tuple[Any, ...]] = [
                ("模板文件", "类别", "级别", "状态", "工作表", "位置", "说明")
            ]
            health_rows.extend(
                (
                    str(template_path),
                    item.category,
                    item.level,
                    item.status,
                    item.sheet_name or "—",
                    item.location or "—",
                    item.message,
                )
                for item in template_items
            )
            match_rows: list[tuple[Any, ...]] = [
                (
                    "报送文件",
                    "匹配结果",
                    "匹配率",
                    "匹配标签数",
                    "检查标签数",
                    "缺少工作表",
                    "说明",
                )
            ]
            match_rows.extend(
                (
                    str(source_path),
                    "通过" if result.matched else "不通过",
                    f"{result.score:.0%}",
                    result.matched_labels,
                    result.checked_labels,
                    "、".join(result.missing_sheets) or "无",
                    result.details,
                )
                for source_path, result in source_matches
            )
            if not source_matches:
                match_rows.append(
                    ("尚未检查报送文件", "—", "—", "—", "—", "—", "—")
                )
            self._write_table(health_sheet, health_rows, freeze=True)
            self._write_table(match_sheet, match_rows, freeze=True)
            workbook.SaveAs(str(path.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
        finally:
            self.close_workbook(workbook)

    def write_summary(
        self,
        path: Path,
        current: list[Issue],
        *,
        sheet_name: str = "本期审核结果",
    ) -> None:
        workbook = self.excel.Workbooks.Add()
        try:
            while workbook.Worksheets.Count > 1:
                workbook.Worksheets(workbook.Worksheets.Count).Delete()
            current_sheet = workbook.Worksheets(1)
            current_sheet.Name = re.sub(r"[\\[\\]:*?/\\\\]", "_", sheet_name or "本期审核结果")[:31] or "本期审核结果"

            self._write_table(
                current_sheet,
                [HISTORY_HEADERS]
                + [self._issue_history_row(item) for item in current],
                freeze=True,
            )
            # “定位单元格”直接跳转至对应机构的审核副本。工作表名另列保留，
            # 超链接显示仍只显示单元格地址，方便筛选和复制。
            for row_index, item in enumerate(current, start=2):
                if not item.audit_file or not item.sheet_name or not item.target_cell:
                    continue
                audit_path = Path(item.audit_file)
                if not audit_path.is_file():
                    continue
                sheet_name = item.sheet_name.replace("'", "''")
                current_sheet.Hyperlinks.Add(
                    Anchor=current_sheet.Cells(row_index, 3),
                    Address=str(audit_path.resolve()),
                    SubAddress=f"'{sheet_name}'!{item.target_cell}",
                    TextToDisplay=item.target_cell,
                )
            workbook.SaveAs(str(path.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
        finally:
            self.close_workbook(workbook)

    @staticmethod
    def _issue_history_row(item: Issue) -> tuple[Any, ...]:
        return (
            Path(item.source_file).name if item.source_file else "",
            item.sheet_name,
            item.target_cell,
            item.severity,
            item.check_field,
            item.detail or item.message,
            _plain(item.target_value),
            _plain(item.comparison_value),
            _plain(item.difference_value),
            item.issue_id,
            item.institution_feedback,
            item.auditor_opinion,
        )

    @staticmethod
    def _issue_result_row(item: Issue) -> tuple[Any, ...]:
        return (
            Path(item.source_file).name if item.source_file else "",
            item.sheet_name,
            item.target_cell,
            item.severity,
            _plain(item.comparison_value if item.comparison_value != "" else item.target_value),
            _plain(item.reference_value),
            _plain(item.difference_value),
            item.detail or item.message,
        )

    @staticmethod
    def _write_table(sheet: Any, values: list[tuple[Any, ...]], *, freeze: bool) -> None:
        if not values:
            return
        width = max(len(row) for row in values)
        normalized = []
        for row in values:
            padded = tuple(row) + ("",) * (width - len(row))
            # COM writes an empty string as a shared-string cell. Excel displays
            # it as blank, but some non-Excel readers can misread the shared
            # string index as a number. None creates a real blank cell instead.
            normalized.append(
                tuple(None if value == "" else value for value in padded)
            )
        target = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(normalized), width))
        target.NumberFormat = "@"
        target.Value2 = tuple(normalized)
        header = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, width))
        header.Font.Bold = True
        header.Interior.Color = 0xD9EAD3
        target.AutoFilter()
        target.EntireColumn.AutoFit()
        for column in range(1, width + 1):
            if sheet.Columns(column).ColumnWidth > 40:
                sheet.Columns(column).ColumnWidth = 40
        if freeze:
            sheet.Activate()
            sheet.Range("A2").Select()
            sheet.Application.ActiveWindow.FreezePanes = True
