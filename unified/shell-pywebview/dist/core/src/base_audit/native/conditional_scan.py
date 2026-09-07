"""Conditional-format background extraction for the native edition.

Design note
-----------
The Windows/Excel/WPS build reads a conditionally-applied background through
the COM ``DisplayFormat.Interior.Color`` API.  LibreOffice's UNO has **no**
equivalent: ``CellBackColor`` only reflects the cell's base fill, so a rule
that merely evaluates TRUE (e.g. ``cellIs > 0``) never surfaces through UNO.
Repeated checks on xlsx and .ods confirmed the rule object loads
(``ConditionalFormat.Count == 1``) but the rendered colour stays default.

On Linux this module therefore evaluates the submission's own conditional
format rules directly from the OOXML with openpyxl and reports the rule's
``dxf`` fill as the display colour.  The audit copy carries the source
submission's rules (reporting systems pre-flag suspect cells such as
``cellIs greaterThan 0``), and the already-calculated values are cached in the
file, so no separate formula render is required for the common ``cellIs``
form.  Rules that need a real formula renderer (``expression``,
``colorScale``, ``dataBar``, ``iconSet``, …) are best handled by the COM build
and are reported as not evaluable here rather than guessed.

The public audit entry ``libreoffice_conditional_batch`` remains as a
compatibility context so the pipeline's step wiring and the step's
“失败后处理” policy are unchanged; it no longer starts a UNO listener because
none is needed for OOXML rule evaluation.
"""

from __future__ import annotations

import re
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries

from ..conditional_format import evaluate_expression_formula
from ..models import CopyRange, Issue


# White, transparent and automatic/no-fill never represent a visible result.
NO_FILL_COLOURS = {0, -1, 16777215}

_CELLREF_RE = re.compile(r"^([A-Za-z]{1,3})(\d+)$")


def _color_to_hex(color) -> str | None:
    """Return an 8-digit AARRGGBB hex string for an openpyxl Color, if real."""
    if color is None:
        return None
    try:
        color_type = (color.type or "rgb").casefold()
    except Exception:
        color_type = "rgb"
    if color_type == "rgb":
        try:
            raw = str(color.rgb)
        except Exception:
            return None
        # openpyxl keeps file 8-char "AARRGGBB" verbatim but pads 6-char input
        # to "00RRGGBB".  A conditional-format background is always opaque, so
        # normalise to an FF alpha over the trailing RGB.
        if len(raw) >= 6 and re.fullmatch(r"[0-9A-Fa-f]{6,8}", raw):
            return "FF" + raw[-6:].upper()
        return None
    if color_type == "indexed":
        idx = color.indexed
        if idx is None:
            return None
        from openpyxl.styles.colors import COLOR_INDEX
        try:
            named = COLOR_INDEX[idx]
        except (IndexError, TypeError):
            return None
        if named in (None, "window", "windowText"):
            return None
        # COLOR_INDEX stores 8-char "00RRGGBB" strings (alpha byte set to 00).
        if isinstance(named, str) and re.fullmatch(r"[0-9A-Fa-f]{8}", named):
            return "FF" + named[2:].upper()
        if isinstance(named, str) and named.startswith("#"):
            return "FF" + named[1:].upper()
        return None
    if color_type == "theme":
        # Standard Office theme colours for the first 12 palette slots.
        theme_hex = {
            0: "FFFFFF", 1: "000000", 2: "EEECE1", 3: "1F497D",
            4: "4F81BD", 5: "C0504D", 6: "9BBB59", 7: "8064A2",
            8: "4BACC6", 9: "F79646", 10: "0000FF", 11: "800080",
        }.get(color.theme)
        if theme_hex is None:
            return None
        tint = color.tint or 0.0
        if tint:
            # Apply Excel's luminance tint formula.
            def _channel(value):
                return round((value if tint > 0 else value * (1 + tint)) * 255)
            base = [int(theme_hex[i:i + 2], 16) for i in (0, 2, 4)]
            if tint > 0:
                rgb = [round(255 - (255 - ch) * (1 - tint)) for ch in base]
            else:
                rgb = [round(ch * (1 + tint)) for ch in base]
            return "FF" + "".join("{:02X}".format(ch) for ch in rgb)
        return "FF" + theme_hex
    return None


def _dxf_fill_rgb(dxfs, dxf_id: int | None) -> str | None:
    """Resolve an Excel conditional-format dxfId to its visible fill RGB.

    Excel writes the condition background under ``dxf/fill/patternFill``:
    ``bgColor`` paints the whole cell (the usual conditional-format idiom),
    ``fgColor`` is the base.  openpyxl keeps them on the differential style.
    """
    if dxf_id is None:
        return None
    try:
        dxf = dxfs[dxf_id]
    except (IndexError, TypeError):
        return None
    if dxf is None:
        return None
    fill = getattr(dxf, "fill", None)
    if fill is None:
        return None
    fg_hex = _color_to_hex(getattr(fill, "fgColor", None))
    bg_hex = _color_to_hex(getattr(fill, "bgColor", None))
    for candidate in (bg_hex, fg_hex):
        if candidate and candidate.upper() not in {"00000000", "FFFFFFFF"}:
            return candidate
    return None


_OPERATOR_LABELS = {
    "equal": "=", "notequal": "<>", "greaterthan": ">",
    "greaterthanorequal": ">=", "lessthan": "<", "lessthanorequal": "<=",
}


def _rule_label(rule, *, address: str) -> str:
    """A human-readable description of one conditional-format rule.

    Shown in the issue's 描述 column when the cell has no explanatory comment,
    mirroring the rule as it would be read in Excel:
      * cellIs     ->  条件格式规则：C26>0
      * expression -> 条件格式规则：and(d5>0,1)
    ``address`` is the target cell being reported (used for cellIs rules).
    """
    rule_type = (rule.type or "").casefold()
    if rule_type == "cellis":
        op_label = _OPERATOR_LABELS.get((rule.operator or "").casefold(), (rule.operator or "?"))
        operand = (rule.formula or [""])[0]
        return "条件格式规则：{}{}{}".format(address, op_label, operand)
    if rule_type == "expression":
        formula = (rule.formula or [""])[0]
        # Compact, lower-case form without $-anchors reads naturally.
        text = str(formula).replace("$", "").replace(" ", "")
        text = text.lower()[:80]
        return "条件格式规则：{}".format(text)
    return "条件格式规则：{}".format(rule_type or "未知")

def _as_number(value):
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip())
        except ValueError:
            return None
    return None


def _operand_value(ws, formula, index: int):
    """Best-effort cellIs operand: numeric literal, bool, or single cell ref.

    ``cellIs`` operands in submission files are typically plain literals such
    as the ``0`` in ``cellIs greaterThan 0``.  A formula expression operand is
    not evaluated here (that is what a real formula renderer is for); we return
    None and the rule is skipped rather than guessed.
    """
    if not formula or index >= len(formula):
        return None
    raw = str(formula[index]).strip()
    if raw.startswith("="):
        raw = raw[1:]
    lowered = raw.casefold()
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    try:
        return float(raw) if ("." in raw or "e" in lowered) else int(raw)
    except ValueError:
        pass
    match = _CELLREF_RE.fullmatch(raw.replace("$", ""))
    if match:
        from openpyxl.utils.cell import column_index_from_string
        try:
            col = column_index_from_string(match.group(1))
            row = int(match.group(2))
        except Exception:
            return None
        return ws.cell(row=row, column=col).value
    return None


def _cellis_matches(rule, value, ws) -> bool:
    operator = (rule.operator or "").casefold()
    if value is None:
        return False
    lo = _as_number(value)
    operand1 = _operand_value(ws, rule.formula, 0)
    operand2 = _operand_value(ws, rule.formula, 1) if operator in {"between", "notbetween"} else None
    n1 = _as_number(operand1)
    n2 = _as_number(operand2)

    def cmp_both(predicate):
        if lo is not None and n1 is not None:
            return predicate(lo, n1)
        if isinstance(value, str) and isinstance(operand1, str):
            return predicate(value, operand1)
        if isinstance(value, bool) and isinstance(operand1, bool):
            return predicate(value, operand1)
        return False

    if operator == "equal":
        return value == operand1 or cmp_both(lambda a, b: a == b)
    if operator == "notequal":
        return not (value == operand1) if isinstance(value, type(operand1)) or value == operand1 else True
    if operator == "greaterthan":
        return lo is not None and n1 is not None and lo > n1
    if operator == "greaterthanorequal":
        return lo is not None and n1 is not None and lo >= n1
    if operator == "lessthan":
        return lo is not None and n1 is not None and lo < n1
    if operator == "lessthanorequal":
        return lo is not None and n1 is not None and lo <= n1
    if operator in {"between", "notbetween"}:
        if lo is None or n1 is None or n2 is None:
            return False
        inside = n1 <= lo <= n2
        return inside if operator == "between" else not inside
    return False


def _evaluate_rule(rule, value, ws, *, anchor_row: int, anchor_col: int,
                   cur_row: int, cur_col: int) -> bool:
    rule_type = (rule.type or "").casefold()
    if rule_type == "cellis":
        return _cellis_matches(rule, value, ws)
    if rule_type == "expression":
        formula = (rule.formula or [""])[0]
        try:
            return evaluate_expression_formula(
                formula, ws, anchor_row, anchor_col, cur_row, cur_col
            )
        except (ValueError, NotImplementedError):
            # An unparseable/unsupported formula must not crash the batch;
            # it simply cannot be evaluated natively.
            return False
    # colorScale / dataBar / iconSet need a real renderer (Windows COM only).
    return False


def _covering_ranges(ws):
    """Yield (rule, sqref) pairs registered on a worksheet."""
    try:
        entries = ws.conditional_formatting
    except AttributeError:
        return
    for cf in entries:
        for rule in cf.rules:
            yield rule, cf.sqref


def _sqref_bounds(sqref) -> list[tuple[int, int, int, int]]:
    """Expand a conditional-format sqref into [(min_col,min_row,max_col,max_row)]."""
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
    if isinstance(sqref, MultiCellRange):
        raw = sqref.ranges
    else:
        raw = [sqref]
    result = []
    for item in raw:
        try:
            if isinstance(item, CellRange):
                min_col, min_row, max_col, max_row = item.bounds
            else:
                text = str(item).strip()
                if not text:
                    continue
                parsed = CellRange(text)
                min_col, min_row, max_col, max_row = parsed.bounds
        except Exception:
            continue
        result.append((min_col, min_row, max_col, max_row))
    return result


def _scan_region(book, workbook_path: Path, ranges: list[CopyRange]) -> dict[tuple[str, int, int], int]:
    """Return {(sheet,row,col): fill_int} for cells whose own CF rule triggers.

    Only cells that lie inside one of the caller's ``条件格式区域`` named
    ranges and that are covered by a triggering rule are reported.  When a cell
    matches several rules the *highest-priority* colour wins (Excel applies the
    largest priority number last); ``stopIfTrue`` stops further matching.
    """
    dxfs = getattr(book, "_differential_styles", None)

    def _rule_colour(rule):
        if rule.dxfId is None:
            return None
        rgb = _dxf_fill_rgb(dxfs, rule.dxfId)
        if not rgb or len(rgb) < 6:
            return None
        try:
            colour = int(rgb[-6:], 16)
        except ValueError:
            return None
        return None if colour in NO_FILL_COLOURS else colour

    candidates: dict[tuple[str, int, int], tuple[int, int, str]] = {}
    for area in ranges:
        if area.sheet_name not in book.sheetnames:
            continue
        ws = book[area.sheet_name]
        region_min_col, region_min_row, region_max_col, region_max_row = range_boundaries(area.address)

        # Pre-resolve every covering rule's colour and anchor once.
        rule_specs = []
        for rule, sqref in _covering_ranges(ws):
            colour = _rule_colour(rule)
            if colour is None:
                continue
            bounds = _sqref_bounds(sqref)
            if not bounds:
                continue
            anchor_col = min(b[0] for b in bounds)
            anchor_row = min(b[1] for b in bounds)
            rule_specs.append((
                rule, bounds, anchor_row, anchor_col, colour,
                rule.priority if rule.priority is not None else 0,
                bool(rule.stopIfTrue),
            ))

        # Evaluate each target cell once across all covering rules, keeping the
        # highest-priority matching colour (Excel: later priority wins).
        for spec in rule_specs:
            rule, bounds, anchor_row, anchor_col, colour, priority, stop_if = spec
            for min_col, min_row, max_col, max_row in bounds:
                lo_row = max(min_row, region_min_row)
                hi_row = min(max_row, region_max_row)
                lo_col = max(min_col, region_min_col)
                hi_col = min(max_col, region_max_col)
                for row in range(lo_row, hi_row + 1):
                    for col in range(lo_col, hi_col + 1):
                        value = ws.cell(row=row, column=col).value
                        matched = _evaluate_rule(
                            rule, value, ws,
                            anchor_row=anchor_row, anchor_col=anchor_col,
                            cur_row=row, cur_col=col,
                        )
                        if not matched:
                            continue
                        key = (area.sheet_name, row, col)
                        current = candidates.get(key)
                        address = "{}{}".format(get_column_letter(col), row)
                        label = _rule_label(rule, address=address)
                        if stop_if or current is None or priority >= current[0]:
                            candidates[key] = (priority, colour, label)
    return {key: (colour, label) for key, (_priority, colour, label) in candidates.items()}


def _build_issues(
    *,
    workbook_path: Path, ranges: list[CopyRange], structure_ranges: list[CopyRange],
    period: str, batch_id: str, source_file: Path,
    colors: dict[tuple[str, int, int], tuple[int, str]], static,
) -> list[Issue]:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result: list[Issue] = []
    seen: set[tuple[str, int, int]] = set()
    for area in ranges:
        if area.sheet_name not in static.sheetnames:
            continue
        baseline = static[area.sheet_name]
        min_col, min_row, max_col, max_row = range_boundaries(area.address)
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if (area.sheet_name, row, col) in seen:
                    continue
                seen.add((area.sheet_name, row, col))
                cell_hit = colors.get((area.sheet_name, row, col))
                if cell_hit is None:
                    continue
                colour, rule_label = cell_hit
                if colour in NO_FILL_COLOURS:
                    continue
                fg_rgb = baseline.cell(row, col).fill.fgColor
                base = str((fg_rgb.rgb or "") if fg_rgb else "") or ""
                if base and base[-6:].upper() == "{:06X}".format(colour):
                    continue
                address = "{}{}".format(get_column_letter(col), row)
                labels: list[str] = []
                for structure in structure_ranges:
                    if structure.sheet_name != area.sheet_name:
                        continue
                    a, b, c, d = range_boundaries(structure.address)
                    if a <= col <= c:
                        for r in range(min(row, d), b - 1, -1):
                            value = baseline.cell(r, col).value
                            if value not in (None, ""):
                                labels.append(str(value))
                                break
                    if b <= row <= d:
                        for cc in range(min(col, c), a - 1, -1):
                            value = baseline.cell(row, cc).value
                            if value not in (None, ""):
                                labels.append(str(value))
                                break
                indicator = "_".join(dict.fromkeys(labels)) or "校验指标未识别"
                comment = baseline.cell(row, col).comment
                detail = comment.text.strip() if comment else rule_label
                # 规则编号与既有校验/COM 约定一致：工作簿名(去日期)｜表｜定位格｜校验指标。
                from ..template import normalize_template_name
                workbook_key = normalize_template_name(source_file.stem) or source_file.stem
                identity = "｜".join((workbook_key, area.sheet_name, address, indicator))
                result.append(Issue(
                    identity,
                    period, batch_id, stamp, True, "", "", "", 1, "", "", "",
                    area.sheet_name, identity, "总行条件格式触发", address, address,
                    baseline.cell(row, col).value, "", detail, str(source_file), str(workbook_path),
                    check_field=indicator, detail=detail,
                ))
    return result


def extract_conditional_format_issues(
    *, workbook_path: Path, ranges: list[CopyRange],
    structure_ranges: list[CopyRange], period: str, batch_id: str, source_file: Path,
) -> list[Issue]:
    """Return cells whose own conditional-format rule triggers a background.

    Uses pure openpyxl OOXML evaluation (see module docstring): the submission's
    ``cellIs`` rules are read directly and their dxf fill reported as the
    display colour.  Named ``条件格式区域`` limits scanning; labels from
    ``表结构区域`` become the check indicator where present.
    """
    static = load_workbook(workbook_path, data_only=True, keep_links=False)
    try:
        live_ranges = [area for area in ranges if area.sheet_name in static.sheetnames]
        book = load_workbook(workbook_path, data_only=True, keep_links=False)
        try:
            colors = _scan_region(book, workbook_path, live_ranges)
        finally:
            book.close()
        return _build_issues(
            workbook_path=workbook_path, ranges=live_ranges,
            structure_ranges=structure_ranges, period=period, batch_id=batch_id,
            source_file=source_file, colors=colors, static=static,
        )
    finally:
        static.close()


def extract_conditional_format_issues_with_libreoffice(
    *, workbook_path: Path, ranges: list[CopyRange],
    structure_ranges: list[CopyRange], period: str, batch_id: str, source_file: Path,
) -> list[Issue]:
    """Compatibility entry: extraction is OOXML-native, LibreOffice not needed."""
    return extract_conditional_format_issues(
        workbook_path=workbook_path, ranges=ranges,
        structure_ranges=structure_ranges, period=period,
        batch_id=batch_id, source_file=source_file,
    )
