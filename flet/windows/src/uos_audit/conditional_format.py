"""Openpyxl conditional-format extraction and expression evaluator."""
from __future__ import annotations
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.formula.translate import Translator
import re
import operator as _operator
from .models import Issue, CopyRange

_CELL_REF = re.compile(r"(?<![A-Z0-9_!])(?P<ref>\$?[A-Z]{1,3}\$?\d+)(?![A-Z0-9_])", re.I)


def _python_condition(formula: str, sheet, origin: str, destination: str) -> bool | None:
    """Evaluate the deliberately small TRUE/FALSE rule subset without UNO.

    This covers the ordinary expression rules used by the templates: cell
    comparisons, arithmetic, AND/OR/NOT, ABS and ISBLANK.  Unknown formulas
    return ``None`` rather than guessing a result.
    """
    try:
        text = Translator(formula, origin=origin).translate_formula(destination)
    except Exception:
        text = formula
    if "!" in text or ":" in text:
        return None
    text = text.lstrip("=").replace("<>", "!=").replace("^", "**")
    text = re.sub(r"(?<![<>=])=(?!=)", "==", text)
    text = re.sub(r"\bTRUE\b", "True", text, flags=re.I)
    text = re.sub(r"\bFALSE\b", "False", text, flags=re.I)
    text = re.sub(r"\bAND\s*\(", "_and(", text, flags=re.I)
    text = re.sub(r"\bOR\s*\(", "_or(", text, flags=re.I)
    text = re.sub(r"\bNOT\s*\(", "_not(", text, flags=re.I)
    text = re.sub(r"\bABS\s*\(", "abs(", text, flags=re.I)
    text = re.sub(r"\bISBLANK\s*\(", "_blank(", text, flags=re.I)
    text = _CELL_REF.sub(lambda match: "_cell({!r})".format(match.group("ref").replace("$", "")), text)
    def cell_value(address: str):
        value = sheet[address].value
        return 0 if value is None else value
    try:
        value = eval(text, {"__builtins__": {}}, {
            "_cell": cell_value, "_and": lambda *items: all(items), "_or": lambda *items: any(items),
            "_not": lambda item: not item, "_blank": lambda item: item in (None, ""), "abs": abs,
        })
    except Exception:
        return None
    return bool(value)


def extract_simple_conditional_format_issues(*, workbook_path: Path, ranges: list[CopyRange], structure_ranges: list[CopyRange], period: str, batch_id: str, source_file: Path) -> tuple[list[Issue], int]:
    """Extract TRUE/FALSE conditional-format rules with openpyxl only.

    Returns ``(issues, unsupported_rule_count)``.  It intentionally does not
    infer rendered colour, data bars, icon sets or colour scales.
    """
    book = load_workbook(workbook_path, read_only=False, data_only=False, keep_links=False)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    issues: list[Issue] = []; unsupported = 0; seen: set[tuple[str, str]] = set()
    try:
        for area in ranges:
            if area.sheet_name not in book.sheetnames:
                continue
            sheet = book[area.sheet_name]
            left, top, right, bottom = range_boundaries(area.address)
            for conditional in sheet.conditional_formatting:
                rules = conditional.rules
                try:
                    origin = str(next(iter(conditional.sqref.ranges))).split(":", 1)[0].replace("$", "")
                except Exception:
                    unsupported += len(rules); continue
                for rule in rules:
                    if rule.type != "expression" or not rule.formula:
                        unsupported += 1; continue
                    for row in range(top, bottom + 1):
                        for col in range(left, right + 1):
                            address = "{}{}".format(get_column_letter(col), row)
                            if address not in conditional.sqref:
                                continue
                            triggered = _python_condition(str(rule.formula[0]), sheet, origin, address)
                            if triggered is None:
                                unsupported += 1; break
                            if not triggered or (area.sheet_name, address) in seen:
                                continue
                            seen.add((area.sheet_name, address))
                            cell = sheet[address]; comment = cell.comment
                            detail = comment.text.strip() if comment else "条件格式规则已触发，请核实"
                            issues.append(Issue("｜".join((source_file.stem, area.sheet_name, address, "条件格式规则")), period, batch_id, stamp, True, "", "", "", 1, "", "", "", area.sheet_name, "条件格式规则", "总行条件格式触发", address, address, cell.value, str(rule.formula[0]), detail, str(source_file), str(workbook_path), check_field="条件格式规则", detail=detail))
    finally:
        book.close()
    return issues, unsupported


# ---------------------------------------------------------------------------
# Excel `expression` conditional-format evaluator (verified, shared).
#
# Ported from the native UOS edition.  Banking submissions encode most
# conditional formats as formulas whose first operand is the range's top-left
# (anchor) cell; Excel re-evaluates the formula for every cell in the range,
# translating *relative* references by the row/col offset from that anchor.
#
# Supported subset (matches every shape observed in the real templates):
#   AND / OR / NOT / ABS / SUM, comparisons = <> < > <= >=, arithmetic + - * /,
#   numeric & string literals incl. "", single cells and A1:B2 ranges, with
#   $ -locking.  Values are read from the audit copy's already-calculated
#   (data_only) cache, so formula cells referenced by a rule already carry the
#   computed result.
# ---------------------------------------------------------------------------

_EXPR_TOKEN_RE = re.compile(r"""\s*(?:
   (?P<num>\d+(?:\.\d+)?) |
   (?P<str>"(?:[^"])*") |
   (?P<ref>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?) |
   (?P<name>[A-Za-z_][A-Za-z0-9_.]*) |
   (?P<op><>|<=|>=|[<>+\-*/=()%,])
 )""", re.VERBOSE)

_EXPR_ARITH = {'+': _operator.add, '-': _operator.sub, '*': _operator.mul, '/': _operator.truediv}


def _expr_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        try:
            if text.endswith('%'):
                return float(text[:-1]) / 100.0
            return float(text)
        except ValueError:
            return None
    return None


def _expr_truthy(value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value != ""
    return value is not None


class _ExprRefs:
    """Resolve cell / range references inside an expression rule."""

    def __init__(self, ws, anchor_row: int, anchor_col: int):
        self.ws = ws
        self.ar = anchor_row
        self.ac = anchor_col

    @staticmethod
    def _coord(ref: str):
        col_abs = ref.startswith("$")
        rest = ref.lstrip("$")
        match = re.match(r"([A-Za-z]+)\$?(\d+)$", rest)
        col_letters, row_text = match.group(1), match.group(2)
        col = 0
        for ch in col_letters.upper():
            col = col * 26 + (ord(ch) - 64)
        row = int(row_text)
        row_abs = ref.count("$") >= 2
        return col, row, col_abs, row_abs

    def cell_value(self, ref, dr, dc):
        col, row, col_abs, row_abs = self._coord(ref)
        col = col if col_abs else col + dc
        row = row if row_abs else row + dr
        return self.ws.cell(row=row, column=col).value

    def range_values(self, ref, dr, dc):
        a, b = ref.split(":")
        c1, r1, c1a, r1a = self._coord(a)
        c2, r2, c2a, r2a = self._coord(b)
        if not c1a:
            c1 += dc
        if not c2a:
            c2 += dc
        if not r1a:
            r1 += dr
        if not r2a:
            r2 += dr
        lo_c, hi_c = sorted((c1, c2))
        lo_r, hi_r = sorted((r1, r2))
        return [
            self.ws.cell(row=row, column=col).value
            for row in range(lo_r, hi_r + 1)
            for col in range(lo_c, hi_c + 1)
        ]


def _expr_call(name, args):
    if name == "AND":
        return all(_expr_truthy(a) for a in args)
    if name == "OR":
        return any(_expr_truthy(a) for a in args)
    if name == "NOT":
        return not _expr_truthy(args[0]) if args else False
    if name == "ABS":
        v = _expr_number(args[0])
        return abs(v) if v is not None else None
    if name == "SUM":
        total = 0.0
        for arg in args:
            if isinstance(arg, list):
                for item in arg:
                    n = _expr_number(item)
                    if n is not None:
                        total += n
            else:
                n = _expr_number(arg)
                if n is not None:
                    total += n
        return total
    raise ValueError("暂不支持的条件格式函数：{}".format(name))


def evaluate_expression_formula(formula: str, ws, anchor_row: int, anchor_col: int,
                                cur_row: int, cur_col: int):
    """Evaluate an Excel CF ``expression`` at one cell (anchor-translated).

    ``ws`` may be any object exposing ``cell(row=…, column=…).value`` (an
    openpyxl worksheet).  Relative references are shifted by ``cur`` - ``anchor``
    while ``$``-locked references stay fixed, so a rule anchored at C5 and
    applied to C26 reads C26 rather than C5.
    """
    dr = cur_row - anchor_row
    dc = cur_col - anchor_col
    refs = _ExprRefs(ws, anchor_row, anchor_col)
    tokens = list(_EXPR_TOKEN_RE.finditer(formula))
    pos = [0]

    def peek():
        return tokens[pos[0]].group() if pos[0] < len(tokens) else None

    def peek_kind():
        return tokens[pos[0]].lastgroup if pos[0] < len(tokens) else None

    def eat():
        token = tokens[pos[0]]
        pos[0] += 1
        return token.group()

    def fail():
        raise ValueError("无法解析的条件格式公式：{}".format(formula))

    def expr():
        left = term()
        while peek_kind() == "op" and peek() in ("<>", "<=", ">=", "<", ">", "="):
            op = eat()
            right = term()
            left = _expr_compare(op, left, right)
        return left

    def term():
        left = factor()
        while peek_kind() == "op" and peek() in ("+", "-"):
            op = eat()
            right = factor()
            if isinstance(left, str) or isinstance(right, str):
                if op == "+":
                    left = str(left or "") + str(right or "")
                else:
                    left = None
            else:
                left = _EXPR_ARITH[op](_expr_number(left) or 0.0, _expr_number(right) or 0.0)
        return left

    def factor():
        left = unary()
        while peek_kind() == "op" and peek() in ("*", "/"):
            op = eat()
            right = unary()
            ln, rn = _expr_number(left), _expr_number(right)
            if ln is None or rn is None:
                left = None
            elif op == "/" and rn == 0:
                left = None
            else:
                left = _EXPR_ARITH[op](ln, rn)
        return left

    def unary():
        if peek_kind() == "op" and peek() == "-":
            eat()
            return -(_expr_number(unary()) or 0.0)
        if peek() == "(":
            eat()
            value = expr()
            if peek() != ")":
                fail()
            eat()
            return value
        kind = peek_kind()
        if kind == "num":
            raw = eat()
            return float(raw) if "." in raw else int(raw)
        if kind == "str":
            return eat()[1:-1]
        if kind == "ref":
            ref = eat()
            if ":" in ref:
                return refs.range_values(ref, dr, dc)
            return refs.cell_value(ref, dr, dc)
        if kind == "name":
            name = eat().upper()
            if name == "TRUE":
                return True
            if name == "FALSE":
                return False
            if peek() != "(":
                fail()
            eat()
            args = []
            if peek() != ")":
                while True:
                    args.append(expr())
                    if peek() == ",":
                        eat()
                    else:
                        break
            if peek() != ")":
                fail()
            eat()
            return _expr_call(name, args)
        fail()

    def _expr_compare(op, a, b):
        a_empty = a is None or a == ""
        b_empty = b is None or b == ""
        if op == "=":
            if a_empty and b_empty:
                return True
            if a_empty or b_empty:
                return False
            return a == b
        if op == "<>":
            if a_empty and b_empty:
                return False
            if a_empty or b_empty:
                return True
            return a != b
        na = _expr_number(0 if a_empty else a)
        nb = _expr_number(0 if b_empty else b)
        if na is None or nb is None:
            return False
        return {"<": na < nb, ">": na > nb, "<=": na <= nb, ">=": na >= nb}[op]

    value = expr()
    return _expr_truthy(value)
