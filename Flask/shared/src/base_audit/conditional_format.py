"""OOXML ``expression`` conditional-format evaluator.

Ported from the native UOS edition (and shared with the Flet Windows build) so
that WPS — which cannot return a rendered ``DisplayFormat`` colour through COM
— can still evaluate the submission's own conditional-format rules directly
from OOXML with openpyxl.

Supported subset (matches every shape observed in the real templates):
  AND / OR / NOT / ABS / SUM, comparisons = <> < > <= >=, arithmetic + - * /,
  numeric & string literals incl. "", single cells and A1:B2 ranges, with
  $ -locking.  Values are read from the audit copy's already-calculated
  (data_only) cache, so formula cells referenced by a rule already carry the
  computed result.
"""

from __future__ import annotations

import re
import operator as _operator


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
