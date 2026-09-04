"""Resolve the worksheets an auxiliary workbook contributes to an audit copy."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable


_SHEET_REF = re.compile(r"(?:'(?P<quoted>[^']+)'|(?P<plain>[A-Za-z0-9_一-鿿]+))!")


@dataclass(frozen=True)
class ExternalSheetPlan:
    sheet_names: tuple[str, ...]
    source: str  # 公式识别 / 全部工作表


def referenced_external_sheets(
    formulas: Iterable[object], available_sheets: Iterable[str]
) -> tuple[str, ...]:
    """Find references such as ``参照表!A:A`` in template formulas."""
    available = tuple(available_sheets)
    index = {name.casefold(): name for name in available}
    found: list[str] = []
    for formula in formulas:
        if not isinstance(formula, str):
            continue
        for match in _SHEET_REF.finditer(formula):
            name = (match.group("quoted") or match.group("plain") or "").strip()
            actual = index.get(name.casefold())
            if actual and actual not in found:
                found.append(actual)
    return tuple(found)


def make_external_sheet_plan(
    *,
    formulas: Iterable[object],
    available_sheets: Iterable[str],
) -> ExternalSheetPlan:
    """按模板公式识别要复制的工作表；识别不到时复制外部文件全部工作表。"""
    detected = referenced_external_sheets(formulas, available_sheets)
    if detected:
        return ExternalSheetPlan(detected, "公式识别")
    return ExternalSheetPlan(tuple(available_sheets), "全部工作表")
