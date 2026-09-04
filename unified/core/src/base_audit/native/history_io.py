"""openpyxl读写历史审核工作簿（native 管线，无 COM 依赖）。

对齐 core history.py 的表名与表头（HISTORY_AUDIT_SHEET / HISTORY_HEADERS），
只重写程序管理的历史表；人工维护的“历史校验说明/审核意见”按规则身份带回
本期结果，更新时永不覆盖。
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook

from ..history import HISTORY_AUDIT_SHEET, HISTORY_HEADERS
from ..models import Issue


def read_history_xlsx(path: Path) -> list[Issue]:
    """读取历史审核工作表为 Issue 列表；文件/表缺失时返回空列表。"""
    if not path.is_file():
        return []
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        sheet = book[HISTORY_AUDIT_SHEET] if HISTORY_AUDIT_SHEET in book.sheetnames else (
            book["历史审核结果"] if "历史审核结果" in book.sheetnames else None
        )
        if sheet is None:
            return []
        rows = list(sheet.iter_rows(values_only=True))
        if len(rows) < 2:
            return []
        headers = {str(value or "").strip(): index for index, value in enumerate(rows[0])}

        def get(row: tuple[object, ...], name: str) -> object:
            index = headers.get(name)
            return row[index] if index is not None and index < len(row) else ""

        result: list[Issue] = []
        for row in rows[1:]:
            source_file, sheet_value, target_cell = (
                str(get(row, key) or "") for key in ("工作簿名", "工作表名", "定位单元格")
            )
            rule_id = str(get(row, "规则编号") or "")
            if not any((source_file, sheet_value, target_cell, rule_id)):
                continue
            indicator = str(get(row, "校验指标") or "")
            identity = rule_id or "｜".join((source_file, sheet_value, target_cell, indicator or "校验指标未识别"))
            result.append(Issue(
                issue_id=identity, period=str(get(row, "数据期") or ""), batch_id="", audit_time="",
                triggered=True, status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                org_code="", org_name="", report_code="", sheet_name=sheet_value, rule_id=identity,
                severity=str(get(row, "错误类型") or "错误"), formula_cell=target_cell, target_cell=target_cell,
                target_value=get(row, "当前值"), formula_result="", message=str(get(row, "描述") or ""),
                source_file=source_file, audit_file="", check_field=indicator, comparison_value=get(row, "对比值"),
                difference_value=get(row, "差值"), detail=str(get(row, "描述") or ""),
                institution_feedback=str(get(row, "历史校验说明") or ""), auditor_opinion=str(get(row, "审核意见") or ""),
            ))
        return result
    finally:
        book.close()


def write_history_xlsx(path: Path, issues: Iterable[Issue]) -> None:
    """只重写程序管理的历史表；工作簿中的其他表（含人工补充列所在表）原样保留。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    book = load_workbook(path) if path.is_file() else Workbook()
    try:
        if HISTORY_AUDIT_SHEET in book.sheetnames:
            del book[HISTORY_AUDIT_SHEET]
        sheet = book.create_sheet(HISTORY_AUDIT_SHEET, 0)
        sheet.append(HISTORY_HEADERS)
        for item in issues:
            sheet.append((
                Path(item.source_file).name if item.source_file else "", item.sheet_name, item.target_cell,
                item.severity, item.check_field, item.detail or item.message, item.target_value,
                item.comparison_value, item.difference_value, item.issue_id,
                item.institution_feedback, item.auditor_opinion,
            ))
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in range(1, len(HISTORY_HEADERS) + 1):
            sheet.column_dimensions[chr(64 + column)].width = 22
        # 新建工作簿残留的空 Sheet 清理；用户已有工作簿不动其他表。
        if "Sheet" in book.sheetnames and len(book.sheetnames) > 1:
            del book["Sheet"]
        book.save(path)
    finally:
        book.close()
