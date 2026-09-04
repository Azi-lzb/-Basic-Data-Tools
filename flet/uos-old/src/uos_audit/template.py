from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass

from .models import AuditRule


REQUIRED_HEADERS = (
    "规则编号",
    "启用",
    "报表代码",
    "工作表",
    "公式单元格",
    "定位单元格",
    "级别",
    "问题说明",
)
ALLOWED_SEVERITIES = {"错误", "核实", "提示"}
CELL_RE = re.compile(r"^\$?[A-Z]{1,3}\$?[1-9][0-9]*$", re.IGNORECASE)
DATE_TOKEN_RE = re.compile(r"20\d{2}(?:[-._]?\d{2}){1,2}")
RULE_ID_RE = re.compile(
    r"(?:规则编号|规则ID)\s*[:：]\s*([^\s，,；;]+)", re.IGNORECASE
)
VALID_RULE_ID_RE = re.compile(r"^[\w.-]{2,64}$", re.UNICODE)
REGION_MARKER_RE = re.compile(r"(?:本地校验区域|联表校验区域)#?\d+")


class TemplateError(ValueError):
    pass


@dataclass(frozen=True)
class ParsedFormulaResult:
    severity: str
    indicator: str
    value: str
    message: str
    comparison_value: str = ""
    reference_value: str = ""
    difference_value: str = ""
    detail: str = ""


def normalize_template_name(value: str) -> str:
    """Remove the leading template mark and changing date suffixes."""
    text = value.strip().lstrip("!！")
    text = DATE_TOKEN_RE.sub("", text)
    return text.rstrip(" _.-")


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def parse_comment_rule_id(comment_text: str) -> str:
    match = RULE_ID_RE.search(comment_text or "")
    if not match:
        return ""
    rule_id = match.group(1).strip()
    if not VALID_RULE_ID_RE.fullmatch(rule_id):
        raise TemplateError(f"规则编号无效：{rule_id}")
    return rule_id


def clean_rule_comment(comment_text: str) -> str:
    text = re.sub(r"^[^\n]{1,30}:\s*", "", comment_text or "").strip()
    text = RULE_ID_RE.sub("", text)
    text = REGION_MARKER_RE.sub("", text)
    return "\n".join(line.strip(" ：:，,") for line in text.splitlines() if line.strip(" ：:，,"))


def parse_formula_result(
    value: object,
    *,
    default_severity: str = "错误",
    default_message: str = "",
) -> ParsedFormulaResult:
    text = _text(value).replace("｜", "|")
    parts = [part.strip() for part in text.split("|")]
    marker = parts[0] if parts else ""
    if any(word in marker for word in ("错误", "硬性", "错")):
        severity = "错误"
    elif "软性" in marker:
        severity = "核实"
    else:
        severity = default_severity

    comparison_value = ""
    reference_value = ""
    difference_value = ""
    detail = ""
    # 标准格式：错误类型|校验指标|描述|当前值|对比值|差值。
    # 三段即可表达一条问题；后三段仅用于需要展示精确取值的规则。旧模板
    # 的数值优先六段格式仍按旧位置解析，避免已经投入使用的公式失效。
    has_named_indicator = bool(re.search(r"[\u4e00-\u9fff]", parts[1] if len(parts) > 1 else ""))
    third_part_is_description = bool(
        re.search(r"[\u4e00-\u9fff]", parts[2] if len(parts) > 2 else "")
    )
    if len(parts) >= 6 and has_named_indicator and third_part_is_description:
        # 标准六段格式：错误类型|校验指标|描述|当前值|对比值|差值
        indicator = parts[1]
        detail = parts[2]
        parsed_value = parts[3]
        comparison_value = parts[4]
        difference_value = "|".join(parts[5:]).strip()
        message = detail or default_message or text
    elif len(parts) >= 6 and has_named_indicator:
        # 旧六段格式：错误类型|校验指标|当前值|对比值|差值|描述
        indicator = parts[1]
        parsed_value = parts[2]
        comparison_value = parts[3]
        difference_value = parts[4]
        detail = "|".join(parts[5:]).strip()
        message = detail or default_message or text
    elif len(parts) >= 5:
        # Earlier structured convention:
        # 错|对比值|参考值|差值|详细说明
        indicator = ""
        comparison_value = parts[1]
        reference_value = parts[2]
        difference_value = parts[3]
        detail = "|".join(parts[4:]).strip()
        parsed_value = difference_value
        message = detail or default_message or text
    elif len(parts) == 4:
        # Existing templates use: 错|校验字段|校验值|详细说明
        indicator = parts[1]
        parsed_value = parts[2]
        difference_value = parsed_value
        detail = parts[3]
        message = detail or default_message or text
    elif len(parts) == 3:
        indicator = parts[1]
        # 标准最简格式：错误类型|校验指标|描述
        parsed_value = ""
        detail = parts[2] or default_message or text
        message = detail
    elif len(parts) == 2:
        indicator = ""
        parsed_value = ""
        detail = parts[1] or default_message or text
        message = parts[1] or default_message or text
    else:
        indicator = ""
        parsed_value = ""
        detail = default_message or text
        message = default_message or text
    if indicator and indicator not in message:
        message = f"{indicator}：{message}"
    return ParsedFormulaResult(
        severity=severity,
        indicator=indicator,
        value=parsed_value,
        message=message,
        comparison_value=comparison_value,
        reference_value=reference_value,
        difference_value=difference_value,
        detail=detail,
    )


def parse_rule_rows(rows: Iterable[Mapping[str, object]]) -> list[AuditRule]:
    rules: list[AuditRule] = []
    seen: set[str] = set()
    errors: list[str] = []

    for row_number, row in enumerate(rows, start=2):
        rule_id = _text(row.get("规则编号"))
        if not rule_id:
            continue
        if rule_id in seen:
            errors.append(f"第{row_number}行规则编号重复：{rule_id}")
            continue
        seen.add(rule_id)

        enabled_text = _text(row.get("启用"))
        enabled = enabled_text not in {"否", "0", "False", "false", "停用"}
        sheet_name = _text(row.get("工作表"))
        formula_cell = _text(row.get("公式单元格")).replace("$", "").upper()
        target_cell = _text(row.get("定位单元格")).replace("$", "").upper()
        severity = _text(row.get("级别")) or "错误"

        if not sheet_name:
            errors.append(f"第{row_number}行未填写工作表：{rule_id}")
        if not CELL_RE.match(formula_cell):
            errors.append(f"第{row_number}行公式单元格无效：{formula_cell or '空'}")
        if not CELL_RE.match(target_cell):
            errors.append(f"第{row_number}行定位单元格无效：{target_cell or '空'}")
        if severity not in ALLOWED_SEVERITIES:
            errors.append(f"第{row_number}行级别无效：{severity}")

        rules.append(
            AuditRule(
                rule_id=rule_id,
                enabled=enabled,
                report_code=_text(row.get("报表代码")),
                sheet_name=sheet_name,
                formula_cell=formula_cell,
                target_cell=target_cell,
                severity=severity,
                message=_text(row.get("问题说明")),
                copy_range=formula_cell,
                result_mode="nonempty",
            )
        )

    if errors:
        raise TemplateError("；".join(errors))
    if not rules:
        raise TemplateError("审核规则表中没有可读取的规则")
    return rules
