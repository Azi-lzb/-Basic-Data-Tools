"""把旧版（5 表）比较配置工作簿升级为新版（4 表）格式。

用法：
    python tools/gen_period_compare_config.py [旧配置簿] [输出路径]

默认：
    源 unified/core/报表采集系统_比较配置.xlsx（缺省回退旧名 跨期比较配置.xlsx）
    目标 unified/core/报表采集系统_比较配置.xlsx（就地覆盖；文件被 Excel 占用时另指定路径）

转换规则：
- 指标参照：删除程序不读的“表单/频度”列，“核对要求”并入“备注”；
- 机构参照：“归属行”更名“地区”（与代码语义、输出列名一致）；
- 警戒区间：删除“序号”列，“是否整行填充”更名“整行填充”；
- 校验规则：合并“特殊指标-自定义”与“复杂校验-自定义”——
  · 同文案的“累计不降”规则按类型各合并为一条，规则内容为指标代码清单，
    旧表中被禁用的代码记入备注，不丢信息；
  · 复杂校验表达式改写为短写（[,,,20203046,] → [20203046]）并按校验性质
    标注级别（勾稽/不可能关系=错误，监管线/占比阈值=核实）；
  · 校验表单不是 20202/20203 的历史测试行（A0000/表1-1）直接丢弃；
  · 取反标识、Thd值列因从未启用不迁移。

生成后立即用新版加载器回读自检，打印规则条数与配置提醒。
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "unified" / "core" / "src"))

from base_audit.period_compare import (  # noqa: E402
    CONFIG_ALERT_SHEET,
    CONFIG_COMPLEX_SHEET,
    CONFIG_INDICATOR_SHEET,
    CONFIG_ORG_SHEET,
    CONFIG_RULE_SHEET,
    CONFIG_SPECIAL_SHEET,
    DEFAULT_RULE_LEVEL,
    RULE_TYPE_ACC_HISTORY,
    RULE_TYPE_ACC_YEAR,
    RULE_TYPE_EXPRESSION,
    load_period_config,
)

DEFAULT_SOURCE = ROOT / "unified" / "core" / "报表采集系统_比较配置.xlsx"
LEGACY_SOURCE = ROOT / "unified" / "core" / "跨期比较配置.xlsx"
DEFAULT_TARGET = DEFAULT_SOURCE

def _level_for(desc: str) -> str:
    """按校验性质定级：勾稽/不可能关系=错误，监管线与占比阈值=核实，其余=提示。"""
    if any(key in desc for key in ("不平衡", "不等于", "层级关系", "大于逾期贷款", "大于贷款总额")):
        return "错误"
    if any(key in desc for key in ("低于", "超过", "过高", "过低", "大于总资产", "杠杆")):
        return "核实"
    return DEFAULT_RULE_LEVEL

_KIND_TO_TYPE = {
    "当年累计指标比上期不应减少": RULE_TYPE_ACC_YEAR,
    "历史累计指标比上期不应减少": RULE_TYPE_ACC_HISTORY,
}


def _text(value) -> str:
    return str(value or "").strip()


def _norm_code(value) -> str:
    text = _text(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _cell_bool(value) -> bool:
    return _text(value) in {"是", "true", "1", "y", "yes"}


def _shorten_expression(expression: str) -> str:
    """[,,,20203046,] → [20203046]；{,,,代码,} → {代码}；规范运算符空格。"""

    def replace(match: re.Match) -> str:
        opener, inner, closer = match.group(1), match.group(2), match.group(3)
        if "," in inner:
            parts = [p.strip() for p in inner.split(",")]
            inner = parts[3] if len(parts) > 3 else ""
        return f"{opener}{inner}{closer}"

    result = re.sub(r"(\[|\{)([^\]}]*)(\]|\})", replace, expression)
    # and/or/not 函数式写法统一为大写 AND/OR/NOT，便于阅读。
    result = re.sub(r"\b(and|or|not)\s*\(", lambda m: m.group(1).upper() + "(", result, flags=re.IGNORECASE)
    # 统一运算符两侧空格（<>、>=、<= 优先于单字符，避免拆散）。
    return re.sub(r"\s*(<>|>=|<=|>|<|\*|\+|-|/)\s*", r" \1 ", result)


def main() -> int:
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else (
        DEFAULT_SOURCE if DEFAULT_SOURCE.is_file() else LEGACY_SOURCE
    )
    target = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_TARGET
    book = load_workbook(source, read_only=True, data_only=True)
    try:
        if CONFIG_RULE_SHEET in book.sheetnames:
            print(f"源文件已是新版 4 表格式，无需迁移：{source}")
            return 0
        problems: list[str] = []

        # ---- 指标参照（旧列序：代码/名称/属性/不转换/核对/大集中名/禁用/表单/频度/核对要求） ----
        indicators: list[list] = []
        if CONFIG_INDICATOR_SHEET in book.sheetnames:
            for row in book[CONFIG_INDICATOR_SHEET].iter_rows(min_row=2, values_only=True):
                code = _norm_code(row[0] if row else None)
                if not code:
                    continue
                indicators.append([
                    code,
                    _text(row[1] if len(row) > 1 else None),
                    _text(row[2] if len(row) > 2 else None),
                    "是" if (len(row) > 3 and _cell_bool(row[3])) else "",
                    "是" if (len(row) > 4 and _cell_bool(row[4])) else "",
                    _text(row[5] if len(row) > 5 else None),
                    "是" if (len(row) > 6 and _cell_bool(row[6])) else "",
                    _text(row[9] if len(row) > 9 else None),   # 核对要求 → 备注
                ])

        # ---- 机构参照（旧列序：名称/代码/类别/承接行/归属行/报表项目/禁用） ----
        orgs: list[list] = []
        if CONFIG_ORG_SHEET in book.sheetnames:
            for row in book[CONFIG_ORG_SHEET].iter_rows(min_row=2, values_only=True):
                name = _text(row[0] if row else None)
                if not name:
                    continue
                orgs.append([
                    name,
                    _text(row[1] if len(row) > 1 else None),
                    _text(row[2] if len(row) > 2 else None),
                    _text(row[3] if len(row) > 3 else None),
                    _text(row[4] if len(row) > 4 else None),
                    _text(row[5] if len(row) > 5 else None),
                    "是" if (len(row) > 6 and _cell_bool(row[6])) else "",
                ])

        # ---- 警戒区间（旧列序：序号/变幅下限/备注/是否整行填充） ----
        alerts: list[list] = []
        if CONFIG_ALERT_SHEET in book.sheetnames:
            for row in book[CONFIG_ALERT_SHEET].iter_rows(min_row=2, values_only=True):
                if not row or row[1] in (None, ""):
                    continue
                try:
                    lower = float(row[1])
                except (TypeError, ValueError):
                    continue
                alerts.append([
                    lower,
                    _text(row[2] if len(row) > 2 else None),
                    "是" if (len(row) > 3 and _cell_bool(row[3])) else "",
                ])

        # ---- 特殊指标-自定义：按规则类型合并为“累计不降”代码清单 ----
        acc_enabled: dict[str, list[str]] = {}
        acc_disabled: dict[str, list[str]] = {}
        acc_desc: dict[str, str] = {}
        unsupported_kinds: set[str] = set()
        if CONFIG_SPECIAL_SHEET in book.sheetnames:
            for row in book[CONFIG_SPECIAL_SHEET].iter_rows(min_row=2, values_only=True):
                code = _norm_code(row[0] if row else None)
                remark = _text(row[2] if len(row) > 2 else None)
                if not code or not remark:
                    continue
                text = remark.replace("。", "").replace("：", ":").strip()
                kind = next((k for k in _KIND_TO_TYPE if text.startswith(k)), None)
                if kind is None:
                    unsupported_kinds.add(_text(remark))
                    continue
                bucket = acc_disabled if (len(row) > 10 and _cell_bool(row[10])) else acc_enabled
                bucket.setdefault(kind, []).append(code)
                acc_desc.setdefault(kind, remark)
        for remark in sorted(unsupported_kinds):
            problems.append(f"特殊指标存在未支持类型，未迁移：{remark}")

        # ---- 复杂校验-自定义 → 表达式规则 ----
        expression_rules: list[list] = []
        if CONFIG_COMPLEX_SHEET in book.sheetnames:
            for row in book[CONFIG_COMPLEX_SHEET].iter_rows(min_row=2, values_only=True):
                rule = _text(row[2] if len(row) > 2 else None)
                desc = _text(row[1] if len(row) > 1 else None)
                if not rule or not desc:
                    continue
                form = _text(row[0] if len(row) > 0 else None)
                if form not in {"20202", "20203"}:
                    problems.append(f"复杂校验“{desc}”（表单 {form or '空'}）为历史测试行，已丢弃")
                    continue
                if len(row) > 3 and _cell_bool(row[3]):
                    problems.append(f"复杂校验“{desc}”带取反标识，历史上从未执行，未迁移")
                    continue
                expression_rules.append([
                    desc,
                    _shorten_expression(rule),
                    _level_for(desc),
                    "是" if (len(row) > 5 and _cell_bool(row[5])) else "",
                ])
    finally:
        book.close()

    # ---- 写出新版 4 表 ----
    out = Workbook()
    sheet = out.active
    sheet.title = CONFIG_INDICATOR_SHEET
    sheet.append(["指标代码", "指标名称", "数据属性", "不转换单位", "大集中核对", "大集中指标名称", "禁用", "备注"])
    for row in indicators:
        sheet.append(row)
    org_sheet = out.create_sheet(CONFIG_ORG_SHEET)
    org_sheet.append(["机构名称", "社会信用代码", "机构类别", "承接行", "地区", "报表项目", "禁用"])
    for row in orgs:
        org_sheet.append(row)
    alert_sheet = out.create_sheet(CONFIG_ALERT_SHEET)
    alert_sheet.append(["变幅下限（小数，0.3=30%）", "备注", "整行填充"])
    for row in alerts:
        alert_sheet.append(row)
    rule_sheet = out.create_sheet(CONFIG_RULE_SHEET)
    rule_sheet.append(["规则编号", "类型", "描述", "规则内容", "级别", "禁用", "备注"])

    rule_rows: list[list] = []
    for kind, rule_type in _KIND_TO_TYPE.items():
        if kind not in acc_enabled and kind not in acc_disabled:
            continue
        note = "旧表停用代码：" + "、".join(acc_disabled[kind]) if kind in acc_disabled else ""
        rule_rows.append([
            acc_desc.get(kind, ""),
            rule_type,
            ",".join(acc_enabled.get(kind, [])),
            "提示",
            "",
            note,
        ])
    for desc, rule, level, disabled in expression_rules:
        rule_rows.append([desc, RULE_TYPE_EXPRESSION, rule, level, disabled, ""])
    for index, row in enumerate(rule_rows, start=1):
        rule_id = f"R{index:03d}"
        desc, rule_type, content, level, disabled, note = row
        rule_sheet.append([rule_id, rule_type, desc, content, level, disabled, note])

    target.parent.mkdir(parents=True, exist_ok=True)
    out.save(target)
    print(f"已生成 {target}")
    print(f"  指标参照 {len(indicators)} 条；机构参照 {len(orgs)} 条；警戒区间 {len(alerts)} 档")
    print(f"  校验规则 {len(rule_rows)} 条（累计不降 {sum(1 for r in rule_rows if r[1] != RULE_TYPE_EXPRESSION)} 条，"
          f"表达式 {sum(1 for r in rule_rows if r[1] == RULE_TYPE_EXPRESSION)} 条）")

    # ---- 回读自检 ----
    config = load_period_config(target)
    active = [r for r in config.rules if not r.disabled]
    if active:
        print("自检：启用规则按类型统计：")
        for rule_type in (RULE_TYPE_ACC_YEAR, RULE_TYPE_ACC_HISTORY, RULE_TYPE_EXPRESSION):
            count = sum(1 for r in active if r.type == rule_type)
            if count:
                print(f"  {rule_type}：{count} 条")
    else:
        print("自检：无启用规则")
    if config.warnings:
        print(f"自检发现 {len(config.warnings)} 条配置提醒：")
        for warning in config.warnings:
            print(f"  - {warning}")
    for problem in problems:
        print(f"迁移备注：{problem}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
