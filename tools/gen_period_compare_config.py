"""从旧 VBA 统计工具工作簿生成 跨期比较配置.xlsx（一次性迁移工具）。

用法：
    python tools/gen_period_compare_config.py [VBA工作簿] [输出路径]

默认：
    源 reference/统计工具-报表工具(基础数据用)v1.0.4.9(1).xls
    目标 unified/core/跨期比较配置.xlsx

只迁移三张参照表（指标参照/机构参照/警戒区间）；特殊指标-自定义、
复杂校验-自定义留二期迁移，此处仅建空 sheet 占位。
"""

from __future__ import annotations

import sys
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE = ROOT / "reference" / "统计工具-报表工具(基础数据用)v1.0.4.9(1).xls"
DEFAULT_TARGET = ROOT / "unified" / "core" / "跨期比较配置.xlsx"


def _text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        return value.strip()
    return str(value)


def _norm_code(value):
    text = _text(value)
    return text[:-2] if text.endswith(".0") else text


def _bool_cn(value):
    if isinstance(value, (int, float)):
        return "是" if value else ""
    return _text(value)


def main() -> int:
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SOURCE
    target = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_TARGET
    book = xlrd.open_workbook(str(source))

    # ---- 指标参照：编码/名称/表单/频度/数据属性/核对要求/不转换/禁用/大集中核对/大集中名 ----
    sh = book.sheet_by_name("指标参照")
    indicators = []
    for r in range(1, sh.nrows):
        code = _norm_code(sh.cell_value(r, 0))
        if not code:
            continue
        indicators.append({
            "code": code,
            "name": _text(sh.cell_value(r, 1)),
            "form": _norm_code(sh.cell_value(r, 2)),
            "freq": _text(sh.cell_value(r, 3)),
            "type": _text(sh.cell_value(r, 4)),
            "hint": _text(sh.cell_value(r, 7)),
            "no_unit": _bool_cn(sh.cell_value(r, 8)),
            "disabled": _bool_cn(sh.cell_value(r, 9)),
            "central": _bool_cn(sh.cell_value(r, 10)),
            "central_name": _text(sh.cell_value(r, 11)),
        })

    # ---- 机构参照：名称/代码/类别/归属行/承接行/停用/报表项目 ----
    sh = book.sheet_by_name("机构参照")
    orgs = []
    for r in range(1, sh.nrows):
        name = _text(sh.cell_value(r, 0))
        if not name:
            continue
        orgs.append({
            "name": name,
            "code": _text(sh.cell_value(r, 4)),
            "org_class": _text(sh.cell_value(r, 8)),
            "region": _text(sh.cell_value(r, 9)),
            "bank_row": _text(sh.cell_value(r, 10)),
            "disabled": _bool_cn(sh.cell_value(r, 16)),
            "report_item": _text(sh.cell_value(r, 14)),
        })

    # ---- 警戒区间：序号/变幅(小数)/备注/整行填充 ----
    sh = book.sheet_by_name("警戒区间")
    alerts = []
    for r in range(1, sh.nrows):
        bound = sh.cell_value(r, 1)
        if bound in (None, ""):
            continue
        alerts.append((
            _norm_code(sh.cell_value(r, 0)),
            float(bound),
            _text(sh.cell_value(r, 2)),
            _bool_cn(sh.cell_value(r, 3)),
        ))

    # ---- 特殊指标-自定义：按 VBA 列布局迁移；人行/机构禁用标记任一非空即禁用 ----
    sh = book.sheet_by_name("特殊指标-自定义")
    specials = []
    for r in range(1, sh.nrows):
        code = _norm_code(sh.cell_value(r, 0))
        remark = _text(sh.cell_value(r, 2))
        if not code or not remark:
            continue
        specials.append((
            code,
            _text(sh.cell_value(r, 1)),
            remark,
            sh.cell_value(r, 3),
            sh.cell_value(r, 4),
            _text(sh.cell_value(r, 5)),
            _text(sh.cell_value(r, 6)),
            _norm_code(sh.cell_value(r, 7)),
            _text(sh.cell_value(r, 8)),
            _text(sh.cell_value(r, 12)),
            "是" if (sh.cell_value(r, 9) or sh.cell_value(r, 10)) else "",
        ))

    # ---- 复杂校验-自定义：表单/描述/规则/取反/Thd/禁用/备注 ----
    sh = book.sheet_by_name("复杂校验-自定义")
    complex_rules = []
    for r in range(1, sh.nrows):
        rule = _text(sh.cell_value(r, 2))
        desc = _text(sh.cell_value(r, 1))
        if not rule or not desc:
            continue
        complex_rules.append((
            _text(sh.cell_value(r, 0)),
            desc,
            rule,
            _bool_cn(sh.cell_value(r, 3)),
            sh.cell_value(r, 5),
            "是" if sh.cell_value(r, 4) else "",
            _text(sh.cell_value(r, 6)),
        ))

    # ---- 写出（列布局与 base_audit.period_compare 读取器一致） ----
    if target.is_file():
        existing = load_workbook(target)
    else:
        existing = None
    out = Workbook()
    sheet = out.active
    sheet.title = "指标参照"
    sheet.append(["指标代码", "指标名称", "数据属性", "是否不转换单位", "是否与大集中核对", "大集中报表查询指标名称", "禁用", "表单", "频度", "核对要求"])
    for item in indicators:
        sheet.append([
            item["code"], item["name"], item["type"], item["no_unit"],
            item["central"], item["central_name"], item["disabled"],
            item["form"], item["freq"], item["hint"],
        ])
    org_sheet = out.create_sheet("机构参照")
    org_sheet.append(["机构名称", "社会信用代码", "机构类别", "承接行", "归属行", "报表项目", "禁用"])
    for item in orgs:
        org_sheet.append([
            item["name"], item["code"], item["org_class"],
            item["bank_row"], item["region"], item["report_item"], item["disabled"],
        ])
    alert_sheet = out.create_sheet("警戒区间")
    alert_sheet.append(["序号", "变幅下限（小数，0.3=30%）", "备注", "是否整行填充"])
    for row in alerts:
        alert_sheet.append(list(row))
    special_sheet = out.create_sheet("特殊指标-自定义")
    special_sheet.append([
        "指标代码", "指标名", "备注", "人民币阀值(亿元)", "美元合计阀值(亿美元)",
        "数据属性", "说明/比较指标", "表单", "详细说明", "绝对值变幅（%）", "禁用",
    ])
    for row in specials:
        special_sheet.append(list(row))
    complex_sheet = out.create_sheet("复杂校验-自定义")
    complex_sheet.append(["校验表单", "校验描述", "校验规则", "取反标识", "Thd值(万元)", "禁用", "备注"])
    for row in complex_rules:
        complex_sheet.append(list(row))

    target.parent.mkdir(parents=True, exist_ok=True)
    out.save(target)
    print(f"已生成 {target}")
    print(f"  指标参照 {len(indicators)} 条（其中大集中核对 {sum(1 for i in indicators if i['central'] == '是')} 条）")
    print(f"  机构参照 {len(orgs)} 条；警戒区间 {len(alerts)} 档")
    print(f"  特殊指标 {len(specials)} 条（禁用 {sum(1 for s in specials if s[10] == '是')}）；复杂校验 {len(complex_rules)} 条（禁用 {sum(1 for c in complex_rules if c[5] == '是')}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
