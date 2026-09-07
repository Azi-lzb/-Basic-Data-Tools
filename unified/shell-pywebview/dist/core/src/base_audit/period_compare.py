"""基础数据季报表跨期比较（第一期：两期比较 + 大集中核对）。

对齐旧 VBA `基础数据季报表跨期比较.bas` 的核心口径：
- 输入为“当期目录 + 上期目录”，支持三种系统导出方式（自动识别）：
  1. 逐机构逐表：`机构代码#日期#01#表单#机构名.xls`，sheet=表单；
  2. 按报表划分：`banks#日期#01#表单.xls`，sheet 名=机构代码；
  3. 按机构划分：`reports#机构代码#日期#01#机构名.xls`，sheet=表单。
  数据工作表统一为：第 1 列指标编号、第 2 列指标名称、第 3 列值，第 4 行起。
- 指标/机构/警戒区间配置放在独立工作簿 `跨期比较配置.xlsx`（用户维护、程序只读）。
- 大集中核对：读用户选择的大集中工作簿（集中系统数据 + 参照表），按
  机构代码→报表项目、指标代码→大集中指标名称映射，差异绝对值 > 0.01 万元
  记“差异超过100元”，单边有值记“谨慎核实”。

纯 openpyxl/xlrd 实现，不启动 Excel/WPS/LibreOffice，Windows 与 UOS 通用。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from openpyxl import Workbook, load_workbook

# —— 口径常量（与 VBA 一致） ——
SOURCE_UNIT = "元"
DEST_UNIT = "万元"
UNIT_FACTOR = 1.0 / 10000.0          # 元 → 万元
CENTRAL_DIFF_TOLERANCE = 0.01        # 万元；差异绝对值超过它记“差异超过100元”
CENTRAL_UNIT_FACTOR = 10000.0        # 集中系统数据单位为亿元 → 万元
INDICATOR_NAME_CODE = "20201001"     # 按报表划分导出里取机构名称的指标编号
INDICATOR_CODE_CODE = "20201002"     # 同上，机构代码

PERIOD_SHEET_HEADERS = [
    "地区", "数据属性", "机构名称", "指标编码", "指标名称", "当期数", "上期数",
    "变动绝对值", "环比变动", "备注", "是否说明", "计算过程", "机构类别", "承接行",
    "社会信用代码", "数据日期", "币种", "频度",
]
CENTRAL_SHEET_HEADERS = [
    "地区", "数据属性", "机构名称", "指标编码", "指标名称", "基础数据值",
    "大集中值", "差异绝对值", "差异幅度", "是否说明",
]

CONFIG_WORKBOOK_NAME = "跨期比较配置.xlsx"
CONFIG_INDICATOR_SHEET = "指标参照"
CONFIG_ORG_SHEET = "机构参照"
CONFIG_ALERT_SHEET = "警戒区间"
CONFIG_SPECIAL_SHEET = "特殊指标-自定义"
CONFIG_COMPLEX_SHEET = "复杂校验-自定义"


class PeriodCompareError(RuntimeError):
    """跨期比较的明确业务失败（目录为空、文件名不合法、配置缺失等）。"""


@dataclass
class IndicatorDef:
    code: str
    name: str = ""
    data_type: str = ""            # 数据属性：余额/累发/个数/百分数…
    no_unit_convert: bool = False  # 是否不转换单位
    check_central: bool = False    # 是否与大集中核对
    central_name: str = ""         # 大集中报表查询指标名称


@dataclass
class OrgDef:
    name: str
    code: str = ""
    org_class: str = ""            # 机构类别
    bank_row: str = ""             # 承接行
    region: str = ""               # 地区/归属行
    report_item: str = ""          # 报表项目（大集中行标识）


@dataclass
class AlertRange:
    lower_bound: float    # 变幅下限（环比小数，0.3 = 30%）
    remark: str
    fill_row: bool = False


@dataclass
class SpecialRule:
    """特殊指标规则（对齐 VBA“特殊指标-自定义”）。

    第一期实现备注文本为“当年累计/历史累计指标比上期不应减少”的规则；
    其他 VBA 规则类型读取保留但跳过执行。
    """

    code: str
    name: str
    remark: str           # 备注文本，同时是规则类型标识与提示内容
    rmb_threshold: float | None = None   # 人民币阀值（亿元）
    usd_threshold: float | None = None
    data_type: str = ""
    explain: str = ""     # 说明/比较指标代码
    form: str = ""
    detailed: str = ""
    change_pct: float | None = None      # 绝对值变幅（%）
    disabled: bool = False

    @property
    def rule_kind(self) -> str:
        text = self.remark.replace("。", "").replace("：", ":").strip()
        if text.startswith("当年累计指标比上期不应减少"):
            return "当年累计不应减少"
        if text.startswith("历史累计指标比上期不应减少"):
            return "历史累计不应减少"
        return "未支持"


@dataclass
class ComplexRule:
    """复杂校验表达式规则（对齐 VBA“复杂校验-自定义”）。

    表达式语法：`[机构,地区,类别,指标代码,数据属性,币种,频度,批次]` 为当期值，
    `{...}` 同结构为上期值；第 4 段（下标 3）是指标代码。``Thd`` 占位符取
    规则自带阈值。支持 Excel 风格 AND/OR/NOT 与 `<>` 运算符。
    """

    form: str
    desc: str
    rule: str
    invert: bool = False
    threshold: float | None = None
    disabled: bool = False
    note: str = ""


@dataclass
class PeriodConfig:
    indicators: dict[str, IndicatorDef] = field(default_factory=dict)
    orgs: dict[str, OrgDef] = field(default_factory=dict)      # key=机构名称
    orgs_by_code: dict[str, OrgDef] = field(default_factory=dict)
    alerts: list[AlertRange] = field(default_factory=list)
    specials: list[SpecialRule] = field(default_factory=list)
    complex_rules: list[ComplexRule] = field(default_factory=list)

    def alert_remark(self, change_pct: float | None) -> str:
        """按 VBA AddbfMark 口径匹配警戒区间。

        配置中的“变幅下限”是小数（0.3 表示 30%）；环比变动以百分数传入
        （41.56 表示 +41.56%）。从大到小找第一个 ``环比% >= 下限×100`` 的
        档位取备注；恰好 -99% 记“缩小100倍及以上”，低于最小档记
        “近缩小100倍及以上”。
        """
        if change_pct is None or round(change_pct, 6) == 0:
            return ""
        matched: str | None = None
        for item in reversed(self.alerts):
            if round(change_pct - item.lower_bound * 100.0, 10) >= 0:
                matched = item.remark
                if round(change_pct - (-99.0), 10) == 0:
                    matched = "缩小100倍及以上"
                break
        if matched is None:
            return "近缩小100倍及以上"
        return matched


@dataclass
class IndicatorValue:
    org_name: str
    org_code: str
    code: str
    name: str
    value: Any            # float 或 str（文字型指标）
    date: str
    form: str             # 表单号 20201/20202/20203


# ---------------------------------------------------------------------------
# 配置读取
# ---------------------------------------------------------------------------

def _cell_bool(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    return str(value).strip() in {"是", "true", "1", "y", "yes"}


def _optional_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def load_period_config(config_path: Path) -> PeriodConfig:
    if not config_path.is_file():
        raise PeriodCompareError(
            f"未找到跨期比较配置：{config_path}。请在程序目录放置{CONFIG_WORKBOOK_NAME}"
            "（指标参照/机构参照/警戒区间），参照说明维护后重试。"
        )
    config = PeriodConfig()
    book = load_workbook(config_path, read_only=True, data_only=True)
    try:
        if CONFIG_INDICATOR_SHEET in book.sheetnames:
            for row in book[CONFIG_INDICATOR_SHEET].iter_rows(min_row=2, values_only=True):
                code = _norm_code(row[0]) if row and row[0] is not None else ""
                if not code or (len(row) > 6 and _cell_bool(row[6])):
                    continue
                config.indicators[code] = IndicatorDef(
                    code=code,
                    name=str(row[1] or "").strip() if len(row) > 1 else "",
                    data_type=str(row[2] or "").strip() if len(row) > 2 else "",
                    no_unit_convert=_cell_bool(row[3]) if len(row) > 3 else False,
                    check_central=_cell_bool(row[4]) if len(row) > 4 else False,
                    central_name=str(row[5] or "").strip() if len(row) > 5 else "",
                )
        if CONFIG_ORG_SHEET in book.sheetnames:
            for row in book[CONFIG_ORG_SHEET].iter_rows(min_row=2, values_only=True):
                name = str(row[0] or "").strip() if row and row[0] else ""
                if not name or (len(row) > 6 and _cell_bool(row[6])):
                    continue
                org = OrgDef(
                    name=name,
                    code=str(row[1] or "").strip() if len(row) > 1 else "",
                    org_class=str(row[2] or "").strip() if len(row) > 2 else "",
                    bank_row=str(row[3] or "").strip() if len(row) > 3 else "",
                    region=str(row[4] or "").strip() if len(row) > 4 else "",
                    report_item=str(row[5] or "").strip() if len(row) > 5 else "",
                )
                config.orgs[name] = org
                if org.code:
                    config.orgs_by_code[org.code] = org
        if CONFIG_ALERT_SHEET in book.sheetnames:
            for row in book[CONFIG_ALERT_SHEET].iter_rows(min_row=2, values_only=True):
                if not row or row[1] is None or row[1] == "":
                    continue
                try:
                    lower = float(row[1])
                except (TypeError, ValueError):
                    continue
                config.alerts.append(
                    AlertRange(
                        lower_bound=lower,
                        remark=str(row[2]).strip() if row[2] is not None else "",
                        fill_row=_cell_bool(row[3]) if len(row) > 3 else False,
                    )
                )
        if CONFIG_SPECIAL_SHEET in book.sheetnames:
            for row in book[CONFIG_SPECIAL_SHEET].iter_rows(min_row=2, values_only=True):
                code = _norm_code(row[0]) if row and row[0] is not None else ""
                remark = str(row[2] or "").strip() if len(row) > 2 else ""
                if not code or not remark:
                    continue
                config.specials.append(SpecialRule(
                    code=code,
                    name=str(row[1] or "").strip(),
                    remark=remark,
                    rmb_threshold=_optional_float(row[3]) if len(row) > 3 else None,
                    usd_threshold=_optional_float(row[4]) if len(row) > 4 else None,
                    data_type=str(row[5] or "").strip() if len(row) > 5 else "",
                    explain=_norm_code(row[6]) if len(row) > 6 else "",
                    form=_norm_code(row[7]) if len(row) > 7 else "",
                    detailed=str(row[8] or "").strip() if len(row) > 8 else "",
                    change_pct=_optional_float(row[9]) if len(row) > 9 else None,
                    disabled=_cell_bool(row[10]) if len(row) > 10 else False,
                ))
        if CONFIG_COMPLEX_SHEET in book.sheetnames:
            for row in book[CONFIG_COMPLEX_SHEET].iter_rows(min_row=2, values_only=True):
                rule = str(row[2] or "").strip() if len(row) > 2 else ""
                desc = str(row[1] or "").strip() if len(row) > 1 else ""
                if not rule or not desc:
                    continue
                config.complex_rules.append(ComplexRule(
                    form=str(row[0] or "").strip(),
                    desc=desc,
                    rule=rule,
                    invert=_cell_bool(row[3]) if len(row) > 3 else False,
                    threshold=_optional_float(row[4]) if len(row) > 4 else None,
                    disabled=_cell_bool(row[5]) if len(row) > 5 else False,
                    note=str(row[6] or "").strip() if len(row) > 6 else "",
                ))
    finally:
        book.close()
    config.alerts.sort(key=lambda item: item.lower_bound)
    return config


def ensure_default_config(config_path: Path) -> bool:
    """配置簿缺失时生成一份带默认警戒区间的模板；返回是否新建。"""
    if config_path.is_file():
        return False
    book = Workbook()
    sheet = book.active
    sheet.title = CONFIG_INDICATOR_SHEET
    sheet.append(["指标代码", "指标名称", "数据属性", "是否不转换单位", "是否与大集中核对", "大集中报表查询指标名称", "禁用"])
    org_sheet = book.create_sheet(CONFIG_ORG_SHEET)
    org_sheet.append(["机构名称", "社会信用代码", "机构类别", "承接行", "归属行", "报表项目", "禁用"])
    alert_sheet = book.create_sheet(CONFIG_ALERT_SHEET)
    alert_sheet.append(["序号", "变幅下限（小数，0.3=30%）", "备注", "是否整行填充"])
    default_alerts = [
        (1, -0.96, "降幅(-96%,-90%],缩小10倍-100倍", "是"),
        (2, -0.9, "降幅(-90%,-80%],缩小5倍-10倍", ""),
        (3, -0.8, "降幅(-80%,-50%],缩小1倍-5倍", ""),
        (4, -0.5, "降幅(-50%,-30%],缩小1倍以内", ""),
        (5, -0.3, "", ""),
        (6, 0.0, "", ""),
        (7, 0.3, "增幅[30%,50%)", ""),
        (8, 0.5, "增幅[50%,1倍)", ""),
        (9, 1.0, "增幅[1倍,5倍)", ""),
        (10, 5.0, "增幅[5倍,10倍)", ""),
        (11, 10.0, "增幅[10倍,96倍)", "是"),
        (12, 96.0, "近增幅100倍以上，请核实", "是"),
    ]
    for row in default_alerts:
        alert_sheet.append(row)
    book.create_sheet("特殊指标-自定义")   # 二期迁移预留
    book.create_sheet("复杂校验-自定义")   # 二期迁移预留
    config_path.parent.mkdir(parents=True, exist_ok=True)
    book.save(config_path)
    return True


# ---------------------------------------------------------------------------
# 两期数据读取（三种导出方式归一化）
# ---------------------------------------------------------------------------

def _norm_code(value: Any) -> str:
    """指标编号统一为整数字符串（xlrd 会把 20201001 读成 float）。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def _read_sheet_rows(sheet_rows: Iterable[Iterable[Any]]) -> list[tuple[str, str, Any]]:
    """按第 4 行起读 (指标编号, 指标名称, 值)；前 3 行为标题/空行/表头。"""
    rows: list[tuple[str, str, Any]] = []
    for index, row in enumerate(sheet_rows):
        if index < 3:
            continue
        values = list(row) + [None, None, None]
        code = _norm_code(values[0])
        if not code:
            continue
        rows.append((code, str(values[1] or "").strip(), values[2]))
    return rows


def _open_read_only(path: Path):
    """按后缀分派 .xls（xlrd）/ .xlsx（openpyxl），返回 (sheet名→行迭代器) 工厂。"""
    suffix = path.suffix.casefold()
    if suffix == ".xls":
        import xlrd

        book = xlrd.open_workbook(str(path))

        def make_rows(sheet_name: str) -> list[list[Any]]:
            sh = book.sheet_by_name(sheet_name)
            return [ [sh.cell_value(r, c) if c < sh.ncols else None for c in range(3)]
                     for r in range(sh.nrows) ]

        return book.sheet_names(), make_rows, lambda: None
    book = load_workbook(path, read_only=True, data_only=True)

    def make_rows_xlsx(sheet_name: str) -> list[list[Any]]:
        return [list(row) + [None, None, None] for row in book[sheet_name].iter_rows(values_only=True)]

    return list(book.sheetnames), make_rows_xlsx, book.close


def parse_source_file(path: Path) -> list[IndicatorValue]:
    """把一个导出文件归一化为指标记录列表（自动识别三种导出方式）。"""
    stem = path.stem
    segments = stem.split("#")
    records: list[IndicatorValue] = []
    sheet_names, make_rows, close = _open_read_only(path)
    try:
        if len(segments) >= 4 and segments[0] in {"banks", "BANKS"}:
            # 按报表划分：banks#日期#01#表单.xls，sheet 名=机构代码
            date, form = segments[1], segments[3]
            for sheet_name in sheet_names:
                rows = _read_sheet_rows(make_rows(sheet_name))
                lookup = {code: value for code, _name, value in rows}
                org_code = _norm_code(sheet_name)
                org_name = str(lookup.get(INDICATOR_NAME_CODE, "") or "")
                for code, name, value in rows:
                    records.append(IndicatorValue(org_name, org_code, code, name, value, date, form))
        else:
            # 逐机构逐表 / 按机构划分：文件名给出机构（或表单），sheet 名含 2020 为表单
            if len(segments) >= 5:
                org_code, date, form_from_name, org_name = segments[0], segments[1], segments[3], segments[4]
            else:
                raise PeriodCompareError(f"无法识别的导出文件名：{path.name}")
            for sheet_name in sheet_names:
                form = sheet_name if re.search(r"202\d{2}", str(sheet_name)) else form_from_name
                rows = _read_sheet_rows(make_rows(sheet_name))
                name_lookup = {code: value for code, _n, value in rows}
                resolved_name = str(name_lookup.get(INDICATOR_NAME_CODE, "") or org_name)
                resolved_code = str(name_lookup.get(INDICATOR_CODE_CODE, "") or org_code)
                for code, name, value in rows:
                    records.append(IndicatorValue(resolved_name, resolved_code, code, name, value, date, form))
    finally:
        close()
    return records


def list_period_pairs(
    current_dir: Path | None,
    previous_dir: Path | None,
) -> list[dict[str, object]]:
    """按“机构 + 表单”配对两期文件，供界面展示谁和谁比。

    三种导出方式的配对键：
    - 逐机构逐表（机构代码#日期#01#表单#机构名）：机构名 + 表单号；
    - 按机构划分（reports#代码#日期#01#机构名）：机构名（文件本身含全部表单）；
    - 按报表划分（banks#日期#01#表单）：表单号（文件本身含全部机构）。
    只在单期出现的文件同样列出，另一侧标记缺失，便于发现漏报。
    """

    def collect(directory: Path | None) -> dict[tuple[str, str], dict[str, str]]:
        result: dict[tuple[str, str], dict[str, str]] = {}
        if directory is None or not directory.is_dir():
            return result
        for path in sorted(directory.iterdir()):
            if not path.is_file() or path.suffix.casefold() not in {".xls", ".xlsx"} or path.name.startswith("~$"):
                continue
            segments = path.stem.split("#")
            if len(segments) >= 4 and segments[0].casefold() == "banks":
                key = ("（按报表划分）", segments[3])
            elif len(segments) >= 5 and segments[0].casefold() == "reports":
                key = (segments[4], "（全表单）")
            elif len(segments) >= 5:
                key = (segments[4], segments[3])
            else:
                key = (path.stem, "")
            result[key] = {"name": path.name, "path": str(path)}
        return result

    cur = collect(current_dir)
    pre = collect(previous_dir)
    rows: list[dict[str, object]] = []
    for key in sorted(set(cur) | set(pre), key=lambda k: (k[0], k[1])):
        cur_item = cur.get(key)
        pre_item = pre.get(key)
        rows.append({
            "org": key[0],
            "form": key[1],
            "curName": cur_item["name"] if cur_item else "",
            "curPath": cur_item["path"] if cur_item else "",
            "preName": pre_item["name"] if pre_item else "",
            "prePath": pre_item["path"] if pre_item else "",
            "matched": cur_item is not None and pre_item is not None,
            "side": "both" if cur_item and pre_item else "cur" if cur_item else "pre",
        })
    return rows


def load_period_directory(
    directory: Path,
    *,
    label: str,
    on_step: Callable[[str], None] | None = None,
) -> dict[tuple[str, str], IndicatorValue]:
    if not directory.is_dir():
        raise PeriodCompareError(f"{label}目录不存在：{directory}")
    files = sorted(
        p for p in directory.iterdir()
        if p.is_file() and p.suffix.casefold() in {".xls", ".xlsx"} and not p.name.startswith("~$")
    )
    if not files:
        raise PeriodCompareError(f"{label}目录中没有 .xls/.xlsx 报送文件：{directory}")
    records: dict[tuple[str, str], IndicatorValue] = {}
    for path in files:
        if on_step:
            on_step(f"读取{label}文件：{path.name}")
        for record in parse_source_file(path):
            records[(record.org_name, record.code)] = record
    return records


# ---------------------------------------------------------------------------
# 两期比较
# ---------------------------------------------------------------------------

def _numeric(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        try:
            return float(text)
        except ValueError:
            return None
    return None


def _apply_unit(value: Any, indicator: IndicatorDef | None) -> Any:
    number = _numeric(value)
    if number is None:
        return str(value).strip() if value is not None else ""
    if indicator is not None and indicator.no_unit_convert:
        return number
    return number * UNIT_FACTOR


def _vba_has_value(record: IndicatorValue, indicator: IndicatorDef | None) -> bool:
    """VBA 读数口径：文字型指标有非空值才读；数值型指标非 0 才读。

    对应 VBA ``If (data_type="文字" And value<>"") Or (data_type<>"文字" And
    ycV(value)<>0)``；ycV 把非数值文本当 0。零值/空行不进入比较与输出。
    """
    if record is None:
        return False
    is_text = indicator is not None and indicator.data_type == "文字"
    if is_text:
        return str(record.value).strip() != ""
    return _numeric(record.value) not in (None, 0.0)


def compare_periods(
    current: dict[tuple[str, str], IndicatorValue],
    previous: dict[tuple[str, str], IndicatorValue],
    config: PeriodConfig,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    keys = list(current.keys()) + [k for k in previous.keys() if k not in current]
    for org_name, code in keys:
        cur = current.get((org_name, code))
        pre = previous.get((org_name, code))
        if cur is None and pre is None:
            continue
        indicator = config.indicators.get(code)
        # VBA 口径：两期读数时零值/空行不进入结果（文字型有值即读）。
        if not _vba_has_value(cur, indicator) and not _vba_has_value(pre, indicator):
            continue
        sample = cur or pre
        org = config.orgs.get(org_name)
        is_text = indicator is not None and indicator.data_type == "文字"
        cur_value = _apply_unit(cur.value, indicator) if cur else None
        pre_value = _apply_unit(pre.value, indicator) if pre else None
        if is_text:
            # VBA 按数据属性判断文字型：即使值恰好是数字也只比较字符串。
            cur_value = str(cur.value).strip() if cur is not None else None
            pre_value = str(pre.value).strip() if pre is not None else None
        if isinstance(cur_value, str) or isinstance(pre_value, str):
            difference = "文字变动" if (cur_value or "") != (pre_value or "") else ""
            change_pct: float | None = None
            remark = "文字变动" if difference else ""
        else:
            cur_num = cur_value if cur_value is not None else 0.0
            pre_num = pre_value if pre_value is not None else 0.0
            if cur is not None and pre is None and cur_num != 0:
                difference, change_pct, remark = "", None, "本期有，上期无"
            elif pre is not None and cur is None and pre_num != 0:
                difference, change_pct, remark = "", None, "本期无，上期有"
            else:
                difference = cur_num - pre_num
                if pre_num != 0:
                    # 与 VBA 一致：环比 = 增减额 / 上期 * 100（上期取原值）。
                    change_pct = difference / pre_num * 100.0
                    remark = config.alert_remark(change_pct)
                else:
                    change_pct = None
                    remark = ""
        rows.append({
            "地区": org.region if org else "",
            "数据属性": indicator.data_type if indicator else "",
            "机构名称": org_name,
            "指标编码": code,
            "指标名称": (cur.name if cur else pre.name if pre else ""),
            "当期数": cur_value,
            "上期数": pre_value,
            "变动绝对值": difference,
            "环比变动": change_pct,
            "备注": remark,
            "机构类别": org.org_class if org else "",
            "承接行": org.bank_row if org else "",
            "社会信用代码": (cur.org_code if cur else pre.org_code if pre else ""),
            "数据日期": (cur.date if cur else pre.date if pre else ""),
            "币种": "人民币",
            "是否说明": "",
            "计算过程": "",
            "频度": "季",
        })
    rows.sort(key=lambda row: (row["机构名称"], row["指标编码"]))
    return rows


# ---------------------------------------------------------------------------
# 大集中比较
# ---------------------------------------------------------------------------

def _load_central_matrix(central_path: Path) -> tuple[dict[tuple[str, str], float], set[str]]:
    """读集中系统数据 sheet，返回 {(行标签, 指标名): 值} 与 {指标名集合}。"""
    if not central_path.is_file():
        raise PeriodCompareError(f"大集中数据文件不存在：{central_path}")
    book = load_workbook(central_path, read_only=True, data_only=True)
    try:
        if "集中系统数据" not in book.sheetnames:
            raise PeriodCompareError("大集中文件缺少“集中系统数据”工作表")
        sheet = book["集中系统数据"]
        matrix: list[list[Any]] = [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        book.close()
    if not matrix:
        raise PeriodCompareError("“集中系统数据”工作表为空")
    # 表头行：首个含非空文本的行（真实文件里表头在数据行之前）；
    # 行标签列：表头为空但其下各行是文本的那一列。
    header_index = next(
        (
            i for i, row in enumerate(matrix)
            if any(isinstance(v, str) and v.strip() for v in row)
        ),
        None,
    )
    if header_index is None:
        raise PeriodCompareError("“集中系统数据”中未找到表头行")
    header = matrix[header_index]
    label_col = next(
        (
            c for c in range(min(4, len(header)))
            if (header[c] is None or str(header[c]).strip() in {"", " "})
            and any(
                r > header_index and c < len(matrix[r]) and isinstance(matrix[r][c], str) and matrix[r][c].strip()
                for r in range(header_index + 1, min(header_index + 5, len(matrix)))
            )
        ),
        None,
    )
    if label_col is None:
        raise PeriodCompareError("“集中系统数据”中未找到机构行标签列")
    values: dict[tuple[str, str], float] = {}
    headers: set[str] = set()
    for c, title in enumerate(header):
        if c == label_col or title is None or str(title).strip() == "":
            continue
        headers.add(str(title).strip())
        for r in range(header_index + 1, len(matrix)):
            if c >= len(matrix[r]) or matrix[r][c] is None:
                continue
            label = matrix[r][label_col]
            if not isinstance(label, str) or not label.strip():
                continue
            number = _numeric(matrix[r][c])
            if number is not None:
                values[(label.strip(), str(title).strip())] = number
    return values, headers


def _load_central_org_map(central_path: Path) -> dict[str, str]:
    """读参照表：社会信用代码 → 报表项目（大集中行标签）。"""
    book = load_workbook(central_path, read_only=True, data_only=True)
    mapping: dict[str, str] = {}
    try:
        if "参照表" not in book.sheetnames:
            return mapping
        rows = [list(row) for row in book["参照表"].iter_rows(values_only=True)]
    finally:
        book.close()
    if not rows:
        return mapping
    header = [str(v or "").strip() for v in rows[0]]
    try:
        code_col = header.index("统一社会信用代码")
        item_col = header.index("报表项目")
    except ValueError:
        return mapping
    for row in rows[1:]:
        if len(row) <= max(code_col, item_col):
            continue
        code = str(row[code_col] or "").strip()
        item = str(row[item_col] or "").strip()
        if code and item:
            mapping[code] = item
    return mapping


def compare_central(
    current: dict[tuple[str, str], IndicatorValue],
    central_path: Path,
    config: PeriodConfig,
) -> list[dict[str, Any]]:
    matrix, _headers = _load_central_matrix(central_path)
    org_map = _load_central_org_map(central_path)
    rows: list[dict[str, Any]] = []
    for (org_name, code), record in sorted(current.items()):
        indicator = config.indicators.get(code)
        if indicator is None or not indicator.check_central or not indicator.central_name:
            continue
        # VBA 口径：只核对当期有值的指标（零值/空行不读）。
        if not _vba_has_value(record, indicator):
            continue
        org = config.orgs.get(org_name)
        org_code = org.code if org and org.code else record.org_code
        row_label = org_map.get(org_code) or (org.report_item if org else "")
        base_value = _apply_unit(record.value, indicator)
        base_num = _numeric(base_value)
        # 集中系统数据以亿元报送，统一换算成万元再比较（输出列同样存万元）。
        central_value = matrix.get((row_label, indicator.central_name)) if row_label else None
        if central_value is not None:
            central_value = central_value * CENTRAL_UNIT_FACTOR
        if base_num is None:
            continue
        if central_value is None:
            difference, ratio, note = None, None, "谨慎核实"
        else:
            difference = abs(base_num - central_value)
            ratio = (difference / abs(base_num) * 100.0) if base_num else None
            note = "差异超过100元" if difference > CENTRAL_DIFF_TOLERANCE else ""
        rows.append({
            "地区": org.region if org else "",
            "数据属性": indicator.data_type,
            "机构名称": org_name,
            "指标编码": code,
            "指标名称": record.name,
            "基础数据值": base_value,
            "大集中值": central_value,
            "差异绝对值": difference,
            "差异幅度": ratio,
            "是否说明": note,
        })
    return rows


# ---------------------------------------------------------------------------
# 特殊指标与复杂校验规则
# ---------------------------------------------------------------------------

def _period_month(date_text: str) -> tuple[int, int] | None:
    """从文件名日期（2026-06-30）解析 (年, 月)；解析失败返回 None。"""
    match = re.search(r"(\d{4})[-./]?(\d{2})", str(date_text))
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def apply_special_rules(
    rows: list[dict[str, Any]],
    current: dict[tuple[str, str], IndicatorValue],
    previous: dict[tuple[str, str], IndicatorValue],
    config: PeriodConfig,
    *,
    on_step: Callable[[str], None] | None = None,
) -> None:
    """执行特殊指标规则，命中行写“是否说明”。

    第一期支持“当年累计/历史累计指标比上期不应减少”：当期值 < 上期值
    （均换算为万元后）即命中。“当年累计”仅在同一年度内比较（跨年累计
    不可比，跳过并写计算过程说明）；“历史累计”跨年也可比。
    其余 VBA 规则类型读取保留但跳过。
    """
    step = on_step or (lambda _text: None)
    rules = [r for r in config.specials if not r.disabled]
    if not rules:
        return
    supported = [r for r in rules if r.rule_kind in {"当年累计不应减少", "历史累计不应减少"}]
    skipped_kinds = sorted({r.rule_kind for r in rules if r.rule_kind not in {"当年累计不应减少", "历史累计不应减少"}})
    if skipped_kinds:
        step(f"特殊指标：{len(supported)} 条执行，跳过未支持类型（{len(rules) - len(supported)} 条）")
    else:
        step(f"特殊指标：执行 {len(supported)} 条")
    cur_years = {_period_month(record.date) for record in current.values()}
    pre_years = {_period_month(record.date) for record in previous.values()}
    cur_year = next(iter({y[0] for y in cur_years if y}), None)
    pre_year = next(iter({y[0] for y in pre_years if y}), None)
    row_map = {(row["机构名称"], row["指标编码"]): row for row in rows}
    hits = 0
    for rule in supported:
        for org_name in {name for name, _code in row_map}:
            row = row_map.get((org_name, rule.code))
            if row is None:
                continue
            cur = current.get((org_name, rule.code))
            pre = previous.get((org_name, rule.code))
            cur_num = _numeric(_apply_unit(cur.value, config.indicators.get(rule.code))) if cur else None
            pre_num = _numeric(_apply_unit(pre.value, config.indicators.get(rule.code))) if pre else None
            if cur_num is None or pre_num is None:
                continue
            same_year = cur_year is not None and cur_year == pre_year
            if rule.rule_kind == "当年累计不应减少" and not same_year:
                row["计算过程"] = (
                    f"跨年累计不比较（当期 {cur.date} / 上期 {pre.date}）"
                    if row.get("计算过程") in (None, "")
                    else row["计算过程"]
                )
                continue
            if cur_num < pre_num:
                row["是否说明"] = rule.remark
                hits += 1
    step(f"特殊指标命中 {hits} 行")


_PLACEHOLDER = re.compile(r"(\[|\{)([^\]}]*)(\]|\})")


class _V(float):
    """表达式求值的包装数值：相等比较带容差，消除二进制浮点误报。

    值已按万元两位舍入（百元精度），``==``/``!=`` 采用 0.005 万元
    （50 元）容差——勾稽关系两侧十进制相同时不应判“不平衡”。
    ``<``/``>`` 等严格比较保持浮点原语义。
    """

    __slots__ = ()

    def _tolerant_eq(self, other) -> bool:
        try:
            return abs(float(self) - float(other)) <= 0.005
        except (TypeError, ValueError):
            return False

    def __eq__(self, other):
        return self._tolerant_eq(other)

    def __ne__(self, other):
        return not self._tolerant_eq(other)

    def __hash__(self):
        return hash(float(self))

    def __add__(self, other):
        return _V(float(self) + float(other))

    def __radd__(self, other):
        return _V(float(other) + float(self))

    def __sub__(self, other):
        return _V(float(self) - float(other))

    def __rsub__(self, other):
        return _V(float(other) - float(self))

    def __mul__(self, other):
        return _V(float(self) * float(other))

    def __rmul__(self, other):
        return _V(float(other) * float(self))

    def __truediv__(self, other):
        return _V(float(self) / float(other))

    def __rtruediv__(self, other):
        return _V(float(other) / float(self))


def _AND(*values) -> bool:
    return all(bool(v) for v in values)


def _OR(*values) -> bool:
    return any(bool(v) for v in values)


def _NOT(value) -> bool:
    return not bool(value)


def _eval_complex_expression(
    expression: str,
    cur_lookup: Callable[[str], float],
    pre_lookup: Callable[[str], float],
) -> tuple[bool, str]:
    """求值一条复杂校验表达式，返回 (是否命中, 代入后的表达式)。

    `[...]` 占位符取当期值、`{...}` 取上期值，均按第 4 段指标代码取该机构
    的值；缺省指标记 0，值为 0 时按 VBA 口径替换为 0.01（避免除零）。
    数值四舍五入两位（VBA RoundRule）。Excel 风格 AND/OR/NOT 与 `<>` 转
    成 Python 求值；表达式来自用户自己的配置簿，与 VBA Evaluate 等价。
    """

    def replace(match: re.Match) -> str:
        parts = [p.strip() for p in match.group(2).split(",")]
        code = parts[3] if len(parts) > 3 else ""
        lookup = cur_lookup if match.group(1) == "[" else pre_lookup
        value = lookup(code)
        return f"_V({0.01 if value == 0 else round(value, 2)!r})"

    filled = _PLACEHOLDER.sub(replace, expression)
    eval_text = filled.replace("<>", "!=")                          # Excel 不等号
    eval_text = re.sub(r"(?<![<>!])=(?!=)", "==", eval_text)        # 单独 = 为相等比较
    for excel, py in (("AND", "_AND"), ("OR", "_OR"), ("NOT", "_NOT")):
        eval_text = re.sub(rf"\b{excel}\b", py, eval_text)
    try:
        result = eval(eval_text, {"__builtins__": {}}, {"_AND": _AND, "_OR": _OR, "_NOT": _NOT, "_V": _V})  # noqa: S307 - 用户配置簿内的受控表达式
        return bool(result), filled
    except Exception:
        return False, filled


def apply_complex_rules(
    rows: list[dict[str, Any]],
    current: dict[tuple[str, str], IndicatorValue],
    previous: dict[tuple[str, str], IndicatorValue],
    config: PeriodConfig,
    *,
    on_step: Callable[[str], None] | None = None,
) -> None:
    """执行复杂校验表达式：逐机构代入指标值，命中行写“是否说明+计算过程”。

    取反标识（VBA qfbz）规则第一期跳过（保留数据，日志注明）。
    """
    step = on_step or (lambda _text: None)
    rules = [r for r in config.complex_rules if not r.disabled]
    inverted = sum(1 for r in rules if r.invert)
    forward = len(rules) - inverted
    if not rules:
        return
    if inverted:
        step(f"复杂校验：执行 {forward} 条（跳过 {inverted} 条取反规则）")
    else:
        step(f"复杂校验：执行 {forward} 条")

    def values_by_org(values: dict[tuple[str, str], IndicatorValue]) -> dict[str, dict[str, float]]:
        result: dict[str, dict[str, float]] = {}
        for (org, code), record in values.items():
            number = _numeric(_apply_unit(record.value, config.indicators.get(code)))
            if number is None:
                continue
            result.setdefault(org, {})[code] = number
        return result

    cur_by_org = values_by_org(current)
    pre_by_org = values_by_org(previous)
    row_map = {(row["机构名称"], row["指标编码"]): row for row in rows}
    hit_rules = 0
    for rule in rules:
        if rule.invert:
            continue
        involved_codes = {
            parts[3].strip()
            for parts in (m.group(2).split(",") for m in _PLACEHOLDER.finditer(rule.rule))
            if len(parts) > 3 and parts[3].strip()
        }
        rule_hit_any_org = False
        for org in sorted(set(cur_by_org) | set(pre_by_org)):
            cur_vals = cur_by_org.get(org, {})
            pre_vals = pre_by_org.get(org, {})
            hit, filled = _eval_complex_expression(
                rule.rule,
                lambda code, _v=cur_vals: _v.get(code, 0.0),
                lambda code, _v=pre_vals: _v.get(code, 0.0),
            )
            if not hit:
                continue
            rule_hit_any_org = True
            for code in involved_codes:
                row = row_map.get((org, code))
                if row is not None and not row.get("是否说明"):
                    row["是否说明"] = rule.desc
                    row["计算过程"] = f"复杂校验命中：{filled}"
        if rule_hit_any_org:
            hit_rules += 1
    step(f"复杂校验命中 {hit_rules} 条规则")




def write_output(
    output_path: Path,
    period_rows: list[dict[str, Any]],
    central_rows: list[dict[str, Any]],
) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "跨期比较"
    sheet.append(PERIOD_SHEET_HEADERS)
    for row in period_rows:
        sheet.append([row.get(header, "") for header in PERIOD_SHEET_HEADERS])
    central_sheet = book.create_sheet("大集中比较")
    central_sheet.append(CENTRAL_SHEET_HEADERS)
    for row in central_rows:
        central_sheet.append([row.get(header, "") for header in CENTRAL_SHEET_HEADERS])
    for column in range(1, len(PERIOD_SHEET_HEADERS) + 1):
        sheet.column_dimensions[sheet.cell(row=1, column=column).column_letter].width = 14
    central_sheet.column_dimensions["A"].width = 14
    output_path.parent.mkdir(parents=True, exist_ok=True)
    book.save(output_path)


def run_period_compare(
    *,
    current_dir: Path,
    previous_dir: Path,
    central_path: Path | None,
    output_dir: Path,
    config_path: Path,
    on_step: Callable[[str], None] | None = None,
) -> Path:
    """执行跨期比较并输出工作簿；返回输出文件路径。"""
    step = on_step or (lambda _text: None)
    config = load_period_config(config_path)
    step(f"已加载配置：指标 {len(config.indicators)} 项，机构 {len(config.orgs)} 家，警戒区间 {len(config.alerts)} 档")
    current = load_period_directory(current_dir, label="当期", on_step=step)
    previous = load_period_directory(previous_dir, label="上期", on_step=step)
    step(f"当期指标值 {len(current)} 条，上期指标值 {len(previous)} 条")
    period_rows = compare_periods(current, previous, config)
    step(f"两期比较完成：{len(period_rows)} 行")
    apply_special_rules(period_rows, current, previous, config, on_step=step)
    apply_complex_rules(period_rows, current, previous, config, on_step=step)
    central_rows: list[dict[str, Any]] = []
    if central_path is not None:
        central_rows = compare_central(current, central_path, config)
        flagged = sum(1 for row in central_rows if row["是否说明"])
        step(f"大集中核对完成：{len(central_rows)} 行，其中 {flagged} 行需要说明")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"跨期比较_{timestamp}.xlsx"
    write_output(output_path, period_rows, central_rows)
    return output_path
