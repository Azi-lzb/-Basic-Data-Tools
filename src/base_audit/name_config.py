from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


MAPPING_SHEET = "模块化功能"
EXTERNAL_FILE_FUNCTION = "修改_外部文件添加"
FORMULA_COPY_FUNCTION = "修改_公式校验复制"
ISSUE_EXTRACT_FUNCTION = "汇总_公式校验结果提取"
STRUCTURE_COMPARE_FUNCTION = "核对_表结构比对"
# 旧版“检查_表结构检查”类型名，初始化配置时自动迁移为新类型名。
LEGACY_STRUCTURE_COMPARE_FUNCTION = "检查_表结构检查"
USED_RANGE_SUMMARY_FUNCTION = "汇总_任意行汇总"
FIXED_ROW_SUMMARY_FUNCTION = "汇总_固定行汇总"
WORKBOOK_TABLE_MERGE_FUNCTION = "汇总_汇总表合并"
NAMED_RANGE_CHECK_FUNCTION = "检查_命名区域存在性"
MERGE_ORG_FILES_FUNCTION = "修改_合并同机构多表"

FEATURE_TYPES = (
    EXTERNAL_FILE_FUNCTION,
    FORMULA_COPY_FUNCTION,
    ISSUE_EXTRACT_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    USED_RANGE_SUMMARY_FUNCTION,
    FIXED_ROW_SUMMARY_FUNCTION,
    WORKBOOK_TABLE_MERGE_FUNCTION,
    NAMED_RANGE_CHECK_FUNCTION,
    MERGE_ORG_FILES_FUNCTION,
)
FLOW_SHEET = "执行流程"
FLOW_HEADERS = (
    "流程名", "顺序", "功能名", "启用", "失败后处理", "是否输出结果", "输出文件名", "处理对象", "备注",
)
# 7 列旧配置（无“输出文件名/处理对象”列）仍视为合法：加载时缺列回退空串，
# 运行 initialize_config() 时增量迁移。
FLOW_LEGACY_HEADERS = (
    "流程名", "顺序", "功能名", "启用", "失败后处理", "是否输出结果", "备注",
)
# 上一版 8 列配置（有“输出文件名”列、无“处理对象”列）。
FLOW_PREVIOUS_HEADERS = (
    "流程名", "顺序", "功能名", "启用", "失败后处理", "是否输出结果", "输出文件名", "备注",
)
# 上一版 9 列配置（“处理对象”列还叫“待处理对象”）。
FLOW_PREVIOUS_9_HEADERS = (
    "流程名", "顺序", "功能名", "启用", "失败后处理", "是否输出结果", "输出文件名", "待处理对象", "备注",
)
CUSTOM_FEATURE_SHEET = "自定义按钮"
CUSTOM_FEATURE_HEADERS = ("按钮名称", "流程名称", "是否显示", "备注")
LEGACY_CUSTOM_FEATURE_HEADERS = ("功能标识", "功能名称", "是否显示", "备注")
MAPPING_HEADERS = (
    "功能名", "功能类型", "命名区域名", "是否限定工作簿", "工作簿关键字", "备注"
)
DEFAULT_FEATURES = (
    ("检查校验区域", NAMED_RANGE_CHECK_FUNCTION, "校验区域", "否", "", "检查模板是否存在“校验区域”命名区域；可填写多个区域名，均必须存在"),
    ("检查表结构区域", NAMED_RANGE_CHECK_FUNCTION, "表结构区域", "否", "", "检查模板是否存在“表结构区域”命名区域"),
    ("检查汇总区域", NAMED_RANGE_CHECK_FUNCTION, "任意行汇总区域、表头区域", "否", "", "检查汇总模板是否同时存在汇总区域和表头区域"),
    ("外部文件添加", EXTERNAL_FILE_FUNCTION, "", "否", "", "将外部工作表复制至审核副本，必须先于公式复制"),
    ("公式校验复制", FORMULA_COPY_FUNCTION, "校验区域", "否", "", "复制公式及格式到审核副本，可输出同名审核副本目录"),
    ("校验结果提取", ISSUE_EXTRACT_FUNCTION, "校验区域", "否", "", "重算后提取错误、核实、提示等结果"),
    ("表结构比对", STRUCTURE_COMPARE_FUNCTION, "表结构区域", "否", "", "逐格精确核对固定表结构"),
    ("任意行汇总", USED_RANGE_SUMMARY_FUNCTION, "任意行汇总区域", "否", "", "模板必须定义同工作表的“表头区域”；“单元格区域.字段名”会附加到每行"),
    ("固定行汇总", FIXED_ROW_SUMMARY_FUNCTION, "固定行汇总区域", "否", "", "模板必须定义同工作表的“表头区域”；严格按框选的固定行汇总"),
    ("汇总表合并", WORKBOOK_TABLE_MERGE_FUNCTION, "", "否", "", "读取所有源工作簿的工作表；表头相同才合并，不同表头分别保留在同一结果工作簿中"),
    ("合并同机构多表", MERGE_ORG_FILES_FUNCTION, "", "否", "", "把源目录下同一机构（文件名主名下划线第一段相同）的所有工作簿合并为一本，所有工作表集中存放，供制作模板后单独跑汇总核查表校验"),
)
DEFAULT_FLOWS = (
    ("汇总核查表校验", 10, "检查校验区域", "是", "停止", "否", "", "", "先确认公式复制和问题提取所需区域存在"),
    ("汇总核查表校验", 20, "检查表结构区域", "是", "停止", "否", "", "", "先确认表结构比对所需区域存在"),
    ("汇总核查表校验", 30, "表结构比对", "是", "停止", "否", "", "", "不匹配文件跳过，正常运行只记入运行日志"),
    ("汇总核查表校验", 40, "外部文件添加", "是", "停止", "否", "", "", "外部表先于公式复制；需要排错时再输出阶段副本"),
    ("汇总核查表校验", 50, "公式校验复制", "是", "停止", "否", "", "外部文件添加", "复制模板公式并强制重算；需要排错时再输出审核副本"),
    ("汇总核查表校验", 60, "校验结果提取", "是", "停止", "是", "汇总核查表校验", "公式校验复制", "输出正式本期审核结果"),
    ("汇总校验结果说明", 10, "检查表结构区域", "是", "停止", "否", "", "", "先确认表结构比对所需区域存在"),
    ("汇总校验结果说明", 20, "检查汇总区域", "是", "停止", "否", "", "", "先确认汇总区域和表头区域存在"),
    ("汇总校验结果说明", 30, "表结构比对", "是", "停止", "否", "", "", "不匹配文件跳过，正常运行只记入运行日志"),
    ("汇总校验结果说明", 40, "任意行汇总", "是", "停止", "是", "汇总校验结果说明", "", "输出正式校验结果与报送说明汇总"),
    ("合并同机构多表", 10, "合并同机构多表", "是", "停止", "是", "合并同机构多表", "", "递归合并同一机构的多个报送工作簿为一本，输出每机构一个 xlsx"),
)
# 旧配置迁移时，为“默认最终输出步骤”回填流程名（键 = 流程名, 功能名）。
DEFAULT_OUTPUT_NAMES = {(row[0], row[2]): row[6] for row in DEFAULT_FLOWS if row[6]}
# 旧配置迁移时，为“默认功能”回填处理对象（键 = 流程名, 功能名）。
DEFAULT_PROCESS_SOURCES = {(row[0], row[2]): row[7] for row in DEFAULT_FLOWS if row[7]}


# 配置表头批注：告诉使用者每列怎么填、如何影响功能输出。
GUIDE_COMMENTS: dict[str, dict[int, str]] = {
    MAPPING_SHEET: {
        1: "本模块显示名，必须唯一；执行流程的“功能名”列用此名引用。",
        2: "功能类型决定模块行为。可选：\n"
           "检查_命名区域存在性：检查模板是否存在指定命名区域，运行日志会记录每个区域所在工作表与覆盖区域；\n"
           "核对_表结构比对：逐格核对固定表头，不匹配的报送文件跳过；\n"
           "修改_外部文件添加：把外部工作表复制进审核副本，须先于公式校验复制；\n"
           "修改_公式校验复制：复制模板公式并强制重算，可输出同名审核副本目录；\n"
           "修改_合并同机构多表：按文件名第一段识别机构，将同一机构多个工作簿的所有工作表合并为一本；\n"
           "汇总_公式校验结果提取：重算后提取错误、核实、提示等结果；\n"
           "汇总_任意行汇总 / 汇总_固定行汇总 / 汇总_汇总表合并：区域汇总。",
        3: "模板中要使用的命名区域名，可填多个用、分隔；对应功能按此名称在模板名称管理器中查找。",
        4: "填“是”时，本模块只在“工作簿关键字”匹配当前模板文件名时才启用。",
        5: "配合“是否限定工作簿=是”使用的模板名关键字。",
        6: "自由说明，也会显示在自定义按钮的悬浮提示等位置。",
    },
    FLOW_SHEET: {
        1: "流程名即界面主按钮文字，应唯一；自定义按钮的“流程名称”也引用此名。",
        2: "步骤执行顺序，建议 10、20、30… 递增。",
        3: "引用“模块化功能”页中的功能名。",
        4: "填“是”才执行该步骤。",
        5: "仅支持“停止”或“跳过”：停止=该步骤出错则中止整个流程；跳过=出错则跳过当前文件继续。",
        6: "只对会产出实际文件的功能生效：外部文件添加 / 公式校验复制输出同名阶段副本目录、审核副本目录，校验结果提取和合并同机构多表输出实际工作簿；填“否”时这些中间文件处理完即清理。检查类与表结构比对始终只记入运行日志，不受此列影响。",
        7: "输出文件名（可选）：仅对“是否输出结果=是”且会产出实际文件的步骤生效；填写后，该步骤的输出文件/目录名从“序号_功能名_时间戳”改为“输出文件名_时间戳”。校验结果提取、任意行/固定行/汇总表合并、合并同机构多表等最终结果建议填流程名；外部文件添加、公式校验复制的中间阶段副本默认留空。",
        8: "处理对象（可选）：本步骤读哪份数据。留空=源数据目录的报送文件；填流程中排在前面的功能名，则读该功能产出的副本，例如校验结果提取填“公式校验复制”读套了公式的审核副本，填“外部文件添加”读加了外部表的副本；任意行/固定行/汇总表合并填“公式校验复制”可对审核副本做汇总。主流程的校验结果提取预填“公式校验复制”、公式校验复制预填“外部文件添加”。处理对象只认流程中排在前面的外部文件添加/公式校验复制等会产副本的功能；引用不匹配或不产副本的功能（如检查类、核对类）时，自动跳过非副本功能，兜底到最近前一个产副本功能（之前没有则兜底到源数据目录），并在运行日志给出提示，不会中断流程。",
        9: "说明。",
    },
    CUSTOM_FEATURE_SHEET: {
        1: "WebView 按钮文字，应唯一。",
        2: "指向“执行流程”页的流程名，点击后执行该流程。",
        3: "填“是”才在界面显示该按钮。",
        4: "说明。",
    },
}


@dataclass(frozen=True)
class FeatureMapping:
    name: str
    feature_type: str
    range_names: tuple[str, ...]
    workbook_limited: bool
    workbook_keyword: str
    remark: str


@dataclass(frozen=True)
class FlowStep:
    flow_name: str
    order: int
    feature_name: str
    on_failure: str
    output_result: bool
    remark: str
    output_name: str = ""
    process_source: str = ""


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _split_names(value: str) -> tuple[str, ...]:
    return tuple(item.strip() for item in re.split(r"[、，,；;\n]", value) if item.strip())


def _limited(value: str) -> bool:
    return value.casefold() in {"是", "y", "yes", "true", "1"}


def _apply_guide_comments(workbook: object) -> None:
    """给配置表头加批注，说明每列怎么填、如何影响功能输出。

    重置或初始化配置后仍然保留，方便直接在 Excel 里查看。
    """
    from openpyxl.comments import Comment

    for sheet_name, column_comments in GUIDE_COMMENTS.items():
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        for column, text in column_comments.items():
            sheet.cell(1, column).comment = Comment(
                text, "基础数据审核工具", height=240, width=380
            )


def initialize_config(config_path: Path) -> Path:
    """Restore default configuration structure without touching history sheets."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from .history import HISTORY_HEADERS
    from .template import normalize_template_name

    config_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = load_workbook(config_path) if config_path.exists() else Workbook()
    try:
        for obsolete in ("区域映射", "自定义功能", "外部工作表"):
            if obsolete in workbook.sheetnames:
                del workbook[obsolete]
        if MAPPING_SHEET in workbook.sheetnames:
            sheet = workbook[MAPPING_SHEET]
            if tuple(_text(cell.value) for cell in sheet[1]) != MAPPING_HEADERS:
                raise RuntimeError("无法初始化配置：‘模块化功能’表头不符合规范")
        else:
            sheet = workbook.active if workbook.active.title == "Sheet" and len(workbook.sheetnames) == 1 else workbook.create_sheet(MAPPING_SHEET, 0)
            sheet.title = MAPPING_SHEET
            sheet.append(MAPPING_HEADERS)
        # 旧功能名“公式校验”统一迁移为“公式校验复制”，避免新旧默认值并存。
        for row_number in range(2, sheet.max_row + 1):
            name_cell = sheet.cell(row_number, 1)
            type_cell = sheet.cell(row_number, 2)
            if _text(name_cell.value) == "公式校验" and _text(type_cell.value) == FORMULA_COPY_FUNCTION:
                name_cell.value = "公式校验复制"
        existing_features = {_text(row[0]) for row in sheet.iter_rows(min_row=2, values_only=True)}
        for row in DEFAULT_FEATURES:
            if row[0] not in existing_features:
                sheet.append(row)
        for row_number in range(2, sheet.max_row + 1):
            type_cell = sheet.cell(row_number, 2)
            if _text(type_cell.value) == LEGACY_STRUCTURE_COMPARE_FUNCTION:
                type_cell.value = STRUCTURE_COMPARE_FUNCTION  # 旧“检查_表结构检查”迁移为新类型名
            if _text(type_cell.value) == STRUCTURE_COMPARE_FUNCTION:
                sheet.cell(row_number, 3, "表结构区域")
                sheet.cell(row_number, 6, "逐格精确核对固定表结构")
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center")
        for column, width in {"A": 18, "B": 16, "C": 24, "D": 18, "E": 28, "F": 42}.items():
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:F{max(2, sheet.max_row)}"
        if FLOW_SHEET in workbook.sheetnames:
            flow_sheet = workbook[FLOW_SHEET]
            actual_headers = tuple(_text(cell.value) for cell in flow_sheet[1])
            if actual_headers != FLOW_HEADERS:
                if actual_headers in (FLOW_LEGACY_HEADERS, FLOW_PREVIOUS_HEADERS, FLOW_PREVIOUS_9_HEADERS):
                    # 旧配置增量迁移，保留用户既有流程行：
                    # 7 列先插入“输出文件名”，再插入“处理对象”；8 列只插入“处理对象”；
                    # 9 列只是把“待处理对象”列改名“处理对象”。
                    if actual_headers == FLOW_LEGACY_HEADERS:
                        flow_sheet.insert_cols(7)
                        flow_sheet.cell(1, 7, "输出文件名")
                    if actual_headers != FLOW_PREVIOUS_9_HEADERS:
                        flow_sheet.insert_cols(8)
                    flow_sheet.cell(1, 8, "处理对象")
                    for row in flow_sheet.iter_rows(min_row=2):
                        if len(row) < len(FLOW_HEADERS):
                            continue
                        if not _text(row[6].value):
                            default = DEFAULT_OUTPUT_NAMES.get((_text(row[0].value), _text(row[2].value)))
                            if default:
                                row[6].value = default
                        if not _text(row[7].value):
                            default = DEFAULT_PROCESS_SOURCES.get((_text(row[0].value), _text(row[2].value)))
                            if default:
                                row[7].value = default
                else:
                    if flow_sheet.max_column > len(FLOW_HEADERS):
                        flow_sheet.delete_cols(len(FLOW_HEADERS) + 1, flow_sheet.max_column - len(FLOW_HEADERS))
                    flow_sheet.delete_rows(1, flow_sheet.max_row)
                    flow_sheet.append(FLOW_HEADERS)
        else:
            flow_sheet = workbook.create_sheet(FLOW_SHEET)
            flow_sheet.append(FLOW_HEADERS)
        existing_steps = {( _text(row[0]), _text(row[2])) for row in flow_sheet.iter_rows(min_row=2, values_only=True)}
        for row in DEFAULT_FLOWS:
            if (row[0], row[2]) not in existing_steps:
                flow_sheet.append(row)
        # 极早期的 9 列试验版曾把“中间副本保留标记”错位写进“处理对象”列，
        # 留下 1、2、3… 等数字。处理对象只能引用前置模块名，数字不可能有效；
        # 初始化时只清理数值型遗留项，不改用户填写的正常文字引用。
        for row in flow_sheet.iter_rows(min_row=2):
            if isinstance(row[7].value, (int, float)):
                row[7].value = ""
        # 执行流程里旧名“公式校验”同步迁移为“公式校验复制”。
        for row in flow_sheet.iter_rows(min_row=2):
            if _text(row[2].value) == "公式校验":
                row[2].value = "公式校验复制"
        for cell in flow_sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for column, width in {"A": 20, "B": 10, "C": 20, "D": 10, "E": 16, "F": 16, "G": 18, "H": 18, "I": 38}.items():
            flow_sheet.column_dimensions[column].width = width
        flow_sheet.freeze_panes = "A2"
        flow_sheet.auto_filter.ref = f"A1:I{max(2, flow_sheet.max_row)}"

        # 自定义按钮由用户维护；按新规范创建空表。
        if CUSTOM_FEATURE_SHEET not in workbook.sheetnames:
            custom_sheet = workbook.create_sheet(CUSTOM_FEATURE_SHEET)
            legacy_rows: list[tuple[object, ...]] = []
        else:
            custom_sheet = workbook[CUSTOM_FEATURE_SHEET]
            existing_headers = tuple(_text(cell.value) for cell in custom_sheet[1])
            # 修复早期初始化产生的空首行。
            second_headers = tuple(_text(cell.value) for cell in custom_sheet[2])
            if not any(existing_headers) and second_headers == CUSTOM_FEATURE_HEADERS:
                custom_sheet.delete_rows(1, 1)
                existing_headers = tuple(_text(cell.value) for cell in custom_sheet[1])
            legacy_rows = []
            if existing_headers and existing_headers != CUSTOM_FEATURE_HEADERS:
                raise RuntimeError(
                    "无法初始化配置：‘自定义按钮’表头应为“按钮名称、流程名称、是否显示、备注”"
                )
        if not any(cell.value for cell in custom_sheet[1]):
            for column, header in enumerate(CUSTOM_FEATURE_HEADERS, start=1):
                custom_sheet.cell(row=1, column=column, value=header)
        for cell in custom_sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for column, width in {"A": 24, "B": 24, "C": 12, "D": 42}.items():
            custom_sheet.column_dimensions[column].width = width
        custom_sheet.freeze_panes = "A2"
        custom_sheet.auto_filter.ref = f"A1:D{max(2, custom_sheet.max_row)}"

        # 历史表与配置文件一同维护；初始化配置不会删除既有记录。
        history_sheet_name = "历史审核结果"
        if history_sheet_name not in workbook.sheetnames:
            if "历史审核记录" in workbook.sheetnames:
                history_sheet = workbook["历史审核记录"]
                history_sheet.title = history_sheet_name
            else:
                history_sheet = workbook.create_sheet(history_sheet_name)
                history_sheet.append(HISTORY_HEADERS)
        else:
            history_sheet = workbook[history_sheet_name]
        current_headers = tuple(_text(cell.value) for cell in history_sheet[1])
        if current_headers != HISTORY_HEADERS:
            old_rows = list(history_sheet.iter_rows(min_row=2, values_only=True))
            positions = {name: index for index, name in enumerate(current_headers) if name}

            def old_value(row: tuple[object, ...], *names: str) -> str:
                for name in names:
                    index = positions.get(name)
                    if index is not None and index < len(row):
                        value = _text(row[index])
                        if value:
                            return value
                return ""

            converted_rows: list[tuple[str, ...]] = []
            for row in old_rows:
                book = old_value(row, "工作簿名", "来源工作簿", "源文件")
                book_name = normalize_template_name(Path(book).stem) if book else "工作簿未识别"
                sheet_name = old_value(row, "工作表名", "工作表")
                location = old_value(row, "定位单元格")
                formula_cell = old_value(row, "公式单元格") or location
                indicator = old_value(row, "校验指标", "校验字段")
                existing_rule = old_value(row, "规则编号", "问题标识", "问题ID")
                rule_id = re.sub(r"｜\d+$", "", existing_rule) if existing_rule else "｜".join(
                    (book_name, sheet_name or "工作表未识别", formula_cell or "公式单元格未识别", indicator or "校验指标未识别")
                )
                if not any((book, sheet_name, location, existing_rule)):
                    continue
                converted_rows.append((
                    Path(book).name if book else "",
                    sheet_name,
                    location,
                    old_value(row, "错误类型", "级别"),
                    indicator,
                    old_value(row, "描述", "详细说明", "问题说明"),
                    old_value(row, "当前值"),
                    old_value(row, "对比值"),
                    old_value(row, "差值"),
                    rule_id,
                    old_value(row, "历史校验说明", "机构反馈"),
                    old_value(row, "审核意见"),
                ))
            history_sheet.delete_rows(1, history_sheet.max_row)
            history_sheet.append(HISTORY_HEADERS)
            for row in converted_rows:
                history_sheet.append(row)
        for cell in history_sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        history_sheet.freeze_panes = "A2"
        history_sheet.auto_filter.ref = f"A1:L{max(2, history_sheet.max_row)}"
        _apply_guide_comments(workbook)
        try:
            workbook.save(config_path)
        except PermissionError as exc:
            raise RuntimeError(f"无法初始化配置：请先关闭已打开的配置文件“{config_path.name}”后重试") from exc
    finally:
        workbook.close()
    return config_path


def reset_default_configuration(config_path: Path) -> Path:
    """Reset only executable module/flow sheets, keeping user-maintained sheets intact."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment, Font, PatternFill
    from .history import HISTORY_HEADERS

    config_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = load_workbook(config_path) if config_path.exists() else Workbook()
    try:
        def get_or_create(sheet_name: str, position: int | None = None):
            if sheet_name in workbook.sheetnames:
                return workbook[sheet_name]
            if workbook.active.title == "Sheet" and len(workbook.sheetnames) == 1:
                sheet = workbook.active
                sheet.title = sheet_name
                return sheet
            return workbook.create_sheet(sheet_name, position)

        def rewrite(sheet_name: str, headers: tuple[str, ...], rows: tuple[tuple[object, ...], ...], widths: dict[str, int]) -> None:
            sheet = get_or_create(sheet_name, 0 if sheet_name == MAPPING_SHEET else None)
            if sheet.max_column > len(headers):
                sheet.delete_cols(len(headers) + 1, sheet.max_column - len(headers))
            sheet.delete_rows(1, sheet.max_row)
            sheet.append(headers)
            for row in rows:
                sheet.append(row)
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center")
            for column, width in widths.items():
                sheet.column_dimensions[column].width = width
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(2, sheet.max_row)}"

        rewrite(
            MAPPING_SHEET, MAPPING_HEADERS, DEFAULT_FEATURES,
            {"A": 18, "B": 16, "C": 24, "D": 18, "E": 28, "F": 42},
        )
        rewrite(
            FLOW_SHEET, FLOW_HEADERS, DEFAULT_FLOWS,
            {"A": 20, "B": 10, "C": 20, "D": 10, "E": 16, "F": 16, "G": 18, "H": 18, "I": 38},
        )
        # These are user-maintained sheets. Create only when absent; never
        # rewrite existing values or headers during a reset.
        if CUSTOM_FEATURE_SHEET not in workbook.sheetnames:
            custom_sheet = workbook.create_sheet(CUSTOM_FEATURE_SHEET)
            custom_sheet.append(CUSTOM_FEATURE_HEADERS)
            for cell in custom_sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
        if "历史审核结果" not in workbook.sheetnames:
            history_sheet = workbook.create_sheet("历史审核结果")
            history_sheet.append(HISTORY_HEADERS)
            for cell in history_sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
        _apply_guide_comments(workbook)
        workbook.save(config_path)
    except PermissionError as exc:
        raise RuntimeError(f"无法重置配置：请先关闭已打开的配置文件“{config_path.name}”后重试") from exc
    finally:
        workbook.close()
    return config_path


def load_feature_mappings(config_path: Path | None, template_path: Path) -> list[FeatureMapping]:
    """Load executable feature rows, keeping template-specific rows preferred."""
    if not config_path or not config_path.is_file():
        return [
            FeatureMapping(row[0], row[1], _split_names(row[2]), False, "", row[5])
            for row in DEFAULT_FEATURES
        ]
    from openpyxl import load_workbook
    workbook = load_workbook(config_path, read_only=True, data_only=True)
    try:
        if MAPPING_SHEET not in workbook.sheetnames:
            return []
        sheet = workbook[MAPPING_SHEET]
        rows = sheet.iter_rows(values_only=True)
        headers = [_text(value) for value in next(rows, ())]
        missing = [header for header in MAPPING_HEADERS if header not in headers]
        if missing:
            raise ValueError("config.xlsx 的“模块化功能”缺少列：" + "、".join(missing))
        positions = {header: headers.index(header) for header in headers}
        template_name = template_path.name.casefold()
        selected: dict[tuple[str, str], tuple[int, FeatureMapping]] = {}
        for row in rows:
            def get(name: str) -> str:
                index = positions[name]
                return _text(row[index]) if index < len(row) else ""

            feature_type = get("功能类型")
            feature_name = get("功能名")
            if not feature_name or feature_type not in FEATURE_TYPES:
                continue
            limited = _limited(get("是否限定工作簿"))
            keyword = get("工作簿关键字")
            if limited and (not keyword or keyword.casefold() not in template_name):
                continue
            feature = FeatureMapping(feature_name, feature_type, _split_names(get("命名区域名")), limited, keyword, get("备注"))
            identity = (feature_name, feature_type)
            priority = 2 if limited else 1
            if identity not in selected or priority >= selected[identity][0]:
                selected[identity] = (priority, feature)
        return [item[1] for item in selected.values()]
    finally:
        workbook.close()


def features_of_type(features: list[FeatureMapping], feature_type: str) -> list[FeatureMapping]:
    return [feature for feature in features if feature.feature_type == feature_type]


def load_flow_features(config_path: Path, flow_name: str) -> tuple[str, ...]:
    """Read enabled feature names in configured order for one execution entry."""
    return tuple(step.feature_name for step in load_flow_steps(config_path, flow_name))


def load_flow_steps(config_path: Path, flow_name: str) -> tuple[FlowStep, ...]:
    """Read enabled executable steps and their failure policy for one flow."""
    from openpyxl import load_workbook
    if not config_path.is_file():
        return ()
    workbook = load_workbook(config_path, read_only=True, data_only=True)
    try:
        if FLOW_SHEET not in workbook.sheetnames:
            return ()
        rows = workbook[FLOW_SHEET].iter_rows(values_only=True)
        headers = [_text(value) for value in next(rows, ())]
        if any(header not in headers for header in FLOW_LEGACY_HEADERS):
            raise ValueError("config.xlsx 的“执行流程”缺少必要列：" + "、".join(header for header in FLOW_LEGACY_HEADERS if header not in headers))
        index = {header: headers.index(header) for header in headers}
        process_source_col = index.get("处理对象", index.get("待处理对象"))
        selected: list[FlowStep] = []
        for row in rows:
            if _text(row[index["流程名"]]) != flow_name or not _limited(_text(row[index["启用"]])):
                continue
            feature_name = _text(row[index["功能名"]])
            if feature_name:
                on_failure = _text(row[index["失败后处理"]]) or "停止"
                if on_failure not in {"停止", "跳过"}:
                    raise ValueError(
                        f"执行流程“{flow_name}”的“失败后处理”只能填写“停止”或“跳过”"
                    )
                selected.append(
                    FlowStep(
                        flow_name=flow_name,
                        order=int(float(row[index["顺序"]] or 0)),
                        feature_name=feature_name,
                        on_failure=on_failure,
                        output_result=_limited(_text(row[index["是否输出结果"]])),
                        remark=_text(row[index.get("备注", -1)]) if "备注" in index else "",
                        output_name=_text(row[index["输出文件名"]]) if "输出文件名" in index and index["输出文件名"] < len(row) else "",
                        process_source=_text(row[process_source_col]) if process_source_col is not None and process_source_col < len(row) else "",
                    )
                )
        return tuple(sorted(selected, key=lambda item: item.order))
    finally:
        workbook.close()


def validate_config(config_path: Path) -> list[str]:
    """Return human-readable configuration errors; never changes user data."""
    from openpyxl import load_workbook
    errors: list[str] = []
    workbook = load_workbook(config_path, read_only=True, data_only=True)
    try:
        for sheet_name, headers in ((MAPPING_SHEET, MAPPING_HEADERS), (FLOW_SHEET, FLOW_HEADERS), (CUSTOM_FEATURE_SHEET, CUSTOM_FEATURE_HEADERS)):
            if sheet_name not in workbook.sheetnames:
                errors.append(f"缺少“{sheet_name}”工作表")
                continue
            actual = tuple(_text(cell.value) for cell in workbook[sheet_name][1])
            if sheet_name == FLOW_SHEET:
                # 7 列、8 列及旧 9 列（“处理对象”还叫“待处理对象”）配置仍视为合法。
                if actual not in (FLOW_HEADERS, FLOW_LEGACY_HEADERS, FLOW_PREVIOUS_HEADERS, FLOW_PREVIOUS_9_HEADERS):
                    errors.append(f"“{sheet_name}”表头不符合规范")
            elif actual != headers:
                errors.append(f"“{sheet_name}”表头不符合规范")
        if errors:
            return errors
        features = [_text(row[0]) for row in workbook[MAPPING_SHEET].iter_rows(min_row=2, values_only=True) if _text(row[0])]
        if len(features) != len(set(features)):
            errors.append("“模块化功能”存在重复功能名")
        feature_set = set(features)
        flows: set[str] = set()
        flow_outputs: dict[str, bool] = {}
        seen_steps: set[tuple[str, str]] = set()
        # 先收集每个流程里功能的最早顺序，供“处理对象”校验引用是否靠前。
        flow_feature_min_order: dict[str, dict[str, int]] = {}
        for row in workbook[FLOW_SHEET].iter_rows(min_row=2, values_only=True):
            f_flow, f_order, f_feature = (_text(row[i]) if i < len(row) else "" for i in range(3))
            if f_flow and f_feature:
                flow_feature_min_order.setdefault(f_flow, {}).setdefault(f_feature, int(float(f_order or 0)))
        for row in workbook[FLOW_SHEET].iter_rows(min_row=2, values_only=True):
            flow, order, feature, enabled = (_text(row[i]) if i < len(row) else "" for i in range(4))
            if not flow or not feature:
                continue
            flows.add(flow)
            on_failure = _text(row[4]) if len(row) > 4 else ""
            if on_failure not in {"停止", "跳过"}:
                errors.append(f"流程“{flow}”的“失败后处理”只能填写“停止”或“跳过”")
            output_name = _text(row[6]) if len(row) > 6 else ""
            if output_name and any(ch in output_name for ch in '\\/:*?"<>|'):
                errors.append(f"流程“{flow}”顺序“{order}”的“输出文件名”包含非法字符（不能包含 \\ / : * ? \" < > |）：{output_name}")
            if _limited(_text(row[3])) and _limited(_text(row[5]) if len(row) > 5 else ""):
                flow_outputs[flow] = True
            if feature not in feature_set:
                errors.append(f"流程“{flow}”引用不存在的模块“{feature}”")
            key = (flow, order)
            if key in seen_steps:
                errors.append(f"流程“{flow}”存在重复顺序“{order}”")
            seen_steps.add(key)
            process_source = _text(row[7]) if len(row) > 7 else ""
            if process_source:
                ref_order = flow_feature_min_order.get(flow, {}).get(process_source)
                if ref_order is None:
                    errors.append(f"流程“{flow}”顺序“{order}”的“处理对象”引用了流程中不存在的功能“{process_source}”")
                elif ref_order >= int(float(order or 0)):
                    errors.append(f"流程“{flow}”顺序“{order}”的“处理对象”引用的功能“{process_source}”必须排在前面")
        for flow in flows:
            if not flow_outputs.get(flow, False):
                errors.append(f"流程“{flow}”没有启用且“是否输出结果=是”的模块")
        for row in workbook[CUSTOM_FEATURE_SHEET].iter_rows(min_row=2, values_only=True):
            button, flow = _text(row[0]), _text(row[1])
            if button and flow and flow not in flows:
                errors.append(f"按钮“{button}”引用不存在的流程“{flow}”")
    finally:
        workbook.close()
    return errors




def matches_named_range(mapping: FeatureMapping, candidate: str) -> bool:
    """Names match exactly or as an underscore-suffixed variant."""
    simple = candidate.rsplit("!", 1)[-1].strip().strip("'")
    return any(
        simple.casefold() == name.casefold()
        or simple.casefold().startswith(name.casefold() + "_")
        for name in mapping.range_names
    )
