from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path


MAPPING_SHEET = "模块化功能"
EXTERNAL_FILE_FUNCTION = "修改_外部文件添加"
FORMULA_COPY_FUNCTION = "修改_公式校验复制"
ISSUE_EXTRACT_FUNCTION = "汇总_公式校验结果提取"
CONDITIONAL_FORMAT_EXTRACT_FUNCTION = "汇总_条件格式结果提取"
AUDIT_RESULT_OUTPUT_FUNCTION = "汇总_审核结果输出"
STRUCTURE_COMPARE_FUNCTION = "核对_表结构比对"
# 旧版“检查_表结构检查”类型名，初始化配置时自动迁移为新类型名。
LEGACY_STRUCTURE_COMPARE_FUNCTION = "检查_表结构检查"
USED_RANGE_SUMMARY_FUNCTION = "汇总_任意行汇总"
FIXED_ROW_SUMMARY_FUNCTION = "汇总_固定行汇总"
WORKBOOK_TABLE_MERGE_FUNCTION = "汇总_汇总表合并"
NAMED_RANGE_CHECK_FUNCTION = "检查_命名区域存在性"
MERGE_ORG_FILES_FUNCTION = "修改_组合联合核查表"
COMBINE_SHEETS_FUNCTION = "修改_组合工作表"
MERGE_ORG_FEATURE = "组合联合核查表"
MERGE_ORG_FLOW = "组合联合核查表"
COMBINE_SHEETS_FEATURE = "组合工作表"
COMBINE_SHEETS_FLOW = "组合工作表"
SOURCE_DIRECTORY_INPUT = "源数据目录"
_OLD_MERGE_ORG_FUNCTION = "修改_合并同机构多表"
_OLD_MERGE_ORG_LABEL = "合并同机构多表"

FEATURE_TYPES = (
    EXTERNAL_FILE_FUNCTION,
    FORMULA_COPY_FUNCTION,
    ISSUE_EXTRACT_FUNCTION,
    CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
    AUDIT_RESULT_OUTPUT_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    USED_RANGE_SUMMARY_FUNCTION,
    FIXED_ROW_SUMMARY_FUNCTION,
    WORKBOOK_TABLE_MERGE_FUNCTION,
    NAMED_RANGE_CHECK_FUNCTION,
    MERGE_ORG_FILES_FUNCTION,
    COMBINE_SHEETS_FUNCTION,
)
# These modules operate on workbooks or result sets directly, never on a
# template named range.  Keeping this explicit also lets the UI prevent
# misleading configuration input.
RANGE_FREE_FEATURE_TYPES = frozenset({
    EXTERNAL_FILE_FUNCTION,
    AUDIT_RESULT_OUTPUT_FUNCTION,
    WORKBOOK_TABLE_MERGE_FUNCTION,
    MERGE_ORG_FILES_FUNCTION,
    COMBINE_SHEETS_FUNCTION,
})
FLOW_SHEET = "执行流程"
FLOW_HEADERS = (
    "流程名", "顺序", "功能名", "启用", "失败后处理", "是否输出结果", "输出文件名", "输入", "输出", "备注",
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
# 上一版 9 列配置（“处理对象”已是正式名称，但尚无“结果集”列）。
FLOW_CURRENT_9_HEADERS = (
    "流程名", "顺序", "功能名", "启用", "失败后处理", "是否输出结果", "输出文件名", "处理对象", "备注",
)
CUSTOM_FEATURE_SHEET = "自定义按钮"
CUSTOM_FEATURE_HEADERS = ("按钮名称", "流程名称", "是否显示", "备注")
LEGACY_CUSTOM_FEATURE_HEADERS = ("功能标识", "功能名称", "是否显示", "备注")
MAPPING_HEADERS = ("功能名", "执行模块", "命名区域名", "输入", "输出", "备注")
DEFAULT_FEATURES = (
    ("检查校验区域", NAMED_RANGE_CHECK_FUNCTION, "校验区域", "模板命名区域", "运行日志", "检查模板是否存在“校验区域”命名区域；可填写多个区域名，均必须存在"),
    ("检查表结构区域", NAMED_RANGE_CHECK_FUNCTION, "表结构区域", "模板命名区域", "运行日志", "检查模板是否存在“表结构区域”命名区域"),
    ("检查汇总区域", NAMED_RANGE_CHECK_FUNCTION, "任意行汇总区域、表头区域", "模板命名区域", "运行日志", "检查汇总模板是否同时存在汇总区域和表头区域"),
    ("外部文件添加", EXTERNAL_FILE_FUNCTION, "", "外部辅助文件、源文件（或上一修改步骤的审核副本）", "已添加外部表的审核副本", "不使用命名区域；将外部工作表复制至审核副本，必须先于公式复制"),
    ("公式校验复制", FORMULA_COPY_FUNCTION, "校验区域", "模板文件、源文件（或已添加外部表的审核副本）", "已复制公式的审核副本", "审核文件是修改后的副本；复制模板公式及格式后用于计算和人工复核"),
    ("校验结果提取", ISSUE_EXTRACT_FUNCTION, "校验区域", "已计算的审核副本", "审核结果集", "重算后提取错误、核实、提示等结果"),
    ("条件格式结果提取", CONDITIONAL_FORMAT_EXTRACT_FUNCTION, "条件格式区域", "源数据目录的条件格式区域", "审核结果集", "提取命名区域内被条件格式实际改变填充色的单元格；单元格批注作为描述，表结构区域用于定位校验指标"),
    ("审核结果输出", AUDIT_RESULT_OUTPUT_FUNCTION, "", "审核结果集", "最终审核结果工作簿", "将同一结果集中的审核问题写入最终工作簿；相同表头追加到同一工作表"),
    ("表结构比对", STRUCTURE_COMPARE_FUNCTION, "表结构区域", "源数据目录、模板表结构区域", "运行日志", "逐格精确核对固定表结构"),
    ("任意行汇总", USED_RANGE_SUMMARY_FUNCTION, "任意行汇总区域", "源数据目录、表头区域", "汇总工作簿", "模板必须定义同工作表的“表头区域”；“单元格区域.字段名”会附加到每行"),
    ("固定行汇总", FIXED_ROW_SUMMARY_FUNCTION, "固定行汇总区域", "源数据目录、表头区域", "汇总工作簿", "模板必须定义同工作表的“表头区域”；严格按框选的固定行汇总"),
    ("汇总表合并", WORKBOOK_TABLE_MERGE_FUNCTION, "", "源数据目录", "汇总工作簿", "不使用命名区域；读取所有源工作簿的工作表，表头相同才合并"),
    (MERGE_ORG_FEATURE, MERGE_ORG_FILES_FUNCTION, "", "源数据目录", "联合核查表", "把源目录下同一机构（文件名主名下划线第一段相同）的所有工作簿组合为联合核查表"),
    (COMBINE_SHEETS_FEATURE, COMBINE_SHEETS_FUNCTION, "", "源数据目录", "组合工作簿", "按“组合工作表分组方案”将任意文件组合为工作簿；支持正则表达式或按名称/关键字分组"),
)
DEFAULT_FLOWS = (
    ("汇总核查表校验", 10, "检查校验区域", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "运行日志", "先确认公式复制和问题提取所需区域存在"),
    ("汇总核查表校验", 20, "检查表结构区域", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "运行日志", "先确认表结构比对所需区域存在"),
    ("汇总核查表校验", 30, "表结构比对", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "运行日志", "不匹配文件跳过，正常运行只记入运行日志"),
    ("汇总核查表校验", 40, "外部文件添加", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "外部文件副本", "外部表先于公式复制；需要排错时再输出阶段副本"),
    ("汇总核查表校验", 50, "公式校验复制", "是", "跳过", "是", "机构审核副本", "外部文件添加", "机构审核副本", "复制模板公式并强制重算，并保留供人工复核的审核副本"),
    ("汇总核查表校验", 60, "校验结果提取", "是", "跳过", "否", "", "公式校验复制", "本期审核结果", "提取公式校验问题，追加到本期审核结果"),
    ("汇总核查表校验", 70, "条件格式结果提取", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "本期审核结果", "提取红色条件格式触发项，追加到本期审核结果"),
    ("汇总核查表校验", 80, "审核结果输出", "是", "跳过", "是", "汇总核查表校验", "本期审核结果", "本期审核结果", "读取本期审核结果并输出正式审核结果工作簿"),
    ("汇总校验结果说明", 10, "检查表结构区域", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "运行日志", "先确认表结构比对所需区域存在"),
    ("汇总校验结果说明", 20, "检查汇总区域", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "运行日志", "先确认汇总区域和表头区域存在"),
    ("汇总校验结果说明", 30, "表结构比对", "是", "跳过", "否", "", SOURCE_DIRECTORY_INPUT, "运行日志", "不匹配文件跳过，正常运行只记入运行日志"),
    ("汇总校验结果说明", 40, "任意行汇总", "是", "跳过", "是", "汇总校验结果说明", SOURCE_DIRECTORY_INPUT, "校验结果与报送说明汇总", "输出正式校验结果与报送说明汇总"),
    (MERGE_ORG_FLOW, 10, MERGE_ORG_FEATURE, "是", "跳过", "是", MERGE_ORG_FLOW, SOURCE_DIRECTORY_INPUT, "联合核查表", "递归组合同一机构的多个报送工作簿，输出每机构一个联合核查表"),
    (COMBINE_SHEETS_FLOW, 10, COMBINE_SHEETS_FEATURE, "是", "跳过", "是", COMBINE_SHEETS_FLOW, SOURCE_DIRECTORY_INPUT, "组合工作表", "按可配置分组方案组合工作表，便于与旧组合联合核查表并行测试"),
)
# 旧配置迁移时，为“默认最终输出步骤”回填流程名（键 = 流程名, 功能名）。
DEFAULT_OUTPUT_NAMES = {(row[0], row[2]): row[6] for row in DEFAULT_FLOWS if row[6]}
# 旧配置迁移时，为“默认功能”回填处理对象（键 = 流程名, 功能名）。
DEFAULT_PROCESS_SOURCES = {(row[0], row[2]): row[7] for row in DEFAULT_FLOWS if row[7]}


# 配置表头批注：告诉使用者每列怎么填、如何影响功能输出。
GUIDE_COMMENTS: dict[str, dict[int, str]] = {
    MAPPING_SHEET: {
        1: "本模块显示名，必须唯一；执行流程的“功能名”列用此名引用。",
        2: "执行模块决定功能行为。可选：\n"
           "检查_命名区域存在性：检查模板是否存在指定命名区域，运行日志会记录每个区域所在工作表与覆盖区域；\n"
           "核对_表结构比对：逐格核对固定表头，不匹配的报送文件跳过；\n"
           "修改_外部文件添加：把外部工作表复制进审核副本，须先于公式校验复制；\n"
           "修改_公式校验复制：复制模板公式并强制重算，可输出同名审核副本目录；\n"
           "修改_组合联合核查表：按文件名第一段识别机构，将同一机构多个工作簿组合为一份联合核查表；\n"
           "汇总_公式校验结果提取：重算后提取错误、核实、提示等结果；\n"
           "汇总_任意行汇总 / 汇总_固定行汇总 / 汇总_汇总表合并：区域汇总。",
        3: "仅填写需要读取模板区域的模块，可填多个并用、分隔；外部文件添加、审核结果输出、汇总表合并、组合联合核查表、组合工作表不使用命名区域，此项会在设置界面禁用。",
        4: "该功能通常读取的内容说明，例如模板命名区域、源数据目录、审核副本或结果集；用于帮助理解，不参与流程调度。",
        5: "该功能通常产生的内容说明，例如运行日志、审核副本、结果集或汇总工作簿；用于帮助理解，不参与流程调度。",
        6: "自由说明。",
    },
    FLOW_SHEET: {
        1: "流程名即界面主按钮文字，应唯一；自定义按钮的“流程名称”也引用此名。",
        2: "步骤执行顺序，建议 10、20、30… 递增。",
        3: "引用“模块化功能”页中的功能名。",
        4: "填“是”才执行该步骤。",
        5: "仅支持“停止”或“跳过”：停止=该步骤出错则中止整个流程；跳过=出错则跳过当前文件继续。",
        6: "填“是”表示保留该步骤的实际输出。检查/核对类填“是”时，其检查明细保存在流程运行日志中，适合单独作为自定义流程执行；修改、汇总类填“是”时保留相应副本或最终工作簿。",
        7: "输出文件名（可选）：仅对“是否输出结果=是”的步骤生效；填写后，该步骤的输出文件/目录名从“序号_功能名_时间戳”改为“输出文件名_时间戳”。最终输出建议填流程名。",
        8: "输入：选择“源数据目录”时读取原始报送文件；也可填写流程中排在前面的外部文件添加或公式校验复制功能名，读取其产出的副本。条件格式结果提取默认必须选择“源数据目录”。“审核结果输出”填写结果集名称，例如“本期审核结果”。",
        9: "输出（可选）：问题提取模块把记录写入指定结果集；“审核结果输出”读取同名结果集并生成最终文件。公式校验结果提取和条件格式结果提取填相同名称，例如“本期审核结果”，将追加到同一工作表。",
        10: "说明。",
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
    output: str = ""
    input_note: str = ""


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
    result_set: str = ""


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


def _initialize_legacy_config(config_path: Path) -> Path:
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
            # 仅迁移精确旧名，不保留别名，避免模块目录中同时出现两套同义功能。
            if _text(name_cell.value) == _OLD_MERGE_ORG_LABEL:
                name_cell.value = MERGE_ORG_FEATURE
            if _text(type_cell.value) == _OLD_MERGE_ORG_FUNCTION:
                type_cell.value = MERGE_ORG_FILES_FUNCTION
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
                if actual_headers in (FLOW_LEGACY_HEADERS, FLOW_PREVIOUS_HEADERS, FLOW_PREVIOUS_9_HEADERS, FLOW_CURRENT_9_HEADERS):
                    # 旧配置增量迁移，保留用户既有流程行：
                    # 7 列先插入“输出文件名”，再插入“处理对象”；8 列只插入“处理对象”；
                    # 9 列只是把“待处理对象”列改名“处理对象”。
                    if actual_headers == FLOW_LEGACY_HEADERS:
                        flow_sheet.insert_cols(7)
                        flow_sheet.cell(1, 7, "输出文件名")
                    if actual_headers not in (FLOW_PREVIOUS_9_HEADERS, FLOW_CURRENT_9_HEADERS):
                        flow_sheet.insert_cols(8)
                    flow_sheet.cell(1, 8, "处理对象")
                    # 旧第 9 列是“备注”，在它之前插入“结果集”。
                    flow_sheet.insert_cols(9)
                    flow_sheet.cell(1, 9, "结果集")
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
            if _text(row[0].value) == _OLD_MERGE_ORG_LABEL:
                row[0].value = MERGE_ORG_FLOW
            if _text(row[2].value) == _OLD_MERGE_ORG_LABEL:
                row[2].value = MERGE_ORG_FEATURE
            if _text(row[6].value) == _OLD_MERGE_ORG_LABEL:
                row[6].value = MERGE_ORG_FLOW
            if _text(row[7].value) == _OLD_MERGE_ORG_LABEL:
                row[7].value = MERGE_ORG_FEATURE
        # 旧流程在迁移前可能因“流程名 + 功能名”不同而被补过一条新默认行；
        # 合并流程本身只能有一个模块，保留原先的首条配置并删去重复默认行。
        merge_rows = [
            row_number
            for row_number in range(2, flow_sheet.max_row + 1)
            if _text(flow_sheet.cell(row_number, 1).value) == MERGE_ORG_FLOW
            and _text(flow_sheet.cell(row_number, 3).value) == MERGE_ORG_FEATURE
        ]
        for row_number in reversed(merge_rows[1:]):
            flow_sheet.delete_rows(row_number, 1)
        for cell in flow_sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for column, width in {"A": 20, "B": 10, "C": 20, "D": 10, "E": 16, "F": 16, "G": 18, "H": 18, "I": 18, "J": 38}.items():
            flow_sheet.column_dimensions[column].width = width
        flow_sheet.freeze_panes = "A2"
        flow_sheet.auto_filter.ref = f"A1:J{max(2, flow_sheet.max_row)}"

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
        for row in custom_sheet.iter_rows(min_row=2):
            if _text(row[0].value) == _OLD_MERGE_ORG_LABEL:
                row[0].value = MERGE_ORG_FLOW
            if _text(row[1].value) == _OLD_MERGE_ORG_LABEL:
                row[1].value = MERGE_ORG_FLOW
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


def _reset_legacy_default_configuration(config_path: Path) -> Path:
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
            FeatureMapping(row[0], row[1], _split_names(row[2]), False, "", row[5], row[4], row[3])
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

            feature_type = get("执行模块")
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
        process_source_col = index.get("输入")
        selected: list[FlowStep] = []
        for row in rows:
            if _text(row[index["流程名"]]) != flow_name or not _limited(_text(row[index["启用"]])):
                continue
            feature_name = _text(row[index["功能名"]])
            if feature_name:
                on_failure = _text(row[index["失败后处理"]]) or "跳过"
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
                if actual not in (FLOW_HEADERS, FLOW_LEGACY_HEADERS, FLOW_PREVIOUS_HEADERS, FLOW_PREVIOUS_9_HEADERS, FLOW_CURRENT_9_HEADERS):
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
            if process_source and process_source != SOURCE_DIRECTORY_INPUT:
                ref_order = flow_feature_min_order.get(flow, {}).get(process_source)
                if ref_order is None:
                    errors.append(f"流程“{flow}”顺序“{order}”的“输入”引用了流程中不存在的功能“{process_source}”")
                elif ref_order >= int(float(order or 0)):
                    errors.append(f"流程“{flow}”顺序“{order}”的“输入”引用的功能“{process_source}”必须排在前面")
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


# 设置中心的数据接口 ---------------------------------------------------------
#
# config.xlsx 仍是唯一的配置来源。下面的函数只把其中三个“程序配置”页转换为
# JSON 友好的行字典，供 WebView 做可视化编辑；历史审核结果工作表不会被读取、
# 重写或传给前端。

DEFAULT_FLOW_NAMES = tuple(dict.fromkeys(row[0] for row in DEFAULT_FLOWS))


def _row_dict(headers: tuple[str, ...], row: tuple[object, ...]) -> dict[str, str]:
    return {
        header: _text(row[index]) if index < len(row) else ""
        for index, header in enumerate(headers)
    }


def load_config_editor_data(config_path: Path) -> dict[str, object]:
    """Return editable configuration data without exposing history records."""
    from openpyxl import load_workbook

    if not config_path.is_file():
        initialize_config(config_path)
    workbook = load_workbook(config_path, read_only=True, data_only=True)
    try:
        def read_rows(sheet_name: str, headers: tuple[str, ...]) -> list[dict[str, str]]:
            if sheet_name not in workbook.sheetnames:
                return []
            sheet = workbook[sheet_name]
            actual = tuple(_text(cell.value) for cell in sheet[1])
            if actual != headers:
                raise ValueError(f"config.xlsx 的“{sheet_name}”表头不符合规范，请先执行重置配置")
            return [
                _row_dict(headers, row)
                for row in sheet.iter_rows(min_row=2, values_only=True)
                if any(_text(value) for value in row)
            ]

        mappings = read_rows(MAPPING_SHEET, MAPPING_HEADERS)
        all_flows = read_rows(FLOW_SHEET, FLOW_HEADERS)
        buttons = read_rows(CUSTOM_FEATURE_SHEET, CUSTOM_FEATURE_HEADERS)
        default_rows = [row for row in all_flows if row["流程名"] in DEFAULT_FLOW_NAMES]
        custom_rows = [row for row in all_flows if row["流程名"] not in DEFAULT_FLOW_NAMES]
        return {
            "featureTypes": list(FEATURE_TYPES),
            "mappingHeaders": list(MAPPING_HEADERS),
            "flowHeaders": list(FLOW_HEADERS),
            "buttonHeaders": list(CUSTOM_FEATURE_HEADERS),
            "modules": mappings,
            "defaultFlows": default_rows,
            "customFlows": custom_rows,
            "buttons": buttons,
            "defaultFlowNames": list(DEFAULT_FLOW_NAMES),
        }
    finally:
        workbook.close()


def _editor_rows(
    rows: object, headers: tuple[str, ...], label: str
) -> list[dict[str, str]]:
    if not isinstance(rows, list):
        raise ValueError(f"设置中心的“{label}”必须是表格行列表")
    result: list[dict[str, str]] = []
    for number, raw in enumerate(rows, start=1):
        if not isinstance(raw, dict):
            raise ValueError(f"设置中心的“{label}”第 {number} 行不是有效记录")
        row = {header: _text(raw.get(header)) for header in headers}
        if any(row.values()):
            result.append(row)
    return result


def validate_config_editor_draft(
    config_path: Path, draft: object
) -> list[str]:
    """Validate a visual-editor draft before it changes config.xlsx."""
    if not isinstance(draft, dict):
        return ["设置中心提交的数据格式无效"]
    try:
        modules = _editor_rows(draft.get("modules", []), MAPPING_HEADERS, "模块化功能")
        custom_flows = _editor_rows(draft.get("customFlows", []), FLOW_HEADERS, "自定义流程")
        buttons = _editor_rows(draft.get("buttons", []), CUSTOM_FEATURE_HEADERS, "自定义按钮")
    except ValueError as exc:
        return [str(exc)]

    errors: list[str] = []
    names = [row["功能名"] for row in modules if row["功能名"]]
    if len(names) != len(set(names)):
        errors.append("模块化功能存在重复功能名")
    for row in modules:
        if not row["功能名"]:
            errors.append("模块化功能存在未填写功能名的行")
        if row["执行模块"] not in FEATURE_TYPES:
            errors.append(f"模块“{row['功能名'] or '未命名'}”的执行模块无效")
        if row["是否限定工作簿"] not in {"", "是", "否"}:
            errors.append(f"模块“{row['功能名'] or '未命名'}”的“是否限定工作簿”只能填是或否")

    all_feature_names = set(names)
    seen_order: set[tuple[str, int]] = set()
    enabled_outputs: dict[str, bool] = {}
    flow_names: set[str] = set()
    for row in custom_flows:
        flow, feature = row["流程名"], row["功能名"]
        if not flow or not feature:
            errors.append("自定义流程的流程名和功能名均不能为空")
            continue
        if flow in DEFAULT_FLOW_NAMES:
            errors.append(f"默认流程“{flow}”只能查看，不能在自定义流程中修改")
        flow_names.add(flow)
        try:
            order = int(float(row["顺序"]))
        except (TypeError, ValueError):
            errors.append(f"流程“{flow}”的顺序必须为数字")
            continue
        if (flow, order) in seen_order:
            errors.append(f"流程“{flow}”存在重复顺序“{order}”")
        seen_order.add((flow, order))
        if feature not in all_feature_names:
            errors.append(f"流程“{flow}”引用不存在的模块“{feature}”")
        if row["启用"] not in {"", "是", "否"}:
            errors.append(f"流程“{flow}”的“启用”只能填是或否")
        if row["失败后处理"] not in {"停止", "跳过"}:
            errors.append(f"流程“{flow}”的“失败后处理”只能填停止或跳过")
        if row["是否输出结果"] not in {"", "是", "否"}:
            errors.append(f"流程“{flow}”的“是否输出结果”只能填是或否")
        if row["启用"] == "是" and row["是否输出结果"] == "是":
            enabled_outputs[flow] = True
        if any(char in row["输出文件名"] for char in '\\/:*?"<>|'):
            errors.append(f"流程“{flow}”的输出文件名包含非法字符")
    for flow in flow_names:
        if not enabled_outputs.get(flow):
            errors.append(f"自定义流程“{flow}”至少要有一个启用步骤输出结果")

    available_flows = set(DEFAULT_FLOW_NAMES) | flow_names
    button_names = [row["按钮名称"] for row in buttons if row["按钮名称"]]
    if len(button_names) != len(set(button_names)):
        errors.append("自定义按钮存在重复按钮名称")
    for row in buttons:
        if not row["按钮名称"] or not row["流程名称"]:
            errors.append("自定义按钮的按钮名称和流程名称均不能为空")
        elif row["流程名称"] not in available_flows:
            errors.append(f"按钮“{row['按钮名称']}”引用不存在的流程“{row['流程名称']}”")
        if row["是否显示"] not in {"", "是", "否"}:
            errors.append(f"按钮“{row['按钮名称'] or '未命名'}”的“是否显示”只能填是或否")
    return errors


def save_config_editor_draft(config_path: Path, draft: object) -> Path:
    """Save modules/custom flows/buttons while preserving defaults and history."""
    errors = validate_config_editor_draft(config_path, draft)
    if errors:
        raise ValueError("设置中心配置不合法：" + "；".join(errors))
    modules = _editor_rows(draft.get("modules", []), MAPPING_HEADERS, "模块化功能")
    custom_flows = _editor_rows(draft.get("customFlows", []), FLOW_HEADERS, "自定义流程")
    buttons = _editor_rows(draft.get("buttons", []), CUSTOM_FEATURE_HEADERS, "自定义按钮")
    initialize_config(config_path)
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    workbook = load_workbook(config_path)
    try:
        # Default flows are not accepted from the client.  Preserve current
        # default rows so a user who still uses Excel to inspect them never
        # loses them when saving a visual-editor draft.
        default_rows = [
            _row_dict(FLOW_HEADERS, row)
            for row in workbook[FLOW_SHEET].iter_rows(min_row=2, values_only=True)
            if _text(row[0]) in DEFAULT_FLOW_NAMES
        ]

        def rewrite(sheet_name: str, headers: tuple[str, ...], rows: list[dict[str, str]], widths: dict[str, int]) -> None:
            sheet = workbook[sheet_name]
            sheet.delete_rows(1, sheet.max_row)
            sheet.append(headers)
            for row in rows:
                sheet.append([row[header] for header in headers])
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center")
            for column, width in widths.items():
                sheet.column_dimensions[column].width = width
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(2, sheet.max_row)}"

        rewrite(MAPPING_SHEET, MAPPING_HEADERS, modules, {"A": 18, "B": 22, "C": 24, "D": 18, "E": 28, "F": 42})
        rewrite(FLOW_SHEET, FLOW_HEADERS, [*default_rows, *custom_flows], {"A": 20, "B": 10, "C": 20, "D": 10, "E": 16, "F": 16, "G": 18, "H": 18, "I": 38})
        rewrite(CUSTOM_FEATURE_SHEET, CUSTOM_FEATURE_HEADERS, buttons, {"A": 24, "B": 24, "C": 12, "D": 42})
        _apply_guide_comments(workbook)
        workbook.save(config_path)
    except PermissionError as exc:
        raise RuntimeError(f"无法保存设置：请先关闭已打开的配置文件“{config_path.name}”后重试") from exc
    finally:
        workbook.close()
    return config_path




def matches_named_range(mapping: FeatureMapping, candidate: str) -> bool:
    """Names match exactly or as an underscore-suffixed variant."""
    simple = candidate.rsplit("!", 1)[-1].strip().strip("'")
    return any(
        simple.casefold() == name.casefold()
        or simple.casefold().startswith(name.casefold() + "_")
        for name in mapping.range_names
    )


# JSON 工作流配置 ------------------------------------------------------------
#
# 审核历史必须让业务人员在 Excel 中直接维护；模块和流程则更适合由界面维护，
# 因而从 config.xlsx 迁出到同目录的“流程配置.json”。下面保留旧函数名，
# 让 service/excel_com 不必知道存储介质已改变。

WORKFLOW_CONFIG_NAME = "流程配置.json"
HISTORY_WORKBOOK_NAME = "历史审核配置.xlsx"
LEGACY_HISTORY_WORKBOOK_NAME = "历史审核说明.xlsx"
WORKFLOW_CONFIG_VERSION = 1
DEFAULT_COMBINE_SHEETS_PLAN = {
    "id": "default",
    "name": "按文件名正则提取组合",
    "mode": "regex",
    "pattern": "(?P<组合>.+?)_金融基础数据-",
    "groups": [],
}


def _default_combine_sheets_plans() -> list[dict[str, object]]:
    """Return fresh defaults; profiles must never share mutable lists/maps."""
    return [{
        **DEFAULT_COMBINE_SHEETS_PLAN,
        "groups": [],
    }]


def workflow_config_path(history_config_path: Path) -> Path:
    history_path = Path(history_config_path)
    # The history workbook intentionally lives beside the EXE.  All
    # application-managed JSON remains in its hidden data directory.
    if history_path.name == HISTORY_WORKBOOK_NAME:
        return history_path.parent / "data" / WORKFLOW_CONFIG_NAME
    return history_path.parent / WORKFLOW_CONFIG_NAME


def _default_module_rows() -> list[dict[str, str]]:
    return [_row_dict(MAPPING_HEADERS, row) for row in DEFAULT_FEATURES]


def _default_flow_rows() -> list[dict[str, str]]:
    return [_row_dict(FLOW_HEADERS, row) for row in DEFAULT_FLOWS]


def _write_workflow(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def _read_workflow(path: Path) -> dict[str, object]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError) as exc:
        raise RuntimeError(f"无法读取流程配置“{path.name}”：{exc}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError(f"流程配置“{path.name}”格式无效")
    return payload


def _legacy_workflow_rows(history_path: Path) -> tuple[list[dict[str, str]], list[dict[str, str]], list[dict[str, str]]]:
    """Read the three former configuration sheets once, without history."""
    from openpyxl import load_workbook
    if not history_path.is_file():
        return [], [], []
    workbook = load_workbook(history_path, read_only=True, data_only=True)
    try:
        def read(sheet_name: str, headers: tuple[str, ...]) -> list[dict[str, str]]:
            if sheet_name not in workbook.sheetnames:
                return []
            sheet = workbook[sheet_name]
            actual = tuple(_text(cell.value) for cell in sheet[1])
            if actual != headers:
                return []
            return [
                _row_dict(headers, row) for row in sheet.iter_rows(min_row=2, values_only=True)
                if any(_text(value) for value in row)
            ]
        return (
            read(MAPPING_SHEET, MAPPING_HEADERS),
            read(FLOW_SHEET, FLOW_HEADERS),
            read(CUSTOM_FEATURE_SHEET, CUSTOM_FEATURE_HEADERS),
        )
    finally:
        workbook.close()


def _ensure_history_workbook(history_path: Path, *, remove_legacy_sheets: bool) -> None:
    """Ensure the history workbook contains the two history sheets; never rewrite manual columns."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
    from .history import (
        HISTORY_HEADERS,
        HISTORY_AUDIT_SHEET,
        LOCAL_VALIDATION_HISTORY_SHEET,
    )

    history_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = load_workbook(history_path) if history_path.is_file() else Workbook()
    try:
        # 旧 sheet 名（历史审核结果 / 历史审核记录）统一迁移为「历史核查表审核」。
        if HISTORY_AUDIT_SHEET not in workbook.sheetnames:
            for legacy_name in ("历史审核结果", "历史审核记录"):
                if legacy_name in workbook.sheetnames:
                    workbook[legacy_name].title = HISTORY_AUDIT_SHEET
                    break
            else:
                sheet = workbook.active if workbook.active.title == "Sheet" and len(workbook.sheetnames) == 1 else workbook.create_sheet(HISTORY_AUDIT_SHEET)
                sheet.title = HISTORY_AUDIT_SHEET
        sheet = workbook[HISTORY_AUDIT_SHEET]
        if not any(cell.value for cell in sheet[1]):
            for column, header in enumerate(HISTORY_HEADERS, start=1):
                sheet.cell(1, column, header)
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:L{max(2, sheet.max_row)}"

        # 本地校验结果历史：只建空表，表头由汇总流程首次追加时写入。
        if LOCAL_VALIDATION_HISTORY_SHEET not in workbook.sheetnames:
            workbook.create_sheet(LOCAL_VALIDATION_HISTORY_SHEET)

        if remove_legacy_sheets:
            for name in (MAPPING_SHEET, FLOW_SHEET, CUSTOM_FEATURE_SHEET, "区域映射", "自定义功能", "外部工作表", "历史审核记录", "历史审核结果"):
                if name in workbook.sheetnames and name not in (HISTORY_AUDIT_SHEET, LOCAL_VALIDATION_HISTORY_SHEET):
                    del workbook[name]
            # A brand-new workbook can retain an unrelated blank Sheet.
            if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
                del workbook["Sheet"]
        workbook.save(history_path)
    except PermissionError as exc:
        raise RuntimeError(f"无法更新历史审核结果：请先关闭“{history_path.name}”后重试") from exc
    finally:
        workbook.close()


def _migrate_legacy_workflow(history_path: Path) -> dict[str, object]:
    """Create JSON from an old Excel configuration, then retain only history."""
    mappings, flows, buttons = _legacy_workflow_rows(history_path)
    default_names = set(DEFAULT_FLOW_NAMES)
    default_module_names = {row["功能名"] for row in _default_module_rows()}
    payload = {
        "version": WORKFLOW_CONFIG_VERSION,
        # Custom modules are retained for compatibility with old flows.  The
        # visual editor deliberately does not expose editing them; new users
        # compose read-only built-in modules into custom flows instead.
        "customModules": [row for row in mappings if row["功能名"] not in default_module_names],
        "customFlows": [row for row in flows if row["流程名"] not in default_names],
        "flowDisplay": {
            row["流程名称"]: True
            for row in buttons
            if row["流程名称"] and row["是否显示"] == "是"
        },
        "combineSheetsPlans": _default_combine_sheets_plans(),
        "activeCombineSheetsPlanId": "default",
    }
    _write_workflow(workflow_config_path(history_path), payload)
    _ensure_history_workbook(history_path, remove_legacy_sheets=True)
    return payload


def _workflow(history_path: Path) -> dict[str, object]:
    path = workflow_config_path(history_path)
    if not path.is_file():
        return _migrate_legacy_workflow(history_path)
    payload = _read_workflow(path)
    if int(payload.get("version") or 0) != WORKFLOW_CONFIG_VERSION:
        raise RuntimeError(f"流程配置“{path.name}”版本不支持，请使用重置流程配置恢复默认")
    if "combineSheetsPlans" not in payload:
        # 单方案版本自动成为首个可切换方案，不丢弃用户已填写的规则。
        legacy_plan = payload.get("combineSheetsPlan")
        if isinstance(legacy_plan, dict):
            payload["combineSheetsPlans"] = [{
                **DEFAULT_COMBINE_SHEETS_PLAN,
                **legacy_plan,
                "id": _text(legacy_plan.get("id")) or "default",
            }]
        else:
            payload["combineSheetsPlans"] = _default_combine_sheets_plans()
        payload["activeCombineSheetsPlanId"] = _text(
            payload.get("activeCombineSheetsPlanId")
        ) or _text(payload["combineSheetsPlans"][0].get("id"))
        payload.pop("combineSheetsPlan", None)
        _write_workflow(path, payload)
    # 自定义按钮已收口为“流程是否显示在主界面”。旧按钮配置首次读取时
    # 自动按其目标流程迁移，按钮文字改用流程名，避免再维护两套名称。
    if "flowDisplay" not in payload or "buttons" in payload:
        display = payload.get("flowDisplay", {})
        if not isinstance(display, dict):
            display = {}
        for row in _workflow_rows(payload, "buttons", CUSTOM_FEATURE_HEADERS):
            flow = row["流程名称"]
            if flow and row["是否显示"] == "是":
                display[flow] = True
        payload["flowDisplay"] = {
            _text(name): bool(value)
            for name, value in display.items()
            if _text(name) and bool(value)
        }
        payload.pop("buttons", None)
        _write_workflow(path, payload)
    return payload


def initialize_config(config_path: Path) -> Path:
    """Migrate program settings and retain historical explanations in Excel."""
    history_path = Path(config_path)
    # Rename the former generic filename once, preserving all manually entered
    # history.  Only the application-owned target name triggers this migration;
    # callers using an arbitrary path (including tests) keep their own filename.
    if history_path.name == HISTORY_WORKBOOK_NAME and not history_path.exists():
        # 旧版曾把历史存成 data/下的 config.xlsx，或根目录的「历史审核说明.xlsx」。
        # 统一迁移到根目录的「历史审核配置.xlsx」，保留所有人工填写内容。
        candidates = (
            history_path.parent / "data" / HISTORY_WORKBOOK_NAME,
            history_path.parent / "data" / LEGACY_HISTORY_WORKBOOK_NAME,
            history_path.parent / LEGACY_HISTORY_WORKBOOK_NAME,
            history_path.parent / "data" / "config.xlsx",
            history_path.with_name("config.xlsx"),
        )
        legacy_path = next((item for item in candidates if item.is_file()), None)
        if legacy_path is not None:
            try:
                legacy_path.replace(history_path)
            except PermissionError as exc:
                raise RuntimeError(
                    f"无法迁移历史审核说明：请先关闭“{legacy_path.name}”后重试"
                ) from exc
    _workflow(history_path)
    _ensure_history_workbook(history_path, remove_legacy_sheets=True)
    return history_path


def reset_default_configuration(config_path: Path) -> Path:
    """Reset visual workflow data, leaving the Excel history workbook untouched."""
    history_path = Path(config_path)
    _ensure_history_workbook(history_path, remove_legacy_sheets=True)
    _write_workflow(workflow_config_path(history_path), {
        "version": WORKFLOW_CONFIG_VERSION,
        "customModules": [],
        "customFlows": [],
        "flowDisplay": {},
        "combineSheetsPlans": _default_combine_sheets_plans(),
        "activeCombineSheetsPlanId": "default",
    })
    return history_path


def _workflow_rows(payload: dict[str, object], key: str, headers: tuple[str, ...]) -> list[dict[str, str]]:
    raw = payload.get(key, [])
    if not isinstance(raw, list):
        return []
    result: list[dict[str, str]] = []
    for item in raw:
        if isinstance(item, dict):
            result.append({header: _text(item.get(header)) for header in headers})
    return result


def load_feature_mappings(config_path: Path | None, template_path: Path) -> list[FeatureMapping]:
    """Load default modules, allowing saved rows to override them by name."""
    rows = _default_module_rows()
    if config_path is not None:
        rows.extend(_workflow_rows(_workflow(Path(config_path)), "customModules", MAPPING_HEADERS))
    selected: dict[str, FeatureMapping] = {}
    for row in rows:
        feature_name, feature_type = row["功能名"], row["执行模块"]
        if not feature_name or feature_type not in FEATURE_TYPES:
            continue
        selected[feature_name] = FeatureMapping(
            feature_name, feature_type,
            () if feature_type in RANGE_FREE_FEATURE_TYPES else _split_names(row["命名区域名"]),
            False, "", row["备注"], row["输出"], row["输入"]
        )
    return list(selected.values())


def load_flow_steps(config_path: Path, flow_name: str) -> tuple[FlowStep, ...]:
    custom_rows = _workflow_rows(_workflow(Path(config_path)), "customFlows", FLOW_HEADERS)
    overridden_flows = {row["流程名"] for row in custom_rows if row["流程名"]}
    rows = [
        *[row for row in _default_flow_rows() if row["流程名"] not in overridden_flows],
        *custom_rows,
    ]
    module_rows = [
        *_default_module_rows(),
        *_workflow_rows(_workflow(Path(config_path)), "customModules", MAPPING_HEADERS),
    ]
    feature_types_by_name = {
        row["功能名"]: row["执行模块"]
        for row in module_rows
        if row["功能名"]
    }
    selected: list[FlowStep] = []
    for row in rows:
        if row["流程名"] != flow_name or not _limited(row["启用"]):
            continue
        on_failure = row["失败后处理"] or "跳过"
        if on_failure not in {"停止", "跳过"}:
            raise ValueError(f"执行流程“{flow_name}”的“失败后处理”只能填写“停止”或“跳过”")
        try:
            order = int(float(row["顺序"] or 0))
        except ValueError as exc:
            raise ValueError(f"执行流程“{flow_name}”的顺序必须是数字") from exc
        output_result = _limited(row["是否输出结果"])
        selected.append(FlowStep(flow_name, order, row["功能名"], on_failure, output_result, row["备注"], row["输出文件名"], row["输入"], row["输出"]))
    return tuple(sorted(selected, key=lambda item: item.order))


def load_flow_features(config_path: Path, flow_name: str) -> tuple[str, ...]:
    return tuple(step.feature_name for step in load_flow_steps(config_path, flow_name))


def validate_config(config_path: Path) -> list[str]:
    """Validate JSON workflow data; config.xlsx is only checked for history presence."""
    try:
        payload = _workflow(Path(config_path))
    except Exception as exc:
        return [str(exc)]
    draft = {
        "customModules": _workflow_rows(payload, "customModules", MAPPING_HEADERS),
        "customFlows": _workflow_rows(payload, "customFlows", FLOW_HEADERS),
        "flowDisplay": payload.get("flowDisplay", {}),
        "combineSheetsPlans": payload.get("combineSheetsPlans", []),
        "activeCombineSheetsPlanId": payload.get("activeCombineSheetsPlanId", ""),
    }
    return validate_config_editor_draft(config_path, draft)


def load_config_editor_data(config_path: Path) -> dict[str, object]:
    payload = _workflow(Path(config_path))
    plans = payload.get("combineSheetsPlans", _default_combine_sheets_plans())
    active_plan_id = _text(payload.get("activeCombineSheetsPlanId"))
    active_plan = next(
        (item for item in plans if isinstance(item, dict) and _text(item.get("id")) == active_plan_id),
        _default_combine_sheets_plans()[0],
    )
    saved_modules = _workflow_rows(payload, "customModules", MAPPING_HEADERS)
    modules_by_name = {row["功能名"]: row for row in _default_module_rows()}
    modules_by_name.update({row["功能名"]: row for row in saved_modules if row["功能名"]})
    modules = list(modules_by_name.values())
    saved_flows = _workflow_rows(payload, "customFlows", FLOW_HEADERS)
    overridden_flow_names = {row["流程名"] for row in saved_flows if row["流程名"]}
    effective_defaults = [
        *[row for row in _default_flow_rows() if row["流程名"] not in overridden_flow_names],
        *[row for row in saved_flows if row["流程名"] in DEFAULT_FLOW_NAMES],
    ]
    effective_custom = [row for row in saved_flows if row["流程名"] not in DEFAULT_FLOW_NAMES]
    return {
        "featureTypes": list(FEATURE_TYPES),
        "mappingHeaders": list(MAPPING_HEADERS),
        "flowHeaders": list(FLOW_HEADERS),
        "modules": modules,
        # 供设置中心的“恢复该页默认”使用。这里必须是未叠加用户覆盖的
        # 初始定义，不能复用上面的 effective modules / flows。
        "initialModules": _default_module_rows(),
        "initialFlows": _default_flow_rows(),
        "defaultFlows": effective_defaults,
        "customFlows": effective_custom,
        "flowDisplay": {
            _text(name): bool(value)
            for name, value in (payload.get("flowDisplay", {}) or {}).items()
            if _text(name) and bool(value)
        },
        "defaultFlowNames": list(DEFAULT_FLOW_NAMES),
        "storage": "流程配置.json",
        # 单方案字段只为旧前端临时兼容；新前端应读取 profiles + 当前 id。
        "combineSheetsPlan": active_plan,
        "combineSheetsPlans": plans,
        "activeCombineSheetsPlanId": active_plan_id,
        "combineSheetsModes": [
            {"id": "regex", "name": "按文件名正则提取"},
            {"id": "name", "name": "按名称分组"},
        ],
    }


def _validate_combine_sheets_plan(plan: object) -> list[str]:
    """Validate the optional settings-centre plan before it is persisted."""
    if not isinstance(plan, dict):
        return ["组合工作表分组方案格式无效"]
    merged = {**DEFAULT_COMBINE_SHEETS_PLAN, **plan}
    mode = _text(merged.get("mode"))
    if mode not in {"regex", "name"}:
        return ["组合工作表分组方式只能为 regex（正则表达式）或 name（按名称分组）"]
    if not _text(merged.get("name")):
        return ["组合工作表分组方案名称不能为空"]
    if mode == "regex":
        pattern = _text(merged.get("pattern"))
        if not pattern:
            return ["按文件名正则提取时必须填写正则表达式"]
        try:
            compiled = re.compile(pattern)
        except re.error as exc:
            return [f"组合工作表正则表达式无效：{exc}"]
        if "组合" not in compiled.groupindex and "group" not in compiled.groupindex and compiled.groups < 1:
            return ["组合工作表正则必须包含命名分组“组合”（或 group），或至少一个括号分组"]
    else:
        groups = merged.get("groups")
        if not isinstance(groups, list) or not groups:
            return ["按名称分组时至少需要一条“组合名称 + 匹配名称/关键字”规则"]
        names: set[str] = set()
        for index, item in enumerate(groups, start=1):
            if not isinstance(item, dict):
                return [f"按名称分组第 {index} 条格式无效"]
            name = _text(item.get("name"))
            raw_keywords = item.get("keywords")
            words = (
                re.split(r"[、，,；;\n]", raw_keywords)
                if isinstance(raw_keywords, str)
                else raw_keywords
            )
            if not name:
                return [f"按名称分组第 {index} 条未填写组合名称"]
            if name in names:
                return [f"按名称分组存在重复组合名称“{name}”"]
            names.add(name)
            if not isinstance(words, list) or not any(_text(word) for word in words):
                return [f"按名称分组“{name}”未填写匹配名称/关键字"]
    return []


def _validate_combine_sheets_profiles(plans: object, active_plan_id: object) -> list[str]:
    if not isinstance(plans, list) or not plans:
        return ["组合工作表至少需要保留一个分组方案"]
    errors: list[str] = []
    ids: set[str] = set()
    names: set[str] = set()
    for index, plan in enumerate(plans, start=1):
        if not isinstance(plan, dict):
            errors.append(f"组合工作表第 {index} 个分组方案格式无效")
            continue
        plan_id, name = _text(plan.get("id")), _text(plan.get("name"))
        if not plan_id:
            errors.append(f"组合工作表第 {index} 个分组方案缺少标识")
        elif plan_id in ids:
            errors.append(f"组合工作表存在重复分组方案标识“{plan_id}”")
        ids.add(plan_id)
        if name and name in names:
            errors.append(f"组合工作表存在重复分组方案名称“{name}”")
        names.add(name)
        errors.extend(_validate_combine_sheets_plan(plan))
    active = _text(active_plan_id)
    if not active:
        errors.append("请指定当前启用的组合工作表分组方案")
    elif active not in ids:
        errors.append("当前启用的组合工作表分组方案不存在")
    return errors


def validate_config_editor_draft(config_path: Path, draft: object) -> list[str]:
    """Validate custom modules, flows and display settings from the visual editor."""
    if not isinstance(draft, dict):
        return ["设置中心提交的数据格式无效"]
    try:
        custom_modules = _editor_rows(draft.get("customModules", []), MAPPING_HEADERS, "自定义模块")
        custom_flows = _editor_rows(draft.get("customFlows", []), FLOW_HEADERS, "自定义流程")
    except ValueError as exc:
        return [str(exc)]
    custom_module_names: set[str] = set()
    errors: list[str] = []
    if "combineSheetsPlans" in draft:
        errors.extend(_validate_combine_sheets_profiles(
            draft["combineSheetsPlans"], draft.get("activeCombineSheetsPlanId")
        ))
    elif "combineSheetsPlan" in draft:
        # 旧界面尚未升级时，继续允许只保存当前方案。
        errors.extend(_validate_combine_sheets_plan(draft["combineSheetsPlan"]))
    for row in custom_modules:
        name = row["功能名"]
        if not name:
            errors.append("自定义模块的功能名不能为空")
        elif name in custom_module_names:
            errors.append(f"模块存在重复功能名“{name}”")
        custom_module_names.add(name)
        if row["执行模块"] not in FEATURE_TYPES:
            errors.append(f"自定义模块“{name or '未命名'}”的执行模块必须从内置模块中选择")
        elif row["执行模块"] in RANGE_FREE_FEATURE_TYPES and row["命名区域名"]:
            errors.append(f"功能“{name or '未命名'}”使用的模块不需要命名区域，请清空“命名区域名”")
    all_module_rows = [*_default_module_rows(), *custom_modules]
    feature_types_by_name = {row["功能名"]: row["执行模块"] for row in all_module_rows}
    feature_names = set(feature_types_by_name)
    flow_names: set[str] = set()
    used_orders: set[tuple[str, int]] = set()
    flow_outputs: dict[str, bool] = {}
    for row in custom_flows:
        flow, feature = row["流程名"], row["功能名"]
        if not flow or not feature:
            errors.append("自定义流程的流程名和功能名均不能为空")
            continue
        flow_names.add(flow)
        if feature not in feature_names:
            errors.append(f"流程“{flow}”引用不存在的内置模块“{feature}”")
        try:
            order = int(float(row["顺序"]))
        except (TypeError, ValueError):
            errors.append(f"流程“{flow}”的顺序必须为数字")
            continue
        if (flow, order) in used_orders:
            errors.append(f"流程“{flow}”存在重复顺序“{order}”")
        used_orders.add((flow, order))
        if row["启用"] not in {"是", "否"}:
            errors.append(f"流程“{flow}”的“启用”只能填是或否")
        if row["失败后处理"] not in {"停止", "跳过"}:
            errors.append(f"流程“{flow}”的“失败后处理”只能填停止或跳过")
        if row["是否输出结果"] not in {"是", "否"}:
            errors.append(f"流程“{flow}”的“是否输出结果”只能填是或否")
        if row["启用"] == "是" and row["是否输出结果"] == "是":
            flow_outputs[flow] = True
        feature_type = feature_types_by_name.get(feature)
        if row["启用"] == "是" and feature_type == AUDIT_RESULT_OUTPUT_FUNCTION and not row["输出"]:
            errors.append(f"流程“{flow}”的“{feature}”必须填写输出")
        if feature_type == AUDIT_RESULT_OUTPUT_FUNCTION and row["启用"] == "是" and row["是否输出结果"] != "是":
            errors.append(f"流程“{flow}”的“{feature}”必须填写“是否输出结果=是”")
        if any(character in row["输出文件名"] for character in '\\/:*?"<>|'):
            errors.append(f"流程“{flow}”的输出文件名包含非法字符")
    for flow in flow_names:
        if not flow_outputs.get(flow):
            errors.append(f"自定义流程“{flow}”至少要有一个启用步骤输出结果")
    flow_display = draft.get("flowDisplay", {})
    if not isinstance(flow_display, dict):
        errors.append("流程显示设置格式无效")
    else:
        all_flows = set(DEFAULT_FLOW_NAMES) | flow_names
        for flow, shown in flow_display.items():
            if _text(flow) not in all_flows:
                errors.append(f"显示设置引用不存在的流程“{_text(flow)}”")
            if not isinstance(shown, bool):
                errors.append(f"流程“{_text(flow)}”的显示设置必须为是或否")
    return errors


def save_config_editor_draft(config_path: Path, draft: object) -> Path:
    errors = validate_config_editor_draft(config_path, draft)
    if errors:
        raise ValueError("设置中心配置不合法：" + "；".join(errors))
    payload = _workflow(Path(config_path))
    payload["customModules"] = _editor_rows(draft.get("customModules", []), MAPPING_HEADERS, "自定义模块")
    payload["customFlows"] = _editor_rows(draft.get("customFlows", []), FLOW_HEADERS, "自定义流程")
    display = draft.get("flowDisplay", {})
    payload["flowDisplay"] = {
        _text(name): bool(value)
        for name, value in display.items()
        if _text(name) and bool(value)
    } if isinstance(display, dict) else {}
    payload.pop("buttons", None)
    if "combineSheetsPlans" in draft:
        payload["combineSheetsPlans"] = [
            {**DEFAULT_COMBINE_SHEETS_PLAN, **plan}
            for plan in draft["combineSheetsPlans"]
        ]
        payload["activeCombineSheetsPlanId"] = _text(
            draft.get("activeCombineSheetsPlanId")
        )
    elif "combineSheetsPlan" in draft:
        # Preserve compatibility with a settings window that only knows one
        # editable scheme: replace the current profile, retain other saved ones.
        plan = draft["combineSheetsPlan"]
        if not isinstance(plan, dict):
            raise ValueError("组合工作表分组方案格式无效")
        active_id = _text(payload.get("activeCombineSheetsPlanId")) or "default"
        replacement = {**DEFAULT_COMBINE_SHEETS_PLAN, **plan, "id": active_id}
        profiles = payload.get("combineSheetsPlans", [])
        if not isinstance(profiles, list):
            profiles = []
        replaced = False
        updated: list[dict[str, object]] = []
        for profile in profiles:
            if isinstance(profile, dict) and _text(profile.get("id")) == active_id:
                updated.append(replacement)
                replaced = True
            elif isinstance(profile, dict):
                updated.append(profile)
        if not replaced:
            updated.append(replacement)
        payload["combineSheetsPlans"] = updated
        payload["activeCombineSheetsPlanId"] = active_id
    _write_workflow(workflow_config_path(Path(config_path)), payload)
    return workflow_config_path(Path(config_path))


def load_combine_sheets_plan(config_path: Path) -> dict[str, object]:
    """Read the independently configurable grouping plan for 组合工作表."""
    payload = _workflow(Path(config_path))
    plans = payload.get("combineSheetsPlans", [])
    active_id = _text(payload.get("activeCombineSheetsPlanId"))
    errors = _validate_combine_sheets_profiles(plans, active_id)
    if errors:
        raise ValueError("；".join(errors))
    plan = next(
        item for item in plans
        if isinstance(item, dict) and _text(item.get("id")) == active_id
    )
    return {**DEFAULT_COMBINE_SHEETS_PLAN, **plan}
