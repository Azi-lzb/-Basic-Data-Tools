"""配置驱动的 UOS 原生流程执行器。

与 ``service.run_flow`` 的 Windows COM 实现保持同一契约：步骤顺序、配置的
处理对象、输出名与失败后处理都真实生效。COM 专属功能（组合联合核查表、
组合工作表、联合模板制作、机构合并）在本管线明确报“暂仅支持 Windows”，
不静默跳过。
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from ..discovery import source_workbooks
from ..feature_log import FeatureLog
from ..name_config import (
    AUDIT_RESULT_OUTPUT_FUNCTION,
    COMBINE_SHEETS_FUNCTION,
    CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
    EXTERNAL_FILE_FUNCTION,
    FORMULA_COPY_FUNCTION,
    ISSUE_EXTRACT_FUNCTION,
    MERGE_ORG_FILES_FUNCTION,
    NAMED_RANGE_CHECK_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    USED_RANGE_SUMMARY_FUNCTION,
    FIXED_ROW_SUMMARY_FUNCTION,
    WORKBOOK_TABLE_MERGE_FUNCTION,
    FlowStep,
    load_feature_mappings,
    load_flow_steps,
)
from ..preflight_xlsx import validate_source_xlsx, write_structure_report_xlsx
from .audit_flow import run_native_audit
from .openpyxl_workbook import named_ranges, read_template, template_structure_values
from .summary import merge_by_first_filename_part, merge_workbook_tables, run_region_summaries

COPY_PRODUCING_FEATURE_TYPES = (EXTERNAL_FILE_FUNCTION, FORMULA_COPY_FUNCTION)
COM_ONLY_FEATURE_TYPES = {MERGE_ORG_FILES_FUNCTION, COMBINE_SHEETS_FUNCTION}
SUMMARY_FEATURE_TYPES = {USED_RANGE_SUMMARY_FUNCTION, FIXED_ROW_SUMMARY_FUNCTION, WORKBOOK_TABLE_MERGE_FUNCTION}


def _selected_sources(
    input_dir: Path, *, recursive: bool, selected_files, excluded_files=(),
) -> list[Path]:
    """只返回工作台确认过的源工作簿；外部辅助文件不混进批处理。"""
    files = source_workbooks(input_dir, recursive=recursive)
    excluded = {Path(path).resolve() for path in excluded_files if path}
    files = [path for path in files if path.resolve() not in excluded]
    if selected_files is None:
        return files
    allowed = {path.resolve() for path in files}
    requested = [Path(path).resolve() for path in selected_files]
    invalid = [path for path in requested if path not in allowed]
    if invalid:
        raise ValueError("选择的待处理文件不在源数据目录中：{}".format(invalid[0]))
    requested_set = set(requested)
    return [path for path in files if path.resolve() in requested_set]


def _step_output_name(step: FlowStep) -> str:
    return (step.output_name or "{:02d}_{}".format(step.order, step.feature_name)).strip()


def _resolve_process_source(steps: tuple[FlowStep, ...], step: FlowStep, copy_producers: set[str]):
    """与 Windows service 相同的处理对象兜底语义。"""
    if not step.process_source or step.process_source == "源数据目录":
        return None, None
    prior = [item for item in steps if item.order < step.order]
    if step.process_source in copy_producers and any(item.feature_name == step.process_source for item in prior):
        return step.process_source, None
    fallback = next((item.feature_name for item in reversed(prior) if item.feature_name in copy_producers), None)
    if fallback:
        return fallback, "“{}”输入“{}”不可用，已使用最近前置副本“{}”".format(step.feature_name, step.process_source, fallback)
    return None, "“{}”输入“{}”不可用，已按源数据目录处理".format(step.feature_name, step.process_source)


def run_native_flow(
    *,
    flow_name: str,
    template_path: Path | None,
    input_dir: Path,
    output_dir: Path,
    history_path: Path,
    config_path: Path,
    external_path: Path | None = None,
    recursive: bool = True,
    period: str = "",
    on_step=None,
    write_flow_logs: bool = True,
    selected_files=None,
) -> dict:
    """执行一个配置流程；返回 {"output": str, "log": str} 供 web 层消费。"""
    from ..name_config import SOURCE_DIRECTORY_INPUT

    say = on_step or (lambda _text: None)
    steps = load_flow_steps(config_path, flow_name)
    if not steps:
        raise ValueError("执行流程“{}”不存在，或没有启用的功能".format(flow_name))
    mappings = {item.name: item for item in load_feature_mappings(config_path, template_path or Path("__无需模板__.xlsx"))}
    missing = [step.feature_name for step in steps if step.feature_name not in mappings]
    if missing:
        raise ValueError("执行流程引用了“模块化功能”中不存在的功能：" + "、".join(missing))

    com_only = [step.feature_name for step in steps if mappings[step.feature_name].feature_type in COM_ONLY_FEATURE_TYPES]
    if com_only:
        raise ValueError(
            "功能“{}”暂仅支持 Windows 外壳（Excel/WPS COM）；统信 UOS 版请使用其余功能".format("、".join(sorted(set(com_only))))
        )

    log = FeatureLog(flow_name, output_dir) if write_flow_logs else None
    feature_types = {mappings[step.feature_name].feature_type for step in steps}
    executable_types = feature_types - {NAMED_RANGE_CHECK_FUNCTION}
    copy_producers = {mapping.name for mapping in mappings.values() if mapping.feature_type in COPY_PRODUCING_FEATURE_TYPES}
    effective_sources: dict[int, str | None] = {}
    for step in steps:
        mapping = mappings[step.feature_name]
        effective, note = (None, None) if mapping.feature_type == AUDIT_RESULT_OUTPUT_FUNCTION else _resolve_process_source(steps, step, copy_producers)
        effective_sources[step.order] = effective
        if note:
            say("提示：" + note)

    try:
        input_files = _selected_sources(
            input_dir, recursive=recursive, selected_files=selected_files, excluded_files=(external_path,),
        )
        if not input_files:
            raise FileNotFoundError("源数据目录没有可处理的 Excel 文件")

        if template_path is None or not template_path.is_file():
            raise FileNotFoundError("该执行流程需要选择有效的模板文件")
        features = [mappings[step.feature_name] for step in steps]

        for step in steps:
            mapping = mappings[step.feature_name]
            if mapping.feature_type != NAMED_RANGE_CHECK_FUNCTION:
                continue
            book = load_workbook(template_path, read_only=True, data_only=False, keep_links=False)
            try:
                areas = named_ranges(book, [mapping])
            finally:
                book.close()
            rows = [(area.sheet_name, area.address, "通过") for area in areas] or [("", "、".join(mapping.range_names), "缺失")]
            if log:
                log.add_sheet(mapping.name, ("工作表", "区域", "结果"), rows)
            if not areas:
                message = "缺少命名区域：" + "、".join(mapping.range_names)
                if step.on_failure == "停止":
                    raise ValueError(message)
                say("跳过：" + message)

        formula_mappings = [item for item in features if item.feature_type == FORMULA_COPY_FUNCTION]
        structure_mappings = [item for item in features if item.feature_type == STRUCTURE_COMPARE_FUNCTION]
        extraction_mappings = [item for item in features if item.feature_type == ISSUE_EXTRACT_FUNCTION]
        definition = read_template(template_path, formula_mappings=formula_mappings, structure_mappings=structure_mappings, extraction_mappings=extraction_mappings)

        processable_files = input_files
        structure_report_path: Path | None = None
        if STRUCTURE_COMPARE_FUNCTION in feature_types:
            expected = template_structure_values(template_path, definition)
            rows, matched_files, source_matches = [], [], []
            for source in input_files:
                if source.suffix.casefold() == ".xls":
                    rows.append((source.name, "跳过", "旧版 .xls 仅支持汇总读取；表结构比对请先转换为 .xlsx"))
                    continue
                match = validate_source_xlsx(source, definition, expected)
                source_matches.append((source, match))
                rows.append((source.name, "通过" if match.matched else "不通过", match.details))
                if match.matched:
                    matched_files.append(source)
            processable_files = matched_files
            if log:
                log.add_sheet("表结构比对", ("文件", "结果", "说明"), rows)
            structure_step = next(step for step in steps if mappings[step.feature_name].feature_type == STRUCTURE_COMPARE_FUNCTION)
            if structure_step.output_result:
                structure_report_path = output_dir / "{}_{}.xlsx".format(_step_output_name(structure_step), datetime.now().strftime("%Y%m%d%H%M%S"))
                write_structure_report_xlsx(structure_report_path, source_matches)
            say("表结构比对：通过 {} 个，跳过 {} 个".format(len(processable_files), len(input_files) - len(processable_files)))

        audit_types = {EXTERNAL_FILE_FUNCTION, FORMULA_COPY_FUNCTION, ISSUE_EXTRACT_FUNCTION, AUDIT_RESULT_OUTPUT_FUNCTION, CONDITIONAL_FORMAT_EXTRACT_FUNCTION}
        summary_steps = [step for step in steps if mappings[step.feature_name].feature_type in SUMMARY_FEATURE_TYPES]
        audit_steps = [step for step in steps if mappings[step.feature_name].feature_type in audit_types]
        conditional_steps = [step for step in audit_steps if mappings[step.feature_name].feature_type == CONDITIONAL_FORMAT_EXTRACT_FUNCTION]
        conditional_mappings = [mappings[step.feature_name] for step in conditional_steps]
        conditional_on_failure = conditional_steps[0].on_failure if conditional_steps else "停止"
        summary_output_steps = [step for step in summary_steps if step.output_result]
        summary_output_name = _step_output_name(summary_output_steps[-1]) if summary_output_steps else ""
        if summary_steps and not summary_output_steps:
            raise ValueError("汇总流程至少应有一个汇总模块填写“是否输出结果=是”")
        if WORKBOOK_TABLE_MERGE_FUNCTION in feature_types and STRUCTURE_COMPARE_FUNCTION in feature_types:
            raise ValueError("“汇总表合并”会自行按表头判断合并，不应与“表结构比对”放在同一流程")
        copy_source = next((effective_sources[step.order] for step in summary_steps if effective_sources[step.order]), None)
        if audit_steps and summary_steps and not copy_source:
            raise ValueError("执行流程不能混用审核功能与汇总功能；汇总步骤的“输入”应指向前置副本功能")

        audit_result = None
        if audit_steps:
            if not processable_files:
                raise ValueError("表结构比对后没有可继续处理的文件")
            formula_step = next((step for step in audit_steps if mappings[step.feature_name].feature_type == FORMULA_COPY_FUNCTION), None)
            external_step = next((step for step in audit_steps if mappings[step.feature_name].feature_type == EXTERNAL_FILE_FUNCTION), None)
            output_step = next((step for step in audit_steps if mappings[step.feature_name].feature_type == AUDIT_RESULT_OUTPUT_FUNCTION and step.output_result), None)
            keep_copies = bool((formula_step and formula_step.output_result) or (external_step and external_step.output_result) or copy_source)
            copy_label_step = formula_step if formula_step and formula_step.output_result else external_step
            suffix = datetime.now().strftime("%Y%m%d%H%M%S")
            copies_dir = output_dir / ("{}_{}".format(_step_output_name(copy_label_step), suffix) if copy_label_step else "_审核工作副本")
            audit_result = run_native_audit(
                input_dir=input_dir, template_path=template_path, output_dir=output_dir,
                audit_copies_dir=copies_dir, config_path=config_path, external_path=external_path,
                history_path=history_path, recursive=recursive, period=period, on_step=say,
                formula_mappings=formula_mappings, structure_mappings=structure_mappings,
                extraction_mappings=extraction_mappings, conditional_mappings=conditional_mappings,
                include_external=EXTERNAL_FILE_FUNCTION in feature_types,
                include_formula_copy=FORMULA_COPY_FUNCTION in feature_types,
                include_formula_extraction=ISSUE_EXTRACT_FUNCTION in feature_types,
                write_summary=bool(output_step),
                summary_output_name=_step_output_name(output_step) if output_step else "",
                update_history=bool({mappings[step.feature_name].feature_type for step in audit_steps} & {ISSUE_EXTRACT_FUNCTION, CONDITIONAL_FORMAT_EXTRACT_FUNCTION}),
                keep_audit_copies=keep_copies,
                stop_on_file_failure=any(step.on_failure == "停止" for step in audit_steps),
                conditional_on_failure=conditional_on_failure,
                selected_files=processable_files,
            )
            if log:
                log.add_sheet("审核结果", ("文件", "问题数", "计算引擎", "结果"), [
                    (item.source_path.name, len(item.issues), audit_result.calculation_engine if not item.error else "未计算", item.error or "成功")
                    for item in audit_result.files
                ])

        if summary_steps:
            if not processable_files:
                raise ValueError("表结构比对后没有可汇总文件")
            if WORKBOOK_TABLE_MERGE_FUNCTION in feature_types:
                step = next(step for step in summary_steps if mappings[step.feature_name].feature_type == WORKBOOK_TABLE_MERGE_FUNCTION)
                path = merge_workbook_tables(input_dir=input_dir, output_dir=output_dir, recursive=recursive, selected_files=processable_files, on_step=say, output_name=_step_output_name(step))
                if log:
                    log.add_sheet("汇总表合并", ("输出文件", "结果"), [(str(path), "成功")])
            else:
                summary_input, summary_selected = input_dir, processable_files
                if audit_result is not None and copy_source:
                    copies_dir = audit_result.copies.get(copy_source)
                    if copies_dir is None:
                        raise ValueError("前置功能“{}”没有保留可汇总的审核副本".format(copy_source))
                    summary_input, summary_selected = copies_dir, None
                path = run_region_summaries(
                    template_path=template_path, input_dir=summary_input, output_dir=output_dir,
                    features=[mappings[step.feature_name] for step in summary_steps],
                    recursive=recursive, selected_files=summary_selected, on_step=say,
                    output_name=summary_output_name, history_config_path=config_path,
                )
                if log:
                    log.add_sheet("区域汇总", ("输出文件", "结果"), [(str(path), "成功")])
            return {"output": str(path), "log": _write_log(log)}

        if audit_result is not None:
            return {"output": str(audit_result.summary_path or audit_result.output_dir), "log": _write_log(log)}
        log_path = _write_log(log)
        return {"output": str(structure_report_path or log_path), "log": log_path}
    except Exception:
        if log:
            log.write()
        raise


def _write_log(log: FeatureLog | None) -> str:
    if log is None:
        return ""
    path = log.write()
    return str(path) if path else ""
