from __future__ import annotations

import hashlib
import shutil
import time
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from .discovery import source_workbooks
from .excel_com import ExcelSession
from .region_summary import RegionSummaryItem, RegionSummaryResult
from .external import ExternalSheetPlan, make_external_sheet_plan
from .history import (
    HISTORY_AUDIT_SHEET,
    LOCAL_VALIDATION_HISTORY_SHEET,
    merge_history,
)
from .name_config import (
    AUDIT_RESULT_OUTPUT_FUNCTION,
    CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
    EXTERNAL_FILE_FUNCTION,
    COMBINE_SHEETS_FUNCTION,
    FeatureMapping,
    FORMULA_COPY_FUNCTION,
    ISSUE_EXTRACT_FUNCTION,
    NAMED_RANGE_CHECK_FUNCTION,
    MERGE_ORG_FILES_FUNCTION,
    FIXED_ROW_SUMMARY_FUNCTION,
    STRUCTURE_COMPARE_FUNCTION,
    USED_RANGE_SUMMARY_FUNCTION,
    WORKBOOK_TABLE_MERGE_FUNCTION,
    SOURCE_DIRECTORY_INPUT,
    features_of_type,
    load_combine_sheets_plan, load_feature_mappings,
    load_flow_steps,
)
from .models import (
    AuditRunResult,
    FileAuditResult,
    PreflightItem,
    PreflightRunResult,
    SourceMatch,
    TemplateDefinition,
)
from .feature_log import FeatureLog
from .merge_org import MergeOrgResult, TemplateMergeResult, run_combine_sheets, run_merge_org, run_template_merge
from .preflight_xlsx import (
    template_structure_values,
    validate_required_sheets_xlsx,
    validate_source_xlsx,
    write_preflight_report_xlsx,
)
from .template import TemplateError


CONFIG_HISTORY_SHEET = HISTORY_AUDIT_SHEET
# 会产出阶段/审核副本的功能类型：只有它们能作为“处理对象”的数据来源。
COPY_PRODUCING_FEATURE_TYPES = (EXTERNAL_FILE_FUNCTION, FORMULA_COPY_FUNCTION)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _organisation_from_name(path: Path) -> tuple[str, str]:
    parts = [part.strip() for part in path.stem.split("_") if part.strip()]
    if len(parts) >= 2:
        return parts[0], parts[1]
    return path.stem, path.stem


def _available_output_path(folder: Path, source: Path) -> Path:
    candidate = folder / f"{source.stem}_审核版.xlsx"
    if not candidate.exists():
        return candidate
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return folder / f"{source.stem}_审核版_{stamp}.xlsx"


def _output_prefix(order: int, feature_name: str, output_name: str = "") -> str:
    """输出前缀：填写“输出文件名”列时用该值，否则沿用“序号_功能名”。"""
    return output_name or f"{order:02d}_{feature_name}"


def _resolve_process_source(steps, step, copy_producers):
    """按兜底语义解析一步的“输入”。

    ``copy_producers`` 是流程中会产副本（阶段副本/审核副本）的功能名集合；
    只有它们能作为数据来源，检查类/核对类等不产副本的功能不能被引用。

    返回 ``(effective, message)``：
    - 留空 → ``(None, None)``：读源数据目录。
    - 匹配流程中排在本步之前、且会产副本的功能名 → ``(该功能名, None)``。
    - 不匹配或引用了不产副本的功能 → 跳过非副本功能，回退到最近前一个产副本功能；
      本步之前没有产副本功能则回退源数据目录，并返回一条提示由调用方写入运行记录。
    """
    if not step.process_source or step.process_source == SOURCE_DIRECTORY_INPUT:
        return None, None
    prior = [s for s in steps if s.order < step.order]
    if step.process_source in copy_producers and any(s.feature_name == step.process_source for s in prior):
        return step.process_source, None
    fallback = next(
        (s.feature_name for s in reversed(prior) if s.feature_name in copy_producers),
        None,
    )
    if fallback is not None:
        return fallback, (
            f"流程“{step.flow_name}”的“{step.feature_name}”输入"
            f"“{step.process_source}”不是本流程中排在其前且会产副本的功能，"
            f"已按最近前一个产副本功能“{fallback}”兜底处理"
        )
    if not prior:
        return None, (
            f"流程“{step.flow_name}”的“{step.feature_name}”是流程第一个功能，"
            f"输入“{step.process_source}”无法匹配前置功能，已按源数据目录兜底处理"
        )
    return None, (
        f"流程“{step.flow_name}”的“{step.feature_name}”输入"
        f"“{step.process_source}”之前没有会产副本的功能，已按源数据目录兜底处理"
    )


def _external_sheet_plan(
    excel: ExcelSession,
    template_workbook: object,
    definition: object,
    external_workbook: object,
) -> ExternalSheetPlan:
    # 模板规则上千条时逐格 COM 取公式是热路径（每格一次跨进程调用，可达数秒）。
    # 按工作表一次性读取 UsedRange.Formula 二维数组，再用单元格坐标取回规则
    # 公式；语义与逐格读取一致（只看启用规则的公式）。
    from .excel_com import _cell_position

    cells_by_sheet: dict[str, set[str]] = {}
    for rule in definition.rules:
        if rule.enabled:
            cells_by_sheet.setdefault(rule.sheet_name, set()).add(rule.formula_cell)
    formulas: list[object] = []
    for sheet_name, cells in sorted(cells_by_sheet.items()):
        sheet = template_workbook.Worksheets(sheet_name)
        used = sheet.UsedRange
        matrix = used.Formula
        top, left = int(used.Row), int(used.Column)
        if not isinstance(matrix, tuple):
            matrix = ((matrix,),)
        width = max(len(row) for row in matrix) if matrix else 0
        for address in sorted(cells):
            row, column = _cell_position(address)
            r, c = row - top, column - left
            if 0 <= r < len(matrix):
                row_values = matrix[r]
                if isinstance(row_values, tuple) and 0 <= c < len(row_values):
                    formulas.append(row_values[c])
                    continue
                if not isinstance(row_values, tuple) and c == 0:
                    formulas.append(row_values)
                    continue
            # 单元格在 UsedRange 之外（理论不应发生）：回退单格读取。
            formulas.append(sheet.Range(address).Formula)
    return make_external_sheet_plan(
        formulas=formulas,
        available_sheets=excel._worksheet_names(external_workbook),
    )
def _source_files(
    input_dir: Path,
    selected_files: list[Path] | None = None,
    *,
    recursive: "bool | int" = False,
    extra_files: list[Path] | None = None,
) -> list[Path]:
    """源目录扫描结果 + 用户手动追加的文件，共同构成待处理清单。"""
    all_files = list(source_workbooks(input_dir, recursive=recursive))
    extra = [
        Path(path)
        for path in (extra_files or [])
        if Path(path).is_file()
        and Path(path).resolve() not in {path.resolve() for path in all_files}
    ]
    all_files.extend(extra)
    all_files.sort(key=lambda path: path.name)
    if selected_files is None:
        return all_files
    allowed = {path.resolve() for path in all_files}
    requested = [Path(path).resolve() for path in selected_files]
    invalid = [path for path in requested if path not in allowed]
    if invalid:
        raise ValueError("选择的待审核文件不在源数据目录中：" + str(invalid[0]))
    return [path for path in all_files if path.resolve() in set(requested)]


class AuditService:
    def __init__(
        self, *, config_path: Path | None = None, engine_preference: str = "自动"
    ) -> None:
        self.config_path = config_path
        self.engine_preference = engine_preference

    def summarize_regions(
        self,
        *,
        template_path: Path,
        input_dir: Path,
        output_dir: Path,
        selected_files: list[Path] | None = None,
        flow_name: str | None = None,
        recursive: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        named_range_features: tuple[FeatureMapping, ...] = (),
        output_name: str | None = None,
        feature_log: Optional[FeatureLog] = None,
        copies_dir: Path | None = None,
    ) -> "RegionSummaryResult":
        from .engines import pipeline_kind
        from .name_config import load_flow_features

        if self.config_path is None:
            raise ValueError("汇总功能需要历史审核配置.xlsx")
        flow_features = load_flow_features(self.config_path, flow_name) if flow_name else None
        if flow_name and not flow_features:
            raise ValueError(f"执行流程“{flow_name}”没有启用的功能")
        if pipeline_kind(self.engine_preference) == "native":
            from .native.summary import merge_workbook_tables, run_region_summaries as native_run_region_summaries

            if named_range_features:
                features = list(named_range_features)
            elif flow_name:
                names = flow_features
                mappings = {item.name: item for item in load_feature_mappings(self.config_path, template_path)}
                features = [mappings[name] for name in names if name in mappings]
            else:
                features = []
            merge_features = [item for item in features if item.feature_type == WORKBOOK_TABLE_MERGE_FUNCTION]
            row_features = [item for item in features if item.feature_type != WORKBOOK_TABLE_MERGE_FUNCTION]
            output_dir = output_dir.resolve()
            if merge_features and not row_features:
                path = merge_workbook_tables(
                    input_dir=input_dir, output_dir=output_dir, recursive=recursive,
                    selected_files=selected_files, on_step=on_step,
                    output_name=output_name or "汇总表合并",
                )
                return RegionSummaryResult(output_path=path, items=[RegionSummaryItem("汇总表合并", 0, "")])
            if not row_features:
                raise ValueError("没有启用“汇总_任意行汇总”或“汇总_固定行汇总”模块")
            path = native_run_region_summaries(
                template_path=template_path, input_dir=input_dir, output_dir=output_dir,
                features=row_features, recursive=recursive, selected_files=selected_files,
                on_step=on_step, output_name=output_name or "区域汇总",
                history_config_path=self.config_path,
            )
            items = [RegionSummaryItem(item.name, 0, "") for item in row_features]
            return RegionSummaryResult(output_path=path, items=items)
        from .region_summary import run_region_summaries
        return run_region_summaries(
            template_path=template_path,
            input_dir=input_dir,
            output_dir=output_dir,
            config_path=self.config_path,
            selected_files=selected_files,
            feature_names=flow_features,
            flow_name=flow_name,
            recursive=recursive,
            on_step=on_step,
            named_range_features=named_range_features,
            output_name=output_name,
            feature_log=feature_log,
            copies_dir=copies_dir,
            engine_preference=self.engine_preference,
        )

    def merge_org_files(
        self,
        *,
        input_dir: Path,
        output_dir: Path,
        period: str = "",
        selected_files: list[Path] | None = None,
        flow_name: str | None = None,
        recursive: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        feature_log: Optional[FeatureLog] = None,
        output_name: str | None = None,
    ) -> MergeOrgResult:
        """Run the standalone preparation flow that merges each institution's workbooks."""
        return run_merge_org(
            input_dir=input_dir,
            output_dir=output_dir,
            period=period,
            selected_files=selected_files,
            flow_name=flow_name,
            recursive=recursive,
            on_step=on_step,
            feature_log=feature_log,
            output_name=output_name,
            engine_preference=self.engine_preference,
        )

    def combine_sheets(
        self,
        *,
        input_dir: Path,
        output_dir: Path,
        period: str = "",
        selected_files: list[Path] | None = None,
        flow_name: str | None = None,
        recursive: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        feature_log: Optional[FeatureLog] = None,
        output_name: str | None = None,
    ) -> MergeOrgResult:
        if self.config_path is None:
            raise ValueError("组合工作表需要流程配置")
        plan = load_combine_sheets_plan(self.config_path)
        if on_step is not None:
            on_step(f"组合分组方案：{plan['name']}（{plan['mode']}）")
        return run_combine_sheets(
            input_dir=input_dir, output_dir=output_dir, period=period,
            selected_files=selected_files, flow_name=flow_name, recursive=recursive,
            on_step=on_step, feature_log=feature_log, output_name=output_name,
            engine_preference=self.engine_preference, grouping_plan=plan,
        )

    def is_standalone_combine_flow(self, flow_name: str) -> bool:
        """Whether a flow needs only source files, not a template or external file.

        The workbench must inspect configured steps rather than compare the
        displayed flow name: users can create a custom button or copy a flow
        that contains the built-in combine module.
        """
        if self.config_path is None:
            return False
        steps = load_flow_steps(self.config_path, flow_name)
        if not steps:
            return False
        mappings = {
            item.name: item
            for item in load_feature_mappings(self.config_path, Path("__无需模板__.xlsx"))
        }
        types = {
            mappings[step.feature_name].feature_type
            for step in steps
            if step.feature_name in mappings
        }
        executable = types - {NAMED_RANGE_CHECK_FUNCTION}
        return executable in ({MERGE_ORG_FILES_FUNCTION}, {COMBINE_SHEETS_FUNCTION})

    def merge_template_files(
        self,
        *,
        base_template: Path,
        source_templates: list[Path],
        on_step: Optional[Callable[[str], None]] = None,
    ) -> TemplateMergeResult:
        """Create a combined template from explicitly selected workbooks.

        This is intentionally outside config-driven audit flows: it is a
        manual, low-frequency template authoring action and never examines the
        workbench's source-data or auto-matched template fields.
        """
        from .engines import pipeline_kind

        if pipeline_kind(self.engine_preference) == "native":
            raise ValueError("制作联合模板暂仅支持 Windows 外壳（Excel/WPS COM）；统信 UOS 版暂不提供该功能。")

        return run_template_merge(
            base_template=base_template,
            source_templates=source_templates,
            on_step=on_step,
            engine_preference=self.engine_preference,
        )

    def run_flow(
        self,
        *,
        flow_name: str,
        template_path: Path | None,
        input_dir: Path,
        output_dir: Path,
        period: str,
        history_path: Path,
        selected_files: list[Path] | None = None,
        external_path: Path | None = None,
        extra_files: list[Path] | None = None,
        recursive: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        strict: bool = True,
        write_flow_logs: bool = True,
    ) -> "AuditRunResult | RegionSummaryResult | PreflightRunResult":
        """Execute one configured flow; its steps are the single source of truth.

        ``strict`` 控制组合前置缺失时的处理：主流程(True)报错，自定义流程(False)只警告。
        """
        if self.config_path is None:
            raise ValueError("按流程执行需要 data/历史审核配置.xlsx")
        # 管线分派（engines 收口平台判断）：UOS/麒麟走原生管线，Windows 走 COM。
        from .engines import pipeline_kind

        if pipeline_kind(self.engine_preference) == "native":
            from .native.flow import run_native_flow

            outcome = run_native_flow(
                flow_name=flow_name,
                template_path=template_path,
                input_dir=input_dir,
                output_dir=output_dir,
                history_path=history_path,
                config_path=self.config_path,
                external_path=external_path,
                recursive=recursive,
                period=period,
                on_step=on_step,
                write_flow_logs=write_flow_logs,
                selected_files=selected_files,
            )
            return _NativeFlowResult(outcome)

        steps = load_flow_steps(self.config_path, flow_name)
        if not steps:
            raise ValueError(f"执行流程“{flow_name}”不存在，或没有启用的功能")
        # “组合联合核查表”不使用模板；给配置读取一个占位文件名只是为了
        # 沿用模块化功能的加载逻辑，不会打开或访问该路径。
        mapping_template = template_path or Path("__无需模板__.xlsx")
        mappings = {item.name: item for item in load_feature_mappings(self.config_path, mapping_template)}
        missing = [step.feature_name for step in steps if step.feature_name not in mappings]
        if missing:
            raise ValueError("执行流程引用了“模块化功能”中不存在的功能：" + "、".join(missing))
        # 处理对象兜底解析：只有会产副本的功能（外部文件添加/公式校验复制）能作为
        # 数据来源；不匹配或引用不产副本功能时，跳过非副本功能回退到最近前一个产
        # 副本功能（本步之前没有产副本功能则回退源数据目录），并给出提示。
        # 主流程与自定义流程都走兜底，不再按 strict 报错。
        copy_producers = {
            mapping.name for mapping in mappings.values()
            if mapping.feature_type in COPY_PRODUCING_FEATURE_TYPES
        }
        effective_sources: dict[int, str | None] = {}
        for step in steps:
            if mappings[step.feature_name].feature_type == AUDIT_RESULT_OUTPUT_FUNCTION:
                # 最终输出读取“输出”同名的结果集，不是工作簿副本；“输入”
                # 仅在界面中显示该结果集名称，不参与副本来源解析。
                effective, message = None, None
            else:
                effective, message = _resolve_process_source(steps, step, copy_producers)
            effective_sources[step.order] = effective
            if message is not None and on_step is not None:
                on_step(f"提示：{message}")
        feature_types = {mappings[step.feature_name].feature_type for step in steps}
        named_range_steps = tuple(
            (step.order, mappings[step.feature_name]) for step in steps
            if mappings[step.feature_name].feature_type == NAMED_RANGE_CHECK_FUNCTION
        )
        # preflight 需要顺序以命名运行日志 sheet 标题；汇总引擎只需要映射本身。
        named_range_checks = named_range_steps
        named_range_mappings = tuple(mapping for _, mapping in named_range_steps)
        executable_types = feature_types - {NAMED_RANGE_CHECK_FUNCTION}
        standalone_combine_types = {MERGE_ORG_FILES_FUNCTION, COMBINE_SHEETS_FUNCTION}
        if executable_types.intersection(standalone_combine_types) and executable_types not in ({MERGE_ORG_FILES_FUNCTION}, {COMBINE_SHEETS_FUNCTION}):
            raise ValueError("“组合联合核查表”或“组合工作表”必须单独成一个流程，不能与其他功能混用")
        if executable_types not in ({MERGE_ORG_FILES_FUNCTION}, {COMBINE_SHEETS_FUNCTION}) and (
            template_path is None or not template_path.is_file()
        ):
            raise FileNotFoundError("该执行流程需要选择有效的模板文件")
        summary_types = {
            USED_RANGE_SUMMARY_FUNCTION,
            FIXED_ROW_SUMMARY_FUNCTION,
            WORKBOOK_TABLE_MERGE_FUNCTION,
        }
        summary_output_steps = [
            (step, mappings[step.feature_name]) for step in steps
            if mappings[step.feature_name].feature_type in summary_types and step.output_result
        ]
        summary_output_name = None
        if summary_output_steps:
            output_step, output_mapping = summary_output_steps[-1]
            summary_output_name = _output_prefix(
                output_step.order, output_mapping.name, output_step.output_name
            )
        # 一个流程生成一份运行日志 xlsx，每个功能一个 sheet；检查/核对类步骤
        # 填“是否输出结果=是”时，运行日志即为其可保留输出，用于单独的检查流程。
        feature_log = FeatureLog(flow_name, output_dir) if write_flow_logs else None
        try:
            if executable_types in ({MERGE_ORG_FILES_FUNCTION}, {COMBINE_SHEETS_FUNCTION}):
                merge_steps = [
                    (step, mappings[step.feature_name]) for step in steps
                    if mappings[step.feature_name].feature_type in standalone_combine_types
                ]
                if len(merge_steps) != 1:
                    raise ValueError("组合流程必须且只能启用一个合并模块")
                merge_step, merge_mapping = merge_steps[0]
                if not merge_step.output_result:
                    raise ValueError("组合流程必须填写“是否输出结果=是”")
                merger = self.merge_org_files if merge_mapping.feature_type == MERGE_ORG_FILES_FUNCTION else self.combine_sheets
                result = merger(
                    input_dir=input_dir,
                    output_dir=output_dir,
                    period=period,
                    selected_files=selected_files,
                    flow_name=flow_name,
                    recursive=recursive,
                    on_step=on_step,
                    feature_log=feature_log,
                    output_name=_output_prefix(
                        merge_step.order, merge_mapping.name, merge_step.output_name
                    ),
                )
            elif executable_types.issubset(summary_types):
                if not summary_output_name:
                    raise ValueError("汇总流程至少应有一个汇总模块填写“是否输出结果=是”")
                result = self.summarize_regions(
                    template_path=template_path, input_dir=input_dir, output_dir=output_dir,
                    selected_files=selected_files, flow_name=flow_name, recursive=recursive,
                    on_step=on_step,
                    named_range_features=named_range_mappings,
                    output_name=summary_output_name,
                    feature_log=feature_log,
                )
            elif executable_types.issubset(summary_types | {STRUCTURE_COMPARE_FUNCTION}):
                summary_steps = [
                    step for step in steps
                    if mappings[step.feature_name].feature_type in summary_types
                ]
                if summary_steps and not summary_output_name:
                    raise ValueError("汇总流程至少应有一个汇总模块填写“是否输出结果=是”")
                if WORKBOOK_TABLE_MERGE_FUNCTION in executable_types:
                    raise ValueError("“汇总表合并”会自行按表头判断合并，不应与“表结构比对”放在同一流程")
                structure_steps = [
                    step for step in steps
                    if mappings[step.feature_name].feature_type == STRUCTURE_COMPARE_FUNCTION
                ]
                if summary_steps:
                    first_summary_order = min(step.order for step in summary_steps)
                    if any(step.order >= first_summary_order for step in structure_steps):
                        raise ValueError("执行流程配置错误：“表结构比对”必须排在汇总功能之前")
                # 表结构比对的结果并入运行日志的“表结构比对”sheet，不再单独输出报告。
                check = self.preflight(
                    template_path=template_path,
                    input_dir=input_dir,
                    output_dir=output_dir,
                    selected_files=selected_files,
                    external_path=external_path,
                    recursive=recursive,
                    summary_feature_names=tuple(
                        step.feature_name for step in steps
                        if mappings[step.feature_name].feature_type in summary_types
                    ),
                    write_report=False,
                    on_step=on_step,
                    named_range_features=named_range_checks,
                    feature_log=feature_log,
                )
                matched_files = [path for path, match in check.source_matches if match.matched]
                skipped_files = tuple((path, match) for path, match in check.source_matches if not match.matched)
                if summary_steps:
                    # 汇总流程：结构通过的文件继续做区域汇总。
                    if not matched_files:
                        names = "、".join(path.name for path, _ in skipped_files)
                        raise ValueError(
                            f"表结构比对后没有可汇总文件，已跳过 {len(skipped_files)} 个文件：{names}"
                        )
                    result = self.summarize_regions(
                        template_path=template_path,
                        input_dir=input_dir,
                        output_dir=output_dir,
                        selected_files=matched_files,
                        flow_name=flow_name,
                        recursive=recursive,
                        on_step=on_step,
                        output_name=summary_output_name,
                        feature_log=feature_log,
                    )
                    result = replace(result, skipped_files=skipped_files)
                else:
                    # 纯表结构检查流程：只做表结构比对并写入运行日志，不汇总。
                    result = check
            elif executable_types.intersection(summary_types):
                # 受控组合：汇总步骤声明“处理对象=外部文件添加/公式校验复制”，
                # 且流程含对应修改步骤时，先跑修改类生成副本，再对副本做汇总。
                summary_steps = [
                    step for step in steps
                    if mappings[step.feature_name].feature_type in summary_types
                ]
                copies_source = next(
                    (effective_sources.get(step.order) for step in summary_steps
                     if effective_sources.get(step.order) in ("外部文件添加", "公式校验复制")),
                    None,
                )
                if copies_source:
                    audit = self.run(
                        template_path=template_path, input_dir=input_dir, output_dir=output_dir,
                        period=period, history_path=history_path, selected_files=selected_files,
                        external_path=external_path, extra_files=extra_files, flow_name=flow_name, recursive=recursive,
                        on_step=on_step, feature_log=feature_log, strict=strict,
                        effective_sources=effective_sources,
                    )
                    copies_dir = (audit.copies or {}).get(copies_source)
                    if copies_dir is None or not copies_dir.is_dir():
                        raise ValueError(
                            f"执行流程“{flow_name}”的“{copies_source}”没有输出可汇总的副本，"
                            f"请确认该步骤“是否输出结果=是”"
                        )
                    result = self.summarize_regions(
                        template_path=template_path, input_dir=input_dir, output_dir=output_dir,
                        selected_files=selected_files, flow_name=flow_name, recursive=recursive,
                        on_step=on_step,
                        named_range_features=named_range_mappings,
                        output_name=summary_output_name,
                        feature_log=feature_log,
                        copies_dir=copies_dir,
                    )
                else:
                    raise ValueError("执行流程不能混用审核功能与汇总功能，请拆分为两个流程")
            else:
                result = self.run(
                    template_path=template_path, input_dir=input_dir, output_dir=output_dir,
                    period=period, history_path=history_path, selected_files=selected_files,
                    external_path=external_path, extra_files=extra_files, flow_name=flow_name, recursive=recursive,
                    on_step=on_step, feature_log=feature_log, strict=strict,
                    effective_sources=effective_sources,
                )
            log_path = feature_log.write() if feature_log is not None else None
            if log_path is not None and result is not None:
                result = replace(result, log_path=log_path)
            return result
        except Exception:
            # 流程中途失败时也写一份已收集 sheet 的运行日志，便于排错。
            try:
                if feature_log is not None:
                    feature_log.write()
            except Exception:
                pass
            raise

    def _history_storage(self, legacy_path: Path) -> tuple[Path, str, Path | None]:
        """Keep operational history in the configured history workbook.

        The supplied path is retained only as a one-time migration source, so
        existing users do not lose their old independent history workbook.
        """
        if self.config_path is not None:
            return self.config_path.resolve(), CONFIG_HISTORY_SHEET, legacy_path.resolve()
        return legacy_path.resolve(), "问题历史", None

    def preflight(
        self,
        *,
        template_path: Path,
        input_dir: Path,
        output_dir: Path,
        selected_files: list[Path] | None = None,
        external_path: Path | None = None,
        extra_files: list[Path] | None = None,
        recursive: bool = False,
        summary_feature_names: tuple[str, ...] | None = None,
        write_report: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        named_range_features: tuple[tuple[int, FeatureMapping], ...] = (),
        feature_log: Optional[FeatureLog] = None,
    ) -> PreflightRunResult:
        from .engines import pipeline_kind

        if pipeline_kind(self.engine_preference) == "native":
            from .native.preflight import run_native_preflight

            return run_native_preflight(
                template_path=template_path,
                input_dir=input_dir,
                output_dir=output_dir,
                config_path=self.config_path,
                selected_files=selected_files,
                external_path=external_path,
                recursive=recursive,
                write_report=write_report,
                on_step=on_step,
                named_range_features=named_range_features,
                feature_log=feature_log,
            )
        template_path = template_path.resolve()
        input_dir = input_dir.resolve()
        output_dir = output_dir.resolve()
        if external_path is not None:
            external_path = external_path.resolve()
        if not template_path.is_file():
            raise FileNotFoundError(f"模板不存在：{template_path}")
        if not input_dir.is_dir():
            raise FileNotFoundError(f"源数据目录不存在：{input_dir}")
        if external_path is not None and not external_path.is_file():
            raise FileNotFoundError(f"外部文件不存在：{external_path}")
        source_files = _source_files(
            input_dir, selected_files, recursive=recursive, extra_files=extra_files
        )
        if not source_files:
            raise ValueError("源数据目录中没有可检查的 .xlsx 文件")
        output_dir.mkdir(parents=True, exist_ok=True)
        batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
        report_path = output_dir / f"审核前检查_{batch_id}.xlsx"

        with ExcelSession(self.engine_preference) as excel:
            template_workbook = excel.open_workbook(template_path, read_only=True)
            external_workbook = None
            external_plan = None
            try:
                try:
                    template_items: list[PreflightItem] = []
                    for _order, mapping in named_range_features:
                        if on_step is not None:
                            on_step(f"正在执行：{mapping.name}")
                        # 先调查再强校验：区域缺失时运行日志仍保留“缺失”记录。
                        if feature_log is not None:
                            survey = excel.survey_named_ranges(template_workbook, mapping)
                            feature_log.add_sheet(
                                mapping.name, ("工作表", "命名区域名", "覆盖区域", "结果"), survey
                            )
                            if on_step is not None:
                                found = [row for row in survey if row[3] == "通过"]
                                if found:
                                    per_sheet: dict[str, int] = {}
                                    for row in found:
                                        per_sheet[row[0]] = per_sheet.get(row[0], 0) + 1
                                    detail = "、".join(f"{sheet} {count} 处" for sheet, count in per_sheet.items())
                                    on_step(f"{mapping.name}：{detail}")
                                else:
                                    on_step(f"{mapping.name}：未找到 {mapping.range_names}")
                        ranges = excel.require_named_ranges(template_workbook, mapping)
                        template_items.append(
                            PreflightItem(
                                "命名区域检查", "提示", "通过", "", "",
                                f"模块“{mapping.name}”已找到 {len(ranges)} 个区域",
                            )
                        )
                        if on_step is not None:
                            on_step(f"完成：{mapping.name}")
                    if not summary_feature_names:
                        # 纯表结构检查流程没有汇总步骤（summary_feature_names 为空），
                        # 同样走模板体检，不做“汇总结构”准备。
                        definition = excel.read_template(
                            template_workbook, config_path=self.config_path
                        )
                        template_items.extend(excel.inspect_template_health(template_workbook, definition))
                    else:
                        mappings = load_feature_mappings(self.config_path, template_path)
                        selected = [item for item in mappings if item.name in summary_feature_names]
                        data_ranges = []
                        for mapping in selected:
                            data_ranges.extend(excel._named_ranges_for_features(template_workbook, [mapping]))
                        headers = excel._named_ranges_for_features(
                            template_workbook,
                            [FeatureMapping("汇总表头", USED_RANGE_SUMMARY_FUNCTION, ("表头区域",), False, "", "")],
                        )
                        if not data_ranges:
                            raise TemplateError("汇总功能未找到命名区域")
                        if not headers:
                            raise TemplateError("汇总功能缺少“表头区域”命名区域")
                        definition = TemplateDefinition(
                            rules=[], copy_ranges=data_ranges, structured=False,
                            structure_ranges=headers,
                        )
                        template_items.append(
                            PreflightItem(
                                "汇总结构", "提示", "通过", "", "",
                                f"将按 {len(headers)} 个表头区域核对汇总文件结构",
                            )
                        )
                    if external_path is not None:
                        external_workbook = excel.open_workbook(external_path, read_only=True)
                        external_plan = _external_sheet_plan(
                            excel, template_workbook, definition, external_workbook,
                        )
                        template_items.append(
                            excel.inspect_external_workbook(
                                external_workbook, external_plan.sheet_names,
                                plan_source=external_plan.source,
                            )
                        )
                except Exception as exc:
                    template_items = [
                        PreflightItem("总体", "错误", "不通过", "", "", str(exc))
                    ]
                    if feature_log is not None:
                        feature_log.add_template_health(template_path, template_items)
                    if write_report:
                        write_preflight_report_xlsx(
                            report_path, template_path, template_items
                        )
                    raise TemplateError(
                        f"模板体检不通过：{exc}"
                        + (f"。体检报告：{report_path}" if write_report else "")
                    ) from exc

                if feature_log is not None:
                    feature_log.add_template_health(template_path, template_items)
                if any(item.level == "错误" for item in template_items):
                    if write_report:
                        write_preflight_report_xlsx(
                            report_path, template_path, template_items
                        )
                    raise TemplateError(
                        "模板体检不通过，请先处理错误项"
                        + (f"。体检报告：{report_path}" if write_report else "")
                    )

                source_matches: list[tuple[Path, SourceMatch]] = []
                expected_structure = template_structure_values(template_workbook, definition)
                if on_step is not None:
                    on_step("正在执行：表结构比对")
                for source_path in source_files:
                    try:
                        match_result = validate_source_xlsx(
                            source_path, definition, expected_structure,
                            external_sheet_names=external_plan.sheet_names if external_plan else (),
                        )
                    except Exception as exc:
                        match_result = SourceMatch(
                            False, 0.0, 0, 0, (), f"无法完成结构匹配：{exc}"
                        )
                    source_matches.append((source_path, match_result))
                if on_step is not None:
                    on_step("完成：表结构比对")
                if feature_log is not None:
                    feature_log.add_sheet(
                        "表结构比对",
                        ("报送文件", "匹配结果", "匹配率", "匹配标签数", "检查标签数", "缺少工作表", "说明"),
                        [
                            (
                                str(source_path),
                                "通过" if result.matched else "不通过",
                                f"{result.score:.0%}",
                                result.matched_labels,
                                result.checked_labels,
                                "、".join(result.missing_sheets) or "无",
                                result.details,
                            )
                            for source_path, result in source_matches
                        ],
                    )
                if write_report:
                    write_preflight_report_xlsx(
                        report_path, template_path, template_items
                    )
            finally:
                if external_workbook is not None:
                    excel.close_workbook(external_workbook)
                excel.close_workbook(template_workbook)

        result = PreflightRunResult(
            report_path=report_path if write_report else None,
            total_files=len(source_matches),
            matched_files=sum(1 for _, item in source_matches if item.matched),
            template_warnings=sum(
                1 for item in template_items if item.level in {"提示", "警告"}
            ),
            source_matches=tuple(source_matches),
            batch_id=batch_id,
        )
        if on_step is not None:
            on_step("完成：表结构比对")
        return result

    def run(
        self,
        *,
        template_path: Path,
        input_dir: Path,
        output_dir: Path,
        period: str,
        history_path: Path,
        selected_files: list[Path] | None = None,
        external_path: Path | None = None,
        extra_files: list[Path] | None = None,
        flow_name: str | None = None,
        flow_instances: list[object] | None = None,
        recursive: bool = False,
        on_step: Optional[Callable[[str], None]] = None,
        feature_log: Optional[FeatureLog] = None,
        strict: bool = True,
        effective_sources: dict[int, str | None] | None = None,
    ) -> AuditRunResult:
        from .engines import pipeline_kind

        if pipeline_kind(self.engine_preference) == "native":
            from .native.audit_flow import run_native_audit

            return run_native_audit(
                template_path=template_path,
                input_dir=input_dir,
                output_dir=output_dir,
                config_path=self.config_path,
                history_path=history_path,
                period=period,
                external_path=external_path,
                selected_files=selected_files,
                recursive=recursive,
                on_step=on_step,
                write_summary=True,
                update_history=True,
            )
        template_path = template_path.resolve()
        input_dir = input_dir.resolve()
        output_dir = output_dir.resolve()
        history_path, history_sheet, legacy_history_path = self._history_storage(history_path)
        if external_path is not None:
            external_path = external_path.resolve()
        if not template_path.is_file():
            raise FileNotFoundError(f"模板不存在：{template_path}")
        if not input_dir.is_dir():
            raise FileNotFoundError(f"源数据目录不存在：{input_dir}")
        if external_path is not None and not external_path.is_file():
            raise FileNotFoundError(f"外部文件不存在：{external_path}")
        if not period.strip():
            raise ValueError("数据期不能为空")

        flow_steps = load_flow_steps(self.config_path, flow_name) if flow_name and self.config_path else ()
        if flow_name and not flow_steps:
            raise ValueError(f"执行流程“{flow_name}”不存在，或没有启用的功能")

        source_files = _source_files(
            input_dir, selected_files, recursive=recursive, extra_files=extra_files
        )
        if not source_files:
            raise ValueError("源数据目录中没有可审核的 .xlsx 文件")

        batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
        audit_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        output_dir.mkdir(parents=True, exist_ok=True)
        file_results: list[FileAuditResult] = []
        raw_issues = []
        result_sets: dict[str, list] = {}
        preflight_path = output_dir / f"审核前检查_{batch_id}.xlsx"
        service_started = time.monotonic()
        source_copy_seconds = 0.0
        open_seconds = 0.0
        extract_seconds = 0.0
        save_seconds = 0.0
        # 模板/配置读取本身也可能失败；异常处理仍应能产出可诊断的体检报告。
        # 成功读取流程后，会在下面按模块配置覆盖该默认值。
        keep_preflight_report = True
        # 审核副本不再单独保留：归“公式校验”输出，临时工作目录处理完即清理。
        audit_folder: Path | None = None
        audit_folder_scratch = False
        # “外部文件添加”的阶段副本目录，供后续“处理对象”指向它的功能使用。
        external_copies_dir: Path | None = None

        with ExcelSession(self.engine_preference) as excel:
            if on_step is not None:
                on_step(f"已连接表格引擎：{excel.engine_name}")
            history = excel.read_history(history_path, sheet_name=history_sheet)
            if not history and legacy_history_path and legacy_history_path.exists():
                history = excel.read_history(legacy_history_path)
            template_workbook = excel.open_workbook(template_path, read_only=True)
            external_workbook = None
            external_plan = None
            try:
                try:
                    definition = excel.read_template(
                        template_workbook, config_path=self.config_path, features=flow_instances
                    )
                    mappings = load_feature_mappings(self.config_path, template_path)
                    mapping_by_name = {item.name: item for item in mappings}
                    missing_features = [step.feature_name for step in flow_steps if step.feature_name not in mapping_by_name]
                    if missing_features:
                        raise ValueError("执行流程引用了“模块化功能”中不存在的功能：" + "、".join(missing_features))
                    selected_mappings = flow_instances if flow_instances is not None else ([mapping_by_name[step.feature_name] for step in flow_steps] if flow_steps else list(mappings))
                    active_types = {item.feature_type for item in selected_mappings}
                    step_mappings = [
                        (step, mapping_by_name[step.feature_name]) for step in flow_steps
                    ]

                    def emit_feature_steps(feature_type: str, phase: str) -> None:
                        if on_step is None:
                            return
                        for step, mapping in step_mappings:
                            if mapping.feature_type == feature_type:
                                on_step(f"{phase}：{step.feature_name}")
                    named_range_steps = [
                        (step, mapping) for step, mapping in step_mappings
                        if mapping.feature_type == NAMED_RANGE_CHECK_FUNCTION
                    ]
                    for step, mapping in named_range_steps:
                        if on_step is not None:
                            on_step(f"正在执行：{step.feature_name}")
                        # 先调查再强校验：区域缺失时运行日志仍保留“缺失”记录。
                        if feature_log is not None:
                            survey = excel.survey_named_ranges(template_workbook, mapping)
                            feature_log.add_sheet(
                                step.feature_name, ("工作表", "命名区域名", "覆盖区域", "结果"), survey
                            )
                            if on_step is not None:
                                found = [row for row in survey if row[3] == "通过"]
                                if found:
                                    per_sheet: dict[str, int] = {}
                                    for row in found:
                                        per_sheet[row[0]] = per_sheet.get(row[0], 0) + 1
                                    detail = "、".join(f"{sheet} {count} 处" for sheet, count in per_sheet.items())
                                    on_step(f"{step.feature_name}：{detail}")
                                else:
                                    on_step(f"{step.feature_name}：未找到 {mapping.range_names}")
                        excel.require_named_ranges(template_workbook, mapping)
                        if on_step is not None:
                            on_step(f"完成：{step.feature_name}")
                    issue_result_steps = [
                        (step, mapping) for step, mapping in step_mappings
                        if step.output_result and mapping.feature_type == AUDIT_RESULT_OUTPUT_FUNCTION
                    ]
                    keep_issue_result = not flow_steps or bool(issue_result_steps)
                    # 阶段快照只保留“外部文件添加”；“公式校验”的工作副本即最终审核副本。
                    checkpoint_steps = [
                        (step, mapping) for step, mapping in step_mappings
                        if step.output_result and mapping.feature_type == EXTERNAL_FILE_FUNCTION
                    ]
                    # 检查类与核对类的记录都并入运行日志；若检查步骤填写“输出结果=是”，
                    # 运行日志就是该流程可保留的输出。仍不再单独写“审核前检查”报告，
                    # 模板体检也进入运行日志。
                    check_report_keep = not flow_steps or any(
                        step.output_result and mapping.feature_type == NAMED_RANGE_CHECK_FUNCTION
                        for step, mapping in step_mappings
                    )
                    keep_preflight_report = check_report_keep
                    if feature_log is not None:
                        check_report_keep = False
                        keep_preflight_report = False
                    # 审核副本是“公式校验”的输出：该步选“是”时，工作副本直接建在
                    # “序号_公式校验_时间戳/”下，完成后即带审核导航表；选“否”时
                    # 工作副本放临时目录，处理完清理，不再单独保留机构审核副本。
                    formula_result_steps = [
                        (step, mapping) for step, mapping in step_mappings
                        if step.output_result and mapping.feature_type == FORMULA_COPY_FUNCTION
                    ]
                    if formula_result_steps:
                        audit_step, audit_mapping = formula_result_steps[-1]
                        audit_folder = output_dir / (
                            f"{_output_prefix(audit_step.order, audit_mapping.name, audit_step.output_name)}_{batch_id}"
                        )
                        audit_folder_scratch = False
                    else:
                        audit_folder = output_dir / "_审核工作副本"
                        audit_folder_scratch = True
                    audit_folder.mkdir(parents=True, exist_ok=True)
                    # 只有当“校验结果提取”解析后的处理对象是“公式校验复制”，
                    # 而流程里又没有公式校验复制时才报错/警告；解析后为源文件/外部文件添加
                    # 或留空时不要求公式复制（提取读源文件自带公式或外部添加副本）。
                    issue_step = next(
                        (step for step, _m in step_mappings
                         if _m.feature_type == ISSUE_EXTRACT_FUNCTION),
                        None,
                    )
                    if issue_step is not None:
                        issue_process_source = (
                            effective_sources.get(issue_step.order)
                            if effective_sources is not None
                            else _resolve_process_source(
                                flow_steps, issue_step,
                                {
                                    mapping.name for mapping in mapping_by_name.values()
                                    if mapping.feature_type in COPY_PRODUCING_FEATURE_TYPES
                                },
                            )[0]
                        ) or ""
                    else:
                        issue_process_source = ""
                    if issue_process_source == "公式校验复制" and FORMULA_COPY_FUNCTION not in active_types:
                        if strict:
                            raise ValueError("执行流程配置错误：“问题提取”前必须启用“公式复制”")
                        if on_step is not None:
                            on_step(
                                "警告：本流程未配置“公式校验复制”，"
                                "校验结果提取取决于源文件自带的校验公式"
                            )
                    type_order = [item.feature_type for item in selected_mappings]
                    if EXTERNAL_FILE_FUNCTION in type_order and FORMULA_COPY_FUNCTION in type_order and type_order.index(EXTERNAL_FILE_FUNCTION) > type_order.index(FORMULA_COPY_FUNCTION):
                        raise ValueError("执行流程配置错误：“外部文件添加”必须排在“公式复制”之前")
                    failure_policy = {
                        mapping.feature_type: step.on_failure.strip() or "停止"
                        for step, mapping in step_mappings
                    }
                    template_items = excel.inspect_template_health(
                        template_workbook, definition
                    )
                    if external_path is not None and EXTERNAL_FILE_FUNCTION in active_types:
                        external_workbook = excel.open_workbook(external_path, read_only=True)
                        external_plan = _external_sheet_plan(
                            excel, template_workbook, definition, external_workbook,
                        )
                        template_items.append(
                            excel.inspect_external_workbook(
                                external_workbook, external_plan.sheet_names,
                                plan_source=external_plan.source,
                            )
                        )
                    if feature_log is not None:
                        feature_log.add_template_health(template_path, template_items)
                except Exception as exc:
                    template_items = [
                        PreflightItem(
                            "总体",
                            "错误",
                            "不通过",
                            "",
                            "",
                            str(exc),
                        )
                    ]
                    if feature_log is not None:
                        feature_log.add_template_health(template_path, template_items)
                    if keep_preflight_report:
                        write_preflight_report_xlsx(
                            preflight_path, template_path, template_items
                        )
                    raise TemplateError(
                        f"模板体检不通过：{exc}"
                        + (f"。体检报告：{preflight_path}" if keep_preflight_report else "")
                    ) from exc

                if any(item.level == "错误" for item in template_items):
                    if keep_preflight_report:
                        write_preflight_report_xlsx(
                            preflight_path, template_path, template_items
                        )
                    raise TemplateError(
                        "模板体检不通过，请先处理错误项"
                        + (f"。体检报告：{preflight_path}" if keep_preflight_report else "")
                    )

                source_matches: list[tuple[Path, SourceMatch]] = []
                processable_files: list[Path] = []
                if STRUCTURE_COMPARE_FUNCTION in active_types:
                    emit_feature_steps(STRUCTURE_COMPARE_FUNCTION, "正在执行")
                    expected_structure = template_structure_values(template_workbook, definition)
                    for source_path in source_files:
                        org_code, org_name = _organisation_from_name(source_path)
                        try:
                            match_result = validate_source_xlsx(
                                source_path, definition, expected_structure,
                                external_sheet_names=external_plan.sheet_names if external_plan else (),
                            )
                        except Exception as exc:
                            match_result = SourceMatch(False, 0.0, 0, 0, (), f"无法完成结构匹配：{exc}")
                        source_matches.append((source_path, match_result))
                        if match_result.matched:
                            processable_files.append(source_path)
                        else:
                            file_results.append(
                                FileAuditResult(source_path=source_path, audit_path=None, org_code=org_code,
                                    org_name=org_name, error="报送文件与模板不匹配：" + match_result.details)
                            )
                    if on_step is not None:
                        matched_count = sum(1 for _, match in source_matches if match.matched)
                        on_step(
                            f"表结构比对：{matched_count} 个文件通过，"
                            f"{len(source_matches) - matched_count} 个跳过"
                        )
                    emit_feature_steps(STRUCTURE_COMPARE_FUNCTION, "完成")
                else:
                    # 未登记“表结构比对”即不打开或比对源文件；后续模块自行处理
                    # 它们真正需要的工作表/外部表冲突，不能再被当作隐含前置检查。
                    source_matches = [
                        (path, SourceMatch(True, 0.0, 0, 0, (), "流程未配置表结构比对，已跳过"))
                        for path in source_files
                    ]
                    processable_files = list(source_files)

                # 表结构不匹配是单个报送文件的问题：记录并跳过该文件，不能
                # 因一个机构的文件异常中止其余机构的审核或说明汇总。“停止”
                # 仅用于实际模块执行过程中的致命错误。
                if feature_log is not None and STRUCTURE_COMPARE_FUNCTION in active_types:
                    structure_feature_name = next(
                        (mapping.name for step, mapping in step_mappings
                         if mapping.feature_type == STRUCTURE_COMPARE_FUNCTION),
                        "表结构比对",
                    )
                    feature_log.add_sheet(
                        structure_feature_name,
                        ("报送文件", "匹配结果", "匹配率", "匹配标签数", "检查标签数", "缺少工作表", "说明"),
                        [
                            (
                                str(source_path),
                                "通过" if result.matched else "不通过",
                                f"{result.score:.0%}",
                                result.matched_labels,
                                result.checked_labels,
                                "、".join(result.missing_sheets) or "无",
                                result.details,
                            )
                            for source_path, result in source_matches
                        ],
                    )
                if check_report_keep:
                    write_preflight_report_xlsx(
                        preflight_path, template_path, template_items
                    )

                # —— 阶段化处理：先复制全部审核副本，再统一做外部文件添加、
                # 公式校验复制，最后统一做校验结果提取。每阶段只同时打开一个
                # 工作簿，处理完立即关闭，避免 Excel 并发打开多工作簿拖慢计算。
                external_log_rows: list[tuple[str, str, str]] = []
                formula_log_rows: list[tuple[str, str, int, str]] = []
                raw_issues = []

                records = []
                for source_path in processable_files:
                    org_code, org_name = _organisation_from_name(source_path)
                    records.append(
                        {
                            "source": source_path,
                            "audit": _available_output_path(audit_folder, source_path),
                            "org_code": org_code,
                            "org_name": org_name,
                            "stage_errors": [],
                            "issues": [],
                        }
                    )

                def stage_log_failure(record, feature_type: str, exc: Exception) -> None:
                    detail = f"{feature_type}：{exc}"
                    record["stage_errors"].append(detail)
                    if on_step is not None:
                        on_step(f"已跳过：{record['source'].name}｜{detail}")

                def abort_run(record, exc: Exception) -> RuntimeError:
                    try:
                        record["audit"].unlink(missing_ok=True)
                    except OSError:
                        pass
                    return RuntimeError(
                        f"执行流程“{flow_name}”处理“{record['source'].name}”失败，已按“停止”中止：{exc}"
                    )

                # 复制：把报送文件一次性复制为审核副本快照，供后续各阶段处理。
                for record in records:
                    source_hash = _sha256(record["source"])
                    started = time.monotonic()
                    shutil.copy2(record["source"], record["audit"])
                    source_copy_seconds += time.monotonic() - started
                    if _sha256(record["source"]) != source_hash:
                        raise abort_run(record, RuntimeError("原始报送文件在审核过程中发生变化"))

                # —— 阶段：外部文件添加（把外部工作表复制进审核副本）——
                if external_workbook is not None and EXTERNAL_FILE_FUNCTION in active_types:
                    emit_feature_steps(EXTERNAL_FILE_FUNCTION, "正在执行")
                    for record in records:
                        workbook = None
                        external_ok = True
                        try:
                            started = time.monotonic()
                            workbook = excel.open_workbook(record["audit"], read_only=False)
                            open_seconds += time.monotonic() - started
                            try:
                                excel.copy_external_sheets(
                                    external_workbook, workbook, external_plan.sheet_names
                                )
                                for step, mapping in checkpoint_steps:
                                    if mapping.feature_type != EXTERNAL_FILE_FUNCTION:
                                        continue
                                    checkpoint_dir = output_dir / (
                                        f"{_output_prefix(step.order, mapping.name, step.output_name)}_{batch_id}"
                                    )
                                    checkpoint_dir.mkdir(parents=True, exist_ok=True)
                                    external_copies_dir = checkpoint_dir
                                    workbook.Save()
                                    workbook.SaveCopyAs(str((checkpoint_dir / record["audit"].name).resolve()))
                                workbook.Save()
                            except Exception as exc:
                                if failure_policy.get(EXTERNAL_FILE_FUNCTION, "停止") != "跳过":
                                    raise
                                external_ok = False
                                stage_log_failure(record, EXTERNAL_FILE_FUNCTION, exc)
                        except Exception as exc:
                            raise abort_run(record, exc) from exc
                        finally:
                            if workbook is not None:
                                excel.close_workbook(workbook)
                        if feature_log is not None:
                            external_log_rows.append((
                                record["source"].name,
                                "、".join(external_plan.sheet_names) if external_plan else "",
                                "成功" if external_ok else "；".join(record["stage_errors"]),
                            ))
                        if on_step is not None:
                            sheets = "、".join(external_plan.sheet_names) if external_plan else "外部工作表"
                            on_step(f"为{record['source'].name}添加了{sheets}")
                    emit_feature_steps(EXTERNAL_FILE_FUNCTION, "完成")

                # —— 阶段：公式校验复制（复制模板公式并强制重算）——
                if FORMULA_COPY_FUNCTION in active_types:
                    emit_feature_steps(FORMULA_COPY_FUNCTION, "正在执行")
                    for record in records:
                        workbook = None
                        formula_ok = True
                        try:
                            started = time.monotonic()
                            workbook = excel.open_workbook(record["audit"], read_only=False)
                            open_seconds += time.monotonic() - started
                            try:
                                excel.apply_rules(template_workbook, workbook, definition)
                                workbook.Save()
                            except Exception as exc:
                                if failure_policy.get(FORMULA_COPY_FUNCTION, "停止") != "跳过":
                                    raise
                                formula_ok = False
                                stage_log_failure(record, FORMULA_COPY_FUNCTION, exc)
                        except Exception as exc:
                            raise abort_run(record, exc) from exc
                        finally:
                            if workbook is not None:
                                excel.close_workbook(workbook)
                        if feature_log is not None:
                            formula_log_rows.append((
                                record["source"].name,
                                str(record["audit"]),
                                sum(1 for rule in definition.rules if rule.enabled),
                                "成功" if formula_ok else "；".join(record["stage_errors"]),
                            ))
                    emit_feature_steps(FORMULA_COPY_FUNCTION, "完成")

                # —— 阶段：校验结果提取 + 收尾（写审核导航、保存、记录结果）——
                extract_active = ISSUE_EXTRACT_FUNCTION in active_types
                conditional_steps = [
                    (step, mapping) for step, mapping in step_mappings
                    if mapping.feature_type == CONDITIONAL_FORMAT_EXTRACT_FUNCTION
                ]
                if extract_active:
                    emit_feature_steps(ISSUE_EXTRACT_FUNCTION, "正在执行")
                for step, _mapping in conditional_steps:
                    if on_step is not None:
                        on_step(f"正在执行：{step.feature_name}")
                for record in records:
                    workbook = None
                    issues = []
                    conditional_issues = []
                    conditional_by_set: dict[str, list] = {}
                    extract_ok = True
                    try:
                        if conditional_steps:
                            source_workbook = None
                            try:
                                # 条件格式默认直接读取原始报送文件；若用户明确把“输入”
                                # 设为前序副本模块，则读取该审核副本。
                                conditional_input = next(
                                    (effective_sources.get(step.order) for step, _mapping in conditional_steps
                                     if effective_sources is not None and effective_sources.get(step.order)),
                                    None,
                                )
                                conditional_path = record["audit"] if conditional_input else record["source"]
                                started = time.monotonic()
                                source_workbook = excel.open_workbook(conditional_path, read_only=True)
                                open_seconds += time.monotonic() - started
                                for step, mapping in conditional_steps:
                                    try:
                                        started = time.monotonic()
                                        extracted = excel.extract_conditional_format_issues(
                                            source_workbook,
                                            mapping=mapping,
                                            structure_ranges=definition.structure_ranges,
                                            period=period.strip(), batch_id=batch_id,
                                            audit_time=audit_time, org_code=record["org_code"],
                                            org_name=record["org_name"], source_file=record["source"],
                                            workbook_path=conditional_path,
                                        )
                                        conditional_issues.extend(extracted)
                                        conditional_by_set.setdefault(step.result_set or "本期审核结果", []).extend(extracted)
                                        unsupported = getattr(excel, "last_conditional_unsupported_count", 0)
                                        if unsupported and on_step is not None:
                                            on_step(f"WPS 条件规则暂不支持：{unsupported} 条规则")
                                        extract_seconds += time.monotonic() - started
                                    except Exception as exc:
                                        if step.on_failure != "跳过":
                                            raise
                                        stage_log_failure(record, step.feature_name, exc)
                            finally:
                                if source_workbook is not None:
                                    excel.close_workbook(source_workbook)
                        started = time.monotonic()
                        workbook = excel.open_workbook(record["audit"], read_only=False)
                        open_seconds += time.monotonic() - started
                        if extract_active:
                            try:
                                # 将 Excel 的实时计算结果写入 xlsx 缓存后，由
                                # openpyxl 成批读取。这样既避免逐格 COM 调用，也
                                # 能可靠保留 #N/A、#REF! 等 Excel 错误值。
                                started = time.monotonic()
                                workbook.Save()
                                save_seconds += time.monotonic() - started
                                started = time.monotonic()
                                issues = excel.extract_issues(
                                    workbook, definition.rules, period=period.strip(), batch_id=batch_id,
                                    audit_time=audit_time, org_code=record["org_code"], org_name=record["org_name"],
                                    source_file=record["source"], audit_file=record["audit"],
                                    extraction_ranges=definition.extraction_ranges,
                                    saved_workbook_path=record["audit"],
                                )
                                extract_seconds += time.monotonic() - started
                            except Exception as exc:
                                if failure_policy.get(ISSUE_EXTRACT_FUNCTION, "停止") != "跳过":
                                    raise
                                extract_ok = False
                                stage_log_failure(record, ISSUE_EXTRACT_FUNCTION, exc)
                        formula_set = next(
                            (step.result_set or "本期审核结果" for step, mapping in step_mappings
                             if mapping.feature_type == ISSUE_EXTRACT_FUNCTION),
                            "本期审核结果",
                        )
                        result_sets.setdefault(formula_set, []).extend(issues)
                        for name, extracted in conditional_by_set.items():
                            result_sets.setdefault(name, []).extend(extracted)
                        issues.extend(conditional_issues)
                        record["issues"] = issues
                        # 历史说明只读带入后，在本次打开期间直接写审核导航，
                        # 避免每家审核副本保存后又重新打开一次。
                        merge_history(history, issues)
                        started = time.monotonic()
                        excel.write_navigation(workbook, issues)
                        workbook.Save()
                        save_seconds += time.monotonic() - started
                    except Exception as exc:
                        raise abort_run(record, exc) from exc
                    finally:
                        if workbook is not None:
                            excel.close_workbook(workbook)
                    if on_step is not None and extract_active and extract_ok:
                        on_step(f"提取问题：{record['source'].name}（{len(issues)} 条）")
                    raw_issues.extend(issues)
                    file_results.append(
                        FileAuditResult(
                            source_path=record["source"],
                            audit_path=record["audit"],
                            org_code=record["org_code"],
                            org_name=record["org_name"],
                            issues=issues,
                            error="；".join(record["stage_errors"]),
                        )
                    )
                if extract_active:
                    emit_feature_steps(ISSUE_EXTRACT_FUNCTION, "完成")
                for step, _mapping in conditional_steps:
                    if on_step is not None:
                        on_step(f"完成：{step.feature_name}")
                if feature_log is not None and external_log_rows:
                    external_feature_name = next(
                        (mapping.name for step, mapping in step_mappings
                         if mapping.feature_type == EXTERNAL_FILE_FUNCTION),
                        "外部文件添加",
                    )
                    feature_log.add_sheet(
                        external_feature_name, ("报送文件", "复制的外部工作表", "结果"), external_log_rows
                    )
                if feature_log is not None and formula_log_rows:
                    formula_feature_name = next(
                        (mapping.name for step, mapping in step_mappings
                         if mapping.feature_type == FORMULA_COPY_FUNCTION),
                        "公式校验复制",
                    )
                    feature_log.add_sheet(
                        formula_feature_name, ("报送文件", "审核副本", "复制公式数", "结果"), formula_log_rows
                    )
            finally:
                if external_workbook is not None:
                    excel.close_workbook(external_workbook)
                excel.close_workbook(template_workbook)

            # “公式校验”未要求输出时，工作副本目录只作临时用途，处理完即清理。
            if audit_folder_scratch and audit_folder is not None:
                shutil.rmtree(audit_folder, ignore_errors=True)

            # “历史审核记录”是简洁的规则说明库，不再维护新增/整改等生命周期台账。
            # 相同规则编号的人工说明和审核意见会自动带入本期审核结果。
            current = raw_issues
            resolved = []
            # merge_history 只把既有人工填写内容带入 current，不会改动历史表。
            merge_history(history, current)
            if feature_log is not None:
                issue_feature_name = (
                    next((mapping.name for step, mapping in step_mappings
                          if mapping.feature_type == ISSUE_EXTRACT_FUNCTION), "校验结果提取")
                    if step_mappings else "校验结果提取"
                )
                feature_log.add_sheet(
                    issue_feature_name,
                    ("报送文件", "工作表", "定位单元格", "级别", "校验字段", "问题说明", "当前值", "对比值", "差值", "规则编号"),
                    [
                        (
                            Path(item.source_file).name if item.source_file else "",
                            item.sheet_name,
                            item.target_cell,
                            item.severity,
                            item.check_field,
                            item.detail or item.message,
                            item.target_value,
                            item.comparison_value,
                            item.difference_value,
                            item.rule_id,
                        )
                        for item in current
                    ],
                )
                for _step, mapping in conditional_steps:
                    feature_log.add_sheet(
                        mapping.name,
                        ("报送文件", "工作表", "定位单元格", "错误类型", "校验指标", "描述", "当前值"),
                        [
                            (
                                Path(item.source_file).name if item.source_file else "",
                                item.sheet_name, item.target_cell, item.severity,
                                item.check_field, item.detail or item.message, item.target_value,
                            )
                            for item in current if item.rule_id == "条件格式填充"
                        ],
                    )
            summary_path = None
            if keep_issue_result:
                outputs = issue_result_steps or [(None, None)]
                for step, mapping in outputs:
                    if step is not None and mapping is not None:
                        result_set = step.result_set or "本期审核结果"
                        path = output_dir / f"{_output_prefix(step.order, mapping.name, step.output_name)}_{batch_id}.xlsx"
                    else:
                        result_set = "本期审核结果"
                        path = output_dir / f"基础数据审核结果_{period}_{batch_id}.xlsx"
                    excel.write_summary(path, result_sets.get(result_set, []), sheet_name=result_set)
                    summary_path = path

        total_seconds = time.monotonic() - service_started
        performance_lines = (
            f"计算引擎 {excel.engine_name or '未知'}",
            f"文件复制 {source_copy_seconds:.1f} 秒",
            f"打开审核副本 {open_seconds:.1f} 秒",
            f"外部表复制 {excel.external_copy_seconds:.1f} 秒",
            f"公式复制 {excel.copy_formula_seconds:.1f} 秒",
            f"Excel 计算 {excel.calculate_seconds:.1f} 秒",
            f"问题提取 {extract_seconds:.1f} 秒",
            f"写入导航并保存 {save_seconds:.1f} 秒",
            f"总计 {total_seconds:.1f} 秒",
        )

        copies: dict[str, Path] = {}
        if external_copies_dir is not None:
            copies["外部文件添加"] = external_copies_dir
        if not audit_folder_scratch and audit_folder is not None:
            copies["公式校验复制"] = audit_folder
        return AuditRunResult(
            batch_id=batch_id,
            period=period.strip(),
            output_dir=output_dir,
            current_issues=current,
            resolved_issues=resolved,
            files=file_results,
            summary_path=summary_path,
            history_path=None,
            preflight_path=preflight_path if preflight_path.exists() else None,
            performance_lines=performance_lines,
            copies=copies,
        )


class _NativeFlowResult:
    """native/flow.py 返回值到 web 层的适配：web 只消费 summary_text()。"""

    def __init__(self, outcome: dict) -> None:
        self._outcome = outcome

    def summary_text(self) -> str:
        parts = []
        output = self._outcome.get("output")
        if output:
            parts.append("流程输出：{}".format(output))
        log = self._outcome.get("log")
        if log:
            parts.append("运行日志：{}".format(log))
        return "\n".join(parts) if parts else "流程完成。"
