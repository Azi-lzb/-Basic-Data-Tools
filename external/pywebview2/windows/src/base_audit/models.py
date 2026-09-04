from __future__ import annotations

from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class AuditRule:
    rule_id: str
    enabled: bool
    report_code: str
    sheet_name: str
    formula_cell: str
    target_cell: str
    severity: str
    message: str
    copy_range: str = ""
    result_mode: str = "nonempty"
    value_from_result: bool = False


@dataclass(frozen=True)
class CopyRange:
    sheet_name: str
    address: str


@dataclass(frozen=True)
class TemplateDefinition:
    rules: list[AuditRule]
    copy_ranges: list[CopyRange]
    structured: bool
    structure_ranges: list[CopyRange] = field(default_factory=list)
    extraction_ranges: list[CopyRange] = field(default_factory=list)


@dataclass(frozen=True)
class PreflightItem:
    category: str
    level: str
    status: str
    sheet_name: str
    location: str
    message: str


@dataclass(frozen=True)
class StructureCheck:
    """One fixed template label compared with its source counterpart."""

    sheet_name: str
    cell_address: str
    expected_value: Any
    actual_value: Any
    matched: bool


@dataclass(frozen=True)
class SourceMatch:
    matched: bool
    score: float
    matched_labels: int
    checked_labels: int
    missing_sheets: tuple[str, ...] = ()
    details: str = ""
    structure_checks: tuple[StructureCheck, ...] = ()


@dataclass(frozen=True)
class PreflightRunResult:
    report_path: Path | None
    total_files: int
    matched_files: int
    template_warnings: int
    source_matches: tuple[tuple[Path, SourceMatch], ...] = ()
    batch_id: str | None = None
    log_path: Path | None = None

    @property
    def failed_files(self) -> int:
        return self.total_files - self.matched_files

    def summary_text(self) -> str:
        text = (
            f"审核前检查完成：通过 {self.matched_files} 个，不通过 {self.failed_files} 个，"
            f"模板提示 {self.template_warnings} 项。"
        )
        text += f"\n检查报告：{self.report_path}" if self.report_path else ""
        if self.log_path:
            text += f"\n运行日志：{self.log_path}"
        return text


@dataclass
class Issue:
    issue_id: str
    period: str
    batch_id: str
    audit_time: str
    triggered: bool
    status: str
    first_seen_period: str
    previous_seen_period: str
    consecutive_count: int
    org_code: str
    org_name: str
    report_code: str
    sheet_name: str
    rule_id: str
    severity: str
    formula_cell: str
    target_cell: str
    target_value: Any
    formula_result: Any
    message: str
    source_file: str
    audit_file: str
    check_field: str = ""
    comparison_value: Any = ""
    reference_value: Any = ""
    difference_value: Any = ""
    detail: str = ""
    institution_feedback: str = ""
    auditor_opinion: str = ""

    def clone(self, **changes: Any) -> "Issue":
        return replace(self, **changes)


@dataclass
class FileAuditResult:
    source_path: Path
    audit_path: Path | None
    org_code: str
    org_name: str
    issues: list[Issue] = field(default_factory=list)
    error: str = ""


@dataclass
class AuditRunResult:
    batch_id: str
    period: str
    output_dir: Path
    current_issues: list[Issue]
    resolved_issues: list[Issue]
    files: list[FileAuditResult]
    summary_path: Path | None = None
    history_path: Path | None = None
    preflight_path: Path | None = None
    performance_lines: tuple[str, ...] = ()
    log_path: Path | None = None
    # 各“副本产出功能”→副本目录，供后续“处理对象”指向该副本的功能使用。
    copies: dict[str, Path] = field(default_factory=dict)

    @property
    def successful_files(self) -> int:
        return sum(1 for item in self.files if not item.error)

    @property
    def failed_files(self) -> int:
        return sum(1 for item in self.files if item.error)

    def summary_text(self) -> str:
        parts = [
            f"审核完成：成功 {self.successful_files} 个文件，失败 {self.failed_files} 个文件。",
            f"本期审核结果 {len(self.current_issues)} 条。",
        ]
        if self.summary_path:
            parts.append(f"问题汇总：{self.summary_path}")
        if self.history_path:
            parts.append(f"历史库：{self.history_path}")
        if self.preflight_path:
            parts.append(f"审核前检查：{self.preflight_path}")
        if self.log_path:
            parts.append(f"运行日志：{self.log_path}")
        skipped = [item.source_path.name for item in self.files if item.error.startswith("报送文件与模板不匹配：")]
        if skipped:
            parts.append(f"已跳过 {len(skipped)} 个结构不匹配文件：" + "、".join(skipped))
        if self.performance_lines:
            parts.append("耗时分解：" + "；".join(self.performance_lines))
        return "\n".join(parts)
