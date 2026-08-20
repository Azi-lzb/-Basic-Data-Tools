"""Merge source workbooks that belong to the same reporting institution.

The result is deliberately a standalone preparation artifact.  It does not
attempt to preserve or repair cross-workbook formulas: users create named
ranges on the merged workbook afterwards, then use the ordinary audit flow.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from .discovery import source_workbooks
from .excel_com import ExcelSession


_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


@dataclass(frozen=True)
class MergeOrgItem:
    organisation: str
    output_path: Path | None
    source_files: tuple[Path, ...]
    sheet_count: int
    status: str
    message: str = ""


@dataclass(frozen=True)
class MergeOrgResult:
    output_dir: Path
    items: tuple[MergeOrgItem, ...]
    total_files: int
    log_path: Path | None = None

    @property
    def successful_items(self) -> int:
        return sum(item.status == "成功" for item in self.items)

    @property
    def failed_items(self) -> int:
        return len(self.items) - self.successful_items

    def summary_text(self) -> str:
        text = (
            f"合并同机构多表完成：成功 {self.successful_items} 家机构，"
            f"失败 {self.failed_items} 家机构，来源文件 {self.total_files} 个。\n"
            f"合并结果目录：{self.output_dir}"
        )
        if self.log_path:
            text += f"\n运行日志：{self.log_path}"
        return text


def _organisation_from_source(path: Path) -> str:
    """Institution key: the first underscore-delimited filename segment."""
    return path.stem.split("_", 1)[0].strip() or path.stem


def _report_type_from_source(path: Path) -> str:
    parts = [part.strip() for part in path.stem.split("_")]
    return parts[1] if len(parts) > 1 and parts[1] else "报表"


def _source_period(path: Path) -> str:
    parts = [part.strip() for part in path.stem.split("_")]
    return parts[3] if len(parts) > 3 else ""


def _safe_sheet_name(value: str, existing: set[str]) -> str:
    """Create a valid unique Excel sheet title, retaining a numeric suffix."""
    base = _INVALID_SHEET_CHARS.sub("_", value).strip("'") or "工作表"
    base = base[:31]
    candidate = base
    index = 2
    while candidate.casefold() in existing:
        suffix = f"_{index}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1
    existing.add(candidate.casefold())
    return candidate


def _selected_sources(input_dir: Path, selected_files: list[Path] | None, recursive: bool) -> list[Path]:
    sources = source_workbooks(input_dir, recursive=recursive)
    if not selected_files:
        return sources
    allowed = {path.resolve() for path in sources}
    requested = {Path(path).resolve() for path in selected_files}
    invalid = requested - allowed
    if invalid:
        raise ValueError("选择的待合并文件不在源数据目录中：" + str(sorted(invalid)[0]))
    return [path for path in sources if path.resolve() in requested]


def _output_folder(input_dir: Path, output_dir: Path, output_name: str) -> Path:
    """Avoid placing recursive inputs in a folder that will be scanned again."""
    input_resolved = input_dir.resolve()
    output_resolved = output_dir.resolve()
    if output_resolved == input_resolved:
        raise ValueError("输出目录不能与源数据目录相同；请使用源数据目录下的“审核结果”或其他独立目录")
    try:
        relative = output_resolved.relative_to(input_resolved)
    except ValueError:
        relative = None
    if relative is not None and "审核结果" not in relative.parts:
        raise ValueError("输出目录位于源数据目录内但不在“审核结果”目录下，递归时可能重复合并；请调整输出目录")
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    return output_resolved / f"{output_name}_{batch_id}"


def run_merge_org(
    *,
    input_dir: Path,
    output_dir: Path,
    period: str = "",
    selected_files: list[Path] | None = None,
    flow_name: str | None = None,
    recursive: bool = True,
    on_step: Optional[Callable[[str], None]] = None,
    feature_log: object | None = None,
    output_name: str | None = None,
) -> MergeOrgResult:
    """Merge all selected workbooks by their filename's first segment."""
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()
    if not input_dir.is_dir():
        raise FileNotFoundError(f"源数据目录不存在：{input_dir}")
    sources = _selected_sources(input_dir, selected_files, recursive)
    if not sources:
        raise ValueError("源数据目录中没有可合并的 .xlsx 文件")

    output_prefix = output_name or "合并同机构多表"
    target_dir = _output_folder(input_dir, output_dir, output_prefix)
    target_dir.mkdir(parents=True, exist_ok=False)
    groups: dict[str, list[Path]] = {}
    for path in sources:
        groups.setdefault(_organisation_from_source(path), []).append(path)

    if on_step is not None:
        on_step("提示：机构识别使用文件名下划线第一段；文件名第一段必须是机构全称")
    items: list[MergeOrgItem] = []
    with ExcelSession() as excel:
        if on_step is not None:
            on_step(f"已连接表格引擎：{excel.engine_name}")
        for organisation, files in sorted(groups.items()):
            dates = sorted({value for value in (_source_period(path) for path in files) if value})
            if len(dates) > 1 and on_step is not None:
                on_step(f"提示：机构“{organisation}”包含多个数据期（{'、'.join(dates)}），请确认")
            output_period = period.strip() or (dates[0] if dates else "未识别日期")
            out_path = target_dir / f"{organisation}_合并_{output_period}.xlsx"
            merged = None
            source_sheet_count = 0
            try:
                if on_step is not None:
                    on_step(f"正在合并：{organisation}（{len(files)} 个文件）")
                merged = excel.excel.Workbooks.Add()
                # Excel's user setting may create several blank sheets.  Keep
                # exactly one placeholder while copying, then delete it before
                # SaveAs so the merged workbook contains source sheets only.
                while merged.Worksheets.Count > 1:
                    merged.Worksheets(merged.Worksheets.Count).Delete()
                placeholder = merged.Worksheets(1)
                existing: set[str] = set()
                for source_path in sorted(files):
                    source_book = None
                    try:
                        source_book = excel.open_workbook(source_path, read_only=True)
                        report_type = _report_type_from_source(source_path)
                        for index in range(1, source_book.Worksheets.Count + 1):
                            source_sheet = source_book.Worksheets(index)
                            source_sheet.Copy(None, merged.Worksheets(merged.Worksheets.Count))
                            copied = merged.Worksheets(merged.Worksheets.Count)
                            copied.Name = _safe_sheet_name(
                                f"{report_type}_{source_sheet.Name}", existing
                            )
                            source_sheet_count += 1
                    finally:
                        if source_book is not None:
                            excel.close_workbook(source_book)
                placeholder.Delete()
                merged.SaveAs(str(out_path), FileFormat=51)
                items.append(MergeOrgItem(
                    organisation, out_path, tuple(files), source_sheet_count, "成功",
                    "命名区域需在合并文件上重建；跨表外部引用仍可能指向原文件，请手动更新",
                ))
                if on_step is not None:
                    on_step(f"完成合并：{organisation}（{source_sheet_count} 张工作表）")
            except Exception as exc:
                items.append(MergeOrgItem(
                    organisation, None, tuple(files), source_sheet_count, "失败", str(exc)
                ))
                if on_step is not None:
                    on_step(f"合并失败：{organisation}（{exc}），已继续其他机构")
            finally:
                if merged is not None:
                    excel.close_workbook(merged, save=False)

    if feature_log is not None:
        feature_log.add_sheet(
            "合并同机构多表",
            ("机构", "输出文件", "来源文件数", "来源工作表数", "结果"),
            [
                (item.organisation, str(item.output_path or ""), len(item.source_files), item.sheet_count,
                 item.status if not item.message else f"{item.status}：{item.message}")
                for item in items
            ],
        )
    return MergeOrgResult(target_dir, tuple(items), len(sources))
