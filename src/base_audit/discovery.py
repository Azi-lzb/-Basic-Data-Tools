from __future__ import annotations

import json
import re
import zipfile
from dataclasses import asdict, dataclass
from difflib import SequenceMatcher
from pathlib import Path
from xml.etree import ElementTree

from .template import normalize_template_name


SOURCE_SUFFIXES = {".xlsx"}
PERIOD_SEPARATED_RE = re.compile(
    r"(?<!\d)(20\d{2})[-._年/](0?[1-9]|1[0-2])(?:月|[-._/](?:0?[1-9]|[12]\d|3[01])日?)?(?!\d)"
)
PERIOD_COMPACT_RE = re.compile(
    r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])?(?!\d)"
)


@dataclass(frozen=True)
class PeriodDetection:
    period: str
    source: str
    conflict: bool = False
    details: str = ""


@dataclass(frozen=True)
class TemplateProfile:
    path: str
    size: int
    mtime_ns: int
    report_name: str
    sheet_names: tuple[str, ...]


@dataclass(frozen=True)
class TemplateSuggestion:
    template_path: Path | None
    confidence: float
    matched: bool
    details: str
    alternatives: tuple[tuple[str, float], ...] = ()


GENERATED_OUTPUT_FOLDERS = {"审核结果", "机构审核副本", "运行中间副本", "测试结果"}


def source_workbooks(input_dir: Path, *, recursive: bool = False) -> list[Path]:
    """List original source workbooks, optionally including institution subfolders."""
    if not input_dir.is_dir():
        return []
    iterator = input_dir.rglob("*.xlsx") if recursive else input_dir.glob("*.xlsx")
    return sorted(
        path
        for path in iterator
        if path.suffix.lower() in SOURCE_SUFFIXES
        and not path.name.startswith(("~$", "!", "！"))
        and "_审核版" not in path.stem
        and not any(part in GENERATED_OUTPUT_FOLDERS for part in path.relative_to(input_dir).parts[:-1])
    )


# 机构偶尔以 doc/docx/pdf 等非 xlsx 附件单独上传说明文件；这类文件不是审核源，
# 但需要在前端“机构说明文件”卡片提醒用户其存在。下面这些扩展名视为系统垃圾，
# 不进入提醒清单。
NON_XLSX_SKIPPED_SUFFIXES = {"", ".tmp", ".lnk", ".ini", ".db", ".log", ".json", ".cache"}


def explanation_files(input_dir: Path, *, recursive: bool = False) -> list[Path]:
    """List non-xlsx attachment/explanation files from the source tree.

    Institutions sometimes upload their explanation as a .doc/.docx/.pdf
    instead of an .xlsx workbook.  These are surfaced in a separate workbench
    card so the user is reminded they exist, without treating them as audit
    sources.  Generated output folders are skipped, matching ``source_workbooks``.
    """
    if not input_dir.is_dir():
        return []
    iterator = input_dir.rglob("*") if recursive else input_dir.glob("*")
    return sorted(
        path
        for path in iterator
        if path.is_file()
        and path.suffix.lower() not in SOURCE_SUFFIXES
        and path.suffix.lower() not in NON_XLSX_SKIPPED_SUFFIXES
        and not path.name.startswith(("~$", "!", "！", "."))
        and not any(part in GENERATED_OUTPUT_FOLDERS for part in path.relative_to(input_dir).parts[:-1])
    )


def _periods_in_text(value: str) -> set[str]:
    periods = {
        f"{match.group(1)}-{int(match.group(2)):02d}"
        for match in PERIOD_SEPARATED_RE.finditer(value)
    }
    periods.update(
        f"{match.group(1)}-{int(match.group(2)):02d}"
        for match in PERIOD_COMPACT_RE.finditer(value)
    )
    return periods


def detect_period(input_dir: Path, *, recursive: bool = False) -> PeriodDetection:
    files = source_workbooks(input_dir, recursive=recursive)
    filename_periods: set[str] = set()
    for path in files:
        filename_periods.update(_periods_in_text(path.stem))
    if len(filename_periods) == 1:
        period = next(iter(filename_periods))
        return PeriodDetection(period, "源文件名", details=f"从源文件名识别为 {period}")
    if len(filename_periods) > 1:
        values = "、".join(sorted(filename_periods))
        return PeriodDetection(
            "", "源文件名", True, f"源文件中识别出多个数据期：{values}"
        )

    # Only fall back to folder names when filenames contain no period at all.
    for folder in (input_dir, *input_dir.parents):
        periods = _periods_in_text(folder.name)
        if len(periods) == 1:
            period = next(iter(periods))
            return PeriodDetection(period, "文件夹名称", details=f"从文件夹名称识别为 {period}")
        if len(periods) > 1:
            values = "、".join(sorted(periods))
            return PeriodDetection(
                "", "文件夹名称", True, f"文件夹名称中识别出多个数据期：{values}"
            )
    return PeriodDetection("", "", details="未识别到数据期，请手动填写")


def read_xlsx_sheet_names(path: Path) -> tuple[str, ...]:
    try:
        with zipfile.ZipFile(path) as archive:
            xml = archive.read("xl/workbook.xml")
    except (OSError, KeyError, zipfile.BadZipFile):
        return ()
    root = ElementTree.fromstring(xml)
    # ElementTree's ``{*}tag`` wildcard is not supported by Python 3.7,
    # which is retained for the Win7 package.  Compare the local XML tag name
    # explicitly so Win7 and current Python versions read the same sheet list.
    return tuple(
        node.attrib["name"]
        for node in root.iter()
        if node.tag.rsplit("}", 1)[-1] == "sheet" and node.attrib.get("name")
    )


class TemplateCatalog:
    def __init__(self, template_dir: Path, index_path: Path) -> None:
        self.template_dir = template_dir
        self.index_path = index_path

    def profiles(self, *, force: bool = False) -> list[TemplateProfile]:
        cached = {} if force else self._load_index()
        profiles: list[TemplateProfile] = []
        for path in sorted(self.template_dir.glob("*.xlsx")):
            if not path.name.startswith(("!", "！")):
                continue
            stat = path.stat()
            key = str(path.resolve())
            item = cached.get(key)
            if (
                item
                and item.get("size") == stat.st_size
                and item.get("mtime_ns") == stat.st_mtime_ns
            ):
                profile = TemplateProfile(
                    path=key,
                    size=stat.st_size,
                    mtime_ns=stat.st_mtime_ns,
                    report_name=str(item.get("report_name") or ""),
                    sheet_names=tuple(item.get("sheet_names") or ()),
                )
            else:
                profile = TemplateProfile(
                    path=key,
                    size=stat.st_size,
                    mtime_ns=stat.st_mtime_ns,
                    report_name=normalize_template_name(path.stem),
                    sheet_names=read_xlsx_sheet_names(path),
                )
            profiles.append(profile)
        self._write_index(profiles)
        return profiles

    def _load_index(self) -> dict[str, dict[str, object]]:
        try:
            content = json.loads(self.index_path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return {}
        if not isinstance(content, dict) or content.get("template_dir") != str(
            self.template_dir.resolve()
        ):
            return {}
        items = content.get("templates", [])
        return {
            str(item.get("path")): item
            for item in items
            if isinstance(item, dict) and item.get("path")
        }

    def _write_index(self, profiles: list[TemplateProfile]) -> None:
        self.index_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": 2,
            "template_dir": str(self.template_dir.resolve()),
            "templates": [asdict(item) for item in profiles],
        }
        temporary = self.index_path.with_suffix(self.index_path.suffix + ".tmp")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporary.replace(self.index_path)


def _compact(value: str) -> str:
    value = PERIOD_SEPARATED_RE.sub("", value)
    value = PERIOD_COMPACT_RE.sub("", value)
    return "".join(character.lower() for character in value if character.isalnum())


def _filename_score(report_name: str, source_stem: str) -> float:
    target = _compact(report_name)
    if not target:
        return 0.0
    raw_parts = [part for part in re.split(r"[_\s]+", source_stem) if part]
    candidates = [source_stem, *raw_parts]
    candidates.extend("_".join(raw_parts[start:]) for start in range(len(raw_parts)))
    best = 0.0
    for value in candidates:
        candidate = _compact(value)
        if not candidate:
            continue
        if target in candidate or candidate in target:
            shorter = min(len(target), len(candidate))
            longer = max(len(target), len(candidate))
            if shorter >= 4:
                best = max(best, 0.92 + 0.08 * shorter / longer)
                continue
        best = max(best, SequenceMatcher(None, target, candidate).ratio())
    return min(best, 1.0)


def _sheet_score(template_sheets: tuple[str, ...], source_sheets: tuple[str, ...]) -> float:
    expected = {name for name in template_sheets if name != "审核规则"}
    actual = set(source_sheets)
    if not expected or not actual:
        return 0.0
    return len(expected & actual) / len(expected)


def _candidate_text(
    scores: list[tuple[float, TemplateProfile, float, float]],
) -> str:
    """Short, actionable candidate explanation for the workbench log."""
    return "；".join(
        f"{Path(profile.path).name}（文件名 {name_score:.0%}，工作表 {sheet_score:.0%}）"
        for _, profile, name_score, sheet_score in scores[:3]
    )


def recommend_template(
    template_dir: Path,
    input_dir: Path,
    index_path: Path,
    *,
    force_index: bool = False,
    recursive: bool = False,
) -> TemplateSuggestion:
    files = source_workbooks(input_dir, recursive=recursive)
    if not files:
        return TemplateSuggestion(None, 0.0, False, "源数据目录中没有可识别的 .xlsx 文件")
    profiles = TemplateCatalog(template_dir, index_path).profiles(force=force_index)
    if not profiles:
        return TemplateSuggestion(None, 0.0, False, "模板文件目录中没有正式模板")

    # Detect mixed report folders before calculating the aggregate recommendation.
    per_file_choices: set[str] = set()
    for source in files:
        source_sheets = read_xlsx_sheet_names(source)
        ranked = sorted(
            (
                (
                    0.65 * _filename_score(item.report_name, source.stem)
                    + 0.35 * _sheet_score(item.sheet_names, source_sheets),
                    item.path,
                )
                for item in profiles
            ),
            reverse=True,
        )
        if ranked and ranked[0][0] >= 0.82:
            if len(ranked) == 1 or ranked[0][0] - ranked[1][0] >= 0.04:
                per_file_choices.add(ranked[0][1])
    if len(per_file_choices) > 1:
        return TemplateSuggestion(
            None,
            0.0,
            False,
            "源数据目录可能包含多种报表，请按报表类型分目录后再审核",
        )

    source_sheets = read_xlsx_sheet_names(files[0])
    scores: list[tuple[float, TemplateProfile, float, float]] = []
    for profile in profiles:
        name_score = sum(
            _filename_score(profile.report_name, path.stem) for path in files
        ) / len(files)
        sheets_score = _sheet_score(profile.sheet_names, source_sheets)
        # Source filenames occasionally omit the report type, while business
        # sheet names are normally stable. Give sheet structure enough weight
        # to resolve those cases without trusting it alone.
        total = 0.65 * name_score + 0.35 * sheets_score
        scores.append((total, profile, name_score, sheets_score))
    scores.sort(key=lambda item: item[0], reverse=True)
    top_total, top, name_score, sheets_score = scores[0]
    runner_up = scores[1][0] if len(scores) > 1 else 0.0
    alternatives = tuple(
        (Path(profile.path).name, round(total, 3))
        for total, profile, _, _ in scores[:3]
    )
    structure_decisive = (
        sheets_score >= 0.95
        and (len(scores) == 1 or sheets_score - scores[1][3] >= 0.20)
    )
    unique_enough = top_total - runner_up >= 0.05 or structure_decisive
    matched = top_total >= 0.72 and unique_enough
    if matched:
        details = (
            f"自动匹配：{Path(top.path).name}（按源文件名称和工作表名称辅助判断）"
        )
        return TemplateSuggestion(
            Path(top.path), top_total, True, details, alternatives
        )
    if top_total >= 0.62 and not unique_enough:
        details = "有多个相近模板，已停止自动选择，请手动确认。候选：" + _candidate_text(scores)
    else:
        details = "未找到可信度足够的模板，请手动选择"
    return TemplateSuggestion(None, top_total, False, details, alternatives)


def classify_source_files(
    template_dir: Path, input_dir: Path, index_path: Path, *, recursive: bool = False
) -> list[dict[str, str]]:
    """Return a lightweight per-file template suggestion for the workbench."""
    files = source_workbooks(input_dir, recursive=recursive)
    profiles = TemplateCatalog(template_dir, index_path).profiles() if template_dir.is_dir() else []
    rows: list[dict[str, str]] = []
    for path in files:
        ranked: list[tuple[float, TemplateProfile]] = []
        source_sheets = read_xlsx_sheet_names(path)
        for profile in profiles:
            score = 0.85 * _filename_score(profile.report_name, path.stem) + 0.15 * _sheet_score(profile.sheet_names, source_sheets)
            ranked.append((score, profile))
        ranked.sort(key=lambda item: item[0], reverse=True)
        if ranked and ranked[0][0] >= 0.72:
            template_name = Path(ranked[0][1].path).name
        else:
            template_name = "未识别"
        rows.append({"path": str(path), "name": path.name, "template": template_name})
    return rows
