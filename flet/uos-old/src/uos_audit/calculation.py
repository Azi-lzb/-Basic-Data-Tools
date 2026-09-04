"""Formula-calculation adapters for the native UOS edition.

The workbook workflow is intentionally independent from the calculation
runtime: openpyxl performs all copying; this module calculates the prepared
``*_审核版.xlsx``. Configured engines are always single-engine: LibreOffice,
Excel COM, WPS COM, IronCalc, or formulas.
"""
from __future__ import annotations

import re
import os
import shutil
import sys
import tempfile
import time
import zipfile
from multiprocessing import get_context
from queue import Empty
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Protocol
from xml.etree import ElementTree

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries, get_column_letter

from .libreoffice import LibreOfficeCalculator
from .models import Issue
from .template import normalize_template_name


ENGINE_COMBINED = "formulas + IronCalc"
ENGINE_FORMULAS = "formulas"
ENGINE_IRONCALC = "IronCalc"
ENGINE_LIBREOFFICE = "LibreOffice Calc"
ENGINE_EXCEL_COM = "Excel COM（测试）"
ENGINE_WPS_COM = "WPS COM（测试）"
# The UI deliberately exposes individual engines only.  Cross-validation was
# useful while developing adapters, but makes a routine audit wait for the
# slowest engine and obscures which engine actually produced a result.
ENGINE_OPTIONS = (ENGINE_LIBREOFFICE,)
ENGINE_PRIORITY = ENGINE_OPTIONS
ENGINE_DEFAULT = ENGINE_LIBREOFFICE
# Historical compatibility alias; UOS formal execution is LibreOffice only.
ENGINE_PYTHON = ENGINE_DEFAULT

# These are deliberately conservative. An unlisted function is not silently
# accepted by the formulas adapter.
PYTHON_COMMON_FUNCTIONS = frozenset({
    "ABS", "ADDRESS", "AND", "AVERAGE", "COLUMN", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
    "IF", "IFERROR", "INDEX", "ISBLANK", "ISERROR", "ISNUMBER", "LEFT", "LEN",
    "LOWER", "MATCH", "MAX", "MID", "MIN", "NOT", "OR", "RIGHT", "ROUND",
    "ROW", "SUM", "SUMIF", "SUMIFS", "TEXT", "TRIM", "UPPER", "VLOOKUP",
})
PYTHON_BLOCKED_FUNCTIONS = frozenset({
    "INDIRECT", "OFFSET", "CELL", "INFO", "RAND", "RANDBETWEEN", "NOW",
    "TODAY", "FILTER", "SORT", "UNIQUE", "SEQUENCE", "LET", "LAMBDA",
    "XLOOKUP", "XMATCH",
})
_FUNCTION = re.compile(r"(?<![A-Z0-9_.])([A-Z][A-Z0-9.]*)\s*\(", re.IGNORECASE)
_RELATIONSHIP_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_PYWIN32_DLL_DIRECTORIES: list[object] = []


def _prepare_pywin32_runtime() -> None:
    """Make pywin32 importable from a Flet embedded-Python distribution.

    Flet's native launcher adds ``site-packages`` to ``sys.path`` directly but
    does not always process its ``.pth`` files.  pywin32 relies on that step to
    add ``win32`` and register ``pywin32_system32`` as a DLL directory.
    Ordinary venv Python has already done this, so this is harmless there.
    """
    if os.name != "nt":
        return

    roots: list[Path] = []
    for item in sys.path:
        if item:
            roots.append(Path(item))
    # In a packaged app the module lives in <bundle>/app/src/uos_audit while
    # pywin32 is in the adjacent <bundle>/site-packages directory.
    try:
        module_path = Path(__file__).resolve()
        roots.extend(parent / "site-packages" for parent in module_path.parents)
    except OSError:
        pass

    seen: set[str] = set()
    for root in roots:
        try:
            root = root.resolve()
        except OSError:
            continue
        key = os.path.normcase(str(root))
        if key in seen:
            continue
        seen.add(key)
        system32 = root / "pywin32_system32"
        if not system32.is_dir():
            continue

        for child in (root / "win32", root / "win32" / "lib", root / "pythonwin"):
            child_text = str(child)
            if child.is_dir() and child_text not in sys.path:
                sys.path.append(child_text)
        try:
            handle = os.add_dll_directory(str(system32))
            _PYWIN32_DLL_DIRECTORIES.append(handle)
        except (AttributeError, OSError):
            # Older Python or a normal site-packages setup may already expose
            # the DLLs through PATH; let the subsequent import report details.
            pass
        return


def _workbook_sheet_parts(path: Path) -> dict[str, str]:
    """Resolve worksheet display names to ZIP part paths without openpyxl save."""
    with zipfile.ZipFile(path) as archive:
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.attrib.get("Id", ""): item.attrib.get("Target", "")
        for item in rels
        if item.tag.rsplit("}", 1)[-1] == "Relationship"
    }
    parts: dict[str, str] = {}
    for sheet in workbook.iter():
        if sheet.tag.rsplit("}", 1)[-1] != "sheet":
            continue
        name, rel_id = sheet.attrib.get("name", ""), sheet.attrib.get(_RELATIONSHIP_ID, "")
        target = targets.get(rel_id, "")
        if name and target:
            parts[name] = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")
    return parts


def _children(cell, local_name: str):
    return [child for child in cell if child.tag.rsplit("}", 1)[-1] == local_name]


def _patch_formula_caches(original_xml: bytes, calculated_xml: bytes) -> bytes:
    """Keep original worksheet presentation and replace only formula caches."""
    original = ElementTree.fromstring(original_xml)
    calculated = ElementTree.fromstring(calculated_xml)
    calculated_cells = {
        cell.attrib.get("r", ""): cell
        for cell in calculated.iter()
        if cell.tag.rsplit("}", 1)[-1] == "c" and cell.attrib.get("r")
    }
    for cell in original.iter():
        if cell.tag.rsplit("}", 1)[-1] != "c" or not _children(cell, "f"):
            continue
        calculated_cell = calculated_cells.get(cell.attrib.get("r", ""))
        if calculated_cell is None:
            continue
        values = _children(calculated_cell, "v")
        inline_strings = _children(calculated_cell, "is")
        if not values and not inline_strings:
            continue
        for prior in _children(cell, "v"):
            cell.remove(prior)
        for prior in _children(cell, "is"):
            cell.remove(prior)
        if values:
            cell.append(deepcopy(values[-1]))
        else:
            # formulas emits a calculated text formula as an inline string
            # value cell (without an <f>). A formula cache must instead be
            # stored as ``t=str`` + ``<v>text</v>`` beside the original <f>.
            namespace = cell.tag.rsplit("}", 1)[0] + "}" if "}" in cell.tag else ""
            cached = ElementTree.Element(namespace + "v")
            cached.text = "".join(inline_strings[-1].itertext())
            cell.append(cached)
        # Keep the original style attribute, but take the cache type so
        # openpyxl/Calc interpret string and boolean formula results properly.
        if inline_strings:
            cell.attrib["t"] = "str"
        elif "t" in calculated_cell.attrib:
            cell.attrib["t"] = calculated_cell.attrib["t"]
        else:
            cell.attrib.pop("t", None)
    return ElementTree.tostring(original, encoding="utf-8", xml_declaration=True)


def merge_formula_caches_preserving_workbook(
    original_path: Path, calculated_path: Path, destination_path: Path
) -> None:
    """Keep every original OOXML part and patch formula ``<v>`` cache nodes.

    The formulas library emits a simplified XLSX. Directly replacing the audit
    copy would lose styles, comments, defined names and conditional formatting.
    This ZIP-level merge avoids an openpyxl re-save, which itself clears cached
    formula values.
    """
    original_parts = _workbook_sheet_parts(original_path)
    calculated_parts = _workbook_sheet_parts(calculated_path)
    source_to_calculated = {
        original_parts[name]: calculated_parts[name]
        for name in original_parts.keys() & calculated_parts.keys()
    }
    with zipfile.ZipFile(original_path) as original, zipfile.ZipFile(calculated_path) as calculated:
        calculated_names = set(calculated.namelist())
        with zipfile.ZipFile(destination_path, "w", compression=zipfile.ZIP_DEFLATED) as merged:
            for info in original.infolist():
                data = original.read(info.filename)
                calculated_part = source_to_calculated.get(info.filename)
                if calculated_part and calculated_part in calculated_names:
                    data = _patch_formula_caches(data, calculated.read(calculated_part))
                merged.writestr(info, data)


class FormulaEngineError(RuntimeError):
    """A calculation engine cannot safely calculate this workbook."""


def _ironcalc_worker(workbook_name: str, calculated_name: str, result_queue: object) -> None:
    """Run IronCalc outside the UI process so a stuck native load can be stopped.

    Some Excel-authored packages make IronCalc's XLSX reader spend a long time
    before returning a ZIP-level error.  A thread cannot safely interrupt that
    native work, while a dedicated process can be terminated without taking
    down Flet or the rest of the batch.
    """
    try:
        import ironcalc as ic

        model = ic.load_from_xlsx(workbook_name, "en", "UTC", "en")
        model.evaluate()
        model.save_to_xlsx(calculated_name)
        result_queue.put(("ok", ""))
    except Exception as exc:
        result_queue.put(("error", "{}".format(exc)))


def _run_ironcalc_with_timeout(workbook_path: Path, calculated_path: Path, timeout_seconds: int) -> None:
    """Calculate one workbook with a bounded IronCalc child process."""
    context = get_context("spawn")
    result_queue = context.Queue()
    process = context.Process(
        target=_ironcalc_worker,
        args=(str(workbook_path), str(calculated_path), result_queue),
        daemon=True,
    )
    process.start()
    process.join(timeout_seconds)
    if process.is_alive():
        process.terminate()
        process.join()
        raise FormulaEngineError(
            "IronCalc 在 {} 秒内未完成。该工作簿可能包含 IronCalc 不兼容的 XLSX 组件；"
            "请在设置中心改用 formulas，或安装 LibreOffice Calc 后使用 LibreOffice 计算。".format(timeout_seconds)
        )
    try:
        status, detail = result_queue.get(timeout=2)
    except Empty:
        status, detail = "error", "IronCalc 子进程异常退出（退出码 {}）".format(process.exitcode)
    finally:
        result_queue.close()
        result_queue.join_thread()
    if status != "ok":
        raise FormulaEngineError(
            "IronCalc 无法读取此 XLSX 审核副本：{}。请改用 formulas，或使用 LibreOffice Calc 对模板基准验证。".format(detail)
        )


@dataclass(frozen=True)
class FormulaCapability:
    formula_count: int
    functions: tuple[str, ...]
    unsupported_functions: tuple[str, ...]
    blockers: tuple[str, ...]
    formula_errors: tuple[tuple[str, str, str], ...] = ()

    @property
    def python_supported(self) -> bool:
        return not self.unsupported_functions and not self.blockers

    def summary(self) -> str:
        parts = ["发现 {} 个公式".format(self.formula_count)]
        if self.functions:
            parts.append("函数：{}".format("、".join(self.functions)))
        if self.unsupported_functions:
            parts.append("未验证函数：{}".format("、".join(self.unsupported_functions)))
        if self.blockers:
            parts.append("兼容性限制：{}".format("；".join(self.blockers)))
        if self.formula_errors:
            parts.append("错误公式：{}".format(len(self.formula_errors)))
        return "；".join(parts)


@dataclass(frozen=True)
class CalculationResult:
    workbook_path: Path
    engine_name: str
    elapsed_seconds: float
    capability: FormulaCapability
    fallback_used: bool = False
    # 只扫描本次校验区域。这里记录的是计算后的缓存错误值，不把普通的
    # 业务提示文字误当成公式错误。
    error_cells: tuple[tuple[str, str, str], ...] = ()
    # 结果仍可用于提取，但不应被界面误认为已经完成双引擎基准验证。
    diagnostic: str = ""


class Calculator(Protocol):
    engine_name: str

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult: ...


def _formula_cache_values(
    formula_path: Path, calculated_path: Path, *, formula_ranges: Iterable[object] = ()
) -> dict[tuple[str, str], object]:
    """Read cached values only for formulas belonging to this audit step."""
    scoped = list(formula_ranges)
    formulas_book = load_workbook(formula_path, read_only=True, data_only=False, keep_links=False)
    values_book = load_workbook(calculated_path, read_only=True, data_only=True, keep_links=False)
    try:
        result: dict[tuple[str, str], object] = {}
        if scoped:
            candidates = []
            for item in scoped:
                if item.sheet_name not in formulas_book.sheetnames:
                    continue
                min_col, min_row, max_col, max_row = range_boundaries(item.address)
                candidates.append((item.sheet_name, formulas_book[item.sheet_name].iter_rows(
                    min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
                )))
        else:
            candidates = [(sheet.title, sheet.iter_rows()) for sheet in formulas_book.worksheets]
        for sheet_name, rows in candidates:
            values_sheet = values_book[sheet_name]
            for row in rows:
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        result[(sheet_name, cell.coordinate)] = values_sheet[cell.coordinate].value
        return result
    finally:
        values_book.close()
        formulas_book.close()


def inspect_formula_capability(workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> FormulaCapability:
    """Read validation formulas once and report Python-engine compatibility.

    When ranges are supplied they are expected to expose ``sheet_name`` and
    ``address`` (the native ``CopyRange`` contract). This keeps automatic mode
    focused on formulas copied by the current audit flow rather than unrelated
    helper formulas elsewhere in the workbook.
    """
    scoped = list(formula_ranges)
    book = load_workbook(workbook_path, read_only=True, data_only=False, keep_links=False)
    try:
        formulas: list[str] = []
        formula_cells: list[tuple[str, str, str]] = []
        candidates = []
        if scoped:
            for item in scoped:
                if item.sheet_name not in book.sheetnames:
                    continue
                min_col, min_row, max_col, max_row = range_boundaries(item.address)
                candidates.append(book[item.sheet_name].iter_rows(
                    min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
                ))
        else:
            candidates = [sheet.iter_rows() for sheet in book.worksheets]
        for rows in candidates:
            for row in rows:
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formulas.append(cell.value)
                        formula_cells.append((str(cell.parent.title), cell.coordinate, cell.value))
    finally:
        book.close()
    functions = sorted({name.upper() for formula in formulas for name in _FUNCTION.findall(formula)})
    unsupported = sorted(set(functions) - PYTHON_COMMON_FUNCTIONS)
    blockers: list[str] = []
    joined = "\n".join(formulas).upper()
    blocked = sorted(set(functions) & PYTHON_BLOCKED_FUNCTIONS)
    if blocked:
        blockers.append("包含不稳定函数：{}".format("、".join(blocked)))
    if "[" in joined and "]" in joined:
        blockers.append("包含外部工作簿引用")
    if "{" in joined or "#" in joined:
        blockers.append("可能包含数组或动态数组公式")
    error_tokens = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")
    formula_errors = tuple(
        (sheet_name, address, next(token for token in error_tokens if token in formula.upper()))
        for sheet_name, address, formula in formula_cells
        if any(token in formula.upper() for token in error_tokens)
    )
    return FormulaCapability(
        len(formulas), tuple(functions), tuple(unsupported), tuple(blockers), formula_errors,
    )


def scan_formula_errors(workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> tuple[tuple[str, str, str], ...]:
    """Return cached Excel error values in the copied validation formula areas."""
    scoped = list(formula_ranges)
    if not scoped:
        return ()
    book = load_workbook(workbook_path, read_only=True, data_only=True, keep_links=False)
    try:
        errors: list[tuple[str, str, str]] = []
        for item in scoped:
            if item.sheet_name not in book.sheetnames:
                continue
            min_col, min_row, max_col, max_row = range_boundaries(item.address)
            for row in book[item.sheet_name].iter_rows(
                min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
            ):
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.upper() in {
                        "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!",
                    }:
                        errors.append((item.sheet_name, cell.coordinate, value.upper()))
        return tuple(errors)
    finally:
        book.close()


class PythonFormulaCalculator:
    """Calculate a simple xlsx with the ``formulas`` package and save caches."""

    engine_name = ENGINE_FORMULAS

    def __init__(self, *, timeout_seconds: int = 180) -> None:
        self.timeout_seconds = timeout_seconds

    @staticmethod
    def require_available() -> None:
        try:
            import formulas  # noqa: F401
        except ImportError as exc:
            raise FormulaEngineError("未安装 formulas 引擎。请安装 requirements.txt，或在设置中显式选择 LibreOffice Calc。") from exc

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        self.require_available()
        workbook_path = workbook_path.resolve()
        if workbook_path.suffix.casefold() != ".xlsx":
            raise FormulaEngineError("formulas 引擎仅支持 .xlsx 审核副本")
        capability = inspect_formula_capability(workbook_path, formula_ranges=formula_ranges)
        # Compatibility inspection is diagnostic only.  Audit must still
        # attempt calculation so #REF!/#NAME? and other error caches can be
        # extracted as problems instead of blocking the whole batch.
        import formulas

        started = time.monotonic()
        try:
            model = formulas.ExcelModel().loads(str(workbook_path)).finish()
            model.calculate()
            with tempfile.TemporaryDirectory(prefix="base_audit_formulas_") as text:
                temp_dir = Path(text)
                output_dir = temp_dir / "output"
                output_dir.mkdir()
                model.write(dirpath=str(output_dir))
                generated = next(output_dir.rglob("*.xlsx"), None)
                if generated is None:
                    raise FormulaEngineError("formulas 未生成计算后的工作簿")
                # Formulas writes a simplified workbook. Preserve the audit
                # copy's styles, comments, names and conditional formatting,
                # then merge back only calculated formula caches.
                merged = temp_dir / "审核副本_已计算.xlsx"
                merge_formula_caches_preserving_workbook(workbook_path, generated, merged)
                shutil.copy2(merged, workbook_path)
        except FormulaEngineError:
            raise
        except Exception as exc:
            raise FormulaEngineError("formulas 计算失败：{}".format(exc)) from exc
        return CalculationResult(
            workbook_path, self.engine_name, time.monotonic() - started, capability,
            error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
        )


class IronCalcCalculator:
    """Calculate via the IronCalc Python binding and merge only formula caches."""

    engine_name = ENGINE_IRONCALC

    def __init__(self, *, timeout_seconds: int = 60) -> None:
        self.timeout_seconds = timeout_seconds

    @staticmethod
    def require_available() -> None:
        try:
            import ironcalc  # noqa: F401
        except ImportError as exc:
            raise FormulaEngineError("未安装 IronCalc 引擎。请安装 requirements.txt，或在设置中选择其他计算引擎。") from exc

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        self.require_available()
        workbook_path = workbook_path.resolve()
        if workbook_path.suffix.casefold() != ".xlsx":
            raise FormulaEngineError("IronCalc 引擎仅支持 .xlsx 审核副本")
        capability = inspect_formula_capability(workbook_path, formula_ranges=formula_ranges)
        started = time.monotonic()
        try:
            with tempfile.TemporaryDirectory(prefix="base_audit_ironcalc_") as text:
                temp_dir = Path(text)
                calculated = temp_dir / "审核副本_ironcalc.xlsx"
                _run_ironcalc_with_timeout(workbook_path, calculated, self.timeout_seconds)
                if not calculated.is_file():
                    raise FormulaEngineError("IronCalc 未生成计算后的工作簿")
                merged = temp_dir / "审核副本_已计算.xlsx"
                merge_formula_caches_preserving_workbook(workbook_path, calculated, merged)
                shutil.copy2(merged, workbook_path)
        except FormulaEngineError:
            raise
        except Exception as exc:
            raise FormulaEngineError("IronCalc 计算失败：{}".format(exc)) from exc
        return CalculationResult(
            workbook_path, self.engine_name, time.monotonic() - started, capability,
            error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
        )


class LibreOfficeAdapter:
    engine_name = ENGINE_LIBREOFFICE

    def __init__(self) -> None:
        self._calculator = LibreOfficeCalculator()

    def available(self) -> tuple[bool, str]:
        try:
            engine = self._calculator.require_available()
            return True, "{}（{}；{}）".format(engine.display, engine.source, self._calculator.version(engine))
        except Exception as exc:
            return False, str(exc)

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        capability = inspect_formula_capability(workbook_path, formula_ranges=formula_ranges)
        result = self._calculator.recalculate(workbook_path)
        return CalculationResult(
            result.workbook_path, self.engine_name, result.elapsed_seconds, capability,
            error_cells=scan_formula_errors(result.workbook_path, formula_ranges=formula_ranges),
        )


class ExcelComCalculator:
    """Recalculate with a private Microsoft Excel instance for Windows testing.

    This adapter is intentionally not a Linux/UOS dependency.  It never uses a
    user's existing Excel process and is offered only as a local baseline when
    the same Flet workbench is started on Windows.
    """

    engine_name = ENGINE_EXCEL_COM
    application_progids = ("Excel.Application",)
    application_label = "Excel"

    def __init__(self) -> None:
        self._batch_excel = None
        self._batch_template = None
        self._batch_external = None
        self._batch_pythoncom = None

    @staticmethod
    def _open_com_workbook(excel, path: Path, *, read_only: bool):
        try:
            return excel.Workbooks.Open(
                str(path.resolve()), UpdateLinks=0, ReadOnly=read_only,
                IgnoreReadOnlyRecommended=True, AddToMru=False,
            )
        except Exception:
            return excel.Workbooks.Open(
                str(path.resolve()), UpdateLinks=0, ReadOnly=read_only
            )

    def begin_batch(self, *, template_path: Path, external_path: Path | None = None) -> None:
        """Start one private COM application and cache batch-level workbooks."""
        if self._batch_excel is not None:
            return
        self.require_available()
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = template = external = None
        try:
            launch_error = None
            for progid in self.application_progids:
                try:
                    excel = win32com.client.DispatchEx(progid)
                    break
                except Exception as exc:
                    launch_error = exc
            if excel is None:
                raise FormulaEngineError("{} COM 无法启动：{}".format(self.application_label, launch_error))
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            try:
                excel.Calculation = -4135  # xlCalculationManual
            except Exception:
                pass
            template = self._open_com_workbook(excel, template_path, read_only=True)
            if external_path:
                external = self._open_com_workbook(excel, external_path, read_only=True)
            self._batch_excel = excel
            self._batch_template = template
            self._batch_external = external
            self._batch_pythoncom = pythoncom
        except Exception:
            for workbook in (external, template):
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            raise

    def end_batch(self) -> None:
        """Close cached workbooks and the isolated COM application once."""
        excel = self._batch_excel
        for workbook in (self._batch_external, self._batch_template):
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
        if excel is not None:
            try:
                while excel.Workbooks.Count:
                    excel.Workbooks(1).Close(False)
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass
        if self._batch_pythoncom is not None:
            self._batch_pythoncom.CoUninitialize()
        self._batch_excel = None
        self._batch_template = None
        self._batch_external = None
        self._batch_pythoncom = None

    @classmethod
    def require_available(cls) -> None:
        if os.name != "nt":
            raise FormulaEngineError("{} COM 仅能在 Windows 环境中使用".format(cls.application_label))
        try:
            _prepare_pywin32_runtime()
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
            import winreg
            for progid in cls.application_progids:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + "\\CLSID"):
                        return
                except FileNotFoundError:
                    continue
            raise FormulaEngineError("未检测到 {} COM 注册信息".format(cls.application_label))
        except ImportError as exc:
            raise FormulaEngineError("未安装 pywin32，无法使用 {} COM 测试引擎".format(cls.application_label)) from exc

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        self.require_available()
        workbook_path = workbook_path.resolve()
        if workbook_path.suffix.casefold() != ".xlsx":
            raise FormulaEngineError("Excel COM 测试引擎仅支持 .xlsx 审核副本")
        capability = inspect_formula_capability(workbook_path, formula_ranges=formula_ranges)
        import pythoncom
        import win32com.client

        started = time.monotonic()
        excel = workbook = None
        com_initialized = False
        staged_path: Path | None = None
        stage_directory = None
        calculation_saved = False

        def open_workbook(path: Path):
            """Open with the full Excel parameter set, then its common subset.

            Excel installations differ in which optional COM arguments they
            accept.  The retry mirrors the Windows implementation's proven
            workbook-opening behaviour.
            """
            try:
                return excel.Workbooks.Open(
                    str(path), UpdateLinks=0, ReadOnly=False,
                    IgnoreReadOnlyRecommended=True, AddToMru=False,
                )
            except Exception:
                return excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)

        try:
            # Flet dispatches workflows from a background thread. COM must be
            # initialised in that exact thread, rather than relying on the UI
            # thread's apartment state.
            pythoncom.CoInitialize()
            com_initialized = True
            launch_error = None
            for progid in self.application_progids:
                try:
                    excel = win32com.client.DispatchEx(progid)
                    break
                except Exception as exc:
                    launch_error = exc
            if excel is None:
                raise FormulaEngineError("{} COM 无法启动：{}".format(self.application_label, launch_error))
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            try:
                workbook = open_workbook(workbook_path)
            except Exception as direct_error:
                # A few Excel versions reject otherwise valid workbooks when
                # their full path contains a deep hierarchy/non-ASCII names or
                # a just-released file handle.  Retrying a byte-identical copy
                # from a short local name is safe for this *test* adapter.  It
                # is deliberately a fallback: relative-link semantics keep
                # their original directory whenever direct opening succeeds.
                stage_directory = tempfile.TemporaryDirectory(prefix="base_audit_excel_com_")
                staged_path = Path(stage_directory.name) / "audit.xlsx"
                shutil.copy2(workbook_path, staged_path)
                try:
                    workbook = open_workbook(staged_path)
                except Exception as staged_error:
                    raise FormulaEngineError(
                        "{} 无法打开审核副本（原路径：{}；临时短路径重试也失败：{}；原始错误：{}）".format(
                            self.application_label,
                            workbook_path, staged_error, direct_error
                        )
                    ) from staged_error
            excel.Calculation = -4105  # xlCalculationAutomatic
            excel.CalculateFullRebuild()
            workbook.Save()
            calculation_saved = True
        except Exception as exc:
            raise FormulaEngineError("{} COM 计算失败：{}".format(self.application_label, exc)) from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
                workbook = None
            if calculation_saved and staged_path is not None and staged_path.exists():
                # Save has completed before Close.  Put the calculated OOXML
                # back at the audit-copy path that downstream extraction uses.
                shutil.copy2(staged_path, workbook_path)
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if com_initialized:
                pythoncom.CoUninitialize()
            if stage_directory is not None:
                stage_directory.cleanup()
        return CalculationResult(
            workbook_path, self.engine_name, time.monotonic() - started, capability,
            error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
        )

    def prepare_audit_copy(
        self,
        *,
        template_path: Path,
        audit_path: Path,
        definition: object,
        external_path: Path | None = None,
        external_sheet_names: Iterable[str] = (),
        apply_formulas: bool = True,
    ) -> CalculationResult:
        """Use the Windows COM workflow for the entire audit-copy mutation.

        This is intentionally separate from :meth:`recalculate`: opening a
        workbook which openpyxl has already rewritten is not equivalent to the
        Windows application.  The Windows edition copies external tabs and
        formula ranges through the *same* private Excel instance, then saves
        Excel's own OOXML.  The Windows-only test option must follow that exact
        lifecycle so it is a meaningful baseline for the native UOS workflow.
        """
        self.require_available()
        import pythoncom
        import win32com.client

        template_path = template_path.resolve()
        audit_path = audit_path.resolve()
        external_path = external_path.resolve() if external_path else None
        external_names = tuple(external_sheet_names)
        formula_ranges = tuple(getattr(definition, "copy_ranges", ()) or ())
        # COM is the formal calculator in this path.  Do not perform the
        # Python-engine capability scan here: it opens and parses the whole
        # audit XLSX once per institution but cannot affect Excel's result.
        # Formula errors are read from the live Excel values during issue
        # extraction below.
        capability = FormulaCapability(0, (), (), ())
        started = time.monotonic()
        using_batch = self._batch_excel is not None
        excel = self._batch_excel
        template = self._batch_template
        external = self._batch_external
        audit = None
        initialized = False
        try:
            if not using_batch:
                pythoncom.CoInitialize()
                initialized = True
                launch_error = None
                for progid in self.application_progids:
                    try:
                        excel = win32com.client.DispatchEx(progid)
                        break
                    except Exception as exc:
                        launch_error = exc
                if excel is None:
                    raise FormulaEngineError("{} COM 无法启动：{}".format(self.application_label, launch_error))
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.ScreenUpdating = False
                excel.EnableEvents = False
                excel.AskToUpdateLinks = False
                try:
                    excel.Calculation = -4135  # xlCalculationManual
                except Exception:
                    pass

            def open_book(path: Path, read_only: bool):
                try:
                    return excel.Workbooks.Open(
                        str(path.resolve()), UpdateLinks=0, ReadOnly=read_only,
                        IgnoreReadOnlyRecommended=True, AddToMru=False,
                    )
                except Exception:
                    return excel.Workbooks.Open(
                        str(path.resolve()), UpdateLinks=0, ReadOnly=read_only
                    )

            if template is None:
                template = open_book(template_path, True)
            audit = open_book(audit_path, False)
            if external_path and external_names:
                if external is None:
                    external = open_book(external_path, True)
                audit_names = {
                    str(audit.Worksheets(index).Name)
                    for index in range(1, audit.Worksheets.Count + 1)
                }
                missing = [
                    name for name in external_names
                    if name not in {
                        str(external.Worksheets(index).Name)
                        for index in range(1, external.Worksheets.Count + 1)
                    }
                ]
                conflicts = [name for name in external_names if name in audit_names]
                if missing:
                    raise FormulaEngineError("外部文件缺少工作表：" + "、".join(missing))
                if conflicts:
                    raise FormulaEngineError("外部文件工作表与报送文件重名，不能覆盖原表：" + "、".join(conflicts))
                for name in external_names:
                    external.Worksheets(name).Copy(
                        None, audit.Worksheets(audit.Worksheets.Count)
                    )

            if apply_formulas:
                audit_names = {
                    str(audit.Worksheets(index).Name)
                    for index in range(1, audit.Worksheets.Count + 1)
                }
                required = {
                    item.sheet_name for item in formula_ranges
                }
                missing = sorted(required - audit_names)
                if missing:
                    raise FormulaEngineError("报送文件缺少工作表：" + "、".join(missing))
                suspended_sheets = self._suspend_sheet_calculation(audit)
                try:
                    copied: set[tuple[str, str]] = set()
                    for item in formula_ranges:
                        identity = (item.sheet_name, item.address)
                        if identity in copied:
                            continue
                        copied.add(identity)
                        source_range = template.Worksheets(item.sheet_name).Range(item.address)
                        target_range = audit.Worksheets(item.sheet_name).Range(item.address)
                        source_range.Copy(Destination=target_range)
                        target_range.Formula = source_range.Formula
                finally:
                    self._restore_sheet_calculation(suspended_sheets)
                if bool(getattr(definition, "structured", False)):
                    rule_sheet = "审核规则"
                    current_names = {
                        str(audit.Worksheets(index).Name)
                        for index in range(1, audit.Worksheets.Count + 1)
                    }
                    if rule_sheet in current_names:
                        audit.Worksheets(rule_sheet).Delete()
                    template.Worksheets(rule_sheet).Copy(
                        After=audit.Worksheets(audit.Worksheets.Count)
                    )
                try:
                    excel.CutCopyMode = False
                except Exception:
                    pass
                # Match the Windows batch calculation: Calculate first, use a
                # full rebuild only when the host rejects normal calculation.
                try:
                    excel.Calculate()
                except Exception:
                    excel.CalculateFullRebuild()
                deadline = time.monotonic() + 120
                while True:
                    try:
                        pending = excel.CalculationState != 0  # xlDone
                    except Exception:
                        pending = False
                    if not pending:
                        break
                    if time.monotonic() >= deadline:
                        raise FormulaEngineError("Excel 公式计算超过 120 秒")
                    time.sleep(0.1)
            audit.Save()
        except Exception as exc:
            raise FormulaEngineError("{} COM 完整审核副本处理失败：{}".format(self.application_label, exc)) from exc
        finally:
            owned_workbooks = (audit,) if using_batch else (external, audit, template)
            for workbook in owned_workbooks:
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
            if excel is not None and not using_batch:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if initialized:
                pythoncom.CoUninitialize()

        return CalculationResult(
            audit_path, self.engine_name, time.monotonic() - started, capability,
            # The configured formula cells are read in bulk later in the same
            # COM session.  Re-opening each large XLSX here solely for an
            # informational error scan costs several seconds per institution.
            error_cells=(),
        )

    def read_issue_values(
        self, *, workbook_path: Path, rules: Iterable[object], extraction_ranges: Iterable[object] = (),
    ) -> dict[tuple[str, str], Any]:
        """Read configured formula/target values with one COM range call per sheet."""
        if self._batch_excel is None:
            raise FormulaEngineError("公式结果提取需要已启动的 Excel/WPS 批量会话")
        allowed = tuple(extraction_ranges)
        enabled = [
            rule for rule in rules
            if getattr(rule, "enabled", False) and (
                not allowed or any(
                    item.sheet_name == rule.sheet_name
                    and self._address_in_range(rule.formula_cell, item.address)
                    for item in allowed
                )
            )
        ]
        if not enabled:
            return {}
        by_sheet: dict[str, list[str]] = {}
        formula_cells: set[tuple[str, str]] = set()
        for rule in enabled:
            by_sheet.setdefault(rule.sheet_name, []).append(rule.formula_cell)
            formula_cells.add((rule.sheet_name, rule.formula_cell))
            if not getattr(rule, "value_from_result", False):
                by_sheet.setdefault(rule.sheet_name, []).append(rule.target_cell)

        workbook = self._open_com_workbook(self._batch_excel, workbook_path.resolve(), read_only=True)
        values: dict[tuple[str, str], Any] = {}
        try:
            for sheet_name, addresses in by_sheet.items():
                sheet = workbook.Worksheets(sheet_name)
                positions = [coordinate_to_tuple(address.replace("$", "")) for address in addresses]
                first_row = min(row for row, _ in positions)
                last_row = max(row for row, _ in positions)
                first_column = min(column for _, column in positions)
                last_column = max(column for _, column in positions)
                area = sheet.Range(
                    "{}{}:{}{}".format(
                        get_column_letter(first_column), first_row,
                        get_column_letter(last_column), last_row,
                    )
                )
                raw = area.Value2
                matrix = self._com_matrix(raw, int(area.Rows.Count), int(area.Columns.Count))
                for address, (row, column) in zip(addresses, positions):
                    values[(sheet_name, address)] = matrix[row - first_row][column - first_column]

            # Multi-cell Value2 reads may turn an Excel error into ``None``.
            # Only those few formula cells need a separate display-text read.
            for sheet_name, address in formula_cells:
                if values.get((sheet_name, address)) is None:
                    text = str(workbook.Worksheets(sheet_name).Range(address).Text or "").strip()
                    if text.upper().startswith(("#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!")):
                        values[(sheet_name, address)] = text
            return values
        finally:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass

    @staticmethod
    def _address_in_range(address: str, range_address: str) -> bool:
        row, column = coordinate_to_tuple(address.replace("$", ""))
        min_col, min_row, max_col, max_row = range_boundaries(range_address.replace("$", ""))
        return min_row <= row <= max_row and min_col <= column <= max_col

    @staticmethod
    def _com_matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
        if rows == 1 and columns == 1:
            return [[value]]
        if rows == 1:
            # pywin32 returns a one-row Excel Range inconsistently: either
            # ``(a, b)`` or ``((a, b),)``.  Always unwrap to one value row.
            if (
                isinstance(value, (tuple, list)) and len(value) == 1
                and isinstance(value[0], (tuple, list))
            ):
                return [list(value[0])]
            return [list(value)]
        if columns == 1:
            if (
                isinstance(value, (tuple, list)) and len(value) == 1
                and isinstance(value[0], (tuple, list))
            ):
                value = value[0]
            return [[item] for item in value]
        return [list(item) for item in value]

    @staticmethod
    def _suspend_sheet_calculation(workbook: Any) -> list[tuple[Any, bool]]:
        """Match the Windows batch path for sheets that ignore app-level manual calc."""
        suspended: list[tuple[Any, bool]] = []
        for index in range(1, workbook.Worksheets.Count + 1):
            sheet = workbook.Worksheets(index)
            try:
                enabled = bool(sheet.EnableCalculation)
                if enabled:
                    sheet.EnableCalculation = False
                suspended.append((sheet, enabled))
            except Exception:
                continue
        return suspended

    @staticmethod
    def _restore_sheet_calculation(sheets: Iterable[tuple[Any, bool]]) -> None:
        for sheet, was_enabled in sheets:
            if not was_enabled:
                continue
            try:
                sheet.EnableCalculation = True
            except Exception:
                pass

    def extract_rendered_conditional_formats(
        self, *, workbook_path: Path, ranges: Iterable[object],
        structure_ranges: Iterable[object], period: str, batch_id: str,
        source_file: Path,
    ) -> list[Issue]:
        """Read Excel/WPS ``DisplayFormat`` colours from the batch COM session."""
        if self._batch_excel is None:
            raise FormulaEngineError("条件格式提取需要已启动的 Excel/WPS 批量会话")
        workbook_path = workbook_path.resolve()
        ranges = tuple(ranges)
        structure_ranges = tuple(structure_ranges)
        all_started = time.monotonic()
        workbook = self._open_com_workbook(self._batch_excel, workbook_path, read_only=True)
        open_seconds = time.monotonic() - all_started
        region_seconds = 0.0
        colour_seconds = 0.0
        candidate_count = 0
        skipped_blank_count = 0
        scope_seconds = 0.0
        scope_count = 0
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        issues: list[Issue] = []
        seen: set[tuple[str, int, int]] = set()
        structure_cache: dict[str, list[tuple[int, int, int, int, list[list[Any]]]]] = {}

        # Conditional-format ranges commonly contain more than a thousand
        # cells.  Read table-structure labels once per range; the former
        # implementation called ``sheet.Cells(...).Value2`` repeatedly for
        # every triggered cell, which dominated the entire audit runtime.
        region_started = time.monotonic()
        for structure in structure_ranges:
            sheet = workbook.Worksheets(structure.sheet_name)
            left, top, right, bottom = range_boundaries(structure.address)
            selected = sheet.Range(structure.address)
            matrix = self._com_matrix(
                selected.Value2, int(selected.Rows.Count), int(selected.Columns.Count)
            )
            structure_cache.setdefault(str(structure.sheet_name), []).append(
                (left, top, right, bottom, matrix)
            )
        region_seconds += time.monotonic() - region_started

        # A named 条件格式区域 may be broader than the actual AppliesTo of its
        # rules.  Cache those rule scopes once per worksheet and use them as a
        # cheap Python-side membership test before any DisplayFormat COM call.
        # If an office suite exposes an incomplete FormatConditions collection,
        # fall back to the explicit named range rather than risking omissions.
        rule_scopes: dict[str, list[tuple[int, int, int, int]]] = {}
        scope_started = time.monotonic()
        for sheet_name in {item.sheet_name for item in ranges}:
            try:
                sheet = workbook.Worksheets(sheet_name)
                conditions = sheet.UsedRange.FormatConditions
                scopes: list[tuple[int, int, int, int]] = []
                for condition_index in range(1, int(conditions.Count) + 1):
                    applies_to = conditions.Item(condition_index).AppliesTo
                    for area_index in range(1, int(applies_to.Areas.Count) + 1):
                        applies_area = applies_to.Areas.Item(area_index)
                        first_row = int(applies_area.Row)
                        first_column = int(applies_area.Column)
                        scopes.append((
                            first_column,
                            first_row,
                            first_column + int(applies_area.Columns.Count) - 1,
                            first_row + int(applies_area.Rows.Count) - 1,
                        ))
                rule_scopes[str(sheet.Name)] = scopes
                scope_count += len(scopes)
            except Exception:
                # No entry means named-range-only fallback for this sheet.
                continue
        scope_seconds = time.monotonic() - scope_started

        def covered_by_rule(sheet_name: str, row: int, column: int) -> bool:
            scopes = rule_scopes.get(sheet_name)
            if scopes is None:
                return True
            return any(
                left <= column <= right and top <= row <= bottom
                for left, top, right, bottom in scopes
            )

        def label_for(sheet, row: int, column: int) -> str:
            areas = structure_cache.get(str(sheet.Name), [])
            if not areas:
                return ""

            def display_value(current_row: int, current_column: int) -> str:
                for left, top, right, bottom, matrix in areas:
                    if not (left <= current_column <= right and top <= current_row <= bottom):
                        continue
                    try:
                        value = matrix[current_row - top][current_column - left]
                    except IndexError:
                        value = None
                    if value is None or isinstance(value, (int, float)) and not isinstance(value, bool):
                        return ""
                    text = str(value).strip()
                    # 表结构区域也覆盖数据列；指标仅拼文本型行头/列表头。
                    return "" if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text.replace(",", "")) else text
                return ""

            min_row = min(item[1] for item in areas)
            min_column = min(item[0] for item in areas)
            left_labels: list[str] = []
            for current_column in range(min_column, column):
                value = display_value(row, current_column)
                if value and value not in left_labels:
                    left_labels.append(value)
            top_labels: list[str] = []
            for current_row in range(min_row, row):
                value = display_value(current_row, column)
                if value and value not in top_labels:
                    top_labels.append(value)
            parts = []
            if left_labels:
                parts.append("_".join(left_labels))
            if top_labels:
                parts.append("_".join(top_labels))
            return "｜".join(parts)

        try:
            for area in ranges:
                region_started = time.monotonic()
                sheet = workbook.Worksheets(area.sheet_name)
                selected = sheet.Range(area.address)
                left, top, right, bottom = range_boundaries(area.address)
                row_count, column_count = int(selected.Rows.Count), int(selected.Columns.Count)
                displayed_values = self._com_matrix(selected.Value2, row_count, column_count)
                region_seconds += time.monotonic() - region_started

                # Excel/WPS does not expose an API that enumerates only cells
                # whose conditional-format formula evaluated TRUE.  FindFormat
                # sees stored formatting rather than DisplayFormat, so it is
                # not a correct substitute.  Value2 is read once for the full
                # named area, then only populated candidates need the costly
                # per-cell rendered-colour comparison.  This preserves the
                # template convention that a reportable conditional-format
                # result always has a displayed cell value, while avoiding a
                # FormatConditions COM call for every cell in a large area.
                for row_index in range(1, row_count + 1):
                    for column_index in range(1, column_count + 1):
                        value = displayed_values[row_index - 1][column_index - 1]
                        if value in (None, ""):
                            skipped_blank_count += 1
                            continue
                        row, column = top + row_index - 1, left + column_index - 1
                        if not covered_by_rule(str(sheet.Name), row, column):
                            continue
                        identity = (str(sheet.Name), row, column)
                        if identity in seen:
                            continue
                        cell = selected.Cells(row_index, column_index)
                        candidate_count += 1
                        colour_started = time.monotonic()
                        try:
                            displayed = int(cell.DisplayFormat.Interior.Color)
                            base = int(cell.Interior.Color)
                        except Exception as exc:
                            raise FormulaEngineError(
                                "{} 无法读取条件格式实际显示颜色".format(self.application_label)
                            ) from exc
                        finally:
                            colour_seconds += time.monotonic() - colour_started
                        if displayed == base:
                            continue
                        seen.add(identity)
                        address = "{}{}".format(get_column_letter(column), row)
                        detail = "条件格式填充已触发，请核实"
                        try:
                            comment = cell.Comment
                            if comment is not None:
                                detail = str(comment.Text() or "").strip() or detail
                        except Exception:
                            pass
                        indicator = label_for(sheet, row, column)
                        identity = "｜".join((
                            normalize_template_name(source_file.stem) or "工作簿未识别",
                            str(sheet.Name), address, indicator or "条件格式填充",
                        ))
                        issues.append(Issue(
                            issue_id=identity,
                            period=period, batch_id=batch_id, audit_time=stamp, triggered=True,
                            status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                            org_code="", org_name="", report_code="", sheet_name=str(sheet.Name),
                            rule_id=identity, severity="总行条件格式触发",
                            formula_cell=address, target_cell=address, target_value=value,
                            formula_result=value, message=detail, source_file=str(source_file.resolve()),
                            audit_file=str(workbook_path), check_field=indicator, detail=detail,
                            display_fill_color=displayed,
                        ))
        finally:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
            self.last_conditional_metrics = {
                "open_seconds": open_seconds,
                "region_seconds": region_seconds,
                "scope_seconds": scope_seconds,
                "scope_count": scope_count,
                "colour_seconds": colour_seconds,
                "candidate_count": candidate_count,
                "skipped_blank_count": skipped_blank_count,
                "total_seconds": time.monotonic() - all_started,
            }
        return issues


class WpsComCalculator(ExcelComCalculator):
    """Windows WPS Spreadsheet COM adapter, kept separate from Excel COM."""

    engine_name = ENGINE_WPS_COM
    application_progids = ("ket.Application", "KET.Application")
    application_label = "WPS"


class CrossValidatedCalculator:
    """Run formulas and IronCalc independently and require equal cached values."""

    engine_name = ENGINE_COMBINED

    def __init__(self) -> None:
        # A batch commonly contains several workbooks exported by the same
        # reporting system. Once IronCalc has rejected that package family,
        # retrying it for every remaining file only wastes time.
        self._ironcalc_disabled_reason = ""

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        workbook_path = workbook_path.resolve()
        capability = inspect_formula_capability(workbook_path, formula_ranges=formula_ranges)
        started = time.monotonic()
        with tempfile.TemporaryDirectory(prefix="base_audit_crosscheck_") as text:
            temp_dir = Path(text)
            formulas_path = temp_dir / "formulas.xlsx"
            ironcalc_path = temp_dir / "ironcalc.xlsx"
            shutil.copy2(workbook_path, formulas_path)
            if self._ironcalc_disabled_reason:
                formulas_result = PythonFormulaCalculator().recalculate(
                    formulas_path, formula_ranges=formula_ranges
                )
                shutil.copy2(formulas_path, workbook_path)
                return CalculationResult(
                    workbook_path, "formulas（IronCalc 已跳过，未交叉验证）",
                    time.monotonic() - started, formulas_result.capability, fallback_used=True,
                    error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
                    diagnostic="本批已发现 IronCalc 与此类 XLSX 不兼容：{}".format(self._ironcalc_disabled_reason),
                )
            shutil.copy2(workbook_path, ironcalc_path)
            # Both engines receive independent copies. Running them in parallel
            # reduces wall-clock latency while avoiding concurrent writes to the
            # retained audit workbook.
            with ThreadPoolExecutor(max_workers=2, thread_name_prefix="formula_check") as executor:
                formulas_future = executor.submit(
                    PythonFormulaCalculator().recalculate, formulas_path, formula_ranges=formula_ranges
                )
                ironcalc_future = executor.submit(
                    IronCalcCalculator().recalculate, ironcalc_path, formula_ranges=formula_ranges
                )
                formulas_result = None
                ironcalc_result = None
                formulas_error = None
                ironcalc_error = None
                try:
                    formulas_result = formulas_future.result()
                except Exception as exc:
                    formulas_error = exc
                try:
                    ironcalc_result = ironcalc_future.result()
                except Exception as exc:
                    ironcalc_error = exc
            if formulas_result is not None and ironcalc_error is not None:
                self._ironcalc_disabled_reason = str(ironcalc_error)
                shutil.copy2(formulas_path, workbook_path)
                return CalculationResult(
                    workbook_path, "formulas（IronCalc 失败，未交叉验证）",
                    time.monotonic() - started, formulas_result.capability, fallback_used=True,
                    error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
                    diagnostic="IronCalc 本次未能计算：{}；已使用 formulas 结果继续提取。".format(ironcalc_error),
                )
            if ironcalc_result is not None and formulas_error is not None:
                shutil.copy2(ironcalc_path, workbook_path)
                return CalculationResult(
                    workbook_path, "IronCalc（formulas 失败，未交叉验证）",
                    time.monotonic() - started, ironcalc_result.capability, fallback_used=True,
                    error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
                    diagnostic="formulas 本次未能计算：{}；已使用 IronCalc 结果继续提取。".format(formulas_error),
                )
            if formulas_error is not None and ironcalc_error is not None:
                raise FormulaEngineError("formulas 计算失败：{}；IronCalc 计算失败：{}".format(formulas_error, ironcalc_error))
            if formulas_result is None or ironcalc_result is None:
                raise FormulaEngineError("组合计算未产生可用结果")
            formulas_values = _formula_cache_values(workbook_path, formulas_path, formula_ranges=formula_ranges)
            ironcalc_values = _formula_cache_values(workbook_path, ironcalc_path, formula_ranges=formula_ranges)
            differences = [
                "{}!{}：formulas={!r}，IronCalc={!r}".format(sheet, address, formulas_values.get((sheet, address)), ironcalc_values.get((sheet, address)))
                for sheet, address in sorted(set(formulas_values) | set(ironcalc_values))
                if formulas_values.get((sheet, address)) != ironcalc_values.get((sheet, address))
            ]
            if differences:
                suffix = "；其余 {} 项".format(len(differences) - 8) if len(differences) > 8 else ""
                # Preserve a usable Python result for issue extraction.  The
                # caller receives the discrepancy in diagnostics; it is not a
                # reason to discard #REF!/#NAME? evidence or stop the batch.
                shutil.copy2(formulas_path, workbook_path)
                return CalculationResult(
                    workbook_path, self.engine_name + "（结果不一致，采用 formulas）",
                    time.monotonic() - started, formulas_result.capability,
                    error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
                )
            # Both engines agree. Keep the formulas copy because its adapter
            # already preserves the source workbook's OOXML parts and caches.
            shutil.copy2(formulas_path, workbook_path)
        return CalculationResult(
            workbook_path, self.engine_name, time.monotonic() - started, capability,
            error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
        )


def get_calculator(engine_name: str) -> Calculator:
    normalized = (engine_name or ENGINE_DEFAULT).strip()
    if normalized == ENGINE_LIBREOFFICE:
        return LibreOfficeAdapter()
    raise FormulaEngineError("统信版只支持 LibreOffice Calc：{}".format(engine_name))


def engine_status() -> dict[str, object]:
    """Return data used by the Flet setting page without starting Calc."""
    libre = LibreOfficeAdapter()
    available, detail = libre.available()
    return {"libreoffice": {"available": available, "detail": detail}}


def engine_available(engine_name: str, status: dict[str, object] | None = None) -> bool:
    current = status or engine_status()
    return engine_name == ENGINE_LIBREOFFICE and bool(dict(current.get("libreoffice", {})).get("available"))


def preferred_engine(status: dict[str, object] | None = None) -> str:
    """Choose one available engine at startup; never use runtime fallback."""
    current = status or engine_status()
    availability = {name: engine_available(name, current) for name in ENGINE_PRIORITY}
    return next((name for name in ENGINE_PRIORITY if availability[name]), ENGINE_DEFAULT)
