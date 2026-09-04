"""Formula-calculation adapters for the Windows Flet edition.

The workbook workflow is intentionally independent from the calculation
runtime: openpyxl performs all copying; this module calculates the prepared
``*_审核版.xlsx``. Configured engines are Excel COM or WPS COM only.
"""
from __future__ import annotations

import re
import os
import shutil
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Protocol

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries, get_column_letter

from .models import Issue
from .template import normalize_template_name
from .conditional_format import evaluate_expression_formula


ENGINE_EXCEL_COM = "Excel COM"
ENGINE_WPS_COM = "WPS COM"
# Windows users do not choose an Office implementation.  The application
# verifies Excel first and falls back to WPS when it actually starts a COM
# server.  A registry key alone is not sufficient: stale Office registrations
# are common after uninstalling Excel.
ENGINE_OFFICE_COM = "Excel/WPS"
ENGINE_WPS_FIRST_COM = "WPS/Excel"
ENGINE_OPTIONS = (ENGINE_OFFICE_COM, ENGINE_WPS_FIRST_COM)
ENGINE_PRIORITY = ENGINE_OPTIONS
ENGINE_DEFAULT = ENGINE_OFFICE_COM
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


class FormulaEngineError(RuntimeError):
    """A calculation engine cannot safely calculate this workbook."""


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
            # Do not inspect only HKEY_CLASSES_ROOT here.  32-bit WPS is often
            # registered in the 32-bit registry view and invisible to this
            # 64-bit Python process. ``probe_available`` performs the decisive
            # DispatchEx test in the COM registry view used at runtime.
            return
        except ImportError as exc:
            raise FormulaEngineError("未安装 pywin32，无法使用 {} COM".format(cls.application_label)) from exc

    @classmethod
    def probe_available(cls) -> str:
        """Start and immediately close a private COM server.

        This is deliberately stronger than ``require_available()``.  It
        prevents an old Excel registry entry from being reported as a usable
        engine when Excel is no longer installed.
        """
        cls.require_available()
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        app = None
        try:
            last_error = None
            for progid in cls.application_progids:
                try:
                    app = win32com.client.DispatchEx(progid)
                    break
                except Exception as exc:
                    last_error = exc
            if app is None:
                raise FormulaEngineError("{} COM 无法启动：{}".format(cls.application_label, last_error))
            try:
                app.Visible = False
                app.DisplayAlerts = False
            except Exception:
                pass
            version = str(getattr(app, "Version", "") or "").strip()
            return version
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        self.require_available()
        workbook_path = workbook_path.resolve()
        if workbook_path.suffix.casefold() != ".xlsx":
            raise FormulaEngineError("Excel COM 仅支持 .xlsx 审核副本")
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

    def _prepare_conditional_format_sheet(self, workbook: Any, sheet: Any) -> None:
        """Allow an office-specific adapter to refresh rendered cell formats.

        Excel updates ``DisplayFormat`` while hidden.  WPS has a separate
        implementation and needs an explicit refresh before its displayed
        colour can be trusted; the override below keeps that workaround out of
        the shared extraction loop.
        """

    @staticmethod
    def _rendered_fill_colour(cell: Any) -> int:
        return int(cell.DisplayFormat.Interior.Color)

    @staticmethod
    def _conditional_rule_text(condition: Any, *, origin: str, destination: str) -> str:
        """Return a readable, destination-relative Excel condition formula.

        ``FormatCondition.Formula1`` is stored relative to the first cell of
        ``AppliesTo``.  Translate it to the triggered cell so a summary shows
        the rule the reviewer can actually verify (for example ``C10>0``),
        rather than a generic "conditional formatting triggered" message.
        """
        formulas: list[str] = []
        for attribute in ("Formula1", "Formula2"):
            try:
                value = str(getattr(condition, attribute) or "").strip()
            except Exception:
                value = ""
            if not value:
                continue
            try:
                value = Translator(value, origin=origin).translate_formula(destination)
            except Exception:
                # Keep COM's original formula when it is not translatable
                # (for example a constant threshold or suite-specific syntax).
                pass
            # 输出是面向审核人员的描述。绝对引用符号只服务于 Excel
            # 复制公式语义，展示为 “C10>0” 比 “$C$10>0” 更易阅读。
            formulas.append(value.lstrip("=").replace("$", ""))

        # “只为包含以下内容的单元格设置格式”属于 Excel 的 xlCellValue
        # 条件格式：Formula1 只保存阈值（例如 0），比较对象与运算符分别
        # 存在 Type/Operator 中。摘要必须还原为“D26>0”，而不是孤立的“0”。
        try:
            condition_type = int(getattr(condition, "Type"))
        except Exception:
            condition_type = None
        try:
            operator = int(getattr(condition, "Operator"))
        except Exception:
            operator = None
        operator_symbols = {
            1: (">=", "<="),  # xlBetween
            2: ("<", ">"),    # xlNotBetween
            3: ("=",),         # xlEqual
            4: ("<>",),        # xlNotEqual
            5: (">",),         # xlGreater
            6: ("<",),         # xlLess
            7: (">=",),        # xlGreaterEqual
            8: ("<=",),        # xlLessEqual
        }
        if condition_type == 1 and formulas and operator in operator_symbols:
            symbols = operator_symbols[operator]
            if operator == 1 and len(formulas) >= 2:
                text = "{}{}{} 且 {}{}{}".format(
                    destination, symbols[0], formulas[0], destination, symbols[1], formulas[1]
                )
            elif operator == 2 and len(formulas) >= 2:
                text = "{}{}{} 或 {}{}{}".format(
                    destination, symbols[0], formulas[0], destination, symbols[1], formulas[1]
                )
            else:
                text = "{}{}{}".format(destination, symbols[0], formulas[0])
            return "条件格式规则：{}".format(text)
        if formulas:
            return "条件格式规则：{}".format("；".join(dict.fromkeys(formulas)))
        try:
            rule_type = str(getattr(condition, "Type"))
        except Exception:
            rule_type = "未提供公式"
        return "条件格式规则：类型 {}（未提供可显示的公式）".format(rule_type)

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
        rule_scopes: dict[str, list[tuple[int, int, int, int, str, Any]]] = {}
        scope_started = time.monotonic()
        for sheet_name in {item.sheet_name for item in ranges}:
            try:
                sheet = workbook.Worksheets(sheet_name)
                conditions = sheet.UsedRange.FormatConditions
                scopes: list[tuple[int, int, int, int, str, Any]] = []
                for condition_index in range(1, int(conditions.Count) + 1):
                    condition = conditions.Item(condition_index)
                    applies_to = condition.AppliesTo
                    for area_index in range(1, int(applies_to.Areas.Count) + 1):
                        applies_area = applies_to.Areas.Item(area_index)
                        first_row = int(applies_area.Row)
                        first_column = int(applies_area.Column)
                        scopes.append((
                            first_column,
                            first_row,
                            first_column + int(applies_area.Columns.Count) - 1,
                            first_row + int(applies_area.Rows.Count) - 1,
                            "{}{}".format(get_column_letter(first_column), first_row),
                            condition,
                        ))
                rule_scopes[str(sheet.Name)] = scopes
                scope_count += len(scopes)
            except Exception:
                # No entry means named-range-only fallback for this sheet.
                continue
        scope_seconds = time.monotonic() - scope_started

        def matching_rules(sheet_name: str, row: int, column: int) -> list[tuple[str, Any]] | None:
            scopes = rule_scopes.get(sheet_name)
            if scopes is None:
                return None
            return [
                (origin, condition)
                for left, top, right, bottom, origin, condition in scopes
                if left <= column <= right and top <= row <= bottom
            ]

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
                self._prepare_conditional_format_sheet(workbook, sheet)
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
                        matched_rules = matching_rules(str(sheet.Name), row, column)
                        if matched_rules == []:
                            continue
                        identity = (str(sheet.Name), row, column)
                        if identity in seen:
                            continue
                        cell = selected.Cells(row_index, column_index)
                        candidate_count += 1
                        colour_started = time.monotonic()
                        try:
                            displayed = self._rendered_fill_colour(cell)
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
                        detail = ""
                        try:
                            comment = cell.Comment
                            if comment is not None:
                                detail = str(comment.Text() or "").strip()
                        except Exception:
                            pass
                        if not detail:
                            destination = "{}{}".format(get_column_letter(column), row)
                            descriptions = [
                                self._conditional_rule_text(condition, origin=origin, destination=destination)
                                for origin, condition in (matched_rules or [])
                            ]
                            detail = "；".join(dict.fromkeys(descriptions)) or "条件格式规则：实际填充色已变化"
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
    # WPS releases have used both KET and ET registrations; installations can
    # also expose only one casing / registry view. Try every known spreadsheet
    # automation ProgID before declaring WPS unavailable.
    application_progids = ("ket.Application", "KET.Application", "et.Application", "ET.Application")
    application_label = "WPS"

    @staticmethod
    def _cellis_operand(sheet: Any, raw: object) -> object:
        """Resolve a simple WPS/OOXML ``cellIs`` threshold without guessing."""
        text = str(raw or "").strip().lstrip("=")
        if not text:
            return None
        if len(text) >= 2 and text[0] == text[-1] == '"':
            return text[1:-1]
        try:
            return float(text)
        except ValueError:
            pass
        match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", text)
        if match:
            return sheet["{}{}".format(match.group(1), match.group(2))].value
        return None

    @classmethod
    def _cellis_matches(cls, rule: Any, value: object, sheet: Any) -> bool:
        if value in (None, ""):
            return False
        formulas = list(rule.formula or [])
        first = cls._cellis_operand(sheet, formulas[0] if formulas else None)
        second = cls._cellis_operand(sheet, formulas[1] if len(formulas) > 1 else None)
        if first is None:
            return False
        try:
            left, right = float(value), float(first)
            upper = float(second) if second is not None else None
        except (TypeError, ValueError):
            left, right, upper = str(value), str(first), str(second) if second is not None else None
        operator = str(rule.operator or "").casefold()
        return {
            "equal": left == right, "notequal": left != right,
            "greaterthan": left > right, "greaterthanorequal": left >= right,
            "lessthan": left < right, "lessthanorequal": left <= right,
            "between": upper is not None and right <= left <= upper,
            "notbetween": upper is not None and not (right <= left <= upper),
        }.get(operator, False)

    @staticmethod
    def _has_dxf_fill(book: Any, rule: Any) -> bool:
        try:
            style = book._differential_styles[rule.dxfId]
            fill = style.fill
            return bool(fill and (fill.patternType or fill.fgColor.type or fill.bgColor.type))
        except Exception:
            return False

    @staticmethod
    def _expression_rule_text(formula: str, origin: str, destination: str) -> str:
        """Return a readable, destination-relative expression rule description.

        Mirrors ``_conditional_rule_text`` for the OOXML fallback: the stored
        ``Formula1`` is relative to the AppliesTo anchor, so translate it to the
        triggered cell (``AND(C26>0,C26>1)`` instead of ``AND(C5>0,C5>1)``).
        """
        text = str(formula or "").strip()
        # openpyxl stores the expression formula without a leading "=", but the
        # Translator only rewrites relative references when it sees a formula
        # (a leading "="), otherwise it returns the text verbatim.
        if text and not text.startswith("="):
            text = "=" + text
        try:
            translated = Translator(text, origin=origin).translate_formula(destination)
        except Exception:
            translated = text
        return "条件格式规则：{}".format(str(translated).lstrip("=").replace("$", ""))

    def _extract_rule_based_conditional_formats(
        self, *, workbook_path: Path, ranges: Iterable[object], structure_ranges: Iterable[object],
        period: str, batch_id: str, source_file: Path,
    ) -> list[Issue]:
        """WPS fallback for rules it exposes but cannot render through COM.

        Evaluates the submission's own OOXML ``cellIs`` and ``expression`` rules
        directly, translating expression references per target cell from the
        rule's anchor (top-left of ``AppliesTo``).  Rules that cannot be
        reliably parsed are not guessed: they yield no trigger and are counted
        in ``last_conditional_metrics["unsupported_count"]`` so the run log can
        report “WPS 条件规则暂不支持”.
        """
        book = load_workbook(workbook_path, data_only=True, keep_links=False)
        try:
            issues: list[Issue] = []
            seen: set[tuple[str, int, int]] = set()
            unsupported = 0
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            symbols = {
                "equal": "=", "notequal": "<>", "greaterthan": ">",
                "greaterthanorequal": ">=", "lessthan": "<", "lessthanorequal": "<=",
                "between": "介于", "notbetween": "不介于",
            }
            structure_cache: dict[str, list[tuple[int, int, int, int, list[list[Any]]]]] = {}
            for structure in structure_ranges:
                if structure.sheet_name not in book.sheetnames:
                    continue
                structure_sheet = book[structure.sheet_name]
                s_left, s_top, s_right, s_bottom = range_boundaries(structure.address)
                structure_cache.setdefault(structure.sheet_name, []).append((
                    s_left, s_top, s_right, s_bottom,
                    [
                        [structure_sheet.cell(row=r, column=c).value for c in range(s_left, s_right + 1)]
                        for r in range(s_top, s_bottom + 1)
                    ],
                ))

            def label_for(sheet_name: str, row: int, column: int) -> str:
                areas = structure_cache.get(sheet_name, [])
                if not areas:
                    return ""

                def display_value(current_row: int, current_column: int) -> str:
                    for area_left, area_top, area_right, area_bottom, matrix in areas:
                        if not (area_left <= current_column <= area_right and area_top <= current_row <= area_bottom):
                            continue
                        try:
                            value = matrix[current_row - area_top][current_column - area_left]
                        except IndexError:
                            value = None
                        if value is None or isinstance(value, (int, float)) and not isinstance(value, bool):
                            return ""
                        text = str(value).strip()
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

            for area in ranges:
                if area.sheet_name not in book.sheetnames:
                    continue
                sheet = book[area.sheet_name]
                left, top, right, bottom = range_boundaries(area.address)
                for conditional in sheet.conditional_formatting:
                    try:
                        applies_ranges = list(conditional.sqref.ranges)
                    except Exception:
                        applies_ranges = []
                    if not applies_ranges:
                        continue
                    # Excel anchors a rule's relative references to the top-left
                    # cell of its AppliesTo.  A multi-area AppliesTo keeps one
                    # anchor: the top-left of the whole range.
                    anchor_col = min(applies.bounds[0] for applies in applies_ranges)
                    anchor_row = min(applies.bounds[1] for applies in applies_ranges)
                    anchor_address = "{}{}".format(get_column_letter(anchor_col), anchor_row)
                    for rule in conditional.rules:
                        rule_type = str(rule.type or "").casefold()
                        if rule_type not in ("cellis", "expression"):
                            unsupported += 1
                            continue
                        if not self._has_dxf_fill(book, rule):
                            continue
                        formula = None
                        if rule_type == "expression":
                            formula = str((rule.formula or [""])[0] or "").strip()
                            if formula.startswith("="):
                                formula = formula[1:].strip()
                            if not formula:
                                unsupported += 1
                                continue
                            # Reject unparseable/unsupported formulas once,
                            # before iterating cells, so they never trigger.
                            try:
                                evaluate_expression_formula(
                                    formula, sheet, anchor_row, anchor_col, anchor_row, anchor_col
                                )
                            except (ValueError, NotImplementedError):
                                unsupported += 1
                                continue
                        for applies in applies_ranges:
                            min_col, min_row, max_col, max_row = applies.bounds
                            for row in range(max(top, min_row), min(bottom, max_row) + 1):
                                for column in range(max(left, min_col), min(right, max_col) + 1):
                                    identity = (area.sheet_name, row, column)
                                    if identity in seen:
                                        continue
                                    value = sheet.cell(row, column).value
                                    if rule_type == "cellis":
                                        matched = self._cellis_matches(rule, value, sheet)
                                    else:
                                        matched = evaluate_expression_formula(
                                            formula, sheet, anchor_row, anchor_col, row, column
                                        )
                                    if not matched:
                                        continue
                                    seen.add(identity)
                                    address = "{}{}".format(get_column_letter(column), row)
                                    if rule_type == "cellis":
                                        threshold = str((rule.formula or [""])[0]).lstrip("=").replace("$", "")
                                        detail = "条件格式规则：{}{}{}".format(
                                            address, symbols.get(str(rule.operator or "").casefold(), "?"), threshold
                                        )
                                    else:
                                        detail = self._expression_rule_text(formula, anchor_address, address)
                                    comment = sheet.cell(row, column).comment
                                    if comment and comment.text and comment.text.strip():
                                        detail = comment.text.strip()
                                    indicator = label_for(area.sheet_name, row, column)
                                    key = normalize_template_name(source_file.stem) or source_file.stem
                                    rule_id = "｜".join((key, area.sheet_name, address, indicator or "条件格式填充"))
                                    issues.append(Issue(
                                        issue_id=rule_id, period=period, batch_id=batch_id, audit_time=stamp,
                                        triggered=True, status="", first_seen_period="", previous_seen_period="",
                                        consecutive_count=1, org_code="", org_name="", report_code="",
                                        sheet_name=area.sheet_name, rule_id=rule_id, severity="总行条件格式触发",
                                        formula_cell=address, target_cell=address,
                                        target_value=sheet.cell(row, column).value,
                                        formula_result=sheet.cell(row, column).value, message=detail,
                                        source_file=str(source_file.resolve()), audit_file=str(workbook_path),
                                        check_field=indicator, detail=detail,
                                    ))
            self.last_conditional_metrics = {
                "mode": "WPS 条件规则计算",
                "issue_count": len(issues),
                "unsupported_count": unsupported,
            }
            return issues
        finally:
            book.close()

    def _prepare_conditional_format_sheet(self, workbook: Any, sheet: Any) -> None:
        """Refresh WPS's rendered conditional-format cache once per sheet.

        WPS COM may otherwise return the static ``Interior.Color`` from
        ``DisplayFormat`` when the private application was launched hidden
        with screen updates disabled.  Activating the sheet and recalculating
        once is materially cheaper than doing it per cell and keeps the COM
        instance private to this task.
        """
        app = self._batch_excel
        if app is None:
            return
        try:
            app.ScreenUpdating = True
        except Exception:
            pass
        try:
            workbook.Activate()
        except Exception:
            pass
        try:
            sheet.Activate()
        except Exception:
            pass
        try:
            app.CalculateFullRebuild()
        except Exception:
            try:
                app.Calculate()
            except Exception:
                pass
        try:
            app.ScreenUpdating = False
        except Exception:
            pass

    def extract_rendered_conditional_formats(self, **kwargs: Any) -> list[Issue]:
        try:
            return super().extract_rendered_conditional_formats(**kwargs)
        except FormulaEngineError as exc:
            # WPS 12.0 returns None for DisplayFormat.Interior.Color even when
            # it has loaded a matching FormatCondition. Use the verified rule
            # model only in that unsupported-renderer case.
            if "实际显示颜色" not in str(exc):
                raise
            return self._extract_rule_based_conditional_formats(**kwargs)


class OfficeComCalculator:
    """Windows Office calculator with one deliberate Excel -> WPS fallback.

    The selected program remains fixed for the entire batch, so a workbook is
    never calculated partly by Excel and partly by WPS.
    """

    engine_name = ENGINE_OFFICE_COM

    def __init__(self, *, prefer_wps: bool = False) -> None:
        self._delegate: ExcelComCalculator | None = None
        self._adapters = (
            (WpsComCalculator, ExcelComCalculator)
            if prefer_wps else (ExcelComCalculator, WpsComCalculator)
        )

    def _choose(self) -> ExcelComCalculator:
        failures: list[str] = []
        for adapter in self._adapters:
            try:
                adapter.probe_available()
                return adapter()
            except Exception as exc:
                failures.append("{}：{}".format(adapter.application_label, exc))
        raise FormulaEngineError("未检测到可用的 Excel 或 WPS。" + "；".join(failures))

    def _active(self) -> ExcelComCalculator:
        if self._delegate is None:
            self._delegate = self._choose()
        return self._delegate

    def begin_batch(self, *, template_path: Path, external_path: Path | None = None) -> None:
        if self._delegate is not None:
            self._delegate.begin_batch(template_path=template_path, external_path=external_path)
            return
        failures: list[str] = []
        for adapter in self._adapters:
            candidate = adapter()
            try:
                adapter.probe_available()
                candidate.begin_batch(template_path=template_path, external_path=external_path)
                self._delegate = candidate
                return
            except Exception as exc:
                failures.append("{}：{}".format(adapter.application_label, exc))
                candidate.end_batch()
        raise FormulaEngineError("无法启动 Excel 或 WPS 批量会话：" + "；".join(failures))

    def end_batch(self) -> None:
        if self._delegate is not None:
            self._delegate.end_batch()

    def prepare_audit_copy(self, *args: Any, **kwargs: Any):
        return self._active().prepare_audit_copy(*args, **kwargs)

    def recalculate(self, *args: Any, **kwargs: Any) -> CalculationResult:
        return self._active().recalculate(*args, **kwargs)

    def extract_rendered_conditional_formats(self, *args: Any, **kwargs: Any):
        return self._active().extract_rendered_conditional_formats(*args, **kwargs)


def get_calculator(engine_name: str) -> Calculator:
    normalized = (engine_name or ENGINE_DEFAULT).strip()
    # Accept the two old saved values once, but always execute the automatic
    # selector. New settings only persist the single Excel/WPS option.
    if normalized in {ENGINE_OFFICE_COM, ENGINE_EXCEL_COM}:
        return OfficeComCalculator(prefer_wps=False)
    if normalized in {ENGINE_WPS_FIRST_COM, ENGINE_WPS_COM}:
        return OfficeComCalculator(prefer_wps=True)
    raise FormulaEngineError("Windows Flet 版只支持 Excel/WPS 或 WPS/Excel：{}".format(engine_name))


def engine_status() -> dict[str, object]:
    """Detect usable Office engines by actually starting their COM server."""
    try:
        version = ExcelComCalculator.probe_available()
        excel_com = {"available": True, "detail": "Microsoft Excel{}".format(" " + version if version else "")}
    except Exception as exc:
        excel_com = {"available": False, "detail": str(exc)}
    try:
        version = WpsComCalculator.probe_available()
        wps_com = {"available": True, "detail": "WPS{}".format(" " + version if version else "")}
    except Exception as exc:
        wps_com = {"available": False, "detail": str(exc)}
    return {"excelCom": excel_com, "wpsCom": wps_com}


def engine_available(engine_name: str, status: dict[str, object] | None = None) -> bool:
    """Return availability for the two COM engines exposed by the Windows edition.

    Stale settings/status payloads that still name removed engines resolve to
    ``False`` here, so they are never presented as part of the Windows product.
    """
    current = status or engine_status()
    availability = {
        ENGINE_OFFICE_COM: bool(dict(current.get("excelCom", {})).get("available"))
        or bool(dict(current.get("wpsCom", {})).get("available")),
        ENGINE_WPS_FIRST_COM: bool(dict(current.get("excelCom", {})).get("available"))
        or bool(dict(current.get("wpsCom", {})).get("available")),
    }
    return availability.get(engine_name, False)


def preferred_engine(status: dict[str, object] | None = None) -> str:
    """Choose one available engine at startup; never use runtime fallback."""
    current = status or engine_status()
    return ENGINE_OFFICE_COM
