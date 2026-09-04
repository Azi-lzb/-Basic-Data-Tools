"""LibreOffice Calc 适配器（UOS/麒麟/Linux）。

包装 ``engines.libreoffice.LibreOfficeCalculator``，使其符合
``engines.protocol.Calculator`` 协议。完整审核流程（模板体检、公式复制、
条件格式渲染读取）接入统一核心属于二期工作，须在真实 UOS 环境验收。
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

from .libreoffice import LibreOfficeCalculator
from .protocol import CalculationResult, scan_formula_errors


class LibreOfficeAdapter:
    engine_name = "LibreOffice Calc"

    def __init__(self) -> None:
        self._calculator = LibreOfficeCalculator()

    def available(self) -> tuple[bool, str]:
        try:
            engine = self._calculator.require_available()
            return True, "{}（{}；{}）".format(engine.display, engine.source, self._calculator.version(engine))
        except Exception as exc:
            return False, str(exc)

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        result = self._calculator.recalculate(workbook_path)
        return CalculationResult(
            result.workbook_path,
            self.engine_name,
            result.elapsed_seconds,
            error_cells=scan_formula_errors(result.workbook_path, formula_ranges=formula_ranges),
        )
