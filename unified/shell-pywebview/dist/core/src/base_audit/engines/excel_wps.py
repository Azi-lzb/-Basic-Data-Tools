"""Microsoft Excel / WPS 表格 COM 适配器（仅 Windows）。

统一核心的完整审核流程（模板体检、公式复制、重算、问题提取、条件格式
渲染读取）由 ``excel_com.ExcelSession`` 门面承担；本模块只提供引擎可用性
探测与协议包装，pywin32/COM 延迟导入，不进入 Linux 发行包。
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterable

from .protocol import CalculationResult, scan_formula_errors

ENGINE_AUTO = "自动"
ENGINE_EXCEL = "Microsoft Excel"
ENGINE_WPS = "WPS 表格"


def com_available(engine_preference: str = ENGINE_AUTO) -> tuple[bool, str]:
    """探测当前 Windows 环境能否启动所选的 Excel/WPS COM 引擎。"""
    if sys.platform != "win32":
        return False, "当前操作系统不支持 Excel/WPS COM；统信 UOS / 麒麟请使用 LibreOffice Calc。"
    try:
        from ..excel_com import ExcelSession
    except Exception as exc:  # pywin32 缺失等
        return False, "Excel/WPS COM 组件不可用：{}".format(exc)
    try:
        with ExcelSession(engine_preference):
            return True, engine_preference if engine_preference != ENGINE_AUTO else ENGINE_EXCEL
    except Exception as exc:
        return False, str(exc)


class ExcelWpsCalculator:
    """按协议包装 ``ExcelSession`` 的最小重算适配器。

    批量审核仍直接使用 ``ExcelSession``（含公式复制与问题提取门面）；
    本适配器供独立重算场景与统一协议消费方使用。
    """

    def __init__(self, engine_preference: str = ENGINE_AUTO) -> None:
        self.engine_preference = engine_preference

    @property
    def engine_name(self) -> str:
        return self.engine_preference

    def recalculate(self, workbook_path: Path, *, formula_ranges: Iterable[object] = ()) -> CalculationResult:
        import time

        from ..excel_com import ExcelSession

        started = time.monotonic()
        with ExcelSession(self.engine_preference) as excel:
            workbook = excel.open_workbook(workbook_path, read_only=False)
            try:
                excel.excel.Calculate()
                engine_name = excel.engine_name or ENGINE_EXCEL
            finally:
                excel.close_workbook(workbook)
        return CalculationResult(
            workbook_path,
            engine_name,
            time.monotonic() - started,
            error_cells=scan_formula_errors(workbook_path, formula_ranges=formula_ranges),
        )
