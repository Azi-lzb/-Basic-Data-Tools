"""计算引擎适配层：按操作系统提供可用引擎清单与提示。

业务层与前端只通过本模块（及 ``protocol.Calculator``）认识引擎；
不得在业务代码中直接判断 COM、UNO 或操作系统。
"""

from __future__ import annotations

import sys

from .excel_wps import ENGINE_AUTO, ENGINE_EXCEL, ENGINE_WPS, com_available
from .libreoffice_adapter import LibreOfficeAdapter
from .protocol import CalculationResult, Calculator, scan_formula_errors

ENGINE_LIBREOFFICE = "LibreOffice Calc"


def available_engines() -> list[dict[str, str]]:
    """当前环境真正可用的计算引擎选项（启动时按操作系统识别）。

    Windows 使用 Microsoft Excel / WPS 表格 COM；统信 UOS / 麒麟使用
    LibreOffice Calc。界面只显示当前环境可用的引擎。
    """
    if sys.platform == "win32":
        return [
            {"value": ENGINE_AUTO, "label": "自动（推荐）"},
            {"value": ENGINE_EXCEL, "label": ENGINE_EXCEL},
            {"value": ENGINE_WPS, "label": ENGINE_WPS},
        ]
    return [
        {"value": ENGINE_AUTO, "label": "自动（推荐）"},
        {"value": ENGINE_LIBREOFFICE, "label": ENGINE_LIBREOFFICE},
    ]


def engine_hint() -> str:
    if sys.platform == "win32":
        return (
            "自动模式优先使用 Microsoft Excel；无法启动时尝试 WPS 表格。"
            "指定引擎后，程序不会再自动切换，适合排查公式兼容性问题。"
        )
    return "统信 UOS / 麒麟环境使用系统安装的 LibreOffice Calc 计算公式和渲染条件格式。"


def valid_engine_values() -> set[str]:
    return {item["value"] for item in available_engines()}


def pipeline_kind(engine_preference: str = ENGINE_AUTO) -> str:
    """当前环境下应使用的审核管线类型，业务层只问结果不判系统。

    - ``"com"``：Windows Excel/WPS 会话（excel_com.ExcelSession 全功能门面）；
    - ``"native"``：openpyxl + soffice 重算的原生管线（native 包）。
    """
    return "com" if sys.platform == "win32" else "native"


def probe_engines() -> list[dict[str, str]]:
    """逐项探测引擎真实可用性；返回 value/available/detail 供诊断界面使用。"""
    status: list[dict[str, str]] = []
    if sys.platform == "win32":
        for preference in (ENGINE_EXCEL, ENGINE_WPS):
            ok, detail = com_available(preference)
            status.append({"value": preference, "available": "1" if ok else "0", "detail": detail})
    else:
        ok, detail = LibreOfficeAdapter().available()
        status.append({"value": ENGINE_LIBREOFFICE, "available": "1" if ok else "0", "detail": detail})
    return status


__all__ = [
    "CalculationResult",
    "Calculator",
    "scan_formula_errors",
    "available_engines",
    "engine_hint",
    "valid_engine_values",
    "pipeline_kind",
    "probe_engines",
]
