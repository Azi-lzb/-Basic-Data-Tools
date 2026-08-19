"""Benchmark Excel calculation strategies on audit copies only.

Usage example:
  python tools/benchmark_calculation.py --template <template.xlsx> --input <source-dir> --external <external.xlsx> --config <config.xlsx> --limit 1
"""

from __future__ import annotations

import argparse
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.excel_com import ExcelSession, XL_CALCULATION_DONE
from base_audit.service import _external_sheet_plan, _organisation_from_name


def wait_calculation(excel: ExcelSession, timeout: float = 120) -> None:
    deadline = time.monotonic() + timeout
    while excel.excel.CalculationState != XL_CALCULATION_DONE:
        if time.monotonic() >= deadline:
            raise TimeoutError("Excel 公式计算超过 120 秒")
        time.sleep(0.1)


def calculate(excel: ExcelSession, workbook: object, strategy: str) -> None:
    if strategy == "普通计算":
        excel.calculate_pending_workbooks()
        return
    if strategy == "逐表计算":
        for index in range(1, workbook.Worksheets.Count + 1):
            workbook.Worksheets(index).Calculate()
        wait_calculation(excel)
        return
    if strategy == "全量重建":
        excel.excel.CalculateFullRebuild()
        wait_calculation(excel)
        return
    if strategy == "仅保存":
        return
    raise ValueError(f"未知策略：{strategy}")


def main() -> int:
    parser = argparse.ArgumentParser(description="基础数据审核计算方式性能对比")
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--external", type=Path, required=True)
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--limit", type=int, default=1)
    parser.add_argument("--strategies", default="普通计算,逐表计算,全量重建,仅保存")
    args = parser.parse_args()

    sources = sorted(
        path for path in args.input.glob("*.xlsx")
        if not path.name.startswith(("~$", "!", "！")) and "_审核版" not in path.stem
    )[: args.limit]
    if not sources:
        raise ValueError("没有可测试的源文件")
    strategies = [item.strip() for item in args.strategies.split(",") if item.strip()]
    result_dir = ROOT / "tests" / "performance_results" / datetime.now().strftime("%Y%m%d%H%M%S")
    result_dir.mkdir(parents=True, exist_ok=True)
    rows: list[tuple[object, ...]] = []

    for strategy in strategies:
        with ExcelSession() as excel:
            template = excel.open_workbook(args.template, read_only=True)
            external = excel.open_workbook(args.external, read_only=True)
            try:
                definition = excel.read_template(template, config_path=args.config)
                plan = _external_sheet_plan(
                    excel, template, definition, external,
                )
                for source in sources:
                    audit_path = result_dir / f"{strategy}_{source.stem}_审核版.xlsx"
                    started = time.monotonic()
                    workbook = None
                    try:
                        shutil.copy2(source, audit_path)
                        workbook = excel.open_workbook(audit_path, read_only=False)
                        copy_started = time.monotonic()
                        excel.copy_external_sheets(external, workbook, plan.sheet_names)
                        excel.apply_rules(template, workbook, definition, calculate=False)
                        copy_seconds = time.monotonic() - copy_started
                        calc_started = time.monotonic()
                        calculate(excel, workbook, strategy)
                        calc_seconds = time.monotonic() - calc_started
                        issues = excel.extract_issues(
                            workbook, definition.rules, period="benchmark", batch_id="benchmark",
                            audit_time="", org_code="", org_name="",
                            source_file=source, audit_file=audit_path,
                            extraction_ranges=definition.extraction_ranges,
                        )
                        save_started = time.monotonic()
                        workbook.Save()
                        save_seconds = time.monotonic() - save_started
                        rows.append((strategy, source.name, "成功", round(time.monotonic() - started, 2), round(copy_seconds, 2), round(calc_seconds, 2), round(save_seconds, 2), len(issues), ""))
                    except Exception as exc:
                        rows.append((strategy, source.name, "失败", round(time.monotonic() - started, 2), "", "", "", "", str(exc)))
                    finally:
                        if workbook is not None:
                            excel.close_workbook(workbook)
            finally:
                excel.close_workbook(external)
                excel.close_workbook(template)

    report = Workbook()
    sheet = report.active
    sheet.title = "计算方式对比"
    sheet.append(("策略", "源文件", "结果", "总耗时(秒)", "复制耗时(秒)", "计算耗时(秒)", "保存耗时(秒)", "问题数", "说明"))
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:I{max(2, sheet.max_row)}"
    for width, column in ((18, "A"), (56, "B"), (12, "C"), (14, "D"), (14, "E"), (14, "F"), (14, "G"), (12, "H"), (48, "I")):
        sheet.column_dimensions[column].width = width
    report_path = result_dir / "计算方式对比.xlsx"
    report.save(report_path)
    print(report_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
