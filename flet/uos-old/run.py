"""Command-line entry point for the native UOS/Linux audit edition."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from uos_audit.calculation import ENGINE_DEFAULT, ENGINE_OPTIONS, LibreOfficeAdapter, get_calculator
from uos_audit.audit import run_native_audit
from uos_audit.name_config import (
    FIXED_ROW_SUMMARY_FUNCTION, USED_RANGE_SUMMARY_FUNCTION,
    load_feature_mappings,
)
from uos_audit.region_summary import run_region_summaries


def main() -> int:
    parser = argparse.ArgumentParser(description="基础数据审核工具（统信/Linux 原生版）")
    parser.add_argument("--check-engine", action="store_true", help="检查 LibreOffice Calc 兼容计算引擎")
    parser.add_argument("--flet", action="store_true", help="启动统信 Flet 工作台（默认）")
    parser.add_argument("--recalculate", type=Path, help="计算并保存指定的 .xlsx 审核副本")
    parser.add_argument("--audit", action="store_true", help="运行统信原生公式审核主链路")
    parser.add_argument("--summary", action="store_true", help="运行统信原生任意行/固定行汇总")
    parser.add_argument("--flow", help="按流程配置执行指定流程名")
    parser.add_argument("--make-union-template", action="store_true", help="以底稿模板制作联合模板")
    parser.add_argument("--base-template", type=Path, help="联合模板的底稿模板")
    parser.add_argument("--join-template", action="append", type=Path, default=[], help="要并入的模板，可重复填写")
    parser.add_argument("--input-dir", type=Path, help="源数据目录")
    parser.add_argument("--template", type=Path, help="模板文件")
    parser.add_argument("--output-dir", type=Path, help="执行结果目录")
    parser.add_argument("--external", type=Path, help="可选的外部辅助工作簿")
    parser.add_argument("--history", type=Path, help="可选的历史审核说明.xlsx")
    parser.add_argument("--config", type=Path, help="可选的流程配置.json")
    parser.add_argument("--recursive", action="store_true", help="递归读取源数据子目录")
    parser.add_argument("--period", default="", help="可选审核期，例如 2026-08")
    parser.add_argument("--feature", action="append", default=[], help="汇总时仅执行指定功能名，可重复填写")
    parser.add_argument("--engine", choices=ENGINE_OPTIONS, default=ENGINE_DEFAULT, help="公式计算引擎（默认：LibreOffice Calc）")
    args = parser.parse_args()
    if args.flet or len(sys.argv) == 1:
        from uos_audit.flet_app import launch_flet
        launch_flet(ROOT)
        return 0
    if args.check_engine:
        available, detail = LibreOfficeAdapter().available()
        if not available:
            parser.error(detail)
        print("LibreOffice 兼容引擎可用：{}".format(detail))
        return 0
    if args.recalculate:
        result = get_calculator(args.engine).recalculate(args.recalculate)
        print("已通过 {} 计算并保存：{}（{:.1f} 秒）".format(result.engine_name, result.workbook_path, result.elapsed_seconds))
        return 0
    if args.audit:
        required = {
            "--input-dir": args.input_dir,
            "--template": args.template,
            "--output-dir": args.output_dir,
        }
        missing = [name for name, value in required.items() if value is None]
        if missing:
            parser.error("--audit 需要同时填写 " + "、".join(missing))
        result = run_native_audit(
            input_dir=args.input_dir, template_path=args.template, output_dir=args.output_dir,
            external_path=args.external, config_path=args.config, recursive=args.recursive,
            history_path=args.history, period=args.period, on_step=print, calculation_engine=args.engine,
        )
        print(result.summary_text())
        return 0
    if args.summary:
        required = {
            "--input-dir": args.input_dir,
            "--template": args.template,
            "--output-dir": args.output_dir,
        }
        missing = [name for name, value in required.items() if value is None]
        if missing:
            parser.error("--summary 需要同时填写 " + "、".join(missing))
        features = [
            item for item in load_feature_mappings(args.config, args.template)
            if item.feature_type in {USED_RANGE_SUMMARY_FUNCTION, FIXED_ROW_SUMMARY_FUNCTION}
            and (not args.feature or item.name in args.feature)
        ]
        output = run_region_summaries(
            template_path=args.template, input_dir=args.input_dir, output_dir=args.output_dir,
            features=features, recursive=args.recursive, on_step=print,
        )
        print("汇总完成：{}".format(output))
        return 0
    if args.flow:
        if args.input_dir is None or args.output_dir is None:
            parser.error("--flow 需要 --input-dir 与 --output-dir")
        from uos_audit.workflow import run_flow
        history = args.history or (ROOT / "历史审核说明.xlsx")
        config = args.config or history
        result = run_flow(flow_name=args.flow, template_path=args.template, input_dir=args.input_dir, output_dir=args.output_dir, history_path=history, config_path=config, external_path=args.external, recursive=args.recursive, period=args.period, on_step=print, calculation_engine=args.engine)
        print(result)
        return 0
    if args.make_union_template:
        if args.base_template is None or not args.join_template:
            parser.error("--make-union-template 需要 --base-template 和至少一个 --join-template")
        from uos_audit.merge_workbooks import create_combined_template
        output, copied, skipped = create_combined_template(base_template=args.base_template, source_templates=args.join_template, on_step=print)
        print("联合模板：{}；复制 {} 张表，跳过 {} 张同名表".format(output, len(copied), len(skipped)))
        return 0
    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
