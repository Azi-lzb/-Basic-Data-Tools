from __future__ import annotations

import argparse
import sys
from pathlib import Path


ROOT = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
HISTORY_WORKBOOK = ROOT / "历史审核说明.xlsx"
SRC = ROOT / "src"
if not getattr(sys, "frozen", False) and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from base_audit.settings import hide_application_data_directory


def main() -> int:
    hide_application_data_directory(ROOT / "data")
    parser = argparse.ArgumentParser(description="基础数据审核工具")
    parser.add_argument("--cli", action="store_true", help="使用命令行模式")
    parser.add_argument("--template", type=Path, help="审核模板")
    parser.add_argument("--input", type=Path, help="报送文件目录")
    parser.add_argument("--output", type=Path, help="审核输出目录")
    parser.add_argument("--period", help="数据期，例如 2026-06")
    parser.add_argument("--external", type=Path, help="外部辅助文件；审核时复制其工作表到审核副本")
    parser.add_argument("--init-config", action="store_true", help="迁移流程配置并保留历史审核说明.xlsx")
    parser.add_argument(
        "--check-only", action="store_true", help="只做模板体检和文件匹配"
    )
    parser.add_argument(
        "--region-summary", action="store_true", help="执行区域配置中的任意行/固定行汇总"
    )
    parser.add_argument(
        "--history",
        type=Path,
        default=None,
        help="历史审核库",
    )
    args = parser.parse_args()

    if args.init_config:
        from base_audit.name_config import initialize_config
        path = initialize_config(HISTORY_WORKBOOK)
        print(f"配置已初始化：{path}（流程配置已迁出；历史审核结果未改动）")
        return 0

    if args.cli:
        from base_audit.name_config import initialize_config
        initialize_config(HISTORY_WORKBOOK)

    if args.cli and args.region_summary:
        missing = [name for name, value in (("--template", args.template), ("--input", args.input), ("--output", args.output)) if not value]
        if missing:
            parser.error("区域汇总缺少参数：" + "、".join(missing))
        from base_audit.service import AuditService
        result = AuditService(config_path=HISTORY_WORKBOOK).summarize_regions(
            template_path=args.template, input_dir=args.input, output_dir=args.output
        )
        print(result.summary_text())
        return 0

    if args.cli:
        missing = [
            name
            for name, value in (
                ("--template", args.template),
                ("--input", args.input),
                ("--output", args.output),
            )
            if not value
        ]
        if not args.check_only and not args.period:
            missing.append("--period")
        if missing:
            parser.error("命令行模式缺少参数：" + "、".join(missing))
        if args.external and not args.external.is_file():
            parser.error("外部文件不存在：" + str(args.external))
        from base_audit.service import AuditService

        if args.external:
            print(f"外部辅助文件：{args.external.resolve()}（将复制到审核副本）")
        if args.check_only:
            result = AuditService(config_path=HISTORY_WORKBOOK).preflight(
                template_path=args.template,
                input_dir=args.input,
                output_dir=args.output,
                external_path=args.external,
            )
            print(result.summary_text())
            return 0 if result.failed_files == 0 else 2
        result = AuditService(config_path=HISTORY_WORKBOOK).run(
            template_path=args.template,
            input_dir=args.input,
            output_dir=args.output,
            period=args.period,
            history_path=args.history or HISTORY_WORKBOOK,
            external_path=args.external,
        )
        print(result.summary_text())
        return 0 if result.failed_files == 0 else 2

    from base_audit.single_instance import acquire, show_already_running_message
    if not acquire():
        show_already_running_message()
        return 0
    from base_audit.web_app import launch_web

    launch_web(ROOT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
