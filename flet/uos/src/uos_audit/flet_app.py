"""Flet desktop workbench for the native UOS edition.

企业级 B 端管理平台风格：左侧深蓝导航 + 右侧浅灰主内容区。
设计规范见项目根目录 CLAUDE.md「UOS 前端设计规范」。
"""
from __future__ import annotations

import asyncio
from pathlib import Path

from .calculation import ENGINE_OPTIONS
from .name_config import COMBINE_SHEETS_FLOW, COMBINE_SHEETS_FUNCTION, RANGE_FREE_FEATURE_TYPES
from .web_app import WebApi


def launch_flet(root: Path) -> None:
    """Start the only graphical workbench supported by the UOS edition."""
    try:
        import flet as ft
    except ImportError as exc:  # pragma: no cover - exercised on UOS machine
        raise RuntimeError("未安装 Flet。请执行：python -m pip install -r requirements.txt") from exc

    def main(page: "ft.Page") -> None:
        api = WebApi(root)
        api._log("欢迎使用基础数据审核工具（统信版）。请先选择源数据目录。")
        page.title = "基础数据审核工具（统信版）"
        page.theme_mode = ft.ThemeMode.LIGHT
        # 全局界面统一采用宋体。Flutter/Flet 会在目标系统未安装该字体时由
        # 系统做字符级回退；UOS 发布环境可安装仿宋 GB2312 作为同类备用字库。
        # 不再把无衬线英文字体放在前面，以免中文随控件默认字体发生漂移。
        CHINESE_UI_FONT = "SimSun"
        page.theme = ft.Theme(font_family=CHINESE_UI_FONT)
        page.padding = 0
        page.bgcolor = "#F5F7FA"
        page.window.min_width = 1180
        page.window.min_height = 740

        # ---- 企业级 B 端配色 ----
        PRIMARY = "#1A3C6E"
        PRIMARY_HOVER = "#15305A"
        BG = "#F5F7FA"
        DANGER = "#E74C3C"
        SUCCESS = "#27AE60"
        WARN = "#F39C12"
        INFO = "#3498DB"
        TEXT = "#1F2A37"
        TEXT_2 = "#5B6B7C"
        TEXT_3 = "#98A6B5"
        BORDER = "#E4E9F0"
        CONSOLE_BG = "#1E1E1E"
        CONSOLE_TEXT = "#7CE38B"

        paths: dict[str, "ft.Text"] = {}
        file_rows = ft.ListView(expand=True, spacing=4)
        custom_actions = ft.Column(spacing=8)
        custom_flow_menu = None
        rendered_custom_flows: tuple[tuple[str, str], ...] | None = None
        # 日志只在有新内容时更新，避免监控刷新打断用户拖选复制；
        # auto_scroll 会在新增记录时自动把最新行带入可视区域。
        log_view = ft.ListView(expand=True, spacing=2, auto_scroll=True)
        rendered_log_text: str | None = None
        status = ft.Text(api.state["status"], color=TEXT_2, size=12)
        busy = ft.ProgressRing(width=16, height=16, stroke_width=2, visible=False)
        output_pin = ft.IconButton(icon=ft.Icons.PUSH_PIN_OUTLINED, tooltip="固定输出目录")
        engine_note = ft.Text(size=11, color=TEXT_2)
        engine_bar = ft.Row(spacing=12)

        def display_name(value: str) -> str:
            value = str(value).strip()
            if not value:
                return "未选择"
            return value.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1]

        def sync_values() -> dict[str, str]:
            return {name: str(api.state.get(name) or "") for name in paths}

        def set_path(name: str, value: str) -> None:
            api.update({name: value})
            render()

        pickers: dict[str, object] = {}
        for name in ("input", "templateDir", "output"):
            picker = ft.FilePicker()
            page.services.append(picker)
            pickers[name] = picker
        for name in ("template", "external"):
            picker = ft.FilePicker()
            page.services.append(picker)
            pickers[name] = picker
        history_picker = ft.FilePicker()
        union_base_picker = ft.FilePicker()
        union_join_picker = ft.FilePicker()
        run_log_export_picker = ft.FilePicker()
        page.services.extend([history_picker, union_base_picker, union_join_picker, run_log_export_picker])

        async def choose(name: str, folder: bool = False) -> None:
            picker = pickers[name]
            if folder:
                value = await picker.get_directory_path(
                    dialog_title="选择" + {"input": "源数据目录", "templateDir": "模板文件目录", "output": "输出目录"}[name]
                )
            else:
                allowed_extensions = ["xlsx"] if name in {"template", "external"} else ["xlsx", "xls"]
                files = await picker.pick_files(allow_multiple=False, allowed_extensions=allowed_extensions)
                value = files[0].path if files else None
            if value:
                set_path(name, value)
                if name in {"input", "templateDir"}:
                    api.refresh_sources()
                render()

        def choose_handler(name: str, folder: bool = False):
            async def handler(event) -> None:
                await choose(name, folder)
            return handler

        async def choose_history(event=None) -> None:
            files = await history_picker.pick_files(
                allow_multiple=False, allowed_extensions=["xlsx"], dialog_title="选择历史审核说明"
            )
            if files:
                try:
                    api.update({"historyPath": files[0].path})
                    settings_history_path.value = str(api.history_path)
                    settings_history_path.tooltip = str(api.history_path)
                except Exception as exc:
                    page.snack_bar = ft.SnackBar(ft.Text("历史审核说明设置失败：{}".format(exc)), open=True)
                render()

        def toggle_pin(event) -> None:
            pinned = not bool(api.state["outputPinned"])
            values: dict[str, object] = {"outputPinned": pinned}
            if not pinned:
                current_input = str(api.state.get("input") or "")
                if current_input:
                    values["output"] = str(Path(current_input) / "执行结果")
            api.update(values)
            render()

        output_pin.on_click = toggle_pin

        def recognize(event=None) -> None:
            api.update(sync_values())
            api.recognize()
            render()

        def inspect_template_formula(event=None) -> None:
            api.update(sync_values())
            api.inspect_current_template()
            render()

        def select_all(event=None) -> None:
            selected = [str(item["path"]) for item in api.state.get("sourceFiles", [])]
            api.set_selected_files(selected)
            render()

        def select_none(event=None) -> None:
            api.set_selected_files([])
            render()

        def refresh_files(event=None) -> None:
            api.update(sync_values())
            api.recognize()
            render()

        def select_one(path: str, checked: bool) -> None:
            selected = set(api.state.get("selectedFiles", []))
            if checked:
                selected.add(path)
            else:
                selected.discard(path)
            api.set_selected_files(sorted(selected))
            render()

        def start(flow: str) -> None:
            api.update(sync_values())
            if not bool(api.state.get("confirmBeforeRun", True)):
                api.start_with_state(flow, selected_files=list(api.state.get("selectedFiles", [])))
                render()
                return
            open_run_confirmation(flow)
            render()

        pending_run_action = ""
        run_confirm_title = ft.Text("确认执行", size=16, weight=ft.FontWeight.W_700, color=TEXT)
        run_confirm_items = ft.Column(spacing=8, scroll=ft.ScrollMode.AUTO, height=280)

        def confirm_run(event=None) -> None:
            nonlocal pending_run_action
            action = pending_run_action
            run_confirm_dialog.open = False
            pending_run_action = ""
            if action:
                api.start_with_state(action, selected_files=list(api.state.get("selectedFiles", [])))
            render()

        def open_run_confirmation(action: str) -> None:
            nonlocal pending_run_action
            try:
                context = api.get_run_context(action)
            except Exception as exc:
                page.snack_bar = ft.SnackBar(ft.Text("无法读取流程输入：{}".format(exc)), open=True)
                page.update()
                return
            pending_run_action = action
            selected = [Path(item).name for item in context["selectedFiles"]]
            selected_text = "已选择 {} 个文件".format(len(selected))
            if selected:
                selected_text += "\n" + "、".join(selected[:12])
                if len(selected) > 12:
                    selected_text += "……另有 {} 个".format(len(selected) - 12)
            items = [("待处理对象", selected_text)]
            if context["templateRequired"]:
                items.append(("模板文件", context["template"] or "未选择"))
            if context["externalRequired"]:
                items.append(("外部辅助文件", context["external"] or "未选择"))
            items.append(("执行结果存放", context["output"] or "未选择"))
            run_confirm_title.value = "确认执行：{}".format(context["flow"])
            run_confirm_items.controls = [
                ft.Container(
                    content=ft.Column([
                        ft.Text(label, size=12, weight=ft.FontWeight.W_600, color=TEXT_2),
                        ft.Text(str(value), size=12, color=TEXT, selectable=True),
                    ], spacing=2),
                    bgcolor="#F5F7FA", border_radius=6, padding=8,
                )
                for label, value in items
            ]
            run_confirm_dialog.open = True
            page.update()

        run_confirm_dialog = ft.AlertDialog(
            modal=True, bgcolor="#FFFFFF", elevation=0,
            title=run_confirm_title,
            content=ft.Container(run_confirm_items, width=560),
            actions=[
                ft.TextButton("取消", on_click=lambda event: setattr(run_confirm_dialog, "open", False) or page.update()),
                ft.FilledButton("确认执行", on_click=confirm_run),
            ],
        )
        page.overlay.append(run_confirm_dialog)

        def execute_selected_custom_flow(event=None) -> None:
            flow_name = str(getattr(custom_flow_menu, "value", "") or "").strip()
            if not flow_name:
                page.snack_bar = ft.SnackBar(ft.Text("请先选择一个自定义流程"), open=True)
                page.update()
                return
            start("flow:" + flow_name)

        async def export_run_log(event=None) -> None:
            """Save the current run record; users can still select any text directly."""
            text = "\n".join(str(item.get("text", item)) for item in api.state.get("log", []))
            if not text:
                page.snack_bar = ft.SnackBar(ft.Text("当前没有可导出的运行记录"), open=True)
                page.update()
                return
            try:
                saved = await run_log_export_picker.save_file(
                    dialog_title="导出运行记录", file_name="运行记录.txt",
                    allowed_extensions=["txt"], src_bytes=text.encode("utf-8"),
                )
                if saved:
                    page.snack_bar = ft.SnackBar(ft.Text("运行记录已导出：{}".format(saved)), open=True)
            except Exception as exc:
                page.snack_bar = ft.SnackBar(ft.Text("导出运行记录失败：{}".format(exc)), open=True)
            page.update()

        async def make_union_template(event) -> None:
            base_files = await union_base_picker.pick_files(
                allow_multiple=False, allowed_extensions=["xlsx"],
                dialog_title="第 1 步：选择基准模板（原文件不会修改）",
            )
            if not base_files:
                return
            join_files = await union_join_picker.pick_files(
                allow_multiple=True, allowed_extensions=["xlsx"],
                dialog_title="第 2 步：选择待并入模板（不要重复选择基准模板）",
            )
            if not join_files:
                page.snack_bar = ft.SnackBar(ft.Text("未选择待并入模板"), open=True)
                page.update()
                return
            api.create_template_merge_from_paths(base_files[0].path, [item.path for item in join_files])
            render()

        # ================= 设置中心工作数据 =================
        sc: dict[str, object] = {
            "editor": {}, "modules": [], "flows": [], "display": {},
            "plans": [], "active_plan": "", "step_rows": [],
            "sel_module": "", "sel_flow": "", "sel_step": 0, "sel_plan": "",
            "empty_flows": [],
        }

        def thin_border() -> "ft.Border":
            return ft.Border(
                top=ft.BorderSide(1, "#D3DCE6"), right=ft.BorderSide(1, "#D3DCE6"),
                bottom=ft.BorderSide(1, "#D3DCE6"), left=ft.BorderSide(1, "#D3DCE6"),
            )

        def module_rows() -> list[dict[str, str]]:
            by_name: dict[str, dict[str, str]] = {}
            for row in sc["editor"].get("modules", []):
                by_name[str(row.get("功能名", ""))] = dict(row)
            for row in sc["modules"]:
                name = str(row.get("功能名", ""))
                if name:
                    by_name[name] = dict(row)
            return list(by_name.values())

        def flow_rows() -> list[dict[str, str]]:
            # 一条流程有多条步骤，不能用“流程名 -> 单行”字典汇总；否则后
            # 写入的步骤会覆盖前面的步骤，设置中心只会显示最后一条。
            all_rows = [
                *(dict(row) for row in sc["editor"].get("defaultFlows", [])),
                *(dict(row) for row in sc["flows"]),
            ]
            # 新建但尚未添加步骤的流程只在会话内存在：没有步骤行就无法落库，
            # 先作为“空流程”加入列表，保证点击“新增流程”后立刻能看到。
            for name in sc["empty_flows"]:
                name = str(name).strip()
                if name and not any(str(r.get("流程名", "")).strip() == name for r in all_rows):
                    all_rows.append({"流程名": name})
            first_by_name: dict[str, dict[str, str]] = {}
            for row in all_rows:
                name = str(row.get("流程名", "")).strip()
                if name and name not in first_by_name:
                    first_by_name[name] = row
            ordered: list[dict[str, str]] = []
            for name in sc["editor"].get("defaultFlowNames", []):
                if name in first_by_name:
                    ordered.append(first_by_name.pop(name))
            ordered.extend(first_by_name.values())
            return ordered

        def _order_int(row: dict[str, str]) -> int:
            try:
                return int(float(str(row.get("顺序", "0") or "0")))
            except (TypeError, ValueError):
                return 0

        def sync_steps_to_flows() -> None:
            flow = str(sc["sel_flow"] or "")
            if not flow:
                return
            keep = [r for r in sc["flows"] if str(r.get("流程名", "")) != flow]
            sc["flows"] = [*keep, *(dict(r) for r in sc["step_rows"])]
            # 流程有步骤后即为真实记录，移出“空流程”清单；自定义流程没有任何
            # 步骤时继续保留在会话级空流程清单里，以便列表仍能显示它。
            default_names = set(sc["editor"].get("defaultFlowNames", []))
            if sc["step_rows"]:
                sc["empty_flows"] = [n for n in sc["empty_flows"] if str(n) != flow]
            elif flow not in default_names and flow not in sc["empty_flows"]:
                sc["empty_flows"].append(flow)

        # ---- 常规设置页控件 ----
        settings_engine = ft.Dropdown(
            label="公式计算引擎", value=str(api.state["calculationEngine"]), width=280, disabled=True,
            options=[ft.DropdownOption(key=name, text=name) for name in ENGINE_OPTIONS],
        )
        settings_logs = ft.Checkbox(label="输出流程运行日志", value=bool(api.state["writeFlowLogs"]))
        settings_confirm = ft.Checkbox(label="执行前确认输入对象", value=bool(api.state.get("confirmBeforeRun", True)))
        settings_custom = ft.Checkbox(label="显示自定义流程按钮", value=bool(api.state["showCustomFeatures"]))
        settings_recursive = ft.Checkbox(label="递归子文件夹", value=bool(api.state["recursive"]))
        engine_detail = ft.Text(size=12, color=TEXT_2)
        settings_history_path = ft.TextField(
            label="历史审核说明", value=str(api.history_path), read_only=True, expand=True,
            tooltip=str(api.history_path),
        )

        def engine_status_text() -> str:
            engines = api.state.get("engineStatus", {})
            libre_info = engines.get("libreoffice", {}) if isinstance(engines, dict) else {}
            detail = str(libre_info.get("detail", "未检测"))
            return "LibreOffice Calc：{}".format(detail)

        def update_engine_status() -> None:
            engine_detail.value = engine_status_text()

        # ---- 模块说明页 ----
        catalog_list = ft.ListView(expand=True, spacing=2)
        catalog_detail = ft.Column(spacing=6, expand=True, scroll=ft.ScrollMode.AUTO)

        def show_catalog(row: dict[str, str]) -> None:
            catalog_detail.controls = [
                ft.Text(str(row.get("功能名", "")), size=16, weight=ft.FontWeight.W_700, color=TEXT),
                ft.Text("执行模块：{}".format(row.get("执行模块", "")), size=13),
                ft.Text("命名区域名：{}".format(row.get("命名区域名", "") or "（无）"), size=13),
                ft.Text("输入：{}".format(row.get("输入", "")), size=13),
                ft.Text("输出：{}".format(row.get("输出", "")), size=13),
                ft.Text("备注：{}".format(row.get("备注", "")), size=13),
            ]
            page.update()

        def render_catalog() -> None:
            catalog_list.controls = []
            groups: dict[str, list[dict[str, str]]] = {}
            for row in module_rows():
                groups.setdefault(str(row.get("执行模块", "")), []).append(row)
            for ftype in sc["editor"].get("featureTypes", []):
                rows = groups.get(str(ftype), [])
                if not rows:
                    continue
                catalog_list.controls.append(ft.Text(str(ftype), size=12, weight=ft.FontWeight.W_600, color=TEXT_2))
                for row in rows:
                    catalog_list.controls.append(ft.ListTile(
                        title=ft.Text(str(row.get("功能名", ""))),
                        dense=True,
                        on_click=lambda e, r=row: show_catalog(r),
                    ))
            page.update()

        # ---- 功能配置页 ----
        module_list = ft.ListView(expand=True, spacing=2)
        module_name = ft.TextField(label="功能名", dense=True, expand=True)
        module_type = ft.Dropdown(label="执行模块", dense=True, expand=True)
        module_range = ft.TextField(label="命名区域名", dense=True, expand=True)
        module_input = ft.TextField(label="输入", dense=True, expand=True)
        module_output = ft.TextField(label="输出", dense=True, expand=True)
        module_note = ft.TextField(label="备注", dense=True, multiline=True, min_lines=2, max_lines=4, expand=True)
        module_error = ft.Text(size=12, color=DANGER)

        def default_io_for_type(ftype: str) -> tuple[str, str]:
            # 按“执行模块”类型取内置模块的默认输入/输出，新建功能时预填供参考。
            for row in sc["editor"].get("modules", []):
                if str(row.get("执行模块", "")) == ftype:
                    return str(row.get("输入", "") or ""), str(row.get("输出", "") or "")
            return "", ""

        def on_module_type_change(event=None) -> None:
            inp, outp = default_io_for_type(str(module_type.value or ""))
            module_input.value = inp
            module_output.value = outp
            page.update()

        module_type.on_change = on_module_type_change

        def select_module(row: dict[str, str]) -> None:
            sc["sel_module"] = str(row.get("功能名", ""))
            module_name.value = str(row.get("功能名", ""))
            module_type.value = str(row.get("执行模块", ""))
            module_range.value = str(row.get("命名区域名", ""))
            module_input.value = str(row.get("输入", ""))
            module_output.value = str(row.get("输出", ""))
            module_note.value = str(row.get("备注", ""))
            module_error.value = ""
            plans_section.visible = str(row.get("执行模块", "")) == COMBINE_SHEETS_FUNCTION
            if plans_section.visible:
                render_plans()
            render_module_list()
            page.update()

        def render_module_list() -> None:
            module_list.controls = []
            groups: dict[str, list[dict[str, str]]] = {}
            for row in module_rows():
                groups.setdefault(str(row.get("执行模块", "")), []).append(row)
            for ftype in sc["editor"].get("featureTypes", []):
                rows = groups.get(str(ftype), [])
                if not rows:
                    continue
                module_list.controls.append(ft.Text(str(ftype), size=12, weight=ft.FontWeight.W_600, color=TEXT_2))
                for row in rows:
                    name = str(row.get("功能名", ""))
                    module_list.controls.append(ft.ListTile(
                        title=ft.Text(name, size=13),
                        dense=True,
                        selected=(name == sc["sel_module"]),
                        on_click=lambda e, r=row: select_module(r),
                    ))
            page.update()

        def save_module(event=None) -> None:
            name = str(module_name.value or "").strip()
            if not name:
                module_error.value = "请填写功能名"
                page.update()
                return
            row: dict[str, str] = {
                "功能名": name,
                "执行模块": str(module_type.value or ""),
                "命名区域名": str(module_range.value or "").strip(),
                "输入": str(module_input.value or "").strip(),
                "输出": str(module_output.value or "").strip(),
                "备注": str(module_note.value or "").strip(),
            }
            if row["执行模块"] in RANGE_FREE_FEATURE_TYPES:
                row["命名区域名"] = ""
            replaced = False
            for index, existing in enumerate(sc["modules"]):
                if str(existing.get("功能名", "")) == name:
                    sc["modules"][index] = row
                    replaced = True
                    break
            if not replaced:
                sc["modules"].append(row)
            sc["sel_module"] = name
            module_error.value = ""
            render_module_list()
            page.update()

        def new_module(event=None) -> None:
            module_name.value = ""
            module_type.value = str(sc["editor"].get("featureTypes", [""])[0]) if sc["editor"].get("featureTypes") else ""
            module_range.value = ""
            module_note.value = ""
            module_error.value = ""
            inp, outp = default_io_for_type(str(module_type.value or ""))
            module_input.value = inp
            module_output.value = outp
            page.update()

        def copy_module(event=None) -> None:
            base = str(module_name.value or "").strip()
            module_name.value = (base + " 副本") if base else "新功能"
            save_module()
            page.update()

        def delete_module(event=None) -> None:
            name = sc["sel_module"]
            sc["modules"] = [r for r in sc["modules"] if str(r.get("功能名", "")) != name]
            rows = module_rows()
            if rows:
                select_module(rows[0])
            else:
                sc["sel_module"] = ""
                module_name.value = ""
                module_error.value = ""
            render_module_list()
            page.update()

        # ---- 组合工作表分组方案 ----
        plan_list = ft.ListView(expand=True, spacing=2, height=120)
        plan_name_field = ft.TextField(label="方案名称", dense=True, width=200)
        plan_mode = ft.Dropdown(label="分组方式", dense=True, width=140, options=[
            ft.DropdownOption(key="regex", text="正则表达式"),
            ft.DropdownOption(key="name", text="按名称分组"),
        ])
        plan_pattern = ft.TextField(label="正则表达式（regex 模式填写）", dense=True, expand=True)
        plan_groups = ft.TextField(
            label="按名称分组规则（每行：组合名称=关键字1、关键字2）",
            dense=True, multiline=True, min_lines=3, max_lines=6, expand=True,
        )
        plan_error = ft.Text(size=12, color=DANGER)

        def format_groups(groups: object) -> str:
            lines: list[str] = []
            for item in groups or []:
                if isinstance(item, dict):
                    lines.append("{}= {}".format(item.get("name", ""), item.get("keywords", "")))
            return "\n".join(lines)

        def parse_groups(text: object) -> list[dict[str, str]]:
            groups: list[dict[str, str]] = []
            for line in str(text or "").splitlines():
                line = line.strip()
                if not line:
                    continue
                if "=" in line:
                    name, keywords = line.split("=", 1)
                else:
                    name, keywords = line, ""
                groups.append({"name": name.strip(), "keywords": keywords.strip()})
            return groups

        def render_plans() -> None:
            plan_list.controls = []
            for plan in sc["plans"]:
                name = str(plan.get("name", ""))
                active = str(plan.get("id", "")) == str(sc["active_plan"])
                plan_list.controls.append(ft.ListTile(
                    title=ft.Text(name + ("（当前启用）" if active else ""), size=13),
                    dense=True,
                    selected=active,
                    on_click=lambda e, p=plan: select_plan(p),
                ))
            page.update()

        def select_plan(plan: dict[str, object]) -> None:
            sc["sel_plan"] = str(plan.get("id", ""))
            plan_name_field.value = str(plan.get("name", ""))
            plan_mode.value = str(plan.get("mode", "regex"))
            plan_pattern.value = str(plan.get("pattern", ""))
            plan_groups.value = format_groups(plan.get("groups", []))
            plan_error.value = ""
            render_plans()
            page.update()

        def new_plan(event=None) -> None:
            n = len(sc["plans"]) + 1
            plan_id = "plan{}".format(n)
            while any(str(p.get("id", "")) == plan_id for p in sc["plans"]):
                n += 1
                plan_id = "plan{}".format(n)
            plan = {"id": plan_id, "name": "新分组方案", "mode": "regex", "pattern": "", "groups": []}
            sc["plans"].append(plan)
            if not sc["active_plan"]:
                sc["active_plan"] = plan_id
            select_plan(plan)
            page.update()

        def save_plan(event=None) -> None:
            name = str(plan_name_field.value or "").strip()
            plan_id = str(sc["sel_plan"] or "")
            if not name:
                plan_error.value = "方案名称不能为空"
                page.update()
                return
            if not plan_id:
                plan_error.value = "请先选择一个方案"
                page.update()
                return
            plan = {
                "id": plan_id,
                "name": name,
                "mode": str(plan_mode.value or "regex"),
                "pattern": str(plan_pattern.value or "").strip(),
                "groups": parse_groups(plan_groups.value),
            }
            for index, item in enumerate(sc["plans"]):
                if str(item.get("id", "")) == plan_id:
                    sc["plans"][index] = plan
                    break
            plan_error.value = ""
            render_plans()
            page.update()

        def delete_plan(event=None) -> None:
            plan_id = str(sc["sel_plan"] or "")
            if not plan_id:
                return
            sc["plans"] = [p for p in sc["plans"] if str(p.get("id", "")) != plan_id]
            if str(sc["active_plan"]) == plan_id:
                sc["active_plan"] = str(sc["plans"][0].get("id", "")) if sc["plans"] else ""
            if sc["plans"]:
                select_plan(sc["plans"][0])
            else:
                sc["sel_plan"] = ""
                plan_name_field.value = ""
                plan_error.value = ""
            render_plans()
            page.update()

        def set_active_plan(event=None) -> None:
            plan_id = str(sc["sel_plan"] or "")
            if not plan_id:
                return
            sc["active_plan"] = plan_id
            render_plans()
            page.update()

        plans_section = ft.Column([
            ft.Divider(height=1),
            ft.Text("组合工作表分组方案", size=13, weight=ft.FontWeight.W_600, color=TEXT),
            plan_list,
            ft.Row([plan_name_field, plan_mode], spacing=8),
            plan_pattern,
            plan_groups,
            ft.Row([
                ft.OutlinedButton("新建方案", on_click=new_plan),
                ft.FilledButton("保存方案", on_click=save_plan),
                ft.OutlinedButton("删除方案", on_click=delete_plan),
                ft.OutlinedButton("设为启用", on_click=set_active_plan),
            ], spacing=6),
            plan_error,
        ], spacing=8, visible=False)

        # ---- 流程编排页 ----
        # Keep scrollable lists on explicit heights.  Flet desktop can render
        # a plain grey placeholder when an expanding ListView sits inside two
        # further expanding Rows/Containers in an AlertDialog.
        flow_list = ft.Column(spacing=2, scroll=ft.ScrollMode.AUTO, expand=True)
        steps_title = ft.Text("步骤", size=12, weight=ft.FontWeight.W_600, color=TEXT_2)
        steps_view = ft.Column(spacing=4, scroll=ft.ScrollMode.AUTO, expand=True)
        # 步骤编辑表单：放进“添加/编辑步骤”弹窗里，由用户选择功能并填写输入输出，
        # 不再默认自动生成一条步骤。同一行里的两个控件用 expand 均分宽度。
        step_flow = ft.TextField(label="流程名", dense=True, disabled=True, expand=True)
        step_order = ft.TextField(label="顺序", dense=True, expand=True)
        step_feature = ft.Dropdown(label="功能名", dense=True, expand=True)
        step_enable = ft.Dropdown(label="启用", dense=True, expand=True, options=[
            ft.DropdownOption(key="是", text="是"), ft.DropdownOption(key="否", text="否"),
        ])
        step_fail = ft.Dropdown(label="失败后处理", dense=True, expand=True, options=[
            ft.DropdownOption(key="停止", text="停止"), ft.DropdownOption(key="跳过", text="跳过"),
        ])
        step_output = ft.Dropdown(label="是否输出结果", dense=True, expand=True, options=[
            ft.DropdownOption(key="是", text="是"), ft.DropdownOption(key="否", text="否"),
        ])
        step_outfile = ft.TextField(label="输出文件名", dense=True)
        step_input = ft.TextField(label="输入", dense=True, expand=True)
        step_output2 = ft.TextField(label="输出", dense=True, expand=True)
        step_note = ft.TextField(label="备注", dense=True)
        step_error = ft.Text(size=12, color=DANGER)

        def on_step_feature_change(event=None) -> None:
            # 选择功能后，按该功能在“功能配置”里维护的输入/输出预填，用户可再改。
            feature = str(step_feature.value or "")
            defaults = {str(r.get("功能名", "")): r for r in module_rows()}
            row = defaults.get(feature)
            if row:
                step_input.value = str(row.get("输入", "") or "")
                step_output2.value = str(row.get("输出", "") or "")
                page.update()

        step_feature.on_change = on_step_feature_change

        def select_step(index: int) -> None:
            open_step_editor("edit", index)

        def select_flow(name: str) -> None:
            sc["sel_flow"] = name
            default_names = set(sc["editor"].get("defaultFlowNames", []))
            # A saved edit to a built-in flow is stored as an override in
            # customFlows; otherwise show the pristine default definition.
            saved_rows = [dict(r) for r in sc["flows"] if str(r.get("流程名", "")) == name]
            source_rows = saved_rows or (sc["editor"].get("defaultFlows", []) if name in default_names else sc["flows"])
            rows = [dict(r) for r in source_rows if str(r.get("流程名", "")) == name]
            rows.sort(key=_order_int)
            sc["step_rows"] = rows
            sc["sel_step"] = 0
            sync_flow_display_switch()
            load_step_form()
            render_flow_list()
            render_steps()
            page.update()

        def render_flow_list() -> None:
            flow_list.controls = []
            default_names = set(sc["editor"].get("defaultFlowNames", []))
            rows = flow_rows()
            defaults = [r for r in rows if str(r.get("流程名", "")) in default_names]
            customs = [r for r in rows if str(r.get("流程名", "")) not in default_names]

            def header(text: str) -> None:
                flow_list.controls.append(ft.Text(text, size=12, weight=ft.FontWeight.W_600, color=TEXT_2))

            def item(row: dict[str, str]) -> None:
                name = str(row.get("流程名", ""))
                flow_list.controls.append(ft.ListTile(
                    title=ft.Text(name, size=13),
                    dense=True,
                    selected=(name == sc["sel_flow"]),
                    on_click=lambda e, n=name: select_flow(n),
                ))

            header("常用流程")
            for row in defaults:
                item(row)
            header("自定义流程")
            for row in customs:
                item(row)
            page.update()

        def render_steps() -> None:
            steps_title.value = "步骤：{}".format(str(sc["sel_flow"] or "未选择"))
            steps_view.controls = []
            rows = sc["step_rows"]
            if not sc["sel_flow"]:
                steps_view.controls.append(ft.Text("请先在左侧选择一个流程", size=12, color=TEXT_2))
            elif not rows:
                steps_view.controls.append(ft.Text("该流程暂无步骤；点击左侧「添加步骤」从功能库选择功能。", size=12, color=TEXT_2))
            for index, row in enumerate(rows):
                left = ft.Container(
                    content=ft.GestureDetector(
                        content=ft.Row([
                            ft.Text(str(row.get("顺序", "")), width=34, size=12),
                            ft.Text(str(row.get("功能名", "")), expand=True, size=12),
                            ft.Text(str(row.get("启用", "")), width=30, size=12),
                            ft.Text(str(row.get("失败后处理", "")), width=44, size=12),
                        ], spacing=4),
                        on_tap=lambda e, i=index: select_step(i),
                    ),
                    bgcolor="#EAF2FF" if index == sc["sel_step"] else None,
                    border_radius=6, padding=ft.Padding(8, 2, 8, 2), expand=True,
                )
                steps_view.controls.append(ft.Row([
                    left,
                    ft.IconButton(ft.Icons.ARROW_UPWARD, icon_size=15, tooltip="上移", on_click=lambda e, i=index, d=-1: move_step(i, d, e)),
                    ft.IconButton(ft.Icons.ARROW_DOWNWARD, icon_size=15, tooltip="下移", on_click=lambda e, i=index, d=1: move_step(i, d, e)),
                    ft.IconButton(ft.Icons.DELETE_OUTLINE, icon_size=15, tooltip="删除本步骤", on_click=delete_step),
                ], spacing=2))
            # The compact two-column workflow view owns the visible cards.
            steps_view.controls = [flow_step_card(index, row) for index, row in enumerate(rows)]
            flow_current_title.value = "当前流程：{}".format(sc["sel_flow"] or "未选择")
            flow_count.value = "共 {} 步".format(len(rows))
            sync_flow_display_switch()
            page.update()

        def load_step_form() -> None:
            rows = sc["step_rows"]
            index = sc["sel_step"]
            if not sc["sel_flow"]:
                step_error.value = "请先选择一个流程"
                return
            if 0 <= index < len(rows):
                row = rows[index]
                step_flow.value = str(row.get("流程名", ""))
                step_order.value = str(row.get("顺序", ""))
                step_feature.value = str(row.get("功能名", ""))
                step_enable.value = str(row.get("启用", "是"))
                step_fail.value = str(row.get("失败后处理", "停止"))
                step_output.value = str(row.get("是否输出结果", "否"))
                step_outfile.value = str(row.get("输出文件名", ""))
                step_input.value = str(row.get("输入", ""))
                step_output2.value = str(row.get("输出", ""))
                step_note.value = str(row.get("备注", ""))
            else:
                step_flow.value = str(sc["sel_flow"])
                step_order.value = ""
                step_feature.value = ""
                step_enable.value = "是"
                step_fail.value = "停止"
                step_output.value = "否"
                step_outfile.value = ""
                step_input.value = ""
                step_output2.value = ""
                step_note.value = ""
            step_error.value = ""

        def save_step(event=None) -> None:
            flow = str(sc["sel_flow"] or "")
            if not flow:
                step_error.value = "请先选择一个流程"
                page.update()
                return
            feature = str(step_feature.value or "")
            order = str(step_order.value or "").strip()
            if not feature or not order:
                step_error.value = "请填写功能名和顺序"
                page.update()
                return
            row: dict[str, str] = {
                "流程名": flow,
                "顺序": order,
                "功能名": feature,
                "启用": str(step_enable.value or "是"),
                "失败后处理": str(step_fail.value or "停止"),
                "是否输出结果": str(step_output.value or "否"),
                "输出文件名": str(step_outfile.value or "").strip(),
                "输入": str(step_input.value or "").strip(),
                "输出": str(step_output2.value or "").strip(),
                "备注": str(step_note.value or "").strip(),
            }
            rows = sc["step_rows"]
            index = sc["sel_step"]
            if 0 <= index < len(rows):
                edited = rows[index]
                edited.update(row)
            else:
                edited = row
                rows.append(edited)
            rows.sort(key=_order_int)
            sc["sel_step"] = next(i for i, r in enumerate(rows) if r is edited)
            sync_steps_to_flows()
            step_error.value = ""
            step_dialog.open = False
            render_steps()
            page.update()

        def open_step_editor(mode: str, index: int = -1) -> None:
            # “添加步骤”弹窗：先选功能库里的功能，再填写顺序/输入/输出等字段，
            # 由用户决定而不是默认生成。编辑既有步骤则复用同一弹窗。
            flow = str(sc["sel_flow"] or "")
            if not flow:
                step_error.value = "请先选择一个流程"
                page.update()
                return
            default_order = ""
            if mode == "add":
                sc["sel_step"] = len(sc["step_rows"])  # 越界 => load_step_form 按空表单初始化
                nxt = 10
                for row in sc["step_rows"]:
                    nxt = max(nxt, _order_int(row) + 10)
                default_order = str(nxt)
            else:
                sc["sel_step"] = index
            load_step_form()
            if default_order:
                step_order.value = default_order
            step_dialog.title.value = ("添加步骤" if mode == "add" else "编辑步骤") + "：{}".format(flow)
            step_error.value = ""
            step_dialog.open = True
            page.update()

        step_dialog = ft.AlertDialog(
            modal=True,
            bgcolor="#FFFFFF",
            elevation=0,
            title_padding=ft.Padding(20, 18, 20, 12),
            content_padding=ft.Padding(20, 8, 20, 8),
            actions_padding=ft.Padding(20, 10, 20, 16),
            title=ft.Text("添加步骤", size=16, weight=ft.FontWeight.W_700, color=TEXT),
            content=ft.Column([
                ft.Row([step_flow, step_order], spacing=12),
                ft.Row([step_feature, step_enable], spacing=12),
                ft.Row([step_fail, step_output], spacing=12),
                ft.Row([step_input, step_output2], spacing=12),
                step_outfile,
                step_note,
                step_error,
            ], width=480, spacing=10, horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
            actions=[
                ft.TextButton("取消", on_click=lambda event: setattr(step_dialog, "open", False) or page.update()),
                ft.FilledButton("保存步骤", on_click=lambda event: save_step(event)),
            ],
        )
        page.overlay.append(step_dialog)

        def move_step(index: int, delta: int, event=None) -> None:
            rows = sc["step_rows"]
            target = index + delta
            if not (0 <= target < len(rows)):
                return
            rows[index], rows[target] = rows[target], rows[index]
            sc["sel_step"] = target
            for i, row in enumerate(rows):
                row["顺序"] = str((i + 1) * 10)
            sync_steps_to_flows()
            render_steps()
            load_step_form()
            page.update()

        def delete_step(event=None) -> None:
            rows = sc["step_rows"]
            index = sc["sel_step"]
            if not (0 <= index < len(rows)):
                return
            rows.pop(index)
            sc["sel_step"] = max(0, min(index, len(rows) - 1))
            sync_steps_to_flows()
            render_steps()
            load_step_form()
            page.update()

        def new_flow(event=None) -> None:
            base = "自定义流程"
            name = base
            n = 2
            taken = {str(r.get("流程名", "")) for r in flow_rows()}
            while name in taken:
                name = "{}{}".format(base, n)
                n += 1
            sc["sel_flow"] = name
            sc["step_rows"] = []
            sc["sel_step"] = 0
            sc["display"][name] = False
            if name not in sc["empty_flows"]:
                sc["empty_flows"].append(name)
            load_step_form()
            render_flow_list()
            render_steps()
            page.update()

        def delete_flow(event=None) -> None:
            flow = str(sc["sel_flow"] or "")
            if not flow:
                return
            if flow in set(sc["editor"].get("defaultFlowNames", [])):
                step_error.value = "默认流程不能删除"
                page.update()
                return
            sc["flows"] = [r for r in sc["flows"] if str(r.get("流程名", "")) != flow]
            sc["display"].pop(flow, None)
            sc["empty_flows"] = [n for n in sc["empty_flows"] if str(n) != flow]
            sc["step_rows"] = []
            sc["sel_step"] = 0
            rows = flow_rows()
            sc["sel_flow"] = str(rows[0].get("流程名", "")) if rows else ""
            if sc["sel_flow"]:
                select_flow(sc["sel_flow"])
            render_steps()
            page.update()

        # ---- 字段与文件行控件 ----
        def open_path_field(name: str) -> None:
            value = str(api.state.get(name) or "")
            if not value:
                page.snack_bar = ft.SnackBar(ft.Text("尚未选择路径"), open=True)
                page.update()
                return
            api.open_path(value)

        def field(name: str, title: str, folder: bool = False, trailing=None, expand: bool = False, compact: bool = False) -> "ft.Control":
            # 路径以可点击文本展示：悬停变手型、点击用系统程序打开、可拖选复制。
            path_text = ft.Text(
                "未选择", size=12.5, color=TEXT_2, selectable=True,
                expand=True, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS,
            )
            path_box = ft.Container(
                content=path_text,
                bgcolor="#FFFFFF",
                border=ft.Border(
                    top=ft.BorderSide(1, "#D3DCE6"), right=ft.BorderSide(1, "#D3DCE6"),
                    bottom=ft.BorderSide(1, "#D3DCE6"), left=ft.BorderSide(1, "#D3DCE6"),
                ),
                border_radius=7, padding=7, expand=True, height=34,
                alignment=ft.Alignment(-1, 0),
            )
            path_field = ft.GestureDetector(
                content=path_box,
                # 精简入口没有“选择”按钮；点击路径框直接调出选择器。
                on_tap=choose_handler(name, folder) if compact else lambda event, n=name: open_path_field(n),
                mouse_cursor=ft.MouseCursor.CLICK,
                expand=True,
            )
            paths[name] = path_text
            if compact:
                return ft.Column([path_field], spacing=0, expand=expand,
                                 horizontal_alignment=ft.CrossAxisAlignment.STRETCH)
            compact_style = ft.ButtonStyle(padding=ft.Padding(8, 1, 8, 1))
            buttons = [ft.OutlinedButton("选择", on_click=choose_handler(name, folder), style=compact_style)]
            if trailing:
                buttons.insert(0, trailing)
            return ft.Column([
                ft.Row([
                    ft.Text(title, weight=ft.FontWeight.W_600, size=12, color=TEXT, expand=True),
                    *buttons,
                ], spacing=6),
                ft.Row([path_field], spacing=0),
            ], spacing=3, expand=expand, horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

        def _file_cell(index: int, item: dict[str, object], path: str, selected: set[str]) -> "ft.Control":
            # 单元格：上行文件名（勾选），下行匹配情况。
            suffix_note = "（.xls 仅支持读取/汇总）" if path.casefold().endswith(".xls") else ""
            label = "{:02d}. {}{}".format(index, item.get("name", ""), " " + suffix_note if suffix_note else "")
            template = str(item.get("template") or "").strip()
            matched = bool(template) and template != "未识别"
            match_text = "已匹配模板：" + template if matched else "未匹配模板"
            return ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Checkbox(
                            value=path in selected,
                            on_change=lambda event, p=path: select_one(p, bool(event.control.value)),
                        ),
                        # Checkbox 自带 label 在当前桌面端会单行截断；改为
                        # 独立小字号文本，让长文件名按卡片宽度自动换行。
                        ft.Text(label, size=11, color=TEXT, expand=True, selectable=True),
                    ], spacing=4, vertical_alignment=ft.CrossAxisAlignment.START),
                    ft.Row([
                        ft.Icon(ft.Icons.CHECK_CIRCLE if matched else ft.Icons.REMOVE_CIRCLE_OUTLINE, size=13, color=SUCCESS if matched else TEXT_3),
                        ft.Text(match_text, size=11, color=SUCCESS if matched else TEXT_2, selectable=True, tooltip=template),
                    ], spacing=4),
                ], spacing=2),
                bgcolor="#F8FAFC",
                border=thin_border(),
                border_radius=6,
                padding=8,
                expand=True,
            )

        def _file_grid(files: list[dict[str, object]], selected: set[str]) -> list["ft.Control"]:
            cells = [_file_cell(index, item, str(item["path"]), selected) for index, item in enumerate(files, start=1)]
            rows: list["ft.Control"] = []
            for i in range(0, len(cells), 2):
                pair = cells[i:i + 2]
                while len(pair) < 2:
                    pair.append(ft.Container(expand=True))
                rows.append(ft.Row([ft.Container(cell, expand=True) for cell in pair], spacing=8))
            return rows

        # ================= 主工作台（三张独立卡片） =================
        def card(content, *, expand: bool = False, padding: int = 20, height: int | None = None) -> "ft.Control":
            return ft.Container(
                content=content,
                bgcolor="#FFFFFF",
                border=ft.Border(
                    top=ft.BorderSide(1, BORDER), right=ft.BorderSide(1, BORDER),
                    bottom=ft.BorderSide(1, BORDER), left=ft.BorderSide(1, BORDER),
                ),
                border_radius=8,
                padding=12,
                expand=expand,
                height=height,
                shadow=ft.BoxShadow(blur_radius=8, spread_radius=0, offset=ft.Offset(0, 2), color="rgba(0,0,0,0.08)"),
            )

        # 上排两卡片：左源数据目录、右模板与目标
        source_card = card(ft.Column([
            field("input", "源数据目录", folder=True),
            ft.Text(
                "汇总核查表：按模板审核待处理文件\n"
                "汇总校验说明：汇总校验结果说明文件\n"
                "组合联合核查表：按机构合并多个报送工作簿\n"
                "制作联合模板：将多个模板工作表合并为联合模板",
                size=10.5, color=TEXT_2, selectable=True,
            ),
            ft.Row([
                ft.Container(expand=True),
                ft.FilledButton("设置中心", icon=ft.Icons.SETTINGS_OUTLINED,
                                on_click=lambda event: open_settings(event)),
            ], alignment=ft.MainAxisAlignment.END),
        ], spacing=7), expand=True, height=204)

        template_tool_style = ft.ButtonStyle(
            padding=ft.Padding(5, 1, 5, 1), text_style=ft.TextStyle(size=10),
        )
        template_tools = ft.Row([
            ft.TextButton("自动识别", on_click=recognize, style=template_tool_style),
            ft.TextButton("检查公式", on_click=inspect_template_formula, style=template_tool_style),
        ], spacing=1, tight=True)

        target_card = card(ft.Column([
            ft.Row([
                field("templateDir", "模板文件目录", folder=True, expand=True),
                field("template", "匹配模板", expand=True, trailing=template_tools),
            ], spacing=8),
            ft.Row([
                field("output", "输出目录", folder=True, trailing=output_pin, expand=True),
                field("external", "外部辅助文件", expand=True),
            ], spacing=8),
        ], spacing=7), expand=True, height=204)

        top_row = ft.Row([source_card, target_card], spacing=16,
                         vertical_alignment=ft.CrossAxisAlignment.STRETCH, height=204)

        # 下半区左栏：待处理文件纵向撑满，列表在卡片内部滚动，便于显示较长文件名。
        file_summary = ft.Text("共 0 张，匹配模板共 0 张", size=12, color=TEXT_2)
        files_card = card(ft.Column([
            ft.Row([
                ft.Text("待处理文件", size=15, weight=ft.FontWeight.W_700, color=TEXT),
                ft.Row([
                    file_summary,
                    ft.OutlinedButton("全选", on_click=select_all, style=ft.ButtonStyle(padding=ft.Padding(8, 3, 8, 3))),
                    ft.OutlinedButton("取消", on_click=select_none, style=ft.ButtonStyle(padding=ft.Padding(8, 3, 8, 3))),
                    ft.OutlinedButton("刷新", on_click=refresh_files, style=ft.ButtonStyle(padding=ft.Padding(8, 3, 8, 3))),
                ], spacing=6),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            file_rows,
        ], spacing=8, expand=True), expand=True)

        execution_section = ft.Column([
            ft.Row([
                ft.FilledButton(
                    "汇总核查表校验", icon=ft.Icons.FACT_CHECK_OUTLINED,
                    tooltip="操作步骤：\n1. 选择源数据目录；\n2. 选择模板文件目录，点击“自动识别”；\n3. 确认匹配模板，必要时选择外部辅助文件；\n4. 勾选待处理文件；\n5. 点击后确认执行。\n提示：需要扫描子目录时，请在设置中心勾选“递归子文件夹”。",
                    on_click=lambda event: start("audit"),
                    style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6)),
                ),
                ft.OutlinedButton(
                    "汇总校验结果说明", icon=ft.Icons.SUMMARIZE_OUTLINED,
                    tooltip="操作步骤：\n1. 选择源数据目录；\n2. 选择对应的汇总模板；\n3. 勾选待处理文件；\n4. 点击后确认执行。\n提示：说明文件分散在子目录时，请在设置中心勾选“递归子文件夹”。",
                    on_click=lambda event: start("summary"),
                ),
                ft.OutlinedButton(
                    "组合联合核查表", icon=ft.Icons.ACCOUNT_TREE_OUTLINED,
                    tooltip="操作步骤：\n1. 选择源数据目录；\n2. 勾选要组合的报送文件；\n3. 点击后确认执行。\n提示：机构文件位于子目录时，请在设置中心勾选“递归子文件夹”；本功能不需要模板或外部辅助文件。",
                    on_click=lambda event: start("merge"),
                ),
                ft.OutlinedButton(
                    "制作联合模板", icon=ft.Icons.AUTO_AWESOME,
                    tooltip="操作步骤：\n1. 点击按钮选择基准模板；\n2. 选择一个或多个待并入模板；\n3. 确认生成联合模板。\n提示：基准模板不会被修改，结果写入其同级“联合模板”目录。",
                    on_click=make_union_template,
                ),
            ], spacing=8, wrap=False),
            ft.Row([
                custom_actions,
                ft.Container(expand=True),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        ], spacing=10, horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

        run_log_section = ft.Column([
            ft.Row([
                ft.Text("运行记录", size=13, weight=ft.FontWeight.W_700, color=TEXT), busy,
                ft.TextButton("导出运行记录", icon=ft.Icons.DOWNLOAD_OUTLINED, on_click=export_run_log),
            ], spacing=8),
            ft.Container(
                content=log_view,
                bgcolor=CONSOLE_BG,
                border_radius=8,
                padding=10,
                expand=True,
            ),
        ], spacing=8, horizontal_alignment=ft.CrossAxisAlignment.STRETCH, expand=True)

        # 下半区与上排保持左右对称：左侧待处理文件与“源数据目录”同为
        # 半屏宽度，右侧为功能入口和运行日志。中间使用一条明确的分隔线，
        # 避免两组卡片视觉上连成一片。
        action_card = card(execution_section, height=158)
        log_card = card(run_log_section, expand=True)
        bottom_section = ft.Row([
            files_card,
            ft.Container(width=1, bgcolor="#D7E0EC", margin=ft.Margin(7, 0, 7, 0)),
            ft.Column([action_card, log_card], spacing=16, expand=True,
                      horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
        ], spacing=0, expand=True, vertical_alignment=ft.CrossAxisAlignment.STRETCH)
        files_card.expand = True
        workbench_view = ft.Column([
            top_row,
            bottom_section,
        ], spacing=12, expand=True,
           horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

        # ================= 设置中心页面 =================
        def guide_block(title: str, lines: list[str]) -> "ft.Control":
            # 指引页放在 ListView 内滚动。不要把文字放进没有确定可用宽度的
            # 嵌套 Row/Column，部分 Flet 桌面后端会因此把正文压缩成不可见区域。
            return ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Container(width=4, height=18, border_radius=2, bgcolor=PRIMARY),
                        ft.Text(title, weight=ft.FontWeight.W_700, size=14, color=TEXT),
                    ], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    *[ft.Row([
                        ft.Container(width=5, height=5, border_radius=3, bgcolor=INFO, margin=ft.Margin(0, 7, 0, 0)),
                        ft.Text(line, size=13, color=TEXT, expand=True, selectable=True),
                    ], spacing=8, vertical_alignment=ft.CrossAxisAlignment.START) for line in lines],
                ], spacing=7, horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
                padding=ft.Padding(4, 4, 12, 14),
            )

        general_content = ft.Column([
            ft.Text("计算引擎设置", size=14, weight=ft.FontWeight.W_700, color=TEXT),
            ft.Row([
                ft.Text("公式计算引擎", size=13, color=TEXT, width=160),
                settings_engine,
            ], spacing=12),
            ft.Text("统信版固定使用 LibreOffice Calc；程序启动时自动检测安装入口与版本。", size=12, color=TEXT_2),
            ft.Divider(height=1, color=BORDER),
            ft.Text("历史审核说明", size=14, weight=ft.FontWeight.W_700, color=TEXT),
            ft.Row([
                settings_history_path,
                ft.OutlinedButton("绑定文件", icon=ft.Icons.FOLDER_OPEN, on_click=choose_history),
            ], spacing=10),
            ft.Text("用于读取和更新历史审核结果；可绑定到程序目录以外的历史审核说明.xlsx。", size=12, color=TEXT_2),
            ft.Divider(height=1, color=BORDER),
            ft.Text("运行选项", size=14, weight=ft.FontWeight.W_700, color=TEXT),
            ft.Row([settings_recursive, ft.Text("递归子文件夹：汇总核查表校验通常不勾选；汇总校验结果说明在机构文件分散于子文件夹时通常勾选。", size=12, color=TEXT_2, expand=True)], spacing=8),
            ft.Row([settings_logs, ft.Text("每次流程运行生成一份运行日志工作簿。", size=12, color=TEXT_2, expand=True)], spacing=8),
            ft.Row([settings_confirm, ft.Text("每次执行前确认待处理对象、模板、外部文件和输出目录；关闭后直接执行。", size=12, color=TEXT_2, expand=True)], spacing=8),
            ft.Row([settings_custom, ft.Text("在主工作台显示“自定义流程”下拉入口。", size=12, color=TEXT_2, expand=True)], spacing=8),
            ft.Divider(height=1, color=BORDER),
            ft.Text("环境提示", size=14, weight=ft.FontWeight.W_700, color=TEXT),
            ft.Container(
                content=ft.Column([
                    ft.Row([ft.Icon(ft.Icons.INFO_OUTLINE, color=WARN, size=18), ft.Text("本版仅使用 LibreOffice Calc。", size=13, color=TEXT, expand=True)], spacing=8),
                    engine_detail,
                    ft.Text("未检测到时，请安装 LibreOffice Calc，并确认终端执行 soffice --version 可输出版本；非 PATH 安装可设置 BASE_AUDIT_SOFFICE。", size=12, color=TEXT_2),
                ], spacing=6),
                bgcolor="#FFF8E6", border_radius=8, padding=12,
            ),
        ], spacing=12, scroll=ft.ScrollMode.AUTO, expand=True)

        guide_content = ft.ListView([
            guide_block("汇总核查表校验", [
                "准备：选择源数据目录（机构报送文件所在目录；一般不勾选“递归子文件夹”）。",
                "准备：在“模板文件目录”选择模板目录，程序自动识别匹配模板；必要时手动更换“匹配模板”或添加“外部辅助文件”。",
                "步骤：核对“待审核文件”清单及匹配情况（已匹配 / 未匹配），点“汇总核查表校验”。",
                "输出：每份文件生成 _审核版 副本与本期审核结果工作簿；运行记录列出命中的校验问题（工作表、定位单元格、级别、校验字段、当前值、对比值、差值、规则编号）。",
                "提示：复杂公式（INDIRECT、OFFSET、动态数组、外部引用）必须使用 LibreOffice Calc，并以模板基准验证结果为准。",
            ]),
            guide_block("汇总校验结果说明", [
                "准备：选择校验结果及报送说明目录，以及对应的汇总模板。",
                "准备：机构文件分散在子文件夹时，在“常规设置”勾选“递归子文件夹”。",
                "步骤：确认清单后点“汇总校验结果说明”。",
                "输出：按模块生成汇总输出工作表；多个区域匹配到同一模块且表头相同才合并为一张表，表头不同分别按原表名输出。",
            ]),
            guide_block("组合联合核查表", [
                "准备：只需源数据目录（可勾选“递归子文件夹”）；不需要模板、外部文件或数据期。",
                "步骤：点“组合联合核查表”执行组合；制作联合模板请使用主界面的独立按钮。",
                "输出：按源文件主名下划线第一段分组，将同一机构的多个工作簿合并为一本；同名工作表保留先复制的一张、后续跳过。",
            ]),
            guide_block("制作联合模板", [
                "说明：这是低频的模板制作动作，不使用页面路径或自动识别模板。",
                "第 1 步：选择基准模板（包含集中系统数据、参照表等公共依赖工作表；原文件不会修改）。",
                "第 2 步：批量选择并入模板（其独有工作表将复制到联合模板；不要重复选择基准模板）。",
                "输出：在基准模板同级“联合模板”目录生成联合模板文件与检查报告；同名工作表保留基准/先复制版本。",
                "注意：所有外部工作簿公式引用会统一改为工作簿内引用；若对应工作表未合并，最终以 #REF! 写入报告提示处理。",
            ]),
        ], spacing=8, padding=ft.Padding(4, 4, 4, 4), expand=True)

        catalog_content = ft.Row([
            ft.Container(catalog_list, width=280, bgcolor="#F8FAFC", border_radius=8, padding=8, border=thin_border()),
            ft.Container(width=1, bgcolor=BORDER),
            catalog_detail,
        ], expand=True, spacing=16)

        modules_content = ft.Row([
            ft.Container(module_list, width=280, bgcolor="#F8FAFC", border_radius=8, padding=8, border=thin_border()),
            ft.Container(width=1, bgcolor=BORDER),
            ft.Column([
                module_name,
                module_type,
                module_range,
                module_input,
                module_output,
                module_note,
                module_error,
                ft.Row([
                    ft.OutlinedButton("新建", icon=ft.Icons.ADD, on_click=new_module),
                    ft.FilledButton("保存本项", on_click=save_module),
                    ft.OutlinedButton("复制", icon=ft.Icons.CONTENT_COPY, on_click=copy_module),
                    ft.OutlinedButton("删除", icon=ft.Icons.DELETE_OUTLINE, on_click=delete_module),
                ], spacing=8),
                plans_section,
            ], spacing=12, scroll=ft.ScrollMode.AUTO, expand=True,
               horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
        ], expand=True, spacing=16)

        flow_current_title = ft.Text("当前流程：未选择", size=15, weight=ft.FontWeight.W_700, color=TEXT)
        flow_count = ft.Text("共 0 步", size=12, color=TEXT_2)
        flow_display_switch = ft.Switch(label="显示在主界面", scale=0.75)

        def flow_display_allowed(flow_name: str) -> bool:
            return bool(flow_name) and (
                flow_name == COMBINE_SHEETS_FLOW
                or flow_name not in set(sc["editor"].get("defaultFlowNames", []))
            )

        def sync_flow_display_switch() -> None:
            flow_name = str(sc["sel_flow"] or "")
            allowed = flow_display_allowed(flow_name)
            flow_display_switch.disabled = not allowed
            flow_display_switch.value = (
                bool(sc["display"].get(flow_name, flow_name == COMBINE_SHEETS_FLOW))
                if allowed else bool(flow_name)
            )

        def change_flow_display(event) -> None:
            flow_name = str(sc["sel_flow"] or "")
            if flow_display_allowed(flow_name):
                sc["display"][flow_name] = bool(event.control.value)
            else:
                sync_flow_display_switch()
            page.update()

        flow_display_switch.on_change = change_flow_display

        def flow_step_card(index: int, row: dict[str, str]) -> "ft.Control":
            def toggle_step(event) -> None:
                row["启用"] = "是" if event.control.value else "否"
                sync_steps_to_flows(); render_steps(); page.update()
            return ft.Container(
                content=ft.Column([
                    ft.Row([ft.Icon(ft.Icons.DRAG_INDICATOR, size=16, color=TEXT_3), ft.Text(str(row.get("顺序", "")), width=32, size=13, weight=ft.FontWeight.W_700, color=PRIMARY), ft.Text(str(row.get("功能名", "")), expand=True, size=13, weight=ft.FontWeight.W_500, color=TEXT), ft.Text("失败：{}".format(row.get("失败后处理", "跳过")), size=11, color=TEXT_2), ft.Switch(value=str(row.get("启用", "是")) == "是", on_change=toggle_step, scale=0.75)], spacing=6),
                    ft.Row([ft.Text("输入：{}".format(row.get("输入", "源数据目录") or "源数据目录"), size=11, color=TEXT_2, expand=True), ft.Icon(ft.Icons.ARROW_FORWARD, size=14, color=TEXT_3), ft.Text("输出：{}".format(row.get("输出", "运行日志") or "运行日志"), size=11, color=TEXT_2, expand=True)], spacing=5),
                ], spacing=3), bgcolor="#F5F7FA", border_radius=8, padding=8, on_click=lambda event, i=index: select_step(i),
            )

        # 流程级操作（新增流程 / 添加步骤 / 删除流程）统一收在模板库卡片底部，
        # 用小号按钮排成一行；点“添加步骤”会弹出步骤编辑弹窗（选功能 + 填输入输出）。
        flow_action_style = ft.ButtonStyle(
            padding=ft.Padding(8, 2, 8, 2),
            text_style=ft.TextStyle(size=12),
        )
        flow_library_panel = ft.Container(ft.Column([
            ft.Text("流程模板库", size=14, weight=ft.FontWeight.W_700, color=TEXT),
            flow_list,
            ft.Divider(height=1),
            ft.Row([
                ft.OutlinedButton("新增流程", on_click=new_flow, style=flow_action_style),
                ft.OutlinedButton("添加步骤", on_click=lambda event: open_step_editor("add"), style=flow_action_style),
                ft.OutlinedButton("删除流程", on_click=delete_flow, style=flow_action_style),
            ], spacing=4, wrap=True),
        ], spacing=6, expand=True, horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
        width=280, height=480, bgcolor="#F5F7FA", padding=12, border_radius=8)
        flow_main_panel = ft.Container(ft.Column([
            ft.Row([flow_current_title, flow_count, ft.Container(expand=True), flow_display_switch], spacing=8),
            ft.Text("步骤列表", size=13, weight=ft.FontWeight.W_600, color=TEXT_2), steps_view,
        ], spacing=8, expand=True), width=500, height=480, padding=12, bgcolor="#FFFFFF")
        flows_content = ft.Row([flow_library_panel, flow_main_panel], spacing=16, height=500)

        settings_error = ft.Text(size=12, color=DANGER)

        # 设置中心左侧竖向目录：常规设置 / 操作指引 / 模块说明 / 功能配置 / 流程编排
        settings_meta = {
            "general": ("常规设置", ft.Icons.TUNE),
            "guide": ("操作指引", ft.Icons.HELP_OUTLINE),
            "catalog": ("模块说明", ft.Icons.MENU_BOOK_OUTLINED),
            "modules": ("功能配置", ft.Icons.VIEW_QUILT_OUTLINED),
            "flows": ("流程编排", ft.Icons.ACCOUNT_TREE_OUTLINED),
        }
        settings_contents = {
            "general": general_content,
            "guide": guide_content,
            "catalog": catalog_content,
            "modules": modules_content,
            "flows": flows_content,
        }
        settings_nav_buttons: dict[str, "ft.Container"] = {}
        settings_nav_icons: dict[str, "ft.Icon"] = {}
        settings_nav_labels: dict[str, "ft.Text"] = {}
        settings_content_area = ft.Container(
            expand=True, content=general_content, padding=ft.Padding(20, 12, 20, 12),
        )

        def switch_settings(name: str) -> None:
            for key in settings_nav_buttons:
                selected = key == name
                settings_nav_buttons[key].bgcolor = "#EAF2FF" if selected else None
                settings_nav_icons[key].color = PRIMARY if selected else TEXT_2
                settings_nav_labels[key].color = PRIMARY if selected else TEXT
                settings_nav_labels[key].weight = ft.FontWeight.W_700 if selected else ft.FontWeight.W_400
            settings_content_area.content = settings_contents[name]
            page.update()

        def settings_nav_item(key: str) -> "ft.Control":
            label, icon = settings_meta[key]
            icon_ctl = ft.Icon(icon, color=TEXT_2, size=16)
            label_ctl = ft.Text(label, color=TEXT, size=13)
            settings_nav_icons[key] = icon_ctl
            settings_nav_labels[key] = label_ctl
            box = ft.Container(
                content=ft.Row([icon_ctl, label_ctl], spacing=8),
                padding=ft.Padding(12, 9, 12, 9),
                border_radius=6,
                margin=ft.Margin(4, 1, 4, 1),
                on_click=lambda e, k=key: switch_settings(k),
                ink=True,
            )
            settings_nav_buttons[key] = box
            return box

        # ================= 设置中心数据加载与保存 =================
        def render_settings() -> None:
            render_catalog()
            render_module_list()
            render_flow_list()
            render_steps()
            render_plans()
            flow_current_title.value = "当前流程：{}".format(sc["sel_flow"] or "未选择")
            flow_count.value = "共 {} 步".format(len(sc["step_rows"]))
            steps_view.controls = [flow_step_card(index, row) for index, row in enumerate(sc["step_rows"])]
            page.update()

        def load_settings() -> None:
            sc["editor"] = api.get_config_editor_data()
            sc["modules"] = [dict(r) for r in sc["editor"].get("customModules", [])]
            sc["flows"] = [dict(r) for r in sc["editor"].get("customFlows", [])]
            sc["empty_flows"] = []
            sc["display"] = dict(sc["editor"].get("flowDisplay", {}))
            sc["plans"] = [dict(r) for r in sc["editor"].get("combineSheetsPlans", [])]
            sc["active_plan"] = str(sc["editor"].get("activeCombineSheetsPlanId", "") or "")
            settings_engine.value = str(api.state["calculationEngine"])
            settings_history_path.value = str(api.history_path)
            settings_history_path.tooltip = str(api.history_path)
            settings_logs.value = bool(api.state["writeFlowLogs"])
            settings_confirm.value = bool(api.state.get("confirmBeforeRun", True))
            settings_custom.value = bool(api.state["showCustomFeatures"])
            settings_recursive.value = bool(api.state["recursive"])
            update_engine_status()
            feature_types = [str(t) for t in sc["editor"].get("featureTypes", [])]
            module_type.options = [ft.DropdownOption(key=t, text=t) for t in feature_types]
            module_names = [str(r.get("功能名", "")) for r in module_rows() if r.get("功能名")]
            step_feature.options = [ft.DropdownOption(key=n, text=n) for n in module_names]
            mrows = module_rows()
            if mrows:
                select_module(mrows[0])
            frows = flow_rows()
            sc["sel_flow"] = str(frows[0].get("流程名", "")) if frows else ""
            sc["step_rows"] = []
            sc["sel_step"] = 0
            if sc["sel_flow"]:
                select_flow(sc["sel_flow"])
            else:
                render_flow_list()
                render_steps()
            sc["sel_plan"] = str(sc["plans"][0].get("id", "")) if sc["plans"] else ""
            if sc["sel_plan"]:
                select_plan(sc["plans"][0])
            else:
                render_plans()
            settings_error.value = ""
            render_settings()
            page.update()

        def save_settings(event=None) -> None:
            try:
                sync_steps_to_flows()
                api.update({
                    "recursive": settings_recursive.value,
                    "calculationEngine": settings_engine.value,
                    "writeFlowLogs": settings_logs.value,
                    "confirmBeforeRun": settings_confirm.value,
                    "showCustomFeatures": settings_custom.value,
                })
                draft = {
                    "customModules": sc["modules"],
                    "customFlows": sc["flows"],
                    "flowDisplay": sc["display"],
                    "combineSheetsPlans": sc["plans"],
                    "activeCombineSheetsPlanId": str(sc["active_plan"] or ""),
                }
                check = api.validate_config_editor_draft(draft)
                if not check["valid"]:
                    raise ValueError("；".join(check["errors"]))
                result = api.save_config_editor_draft(draft)
                if not result["ok"]:
                    raise ValueError("；".join(result["errors"]))
                page.snack_bar = ft.SnackBar(ft.Text("设置已保存"), open=True)
                settings_dialog.open = False
            except Exception as exc:
                settings_error.value = "保存失败：" + str(exc)
            page.update()

        def reset_settings(event=None) -> None:
            sc["modules"] = []
            sc["flows"] = []
            sc["display"] = {}
            sc["empty_flows"] = []
            sc["plans"] = [dict(r) for r in sc["editor"].get("combineSheetsPlans", [])]
            sc["active_plan"] = str(sc["editor"].get("activeCombineSheetsPlanId", "") or "")
            settings_error.value = ""
            mrows = module_rows()
            sc["sel_module"] = str(mrows[0].get("功能名", "")) if mrows else ""
            if mrows:
                select_module(mrows[0])
            frows = flow_rows()
            sc["sel_flow"] = str(frows[0].get("流程名", "")) if frows else ""
            sc["step_rows"] = []
            sc["sel_step"] = 0
            if sc["sel_flow"]:
                select_flow(sc["sel_flow"])
            sc["sel_plan"] = str(sc["plans"][0].get("id", "")) if sc["plans"] else ""
            if sc["sel_plan"]:
                select_plan(sc["plans"][0])
            render_settings()
            page.update()

        # ================= 顶部条 + 设置中心弹窗 =================
        def open_settings(event=None) -> None:
            load_settings()
            switch_settings("general")
            settings_dialog.open = True
            page.update()

        settings_dialog = ft.AlertDialog(
            modal=True,
            # AlertDialog 默认的 Material surface 是灰色，内容 Container 又是
            # 白色时会形成“灰底套白卡”的两层视觉。统一由弹窗本身提供白底。
            bgcolor="#FFFFFF",
            elevation=0,
            title_padding=ft.Padding(20, 18, 20, 12),
            content_padding=0,
            actions_padding=ft.Padding(20, 10, 20, 16),
            title=ft.Text("设置中心", size=16, weight=ft.FontWeight.W_700, color=TEXT),
            content=ft.Container(
                ft.Column([
                    settings_error,
                    ft.Row([
                        # 左侧竖向目录（嵌入大弹窗左栏，白底一体，右侧分隔线）
                        ft.Container(
                            content=ft.Column([
                                *[settings_nav_item(k) for k in settings_meta],
                            ], spacing=2),
                            width=180,
                            padding=0,
                            border=ft.Border(right=ft.BorderSide(1, BORDER)),
                        ),
                        # 右侧内容区（撑满剩余宽度）
                        settings_content_area,
                    ], expand=True, spacing=0),
                ], spacing=8, expand=True),
                bgcolor="#FFFFFF",
                width=1000, height=620,
            ),
            actions=[
                ft.TextButton("恢复默认", on_click=lambda event: reset_settings(event)),
                ft.TextButton("取消", on_click=lambda event: setattr(settings_dialog, "open", False) or page.update()),
                ft.FilledButton("保存", on_click=lambda event: save_settings(event)),
            ],
        )
        page.overlay.append(settings_dialog)

        main_content = ft.Container(
            content=workbench_view,
            expand=True,
            bgcolor=BG,
            padding=24,
        )

        # ================= 渲染循环 =================
        def update_engine_bar() -> None:
            engines = api.state.get("engineStatus", {})
            libre_info = engines.get("libreoffice", {}) if isinstance(engines, dict) else {}

            def seg(ok: bool, label: str, detail: str) -> "ft.Control":
                color = SUCCESS if ok else DANGER
                return ft.Row([
                    ft.Icon(ft.Icons.CHECK_CIRCLE if ok else ft.Icons.ERROR_OUTLINE, color=color, size=15),
                    ft.Text("{} {}".format(label, "✓" if ok else "✗"), size=12, color=TEXT),
                    ft.Text(detail, size=12, color=TEXT_2),
                ], spacing=4)

            engine_bar.controls = [
                seg(bool(libre_info.get("available")), "LibreOffice", str(libre_info.get("detail", ""))),
            ]

        def render() -> None:
            nonlocal rendered_log_text, custom_flow_menu, rendered_custom_flows
            state = api.get_state()
            for name, control in paths.items():
                value = str(state.get(name) or "")
                control.value = display_name(value)
                control.color = TEXT if value else TEXT_2
                control.tooltip = value if value else None
            engine_note.value = engine_status_text()
            output_pin.icon = ft.Icons.PUSH_PIN if state.get("outputPinned") else ft.Icons.PUSH_PIN_OUTLINED
            output_pin.tooltip = "输出目录已固定" if state.get("outputPinned") else "跟随源数据目录"
            status.value = str(state.get("status") or "就绪")
            busy.visible = bool(state.get("busy"))
            update_engine_bar()
            selected = set(state.get("selectedFiles", []))
            custom_flows = api.get_custom_features()
            custom_flow_signature = tuple((str(item["flow"]), str(item["name"])) for item in custom_flows)
            if custom_flow_signature != rendered_custom_flows:
                custom_flow_menu = ft.Dropdown(
                    hint_text="选择自定义流程", width=250, height=40, dense=True,
                    options=[ft.DropdownOption(key=flow, text=name) for flow, name in custom_flow_signature],
                )
                custom_flow_menu.text_size = 12
                custom_actions.controls = ([ft.Row([
                    ft.Container(custom_flow_menu, width=250, height=44, alignment=ft.Alignment(-1, 0)),
                    ft.OutlinedButton("执行", icon=ft.Icons.PLAY_ARROW,
                                      on_click=execute_selected_custom_flow),
                ], spacing=8)] if custom_flow_menu.options else [])
                rendered_custom_flows = custom_flow_signature
            files = state.get("sourceFiles", [])
            if files:
                file_rows.controls = _file_grid(files, selected)
            else:
                file_rows.controls = []
            matched_count = sum(
                1 for item in files
                if str(item.get("template") or "").strip() not in {"", "未识别"}
            )
            file_summary.value = "共 {} 张，匹配模板共 {} 张".format(len(files), matched_count)
            log_lines = [str(item.get("text", item)) for item in state.get("log", [])]
            current_log_text = "\n".join(log_lines)
            if current_log_text != rendered_log_text:
                log_view.controls = [ft.Text(
                    current_log_text, color=CONSOLE_TEXT, size=12,
                    selectable=True, font_family=CHINESE_UI_FONT)]
                rendered_log_text = current_log_text
            page.update()

        async def monitor() -> None:
            while True:
                render()
                await asyncio.sleep(0.8)

        page.add(
            main_content,
        )
        load_settings()
        render()
        page.run_task(monitor)

    ft.app(target=main, view=ft.AppView.FLET_APP)
