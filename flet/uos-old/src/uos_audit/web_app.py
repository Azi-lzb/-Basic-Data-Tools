"""Stateful controller shared by the Flet desktop workbench and CLI.

Despite its historical filename, this module has no pywebview, GTK or WebKit
dependency.  It owns only workflow state and background task scheduling.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from .calculation import ENGINE_DEFAULT, ENGINE_OPTIONS, engine_available, engine_status, inspect_formula_capability, preferred_engine
from .discovery import detect_period, explanation_files, recommend_template, source_workbooks
from .merge_workbooks import create_combined_template
from .name_config import (
    COMBINE_SHEETS_FLOW,
    COMBINE_SHEETS_FUNCTION,
    EXTERNAL_FILE_FUNCTION,
    FORMULA_COPY_FUNCTION,
    HISTORY_WORKBOOK_NAME,
    MERGE_ORG_FILES_FUNCTION,
    NAMED_RANGE_CHECK_FUNCTION,
    features_of_type,
    initialize_config,
    load_feature_mappings,
    load_flow_steps,
    load_config_editor_data,
    reset_default_configuration,
    save_config_editor_draft,
    validate_config_editor_draft,
)
from .openpyxl_workbook import read_template
from .workflow import run_flow


class WebApi:
    """API contract used by the shared local HTML workbench."""

    def __init__(self, root: Path) -> None:
        self.root = root
        self.settings_path = root / "data" / "用户设置.json"
        preferences = self._load_preferences()
        default_history = (root / HISTORY_WORKBOOK_NAME).resolve()
        saved_history = Path(str(preferences.get("historyPath") or "")).expanduser()
        self.history_path = saved_history.resolve() if saved_history.is_file() else default_history
        initialize_config(self.history_path)
        detected_engines = engine_status()
        self.state: dict[str, Any] = {
            "busy": False,
            "status": "就绪（统信原生审核引擎）",
            "log": [],
            "input": str(preferences.get("input") or ""), "templateDir": str(preferences.get("templateDir") or ""),
            "template": str(preferences.get("template") or ""), "templateManual": bool(preferences.get("templateManual", False)),
            "external": str(preferences.get("external") or ""), "output": str(preferences.get("output") or ""),
            "outputAuto": bool(preferences.get("outputAuto", True)), "outputPinned": bool(preferences.get("outputPinned", False)),
            "recursive": bool(preferences.get("recursive", False)),
            "writeFlowLogs": bool(preferences.get("writeFlowLogs", False)),
            "confirmBeforeRun": bool(preferences.get("confirmBeforeRun", True)),
            # The default is detected on every launch. A saved choice must not
            # force an unavailable engine or bypass the agreed priority.
            "calculationEngine": preferred_engine(detected_engines),
            "engineStatus": detected_engines, "showCustomFeatures": bool(preferences.get("showCustomFeatures", True)),
            "historyPath": str(self.history_path),
            "sourceFiles": [], "selectedFiles": [], "mixedTemplates": [],
            "explanationFiles": [], "detectedPeriod": "", "lastOutput": "",
            "formulaCapability": {},
        }
        if self.state["calculationEngine"] not in ENGINE_OPTIONS:
            self.state["calculationEngine"] = ENGINE_DEFAULT
        if self.state["input"]:
            self._recognize(allow_template_auto=True)

    def get_state(self) -> dict[str, Any]:
        return self.state

    def update(self, values: dict[str, Any]) -> dict[str, Any]:
        if not isinstance(values, dict):
            return self.state
        for field in ("input", "templateDir", "template", "external", "output"):
            if field in values:
                self._set_path(field, str(values[field] or "").strip())
        if "historyPath" in values:
            self._set_history_path(str(values["historyPath"] or "").strip())
        for field in ("recursive", "writeFlowLogs", "confirmBeforeRun", "showCustomFeatures", "outputPinned"):
            if field in values:
                self.state[field] = self._as_bool(values[field])
        if "outputPinned" in values:
            self.state["outputAuto"] = not self.state["outputPinned"]
        if "recursive" in values:
            self._recognize(allow_template_auto=False)
        if "calculationEngine" in values:
            selected = str(values["calculationEngine"] or ENGINE_DEFAULT).strip()
            if selected not in ENGINE_OPTIONS:
                raise ValueError("不支持的计算引擎：{}".format(selected))
            latest_status = engine_status()
            if not engine_available(selected, latest_status):
                raise ValueError("计算引擎“{}”当前不可用，请选择已检测到的引擎".format(selected))
            self.state["engineStatus"] = latest_status
            self.state["calculationEngine"] = selected
        else:
            self.state["engineStatus"] = engine_status()
        if any(field in values for field in (
            "input", "templateDir", "template", "external", "output", "outputPinned", "historyPath",
            "recursive", "writeFlowLogs", "confirmBeforeRun", "showCustomFeatures", "calculationEngine",
        )):
            self._save_preferences()
        return self.state

    def recognize(self) -> dict[str, Any]:
        self._recognize(force_template=True)
        return self.state

    def refresh_sources(self) -> dict[str, Any]:
        self._recognize(allow_template_auto=False)
        return self.state

    def recommend_template(self) -> dict[str, Any]:
        return self.recognize()

    def inspect_current_template(self) -> dict[str, Any]:
        """Read-only preflight for the configured Python formula engine."""
        path = Path(str(self.state.get("template") or ""))
        if not path.is_file():
            self._log("请先选择有效的模板文件，再检查公式兼容性")
            return self.state
        try:
            mappings = load_feature_mappings(self.history_path, path)
            definition = read_template(
                path, formula_mappings=features_of_type(mappings, FORMULA_COPY_FUNCTION)
            )
            capability = inspect_formula_capability(path, formula_ranges=definition.copy_ranges)
            payload = {
                "formulaCount": capability.formula_count,
                "functions": list(capability.functions),
                "unsupportedFunctions": list(capability.unsupported_functions),
                "blockers": list(capability.blockers),
                "formulaErrors": [
                    {"sheet": sheet, "cell": cell, "error": error}
                    for sheet, cell, error in capability.formula_errors
                ],
                "pythonSupported": capability.python_supported,
            }
            self.state["formulaCapability"] = payload
            self._log("模板公式检查：{}".format(capability.summary()))
            for sheet, cell, error in capability.formula_errors:
                self._log("错误公式定位：{}!{}（{}）".format(sheet, cell, error), detail=True)
        except Exception as exc:
            self._log("模板公式检查失败：{}".format(exc))
        return self.state

    def set_selected_files(self, paths: list[str]) -> dict[str, Any]:
        allowed = {str(item["path"]) for item in self.state["sourceFiles"]}
        self.state["selectedFiles"] = [str(path) for path in (paths or []) if str(path) in allowed]
        return self.state

    def start_with_state(self, action: str, values: dict[str, Any] | None = None, selected_files: list[str] | None = None) -> bool:
        if values:
            self.update(values)
        if selected_files is not None:
            self.set_selected_files(selected_files)
        return self.start(action)

    def start_flow_with_state(self, flow_name: str, values: dict[str, Any] | None = None, selected_files: list[str] | None = None) -> bool:
        name = str(flow_name or "").strip()
        if not name:
            self._log("执行流程名称不能为空")
            return False
        return self.start_with_state("flow:" + name, values, selected_files)

    def start(self, action: str) -> bool:
        if self.state["busy"]:
            return False
        self.state["busy"] = True
        label = {"audit": "汇总核查表校验", "summary": "汇总校验结果说明"}.get(action, action[5:] if action.startswith("flow:") else action)
        self.state["status"] = "正在处理，请勿关闭窗口……"
        self._log("开始执行：{}".format(label))
        threading.Thread(target=self._worker, args=(action,), daemon=True).start()
        return True

    def get_config_editor_data(self) -> dict[str, object]:
        return load_config_editor_data(self.history_path)

    def validate_config_editor_draft(self, draft: object) -> dict[str, object]:
        errors = validate_config_editor_draft(self.history_path, draft)
        return {"valid": not errors, "errors": errors}

    def save_config_editor_draft(self, draft: object) -> dict[str, object]:
        if self.state["busy"]:
            return {"ok": False, "errors": ["任务正在运行，暂不能修改配置"]}
        try:
            save_config_editor_draft(self.history_path, draft)
        except Exception as exc:
            return {"ok": False, "errors": [str(exc)]}
        self._log("已保存功能、流程和主界面显示设置")
        return {"ok": True, "errors": [], "data": self.get_config_editor_data()}

    def reset_config(self) -> dict[str, object]:
        reset_default_configuration(self.history_path)
        self._log("已恢复默认模块和流程；历史审核说明未修改")
        return self.get_config_editor_data()

    def get_custom_features(self) -> list[dict[str, str]]:
        # “组合工作表”从主按钮移入自定义流程下拉框，并始终作为可用的
        # 默认项，避免首次使用时该下拉框为空。
        if not self.state["showCustomFeatures"]:
            return []
        editor = self.get_config_editor_data()
        shown = editor.get("flowDisplay", {})
        names: list[str] = []
        for row in editor.get("customFlows", []):
            name = str(row.get("流程名") or "").strip()
            if name and shown.get(name) and name not in names:
                names.append(name)
        if shown.get(COMBINE_SHEETS_FLOW, True) and COMBINE_SHEETS_FLOW not in names:
            names.append(COMBINE_SHEETS_FLOW)
        return [{"id": name, "name": name, "flow": name, "remark": ""} for name in names]

    def run_custom(self, flow_name: str) -> bool:
        return self.start("flow:" + str(flow_name or "").strip())

    def get_run_context(self, action: str) -> dict[str, Any]:
        """Describe the selected flow inputs for the pre-run confirmation UI."""
        flow = action[5:] if action.startswith("flow:") else {
            "audit": "汇总核查表校验",
            "summary": "汇总校验结果说明",
            "merge": "组合联合核查表",
            "combine": "组合工作表",
        }.get(action, action)
        steps = load_flow_steps(self.history_path, flow)
        mappings = {
            item.name: item
            for item in load_feature_mappings(self.history_path, Path("__无需模板__.xlsx"))
        }
        types = {
            mappings[step.feature_name].feature_type
            for step in steps
            if step.feature_name in mappings
        }
        selected = list(self.state.get("selectedFiles") or [])
        return {
            "action": action,
            "flow": flow,
            "selectedFiles": selected,
            "templateRequired": self._flow_requires_template(flow),
            "externalRequired": EXTERNAL_FILE_FUNCTION in types,
            "template": str(self.state.get("template") or ""),
            "external": str(self.state.get("external") or ""),
            "output": str(self.state.get("output") or ""),
        }

    def create_template_merge(self) -> bool:
        """Retained for callers; Flet should use explicit file paths instead."""
        self._log("请在 Flet 工作台中选择基准模板和待并入模板")
        return False

    def create_template_merge_from_paths(self, base: str | Path, joins: list[str | Path]) -> bool:
        """Create a union template from paths selected by the Flet file picker."""
        if self.state["busy"]:
            return False
        base_path = Path(base).resolve()
        selected = [Path(item).resolve() for item in joins if Path(item).resolve() != base_path]
        if not base_path.is_file() or not selected:
            self._log("制作联合模板需要一个基准模板和至少一个待并入模板")
            return False
        self.state["busy"] = True
        self._log("开始制作联合模板：基准 {}；并入 {} 个模板".format(base_path.name, len(selected)))
        threading.Thread(target=self._template_merge_worker, args=(base_path, selected), daemon=True).start()
        return True

    def open_path(self, path: str) -> dict[str, Any]:
        target = Path(path)
        if not target.exists():
            self._log("未找到路径：{}".format(target))
            return self.state
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                subprocess.Popen(["xdg-open", str(target)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except OSError as exc:
            self._log("打开路径失败：{}".format(exc))
        return self.state

    def open_last_output(self) -> dict[str, Any]:
        """Open the actual output produced by the most recent completed flow."""
        value = str(self.state.get("lastOutput") or "")
        if not value:
            self._log("本次尚未生成可打开的结果文件")
            return self.state
        return self.open_path(value)

    def open_user_guide(self) -> dict[str, Any]:
        guide = self.root / "基础数据审核工具使用说明.docx"
        if not guide.is_file():
            self._log("统信版尚未附带使用说明：{}".format(guide))
            return self.state
        return self.open_path(str(guide))

    def _set_path(self, field: str, value: str) -> None:
        previous = self.state.get(field, "")
        self.state[field] = value
        if field == "template" and value and value != previous:
            self.state["templateManual"] = True
        if field == "templateDir" and value != previous:
            self.state["templateManual"] = False
        if field == "input" and value and not self.state["outputPinned"]:
            self.state["output"] = str(Path(value) / "执行结果")
            self.state["outputAuto"] = True
        if field == "output" and value != previous:
            self.state["outputPinned"] = True
            self.state["outputAuto"] = False

    def _set_history_path(self, value: str) -> None:
        if not value:
            return
        path = Path(value).expanduser().resolve()
        if path.suffix.casefold() != ".xlsx" or not path.is_file():
            raise ValueError("历史审核说明必须是已存在的 .xlsx 文件")
        if path == self.history_path:
            return
        initialize_config(path)
        self.history_path = path
        self.state["historyPath"] = str(path)
        self._log("已绑定历史审核说明：{}".format(path))

    def _load_preferences(self) -> dict[str, object]:
        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError, TypeError):
            return {}

    def _save_preferences(self) -> None:
        self.settings_path.parent.mkdir(parents=True, exist_ok=True)
        data = {
            key: self.state[key]
            for key in (
                "input", "templateDir", "template", "templateManual", "external", "output", "outputAuto", "outputPinned",
                "recursive", "writeFlowLogs", "confirmBeforeRun", "showCustomFeatures", "calculationEngine",
            )
        }
        data["historyPath"] = str(self.history_path)
        temporary = self.settings_path.with_suffix(".json.tmp")
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(self.settings_path)

    def _recognize(self, *, force_template: bool = False, allow_template_auto: bool = False) -> None:
        input_dir = Path(str(self.state["input"] or ""))
        template_dir = Path(str(self.state["templateDir"] or ""))
        if not input_dir.is_dir():
            self.state["sourceFiles"] = []
            self.state["selectedFiles"] = []
            self.state["explanationFiles"] = []
            return
        recursive = bool(self.state["recursive"])
        files = source_workbooks(input_dir, recursive=recursive)
        selected_before = set(self.state["selectedFiles"])
        template_name = Path(self.state["template"]).name if self.state["template"] else "未识别"
        self.state["sourceFiles"] = [{"path": str(path), "name": path.name, "template": template_name} for path in files]
        self.state["selectedFiles"] = [str(path) for path in files if not selected_before or str(path) in selected_before]
        self.state["explanationFiles"] = [{"path": str(path), "name": path.name} for path in explanation_files(input_dir, recursive=recursive)]
        self.state["detectedPeriod"] = detect_period(input_dir, recursive=recursive).period
        if not self.state["outputPinned"]:
            self.state["output"] = str(input_dir / "执行结果")
            self.state["outputAuto"] = True
        if template_dir.is_dir() and (force_template or (allow_template_auto and not self.state["templateManual"])):
            result = recommend_template(template_dir, input_dir, self.root / "data" / "模板索引.json", recursive=recursive)
            if result.template_path and result.matched:
                self.state["template"] = str(result.template_path)
                self.state["templateManual"] = False
                for item in self.state["sourceFiles"]:
                    item["template"] = result.template_path.name
            self.state["mixedTemplates"] = [name for name, _score in result.alternatives]
            self._log(result.details)

    def _flow_requires_template(self, flow_name: str) -> bool:
        """Use configured feature types, never just a displayed flow name."""
        steps = load_flow_steps(self.history_path, flow_name)
        mappings = {
            item.name: item
            for item in load_feature_mappings(self.history_path, Path("__无需模板__.xlsx"))
        }
        types = {
            mappings[step.feature_name].feature_type
            for step in steps
            if step.feature_name in mappings
        }
        return (types - {NAMED_RANGE_CHECK_FUNCTION}) not in (
            {MERGE_ORG_FILES_FUNCTION}, {COMBINE_SHEETS_FUNCTION}, set()
        )

    def _worker(self, action: str) -> None:
        started = time.monotonic()
        try:
            self._log("正在刷新待审核文件清单……")
            self._recognize(allow_template_auto=False)
            selected = [Path(path) for path in self.state["selectedFiles"]]
            if not selected:
                raise ValueError("请至少勾选一个待审核文件")
            flow = action[5:] if action.startswith("flow:") else {"audit": "汇总核查表校验", "summary": "汇总校验结果说明", "merge": "组合联合核查表", "combine": "组合工作表"}.get(action, action)
            requires_template = self._flow_requires_template(flow)
            if requires_template and not self.state["template"]:
                raise ValueError("请先选择模板文件")
            output = Path(str(self.state["output"] or Path(self.state["input"]) / "执行结果"))
            self._log("待处理文件：{} 个；正在读取执行流程……".format(len(selected)))
            self._log("计算引擎：{}".format(self.state["calculationEngine"]), detail=True)
            result = run_flow(
                flow_name=flow, template_path=Path(self.state["template"]) if requires_template else None,
                input_dir=Path(self.state["input"]), output_dir=output, history_path=self.history_path,
                config_path=self.history_path, external_path=Path(self.state["external"]) if self.state["external"] else None,
                recursive=bool(self.state["recursive"]), period=str(self.state.get("detectedPeriod") or ""),
                selected_files=selected, on_step=self._log_detail, write_flow_logs=bool(self.state["writeFlowLogs"]),
                calculation_engine=str(self.state["calculationEngine"]),
            )
            self.state["status"] = "流程完成"
            self.state["lastOutput"] = str(result.get("output") or "")
            self._log("流程完成：{}".format(self.state["lastOutput"] or "无文件输出"))
        except Exception as exc:
            self.state["status"] = "执行失败"
            self._log("执行失败：{}".format(exc))
        finally:
            self._log("任务结束，本次耗时：{:.1f} 秒".format(time.monotonic() - started))
            self.state["busy"] = False

    def _template_merge_worker(self, base: Path, joins: list[Path]) -> None:
        started = time.monotonic()
        try:
            output, copied, skipped = create_combined_template(base_template=base, source_templates=joins, on_step=self._log_detail)
            self.state["status"] = "联合模板已生成"
            self.state["lastOutput"] = str(output)
            self._log("联合模板：{}；复制 {} 张表，跳过 {} 张同名表".format(output, len(copied), len(skipped)))
        except Exception as exc:
            self.state["status"] = "制作联合模板失败"
            self._log("制作联合模板失败：{}".format(exc))
        finally:
            self._log("任务结束，本次耗时：{:.1f} 秒".format(time.monotonic() - started))
            self.state["busy"] = False

    @staticmethod
    def _as_bool(value: object) -> bool:
        return value if isinstance(value, bool) else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}

    def _log(self, text: str, detail: bool = False) -> None:
        item = {"text": "[{}] {}".format(datetime.now().strftime("%H:%M:%S"), text), "detail": bool(detail)}
        self.state["log"] = [*self.state["log"][-99:], item]

    def _log_detail(self, text: str) -> None:
        self._log(text, detail=True)
