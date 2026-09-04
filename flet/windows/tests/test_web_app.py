from __future__ import annotations

from pathlib import Path
import time

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from uos_audit.calculation import ENGINE_OFFICE_COM, preferred_engine
from uos_audit.name_config import load_feature_mappings
from uos_audit.web_app import WebApi


def _book(path: Path) -> None:
    book = Workbook()
    book.active.title = "数据"
    book.save(path)
    book.close()


def test_startup_uses_the_single_automatic_office_engine(tmp_path: Path, monkeypatch) -> None:
    statuses = {
        "formulas": {"available": True, "detail": "ok"},
        "ironcalc": {"available": True, "detail": "ok"},
        "libreoffice": {"available": True, "detail": "ok"},
        "excelCom": {"available": True, "detail": "ok"},
        "wpsCom": {"available": True, "detail": "ok"},
    }
    monkeypatch.setattr("uos_audit.web_app.engine_status", lambda: statuses)
    root = tmp_path / "uos_audit"
    root.mkdir()
    api = WebApi(root)
    assert api.state["calculationEngine"] == ENGINE_OFFICE_COM


def test_flet_controller_file_selection_and_engine_setting(tmp_path: Path, monkeypatch) -> None:
    root = tmp_path / "uos_audit"
    root.mkdir()
    input_dir = tmp_path / "source"
    input_dir.mkdir()
    _book(input_dir / "机构A_在线核查表.xlsx")

    status = {
        "excelCom": {"available": False, "detail": "not installed"},
        "wpsCom": {"available": True, "detail": "WPS"},
    }
    monkeypatch.setattr("uos_audit.web_app.engine_status", lambda: status)
    api = WebApi(root)
    api.update({"input": str(input_dir), "recursive": False})
    api.refresh_sources()

    assert api.state["output"].endswith("执行结果")
    assert [item["name"] for item in api.state["sourceFiles"]] == ["机构A_在线核查表.xlsx"]
    api.set_selected_files([api.state["sourceFiles"][0]["path"], "not-a-workbook.xlsx"])
    assert len(api.state["selectedFiles"]) == 1
    api.update({"calculationEngine": ENGINE_OFFICE_COM})
    assert api.state["calculationEngine"] == ENGINE_OFFICE_COM
    reopened = WebApi(root)
    assert reopened.state["calculationEngine"] == preferred_engine(reopened.state["engineStatus"])


def test_run_confirmation_preference_and_context(tmp_path: Path) -> None:
    root = tmp_path / "uos_audit"
    root.mkdir()
    api = WebApi(root)

    assert api.state["confirmBeforeRun"] is True
    api.update({"confirmBeforeRun": False})
    assert WebApi(root).state["confirmBeforeRun"] is False

    context = api.get_run_context("merge")
    assert context["flow"] == "组合联合核查表"
    assert context["templateRequired"] is False
    assert context["externalRequired"] is False


def test_flet_controller_dispatches_audit_flow_with_selected_files(tmp_path: Path, monkeypatch) -> None:
    root = tmp_path / "uos_audit"
    root.mkdir()
    input_dir = tmp_path / "source"
    input_dir.mkdir()
    source_path = input_dir / "机构A_在线核查表.xlsx"
    _book(source_path)
    template_path = tmp_path / "模板.xlsx"
    _book(template_path)
    external_path = tmp_path / "外部.xlsx"
    _book(external_path)

    received: dict[str, object] = {}

    def fake_run_flow(**kwargs):
        received.update(kwargs)
        return {"output": str(tmp_path / "执行结果" / "结果.xlsx"), "log": ""}

    monkeypatch.setattr("uos_audit.web_app.run_flow", fake_run_flow)
    api = WebApi(root)
    api.update({
        "input": str(input_dir), "template": str(template_path),
        "external": str(external_path), "calculationEngine": ENGINE_OFFICE_COM,
    })
    api.refresh_sources()
    api.set_selected_files([str(source_path)])

    assert api.start("audit")
    deadline = time.monotonic() + 2
    while api.state["busy"] and time.monotonic() < deadline:
        time.sleep(0.01)

    assert not api.state["busy"]
    assert received["flow_name"] == "汇总核查表校验"
    assert received["template_path"] == template_path
    assert received["external_path"] == external_path
    assert received["selected_files"] == [source_path]
    assert received["calculation_engine"] == ENGINE_OFFICE_COM
    assert api.state["lastOutput"].endswith("结果.xlsx")


def test_custom_standalone_flow_does_not_require_a_template(tmp_path: Path) -> None:
    """A renamed combine flow must be detected from its configured module."""
    root = tmp_path / "uos_audit"
    root.mkdir()
    api = WebApi(root)
    editor = api.get_config_editor_data()
    saved = api.save_config_editor_draft({
        "customModules": [],
        "customFlows": [{
            "流程名": "自定义组合", "顺序": "10", "功能名": "组合联合核查表",
            "启用": "是", "失败后处理": "跳过", "是否输出结果": "是",
            "输出文件名": "自定义组合", "输入": "源数据目录", "输出": "联合核查表", "备注": "",
        }],
        "flowDisplay": {"自定义组合": True},
        "combineSheetsPlans": editor["combineSheetsPlans"],
        "activeCombineSheetsPlanId": editor["activeCombineSheetsPlanId"],
    })
    assert saved["ok"]
    assert not api._flow_requires_template("自定义组合")


def test_flet_settings_draft_persists_custom_module(tmp_path: Path) -> None:
    root = tmp_path / "uos_audit"
    root.mkdir()
    api = WebApi(root)
    editor = api.get_config_editor_data()
    assert "customModules" in editor
    draft = {
        "customModules": [{
            "功能名": "测试固定行汇总", "执行模块": "汇总_固定行汇总",
            "命名区域名": "固定行汇总区域", "输入": "源数据目录",
            "输出": "测试汇总", "备注": "测试",
        }],
        "customFlows": [{
            "流程名": "测试流程", "顺序": "10", "功能名": "测试固定行汇总",
            "启用": "是", "失败后处理": "跳过", "是否输出结果": "是",
            "输出文件名": "测试流程", "输入": "源数据目录", "输出": "测试汇总", "备注": "",
        }],
        "flowDisplay": {"测试流程": True},
        "combineSheetsPlans": editor["combineSheetsPlans"],
        "activeCombineSheetsPlanId": editor["activeCombineSheetsPlanId"],
    }
    saved = api.save_config_editor_draft(draft)
    assert saved["ok"]
    names = {item.name for item in load_feature_mappings(api.history_path, root / "不存在.xlsx")}
    assert "测试固定行汇总" in names
    api.update({"showCustomFeatures": True})
    assert api.get_custom_features()[0]["flow"] == "测试流程"


def test_flet_controller_inspects_template_formula_compatibility(tmp_path: Path) -> None:
    root = tmp_path / "uos_audit"
    root.mkdir()
    template = Workbook()
    sheet = template.active
    sheet.title = "报表"
    sheet["A1"] = "=INDIRECT(\"B1\")"
    template.defined_names.add(DefinedName("校验区域", attr_text="'报表'!$A$1"))
    template_path = tmp_path / "模板.xlsx"
    template.save(template_path)
    template.close()

    api = WebApi(root)
    api.update({"template": str(template_path)})
    api.inspect_current_template()

    capability = api.state["formulaCapability"]
    assert capability["formulaCount"] == 1
    assert not capability["pythonSupported"]
    assert "INDIRECT" in capability["blockers"][0]
