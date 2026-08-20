from __future__ import annotations

import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.web_app import WebApi


class _Window:
    def __init__(self, selected: str) -> None:
        self.selected = selected
        self.calls: list[tuple[object, dict[str, object]]] = []

    def create_file_dialog(self, dialog_type: object, **kwargs: object) -> tuple[str, ...]:
        self.calls.append((dialog_type, kwargs))
        return (self.selected,)


class WebAppCompatibilityTests(unittest.TestCase):
    def test_old_pywebview_dialog_constants_are_supported(self) -> None:
        window = _Window("C:/审核输出")
        old_webview = SimpleNamespace(
            FOLDER_DIALOG=20,
            OPEN_DIALOG=10,
            windows=[window],
        )
        with patch.dict(sys.modules, {"webview": old_webview}):
            api = WebApi(ROOT)
            self.assertEqual(api.choose_folder("output"), "C:/审核输出")
            self.assertEqual(window.calls[0][0], 20)
            self.assertEqual(api.choose_file("external"), "C:/审核输出")
            self.assertEqual(window.calls[1][0], 10)

    def test_new_pywebview_file_dialog_enum_is_supported(self) -> None:
        window = _Window("C:/模板.xlsx")
        file_dialog = SimpleNamespace(FOLDER=101, OPEN=102)
        new_webview = SimpleNamespace(
            FileDialog=file_dialog,
            FOLDER_DIALOG=20,
            OPEN_DIALOG=10,
            windows=[window],
        )
        with patch.dict(sys.modules, {"webview": new_webview}):
            api = WebApi(ROOT)
            self.assertEqual(api.choose_folder("output"), "C:/模板.xlsx")
            self.assertEqual(window.calls[0][0], 101)
            self.assertEqual(api.choose_file("external"), "C:/模板.xlsx")
            self.assertEqual(window.calls[1][0], 102)

    def test_start_with_state_syncs_selected_files_before_starting(self) -> None:
        api = WebApi(ROOT)
        api.state["sourceFiles"] = [{
            "path": "C:/报送目录/机构A.xlsx",
            "name": "机构A.xlsx",
            "template": "测试模板.xlsx",
        }]
        with patch.object(api, "start", return_value=True) as start:
            result = api.start_with_state(
                "audit",
                {"input": "C:/报送目录"},
                ["C:/报送目录/机构A.xlsx"],
            )
        self.assertTrue(result)
        self.assertEqual(api.state["input"], "C:/报送目录")
        self.assertEqual(api.state["selectedFiles"], ["C:/报送目录/机构A.xlsx"])
        start.assert_called_once_with("audit", strict=True)

    def test_unpinned_output_follows_input_with_execution_results_name(self) -> None:
        api = WebApi(ROOT)
        api.update({"input": "C:/报送目录", "outputPinned": False})
        self.assertEqual(api.state["output"], str(Path("C:/报送目录") / "执行结果"))

    def test_flow_start_logs_before_background_worker_for_win7_compatibility(self) -> None:
        api = WebApi(ROOT)
        with patch("base_audit.web_app.threading.Thread") as thread:
            self.assertTrue(api.start("flow:汇总核查表校验"))
        last = api.state["log"][-1]
        self.assertIn("开始执行：汇总核查表校验", last["text"])
        self.assertFalse(last["detail"])
        thread.return_value.start.assert_called_once()


if __name__ == "__main__":
    unittest.main()
