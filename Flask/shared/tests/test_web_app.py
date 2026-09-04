from __future__ import annotations

import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.web_app import WebApi
from base_audit.name_config import HISTORY_WORKBOOK_NAME, initialize_config
from base_audit.history import HISTORY_AUDIT_SHEET


class WebAppCompatibilityTests(unittest.TestCase):
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

    def test_calculation_engine_setting_is_persisted_in_state(self) -> None:
        with TemporaryDirectory() as folder:
            api = WebApi(Path(folder))
            api.update({"calculationEngine": "WPS 表格"})
            self.assertEqual("WPS 表格", api.state["calculationEngine"])
            api.update({"calculationEngine": "不存在的引擎"})
            self.assertEqual("WPS 表格", api.state["calculationEngine"])

    def test_confirm_before_run_defaults_on_and_toggles(self) -> None:
        with TemporaryDirectory() as folder:
            api = WebApi(Path(folder))
            # 默认开启：误点主按钮先看到文件清单。
            self.assertTrue(api.state["confirmBeforeRun"])
            api.update({"confirmBeforeRun": False})
            self.assertFalse(api.state["confirmBeforeRun"])
            # 重启后保持关闭（写入用户设置）。
            reloaded = WebApi(Path(folder))
            self.assertFalse(reloaded.state["confirmBeforeRun"])

    def test_history_page_is_paginated_and_filterable(self) -> None:
        with TemporaryDirectory() as folder:
            root = Path(folder)
            config = root / HISTORY_WORKBOOK_NAME
            initialize_config(config)
            from openpyxl import load_workbook
            workbook = load_workbook(config)
            sheet = workbook[HISTORY_AUDIT_SHEET]
            sheet.append(("机构A.xlsx", "贷款表", "A1", "错误", "余额", "余额异常", "10", "8", "2", "规则A", "说明A", "通过"))
            sheet.append(("机构B.xlsx", "贷款表", "B2", "软性", "利率", "利率核实", "", "", "", "规则B", "说明B", "待核实"))
            workbook.save(config)
            workbook.close()
            api = WebApi(root)
            result = api.get_history_page({"page": 1, "pageSize": 10, "keyword": "余额", "errorType": "错误"})
        self.assertEqual(1, result["total"])
        self.assertEqual("机构A.xlsx", result["items"][0]["工作簿名"])
        self.assertIn("软性", result["errorTypes"])
        self.assertIn("通过", result["opinions"])

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
