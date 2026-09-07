from __future__ import annotations

import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.discovery import (
    TemplateCatalog,
    _compact,
    detect_period,
    recommend_template,
    source_workbooks,
)
from base_audit.settings import SettingsStore, UserSettings, add_favorite, remember_path
from base_audit.template import normalize_template_name


WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="数据表" sheetId="1"/></sheets>
</workbook>"""


def make_xlsx(path: Path, *sheet_names: str) -> None:
    names = sheet_names or ("数据表",)
    sheets = "".join(
        f'<sheet name="{name}" sheetId="{index}"/>'
        for index, name in enumerate(names, start=1)
    )
    xml = WORKBOOK_XML.replace(
        '<sheets><sheet name="数据表" sheetId="1"/></sheets>',
        f"<sheets>{sheets}</sheets>",
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", xml)


class DiscoveryTests(unittest.TestCase):
    def test_recursive_depth_levels(self) -> None:
        """0 层=仅根目录；1 层=下一级文件夹；负数/True=最深处。"""
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "root.xlsx").touch()
            level1 = root / "机构甲"
            level1.mkdir()
            (level1 / "a.xlsx").touch()
            level2 = level1 / "2026-06"
            level2.mkdir()
            (level2 / "b.xlsx").touch()
            assert len(source_workbooks(root, recursive=0)) == 1
            assert len(source_workbooks(root, recursive=1)) == 2
            assert len(source_workbooks(root, recursive=2)) == 3
            assert len(source_workbooks(root, recursive=-1)) == 3
            assert len(source_workbooks(root, recursive=True)) == 3

    def test_extra_files_join_pending_list(self) -> None:
        """手动追加的文件（源数据目录之外）进入待处理清单并可被勾选。"""
        from base_audit.service import _source_files
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "in_list.xlsx").touch()
            outside = Path(folder) / "追加.xlsx"
            outside.touch()
            files = _source_files(root, extra_files=[outside])
            names = [path.name for path in files]
            assert "in_list.xlsx" in names
            assert "追加.xlsx" in names
            # 勾选追加文件时不在 allowed 集合会被拒绝——放行后应成功
            picked = _source_files(root, selected_files=[outside], extra_files=[outside])
            assert [path.name for path in picked] == ["追加.xlsx"]


    def test_period_is_read_from_source_filenames(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "001_甲银行_单位贷款_2026年6月.xlsx").touch()
            (root / "002_乙银行_单位贷款_20260630.xlsx").touch()
            result = detect_period(root)
        self.assertEqual(result.period, "2026-06")
        self.assertEqual(result.source, "源文件名")
        self.assertFalse(result.conflict)

    def test_period_conflict_is_not_silently_selected(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "001_甲银行_单位贷款_2026-05.xlsx").touch()
            (root / "002_乙银行_单位贷款_2026-06.xlsx").touch()
            result = detect_period(root)
        self.assertTrue(result.conflict)
        self.assertEqual(result.period, "")
        self.assertIn("多个数据期", result.details)

    def test_period_falls_back_to_folder_name(self) -> None:
        with tempfile.TemporaryDirectory(prefix="报送_2026-07_") as folder:
            root = Path(folder)
            (root / "甲银行_单位贷款.xlsx").touch()
            result = detect_period(root)
        self.assertEqual(result.period, "2026-07")
        self.assertEqual(result.source, "文件夹名称")

    def test_recursive_source_files_include_institution_folders_but_not_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "机构甲").mkdir()
            (root / "历史材料_skip" / "机构审核副本").mkdir(parents=True)
            (root / "机构甲" / "甲银行_单位贷款_202607.xlsx").touch()
            (root / "历史材料_skip" / "机构审核副本" / "甲银行_审核版.xlsx").touch()
            files = source_workbooks(root, recursive=True)
        self.assertEqual([path.name for path in files], ["甲银行_单位贷款_202607.xlsx"])

    def test_recursive_source_files_skip_current_execution_output_folder(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "执行结果").mkdir()
            (root / "机构甲").mkdir()
            (root / "执行结果" / "甲银行_审核版.xlsx").touch()
            (root / "机构甲" / "甲银行_单位贷款_202607.xlsx").touch()
            files = source_workbooks(root, recursive=True)
        self.assertEqual([path.name for path in files], ["甲银行_单位贷款_202607.xlsx"])

    def test_template_catalog_reuses_unchanged_index(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template_dir = root / "templates"
            template_dir.mkdir()
            template = template_dir / "！单位贷款202606.xlsx"
            make_xlsx(template, "单位贷款")
            index = root / "data" / "模板索引.json"
            profiles = TemplateCatalog(template_dir, index).profiles()
            payload = json.loads(index.read_text(encoding="utf-8"))
            second = TemplateCatalog(template_dir, index).profiles()
        self.assertEqual(profiles, second)
        self.assertEqual(payload["template_dir"], str(template_dir.resolve()))
        self.assertEqual(payload["templates"][0]["sheet_names"], ["单位贷款"])

    def test_template_catalog_does_not_reuse_another_directory(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            first_dir = root / "模板甲"
            second_dir = root / "模板乙"
            first_dir.mkdir()
            second_dir.mkdir()
            make_xlsx(first_dir / "！单位贷款.xlsx", "单位贷款")
            make_xlsx(second_dir / "！个人贷款.xlsx", "个人贷款")
            index = root / "模板索引.json"
            TemplateCatalog(first_dir, index).profiles()
            profiles = TemplateCatalog(second_dir, index).profiles()
            payload = json.loads(index.read_text(encoding="utf-8"))
        self.assertEqual([Path(item.path).name for item in profiles], ["！个人贷款.xlsx"])
        self.assertEqual(payload["template_dir"], str(second_dir.resolve()))

    def test_name_and_sheet_recommend_the_template(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template_dir = root / "templates"
            input_dir = root / "input"
            template_dir.mkdir()
            input_dir.mkdir()
            make_xlsx(template_dir / "！金融基础数据-单位贷款202606.xlsx", "单位贷款")
            make_xlsx(template_dir / "！金融基础数据-个人贷款202606.xlsx", "个人贷款")
            make_xlsx(input_dir / "001_甲银行_金融基础数据-单位贷款_2026-06.xlsx", "单位贷款")
            result = recommend_template(template_dir, input_dir, root / "index.json")
        self.assertTrue(result.matched)
        self.assertEqual(result.template_path.name, "！金融基础数据-单位贷款202606.xlsx")

    def test_combined_workbook_is_matched_by_sheet_set_not_filename(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template_dir = root / "templates"
            input_dir = root / "input"
            template_dir.mkdir()
            input_dir.mkdir()
            make_xlsx(template_dir / "！个人贷款模板.xlsx", "个人贷款", "个人客户")
            make_xlsx(template_dir / "！单位贷款模板.xlsx", "单位贷款", "单位客户")
            make_xlsx(
                template_dir / "！任意名称.xlsx",
                "个人贷款", "个人客户", "单位贷款", "单位客户", "集中系统数据", "参照表",
            )
            make_xlsx(
                input_dir / "机构A_合并_2026-07-31.xlsx",
                "个人贷款", "个人客户", "单位贷款", "单位客户",
            )
            result = recommend_template(template_dir, input_dir, root / "index.json")
        self.assertTrue(result.matched)
        self.assertEqual(result.template_path.name, "！任意名称.xlsx")
        self.assertIn("工作表集合匹配率", result.details)
        self.assertIn("已忽略未能识别", result.details)

    def test_unmatched_summary_workbook_first_is_ignored(self) -> None:
        # “汇总信息”这类非报送工作簿可能排在机构文件之前；识别应基于
        # 全部源文件的共识，而不是只看第一个文件。
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template_dir = root / "templates"
            input_dir = root / "input"
            template_dir.mkdir()
            input_dir.mkdir()
            make_xlsx(template_dir / "！单位贷款模板.xlsx", "单位贷款")
            make_xlsx(template_dir / "！个人贷款模板.xlsx", "个人贷款")
            make_xlsx(input_dir / "00_汇总信息.xlsx", "汇总")
            make_xlsx(input_dir / "机构A_报送.xlsx", "单位贷款")
            make_xlsx(input_dir / "机构B_报送.xlsx", "单位贷款")
            result = recommend_template(template_dir, input_dir, root / "index.json")
        self.assertTrue(result.matched)
        self.assertEqual(result.template_path.name, "！单位贷款模板.xlsx")
        self.assertIn("已忽略未能识别", result.details)

    def test_mixed_reports_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template_dir = root / "templates"
            input_dir = root / "input"
            template_dir.mkdir()
            input_dir.mkdir()
            make_xlsx(template_dir / "！金融基础数据-单位贷款202606.xlsx", "单位贷款")
            make_xlsx(template_dir / "！金融基础数据-个人贷款202606.xlsx", "个人贷款")
            make_xlsx(input_dir / "001_甲银行_金融基础数据-单位贷款_2026-06.xlsx", "单位贷款")
            make_xlsx(input_dir / "001_甲银行_金融基础数据-个人贷款_2026-06.xlsx", "个人贷款")
            result = recommend_template(template_dir, input_dir, root / "index.json")
        self.assertFalse(result.matched)
        self.assertIn("多种报表", result.details)

    def test_plain_named_template_is_recognized_without_bang_prefix(self) -> None:
        # “！”前缀只是命名习惯，不应作为模板识别的硬性条件。
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template_dir = root / "templates"
            input_dir = root / "input"
            template_dir.mkdir()
            input_dir.mkdir()
            make_xlsx(template_dir / "金融基础数据-单位贷款202606.xlsx", "单位贷款")
            make_xlsx(template_dir / "~$临时锁文件.xlsx", "单位贷款")
            make_xlsx(input_dir / "001_甲银行_金融基础数据-单位贷款_2026-06.xlsx", "单位贷款")
            profiles = TemplateCatalog(template_dir, root / "index.json").profiles()
            result = recommend_template(template_dir, input_dir, root / "index.json")
        self.assertEqual(
            [Path(item.path).name for item in profiles],
            ["金融基础数据-单位贷款202606.xlsx"],
        )
        self.assertTrue(result.matched)

class SettingsTests(unittest.TestCase):
    def test_settings_round_trip_and_recent_limit(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            values: list[str] = []
            for index in range(12):
                values = remember_path(values, str(root / str(index)))
            settings = UserSettings(
                last_input_dir=str(root),
                last_template_dir=str(root / "模板"),
                last_external_file=str(root / "外部数据.xlsx"),
                favorite_input_dirs=add_favorite([], "月报", str(root / "月报")),
                favorite_template_dirs=add_favorite(
                    [], "正式模板", str(root / "模板")
                ),
                favorite_external_files=add_favorite(
                    [], "外部数据", str(root / "外部数据.xlsx")
                ),
                recent_input_dirs=values,
                recent_template_dirs=[str(root / "模板")],
                write_flow_logs=False,
            )
            store = SettingsStore(root / "用户设置.json")
            store.save(settings)
            loaded = store.load()
        self.assertEqual(loaded.favorite_input_dirs[0].name, "月报")
        self.assertEqual(loaded.favorite_template_dirs[0].name, "正式模板")
        self.assertEqual(loaded.favorite_external_files[0].name, "外部数据")
        self.assertTrue(loaded.last_template_dir.endswith("模板"))
        self.assertTrue(loaded.last_external_file.endswith("外部数据.xlsx"))
        self.assertFalse(loaded.output_pinned)
        self.assertFalse(loaded.write_flow_logs)
        self.assertEqual(len(loaded.recent_template_dirs), 1)
        self.assertEqual(len(loaded.recent_input_dirs), 10)
        self.assertTrue(loaded.recent_input_dirs[0].endswith("11"))
