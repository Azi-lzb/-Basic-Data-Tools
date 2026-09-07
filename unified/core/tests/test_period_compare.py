import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from openpyxl import Workbook, load_workbook

from base_audit import period_compare as pc


def _make_period_workbook(path: Path, rows, title="20201"):
    book = Workbook()
    sheet = book.active
    sheet.title = title
    sheet.append([f"{title}金融机构（法人）基本情况统计表", "", ""])
    sheet.append([None, None, None])
    sheet.append(["指标编号", "指标名称", "本期情况"])
    for row in rows:
        sheet.append(row)
    book.save(path)


def _make_config(path: Path):
    book = Workbook()
    sheet = book.active
    sheet.title = pc.CONFIG_INDICATOR_SHEET
    sheet.append(["指标代码", "指标名称", "数据属性", "是否不转换单位", "是否与大集中核对", "大集中报表查询指标名称", "禁用"])
    sheet.append([20201001, "金融机构名称", "文字", None, None, None])
    sheet.append([20202001, "存款余额", "余额", None, "是", "各项存款"])
    sheet.append([20202002, "机构户数", "个数", "是", None, None])
    sheet.append([20202003, "文字说明", "文字", None, None, None])
    org = book.create_sheet(pc.CONFIG_ORG_SHEET)
    org.append(["机构名称", "社会信用代码", "机构类别", "承接行", "归属行", "报表项目", "禁用"])
    org.append(["甲银行", "91440000AA", "农商行", "承接一部", "惠州", "j01（甲银行）"])
    alert = book.create_sheet(pc.CONFIG_ALERT_SHEET)
    alert.append(["序号", "变幅下限（小数，0.3=30%）", "备注", "是否整行填充"])
    alert.append([1, -0.3, "", None])
    alert.append([2, 0.0, "", None])
    alert.append([3, 0.3, "增幅[30%,50%)", "是"])
    alert.append([4, 1.0, "增幅[1倍,5倍)", None])
    book.save(path)


def _make_central(path: Path, *, value_yi=0.2831236):
    """集中系统数据单位为亿元；value_yi=0.2831236 亿元 = 283.1236 万元。"""
    book = Workbook()
    sheet = book.active
    sheet.title = "集中系统数据"
    sheet.append([None, None, None])
    sheet.append([None, " ", "各项存款"])
    sheet.append([None, "j01（甲银行）", value_yi])
    ref = book.create_sheet("参照表")
    ref.append(["统一社会信用代码", "报表项目", "机构地区"])
    ref.append(["91440000AA", "j01（甲银行）", "惠州"])
    book.save(path)


def _make_config_v2(path: Path):
    """新版 4 表配置：表头名驱动、短写/名称引用表达式、级别、代码清单规则。

    列顺序故意与旧格式不同，且包含未知类型/非法级别/语法错误/未知名称等
    异常行，用于同时验证加载期校验。
    """
    book = Workbook()
    sheet = book.active
    sheet.title = pc.CONFIG_INDICATOR_SHEET
    sheet.append(["禁用", "指标代码", "指标名称", "数据属性", "不转换单位", "大集中核对", "大集中指标名称", "备注"])
    sheet.append([None, 20201001, "金融机构名称", "文字", None, None, None])
    sheet.append([None, 20202001, "各项存款", "余额", None, "是", "各项存款"])
    sheet.append([None, 20202002, "各项贷款", "余额", None, None, None])
    sheet.append([None, 20202003, "资产总计", "余额", None, None, None])
    sheet.append([None, 20202004, "负债合计", "余额", None, None, None])
    sheet.append([None, 20202005, "所有者权益", "余额", None, None, None])
    sheet.append([None, 20202006, "备用指标", "余额", None, None, None])
    sheet.append([None, 20202007, "备用指标", "余额", None, None, None])   # 重名 → 名称引用禁用
    org = book.create_sheet(pc.CONFIG_ORG_SHEET)
    org.append(["机构名称", "社会信用代码", "机构类别", "承接行", "地区", "报表项目", "禁用"])
    org.append(["甲银行", "91440000AA", "农商行", "承接一部", "惠州", "j01（甲银行）"])
    alert = book.create_sheet(pc.CONFIG_ALERT_SHEET)
    alert.append(["变幅下限（小数，0.3=30%）", "备注", "整行填充"])
    alert.append([-0.3, "", None])
    alert.append([0.3, "增幅[30%,50%)", "是"])
    rules = book.create_sheet(pc.CONFIG_RULE_SHEET)
    rules.append(["规则编号", "类型", "描述", "规则内容", "级别", "禁用", "备注"])
    rules.append(["R001", "累计不降(当年)", "当年累计指标比上期不应减少。", "20202002,20202003", "提示", None, None])
    rules.append(["R002", "表达式", "资产负债表不平衡", "[资产总计] <> [负债合计] + [所有者权益]", "错误", None, None])
    rules.append(["R003", "表达式", "贷款大于存款", "or([各项贷款] > [各项存款] , [各项贷款] > 99999)", "核实", None, None])
    rules.append(["R004", "表达式", "禁用规则不执行", "[20202001] > 0", "提示", "是", None])
    rules.append(["R005", "未支持类型", "未知类型应跳过", "[20202001] > 0", "提示", None, None])
    rules.append(["R006", "表达式", "级别无效按提示", "[20202001] > 0", "严重", None, None])
    rules.append(["R007", "表达式", "语法错误进告警", "[20202001] >", "错误", None, None])
    rules.append(["R008", "表达式", "未知名称进告警", "[不存在的指标] > 0", "错误", None, None])
    book.save(path)


class PeriodCompareV2ConfigTests(unittest.TestCase):
    """新版（4 表、表头名驱动）配置簿的加载、校验与规则执行。"""

    def setUp(self):
        import tempfile

        self.tmp = Path(tempfile.mkdtemp())
        self.config_path = self.tmp / "config_v2.xlsx"
        _make_config_v2(self.config_path)
        self.config = pc.load_period_config(self.config_path)

    def _prepare_periods(self):
        """两期数据：贷款下降（累计不降命中）、资产负债不平衡（表达式命中）。"""
        cur_dir = self.tmp / "v2_cur"
        pre_dir = self.tmp / "v2_pre"
        cur_dir.mkdir()
        pre_dir.mkdir()
        common_head = [
            [20201001, "金融机构名称", "甲银行"],
            [20201002, "金融机构代码", "91440000AA"],
        ]
        rows_cur = common_head + [
            [20202001, "各项存款", 500_000],     # 50 万元
            [20202002, "各项贷款", 800_000],     # 80 万元（上期 100 万 → 下降命中）
            [20202003, "资产总计", 1_300_000],   # 130 万
            [20202004, "负债合计", 700_000],     # 70 万
            [20202005, "所有者权益", 500_000],   # 50 万 → 70+50 ≠ 130 不平衡
        ]
        rows_pre = common_head + [
            [20202001, "各项存款", 400_000],
            [20202002, "各项贷款", 1_000_000],
            [20202003, "资产总计", 1_200_000],
            [20202004, "负债合计", 700_000],
            [20202005, "所有者权益", 500_000],
        ]
        for path, rows in (
            (cur_dir / "91440000AA#2026-06-30#01#20202#甲银行.xlsx", rows_cur),
            (pre_dir / "91440000AA#2026-03-31#01#20202#甲银行.xlsx", rows_pre),
        ):
            book = Workbook()
            sheet = book.active
            sheet.title = "20202"
            sheet.append(["t", "", ""])
            sheet.append([None, None, None])
            sheet.append(["指标编号", "指标名称", "本期情况"])
            for row in rows:
                sheet.append(row)
            book.save(path)
        return (
            pc.load_period_directory(cur_dir, label="当期"),
            pc.load_period_directory(pre_dir, label="上期"),
        )

    def test_header_driven_loading_with_warnings(self):
        config = self.config
        # 表头名驱动：禁用列在最前也能正确读取
        self.assertIn("20202001", config.indicators)
        self.assertEqual(config.indicators["20202001"].central_name, "各项存款")
        self.assertEqual(config.orgs["甲银行"].region, "惠州")
        self.assertEqual(len(config.alerts), 2)
        # R005 未知类型跳过、R006 非法级别回落、R007 语法错误、R008 未知名称、重名指标
        self.assertFalse(any(r.rule_id == "R005" for r in config.rules))
        self.assertEqual(next(r for r in config.rules if r.rule_id == "R006").level, "提示")
        warnings = "\n".join(config.warnings)
        self.assertIn("未知类型", warnings)
        self.assertIn("级别", warnings)
        self.assertIn("语法错误", warnings)
        self.assertIn("不存在的指标", warnings)
        self.assertIn("重复", warnings)
        # 禁用规则保留在配置里但启用规则不含它
        self.assertTrue(next(r for r in config.rules if r.rule_id == "R004").disabled)
        self.assertEqual(len([r for r in config.effective_rules() if not r.disabled]), 6)

    def test_rules_hit_with_short_name_and_lowercase_or(self):
        config = self.config
        cur, pre = self._prepare_periods()
        rows = pc.compare_periods(cur, pre, config)
        pc.apply_special_rules(rows, cur, pre, config)
        pc.apply_complex_rules(rows, cur, pre, config, on_step=lambda _s: None)
        by_code = {row["指标编码"]: row for row in rows}
        # 累计不降（代码清单）：贷款下降命中，级别=提示
        self.assertEqual(by_code["20202002"]["是否说明"], "当年累计指标比上期不应减少。")
        self.assertEqual(by_code["20202002"]["级别"], "提示")
        # 名称引用 + 小写 or() 的表达式命中，级别=核实
        self.assertEqual(by_code["20202001"]["是否说明"], "贷款大于存款")
        self.assertEqual(by_code["20202001"]["级别"], "核实")
        # 资产负债不平衡：名称引用表达式命中所有涉及指标行，级别=错误，计算过程带规则编号
        for code in ("20202003", "20202004", "20202005"):
            self.assertEqual(by_code[code]["是否说明"], "资产负债表不平衡")
            self.assertEqual(by_code[code]["级别"], "错误")
            self.assertIn("R002", by_code[code]["计算过程"])

    def test_v2_accumulation_skips_cross_year(self):
        config = self.config
        d1 = self.tmp / "v2x_cur"
        d2 = self.tmp / "v2x_pre"
        d1.mkdir()
        d2.mkdir()
        for path, value in (
            (d1 / "91440000AA#2026-06-30#01#20202#甲银行.xlsx", 100.0),
            (d2 / "91440000AA#2025-12-31#01#20202#甲银行.xlsx", 500.0),
        ):
            book = Workbook()
            sheet = book.active
            sheet.title = "20202"
            sheet.append(["t", "", ""])
            sheet.append([None, None, None])
            sheet.append(["指标编号", "指标名称", "本期情况"])
            sheet.append([20201001, "金融机构名称", "甲银行"])
            sheet.append([20201002, "金融机构代码", "91440000AA"])
            sheet.append([20202002, "各项贷款", value])
            book.save(path)
        cur = pc.load_period_directory(d1, label="当期")
        pre = pc.load_period_directory(d2, label="上期")
        rows = pc.compare_periods(cur, pre, config)
        pc.apply_special_rules(rows, cur, pre, config)
        row = next(r for r in rows if r["指标编码"] == "20202002")
        self.assertEqual(row["是否说明"], "")
        self.assertIn("跨年累计不比较", row["计算过程"])

    def test_central_tolerance_configurable(self):
        cur, _pre = self._prepare_periods()
        central_path = self.tmp / "central.xlsx"
        _make_central(central_path, value_yi=0.005005)   # 50.05 万元，基础 50 万元 → 差 500 元
        rows = pc.compare_central(cur, central_path, self.config)
        deposit = next(r for r in rows if r["指标编码"] == "20202001")
        self.assertEqual(deposit["是否说明"], "差异超过100元")
        # 容差放宽到 0.1 万元（1000 元）后不再标记
        rows = pc.compare_central(cur, central_path, self.config, tolerance=0.1)
        deposit = next(r for r in rows if r["指标编码"] == "20202001")
        self.assertEqual(deposit["是否说明"], "")

    def test_run_accepts_tolerance_and_outputs_level_column(self):
        cur_dir, pre_dir = self._prepare_periods()
        out_dir = self.tmp / "v2_out"
        output = pc.run_period_compare(
            current_dir=cur_dir,
            previous_dir=pre_dir,
            central_path=None,
            output_dir=out_dir,
            config_path=self.config_path,
            central_tolerance_yuan=1000.0,
        )
        self.assertTrue(output.is_file())
        book = load_workbook(output, read_only=True)
        try:
            headers = [cell.value for cell in next(book["两期对比"].iter_rows(max_row=1))]
            self.assertEqual(headers, pc.PERIOD_SHEET_HEADERS)
            self.assertIn("级别", headers)
        finally:
            book.close()


class PeriodCompareTests(unittest.TestCase):
    def setUp(self):
        import tempfile

        self.tmp = Path(tempfile.mkdtemp())
        self._pair_seq = 0

    def _prepare_periods(self):
        cur_dir = self.tmp / "cur"
        pre_dir = self.tmp / "pre"
        cur_dir.mkdir()
        pre_dir.mkdir()
        rows_cur = [
            [20201001, "金融机构名称", "甲银行"],
            [20201002, "金融机构代码", "91440000AA"],
            [20202001, "存款余额", 2_831_236],      # 元 → 283.1236 万元
            [20202002, "机构户数", 12],
            [20202003, "文字说明", "正常"],
            [20202004, "当期独有", 5],
        ]
        rows_pre = [
            [20201001, "金融机构名称", "甲银行"],
            [20201002, "金融机构代码", "91440000AA"],
            [20202001, "存款余额", 2_000_000],      # → 200 万元，环比 +41.56%
            [20202002, "机构户数", 12],
            [20202003, "文字说明", "变更后"],
            [20202009, "上期独有", 7],
        ]
        _make_period_workbook(cur_dir / "91440000AA#2026-06-30#01#20201#甲银行.xlsx", rows_cur)
        _make_period_workbook(pre_dir / "91440000AA#2026-03-31#01#20201#甲银行.xlsx", rows_pre)
        return cur_dir, pre_dir

    def test_filename_parsing_and_pairing(self):
        cur_dir, pre_dir = self._prepare_periods()
        current = pc.load_period_directory(cur_dir, label="当期")
        previous = pc.load_period_directory(pre_dir, label="上期")
        self.assertEqual(current[("甲银行", "20202001")].value, 2_831_236)
        self.assertIn(("甲银行", "20202004"), current)
        self.assertNotIn(("甲银行", "20202004"), previous)
        self.assertIn(("甲银行", "20202009"), previous)

    def test_pair_list_expands_banks_workbook_by_credit_code_sheet(self):
        """按报表划分时，一个工作簿内每个机构子表都必须单独配对展示。"""
        cur_dir = self.tmp / "banks_cur"
        pre_dir = self.tmp / "banks_pre"
        cur_dir.mkdir()
        pre_dir.mkdir()

        def make_banks_book(path: Path, date: str) -> None:
            book = Workbook()
            book.remove(book.active)
            for credit_code, org_name in (("91440000AA", "甲银行"), ("91440000BB", "乙银行")):
                sheet = book.create_sheet(credit_code)
                sheet.append(["20201报表", "", ""])
                sheet.append([None, None, None])
                sheet.append(["指标编号", "指标名称", "本期情况"])
                sheet.append([20201001, "金融机构名称", org_name])
                sheet.append([20201002, "金融机构代码", credit_code])
                sheet.append([20201003, "测试指标", 1])
            book.save(path)

        make_banks_book(cur_dir / "banks#2026-06-30#01#20201.xlsx", "2026-06-30")
        make_banks_book(pre_dir / "banks#2026-03-31#01#20201.xlsx", "2026-03-31")
        pairs = pc.list_period_pairs(cur_dir, pre_dir)

        self.assertEqual(len(pairs), 2)
        self.assertEqual({row["orgName"] for row in pairs}, {"甲银行", "乙银行"})
        self.assertTrue(all(row["form"] == "20201" for row in pairs))
        self.assertTrue(all(row["curDate"] == "2026-06-30" for row in pairs))
        self.assertTrue(all(row["preDate"] == "2026-03-31" for row in pairs))
        self.assertEqual({row["curSheet"] for row in pairs}, {"91440000AA", "91440000BB"})

    def test_pair_list_keeps_report_name_sheet_as_form(self):
        """按机构划分时，中文报表名子表应作为表单身份保留。"""
        cur_dir = self.tmp / "reports_cur"
        pre_dir = self.tmp / "reports_pre"
        cur_dir.mkdir()
        pre_dir.mkdir()

        def make_reports_book(path: Path) -> None:
            book = Workbook()
            sheet = book.active
            sheet.title = "金融机构信贷收支表"
            sheet.append(["报表标题", "", ""])
            sheet.append([None, None, None])
            sheet.append(["指标编号", "指标名称", "本期情况"])
            sheet.append([20201001, "金融机构名称", "甲银行"])
            sheet.append([20201002, "金融机构代码", "91440000AA"])
            sheet.append([20202001, "测试指标", 1])
            book.save(path)

        make_reports_book(cur_dir / "reports#91440000AA#2026-06-30#01#甲银行.xlsx")
        make_reports_book(pre_dir / "reports#91440000AA#2026-03-31#01#甲银行.xlsx")
        pairs = pc.list_period_pairs(cur_dir, pre_dir)

        self.assertEqual(len(pairs), 1)
        self.assertEqual(pairs[0]["orgName"], "甲银行")
        self.assertEqual(pairs[0]["form"], "金融机构信贷收支表")
        self.assertEqual(pairs[0]["curSheet"], "金融机构信贷收支表")

    def test_zero_rows_are_dropped_like_vba(self):
        """VBA 口径：两期都为 0/空的行不进入比较结果。"""
        cur_dir, pre_dir = self._prepare_periods()
        config_path = self.tmp / "config.xlsx"
        _make_config(config_path)
        # 追加两期均为 0 的指标行
        book_path = cur_dir / "91440000AA#2026-06-30#01#20201#甲银行.xlsx"
        from openpyxl import load_workbook as lw
        book = lw(book_path)
        book["20201"].append([20202006, "两期零值", 0])
        book.save(book_path)
        book_path2 = pre_dir / "91440000AA#2026-03-31#01#20201#甲银行.xlsx"
        book = lw(book_path2)
        book["20201"].append([20202006, "两期零值", 0])
        book.save(book_path2)
        config = pc.load_period_config(config_path)
        rows = pc.compare_periods(
            pc.load_period_directory(cur_dir, label="当期"),
            pc.load_period_directory(pre_dir, label="上期"),
            config,
        )
        self.assertNotIn("20202006", {row["指标编码"] for row in rows})
        # 20201001/20201002（机构名称/代码）保留输出，与 VBA 一致
        self.assertIn("20201001", {row["指标编码"] for row in rows})

    def test_central_skips_both_zero(self):
        """大集中核对：基础值与大集中值都为空/0 的行不输出。"""
        cur_dir, _ = self._prepare_periods()
        config_path = self.tmp / "config.xlsx"
        _make_config(config_path)
        central_path = self.tmp / "central.xlsx"
        _make_central(central_path)
        config = pc.load_period_config(config_path)
        rows = pc.compare_central(pc.load_period_directory(cur_dir, label="当期"), central_path, config)
        codes = {row["指标编码"] for row in rows}
        # 20202001 存款非零且勾选核对 → 在；未勾选核对的指标 → 不在
        self.assertIn("20202001", codes)
        self.assertNotIn("20202002", codes)

    def test_compare_periods_units_and_remarks(self):
        cur_dir, pre_dir = self._prepare_periods()
        config_path = self.tmp / "config.xlsx"
        _make_config(config_path)
        config = pc.load_period_config(config_path)
        rows = pc.compare_periods(
            pc.load_period_directory(cur_dir, label="当期"),
            pc.load_period_directory(pre_dir, label="上期"),
            config,
        )
        by_code = {row["指标编码"]: row for row in rows}
        # 元→万元换算 + 环比（除以原上期值）
        deposit = by_code["20202001"]
        self.assertAlmostEqual(deposit["当期数"], 283.1236, places=4)
        self.assertAlmostEqual(deposit["变动绝对值"], 83.1236, places=4)
        self.assertAlmostEqual(deposit["环比变动"], 41.5618, places=3)
        self.assertEqual(deposit["备注"], "增幅[30%,50%)")
        # 空备注档位（0.0）必须保留：小幅变动不产生提示
        self.assertEqual(by_code["20202002"]["备注"], "")
        self.assertEqual(deposit["机构类别"], "农商行")
        self.assertEqual(deposit["承接行"], "承接一部")
        # 不转换单位（个数）且无变化
        count = by_code["20202002"]
        self.assertEqual(count["当期数"], 12)
        self.assertEqual(count["变动绝对值"], 0)
        # 文字变动
        text = by_code["20202003"]
        self.assertEqual(text["变动绝对值"], "文字变动")
        # 单边
        self.assertEqual(by_code["20202004"]["备注"], "本期有，上期无")
        self.assertEqual(by_code["20202009"]["备注"], "本期无，上期有")

    def test_central_compare(self):
        cur_dir, _ = self._prepare_periods()
        config_path = self.tmp / "config.xlsx"
        _make_config(config_path)
        central_path = self.tmp / "central.xlsx"
        _make_central(central_path, value_yi=0.02831236)   # = 283.1236 万元，与基础数据一致
        config = pc.load_period_config(config_path)
        rows = pc.compare_central(pc.load_period_directory(cur_dir, label="当期"), central_path, config)
        deposit = next(row for row in rows if row["指标编码"] == "20202001")
        self.assertAlmostEqual(deposit["基础数据值"], 283.1236, places=4)
        self.assertAlmostEqual(deposit["大集中值"], 283.1236, places=4)
        self.assertLessEqual(deposit["差异绝对值"], pc.CENTRAL_DIFF_TOLERANCE)
        self.assertEqual(deposit["是否说明"], "")
        # 制造 0.02 万元差异（+2e-6 亿元）→ 差异超过100元
        _make_central(central_path, value_yi=0.02831436)
        rows = pc.compare_central(pc.load_period_directory(cur_dir, label="当期"), central_path, config)
        deposit = next(row for row in rows if row["指标编码"] == "20202001")
        self.assertEqual(deposit["是否说明"], "差异超过100元")
        # 未勾选核对的指标不进入结果
        self.assertNotIn("20202002", {row["指标编码"] for row in rows})

    def test_run_outputs_workbook(self):
        cur_dir, pre_dir = self._prepare_periods()
        config_path = self.tmp / "config.xlsx"
        _make_config(config_path)
        central_path = self.tmp / "central.xlsx"
        _make_central(central_path)
        out_dir = self.tmp / "out"
        output = pc.run_period_compare(
            current_dir=cur_dir,
            previous_dir=pre_dir,
            central_path=central_path,
            output_dir=out_dir,
            config_path=config_path,
        )
        self.assertTrue(output.is_file())
        book = load_workbook(output, read_only=True)
        try:
            self.assertIn("两期对比", book.sheetnames)
            self.assertIn("大集中对比", book.sheetnames)
            headers = [cell.value for cell in next(book["两期对比"].iter_rows(max_row=1))]
            self.assertEqual(headers, pc.PERIOD_SHEET_HEADERS)
        finally:
            book.close()

    def test_empty_directory_fails_clearly(self):
        empty = self.tmp / "empty"
        empty.mkdir()
        with self.assertRaises(pc.PeriodCompareError):
            pc.load_period_directory(empty, label="当期")

    def _period_pair(self, cur_value, pre_value, cur_date="2026-06-30", pre_date="2026-03-31", code=20203003):
        seq = str(self._pair_seq) + "_"
        self._pair_seq += 1
        d1 = self.tmp / f"cur_{seq}{code}_{cur_value}_{cur_date}"
        d2 = self.tmp / f"pre_{seq}{code}_{pre_value}_{pre_date}"
        d1.mkdir()
        d2.mkdir()
        for d, value, date in ((d1, cur_value, cur_date), (d2, pre_value, pre_date)):
            book = Workbook()
            sheet = book.active
            sheet.title = "20203"
            sheet.append(["t", "", ""])
            sheet.append([None, None, None])
            sheet.append(["指标编号", "指标名称", "本期情况"])
            sheet.append([20201001, "金融机构名称", "甲银行"])
            sheet.append([20201002, "金融机构代码", "91440000AA"])
            sheet.append([code, "利息收入", value])
            book.save(d / f"91440000AA#{date}#01#20203#甲银行.xlsx")
        return pc.load_period_directory(d1, label="当期"), pc.load_period_directory(d2, label="上期")

    def test_special_rule_hits_on_decrease_and_skips_cross_year(self):
        _make_config(self.tmp / "config.xlsx")
        config = pc.load_period_config(self.tmp / "config.xlsx")
        config.specials.append(pc.SpecialRule(code="20203003", name="利息收入", remark="当年累计指标比上期不应减少。"))
        # 同年递减 → 命中
        cur, pre = self._period_pair(100.0, 500.0)
        rows = pc.compare_periods(cur, pre, config)
        pc.apply_special_rules(rows, cur, pre, config)
        row = next(r for r in rows if r["指标编码"] == "20203003")
        self.assertEqual(row["是否说明"], "当年累计指标比上期不应减少。")
        # 跨年递减 → 跳过并写计算过程
        cur2, pre2 = self._period_pair(100.0, 500.0, cur_date="2026-06-30", pre_date="2025-12-31")
        rows2 = pc.compare_periods(cur2, pre2, config)
        pc.apply_special_rules(rows2, cur2, pre2, config)
        row2 = next(r for r in rows2 if r["指标编码"] == "20203003")
        self.assertEqual(row2["是否说明"], "")
        self.assertIn("跨年累计不比较", row2["计算过程"])
        # 递增不命中
        cur3, pre3 = self._period_pair(900.0, 500.0)
        rows3 = pc.compare_periods(cur3, pre3, config)
        pc.apply_special_rules(rows3, cur3, pre3, config)
        row3 = next(r for r in rows3 if r["指标编码"] == "20203003")
        self.assertEqual(row3["是否说明"], "")

    def test_complex_rule_expression(self):
        _make_config(self.tmp / "config.xlsx")
        config = pc.load_period_config(self.tmp / "config.xlsx")
        config.complex_rules.append(pc.ComplexRule(
            form="20203", desc="递减且贷大于存",
            rule="AND([,,,20203003,] < {,,,20203003,} , [,,,20202002,] > [,,,20202001,])",
        ))
        config.complex_rules.append(pc.ComplexRule(
            form="20203", desc="取反规则跳过", rule="[,,,20203003,] > 0", invert=True,
        ))
        cur, pre = self._period_pair(100.0, 500.0)
        rows = pc.compare_periods(cur, pre, config)
        pc.apply_complex_rules(rows, cur, pre, config, on_step=lambda _s: None)
        # 20202001/20202002 不在该机构数据里 → 按 0 代入：0 > 0 为 False → AND 不命中
        row = next(r for r in rows if r["指标编码"] == "20203003")
        self.assertEqual(row["是否说明"], "")

    def test_complex_rule_hits_with_full_data(self):
        _make_config(self.tmp / "config.xlsx")
        config = pc.load_period_config(self.tmp / "config.xlsx")
        config.complex_rules.append(pc.ComplexRule(
            form="20203", desc="拨备覆盖率低于100%",
            rule="[,,,20202017,] < ([,,,20202012,]+[,,,20202013,]+[,,,20202014,])",
        ))
        d1 = self.tmp / "cur_full"
        d2 = self.tmp / "pre_full"
        d1.mkdir()
        d2.mkdir()
        for d, factor in ((d1, 1.0), (d2, 1.0)):
            book = Workbook()
            sheet = book.active
            sheet.title = "20202"
            sheet.append(["t", "", ""])
            sheet.append([None, None, None])
            sheet.append(["指标编号", "指标名称", "本期情况"])
            sheet.append([20201001, "金融机构名称", "甲银行"])
            sheet.append([20201002, "金融机构代码", "91440000AA"])
            sheet.append([20202002, "各项贷款", 10000 * factor])
            sheet.append([20202012, "次级类贷款", 300 * factor])
            sheet.append([20202013, "可疑类贷款", 200 * factor])
            sheet.append([20202014, "损失类贷款", 100 * factor])
            sheet.append([20202017, "贷款减值准备", 100 * factor])
            book.save(d / f"91440000AA#{'2026-06-30' if d is d1 else '2026-03-31'}#01#20202#甲银行.xlsx")
        cur = pc.load_period_directory(d1, label="当期")
        pre = pc.load_period_directory(d2, label="上期")
        rows = pc.compare_periods(cur, pre, config)
        pc.apply_complex_rules(rows, cur, pre, config, on_step=lambda _s: None)
        # 拨备 100 万 < 不良 600 万 → 命中；涉及 5 个指标行都写说明
        hit_rows = [r for r in rows if r["是否说明"] == "拨备覆盖率低于100%"]
        self.assertEqual({r["指标编码"] for r in hit_rows}, {"20202017", "20202012", "20202013", "20202014"})
        # 值经元→万元换算：拨备 100 元 → 0.01 万元
        self.assertIn("0.01", hit_rows[0]["计算过程"])

    def test_complex_rule_balance_float_tolerance(self):
        """资产负债平衡表两侧十进制相等时，浮点误差不得判“不平衡”。"""
        _make_config(self.tmp / "config.xlsx")
        config = pc.load_period_config(self.tmp / "config.xlsx")
        config.complex_rules.append(pc.ComplexRule(
            form="20202", desc="资产负债表不平衡",
            rule="[,,,20202003,] <> [,,,20202004,] + [,,,20202005,]",
        ))
        d1 = self.tmp / "cur_bal"
        d2 = self.tmp / "pre_bal"
        d1.mkdir()
        d2.mkdir()
        for d, date in ((d1, "2026-06-30"), (d2, "2026-03-31")):
            book = Workbook()
            sheet = book.active
            sheet.title = "20202"
            sheet.append(["t", "", ""])
            sheet.append([None, None, None])
            sheet.append(["指标编号", "指标名称", "本期情况"])
            sheet.append([20201001, "金融机构名称", "甲银行"])
            sheet.append([20201002, "金融机构代码", "91440000AA"])
            sheet.append([20202003, "资产总计", 54579487200000.0])   # 元 → 5,457,948.72 万
            sheet.append([20202004, "负债总计", 47513967300000.0])   # 4,751,396.73 万
            sheet.append([20202005, "所有者权益", 7065519900000.0])  # 706,551.99 万
            book.save(d / f"91440000AA#{date}#01#20202#甲银行.xlsx")
        cur = pc.load_period_directory(d1, label="当期")
        pre = pc.load_period_directory(d2, label="上期")
        rows = pc.compare_periods(cur, pre, config)
        pc.apply_complex_rules(rows, cur, pre, config, on_step=lambda _s: None)
        row = next(r for r in rows if r["指标编码"] == "20202003")
        self.assertEqual(row["是否说明"], "", "两侧十进制相同时不得报不平衡")


if __name__ == "__main__":
    unittest.main()
