from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "基础数据审核工具使用说明.docx"

BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
NAVY = "17365D"
MUTED = "666666"
LIGHT_BLUE = "E8EEF5"
LIGHT_GRAY = "F4F6F9"
GOLD = "9C6500"
RED = "9B1C1C"
TABLE_WIDTH = 9360
TABLE_INDENT = 120


def set_run_font(
    run,
    *,
    name: str = "Calibri",
    east_asia: str = "微软雅黑",
    size: float | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
) -> None:
    run.font.name = name
    fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    fonts.set(qn("w:ascii"), name)
    fonts.set(qn("w:hAnsi"), name)
    fonts.set(qn("w:eastAsia"), east_asia)
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_style_font(style, size: float, color: str = "000000", bold: bool = False) -> None:
    style.font.name = "Calibri"
    style._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    style._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor.from_string(color)
    style.font.bold = bold


def configure_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    set_style_font(normal, 11)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    for name, size, color, before, after in (
        ("Heading 1", 16, BLUE, 18, 10),
        ("Heading 2", 13, BLUE, 14, 7),
        ("Heading 3", 12, DARK_BLUE, 10, 5),
    ):
        style = doc.styles[name]
        set_style_font(style, size, color, True)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.keep_with_next = True

    code = doc.styles.add_style("Code Block", 1)
    set_style_font(code, 9.5, "333333")
    code.font.name = "Consolas"
    code._element.rPr.rFonts.set(qn("w:ascii"), "Consolas")
    code._element.rPr.rFonts.set(qn("w:hAnsi"), "Consolas")
    code.paragraph_format.space_before = Pt(4)
    code.paragraph_format.space_after = Pt(7)
    code.paragraph_format.line_spacing = 1.15
    code.paragraph_format.left_indent = Inches(0.12)
    code.paragraph_format.right_indent = Inches(0.12)


def set_cell_margins(cell, *, top=80, start=120, bottom=80, end=120) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_geometry(table, widths: list[int]) -> None:
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths)))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.first_child_found_in("w:tblInd")
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(TABLE_INDENT))
    tbl_ind.set(qn("w:type"), "dxa")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)

    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths[index] / 1440)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            tc_w = cell._tc.get_or_add_tcPr().first_child_found_in("w:tcW")
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                cell._tc.get_or_add_tcPr().append(tc_w)
            tc_w.set(qn("w:w"), str(widths[index]))
            tc_w.set(qn("w:type"), "dxa")


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def add_table(
    doc: Document,
    headers: list[str],
    rows: list[list[str]],
    widths: list[int],
    *,
    keep_together: bool = True,
    body_font_size: float = 10,
):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    set_table_geometry(table, widths)
    set_repeat_table_header(table.rows[0])
    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        shade_cell(cell, LIGHT_BLUE)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(header)
        set_run_font(run, size=10, color=NAVY, bold=True)
    for row_values in rows:
        row = table.add_row()
        for index, value in enumerate(row_values):
            cell = row.cells[index]
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(value)
            set_run_font(run, size=body_font_size)
    if keep_together:
        for row in table.rows[:-1]:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.keep_with_next = True
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_paragraph_shading(paragraph, fill: str, border_color: str = BLUE) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    p_pr.append(shd)
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:space"), "6")
    left.set(qn("w:color"), border_color)
    borders.append(left)
    p_pr.append(borders)


def add_callout(doc: Document, label: str, text: str, *, caution: bool = False) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.12)
    p.paragraph_format.right_indent = Inches(0.08)
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.2
    add_paragraph_shading(p, "FFF8E8" if caution else LIGHT_GRAY, GOLD if caution else BLUE)
    label_run = p.add_run(label + "：")
    set_run_font(label_run, size=10.5, color=GOLD if caution else NAVY, bold=True)
    text_run = p.add_run(text)
    set_run_font(text_run, size=10.5)


def add_code(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="Code Block")
    add_paragraph_shading(p, "F2F2F2", "B7B7B7")
    run = p.add_run(text)
    set_run_font(run, name="Consolas", east_asia="等线", size=9.5, color="333333")


def create_numbering(doc: Document, *, bullet: bool = False) -> int:
    numbering = doc.part.numbering_part.element
    abstract_ids = [int(node.get(qn("w:abstractNumId"))) for node in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(node.get(qn("w:numId"))) for node in numbering.findall(qn("w:num"))]
    abstract_id = (max(abstract_ids) + 1) if abstract_ids else 1
    num_id = (max(num_ids) + 1) if num_ids else 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)
    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    level.append(start)
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), "bullet" if bullet else "decimal")
    level.append(num_fmt)
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), "·" if bullet else "%1.")
    level.append(lvl_text)
    suff = OxmlElement("w:suff")
    suff.set(qn("w:val"), "tab")
    level.append(suff)
    p_pr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "num")
    tab.set(qn("w:pos"), "540")
    tabs.append(tab)
    p_pr.append(tabs)
    indent = OxmlElement("w:ind")
    indent.set(qn("w:left"), "540")
    indent.set(qn("w:hanging"), "270")
    p_pr.append(indent)
    level.append(p_pr)
    abstract.append(level)
    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        numbering.insert(list(numbering).index(first_num), abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering.append(num)
    return num_id


def add_list_item(doc: Document, text: str, num_id: int, *, bold_prefix: str = "") -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    p_pr = p._p.get_or_add_pPr()
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num = OxmlElement("w:numId")
    num.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl)
    num_pr.append(num)
    p_pr.append(num_pr)
    if bold_prefix and text.startswith(bold_prefix):
        first = p.add_run(bold_prefix)
        set_run_font(first, bold=True)
        rest = p.add_run(text[len(bold_prefix) :])
        set_run_font(rest)
    else:
        run = p.add_run(text)
        set_run_font(run)


def add_heading(doc: Document, text: str, level: int, *, page_break: bool = False) -> None:
    p = doc.add_paragraph(text, style=f"Heading {level}")
    p.paragraph_format.page_break_before = page_break


def add_page_field(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("第 ")
    set_run_font(run, size=9, color=MUTED)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    paragraph._p.append(fld)
    run = paragraph.add_run(" 页")
    set_run_font(run, size=9, color=MUTED)


def build_document() -> None:
    """Build the current, concise operator guide used by the desktop workbench."""
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)
    configure_styles(doc)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = header.add_run("基础数据审核工具｜使用说明")
    set_run_font(run, size=9, color=MUTED)
    add_page_field(section.footer.paragraphs[0])

    kicker = doc.add_paragraph()
    kicker.paragraph_format.space_before = Pt(10)
    kicker.paragraph_format.space_after = Pt(2)
    run = kicker.add_run("操作手册")
    set_run_font(run, size=10, color=GOLD, bold=True)

    title = doc.add_paragraph()
    title.paragraph_format.space_after = Pt(6)
    run = title.add_run("基础数据审核工具使用说明")
    set_run_font(run, size=28, color=NAVY, bold=True)
    subtitle = doc.add_paragraph()
    subtitle.paragraph_format.space_after = Pt(16)
    run = subtitle.add_run("批量审核金融机构报送数据，保留公式复核与历史说明")
    set_run_font(run, size=12.5, color=MUTED)

    add_callout(
        doc,
        "使用原则",
        "原始报送文件、模板文件和外部文件均不修改。程序只在输出目录生成审核副本、审核结果和汇总文件。",
    )

    add_heading(doc, "1. 日常使用", 1)
    steps = create_numbering(doc)
    for text in (
        "双击“基础数据审核工具.exe”打开工作台。",
        "选择源数据目录。首次选择或点击“自动识别模板”时，程序会推荐匹配模板并列出待处理文件。",
        "确认模板文件目录、模板文件和外部文件；日常只需调整源数据目录，其他路径会记住上次选择。",
        "确认输出目录。默认跟随源数据目录并使用“执行结果”子目录；点亮书钉后可固定输出目录。",
        "在待处理文件列表中取消本次不需要处理的文件。不会删除源文件。",
        "点击需要执行的流程，例如“汇总核查表校验”或“汇总校验结果说明”。",
        "任务结束后，查看页面运行记录；如已在“设置”中开启“输出流程运行日志”，再核对运行日志中的失败文件。随后查看问题汇总和审核副本。",
    ):
        add_list_item(doc, text, steps)

    add_heading(doc, "1.1 主界面的常用流程", 2)
    add_table(
        doc,
        ["流程", "适用场景", "输出"],
        [
            ["汇总核查表校验", "审核同一类报表的多家机构报送文件。", "机构审核副本、本期审核结果；可选流程运行日志。"],
            ["汇总校验结果说明", "汇总机构在系统中下载的校验结果及报送说明。", "汇总结果工作簿；可选流程运行日志。"],
            ["组合联合核查表", "将同一机构分散在多个报送工作簿中的工作表组合为一本，供制作联合模板或跨表核查。", "每家机构一本组合工作簿。"],
        ],
        [2100, 3450, 3810],
        keep_together=False,
    )

    add_heading(doc, "1.2 递归子文件夹", 2)
    bullet = create_numbering(doc, bullet=True)
    add_list_item(doc, "“汇总核查表校验”通常不勾选。源数据目录一般就是本期同类报表所在目录。", bullet)
    add_list_item(doc, "“汇总校验结果说明”在机构文件分散于子文件夹时通常勾选。程序会向下查找 xlsx 文件。", bullet)
    add_list_item(doc, "程序固定跳过“执行结果”目录；其他不需处理的目录只要名称包含“_skip”也会跳过，例如“历史材料_skip”。", bullet)

    add_heading(doc, "2. 审核后看什么", 1)
    add_table(
        doc,
        ["文件或目录", "用途"],
        [
            ["机构审核副本", "文件名为“原文件名_审核版.xlsx”。包含复制后的公式；从结果表的“定位单元格”可直接跳转复核。"],
            ["基础数据审核结果_*.xlsx", "包含“本期审核结果”，定位单元格可跳转至对应审核副本复核。"],
            ["流程名_运行日志_时间戳.xlsx", "仅在“设置 - 输出流程运行日志”开启时生成。一个功能一个工作表：检查类列出各命名区域所在工作表与覆盖区域，表结构比对列出各文件匹配结果，外部文件添加 / 公式校验复制列出复制记录，另有模板体检工作表。"],
            ["机构审核副本_时间戳", "“汇总核查表校验”默认保留。每家机构文件名为“原文件名_审核版.xlsx”，包含已复制的公式和审核导航；本期审核结果中的定位单元格可直接跳转至对应位置。"],
            ["序号_功能名_时间戳（或 输出文件名_时间戳）", "修改类、汇总类步骤仅在“是否输出结果=是”时保留实际文件。检查类、核对类填“是”时，其检查明细保留在流程运行日志中；填写“输出文件名”后，以该值替换“序号_功能名”作为前缀。"],
        ],
        [3150, 6210],
        keep_together=False,
    )
    add_callout(
        doc,
        "先看运行记录",
        "本期审核结果为 0 不一定代表没有问题。应先确认页面运行记录没有失败文件；需要逐项追溯时，再临时开启“输出流程运行日志”后重跑。",
        caution=True,
    )

    add_heading(doc, "3. 公式结果的写法", 1)
    p = doc.add_paragraph("校验公式保存在模板中。条件不触发时应返回空字符串；触发时使用英文竖线“|”分隔内容。")
    set_run_font(p.runs[0])
    add_table(
        doc,
        ["推荐格式", "含义"],
        [
            ["错误类型|校验指标|描述", "最简格式。适用于只需要说明问题的情况。"],
            ["错误类型|校验指标|描述|当前值|对比值|差值", "需要在结果表中明确展示取值时使用；三个值均可为文字或数字。"],
        ],
        [4300, 5060],
    )
    add_code(doc, '=IF(A11>0,"错误|字段其他|不应有数，请核实|"&A11&"||"&A11,"")')
    add_callout(
        doc,
        "提取标记",
        "通常使用“错、错误、硬性、软性”作为第一段。普通说明文字不要放在这些标记开头，否则会被当作问题提取。",
    )

    add_heading(doc, "4. 怎样制作和维护模板", 1)
    p = doc.add_paragraph(
        "模板必须与机构报送文件使用相同的工作表名称和单元格位置。建议从一份确认无误的标准报送空表复制出模板，"
        "在模板中加入校验公式，再通过 Excel 原生的“名称管理器”标出程序应处理的位置。模板文件名应以“！”或“!”开头。"
    )
    set_run_font(p.runs[0])
    add_callout(
        doc,
        "区域划分方式已经改变",
        "现在只通过 Excel 命名区域划分功能范围，不再通过单元格批注确定区域。批注可以保留规则说明，但不会决定公式复制、结构比对或汇总范围。",
        caution=True,
    )

    add_heading(doc, "4.1 在 Excel 中建立命名区域", 2)
    named_steps = create_numbering(doc)
    for text in (
        "打开模板，先切换到需要划分区域的业务工作表。",
        "选中连续区域；若区域不连续，可按住 Ctrl 依次选中多块区域或多个单元格。",
        "打开 Excel“公式”选项卡，点击“定义名称”或“名称管理器 - 新建”。",
        "在“名称”中填写本说明规定的名称；“范围”选择“工作簿”，不要只限定到当前工作表。",
        "确认“引用位置”指向正确的工作表和绝对地址，例如“=贷款明细!$A$1:$F$3”。点击确定。",
        "重新打开“名称管理器”，逐项检查名称、引用工作表和区域边界；保存并关闭模板。",
    ):
        add_list_item(doc, text, named_steps)
    add_callout(
        doc,
        "同类区域有多块时",
        "一个名称可直接引用不连续区域。也可以建立“名称_后缀”，例如“表结构区域_01”“表结构区域_02”或“校验区域_贷款”“校验区域_担保”。程序识别基础名称以及紧跟下划线的变体；不要写成“表结构区域01”或自行改用空格、短横线。",
    )

    add_heading(doc, "4.1.1 联合模板的命名区域", 2)
    p = doc.add_paragraph(
        "单个模板内可继续使用“表结构区域_001”“校验区域_001”等简短名称。制作联合模板时，程序会自动把每个区域转换为工作簿级唯一名称，"
        "在原名称后追加工作表名，例如“表结构区域_001_存量个人贷款信息”。"
    )
    set_run_font(p.runs[0])
    add_callout(
        doc,
        "不需要手工改旧模板",
        "自动加后缀只发生在新生成的联合模板中，避免单位贷款和个人贷款等多个模板合并后出现同名区域冲突。功能配置仍只填写“表结构区域”“校验区域”等基础名称，程序按“基础名称_后缀”识别。已有联合模板不会自动修复，应重新制作。",
    )

    add_heading(doc, "4.1.2 制作联合模板（跨表核查时使用）", 2)
    p = doc.add_paragraph(
        "当一套核查公式需要同时引用单位贷款、个人贷款或其他多类报表时，先制作联合模板，再使用联合后的报送工作簿执行普通“汇总核查表校验”。该操作只生成新文件，不会修改任何原始模板。"
    )
    set_run_font(p.runs[0])
    merge_steps = create_numbering(doc)
    for text in (
        "在主界面“组合联合核查表”按钮上使用右键，选择“制作联合模板”。",
        "第一步选择“底稿模板”。底稿应优先选已包含集中系统数据、参照表等外部依赖工作表的模板；它的同名工作表和样式优先保留。",
        "第二步批量选择需要并入的其他模板，例如单位贷款模板与个人贷款模板。",
        "程序复制底稿后，把其他模板中不重名的工作表复制进去；同名工作表不覆盖底稿，直接跳过，以避免误替换外部依赖表或公式表。",
        "程序把每张业务表的命名区域改为工作簿级唯一名称，并在名称后附加工作表名；然后生成“联合模板”和“检查报告”。",
        "打开检查报告，确认“工作表处理清单”中需要的工作表均为“已复制”或“保留为底稿”，确认“命名区域处理清单”没有失败记录。",
        "在联合模板中补充跨表校验公式，并在名称管理器抽查校验区域、表结构区域的引用位置；保存后把联合模板放入模板文件目录，通过“自动识别模板”或手动选择用于联合报送工作簿。",
    ):
        add_list_item(doc, text, merge_steps)
    add_callout(
        doc,
        "两个容易混淆的动作",
        "“组合联合核查表”是把同一机构的多个报送文件组合为一本，用于生成联合源数据；“制作联合模板”是把多个模板组合为一个模板，用于写跨表公式。两者通常配套使用，但也可以分别使用。",
        caution=True,
    )

    add_heading(doc, "4.2 每个区域的作用和划分位置", 2)
    add_table(
        doc,
        ["命名区域", "应该框选哪里", "作用与注意事项"],
        [
            [
                "校验区域",
                "框选模板中已经写入校验公式的单元格，以及确实需要随公式一起复制的格式范围。",
                "程序把区域内的公式和格式复制到审核副本的相同坐标，并从计算结果中提取问题。不要框入机构原始数据输入区，否则可能覆盖报送值。",
            ],
            [
                "表结构区域",
                "框选固定表名、固定列标题、固定行标题等用于确认表样的非空单元格；可以只选零散关键单元格。",
                "审核前按相同工作表、相同坐标逐格严格比较。不要包含机构名称、机构代码、统一社会信用代码、数据日期、金额、数量等每期或每家会变化的内容。",
            ],
            [
                "条件格式区域",
                "框选源文件中可能触发条件格式的单元格范围。该区域通常在原始报送文件中已存在，无需从模板复制。",
                "用于“条件格式结果提取”。程序读取实际显示填充色发生变化的单元格，不限定为红色；以单元格批注作为描述，并借助表结构区域生成行、列指标。",
            ],
            [
                "表头区域",
                "框选汇总数据对应的那一行列标题，只选表头行；列范围应与同工作表的汇总区域有交集。",
                "为“任意行汇总区域”或“固定行汇总区域”提供输出列名。每个需要汇总的工作表都必须有对应表头区域。",
            ],
            [
                "任意行汇总区域",
                "在数据首行框选要汇总的列，通常从表头下一行开始；区域的左右边界决定输出哪些列。",
                "程序从该起始行一直读取到源工作表实际使用区域的最后一行，适合机构填报行数不固定的明细表。模板不必预先框到大量空白行。",
            ],
            [
                "固定行汇总区域",
                "完整框选需要汇总的固定数据行，并让左右边界覆盖需要输出的列。",
                "程序只读取框选的起止行，适合指标位置固定、行数固定的报表；区域之外的行不会汇总。",
            ],
            [
                "单元格区域.字段名",
                "框选一个需要附加到每条汇总记录的固定信息单元格，例如数据日期或机构名称所在格。",
                "字段名会成为输出列名，例如“单元格区域.数据日期”。同工作表汇总使用；每个名称应指向一个单元格。",
            ],
            [
                "全局单元格区域.字段名",
                "框选一个需要附加到所有汇总工作表记录的公共信息单元格。",
                "适合工作簿级机构信息或数据期。字段会加入所有汇总结果，即使数据区位于其他工作表。",
            ],
        ],
        [1850, 3500, 4010],
        keep_together=False,
    )

    add_heading(doc, "4.3 表结构区域怎样选择", 2)
    bullet = create_numbering(doc, bullet=True)
    add_list_item(doc, "应选：固定工作表标题、固定栏目名称、固定指标名称，以及能唯一识别表样的关键文字或数字代码。", bullet)
    add_list_item(doc, "不应选：机构名称、机构编码、统一社会信用代码、数据日期、统计期、金额、数量、备注和其他报送内容。", bullet)
    add_list_item(doc, "模板空白单元格会跳过；模板非空单元格全部参与比较。不要用大矩形把大量无关单元格一并框入。", bullet)
    add_list_item(doc, "文本按原文严格比较，前后空格、符号和日期不会自动清洗；数字与数字按数值比较，但数字 1001 与文本“1001”视为不同。", bullet)
    add_list_item(doc, "合并单元格只需包含其左上角实际存值单元格；框选整个合并区域也可以，其余空白格会跳过。", bullet)
    add_list_item(doc, "同一工作簿有多张业务表时，应分别在各工作表建立带后缀的区域，例如“表结构区域_贷款”“表结构区域_担保”。", bullet)
    add_callout(
        doc,
        "推荐做法",
        "优先选择少量、稳定、分散且能代表表样的固定表头。结构检查默认要求至少 90% 的非空模板单元格匹配；选入动态内容会导致正确报送文件被判为结构不匹配。",
    )

    add_heading(doc, "4.4 校验区域怎样选择", 2)
    bullet = create_numbering(doc, bullet=True)
    add_list_item(doc, "先把校验公式写在模板的目标位置，确认公式引用的是报送表中的正确单元格。", bullet)
    add_list_item(doc, "框选公式单元格及其需要复制的格式。多个相隔较远的公式应使用不连续区域或带后缀的多个校验区域。", bullet)
    add_list_item(doc, "不要为了方便而框选整张工作表、整列或包含机构录入数据的大区域。程序会把所选范围复制到审核副本的相同位置。", bullet)
    add_list_item(doc, "公式未触发时返回空字符串；触发时按第 3 节格式返回问题内容。", bullet)
    add_list_item(doc, "如果公式依赖外部工作表，先确认公式中的工作表名称与外部文件实际名称一致，并在执行流程中让“外部文件添加”排在“公式校验复制”之前。", bullet)

    add_heading(doc, "4.5 汇总区域怎样配套", 2)
    p = doc.add_paragraph(
        "汇总功能以工作表为单位配对区域。同一工作表中的“表头区域”和“任意行汇总区域”或“固定行汇总区域”必须覆盖共同列。"
        "例如表头在 A5:C5，变动行数的数据从第 6 行开始，可把 A5:C5 命名为“表头区域”，把 A6:C6 命名为“任意行汇总区域”；"
        "若只汇总第 6 至第 10 行，则把 A6:C10 命名为“固定行汇总区域”。"
    )
    set_run_font(p.runs[0])
    add_callout(
        doc,
        "划分完成后的检查",
        "在名称管理器中逐项点击名称，确认 Excel 高亮了预期工作表和单元格。然后运行“汇总核查表校验”，查看运行日志中的“模板体检”和“表结构比对”工作表；模板体检或文件匹配不通过时，不要修改机构报送文件，应返回模板核对名称、坐标和固定表头值。",
        caution=True,
    )

    add_heading(doc, "4.6 设置中心：模块、功能和流程", 2)
    p = doc.add_paragraph(
        "日常审核不需要修改设置。需要组合新的流程时，在工作台右上角打开“设置”，进入“进阶使用”后维护。"
        "模块和流程都可直接调整；恢复默认设置即可回到程序初始配置。"
    )
    set_run_font(p.runs[0])
    add_table(
        doc,
        ["位置", "维护内容"],
        [
            ["模块说明", "模块是程序已经实现的固定能力，例如命名区域存在性检查、表结构比对、外部文件添加、公式校验复制、公式结果提取、任意行汇总、固定行汇总、组合工作表和汇总表合并。"],
            ["功能配置", "功能是给模块起业务名称并指定作用区域。可修改功能名、执行模块、命名区域名、输入、输出和备注；同一执行模块可配置多个功能。命名区域名支持多个名称，以逗号、顿号或分号分隔。"],
            ["流程编排", "流程把功能按顺序组合。每一步设置启用、失败后处理（停止/跳过）、是否输出结果、输出文件名、输入和输出结果集；检查类、核对类填“是”时会把明细保留在流程运行日志中。勾选“显示在主界面”后，流程名成为主界面入口。"],
            ["组合工作表功能参数", "选中“组合工作表”功能时，右侧显示分组方案。可按名称/关键字或正则表达式创建多个方案，选择当前使用方案后，通过自定义流程运行该功能。"],
        ],
        [2400, 6960],
    )
    add_callout(
        doc,
        "恢复默认设置",
        "设置中心的“恢复默认设置”只清空自定义功能、自定义流程及主界面显示设置，并恢复内置默认状态；不会修改根目录的“历史审核说明.xlsx”。",
    )

    add_heading(doc, "5. 外部文件", 1)
    p = doc.add_paragraph("外部文件用于模板中引用辅助数据的公式。执行“外部文件添加”模块时，程序会在公式复制前把需要的工作表复制到审核副本。")
    set_run_font(p.runs[0])
    bullet = create_numbering(doc, bullet=True)
    add_list_item(doc, "程序优先从模板公式识别引用的工作表；识别不到才复制外部文件的全部工作表。", bullet)
    add_list_item(doc, "外部工作表名称必须与公式引用名称一致；若审核副本已有同名工作表，程序会停止并提示，绝不覆盖。", bullet)
    add_list_item(doc, "外部表复制在审核副本内完成，原始报送文件和外部文件保持不变。", bullet)

    add_heading(doc, "6. 历史审核说明", 1)
    p = doc.add_paragraph("“本期审核结果”和程序根目录“历史审核说明.xlsx”中的“历史审核结果”使用同一组表头：")
    set_run_font(p.runs[0])
    add_code(doc, "工作簿名｜工作表名｜定位单元格｜错误类型｜校验指标｜描述｜当前值｜对比值｜差值｜规则编号｜历史校验说明｜审核意见")
    p = doc.add_paragraph("审核人员可将需要长期保留的问题整行复制到“历史审核结果”，填写“历史校验说明”和“审核意见”。下次同规则编号出现时，程序自动带出这两项人工内容。历史文件由用户直接用 Excel 打开维护，不在设置中心显示。")
    set_run_font(p.runs[0])
    add_callout(
        doc,
        "多条历史说明",
        "规则编号按“工作簿｜工作表｜公式单元格｜校验指标”生成，不附加序号。同一规则需要保留多条解释或意见时，直接在“历史审核结果”增加多行；程序会把这些内容合并带入本期审核结果。",
    )

    add_heading(doc, "7. 与旧 VBA 工具的对比", 1)
    add_table(
        doc,
        ["方面", "当前工具", "旧 VBA 方式"],
        [
            ["模板维护", "使用 Excel 原生命名区域定位，支持不连续区域；联合模板自动将重名区域转为工作簿级唯一名称。", "主要依赖单元格批注的行、列区域组合；不连续区域和多模板合并维护成本较高。"],
            ["审核复核", "原始报送文件只读，所有公式和外部表写入审核副本；结果中的定位单元格可直接跳回审核副本。", "可复制校验区域并提取问题，但过程文件、定位和结果留存规则更多依赖人工约定。"],
            ["流程调整", "模块、功能、流程分层配置；可调整步骤顺序、启用状态、输入、输出和失败处理，不需要改 Python。", "通常需要修改 VBA 代码、面板参数或模板宏。"],
            ["批量与追溯", "可递归扫描、自动推荐模板、筛除不匹配文件、按需输出运行日志，并保留历史审核说明。", "批处理和汇总逻辑集中在宏内，跨人员交接时更依赖对 VBA 代码的熟悉程度。"],
            ["跨表准备", "可组合联合核查表、制作联合模板，并保留外部文件先添加后复制公式的顺序。", "往往需要手工整理多个工作簿或为跨表关系额外维护宏逻辑。"],
        ],
        [1600, 3970, 3980],
        keep_together=False,
    )
    add_callout(
        doc,
        "需要如实保留的限制",
        "当前 Windows 发行版仍依赖 Microsoft Excel 或 Windows 版 WPS 计算公式，因此公式兼容性与原有 Excel/VBA 流程较接近；统信 UOS、麒麟等 Linux 环境不能直接运行本 EXE，后续需要单独采用 LibreOffice 计算引擎并用真实模板验证。",
        caution=True,
    )

    add_heading(doc, "8. 常见问题", 1)
    add_table(
        doc,
        ["现象", "处理方法"],
        [
            ["未找到可信度足够的模板", "确认源数据目录没有混放多类报表；必要时手动选择模板。"],
            ["公式出现 #REF!", "先确认外部文件已选择，且需要的工作表已在“外部文件添加”步骤中复制到审核副本。"],
            ["文件被占用", "关闭 Excel 中打开的模板、源文件、历史审核说明和上次输出结果后重试。"],
            ["某些文件没有结果", "查看“运行日志”（表结构比对、公式校验复制等工作表），确认是否因结构不匹配或工作表缺失而跳过。"],
            ["需要更改一个流程", "在“设置 → 进阶使用 → 流程编排”中建立自定义流程；无需修改 Python 程序。"],
        ],
        [2500, 6860],
        keep_together=False,
    )

    doc.core_properties.title = "基础数据审核工具使用说明"
    doc.core_properties.subject = "金融机构基础数据报表审核操作手册"
    doc.core_properties.author = "基础数据审核工作组"
    doc.core_properties.keywords = "基础数据审核, Excel, 模板校验, 历史审核结果"
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_document()
