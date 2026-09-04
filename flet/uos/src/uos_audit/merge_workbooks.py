"""Native implementations of the two Windows workbook-combination modules."""
from __future__ import annotations
import re
from copy import copy
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from .discovery import source_workbooks

INVALID = re.compile(r"[\[\]:*?/\\]")
EXTERNAL = re.compile(r"'(?:[^']*\[[^]]+\])([^']+)'!")

def _sheet_name(value, used):
    base = INVALID.sub("_", value)[:31] or "工作表"; out=base; n=2
    while out.casefold() in used:
        tail="_"+str(n); out=base[:31-len(tail)]+tail; n+=1
    used.add(out.casefold()); return out

def _copy_sheet(source, target, name, style_cache=None):
    """Copy one sheet while registering each visible style only once.

    This is the same style-cache approach used by the Windows implementation.
    Assigning a source workbook's ``_style`` directly is unsafe; rebuilding
    font/fill/border for every cell is safe but disproportionately slow.
    """
    from openpyxl.cell.cell import Cell, MergedCell

    style_cache = style_cache if style_cache is not None else {}
    sheet=target.create_sheet(name); sheet.sheet_state=source.sheet_state; sheet.freeze_panes=source.freeze_panes
    for row in source.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            dst=sheet.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                key=(id(source.parent), tuple(cell._style))
                target_style=style_cache.get(key)
                if target_style is None:
                    # Register the public style components once in the target
                    # workbook, then safely reuse its target-local style IDs.
                    prototype=Cell(sheet, row=1, column=1)
                    prototype.font=copy(cell.font); prototype.fill=copy(cell.fill)
                    prototype.border=copy(cell.border); prototype.alignment=copy(cell.alignment)
                    prototype.number_format=cell.number_format; prototype.protection=copy(cell.protection)
                    target_style=copy(prototype._style)
                    style_cache[key]=target_style
                dst._style=copy(target_style)
            if cell.comment: dst.comment=copy(cell.comment)
            if cell.hyperlink: dst._hyperlink=copy(cell.hyperlink)
    for key, dim in source.row_dimensions.items():
        sheet.row_dimensions[key].height, sheet.row_dimensions[key].hidden=dim.height,dim.hidden
        sheet.row_dimensions[key].outline_level=dim.outline_level
    for key, dim in source.column_dimensions.items():
        sheet.column_dimensions[key].width, sheet.column_dimensions[key].hidden=dim.width,dim.hidden
        sheet.column_dimensions[key].outline_level=dim.outline_level
    for merged in source.merged_cells.ranges: sheet.merge_cells(str(merged))
    sheet.sheet_view.showGridLines=source.sheet_view.showGridLines
    sheet.sheet_properties.tabColor=source.sheet_properties.tabColor
    return sheet


def _safe_defined_name(value: str) -> str:
    """Use a worksheet name as an Excel-name suffix without invalid symbols."""
    cleaned = re.sub(r"[^0-9A-Za-z_一-鿿]", "_", value)
    return cleaned.strip("_") or "工作表"


def _copy_defined_names(source_book, target_book, sheet_map: dict[str, str]) -> list[str]:
    """Copy range names belonging to copied sheets into a combined template.

    Excel permits a workbook name and a sheet-scoped name with the same text,
    but this is ambiguous after templates are combined.  The combined template
    therefore uses one workbook-scoped name per copied sheet.  Collision names
    receive the target sheet suffix, e.g. ``表结构区域_001_存量个人贷款信息``;
    configured prefix matching keeps ``表结构区域`` modules working normally.
    Formula compatibility names (such as ``_xlfn.SUMIFS``) have no cell
    destination and are deliberately ignored.
    """
    existing = {item.name.casefold() for item in target_book.defined_names.values()}
    copied: list[str] = []
    for defined in source_book.defined_names.values():
        try:
            destinations = tuple(defined.destinations)
        # Malformed/legacy defined names (notably a broken local reference)
        # can raise IndexError inside openpyxl's destination parser.  They do
        # not belong to a copied worksheet, so skip them instead of losing an
        # otherwise usable combined template.
        except (AttributeError, IndexError, TypeError, ValueError):
            continue
        for source_sheet, address in destinations:
            target_sheet = sheet_map.get(source_sheet)
            if not target_sheet:
                continue
            base = defined.name
            suffix = "_" + _safe_defined_name(target_sheet)
            candidate = base if base.casefold().endswith(suffix.casefold()) else base + suffix
            number = 2
            while candidate.casefold() in existing:
                candidate = "{}_{}".format(base + suffix, number)
                number += 1
            target_book.defined_names.add(
                DefinedName(candidate, attr_text="'{}'!{}".format(target_sheet.replace("'", "''"), address))
            )
            existing.add(candidate.casefold())
            copied.append(candidate)
    return copied


def _suffix_existing_defined_names(book) -> list[str]:
    """Give every range name already in the base template a sheet suffix.

    A combined template previously had unsuffixed base names plus suffixed
    imported names, which made the Name Manager misleading.  Split multi-area
    names into one workbook-scoped name per sheet and keep the configured
    prefix (``表结构区域`` etc.) intact for the audit workflow.
    """
    existing = {item.name.casefold() for item in book.defined_names.values()}
    replacements: list[tuple[str, list[DefinedName]]] = []
    copied: list[str] = []
    for defined in list(book.defined_names.values()):
        try:
            destinations = tuple(defined.destinations)
        except (AttributeError, IndexError, TypeError, ValueError):
            continue
        destination_names: list[DefinedName] = []
        for sheet_name, address in destinations:
            if sheet_name not in book.sheetnames:
                continue
            suffix = "_" + _safe_defined_name(sheet_name)
            if defined.name.casefold().endswith(suffix.casefold()):
                candidate = defined.name
            else:
                candidate = defined.name + suffix
            number = 2
            while candidate.casefold() in existing and candidate.casefold() != defined.name.casefold():
                candidate = "{}_{}".format(defined.name + suffix, number)
                number += 1
            # Already normalized names are left in place rather than deleting
            # and recreating them, preserving any workbook-local metadata.
            if candidate == defined.name and len(destinations) == 1:
                destination_names = []
                break
            existing.add(candidate.casefold())
            destination_names.append(DefinedName(candidate, attr_text="'{}'!{}".format(sheet_name.replace("'", "''"), address)))
        if destination_names:
            replacements.append((defined.name, destination_names))
    for original, names in replacements:
        try:
            del book.defined_names[original]
        except KeyError:
            continue
        for name in names:
            book.defined_names.add(name)
            copied.append(name.name)
    return copied

def _period(path):
    parts=path.stem.split("_"); return parts[3] if len(parts)>3 else ""

def merge_by_first_filename_part(
    *, input_dir: Path, output_dir: Path, recursive=True, selected_files=None,
    output_name="组合联合核查表", on_step=None,
):
    """`修改_组合联合核查表`: one output per first underscore segment."""
    say=on_step or (lambda _ : None); files=source_workbooks(input_dir,recursive=recursive)
    if selected_files:
        allowed={p.resolve() for p in files}; files=[Path(p) for p in selected_files if Path(p).resolve() in allowed]
    groups=defaultdict(list)
    for path in files: groups[path.stem.split("_",1)[0] or path.stem].append(path)
    target_dir=output_dir/(str(output_name or "组合联合核查表")+"_"+datetime.now().strftime("%Y%m%d%H%M%S")); target_dir.mkdir(parents=True,exist_ok=True)
    records=[]
    for group, paths in sorted(groups.items()):
        out=target_dir/(group+"_合并_"+(_period(paths[0]) or "本期")+".xlsx"); book=Workbook(); del book[book.sheetnames[0]]; used=set(); copied=[]
        try:
            for path in paths:
                source=load_workbook(path,data_only=False,keep_links=False)
                try:
                    style_cache={}
                    for sheet in source.worksheets:
                        name=_sheet_name(sheet.title,used); _copy_sheet(sheet,book,name,style_cache); copied.append(path.name+"｜"+sheet.title)
                finally: source.close()
            if not book.sheetnames: book.create_sheet("空工作表")
            book.save(out); records.append((group,out.name,"；".join(copied),"成功")); say("已组合：{}（合并源文件 {} 个）".format(group,len(paths)))
        except Exception as exc: records.append((group,"","","失败："+str(exc)))
        finally: book.close()
    return target_dir, records

def group_key_from_plan(path: Path, plan: dict) -> str:
    mode=str(plan.get("mode") or "regex"); stem=path.stem
    if mode=="regex":
        match=re.search(str(plan.get("pattern") or ""),stem)
        if not match: raise ValueError("未匹配正则“{}”".format(plan.get("pattern") or ""))
        return str(match.groupdict().get("组合") or match.groupdict().get("group") or (match.group(1) if match.groups() else "")).strip() or "未识别组合"
    if mode=="name":
        found=[]
        for item in plan.get("groups",[]):
            name=str(item.get("name") or "").strip(); words=item.get("keywords",[])
            if isinstance(words,str): words=re.split(r"[、，,；;\n]",words)
            if name and any(str(word).strip().casefold() in stem.casefold() for word in words if str(word).strip()): found.append(name)
        if len(found)!=1: raise ValueError("未匹配名称/关键字" if not found else "同时匹配多个组合："+"、".join(found))
        return found[0]
    raise ValueError("不支持的分组方式：{}".format(mode))

def combine_by_plan(
    *, input_dir: Path, output_dir: Path, plan: dict, recursive=True,
    selected_files=None, output_name="组合工作表", on_step=None,
):
    """`修改_组合工作表`: uses regex or name/keyword grouping profile."""
    files=source_workbooks(input_dir,recursive=recursive)
    if selected_files:
        allowed={path.resolve() for path in files}
        files=[Path(path).resolve() for path in selected_files if Path(path).resolve() in allowed]
    groups=defaultdict(list)
    for path in files:
        try: groups[group_key_from_plan(path,plan)].append(path)
        except ValueError: continue
    folder=output_dir/(str(output_name or "组合工作表")+"_"+datetime.now().strftime("%Y%m%d%H%M%S")); folder.mkdir(parents=True,exist_ok=True); rows=[]; say=on_step or (lambda _:None)
    for name,paths in sorted(groups.items()):
        output=folder/(name+"_组合.xlsx"); book=Workbook(); del book[book.sheetnames[0]]; used=set(); sheets=[]
        try:
            for path in paths:
                source=load_workbook(path,data_only=False,keep_links=False)
                try:
                    style_cache={}
                    for sheet in source.worksheets:
                        if sheet.title.casefold() in used: continue
                        _copy_sheet(sheet,book,_sheet_name(sheet.title,used),style_cache); sheets.append(path.name+"｜"+sheet.title)
                finally: source.close()
            if not book.sheetnames: book.create_sheet("空工作表")
            book.save(output); rows.append((name,output.name,"；".join(sheets),"成功")); say("已组合：{}（合并源文件 {} 个，生成工作簿 1 个）".format(name,len(paths)))
        except Exception as exc: rows.append((name,"","","失败："+str(exc)))
        finally: book.close()
    return folder,rows

def create_combined_template(*, base_template: Path, source_templates: list[Path], on_step=None):
    """Low-frequency template composition: base stays untouched; duplicate sheets skip."""
    say=on_step or (lambda _ : None); base=load_workbook(base_template,data_only=False,keep_links=True)
    try:
        renamed_base_names = _suffix_existing_defined_names(base)
        if renamed_base_names:
            say("已规范基准模板 {} 个命名区域：{}".format(len(renamed_base_names), "、".join(renamed_base_names[:8])))
        used={name.casefold() for name in base.sheetnames}; copied=[]; skipped=[]
        for path in source_templates:
            source=load_workbook(path,data_only=False,keep_links=True)
            try:
                style_cache = {}
                sheet_map = {}
                for sheet in source.worksheets:
                    if sheet.title.casefold() in used: skipped.append((path.name,sheet.title)); continue
                    target_name = _sheet_name(sheet.title,used)
                    _copy_sheet(sheet,base,target_name,style_cache)
                    sheet_map[sheet.title] = target_name
                    copied.append((path.name,sheet.title))
                copied_names = _copy_defined_names(source, base, sheet_map)
                if copied_names:
                    say("已复制 {} 个命名区域：{}".format(path.name, "、".join(copied_names[:8])))
            finally: source.close()
        for sheet in base.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value,str) and cell.value.startswith("="):
                        cell.value=EXTERNAL.sub(lambda m:"'"+m.group(1).replace("'","''")+"'!",cell.value)
        folder=base_template.parent/"联合模板"; folder.mkdir(exist_ok=True); stamp=datetime.now().strftime("%Y%m%d%H%M%S")
        out=folder/("！联合模板_"+"+".join(p.stem.lstrip("！") for p in [base_template,*source_templates])[:120]+"_"+stamp+".xlsx")
        base.save(out); say("联合模板已生成：{}".format(out.name)); return out,copied,skipped
    finally: base.close()
