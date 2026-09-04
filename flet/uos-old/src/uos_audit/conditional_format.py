"""LibreOffice UNO conditional-format extraction for the native edition."""
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import subprocess
import time
from contextlib import contextmanager
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from .models import AuditRule, Issue, CopyRange
from .template import TemplateError
from .libreoffice import find_calc_engine

@contextmanager
def uno_service(command_prefix: tuple[str, ...] = ("soffice",)):
    """Start an isolated headless UNO listener for one extraction batch."""
    process = subprocess.Popen([
        *command_prefix, "--headless", "--nologo", "--nodefault", "--nolockcheck",
        "--nofirststartwizard", "--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ServiceManager",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        # LibreOffice needs a short initialisation period; connection itself is
        # retried by the caller's first document request.
        time.sleep(1.2)
        yield
    finally:
        process.terminate()
        try: process.wait(timeout=8)
        except subprocess.TimeoutExpired: process.kill()

def _connect(path: Path):
    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError as exc:
        raise RuntimeError("统信条件格式提取需要 python3-uno。请安装 python3-uno，并使用系统 Python 启动程序。") from exc
    local=uno.getComponentContext(); resolver=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
    deadline = time.monotonic() + 8
    while True:
        try:
            context=resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
            desktop=context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context)
            break
        except Exception as exc:
            if time.monotonic() >= deadline:
                raise RuntimeError("LibreOffice UNO 服务未启动；条件格式提取需要可用的 python3-uno。") from exc
            time.sleep(0.2)
    hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
    # 强制 Calc 在本次 UNO 会话中更新公式和条件格式，不能只依赖 xlsx
    # 文件内可能陈旧的缓存值。
    update_mode = PropertyValue(); update_mode.Name = "UpdateDocMode"; update_mode.Value = 3
    document = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(str(path.resolve())), "_blank", 0, (hidden, update_mode)
    )
    if document is None:
        raise RuntimeError("LibreOffice UNO 无法打开审核副本：{}".format(path))
    try:
        document.calculateAll()
        if hasattr(document, "refresh"):
            document.refresh()
        time.sleep(0.25)
    except Exception as exc:
        raise RuntimeError("LibreOffice UNO 重算条件格式失败：{}".format(exc)) from exc
    return document

def extract_conditional_format_issues(*, workbook_path: Path, ranges: list[CopyRange], structure_ranges: list[CopyRange], period: str, batch_id: str, source_file: Path) -> list[Issue]:
    """Return cells with an active non-default calculated background colour.

    Named ``条件格式区域`` limits scanning; this intentionally avoids scanning a
    full sheet.  Row/column labels from ``表结构区域`` are concatenated as the
    check indicator where they exist.
    """
    document=_connect(workbook_path); static=load_workbook(workbook_path,read_only=True,data_only=True,keep_links=False)
    try:
        stamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"); result=[]; seen=set()
        for area in ranges:
            if area.sheet_name not in static.sheetnames: continue
            sheet=document.Sheets.getByName(area.sheet_name); baseline=static[area.sheet_name]
            min_col,min_row,max_col,max_row=range_boundaries(area.address)
            for row in range(min_row,max_row+1):
                for col in range(min_col,max_col+1):
                    if (area.sheet_name, row, col) in seen:
                        continue
                    seen.add((area.sheet_name, row, col))
                    cell=sheet.getCellByPosition(col-1,row-1)
                    color=int(cell.CellBackColor)
                    # White/default (16777215) and transparent/no fill (0) do
                    # not constitute a visible condition result.  A normal
                    # static fill is skipped by comparing OOXML base fill.
                    base=baseline.cell(row,col).fill.fgColor.rgb or ""
                    if color in {0,16777215} or (base and base[-6:].upper()=="{:06X}".format(color)): continue
                    address="{}{}".format(get_column_letter(col),row); labels=[]
                    for structure in structure_ranges:
                        if structure.sheet_name!=area.sheet_name: continue
                        a,b,c,d=range_boundaries(structure.address)
                        if a<=col<=c:
                            for r in range(min(row,d),b-1,-1):
                                value=baseline.cell(r,col).value
                                if value not in (None,""): labels.append(str(value)); break
                        if b<=row<=d:
                            for cc in range(min(col,c),a-1,-1):
                                value=baseline.cell(row,cc).value
                                if value not in (None,""): labels.append(str(value)); break
                    indicator="_".join(dict.fromkeys(labels)); comment=baseline.cell(row,col).comment
                    detail=comment.text.strip() if comment else "条件格式填充已触发，请核实"
                    rule=AuditRule("条件格式填充",True,"",area.sheet_name,address,address,"总行条件格式触发",detail)
                    result.append(Issue("｜".join((source_file.stem,area.sheet_name,address,indicator or "校验指标未识别")),period,batch_id,stamp,True,"","","",1,"","","",area.sheet_name,"条件格式填充","总行条件格式触发",address,address,baseline.cell(row,col).value,"",detail,str(source_file),str(workbook_path),check_field=indicator,detail=detail,display_fill_color=color))
        return result
    finally:
        static.close(); document.close(True)


def extract_conditional_format_issues_with_libreoffice(
    *, workbook_path: Path, ranges: list[CopyRange], structure_ranges: list[CopyRange],
    period: str, batch_id: str, source_file: Path,
) -> list[Issue]:
    """Read actual Calc-rendered conditional-format colours when UNO is available."""
    engine = find_calc_engine()
    if engine is None:
        raise RuntimeError("未找到 LibreOffice Calc")
    try:
        import uno  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("已找到 LibreOffice，但当前 Python 未安装 python3-uno") from exc
    # 玲珑版同样尝试通过宿主 python3-uno 连接本地 UNO socket；过去直接
    # 拒绝它，导致已能计算公式的玲珑 LibreOffice 永远无法提取条件格式。
    with uno_service(engine.command_prefix):
        return extract_conditional_format_issues(
            workbook_path=workbook_path, ranges=ranges,
            structure_ranges=structure_ranges, period=period,
            batch_id=batch_id, source_file=source_file,
        )


@contextmanager
def libreoffice_conditional_batch():
    """Keep one isolated LibreOffice/UNO listener for the whole audit batch."""
    engine = find_calc_engine()
    if engine is None:
        raise RuntimeError("未找到 LibreOffice Calc")
    try:
        import uno  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("已找到 LibreOffice，但当前 Python 未安装 python3-uno") from exc
    with uno_service(engine.command_prefix):
        yield
