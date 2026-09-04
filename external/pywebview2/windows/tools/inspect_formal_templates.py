from __future__ import annotations

import json
from pathlib import Path

import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = ROOT / "2026-07-31" / "模板文件"
XL_CELL_TYPE_FORMULAS = -4123


def main() -> None:
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        for path in sorted(TEMPLATE_DIR.glob("！*.xlsx")):
            workbook = excel.Workbooks.Open(
                str(path), UpdateLinks=0, ReadOnly=True, AddToMru=False
            )
            try:
                report = {"file": path.name, "sheets": []}
                for sheet in workbook.Worksheets:
                    used = sheet.UsedRange
                    formulas = []
                    try:
                        formula_cells = used.SpecialCells(XL_CELL_TYPE_FORMULAS)
                        for cell in formula_cells.Cells:
                            formulas.append(
                                {
                                    "cell": str(cell.Address).replace("$", ""),
                                    "formula": cell.Formula,
                                }
                            )
                            if len(formulas) >= 12:
                                break
                    except Exception:
                        pass
                    comments = []
                    try:
                        for comment in sheet.Comments:
                            comments.append(
                                {
                                    "cell": str(comment.Parent.Address).replace("$", ""),
                                    "text": comment.Text(),
                                }
                            )
                            if len(comments) >= 30:
                                break
                    except Exception:
                        pass
                    report["sheets"].append(
                        {
                            "name": sheet.Name,
                            "used_range": str(used.Address).replace("$", ""),
                            "rows": used.Rows.Count,
                            "columns": used.Columns.Count,
                            "formula_count": len(formulas),
                            "formula_samples": formulas,
                            "comment_count": sheet.Comments.Count,
                            "comment_samples": comments,
                        }
                    )
                print(json.dumps(report, ensure_ascii=False))
            finally:
                workbook.Close(False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
