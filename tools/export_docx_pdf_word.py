from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client


input_path = Path(sys.argv[1]).resolve()
output_path = Path(sys.argv[2]).resolve()
output_path.parent.mkdir(parents=True, exist_ok=True)

pythoncom.CoInitialize()
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
document = None
try:
    document = word.Documents.Open(str(input_path), ReadOnly=True, AddToRecentFiles=False)
    document.Fields.Update()
    document.ExportAsFixedFormat(str(output_path), 17)
finally:
    if document is not None:
        document.Close(False)
    word.Quit()
    pythoncom.CoUninitialize()

print(output_path)
