from __future__ import annotations

import sys
from pathlib import Path

import pypdfium2 as pdfium


input_path = Path(sys.argv[1]).resolve()
output_dir = Path(sys.argv[2]).resolve()
output_dir.mkdir(parents=True, exist_ok=True)
pdf = pdfium.PdfDocument(input_path)
for index in range(len(pdf)):
    page = pdf[index]
    bitmap = page.render(scale=2.0)
    image = bitmap.to_pil()
    image.save(output_dir / f"page-{index + 1}.png")
    page.close()
pdf.close()
print(f"pages={index + 1 if 'index' in locals() else 0} output={output_dir}")
