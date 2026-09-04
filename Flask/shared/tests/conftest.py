"""统一测试导入路径：无论从哪个目录运行 pytest，都能同时以
``base_audit`` 和 ``src.base_audit`` 两种风格导入共享核心。"""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent.parent / "src"
for entry in (str(_SRC), str(_SRC.parent)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
