"""Rename the ASCII Win7 staging executable to its public Chinese name.

PyInstaller 4.10's bundled pefile can fail while writing a Unicode output
name in the Python 3.7 Win7 build environments.  The package itself is built
with an ASCII name and this small Unicode-safe Python step runs afterwards.
"""

from __future__ import print_function

import sys
import shutil
from pathlib import Path


NAMES = {
    "x64": ("BaseAuditTool_Win7.exe", "基础数据审核工具_Win7.exe"),
    "x86": ("BaseAuditTool_Win7_x86.exe", "基础数据审核工具_Win7_x86.exe"),
}


def main():
    target = (sys.argv[1] if len(sys.argv) > 1 else "").lower()
    if target not in NAMES:
        raise SystemExit("Usage: finalize_win7_build.py x64|x86")
    source_name, public_name = NAMES[target]
    dist_dir = Path.cwd() / "dist"
    source = dist_dir / source_name
    destination = dist_dir / public_name
    if not source.is_file():
        raise SystemExit("Win7 staging EXE was not generated: {}".format(source))
    if destination.exists():
        destination.unlink()
    source.replace(destination)
    print("Created {}".format(destination))
    _copy_companions(dist_dir)


def _copy_companions(dist_dir):
    """Copy non-code companion files without failing a completed EXE build.

    Word can keep the existing guide open through a memory-mapped section.
    In that case replacing the guide is impossible, but the newly built EXE
    remains valid, so emit a short warning instead of letting cmd show a raw
    system error and make the build look broken.
    """
    project_root = Path.cwd()
    data_dir = dist_dir / "data"
    data_dir.mkdir(exist_ok=True)
    history_config = project_root / "data" / "config.xlsx"
    output_config = data_dir / "config.xlsx"
    if history_config.is_file() and not output_config.exists():
        shutil.copy2(str(history_config), str(output_config))
    for guide in project_root.glob("*.docx"):
        try:
            shutil.copy2(str(guide), str(dist_dir / guide.name))
        except (OSError, PermissionError) as exc:
            print("WARNING: guide is in use; skipped update: {} ({})".format(guide.name, exc))


if __name__ == "__main__":
    main()
