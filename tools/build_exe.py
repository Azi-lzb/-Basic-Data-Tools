"""Unicode-safe PyInstaller entry point for the Windows batch launcher."""

from __future__ import annotations

from pathlib import Path
import shutil


def main() -> None:
    import PyInstaller.__main__

    root = Path(__file__).resolve().parents[1]
    spec = root / "基础数据审核工具.spec"
    if not spec.is_file():
        raise FileNotFoundError(f"Missing spec file: {spec}")
    PyInstaller.__main__.run(["--noconfirm", "--clean", str(spec)])
    # 使用说明与 EXE 同级，供工作台的小书入口以系统默认程序打开。
    shutil.copy2(
        root / "基础数据审核工具使用说明.docx",
        root / "dist" / "基础数据审核工具使用说明.docx",
    )


if __name__ == "__main__":
    main()
