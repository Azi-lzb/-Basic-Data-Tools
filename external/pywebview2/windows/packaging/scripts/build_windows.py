"""Unicode-safe PyInstaller entry point for the Windows batch launcher."""

from __future__ import annotations

from pathlib import Path
import shutil


def copy_user_guide(root: Path) -> None:
    """Copy the optional local guide before PyInstaller mutates its process state.

    The executable is the build artefact. A local guide being open or a rare
    Windows path-normalisation issue must never turn a completed EXE build into
    a failed build.
    """
    source = root / "基础数据审核工具使用说明.docx"
    target_dir = root / "dist"
    target = target_dir / source.name
    if not source.is_file():
        print(f"Warning: user guide was not found and was not copied: {source}")
        return
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        # Do this before PyInstaller.__main__.run().  Some older PyInstaller
        # environments alter path handling while building a one-file package.
        shutil.copy2(source, target)
    except OSError as exc:
        print(f"Warning: executable was built, but user guide could not be copied: {exc}")


def main() -> None:
    import PyInstaller.__main__

    root = Path(__file__).resolve().parents[2]
    spec = root / "packaging" / "specs" / "基础数据审核工具.spec"
    if not spec.is_file():
        raise FileNotFoundError(f"Missing spec file: {spec}")
    # 使用说明与 EXE 同级，供工作台的小书入口以系统默认程序打开。
    # 该文件复制失败只给出警告，不能影响 EXE 本体构建。
    copy_user_guide(root)
    PyInstaller.__main__.run(["--noconfirm", "--clean", str(spec)])


if __name__ == "__main__":
    main()
