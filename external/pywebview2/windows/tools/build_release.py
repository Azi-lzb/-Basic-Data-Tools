from __future__ import annotations

import hashlib
import shutil
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
EXE = ROOT / "dist" / "基础数据审核工具.exe"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def available_release_dir() -> Path:
    release_root = ROOT / "release"
    release_root.mkdir(parents=True, exist_ok=True)
    base = release_root / f"基础数据审核工具_{datetime.now():%Y%m%d}"
    if not base.exists():
        return base
    return release_root / f"{base.name}_{datetime.now():%H%M%S}"


def main() -> None:
    if not EXE.is_file():
        raise FileNotFoundError(f"未找到打包后的程序：{EXE}")

    release_dir = available_release_dir()
    release_dir.mkdir()
    template_dir = release_dir / "templates"
    template_dir.mkdir()
    data_dir = release_dir / "data"
    data_dir.mkdir()

    shutil.copy2(EXE, release_dir / EXE.name)
    template_source = ROOT / "2026-07-31" / "模板文件"
    if not template_source.is_dir():
        template_source = ROOT / "templates"
    for template in sorted(template_source.glob("*.xlsx")):
        if template.name.startswith(("!", "！")):
            shutil.copy2(template, template_dir / template.name)
    for name in ("启动pywebview2-windows.bat", "基础数据审核工具使用说明.docx", "README.md"):
        shutil.copy2(ROOT / name, release_dir / name)
    history_source = ROOT / "历史审核说明.xlsx"
    if history_source.is_file():
        shutil.copy2(history_source, release_dir / history_source.name)

    (data_dir / "目录说明.txt").write_text(
        "本目录保存程序内部 JSON 设置（流程配置、常用路径和模板索引）。\n"
        "程序首次启动时会将本目录设为隐藏；升级时请保留整个 data 目录。\n",
        encoding="utf-8-sig",
    )
    (release_dir / "发布说明.txt").write_text(
        "基础数据审核工具\n\n"
        "运行条件：Windows 10/11，并已安装 Microsoft Excel。\n"
        "启动方法：双击“基础数据审核工具.exe”或“启动pywebview2-windows.bat”。\n"
        "模板位置：默认是 templates 目录（本次从 2026-07-31/模板文件 复制），也可在工作台选择其他模板文件目录。\n"
        "模板匹配：根据源数据，只在当前选择的模板文件目录中自动匹配。\n"
        "历史审核说明：根目录的 历史审核说明.xlsx；程序升级时不要覆盖或删除。\n"
        "程序设置：data 目录（默认隐藏）；程序升级时不要覆盖或删除。\n"
        "常用路径：选中后点击“使用选中路径”或直接双击即可回填。\n"
        "审核期：用于结果命名和历史问题对比，不参与公式计算。\n"
        "公式拆列：五段公式结果会拆为对比值、参考值、差值和详细说明，并保留公式原文。\n"
        "外部文件：完整审核时将模板公式所需工作表复制到审核副本；源文件和外部文件不修改。\n"
        "模块、流程和按钮由设置中心维护，保存在 data/流程配置.json；如误删，可执行 基础数据审核工具.exe --init-config 恢复，历史审核说明不会删除。\n",
        encoding="utf-8-sig",
    )

    manifest_lines = []
    for path in sorted(item for item in release_dir.rglob("*") if item.is_file()):
        relative = path.relative_to(release_dir)
        manifest_lines.append(f"{sha256(path)}  {relative}")
    (release_dir / "SHA256.txt").write_text(
        "\n".join(manifest_lines) + "\n", encoding="utf-8"
    )

    archive = shutil.make_archive(str(release_dir), "zip", release_dir.parent, release_dir.name)
    print(f"release={release_dir}")
    print(f"archive={archive}")


if __name__ == "__main__":
    main()
