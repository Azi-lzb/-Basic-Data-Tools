"""LibreOffice Calc calculation support for UOS/Deepin and Linux.

Both direct DEB installations and the UOS application-store (Linglong)
edition are supported. Calculation always uses a private profile and staged
files, so it never touches a user's open LibreOffice session or a source file.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path


LINGLONG_APP_DEFAULT = "org.libreoffice.libreoffice"


class LibreOfficeUnavailableError(RuntimeError):
    """Raised when a native LibreOffice Calc process cannot be used."""


@dataclass(frozen=True)
class CalcEngine:
    name: str
    source: str
    command_prefix: tuple[str, ...]
    display: str


@dataclass(frozen=True)
class LibreOfficeCalculationResult:
    workbook_path: Path
    engine_display: str
    elapsed_seconds: float


def _executable(path: Path) -> Path | None:
    try:
        return path if path.is_file() and os.access(path, os.X_OK) else None
    except OSError:
        return None


def _direct_candidates() -> list[tuple[Path, str]]:
    candidates: list[tuple[Path, str]] = []
    configured = os.environ.get("BASE_AUDIT_SOFFICE", "").strip()
    if configured:
        candidates.append((Path(configured).expanduser(), "环境变量 BASE_AUDIT_SOFFICE"))
    configured_dir = os.environ.get("BASE_AUDIT_LIBREOFFICE_DIR", "").strip()
    if configured_dir:
        root = Path(configured_dir).expanduser()
        candidates.extend((root / name, "环境变量 BASE_AUDIT_LIBREOFFICE_DIR") for name in (
            "soffice", "libreoffice", "scalc", "program/soffice",
        ))
    for command in ("soffice", "libreoffice", "scalc"):
        located = shutil.which(command)
        if located:
            candidates.append((Path(located), "PATH"))
    candidates.extend((Path(item), "系统路径") for item in (
        "/usr/bin/soffice", "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice", "/usr/lib64/libreoffice/program/soffice",
    ))
    for pattern in ("/opt/libreoffice*/program/soffice", "/opt/LibreOffice*/program/soffice"):
        candidates.extend((candidate, "版本化 DEB 安装目录") for candidate in sorted(Path("/").glob(pattern.lstrip("/"))))
    home = Path.home()
    for root in (home / "Desktop", home / "Downloads", Path("/opt")):
        if not root.is_dir():
            continue
        for pattern in ("*/program/soffice", "libreoffice*/program/soffice", "LibreOffice*/program/soffice"):
            try:
                candidates.extend((candidate, "手动安装目录") for candidate in root.glob(pattern))
            except OSError:
                continue
    return candidates


def _find_direct_engine() -> CalcEngine | None:
    for candidate, source in _direct_candidates():
        executable = _executable(candidate)
        if executable is not None:
            return CalcEngine("LibreOffice Calc", source, (str(executable),), str(executable))
    return None


def _find_linglong_engine() -> CalcEngine | None:
    ll_cli = shutil.which("ll-cli")
    if not ll_cli:
        return None
    app_id = os.environ.get("BASE_AUDIT_LINGLONG_APP", LINGLONG_APP_DEFAULT).strip() or LINGLONG_APP_DEFAULT
    try:
        listed = subprocess.run(
            [ll_cli, "list"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, timeout=15, check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    if listed.returncode != 0 or app_id not in (listed.stdout or "") + "\n" + (listed.stderr or ""):
        return None
    prefix = (ll_cli, "run", app_id, "--", "soffice")
    return CalcEngine("LibreOffice Calc（玲珑商店版）", "玲珑商店版", prefix, "ll-cli run {} -- soffice".format(app_id))


def find_calc_engine() -> CalcEngine | None:
    """Find the configured direct or Linglong LibreOffice Calc entry."""
    mode = os.environ.get("BASE_AUDIT_ENGINE", "auto").strip().casefold() or "auto"
    if mode not in {"auto", "direct", "ll-cli"}:
        mode = "auto"
    if mode in {"auto", "direct"}:
        direct = _find_direct_engine()
        if direct is not None:
            return direct
    if mode in {"auto", "ll-cli"}:
        return _find_linglong_engine()
    return None


def find_soffice() -> Path | None:
    """Backward-compatible direct-entry lookup; never returns Linglong."""
    engine = _find_direct_engine()
    return Path(engine.command_prefix[0]) if engine is not None else None


class LibreOfficeCalculator:
    """Recalculate an xlsx audit copy with a private LibreOffice profile."""

    engine_name = "LibreOffice Calc"

    def __init__(self, executable: Path | None = None, *, timeout_seconds: int = 300) -> None:
        self.executable = executable
        self.timeout_seconds = timeout_seconds

    def require_available(self) -> CalcEngine:
        if self.executable is not None:
            executable = _executable(self.executable)
            if executable is not None:
                return CalcEngine("LibreOffice Calc", "指定入口", (str(executable),), str(executable))
        engine = find_calc_engine()
        if engine is None:
            raise LibreOfficeUnavailableError(
                "未找到 LibreOffice Calc。可安装 LibreOffice、设置 BASE_AUDIT_SOFFICE 指向 soffice、"
                "设置 BASE_AUDIT_LIBREOFFICE_DIR 指向安装目录，或确认应用商店已安装 LibreOffice。"
            )
        return engine

    def version(self, engine: CalcEngine | None = None) -> str:
        """Return the installed Calc version without opening a workbook."""
        selected = engine or self.require_available()
        try:
            completed = subprocess.run(
                [*selected.command_prefix, "--version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=15,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired):
            return "版本未识别"
        text = "\n".join(item for item in (completed.stdout, completed.stderr) if item).strip()
        return text.splitlines()[0].strip() if completed.returncode == 0 and text else "版本未识别"

    def recalculate(self, workbook_path: Path) -> LibreOfficeCalculationResult:
        engine = self.require_available()
        workbook_path = workbook_path.resolve()
        if not workbook_path.is_file():
            raise FileNotFoundError("待计算工作簿不存在：{}".format(workbook_path))
        if workbook_path.suffix.casefold() != ".xlsx":
            raise ValueError("LibreOffice 计算引擎当前仅支持 .xlsx 审核副本")

        started = time.monotonic()
        # 应用商店的玲珑版 LibreOffice 通常无法读取 /tmp 中的宿主文件。
        # 将临时副本放到用户 HOME 的缓存目录，且采用 ASCII 文件名和 file URI，
        # 同时规避中文路径/文件名在 soffice CLI 中被错误解析为 source 的问题。
        cache_root = Path(os.environ.get("XDG_CACHE_HOME", "")).expanduser() if os.environ.get("XDG_CACHE_HOME") else Path.home() / ".cache"
        cache_root = cache_root / "base_audit_lo"
        try:
            cache_root.mkdir(parents=True, exist_ok=True)
        except OSError:
            # 开发/受限环境可没有 HOME 缓存权限；统信玲珑版仍优先走上面的
            # HOME 缓存路径，只有创建失败时才退回系统临时目录。
            cache_root = Path(tempfile.gettempdir()) / "base_audit_lo"
            cache_root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="run_", dir=cache_root) as temp_text:
            temp_dir = Path(temp_text)
            source_dir, output_dir, profile_dir = temp_dir / "source", temp_dir / "output", temp_dir / "profile"
            source_dir.mkdir()
            output_dir.mkdir()
            profile_dir.mkdir()
            staged = source_dir / "audit_input.xlsx"
            shutil.copy2(workbook_path, staged)
            command = [
                *engine.command_prefix, "--headless", "--nologo", "--nodefault", "--nolockcheck",
                "--nofirststartwizard", "-env:UserInstallation=" + profile_dir.as_uri(),
                "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(output_dir), staged.as_uri(),
            ]
            try:
                completed = subprocess.run(
                    command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                    timeout=self.timeout_seconds, check=False,
                )
            except subprocess.TimeoutExpired as exc:
                raise TimeoutError("LibreOffice 公式计算超过 {} 秒：{}".format(self.timeout_seconds, workbook_path.name)) from exc
            generated = output_dir / staged.name
            if completed.returncode != 0 or not generated.is_file():
                details = "\n".join(part for part in (completed.stderr, completed.stdout) if part).strip()
                hint = ""
                if "source could not be loaded" in details.casefold():
                    hint = "；已改用用户缓存目录暂存输入文件。若仍失败，请确认当前 LibreOffice/玲珑应用可访问 {}".format(cache_root)
                raise LibreOfficeUnavailableError(
                    "LibreOffice 未能计算“{}”（{}）：{}{}".format(workbook_path.name, engine.display, (details or "无输出")[:800], hint)
                )
            shutil.copy2(generated, workbook_path)
        return LibreOfficeCalculationResult(workbook_path, engine.display, time.monotonic() - started)
