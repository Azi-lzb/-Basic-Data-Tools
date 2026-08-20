from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path


MAX_RECENT_PATHS = 10


@dataclass(frozen=True)
class FavoritePath:
    name: str
    path: str


@dataclass
class UserSettings:
    last_input_dir: str = ""
    last_template_dir: str = ""
    last_external_file: str = ""
    last_output_dir: str = ""
    output_pinned: bool = False
    recursive_folders: bool = True
    favorite_input_dirs: list[FavoritePath] = field(default_factory=list)
    favorite_template_dirs: list[FavoritePath] = field(default_factory=list)
    favorite_external_files: list[FavoritePath] = field(default_factory=list)
    favorite_output_dirs: list[FavoritePath] = field(default_factory=list)
    recent_input_dirs: list[str] = field(default_factory=list)
    recent_template_dirs: list[str] = field(default_factory=list)
    recent_external_files: list[str] = field(default_factory=list)
    recent_output_dirs: list[str] = field(default_factory=list)


class SettingsStore:
    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self) -> UserSettings:
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return UserSettings()
        if not isinstance(payload, dict):
            return UserSettings()
        last_input_dir = str(payload.get("last_input_dir") or "")
        last_output_dir = str(payload.get("last_output_dir") or "")
        # 旧设置没有书钉字段时，保留用户原先手动改过的非默认输出目录。
        # “审核结果”“执行结果”“执行结果_skip”均为历次默认值；都不应
        # 被误判为用户手动固定的目录。
        legacy_pinned = bool(
            last_input_dir
            and last_output_dir
            and last_output_dir not in {
                str(Path(last_input_dir) / "审核结果"),
                str(Path(last_input_dir) / "执行结果"),
                str(Path(last_input_dir) / "执行结果_skip"),
            }
        )
        return UserSettings(
            last_input_dir=last_input_dir,
            last_template_dir=str(payload.get("last_template_dir") or ""),
            last_external_file=str(payload.get("last_external_file") or ""),
            last_output_dir=last_output_dir,
            output_pinned=bool(payload.get("output_pinned", legacy_pinned)),
            # 说明报送文件通常按机构放在子目录，默认保持递归发现。
            recursive_folders=bool(payload.get("recursive_folders", True)),
            favorite_input_dirs=self._favorites(payload.get("favorite_input_dirs")),
            favorite_template_dirs=self._favorites(
                payload.get("favorite_template_dirs")
            ),
            favorite_external_files=self._favorites(
                payload.get("favorite_external_files")
            ),
            favorite_output_dirs=self._favorites(payload.get("favorite_output_dirs")),
            recent_input_dirs=self._paths(payload.get("recent_input_dirs")),
            recent_template_dirs=self._paths(payload.get("recent_template_dirs")),
            recent_external_files=self._paths(payload.get("recent_external_files")),
            recent_output_dirs=self._paths(payload.get("recent_output_dirs")),
        )

    def save(self, settings: UserSettings) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"version": 1, **asdict(settings)}
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporary.replace(self.path)

    @staticmethod
    def _paths(value: object) -> list[str]:
        if not isinstance(value, list):
            return []
        return [str(item) for item in value if isinstance(item, str) and item.strip()][
            :MAX_RECENT_PATHS
        ]

    @staticmethod
    def _favorites(value: object) -> list[FavoritePath]:
        if not isinstance(value, list):
            return []
        result = []
        for item in value:
            if not isinstance(item, dict) or not item.get("path"):
                continue
            result.append(
                FavoritePath(
                    str(item.get("name") or Path(str(item["path"])).name),
                    str(item["path"]),
                )
            )
        return result


def remember_path(items: list[str], value: str) -> list[str]:
    normalized = str(Path(value).resolve())
    result = [normalized]
    result.extend(item for item in items if str(Path(item)) != normalized)
    return result[:MAX_RECENT_PATHS]


def add_favorite(
    items: list[FavoritePath], name: str, value: str
) -> list[FavoritePath]:
    normalized = str(Path(value).resolve())
    result = [item for item in items if str(Path(item.path)) != normalized]
    result.append(FavoritePath(name.strip() or Path(normalized).name, normalized))
    return result
