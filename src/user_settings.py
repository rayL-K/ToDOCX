"""用户设置存储。"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

from .app_paths import get_settings_path
from .config import BUILTIN_DEFAULTS
from .diagnostics import get_logger, log_exception
from .persistence import atomic_write_json


@dataclass
class UserSettings:
    """用户持久化设置。"""

    schema_version: int = BUILTIN_DEFAULTS.settings_schema_version
    last_output_dir: str = ""


class UserSettingsStore:
    """用户设置读写。"""

    def __init__(self, settings_path: str | Path | None = None) -> None:
        self.settings_path = Path(settings_path) if settings_path else get_settings_path()
        self.logger = get_logger("settings")

    def load(self) -> UserSettings:
        """加载设置；损坏时回退到默认值。"""

        if not self.settings_path.exists():
            return UserSettings()

        try:
            with open(self.settings_path, "r", encoding="utf-8") as handle:
                raw_data = json.load(handle)
        except json.JSONDecodeError as error:
            log_exception(
                self.logger,
                "读取用户设置失败，已回退默认值",
                error,
                path=str(self.settings_path),
            )
            return UserSettings()
        except OSError as error:
            log_exception(
                self.logger,
                "打开用户设置失败，已回退默认值",
                error,
                path=str(self.settings_path),
            )
            return UserSettings()

        return self._from_dict(raw_data)

    def save(self, settings: UserSettings) -> None:
        """保存设置。"""

        atomic_write_json(self.settings_path, asdict(settings))

    def update_last_output_dir(self, output_dir: str) -> UserSettings:
        """更新最近一次输出目录。"""

        settings = self.load()
        settings.last_output_dir = output_dir
        self.save(settings)
        return settings

    @staticmethod
    def _from_dict(raw_data: dict) -> UserSettings:
        settings = UserSettings()
        if isinstance(raw_data, dict):
            last_output_dir = raw_data.get("last_output_dir")
            if isinstance(last_output_dir, str):
                settings.last_output_dir = last_output_dir
        return settings
