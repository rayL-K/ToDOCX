"""应用目录与文件路径。"""

from __future__ import annotations

import os
from pathlib import Path

from .config import BUILTIN_DEFAULTS


def get_user_data_dir() -> Path:
    """获取用户专属数据目录。"""

    appdata = os.getenv("APPDATA")
    if appdata:
        return Path(appdata) / BUILTIN_DEFAULTS.app_dir_name
    return Path.home() / ".config" / BUILTIN_DEFAULTS.app_dir_name


def ensure_user_data_dir() -> Path:
    """确保用户数据目录存在。"""

    path = get_user_data_dir()
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_templates_dir() -> Path:
    """模板目录（项目内 templates/）。"""

    path = Path(__file__).resolve().parent.parent / "templates"
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_logs_dir() -> Path:
    """日志目录。"""

    path = ensure_user_data_dir() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_log_file_path() -> Path:
    """主日志文件路径。"""

    return get_logs_dir() / "todocx.log"


def get_settings_path() -> Path:
    """用户设置文件路径。"""

    return ensure_user_data_dir() / "settings.json"
