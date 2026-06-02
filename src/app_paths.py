"""应用目录与文件路径。"""

from __future__ import annotations

import os
import sys
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


def _is_frozen() -> bool:
    """判断是否在 PyInstaller 打包的 exe 中运行。"""
    return getattr(sys, "frozen", False)


def get_templates_dir() -> Path:
    """模板目录。

    - 开发阶段：项目根目录下的 templates/
    - 打包分发后：%%APPDATA%%\ToDOCX\templates\（首次自动从包内复制）
    """

    if _is_frozen():
        # 打包运行 —— 用用户目录
        user_dir = ensure_user_data_dir() / "templates"
        user_dir.mkdir(parents=True, exist_ok=True)
        return user_dir

    # 开发运行 —— 用项目内目录
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
