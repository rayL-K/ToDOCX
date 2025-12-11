"""版本信息 - 单一来源"""

__version__ = "1.0.1"
__author__ = "rayL_K"
__app_name__ = "ToDOCX"

# 用于 PyInstaller 版本信息
VERSION_INFO = {
    "version": __version__,
    "company_name": __author__,
    "product_name": __app_name__,
    "description": "ToDOCX - 琐碎排版，一键告别",
    "copyright": f"Copyright © 2025 {__author__}",
}


def get_version_string() -> str:
    """获取显示用的版本字符串"""
    return f"v{__version__} by {__author__}"


def get_version_tuple() -> tuple:
    """获取版本元组，用于 PyInstaller"""
    parts = __version__.split(".")
    # 补齐为4位
    while len(parts) < 4:
        parts.append("0")
    return tuple(int(p) for p in parts[:4])
