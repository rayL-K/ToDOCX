"""应用内置默认配置。"""

from dataclasses import dataclass

APP_NAME = "ToDOCX"
APP_DIR_NAME = APP_NAME
SETTINGS_SCHEMA_VERSION = 1
TEMPLATE_SCHEMA_VERSION = 1
MAX_IMAGE_BYTES = 10 * 1024 * 1024

SUPPORTED_INPUT_EXTENSIONS = (".docx", ".doc", ".md", ".markdown", ".tex")
SUPPORTED_IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")
SUPPORTED_EMBEDDED_IMAGE_MIME_TYPES = (
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/bmp",
    "image/webp",
)


@dataclass(frozen=True)
class BuiltinDefaults:
    """代码内置默认值。"""

    app_name: str = APP_NAME
    app_dir_name: str = APP_DIR_NAME
    settings_schema_version: int = SETTINGS_SCHEMA_VERSION
    template_schema_version: int = TEMPLATE_SCHEMA_VERSION
    max_image_bytes: int = MAX_IMAGE_BYTES
    supported_input_extensions: tuple[str, ...] = SUPPORTED_INPUT_EXTENSIONS
    supported_image_extensions: tuple[str, ...] = SUPPORTED_IMAGE_EXTENSIONS
    supported_embedded_image_mime_types: tuple[str, ...] = SUPPORTED_EMBEDDED_IMAGE_MIME_TYPES


BUILTIN_DEFAULTS = BuiltinDefaults()

# 标准字号映射（字号名称 -> 磅值）
FONT_SIZE_MAP = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}

# 字号列表（用于下拉框）
FONT_SIZE_OPTIONS = list(FONT_SIZE_MAP.keys())

# 行间距类型
LINE_SPACING_TYPES = {
    "单倍行距": {"type": "multiple", "value": 1.0},
    "1.5倍行距": {"type": "multiple", "value": 1.5},
    "2倍行距": {"type": "multiple", "value": 2.0},
    "固定值": {"type": "exact", "value": 20},  # 默认20磅
}

# 默认样式配置（使用新规范）
DEFAULT_STYLES = {
    "heading1": {
        "font_name_cn": "宋体",
        "font_name_en": "Times New Roman",
        "font_size": "小三",  # 15磅
        "bold": True,
        "color": "#000000",
        "space_before": 12,
        "space_after": 6,
        "line_spacing_type": "固定值",
        "line_spacing_value": 20,
        "alignment": "left",
    },
    "heading2": {
        "font_name_cn": "黑体",
        "font_name_en": "Times New Roman",
        "font_size": "四号",  # 14磅
        "bold": True,
        "color": "#000000",
        "space_before": 10,
        "space_after": 5,
        "line_spacing_type": "固定值",
        "line_spacing_value": 20,
        "alignment": "left",
    },
    "heading3": {
        "font_name_cn": "黑体",
        "font_name_en": "Times New Roman",
        "font_size": "小四",  # 12磅
        "bold": True,
        "color": "#000000",
        "space_before": 8,
        "space_after": 4,
        "line_spacing_type": "固定值",
        "line_spacing_value": 20,
        "alignment": "left",
    },
    "heading4": {
        "font_name_cn": "黑体",
        "font_name_en": "Times New Roman",
        "font_size": "小四",
        "bold": False,
        "color": "#000000",
        "space_before": 6,
        "space_after": 3,
        "line_spacing_type": "固定值",
        "line_spacing_value": 20,
        "alignment": "left",
    },
    "body": {
        "font_name_cn": "宋体",
        "font_name_en": "Times New Roman",
        "font_size": "小四",  # 12磅
        "bold": False,
        "color": "#000000",
        "space_before": 0,
        "space_after": 0,
        "line_spacing_type": "固定值",
        "line_spacing_value": 20,
        "first_line_indent": 2,  # 首行缩进字符数
        "alignment": "left",
    },
    "caption": {  # 图表标题样式
        "font_name_cn": "黑体",
        "font_name_en": "Times New Roman",
        "font_size": "小五",  # 9磅
        "bold": False,
        "color": "#000000",
        "space_before": 6,
        "space_after": 6,
        "line_spacing_type": "1.5倍行距",
        "line_spacing_value": 1.5,
        "alignment": "center",
    },
    "code": {
        "font_name_cn": "宋体",
        "font_name_en": "Consolas",
        "font_size": "五号",
        "bold": False,
        "color": "#333333",
        "background": "#f5f5f5",
        "space_before": 6,
        "space_after": 6,
        "line_spacing_type": "单倍行距",
        "line_spacing_value": 1.0,
        "alignment": "left",
    },
    "quote": {
        "font_name_cn": "楷体",
        "font_name_en": "Times New Roman",
        "font_size": "小四",
        "bold": False,
        "italic": True,
        "color": "#666666",
        "space_before": 6,
        "space_after": 6,
        "line_spacing_type": "固定值",
        "line_spacing_value": 20,
        "left_indent": 1,
        "alignment": "left",
    },
    "image": {
        "alignment": "center",
        "max_width": 15,  # cm
        "space_before": 6,
        "space_after": 6,
        "line_spacing_type": "1.5倍行距",
        "line_spacing_value": 1.5,
    },
    "table": {
        "font_name_cn": "宋体",
        "font_name_en": "Times New Roman",
        "font_size": "五号",
        "header_bold": True,
        "alignment": "center",
    },
    "formula": {
        "alignment": "center",
        "space_before": 6,
        "space_after": 6,
        "line_spacing_type": "1.5倍行距",
        "line_spacing_value": 1.5,
    },
}

def get_font_size_pt(size_name: str) -> float:
    """获取字号对应的磅值"""
    return FONT_SIZE_MAP.get(size_name, 12)
