"""共享的样式处理工具。"""

from copy import deepcopy
from typing import Any, Dict, Iterable, Mapping, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from .config import DEFAULT_STYLES, get_font_size_pt


ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

DEFAULT_MULTIPLE_SPACING = {
    "单倍行距": 1.0,
    "1.5倍行距": 1.5,
    "2倍行距": 2.0,
    "倍数行距": 1.5,
}


def merge_styles(overrides: Optional[Mapping[str, Any]] = None) -> Dict[str, Any]:
    """深度合并样式，保证部分覆盖不会丢默认项。"""
    merged = deepcopy(DEFAULT_STYLES)
    if not overrides:
        return merged

    for key, value in overrides.items():
        if isinstance(value, Mapping) and isinstance(merged.get(key), Mapping):
            merged[key] = {**merged[key], **value}
        else:
            merged[key] = deepcopy(value)

    return merged


def get_style(styles: Optional[Mapping[str, Any]], type_id: str, fallback: str = "body") -> Dict[str, Any]:
    """获取指定类型的样式。"""
    if not styles:
        styles = DEFAULT_STYLES
    return dict(styles.get(type_id) or styles.get(fallback) or {})


def get_font_size_value(style_config: Mapping[str, Any], default: float = 12.0) -> float:
    """将字号配置转换为磅值。"""
    size = style_config.get("font_size", default)
    if isinstance(size, str):
        return float(get_font_size_pt(size))
    try:
        return float(size)
    except (TypeError, ValueError):
        return float(default)


def apply_alignment(paragraph_format, alignment: Optional[str], default: str = "left") -> None:
    """应用段落对齐方式。"""
    paragraph_format.alignment = ALIGNMENT_MAP.get(alignment or default, ALIGNMENT_MAP[default])


def _normalize_line_spacing(style_config: Mapping[str, Any]) -> tuple[str, float]:
    spacing_type = style_config.get("line_spacing_type", "1.5倍行距")
    spacing_value = style_config.get(
        "line_spacing_value",
        DEFAULT_MULTIPLE_SPACING.get(spacing_type, 1.5),
    )

    if isinstance(spacing_value, str):
        try:
            spacing_value = float(spacing_value)
        except ValueError:
            spacing_value = DEFAULT_MULTIPLE_SPACING.get(spacing_type, 1.5)

    if spacing_type == "固定值":
        return "fixed", float(spacing_value or 20)

    return "multiple", float(spacing_value or DEFAULT_MULTIPLE_SPACING.get(spacing_type, 1.5))


def apply_line_spacing(paragraph_format, style_config: Mapping[str, Any]) -> None:
    """应用行距设置。"""
    spacing_mode, spacing_value = _normalize_line_spacing(style_config)

    if spacing_mode == "fixed":
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph_format.line_spacing = Pt(spacing_value)
        return

    paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph_format.line_spacing = spacing_value


def _hex_to_rgb(color: str) -> Optional[RGBColor]:
    if not isinstance(color, str):
        return None
    color = color.strip()
    if not color.startswith("#") or len(color) != 7:
        return None
    try:
        return RGBColor(int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16))
    except ValueError:
        return None


def apply_run_style(
    run,
    style_config: Mapping[str, Any],
    *,
    default_cn: str = "宋体",
    default_en: str = "Times New Roman",
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
    default_size: float = 12.0,
) -> None:
    """把共享字体样式应用到 run。"""
    font_cn = style_config.get(
        "font_name_cn",
        style_config.get("font_cn", style_config.get("font_name", default_cn)),
    )
    font_en = style_config.get(
        "font_name_en",
        style_config.get("font_en", style_config.get("font_name", default_en)),
    )
    font_size = get_font_size_value(style_config, default_size)

    run.font.name = font_en
    run.font.size = Pt(font_size)

    if bold is None and "bold" in style_config:
        bold = bool(style_config.get("bold"))
    if bold is not None:
        run.font.bold = bool(bold)

    if italic is None and "italic" in style_config:
        italic = bool(style_config.get("italic"))
    if italic is not None:
        run.font.italic = bool(italic)

    if color is None:
        color = style_config.get("color")
    rgb_color = _hex_to_rgb(color)
    if rgb_color is not None:
        run.font.color.rgb = rgb_color

    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)

    r_fonts.set(qn("w:ascii"), font_en)
    r_fonts.set(qn("w:hAnsi"), font_en)
    r_fonts.set(qn("w:eastAsia"), font_cn)
    r_fonts.set(qn("w:cs"), font_en)


def apply_runs_style(
    runs: Iterable,
    style_config: Mapping[str, Any],
    *,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
    default_cn: str = "宋体",
    default_en: str = "Times New Roman",
    default_size: float = 12.0,
) -> None:
    """批量设置多个 run 的字体样式。"""
    for run in runs:
        apply_run_style(
            run,
            style_config,
            default_cn=default_cn,
            default_en=default_en,
            bold=bold,
            italic=italic,
            color=color,
            default_size=default_size,
        )
