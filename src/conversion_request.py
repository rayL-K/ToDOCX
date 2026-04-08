"""一次转换请求的规范化与校验。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .config import BUILTIN_DEFAULTS
from .errors import ValidationError
from .resource_policy import ResourcePolicy


def _classify_input_type(input_path: Path) -> str:
    suffix = input_path.suffix.lower()
    if suffix in {".md", ".markdown"}:
        return "markdown"
    if suffix == ".tex":
        return "latex"
    if suffix in {".docx", ".doc"}:
        return "docx"
    raise ValidationError(
        f"不支持的文件格式：{suffix or '无扩展名'}",
        code="TODX101",
        hint="请选择 DOCX、Markdown 或 LaTeX 文件。",
    )


@dataclass(frozen=True)
class ConversionRequest:
    """UI 到转换层之间的稳定请求对象。"""

    input_path: Path
    output_dir: Path
    output_path: Path
    input_type: str
    styles: dict[str, Any]
    paragraph_mappings: dict[int, str] = field(default_factory=dict)
    type_overrides: dict[str, str] = field(default_factory=dict)
    resource_policy: ResourcePolicy = field(default_factory=ResourcePolicy)


def build_conversion_request(
    input_file: str,
    output_dir_text: str,
    styles: dict[str, Any],
    *,
    paragraph_mappings: dict[int, str] | None = None,
    type_overrides: dict[str, str] | None = None,
    resource_policy: ResourcePolicy | None = None,
) -> ConversionRequest:
    """从原始输入构建并校验转换请求。"""

    input_path = Path(input_file).expanduser()
    try:
        input_path = input_path.resolve(strict=True)
    except FileNotFoundError as error:
        raise ValidationError(
            "输入文件不存在。",
            code="TODX102",
            hint="请重新选择待转换文件。",
        ) from error

    if not input_path.is_file():
        raise ValidationError(
            "输入路径不是可读取的文件。",
            code="TODX103",
            hint="请重新选择单个 DOCX、Markdown 或 LaTeX 文件。",
        )

    if input_path.suffix.lower() not in BUILTIN_DEFAULTS.supported_input_extensions:
        raise ValidationError(
            f"不支持的文件格式：{input_path.suffix}",
            code="TODX104",
            hint="请选择 DOCX、Markdown 或 LaTeX 文件。",
        )

    raw_output_dir = Path(output_dir_text).expanduser() if output_dir_text else input_path.parent
    if not raw_output_dir.is_absolute():
        raw_output_dir = input_path.parent / raw_output_dir
    output_dir = raw_output_dir.resolve()

    if output_dir.exists() and not output_dir.is_dir():
        raise ValidationError(
            "输出路径不是文件夹。",
            code="TODX105",
            hint="请改为选择一个存在的目录，或留空使用源文件所在目录。",
        )

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        raise ValidationError(
            "无法创建或写入输出目录。",
            code="TODX106",
            hint="请改用当前用户有写权限的目录。",
        ) from error

    output_path = output_dir / f"{input_path.stem}_formatted.docx"

    return ConversionRequest(
        input_path=input_path,
        output_dir=output_dir,
        output_path=output_path,
        input_type=_classify_input_type(input_path),
        styles=dict(styles),
        paragraph_mappings=dict(paragraph_mappings or {}),
        type_overrides=dict(type_overrides or {}),
        resource_policy=resource_policy or ResourcePolicy(),
    )
