"""本地资源访问策略。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote, urlparse

from .config import BUILTIN_DEFAULTS
from .errors import ResourceAccessError


@dataclass(frozen=True)
class ResourcePolicy:
    """Markdown 资源访问策略。"""

    mode: str = "local_only"
    max_image_bytes: int = BUILTIN_DEFAULTS.max_image_bytes
    allowed_image_extensions: tuple[str, ...] = BUILTIN_DEFAULTS.supported_image_extensions
    allowed_embedded_image_mime_types: tuple[str, ...] = (
        BUILTIN_DEFAULTS.supported_embedded_image_mime_types
    )

    def resolve_local_image(self, base_dir: Path, src: str) -> Path:
        """解析并校验本地图片路径。"""

        raw_src = (src or "").strip()
        if not raw_src:
            raise ResourceAccessError(
                "图片路径不能为空。",
                code="TODX201",
                hint="请把图片放在 Markdown 文件同目录或子目录，再使用相对路径引用。",
            )

        parsed = urlparse(raw_src)
        if parsed.scheme in {"http", "https"}:
            raise ResourceAccessError(
                "Markdown 不支持远程图片链接。",
                code="TODX202",
                hint="请先把图片下载到本地，再引用当前文档目录中的相对路径。",
            )

        if parsed.scheme and parsed.scheme != "file":
            raise ResourceAccessError(
                "图片链接格式不受支持。",
                code="TODX203",
                hint="请改用当前文档目录中的本地图片相对路径。",
            )

        candidate = Path(unquote(parsed.path if parsed.scheme == "file" else raw_src))
        if candidate.is_absolute():
            raise ResourceAccessError(
                "Markdown 图片必须使用当前文档目录内的相对路径。",
                code="TODX204",
                hint="请把图片移动到文档同目录或子目录后再引用。",
            )

        resolved_base = base_dir.resolve()
        resolved_path = (resolved_base / candidate).resolve()
        if resolved_path != resolved_base and resolved_base not in resolved_path.parents:
            raise ResourceAccessError(
                "Markdown 图片路径超出了当前文档目录。",
                code="TODX205",
                hint="请把图片移动到当前 Markdown 文件目录或其子目录中。",
            )

        if resolved_path.suffix.lower() not in self.allowed_image_extensions:
            raise ResourceAccessError(
                "图片格式不受支持。",
                code="TODX206",
                hint="请使用 PNG、JPG、JPEG、GIF、BMP 或 WEBP 图片。",
            )

        if not resolved_path.exists() or not resolved_path.is_file():
            raise ResourceAccessError(
                "图片文件不存在。",
                code="TODX207",
                hint="请检查 Markdown 中的相对路径是否正确。",
            )

        self.validate_resource_size(resolved_path.stat().st_size, resolved_path.name)
        return resolved_path

    def validate_embedded_image_mime(self, mime_type: str) -> None:
        """校验 data URL 图片 MIME。"""

        if mime_type not in self.allowed_embedded_image_mime_types:
            raise ResourceAccessError(
                "嵌入图片格式不受支持。",
                code="TODX208",
                hint="请改用 PNG、JPG、JPEG、GIF、BMP 或 WEBP 格式。",
            )

    def validate_resource_size(self, size_bytes: int, label: str = "图片") -> None:
        """校验资源大小。"""

        if size_bytes > self.max_image_bytes:
            max_size_mb = max(self.max_image_bytes // (1024 * 1024), 1)
            raise ResourceAccessError(
                f"{label} 过大，无法安全导入。",
                code="TODX209",
                hint=f"请将资源压缩到 {max_size_mb} MB 以内后重试。",
            )
