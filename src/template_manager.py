"""模板管理模块。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from .app_paths import get_templates_dir
from .config import BUILTIN_DEFAULTS, DEFAULT_STYLES
from .diagnostics import get_logger, log_exception, log_event
from .errors import TemplateStorageError
from .persistence import atomic_write_json


class TemplateManager:
    """样式模板管理器"""

    SUPPORTED_SCHEMA_VERSIONS = {BUILTIN_DEFAULTS.template_schema_version}

    def __init__(self, template_dir: str | None = None):
        if template_dir is None:
            self.template_dir = get_templates_dir()
            self.legacy_template_dir = Path(__file__).resolve().parent.parent / "templates"
        else:
            self.template_dir = Path(template_dir)
            self.legacy_template_dir = None

        self.template_dir.mkdir(parents=True, exist_ok=True)
        self.logger = get_logger("templates")
        self._migrate_legacy_templates()
        self._ensure_builtin_templates()

    @staticmethod
    def _normalize_template_name(name: str) -> str:
        """将模板名转换为安全文件名。"""
        safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).strip()
        safe_name = safe_name.replace(" ", "_")
        if not safe_name:
            raise ValueError("模板名称不能为空")
        return safe_name

    def _get_template_path(self, name: str) -> Path:
        """获取模板文件路径。"""
        return self.template_dir / f"{self._normalize_template_name(name)}.json"

    @staticmethod
    def _build_template_payload(name: str, styles: Dict[str, Any], description: str = "") -> Dict[str, Any]:
        """构建模板落盘数据。"""

        return {
            "schema_version": BUILTIN_DEFAULTS.template_schema_version,
            "name": name,
            "description": description,
            "styles": styles,
        }

    @classmethod
    def _validate_template_payload(
        cls,
        data: Dict[str, Any],
        *,
        expected_name: str | None = None,
    ) -> Dict[str, Any]:
        """校验模板结构，只接受明确合法的数据。"""

        if not isinstance(data, dict):
            raise TemplateStorageError(
                "模板文件格式无效。",
                code="TODX303",
                hint="请删除该模板后重新保存。",
            )

        schema_version = data.get("schema_version", BUILTIN_DEFAULTS.template_schema_version)
        if not isinstance(schema_version, int) or schema_version not in cls.SUPPORTED_SCHEMA_VERSIONS:
            raise TemplateStorageError(
                "模板版本不受支持。",
                code="TODX309",
                hint="请使用当前版本重新保存模板。",
            )

        name = data.get("name")
        if not isinstance(name, str) or not name.strip():
            raise TemplateStorageError(
                "模板名称无效。",
                code="TODX310",
                hint="请删除该模板后重新保存。",
            )

        if expected_name is not None and name != expected_name:
            raise TemplateStorageError(
                "模板名称与目标文件不一致。",
                code="TODX311",
                hint="请重新保存该模板，避免名称和文件不一致。",
            )

        description = data.get("description", "")
        if not isinstance(description, str):
            raise TemplateStorageError(
                "模板描述格式无效。",
                code="TODX312",
                hint="请重新保存该模板。",
            )

        styles = data.get("styles")
        if not isinstance(styles, dict):
            raise TemplateStorageError(
                "模板样式结构无效。",
                code="TODX313",
                hint="请重新保存该模板。",
            )

        for section_name, section_value in styles.items():
            if not isinstance(section_name, str) or not section_name.strip():
                raise TemplateStorageError(
                    "模板样式包含无效分组。",
                    code="TODX314",
                    hint="请重新保存该模板。",
                )
            if not isinstance(section_value, dict):
                raise TemplateStorageError(
                    "模板样式分组必须是对象。",
                    code="TODX315",
                    hint="请重新保存该模板。",
                )

        return {
            "schema_version": schema_version,
            "name": name,
            "description": description,
            "styles": styles,
        }

    def _read_template_file(self, file_path: Path) -> Dict[str, Any]:
        """读取模板文件。"""

        try:
            with open(file_path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except json.JSONDecodeError as error:
            raise TemplateStorageError(
                "模板文件已损坏，无法读取。",
                code="TODX301",
                hint="请删除该模板，或用备份恢复后再试。",
            ) from error
        except OSError as error:
            raise TemplateStorageError(
                "模板文件无法读取。",
                code="TODX302",
                hint="请确认模板目录可访问，或稍后重试。",
            ) from error

        return self._validate_template_payload(data)

    def _migrate_legacy_templates(self) -> None:
        """把旧模板目录中的模板迁到用户目录。"""

        if not self.legacy_template_dir or not self.legacy_template_dir.exists():
            return

        migrated = 0
        for legacy_file in self.legacy_template_dir.glob("*.json"):
            target_file = self.template_dir / legacy_file.name
            if target_file.exists():
                continue
            try:
                legacy_data = self._read_template_file(legacy_file)
                normalized_payload = self._build_template_payload(
                    legacy_data["name"],
                    legacy_data["styles"],
                    legacy_data.get("description", ""),
                )
                atomic_write_json(target_file, normalized_payload)
                migrated += 1
            except TemplateStorageError as error:
                log_exception(
                    self.logger,
                    "旧模板格式无效，已跳过迁移",
                    error,
                    source=str(legacy_file),
                )
            except OSError as error:
                log_exception(
                    self.logger,
                    "迁移旧模板失败",
                    error,
                    source=str(legacy_file),
                    target=str(target_file),
                )

        if migrated:
            log_event(
                self.logger,
                "已迁移旧模板",
                count=migrated,
                source=str(self.legacy_template_dir),
                target=str(self.template_dir),
            )
    
    def save_template(self, name: str, styles: Dict[str, Any], description: str = "") -> str:
        """保存模板

        Args:
            name: 模板名称
            styles: 样式配置
            description: 模板描述

        Returns:
            模板文件路径
        """
        file_path = self._get_template_path(name)
        template_payload = self._build_template_payload(name, styles, description)
        self._validate_template_payload(template_payload, expected_name=name)

        try:
            atomic_write_json(file_path, template_payload)
        except OSError as error:
            raise TemplateStorageError(
                "模板保存失败。",
                code="TODX304",
                hint="请确认当前用户对模板目录有写权限。",
            ) from error

        log_event(self.logger, "模板已保存", template=str(file_path), name=name)
        return str(file_path)

    def load_template(self, name: str) -> Optional[Dict[str, Any]]:
        """加载模板样式配置

        Args:
            name: 模板名称

        Returns:
            模板样式配置字典，如果不存在返回 None
        """
        try:
            file_path = self._get_template_path(name)
        except ValueError as error:
            raise TemplateStorageError(
                "模板名称无效。",
                code="TODX305",
                hint="请使用字母、数字、空格、下划线或中划线命名模板。",
            ) from error

        if not file_path.exists():
            return None

        data = self._read_template_file(file_path)
        return data.get("styles", {})
    
    def delete_template(self, name: str) -> bool:
        """删除模板
        
        Args:
            name: 模板名称
            
        Returns:
            是否删除成功
        """
        try:
            file_path = self._get_template_path(name)
        except ValueError as error:
            raise TemplateStorageError(
                "模板名称无效。",
                code="TODX306",
                hint="请重新选择要删除的模板。",
            ) from error

        if file_path.exists():
            try:
                file_path.unlink()
            except OSError as error:
                raise TemplateStorageError(
                    "模板删除失败。",
                    code="TODX307",
                    hint="请确认模板文件未被占用后重试。",
                ) from error
            log_event(self.logger, "模板已删除", template=str(file_path), name=name)
            return True
        return False
    
    def list_templates(self) -> List[Dict[str, Any]]:
        """列出所有用户模板

        Returns:
            模板列表，每个元素包含 name、description、file、status
        """
        templates = []

        for file_path in sorted(self.template_dir.glob("*.json")):
            try:
                data = self._read_template_file(file_path)
                templates.append({
                    "name": data["name"],
                    "description": data["description"],
                    "file": str(file_path),
                    "status": "ok",
                })
            except TemplateStorageError as error:
                log_exception(
                    self.logger,
                    "模板索引扫描发现损坏文件",
                    error,
                    template=str(file_path),
                )
                templates.append({
                    "name": file_path.stem,
                    "description": "模板文件损坏，无法读取。",
                    "file": str(file_path),
                    "status": "corrupted",
                })

        return templates

    def _ensure_builtin_templates(self) -> None:
        """确保内置模板已写入用户目录。仅在文件不存在时写入。"""
        builtins = [("默认样式", DEFAULT_STYLES, "ToDOCX 内置默认样式")]
        for name, styles, description in builtins:
            file_path = self._get_template_path(name)
            if not file_path.exists():
                try:
                    payload = self._build_template_payload(name, styles, description)
                    self._validate_template_payload(payload, expected_name=name)
                    atomic_write_json(file_path, payload)
                    log_event(self.logger, "内置模板已落盘", template=str(file_path), name=name)
                except (TemplateStorageError, OSError) as error:
                    log_exception(self.logger, "内置模板落盘失败", error, name=name)

    def get_builtin_templates(self) -> Dict[str, Dict[str, Any]]:
        """获取内置模板（从用户目录读取已落盘的文件）"""
        result: Dict[str, Dict[str, Any]] = {}
        for name, styles, _ in [("默认样式", DEFAULT_STYLES, "ToDOCX 内置默认样式")]:
            file_path = self._get_template_path(name)
            if file_path.exists():
                try:
                    data = self._read_template_file(file_path)
                    result[name] = data.get("styles", styles)
                    continue
                except TemplateStorageError:
                    pass
            result[name] = styles  # fallback 到代码常量
        return result
    
    def rename_template(self, old_name: str, new_name: str) -> bool:
        """重命名模板

        Args:
            old_name: 原模板名称
            new_name: 新模板名称

        Returns:
            是否重命名成功
        """
        file_path = self._get_template_path(old_name)
        if not file_path.exists():
            return False

        data = self._read_template_file(file_path)

        new_path = self._get_template_path(new_name)
        if new_path.exists() and new_path != file_path:
            raise TemplateStorageError(
                "同名模板已存在。",
                code="TODX308",
                hint="请换一个模板名称，或先删除已有同名模板。",
            )

        self.save_template(
            new_name,
            data.get("styles", {}),
            data.get("description", ""),
        )

        if new_path != file_path:
            self.delete_template(old_name)

        log_event(
            self.logger,
            "模板已重命名",
            old_name=old_name,
            new_name=new_name,
        )
        return True
