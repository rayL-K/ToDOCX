"""模板管理模块"""

import json
from pathlib import Path
from typing import Dict, Any, List, Optional

from .config import DEFAULT_STYLES


class TemplateManager:
    """样式模板管理器"""
    
    def __init__(self, template_dir: str = None):
        if template_dir is None:
            # 默认模板目录
            self.template_dir = Path(__file__).parent.parent / "templates"
        else:
            self.template_dir = Path(template_dir)
        
        # 确保目录存在
        self.template_dir.mkdir(parents=True, exist_ok=True)

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
    
    def save_template(self, name: str, styles: Dict[str, Any], description: str = "") -> str:
        """保存模板
        
        Args:
            name: 模板名称
            styles: 样式配置
            description: 模板描述
            
        Returns:
            模板文件路径
        """
        template_data = {
            "name": name,
            "description": description,
            "styles": styles
        }

        file_path = self._get_template_path(name)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=2)
        
        return str(file_path)

    def load_template_data(self, name: str) -> Optional[Dict[str, Any]]:
        """加载完整模板信息。"""
        try:
            file_path = self._get_template_path(name)
        except ValueError:
            return None

        if not file_path.exists():
            return None

        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def load_template(self, name: str) -> Optional[Dict[str, Any]]:
        """加载模板
        
        Args:
            name: 模板名称
            
        Returns:
            模板样式配置，如果不存在返回None
        """
        template_data = self.load_template_data(name)
        if not template_data:
            return None

        return template_data.get("styles", {})
    
    def delete_template(self, name: str) -> bool:
        """删除模板
        
        Args:
            name: 模板名称
            
        Returns:
            是否删除成功
        """
        try:
            file_path = self._get_template_path(name)
        except ValueError:
            return False

        if file_path.exists():
            file_path.unlink()
            return True
        return False
    
    def list_templates(self) -> List[Dict[str, str]]:
        """列出所有模板
        
        Returns:
            模板列表，每个元素包含name和description
        """
        templates = []
        
        for file_path in sorted(self.template_dir.glob("*.json")):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    templates.append({
                        "name": data.get("name", file_path.stem),
                        "description": data.get("description", ""),
                        "file": str(file_path)
                    })
            except:
                continue
        
        return templates
    
    def get_builtin_templates(self) -> Dict[str, Dict[str, Any]]:
        """获取内置模板"""
        return {
            "默认样式": DEFAULT_STYLES,
        }
    
    def rename_template(self, old_name: str, new_name: str) -> bool:
        """重命名模板
        
        Args:
            old_name: 原模板名称
            new_name: 新模板名称
            
        Returns:
            是否重命名成功
        """
        template_data = self.load_template_data(old_name)
        if template_data is None:
            return False
        
        self.save_template(
            new_name,
            template_data.get("styles", {}),
            template_data.get("description", ""),
        )
        
        self.delete_template(old_name)
        
        return True
