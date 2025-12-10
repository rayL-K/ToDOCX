"""模板管理模块"""

import json
import os
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
        
        # 清理文件名
        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_name = safe_name.replace(' ', '_')
        
        file_path = self.template_dir / f"{safe_name}.json"
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=2)
        
        return str(file_path)
    
    def load_template(self, name: str) -> Optional[Dict[str, Any]]:
        """加载模板
        
        Args:
            name: 模板名称
            
        Returns:
            模板样式配置，如果不存在返回None
        """
        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_name = safe_name.replace(' ', '_')
        
        file_path = self.template_dir / f"{safe_name}.json"
        
        if not file_path.exists():
            return None
        
        with open(file_path, 'r', encoding='utf-8') as f:
            template_data = json.load(f)
        
        return template_data.get("styles", {})
    
    def delete_template(self, name: str) -> bool:
        """删除模板
        
        Args:
            name: 模板名称
            
        Returns:
            是否删除成功
        """
        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_name = safe_name.replace(' ', '_')
        
        file_path = self.template_dir / f"{safe_name}.json"
        
        if file_path.exists():
            os.remove(file_path)
            return True
        return False
    
    def list_templates(self) -> List[Dict[str, str]]:
        """列出所有模板
        
        Returns:
            模板列表，每个元素包含name和description
        """
        templates = []
        
        for file_path in self.template_dir.glob("*.json"):
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
        # 加载原模板
        styles = self.load_template(old_name)
        if styles is None:
            return False
        
        # 保存为新名称
        self.save_template(new_name, styles)
        
        # 删除原模板
        self.delete_template(old_name)
        
        return True
