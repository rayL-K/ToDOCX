"""DOCX 转 Markdown 模块"""

import os
import re
from pathlib import Path
import mammoth
from docx import Document
from bs4 import BeautifulSoup


class DocxToMarkdown:
    """DOCX 转 Markdown 转换器"""
    
    def __init__(self):
        self.supported_extensions = ['.docx', '.doc']
    
    def convert_to_markdown(self, input_path: str, output_path: str = None,
                           progress_callback=None) -> str:
        """将DOCX转换为Markdown
        
        Args:
            input_path: DOCX文件路径
            output_path: 输出Markdown文件路径（可选）
            progress_callback: 进度回调函数
            
        Returns:
            Markdown内容或输出文件路径
        """
        input_path = Path(input_path)
        if not input_path.exists():
            raise FileNotFoundError(f"文件不存在: {input_path}")
        
        if input_path.suffix.lower() not in self.supported_extensions:
            raise ValueError(f"不支持的文件格式: {input_path.suffix}")
        
        if progress_callback:
            progress_callback(10, "读取Word文档...")
        
        # 使用mammoth进行转换
        with open(input_path, 'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
        
        if progress_callback:
            progress_callback(50, "转换为Markdown...")
        
        # 将HTML转换为Markdown
        md_content = self._html_to_markdown(html)
        
        if progress_callback:
            progress_callback(90, "处理完成...")
        
        if output_path:
            output_path = Path(output_path)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(md_content)
            
            if progress_callback:
                progress_callback(100, "保存完成")
            
            return str(output_path)
        
        if progress_callback:
            progress_callback(100, "转换完成")
        
        return md_content
    
    def _html_to_markdown(self, html: str) -> str:
        """将HTML转换为Markdown"""
        soup = BeautifulSoup(html, 'lxml')
        
        # 处理各种元素
        md_lines = []
        
        for element in soup.body.children if soup.body else soup.children:
            md = self._process_element(element)
            if md:
                md_lines.append(md)
        
        return '\n\n'.join(md_lines)
    
    def _process_element(self, element, depth=0) -> str:
        """处理单个HTML元素"""
        if element.name is None:
            text = str(element).strip()
            return text if text else ''
        
        if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = int(element.name[1])
            text = element.get_text().strip()
            return f"{'#' * level} {text}"
        
        elif element.name == 'p':
            content = self._process_inline(element)
            return content.strip() if content.strip() else ''
        
        elif element.name == 'ul':
            items = []
            for li in element.find_all('li', recursive=False):
                text = self._process_inline(li)
                items.append(f"- {text}")
            return '\n'.join(items)
        
        elif element.name == 'ol':
            items = []
            for i, li in enumerate(element.find_all('li', recursive=False), 1):
                text = self._process_inline(li)
                items.append(f"{i}. {text}")
            return '\n'.join(items)
        
        elif element.name == 'blockquote':
            text = element.get_text().strip()
            lines = text.split('\n')
            return '\n'.join([f"> {line}" for line in lines])
        
        elif element.name in ['pre', 'code']:
            code = element.get_text()
            return f"```\n{code}\n```"
        
        elif element.name == 'table':
            return self._process_table(element)
        
        elif element.name == 'img':
            src = element.get('src', '')
            alt = element.get('alt', '')
            return f"![{alt}]({src})"
        
        elif element.name == 'hr':
            return '---'
        
        elif element.name in ['div', 'section', 'article', 'span']:
            results = []
            for child in element.children:
                result = self._process_element(child, depth + 1)
                if result:
                    results.append(result)
            return '\n\n'.join(results)
        
        else:
            return element.get_text().strip()
    
    def _process_inline(self, element) -> str:
        """处理行内元素"""
        if element.name is None:
            return str(element)
        
        result = []
        for child in element.children:
            if child.name is None:
                result.append(str(child))
            elif child.name == 'strong' or child.name == 'b':
                text = child.get_text()
                result.append(f"**{text}**")
            elif child.name == 'em' or child.name == 'i':
                text = child.get_text()
                result.append(f"*{text}*")
            elif child.name == 'code':
                text = child.get_text()
                result.append(f"`{text}`")
            elif child.name == 'a':
                text = child.get_text()
                href = child.get('href', '')
                result.append(f"[{text}]({href})")
            elif child.name == 'img':
                src = child.get('src', '')
                alt = child.get('alt', '')
                result.append(f"![{alt}]({src})")
            elif child.name == 'br':
                result.append('\n')
            else:
                result.append(self._process_inline(child))
        
        return ''.join(result)
    
    def _process_table(self, table) -> str:
        """处理表格"""
        rows = table.find_all('tr')
        if not rows:
            return ''
        
        md_rows = []
        
        for i, row in enumerate(rows):
            cells = row.find_all(['th', 'td'])
            cell_texts = [cell.get_text().strip().replace('|', '\\|') for cell in cells]
            md_rows.append('| ' + ' | '.join(cell_texts) + ' |')
            
            # 添加表头分隔行
            if i == 0:
                separator = '| ' + ' | '.join(['---'] * len(cells)) + ' |'
                md_rows.append(separator)
        
        return '\n'.join(md_rows)
