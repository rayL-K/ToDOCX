"""Markdown 转 DOCX 转换模块"""

import re
import os
import base64
from pathlib import Path
from io import BytesIO
import markdown
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image
import httpx

from .config import DEFAULT_STYLES, get_font_size_pt, FONT_SIZE_MAP


class MarkdownConverter:
    """Markdown 转换器"""
    
    def __init__(self, styles: dict = None):
        self.styles = styles or DEFAULT_STYLES
        self.supported_extensions = ['.md', '.markdown']
        self.image_cache = {}
    
    def convert_to_docx(self, input_path: str, output_path: str = None,
                        progress_callback=None, styles: dict = None) -> str:
        """将Markdown转换为DOCX
        
        Args:
            input_path: Markdown文件路径
            output_path: 输出DOCX文件路径（可选）
            progress_callback: 进度回调函数
            styles: 自定义样式（可选）
            
        Returns:
            输出文件路径
        """
        input_path = Path(input_path)
        if not input_path.exists():
            raise FileNotFoundError(f"文件不存在: {input_path}")
        
        if input_path.suffix.lower() not in self.supported_extensions:
            raise ValueError(f"不支持的文件格式: {input_path.suffix}")
        
        if output_path is None:
            output_path = input_path.with_suffix('.docx')
        else:
            output_path = Path(output_path)
        
        if styles:
            self.styles = {**DEFAULT_STYLES, **styles}
        
        # 读取Markdown内容
        with open(input_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
        
        # 获取Markdown文件所在目录，用于解析相对路径图片
        self.base_dir = input_path.parent
        
        if progress_callback:
            progress_callback(10, "解析Markdown内容...")
        
        # 转换为DOCX
        doc = self._md_to_docx(md_content, progress_callback)
        
        if progress_callback:
            progress_callback(90, "保存文档...")
        
        doc.save(str(output_path))
        
        if progress_callback:
            progress_callback(100, "转换完成")
        
        return str(output_path)
    
    def convert_from_string(self, md_content: str, output_path: str,
                            progress_callback=None, styles: dict = None,
                            base_dir: str = None) -> str:
        """从字符串转换Markdown到DOCX
        
        Args:
            md_content: Markdown内容字符串
            output_path: 输出DOCX文件路径
            progress_callback: 进度回调函数
            styles: 自定义样式
            base_dir: 图片基础目录
            
        Returns:
            输出文件路径
        """
        if styles:
            self.styles = {**DEFAULT_STYLES, **styles}
        
        self.base_dir = Path(base_dir) if base_dir else Path.cwd()
        
        if progress_callback:
            progress_callback(10, "解析Markdown内容...")
        
        doc = self._md_to_docx(md_content, progress_callback)
        
        if progress_callback:
            progress_callback(90, "保存文档...")
        
        doc.save(str(output_path))
        
        if progress_callback:
            progress_callback(100, "转换完成")
        
        return str(output_path)
    
    def _md_to_docx(self, md_content: str, progress_callback=None) -> Document:
        """将Markdown内容转换为Document对象"""
        doc = Document()
        
        # 设置默认样式
        self._setup_styles(doc)
        
        # 预处理：提取并保护代码块和公式
        code_blocks = []
        formulas = []
        
        # 保护代码块
        def save_code_block(match):
            code_blocks.append(match.group(0))
            return f"<<<CODE_BLOCK_{len(code_blocks) - 1}>>>"
        
        md_content = re.sub(r'```[\s\S]*?```', save_code_block, md_content)
        
        # 保护行内代码
        inline_codes = []
        def save_inline_code(match):
            inline_codes.append(match.group(1))
            return f"<<<INLINE_CODE_{len(inline_codes) - 1}>>>"
        
        md_content = re.sub(r'`([^`]+)`', save_inline_code, md_content)
        
        # 保护公式块
        def save_formula_block(match):
            formulas.append(match.group(0))
            return f"<<<FORMULA_BLOCK_{len(formulas) - 1}>>>"
        
        md_content = re.sub(r'\$\$[\s\S]*?\$\$', save_formula_block, md_content)
        
        # 保护行内公式
        inline_formulas = []
        def save_inline_formula(match):
            inline_formulas.append(match.group(1))
            return f"<<<INLINE_FORMULA_{len(inline_formulas) - 1}>>>"
        
        md_content = re.sub(r'\$([^\$]+)\$', save_inline_formula, md_content)
        
        # 转换为HTML（用于解析复杂结构）
        html = markdown.markdown(
            md_content,
            extensions=['tables', 'fenced_code', 'toc', 'nl2br']
        )
        
        soup = BeautifulSoup(html, 'lxml')
        
        if progress_callback:
            progress_callback(30, "转换文档结构...")
        
        # 处理每个元素
        total_elements = len(soup.find_all(True))
        processed = 0
        
        for element in soup.body.children if soup.body else soup.children:
            self._process_element(doc, element, code_blocks, inline_codes, 
                                 formulas, inline_formulas)
            processed += 1
            if progress_callback and total_elements > 0:
                progress = 30 + int(60 * processed / total_elements)
                progress_callback(min(progress, 90), "转换文档内容...")
        
        return doc
    
    def _get_font_size(self, style_config):
        """获取字体大小（磅值）"""
        size = style_config.get('font_size', 12)
        if isinstance(size, str):
            # 标准字号名称，如"小四"
            return get_font_size_pt(size)
        return size
    
    def _apply_line_spacing(self, paragraph_format, style_config):
        """应用行间距设置"""
        spacing_type = style_config.get('line_spacing_type', '1.5倍行距')
        spacing_value = style_config.get('line_spacing_value', 1.5)
        
        # 确保spacing_value是数值
        if isinstance(spacing_value, str):
            try:
                spacing_value = float(spacing_value)
            except:
                spacing_value = 20 if spacing_type == '固定值' else 1.5
        
        if spacing_type == '固定值':
            # 固定行距（磅值）
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            paragraph_format.line_spacing = Pt(float(spacing_value))
        else:
            # 倍数行距
            paragraph_format.line_spacing = float(spacing_value)
    
    def _setup_styles(self, doc: Document):
        """设置文档样式"""
        styles = doc.styles
        
        # 设置正文样式
        try:
            normal_style = styles['Normal']
            normal_font = normal_style.font
            body_style = self.styles.get('body', {})
            
            # 西文字体
            font_en = body_style.get('font_name_en', body_style.get('font_name', 'Times New Roman'))
            normal_font.name = font_en
            
            # 字号
            font_size = self._get_font_size(body_style)
            normal_font.size = Pt(font_size)
            
            # 中文字体
            font_cn = body_style.get('font_name_cn', body_style.get('font_name', '宋体'))
            normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
        except:
            pass
        
        # 创建各级标题样式
        for i in range(1, 5):
            style_name = f'Heading {i}'
            heading_key = f'heading{i}'
            
            if heading_key in self.styles:
                try:
                    style = styles[style_name]
                    font = style.font
                    heading_style = self.styles[heading_key]
                    
                    # 西文字体
                    font_en = heading_style.get('font_name_en', heading_style.get('font_name', 'Times New Roman'))
                    font.name = font_en
                    
                    # 字号
                    font_size = self._get_font_size(heading_style)
                    font.size = Pt(font_size)
                    
                    font.bold = heading_style.get('bold', True)
                    
                    # 中文字体
                    font_cn = heading_style.get('font_name_cn', heading_style.get('font_name', '宋体'))
                    style._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
                except:
                    pass
    
    def _process_element(self, doc, element, code_blocks, inline_codes, 
                        formulas, inline_formulas):
        """处理单个HTML元素"""
        if element.name is None:
            # 纯文本
            text = str(element).strip()
            if text:
                # 恢复特殊内容
                text = self._restore_special_content(
                    text, code_blocks, inline_codes, formulas, inline_formulas
                )
                if text.strip():
                    p = doc.add_paragraph(text)
                    self._apply_body_style(p)
            return
        
        if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = int(element.name[1])
            text = element.get_text()
            text = self._restore_special_content(
                text, code_blocks, inline_codes, formulas, inline_formulas
            )
            # 使用普通段落而非add_heading，以便完全控制样式
            heading = doc.add_paragraph()
            run = heading.add_run(text)
            self._apply_heading_style(heading, level)
            
        elif element.name == 'p':
            text = element.get_text()
            
            # 检查是否是特殊内容
            if '<<<CODE_BLOCK_' in text:
                match = re.search(r'<<<CODE_BLOCK_(\d+)>>>', text)
                if match:
                    idx = int(match.group(1))
                    self._add_code_block(doc, code_blocks[idx])
                    return
            
            if '<<<FORMULA_BLOCK_' in text:
                match = re.search(r'<<<FORMULA_BLOCK_(\d+)>>>', text)
                if match:
                    idx = int(match.group(1))
                    self._add_formula(doc, formulas[idx])
                    return
            
            # 检查是否包含图片
            img = element.find('img')
            if img:
                self._add_image(doc, img.get('src', ''), img.get('alt', ''))
                return
            
            text = self._restore_special_content(
                text, code_blocks, inline_codes, formulas, inline_formulas
            )
            if text.strip():
                p = doc.add_paragraph(text)
                self._apply_body_style(p)
                
        elif element.name == 'ul':
            for li in element.find_all('li', recursive=False):
                text = li.get_text()
                text = self._restore_special_content(
                    text, code_blocks, inline_codes, formulas, inline_formulas
                )
                p = doc.add_paragraph(text, style='List Bullet')
                
        elif element.name == 'ol':
            for li in element.find_all('li', recursive=False):
                text = li.get_text()
                text = self._restore_special_content(
                    text, code_blocks, inline_codes, formulas, inline_formulas
                )
                p = doc.add_paragraph(text, style='List Number')
                
        elif element.name == 'blockquote':
            text = element.get_text()
            text = self._restore_special_content(
                text, code_blocks, inline_codes, formulas, inline_formulas
            )
            p = doc.add_paragraph(text)
            self._apply_quote_style(p)
            
        elif element.name == 'pre':
            code = element.find('code')
            if code:
                self._add_code_block(doc, code.get_text())
            else:
                self._add_code_block(doc, element.get_text())
                
        elif element.name == 'table':
            self._add_table(doc, element)
            
        elif element.name == 'img':
            self._add_image(doc, element.get('src', ''), element.get('alt', ''))
            
        elif element.name == 'hr':
            # 添加分隔线
            p = doc.add_paragraph()
            p.add_run('─' * 50)
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        elif element.name in ['div', 'section', 'article']:
            # 递归处理容器元素
            for child in element.children:
                self._process_element(doc, child, code_blocks, inline_codes,
                                     formulas, inline_formulas)
    
    def _restore_special_content(self, text, code_blocks, inline_codes, 
                                 formulas, inline_formulas):
        """恢复特殊内容（代码、公式）"""
        # 恢复代码块
        for i, code in enumerate(code_blocks):
            text = text.replace(f'<<<CODE_BLOCK_{i}>>>', '')
        
        # 恢复行内代码
        for i, code in enumerate(inline_codes):
            text = text.replace(f'<<<INLINE_CODE_{i}>>>', f'「{code}」')
        
        # 恢复公式块
        for i, formula in enumerate(formulas):
            text = text.replace(f'<<<FORMULA_BLOCK_{i}>>>', '')
        
        # 恢复行内公式
        for i, formula in enumerate(inline_formulas):
            text = text.replace(f'<<<INLINE_FORMULA_{i}>>>', f'[公式: {formula}]')
        
        return text
    
    def _apply_body_style(self, paragraph):
        """应用正文样式"""
        style = self.styles.get('body', {})
        pf = paragraph.paragraph_format
        
        # 行距
        self._apply_line_spacing(pf, style)
        
        # 段前段后间距
        pf.space_before = Pt(style.get('space_before', 0))
        pf.space_after = Pt(style.get('space_after', 0))
        
        # 首行缩进
        indent = style.get('first_line_indent', 2)
        if indent > 0:
            font_size = self._get_font_size(style)
            pf.first_line_indent = Pt(font_size * indent)
        
        # 对齐方式
        alignment = style.get('alignment', 'left')
        if alignment == 'justify':
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif alignment == 'center':
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alignment == 'right':
            pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # 字体设置
        font_cn = style.get('font_name_cn', style.get('font_name', '宋体'))
        font_en = style.get('font_name_en', style.get('font_name', 'Times New Roman'))
        font_size = self._get_font_size(style)
        
        for run in paragraph.runs:
            run.font.name = font_en
            run.font.size = Pt(font_size)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
    
    def _apply_heading_style(self, heading, level):
        """应用标题样式"""
        style_key = f'heading{min(level, 4)}'
        style = self.styles.get(style_key, {})
        
        pf = heading.paragraph_format
        pf.space_before = Pt(style.get('space_before', 12))
        pf.space_after = Pt(style.get('space_after', 6))
        self._apply_line_spacing(pf, style)
        
        # 对齐方式
        alignment = style.get('alignment', 'left')
        if alignment == 'center':
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        font_cn = style.get('font_name_cn', style.get('font_name', '宋体'))
        font_en = style.get('font_name_en', style.get('font_name', 'Times New Roman'))
        font_size = self._get_font_size(style)
        
        for run in heading.runs:
            run.font.name = font_en
            run.font.size = Pt(font_size)
            run.font.bold = style.get('bold', True)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
    
    def _apply_quote_style(self, paragraph):
        """应用引用样式"""
        style = self.styles.get('quote', {})
        pf = paragraph.paragraph_format
        
        pf.left_indent = Cm(style.get('left_indent', 1))
        pf.space_before = Pt(style.get('space_before', 6))
        pf.space_after = Pt(style.get('space_after', 6))
        pf.line_spacing = style.get('line_spacing', 1.5)
        
        font_size = self._get_font_size(style) if style.get('font_size') else 11
        font_name = style.get('font_name_cn', style.get('font_name', '楷体'))
        
        for run in paragraph.runs:
            run.font.name = style.get('font_name_en', font_name)
            run.font.size = Pt(font_size)
            run.font.italic = style.get('italic', True)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            
            color = style.get('color', '#666666')
            if color.startswith('#'):
                r = int(color[1:3], 16)
                g = int(color[3:5], 16)
                b = int(color[5:7], 16)
                run.font.color.rgb = RGBColor(r, g, b)
    
    def _add_code_block(self, doc, code_text):
        """添加代码块"""
        style = self.styles.get('code', {})
        
        # 清理代码文本
        if code_text.startswith('```'):
            lines = code_text.split('\n')
            # 移除首尾的 ```
            lines = lines[1:-1] if lines[-1].strip() == '```' else lines[1:]
            code_text = '\n'.join(lines)
        
        # 创建代码段落
        font_size = self._get_font_size(style) if style.get('font_size') else 10
        font_name = style.get('font_name_en', style.get('font_name', 'Consolas'))
        
        for line in code_text.split('\n'):
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Consolas')
            
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = style.get('line_spacing', 1.2)
            
            # 添加背景色（通过底纹）
            self._add_shading(p, style.get('background', '#f5f5f5'))
    
    def _add_shading(self, paragraph, color):
        """为段落添加底纹"""
        if color.startswith('#'):
            color = color[1:]
        
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), color)
        paragraph._p.get_or_add_pPr().append(shading)
    
    def _add_formula(self, doc, formula_text):
        """添加公式"""
        style = self.styles.get('formula', {})
        
        # 清理公式文本
        formula_text = formula_text.strip()
        if formula_text.startswith('$$'):
            formula_text = formula_text[2:]
        if formula_text.endswith('$$'):
            formula_text = formula_text[:-2]
        formula_text = formula_text.strip()
        
        p = doc.add_paragraph()
        run = p.add_run(f'[公式: {formula_text}]')
        run.font.name = 'Cambria Math'
        run.font.size = Pt(12)
        
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(style.get('space_before', 6))
        pf.space_after = Pt(style.get('space_after', 6))
    
    def _add_image(self, doc, src, alt=''):
        """添加图片"""
        style = self.styles.get('image', {})
        
        try:
            # 判断图片来源
            if src.startswith('data:image'):
                # Base64 图片
                image_data = self._decode_base64_image(src)
            elif src.startswith(('http://', 'https://')):
                # 网络图片
                image_data = self._download_image(src)
            else:
                # 本地图片
                img_path = self.base_dir / src if hasattr(self, 'base_dir') else Path(src)
                if not img_path.exists():
                    p = doc.add_paragraph(f'[图片: {alt or src}]')
                    pf = p.paragraph_format
                    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    return
                image_data = str(img_path)
            
            # 添加图片
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(style.get('space_before', 6))
            pf.space_after = Pt(style.get('space_after', 6))
            
            run = p.add_run()
            
            if isinstance(image_data, str):
                # 文件路径
                run.add_picture(image_data, width=Cm(style.get('max_width', 15)))
            else:
                # BytesIO
                run.add_picture(image_data, width=Cm(style.get('max_width', 15)))
            
            # 添加图片说明（使用caption样式）
            if alt:
                caption_p = doc.add_paragraph(alt)
                caption_style = self.styles.get('caption', {})
                
                caption_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self._apply_line_spacing(caption_p.paragraph_format, caption_style)
                
                font_cn = caption_style.get('font_name_cn', '黑体')
                font_en = caption_style.get('font_name_en', 'Times New Roman')
                font_size = self._get_font_size(caption_style) if caption_style.get('font_size') else 9
                
                for run in caption_p.runs:
                    run.font.name = font_en
                    run.font.size = Pt(font_size)
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
                    
        except Exception as e:
            # 图片加载失败，显示占位符
            p = doc.add_paragraph(f'[图片加载失败: {alt or src}]')
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def _decode_base64_image(self, data_url):
        """解码Base64图片"""
        # 提取base64数据
        if ',' in data_url:
            data_url = data_url.split(',')[1]
        
        image_data = base64.b64decode(data_url)
        return BytesIO(image_data)
    
    def _download_image(self, url):
        """下载网络图片"""
        if url in self.image_cache:
            return BytesIO(self.image_cache[url])
        
        try:
            with httpx.Client(timeout=30.0) as client:
                response = client.get(url)
                response.raise_for_status()
                self.image_cache[url] = response.content
                return BytesIO(response.content)
        except:
            raise Exception(f"无法下载图片: {url}")
    
    def _add_table(self, doc, table_element):
        """添加表格"""
        style = self.styles.get('table', {})
        
        rows = table_element.find_all('tr')
        if not rows:
            return
        
        # 计算列数
        first_row = rows[0]
        cols = len(first_row.find_all(['th', 'td']))
        
        if cols == 0:
            return
        
        # 创建表格
        table = doc.add_table(rows=len(rows), cols=cols)
        table.style = 'Table Grid'
        
        for i, row in enumerate(rows):
            cells = row.find_all(['th', 'td'])
            for j, cell in enumerate(cells):
                if j < cols:
                    table_cell = table.rows[i].cells[j]
                    table_cell.text = cell.get_text().strip()
                    
                    # 设置字体
                    table_font_size = self._get_font_size(style) if style.get('font_size') else 10
                    table_font_cn = style.get('font_name_cn', style.get('font_name', '宋体'))
                    table_font_en = style.get('font_name_en', table_font_cn)
                    
                    for p in table_cell.paragraphs:
                        for run in p.runs:
                            run.font.name = table_font_en
                            run.font.size = Pt(table_font_size)
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), table_font_cn)
                            
                            # 表头加粗
                            if cell.name == 'th' and style.get('header_bold', True):
                                run.font.bold = True
