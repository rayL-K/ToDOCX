"""LaTeX 文档分析模块 - 用于预览和格式识别"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass


@dataclass
class LatexParagraphInfo:
    """LaTeX 段落信息"""
    index: int
    text: str  # 显示文本（去掉命令）
    raw_text: str  # 原始文本（包含命令）
    line_start: int  # 起始行号
    line_end: int  # 结束行号
    element_type: str  # heading1, heading2, body, caption, code, quote 等
    original_type: str  # 原始识别的类型


@dataclass
class LatexFormatGroup:
    """LaTeX 格式分组"""
    signature: str
    element_type: str
    original_type: str
    paragraph_indices: List[int]
    sample_text: str


class LatexAnalyzer:
    """LaTeX 文档分析器"""
    
    # LaTeX 标题命令映射
    HEADING_COMMANDS = {
        r'\chapter': 'heading1',
        r'\section': 'heading1',
        r'\subsection': 'heading2',
        r'\subsubsection': 'heading3',
        r'\paragraph': 'heading4',
        r'\subparagraph': 'heading4',
    }
    
    # 特殊环境映射（需要显示的内容环境）
    CONTENT_ENVIRONMENTS = {
        'abstract': 'quote',
        'quote': 'quote',
        'quotation': 'quote',
    }
    
    # 图片环境（提取 caption）
    FIGURE_ENVIRONMENTS = {'figure'}
    
    # 表格环境（完整保留内容）
    TABLE_ENVIRONMENTS = {'table'}
    
    # 代码环境（保留完整内容和格式）
    CODE_ENVIRONMENTS = {'verbatim', 'lstlisting', 'minted', 'listing'}
    
    # 列表环境（提取 \item 内容）
    LIST_ENVIRONMENTS = {'itemize', 'enumerate', 'description'}
    
    # 公式环境（保留内容用于渲染）
    FORMULA_ENVIRONMENTS = {'equation', 'equation*', 'align', 'align*', 'gather', 'gather*', 'multline', 'multline*'}
    
    # 完全跳过的环境（包括其内容）
    SKIP_ENVIRONMENTS = {
        'tikzpicture', 'pgfpicture',
        'titlepage',  # 封面页
    }
    
    # 透明环境（跳过环境标记，但处理内部内容）
    # tabular 作为 table 的子环境，应该透明处理
    TRANSPARENT_ENVIRONMENTS = {'center', 'flushleft', 'flushright', 'minipage', 'tabular', 'tabular*', 'longtable'}
    
    # 纯命令行（应该跳过的命令）
    SKIP_COMMANDS = {
        r'\setlength', r'\newcommand', r'\renewcommand', r'\newenvironment',
        r'\def', r'\let', r'\makeatletter', r'\makeatother',
        r'\pagestyle', r'\thispagestyle', r'\pagenumbering',
        r'\setcounter', r'\addtocounter', r'\stepcounter',
        r'\bibliographystyle', r'\bibliography', r'\printbibliography',
        r'\tableofcontents', r'\listoffigures', r'\listoftables',
        r'\maketitle', r'\title', r'\author', r'\date',
        r'\usepackage', r'\RequirePackage', r'\documentclass',
        r'\input', r'\include', r'\includeonly',
        r'\newpage', r'\clearpage', r'\cleardoublepage',
        r'\vspace', r'\hspace', r'\vfill', r'\hfill',
        r'\centering', r'\raggedleft', r'\raggedright',
        r'\small', r'\large', r'\Large', r'\LARGE', r'\huge', r'\Huge',
        r'\normalsize', r'\footnotesize', r'\scriptsize', r'\tiny',
        r'\textwidth', r'\linewidth', r'\columnwidth',
        r'\label', r'\ref', r'\pageref', r'\eqref',
        r'\nocite', r'\phantom', r'\hphantom', r'\vphantom',
    }
    
    def __init__(self):
        self.paragraphs: List[LatexParagraphInfo] = []
        self.format_groups: Dict[str, LatexFormatGroup] = {}
        self.file_path: Optional[Path] = None
        self.raw_content: str = ""
        self.lines: List[str] = []
    
    def load_document(self, file_path: str) -> bool:
        """加载 LaTeX 文档"""
        try:
            self.file_path = Path(file_path)
            with open(file_path, 'r', encoding='utf-8') as f:
                self.raw_content = f.read()
            self.lines = self.raw_content.split('\n')
            self._analyze_structure()
            self._group_by_type()
            return True
        except Exception as e:
            print(f"加载LaTeX文档失败: {e}")
            return False
    
    def _analyze_structure(self):
        """分析 LaTeX 文档结构"""
        self.paragraphs = []
        
        in_document = False
        current_para_lines = []
        current_start_line = 0
        para_index = 0
        
        # 环境栈：[(env_name, start_line, content_lines)]
        env_stack = []
        skip_depth = 0  # 跳过环境的嵌套深度
        
        for i, line in enumerate(self.lines):
            stripped = line.strip()
            
            # 检查 \begin{document}
            if r'\begin{document}' in stripped:
                in_document = True
                continue
            
            # 检查 \end{document}
            if r'\end{document}' in stripped:
                break
            
            if not in_document:
                continue
            
            # 跳过注释
            if stripped.startswith('%'):
                continue
            
            # 检查环境开始
            env_begin = re.match(r'\\begin\{(\w+\*?)\}', stripped)
            if env_begin:
                env_name = env_begin.group(1)
                
                # 完全跳过的环境
                if env_name in self.SKIP_ENVIRONMENTS:
                    skip_depth += 1
                    continue
                
                # 透明环境：收集到父环境（如果有），否则忽略
                if env_name in self.TRANSPARENT_ENVIRONMENTS:
                    if env_stack:
                        env_stack[-1][2].append((i, line))
                    continue
                
                # 先保存之前的段落
                if current_para_lines and not env_stack:
                    para_index = self._add_paragraph(para_index, current_para_lines,
                                       current_start_line, i - 1)
                    current_para_lines = []
                
                # 压入环境栈（包括表格、代码、公式等需要完整保留的环境）
                env_stack.append((env_name, i, [(i, line)]))  # 包含 \begin 行
                continue
            
            # 检查环境结束
            env_end = re.match(r'\\end\{(\w+\*?)\}', stripped)
            if env_end:
                env_name = env_end.group(1)
                
                # 跳过环境结束
                if env_name in self.SKIP_ENVIRONMENTS:
                    skip_depth = max(0, skip_depth - 1)
                    continue
                
                # 透明环境结束：收集到父环境（如果有），否则忽略
                if env_name in self.TRANSPARENT_ENVIRONMENTS:
                    if env_stack:
                        env_stack[-1][2].append((i, line))
                    continue
                
                # 弹出环境栈并处理
                if env_stack and env_stack[-1][0] == env_name:
                    env_info = env_stack.pop()
                    env_info[2].append((i, line))  # 包含 \end 行
                    para_index = self._process_environment(
                        para_index, env_info[0], env_info[1], i, env_info[2]
                    )
                continue
            
            # 如果在跳过的环境内，继续跳过
            if skip_depth > 0:
                continue
            
            # 跳过空行（用于分隔段落）
            if not stripped:
                if current_para_lines and not env_stack:
                    para_index = self._add_paragraph(para_index, current_para_lines,
                                       current_start_line, i - 1)
                    current_para_lines = []
                continue
            
            # 检查是否是纯命令行（应该跳过）
            if self._is_skip_command(stripped):
                continue
            
            # 如果在环境内部，收集内容
            if env_stack:
                env_stack[-1][2].append((i, line))
                continue
            
            # 检查标题命令
            heading_match = None
            for cmd in self.HEADING_COMMANDS:
                if cmd in stripped:
                    heading_match = cmd
                    break
            
            if heading_match:
                # 先保存之前的段落
                if current_para_lines:
                    para_index = self._add_paragraph(para_index, current_para_lines,
                                       current_start_line, i - 1)
                    current_para_lines = []
                
                # 添加标题
                para_index = self._add_heading_paragraph(para_index, line, i, heading_match)
                continue
            
            # 普通文本行 - 检查是否有实际内容
            if self._has_text_content(stripped):
                if not current_para_lines:
                    current_start_line = i
                current_para_lines.append(line)
        
        # 处理最后一个段落
        if current_para_lines:
            self._add_paragraph(para_index, current_para_lines,
                               current_start_line, len(self.lines) - 1)
    
    def _process_environment(self, para_index: int, env_name: str, 
                            start_line: int, end_line: int, 
                            content: List[Tuple[int, str]]) -> int:
        """处理环境内容，返回更新后的 para_index"""
        
        # 列表环境：提取每个 \item
        if env_name in self.LIST_ENVIRONMENTS:
            current_item_lines = []
            item_start = start_line
            
            for line_num, line in content:
                stripped = line.strip()
                if stripped.startswith(r'\item'):
                    # 保存之前的 item
                    if current_item_lines:
                        para_index = self._add_list_item(para_index, current_item_lines, item_start)
                    # 开始新 item
                    current_item_lines = [line]
                    item_start = line_num
                elif current_item_lines:
                    current_item_lines.append(line)
            
            # 最后一个 item
            if current_item_lines:
                para_index = self._add_list_item(para_index, current_item_lines, item_start)
            
            return para_index
        
        # 表格环境：完整保留
        if env_name in self.TABLE_ENVIRONMENTS:
            raw_text = '\n'.join([line for _, line in content])
            # 提取 caption
            caption_match = re.search(r'\\caption\{([^}]*)\}', raw_text)
            if caption_match:
                display_text = f"[表格] {caption_match.group(1)}"
            else:
                display_text = "[表格]"
            
            para = LatexParagraphInfo(
                index=para_index,
                text=display_text[:100],
                raw_text=raw_text,
                line_start=start_line,
                line_end=end_line,
                element_type='table',
                original_type='table'
            )
            self.paragraphs.append(para)
            return para_index + 1
        
        # 图片环境：提取 caption
        if env_name in self.FIGURE_ENVIRONMENTS:
            raw_text = '\n'.join([line for _, line in content])
            caption_match = re.search(r'\\caption\{([^}]*)\}', raw_text)
            if caption_match:
                display_text = f"[图片] {caption_match.group(1)}"
                para = LatexParagraphInfo(
                    index=para_index,
                    text=display_text[:100],
                    raw_text=raw_text,
                    line_start=start_line,
                    line_end=end_line,
                    element_type='caption',
                    original_type='caption'
                )
                self.paragraphs.append(para)
                return para_index + 1
            return para_index
        
        # 代码环境：完整保留格式
        if env_name in self.CODE_ENVIRONMENTS:
            raw_text = '\n'.join([line for _, line in content])
            # 提取 caption 如果有
            caption_match = re.search(r'caption=([^,\]]+)', raw_text)
            if caption_match:
                display_text = f"[代码] {caption_match.group(1)}"
            else:
                # 提取前两行代码作为预览
                code_lines = [line for _, line in content[1:-1] if line.strip()]
                preview = ' '.join(l.strip() for l in code_lines[:2])[:50]
                display_text = f"[代码] {preview}..."
            
            para = LatexParagraphInfo(
                index=para_index,
                text=display_text[:100],
                raw_text=raw_text,
                line_start=start_line,
                line_end=end_line,
                element_type='code',
                original_type='code'
            )
            self.paragraphs.append(para)
            return para_index + 1
        
        # 公式环境：完整保留
        if env_name in self.FORMULA_ENVIRONMENTS:
            raw_text = '\n'.join([line for _, line in content])
            # 提取公式内容作为预览
            formula_lines = [line.strip() for _, line in content[1:-1] if line.strip()]
            preview = ' '.join(formula_lines)[:50]
            display_text = f"[公式] {preview}..."
            
            para = LatexParagraphInfo(
                index=para_index,
                text=display_text[:100],
                raw_text=raw_text,
                line_start=start_line,
                line_end=end_line,
                element_type='formula',
                original_type='formula'
            )
            self.paragraphs.append(para)
            return para_index + 1
        
        # 引用环境
        if env_name in self.CONTENT_ENVIRONMENTS:
            raw_text = '\n'.join([line for _, line in content])
            display_text = self._clean_latex(raw_text)
            if display_text.strip():
                para = LatexParagraphInfo(
                    index=para_index,
                    text=display_text[:100],
                    raw_text=raw_text,
                    line_start=start_line,
                    line_end=end_line,
                    element_type='quote',
                    original_type='quote'
                )
                self.paragraphs.append(para)
                return para_index + 1
        
        return para_index
    
    def _add_list_item(self, para_index: int, lines: List[str], start_line: int) -> int:
        """添加列表项"""
        raw_text = '\n'.join(lines)
        # 清理 \item 命令
        display_text = re.sub(r'\\item\s*(\[[^\]]*\])?\s*', '', raw_text)
        display_text = self._clean_latex(display_text)
        
        if display_text.strip():
            para = LatexParagraphInfo(
                index=para_index,
                text=display_text[:100],
                raw_text=raw_text,
                line_start=start_line,
                line_end=start_line + len(lines) - 1,
                element_type='body',  # 列表项当作正文
                original_type='body'
            )
            self.paragraphs.append(para)
            return para_index + 1
        return para_index
    
    def _is_skip_command(self, line: str) -> bool:
        """检查是否是应该跳过的纯命令行"""
        # 检查是否以跳过的命令开头
        for cmd in self.SKIP_COMMANDS:
            if line.startswith(cmd):
                return True
        
        # 检查是否是只包含格式命令的行（没有实际文本）
        # 如 \centering, \large 等单独成行
        if re.match(r'^\\[a-zA-Z]+\s*$', line):
            return True
        
        # 检查是否是 \xxx{} 或 \xxx[] 形式但没有可见文本
        if re.match(r'^\\[a-zA-Z]+(\[[^\]]*\])?(\{[^}]*\})?\s*$', line):
            # 提取可能的文本内容
            text = re.sub(r'\\[a-zA-Z]+', '', line)
            text = re.sub(r'\[[^\]]*\]', '', text)
            text = re.sub(r'\{([^}]*)\}', r'\1', text)
            if not text.strip():
                return True
        
        return False
    
    def _has_text_content(self, line: str) -> bool:
        """检查行是否包含实际文本内容"""
        # 移除所有 LaTeX 命令
        text = re.sub(r'\\[a-zA-Z]+\*?(\[[^\]]*\])?(\{[^}]*\})?', '', line)
        text = re.sub(r'[{}$$]', '', text)
        text = text.strip()
        
        # 如果清理后还有内容，则认为有文本
        return len(text) > 0
    
    def _add_paragraph(self, index: int, lines: List[str], 
                       start_line: int, end_line: int) -> int:
        """添加普通段落，返回更新后的 para_index"""
        raw_text = '\n'.join(lines)
        display_text = self._clean_latex(raw_text)
        
        if not display_text.strip():
            return index
        
        para = LatexParagraphInfo(
            index=index,
            text=display_text[:100],
            raw_text=raw_text,
            line_start=start_line,
            line_end=end_line,
            element_type='body',
            original_type='body'
        )
        self.paragraphs.append(para)
        return index + 1
    
    def _add_heading_paragraph(self, index: int, line: str, 
                               line_num: int, cmd: str) -> int:
        """添加标题段落，返回更新后的 para_index"""
        # 提取标题文本
        match = re.search(r'\\(?:sub)*(?:section|chapter|paragraph)\*?\{([^}]*)\}', line)
        if match:
            display_text = match.group(1)
        else:
            display_text = self._clean_latex(line)
        
        element_type = self.HEADING_COMMANDS.get(cmd, 'heading1')
        
        para = LatexParagraphInfo(
            index=index,
            text=display_text[:100],
            raw_text=line,
            line_start=line_num,
            line_end=line_num,
            element_type=element_type,
            original_type=element_type
        )
        self.paragraphs.append(para)
        return index + 1
    
    def _add_environment_paragraph(self, index: int, lines: List[str],
                                   start_line: int, end_line: int, env_name: str) -> bool:
        """添加环境段落，返回是否成功添加"""
        raw_text = '\n'.join(lines)
        
        # 代码环境
        if env_name in self.CODE_ENVIRONMENTS:
            # 提取代码内容作为预览
            code_lines = lines[1:-1] if len(lines) > 2 else lines
            display_text = "[代码] " + ' '.join(l.strip() for l in code_lines[:2])[:60]
            element_type = 'code'
        # 内容环境（abstract, quote, figure, table等）
        elif env_name in self.CONTENT_ENVIRONMENTS:
            element_type = self.CONTENT_ENVIRONMENTS[env_name]
            display_text = self._clean_latex(raw_text)
            
            # 图表环境特殊处理 - 尝试提取 caption
            if env_name in ['figure', 'table']:
                caption_match = re.search(r'\\caption\{([^}]*)\}', raw_text)
                if caption_match:
                    display_text = f"[{env_name}] {caption_match.group(1)}"
        else:
            # 其他未知环境，尝试提取内容
            display_text = self._clean_latex(raw_text)
            element_type = 'body'
        
        # 如果没有有效内容，跳过
        if not display_text.strip():
            return False
        
        para = LatexParagraphInfo(
            index=index,
            text=display_text[:100],
            raw_text=raw_text,
            line_start=start_line,
            line_end=end_line,
            element_type=element_type,
            original_type=element_type
        )
        self.paragraphs.append(para)
        return True
    
    def _clean_latex(self, text: str) -> str:
        """清理 LaTeX 命令，提取纯文本"""
        # 移除常见命令
        text = re.sub(r'\\(?:textbf|textit|emph|underline)\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\(?:section|subsection|subsubsection|chapter|paragraph)\*?\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\begin\{\w+\}', '', text)
        text = re.sub(r'\\end\{\w+\}', '', text)
        text = re.sub(r'\\caption\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\label\{[^}]*\}', '', text)
        text = re.sub(r'\\ref\{[^}]*\}', '[ref]', text)
        text = re.sub(r'\\cite\{[^}]*\}', '[cite]', text)
        text = re.sub(r'\\\w+', '', text)  # 移除其他命令
        text = re.sub(r'\{|\}', '', text)  # 移除花括号
        text = re.sub(r'\s+', ' ', text)  # 合并空白
        return text.strip()
    
    def _group_by_type(self):
        """按类型分组"""
        self.format_groups = {}
        
        for para in self.paragraphs:
            sig = para.element_type
            if sig not in self.format_groups:
                self.format_groups[sig] = LatexFormatGroup(
                    signature=sig,
                    element_type=para.element_type,
                    original_type=para.original_type,
                    paragraph_indices=[para.index],
                    sample_text=para.text[:50]
                )
            else:
                self.format_groups[sig].paragraph_indices.append(para.index)
    
    def assign_type_to_paragraph(self, para_index: int, element_type: str):
        """为指定段落分配类型"""
        for para in self.paragraphs:
            if para.index == para_index:
                para.element_type = element_type
                break
    
    def get_paragraph_by_index(self, index: int) -> Optional[LatexParagraphInfo]:
        """根据索引获取段落"""
        for para in self.paragraphs:
            if para.index == index:
                return para
        return None
