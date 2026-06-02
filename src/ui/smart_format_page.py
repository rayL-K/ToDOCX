"""智能排版页面 - 交互式预览版"""

from pathlib import Path
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QGroupBox, QFormLayout, QComboBox, QSpinBox,
    QDoubleSpinBox, QCheckBox, QLineEdit, QTabWidget, QScrollArea,
    QFrame, QSizePolicy, QMessageBox, QSplitter, QListWidget,
    QListWidgetItem, QInputDialog, QRadioButton, QButtonGroup,
    QMenu, QAction, QTreeWidget, QTreeWidgetItem
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QCursor


class NoWheelComboBox(QComboBox):
    """禁用鼠标滚轮切换选项的下拉框"""
    def wheelEvent(self, event):
        # 忽略滚轮事件，防止误操作
        event.ignore()

from .components import FileDropZone, ProgressWidget
from ..config import FONT_SIZE_OPTIONS
from ..conversion_request import build_conversion_request
from ..diagnostics import get_log_path, get_logger, log_event, log_exception
from ..template_manager import TemplateManager
from ..docx_analyzer import DocxAnalyzer
from ..errors import OperationCancelledError, ToDOCXError, user_message_for_error
from ..latex_analyzer import LatexAnalyzer
from ..resource_policy import ResourcePolicy
from ..user_settings import UserSettingsStore


# UI样式
UI_STYLE = """
    QWidget { font-size: 12px; }
    QGroupBox { font-size: 12px; font-weight: bold; padding: 6px; margin-top: 10px; }
    QGroupBox::title { padding: 0 4px; }
    QLabel { font-size: 12px; }
    QPushButton { font-size: 12px; padding: 4px 8px; min-height: 22px; }
    QComboBox { font-size: 12px; padding: 2px; min-height: 22px; }
    QSpinBox { font-size: 12px; padding: 2px; min-height: 22px; }
    QDoubleSpinBox { font-size: 12px; padding: 2px; min-height: 22px; }
    QCheckBox { font-size: 12px; }
    QRadioButton { font-size: 12px; }
    QLineEdit { font-size: 12px; padding: 3px; min-height: 22px; }
    QListWidget { font-size: 12px; }
    QTreeWidget { font-size: 12px; }
    QTabWidget::pane { font-size: 12px; border: 1px solid #ccc; }
    QTabBar::tab { font-size: 12px; padding: 4px 8px; }
"""

ELEMENT_TYPES = [
    ("original", "原格式"),
    ("heading1", "一级标题"),
    ("heading2", "二级标题"),
    ("heading3", "三级标题"),
    ("heading4", "四级标题"),
    ("body", "正文"),
    ("caption", "图表标题"),
    ("code", "代码"),
    ("table", "表格"),
    ("formula", "公式"),
    ("quote", "引用"),
]

ELEMENT_TYPE_NAMES = {t[0]: t[1] for t in ELEMENT_TYPES}


class ConvertWorker(QThread):
    """转换工作线程"""

    progress = pyqtSignal(int, str)
    convert_finished = pyqtSignal(str)
    error = pyqtSignal(object)
    cancelled = pyqtSignal(str)

    def __init__(self, converter_func, *args, **kwargs):
        super().__init__()
        self.converter_func = converter_func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            def progress_callback(value, message):
                if self.isInterruptionRequested():
                    raise OperationCancelledError(
                        "转换已取消。",
                        code="TODX901",
                    )
                self.progress.emit(value, message)

            self.kwargs['progress_callback'] = progress_callback
            result = self.converter_func(*self.args, **self.kwargs)
            if self.isInterruptionRequested():
                raise OperationCancelledError(
                    "转换已取消。",
                    code="TODX901",
                )
            self.convert_finished.emit(result)
        except OperationCancelledError as error:
            self.cancelled.emit(error.user_message)
        except Exception as error:
            self.error.emit(error)

    def cancel(self):
        """请求取消当前任务。"""

        self.requestInterruption()



class FileLoadWorker(QThread):
    """文件异步加载工作线程"""

    load_finished = pyqtSignal(object)  # dict: {seq, type, ...} — 避开 QThread.finished 命名冲突
    error_py = pyqtSignal(object)  # Exception

    def __init__(self, file_path: str, seq: int = 0):
        super().__init__()
        self.file_path = file_path
        self.seq = seq

    def run(self):
        try:
            path_lower = self.file_path.lower()
            if path_lower.endswith('.docx'):
                from ..docx_analyzer import DocxAnalyzer
                analyzer = DocxAnalyzer()
                analyzer.load_document(self.file_path)
                if self.isInterruptionRequested():
                    return
                self.load_finished.emit({
                    'seq': self.seq,
                    'type': 'docx',
                    'analyzer': analyzer,
                    'paragraphs_count': len(analyzer.paragraphs),
                    'groups_count': len(analyzer.format_groups),
                })
            elif path_lower.endswith('.tex'):
                from ..latex_analyzer import LatexAnalyzer
                analyzer = LatexAnalyzer()
                analyzer.load_document(self.file_path)
                if self.isInterruptionRequested():
                    return
                self.load_finished.emit({
                    'seq': self.seq,
                    'type': 'latex',
                    'analyzer': analyzer,
                    'paragraphs_count': len(analyzer.paragraphs),
                })
            else:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                if self.isInterruptionRequested():
                    return
                paragraphs = SmartFormatPage._parse_markdown(content)
                if self.isInterruptionRequested():
                    return
                self.load_finished.emit({
                    'seq': self.seq,
                    'type': 'markdown',
                    'content': content,
                    'paragraphs': paragraphs,
                })
        except Exception as e:
            # 取消后抛出的异常不再传给主线程
            if self.isInterruptionRequested():
                return
            self.error_py.emit(e)

    def cancel(self):
        """请求取消当前任务。"""
        self.requestInterruption()

class SmartFormatPage(QWidget):
    """智能排版页面 - 交互式预览版"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.worker = None
        self.load_worker = None
        self._load_seq = 0
        self._zombie_workers = []  # 持有已取消但线程未结束的 worker 引用
        self.analyzer = DocxAnalyzer()
        self.latex_analyzer = None  # LaTeX 分析器
        self.template_manager = TemplateManager()
        self.settings_store = UserSettingsStore()
        self.user_settings = self.settings_store.load()
        self.logger = get_logger("smart_format_page")
        self.format_mappings = {}
        self.current_file_type = None  # 'docx', 'latex', 'markdown'
        self._template_tab_refreshed = False

        self.setStyleSheet(UI_STYLE)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.setContentsMargins(10, 10, 10, 10)

        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)

        # 左侧：设置面板
        left_panel = self._create_left_panel()
        splitter.addWidget(left_panel)

        # 右侧：交互式预览面板
        right_panel = self._create_right_panel()
        splitter.addWidget(right_panel)

        # 调整左右面板比例，进一步增大右侧预览区
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 4)
        splitter.setSizes([600, 600])
        layout.addWidget(splitter)

    def _create_left_panel(self):
        """创建左侧设置面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 4, 0)
        layout.setSpacing(4)

        # 文件选择区
        self.file_zone = FileDropZone(accept_extensions=['.docx', '.md', '.markdown', '.tex'])
        self.file_zone.setMaximumHeight(70)
        self.file_zone.fileSelected.connect(self._on_file_selected)
        layout.addWidget(self.file_zone)

        # 创建标签页
        self.tab_widget = QTabWidget()

        # 样式设置标签页
        style_tab = self._create_style_tab()
        self.tab_widget.addTab(style_tab, "样式")

        # 模板管理标签页
        template_tab = self._create_template_tab()
        self.tab_widget.addTab(template_tab, "模板")
        self.tab_widget.currentChanged.connect(self._on_tab_changed)
        layout.addWidget(self.tab_widget)

        # 输出路径设置（移到左侧底部）
        output_group = QGroupBox("输出路径")
        output_layout = QHBoxLayout(output_group)
        output_layout.setContentsMargins(4, 8, 4, 4)
        output_layout.setSpacing(4)
        
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("默认与源文件同目录")
        if self.user_settings.last_output_dir:
            self.output_path.setText(self.user_settings.last_output_dir)
        self.browse_btn = QPushButton("...")
        self.browse_btn.setMaximumWidth(30)
        self.browse_btn.clicked.connect(self._browse_output)
        output_layout.addWidget(self.output_path)
        output_layout.addWidget(self.browse_btn)
        
        layout.addWidget(output_group)

        # 进度显示
        self.progress_widget = ProgressWidget()
        layout.addWidget(self.progress_widget)

        # 操作按钮
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(6)

        self.clear_btn = QPushButton("清除")
        self.clear_btn.clicked.connect(self._clear)

        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_convert)

        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setStyleSheet("background-color: #3498db; color: white; font-weight: bold;")
        self.convert_btn.clicked.connect(self._start_convert)

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addWidget(self.convert_btn)

        layout.addLayout(btn_layout)

        return panel

    def _create_right_panel(self):
        """创建右侧交互式预览面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(4, 0, 0, 0)
        layout.setSpacing(4)

        # 预览标题与操作提示
        header_layout = QHBoxLayout()
        header_layout.setSpacing(4)
        preview_title = QLabel("预览")
        preview_title.setStyleSheet("font-size: 12px; font-weight: bold; color: #2c3e50;")
        header_layout.addWidget(preview_title)

        self.preview_hint = QLabel("右键修改类型")
        self.preview_hint.setStyleSheet("color: #999; font-size: 9px;")
        header_layout.addWidget(self.preview_hint)
        header_layout.addStretch()

        # 范围选择
        self.scope_combo = QComboBox()
        self.scope_combo.addItems(["全文", "选中"])
        self.scope_combo.setMaximumWidth(60)
        self.scope_combo.currentIndexChanged.connect(self._on_scope_changed)
        header_layout.addWidget(QLabel("范围:"))
        header_layout.addWidget(self.scope_combo)

        layout.addLayout(header_layout)

        # 交互式段落列表
        self.paragraph_tree = QTreeWidget()
        self.paragraph_tree.setHeaderLabels(["类型", "内容预览"])
        self.paragraph_tree.setColumnWidth(0, 120)
        self.paragraph_tree.header().setStretchLastSection(True)
        self.paragraph_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.paragraph_tree.customContextMenuRequested.connect(self._show_context_menu)
        self.paragraph_tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        self._update_tree_selection_style()

        layout.addWidget(self.paragraph_tree)

        # 格式信息
        self.format_info_label = QLabel("选择DOCX文件后显示内容")
        self.format_info_label.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(self.format_info_label)

        return panel

    def _on_scope_changed(self, index):
        """范围选择变化"""
        self._update_tree_selection_style()

    def _update_tree_selection_style(self):
        """根据范围选择更新树的选择样式"""
        is_global = self.scope_combo.currentIndex() == 0
        
        if is_global:
            # 全文应用：黄色选中
            self.paragraph_tree.setStyleSheet("""
                QTreeWidget {
                    background-color: #ffffff;
                    border: 1px solid #e0e6ed;
                    border-radius: 4px;
                }
                QTreeWidget::item {
                    padding: 2px;
                    border-bottom: 1px solid #f0f0f0;
                }
                QTreeWidget::item:selected {
                    background-color: #fff3cd;
                    color: #856404;
                }
            """)
        else:
            # 仅选中：蓝色选中
            self.paragraph_tree.setStyleSheet("""
                QTreeWidget {
                    background-color: #ffffff;
                    border: 1px solid #e0e6ed;
                    border-radius: 4px;
                }
                QTreeWidget::item {
                    padding: 2px;
                    border-bottom: 1px solid #f0f0f0;
                }
                QTreeWidget::item:selected {
                    background-color: #cce5ff;
                    color: #004085;
                }
            """)

    def _show_context_menu(self, position):
        """显示右键菜单"""
        items = self.paragraph_tree.selectedItems()
        if not items:
            return

        menu = QMenu()

        # 添加类型选项
        for type_id, type_name in ELEMENT_TYPES:
            action = QAction(type_name, self)
            action.setData(type_id)
            action.triggered.connect(lambda checked, t=type_id: self._set_selected_type(t))
            menu.addAction(action)

        menu.exec_(QCursor.pos())

    def _set_selected_type(self, type_id):
        """设置选中项的类型"""
        items = self.paragraph_tree.selectedItems()
        file_type = getattr(self, 'current_file_type', None)
        changed_signatures = set()

        for item in items:
            sig = item.data(0, Qt.UserRole)
            if sig:
                changed_signatures.add(sig)
                if type_id == "original":
                    # 恢复原格式：删除映射
                    self.format_mappings.pop(sig, None)
                    if file_type == 'docx':
                        # DOCX: 从分析器获取原始类型
                        group = self.analyzer.format_groups.get(sig)
                        original_type = group.original_type if group and group.original_type else "body"
                        self.analyzer.assign_type_to_format(sig, original_type)
                else:
                    # 更新映射
                    self.format_mappings[sig] = type_id
                    if file_type == 'docx':
                        # DOCX: 更新分析器
                        self.analyzer.assign_type_to_format(sig, type_id)

        if not changed_signatures:
            return

        for i in range(self.paragraph_tree.topLevelItemCount()):
            tree_item = self.paragraph_tree.topLevelItem(i)
            if tree_item.data(0, Qt.UserRole) not in changed_signatures:
                continue

            if file_type == 'latex':
                self._refresh_latex_item_type(tree_item)
            elif file_type == 'markdown':
                self._refresh_markdown_item_type(tree_item)
            else:
                self._refresh_item_type(tree_item)


    def _on_tab_changed(self, index: int) -> None:
        """标签页切换：首次切换到模板页时刷新列表"""
        if index == 1 and not self._template_tab_refreshed:
            self._template_tab_refreshed = True
            self._refresh_template_list()
    def _create_style_tab(self):
        """创建样式设置标签页"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(2)
        layout.setContentsMargins(4, 4, 4, 4)

        # 各级标题设置
        self._create_heading_settings(layout)

        # 正文设置
        self._create_body_settings(layout)

        # 图表标题设置
        self._create_caption_settings(layout)
        
        # 代码样式设置
        self._create_code_settings(layout)

        layout.addStretch()
        scroll.setWidget(widget)
        return scroll

    def _create_heading_settings(self, parent_layout):
        """创建标题设置"""
        group = QGroupBox("标题")
        layout = QVBoxLayout(group)
        layout.setSpacing(2)
        layout.setContentsMargins(4, 8, 4, 4)

        self.heading_widgets = {}

        for i in range(1, 5):
            row_layout = QHBoxLayout()
            row_layout.setSpacing(4)

            label = QLabel(f"{i}级:")
            label.setFixedWidth(25)
            row_layout.addWidget(label)

            # 中文字体
            font_cn_combo = NoWheelComboBox()
            font_cn_combo.addItems(["宋体", "黑体", "微软雅黑", "楷体", "仿宋"])
            font_cn_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
            font_cn_combo.setMinimumContentsLength(3)
            row_layout.addWidget(font_cn_combo)

            # 西文字体
            font_en_combo = NoWheelComboBox()
            font_en_combo.addItems(["Times New Roman", "Arial", "Calibri"])
            font_en_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
            font_en_combo.setMinimumContentsLength(6)
            row_layout.addWidget(font_en_combo)

            # 字号
            size_combo = NoWheelComboBox()
            size_combo.addItems(FONT_SIZE_OPTIONS)
            size_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
            size_combo.setMinimumContentsLength(2)
            default_sizes = ["小三", "四号", "小四", "小四"]
            idx = size_combo.findText(default_sizes[i - 1])
            if idx >= 0:
                size_combo.setCurrentIndex(idx)
            row_layout.addWidget(size_combo)

            # 加粗
            bold_check = QCheckBox("粗")
            bold_check.setChecked(i <= 3)
            row_layout.addWidget(bold_check)

            row_layout.addStretch()
            layout.addLayout(row_layout)

            self.heading_widgets[f"heading{i}"] = {
                "font_cn": font_cn_combo,
                "font_en": font_en_combo,
                "size": size_combo,
                "bold": bold_check
            }

        parent_layout.addWidget(group)

    def _create_body_settings(self, parent_layout):
        """创建正文设置"""
        group = QGroupBox("正文")
        layout = QFormLayout(group)
        layout.setSpacing(2)
        layout.setContentsMargins(4, 8, 4, 4)

        # 字体行
        font_row = QHBoxLayout()
        font_row.setSpacing(4)

        self.body_font_cn = NoWheelComboBox()
        self.body_font_cn.addItems(["宋体", "黑体", "微软雅黑", "楷体", "仿宋"])
        self.body_font_cn.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.body_font_cn.setMinimumContentsLength(3)
        font_row.addWidget(self.body_font_cn)

        self.body_font_en = NoWheelComboBox()
        self.body_font_en.addItems(["Times New Roman", "Arial", "Calibri"])
        self.body_font_en.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.body_font_en.setMinimumContentsLength(6)
        font_row.addWidget(self.body_font_en)

        self.body_size = NoWheelComboBox()
        self.body_size.addItems(FONT_SIZE_OPTIONS)
        self.body_size.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.body_size.setMinimumContentsLength(2)
        idx = self.body_size.findText("小四")
        if idx >= 0:
            self.body_size.setCurrentIndex(idx)
        font_row.addWidget(self.body_size)
        font_row.addStretch()

        layout.addRow("字体:", font_row)

        # 行间距
        spacing_row = QHBoxLayout()
        spacing_row.setSpacing(4)

        self.spacing_type_group = QButtonGroup(self)
        self.spacing_multiple_radio = QRadioButton("倍数")
        self.spacing_exact_radio = QRadioButton("固定")
        self.spacing_exact_radio.setChecked(True)
        self.spacing_type_group.addButton(self.spacing_multiple_radio, 0)
        self.spacing_type_group.addButton(self.spacing_exact_radio, 1)

        spacing_row.addWidget(self.spacing_multiple_radio)
        self.spacing_multiple_spin = QDoubleSpinBox()
        self.spacing_multiple_spin.setRange(1.0, 3.0)
        self.spacing_multiple_spin.setValue(1.5)
        self.spacing_multiple_spin.setSingleStep(0.25)
        self.spacing_multiple_spin.setMaximumWidth(50)
        self.spacing_multiple_spin.setEnabled(False)
        spacing_row.addWidget(self.spacing_multiple_spin)

        spacing_row.addWidget(self.spacing_exact_radio)
        self.spacing_exact_spin = QSpinBox()
        self.spacing_exact_spin.setRange(10, 50)
        self.spacing_exact_spin.setValue(20)
        self.spacing_exact_spin.setMaximumWidth(50)
        spacing_row.addWidget(self.spacing_exact_spin)
        spacing_row.addWidget(QLabel("磅"))
        spacing_row.addStretch()

        self.spacing_multiple_radio.toggled.connect(self._on_spacing_type_changed)
        layout.addRow("行距:", spacing_row)

        # 首行缩进和对齐
        indent_row = QHBoxLayout()
        indent_row.setSpacing(4)

        self.body_indent = QSpinBox()
        self.body_indent.setRange(0, 4)
        self.body_indent.setValue(2)
        self.body_indent.setMaximumWidth(40)
        indent_row.addWidget(self.body_indent)
        indent_row.addWidget(QLabel("字符"))

        indent_row.addSpacing(10)
        indent_row.addWidget(QLabel("对齐:"))
        self.body_align = NoWheelComboBox()
        self.body_align.addItems(["左", "两端", "中", "右"])
        self.body_align.setMaximumWidth(50)
        indent_row.addWidget(self.body_align)
        indent_row.addStretch()

        layout.addRow("缩进:", indent_row)

        parent_layout.addWidget(group)

    def _create_caption_settings(self, parent_layout):
        """创建图表标题设置"""
        group = QGroupBox("图表标题")
        layout = QHBoxLayout(group)
        layout.setSpacing(4)
        layout.setContentsMargins(4, 8, 4, 4)

        self.caption_font_cn = NoWheelComboBox()
        self.caption_font_cn.addItems(["黑体", "宋体", "微软雅黑"])
        self.caption_font_cn.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.caption_font_cn.setMinimumContentsLength(3)
        layout.addWidget(self.caption_font_cn)

        self.caption_font_en = NoWheelComboBox()
        self.caption_font_en.addItems(["Times New Roman", "Arial"])
        self.caption_font_en.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.caption_font_en.setMinimumContentsLength(6)
        layout.addWidget(self.caption_font_en)

        self.caption_size = NoWheelComboBox()
        self.caption_size.addItems(FONT_SIZE_OPTIONS)
        self.caption_size.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.caption_size.setMinimumContentsLength(2)
        idx = self.caption_size.findText("小五")
        if idx >= 0:
            self.caption_size.setCurrentIndex(idx)
        layout.addWidget(self.caption_size)
        layout.addStretch()

        parent_layout.addWidget(group)

    def _create_code_settings(self, parent_layout):
        """创建代码样式设置"""
        group = QGroupBox("代码")
        layout = QHBoxLayout(group)
        layout.setSpacing(4)
        layout.setContentsMargins(4, 8, 4, 4)

        self.code_font = NoWheelComboBox()
        self.code_font.addItems(["Consolas", "Courier New", "Monaco", "等线"])
        self.code_font.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.code_font.setMinimumContentsLength(6)
        layout.addWidget(self.code_font)

        self.code_size = NoWheelComboBox()
        self.code_size.addItems(FONT_SIZE_OPTIONS)
        self.code_size.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.code_size.setMinimumContentsLength(2)
        idx = self.code_size.findText("小五")
        if idx >= 0:
            self.code_size.setCurrentIndex(idx)
        layout.addWidget(self.code_size)
        layout.addStretch()

        parent_layout.addWidget(group)

    def _on_spacing_type_changed(self, checked):
        """行距类型切换"""
        self.spacing_multiple_spin.setEnabled(checked)
        self.spacing_exact_spin.setEnabled(not checked)

    def _create_template_tab(self):
        """创建模板管理标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(4)
        layout.setContentsMargins(4, 4, 4, 4)

        self.template_list = QListWidget()
        layout.addWidget(self.template_list)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(4)

        load_btn = QPushButton("加载")
        load_btn.clicked.connect(self._load_template)
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self._save_template)
        rename_btn = QPushButton("重命名")
        rename_btn.clicked.connect(self._rename_template)
        del_btn = QPushButton("删除")
        del_btn.clicked.connect(self._delete_template)

        btn_layout.addWidget(load_btn)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(rename_btn)
        btn_layout.addWidget(del_btn)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)
        return widget

    def _refresh_template_list(self):
        """刷新模板列表"""
        self.template_list.clear()
        templates = self.template_manager.list_templates()
        for tpl in templates:
            status = tpl.get("status", "ok")
            prefix = "[损坏]" if status != "ok" else ""
            label = f"{prefix} {tpl['name']}" if prefix else tpl["name"]
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, ("user", tpl["name"], status))
            item.setToolTip(f"{tpl.get('description', '')}\n{tpl['file']}")
            self.template_list.addItem(item)

    def _load_template(self):
        """加载模板"""
        current = self.template_list.currentItem()
        if not current:
            QMessageBox.warning(self, "提示", "请先选择模板")
            return

        tpl_data = current.data(Qt.UserRole)
        tpl_type, tpl_name = tpl_data[0], tpl_data[1]
        tpl_status = tpl_data[2] if len(tpl_data) > 2 else "ok"
        if tpl_status != "ok":
            QMessageBox.warning(self, "提示", "该模板文件已损坏，请删除后重新保存。")
            return
        try:
            styles = self.template_manager.load_template(tpl_name)
        except Exception as error:
            self._show_error("加载模板失败", error)
            return

        if styles:
            self._apply_styles_to_ui(styles)
            QMessageBox.information(self, "成功", f"已加载: {tpl_name}")

    def _save_template(self):
        """保存模板"""
        name, ok = QInputDialog.getText(self, "保存模板", "模板名称:")
        if ok and name:
            styles = self._get_current_styles()
            try:
                file_path = self.template_manager.save_template(name, styles)
            except Exception as error:
                self._show_error("保存模板失败", error)
                return
            self._refresh_template_list()
            QMessageBox.information(self, "保存成功", f"模板已保存至:\n{file_path}")

    def _rename_template(self):
        """重命名模板"""
        current = self.template_list.currentItem()
        if not current:
            QMessageBox.warning(self, "提示", "请先选择模板")
            return
        tpl_data = current.data(Qt.UserRole)
        tpl_type, tpl_name = tpl_data[0], tpl_data[1]
        new_name, ok = QInputDialog.getText(self, "重命名模板", "新名称:", text=tpl_name)
        if ok and new_name and new_name != tpl_name:
            try:
                renamed = self.template_manager.rename_template(tpl_name, new_name)
            except Exception as error:
                self._show_error("重命名模板失败", error)
                return
            if renamed:
                self._refresh_template_list()
                QMessageBox.information(self, "成功", f"已重命名为: {new_name}")

    def _delete_template(self):
        """删除模板"""
        current = self.template_list.currentItem()
        if not current:
            return
        tpl_data = current.data(Qt.UserRole)
        tpl_type, tpl_name = tpl_data[0], tpl_data[1]
        reply = QMessageBox.question(self, "确认", f"删除 '{tpl_name}'?")
        if reply == QMessageBox.Yes:
            try:
                self.template_manager.delete_template(tpl_name)
            except Exception as error:
                self._show_error("删除模板失败", error)
                return
            self._refresh_template_list()

    def _apply_styles_to_ui(self, styles):
        """将样式应用到UI"""
        for i in range(1, 5):
            key = f"heading{i}"
            if key in styles and key in self.heading_widgets:
                widgets = self.heading_widgets[key]
                s = styles[key]
                if "font_name_cn" in s:
                    idx = widgets["font_cn"].findText(s["font_name_cn"])
                    if idx >= 0:
                        widgets["font_cn"].setCurrentIndex(idx)
                if "font_name_en" in s:
                    idx = widgets["font_en"].findText(s["font_name_en"])
                    if idx >= 0:
                        widgets["font_en"].setCurrentIndex(idx)
                if "font_size" in s:
                    idx = widgets["size"].findText(str(s["font_size"]))
                    if idx >= 0:
                        widgets["size"].setCurrentIndex(idx)
                if "bold" in s:
                    widgets["bold"].setChecked(s["bold"])

        if "body" in styles:
            body = styles["body"]
            if "font_name_cn" in body:
                idx = self.body_font_cn.findText(body["font_name_cn"])
                if idx >= 0:
                    self.body_font_cn.setCurrentIndex(idx)
            if "font_name_en" in body:
                idx = self.body_font_en.findText(body["font_name_en"])
                if idx >= 0:
                    self.body_font_en.setCurrentIndex(idx)
            if "font_size" in body:
                idx = self.body_size.findText(str(body["font_size"]))
                if idx >= 0:
                    self.body_size.setCurrentIndex(idx)
            if "line_spacing_type" in body:
                if body["line_spacing_type"] == "固定值":
                    self.spacing_exact_radio.setChecked(True)
                    val = body.get("line_spacing_value", 20)
                    if isinstance(val, (int, float)):
                        self.spacing_exact_spin.setValue(int(val))
                else:
                    self.spacing_multiple_radio.setChecked(True)
                    val = body.get("line_spacing_value", 1.5)
                    if isinstance(val, (int, float)):
                        self.spacing_multiple_spin.setValue(float(val))
            if "first_line_indent" in body:
                self.body_indent.setValue(body["first_line_indent"])
            if "alignment" in body:
                align_map = {"left": 0, "justify": 1, "center": 2, "right": 3}
                self.body_align.setCurrentIndex(align_map.get(body["alignment"], 0))

        if "caption" in styles:
            cap = styles["caption"]
            if "font_name_cn" in cap:
                idx = self.caption_font_cn.findText(cap["font_name_cn"])
                if idx >= 0:
                    self.caption_font_cn.setCurrentIndex(idx)
            if "font_name_en" in cap:
                idx = self.caption_font_en.findText(cap["font_name_en"])
                if idx >= 0:
                    self.caption_font_en.setCurrentIndex(idx)
            if "font_size" in cap:
                idx = self.caption_size.findText(str(cap["font_size"]))
                if idx >= 0:
                    self.caption_size.setCurrentIndex(idx)

    def _get_current_styles(self):
        """获取当前UI中的样式配置"""
        styles = {}

        # 行距设置
        if self.spacing_exact_radio.isChecked():
            line_spacing_type = "固定值"
            line_spacing_value = self.spacing_exact_spin.value()
        else:
            line_spacing_type = "倍数行距"
            line_spacing_value = self.spacing_multiple_spin.value()

        # 标题样式
        for i in range(1, 5):
            key = f"heading{i}"
            if key in self.heading_widgets:
                widgets = self.heading_widgets[key]
                styles[key] = {
                    "font_name_cn": widgets["font_cn"].currentText(),
                    "font_name_en": widgets["font_en"].currentText(),
                    "font_size": widgets["size"].currentText(),
                    "bold": widgets["bold"].isChecked(),
                    "line_spacing_type": line_spacing_type,
                    "line_spacing_value": line_spacing_value,
                    "alignment": "left",
                }

        # 正文样式
        align_map = {0: "left", 1: "justify", 2: "center", 3: "right"}
        styles["body"] = {
            "font_name_cn": self.body_font_cn.currentText(),
            "font_name_en": self.body_font_en.currentText(),
            "font_size": self.body_size.currentText(),
            "bold": False,
            "line_spacing_type": line_spacing_type,
            "line_spacing_value": line_spacing_value,
            "first_line_indent": self.body_indent.value(),
            "alignment": align_map[self.body_align.currentIndex()],
        }

        # 图表标题样式
        styles["caption"] = {
            "font_name_cn": self.caption_font_cn.currentText(),
            "font_name_en": self.caption_font_en.currentText(),
            "font_size": self.caption_size.currentText(),
            "bold": False,
            "line_spacing_type": "1.5倍行距",
            "line_spacing_value": 1.5,
            "alignment": "center",
        }

        # 代码样式
        styles["code"] = {
            "font_name": self.code_font.currentText(),
            "font_size": self.code_size.currentText(),
            "line_spacing_type": "固定值",
            "line_spacing_value": 14,
        }

        # 图片和公式
        styles["image"] = {
            "alignment": "center",
            "max_width": 15,
            "line_spacing_type": "1.5倍行距",
            "line_spacing_value": 1.5,
        }
        styles["formula"] = {
            "alignment": "center",
            "line_spacing_type": "1.5倍行距",
            "line_spacing_value": 1.5,
        }

        return styles

    def _on_file_selected(self, file_path):
        """文件选择后的处理（异步加载）"""
        self.paragraph_tree.clear()
        self.format_mappings = {}  # 清空之前的映射

        # 如果已有加载中的任务：断开信号、丢回僵尸池避免 GC 提前销毁 QThread
        if self.load_worker is not None:
            try:
                self.load_worker.load_finished.disconnect(self._on_load_finished)
                self.load_worker.error_py.disconnect(self._on_load_error)
            except TypeError:
                pass
            self.load_worker.cancel()
            self._zombie_workers.append(self.load_worker)
            self.load_worker = None
        self.current_file_type = None  # 记录当前文件类型

        # 显示加载状态
        self.format_info_label.setText("正在加载文件...")
        self.progress_widget.reset()
        self.progress_widget.set_progress(0, "正在加载文件...")
        self._set_loading_state(True)
        self._load_seq += 1

        self.load_worker = FileLoadWorker(file_path, seq=self._load_seq)

        file_lower = file_path.lower()
        if file_lower.endswith('.docx'):
            self.current_file_type = 'docx'
        elif file_lower.endswith('.tex'):
            self.current_file_type = 'latex'
        else:
            self.current_file_type = 'markdown'

        self.load_worker.load_finished.connect(self._on_load_finished)
        self.load_worker.error_py.connect(self._on_load_error)
        self.load_worker.start()

    def _on_load_finished(self, data: dict):
        """异步加载完成后的处理"""
        # 丢弃过时 worker 的信号（用户已切换文件）
        if data.get('seq', -1) != self._load_seq:
            return

        self._set_loading_state(False)
        # 不释放 worker — 工作线程的 run() 可能还在清理中

        file_type = data['type']
        if file_type == 'docx':
            self.analyzer = data['analyzer']
            self._populate_paragraph_tree()
            self.format_info_label.setText(
                f"共 {data['paragraphs_count']} 段，{data['groups_count']} 种格式"
            )
        elif file_type == 'latex':
            self.latex_analyzer = data['analyzer']
            self._populate_latex_tree()
            self.format_info_label.setText(
                f"LaTeX文档：共 {data['paragraphs_count']} 段"
            )
        elif file_type == 'markdown':
            self.md_paragraphs = data['paragraphs']
            self._populate_markdown_tree()
            self.format_info_label.setText(
                f"Markdown文件：共 {len(self.md_paragraphs)} 段"
            )
        self.progress_widget.set_progress(100, "加载完成")

    def _on_load_error(self, error: Exception):
        """异步加载失败处理"""
        self._set_loading_state(False)
        # 不释放 worker — 同上
        self.format_info_label.setText("文件加载失败")
        self.progress_widget.set_error("加载失败")
        self._show_error("读取文件失败", error)

    def _set_loading_state(self, is_loading: bool) -> None:
        """切换加载中的 UI 状态"""
        self.file_zone.setEnabled(not is_loading)
        self.convert_btn.setEnabled(not is_loading)
        self.clear_btn.setEnabled(not is_loading)
        # blockSignals 彻底阻断禁用状态下 Qt 版本边缘情况可能投递的信号
        self.clear_btn.blockSignals(is_loading)
    def _populate_latex_tree(self):
        """填充 LaTeX 段落树"""
        self.paragraph_tree.clear()

        for para in self.latex_analyzer.paragraphs:
            item = QTreeWidgetItem(["", para.text[:80]])
            sig = f"latex_para_{para.index}"
            item.setData(0, Qt.UserRole, sig)
            item.setData(1, Qt.UserRole, para.index)
            # 存储原始类型
            item.setData(2, Qt.UserRole, para.original_type)
            self.paragraph_tree.addTopLevelItem(item)

        # 刷新显示
        for i in range(self.paragraph_tree.topLevelItemCount()):
            self._refresh_latex_item_type(self.paragraph_tree.topLevelItem(i))

    def _refresh_latex_item_type(self, item):
        """刷新 LaTeX 段落项的类型显示"""
        sig = item.data(0, Qt.UserRole)
        if not sig or not sig.startswith("latex_para_"):
            return

        original_type = item.data(2, Qt.UserRole) or "body"

        # 判断是否用户自定义类型
        if sig in self.format_mappings:
            type_id = self.format_mappings[sig]
            is_original = False
        else:
            type_id = original_type
            is_original = True

        base_text = ELEMENT_TYPE_NAMES.get(type_id, "正文")

        if is_original:
            display_text = f"{base_text}（原）"
            color = Qt.gray
        else:
            display_text = base_text
            color = Qt.black

        item.setText(0, display_text)
        item.setForeground(0, color)

    @staticmethod
    def _parse_markdown(content: str) -> list:
        """解析Markdown内容，识别各段落类型
        
        Returns:
            [(index, type_id, text, original_text), ...]
        """
        import re
        paragraphs = []
        lines = content.split('\n')
        
        in_code_block = False
        code_block_content = []
        para_idx = 0
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # 代码块开始/结束
            if stripped.startswith('```'):
                if in_code_block:
                    # 代码块结束
                    code_text = '\n'.join(code_block_content)
                    if code_text.strip():
                        paragraphs.append((para_idx, 'code', code_text[:80], code_text))
                        para_idx += 1
                    code_block_content = []
                    in_code_block = False
                else:
                    # 代码块开始
                    in_code_block = True
                i += 1
                continue
            
            if in_code_block:
                code_block_content.append(line)
                i += 1
                continue
            
            # 空行跳过
            if not stripped:
                i += 1
                continue
            
            # 标题识别
            if stripped.startswith('######'):
                paragraphs.append((para_idx, 'heading4', stripped[6:].strip()[:80], stripped))
                para_idx += 1
            elif stripped.startswith('#####'):
                paragraphs.append((para_idx, 'heading4', stripped[5:].strip()[:80], stripped))
                para_idx += 1
            elif stripped.startswith('####'):
                paragraphs.append((para_idx, 'heading4', stripped[4:].strip()[:80], stripped))
                para_idx += 1
            elif stripped.startswith('###'):
                paragraphs.append((para_idx, 'heading3', stripped[3:].strip()[:80], stripped))
                para_idx += 1
            elif stripped.startswith('##'):
                paragraphs.append((para_idx, 'heading2', stripped[2:].strip()[:80], stripped))
                para_idx += 1
            elif stripped.startswith('#'):
                paragraphs.append((para_idx, 'heading1', stripped[1:].strip()[:80], stripped))
                para_idx += 1
            # 引用
            elif stripped.startswith('>'):
                paragraphs.append((para_idx, 'quote', stripped[1:].strip()[:80], stripped))
                para_idx += 1
            # 图片
            elif re.match(r'^!\[.*\]\(.*\)$', stripped):
                paragraphs.append((para_idx, 'caption', stripped[:80], stripped))
                para_idx += 1
            # 公式块
            elif stripped.startswith('$$') or stripped.startswith('$'):
                paragraphs.append((para_idx, 'formula', stripped[:80], stripped))
                para_idx += 1
            # 普通正文
            else:
                paragraphs.append((para_idx, 'body', stripped[:80], stripped))
                para_idx += 1
            
            i += 1
        
        return paragraphs

    def _populate_markdown_tree(self):
        """填充Markdown段落树"""
        self.paragraph_tree.clear()
        
        for para_idx, type_id, preview_text, original_text in self.md_paragraphs:
            item = QTreeWidgetItem(["", preview_text])
            # 使用类型作为签名（同类型共享）
            sig = f"md_type_{type_id}"
            item.setData(0, Qt.UserRole, sig)
            item.setData(1, Qt.UserRole, para_idx)
            item.setData(2, Qt.UserRole, type_id)  # 原始类型
            self.paragraph_tree.addTopLevelItem(item)
        
        # 刷新显示
        for i in range(self.paragraph_tree.topLevelItemCount()):
            self._refresh_markdown_item_type(self.paragraph_tree.topLevelItem(i))

    def _refresh_markdown_item_type(self, item):
        """刷新Markdown段落项的类型显示"""
        sig = item.data(0, Qt.UserRole)
        if not sig or not sig.startswith("md_type_"):
            return
        
        original_type = item.data(2, Qt.UserRole) or "body"
        
        # 判断是否用户自定义类型
        if sig in self.format_mappings:
            type_id = self.format_mappings[sig]
            is_original = False
        else:
            type_id = original_type
            is_original = True
        
        base_text = ELEMENT_TYPE_NAMES.get(type_id, "正文")
        
        if is_original:
            display_text = f"{base_text}（原）"
            color = Qt.gray
        else:
            display_text = base_text
            color = Qt.black
        
        item.setText(0, display_text)
        item.setForeground(0, color)

    def _populate_paragraph_tree(self):
        """填充段落树"""
        self.paragraph_tree.clear()

        for para in self.analyzer.paragraphs:
            group = self.analyzer.format_groups.get(para.format_signature)
            # 先占位，稍后统一根据当前映射和原始类型刷新显示
            item = QTreeWidgetItem(["", para.text[:80]])
            item.setData(0, Qt.UserRole, para.format_signature)
            item.setData(1, Qt.UserRole, para.index)
            self.paragraph_tree.addTopLevelItem(item)

        # 填充完成后统一刷新显示文本和颜色
        for i in range(self.paragraph_tree.topLevelItemCount()):
            self._refresh_item_type(self.paragraph_tree.topLevelItem(i))

    def _refresh_item_type(self, item):
        """根据当前映射/原始类型刷新单个条目的类型显示"""
        sig = item.data(0, Qt.UserRole)
        if not sig:
            return

        group = self.analyzer.format_groups.get(sig)

        # 判断是否用户自定义类型
        if sig in self.format_mappings:
            type_id = self.format_mappings[sig]
            is_original = False
        else:
            # 使用原始识别类型
            if group and getattr(group, "original_type", ""):
                type_id = group.original_type
            elif group:
                type_id = group.suggested_type
            else:
                type_id = "body"
            is_original = True

        base_text = ELEMENT_TYPE_NAMES.get(type_id, "正文")

        # 原格式增加灰色小“（原）”标记
        if is_original:
            display_text = f"{base_text}（原）"
            color = Qt.gray
            font = item.font(0)
            # 字体略小一点，保证区分
            if font.pointSize() > 0:
                font.setPointSize(max(font.pointSize() - 1, 8))
            item.setFont(0, font)
        else:
            display_text = base_text
            color = Qt.black

        item.setText(0, display_text)
        item.setForeground(0, color)

    def _browse_output(self):
        """浏览输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", self.output_path.text())
        if dir_path:
            self.output_path.setText(dir_path)
            self._remember_output_dir(dir_path)

    def _clear(self):
        """清除"""
        # 禁用状态下的事件丢弃（Qt 版本边缘情况）
        if not self.clear_btn.isEnabled():
            return
        # 取消正在进行的加载或转换（丢入僵尸池防止 GC 提前销毁 QThread）
        # 注意：僵尸只追加不清理，isRunning() 返回 False 后 C++ cleanup 可能仍在进行
        self._load_seq += 1
        if self.load_worker and self.load_worker.isRunning():
            self.load_worker.cancel()
            self._zombie_workers.append(self.load_worker)
            self.load_worker = None
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self._zombie_workers.append(self.worker)
            self.worker = None
        self.file_zone.clear()
        self.paragraph_tree.clear()
        self.format_info_label.setText("选择DOCX文件后显示内容")
        self.progress_widget.reset()
        self.format_mappings = {}
        self.current_file_type = None
        self._set_loading_state(False)

    def _start_convert(self):
        """开始转换"""
        # 如果在转换中，取消旧任务（丢入僵尸池）
        if self.worker and self.worker.isRunning():
            try:
                self.worker.convert_finished.disconnect(self._on_convert_finished)
                self.worker.error.disconnect(self._on_convert_error)
                self.worker.cancelled.disconnect(self._on_convert_cancelled)
            except TypeError:
                pass
            self.worker.cancel()
            self._zombie_workers.append(self.worker)
            self.worker = None
        input_file = self.file_zone.get_file()
        if not input_file:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return

        styles = self._get_current_styles()
        explicit_output_dir = self.output_path.text().strip()
        request_kwargs = {
            "input_file": input_file,
            "output_dir_text": explicit_output_dir,
            "styles": styles,
            "resource_policy": ResourcePolicy(),
        }

        input_suffix = Path(input_file).suffix.lower()
        if input_suffix == ".tex":
            request_kwargs["paragraph_mappings"] = self._get_latex_paragraph_mappings()
        elif input_suffix in [".docx", ".doc"]:
            paragraph_mappings = self._get_docx_paragraph_mappings()
            if not paragraph_mappings:
                QMessageBox.information(self, "提示", "当前文档没有可应用格式的段落")
                return
            request_kwargs["paragraph_mappings"] = paragraph_mappings
        else:
            request_kwargs["type_overrides"] = self._get_markdown_type_overrides()

        try:
            request = build_conversion_request(**request_kwargs)
        except Exception as error:
            self._show_error("无法开始转换", error)
            return

        if explicit_output_dir:
            self._remember_output_dir(str(request.output_dir))

        # 根据文件类型选择不同的处理方式
        if request.input_type == "latex":
            from ..latex_formatter import convert_latex_to_docx

            self.worker = ConvertWorker(
                convert_latex_to_docx,
                str(request.input_path),
                str(request.output_path),
                paragraph_mappings=request.paragraph_mappings,
                styles=request.styles,
            )
        elif request.input_type == "docx":
            from ..formatter import SmartFormatter

            formatter = SmartFormatter()
            self.worker = ConvertWorker(
                formatter.apply_selective_format,
                str(request.input_path),
                str(request.output_path),
                paragraph_mappings=request.paragraph_mappings,
                styles=request.styles,
            )
        else:
            from ..formatter import SmartFormatter

            formatter = SmartFormatter()
            self.worker = ConvertWorker(
                formatter.format_document,
                str(request.input_path),
                str(request.output_path),
                styles=request.styles,
                type_overrides=request.type_overrides,
                resource_policy=request.resource_policy,
                use_ai=False,
            )
        
        self.worker.progress.connect(self.progress_widget.set_progress)
        self.worker.convert_finished.connect(self._on_convert_finished)
        self.worker.error.connect(self._on_convert_error)
        self.worker.cancelled.connect(self._on_convert_cancelled)

        self.progress_widget.reset()
        self._set_worker_state(True)
        log_event(
            self.logger,
            "用户发起转换",
            input=str(request.input_path),
            output=str(request.output_path),
            input_type=request.input_type,
        )
        self.worker.start()

    def _set_worker_state(self, is_running: bool) -> None:
        """统一切换转换中的 UI 状态。"""

        self.convert_btn.setEnabled(not is_running)
        self.cancel_btn.setEnabled(is_running)
        self.clear_btn.setEnabled(not is_running)
        self.browse_btn.setEnabled(not is_running)
        self.convert_btn.blockSignals(is_running)
        self.clear_btn.blockSignals(is_running)
        self.browse_btn.blockSignals(is_running)

    def _remember_output_dir(self, output_dir: str) -> None:
        """持久化最近使用的输出目录。"""

        try:
            self.user_settings = self.settings_store.update_last_output_dir(output_dir)
        except Exception as error:
            log_exception(
                self.logger,
                "记录最近输出目录失败",
                error,
                output_dir=output_dir,
            )

    def _show_error(self, title: str, error: Exception) -> None:
        """展示用户可读错误，并记录开发者日志。"""

        if isinstance(error, ToDOCXError):
            log_event(
                self.logger,
                title,
                code=error.code,
                message=error.user_message,
            )
            message = user_message_for_error(error)
        else:
            log_exception(
                self.logger,
                title,
                error,
                log_path=str(get_log_path()),
            )
            message = (
                f"{user_message_for_error(error)}\n\n"
                f"日志文件：{get_log_path()}"
            )
        QMessageBox.critical(self, title, message)

    def _cancel_convert(self):
        """请求取消当前转换。"""

        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.progress_widget.set_error("正在取消，请稍候...")

    def _get_docx_paragraph_mappings(self) -> dict:
        """获取 DOCX 当前应应用到各段落的完整类型映射。"""
        paragraph_mappings = {}
        
        for sig, group in self.analyzer.format_groups.items():
            current_type = self.format_mappings.get(sig)
            if not current_type:
                current_type = group.original_type or group.suggested_type or "body"

            for para_idx in group.paragraph_indices:
                paragraph_mappings[para_idx] = current_type
        
        return paragraph_mappings

    def _get_latex_paragraph_mappings(self) -> dict:
        """获取 LaTeX 文件用户修改过类型的段落映射
        
        Returns:
            {段落索引: 新类型} 的字典
        """
        paragraph_mappings = {}
        
        for sig, new_type in self.format_mappings.items():
            if sig.startswith("latex_para_"):
                para_index = int(sig.replace("latex_para_", ""))
                paragraph_mappings[para_index] = new_type

        return paragraph_mappings

    def _get_markdown_type_overrides(self) -> dict:
        """获取 Markdown 识别类型到目标类型的映射。"""
        type_overrides = {}

        for sig, new_type in self.format_mappings.items():
            if sig.startswith("md_type_"):
                original_type = sig.replace("md_type_", "")
                type_overrides[original_type] = new_type
        
        return type_overrides

    def _on_convert_finished(self, output_path):
        """转换完成"""
        self._set_worker_state(False)
        self.progress_widget.set_success("完成")
        log_event(self.logger, "转换完成", output=output_path)
        self.worker = None
        QMessageBox.information(self, "完成", f"已保存到:\n{output_path}")

    def _on_convert_error(self, error):
        """转换出错"""
        self._set_worker_state(False)
        message = user_message_for_error(error)
        if not isinstance(error, ToDOCXError):
            log_exception(
                self.logger,
                "转换失败",
                error,
                log_path=str(get_log_path()),
            )
            message = f"{message}\n\n日志文件：{get_log_path()}"
        self.progress_widget.set_error("转换失败")
        self.worker = None
        QMessageBox.critical(self, "错误", f"转换失败：\n{message}")

    def _on_convert_cancelled(self, message: str):
        """转换取消。"""

        self._set_worker_state(False)
        self.progress_widget.set_error(message)
        log_event(self.logger, "转换已取消")
        self.worker = None
        QMessageBox.information(self, "已取消", message)
