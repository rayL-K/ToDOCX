"""智能排版页面 - 交互式预览版"""

import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QGroupBox, QFormLayout, QComboBox, QSpinBox,
    QDoubleSpinBox, QCheckBox, QLineEdit, QTabWidget, QScrollArea,
    QFrame, QSizePolicy, QMessageBox, QSplitter, QListWidget,
    QListWidgetItem, QInputDialog, QRadioButton, QButtonGroup,
    QMenu, QAction, QTreeWidget, QTreeWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QEvent
from PyQt5.QtGui import QFont, QCursor


class NoWheelComboBox(QComboBox):
    """禁用鼠标滚轮切换选项的下拉框"""
    def wheelEvent(self, event):
        # 忽略滚轮事件，防止误操作
        event.ignore()

from .components import FileDropZone, ProgressWidget, SectionHeader, StyledButton
from ..config import (
    DEFAULT_STYLES, FONT_SIZE_MAP,
    FONT_SIZE_OPTIONS, get_font_size_pt
)
from ..template_manager import TemplateManager
from ..docx_analyzer import DocxAnalyzer
from ..latex_analyzer import LatexAnalyzer


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
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, converter_func, *args, **kwargs):
        super().__init__()
        self.converter_func = converter_func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            def progress_callback(value, message):
                self.progress.emit(value, message)

            self.kwargs['progress_callback'] = progress_callback
            result = self.converter_func(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class SmartFormatPage(QWidget):
    """智能排版页面 - 交互式预览版"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.worker = None
        self.analyzer = DocxAnalyzer()
        self.latex_analyzer = None  # LaTeX 分析器
        self.template_manager = TemplateManager()
        self.format_mappings = {}
        self.current_file_type = None  # 'docx', 'latex', 'markdown'

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

        layout.addWidget(self.tab_widget)

        # 输出路径设置（移到左侧底部）
        output_group = QGroupBox("输出路径")
        output_layout = QHBoxLayout(output_group)
        output_layout.setContentsMargins(4, 8, 4, 4)
        output_layout.setSpacing(4)
        
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("默认与源文件同目录")
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

        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setStyleSheet("background-color: #3498db; color: white; font-weight: bold;")
        self.convert_btn.clicked.connect(self._start_convert)

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()
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

        for item in items:
            sig = item.data(0, Qt.UserRole)
            if sig:
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

        # 更新所有相同签名的项的显示
        if items:
            first_sig = items[0].data(0, Qt.UserRole)
            if first_sig:
                for i in range(self.paragraph_tree.topLevelItemCount()):
                    tree_item = self.paragraph_tree.topLevelItem(i)
                    if tree_item.data(0, Qt.UserRole) == first_sig:
                        if file_type == 'latex':
                            self._refresh_latex_item_type(tree_item)
                        elif file_type == 'markdown':
                            self._refresh_markdown_item_type(tree_item)
                        else:
                            self._refresh_item_type(tree_item)

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
        self._refresh_template_list()
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
        builtin = self.template_manager.get_builtin_templates()
        for name in builtin.keys():
            item = QListWidgetItem(f"[内置] {name}")
            item.setData(Qt.UserRole, ("builtin", name))
            self.template_list.addItem(item)

        user_templates = self.template_manager.list_templates()
        for tpl in user_templates:
            item = QListWidgetItem(f"[用户] {tpl['name']}")
            item.setData(Qt.UserRole, ("user", tpl['name']))
            self.template_list.addItem(item)

    def _load_template(self):
        """加载模板"""
        current = self.template_list.currentItem()
        if not current:
            QMessageBox.warning(self, "提示", "请先选择模板")
            return

        tpl_type, tpl_name = current.data(Qt.UserRole)
        if tpl_type == "builtin":
            styles = self.template_manager.get_builtin_templates().get(tpl_name, {})
        else:
            styles = self.template_manager.load_template(tpl_name)

        if styles:
            self._apply_styles_to_ui(styles)
            QMessageBox.information(self, "成功", f"已加载: {tpl_name}")

    def _save_template(self):
        """保存模板"""
        name, ok = QInputDialog.getText(self, "保存模板", "模板名称:")
        if ok and name:
            styles = self._get_current_styles()
            self.template_manager.save_template(name, styles)
            self._refresh_template_list()

    def _rename_template(self):
        """重命名模板"""
        current = self.template_list.currentItem()
        if not current:
            QMessageBox.warning(self, "提示", "请先选择模板")
            return
        tpl_type, tpl_name = current.data(Qt.UserRole)
        if tpl_type == "builtin":
            QMessageBox.warning(self, "提示", "内置模板不能重命名")
            return
        new_name, ok = QInputDialog.getText(self, "重命名模板", "新名称:", text=tpl_name)
        if ok and new_name and new_name != tpl_name:
            if self.template_manager.rename_template(tpl_name, new_name):
                self._refresh_template_list()
                QMessageBox.information(self, "成功", f"已重命名为: {new_name}")
            else:
                QMessageBox.warning(self, "失败", "重命名失败")

    def _delete_template(self):
        """删除模板"""
        current = self.template_list.currentItem()
        if not current:
            return
        tpl_type, tpl_name = current.data(Qt.UserRole)
        if tpl_type == "builtin":
            QMessageBox.warning(self, "提示", "内置模板不能删除")
            return
        reply = QMessageBox.question(self, "确认", f"删除 '{tpl_name}'?")
        if reply == QMessageBox.Yes:
            self.template_manager.delete_template(tpl_name)
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
        """文件选择后的处理"""
        self.paragraph_tree.clear()
        self.format_mappings = {}  # 清空之前的映射
        self.current_file_type = None  # 记录当前文件类型

        if file_path.lower().endswith('.docx'):
            self.current_file_type = 'docx'
            if self.analyzer.load_document(file_path):
                self._populate_paragraph_tree()
                groups = self.analyzer.get_format_summary()
                self.format_info_label.setText(f"共 {len(self.analyzer.paragraphs)} 段，{len(groups)} 种格式")
            else:
                self.format_info_label.setText("无法加载文档")
        elif file_path.lower().endswith('.tex'):
            self.current_file_type = 'latex'
            self.latex_analyzer = LatexAnalyzer()
            if self.latex_analyzer.load_document(file_path):
                self._populate_latex_tree()
                self.format_info_label.setText(f"LaTeX文档：共 {len(self.latex_analyzer.paragraphs)} 段")
            else:
                self.format_info_label.setText("无法加载LaTeX文档")
        else:
            self.current_file_type = 'markdown'
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.md_paragraphs = self._parse_markdown(content)
                self._populate_markdown_tree()
                self.format_info_label.setText(f"Markdown文件：共 {len(self.md_paragraphs)} 段")
            except Exception as e:
                self.format_info_label.setText(f"读取失败: {e}")

    def _populate_latex_tree(self):
        """填充 LaTeX 段落树"""
        self.paragraph_tree.clear()

        for para in self.latex_analyzer.paragraphs:
            item = QTreeWidgetItem(["", para.text[:80]])
            # 使用元素类型作为签名（同类型共享签名）
            sig = f"latex_type_{para.original_type}"
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
        if not sig or not sig.startswith("latex_type_"):
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

    def _parse_markdown(self, content: str) -> list:
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

    def _clear(self):
        """清除"""
        self.file_zone.clear()
        self.paragraph_tree.clear()
        self.format_info_label.setText("选择DOCX文件后显示内容")
        self.progress_widget.reset()
        self.format_mappings = {}

    def _start_convert(self):
        """开始转换"""
        input_file = self.file_zone.get_file()
        if not input_file:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return

        input_path = Path(input_file)
        # 输出目录：优先使用用户指定，否则使用源文件所在目录
        output_dir = self.output_path.text().strip()
        if not output_dir:
            output_dir = str(input_path.parent)
        os.makedirs(output_dir, exist_ok=True)

        styles = self._get_current_styles()

        # 根据文件类型选择不同的处理方式
        if input_path.suffix.lower() == '.tex':
            # LaTeX 文件：输出 .docx
            output_file = os.path.join(output_dir, f"{input_path.stem}_formatted.docx")
            paragraph_mappings = self._get_latex_paragraph_mappings()
            
            from ..latex_formatter import convert_latex_to_docx
            self.worker = ConvertWorker(
                convert_latex_to_docx,
                input_file,
                output_file,
                paragraph_mappings=paragraph_mappings,
                styles=styles
            )
        elif input_path.suffix.lower() in ['.docx', '.doc']:
            # DOCX 文件：选择性格式化
            output_file = os.path.join(output_dir, f"{input_path.stem}_formatted.docx")
            paragraph_mappings = self._get_modified_paragraph_mappings()
            
            if not paragraph_mappings:
                QMessageBox.information(self, "提示", "没有修改任何段落类型，无需转换")
                return
            
            from ..formatter import SmartFormatter
            formatter = SmartFormatter()
            self.worker = ConvertWorker(
                formatter.apply_selective_format,
                input_file,
                output_file,
                paragraph_mappings=paragraph_mappings,
                styles=styles
            )
        else:
            # Markdown 文件：全量转换为 DOCX
            output_file = os.path.join(output_dir, f"{input_path.stem}_formatted.docx")
            
            from ..formatter import SmartFormatter
            formatter = SmartFormatter()
            self.worker = ConvertWorker(
                formatter.format_document,
                input_file,
                output_file,
                styles=styles,
                use_ai=False
            )
        
        self.worker.progress.connect(self.progress_widget.set_progress)
        self.worker.finished.connect(self._on_convert_finished)
        self.worker.error.connect(self._on_convert_error)

        self.convert_btn.setEnabled(False)
        self.worker.start()

    def _get_modified_paragraph_mappings(self) -> dict:
        """获取用户手动修改过类型的段落索引和类型映射
        
        Returns:
            {段落索引: 类型} 的字典，只包含用户修改过的段落
        """
        paragraph_mappings = {}
        
        # format_mappings: {格式签名: 类型}
        # 只有用户手动改过类型的才在这里
        for sig, type_id in self.format_mappings.items():
            group = self.analyzer.format_groups.get(sig)
            if group:
                # 获取该格式组包含的所有段落索引
                for para_idx in group.paragraph_indices:
                    paragraph_mappings[para_idx] = type_id
        
        return paragraph_mappings

    def _get_latex_paragraph_mappings(self) -> dict:
        """获取 LaTeX 文件用户修改过类型的段落映射
        
        Returns:
            {段落索引: 新类型} 的字典
        """
        paragraph_mappings = {}
        
        for sig, new_type in self.format_mappings.items():
            if sig.startswith("latex_type_"):
                # 从签名中提取原始类型
                original_type = sig.replace("latex_type_", "")
                # 找到所有该类型的段落
                for para in self.latex_analyzer.paragraphs:
                    if para.original_type == original_type:
                        paragraph_mappings[para.index] = new_type
        
        return paragraph_mappings

    def _on_convert_finished(self, output_path):
        """转换完成"""
        self.convert_btn.setEnabled(True)
        self.progress_widget.set_success("完成")
        QMessageBox.information(self, "完成", f"已保存到:\n{output_path}")

    def _on_convert_error(self, error_msg):
        """转换出错"""
        self.convert_btn.setEnabled(True)
        self.progress_widget.set_error(f"失败: {error_msg}")
        QMessageBox.critical(self, "错误", f"转换失败:\n{error_msg}")
