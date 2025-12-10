"""Qt样式表"""

MAIN_STYLE = """
QMainWindow {
    background-color: #f5f7fa;
}

QWidget {
    font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
}

/* 侧边栏 */
#sidebar {
    background-color: #2c3e50;
    border: none;
}

#sidebar QPushButton {
    background-color: transparent;
    color: #ecf0f1;
    border: none;
    border-radius: 8px;
    padding: 12px 16px;
    text-align: left;
    font-size: 14px;
    margin: 4px 8px;
}

#sidebar QPushButton:hover {
    background-color: #34495e;
}

#sidebar QPushButton:checked {
    background-color: #3498db;
    color: white;
}

#sidebar QPushButton:pressed {
    background-color: #2980b9;
}

/* 主内容区 */
#mainContent {
    background-color: #ffffff;
    border-radius: 12px;
    margin: 16px;
}

/* 卡片样式 */
.card {
    background-color: #ffffff;
    border-radius: 12px;
    border: 1px solid #e0e6ed;
}

/* 标题 */
#pageTitle {
    font-size: 24px;
    font-weight: bold;
    color: #2c3e50;
    padding: 20px;
}

/* 按钮样式 */
QPushButton {
    background-color: #3498db;
    color: white;
    border: none;
    border-radius: 6px;
    padding: 10px 20px;
    font-size: 14px;
    font-weight: 500;
}

QPushButton:hover {
    background-color: #2980b9;
}

QPushButton:pressed {
    background-color: #1f6dad;
}

QPushButton:disabled {
    background-color: #bdc3c7;
    color: #7f8c8d;
}

QPushButton#primaryBtn {
    background-color: #27ae60;
}

QPushButton#primaryBtn:hover {
    background-color: #219a52;
}

QPushButton#secondaryBtn {
    background-color: #95a5a6;
}

QPushButton#secondaryBtn:hover {
    background-color: #7f8c8d;
}

QPushButton#dangerBtn {
    background-color: #e74c3c;
}

QPushButton#dangerBtn:hover {
    background-color: #c0392b;
}

/* 文件选择区 */
#fileDropZone {
    background-color: #f8fafc;
    border: 2px dashed #cbd5e1;
    border-radius: 12px;
    min-height: 150px;
}

#fileDropZone:hover {
    border-color: #3498db;
    background-color: #eff6ff;
}

/* 输入框 */
QLineEdit, QTextEdit {
    border: 1px solid #e0e6ed;
    border-radius: 6px;
    padding: 10px;
    background-color: #ffffff;
    font-size: 14px;
}

QLineEdit:focus, QTextEdit:focus {
    border-color: #3498db;
    outline: none;
}

/* 下拉框 */
QComboBox {
    border: 1px solid #e0e6ed;
    border-radius: 6px;
    padding: 8px 12px;
    background-color: #ffffff;
    font-size: 14px;
    min-width: 150px;
}

QComboBox:hover {
    border-color: #3498db;
}

QComboBox::drop-down {
    border: none;
    width: 30px;
}

QComboBox::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #7f8c8d;
    margin-right: 10px;
}

QComboBox QAbstractItemView {
    border: 1px solid #e0e6ed;
    border-radius: 6px;
    background-color: #ffffff;
    selection-background-color: #3498db;
}

/* 进度条 */
QProgressBar {
    border: none;
    border-radius: 6px;
    background-color: #e0e6ed;
    height: 8px;
    text-align: center;
}

QProgressBar::chunk {
    background-color: #3498db;
    border-radius: 6px;
}

/* 标签 */
QLabel {
    color: #2c3e50;
}

QLabel#subtitle {
    color: #7f8c8d;
    font-size: 13px;
}

/* 分组框 */
QGroupBox {
    font-weight: bold;
    border: 1px solid #e0e6ed;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 12px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 8px;
    color: #2c3e50;
}

/* 复选框 */
QCheckBox {
    spacing: 8px;
    font-size: 14px;
}

QCheckBox::indicator {
    width: 18px;
    height: 18px;
    border-radius: 4px;
    border: 2px solid #cbd5e1;
}

QCheckBox::indicator:checked {
    background-color: #3498db;
    border-color: #3498db;
}

QCheckBox::indicator:hover {
    border-color: #3498db;
}

/* 单选按钮 */
QRadioButton {
    spacing: 8px;
    font-size: 14px;
}

QRadioButton::indicator {
    width: 18px;
    height: 18px;
    border-radius: 9px;
    border: 2px solid #cbd5e1;
}

QRadioButton::indicator:checked {
    background-color: #3498db;
    border-color: #3498db;
}

/* 滚动条 */
QScrollBar:vertical {
    border: none;
    background-color: #f0f0f0;
    width: 10px;
    border-radius: 5px;
}

QScrollBar::handle:vertical {
    background-color: #c0c0c0;
    border-radius: 5px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background-color: #a0a0a0;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

/* SpinBox */
QSpinBox, QDoubleSpinBox {
    border: 1px solid #e0e6ed;
    border-radius: 6px;
    padding: 6px 10px;
    background-color: #ffffff;
    font-size: 14px;
}

QSpinBox:focus, QDoubleSpinBox:focus {
    border-color: #3498db;
}

/* TabWidget */
QTabWidget::pane {
    border: 1px solid #e0e6ed;
    border-radius: 8px;
    background-color: #ffffff;
}

QTabBar::tab {
    background-color: #f5f7fa;
    border: 1px solid #e0e6ed;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    padding: 10px 20px;
    margin-right: 2px;
}

QTabBar::tab:selected {
    background-color: #ffffff;
    border-bottom: 1px solid #ffffff;
}

QTabBar::tab:hover:!selected {
    background-color: #e8ecf0;
}

/* 状态标签 */
#statusLabel {
    color: #7f8c8d;
    font-size: 12px;
    padding: 4px 8px;
}

#successLabel {
    color: #27ae60;
}

#errorLabel {
    color: #e74c3c;
}
"""
