"""主窗口"""

import os

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QFrame, QLabel
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

from .styles import MAIN_STYLE
from .smart_format_page import SmartFormatPage


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ToDOCX - 智能排版工具")
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)
        
        # 设置窗口图标
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_dir, "ui", "docx.ico")
        self.setWindowIcon(QIcon(icon_path))
        
        # 应用样式
        self.setStyleSheet(MAIN_STYLE)
        
        self._setup_ui()
    
    def _setup_ui(self):
        # 中央部件
        central = QWidget()
        self.setCentralWidget(central)
        
        # 主布局
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 顶部标题栏
        header = self._create_header()
        main_layout.addWidget(header)
        
        # 智能排版页面（主内容）
        self.format_page = SmartFormatPage()
        main_layout.addWidget(self.format_page)
    
    def _create_header(self):
        """创建顶部标题栏"""
        header = QFrame()
        header.setObjectName("header")
        header.setFixedHeight(60)
        header.setStyleSheet("""
            QFrame#header {
                background-color: #2c3e50;
                border-bottom: 1px solid #34495e;
            }
        """)
        
        layout = QHBoxLayout(header)
        layout.setContentsMargins(20, 0, 20, 0)
        
        # Logo 和标题
        logo_layout = QVBoxLayout()
        logo_layout.setSpacing(2)
        
        logo_label = QLabel("ToDOCX")
        logo_label.setStyleSheet("""
            color: #ffffff;
            font-size: 20px;
            font-weight: bold;
        """)
        logo_layout.addWidget(logo_label)
        
        subtitle = QLabel("琐碎排版，一键告别")
        subtitle.setStyleSheet("color: #95a5a6; font-size: 14px;")
        logo_layout.addWidget(subtitle)
        
        layout.addLayout(logo_layout)
        layout.addStretch()
        
        # 版本号与作者信息
        version_label = QLabel("v1.0.0 by rayL_K")
        version_label.setStyleSheet("color: #7f8c8d; font-size: 14px;")
        layout.addWidget(version_label)
        
        return header
