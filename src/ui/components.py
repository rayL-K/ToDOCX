"""UI 组件"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QProgressBar, QFrame, QSizePolicy
)
from PyQt5.QtCore import Qt, pyqtSignal, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent


class FileDropZone(QFrame):
    """文件拖放区域"""
    
    fileSelected = pyqtSignal(str)
    
    def __init__(self, accept_extensions: list = None, parent=None):
        super().__init__(parent)
        self.accept_extensions = accept_extensions or ['.docx', '.md', '.markdown', '.tex']
        self.selected_file = None
        
        self.setObjectName("fileDropZone")
        self.setAcceptDrops(True)
        self.setMinimumHeight(150)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        
        # 图标
        self.icon_label = QLabel("📁")
        self.icon_label.setStyleSheet("font-size: 48px;")
        self.icon_label.setAlignment(Qt.AlignCenter)
        
        # 提示文本
        self.hint_label = QLabel("拖放文件到此处，或点击选择文件")
        self.hint_label.setStyleSheet("color: #7f8c8d; font-size: 14px;")
        self.hint_label.setAlignment(Qt.AlignCenter)
        
        # 支持格式提示
        ext_text = "支持格式: " + ", ".join(self.accept_extensions)
        self.format_label = QLabel(ext_text)
        self.format_label.setStyleSheet("color: #bdc3c7; font-size: 12px;")
        self.format_label.setAlignment(Qt.AlignCenter)
        
        # 已选文件显示
        self.file_label = QLabel()
        self.file_label.setStyleSheet("color: #27ae60; font-size: 14px; font-weight: bold;")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.hide()
        
        layout.addWidget(self.icon_label)
        layout.addWidget(self.hint_label)
        layout.addWidget(self.format_label)
        layout.addWidget(self.file_label)
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._select_file()
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            file_path = url.toLocalFile()
            if self._is_valid_file(file_path):
                event.acceptProposedAction()
                self.setStyleSheet("""
                    #fileDropZone {
                        background-color: #e8f4fc;
                        border: 2px dashed #3498db;
                    }
                """)
    
    def dragLeaveEvent(self, event):
        self.setStyleSheet("")
    
    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet("")
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            file_path = url.toLocalFile()
            if self._is_valid_file(file_path):
                self._set_file(file_path)
                event.acceptProposedAction()
    
    def _select_file(self):
        ext_filter = "支持的文件 (" + " ".join([f"*{ext}" for ext in self.accept_extensions]) + ")"
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择文件", "", ext_filter
        )
        if file_path:
            self._set_file(file_path)
    
    def _is_valid_file(self, file_path: str) -> bool:
        return any(file_path.lower().endswith(ext) for ext in self.accept_extensions)
    
    def _set_file(self, file_path: str):
        self.selected_file = file_path
        
        # 获取文件名
        file_name = file_path.split('/')[-1].split('\\')[-1]
        
        self.icon_label.setText("✅")
        self.hint_label.hide()
        self.format_label.hide()
        self.file_label.setText(file_name)
        self.file_label.show()
        
        self.fileSelected.emit(file_path)
    
    def clear(self):
        """清除已选文件"""
        self.selected_file = None
        self.icon_label.setText("📁")
        self.hint_label.show()
        self.format_label.show()
        self.file_label.hide()
    
    def get_file(self) -> str:
        """获取已选文件路径"""
        return self.selected_file


class ProgressWidget(QWidget):
    """进度显示组件"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        
        # 状态文本
        self.status_label = QLabel("准备就绪")
        self.status_label.setObjectName("statusLabel")
        
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.status_label)
    
    def set_progress(self, value: int, message: str = None):
        """设置进度"""
        self.progress_bar.setValue(value)
        self.status_label.setStyleSheet("")
        if message:
            self.status_label.setText(message)
    
    def reset(self):
        """重置进度"""
        self.progress_bar.setValue(0)
        self.status_label.setText("准备就绪")
        self.status_label.setStyleSheet("")
    
    def set_success(self, message: str = "完成"):
        """设置成功状态"""
        self.progress_bar.setValue(100)
        self.status_label.setText(message)
        self.status_label.setStyleSheet("color: #27ae60;")
    
    def set_error(self, message: str = "出错了"):
        """设置错误状态"""
        self.status_label.setText(message)
        self.status_label.setStyleSheet("color: #e74c3c;")


class StyledButton(QPushButton):
    """样式化按钮"""
    
    def __init__(self, text: str, style_type: str = "primary", parent=None):
        super().__init__(text, parent)
        
        style_map = {
            "primary": "primaryBtn",
            "secondary": "secondaryBtn",
            "danger": "dangerBtn",
        }
        
        if style_type in style_map:
            self.setObjectName(style_map[style_type])


class SectionHeader(QWidget):
    """区块标题"""
    
    def __init__(self, title: str, subtitle: str = None, parent=None):
        super().__init__(parent)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 12)
        
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #2c3e50;")
        layout.addWidget(title_label)
        
        if subtitle:
            subtitle_label = QLabel(subtitle)
            subtitle_label.setObjectName("subtitle")
            layout.addWidget(subtitle_label)
