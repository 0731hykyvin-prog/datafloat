"""
Qt 兼容层：优先 PySide2 (Win7)，回退 PySide6 (Mac 开发)。
两者在本项目中的 API 完全一致，无需任何条件判断。
"""

try:
    from PySide2.QtCore import Qt
    from PySide2.QtWidgets import (
        QCheckBox, QComboBox, QDialog, QFileDialog, QFormLayout, QFrame,
        QGridLayout, QGroupBox, QHBoxLayout, QInputDialog, QLabel,
        QLineEdit, QListWidget, QMessageBox, QPushButton, QSpinBox,
        QSplitter, QTabWidget, QTableWidget, QTableWidgetItem,
        QTextEdit, QVBoxLayout, QWidget,
    )
    from PySide2.QtWidgets import QApplication
except ImportError:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import (
        QCheckBox, QComboBox, QDialog, QFileDialog, QFormLayout, QFrame,
        QGridLayout, QGroupBox, QHBoxLayout, QInputDialog, QLabel,
        QLineEdit, QListWidget, QMessageBox, QPushButton, QSpinBox,
        QSplitter, QTabWidget, QTableWidget, QTableWidgetItem,
        QTextEdit, QVBoxLayout, QWidget,
    )
    from PySide6.QtWidgets import QApplication
