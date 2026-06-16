import os

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)


class PreviewWindow(QDialog):
    def __init__(self, file_path):
        super().__init__()

        self.setWindowTitle("数据预览")
        self.resize(1040, 640)
        self.file_path = file_path

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel(os.path.basename(file_path))
        title.setObjectName("previewTitle")
        self.info_label = QLabel("正在读取...")
        self.info_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header.addWidget(title, 1)
        header.addWidget(self.info_label)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)

        layout.addLayout(header)
        layout.addWidget(self.table, 1)

        self.apply_style()
        self.load_data()

    def apply_style(self):
        self.setStyleSheet(
            """
            QDialog {
                background: #f5f7fb;
                color: #1f2937;
                font-family: "PingFang SC", "Microsoft YaHei", Arial;
                font-size: 14px;
            }
            QLabel#previewTitle {
                font-size: 20px;
                font-weight: 700;
                color: #111827;
            }
            QTableWidget {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 8px;
                gridline-color: #e5e7eb;
                selection-background-color: #dbeafe;
                selection-color: #111827;
            }
            QHeaderView::section {
                background: #f3f4f6;
                border: 0;
                border-bottom: 1px solid #d1d5db;
                padding: 8px;
                font-weight: 600;
            }
            """
        )

    def load_data(self):
        ext = os.path.splitext(self.file_path)[1].lower()

        if ext == ".csv":
            df = pd.read_csv(self.file_path, dtype=str)
        elif ext in {".xlsx", ".xlsm"}:
            df = pd.read_excel(self.file_path, engine="openpyxl", dtype=str)
        elif ext == ".xls":
            df = pd.read_excel(self.file_path, engine="xlrd", dtype=str)
        else:
            raise ValueError(f"不支持的文件格式：{ext}")

        total_rows = len(df)
        df = df.head(100).fillna("")
        self.info_label.setText(f"预览 {len(df)} / {total_rows} 行，{len(df.columns)} 列")

        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])

        for row in range(len(df)):
            for col in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iloc[row, col]))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, col, item)

        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)
