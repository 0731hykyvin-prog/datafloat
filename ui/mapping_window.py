from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from core.template_manager import TemplateManager


class MappingWindow(QDialog):
    def __init__(self, columns):
        super().__init__()

        self.setWindowTitle("字段映射")
        self.resize(620, 560)

        self.columns = columns
        self.tm = TemplateManager()
        self.mapping_boxes = {}

        self.init_ui()
        self.apply_style()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title = QLabel("创建字段映射模板")
        title.setObjectName("mappingTitle")
        hint = QLabel("将源文件字段匹配到平台标准字段，未使用的字段保持“空”。")
        hint.setObjectName("mappingHint")

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        standard_fields = [
            "本方号码",
            "对方号码",
            "开始时间",
            "结束时间",
            "通话时长",
            "呼叫类型",
            "小区号",
            "基站号",
            "IMSI",
            "IMEI",
        ]

        for field in standard_fields:
            combo = QComboBox()
            combo.addItem("空")
            combo.addItems(self.columns)
            self.mapping_boxes[field] = combo
            form.addRow(field, combo)

        button_bar = QHBoxLayout()
        button_bar.addStretch(1)
        self.btn_cancel = QPushButton("取消")
        self.btn_save = QPushButton("保存模板")
        self.btn_save.setObjectName("primaryButton")
        button_bar.addWidget(self.btn_cancel)
        button_bar.addWidget(self.btn_save)

        layout.addWidget(title)
        layout.addWidget(hint)
        layout.addLayout(form, 1)
        layout.addLayout(button_bar)

        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save.clicked.connect(self.save_mapping)

    def apply_style(self):
        self.setStyleSheet(
            """
            QDialog {
                background: #f5f7fb;
                color: #1f2937;
                font-family: "PingFang SC", "Microsoft YaHei", Arial;
                font-size: 14px;
            }
            QLabel#mappingTitle {
                font-size: 20px;
                font-weight: 700;
                color: #111827;
            }
            QLabel#mappingHint {
                color: #6b7280;
            }
            QComboBox {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 7px;
                padding: 7px;
                min-height: 22px;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 7px;
                padding: 8px 14px;
            }
            QPushButton#primaryButton {
                background: #2563eb;
                border-color: #2563eb;
                color: white;
                font-weight: 600;
            }
            """
        )

    def save_mapping(self):
        mapping = {}
        for std_field, combo in self.mapping_boxes.items():
            value = combo.currentText()
            if value != "空":
                mapping[std_field] = value

        if not mapping:
            QMessageBox.warning(self, "保存模板", "请至少选择一个字段映射。")
            return

        name, ok = QInputDialog.getText(self, "保存模板", "请输入模板名称")
        if not ok or not name.strip():
            return

        self.tm.save_template(name.strip(), mapping)
        QMessageBox.information(self, "保存模板", f"模板已保存：{name.strip()}")
        self.accept()
