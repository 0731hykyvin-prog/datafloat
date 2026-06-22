"""
电信原始话单处理面板
PDF → XLSX 批量转换 + 清洗
"""

import os
import threading

from ui.qt_compat import (
    Qt,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.dianxing_cleaning import batch_convert_folder


class TelecomPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.folder_path = None
        self._busy = False
        self.init_ui()

    def init_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        # ── 左侧控制 ──
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(16, 16, 16, 16)
        cl.setSpacing(14)

        # 数据源
        g1 = QGroupBox("数据源")
        g1l = QVBoxLayout(g1)
        self.lbl_folder = QLabel("未选择文件夹")
        self.lbl_folder.setObjectName("pathLabel")
        self.lbl_folder.setWordWrap(True)
        btn_folder = QPushButton("选择 PDF 文件夹")
        btn_folder.setObjectName("primaryButton")
        btn_folder.clicked.connect(self.select_folder)
        g1l.addWidget(self.lbl_folder)
        g1l.addWidget(btn_folder)

        # 操作
        g2 = QGroupBox("操作")
        g2l = QVBoxLayout(g2)
        lbl_hint = QLabel("将文件夹内所有 PDF 转为 XLSX\n文件以业务号码命名")
        lbl_hint.setObjectName("pathLabel")
        self.btn_convert = QPushButton("开始转换")
        self.btn_convert.setObjectName("primaryButton")
        self.btn_convert.clicked.connect(lambda: self._run_async(self.run_convert))
        self.btn_output = QPushButton("打开输出目录")
        self.btn_output.clicked.connect(self.open_output)
        g2l.addWidget(lbl_hint)
        g2l.addWidget(self.btn_convert)
        g2l.addWidget(self.btn_output)

        cl.addWidget(g1)
        cl.addWidget(g2)
        cl.addStretch()

        # ── 右侧输出 ──
        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        ol.setContentsMargins(16, 16, 16, 16)
        ol.setSpacing(10)

        header = QLabel("转换日志")
        header.setStyleSheet("font-size:16px;font-weight:700;color:#111827;")

        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setColumnCount(3)
        self.result_table.setHorizontalHeaderLabels(["源文件", "输出文件", "业务号码"])

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(120)
        self.log_box.setPlaceholderText("转换日志...")

        ol.addWidget(header)
        ol.addWidget(self.result_table, 1)
        ol.addWidget(self.log_box)

        root.addWidget(ctrl, 1)
        root.addWidget(out, 4)

    # ══════════════════════════════════════════════
    # 操作
    # ══════════════════════════════════════════════

    def select_folder(self):
        path = QFileDialog.getExistingDirectory(self, "选择 PDF 文件夹")
        if not path:
            return
        self.folder_path = path
        self.lbl_folder.setText(path)
        pdfs = [f for f in os.listdir(path) if f.lower().endswith(".pdf")]
        self.log(f"📁 {path}")
        self.log(f"   找到 {len(pdfs)} 个 PDF 文件")

    def run_convert(self):
        if not self.folder_path:
            self.log("请先选择文件夹")
            return
        self.log("⏳ 开始批量转换...")
        results, logs = batch_convert_folder(self.folder_path)
        for line in logs:
            self.log(line)
        self._show_results(results)

    def open_output(self):
        if self.folder_path:
            out_dir = os.path.join(self.folder_path, "xlsx输出")
            os.makedirs(out_dir, exist_ok=True)
            import subprocess, sys
            if sys.platform == "darwin":
                subprocess.run(["open", out_dir])
            elif sys.platform == "win32":
                os.startfile(out_dir)
            self.log(f"📂 {out_dir}")

    # ══════════════════════════════════════════════
    # 线程安全
    # ══════════════════════════════════════════════

    def _run_async(self, func):
        if self._busy:
            return
        self._busy = True
        self.btn_convert.setEnabled(False)
        self.setCursor(Qt.WaitCursor)
        def wrapper():
            try:
                func()
            except Exception as e:
                self.log(f"❌ {e}")
            finally:
                self._busy = False
                self.btn_convert.setEnabled(True)
                self.setCursor(Qt.ArrowCursor)
        threading.Thread(target=wrapper, daemon=True).start()

    # ══════════════════════════════════════════════
    # 工具
    # ══════════════════════════════════════════════

    def _show_results(self, results):
        self.result_table.setRowCount(len(results))
        for i, (src, dst, phone) in enumerate(results):
            for j, val in enumerate([
                os.path.basename(src),
                os.path.basename(dst),
                phone or "未识别",
            ]):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.result_table.setItem(i, j, item)
        self.result_table.resizeColumnsToContents()
        self.log(f"✅ 完成: {len(results)} 个文件已转换")

    def log(self, msg):
        self.log_box.append(str(msg))
