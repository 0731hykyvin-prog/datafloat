"""
视频布控分析面板
"""

import os
import queue
import threading

from ui.qt_compat import (
    Qt,
    QTimer,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.video_analyzer_core import process_video_data


class VideoPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._busy = False
        self._buttons = []
        self._log_queue = queue.Queue()

        self.init_ui()

        self._log_timer = QTimer()
        self._log_timer.timeout.connect(self._flush_logs)
        self._log_timer.start(100)

    def _set_busy(self, busy):
        self._busy = busy
        self.setCursor(Qt.WaitCursor if busy else Qt.ArrowCursor)
        for b in self._buttons:
            b.setEnabled(not busy)

    def _run_async(self, fn):
        if self._busy:
            return
        self._set_busy(True)
        def w():
            try:
                fn()
            except Exception as e:
                self.log(f"❌ {e}")
            finally:
                self._set_busy(False)
        threading.Thread(target=w, daemon=True).start()

    def init_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        # 左侧控制
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(16, 16, 16, 16)
        cl.setSpacing(12)

        # 输入
        g1 = QGroupBox("输入设置")
        g1l = QVBoxLayout(g1)

        g1l.addWidget(QLabel("CSV 文件夹"))
        self.csv_folder = QLineEdit()
        self.csv_folder.setReadOnly(True)
        g1l.addWidget(self.csv_folder)
        b1 = QPushButton("选择文件夹")
        b1.clicked.connect(lambda: self._pick_folder(self.csv_folder))
        self._buttons.append(b1)
        g1l.addWidget(b1)

        g1l.addWidget(QLabel("人员库文件(可选)"))
        self.person_db = QLineEdit()
        self.person_db.setReadOnly(True)
        g1l.addWidget(self.person_db)
        b2 = QPushButton("选择文件")
        b2.clicked.connect(lambda: self._pick_file(self.person_db))
        self._buttons.append(b2)
        g1l.addWidget(b2)

        g1l.addWidget(QLabel("输出路径"))
        out_row = QHBoxLayout()
        self.output_folder = QLineEdit("C:/Users/Administrator/Desktop/分析结果")
        out_row.addWidget(self.output_folder, 1)
        b3 = QPushButton("选择")
        b3.setFixedWidth(60)
        b3.clicked.connect(lambda: self._pick_folder(self.output_folder))
        self._buttons.append(b3)
        out_row.addWidget(b3)
        g1l.addLayout(out_row)
        cl.addWidget(g1)

        # 参数
        g2 = QGroupBox("参数设置")
        g2l = QVBoxLayout(g2)
        g2l.addWidget(QLabel("相似度阈值"))
        self.spin_sim = QDoubleSpinBox()
        self.spin_sim.setRange(0.5, 1.0)
        self.spin_sim.setSingleStep(0.01)
        self.spin_sim.setValue(0.9)
        g2l.addWidget(self.spin_sim)
        g2l.addWidget(QLabel("去重时间窗口(秒)"))
        self.spin_tw = QSpinBox()
        self.spin_tw.setRange(1, 30)
        self.spin_tw.setValue(5)
        g2l.addWidget(self.spin_tw)
        cl.addWidget(g2)

        btn = QPushButton("开始分析")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(lambda: self._run_async(self._do_analysis))
        self._buttons.append(btn)
        cl.addWidget(btn)
        cl.addStretch()

        # 右侧输出
        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        ol.setContentsMargins(16, 16, 16, 16)
        ol.setSpacing(10)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(150)

        ol.addWidget(self.table, 1)
        ol.addWidget(self.log_box)

        root.addWidget(ctrl, 1)
        root.addWidget(out, 4)

    def _do_analysis(self):
        folder = self.csv_folder.text()
        if not folder:
            self.log("请选择 CSV 文件夹")
            return
        self.log_box.clear()
        self.log("⏳ 开始分析...")

        output_path, summary, logs = process_video_data(
            folder,
            self.person_db.text(),
            self.output_folder.text(),
            self.spin_sim.value(),
            self.spin_tw.value(),
            log_callback=lambda m: self.log(m),
        )

        if output_path:
            self.log(f"✅ 输出: {output_path}")
            # 加载结果到表格
            import pandas as pd
            df = pd.read_excel(output_path)
            self._show_table(df)

    def _show_table(self, df):
        self.table.setRowCount(0)
        if df is None or df.empty:
            return
        show = df.head(300)
        self.table.setColumnCount(len(show.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in show.columns])
        self.table.setRowCount(len(show))
        for ri, (_, row) in enumerate(show.iterrows()):
            for ci, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(ri, ci, item)
        self.table.resizeColumnsToContents()

    def log(self, msg):
        """线程安全日志。"""
        self._log_queue.put(str(msg))

    def _flush_logs(self):
        while not self._log_queue.empty():
            try:
                self.log_box.append(self._log_queue.get_nowait())
            except queue.Empty:
                break

    def _pick_folder(self, entry):
        path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if path:
            entry.setText(path)

    def _pick_file(self, entry):
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel (*.xlsx *.xls)")
        if path:
            entry.setText(path)
