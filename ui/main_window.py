"""
Datafloat V3.0 主窗口
TAB 1: 话单分析（通信记录合并+分析）
TAB 2: 银行交易分析（流水快进快出+风险评分）
"""

import os

from ui.qt_compat import (
    Qt,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.analytics import (
    get_data_summary,
    get_global_top_contacts,
    get_night_calls,
    get_risk_contacts,
    get_top_contacts_by_user,
)
from core.excel_engine import merge_excel_files
from core.file_loader import scan_files
from core.template_manager import TemplateManager
from ui.bank_panel import BankPanel
from ui.mapping_window import MappingWindow
from ui.preview_window import PreviewWindow


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Datafloat 数据处理平台 V3.0")
        self.resize(1100, 700)

        # ── 话单分析状态 ──
        self.current_files = []
        self.merged_df = None
        self.template_manager = TemplateManager()
        self.metric_labels = {}

        self.init_ui()
        self.apply_style()
        self.refresh_templates()
        self.update_call_summary()

    # ══════════════════════════════════════════════
    # UI 布局
    # ══════════════════════════════════════════════

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # 标题栏
        root.addWidget(self._build_header())

        # 两个 Tab
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_call_tab(), "📞 话单分析")
        self.tabs.addTab(BankPanel(), "🏦 银行交易分析")
        root.addWidget(self.tabs, 1)

    def _build_header(self):
        frame = QFrame()
        frame.setObjectName("header")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(16, 12, 16, 12)

        title = QLabel("Datafloat 数据处理平台 V3.0")
        title.setObjectName("title")
        sub = QLabel("话单分析 · 银行交易分析  |  Win7/Win10/Win11 通用")
        sub.setObjectName("subtitle")
        box = QVBoxLayout()
        box.addWidget(title)
        box.addWidget(sub)

        self.status_badge = QLabel("就绪")
        self.status_badge.setObjectName("statusBadge")
        self.status_badge.setAlignment(Qt.AlignCenter)

        layout.addLayout(box, 1)
        layout.addWidget(self.status_badge)
        return frame

    # ══════════════════════════════════════════════
    # TAB 1: 话单分析
    # ══════════════════════════════════════════════

    def _build_call_tab(self):
        wrap = QWidget()
        root = QHBoxLayout(wrap)
        root.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_call_sidebar())
        splitter.addWidget(self._build_call_workspace())
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 5)
        root.addWidget(splitter)
        return wrap

    def _build_call_sidebar(self):
        panel = QFrame()
        panel.setObjectName("sidePanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # 数据源
        src = QGroupBox("数据源")
        sl = QVBoxLayout(src)
        self.folder_label = QLabel("未选择文件夹")
        self.folder_label.setObjectName("pathLabel")
        self.folder_label.setWordWrap(True)
        btn_folder = QPushButton("选择话单文件夹")
        btn_folder.setObjectName("primaryButton")
        btn_file = QPushButton("选择话单文件")
        btn_file.setObjectName("primaryButton")
        self.file_list = QListWidget()
        self.file_list.setAlternatingRowColors(True)
        sl.addWidget(self.folder_label)
        sl.addWidget(btn_folder)
        sl.addWidget(btn_file)
        sl.addWidget(QLabel("已选文件"))
        sl.addWidget(self.file_list, 1)
        btn_folder.clicked.connect(self.select_folder)
        btn_file.clicked.connect(self.select_files)
        self.file_list.itemDoubleClicked.connect(self.preview_file)

        # 处理
        proc = QGroupBox("处理与分析")
        pl = QVBoxLayout(proc)
        pl.addWidget(QLabel("数据模板"))
        self.template_box = QComboBox()
        btn_map = QPushButton("字段映射")

        # 输出目录
        pl.addWidget(QLabel("输出目录"))
        out_row = QHBoxLayout()
        self.output_label = QLabel("程序默认目录")
        self.output_label.setObjectName("pathLabel")
        self.output_label.setWordWrap(True)
        btn_out = QPushButton("选择")
        btn_out.setFixedWidth(60)
        out_row.addWidget(self.output_label, 1)
        out_row.addWidget(btn_out)
        pl.addLayout(out_row)
        self.output_dir = None
        btn_out.clicked.connect(self.select_output_dir)

        btn_merge = QPushButton("合并处理")
        btn_merge.setObjectName("primaryButton")
        btn_direct = QPushButton("直接处理当前文件")
        pl.addWidget(self.template_box)
        pl.addWidget(btn_map)
        pl.addWidget(btn_merge)
        pl.addWidget(btn_direct)
        btn_map.clicked.connect(self.open_mapping)
        btn_merge.clicked.connect(lambda: self.start_process(mode="merge"))
        btn_direct.clicked.connect(lambda: self.start_process(mode="direct"))

        layout.addWidget(src, 3)
        layout.addWidget(proc)
        return panel

    def _build_call_workspace(self):
        panel = QFrame()
        panel.setObjectName("workPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # 概览指标
        self.summary_grid = QGridLayout()
        summary_frame = QFrame()
        summary_frame.setObjectName("summaryPanel")
        summary_frame.setLayout(self.summary_grid)
        layout.addWidget(summary_frame)

        # 按钮栏
        bar = QHBoxLayout()
        self.btn_summary = QPushButton("刷新概览")
        self.btn_top = QPushButton("高频联系人")
        self.btn_night = QPushButton("深夜通话")
        self.btn_risk = QPushButton("风险关系")
        for b in [self.btn_summary, self.btn_top, self.btn_night, self.btn_risk]:
            bar.addWidget(b)
        bar.addStretch()
        layout.addLayout(bar)

        # 结果区域
        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setSortingEnabled(True)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("运行日志")
        self.log_box.setMaximumHeight(140)

        layout.addWidget(self.result_table, 1)
        layout.addWidget(self.log_box)

        self.btn_summary.clicked.connect(self.run_call_summary)
        self.btn_top.clicked.connect(self.run_call_top)
        self.btn_night.clicked.connect(self.run_call_night)
        self.btn_risk.clicked.connect(self.run_call_risk)
        return panel

    # ══════════════════════════════════════════════
    # 话单操作
    # ══════════════════════════════════════════════

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择话单文件夹")
        if not folder:
            return
        self.folder_label.setText(folder)
        self.file_list.clear()
        self.current_files = scan_files(folder)
        for fp in self.current_files:
            self.file_list.addItem(os.path.basename(fp))
        self.status_badge.setText("已导入")
        self.log(f"扫描文件夹: {folder}")
        self.log(f"找到 {len(self.current_files)} 个文件")

    def select_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择话单文件", "",
            "表格文件 (*.xlsx *.xls *.xlsm *.csv);;所有文件 (*)"
        )
        if not paths:
            return
        self.folder_label.setText(f"已选择 {len(paths)} 个文件")
        self.file_list.clear()
        self.current_files = paths
        for fp in self.current_files:
            self.file_list.addItem(os.path.basename(fp))
        self.status_badge.setText("已导入")
        self.log(f"直接选择了 {len(self.current_files)} 个文件")

    def select_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not folder:
            return
        self.output_dir = folder
        self.output_label.setText(folder)
        self.log(f"输出目录: {folder}")

    def preview_file(self, item):
        idx = self.file_list.row(item)
        if 0 <= idx < len(self.current_files):
            PreviewWindow(self.current_files[idx]).exec()

    def start_process(self, mode="merge"):
        if not self.current_files:
            self.log("请先选择文件夹或文件")
            return
        tmpl = self.template_box.currentText()
        mapping = None
        if tmpl and tmpl != "自动识别":
            mapping = self.template_manager.load_template(tmpl)
            if mapping:
                self.log(f"加载模板: {tmpl}")

        if mode == "direct":
            # 直接处理：当前列表选中的文件
            idx = self.file_list.currentRow()
            if idx < 0 or idx >= len(self.current_files):
                self.log("请在文件列表中点击选中要直接处理的文件")
                return
            target = [self.current_files[idx]]
            self.log(f"直接处理: {os.path.basename(target[0])}")
            result = merge_excel_files(target, "processed.xlsx", mapping, self.output_dir)
        else:
            # 合并处理：所有文件
            self.log(f"合并处理 {len(self.current_files)} 个文件 ...")
            result = merge_excel_files(self.current_files, "merged.xlsx", mapping, self.output_dir)

        if not result or result.get("df") is None or result["df"].empty:
            self.log("处理失败：没有可合并的数据")
            return
        self.merged_df = result["df"]
        self.status_badge.setText("处理完成")
        for line in result.get("logs", []):
            self.log(line)
        self.update_call_summary()
        self.run_call_summary()

    def update_call_summary(self):
        summary = get_data_summary(self.merged_df)
        metrics = [
            ("总记录数", "0"), ("来源文件数", "0"), ("本方号码数", "0"),
            ("对方号码数", "0"), ("关键字段完整率", "0%"), ("时间范围", "暂无"),
        ]
        if not self.metric_labels:
            for idx, (name, val) in enumerate(metrics):
                box = QFrame()
                bl = QVBoxLayout(box)
                bl.setContentsMargins(14, 8, 14, 8)
                vl = QLabel(val)
                vl.setObjectName("metricValue")
                vl.setWordWrap(True)
                nl = QLabel(name)
                nl.setObjectName("metricName")
                bl.addWidget(vl)
                bl.addWidget(nl)
                self.metric_labels[name] = vl
                self.summary_grid.addWidget(box, idx // 3, idx % 3)
        for name, label in self.metric_labels.items():
            label.setText(str(summary.get(name, "暂无")))

    def _need_data(self):
        if self.merged_df is None or self.merged_df.empty:
            self.log("请先完成数据处理")
            return False
        return True

    def run_call_summary(self):
        if not self._need_data():
            return
        s = get_data_summary(self.merged_df)
        lines = ["数据概览"]
        for k, v in s.items():
            lines.append(f"  {k}: {v}")
        top = get_global_top_contacts(self.merged_df, top_n=5)
        if not top.empty:
            lines.append("\n高频对方号码 TOP 5:")
            for _, r in top.iterrows():
                lines.append(f"  {r['对方号码']}  {int(r['次数'])}次")
        self.log("\n".join(lines))
        self.show_call_table(top)

    def run_call_top(self):
        if not self._need_data():
            return
        data = get_top_contacts_by_user(self.merged_df)
        gt = get_global_top_contacts(self.merged_df)
        if not data:
            self.log("无可分析的联系人数据")
            return
        lines = ["分本方号码高频联系人"]
        for user, tbl in data.items():
            lines.append(f"\n{user}")
            for _, r in tbl.iterrows():
                lines.append(f"  {r['对方号码']}  {int(r['次数'])}次")
        self.log("\n".join(lines))
        self.show_call_table(gt)

    def run_call_night(self):
        if not self._need_data():
            return
        tbl = get_night_calls(self.merged_df)
        self.log(f"深夜通话: 共 {len(tbl)} 条 (00:00-05:59)")
        self.show_call_table(tbl)

    def run_call_risk(self):
        if not self._need_data():
            return
        tbl = get_risk_contacts(self.merged_df)
        if tbl.empty:
            self.log("未发现达到阈值的风险关系")
            return
        lines = ["风险关系 TOP 10"]
        for _, r in tbl.head(10).iterrows():
            lines.append(
                f"{r['本方号码']} -> {r['对方号码']}  "
                f"风险分{int(r['风险分'])} "
                f"(总{int(r['总次数'])}/深夜{int(r['深夜次数'])}/短{int(r['短通话次数'])})"
            )
        self.log("\n".join(lines))
        self.show_call_table(tbl)

    def show_call_table(self, df):
        self.result_table.setSortingEnabled(False)
        self.result_table.clear()
        if df is None or df.empty:
            self.result_table.setRowCount(0)
            self.result_table.setColumnCount(0)
            self.result_table.setSortingEnabled(True)
            return
        td = df.fillna("").copy()
        self.result_table.setRowCount(len(td))
        self.result_table.setColumnCount(len(td.columns))
        self.result_table.setHorizontalHeaderLabels([str(c) for c in td.columns])
        for ri, (_, row) in enumerate(td.iterrows()):
            for ci, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.result_table.setItem(ri, ci, item)
        self.result_table.resizeColumnsToContents()
        self.result_table.setSortingEnabled(True)

    def open_mapping(self):
        if not self.current_files:
            self.log("请先选择文件夹")
            return
        try:
            import pandas as pd
            f0 = self.current_files[0]
            if f0.endswith(".csv"):
                df = pd.read_csv(f0, dtype=str, nrows=5)
            else:
                df = pd.read_excel(f0, dtype=str, nrows=5)
        except Exception as e:
            self.log(f"读取字段失败: {e}")
            return
        dlg = MappingWindow(list(df.columns))
        if dlg.exec():
            self.refresh_templates()
            self.log("字段映射模板已保存")

    def refresh_templates(self):
        self.template_box.clear()
        tmpls = self.template_manager.list_templates()
        self.template_box.addItems(["自动识别"] + tmpls)

    def log(self, msg):
        self.log_box.append(str(msg))

    # ══════════════════════════════════════════════
    # 样式
    # ══════════════════════════════════════════════

    def apply_style(self):
        self.setStyleSheet("""
            QWidget {
                background: #f0f2f5;
                color: #1f2937;
                font-family: "Microsoft YaHei", "PingFang SC", Arial;
                font-size: 13px;
            }
            QFrame#header {
                background: #fff;
                border: 1px solid #d1d5db;
                border-radius: 8px;
            }
            QLabel#title {
                font-size: 22px;
                font-weight: 700;
                color: #111827;
            }
            QLabel#subtitle {
                color: #6b7280;
                font-size: 12px;
            }
            QLabel#statusBadge {
                background: #dbeafe;
                color: #1d4ed8;
                border-radius: 6px;
                padding: 6px 14px;
                font-weight: 600;
            }
            QLabel#metricValue {
                font-size: 20px;
                font-weight: 700;
                color: #111827;
            }
            QLabel#metricName, QLabel#pathLabel {
                color: #6b7280;
                font-size: 12px;
            }
            QFrame#sidePanel, QFrame#workPanel, QFrame#summaryPanel, QGroupBox {
                background: #fff;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
            }
            QGroupBox {
                margin-top: 10px;
                padding: 10px;
                font-weight: 600;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
            QPushButton {
                background: #fff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 7px 12px;
                min-height: 18px;
            }
            QPushButton:hover {
                background: #f3f4f6;
            }
            QPushButton#primaryButton {
                background: #2563eb;
                border-color: #2563eb;
                color: #fff;
                font-weight: 600;
            }
            QPushButton#primaryButton:hover {
                background: #1d4ed8;
            }
            QComboBox, QListWidget, QTextEdit, QTableWidget, QSpinBox, QLineEdit {
                background: #fff;
                border: 1px solid #d1d5db;
                border-radius: 5px;
                padding: 5px;
                selection-background-color: #dbeafe;
            }
            QHeaderView::section {
                background: #f3f4f6;
                border: 0;
                border-bottom: 1px solid #d1d5db;
                padding: 6px;
                font-weight: 600;
            }
            QTabWidget::pane {
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                background: #fff;
            }
            QTabBar::tab {
                background: #e5e7eb;
                border: 1px solid #d1d5db;
                padding: 8px 20px;
                margin-right: 3px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background: #fff;
                color: #2563eb;
                font-weight: 600;
            }
        """)
