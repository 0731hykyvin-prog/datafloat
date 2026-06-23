"""
Datafloat V3.0 主窗口
TAB 1: 话单分析（通信记录合并+分析）
TAB 2: 银行交易分析（流水快进快出+风险评分）
"""

import os
import threading

import pandas as pd

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
    QMessageBox,
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
    get_night_calls_by_user,
    get_night_calls_summary,
    get_risk_by_user,
    get_risk_contacts,
    get_top_contacts_by_user,
)
from core.excel_engine import merge_excel_files
from core.file_loader import scan_files
from core.template_manager import TemplateManager
from ui.bank_panel import BankPanel
from ui.bank_tools_panel import BankToolsPanel
from ui.mapping_window import MappingWindow
from ui.preview_window import PreviewWindow
from ui.video_panel import VideoPanel

MAX_TABLE_ROWS = 500  # 表格最多显示行数，保证UI流畅


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Datafloat 数据处理平台 V3.0")
        self.resize(1100, 700)

        self.current_files = []
        self.merged_df = None
        self.template_manager = TemplateManager()
        self.metric_labels = {}
        self._busy = False

        self.init_ui()
        self.apply_style()
        self.refresh_templates()
        self.update_call_summary()

    # ══════════════════════════════════════════════
    # 线程安全 — 防止重复点击 + UI 反馈
    # ══════════════════════════════════════════════

    def _set_busy(self, busy):
        """设置忙碌状态：禁用按钮，改鼠标指针。"""
        self._busy = busy
        self.setCursor(Qt.WaitCursor if busy else Qt.ArrowCursor)
        for attr in ['btn_summary', 'btn_top', 'btn_night', 'btn_risk']:
            btn = getattr(self, attr, None)
            if btn:
                btn.setEnabled(not busy)

    def _run_async(self, func, *args, **kwargs):
        """在后台线程执行 func，执行期间禁用按钮。"""
        if self._busy:
            return
        self._set_busy(True)
        def wrapper():
            try:
                func(*args, **kwargs)
            except Exception as e:
                self.log(f"❌ 错误: {e}")
            finally:
                self._set_busy(False)
        threading.Thread(target=wrapper, daemon=True).start()

    # ══════════════════════════════════════════════
    # UI 布局
    # ══════════════════════════════════════════════

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)
        root.addWidget(self._build_header())
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_call_tab(), "📞 话单分析")
        self.tabs.addTab(BankPanel(), "🏦 银行交易分析")
        self.tabs.addTab(BankToolsPanel(), "🔧 银行工具集")
        self.tabs.addTab(VideoPanel(), "📹 视频布控")
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
    # TAB 1: 话单分析 — 侧边栏
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
        self.folder_label = QLabel("未选择文件")
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
        btn_merge.clicked.connect(lambda: self._run_async(self.start_process, "merge"))
        btn_direct.clicked.connect(lambda: self._run_async(self.start_process, "direct"))

        layout.addWidget(src, 3)
        layout.addWidget(proc)
        return panel

    # ══════════════════════════════════════════════
    # TAB 1: 话单分析 — 工作区
    # ══════════════════════════════════════════════

    def _build_call_workspace(self):
        panel = QFrame()
        panel.setObjectName("workPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        self.summary_grid = QGridLayout()
        summary_frame = QFrame()
        summary_frame.setObjectName("summaryPanel")
        summary_frame.setLayout(self.summary_grid)
        layout.addWidget(summary_frame)

        bar = QHBoxLayout()
        self.btn_summary = QPushButton("刷新概览（高频联系人）")
        self.btn_top = QPushButton("高频联系人明细")
        self.btn_night = QPushButton("深夜通话分析")
        self.btn_risk = QPushButton("风险关系分析")
        for b in [self.btn_summary, self.btn_top, self.btn_night, self.btn_risk]:
            bar.addWidget(b)
        bar.addStretch()
        layout.addLayout(bar)

        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setSortingEnabled(True)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("运行日志")
        self.log_box.setMaximumHeight(140)

        layout.addWidget(self.result_table, 1)
        layout.addWidget(self.log_box)

        self.btn_summary.clicked.connect(lambda: self._run_async(self.run_call_summary))
        self.btn_top.clicked.connect(lambda: self._run_async(self.run_call_top))
        self.btn_night.clicked.connect(lambda: self._run_async(self.run_call_night))
        self.btn_risk.clicked.connect(lambda: self._run_async(self.run_call_risk))
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

        self.status_badge.setText("处理中...")
        if mode == "direct":
            idx = self.file_list.currentRow()
            if idx < 0 or idx >= len(self.current_files):
                self.log("请在文件列表中点击选中要直接处理的文件")
                self.status_badge.setText("就绪")
                return
            target = [self.current_files[idx]]
            self.log(f"直接处理: {os.path.basename(target[0])}")
            result = merge_excel_files(target, "processed.xlsx", mapping, self.output_dir)
        else:
            self.log(f"合并处理 {len(self.current_files)} 个文件 ...")
            result = merge_excel_files(self.current_files, "merged.xlsx", mapping, self.output_dir)

        if not result or result.get("df") is None or result["df"].empty:
            self.log("处理失败：没有可合并的数据")
            self.status_badge.setText("就绪")
            return
        self.merged_df = result["df"]
        self.status_badge.setText("处理完成")
        for line in result.get("logs", []):
            self.log(line)
        self.update_call_summary()
        self.run_call_summary()
        self.status_badge.setText("就绪")

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

    def _get_out_dir(self):
        return self.output_dir if self.output_dir else os.getcwd()

    def _save_excel(self, df, filename):
        if df is None or df.empty:
            return
        out_dir = self._get_out_dir()
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, filename)
        df.to_excel(path, index=False)
        self.log(f"📁 已导出: {path}")

    # ══════════════════════════════════════════════
    # 快速表格渲染（限制行数保证流畅）
    # ══════════════════════════════════════════════

    def show_call_table(self, df):
        self.result_table.setSortingEnabled(False)
        self.result_table.clear()
        if df is None or df.empty:
            self.result_table.setRowCount(0)
            self.result_table.setColumnCount(0)
            self.result_table.setSortingEnabled(True)
            return
        # 限制显示行数
        show = df.head(MAX_TABLE_ROWS)
        td = show.fillna("").copy()
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
        total = len(df)
        if total > MAX_TABLE_ROWS:
            self.log(f"(表格显示前 {MAX_TABLE_ROWS} 行，共 {total} 行)")

    def open_mapping(self):
        if not self.current_files:
            self.log("请先选择文件夹")
            return
        try:
            f0 = self.current_files[0]
            if f0.endswith(".csv"):
                df = pd.read_csv(f0, dtype=str, nrows=5)
            else:
                df = pd.read_excel(f0, dtype=str, nrows=5)
        except Exception as e:
            self.log(f"读取字段失败: {e}")
            return
        MappingWindow(list(df.columns)).exec()
        self.refresh_templates()
        self.log("字段映射模板已保存")

    def refresh_templates(self):
        self.template_box.clear()
        tmpls = self.template_manager.list_templates()
        self.template_box.addItems(["自动识别"] + tmpls)

    def log(self, msg):
        self.log_box.append(str(msg))

    # ══════════════════════════════════════════════
    # 分析方法（在后台线程中执行）
    # ══════════════════════════════════════════════

    def run_call_summary(self):
        if not self._need_data():
            return
        s = get_data_summary(self.merged_df)
        user_contacts = get_top_contacts_by_user(self.merged_df, top_n=10)
        lines = ["══ 数据概览 ══"]
        for k, v in s.items():
            lines.append(f"  {k}: {v}")
        rows = []
        if user_contacts:
            lines.append("\n══ 各方号码高频联系人 ══")
            for user, tbl in user_contacts.items():
                lines.append(f"\n【{user}】")
                for _, r in tbl.iterrows():
                    lines.append(f"  → {r['对方号码']}   {int(r['次数'])}次")
                tbl_copy = tbl.copy()
                tbl_copy.insert(0, "本方号码", user)
                rows.append(tbl_copy)
        self.log("\n".join(lines))
        all_data = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        self.show_call_table(all_data)
        self._save_excel(all_data, "01_高频联系人_按号码.xlsx")

    def run_call_top(self):
        if not self._need_data():
            return
        data = get_top_contacts_by_user(self.merged_df, top_n=20)
        if not data:
            self.log("无可分析的联系人数据")
            return
        lines = ["══ 高频联系人明细 TOP20 ══"]
        rows = []
        for user, tbl in data.items():
            lines.append(f"\n【{user}】 共联系 {len(tbl)} 个号码")
            for _, r in tbl.iterrows():
                lines.append(f"  {r['对方号码']}   {int(r['次数'])}次")
            tbl_copy = tbl.copy()
            tbl_copy.insert(0, "本方号码", user)
            rows.append(tbl_copy)
        self.log("\n".join(lines))
        all_data = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        self.show_call_table(all_data)
        self._save_excel(all_data, "02_高频联系人_明细.xlsx")

    def run_call_night(self):
        if not self._need_data():
            return
        night_by_user = get_night_calls_by_user(self.merged_df)
        summary = get_night_calls_summary(self.merged_df)
        if not night_by_user:
            self.log("无深夜通话记录 (00:00-05:59)")
            return
        lines = ["══ 深夜通话分析 (00:00-05:59) ══"]
        rows = []
        for user, tbl in night_by_user.items():
            cnt = len(tbl)
            lines.append(f"\n【{user}】 深夜通话 {cnt} 条")
            for _, r in tbl.head(20).iterrows():
                lines.append(f"  {r['对方号码']}  {r.get('开始时间','')}")
            tbl_copy = tbl.copy()
            tbl_copy.insert(0, "本方号码", user)
            rows.append(tbl_copy)
        self.log("\n".join(lines))
        all_data = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        self.show_call_table(all_data)
        self._save_excel(all_data, "03_深夜通话_按号码.xlsx")
        if not summary.empty:
            self._save_excel(summary, "03_深夜通话_汇总.xlsx")

    def run_call_risk(self):
        if not self._need_data():
            return
        risk_by_user = get_risk_by_user(self.merged_df)
        if not risk_by_user:
            self.log("未发现达到阈值的风险关系")
            return
        lines = ["══ 风险关系分析 ══"]
        rows = []
        for user, tbl in risk_by_user.items():
            lines.append(f"\n【{user}】")
            for _, r in tbl.head(15).iterrows():
                lines.append(
                    f"  → {r['对方号码']}  "
                    f"风险分{int(r['风险分'])} "
                    f"(总{int(r['总次数'])}/深夜{int(r['深夜次数'])}/短{int(r['短通话次数'])})"
                )
            rows.append(tbl)
        self.log("\n".join(lines))
        all_data = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        self.show_call_table(all_data)
        self._save_excel(all_data, "04_风险关系_按号码.xlsx")

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
            QPushButton:disabled {
                background: #e5e7eb;
                color: #9ca3af;
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
            QPushButton#primaryButton:disabled {
                background: #93c5fd;
                border-color: #93c5fd;
                color: #fff;
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
