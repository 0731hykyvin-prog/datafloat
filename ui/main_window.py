import os

from PySide2.QtCore import Qt
from PySide2.QtWidgets import (
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QComboBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
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
from core.graph_analyzer import build_call_graph
from core.template_manager import TemplateManager
from ui.graph_widget import GraphWidget
from ui.mapping_window import MappingWindow
from ui.preview_window import PreviewWindow


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("数据处理平台 V2.0")
        self.resize(1180, 760)

        self.current_files = []
        self.merged_df = None
        self.template_manager = TemplateManager()
        self.metric_labels = {}

        self.init_ui()
        self.apply_style()
        self.refresh_templates()
        self.update_summary()

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        header = self.create_header()
        root.addWidget(header)

        self.summary_grid = QGridLayout()
        summary_panel = QFrame()
        summary_panel.setObjectName("summaryPanel")
        summary_panel.setLayout(self.summary_grid)
        root.addWidget(summary_panel)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self.create_sidebar())
        splitter.addWidget(self.create_workspace())
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 5)
        root.addWidget(splitter, 1)

    def create_header(self):
        frame = QFrame()
        frame.setObjectName("header")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(18, 16, 18, 16)

        title_box = QVBoxLayout()
        title = QLabel("数据处理与通信分析平台")
        title.setObjectName("title")
        subtitle = QLabel("批量合并话单数据，快速定位高频联系人、深夜通话与风险关系。")
        subtitle.setObjectName("subtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        self.status_badge = QLabel("等待导入")
        self.status_badge.setObjectName("statusBadge")
        self.status_badge.setAlignment(Qt.AlignCenter)

        layout.addLayout(title_box, 1)
        layout.addWidget(self.status_badge)
        return frame

    def create_sidebar(self):
        panel = QFrame()
        panel.setObjectName("sidePanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        source_group = QGroupBox("数据源")
        source_layout = QVBoxLayout(source_group)

        self.folder_label = QLabel("未选择文件夹")
        self.folder_label.setObjectName("pathLabel")
        self.folder_label.setWordWrap(True)

        self.btn_select = QPushButton("选择文件夹")
        self.btn_select.setObjectName("primaryButton")

        self.file_list = QListWidget()
        self.file_list.setAlternatingRowColors(True)

        source_layout.addWidget(self.folder_label)
        source_layout.addWidget(self.btn_select)
        source_layout.addWidget(QLabel("文件列表"))
        source_layout.addWidget(self.file_list, 1)

        process_group = QGroupBox("处理流程")
        process_layout = QVBoxLayout(process_group)

        self.template_box = QComboBox()
        self.btn_mapping = QPushButton("字段映射")
        self.btn_run = QPushButton("开始处理")
        self.btn_run.setObjectName("primaryButton")

        process_layout.addWidget(QLabel("数据模板"))
        process_layout.addWidget(self.template_box)
        process_layout.addWidget(self.btn_mapping)
        process_layout.addWidget(self.btn_run)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(140)
        self.log_box.setPlaceholderText("运行日志会显示在这里")

        layout.addWidget(source_group, 3)
        layout.addWidget(process_group)
        layout.addWidget(QLabel("运行日志"))
        layout.addWidget(self.log_box, 1)

        self.btn_select.clicked.connect(self.select_folder)
        self.btn_run.clicked.connect(self.start_process)
        self.btn_mapping.clicked.connect(self.open_mapping)
        self.file_list.itemDoubleClicked.connect(self.preview_file)

        return panel

    def create_workspace(self):
        panel = QFrame()
        panel.setObjectName("workPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        action_bar = QHBoxLayout()
        self.btn_summary = QPushButton("刷新概览")
        self.btn_top_contacts = QPushButton("高频联系人")
        self.btn_night_calls = QPushButton("深夜通话")
        self.btn_risk = QPushButton("风险关系")
        self.btn_graph = QPushButton("通信图谱")

        for btn in [
            self.btn_summary,
            self.btn_top_contacts,
            self.btn_night_calls,
            self.btn_risk,
            self.btn_graph,
        ]:
            action_bar.addWidget(btn)
        action_bar.addStretch(1)

        self.tabs = QTabWidget()
        self.analysis_box = QTextEdit()
        self.analysis_box.setReadOnly(True)
        self.analysis_box.setPlaceholderText("处理数据后，可在这里查看分析结论。")

        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setSortingEnabled(True)

        self.graph_widget = GraphWidget()

        self.tabs.addTab(self.analysis_box, "分析报告")
        self.tabs.addTab(self.result_table, "明细表")
        self.tabs.addTab(self.graph_widget, "关系图谱")

        layout.addLayout(action_bar)
        layout.addWidget(self.tabs, 1)

        self.btn_summary.clicked.connect(self.run_summary)
        self.btn_top_contacts.clicked.connect(self.run_top_contacts)
        self.btn_night_calls.clicked.connect(self.run_night_calls)
        self.btn_risk.clicked.connect(self.run_risk_contacts)
        self.btn_graph.clicked.connect(self.run_graph_analysis)

        return panel

    def apply_style(self):
        self.setStyleSheet(
            """
            QWidget {
                background: #f5f7fb;
                color: #1f2937;
                font-family: "PingFang SC", "Microsoft YaHei", Arial;
                font-size: 14px;
            }
            QFrame#header {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
            }
            QLabel#title {
                font-size: 24px;
                font-weight: 700;
                color: #111827;
            }
            QLabel#subtitle {
                color: #6b7280;
                margin-top: 4px;
            }
            QLabel#statusBadge {
                background: #eef6ff;
                color: #1d4ed8;
                border: 1px solid #bfdbfe;
                border-radius: 8px;
                padding: 8px 16px;
                font-weight: 600;
                min-width: 92px;
            }
            QFrame#summaryPanel,
            QFrame#sidePanel,
            QFrame#workPanel,
            QGroupBox {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
            }
            QGroupBox {
                margin-top: 12px;
                padding: 12px;
                font-weight: 600;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #374151;
            }
            QLabel#metricValue {
                font-size: 22px;
                font-weight: 700;
                color: #111827;
            }
            QLabel#metricName,
            QLabel#pathLabel {
                color: #6b7280;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 7px;
                padding: 8px 12px;
                min-height: 20px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
            QPushButton#primaryButton {
                background: #2563eb;
                border-color: #2563eb;
                color: white;
                font-weight: 600;
            }
            QPushButton#primaryButton:hover {
                background: #1d4ed8;
            }
            QComboBox,
            QListWidget,
            QTextEdit,
            QTableWidget {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 7px;
                padding: 6px;
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
            QTabWidget::pane {
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                background: #ffffff;
            }
            QTabBar::tab {
                background: #eef2f7;
                border: 1px solid #d1d5db;
                padding: 8px 16px;
                margin-right: 4px;
                border-top-left-radius: 7px;
                border-top-right-radius: 7px;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #2563eb;
                font-weight: 600;
            }
            """
        )

    def refresh_templates(self):
        self.template_box.clear()
        templates = self.template_manager.list_templates()
        self.template_box.addItems(["自动识别"] + templates)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if not folder:
            return

        self.folder_label.setText(folder)
        self.file_list.clear()
        self.current_files = scan_files(folder)

        for file_path in self.current_files:
            self.file_list.addItem(os.path.basename(file_path))

        self.status_badge.setText("已导入")
        self.log(f"扫描到 {len(self.current_files)} 个文件")
        self.update_summary()

    def preview_file(self, item):
        index = self.file_list.row(item)
        if index < 0 or index >= len(self.current_files):
            return

        dlg = PreviewWindow(self.current_files[index])
        dlg.exec()

    def start_process(self):
        if not self.current_files:
            self.log("请先选择文件夹")
            return

        template_name = self.template_box.currentText()
        mapping = None

        if template_name and template_name != "自动识别":
            mapping = self.template_manager.load_template(template_name)
            if mapping:
                self.log(f"加载模板成功：{template_name}")

        result = merge_excel_files(self.current_files, "merged.xlsx", mapping)
        if not result or result.get("df") is None or result["df"].empty:
            if result and result.get("logs"):
                for line in result["logs"]:
                    self.log(line)
            else:
                self.log("处理失败：没有可合并的数据")
            return

        self.merged_df = result["df"]
        self.status_badge.setText("处理完成")

        # 输出日志到 UI
        for line in result.get("logs", []):
            self.log(line)
        self.update_summary()
        self.run_summary()

    def update_summary(self):
        summary = get_data_summary(self.merged_df)
        if not self.metric_labels:
            metrics = [
                ("总记录数", "0"),
                ("来源文件数", "0"),
                ("本方号码数", "0"),
                ("对方号码数", "0"),
                ("关键字段完整率", "0%"),
                ("时间范围", "暂无"),
            ]
            for index, (name, value) in enumerate(metrics):
                box = QFrame()
                box_layout = QVBoxLayout(box)
                box_layout.setContentsMargins(14, 10, 14, 10)
                value_label = QLabel(value)
                value_label.setObjectName("metricValue")
                name_label = QLabel(name)
                name_label.setObjectName("metricName")
                value_label.setWordWrap(True)
                box_layout.addWidget(value_label)
                box_layout.addWidget(name_label)
                self.metric_labels[name] = value_label
                self.summary_grid.addWidget(box, index // 3, index % 3)

        for name, label in self.metric_labels.items():
            label.setText(str(summary.get(name, "暂无")))

    def require_data(self):
        if self.merged_df is None or self.merged_df.empty:
            self.analysis_box.setText("请先选择文件夹并完成数据处理。")
            self.tabs.setCurrentWidget(self.analysis_box)
            return False
        return True

    def run_summary(self):
        if not self.require_data():
            return

        summary = get_data_summary(self.merged_df)
        lines = ["========== 数据概览 =========="]
        for key, value in summary.items():
            lines.append(f"{key}: {value}")

        top_contacts = get_global_top_contacts(self.merged_df, top_n=5)
        if not top_contacts.empty:
            lines.append("\n========== 高频对方号码 TOP 5 ==========")
            for _, row in top_contacts.iterrows():
                lines.append(
                    f"{row['对方号码']}  {int(row['次数'])}次，涉及{int(row['涉及本方号码数'])}个本方号码"
                )

        self.analysis_box.setText("\n".join(lines))
        self.show_table(top_contacts)
        self.tabs.setCurrentWidget(self.analysis_box)

    def run_top_contacts(self):
        if not self.require_data():
            return

        data = get_top_contacts_by_user(self.merged_df)
        global_top = get_global_top_contacts(self.merged_df)
        if not data:
            self.analysis_box.setText("没有可分析的联系人数据。")
            self.show_table(global_top)
            return

        lines = ["========== 分本方号码高频联系人 =========="]
        for user, table in data.items():
            lines.append(f"\n{user}")
            for _, row in table.iterrows():
                lines.append(f"  {row['对方号码']}  {int(row['次数'])}次")

        self.analysis_box.setText("\n".join(lines))
        self.show_table(global_top)
        self.tabs.setCurrentWidget(self.analysis_box)

    def run_night_calls(self):
        if not self.require_data():
            return

        table = get_night_calls(self.merged_df)
        self.analysis_box.setText(
            f"========== 深夜通话 ==========\n共发现 {len(table)} 条 00:00-05:59 通话记录，明细见表格页。"
        )
        self.show_table(table)
        self.tabs.setCurrentWidget(self.result_table)

    def run_risk_contacts(self):
        if not self.require_data():
            return

        table = get_risk_contacts(self.merged_df)
        if table.empty:
            self.analysis_box.setText("暂未发现达到阈值的风险关系。")
        else:
            lines = ["========== 风险关系 TOP 10 =========="]
            for _, row in table.head(10).iterrows():
                lines.append(
                    f"{row['本方号码']} -> {row['对方号码']}  风险分{int(row['风险分'])} "
                    f"(总{int(row['总次数'])}次 / 深夜{int(row['深夜次数'])}次 / 短通话{int(row['短通话次数'])}次)"
                )
            self.analysis_box.setText("\n".join(lines))

        self.show_table(table)
        self.tabs.setCurrentWidget(self.analysis_box)

    def run_graph_analysis(self):
        if not self.require_data():
            return

        graph_df = build_call_graph(self.merged_df)
        if graph_df is None or graph_df.empty:
            self.analysis_box.setText("无可绘制的通信关系数据。")
            return

        graph_df = graph_df.sort_values("次数", ascending=False).head(50)
        self.graph_widget.draw_graph(graph_df)
        self.show_table(graph_df)
        self.analysis_box.setText(f"通信关系图谱已生成，共展示前 {len(graph_df)} 条高频关系。")
        self.tabs.setCurrentWidget(self.graph_widget)

    def show_table(self, df):
        self.result_table.setSortingEnabled(False)
        self.result_table.clear()

        if df is None or df.empty:
            self.result_table.setRowCount(0)
            self.result_table.setColumnCount(0)
            self.result_table.setSortingEnabled(True)
            return

        table_df = df.fillna("").copy()
        self.result_table.setRowCount(len(table_df))
        self.result_table.setColumnCount(len(table_df.columns))
        self.result_table.setHorizontalHeaderLabels([str(c) for c in table_df.columns])

        for row_index, (_, row) in enumerate(table_df.iterrows()):
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.result_table.setItem(row_index, col_index, item)

        self.result_table.resizeColumnsToContents()
        self.result_table.setSortingEnabled(True)

    def open_mapping(self):
        if not self.current_files:
            self.log("请先选择文件夹")
            return

        try:
            import pandas as pd

            first_file = self.current_files[0]
            if first_file.endswith(".csv"):
                df = pd.read_csv(first_file, dtype=str, nrows=5)
            else:
                df = pd.read_excel(first_file, dtype=str, nrows=5)
        except Exception as exc:
            QMessageBox.warning(self, "字段映射", f"读取字段失败：{exc}")
            return

        dlg = MappingWindow(list(df.columns))
        if dlg.exec():
            self.refresh_templates()
            self.log("字段映射模板已更新")

    def log(self, msg):
        self.log_box.append(str(msg))
