"""
银行交易分析 UI 面板
"""

import os

import pandas as pd
from ui.qt_compat import (
    Qt,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.bank_analyzer import (
    calculate_risk_score,
    export_person_summary,
    filter_atm_withdrawals,
    filter_delivery_trades,
    filter_night_trades,
    filter_quick_in_out,
    filter_sensitive_amount,
    load_bank_data,
)


class BankPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self.current_result = None
        self.init_ui()

    def init_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_control())
        splitter.addWidget(self._build_output())
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 5)
        root.addWidget(splitter)

    def _build_control(self):
        panel = QFrame()
        panel.setObjectName("sidePanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        # ── 文件选择 ──
        file_group = QGroupBox("数据导入")
        fg = QVBoxLayout(file_group)
        self.lbl_file = QLabel("未选择文件")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setObjectName("pathLabel")
        btn_file = QPushButton("选择银行流水 Excel")
        btn_file.clicked.connect(self.select_file)
        fg.addWidget(self.lbl_file)
        fg.addWidget(btn_file)

        # ── 筛选参数 ──
        filter_group = QGroupBox("筛选条件")
        fl = QVBoxLayout(filter_group)

        # 敏感金额
        fl.addWidget(QLabel("敏感金额:"))
        amt_row = QHBoxLayout()
        self.spin_base = QSpinBox()
        self.spin_base.setRange(100, 999999)
        self.spin_base.setValue(500)
        self.spin_step = QSpinBox()
        self.spin_step.setRange(1, 99999)
        self.spin_step.setValue(100)
        amt_row.addWidget(QLabel("≥"))
        amt_row.addWidget(self.spin_base)
        amt_row.addWidget(QLabel("且 %"))
        amt_row.addWidget(self.spin_step)
        amt_row.addWidget(QLabel("=0"))
        amt_row.addStretch()
        fl.addLayout(amt_row)

        # 高频阈值
        hl = QHBoxLayout()
        hl.addWidget(QLabel("高频阈值(>"))
        self.spin_freq = QSpinBox()
        self.spin_freq.setRange(1, 99999)
        self.spin_freq.setValue(50)
        hl.addWidget(self.spin_freq)
        hl.addWidget(QLabel("条)"))
        hl.addStretch()

        # 深夜时间
        hl2 = QHBoxLayout()
        hl2.addWidget(QLabel("深夜:"))
        self.spin_night_start = QSpinBox()
        self.spin_night_start.setRange(0, 23)
        self.spin_night_start.setValue(21)
        hl2.addWidget(self.spin_night_start)
        hl2.addWidget(QLabel(":00 -"))
        self.spin_night_end = QSpinBox()
        self.spin_night_end.setRange(0, 23)
        self.spin_night_end.setValue(5)
        hl2.addWidget(self.spin_night_end)
        hl2.addWidget(QLabel(":00"))
        hl2.addStretch()

        fl.addLayout(hl)
        fl.addLayout(hl2)

        # ── 筛选按钮 ──
        act_group = QGroupBox("执行分析")
        al = QVBoxLayout(act_group)
        self.btn_quick = QPushButton("快进快出筛选")
        self.btn_quick.setObjectName("primaryButton")
        self.btn_sensitive = QPushButton("敏感金额筛选")
        self.btn_night = QPushButton("深夜交易筛选")
        self.btn_atm = QPushButton("ATM取款")
        self.btn_delivery = QPushButton("配送交易")
        self.btn_risk = QPushButton("风险评分汇总")
        self.btn_risk.setObjectName("primaryButton")
        self.btn_export = QPushButton("导出人员名单")

        for b in [self.btn_quick, self.btn_sensitive, self.btn_night,
                   self.btn_atm, self.btn_delivery, self.btn_risk, self.btn_export]:
            al.addWidget(b)

        self.btn_quick.clicked.connect(self.run_quick_in_out)
        self.btn_sensitive.clicked.connect(self.run_sensitive)
        self.btn_night.clicked.connect(self.run_night)
        self.btn_atm.clicked.connect(self.run_atm)
        self.btn_delivery.clicked.connect(self.run_delivery)
        self.btn_risk.clicked.connect(self.run_risk_score)
        self.btn_export.clicked.connect(self.run_export_persons)

        layout.addWidget(file_group)
        layout.addWidget(filter_group)
        layout.addWidget(act_group)
        layout.addStretch()
        return panel

    def _build_output(self):
        panel = QFrame()
        panel.setObjectName("workPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("分析结果会显示在这里")
        self.log_box.setMaximumHeight(160)

        layout.addWidget(self.table, 1)
        layout.addWidget(self.log_box)
        return panel

    # ══════════════════════════════════════════════
    # 操作
    # ══════════════════════════════════════════════

    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择银行流水 Excel", "", "Excel (*.xlsx *.xls *.csv)")
        if not path:
            return
        self.lbl_file.setText(path)
        self.log("正在加载数据...")
        try:
            self.df = load_bank_data(path)
            self.log(f"加载完成: {len(self.df):,} 条")
            self.log(f"收付标志: 进={(self.df['收付标志']=='进').sum():,}, 出={(self.df['收付标志']=='出').sum():,}")
        except Exception as e:
            self.log(f"加载失败: {e}")

    def _ensure_data(self):
        if self.df is None or self.df.empty:
            self.log("请先选择数据文件")
            return False
        return True

    def _show_result(self, df, logs):
        for line in logs:
            self.log(line)
        self.show_table(df)
        self.current_result = df

    def run_quick_in_out(self):
        if not self._ensure_data():
            return
        threshold = self.spin_freq.value()
        result, logs = filter_quick_in_out(self.df, min_hits=threshold)
        self._show_result(result, logs)

    def run_sensitive(self):
        if not self._ensure_data():
            return
        base = self.spin_base.value()
        step = self.spin_step.value()
        src = self.current_result if self.current_result is not None else self.df
        result, logs = filter_sensitive_amount(src, base=base, step=step)
        self._show_result(result, logs)

    def run_night(self):
        if not self._ensure_data():
            return
        sh = self.spin_night_start.value()
        eh = self.spin_night_end.value()
        src = self.current_result if self.current_result is not None else self.df
        result, logs = filter_night_trades(src, start_h=sh, end_h=eh)
        self._show_result(result, logs)

    def run_atm(self):
        if not self._ensure_data():
            return
        src = self.current_result if self.current_result is not None else self.df
        result, logs = filter_atm_withdrawals(src)
        self._show_result(result, logs)

    def run_delivery(self):
        if not self._ensure_data():
            return
        src = self.current_result if self.current_result is not None else self.df
        result, logs = filter_delivery_trades(src)
        self._show_result(result, logs)

    def run_risk_score(self):
        if not self._ensure_data():
            return
        src = self.current_result if self.current_result is not None else self.df
        acc_stats, person_stats = calculate_risk_score(src)
        self.log(f"风险评分: {len(acc_stats)} 账户, {len(person_stats)} 人")
        self.show_table(person_stats)
        self.current_result = person_stats

    def run_export_persons(self):
        if not self._ensure_data():
            return
        src = self.current_result if self.current_result is not None else self.df
        path, summary = export_person_summary(src)
        self.log(f"人员名单已导出: {path}")
        self.log(f"共 {len(summary)} 人")
        self.show_table(summary)

    # ══════════════════════════════════════════════
    # 工具
    # ══════════════════════════════════════════════

    def show_table(self, df):
        self.table.setSortingEnabled(False)
        self.table.clear()
        if df is None or df.empty:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.table.setSortingEnabled(True)
            return
        td = df.fillna("").copy()
        self.table.setRowCount(len(td))
        self.table.setColumnCount(len(td.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in td.columns])
        for ri, (_, row) in enumerate(td.iterrows()):
            for ci, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(ri, ci, item)
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

    def log(self, msg):
        self.log_box.append(str(msg))
