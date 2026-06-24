"""
银行交易分析面板 V3.1
6 个子模块：核心分析 / CSV合并 / 补全 / 行为 / 关联 / 筛选
"""

import os
import queue
import threading
from datetime import datetime

import pandas as pd

from ui.qt_compat import (
    Qt,
    QCheckBox,
    QTimer,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.bank_analyzer import (
    calculate_risk_score,
    export_person_summary,
    filter_night_trades,
    filter_quick_in_out,
    filter_sensitive_amount,
    load_bank_data,
)
from core.bank_analyzer_ext import (
    complete_transactions,
    merge_csv_folder,
)

MAX_ROWS = 500


# ══════════════════════════════════════════════════════
# 主面板
# ══════════════════════════════════════════════════════

class BankPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self.current_result = None
        self.output_dir = "银行交易分析结果"
        self._busy = False
        self._buttons = []
        self._log_widgets = []
        self._log_queue = queue.Queue()

        self.init_ui()

        self._log_timer = QTimer()
        self._log_timer.timeout.connect(self._flush_logs)
        self._log_timer.start(100)

    # ── 线程安全 ──
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
                self._log_last(f"❌ {e}")
            finally:
                self._set_busy(False)
        threading.Thread(target=w, daemon=True).start()

    def _log_last(self, msg):
        """线程安全日志。"""
        self._log_queue.put(str(msg))

    def _flush_logs(self):
        while not self._log_queue.empty():
            try:
                msg = self._log_queue.get_nowait()
                if self._log_widgets:
                    self._log_widgets[-1].append(msg)
            except queue.Empty:
                break

    # ── UI ──
    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        self.subtabs = QTabWidget()
        self.subtabs.addTab(self._tab_core(), "📊 核心分析")
        self.subtabs.addTab(self._tab_merge(), "📦 CSV合并")
        self.subtabs.addTab(self._tab_complete(), "🔧 补全交易")
        root.addWidget(self.subtabs)

    # ── 公共工具 ──
    def _ensure_dir(self):
        os.makedirs(self.output_dir, exist_ok=True)

    def _save_excel(self, df, filename):
        if df is None or df.empty:
            return
        self._ensure_dir()
        p = os.path.join(self.output_dir, filename)
        df.to_excel(p, index=False)
        self._log_last(f"📁 {p}")

    def _show_table(self, df, table):
        table.setRowCount(0)
        if df is None or df.empty:
            return
        show = df.head(MAX_ROWS)
        table.setColumnCount(len(show.columns))
        table.setHorizontalHeaderLabels([str(c) for c in show.columns])
        table.setRowCount(len(show))
        for ri, (_, row) in enumerate(show.iterrows()):
            for ci, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(ri, ci, item)
        table.resizeColumnsToContents()

    def _panel(self):
        """返回统一的 (control_frame, output_frame, table, log) 布局"""
        w = QWidget()
        l = QHBoxLayout(w)
        l.setContentsMargins(0, 0, 0, 0)
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        ctrl.setMinimumWidth(260)
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(12, 12, 12, 12)
        cl.setSpacing(10)

        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        ol.setContentsMargins(12, 12, 12, 12)
        ol.setSpacing(8)
        table = QTableWidget()
        table.setAlternatingRowColors(True)
        tlog = QTextEdit()
        tlog.setReadOnly(True)
        tlog.setMaximumHeight(130)
        self._log_widgets.append(tlog)
        ol.addWidget(table, 1)
        ol.addWidget(tlog)

        l.addWidget(ctrl, 1)
        l.addWidget(out, 4)
        return w, ctrl, cl, table, tlog

    def _make_btn(self, text, fn, primary=False):
        b = QPushButton(text)
        if primary:
            b.setObjectName("primaryButton")
        b.clicked.connect(lambda: self._run_async(fn))
        self._buttons.append(b)
        return b

    def _file_row(self, parent, label, attr):
        r = QHBoxLayout()
        r.addWidget(QLabel(label))
        e = QLineEdit()
        e.setReadOnly(True)
        setattr(self, attr, e)
        r.addWidget(e, 1)
        b = QPushButton("选")
        b.setFixedWidth(40)
        r.addWidget(b)
        parent.addLayout(r)
        return b

    # ══════════════════════════════════════════════
    # Tab 1: 核心分析
    # ══════════════════════════════════════════════

    def _tab_core(self):
        w, ctrl, cl, table, tlog = self._panel()

        # ── 数据导入 ──
        g1 = QGroupBox("导入")
        g1l = QVBoxLayout(g1)
        self.lbl_file = QLabel("未选择文件")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setObjectName("pathLabel")
        g1l.addWidget(self.lbl_file)
        bf = QPushButton("选择银行流水 Excel")
        bf.setObjectName("primaryButton")
        bf.clicked.connect(self._sel_file)
        g1l.addWidget(bf)

        g1l.addWidget(QLabel("输出目录"))
        orow = QHBoxLayout()
        self.out_label = QLabel("银行交易分析结果")
        self.out_label.setObjectName("pathLabel")
        self.out_label.setWordWrap(True)
        ob = QPushButton("选")
        ob.setFixedWidth(40)
        ob.clicked.connect(self._sel_outdir)
        orow.addWidget(self.out_label, 1)
        orow.addWidget(ob)
        g1l.addLayout(orow)
        cl.addWidget(g1)

        # ── 筛选条件（多选，可同时执行）──
        g2 = QGroupBox("筛选条件")
        g2l = QVBoxLayout(g2)

        # 快进快出 + 高频阈值
        self.chk_quick = QCheckBox("快进快出")
        self.chk_quick.setChecked(True)
        g2l.addWidget(self.chk_quick)
        hr1 = QHBoxLayout()
        hr1.addWidget(QLabel("高频阈值 >"))
        self.spin_freq = QSpinBox(); self.spin_freq.setRange(1, 99999); self.spin_freq.setValue(50)
        hr1.addWidget(self.spin_freq); hr1.addWidget(QLabel("条"))
        hr1.addStretch()
        g2l.addLayout(hr1)

        # 敏感金额
        self.chk_amount = QCheckBox("敏感金额")
        self.chk_amount.setChecked(True)
        g2l.addWidget(self.chk_amount)
        hr2 = QHBoxLayout()
        hr2.addWidget(QLabel("≥"))
        self.spin_base = QSpinBox(); self.spin_base.setRange(100, 999999); self.spin_base.setValue(500)
        hr2.addWidget(self.spin_base)
        hr2.addWidget(QLabel("且 %"))
        self.spin_step = QSpinBox(); self.spin_step.setRange(1, 99999); self.spin_step.setValue(100)
        hr2.addWidget(self.spin_step); hr2.addWidget(QLabel("=0"))
        hr2.addStretch()
        g2l.addLayout(hr2)

        # 深夜交易
        self.chk_night = QCheckBox("深夜交易")
        self.chk_night.setChecked(True)
        g2l.addWidget(self.chk_night)
        hr3 = QHBoxLayout()
        self.spin_nstart = QSpinBox(); self.spin_nstart.setRange(0, 23); self.spin_nstart.setValue(21)
        hr3.addWidget(self.spin_nstart)
        hr3.addWidget(QLabel(":00 —"))
        self.spin_nend = QSpinBox(); self.spin_nend.setRange(0, 23); self.spin_nend.setValue(5)
        hr3.addWidget(self.spin_nend); hr3.addWidget(QLabel(":00"))
        hr3.addStretch()
        g2l.addLayout(hr3)

        # 风险评分
        self.chk_risk = QCheckBox("风险评分 + 人员汇总")
        self.chk_risk.setChecked(True)
        g2l.addWidget(self.chk_risk)
        cl.addWidget(g2)

        # ── 一键执行 ──
        cl.addWidget(self._make_btn("🚀 执行筛选", self._run_core_all, True))
        cl.addStretch()

        self._table_core = table
        return w

    def _sel_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择银行流水", "", "Excel (*.xlsx *.xls *.csv)")
        if p:
            self.lbl_file.setText(p)
            try:
                self.df = load_bank_data(p)
                self.current_result = None
                self._show_table(self.df, self._table_core)
                self._log_last(f"✅ 加载: {len(self.df):,}条  |  进{(self.df['收付标志']=='进').sum():,}  出{(self.df['收付标志']=='出').sum():,}")
            except Exception as e:
                self._log_last(f"加载失败: {e}")

    def _sel_outdir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.output_dir = d
            self.out_label.setText(d)

    def _ensure_df(self):
        if self.df is None or self.df.empty:
            self._log_last("请先选择数据文件")
            return False
        return True

    def _run_core_all(self):
        """一键执行所有勾选的筛选条件。"""
        if not self._ensure_df():
            return
        src = self.df
        self._log_last("=" * 40)

        # 1. 快进快出
        if self.chk_quick.isChecked():
            r, logs = filter_quick_in_out(src, min_hits=self.spin_freq.value())
            src = r
            for l in logs: self._log_last(l)
            self._show_table(r, self._table_core)
            self._save_excel(r, "B01_快进快出.xlsx")

        # 2. 敏感金额
        if self.chk_amount.isChecked():
            r, logs = filter_sensitive_amount(src, base=self.spin_base.value(), step=self.spin_step.value())
            src = r
            for l in logs: self._log_last(l)
            self._show_table(r, self._table_core)
            self._save_excel(r, "B02_敏感金额.xlsx")

        # 3. 深夜交易
        if self.chk_night.isChecked():
            r, logs = filter_night_trades(src, start_h=self.spin_nstart.value(), end_h=self.spin_nend.value())
            src = r
            for l in logs: self._log_last(l)
            self._show_table(r, self._table_core)
            self._save_excel(r, "B03_深夜交易.xlsx")

        # 4. 风险评分
        if self.chk_risk.isChecked():
            acc, person = calculate_risk_score(src)
            self._log_last(f"风险评分: {len(acc)}账户, {len(person)}人")
            self._show_table(person, self._table_core)
            self._save_excel(acc, "B04_账户风险.xlsx")
            self._save_excel(person, "B05_人员风险.xlsx")
            # 人员名单
            p, s = export_person_summary(person, self.output_dir)
            self._log_last(f"人员名单: {p} ({len(s)}人)")

        self.current_result = src
        self._log_last("=" * 40)
        self._log_last(f"✅ 筛选完成，共 {len(src):,} 条")

    # ══════════════════════════════════════════════
    # Tab 2: CSV合并
    # ══════════════════════════════════════════════

    def _tab_merge(self):
        w, ctrl, cl, table, tlog = self._panel()
        g = QGroupBox("CSV文件夹")
        gl = QVBoxLayout(g)
        self.merge_entry = QLineEdit()
        self.merge_entry.setReadOnly(True)
        gl.addWidget(self.merge_entry)
        b = QPushButton("选择文件夹")
        b.clicked.connect(lambda: self._pick_dir(self.merge_entry))
        gl.addWidget(b)
        cl.addWidget(g)
        cl.addWidget(self._make_btn("开始合并", self._do_merge, True))
        cl.addStretch()
        self._table_merge = table
        return w

    def _do_merge(self):
        f = self.merge_entry.text()
        if not f: self._log_last("请选择文件夹"); return
        r = merge_csv_folder(f, log_callback=lambda m: self._log_last(m))
        pd.DataFrame([{"类型": k, "文件": v} for k, v in r.items()]).pipe(self._show_table, self._table_merge)

    # ══════════════════════════════════════════════
    # Tab 3: 补全
    # ══════════════════════════════════════════════

    def _tab_complete(self):
        w, ctrl, cl, table, tlog = self._panel()
        g = QGroupBox("文件选择")
        gl = QVBoxLayout(g)
        b1 = self._file_row(gl, "交易明细:", "comp_trans")
        b1.clicked.connect(lambda: self._pick_file(self.comp_trans))
        b2 = self._file_row(gl, "账户信息:", "comp_acc")
        b2.clicked.connect(lambda: self._pick_file(self.comp_acc))
        cl.addWidget(g)
        cl.addWidget(self._make_btn("开始补全", self._do_complete, True))
        cl.addStretch()
        self._table_comp = table
        return w

    def _do_complete(self):
        t = self.comp_trans.text(); a = self.comp_acc.text()
        if not t or not a: self._log_last("请选择两个文件"); return
        df, p = complete_transactions(t, a, log_callback=lambda m: self._log_last(m))
        self._show_table(df.head(200), self._table_comp)


    # ── 文件选择工具 ──
    def _pick_dir(self, entry):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if d: entry.setText(d)

    def _pick_file(self, entry):
        p, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel (*.xlsx *.xls *.csv)")
        if p: entry.setText(p)
