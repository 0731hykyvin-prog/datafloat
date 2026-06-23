"""
银行交易工具集
CSV合并 / 补全 / 行为分析 / 关联分析 / 高级筛选
"""

import os
import threading

from ui.qt_compat import (
    Qt,
    QComboBox,
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

from core.bank_analyzer_ext import (
    advanced_filter,
    analyze_account_behavior,
    analyze_relations,
    complete_transactions,
    merge_csv_folder,
)


class BankToolsPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._busy = False
        self.dfs = {}  # 缓存各步骤的 DataFrame
        self.init_ui()

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
                self._log(f"❌ {e}")
            finally:
                self._set_busy(False)
        threading.Thread(target=w, daemon=True).start()

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        self._buttons = []
        self._log_widgets = []

        subtabs = QTabWidget()
        subtabs.addTab(self._tab_merge(), "合并CSV")
        subtabs.addTab(self._tab_complete(), "补全交易")
        subtabs.addTab(self._tab_behavior(), "行为分析")
        subtabs.addTab(self._tab_relation(), "关联分析")
        subtabs.addTab(self._tab_filter(), "高级筛选")
        root.addWidget(subtabs)

    def _log(self, msg, widget=None):
        w = widget or self._log_widgets[-1] if self._log_widgets else None
        if w:
            w.append(str(msg))

    def _show_table(self, df, table):
        table.setRowCount(0)
        if df is None or df.empty:
            return
        show = df.head(300)
        table.setColumnCount(len(show.columns))
        table.setHorizontalHeaderLabels([str(c) for c in show.columns])
        table.setRowCount(len(show))
        for ri, (_, row) in enumerate(show.iterrows()):
            for ci, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(ri, ci, item)
        table.resizeColumnsToContents()

    def _file_row(self, parent, label, var_attr):
        row = QHBoxLayout()
        row.addWidget(QLabel(label))
        entry = QLineEdit()
        entry.setReadOnly(True)
        setattr(self, var_attr, entry)
        row.addWidget(entry, 1)
        btn = QPushButton("选择")
        btn.setFixedWidth(60)
        row.addWidget(btn)
        parent.addLayout(row)
        return btn

    # ══════════════════════════════════════════════
    # Tab 1: 合并 CSV
    # ══════════════════════════════════════════════

    def _tab_merge(self):
        w = QWidget()
        l = QHBoxLayout(w)
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(12)

        g = QGroupBox("CSV 文件夹")
        gl = QVBoxLayout(g)
        self.merge_folder_entry = QLineEdit()
        self.merge_folder_entry.setReadOnly(True)
        gl.addWidget(self.merge_folder_entry)
        b = QPushButton("选择文件夹")
        b.clicked.connect(lambda: self._pick_folder(self.merge_folder_entry))
        self._buttons.append(b)
        gl.addWidget(b)
        cl.addWidget(g)

        btn = QPushButton("开始合并")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(lambda: self._run_async(self._do_merge))
        self._buttons.append(btn)
        cl.addWidget(btn)
        cl.addStretch()

        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        self.merge_table = QTableWidget()
        self.merge_log = QTextEdit()
        self.merge_log.setReadOnly(True)
        self._log_widgets.append(self.merge_log)
        ol.addWidget(self.merge_table, 1)
        ol.addWidget(self.merge_log)
        ol.addWidget(QLabel("合并的文件将输出到 CSV 所在文件夹"))

        l.addWidget(ctrl, 1)
        l.addWidget(out, 4)
        return w

    def _do_merge(self):
        folder = self.merge_folder_entry.text()
        if not folder:
            self._log("请先选择文件夹", self.merge_log)
            return
        self.merge_log.clear()
        result = merge_csv_folder(folder, log_callback=lambda m: self._log(m, self.merge_log))
        self.dfs["merged"] = result

    # ══════════════════════════════════════════════
    # Tab 2: 补全
    # ══════════════════════════════════════════════

    def _tab_complete(self):
        w = QWidget()
        l = QHBoxLayout(w)
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(12)

        g = QGroupBox("文件选择")
        gl = QVBoxLayout(g)
        b1 = self._file_row(gl, "交易明细:", "complete_trans_entry")
        b1.clicked.connect(lambda: self._pick_file(self.complete_trans_entry))
        self._buttons.append(b1)
        b2 = self._file_row(gl, "账户信息:", "complete_acc_entry")
        b2.clicked.connect(lambda: self._pick_file(self.complete_acc_entry))
        self._buttons.append(b2)
        cl.addWidget(g)

        btn = QPushButton("开始补全")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(lambda: self._run_async(self._do_complete))
        self._buttons.append(btn)
        cl.addWidget(btn)
        cl.addStretch()

        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        self.complete_table = QTableWidget()
        self.complete_log = QTextEdit()
        self.complete_log.setReadOnly(True)
        self._log_widgets.append(self.complete_log)
        ol.addWidget(self.complete_table, 1)
        ol.addWidget(self.complete_log)

        l.addWidget(ctrl, 1)
        l.addWidget(out, 4)
        return w

    def _do_complete(self):
        t = self.complete_trans_entry.text()
        a = self.complete_acc_entry.text()
        if not t or not a:
            self._log("请选择两个文件", self.complete_log)
            return
        self.complete_log.clear()
        df, path = complete_transactions(t, a, log_callback=lambda m: self._log(m, self.complete_log))
        self.dfs["completed"] = df
        self._show_table(df.head(100), self.complete_table)

    # ══════════════════════════════════════════════
    # Tab 3: 行为分析
    # ══════════════════════════════════════════════

    def _tab_behavior(self):
        w = QWidget()
        l = QHBoxLayout(w)
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(12)

        g = QGroupBox("文件选择")
        gl = QVBoxLayout(g)
        b = self._file_row(gl, "交易明细:", "behavior_file_entry")
        b.clicked.connect(lambda: self._pick_file(self.behavior_file_entry))
        self._buttons.append(b)
        cl.addWidget(g)

        g2 = QGroupBox("参数")
        g2l = QVBoxLayout(g2)
        g2l.addWidget(QLabel("高频阈值(次)"))
        self.spin_freq = QSpinBox()
        self.spin_freq.setRange(1, 99999)
        self.spin_freq.setValue(10)
        g2l.addWidget(self.spin_freq)
        g2l.addWidget(QLabel("大额阈值(元)"))
        self.spin_amt = QSpinBox()
        self.spin_amt.setRange(100, 99999999)
        self.spin_amt.setValue(50000)
        g2l.addWidget(self.spin_amt)
        cl.addWidget(g2)

        btn = QPushButton("开始分析")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(lambda: self._run_async(self._do_behavior))
        self._buttons.append(btn)
        cl.addWidget(btn)
        cl.addStretch()

        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        self.behavior_table = QTableWidget()
        self.behavior_log = QTextEdit()
        self.behavior_log.setReadOnly(True)
        self._log_widgets.append(self.behavior_log)
        ol.addWidget(self.behavior_table, 1)
        ol.addWidget(self.behavior_log)

        l.addWidget(ctrl, 1)
        l.addWidget(out, 4)
        return w

    def _do_behavior(self):
        fp = self.behavior_file_entry.text()
        if not fp:
            self._log("请选择文件", self.behavior_log)
            return
        self.behavior_log.clear()
        df = pd.read_excel(fp)
        import pandas as pd
        results = analyze_account_behavior(
            df,
            freq_threshold=self.spin_freq.value(),
            amount_threshold=self.spin_amt.value(),
            log_callback=lambda m: self._log(m, self.behavior_log),
        )
        hf = results.get("高频账户")
        self._show_table(hf, self.behavior_table)
        if hf is not None and not hf.empty:
            out = os.path.join(os.path.dirname(fp), f"账户行为分析_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for k, v in results.items():
                    if isinstance(v, pd.DataFrame) and not v.empty:
                        v.to_excel(writer, sheet_name=k, index=False)
            self._log(f"✅ 已保存: {out}", self.behavior_log)

    # ══════════════════════════════════════════════
    # Tab 4: 关联分析
    # ══════════════════════════════════════════════

    def _tab_relation(self):
        w = QWidget()
        l = QHBoxLayout(w)
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(12)

        g = QGroupBox("文件选择")
        gl = QVBoxLayout(g)
        b = self._file_row(gl, "交易明细:", "relation_file_entry")
        b.clicked.connect(lambda: self._pick_file(self.relation_file_entry))
        self._buttons.append(b)
        cl.addWidget(g)

        g2 = QGroupBox("参数")
        g2l = QVBoxLayout(g2)
        g2l.addWidget(QLabel("最小交易次数"))
        self.spin_rel_cnt = QSpinBox()
        self.spin_rel_cnt.setRange(1, 99999)
        self.spin_rel_cnt.setValue(3)
        g2l.addWidget(self.spin_rel_cnt)
        g2l.addWidget(QLabel("最小总金额(元)"))
        self.spin_rel_amt = QSpinBox()
        self.spin_rel_amt.setRange(100, 99999999)
        self.spin_rel_amt.setValue(10000)
        g2l.addWidget(self.spin_rel_amt)
        cl.addWidget(g2)

        btn = QPushButton("开始分析")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(lambda: self._run_async(self._do_relation))
        self._buttons.append(btn)
        cl.addWidget(btn)
        cl.addStretch()

        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        self.relation_table = QTableWidget()
        self.relation_log = QTextEdit()
        self.relation_log.setReadOnly(True)
        self._log_widgets.append(self.relation_log)
        ol.addWidget(self.relation_table, 1)
        ol.addWidget(self.relation_log)

        l.addWidget(ctrl, 1)
        l.addWidget(out, 4)
        return w

    def _do_relation(self):
        fp = self.relation_file_entry.text()
        if not fp:
            self._log("请选择文件", self.relation_log)
            return
        self.relation_log.clear()
        df = pd.read_excel(fp)
        result = analyze_relations(
            df,
            min_count=self.spin_rel_cnt.value(),
            min_amount=self.spin_rel_amt.value(),
            log_callback=lambda m: self._log(m, self.relation_log),
        )
        self._show_table(result, self.relation_table)
        if not result.empty:
            out = os.path.join(os.path.dirname(fp), f"关联分析_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            result.to_excel(out, index=False)
            self._log(f"✅ 已保存: {out}", self.relation_log)

    # ══════════════════════════════════════════════
    # Tab 5: 高级筛选
    # ══════════════════════════════════════════════

    def _tab_filter(self):
        w = QWidget()
        l = QHBoxLayout(w)
        ctrl = QFrame()
        ctrl.setObjectName("sidePanel")
        cl = QVBoxLayout(ctrl)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(12)

        g = QGroupBox("文件选择")
        gl = QVBoxLayout(g)
        b = self._file_row(gl, "交易明细:", "filter_file_entry")
        b.clicked.connect(lambda: self._pick_file(self.filter_file_entry))
        self._buttons.append(b)
        cl.addWidget(g)

        g2 = QGroupBox("筛选条件")
        g2l = QVBoxLayout(g2)
        g2l.addWidget(QLabel("金额范围"))
        ar = QHBoxLayout()
        self.f_amin = QLineEdit()
        self.f_amin.setPlaceholderText("最小")
        self.f_amax = QLineEdit()
        self.f_amax.setPlaceholderText("最大")
        ar.addWidget(self.f_amin)
        ar.addWidget(QLabel("~"))
        ar.addWidget(self.f_amax)
        g2l.addLayout(ar)
        g2l.addWidget(QLabel("关键词(户名/摘要)"))
        self.f_kw = QLineEdit()
        g2l.addWidget(self.f_kw)
        cl.addWidget(g2)

        btn = QPushButton("执行筛选")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(lambda: self._run_async(self._do_filter))
        self._buttons.append(btn)
        cl.addWidget(btn)
        cl.addStretch()

        out = QFrame()
        out.setObjectName("workPanel")
        ol = QVBoxLayout(out)
        self.filter_table = QTableWidget()
        self.filter_log = QTextEdit()
        self.filter_log.setReadOnly(True)
        self._log_widgets.append(self.filter_log)
        ol.addWidget(self.filter_table, 1)
        ol.addWidget(self.filter_log)
        ol.addWidget(QLabel("筛选结果可直接复制或导出到源文件目录"))

        l.addWidget(ctrl, 1)
        l.addWidget(out, 4)
        return w

    def _do_filter(self):
        fp = self.filter_file_entry.text()
        if not fp:
            self._log("请选择文件", self.filter_log)
            return
        self.filter_log.clear()
        df = pd.read_excel(fp)
        result, stats = advanced_filter(
            df,
            amount_min=self.f_amin.text(),
            amount_max=self.f_amax.text(),
            keyword=self.f_kw.text(),
            log_callback=lambda m: self._log(m, self.filter_log),
        )
        self.dfs["filtered"] = result
        self._show_table(result, self.filter_table)
        self._log(f"原始{stats['原始']}条 → 筛选{stats['筛选后']}条", self.filter_log)
        if not result.empty:
            out = os.path.join(os.path.dirname(fp), f"筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            result.to_excel(out, index=False)
            self._log(f"✅ 已保存: {out}", self.filter_log)

    # ══════════════════════════════════════════════
    # 文件选择辅助
    # ══════════════════════════════════════════════

    def _pick_folder(self, entry):
        path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if path:
            entry.setText(path)

    def _pick_file(self, entry):
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel (*.xlsx *.xls *.csv)")
        if path:
            entry.setText(path)


# 需要全局引入
import pandas as pd
from datetime import datetime
