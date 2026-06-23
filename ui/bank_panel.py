"""
银行交易分析面板 V3.1
6 个子模块：核心分析 / CSV合并 / 补全 / 行为 / 关联 / 筛选
"""

import os
import threading
from datetime import datetime

import pandas as pd

from ui.qt_compat import (
    Qt,
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
    advanced_filter,
    analyze_account_behavior,
    analyze_relations,
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

        self.init_ui()

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
        if self._log_widgets:
            self._log_widgets[-1].append(str(msg))

    # ── UI ──
    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        self.subtabs = QTabWidget()
        self.subtabs.addTab(self._tab_core(), "📊 核心分析")
        self.subtabs.addTab(self._tab_merge(), "📦 CSV合并")
        self.subtabs.addTab(self._tab_complete(), "🔧 补全交易")
        self.subtabs.addTab(self._tab_behavior(), "📈 行为分析")
        self.subtabs.addTab(self._tab_relation(), "🔗 关联分析")
        self.subtabs.addTab(self._tab_filter(), "🔍 高级筛选")
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

        # 数据导入
        g1 = QGroupBox("数据导入")
        g1l = QVBoxLayout(g1)
        self.lbl_file = QLabel("未选择文件")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setObjectName("pathLabel")
        g1l.addWidget(self.lbl_file)
        bf = self._make_btn("选择银行流水", self._sel_file, True)
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

        # 参数
        g2 = QGroupBox("参数")
        g2l = QVBoxLayout(g2)
        g2l.addWidget(QLabel("敏感金额 ≥"))
        sb = QSpinBox(); sb.setRange(100, 999999); sb.setValue(500)
        self.spin_base = sb; g2l.addWidget(sb)
        g2l.addWidget(QLabel("金额公差 %"))
        ss = QSpinBox(); ss.setRange(1, 99999); ss.setValue(100)
        self.spin_step = ss; g2l.addWidget(ss)
        g2l.addWidget(QLabel("高频阈值 >"))
        sf = QSpinBox(); sf.setRange(1, 99999); sf.setValue(50)
        self.spin_freq = sf; g2l.addWidget(sf)
        g2l.addWidget(QLabel("深夜时段"))
        hr = QHBoxLayout()
        sn = QSpinBox(); sn.setRange(0, 23); sn.setValue(21)
        se = QSpinBox(); se.setRange(0, 23); se.setValue(5)
        hr.addWidget(sn); hr.addWidget(QLabel("-")); hr.addWidget(se)
        self.spin_nstart = sn; self.spin_nend = se
        g2l.addLayout(hr)
        cl.addWidget(g2)

        # 按钮
        g3 = QGroupBox("分析")
        g3l = QVBoxLayout(g3)
        for txt, fn in [
            ("快进快出", self._run_core_quick),
            ("敏感金额", self._run_core_sensitive),
            ("深夜交易", self._run_core_night),
            ("风险评分", self._run_core_risk),
            ("导出人员名单", self._run_core_export),
        ]:
            g3l.addWidget(self._make_btn(txt, fn, txt == "快进快出"))
        cl.addWidget(g3)
        cl.addStretch()

        self._table_core = table
        return w

    def _sel_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择银行流水", "", "Excel (*.xlsx *.xls *.csv)")
        if p:
            self.lbl_file.setText(p)
            try:
                self.df = load_bank_data(p)
                self._log_last(f"加载: {len(self.df):,}条  进{(self.df['收付标志']=='进').sum():,}  出{(self.df['收付标志']=='出').sum():,}")
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

    def _run_core_quick(self):
        if not self._ensure_df(): return
        r, logs = filter_quick_in_out(self.df, min_hits=self.spin_freq.value())
        self.current_result = r
        for l in logs:
            self._log_last(l)
        self._show_table(r, self._table_core)
        self._save_excel(r, "B01_快进快出.xlsx")

    def _run_core_sensitive(self):
        if not self._ensure_df(): return
        src = self.current_result if self.current_result is not None else self.df
        r, logs = filter_sensitive_amount(src, base=self.spin_base.value(), step=self.spin_step.value())
        self.current_result = r
        for l in logs:
            self._log_last(l)
        self._show_table(r, self._table_core)
        self._save_excel(r, "B02_敏感金额.xlsx")

    def _run_core_night(self):
        if not self._ensure_df(): return
        src = self.current_result if self.current_result is not None else self.df
        r, logs = filter_night_trades(src, start_h=self.spin_nstart.value(), end_h=self.spin_nend.value())
        self.current_result = r
        for l in logs:
            self._log_last(l)
        self._show_table(r, self._table_core)
        self._save_excel(r, "B03_深夜交易.xlsx")

    def _run_core_risk(self):
        if not self._ensure_df(): return
        src = self.current_result if self.current_result is not None else self.df
        acc, person = calculate_risk_score(src)
        self._log_last(f"风险评分: {len(acc)}账户, {len(person)}人")
        self._show_table(person, self._table_core)
        self.current_result = person
        self._save_excel(acc, "B04_账户风险.xlsx")
        self._save_excel(person, "B05_人员风险.xlsx")

    def _run_core_export(self):
        if not self._ensure_df(): return
        src = self.current_result if self.current_result is not None else self.df
        p, s = export_person_summary(src, self.output_dir)
        self._log_last(f"人员名单: {p} ({len(s)}人)")
        self._show_table(s, self._table_core)

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

    # ══════════════════════════════════════════════
    # Tab 4: 行为分析
    # ══════════════════════════════════════════════

    def _tab_behavior(self):
        w, ctrl, cl, table, tlog = self._panel()
        g = QGroupBox("文件 & 参数")
        gl = QVBoxLayout(g)
        b = self._file_row(gl, "交易明细:", "behav_file")
        b.clicked.connect(lambda: self._pick_file(self.behav_file))
        gl.addWidget(QLabel("高频阈值(次)"))
        self.b_freq = QSpinBox(); self.b_freq.setRange(1, 99999); self.b_freq.setValue(10)
        gl.addWidget(self.b_freq)
        gl.addWidget(QLabel("大额阈值(元)"))
        self.b_amt = QSpinBox(); self.b_amt.setRange(100, 99999999); self.b_amt.setValue(50000)
        gl.addWidget(self.b_amt)
        cl.addWidget(g)
        cl.addWidget(self._make_btn("开始分析", self._do_behavior, True))
        cl.addStretch()
        self._table_behav = table
        return w

    def _do_behavior(self):
        f = self.behav_file.text()
        if not f: self._log_last("请选择文件"); return
        df = pd.read_excel(f)
        r = analyze_account_behavior(df, self.b_freq.value(), self.b_amt.value(),
                                      log_callback=lambda m: self._log_last(m))
        hf = r.get("高频账户")
        self._show_table(hf, self._table_behav)
        if hf is not None and not hf.empty:
            out = os.path.join(os.path.dirname(f), f"行为分析_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
            with pd.ExcelWriter(out) as wb:
                for k, v in r.items():
                    if isinstance(v, pd.DataFrame) and not v.empty:
                        v.to_excel(wb, sheet_name=k, index=False)
            self._log_last(f"📁 {out}")

    # ══════════════════════════════════════════════
    # Tab 5: 关联分析
    # ══════════════════════════════════════════════

    def _tab_relation(self):
        w, ctrl, cl, table, tlog = self._panel()
        g = QGroupBox("文件 & 参数")
        gl = QVBoxLayout(g)
        b = self._file_row(gl, "交易明细:", "rel_file")
        b.clicked.connect(lambda: self._pick_file(self.rel_file))
        gl.addWidget(QLabel("最小交易次数"))
        self.r_cnt = QSpinBox(); self.r_cnt.setRange(1, 99999); self.r_cnt.setValue(3)
        gl.addWidget(self.r_cnt)
        gl.addWidget(QLabel("最小总金额(元)"))
        self.r_amt = QSpinBox(); self.r_amt.setRange(100, 99999999); self.r_amt.setValue(10000)
        gl.addWidget(self.r_amt)
        cl.addWidget(g)
        cl.addWidget(self._make_btn("开始分析", self._do_relation, True))
        cl.addStretch()
        self._table_rel = table
        return w

    def _do_relation(self):
        f = self.rel_file.text()
        if not f: self._log_last("请选择文件"); return
        df = pd.read_excel(f)
        r = analyze_relations(df, self.r_cnt.value(), self.r_amt.value(),
                               log_callback=lambda m: self._log_last(m))
        self._show_table(r, self._table_rel)
        if not r.empty:
            self._save_excel(r, "关联分析.xlsx")

    # ══════════════════════════════════════════════
    # Tab 6: 高级筛选
    # ══════════════════════════════════════════════

    def _tab_filter(self):
        w, ctrl, cl, table, tlog = self._panel()
        g = QGroupBox("文件 & 条件")
        gl = QVBoxLayout(g)
        b = self._file_row(gl, "交易明细:", "filt_file")
        b.clicked.connect(lambda: self._pick_file(self.filt_file))
        gl.addWidget(QLabel("金额范围"))
        ar = QHBoxLayout()
        self.f_amin = QLineEdit(); self.f_amin.setPlaceholderText("最小")
        self.f_amax = QLineEdit(); self.f_amax.setPlaceholderText("最大")
        ar.addWidget(self.f_amin); ar.addWidget(QLabel("~")); ar.addWidget(self.f_amax)
        gl.addLayout(ar)
        gl.addWidget(QLabel("关键词"))
        self.f_kw = QLineEdit()
        gl.addWidget(self.f_kw)
        cl.addWidget(g)
        cl.addWidget(self._make_btn("执行筛选", self._do_filter, True))
        cl.addStretch()
        self._table_filt = table
        return w

    def _do_filter(self):
        f = self.filt_file.text()
        if not f: self._log_last("请选择文件"); return
        df = pd.read_excel(f)
        r, s = advanced_filter(df,
            amount_min=self.f_amin.text(), amount_max=self.f_amax.text(),
            keyword=self.f_kw.text(),
            log_callback=lambda m: self._log_last(m))
        self._log_last(f"原始{s['原始']}条 → {s['筛选后']}条")
        self._show_table(r, self._table_filt)
        self._save_excel(r, "筛选结果.xlsx")

    # ── 文件选择工具 ──
    def _pick_dir(self, entry):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if d: entry.setText(d)

    def _pick_file(self, entry):
        p, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel (*.xlsx *.xls *.csv)")
        if p: entry.setText(p)
