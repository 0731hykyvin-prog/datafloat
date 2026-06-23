"""
银行交易分析扩展模块
=====================
CSV合并 / 补全 / 行为分析 / 关联分析 / 高级筛选
（从 bank_analyzer_windows.py 提炼，去除 tkinter）
"""

import os
import glob
import re
from collections import defaultdict
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import load_workbook


# ══════════════════════════════════════════════════════
# 工具函数
# ══════════════════════════════════════════════════════

def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text)
    for char in ['\t', '\n', '\r', '　']:
        text = text.replace(char, '')
    return text.strip()


def read_csv_auto_encoding(file_path):
    encodings = ['utf-8', 'gb18030', 'gb2312', 'gbk', 'latin1']
    for enc in encodings:
        try:
            return pd.read_csv(file_path, dtype=str, encoding=enc, on_bad_lines='skip')
        except Exception:
            continue
    raise Exception(f"无法读取: {file_path}")


# ══════════════════════════════════════════════════════
# 1. CSV 合并
# ══════════════════════════════════════════════════════

def merge_csv_folder(folder_path, log_callback=None):
    """
    扫描文件夹下所有 CSV，按文件名分类合并。
    - 文件名含"交易明细" → 合并为交易明细
    - 文件名含"账户信息" → 合并为账户信息
    - 文件名含"合并结果" → 跳过
    返回: {"交易明细": output_path, "账户信息": output_path}, logs
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    logs = []
    csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
    log(f"扫描到 {len(csv_files)} 个 CSV 文件")

    trans_files = []
    acc_files = []
    skipped = []

    for f in csv_files:
        name = os.path.basename(f)
        if "合并结果" in name:
            skipped.append(name)
        elif "交易明细" in name:
            trans_files.append(f)
        elif "账户信息" in name:
            acc_files.append(f)
        else:
            skipped.append(name)

    log(f"交易明细: {len(trans_files)} 个, 账户信息: {len(acc_files)} 个, 跳过: {len(skipped)} 个")

    results = {}
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 合并交易明细
    if trans_files:
        all_data = []
        for f in trans_files:
            try:
                df = read_csv_auto_encoding(f)
                df.columns = [clean_text(c) for c in df.columns]
                df["来源文件"] = os.path.basename(f)
                all_data.append(df)
                log(f"  ✓ {os.path.basename(f)}: {len(df)}行")
            except Exception as e:
                log(f"  ✗ {os.path.basename(f)}: {e}")
        if all_data:
            merged = pd.concat(all_data, ignore_index=True)
            path = os.path.join(folder_path, f"银行交易明细合并结果_{timestamp}.xlsx")
            merged.to_excel(path, index=False)
            results["交易明细"] = path
            log(f"✅ 交易明细: {os.path.basename(path)} ({len(merged)}行)")

    # 合并账户信息
    if acc_files:
        all_data = []
        for f in acc_files:
            try:
                df = read_csv_auto_encoding(f)
                df.columns = [clean_text(c) for c in df.columns]
                df["来源文件"] = os.path.basename(f)
                all_data.append(df)
                log(f"  ✓ {os.path.basename(f)}: {len(df)}行")
            except Exception as e:
                log(f"  ✗ {os.path.basename(f)}: {e}")
        if all_data:
            merged = pd.concat(all_data, ignore_index=True)
            path = os.path.join(folder_path, f"银行账户信息合并结果_{timestamp}.xlsx")
            merged.to_excel(path, index=False)
            results["账户信息"] = path
            log(f"✅ 账户信息: {os.path.basename(path)} ({len(merged)}行)")

    return results


# ══════════════════════════════════════════════════════
# 2. 补全交易明细
# ══════════════════════════════════════════════════════

def complete_transactions(trans_path, acc_path, deduplicate=True, log_callback=None):
    """
    用账户信息补全交易明细的缺失字段（卡号/户名/证件号）。
    返回: (补全后 DataFrame, 输出路径, 日志列表)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    df_trans = pd.read_excel(trans_path)
    df_acc = pd.read_excel(acc_path)

    log(f"交易明细: {len(df_trans)}行, 账户信息: {len(df_acc)}行")

    # 去重
    if deduplicate and "交易流水号" in df_trans.columns:
        before = len(df_trans)
        df_trans = df_trans.drop_duplicates(subset=["交易流水号"], keep='first')
        log(f"去重: {before} → {len(df_trans)} 行")

    # 填充交易卡号
    if "交易卡号" in df_trans.columns and "交易账号" in df_trans.columns:
        if "交易账号" in df_acc.columns and "交易卡号" in df_acc.columns:
            empty = df_trans["交易卡号"].isna()
            card_map = df_acc.dropna(subset=["交易卡号"]).set_index("交易账号")["交易卡号"].to_dict()
            df_trans.loc[empty, "交易卡号"] = df_trans.loc[empty, "交易账号"].map(card_map)
            filled = df_trans.loc[empty, "交易卡号"].notna().sum()
            log(f"填充交易卡号: {filled}/{empty.sum()}条")

    # 填充账户信息
    for col in ["账户开户名称", "开户人证件号码"]:
        if col not in df_trans.columns:
            df_trans[col] = None

    if "交易账号" in df_acc.columns:
        for _, row in df_acc.iterrows():
            acc = str(row.get("交易账号", ""))
            if acc and acc != "nan":
                mask = (df_trans["交易账号"].astype(str) == acc)
                for col in ["账户开户名称", "开户人证件号码"]:
                    if col in df_acc.columns:
                        df_trans.loc[mask & df_trans[col].isna(), col] = str(row.get(col, ""))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.dirname(trans_path)
    out_path = os.path.join(out_dir, f"交易明细_补全结果_{timestamp}.xlsx")
    df_trans.to_excel(out_path, index=False)
    log(f"✅ 补全完成: {os.path.basename(out_path)}")

    return df_trans, out_path


# ══════════════════════════════════════════════════════
# 3. 账户行为分析
# ══════════════════════════════════════════════════════

def analyze_account_behavior(df, freq_threshold=10, amount_threshold=50000,
                              time_window_days=30, log_callback=None):
    """
    高频/大额账户识别。
    返回: {"高频账户": DataFrame, "大额账户": DataFrame, "输出路径": str}
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    if "交易时间" in df.columns:
        df["_dt"] = pd.to_datetime(df["交易时间"], errors="coerce")
    if "交易金额" in df.columns:
        df["_amt"] = pd.to_numeric(df["交易金额"], errors="coerce")

    account_col = "交易账号"
    if account_col not in df.columns:
        account_col = "交易卡号"

    results = {}
    if account_col in df.columns:
        groups = df.groupby(account_col)
        log(f"活跃账户: {len(groups)} 个")

        high_freq = []
        high_amt = []
        for acc, grp in groups:
            cnt = len(grp)
            total = grp["_amt"].sum() if "_amt" in grp.columns else 0
            if cnt >= freq_threshold:
                high_freq.append({"账户": acc, "交易次数": cnt, "总金额": round(total, 2)})
            if total >= amount_threshold:
                high_amt.append({"账户": acc, "交易次数": cnt, "总金额": round(total, 2)})

        results["高频账户"] = pd.DataFrame(high_freq).sort_values("交易次数", ascending=False)
        results["大额账户"] = pd.DataFrame(high_amt).sort_values("总金额", ascending=False)
        log(f"高频账户(≥{freq_threshold}次): {len(high_freq)} 个")
        log(f"大额账户(≥{amount_threshold:,}元): {len(high_amt)} 个")

    return results


# ══════════════════════════════════════════════════════
# 4. 关联分析
# ══════════════════════════════════════════════════════

def analyze_relations(df, min_count=3, min_amount=10000, log_callback=None):
    """
    发现账户间交易关联关系。
    返回: DataFrame(来源账户, 目标账户, 交易次数, 总金额)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    relations = defaultdict(lambda: {"count": 0, "total": 0})

    for _, row in df.iterrows():
        from_acc = str(row.get("交易账号", "")).strip()
        to_acc = str(row.get("交易方账号", "")).strip()
        amt = pd.to_numeric(row.get("交易金额", 0), errors="coerce") or 0

        if from_acc and to_acc and from_acc != "nan" and to_acc != "nan" and from_acc != to_acc:
            key = (from_acc, to_acc)
            relations[key]["count"] += 1
            relations[key]["total"] += amt

    significant = []
    for (a, b), d in relations.items():
        if d["count"] >= min_count and d["total"] >= min_amount:
            significant.append({
                "来源账户": a, "目标账户": b,
                "交易次数": d["count"], "总金额": round(d["total"], 2),
            })

    result = pd.DataFrame(significant).sort_values("交易次数", ascending=False)
    log(f"发现 {len(result)} 个显著关联 (≥{min_count}次, ≥{min_amount:,}元)")
    return result


# ══════════════════════════════════════════════════════
# 5. 高级筛选
# ══════════════════════════════════════════════════════

def advanced_filter(df, amount_min=None, amount_max=None,
                     date_start=None, date_end=None,
                     direction=None, keyword=None, log_callback=None):
    """
    多条件组合筛选。
    返回: 筛选后 DataFrame, 统计信息 dict
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    original = len(df)
    df = df.copy()

    if "交易时间" in df.columns:
        df["_dt"] = pd.to_datetime(df["交易时间"], errors="coerce")
    if "交易金额" in df.columns:
        df["_amt"] = pd.to_numeric(df["交易金额"], errors="coerce")

    if amount_min is not None and amount_min != "":
        df = df[df["_amt"] >= float(amount_min)]
        log(f"金额≥{amount_min}: {len(df)}条")
    if amount_max is not None and amount_max != "":
        df = df[df["_amt"] <= float(amount_max)]
        log(f"金额≤{amount_max}: {len(df)}条")
    if date_start and str(date_start).strip():
        df = df[df["_dt"] >= pd.to_datetime(date_start)]
        log(f"时间≥{date_start}: {len(df)}条")
    if date_end and str(date_end).strip():
        df = df[df["_dt"] <= pd.to_datetime(date_end)]
        log(f"时间≤{date_end}: {len(df)}条")
    if keyword and keyword.strip():
        kw = keyword.strip()
        mask = pd.Series(False, index=df.index)
        for col in ["交易方户名", "对手户名", "摘要说明", "账户开户名称"]:
            if col in df.columns:
                mask |= df[col].astype(str).str.contains(kw, na=False, case=False)
        df = df[mask]
        log(f"关键词'{kw}': {len(df)}条")

    stats = {
        "原始": original, "筛选后": len(df),
        "总金额": round(df["_amt"].sum(), 2) if "_amt" in df.columns else 0,
    }
    return df, stats
