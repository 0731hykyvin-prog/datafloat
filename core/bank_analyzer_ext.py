"""
银行交易分析扩展模块
=====================
CSV合并 / 补全 / 行为分析 / 关联分析 / 高级筛选
（从 bank_analyzer_windows.py 提炼，去除 tkinter）
"""

import os
import glob
from datetime import datetime

from collections import defaultdict
import pandas as pd

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
    - 文件名含"银行交易明细" → 合并为交易明细
    - 文件名含"银行账户信息" → 合并为账户信息
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
        elif "银行交易明细" in name:
            trans_files.append(f)
        elif "银行账户信息" in name:
            acc_files.append(f)
        else:
            skipped.append(name)

    log(f"银行交易明细: {len(trans_files)} 个, 银行账户信息: {len(acc_files)} 个, 跳过: {len(skipped)} 个")

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
        ccb_count = 0
        for f in acc_files:
            try:
                df = read_csv_auto_encoding(f)
                df.columns = [clean_text(c) for c in df.columns]

                # 建设银行特殊清洗：卡号/账号只保留前19位数字
                if "建设银行" in os.path.basename(f):
                    for col in ["交易卡号", "交易账号"]:
                        if col in df.columns:
                            df[col] = df[col].astype(str).str.extract(r'(\d{19})')[0]
                            ccb_count += 1
                    log(f"  🔧 {os.path.basename(f)}: 建设银行清洗（卡号/账号→19位）")

                df["来源文件"] = os.path.basename(f)
                all_data.append(df)
                log(f"  ✓ {os.path.basename(f)}: {len(df)}行")
            except Exception as e:
                log(f"  ✗ {os.path.basename(f)}: {e}")
        if ccb_count:
            log(f"  共清洗 {ccb_count} 个建设银行字段")

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

def complete_transactions(trans_path, acc_path, log_callback=None):
    """
    用账户信息补全交易明细的缺失字段（卡号/户名/证件号）。

    匹配逻辑：
    1. 优先用"交易账号"匹配
    2. 匹配不上则用"交易卡号"匹配
    3. 补全：交易卡号、账户开户名称、开户人证件号码
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    df_trans = pd.read_excel(trans_path)
    df_acc = pd.read_excel(acc_path)

    # 清洗：去除所有字符串字段的制表符和首尾空白
    for df_ in (df_trans, df_acc):
        for col in df_.columns:
            df_[col] = df_[col].astype(str).str.replace("\t", "").str.strip()
            df_[col] = df_[col].replace(["nan", "None", ""], pd.NA)

    log(f"交易明细: {len(df_trans)}行, 账户信息: {len(df_acc)}行")

    # 确保补全目标列存在（字符串类型避免 dtype 冲突）
    target_cols = ["交易卡号", "交易方户名", "交易方证件号码"]
    for col in target_cols:
        if col not in df_trans.columns:
            df_trans[col] = ""
        df_trans[col] = df_trans[col].astype(object)

    def _clean_num(val):
        """清洗数字字段：去科学计数法、去小数点、去空格。"""
        if pd.isna(val):
            return ""
        s = str(val).strip()
        try:
            s = str(int(float(s)))  # 处理科学计数法如 6.22e+18
        except ValueError:
            pass
        return s

    # 构建两个索引
    acc_by_account = {}
    acc_by_card = {}
    for _, row in df_acc.iterrows():
        account = _clean_num(row.get("交易账号"))
        card = _clean_num(row.get("交易卡号"))
        info = {
            "交易卡号": _clean_num(row.get("交易卡号")),
            "交易方户名": str(row.get("账户开户名称", "")).strip() if pd.notna(row.get("账户开户名称")) else "",
            "交易方证件号码": str(row.get("开户人证件号码", "")).strip() if pd.notna(row.get("开户人证件号码")) else "",
        }
        if account:
            acc_by_account[account] = info
        if card:
            acc_by_card[card] = info

    log(f"索引构建: 按账号{len(acc_by_account)}条, 按卡号{len(acc_by_card)}条")

    filled_1 = 0
    filled_2 = 0
    # 账户信息表字段 → 交易明细表字段 的映射
    fill_map = {"交易卡号": "交易卡号", "交易方户名": "交易方户名", "交易方证件号码": "交易方证件号码"}

    # ═══ 第一遍：交易账号 对 交易账号 ═══
    for idx, row in df_trans.iterrows():
        trans_account = _clean_num(row.get("交易账号"))
        if not trans_account:
            continue
        match = acc_by_account.get(trans_account)
        if match:
            for dst_col in fill_map:
                current = str(df_trans.at[idx, dst_col]).strip() if pd.notna(df_trans.at[idx, dst_col]) else ""
                if not current or current in ("nan", "None", ""):
                    df_trans.at[idx, dst_col] = match.get(dst_col, "")
            filled_1 += 1

    # ═══ 第二遍：交易卡号 对 交易卡号 ═══
    for idx, row in df_trans.iterrows():
        trans_card = _clean_num(row.get("交易卡号"))
        if not trans_card:
            continue
        match = acc_by_card.get(trans_card)
        if match:
            for dst_col in fill_map:
                current = str(df_trans.at[idx, dst_col]).strip() if pd.notna(df_trans.at[idx, dst_col]) else ""
                if not current or current in ("nan", "None", ""):
                    df_trans.at[idx, dst_col] = match.get(dst_col, "")
            filled_2 += 1

    log(f"第一遍(交易账号匹配): {filled_1}条")
    log(f"第二遍(交易卡号匹配): {filled_2}条")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.dirname(trans_path)
    out_path = os.path.join(out_dir, f"交易明细_补全结果_{timestamp}.xlsx")
    df_trans.to_excel(out_path, index=False)
    log(f"✅ 补全完成: {os.path.basename(out_path)}")

    return df_trans, out_path

# ══════════════════════════════════════════════════════
# 3. 账户行为分析
# ══════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════
# 4. 关联分析
# ══════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════
# 5. 高级筛选
# ══════════════════════════════════════════════════════

