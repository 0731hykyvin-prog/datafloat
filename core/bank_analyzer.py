"""
银行交易分析引擎
=================
快进快出、敏感金额、深夜交易、高频账户、ATM取款、配送交易、风险评分。
"""

import os
from datetime import time

import pandas as pd


# ══════════════════════════════════════════════════════
# 数据加载
# ══════════════════════════════════════════════════════

def load_bank_data(file_path):
    """加载银行流水 Excel，统一 dtype=str，转换金额/时间为数值。"""
    df = pd.read_excel(file_path, dtype=str)
    df["交易金额_num"] = pd.to_numeric(df["交易金额"], errors="coerce").fillna(0)
    df["交易时间_dt"] = pd.to_datetime(df["交易时间"], errors="coerce")
    return df


# ══════════════════════════════════════════════════════
# 筛选器
# ══════════════════════════════════════════════════════

def filter_quick_in_out(df, min_hits=50):
    """
    快进快出：同一 交易时间+交易流水号 下有进有出，各保留1条。
    返回: 筛选后 DataFrame, 日志列表
    """
    logs = []
    valid = df.dropna(subset=["交易流水号", "交易时间", "收付标志"])

    g = valid.groupby(["交易时间", "交易流水号"])
    mask = g["收付标志"].transform(lambda x: ("进" in x.values) and ("出" in x.values))
    candidates = valid[mask]
    unique_groups = candidates.groupby(["交易时间", "交易流水号"]).ngroups

    result_rows = []
    for _, grp in candidates.groupby(["交易时间", "交易流水号"]):
        jin = grp[grp["收付标志"] == "进"]
        chu = grp[grp["收付标志"] == "出"]
        if len(jin) > 0:
            result_rows.append(jin.iloc[0])
        if len(chu) > 0:
            result_rows.append(chu.iloc[0])

    result = pd.DataFrame(result_rows)
    jin_n = (result["收付标志"] == "进").sum()
    chu_n = (result["收付标志"] == "出").sum()

    logs.append(f"快进快出: {len(result):,} 条 (进{jin_n:,}/出{chu_n:,})  |  {unique_groups:,} 组")

    # 按账户统计命中次数，只保留 > min_hits 的
    counts = result["交易账号"].value_counts()
    keep = counts[counts > min_hits].index
    result = result[result["交易账号"].isin(keep)]
    logs.append(f"高频过滤(>{min_hits}条): 保留 {len(result):,} 条  |  {len(keep)} 个账户")

    return result, logs


def filter_sensitive_amount(df, base=500, step=100):
    """
    敏感金额：金额 >= base 且为 step 的整数倍。
    """
    logs = []
    mask = (df["交易金额_num"] >= base) & (df["交易金额_num"] % step == 0)
    result = df[mask].copy()
    logs.append(f"敏感金额(≥{base},%{step}=0): {len(result):,} 条")
    return result, logs


def filter_night_trades(df, start_h=21, end_h=5):
    """
    深夜交易：交易时间在 [start_h, 24) 或 [0, end_h]。
    """
    logs = []
    valid = df.dropna(subset=["交易时间_dt"])
    hours = valid["交易时间_dt"].dt.hour
    mask = (hours >= start_h) | (hours < end_h)
    result = valid[mask].copy()
    logs.append(f"深夜交易({start_h}:00-{end_h}:00): {len(result):,} 条")
    return result, logs


def filter_atm_withdrawals(df):
    """ATM取款：摘要说明含'取款'。"""
    logs = []
    mask = df["摘要说明"].fillna("").str.contains("取款", na=False)
    result = df[mask].copy()
    logs.append(f"ATM取款: {len(result):,} 条")
    return result, logs


def filter_delivery_trades(df, keywords=None):
    """配送交易：对手户名含指定关键词。"""
    if keywords is None:
        keywords = ["uu跑腿", "京东秒送", "翠鸟配送", "顺丰速运", "达达秒送"]
    logs = []
    pattern = "|".join(keywords)
    mask = df["对手户名"].fillna("").str.contains(pattern, case=False, na=False)
    result = df[mask].copy()
    logs.append(f"配送交易: {len(result):,} 条")
    return result, logs


# ══════════════════════════════════════════════════════
# 风险评分
# ══════════════════════════════════════════════════════

def calculate_risk_score(df):
    """
    按账户统计各维度命中次数，加权计算风险分。
    权重: 深夜交易=1, 敏感金额=1, 深夜+敏感=2, ATM=2, 配送=2
    """
    mask_night = False
    mask_sensitive = False
    mask_atm = False
    mask_delivery = False

    if "交易时间_dt" in df.columns:
        hours = df["交易时间_dt"].dt.hour
        mask_night = (hours >= 21) | (hours < 5)
    if "交易金额_num" in df.columns:
        mask_sensitive = (df["交易金额_num"] >= 500) & (df["交易金额_num"] % 100 == 0)
    if "摘要说明" in df.columns:
        mask_atm = df["摘要说明"].fillna("").str.contains("取款", na=False)
    if "对手户名" in df.columns:
        kw = ["uu跑腿", "京东秒送", "翠鸟配送", "顺丰速运", "达达秒送"]
        mask_delivery = df["对手户名"].fillna("").str.contains("|".join(kw), case=False, na=False)

    # 分组统计
    grp = df.groupby(["交易账号", "账户开户名称"], dropna=False)
    stats = pd.DataFrame({
        "深夜交易次数": mask_night.groupby([df["交易账号"], df["账户开户名称"]]).sum().astype(int),
        "敏感金额次数": mask_sensitive.groupby([df["交易账号"], df["账户开户名称"]]).sum().astype(int),
        "ATM取款次数": mask_atm.groupby([df["交易账号"], df["账户开户名称"]]).sum().astype(int),
        "配送交易次数": mask_delivery.groupby([df["交易账号"], df["账户开户名称"]]).sum().astype(int),
        "总交易次数": grp.size(),
    }).fillna(0).astype(int).reset_index()

    # 深夜+敏感 = 同时命中
    night_sensitive = (mask_night & mask_sensitive)
    ns = night_sensitive.groupby([df["交易账号"], df["账户开户名称"]]).sum().astype(int).reset_index(name="深夜敏感金额次数")

    stats = stats.merge(ns, on=["交易账号", "账户开户名称"], how="left")
    stats["深夜敏感金额次数"] = stats["深夜敏感金额次数"].fillna(0).astype(int)

    # 风险分 = 深夜*1 + 敏感*1 + 深夜敏感*2 + ATM*2 + 配送*2
    stats["风险评分"] = (
        stats["深夜交易次数"] * 1
        + stats["敏感金额次数"] * 1
        + stats["深夜敏感金额次数"] * 2
        + stats["ATM取款次数"] * 2
        + stats["配送交易次数"] * 2
    )
    stats = stats.sort_values("风险评分", ascending=False)

    # 按人员汇总
    person = stats.groupby("账户开户名称").agg({
        "深夜交易次数": "sum", "敏感金额次数": "sum",
        "深夜敏感金额次数": "sum", "ATM取款次数": "sum",
        "配送交易次数": "sum", "风险评分": "sum",
    }).reset_index().sort_values("风险评分", ascending=False)

    return stats, person


def export_person_summary(df, output_dir="银行交易分析结果"):
    """导出人员汇总表：账户开户名称 + 开户人证件号码 + 风险评分。"""
    os.makedirs(output_dir, exist_ok=True)
    if "开户人证件号码" in df.columns:
        summary = df[["账户开户名称", "开户人证件号码"]].drop_duplicates().sort_values("账户开户名称")
    else:
        summary = df[["账户开户名称"]].drop_duplicates().sort_values("账户开户名称")

    path = os.path.join(output_dir, "人员汇总表.xlsx")
    summary.to_excel(path, index=False)
    return path, summary
