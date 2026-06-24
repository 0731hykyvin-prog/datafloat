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

def filter_quick_in_out(df):
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
