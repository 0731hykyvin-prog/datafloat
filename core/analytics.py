"""
数据分析引擎
=============
以本方号码为基准，分析高频联系人、深夜通话、风险关系。
单号码 / 多号码 统一按本方号码分组输出。
"""

import pandas as pd
import warnings


def _empty_df(columns):
    return pd.DataFrame(columns=columns)


def _to_datetime(series):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        return pd.to_datetime(series, errors="coerce")


def _to_duration_seconds(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


# ══════════════════════════════════════════════════════
# 数据概览
# ══════════════════════════════════════════════════════

def get_data_summary(df):
    """整体数据概览（不分组）。"""
    if df is None or df.empty:
        return {
            "总记录数": 0, "来源文件数": 0, "本方号码数": 0,
            "对方号码数": 0, "时间范围": "暂无", "关键字段完整率": "0%",
        }
    key_fields = [c for c in ["本方号码", "对方号码", "开始时间"] if c in df.columns]
    completeness = 0
    if key_fields:
        completeness = round(df[key_fields].notna().mean().mean() * 100, 2)
    time_range = "暂无"
    if "开始时间" in df.columns:
        times = _to_datetime(df["开始时间"]).dropna()
        if not times.empty:
            time_range = f"{times.min():%Y-%m-%d %H:%M} 至 {times.max():%Y-%m-%d %H:%M}"
    return {
        "总记录数": len(df),
        "来源文件数": df["来源文件"].nunique() if "来源文件" in df.columns else 0,
        "本方号码数": df["本方号码"].nunique() if "本方号码" in df.columns else 0,
        "对方号码数": df["对方号码"].nunique() if "对方号码" in df.columns else 0,
        "时间范围": time_range,
        "关键字段完整率": f"{completeness}%",
    }


# ══════════════════════════════════════════════════════
# 高频联系人 — 按本方号码分组
# ══════════════════════════════════════════════════════

def get_top_contacts_by_user(df, top_n=10):
    """
    按本方号码分组，返回 dict: {本方号码: DataFrame(对方号码, 次数)}
    """
    if df is None or df.empty:
        return {}
    if "本方号码" not in df.columns or "对方号码" not in df.columns:
        return {}
    clean = df.dropna(subset=["本方号码", "对方号码"])
    result = {}
    for user, group in clean.groupby("本方号码"):
        top = (
            group.groupby("对方号码")
            .size()
            .reset_index(name="次数")
            .sort_values("次数", ascending=False)
            .head(top_n)
        )
        result[user] = top
    return result


def get_global_top_contacts(df, top_n=20):
    """全局高频对方号码（跨所有本方号码）。"""
    columns = ["对方号码", "次数", "涉及本方号码数"]
    if df is None or df.empty or "对方号码" not in df.columns:
        return _empty_df(columns)
    clean = df.dropna(subset=["对方号码"])
    if clean.empty:
        return _empty_df(columns)
    result = clean.groupby("对方号码").agg(次数=("对方号码", "size"))
    if "本方号码" in clean.columns:
        result["涉及本方号码数"] = clean.groupby("对方号码")["本方号码"].nunique()
    else:
        result["涉及本方号码数"] = 0
    return result.reset_index().sort_values("次数", ascending=False).head(top_n)


# ══════════════════════════════════════════════════════
# 深夜通话 — 按本方号码分组
# ══════════════════════════════════════════════════════

def get_night_calls_by_user(df, start_hour=0, end_hour=5):
    """
    按本方号码分组，返回 dict: {本方号码: DataFrame(深夜通话明细)}
    """
    if df is None or df.empty or "开始时间" not in df.columns:
        return {}
    if "本方号码" not in df.columns:
        return {}
    work = df.dropna(subset=["开始时间", "本方号码"]).copy()
    work["_dt"] = _to_datetime(work["开始时间"])
    work = work.dropna(subset=["_dt"])
    hours = work["_dt"].dt.hour
    mask = (hours >= start_hour) & (hours <= end_hour)
    night = work[mask].copy()
    night.drop(columns=["_dt"], inplace=True)

    result = {}
    for user, group in night.groupby("本方号码"):
        cols = [c for c in ["对方号码", "开始时间", "通话时长", "呼叫类型"] if c in group.columns]
        result[user] = group[cols].sort_values("开始时间")
    return result


def get_night_calls_summary(df, start_hour=0, end_hour=5):
    """
    深夜通话汇总（不分明细）：按本方号码，统计深夜通话次数、涉及对方号码数。
    返回 DataFrame: 本方号码, 深夜通话次数, 涉及对方号码数
    """
    if df is None or df.empty or "开始时间" not in df.columns:
        return _empty_df(["本方号码", "深夜通话次数", "涉及对方号码数"])
    work = df.dropna(subset=["开始时间", "本方号码"]).copy()
    work["_dt"] = _to_datetime(work["开始时间"])
    work = work.dropna(subset=["_dt"])
    hours = work["_dt"].dt.hour
    mask = (hours >= start_hour) & (hours <= end_hour)
    night = work[mask]
    if night.empty:
        return _empty_df(["本方号码", "深夜通话次数", "涉及对方号码数"])
    summary = night.groupby("本方号码").agg(
        深夜通话次数=("对方号码", "size"),
        涉及对方号码数=("对方号码", "nunique"),
    ).reset_index().sort_values("深夜通话次数", ascending=False)
    return summary


# ══════════════════════════════════════════════════════
# 风险关系 — 按（本方号码, 对方号码）分组
# ══════════════════════════════════════════════════════

def get_risk_contacts(df, min_calls=5, night_weight=2):
    """
    按 (本方号码, 对方号码) 统计风险关系。
    返回 DataFrame: 本方号码, 对方号码, 总次数, 深夜次数, 短通话次数, 风险分
    """
    columns = ["本方号码", "对方号码", "总次数", "深夜次数", "短通话次数", "风险分"]
    if df is None or df.empty:
        return _empty_df(columns)
    required = {"本方号码", "对方号码"}
    if not required.issubset(df.columns):
        return _empty_df(columns)

    work = df.dropna(subset=["本方号码", "对方号码"]).copy()
    if work.empty:
        return _empty_df(columns)

    # 深夜标记
    if "开始时间" in work.columns:
        dt = _to_datetime(work["开始时间"])
        work["_is_night"] = dt.dt.hour.between(0, 5).fillna(False)
    else:
        work["_is_night"] = False

    # 短通话标记
    if "通话时长" in work.columns:
        dur = _to_duration_seconds(work["通话时长"])
        work["_is_short"] = dur.between(1, 5)
    else:
        work["_is_short"] = False

    result = (
        work.groupby(["本方号码", "对方号码"])
        .agg(
            总次数=("对方号码", "size"),
            深夜次数=("_is_night", "sum"),
            短通话次数=("_is_short", "sum"),
        )
        .reset_index()
    )
    result["风险分"] = (
        result["总次数"] + result["深夜次数"] * night_weight + result["短通话次数"]
    )
    return (
        result[result["总次数"] >= min_calls]
        .sort_values(["风险分", "总次数"], ascending=False)
        .head(50)
    )


def get_risk_by_user(df, min_calls=5, night_weight=2):
    """
    按本方号码分组返回风险关系。dict: {本方号码: DataFrame}
    """
    all_risk = get_risk_contacts(df, min_calls, night_weight)
    if all_risk.empty:
        return {}
    result = {}
    for user, group in all_risk.groupby("本方号码"):
        result[user] = group.sort_values("风险分", ascending=False)
    return result
