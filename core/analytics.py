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


def get_data_summary(df):
    if df is None or df.empty:
        return {
            "总记录数": 0,
            "来源文件数": 0,
            "本方号码数": 0,
            "对方号码数": 0,
            "时间范围": "暂无",
            "关键字段完整率": "0%",
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


def get_top_contacts_by_user(df, top_n=10):
    if df is None or df.empty:
        return None

    if "本方号码" not in df.columns or "对方号码" not in df.columns:
        return None

    clean_df = df.dropna(subset=["本方号码", "对方号码"])

    result = {}
    for user, group in clean_df.groupby("本方号码"):
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
    if df is None or df.empty or "对方号码" not in df.columns:
        return _empty_df(["对方号码", "次数", "涉及本方号码数"])

    group_cols = ["对方号码"]
    clean_df = df.dropna(subset=group_cols)
    if clean_df.empty:
        return _empty_df(["对方号码", "次数", "涉及本方号码数"])

    result = clean_df.groupby("对方号码").agg(次数=("对方号码", "size"))
    if "本方号码" in clean_df.columns:
        result["涉及本方号码数"] = clean_df.groupby("对方号码")["本方号码"].nunique()
    else:
        result["涉及本方号码数"] = 0

    return result.reset_index().sort_values("次数", ascending=False).head(top_n)


def get_night_calls(df, start_hour=0, end_hour=5):
    columns = ["本方号码", "对方号码", "开始时间", "呼叫类型", "通话时长"]
    if df is None or df.empty or "开始时间" not in df.columns:
        return _empty_df(columns)

    work_df = df.copy()
    work_df["_开始时间_dt"] = _to_datetime(work_df["开始时间"])
    work_df = work_df.dropna(subset=["_开始时间_dt"])
    if work_df.empty:
        return _empty_df(columns)

    hours = work_df["_开始时间_dt"].dt.hour
    if start_hour <= end_hour:
        mask = (hours >= start_hour) & (hours <= end_hour)
    else:
        mask = (hours >= start_hour) | (hours <= end_hour)

    available_cols = [c for c in columns if c in work_df.columns]
    return work_df.loc[mask, available_cols].sort_values("开始时间").head(200)


def get_risk_contacts(df, min_calls=5, night_weight=2):
    columns = ["本方号码", "对方号码", "总次数", "深夜次数", "短通话次数", "风险分"]
    if df is None or df.empty:
        return _empty_df(columns)

    required = {"本方号码", "对方号码"}
    if not required.issubset(df.columns):
        return _empty_df(columns)

    work_df = df.dropna(subset=["本方号码", "对方号码"]).copy()
    if work_df.empty:
        return _empty_df(columns)

    if "开始时间" in work_df.columns:
        dt = _to_datetime(work_df["开始时间"])
        work_df["_is_night"] = dt.dt.hour.between(0, 5).fillna(False)
    else:
        work_df["_is_night"] = False

    if "通话时长" in work_df.columns:
        duration = _to_duration_seconds(work_df["通话时长"])
        work_df["_is_short"] = duration.between(1, 5)
    else:
        work_df["_is_short"] = False

    result = (
        work_df.groupby(["本方号码", "对方号码"])
        .agg(
            总次数=("对方号码", "size"),
            深夜次数=("_is_night", "sum"),
            短通话次数=("_is_short", "sum"),
        )
        .reset_index()
    )
    result["风险分"] = result["总次数"] + result["深夜次数"] * night_weight + result["短通话次数"]

    return (
        result[result["总次数"] >= min_calls]
        .sort_values(["风险分", "总次数"], ascending=False)
        .head(50)
    )
