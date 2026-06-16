# core/transformers.py

import os
import re
import pandas as pd
from datetime import timedelta
# 本方号码补全
from core.utils import extract_phone_from_filename

def fill_local_number(df, file_path):
    """
    如果本方号码为空
    自动使用文件名中的手机号补全
    """

    phone = extract_phone_from_filename(
        file_path
    )

    if not phone:
        return df

    # 本方号码字段不存在
    if "本方号码" not in df.columns:

        df["本方号码"] = phone

        print(
            f"自动补全本方号码: {phone}"
        )

        return df

    # 字段存在但全部为空
    if df["本方号码"].isna().all():

        df["本方号码"] = phone

        print(
            f"自动补全本方号码: {phone}"
        )

    return df
# 标准化通话时长
def normalize_duration(value):

    if pd.isna(value):
        return None

    value = str(value).strip()

    # 纯秒

    if value.isdigit():
        return int(value)

    # xx分xx秒

    minute = 0
    second = 0

    m = re.search(r'(\d+)分', value)
    s = re.search(r'(\d+)秒', value)

    if m:
        minute = int(m.group(1))

    if s:
        second = int(s.group(1))

    return minute * 60 + second
# 标准化通话开始时间
def normalize_datetime(value):

    if pd.isna(value):
        return None

    try:

        return pd.to_datetime(
            value
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )

    except:

        return None
# 补全通话结束时间
def fill_end_time(df):

    if "开始时间" not in df.columns:
        return df

    if "通话时长" not in df.columns:
        return df

    if "结束时间" in df.columns:

        if df["结束时间"].notna().any():
            return df

    end_times = []

    for _, row in df.iterrows():

        try:

            start = pd.to_datetime(
                row["开始时间"]
            )

            duration = normalize_duration(
                row["通话时长"]
            )

            end_time = start + timedelta(
                seconds=duration
            )

            end_times.append(
                end_time.strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

        except:

            end_times.append(None)

    df["结束时间"] = end_times

    print("自动生成结束时间")

    return df