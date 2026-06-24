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
# 通话时长 → 秒
def normalize_duration(value):
    """
    将各种格式的通话时长统一转为秒数（int）。
    支持: 纯秒"120" / 分秒"2分30秒" / 时分秒"1时2分3秒" / 浮点"120.0" / Excel时间"0:02:30"
    """
    if pd.isna(value):
        return None

    raw = str(value).strip()
    if not raw or raw in ("nan", "None", ""):
        return None

    # 尝试直接转 float（纯秒"120" 或 "120.0"）
    try:
        return int(round(float(raw)))
    except ValueError:
        pass

    # "1时2分3秒" / "2分30秒" / "30秒"
    total = 0
    h = re.search(r'(\d+)\s*时', raw)
    m = re.search(r'(\d+)\s*分', raw)
    s = re.search(r'(\d+)\s*秒', raw)
    if h:
        total += int(h.group(1)) * 3600
    if m:
        total += int(m.group(1)) * 60
    if s:
        total += int(s.group(1))
    if total > 0:
        return total

    # Excel 时间格式 "0:02:30" (时:分:秒) 或 "02:30" (分:秒)
    parts = raw.replace("：", ":").split(":")
    if len(parts) == 3:
        return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    if len(parts) == 2:
        return int(parts[0]) * 60 + int(parts[1])

    return 0
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
    """
    有"通话时长"但无"结束时间"时：
    1. 将通话时长统一转为秒
    2. 开始时间 + 时长秒 = 结束时间
    """
    if "开始时间" not in df.columns or "通话时长" not in df.columns:
        return df

    # 如果已有有效的结束时间，跳过
    if "结束时间" in df.columns and df["结束时间"].notna().any():
        return df

    end_times = []
    skipped = 0

    for _, row in df.iterrows():
        try:
            start = pd.to_datetime(row["开始时间"])
            secs = normalize_duration(row["通话时长"])
            if secs is None or secs < 0:
                end_times.append(None)
                skipped += 1
            else:
                end_times.append((start + timedelta(seconds=secs)).strftime("%Y-%m-%d %H:%M:%S"))
        except Exception:
            end_times.append(None)
            skipped += 1

    df["结束时间"] = end_times
    generated = len(end_times) - skipped
    if generated > 0:
        print(f"自动生成结束时间: {generated}/{len(end_times)}条")
    return df