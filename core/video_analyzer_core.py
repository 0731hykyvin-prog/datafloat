"""
视频布控数据分析引擎
=====================
从 video_analyzer.py 提炼，去除 tkinter。
"""

import os
from datetime import datetime
from pathlib import Path

import pandas as pd


def process_video_data(input_folder, person_db_path, output_folder,
                       similarity_threshold=0.9, time_window=5, log_callback=None):
    """
    视频布控数据处理主函数。

    参数:
        input_folder: CSV 文件夹路径
        person_db_path: 人员库 Excel 路径（含"公民身份号码"和"姓名"列）
        output_folder: 输出文件夹
        similarity_threshold: 相似度阈值（默认 0.9）
        time_window: 去重时间窗口秒数（默认 5）
        log_callback: 日志回调函数

    返回: (输出文件路径, 统计 dict, 日志列表)
    """
    def log(msg):
        logs.append(msg)
        if log_callback:
            log_callback(msg)

    logs = []
    try:
        log("=" * 50)
        log("开始处理视频布控数据")
        log("=" * 50)

        column_names = [
            "目标类型", "预警任务类型", "预警任务名称", "预警任务目标",
            "告警设备", "告警时间", "预警任务原因", "证件类型", "证件号码", "相似度",
        ]

        root_path = Path(input_folder)
        csv_files = list(root_path.rglob("*.csv"))
        if not csv_files:
            log("未找到 CSV 文件")
            return None, {}, logs

        log(f"找到 {len(csv_files)} 个 CSV 文件")

        def clean_excel_string(value):
            if isinstance(value, str) and value.startswith('="') and value.endswith('"'):
                return value[2:-1]
            return value

        all_data = []
        for i, f in enumerate(csv_files):
            try:
                df = pd.read_csv(f, encoding="utf-8", skiprows=5,
                                 names=column_names, on_bad_lines="skip")
                df = df.dropna(how="all")
                if len(df) > 0:
                    for col in df.columns:
                        df[col] = df[col].apply(clean_excel_string)
                    all_data.append(df)
                    log(f"  [{i+1}/{len(csv_files)}] {f.name}: {len(df)}行 ✓")
            except Exception as e:
                log(f"  [{i+1}/{len(csv_files)}] {f.name}: ✗ {e}")

        if not all_data:
            log("没有成功读取任何文件")
            return None, {}, logs

        result = pd.concat(all_data, ignore_index=True)
        result["相似度"] = pd.to_numeric(result["相似度"], errors="coerce")
        result["告警时间"] = pd.to_datetime(result["告警时间"], errors="coerce")

        # 相似度筛选
        result = result.dropna(subset=["相似度"])
        result = result[result["相似度"] >= similarity_threshold].copy()
        log(f"相似度 ≥ {similarity_threshold}: {len(result)} 条")

        # 去重
        result = result.dropna(subset=["告警时间"])
        if len(result) > 0:
            deduped_list = []
            for id_num, group in result.groupby("证件号码"):
                group = group.sort_values("告警时间").copy()
                group["_cluster"] = 0
                cid, last = 0, None
                for idx in group.index:
                    t = group.loc[idx, "告警时间"]
                    if last is None or (t - last).total_seconds() > time_window:
                        cid += 1
                    group.loc[idx, "_cluster"] = cid
                    last = t
                # 每个 cluster 取第一条有告警设备的
                for _, sub in group.groupby("_cluster"):
                    non_empty = sub[sub["告警设备"].notna() & (sub["告警设备"] != "")]
                    row = non_empty.iloc[0] if len(non_empty) > 0 else sub.iloc[0]
                    deduped_list.append(row)

            result = pd.DataFrame(deduped_list)
            if "_cluster" in result.columns:
                result.drop(columns=["_cluster"], inplace=True)
        log(f"去重后: {len(result)} 条")

        # 深夜标签
        result["告警时间"] = pd.to_datetime(result["告警时间"], errors="coerce")
        result["小时"] = result["告警时间"].dt.hour
        result["深夜出行"] = result["小时"].apply(lambda x: 1 if (x >= 23 or x <= 5) else 0)
        night_count = int(result["深夜出行"].sum())
        log(f"深夜出行: {night_count} 条")

        # 统计
        night_stats = result[result["深夜出行"] == 1].groupby("证件号码").size().reset_index(name="深夜出行次数")
        total_stats = result.groupby("证件号码").size().reset_index(name="总出行次数")
        stats = total_stats.merge(night_stats, on="证件号码", how="left")
        stats["深夜出行次数"] = stats["深夜出行次数"].fillna(0).astype(int)
        stats["深夜出行占比"] = (stats["深夜出行次数"] / stats["总出行次数"] * 100).round(2)

        # 匹配姓名
        if person_db_path and os.path.exists(person_db_path):
            try:
                person_df = pd.read_excel(person_db_path)
                if "公民身份号码" in person_df.columns and "姓名" in person_df.columns:
                    name_map = dict(zip(person_df["公民身份号码"].astype(str), person_df["姓名"]))
                    stats["姓名"] = stats["证件号码"].astype(str).map(name_map).fillna("未匹配")
                    matched = (stats["姓名"] != "未匹配").sum()
                    log(f"姓名匹配: {matched}/{len(stats)} 人")
                else:
                    stats["姓名"] = "未匹配"
            except Exception as e:
                log(f"人员库读取失败: {e}")
                stats["姓名"] = "未匹配"
        else:
            stats["姓名"] = "未匹配"

        # 最终输出
        final = stats[["姓名", "证件号码", "深夜出行次数", "总出行次数", "深夜出行占比"]].copy()
        final = final.sort_values("深夜出行次数", ascending=False)
        final.insert(0, "序号", range(1, len(final) + 1))

        os.makedirs(output_folder, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_folder, f"视频布控分析报告_{ts}.xlsx")
        final.to_excel(output_path, index=False)

        summary = {
            "总人数": len(final),
            "深夜出行总次数": int(final["深夜出行次数"].sum()),
            "深夜出行人数": int((final["深夜出行次数"] > 0).sum()),
            "输出文件": output_path,
        }
        log(f"✅ 完成: {summary}")
        return output_path, summary, logs

    except Exception as e:
        log(f"❌ 失败: {e}")
        import traceback
        log(traceback.format_exc())
        return None, {}, logs
