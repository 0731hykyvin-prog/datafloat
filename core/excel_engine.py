import os
import sys

import pandas as pd

from core.mapper import apply_mapping
from core.transformers import fill_end_time, fill_local_number, normalize_datetime
from core.validator import analyze_quality, calculate_match_rate


def _get_output_dir():
    """返回 exe 同级目录（frozen 时）或当前工作目录，保证可写。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.getcwd()


def merge_excel_files(file_list, output_filename="merged.xlsx", mapping=None):
    """
    合并多个 Excel / CSV 文件，返回统一的 DataFrame 并写出 Excel。

    返回:
        dict 或 None — 包含 "df", "output_file", "logs" 三个键；
        logs 是字符串列表，调用方可写入 UI 日志框。
        None 表示没有可合并的数据。
    """
    logs = []
    dfs = []

    for file in file_list:
        try:
            ext = os.path.splitext(file)[1].lower()

            if ext in {".xlsx", ".xlsm"}:
                df = pd.read_excel(file, engine="openpyxl", dtype=str)
            elif ext == ".xls":
                df = pd.read_excel(file, engine="xlrd", dtype=str)
            elif ext == ".csv":
                df = pd.read_csv(file, dtype=str)
            else:
                continue

            # ── 1. 字段映射（如果提供） ──
            hit_report = {}
            missing_report = {}
            match_rate = 0

            if mapping:
                df, hit_report, missing_report = apply_mapping(df, mapping)
                match_rate = calculate_match_rate(mapping, hit_report)

            # ── 2. 本方号码补全（始终执行） ──
            df = fill_local_number(df, file)

            # ── 3. 时间标准化（始终执行） ──
            if "开始时间" in df.columns:
                df["开始时间"] = df["开始时间"].apply(normalize_datetime)

            # ── 4. 自动生成结束时间（始终执行） ──
            df = fill_end_time(df)

            # ── 5. 质量分析 ──
            standard_fields = list(mapping.keys()) if mapping else []
            report = analyze_quality(df, standard_fields)

            # ── 6. 收集日志 ──
            logs.append("=" * 50)
            logs.append(f"文件: {os.path.basename(file)}")

            if mapping:
                logs.append("")
                logs.append("📌 字段映射结果:")
                for k, v in hit_report.items():
                    logs.append(f"  ✔ {k} <- {v}")
                if missing_report:
                    logs.append("")
                    logs.append("⚠ 缺失字段:")
                    for k, v in missing_report.items():
                        logs.append(f"  ✘ {k} 未找到（尝试: {v}）")
                logs.append(f"")
                logs.append(f"📊 字段匹配率: {match_rate}%")

            logs.append("")
            logs.append("📊 话单质量报告:")
            logs.append(f"  完整率: {report['completeness']}%")
            logs.append(f"  时间质量: {report.get('time_quality', 0)}%")
            logs.append(f"  是否可分析: {report['usable']}")
            logs.append("  字段状态:")
            for k, v in report.get("field_status", {}).items():
                logs.append(f"    {k}: {v}")

            # ── 7. 统一为标准字段结构（始终执行） ──
            standard_columns = [
                "本方号码", "对方号码", "开始时间", "结束时间",
                "通话时长", "呼叫类型", "小区号", "基站号",
                "IMSI", "IMEI",
            ]
            df = df.reindex(columns=[c for c in standard_columns if c in df.columns])

            # ── 8. 标记来源文件 ──
            df["来源文件"] = os.path.basename(file)

            dfs.append(df)

        except Exception as e:
            logs.append(f"❌ 读取失败: {file}")
            logs.append(f"   {e}")

    if not dfs:
        logs.append("❌ 处理失败：没有可合并的数据")
        return {"df": pd.DataFrame(), "output_file": None, "logs": logs}

    # ── 合并 ──
    result = pd.concat(dfs, ignore_index=True)

    # ── 写出 Excel ──
    output_path = os.path.join(_get_output_dir(), output_filename)
    result.to_excel(output_path, index=False)

    logs.append("")
    logs.append(f"✅ 合并完成，共 {len(result)} 条记录")
    logs.append(f"📁 输出文件: {output_path}")

    return {
        "df": result,
        "output_file": output_path,
        "logs": logs,
    }
