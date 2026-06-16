# core/validator.py

def calculate_match_rate(mapping, hit_report):

    total = len(mapping)

    success = len(hit_report)

    if total == 0:
        return 0

    return round(success / total * 100, 2)
def analyze_quality(df, standard_fields):
    """
    输出话单质量报告
    """

    report = {}

    total_fields = len(standard_fields)
    hit_fields = 0

    field_status = {}

    for field in standard_fields:

        if field in df.columns and df[field].notna().any():

            field_status[field] = "✔"
            hit_fields += 1

        else:

            field_status[field] = "✘"

    # 完整率
    completeness = round(hit_fields / total_fields * 100, 2)

    report["field_status"] = field_status
    report["completeness"] = completeness

    # 时间质量
    if "开始时间" in df.columns:

        valid_time = df["开始时间"].notna().sum()
        total = len(df)

        report["time_quality"] = round(valid_time / total * 100, 2)
    else:
        report["time_quality"] = 0

    # 是否可分析
    report["usable"] = completeness >= 60

    return report