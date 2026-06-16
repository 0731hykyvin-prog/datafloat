import os
import pandas as pd
from datetime import time

# =========================
# 文件路径
# =========================

input_file = r"/Users/kyvinmacbook/Desktop/禁毒项目/交易明细_最终结果_20260608_195745.xlsx"

output_dir = "交易分析结果"
os.makedirs(output_dir, exist_ok=True)

# =========================
# 读取数据
# =========================

df = pd.read_excel(
    input_file,
    dtype=str
)

df["交易金额"] = pd.to_numeric(
    df["交易金额"],
    errors="coerce"
).fillna(0)

df["交易时间"] = pd.to_datetime(
    df["交易时间"],
    errors="coerce"
)

# =========================
# 时间处理
# =========================

df["交易时间"] = pd.to_datetime(
    df["交易时间"],
    errors="coerce"
)

# =========================
# 深夜交易判断
# 21:00~05:00
# =========================

def is_night_trade(dt):
    if pd.isna(dt):
        return 0

    t = dt.time()

    if t >= time(21, 0, 0):
        return 1

    if t < time(5, 0, 0):
        return 1

    return 0


df["深夜交易标志"] = df["交易时间"].apply(is_night_trade)

# =========================
# 金额处理
# =========================

df["交易金额"] = pd.to_numeric(
    df["交易金额"],
    errors="coerce"
).fillna(0)

# =========================
# 敏感金额判断
# >=500 且 50倍数
# =========================

df["敏感金额标志"] = (
    (df["交易金额"] >= 500)
    &
    (df["交易金额"] % 50 == 0)

).astype(int)

# ==================================================
# 01 深夜交易明细（仅统计进账，流水号去重）
# ==================================================

night_df = df[
    (df["深夜交易标志"] == 1)
    &
    (df["收付标志"] == "进")
].copy()

night_df = night_df.drop_duplicates(
    subset=["交易流水号"]
)

night_df.to_excel(
    os.path.join(
        output_dir,
        "01_深夜交易明细.xlsx"
    ),
    index=False
)

# ==================================================
# 02 敏感金额交易明细（仅统计进账，流水号去重）
# ==================================================

sensitive_df = df[
    (df["敏感金额标志"] == 1)
    &
    (df["收付标志"] == "进")
    & (df["对手户名"] != df["账户开户名称"])
    
].copy()

sensitive_df = sensitive_df.drop_duplicates(
    subset=["交易流水号"]
)

sensitive_df.to_excel(
    os.path.join(
        output_dir,
        "02_敏感金额交易明细.xlsx"
    ),
    index=False
)

# ==================================================
# 03 深夜交易次数统计
# ==================================================

night_count_df = (
    night_df
    .groupby(
        ["交易账号", "账户开户名称"],
        dropna=False
    )
    .size()
    .reset_index(name="深夜交易次数")
    .sort_values(
        "深夜交易次数",
        ascending=False
    )
)

night_count_df.to_excel(
    os.path.join(
        output_dir,
        "03_深夜交易次数统计.xlsx"
    ),
    index=False
)

# ==================================================
# 04 深夜+敏感金额（仅统计进账，流水号去重）
# ==================================================

night_sensitive_df = df[
    (df["深夜交易标志"] == 1)
    &
    (df["敏感金额标志"] == 1)
    &
    (df["收付标志"] == "进")
].copy()

night_sensitive_df = night_sensitive_df.drop_duplicates(
    subset=["交易流水号"]
)

night_sensitive_df.to_excel(
    os.path.join(
        output_dir,
        "04_深夜敏感金额交易.xlsx"
    ),
    index=False
)

# ==================================================
# 05 ATM取款明细
# ==================================================

atm_df = df[
    df["摘要说明"]
    .fillna("")
    .str.contains(
        "取款",
        na=False
    )
].copy()

atm_df.to_excel(
    os.path.join(
        output_dir,
        "05_ATM取款明细.xlsx"
    ),
    index=False
)

# ==================================================
# 06 配送交易明细
# ==================================================

delivery_keywords = [
    "uu跑腿",
    "京东秒送",
    "翠鸟配送",
    "顺丰速运",
    "达达秒送"
]

delivery_df = df[
    df["对手户名"]
    .fillna("")
    .str.contains(
        "|".join(delivery_keywords),
        case=False,
        na=False
    )
].copy()

delivery_df.to_excel(
    os.path.join(
        output_dir,
        "06_配送交易明细.xlsx"
    ),
    index=False
)

# ==================================================
# 统计模块
# ==================================================

night_stat = (
    night_df
    .groupby(
        ["交易账号", "账户开户名称"],
        dropna=False
    )
    .size()
    .reset_index(name="深夜交易次数")
)

sensitive_stat = (
    sensitive_df
    .groupby(
        ["交易账号", "账户开户名称"],
        dropna=False
    )
    .size()
    .reset_index(name="敏感金额次数")
)

night_sensitive_stat = (
    night_sensitive_df
    .groupby(
        ["交易账号", "账户开户名称"],
        dropna=False
    )
    .size()
    .reset_index(name="深夜敏感金额次数")
)

atm_stat = (
    atm_df
    .groupby(
        ["交易账号", "账户开户名称"],
        dropna=False
    )
    .size()
    .reset_index(name="ATM取款次数")
)

delivery_stat = (
    delivery_df
    .groupby(
        ["交易账号", "账户开户名称"],
        dropna=False
    )
    .size()
    .reset_index(name="配送交易次数")
)

# ==================================================
# 风险评分
# ==================================================

risk_df = (
    night_stat
    .merge(
        sensitive_stat,
        how="outer",
        on=["交易账号", "账户开户名称"]
    )
    .merge(
        night_sensitive_stat,
        how="outer",
        on=["交易账号", "账户开户名称"]
    )
    .merge(
        atm_stat,
        how="outer",
        on=["交易账号", "账户开户名称"]
    )
    .merge(
        delivery_stat,
        how="outer",
        on=["交易账号", "账户开户名称"]
    )
)

risk_df = risk_df.fillna(0)

# ==================================================
# 风险评分公式
# ==================================================

risk_df["风险评分"] = (
    risk_df["深夜交易次数"] * 1
    +
    risk_df["敏感金额次数"] * 1
    +
    risk_df["深夜敏感金额次数"] * 2
    +
    risk_df["ATM取款次数"] * 2
    +
    risk_df["配送交易次数"] * 2
)

risk_df = risk_df.sort_values(
    "风险评分",
    ascending=False
)

risk_df.to_excel(
    os.path.join(
        output_dir,
        "07_账户风险评分.xlsx"
    ),
    index=False
)

# ==================================================
# 08 人员风险评分汇总
# ==================================================

person_risk_df = (
    risk_df
    .groupby(
        "账户开户名称",
        dropna=False
    )
    .agg({
        "深夜交易次数":"sum",
        "敏感金额次数":"sum",
        "深夜敏感金额次数":"sum",
        "ATM取款次数":"sum",
        "配送交易次数":"sum",
        "风险评分":"sum"
    })
    .reset_index()
)

person_risk_df = person_risk_df.sort_values(
    "风险评分",
    ascending=False
)

person_risk_df.to_excel(
    os.path.join(
        output_dir,
        "08_人员风险评分汇总.xlsx"
    ),
    index=False
)
# ==================================================
# 完成
# ==================================================

print("=" * 50)
print("分析完成")
print(f"输出目录：{output_dir}")
print("=" * 50)
print(f"总记录数：{len(df):,}")
print(f"深夜交易：{len(night_df):,}")
print(f"敏感金额：{len(sensitive_df):,}")
print(f"深夜敏感金额：{len(night_sensitive_df):,}")
print(f"ATM取款：{len(atm_df):,}")
print(f"配送交易：{len(delivery_df):,}")
print("=" * 50)