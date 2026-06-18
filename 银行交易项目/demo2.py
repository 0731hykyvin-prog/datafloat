"""
快进快出筛选模块
=================
同一"交易时间"+"交易流水号"下，同时存在"进"和"出"记录时，
各保留一条，其余剔除。

此为第一步筛选，后续可叠加深夜交易、敏感金额等分析。
"""

import os
import pandas as pd

# ══════════════════════════════════════════════════════
# 配置
# ══════════════════════════════════════════════════════

INPUT_FILE = "/Users/kyvinmacbook/Desktop/禁毒项目/交易明细_最终结果_20260608_195745.xlsx"
OUTPUT_DIR = "快进快出分析结果"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ══════════════════════════════════════════════════════
# 读取数据
# ══════════════════════════════════════════════════════
print("正在读取数据...")
df = pd.read_excel(INPUT_FILE, dtype=str)

total_before = len(df)
print(f"原始数据: {total_before:,} 行")
print(f"收付标志分布: 进={(df['收付标志']=='进').sum():,} , 出={(df['收付标志']=='出').sum():,}")

# ══════════════════════════════════════════════════════
# 第一步：快进快出筛选
# 规则：同一 交易时间 + 交易流水号 → 保留 进/出 各一条
# ══════════════════════════════════════════════════════

print("\n" + "=" * 50)
print("第一步：快进快出筛选")
print("规则: 同一交易时间+流水号，进/出各保留1条")
print("=" * 50)

# 1. 剔除流水号为空的记录（无法判断是否同一笔）
valid = df.dropna(subset=["交易流水号", "交易时间", "收付标志"]).copy()
dropped_null = total_before - len(valid)
print(f"\n  ├─ 流水号为空的记录: {dropped_null:,} 条（剔除）")

# 2. 找出同时有 进+出 的 (交易时间, 流水号) 组合
g = valid.groupby(["交易时间", "交易流水号"])
has_both = g["收付标志"].transform(
    lambda x: ("进" in x.values) and ("出" in x.values)
)
candidates = valid[has_both]

total_candidates = len(candidates)
unique_groups = candidates.groupby(["交易时间", "交易流水号"]).ngroups
print(f"  ├─ 快进快出候选记录: {total_candidates:,} 条")
print(f"  ├─ 涉及唯一（时间+流水号）组合: {unique_groups:,} 组")

# 3. 每组只保留 进1条 + 出1条
result_rows = []
for _, group in candidates.groupby(["交易时间", "交易流水号"]):
    jin_rows = group[group["收付标志"] == "进"]
    chu_rows = group[group["收付标志"] == "出"]

    # 各取第一条
    if len(jin_rows) > 0:
        result_rows.append(jin_rows.iloc[0])
    if len(chu_rows) > 0:
        result_rows.append(chu_rows.iloc[0])

result = pd.DataFrame(result_rows)
dropped_duplicates = total_candidates - len(result)
print(f"  ├─ 每组去重后保留: {len(result):,} 条（每组进1条 + 出1条）")
print(f"  └─ 剔除的重复项: {dropped_duplicates:,} 条")

# ══════════════════════════════════════════════════════
# 第二步：敏感金额筛选
# 规则：交易金额 ≥ 500 且金额 % 100 == 0（即 500 为基准，100 为公差的数列）
# ══════════════════════════════════════════════════════

print("\n" + "=" * 50)
print("第二步：敏感金额筛选")
print("规则: 金额 ≥ 500 且为 100 的整数倍")
print("=" * 50)

# 转换金额为数值
result["交易金额_num"] = pd.to_numeric(result["交易金额"], errors="coerce").fillna(0)

sensitive_mask = (result["交易金额_num"] >= 500) & (result["交易金额_num"] % 100 == 0)
result_step2 = result[sensitive_mask].copy()
result_step2.drop(columns=["交易金额_num"], inplace=True)

dropped_amount = len(result) - len(result_step2)
print(f"\n  ├─ 快进快出结果: {len(result):,} 条")
print(f"  ├─ 金额不满足条件的: {dropped_amount:,} 条（剔除）")
print(f"  ├─ 敏感金额保留: {len(result_step2):,} 条")
print(f"  ├─ 涉及账户: {result_step2['交易账号'].nunique():,} 个")
print(f"  └─ 涉及人员: {result_step2['账户开户名称'].nunique():,} 人")

# 金额分布
amounts = pd.to_numeric(result_step2["交易金额"], errors="coerce").dropna()
if len(amounts) > 0:
    print(f"\n  金额分布: 最小={amounts.min():.0f}  最大={amounts.max():.0f}  中位={amounts.median():.0f}")
    # Top 10 高频金额
    top_amounts = amounts.value_counts().head(10)
    print("  高频金额 TOP 10:")
    for amt, cnt in top_amounts.items():
        print(f"    {amt:>10.0f}  →  {cnt:>6,} 次")

# ══════════════════════════════════════════════════════
# 第三步：高频账户筛选
# 规则：前两步过滤后，交易次数 > 50 条的账户才保留
# ══════════════════════════════════════════════════════

print("\n" + "=" * 50)
print("第三步：高频账户筛选")
print("规则: 命中前两步的条数 > 50 条")
print("=" * 50)

# 按交易账号统计命中次数
account_counts = result_step2["交易账号"].value_counts()
high_freq_accounts = account_counts[account_counts > 50].index.tolist()

result_step3 = result_step2[result_step2["交易账号"].isin(high_freq_accounts)].copy()

dropped_low = len(result_step2) - len(result_step3)
print(f"\n  ├─ 上一步结果: {len(result_step2):,} 条")
print(f"  ├─ ≤50 条的低频账户: {dropped_low:,} 条（剔除）")
print(f"  ├─ 高频账户保留: {len(result_step3):,} 条")
print(f"  ├─ 涉及账户: {result_step3['交易账号'].nunique():,} 个")
print(f"  └─ 涉及人员: {result_step3['账户开户名称'].nunique():,} 人")

# 每个高频账户的明细
print(f"\n  高频账户 TOP 10:")
for acc in high_freq_accounts[:10]:
    acc_data = result_step3[result_step3["交易账号"] == acc]
    name = acc_data["账户开户名称"].iloc[0]
    jin = (acc_data["收付标志"] == "进").sum()
    chu = (acc_data["收付标志"] == "出").sum()
    print(f"    {acc}  {name}  总{len(acc_data):>4,}条  进{jin:>4,}  出{chu:>4,}")

# ══════════════════════════════════════════════════════
# 导出结果
# ══════════════════════════════════════════════════════

output_file_1 = os.path.join(OUTPUT_DIR, "01_快进快出筛选结果.xlsx")
result.to_excel(output_file_1, index=False)

output_file_2 = os.path.join(OUTPUT_DIR, "02_快进快出+敏感金额.xlsx")
result_step2.to_excel(output_file_2, index=False)

output_file_3 = os.path.join(OUTPUT_DIR, "03_高频账户_大于50条.xlsx")
result_step3.to_excel(output_file_3, index=False)

# ══════════════════════════════════════════════════════
# 统计摘要
# ══════════════════════════════════════════════════════

jin_count = (result["收付标志"] == "进").sum()
chu_count = (result["收付标志"] == "出").sum()
jin2 = (result_step2["收付标志"] == "进").sum()
chu2 = (result_step2["收付标志"] == "出").sum()
jin3 = (result_step3["收付标志"] == "进").sum()
chu3 = (result_step3["收付标志"] == "出").sum()

print("\n" + "=" * 50)
print("筛选完成 — 总览")
print("=" * 50)
print(f"{'原始数据:':<20s} {total_before:>10,} 条")
print(f"{'① 快进快出:':<20s} {len(result):>10,} 条  (进 {jin_count:,} / 出 {chu_count:,})")
print(f"{'② +敏感金额:':<20s} {len(result_step2):>10,} 条  (进 {jin2:,} / 出 {chu2:,})")
print(f"{'③ +高频(>50条):':<20s} {len(result_step3):>10,} 条  (进 {jin3:,} / 出 {chu3:,})")
print(f"{'涉及账户:':<20s} {result_step3['交易账号'].nunique():>10,} 个")
print(f"{'涉及人员:':<20s} {result_step3['账户开户名称'].nunique():>10,} 人")
print(f"\n输出文件:")
print(f"  ① {output_file_1}")
print(f"  ② {output_file_2}")
print(f"  ③ {output_file_3}")

# ══════════════════════════════════════════════════════
# 第四步：人员汇总表（仅账户开户名称 + 证件号码）
# ══════════════════════════════════════════════════════

person_summary = (
    result_step3[["账户开户名称", "开户人证件号码"]]
    .drop_duplicates()
    .sort_values("账户开户名称")
    .reset_index(drop=True)
)

output_file_4 = os.path.join(OUTPUT_DIR, "04_人员汇总表.xlsx")
person_summary.to_excel(output_file_4, index=False)

print(f"  ④ {output_file_4}")
print(f"\n人员名单 ({len(person_summary)} 人):")
for _, row in person_summary.iterrows():
    print(f"  {row['账户开户名称']:<8s}  {row['开户人证件号码']}")
