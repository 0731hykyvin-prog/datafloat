import pandas as pd
from collections import defaultdict


def build_call_graph(df):
    """
    构建通信关系图（双向合并）
    """

    if "本方号码" not in df.columns or "对方号码" not in df.columns:
        return None

    df = df.dropna(subset=["本方号码", "对方号码"])

    graph = defaultdict(int)

    # 遍历记录
    for _, row in df.iterrows():

        a = str(row["本方号码"])
        b = str(row["对方号码"])

        # 🔥关键：双向归一（排序后合并）
        key = tuple(sorted([a, b]))

        graph[key] += 1

    # 转 DataFrame
    result = pd.DataFrame([
        {"节点A": k[0], "节点B": k[1], "次数": v}
        for k, v in graph.items()
    ])

    return result