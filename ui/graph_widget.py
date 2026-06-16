import os
import tempfile

# 必须最先设置 matplotlib 配置目录（PyInstaller 打包后文件系统可能是只读的）
os.environ.setdefault(
    "MPLCONFIGDIR",
    os.path.join(tempfile.gettempdir(), "datafloat_matplotlib"),
)

import matplotlib
import networkx as nx
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PySide2.QtWidgets import QVBoxLayout, QWidget

# 延迟设置后端 — 必须在 QApplication 创建之后才能调用
_backend_set = False


def _ensure_backend():
    global _backend_set
    if not _backend_set:
        matplotlib.use("Qt5Agg")
        _backend_set = True


class GraphWidget(QWidget):
    def __init__(self, graph_df=None):
        super().__init__()

        self._layout = QVBoxLayout()
        self.setLayout(self._layout)

        _ensure_backend()
        import matplotlib.pyplot as plt

        self.figure = plt.figure()
        self.canvas = FigureCanvas(self.figure)
        self._layout.addWidget(self.canvas)

        if graph_df is not None:
            self.draw_graph(graph_df)

    def draw_graph(self, df):
        import matplotlib.pyplot as plt

        _ensure_backend()

        if df is None or df.empty:
            return

        required_cols = {"节点A", "节点B", "次数"}
        if not required_cols.issubset(df.columns):
            return

        self.figure.clf()
        ax = self.figure.add_subplot(111)

        G = nx.Graph()

        for _, row in df.iterrows():
            try:
                a = str(row["节点A"])
                b = str(row["节点B"])
                w = float(row["次数"])
                if a == "" or b == "":
                    continue
                G.add_edge(a, b, weight=w)
            except Exception:
                continue

        if len(G.nodes) == 0:
            ax.set_title("无有效关系数据")
            self.canvas.draw()
            return

        pos = nx.spring_layout(G, k=0.8, seed=42)

        nx.draw_networkx_nodes(
            G, pos, node_size=800, node_color="lightblue", ax=ax
        )

        edges = G.edges(data=True)
        nx.draw_networkx_edges(
            G,
            pos,
            width=[max(0.5, d["weight"] * 0.2) for _, _, d in edges],
            ax=ax,
        )

        nx.draw_networkx_labels(G, pos, font_size=8, ax=ax)

        ax.set_title("通信关系图谱")
        ax.axis("off")
        self.canvas.draw()
