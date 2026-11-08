"""
预览：通用「数据分析看板 · 高级科技风格」
主题：销售趋势 / 品类占比 / 区域分布 / 用户活跃 / 库存周转 / 转化漏斗
- 顶部标题 + 4 KPI 卡 + 6 图表卡 + 底部 AI 结论条
- 每张图表卡右上角：↓PNG / + 报表
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Circle
import numpy as np

BG = "#0a1a3a"; PANEL = "#0f2349"; PANEL_EDGE = "#1e3a6f"
CYAN = "#2dd4bf"; NEON = "#38bdf8"; PURPLE = "#a855f7"; PINK = "#ec4899"
ORANGE = "#f59e0b"; LIME = "#84cc16"; TEXT = "#e2e8f0"; SUBTEXT = "#94a3b8"

plt.rcParams["font.sans-serif"] = ["PingFang SC", "Heiti TC", "Microsoft YaHei", "Arial Unicode MS", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

fig = plt.figure(figsize=(19.2, 10.8), facecolor=BG)

# === 顶部标题 ===
t = fig.add_axes([0, 0.92, 1, 0.08]); t.set_facecolor(BG); t.axis("off")
t.set_xlim(0, 1); t.set_ylim(0, 1)
t.text(0.5, 0.65, "数据分析看板", ha="center", va="center", fontsize=30, color=CYAN, weight="bold")
t.text(0.5, 0.18, "销售 · 品类 · 区域 · 用户 · 库存 · 转化", ha="center", va="center", fontsize=14, color=SUBTEXT)
t.plot([0.05, 0.95], [0.95, 0.95], color=NEON, lw=1, alpha=0.6)

# === 4 KPI ===
kpis = [
    ("总销售额", "¥826.5万", "+12.3%", CYAN),
    ("订单数", "18,420", "+8.1%", NEON),
    ("活跃用户", "9,635", "+5.6%", PURPLE),
    ("客单价", "¥448.7", "-1.2%", PINK),
]
for i, (label, val, delta, color) in enumerate(kpis):
    x0 = 0.04 + i * 0.235
    ax = fig.add_axes([x0, 0.79, 0.21, 0.10])
    ax.set_facecolor(BG); ax.axis("off"); ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.add_patch(FancyBboxPatch((0.02, 0.05), 0.96, 0.9, boxstyle="round,pad=0.02,rounding_size=0.04",
                                linewidth=1.2, edgecolor=PANEL_EDGE, facecolor=PANEL))
    ax.add_patch(mpatches.Rectangle((0.02, 0.05), 0.05, 0.9, color=color, alpha=0.85))
    ax.text(0.12, 0.72, label, fontsize=12, color=SUBTEXT)
    ax.text(0.12, 0.40, val, fontsize=20, color=TEXT, weight="bold")
    ax.text(0.95, 0.40, delta, fontsize=12, color=color, ha="right", weight="bold")

def draw_card(x, y, w, h, title, ax_fn, accent=CYAN):
    cax = fig.add_axes([x, y, w, h]); cax.set_facecolor(BG); cax.axis("off")
    cax.set_xlim(0, 1); cax.set_ylim(0, 1)
    cax.add_patch(FancyBboxPatch((0.005, 0.005), 0.99, 0.99, boxstyle="round,pad=0.01,rounding_size=0.03",
                                 linewidth=1.3, edgecolor=PANEL_EDGE, facecolor=PANEL))
    cax.text(0.04, 0.92, title, fontsize=13, color=TEXT, weight="bold")
    cax.plot([0.04, 0.5], [0.88, 0.88], color=accent, lw=2, alpha=0.9)
    cax.add_patch(mpatches.Rectangle((0.80, 0.92), 0.07, 0.06, facecolor=PANEL_EDGE, edgecolor=accent, lw=0.8))
    cax.text(0.835, 0.95, "↓PNG", ha="center", va="center", fontsize=7, color=TEXT)
    cax.add_patch(mpatches.Rectangle((0.89, 0.92), 0.09, 0.06, facecolor=accent, edgecolor=accent, alpha=0.9))
    cax.text(0.935, 0.95, "+ 报表", ha="center", va="center", fontsize=7, color="#0a1a3a", weight="bold")
    fig.add_axes([x + 0.02, y + 0.06, w - 0.04, h - 0.18])
    ax_fn(plt.gca())

# 1 月度销售趋势 - 双线（销售额 / 订单数）
def c1(ax):
    ax.set_facecolor(PANEL)
    m = np.arange(1, 13)
    sales = 60 + 8*np.sin(m/1.6) + np.random.default_rng(1).normal(0, 3, 12)
    orders = 1200 + 150*np.cos(m/2.0) + np.random.default_rng(2).normal(0, 30, 12)
    ax.plot(m, sales, color=CYAN, lw=2.2, marker="o", markersize=4, label="销售额(万)")
    ax.plot(m, orders/10, color=NEON, lw=2.2, marker="s", markersize=4, label="订单数(十)", alpha=0.7)
    ax.fill_between(m, sales, alpha=0.15, color=CYAN)
    ax.set_xticks(m); ax.set_xticklabels([f"{i}月" for i in m], color=SUBTEXT, fontsize=8)
    ax.tick_params(colors=SUBTEXT, labelsize=8)
    for s in ax.spines.values(): s.set_color(PANEL_EDGE)
    ax.grid(color=PANEL_EDGE, lw=0.5, alpha=0.4); ax.set_axisbelow(True)
    ax.legend(loc="upper left", facecolor=PANEL, edgecolor=PANEL_EDGE, labelcolor=TEXT, fontsize=8)

# 2 品类占比 - 环形
def c2(ax):
    ax.set_facecolor(PANEL); ax.axis("off"); ax.set_xlim(-1.3, 1.3); ax.set_ylim(-1.2, 1.3)
    sizes = [38, 22, 18, 12, 10]
    labels = ["服饰 38%", "美妆 22%", "数码 18%", "食品 12%", "其他 10%"]
    colors = [CYAN, NEON, PURPLE, PINK, ORANGE]
    start = 90
    for s, c in zip(sizes, colors):
        end = start - s * 3.6
        w = mpatches.Wedge((0, 0), 0.85, end, start, width=0.32, facecolor=c, edgecolor=BG, lw=2)
        ax.add_patch(w)
        start = end
    ax.add_patch(Circle((0, 0), 0.45, facecolor=PANEL, edgecolor=PANEL_EDGE, lw=1))
    ax.text(0, 0.05, "826", ha="center", va="center", fontsize=22, color=TEXT, weight="bold")
    ax.text(0, -0.22, "万总额", ha="center", va="center", fontsize=10, color=SUBTEXT)
    for i, (lab, col) in enumerate(zip(labels, colors)):
        ax.plot([0.95, 1.05], [0.9 - i*0.22, 0.9 - i*0.22], color=col, lw=4)
        ax.text(1.08, 0.9 - i*0.22, lab, color=TEXT, fontsize=8, va="center")

# 3 区域销售分布 - 渐变横向柱
def c3(ax):
    ax.set_facecolor(PANEL)
    regions = ["华东", "华南", "华北", "西南", "东北", "西北"]
    vals = [285, 218, 156, 88, 52, 27]
    cmap = plt.cm.colors.LinearSegmentedColormap.from_list("cn", [NEON, CYAN])
    bars = ax.barh(regions, vals, color=cmap(np.linspace(0.2, 0.95, 6)), edgecolor="none")
    for b, v in zip(bars, vals):
        ax.text(v + 5, b.get_y() + b.get_height()/2, f"{v}万", va="center", fontsize=9, color=TEXT)
    ax.invert_yaxis()
    ax.tick_params(colors=SUBTEXT, labelsize=9)
    for s in ax.spines.values(): s.set_color(PANEL_EDGE)
    ax.grid(axis="x", color=PANEL_EDGE, lw=0.5, alpha=0.4); ax.set_axisbelow(True)

# 4 用户活跃度 - 多日热力（用矩形块模拟）
def c4(ax):
    ax.set_facecolor(PANEL); ax.set_xlim(0, 24); ax.set_ylim(0, 7)
    rng = np.random.default_rng(3)
    data = rng.uniform(0, 1, (7, 24))
    cmap = plt.cm.colors.LinearSegmentedColormap.from_list("hm", [PANEL, NEON, CYAN])
    for i in range(7):
        for j in range(24):
            color = cmap(data[i, j])
            ax.add_patch(mpatches.Rectangle((j, 6 - i), 0.9, 0.85, facecolor=color, edgecolor=PANEL))
    ax.set_xticks(np.arange(0, 24, 4) + 0.5)
    ax.set_xticklabels([f"{h}h" for h in range(0, 24, 4)], color=SUBTEXT, fontsize=8)
    ax.set_yticks(np.arange(0, 7) + 0.5)
    ax.set_yticklabels(["周日", "周六", "周五", "周四", "周三", "周二", "周一"], color=SUBTEXT, fontsize=8)
    ax.tick_params(colors=SUBTEXT, labelsize=8)
    for s in ax.spines.values(): s.set_color(PANEL_EDGE)

# 5 库存周转 Top 品类 - 渐变柱
def c5(ax):
    ax.set_facecolor(PANEL)
    cats = ["手机壳", "洗面奶", "蓝牙耳机", "运动鞋", "保温杯"]
    vals = [92, 78, 71, 65, 58]
    turnover = [12, 9, 14, 8, 11]
    cmap = plt.cm.colors.LinearSegmentedColormap.from_list("pl", [PURPLE, PINK])
    x = np.arange(len(cats))
    ax.bar(x - 0.18, vals, width=0.36, color=cmap(np.linspace(0.1, 0.9, 5)), label="库存量")
    ax.bar(x + 0.18, turnover, width=0.36, color=CYAN, alpha=0.85, label="周转天数")
    ax.set_xticks(x); ax.set_xticklabels(cats, color=SUBTEXT, fontsize=9)
    ax.tick_params(colors=SUBTEXT, labelsize=8)
    for s in ax.spines.values(): s.set_color(PANEL_EDGE)
    ax.grid(axis="y", color=PANEL_EDGE, lw=0.5, alpha=0.4); ax.set_axisbelow(True)
    ax.legend(loc="upper right", facecolor=PANEL, edgecolor=PANEL_EDGE, labelcolor=TEXT, fontsize=8)

# 6 转化漏斗 - 横向梯形条
def c6(ax):
    ax.set_facecolor(PANEL)
    stages = ["浏览", "加购", "下单", "支付", "复购"]
    vals = [100, 62, 38, 28, 14]
    colors = [CYAN, NEON, PURPLE, PINK, ORANGE]
    y = np.arange(len(stages))
    for i, (s, v, c) in enumerate(zip(stages, vals, colors)):
        ax.barh(i, v, height=0.7, color=c, alpha=0.85, edgecolor="none")
        ax.text(v + 1.5, i, f"{v}%", va="center", fontsize=10, color=TEXT, weight="bold")
        ax.text(0.5, i, s, va="center", ha="left", fontsize=10, color=BG, weight="bold")
    ax.set_yticks([]); ax.set_xticks([])
    ax.set_xlim(0, 115)
    for s in ax.spines.values(): s.set_visible(False)

card_w = 0.305; card_h = 0.30
positions = [
    (0.025, 0.46, "月度销售趋势", c1, CYAN),
    (0.345, 0.46, "品类销售占比", c2, NEON),
    (0.665, 0.46, "区域销售分布", c3, CYAN),
    (0.025, 0.14, "用户活跃时段", c4, NEON),
    (0.345, 0.14, "库存周转 TOP 品类", c5, PURPLE),
    (0.665, 0.14, "用户转化漏斗", c6, PINK),
]
for x, y, title, fn, accent in positions:
    draw_card(x, y, card_w, card_h, title, fn, accent)

# === 底部 ===
foot = fig.add_axes([0, 0.02, 1, 0.10]); foot.set_facecolor(BG); foot.axis("off")
foot.set_xlim(0, 1); foot.set_ylim(0, 1)
foot.add_patch(FancyBboxPatch((0.02, 0.1), 0.96, 0.8, boxstyle="round,pad=0.02,rounding_size=0.03",
                              linewidth=1, edgecolor=PANEL_EDGE, facecolor=PANEL))
foot.text(0.04, 0.55, "AI 总结", fontsize=12, color=CYAN, weight="bold")
foot.text(0.04, 0.22,
          "整体销售呈上升趋势，服饰类目贡献最高；华东、华南为核心市场；用户活跃集中在午间与晚间 20–22 点；蓝牙耳机周转最快，转化漏斗中加购到下单环节流失较大。",
          fontsize=11, color=TEXT)
foot.text(0.62, 0.55, "导出图片", fontsize=10, color=TEXT,
          bbox=dict(boxstyle="round,pad=0.4", fc=PANEL, ec=PANEL_EDGE))
foot.text(0.74, 0.55, "导出 PDF", fontsize=10, color=TEXT,
          bbox=dict(boxstyle="round,pad=0.4", fc=PANEL, ec=PANEL_EDGE))
foot.text(0.85, 0.55, "对话调整", fontsize=10, color=BG, weight="bold",
          bbox=dict(boxstyle="round,pad=0.4", fc=CYAN, ec=CYAN))

out = "/Users/xiahua/trea/数据/preview/prebuilt_dashboard_preview.png"
plt.savefig(out, dpi=120, facecolor=BG, bbox_inches="tight", pad_inches=0.1)
print("OK", out)