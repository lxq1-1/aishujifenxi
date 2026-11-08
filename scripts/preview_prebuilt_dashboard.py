"""
预览生成：模拟「图表编辑弹窗 · 前置可视化看板」高级科技风格效果。
- 顶部标题栏 + 4 个 KPI 卡 + 6 个图表卡（柱状/雷达/折线/仪表盘/面积/环形）
- 每张图表卡右上角两个小图标（下载PNG / + 添加至报表）
- 风格：深蓝 #0a1a3a 底 + 青绿 #2dd4bf / 霓虹 #38bdf8 / 紫粉渐变
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Wedge, Circle
import numpy as np

# 颜色（高级科技风格）
BG = "#0a1a3a"
PANEL = "#0f2349"
PANEL_EDGE = "#1e3a6f"
CYAN = "#2dd4bf"
NEON = "#38bdf8"
PURPLE = "#a855f7"
PINK = "#ec4899"
ORANGE = "#f59e0b"
TEXT = "#e2e8f0"
SUBTEXT = "#94a3b8"

plt.rcParams["font.sans-serif"] = ["PingFang SC", "Heiti TC", "Microsoft YaHei", "Arial Unicode MS", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

fig = plt.figure(figsize=(19.2, 10.8), facecolor=BG)

# ============= 顶部标题栏 =============
title_ax = fig.add_axes([0, 0.92, 1, 0.08])
title_ax.set_facecolor(BG)
title_ax.set_xlim(0, 1); title_ax.set_ylim(0, 1); title_ax.axis("off")
title_ax.text(0.5, 0.65, "AI 数据异常分析看板", ha="center", va="center", fontsize=28, color=CYAN, weight="bold")
title_ax.text(0.5, 0.18, "业绩复盘 · 异常追踪 · 贡献拆解", ha="center", va="center", fontsize=14, color=SUBTEXT)
# 顶部装饰线
title_ax.plot([0.05, 0.95], [0.95, 0.95], color=NEON, lw=1, alpha=0.6)

# ============= KPI 4 卡 =============
kpis = [
    ("总营收", "¥1,286万", "+18.6%", CYAN),
    ("同比增长", "¥+18.6万", "+18.6%", NEON),
    ("异常波动", "12 处", "—", PURPLE),
    ("高风险项", "3 项", "—", PINK),
]
for i, (label, val, delta, color) in enumerate(kpis):
    x0 = 0.04 + i * 0.235
    ax = fig.add_axes([x0, 0.79, 0.21, 0.10])
    ax.set_facecolor(BG); ax.axis("off"); ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    box = FancyBboxPatch((0.02, 0.05), 0.96, 0.9, boxstyle="round,pad=0.02,rounding_size=0.04",
                         linewidth=1.2, edgecolor=PANEL_EDGE, facecolor=PANEL)
    ax.add_patch(box)
    ax.add_patch(mpatches.Rectangle((0.02, 0.05), 0.05, 0.9, color=color, alpha=0.85))
    ax.text(0.12, 0.72, label, fontsize=12, color=SUBTEXT)
    ax.text(0.12, 0.40, val, fontsize=20, color=TEXT, weight="bold")
    ax.text(0.85, 0.40, delta, fontsize=12, color=color, ha="right", weight="bold")

# ============= 图表卡片辅助函数 =============
def draw_card(x, y, w, h, title, ax_fn, accent=CYAN):
    # 卡片底框
    card_ax = fig.add_axes([x, y, w, h])
    card_ax.set_facecolor(BG); card_ax.axis("off"); card_ax.set_xlim(0, 1); card_ax.set_ylim(0, 1)
    box = FancyBboxPatch((0.005, 0.005), 0.99, 0.99, boxstyle="round,pad=0.01,rounding_size=0.03",
                         linewidth=1.3, edgecolor=PANEL_EDGE, facecolor=PANEL)
    card_ax.add_patch(box)
    # 标题
    card_ax.text(0.04, 0.92, title, fontsize=13, color=TEXT, weight="bold")
    card_ax.plot([0.04, 0.5], [0.88, 0.88], color=accent, lw=2, alpha=0.9)
    # 右上角两个小图标（下载PNG / + 添加至报表）
    # 下载
    card_ax.add_patch(mpatches.Rectangle((0.80, 0.92), 0.07, 0.06, facecolor=PANEL_EDGE, edgecolor=accent, lw=0.8))
    card_ax.text(0.835, 0.95, "↓PNG", ha="center", va="center", fontsize=7, color=TEXT)
    # 添加至报表
    card_ax.add_patch(mpatches.Rectangle((0.89, 0.92), 0.09, 0.06, facecolor=accent, edgecolor=accent, alpha=0.9))
    card_ax.text(0.935, 0.95, "+ 报表", ha="center", va="center", fontsize=7, color="#0a1a3a", weight="bold")
    # 内嵌图区
    inner = fig.add_axes([x + 0.02, y + 0.06, w - 0.04, h - 0.18])
    ax_fn(inner)

# ============= 6 张图表 =============

# 1 部门业绩对比 - 渐变柱
def chart1(ax):
    ax.set_facecolor(PANEL)
    cats = ["销售一部", "电商二部", "线下三部"]
    vals = [240, 185, -96]
    colors = [CYAN, NEON, PINK]
    bars = ax.bar(cats, vals, color=colors, edgecolor="none", width=0.55)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width()/2, v + (8 if v >= 0 else -18), f"{v:+d}万", ha="center",
                fontsize=10, color=TEXT, weight="bold")
    ax.set_ylim(-130, 290)
    ax.tick_params(colors=SUBTEXT, labelsize=9)
    ax.set_facecolor(PANEL)
    for s in ax.spines.values(): s.set_color(PANEL_EDGE)
    ax.grid(axis="y", color=PANEL_EDGE, lw=0.5, alpha=0.4)
    ax.set_axisbelow(True)

# 2 异常关联 - 雷达
def chart2(ax):
    ax.set_facecolor(PANEL)
    labels = ["销售", "成本", "客户", "库存", "人力", "渠道"]
    vals = [0.92, 0.78, 0.65, 0.85, 0.55, 0.70]
    angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist()
    vals += vals[:1]; angles += angles[:1]
    ax = plt.gca(); ax.remove()
    polar = fig.add_axes([0.355, 0.12, 0.275, 0.45], projection="polar")
    polar.set_facecolor(PANEL)
    polar.plot(angles, vals, color=CYAN, lw=2)
    polar.fill(angles, vals, color=CYAN, alpha=0.25)
    polar.set_xticks(angles[:-1]); polar.set_xticklabels(labels, color=SUBTEXT, fontsize=9)
    polar.set_yticks([0.2, 0.4, 0.6, 0.8, 1.0])
    polar.set_yticklabels(["", "", "", "", ""], color=SUBTEXT, fontsize=7)
    polar.grid(color=PANEL_EDGE, alpha=0.6)
    polar.spines["polar"].set_color(PANEL_EDGE)

# 3 月度新增趋势 - 紫粉渐变柱
def chart3(ax):
    ax.set_facecolor(PANEL)
    months = ["1月", "2月", "3月", "4月", "5月", "6月"]
    vals = [120, 185, 220, 260, 198, 240]
    cmap = plt.cm.colors.LinearSegmentedColormap.from_list("pp", [PURPLE, PINK])
    colors = [cmap(i/5) for i in range(6)]
    ax.bar(months, vals, color=colors, edgecolor="none", width=0.55)
    ax.tick_params(colors=SUBTEXT, labelsize=9)
    for s in ax.spines.values(): s.set_color(PANEL_EDGE)
    ax.grid(axis="y", color=PANEL_EDGE, lw=0.5, alpha=0.4)
    ax.set_axisbelow(True)

# 4 整体健康度 - 仪表盘
def chart4(ax):
    ax.set_facecolor(PANEL); ax.axis("off")
    ax.set_xlim(-1.2, 1.2); ax.set_ylim(-1.2, 1.2)
    # 背景半圆
    theta = np.linspace(np.pi, 0, 100)
    ax.plot(np.cos(theta), np.sin(theta), color=PANEL_EDGE, lw=14, solid_capstyle="round")
    # 进度
    score = 0.82
    end = np.pi * (1 - score)
    theta2 = np.linspace(np.pi, end, 100)
    ax.plot(np.cos(theta2), np.sin(theta2), color=CYAN, lw=14, solid_capstyle="round")
    ax.text(0, 0.05, "82", ha="center", va="center", fontsize=34, color=TEXT, weight="bold")
    ax.text(0, -0.30, "良好", ha="center", va="center", fontsize=13, color=CYAN)

# 5 业绩走势与异常点 - 面积+折线+标点
def chart5(ax):
    ax.set_facecolor(PANEL)
    x = np.arange(0, 12)
    y = 200 + 40*np.sin(x/1.5) + np.random.default_rng(2).normal(0, 6, 12)
    y[5] = 360; y[8] = 110  # 异常点
    ax.fill_between(x, y, 100, color=PURPLE, alpha=0.35)
    ax.plot(x, y, color=CYAN, lw=2.2, marker="o", markersize=4, markerfacecolor=NEON)
    # 异常红点
    ax.scatter([5, 8], [360, 110], color=PINK, s=80, zorder=5, edgecolor="white", lw=1.2)
    ax.text(5, 380, "异常↑", color=PINK, fontsize=8, ha="center")
    ax.text(8, 80, "异常↓", color=PINK, fontsize=8, ha="center")
    ax.set_xticks(x); ax.set_xticklabels([f"{i+1}日" for i in x], color=SUBTEXT, fontsize=8)
    ax.tick_params(colors=SUBTEXT, labelsize=8)
    for s in ax.spines.values(): s.set_color(PANEL_EDGE)
    ax.grid(color=PANEL_EDGE, lw=0.5, alpha=0.4)
    ax.set_axisbelow(True)

# 6 目标完成率 - 紫粉环形
def chart6(ax):
    ax.set_facecolor(PANEL); ax.axis("off")
    ax.set_xlim(-1.2, 1.2); ax.set_ylim(-1.2, 1.2)
    circle_out = Circle((0, 0), 0.9, facecolor="none", edgecolor=PANEL_EDGE, lw=10)
    ax.add_patch(circle_out)
    # 进度弧 - 76%
    theta = np.linspace(-np.pi/2, -np.pi/2 + 2*np.pi*0.76, 100)
    xs = 0.9*np.cos(theta); ys = 0.9*np.sin(theta)
    ax.plot(xs, ys, color=PURPLE, lw=10, solid_capstyle="round")
    # 第二圈浅
    theta2 = np.linspace(-np.pi/2 + 2*np.pi*0.76, -np.pi/2 + 2*np.pi*0.96, 100)
    ax.plot(0.9*np.cos(theta2), 0.9*np.sin(theta2), color=NEON, lw=10, solid_capstyle="round")
    ax.text(0, 0.08, "76%", ha="center", va="center", fontsize=22, color=TEXT, weight="bold")
    ax.text(0, -0.25, "8/16", ha="center", va="center", fontsize=11, color=SUBTEXT)
    # 右侧图例
    ax.text(1.05, 0.40, "● 新综完成率  3,420", color=SUBTEXT, fontsize=8, va="center")
    ax.text(1.05, 0.10, "● 新增客户    3,420", color=SUBTEXT, fontsize=8, va="center")
    ax.text(1.05, -0.20, "● 流失客户    816", color=SUBTEXT, fontsize=8, va="center")

# 布局 6 卡（grid 3列 x 2行）
card_w = 0.305; card_h = 0.30
positions = [
    (0.025, 0.46, "部门业绩对比", chart1, CYAN),
    (0.345, 0.46, "异常关联网络", chart2, NEON),
    (0.665, 0.46, "月度新增趋势", chart3, PURPLE),
    (0.025, 0.14, "整体健康度", chart4, CYAN),
    (0.345, 0.14, "业绩走势与异常点", chart5, PURPLE),
    (0.665, 0.14, "目标完成率", chart6, NEON),
]
for x, y, title, fn, accent in positions:
    draw_card(x, y, card_w, card_h, title, fn, accent)

# ============= 底部 AI 结论条 =============
foot = fig.add_axes([0, 0.02, 1, 0.10]); foot.set_facecolor(BG); foot.axis("off")
foot.set_xlim(0, 1); foot.set_ylim(0, 1)
foot.add_patch(FancyBboxPatch((0.02, 0.1), 0.96, 0.8, boxstyle="round,pad=0.02,rounding_size=0.03",
                              linewidth=1, edgecolor=PANEL_EDGE, facecolor=PANEL))
foot.text(0.04, 0.55, "AI 分析结论", fontsize=12, color=CYAN, weight="bold")
foot.text(0.04, 0.22, "整体业绩同比增长正向增长，增长主要来自销售一部与电商二部；线下三部负增长并伴随较高异常波动。",
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