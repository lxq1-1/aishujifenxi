"""ECharts option → matplotlib PNG 渲染器。

背景：
    浏览器上的数据图表是用 ECharts 画的（基于 canvas，前端可交互）。
    但下载下来的 Word / Excel 是离线文件，python-docx + openpyxl 都没法
    在文档里复刻 ECharts 交互图——过去只能用 JSON 代码块占位，
    用户体验割裂。
    本模块做的事：
        把 ECharts 的 ``option`` JSON 解析出来,用 matplotlib 重新画一遍,
        输出成 PNG 字节,Word/Excel 里能嵌入真实图片。

支持的图表类型（与 ai_service._CHART_TYPES 白名单一致）：
    line / bar / pie / scatter / radar / heatmap / funnel

输入约定：
    每个 chart_option 由 ai_service.chart_edit() / generate_report()
    生成,大致形态为::

        {
          "title": {"text": "...", ...},
          "tooltip": {...},
          "legend": {...},
          "series": [
            {
              "type": "pie",
              "data": [{"name": "福州朴朴", "value": 13}, ...],
              "radius": ["40%", "70%"],   # pie 时此字段决定环形/饼图
            },
            ...
          ],
        }

输出：
    :func:`render_chart_to_png` 返回 ``bytes``,可直接喂给
    python-docx 的 ``add_picture()`` 或 openpyxl 的 ``add_image()``。
    图表渲染失败时返回 ``None``,调用方走"原 JSON 配置占位"分支。
"""
from __future__ import annotations

import contextvars
import io
import logging
from typing import Any

import matplotlib
# 强制 Agg 后端：服务端无显示器
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
# 【chart-visual-upgrade 新增】以下均为 matplotlib 自带子模块,不引入任何第三方依赖:
#   colors : LinearSegmentedColormap(柱图竖向渐变) / to_rgb / is_color_like(颜色校验)
#   lines  : Line2D(标题装饰线 + 图例代理句柄)
#   patches: FancyBboxPatch(圆角柱) / Patch(图例色块)
from matplotlib.colors import LinearSegmentedColormap, is_color_like, to_rgb
from matplotlib.lines import Line2D
from matplotlib.patches import FancyBboxPatch, Patch
import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)


# ============================================================
# 【chart-render-smart-2026-10 新增】汇总行识别 + 图表有效性校验
# ------------------------------------------------------------
# 背景:
#   用户反馈"图表编辑"生成的图表存在三类问题:
#     1) AI 不分析数据,每个 sheet 固定生成 6~8 张,导致出现空白图或重复图;
#     2) 部分图(年度目标 vs 年度实际 散点)无数据,仍然渲染出空图;
#     3) 环形图把"全年合计"与各月份并列做占比,合计行混入对比维度;
#   这里在 chart_renderer 层做三道前置处理:
#     a) 剥离"合计 / 总计 / 全年合计 / 汇总 / 平均"等汇总行
#     b) 每张图 append 前做有效性检查(空值/单点/全相等自动改图型或丢弃)
#     c) 根据 sheet 数据丰度自适应决定生成图表数量(不再是固定 6~8 张)
#     d) 【2026-10 新增】日期列若粒度过细(到秒/到天)→ 自动按日/月/季聚合,
#        避免出现"11 个不同时间戳=11 个等分切片"这种乱图
# ============================================================
# ============================================================

# 汇总行识别关键词:出现在分类列(通常是行首列)时,该行应排除出图表数据
_SUMMARY_KEYWORDS = (
    "合计", "总计", "全年合计", "全年总计", "总合计", "汇总",
    "平均", "总计平均", "合计平均", "小计", "总和小计",
    "total", "sum", "subtotal", "grand total", "average", "avg",
)


# ============================================================
# 图表主题上下文（2026-10 新增）
# ------------------------------------------------------------
# 背景:
#   同一个 render_chart_to_png 有两类使用场景,底色要求正好相反:
#     - 网页报告页 / 深色 UI   → 深底浅字(dark),与页面深蓝主题融为一体
#     - 导出 PPT / Word / 打印  → 白底深字(light),深色图块印出来很脏,
#                                而且深色图压在浅色文字上会看不清
#   之前图表底色是写死的 #0f172a(深色),导致导出的 PPT 页面是白底配深色图,
#   深色块还常常溢出页面边界压住正文 —— 用户反馈"排版不能直接用"。
#
# 做法:
#   用一个模块级上下文记录当前主题,各 _render_* 内部的文字色从这里取,
#   不必给每个渲染函数都加参数(会污染一堆签名)。
#   用 contextvars 而非普通全局变量,保证多线程(导出并发)互不干扰。
# ============================================================

# 当前主题下的图表配色。默认 dark(保持历史行为)。
_CHART_THEME: contextvars.ContextVar[str] = contextvars.ContextVar(
    "chart_theme", default="dark"
)

# 主题 → (背景, 网格/轴线, 刻度文字, 标题, 数据标签)
_CHART_PALETTES: dict[str, dict[str, str]] = {
    "dark": {
        "bg": "#0f172a", "grid": "#1e293b", "tick": "#94a3b8",
        "title": "#e2e8f0", "label": "#94a3b8", "muted": "#64748b",
    },
    "light": {
        "bg": "#ffffff", "grid": "#e5e7eb", "tick": "#4b5563",
        "title": "#111827", "label": "#4b5563", "muted": "#9ca3af",
    },
}


def _palette() -> dict[str, str]:
    """取当前主题的图表配色。"""
    return _CHART_PALETTES.get(_CHART_THEME.get(), _CHART_PALETTES["dark"])


# ============================================================
# 【chart-visual-upgrade 2026-10 新增】主题联动色盘
# ------------------------------------------------------------
# 背景(代码级根因):
#   模块级 _PALETTE 是写死的"科技霓虹"色板,_colors() 只认它,
#   **完全忽略 option.color**。而 apply_style() 会把 option.color 覆盖成
#   风格色盘,于是 _STYLE_PRESETS 里的 3 套风格
#   (tech_advanced 深蓝青绿 / business_pro 白底深蓝橙 / minimalist 浅灰蓝绿)
#   在导出的 PNG 里全部失效 —— 3 种风格导出来长得一模一样。
#
# 做法:
#   与 _CHART_THEME 完全同款:渲染期间用 contextvars 注入"当前这张图的色盘",
#   _colors() 优先取注入色盘,取不到才回退 _PALETTE。
#   选 contextvars 而非普通全局变量,是为了保证并发导出(多线程/多任务)
#   互不串色 —— 报表导出常常是批量并发的。
# ============================================================
_CHART_SERIES_COLORS: contextvars.ContextVar[list[str] | None] = contextvars.ContextVar(
    "chart_series_colors", default=None
)


def _is_summary_label(value: Any) -> bool:
    """判断一个分类标签值是否为「汇总行」(合计/总计/全年合计 等)。

    背景:用户上传的 Excel/CSV 经常在最末或最前出现「合计 / 全年合计 /
    总计 / 平均」这种汇总行,如果把它和「10月 / 11月 / 12月」一起塞进
    环形图,「合计」会被当作一个独立切片(月度占比图把全年合计当成
    第 13 个月)—— 这正是用户截图里看到的"全年合计 vs 各月"。
    这里在数据预处理阶段把这些行剔除。

    判定策略(更严格一些,避免误杀真实业务数据):
        - 转 str 后 strip + lower
        - 含上述关键词(精确包含,非前缀);同时排除「纯数字 / 空值」
        - 全角括号、空白不影响匹配
    """
    if value is None:
        return False
    s = str(value).strip().lower()
    if not s:
        return False
    # 纯数字 / 日期不算汇总(如 "12.0" 不该被误判为 "总计")
    try:
        float(s.replace(",", ""))
        return False
    except (ValueError, TypeError):
        pass
    # 含任一关键词即视为汇总行
    return any(kw in s for kw in _SUMMARY_KEYWORDS)


def _strip_summary_rows(df: pd.DataFrame, cat_cols: list[str]) -> pd.DataFrame:
    """在不影响 cat_cols 之外结构的前提下,剔除 cat_cols 上的汇总行。

    实现要点:
        - 只剔除 cat_cols(分类列)中**任何一个**为汇总行的行(交集为空会误杀)
        - 其它数值列 / 日期列不参与判断(它们的"汇总"语义不一样)
        - 找不到 cat_cols / df 为空时原样返回
        - 剔除后行数为 0 → 返回原 df,避免下游连兜底图都画不出来
    """
    if df is None or len(df) == 0:
        return df
    valid_cat_cols = [c for c in (cat_cols or []) if c in df.columns]
    if not valid_cat_cols:
        return df
    try:
        # 任一分类列命中汇总关键词就剔除
        mask = pd.Series(False, index=df.index)
        for c in valid_cat_cols:
            mask = mask | df[c].apply(_is_summary_label)
        cleaned = df.loc[~mask]
        # 兜底:全被剔除 → 原样返回
        if len(cleaned) == 0:
            return df
        return cleaned.reset_index(drop=True)
    except Exception:
        return df


def _is_chart_meaningful(opt: dict) -> bool:
    """校验一张 ECharts option 是否「值得渲染」。

    触发场景(2026-10 用户反馈):
        - 「年度目标 vs 年度实际 散点」:数据只有 1 行 / 全相等时画出来是空白;
        - 「按分类 X 饼图」:所有切片值相同(全 0 / 全相等)→ 饼图空;
        - 「Top N 柱图」:N=1 时柱形单点不可读。
    这里做最基本的有效性检查,问题图直接丢弃,不再污染看板。

    检查项:
        - series 必须存在
        - 至少一个系列至少 2 个数据点(避免单点)
        - 数值不能全相等(全相等看不出差异)
        - 数值不能全为 0/NaN
    """
    if not isinstance(opt, dict):
        return False
    series_list = opt.get("series")
    if not isinstance(series_list, list) or not series_list:
        return False
    chart_type = str(series_list[0].get("type") or "bar").lower()
    # 累计至少有效(非零非 NaN)的数值数
    total_valid = 0
    all_vals: list[float] = []
    for s in series_list:
        data = s.get("data")
        if not isinstance(data, list):
            continue
        for item in data:
            v = None
            if isinstance(item, dict):
                v = item.get("value")
                if isinstance(v, list):  # 雷达 [v1, v2, ...]
                    for x in v:
                        if isinstance(x, (int, float)) and x == x:  # NaN 检查
                            all_vals.append(float(x))
                            total_valid += 1
                    continue
            elif isinstance(item, (int, float)):
                v = item
            elif isinstance(item, list) and len(item) >= 2:
                # 散点 [x, y] / heatmap [x, y, v]
                v = item[-1]
            if isinstance(v, (int, float)) and v == v:  # NaN check
                all_vals.append(float(v))
                total_valid += 1
    # 1) 没有任何有效数值 → 空图
    if total_valid == 0:
        return False
    # 2) 散点图至少要 2 个点(单点散点没意义)
    if chart_type == "scatter" and total_valid < 2:
        return False
    # 3) 数值全相等(包括全 0)→ 看不出差异
    if all_vals and min(all_vals) == max(all_vals):
        return False
    # 4) 至少 2 个有效数值点(饼图只 1 切片 → 无意义)
    if total_valid < 2:
        return False
    # 5) 【2026-10 新增】饼图/环形图切片数 > 8 时拒绝。
    #    背景:用户截图场景 11 行日级数据 → 11 切片环形饼图,小到看不清标签。
    #    饼图本身没有"TopN 缩到 8"的概念,所以在有效性层硬拦截,让上游
    #    走"按日聚合"路径(把 11 个日 → 1 个月 → 兜底)或干脆不画饼。
    #    横 Top N / 折线不受此限(横 Top N 默认只取 top_n)。
    if chart_type == "pie" and total_valid > 8:
        return False
    return True


def _profile_dataframe(df: pd.DataFrame) -> dict[str, Any]:
    """给一份 df 做"图表可行性画像",用于 AI 决定生成几张图、用哪些图型。

    输出字段:
        row_count           : 有效行数(剔除汇总行后)
        col_count           : 总列数
        numeric_cols        : 数值列数
        date_cols           : 日期列数
        cat_cols            : 分类列数
        unique_in_primary   : 主分类的唯一值数
        has_time_series     : 是否有日期列(时间序列场景)
        has_two_dimensions  : 是否有 ≥2 个分类列(可做堆叠 / 交叉分析)
        has_two_numerics    : 是否有 ≥2 个数值列(可做散点 / 多系列)
        is_sparse           : 是否数据稀薄(行<5 或数值列<1)
        recommended_count   : 建议生成的图表数量(自适应,3~6 张)
    """
    if df is None or len(df) == 0:
        return {
            "row_count": 0, "col_count": 0,
            "numeric_cols": [], "date_cols": [], "cat_cols": [],
            "unique_in_primary": 0,
            "has_time_series": False, "has_two_dimensions": False,
            "has_two_numerics": False, "is_sparse": True,
            "recommended_count": 1,
        }
    numeric_cols, date_cols, cat_cols, _ = _split_columns(df)
    primary_cat = date_cols[0] if date_cols else (cat_cols[0] if cat_cols else None)
    if primary_cat and primary_cat in df.columns:
        try:
            unique = int(df[primary_cat].dropna().nunique())
        except Exception:
            unique = 0
    else:
        unique = 0
    has_time_series = bool(date_cols)
    has_two_dimensions = len(cat_cols) >= 2
    has_two_numerics = len(numeric_cols) >= 2
    is_sparse = len(df) < 5 or not numeric_cols
    # 自适应图表数量(取代旧的"固定 6~8 张"):
    #   - 数据稀薄(行<5 或无数值):1 张分布图
    #   - 单维度无时间:3 张(横 Top / 饼 / 竖柱)
    #   - 含时间维度 + 单一数值:4 张(横 Top / 饼 / 趋势 / 占比)
    #   - 双维度 + 双数值:6 张(全部模板)
    if is_sparse:
        rec = 1
    elif has_time_series and has_two_dimensions and has_two_numerics:
        rec = 6
    elif has_time_series and has_two_numerics:
        rec = 5
    elif has_time_series or has_two_dimensions:
        rec = 4
    else:
        rec = 3
    return {
        "row_count": int(len(df)),
        "col_count": int(len(df.columns)),
        "numeric_cols": numeric_cols,
        "date_cols": date_cols,
        "cat_cols": cat_cols,
        "unique_in_primary": unique,
        "has_time_series": has_time_series,
        "has_two_dimensions": has_two_dimensions,
        "has_two_numerics": has_two_numerics,
        "is_sparse": is_sparse,
        "recommended_count": int(rec),
    }


# ============================================================
# 【chart-render-smart-2026-10 新增】日期粒度自适应
# ------------------------------------------------------------
# 背景(用户截图 2026-10-20):
#   用户上传的 Excel 「日期」列实际是精确到秒的时间戳
#   (如 2026-10-15 00:00:00、2026-10-16 00:00:00…),共 11 行。
#   旧实现把这一列当成普通分类列 → Top 10 横轴变成 11 个时间戳;
#   环形图变成 11 个等分小切片;趋势图虽画出来但粒度过细无意义。
# 这里:
#   1) 检测日期列的真实粒度(秒/分/时/日/月/季/年/跨年)
#   2) 若唯一值数 > 12 且粒度 < 日 → 自动向上聚合(日→月、月→季、季→年)
#      让唯一值落到 ≤12 个可读分类
#   3) 返回一个新 df(分类列被替换为聚合后的"友好标签"),并把该列从
#      "日期列"语义降级为"分类列",避免后续模板再走"按日分桶"路径
# ============================================================

def _detect_date_granularity(series: pd.Series) -> str:
    """识别一个日期序列的真实粒度,返回 "second"/"minute"/"hour"/"day"/"month"/"quarter"/"year"。

    判定依据:把所有非空样本转 datetime 后,看连续两个样本的差值众数落在哪个区间。
    这样比"看字符串里有没有时分秒"更稳——既能识别 datetime64,也能识别
    "2026-10-15" 这种纯日期字符串。
    """
    try:
        dt = pd.to_datetime(series.dropna(), errors="coerce").dropna()
    except Exception:
        return "day"
    if len(dt) < 2:
        return "day"
    try:
        diffs = dt.sort_values().diff().dropna()
        if len(diffs) == 0:
            return "day"
        median = diffs.median()
    except Exception:
        return "day"
    sec = median.total_seconds()
    if sec < 60:
        return "second"
    if sec < 3600:
        return "minute"
    if sec < 86400:
        return "hour"
    if sec < 86400 * 32:  # 月内任意天数都算"日级"(含 1 天)
        return "day"
    if sec < 86400 * 95:  # ~3 个月
        return "month"
    if sec < 86400 * 380:  # ~12 个月
        return "quarter"
    return "year"


def _aggregate_date_column(
    df: pd.DataFrame, date_col: str, granularity: str
) -> tuple[pd.DataFrame, str]:
    """按指定粒度把日期列聚合到「友好标签」列,返回 (新 df, 新分类列名)。

    设计要点:
        - 聚合粒度 = `granularity` 向上取整(秒→分→时→日→月→季→年)
        - 聚合方式:按新分类列 groupby 数值列 sum
        - 如果该日期列同时含时分秒,聚合后会丢失时分秒信息 —— 标题里
          会改用「按月」「按日」等粒度名,而不是原列名
        - 失败 / 全空 → 原样返回 (df, 原列名),避免误伤
        - **强制聚合条件**:若当前粒度 ≥ day 但原始 uniq > 12 → 必须向上爬;
          若 < day(uniqu 已经是"按秒") → 强制向上到 day(秒级时间戳几乎一定
          是数据导入时未格式化造成的,直接升到"按日"才有意义)
    """
    if not date_col or date_col not in df.columns or granularity not in (
        "second", "minute", "hour", "day", "month", "quarter", "year"
    ):
        return df, date_col
    try:
        dt = pd.to_datetime(df[date_col], errors="coerce")
        if dt.isna().all():
            return df, date_col
        ladder = ["second", "minute", "hour", "day", "month", "quarter", "year"]
        idx = ladder.index(granularity)
        uniq = dt.dropna().nunique()

        # 动态爬升:从当前粒度开始,沿阶梯向上,找到"唯一值数 ∈ [2, 8]"
        # 的目标粒度。例如 11 行日级数据 uniq=11,需要升到月,
        # 但若整批数据全在同一月(如 uni=11 个不同日期都在 10 月),升到月
        # 后只剩 1 个值,需要继续升到季/年;
        # 但又不能升到年(此时 1 年 = 1 个值,会被 _is_chart_meaningful 拒)。
        # 所以目标粒度的 uniq 必须 >= 2 且 <= 8;若爬到 year 仍 < 2,
        # 就保留原粒度让 _is_chart_meaningful 自然拒绝(并出兜底图)。
        target = None
        for hop in ladder[idx:]:
            try:
                if hop == "second":
                    probe = dt.dt.strftime("%Y-%m-%d %H:%M:%S")
                elif hop == "minute":
                    probe = dt.dt.strftime("%Y-%m-%d %H:%M")
                elif hop == "hour":
                    probe = dt.dt.strftime("%Y-%m-%d %H:00")
                elif hop == "day":
                    probe = dt.dt.strftime("%Y-%m-%d")
                elif hop == "month":
                    probe = dt.dt.strftime("%Y-%m")
                elif hop == "quarter":
                    probe = dt.dt.year.astype(str) + "-Q" + dt.dt.quarter.astype(str)
                else:
                    probe = dt.dt.year.astype(str)
                probe_uniq = probe.dropna().nunique()
            except Exception:
                break
            if 2 <= probe_uniq <= 8:
                target = hop
                break
            if probe_uniq < 2:
                # 再向上爬只会更少,放弃聚合
                break
        if target is None:
            return df, date_col
        # 实际生成标签
        try:
            if target == "day":
                labels = dt.dt.strftime("%Y-%m-%d")
            elif target == "month":
                labels = dt.dt.strftime("%Y-%m")
            elif target == "quarter":
                labels = dt.dt.year.astype(str) + "-Q" + dt.dt.quarter.astype(str)
            else:
                labels = dt.dt.year.astype(str)
        except Exception:
            return df, date_col
        # 把新标签当成普通分类列使用:加 __date_agg__ 列,按它 groupby 数值列求和
        new_df = df.copy()
        new_df["__date_agg__"] = labels
        numeric_in_df = [c for c in df.columns
                         if c != date_col and pd_is_numeric_dtype(new_df[c])]
        if numeric_in_df:
            agg = new_df.groupby("__date_agg__", as_index=False)[numeric_in_df].sum()
            # 保留非数值列的首条(避免丢失维度信息)
            non_num = [c for c in df.columns if c not in numeric_in_df and c != date_col]
            if non_num:
                sample = new_df.groupby("__date_agg__", as_index=False)[non_num].first()
                agg = sample.merge(agg, on="__date_agg__", how="outer")
            agg = agg.reset_index(drop=True)
            return agg, "__date_agg__"
        # 没数值列 → 直接替换标签
        new_df = new_df.drop(columns=[date_col]).rename(columns={"__date_agg__": date_col})
        return new_df.reset_index(drop=True), date_col
    except Exception:
        return df, date_col


# 粒度名 → 中文显示(用于标题)
_DATE_GRANULARITY_LABEL_CN = {
    "second": "按秒", "minute": "按分", "hour": "按小时",
    "day": "按日", "month": "按月", "quarter": "按季", "year": "按年",
}


# ============================================================
# 【chart-render-smart-2026-10 新增】混合出图:本地骨架 + LLM 智能筛选
# ------------------------------------------------------------
# 背景(用户 2026-10-20 反馈):
#   用户希望「混合」出图——既保留本地规则出图的"稳",又让 AI 决定保留哪些图、
#   给出更自然的标题。现阶段硬性痛点是「图表过多且很多是空的」,先解决这个。
#
# 设计:
#   1) 第一道:本地 build_dashboard_options 出图骨架(已有)
#   2) 第二道:chart_ai_curate(sheets, chart_data, style_key)
#       - 输入:本地出的全部图表 + 每张图的"信息密度"(切片数 / 数值跨度 / 趋势单调性)
#       - 输出:1~3 张最有信息量的图,标题改成更自然的中文(由 LLM 重写)
#       - LLM 不可用 → 直接按本地骨架的"信息密度评分"排序取 Top N,完全无 LLM 依赖
#   3) 调用方:workbench.preview_charts 加一个 ?ai_curate=1 query 或 payload 字段,
#       开关打开时走混合路径,关闭时维持本地路径(向后兼容)。
# ============================================================

# 信息密度评分阈值:低于此分的图视为"无信息量",优先被剔除
_INFORMATIVENESS_MIN_SCORE = 0.35


def _chart_informativeness(opt: dict) -> float:
    """给一张 ECharts option 打"信息密度"分,0~1 之间。

    评分维度:
      - 切片/数据点数:太少(<3)或太多(>20)扣分
      - 数值方差:全相等扣 0 分;变异系数越大分越高
      - 趋势单调性:折线图完全单调(线性)略加分,完全无序略减分
      - 饼图切片过多:用户痛点,直接扣到 0.3 以下

    返回:
        0~1 的浮点分。>= 0.35 视为"值得保留"。
    """
    try:
        series_list = opt.get("series") or []
        if not series_list:
            return 0.0
        chart_type = str(series_list[0].get("type") or "bar").lower()
        vals: list[float] = []
        for s in series_list:
            data = s.get("data") or []
            for item in data:
                v = None
                if isinstance(item, dict):
                    v = item.get("value")
                    if isinstance(v, list):
                        vals.extend(float(x) for x in v if isinstance(x, (int, float)))
                        continue
                elif isinstance(item, (int, float)):
                    v = item
                elif isinstance(item, list) and len(item) >= 2:
                    v = item[-1]
                if isinstance(v, (int, float)):
                    vals.append(float(v))
        if not vals or min(vals) == max(vals):
            return 0.0
        n = len(vals)
        # 1) 切片数打分:3~8 最理想
        if 3 <= n <= 8:
            size_score = 1.0
        elif n < 3:
            size_score = 0.3
        elif n <= 15:
            size_score = 0.6
        else:
            size_score = 0.3
        # 2) 饼图切片过多(用户截图痛点)直接扣
        if chart_type == "pie" and n > 8:
            return 0.2
        # 3) 数值变异系数:cv = std / mean
        import statistics
        mean = statistics.mean(vals)
        std = statistics.pstdev(vals) if len(vals) > 1 else 0.0
        cv = (std / abs(mean)) if abs(mean) > 1e-9 else 0.0
        # cv 越大分越高;经验:cv=0.3 → 0.7, cv=1 → 1.0, cv≥2 → 1.0
        var_score = max(0.0, min(1.0, cv * 0.7 + 0.3))
        return round(size_score * 0.5 + var_score * 0.5, 3)
    except Exception:
        return 0.5  # 解析失败时给个中间分,避免被一刀切


def ai_curate_charts(
    charts: list[dict],
    df_summary: str | None = None,
    target_count: int = 3,
    style_key: str | None = None,
) -> list[dict]:
    """从本地骨架中精选 1~target_count 张图,并用 LLM 润色标题。

    实现策略(双层保险):
      L1) 本地信息密度评分:无 LLM 也能跑,失败兜底即生效
      L2) LLM 重排 + 润色:在 L1 排序后,让 LLM 重写每张图的标题,让文案更自然

    Args:
        charts: 本地生成的图表列表(每项含 title/option/style_key)
        df_summary: 数据集一句话描述(可选,辅助 LLM 理解业务背景)
        target_count: 目标保留张数(默认 3)
        style_key: 风格 key,透传给 option 的 style_key

    Returns:
        精选后的图表列表(可能少于 target_count,也可能为 0 表示本地骨架全部不合格)
    """
    if not charts:
        return []

    # ---- L1:本地打分 + 排序 ----
    scored = [
        (c, _chart_informativeness(c.get("option") or {}))
        for c in charts
        if isinstance(c, dict)
    ]
    scored.sort(key=lambda x: x[1], reverse=True)
    candidates = [c for c, s in scored if s >= _INFORMATIVENESS_MIN_SCORE]

    # 兜底:全部低于阈值 → 仍取分数最高的 1 张,避免 0 张
    if not candidates and scored:
        candidates = [scored[0][0]]
    candidates = candidates[: max(1, int(target_count))]
    if not candidates:
        return []

    # ---- L2:LLM 润色标题(失败 → 保留原标题)----
    try:
        from app.services.llm_client import llm  # 局部 import,避免启动期循环
    except Exception:
        llm = None

    if llm is None:
        return candidates

    # 构造精简的图表摘要给 LLM(避免 prompt 过大)
    brief_lines = []
    for i, c in enumerate(candidates, 1):
        opt = c.get("option") or {}
        series = opt.get("series") or []
        ctype = series[0].get("type") if series else "bar"
        # 取前几个 data 点作为摘要
        sample = []
        for s in series[:1]:
            d = s.get("data") or []
            for item in d[:5]:
                if isinstance(item, dict):
                    sample.append(f"{item.get('name')}={item.get('value')}")
                elif isinstance(item, (int, float)):
                    sample.append(str(item))
                elif isinstance(item, list) and len(item) >= 2:
                    sample.append(f"({item[0]},{item[1]})")
        brief_lines.append(
            f"[{i}] 原标题:{c.get('title','')} | 类型:{ctype} | 前5样本:{'、'.join(sample)[:80]}"
        )

    system_prompt = (
        "你是一名严谨的数据分析师,根据真实数据特征为图表起一个更自然、信息量更大的标题。"
        "**严禁编造任何数据**——你只能基于下面给的真实样本润色措辞/指出重点维度,"
        "不要发明样本中没有的数字/分类。"
    )
    user_prompt = (
        f"数据集概览:{(df_summary or '未提供')[:200]}\n\n"
        f"下列是 {len(candidates)} 张本地规则生成的图表及其真实数据样本:\n"
        + "\n".join(brief_lines)
        + "\n\n请严格按 JSON 输出:\n"
        '{"titles": [{"index": 1, "title": "..."}, {"index": 2, "title": "..."}, ...]}\n'
        "要求:\n"
        "  - title 用中文,不超过 18 字,口语化、信息量大(如「日均订单量排行 Top 10」"
        "「销售额月度对比」「新增 vs 流失客户分布」);\n"
        "  - 不要直接复读原标题;若数据样本含具体数值,可以在标题里点出极值方向"
        "(如「10月订单量最高」),但只能用样本里出现过的数字;\n"
        "  - 不要编造样本里没有的字段/分类;\n"
        "  - index 必须严格对应上面的 [N] 编号。"
    )

    try:
        result = llm.chat_json(system_prompt, user_prompt, temperature=0.3)
        new_titles = {}
        if isinstance(result, dict) and isinstance(result.get("titles"), list):
            for item in result["titles"]:
                if not isinstance(item, dict):
                    continue
                idx = item.get("index")
                title = item.get("title")
                if isinstance(idx, int) and isinstance(title, str) and title.strip():
                    new_titles[idx] = title.strip()
        # 应用新标题(只在 LLM 返回了有效标题时才覆盖,否则保留原标题)
        for i, c in enumerate(candidates, 1):
            if i in new_titles:
                c["title"] = new_titles[i]
                # option.title.text 也同步更新,前端 ECharts 才看得到
                try:
                    if isinstance(c.get("option"), dict):
                        if isinstance(c["option"].get("title"), dict):
                            c["option"]["title"]["text"] = new_titles[i]
                        else:
                            c["option"]["title"] = new_titles[i]
                except Exception:
                    pass
    except Exception:
        # LLM 失败 → 保留本地标题,不抛错
        pass
    return candidates


# ============================================================
# 中文字体注册：兼容 Linux / macOS / Windows
# ============================================================
def _register_chinese_font() -> str:
    """自动找一个可用的中文字体,返回字体名;找不到则用 DejaVu(豆腐字兜底)。

    ECharts 前端一般用 PingFang/宋体/微软雅黑,matplotlib 没原生支持。
    本函数按常见系统字体名列表顺序探测;找到就用,没找到 fallback 到
    默认 sans-serif(中文方块但不崩)。
    """
    candidates = [
        "Heiti SC", "PingFang SC", "Songti SC", "STHeiti", "STSong",
        "Hiragino Sans GB", "Microsoft YaHei", "SimHei", "SimSun",
        "Noto Sans CJK SC", "Noto Sans CJK", "Source Han Sans SC",
        "WenQuanYi Zen Hei", "WenQuanYi Micro Hei",
        "Arial Unicode MS",
    ]
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in installed:
            plt.rcParams["font.sans-serif"] = [name] + plt.rcParams.get(
                "font.sans-serif", []
            )
            plt.rcParams["axes.unicode_minus"] = False
            return name
    plt.rcParams["axes.unicode_minus"] = False
    return "DejaVu Sans"


_CN_FONT_NAME = _register_chinese_font()
logger.info("ECharts→matplotlib 使用字体: %s", _CN_FONT_NAME)


# ============================================================
# 安全色板（与前端 ReportView 的 chart 主题保持一致）
# ============================================================
_PALETTE = [
    "#22d3ee", "#5eead4", "#a78bfa", "#fbbf24", "#f87171", "#4ade80",
    "#60a5fa", "#f472b6", "#facc15", "#34d399", "#818cf8", "#fb7185",
]


def _normalize_color(color: Any) -> str | None:
    """把任意颜色值规整为 matplotlib / ECharts 都能识别的字符串。

    非法颜色(非字符串 / ECharts 渐变对象 / 空串)一律返回 None,
    由调用方回退默认色板,避免渲染期抛异常导致整张图直接渲染失败。
    """
    if not isinstance(color, str):
        return None
    c = color.strip()
    if not c:
        return None
    try:
        if is_color_like(c):
            return c
    except Exception:
        return None
    return None


def _extract_series_colors(chart: Any) -> list[str]:
    """从 chart 里解析出 option.color(风格色盘),返回合法颜色列表。

    解析规则与 _render_chart_to_png_inner 保持完全一致:
        option = chart["option"] if isinstance(chart["option"], dict) else chart
    之所以要"再解析一次",是因为 render_chart_to_png 需要在进入渲染前
    就把色盘注入 contextvar,好让 _colors() 在任何 _render_* 里都能取到。
    """
    if not isinstance(chart, dict):
        return []
    option = chart.get("option")
    if not isinstance(option, dict):
        option = chart
    raw = option.get("color")
    if not isinstance(raw, (list, tuple)):
        return []
    colors: list[str] = []
    for item in raw:
        norm = _normalize_color(item)
        if norm:
            colors.append(norm)
    return colors


def _colors(n: int) -> list[str]:
    """返回 n 个颜色,不足就循环。

    【2026-10 chart-visual-upgrade 升级】
    优先使用当前渲染上下文注入的色盘(option.color,即风格色盘);
    未注入时回退到模块级 _PALETTE。
    这样 tech_advanced / business_pro / minimalist 三套风格在导出的 PNG 里
    才会真正生效 —— 此前 3 种风格导出来完全一样,用户反馈"图表太普通"。
    """
    if n <= 0:
        return []
    injected = _CHART_SERIES_COLORS.get()
    if injected:
        return [injected[i % len(injected)] for i in range(n)]
    return [_PALETTE[i % len(_PALETTE)] for i in range(n)]


# ============================================================
# 标题/坐标文字清洗（图表文字也可能含 <think> / 英文短词）
# ============================================================
def _clean_label(s: Any, max_len: int = 64) -> str:
    """标题/坐标轴文字清洗 + 长度截断。"""
    from app.core.json_utils import clean_llm_text  # 局部导入,避免启动循环

    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = clean_llm_text(s).strip()
    if not s:
        return ""
    s = s.replace("\n", " ").replace("\r", " ")
    if len(s) > max_len:
        s = s[:max_len] + "…"
    return s


# ============================================================
# 数据归一化 helpers
# ============================================================
def _series_data_to_pairs(series: dict) -> tuple[list[str], list[float]]:
    """把 series.data 归一为 (names[], values[])。"""
    data = series.get("data") or []
    names: list[str] = []
    values: list[float] = []
    if not isinstance(data, list):
        data = [data]
    for i, item in enumerate(data):
        if isinstance(item, dict):
            name = (item.get("name") or item.get("label")
                    or item.get("x") or f"系列{i + 1}")
            val = item.get("value")
            if val is None and isinstance(item.get("value"), list):
                # 雷达散点：{"value": [v1, v2, ...]}
                v_arr = item["value"]
                val = v_arr[0] if v_arr else 0
            try:
                val = float(val if val is not None else 0)
            except (TypeError, ValueError):
                val = 0.0
            names.append(_clean_label(name, 32))
            values.append(val)
        else:
            try:
                values.append(float(item))
            except (TypeError, ValueError):
                values.append(0.0)
            names.append(f"系列{len(names) + 1}")
    return names, values


def _x_axis_categories(series: dict, names: list[str], values: list[float]) -> list[str]:
    """对 bar/line,从 xAxis 取分类轴;没有就用默认 1..N。"""
    if names and any(n and not n.startswith("系列") for n in names):
        return names
    return [str(i + 1) for i in range(len(values))]


# ============================================================
# 【chart-visual-upgrade 2026-10 新增】图表质感升级工具
# ------------------------------------------------------------
# 背景:
#   用户反馈"我想要的是可视化报表,现在生成的图表太普通了"。
#   代码级根因:matplotlib 兜底渲染用的是 ax.bar 默认样式 —— 纯色平涂、
#   直角柱、无数值标签、网格质感弱;折线图更是纯线条,没有渐变面积、
#   没有数据点标签;标题永远是居中的默认样式。
#   下面这一组 helper 负责把这些"质感"补齐,且**全部可降级**:
#   任何一步出错都回退到 matplotlib 原生画法,绝不因为装饰而整图渲染失败。
# ============================================================


def _format_number(value: Any) -> str:
    """把数值格式化成图表标签用的短文本(带千分位,去掉无意义的小数 0)。"""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return ""
    # NaN / inf 都不该出现在标签里
    if v != v or v in (float("inf"), float("-inf")):
        return ""
    av = abs(v)
    if av >= 1000:
        return f"{v:,.0f}"
    if av >= 100:
        return f"{v:,.1f}"
    if av >= 10:
        return f"{v:.1f}"
    if float(v).is_integer():
        return str(int(v))
    return f"{v:.2f}"


def _blend(color: str, target: str, ratio: float) -> str:
    """把 color 朝 target(通常是白/黑)方向混合 ratio(0~1),返回 #rrggbb。

    用于生成渐变的"高光端"颜色:柱图/面积图的顶部更亮,视觉上更有立体感。
    """
    try:
        c = np.array(to_rgb(color), dtype=float)
        t = np.array(to_rgb(target), dtype=float)
    except Exception:
        return color
    mixed = c + (t - c) * float(ratio)
    return "#%02x%02x%02x" % tuple(
        int(round(min(1.0, max(0.0, ch)) * 255)) for ch in mixed
    )


def _hex_to_rgba(color: str, alpha: float = 0.35) -> str:
    """十六进制颜色 → rgba(...) 字符串(ECharts 发光阴影 shadowColor 用)。"""
    try:
        r, g, b = to_rgb(color)
    except Exception:
        return f"rgba(45, 212, 191, {alpha})"
    return (
        f"rgba({int(round(r * 255))}, {int(round(g * 255))}, "
        f"{int(round(b * 255))}, {alpha})"
    )


def _linear_gradient(color: str, lighten: float = 0.32) -> dict:
    """构造 ECharts 的 LinearGradient(**纯 dict,可 JSON 序列化**)。

    后端没有 echarts.graphic 的 Python 实现(那是 JS 端 API),
    所以只能手写官方结构:
        {"type": "linear", "x": 0, "y": 0, "x2": 0, "y2": 1,
         "colorStops": [{"offset": 0, "color": "#..."},
                        {"offset": 1, "color": "#..."}]}
    前端 ECharts 5 能识别 type 为 "linear" 的颜色对象,直接当 color 用。
    """
    top = _blend(color, "#ffffff", lighten)
    return {
        "type": "linear",
        "x": 0, "y": 0, "x2": 0, "y2": 1,
        "colorStops": [
            {"offset": 0, "color": top},
            {"offset": 1, "color": color},
        ],
    }


class _RoundedBar(FancyBboxPatch):
    """圆角柱:在**真正绘制的那一刻**按当前数据变换动态修正圆角比例。

    原理(见 matplotlib 源码 FancyBboxPatch.get_path):
        框架会先把 y 方向按 mutation_aspect 压扁 → 套"圆角"模板 → 再把 y 还原。
        因此只有当 mutation_aspect == (1 个 x 数据单位对应的像素) /
        (1 个 y 数据单位对应的像素) 时,圆角在屏幕上才是**正圆**,
        否则柱子的圆角会被拉成椭圆(横向柱尤其明显)。
        而 tight_layout() 会在渲染之后再次调整坐标区大小,渲染期算出来的比例
        会过期,所以这里重写 get_path(),在 draw 时刻重新计算,任何画布尺寸都正确。
    """

    def __init__(self, xy, width, height, *, rounding_px: float = 4.0, **kwargs):
        # boxstyle 先给一个占位值,真正绘制时(get_path)会按像素目标重算
        super().__init__(xy, width, height, boxstyle="round,pad=0,rounding_size=0",
                         **kwargs)
        self._rounding_px = float(rounding_px)

    def get_path(self):
        """绘制时刻动态计算 mutation_aspect / 圆角半径。"""
        try:
            trans = self.get_transform()
            ox, oy = trans.transform((0.0, 0.0))
            px_per_x = abs(trans.transform((1.0, 0.0))[0] - ox)
            px_per_y = abs(trans.transform((0.0, 1.0))[1] - oy)
            if px_per_x > 1e-9 and px_per_y > 1e-9:
                aspect = px_per_x / px_per_y
                self.set_mutation_aspect(aspect)
                # 目标圆角(像素) → x 数据单位
                radius = self._rounding_px / px_per_x
                # 半径不能超过柱子自身尺寸的一半,否则形状会自交变成怪形
                radius = min(
                    radius,
                    abs(self.get_width()) / 2.0,
                    abs(self.get_height()) / (2.0 * aspect),
                )
                if radius > 0:
                    self.set_boxstyle(f"round,pad=0,rounding_size={radius}")
        except Exception:  # pragma: no cover - 兜底:用默认矩形画
            pass
        return super().get_path()


def _draw_bar_gradient(ax, x0: float, x1: float, y0: float, y1: float,
                       color: str, horizontal: bool = False,
                       zorder: float = 3.02):
    """在柱子区域铺一层线性渐变(imshow + LinearSegmentedColormap)。

    返回 matplotlib 图像对象(供 set_clip_path 裁剪到圆角形状);
    失败返回 None —— 调用方保留纯色圆角柱,不影响出图。
    """
    try:
        # 底部/左侧是主色,顶部/右侧是高光色,形成"打光"效果
        cmap = LinearSegmentedColormap.from_list(
            f"chart_bar_grad_{str(color).lstrip('#')}",
            [color, _blend(color, "#ffffff", 0.30)],
        )
        if horizontal:
            grad = np.linspace(0, 1, 256).reshape(1, 256)
        else:
            grad = np.linspace(0, 1, 256).reshape(256, 1)
        img = ax.imshow(
            grad, extent=(x0, x1, y0, y1), origin="lower", aspect="auto",
            cmap=cmap, interpolation="bicubic", zorder=zorder,
        )
        return img
    except Exception as exc:
        logger.debug("柱状图渐变层绘制失败(回退纯色): %s", exc)
        return None


def _draw_bars(ax, positions: list[float], values: list[float], width: float,
               color: str, is_horizontal: bool) -> None:
    """绘制一组柱子:圆角(FancyBboxPatch) + 渐变。

    稳健性:任何异常都回退到 matplotlib 原生 ax.bar / ax.barh,
    并清理掉已经画了一半的圆角柱,避免出现"半张图是圆角半张图是方柱"。
    """
    patches: list[Any] = []
    images: list[Any] = []
    try:
        for pos, val in zip(positions, values):
            v = float(val)
            if is_horizontal:
                x0, y0 = min(0.0, v), pos - width / 2.0
                w, h = abs(v), width
            else:
                x0, y0 = pos - width / 2.0, min(0.0, v)
                w, h = width, abs(v)
            if w <= 0 or h <= 0:
                continue  # 值为 0 的柱子没有可见面积,跳过(标签仍会画)
            patch = _RoundedBar(
                (x0, y0), w, h, rounding_px=4.0,
                facecolor=color, edgecolor="none", zorder=3,
            )
            ax.add_patch(patch)
            patches.append(patch)
            img = _draw_bar_gradient(
                ax, x0, x0 + w, y0, y0 + h, color, horizontal=is_horizontal
            )
            if img is not None:
                try:
                    img.set_clip_path(patch)
                    images.append(img)
                except Exception:
                    pass
    except Exception as exc:
        logger.warning("圆角柱绘制失败,回退默认柱状图: %s", exc)
        for artist in patches + images:
            try:
                artist.remove()
            except Exception:
                pass
        try:
            if is_horizontal:
                ax.barh(positions, values, height=width, color=color, zorder=3)
            else:
                ax.bar(positions, values, width=width, color=color, zorder=3)
        except Exception:
            pass


def _add_bar_labels(ax, positions: list[float], values: list[float],
                    is_horizontal: bool) -> None:
    """柱状图数据标签:竖直柱标在柱顶,横向柱标在柱右侧。"""
    pal = _palette()
    for pos, val in zip(positions, values):
        try:
            v = float(val)
        except (TypeError, ValueError):
            continue
        text = _format_number(v)
        if not text:
            continue
        try:
            if is_horizontal:
                # 横向条形图:标签放在柱子右端外侧
                ax.annotate(
                    text, xy=(v, pos), xytext=(6, 0), textcoords="offset points",
                    ha="left", va="center", fontsize=9, color=pal["label"], zorder=5,
                )
            elif v < 0:
                ax.annotate(
                    text, xy=(pos, v), xytext=(0, -5), textcoords="offset points",
                    ha="center", va="top", fontsize=9, color=pal["label"], zorder=5,
                )
            else:
                ax.annotate(
                    text, xy=(pos, v), xytext=(0, 4), textcoords="offset points",
                    ha="center", va="bottom", fontsize=9, color=pal["label"], zorder=5,
                )
        except Exception:
            continue


def _fill_line_area(ax, xs: list[int], values: list[Any], color: str) -> None:
    """折线下方的渐变面积。

    做法:fill_between 传入**逐顶点 alpha 数组**(上沿浓、下沿淡),
    形成天然的"面积渐隐"效果 —— 比纯色半透明填充更有质感。
    不支持数组 alpha 的旧版本会自动退化为固定 alpha。
    """
    try:
        ys = [np.nan if v is None else float(v) for v in values]
        alpha = np.linspace(0.32, 0.02, max(2, len(xs)))
        ax.fill_between(xs, ys, 0, color=color, alpha=alpha, zorder=2)
    except Exception as exc:
        logger.debug("折线面积渐变失败(回退纯色面积): %s", exc)
        try:
            ys = [np.nan if v is None else float(v) for v in values]
            ax.fill_between(xs, ys, 0, color=color, alpha=0.15, zorder=2)
        except Exception:
            pass


def _add_line_labels(ax, xs: list[int], values: list[Any]) -> None:
    """折线数据点标签。

    避让策略:点少(<=12)时全部标注;点多时只标注首点、末点、最大值、最小值,
    并把字号调小 —— 否则标签会互相压叠成一团,反而更难读。
    """
    pal = _palette()
    points = [
        (x, float(v)) for x, v in zip(xs, values)
        if v is not None and v == v
    ]
    if not points:
        return
    total = len(points)
    if total <= 12:
        targets, fontsize = points, 9
    else:
        vals = [p[1] for p in points]
        keep = {0, total - 1, vals.index(max(vals)), vals.index(min(vals))}
        targets = [p for i, p in enumerate(points) if i in keep]
        fontsize = 8
    for x, v in targets:
        text = _format_number(v)
        if not text:
            continue
        try:
            ax.annotate(
                text, xy=(x, v), xytext=(0, 7), textcoords="offset points",
                ha="center", va="bottom", fontsize=fontsize,
                color=pal["label"], zorder=5,
            )
        except Exception:
            continue


def _decorate_title(ax, title: str, accent_color: str,
                    is_polar: bool = False) -> None:
    """标题装饰:左对齐标题 + 标题下方的 accent 短装饰线。

    为什么:居中标题 + 无装饰是老式"Excel 截图"观感;
    左对齐标题 + 一小段主色装饰线是现代数据看板的通用语言,
    能显著提升"这看起来像一份正经可视化报表"的观感。
    """
    pal = _palette()
    text = _clean_label(title, 80)
    try:
        ax.set_title(text or "", fontsize=14, loc="left", pad=18, color=pal["title"])
    except Exception:
        try:
            ax.set_title(text or "", fontsize=14, pad=14)
        except Exception:
            pass
    if is_polar or not text or not accent_color:
        return
    try:
        # 装饰线画在坐标区坐标系里(1.012 = 略高于坐标区上沿),
        # clip_on=False 保证不会被裁掉;长度只有坐标区宽度的 5.5%,够点缀不抢戏。
        line = Line2D(
            [0.0, 0.055], [1.012, 1.012], transform=ax.transAxes,
            color=accent_color, linewidth=2.6, solid_capstyle="round",
            clip_on=False, zorder=6,
        )
        ax.add_line(line)
    except Exception:
        pass


def _hide_spines(ax) -> None:
    """隐藏上/右边框,只留左/下,让图表更干净(既有行为,保留)。"""
    for key in ("top", "right"):
        try:
            ax.spines[key].set_visible(False)
        except Exception:
            pass


def _style_grid(ax, axis: str = "both") -> None:
    """网格质感:虚线 + 低透明度 + 置于数据层之下。

    实线网格很"硬",会把图表切成格子;虚线 + 半透明能让网格退到背景层,
    既保留读数参考,又不喧宾夺主。
    """
    pal = _palette()
    try:
        ax.set_axisbelow(True)
        ax.grid(True, axis=axis, linestyle="--", linewidth=0.8,
                alpha=0.7, color=pal["grid"])
    except Exception:
        pass


def _style_legend(ax, handles: list[Any]) -> None:
    """绘制图例(代理句柄),文字颜色跟随主题 —— light 主题不会写白字。"""
    pal = _palette()
    try:
        legend = ax.legend(handles=handles, loc="best", fontsize=10, frameon=False)
        for text in legend.get_texts():
            text.set_color(pal["label"])
    except Exception:
        pass


# ============================================================
# 各种图型渲染
# ============================================================
def _render_pie(ax, series_list: list[dict], title: str) -> None:
    """渲染饼图/环形图(根据 series.radius 字段判断)。"""
    pal = _palette()
    merged_names: list[str] = []
    merged_vals: list[float] = []
    for s in series_list:
        n, v = _series_data_to_pairs(s)
        merged_names.extend(n)
        merged_vals.extend(v)
    if not merged_vals:
        ax.text(0.5, 0.5, "(无数据)", ha="center", va="center", fontsize=14,
                color=pal["tick"])
        return
    is_donut = any(isinstance(s.get("radius"), list) for s in series_list)
    colors = _colors(len(merged_vals))
    if is_donut:
        wedges, texts, autotexts = ax.pie(
            merged_vals, labels=merged_names, colors=colors,
            wedgeprops=dict(width=0.42, edgecolor=pal["bg"], linewidth=1.5),
            autopct="%1.1f%%", startangle=90, pctdistance=0.78,
            textprops={"fontsize": 10, "color": pal["label"]},
        )
        for t in autotexts:
            t.set_color("#ffffff"); t.set_fontsize(10)
        # 【B3.5】环形图中心显示总计:环形图的"圆心"天然适合放汇总数字,
        # 一眼就能看到总量,这是环形图相对普通饼图的核心信息优势。
        total = sum(merged_vals)
        total_text = _format_number(total)
        if total_text:
            ax.text(
                0, 0, f"总计\n{total_text}", ha="center", va="center",
                fontsize=13, fontweight="bold", color=pal["title"],
                linespacing=1.5, zorder=6,
            )
    else:
        ax.pie(merged_vals, labels=merged_names, colors=colors,
               autopct="%1.1f%%", startangle=90,
               textprops={"fontsize": 10, "color": pal["label"]},
               wedgeprops=dict(edgecolor=pal["bg"], linewidth=1.5))
    _decorate_title(ax, title, colors[0] if colors else pal["title"])


def _render_bar(ax, series_list: list[dict], x_categories: list[str], title: str,
                is_horizontal: bool = False) -> None:
    """柱状图(竖直/水平)。

    【chart-visual-upgrade】圆角柱 + 竖向渐变 + 柱顶数值标签 + 虚线网格。
    """
    if not series_list:
        return
    n_series = len(series_list)
    width = 0.8 / max(1, n_series)
    colors = _colors(n_series)
    pal = _palette()
    handles: list[Any] = []
    max_val = 0.0
    min_val = 0.0
    for i, s in enumerate(series_list):
        _, values = _series_data_to_pairs(s)
        if not values:
            values = [0] * len(x_categories)
        if len(values) < len(x_categories):
            values = values + [0] * (len(x_categories) - len(values))
        # 数据比分类多时截断:否则柱子会画到坐标区外,还会让刻度对不上
        values = values[: len(x_categories)]
        offset = (i - (n_series - 1) / 2) * width
        positions = [(j + offset) for j in range(len(x_categories))]
        label = _clean_label(s.get("name"), 32) or f"系列{i + 1}"
        _draw_bars(ax, positions, values, width, colors[i], is_horizontal)
        _add_bar_labels(ax, positions, values, is_horizontal)
        # 图例用代理色块(因为柱子是手动画的 patch,不会被自动收集)
        handles.append(Patch(facecolor=colors[i], edgecolor="none", label=label))
        if values:
            max_val = max(max_val, max(values))
            min_val = min(min_val, min(values))
    # 刻度
    if is_horizontal:
        ax.set_yticks(range(len(x_categories)))
        ax.set_yticklabels(x_categories, fontsize=10)
    else:
        ax.set_xticks(range(len(x_categories)))
        ax.set_xticklabels(
            x_categories, fontsize=10,
            rotation=30 if any(len(c) > 6 for c in x_categories) else 0,
        )
    # 数值轴范围:必须显式设置 —— 渐变层用的 imshow 会把坐标范围改成图片范围
    span = (max_val - min_val) or max(abs(max_val), 1.0)
    try:
        if is_horizontal:
            ax.set_xlim(min(0.0, min_val) - span * 0.02,
                        max_val + span * 0.18)
            ax.set_ylim(-0.6, max(0, len(x_categories) - 1) + 0.6)
        else:
            ax.set_ylim(min(0.0, min_val) - span * 0.02,
                        max_val + span * 0.18)
            ax.set_xlim(-0.6, max(0, len(x_categories) - 1) + 0.6)
    except Exception:
        pass
    _decorate_title(ax, title, colors[0] if colors else pal["title"])
    if n_series > 1:
        _style_legend(ax, handles)
    _hide_spines(ax)
    _style_grid(ax, axis=("x" if is_horizontal else "y"))


def _render_line(ax, series_list: list[dict], x_categories: list[str], title: str) -> None:
    """折线图。

    【chart-visual-upgrade】渐变面积 + 数据点标签 + 圆点描边 + 虚线网格。
    """
    if not series_list:
        return
    pal = _palette()
    colors = _colors(len(series_list))
    xs = list(range(len(x_categories)))
    handles: list[Any] = []
    for i, s in enumerate(series_list):
        _, values = _series_data_to_pairs(s)
        if len(values) < len(xs):
            values = values + [None] * (len(xs) - len(values))
        values = values[: len(xs)]  # 超长截断,避免 x/y 长度不等直接报错
        label = _clean_label(s.get("name"), 32) or f"系列{i + 1}"
        ax.plot(xs, values, marker="o", color=colors[i], label=label,
                linewidth=2.2, markersize=6, markeredgecolor=pal["bg"],
                markeredgewidth=1.2, zorder=4, solid_capstyle="round")
        _fill_line_area(ax, xs, values, colors[i])
        _add_line_labels(ax, xs, values)
        handles.append(Line2D([0], [0], color=colors[i], marker="o",
                              markersize=6, linewidth=2.2, label=label))
    ax.set_xticks(xs)
    ax.set_xticklabels(
        x_categories, fontsize=10,
        rotation=30 if any(len(c) > 6 for c in x_categories) else 0,
    )
    _decorate_title(ax, title, colors[0] if colors else pal["title"])
    if len(series_list) > 1 or any(s.get("name") for s in series_list):
        _style_legend(ax, handles)
    _hide_spines(ax)
    _style_grid(ax, axis="both")


def _render_scatter(ax, series_list: list[dict], title: str) -> None:
    """散点图。"""
    if not series_list:
        return
    pal = _palette()
    colors = _colors(len(series_list))
    for i, s in enumerate(series_list):
        data = s.get("data") or []
        if isinstance(data, list) and data and isinstance(data[0], list):
            xs = [d[0] for d in data]; ys = [d[1] for d in data]
        else:
            _, values = _series_data_to_pairs(s)
            xs = list(range(len(values))); ys = values
        ax.scatter(xs, ys, color=colors[i], label=_clean_label(s.get("name"), 32),
                   s=50, alpha=0.7, edgecolors="white")
    _decorate_title(ax, title, colors[0] if colors else pal["title"])
    _hide_spines(ax)
    _style_grid(ax, axis="both")


def _render_radar(ax, series_list: list[dict], title: str) -> None:
    """雷达图。

    注意：``ax`` 必须是 polar 投影的 Axes,由调用方传进来。
    """
    if not series_list:
        return
    indicator: list[str] = []
    for s in series_list:
        names, _ = _series_data_to_pairs(s)
        if names and any(n and not n.startswith("系列") for n in names):
            indicator = names; break
    if not indicator:
        first = series_list[0].get("data") or []
        if isinstance(first, list) and first:
            if isinstance(first[0], dict) and isinstance(first[0].get("value"), list):
                indicator = [f"维度{i + 1}" for i in range(len(first[0]["value"]))]
            else:
                indicator = [f"维度{i + 1}" for i in range(len(first))]
    angles = np.linspace(0, 2 * np.pi, max(1, len(indicator)), endpoint=False).tolist()
    angles += angles[:1]
    ax.set_theta_offset(np.pi / 2); ax.set_theta_direction(-1)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(indicator[: len(angles) - 1], fontsize=10)
    colors = _colors(len(series_list))
    for i, s in enumerate(series_list):
        data = s.get("data") or []
        values: list[float] = []
        if isinstance(data, list) and data:
            for item in data:
                if isinstance(item, dict):
                    v = item.get("value")
                    if isinstance(v, list):
                        values.extend(float(x) for x in v if isinstance(x, (int, float)))
                    elif isinstance(v, (int, float)):
                        values.append(float(v))
        if not values:
            continue
        values += values[:1]
        label = _clean_label(s.get("name"), 32) or f"系列{i + 1}"
        ax.plot(angles, values, color=colors[i], linewidth=2, label=label)
        ax.fill(angles, values, color=colors[i], alpha=0.18)
    # 雷达图是 polar 投影,坐标区是正方形,装饰线会压到雷达网格外,
    # 因此只做"左对齐标题",不加 accent 短线。
    _decorate_title(ax, title, colors[0] if colors else _palette()["title"],
                    is_polar=True)
    ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.0), fontsize=9)


def _render_heatmap(ax, series_list: list[dict], title: str) -> None:
    """热力图。data = [[x, y, value], ...]。"""
    if not series_list:
        return
    all_pts: list[tuple[int, int, float]] = []
    for s in series_list:
        data = s.get("data") or []
        if isinstance(data, list):
            for item in data:
                if isinstance(item, (list, tuple)) and len(item) >= 3:
                    all_pts.append((int(item[0]), int(item[1]), float(item[2])))
    if not all_pts:
        ax.text(0.5, 0.5, "(无数据)", ha="center", va="center",
                fontsize=14, color="#94a3b8")
        return
    xs = [p[0] for p in all_pts]; ys = [p[1] for p in all_pts]
    vals = [p[2] for p in all_pts]
    heat = np.full((max(ys) + 1, max(xs) + 1), np.nan)
    for x, y, v in all_pts:
        heat[y, x] = v
    im = ax.imshow(heat, aspect="auto", cmap="YlOrRd")
    plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    _decorate_title(ax, title, _palette()["title"])


def _render_funnel(ax, series_list: list[dict], title: str) -> None:
    """漏斗图：水平 bar,顶部对齐。"""
    if not series_list:
        return
    names: list[str] = []; values: list[float] = []
    for s in series_list:
        n, v = _series_data_to_pairs(s)
        names.extend(n); values.extend(v)
    if not values:
        return
    colors = _colors(len(values))
    ys = list(range(len(values)))[::-1]
    pal = _palette()
    for i, (n, v) in enumerate(zip(names, values)):
        ax.barh(ys[i], v, color=colors[i], label=n)
        ax.text(v, ys[i], f"  {n} {v:.0f}", va="center",
                fontsize=9, color=pal["label"])
    ax.set_yticks(ys); ax.set_yticklabels([""] * len(values))
    _decorate_title(ax, title, colors[0] if colors else pal["title"])
    _hide_spines(ax)


# ============================================================
# 入口：render_chart_to_png
# ============================================================
def render_chart_to_png(chart: dict, width: int = 8, height: int = 5,
                         dpi: int = 150, theme: str = "dark") -> bytes | None:
    """把单张 chart_option 渲染为 PNG 字节。

    Args:
        chart: 单张图表配置,形态见模块顶部说明。
        width / height / dpi: 输出图大小。
        theme: ``dark``(默认,深底浅字,适合网页/深色 UI)或
               ``light``(白底深字,适合打印与放进浅色 PPT/Word 页面)。

    Returns:
        PNG 字节。渲染失败时返回 ``None``,调用方应退回到"贴 JSON 配置"分支。

    【2026-10 chart-visual-upgrade】签名与返回语义保持不变,只新增了
    "把 option.color 注入上下文"这一步 —— 让 3 套风格色盘在导出件里真正生效。
    """
    # 设置图表主题上下文:各 _render_* 内部的文字色从这里取,
    # 避免给每个渲染函数都加 theme 参数。
    token = _CHART_THEME.set(theme)
    # 【新增】把 option.color(风格色盘)注入上下文,_colors() 优先取它。
    # 用 contextvars 而非全局变量,保证并发导出互不串色。
    color_token = None
    try:
        palette = _extract_series_colors(chart)
        if palette:
            color_token = _CHART_SERIES_COLORS.set(palette)
        return _render_chart_to_png_inner(
            chart, width=width, height=height, dpi=dpi, theme=theme
        )
    finally:
        # 无论成功失败都要还原,否则会把色盘/主题泄漏给后续调用
        if color_token is not None:
            _CHART_SERIES_COLORS.reset(color_token)
        _CHART_THEME.reset(token)


def _render_chart_to_png_inner(chart: dict, width: int, height: int,
                               dpi: int, theme: str) -> bytes | None:
    """render_chart_to_png 的真正实现(主题上下文已在调用方设置)。"""
    try:
        # ---- 解析 ----
        option = chart.get("option") if isinstance(chart.get("option"), dict) else chart
        series_list = option.get("series") or []
        if not isinstance(series_list, list) or not series_list:
            return None
        chart_type = (series_list[0].get("type") or "bar").lower()

        title_obj = option.get("title") or chart.get("title") or ""
        if isinstance(title_obj, dict):
            title = title_obj.get("text") or ""
        else:
            title = title_obj

        x_axis_obj = option.get("xAxis")
        y_axis_obj = option.get("yAxis")
        x_categories: list[str] = []
        # 先读 xAxis.data(经典笛卡尔坐标系:分类在 x 轴、值在 y 轴)
        if isinstance(x_axis_obj, dict):
            data = x_axis_obj.get("data")
            if isinstance(data, list):
                x_categories = [_clean_label(d, 32) for d in data]
        # 水平条形图修复：ECharts 中水平 bar 用"xAxis.type=value + yAxis.type=category"，
        # 此时分类数据其实在 yAxis.data 里。如果 xAxis 没 data 但 yAxis 有,
        # 把 yAxis 的分类当 categories 用,这是 matplotlib 渲染水平条形图所必须的。
        if (
            not x_categories
            and isinstance(y_axis_obj, dict)
            and isinstance(y_axis_obj.get("data"), list)
        ):
            x_categories = [_clean_label(d, 32) for d in y_axis_obj["data"]]
        if not x_categories:
            n, v = _series_data_to_pairs(series_list[0])
            x_categories = _x_axis_categories(series_list[0], n, v)

        # ---- 画布 ----
        # 雷达图必须 polar 投影;其它用默认 cartesian
        projection = "polar" if chart_type == "radar" else None
        fig, ax = plt.subplots(
            figsize=(width, height), dpi=dpi, subplot_kw={"projection": projection} if projection else None,
        )
        # 配色:light 主题是白底深字(放进浅色 PPT / 打印场景);
        # dark 主题保持原样(网页深色 UI 用)。
        if theme == "light":
            bg, grid, tick, title_c = "#ffffff", "#e5e7eb", "#4b5563", "#111827"
        else:
            bg, grid, tick, title_c = "#0f172a", "#1e293b", "#94a3b8", "#e2e8f0"
        fig.patch.set_facecolor(bg)
        if projection is None:
            ax.set_facecolor(bg)
            for sp in ax.spines.values():
                sp.set_color(grid)
            ax.tick_params(colors=tick)
        else:
            ax.set_facecolor(bg)
            ax.spines["polar"].set_color(grid)
            ax.tick_params(colors=tick)
        ax.title.set_color(title_c)

        # ---- 分发 ----
        if chart_type == "pie":
            _render_pie(ax, series_list, title)
        elif chart_type == "bar":
            is_h = isinstance(x_axis_obj, dict) and "value" in str(
                x_axis_obj.get("type", "")
            ).lower()
            _render_bar(ax, series_list, x_categories, title, is_horizontal=is_h)
        elif chart_type == "line":
            _render_line(ax, series_list, x_categories, title)
        elif chart_type == "scatter":
            _render_scatter(ax, series_list, title)
        elif chart_type == "radar":
            _render_radar(ax, series_list, title)
        elif chart_type == "heatmap":
            _render_heatmap(ax, series_list, title)
        elif chart_type == "funnel":
            _render_funnel(ax, series_list, title)
        else:
            _render_bar(ax, series_list, x_categories, title)

        fig.tight_layout()
        buf = io.BytesIO()
        # 【关键】不使用 bbox_inches="tight"。
        #   "tight" 会按内容裁掉画布外的空白,使输出 PNG 的宽高比**脱离** figsize
        #   的控制 —— 例如 figsize=(8,5) 的柱状图,裁完可能变成 8:4.78。
        #   导出 PPT 时按宽度固定缩放,高度就会跟着失控:
        #   实测 7.7" 宽的图在页面上高 4.78",底部落到 9.13",
        #   而画布只有 7.5" —— 图直接溢出页面并压住正文。
        #   改成固定画布尺寸,宽高比才是可预期的(width/height)。
        fig.savefig(buf, format="png", facecolor=fig.get_facecolor())
        plt.close(fig)
        buf.seek(0)
        return buf.getvalue()
    except Exception as exc:
        logger.warning("ECharts→matplotlib 渲染失败: %s", exc)
        try:
            plt.close("all")
        except Exception:
            pass
        return None


# ============================================================
# 【chart-edit-prebuilt-dashboard 新增】前置可视化看板:3 套风格 + 规则引擎
# ------------------------------------------------------------
# 背景:
#   用户在工作台上传数据集、尚未执行【开始 AI 分析】时,图表编辑弹窗需要
#   "前置"生成 6~8 张可视化看板,供用户挑选并加入报表。本模块在 chart_renderer
#   里补齐这套能力的"风格常量 + 规则生成 option"两部分:
#       - 3 套风格常量(配色 / 字体 / 网格 / 标签)
#       - apply_style(option, style_key) :给已有 ECharts option 覆盖风格
#       - build_dashboard_options(df, style_key) :基于 Pandas df 推断字段类型
#           + 规则生成 N 张图表 option(柱/折/饼/雷达/面积/堆叠/横向/漏斗等),
#           整看板一次性返回 [{title, option, style_key}, ...]
#
# 设计要点:
#   - **不调 LLM**:纯本地 Pandas + 规则,前端打开弹窗 → 点【生成新图表】即可秒级出图。
#   - **风格字段全部直接覆盖 option**:与 ai_service 生成的结构兼容,前端 ECharts 可直接渲染。
#   - **稳健降级**:df 列数 / 数值列不足时返回尽可能多的图(>=3 张),不让看板空荡。
# ============================================================

# 3 套图表风格常量。
# color:  ECharts option.color(系列调色盘)
# textStyle: 全局字体
# grid:   网格 / 背景 / 标签色
STYLE_TECH_ADVANCED = "tech_advanced"  # 高级科技风格(默认)
STYLE_BUSINESS_PRO = "business_pro"    # 商务职业风格
STYLE_MINIMALIST = "minimalist"        # 极简汇风格


_STYLE_PRESETS: dict[str, dict] = {
    STYLE_TECH_ADVANCED: {
        "label": "高级科技风格",
        # 青绿 + 霓虹 + 紫粉,深色科技感
        "color": ["#2dd4bf", "#38bdf8", "#a855f7", "#ec4899", "#f59e0b",
                   "#22d3ee", "#818cf8", "#34d399"],
        "textStyle": {"color": "#e2e8f0", "fontFamily": "PingFang SC, Microsoft YaHei, sans-serif"},
        "title": {"textStyle": {"color": "#2dd4bf", "fontWeight": "bold"}},
        "grid": {"borderColor": "#1e3a6f", "backgroundColor": "#0f2349"},
        "axisLine": {"lineStyle": {"color": "#1e3a6f"}},
        "axisLabel": {"color": "#94a3b8"},
        "legend": {"textStyle": {"color": "#e2e8f0"}},
        "tooltip": {"backgroundColor": "#0f172a", "borderColor": "#2dd4bf",
                     "textStyle": {"color": "#e2e8f0"}},
    },
    STYLE_BUSINESS_PRO: {
        "label": "商务职业风格",
        # 深蓝 + 橙色 + 商务灰,白底专业
        "color": ["#1e3a8a", "#f59e0b", "#0e7490", "#b91c1c", "#4b5563",
                   "#7c3aed", "#0369a1", "#65a30d"],
        "textStyle": {"color": "#1f2937", "fontFamily": "PingFang SC, Microsoft YaHei, sans-serif"},
        "title": {"textStyle": {"color": "#1e3a8a", "fontWeight": "bold"}},
        "grid": {"borderColor": "#e5e7eb", "backgroundColor": "#ffffff"},
        "axisLine": {"lineStyle": {"color": "#9ca3af"}},
        "axisLabel": {"color": "#4b5563"},
        "legend": {"textStyle": {"color": "#1f2937"}},
        "tooltip": {"backgroundColor": "#ffffff", "borderColor": "#1e3a8a",
                     "textStyle": {"color": "#1f2937"}},
    },
    STYLE_MINIMALIST: {
        "label": "极简汇风格",
        # 蓝绿主色 + 灰轴,大量留白
        "color": ["#10b981", "#64748b", "#0ea5e9", "#f97316", "#a3a3a3",
                   "#84cc16", "#06b6d4", "#facc15"],
        "textStyle": {"color": "#475569", "fontFamily": "PingFang SC, Microsoft YaHei, sans-serif"},
        "title": {"textStyle": {"color": "#0f172a", "fontWeight": "normal"}},
        "grid": {"borderColor": "#e2e8f0", "backgroundColor": "#f8fafc"},
        "axisLine": {"show": False},
        "axisLabel": {"color": "#64748b"},
        "legend": {"textStyle": {"color": "#475569"}},
        "tooltip": {"backgroundColor": "#f8fafc", "borderColor": "#cbd5e1",
                     "textStyle": {"color": "#0f172a"}},
    },
}


def get_style_preset(style_key: str) -> dict:
    """取风格预设;非法 key 兜底为高级科技风格(默认)。"""
    return _STYLE_PRESETS.get(style_key, _STYLE_PRESETS[STYLE_TECH_ADVANCED])


def list_style_keys() -> list[dict]:
    """返回风格枚举,供前端下拉框使用。

    形态:[{key, label}, ...],前端 el-option 直接 v-for。
    """
    return [
        {"key": k, "label": v["label"]}
        for k, v in _STYLE_PRESETS.items()
    ]


def apply_style(option: dict, style_key: str) -> dict:
    """把 ECharts option 按指定风格覆盖(不破坏已有 series / title 文本)。

    - color: 系列调色盘
    - textStyle: 全局字体 + 颜色
    - title.textStyle / legend.textStyle / tooltip.textStyle
    - 各类 backgroundColor / borderColor

    Returns:
        新的 option dict(避免污染入参;若风格非法,返回入参副本)。
    """
    if not isinstance(option, dict):
        return {}
    preset = get_style_preset(style_key)
    out = json_safe_copy(option)

    out["color"] = preset["color"]
    out["textStyle"] = {**(out.get("textStyle") or {}), **preset["textStyle"]}

    # title 既有 text(用户自定义)也有 textStyle(风格颜色)
    title_obj = out.get("title")
    if isinstance(title_obj, dict):
        out["title"] = {
            **title_obj,
            "textStyle": {**(title_obj.get("textStyle") or {}),
                          **preset["title"]["textStyle"]},
        }
    elif isinstance(title_obj, str) and title_obj:
        out["title"] = {"text": title_obj,
                         "textStyle": dict(preset["title"]["textStyle"])}

    # legend / tooltip 同步注入
    legend_obj = out.get("legend")
    if isinstance(legend_obj, dict):
        out["legend"] = {
            **legend_obj,
            "textStyle": {**(legend_obj.get("textStyle") or {}),
                          **preset["legend"]["textStyle"]},
        }
    tooltip_obj = out.get("tooltip")
    if isinstance(tooltip_obj, dict):
        out["tooltip"] = {
            **tooltip_obj,
            "backgroundColor": preset["tooltip"]["backgroundColor"],
            "borderColor": preset["tooltip"]["borderColor"],
            "textStyle": {**(tooltip_obj.get("textStyle") or {}),
                          **preset["tooltip"]["textStyle"]},
        }

    # grid(笛卡尔坐标系)
    grid_obj = out.get("grid")
    if isinstance(grid_obj, dict):
        out["grid"] = {**grid_obj,
                       "borderColor": preset["grid"]["borderColor"],
                       "backgroundColor": preset["grid"]["backgroundColor"]}
    elif "xAxis" in out or "yAxis" in out:
        # 用户没显式给 grid 时补一个默认值(ECharts 不补也能跑,但加底色更稳)
        out["grid"] = {
            "left": "8%", "right": "8%", "top": "16%", "bottom": "8%",
            "containLabel": True,
            "borderColor": preset["grid"]["borderColor"],
            "backgroundColor": preset["grid"]["backgroundColor"],
        }

    # xAxis / yAxis 同步轴线 / 标签色
    for axis_key in ("xAxis", "yAxis"):
        axes = out.get(axis_key)
        if not isinstance(axes, list):
            axes = [axes] if isinstance(axes, dict) else []
        for ax in axes:
            if not isinstance(ax, dict):
                continue
            ax["axisLine"] = {**(ax.get("axisLine") or {}),
                              **preset["axisLine"]}
            ax["axisLabel"] = {**(ax.get("axisLabel") or {}),
                               **preset["axisLabel"]}
        if len(axes) == 1:
            out[axis_key] = axes[0]
        elif len(axes) > 1:
            out[axis_key] = axes

    # 【chart-visual-upgrade 新增】注入图表质感属性(圆角 / 标签 / symbolSize /
    # 标签避让 / 悬停发光)。
    # 放在 apply_style 里,是因为它是**所有图表的统一出口**(本地规则引擎、
    # AI 生成、用户改图都会走这里),在这里统一注入才能保证
    # 前端预览 / 前端截图 / 报告页 / 导出 PNG 四处观感一致,
    # 不会出现"某条路径生成的图还是老样子"。
    _enhance_series_visuals(out, preset["color"])

    return out


def build_dashboard_options_per_sheet(
    sheets: list[dict],
    style_key: str = STYLE_TECH_ADVANCED,
) -> list[dict]:
    """【chart-edit-prebuilt-dashboard 新增】每个 sheet 独立生成图表,合并返回。

    输入:
        sheets: [{name: str, df: pd.DataFrame}, ...],由 pandas_engine.read_dataframe_per_sheet 返回。

    输出:
        [
          {
            "sheet": "<sheet 名>",
            "charts": [{"title", "option", "style_key", "sheet"}, ...],
          },
          ...
        ]
        - 每个 sheet 内部按其字段自适应生成 3~6 张图(不再固定 7 张)
        - 空 sheet 或字段不足时跳过
        - 兜底:全部 sheet 都生成不出图时,返回 1 张"数据分布"占位图

    设计要点:
        1) **每个 sheet 独立分析字段**,避免不同 sheet 列结构差异把列名搞乱
           (concat 后会出现大量 NaN 列,真正有用的字段被稀释)
        2) **按 sheet 分组渲染**,前端看板按 sheet 分块展示,
           数据归属一目了然(用户能直接看到"销售明细"、"目标设定"、"渠道分析"、
           "产品分析" 4 个 sheet 各自的图表)。
        3) 每个 sheet 内部生成 2~6 张,而不是固定的 6~8 张 —
           字段少就少出,避免出现空图 / 重复图。
    """
    groups: list[dict] = []
    total_charts = 0
    for s in sheets or []:
        if not isinstance(s, dict):
            continue
        name = str(s.get("name") or "")
        df = s.get("df")
        if df is None or len(df) == 0:
            continue
        # 每个 sheet 独立生成(每张图都标记 sheet 来源,前端好分组)
        charts = build_dashboard_options(df, style_key=style_key)
        for c in charts:
            c["sheet"] = name  # 给单图补 sheet 字段(若函数没填则补)
        if not charts:
            continue
        groups.append({"sheet": name, "charts": charts})
        total_charts += len(charts)

    # 全部 sheet 都生成不出图 → 用第一个 sheet 的 df 兜底出 1 张
    if not total_charts and sheets:
        first = sheets[0]
        df = first.get("df")
        if df is not None and len(df) > 0:
            charts = build_dashboard_options(df, style_key=style_key)
            if charts:
                for c in charts:
                    c["sheet"] = str(first.get("name") or "")
                groups.append({"sheet": str(first.get("name") or ""), "charts": charts})

    return groups


def json_safe_copy(obj: Any) -> Any:
    """深拷贝 dict / list(其他类型直接返回)。避免污染入参。"""
    if isinstance(obj, dict):
        return {k: json_safe_copy(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [json_safe_copy(v) for v in obj]
    return obj


def _split_columns(df) -> tuple[list[str], list[str], list[str]]:
    """把 df 的列分为 (数值列, 日期列, 分类列)。

    判定优先级:
        1) 日期列: dtype=datetime / 字符串 50%+ 能解析成 ISO 日期 / 含 "/" "-"
        2) 数值列: dtype numeric 且非 bool;若 object 含 "50%"+ 能转 numeric 也算
        3) 分类列: 其它

    【商品条码智能识别】返回结果附带 product_barcode_cols,通过列名关键词
    + 数据特征识别"商品条码 / SKU / 商品编码"等高价值分析维度,让 AI 推荐图表
    时能以条码作为分析对象,而不是被当作 ID 跳过。

    Returns:
        (numeric_cols, date_cols, cat_cols, product_barcode_cols)
    """
    if df is None or df.empty:
        return [], [], [], []
    numeric_cols: list[str] = []
    date_cols: list[str] = []
    cat_cols: list[str] = []
    product_barcode_cols: list[str] = []  # 【新增】商品条码/SKU 类高价值维度
    import datetime as _dt
    import re as _re_iso
    iso_pattern = _re_iso.compile(r"^(19|20)\d{2}[-/]\d{1,2}[-/]\d{1,2}")
    # 【新增】商品条码/SKU 识别关键词
    _BARCODE_KEYWORDS = ("条码", "sku", "SKU", "商品编码", "商品码", "barcode", "product_code", "item_code", "product_id")
    for c in df.columns:
        s = df[c]
        c_lower = str(c).lower()
        # 1) dtype=datetime64 → 日期列
        try:
            if s.dtype == "datetime64[ns]":
                date_cols.append(str(c))
                continue
        except Exception:
            pass
        # 2) 数值列
        is_num = False
        try:
            is_num = pd_is_numeric_dtype(s) and not pd_is_bool_dtype(s)
        except Exception:
            is_num = False
        if not is_num:
            try:
                coerced = pd.to_numeric(s, errors="coerce")
                if coerced.notna().sum() >= max(1, int(len(s) * 0.5)):
                    is_num = True
            except Exception:
                pass
        if is_num:
            # 【新增】数字形式的商品条码识别:
            # 数字看起来像条码(>=10 位 + 唯一值比例适中)→ 提取为条码字段,
            # 同时保留为数值列(向后兼容,但后续在条码专属逻辑里优先使用)
            c_lower_check = str(c).lower()
            is_barcode_by_num = (any(kw in c_lower_check for kw in (k.lower() for k in _BARCODE_KEYWORDS))
                                 or _looks_like_barcode_number(s))
            if is_barcode_by_num:
                n_unique_num = s.dropna().nunique()
                n_total_num = max(len(s), 1)
                if 5 <= n_unique_num <= int(n_total_num * 0.8):
                    product_barcode_cols.append(str(c))
                    # 注意:不 continue,仍加入 numeric_cols,保留向后兼容
            numeric_cols.append(str(c))
            continue
        # 3) 是否日期字符串(2026-10-12 / 2026/10/12 / 2026年10月12日 等)
        try:
            non_null = s.dropna().astype(str)
            if len(non_null) > 0:
                date_ratio = non_null.str.match(iso_pattern).mean()
                if date_ratio >= 0.5:
                    # 再用 datetime 严格验证
                    try:
                        _dt.datetime.fromisoformat(non_null.iloc[0].replace("/", "-"))
                        date_cols.append(str(c))
                        continue
                    except Exception:
                        pass
        except Exception:
            pass

        # 【新增】商品条码/SKU 识别:
        # 命中关键词 + 唯一值数量适中(>5 且 <数据行数*0.8)→ 商品条码列
        # 这种列本质是分类维度,但价值远高于普通类目,应作为 Top 维度优先分析
        is_barcode = any(kw in c_lower for kw in (k.lower() for k in _BARCODE_KEYWORDS))
        if is_barcode:
            n_unique = s.dropna().nunique()
            n_total = max(len(s), 1)
            if 5 <= n_unique <= int(n_total * 0.8):
                product_barcode_cols.append(str(c))
                # 同时也作为类目维度(保留向后兼容)
                cat_cols.append(str(c))
                continue
        cat_cols.append(str(c))
    return numeric_cols, date_cols, cat_cols, product_barcode_cols


def pd_is_numeric_dtype(s) -> bool:
    """薄包装:剥离 pandas API 差异,1.x/2.x 都跑。"""
    try:
        return bool(pd.api.types.is_numeric_dtype(s))
    except Exception:
        return False


def pd_is_bool_dtype(s) -> bool:
    try:
        return bool(pd.api.types.is_bool_dtype(s))
    except Exception:
        return False


def _looks_like_barcode_number(s: pd.Series) -> bool:
    """判断一个数值列是否"看起来像条码"。

    判定规则(同时满足):
      1) 中位数绝对值 >= 1e9(条码都是 10+ 位)
      2) 至少 50% 样本为整数(条码不含小数)
      3) 唯一值数量 >= 5(不是常量)
    """
    try:
        nums = pd.to_numeric(s, errors="coerce").dropna()
        if len(nums) < 5:
            return False
        # 整数判定:与 round 后的值差异为 0
        int_ratio = (nums == nums.round()).mean()
        if int_ratio < 0.5:
            return False
        # 数值规模
        if nums.abs().median() < 1e9:
            return False
        # 唯一值
        if nums.nunique() < 5:
            return False
        return True
    except Exception:
        return False


def _safe_group_sum(df, cat_col: str, num_col: str, top_n: int = 8) -> tuple[list, list]:
    """按分类列分组聚合数值列,返回 (categories[], values[])。

    - 空值 / 非数值会被丢弃
    - 超过 top_n 时只保留 top_n(by value desc)
    """
    if cat_col not in df.columns or num_col not in df.columns:
        return [], []
    try:
        sub = df[[cat_col, num_col]].dropna()
        # 把 num_col 转 float,失败置 NaN,再 drop
        sub[num_col] = pd.to_numeric(sub[num_col], errors="coerce")
        sub = sub.dropna()
        if sub.empty:
            return [], []
        grouped = sub.groupby(cat_col)[num_col].sum().sort_values(ascending=False)
        if len(grouped) > top_n:
            grouped = grouped.head(top_n)
        return [str(k) for k in grouped.index.tolist()], [float(v) for v in grouped.values]
    except Exception:
        return [], []


def _series_option(
    chart_type: str,
    name: str,
    data: list,
    extra: dict | None = None,
    *,
    show_label: bool = True,
    label_position: str | None = None,
    horizontal: bool = False,
    color_hint: str | None = None,
    palette: list[str] | None = None,
) -> dict:
    """构造一个 series dict,统一 data 形态,并注入"图表质感"视觉属性。

    【2026-10 chart-visual-upgrade 升级】
    原来这里只写 {"type","name","data"} 三个字段,前端 ECharts 只能画出
    "默认样式"的图(直角柱、无渐变、无标签),用户反馈"图表太普通"。
    现在按图型补齐视觉属性(前端预览 / 前端截图 / 报告页三处同步生效):
        - bar  : itemStyle.borderRadius(圆角)+ itemStyle.color(渐变)+ label + 发光
        - pie  : 逐扇区渐变(itemStyle.color)+ 外部标签 + labelLine
        - line : areaStyle 渐变 + symbolSize + label + labelLayout(避让) + 发光

    参数说明:
        show_label:    是否显示数值标签(默认开;调用方可关)
        label_position: 标签位置(柱 top/right、线 top、饼 inside/outside)
        horizontal:    是否横向条形图(影响圆角方向与标签位置)
        color_hint:    该系列的主色(用于渐变/发光);缺省用默认色盘首色
        palette:       完整色盘(饼图逐扇区取色用)

    向后兼容:
        extra 里若显式给了 label / areaStyle / itemStyle / emphasis,
        会与默认值**合并**(显式项优先),因此既有调用方行为完全不变。
    """
    ct = (chart_type or "").lower()
    series: dict[str, Any] = {"name": name, "type": chart_type}

    if ct == "pie":
        # 饼图 data 统一成 {"name","value"};顺带做容错(脏数据不炸整图)
        items: list[dict[str, Any]] = []
        for item in (data or []):
            if not isinstance(item, dict):
                continue
            try:
                value = float(item.get("value", 0) or 0)
            except (TypeError, ValueError):
                value = 0.0
            items.append({"name": str(item.get("name", "")), "value": value})
        series["data"] = items
    else:
        series["data"] = data

    # 系列主色:调用方通常传风格色盘里的颜色,保证渐变与整体配色一致
    base_color = _normalize_color(color_hint) or _PALETTE[0]

    if ct == "bar":
        # 横向条形图的圆角在右侧(数据末端),竖直柱在顶部
        series["itemStyle"] = {
            "color": _linear_gradient(base_color),
            "borderRadius": [0, 4, 4, 0] if horizontal else [4, 4, 0, 0],
        }
        series["label"] = {
            "show": bool(show_label),
            "position": label_position or ("right" if horizontal else "top"),
        }
        series["emphasis"] = {
            "focus": "series",
            "itemStyle": {
                "shadowBlur": 12,
                "shadowColor": _hex_to_rgba(base_color, 0.35),
            },
        }
    elif ct == "pie":
        # 注意:饼图**不能**在 series 级设 itemStyle.color —— 那会让所有扇区同色。
        # 所以 series 级只保留描边,颜色下沉到 data[i].itemStyle 逐扇区给。
        series["itemStyle"] = {"borderColor": "#ffffff", "borderWidth": 2}
        series["label"] = {
            "show": bool(show_label),
            "position": label_position or "outside",
        }
        series["labelLine"] = {"show": bool(show_label)}
        series["emphasis"] = {
            "scale": True,
            "itemStyle": {
                "shadowBlur": 12,
                "shadowColor": _hex_to_rgba(base_color, 0.35),
            },
        }
        pie_palette = [c for c in (palette or []) if _normalize_color(c)]
        if pie_palette:
            for idx, item in enumerate(series["data"]):
                item["itemStyle"] = {
                    "color": _linear_gradient(pie_palette[idx % len(pie_palette)]),
                }
    elif ct == "line":
        series["symbolSize"] = 7
        series["areaStyle"] = {
            "color": _linear_gradient(base_color),
            "opacity": 0.28,
        }
        series["label"] = {
            "show": bool(show_label),
            "position": label_position or "top",
        }
        # ECharts 5 的标签避让:点多了自动隐藏重叠的标签
        series["labelLayout"] = {"hideOverlap": True}
        series["emphasis"] = {
            "focus": "series",
            "itemStyle": {
                "shadowBlur": 12,
                "shadowColor": _hex_to_rgba(base_color, 0.35),
            },
        }
    else:
        # 其它图型(scatter/radar/funnel/heatmap)只补标签开关,不强行改样式
        series["label"] = {
            "show": bool(show_label),
            "position": label_position or "top",
        }

    # extra 合并:显式配置优先,但不整块覆盖(避免把渐变/圆角冲掉)
    if extra:
        for key, value in extra.items():
            if isinstance(value, dict) and isinstance(series.get(key), dict):
                merged = {**series[key], **value}
                # itemStyle / labelLine 这类嵌套字典再合并一层
                for sub_key in ("itemStyle", "labelLine"):
                    if isinstance(value.get(sub_key), dict) and isinstance(
                        merged.get(sub_key), dict
                    ):
                        merged[sub_key] = {**merged[sub_key], **value[sub_key]}
                series[key] = merged
            else:
                series[key] = value
    return series


def _is_horizontal_option(option: dict) -> bool:
    """判断 option 里的柱状图是不是"横向条形图"。

    ECharts 约定:横向 bar = xAxis.type="value" + yAxis.type="category",
    竖直 bar = xAxis.type="category" + yAxis.type="value"。
    """
    def _axis_type(value: Any) -> str:
        if isinstance(value, dict):
            return str(value.get("type") or "").lower()
        if isinstance(value, list) and value and isinstance(value[0], dict):
            return str(value[0].get("type") or "").lower()
        return ""

    x_type = _axis_type(option.get("xAxis"))
    y_type = _axis_type(option.get("yAxis"))
    return x_type == "value" and y_type == "category"


def _enhance_series_visuals(option: dict, palette: list[str]) -> dict:
    """给 option 里的每个 series 补齐"图表质感"属性(原地修改并返回)。

    设计取舍(很重要):
        这里**只注入与调色盘无关**的属性(圆角、标签、symbolSize、
        labelLayout 避让、饼图描边),**不写死 itemStyle.color / areaStyle.color**。
        原因:option.color 才是"最终生效的调色盘"(前端 ECharts 按 series 顺序取色),
        如果在 series 级写死颜色,用户在 workbench 里说"配色换成红色系"时
        就会失效 —— 用户指令必须永远胜出。
        具体颜色由 _series_option() 在**已知最终色盘**的地方注入。

    幂等:重复调用不会叠加出脏字段(全部走 setdefault)。
    """
    series_list = option.get("series")
    if not isinstance(series_list, list):
        return option
    colors = [c for c in (palette or []) if _normalize_color(c)] or list(_PALETTE)
    horizontal = _is_horizontal_option(option)
    for idx, series in enumerate(series_list):
        if not isinstance(series, dict):
            continue
        try:
            ct = str(series.get("type") or "").lower()
            base = colors[idx % len(colors)]

            if ct == "bar":
                item_style = series.get("itemStyle")
                item_style = dict(item_style) if isinstance(item_style, dict) else {}
                item_style.setdefault(
                    "borderRadius", [0, 4, 4, 0] if horizontal else [4, 4, 0, 0]
                )
                series["itemStyle"] = item_style
                label = series.get("label")
                label = dict(label) if isinstance(label, dict) else {}
                label.setdefault("show", True)
                label.setdefault("position", "right" if horizontal else "top")
                series["label"] = label
            elif ct == "line":
                series.setdefault("symbolSize", 7)
                area_style = series.get("areaStyle")
                area_style = dict(area_style) if isinstance(area_style, dict) else {}
                # 只给透明度,不写死颜色(让 ECharts 用 option.color 派生)
                area_style.setdefault("opacity", 0.28)
                series["areaStyle"] = area_style
                label = series.get("label")
                label = dict(label) if isinstance(label, dict) else {}
                label.setdefault("show", True)
                label.setdefault("position", "top")
                series["label"] = label
                layout = series.get("labelLayout")
                layout = dict(layout) if isinstance(layout, dict) else {}
                layout.setdefault("hideOverlap", True)
                series["labelLayout"] = layout
            elif ct == "pie":
                item_style = series.get("itemStyle")
                item_style = dict(item_style) if isinstance(item_style, dict) else {}
                item_style.setdefault("borderColor", "#ffffff")
                item_style.setdefault("borderWidth", 2)
                series["itemStyle"] = item_style
                label = series.get("label")
                label = dict(label) if isinstance(label, dict) else {}
                label.setdefault("show", True)
                label.setdefault("position", "outside")
                series["label"] = label
                series.setdefault("labelLine", {"show": True})

            # 强调态发光:悬停时的高光,让图表"活"起来
            emphasis = series.get("emphasis")
            emphasis = dict(emphasis) if isinstance(emphasis, dict) else {}
            em_item = emphasis.get("itemStyle")
            em_item = dict(em_item) if isinstance(em_item, dict) else {}
            em_item.setdefault("shadowBlur", 12)
            em_item.setdefault("shadowColor", _hex_to_rgba(base, 0.35))
            emphasis["itemStyle"] = em_item
            series["emphasis"] = emphasis
        except Exception as exc:
            # 单个 series 失败不影响其它 series
            logger.debug("注入图表质感属性失败(已跳过该 series): %s", exc)
    return option


def build_dashboard_options(df, style_key: str = STYLE_TECH_ADVANCED) -> list[dict]:
    """根据 df 生成 3~6 张前置看板图表的列表(本地规则引擎,无 LLM)。

    输入:
        df: pd.DataFrame(读归一后的"原始数据")
        style_key: 风格 key;非法 → 默认 tech_advanced

    输出:
        [{"title": str, "option": dict, "style_key": str, "sheet": str?}, ...]
        - sheet 字段:多 sheet 数据时携带来源 sheet 名,前端按 sheet 分组渲染
        - 单 sheet / CSV 时,sheet 字段可能为空

    【2026-10 升级】结合 _profile_dataframe 画像自适应:
        1) 数据稀薄(行<5 或无数值)→ 仅 1 张分布图;
        2) 单分类无时间维度 → 3 张(横 Top / 饼 / 竖柱);
        3) 含时间序列 + 单一数值 → 4 张(+ 趋势);
        4) 多维度(2 分类 + 2 数值) → 最多 6 张。
        同时在生成前剔除「合计 / 全年合计 / 总计 / 平均」等汇总行(避免混入
        占比维度),并对每张 option 做 _is_chart_meaningful 校验,问题图直接丢弃。
    """
    out: list[dict] = []
    if df is None or len(df) == 0 or len(df.columns) == 0:
        return out

    # ---- 第一步:数据画像 → 决定预算图表数量 ----
    profile = _profile_dataframe(df)

    # 【chart-visual-upgrade 新增】当前风格的色盘。
    # 局部规则引擎生成的 series 会用它给渐变/发光取色 —— 这样"风格色盘"
    # 既是 option.color,也是 series 渐变的取色来源,两者天然一致。
    style_palette: list[str] = list(
        get_style_preset(style_key).get("color") or _PALETTE
    )

    # ---- 第二步:剔除汇总行(合计/总计/全年合计/平均 等)----
    # 必须在做类型判定 / groupby 之前完成,否则会被算进切片。
    numeric_cols, date_cols, cat_cols, product_barcode_cols = _split_columns(df)
    df_clean = _strip_summary_rows(df, cat_cols + date_cols)
    # 汇总行剔除后重新跑类型判定(行数变化、唯一值变化)
    if len(df_clean) != len(df):
        numeric_cols, date_cols, cat_cols, product_barcode_cols = _split_columns(df_clean)

    # ---- 第二点五步:日期粒度自适应(2026-10 新增)----
    # 用户截图场景:「日期」列实际是 "2026-10-15 00:00:00" 这种时间戳,
    # 11 行 → Top 10 变成 11 个时间戳、环形图变成 11 个等分切片。
    # 这里若日期列唯一值 > 12 且粒度 < 日,自动按日/月/季聚合到 ≤12 个分类,
    # 并把日期列降级为普通分类列(后续模板按分类走,不再按时间序列模板走)。
    _date_agg_granularity = None  # 记录实际聚合粒度,用于标题显示
    if date_cols:
        d0 = date_cols[0]
        gran = _detect_date_granularity(df_clean[d0]) if d0 in df_clean.columns else "day"
        # 唯一值 > 12 才需要聚合;否则保留原始粒度
        try:
            uniq = int(df_clean[d0].dropna().nunique()) if d0 in df_clean.columns else 0
        except Exception:
            uniq = 0
        if uniq > 8 and gran in ("second", "minute", "hour", "day"):
            new_df, new_col = _aggregate_date_column(df_clean, d0, gran)
            if new_col != d0 or len(new_df) != len(df_clean):
                df_clean = new_df
                # 日期列被聚合后,从 date_cols 移除(语义上变成普通分类),
                # 把它塞进 cat_cols 让后续"按分类 TopN"路径接管
                date_cols = []
                cat_cols = [new_col] + [c for c in cat_cols if c != new_col]
                _date_agg_granularity = gran
                # 重新跑类型判定
                numeric_cols, _, cat_cols, product_barcode_cols = _split_columns(df_clean)

    # ---- 字段优先级选 ----
    # 1) 主分类:日期列 > cat_cols(0)
    #    日期列天然适合做时间序列 X 轴,比 cat 更友好
    primary_cat = date_cols[0] if date_cols else (cat_cols[0] if cat_cols else None)
    # 2) 主数值:numeric_cols(0);没数值列就 fallback 到"按分类列计数"
    primary_num = numeric_cols[0] if numeric_cols else None
    # 没数值时把 primary_cat 标记成"以计数为主",生成 count 类图
    use_count_mode = not bool(numeric_cols)
    # 3) 第二分类(给"按渠道分类"用)
    second_cat = cat_cols[0] if cat_cols and (primary_cat in date_cols or not date_cols) else (cat_cols[1] if len(cat_cols) > 1 else None)
    if primary_cat and primary_cat in cat_cols:
        # 把 primary_cat 从 cat_cols 抽出来
        cat_cols_for_group = [c for c in cat_cols if c != primary_cat]
        second_cat = cat_cols_for_group[0] if cat_cols_for_group else None

    # 当 use_count_mode 时,用 groupby primary_cat count(*) 当数据
    if use_count_mode and primary_cat:
        try:
            cnt = df_clean.groupby(primary_cat).size().sort_values(ascending=False).head(10)
            primary_num = "__count__"
            # 我们用 _safe_group_sum 不行(num_col 必须存在),
            # 改用一个特殊路径:在生成图时直接读 cnt。
        except Exception:
            pass

    # ---- 第三步:用一个计数器限制生成数量,取代旧的"无脑跑满所有模板" ----
    _chart_budget = profile["recommended_count"]

    def _push(opt: dict, title: str) -> None:
        """校验 + 入库。一站式:_is_chart_meaningful 检查 + 应用风格 + 上限保护。"""
        nonlocal _chart_budget
        if _chart_budget <= 0:
            return
        if not _is_chart_meaningful(opt):
            return
        out.append({
            "title": title,
            "option": apply_style(opt, style_key),
            "style_key": style_key,
        })
        _chart_budget -= 1

    # ============================================================
    # 【商品条码专属多维度图表】商品条码作为高价值 Top 维度,
    # 自动与所有数值字段组合,生成 Top N / 占比 / 趋势 多角度图,
    # 解决"商品条码 × 数量"单一组合导致的图表重复问题。
    # 【2026-10 改造】所有 out.append 改走 _push(校验+限额),
    #               groupby 数据源从 df 改为 df_clean(已剔除汇总行)。
    # ============================================================
    if product_barcode_cols and numeric_cols:
        barcode_col = product_barcode_cols[0]
        # 用多个数值字段分别生成"商品条码 × 数量/金额/退货率"图
        # 避免只跟 primary_num 组合造成的图重复
        for num_col in numeric_cols[:4]:  # 最多取 4 个数值字段
            if num_col == primary_num and len(numeric_cols) == 1:
                # 只有 1 个数值时,沿用主流程(避免重复)
                continue
            cats, vals = _safe_group_sum(df_clean, barcode_col, num_col, top_n=10)
            if cats and vals:
                opt = {
                    "title": {"text": f"{num_col} 按 {barcode_col} Top 10"},
                    "tooltip": {"trigger": "axis"},
                    "legend": {"show": False},
                    "xAxis": {"type": "value"},
                    "yAxis": {"type": "category", "data": list(reversed(cats))},
                    "series": [_series_option("bar", num_col,
                                              list(reversed(vals)),
                                              {"label": {"show": True, "position": "right"}},
                                              horizontal=True,
                                              color_hint=style_palette[0],
                                              palette=style_palette)],
                }
                _push(opt, opt["title"]["text"])

            # 同字段再生成一张饼图(占比角度)
            if num_col != primary_num or len(numeric_cols) > 1:
                cats2, vals2 = _safe_group_sum(df_clean, barcode_col, num_col, top_n=6)
                if cats2 and vals2:
                    opt = {
                        "title": {"text": f"{num_col} 占比(按{barcode_col})"},
                        "tooltip": {"trigger": "item"},
                        "legend": {"bottom": 0},
                        "series": [_series_option("pie", num_col,
                                                   [{"name": c, "value": v} for c, v in zip(cats2, vals2)],
                                                   {"radius": ["35%", "65%"]},
                                                   color_hint=style_palette[0],
                                                   palette=style_palette)],
                    }
                    _push(opt, opt["title"]["text"])
        # 商品条码 × 时间 趋势:展示每个条码随时间的走势(只取 Top 5 条码避免图太乱)
        if date_cols and len(product_barcode_cols) >= 1:
            date_col = date_cols[0]
            try:
                # 选数量最大的 5 个条码作为展示对象
                top_barcodes = (df_clean.groupby(barcode_col)[primary_num or numeric_cols[0]]
                                .sum().sort_values(ascending=False).head(5).index.tolist())
                if top_barcodes:
                    sub = df_clean[df_clean[barcode_col].isin(top_barcodes)]
                    num_for_line = primary_num or numeric_cols[0]
                    pivot = (sub.groupby([date_col, barcode_col])[num_for_line]
                             .sum().unstack(fill_value=0))
                    # 如果日期列是字符串,排序;如果是 datetime,按时间排
                    try:
                        pivot.index = pd.to_datetime(pivot.index)
                        pivot = pivot.sort_index()
                        x_data = [d.strftime("%Y-%m-%d") for d in pivot.index]
                    except Exception:
                        x_data = [str(d) for d in pivot.index]
                    series_list = []
                    for bc_idx, bc in enumerate(top_barcodes):
                        if bc in pivot.columns:
                            series_list.append(_series_option(
                                "line", str(bc),
                                [float(v) for v in pivot[bc].tolist()],
                                {"smooth": True},
                                # 每条线取色盘里的不同颜色,渐变/发光同步跟随
                                color_hint=style_palette[bc_idx % len(style_palette)],
                            ))
                    if series_list:
                        opt = {
                            "title": {"text": f"{num_for_line} 走势(Top 5 {barcode_col})"},
                            "tooltip": {"trigger": "axis"},
                            "legend": {},
                            "xAxis": {"type": "category", "data": x_data},
                            "yAxis": {"type": "value"},
                            "series": series_list,
                        }
                        _push(opt, opt["title"]["text"])
            except Exception:
                pass

    # ---- helper:根据 primary_cat / primary_num 取聚合数据 ----
    # - use_count_mode:对 primary_cat 做 groupby 计数
    # - 否则调 _safe_group_sum
    # 【2026-10 改造】数据源改用 df_clean
    def _agg(cat_col, num_col, top_n=10):
        if use_count_mode:
            try:
                cnt = df_clean.groupby(cat_col).size().sort_values(ascending=False).head(top_n)
                # 转成 _safe_group_sum 形态 (cats, vals)
                return [str(k) for k in cnt.index.tolist()], [int(v) for v in cnt.values]
            except Exception:
                return [], []
        return _safe_group_sum(df_clean, cat_col, num_col, top_n=top_n)

    # 显示名:__count__ 时换成"数量"
    def _display_num(name):
        return "数量" if name == "__count__" else name

    # --- 1. 横向柱状 TopN ---
    if primary_cat and primary_num:
        cats, vals = _agg(primary_cat, primary_num, top_n=10)
        if cats and vals:
            opt = {
                "title": {"text": f"{_display_num(primary_num)} 按 {primary_cat} Top 10"},
                "tooltip": {"trigger": "axis"},
                "legend": {"show": False},
                "xAxis": {"type": "value"},
                "yAxis": {"type": "category", "data": list(reversed(cats))},
                "series": [_series_option("bar", _display_num(primary_num),
                                            list(reversed(vals)),
                                            {"label": {"show": True, "position": "right"}},
                                            horizontal=True,
                                            color_hint=style_palette[0],
                                            palette=style_palette)],
            }
            _push(opt, opt["title"]["text"])

    # --- 2. 饼图(占比)---
    if primary_cat and primary_num:
        cats, vals = _agg(primary_cat, primary_num, top_n=6)
        if cats and vals:
            opt = {
                "title": {"text": f"{_display_num(primary_num)} 占比({primary_cat})"},
                "tooltip": {"trigger": "item"},
                "legend": {"bottom": 0},
                "series": [_series_option("pie", _display_num(primary_num),
                                            [{"name": c, "value": v} for c, v in zip(cats, vals)],
                                            {"radius": ["35%", "65%"]},
                                            color_hint=style_palette[0],
                                            palette=style_palette)],
            }
            _push(opt, opt["title"]["text"])

    # --- 3. 柱状(主数值,首列分组)---
    if primary_cat and primary_num:
        cats, vals = _agg(primary_cat, primary_num, top_n=8)
        if cats and vals:
            opt = {
                "title": {"text": f"{_display_num(primary_num)} 分布({primary_cat})"},
                "tooltip": {"trigger": "axis"},
                "legend": {"show": False},
                "xAxis": {"type": "category", "data": cats},
                "yAxis": {"type": "value"},
                "series": [_series_option("bar", _display_num(primary_num), vals,
                                            {"label": {"show": True, "position": "top"}},
                                            color_hint=style_palette[0],
                                            palette=style_palette)],
            }
            _push(opt, opt["title"]["text"])

    # --- 4. 折线(时间序列 / 第二数值列)---
    if date_cols and primary_num:
        # 有日期列 → 时间序列折线
        num_for_line = numeric_cols[0] if numeric_cols else None
        if num_for_line:
            cats, vals = _agg(date_cols[0], num_for_line, top_n=12)
            if cats and vals:
                opt = {
                    "title": {"text": f"{_display_num(num_for_line)} 趋势(按 {date_cols[0]})"},
                    "tooltip": {"trigger": "axis"},
                    "legend": {"show": False},
                    "xAxis": {"type": "category", "data": cats},
                    "yAxis": {"type": "value"},
                    "series": [_series_option("line", _display_num(num_for_line), vals, {
                        "smooth": True,
                        "areaStyle": {"opacity": 0.25},
                        "label": {"show": True},
                    }, color_hint=style_palette[0], palette=style_palette)],
                }
                _push(opt, opt["title"]["text"])
    elif len(numeric_cols) >= 2:
        # 没日期列但有多数值 → 第二数值列趋势
        num2 = numeric_cols[1]
        if primary_cat:
            cats, vals = _agg(primary_cat, num2, top_n=12)
        else:
            try:
                s = pd.to_numeric(df_clean[num2], errors="coerce").dropna().head(12)
                cats = [str(i + 1) for i in range(len(s))]
                vals = [float(v) for v in s.values]
            except Exception:
                cats, vals = [], []
        if cats and vals:
            opt = {
                "title": {"text": f"{num2} 趋势"},
                "tooltip": {"trigger": "axis"},
                "legend": {"show": False},
                "xAxis": {"type": "category", "data": cats},
                "yAxis": {"type": "value"},
                "series": [_series_option("line", num2, vals, {
                    "smooth": True,
                    "areaStyle": {"opacity": 0.25},
                    "label": {"show": True},
                }, color_hint=style_palette[0], palette=style_palette)],
            }
            _push(opt, opt["title"]["text"])

    # --- 5. 按第二分类拆分(堆叠柱)---
    # primary_cat + primary_num + second_cat:堆叠柱"primary_num 按 second_cat 拆分(按 primary_cat 分组)"
    if primary_cat and primary_num and second_cat:
        try:
            top_n = 6
            grouped_main = _agg(primary_cat, primary_num, top_n=top_n)
            top_cats = grouped_main[0]
            if top_cats:
                # 对每个 top_cat,按 second_cat 聚合计数 / 求和
                series_list = []
                sec_vals_set = set()
                sec_to_vals = {}
                for cat in top_cats:
                    sub = df_clean[df_clean[primary_cat] == cat]
                    try:
                        if use_count_mode:
                            cnt = sub.groupby(second_cat).size()
                            d = {str(k): int(v) for k, v in cnt.items()}
                        else:
                            sub2 = sub[[second_cat, primary_num]].dropna()
                            sub2[primary_num] = pd.to_numeric(sub2[primary_num], errors="coerce")
                            sub2 = sub2.dropna()
                            d = sub2.groupby(second_cat)[primary_num].sum().to_dict()
                            d = {str(k): float(v) for k, v in d.items()}
                    except Exception:
                        d = {}
                    sec_to_vals[cat] = d
                    sec_vals_set.update(d.keys())
                sec_vals_list = sorted(sec_vals_set)[:6]
                for sec_idx, sec in enumerate(sec_vals_list):
                    arr = [float(sec_to_vals.get(c, {}).get(sec, 0)) for c in top_cats]
                    series_list.append(_series_option(
                        "bar", sec, arr,
                        # 堆叠柱的每个分段取色盘里的不同颜色 + 渐变
                        color_hint=style_palette[sec_idx % len(style_palette)],
                    ))
                if series_list:
                    opt = {
                        "title": {"text": f"{_display_num(primary_num)} 按 {second_cat} 拆分(按 {primary_cat})"},
                        "tooltip": {"trigger": "axis"},
                        "legend": {},
                        "xAxis": {"type": "category", "data": top_cats},
                        "yAxis": {"type": "value"},
                        "series": series_list,
                    }
                    _push(opt, opt["title"]["text"])
        except Exception:
            pass

    # --- 6. 第二分类的占比饼图 ---
    if second_cat:
        try:
            if use_count_mode:
                cnt = df_clean.groupby(second_cat).size().sort_values(ascending=False).head(6)
                cats, vals = list(cnt.index.astype(str)), [int(v) for v in cnt.values]
            else:
                cats, vals = _safe_group_sum(df_clean, second_cat, primary_num if primary_num and primary_num != "__count__" else (numeric_cols[0] if numeric_cols else second_cat), top_n=6)
            if cats and vals:
                opt = {
                    "title": {"text": f"{second_cat} 占比"},
                    "tooltip": {"trigger": "item"},
                    "legend": {"bottom": 0},
                    "series": [_series_option("pie", second_cat,
                                                [{"name": c, "value": v} for c, v in zip(cats, vals)],
                                                {"radius": ["35%", "65%"]},
                                                color_hint=style_palette[0],
                                                palette=style_palette)],
                }
                _push(opt, opt["title"]["text"])
        except Exception:
            pass

    # --- 7. 散点(num1 vs num2)---
    if len(numeric_cols) >= 2:
        try:
            sub = df_clean[[numeric_cols[0], numeric_cols[1]]].apply(
                lambda s: pd.to_numeric(s, errors="coerce")
            ).dropna().head(200)
            data = [[float(r[0]), float(r[1])] for r in sub.values]
            if data:
                opt = {
                    "title": {"text": f"{numeric_cols[0]} vs {numeric_cols[1]} 散点"},
                    "tooltip": {"trigger": "item"},
                    "xAxis": {"type": "value", "name": numeric_cols[0]},
                    "yAxis": {"type": "value", "name": numeric_cols[1]},
                    "series": [{
                        "name": f"{numeric_cols[0]} × {numeric_cols[1]}",
                        "type": "scatter",
                        "data": data,
                        "symbolSize": 8,
                    }],
                }
                _push(opt, opt["title"]["text"])
        except Exception:
            pass

    # --- 兜底:确保至少 1 张图 ---
    # 注:兜底不强制 _push(否则会被预算拦掉导致 0 张),直接 append
    if not out:
        try:
            s = df_clean.iloc[:, 0]
            opt = {
                "title": {"text": "数据分布(第 1 列)"},
                "tooltip": {"trigger": "axis"},
                "xAxis": {"type": "category",
                          "data": [str(i) for i in range(min(20, len(s)))]},
                "yAxis": {"type": "value"},
                "series": [_series_option(
                    "bar", str(df_clean.columns[0]),
                    [float(v) if pd_is_numeric_dtype(s) else 1
                     for v in s.head(20)],
                    color_hint=style_palette[0],
                    palette=style_palette,
                )],
            }
            out.append({"title": opt["title"]["text"],
                          "option": apply_style(opt, style_key),
                          "style_key": style_key})
        except Exception:
            pass

    return out
