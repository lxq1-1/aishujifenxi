"""分析编排服务：串联解析 → 清洗 → 指标 → AI。"""
from __future__ import annotations

import logging
import re
from typing import Any

from sqlalchemy.orm import Session

from app.core.json_utils import dumps
from app.models.dataset import Dataset
from app.models.report_template import ReportTemplate
from app.models.task import Task
from app.services import ai_service, pandas_engine as pe

logger = logging.getLogger(__name__)


def _generate_local_section_text(section_key: str, indicators: dict[str, Any]) -> str:
    """【已废弃】原"LLM 占位文案 → 本地拼接"的兜底函数。

    【2026-10 移除依据】
        旧实现用该函数把"29 行 × 24 列"这种纯结构化拼接写进报告,
        看似"有数字",实际是把 AI 报告降级成了指标 dump。
        与"AI 真写分析"的产品诉求完全背离,因此 _run_analysis_core
        已不再调用本函数;函数体保留为空串返回,避免误引用。

    新策略:
        - LLM 写出正文 → 原样保留
        - LLM 真失败 → 主链路异常 → task.status='failed',根本不会到这里
    """
    logger.warning(
        "[analyzer] _generate_local_section_text 已被废弃,"
        "调用者请改为让 LLM 直接生成章节正文。"
    )
    return ""


def _merge_sheet_indicators(sheet_indicators_list: list[dict]) -> dict[str, Any]:
    """把多张 sheet 的 compute_indicators 输出合并成一个"顶层汇总" dict。

    设计要点(2026-10 多 sheet 分项):
      - 旧实调用 read_dataframe_all_sheets 把所有 sheet 拼成一张大表后再算一次指标,
        LLM 拿到的指标会把跨表结构差异补的 NaN 也算进去,严重失真。
      - 现在改为逐 sheet 算指标,再合并成顶层 indicators,key 加 sheet 前缀避免撞名。
      - 顶层同时保留"行数/列数/缺失数"的总和,便于 Prompt 给出整体规模。
      - 聚合类字段(numeric_summary/topn/by_group/trend/text_value_counts)按 sheet 拆分,
        不再做跨 sheet 求和,避免数据语义错乱。

    Args:
        sheet_indicators_list: 每张 sheet 的指标 dict,来自 pe.compute_indicators()。
                                其中"empty"=True 的 sheet(0 行)会被跳过。

    Returns:
        合并后的 indicators dict,key 命名约定:
            - row_count / column_count / missing_top5 / missing_ratio_top5: 跨 sheet 求和
            - numeric_summary / topn / by_group / trend / text_value_counts: 拆 sheet 存
              (key 形如 "numeric_summary::销售明细"),避免跨 sheet 数据相加
            - columns: 跨 sheet 合并(同名列只保留首个 sheet 的判定)
    """
    if not sheet_indicators_list:
        return {}

    merged: dict[str, Any] = {
        "row_count": 0,
        "column_count": 0,
        "columns": [],
        "numeric_summary": {},
        "topn": {},
        "by_group": {},
        "trend": {},
        "missing_top5": {},
        "missing_ratio_top5": {},
        "text_value_counts": {},
    }
    # 跨 sheet 累计的缺失计数(后续按总值排序取 top5)
    overall_missing_counter: dict[str, int] = {}
    # 【2026-10 新增】KPI 宽表可能出现在多张 sheet(如"年度核心KPI" +
    # "季度核心KPI"),这里按 sheet 名分键存放,与其它聚合字段口径一致,
    # 避免两个 sheet 的 KPI 记录混成一张表导致指标重名冲突。
    kpi_by_sheet: dict[str, Any] = {}

    for s_ind in sheet_indicators_list:
        if not isinstance(s_ind, dict):
            continue
        if s_ind.get("empty"):
            # 空 sheet,只贡献 row_count=0
            continue
        s_name = s_ind.get("name") or ""
        merged["row_count"] += int(s_ind.get("row_count", 0))
        merged["column_count"] += int(s_ind.get("column_count", 0))
        # columns 字段用同 sheet 跨 sheet 合并(同名列保留首次出现的类型)
        seen_col_names = {str(r.get("name")) for r in merged["columns"]}
        for col in (s_ind.get("columns") or []):
            cn = str(col.get("name"))
            if cn not in seen_col_names:
                merged["columns"].append(col)
                seen_col_names.add(cn)
        # 拆 sheet 存的字段,key 加前缀
        prefix = f"{s_name}::" if s_name else ""
        for field in ("numeric_summary", "topn", "by_group", "trend", "text_value_counts"):
            inner = s_ind.get(field) or {}
            for k, v in inner.items():
                merged[field][f"{prefix}{k}"] = v
        # 【2026-10 新增】KPI 宽表:按 sheet 分键合并。
        # 不加前缀直接覆盖会导致多张 KPI sheet 只剩最后一张,
        # 前面 sheet 的真实指标值会从数值池消失,进而让图表防编造校验
        # 误判"真实图=编造图"(实测 5000/4680 会被拦)。
        kpi = s_ind.get("kpi_matrix")
        if isinstance(kpi, dict) and kpi.get("records"):
            kpi_by_sheet[f"{s_name}::KPI" if s_name else "KPI"] = kpi
        # missing_top5 / missing_ratio_top5 跨 sheet 累加后统一取 top5
        for k, v in (s_ind.get("missing_top5") or {}).items():
            overall_missing_counter[k] = overall_missing_counter.get(k, 0) + int(v)

    # 存在 KPI 宽表时挂到顶层;多张 sheet 时按 sheet 分键,
    # 单张 sheet 时额外提供一份无前缀的 "kpi_matrix"(最常用形态,便于前端/导出直读)
    if kpi_by_sheet:
        merged["kpi_matrix_by_sheet"] = kpi_by_sheet
        if len(kpi_by_sheet) == 1:
            merged["kpi_matrix"] = next(iter(kpi_by_sheet.values()))

    # 跨 sheet 缺失 top5(按缺失数降序)
    top5 = sorted(overall_missing_counter.items(), key=lambda x: -x[1])[:5]
    merged["missing_top5"] = {k: v for k, v in top5}
    # 缺失率(%):基于顶层行数加权
    total_rows = max(1, merged["row_count"])
    merged["missing_ratio_top5"] = {
        k: round(v / total_rows * 100, 1) for k, v in top5
    }
    return merged


def _sheet_indicators_to_charts(sheet_indicators_list: list[dict]) -> list[dict]:
    """按 sheet 分别调用 indicators_to_charts,产出多 sheet 兜底图表列表。

    设计要点(2026-10 多 sheet 分项):
      - 旧实现直接对合并大表的 indicators 调 indicators_to_charts,
        会用"列并集 + NaN 填充"后的扭曲数据画图(柱图 X 轴混入不相关维度)。
      - 现在按 sheet 分别画图,每张 sheet 内部图表前追加 "[sheet 名] " 前缀
        便于区分;空 sheet 跳过(0 行无法绘图)。

    Args:
        sheet_indicators_list: 每张 sheet 的 indicators dict(由 compute_indicators 产出)。
                              元素含 "empty"=True 的会被跳过。

    Returns:
        list[dict]: 每个元素是 {type, title, option} 的 ECharts option 包装,
                    title 已加 "[sheet 名] " 前缀;按 sheet 顺序连接。
    """
    out: list[dict] = []
    for s_ind in sheet_indicators_list or []:
        if not isinstance(s_ind, dict):
            continue
        if s_ind.get("empty"):
            continue
        s_name = s_ind.get("name") or ""
        # 去掉 indicators 里的 sheet 装饰字段,避免 indicators_to_charts 处理异常
        local_ind = {k: v for k, v in s_ind.items() if k not in ("name", "empty")}
        sheet_charts = pe.indicators_to_charts(local_ind) or []
        for c in sheet_charts:
            if not isinstance(c, dict):
                continue
            title = c.get("title") or ""
            if s_name and not title.startswith(f"[{s_name}]"):
                c["title"] = f"[{s_name}] {title}"
            c["sheet"] = s_name
            out.append(c)
    return out


def _chart_title(chart: dict) -> str:
    """从 chart 字典里取出标题字符串(用于去重比对)。

    兼容 title 字段是 dict / str / 缺失 三种形态;统一返回 str(strip 后)。
    """
    if not isinstance(chart, dict):
        return ""
    raw = chart.get("title")
    if isinstance(raw, dict):
        return str(raw.get("text") or "").strip()
    if isinstance(raw, str):
        return raw.strip()
    return ""


def _first_series_type(option: dict) -> str:
    """从 ECharts option 里取出第一个 series 的 type,fallback 留空串。"""
    if not isinstance(option, dict):
        return ""
    series = option.get("series")
    if isinstance(series, list) and series:
        first = series[0]
        if isinstance(first, dict):
            return str(first.get("type") or "").strip()
    return ""


# ==================== 图表防编造校验（2026-10）====================
# 背景:
#     Prompt 里三处写了"严禁编造未给出的数字",但**后端零校验**。
#     实测 task #65(年度复盘)里 LLM 交出的图表全是编的:
#         毛利率(%)  series=[38, 36, 34]      ← 源数据真实值 30.0~33.3
#         增量贡献占比 series=[62, 21, 12, 5]  ← 源数据里根本不存在
#         当前水平 vs 季度目标 [72,65,80] vs [100,100,100]  ← 不存在
#     这些假数据会一路流进前端图表、PDF/Word/PPT 导出,对管理层是硬伤。
#
# 策略(不靠"再喊一遍不要编造",而是**用真实数据兜底**):
#     1) 把 indicators 里所有真实数值收集成一个"数值池"(允许 ±容差匹配);
#     2) 逐张图校验每个 series 的数值是否都能在池里找到出处;
#     3) 全部有出处 → 放行(图表可以保留);
#        有编造 → 用 pandas 的真实派生数据**整张替换**该图;
#        两者都拿不到 → 丢弃该图(宁可少一张,不可假一张)。
#     这样即使 LLM 编造,最终落库/导出的也一定是真数据。

# 容差常量已内联到 _value_in_pool(绝对 0.01 + 大额 0.1%),不再单独定义。

# 单条 series 至少要命中多少比例的数值才算"有出处"。
# 为什么是 0.9 而不是更宽松的 0.6:
#     实测 LLM 编造的「毛利率 [38,36,34]」里若有 2 个值能蒙混命中,
#     阈值 0.6 时 2/3=0.67 判定放行,假图就漏过去了。
#     真实图的每个值都来自指标池,命中率应接近 1.0;
#     留 10% 余量只用于容忍极个别点的四舍五入偏差。
_TRACE_MIN_HIT_RATIO = 0.9


def _collect_indicator_values(indicators: dict[str, Any]) -> list[float]:
    """递归收集 indicators 里出现的所有数值,构成"真实数值池"。

    覆盖: numeric_summary 的聚合值、kpi_matrix(_records 与 _by_sheet)的逐指标值、
    trend 的 y / 总计 / 峰值 / 环比、by_group 的 value、text_value_counts 的 value。
    刻意不区分量纲 —— 因为我们只验证"这个数字是否在数据里出现过",
    量纲是否合理由 generate_report 阶段的指标构造负责。
    """
    vals: list[float] = []

    def walk(node: Any) -> None:
        if isinstance(node, bool):
            return
        if isinstance(node, (int, float)):
            try:
                f = float(node)
                if f == f and abs(f) != float("inf"):
                    vals.append(f)
                    # 【关键】比率类指标的百分数形态也要入池。
                    #   源数据里毛利率存 0.338,但报告正文/明细表里
                    #   按业务习惯写「33.8%」——若池里只有 0.338,
                    #   溯源会把真实值 33.8 误判为"编造"(实测踩坑)。
                    #   这里把 |f| ≤ 1.5 的值同时按 ×100 收录,
                    #   让「33.8%」「11.43%」这类百分数表述都能溯源成功。
                    #   为什么限定 |f| ≤ 1.5:金额/单量不可能落在这个区间,
                    #   避免把 4680 变成 468000 混进池子造成假命中。
                    if abs(f) <= 1.5:
                        vals.append(round(f * 100, 6))
            except (TypeError, ValueError):
                pass
            return
        if isinstance(node, dict):
            # 【关键】dict 形态额外收录"派生差值"。
            #   KPI 记录里有 目标/实际,但**差值(实际-目标)** 是
            #   _build_section_digest 现算出来的(如 4680-5000=-320),
            #   它不会作为独立字段落进 indicators,导致正文里的
            #   「差值 -320 万元」被溯源误判为编造(实测踩坑)。
            #   这里对同时含"实际/actual"与"目标/target"的 dict 补算差值。
            _a = node.get("实际", node.get("actual"))
            _t = node.get("目标", node.get("target"))
            if (
                isinstance(_a, (int, float)) and not isinstance(_a, bool)
                and isinstance(_t, (int, float)) and not isinstance(_t, bool)
            ):
                try:
                    _gap = round(float(_a) - float(_t), 6)
                    vals.append(_gap)
                    # 比率类指标的差值要额外收录"百分点形态":
                    #   毛利率目标 0.35、实际 0.338 → 差值 -0.012,
                    #   但正文按业务习惯写成「-1.2 个百分点」(即 -0.012×100)。
                    #   不收录这个形态,溯源会把真实值 -1.2 误判为编造(实测踩坑)。
                    #   判据:差值绝对值 ≤ 0.05 才视为比率差值(金额差值不会这么小)。
                    if abs(_gap) <= 0.05:
                        vals.append(round(_gap * 100, 6))
                except (TypeError, ValueError):
                    pass
            for v in node.values():
                walk(v)
            return
        if isinstance(node, list):
            for v in node:
                walk(v)

    # 只走"业务数值"字段,跳过 columns/topn 这类原始行数据。
    # 注意 kpi_matrix_by_sheet 必须显式列出:多 sheet 场景下真实 KPI 值
    # 只存在于该字段里,不列会让"KPI 目标vs实际"这类真实图被误判为编造。
    for field in (
        "numeric_summary", "kpi_matrix", "kpi_matrix_by_sheet",
        "trend", "by_group", "text_value_counts",
    ):
        walk(indicators.get(field))
    return vals


def _value_in_pool(v: float, pool: list[float]) -> bool:
    """判断数值 v 是否能在真实数值池里找到出处。

    容差策略(2026-10 修订):**绝对容差为主,相对容差为辅**。
    早期只用相对容差 0.5%,在 100 附近的值容差达到 ±0.5 ——
    于是正文里的 99.99 会"命中"池中的 100.0,编造数字被放过。
    现在改为:
        1) 绝对容差 ≤ 0.01 —— 覆盖四舍五入/浮点误差(如 33.8 vs 33.80000001);
        2) 数值较大(≥100)时额外允许 0.1% 相对容差 —— 覆盖大额数值
           的展示精度差异(如 4680 vs 4680.5),但不会把 99.99 放进 100。
    这样既不误杀真实值,也不放过近似编造值。
    """
    if v == 0:
        # 0 常由"None→0"产生,单独放行以免误杀
        return True
    for p in pool:
        ap = abs(p)
        if ap < 1e-12:
            continue
        diff = abs(v - p)
        # 绝对容差:小数级别的舍入误差一律放行(同号才比)。
        # 用 1e-9 极小的 epsilon 收紧边界 —— 否则 99.99 与 100.0 的差
        # 恰好等于 0.01 会被判定为"有出处",编造数字被放过(实测踩坑)。
        if diff < 0.01 - 1e-9:
            return True
        # 大额容差:**仅对真正的金额/数量级数值**放开(≥1000),
        # 让 4680 与 4680.4 这类展示精度差异能匹配。
        # 为什么下限是 1000 而不是 100:百分数天然落在 100 附近
        # (达成率 99.9% vs 100%),若对 ≥100 放开千分之一容差,
        # 99.99 就会命中 100.0,编造数字被放过(实测踩坑)。
        if ap >= 1000 and diff < ap * 0.001:
            return True
    return False


def _series_values(chart: dict) -> list[float]:
    """抽取一张图里所有 series 的数值点(兼容 ECharts 的多种 data 形态)。"""
    out: list[float] = []
    opt = chart.get("option")
    if not isinstance(opt, dict):
        return out
    series = opt.get("series")
    if not isinstance(series, list):
        return out
    for s in series:
        if not isinstance(s, dict):
            continue
        data = s.get("data")
        if not isinstance(data, list):
            continue
        for pt in data:
            # 形态 1:数值
            if isinstance(pt, bool):
                continue
            if isinstance(pt, (int, float)):
                try:
                    f = float(pt)
                    if f == f:
                        out.append(f)
                except (TypeError, ValueError):
                    pass
                continue
            # 形态 2:饼图 {name, value}
            if isinstance(pt, dict):
                v = pt.get("value")
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    try:
                        f = float(v)
                        if f == f:
                            out.append(f)
                    except (TypeError, ValueError):
                        pass
    return out


def _collect_indicator_text(indicators: dict[str, Any]) -> set[str]:
    """收集 indicators 里出现过的文本(列名/指标名/分类名/周期名)。

    用途:判断图表 X 轴的分类标签是不是模型凭空造的。
    真实分类名("线上"/"门店"/"1月")一定能在这里找到;
    编造的("业务线A"/"销售费用")找不到。
    """
    texts: set[str] = set()

    def walk(node: Any) -> None:
        if isinstance(node, str):
            s = node.strip()
            if s:
                texts.add(s)
            return
        if isinstance(node, dict):
            for k, v in node.items():
                if isinstance(k, str):
                    texts.add(k.strip())
                walk(v)
            return
        if isinstance(node, list):
            for v in node:
                walk(v)

    for field in ("kpi_matrix", "kpi_matrix_by_sheet", "trend", "by_group",
                  "text_value_counts", "columns", "period_cols"):
        walk(indicators.get(field))
    return {t for t in texts if t}


def _chart_axis_labels(chart: dict) -> list[str]:
    """取图表 X 轴分类标签(只取 category 类型轴,value 轴无数值含义)。

    用于判断"图是否名不副实":若 X 轴写着「业务线A/业务线B/业务线C」,
    说明 LLM 造了一批源数据里根本不存在的分类。
    """
    opt = chart.get("option")
    if not isinstance(opt, dict):
        return []
    x = opt.get("xAxis")
    if not isinstance(x, dict):
        # 多系列图时 xAxis 可能是数组
        if isinstance(x, list) and x:
            for cand in x:
                if isinstance(cand, dict) and cand.get("type") == "category":
                    return [str(v) for v in (cand.get("data") or [])]
        return []
    if x.get("type") != "category":
        return []
    return [str(v) for v in (x.get("data") or [])]


def verify_chart_data(chart: dict, indicators: dict[str, Any]) -> tuple[bool, list[float]]:
    """校验单张图的数据真实性。

    返回: (是否通过, 图中全部数值点)
        - 无 series / 无数据数组(纯占位图)→ 视为通过,不做干预;
        - **有 series 但 data 全为空数组** → 不通过(空图,见下);
        - 有数值点且全部能在指标池找到出处 → 通过;
        - 有数值点但命中率不足 → 不通过(判定为编造)。

    【2026-10 重要补充:空图判定】
        实测 deepseek-flash 交出的图表长这样:
            {"title": "各业务线收入增量贡献度拆解",
             "option": {"series": [{"name": "存量复购", "data": []}, ...],
                        "xAxis": {"data": ["业务线A", "业务线B", "业务线C", "业务线D"]}}}
        series 有名字、x 轴有分类,但 **data 全是 []** —— 前端会渲染出
        一张只有坐标轴和图例、没有任何柱子的空白图。
        旧实现把"无数值点"一律当通过,于是这种空图全部漏过,
        用户看到的就是"报告里一堆空图表"。
        现在只要 series 存在却没有任何数值,就判为不通过,
        交由 _enforce_chart_data_integrity 用真实数据替换。
    """
    vals = _series_values(chart)
    if vals:
        pool = _collect_indicator_values(indicators)
        if not pool:
            # 没有任何真实数值可对照,无法判定;保守放行,
            # 避免在极简数据集上把所有图都误杀
            return True, vals
        hits = sum(1 for v in vals if _value_in_pool(v, pool))
        ratio = hits / len(vals)
        if ratio >= _TRACE_MIN_HIT_RATIO:
            return True, vals
        # 数值没出处时再看分类轴:如果 x 轴分类名在源数据里也找不到,
        # 说明这张图从头到尾都是编的(指标名/分类名全靠想象),
        # 补充记录到异常日志,便于排查"模型编造整张图"的情况。
        labels = _chart_axis_labels(chart)
        if labels:
            text_pool = _collect_indicator_text(indicators)
            hit_lbl = sum(1 for l in labels if any(l in t or t in l for t in text_pool))
            logger.debug(
                "[verify_chart_data] 图表 %s 数值无出处;分类轴命中率 %d/%d",
                chart.get("title"), hit_lbl, len(labels),
            )
        return False, vals

    # ---- 走到这里说明 vals 为空:区分"纯占位图"与"空图" ----
    opt = chart.get("option")
    series = opt.get("series") if isinstance(opt, dict) else None
    has_series = isinstance(series, list) and any(isinstance(s, dict) for s in series)
    if has_series:
        # 有 series 定义但没有任何数值点 → 空图,判为不通过
        return False, []
    # 连 series 都没有(如 build_nl_query_chart_option 的占位图)→ 放行
    return True, []


# 图表标题里的"无意义词":命中则该图不适合作为替换候选。
# 实测踩坑:KPI 表有一列叫「备注」,pandas 会把它当文本列生成
# 《[年度核心KPI] 备注 频次分布》饼图 —— 这图对报告毫无价值,
# 但因为它是 pie,曾被"结构章 → pie"的规则选去替换编造图,
# 结果"收入与增长结构"章挂了一张"备注词频"的饼图,比原编造图更糟。
_CHART_NOISE_WORDS = ("备注", "说明", "单位", "频次分布", "缺失率")


def _replace_chart_with_real_data(
    section_name: str,
    sheet_charts: list[dict],
    used_titles: set[str] | None = None,
) -> dict | None:
    """为一张编造的图挑选最合适的真实数据图替换。

    选择策略:
        1) 排除标题含"备注/说明/单位/频次/缺失率"的图(无业务信息量);
        2) 按章节语义打分(指标章→KPI 图;趋势章→折线图;结构章→饼图;
           风险/客户章→对应维度图);
        3) 已被其它章节用过的图降权,避免 6 个章节挂同一张图;
        4) 仍无匹配则取第一张可用图。

    Args:
        section_name: 章节名,用于语义匹配。
        sheet_charts:  pandas 产出的真实图表候选池。
        used_titles:   已被其他章节占用的图表标题(原地累加,避免重复挂载)。
    """
    if not sheet_charts:
        return None
    name = (section_name or "")
    used = used_titles or set()

    # 候选过滤:先剔除无意义图,再剔除已用过的
    usable = [
        c for c in sheet_charts
        if isinstance(c, dict)
        and not any(w in str(c.get("title") or "") for w in _CHART_NOISE_WORDS)
    ]
    if not usable:
        usable = [c for c in sheet_charts if isinstance(c, dict)]
    if not usable:
        return None

    def score(c: dict) -> int:
        title = str(c.get("title") or "")
        ctype = str(c.get("type") or "")
        s = 0
        # 已用过的图大幅降权(但不排除,实在没有别的可给时仍可复用)
        if title in used:
            s -= 20
        # 指标 / 概况 / 成本盈利章 → KPI 对比类柱图
        if any(k in name for k in ("指标", "KPI", "业绩", "达成", "总体", "概况", "经营", "成本", "毛利", "盈利")):
            if any(k in title for k in ("目标", "实际", "达成率")):
                s += 30
        # 趋势 / 增长章 → 折线图
        if any(k in name for k in ("趋势", "走势", "增长", "月度", "季度", "规模")):
            if ctype == "line":
                s += 30
            elif "环比" in title:
                s += 15
        # 结构 / 渠道 / 产品章 → 饼图(占比)
        if any(k in name for k in ("结构", "占比", "渠道", "产品", "区域", "板块")):
            if ctype == "pie":
                s += 30
        # 客户维度章 → 含客单价/新客户 等 KPI 的图
        if any(k in name for k in ("客户", "用户", "客群")):
            if any(k in title for k in ("目标", "实际", "达成率")):
                s += 20
        # 风险章 → 未达标相关的达成率图
        if any(k in name for k in ("风险", "问题", "归因", "异常", "不足")):
            if "达成率" in title:
                s += 20
        return s

    ranked = sorted(usable, key=score, reverse=True)
    return ranked[0] if ranked else None


def _enforce_chart_data_integrity(
    by_section: dict,
    indicators: dict[str, Any],
    sheet_charts: list[dict],
) -> tuple[dict, list[str]]:
    """校验并修复所有图表的数据真实性。

    返回: (修复后的 by_section, 修复说明列表)
        - 通过校验的图原样保留;
        - 编造的图用 pandas 真实数据替换;
        - 无真实数据可替换的图被丢弃(记录说明)。
    """
    notes: list[str] = []
    fixed: dict = {}
    # 记录已被占用的真实图表标题,避免多个章节挂同一张图
    # (实测:6 个章节全挂上《各指标目标 vs 实际》,报告观感很差)
    used_titles: set[str] = set()

    for sec, arr in by_section.items():
        if not isinstance(arr, list):
            continue
        kept: list[dict] = []
        # 章节内去重:LLM 常对同一章节交出多张同 data 的图,
        # 全部被替换后会变成"同一张真实图在同一章节里重复 3 遍"。
        sec_titles: set[str] = set()
        for c in arr:
            if not isinstance(c, dict):
                continue
            ok, vals = verify_chart_data(c, indicators)
            if ok:
                t = _chart_title(c)
                # 章节内同名图只保留第一张
                if t and t in sec_titles:
                    notes.append(
                        f"章节「{sec}」存在重复图表《{t}》,已去重只保留一张"
                    )
                    continue
                kept.append(c)
                if t:
                    sec_titles.add(t)
                used_titles.add(t)
                continue
            # 不通过 → 分两种情形给出不同的说明文案
            #   vals 为空  = 空图(series 有定义但 data 全是 [])
            #   vals 非空  = 数值编造(有数字但源数据里找不到)
            why = (
                "图表数据为空(模型只给了图表结构、未填数值)"
                if not vals
                else f"图表数值(如 {vals[:3]})在源数据中无出处"
            )
            # 编造/空图 → 尝试用真实数据替换
            repl = _replace_chart_with_real_data(sec, sheet_charts, used_titles)
            if repl is not None:
                # 替换时保留一个"数据已校正"的标记,便于前端/导出溯源
                new_chart = dict(repl)
                new_chart["data_corrected"] = True
                if vals:
                    new_chart["replaced_fake_values"] = vals[:12]
                # 同一张真实图被复用时(真实图数量少于图表数时必然会这样),
                # 加序号后缀区分,避免前端/导出出现多个同名图表。
                base_title = _chart_title(new_chart)
                if base_title in used_titles or base_title in sec_titles:
                    new_chart["title"] = f"{base_title} · {sec}"
                    new_chart["chart_shared_from"] = base_title
                    # 后缀仍撞车(同一章节内重复替换)时再追加序号
                    t2 = _chart_title(new_chart)
                    if t2 in sec_titles:
                        dup_i = 2
                        while f"{base_title} · {sec} ({dup_i})" in sec_titles:
                            dup_i += 1
                        new_chart["title"] = f"{base_title} · {sec} ({dup_i})"
                t_new = _chart_title(new_chart)
                kept.append(new_chart)
                sec_titles.add(t_new)
                used_titles.add(t_new)
                notes.append(
                    f"章节「{sec}」图表《{c.get('title') or '未命名'}》{why},"
                    f"已替换为真实数据图《{t_new}》"
                )
            else:
                notes.append(
                    f"章节「{sec}」图表《{c.get('title') or '未命名'}》{why},"
                    f"且无可用真实数据替换,已移除该图"
                )
        if kept:
            fixed[sec] = kept
        else:
            fixed[sec] = []
    return fixed, notes


# ==================== 正文数字密度校验（2026-10）====================
# 背景:
#     task #65 的正文长这样,通篇没有一个具体数字:
#       "核心结论:整体经营处于「收入低速增长、利润承压」的剪刀差区间,
#        增长质量而非增长速度是当前第一矛盾。"
#       "毛利率连续下滑,降幅大于费用率的优化幅度。"
#     这种"正确的废话"对管理层零信息量。Prompt 里已经加了数字密度硬约束,
#     但**小模型不一定遵守**(实测 deepseek-flash 在长上下文下容易退化成
#     定性套话),所以后端必须有一层可验证的兜底。
#
# 策略(不重写正文,只做检测 + 定向补数):
#     1) 逐章节统计"真实数值"的引用次数;
#     2) 达标 → 原样保留;
#        不达标 → 在章节开头**追加**一段由 indicators 派生的"关键数据摘要",
#              确保读者进入正文前就看到具体数字。
#     为什么是"追加"而不是"重写":重写要再调一次 LLM(单次 600s 超时风险),
#     且可能把 LLM 已经写对的分析改坏;追加是确定性操作、零延迟、零风险。

# 系统补入内容的标记前缀(2026-10 新增)
#
# 用途:enforce_numeric_density 会给"零数字章节"补一段真实数据摘要,
# 这段摘要的数字 100% 来自 indicators,可信;但如果把它和模型原文
# 一起送进 audit_report_numbers 做溯源判定,会稀释编造比例
# (35 个可信数字 + 23 个编造数字 → 无出处率被算成 40%,
#  看起来像"中度污染",实际是原文 100% 编造)。
# 用这个前缀标记"这段是系统写的、免检",落库前再统一剥掉。
_AUDIT_SKIP_MARKER = "\x00__SYS_DIGEST__\n"


# 正则匹配"带数字"的片段:支持千分位、小数、百分号、中文单位
_NUMBER_RE = re.compile(
    r"\d+(?:,\d{3})*(?:\.\d+)?(?:\s*(?:%|个百分点|万元|亿元|万|亿|元|单|个|人|次|天|月|年|倍))?"
)


def _count_numbers(text: str) -> int:
    """统计文本里出现的数字个数(作为"数字密度"的粗略度量)。

    故意用宽松匹配:只要出现阿拉伯数字就算 1 个,
    目的是识别"完全零数字"的段落,而不是做精确的数字溯源
    (精确溯源由图表侧的 verify_chart_data 负责)。
    """
    if not text:
        return 0
    return len(_NUMBER_RE.findall(text))


def _fmt_num(v: float) -> str:
    """数值转业务可读文本:去掉无意义的浮点尾巴。

    4680.0 → "4680";94.59 → "94.59";0.338 → "0.338"。
    报告里出现"4680.0 万元"会让管理层觉得是程序直接 dump 的,
    去掉 .0 是最基础的"像人写的"要求。
    """
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return f"{f:.4g}" if abs(f) >= 1e6 else str(round(f, 4)).rstrip("0").rstrip(".")


def _build_section_digest(section_name: str, indicators: dict[str, Any]) -> str:
    """为"零数字章节"生成一段关键数据摘要(纯 pandas 真实数据,零 LLM 参与)。

    **按章节语义选不同的数据源**,避免所有章节都贴同一段 KPI 摘要
    (实测 6 个章节全部贴同一段,既冗余又与章节主题脱节):

        指标/达成/业绩章 → KPI 明细(最全,列全量指标)
        趋势/增长/月度章 → 周期趋势(峰值/谷值/合计/环比方向)
        结构/渠道/产品章 → 分组占比 Top 项
        风险/问题/异常章 → KPI 未达标指标 + 缺失率 Top
        其它章节         → KPI 概览(仅达标/未达标数 + 极值)

    返回空串表示无可用数据(此时保持原文不变,不强加内容)。
    """
    name = section_name or ""
    parts: list[str] = []

    # 取 KPI(多 sheet 时取第一条有记录的)
    kpi = indicators.get("kpi_matrix")
    if not (isinstance(kpi, dict) and kpi.get("records")):
        for v in (indicators.get("kpi_matrix_by_sheet") or {}).values():
            if isinstance(v, dict) and v.get("records"):
                kpi = v
                break
    records = (kpi or {}).get("records") or [] if isinstance(kpi, dict) else []

    def kpi_segs(recs: list[dict], with_yoy: bool = True) -> list[str]:
        """把若干 KPI 记录转成「指标 实际/目标(差值),达成率%,同比%」文本。"""
        out: list[str] = []
        for rec in recs[:8]:
            mname = rec.get("指标")
            if not mname:
                continue
            unit = rec.get("单位")
            suffix = unit if unit and unit != "%" else ""
            seg = f"{mname}"
            act, tgt = rec.get("实际"), rec.get("目标")
            rate, yoy = rec.get("达成率%"), rec.get("同比%")
            # 单位是 % 的比率类指标(毛利率/留存率),实际值应按百分数展示,
            # 否则会出现"毛利率 实际 0.338"这种管理层读不懂的写法
            ratio_like = (unit == "%")
            act_disp = _fmt_num(act * 100) + "%" if (
                ratio_like and isinstance(act, (int, float))) else _fmt_num(act) + suffix
            tgt_disp = _fmt_num(tgt * 100) + "%" if (
                ratio_like and isinstance(tgt, (int, float))) else _fmt_num(tgt) + suffix
            if isinstance(act, (int, float)):
                seg += f" 实际 {act_disp}"
            if isinstance(tgt, (int, float)):
                seg += f"/目标 {tgt_disp}"
                if isinstance(act, (int, float)):
                    gap = act - tgt
                    gap_disp = (
                        _fmt_num(gap * 100) + "个百分点" if ratio_like
                        else _fmt_num(gap) + suffix
                    )
                    seg += f"(差值 {gap_disp})"
            if isinstance(rate, (int, float)):
                seg += f",达成率 {_fmt_num(rate)}%"
            if with_yoy and isinstance(yoy, (int, float)):
                seg += f",同比 {'+' if yoy >= 0 else ''}{_fmt_num(yoy)}%"
            out.append(seg)
        return out

    # ---------- 按章节语义分流 ----------
    # 优先级顺序很重要:先判"专属数据源"章节(趋势/结构/风险),
    # 再判 KPI 概览章。判定关键词刻意收窄,避免
    # "客户与市场维度"被"市场"命中结构章(实际该用 KPI 的客户类指标)。
    is_metric_chapter = any(k in name for k in ("指标", "KPI", "业绩", "达成", "总体", "概况", "经营", "盈利", "成本", "毛利", "价格"))
    is_trend_chapter = any(k in name for k in ("趋势", "走势", "增长", "月度", "季度", "规模"))
    is_structure_chapter = any(k in name for k in ("结构", "占比", "渠道", "产品", "区域", "板块", "品类"))
    is_risk_chapter = any(k in name for k in ("风险", "问题", "归因", "异常", "不足", "挑战"))

    # 章节主题 → 相关 KPI 指标关键词(用于只挑相关指标,避免每章都贴全量 7 个)
    _THEME_KPI_HINTS: list[tuple[tuple[str, ...], tuple[str, ...]]] = [
        (("客户", "用户", "市场", "客群"), ("客户", "客单", "留存", "复购", "新客")),
        (("成本", "毛利", "利润", "盈利", "费用"),
         ("毛利", "利润", "成本", "费用", "净利")),
        (("价格", "定价", "单价"), ("客单", "单价", "价格", "毛利")),
        (("风险", "问题", "挑战", "不足"), ("回款", "留存", "毛利", "达成")),
    ]

    # ---- 趋势/增长章:优先用周期趋势 ----
    # 注意 is_metric_chapter 里含"增长"之外的收入类词不含"收入"了,
    # "收入与增长结构"既含"增长"(趋势)又含"结构"(占比),此时优先走趋势
    if (is_trend_chapter and not is_metric_chapter) or not records:
        for key, tr in (indicators.get("trend") or {}).items():
            if not isinstance(tr, dict) or not tr.get("x"):
                continue
            total, peak, trough = tr.get("总计"), tr.get("峰值"), tr.get("谷值")
            mom = tr.get("环比变化率%") or []
            xs, ys = tr.get("x") or [], tr.get("y") or []
            if not (isinstance(peak, dict) and isinstance(trough, dict)):
                continue
            seg = f"{key}:"
            if isinstance(total, (int, float)):
                seg += f" 合计 {_fmt_num(total)}"
            if xs and ys:
                seg += (f",从 {xs[0]}({_fmt_num(ys[0])})到 {xs[-1]}({_fmt_num(ys[-1])})")
            seg += f",最高 {peak.get('周期')}({_fmt_num(peak.get('值'))})"
            seg += f",最低 {trough.get('周期')}({_fmt_num(trough.get('值'))})"
            # 环比方向:取最后两个有值的点,给出趋势方向判断
            tail = [(x, m) for x, m in zip(xs[1:], mom[1:]) if isinstance(m, (int, float))]
            if tail:
                seg += f",{tail[-1][0]}环比{tail[-1][1]:+.1f}%"
            parts.append(seg)
            break

    # ---- 结构/渠道/产品章:用分组占比 ----
    # 只在"确实是结构章"时走;风险章即使含"市场/渠道"等词也不走这条,
    # 风险章要的是"哪些指标没达标",不是"哪个渠道占比高"。
    if not parts and (is_structure_chapter and not is_risk_chapter):
        for key, items in (indicators.get("by_group") or {}).items():
            if not isinstance(items, list) or not items:
                continue
            total = sum(float(it.get("value") or 0) for it in items
                        if isinstance(it, dict))
            segs = []
            for it in items[:5]:
                if not isinstance(it, dict):
                    continue
                v = it.get("value")
                pct = f"({float(v) / total * 100:.1f}%)" if total else ""
                segs.append(f"{it.get('name')} {_fmt_num(v)}{pct}")
            if segs:
                parts.append(f"{key}:" + "、".join(segs))
            break

    # ---- 指标/概况章,或前面都没取到:用 KPI ----
    if not parts and records:
        # 先按章节主题挑相关指标;"总体经营概况"这类无明确主题的章节
        # 挑不到 → 退回全量(概览章本来就需要看全部指标)
        scoped: list[dict] = []
        for themes, kpi_keys in _THEME_KPI_HINTS:
            if any(t in name for t in themes):
                scoped = [
                    r for r in records
                    if any(k in str(r.get("指标") or "") for k in kpi_keys)
                ]
                if scoped:
                    break
        if is_risk_chapter:
            # 风险章聚焦"未达标 + 超额"两个方向,更贴合风险视角
            base = scoped or records
            missed = [r for r in base
                      if isinstance(r.get("达成率%"), (int, float)) and r["达成率%"] < 100]
            exceeded = [r for r in base
                        if isinstance(r.get("达成率%"), (int, float)) and r["达成率%"] >= 100]
            if missed:
                parts.append("未达标指标:" + ";".join(kpi_segs(missed)))
            if exceeded:
                parts.append("超额达标指标:" + ";".join(kpi_segs(exceeded)))
            if not parts:
                parts = kpi_segs(base)
        else:
            parts = kpi_segs(scoped or records)

    # ---- 仍无数据:用缺失率 ----
    if not parts:
        miss = indicators.get("missing_ratio_top5") or {}
        if miss:
            parts.append("字段缺失率:" + "、".join(
                f"{k} {v}%" for k, v in list(miss.items())[:5]))

    if not parts:
        return ""
    # 不加「【本期关键数据】」这类标签:标签本身是系统自述,
    # 且会让这段读起来像"数据罗列"而不是分析。
    # 直接以"销售额 实际 4680万元/目标 5000万元…"的形式呈现,
    # 配合图表与明细表,读者自然知道这是数据支撑。
    return ";".join(parts) + "。"


def enforce_numeric_density(
    sections: dict,
    indicators: dict[str, Any],
    min_numbers: int = 3,
) -> tuple[dict, list[str]]:
    """检测"零数字/低数字密度"章节并补入关键数据摘要。

    返回: (补齐后的 sections, 补齐说明列表)
    """
    notes: list[str] = []
    if not isinstance(sections, dict) or not sections:
        return sections, notes

    result: dict[str, Any] = {}
    for key, val in sections.items():
        # 非字符串章节(合同对比的 list / dict)不适用本规则,原样保留
        if not isinstance(val, str):
            result[key] = val
            continue
        cnt = _count_numbers(val)
        if cnt >= min_numbers:
            result[key] = val
            continue
        digest = _build_section_digest(key, indicators)
        if not digest:
            # 没有可用真实数据 → 保持原文,不做无意义的加前缀
            result[key] = val
            continue
        result[key] = _AUDIT_SKIP_MARKER + f"{digest}\n\n{val}"
        notes.append(
            f"章节「{key}」正文仅含 {cnt} 个数字,"
            f"已在开头补充本期关键数据摘要以提升信息量"
        )
    return result, notes


# ==================== 正文数字溯源（2026-10）====================
# 背景(实测):
#     图表加了防编造校验后,图表数据已经全真。但**正文**仍会编数字:
#     task#65 重跑后「三、成本与盈利能力」写的是:
#         "本期毛利率31.2%、同比-1.8个百分点,低毛利的A线收入占比提升4.4个百分点,
#          拉低综合毛利率约1.1个百分点;期间费用率18.4%…投入产出比约9:1"
#     逐个核数:16 个数字里 **13 个在源数据中找不到出处**。
#     真实毛利率是 33.8%,不是 31.2%;"A线"这个业务线根本不存在;
#     费用率拆解更是凭空编的。
#     这类"看起来极其专业但全是编的"内容对管理层危害最大 ——
#     比"收入低速增长"这种没数字的话更危险,因为它看起来可信。
#
# 策略(不重写正文,只做检测 + 标注):
#     逐章节逐数字溯源,统计"无出处数字"的占比;
#     超过阈值 → 判定该章可信度不足,在章节开头插入醒目提示,
#     并把无出处数字列进 annotations,让读者知道哪些数字要复核。
# 为什么只提示不删除:
#     删除单个数字会破坏句子语义(数字是句子的一部分),
#     重写要再调 LLM(600s 超时风险且可能改坏已写对的内容)。
#     明确告知读者"以下数字需复核"是更诚实、更安全的做法。

# 单章无出处数字占比超过该值 → 判定不可信(0.35 意味着超过 1/3 的数字是编的)
_TEXT_FABRICATED_RATIO_THRESHOLD = 0.35
# 单章至少要有这么多个数字才做判定(数字太少时统计噪声大)
_TEXT_MIN_NUMBERS_FOR_CHECK = 4

# 异常批注单条最大长度(2026-10 新增)
# 批注区的定位是"一句话点出异常",模型却常写成 100+ 字的建议书,
# 把异常批注区淹掉。超出部分截断,完整分析交给正文与明细表。
_ANNOTATION_MAX_LEN = 80


def _extract_text_numbers(text: str) -> list[str]:
    """抽取正文里的数字字符串(保留原样,便于回填展示)。"""
    if not text:
        return []
    return re.findall(r"\d+(?:,\d{3})*(?:\.\d+)?", text)


# 带符号的数字匹配(溯源用)。
# 为什么必须捕获符号:指标池里同时存在 +12.5(同比)和 -11.8(环比),
# 若只按绝对值匹配,正文里编造的「销售费用率 11.8%」会命中池中的
# -11.8(环比)而被误判为"有出处",假数字就漏过去了。
_NUMBER_SIGN_RE = re.compile(
    r"(?P<sign>[-+])?\s*(?P<num>\d+(?:,\d{3})*(?:\.\d+)?)"
)


def audit_section_numbers(
    section_name: str,
    text: str,
    pool: list[float],
) -> list[str]:
    """找出章节正文中"在源数据里找不到出处"的数字。

    判定细节:
        - 逐个数字尝试在数值池里做同量纲匹配(复用 _value_in_pool);
        - 负号不影响判定(只看绝对值);
        - 「1」「2」这类章节序号/枚举值不计入,避免误报
          (所以 min 值设为 1,纯数字 1 若池里没有也不报)。

    返回: 无出处的数字字符串列表(去重后)。
    """
    if not isinstance(text, str) or not text:
        return []
    bad: list[str] = []
    for m in _NUMBER_SIGN_RE.finditer(text):
        raw = m.group("num")
        # 带符号时按带符号值溯源(「环比 -11.8%」应匹配池中的 -11.8);
        # 不带符号时按原值溯源 —— 绝不能取绝对值,
        # 否则正文里编造的「费用率 11.8%」会命中池中的环比 -11.8 而被放过。
        signed = m.group("sign")
        try:
            f = float(raw.replace(",", ""))
        except ValueError:
            continue
        if signed:
            f = -f if signed == "-" else f
        # 序号/枚举:1~20 的整数且不带小数点,多为"第3条""两类"这类计数,
        # 跳过以免把正常表述误判成编造
        if "." not in raw and "," not in raw and f.is_integer() and 1 <= abs(f) <= 20:
            continue
        if _value_in_pool(f, pool):
            continue
        if raw not in bad:
            bad.append(raw)
    return bad


def audit_report_numbers(
    sections: dict,
    indicators: dict[str, Any],
    ratio_threshold: float = _TEXT_FABRICATED_RATIO_THRESHOLD,
) -> tuple[dict, list[str]]:
    """检测正文数字真实性,并按三档策略处置不可信章节。

    Returns:
        (处理后的 sections, 告警列表)

    【只检查模型原文,不检查系统补的摘要】
        sections 里可能带 `_AUDIT_SKIP_MARKER` 前缀 —— 那是
        enforce_numeric_density 补入的「本期关键数据摘要」,
        它的数字全部直接来自 indicators,天然可信。
        若把它一起算进分母,会**稀释**编造比例:
        摘要贡献 35 个可信数字 + 原文 23 个编造数字,
        无出处率被算成 40%(看着像"中度污染"),
        而真实情况是原文 100% 编造。
        所以这里必须先剥离系统摘要,只对模型原文做判定,
        判定完成后再把摘要拼回去。
    """
    pool = _collect_indicator_values(indicators)
    if not pool:
        return sections, []

    notes: list[str] = []
    result: dict[str, Any] = {}
    for key, val in sections.items():
        if not isinstance(val, str):
            result[key] = val
            continue

        # 剥离系统补的摘要(若有),只对模型原文做溯源判定
        prefix = ""
        body = val
        if body.startswith(_AUDIT_SKIP_MARKER):
            prefix = _AUDIT_SKIP_MARKER
            body = body[len(_AUDIT_SKIP_MARKER):]

        raw_nums = _extract_text_numbers(body)
        # 过滤掉序号类数字后的"实质数字"数量
        real_nums = [
            n for n in raw_nums
            if not ("." not in n and "," not in n and n.isdigit() and 1 <= int(n) <= 20)
        ]
        if len(real_nums) < _TEXT_MIN_NUMBERS_FOR_CHECK:
            result[key] = prefix + body
            continue
        bad = audit_section_numbers(key, body, pool)
        if not bad:
            result[key] = prefix + body
            continue
        ratio = len(bad) / len(real_nums)

        # ---- 分档处置(2026-10 修订)----
        # 【核心原则:宁可少写,不可造假;但更不能因为怕造假就不分析】
        #   实测教训(两轮踩坑):
        #     第一轮:只加"⚠️数据核对提示"→ 正文读起来仍像模像样,
        #             用户很可能直接跳过,危害更大;
        #     第二轮:重度污染整章替换成数据摘要 → 把模型写对的分析
        #             (定性判断/归因/建议)连同假数字一起扔了,
        #             报告退化成"只有数据没有分析",用户明确否决。
        #   现在统一为:剔除含编造数字的**句子**,其余分析全部保留。
        #   一句话里混进假数字就删整句,不会伤及相邻的正确分析。
        #
        # 【两档】
        #   1. 轻度(ratio ≤ 0.35):原文可用,段末加一句轻量提示(不列数字)。
        #   2. 中度/重度(ratio > 0.35):剔除含编造数字的句子,保留其余;
        #      只有在剔除后什么都不剩时,才退回纯数据摘要兜底。
        if ratio <= 0.35:
            # 轻度污染:原文基本可信,只在**段末**加一句轻量提示。
            # 【措辞调整】用户反馈"数据核对提示"这种系统自述太多余,
            # 这里压缩成最短的一句,并且不列具体数字(列了等于把编造值
            # 又抄一遍留在报告里,导出 PDF / 全文搜索时照样能看到)。
            result[key] = prefix + (
                f"{body}\n\n(注:本章 {len(bad)} 处数值请以下方明细表为准)"
            )
            notes.append(
                f"章节「{key}」正文有 {len(bad)} 个数字在源数据中无出处"
                f"(占 {round(ratio * 100)}%),已在段末标注以明细表为准"
            )
            continue

        # ---- 中度 / 重度:一律走"剔除编造句 + 保留其余" ----
        # 【2026-10 重要修订】
        #   旧逻辑:重度污染(>60%)整章替换为数据摘要。实测后果是
        #   「核心结论」章变成:
        #       "本数据集未覆盖本章所需的完整明细数据,以下为可核对的真实数据指标。
        #        (原文中引用的 9 个数值在源数据中均无对应,已不予采纳)
        #        【本期关键数据】销售额 实际 4680万元/目标 5000万元…"
        #   用户反馈:**报告不能只是罗列数据,要对数据进行分析、给方案**。
        #   整章替换等于把模型写了一半的分析(定性判断、归因、建议)
        #   连同假数字一起扔掉,只剩一堆数字 —— 这比"带假数字的分析"更糟,
        #   因为它连"分析"都没有了。
        #
        # 【新逻辑】无论比例多高,都只做一件事:剔除含编造数字的句子。
        #   模型写对的分析(结构判断、趋势解读、行动建议)只要不含假数字,
        #   就完整保留;只有引用了源数据里不存在的数值的句子才被删。
        cleaned = _drop_sentences_with_numbers(body, set(bad))

        # 剔除后仍有实质内容(有数字 且 有文字)→ 保留清洗后的分析
        if _count_numbers(cleaned) >= 2 and len(cleaned.strip()) >= 60:
            result[key] = prefix + cleaned
            notes.append(
                f"章节「{key}」正文有 {len(bad)} 个数字无出处"
                f"(占 {round(ratio * 100)}%),已剔除含这些数字的句子"
            )
            continue

        # 剔除后什么都不剩(模型整段都在编)→ 此时才退回数据摘要兜底,
        # 但**不写任何"未覆盖/不予采纳"的说教话术**,只给数据。
        result[key] = prefix + _replace_untrusted_section(key, indicators, bad)
        notes.append(
            f"章节「{key}」正文有 {len(bad)} 个数字无出处"
            f"(占 {round(ratio * 100)}%),原分析均引用了不存在的数值,"
            f"已替换为真实数据摘要"
        )
    return result, notes


# 句子切分用的分隔符:中文句号/分号/换行/问号/感叹号
_SENT_SPLIT_RE = re.compile(r"(?<=[。;;!?\n])")


def _drop_sentences_with_numbers(text: str, bad_numbers: set[str]) -> str:
    """剔除"包含指定编造数字"的句子,保留其余句子。

    为什么按句子切:数字往往内嵌在句子中间(「本期营业收入 12.86 亿元」),
    只删数字会让句子变成「本期营业收入 亿元」这种病句,
    整句删除才能保证剩余文本语义完整、可读。
    """
    if not text:
        return text
    bad_floats: set[float] = set()
    for b in bad_numbers:
        try:
            # 保留符号:「环比 -11.8%」与「费用率 11.8%」必须区分,
            # 否则剔除句子时会误伤正确的环比表述
            bad_floats.add(float(b))
        except ValueError:
            continue
    kept: list[str] = []
    for sent in _SENT_SPLIT_RE.split(text):
        if not sent.strip():
            continue
        # 与 audit_section_numbers 保持同一套带符号匹配口径
        sent_bads = [
            m.group("num") for m in _NUMBER_SIGN_RE.finditer(sent)
            if (lambda f: any(
                abs(f - bf) < 0.01 - 1e-9
                or (abs(bf) >= 1000 and abs(f - bf) < abs(bf) * 0.001)
                for bf in bad_floats
            ))(_signed_num(m))
        ]
        if sent_bads:
            # 该句含编造数字 → 整句剔除
            continue
        kept.append(sent)
    return "".join(kept).strip()


def _signed_num(m: "re.Match[str]") -> float:
    """把 _NUMBER_SIGN_RE 的匹配结果转成带符号的浮点数。"""
    try:
        f = float(m.group("num").replace(",", ""))
    except (ValueError, AttributeError):
        return 0.0
    return -f if m.group("sign") == "-" else f


def _dedupe_annotations(items: list[dict]) -> list[dict]:
    """去掉内容重复的异常批注。

    为什么要去重:批注经过限长截断 + 数值标记后,不同指标的批注
    可能变成完全相同的占位文字(如两条都变成
    "该指标存在异常…〔数值待核〕"),在异常批注区并排出现很傻。
    这里按 (level, metric, comment) 三元组去重,保留首条。
    """
    seen: set[tuple] = set()
    out: list[dict] = []
    for it in items or []:
        if not isinstance(it, dict):
            continue
        key = (
            str(it.get("level") or ""),
            str(it.get("metric") or ""),
            str(it.get("comment") or "")[:60],
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(it)
    return out


def _strip_fabricated_numbers(text: str, bad_numbers: set[str]) -> str:
    """把句中"无出处的数字"替换为占位符,**保留句子其余部分**。

    【为什么不整句删除】批注是"一句话点出异常",信息密度极高,
    例如「2月销售额250万元为全年最低点,需关注节后需求回落」——
    如果因为句子里混进一个编造数字就整句删掉,用户会看到
    "该指标存在异常,建议核对明细数据"这种毫无信息量的占位文字,
    反而不如保留原句、只把可疑数字标记出来。
    与 _drop_sentences_with_numbers 的区别:
        - _drop_sentences_with_numbers:用于**正文**,句子冗长可丢弃
        - _strip_fabricated_numbers:用于**批注**,句子精炼不可丢弃
    """
    if not text:
        return text
    bad_floats: set[float] = set()
    for b in bad_numbers:
        try:
            bad_floats.add(float(b))
        except ValueError:
            continue
    if not bad_floats:
        return text

    def _repl(m: "re.Match[str]") -> str:
        f = _signed_num(m)
        if any(
            abs(f - bf) < 0.01 - 1e-9
            or (abs(bf) >= 1000 and abs(f - bf) < abs(bf) * 0.001)
            for bf in bad_floats
        ):
            return "〔数值待核〕"
        return m.group(0)

    return _NUMBER_SIGN_RE.sub(_repl, text)


def _replace_untrusted_section(
    section_name: str,
    indicators: dict[str, Any],
    bad_numbers: list[str],
) -> str:
    """模型分析已全部因编造数字被剔除时的兜底正文:**只给数据,不加任何话术**。

    【2026-10 修订】
        旧版本会写:
            "本数据集未覆盖本章所需的完整明细数据,以下为可核对的真实数据指标。
             (原文中引用的 9 个数值在源数据中均无对应,已不予采纳)"
        用户明确要求**不要出现这种话**:
            - "未覆盖"是在替用户下结论(也许用户觉得该有,只是没提供),
              报告不该自己声明"数据不够";
            - "不予采纳"是在解释系统的内部处理过程,读者关心的是结论,
              不是系统做了什么;
            - 大段声明挤占了正文空间,反而把真正该看的分析挤没了。

        现在只输出真实数据本身,让明细表和图表去承载"数据从哪来"的信息。
    """
    digest = _build_section_digest(section_name, indicators)
    return digest or ""


# ==================== 章节正文分段（2026-10）====================
# 背景:
#     用户反馈"章节要分段,不要全部都在一段上"。实测模型的输出有两种形态:
#         合规:「结论句。\n\n数据支撑句。\n\n建议句。」→ 已有换行,无需处理
#         不合规:「结论句。数据支撑句。原因句。建议句。」→ 500 字一整块
#     Prompt 已明确要求分段(见 prompt_builder 的【分段要求】),
#     但小模型不一定遵守,这里做确定性兜底。
#
# 为什么在前端做而不是后端:
#     后端改的话,导出 PDF/Word 时也要重新处理一遍,且三处逻辑容易漂移;
#     前端一处改动,预览与页面展示同时生效,导出侧由 cleanText/渲染层
#     统一按 \n 拆段(white-space: pre-wrap 已支持),行为一致。

# 语义转折标记:出现在句首时,视为新论点的起点
_PARAGRAPH_MARKERS = (
    "其一", "其二", "其三", "其四",
    "第一", "第二", "第三", "第四",
    "首先", "其次", "再次", "最后",
    "一方面", "另一方面",
    "结论", "核心结论", "总体来看", "综上",
)


def split_paragraphs(text: str, max_len: int = 200) -> list[str]:
    """把单段长文本按语义与长度拆成多段。

    Args:
        text:     章节正文(可能已含 \\n\\n 分段,也可能是一整块)。
        max_len:  单段目标最大字数;超过则按句号强制切分。

    Returns:
        段落列表,段间以 \\n\\n 拼接后的字符串由调用方使用。
        已有换行的地方**保留原样**,不做重排,避免破坏模型的分段意图。
    """
    if not text or not isinstance(text, str):
        return []
    # 已有空行分段 → 尊重模型意图,只对超长段做二次切分
    blocks = [b for b in re.split(r"\n\s*\n", text) if b.strip()]
    if len(blocks) <= 1 and "\n" in text:
        blocks = [b for b in text.split("\n") if b.strip()]
    if len(blocks) <= 1:
        blocks = [text]

    out: list[str] = []
    for blk in blocks:
        blk = blk.strip()
        if not blk:
            continue
        if len(blk) <= max_len:
            out.append(blk)
            continue
        # 超长段:先按句号切,再把短句合并到 ~max_len 为一段
        sents = [s.strip() for s in re.split(r"(?<=[。!?])", blk) if s.strip()]
        cur = ""
        for s in sents:
            if not cur:
                cur = s
            elif len(cur) + len(s) <= max_len:
                cur += s
            else:
                out.append(cur)
                cur = s
        if cur:
            out.append(cur)

    # 段内以"其一/其二"等标记开头时,视为隐式分段,额外拆开
    final: list[str] = []
    for blk in out:
        sents = [s.strip() for s in re.split(r"(?<=[。!?])", blk) if s.strip()]
        cur = ""
        for s in sents:
            starts_marker = any(s.startswith(m) for m in _PARAGRAPH_MARKERS)
            if cur and (starts_marker or len(cur) >= max_len):
                final.append(cur)
                cur = s
            else:
                cur += s
        if cur:
            final.append(cur)
    return [p for p in final if p.strip()]


# ==================== 缺图章节 Pandas 兜底补齐 ====================
# 背景:LLM 即便被 prompt 强约束「每个章节都必须配图」,仍可能对
# 「行动计划 / 问题风险 / 附录」等看似纯文字的章节留空数组。
# 下面的工具函数用于在后端检测到这种"漏图章节"时,
# 基于已有的 indicators 派生一张与章节主题最契合的 Pandas 兜底图,
# 保证每个 sections.key 在 chart_option.by_section 里都至少有 1 张图。

# 与 LLM 配色要求保持一致的安全色板
_FALLBACK_COLORS = ["#22d3ee", "#5eead4", "#a78bfa", "#fbbf24", "#f87171", "#4ade80"]


def _fb_base_option(title: str) -> dict:
    """兜底图的公共 option 骨架(暗色主题,与前端报告页风格一致)。"""
    return {
        "title": {"text": title, "textStyle": {"color": "#5eead4"}},
        "tooltip": {"trigger": "axis"},
    }


def _fb_line(title: str, x: list, y: list) -> dict:
    opt = _fb_base_option(title)
    opt.update({
        "xAxis": {"type": "category", "data": x, "axisLabel": {"color": "#94a3b8"}},
        "yAxis": {"type": "value", "axisLabel": {"color": "#94a3b8"}},
        "series": [{
            "type": "line", "data": y, "smooth": True,
            "lineStyle": {"color": _FALLBACK_COLORS[0]},
            "areaStyle": {"color": "rgba(34, 211, 238, 0.25)"},
        }],
    })
    return {"title": title, "type": "line", "option": opt}


def _fb_pie(title: str, items: list[dict]) -> dict:
    opt = _fb_base_option(title)
    opt["tooltip"] = {"trigger": "item"}
    opt.update({
        "legend": {"textStyle": {"color": "#94a3b8"}},
        "series": [{
            "type": "pie",
            "radius": ["40%", "70%"],
            "data": [{"name": it.get("name"), "value": it.get("value")} for it in items],
            "label": {"color": "#cbd5e1"},
        }],
    })
    return {"title": title, "type": "pie", "option": opt}


def _fb_bar(title: str, items: list[tuple[str, float]], suffix: str = "") -> dict:
    """items: [(name, value), ...] 横向 bar,适合 TopN / 指标速览。"""
    names = [n for n, _ in items]
    values = [v for _, v in items]
    opt = _fb_base_option(title + (f"({suffix})" if suffix else ""))
    opt.update({
        "grid": {"left": 90, "right": 30, "top": 40, "bottom": 20},
        "xAxis": {"type": "value", "axisLabel": {"color": "#94a3b8"}},
        "yAxis": {"type": "category", "data": names, "axisLabel": {"color": "#94a3b8"}},
        "series": [{
            "type": "bar",
            "data": values,
            "itemStyle": {"color": _FALLBACK_COLORS[0]},
            "label": {"show": True, "position": "right", "color": "#cbd5e1"},
        }],
    })
    return {"title": title, "type": "bar", "option": opt}


def _fallback_chart_for_section(section_name: str, indicators: dict) -> dict | None:
    """为缺图章节挑选/生成一张 Pandas 兜底图。

    策略(按章节名关键词优先级匹配):
        1. 含「趋势/走势/同比/环比」 → 用 trend 画 line
        2. 含「占比/结构/分布/渠道」 → 用 by_group 画 pie
        3. 含「Top/排名/客户/产品」 → 用 by_group 画 horizontal bar
        4. 含「问题/风险/异常」     → 用 missing_ratio_top5 画 bar(缺失率)
        5. 其他(含「行动计划/摘要/附录/清单」) → 用 numeric_summary 画 bar(关键指标速览)
        6. 全部字段都为空 → 用 text_value_counts 画 pie(文本列频次)
        7. 还是空 → 返回 None(调用方决定要不要跳过)

    返回:
        chart dict({title, type, option}),无法生成时返回 None。
    """
    name = (section_name or "").strip()
    trend = indicators.get("trend") or {}
    by_group = indicators.get("by_group") or {}
    numeric = indicators.get("numeric_summary") or {}
    miss = indicators.get("missing_ratio_top5") or {}
    text_counts = indicators.get("text_value_counts") or {}

    # 1. 趋势类
    if any(k in name for k in ("趋势", "走势", "同比", "环比", "增长")) and trend:
        k, v = next(iter(trend.items()))
        return _fb_line(f"{name} · {k}", v.get("x") or [], v.get("y") or [])

    # 2. 占比/结构类
    if any(k in name for k in ("占比", "结构", "分布", "渠道")) and by_group:
        k, items = next(iter(by_group.items()))
        return _fb_pie(f"{name} · {k}", items)

    # 3. Top/排名类(用 by_group 画横向 bar)
    if any(k in name for k in ("Top", "TOP", "排名", "客户", "产品", "SKU")) and by_group:
        k, items = next(iter(by_group.items()))
        pairs = [(str(it.get("name")), float(it.get("value") or 0)) for it in items[:8]]
        return _fb_bar(f"{name} · {k}", pairs)

    # 4. 问题/风险类(用缺失率画 bar,最直观反映"数据问题")
    if any(k in name for k in ("问题", "风险", "异常", "缺失")) and miss:
        pairs = [(str(k), float(v)) for k, v in list(miss.items())[:5]]
        return _fb_bar(f"{name} · 字段缺失率 Top5", pairs, suffix="%")

    # 5. 默认:关键指标速览 bar(用 numeric_summary 的 sum 值)
    if numeric:
        pairs = []
        for col, agg in list(numeric.items())[:6]:
            if isinstance(agg, dict):
                v = agg.get("sum") or agg.get("mean") or 0
                try:
                    pairs.append((str(col), float(round(v, 2))))
                except (TypeError, ValueError):
                    continue
        if pairs:
            return _fb_bar(f"{name} · 关键指标速览", pairs)

    # 6. 缺失率(当 numeric 为空时的兜底)
    if miss:
        pairs = [(str(k), float(v)) for k, v in list(miss.items())[:5]]
        return _fb_bar(f"{name} · 字段缺失率 Top5", pairs, suffix="%")

    # 7. 文本列频次 pie(最后兜底)
    for k, items in text_counts.items():
        if items:
            return _fb_pie(f"{name} · {k} 频次分布", items)

    return None


def _normalize_by_section_keys(
    by_section: dict,
    sections_map: dict,
) -> dict:
    """把 by_section 中"未与任何 sections.key 匹配"的章节合并到「数据图表」。

    触发场景:
        LLM 即便被 prompt_builder 反复约束"by_section 的 key 必须严格等于
        sections 的 key",仍然常见两类脏输出:
          1) 把所有图堆在一个固定 key 里,如「数据图表」「图表」「charts」
             —— 显然 key 与模板章节(执行摘要/核心销售指标/...)完全对不上,
             这些图与章节正文脱节,前端整页只剩末尾兜底区;
          2) 章节名拼写差异,如「核心销售指标」 vs 「核心销售指标:」 /
             「核心销售指标分析」 —— 旧版走严格相等匹配会把它们推到孤儿区,
             章节下方只剩正文、不挂图。

    【2026-10 修复】改宽容匹配:
        - 复用 ai_service._loosen_section_key 的"宽松化"能力(剥前缀编号、
          去尾部冒号、切尾部破折号说明、首尾空白修剪),让 LLM 输出的「核心销售指标:」
          「核心销售指标分析」等同形态章节都能挂到模板的「核心销售指标」下;
        - 与 sections 归一化(_normalize_sections)用同一套算法,保证
          sections 与 by_section 的"宽容口径"完全一致;
        - 即使宽容匹配仍然挂不上(典型为 LLM 把所有图堆到「数据图表」),
          才推到孤儿区。

    返回:
        新的 by_section,结构仍是 `{章节名: [ChartItem, ...]}`。
        - 与 sections_map 任一 key 严格相等 or 宽容匹配的 by_section 章节
          归入对应模板章节;
        - 无法匹配的章节合并到「数据图表」;
        - 若 sections_map 为空(被 ai_service 归一化掉/失败),所有章节原样
          返回(不做"孤儿化",避免误伤)。
    """
    if not isinstance(by_section, dict) or not by_section:
        return by_section if isinstance(by_section, dict) else {}
    if not isinstance(sections_map, dict) or not sections_map:
        return dict(by_section)

    # 局部导入避免循环(ai_service 不依赖 analyzer)
    from app.services.ai_service import _loosen_section_key

    # 建立"宽松后 allowed_keys 形态 → 原 key"映射,与 _normalize_sections 策略一致
    allowed_loose_map: dict[str, str] = {}
    for ak in sections_map.keys():
        for loose in _loosen_section_key(str(ak)):
            allowed_loose_map.setdefault(loose, str(ak))

    kept: dict[str, list] = {}
    orphan: list[dict] = []
    # 用于在"宽容命中"时,把多张 key 合并到同一章节时按 title 去重
    merged_titles_per_sec: dict[str, set[str]] = {}

    for k, arr in by_section.items():
        if not isinstance(arr, list):
            continue

        # 1) 严格相等优先(精确挂载,保留章节 key 字面一致)
        if k in sections_map:
            kept[k] = list(arr)
            continue

        # 2) 宽松匹配:把 k 宽松化后反查 allowed_loose_map
        target_sec: str | None = None
        for loose in _loosen_section_key(str(k)):
            if loose in allowed_loose_map:
                target_sec = allowed_loose_map[loose]
                break

        if target_sec:
            # 归并到目标章节,按 title 去重避免重复展示
            seen = merged_titles_per_sec.setdefault(target_sec, set())
            existing = kept.get(target_sec) or []
            bucket = list(existing)
            for c in arr:
                if not isinstance(c, dict):
                    bucket.append(c)
                    continue
                t = _chart_title(c)
                if t and t in seen:
                    continue
                if t:
                    seen.add(t)
                bucket.append(c)
            kept[target_sec] = bucket
        else:
            # 3) 完全挂不上 → 进孤儿区
            orphan.extend(arr)

    if orphan:
        bucket = kept.get("数据图表", [])
        if not isinstance(bucket, list):
            bucket = []
        # 孤儿区内部同样按 title 去重
        seen_titles: set[str] = set()
        # 先把已有 bucket 的 title 收集起来(宽松匹配的章节已经收集了,
        # 但 bucket 兜底区可能与其它已有 title 重)
        for c in bucket:
            if isinstance(c, dict):
                t = _chart_title(c)
                if t:
                    seen_titles.add(t)
        dedup_orphan: list[dict] = []
        for c in orphan:
            if not isinstance(c, dict):
                dedup_orphan.append(c)
                continue
            t = _chart_title(c)
            if t and t in seen_titles:
                continue
            if t:
                seen_titles.add(t)
            dedup_orphan.append(c)
        kept["数据图表"] = bucket + dedup_orphan
    return kept


def _ensure_all_sections_have_charts(
    sections: dict,
    by_section: dict,
    indicators: dict,
) -> dict:
    """检测缺图章节并用 Pandas 兜底图补齐,返回补齐后的 by_section。

    规则:
        - 以 sections 的 key 为唯一事实来源,任何 by_section 中缺失或空数组的章节
          都会被补一张 `_fallback_chart_for_section` 生成的图;
        - by_section 中 LLM 已给出的图原样保留(不覆盖);
        - 如果 `_fallback_chart_for_section` 返回 None(indicators 完全为空),
          该章节仍保留空数组,不强行造图(避免编造数据)。
    """
    if not isinstance(sections, dict) or not sections:
        return by_section
    result = dict(by_section) if isinstance(by_section, dict) else {}
    for sec_key in sections.keys():
        current = result.get(sec_key)
        # 已存在且非空 → 跳过
        if isinstance(current, list) and current:
            continue
        # 缺图 → 用 Pandas 兜底补 1 张
        fb = _fallback_chart_for_section(sec_key, indicators)
        if fb:
            result[sec_key] = [fb]
        elif sec_key not in result:
            # 连 Pandas 兜底都生成不出来 → 保留空数组,由前端自己决定显示
            result[sec_key] = []
    return result


def run_analysis(
    db: Session,
    task: Task,
    cleaning_rules: dict[str, bool] | None = None,
    analysis_intent: str = "",
    feedback: str = "",
    selected_charts: list[dict] | None = None,
) -> Task:
    """执行一次端到端分析并更新 task 字段。

    入参:
        db: SQLAlchemy Session。
        task: 当前任务。
        cleaning_rules: 清洗规则(同原行为)。
        analysis_intent: 用户在工作台临时输入的分析重点(自由长度,已放开 500 上限)。
                         非空时:
                         1) 透传给 prompt_builder,作为第 6 段注入 Prompt;
                         2) 在报告页「异常批注」区追加一条 info 级标注,
                            让用户无需回到工作台也能看到本次的意图。
        feedback: 用户对已生成报告提的修改意见(自由长度,已放开 500 上限)。
                  非空时:仅透传给 prompt_builder 作为第 7 段注入 Prompt;
                  **不**追加 info 标注(意见本身会落库到 task.feedback_text,
                  不需要冗余展示在 annotations)。
        selected_charts: 【chart-edit-prebuilt-dashboard 新增】用户在工作台
                  图表编辑弹窗里前置预选的图表。
                  形态:[{title, option, style_key}, ...],会与 AI 自动生成图表
                  按 title 去重合并,预选优先;空列表 / None 走原逻辑。
    """
    cleaning_rules = cleaning_rules or {}

    # 【2026-10 修复·再修订】LLM 失败 → 明确失败,绝不悄悄出"假报告"。
    # 旧实现走 _run_offline_fallback,用户拿到的报告正文是「29 行 × 24 列」这种
    # 纯结构化拼接,**没有任何 AI 分析**。这与"AI 撰写年度报告"的诉求完全背离。
    # 新策略:
    #   - LLM 链路(generate_report → summarize_report → annotate_anomalies)
    #     包到外层 try 里,失败时直接写 task.status='failed' + task.error_message,
    #     让前端工作台轮询拿到明确错误。
    #   - 关闭 _run_offline_fallback 调用,保留函数定义以备回滚/单元测试。
    #   - 文件读取失败 / 模板不存在 这类**入参错误**仍走各自原分支,不影响。
    try:
        _run_analysis_core(
            db, task, cleaning_rules,
            analysis_intent=analysis_intent,
            feedback=feedback,
            selected_charts=selected_charts,
        )
    except Exception as exc:
        err = f"{type(exc).__name__}: {exc}"
        logger.warning("[run_analysis] 主链路失败,标记 task=failed: %s", err)
        if getattr(task, "status", None) == "processing":
            try:
                task.status = "failed"
                task.error_message = (
                    f"AI 生成失败:{err[:500]}。"
                    "请检查 OPENAI_API_KEY / BASE_URL / 模型名配置,"
                    "或稍后重试。"
                )
                db.commit()
                db.refresh(task)
            except Exception:
                logger.exception("[run_analysis] 写入 failed 状态时再次出错")
        # 不 raise:BackgroundTasks 已经把 task.status 落库,再 raise 只会让上层
        # 把它当成 BackgroundTasks 自身的故障记录,与本次失败原因混在一起。
    return task


def _run_offline_fallback(db: Session, task: Task, llm_err: str) -> None:
    """LLM 不可用时的本地兜底:不调 LLM,纯 pandas 产出报告骨架 + 图表。

    流程:
        1) 重读 dataset + template
        2) pandas 清洗 + 算 indicators
        3) 按模板骨架填充 sections(每节写「AI 不可达,本地指标如下」)
        4) summary 顶部注明「AI 不可用」,让用户知道本次不是 AI 报告
        5) 按 sections 派生补齐 by_section 图表
        6) 写 task.status='success' + indicators/report/chart_option 全部落库

    不写 failed 的理由:用户点了按钮没拿到东西会很失望,
    拿到一份「机器生成的本地报告 + 数据图」远比 500 报错更有用。
    """
    ds = db.get(Dataset, task.dataset_id)
    if not ds:
        task.status = "failed"
        task.error_message = "数据集不存在"
        db.commit()
        return
    tpl = db.get(ReportTemplate, task.template_id)
    if not tpl:
        task.status = "failed"
        task.error_message = "模板不存在"
        db.commit()
        return

    try:
        df, _ = pe.read_dataframe_all_sheets(ds.file_path)
        df_clean = pe.clean(df, {})
        indicators = pe.compute_indicators(df_clean)
    except Exception as exc:
        # 兜底连数据都读不出来,只能走 failed
        raise RuntimeError(f"本地兜底读数据失败: {exc}") from exc

    # 按模板骨架填 sections
    from app.core.prompt_builder import _parse_section_titles
    allowed_keys = _parse_section_titles(tpl.structure)
    sections = {
        k: (
            f"本期数据未覆盖(AI 服务暂时不可达:{llm_err[:80]}),"
            f"请查看下方「关键指标 / 数据图表」获取本地派生结果。"
        )
        for k in allowed_keys
    }

    # pandas 兜底图:每个章节都补一张(若有 indicators 派生数据)
    by_section: dict = {}
    charts = pe.indicators_to_charts(indicators)
    if charts:
        by_section["数据图表"] = charts
    by_section = _ensure_all_sections_have_charts(sections, by_section, indicators)

    summary = (
        f"⚠️ AI 服务暂时不可达({llm_err[:120]}),"
        f"本报告由本地指标引擎生成,不是 AI 撰写的内容。"
        f"建议稍后在工作台重新点击「开始 AI 分析」获取完整报告。"
    )

    task.indicators_json = dumps(indicators)
    task.report_content = dumps({
        "title": tpl.name,
        "sections": sections,
        "summary": summary,
        # 标记本次是本地兜底,前端报告页可据此展示提示条
        "offline_fallback": True,
        "offline_reason": llm_err[:300],
    })
    task.chart_option_json = dumps({
        "by_section": by_section,
        "indicators": indicators,
    })
    task.annotations_json = dumps([])
    task.status = "success"
    task.error_message = ""  # 兜底成功则清空错误
    db.commit()
    db.refresh(task)
    logger.info(
        "[run_analysis] 本地兜底成功:sections=%d, by_section keys=%s",
        len(sections), list(by_section.keys()),
    )


def _shrink_indicators_for_llm(indicators: dict[str, Any]) -> dict[str, Any]:
    """为 LLM 调用构造一份瘦身后的 indicators,大幅压缩 token。

    保留(LLM 写正文真的用得上的):
        - 基础信息:row_count / column_count
        - **kpi_matrix / kpi_matrix_by_sheet**:逐指标的目标/实际/达成率/同比,
          这是年度报告信息密度最高的结构,**必须传给 LLM**
        - 数值汇总:numeric_summary(sum/mean/max/min/median)
        - 按文本列聚合:by_group(top8 名次+值)
        - 时间趋势:trend(含总计/峰值/谷值/环比)
        - 缺失:missing_ratio_top5
        - sheet 明细:sheets

    剥离(LLM 写正文用不上,只是给前端图表用的):
        - **columns**:19 个列名的 name/type 列表,约 1.5k 字符。
          模型不需要知道"有哪些列",它只需要知道"有哪些指标、多少数据"。
          实测:带上 columns 时,模型会答"可用于经营判断的指标记录为 0 项" ——
          它把列名当成了数据行去数,反而没看到 kpi_matrix 里的真实指标。
        - **topn**:每列前 5 行原始值,年度复盘场景独占 ~10k 字符。
        - text_value_counts:裁剪到每列 5 条。
        - sheet_indicators:与已展开的顶层字段重复。

    不修改原 indicators,前端图表仍按原 indicators 渲染。
    """
    if not isinstance(indicators, dict):
        return {}
    slim: dict[str, Any] = {}

    # ---- KPI 宽表(最重要,优先保留) ----
    # 【2026-10 修复】实测漏传 kpi_matrix 导致模型完全看不到核心指标,
    # 写出"可用于经营判断的指标记录为 0 项"这种荒谬结论。
    kpi = indicators.get("kpi_matrix")
    if isinstance(kpi, dict) and kpi.get("records"):
        slim["kpi_matrix"] = {
            "dimensions": kpi.get("dimensions") or [],
            "records": kpi.get("records") or [],
        }
    elif isinstance(indicators.get("kpi_matrix_by_sheet"), dict):
        slim["kpi_matrix_by_sheet"] = indicators["kpi_matrix_by_sheet"]

    # ---- 基础信息 ----
    for key in ("row_count", "column_count"):
        if key in indicators:
            slim[key] = indicators[key]

    # ---- 周期趋势 ----
    # 【2026-10 精简】只保留 x/y 与关键统计,丢掉 raw 里其它冗余字段。
    # trend 是"月度走势"的唯一来源,必须留;但每期数值已在 y 里,
    # 不需要再重复存峰值/谷值对象(占 400+ 字符)。
    tr_out: dict[str, Any] = {}
    for key, tr in (indicators.get("trend") or {}).items():
        if not isinstance(tr, dict) or not tr.get("x"):
            continue
        item: dict[str, Any] = {"x": tr.get("x") or [], "y": tr.get("y") or []}
        mom = tr.get("环比变化率%")
        if isinstance(mom, list) and mom:
            item["环比变化率%"] = mom
        for stat in ("总计", "峰值", "谷值"):
            if stat in tr:
                item[stat] = tr[stat]
        tr_out[key] = item
        if len(tr_out) >= 2:      # 最多 2 条趋势,再多会挤爆 prompt
            break
    if tr_out:
        slim["trend"] = tr_out

    # ---- 分组构成 ----
    if indicators.get("by_group"):
        slim["by_group"] = indicators["by_group"]

    # ---- 数值列汇总 ----
    # 【2026-10 精简】**只保留 kpi_matrix 未覆盖的列**。
    # 年度复盘场景里,销售额/订单量/客单价/毛利率 等列的 sum/mean/max/min
    # 在 kpi_matrix 与 trend 里已经出现过,再重复一遍纯浪费 token
    # (实测 numeric_summary 占 1490 字符,其中约 70% 与 KPI 重复)。
    ns = indicators.get("numeric_summary") or {}
    kpi_cols = {
        (r.get("指标"), r.get("单位")) for r in (kpi.get("records") or [])
    } if isinstance(kpi, dict) else set()
    ns_slim: dict[str, Any] = {}
    for col, agg in ns.items():
        if not isinstance(agg, dict):
            continue
        # 列名形如 "sheet名::销售额(万元)",含 KPI 指标名或单位的列跳过
        base = str(col).split("::")[-1]
        if any((m and m in base) or (u and u in base) for m, u in kpi_cols):
            continue
        ns_slim[col] = {
            k: agg.get(k) for k in ("sum", "mean", "max", "min")
            if k in agg
        }
        if len(ns_slim) >= 6:
            break
    if ns_slim:
        slim["numeric_summary"] = ns_slim

    # ---- 缺失率(只留百分比口径,计数口径对写报告没用) ----
    mr = indicators.get("missing_ratio_top5")
    if mr:
        slim["missing_ratio_top5"] = mr

    return slim


def _run_analysis_core(
    db: Session,
    task: Task,
    cleaning_rules: dict[str, bool] | None,
    *,
    analysis_intent: str,
    feedback: str,
    selected_charts: list[dict] | None,
) -> None:
    """run_analysis 的核心实现:被外层 try/except 包住。

    抽出独立函数,只为让外层能精准捕获 LLM 抛出的网络/解析/超时异常,
    把 task 状态明确写成 'failed',而不是让它一直卡在 'processing'。
    """
    # 1. 读取数据
    ds: Dataset | None = db.get(Dataset, task.dataset_id)
    if not ds:
        task.status = "failed"
        task.error_message = "数据集不存在"
        db.commit()
        return task

    try:
        # 【多 sheet 分项 2026-10】改用按 sheet 独立读取:
        #   旧实调用 read_dataframe_all_sheets 把数据按"列名并集"纵向拼成一张大表,
        #   这会让跨表结构差异补出大量 NaN,LLM 拿到的指标严重失真。
        #   现在改为 read_dataframe_per_sheet 逐 sheet 读,各 sheet 独立清洗 + 算指标,
        #   最后在 indicators 里同时提供"sheet_indicators"+顶层汇总两层。
        per_sheets = pe.read_dataframe_per_sheet(ds.file_path)
    except Exception as exc:
        task.status = "failed"
        task.error_message = f"文件读取失败: {exc}"
        db.commit()
        return task

    # sheet 明细:[{name, rows, cols, indicators}, ...]
    sheet_info: list[dict] = []
    sheet_indicators_list: list[dict] = []
    cleaned_frames: list[tuple[str, dict]] = []  # (sheet_name, indicators)

    for entry in per_sheets:
        s_name = entry["name"]
        s_df = entry["df"]
        sheet_info.append({
            "name": s_name,
            "rows": int(len(s_df)),
            "cols": int(len(s_df.columns)),
        })
        if len(s_df) == 0:
            # 空 sheet:跳过清洗/指标,避免 AI 拿空数据生成"无数据"段落
            sheet_indicators_list.append({"name": s_name, "row_count": 0, "empty": True})
            continue
        try:
            s_clean = pe.clean(s_df, cleaning_rules)
        except Exception as exc:
            logger.warning("clean(sheet=%s) failed: %s", s_name, exc)
            s_clean = s_df
        s_ind = pe.compute_indicators(s_clean)
        s_ind["name"] = s_name
        sheet_indicators_list.append(s_ind)
        cleaned_frames.append((s_name, s_ind))

    # 顶层汇总:把每张 sheet 的 indicators 合并成一个 dict,key 加 sheet 前缀避免撞名
    indicators = _merge_sheet_indicators(sheet_indicators_list)
    indicators["sheets"] = sheet_info
    # 把分项指标也附在 indicators 里,前端/报告页可下钻查看
    indicators["sheet_indicators"] = sheet_indicators_list
    # 4. 调 LLM
    tpl: ReportTemplate | None = db.get(ReportTemplate, task.template_id)
    if not tpl:
        task.status = "failed"
        task.error_message = "模板不存在"
        db.commit()
        return task

    # 模板上的「描述」被作为业务背景注入 Prompt,以影响 LLM 的口径、受众与语气。
    # 没填则为空,prompt_builder 会跳过该段,行为兼容旧实现。

    # 【2026-10 性能修复】喂给 LLM 的 indicators 瘦身
    # 实测场景:年度复盘数据集 indicators JSON 序列化后 ~20k 字符 (~13000+ token),
    # 单次 chat 耗时 37s 仍未结束;三次串行累计 ~110s,直接撞 60s 单次超时。
    # 根因:pandas_engine.compute_indicators 把 topn(每列前 5 行原始值)、
    # by_group(text_value_counts 全列频次)都塞进 indicators。这些**派生原始数据**
    # 是给前端图表用的,LLM 写正文根本不需要 —— LLM 只看数字汇总就够了。
    # 修法:对传给 LLM 的 indicators 做一次瘦身,保留 numeric_summary / by_group /
    # trend 等「汇总数字」,剥离 topn 里的原始行(text_value_counts 也可裁剪)。
    # 注意:不要修改原 indicators,本地分析用的图仍按原 indicators 渲染。
    indicators_for_llm = _shrink_indicators_for_llm(indicators)

    report = ai_service.generate_report(
        category=tpl.category,
        structure=tpl.structure,
        indicators=indicators_for_llm,
        template_description=tpl.description or "",
        analysis_intent=analysis_intent,
        feedback=feedback,
        # 把 sheet 明细透传给 prompt_builder,仅在 len > 1 时插入【数据来源】段
        sheets=indicators.get("sheets") or [],
    )

    # 5. 【2026-10 改回串行】摘要 & 批注
    # 旧版曾用 ThreadPoolExecutor(max_workers=2) 并发跑 summarize+annotate,
    # 实测 2026-10-20 regenerate 失败栈追到本段:
    #   - generate_report 主调用 in_len=12132 的 reasoning 思考 + JSON 输出 ~500s;
    #   - 同进程并发起 summarize + annotate 后,**Python GIL 仍让两个 httpx 读
    #     socket 时相互让出 CPU**,反而把每路 LLM 调用都拉长到 600s+,
    #     然后撞超时,外层 _run_analysis_core 走 except → task=failed。
    # 新策略:
    #   - 串行执行 summarize → annotate,单次总耗时 ≈ 700s;
    #   - 配合 chat() 超时已放宽到 600s,主调用不再撞超时;
    #   - 串行让 LLM 服务端压力减小(避免两个 reasoning 同时打满),
    #     也避免同进程两个 httpx 在 SSL 读取时抢 GIL。
    # 性能上:相比并发的 ~600s,串行 ~700s,只多 ~100s,但稳定性好得多。
    #   如果以后发现 thinking 模型稳定 < 60s,再考虑改回并发。
    summary = ai_service.summarize_report(report, indicators)
    # 【2026-10 新增】摘要溯源复核。
    #   摘要在报告顶部最显眼,模型在这里编数字(如"营业收入 12.86 亿元")
    #   的危害远大于正文。Prompt 已约束"只能用指标里的数字",
    #   但小模型不保证遵守,这里做一次确定性复核:
    #   摘要里的数字若无出处 → 整条摘要作废,换成由 indicators 拼出的
    #   真实数据概要(宁可朴素,也不能出错)。
    _pool_for_summary = _collect_indicator_values(indicators)
    if isinstance(summary, str) and summary.strip() and _pool_for_summary:
        _s_bad = audit_section_numbers("摘要", summary, _pool_for_summary)
        # 摘要只有 3 句,容错更低:有 1 个无出处数字即判定不可信
        if _s_bad:
            logger.info(
                "[run_analysis] 摘要含无出处数字 %s,已改用真实数据概要", _s_bad
            )
            summary = _build_section_digest("总体经营概况", indicators) or (
                "本期核心数据见下方各章节明细。"
            )
    annotations = ai_service.annotate_anomalies(report, indicators)

    # 【2026-10 新增】异常批注的数值/指标名也要溯源。
    #   批注区的 value 会直接渲染成数字展示给用户,模型编造会造成
    #   "数据看板出现不存在的指标值"。这里做两层清洗:
    #     ① value 不是真实数值 → 置为 None(前端渲染成「—」)
    #     ② 清理掉 comments 里的编造句子(复用正文的剔除逻辑)
    _pool_for_anno = _collect_indicator_values(indicators)
    if isinstance(annotations, list) and _pool_for_anno:
        _cleaned_anno: list = []
        for _a in annotations:
            if not isinstance(_a, dict):
                continue
            _a = dict(_a)
            _v = _a.get("value")
            if isinstance(_v, bool):
                _a["value"] = None
            elif isinstance(_v, (int, float)):
                if not _value_in_pool(float(_v), _pool_for_anno):
                    logger.info(
                        "[run_analysis] 异常批注数值 %s 无出处,已置空", _v
                    )
                    _a["value"] = None
            _cmt = _a.get("comment")
            if isinstance(_cmt, str) and _cmt:
                _c_bad = audit_section_numbers("批注", _cmt, _pool_for_anno)
                if _c_bad:
                    # 批注信息密度高,只把编造数字标成〔数值待核〕,
                    # 不整句删除(删了只剩"该指标存在异常"这种空话)。
                    _a["comment"] = _strip_fabricated_numbers(_cmt, set(_c_bad))
                    logger.info(
                        "[run_analysis] 异常批注含 %d 个无出处数值,已标记待核",
                        len(_c_bad),
                    )
                # 【2026-10 新增】批注限长。
                #   模型常把批注写成整段建议书(如"建议在下一经营例会前明确
                #   门店单店产出整改责任人与月度恢复目标"),单条 100+ 字,
                #   异常批注区会被大段文字淹没,反而看不清重点。
                #   批注的定位是"一句话点出异常",不是替代正文;
                #   完整分析交给正文与明细表。
                if len(_a["comment"]) > _ANNOTATION_MAX_LEN:
                    _a["comment"] = (
                        _a["comment"][:_ANNOTATION_MAX_LEN].rstrip()
                        + "…(详见正文分析)"
                    )
            _cleaned_anno.append(_a)
        annotations = _dedupe_annotations(_cleaned_anno)

    # 6. 【按章节挂载 · 兜底】sections 落库前先确保有结构(可被 LLM 归一化挡掉时):
    #    若 LLM 经过 ai_service._normalize_sections 归一化后 sections 仍为空,
    #    说明 LLM 没有给出任何合模板章节清单的 key,这种情况下报告页会
    #    出现"白页"体验。这里按模板 tpl.structure 解析出的章节清单,
    #    为每个章节填一段「本期数据未覆盖,建议补充 XX 后重新生成」的占位文案,
    #    让后续 by_section 兜底按章节填 Pandas 派生图,前端报告页至少能看到
    #    完整骨架而不是白页。
    final_sections = report.get("sections") or {}
    if not final_sections and tpl and tpl.structure:
        from app.core.prompt_builder import _parse_section_titles
        allowed_keys = _parse_section_titles(tpl.structure)
        if allowed_keys:
            final_sections = {
                k: "本期数据未覆盖,建议补充该章节相关数据后重新生成。"
                for k in allowed_keys
            }
            logger.info(
                "[run_analysis] sections 归一化为空,按模板骨架(%d 章节)填充兜底占位",
                len(allowed_keys),
            )
    # 同步回 report["sections"],让下面 by_section 兜底也能用上骨架
    report["sections"] = final_sections

    # 6.5 【2026-10 再修订】不再用本地拼接覆盖 LLM 写出的章节正文。
    # 旧实现检测到"本期数据未覆盖"等占位字面量时,会用 ``_generate_local_section_text``
    # 派生一段"基于 29 行 × 24 列"的本地段落,**这与"让 AI 真写分析"的诉求完全背离**。
    # 新策略:
    #   - LLM 写出的就是最终结果。即便它写得短,也是用户授权的判断(数据稀薄、模型保守)。
    #   - LLM 真失败 → 主链路异常已写 task.status='failed',**根本走不到这里**。
    #   - 保留 placeholder_markers 常量仅作日志审计,不修改 sections 内容。
    placeholder_markers = (
        "本期数据未覆盖", "未覆盖", "建议补充", "建议补充后",
        "重新生成", "数据暂未提供", "暂无数据",
    )
    placeholder_hits = sum(
        1 for v in final_sections.values()
        if isinstance(v, str) and len(v.strip()) < 60
        and any(mk in v for mk in placeholder_markers)
    )
    if placeholder_hits:
        logger.info(
            "[run_analysis] 检测到 %d 个章节使用了占位文案,保留 LLM 原样输出"
            "(不再用本地拼接覆盖,避免把 AI 报告变成 pandas 拼接)",
            placeholder_hits,
        )

    # 6.6 【2026-10 新增】正文数字密度兜底。
    #     实测 task #65:LLM 正文全是"收入低速增长、利润承压的剪刀差区间"
    #     这类零数字套话。Prompt 已加硬约束,但小模型在长上下文下不一定遵守,
    #     所以后端做一次确定性检测:数字不足的章节在开头补一段真实数据摘要。
    #     放在图表兜底之前,因为摘要内容只依赖 indicators,与图表无关。
    final_sections, density_notes = enforce_numeric_density(final_sections, indicators)
    if density_notes:
        logger.info(
            "[run_analysis] 数字密度兜底补齐 %d 个章节:\n%s",
            len(density_notes), "\n".join(density_notes),
        )

    # 6.65 【2026-10 新增】正文数字溯源。
    #     图表已经防编造了,但正文仍会编 —— 实测「成本与盈利能力」章
    #     16 个数字里 13 个查无出处(毛利率写成 31.2%,真实是 33.8%;
    #     还虚构了"A线"业务线和一整套费用率拆解)。
    #     这种"看起来专业但全是编的"内容比没数字的话危害更大。
    #     这里逐章溯源,无出处数字占比超阈值的章节加核对提示。
    #     必须放在密度补齐之后:补齐摘要里的数字全部来自 indicators,
    #     不会误判;若放在之前,摘要本身尚未写入。
    final_sections, audit_notes = audit_report_numbers(final_sections, indicators)
    if audit_notes:
        logger.info(
            "[run_analysis] 正文数字溯源告警:\n%s", "\n".join(audit_notes)
        )

    # 6.7 【2026-10 新增】章节明细数据表。
    #     正文有数字之后,管理层还需要能**核对**这些数字的来源,
    #     所以为每个章节挂一张与该章节主题匹配的明细表(KPI 明细 /
    #     周期走势 / 结构占比 / 未达标清单)。前端渲染成表格,
    #     导出 PDF/Word 时也会落成真表格。
    detail_tables = pe.build_section_detail_tables(indicators, list(final_sections.keys()))
    if detail_tables:
        logger.info("[run_analysis] 已生成章节明细表:%s", list(detail_tables.keys()))

    # 7. 兜底图表:分两层保障「每个 sections.key 都至少有 1 张图」
    #    第一层(by_section 完全为空):LLM 完全没给图,走 Pandas 兜底,
    #        先塞到「数据图表」区,再按 sections 逐章节派生补齐;
    #    第二层(by_section 非空但漏章节):LLM 只给部分章节配图,
    #        用 `_ensure_all_sections_have_charts` 为缺图章节补 Pandas 兜底图。
    #    【chart-edit-prebuilt-dashboard 新增】兜底完成后,与用户在弹窗里
    #    预选的 selected_charts 按 title 去重合并,预选优先(避免标题撞车)。
    co = report.get("chart_option") or {}
    by_section = co.get("by_section") if isinstance(co, dict) else None
    sections_map = final_sections  # 用上面的兜底骨架 sections

    # 【按章节挂载 · 关键】归一化 by_section 的 key:把与 sections 任一 key
    # 都不匹配的章节(如 LLM 把所有图都堆到「数据图表」「图表」「charts」等)
    # 强制合并到「数据图表」孤儿区,避免前端报告页出现"图堆在某个诡异 key 里"
    # 与章节正文脱节。归一化后再走缺图补齐 + 兜底。
    if isinstance(by_section, dict):
        by_section = _normalize_by_section_keys(by_section, sections_map)

    # 【2026-10 新增】先算出"pandas 真实图表池",供后续两处复用:
    #   ① by_section 完全为空时的兜底图;
    #   ② 图表防编造校验时,给 LLM 编造的图提供真实替代品。
    # 放在这里统一算一次,避免重复调用 indicators_to_charts。
    real_sheet_charts = _sheet_indicators_to_charts(sheet_indicators_list)

    if not by_section:
        # 【多 sheet 分项 2026-10】按 sheet 分别生成兜底图表,
        # 把每张 sheet 的图表放到独立区(数据图表::sheet 名),便于下钻。
        base_by_section: dict = {"数据图表": list(real_sheet_charts)}
        # 即使走「数据图表」兜底,也尽力为每个章节派生一张图,
        # 避免新报告页上某些章节只剩纯文字
        base_by_section = _ensure_all_sections_have_charts(
            sections_map, base_by_section, indicators,
        )
        report["chart_option"] = {
            "by_section": base_by_section,
            "indicators": indicators,
        }
    else:
        # by_section 非空:先补 indicators 字段(供导出/前端兜底使用)
        if isinstance(co, dict) and not co.get("indicators"):
            co["indicators"] = indicators
        # 再为缺图章节(LLM 漏配的章节)补 Pandas 兜底图,
        # 保证每个 sections.key 在 by_section 里都至少有 1 张图
        co["by_section"] = _ensure_all_sections_have_charts(
            sections_map, by_section, indicators,
        )
        report["chart_option"] = co

    # ==== 【2026-10 新增】挂载章节明细数据表 ====
    # 放在 chart_option.detail_tables 而非 report_content.sections 里,
    # 是为了让"结构化数据"与"AI 文本"两条线保持分离:
    #   - sections  仍是纯文本(导出 PDF/Word 时按段落排版);
    #   - detail_tables 是结构化表格数据(导出时落成真表格,前端渲染成 <table>)。
    # 混写进 sections 会污染 LLM 的输出契约(sections 只该是字符串),
    # 且会让 _normalize_sections 的 str() 化把表格变成 JSON 字符串。
    if detail_tables:
        _co_tbl = report.get("chart_option")
        if not isinstance(_co_tbl, dict):
            _co_tbl = {}
            report["chart_option"] = _co_tbl
        _co_tbl["detail_tables"] = detail_tables

    # ==== 【2026-10 新增】图表防编造校验 ====
    # LLM 交出的图表 series 数值若在 indicators 里找不到出处,一律判为编造,
    # 用 pandas 真实数据替换(或移除)。必须在用户预选图表合并**之前**做,
    # 这样"用户预选的图"是用户自己从真实图表里挑的,不需要也不应该被改。
    co_after = report.get("chart_option") or {}
    by_sec_after = co_after.get("by_section") if isinstance(co_after, dict) else None
    if isinstance(by_sec_after, dict) and by_sec_after:
        fixed_by_section, chart_notes = _enforce_chart_data_integrity(
            by_sec_after, indicators, real_sheet_charts,
        )
        if chart_notes:
            logger.info(
                "[run_analysis] 图表防编造校验介入 %d 处:\n%s",
                len(chart_notes), "\n".join(chart_notes),
            )
            # 把修复说明挂到 chart_option,前端可展示"数据已校正"提示,
            # 同时写入 annotations 让管理层在报告页就能看到提醒
            co_after["by_section"] = fixed_by_section
            co_after["data_integrity_notes"] = chart_notes
            report["chart_option"] = co_after

    # === 【chart-edit-prebuilt-dashboard】合并用户预选图表 ===
    if selected_charts:
        # 把所有 by_section 内的图先打成 (title, chart) 索引,合并后再按原结构写回
        co_now = report.get("chart_option") or {}
        by_sec = co_now.get("by_section") if isinstance(co_now, dict) else None
        if not isinstance(by_sec, dict):
            by_sec = {"数据图表": []}
        # 合并:先按 title 收集已存在的图,再决定每个 by_section 内的最终列表
        existing_titles: set[str] = set()
        for arr in by_sec.values():
            if isinstance(arr, list):
                for c in arr:
                    if isinstance(c, dict):
                        t = _chart_title(c)
                        if t:
                            existing_titles.add(t)
        # 用户预选的图:跳过分工标题与现有标题都冲突的,但预选优先(走"覆盖")
        merged: dict[str, list] = {}
        for sec, arr in by_sec.items():
            if isinstance(arr, list):
                merged[sec] = list(arr)
        orphan_selected: list[dict] = []
        for sc in selected_charts:
            if not isinstance(sc, dict):
                continue
            t = (sc.get("title") or "").strip()
            if not t:
                continue
            chart_obj = {
                "title": t,
                "type": sc.get("type") or _first_series_type(sc.get("option") or {}),
                "option": sc.get("option") or {},
                "style_key": sc.get("style_key") or "",
                # 标记来源:前端预选
                "source": "user_selected",
            }
            # 把图表放进与已存在 title 相同章节(若有"数据图表"兜底区,优先放这里)
            placed = False
            for sec, arr in merged.items():
                if any(_chart_title(c) == t for c in arr):
                    # 同标题:覆盖旧图(预选优先)
                    merged[sec] = [chart_obj if _chart_title(c) == t else c for c in arr]
                    placed = True
                    break
            if not placed:
                # 没撞标题:放进"数据图表"章节
                if "数据图表" in merged:
                    merged["数据图表"].append(chart_obj)
                else:
                    merged["数据图表"] = [chart_obj]
        # 若用户预选全部撞标题(都被合并),也保留一份在"用户预选"section 方便追溯
        if orphan_selected:
            merged.setdefault("用户预选", []).extend(orphan_selected)
        co_now["by_section"] = merged
        report["chart_option"] = co_now
        # 同时把 selected_charts 落到 task.selected_charts_json,
        # 留痕便于报告页/导出端再次读取(虽然 by_section 已合并,但保留原 list 也方便排查)
        try:
            task.selected_charts_json = dumps(selected_charts)
        except Exception:
            task.selected_charts_json = None

    # 7. 把用户在「开始 AI 分析」前临时输入的"分析重点"作为 info 级标注,
    #    追加到 annotations 的**前端**(index=0),让用户在报告页「异常批注」
    #    区第一眼看到自己刚才输入的重点,无需回到工作台。
    #    仅当 analysis_intent 非空 + feedback 为空时(feedback 路径独立,意见
    #    已在 feedback_text 落库,不在 annotations 重复展示)。
    intent_norm = (analysis_intent or "").strip()
    if intent_norm and not (feedback or "").strip():
        intent_note = {
            "metric": "分析重点",
            "level": "info",
            "comment": f"用户本次分析重点:{intent_norm[:200]}",
        }
        if isinstance(annotations, list):
            annotations.insert(0, intent_note)
        else:
            annotations = [intent_note]

    # 8. 【2026-10 新增】把图表防编造校验的修复说明写进异常批注。
    #    为什么要有:LLM 编造的图被替换/移除后,用户看到的图与 AI 原本
    #    "想画的图"可能不一致。若不告知,用户会怀疑"是不是数据出问题了"。
    #    用 warn 级别提示,内容说明"原图数值无出处 → 已换成真实数据图"。
    _co_final = report.get("chart_option") or {}
    _integrity_notes = (
        _co_final.get("data_integrity_notes")
        if isinstance(_co_final, dict) else None
    )
    if _integrity_notes:
        _joined = ";".join(str(n) for n in _integrity_notes)[:500]
        _integrity_note = {
            "metric": "图表数据校验",
            "level": "warn",
            "comment": f"检测到 {len(_integrity_notes)} 处图表数值在源数据中无出处,"
                       f"已用真实数据替换或移除:{_joined}",
        }
        if isinstance(annotations, list):
            annotations.append(_integrity_note)
        else:
            annotations = [_integrity_note]

    # 9. 【2026-10 新增】把"正文数字密度补齐"记录写进异常批注(info 级)。
    #    用 info 而非 warn:这不是错误,而是系统主动提升了报告的信息量,
    #    让用户知道章节开头那段数字是系统补的、来源是真实数据。
    if density_notes:
        _dens_note = {
            "metric": "数字密度增强",
            "level": "info",
            "comment": f"{len(density_notes)} 个章节的正文数字密度不足,"
                       f"已在章节开头补充本期真实关键数据摘要",
        }
        if isinstance(annotations, list):
            annotations.append(_dens_note)
        else:
            annotations = [_dens_note]

    # 10. 【2026-10 新增】正文数字溯源告警(danger 级,排最前)。
    #     这是本次修复里最需要让用户知道的风险:AI 在正文里写了
    #     看起来专业、但源数据里根本不存在的数字。必须在报告顶部警示,
    #     且用 danger 级(与 warn/info 区分),避免被当成普通提示忽略。
    if audit_notes:
        _audit_note = {
            "metric": "正文数字核对",
            "level": "danger",
            "comment": f"检测到 {len(audit_notes)} 个章节的正文引用了源数据中不存在的数字,"
                       f"已按可信度自动处置(剔除问题句子或替换为真实数据):"
                       f"{';'.join(audit_notes)[:400]}",
        }
        if isinstance(annotations, list):
            annotations.insert(0, _audit_note)
        else:
            annotations = [_audit_note]

    # 11. 【2026-10 新增】剥离系统摘要标记前缀(落库前最后一步)。
    #     _AUDIT_SKIP_MARKER 只用于让 audit_report_numbers 识别
    #     "这段是系统写的、免检",它绝不能出现在用户看到的报告里。
    if isinstance(final_sections, dict):
        for _k, _v in final_sections.items():
            if isinstance(_v, str) and _v.startswith(_AUDIT_SKIP_MARKER):
                final_sections[_k] = _v[len(_AUDIT_SKIP_MARKER):].lstrip("\n")

    task.indicators_json = dumps(indicators)
    task.report_content = dumps({
        "title": report.get("title", tpl.name),
        "sections": final_sections,
        "summary": summary,
    })
    task.chart_option_json = dumps(report.get("chart_option", {}))
    task.annotations_json = dumps(annotations)
    task.status = "success"
    db.commit()
    db.refresh(task)
    # 注意:core 函数不返回 task,统一由外层 run_analysis return,
    # 避免"内层 return" 屏蔽外层的 try/except。
    return None
