"""图表防编造校验回归测试。

用途:
    验证 analyzer.verify_chart_data 能准确区分"真实图"与"LLM 编造的图"。
    用例直接取自 task #65(年度复盘)里 LLM 实际编造的数据,确保
    历史上真实出现过的编造形态都能被拦住。

为什么必须有这个测试(两次实测踩坑):
    1. `_merge_sheet_indicators` 未合并 kpi_matrix → 真实 KPI 值不在数值池
       → 「目标vs实际」被误判为编造(假阴性)。
    2. 数值池允许 ×100 换算 → 指标里的 0.38 被当成 38%,让编造的
       「毛利率 [38,36,34]」蒙混过关(假阳性)。
    这两个问题都不是"看起来能跑"能发现的,必须用真实数据断言。

用法:
    cd /Users/xiahua/trea/数据/backend
    python3 ../scripts/verify_chart_integrity.py
"""
from __future__ import annotations

import sys
from pathlib import Path

BACKEND = Path(__file__).resolve().parent.parent / "backend"
sys.path.insert(0, str(BACKEND))

from app.services import pandas_engine as pe  # noqa: E402
from app.services.analyzer import (  # noqa: E402
    _merge_sheet_indicators,
    verify_chart_data,
)

DATA = BACKEND / "backend/uploads/20261019_134239_年度复盘_统计指标数据.xlsx"

_FAILS: list[str] = []


def build_top_indicators() -> dict:
    """按 analyzer 主链路的方式算出顶层合并指标。"""
    per_sheets = pe.read_dataframe_per_sheet(DATA)
    s_inds = []
    for e in per_sheets:
        df = e["df"]
        if len(df) == 0:
            continue
        ind = pe.compute_indicators(pe.clean(df, {}))
        ind["name"] = e["name"]
        s_inds.append(ind)
    return _merge_sheet_indicators(s_inds)


def case(title: str, option: dict) -> dict:
    return {"title": title, "type": "bar", "option": option}


# (用例名, chart, 期望是否放行)
CASES: list[tuple[str, dict, bool]] = [
    # ---------- 必须拦截:task #65 里 LLM 真实编造过的数据 ----------
    ("编造·毛利率折线[38,36,34]", case(
        "毛利率(%)",
        {"series": [{"type": "line", "data": [38, 36, 34]}]},
    ), False),
    ("编造·增量贡献占比[62,21,12,5]", case(
        "增量贡献占比",
        {"series": [{"type": "bar", "data": [62, 21, 12, 5]}]},
    ), False),
    ("编造·收入集中度饼图[48,52]", case(
        "收入集中度",
        {"series": [{"type": "pie", "data": [
            {"name": "头部客户收入占比", "value": 48},
            {"name": "其他客户收入占比", "value": 52},
        ]}]},
    ), False),
    ("编造·当前水平vs季度目标[72,65,80]", case(
        "当前水平",
        {"series": [
            {"name": "当前水平", "type": "bar", "data": [72, 65, 80]},
            {"name": "季度目标", "type": "bar", "data": [100, 100, 100]},
        ]},
    ), False),

    # ---------- 必须放行:pandas 真实产出的数据 ----------
    ("真实·月度销售额折线", case(
        "月度销售额",
        {"series": [{"type": "line", "data": [
            280, 250, 360, 380, 410, 430, 400, 420, 410, 460, 440, 440]}]},
    ), True),
    ("真实·KPI目标vs实际", case(
        "目标vs实际",
        {"series": [
            {"name": "目标", "type": "bar",
             "data": [5000, 60000, 850, 0.35, 1200, 0.7, 0.95]},
            {"name": "实际", "type": "bar",
             "data": [4680, 58200, 804, 0.338, 1350, 0.66, 0.91]},
        ]},
    ), True),
    ("真实·渠道占比饼图", case(
        "渠道占比",
        {"series": [{"type": "pie", "data": [
            {"name": "线上", "value": 1950},
            {"name": "门店", "value": 1200},
            {"name": "代理商", "value": 900},
            {"name": "分销", "value": 630},
        ]}]},
    ), True),
    ("真实·KPI达成率%(百分数形态)", case(
        "达成率",
        {"series": [{"type": "bar",
                     "data": [93.6, 97.0, 94.59, 96.57, 112.5, 94.29, 95.79]}]},
    ), True),
    ("真实·KPI同比%(百分数形态)", case(
        "同比",
        {"series": [{"type": "bar",
                     "data": [11.43, 7.78, 3.34, 4.0, 35.0, 6.45, 3.41]}]},
    ), True),
    ("真实·月度环比变化率(含 null 首项)", case(
        "环比",
        {"series": [{"type": "bar", "data": [
            None, -10.7, 44.0, 5.6, 7.9, 4.9, -7.0, 5.0, -2.4, 12.2, -4.3, 0.0]}]},
    ), True),

    # ---------- 必须拦截:空图(实测 deepseek-flash 的典型失败形态) ----------
    # 模型只给了图表骨架(x 轴有编造的分类名、series 有名字),
    # 但 data 全是 [] —— 前端会渲染出只有坐标轴的空白图。
    ("空图·业务线增量贡献度(data 全空)", {
        "title": "各业务线收入增量贡献度拆解", "type": "bar",
        "option": {
            "xAxis": {"type": "category",
                      "data": ["业务线A", "业务线B", "业务线C", "业务线D"]},
            "series": [
                {"name": "存量复购", "type": "bar", "data": []},
                {"name": "新客开拓", "type": "bar", "data": []},
                {"name": "价格调整", "type": "bar", "data": []},
            ],
        },
    }, False),
    ("空图·单 series data 为空", {
        "title": "空图", "type": "line",
        "option": {"series": [{"type": "line", "data": []}]},
    }, False),

    # ---------- 边界:完全不构成图的占位结构不做干预 ----------
    # 注意:"有 series 但 data 为空"已归入上面的"空图"用例(应拦截);
    # 这里只验证"连 series 数组都没有"的纯占位图仍放行,
    # 以兼容 pandas_engine.build_nl_query_chart_option 的占位图
    # (那种图 series:[] 是故意留空的,前端靠它显示标题)。
    ("边界·无 series 键", {"title": "纯占位图"}, True),
    ("边界·series 为空数组", case("占位图", {"series": []}), True),
]


def main() -> int:
    if not DATA.exists():
        print(f"[ERROR] 数据集不存在: {DATA}")
        return 2

    print("=" * 78)
    print("图表防编造校验回归测试")
    print("=" * 78)
    top = build_top_indicators()
    print(f"顶层指标字段: {sorted(top.keys())}")
    print()

    for label, chart, expect_pass in CASES:
        got_pass, vals = verify_chart_data(chart, top)
        ok = (got_pass == expect_pass)
        mark = "PASS" if ok else "FAIL"
        verdict = "放行" if got_pass else "拦截"
        expect = "放行" if expect_pass else "拦截"
        print(f"  [{mark}] {label}")
        print(f"         期望={expect} 实际={verdict}  数值点={vals[:6]}")
        if not ok:
            _FAILS.append(label)

    print()
    print("=" * 78)
    if _FAILS:
        print(f"{len(_FAILS)} 项失败:")
        for f in _FAILS:
            print(f"  - {f}")
        return 1
    print(f"全部 {len(CASES)} 项用例通过:编造图被拦、真实图放行。")
    return 0


if __name__ == "__main__":
    sys.exit(main())