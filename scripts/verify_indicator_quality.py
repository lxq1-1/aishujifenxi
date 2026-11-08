"""指标引擎质量验证脚本。

用途:
    验证 2026-10「年度报告只有文字」修复后,pandas_engine 喂给 LLM 的
    指标是否**干净的真实业务数据**。逐张 sheet 打印关键指标并做断言,
    任何一个断言失败即 exit 1,可直接用于回归卡点。

背景(修复前的三类问题):
    1. 合计行混入聚合 → sum 翻倍(9360 vs 真实 4680)
    2. KPI 宽表被跨量纲求和 → sum=67052(5000万元+60000单+850元+0.35比率)
    3. 月份列未识别为时间轴 → trend 为空,LLM 无走势素材

用法:
    cd /Users/xiahua/trea/数据/backend
    python3 ../scripts/verify_indicator_quality.py
    python3 ../scripts/verify_indicator_quality.py --file <xlsx路径>   # 自定义数据集
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# 允许从 backend 目录直接运行本脚本
BACKEND = Path(__file__).resolve().parent.parent / "backend"
sys.path.insert(0, str(BACKEND))

from app.services import pandas_engine as pe  # noqa: E402


# 断言失败累计数;>0 时进程退出码为 1
_FAILURES: list[str] = []


def check(cond: bool, label: str, detail: str = "") -> None:
    """断言并打印 PASS/FAIL。失败记录到 _FAILURES,不影响后续检查。"""
    mark = "PASS" if cond else "FAIL"
    line = f"  [{mark}] {label}"
    if detail:
        line += f" — {detail}"
    print(line)
    if not cond:
        _FAILURES.append(label)


def section(title: str) -> None:
    print()
    print("=" * 78)
    print(title)
    print("=" * 78)


def verify_sheet(name: str, ind: dict) -> None:
    """逐张 sheet 打印指标并断言。"""
    print()
    print(f"### sheet: {name}")

    removed = ind.get("total_rows_removed")
    print(f"  行数: {ind.get('row_count')}"
          + (f"(原始 {ind.get('raw_row_count')},已剔除汇总行 {removed})"
             if removed else ""))
    if ind.get("period_cols"):
        print(f"  周期列: {ind['period_cols']}")

    # ---- 检查 1:汇总行不得进入任何 by_group / text_value_counts ----
    for field in ("by_group", "text_value_counts"):
        for key, items in (ind.get(field) or {}).items():
            if not isinstance(items, list):
                continue
            bad = [
                str(it.get("name"))
                for it in items
                if isinstance(it, dict) and pe.is_total_label(it.get("name"))
            ]
            check(not bad, f"{field}[{key}] 无汇总伪类别", f"发现 {bad}" if bad else "")

    # ---- 检查 2:KPI 宽表应产出 kpi_matrix,且不再进 numeric_summary ----
    kpi = ind.get("kpi_matrix")
    if kpi:
        print(f"  KPI 维度: {kpi.get('dimensions')}")
        for rec in (kpi.get("records") or [])[:8]:
            print(f"    · {rec}")
        check(len(kpi.get("records") or []) >= 2, "KPI 记录数 ≥ 2",
              f"实际 {len(kpi.get('records') or [])} 条")
        # KPI 角色列不应出现在 numeric_summary(那会导致跨量纲求和)
        role_cols = set((kpi.get("role_cols") or {}).values())
        leaked = role_cols & set((ind.get("numeric_summary") or {}).keys())
        check(not leaked, "KPI 角色列未混入 numeric_summary",
              f"泄漏列 {sorted(leaked)}" if leaked else "")
        # 记录的字段必须有真实数值,不能是 None
        no_value = [
            r.get("指标") for r in kpi.get("records") or []
            if r.get("实际") is None and r.get("目标") is None
        ]
        check(not no_value, "KPI 每条记录都带真实数值", f"缺数值指标 {no_value}" if no_value else "")

    # ---- 检查 3:周期趋势必须有序且不含汇总行 ----
    for key, tr in (ind.get("trend") or {}).items():
        if not isinstance(tr, dict) or "环比变化率%" not in tr:
            continue
        xs = tr.get("x") or []
        ys = tr.get("y") or []
        print(f"  趋势[{key}]:")
        print(f"    x   = {xs}")
        print(f"    y   = {ys}")
        print(f"    总计={tr.get('总计')} 均值={tr.get('均值')}")
        print(f"    峰值={tr.get('峰值')} 谷值={tr.get('谷值')}")
        print(f"    环比变化率% = {tr.get('环比变化率%')}")
        bad_x = [x for x in xs if pe.is_total_label(x)]
        check(not bad_x, f"趋势[{key}] X 轴无汇总行", f"发现 {bad_x}" if bad_x else "")
        # 顺序校验:解析出的周期序号应单调递增
        ords = [pe.period_sort_key(x)[1] for x in xs]
        check(ords == sorted(ords), f"趋势[{key}] X 轴按真实周期排序",
              f"实际顺序 {xs}")
        check(len(xs) == len(ys), f"趋势[{key}] x/y 长度一致")
        # 总计应等于各期之和
        if ys:
            check(abs(sum(ys) - (tr.get("总计") or 0)) < 0.01,
                  f"趋势[{key}] 总计 = 各期之和",
                  f"{tr.get('总计')} vs {round(sum(ys), 2)}")

    # ---- 检查 4:numeric_summary 不应出现跨量纲的极端值 ----
    ns = ind.get("numeric_summary") or {}
    for col, agg in ns.items():
        if not isinstance(agg, dict):
            continue
        s = agg.get("sum")
        # 订单量/金额这类正常业务值可能到几万,但 1e10 以上基本是量纲混算
        if isinstance(s, (int, float)):
            check(abs(s) < 1e10, f"numeric_summary[{col}] 无跨量纲异常",
                  f"sum={s}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="自定义 Excel/CSV 路径;默认用年度复盘数据集")
    args = ap.parse_args()

    if args.file:
        path = Path(args.file)
    else:
        path = BACKEND / "backend/uploads/20261019_134239_年度复盘_统计指标数据.xlsx"

    if not path.exists():
        print(f"[ERROR] 数据集不存在: {path}")
        return 2

    section(f"验证数据集: {path.name}")
    per_sheets = pe.read_dataframe_per_sheet(path)
    print(f"共 {len(per_sheets)} 张 sheet")

    for entry in per_sheets:
        s_name = entry["name"]
        s_df = entry["df"]
        if len(s_df) == 0:
            continue
        ind = pe.compute_indicators(pe.clean(s_df, {}))
        verify_sheet(s_name, ind)

        # 图表产出验证:确认 KPI/周期图真的生成了
        charts = pe.indicators_to_charts(ind)
        if charts:
            print(f"  产出图表 {len(charts)} 张:")
            for c in charts:
                opt = c.get("option") or {}
                series = opt.get("series") or []
                names = [s.get("name", s.get("type")) for s in series if isinstance(s, dict)]
                print(f"    · [{c.get('type')}] {c.get('title')} series={names}")
        check(bool(charts), f"[{s_name}] 至少产出 1 张图")

    section("验证结果")
    if _FAILURES:
        print(f"共 {len(_FAILURES)} 项断言失败:")
        for f in _FAILURES:
            print(f"  - {f}")
        return 1
    print("全部断言通过:指标为干净的真实业务数据。")
    return 0


if __name__ == "__main__":
    sys.exit(main())