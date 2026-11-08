"""端到端离线演练:不调 LLM,验证四层修复能协同工作。

覆盖用户反馈的原始问题「生成的年度报告不应该只有文字,内容应该还是数据吧」:
    第 1 层 指标引擎     → 合计行剔除 / KPI 宽表 / 周期趋势
    第 3 层 数字密度     → 零数字章节补真实数据摘要
    第 4 层 明细数据表   → 每章节挂结构化表格
    图表防编造           → 用 task #65 的真实编造图验证会被拦

为什么是"离线演练"而不是真跑一次分析:
    真跑要调 3 次 LLM(generate + summarize + annotate),实测单次 600s+,
    而本脚本能在 1 秒内验证**确定性的部分**——即 LLM 之外的兜底逻辑。
    这部分才是本次修复的主体(LLM 是否遵守 Prompt 属于概率性行为)。

用法:
    cd /Users/xiahua/trea/数据/backend
    python3 ../scripts/verify_report_pipeline.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

BACKEND = Path(__file__).resolve().parent.parent / "backend"
sys.path.insert(0, str(BACKEND))

from app.services import pandas_engine as pe  # noqa: E402
from app.services.analyzer import (  # noqa: E402
    _merge_sheet_indicators,
    _sheet_indicators_to_charts,
    _collect_indicator_values,
    _count_numbers,
    enforce_numeric_density,
    verify_chart_data,
    _enforce_chart_data_integrity,
    audit_report_numbers,
    audit_section_numbers,
)
from app.core.prompt_builder import build_user_prompt  # noqa: E402

DATA = BACKEND / "backend/uploads/20261019_134239_年度复盘_统计指标数据.xlsx"
FAKE_REPORT = Path("/tmp/t65.json")   # task #65 的真实产物(LLM 编造版)

_FAILS: list[str] = []


def check(cond: bool, label: str, detail: str = "") -> None:
    mark = "PASS" if cond else "FAIL"
    print(f"  [{mark}] {label}" + (f" — {detail}" if detail else ""))
    if not cond:
        _FAILS.append(label)


def main() -> int:
    print("=" * 78)
    print("端到端离线演练:年度报告「不只是文字」")
    print("=" * 78)

    # ---------- 准备数据 ----------
    per_sheets = pe.read_dataframe_per_sheet(DATA)
    s_inds = []
    for e in per_sheets:
        df = e["df"]
        if len(df) == 0:
            continue
        ind = pe.compute_indicators(pe.clean(df, {}))
        ind["name"] = e["name"]
        s_inds.append(ind)
    indicators = _merge_sheet_indicators(s_inds)

    print()
    print("【第 1 层】指标引擎 —— 喂给 LLM 的是不是干净真实数据?")
    kpi = indicators.get("kpi_matrix") or {}
    recs = kpi.get("records") or []
    check(len(recs) >= 5, "KPI 宽表识别出逐指标记录", f"{len(recs)} 条")
    if recs:
        r0 = recs[0]
        print(f"         示例:{r0}")
        check(
            r0.get("指标") and r0.get("实际") is not None and r0.get("目标") is not None,
            "KPI 记录含指标名/目标/实际",
        )
    check("年度核心KPI::年度目标" not in (indicators.get("numeric_summary") or {}),
          "KPI 列未跨量纲求和(不再有 sum=67052)")
    trend_keys = [k for k, v in (indicators.get("trend") or {}).items()
                  if isinstance(v, dict) and v.get("x")]
    check(bool(trend_keys), "周期趋势已产出(不再是空 trend)")
    for tk in trend_keys:
        tv = indicators["trend"][tk]
        print(f"         {tk}: {tv['x'][:4]}... 合计={tv.get('总计')} 峰值={tv.get('峰值')}")
        check("全年合计" not in tv["x"], f"[{tk}] 趋势不含汇总行")

    print()
    print("【第 2 层】图表防编造 —— LLM 编的图会被拦吗?")
    if not FAKE_REPORT.exists():
        print(f"  [SKIP] 找不到 {FAKE_REPORT},跳过(该文件是 task #65 的历史产物)")
    else:
        fake_co = json.loads(Path("/tmp/c65.json").read_text()) \
            if Path("/tmp/c65.json").exists() else {}
        fake_by_section = (fake_co or {}).get("by_section") or {}
        total_charts = blocked = 0
        for sec, arr in fake_by_section.items():
            for c in arr:
                if not isinstance(c, dict):
                    continue
                total_charts += 1
                ok, vals = verify_chart_data(c, indicators)
                if not ok:
                    blocked += 1
                    print(f"         拦截「{sec}」《{c.get('title')}》数值={vals[:4]}")
        check(total_charts > 0, "读到 task #65 的历史编造图",
              f"共 {total_charts} 张")
        if total_charts:
            check(blocked == total_charts,
                  "历史编造图全部被拦",
                  f"{blocked}/{total_charts}")

            # ---- 修复效果验证:替换后的图必须是真实数据,且不能 6 章挂同一张 ----
            real_charts = _sheet_indicators_to_charts(s_inds)
            fixed_by_sec, fix_notes = _enforce_chart_data_integrity(
                fake_by_section, indicators, real_charts
            )
            print("         替换后各章节图表:")
            got_titles: list[str] = []
            for sec, arr in fixed_by_sec.items():
                for c in arr:
                    if not isinstance(c, dict):
                        continue
                    t = str(c.get("title"))
                    got_titles.append(t)
                    shared = c.get("chart_shared_from")
                    mark = f" ← 复用自《{shared}》" if shared else ""
                    print(f"           {sec}: {t}{mark}")
            check(len(fix_notes) > 0, "产生了图表校正记录")
            check(all(
                verify_chart_data(c, indicators)[0]
                for arr in fixed_by_sec.values() for c in arr
                if isinstance(c, dict)
            ), "替换后的图表全部通过真实性校验")
            dup_titles = {t for t in got_titles if got_titles.count(t) > 1}
            check(not dup_titles, "没有多个章节挂同一张同名图",
                  f"重复: {sorted(dup_titles)}" if dup_titles else "")
            noise = [t for t in got_titles
                     if any(w in t for w in ("备注", "频次分布"))]
            check(not noise, "没有把「备注频次」这类无意义图挂进章节",
                  f"发现: {noise}" if noise else "")

    print()
    print("【第 3 层】正文数字密度 —— 有数字吗?")
    sections = {
        "一、总体经营概况":
            "核心结论:整体经营处于「收入低速增长、利润承压」的剪刀差区间,"
            "增长质量而非增长速度是当前第一矛盾。",
        "二、收入与增长结构":
            "收入增长是「单引擎驱动」,结构失衡程度高于增速回落到何种水平本身。",
        "三、成本与盈利能力":
            "利润下滑的主因不是费用失控,而是毛利空间收窄叠加固定成本摊薄不足。",
        "四、客户与市场维度":
            "客户盘子「数量稳、质量降」,单客价值与复购是当前最需要盯住的两个指标。",
        "五、主要风险提示":
            "有三类风险未被当期指标充分反映,但会在后续两到四个周期兑现。",
        "六、下一步行动建议":
            "建议以「稳毛利、优结构、提单客」为三条主线,在一个季度内落地。",
    }
    before = {k: _count_numbers(v) for k, v in sections.items()}
    print(f"         修复前:{before}")
    check(all(v == 0 for v in before.values()), "修复前:六章全是零数字")

    fixed, notes = enforce_numeric_density(sections, indicators)
    after = {k: _count_numbers(v) for k, v in fixed.items()}
    print(f"         修复后:{after}")
    check(all(v >= 3 for v in after.values()), "修复后:每章至少 3 个数字")
    check(len(notes) == len(sections), "全部章节都被补齐")

    print()
    print("【第 4 层】章节明细数据表 —— 能核对数字来源吗?")
    tables = pe.build_section_detail_tables(indicators, list(fixed.keys()))
    check(len(tables) > 0, "已生成明细表", f"{len(tables)} 张")
    for sec, t in tables.items():
        print(f"         【{sec}】{t['title']}({t['total_rows']} 行)")
    check(all(t.get("columns") and t.get("rows") for t in tables.values()),
          "每张表都有表头与数据行")

    # ---------- Prompt 是否带上了 KPI 指引 ----------
    print()
    print("【Prompt】KPI 写作指引是否注入?")
    prompt = build_user_prompt(
        category="年度",
        structure="\n".join(f"{i+1}. {k}" for i, k in enumerate(fixed.keys())),
        indicators=indicators,
        template_description="年度经营复盘",
    )
    check("KPI 明细表" in prompt, "KPI 写作指引已注入 Prompt")
    check("数字密度硬约束" in prompt, "数字密度硬约束已注入 Prompt")
    check("达成率%" in prompt, "百分数字段已随指标传给 LLM")
    check("可用数字白名单" in prompt, "可用数字白名单已注入 Prompt")
    # 白名单必须不含浮点尾巴(它是模型的直接抄写源)
    _wl_seg = prompt.split("【可用数字白名单")[1].split("【")[0] \
        if "【可用数字白名单" in prompt else ""
    check("4680.0" not in _wl_seg,
          "白名单里无浮点尾巴(4680 而非 4680.0)")
    # 【2026-10】数据必须排在最前:模型长上下文注意力衰减,
    # 数据放在 1 万字符之后会被忽略,导致模型声称"未附可核验的经营指标"
    check(prompt.index("【本期真实数据】") < prompt.index("【章节清单"),
          "Prompt 里真实数据排在章节清单之前(占据模型注意力黄金位置)")
    # 无 KPI 场景不应注入 KPI 指引
    prompt_no_kpi = build_user_prompt(
        category="年度",
        structure="1. 执行摘要",
        indicators={"numeric_summary": {"a": {"sum": 1}}, "trend": {}, "by_group": {}},
    )
    check("KPI 明细表 · 必须逐条引用" not in prompt_no_kpi,
          "无 KPI 数据时不注入 KPI 指引(避免噪声)")

    # ---------- 正文数字溯源(第 5 层:正文也会编数字) ----------
    print()
    print("【第 5 层】正文数字溯源 —— 编造的数字会被标注吗?")
    pool = _collect_indicator_values(indicators)
    # 用实测跑出来的真实产物:模型编造了费用率/现金流等源数据里没有的字段
    real_sec = "三、成本费用与盈利能力"
    real_text = (
        "费用刚性是利润下滑的第二根支柱:期间费用率 22.4%,同比上升 1.1 个百分点,"
        "其中销售费用率 11.8%、管理费用率 6.1%。"
        "经营活动现金流净额 0.71 亿元,净利润现金含量由 1.10 降至 0.76。"
    )
    bad_nums = audit_section_numbers(real_sec, real_text, pool)
    print(f"         无出处数字: {bad_nums}")
    check("22.4" in bad_nums and "11.8" in bad_nums,
          "编造的「期间费用率 22.4%」被识别为无出处")
    check("6.1" in bad_nums,
          "编造的「管理费用率 6.1%」被识别为无出处")
    # 源数据里确实没有费用率/现金流这些字段,所以 1.1 也应是"无出处"
    # (实测池中 0~2 区间只有 0.01/0.033/0.04… 这类同比达成率,没有 1.1)
    check("1.1" in bad_nums,
          "编造的「同比上升 1.1 个百分点」被识别为无出处")

    # 反向验证:真实存在的数字不应被误报
    real_only_text = "销售额实际 4680 万元,达成率 93.6%,同比 11.43%。"
    none_bad = audit_section_numbers(real_sec, real_only_text, pool)
    check(not none_bad,
          "全部数字有出处的段落不产生误报",
          f"误报 {none_bad}" if none_bad else "")

    audited, audit_notes = audit_report_numbers(
        {real_sec: real_text}, indicators
    )
    check(bool(audit_notes), "不可信章节产生告警")
    check("未能在源数据中找到对应" not in audited[real_sec]
          or audited[real_sec].startswith("⚠️"),
          "重度污染章节按新策略处置")

    # ---- 分档处置验证(2026-10 新增)----
    # 重度污染 → 剔除编造句,但**保留模型的分析与建议**(2026-10 修订)
    # 用例刻意构造成"一半假数字 + 一半真实分析",验证:
    #   假数字句子被删 → 真分析(含行动建议)必须留下
    heavy_text = (
        "核心结论:本期经营整体承压。"
        "营业收入12.86亿元,同比下降5.2%,环比下降2.7%,利润端同步走弱。"
        "期间费用率22.4%,其中销售费用率11.8%,管理费用率6.1%,费用侵蚀利润。"
        "经营活动现金流净额0.71亿元,现金含量由1.10降至0.76。"
        "毛利率33.8%,较目标35%低1.2个百分点,是拖累利润的主因。"
        "订单量58200单,达成率97%。客单价804元,达成率94.59%,量的增长基本对冲了价的下滑。"
        "建议由销售负责人牵头,在Q2将客单价提升至850元,并按月复盘达成进度。"
    )
    heavy_out, heavy_notes = audit_report_numbers({"一、经营概览": heavy_text}, indicators)
    heavy_body = heavy_out["一、经营概览"]
    print(f"         重度处置结果: {heavy_body[:150]}")
    check("12.86" not in heavy_body and "22.4" not in heavy_body,
          "重度污染章节的编造数字已从正文移除")
    # 【核心诉求】用户的原话:"报告不能只是说数据,要对数据进行分析和给方案"
    check("毛利率" in heavy_body and "拖累利润" in heavy_body,
          "重度污染章节保留了模型的核心分析判断")
    check("建议" in heavy_body and "销售负责人" in heavy_body,
          "重度污染章节保留了模型的行动方案")
    # 【话术清理】用户明确要求不要出现"未覆盖/不予采纳"
    check("未覆盖" not in heavy_body and "不予采纳" not in heavy_body,
          "正文不含'未覆盖/不予采纳'类系统自述话术")

    # 轻度污染(≤35%)→ 保留原文 + 段末轻量提示
    light_text = (
        "本期销售额 4680 万元,目标 5000 万元,达成率 93.6%,同比 11.43%;"
        "订单量 58200 单,达成率 97%;客单价 804 元,达成率 94.59%;"
        "毛利率 33.8%,同比 4%;新客户数 1350 个,达成率 112.5%;"
        "另有一处待核数据 99.99。毛利率承压,需重点关注。"
    )
    light_out, light_notes = audit_report_numbers({"一、经营概览": light_text}, indicators)
    light_body = light_out["一、经营概览"]
    check("4680" in light_body and "毛利率承压" in light_body,
          "轻度污染章节完整保留原文与结论")
    check(light_body.rstrip().endswith("(注:本章 1 处数值请以下方明细表为准)"),
          "轻度污染章节在段末追加轻量核对提示",
          f"实际结尾: {light_body.rstrip()[-40:]}")
    check("数据核对提示" not in light_body,
          "轻度污染不再使用'数据核对提示'标题式话术")

    # 系统摘要标记:必须被识别为免检,且不参与编造率计算
    from app.services.analyzer import _AUDIT_SKIP_MARKER
    marked = {"一、经营概览": _AUDIT_SKIP_MARKER + heavy_text}
    m_out, _ = audit_report_numbers(marked, indicators)
    check(m_out["一、经营概览"].startswith(_AUDIT_SKIP_MARKER),
          "系统摘要前缀在处置后仍保留(供后续剥离)")
    check("12.86" not in m_out["一、经营概览"],
          "系统摘要不稀释判定:原文编造数字仍被识别")

    # 全部数字有出处的章节,不应被加提示
    good_sec = "一、经营概览"
    good_text = (
        "本期销售额实际 4680 万元,目标 5000 万元,达成率 93.6%,同比 11.43%;"
        "订单量实际 58200 单,达成率 97%;客单价实际 804 元,达成率 94.59%;"
        "毛利率 33.8%,同比 4%;新客户数实际 1350 个,达成率 112.5%。"
    )
    audited2, notes2 = audit_report_numbers({good_sec: good_text}, indicators)
    check(not notes2, "全部数字有出处的章节不产生告警",
          f"实际告警 {notes2}" if notes2 else "")
    check(not audited2[good_sec].startswith("⚠️"),
          "可信章节正文保持原样")

    # 比率类指标的百分数形态必须能溯源成功
    # (源数据存 0.338,正文按业务习惯写 33.8%,早期版本会误判为编造)
    pct_text = "毛利率 33.8%,目标 35%,达成率 96.57%,同比 4.0%。"
    pct_bad = audit_section_numbers(good_sec, pct_text, pool)
    check(not pct_bad,
          "百分数形态(33.8%)能溯源到源数据(0.338)",
          f"误报 {pct_bad}" if pct_bad else "")

    # ---- 分段验证(2026-10 新增:用户要求"章节要分段") ----
    print()
    print("【分段】章节正文会按语义/长度拆段吗?")
    from app.services.analyzer import split_paragraphs
    long_text = (
        "核心结论:增长由订单量驱动。销售额4680万元,同比11.43%。"
        "其一,订单量同比增7.78%,而客单价仅增3.34%,说明增长主要来自单量扩张。"
        "其二,渠道结构分化明显,线上占比41.67%,门店同比-7.69%。"
        "建议由销售负责人牵头,将客单价提升至850元并按月复盘。"
    )
    paras = split_paragraphs(long_text)
    for i, p in enumerate(paras, 1):
        print(f"         [{i}] {p[:50]}")
    check(len(paras) >= 3,
          "单段长文被拆成 3 段以上", f"实际 {len(paras)} 段")
    check(any(p.startswith("其一") for p in paras),
          "「其一」等语义标记处正确断段")
    check(all(len(p) <= 260 for p in paras),
          "每段长度可控(无超长段落)")
    keep = split_paragraphs("第一段结论。\n\n第二段结论。")
    check(len(keep) == 2, "模型已有的空行分段被保留",
          f"实际 {len(keep)} 段")

    print()
    print("=" * 78)
    if _FAILS:
        print(f"{len(_FAILS)} 项失败:")
        for f in _FAILS:
            print(f"  - {f}")
        return 1
    print("全链路离线演练通过:报告正文有数字、图表不编造、明细表可核对。")
    return 0


if __name__ == "__main__":
    sys.exit(main())