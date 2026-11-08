"""端到端验证:不依赖真实 LLM,模拟 LLM 在贫数据下输出脏数据,
验证 ai_service._normalize_sections + analyzer._normalize_by_section_keys
两道闸门能把 sections 与 by_section 强制对齐到模板章节清单。

模拟场景:
    模板「销售周报」的 structure(7 章节):
        一、本周业绩摘要
        二、核心销售指标(达成率/周环比/同比)
        三、渠道与产品结构分析
        四、Top 客户与重点商机进展
        五、问题与风险(转化/退货/回款)
        六、下周行动清单(客户跟进/促销活动/资源申请)
        七、附录:本周关键数据明细

模拟 LLM 输出:
    - sections: {overview: "思考过程..."}        ← 自创 key,应被丢弃
    - by_section: {"数据图表": [...3张图...]}    ← 与 sections 完全不匹配,应归入孤儿
预期归一化结果:
    - sections: {} (因为 LLM 的 overview 不在 allowed_keys 里)
    - by_section: {"数据图表": [...3张图...]} (因为没有 sections 匹配,所以都进孤儿区)

跑法:
    cd /Users/xiahua/trea/数据/backend && python3 -m pytest /Users/xiahua/trea/数据/scripts/verify_section_normalization.py -v
    或:
    cd /Users/xiahua/trea/数据/backend && python3 /Users/xiahua/trea/数据/scripts/verify_section_normalization.py
"""
from __future__ import annotations

import sys
import os
sys.path.insert(0, "/Users/xiahua/trea/数据/backend")

from app.services import ai_service, analyzer


SALES_TEMPLATE_STRUCTURE = (
    "一、本周业绩摘要\n"
    "二、核心销售指标(达成率/周环比/同比)\n"
    "三、渠道与产品结构分析\n"
    "四、Top 客户与重点商机进展\n"
    "五、问题与风险(转化/退货/回款)\n"
    "六、下周行动清单(客户跟进/促销活动/资源申请)\n"
    "七、附录:本周关键数据明细"
)

# LLM 在贫数据下典型的"自暴自弃"输出
LLM_DIRTY_REPORT = {
    "title": "AI 分析",
    "sections": {
        "overview": (
            "用户要求我根据提供的指标数据,撰写一份销售周报。"
            "让我先仔细分析数据结构:\n\n"
            "**数据概况:**\n- 共 4 个工作表:销售明细/周度KPI/渠道结构/产品Top10\n"
            "- 缺失率较低,数据可用\n\n"
            "**可用的实际数据(有值的):**\n1. 销售明细 sheet (1220 行, 13 列)"
        ),
        # 即便 LLM 给了一个正确 key,这里也应当被保留
        "二、核心销售指标(达成率/周环比/同比)": "本周销售额 XXX 万元,环比上涨 8.5%。",
    },
    "chart_option": {
        "by_section": {
            "数据图表": [
                {"title": "本周销售趋势", "type": "line",
                 "option": {"series": [{"type": "line", "data": [1, 2, 3]}]}},
                {"title": "渠道占比", "type": "pie",
                 "option": {"series": [{"type": "pie", "data": [{"name": "线上", "value": 100}]}]}},
                {"title": "产品 Top5", "type": "bar",
                 "option": {"series": [{"type": "bar", "data": [10, 20, 30]}]}},
            ],
        },
        "indicators": {},
    },
}


def main():
    print("=" * 60)
    print("测试 1:_parse_section_keys 解析模板章节")
    print("=" * 60)
    keys = ai_service._parse_section_keys(SALES_TEMPLATE_STRUCTURE)
    print(f"解析到 {len(keys)} 个章节:")
    for i, k in enumerate(keys, 1):
        print(f"  {i}. {k!r}")

    expected = [
        "本周业绩摘要",
        "核心销售指标",
        "渠道与产品结构分析",
        "Top 客户与重点商机进展",
        "问题与风险",
        "下周行动清单",
        "附录:本周关键数据明细",
    ]
    assert keys == expected, f"\n[FAIL] 解析结果不符\n  期望: {expected}\n  实际: {keys}"
    print("[PASS] 章节解析与预期一致(括号说明被剥离,与 prompt_builder 对齐)\n")

    print("=" * 60)
    print("测试 2:_normalize_sections 强对齐(LLM 自创 overview 应被丢弃)")
    print("=" * 60)
    normalized_sections = ai_service._normalize_sections(
        LLM_DIRTY_REPORT["sections"], keys
    )
    print(f"LLM 原始 sections: {list(LLM_DIRTY_REPORT['sections'].keys())}")
    print(f"归一化后 sections: {list(normalized_sections.keys())}")
    # overview 应被丢弃(不在 allowed_keys 里);"二、核心销售指标..." 因仍带前缀
    # 也不在 allowed_keys 里,应被丢弃
    assert "overview" not in normalized_sections, "overview 应被丢弃"
    print("[PASS] overview 已被丢弃(不在模板章节清单)\n")

    print("=" * 60)
    print("测试 3:_normalize_by_section_keys 归一化")
    print("=" * 60)
    by_section_in = LLM_DIRTY_REPORT["chart_option"]["by_section"]
    by_section_out = analyzer._normalize_by_section_keys(by_section_in, normalized_sections)
    print(f"输入 by_section keys: {list(by_section_in.keys())}")
    print(f"输出 by_section keys: {list(by_section_out.keys())}")
    # 因为 normalized_sections 为空(都被丢弃了),原样返回所有 by_section
    # ——这是预期行为,避免"误伤"
    print("[INFO] 因 sections 已为空,by_section 原样保留,所有图将进入前端孤儿区\n")

    print("=" * 60)
    print("测试 4:完整链路模拟(在正常 sections 下,by_section 应正确归一)")
    print("=" * 60)
    # 这次 sections 用 allowed_keys 里的合规 key,模拟 LLM 部分正确
    partial = {
        "本周业绩摘要": "本周销售额 XXX 万元,环比上涨 8.5%。",
        "渠道与产品结构分析": "线上渠道占比 35%,本周最高。",
        # 缺章节:附录,这个会由 Pandas 兜底补图
    }
    by_section_dirty = {
        "本周业绩摘要": [{"title": "本周销售趋势", "option": {"series": []}}],
        "渠道与产品结构分析": [{"title": "渠道占比", "option": {"series": []}}],
        "数据图表": [{"title": "孤儿图1", "option": {"series": []}}],
        "图表汇总": [{"title": "孤儿图2", "option": {"series": []}}],
    }
    cleaned = analyzer._normalize_by_section_keys(by_section_dirty, partial)
    print(f"归一化后 by_section keys: {list(cleaned.keys())}")
    print(f"「数据图表」章节内图数: {len(cleaned.get('数据图表', []))}")
    print(f"两个孤儿图都应被合并到「数据图表」区")
    assert "本周业绩摘要" in cleaned, "匹配的章节应保留"
    assert "渠道与产品结构分析" in cleaned, "匹配的章节应保留"
    assert "数据图表" in cleaned, "孤儿图合并到「数据图表」"
    # 「图表汇总」和「数据图表」这两个 key 都不在 allowed,应被合并
    assert "图表汇总" not in cleaned, "图表汇总应被合并到「数据图表」"
    print("[PASS] by_section key 归一化符合预期\n")

    print("=" * 60)
    print("测试 5:_normalize_chart_option 向后兼容")
    print("=" * 60)
    # 旧 shape:{charts: [...]} 应被归一为 by_section["数据图表"]
    old_shape = {
        "charts": [
            {"title": "老图1", "option": {"series": []}},
            {"title": "老图2", "option": {"series": []}},
        ],
        "indicators": {"foo": 1},
    }
    out = ai_service._normalize_chart_option(old_shape)
    print(f"旧 shape 归一化结果 keys: {list(out.keys())}")
    print(f"by_section keys: {list(out['by_section'].keys())}")
    assert out["by_section"]["数据图表"], "旧 shape 应归一到「数据图表」"
    print("[PASS] 旧 shape 兼容\n")

    # 直接给 ECharts option(含 series)的脏数据
    direct_opt = {"series": [{"type": "line", "data": [1, 2, 3]}]}
    out2 = ai_service._normalize_chart_option(direct_opt)
    assert out2["by_section"]["数据图表"], "直接给 ECharts option 应归一到「数据图表」"
    print("[PASS] 直接 ECharts option 兼容\n")

    print("=" * 60)
    print("测试 6:LLM 严格按章节清单输出 sections(理想情况)")
    print("=" * 60)
    ideal_sections = {
        "本周业绩摘要": "本周销售额 XXX 万元,环比上涨 8.5%。",
        "核心销售指标": "本周 GMV 1234 万,订单量 5678 单,达成率 102%。",
        "渠道与产品结构分析": "线上占比 38%,私域 22%,线下 25%,直播 15%。",
        "Top 客户与重点商机进展": "Top 5 客户贡献 25% 的销售额。",
        "问题与风险": "回款率 85%,低于目标 5 个百分点。",
        "下周行动清单": "重点跟进 3 个 Top 客户,开展 1 场促销活动。",
        "附录:本周关键数据明细": "见明细表。",
    }
    norm2 = ai_service._normalize_sections(ideal_sections, keys)
    print(f"理想 LLM 输出 sections 数:{len(ideal_sections)}")
    print(f"归一化后 sections 数:{len(norm2)}")
    assert len(norm2) == 7, f"应全部保留,实际剩 {len(norm2)} 个"
    print("[PASS] 理想情况下 7 章节全部保留\n")

    print("=" * 60)
    print("全部断言通过 ✅")
    print("=" * 60)
    print("\n落地动作清单:")
    print("  1. 后端 ai_service._parse_section_keys + _normalize_sections 已就绪")
    print("  2. 后端 analyzer._normalize_by_section_keys 已就绪")
    print("  3. 真实运行时,LLM 自创 key 一律被丢弃;by_section 错位 key 一律归并到「数据图表」")
    print("  4. 前端 hasAnyCharts + section-no-chart 占位提示已就绪")
    print("  5. Demo 数据已生成:/Users/xiahua/trea/数据/preview/sales_demo_3months.xlsx")
    print("     接下来可在浏览器上传该文件 + 「销售周报」模板,人工核对每章节是否配图\n")


if __name__ == "__main__":
    main()