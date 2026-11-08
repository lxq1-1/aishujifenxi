"""LLM 输出能力探针:定位"空 sections"的真正原因。

背景:
    2026-10-20 报告生成连续失败,MiniMax-M3.1-Flash 稳定返回
        {"title":"...","sections":{},"chart_option":{...}}
    需要区分三种可能:
        A. Prompt 太长/位置不对,模型没读到数据 → 精简/调顺序能解决
        B. 模型能力不足,读到了也不写 → 换模型能解决
        C. 指令冲突,模型被某条规则卡住 → 改指令能解决

用法:
    cd /Users/xiahua/trea/数据/backend
    python3 ../scripts/probe_llm_sections.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

BACKEND = Path(__file__).resolve().parent.parent / "backend"
sys.path.insert(0, str(BACKEND))

from app.core.prompt_builder import SYSTEM_PROMPT, build_user_prompt  # noqa: E402
from app.services import pandas_engine as pe  # noqa: E402
from app.services.analyzer import _merge_sheet_indicators, _shrink_indicators_for_llm  # noqa: E402
from app.services.llm_client import llm  # noqa: E402

DATA = BACKEND / "backend/uploads/20261019_134239_年度复盘_统计指标数据.xlsx"
SECTIONS = "1. 经营结论\n2. 指标对比\n3. 根因分析\n4. 行动计划\n5. 风险提示"


def _indicators() -> dict:
    per = pe.read_dataframe_per_sheet(DATA)
    sl = []
    for e in per:
        if len(e["df"]) == 0:
            continue
        i = pe.compute_indicators(pe.clean(e["df"], {}))
        i["name"] = e["name"]
        sl.append(i)
    return _shrink_indicators_for_llm(_merge_sheet_indicators(sl))


def probe(label: str, user: str, use_system: bool = True) -> dict:
    """发一次请求,返回解析结果摘要。"""
    print()
    print(f"──── {label} (user {len(user)} 字符) ────")
    try:
        if use_system:
            raw = llm.chat(SYSTEM_PROMPT, user, json_mode=True)
        else:
            raw = llm.chat(SYSTEM_PROMPT, user, json_mode=True)
    except Exception as exc:
        print(f"  调用失败: {type(exc).__name__}: {exc}")
        return {}
    # chat_json 已返回 dict;chat 返回 str。统一成 dict
    if isinstance(raw, dict):
        return _summarize(label, raw)
    s = str(raw).strip()
    if s.startswith("```"):
        s = s.split("```", 2)[1]
        if s.startswith("json"):
            s = s[4:]
    try:
        data = json.loads(s)
    except Exception:
        print(f"  非 JSON 输出,前 300 字: {s[:300]}")
        return {}
    return _summarize(label, data)


def _summarize(label: str, data: dict) -> dict:
    secs = data.get("sections") or {}
    print(f"  title      = {str(data.get('title'))[:60]}")
    print(f"  sections   = {len(secs)} 章 {list(secs)[:6]}")
    for k, v in list(secs.items())[:1]:
        print(f"    [{k}] {str(v)[:200]}")
    ch = ((data.get("chart_option") or {}).get("by_section") or {})
    print(f"  by_section = {len(ch)} 章")
    return data


def main() -> int:
    ind = _indicators()
    full = build_user_prompt("年度", SECTIONS, ind, "年度复盘")

    # ---- 探针 1:完整 Prompt(复现线上问题) ----
    probe("① 完整 Prompt(线上同款)", full)

    # ---- 探针 2:极简 Prompt,只给数据 + 5 个章节名 ----
    minimal = (
        "【本期真实数据】\n"
        + json.dumps(ind, ensure_ascii=False, indent=2, default=str)
        + "\n\n【任务】为下列 5 个章节各写 150~250 字中文分析,"
          "必须引用上面的真实数字,禁止编造。\n"
          + SECTIONS
        + '\n\n严格输出 JSON:{"sections":{"<章节名>":"<正文>"}}'
    )
    probe("② 极简 Prompt(无 system/无写作规范)", minimal, use_system=False)

    # ---- 探针 3:真实数据 + 详细要求,但不带图表要求 ----
    no_chart = build_user_prompt("年度", SECTIONS, ind, "年度复盘")
    for kw in ("chart_option", "by_section", "可视化要求", "图表"):
        pass
    import re as _re
    stripped = _re.sub(r"【可视化要求.*?(?=\n\n【|$)", "", no_chart, flags=_re.S)
    stripped = _re.sub(r'"chart_option".*?"\n  \}', "", stripped, flags=_re.S)
    probe("③ 去掉可视化/图表要求", stripped)

    print()
    print("=" * 78)
    print("判读:")
    print("  ① 空 + ② 有内容 → 原 Prompt 过长或规则冲突(能力够,表达受阻)")
    print("  ①② 都空            → 模型对该任务能力不足,需换更强模型")
    print("  ① 空 + ③ 有内容   → 图表/可视化要求把模型压垮,应拆成两次调用")
    return 0


if __name__ == "__main__":
    sys.exit(main())