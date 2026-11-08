"""演示稿 Prompt 构建器：把「一份报告」变成「一份演示稿大纲」。

为什么需要独立的一套 Prompt:
    报告 Prompt(`prompt_builder.py`)的目标是**写全**:章节齐全、论证充分、
    数字密集,读者是"通读全文的人"。
    演示稿的目标完全相反:**讲透一页**:一页一个判断、几条要点、一张图、一段讲稿,
    读者是"听汇报、看投影的人"。
    用同一套 Prompt 必然产出"报告被切成块"的结果,那不是演示稿。

所以这里另起一套 system prompt,核心约束:
    1. **结论先行**——每页标题必须是一个判断句,不是章节名;
    2. **bullet 化**——每页 3~5 条要点,每条一行,不是段落;
    3. **图表引用**——每页最多引用一张可用图表(用 chart_ref 指名);
    4. **不编数字**——所有数字必须来自报告正文/指标,后端还会再做一次溯源;
    5. **禁止思考过程**——沿用项目里对 reasoning 模型的一贯要求。

输出是纯 JSON,结构:
    {
      "title": str,
      "subtitle": str,
      "slides": [
        {
          "kind": "cover|agenda|content|chart|metrics|conclusion",
          "title": str,
          "bullets": [str, ...],
          "chart_ref": str,     # 只能填【可用图表清单】里的 id 或标题
          "notes": str          # 演讲备注(讲稿)
        }
      ]
    }
"""
from __future__ import annotations

import json
from typing import Any

from app.core.json_utils import json_default


# ============================================================
# 演示稿专用系统提示
# ============================================================
SYSTEM_PROMPT_DECK = (
    "你是一名资深商业演示顾问,专长是把「数据分析师写的长报告」改写成"
    "「高管愿意听完的演示稿」。\n\n"
    "【演示稿与报告的根本区别】\n"
    "  报告是给人**读**的,演示稿是给人**听**的。听众不会读满屏文字,\n"
    "  他们只看标题、扫要点、看图。所以演示稿追求的是**信息密度**与**节奏**,\n"
    "  而不是覆盖面。你要做的不是「把报告搬进幻灯片」,而是**重新组织叙事**。\n\n"
    "【第 1 条 · 结论先行(最重要)】\n"
    "  每页的 title 必须是一个**判断句 / 结论句**,不是章节名、不是话题标签。\n"
    "    反例(章节名,不合格):「总体经营概况」「问题与风险分析」「行动计划」\n"
    "    正例(判断句,合格):\n"
    "      「销售增长 11.43%,但较目标仍差 320 万元」\n"
    "      「增长质量是短板:新客超额,客单价持续未达标」\n"
    "      「风险集中在回款与留存两项硬指标上」\n"
    "    判断标准:标题读出来就像一句可以直接说出口的话。\n\n"
    "【第 2 条 · 要点化】\n"
    "  每页 bullets 给 **3~5 条**,每条**一行**、≤ 60 字,是「关键事实」而不是「话题」。\n"
    "    - 每条要点应能被独立听懂,且尽量带一个具体数字或明确判断;\n"
    "    - 禁止把正文段落直接搬进 bullets(那是复制粘贴,不是提炼);\n"
    "    - 禁止出现「本章介绍了…」「综上所述…」这类元叙述;\n"
    "    - 宁可少写一条,也不要写一句空话凑数。\n\n"
    "【第 3 条 · 叙事线】\n"
    "  6~15 页,构成一条完整叙事线,推荐结构(可按素材调整,但要有起承转合):\n"
    "    1) 开局(总体结论:一句话说清本期表现);\n"
    "    2) 亮点(哪些方面做得好,靠什么);\n"
    "    3) 问题(哪里没达标,差多少);\n"
    "    4) 归因(为什么会这样);\n"
    "    5) 行动(下一步做什么、谁做、做到什么程度);\n"
    "  不要求严格按报告章节顺序切分,可以**跨章节重组**,\n"
    "  但所有内容必须来自下面的【报告正文】与【关键指标】,不得引入外部信息。\n\n"
    "【第 4 条 · 图表引用】\n"
    "  每页最多引用 **1 张**图表,填在 chart_ref 字段。\n"
    "    - chart_ref **只能**填【可用图表清单】里给出的 id 或图表标题,逐字一致;\n"
    "    - 找不到合适图表时,chart_ref 填空字符串(\"\");\n"
    "    - 同一张图**不要**在多页重复引用(后端也会做去重,但你应主动避免);\n"
    "    - 图要能**佐证本页标题的结论**,不是「有图就挂」。\n\n"
    "【第 5 条 · 演讲备注】\n"
    "  **每页都必须写 notes,不允许留空**。notes 是 2~4 句**讲稿**(≤ 300 字),\n"
    "  是给汇报人照着念的口播稿:\n"
    "    - 先把本页结论说清楚,再点出最该讲的一个数字,最后给一个过渡;\n"
    "    - 用口语,像真的在讲,不要写成「本页要点:1. 2. 3.」这种提纲;\n"
    "    - notes 里的数字同样必须来自真实数据。\n"
    "  【重要】不要偷懒省略 notes:漏写的页在导出 PPT 时会没有讲稿,\n"
    "  汇报人只能干念标题。你每一页都要替他把这段话讲出来。\n\n"
    "【第 6 条 · 数字纪律】\n"
    "  - 所有数字**必须**能在【报告正文】或【关键指标】里找到原值,**严禁编造**;\n"
    "  - **不得换算单位**(白名单写 4680 万元就写 4680 万元,不要写成 0.468 亿元);\n"
    "  - 不确定的数字**直接不写**,不要猜一个看起来合理的值;\n"
    "  - 后端会对你的要点做一次数字溯源,查不到出处的整条要点会被删除,\n"
    "    所以宁可写得少,也不要写得虚。\n\n"
    "【第 7 条 · 输出格式硬约束】\n"
    "  你的整篇回复**必须是合法 JSON**,第一字符必须是 `{`,最后一字符必须是 `}`。\n"
    "  严禁:```json 代码块包裹 / markdown 标题 / markdown 表格 / JSON 前的导言文字。\n"
    "  严禁出现 <think>、<thinking> 等思考标签,也不要写 Let me analyze、\n"
    "  Looking at 之类的英文推导过程。中文回答,结构化输出。\n"
)


# ============================================================
# 输出 schema(与 SYSTEM_PROMPT_DECK 配套,单独拼进 user 消息)
# ============================================================
_DECK_SCHEMA_HINT = (
    "【输出 JSON 格式】严格按下面的结构输出:\n"
    "{\n"
    '  "title": "演示稿主标题(中文,≤ 24 字)",\n'
    '  "subtitle": "副标题(中文,≤ 40 字,一句话点明本期核心结论)",\n'
    '  "slides": [\n'
    "    {\n"
    '      "kind": "content",\n'
    '      "title": "该页的结论句(判断句,不是章节名)",\n'
    '      "bullets": [\n'
    '        "要点一(一行,≤ 60 字)",\n'
    '        "要点二",\n'
    '        "要点三"\n'
    "      ],\n"
    '      "chart_ref": "可用图表清单里的 id 或标题(没有就写空字符串)",\n'
    '      "notes": "该页的演讲讲稿(2~4 句)"\n'
    "    }\n"
    "  ]\n"
    "}\n"
    "【字段说明】\n"
    "  - kind 取值:cover(封面,可不输出) / agenda(目录,可不输出) /\n"
    "    content(内容页,默认) / chart(以图为主) / metrics(以指标为主) /\n"
    "    conclusion(总结页)。除 cover/agenda 外一律用 content 最稳妥。\n"
    "  - slides 数量控制在 6~15 页。\n"
    "  - bullets 每页 3~5 条;某页实在没有要点就给 1~2 条,但不要留空数组。\n"
)


# ============================================================
# 各段落渲染
# ============================================================
def _render_report_sections(report: dict[str, Any], max_chars: int = 8000) -> str:
    """把报告正文渲染成给 LLM 看的文本,总长受控。

    为什么要有上限:
        报告正文动辄 1~2 万字,全塞进 prompt 会让 reasoning 模型思考链爆炸、
        生成极慢(实测主报告链路 in_len 12k 就要 ~500s)。
        演示稿只需要"骨架 + 关键论点",所以按章节均摊预算截断。
    """
    sections = (report or {}).get("sections") or {}
    if not isinstance(sections, dict) or not sections:
        return "(报告正文为空)"

    entries: list[tuple[str, str]] = []
    for k, v in sections.items():
        title = str(k or "").strip()
        if isinstance(v, str):
            body = v.strip()
        elif isinstance(v, dict):
            # 结构化章节(合同对比的差异表等)转成 JSON 摘要,
            # 演示稿只需知道"这里有明细数据",不需要逐行读。
            body = json.dumps(v, ensure_ascii=False, default=json_default)[:600]
        elif isinstance(v, list):
            body = json.dumps(v, ensure_ascii=False, default=json_default)[:600]
        else:
            body = str(v)
        if not title:
            continue
        entries.append((title, body))

    if not entries:
        return "(报告正文为空)"

    # 每章均摊预算:至少 400 字,最多 1500 字,总预算 max_chars
    per_section = max(400, min(1500, max_chars // max(1, len(entries))))
    lines: list[str] = []
    used = 0
    for title, body in entries:
        chunk = body[:per_section]
        if len(body) > per_section:
            chunk += "…（本节已截断）"
        block = f"【{title}】\n{chunk}"
        if used + len(block) > max_chars:
            break
        lines.append(block)
        used += len(block)

    return "\n\n".join(lines) if lines else "(报告正文为空)"


def _render_indicators(indicators: dict[str, Any]) -> str:
    """关键指标段:优先复用报告链路的瘦身口径,避免 prompt 爆炸。

    复用 `analyzer._shrink_indicators_for_llm` 而不是自己写一份,
    是为了保证"演示稿看到的数字"与"报告看到的数字"来自同一份瘦身逻辑,
    不会出现"报告里没这个指标、演示稿里却有"的口径漂移。
    """
    try:
        from app.services.analyzer import _shrink_indicators_for_llm
        slim = _shrink_indicators_for_llm(indicators)
    except Exception:
        slim = indicators if isinstance(indicators, dict) else {}
    if not slim:
        return "(无指标数据)"
    text = json.dumps(slim, ensure_ascii=False, default=json_default)
    # 指标段最多 4000 字符:演示稿只需要"有哪些数字可用",不需要全部细节
    if len(text) > 4000:
        text = text[:4000] + "…（已截断）"
    return text


def _render_chart_catalog(chart_catalog: list[dict[str, Any]] | None) -> str:
    """可用图表清单段:告诉 LLM "有哪些图可以引用",以及每个图的 id。"""
    if not chart_catalog:
        return "(本报告没有可用图表,chart_ref 一律填空字符串)"
    lines = []
    for item in chart_catalog:
        lines.append(
            f"- id={item.get('id', '')} | 标题={item.get('title', '')}"
            f" | 所属章节={item.get('section', '')}"
        )
    return "\n".join(lines)


def build_deck_prompt(
    report: dict[str, Any],
    indicators: dict[str, Any],
    chart_catalog: list[dict[str, Any]] | None = None,
) -> str:
    """拼装演示稿生成的 user 消息。

    Args:
        report:       已生成的结构化报告(title / sections / summary)。
        indicators:   指标快照(用于给 LLM 提供"可引用的真实数字")。
        chart_catalog:可用图表清单(build_chart_catalog 产出)。

    Returns:
        可直接交给 ``llm.chat_json`` 的 user 消息文本。
    """
    title = str((report or {}).get("title") or "").strip()
    summary = str((report or {}).get("summary") or "").strip()

    chunks: list[str] = [
        "请把下面这份数据报告,改写成一份可以直接拿去汇报的**演示稿大纲**。",
        _DECK_SCHEMA_HINT,
        "【报告标题】\n" + (title or "(无)"),
        "【报告摘要】\n" + (summary[:500] if summary else "(无)"),
        "【报告正文】\n" + _render_report_sections(report),
        "【关键指标】\n" + _render_indicators(indicators),
        "【可用图表清单】\n" + _render_chart_catalog(chart_catalog),
        (
            "【再次提醒】\n"
            "  - 标题写结论,不写章节名;\n"
            "  - 每页 3~5 条要点,一行一条;\n"
            "  - 数字只准用上面出现过的,不确定就不写;\n"
            "  - 只输出合法 JSON,不要任何解释文字。"
        ),
    ]
    return "\n\n".join(chunks)
