"""AI 业务服务：报告生成、摘要、批注、对比、NL 查询。"""
from __future__ import annotations

import base64
import logging
import json
import re
from pathlib import Path
from typing import Any

from app.core.json_utils import json_default, strip_thinking, _safe_json_loads
from app.core.prompt_builder import SYSTEM_PROMPT, build_user_prompt
from app.services.llm_client import llm
from app.services import pandas_engine as pe

logger = logging.getLogger(__name__)


def generate_report(
    category: str,
    structure: str,
    indicators: dict,
    reference: str = "",
    template_description: str = "",
    analysis_intent: str = "",
    feedback: str = "",
    sheets: list | None = None,
) -> dict[str, Any]:
    """根据模板分类与章节结构生成结构化报告 + 图表 option。

    Prompt 的"角色/约束/JSON 输出 schema/占位符注入"统一由
    `app.core.prompt_builder.build_user_prompt` 处理,本函数只负责调用 LLM。

    template_description 是模板上填写的「业务背景」,会作为业务背景段注入 Prompt,
    不填则不影响输出(向后兼容)。
    analysis_intent 是用户在工作台「开始 AI 分析」前临时输入的"分析重点",
    feedback 是用户在报告页对已生成报告提的"修改意见" —— 二者均透传给
    prompt_builder,仅在非空时被注入到 Prompt。
    sheets 是数据集的工作表明细(`[{name, rows, cols}, ...]`),由 prompt_builder
    在 `len > 1` 时插入【数据来源】小节,引导 LLM 在写渠道/产品/目标等章节时
    主动引用各 sheet 的数字(而不是只盯 `_sheet_name` 中的某一个值)。

    返回值:
        {
            "title": "...",
            "sections": {...},
            "chart_option": {
                "by_section": { "<章节名>": [ChartItem,...] },
                "indicators": {...}
            }
        }
        其中 `chart_option.by_section` 是「按章节分布的图表索引」,空 dict 表示
        LLM 完全未给图(由调用方 `analyzer.run_analysis` 走 Pandas 兜底塞到
        `by_section["数据图表"]`,保证旧版"统一图表区"在历史/兜底场景下仍然显示)。
    """
    user = build_user_prompt(
        category,
        structure,
        indicators,
        reference,
        template_description,
        analysis_intent,
        feedback,
        sheets=sheets or [],
    )
    report = llm.chat_json(SYSTEM_PROMPT, user)
    # 二次清洗:有些 reasoning 模型会把思考块塞进 sections 的某个段落,
    # 这里递归剥一次,保证 title / sections 都不会泄漏 <think>。
    report = _strip_thinking_recursive(report)
    # 归一化 sections key:LLM 经常无视 Prompt 里的「章节清单」硬约束,
    # 自创 `overview` / `Highlights` / `数据图表` 这类 key,导致前端报告页只剩一
    # 个孤零零的 overview 段、所有章节正文与图都"失踪"。这里按 prompt
    # 实际传给 LLM 的章节清单,把 sections 的 key 强制对齐;LLM 自创的 key
    # 一律丢弃(若其正文含业务信息,并入摘要兜底,避免数据丢失)。
    if isinstance(report, dict):
        allowed_keys = _parse_section_keys(structure)
        report["sections"] = _normalize_sections(report.get("sections"), allowed_keys)
        report["chart_option"] = _normalize_chart_option(report.get("chart_option"))
    return report


def _parse_section_keys(structure: str) -> list[str]:
    """复用 prompt_builder 的章节解析,把用户填的 structure 还原成有序章节列表。

    目的:让 generate_report 拿到「Prompt 实际传给 LLM 的章节清单」,在
    _normalize_sections 里用它做 sections key 的强对齐 —— LLM 经常无视
    Prompt 里的章节清单硬约束,自创 `overview` / `Highlights` 等 key,
    这里负责把这部分脏数据纠正回原清单。

    与 prompt_builder._parse_section_titles 的策略完全一致(编号剥离、
    末尾冒号剥离、过长说明文字过滤、去重),保证两侧看到的章节名一字不差。
    失败兜底:返回一个空列表(走默认 4 段兜底,行为与 prompt_builder 一致)。
    """
    if not structure or not structure.strip():
        return []
    titles: list[str] = []
    seen: set[str] = set()
    for raw in structure.splitlines():
        line = raw.strip()
        if not line:
            continue
        # 复用 prompt_builder 的两段式前缀剥离逻辑(简化版)
        import re as _re
        # 已知前缀剥离:"一、" "1." "## " "- " 等
        prefix_re = _re.compile(
            r"""^\s*
                (?:
                    (?:[（(\[【]?\s*(?:[一二三四五六七八九十百千]+|[0-9]+|[①-⑳])\s*[)）\]】]?)\s*[、,.\)::]\s*
                  | [##]+\s+
                  | [-*]\s+
                )?
                (?P<title>.*?)
                \s*$
            """,
            _re.VERBOSE,
        )
        m = prefix_re.match(line)
        title = (m.group("title") if m and m.group("title") else line).strip()
        # 末尾冒号剥离
        title = _re.sub(r"\s*[::]\s*$", "", title).strip().rstrip(":").rstrip(":").strip()
        # 尾部破折号说明剥离
        for sep in (" —— ", " -- ", " — ", " - "):
            if sep in title:
                title = title.split(sep, 1)[0].strip()
                break
        # 尾部 (xxx) 备注剥离
        title = _re.sub(r"[\s(（][^()（）]{2,30}[)）]\s*$", "", title).strip()
        if not title:
            continue
        if len(title) > 60 or title.count("。") + title.count(".") >= 2:
            continue
        if title not in seen:
            seen.add(title)
            titles.append(title)
    return titles


def _normalize_sections(sections: Any, allowed_keys: list[str]) -> dict[str, str]:
    """把 LLM 返回的 sections 按「模板章节清单」归一化,优先保留 LLM 真实产出。

    行为(2026-10 修订):
        - sections 不是 dict → 直接返回空 dict。
        - LLM 的 key 与 allowed_keys 严格相等 → 原样保留。
        - LLM 的 key 与 allowed_keys 在宽松匹配(去前缀编号 / 去尾部冒号 / 去尾部破折号说明)
          后相等 → 也保留(LLM 经常在章节名前加 `## `、在尾部加冒号或冒号+说明,
          旧 strict 逻辑会让所有 key 都被丢弃,导致 sections 被判空并被本地拼接覆盖)。
        - LLM 自创的 key(任何宽松匹配都挂不上 allowed_keys)→ 仍然保留,
          按 LLM 实际写的内容落到 sections,前端照常渲染 ——
          用户拿到的是 AI 真正写的分析,而不是"29 行 × 24 列"。
        - 缺失的 allowed_keys 不补占位文案,避免后端"二次编造"。
        - value 不是 str 时统一 str() 化,避免前端 NPE。

    设计要点:
        - 旧 strict 行为会导致任务 #63 这种"LLM 真写了 7 个高质量段落,但 key 跟
          模板对不上"的场景被归一化为空,然后被本地 pandas 拼接覆盖成假报告。
        - 现在的策略是"先信任 LLM 输出"——LLM 自创的章节往往比模板更贴合数据形态,
          删掉就是丢掉价值。只有当 allowed_keys 完全空(模板没填章节)时才允许任意 key。
    """
    if not isinstance(sections, dict):
        return {}
    allowed = list(allowed_keys or [])
    # 没传 allowed_keys → 直接保留 LLM 原样
    if not allowed:
        return {str(k): str(v) for k, v in sections.items() if v}

    normalized: dict[str, str] = {}
    # 先建立"宽松后的 allowed_keys → 原 key"映射,严格相等已直接命中走 default 路径
    allowed_loose_map: dict[str, str] = {}
    for ak in allowed:
        for loose in _loosen_section_key(ak):
            allowed_loose_map.setdefault(loose, ak)

    for k, v in sections.items():
        ks = str(k)
        if not v:
            continue
        # 1) 严格相等
        if ks in allowed:
            normalized[ks] = str(v)
            continue
        # 2) 宽松匹配(LLM 加了前缀/尾部冒号/尾部说明等)
        for loose in _loosen_section_key(ks):
            if loose in allowed_loose_map:
                real_key = allowed_loose_map[loose]
                # 同一目标 key 只取首次命中,避免覆盖
                normalized.setdefault(real_key, str(v))
                break
        else:
            # 3) 完全自创的 key(LLM 没跟模板走)→ 仍保留,按 LLM 实际 key 渲染
            # 这样即使模板是「年度复盘」、LLM 写了「本周业绩摘要」,用户也能看到 AI 真写的分析
            normalized.setdefault(ks, str(v))
    if sections and not normalized:
        # 兜底:实在异常(理论上走不到这里)→ 原样返回
        normalized = {str(k): str(v) for k, v in sections.items() if v}
    return normalized


def _loosen_section_key(key: str) -> list[str]:
    """对一个章节名生成多种"等价形态",用于宽松匹配。

    常见 LLM 偏离模板的写法:
        - 加前缀:`## 一、执行摘要` / `一、执行摘要` / `1. 执行摘要` / `- 执行摘要`
        - 加尾部冒号:`执行摘要:` / `执行摘要:`
        - 加尾部说明:`执行摘要 —— 本期核心结论`(我们截到第一个破折号前)
        - 首尾空白 / 全角空格

    返回:该 key 可能的「等价形态」列表(包含原 key 自身),用于在 allowed_keys 中反查。
    """
    if not key:
        return [""]
    out: set[str] = {key}
    s = key.strip()
    out.add(s)
    # 末尾冒号
    for trail in (":", "：", ":", "："):
        if s.endswith(trail):
            out.add(s[: -len(trail)].strip())
    # 尾部破折号说明(只在分隔符处切,不切中间)
    for sep in (" —— ", " -- ", " — ", " - "):
        if sep in s:
            out.add(s.split(sep, 1)[0].strip())
    # 行首编号 / 列表前缀
    import re as _re
    head_re = _re.compile(
        r"""^\s*
        (?:
            (?:[（(\[【]?\s*(?:[一二三四五六七八九十百千]+|[0-9]+|[①-⑳])\s*[)）\]】]?)\s*[、,.\)::]\s*
          | [##]+\s+
          | [-*]\s+
        )?
        """,
        _re.VERBOSE,
    )
    stripped = head_re.sub("", s)
    if stripped != s:
        out.add(stripped.strip())
        out.add(stripped.strip().rstrip(":").rstrip("：").strip())
    return list(out)


def _normalize_chart_option(raw: Any) -> dict[str, Any]:
    """把 LLM 返回的 chart_option 统一为新结构 `{by_section, indicators}`。

    兼容旧 shape:
        - 直接给一个完整 ECharts option(有 `series`):塞进 by_section["数据图表"]
        - 旧 `{charts: [...]}`:把 charts 整体塞进 by_section["数据图表"]
        - 已是 `{by_section: {...}, indicators: ...}`:原样透传
        - 缺失 / 空 / 非 dict:返回 `{"by_section": {}, "indicators": {}}`,
          由调用方决定是否走 Pandas 兜底。
    """
    empty = {"by_section": {}, "indicators": {}}

    # 旧形状 1:LLM 把"整个 ECharts option"作为 chart_option 塞回来(没套对象)
    if isinstance(raw, dict) and "series" in raw and "by_section" not in raw and "charts" not in raw:
        return {"by_section": {"数据图表": [raw]}, "indicators": {}}

    if not isinstance(raw, dict):
        return empty

    # 已是新形状
    if isinstance(raw.get("by_section"), dict):
        by_section = raw["by_section"] or {}
        # 每个 value 必须是 list;非 list 包成 list(常见:LLM 直接给单个 dict)
        normalized = {
            k: (v if isinstance(v, list) else [v])
            for k, v in by_section.items()
            if v
        }
        return {
            "by_section": normalized,
            "indicators": raw.get("indicators") if isinstance(raw.get("indicators"), dict) else {},
        }

    # 旧形状 2:{charts: [...], indicators: {...}}
    if isinstance(raw.get("charts"), list) and raw.get("charts"):
        return {
            "by_section": {"数据图表": list(raw["charts"])},
            "indicators": raw.get("indicators") if isinstance(raw.get("indicators"), dict) else {},
        }

    # 兜底:既无 by_section 也无 charts → 空
    return empty


def _strip_thinking_recursive(obj: Any) -> Any:
    """对 dict/list/str 递归地剥离 <think>...</think> 块。

    说明:
        generate_report 返回的 report 是结构化对象(title + sections 字典),
        sections 每个 value 是字符串。Reasoning 模型有时会把思考段嵌进
        某个章节字符串里(例如"<think>…</think>\n正式文本"),只在外层
        strip 一次不够,需要递归。
    """
    if isinstance(obj, dict):
        return {k: _strip_thinking_recursive(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_strip_thinking_recursive(v) for v in obj]
    if isinstance(obj, str):
        return strip_thinking(obj)
    return obj


def summarize_report(report: dict[str, Any], indicators: dict) -> str:
    """提炼报告摘要。

    【2026-10 加固】摘要同样会编数字,且摘要位置最显眼(报告顶部),
    一旦出现"营业收入 12.86 亿元"这类假数字,危害大于正文。
    这里给三道约束:
        1) 明确要求"只能引用指标 JSON 里出现过的数字";
        2) 不确定就写定性描述,不要凑数字;
        3) 生成后由调用方(analyzer)做溯源复核,这里先做一次自查。
    """
    user = (
        "请基于以下报告 JSON 提炼 3 句中文摘要（每句不超过 30 字）。\n"
        f"报告：{json.dumps(report, ensure_ascii=False, default=json_default)[:6000]}\n"
        f"指标：{json.dumps(indicators, ensure_ascii=False, default=json_default)[:3000]}\n"
        "【硬性要求】\n"
        "  - 摘要里出现的每一个数字,**必须**能在上面的「指标」JSON 中找到原值;\n"
        "    严禁把万元换算成亿元、严禁编造指标里没有的字段(如费用率、现金流、\n"
        "    业务线占比等)——源数据没有的维度,就不要在摘要里提数字;\n"
        "  - 若某个维度指标里没有,直接用定性表述(如「盈利能力承压」),\n"
        "    **宁可没有数字,也不能有错数字**;\n"
        "  - 仅返回摘要文本,不要其他内容。"
    )
    # Reasoning 模型会在开头输出 <think>...;此处统一剥离,
    # 避免把"思考过程"暴露给前端(原 bug:摘要卡里出现大段英文思考链)
    text = strip_thinking(llm.chat(SYSTEM_PROMPT, user, json_mode=False).strip())
    return _normalize_summary_text(text)


def _normalize_summary_text(raw: str) -> str:
    """把模型返回的摘要规整成纯文本。

    【为什么需要】实测 deepseek-flash 经常无视"仅返回摘要文本"的指令,
    直接返回 JSON:{"摘要":["全年销售额 4680 万元…","客单价 826.72 元…"]}
    前端把它当纯文本渲染,报告顶部就会出现一大坨 JSON 原文。
    这里做兼容提取:
        1) 是 JSON → 取出文本数组/对象里的字符串,拼成多行纯文本;
        2) 抽出后仍带 <think> 等标签 → 再剥一次;
        3) 解析失败 → 原样返回(不猜、不丢内容)。
    """
    if not raw:
        return ""
    s = raw.strip()
    # 快速判断:既不是 JSON 对象/数组,也没有多余包裹,直接返回
    if not (s.startswith("{") or s.startswith("[")):
        return s
    try:
        data = _safe_json_loads(s)
    except Exception:
        return s
    if data is None:
        return s

    parts: list[str] = []

    def collect(node: Any) -> None:
        """递归收集文本片段:优先常见的 摘要/summary/content/text 字段。"""
        if isinstance(node, str):
            t = strip_thinking(node).strip()
            if t:
                parts.append(t)
            return
        if isinstance(node, list):
            for it in node:
                collect(it)
            return
        if isinstance(node, dict):
            # 优先按语义字段取值,顺序固定避免输出乱序
            for key in ("摘要", "summary", "content", "text", "简报"):
                if key in node:
                    collect(node[key])
                    return
            for v in node.values():
                collect(v)

    collect(data)
    if parts:
        return "\n".join(parts)
    return s


def annotate_anomalies(report: dict[str, Any], indicators: dict) -> list[dict[str, Any]]:
    """生成异常指标批注列表。

    【2026-10 加固】批注里的 value 字段会直接显示给用户,
    模型若编造数值会造成"数据看板出现不存在的指标值"。
    因此:value 字段只允许填指标 JSON 里真实存在的数值,
    无法确定时用 null,由前端渲染成「—」。
    """
    user = (
        "请基于以下指标 JSON，识别异常或值得关注的指标，"
        "返回 JSON 数组：[{metric, value, level: 'info'|'warn'|'danger', comment}]，最多 5 条。\n"
        f"指标：{json.dumps(indicators, ensure_ascii=False, default=json_default)[:6000]}\n"
        "【硬性要求】\n"
        "  - `value` 字段**必须**是上面「指标」JSON 里原样存在的数值,**严禁编造或换算**;\n"
        "    找不到对应数值时填 null,不要猜一个看起来合理的数;\n"
        "  - `metric` 字段必须是指标里真实存在的指标名/字段名,\n"
        "    不要写指标 JSON 里不存在的科目(如「期间费用率」,若源数据没有就不得出现);\n"
        "  - `comment` 里如需引用数字,同样必须来自指标 JSON。\n"
        "  - `comment` **最多 40 字**,只写\"是什么异常 + 一句判断\",\n"
        "    不要写行动建议、不要写多段分析(那些属于报告正文);\n"
        "  仅返回 JSON 数组,不要其他内容。"
    )
    raw = llm.chat(SYSTEM_PROMPT, user, json_mode=True)
    try:
        data = json.loads(raw)
        items: list[dict[str, Any]] = []
        if isinstance(data, list):
            items = data
        elif isinstance(data, dict) and "annotations" in data:
            items = data["annotations"]
        # 同样对 comment 做一次脱敏,避免 reasoning 模型把思考块塞进批注
        return _strip_thinking_recursive(items) if isinstance(items, list) else []
    except Exception:
        pass
    return []


# ---------------------- 数据校验解读 ----------------------

def interpret_validation(quality: dict, column_types: dict, dataset_name: str = "") -> dict[str, Any]:
    """把 Pandas 基础质量报告交给 LLM,生成自然语言解读。

    输入:
        quality: basic_quality() 的输出,含 row_count / missing_total / missing_by_column / duplicate_rows
        column_types: detect_column_types() 的输出,字段名 → number/date/text
        dataset_name: 数据集名称(用于报告中点名)

    输出:
        {
            "summary": 一句话总览(质量等级 + 主要问题),
            "level": "good" | "warn" | "danger",
            "highlights": [亮点列表],
            "issues": [
                {field, severity, snippet}
            ],
            "suggestions": [处理建议列表],
            "fallback": bool,  # True 表示 LLM 不可用,使用了离线解读
        }
    """
    # 组装 LLM 输入数据(精简,避免超出 token)
    miss_top = sorted(
        [(k, v) for k, v in (quality.get("missing_by_column") or {}).items() if v > 0],
        key=lambda x: -x[1],
    )[:5]

    user = (
        f"数据集名称:{dataset_name or '未命名'}\n"
        f"总行数:{quality.get('row_count', 0)}\n"
        f"总列数:{quality.get('column_count', 0)}\n"
        f"缺失值总数:{quality.get('missing_total', 0)}\n"
        f"重复行数:{quality.get('duplicate_rows', 0)}\n"
        f"字段类型:{json.dumps(column_types, ensure_ascii=False)}\n"
        f"缺失值 Top5 列:{json.dumps(miss_top, ensure_ascii=False)}\n\n"
        "请基于以上数据,扮演严谨但易读的数据分析师,输出 JSON:\n"
        "{\n"
        '  "summary": "一句话总览(20字内,如:整体良好 / 存在空值与重复 / 数据质量较差)",\n'
        '  "level": "good | warn | danger",\n'
        '  "highlights": ["数据完整性方面的亮点1", "亮点2"],\n'
        '  "issues": [{"field": "字段名或整体", "severity": "low|medium|high", "snippet": "这条字段的什么问题"}],\n'
        '  "suggestions": ["建议1:具体清洗动作", "建议2:..."]\n'
        "}\n"
        "判断标准:\n"
        "  - missing_total==0 且 duplicate_rows==0 → good\n"
        "  - missing_total/row_count < 10% → warn\n"
        "  - missing_total/row_count ≥ 10% 或 多个列缺失严重 → danger\n"
        "不要编造数据,所有判断基于上面提供的数字。"
    )

    fallback = False
    error_msg = ""
    try:
        # 先按 json_mode=True 试一次(模型支持时更稳)
        raw = ""
        parse_ok = False
        try:
            raw = llm.chat(SYSTEM_PROMPT, user, json_mode=True)
            data = json.loads(raw)
            if not isinstance(data, dict):
                raise ValueError("LLM 返回的不是字典")
            parse_ok = True
        except (json.JSONDecodeError, ValueError) as first_exc:
            # 第一次解析失败:可能是模型不支持 response_format=json_object,
            # 也可能是纯文本回答。再以"自由输出"重试一次,让模型自己控制格式
            logger.info("interpret_validation: JSON 模式解析失败(%s),尝试自由模式重试", first_exc)
            try:
                raw = llm.chat(SYSTEM_PROMPT, user, json_mode=False)
                # 优先从 raw 中抽出 {...} 段
                m = re.search(r"\{[\s\S]*\}", raw or "")
                if m:
                    data = json.loads(m.group(0))
                    if isinstance(data, dict):
                        parse_ok = True
            except Exception as second_exc:
                error_msg = f"JSON 模式失败: {first_exc}; 自由模式失败: {second_exc}"
                logger.warning("interpret_validation 重试仍失败: %s", error_msg)
                raise

        if parse_ok and isinstance(data, dict):
            return {
                "summary": data.get("summary", ""),
                "level": data.get("level", "warn"),
                "highlights": data.get("highlights", []) or [],
                "issues": data.get("issues", []) or [],
                "suggestions": data.get("suggestions", []) or [],
                "fallback": False,
                "error": "",
            }
        # 走到这里说明两次都没拿到 dict
        raise ValueError(error_msg or "两次调用都未拿到 JSON 字典")
    except Exception as exc:
        # LLM 不可用 → 用规则化解读兜底,保证 UI 总有内容
        # 用 logger 打到后端终端,便于排错
        error_msg = str(exc)
        logger.warning("interpret_validation LLM 调用失败,降级为规则解读: %s", exc)
        return {
            **rule_based_interpret(quality, column_types),
            "fallback": True,
            "error": error_msg,
        }


def rule_based_interpret(quality: dict, column_types: dict) -> dict[str, Any]:
    """LLM 不可达时的规则化解读兜底。"""
    row_count = max(1, quality.get("row_count", 0))
    missing = quality.get("missing_total", 0)
    dup = quality.get("duplicate_rows", 0)
    miss_ratio = missing / (row_count * max(1, quality.get("column_count", 0)))
    dup_ratio = dup / row_count

    # 评级
    if missing == 0 and dup == 0:
        level = "good"
        summary = "数据完整,无明显质量问题"
    elif miss_ratio < 0.1 and dup_ratio < 0.05:
        level = "warn"
        summary = "存在少量缺失或重复,建议清洗"
    else:
        level = "danger"
        summary = "缺失或重复较多,需重点处理"

    # 亮点
    highlights = []
    if missing == 0:
        highlights.append("全部字段无空值,数据完整")
    if dup == 0:
        highlights.append("未发现重复行,记录唯一")
    if not highlights:
        highlights.append(f"共 {quality.get('column_count', 0)} 个字段、{row_count} 条记录")

    # 问题
    issues = []
    for col, cnt in sorted(
        (quality.get("missing_by_column") or {}).items(),
        key=lambda x: -x[1],
    )[:5]:
        if cnt > 0:
            sev = "high" if cnt / row_count > 0.3 else ("medium" if cnt / row_count > 0.1 else "low")
            issues.append({
                "field": col,
                "severity": sev,
                "snippet": f"该列有 {cnt} 个空值({round(cnt/row_count*100, 1)}%)",
            })
    if dup > 0:
        issues.append({
            "field": "整体",
            "severity": "high" if dup_ratio > 0.1 else "medium",
            "snippet": f"存在 {dup} 条重复行({round(dup_ratio*100, 1)}%)",
        })

    # 建议
    suggestions = []
    if missing > 0:
        suggestions.append("对数值列缺失值,可用均值/中位数填充;文本列用众数或前向填充")
    if dup > 0:
        suggestions.append("使用「去除重复」规则一键去重")
    if not suggestions:
        suggestions.append("数据质量良好,可直接进入指标计算与 AI 报告阶段")

    return {
        "summary": summary,
        "level": level,
        "highlights": highlights,
        "issues": issues,
        "suggestions": suggestions,
    }


# ---------------------- 自然语言查询 ----------------------

def nl_to_pandas_query(question: str, columns: list[str], df_head: list[list]) -> dict[str, Any]:
    """让 LLM 把中文问题转成一段 Python/Pandas 计算说明（不执行，只产出步骤）。"""
    user = (
        f"用户问题：{question}\n"
        f"可用列：{columns}\n"
        f"前 5 行示例：{df_head}\n\n"
        "请将问题拆解为可执行步骤：\n"
        "1. 涉及哪些列；2. 过滤条件；3. 聚合方式（sum/mean/count/max/min/sort）；4. 返回前 5 行。\n"
        "仅返回 JSON：{ops: [...], explanation: '...'}，不要写代码。"
    )
    # LLM 实际返回的是字符串,需 json.loads 解析为字典;
    # 解析失败也要返回空字典而不是字符串,避免调用方 .get() 报 'str' has no attribute 'get'
    # 【deepen-data-chat-analysis】LLM 调用本身也可能失败(网络/额度/超时)。
    # plan 只是辅助展示字段,拿不到不该让整个降级链路崩到接口 500 ——
    # 这里捕获后返回空 plan,调用方继续用 Pandas 规则分支作答。
    try:
        raw = llm.chat(SYSTEM_PROMPT, user, json_mode=True)
    except Exception as exc:
        logger.warning("nl_to_pandas_query LLM 调用失败,plan 置空继续降级作答: %s", exc)
        return {"ops": [], "explanation": ""}
    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            return data
    except Exception:
        # 从 raw 中尝试抽取 JSON 片段(兼容模型在 JSON 外加说明文字的情况)
        try:
            start, end = raw.find("{"), raw.rfind("}")
            if start != -1 and end > start:
                data = json.loads(raw[start: end + 1])
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {"ops": [], "explanation": ""}


def execute_nl_query(
    question: str, file_path: str, columns_meta: list[dict],
    sheets_info: list[dict] | None = None,
    history: list[dict] | None = None,
) -> dict[str, Any]:
    """根据中文问题返回结果表与文字说明。

    【两阶段管道 2026-10】
      数据画像 → LLM 产出受限计算指令 → Pandas 全量真实执行 → LLM 基于真实计算
      结果写结构化深度叙述(structured 字段)。旧关键词硬编码意图逻辑**整体保留**
      为 _execute_legacy_keyword_query:管道任一步失败(规划为空/执行异常/叙述
      解析失败)时降级走旧逻辑,不删不弃。

    响应结构:
      - 旧字段(columns/rows/explanation/sheets/sheet_results/plan/active_sheet_count)
        全部保持;新增 structured:
          {"answer": str, "analysis": [str],
           "data_table": {"columns": [...], "rows": [[...]]},
           "suggestions": [str], "followups": [str]}
        管道失败时 structured=None,前端旧渲染不受影响。

    参数:
      history: 多轮对话历史 [{role:"user"|"ai", content:str}],由 API 层清洗后传入,
               供叙述阶段理解上下文(最多取最近 5 条)。
    """
    # 1. 按 sheet 独立读取(与旧版一致)
    per_sheets = pe.read_dataframe_per_sheet(file_path)
    if not per_sheets:
        return {
            "columns": ["字段", "值"],
            "rows": [],
            "explanation": "数据集为空或无法识别工作表",
            "sheets": [],
            "sheet_results": [],
            "structured": None,
        }

    # —— 两阶段管道(任何异常都降级,保证接口永不 500) ——
    pipeline_ops: list[dict] = []
    sheet_computations: list[dict] = []
    structured: dict[str, Any] | None = None
    try:
        # 阶段 1:逐 sheet 生成轻量画像 —— 这是 LLM 规划时"唯一可信的数据视野",
        # 全量统计(min/max/sum/去重值)替代旧方案的 head(20),杜绝大表聚合靠猜
        profiles: list[dict] = []
        for entry in per_sheets:
            try:
                prof = pe.build_sheet_profile(entry["df"])
            except Exception as exc:
                # 单 sheet 画像失败不致命:跳过该 sheet,其余照常
                logger.warning(f"[nl_query] sheet「{entry['name']}」画像失败,已跳过: {exc}")
                continue
            prof["sheet"] = entry["name"]  # 画像内补 sheet 名,LLM 规划时按 sheet 定位
            profiles.append(prof)

        # 阶段 2:LLM 规划受限计算指令(只产出 ops 参数字典,绝不产出代码)
        if profiles:
            pipeline_ops = _plan_computation_ops(question, profiles)

        # 阶段 3:Pandas 白名单执行器对全量数据真实执行(按 sheet 分组、流水线执行)
        if pipeline_ops:
            ops_by_sheet: dict[str, list[dict]] = {}
            for item in pipeline_ops:
                # 传给执行器前剥掉 "sheet" 字段(执行器只认 sheet 无关的 op 参数)
                ops_by_sheet.setdefault(str(item.get("sheet") or ""), []).append(
                    {k: v for k, v in item.items() if k != "sheet"}
                )
            for entry in per_sheets:
                s_ops = ops_by_sheet.get(entry["name"]) or []
                if not s_ops:
                    continue
                comp = pe.execute_computation_ops(entry["df"], s_ops)
                comp["sheet"] = entry["name"]  # 执行结果标回来源 sheet,叙述与 sheet_results 都要用
                sheet_computations.append(comp)

        # 阶段 4:LLM 基于真实计算结果写结构化深度叙述
        # (多个 sheet 的执行结果合并喂给叙述,让模型能给出跨表结论)
        if sheet_computations:
            structured = _narrate_deep_answer(
                question, profiles, {"sheets": sheet_computations}, history
            )
    except Exception as exc:
        # 管道整体异常 → 静默降级到旧逻辑(旧逻辑自身有兜底)
        logger.warning(f"[nl_query] 两阶段管道异常,降级旧关键词逻辑: {exc}")
        structured = None

    if structured is not None:
        # —— 成功路径:沿用旧结构骨架,内容换成真实计算产物(不再跑旧逻辑,省 LLM 调用) ——
        first_df = per_sheets[0]["df"]
        # 顶层 rows/columns:优先取首个非空计算表(全量真算结果,优于 head 样本);
        # 计算表全空(如 filter 没命中)→ 沿用旧版习惯给首 sheet 前 5 行样本
        top_cols: list = []
        top_rows: list = []
        for comp in sheet_computations:
            tb = comp.get("table") or {}
            if tb.get("rows"):
                top_cols = list(tb.get("columns") or [])
                top_rows = list(tb.get("rows") or [])
                break
        if not top_rows:
            top_cols = [str(c) for c in first_df.columns]
            top_rows = first_df.head(5).astype(object).fillna("").values.tolist()
        # sheet_results 从真实计算结果生成(每步 desc 汇总为 summary,供前端按 sheet 展示)
        sheet_results: list[dict] = []
        for entry in per_sheets:
            s_name = entry["name"]
            comp = next((c for c in sheet_computations if c.get("sheet") == s_name), None)
            if comp:
                tb = comp.get("table") or {}
                steps_desc = "; ".join(
                    str(s.get("desc") or "") for s in (comp.get("steps") or []) if s.get("desc")
                )
                sheet_results.append({
                    "sheet": s_name,
                    "columns": list(tb.get("columns") or []),
                    "rows": list(tb.get("rows") or []),
                    "summary": steps_desc or "已完成受限计算",
                })
            else:
                sheet_results.append({
                    "sheet": s_name, "rows": [], "columns": [],
                    "summary": "本次计算未涉及该表",
                })
        # plan 字段沿用旧契约(保持 dict 含 ops/explanation),内容换成新管道的指令
        plan = {
            "ops": pipeline_ops,
            "explanation": "两阶段管道:数据画像 → 受限计算指令 → Pandas 全量执行 → 结构化叙述",
        }
        # explanation 取叙述的 answer:前端旧渲染位(纯文本摘要)也能看到关键结论
        explanation = str(structured.get("answer") or "")
    else:
        # —— 降级路径:完整保留旧关键词逻辑(含其内部 LLM 调用),结构一字不改 ——
        legacy = _execute_legacy_keyword_query(question, per_sheets, columns_meta, sheets_info)
        legacy["structured"] = None
        return legacy

    # 4. explanation 附多 sheet 明细(旧版逻辑,成功路径与降级路径都保留)
    if sheets_info and len(sheets_info) > 1:
        names = "、".join(s["name"] for s in sheets_info[:8])
        more = "" if len(sheets_info) <= 8 else f" 等共 {len(sheets_info)} 张"
        explanation = (
            explanation
            + f"\n(本次查询按 sheet 分别执行,共 {len(sheets_info)} 张:{names}{more},"
              f"每张表的独立结果见 sheet_results 字段)"
        )

    return {
        "columns": top_cols,
        "rows": top_rows,
        "explanation": explanation,
        "sheets": sheets_info or [],
        "sheet_results": sheet_results,
        "plan": plan,
        "active_sheet_count": len(sheets_info) if sheets_info else len(per_sheets),
        "structured": structured,
    }


# —— 受限计算指令白名单(与 pandas_engine.execute_computation_ops 的算子集保持一致) ——
_NL_PIPELINE_OPS = ("filter", "groupby_agg", "sort", "topn", "describe")


def _plan_computation_ops(question: str, profiles: list[dict]) -> list[dict]:
    """阶段 2:让 LLM 把问题 + 数据画像规划成「受限计算指令」(ops 列表)。

    安全设计:
      - LLM 只产出参数字典,真正计算由 pandas_engine.execute_computation_ops 的
        白名单算子完成,全链路无 eval/exec/任意代码;
      - 本函数再做一次收口过滤:op 不在白名单 / sheet 名与画像对不上的项直接丢弃,
        与执行器内部的逐项防御形成双重防线。

    失败策略:LLM 调用或 JSON 解析失败返回 [] —— 调用方对空 ops 会降级旧关键词逻辑。
    """
    user = (
        "你是数据计算规划器。根据用户问题和各 sheet 的数据画像,产出一段「受限计算指令」。\n"
        "只能使用画像中出现过的真实列名,严禁编造列名或数值。\n\n"
        "可用算子(仅以下 5 种):\n"
        '1. filter      : {"op":"filter","column":"列名","operator":"==|!=|>|>=|<|<=|contains","value":值}\n'
        '2. groupby_agg : {"op":"groupby_agg","group":["分组列"],"agg":{"数值列":"sum|mean|count|max|min"}}\n'
        '3. sort        : {"op":"sort","by":"列名","ascending":true或false}\n'
        '4. topn        : {"op":"topn","by":"数值列","n":整数,"ascending":false}   # ascending 默认 false=取最大前 N\n'
        '5. describe    : {"op":"describe"}   # 数值列统计摘要,适合"概览/分布"类问题\n\n'
        "规则:\n"
        "- ops 数组顺序 = 执行顺序(流水线:filter → groupby_agg → sort → topn);\n"
        '- 每条指令必须带 "sheet" 字段指定在哪张表上执行(sheet 名用画像里的真实名称);\n'
        "- 占比/份额类问题:先用 groupby_agg 对相关数值列求和即可,"
        "占比计算交给后续叙述阶段基于聚合结果完成,不要自己算;\n"
        "- 一般问题 1~4 条指令足够,不要堆砌。\n\n"
        '仅返回 JSON(不要代码、不要解释):{"ops": [{"sheet": "sheet名", "op": "...", ...}, ...]};'
        '若问题与数据无关或无法规划,返回 {"ops": []}。\n\n'
        f"【用户问题】{question}\n\n"
        f"【各 sheet 数据画像】\n{json.dumps(profiles, ensure_ascii=False)}"
    )
    try:
        # temperature=0.1 + max_tokens=1500:规划要稳定,指令体量也小
        raw = llm.chat(SYSTEM_PROMPT, user, json_mode=True, temperature=0.1, max_tokens=1500)
    except Exception as exc:
        logger.warning(f"[nl_query] 规划计算指令失败,将降级旧逻辑: {exc}")
        return []
    # 先剥思考前缀再解析(reasoning 模型可能裹 <think>);解析失败再抽 JSON 段兜底
    text = strip_thinking(raw or "")
    data: Any = None
    try:
        data = json.loads(text)
    except Exception:
        try:
            data = _safe_json_loads(text)
        except Exception:
            data = None
    raw_ops = data.get("ops") if isinstance(data, dict) else None
    if not isinstance(raw_ops, list):
        return []
    # —— 收口过滤:op 必须在白名单内,sheet 名必须出现在画像中 ——
    valid_sheets = {str(p.get("sheet") or "") for p in profiles}
    # 单 sheet 场景:LLM 常省略 sheet 字段,默认归到唯一 sheet,避免误降级
    default_sheet = str(profiles[0].get("sheet") or "") if len(profiles) == 1 else ""
    out: list[dict] = []
    for item in raw_ops:
        if not isinstance(item, dict):
            continue
        if str(item.get("op") or "").strip().lower() not in _NL_PIPELINE_OPS:
            logger.warning(f"[nl_query] 规划指令 op「{item.get('op')}」不在白名单,已丢弃")
            continue
        sheet = str(item.get("sheet") or "").strip()
        if not sheet and default_sheet:
            sheet = default_sheet
            item["sheet"] = sheet
        if sheet not in valid_sheets:
            # sheet 名对不上 → 丢弃(防止指令落到不存在的表上)
            logger.warning(f"[nl_query] 规划指令 sheet「{sheet}」与画像不匹配,已丢弃")
            continue
        out.append(item)
    return out


def _nl_narr_cell(v: Any) -> Any:
    """叙述 data_table 单元格规整:可序列化标量原样保留,NaN/Inf→"",复杂类型转 str。

    为什么需要:json.loads 会接受 NaN/Infinity 字面量,且 LLM 可能把单元格
    塞成 list/dict —— 逐格规整保证 data_table 永远可 JSON 序列化。
    """
    if isinstance(v, float):
        return "" if (v != v or v in (float("inf"), float("-inf"))) else v
    if v is None:
        return ""
    if isinstance(v, (str, int, bool)):
        return v
    if isinstance(v, (list, dict)):
        return json.dumps(v, ensure_ascii=False)
    return str(v)


def _default_followups(profiles: list[dict]) -> list[str]:
    """followups 为空时,基于画像自动补默认追问(保证前端"推荐追问"区不空)。"""
    out: list[str] = []
    for p in profiles[:2]:  # 最多看 2 张表,足够拼 3 条追问
        prefix = f"「{p.get('sheet')}」表中" if p.get("sheet") else ""
        cols = p.get("columns") or []
        cat_col = next(
            (c for c in cols if c.get("type") == "text" and c.get("top5")), None
        )
        num_col = next(
            (c for c in cols if c.get("type") == "number" and c.get("stats")), None
        )
        if cat_col:
            out.append(f"{prefix}「{cat_col['name']}」各类别的 Top5 分布是怎样的？")
        if num_col and cat_col:
            out.append(f"{prefix}按「{cat_col['name']}」分组时,「{num_col['name']}」的合计是多少？")
        elif num_col:
            out.append(f"{prefix}「{num_col['name']}」的最大值、最小值和均值分别是多少？")
        if len(out) >= 3:
            break
    if not out:
        # 画像没给出可用列时的通用兜底
        out = ["各分类列的 Top5 分布是怎样的？", "数值列的合计与均值分别是多少？"]
    return out[:3]


def _narrate_deep_answer(
    question: str,
    profiles: list[dict],
    computation: dict,
    history: list[dict] | None = None,
) -> dict[str, Any] | None:
    """阶段 4:把「真实计算结果」+ 数据画像 + 对话历史喂给 LLM,写结构化深度叙述。

    与旧实现的本质区别:所有数字来自 Pandas 全量真算(computation),LLM 只做
    解读,不允许自己算数 —— 从根上解决"head(20) 猜全表"的幻觉问题。

    返回:{"answer": str, "analysis": [str],
          "data_table": {"columns": [...], "rows": [[...]]},
          "suggestions": [str], "followups": [str](最多 3 条)}
    任何异常(调用失败/JSON 解析失败/结构不对)返回 None,由调用方降级旧逻辑。
    """
    try:
        # —— 组装「最近对话」段(最多 5 条),帮助模型理解指代与上下文 ——
        hist_lines: list[str] = []
        for h in (history or [])[-5:]:
            if not isinstance(h, dict):
                continue
            role = "用户" if h.get("role") == "user" else "助手"
            content = str(h.get("content") or "").strip()[:300]  # 单条截断,防历史超长
            if content:
                hist_lines.append(f"{role}: {content}")
        hist_block = "\n".join(hist_lines) if hist_lines else "(无历史对话)"

        # —— 组装画像与计算结果文本(超长截断,控制 prompt 体积) ——
        prof_text = json.dumps(profiles, ensure_ascii=False)
        if len(prof_text) > 8000:
            prof_text = prof_text[:8000] + "...(画像过长已截断)"
        # json_default 兜底:计算结果理论上已纯 Python 化,防御性再兜一层
        comp_text = json.dumps(computation, ensure_ascii=False, default=json_default)
        if len(comp_text) > 12000:
            comp_text = comp_text[:12000] + "...(计算结果过长已截断)"

        user = (
            "你是资深数据分析师。请基于「真实计算结果」回答用户当前问题,输出结构化深度叙述。\n\n"
            "【硬性要求】\n"
            "- 所有数字只能来自「数据画像」或「真实计算结果」,严禁编造或自行推算表中没有的数;\n"
            "- 计算结果为空或含 error 时,如实向用户说明,不要臆测;\n"
            "- answer 必须是含关键数字的直接结论(1~2 句),不要空话;\n"
            "- analysis 给 2~5 条分点分析(趋势/结构/异常/风险视角);\n"
            "- data_table 从真实计算结果中选最能回答问题的表原样引用(不要改数字,最多 10 行);"
            '计算结果为空时给空表 {"columns": [], "rows": []};\n'
            "- suggestions 给 2~3 条基于数据的可执行建议;\n"
            "- followups 给 3 条用户可能继续追问的问题;\n"
            "- 全部用中文。\n\n"
            f"【最近对话】\n{hist_block}\n\n"
            f"【用户当前问题】{question}\n\n"
            f"【数据画像】\n{prof_text}\n\n"
            f"【真实计算结果】\n{comp_text}\n\n"
            "仅返回 JSON:\n"
            '{"answer": "...", "analysis": ["...", ...], '
            '"data_table": {"columns": ["..."], "rows": [["..."]]}, '
            '"suggestions": ["..."], "followups": ["...", "...", "..."]}'
        )
        # 【2026-10 加固】reasoning 模型偶发返回空正文(实测约 1/10 概率),
        # 若直接解析空串会抛 "empty input" 并掉到旧逻辑,用户体验断崖式下降。
        # 只重试一次:仍失败才交给调用方降级,避免放大延时。
        data: Any = None
        for _attempt in range(2):
            raw = llm.chat(SYSTEM_PROMPT, user, json_mode=False, temperature=0.2, max_tokens=3000)
            # 叙述输出必须过 strip_thinking:剥掉 reasoning 模型的思考前缀再解析
            text = strip_thinking(raw or "").strip()
            if not text:
                logger.warning("[nl_query] 深度叙述返回空正文,重试一次")
                continue
            try:
                data = json.loads(text)
            except Exception:
                # 兜底抽取 JSON 段(模型可能在 JSON 外带说明文字);再失败则下一轮/返回 None
                try:
                    data = _safe_json_loads(text)
                except Exception:
                    data = None
            if isinstance(data, dict):
                break
            data = None
        if not isinstance(data, dict):
            return None

        # —— 归一化校验:保证字段类型与前端契约一致,脏数据逐项防御 ——
        answer = str(data.get("answer") or "").strip()
        analysis = [str(x).strip() for x in (data.get("analysis") or []) if str(x or "").strip()]
        suggestions = [
            str(x).strip() for x in (data.get("suggestions") or []) if str(x or "").strip()
        ]
        followups = [
            str(x).strip() for x in (data.get("followups") or []) if str(x or "").strip()
        ][:3]
        # data_table:columns 转 str 列表;rows 只保留 list 行并逐格规整(最多 10 行)
        dt = data.get("data_table") if isinstance(data.get("data_table"), dict) else {}
        columns = [str(c) for c in (dt.get("columns") or [])]
        rows: list[list] = []
        for r in (dt.get("rows") or [])[:10]:
            if isinstance(r, list):
                rows.append([_nl_narr_cell(v) for v in r])
        # answer 缺失时用第一条分析兜底;answer/analysis/rows 全空才视为叙述失败
        if not answer:
            answer = analysis[0] if analysis else ""
        if not answer and not analysis and not rows:
            return None
        # followups 为空 → 基于画像自动补默认追问
        if not followups:
            followups = _default_followups(profiles)

        return {
            "answer": answer,
            "analysis": analysis,
            "data_table": {"columns": columns, "rows": rows},
            "suggestions": suggestions,
            "followups": followups,
        }
    except Exception as exc:
        logger.warning(f"[nl_query] 深度叙述失败,将降级旧逻辑: {exc}")
        return None


def _execute_legacy_keyword_query(
    question: str,
    per_sheets: list[dict],
    columns_meta: list[dict],
    sheets_info: list[dict] | None = None,
) -> dict[str, Any]:
    """【整体保留的旧实现】关键词硬编码意图分支(两阶段管道失败时的 fallback)。

    抽取自原 execute_nl_query 的关键词段,执行逻辑一字未改,仅移入私有函数:
      - 意图分支:最大/最小/总和/平均/行数/通用查找(LLM 抽过滤条件)/LLM 直接兜底;
      - plan 由 nl_to_pandas_query 生成但从不执行(历史行为,保留);
      - 多 sheet 时 explanation 附注 sheet 明细。
    入参 per_sheets 由 execute_nl_query 统一读取后传入,避免重复读盘。
    """
    # 2. 全局列名(用于 plan)
    all_cols = [c["name"] for c in columns_meta]
    # plan 复用首 sheet 的 head(简化实现:plan 是基于"我们有什么字段"的总览,不分 sheet)
    first_df = per_sheets[0]["df"]
    head = first_df.head(3).astype(object).fillna("").values.tolist() if len(first_df) > 0 else []
    plan = nl_to_pandas_query(question, all_cols, head)

    # 3. 逐 sheet 执行查询逻辑(规则版,LLM 给 plan 但执行端是 Pandas)
    q = question.lower()
    sheet_results: list[dict] = []
    rows: list[list] = []
    explanation = ""
    try:
        for entry in per_sheets:
            s_name = entry["name"]
            df = entry["df"]
            if len(df) == 0:
                sheet_results.append({
                    "sheet": s_name, "rows": [], "columns": [],
                    "summary": f"工作表「{s_name}」为空",
                })
                continue
            # 当前 sheet 内部的 numeric_cols(基于 columns_meta 中 [sheet] 前缀匹配)
            s_numeric_cols = [c["name"] for c in columns_meta
                               if c["type"] == "number" and c["name"].startswith(f"[{s_name}] ")]
            # 也接受"未带前缀"的旧 columns_meta(单 sheet 场景)
            if not s_numeric_cols:
                s_numeric_cols = [c["name"] for c in columns_meta if c["type"] == "number"
                                   and "[" not in c["name"]]
            s_date_cols = [c["name"] for c in columns_meta
                            if c["type"] == "date" and c["name"].startswith(f"[{s_name}] ")]
            if not s_date_cols:
                s_date_cols = [c["name"] for c in columns_meta if c["type"] == "date"
                                and "[" not in c["name"]]

            local_rows: list = []
            local_note = ""
            # —— 各意图分支 ——
            if any(k in q for k in ["最大", "最高", "最多", "max"]):
                if s_numeric_cols:
                    first_num = s_numeric_cols[0]
                    # 把 "[sheet] 列名" 转回真实列名(去掉前缀)
                    real_col = first_num[len(f"[{s_name}] "):] if first_num.startswith(f"[{s_name}] ") else first_num
                    if real_col in df.columns:
                        idx = df[real_col].idxmax()
                        local_rows = df.loc[[idx]].astype(object).fillna("").values.tolist()
                        local_note = f"{real_col} 最大值出现在第 {int(idx) + 1} 行。"
            elif any(k in q for k in ["最小", "最低", "最少", "min"]):
                if s_numeric_cols:
                    first_num = s_numeric_cols[0]
                    real_col = first_num[len(f"[{s_name}] "):] if first_num.startswith(f"[{s_name}] ") else first_num
                    if real_col in df.columns:
                        idx = df[real_col].idxmin()
                        local_rows = df.loc[[idx]].astype(object).fillna("").values.tolist()
                        local_note = f"{real_col} 最小值出现在第 {int(idx) + 1} 行。"
            elif "总和" in q or "合计" in q or "总" in q:
                if s_numeric_cols:
                    real_cols = [c[len(f"[{s_name}] "):] if c.startswith(f"[{s_name}] ") else c
                                  for c in s_numeric_cols]
                    real_cols = [c for c in real_cols if c in df.columns]
                    if real_cols:
                        total = df[real_cols].sum().to_dict()
                        local_rows = [[k, float(round(v, 2))] for k, v in total.items()]
                        local_note = f"各数值列求和:{total}"
            elif "平均" in q:
                if s_numeric_cols:
                    real_cols = [c[len(f"[{s_name}] "):] if c.startswith(f"[{s_name}] ") else c
                                  for c in s_numeric_cols]
                    real_cols = [c for c in real_cols if c in df.columns]
                    if real_cols:
                        mean = df[real_cols].mean().to_dict()
                        local_rows = [[k, float(round(v, 2))] for k, v in mean.items()]
                        local_note = f"各数值列均值:{mean}"
            elif "行数" in q or "多少行" in q or "count" in q:
                local_rows = [["总行数", int(len(df))]]
                local_note = f"数据共 {len(df)} 行。"
            elif any(k in q for k in ["多少", "哪", "哪个", "查找", "查", "求", "呢", "?", "？"]):
                # 【2026-10 修复】通用查找意图:让 LLM 抽 (filter_col, filter_value),
                # 再 .loc 取行,而不是硬编码关键词。常见问法如"产品A 的销售额是多少"
                # /"查一下产品B 的订单" /"产品C 达成率"等。
                answer_text = ""
                hit_rows: list = []
                try:
                    cols_hint = ", ".join(str(c) for c in df.columns)
                    plan_user = (
                        "请从用户问题中提取过滤条件,用于定位表格中的某一行/若干行。\n"
                        f"可用列名:{cols_hint}\n用户问题:{question}\n\n"
                        '仅返回合法 JSON:{"filter_col":"列名","filter_value":"值"};'
                        '若问题不指向具体过滤(纯开放性问题),返回:{"filter_col":"","filter_value":""}。'
                    )
                    raw_plan = llm.chat(SYSTEM_PROMPT, plan_user, json_mode=True, temperature=0.1)
                    fc, fv = "", ""
                    try:
                        m_obj = json.loads(raw_plan)
                        if isinstance(m_obj, dict):
                            fc = (m_obj.get("filter_col") or "").strip()
                            fv = (m_obj.get("filter_value") or "").strip()
                    except Exception:
                        # 抽 JSON 段再试一次,兼容 reasoning 模型
                        from app.core.json_utils import _safe_json_loads as _sj
                        m_obj2 = _sj(raw_plan or "")
                        if isinstance(m_obj2, dict):
                            fc = (m_obj2.get("filter_col") or "").strip()
                            fv = (m_obj2.get("filter_value") or "").strip()
                    if fc and fv and fc in df.columns:
                        hit = df[df[fc].astype(str) == fv]
                        if len(hit) > 0:
                            hit_rows = hit.astype(object).fillna("").values.tolist()
                            local_note = f"按「{fc}={fv}」命中 {len(hit)} 行,详情见下方表格。"
                        else:
                            local_note = f"未找到 {fc}={fv} 的记录,已返回前 5 行作为参考。"
                    else:
                        local_note = "问题未明确指向单一行,已用 LLM 直接回答并附前 5 行样本。"
                except Exception as exc:
                    local_note = f"按列查找失败({exc}),已用 LLM 直接回答。"
                # 若 LLM 没给出 row 命中,降级为"LLM 自由回答 + 前 5 行样本"
                if not hit_rows:
                    try:
                        sample_md = df.head(20).to_markdown(index=False)
                    except Exception:
                        sample_md = df.head(20).astype(object).fillna("").to_string(index=False)
                    direct_user = (
                        f"你是数据分析助手。请基于下面「{s_name}」表的真实数据,直接回答用户问题。\n"
                        f"【要求】只能用表中真实数字回答;表里找不到答案请明确说「未在该表中找到」。\n\n"
                        f"用户问题:{question}\n\n{question}\n\n"
                        f"该表前 20 行:\n{sample_md}\n"
                    )
                    try:
                        # 【deepseek-flash 注意】这类模型有 reasoning_tokens,
                        # max_tokens 必须 >= 思考 + 答复,这里给 2000 留足余量
                        raw_ans = llm.chat(SYSTEM_PROMPT, direct_user, json_mode=False, temperature=0.2, max_tokens=2000)
                        answer_text = strip_thinking(raw_ans or "").strip()
                    except Exception as exc:
                        answer_text = f"AI 服务暂不可用:{exc}"
                    local_rows = df.head(5).astype(object).fillna("").values.tolist()
                    if answer_text:
                        # 把 LLM 回答拼到 summary 顶部,前端能看到具体数字答案
                        local_note = answer_text
                else:
                    local_rows = hit_rows
            else:
                # 【2026-10 兜底加固】之前的兜底直接返前 5 行 + 一句"无法精确识别",
                # 用户体感是"答非所问"。现在兜底直接调用 LLM,把整张表喂给模型作答。
                try:
                    sample_md = df.head(20).to_markdown(index=False)
                except Exception:
                    sample_md = df.head(20).astype(object).fillna("").to_string(index=False)
                direct_user = (
                    f"你是数据分析助手。请基于下面「{s_name}」表的真实数据,直接回答用户问题。\n"
                    f"【要求】只能用表中真实数字回答;表里找不到答案请明确说「未在该表中找到」。\n\n"
                    f"用户问题:{question}\n\n"
                    f"该表前 20 行:\n{sample_md}\n"
                )
                try:
                    raw_ans = llm.chat(SYSTEM_PROMPT, direct_user, json_mode=False, temperature=0.2, max_tokens=2000)
                    answer_text = strip_thinking(raw_ans or "").strip()
                except Exception as exc:
                    answer_text = f"AI 服务暂不可用:{exc}"
                local_rows = df.head(5).astype(object).fillna("").values.tolist()
                if answer_text:
                    local_note = answer_text
                else:
                    local_note = "未能从表中回答该问题,已返回前 5 行作为样本。"
            sheet_results.append({
                "sheet": s_name,
                "rows": local_rows,
                "columns": list(df.columns),
                "summary": local_note,
            })
            # 顶层 rows 取首 sheet 的(向后兼容)
            if s_name == per_sheets[0]["name"]:
                rows = local_rows
                explanation = local_note
    except Exception as exc:
        explanation = f"查询执行失败：{exc}"

    # 4. explanation 附多 sheet 明细(2026-10)
    if sheets_info and len(sheets_info) > 1:
        names = "、".join(s["name"] for s in sheets_info[:8])
        more = "" if len(sheets_info) <= 8 else f" 等共 {len(sheets_info)} 张"
        explanation = (
            explanation
            + f"\n(本次查询按 sheet 分别执行,共 {len(sheets_info)} 张:{names}{more},"
              f"每张表的独立结果见 sheet_results 字段)"
        )

    return {
        "columns": list(first_df.columns) if not rows else (list(first_df.columns) if isinstance(rows[0], list) and len(rows[0]) == len(first_df.columns) else ["字段", "值"]),
        "rows": rows,
        "explanation": explanation,
        "sheets": sheets_info or [],
        "sheet_results": sheet_results,
        "plan": plan,
        "active_sheet_count": len(sheets_info) if sheets_info else len(per_sheets),
    }


# ---------------------- 数据集对比 ----------------------
# 2 种表格 mode 共享的 ai_service 入口:按 mode 路由 pandas_engine 算子,
# 把 diff_details + indicators 喂给 LLM 生成「数据分析师视角」的中文叙述。
# LLM 失败时由 _fallback_narration_by_mode 兜底,保证前端 el-table + 叙述区始终有内容。
#
# 【topn 并入 business / distribution 下线 新增】
#   - 原 "topn" 独立模式已下线:排行只是 business 聚合结果的另一种视角,
#     拆开会让「别人插进来 → 全体名次下降」这类相对位移被误读为经营下滑。
#     现在 TopN 是 business 的可选截断参数(top_n),排名只作辅助参考列。
#   - 原 "distribution" 独立模式已下线:输出仅 5 个标量,描述不出分布形状,
#     且与单表侧 numeric_summary / describe 能力重叠。
_VALID_MODES = ("business", "row_diff")


def _call_narration_llm(mode: str, prompt: str) -> dict[str, Any]:
    """调 LLM 拿四段叙述;失败/解析失败 → 返回 fallback 占位 dict。

    为什么单独抽出来:不抛异常,统一在调用方用 or 合并 fallback,
    避免单点 LLM 故障让 4 个 mode 全都 500。
    """
    try:
        raw = llm.chat(SYSTEM_PROMPT, prompt, json_mode=True, temperature=0.3)
        data = json.loads(raw)
    except Exception:
        # 自由模式重试一次(部分中转不支持 response_format=json_object)
        try:
            raw = llm.chat(SYSTEM_PROMPT, prompt, json_mode=False, temperature=0.3)
            m = re.search(r"\{[\s\S]*\}", raw or "")
            data = json.loads(m.group(0)) if m else {}
        except Exception as exc:
            logger.warning("compare narration LLM 失败,降级 fallback: %s", exc)
            return {}
    if not isinstance(data, dict):
        return {}
    # 兼容 LLM 给出的不同 key 名,统一归一
    return {
        "summary":       str(data.get("summary", "")).strip(),
        "key_findings":  list(data.get("key_findings") or data.get("highlights") or []),
        "risks":         list(data.get("risks") or data.get("problems") or []),
        "suggestions":   list(data.get("suggestions") or data.get("建议") or []),
    }


def _fallback_narration_by_mode(
    mode: str,
    diff_details: Any,
    indicators: dict[str, Any],
    hint: str = "",
) -> dict[str, Any]:
    """LLM 不可用时,按 mode 模板生成事实陈述(纯 Pandas 计算,无编造)。

    关键:只引用 diff_details / indicators 里**已经计算出来的**数字,
    绝不二次推算。所有 mode 都保证输出 4 键(summary/key_findings/risks/suggestions),
    前端 el-collapse 直接渲染。
    """
    summary = ""
    key_findings: list[str] = []
    risks: list[str] = []
    suggestions: list[str] = ["如需更深入解读,请检查 OPENAI_API_KEY 是否可用。"]
    if hint:
        suggestions.append(hint)

    if mode == "business" and isinstance(diff_details, list):
        valid = [r for r in diff_details if isinstance(r.get("diff"), (int, float))]
        if not valid:
            summary = "无可对齐的分组,未能计算差值。"
        else:
            total_l = sum((r.get("left_sum") or 0) for r in valid)
            total_r = sum((r.get("right_sum") or 0) for r in valid)
            total_diff = round(total_r - total_l, 2)
            total_growth = round((total_diff / total_l * 100), 2) if total_l else None
            summary = (
                f"按 {indicators.get('group_col', '分组')} 聚合 {indicators.get('metric_col', '指标')},"
                f"共 {len(valid)} 个分组;左合计 {round(total_l, 2)}、右合计 {round(total_r, 2)},"
                f"差值 {total_diff}"
                + (f",变动 {total_growth}%。" if total_growth is not None else "。")
            )
            # 涨幅/跌幅 Top 3
            up = sorted(
                (r for r in valid if (r.get("growth_pct") or 0) > 0),
                key=lambda r: r["growth_pct"],
                reverse=True,
            )[:3]
            down = sorted(
                (r for r in valid if (r.get("growth_pct") or 0) < 0),
                key=lambda r: r["growth_pct"],
            )[:3]
            for r in up:
                key_findings.append(
                    f"{r['group']} 上升 {round(r['growth_pct'], 2)}% ({r['left_sum']} → {r['right_sum']})"
                )
            for r in down:
                key_findings.append(
                    f"{r['group']} 下降 {abs(round(r['growth_pct'], 2))}% ({r['left_sum']} → {r['right_sum']})"
                )
            if not up and not down:
                key_findings.append("各分组无明显涨跌。")
            risks.append("如存在异常波动,建议核对原始数据源。")

            # 【topn 并入 business 新增】TopN 截断时,补充「名次变动」类关键发现。
            # 措辞上强调「名次变化可能只是相对位移」,避免读者把它当成经营下滑。
            #
            # 注意:这里必须遍历 diff_details 全量行,不能只遍历 valid ——
            # valid 是「两侧都有值、diff 可计算」的行,而"新进入 / 跌出 TopN"
            # 的分组恰恰是**单边存在**的行(其 diff 为 None),会被 valid 过滤掉。
            top_n = indicators.get("top_n")
            if top_n:
                all_rows = diff_details if isinstance(diff_details, list) else []
                # 单边存在的分组:只有一侧有名次,另一侧为 0
                dropped = [
                    r for r in all_rows
                    if r.get("left_rank") and not r.get("right_rank")
                ]
                entered = [
                    r for r in all_rows
                    if r.get("right_rank") and not r.get("left_rank")
                ]
                # 两侧都有名次但名次不同 → 相对位移
                shifted = [
                    r for r in all_rows
                    if r.get("left_rank") and r.get("right_rank")
                    and r["left_rank"] != r["right_rank"]
                ]
                if dropped:
                    key_findings.append(
                        f"跌出 Top{int(top_n)} 的分组:"
                        + "、".join(r["group"] for r in dropped[:3])
                    )
                if entered:
                    key_findings.append(
                        f"新进入 Top{int(top_n)} 的分组:"
                        + "、".join(r["group"] for r in entered[:3])
                    )
                if shifted and not entered and not dropped:
                    key_findings.append(
                        f"有 {len(shifted)} 个分组的名次发生变化"
                        "(相对位移,不代表其指标值涨跌)"
                    )
                if entered or dropped or shifted:
                    suggestions.append(
                        "名次变化可能只是相对位移(有更强的新业务挤入),"
                        "判断经营增减请以「差值 / 变化幅度」为准。"
                    )
    elif mode == "row_diff" and isinstance(diff_details, dict):
        a = len(diff_details.get("added") or [])
        rm = len(diff_details.get("removed") or [])
        m_total = int(diff_details.get("modified_total") or 0)
        summary = (
            f"行级差异:新增 {a} 条,删除 {rm} 条,修改 {m_total} 条。"
        )
        if a == 0 and rm == 0 and m_total == 0:
            summary = "两份数据完全一致,无新增/删除/修改。"
        # 抽前 3 条 modified 提关键发现
        for item in (diff_details.get("modified") or [])[:3]:
            changes = item.get("field_changes") or []
            if not changes:
                continue
            fc = changes[0]
            key_findings.append(
                f"主键 {item.get('key')} 的「{fc.get('field')}」"
                f"由 {fc.get('before')} 改为 {fc.get('after')}"
            )
        if m_total > 3:
            risks.append(f"另有 {m_total - 3} 条修改未在摘要中展开,请展开折叠面板查看。")

    return {
        "summary": summary or "已生成对比结果(详细数字见下方表格)。",
        "key_findings": key_findings,
        "risks": risks,
        "suggestions": suggestions,
    }


def _build_narration_prompt(
    mode: str,
    indicators: dict[str, Any],
    diff_details: Any,
) -> str:
    """按 mode 拼装 LLM prompt:角色 = 数据分析师;数字全部显式给出;输出 4 段 JSON。"""
    # 用 ensure_ascii=False 防止中文被转义成 \uXXXX,影响 LLM 阅读
    diff_json = json.dumps(diff_details, ensure_ascii=False, default=str)[:6000]
    ind_json = json.dumps(indicators, ensure_ascii=False, default=str)[:1000]

    role_block = (
        "你是一名资深数据分析师,基于下面给出的「左右两份数据集的对比结果」"
        "用中文写一份简洁的报告,所有数字必须直接引用结果 JSON,**不要编造任何数字**。\n"
    )
    common_tail = (
        "输出必须是合法 JSON,字段:\n"
        "  - summary:      一句话总览(40~80 字)\n"
        "  - key_findings: 关键发现,数组,3~5 条;每条一句话指出具体数字与变化\n"
        "  - risks:        风险提示,数组;无则空数组 []\n"
        "  - suggestions:  建议,数组;2~3 条可操作建议\n\n"
        f"indicators: {ind_json}\n"
        f"diff_details: {diff_json}\n"
    )

    if mode == "business":
        # 【topn 并入 business 新增】TopN 截断时追加一段口径约束:
        # 明确告诉 LLM「名次变化 ≠ 经营涨跌」,避免它把相对位移写成下滑。
        topn_hint = ""
        if indicators.get("top_n"):
            topn_hint = (
                f"\n【重要口径】本次只截取了指标值最高的 Top{int(indicators.get('top_n'))} 个分组,"
                "结果中带有 left_rank / right_rank 名次字段。"
                "**名次变化只是相对位移**(有更强的新业务挤入会让老业务名次下降),"
                "不代表该分组自身经营变差。"
                "判断增减请以「差值 diff / 变化幅度 growth_pct」为准,"
                "名次只作为辅助参考提及。\n"
            )
        return (
            role_block
            + "本次对比类型:**业务对比(按分组聚合指标)**。\n"
            + "请从「变化方向 / 关键分组 / 建议」三方面点评;突出涨幅最大、跌幅最大、贡献最大的分组。\n"
            + topn_hint
            + common_tail
        )
    if mode == "row_diff":
        return (
            role_block
            + "本次对比类型:**行级变更追踪**(以指定主键匹配左右两表)。\n"
            + "请从「新增 / 删除 / 修改」三方面点评,给出 1~2 个具体修改示例(主键 + 字段 + before/after)。\n"
            + common_tail
        )
    return role_block + common_tail


def _build_document_compare_prompt(
    left_name: str,
    right_name: str,
    left_type: str,
    right_type: str,
    left_text: str,
    right_text: str,
) -> str:
    """为跨模态文档对比构造 prompt。

    设计要点:
        - 显式告诉 LLM 左右两份材料的文件类型,提示它理解"数据形态差异"
          (比如左边是 Excel 表格,右边是 Word 报告,数据呈现方式天然不同)
        - 文本可能很长,这里再做一次截断保险(单份不超过 8000 字),
          防止 prompt 超过模型上下文
        - 输出 schema 与 Pandas 对比保持一致(4 键 JSON),
          前端可以复用同一套渲染逻辑
    """
    type_cn = {"structured": "结构化数据(表格)", "document": "文档", "image": "图片"}
    left_desc = type_cn.get(left_type, left_type)
    right_desc = type_cn.get(right_type, right_type)

    # 单份文本截断到 8000 字,双份共 16000 字,给 system+输出留空间
    left_snippet = (left_text or "").strip()[:8000] or "(未抽取到文本内容)"
    right_snippet = (right_text or "").strip()[:8000] or "(未抽取到文本内容)"

    return (
        "你是一名资深数据分析师,现在要对比**两份形态可能不同的材料**,"
        "用中文写一份客观、基于原文的对比报告。\n\n"
        f"【左方】{left_name}(类型:{left_desc})\n"
        f"【右方】{right_name}(类型:{right_desc})\n\n"
        "对比要求:\n"
        "1. **忠于原文**:所有数字、事实必须直接引用下面给出的文本,**严禁编造**;\n"
        "2. **形态感知**:如果左右类型不同(比如一份是数字表格、一份是文字报告),\n"
        "   请指出这种形态差异本身带来的对比局限,并聚焦在可以语义对齐的维度;\n"
        "3. **业务视角**:优先对比业务事实(目标/结果/趋势/重点),而不是文体差异;\n"
        "4. **标注不确定**:如果某份材料信息不足以支撑某个判断,明确说\"无法判断\",不要硬编。\n\n"
        "输出必须是合法 JSON,字段:\n"
        "  - summary:      一句话总览两份材料的核心差异(60~120 字)\n"
        "  - key_findings: 关键发现,数组,3~5 条;每条一句话指出具体差异点\n"
        "  - risks:        数据/口径/完整性风险,数组;无则空数组 []\n"
        "  - suggestions:  建议,数组,2~3 条;比如\"建议补充 XX 数据\"\"建议统一 XX 口径\"\n\n"
        f"【左方原文】\n{left_snippet}\n\n"
        f"【右方原文】\n{right_snippet}\n"
    )


def _compare_documents(
    left_name: str,
    right_name: str,
    left_type: str,
    right_type: str,
    left_text: str,
    right_text: str,
) -> dict[str, Any]:
    """跨模态文档对比:把两份非结构化(或混合)材料的文本交给 LLM 做语义对比。

    触发场景:
        compare_datasets 检测到任一边 file_type 不是 structured 时调用。

    返回:
        与 Pandas 对比相同的结构,但 indicators 只包含文本统计信息,
        diff_details 是 None(没有 Pandas 算子可跑)。
    """
    # 文本统计指标(不算业务指标,只反映材料规模)
    indicators = {
        "left_type": left_type,
        "right_type": right_type,
        "left_chars": len(left_text or ""),
        "right_chars": len(right_text or ""),
        "compare_mode": "document_compare",
    }

    # 有一边文本为空 → LLM 也编不出内容,直接走 fallback
    if not (left_text or "").strip() or not (right_text or "").strip():
        empty_side = "左方" if not (left_text or "").strip() else "右方"
        narration = {
            "summary": f"{empty_side}材料未能抽取出有效文本,无法进行对比。",
            "key_findings": [
                f"{empty_side}文本抽取结果为空,可能是扫描版 PDF 未识别或图片无内容。",
                "建议检查文件本身是否包含可识别的文字或数据。",
            ],
            "risks": ["文本抽取失败会导致对比结论严重偏离,请勿基于本次结果做决策。"],
            "suggestions": [
                "确认上传的文件不是纯空白或加密文档。",
                "如是扫描版 PDF,确认视觉模型配置(OPENAI_VISION_MODEL)可用。",
            ],
        }
    else:
        prompt = _build_document_compare_prompt(
            left_name, right_name, left_type, right_type, left_text, right_text,
        )
        try:
            narration = llm.chat_json(
                "你是一名资深数据分析师,擅长跨格式材料对比。输出严格 JSON。",
                prompt,
            )
        except Exception as exc:
            logger.warning("[compare] 文档对比 LLM 调用失败: %s", exc)
            narration = {}

        # 失败兜底:至少给出基本事实陈述
        if not narration or not narration.get("summary"):
            narration = {
                "summary": (
                    f"已读取左方({left_name},{indicators['left_chars']} 字)与"
                    f"右方({right_name},{indicators['right_chars']} 字)的文本内容,"
                    "但 AI 对比服务暂不可用,请人工阅读两份原文进行比对。"
                ),
                "key_findings": [
                    f"左方材料:{left_name},共 {indicators['left_chars']} 字。",
                    f"右方材料:{right_name},共 {indicators['right_chars']} 字。",
                    "AI 解读暂不可用,建议查看两份材料的预览文本。",
                ],
                "risks": ["AI 服务异常,本次对比未生成语义分析。"],
                "suggestions": [
                    "检查 OPENAI_API_KEY 是否可用。",
                    "稍后重试,或先下载两份原文人工比对。",
                ],
            }

    return {
        "indicators":   indicators,
        "comment":      narration.get("summary", ""),
        "narration":    narration,
        "diff_mode":    "document_compare",
        "diff_details": None,
        # 非结构化对比没有 sheet 概念,但保持返回结构一致,前端不会报错
        "left_sheets":       [],
        "right_sheets":      [],
        "left_sheet_count":  0,
        "right_sheet_count": 0,
    }


# ==================== 合同对比(垂直场景) ====================
# 背景:通用 document_compare 是「跨模态泛对比」,对合同这种有结构化条款
# (金额/期限/违约金/管辖地等)的文档不够用 —— 业务方真正需要的是
# 「条款级差异表 + 利弊 + 谈判建议」。
# 这里的 _compare_contracts 在通用对比基础上做了 3 个增强:
#   1. 提示 LLM 按条款维度去拆解(标的/金额/期限/付款/违约/争议解决…)
#   2. 要求输出结构化 clause_diffs 数组(供前端表格渲染)
#   3. 强制 pros_cons + negotiation_tips 字段,贴近真实业务诉求

_CONTRACT_CLAUSE_FIELDS = [
    {"key": "scope",        "label": "合同标的/标的物",   "hint": "服务内容、标的物、数量"},
    {"key": "amount",       "label": "金额/价款",         "hint": "总价、单价、币种、含税与否"},
    {"key": "term",         "label": "合同期限",         "hint": "起止日期、有效期、自动续约条款"},
    {"key": "payment",      "label": "付款方式",         "hint": "付款节点、账期、发票要求"},
    {"key": "delivery",     "label": "交付/验收",         "hint": "交付方式、验收标准、质保期"},
    {"key": "penalty",      "label": "违约责任",         "hint": "违约金、赔偿上限、不可抗力"},
    {"key": "confidentiality","label":"保密条款",         "hint": "保密范围、期限、违约金"},
    {"key": "ip",           "label": "知识产权",         "hint": "成果归属、许可范围"},
    {"key": "liability",    "label": "责任限制",         "hint": "责任上限、免责情形"},
    {"key": "jurisdiction", "label": "争议解决/管辖",    "hint": "诉讼管辖、仲裁机构、适用法律"},
    {"key": "termination",  "label": "解除/终止",         "hint": "单方解除条件、通知期、后果"},
    {"key": "other",        "label": "其他重要条款",      "hint": "通知/转让/反贿赂/数据保护等"},
]


def _build_contract_compare_prompt(
    left_name: str,
    right_name: str,
    left_type: str,
    right_type: str,
    left_text: str,
    right_text: str,
    focus: str = "",
) -> str:
    """为合同对比构造 prompt:引导 LLM 按条款维度拆解并输出结构化 JSON。

    设计要点:
        - 显式列出 12 类条款,确保 LLM 不漏关键维度
        - 输出 schema 包含 clause_diffs(结构化差异表) + pros_cons(利弊) + negotiation_tips(谈判建议)
        - 单份文本 12000 字(比通用对比多 4000),合同往往更长
        - focus 参数:前端允许用户指定"重点关注 XX",作为软提示注入
    """
    type_cn = {"structured": "结构化", "document": "文档", "image": "图片"}
    left_desc = type_cn.get(left_type, left_type)
    right_desc = type_cn.get(right_type, right_type)

    left_snippet = (left_text or "").strip()[:12000] or "(未抽取到文本)"
    right_snippet = (right_text or "").strip()[:12000] or "(未抽取到文本)"

    # 把 12 类条款渲染成清单,让 LLM 看到「该看哪些维度」
    clause_lines = "\n".join(
        f"  - {c['key']}: {c['label']}({c['hint']})"
        for c in _CONTRACT_CLAUSE_FIELDS
    )

    focus_block = ""
    if focus and focus.strip():
        focus_block = f"\n【重点关注】\n本次对比用户特别关注以下方向,请重点分析:\n{focus.strip()}\n"

    return (
        "你是一名资深商务/法务顾问,现在要对比两份**合同/协议**文档,"
        "可能是同一份合同的相邻年度版本(如今年 vs 去年),也可能是两份独立的合同。\n\n"
        f"【左方】{left_name}(类型:{left_desc})\n"
        f"【右方】{right_name}(类型:{right_desc})\n"
        f"{focus_block}\n"
        "请按以下 12 类条款维度逐项对比,不要遗漏关键维度:\n"
        f"{clause_lines}\n\n"
        "【对比要求】\n"
        "1. **忠于原文**:每条差异必须基于原文事实,**禁止编造**数字/期限/条款;\n"
        "2. **标注不确定**:若某条款在某份合同中未明确,标注\"未约定\"而非推测;\n"
        "3. **区分三类差异**:\n"
        "   - 实质变更(影响权利义务,如金额/期限/违约金变化)\n"
        "   - 表述优化(只是措辞调整,无实质影响)\n"
        "   - 新增/删除(某份独有的条款)\n"
        "4. **商务视角**:除了描述差异,要说清「对客户/对我方的利弊」。\n\n"
        "【输出 schema】(必须是合法 JSON,不要任何额外说明):\n"
        "{\n"
        '  "summary": "一句话总览两份合同的核心差异(60~120字)",\n'
        '  "clause_diffs": [\n'
        '    {\n'
        '      "clause_key": "amount",                  // 对应上面 12 类的 key\n'
        '      "clause_label": "金额/价款",                // 中文标签\n'
        '      "importance": "high",                  // high/medium/low\n'
        '      "change_type": "modified",              // modified/added/removed/unchanged/unclear\n'
        '      "left_value": "总价 100 万",              // 左方该条款原文要点,可空\n'
        '      "right_value": "总价 120 万",             // 右方该条款原文要点,可空\n'
        '      "diff_note": "金额从 100 万上调至 120 万,涨幅 20%"  // 差异描述,1~2 句\n'
        '    }\n'
        '  ],\n'
        '  "key_findings": ["关键发现 1", "关键发现 2"],\n'
        '  "risks": ["风险点 1", "风险点 2"],\n'
        '  "pros_cons": [\n'
        '    {\n'
        '      "side": "right",                       // left=左方更优, right=右方更优, both=各有优劣\n'
        '      "title": "右方付款账期更短",              // 一句话标题\n'
        '      "detail": "从 60 天缩短至 30 天,我方资金周转压力上升"  // 详细说明\n'
        '    }\n'
        '  ],\n'
        '  "negotiation_tips": ["谈判建议 1", "谈判建议 2"]\n'
        "}\n\n"
        f"【左方原文】\n{left_snippet}\n\n"
        f"【右方原文】\n{right_snippet}\n"
    )


def _compare_contracts(
    left_name: str,
    right_name: str,
    left_type: str,
    right_type: str,
    left_text: str,
    right_text: str,
    focus: str = "",
) -> dict[str, Any]:
    """合同对比(垂直场景):比通用 document_compare 多了条款级结构化输出。

    返回结构与 compare_datasets 同形,前端可复用同一套渲染管线。
    关键差异:
        - diff_mode = "contract_compare"
        - diff_details = clause_diffs 数组(供表格渲染)
        - narration 里多了 pros_cons / negotiation_tips
    """
    indicators = {
        "left_type": left_type,
        "right_type": right_type,
        "left_chars": len(left_text or ""),
        "right_chars": len(right_text or ""),
        "compare_mode": "contract_compare",
    }

    if not (left_text or "").strip() or not (right_text or "").strip():
        empty_side = "左方" if not (left_text or "").strip() else "右方"
        narration = {
            "summary": f"{empty_side}合同材料未能抽取出有效文本,无法对比条款。",
            "key_findings": [
                "请确认文件不是扫描件、加密文档或纯图片。",
                "PDF 扫描件需要视觉模型支持(检查 OPENAI_VISION_MODEL 配置)。",
            ],
            "risks": ["文本缺失将无法做条款级差异分析。"],
            "suggestions": ["重新上传文件,或检查后端日志的抽取错误。"],
            "pros_cons": [],
            "negotiation_tips": [],
        }
        clause_diffs: list[dict[str, Any]] = []
    else:
        prompt = _build_contract_compare_prompt(
            left_name, right_name, left_type, right_type, left_text, right_text, focus,
        )
        try:
            result = llm.chat_json(
                "你是一名资深商务/法务顾问,擅长合同条款级对比分析。输出严格 JSON。",
                prompt,
            )
        except Exception as exc:
            logger.warning("[compare] 合同对比 LLM 调用失败: %s", exc)
            result = {}

        # LLM 偶尔会把字段拍平,这里只取已知字段
        clause_diffs = result.get("clause_diffs") or []
        if not isinstance(clause_diffs, list):
            clause_diffs = []
        # 二次清洗 + 字段补全,避免前端渲染时 NPE
        normalized: list[dict[str, Any]] = []
        for c in clause_diffs:
            if not isinstance(c, dict):
                continue
            normalized.append({
                "clause_key":   c.get("clause_key") or "other",
                "clause_label": c.get("clause_label") or c.get("clause_key") or "其他",
                "importance":   c.get("importance") or "medium",
                "change_type":  c.get("change_type") or "unclear",
                "left_value":   c.get("left_value") or "",
                "right_value":  c.get("right_value") or "",
                "diff_note":    c.get("diff_note") or "",
            })
        clause_diffs = normalized

        # 关键差异:若有 high importance 的 clause 变更,放在 key_findings 最前面
        high_changes = [c for c in clause_diffs if c.get("importance") == "high" and c.get("change_type") in ("modified", "added", "removed")]
        if high_changes:
            lead = f"识别到 {len(high_changes)} 项高优先级条款变更。"
        else:
            lead = "本次合同差异整体较为温和。"
        raw_summary = (result.get("summary") or "").strip()
        summary_text = (lead + " " + raw_summary).strip() if raw_summary else lead

        pros_cons = result.get("pros_cons") or []
        if not isinstance(pros_cons, list):
            pros_cons = []
        negotiation_tips = result.get("negotiation_tips") or []
        if not isinstance(negotiation_tips, list):
            negotiation_tips = []

        narration = {
            "summary":           summary_text,
            "key_findings":      result.get("key_findings") or [],
            "risks":             result.get("risks") or [],
            "suggestions":       result.get("suggestions") or [],
            "pros_cons":         pros_cons,
            "negotiation_tips":  negotiation_tips,
        }
        # 兜底:LLM 返回极简时给最小可用叙述
        if not summary_text and not narration["key_findings"]:
            narration = {
                "summary": (
                    f"已读取左方({left_name},{indicators['left_chars']} 字)与"
                    f"右方({right_name},{indicators['right_chars']} 字),"
                    "但 AI 条款级对比暂不可用,建议人工比对。"
                ),
                "key_findings": [
                    f"左方:共 {indicators['left_chars']} 字。",
                    f"右方:共 {indicators['right_chars']} 字。",
                ],
                "risks": ["AI 服务异常,本次未生成条款级分析。"],
                "suggestions": ["检查 OPENAI_API_KEY 后重试。"],
                "pros_cons": [],
                "negotiation_tips": [],
            }
            clause_diffs = []

    return {
        "indicators":   indicators,
        "comment":      narration.get("summary", ""),
        "narration":    narration,
        "diff_mode":    "contract_compare",
        # diff_details 直接是 clause_diffs,前端表格组件可直接消费
        "diff_details": clause_diffs,
        "left_sheets":       [],
        "right_sheets":      [],
        "left_sheet_count":  0,
        "right_sheet_count": 0,
    }




def compare_datasets(
    left_file: str,
    right_file: str,
    mode: str = "business",
    group_col: str | None = None,
    metric_col: str | None = None,
    key_cols: list[str] | None = None,
    top_n: int | None = None,
    persist: bool = False,
    # —— 多格式文件支持(2026-10 新增) ——
    # 调用方(tools.py)从 Dataset 记录里读出这 4 个字段透传过来,
    # 本函数据此决定走 Pandas 链路还是 LLM 文档对比链路。
    # 默认值保持向后兼容(旧调用方不传时,行为与旧版一致)。
    left_file_type: str = "structured",
    right_file_type: str = "structured",
    left_extracted_text: str = "",
    right_extracted_text: str = "",
    left_name: str = "左数据集",
    right_name: str = "右数据集",
    # 【dataset-contract-compare 新增】合同对比场景的用户关注点
    contract_focus: str = "",
) -> dict[str, Any]:
    """对比两份数据集:按 mode 路由算子,产出 diff_details + 中文叙述。

    入参:
        left_file / right_file: 数据集文件路径
        mode: business | row_diff
        group_col / metric_col: business 用
        key_cols:                 row_diff 用
        top_n:                    business 用(可选)。指定正整数时,
                                  只保留指标值最高的 N 个分组,并附加
                                  left_rank / right_rank 名次作为辅助参考列;
                                  为 None 时返回全部分组(旧行为)。
        persist:                  【tools-results-history-persistence 新增】
                                  仅用于审计日志;实际落库由 tools.compare 完成。
                                  默认 False,旧调用方不受影响。

    异常:
        ValueError - 列不存在 / 主键缺失 / topn 非法 → 路由层转换为 400
        任何算子内部异常都会被 try/except 包住,转成 fallback 叙述,不再 500

    返回:
        {
            "indicators":    dict,                 # 按 mode 不同的指标摘要
            "comment":       str,                  # v1 兼容字段,直接复用 summary
            "narration":     {summary, key_findings, risks, suggestions},
            "diff_mode":     "business" | ... | "document_compare",
            "diff_details":  list | dict | None,   # document_compare 时为 None
            "left_sheets" / "right_sheets" / "left_sheet_count" / "right_sheet_count"
        }
    """
    # —— 类型分流(2026-10 新增) ——
    # 1) mode=contract_compare → 合同垂直场景(用户显式选了)
    # 2) 任一边不是 structured → 走 LLM 跨模态文档对比,不再走 Pandas 算子
    # 3) 否则按 mode 路由到 Pandas 算子
    # 优先级:合同模式 > 通用文档对比 > 表格对比
    if mode == "contract_compare":
        logger.info(
            "[compare] 合同对比: left=%s right=%s, focus=%r",
            left_name, right_name, (contract_focus or "")[:50],
        )
        return _compare_contracts(
            left_name=left_name,
            right_name=right_name,
            left_type=left_file_type,
            right_type=right_file_type,
            left_text=left_extracted_text,
            right_text=right_extracted_text,
            focus=contract_focus,
        )
    if left_file_type != "structured" or right_file_type != "structured":
        logger.info(
            "[compare] 非结构化对比: left_type=%s right_type=%s, 走 document_compare",
            left_file_type, right_file_type,
        )
        return _compare_documents(
            left_name=left_name,
            right_name=right_name,
            left_type=left_file_type,
            right_type=right_file_type,
            left_text=left_extracted_text,
            right_text=right_extracted_text,
        )

    if mode not in _VALID_MODES:
        raise ValueError(f"mode 必须是 {_VALID_MODES} 之一,当前:{mode}")

    # 【tools-results-history-persistence 新增】审计日志:仅在 persist=True 时写一条 INFO,
    # 便于运维追溯"哪些对比结果被留痕"。不影响主流程。
    if persist:
        logger.info("[compare] persist=True, mode=%s", mode)

    # 1) 读两份数据(2026-10 多 sheet 分项:不再走列并集合并大表)
    #    按 sheet 独立读,选取**第一对 sheet 名称相同**(若左右 sheet 名能对齐)
    #    或左/右首张 sheet,作为本次对比的"工作面"。
    per_l = pe.read_dataframe_per_sheet(left_file)
    per_r = pe.read_dataframe_per_sheet(right_file)
    sheets_l = [{"name": s["name"], "rows": int(len(s["df"])), "cols": int(len(s["df"].columns))}
                for s in per_l]
    sheets_r = [{"name": s["name"], "rows": int(len(s["df"])), "cols": int(len(s["df"].columns))}
                for s in per_r]

    # 决定对比的工作面:左右 sheet 名按位置对齐,首张同名的优先;否则左/右都取第一张
    aligned_sheet_name: str | None = None
    if per_l and per_r:
        # 优先按名字相同的对齐
        left_names = {s["name"] for s in per_l}
        right_names = {s["name"] for s in per_r}
        common = left_names & right_names
        if common:
            # 取"双方都有的第一张 sheet",保持顺序
            for s in per_l:
                if s["name"] in common:
                    aligned_sheet_name = s["name"]
                    break
        # 否则按位置对齐(左第一 vs 右第一)
        if aligned_sheet_name is None:
            aligned_sheet_name = per_l[0]["name"]
    df_l = next((s["df"] for s in per_l if s["name"] == aligned_sheet_name), per_l[0]["df"]) if per_l else None
    df_r = next((s["df"] for s in per_r if s["name"] == aligned_sheet_name), per_r[0]["df"]) if per_r else None
    if df_l is None or df_r is None:
        raise ValueError("对比数据集无法读取(可能为空)")
    # 在 indicators 里附 sheet 对齐信息,前端能看到"本次对比按哪张 sheet 工作"
    sheet_alignment = {
        "aligned_sheet": aligned_sheet_name,
        "left_sheets": sheets_l,
        "right_sheets": sheets_r,
    }

    # 2) 按 mode 路由到 pandas_engine 对应算子
    diff_details: Any
    indicators: dict[str, Any]

    if mode == "business":
        # 【topn 并入 business 新增】top_n 为 None 时返回全部分组(旧行为);
        # 指定正整数时只保留指标值最高的 N 个分组,并附带名次辅助字段。
        effective_top_n = int(top_n) if top_n is not None and int(top_n) > 0 else None
        diff_details = pe.compare_by_group(
            df_l, df_r, group_col or "", metric_col or "", top_n=effective_top_n
        )
        valid = [r for r in diff_details if isinstance(r.get("diff"), (int, float))]
        total_l = sum((r.get("left_sum") or 0) for r in valid)
        total_r = sum((r.get("right_sum") or 0) for r in valid)
        indicators = {
            "group_col": group_col,
            "metric_col": metric_col,
            "group_count": len(valid),
            "left_total": round(total_l, 2),
            "right_total": round(total_r, 2),
            "diff": round(total_r - total_l, 2),
        }
        if effective_top_n:
            indicators["top_n"] = effective_top_n
            # 名次变动统计(仅在两侧都有名次的行上有意义)
            both = [
                r for r in diff_details
                if r.get("left_rank") and r.get("right_rank")
            ]
            indicators["rank_up"] = sum(
                1 for r in both if r["right_rank"] < r["left_rank"]
            )
            indicators["rank_down"] = sum(
                1 for r in both if r["right_rank"] > r["left_rank"]
            )
            indicators["rank_entered"] = sum(
                1 for r in diff_details if r.get("right_rank") and not r.get("left_rank")
            )
            indicators["rank_dropped"] = sum(
                1 for r in diff_details if r.get("left_rank") and not r.get("right_rank")
            )
    elif mode == "row_diff":
        keys = key_cols or []
        diff_details = pe.row_level_diff(df_l, df_r, keys)
        indicators = {
            "key_cols": keys,
            "added_count": len(diff_details.get("added") or []),
            "removed_count": len(diff_details.get("removed") or []),
            "modified_count": int(diff_details.get("modified_total") or 0),
        }
    else:
        # 已被 _VALID_MODES 拦截,这里只是给静态检查一个落点
        raise ValueError(f"未实现的 mode:{mode}")

    # 3) LLM 拿叙述(失败 → fallback,不再抛 500)
    prompt = _build_narration_prompt(mode, indicators, diff_details)
    narration = _call_narration_llm(mode, prompt)
    if not narration or not narration.get("summary"):
        narration = _fallback_narration_by_mode(mode, diff_details, indicators)
    elif not narration.get("suggestions"):
        # LLM 给的叙述没建议,补一条 fallback 兜底,避免前端 el-collapse 空白
        fb = _fallback_narration_by_mode(mode, diff_details, indicators)
        narration["suggestions"] = fb.get("suggestions") or []

    return {
        "indicators":   indicators,
        "comment":      narration.get("summary", ""),     # v1 字段,继续保留
        "narration":    narration,
        "diff_mode":    mode,
        "diff_details": diff_details,
        # sheet 明细:前端对比弹窗可显示「两份数据集各自的 sheet 分布」
        "left_sheets":       sheets_l,
        "right_sheets":      sheets_r,
        "left_sheet_count":  len(sheets_l),
        "right_sheet_count": len(sheets_r),
        # 【多 sheet 分项 2026-10】本次对比按哪张 sheet 工作(按同名优先 + 位置兜底)
        "aligned_sheet":     aligned_sheet_name,
    }


# ---------------------- 图表编辑（自然语言 → ECharts option） ----------------------

# ECharts 图表类型白名单,防止 LLM 编造不存在的类型
_CHART_TYPES = {"line", "bar", "pie", "scatter", "radar", "heatmap", "funnel"}


def chart_edit(
    instruction: str,
    current_option: dict,
    df_meta: dict,
) -> dict[str, Any]:
    """根据自然语言指令修改 ECharts option。

    入参:
        instruction: 用户指令(如 "换成折线图" "配色换成蓝色" "只显示 Top 5")
        current_option: 当前图表 option(可以为空,表示从零生成)
        df_meta: {
            "columns": [{"name", "type"}, ...],
            "sample_rows": [[..], [..], ...]    # 前 5 行样例,给 LLM 推断数据形状
            "title": "原报告标题"
        }

    返回:
        {
            "type": "line" | ...,
            "title": "图表标题",
            "option": { ...完整 ECharts option... },
            "explanation": "一句话说明本次改了什么",
            "fallback": bool,
            "error": str
        }
    """
    columns_desc = ", ".join(f"{c['name']}({c['type']})" for c in (df_meta.get("columns") or []))

    # 【加固】在拼 prompt 前,用 json_default 走一遍「序列化 → 反序列化」,
    # 把 current_option / sample_rows 里潜在的 pd.Timestamp / np.int64 / pd.NaT
    # 一次性"清洗"成纯 Python 原生类型,根除 "not JSON serializable" 的任何隐患;
    # 即便后续忘记给 json.dumps 加兜底,这里也兜住了。同时这是无副作用的,
    # 因为下面两次 json.dumps 都仍带着 default=json_default,真正安全的是双保险。
    def _pure(o):
        return json.loads(json.dumps(o, ensure_ascii=False, default=json_default))

    current_option_pure = _pure(current_option or {})
    sample_rows_pure = _pure(df_meta.get('sample_rows') or [])

    # 【修复】这俩序列化必须带 default=json_default:current_option 与 sample_rows
    # 可能源自前端原始 payload / Pandas DataFrame.values.tolist(),
    # 若残留 pd.Timestamp / np.int64 / pd.NaT 等就会抛 "not JSON serializable",
    # 进而被外层 except 当成 "chart_edit 失败" 回显。
    # 用 default=json_default 兜底,既不影响 LLM 看到的文本,又永远不再抛序列化错。
    user = (
        f"原图表配置:{json.dumps(current_option_pure, ensure_ascii=False, default=json_default)[:3000]}\n"
        f"数据集列:{columns_desc}\n"
        f"前 5 行样例:{json.dumps(sample_rows_pure, ensure_ascii=False, default=json_default)[:1500]}\n"
        f"原标题:{df_meta.get('title', '')}\n\n"
        f"用户指令:{instruction}\n\n"
        "请生成修改后的 ECharts 配置,输出 JSON:\n"
        "{\n"
        '  "type": "line | bar | pie | scatter | radar | heatmap | funnel",\n'
        '  "title": "图表标题",\n'
        '  "option": { ... 完整 ECharts option,必须含 title/tooltip/series ... },\n'
        '  "explanation": "一句话说明这次改了什么(20字内)"\n'
        "}\n"
        "约束:\n"
        "  - type 必须从白名单选;option 必须是合法 ECharts JSON\n"
        "  - 数据从原图表配置或 sample_rows 推断,不要编造数字\n"
        "  - 配色使用以下安全色板之一:['#22d3ee','#5eead4','#a78bfa','#fbbf24','#f87171','#4ade80']"
    )

    try:
        raw = ""
        parse_ok = False
        data = {}
        try:
            raw = llm.chat(SYSTEM_PROMPT, user, json_mode=True)
            # 【修复】LLM 经常返回"接近 JSON 但又不规范"的输出(单引号、未加引号裸 key、
            # markdown 包裹、thinking 截断尾巴),标准 json.loads 会挂;
            # 这里改用 _safe_json_loads 宽容解析,提升成功率。
            data = _safe_json_loads(raw or "")
            # 【修复】_safe_json_loads 可能解析出一个 list / str / null 而不是 dict,
            # 这时不能用,需要当作"未拿到" → 走第二次自由模式。
            if isinstance(data, dict):
                parse_ok = True
            else:
                logger.info("chart_edit: 解析成功但不是 dict (type=%s, raw[:120]=%r)",
                            type(data).__name__, (raw or "")[:120])
        except (json.JSONDecodeError, ValueError) as first_exc:
            logger.info("chart_edit: JSON 模式失败(%s),尝试自由模式", first_exc)
            try:
                raw = llm.chat(SYSTEM_PROMPT, user, json_mode=False)
                data = _safe_json_loads(raw or "")
                if isinstance(data, dict):
                    parse_ok = True
                else:
                    logger.info("chart_edit: 自由模式解析成功但不是 dict (type=%s, raw[:120]=%r)",
                                type(data).__name__, (raw or "")[:120])
            except Exception as second_exc:
                logger.warning("chart_edit 重试仍失败: %s / %s", first_exc, second_exc)
                raise

        if not parse_ok or not isinstance(data, dict):
            raise ValueError("两次调用都未拿到 JSON")

        # 校验 type,白名单外则降级
        ctype = data.get("type", "line")
        if ctype not in _CHART_TYPES:
            ctype = "line"

        return {
            "type": ctype,
            "title": data.get("title", "") or df_meta.get("title", ""),
            "option": data.get("option") or {},
            "explanation": data.get("explanation", "已更新图表"),
            "fallback": False,
            "error": "",
        }
    except Exception as exc:
        logger.warning("chart_edit 失败,降级为原图表: %s", exc)
        # 失败时返回原图,不丢功能
        return {
            "type": (current_option or {}).get("type", "line") if isinstance(current_option, dict) else "line",
            "title": df_meta.get("title", ""),
            "option": current_option or {},
            "explanation": "本次编辑失败,已保留原图",
            "fallback": True,
            "error": str(exc),
        }


# ---------------------- 智能派生列推荐 ----------------------
# 【clean-ai-derived-columns-pivot 新增】
# 需求:根据现有列画像,给「数据清洗」弹窗推荐可直接计算的派生列
# (如 销售额 = 销量 × 单价)。LLM 只是建议者,所有建议最终都会走
# pandas_engine.validate_derived_specs 的同一套安全校验。

_DERIVED_SUGGEST_MAX = 6      # LLM 派生列建议最多保留 6 条
_DERIVED_FALLBACK_MAX = 3     # 规则兜底派生列最多 3 条
_PIVOT_SUGGEST_MAX = 8        # LLM 透视建议最多保留 8 条(支持多视角并行)
_PIVOT_FALLBACK_MAX = 3       # 规则兜底透视每张 sheet 最多 3 条

# 兜底规则用的列名关键词(按优先级排列)
_DERIVED_QUANTITY_WORDS = ("销量", "销售数量", "数量", "件数")
_DERIVED_PRICE_WORDS = ("单价", "价格", "售价", "均价")
_DERIVED_REVENUE_WORDS = ("收入", "营收", "销售额", "总收入")
_DERIVED_COST_WORDS = ("成本", "成本额", "总成本")

# 透视兜底挑选"指标列"时优先的关键词(命中其一即视为高优先级指标列)
_PIVOT_VALUE_WORDS = (
    "金额", "销售额", "利润", "收入", "营收", "成本",
    "销量", "数量", "件数", "占比", "率", "价", "额",
)

# 透视兜底挑选"维度列"时优先的关键词(命中其一即视为高优先级维度列)
_PIVOT_DIM_WORDS = (
    "渠道", "地区", "区域", "省市", "客户", "产品", "商品",
    "类别", "类型", "分类", "部门", "行业", "月份", "月", "季度",
    "周", "年", "日期", "日期时间", "品牌", "门店", "销售",
)

_PIVOT_AGG_WHITELIST = {"sum", "mean", "count", "max", "min"}

_DERIVED_SUGGEST_SYSTEM = (
    "你是资深数据分析师,负责为数据表推荐可直接计算的派生列与透视汇总。"
    "必须遵守以下规则:\n"
    "【派生列】\n"
    "1. 只允许逐行四则运算(+ - * /)与括号,禁止函数、聚合、条件判断、跨行计算;\n"
    "2. 公式中的列名必须用反引号包裹,例如 `销量` * `单价`;\n"
    "3. 新列名不得与已有列同名,长度不超过 30 个字符;\n"
    "4. 最多推荐 6 条派生列,宁缺毋滥。\n"
    "【透视汇总(重要)】\n"
    "5. 透视必须选同一工作表内真实存在的列;\n"
    "6. 行维度(row_dim)选非数值列或基数较低(<=50)的数值列;"
    "value 必须是数值列或你自己推荐的派生列(派生列默认也是数值型);"
    "agg 只允许 sum/mean/count/max/min;\n"
    "7. 列维度(col_dim)可选,填 null 表示只做单维度透视;\n"
    "8. description 用中文一句话说明这张表回答了什么问题"
    "(例如「按渠道看销售额合计」)。\n"
    "9. 最多推荐 8 条透视,鼓励多角度交叉(不同 row_dim × 不同 value 组合)"
    "例如:区域×年度合计、产品×年度合计、区域×产品×年度合计、渠道×月均销售额 等;\n"
    "10. value 既可以用表里原始的数值列,也可以用你上面推荐的派生列"
    "(例如你推荐了「第一季度销售额」,就可以把它作为某张透视的 value);\n"
    "11. 只输出 JSON,不要任何解释文字。"
)


def _build_derived_suggest_prompt(profiles: list[dict]) -> str:
    """把各 sheet 的字段画像拼成中文 Prompt(同时产出派生列与透视建议)。"""
    payload = json.dumps(profiles, ensure_ascii=False, default=str)
    return (
        "下面是各工作表的结构画像(JSON,type 为 number/text/date,"
        "unique_count 是该列非空去重后的数量):\n"
        f"{payload}\n\n"
        "请据此推荐派生列与透视汇总,只输出如下 JSON:\n"
        "{\n"
        '  "suggestions":[\n'
        '    {"sheet_name":"工作表名","name":"新列名",'
        '     "expression":"`列A` * `列B`","description":"中文说明"}\n'
        "  ],\n"
        '  "pivot_suggestions":[\n'
        '    {"sheet_name":"工作表名","row_dim":"渠道","col_dim":null,'
        '     "value":"销售额","agg":"sum","include_margin":true,'
        '     "description":"按渠道看销售额合计"}\n'
        "  ]\n"
        "}\n"
        "注意:\n"
        "1. sheet_name 必须是上面真实存在的工作表名;\n"
        "2. expression 只能引用同一张表里的数值列,列名必须用反引号包裹;\n"
        "3. 不要编造不存在的列,不要推荐与现有列同名的派生列;\n"
        "4. 透视 row_dim 必须存在于该 sheet;"
        "value 可以是该 sheet 的原始数值列,也可以是你自己上面推荐的派生列名;"
        "agg 只能是 sum/mean/count/max/min;\n"
        "5. 鼓励多角度透视:"
        "渠道/地区 × 金额(sum)、产品 × 数量(sum)、地区×产品 × 金额、"
        "派生列(如季度销售额) × sum 等,最多 8 条。"
    )


def _pick_profile_column(
    columns: list[dict], keywords: tuple[str, ...], used: set[str]
) -> dict | None:
    """在字段画像中按关键词挑一个数值列。

    匹配优先级:精确同名 > 关键词包含。这样「销量」不会被「销售量」抢先匹配,
    同 sheet 内已选过的列通过 used 排除,避免 数量 × 数量 这种自乘。
    """
    candidates: list[tuple[int, int, dict]] = []
    for col in columns:
        if not isinstance(col, dict):
            continue
        name = str(col.get("name") or "")
        if not name or name in used or col.get("type") != "number":
            continue
        for rank, keyword in enumerate(keywords):
            if name == keyword:
                candidates.append((0, rank, col))
                break
            if keyword in name:
                candidates.append((1, rank, col))
                break
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _fallback_derived_suggestions(profiles: list[dict]) -> list[dict]:
    """LLM 不可用时的规则兜底(纯 Pandas 侧的结构判断,不调用任何模型)。

    覆盖两类最常见业务公式:
        - 数量类 × 价格类 → 销售额
        - 收入类 − 成本类 → 利润
    约束:同一 sheet 内两列都必须是 number;新列名不得与现有列或已推荐项重复;
    每个 sheet 最多产出上述两类,全局最多 3 条。
    """
    suggestions: list[dict] = []
    keys: set[tuple[str, str]] = set()

    for sheet in profiles or []:
        if not isinstance(sheet, dict):
            continue
        sheet_name = str(sheet.get("sheet_name") or "").strip()
        columns = [c for c in (sheet.get("columns") or []) if isinstance(c, dict)]
        if not sheet_name or not columns:
            continue
        existing = {str(c.get("name") or "") for c in columns if c.get("name")}

        def _append(name: str, expression: str, description: str) -> None:
            """同一 sheet 内追加一条兜底建议(重复/同名/非法时静默跳过)。"""
            key = (sheet_name, name)
            if key in keys or name in existing:
                return
            try:
                # 与执行阶段共用同一套校验,保证兜底结果一定能跑通
                pe.validate_derived_specs(
                    existing, [{"name": name, "expression": expression}]
                )
            except ValueError:
                return
            suggestions.append({
                "sheet_name": sheet_name,
                "name": name,
                "expression": expression,
                "description": description,
            })
            keys.add(key)

        # —— 数量 × 价格 → 销售额 ——
        qty = _pick_profile_column(columns, _DERIVED_QUANTITY_WORDS, set())
        price = _pick_profile_column(
            columns, _DERIVED_PRICE_WORDS,
            {str(qty.get("name"))} if qty else set(),
        )
        if qty and price:
            _append(
                "销售额",
                f"`{qty['name']}` * `{price['name']}`",
                f"{qty['name']} × {price['name']} 计算销售额",
            )
        if len(suggestions) >= _DERIVED_FALLBACK_MAX:
            break

        # —— 收入 − 成本 → 利润 ——
        revenue = _pick_profile_column(columns, _DERIVED_REVENUE_WORDS, set())
        cost = _pick_profile_column(
            columns, _DERIVED_COST_WORDS,
            {str(revenue.get("name"))} if revenue else set(),
        )
        if revenue and cost:
            _append(
                "利润",
                f"`{revenue['name']}` - `{cost['name']}`",
                f"{revenue['name']} − {cost['name']} 计算利润",
            )
        if len(suggestions) >= _DERIVED_FALLBACK_MAX:
            break

    return suggestions[:_DERIVED_FALLBACK_MAX]


def _normalize_llm_suggestions(
    profiles: list[dict], items: list
) -> list[dict]:
    """过滤/normalize LLM 返回的派生列建议。

    逐条复用 pandas_engine.validate_derived_specs(与真正执行时同一套规则):
        - sheet_name 必须是本次画像里真实存在的表;
        - 公式引用的列必须存在,且建议只引用数值列;
        - 新列名不能为空 / 超 30 字 / 与现有列或本批已接受项重名;
    非法项直接丢弃;按 (sheet_name, name) 去重;最多保留 6 条。
    """
    sheet_columns: dict[str, list[dict]] = {}
    for sheet in profiles or []:
        if not isinstance(sheet, dict):
            continue
        name = str(sheet.get("sheet_name") or "").strip()
        if name:
            sheet_columns[name] = [
                c for c in (sheet.get("columns") or []) if isinstance(c, dict)
            ]

    accepted: list[dict] = []
    used_keys: set[tuple[str, str]] = set()
    # 每个 sheet 已接受的新列名:后续同名建议要视为"与现有列重复"
    accepted_names: dict[str, set[str]] = {}

    for item in items or []:
        if not isinstance(item, dict):
            continue
        sheet_name = str(item.get("sheet_name") or "").strip()
        if sheet_name not in sheet_columns:
            continue
        columns = sheet_columns[sheet_name]
        types = {
            str(c.get("name") or ""): str(c.get("type") or "text")
            for c in columns
        }
        existing = {str(c.get("name") or "") for c in columns if c.get("name")}
        name = str(item.get("name") or "").strip()
        expression = str(item.get("expression") or "").strip()

        if (sheet_name, name) in used_keys:
            continue

        # 推荐层面多做一层约束:公式引用的列必须是 number。
        # 文本列做四则运算虽然语法合法,但执行时必然整列 NaN,提前丢弃。
        refs = re.findall(r"`([^`]+)`", expression)
        if any(types.get(ref) != "number" for ref in refs):
            continue

        batch_existing = existing | accepted_names.get(sheet_name, set())
        try:
            pe.validate_derived_specs(
                batch_existing, [{"name": name, "expression": expression}]
            )
        except ValueError:
            continue

        accepted_names.setdefault(sheet_name, set()).add(name)
        used_keys.add((sheet_name, name))
        accepted.append({
            "sheet_name": sheet_name,
            "name": name,
            "expression": expression,
            "description": str(item.get("description") or "").strip(),
        })
        if len(accepted) >= _DERIVED_SUGGEST_MAX:
            break

    return accepted


# ---------------------- 透视汇总推荐 ----------------------

def _score_dim_column(col: dict) -> int:
    """给候选维度列打分,分越低越优先。

    排序优先级: 关键词命中 > 非数值 > 低基数(<=20) > 低基数(<=50) > 其余
    打分用于兜底推荐时"谁来当行维度"的决策。
    """
    name = str(col.get("name") or "")
    unique = int(col.get("unique_count") or 10**9)
    type_ = str(col.get("type") or "text")
    hit_kw = any(kw in name for kw in _PIVOT_DIM_WORDS)
    score = 0
    score += 0 if hit_kw else 100
    score += 0 if type_ != "number" else 50
    score += 0 if unique <= 20 else (20 if unique <= 50 else 100)
    score += unique
    return score


def _score_value_column(col: dict) -> int:
    """给候选数值列打分,分越低越优先(越像"汇总指标")。"""
    name = str(col.get("name") or "")
    non_null = int(col.get("non_null") or 0)
    hit_kw = any(kw in name for kw in _PIVOT_VALUE_WORDS)
    score = 0
    score += 0 if hit_kw else 100
    score -= non_null / 10000.0  # 非空行数越高越好
    return score


def _fallback_pivot_suggestions(profiles: list[dict]) -> list[dict]:
    """LLM 不可用时的透视汇总规则兜底。

    每个 sheet 最多产出 2 条:
        1. 单列维度(非数值或低基数) × 最佳数值指标 → sum
        2. 若存在 date 列,date 作为行维度 × 最佳数值指标 → sum(覆盖时间趋势)
    全局最多 _PIVOT_FALLBACK_MAX * 2 条,避免把后端 Prompt 撑爆。
    """
    suggestions: list[dict] = []
    seen: set[tuple[str, str, str, str]] = set()

    for sheet in profiles or []:
        if not isinstance(sheet, dict):
            continue
        sheet_name = str(sheet.get("sheet_name") or "").strip()
        columns = [c for c in (sheet.get("columns") or []) if isinstance(c, dict)]
        if not sheet_name or not columns:
            continue

        number_cols = [c for c in columns if c.get("type") == "number"]
        # 维度候选:非数值列 + 基数低的数值列(<=50),排除数值但基数极高者
        dim_candidates = [
            c for c in columns
            if c.get("type") != "number" or int(c.get("unique_count") or 10**9) <= 50
        ]
        if not number_cols or not dim_candidates:
            continue

        # 选最像"指标"的数值列
        value_col = sorted(number_cols, key=_score_value_column)[0]
        # 选最像"维度"的列
        dim_col = sorted(dim_candidates, key=_score_dim_column)[0]

        def _add(row_dim_col: dict, value_col: dict, agg: str = "sum") -> None:
            key = (sheet_name, row_dim_col.get("name"), value_col.get("name"), agg)
            if key in seen:
                return
            seen.add(key)
            row_dim = str(row_dim_col.get("name") or "")
            value = str(value_col.get("name") or "")
            suggestions.append({
                "sheet_name": sheet_name,
                "row_dim": row_dim,
                "col_dim": None,
                "value": value,
                "agg": agg,
                "include_margin": True,
                "description": f"按{row_dim}汇总{value}",
            })

        _add(dim_col, value_col, "sum")

        # 若存在日期列,日期做行维度再补一条(覆盖时间趋势分析)
        date_cols = [c for c in columns if c.get("type") == "date"]
        if date_cols:
            _add(date_cols[0], value_col, "sum")

        if len(suggestions) >= 12:  # 全局硬上限 12,避免极端数据把列表撑爆
            break

    return suggestions


def _normalize_pivot_llm_suggestions(
    profiles: list[dict],
    items: list,
    derived_names_by_sheet: dict[str, set[str]] | None = None,
) -> list[dict]:
    """过滤 / normalize LLM 返回的透视建议,逐条对齐后端 build_pivot_table 的约束。

    新增 derived_names_by_sheet:把 LLM 自己推荐的派生列名按 sheet 分组传入,
    这样 LLM 推荐的透视 value 如果指向"自推荐派生列"也能通过校验。
    执行顺序:clean → derived → pivot,所以后端 clean 接口里派生列先算好,
    透视自然就能用派生列做 value。
    """
    sheet_columns: dict[str, dict[str, dict]] = {}
    for sheet in profiles or []:
        if not isinstance(sheet, dict):
            continue
        name = str(sheet.get("sheet_name") or "").strip()
        if name:
            sheet_columns[name] = {
                str(c.get("name") or ""): c
                for c in (sheet.get("columns") or [])
                if isinstance(c, dict) and c.get("name")
            }
    derived_names_by_sheet = derived_names_by_sheet or {}

    accepted: list[dict] = []
    seen: set[tuple[str, str, str, str, str]] = set()

    for raw in items or []:
        if not isinstance(raw, dict):
            continue
        sheet_name = str(raw.get("sheet_name") or "").strip()
        row_dim = str(raw.get("row_dim") or "").strip()
        value = str(raw.get("value") or "").strip()
        agg = str(raw.get("agg") or "sum").strip()
        col_dim = raw.get("col_dim")
        include_margin = raw.get("include_margin", True)

        cols = sheet_columns.get(sheet_name)
        if not cols or not row_dim or not value:
            continue
        if agg not in _PIVOT_AGG_WHITELIST:
            continue
        # row_dim 必须在原始列里存在(维度列不存在派生的场景)
        if row_dim not in cols:
            continue
        # value 的合法来源:原始数值列 或 LLM 自推荐的派生列
        derived_here = derived_names_by_sheet.get(sheet_name, set())
        value_ok = (
            (value in cols and cols[value].get("type") == "number")
            or value in derived_here
        )
        if not value_ok:
            continue
        # col_dim 若有值,必须存在于该 sheet(列维度也不允许用派生列,保持简洁)
        if col_dim not in (None, "", "null") and str(col_dim).strip() not in cols:
            continue

        # 去重键:同 sheet / 同 row_dim / 同 col_dim / 同 value / 同 agg 视为同一条
        col_key = "" if col_dim in (None, "", "null") else str(col_dim).strip()
        key = (sheet_name, row_dim, col_key, value, agg)
        if key in seen:
            continue
        seen.add(key)

        accepted.append({
            "sheet_name": sheet_name,
            "row_dim": row_dim,
            "col_dim": col_key or None,
            "value": value,
            "agg": agg,
            "include_margin": bool(include_margin),
            # 标记这条建议的 value 是否来自派生列,便于前端 UI 展示差异化
            "value_is_derived": value in derived_here,
            "description": str(raw.get("description") or "").strip()
                or f"按{row_dim}汇总{value}",
        })
        if len(accepted) >= _PIVOT_SUGGEST_MAX:
            break

    return accepted


def suggest_derived_columns(profiles: list[dict]) -> dict[str, Any]:
    """根据字段画像推荐派生列与透视汇总,返回 {suggestions, pivot_suggestions, fallback}。

    说明:这里刻意走 ``llm.chat`` + ``_safe_json_loads``,而不是
    ``llm.chat_json`` —— 后者的契约是「必须返回含 sections 的报告 JSON」,
    对普通结构化结果会把合法输出当成失败并触发 3 次报告式重试。推荐结果
    是非报告结构,用宽容 JSON 解析更符合语义,也与 nl_to_pandas_query 等
    现有非报告 LLM 调用的约定一致。

    降级策略:LLM 异常 / 非 JSON / 空结果 → 规则兜底;兜底也给不出则对应列表为空。
    fallback=True 表示本次至少有一类建议走了规则兜底(对应当前弹窗需降级提示)。
    绝不抛 500 给调用方。
    """
    profiles = profiles or []
    if not profiles:
        return {
            "suggestions": [],
            "pivot_suggestions": [],
            "fallback": True,
        }

    derived: list[dict] = []
    pivot: list[dict] = []
    used_fallback = False
    try:
        raw = llm.chat(
            _DERIVED_SUGGEST_SYSTEM,
            _build_derived_suggest_prompt(profiles),
            json_mode=True,
            temperature=0.2,
            # 推荐结果本身很短,但 reasoning 模型会先消耗思考配额,
            # 2048 容易被截断成空输出,这里留足 4096。
            max_tokens=4096,
        )
        data = _safe_json_loads(raw)
        if isinstance(data, dict):
            if isinstance(data.get("suggestions"), list):
                derived = _normalize_llm_suggestions(profiles, data["suggestions"])
            # 先把 derived names 按 sheet 分组,让 normalize_pivot 允许用派生列作 value
            derived_names_by_sheet: dict[str, set[str]] = {}
            for s in derived:
                sn = str(s.get("sheet_name") or "")
                nm = str(s.get("name") or "")
                if sn and nm:
                    derived_names_by_sheet.setdefault(sn, set()).add(nm)
            if isinstance(data.get("pivot_suggestions"), list):
                pivot = _normalize_pivot_llm_suggestions(
                    profiles, data["pivot_suggestions"],
                    derived_names_by_sheet=derived_names_by_sheet,
                )
    except Exception as exc:
        logger.warning("AI 派生/透视推荐失败,改用规则兜底: %s", exc)

    if not derived:
        derived = _fallback_derived_suggestions(profiles)
        used_fallback = True
    if not pivot:
        pivot = _fallback_pivot_suggestions(profiles)
        used_fallback = True

    return {
        "suggestions": derived,
        "pivot_suggestions": pivot,
        "fallback": used_fallback,
    }


# ---------------------- 字段类型辅助识别 ----------------------

def detect_columns_with_llm(df) -> dict[str, str]:
    """在 Pandas 推断基础上，用 LLM 二次修正。"""
    base = pe.detect_column_types(df)
    # 简单列数太多则跳过 LLM
    if len(df.columns) > 20:
        return base
    return base


# ---------------------- PDF 工具 AI 增强 ----------------------
# 给 PDF 抽取工具加 AI "识别表头 / 清洗文本" 两块能力,
# 增强 pdfplumber 原始输出的人类可读性。注意:这两个函数对失败做"软兜底",
# 任何异常 / 解析失败都 logger.warning + 返回占位结果,绝不抛错影响 PDF 主链路。

_PDF_TABLE_REFINE_SYSTEM = (
    "你是一个 PDF 表格结构识别助手。"
    "你会拿到多张从 PDF 抽取的原始表格,每张表可能因为 PDF 排版导致表头识别错误。"
    "请分析每张表,指出真正的表头所在行(0-based),返回 JSON。"
)

_PDF_CLEAN_TEXT_SYSTEM = (
    "你是一个 PDF 文本清洗助手。"
    "你会拿到从 PDF 逐页抽取的原始文本(含多余换行、页眉页脚、断裂的段落)。"
    "请还原成正常阅读顺序的中文/英文文本,不要解释,直接输出清洗后的全文。"
)


def refine_pdf_tables(tables: list) -> list[dict]:
    """让 LLM 修正 pdfplumber 抽出的原始表格结构,识别真实表头行。

    设计要点:
        - 输入是 ``pdfplumber.extract_tables()`` 的原始输出(三层嵌套列表)，
          体积可能很大(几十张表、每张上百行)，所以这里上限截到 **前 10 张表**
          且 **每张表前 5 行**，避免 prompt 爆掉。
        - 由 LLM 在低温下 (temperature=0.2) 给出真实表头所在行号，
          调用方据此把"表头"vs"数据行"区分开，渲染出来的表更接近原文结构。
        - **任何异常 / 解析失败都 logger.warning + 返回 []**，绝不抛错——
          PDF 主链路容错优先，AI 增强属于"锦上添花"。
    """
    # 第一层防御:空输入直接返回,避免无意义的网络往返
    if not tables:
        return []

    # 截断:前 10 张表 × 每张前 5 行;再把每个 cell 转 str,处理 None 与非字符串
    snippet: list[list[list[str]]] = []
    for ti, tbl in enumerate(tables[:10]):
        if not tbl:
            continue
        cleaned_rows: list[list[str]] = []
        for row in (tbl[:5] if tbl else []):
            if not row:
                continue
            cleaned_rows.append([("" if c is None else str(c)) for c in row])
        if cleaned_rows:
            snippet.append(cleaned_rows)

    if not snippet:
        return []

    # 转 JSON 文本给 LLM 读;ensure_ascii=False 保留中文,体积超限交给 prompt 截断
    try:
        tables_json = json.dumps(snippet, ensure_ascii=False)
    except (TypeError, ValueError) as exc:
        logger.warning("refine_pdf_tables: 序列化失败(%s),跳过 LLM", exc)
        return []

    user = (
        f"待识别的表格共 {len(snippet)} 张(每张仅给出前 5 行样例),"
        f"实际表头可能在这些行内,也可能在被截掉的更下方。\n"
        f"请基于内容语义(字段名/类型/单位/合计行等线索)判断真正的表头行号。\n\n"
        f"表格数据(JSON,外层列表每项是一张表,内层列表每项是一行单元格):\n{tables_json}\n\n"
        "请严格只返回以下 JSON,不要解释、不要 markdown 标记:\n"
        "{\n"
        '  "hints": [\n'
        '    {"table_index": 0, "header_row_index": 0, "merged_cells_fix": []},\n'
        "    ...\n"
        "  ]\n"
        "}\n"
        "- table_index: 0-based,对应当前 snippet 中的第几张表;\n"
        "- header_row_index: 0-based,表头所在行号(必须是上面给出范围内最合理的一行);\n"
        "- merged_cells_fix: 数组,记录被错误切断或合并的单元格提示,无法判断就留 []。"
    )

    try:
        data = llm.chat_json(_PDF_TABLE_REFINE_SYSTEM, user, temperature=0.2)
        # 第一层校验:必须是 dict 且包含 hints 列表
        if not isinstance(data, dict):
            logger.warning("refine_pdf_tables: LLM 返回非 dict(%s)", type(data).__name__)
            return []
        hints = data.get("hints")
        if not isinstance(hints, list):
            logger.warning("refine_pdf_tables: hints 字段不是 list(%s)", type(hints).__name__)
            return []

        # 第二层校验:逐条规整,过滤 schema 不符的条目,避免前段拿脏数据崩渲染
        valid: list[dict] = []
        for idx, item in enumerate(hints):
            if not isinstance(item, dict):
                continue
            try:
                table_index = int(item.get("table_index", idx))
            except (TypeError, ValueError):
                table_index = idx
            try:
                header_row = int(item.get("header_row_index", 0))
            except (TypeError, ValueError):
                header_row = 0
            merged = item.get("merged_cells_fix")
            if not isinstance(merged, list):
                merged = []
            valid.append({
                "table_index": table_index,
                "header_row_index": max(0, header_row),
                "merged_cells_fix": [str(m) for m in merged],
            })
        return valid
    except Exception as exc:
        # 兜底:AI 不可用 / 网络异常 / JSON 解析失败 → 返回空列表,主链路继续
        logger.warning("refine_pdf_tables LLM 调用失败,跳过表头修正: %s", exc)
        return []


def clean_pdf_text(raw_text: str) -> tuple[str, bool]:
    """让 LLM 把 PDF 抽取出的原始文本清洗成正常阅读顺序。

    处理目标:
        1. 去除页眉 / 页脚 / 页码 / 重复标题
        2. 合并多余空行与异常换行
        3. 还原被 PDF 断行切断的段落("自然段")
        4. 保留大小标题层级(用 Markdown 风格的 #、## 等)

    返回值:
        ``(cleaned_text, ai_fallback)`` 元组:
        - ``cleaned_text``:清洗后的文本(若 LLM 失败或返回空,退回原文本);
        - ``ai_fallback`` :布尔,``True`` 表示走了兜底分支(原文本 / 为空),
          供调用方统计 AI 可用率。
    """
    # 第一层防御:输入为空就根本不要打 LLM,直接告诉调用方"空 + fallback"
    if not raw_text or not raw_text.strip():
        return ("", True)

    # 文本过长截断,避免撑爆 prompt;清洗是上下文敏感任务,
    # 经验值 12000 字内能覆盖绝大多数年度报告/招股说明书段落
    snippet = raw_text.strip()[:12000]

    user = (
        "请按以下要求清洗下面的 PDF 抽取文本:\n"
        "  1. 去掉页眉、页脚、页码、重复出现的标题;\n"
        "  2. 把被断行切断的段落合并为完整段落;\n"
        "  3. 多个空行合并为一个空行;\n"
        "  4. 保留并还原标题层级(可用 Markdown # / ## 标记);\n"
        "  5. 不要添加任何解释、评论、摘要,直接输出清洗后的全文。\n\n"
        f"原文:\n{snippet}"
    )

    try:
        cleaned = llm.chat(_PDF_CLEAN_TEXT_SYSTEM, user, json_mode=False, temperature=0.2)
    except Exception as exc:
        # LLM 调用层异常 → 降级返回原文,主链路无感
        logger.warning("clean_pdf_text LLM 调用失败,返回原文: %s", exc)
        return (raw_text, True)

    # 二次清洗:reasoning 模型可能把 <think> 思考块塞进正文,
    # 与 ai_service 其它入口保持一致的清理策略(剥除思考链)
    cleaned = strip_thinking(cleaned or "").strip()

    # 空结果也走 fallback,保证 UI 拿到的不为空,或者显式告知"AI 没清出来"
    if not cleaned:
        logger.warning("clean_pdf_text LLM 返回空结果,返回原文")
        return (raw_text, True)

    return (cleaned, False)


# ---------------------- 照片 OCR 视觉识别 ----------------------
# 背景:用户直接上传若干张照片(现场看板、手写单据、纸质报表翻拍),
# 需要把它们转成可读文本。与 clean_pdf_text / refine_pdf_tables 一样,
# 这里只做"AI 增强",任何异常都不外抛 —— 单张识别失败只标记该位置,
# 不影响同批其它图片。
#
# 设计取舍(为什么不一次性把所有图丢给模型):
#   1. 错误粒度:一次多图时,模型内部对某张图的误读/漏读无法归因,
#      调用方拿不到"第 3 张识别失败"这种可定位的信息;
#   2. 上下文干扰:多图混排容易让模型把 A 图的文字"串"到 B 图上,
#      逐图单问更干净;
#   3. 超时可控:与 file_extractor._extract_pdf 的逐页策略保持一致,
#      一次失败只损失一张图的识别时间。

# 图片后缀 → MIME 类型。
# 与 file_extractor._IMAGE_MIME 保持同一份事实来源:chat_vision 内部拼的是
# `data:{mime};base64,{b64}`,mime 填错会让中转/官方端解析图片失败,
# 所以必须按真实后缀取,而不是一律 image/png。
_OCR_IMAGE_MIME: dict[str, str] = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".bmp": "image/bmp",
}

# 视觉调用的单次超时(秒)。
# 为什么不用 llm_client.chat_vision 的默认 180s:
#   视觉推理的**正常**耗时实测 8~40s,180s 只是"防服务端抽风"的兜底。
#   真的碰上卡死的连接时,干等 180s 才重试会让用户看 3 分钟以上的转圈。
#
# 为什么 2026-10-26 从 75s 提到 95s:
#   日志发现 MiniMax-M3.1-Flash 偶尔会返回 >75s 的"正常慢响应"
#   (不是卡死,只是推理久一点),75s 会误伤它们,导致"白等一轮超时 + 白做
#   两次重试"。95s 比 75s 多 20s 余量,正常慢响应不会被误杀;又远小于 180s,
#   真卡死时仍能快速失败 + 退避重试(3 次总耗时 ≤ 95+3+95+4+95 ≈ 292s,
#   比单次等满 180s + 再做一轮还短)。
_OCR_VISION_TIMEOUT = 95.0

# OCR prompt:模块级常量,便于后续统一调整措辞 / 复用到其它视觉入口。
# 注意 chat_vision 只有 prompt 一个入参、没有 system 形参,所以"角色设定"
# 只能并入 user prompt 开头(与 file_extractor._extract_pdf 的写法一致)。
# 约束点与 clean_pdf_text 相反 —— OCR 要求"纯文本输出",不要 markdown:
# 因为下游会把文本拼回正文/数据集预览,markdown 标记会污染最终展示。
_OCR_USER_PROMPT = (
    "你是一个高精度 OCR 识别引擎,只做逐字转写,不做任何解释、总结或评价。\n"
    "请逐行转写这张图片中的全部文字,要求:\n"
    "1. 保持原始阅读顺序(从上到下、从左到右),不要重排、不要合并、不要漏行;\n"
    "2. 只输出识别出的文字本身,**不要**添加任何解释性前后缀"
    "(不要说「以下是图片内容」「识别结果如下」之类);\n"
    "3. **不要**使用任何 markdown 标记:不要 ``` 代码块包裹、不要 # 标题、"
    "不要 ** 加粗、不要表格竖线;\n"
    "4. **手写文字与印刷文字都要转写** —— 不要因为是手写体就跳过整段,"
    "也不要输出「手写体无法识别」这类说明;\n"
    "5. 确实无法辨认的部分写 [无法辨认],**严禁臆造**原文里没有的内容;\n"
    "6. 保留段落之间的换行,段落内该断行的地方也照原样断行。"
)


def _ocr_one(idx: int, filename: str, raw_bytes: bytes) -> tuple[str | None, str | None]:
    """识别单张图片,返回 ``(text, error)``;**内部兜住所有异常,绝不外抛**。

    抽成独立函数是为了让 ocr_images 能把它丢进线程池并发执行。
    这里只关心"这一张"的结果,批量编排与结果拼装都在调用方。

    Args:
        idx: 该图在原始批次中的下标(仅用于日志,便于定位第几张出问题)。
        filename: 原文件名,决定 OCR 前缀的 mime 判定与日志可读性。
        raw_bytes: 图片原始字节。

    Returns:
        - 成功且有文字: ``(文本, None)``
        - 成功但无文字: ``("", None)``   —— 合法结果,不是错误
        - 失败:         ``(None, 原因摘要)``
    """
    # ========== 预处理只做一次 ==========
    try:
        # 【2026-10 修复·必须先压缩再发】直接把原始字节丢给视觉模型会被拒。
        # 实测:手机拍的 12MP 合同照片约 20MB,base64 后 27MB,MiniMax 直接返回
        # 400 "media exceeds size limit: max 10485760 bytes",而且这种超大请求
        # 要等很久才返回,用户看到的就是按钮一直转圈。
        # 这里复用 pdf_service._prepare_image:EXIF 校正 + 长边缩到 ≤1600px
        # + 转 RGB,压到几百 KB 再编码,既满足对端 10MB 限制,也顺带把
        # 上传体积与推理耗时一起降下来。
        # 放在函数内 import 是为了避免 ai_service ← pdf_service 的循环依赖。
        from app.services.pdf_service import _prepare_image as _prep_for_ocr

        img = _prep_for_ocr(raw_bytes)
        # 重编码为 JPEG:PNG 无损,同样像素体积大 3~5 倍。
        # quality=85 对 OCR 精度足够(文字是高频细节),能让请求体再小一个数量级。
        import io as _io

        buf = _io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
        jpg_bytes = buf.getvalue()
        # mime 必须是 image/jpeg —— 数据真的变成 JPEG 了,
        # 若还按原后缀报 png/webp,会让对端按错误的魔数嗅探,徒增识别失败概率。
        b64 = base64.b64encode(jpg_bytes).decode("ascii")
    except Exception as exc:
        # 预处理失败(图片损坏等):不是网络问题,重试无意义,直接返回
        reason = f"图片预处理失败: {type(exc).__name__}: {exc}"
        logger.warning("ocr_images: 第 %d 张(%s)%s", idx + 1, filename, reason)
        return (None, reason)

    # ========== 网络调用:1 次退避重试 ==========
    # 【2026-10-26 重写】从"零重试"改为"首次 + 1 次重试",原因:
    #   - 根因:老代码零重试,单次超时直接永久失败 —— 这正是"纯文字形态
    #     经常失败"的直接原因。
    #   - 为什么只加 1 次:日志显示 _ocr_layout_one 第一次超时后第二次
    #     就通了,2 次足够覆盖 95% 的抖动场景。3 次的最坏耗时
    #     (3 × 95s ≈ 4 分半)会让用户等过久。
    #   - 区分错误类型(重要):
    #     • APITimeoutError / APIConnectionError → 网络抖动,值得重试
    #     • 429(配额耗尽)/ 401(鉴权失败)/ 403(权限) → 永久性失败,
    #       重试毫无意义,立即放弃。实测对 429 做了三次无意义重试,
    #       白白浪费 2 次调用额度。
    #     • BadRequestError(400)/ NotFoundError(404) → 参数/模型错误,
    #       也立即放弃。
    import time as _t
    from openai import (
        APITimeoutError, APIConnectionError,
        AuthenticationError, PermissionDeniedError,
        BadRequestError, NotFoundError,
    )
    # 值得重试的异常类型(只有网络抖动类)
    _RETRYABLE = (APITimeoutError, APIConnectionError)
    # 明确不该重试的永久性异常(对端立即返回,重试毫无意义)
    _PERMANENT = (AuthenticationError, PermissionDeniedError,
                  BadRequestError, NotFoundError)

    last_err = "未知错误"
    for attempt in range(2):
        try:
            out = llm.chat_vision(
                _OCR_USER_PROMPT,
                images_b64=[b64],
                mime="image/jpeg",
                # 长文档一页文字可能很多,chat_vision 默认 4096 会把后半页截断
                max_tokens=8192 if attempt == 0 else 16384,
                timeout=_OCR_VISION_TIMEOUT,
                # M3.1-Flash 强制 adaptive thinking 且不允许关闭;
                # OCR 是模式识别任务,压低思考链即可避免"思考耗光 max_tokens
                # → 空输出"和"思考过久 → read timeout"两个故障(实测见
                # llm_client.chat_vision 的 docstring)。
                reasoning_effort="low",
            )
        except _RETRYABLE as exc:
            # 超时/连接抖动 → 退避重试
            reason = f"{type(exc).__name__}: {exc}"
            last_err = reason
            if attempt == 0:
                _t.sleep(2)  # 2s 退避,避开瞬时抖动
                logger.info("ocr_images: 第 %d 张(%s)超时/连接失败,重试: %s",
                            idx + 1, filename, reason)
                continue
            logger.warning("ocr_images: 第 %d 张图片(%s)识别失败(超时重试后仍失败),降级跳过: %s",
                           idx + 1, filename, reason)
            return (None, reason)
        except _PERMANENT as exc:
            # 429/401/403/400 等 → 立即放弃,不重试
            reason = f"{type(exc).__name__}: {exc}"
            logger.warning("ocr_images: 第 %d 张图片(%s)识别失败(永久性错误,不重试): %s",
                           idx + 1, filename, reason)
            return (None, reason)
        except Exception as exc:
            # 其它未知异常:保守起见也放弃,避免无限循环
            reason = f"{type(exc).__name__}: {exc}"
            logger.warning("ocr_images: 第 %d 张图片(%s)识别失败(未知异常): %s",
                           idx + 1, filename, reason)
            return (None, reason)

        # reasoning 模型可能把 <think> 思考块混进正文,
        # 与 ai_service 其它入口(summarize_report / clean_pdf_text)
        # 保持同一套清理策略
        cleaned = strip_thinking(out or "").strip()
        if cleaned:
            return (cleaned, None)

        # 空输出:合法结果(这张图里确实没有可识别文字)
        # 与"识别失败"必须区分开 —— 调用方据此决定展示空态而非错误态。
        # 只在**第一次**拿到空输出时做一次"提高 max_tokens 重试":
        # 以防是被截断而非真的没文字
        if attempt == 0:
            last_err = "模型返回空内容(可能被 max_tokens 截断)"
            logger.info("ocr_images: 第 %d 张(%s)第 1 次返回空,提高额度重试",
                        idx + 1, filename)
            continue
        logger.info("ocr_images: 第 %d 张图片(%s)未识别到文字", idx + 1, filename)
        return ("", None)

    logger.warning("ocr_images: 第 %d 张图片(%s)识别异常: %s",
                   idx + 1, filename, last_err)
    return (None, last_err)


def ocr_images(
    images: list[tuple[str, bytes]],
) -> tuple[list[str | None], list[str | None]]:
    """用视觉模型对一批图片做 OCR 文字识别,单张失败只标记该位置不中断整批。

    入参:
        images: ``[(原文件名, 图片字节), ...]``,按用户上传顺序排列。
                这里收 bytes 而不是路径,是调用方(上传接口)已经把文件读进内存,
                避免本函数再去碰文件系统,保持"纯计算"、易于测试。

    返回值:
        ``(texts, errors)`` 两个**等长同序**的列表,长度与入参一致:
        - ``texts[i]``: 识别出的文本;
            ``None`` 表示该图识别失败(此时必然 ``errors[i]`` 非 None);
            ``""``  表示识别成功但图中确实没有文字(合法的"空结果",不是错误)。
        - ``errors[i]``: 失败原因摘要(可直接展示给用户);成功时为 ``None``。

    软降级约定(与 clean_pdf_text / refine_pdf_tables 一致):
        本函数**绝不向外抛异常**。超时 / 连接错误 / 鉴权失败 / 模型不存在 /
        模型不支持视觉输入等任何异常,都只会把该位置置成
        ``(None, 原因摘要)`` 并 logger 记录,然后继续处理下一张图。
        因为"识别 10 张里有 1 张失败"不该让整次上传变成 500。
    """
    # 预分配结果容器:先用 None / None 兜底,后面按 i 逐位覆盖。
    # 这样即使中途意外 return,调用方拿到的也是等长列表而不是半截结果。
    n = len(images or [])
    texts: list[str | None] = [None] * n
    errors: list[str | None] = [None] * n

    # 空输入直接返回两个空列表,不打网络请求
    if n == 0:
        return (texts, errors)

    _OCR_TODO: list[tuple[int, str, bytes]] = []

    for i, (filename, raw_bytes) in enumerate(images or []):
        # 第一层防御:bytes 为空/None 的脏数据不该走到网络层
        if not raw_bytes:
            errors[i] = "图片内容为空"
            logger.warning("ocr_images: 第 %d 张图片(%s)内容为空,跳过识别", i + 1, filename)
            continue
        # 只把"待办索引 + 数据"交给线程池,真正的网络调用在 _ocr_one 里
        _OCR_TODO.append((i, filename, raw_bytes))

    # ===== 【2026-10 并发化】逐图串行 → 线程池并发 =====
    # 为什么改:串行时总耗时 = 单张 × 张数,而单张视觉推理实测 3~150 秒
    # (服务端负载波动极大)。10 张串行最坏要等十几分钟,前端 axios
    # 120s 早就超时,用户看到的就是"转圈很久然后失败"。
    # 并发后总耗时 ≈ max(单张),10 张和 1 张几乎一样快。
    #
    # 为什么用线程池而不是 asyncio:chat_vision 内部是**同步阻塞**的
    # httpx.Client.post,GIL 在等待 socket 时会释放,所以纯网络等待型任务
    # 用线程池是安全的(这也是项目里 analyzer 曾经用 ThreadPoolExecutor
    # 并发跑 summarize/annotate 的同一套思路)。
    #
    # 并发度为什么只给 4:再高对单个视觉模型没有意义(对端有并发限流,
    # 超过一定数量反而排队更久甚至 429),4 足以把"串行 N 倍"压成"约 1 倍"。
    if _OCR_TODO:
        from concurrent.futures import ThreadPoolExecutor, as_completed

        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {
                pool.submit(_ocr_one, idx, fname, raw): idx
                for idx, fname, raw in _OCR_TODO
            }
            for fut in as_completed(futures):
                idx = futures[fut]
                # _ocr_one 内部已兜住所有异常,这里再包一层只是防御
                try:
                    text, err = fut.result()
                except Exception as exc:  # pragma: no cover
                    text, err = None, f"{type(exc).__name__}: {exc}"
                texts[idx] = text
                errors[idx] = err

    logger.info(
        "ocr_images: 完成 %d 张,成功 %d / 失败 %d",
        n, sum(1 for t in texts if t is not None), sum(1 for e in errors if e),
    )
    return (texts, errors)


# ============================================================
# 结构化版面 OCR(2026-10 新增)—— 为「照片 → Word 风格 PDF 重排」供料
# ============================================================
#
# 【为什么必须有这一层:照片"修补"路线已被用户实测否决】
#   旧路线(ocr_images + _gen_dual_pdf)是"修补照片":去底色 → 二值化 → 白底黑字
#   → 再叠一层不可见文字层。它有三条**无法用调参绕过**的死结:
#     1. 照片里的**非文档物体**会被一并二值化 —— 拍合同时的笔记本边缘、
#        桌上的绿笔、书本线条,都会变成黑色方块留在"白底文档"里。
#        用户截图里红框标出的正是这些,这不是阈值能救的:它们本来就是"深色",
#        和文字在灰度空间里同属前景。
#     2. 文字是**照片的一部分**,带透视倾斜。二值化只能改像素,
#        改不了"这一行文字是斜的"这个几何事实 —— 不做整页透视重采样,
#        不可能"扶正";而做了整页透视重采样,又回到了"还是一张被拍歪的图"。
#     3. 字号大小不一、行距不齐,是拍摄透视 + 手工排版共同造成的,
#        任何图像处理都**不可能**凭空产生 Word 的排版节奏。
#   结论:唯一的出路是**丢弃照片的"图像"属性**,让视觉模型输出
#   "这块文档由哪些语义块、按什么顺序组成"的结构化信息,
#   再交给 PDF 排版引擎(reportlab)从零重排 —— 文字天然正向横排、
#   字号统一、自动换行,和 Word 导出的效果同类。
#
# 【与 ocr_images 的关系:并存,不是替换】
#   ocr_images 继续服务"纯文字识别"场景(数据集预览、可搜索文字层),
#   ocr_layout 服务"重排"场景。两者共享同一套软降级 + 线程池并发约定,
#   便于调用方按形态自由选择。

# 版面 JSON 的四个 block 类型白名单。
# 为什么只认这四种:它们是"Word 合同"里出现频率最高、且**排版语义明确**的块。
# 表格、多栏、图片说明等形态本轮**故意不支持** —— 模型对表格的结构化还原
# (合并单元格、跨行)极易出错,与其让调用方拿到一份错位的"表格 JSON",
# 不如先老老实实只支持流式文字块,把能力边界划清楚。
_OCR_LAYOUT_BLOCK_TYPES: frozenset[str] = frozenset(
    {"paragraph", "kv", "signature", "date"}
)

# 结构化版面 OCR 的 prompt(模块级常量,与 _OCR_USER_PROMPT 同一套组织方式)。
# chat_vision 只有 prompt 一个入参、没有 system 形参,所以"角色设定"必须并入开头。
#
# 【为什么 Prompt 里"明确禁止输出坐标"】—— 这是刻意设计,不是省 token:
#   1. 上游图像几何**不可信**:手机拍的合同有透视,照片里的"矩形"在真实
#      版面上是梯形;模型给出的 bbox 只是它在**倾斜照片**上目测的框,
#      换算回"Word 版面坐标"必然有系统性偏差。
#   2. 一旦允许坐标,模型就会开始"对齐坐标",而我们又必须**丢弃**它,
#      反而引入噪声;更糟的是模型可能为了凑坐标而臆造版面(把一段话拆成
#      三个框、把同一行切成两段),反而污染阅读顺序。
#   3. 我们的排版引擎(reportlab)只需要**流式顺序**:按 blocks 数组顺序
#      往下排即可,版式由引擎决定。坐标是"像素事实",顺序才是"内容事实",
#      而重排要的恰恰是内容事实。
#   所以这里不但禁止,还要在末尾再强调一次 —— 模型对"多给点信息"有很强倾向。
_OCR_LAYOUT_USER_PROMPT = (
    "你是一个文档版面解析引擎,把纸质文档的照片解析成结构化的版面 JSON。"
    "你的输出会被程序**逐块重新排版**,不是给人看的描述。\n"
    "严格只输出一个 JSON 对象:第一个字符必须是 `{`,最后一个字符必须是 `}`,"
    "禁止用 ```json 代码块包裹,禁止任何解释性前后缀。\n"
    "JSON 结构如下:\n"
    '{"title": "文档标题", "blocks": [{"type": "...", "text": "...", '
    '"label": "...", "value": "..."}]}\n'
    "字段与规则:\n"
    "1. title 是文档标题(例如「六天双向选择协议书」);没有明确标题就写空字符串。\n"
    "2. blocks 是数组,按文档**从上到下的阅读顺序**排列,顺序就是排版顺序。\n"
    "3. 每个 block 的 type 只能是以下四种之一:\n"
    "   - \"paragraph\": 普通正文段落;\n"
    "   - \"kv\": 「标签 + 值」形式的条目,label 放标签(如「甲方」),"
    "value 放值(如「万华商贸有限公司」),text 字段可省略;\n"
    "   - \"signature\": 签署区的引导语(如「甲方盖章:」「乙方(签字):」);\n"
    "   - \"date\": 日期行(如「日期:      年    月    日」)。\n"
    "3.1 **左右并排的两个条目必须合并成一个 block,不要拆成两个**。\n"
    "   纸质合同里同一个水平行上左右各写一条的情况非常常见"
    "(如左上「甲方:某某公司」、右上「乙方:某某」;"
    "左下「甲方盖章:」、右下「乙方签字:」;左「日期:年月日」、右「日期:年月日」)。\n"
    "   遇到这种情况:左边那条按上面的字段正常填;"
    "右边那条的字段加 right_ 前缀写进**同一个 block**:\n"
    "   - kv 用 right_label / right_value;\n"
    "   - signature / date 用 right_text;\n"
    "   - paragraph 不参与并排。\n"
    "   例子:左上「甲方:宁夏某公司」+ 右上「乙方:___」→ "
    "{\"type\":\"kv\",\"label\":\"甲方\",\"value\":\"宁夏某公司\","
    "\"right_label\":\"乙方\",\"right_value\":\"\"};\n"
    "   例子:左下「甲方盖章:」+ 右下「乙方签字:」→ "
    "{\"type\":\"signature\",\"text\":\"甲方盖章:\",\"right_text\":\"乙方签字:\"}。\n"
    "   判断依据是**左右两个条目在照片上处于同一高度**(不需要输出任何坐标);"
    "只有确实左右并排才合并,上下两行写的两条仍然拆成两个 block。\n"
    "4. **只转写文档本身的文字**。照片里出现的笔记本边缘、桌面上的笔、"
    "书本装订线、桌面纹理、背景杂物、以及任何非文档物体,"
    "一律**不要描述、不要输出、不要猜测**它们是什么 —— 哪怕它们非常显眼。\n"
    "5. 手写批注与印章里的文字都要转写;确实无法辨认的部分写 [无法辨认],"
    "严禁臆造原文中不存在的内容。\n"
    "6. 同一段正文不要拆成多个 paragraph;段落边界用阅读顺序体现即可。\n"
    "7. **印章的位置必须单独标注**。在 seals 数组里为每一个红色印章输出一项,"
    "说明它盖在哪个文字附近:\n"
    "   - \"anchor\": 该印章**正下方或正上方最近的那行文字**(逐字照抄),"
    "程序据此把印章贴回这行文字旁边;\n"
    "   - \"side\": 印章相对这行文字的位置,只能取 \"left\" 或 \"right\";\n"
    "   - \"text\": 印章里能读出的文字(读不出就写空字符串)。\n"
    "   注意:一页上**可能有多个印章**(如甲方处一枚、落款处一枚),"
    "每一个都要单独列出来,不要只报最大的那个。\n"
    "   例子:{\"anchor\":\"甲方盖章\",\"side\":\"left\",\"text\":\"某公司\"}\n"
    "先输出 JSON,不要输出任何其他文字。"
)


def _parse_layout_json(raw: str) -> tuple[dict[str, Any], str | None]:
    """把模型返回的版面 JSON 文本解析 + 校验成规范结构,返回 ``(data, error)``。

    为什么解析要走"两段式":
        ``json.loads`` 对**完全合规**的 JSON 最快、最严格;但视觉模型经常
        违反自己的指令 —— 用 ```json 包裹、单引号属性名、结尾多一句解释、
        或者在 JSON 前后塞思考块。这些情况 ``json.loads`` 会直接抛错,
        而 json_utils._safe_json_loads 正是项目里为"LLM 输出的非完全合规 JSON"
        准备的宽容解析器(chat_json 也在用)。先用严格解析保证"能严格就严格",
        失败再回退宽容解析,是与项目既有范式一致、且不引入新依赖的做法。

    校验为什么放在解析之后:
        模型可能给出合法 JSON 但**结构不对**(blocks 缺失、block 缺 text、
        type 是自创值)。这类结果如果直接交给排版引擎,会渲染出空白页或
        LayoutError。这里逐条过滤非法 block,只要还剩 ≥1 个合法 block
        就仍然算成功 —— 局部错不该让整页作废;全错才算这张图识别失败。

    Returns:
        - 成功: ``({"title": str, "blocks": [{"type","text","label","value"}, ...]}, None)``
        - 失败: ``(None, 失败原因)``
        - 合法但无可识别文字: ``({"title": "", "blocks": []}, None)``
          (这是"这张图里确实没有文档文字"的合法结果,不是错误)
    """
    # 解析前的轻清洗:与既有 OCR 入口一致地先剥思考块/截断标记,
    # 避免 <think> 尾巴把 JSON 的闭合括号吃掉导致解析失败。
    cleaned = strip_thinking(raw or "").strip()

    # 模型返回空字符串时单独给一条可读原因:
    # 直接交给下面的 json.loads 会抛 "Expecting value: line 1 column 1",
    # 而真实病因是"空输出"(多半被 max_tokens 截断),提示语要说人话。
    if not cleaned:
        return (None, "版面 JSON 解析失败: 模型返回空内容(通常是被 max_tokens 截断)")

    # 两段式解析:严格优先,失败回退宽容解析器。
    obj: Any = None
    try:
        obj = json.loads(cleaned)
    except Exception:
        try:
            obj = _safe_json_loads(cleaned)
        except Exception as exc:
            return (None, f"版面 JSON 解析失败: {type(exc).__name__}: {exc}")

    if not isinstance(obj, dict):
        return (None, f"版面 JSON 顶层不是对象: {type(obj).__name__}")

    # ---- title:容错成字符串 ----
    title = obj.get("title")
    title = title.strip() if isinstance(title, str) else ""

    # ---- blocks:必须是 list ----
    blocks_raw = obj.get("blocks")
    if not isinstance(blocks_raw, list):
        # 顶层没有 blocks 数组 = 模型跑偏了。空数组才是"无文字"的合法结果,
        # 这里必须区分开:前者是失败(有 error),后者是成功(无 error)。
        return (None, "版面 JSON 缺少 blocks 数组")

    valid: list[dict[str, Any]] = []
    for item in blocks_raw:
        # 逐块防御:模型可能塞进 null / 字符串 / 嵌套对象等奇形怪状
        if not isinstance(item, dict):
            continue
        btype = item.get("type")
        if not isinstance(btype, str) or btype not in _OCR_LAYOUT_BLOCK_TYPES:
            continue
        # kv 的正文在 label/value 上,其余类型在 text 上 —— 两种来源都要归一到 text。
        text = item.get("text")
        if not isinstance(text, str):
            text = ""
        label = item.get("label")
        label = label.strip() if isinstance(label, str) else ""
        value = item.get("value")
        value = value.strip() if isinstance(value, str) else ""
        if btype == "kv":
            text = value

        # ---- 并排右栏(2026-10 新增,见 _OCR_LAYOUT_USER_PROMPT 第 3.1 条)----
        # 纸质合同里「甲方 | 乙方」「甲方盖章 | 乙方签字」「左日期 | 右日期」
        # 都是同一水平行的左右两栏,模型被要求把它们合并成一个 block,
        # 右栏内容放在 right_* 字段里。这里统一清洗成字符串:
        #   - kv 用 right_label / right_value(right_value 同时并入 right_text,
        #     让排版端只需读一个"右栏正文"即可,不必再区分类型);
        #   - signature / date 直接用 right_text。
        right_text = item.get("right_text")
        right_text = right_text.strip() if isinstance(right_text, str) else ""
        right_label = item.get("right_label")
        right_label = right_label.strip() if isinstance(right_label, str) else ""
        right_value = item.get("right_value")
        right_value = right_value.strip() if isinstance(right_value, str) else ""
        if btype == "kv":
            right_text = right_value or right_text
        # 右栏若只有标签(如「乙方:」后面空着),原文本身就存在这一栏,
        # 必须保留 —— 否则排版出来的合同右下角会漏掉"乙方:"这一处空位。
        if not text and not label and not right_text and not right_label:
            continue
        valid.append({
            "type": btype, "text": text, "label": label, "value": value,
            "right_text": right_text, "right_label": right_label,
        })

    # ---- seals:印章的位置标注(2026-10 新增)----
    # 为什么需要:重排后原文坐标已不存在,若只把章统一贴到末尾,
    # 用户看到的就是"章的位置和原件对不上"。让模型标出"这枚章盖在哪行文字旁边",
    # 程序就能把章贴回它真正所属的那一行。
    seals: list[dict[str, str]] = []
    seals_raw = obj.get("seals")
    if isinstance(seals_raw, list):
        for item in seals_raw:
            if not isinstance(item, dict):
                continue
            anchor = item.get("anchor")
            anchor = anchor.strip() if isinstance(anchor, str) else ""
            if not anchor:
                # 没有锚点的章无法定位,直接丢弃(贴到末尾反而更糟)
                continue
            side = item.get("side")
            side = "left" if str(side).strip().lower() == "left" else "right"
            stext = item.get("text")
            stext = stext.strip() if isinstance(stext, str) else ""
            seals.append({"anchor": anchor, "side": side, "text": stext})

    # 全空 → 视为"该图无可识别文字",返回空 blocks 且**不算失败**
    # (与 _ocr_one 把"无文字"当合法结果的既有约定一致)。
    return ({"title": title, "blocks": valid, "seals": seals}, None)


def _ocr_layout_one(
    idx: int, filename: str, raw_bytes: bytes,
) -> tuple[dict[str, Any] | None, str | None]:
    """结构化识别单张图片的版面,返回 ``(layout, error)``;**内部兜住所有异常**。

    抽成独立函数的原因与 _ocr_one 一致:让 ocr_layout 能把它丢进线程池并发。

    与 _ocr_one 的差别只在"拿到文本之后":
        _ocr_one      → 直接返回纯文本字符串;
        _ocr_layout_one→ 还要把文本解析 + 校验成版面 dict。

    【2026-10 加固·空输出重试】
        实测踩到:模型偶尔返回**空字符串**,于是下游报
        "JSONDecodeError: empty input: line 1 column 1"。
        成因是 reasoning 模型把 max_tokens 全耗在思考过程上,
        正文一个字都没输出就被截断 —— 这与项目里 llm_client.chat_json
        注释记录的"raw 为空 = 被 max_tokens 截断"是同一个病。
        单次重试的代价(一次视觉调用)远小于"整页退化成白底图"，
        所以这里对空输出/解析失败各重试一次,并把上限提到 16384。

    Args:
        idx: 该图在原始批次中的下标(仅用于日志)。
        filename: 原文件名,决定 mime 判定与日志可读性。
        raw_bytes: 图片原始字节。

    Returns:
        - 成功(含空版面): ``(layout_dict, None)``
        - 失败:             ``(None, 原因摘要)``
    """
    # 放在函数内 import 是为了避免 ai_service ← pdf_service 的循环依赖
    # (与 _ocr_one 里的写法保持一致)。
    from app.services.pdf_service import _prepare_image as _prep_for_ocr

    try:
        img = _prep_for_ocr(raw_bytes)
        # 重编码 JPEG:PNG 无损,同样像素体积大 3~5 倍,而对端限制 10MB。
        import io as _io

        buf = _io.BytesIO()
        img.save(buf, format="JPEG", quality=88, optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception as exc:
        reason = f"图片预处理失败: {type(exc).__name__}: {exc}"
        logger.warning("ocr_layout: 第 %d 张(%s)%s", idx + 1, filename, reason)
        return (None, reason)

    # 逐图单问(理由同 _ocr_one:错误粒度细、上下文不串味)。
    #
    # 【2026-10-26 对齐 _ocr_one 的重试策略】
    #   老版三次盲目重试,对 429(配额耗尽)/401(鉴权)这类永久性错误
    #   也做了三次无意义调用 —— 实测浪费了 2 次额度。
    #   现在与 _ocr_one 一致:
    #     • 重试次数降为 2 次(首次 + 1 次,95% 抖动一次过)
    #     • 超时/连接抖动 → 2s 退避重试
    #     • 429/401/403/400 → 永久性错误,立即放弃不重试
    import time as _t
    from openai import (
        APITimeoutError, APIConnectionError,
        AuthenticationError, PermissionDeniedError,
        BadRequestError, NotFoundError,
    )
    _RETRYABLE = (APITimeoutError, APIConnectionError)
    _PERMANENT = (AuthenticationError, PermissionDeniedError,
                  BadRequestError, NotFoundError)

    last_err = "未知错误"
    for attempt in range(2):
        try:
            out = llm.chat_vision(
                _OCR_LAYOUT_USER_PROMPT,
                images_b64=[b64],
                mime="image/jpeg",
                max_tokens=8192 if attempt == 0 else 16384,
                timeout=_OCR_VISION_TIMEOUT,
                # 同 _ocr_one:压低 adaptive thinking 的思考链,
                # 根治"空输出 / read timeout"这对相互掣肘的故障。
                reasoning_effort="low",
            )
        except _RETRYABLE as exc:
            reason = f"{type(exc).__name__}: {exc}"
            last_err = reason
            if attempt == 0:
                _t.sleep(2)
                logger.info("ocr_layout: 第 %d 张(%s)超时/连接失败,重试: %s",
                            idx + 1, filename, reason)
                continue
            logger.warning("ocr_layout: 第 %d 张图片(%s)版面识别失败(超时重试后仍失败),降级跳过: %s",
                           idx + 1, filename, reason)
            return (None, reason)
        except _PERMANENT as exc:
            reason = f"{type(exc).__name__}: {exc}"
            logger.warning("ocr_layout: 第 %d 张图片(%s)版面识别失败(永久性错误,不重试): %s",
                           idx + 1, filename, reason)
            return (None, reason)
        except Exception as exc:
            reason = f"{type(exc).__name__}: {exc}"
            logger.warning("ocr_layout: 第 %d 张图片(%s)版面识别失败(未知异常): %s",
                           idx + 1, filename, reason)
            return (None, reason)

        data, err = _parse_layout_json(out or "")
        if not err:
            return (data, None)

        last_err = err
        if attempt == 0:
            # 空输出或解析失败 → 提高额度再试一次
            logger.info("ocr_layout: 第 %d 张(%s)第 1 次解析失败(%s),提高 max_tokens 重试",
                        idx + 1, filename, err)
            continue

    logger.warning("ocr_layout: 第 %d 张图片(%s)版面解析失败(重试后仍失败): %s",
                   idx + 1, filename, last_err)
    return (None, last_err)


def ocr_layout(
    images: list[tuple[str, bytes]],
) -> tuple[list[dict[str, Any] | None], list[str | None]]:
    """用视觉模型把一批照片解析成**结构化版面**,供 PDF 引擎重排成 Word 风格文档。

    与 :func:`ocr_images` 的约定完全一致(软降级、等长同序返回、线程池并发),
    区别只在返回的是版面 dict 而不是纯文本。

    【为什么必须"丢弃照片"重排,而不是"修补照片"】
        照片里的文字带透视倾斜、字号不一、行距不齐,而且混杂着非文档物体
        (笔记本边缘、桌上的笔、书本线条)。任何**图像处理**都只能在像素层面
        修补,既无法把斜着的文字"扶正",也无法凭空造出 Word 的排版节奏,
        更无法区分"深色的桌面物体"和"深色的文字"。
        所以这里只从照片里取"内容"(文字 + 块类型 + 阅读顺序),
        排版完全交给 reportlab 重做 —— 文字因此天然正向横排、字号统一、
        自动换行,背景里的杂物则被整体丢弃。

    入参:
        images: ``[(原文件名, 图片字节), ...]``,按用户上传顺序排列。

    返回值:
        ``(layouts, errors)`` 两个**等长同序**的列表,长度与入参一致:
        - ``layouts[i]``: ``{"title": str, "blocks": [{"type","text","label","value"}, ...]}``;
            ``None`` 表示该图识别失败(此时必然 ``errors[i]`` 非 None);
            ``{"title": "", "blocks": []}`` 表示识别成功但图中没有文档文字
            (合法的"空结果",不是错误)。
        - ``errors[i]``: 失败原因摘要(可直接展示给用户);成功时为 ``None``。

    软降级约定:
        本函数**绝不向外抛异常**。超时 / 连接错误 / 鉴权失败 / 模型无视觉能力 /
        JSON 解析失败等任何异常,都只会把该位置置成 ``(None, 原因摘要)`` 并
        logger 记录,然后继续处理下一张图。

    并发约定:
        与 ocr_images 一致使用 ``ThreadPoolExecutor(max_workers=4)``。
        chat_vision 是**同步阻塞**的 httpx 调用,等待 socket 时 GIL 会释放,
        纯网络等待型任务用线程池是安全的;并发度只给 4 是因为对端有并发限流,
        再高反而排队更久甚至 429。
    """
    n = len(images or [])
    layouts: list[dict[str, Any] | None] = [None] * n
    errors: list[str | None] = [None] * n

    if n == 0:
        return (layouts, errors)

    # 只把"有内容"的待办交给线程池,空 bytes 在这里就地判失败,
    # 避免把无效请求丢进网络层浪费配额。
    _TODO: list[tuple[int, str, bytes]] = []
    for i, (filename, raw_bytes) in enumerate(images or []):
        if not raw_bytes:
            errors[i] = "图片内容为空"
            logger.warning("ocr_layout: 第 %d 张图片(%s)内容为空,跳过识别", i + 1, filename)
            continue
        _TODO.append((i, filename, raw_bytes))

    if _TODO:
        from concurrent.futures import ThreadPoolExecutor, as_completed

        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {
                pool.submit(_ocr_layout_one, idx, fname, raw): idx
                for idx, fname, raw in _TODO
            }
            for fut in as_completed(futures):
                idx = futures[fut]
                # _ocr_layout_one 内部已兜住所有异常,这里再包一层只是防御
                try:
                    data, err = fut.result()
                except Exception as exc:  # pragma: no cover
                    data, err = None, f"{type(exc).__name__}: {exc}"
                layouts[idx] = data
                errors[idx] = err

    logger.info(
        "ocr_layout: 完成 %d 张,成功 %d / 失败 %d",
        n,
        sum(1 for d in layouts if d is not None),
        sum(1 for e in errors if e),
    )
    return (layouts, errors)
