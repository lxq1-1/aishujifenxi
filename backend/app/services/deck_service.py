"""演示稿服务：把报告重构为演示稿大纲,并做归一化 + 数字溯源 + 缓存。

整体链路:
    report + indicators + chart_option
        → build_chart_catalog()      摊平可用图表清单
        → generate_deck()            调 LLM 产出大纲 JSON
        → _normalize_deck()          归一化(kind 白名单 / 页数 / bullet 长度 / 标题去空)
        → _audit_deck_numbers()      数字溯源(无出处的 bullet 整条剔除)
        → ensure_deck()              落 deck 表(缓存),供导出与预览复用

设计要点:
    1. **不抛异常**:`ensure_deck` 内部全 try/except,失败只写 failed 记录。
       演示稿是"锦上添花",绝不能因为它把报告主链路搞挂。
    2. **数字只删不改**:查不到出处的 bullet 整条删掉,不做数值替换。
       PPT 是汇报材料,宁可少一条也不要错一个数字。
    3. **缓存优先**:命中 ready 记录就不调 LLM;导出时的现场生成只是兜底。
"""
from __future__ import annotations

import json
import logging
import re
from typing import Any

from sqlalchemy import select
from sqlalchemy.orm import Session

from app.core.deck_prompt import SYSTEM_PROMPT_DECK, build_deck_prompt
from app.core.json_utils import clean_llm_text, dumps, json_default
from app.models.deck import Deck
from app.services.llm_client import llm

logger = logging.getLogger(__name__)


# ============================================================
# 归一化常量
# ============================================================
# 允许的页面类型。非法值一律降级为 content。
_DECK_KINDS = {"cover", "agenda", "content", "chart", "metrics", "conclusion"}

# 页数硬上限:演示稿不是报告,超过 15 页就该砍。
_MIN_SLIDES = 1
_MAX_SLIDES = 15

# 每页 bullet 条数与长度上限
_MAX_BULLETS = 5
_MAX_BULLET_LEN = 60

# 演讲备注长度上限
_MAX_NOTES_LEN = 300

# 标题 / 副标题长度上限
_MAX_TITLE_LEN = 60
_MAX_SUBTITLE_LEN = 80


# ============================================================
# 1) 图表清单
# ============================================================
def build_chart_catalog(
    report: dict[str, Any],
    chart_option: dict[str, Any],
) -> list[dict[str, Any]]:
    """把报告的图表摊平成「可用图表清单」。

    形态统一为::

        [{"id": "chart_1", "title": "月度销售趋势", "section": "总体经营概况", "chart": {...}}, ...]

    `id` 稳定且可读(chart_1 / chart_2 …),Prompt 里让 LLM 引用它;
    `title` 用于二次匹配(LLM 可能填标题而不是 id);
    `section` 说明这张图原本挂在哪个章节,便于 LLM 判断该配给哪一页。
    """
    if not isinstance(chart_option, dict):
        chart_option = {}

    # 优先用 report 内嵌的 chart_option(与 _chart_list 的口径一致:
    # 图表编辑 API 写回 DB 前,report 里可能是更新的那份)
    opt = chart_option
    rep_opt = (report or {}).get("chart_option")
    if isinstance(rep_opt, dict) and rep_opt:
        opt = rep_opt

    catalog: list[dict[str, Any]] = []

    def _add(section: str, chart: dict[str, Any]) -> None:
        if not isinstance(chart, dict):
            return
        title = _chart_title_of(chart)
        catalog.append({
            "id": f"chart_{len(catalog) + 1}",
            "title": title,
            "section": section,
            "chart": chart,
        })

    by_section = opt.get("by_section")
    if isinstance(by_section, dict):
        for section, arr in by_section.items():
            if not isinstance(arr, list):
                continue
            for c in arr:
                _add(str(section), c)

    # 旧形态:{charts: [...]} / 顶层 {series: [...]} —— 同样收进清单,
    # 否则老报告的图会让 LLM 没有可引用的对象。
    if not catalog:
        charts = opt.get("charts")
        if isinstance(charts, list):
            for c in charts:
                _add("数据图表", c)
        elif opt.get("series"):
            _add("数据图表", opt)

    return catalog


def _chart_title_of(chart: dict[str, Any]) -> str:
    """取图表标题(ECharts option 里 title 可能是 dict 或 str)。"""
    raw = chart.get("title")
    if isinstance(raw, dict):
        t = raw.get("text") or ""
    elif isinstance(raw, str):
        t = raw
    else:
        t = ""
    t = clean_llm_text(str(t)).strip() if t else ""
    if not t:
        opt = chart.get("option")
        if isinstance(opt, dict):
            raw2 = opt.get("title")
            if isinstance(raw2, dict):
                t = str(raw2.get("text") or "").strip()
            elif isinstance(raw2, str):
                t = raw2.strip()
    return t or "未命名图表"


# ============================================================
# 2) 归一化
# ============================================================
def _clean_one_line(text: Any, limit: int) -> str:
    """清洗成单行短文本:去思考标签、去换行、去首尾空白、超长截断。"""
    if not isinstance(text, str):
        text = "" if text is None else str(text)
    cleaned = clean_llm_text(text)
    cleaned = cleaned.replace("\r", " ").replace("\n", " ").strip()
    cleaned = " ".join(cleaned.split())
    if len(cleaned) > limit:
        # 截断后补省略号,并保证**总长不超过 limit**(spec 要求 ≤60 字 / ≤300 字)
        cleaned = cleaned[: max(1, limit - 1)].rstrip() + "…"
    return cleaned


def _normalize_deck(
    raw: Any,
    report: dict[str, Any],
) -> dict[str, Any] | None:
    """把 LLM 返回的大纲归一化成稳定结构。

    归一化规则(与 spec 一致):
        - kind 白名单,非法值降级为 content;
        - slides 页数裁到 [_MIN_SLIDES, _MAX_SLIDES];
        - 每页 bullets ≤ 5 条、每条 ≤ 60 字;
        - notes ≤ 300 字;
        - title 为空的页直接丢弃;
        - title / subtitle 缺省时用报告标题 / 报告摘要首句兜底。
    """
    if not isinstance(raw, dict):
        return None

    raw_slides = raw.get("slides")
    if not isinstance(raw_slides, list):
        return None

    slides: list[dict[str, Any]] = []
    for item in raw_slides:
        if not isinstance(item, dict):
            continue
        title = _clean_one_line(item.get("title"), _MAX_TITLE_LEN)
        if not title:
            # 标题为空的页没有信息量,直接丢弃(避免 PPT 上出现空标题页)
            continue

        kind = str(item.get("kind") or "content").strip().lower()
        if kind not in _DECK_KINDS:
            kind = "content"

        bullets: list[str] = []
        raw_bullets = item.get("bullets")
        if isinstance(raw_bullets, list):
            for b in raw_bullets:
                text = _clean_one_line(b, _MAX_BULLET_LEN)
                if text:
                    bullets.append(text)
                if len(bullets) >= _MAX_BULLETS:
                    break
        elif isinstance(raw_bullets, str):
            # 模型偶尔把 bullets 写成一段带换行的文本,拆行兜底
            for line in raw_bullets.replace("\r", "\n").split("\n"):
                text = _clean_one_line(line, _MAX_BULLET_LEN)
                if text:
                    bullets.append(text)
                if len(bullets) >= _MAX_BULLETS:
                    break

        notes = _clean_one_line(item.get("notes"), _MAX_NOTES_LEN)
        chart_ref = _clean_one_line(item.get("chart_ref"), 120)

        slides.append({
            "kind": kind,
            "title": title,
            "bullets": bullets,
            "chart_ref": chart_ref,
            "notes": notes,
        })
        if len(slides) >= _MAX_SLIDES:
            break

    if len(slides) < _MIN_SLIDES:
        return None

    # 标题 / 副标题兜底
    deck_title = _clean_one_line(raw.get("title"), _MAX_TITLE_LEN)
    if not deck_title:
        deck_title = _clean_one_line((report or {}).get("title"), _MAX_TITLE_LEN) or "数据分析汇报"

    deck_subtitle = _clean_one_line(raw.get("subtitle"), _MAX_SUBTITLE_LEN)
    if not deck_subtitle:
        summary = (report or {}).get("summary")
        if isinstance(summary, str) and summary.strip():
            first = summary.strip().split("\n")[0]
            deck_subtitle = _clean_one_line(first, _MAX_SUBTITLE_LEN)

    return {
        "title": deck_title,
        "subtitle": deck_subtitle,
        "slides": slides,
    }


# ============================================================
# 3) 数字溯源
# ============================================================
def _build_deck_number_pool(
    report: dict[str, Any],
    indicators: dict[str, Any],
) -> list[float]:
    """构造演示稿专用的"真实数值池"。

    比报告链路的池多两层,原因是演示稿的表达习惯与报告不同:

    1. **报告正文里的数字也算数**
       报告正文经过 analyzer 一轮溯源后,留下来的数字是"报告已认可"的。
       演示稿的要点是从报告正文提炼的,若只认 indicators,
       会把「超额 12.5%」这类正文里已写明的派生比例误判成编造。

    2. **补一份绝对值形态**
       indicators 里的差值是**带符号**的(目标 5000 − 实际 4680 = −320),
       但人写缺口时习惯写正数("还差 320 万元")。
       只认带符号值,就会把完全正确的 320 判成"无出处"。
       这里为每个池元素补上 |x|,让正负两种表述都能溯源成功。

    Args:
        report:     已生成的结构化报告(取 sections 正文里的数字)。
        indicators: 指标快照。

    Returns:
        可用于溯源判定的数值列表(可能为空)。
    """
    try:
        from app.services.analyzer import _collect_indicator_values
    except Exception:
        return []

    pool = list(_collect_indicator_values(indicators))

    # ---- 1) 报告正文里的数字 ----
    sections = (report or {}).get("sections") or {}
    if isinstance(sections, dict):
        for body in sections.values():
            if not isinstance(body, str):
                continue
            for m in re.finditer(r"\d+(?:,\d{3})*(?:\.\d+)?", body):
                try:
                    pool.append(float(m.group(0).replace(",", "")))
                except ValueError:
                    continue

    # ---- 2) 绝对值形态(补正负表述差异)----
    for v in list(pool):
        pool.append(abs(v))

    return pool


def _audit_deck_numbers(
    deck: dict[str, Any],
    report: dict[str, Any],
    indicators: dict[str, Any],
) -> dict[str, Any]:
    """剔除 / 标注引用了「源数据里不存在」的数字。

    做法:复用报告链路已经验证过的溯源机制
        analyzer._collect_indicator_values  → 真实数值池(经 _build_deck_number_pool 扩展)
        analyzer.audit_section_numbers       → 找出正文里无出处的数字

    处置策略:
        - 一条 bullet 里只要含无出处数字 → **整条剔除**(bullet 短,
          留半句反而更别扭);
        - 一页 bullet 全被剔除 → 用一句提示占位,不留空白页;
        - 标题 / 讲稿里的无出处数字 → **只替换该数字**为「〔数值待核〕」,
          其余文字全部保留。标题是 28pt 大标题、讲稿是口播稿,
          两者都是"人当场要念出来"的内容,整段清空等于这一页没内容。
    """
    try:
        from app.services.analyzer import (
            _strip_fabricated_numbers,
            audit_section_numbers,
        )
    except Exception:
        return deck

    pool = _build_deck_number_pool(report, indicators)
    if not pool:
        # 没有指标池(极老数据)→ 无法判定,原样保留
        return deck

    audited_slides: list[dict[str, Any]] = []
    for slide in deck.get("slides") or []:
        if not isinstance(slide, dict):
            continue
        title = slide.get("title") or ""

        # ---- 要点:含无出处数字的整条剔除 ----
        # bullet 是"独立可读的一行",半句带「〔数值待核〕」的 bullet 读起来很别扭,
        # 整条删掉更干净(宁可少一条,也不要错数字)。
        kept: list[str] = []
        for bullet in slide.get("bullets") or []:
            bad = audit_section_numbers(title, bullet, pool)
            if bad:
                logger.info(
                    "[deck] bullet 含无出处数字 %s,已剔除: %s", bad, bullet[:40]
                )
                continue
            kept.append(bullet)

        had_bullets = bool(slide.get("_had_bullets"))
        slide["bullets"] = kept
        # 原本有要点、现在被删光 → 补占位提示,不留空白页
        if had_bullets and not kept:
            slide["bullets"] = ["本页要点未通过数字核对,已省略"]

        # ---- 标题 / 讲稿:替换可疑数字,而不是整段丢弃 ----
        # 标题是 28pt 大标题、讲稿是口播稿,两者都是"人当场要念出来"的内容,
        # 整段清空等于这一页没内容 —— 代价太大。
        # 这里只把**个别**可疑数字替换成「〔数值待核〕」,其余文字全部保留:
        #   - 标题里出现「〔数值待核〕」= 提醒读者这处数值待确认;
        #   - 讲稿里出现「〔数值待核〕」= 汇报人不会念出一个错数字。
        # 复用 analyzer._strip_fabricated_numbers(与批注清洗同一套逻辑)。
        for field in ("title", "notes"):
            text = slide.get(field)
            if not isinstance(text, str) or not text:
                continue
            bad = audit_section_numbers(title, text, pool)
            if bad:
                logger.info(
                    "[deck] %s 含无出处数字 %s,已标记为数值待核", field, bad
                )
                slide[field] = _strip_fabricated_numbers(text, set(bad))

        slide.pop("_had_bullets", None)
        audited_slides.append(slide)

    deck["slides"] = audited_slides
    return deck


def _mark_had_bullets(deck: dict[str, Any]) -> dict[str, Any]:
    """在溯源前打标记:记录哪些页原本是有 bullet 的。

    需要区分两种"空":
        - 模型本来就没给 bullet → 保持空(渲染时不画 bullet 区)
        - 模型给了但被溯源删光 → 补占位提示
    """
    for slide in deck.get("slides") or []:
        if isinstance(slide, dict):
            slide["_had_bullets"] = bool(slide.get("bullets"))
    return deck


# ============================================================
# 4) 生成
# ============================================================
def generate_deck(
    report: dict[str, Any],
    indicators: dict[str, Any],
    chart_option: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """调 LLM 生成并归一化演示稿大纲。

    Raises:
        Exception: LLM 调用失败或输出无法解析为有效大纲时抛出,
                    由调用方(ensure_deck)转成 failed 记录。
    """
    catalog = build_chart_catalog(report, chart_option or {})
    user = build_deck_prompt(report, indicators, catalog)

    # 【为什么不用 llm.chat_json】
    #   chat_json 是**为报告链路**设计的:它只在返回的 JSON 里找到 "sections"
    #   字段时才算成功,否则判定"模型跑偏"并进入重试。
    #   演示稿大纲的结构是 {title, subtitle, slides},根本没有 sections,
    #   用 chat_json 会被一律丢弃 → 最终报"无法解析为有效大纲"。
    #   所以这里走 chat() 拿原始文本,再用 _safe_json_loads 做宽容解析
    #   (能处理 ```json 包裹、单引号、裸 key 等 LLM 常见毛病)。
    raw = llm.chat(SYSTEM_PROMPT_DECK, user, json_mode=True)
    try:
        obj = json.loads(raw)
    except Exception:
        from app.core.json_utils import _safe_json_loads
        obj = _safe_json_loads(raw)

    deck = _normalize_deck(obj, report)
    if not deck:
        raise ValueError(
            f"演示稿输出无法解析为有效大纲(raw 头 200 字: {str(raw)[:200]!r})"
        )

    deck = _mark_had_bullets(deck)
    deck = _audit_deck_numbers(deck, report, indicators)
    return deck


# ============================================================
# 5) 缓存读写
# ============================================================
def get_deck(db: Session, task_id: int) -> dict[str, Any] | None:
    """读该任务最新一条 `ready` 大纲;没有则返回 None。"""
    row = db.query(Deck).filter(
        Deck.task_id == task_id,
        Deck.status == "ready",
    ).order_by(Deck.id.desc()).first()
    if not row:
        return None
    try:
        data = json.loads(row.deck_json or "{}")
    except Exception:
        return None
    return data if isinstance(data, dict) and data.get("slides") else None


def get_latest_deck_row(db: Session, task_id: int) -> Deck | None:
    """读该任务最新一条 deck 记录(不论状态),用于判断"是否正在生成"。"""
    return db.query(Deck).filter(
        Deck.task_id == task_id,
    ).order_by(Deck.id.desc()).first()


def ensure_deck(
    db: Session,
    task,
    force: bool = False,
) -> dict[str, Any] | None:
    """确保该任务有可用演示稿:命中缓存直接返回,否则生成并落库。

    **本函数不抛异常**:任何失败都转成 `deck` 表里的 failed 记录并返回 None。
    调用方(报告分析完成后的预生成、导出兜底)因此不会因演示稿而失败。

    Args:
        db:    数据库会话。
        task:  Task 对象。
        force: True 时忽略缓存,强制重新生成。
    """
    try:
        task_id = int(task.id)

        if not force:
            cached = get_deck(db, task_id)
            if cached:
                return cached

        # 标记"生成中":并发请求看到这条记录就不会再起一次 LLM
        row = Deck(task_id=task_id, status="generating", deck_json="")
        db.add(row)
        db.commit()
        db.refresh(row)

        try:
            raw_report = task.report_content or ""
            report = json.loads(raw_report) if raw_report else {}
            indicators = json.loads(task.indicators_json or "{}") or {}
            chart_option = json.loads(task.chart_option_json or "{}") or {}

            deck = generate_deck(report, indicators, chart_option)
        except Exception as exc:
            logger.warning("[deck] 生成失败(task=%s): %s", task_id, exc)
            row.status = "failed"
            row.error_message = f"{type(exc).__name__}: {exc}"[:500]
            row.deck_json = ""
            db.commit()
            return None

        row.status = "ready"
        row.deck_json = dumps(deck)
        row.error_message = ""
        db.commit()
        db.refresh(row)
        return deck
    except Exception as exc:  # pragma: no cover - 兜底:连数据库异常都不外抛
        logger.warning("[deck] ensure_deck 外层异常: %s", exc)
        return None


# ============================================================
# 6) 讲义(Markdown)
# ============================================================
def deck_to_markdown(deck: dict[str, Any], task_name: str = "") -> str:
    """把演示稿渲染成 Markdown 讲义(标题 + 每页要点 + 讲稿)。

    用途:汇报人打印/二次编辑;也可以直接贴进邮件或群消息。
    """
    if not isinstance(deck, dict) or not deck.get("slides"):
        return ""

    lines: list[str] = []
    title = str(deck.get("title") or "演示稿").strip()
    subtitle = str(deck.get("subtitle") or "").strip()
    lines.append(f"# {title}")
    if subtitle:
        lines.append(f"\n> {subtitle}")
    if task_name:
        lines.append(f"\n- 报告任务：{task_name}")

    for idx, slide in enumerate(deck.get("slides") or [], start=1):
        if not isinstance(slide, dict):
            continue
        lines.append("")
        lines.append(f"## 第 {idx} 页 · {slide.get('title', '')}")
        for bullet in slide.get("bullets") or []:
            lines.append(f"- {bullet}")
        notes = str(slide.get("notes") or "").strip()
        if notes:
            lines.append("")
            lines.append(f"**讲稿**：{notes}")

    lines.append("")
    lines.append("---")
    lines.append("（本页为结束页 · 谢谢观看）")
    return "\n".join(lines)
