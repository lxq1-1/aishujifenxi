"""报告导出：PDF / Word / Markdown / Excel。

对齐目标：
    系统页面（ReportView.vue）展示的完整报告 = 标题 + AI 摘要 + sections +
    数据图表 + 关键指标表 + 异常批注。导出文件必须与之一致，否则用户下载下来的
    内容会"少一块"，体验割裂。本文件统一按以下顺序产出：

        1. 标题
        2. AI 摘要（若有）
        3. 各章节正文（清洗后）
        4. 数据图表
              - Word / Excel：matplotlib 真实图（与前端 ECharts 同数据）
              - PDF / Markdown：option 摘要（PDF 难以原生复刻矢量图，MD 适合文本）
        5. 关键指标表（numeric_summary / missing_ratio_top5 / text_value_counts）
        6. 异常批注（annotations_json）

文本清洗策略：
    所有写到导出文件里的中文字段都必须先过 :func:`clean_llm_text`：
        strip_thinking()         —— 剥 <think> 类标签包裹的思考块
        strip_english_reasoning()—— 剥"明文写在正文里"的英文推导段
    后者是新加的：reasoning 模型（DeepSeek-R1 / QwQ / MiniMax-M2.7 等）
    会无视 SYSTEM_PROMPT 的"禁止英文推导"约束，把思考过程写在 sections
    正文中，例如：
        "Let me carefully analyze the data provided. Looking at the indicator data:
         - 36 rows total - 25 columns ..."
    这条规则由 :func:`json_utils.strip_english_reasoning` 实现。

图表真实化策略：
    Word / Excel 通过 :mod:`chart_renderer` 把 ECharts option 渲染为 PNG，
    嵌入文档。PDF / Markdown 仍用 option 摘要：PDF 静态文本足够阅读，
    Markdown 看图不便但能保留完整数据结构便于二次使用。
"""
from __future__ import annotations

import io
import json
import re
from typing import Any

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# python-pptx 的 RGBColor（PPT 主题字典与 export_pptx 都依赖，提前到模块顶层 import，
# 避免函数内局部 import 触发循环依赖/启动开销）。
from pptx.dml.color import RGBColor

from app.core.json_utils import clean_llm_text, clean_llm_obj, json_default
from app.services.chart_renderer import render_chart_to_png

# 注册中文字体（ReportLab 自带 CID 字体，无需外部 ttf）
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))


# ============================================================
# PPT 主题字典：4 套主题供前端弹窗选择
# 字段:
#   title_bg   标题页/尾页整页背景色
#   accent     顶部色条 + 表格头底色
#   text       主正文颜色
#   muted      副标题/降级文本颜色
#   title_text 标题页主标题文字颜色（可能在白底上要压暗）
#   background 兜底背景（目前未使用，可与 title_bg 同）
# ============================================================
_PPT_THEMES = {
    "deep_cyan": {
        "title_bg":   RGBColor(0x0e, 0x74, 0x90),
        "accent":     RGBColor(0x22, 0xd3, 0xee),
        "text":       RGBColor(0x1f, 0x29, 0x37),
        "muted":      RGBColor(0x6b, 0x72, 0x80),
        "title_text": RGBColor(0xff, 0xff, 0xff),
        "background": RGBColor(0xff, 0xff, 0xff),
    },
    "midnight": {
        "title_bg":   RGBColor(0x0f, 0x17, 0x2a),
        "accent":     RGBColor(0x7d, 0xd3, 0xfc),
        "text":       RGBColor(0xe2, 0xe8, 0xf0),
        "muted":      RGBColor(0x94, 0xa3, 0xb8),
        "title_text": RGBColor(0xff, 0xff, 0xff),
        "background": RGBColor(0x0f, 0x17, 0x2a),
    },
    "light": {
        "title_bg":   RGBColor(0xff, 0xff, 0xff),
        "accent":     RGBColor(0x0e, 0x74, 0x90),
        "text":       RGBColor(0x1f, 0x29, 0x37),
        "muted":      RGBColor(0x6b, 0x72, 0x80),
        "title_text": RGBColor(0x0e, 0x74, 0x90),
        "background": RGBColor(0xff, 0xff, 0xff),
    },
    "business": {
        "title_bg":   RGBColor(0x1e, 0x3a, 0x8a),
        "accent":     RGBColor(0xfb, 0xbf, 0x24),
        "text":       RGBColor(0x11, 0x18, 0x27),
        "muted":      RGBColor(0x4b, 0x55, 0x63),
        "title_text": RGBColor(0xff, 0xff, 0xff),
        "background": RGBColor(0xf5, 0xf1, 0xe6),
    },
}


def _theme_palette(theme: str) -> dict:
    """根据 theme key 取主题配色 dict；非法 key 抛 ValueError（由 API 层转 400）。"""
    if theme not in _PPT_THEMES:
        raise ValueError(
            f"不支持的主题: {theme};可选 "
            "deep_cyan/midnight/light/business"
        )
    return _PPT_THEMES[theme]


# ============================================================
# 公共：task → 统一字典（导出所有格式都从这里取数据）
# ============================================================
def _read_task(task) -> dict:
    """把 ORM task 转 dict,并把 report_content 解析为 Python 对象。"""
    raw_report = task.report_content or ""
    try:
        report = json.loads(raw_report) if raw_report else {}
    except Exception:
        report = {"raw": raw_report}
    return {
        "id": task.id,
        "name": task.name,
        "status": task.status,
        "report": report or {},          # title / sections / summary / chart_option
        "chart_option": _safe_json_load(task.chart_option_json, default={}),
        "indicators": _safe_json_load(task.indicators_json, default={}),
        "annotations": _safe_json_load(task.annotations_json, default=[]),
    }


def _safe_json_load(s, default):
    """小心地 json.loads,失败时返回 default。"""
    if not s:
        return default
    try:
        v = json.loads(s)
        return v if v is not None else default
    except Exception:
        return default


def _detail_tables(data: dict) -> dict:
    """取章节明细数据表(2026-10 新增)。

    数据落在 task.chart_option_json 的 detail_tables 字段,
    与前端 ReportView 读的是同一份,保证"页面所见 = 导出所得"。
    老报告没有这个字段,返回空 dict(导出行为与旧实现完全一致)。
    """
    co = (data or {}).get("chart_option") or {}
    if not isinstance(co, dict):
        return {}
    dt = co.get("detail_tables")
    return dt if isinstance(dt, dict) else {}


def _lookup_detail_table(dt_all: dict, section: str) -> dict | None:
    """按章节名取明细表,带一次宽松兜底匹配。

    为什么需要兜底:章节 key 可能一侧带编号前缀("一、总体经营概况")、
    另一侧不带("总体经营概况"),精确匹配会漏掉表格。
    复用 ai_service._loosen_section_key 的宽松化能力(局部导入避免循环依赖)。
    """
    if not dt_all:
        return None
    exact = dt_all.get(section)
    if exact:
        return exact
    try:
        from app.services.ai_service import _loosen_section_key
    except Exception:
        return None
    target = set(_loosen_section_key(str(section)))
    for cand_key, cand_tbl in dt_all.items():
        if set(_loosen_section_key(str(cand_key))) & target:
            return cand_tbl
    return None


def _fmt_cell(v) -> str:
    """表格单元格转文本:None → "—",数字去掉无意义小数尾巴。"""
    if v is None:
        return "—"
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, (int, float)):
        f = float(v)
        if f == int(f) and abs(f) < 1e15:
            return str(int(f))
        return str(round(f, 4)).rstrip("0").rstrip(".")
    return str(v)


def _detail_table_to_markdown(table: dict) -> str:
    """把一张明细表渲染成 markdown 表格文本(供 Markdown/Excel 导出)。"""
    if not isinstance(table, dict):
        return ""
    cols = table.get("columns") or []
    rows = table.get("rows") or []
    if not cols or not rows:
        return ""
    lines: list[str] = []
    title = table.get("title")
    if title:
        lines.append(f"**{title}**")
        lines.append("")
    lines.append("| " + " | ".join(str(c) for c in cols) + " |")
    lines.append("| " + " | ".join("---" for _ in cols) + " |")
    for r in rows:
        if not isinstance(r, (list, tuple)):
            continue
        # 列数不齐时补齐/截断,避免 markdown 表格串列
        cells = [_fmt_cell(c) for c in list(r)[: len(cols)]]
        cells += ["—"] * (len(cols) - len(cells))
        lines.append("| " + " | ".join(cells) + " |")
    note = table.get("note")
    if note:
        lines.append("")
        lines.append(f"> {note}")
    return "\n".join(lines)


# ============================================================
# 公共：报告字段提取（统一清洗逻辑）
# ============================================================
def _title(report: dict, fallback: str = "未命名报告") -> str:
    raw = (report or {}).get("title") or fallback
    return clean_llm_text(raw) if isinstance(raw, str) else raw or fallback


def _summary(report: dict) -> str:
    raw = (report or {}).get("summary") or ""
    if not isinstance(raw, str):
        return ""
    return clean_llm_text(raw)


def _sections(report: dict) -> list[tuple[str, str]]:
    """取章节列表 [(章节名, 清洗后的正文), ...]。

    【2026-10 新增】正文统一在此处做分段处理:
        用户要求"章节要分段,不要全部都在一段上"。模型有时会写成
        500 字不带换行的整块文字,导出 PDF/Word 后就是一整块密排,
        读者会直接跳过。这里复用 analyzer.split_paragraphs 的同一套
        口径(尊重已有空行 → 超长段按句号切 → 按语义标记隐式拆段),
        把段落用 "\n\n" 连接返回。

        为什么放在 `_sections` 而不是在各导出器里:
        markdown / docx / pdf / excel 四个导出器都调用它,
        在此处分段一次即可,四处行为一致、不会漂移。
    """
    from app.services.analyzer import split_paragraphs

    out: list[tuple[str, str]] = []
    sections = (report or {}).get("sections") or {}
    for k, v in sections.items():
        title = clean_llm_text(str(k)) if k else ""
        # 【row-diff-detail-expansion 新增】"对比详情" 在 row_diff 模式下是结构化 dict,
        # 渲染成 markdown 表格(导出 PDF/Word/Excel 时仍能展示新增/删除明细),
        # 避免 clean_llm_text 收到 dict 报错导致整份报告导出失败。
        if isinstance(v, dict) and (
            isinstance(v.get("added"), list)
            and isinstance(v.get("removed"), list)
        ):
            body = _render_diff_detail_to_markdown(v)
        else:
            body = clean_llm_text(v) if isinstance(v, str) else (str(v) if v is not None else "")
            # 纯文本章节才需要分段(struct dict 已是表格/列表结构)
            if body:
                paras = split_paragraphs(body)
                if paras:
                    body = "\n\n".join(paras)
        if not title:
            continue
        out.append((title, body))
    return out


def _render_diff_detail_to_markdown(payload: dict) -> str:
    """把"对比详情·结构化 dict"渲染成 markdown 文本(供 PDF/Word/Excel 导出)。

    与前端 ReportView 的 diff-detail 渲染一一对应,但用 markdown 表达:
        - 顶部写汇总(summary)
        - 新增 / 删除 块用 markdown 表格
        - 修改 块按"主键: 字段A 前 → 后" 列表
    """
    lines: list[str] = []
    summary = payload.get("summary", "")
    if summary:
        lines.append(str(summary))
        lines.append("")

    def _rows_to_md(rows: list, label: str) -> None:
        """把 list[dict] 渲染成 markdown 表格(列取并集,截断前 50 行)。"""
        if not rows:
            return
        cols: list[str] = []
        seen: set[str] = set()
        for r in rows:
            if not isinstance(r, dict):
                continue
            for ck in r.keys():
                if ck not in seen:
                    seen.add(ck)
                    cols.append(str(ck))
        if not cols:
            return
        lines.append(f"**{label}** ({len(rows)} 条):")
        lines.append("")
        lines.append("| " + " | ".join(cols) + " |")
        lines.append("| " + " | ".join("---" for _ in cols) + " |")
        for r in rows[:50]:
            if not isinstance(r, dict):
                continue
            cells: list[str] = []
            for c in cols:
                v = r.get(c, "")
                if v is None:
                    cells.append("—")
                else:
                    cells.append(
                        str(v).replace("\n", " ").replace("|", "\\|")
                    )
            lines.append("| " + " | ".join(cells) + " |")
        if len(rows) > 50:
            lines.append("")
            lines.append(f"_仅展示前 50 条,还有 {len(rows) - 50} 条未展开_")
        lines.append("")

    _rows_to_md(payload.get("added") or [], "新增明细")
    _rows_to_md(payload.get("removed") or [], "删除明细")

    modified = payload.get("modified") or []
    if modified:
        lines.append(f"**修改明细** ({payload.get('modified_total', len(modified))} 条):")
        lines.append("")
        for m in modified[:50]:
            if not isinstance(m, dict):
                continue
            lines.append(f"  - 主键:{m.get('key', '—')}")
            for fd in (m.get("field_diffs") or []):
                if not isinstance(fd, dict):
                    continue
                field = fd.get("field", "")
                before = fd.get("before", "—")
                after = fd.get("after", "—")
                lines.append(
                    f"    - {field}:{before} → {after}"
                )
        if payload.get("truncated_modified"):
            lines.append("")
            shown = payload.get("shown_modified", len(modified))
            lines.append(f"_仅展示前 {shown} 条修改_")
        lines.append("")

    if not lines:
        return "无对比明细"
    return "\n".join(lines).rstrip()


def _chart_list(report: dict, chart_option: dict) -> list[dict]:
    """归一化为图表列表。

    兼容多种 shape:
        - 旧 shape: ``{"charts": [chart1, chart2, ...]}``
        - 旧 shape: 直接给一张完整 ECharts option(顶层带 ``series``)
        - 新 shape: ``{"by_section": {<章节名>: [chart, ...]}}`` —— 当前 ai_service
          归一化后的主流形态,Word/PDF/Markdown 导出会经过这里。

    优先级:``report["chart_option"]`` > 入参 ``chart_option`` —— 避免
    调方传错时机(例如图表编辑 API 还没写回 DB)造成下载到旧图。
    """
    opt = report.get("chart_option") if isinstance(report.get("chart_option"), dict) else None
    if opt is None:
        opt = chart_option if isinstance(chart_option, dict) else {}
    if not isinstance(opt, dict):
        return []

    # 新 shape:by_section(ai_service._normalize_chart_option 当前正在用的形态)。
    # 把所有章节的图打平成一个列表,导出层用不上"按章节挂载",
    # 只关心"这张报告一共有哪些图、要分别嵌到数据图表区"。
    if isinstance(opt.get("by_section"), dict):
        out: list[dict] = []
        for arr in opt["by_section"].values():
            if isinstance(arr, list):
                out.extend(c for c in arr if isinstance(c, dict))
        return out

    # 旧 shape 1:{charts: [chart, ...]}
    if isinstance(opt.get("charts"), list):
        return [c for c in opt["charts"] if isinstance(c, dict)]

    # 旧 shape 2:整张 ECharts option 漏到顶层
    if opt.get("series"):
        return [opt]

    return []


def _collect_section_charts(report: dict, chart_option: dict) -> tuple[dict, list]:
    """把 chart_option 归一化为 (by_section_map, orphan_list)。

    - 新 shape {by_section: {<章节名>: [charts...]}}:
        * 若 <章节名> 能在 report.sections 里完全匹配 → 进 by_section_map
        * 若 key 不匹配 report.sections 任何 key → 视为"孤儿",汇总到 orphan_list
    - 旧 shape {charts: [...]} 或 顶层 {series: [...]} → 全部进 orphan_list
    - 其它（空 / 错形态）→ 返回 ({}, [])
    """
    if not isinstance(chart_option, dict):
        chart_option = {}
    section_keys = set()
    secs = (report or {}).get("sections") or {}
    if isinstance(secs, dict):
        section_keys = set(secs.keys())

    by_section_map: dict = {}
    orphan_list: list = []

    by_sec = chart_option.get("by_section")
    if isinstance(by_sec, dict) and by_sec:
        for k, arr in by_sec.items():
            if not isinstance(arr, list):
                continue
            charts = [c for c in arr if isinstance(c, dict)]
            if not charts:
                continue
            if k in section_keys:
                by_section_map[k] = charts
            else:
                orphan_list.extend(charts)
        if by_section_map or orphan_list:
            return by_section_map, orphan_list

    # 旧 shape
    if isinstance(chart_option.get("charts"), list):
        orphan_list.extend(c for c in chart_option["charts"] if isinstance(c, dict))
    elif chart_option.get("series"):
        orphan_list.append(chart_option)
    return by_section_map, orphan_list


def _numeric_summary(indicators: dict) -> dict:
    return (indicators or {}).get("numeric_summary") or {}


def _fallback_indicators(indicators: dict) -> tuple[dict, dict]:
    """返回 (missing_ratio, text_value_counts)。"""
    miss = (indicators or {}).get("missing_ratio_top5") or {}
    tv = (indicators or {}).get("text_value_counts") or {}
    return miss, tv


def _annotations(annotations) -> list[dict]:
    out: list[dict] = []
    if not isinstance(annotations, list):
        return out
    for a in annotations:
        if not isinstance(a, dict):
            continue
        c = a.get("comment") or ""
        out.append({
            "metric": clean_llm_text(str(a.get("metric") or "")),
            "value": a.get("value"),
            "level": a.get("level") or "info",
            "comment": clean_llm_text(c) if isinstance(c, str) else str(c),
        })
    return out


# ============================================================
# 图表渲染 helper（Word / Excel 用）：渲染失败时退化为 JSON 占位
# ============================================================
def _render_chart_png_or_none(chart: dict, theme: str = "dark",
                             size: tuple[int, int] = (8, 5)) -> bytes | None:
    """调用 chart_renderer 渲染单张图,失败返回 None。

    Args:
        chart: 图表配置。
        theme: ``dark``(深底浅字,网页/Word 用)或 ``light``(白底深字,
               放进 PPT 页面用)。PPT 页面是白底,深色图块压在上面既不协调
               也会盖住正文,必须用 light。
        size: 渲染尺寸(宽,高)英寸。默认 8×5;
              演示稿 PPT 的图表区是扁的,会传 8×3.75 让图正好填满下半区。
    """
    try:
        return render_chart_to_png(
            chart, width=size[0], height=size[1], dpi=150, theme=theme,
        )
    except Exception:
        return None


def _chart_label(chart: dict, index: int) -> str:
    """取图表的标题文字（清洗 + 兜底）。"""
    raw_title = chart.get("title")
    if isinstance(raw_title, dict):
        t = raw_title.get("text") or f"图表 {index}"
    elif isinstance(raw_title, str) and raw_title:
        t = raw_title
    else:
        t = f"图表 {index}"
    return clean_llm_text(str(t)).strip() or f"图表 {index}"


def _chart_option_json(chart: dict) -> str:
    """导出 option JSON 文本(用于"渲染失败退化"分支 / Markdown。)。"""
    opt = chart.get("option") if isinstance(chart.get("option"), dict) else chart
    return json.dumps(opt, ensure_ascii=False, default=json_default, indent=2)


def _sanitize_excel_sheet_title(title: str) -> str:
    """Excel sheet 名合规化:截断到 31 字符,过滤 ``[ ] / \\ ? * :`` 等禁用字符。

    openpyxl 对包含以上字符的标题会直接抛 ``Invalid character found in sheet title``。
    中文标点如 `:`、`(`、`)`、`?`、`*` 在中文报告的章节标题中很常见(例如"附录:本周关键数据明细"),
    必须清洗。
    """
    if not isinstance(title, str):
        title = str(title)
    bad_chars = set("[]/\\?*:")
    cleaned = "".join("_" if ch in bad_chars else ch for ch in title)
    cleaned = cleaned.strip("'") or "Sheet"
    return cleaned[:31]


# ============================================================
# Markdown（图表用 option 摘要：MD 文本格式方便二次使用）
# ============================================================
def export_markdown(task) -> bytes:
    """导出 Markdown：摘要 + 章节正文(章节下方 inline 该章节图表)+ 孤儿图兜底 + 关键指标 + 异常批注。

    说明:前端 ReportView 把每张图挂在对应章节的正文下方,导出 MD 也保持一致:
    - 章节顺序由 ``report.sections`` 决定,每章节后紧随 ``by_section[章节名]`` 的图;
    - 未挂章节的图与旧 shape ``{charts:[...]}`` 汇总到末尾的"数据图表"section。
    """
    data = _read_task(task)
    report = data["report"]

    lines: list[str] = [f"# {_title(report, data['name'])}\n"]

    summary = _summary(report)
    if summary:
        lines.append(f"> **AI 摘要**：{summary}\n")

    # 归一化一次,后续 inline + 兜底复用
    by_section_charts, orphan_charts = _collect_section_charts(
        report, data["chart_option"]
    )
    # 【2026-10 新增】章节明细数据表,与前端 ReportView 同源
    dt_all = _detail_tables(data)

    for sec, content in _sections(report):
        lines.append(f"## {sec}\n\n{content}\n")
        sec_charts = by_section_charts.get(sec) or []
        for ci, c in enumerate(sec_charts, 1):
            t = _chart_label(c, ci)
            lines.append(f"### {t}\n")
            lines.append("```json")
            lines.append(_chart_option_json(c))
            lines.append("```\n")
        # 【2026-10 新增】章节明细数据表:与前端同源,保证"页面所见 = 导出所得"
        _dt = _lookup_detail_table(dt_all, sec)
        _md = _detail_table_to_markdown(_dt) if _dt else ""
        if _md:
            lines.append(_md)
            lines.append("")

    # 孤儿图表兜底:旧报告 {charts:[...]} 或 by_section 有不匹配章节名
    if orphan_charts:
        lines.append("\n## 数据图表\n")
        for i, c in enumerate(orphan_charts, 1):
            t = _chart_label(c, i)
            lines.append(f"### {t}\n")
            lines.append("```json")
            lines.append(_chart_option_json(c))
            lines.append("```\n")

    num = _numeric_summary(data["indicators"])
    if num:
        lines.append("\n## 关键指标\n")
        lines.append("| 指标 | 合计 | 均值 | 最大 | 最小 |")
        lines.append("|---|---|---|---|---|")
        for k, v in num.items():
            lines.append(
                f"| {k} | {v.get('sum', '')} | {v.get('mean', '')} "
                f"| {v.get('max', '')} | {v.get('min', '')} |"
            )

    miss, tv = _fallback_indicators(data["indicators"])
    if miss or tv:
        if not num:
            lines.append("\n## 关键指标\n")
        if miss:
            lines.append("\n### 缺失率 Top5\n")
            lines.append("| 字段 | 缺失率 |")
            lines.append("|---|---|")
            for k, p in miss.items():
                lines.append(f"| {k} | {p}% |")
        if tv:
            for col, items in tv.items():
                title = clean_llm_text(str(col))
                lines.append(f"\n### {title} · 频次 Top8\n")
                lines.append("| 取值 | 出现次数 |")
                lines.append("|---|---|")
                for it in items[:8]:
                    name = clean_llm_text(str(it.get("name", "")))
                    lines.append(f"| {name} | {it.get('value', '')} |")

    annos = _annotations(data["annotations"])
    if annos:
        lines.append("\n## 异常批注\n")
        for a in annos:
            lines.append(f"- **[{a['level']}] {a['metric']}**：{a['comment']}")

    return ("\n".join(lines)).encode("utf-8")


# ============================================================
# Word（python-docx）：图表用 matplotlib PNG 真实嵌入
# ============================================================
def export_docx(task) -> bytes:
    """导出 Word：标题 + 摘要 + 章节 + 真实图表 PNG + 关键指标 + 异常批注。"""
    from docx import Document
    from docx.shared import Inches, Pt

    doc = Document()
    data = _read_task(task)
    report = data["report"]

    # 1. 主标题
    doc.add_heading(_title(report, data["name"]), level=1)

    # 2. AI 摘要
    summary = _summary(report)
    if summary:
        p = doc.add_paragraph()
        run = p.add_run("AI 摘要：")
        run.bold = True
        p.add_run(summary)

    # 3. 章节正文 + 该章节内嵌的图表（与前端 ReportView 一致,inline 在章节下方）
    by_section_charts, orphan_charts = _collect_section_charts(
        report, data["chart_option"]
    )
    # 【2026-10 新增】章节明细数据表,与前端同源
    dt_all = _detail_tables(data)
    for sec, content in _sections(report):
        doc.add_heading(sec, level=2)
        # 【2026-10 新增】按段落写入 Word:每个段落一个 add_paragraph,
        # Word 会在段间自动留白。若把带 \n 的整块丢进单个 add_paragraph,
        # 导出后仍是一整块密排文字(Word 不把 \n 当段落分隔)。
        for para in (content or "").split("\n\n"):
            if para.strip():
                doc.add_paragraph(para.strip())
        # ---- 章节明细数据表(2026-10 新增):渲染成真正的 Word 表格 ----
        _dt = _lookup_detail_table(dt_all, sec)
        if _dt:
            _cols = _dt.get("columns") or []
            _rows = _dt.get("rows") or []
            if _cols and _rows:
                _title_txt = _dt.get("title")
                if _title_txt:
                    _tp = doc.add_paragraph()
                    _tr = _tp.add_run(str(_title_txt))
                    _tr.bold = True
                _table = doc.add_table(rows=1, cols=len(_cols))
                _table.style = "Light Grid Accent 1"
                for _ci, _cname in enumerate(_cols):
                    _table.rows[0].cells[_ci].text = str(_cname)
                for _r in _rows:
                    if not isinstance(_r, (list, tuple)):
                        continue
                    _cells = _table.add_row().cells
                    for _ci in range(len(_cols)):
                        _val = _r[_ci] if _ci < len(_r) else None
                        _cells[_ci].text = _fmt_cell(_val)
                _note = _dt.get("note")
                if _note:
                    doc.add_paragraph(str(_note))
        sec_charts = by_section_charts.get(sec) or []
        for ci, c in enumerate(sec_charts, 1):
            t = _chart_label(c, ci)
            doc.add_heading(t, level=3)
            png = _render_chart_png_or_none(c)
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(6))
                last_para = doc.paragraphs[-1]
                last_para.alignment = 1  # 居中
            else:
                doc.add_paragraph("(图表渲染失败,以下为 ECharts 配置摘要)")
                doc.add_paragraph(_chart_option_json(c))

    # 4. 孤儿图表兜底(旧 shape / 未挂章节的图统一汇总到末尾的「数据图表」)
    if orphan_charts:
        doc.add_heading("数据图表", level=2)
        for i, c in enumerate(orphan_charts, 1):
            t = _chart_label(c, i)
            doc.add_heading(t, level=3)
            png = _render_chart_png_or_none(c)
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(6))
                last_para = doc.paragraphs[-1]
                last_para.alignment = 1
            else:
                doc.add_paragraph(_chart_option_json(c))

    # 5. 关键指标表
    num = _numeric_summary(data["indicators"])
    if num:
        doc.add_heading("关键指标", level=2)
        table = doc.add_table(rows=1, cols=5)
        table.style = "Light Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "指标"; hdr[1].text = "合计"; hdr[2].text = "均值"
        hdr[3].text = "最大"; hdr[4].text = "最小"
        for k, v in num.items():
            row = table.add_row().cells
            row[0].text = str(k)
            row[1].text = str(v.get("sum", ""))
            row[2].text = str(v.get("mean", ""))
            row[3].text = str(v.get("max", ""))
            row[4].text = str(v.get("min", ""))

    # 兜底指标
    miss, tv = _fallback_indicators(data["indicators"])
    if miss:
        doc.add_heading("缺失率 Top5", level=3)
        table = doc.add_table(rows=1, cols=2)
        table.style = "Light Grid"
        table.rows[0].cells[0].text = "字段"
        table.rows[0].cells[1].text = "缺失率"
        for k, p in miss.items():
            row = table.add_row().cells
            row[0].text = str(k)
            row[1].text = f"{p}%"
    if tv:
        for col, items in tv.items():
            title = clean_llm_text(str(col))
            doc.add_heading(f"{title} · 频次 Top8", level=3)
            table = doc.add_table(rows=1, cols=2)
            table.style = "Light Grid"
            table.rows[0].cells[0].text = "取值"
            table.rows[0].cells[1].text = "出现次数"
            for it in items[:8]:
                row = table.add_row().cells
                row[0].text = clean_llm_text(str(it.get("name", "")))
                row[1].text = str(it.get("value", ""))

    # 6. 异常批注
    annos = _annotations(data["annotations"])
    if annos:
        doc.add_heading("异常批注", level=2)
        for a in annos:
            doc.add_paragraph(f"[{a['level']}] {a['metric']}: {a['comment']}")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ============================================================
# Excel（openpyxl）：图表用 matplotlib PNG 嵌入 cell
# ============================================================
def export_excel(task) -> bytes:
    """导出 Excel：多 sheet + 真实图表 PNG。

    Sheet 列表：
        - 报告      （章节 + 内容）
        - 摘要      （AI 摘要,如有）
        - 关键指标  （numeric_summary）
        - 数据图表  （PNG 一行一张）
        - 字段明细  （missing_ratio_top5 / text_value_counts）
        - 异常批注
    """
    wb = Workbook()
    data = _read_task(task)
    report = data["report"]

    # 归一化一次,供后续按章节分 sheet 与孤儿兜底 sheet 复用
    by_section_charts, orphan_charts = _collect_section_charts(
        report, data["chart_option"]
    )

    # 报告正文（首 sheet）
    ws = wb.active
    ws.title = "报告"
    ws.append(["章节", "内容"])
    for sec, content in _sections(report):
        ws.append([sec, content or ""])

    # 摘要 sheet
    summary = _summary(report)
    if summary:
        ws_sum = wb.create_sheet("摘要")
        ws_sum.append(["标题", _title(report, data["name"])])
        ws_sum.append(["AI 摘要", summary])

    # 关键指标 sheet
    num = _numeric_summary(data["indicators"])
    if num:
        ws2 = wb.create_sheet("关键指标")
        ws2.append(["指标", "合计", "均值", "最大", "最小"])
        for k, v in num.items():
            ws2.append([k, v.get("sum"), v.get("mean"), v.get("max"), v.get("min")])

    # 兜底指标 sheet
    miss, tv = _fallback_indicators(data["indicators"])
    if miss or tv:
        ws_fb = wb.create_sheet("字段明细")
        if miss:
            ws_fb.append(["字段", "缺失率(%)"])
            for k, p in miss.items():
                ws_fb.append([k, p])
        if tv:
            ws_fb.append([])
            for col, items in tv.items():
                ws_fb.append([f"{col} 频次 Top8", "出现次数"])
                for it in items[:8]:
                    ws_fb.append([it.get("name", ""), it.get("value", "")])

    # 图表按章节分 sheet(每个有数据的章节一张)。Excel 没有"页面流"做不了 inline,
    # 但可以用"图 - 章节名"这种 sheet 名明确表达归属,与前端 ReportView 语义一致。
    for sec_name, sec_charts in by_section_charts.items():
        if not sec_charts:
            continue
        sheet_title = _sanitize_excel_sheet_title(f"图 - {sec_name}")
        ws_chart = wb.create_sheet(sheet_title)
        for i, c in enumerate(sec_charts, 1):
            t = _chart_label(c, i)
            png = _render_chart_png_or_none(c)
            # 标题占 A 列;图片锚定在 B 列同行;14 行一张图,行高撑大装图片
            ws_chart.cell(row=i * 14 - 13, column=1, value=f"{sec_name} · {t}")
            if png:
                img = XLImage(io.BytesIO(png))
                img.width = 720
                img.height = 480
                ws_chart.add_image(img, f"B{i * 14 - 13}")
            else:
                ws_chart.cell(row=i * 14 - 12, column=1, value=_chart_option_json(c))
            for r in range(i * 14 - 13, i * 14 + 1):
                ws_chart.row_dimensions[r].height = 18

    # 孤儿图表汇总到「数据图表」sheet(旧 shape / 未挂章节的图)
    if orphan_charts:
        ws_c = wb.create_sheet("数据图表")
        for i, c in enumerate(orphan_charts, 1):
            t = _chart_label(c, i)
            png = _render_chart_png_or_none(c)
            ws_c.cell(row=i * 14 - 13, column=1, value=t)
            if png:
                img = XLImage(io.BytesIO(png))
                img.width = 720
                img.height = 480
                ws_c.add_image(img, f"B{i * 14 - 13}")
            else:
                ws_c.cell(row=i * 14 - 12, column=1, value=_chart_option_json(c))
            for r in range(i * 14 - 13, i * 14 + 1):
                ws_c.row_dimensions[r].height = 18

    # 异常批注 sheet
    annos = _annotations(data["annotations"])
    if annos:
        ws_a = wb.create_sheet("异常批注")
        ws_a.append(["级别", "指标", "值", "说明"])
        for a in annos:
            ws_a.append([a["level"], a["metric"], a["value"], a["comment"]])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# ============================================================
# PDF（ReportLab）：图表用 option 摘要（PDF 难嵌入矢量交互图,文本摘要可读性可）
# ============================================================
def export_pdf(task) -> bytes:
    """导出 PDF：与前端显示保持一致结构,中文 STSong-Light 字体。"""
    data = _read_task(task)
    report = data["report"]

    bio = io.BytesIO()
    doc = SimpleDocTemplate(
        bio, pagesize=A4,
        leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40,
    )
    styles = getSampleStyleSheet()
    zh_title = ParagraphStyle("zh_title", parent=styles["Title"],
                              fontName="STSong-Light", fontSize=20, leading=26)
    zh_h2 = ParagraphStyle("zh_h2", parent=styles["Heading2"],
                            fontName="STSong-Light", fontSize=14, leading=20)
    zh_h3 = ParagraphStyle("zh_h3", parent=styles["Heading3"],
                            fontName="STSong-Light", fontSize=12, leading=18)
    zh_body = ParagraphStyle("zh_body", parent=styles["BodyText"],
                             fontName="STSong-Light", fontSize=11, leading=18)
    zh_summary = ParagraphStyle(
        "zh_summary", parent=zh_body,
        backColor=colors.HexColor("#0e7490"), textColor=colors.white,
        borderPadding=8, leading=18,
    )
    zh_note = ParagraphStyle("zh_note", parent=zh_body,
                            textColor=colors.grey, fontSize=10, leading=15)

    story: list = [
        Paragraph(_title(report, data["name"]), zh_title),
        Spacer(1, 12),
    ]

    # 2. AI 摘要
    summary = _summary(report)
    if summary:
        story.append(Paragraph("AI 摘要", zh_h2))
        story.append(Paragraph(_escape(summary), zh_summary))
        story.append(Spacer(1, 8))

    # 3. 各章节 + 该章节内嵌的图表（与前端 ReportView 一致）
    by_section_charts, orphan_charts = _collect_section_charts(
        report, data["chart_option"]
    )
    for sec, content in _sections(report):
        story.append(Paragraph(_escape(sec), zh_h2))
        # 【2026-10 新增】按段落生成 Paragraph,段间才有留白。
        # 旧实现把带 \n 的整块丢进一个 Paragraph 再 replace("<br/>"),
        # PDF 里行与行之间没有任何间距,导出后仍是一整块密排文字。
        if content:
            for para in content.split("\n\n"):
                para = para.strip()
                if para:
                    story.append(
                        Paragraph(_escape(para).replace("\n", "<br/>"), zh_body)
                    )
        sec_charts = by_section_charts.get(sec) or []
        for ci, c in enumerate(sec_charts, 1):
            t = _chart_label(c, ci)
            story.append(Paragraph(_escape(f"{sec} · {t}"), zh_h3))
            png = _render_chart_png_or_none(c)
            if png:
                from reportlab.platypus import Image as RLImage
                img = RLImage(io.BytesIO(png), width=400, height=260)
                story.append(img)
            else:
                opt_str = _chart_option_json(c)
                if len(opt_str) > 1500:
                    opt_str = opt_str[:1500] + "\n... (option 过长,已截断)"
                story.append(Paragraph(
                    "（PDF 中图表渲染失败,以下为 ECharts 配置摘要。）",
                    zh_note,
                ))
                story.append(Paragraph(
                    _escape(opt_str).replace("\n", "<br/>"),
                    zh_note,
                ))
            story.append(Spacer(1, 6))
        story.append(Spacer(1, 8))

    # 4. 孤儿图表兜底（未挂章节 / 旧 shape 全部汇总到末尾的「数据图表」）
    if orphan_charts:
        story.append(Paragraph("数据图表", zh_h2))
        for i, c in enumerate(orphan_charts, 1):
            t = _chart_label(c, i)
            story.append(Paragraph(_escape(t), zh_h3))
            png = _render_chart_png_or_none(c)
            if png:
                from reportlab.platypus import Image as RLImage
                img = RLImage(io.BytesIO(png), width=400, height=260)
                story.append(img)
            else:
                opt_str = _chart_option_json(c)
                if len(opt_str) > 1500:
                    opt_str = opt_str[:1500] + "\n... (option 过长,已截断)"
                story.append(Paragraph(
                    "（PDF 中图表渲染失败,以下为 ECharts 配置摘要。）",
                    zh_note,
                ))
                story.append(Paragraph(
                    _escape(opt_str).replace("\n", "<br/>"),
                    zh_note,
                ))
            story.append(Spacer(1, 6))

    # 5. 关键指标
    num = _numeric_summary(data["indicators"])
    if num:
        story.append(Paragraph("关键指标", zh_h2))
        tbl_data = [["指标", "合计", "均值", "最大", "最小"]]
        for k, v in num.items():
            tbl_data.append([
                str(k),
                str(v.get("sum", "")),
                str(v.get("mean", "")),
                str(v.get("max", "")),
                str(v.get("min", "")),
            ])
        t = Table(tbl_data, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0e7490")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))

    # 兜底指标
    miss, tv = _fallback_indicators(data["indicators"])
    if miss or tv:
        if not num:
            story.append(Paragraph("关键指标", zh_h2))
        if miss:
            story.append(Paragraph("缺失率 Top5", zh_h3))
            tbl_data = [["字段", "缺失率"]]
            for k, p in miss.items():
                tbl_data.append([str(k), f"{p}%"])
            t = Table(tbl_data, hAlign="LEFT")
            t.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0e7490")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]))
            story.append(t)
            story.append(Spacer(1, 6))
        if tv:
            for col, items in tv.items():
                title = clean_llm_text(str(col))
                story.append(Paragraph(f"{_escape(title)} · 频次 Top8", zh_h3))
                tbl_data = [["取值", "出现次数"]]
                for it in items[:8]:
                    tbl_data.append([
                        clean_llm_text(str(it.get("name", ""))),
                        str(it.get("value", "")),
                    ])
                t = Table(tbl_data, hAlign="LEFT")
                t.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
                    ("FONTSIZE", (0, 0), (-1, -1), 10),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0e7490")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ]))
                story.append(t)
                story.append(Spacer(1, 6))

    # 6. 异常批注
    annos = _annotations(data["annotations"])
    if annos:
        story.append(Paragraph("异常批注", zh_h2))
        for a in annos:
            tag = f"[{a['level']}] {a['metric']}"
            story.append(Paragraph(
                f"<b>{_escape(tag)}</b>: {_escape(a['comment'])}", zh_body,
            ))
            story.append(Spacer(1, 4))

    doc.build(story)
    return bio.getvalue()


def _escape(s: str) -> str:
    """ReportLab Paragraph 用的是 HTML 子集,转义 & < > 。"""
    if not isinstance(s, str):
        s = str(s)
    return (
        s.replace("&", "&amp;")
         .replace("<", "&lt;")
         .replace(">", "&gt;")
    )


# ============================================================
# PPT (python-pptx):16:9 演示稿,标题+摘要+章节+图表+指标+批注
# 设计目标：
#   - 16:9 标准 12.5"×7.5",会议投屏友好
#   - 章节正文过长时按 800 字/块 自动拆页,不截断
#   - 图表走 PNG 嵌入（与 Word/Excel 一致),失败退化为 ECharts option JSON
#   - 与 PDF/Word 字段顺序保持一致:标题→摘要→章节→图表→关键指标→兜底→批注→致谢
#   - 主题色复用项目品牌色:#0e7490(深青,标题/背景) + #22d3ee(强调)
# ============================================================


def _split_section_text(text: str, max_chars: int = 800) -> list[str]:
    """把章节正文按 max_chars 切为若干块。

    切块策略(简单可预测,便于审稿):
        1. 先按换行 \\n 切为段落
        2. 单段超长时按句号"。"再切
        3. 单句仍超长则硬切到 max_chars
    返回的列表不空;若 text 为空,返回 [""]。
    """
    if not text:
        return [""]
    cleaned = str(text).strip()
    if len(cleaned) <= max_chars:
        return [cleaned]

    # 先按 \n 切
    paragraphs = [p.strip() for p in cleaned.split("\n") if p.strip()]
    if not paragraphs:
        return [cleaned[:max_chars]]

    chunks: list[str] = []
    cur = ""
    for para in paragraphs:
        # 当前块 + 段落 不会超过上限 → 合并
        candidate = (cur + "\n" + para).strip() if cur else para
        if len(candidate) <= max_chars:
            cur = candidate
            continue
        # 超出 → 先刷出当前块
        if cur:
            chunks.append(cur)
            cur = ""
        # 单段超长 → 按 "。" 切句
        if len(para) > max_chars:
            sentences = [s.strip() for s in re.split(r"(?<=[。!?])\s*", para) if s.strip()]
            buf = ""
            for sent in sentences:
                cand = (buf + sent).strip() if buf else sent
                if len(cand) <= max_chars:
                    buf = cand
                else:
                    if buf:
                        chunks.append(buf)
                        buf = ""
                    # 单句仍超长 → 硬切
                    while len(sent) > max_chars:
                        chunks.append(sent[:max_chars])
                        sent = sent[max_chars:]
                    buf = sent
            if buf:
                chunks.append(buf)
        else:
            cur = para
    if cur:
        chunks.append(cur)
    return chunks or [""]


def export_pptx(task, theme: str = "deep_cyan") -> bytes:
    """导出 PPT(16:9):标题页 + 摘要页 + 各章节页 + 图表页 + 关键指标页 + 异常批注页 + 尾页。

    实现要点:
        - 12.5"×7.5" (16:9),默认蓝/白配色,与项目风格一致
        - 章节正文 800 字/页 自动拆页(超长不分页则会溢出)
        - 图表渲染沿用 chart_renderer.render_chart_to_png(与 Word/Excel 一致)
        - 关键指标 / 缺失率 / 频次 / 批注 数据为空时**跳过对应页**,不渲染空页
        - theme 参数控制主题色板(deep_cyan / midnight / light / business)
        - chart_option.by_section 下的图会嵌入对应章节正文下方;未挂章节的图走末尾"数据图表"页
    """
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    data = _read_task(task)
    report = data["report"]

    # 16:9 12.5"×7.5" (PowerPoint widescreen 默认值)
    prs = Presentation()
    prs.slide_width = Inches(12.5)
    prs.slide_height = Inches(7.5)

    # 主题色:由 theme 决定。_theme_palette 会做合法性校验,非法 key 直接抛 ValueError。
    palette = _theme_palette(theme)
    COLOR_TITLE_BG = palette["title_bg"]
    COLOR_ACCENT = palette["accent"]
    # 【2026-10 修订】内容页正文固定用深色。
    #   之前正文用 palette["text"],而 midnight 主题的 text 是 #e2e8f0(浅色)——
    #   内容页底色其实是白底(PPT 默认),浅色字落在白底上等于隐形,
    #   四套主题里有一半不可用。主题色只用于标题页/尾页/色条。
    COLOR_TEXT = RGBColor(0x1f, 0x29, 0x37)
    COLOR_MUTED = RGBColor(0x6b, 0x72, 0x80)
    # 标题页 / 尾页主标题文字色。light 主题下白底要压暗,其它主题用白。
    COLOR_TITLE_TEXT = palette["title_text"]
    # 兼容老代码:某些表格"白字"场景仍保留纯白常量(不在主题里换)。
    COLOR_WHITE = RGBColor(0xff, 0xff, 0xff)

    blank_layout = prs.slide_layouts[6]  # 空白 layout,自己画框线与文字

    def _add_blank_slide():
        return prs.slides.add_slide(blank_layout)

    def _fill_solid(slide, rgb_color):
        """整页填一个底色(用于封面/章节分隔)。"""
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = rgb_color

    def _add_textbox(slide, left, top, width, height, text,
                     font_size=18, bold=False, color=COLOR_TEXT,
                     align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name=None):
        """快捷加一个 textbox,返回 shape 便于追加段落。"""
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = Inches(0.05)
        tf.margin_right = Inches(0.05)
        tf.margin_top = Inches(0.05)
        tf.margin_bottom = Inches(0.05)
        p = tf.paragraphs[0]
        p.alignment = align
        if isinstance(text, str):
            text_list = [text]
        else:
            text_list = text
        for i, t in enumerate(text_list):
            if i == 0:
                run = p.add_run()
            else:
                pp = tf.add_paragraph()
                pp.alignment = align
                run = pp.add_run()
            run.text = t
            f = run.font
            f.size = Pt(font_size)
            f.bold = bold
            f.color.rgb = color
            if font_name:
                f.name = font_name
        return tb

    def _add_horizontal_bar(slide, color, top=None):
        """顶部一条小色条作为视觉锚定。"""
        from pptx.shapes.autoshape import Shape
        from pptx.enum.shapes import MSO_SHAPE
        if top is None:
            top = Inches(0.0)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, top,
                                     prs.slide_width, Inches(0.15))
        bar.line.fill.background()
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        return bar

    # ============================================================
    # 1. 标题页
    # ============================================================
    slide = _add_blank_slide()
    _fill_solid(slide, COLOR_TITLE_BG)
    _add_horizontal_bar(slide, COLOR_ACCENT, top=Inches(3.0))

    title = _title(report, data["name"])
    _add_textbox(
        slide, Inches(1.0), Inches(2.0), Inches(10.5), Inches(1.5),
        title, font_size=44, bold=True, color=COLOR_TITLE_TEXT,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    _add_textbox(
        slide, Inches(1.0), Inches(3.4), Inches(10.5), Inches(0.6),
        f"任务: {data['name']}", font_size=18, color=COLOR_TITLE_TEXT,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    create_time_str = ""
    if getattr(task, "create_time", None):
        try:
            create_time_str = task.create_time.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            create_time_str = str(task.create_time)
    _add_textbox(
        slide, Inches(1.0), Inches(4.0), Inches(10.5), Inches(0.5),
        f"生成时间:{create_time_str}" if create_time_str else "",
        font_size=14, color=COLOR_TITLE_TEXT,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    _add_textbox(
        slide, Inches(1.0), Inches(6.2), Inches(10.5), Inches(0.6),
        "AI 数据处理工作台 · 自动生成",
        font_size=14, color=COLOR_TITLE_TEXT,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )

    # ============================================================
    # 2. 摘要页(有摘要时)
    # ============================================================
    summary_text = _summary(report)
    if summary_text:
        slide = _add_blank_slide()
        _add_horizontal_bar(slide, COLOR_ACCENT)
        _add_textbox(
            slide, Inches(0.6), Inches(0.4), Inches(11.3), Inches(0.8),
            "AI 摘要", font_size=28, bold=True, color=COLOR_TITLE_BG,
        )
        # 摘要框:浅青底
        from pptx.enum.shapes import MSO_SHAPE
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.6), Inches(1.4), Inches(11.3), Inches(5.5),
        )
        box.fill.solid()
        # 用主题里的 background 兜底色（light 主题为白,business 主题为米白等）,
        # 避免在白底主题上与极浅青撞色看不见。
        box.fill.fore_color.rgb = palette["background"]
        box.line.color.rgb = COLOR_ACCENT
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.3)
        tf.margin_right = Inches(0.3)
        tf.margin_top = Inches(0.3)
        tf.margin_bottom = Inches(0.3)
        # 摘要超长时也分块
        chunks = _split_section_text(summary_text, max_chars=600)
        for i, chunk in enumerate(chunks):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = chunk
            run.font.size = Pt(20)
            run.font.color.rgb = COLOR_TEXT

    # ============================================================
    # 3. 章节页(每个 section,正文过长自动分页)
    # ============================================================
    # 把 chart_option 提前归一化为「按章节挂载 + 孤儿」,供步骤 3 与步骤 4 共享
    by_section_charts, orphan_charts = _collect_section_charts(
        report, data["chart_option"]
    )
    sections = _sections(report)
    for sec_name, sec_body in sections:
        chunks = _split_section_text(sec_body, max_chars=800)
        total = len(chunks)
        # 该章节对应的 inline 图表(可能为空)。
        # 修根因:之前章节 inline 图没进 PPT 是因为 chart_option.by_section 没有被
        # 任何路径消费。现在采用"读-图文就近"原则:
        #   - 单张图(或 0 张)的章节:文字上半页 / 图下半页共占同一张幻灯片,
        #     阅读路径自然(文字 → 抬眼看到对应图)。
        #   - 多张图(≥2)的章节:第一张图放在该章节**最后一个正文 chunk 同页**,
        #     后续每张图追加独立图页(避免同页挤压),标题「{章节名} · 图 N · {图名}」。
        # 这样既解决"文字描述没有图表数据支持"(图离文字只差几厘米),
        # 也避免"图表堆叠重叠"的旧问题。
        sec_inline_charts = by_section_charts.get(sec_name) or []

        # 章节正文页:分页生成时,**最后一块正文所在页**放"上半文 / 下半图"
        # 其余前置页只放正文(不需要每页都重复图)
        last_idx = total  # 1-based
        for idx, chunk in enumerate(chunks, start=1):
            slide = _add_blank_slide()
            _add_horizontal_bar(slide, COLOR_TITLE_BG)
            heading = sec_name if total == 1 else f"{sec_name} ({idx}/{total})"
            _add_textbox(
                slide, Inches(0.6), Inches(0.4), Inches(11.3), Inches(0.8),
                heading, font_size=28, bold=True, color=COLOR_TITLE_BG,
            )

            # 判断"这页是否要同页放图":在该章节最后一块正文页放**第一张**图(只剩
            # 1 张图时等价于唯一的图;多张图时也保证至少一张与文字同时呈现)。
            # 同页放图,意味着本张幻灯片是"上半正文 / 下半图"布局,
            # 否则是纯正文占满整页布局。
            show_inline_here = (
                idx == last_idx
                and len(sec_inline_charts) >= 1
            )

            if show_inline_here:
                # 上半页正文: y=1.5" → 4.5",高 3",字号 18pt
                _add_textbox(
                    slide, Inches(0.8), Inches(1.5), Inches(10.9),
                    Inches(3.0), chunk, font_size=18, color=COLOR_TEXT,
                )
                # 下半页图: y=4.7" → 7.0",宽 8" 居中
                first_chart = sec_inline_charts[0]
                # light 主题(白底深字):PPT 页面是白底,深色图块压在上面
                # 既不协调,也会盖住浅色正文。尺寸用扁图比例正好填满下半区。
                png = _render_chart_png_or_none(
                    first_chart, theme="light", size=(8, 3.75),
                )
                if png:
                    # 等比缩放塞进 (4.7" → 7.0") 的图区,绝不允许溢出页面。
                    # 之前只给 width=8",高度随 PNG 比例走,实测底部落到 9.70"
                    # (画布只有 7.5")—— 图直接盖住正文。
                    _place_picture_fit(
                        slide, png, top=4.7, max_w=9.0, max_h=2.3,
                    )
                else:
                    _add_textbox(
                        slide, Inches(0.8), Inches(4.7),
                        Inches(10.9), Inches(0.4),
                        "(图表渲染失败,以下为 ECharts 配置摘要)",
                        font_size=13, color=COLOR_MUTED,
                    )
                    opt_str = _chart_option_json(first_chart)
                    if len(opt_str) > 1200:
                        opt_str = opt_str[:1200] + "\n... (option 过长,已截断)"
                    _add_textbox(
                        slide, Inches(0.8), Inches(5.15),
                        Inches(10.9), Inches(1.7),
                        opt_str, font_size=10, color=COLOR_MUTED,
                    )
            else:
                # 纯正文页:占满 y=1.5" → 6.6"(高度 5.1",字号 20pt)
                _add_textbox(
                    slide, Inches(0.8), Inches(1.5), Inches(10.9),
                    Inches(5.1), chunk, font_size=20, color=COLOR_TEXT,
                )

        # 多张图章节(≥2 张):**第一张**已经放在最后一块正文同页,本循环只追加
        # 第 2 张、第 3 张、...为独立图页(避免重复)。
        # 0 张 / 1 张图章节:这里跳过(0 张无图;1 张已经在同页放完)。
        if len(sec_inline_charts) >= 2:
            for j, c in enumerate(sec_inline_charts[1:], start=2):
                chart_label = _chart_label(c, j)
                slide = _add_blank_slide()
                _add_horizontal_bar(slide, COLOR_TITLE_BG)
                _add_textbox(
                    slide, Inches(0.6), Inches(0.4),
                    Inches(11.3), Inches(0.8),
                    f"{sec_name} · 图 {j} · {chart_label}",
                    font_size=24, bold=True, color=COLOR_TITLE_BG,
                )
                png = _render_chart_png_or_none(c, theme="light", size=(8, 3.75))
                if png:
                    # 整页图:等比缩放塞进 (1.5" → 6.9") 的区域,水平居中。
                    # 之前只给 width=9.5",高度随 PNG 比例走会溢出页面。
                    _place_picture_fit(
                        slide, png, top=1.5, max_w=10.5, max_h=5.4,
                    )
                else:
                    _add_textbox(
                        slide, Inches(0.8), Inches(1.7),
                        Inches(10.9), Inches(0.5),
                        "(图表渲染失败,以下为 ECharts 配置摘要)",
                        font_size=14, color=COLOR_MUTED,
                    )
                    opt_str = _chart_option_json(c)
                    if len(opt_str) > 1800:
                        opt_str = opt_str[:1800] + "\n... (option 过长,已截断)"
                    _add_textbox(
                        slide, Inches(0.8), Inches(2.3),
                        Inches(10.9), Inches(4.8),
                        opt_str, font_size=11, color=COLOR_MUTED,
                    )

    # ============================================================
    # 4. 图表页(每张图独立成页,PNG 嵌入,失败退化)
    # ============================================================
    # by_section_charts / orphan_charts 已在步骤 3 顶部归一化;
    # 这里只渲染剩余的孤儿图(inline 图已经在 3 中嵌到章节页)。
    for i, c in enumerate(orphan_charts, 1):
        slide = _add_blank_slide()
        _add_horizontal_bar(slide, COLOR_TITLE_BG)
        t = _chart_label(c, i)
        _add_textbox(
            slide, Inches(0.6), Inches(0.4), Inches(11.3), Inches(0.8),
            f"数据图表 · {t}", font_size=24, bold=True, color=COLOR_TITLE_BG,
        )
        png = _render_chart_png_or_none(c, theme="light", size=(8, 3.75))
        if png:
            # 整页图:等比缩放塞进 (1.5" → 6.9") 的区域,水平居中,不溢出页面。
            _place_picture_fit(
                slide, png, top=1.5, max_w=10.5, max_h=5.4,
            )
        else:
            # 渲染失败 → 显示退化说明
            _add_textbox(
                slide, Inches(0.8), Inches(1.7), Inches(10.9), Inches(0.5),
                "（图表渲染失败,以下为 ECharts 配置摘要）",
                font_size=14, color=COLOR_MUTED, anchor=MSO_ANCHOR.MIDDLE,
            )
            opt_str = _chart_option_json(c)
            if len(opt_str) > 1800:
                opt_str = opt_str[:1800] + "\n... (option 过长,已截断)"
            _add_textbox(
                slide, Inches(0.8), Inches(2.3), Inches(10.9), Inches(4.8),
                opt_str, font_size=11, color=COLOR_MUTED,
            )

    # ============================================================
    # 5. 关键指标页(numeric_summary)
    # ============================================================
    num = _numeric_summary(data["indicators"])
    if num:
        slide = _add_blank_slide()
        _add_horizontal_bar(slide, COLOR_TITLE_BG)
        _add_textbox(
            slide, Inches(0.6), Inches(0.4), Inches(11.3), Inches(0.8),
            "关键指标", font_size=28, bold=True, color=COLOR_TITLE_BG,
        )
        rows = [["指标", "合计", "均值", "最大", "最小"]]
        for k, v in num.items():
            rows.append([
                str(k),
                str(v.get("sum", "")),
                str(v.get("mean", "")),
                str(v.get("max", "")),
                str(v.get("min", "")),
            ])
        # 插入表格: 头部放正中区域
        table_left = Inches(1.0)
        table_top = Inches(1.5)
        table_width = Inches(10.5)
        n_rows = len(rows)
        n_cols = len(rows[0]) if rows else 1
        table_height = Inches(0.4 * (n_rows + 1))
        tbl_shape = slide.shapes.add_table(n_rows, n_cols, table_left,
                                            table_top, table_width,
                                            table_height)
        tbl = tbl_shape.table
        # 列宽：指标列 3",其余均分
        tbl.columns[0].width = Inches(3.5)
        rest = Inches(7.0)
        for i in range(1, n_cols):
            tbl.columns[i].width = Inches(7.0 / max(1, n_cols - 1))
        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                cell = tbl.cell(r_idx, c_idx)
                cell.text = ""
                tf = cell.text_frame
                tf.margin_left = Inches(0.08)
                tf.margin_right = Inches(0.08)
                tf.margin_top = Inches(0.04)
                tf.margin_bottom = Inches(0.04)
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = str(val)
                if r_idx == 0:
                    run.font.bold = True
                    run.font.color.rgb = COLOR_WHITE
                    run.font.size = Pt(16)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = COLOR_TITLE_BG
                else:
                    run.font.color.rgb = COLOR_TEXT
                    run.font.size = Pt(13)
                    if r_idx % 2 == 0:
                        cell.fill.solid()
                        # 关键指标表交替行底色:light 主题用白,business 用米白,
                        # 避免与白底表格混在一起看不出来是隔行。
                        cell.fill.fore_color.rgb = palette["background"]

    # ============================================================
    # 6. 兜底指标页(missing_ratio_top5 / text_value_counts)
    # ============================================================
    miss, tv = _fallback_indicators(data["indicators"])
    has_miss = bool(miss)
    has_tv = bool(tv)
    if has_miss or has_tv:
        slide = _add_blank_slide()
        _add_horizontal_bar(slide, COLOR_TITLE_BG)
        _add_textbox(
            slide, Inches(0.6), Inches(0.4), Inches(11.3), Inches(0.8),
            "字段明细", font_size=28, bold=True, color=COLOR_TITLE_BG,
        )
        top = 1.4
        if has_miss:
            _add_textbox(
                slide, Inches(0.6), Inches(top), Inches(11.3), Inches(0.5),
                "缺失率 Top5", font_size=18, bold=True, color=COLOR_ACCENT,
            )
            miss_rows = [["字段", "缺失率"]]
            for k, p in miss.items():
                miss_rows.append([str(k), f"{p}%"])
            tbl_shape = slide.shapes.add_table(
                len(miss_rows), 2,
                Inches(1.0), Inches(top + 0.6),
                Inches(10.5), Inches(0.35 * (len(miss_rows) + 1)),
            )
            tbl = tbl_shape.table
            tbl.columns[0].width = Inches(4.5)
            tbl.columns[1].width = Inches(6.0)
            for r_idx, row in enumerate(miss_rows):
                for c_idx, val in enumerate(row):
                    cell = tbl.cell(r_idx, c_idx)
                    cell.text = ""
                    tf = cell.text_frame
                    p = tf.paragraphs[0]
                    run = p.add_run()
                    run.text = str(val)
                    if r_idx == 0:
                        run.font.bold = True
                        run.font.color.rgb = COLOR_WHITE
                        run.font.size = Pt(14)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = COLOR_TITLE_BG
                    else:
                        run.font.color.rgb = COLOR_TEXT
                        run.font.size = Pt(12)
            top += 0.35 * (len(miss_rows) + 1) + 0.4

        if has_tv:
            for col, items in tv.items():
                title_tv = clean_llm_text(str(col))
                _add_textbox(
                    slide, Inches(0.6), Inches(top), Inches(11.3), Inches(0.5),
                    f"{title_tv} · 频次 Top8",
                    font_size=18, bold=True, color=COLOR_ACCENT,
                )
                tv_rows = [["取值", "出现次数"]]
                for it in items[:8]:
                    name = clean_llm_text(str(it.get("name", "")))
                    tv_rows.append([name, str(it.get("value", ""))])
                tbl_shape = slide.shapes.add_table(
                    len(tv_rows), 2,
                    Inches(1.0), Inches(top + 0.6),
                    Inches(10.5), Inches(0.35 * (len(tv_rows) + 1)),
                )
                tbl = tbl_shape.table
                tbl.columns[0].width = Inches(4.5)
                tbl.columns[1].width = Inches(6.0)
                for r_idx, row in enumerate(tv_rows):
                    for c_idx, val in enumerate(row):
                        cell = tbl.cell(r_idx, c_idx)
                        cell.text = ""
                        tf = cell.text_frame
                        p = tf.paragraphs[0]
                        run = p.add_run()
                        run.text = str(val)
                        if r_idx == 0:
                            run.font.bold = True
                            run.font.color.rgb = COLOR_WHITE
                            run.font.size = Pt(14)
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = COLOR_TITLE_BG
                        else:
                            run.font.color.rgb = COLOR_TEXT
                            run.font.size = Pt(12)
                top += 0.35 * (len(tv_rows) + 1) + 0.4

    # ============================================================
    # 7. 异常批注页
    # ============================================================
    annos = _annotations(data["annotations"])
    if annos:
        slide = _add_blank_slide()
        _add_horizontal_bar(slide, COLOR_TITLE_BG)
        _add_textbox(
            slide, Inches(0.6), Inches(0.4), Inches(11.3), Inches(0.8),
            "异常批注", font_size=28, bold=True, color=COLOR_TITLE_BG,
        )
        anno_rows = [["级别", "指标", "说明"]]
        for a in annos:
            anno_rows.append([str(a.get("level", "")), str(a.get("metric", "")),
                              str(a.get("comment", ""))])
        n = len(anno_rows)
        tbl_shape = slide.shapes.add_table(
            n, 3,
            Inches(0.5), Inches(1.4),
            Inches(11.5), Inches(0.4 * (n + 1)),
        )
        tbl = tbl_shape.table
        tbl.columns[0].width = Inches(2.0)
        tbl.columns[1].width = Inches(3.0)
        tbl.columns[2].width = Inches(6.5)
        for r_idx, row in enumerate(anno_rows):
            for c_idx, val in enumerate(row):
                cell = tbl.cell(r_idx, c_idx)
                cell.text = ""
                tf = cell.text_frame
                tf.word_wrap = True
                tf.margin_left = Inches(0.08)
                tf.margin_right = Inches(0.08)
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = str(val)
                if r_idx == 0:
                    run.font.bold = True
                    run.font.color.rgb = COLOR_WHITE
                    run.font.size = Pt(14)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = COLOR_TITLE_BG
                else:
                    run.font.color.rgb = COLOR_TEXT
                    run.font.size = Pt(11)
                    # 按 level 着色左边那一格(level=danger/warn/info)
                    if c_idx == 0:
                        level_str = str(val).lower()
                        if "danger" in level_str:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(0xfe, 0xe2, 0xe2)
                        elif "warn" in level_str:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(0xfe, 0xf3, 0xc7)

    # ============================================================
    # 8. 尾页(致谢)
    # ============================================================
    slide = _add_blank_slide()
    _fill_solid(slide, COLOR_TITLE_BG)
    _add_horizontal_bar(slide, COLOR_ACCENT, top=Inches(3.5))
    _add_textbox(
        slide, Inches(1.0), Inches(2.8), Inches(10.5), Inches(1.5),
        "谢谢观看", font_size=54, bold=True, color=COLOR_TITLE_TEXT,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    _add_textbox(
        slide, Inches(1.0), Inches(4.4), Inches(10.5), Inches(0.6),
        "本报告由 AI 数据处理工作台自动生成",
        font_size=16, color=COLOR_TITLE_TEXT,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )

    # 防御性兜底:防止外部绕过 helper 直接传非法值。_theme_palette 已经校验过一次,
    # 但此处仍二次防御,API 层会捕获并转 400。
    if theme not in _PPT_THEMES:
        raise ValueError(f"不支持的主题: {theme}")

    # 序列化
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()


# ============================================================
# AI 演示稿 PPT 导出（Spec: generate-ai-presentation-deck）
# ------------------------------------------------------------
# 与上面的 export_pptx（报告全文版）是**两条并列分支**：
#   - export_pptx      —— 把报告正文按 800 字切块搬进幻灯片（老行为，不改）
#   - export_deck_pptx —— 按 AI 重构的「结论 + 要点」大纲渲染（新增）
#
# 版式差异是本质的：
#   报告全文版 = 满屏文字（适合打印阅读）
#   演示稿版   = 大标题(结论句) + 3~5 条 bullet(+ 可选配图) + 演讲备注
#
# 每页还会写入 speaker notes(PowerPoint 演讲者视图可见)，
# 汇报人可以直接照着念。
# ============================================================

# 演示稿页面各区域的坐标(英寸)。12.5"×7.5" 画布。
#
# 【版式约束(2026-10 修订)】
#   用户实测反馈:"图表应该适应页面,这样排版并不是能直接用"。
#   根因是图表按固定宽度(7.7")插入,而 PNG 自身的宽高比不受控
#   (chart_renderer 之前用 bbox_inches="tight",裁掉了画布外的空白,
#   输出比例脱离 figsize 控制),导致图片实际高度 4.78",
#   底部落到 9.13" —— 而画布只有 7.5",图直接溢出页面并压住正文。
#
#   现在的规则:
#     1) 图表 PNG 一律用 light 主题(白底深字)渲染 —— PPT 页面是白底,
#        深色图块压在上面既不协调,深色图还会盖住浅色文字;
#     2) 插入时**同时**限制宽度与高度,保持原始宽高比(等比缩放),
#        任何情况下都不越出页面内容区;
#     3) 图表区的高度按"实际可用高度"倒推,而不是先定坐标再塞图。
_DECK_PAGE_WIDTH = 12.5           # 画布宽
_DECK_PAGE_HEIGHT = 7.5           # 画布高
_DECK_MARGIN_X = 0.8              # 左右页边距
_DECK_CONTENT_WIDTH = 10.9        # 内容区宽(12.5 - 2*0.8)
_DECK_TITLE_TOP = 0.5             # 标题区顶部
_DECK_TITLE_HEIGHT = 1.3          # 标题区高度(允许两行)
_DECK_BULLET_TOP = 1.95           # 要点区顶部
_DECK_PAGE_BOTTOM = 7.1           # 内容区底线(下方留 0.4" 呼吸位)


def _place_picture_fit(slide, png: bytes, top: float,
                       max_w: float, max_h: float,
                       page_w: float = 12.5) -> None:
    """把图片等比缩放后放进 (top, max_w, max_h) 的区域里,水平居中。

    为什么需要这个函数:
        之前是 `add_picture(..., width=Inches(8.0))` —— 只给宽度,
        高度由 PNG 自身的宽高比决定。而 PNG 比例不受控
        (chart_renderer 用 bbox_inches="tight" 裁掉了画布外空白,
        比例会随内容变化),实测 8" 宽的图高 5",底部落到 9.70",
        而画布只有 7.5" —— 图溢出页面,直接盖住正文,导出的 PPT 根本没法用。

    现在同时约束宽与高,取两者的**较小**缩放比,保证:
        - 宽高比不变(图片不变形)
        - 既不超宽也不超高,永远落在给定区域内
    """
    from pptx.util import Inches

    try:
        from PIL import Image
        with Image.open(io.BytesIO(png)) as im:
            px_w, px_h = im.size
    except Exception:
        # 拿不到尺寸时退化为"只按宽度",但仍用 max_w 兜底
        px_w, px_h = 8, 5

    if px_w <= 0 or px_h <= 0:
        return

    ratio = min(max_w / px_w, max_h / px_h)
    w = px_w * ratio
    h = px_h * ratio
    left = (page_w - w) / 2

    slide.shapes.add_picture(
        io.BytesIO(png),
        Inches(left), Inches(top),
        width=Inches(w), height=Inches(h),
    )


def export_deck_pptx(task, deck: dict, theme: str = "deep_cyan") -> bytes:
    """把 AI 演示稿大纲渲染成 16:9 PPTX(每页含演讲备注)。

    Args:
        task:  Task 对象(取任务名 / 创建时间)。
        deck:  演示稿大纲 {"title","subtitle","slides":[{kind,title,bullets,chart_ref,notes}]}。
        theme: 主题 key(deep_cyan / midnight / light / business)。

    Returns:
        pptx 文件字节。

    页序:封面 → (可选)目录 → 内容页×N → 结束页。
    """
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE

    if not isinstance(deck, dict) or not deck.get("slides"):
        raise ValueError("演示稿内容为空,无法导出 PPT")

    palette = _theme_palette(theme)
    color_title_bg = palette["title_bg"]
    color_accent = palette["accent"]
    color_muted = palette["muted"]
    color_title_text = palette["title_text"]
    color_background = palette["background"]

    # 【内容页文字色固定用深色,不跟随主题】(2026-10 修订)
    #   用户实测反馈:"生成的 ppt,文字应该是黑色"。
    #   根因有两层:
    #     1) 之前内容页正文用 palette["text"],而 midnight 主题的 text 是
    #        #e2e8f0(浅色)—— 内容页底色其实是白底(PPT 默认),
    #        浅色字落在白底上等于隐形,四套主题里有一半不可用;
    #     2) 标题用 palette["title_bg"](deep_cyan 是 #0e7490 深青),
    #        在白底上偏淡,投影仪上更看不清。
    #   演示稿是"拿去讲/拿去打印"的材料,内容页应当**始终**是白底深字,
    #   主题色只通过顶部色条体现 —— 这样四套主题下都清晰可读。
    color_page_title = RGBColor(0x11, 0x18, 0x27)   # 近黑:结论句
    color_page_text = RGBColor(0x1f, 0x29, 0x37)    # 深灰:要点正文

    prs = Presentation()
    prs.slide_width = Inches(12.5)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def _new_slide():
        return prs.slides.add_slide(blank_layout)

    def _fill_solid(slide, rgb_color):
        """整页填底色(封面/结束页用)。"""
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb_color

    def _add_bar(slide, color, top=None):
        """顶部一条色条作为视觉锚定。"""
        if top is None:
            top = Inches(0.0)
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, top, prs.slide_width, Inches(0.15)
        )
        bar.line.fill.background()
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        return bar

    def _add_text(slide, left, top, width, height, text,
                  font_size=18, bold=False, color=None,
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                  line_spacing=None, space_after=None):
        """加一个文本框;text 可以是 str 或 list[str](每项一个段落)。"""
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = Inches(0.04)
        tf.margin_right = Inches(0.04)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        lines = [text] if isinstance(text, str) else list(text)
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if line_spacing is not None:
                p.line_spacing = line_spacing
            if space_after is not None:
                p.space_after = space_after
            run = p.add_run()
            run.text = str(line)
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = color if color is not None else color_text
        return tb

    def _set_notes(slide, text: str):
        """写入演讲备注(PowerPoint 演讲者视图可见)。"""
        if not text:
            return
        try:
            slide.notes_slide.notes_text_frame.text = str(text)
        except Exception:
            # 备注写入失败不影响主流程(某些 pptx 变体可能不支持 notes_slide)
            pass

    # ---------- 图表匹配 ----------
    # AI 通过 chart_ref 引用图表(填 id 或标题)。这里做一次匹配:
    #   1) 先按 id 精确匹配;2) 再按标题精确匹配;3) 最后按标题包含匹配。
    # 同一张图不会被两页重复使用(used 集合)。
    from app.services.deck_service import build_chart_catalog

    try:
        raw_report = task.report_content or ""
        report = json.loads(raw_report) if raw_report else {}
    except Exception:
        report = {}
    try:
        chart_option = json.loads(task.chart_option_json or "{}") or {}
    except Exception:
        chart_option = {}

    catalog = build_chart_catalog(report, chart_option)
    used_chart_ids: set[str] = set()

    def _match_chart(ref: str) -> dict | None:
        """按 chart_ref 找一张未使用过的图表;找不到返回 None。"""
        ref = (ref or "").strip()
        if not ref:
            return None
        # 1) id 精确
        for item in catalog:
            if item["id"] == ref and item["id"] not in used_chart_ids:
                used_chart_ids.add(item["id"])
                return item
        # 2) 标题精确
        for item in catalog:
            if item.get("title") == ref and item["id"] not in used_chart_ids:
                used_chart_ids.add(item["id"])
                return item
        # 3) 标题包含(LLM 可能写了个别字)
        for item in catalog:
            title = str(item.get("title") or "")
            if title and (ref in title or title in ref) and item["id"] not in used_chart_ids:
                used_chart_ids.add(item["id"])
                return item
        return None

    # ---------- 1. 封面页 ----------
    slide = _new_slide()
    _fill_solid(slide, color_title_bg)
    _add_bar(slide, color_accent, top=Inches(3.0))
    _add_text(
        slide, Inches(1.0), Inches(1.9), Inches(10.5), Inches(1.4),
        str(deck.get("title") or "数据分析汇报"),
        font_size=44, bold=True, color=color_title_text,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    subtitle = str(deck.get("subtitle") or "").strip()
    if subtitle:
        _add_text(
            slide, Inches(1.0), Inches(3.35), Inches(10.5), Inches(0.7),
            subtitle, font_size=20, color=color_title_text,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
    _add_text(
        slide, Inches(1.0), Inches(4.15), Inches(10.5), Inches(0.5),
        f"任务: {getattr(task, 'name', '') or ''}",
        font_size=15, color=color_title_text,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    create_time_str = ""
    if getattr(task, "create_time", None):
        try:
            create_time_str = task.create_time.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            create_time_str = str(task.create_time)
    if create_time_str:
        _add_text(
            slide, Inches(1.0), Inches(4.62), Inches(10.5), Inches(0.5),
            f"生成时间:{create_time_str}",
            font_size=13, color=color_title_text,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
    _add_text(
        slide, Inches(1.0), Inches(6.2), Inches(10.5), Inches(0.6),
        "AI 数据处理工作台 · 自动生成",
        font_size=13, color=color_title_text,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )

    slides = [s for s in (deck.get("slides") or []) if isinstance(s, dict)]

    # ---------- 2. 目录页(仅当大纲里存在 agenda 页时) ----------
    agenda = next((s for s in slides if s.get("kind") == "agenda"), None)
    content_slides = [s for s in slides if s.get("kind") != "agenda"]
    if agenda is not None:
        slide = _new_slide()
        _add_bar(slide, color_accent)
        _add_text(
            slide, Inches(_DECK_MARGIN_X), Inches(_DECK_TITLE_TOP),
            Inches(_DECK_CONTENT_WIDTH), Inches(_DECK_TITLE_HEIGHT),
            "汇报提纲", font_size=28, bold=True, color=color_page_title,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        items = list(agenda.get("bullets") or [])
        if not items:
            # 大纲页没给 bullets 时,自动列出后续页标题(提纲页本来就是为了导览)
            items = [str(s.get("title") or "") for s in content_slides][:10]
        _add_text(
            slide, Inches(_DECK_MARGIN_X + 0.1), Inches(_DECK_BULLET_TOP),
            Inches(_DECK_CONTENT_WIDTH),
            Inches(_DECK_PAGE_BOTTOM - _DECK_BULLET_TOP),
            [f"· {it}" for it in items],
            font_size=19, color=color_page_text,
            line_spacing=1.5, space_after=Pt(14),
        )
        _set_notes(slide, str(agenda.get("notes") or ""))

    # ---------- 3. 内容页 ----------
    for item in content_slides:
        title = str(item.get("title") or "").strip()
        if not title:
            continue
        bullets = [str(b) for b in (item.get("bullets") or []) if str(b).strip()]

        slide = _new_slide()
        _add_bar(slide, color_accent)
        # 标题区:结论句,大字粗体。允许两行
        _add_text(
            slide, Inches(_DECK_MARGIN_X), Inches(_DECK_TITLE_TOP),
            Inches(_DECK_CONTENT_WIDTH), Inches(_DECK_TITLE_HEIGHT),
            title, font_size=28, bold=True, color=color_page_title,
            anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15,
        )

        # 先把配图渲染出来 —— 需要知道它的真实宽高比才能给文字区留位置
        matched = _match_chart(str(item.get("chart_ref") or ""))
        png = None
        if matched is not None:
            # light 主题:白底深字。PPT 页面是白底,深色图块压在上面既不协调,
            # 而且深色大色块会盖住浅色正文(用户实测截图里的主要问题)。
            # 尺寸用 8×3.75(≈2.13:1 扁图),正好填满下半区,不浪费横向空间。
            png = _render_chart_png_or_none(
                matched.get("chart") or {}, theme="light", size=(8, 3.75),
            )

        # ---- 版式分配:按"有图 / 无图"和 bullet 条数自适应 ----
        bullet_lines = [f"· {b}" for b in bullets]
        if png is None:
            # 无图:要点区占满剩余空间,行距放宽,看着舒展
            _add_text(
                slide, Inches(_DECK_MARGIN_X + 0.1), Inches(_DECK_BULLET_TOP),
                Inches(_DECK_CONTENT_WIDTH),
                Inches(_DECK_PAGE_BOTTOM - _DECK_BULLET_TOP),
                bullet_lines,
                font_size=19, color=color_page_text,
                line_spacing=1.5, space_after=Pt(14),
            )
        else:
            # 有图:要点区只占上半,图占下半。两区的高度由 bullet 条数决定,
            # 图按可用高度**等比**缩放,保证不越出页面。
            n = max(1, len(bullet_lines))
            # 每条要点约 0.42"(19pt 字 + 行距 + 段后距),再加一点余量
            bullet_h = min(2.9, 0.42 * n + 0.15)
            bullet_bottom = _DECK_BULLET_TOP + bullet_h

            _add_text(
                slide, Inches(_DECK_MARGIN_X + 0.1), Inches(_DECK_BULLET_TOP),
                Inches(_DECK_CONTENT_WIDTH), Inches(bullet_h),
                bullet_lines,
                font_size=17, color=color_page_text,
                line_spacing=1.4, space_after=Pt(8),
            )

            # 图表区:从要点区下方 0.2" 开始,到页面底线结束
            chart_top = bullet_bottom + 0.2
            chart_max_h = _DECK_PAGE_BOTTOM - chart_top
            if chart_max_h > 0.4:
                _place_picture_fit(
                    slide, png,
                    top=chart_top,
                    max_w=_DECK_CONTENT_WIDTH,
                    max_h=chart_max_h,
                )

        # 演讲备注
        _set_notes(slide, str(item.get("notes") or ""))

    # ---------- 4. 结束页 ----------
    slide = _new_slide()
    _fill_solid(slide, color_title_bg)
    _add_bar(slide, color_accent, top=Inches(3.5))
    _add_text(
        slide, Inches(1.0), Inches(2.8), Inches(10.5), Inches(1.5),
        "谢谢观看", font_size=54, bold=True, color=color_title_text,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    _add_text(
        slide, Inches(1.0), Inches(4.4), Inches(10.5), Inches(0.6),
        "本演示稿由 AI 数据处理工作台自动生成",
        font_size=16, color=color_title_text,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )

    # 防御性兜底(与 export_pptx 一致)
    if theme not in _PPT_THEMES:
        raise ValueError(f"不支持的主题: {theme}")

    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()
