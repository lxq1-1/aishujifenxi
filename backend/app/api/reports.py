"""报告导出 API。"""
import base64
import json
import re
from datetime import datetime
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import Response
from sqlalchemy.orm import Session

from app.core.database import get_db
from app.models.task import Task
from app.services import exporter
# 【export-visual-report 新增】可视化报表生成链路（复用既有章节图表归一化 + 报表渲染）
from app.services.dashboard_report import (
    _adapt_chart_groups,
    build_visual_report as _build_visual_report,
    render_visual_report_excel as _render_visual_report_excel,
)

router = APIRouter(prefix="/api/reports", tags=["reports"])

# PPT 主题白名单:仅 pptx 格式接受非默认主题。
ALLOWED_THEMES = {"deep_cyan", "midnight", "light", "business"}

# PPT 导出来源白名单:
#   report —— 报告全文版(把正文切块搬进幻灯片,老行为,默认)
#   ai     —— AI 演示稿(结论 + 要点 + 配图 + 演讲备注)
ALLOWED_SOURCES = {"report", "ai"}


@router.get("/{task_id}/export")
def export_report(
    task_id: int,
    format: str = "md",
    theme: str = "deep_cyan",
    source: str = "report",
    db: Session = Depends(get_db),
):
    """导出报告。

    Args:
        task_id: 任务 ID。
        format: 导出格式,支持 pdf/docx/word/pptx/powerpoint/excel/xlsx/md/markdown。
        theme: PPT 主题(仅 pptx 格式生效),可选 deep_cyan/midnight/light/business。
        source: PPT 导出来源(仅 pptx 格式生效),可选 report(报告全文版,默认)/ai(AI 演示稿)。
    """
    task = db.get(Task, task_id)
    if not task:
        raise HTTPException(404, "任务不存在")

    # 主题合法性统一校验:非法 theme 在所有分支之前直接 400,避免下游 export_pptx
    # 内部 ValueError 透出或被误传到非 pptx 分支。
    if theme not in ALLOWED_THEMES:
        raise HTTPException(
            400, f"不支持的主题: {theme};可选 "
                 f"deep_cyan/midnight/light/business"
        )

    # 导出来源合法性校验(同样在所有分支之前,避免被非 pptx 分支静默吞掉)。
    src = (source or "report").strip().lower()
    if src not in ALLOWED_SOURCES:
        raise HTTPException(
            400, f"不支持的导出来源: {source};可选 ai/report"
        )

    fmt = format.lower()
    if fmt == "pdf":
        # 非 pptx 的格式:必须使用默认主题,否则拒绝
        if theme != "deep_cyan":
            raise HTTPException(400, "主题仅在 pptx 格式下生效")
        if src != "report":
            raise HTTPException(400, "source 参数仅在 pptx 格式下生效")
        data = exporter.export_pdf(task)
        media = "application/pdf"
        fname = f"report_{task_id}.pdf"
    elif fmt == "docx" or fmt == "word":
        if theme != "deep_cyan":
            raise HTTPException(400, "主题仅在 pptx 格式下生效")
        if src != "report":
            raise HTTPException(400, "source 参数仅在 pptx 格式下生效")
        data = exporter.export_docx(task)
        media = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        fname = f"report_{task_id}.docx"
    elif fmt in ("pptx", "powerpoint"):
        # pptx 接受任意合法 theme(已在校验里通过)
        if src == "ai":
            # AI 演示稿分支:命中缓存直接用;没有则现场生成一次(兜底路径)。
            from app.services import deck_service
            deck = deck_service.get_deck(db, task_id)
            if not deck:
                if getattr(task, "status", None) != "success":
                    raise HTTPException(409, "任务尚未完成,无法生成演示稿")
                deck = deck_service.ensure_deck(db, task)
                if not deck:
                    raise HTTPException(
                        500, "演示稿生成失败,请在报告页重试生成后再导出"
                    )
            data = exporter.export_deck_pptx(task, deck, theme=theme)
        else:
            data = exporter.export_pptx(task, theme=theme)
        media = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        fname = f"report_{task_id}.pptx"
    elif fmt == "excel" or fmt == "xlsx":
        if theme != "deep_cyan":
            raise HTTPException(400, "主题仅在 pptx 格式下生效")
        if src != "report":
            raise HTTPException(400, "source 参数仅在 pptx 格式下生效")
        data = exporter.export_excel(task)
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        fname = f"report_{task_id}.xlsx"
    elif fmt in ("md", "markdown"):
        if theme != "deep_cyan":
            raise HTTPException(400, "主题仅在 pptx 格式下生效")
        if src != "report":
            raise HTTPException(400, "source 参数仅在 pptx 格式下生效")
        data = exporter.export_markdown(task)
        media = "text/markdown; charset=utf-8"
        fname = f"report_{task_id}.md"
    else:
        raise HTTPException(400, f"不支持的导出格式: {format};支持 pdf/docx/excel/md/pptx")

    return Response(
        content=data,
        media_type=media,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


# ============================================================
# AI 演示稿（Spec: generate-ai-presentation-deck）
# ------------------------------------------------------------
# 与「导出 PPT」的区别:
#   导出 PPT  —— 一次性下载文件
#   演示稿     —— 一份**持久化**的产物(大纲),报告页会展示状态、支持预览、
#                 支持导出讲义,并且是导出 PPT(source=ai)的输入。
# ============================================================

@router.post("/{task_id}/deck")
def generate_deck_endpoint(
    task_id: int,
    force: bool = False,
    db: Session = Depends(get_db),
):
    """生成(或复用)该任务的 AI 演示稿大纲。

    返回:
        {
          "status": "ready" | "failed" | "generating",
          "deck": {...} | None,
          "error": str | None
        }
    """
    from app.services import deck_service

    task = db.get(Task, task_id)
    if not task:
        raise HTTPException(404, "任务不存在")
    if getattr(task, "status", None) != "success":
        raise HTTPException(409, "任务尚未完成,无法生成演示稿")

    # 命中缓存:直接返回,不重复调 LLM
    if not force:
        cached = deck_service.get_deck(db, task_id)
        if cached:
            return {"status": "ready", "deck": cached, "error": None}

    # 已有"生成中"记录:说明另一个请求正在跑,不另起一次 LLM 调用
    latest = deck_service.get_latest_deck_row(db, task_id)
    if latest is not None and latest.status == "generating":
        # 生成中的记录超过 10 分钟视为"卡死",允许重新发起
        try:
            age = (datetime.now() - latest.create_time).total_seconds()
        except Exception:
            age = 0
        if age < 600:
            return {"status": "generating", "deck": None, "error": None}

    deck = deck_service.ensure_deck(db, task, force=force)
    if not deck:
        row = deck_service.get_latest_deck_row(db, task_id)
        error = (row.error_message if row is not None and row.error_message
                 else "演示稿生成失败")
        return {"status": "failed", "deck": None, "error": error}
    return {"status": "ready", "deck": deck, "error": None}


@router.get("/{task_id}/deck")
def get_deck_endpoint(
    task_id: int,
    theme: str = "deep_cyan",
    db: Session = Depends(get_db),
):
    """返回该任务的演示稿大纲 + 逐页缩略图(供前端预览面板渲染)。

    没有演示稿时返回 `{status: "none", slides: []}`(不报错),
    前端据此显示「尚未生成 + 生成按钮」。
    """
    from app.services import deck_service

    task = db.get(Task, task_id)
    if not task:
        raise HTTPException(404, "任务不存在")

    deck = deck_service.get_deck(db, task_id)
    if not deck:
        return {"status": "none", "deck": None, "slides": []}

    slides = _build_deck_preview_slides(task, deck)
    return {"status": "ready", "deck": deck, "slides": slides}


@router.get("/{task_id}/deck/script.md")
def get_deck_script(
    task_id: int,
    db: Session = Depends(get_db),
):
    """导出演示稿讲义(Markdown):标题 + 每页要点 + 演讲备注。"""
    from app.services import deck_service

    task = db.get(Task, task_id)
    if not task:
        raise HTTPException(404, "任务不存在")

    deck = deck_service.get_deck(db, task_id)
    if not deck:
        raise HTTPException(404, "演示稿尚未生成")

    md = deck_service.deck_to_markdown(deck, getattr(task, "name", "") or "")
    return Response(
        content=md,
        media_type="text/markdown; charset=utf-8",
        headers={
            "Content-Disposition": f'attachment; filename="report_{task_id}_script.md"'
        },
    )


def _build_deck_preview_slides(task, deck: dict) -> list[dict]:
    """构造演示稿的逐页缩略图数据。

    结构与现有 `preview_charts` 对齐,便于前端复用同一套卡片渲染:
        {
          "kind": "cover" | "agenda" | "content",
          "title": str,
          "bullets": [str, ...],
          "chart_png_b64": str | None,
          "notes": str,
        }
    """
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
    used: set[str] = set()

    def _match(ref: str) -> dict | None:
        ref = (ref or "").strip()
        if not ref:
            return None
        for item in catalog:
            if item["id"] == ref and item["id"] not in used:
                used.add(item["id"])
                return item
        for item in catalog:
            if item.get("title") == ref and item["id"] not in used:
                used.add(item["id"])
                return item
        for item in catalog:
            title = str(item.get("title") or "")
            if title and (ref in title or title in ref) and item["id"] not in used:
                used.add(item["id"])
                return item
        return None

    slides: list[dict] = [{
        "kind": "cover",
        "title": str(deck.get("title") or "演示稿"),
        "bullets": [str(deck.get("subtitle") or "")] if deck.get("subtitle") else [],
        "chart_png_b64": None,
        "notes": "",
    }]

    for item in deck.get("slides") or []:
        if not isinstance(item, dict):
            continue
        kind = str(item.get("kind") or "content")
        png_b64 = None
        matched = _match(str(item.get("chart_ref") or ""))
        if matched is not None:
            png = _render_thumb(exporter, matched.get("chart") or {})
            if png:
                png_b64 = base64.b64encode(png).decode("ascii")
        slides.append({
            "kind": kind,
            "title": str(item.get("title") or ""),
            "bullets": [str(b) for b in (item.get("bullets") or [])],
            "chart_png_b64": png_b64,
            "notes": str(item.get("notes") or ""),
        })

    slides.append({
        "kind": "ending",
        "title": "谢谢观看",
        "bullets": [],
        "chart_png_b64": None,
        "notes": "",
    })
    return slides


@router.get("/{task_id}/preview_charts")
def preview_charts(
    task_id: int,
    theme: str = "deep_cyan",
    db: Session = Depends(get_db),
):
    """返回当前报告在指定主题下的逐页渲染摘要(标题/正文/base64 PNG),
    用于前端 PPT 选择主题弹窗的逐页预览。"""
    from app.services import exporter as _exp

    if theme not in ALLOWED_THEMES:
        raise HTTPException(
            400, f"不支持的主题: {theme};可选 "
                 f"deep_cyan/midnight/light/business"
        )

    task = db.get(Task, task_id)
    if not task:
        raise HTTPException(404, "任务不存在")
    if getattr(task, "status", None) != "success":
        raise HTTPException(409, "任务尚未完成,无法预览")

    slides = _build_preview_slides(task, theme)
    return {"theme": theme, "slides": slides}


def _build_preview_slides(task, theme: str) -> list[dict]:
    """复用 exporter 的归一化逻辑,产出每个 PPT 页的"摘要视图"。

    每个 slide 字典形态:
        {
            "kind": "title" | "summary" | "section" | "chart"
                  | "metrics" | "fallback" | "annotation" | "ending",
            "title": str,
            "body":  str,
            "png_b64": str | None,   # base64 编码后的 PNG 字节
        }
    """
    from app.services import exporter as _exp

    data = _exp._read_task(task)
    report = data["report"]

    slides: list[dict] = []

    # 1. 标题页
    slides.append({
        "kind": "title",
        "title": _exp._title(report, data["name"]),
        "body": f"任务: {data['name']}",
        "png_b64": None,
    })

    # 2. 摘要页
    summary = _exp._summary(report)
    if summary:
        slides.append({
            "kind": "summary",
            "title": "AI 摘要",
            "body": summary,
            "png_b64": None,
        })

    # 3. 章节页(支持拆分多页) + inline 图
    by_section_charts, orphan_charts = _exp._collect_section_charts(
        report, data["chart_option"],
    )
    for sec_name, sec_body in _exp._sections(report):
        chunks = _exp._split_section_text(sec_body, max_chars=800)
        sec_inline = by_section_charts.get(sec_name) or []
        total = len(chunks)
        for idx, chunk in enumerate(chunks, start=1):
            heading = sec_name if total == 1 else f"{sec_name} ({idx}/{total})"
            png_b64 = None
            if sec_inline:
                # 取该章节第一张图的 PNG 作为缩略图
                png = _render_thumb(_exp, sec_inline[0])
                if png:
                    png_b64 = base64.b64encode(png).decode("ascii")
            slides.append({
                "kind": "section",
                "title": heading,
                "body": chunk[:120],
                "png_b64": png_b64,
            })

    # 4. 孤儿图表(末尾「数据图表」兜底页,每张图一张 slide)
    for i, c in enumerate(orphan_charts, 1):
        label = _exp._chart_label(c, i)
        png = _render_thumb(_exp, c)
        png_b64 = base64.b64encode(png).decode("ascii") if png else None
        slides.append({
            "kind": "chart",
            "title": label,
            "body": "" if png_b64 else "(图表渲染失败,以下为 ECharts 配置摘要)",
            "png_b64": png_b64,
        })

    # 5. 关键指标
    num = _exp._numeric_summary(data["indicators"])
    if num:
        body_lines = ["| 指标 | 合计 | 均值 | 最大 | 最小 |"]
        body_lines.append("|---|---|---|---|---|")
        for k, v in num.items():
            body_lines.append(
                f"| {k} | {v.get('sum', '')} | {v.get('mean', '')} "
                f"| {v.get('max', '')} | {v.get('min', '')} |"
            )
        slides.append({
            "kind": "metrics",
            "title": "关键指标",
            "body": "\n".join(body_lines),
            "png_b64": None,
        })

    # 6. 兜底指标
    miss, tv = _exp._fallback_indicators(data["indicators"])
    if miss or tv:
        body_lines: list[str] = []
        if miss:
            body_lines.append("缺失率 Top5")
            for k, p in miss.items():
                body_lines.append(f"  {k}: {p}%")
        if tv:
            for col, items in tv.items():
                body_lines.append(f"{col} · 频次 Top8")
                for it in items[:8]:
                    body_lines.append(
                        f"  {it.get('name', '')}: {it.get('value', '')}"
                    )
        slides.append({
            "kind": "fallback",
            "title": "字段明细",
            "body": "\n".join(body_lines),
            "png_b64": None,
        })

    # 7. 异常批注
    annos = _exp._annotations(data["annotations"])
    if annos:
        body_lines = []
        for a in annos:
            body_lines.append(
                f"[{a['level']}] {a['metric']} - {a['comment']}"
            )
        slides.append({
            "kind": "annotation",
            "title": "异常批注",
            "body": "\n".join(body_lines),
            "png_b64": None,
        })

    # 8. 尾页
    slides.append({
        "kind": "ending",
        "title": "谢谢观看",
        "body": "",
        "png_b64": None,
    })

    return slides


def _render_thumb(_exp, chart: dict) -> bytes | None:
    """缩略图专用:小尺寸 PNG。"""
    try:
        return _exp.render_chart_to_png(chart, width=4.8, height=2.7, dpi=100)
    except Exception:
        return None


# ============================================================
# 【export-visual-report 新增】导出 Excel 可视化报表（报告页入口）
# ------------------------------------------------------------
# 与既有 /export（pdf/docx/pptx/xlsx）的区别：
#   既有 excel 导出 = 报告正文 + 章节配图的流水账，没有指标概览；
#   本接口 = 结构化报表：KPI 指标卡 + AI 结论 + 数据明细 + 图表洞察。
#
# 为什么报告页入口"零 LLM 调用"：
#   task.report_content.summary 已是 AI 生成并已过既有审计的摘要，
#   直接复用即可 —— 既省一次串行 LLM（60s 级），又天然不会引入新编造数字。
#
# 前端截图匹配约定（与 ReportView.vue 严格一致）：
#   分组名 = 章节名（chart_option.by_section 的 key）
#   孤儿图（不在章节列表里的 key）统一归入「数据图表」组
#   单图 key = f"{分组名}|{序号}"，序号从 1 开始
#   匹配不上的图不报错，自动走 matplotlib 高质兜底渲染。
# ============================================================
@router.post("/{task_id}/visual-report")
def export_visual_report(
    task_id: int,
    payload: dict | None = None,
    db: Session = Depends(get_db),
):
    """把报告页上的图表导出为一份 Excel 可视化报表。

    请求体（全部可选）：
        include_ai: bool，默认 True
        charts: [{key, png_base64}]，前端离屏截图（按 key 与后端图表匹配）

    返回：xlsx 文件流（attachment）
    """
    payload = payload or {}
    task = db.get(Task, task_id)
    if not task:
        raise HTTPException(404, "任务不存在")

    include_ai = bool(payload.get("include_ai", True))
    shot_list = payload.get("charts") if isinstance(payload.get("charts"), list) else []

    # ---- 1) 读任务落库数据 ----
    report = _loads(task.report_content) or {}
    chart_option = _loads(task.chart_option_json) or {}
    indicators = _loads(task.indicators_json) or {}
    summary = (report.get("summary") or "").strip()

    # ---- 2) 图表按章节分组（复用既有归一化，孤儿图归「数据图表」）----
    by_section, orphan = exporter._collect_section_charts(report, chart_option)
    chart_groups = _adapt_chart_groups("by_section", by_section)
    if orphan:
        chart_groups.extend(_adapt_chart_groups("charts", orphan))

    # ---- 3) 前端截图按 key 回填到对应图表上 ----
    # shot_map: {"章节名|1": "<base64>"}
    shot_map = {
        str(s.get("key")): s.get("png_base64")
        for s in shot_list
        if isinstance(s, dict) and s.get("key") and s.get("png_base64")
    }
    if shot_map:
        for g in chart_groups:
            for c in g.get("charts") or []:
                key = c.get("key")
                if key and shot_map.get(key):
                    # 命中即用前端高保真截图；未命中保持 None → 后端兜底渲染
                    c["png_base64"] = shot_map[key]

    # ---- 4) 章节明细表（追加到「数据明细」Sheet 的后半段）----
    detail_tables = chart_option.get("detail_tables") if isinstance(chart_option, dict) else None

    # ---- 5) 组装模型 → 渲染 ----
    report_model = _build_visual_report(
        title=f"{task.name or '分析报告'} · 可视化报表",
        # 报告页没有"图表风格"概念，统一用商务职业风格：
        # 白底浅色专业版式，与报告页导出 PDF/Word 的观感一致。
        style_key="business_pro",
        indicators=indicators,
        chart_groups=chart_groups,
        dataset_name="",
        # 已有 AI 摘要 → 直接复用（source="report"），零额外 LLM 调用
        summary=summary or None,
        detail_tables=detail_tables,
        include_ai=include_ai,
    )
    try:
        content = _render_visual_report_excel(report_model)
    except Exception as exc:
        raise HTTPException(500, f"报表生成失败:{exc}") from exc

    safe_name = re.sub(r'[\\/:*?"<>|\r\n]+', "_", (task.name or "report")).strip() or "report"
    fname = f"visual_report_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return Response(
        content=content,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        # HTTP 头只支持 latin-1，中文文件名必须 RFC 5987 双写（详见 _attachment_header）
        headers={"Content-Disposition": _attachment_header(fname)},
    )


def _attachment_header(fname: str) -> str:
    """构造 Content-Disposition（支持中文文件名）。

    直接把中文塞进 filename="..." 会触发 Starlette 的 latin-1 编码异常，
    因此双写：ASCII 兜底名 + filename*=UTF-8'' 百分号编码。
    前端 downloadBlobResponse 优先解析 filename*，能拿到正确中文名。
    """
    ascii_name = re.sub(r"[^A-Za-z0-9._-]+", "_", fname) or "visual_report.xlsx"
    return (
        f'attachment; filename="{ascii_name}"; '
        f"filename*=UTF-8''{quote(fname)}"
    )


def _loads(raw) -> dict:
    """把任务里存成 JSON 字符串的字段安全解析为 dict（解析失败返回空 dict）。"""
    if isinstance(raw, dict):
        return raw
    if not isinstance(raw, str) or not raw.strip():
        return {}
    try:
        parsed = json.loads(raw)
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}
