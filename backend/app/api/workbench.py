"""工作台同步执行接口。"""
from __future__ import annotations

import logging
import re
from datetime import datetime
from urllib.parse import quote
from typing import Any

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException
from fastapi.responses import Response
from sqlalchemy.orm import Session

from app.core.database import get_db
from app.models.dataset import Dataset
from app.models.report_template import ReportTemplate
from app.models.task import Task
from app.services import chart_renderer as cr
from app.services import pandas_engine as pe
# 【export-visual-report 新增】可视化报表生成链路
#   - _merge_sheet_indicators：多 sheet 指标合并（与 AI 分析同口径，避免两套合并逻辑漂移）
#   - dashboard_report：KPI 卡 / 明细 / 洞察 / AI 结论 / xlsx 渲染
from app.services.analyzer import _merge_sheet_indicators
from app.services.dashboard_report import (
    _adapt_chart_groups,
    build_visual_report as _build_visual_report,
    render_visual_report_excel as _render_visual_report_excel,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/workbench", tags=["workbench"])


# 注意:此前对 analysis_intent 设了 500 字硬上限,与"开放式反馈"理念冲突。
# 现已放开:仅做 strip 与非空判断,不做长度校验;若担心 Prompt 过大,
# analyzer / prompt_builder 处会按 token 上限做兜底截断(那个截断不影响用户输入存储)。


def _parse_intent(payload: dict) -> str:
    """从请求体里解析 analysis_intent,仅做 strip,不限制长度。"""
    return (payload.get("analysis_intent") or "").strip()


def _parse_selected_charts(payload: dict) -> list[dict]:
    """从请求体里解析 selected_charts。

    兼容两种形态:
        - 直接传 list:[{title, option, style_key}, ...]
        - 用字符串(JSON 序列化):极少数前端可能误传,做兜底解析
    """
    raw = payload.get("selected_charts")
    if raw is None:
        return []
    if isinstance(raw, list):
        return [c for c in raw if isinstance(c, dict)]
    if isinstance(raw, str) and raw.strip():
        import json as _json
        try:
            parsed = _json.loads(raw)
            return [c for c in parsed if isinstance(c, dict)] if isinstance(parsed, list) else []
        except Exception:
            return []
    return []


@router.post("/analyze")
def analyze(payload: dict, background: BackgroundTasks, db: Session = Depends(get_db)):
    """异步执行分析(与 /api/tasks 同模型:提交立即返回 task_id)。

    【2026-10 行为变更】原来是同步阻塞调用,跑完才返回 —— 单次 LLM 超时就会
    把整个工作台"卡死"到 axios 120s 超时,看上去就是「AI 断开了」。
    现在改为:
        1) 校验入参 → 建 task(status='processing')
        2) 把 run_analysis 丢进 BackgroundTasks,立即返回 {task_id, status}
        3) 前端轮询 GET /api/tasks/{id},直到 status=success / failed
        4) 失败时 task.error_message 已被 analyzer 落库,前端可读

    请求体字段:
        dataset_id:         必填
        template_id:        必填
        cleaning_rules:     可选,清洗规则 dict
        name:               可选,task 名称
        use_knowledge_base: 可选,bool
        kb_ids:             可选
        analysis_intent:    可选,分析重点(自由长度,后端不再截断)
        selected_charts:    可选,list[dict],用户在「图表编辑」前置预选的图表,
                            analyzer 会按 title 去重合并入最终报告。
    """
    dataset_id = payload.get("dataset_id")
    template_id = payload.get("template_id")
    cleaning_rules = payload.get("cleaning_rules") or {}
    use_kb = bool(payload.get("use_knowledge_base", False))
    kb_ids = payload.get("kb_ids") or ""
    # 仅做 strip,不限制长度 —— 由 prompt_builder 处按 token 兜底
    analysis_intent = _parse_intent(payload)
    selected_charts = _parse_selected_charts(payload)

    if not dataset_id or not template_id:
        raise HTTPException(400, "dataset_id 与 template_id 必填")
    if not db.get(Dataset, dataset_id):
        raise HTTPException(404, "数据集不存在")
    if not db.get(ReportTemplate, template_id):
        raise HTTPException(404, "模板不存在")

    task = Task(
        name=payload.get("name") or "工作台分析",
        dataset_id=dataset_id,
        template_id=template_id,
        status="processing",
        use_knowledge_base=use_kb,
        kb_ids=kb_ids,
    )
    db.add(task)
    db.commit()
    db.refresh(task)

    # 把分析扔到后台,立即返回 task_id。
    # BackgroundTasks 的回调里 analyzer.run_analysis 已经在外层包了 try/except:
    # 一旦 LLM 链路抛错,会自动把 task.status 写成 'failed' + 写 error_message,
    # 不会再让 task 永远卡在 processing。
    background.add_task(
        _do_workbench_analyze,
        task.id,
        cleaning_rules,
        analysis_intent,
        selected_charts,
    )

    return {
        "task_id": task.id,
        "status": task.status,   # processing
    }


def _do_workbench_analyze(
    task_id: int,
    cleaning_rules: dict,
    analysis_intent: str,
    selected_charts: list[dict],
) -> None:
    """BackgroundTasks 回调:开新 Session,跑 run_analysis。

    抽出来只为与 /api/tasks._do_analyze 同样写法一致 —— 复用 analyzer.run_analysis,
    让 30~60s 的同步分析不阻塞 HTTP 响应。

    analyzer.run_analysis 的外层 try/except 已经处理 LLM 抛错,
    这里仅做 Session 生命周期管理,不再重复捕获。
    """
    from app.core.database import SessionLocal
    from app.services.analyzer import run_analysis

    db = SessionLocal()
    try:
        task = db.get(Task, task_id)
        if task:
            run_analysis(
                db,
                task,
                cleaning_rules,
                analysis_intent=analysis_intent,
                selected_charts=selected_charts,
            )
    finally:
        db.close()


# ============================================================
# 【chart-edit-prebuilt-dashboard 新增】3 个新接口
# ------------------------------------------------------------
# 1) GET  /api/workbench/charts/styles
#    返回 3 套风格枚举,供前端 el-option 下拉使用
# 2) POST /api/workbench/charts/preview
#    前置生成可视化看板:dataset_id + style_key → 6~8 张 option
#    **不走 LLM**,纯本地 Pandas + 规则引擎
# 3) POST /api/workbench/charts/modify
#    自然语言指令改图:instruction + target_chart → 新 option
#    关键词规则引擎(折线/饼图/蓝绿/Top N/横向柱/同比线),命中即改;
#    未命中时回退 LLM 兜底(与既有 chart_edit 工具相同的语义)
# ============================================================


@router.get("/charts/styles")
def list_chart_styles() -> dict[str, Any]:
    """返回 3 套图表风格枚举(供前端下拉)。

    形态:{"styles": [{"key": "tech_advanced", "label": "高级科技风格"}, ...]}
    """
    return {"styles": cr.list_style_keys()}


@router.post("/charts/preview")
def preview_charts(payload: dict, db: Session = Depends(get_db)) -> dict[str, Any]:
    """前置生成可视化看板(无需先跑 AI 分析)。

    请求体:
        dataset_id: 必填
        style_key:  可选,默认 tech_advanced
        ai_curate:  可选(bool),默认 False。
                    True → 走【混合出图】:本地骨架 + AI 精选 Top 3 + 标题润色。
                           解决「图表过多且很多空」的痛点(2026-10 用户反馈)。
                           LLM 不可用时自动降级为本地精选,不会失败。
                    False → 维持旧路径:每个 sheet 完整出图(2~6 张)。
        target_count: 可选(int),ai_curate=True 时生效,默认 3。
                     控制每个 sheet 最终保留的图表数。

    流程:
        1) 读 dataset → 解析 Excel/CSV
           【chart-edit-prebuilt-dashboard 修复】
           改用 read_dataframe_per_sheet,每个 sheet 独立 DataFrame,
           并自动识别"第一行是占位说明行"的 Excel。
           旧实现把所有 sheet concat 成一张大表,导致列名错乱 +
           多 sheet 数据被稀释。
        2) 调用 chart_renderer.build_dashboard_options_per_sheet →
           每个 sheet 独立生成 2~6 张图表 option(本地规则,已含:
              汇总行剔除、日期粒度自适应、空白图过滤、信息密度评分)
        3) 【ai_curate=True 时】再调 cr.ai_curate_charts 对每个 sheet 的图
           列表做精选 + 标题润色;LLM 失败时退化为按信息密度排序取 Top N,
           永远保证 UI 上有图。
        4) 不调 LLM(若 ai_curate=False);ai_curate=True 时按需调 LLM

    返回:
        {
          "style_key": "...",
          "dataset_id": ...,
          "sheet_count": N,
          "chart_count": M,
          "ai_curated": bool,    # 是否走了 AI 精选路径(供前端展示标识)
          "sheets": [
            {"sheet": "<sheet1>", "charts": [{title, option, style_key, sheet}, ...]},
            ...
          ]
        }
    """
    dataset_id = payload.get("dataset_id")
    style_key = payload.get("style_key") or cr.STYLE_TECH_ADVANCED
    ai_curate = bool(payload.get("ai_curate", False))
    target_count = int(payload.get("target_count") or 3)

    if not dataset_id:
        raise HTTPException(400, "dataset_id 必填")
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")

    try:
        sheets = pe.read_dataframe_per_sheet(ds.file_path)
    except Exception as exc:
        raise HTTPException(422, f"文件读取失败:{exc}") from exc

    # 过滤空 sheet,保留所有有数据的 sheet
    sheets = [s for s in sheets if s.get("df") is not None and len(s["df"]) > 0]
    if not sheets:
        raise HTTPException(422, "数据集为空,无法生成可视化看板")

    groups = cr.build_dashboard_options_per_sheet(sheets, style_key=style_key)
    if not groups:
        raise HTTPException(422, "数据集字段不足,无法生成图表")

    # ---- 【2026-10 混合出图】ai_curate=True 时做精选 + 标题润色 ----
    if ai_curate:
        for g in groups:
            df = next((s["df"] for s in sheets if s.get("name") == g["sheet"]), None)
            df_summary = None
            try:
                if df is not None and len(df) > 0:
                    cols = ", ".join(str(c) for c in list(df.columns)[:6])
                    df_summary = f"{g['sheet']}({len(df)} 行,字段:{cols})"
            except Exception:
                df_summary = g.get("sheet") or ""
            g["charts"] = cr.ai_curate_charts(
                g["charts"], df_summary=df_summary,
                target_count=target_count, style_key=style_key,
            )

    # 拍平累计 chart 数(供前端一次性展示)
    chart_count = sum(len(g["charts"]) for g in groups)

    return {
        "sheets": groups,
        "style_key": style_key,
        "dataset_id": dataset_id,
        "sheet_count": len(groups),
        "chart_count": chart_count,
        "ai_curated": ai_curate,
    }


@router.post("/charts/modify")
def modify_chart(payload: dict, db: Session = Depends(get_db)) -> dict[str, Any]:
    """自然语言 / 快捷指令改图(本地规则引擎 + LLM 兜底)。

    请求体:
        dataset_id:     可选,有则用于 LLM 兜底取指标
        instruction:    必填,如 "换成折线图 + 配色换成蓝绿色"
        target_chart:   可选,{"title", "option", "style_key", "type"}
                        不传则视为"整套重新生成"的入口

    返回:
        {"title", "option", "style_key", "type", "explanation", "fallback"}
        - fallback=True 时表示规则/LLM 都失败,前端应保留原图
    """
    import json as _json
    from app.services.ai_service import chart_edit as ai_chart_edit  # 局部 import 避免循环

    instruction = (payload.get("instruction") or "").strip()
    target_chart = payload.get("target_chart") or {}
    style_key = payload.get("style_key") or target_chart.get("style_key") or cr.STYLE_TECH_ADVANCED
    dataset_id = payload.get("dataset_id")

    if not instruction:
        raise HTTPException(400, "instruction 必填")

    target_option = target_chart.get("option") if isinstance(target_chart, dict) else None
    target_type = (target_chart.get("type")
                    or _first_series_type(target_option or {})
                    or "bar")
    target_title = target_chart.get("title") if isinstance(target_chart, dict) else None

    # 【修复】配色诉求前置预处理:用户提到"配色/颜色/色板/color"等,
    # 或"换/改/变/用 + 颜色名"也算配色诉求,我们要确保传给 LLM 的 target_option
    # 里没有硬编码 series color,否则 LLM 直接 copy 旧色导致"换了配色但看起来没动"。
    # 这里提前清空,顺便把 base.color 也清掉,让 LLM 完全自由选择调色板。
    text_lower = (instruction or "").lower()
    palette_keywords = ("配色", "颜色", "色板", "色调", "配色方案", "color")
    color_named = (
        "红", "橙", "黄", "绿", "青", "蓝", "紫", "粉", "黑", "白", "灰", "棕",
        "金", "银", "玫红", "番茄", "草绿", "森林", "天空", "海蓝", "珊瑚", "莫兰迪",
        "马卡龙", "暖色", "冷色", "暗色", "亮色", "霓虹", "金属",
    )
    has_color_word = any(c in text_lower for c in color_named)
    has_color_action = any(v in text_lower for v in ("换", "改", "变", "用", "调", "做", "成"))
    explicit_color = any(p in text_lower for p in palette_keywords)
    is_color_intent = explicit_color or (has_color_word and has_color_action)
    if is_color_intent and isinstance(target_option, dict):
        _strip_hardcoded_colors(target_option)

    # ---- 1. 关键词规则引擎(命中即改,不再走 LLM)----
    rule_hit = _apply_rule(instruction, target_option or {}, target_type, target_title)
    if rule_hit:
        # 【修复】配色类指令(命中"红色/蓝绿色/..."palette_map)的优先级必须最高。
        # apply_style 会无条件把 out["color"] 写回 style_key 默认色板,
        # 导致用户的"配色换成蓝绿色"被主题色板覆盖回蓝绿色(巧合)/ 青色。
        # 这里记下用户选的调色板,apply_style 之后再写回一次,确保用户指令永远胜出。
        user_palette = (
            rule_hit["option"].get("color")
            if isinstance(rule_hit.get("option"), dict)
            else None
        )
        # 风格再 apply 一次,确保返回的 option 一定带 style_key 的色板
        rule_hit["option"] = cr.apply_style(rule_hit["option"], style_key)
        if user_palette:
            rule_hit["option"]["color"] = user_palette
        rule_hit["style_key"] = style_key
        return rule_hit

    # ---- 2. LLM 兜底(走既有 ai_service.chart_edit)----
    if not dataset_id:
        # 缺数据集时,直接退化:不动图,让前端走"重新生成看板"
        return {
            "title": target_title or "图表",
            "type": target_type,
            "option": target_option or {},
            "style_key": style_key,
            "explanation": "指令未命中规则且缺少 dataset_id,未能生成新图表;请重新生成整套看板。",
            "fallback": True,
            "error": "no_dataset_id",
        }
    if not db.get(Dataset, dataset_id):
        raise HTTPException(404, "数据集不存在")

    try:
        # 【修复】ai_service.chart_edit 的实际签名是 (instruction, current_option, df_meta),
        # 之前误传 5 个 kwarg,会抛 "unexpected keyword argument 'dataset_id'"。
        # 这里按正确签名调用:current_option 直接传 dict(函数内部会自己序列化清洗),
        # df_meta 把 prompt 所需字段塞进去。
        from app.services.pandas_engine import detect_column_types  # 局部 import
        import json as _json2
        try:
            _df = None
            # 【多 sheet 分项 2026-10】取当前数据集,做一次最轻量读取,
            # 仅用于拼 df_meta 的 columns/sample_rows。
            # 不再用 read_dataframe_all_sheets(列并集大表会让 columns 混入无关字段),
            # 改为读**第一张 sheet**——chart_edit 通常针对当前/当前 sheets 的图,用首 sheet 提供上下文够用。
            from app.models.dataset import Dataset as _Dataset
            _ds = db.get(_Dataset, int(dataset_id))
            if _ds:
                from app.services import pandas_engine as _pe
                _per_sheets = _pe.read_dataframe_per_sheet(_ds.file_path)
                if _per_sheets:
                    _df = _per_sheets[0]["df"]
                _types = _pe.detect_column_types(_df) if _df is not None else {}
                df_meta_inner = {
                    "columns": [{"name": c, "type": _types.get(c, "text")} for c in (_df.columns if _df is not None else [])],
                    "sample_rows": (_df.head(5).astype(object).where(_df.head(5).notna(), "").values.tolist() if _df is not None else []),
                    "title": target_title or "图表",
                }
            else:
                df_meta_inner = {"columns": [], "sample_rows": [], "title": target_title or "图表"}
        except Exception:
            df_meta_inner = {"columns": [], "sample_rows": [], "title": target_title or "图表"}

        ai_resp = ai_chart_edit(
            instruction=instruction,
            current_option=target_option or {},
            df_meta=df_meta_inner,
        )
        # ai_chart_edit 已经走完规则+LLM,直接拿结果
        new_option = ai_resp.get("option") if isinstance(ai_resp, dict) else None
        new_type = ai_resp.get("type") or target_type
        new_title = ai_resp.get("title") or target_title
        explanation = ai_resp.get("explanation") or "已根据指令改图(LLM 兜底)"
        is_llm_failed = bool(ai_resp.get("fallback")) if isinstance(ai_resp, dict) else False
        err_text = (ai_resp.get("error") if isinstance(ai_resp, dict) else "") or ""

        # 【修复】LLM 完全解析失败时(两次都未拿到 JSON),前端展示"两次调用都未拿到 JSON"
        # 体验很糟。这里做一次"本地降级"——
        # 1) 拿到目标 series 列表,识别是 bar/line/pie 等类型
        # 2) 从 instruction 提取颜色词,匹配一个基础调色板
        # 3) 强制给 series.itemStyle.color 填上新颜色,确保视觉上"用户说改就改"
        # 4) 标题保持原标题,explanation 标注"本地兜底"
        # 【修复】LLM 解析失败 + 用户是配色诉求时,做本地降级。
        # 原 trigger 条件太严(error 字符串包含特定字样),但实际 LLM 失败模式很多。
        # 关键:不能对所有失败都改色——用户输「只显示 Top 5」时,LLM 失败也不该瞎改颜色。
        # 必须确认 instruction 真的是配色诉求才走本地降级。
        text_low = (instruction or "").lower()
        is_color_intent_local = (
            any(p in text_low for p in ("配色", "颜色", "色板", "色调", "配色方案", "color"))
            or (
                any(c in text_low for c in ("红", "橙", "黄", "绿", "青", "蓝", "紫", "粉",
                                              "黑", "白", "灰", "棕", "金", "银", "玫红",
                                              "番茄", "草绿", "森林", "天空", "海蓝", "珊瑚",
                                              "莫兰迪", "马卡龙", "暖色", "冷色", "暗色",
                                              "亮色", "霓虹", "金属"))
                and any(v in text_low for v in ("换", "改", "变", "用", "调", "做", "成"))
            )
        )
        if is_llm_failed and is_color_intent_local:
            # 简单本地调色板(中文色名 → 调色板)
            local_palette = None
            palette_table = [
                (("绿色", "绿", "草绿", "森林"), ["#10b981", "#22d3ee", "#84cc16", "#4ade80", "#34d399"]),
                (("蓝绿", "青色", "青", "天空", "海蓝"), ["#10b981", "#06b6d4", "#22d3ee", "#84cc16", "#0ea5e9"]),
                (("蓝色", "蓝"), ["#3b82f6", "#60a5fa", "#0ea5e9", "#1d4ed8", "#2563eb"]),
                (("红色", "红", "番茄", "玫红", "珊瑚"), ["#ef4444", "#f87171", "#dc2626", "#fb7185", "#f43f5e"]),
                (("紫色", "紫", "莫兰迪"), ["#a855f7", "#8b5cf6", "#c084fc", "#7c3aed", "#9333ea"]),
                (("橙色", "橙", "金", "黄", "金棕"), ["#f59e0b", "#fb923c", "#fbbf24", "#facc15", "#f97316"]),
                (("粉", "玫"), ["#ec4899", "#f472b6", "#fb7185", "#f9a8d4", "#fbcfe8"]),
                (("黑", "暗", "金属"), ["#475569", "#334155", "#1e293b", "#0f172a", "#64748b"]),
            ]
            for keywords, pal in palette_table:
                if any(k in text_low for k in keywords):
                    local_palette = pal
                    break
            if local_palette is None:
                # 没识别出具体颜色:用主题默认调色板中前 1 个色作为兜底
                try:
                    preset = cr.get_style_preset(style_key)
                    local_palette = preset.get("color", [])[:5] if isinstance(preset, dict) else None
                except Exception:
                    local_palette = None
            if local_palette:
                # 取原 target_option 拷贝作为 base,改 series 颜色
                base = json_safe_copy_dict(target_option or {})
                if not base:
                    base = {"title": {"text": target_title or "图表"}, "series": []}
                for i, s in enumerate(base.get("series") or []):
                    if not isinstance(s, dict):
                        continue
                    color = local_palette[i % len(local_palette)]
                    # 强制设 itemStyle.color
                    if not isinstance(s.get("itemStyle"), dict):
                        s["itemStyle"] = {}
                    s["itemStyle"]["color"] = color
                    # line / area 顺手清掉
                    if isinstance(s.get("lineStyle"), dict):
                        s.pop("lineStyle", None)
                base["color"] = local_palette
                new_option = base
                new_type = target_type
                new_title = target_title or "图表"
                explanation = "LLM 解析失败,已按指令本地降级配色"

        # 【修复】如果用户提到"配色/颜色",让 LLM 的 color 胜出,不要被 apply_style
        #   的主题色板覆盖。同 modify_chart 配色诉求路径,逻辑统一。
        user_color = (
            new_option.get("color")
            if isinstance(new_option, dict) and "color" in new_option
            else None
        )
        styled = cr.apply_style(new_option or {}, style_key)
        if user_color:
            styled["color"] = user_color
        return {
            "title": new_title,
            "type": new_type,
            "option": styled,
            "style_key": style_key,
            "explanation": explanation,
            "fallback": bool(ai_resp.get("fallback")) if isinstance(ai_resp, dict) else False,
            "error": ai_resp.get("error") if isinstance(ai_resp, dict) else None,
        }
    except Exception as exc:
        return {
            "title": target_title or "图表",
            "type": target_type,
            "option": target_option or {},
            "style_key": style_key,
            "explanation": f"指令处理失败:{exc}",
            "fallback": True,
            "error": str(exc),
        }


def _first_series_type(option: dict) -> str:
    if not isinstance(option, dict):
        return ""
    series = option.get("series")
    if isinstance(series, list) and series:
        first = series[0]
        if isinstance(first, dict):
            return str(first.get("type") or "").strip()
    return ""


def _apply_rule(instruction: str, target_option: dict, target_type: str,
                 target_title: str | None) -> dict | None:
    """关键词规则引擎:命中常用指令时直接改写 option。

    支持的指令(中文):
        - 换成折线图 / line
        - 换成柱状图 / bar
        - 换成饼图 / pie
        - 改成横向柱状图 / horizontal
        - 配色换成蓝绿色 / 配色换成绿色 / 配色换成橙色
        - 只显示 Top 5 / Top 10
        - 加一条同比增长折线 / 同比线

    返回:
        None → 未命中,调用方走 LLM 兜底
        dict  → {"title","type","option","explanation"},调用方再 apply_style
    """
    text = (instruction or "").lower()
    # 复制 option 避免污染
    base = json_safe_copy_dict(target_option)
    if not base:
        base = {"title": {"text": target_title or "图表"}, "series": []}

    # 类型替换
    type_map = {
        "折线图": "line", "折线": "line", "line": "line",
        "柱状图": "bar", "柱状": "bar", "bar": "bar",
        "饼图": "pie", "饼": "pie", "pie": "pie",
        "雷达图": "radar", "雷达": "radar",
        "漏斗图": "funnel", "漏斗": "funnel",
        "散点图": "scatter", "散点": "scatter",
        "面积图": "area",
    }
    new_type = target_type
    for k, v in type_map.items():
        if k in text and ("换" in text or "改" in text or "变" in text):
            new_type = v
            break

    # 横向柱
    if "横向" in text and ("柱" in text or "bar" in text):
        # ECharts 横向柱:xAxis.value + yAxis.category(data)
        xa = base.get("xAxis")
        ya = base.get("yAxis")
        if isinstance(xa, dict) and isinstance(ya, dict):
            xa_data = ya.get("data") if isinstance(ya.get("data"), list) else []
            ya_data = xa.get("data") if isinstance(xa.get("data"), list) else []
            if xa_data and not ya_data:
                base["xAxis"] = {**xa, "type": "value"}
                base["yAxis"] = {**ya, "type": "category", "data": xa_data}
        new_type = "bar"

    # Top N 截断
    import re
    m = re.search(r"top\s*(\d+)|只显示\s*(\d+)", instruction or "")
    if m:
        n = int(m.group(1) or m.group(2) or 5)
        for s in base.get("series") or []:
            if not isinstance(s, dict):
                continue
            data = s.get("data")
            if isinstance(data, list) and data and not isinstance(data[0], dict):
                s["data"] = data[:n]
            elif isinstance(data, list) and data and isinstance(data[0], dict):
                s["data"] = data[:n]
        # 同时把 xAxis.data / yAxis.data 截断到 N(取前 N)
        for axis_key in ("xAxis", "yAxis"):
            ax = base.get(axis_key)
            if isinstance(ax, dict) and isinstance(ax.get("data"), list):
                ax["data"] = ax["data"][:n]

    # 【修复】配色诉求:规则引擎完全放手,统一交给 LLM。
    # 老代码把"配色换成红色"硬映射成 ["#ef4444", ...] 这种固定数组,
    # 关键词字典永远覆盖不完人类的颜色描述(淡红/番茄红/玫瑰金/莫兰迪...),
    # 用户实际想要的色调整天被关键词挡在外面,体验非常受限。
    # 改动:检测到"配色/颜色/色板/color"等诉求,或任意"颜色名+换/改/变/用"组合,
    # 都直接 return None,让 modify_chart 走 LLM 兜底。
    # 硬编码色的清理工作由 modify_chart 调用方在此之前完成(避免污染 LLM prompt)。
    palette_keywords = ("配色", "颜色", "色板", "色调", "配色方案", "color")
    # 扩展:常见中文颜色名 + 修饰词,搭配"换/改/变/用"时也算配色诉求
    color_named = (
        "红", "橙", "黄", "绿", "青", "蓝", "紫", "粉", "黑", "白", "灰", "棕",
        "金", "银", "玫红", "番茄", "草绿", "森林", "天空", "海蓝", "珊瑚", "莫兰迪",
        "马卡龙", "暖色", "冷色", "暗色", "亮色", "霓虹", "金属",
    )
    has_color_word = any(c in text for c in color_named)
    has_color_action = any(v in text for v in ("换", "改", "变", "用", "调", "做", "成"))
    explicit_color = any(p in text for p in palette_keywords)
    mentions_color = explicit_color or (has_color_word and has_color_action)
    if mentions_color:
        return None

    # 同比线:加一条新 series(name="同比")
    if ("同比" in text or "增长" in text) and "线" in text:
        series = base.get("series") or []
        if series and isinstance(series[0], dict):
            first = series[0]
            data = first.get("data") or []
            if isinstance(data, list) and data and not isinstance(data[0], dict):
                # 简单算同比:每个值 / 前一个值 - 1,首位置 0
                yoy = [0]
                for i in range(1, len(data)):
                    try:
                        prev = float(data[i - 1])
                        cur = float(data[i])
                        yoy.append(round((cur / prev - 1) * 100, 2) if prev else 0)
                    except Exception:
                        yoy.append(0)
                # 仅折线/柱图叠加,饼图 / 雷达不处理
                if first.get("type") in ("bar", "line"):
                    series.append({
                        "name": "同比(%)",
                        "type": "line",
                        "yAxisIndex": 1,
                        "data": yoy,
                        "smooth": True,
                        "lineStyle": {"width": 2, "type": "dashed"},
                        "itemStyle": {"color": "#f59e0b"},
                        "label": {"show": True, "formatter": "{c}%"},
                    })
                    # 追加第二条 yAxis
                    ya = base.get("yAxis")
                    if isinstance(ya, dict):
                        base["yAxis"] = [ya, {"type": "value", "name": "同比(%)",
                                                 "position": "right",
                                                 "axisLabel": {"formatter": "{value}%"}}]
                    elif isinstance(ya, list):
                        ya.append({"type": "value", "name": "同比(%)",
                                    "position": "right",
                                    "axisLabel": {"formatter": "{value}%"}})
                    base["yAxisIndex"] = 0  # 占位,具体 series 已各自指定

    # 把首 series 的 type 改成新类型(若不是改横向/同比)
    if new_type != target_type:
        for s in base.get("series") or []:
            if isinstance(s, dict):
                s["type"] = new_type

    if (new_type == target_type
        and "color" not in base
        and "横向" not in (instruction or "")
        and not m
        and "同比" not in (instruction or "")):
        # 啥规则都没命中
        return None

    return {
        "title": target_title or (base.get("title") or {}).get("text") or "图表",
        "type": new_type,
        "option": base,
        "explanation": f"已按指令修改:类型={new_type}",
        "fallback": False,
    }


def json_safe_copy_dict(obj: dict) -> dict:
    """仅 dict 深拷贝,其他类型直接返回。供 _apply_rule 局部使用。"""
    if not isinstance(obj, dict):
        return {}
    import copy
    try:
        return copy.deepcopy(obj)
    except Exception:
        return dict(obj)


def _strip_hardcoded_colors(option: dict) -> None:
    """把 ECharts option 里所有硬编码的 color(series.itemStyle/lineStyle/areaStyle、
    series.data[].itemStyle、series.data[].color 等)全部清空,让 LLM 重新选择调色板时不会
    因为旧色残留而"看起来没换"。

    直接 in-place 修改,不需要返回值。
    """
    if not isinstance(option, dict):
        return
    for s in option.get("series") or []:
        if not isinstance(s, dict):
            continue
        # 清掉 itemStyle.color
        item_style = s.get("itemStyle")
        if isinstance(item_style, dict) and "color" in item_style:
            item_style.pop("color", None)
        # 折线 lineStyle.color
        line_style = s.get("lineStyle")
        if isinstance(line_style, dict) and "color" in line_style:
            line_style.pop("color", None)
        # 区域 areaStyle.color
        area_style = s.get("areaStyle")
        if isinstance(area_style, dict) and "color" in area_style:
            area_style.pop("color", None)
        # 顶层 color(部分 series 自己挂 color)
        if "color" in s:
            s.pop("color", None)
        # data[].itemStyle.color、data[].color、data[].value 一并清
        inner_data = s.get("data")
        if isinstance(inner_data, list):
            for di, d in enumerate(inner_data):
                if not isinstance(d, dict):
                    continue
                # 饼图 data 项常见 {name, value, itemStyle:{color}}
                inner_is = d.get("itemStyle")
                if isinstance(inner_is, dict) and "color" in inner_is:
                    inner_is.pop("color", None)
                if "color" in d:
                    d.pop("color", None)
                inner_data[di] = d
    # 清掉 base.color(让 LLM 自由选色,而不是被 option.color 调色板锁住)
    if "color" in option:
        option.pop("color", None)
    # label / axisLabel / tooltip 里也可能钉死颜色,留作 safety,只清"color"键
    # (避免波及 textStyle 等其他字段)
    label = option.get("label")
    if isinstance(label, dict) and "color" in label:
        label.pop("color", None)


# ============================================================
# 【export-visual-report 新增】导出 Excel 可视化报表（弹窗入口）
# ------------------------------------------------------------
# 与「下载整个看板」的区别：
#   看板 = 纯位图长图，不可编辑、不可筛选、无 KPI、无结论、无明细；
#   报表 = 结构化 xlsx，含 KPI 卡 + AI 结论 + 数据明细 + 图表洞察。
#
# 为什么指标要在后端"重读数据源"再算：
#   前端提交的只有 dataset_id 与图表 option，KPI 卡与明细表必须来自
#   Pandas 实算的真实指标，才能复用项目既有的"防编造数字"体系
#   （数值池 + 溯源复核），而不是信任前端传来的数字。
#
# 为什么图表 option 直接用前端传回的：
#   保证"所见即所得" —— 用户在看板上看到的图就是报表里的图，
#   不会因为后端重新跑规则引擎而变成另一批图。
# ============================================================
@router.post("/charts/report")
def export_chart_report(payload: dict, db: Session = Depends(get_db)) -> Response:
    """把用户在弹窗里勾选的图表导出为一份 Excel 可视化报表。

    请求体：
        dataset_id: 必填，用于后端重读数据源、实算 KPI 与明细
        style_key:  可选，非法时回落 tech_advanced
        include_ai: 可选(bool)，默认 True，控制是否尝试生成 AI 结论
        charts:     必填非空，[{title, option, sheet?, png_base64?}]
                    png_base64 为前端离屏截图（纯 base64，无 dataURL 前缀），
                    缺失时后端用 matplotlib 高质兜底渲染

    返回：xlsx 文件流（attachment）
    """
    dataset_id = payload.get("dataset_id")
    style_key = payload.get("style_key") or cr.STYLE_TECH_ADVANCED
    include_ai = bool(payload.get("include_ai", True))
    raw_charts = payload.get("charts") or []

    # ---- 入参校验 ----
    if not dataset_id:
        raise HTTPException(400, "dataset_id 必填")
    if not isinstance(raw_charts, list) or not raw_charts:
        # 用 422 表达"语义上无法处理"，前端会直接把 detail 文案弹给用户
        raise HTTPException(422, "请先勾选要导出的图表，再生成可视化报表")
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")

    # ---- 1) 重读数据源 → 逐 sheet 算指标 → 合并 ----
    # 复用 preview_charts 的既有链路：read_dataframe_per_sheet 每个 sheet 独立 DataFrame
    # （避免跨 sheet concat 把列名错乱），再逐 sheet compute_indicators 后合并。
    try:
        sheets = pe.read_dataframe_per_sheet(ds.file_path)
    except Exception as exc:
        raise HTTPException(422, f"文件读取失败:{exc}") from exc

    sheets = [s for s in sheets if s.get("df") is not None and len(s["df"]) > 0]
    if not sheets:
        raise HTTPException(422, "数据集为空,无法生成可视化报表")

    sheet_indicators: list[dict[str, Any]] = []
    for s in sheets:
        try:
            ind = pe.compute_indicators(s["df"])
            ind["name"] = s.get("name") or ""
            sheet_indicators.append(ind)
        except Exception as exc:
            # 单 sheet 算指标失败不影响整份报表（其它 sheet 的 KPI 仍然有效）
            logger.warning("报表导出：sheet %s 指标计算失败: %s", s.get("name"), exc)

    indicators = _merge_sheet_indicators(sheet_indicators) if sheet_indicators else {}

    # ---- 2) 图表分组（按前端传入的 sheet 字段分组，无 sheet 归「默认」）----
    # 直接用前端传回的 option，保证"所见即所得"。
    chart_groups = _adapt_chart_groups("direct", raw_charts)
    if not chart_groups:
        raise HTTPException(422, "图表数据为空,无法生成可视化报表")

    # ---- 3) 组装模型 → 渲染 xlsx ----
    report_model = _build_visual_report(
        title=f"{ds.name or '数据集'} · 可视化报表",
        style_key=style_key,
        indicators=indicators,
        chart_groups=chart_groups,
        dataset_name=ds.name or "",
        detail_tables=None,
        include_ai=include_ai,
    )
    try:
        content = _render_visual_report_excel(report_model)
    except Exception as exc:
        logger.exception("可视化报表渲染失败")
        raise HTTPException(500, f"报表生成失败:{exc}") from exc

    # ---- 4) 文件名：visual_report_<数据集名>_<时间戳>.xlsx ----
    # 数据集名可能含中文标点/空格/斜杠，用 safe 版清洗，避免浏览器下载时文件名乱码或非法
    safe_name = re.sub(r'[\\/:*?"<>|\r\n]+', "_", (ds.name or "dataset")).strip() or "dataset"
    fname = f"visual_report_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return Response(
        content=content,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={"Content-Disposition": _attachment_header(fname)},
    )


def _attachment_header(fname: str) -> str:
    """构造 Content-Disposition（支持中文文件名）。

    HTTP 头只能承载 latin-1 字节，直接把中文名塞进 filename="..." 会抛
    UnicodeEncodeError（Starlette 会做 latin-1 编码）。标准做法是 RFC 5987 双写：
      - filename=<ASCII 兜底名>   给老浏览器
      - filename*=UTF-8''<百分号编码>  给现代浏览器（前端 downloadBlobResponse 优先读这个）
    """
    ascii_name = re.sub(r"[^A-Za-z0-9._-]+", "_", fname) or "visual_report.xlsx"
    return (
        f'attachment; filename="{ascii_name}"; '
        f"filename*=UTF-8''{quote(fname)}"
    )
