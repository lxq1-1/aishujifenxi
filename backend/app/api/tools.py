"""辅助工具 API：合并 / 校验 / NL 查询 / 对比。"""
from __future__ import annotations

import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from app.core.config import settings
from app.core.database import get_db
from app.models.dataset import Dataset, INTERNAL_DATASET_STATUS
from app.models.report_template import ReportTemplate
from app.models.task import Task
from app.services import ai_service, pandas_engine as pe, pdf_service

# 给模块顶层补一份 `pd` 别名,供 /clean 在多 sheet 落盘时使用
# (沿用 pandas_engine 内部的同一份 pd,避免重复 import)
import pandas as pd  # noqa: E402

router = APIRouter(prefix="/api/tools", tags=["tools"])


# ============================================================================
# 【tools-results-history-persistence 新增】3 个工具落库共享工具函数
# 设计要点:
#   - /validate / /nl_query / /compare 三个工具都接收可选 persist: bool
#   - persist=True 且工具计算成功时,新建 Task 并把响应里塞 task_id / task_url
#   - persist=False 或失败时**不**碰 db,与现有行为完全一致
#   - 落库时**不**调 LLM,所有图由 pandas_engine.build_*_chart_option 直接生成
# ============================================================================

# 内置工具模板的 name(由 seed.py 启动时保证存在)
_VALIDATION_TEMPLATE_NAME = "数据校验报告"
_NL_TEMPLATE_NAME = "数据问答报告"
_COMPARE_TEMPLATE_NAME = "数据对比报告"

# diff_mode → 中文标签(用于对比任务名拼接)
_DIFF_MODE_CN_MAP: dict[str, str] = {
    "business": "业务对比",
    "row_diff": "变更追踪",
    # 【topn 并入 business / distribution 下线 新增】
    # 原 "distribution" / "topn" 独立模式已下线:
    #   - topn  → 并入 business 的 top_n 截断参数,任务名不再单独出现
    #   - distribution → 与单表统计能力重叠且输出信息量不足,直接下线
    # 【dataset-contract-compare-report 新增】合同对比标签,
    # 同时也是 ReportView 识别"这是合同报告"的关键特征
    "contract_compare": "合同对比",
}


def _build_contract_report_sections(
    name: str,
    narration: dict,
    clause_diffs: list,
) -> dict:
    """为合同对比场景拼装报告 sections(垂直场景模板,不带空块)。

    输出 4 个板块:
        1. 合同摘要   - str: AI 总结 + 高优先级变更数(纯文本)
        2. 条款差异表 - list[dict]: 12 条核心条款结构化数据,前端按条渲染
        3. 利弊分析   - list[dict]: 商务视角的左右方利弊,前端卡片化
        4. 谈判建议   - list[str]: 逐条建议,前端列表渲染(加 ⭐ 标记)

    设计要点:
        - 4 个板块**一定有内容**(保证前端不留空 h2)
        - 不写「数据图表/关键指标/异常批注」,避免空块
        - 谈判建议放在最后但加高亮标记,前端在 h2 上有 emoji,顺序自然引导
        - 【report-sections-structured 新增】板块值可以是 list/dict 而非纯 str,
          前端按值的类型分支渲染(报告页 contract_compare 分支下走结构化分支)

    报告页 sections 值的类型约定:
        - str : 通用报告模式,作为段落文本渲染(<p>{{ content }}</p>)
        - list : 合同报告的「条款差异表/利弊分析/谈判建议」,每项是一个 dict
        - dict : 预留扩展(目前合同场景不直接用,但前端的兜底分支会处理)
    """
    high_changes = [
        c for c in (clause_diffs or [])
        if c.get("importance") == "high"
        and c.get("change_type") in ("modified", "added", "removed")
    ]
    summary_text = narration.get("summary") or ""
    if high_changes:
        summary_text = f"⚠️ 识别到 {len(high_changes)} 项高优先级条款变更。\n\n{summary_text}".strip()

    # 【report-sections-structured 新增】条款差异表 → list[dict]
    # 每条用结构化对象,前端按「卡」渲染,自然带换行/分隔,不再被 <p> 压扁
    clause_list: list[dict] = []
    for c in clause_diffs or []:
        if not isinstance(c, dict):
            continue
        clause_list.append({
            "clause_label": c.get("clause_label") or c.get("clause_key") or "—",
            "importance":   c.get("importance") or "medium",
            "change_type":  c.get("change_type") or "unclear",
            "left_value":   c.get("left_value") or "",
            "right_value":  c.get("right_value") or "",
            "diff_note":    c.get("diff_note") or "",
        })
    if not clause_list:
        # 兜底:也用 list 形式,让前端分支能正常渲染"空"占位
        clause_list = [{"empty": True, "diff_note": "未识别到差异条款"}]

    # 【report-sections-structured 新增】利弊分析 → list[dict]
    # 每条带 side,前端按 left/right/both 分色卡片
    pros_cons: list[dict] = []
    for pc in (narration.get("pros_cons") or []):
        if not isinstance(pc, dict):
            continue
        side = pc.get("side") or "both"
        if side not in ("left", "right", "both"):
            side = "both"
        pros_cons.append({
            "side":   side,
            "title":  pc.get("title", ""),
            "detail": pc.get("detail", ""),
        })
    if not pros_cons:
        pros_cons = [{"side": "both", "title": "未生成利弊分析", "detail": ""}]

    # 【report-sections-structured 新增】谈判建议 → list[str]
    # 一行一项,前端用 <ul> 渲染并加 ⭐ icon
    tips: list[str] = [str(t) for t in (narration.get("negotiation_tips") or []) if t]
    if not tips:
        tips = ["暂无额外谈判建议"]

    return {
        "合同摘要":   summary_text or "已完成合同对比",
        "条款差异表": clause_list,    # type: ignore[dict-item]
        "利弊分析":   pros_cons,      # type: ignore[dict-item]
        "谈判建议":   tips,           # type: ignore[dict-item]
    }


def _find_builtin_template_id(db: Session, name: str) -> int | None:
    """按 name 查内置模板 id;不存在返回 None。

    用于三个工具的 persist 路径,模板不存在时让路由返回 500 提示。
    """
    tpl = (
        db.query(ReportTemplate)
        .filter(ReportTemplate.name == name)
        .first()
    )
    return tpl.id if tpl else None


def _now_str() -> str:
    """统一的时间戳格式:YYYY-MM-DD HH:MM"""
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _persist_validation_task(
    db: Session,
    ds: Dataset,
    result: dict,
    name_override: str | None,
) -> dict:
    """数据校验·落库为历史任务。

    Args:
        db: SQLAlchemy Session
        ds: 校验目标 dataset
        result: validate 路由返回的 dict(本函数会原地塞 task_id / task_url)
        name_override: 用户入参 name;空时由本函数拼默认名

    Returns:
        原 result(已塞 task_id / task_url;persist=False 时上层不会调用)

    Raises:
        HTTPException 500: 内置"数据校验报告"模板不存在
    """
    tpl_id = _find_builtin_template_id(db, _VALIDATION_TEMPLATE_NAME)
    if tpl_id is None:
        raise HTTPException(
            500,
            "数据校验模板未初始化,请重启后端(seed.py 启动幂等会补齐)",
        )

    ai = result.get("ai_interpret") or {}
    quality = result.get("quality") or {}
    column_types = result.get("column_types") or {}
    issues = ai.get("issues") or []

    # name 入参优先;否则拼默认
    name = (name_override or "").strip() or f"{ds.name} · 数据校验({_now_str()})"

    # 问题清单文本化
    issues_text = _render_issues_to_text(issues)

    # 三个 JSON 字段
    indicators_json = json.dumps(
        {"quality": quality, "column_types": column_types},
        ensure_ascii=False,
    )
    report_content = json.dumps(
        {
            "title": name,
            "summary": ai.get("summary") or "",
            "sections": {
                "校验概览": ai.get("summary") or "",
                "问题清单": issues_text,
                "处理建议": ai.get("suggestions") or [],
            },
        },
        ensure_ascii=False,
    )
    chart_option_json = json.dumps(
        {"by_section": {"数据图表": pe.build_validation_chart_option(quality, column_types)}},
        ensure_ascii=False,
    )

    task = Task(
        name=name,
        dataset_id=ds.id,
        template_id=tpl_id,
        status="success",
        indicators_json=indicators_json,
        report_content=report_content,
        chart_option_json=chart_option_json,
        annotations_json="[]",
        error_message="",
    )
    db.add(task)
    db.commit()
    db.refresh(task)
    result["task_id"] = task.id
    result["task_url"] = f"/reports/{task.id}"
    return result


def _persist_nl_task(
    db: Session,
    ds: Dataset,
    result: dict,
    question: str,
    name_override: str | None,
) -> dict:
    """数据对话·落库为历史任务。"""
    tpl_id = _find_builtin_template_id(db, _NL_TEMPLATE_NAME)
    if tpl_id is None:
        raise HTTPException(
            500,
            "数据问答模板未初始化,请重启后端(seed.py 启动幂等会补齐)",
        )

    columns = result.get("columns") or []
    rows = result.get("rows") or []
    explanation = result.get("explanation") or ""
    sheets_n = int(result.get("active_sheet_count") or 0)
    # 【deepen-data-chat-analysis】两阶段管道产出的结构化深度回答(可能为 None=LLM 降级)
    structured = result.get("structured") if isinstance(result.get("structured"), dict) else None

    # name 入参优先;否则按"<数据集名> · 问答:<问题前 30 字>..."拼接
    name = (name_override or "").strip()
    if not name:
        q = (question or "").strip()
        q_short = q[:30] + ("..." if len(q) > 30 else "")
        name = f"{ds.name} · 问答:{q_short}"

    if structured:
        # —— 结构化路径:结论/分析/建议/追问各占一个 section ——
        # 为什么分 section 而不是塞进一坨 summary:报告页按 section 渲染,
        # 用户能直接看到「结论→分析→明细→建议→可继续追问」的完整链路。
        st_answer = str(structured.get("answer") or "").strip()
        st_analysis = [
            str(x).strip() for x in (structured.get("analysis") or []) if str(x or "").strip()
        ]
        st_suggestions = [
            str(x).strip() for x in (structured.get("suggestions") or []) if str(x or "").strip()
        ]
        st_followups = [
            str(x).strip() for x in (structured.get("followups") or []) if str(x or "").strip()
        ]
        st_table = structured.get("data_table") if isinstance(structured.get("data_table"), dict) else {}
        st_cols = [str(c) for c in (st_table.get("columns") or [])]
        st_rows = list(st_table.get("rows") or [])
        # 支撑表优先用叙述层挑出的表;为空时回落到顶层 rows(通常是首 sheet 前 5 行样本)
        chart_cols = st_cols or columns
        chart_rows = st_rows or rows
        rows_text = _render_rows_to_markdown(chart_cols, chart_rows)
        summary = st_answer or explanation
        sections = {
            "问题": question,
            "AI 结论": st_answer,
            "详细分析": st_analysis,      # list[str] —— 报告页已有 list 板块渲染支持
            "结果明细": rows_text,
            "处理建议": st_suggestions,
            "推荐追问": st_followups,
        }
    else:
        # —— 降级路径:与旧版完全一致,保证历史报告格式不突变 ——
        chart_cols, chart_rows = columns, rows
        rows_text = _render_rows_to_markdown(columns, rows)
        summary = explanation
        sections = {
            "问题": question,
            "AI 回答": explanation,
            "结果明细": rows_text,
        }

    indicators_json = json.dumps(
        {
            "question": question,
            "row_count": len(chart_rows),
            "sheets": sheets_n,
            "structured_present": structured is not None,
        },
        ensure_ascii=False,
    )
    report_content = json.dumps(
        {
            "title": question,
            "summary": summary,
            "sections": sections,
        },
        ensure_ascii=False,
    )
    chart_option_json = json.dumps(
        {"by_section": {"数据图表": pe.build_nl_query_chart_option(chart_rows, chart_cols)}},
        ensure_ascii=False,
    )

    task = Task(
        name=name,
        dataset_id=ds.id,
        template_id=tpl_id,
        status="success",
        indicators_json=indicators_json,
        report_content=report_content,
        chart_option_json=chart_option_json,
        annotations_json="[]",
        error_message="",
    )
    db.add(task)
    db.commit()
    db.refresh(task)
    result["task_id"] = task.id
    result["task_url"] = f"/reports/{task.id}"
    return result


def _persist_compare_task(
    db: Session,
    ds_l: Dataset,
    ds_r: Dataset,
    result: dict,
    mode: str,
) -> dict:
    """数据集对比·落库为历史任务(沿用对比 spec 设计)。

    【dataset-contract-compare-report 新增】按 mode 分流报告模板:
        - 表格/通用对比 4 种:沿用旧的 4 板块(关键发现/对比详情/风险/建议)+ 数据图表
        - 合同对比:        用 _build_contract_report_sections 拼装 4 个专属板块,
                           图表与关键指标留空,避免前端出现"对比结果为空"等无意义空块
    """
    tpl_id = _find_builtin_template_id(db, _COMPARE_TEMPLATE_NAME)
    if tpl_id is None:
        raise HTTPException(
            500,
            "数据对比模板未初始化,请重启后端(seed.py 启动幂等会补齐)",
        )

    mode_cn = _DIFF_MODE_CN_MAP.get(mode, mode)
    name = f"{ds_l.name} vs {ds_r.name} · {mode_cn}({_now_str()})"

    indicators = result.get("indicators") or {}
    narration = result.get("narration") or {}
    diff_details = result.get("diff_details")

    # —— 模式分流:合同对比用专属模板,其他用通用模板 ——
    is_contract = (mode == "contract_compare")
    if is_contract:
        sections = _build_contract_report_sections(name, narration, diff_details)
        # 合同场景没有结构化指标,indicators 存原文即可,前端在关键指标区块会判空
        chart_option = {"by_section": {}, "indicators": indicators}
    else:
        diff_detail_payload = _summarize_diff_details_to_text(diff_details, mode)
        sections = {
            "关键发现": narration.get("key_findings") or [],
            "对比详情": diff_detail_payload,
            "风险": narration.get("risks") or [],
            "建议": narration.get("suggestions") or [],
        }
        chart_option = {"by_section": {"数据图表": pe.build_compare_chart_option(
            mode, diff_details, ds_l.name, ds_r.name,
        )}}

    indicators_json = json.dumps(indicators, ensure_ascii=False, default=str)
    report_content = json.dumps(
        {
            "title": name,
            "summary": narration.get("summary") or "",
            "sections": sections,
            # 【dataset-contract-compare-report 新增】把 mode 写进 report_content,
            # 前端 ReportView 拿到后可直接判断"这是合同报告",
            # 避免依赖解析 task.name(命名格式会变)
            "mode": mode,
        },
        ensure_ascii=False,
    )
    chart_option_json = json.dumps(chart_option, ensure_ascii=False)

    task = Task(
        name=name,
        dataset_id=ds_l.id,
        template_id=tpl_id,
        status="success",
        indicators_json=indicators_json,
        report_content=report_content,
        chart_option_json=chart_option_json,
        annotations_json="[]",
        error_message="",
    )
    db.add(task)
    db.commit()
    db.refresh(task)
    result["task_id"] = task.id
    result["task_url"] = f"/reports/{task.id}"
    return result


def _render_issues_to_text(issues: list) -> str:
    """把 issues[] 渲染成"【严重度】字段: 说明"多行文本。"""
    if not issues:
        return "无明显问题"
    sev_label = {"low": "轻微", "medium": "中等", "high": "严重"}
    lines: list[str] = []
    for iss in issues:
        if not isinstance(iss, dict):
            continue
        sev = iss.get("severity") or "low"
        sev_cn = sev_label.get(sev, sev)
        fld = iss.get("field") or "—"
        snippet = iss.get("snippet") or ""
        lines.append(f"【{sev_cn}】{fld}:{snippet}")
    return "\n".join(lines) if lines else "无明显问题"


def _render_rows_to_markdown(columns: list, rows: list) -> str:
    """把 columns + rows 渲染成 markdown 表格。空 rows → "无结果"。"""
    if not rows:
        return "无结果"
    cols = list(columns) if columns else []
    if not cols:
        # 没有列名,只把行数写出来(避免 markdown 错位)
        return f"共 {len(rows)} 行(无列名)"

    # 表头
    header = "| " + " | ".join(str(c) for c in cols) + " |"
    sep = "| " + " | ".join("---" for _ in cols) + " |"
    # 数据行
    body_lines: list[str] = []
    for r in rows[:200]:  # 防止表格过大,前 200 行
        if isinstance(r, list):
            cells = r
        elif isinstance(r, dict):
            cells = [r.get(c, "") for c in cols]
        else:
            cells = [str(r)]
        # 补齐/截断列数
        if len(cells) < len(cols):
            cells = list(cells) + [""] * (len(cols) - len(cells))
        else:
            cells = list(cells[: len(cols)])
        # 把 None / 复杂对象序列化
        from app.core.json_utils import json_default
        cells = [
            ("" if c is None else (c if isinstance(c, (str, int, float)) else json.dumps(c, ensure_ascii=False, default=json_default)))
            for c in cells
        ]
        # 转 str + 替换换行
        cells = [str(c).replace("\n", " ").replace("|", "\\|") for c in cells]
        body_lines.append("| " + " | ".join(cells) + " |")
    return "\n".join([header, sep, *body_lines])


def _summarize_diff_details_to_text(diff_details, mode: str) -> Any:
    """按 mode 把 diff_details 折成"对比详情"章节的内容。

    返回值类型:
        - row_diff: dict(结构化明细,前端按 added/removed/modified 分别渲染表格)
        - 其他 mode: 仍是纯字符串,与旧报告兼容

    ## 【row-diff-detail-expansion 新增】返回值类型变更说明
    原实现 row_diff 模式只回 "新增 N 条,删除 N 条,修改 N 条" 一行汇总,
    落到报告页"对比详情"section 后,用户只能看到计数,看不到具体的行。
    现改为结构化 dict:
        {
            "summary":           "新增 N 条,删除 N 条,修改 N 条(...)",
            "added":             [ {<key>: <val>, ...}, ... ]   # 新增行明细
            "removed":           [ {<key>: <val>, ...}, ... ]   # 删除行明细
            "modified":          [ {key, field_diffs:[{field, before, after}]} ]
            "modified_total":    int,                          # 真实修改总数
            "truncated_modified": bool,                        # modified 是否被截断
            "added_total":       int,                          # 新增总数
            "removed_total":     int,                          # 删除总数
        }
    前端 ReportView 在渲染该章节时,值是 dict 就走结构化分支;
    other mode 字符串值仍按原 white-space: pre-wrap 段落渲染,向后兼容。

    ## 【topn 并入 business 新增】
    business 模式在指定 top_n 时,明细行会带 left_rank / right_rank,
    这里把它们追加到文本行尾作为**辅助参考**;名次变化只是相对位移,
    不代表该分组自身经营涨跌(判断增减请看「差值 / 变化幅度」)。
    """
    if diff_details is None:
        return "无对比明细"
    try:
        if mode == "business":
            items = diff_details if isinstance(diff_details, list) else []
            if not items:
                return "无对比明细"
            lines = [f"共 {len(items)} 个分组。"]
            for d in items[:20]:
                grp = d.get("group")
                ls = d.get("left_sum")
                rs = d.get("right_sum")
                gp = d.get("growth_pct")
                # 【topn 并入 business 新增】有 top_n 时附带名次(辅助参考)
                lr = d.get("left_rank")
                rr = d.get("right_rank")
                rank_txt = ""
                if lr and rr:
                    rank_txt = f",左名次={lr},右名次={rr}"
                lines.append(
                    f"- {grp}:左={ls},右={rs},变化幅度={gp if gp is not None else '—'}%{rank_txt}"
                )
            if len(items) > 20:
                lines.append(f"  ……还有 {len(items) - 20} 组未展开")
            return "\n".join(lines)

        if mode == "row_diff":
            # 【row-diff-detail-expansion 新增】row_diff 走结构化明细,
            # 报告页"对比详情"section 会按 added/removed/modified 分块渲染表格
            d = diff_details if isinstance(diff_details, dict) else {}
            added_list = d.get("added") or []
            removed_list = d.get("removed") or []
            modified_list = d.get("modified") or []
            modified_total = int(d.get("modified_total") or len(modified_list))
            shown_modified = len(modified_list)

            summary_text = (
                f"新增 {len(added_list)} 条,删除 {len(removed_list)} 条,"
                f"修改 {modified_total} 条(已展示前 {shown_modified} 条)"
            )

            # 兜底:added/removed 不是 list[dict] 时降级为 []
            # 【row-diff-detail-expansion 新增】识别 row_level_diff 的形状
            # ({key, row});把 row 展开为顶层字段,前端可按列渲染表格。
            def _safe_rows(rows: Any) -> list[dict]:
                if not isinstance(rows, list):
                    return []
                out: list[dict] = []
                for r in rows:
                    if not isinstance(r, dict):
                        continue
                    row_inner = r.get("row")
                    if isinstance(row_inner, dict):
                        # 标准形状:{key: ..., row: {<字段>: ...}}</>,
                        # 把 row 内的字段展开,去掉内部的 _sheet_name 这种系统列,
                        # 前端按字段名渲染表格更直观
                        flat: dict = {}
                        for k, v in row_inner.items():
                            if str(k).startswith("_"):
                                # 跳过 _sheet_name / __k__ 这类系统列
                                continue
                            if v is None or isinstance(v, (str, int, float, bool)):
                                flat[k] = v
                            else:
                                try:
                                    flat[k] = str(v)
                                except Exception:
                                    flat[k] = "—"
                        # 主键列也加上,便于表格里看
                        if "key" in r and r["key"] not in flat:
                            flat["主键"] = r["key"]
                        out.append(flat)
                    else:
                        # 旧形状或退化形状:{<字段名>: ...} 直接展开
                        flat2: dict = {}
                        for k, v in r.items():
                            if str(k).startswith("_"):
                                continue
                            if v is None or isinstance(v, (str, int, float, bool)):
                                flat2[k] = v
                            else:
                                try:
                                    flat2[k] = str(v)
                                except Exception:
                                    flat2[k] = "—"
                        out.append(flat2)
                return out

            added_clean = _safe_rows(added_list)
            removed_clean = _safe_rows(removed_list)

            # modified 形状: {key: <主键值>, field_changes:[{field, before, after}]}
            # (后端 pandas_engine 的字段名是 field_changes,前端用 field_diffs 也兼容)
            modified_clean: list[dict] = []
            for m in modified_list:
                if not isinstance(m, dict):
                    continue
                fd_raw = m.get("field_changes") or m.get("field_diffs") or []
                fd: list[dict] = []
                for f in fd_raw:
                    if isinstance(f, dict):
                        fd.append({
                            "field":  f.get("field", ""),
                            "before": f.get("before", ""),
                            "after":  f.get("after", ""),
                        })
                modified_clean.append({
                    "key": m.get("key"),
                    "field_diffs": fd,
                })

            return {
                "summary":            summary_text,
                "added":              added_clean,
                "removed":            removed_clean,
                "modified":           modified_clean,
                "added_total":        len(added_clean),
                "removed_total":      len(removed_clean),
                "modified_total":     modified_total,
                "shown_modified":     shown_modified,
                "truncated_modified": modified_total > shown_modified,
            }

    except Exception:
        return "对比详情渲染失败"
    return "无对比明细"


def _parse_persist_flag(payload: dict, *, allow_name: bool = False) -> tuple[bool, str | None]:
    """解析 payload 里的 persist 与可选 name。

    Returns:
        (persist_bool, name_str_or_None)
        persist 缺省 / 非 bool → (False, None) 或 (False, name);非 bool 显式给 → 400
        这里采用宽松策略:缺省视为 False;非 bool 字符串时尝试解析,失败返回 400。

    说明:
        - allow_name=True 时,解析 payload["name"](数据校验 / 数据对话支持);
          allow_name=False 时,忽略 payload["name"](对比场景下对比 spec 不需要该字段)。
    """
    raw = payload.get("persist", None)
    if raw is None:
        persist = False
    elif isinstance(raw, bool):
        persist = raw
    elif isinstance(raw, str):
        s = raw.strip().lower()
        if s in ("true", "1", "yes", "on"):
            persist = True
        elif s in ("false", "0", "no", "off", ""):
            persist = False
        else:
            raise HTTPException(400, f"persist 必须是 bool,当前:{raw!r}")
    else:
        raise HTTPException(400, f"persist 必须是 bool,当前:{raw!r}")

    name: str | None = None
    if allow_name:
        nv = payload.get("name")
        if isinstance(nv, str):
            name = nv.strip() or None
    return persist, name


# ---------------------------------------------------------------------------
# 工具函数(供 /clean 多 sheet 路径复用):推断哪些列在清洗后被「日期化」
# ---------------------------------------------------------------------------
def _compute_date_unified(
    original_df: "pd.DataFrame",
    cleaned_df: "pd.DataFrame",
    out: list[str],
    pd_mod,
) -> None:
    """对比清洗前后 dtype / 可解析度,把"被统一为日期"的列名追加到 out。

    设计:
        - 复用原 /clean 路径的两条判定规则:(1) 原非日期 → 清洗后变 datetime;
          (2) 原 object/string → 清洗后可解析数显著提升。
        - 各 sheet 累加时同一列名会重复出现,这里去重交给调用方(out 用 list
          配 + 上层做 set 检查即可)。
        - `pd_mod` 传入 pandas 模块,避免函数内反复 import 影响性能。
    """
    common_cols = [c for c in original_df.columns if c in cleaned_df.columns]
    for c in common_cols:
        try:
            before_dt = pd_mod.api.types.is_datetime64_any_dtype(original_df[c])
            after_dt = pd_mod.api.types.is_datetime64_any_dtype(cleaned_df[c])
            if (not before_dt) and after_dt:
                if c not in out:
                    out.append(str(c))
                continue
            if pd_mod.api.types.is_object_dtype(original_df[c]) or pd_mod.api.types.is_string_dtype(original_df[c]):
                parsed = pd_mod.to_datetime(cleaned_df[c], errors="coerce")
                if parsed.notna().sum() >= max(1, int(len(cleaned_df) * 0.5)):
                    before_parsed = pd_mod.to_datetime(original_df[c], errors="coerce")
                    if before_parsed.notna().sum() < parsed.notna().sum():
                        if c not in out:
                            out.append(str(c))
        except Exception:
            continue


@router.post("/merge")
def merge(payload: dict, db: Session = Depends(get_db)):
    """多表合并：支持两种模式，按 mode 字段路由。

    请求体:
        {
            "dataset_ids": [int, ...],   # 至少 2 个
            "name":        str,          # 新数据集名,缺省 "合并数据集"
            "mode":        "flat" | "multi_sheet",  # 缺省 "flat",向后兼容
        }

    模式说明:
        - flat        :把所有数据集的 DataFrame 纵向拼接为一整张表,
                       写入单个 .xlsx 的默认 sheet(行为与历史完全一致)。
        - multi_sheet :每个数据集对应输出 .xlsx 里的 1 个 sheet,
                       sheet 名 = 数据集名(>31 字符截断 / 重名自动加 _2/_3 后缀,
                       兼容 Excel 限制)。

    响应:
        {
            "id":          新 dataset.id,
            "row_count":   总行数(multi_sheet 时为各 sheet 行数之和),
            "sheet_count": 实际写入的 sheet 数(flat=1, multi_sheet=K),
            "mode":        实际采用的模式
        }
    """
    dataset_ids = payload.get("dataset_ids") or []
    new_name = (payload.get("name") or "合并数据集").strip()
    # mode 缺省视为 "flat",对老客户端完全透明
    mode = (payload.get("mode") or "flat").strip()
    if mode not in ("flat", "multi_sheet"):
        raise HTTPException(400, "mode 必须是 flat 或 multi_sheet")
    if len(dataset_ids) < 2:
        raise HTTPException(400, "至少选择 2 个数据集")

    # 读取所有数据集的 DataFrame + 源数据集名(用于 sheet 名)
    frames: list = []
    src_names: list[str] = []
    for did in dataset_ids:
        ds = db.get(Dataset, did)
        if not ds:
            continue
        try:
            frames.append(pe.read_dataframe(ds.file_path))
            src_names.append(ds.name or f"dataset_{did}")
        except Exception as exc:
            raise HTTPException(400, f"读取数据集 {did} 失败: {exc}")

    if not frames:
        raise HTTPException(400, "所选数据集均不可读,请检查文件")

    save_path = settings.upload_dir / f"merge_{Path(new_name).stem}.xlsx"

    if mode == "flat":
        # —— 旧行为:纵向拼接为一整张表,写一个默认 sheet ——
        merged = pd_concat(frames)
        merged.to_excel(save_path, index=False)
        row_count = int(len(merged))
        sheet_count = 1
    else:
        # —— 新行为:每个数据集对应 1 个 sheet,所有 sheet 写入同一份 xlsx ——
        import re as _re

        def _safe_sheet_name(raw: str, used: set[str]) -> str:
            """生成合法且不重复的 sheet 名:
            - 去路径分隔符 / 冒号 / 方括号等 Excel 非法字符
            - 截断到 31 字符
            - 与已用名重复时追加 _2 / _3 ...
            """
            # 1) 把非法字符替换为下划线
            cleaned = _re.sub(r"[\\/\:\*\?\[\]]", "_", str(raw)).strip() or "sheet"
            # 2) Excel 单 sheet 名最长 31 字符
            cleaned = cleaned[:31]
            # 3) 处理重名
            base = cleaned
            n = 2
            while cleaned in used:
                suffix = f"_{n}"
                cleaned = (base[: 31 - len(suffix)] + suffix)
                n += 1
            used.add(cleaned)
            return cleaned

        used_names: set[str] = set()
        with pd_excel_writer(save_path) as writer:
            for df, name in zip(frames, src_names):
                sheet = _safe_sheet_name(name, used_names)
                df.to_excel(writer, sheet_name=sheet, index=False)
        # multi_sheet 模式下,row_count 取所有 sheet 行数之和(便于前端展示)
        row_count = int(sum(len(df) for df in frames))
        sheet_count = len(frames)

    ds = Dataset(
        name=new_name,
        file_name=save_path.name,
        file_path=str(save_path),
        row_count=row_count,
        status="已清洗",
    )
    db.add(ds)
    db.commit()
    db.refresh(ds)
    return {
        "id": ds.id,
        "row_count": ds.row_count,
        "sheet_count": sheet_count,
        "mode": mode,
    }


def pd_concat(frames):
    import pandas as pd
    return pd.concat(frames, ignore_index=True, sort=False)


def pd_excel_writer(path):
    """封装 pd.ExcelWriter,便于测试时 mock。"""
    import pandas as pd
    return pd.ExcelWriter(path, engine="openpyxl")


@router.post("/validate")
def validate(payload: dict, db: Session = Depends(get_db)):
    """数据校验:按 sheet 分别产出基础质量报告 + AI 自然语言解读。

    支持两种调用方式:
      1. 传 dataset_id: 后端自动从数据库查 file_path(推荐)
      2. 传 file_path: 直接使用传入路径(向后兼容)

    改版要点(多 sheet 不再联动成一张大表):
      - 旧实现用 read_dataframe_all_sheets 走"列名并集 + 缺失列补 NaN"生成合并大表,
        这会让每张 sheet 自己的字段空值被叠加统计,5 sheet 文件动辄上万"假缺失值"。
      - 现在改用 read_dataframe_per_sheet 逐 sheet 独立分析,各自产出质量报告。
      - 顶层 quality/column_types 改为按 sheet 加权汇总(行数加权的列缺失率,
        而非错误地把跨表 NaN 也算进来),与各 sheet 报告保持一致。
      - sheet_reports 列表按 sheet 顺序输出,前端可下钻查看每张表真实情况。
      - raw_preview 取自第一张 sheet(避免拉一张"由 NaN 列堆起来的怪表")。

    返回结构:
        {
            "quality":        { row_count, column_count, missing_total, ... },  # 加权汇总
            "column_types":   { 字段名: 类型 },                                  # 各 sheet 合并,同名取前者
            "ai_interpret":   { summary, level, ... },                          # 基于汇总判断的整体解读
            "raw_preview":    { columns, rows, total },                          # 第一张 sheet 的前 10 行
            "sheets":         [{ name, rows, cols }, ... ],                      # 全表清单
            "active_sheet_count": int,
            "sheet_reports":  [                                               # 多 sheet 分项报告(2026-10 新增)
                {
                  "name": "销售明细",
                  "rows": 1220, "cols": 13,
                  "quality": {...},
                  "column_types": {...},
                  "ai_interpret": {...},
                  "raw_preview": {...}
                },
                ...
            ],
        }
    """
    dataset_id = payload.get("dataset_id")
    file_path = payload.get("file_path")
    dataset_name: str = ""

    # 优先用 dataset_id 解析路径
    if dataset_id:
        ds = db.get(Dataset, int(dataset_id))
        if not ds:
            raise HTTPException(404, "数据集不存在")
        file_path = ds.file_path
        dataset_name = ds.name

    if not file_path or not Path(file_path).exists():
        raise HTTPException(400, "file_path 无效")

    # 【多 sheet 分项分析 2026-10】按 sheet 独立读取,各自计算 basic_quality + 字段类型
    try:
        per_sheets = pe.read_dataframe_per_sheet(file_path)
    except Exception as exc:
        raise HTTPException(400, f"读取失败: {exc}")

    # sheets 明细(与 /preview 接口对齐):[{name, rows, cols}, ...]
    sheets = [
        {"name": s["name"], "rows": int(len(s["df"])), "cols": int(len(s["df"].columns))}
        for s in per_sheets
    ]

    # —— 顶层汇总:行数加权 ——
    # row_count / column_count / missing_total 都按 sheet 简单求和;
    # duplicate_rows 同样求和(各 sheet 内的重复,与跨 sheet 无关)。
    # 顶层"整体缺失率"= missing_total / row_count,与单 sheet 缺失率加权平均一致。
    overall_row_count = 0
    overall_col_count = 0
    overall_missing = 0
    overall_duplicate = 0
    overall_missing_by_col: dict[str, int] = {}
    # column_types 用"先到先得"原则:同名列在多 sheet 出现时,以前面 sheet 的判定为准;
    # 命名约定("订单号"等)应一致,这里只兜底。
    overall_column_types: dict[str, str] = {}

    sheet_reports: list[dict] = []
    for entry in per_sheets:
        s_name = entry["name"]
        s_df = entry["df"]
        # 跳过完全空的 sheet(避免 AI 对 0 行数据生成"无数据"报告)
        if len(s_df) == 0:
            sheet_reports.append({
                "name": s_name,
                "rows": 0, "cols": int(len(s_df.columns)),
                "quality": {
                    "row_count": 0,
                    "column_count": int(len(s_df.columns)),
                    "missing_total": 0,
                    "missing_by_column": {},
                    "duplicate_rows": 0,
                },
                "column_types": {},
                "ai_interpret": {
                    "summary": f"工作表「{s_name}」为空,无数据可校验",
                    "level": "good",
                    "highlights": [],
                    "issues": [],
                    "suggestions": [],
                    "fallback": True,
                    "error": "",
                },
                "raw_preview": pe.dataframe_to_rows(s_df, 10),
            })
            continue

        s_quality = pe.basic_quality(s_df)
        s_types = pe.detect_column_types(s_df)
        # 调用 AI 做该 sheet 的自然语言解读,提示词前缀带上 sheet 名,
        # 让 LLM 知道"这是文件里的第 N 张表",报告更有针对性
        s_ai = ai_service.interpret_validation(
            s_quality, s_types,
            dataset_name=f"{dataset_name} / 工作表「{s_name}」",
        )

        # 汇总
        overall_row_count += int(s_quality.get("row_count", 0))
        overall_col_count = max(overall_col_count, int(s_quality.get("column_count", 0)))
        overall_missing += int(s_quality.get("missing_total", 0))
        overall_duplicate += int(s_quality.get("duplicate_rows", 0))
        for col, miss in (s_quality.get("missing_by_column") or {}).items():
            overall_missing_by_col[col] = overall_missing_by_col.get(col, 0) + int(miss)
        for col, t in s_types.items():
            if col not in overall_column_types:
                overall_column_types[col] = t

        sheet_reports.append({
            "name": s_name,
            "rows": int(len(s_df)),
            "cols": int(len(s_df.columns)),
            "quality": s_quality,
            "column_types": s_types,
            "ai_interpret": s_ai,
            "raw_preview": pe.dataframe_to_rows(s_df, 10),
        })

    overall_quality = {
        "row_count": overall_row_count,
        "column_count": overall_col_count,
        "missing_total": overall_missing,
        "missing_by_column": overall_missing_by_col,
        "duplicate_rows": overall_duplicate,
    }
    # 顶层 ai_interpret 用同样的 prompt,但 dataset_name 表明"跨 N 张 sheet"
    ai_interpret = ai_service.interpret_validation(
        overall_quality, overall_column_types,
        dataset_name=f"{dataset_name} (共 {len(per_sheets)} 个工作表)",
    )

    # 顶层 raw_preview 取自第一张 sheet(避免拉一张"由 NaN 列堆起来的怪表")
    if per_sheets:
        raw_preview = pe.dataframe_to_rows(per_sheets[0]["df"], 10)
    elif sheet_reports:
        # 兜底:per_sheets 为空(极端),用第一份 sheet_report 的 raw_preview
        raw_preview = sheet_reports[0]["raw_preview"]
    else:
        raw_preview = {"columns": [], "rows": [], "total": 0}

    result = {
        "quality": overall_quality,
        "column_types": overall_column_types,
        "ai_interpret": ai_interpret,
        "raw_preview": raw_preview,
        "sheets": sheets,
        "active_sheet_count": len(sheets),
        # 【多 sheet 分项 2026-10】每张表各自的报告,前端按 tab 展示
        "sheet_reports": sheet_reports,
    }

    # 【tools-results-history-persistence 新增】可选落库
    persist, name_override = _parse_persist_flag(payload, allow_name=True)
    if persist:
        _persist_validation_task(db, ds, result, name_override)

    return result


@router.post("/nl_query")
def nl_query(payload: dict, db: Session = Depends(get_db)):
    """自然语言查询。

    改版要点(2026-10 多 sheet 分项):
      - 不再传 read_dataframe_all_sheets 的合并大表给 LLM(列并集会污染上下文),
        而是按 sheet 分别独立产 columns_meta + sheet 明细,让 LLM 知道有哪些 sheet 可选。
      - ai_service.execute_nl_query 内部也改为按 sheet 提供数据(详见该函数改造)。
    """
    dataset_id = payload.get("dataset_id")
    question = (payload.get("question") or "").strip()
    if not dataset_id or not question:
        raise HTTPException(400, "dataset_id 与 question 必填")
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")

    # 【多 sheet 分项 2026-10】按 sheet 独立提供 columns_meta,
    # 不再用合并大表的 columns(混合了 N 字段,LLM 易混淆)。
    # 顶层 columns_meta 用"每张 sheet 的列名+类型"拼接,key 加 "[sheet 名] " 前缀避免撞名。
    try:
        per_sheets = pe.read_dataframe_per_sheet(ds.file_path)
    except Exception as exc:
        raise HTTPException(400, f"读取失败: {exc}")
    columns_meta: list[dict] = []
    sheets_info: list[dict] = []
    for entry in per_sheets:
        s_name = entry["name"]
        s_df = entry["df"]
        sheets_info.append({"name": s_name, "rows": int(len(s_df)), "cols": int(len(s_df.columns))})
        s_types = pe.detect_column_types(s_df)
        for c, t in s_types.items():
            columns_meta.append({"name": f"[{s_name}] {c}" if s_name else c, "type": t})
    # 兜底:全空数据集时给空 columns
    if not columns_meta:
        columns_meta = []

    # 【deepen-data-chat-analysis】多轮对话历史清洗:
    # 只接受 role ∈ {user, ai} 且 content 非空的对象;单条截断 1000 字,
    # 整体截取最近 10 条(≈5 轮)——既支持指代消解,又防止 prompt 膨胀拖慢响应。
    raw_history = payload.get("history") or []
    history: list[dict] = []
    if isinstance(raw_history, list):
        for item in raw_history:
            if not isinstance(item, dict):
                continue
            role = str(item.get("role") or "").strip()
            if role not in ("user", "ai"):
                continue
            content = str(item.get("content") or "").strip()
            if not content:
                continue
            history.append({"role": role, "content": content[:1000]})
    history = history[-10:]

    result = ai_service.execute_nl_query(
        question, ds.file_path, columns_meta, sheets_info=sheets_info, history=history
    )

    # 【tools-results-history-persistence 新增】可选落库
    persist, name_override = _parse_persist_flag(payload, allow_name=True)
    if persist:
        _persist_nl_task(db, ds, result, question, name_override)

    return result


@router.post("/compare")
def compare(payload: dict, db: Session = Depends(get_db)):
    """数据集对比:支持 2 种表格 mode(业务对比 / 变更追踪)+ 2 种自动分流模式
    (合同对比 / 跨模态文档对比)。

    请求体:
        {
            "left_dataset_id":  int,  # 必填
            "right_dataset_id": int,  # 必填
            "mode": "business" | "row_diff",   # 缺省 business
            # 以下按 mode 条件必填:
            "group_col":  str,           # business
            "metric_col": str,           # business
            "key_cols":   [str, ...],    # row_diff(≥1)
            "top_n":      int | None,    # business 可选。指定正整数时只保留
                                        # 指标值最高的 N 个分组,并附带名次辅助列;
                                        # 不传 / 0 / None 时返回全部分组
        }

    响应:
        {
            "indicators":        dict,                 # 按 mode 不同的指标摘要
            "comment":           str,                  # v1 兼容字段
            "narration":         {summary, key_findings, risks, suggestions},
            "diff_mode":         "business" | ...,
            "diff_details":      list | dict,          # 详见 spec
            "left_dataset_name": str,
            "right_dataset_name":str,
            "left_sheet_count":  int,
            "right_sheet_count": int,
        }

    异常:
        400  参数缺失/非法/列不存在/主键缺失 → 中文 message
        404  数据集不存在
        500  仅在文件无法读取等不可恢复错误时返回
    """
    left = payload.get("left_dataset_id")
    right = payload.get("right_dataset_id")
    if not left or not right:
        raise HTTPException(400, "left_dataset_id 与 right_dataset_id 必填")
    ds_l = db.get(Dataset, left)
    ds_r = db.get(Dataset, right)
    if not ds_l or not ds_r:
        raise HTTPException(404, "数据集不存在")

    # —— 类型分流前置(2026-10 多格式扩展) ——
    # 1) 用户显式选了 contract_compare(合同垂直场景)→ 强制走 _compare_contracts
    #    即使两边都是 structured(合同如果是 Excel 也合理),不依赖 is_doc_compare
    # 2) 用户没选 + 任一边非 structured → 走 document_compare
    # 3) 否则 → 2 种 Pandas mode(business / row_diff)
    left_type = getattr(ds_l, "file_type", "structured")
    right_type = getattr(ds_r, "file_type", "structured")
    is_doc_compare = (left_type != "structured") or (right_type != "structured")

    # 用户请求的 mode(contract_compare 是新垂直场景)
    req_mode = (payload.get("mode") or "").strip()
    is_contract = req_mode == "contract_compare"

    # 默认值声明(Pandas 分支会重新赋值,document/contract 分支保持 None)
    group_col: str | None = None
    metric_col: str | None = None
    key_cols: list[str] = []
    # 【topn 并入 business 新增】None = 不截断(返回全部分组)
    top_n: int | None = None
    contract_focus: str = ""

    if is_contract:
        # 【dataset-contract-compare 新增】合同对比垂直场景
        # 合同往往左(去年)右(今年)同类型,但内容是长文本,适合 LLM 条款级分析
        mode = "contract_compare"
        contract_focus = (payload.get("focus") or "").strip()
        if getattr(ds_l, "extract_status", "success") == "failed":
            raise HTTPException(400, f"左数据集「{ds_l.name}」内容抽取失败,无法参与合同对比")
        if getattr(ds_r, "extract_status", "success") == "failed":
            raise HTTPException(400, f"右数据集「{ds_r.name}」内容抽取失败,无法参与合同对比")
    elif is_doc_compare:
        # 非结构化对比无需 mode 参数,固定 document_compare,前端无需选择
        mode = "document_compare"
        if getattr(ds_l, "extract_status", "success") == "failed":
            raise HTTPException(400, f"左数据集「{ds_l.name}」内容抽取失败,无法参与对比")
        if getattr(ds_r, "extract_status", "success") == "failed":
            raise HTTPException(400, f"右数据集「{ds_r.name}」内容抽取失败,无法参与对比")
    else:
        # —— 解析并校验 mode 必填字段 ——(只在双 structured 时校验)
        mode = (req_mode or "business").strip()
        # 【topn 并入 business / distribution 下线 新增】
        # 原先的 "topn" / "distribution" 独立模式已下线,不再接受这两个 mode。
        valid_modes = ("business", "row_diff")
        if mode not in valid_modes:
            raise HTTPException(400, f"mode 必须是 {valid_modes} 之一,当前:{mode}")

        group_col = payload.get("group_col") or None
        metric_col = payload.get("metric_col") or None
        key_cols_raw = payload.get("key_cols")
        if isinstance(key_cols_raw, list):
            key_cols = [str(x) for x in key_cols_raw if x not in (None, "")]
        elif isinstance(key_cols_raw, str) and key_cols_raw:
            key_cols = [c.strip() for c in key_cols_raw.split(",") if c.strip()]

        # top_n 是 business 的可选截断参数:
        #   不传 / 空 / 0 → None(返回全部分组,旧行为)
        #   正整数 N → 只保留指标值最高的 N 个分组
        top_n_raw = payload.get("top_n")
        if top_n_raw in (None, ""):
            top_n = None
        else:
            try:
                top_n = int(top_n_raw)
            except (TypeError, ValueError):
                raise HTTPException(400, f"top_n 必须是整数,当前:{top_n_raw}")
            if top_n == 0:
                top_n = None
            elif top_n < 0 or top_n > 100:
                raise HTTPException(400, "top_n 必须在 1~100 之间(不截断请留空)")

        if mode == "business":
            if not group_col or not metric_col:
                raise HTTPException(400, "业务对比需要 group_col + metric_col")
        elif mode == "row_diff":
            if not key_cols:
                raise HTTPException(400, "变更追踪至少选择 1 个主键列(key_cols)")

    # —— 调用 ai_service(算子内部异常由 ValueError 抛出 → 400) ——
    try:
        # 【tools-results-history-persistence 新增】落库开关透传
        persist, _ = _parse_persist_flag(payload, allow_name=False)
        result = ai_service.compare_datasets(
            ds_l.file_path,
            ds_r.file_path,
            mode=mode,
            group_col=group_col,
            metric_col=metric_col,
            key_cols=key_cols,
            top_n=top_n,
            persist=persist,
            # —— 多格式文件支持(2026-10 新增):透传类型与文本,让 ai_service 走 LLM 文档对比 ——
            left_file_type=left_type,
            right_file_type=right_type,
            left_extracted_text=ds_l.extracted_text or "",
            right_extracted_text=ds_r.extracted_text or "",
            left_name=ds_l.name,
            right_name=ds_r.name,
            # 【dataset-contract-compare 新增】合同对比场景的 focus 软提示
            contract_focus=contract_focus,
        )
    except ValueError as exc:
        # 列不存在 / 主键缺失 / topn 非法 等结构化校验错误
        raise HTTPException(400, f"对比参数无效:{exc}")
    except HTTPException:
        raise
    except Exception as exc:
        # 文件损坏、磁盘错误等不可恢复错误,保留 500
        raise HTTPException(500, f"对比失败:{exc}")

    # —— 附 dataset 名 + sheet 数(前端展示用,不参与计算) ——
    result["left_dataset_name"] = ds_l.name
    result["right_dataset_name"] = ds_r.name

    # 【tools-results-history-persistence 新增】可选落库(必须放在 dataset 名写回之后,
    # 工具函数内部不会读 left_dataset_name / right_dataset_name,但语义上更稳)
    if persist:
        _persist_compare_task(db, ds_l, ds_r, result, mode)

    return result


# ----------------------------------------------------------------------
# POST /api/tools/chart_edit
# ----------------------------------------------------------------------
# 工作台辅助工具「图表编辑」独立入口:
#   - 自然语言指令(如"换成折线图""配色换成蓝色""只显示 Top 5")
#   - 后端拼好"当前图表配置 + 数据集元信息 + 指令"喂给 LLM
#   - LLM 返回新的 ECharts option,前端立即渲染对比
# ----------------------------------------------------------------------
@router.post("/chart_edit")
def chart_edit(payload: dict, db: Session = Depends(get_db)):
    """工作台辅助工具·图表编辑入口。

    请求体 (JSON):
        {
            "dataset_id":   int,            # 必填
            "instruction":  str,            # 必填,自然语言指令
            "current_option": dict | null,  # 当前图表 option;为空则从零生成
            "current_type":  "line"|...,    # 当前图表类型,可选
            "title":         str            # 当前图表标题,可选
        }

    响应:
        {
            "type": "line",
            "title": "...",
            "option": { ...完整 ECharts option... },
            "explanation": "一句话说明",
            "fallback": bool,
            "error": str
        }
    """
    dataset_id = payload.get("dataset_id")
    instruction = (payload.get("instruction") or "").strip()

    if not dataset_id:
        raise HTTPException(400, "dataset_id 必填")
    if not instruction:
        raise HTTPException(400, "instruction 必填")

    ds = db.get(Dataset, int(dataset_id))
    if not ds:
        raise HTTPException(404, "数据集不存在")

    # 可选:task_id 与 chart_index 用于把新图表直接持久化回历史任务
    task_id = payload.get("task_id")
    chart_index = payload.get("chart_index")  # 0-based

    try:
        # 【多 sheet 分项 2026-10】改用 read_dataframe_per_sheet 读首张 sheet:
        # 不再用 read_dataframe_all_sheets(列并集会让 LLM 看到无关字段,
        # 改图配色等指令可能被无关系列干扰)。
        # 用户可以传 active_sheet 指定具体 sheet;不传则默认第一张。
        active_sheet = (payload.get("active_sheet") or "").strip()
        per_sheets = pe.read_dataframe_per_sheet(ds.file_path)
        sheets = [
            {"name": s["name"], "rows": int(len(s["df"])), "cols": int(len(s["df"].columns))}
            for s in per_sheets
        ]
        # 选定工作面
        if active_sheet:
            df = next((s["df"] for s in per_sheets if s["name"] == active_sheet), None)
        else:
            df = per_sheets[0]["df"] if per_sheets else None
        if df is None:
            df = per_sheets[0]["df"]
        types = pe.detect_column_types(df)
        # 【修复】把 sample_rows 完全 Python 化,避免时间列残留 pd.Timestamp。
        from app.core.json_utils import json_default
        head5 = df.head(5)
        sample_rows = json.loads(  # 走 json_default,顺手把所有 numpy/Timestamp 转纯 Python
            json.dumps(
                head5.astype(object).where(head5.notna(), "").values.tolist(),
                default=json_default,
            )
        )
        df_meta = {
            "columns": [{"name": c, "type": types[c]} for c in df.columns],
            "sample_rows": sample_rows,
            "title": payload.get("title") or f"{ds.name} 图表",
            # sheet 明细(2026-10):前端可下钻切换工作面
            "sheets": sheets,
            "active_sheet_count": len(sheets),
            "active_sheet": active_sheet or (sheets[0]["name"] if sheets else ""),
        }
        # current_option 兼容两种形态:① 直接 option dict;② {type, option}
        co = payload.get("current_option")
        if isinstance(co, dict) and "option" in co and "type" in co:
            current_option = co.get("option") or {}
        else:
            current_option = co or {}

        result = ai_service.chart_edit(
            instruction=instruction,
            current_option=current_option,
            df_meta=df_meta,
        )

        # 把新图表写回 task.chart_option_json(覆盖原 charts[index],不传 index 则删 default)
        if task_id:
            from app.models.task import Task
            task = db.get(Task, int(task_id))
            if task:
                import json as _json
                try:
                    stored = _json.loads(task.chart_option_json or "{}")
                except Exception:
                    stored = {}
                charts = stored.get("charts") or []
                new_item = {
                    "type": result.get("type", "line"),
                    "title": result.get("title") or df_meta["title"],
                    "option": result.get("option") or {},
                }
                if isinstance(chart_index, int) and 0 <= chart_index < len(charts):
                    charts[chart_index] = new_item
                elif charts:
                    charts[0] = new_item
                else:
                    charts.append(new_item)
                stored["charts"] = charts
                task.chart_option_json = _json.dumps(stored, ensure_ascii=False, default=str)
                db.commit()
                # 告诉前端已持久化,便于跳转到报告预览
                result["saved_to_task_id"] = int(task_id)

        return result
    except HTTPException:
        raise
    except Exception as exc:
        logger_msg = f"chart_edit 失败: {exc}"
        raise HTTPException(400, logger_msg)


# ----------------------------------------------------------------------
# POST /api/tools/chart_apply
# ----------------------------------------------------------------------
# 用户在图表编辑弹窗里点「应用到任务」时调用,把候选图显式落库到指定历史任务。
# 与 chart_edit 解耦:chart_edit 只生成候选图,chart_apply 才落库,避免自动覆盖。
# ----------------------------------------------------------------------
@router.post("/chart_apply")
def chart_apply(payload: dict, db: Session = Depends(get_db)):
    """把候选图表写入指定历史任务。

    请求体 (JSON):
        {
            "task_id":     int,                  # 必填
            "chart_index": int,                  # 选填;缺省/越界则覆盖 charts[0],空时 push
            "type":        str,                  # 必填,如 line/bar/pie
            "title":       str,                  # 选填
            "option":      dict                  # 必填,ECharts option
        }

    响应:
        {
            "task_id":      int,
            "chart_option": { "charts": [...] }   # 落库后的完整 chart_option
        }
    """
    task_id = payload.get("task_id")
    chart_index = payload.get("chart_index")
    chart_type = payload.get("type")
    chart_title = payload.get("title")
    chart_option = payload.get("option")

    if not task_id:
        raise HTTPException(400, "task_id 必填")
    if not chart_type:
        raise HTTPException(400, "type 必填")
    if not isinstance(chart_option, dict) or not chart_option:
        raise HTTPException(400, "option 必填且必须是非空对象")

    from app.models.task import Task
    task = db.get(Task, int(task_id))
    if not task:
        raise HTTPException(404, "任务不存在")

    try:
        try:
            existing = json.loads(task.chart_option_json or "{}")
        except Exception:
            existing = {}
        charts = existing.get("charts") or []
        new_item = {
            "type": chart_type,
            "title": chart_title or "图表",
            "option": chart_option,
        }
        if isinstance(chart_index, int) and 0 <= chart_index < len(charts):
            charts[chart_index] = new_item
        elif charts:
            charts[0] = new_item
        else:
            charts.append(new_item)
        existing["charts"] = charts

        # 【修复】同步更新 by_section 里"出现在 charts 中的章节"的对应项,
        # 这样报告页 ReportView 才能即时显示最新颜色/类型。
        # 为什么?ReportView 优先从 chart_option.by_section[<章节名>] 取图表;
        # 旧版 chart_apply 只写顶层 charts,导致 by_section 仍是 LLM 最初生成的
        # 黄色版本,前端报告显示"应用了紫但报告里还是黄"。
        # 现在:遍历 by_section,对每个标题等于新图表标题的项做同位置替换;
        # 对"数据图表"这种通用 key 的 by_section,如果 charts 里只有一张,
        # 也把它当作"通用图表"顺手替换(对齐多数报告的渲染路径)。
        section = existing.get("by_section") or {}
        if isinstance(section, dict) and section:
            for k, arr in section.items():
                if not isinstance(arr, list):
                    continue
                # 如果该章节的某项标题 == 新图表标题,优先同位置替换
                replaced = False
                for i, c in enumerate(arr):
                    if not isinstance(c, dict):
                        continue
                    old_title = (c.get("title") or "").strip()
                    new_title = (chart_title or "").strip()
                    if new_title and old_title == new_title:
                        arr[i] = new_item
                        replaced = True
                        break
                # 否则:该章节只有 1 张图时直接替换,保持"按章节统一"的视觉
                if not replaced and len(arr) == 1:
                    arr[0] = new_item
                    replaced = True
                # 否则:按 chart_index 同位置替换(若 chart_index 在范围内)
                if not replaced and isinstance(chart_index, int) and 0 <= chart_index < len(arr):
                    arr[chart_index] = new_item

        task.chart_option_json = json.dumps(existing, ensure_ascii=False, default=str)
        db.commit()
        return {"task_id": int(task_id), "chart_option": existing}
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(400, f"chart_apply 失败: {exc}")


# ----------------------------------------------------------------------
# GET /api/tools/schema
# POST /api/tools/suggest_columns
# 【clean-ai-derived-columns-pivot 新增】数据清洗弹窗的字段结构 / AI 推荐入口
# ----------------------------------------------------------------------

_SUGGEST_MAX_SHEETS = 8       # 送给 LLM 的 sheet 上限
_SUGGEST_MAX_COLUMNS = 40     # 每个 sheet 参与画像的列上限
_SUGGEST_MAX_SAMPLES = 5      # 每列最多取 5 个非空去重样本


def _sample_profile_value(value: Any) -> Any:
    """把样本值转成可 JSON 序列化的原生类型(np.int64 / Timestamp 等)。"""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return str(value)


@router.get("/schema")
def dataset_schema(dataset_id: int, db: Session = Depends(get_db)):
    """返回数据集所有 sheet 的字段结构(列名 / 类型 / 非空数 / 唯一值数)。

    只读接口:不调用 LLM、不写库、不创建文件。
    """
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")

    try:
        sheet_frames = pe.read_dataframe_per_sheet(ds.file_path)
    except Exception as exc:
        raise HTTPException(400, f"读取失败:{exc}")

    sheets = [
        {
            "name": str(entry["name"]),
            "row_count": int(len(entry["df"])),
            # column_profiles 内部复用 detect_column_types,列顺序与 Excel 一致
            "columns": pe.column_profiles(entry["df"]),
        }
        for entry in sheet_frames
    ]
    return {"dataset_id": dataset_id, "sheets": sheets}


@router.post("/suggest_columns")
def suggest_columns(payload: dict, db: Session = Depends(get_db)):
    """AI 推荐派生列。

    请求体: {"dataset_id": int}
    响应:   {"suggestions": [...], "fallback": bool}

    约束:
        - 只读数据集并构造采样画像,不落库、不创建 task;
        - 采样上限:sheet ≤ 8、每 sheet 列 ≤ 40、每列非空去重样本 ≤ 5;
        - LLM 失败/无结果时走规则兜底,仍返回 200。
    """
    if "dataset_id" not in payload or payload["dataset_id"] in (None, ""):
        raise HTTPException(400, "dataset_id 必填")
    try:
        dataset_id = int(payload["dataset_id"])
    except (TypeError, ValueError) as exc:
        raise HTTPException(400, f"dataset_id 必须是整数:{exc}")

    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")

    try:
        sheet_frames = pe.read_dataframe_per_sheet(ds.file_path)
    except Exception as exc:
        raise HTTPException(400, f"读取失败:{exc}")

    # 构造轻量画像:只带列名/类型/样本,避免把整表数据塞进 Prompt
    profiles: list[dict] = []
    for entry in sheet_frames[:_SUGGEST_MAX_SHEETS]:
        df = entry["df"]
        types = pe.detect_column_types(df)
        columns: list[dict] = []
        for col in list(df.columns)[:_SUGGEST_MAX_COLUMNS]:
            series = df[col]
            name = str(col)
            samples = [
                _sample_profile_value(v)
                for v in series.dropna().drop_duplicates()
                .head(_SUGGEST_MAX_SAMPLES).tolist()
            ]
            columns.append({
                "name": name,
                "type": types.get(name, "text"),
                "non_null": int(series.notna().sum()),
                "unique_count": int(series.dropna().nunique()),
                "samples": samples,
            })
        profiles.append({
            "sheet_name": str(entry["name"]),
            "row_count": int(len(df)),
            "columns": columns,
        })

    try:
        result = ai_service.suggest_derived_columns(profiles)
    except Exception as exc:
        # 推荐属于辅助能力,任何意外都降级为空建议,不让弹窗整体报错
        logger.warning("suggest_columns 推荐失败: %s", exc)
        result = {
            "suggestions": [],
            "pivot_suggestions": [],
            "fallback": True,
        }

    return {
        "suggestions": list(result.get("suggestions") or []),
        "pivot_suggestions": list(result.get("pivot_suggestions") or []),
        "fallback": bool(result.get("fallback", True)),
    }


# ----------------------------------------------------------------------
# POST /api/tools/clean
# ----------------------------------------------------------------------
# 工作台辅助工具「数据清洗」独立入口:
#   - 与 /workbench 「开始 AI 分析」主路径解耦,不会写 task 表、不调任何 LLM
#   - 复用 pandas_engine.read_dataframe / clean / fill_missing / basic_quality,不动算子
#   - 清洗结果以新 Dataset 行落到「数据集管理」,与上传数据集同表同列表
# 字段语义见 spec:
#   workbench-standalone-cleansing (6 项原始 cleaned_stats) +
#   workbench-cleansing-missing-data (新增 4 项 fill_missing 字段)
# ----------------------------------------------------------------------
@router.post("/clean")
def clean(payload: dict, db: Session = Depends(get_db)):
    """工作台辅助工具·数据清洗入口。

    请求体 (JSON):
        {
            "dataset_id": int,                       # 必填
            "cleaning_rules": {                      # 选填,缺省视为全 false
                "drop_empty_rows":  bool,
                "drop_duplicates":  bool,
                "unify_date":       bool,
                "mark_anomaly":     bool,
                "fill_missing":     bool,            # 【workbench-cleansing-missing-data 新增】
                "fill_missing_policy": {strategy, scope}  # 【同上新增】
            },
            "name": str                              # 选填,留空则用 <原名>_cleaned
        },
        "keep_formulas": bool,                   # 【clean-keep-formulas 新增】选填,
                                                 # True 时走 openpyxl 公式保留回写(仅 Excel 生效)
        "derived_columns": [                     # 【clean-ai-derived-columns-pivot 新增】选填
            {"sheet_name": str, "name": str, "expression": "`销量` * `单价`"}
        ],
        "pivot": {                               # 【clean-ai-derived-columns-pivot 新增】选填
            "sheet_name": str, "row_dim": str, "col_dim": str,
            "value": str, "agg": "sum/mean/count/max/min", "include_margin": bool
        }

    响应:
        {
            "id": 新 dataset.id,
            "name": 新 dataset.name,
            "file_name": 新 dataset.file_name,
            "row_count": 新 dataset.row_count,
            "cleaned_stats": {
                "original_rows":        int,
                "cleaned_rows":         int,
                "drop_empty_rows_count": int,
                "drop_duplicates_count": int,
                "anomaly_marked_count":  int,
                "date_unified_columns": [str, ...],
                # ——【workbench-cleansing-missing-data 新增】——
                "missing_filled_count":  int,
                "dropped_rows_by_missing": int,
                "fill_strategy_used":   str | None,
                "fill_scope_applied":   str | None,
                # ——【clean-ai-derived-columns-pivot 新增】——
                "derived_columns_created":  [{sheet_name, name}, ...],
                "derived_columns_warnings": [str, ...],
                "pivot_sheet_name": str | None,
                "pivot_rows":       int,
                "pivot_row_dim":    str | None,
                "pivot_value":      str | None,
            },
            "left_dataset_name": 原数据集名
        }
    """
    # ------------------------------------------------------------------
    # 1. 参数校验:dataset_id 必填;cleaning_rules 各项缺省视为 false
    # ------------------------------------------------------------------
    if "dataset_id" not in payload or payload["dataset_id"] in (None, ""):
        raise HTTPException(400, "dataset_id 必填")
    try:
        dataset_id = int(payload["dataset_id"])
    except (TypeError, ValueError) as exc:
        raise HTTPException(400, f"dataset_id 必须是整数:{exc}")

    # 解析清洗规则:6 项均为 bool,缺省/未知键视为 false
    raw_rules = payload.get("cleaning_rules") or {}
    if not isinstance(raw_rules, dict):
        raw_rules = {}
    rules = {
        "drop_empty_rows": bool(raw_rules.get("drop_empty_rows", False)),
        "drop_duplicates": bool(raw_rules.get("drop_duplicates", False)),
        "unify_date":      bool(raw_rules.get("unify_date", False)),
        "mark_anomaly":    bool(raw_rules.get("mark_anomaly", False)),
        # ——【workbench-cleansing-missing-data 新增】——
        "fill_missing":    bool(raw_rules.get("fill_missing", False)),
    }
    # fill_missing_policy 透传(仅在 fill_missing=True 时校验形状)
    if rules["fill_missing"]:
        policy = raw_rules.get("fill_missing_policy")
        if isinstance(policy, dict):
            rules["fill_missing_policy"] = policy
        # 缺省/None/非 dict 都由 pandas_engine 兜底为 {strategy: mean, scope: numeric}

    # ——【clean-keep-formulas】公式保留回写开关 ——
    # True 且源文件是 Excel 时,走 openpyxl 回写路径(保留公式/样式/多 sheet 布局);
    # CSV 无公式概念,自动降级为普通 pandas 落盘。
    keep_formulas = bool(payload.get("keep_formulas", False))
    keep_formulas_report = None  # 非 keep_formulas 路径下响应里返回 null

    # ——【clean-ai-derived-columns-pivot】解析派生列 ——
    # 结构:[{sheet_name, name, expression}];按 sheet 分组,便于逐 sheet 执行。
    raw_derived = payload.get("derived_columns")
    derived_by_sheet: dict[str, list[dict[str, str]]] = {}
    if raw_derived is not None:
        if not isinstance(raw_derived, list):
            raise HTTPException(400, "派生列配置不完整")
        for item in raw_derived:
            if not isinstance(item, dict):
                raise HTTPException(400, "派生列配置不完整")
            item_sheet = str(item.get("sheet_name") or "").strip()
            item_name = str(item.get("name") or "").strip()
            item_expr = str(item.get("expression") or "").strip()
            # 三个字段缺一不可;名称/公式的语法校验交给 pandas_engine 统一处理
            if not item_sheet or not item_name or not item_expr:
                raise HTTPException(400, "派生列配置不完整")
            derived_by_sheet.setdefault(item_sheet, []).append({
                "name": item_name,
                "expression": item_expr,
            })

    # ——【clean-ai-derived-columns-pivot】解析透视汇总配置(支持多张并行) ——
    # 同时兼容旧契约 `pivot: {单对象}` 与新契约 `pivots: [{...}, {...}]`。
    # 每张 pivot 的配置结构不变:{sheet_name, row_dim, col_dim, value, agg, include_margin}。
    raw_pivots: list[dict[str, Any]] = []
    raw_pivot_single = payload.get("pivot")
    raw_pivot_list = payload.get("pivots")
    if isinstance(raw_pivot_list, list) and raw_pivot_list:
        # 新契约:数组
        raw_pivots = [p for p in raw_pivot_list if isinstance(p, dict)]
    elif isinstance(raw_pivot_single, dict):
        # 旧契约:单对象,自动包一层向后兼容
        raw_pivots = [raw_pivot_single]

    pivots: list[dict[str, Any]] = []
    for idx, p in enumerate(raw_pivots):
        p_sheet = str(p.get("sheet_name") or "").strip()
        p_row = str(p.get("row_dim") or "").strip()
        p_col = str(p.get("col_dim") or "").strip()
        p_value = str(p.get("value") or "").strip()
        if not p_sheet or not p_row or not p_value:
            raise HTTPException(400, f"第 {idx + 1} 张透视表配置不完整")
        pivots.append({
            "sheet_name": p_sheet,
            "row_dim": p_row,
            "col_dim": p_col or None,
            "value": p_value,
            "agg": str(p.get("agg") or "sum").strip().lower(),
            "include_margin": bool(p.get("include_margin", False)),
        })

    # 校验 dataset 是否存在
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")

    # ——【clean-ai-derived-columns-pivot】keep_formulas 冲突前置拦截 ——
    # 必须发生在读文件/写文件/建 Dataset 之前,否则会产出半成品数据集。
    if keep_formulas and (derived_by_sheet or pivots):
        raise HTTPException(
            400,
            "保留公式回写模式下不支持派生列或透视汇总,请取消其中一项后重试",
        )

    # 派生列 / 透视统计的默认值(未启用时按此回退,保证响应结构稳定)
    derived_created: list[dict[str, str]] = []
    derived_warnings: list[str] = []
    # 【clean-derived-excel-formula 新增】实际写入 Excel 公式的单元格数(0 = 未启用/注入失败)
    derived_formula_cells = 0
    # 多张透视统计:每张一条 {sheet_name, row_dim, value, rows},空列表表示未启用
    pivot_sheets: list[dict[str, Any]] = []

    # ------------------------------------------------------------------
    # 2. 复用 pandas_engine 算子,执行清洗并采集差量统计
    #
    # 【改版要点:per_sheet 独立清洗】
    #   - 旧实现用 read_dataframe_all_sheets 先做"列名并集"纵向合并、再按
    #     `_sheet_name` 分组清洗。4 张列结构完全不同的 sheet(明细/目标/渠道/
    #     产品)会被缝成一张十几列、大半是 NaN 的"缝合怪",落盘后每个 sheet
    #     都拖着其它表的空列,用户看到的就是"清洗完数据很乱"。
    #   - 现改用 read_dataframe_per_sheet:每个 sheet 保持**自己的列结构**独立
    #     读取(内部走智能表头,标题行/空列不再污染列名),清洗后按原 sheet 名
    #     分别写回,各 sheet 列数与原文件一一对应。
    #   - 差量统计在每 sheet 内部累加,合计成整份数据集的总指标。
    # ------------------------------------------------------------------
    try:
        sheet_frames = pe.read_dataframe_per_sheet(ds.file_path)  # [{name, df}, ...]
        if not sheet_frames:
            raise ValueError("文件中没有可读取的工作表")

        # ——【clean-ai-derived-columns-pivot】工作表名必须真实存在 ——
        # 派生列 / 透视的 sheet_name 写错时,提前报 400,避免落盘后才发现配置无效。
        known_sheets = {str(entry["name"]) for entry in sheet_frames}
        for _sname in derived_by_sheet:
            if _sname not in known_sheets:
                raise ValueError(f"工作表不存在:{_sname}")
        for _piv in pivots:
            if _piv["sheet_name"] not in known_sheets:
                raise ValueError(f"工作表不存在:{_piv['sheet_name']}")

        original_rows = sum(int(len(entry["df"])) for entry in sheet_frames)
        # 原始 NaN 总数:用于后面计算 fill_missing 真实填充的单元格数
        original_na_total = sum(
            int(entry["df"].isna().sum().sum()) for entry in sheet_frames
        )

        # 累计差量统计(各 sheet 累加)
        cleaned_rows_total = 0
        cleaned_na_total = 0
        drop_empty_rows_count = 0
        drop_duplicates_count = 0
        anomaly_marked_count = 0
        date_unified_columns: list[str] = []
        # 用于保存各 sheet 的清洗结果,后续 ExcelWriter 多 sheet 落盘
        cleaned_sheets: dict[str, "pd.DataFrame"] = {}
        # 【clean-derived-excel-formula 新增】实际应用成功的派生列配置
        # ({原 sheet 名: [{name, expression}]}),导出后用于把静态值升级为活公式
        derived_specs_applied: dict[str, list[dict[str, str]]] = {}
        # 【clean-derived-excel-formula 新增】原 sheet 名 → 文件内实际 sheet 名
        written_sheet_names: dict[str, str] = {}

        import re as _re
        import pandas as _pd
        for entry in sheet_frames:
            sub_name = str(entry["name"])
            sub_original = entry["df"]

            # ——【clean-keep-formulas】布局型 sheet 原样保留 ——
            # "目标设定"这类键值面板(标签行+空行分隔+合并标题条)没有真表头,
            # smart_header 只能给出"列N"占位列名。它们不是数据表:
            # 若参与 drop_empty_rows/drop_duplicates,空行分隔被删、内容上移,
            # 写回时会撞上合并单元格(MergedCell 只读)导致值静默丢失、布局全毁,
            # 并把渠道分析/产品分析里 =目标设定!B11 这类跨表引用全部带歪
            # (表现为"产品目标"列全部变 "--")。
            # 因此 keep_formulas 模式下:此类 sheet 不清洗、按原值原位回写,
            # 等效"未改动",模板布局与公式引用 100% 保真。
            if keep_formulas:
                n_cols = len(sub_original.columns)
                n_placeholder = sum(
                    1 for c in sub_original.columns
                    if _re.fullmatch(r"列\d+", str(c))
                )
                if n_cols and n_placeholder / n_cols >= 0.5:
                    cleaned_sheets[sub_name] = sub_original.copy()
                    cleaned_rows_total += int(len(sub_original))
                    cleaned_na_total += int(sub_original.isna().sum().sum())
                    continue

            sub_cleaned = pe.clean(sub_original, rules)
            # 差量统计基于"清洗后、派生列前"的表:派生列/透视不应改变
            # missing_filled_count 等既有指标的语义。
            cleaned_rows_total += int(len(sub_cleaned))
            cleaned_na_total += int(sub_cleaned.isna().sum().sum())

            # ——【clean-ai-derived-columns-pivot】该 sheet 的派生列 ——
            # 执行顺序严格为:clean → 派生列 → 透视(透视在循环外统一做)。
            sub_specs = derived_by_sheet.pop(sub_name, [])
            if sub_specs:
                sub_cleaned, derived_report = pe.apply_derived_columns(
                    sub_cleaned, sub_specs
                )
                # 响应按 {sheet_name, name} 记录,便于前端 / 调用方定位是哪张表的新列
                derived_created.extend(
                    {"sheet_name": sub_name, "name": str(col_name)}
                    for col_name in (derived_report.get("created") or [])
                )
                derived_warnings.extend(derived_report.get("warnings") or [])
                # 【clean-derived-excel-formula 新增】记录表达式,导出后升级为活公式
                # (apply_derived_columns 失败会整体抛错,走到这里的都是已创建列)
                derived_specs_applied[sub_name] = [
                    {"name": str(s["name"]), "expression": str(s["expression"])}
                    for s in sub_specs
                ]

            cleaned_sheets[sub_name] = sub_cleaned

            # —— drop_empty_rows 差量(每 sheet 独立计) ——
            if rules["drop_empty_rows"]:
                after_drop_empty = int(sub_original.dropna(how="all").shape[0])
                drop_empty_rows_count += int(len(sub_original)) - after_drop_empty

            # —— drop_duplicates 差量 ——
            if rules["drop_duplicates"]:
                before_dup = int(sub_original.duplicated().sum())
                after_dup = int(sub_cleaned.duplicated().sum())
                drop_duplicates_count += before_dup - after_dup

            # —— anomaly_marked_count ——
            if rules["mark_anomaly"]:
                anomaly_cols = [c for c in sub_cleaned.columns if str(c).endswith("_is_outlier")]
                anomaly_marked_count += int(sum(int(sub_cleaned[c].sum()) for c in anomaly_cols))

            # —— date_unified_columns(各 sheet 累加去重) ——
            if rules["unify_date"]:
                _compute_date_unified(
                    sub_original, sub_cleaned, date_unified_columns, _pd
                )

        cleaned_rows = cleaned_rows_total

        # ——【workbench-cleansing-missing-data】缺失值填充差量统计 ——
        if rules["fill_missing"]:
            missing_filled_count = max(0, original_na_total - cleaned_na_total)
            strategy = rules.get("fill_missing_policy", {}).get("strategy", "mean")
            scope = rules.get("fill_missing_policy", {}).get("scope", "numeric")
            if strategy == "drop":
                # 逐 sheet 统计"因缺失被删除的行数":
                #   before = 原 df 中目标列含 NA 的行数
                #   after  = 清洗后 df 中目标列仍含 NA 的行数(drop 策略下应为 0)
                # 各 sheet 的目标列按其自身列结构确定,不再依赖合并大表。
                dropped_rows_by_missing = 0
                for entry in sheet_frames:
                    sub_original = entry["df"]
                    if scope == "numeric":
                        target_cols = list(sub_original.select_dtypes(include="number").columns)
                    else:
                        target_cols = list(sub_original.columns)
                    if not target_cols:
                        continue
                    before_na_rows = int(sub_original[target_cols].isna().any(axis=1).sum())
                    sub_cleaned = cleaned_sheets[str(entry["name"])]
                    usable = [c for c in target_cols if c in sub_cleaned.columns]
                    after_na_rows = int(sub_cleaned[usable].isna().any(axis=1).sum()) if usable else 0
                    dropped_rows_by_missing += max(0, before_na_rows - after_na_rows)
            else:
                dropped_rows_by_missing = 0
            fill_strategy_used = strategy
            fill_scope_applied = scope
        else:
            missing_filled_count = 0
            dropped_rows_by_missing = 0
            fill_strategy_used = None
            fill_scope_applied = None

        # ——【clean-ai-derived-columns-pivot】循环生成多张透视汇总表 ——
        # 每张透视的数据源都是该 sheet「清洗 + 派生列」后的结果;
        # 透视行不计入 cleaned_rows(仍只统计原始数据行)。
        # sheet 命名策略:
        #   - 单张透视:沿用原固定名「透视汇总」保持向后兼容;
        #   - 多张透视:自动命名为 `透视_<row_dim>_<value>[_<col_dim>]`,
        #     如「透视_区域_年度合计」「透视_区域_产品_年度合计」;
        #   - 自动去 Excel 非法字符,重名加 `_2/_3` 后缀。
        for _piv in pivots:
            p_sheet_src = _piv["sheet_name"]
            src_df = cleaned_sheets.get(p_sheet_src)
            if src_df is None:
                raise ValueError(f"工作表不存在:{p_sheet_src}")
            # 后端再校验一次 value/row_dim/col_dim 是否在 src_df 里:
            # 派生列已在同 sheet apply_derived_columns 后追加,所以 value 如果是派生列,
            # 此处一定存在;如果不是派生列,也要求原始列存在。
            for _need_col in (_piv["row_dim"], _piv["value"]):
                if _need_col not in src_df.columns:
                    raise ValueError(
                        f"工作表 {p_sheet_src} 中不存在列:{_need_col}"
                    )
            if _piv.get("col_dim") and _piv["col_dim"] not in src_df.columns:
                raise ValueError(
                    f"工作表 {p_sheet_src} 中不存在列:{_piv['col_dim']}"
                )

            pivot_df = pe.build_pivot_table(src_df, _piv)

            # --- sheet 名构造 ---
            _sanitize = lambda s: _re.sub(r"[\\/\:\*\?\[\]]", "_", str(s)).strip()[:20] or "透视"
            if len(pivots) == 1:
                _base = "透视汇总"
            else:
                _parts = ["透视", _sanitize(_piv["row_dim"]), _sanitize(_piv["value"])]
                if _piv.get("col_dim"):
                    _parts.insert(2, _sanitize(_piv["col_dim"]))
                _base = "_".join(_parts)[:31]  # Excel sheet 名最长 31 字
            _candidate = _base
            _suffix = 2
            while _candidate in cleaned_sheets:
                _candidate = f"{_base}_{_suffix}"
                _suffix += 1

            cleaned_sheets[_candidate] = pivot_df
            pivot_sheets.append({
                "sheet_name": _candidate,
                "row_dim": _piv["row_dim"],
                "value": _piv["value"],
                "rows": int(len(pivot_df)),
            })

        # ------------------------------------------------------------------
        # 3. 写出清洗结果 Excel + 落库 Dataset 行
        # 单 sheet(CSV / 单表 Excel)保持旧行为直接写默认 Sheet1;
        # 多 sheet 用 ExcelWriter 按原 sheet 名分别落盘,保留多 sheet 结构。
        # 注意:单 sheet + 透视 = 2 张 sheet,会自然落入下面的多 sheet分支。
        # ------------------------------------------------------------------
        from datetime import datetime
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_stem = Path(ds.file_name).stem or ds.name or "dataset"
        save_path = settings.upload_dir / f"clean_{ts}_{original_stem}.xlsx"
        settings.upload_dir.mkdir(parents=True, exist_ok=True)

        if len(cleaned_sheets) == 1 and not keep_formulas:
            # 单 sheet / CSV:旧行为,to_excel 写到 Sheet1
            only_sheet_name = next(iter(cleaned_sheets))
            cleaned_sheets[only_sheet_name].to_excel(save_path, index=False)
            # 【clean-derived-excel-formula 新增】单表分支文件内 sheet 名固定 Sheet1
            written_sheet_names = {only_sheet_name: "Sheet1"}
        elif keep_formulas:
            # ——【clean-keep-formulas】公式保留回写路径 ——
            # 用 openpyxl 打开原工作簿(公式视图),只把清洗后的值写回数据区,
            # 公式/样式/合并单元格/多 sheet 布局全部保留;输出文件名带
            # keepformula_ 前缀以便区分。
            save_path = settings.upload_dir / f"clean_keepformula_{ts}_{original_stem}.xlsx"
            keep_formulas_report = pe.write_cleaned_keep_formulas(
                ds.file_path, save_path, cleaned_sheets
            )
        else:
            # 多 sheet:ExcelWriter 按 sheet 名分别落盘
            import re as _re
            used_names: set[str] = set()

            def _safe_sheet_name(raw: str) -> str:
                """与 /tools/merge multi_sheet 同款的 sheet 名清洗逻辑,
                去 Excel 非法字符 + 截断 31 字 + 重名加 _2/_3 后缀。
                """
                cleaned = _re.sub(r"[\\/\:\*\?\[\]]", "_", str(raw)).strip() or "sheet"
                cleaned = cleaned[:31]
                base = cleaned
                n = 2
                while cleaned in used_names:
                    suffix = f"_{n}"
                    cleaned = (base[: 31 - len(suffix)] + suffix)
                    n += 1
                used_names.add(cleaned)
                return cleaned

            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                for sname, sdf in cleaned_sheets.items():
                    safe = _safe_sheet_name(sname)
                    sdf.to_excel(writer, sheet_name=safe, index=False)
                    # 【clean-derived-excel-formula 新增】记录原 sheet 名 → 文件内名
                    written_sheet_names[sname] = safe

        # ——【clean-derived-excel-formula 新增】派生列升级为 Excel 活公式 ——
        # 在已落盘的 xlsx 上把派生列单元格改写为「活公式 + 缓存值」双写形态:
        # Excel/WPS 打开能看到 =B2*C2 并联动重算;本工具/ pandas 仍读到缓存值。
        # keep_formulas 模式与派生列互斥(前置 400),此处天然不涉及。
        # 注入失败仅记 warning,不影响已生成的静态值文件。
        if derived_specs_applied:
            try:
                derived_formula_cells = pe.inject_derived_formulas(
                    save_path, written_sheet_names, cleaned_sheets,
                    derived_specs_applied,
                )
            except Exception as exc:
                logger.warning("派生列公式注入失败(保留静态值): %s", exc)
                derived_warnings.append(f"派生列公式写入失败,已保留静态值:{exc}")

        # —— name 兜底 + 唯一性 ——
        candidate = (payload.get("name") or "").strip()
        if not candidate:
            candidate = f"{ds.name}_cleaned"
        final_name = candidate
        if db.query(Dataset).filter(Dataset.name == final_name).first():
            n = 1
            while db.query(Dataset).filter(Dataset.name == f"{candidate}_{n}").first():
                n += 1
            final_name = f"{candidate}_{n}"

        new_ds = Dataset(
            name=final_name,
            file_name=save_path.name,
            file_path=str(save_path),
            row_count=cleaned_rows,
            status="已清洗",
        )
        db.add(new_ds)
        db.commit()
        db.refresh(new_ds)
    except HTTPException:
        raise
    except ValueError as exc:
        # pandas_engine 抛的策略非法 / 文本列不允许等校验错误 → 400
        raise HTTPException(400, str(exc))
    except Exception as exc:
        raise HTTPException(400, f"清洗失败:{exc}")

    # ------------------------------------------------------------------
    # 4. 组装返回:核心元数据 + 10 项 cleaned_stats + 左数据集名
    # ------------------------------------------------------------------
    cleaned_stats = {
        "original_rows": original_rows,
        "cleaned_rows": cleaned_rows,
        "drop_empty_rows_count": drop_empty_rows_count,
        "drop_duplicates_count": drop_duplicates_count,
        "anomaly_marked_count": anomaly_marked_count,
        "date_unified_columns": date_unified_columns,
        # ——【workbench-cleansing-missing-data 新增】——
        "missing_filled_count": missing_filled_count,
        "dropped_rows_by_missing": dropped_rows_by_missing,
        "fill_strategy_used": fill_strategy_used,
        "fill_scope_applied": fill_scope_applied,
        # ——【clean-ai-derived-columns-pivot 新增】——
        # 未启用派生列/透视时按 []/[] 回退,保证前端结构稳定。
        # 向后兼容:保留旧单字段(取第一张透视的值),新字段 pivot_sheets 给完整列表。
        "derived_columns_created": derived_created,
        "derived_columns_warnings": derived_warnings,
        # ——【clean-derived-excel-formula 新增】——
        # 导出文件中派生列已写入 Excel 活公式的单元格数;0 表示未启用或注入失败
        # (失败时已回退为静态值,详情见 derived_columns_warnings)
        "derived_formula_cells": derived_formula_cells,
        "pivot_sheet_name": pivot_sheets[0]["sheet_name"] if pivot_sheets else None,
        "pivot_rows": pivot_sheets[0]["rows"] if pivot_sheets else 0,
        "pivot_row_dim": pivot_sheets[0]["row_dim"] if pivot_sheets else None,
        "pivot_value": pivot_sheets[0]["value"] if pivot_sheets else None,
        "pivot_sheets": pivot_sheets,  # [{sheet_name, row_dim, value, rows}], 多张完整信息
    }
    return {
        "id": new_ds.id,
        "name": new_ds.name,
        "file_name": new_ds.file_name,
        "row_count": new_ds.row_count,
        "cleaned_stats": cleaned_stats,
        "left_dataset_name": ds.name,
        # ——【clean-keep-formulas 新增】公式保留回写报告(未启用时为 null) ——
        "keep_formulas_report": keep_formulas_report,
    }


# =============================================================================
# 【workbench-ai-pdf-tools 新增】AI PDF 处理 4 个端点
# 设计:
#   - 4 个 POST 端点路径统一 /api/tools/pdf/<action>
#   - 接收 multipart/form-data,统一校验后缀 .pdf 与 size ≤ 50MB
#   - 输出文件统一放 settings.upload_dir,文件名带 pdf<action>_<ts>_<原名> 前缀
#   - 单独提供 GET /api/tools/pdf/download/{file_name} 给前端下载,白名单前缀
#   - AI 增强(转表格的表头识别 / 提取文字的清洗)失败自动降级,不抛错
# =============================================================================

# 文件大小上限 50MB(模块级常量,避免在 config 里新增字段)
_MAX_PDF_SIZE = 50 * 1024 * 1024

# download 路由允许的文件前缀(防路径穿越)
_PDF_DOWNLOAD_PREFIXES = (
    pdf_service.PREFIX_TO_WORD,
    pdf_service.PREFIX_TO_EXCEL,
    pdf_service.PREFIX_EXTRACT,
    pdf_service.PREFIX_COMPRESS,
    # 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 的产物前缀。
    # 只增不减:原有 4 个前缀的下载行为完全不变。
    pdf_service.PREFIX_IMG2PDF,
)


def _check_pdf_upload(file: UploadFile, content: bytes) -> None:
    """统一校验:后缀必须是 .pdf、大小 ≤ 50MB。

    Raises:
        HTTPException 400 / 413
    """
    name = (file.filename or "").lower()
    if not name.endswith(".pdf"):
        raise HTTPException(400, "仅支持 PDF 文件")
    if len(content) > _MAX_PDF_SIZE:
        raise HTTPException(413, f"PDF 文件超过 {_MAX_PDF_SIZE // (1024 * 1024)}MB 上限")
    if len(content) == 0:
        raise HTTPException(400, "PDF 文件为空")


# ===== 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 的上传校验与形态常量 =====
# 图片类型白名单:与 file_extractor._IMAGE_MIME 保持一致的口径
_IMAGE_SUFFIXES = (".png", ".jpg", ".jpeg", ".webp", ".bmp")
# 单次上传张数上限(与 pdf_service._MAX_IMAGES 对齐,双保险)
_MAX_IMAGE_COUNT = 20
# 单张图片大小上限 20MB(手机照片 12MP 约 3~5MB,20MB 足够覆盖高像素机型)
_MAX_IMAGE_SIZE = 20 * 1024 * 1024

# 三种 PDF 形态 → 中文名(前端单选组、历史任务标题共用)
_IMG_PDF_MODES: dict[str, str] = {
    "dual": "双层 PDF",
    "text": "纯文字 PDF",
    "image": "图片型 PDF",
}


def _check_image_uploads(files: list[UploadFile]) -> None:
    """校验照片转 PDF 的上传文件列表。

    Raises:
        HTTPException 400(后缀/数量/空文件) / 413(单张超大)
    """
    if not files:
        raise HTTPException(400, "请至少上传一张图片")
    if len(files) > _MAX_IMAGE_COUNT:
        raise HTTPException(
            400, f"单次最多支持 {_MAX_IMAGE_COUNT} 张图片,当前 {len(files)} 张"
        )
    for i, f in enumerate(files):
        name = (f.filename or "").lower()
        if not name.endswith(_IMAGE_SUFFIXES):
            raise HTTPException(
                400,
                f"第 {i + 1} 个文件不是支持的图片格式"
                f"(仅支持 {'/'.join(s.lstrip('.') for s in _IMAGE_SUFFIXES)})",
            )
        if f.size is not None and f.size > _MAX_IMAGE_SIZE:
            raise HTTPException(
                413,
                f"第 {i + 1} 张图片超过 "
                f"{_MAX_IMAGE_SIZE // (1024 * 1024)}MB 上限",
            )


def _pdf_download_url(file_name: str) -> str:
    """下载 URL 助手。

    返回**纯相对路径** ``/tools/pdf/download/<file>``(不带 /api 前缀)。
    前端 axios 实例 baseURL 已经是 /api,会自动拼接;若返回带 /api 前缀的
    绝对路径,会被 axios 重复拼前缀导致 /api/api/... 404。
    """
    return f"/tools/pdf/download/{file_name}"


# =============================================================================
# 【pdf-tools-history-persistence 新增】PDF 工具结果落库为历史任务
#
# 设计要点:
#   1. PDF 工具没有"原始 dataset"概念(Task.dataset_id NOT NULL),
#      这里新建一个 status="PDF工具产物" 的 Dataset 行作为锚点,
#      file_path 指向原始 PDF 字节(临时落盘)。
#   2. 前端 /data-datasets 列表按 status 过滤掉"PDF工具产物",不会污染。
#   3. 4 种 PDF 功能各自对应一个内置模板(seed.py 已写入):
#      - "PDF 转 Word 报告"   / "PDF 转表格 报告"
#      - "PDF 提取文字 报告"  / "PDF 压缩 报告"
#   4. report_content.sections 完全由后端组装,不调 LLM。
#   5. 落库时若模板缺失 → 500(让用户感知;不会写出半成品 task)。
# =============================================================================

# 内置模板名 → 用于 /persist 时的 template_id 查找
_PDF_TPL_NAME: dict[str, str] = {
    "to_word":    "PDF 转 Word 报告",
    "to_excel":   "PDF 转表格 报告",
    "extract":    "PDF 提取文字 报告",
    "compress":   "PDF 压缩 报告",
    # 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 落库模板
    "image_to_pdf": "照片转 PDF 报告",
}

# 工具类型 → 前端"工具中文名"卡片文案
_PDF_TOOL_LABEL: dict[str, str] = {
    "to_word":    "PDF 转 Word",
    "to_excel":   "PDF 转表格",
    "extract":    "PDF 提取文字",
    "compress":   "PDF 压缩",
    # 【pdf-photo-to-pdf-ocr 新增】
    "image_to_pdf": "照片转 PDF",
}


def _ensure_pdf_anchor_dataset(
    db: Session,
    *,
    pdf_content: bytes,
    original_name: str,
) -> Dataset:
    """为 PDF 工具产物创建/获取锚点 Dataset。

    锚点行为:
        - 把 PDF 字节落盘到 settings.upload_dir/pdf_anchor_<ts>_<原名>.pdf
        - Dataset.file_path 指向这份 PDF
        - status=INTERNAL_DATASET_STATUS —— 内部锚点,列表接口按此标记过滤,
          不出现在「数据集」页与工作台的下拉里(见 models/dataset.py 常量说明)
        - name 用 stem_<ts> 命名,同名追加 _N

    同一份 PDF 在一次请求内复用 anchor 即可,不做 hash 去重
    (因为 PDF 工具本身是单次请求落库,不存在"同一文件多次保存"的诉求)。
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = Path(original_name or "pdf").stem
    safe_stem = re.sub(r"[\\/:\*\?\"<>\|]", "_", stem).strip() or "pdf"
    # 锚点文件名
    anchor_name = f"{safe_stem}_pdfanchor_{ts}.pdf"
    anchor_path = settings.upload_dir / anchor_name
    settings.upload_dir.mkdir(parents=True, exist_ok=True)
    anchor_path.write_bytes(pdf_content)

    # 锚点 dataset 名(用户列表里看不到,但要保证不冲突)
    ds_base = f"{safe_stem}_pdf_{ts}"
    ds_name = ds_base
    if db.query(Dataset).filter(Dataset.name == ds_name).first():
        n = 1
        while db.query(Dataset).filter(Dataset.name == f"{ds_base}_{n}").first():
            n += 1
        ds_name = f"{ds_base}_{n}"

    new_ds = Dataset(
        name=ds_name,
        file_name=anchor_name,
        file_path=str(anchor_path),
        row_count=0,
        status=INTERNAL_DATASET_STATUS,
    )
    db.add(new_ds)
    db.commit()
    db.refresh(new_ds)
    return new_ds


def _build_pdf_report_content(
    *,
    tool: str,
    anchor: Dataset,
    original_name: str,
    result: dict,
) -> dict[str, Any]:
    """组装 report_content(供 Task.report_content 落库 + 前端 ReportView 渲染)。

    返回 dict 形如:
        {
            "title": ...,
            "summary": ...,
            "sections": {
                "原文件信息": "...",
                "转换结果": "..." | 列表型数据,
                "下载产物": "...",
            },
            "indicators": { ... }
        }
    """
    orig_size = int(result.get("original_size", 0) or 0)
    out_size = int(result.get("output_size", 0) or 0)
    title = f"{Path(original_name).stem} · { _PDF_TOOL_LABEL[tool] }"
    sections: dict[str, Any] = {
        "原文件信息": (
            f"文件名:{original_name}\n"
            f"原始大小:{_format_size(orig_size)}\n"
            f"锚点数据集:{anchor.name} (#{anchor.id})"
        ),
    }
    summary = ""
    # 产物元信息(供前端 ReportView 渲染"下载产物"按钮)
    pdf_artifact: dict[str, Any] = {
        "tool": tool,
        "file_name": result.get("file_name"),
        # 注意:用纯相对路径,前端 axios baseURL='/api' 自动拼接
        "download_url": _pdf_download_url(result.get("file_name") or ""),
        "original_size": orig_size,
        "output_size": out_size,
    }
    if tool == "to_word":
        sections["转换结果"] = (
            f"输出文件名:{result.get('file_name')}\n"
            f"输出大小:{_format_size(out_size)}\n"
            f"压缩/放大率:{round((1 - out_size / orig_size) * 100, 2) if orig_size else 0}%"
        )
        summary = f"PDF 已转换为 Word 文档,产物大小 {_format_size(out_size)}。"
        pdf_artifact.update({
            "kind": "docx",
            "label": "下载 Word 文件",
            "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        })
    elif tool == "to_excel":
        sheets = result.get("sheets") or []
        sheet_lines = [
            f"- {s.get('sheet')}:{s.get('rows')} 行 × {s.get('cols')} 列"
            for s in sheets
        ]
        sections["识别到的表格"] = (
            f"共识别 {result.get('table_count', 0)} 个表格:\n"
            + "\n".join(sheet_lines) if sheet_lines else "未识别到表格"
        )
        sections["AI 辅助说明"] = (
            "已使用 LLM 辅助识别表头" if result.get("ai_used")
            else "未启用 AI 辅助,使用纯库结果"
        )
        sections["下载产物"] = (
            f"输出文件:{result.get('file_name')}\n"
            f"文件大小:{_format_size(out_size)}"
        )
        summary = f"已从 PDF 抽取 {result.get('table_count', 0)} 个表格,生成多 Sheet Excel。"
        pdf_artifact.update({
            "kind": "xlsx",
            "label": "下载 Excel 文件",
            "table_count": int(result.get("table_count", 0) or 0),
            "sheets": sheets,
            "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        })
    elif tool == "extract":
        page_count = int(result.get("page_count", 0) or 0)
        char_count = int(result.get("char_count", 0) or 0)
        sections["提取概览"] = (
            f"总页数:{page_count}\n"
            f"字符数:{char_count}\n"
            f"AI 清洗:{'失败,使用原文' if result.get('ai_fallback') else '已清洗'}"
        )
        # 全文塞到 sections["提取内容"] 里(可能较长)
        full_text = (result.get("text") or "").strip()
        # 单章节展示全部文本;前端可滚动
        sections["提取内容"] = full_text or "(空)"
        summary = f"从 {page_count} 页 PDF 中提取 {char_count} 字。"
        pdf_artifact.update({
            "kind": "txt",
            "label": "下载 TXT 文件",
            "page_count": page_count,
            "char_count": char_count,
            "ai_fallback": bool(result.get("ai_fallback")),
            "mime": "text/plain",
        })
    elif tool == "compress":
        ratio = float(result.get("compression_ratio", 0) or 0)
        sections["压缩结果"] = (
            f"原始大小:{_format_size(orig_size)}\n"
            f"输出大小:{_format_size(out_size)}\n"
            f"压缩比:{ratio}%\n"
            f"提示:{'已优化,体积变化不大' if ratio <= 0 else '已成功压缩'}"
        )
        summary = f"PDF 压缩完成,压缩比 {ratio}%。"
        pdf_artifact.update({
            "kind": "pdf",
            "label": "下载压缩后的 PDF",
            "compression_ratio": ratio,
            "mime": "application/pdf",
        })
    elif tool == "image_to_pdf":
        # 【pdf-photo-to-pdf-ocr 新增】照片转 PDF(OCR)
        page_count = int(result.get("page_count", 0) or 0)
        ocr_success = int(result.get("ocr_success", 0) or 0)
        ocr_failed = int(result.get("ocr_failed", 0) or 0)
        mode_label = _IMG_PDF_MODES.get(
            str(result.get("mode") or "dual"), "双层 PDF"
        )
        # 识别结果可能很长,截断后入库,避免 report_content 体积失控
        ocr_text = (result.get("ocr_text") or "").strip()
        text_preview = ocr_text[:4000] if ocr_text else "(未识别到文字)"
        if ocr_text and len(ocr_text) > 4000:
            text_preview += "\n……(内容过长已截断,完整内容请下载 PDF 查看)"

        sections["识别结果"] = (
            f"输出形态:{mode_label}\n"
            f"总页数:{page_count}\n"
            f"识别成功:{ocr_success} 张\n"
            f"识别失败:{ocr_failed} 张\n"
            f"识别到的文字:\n{text_preview}"
        )
        sections["产物下载"] = (
            f"输出文件:{result.get('file_name')}\n"
            f"文件大小:{_format_size(out_size)}"
        )
        summary = (
            f"已将 {page_count} 张照片合并为 {mode_label},"
            f"成功识别 {ocr_success} 张。"
        )
        pdf_artifact.update({
            "kind": "pdf",
            "label": "下载照片转出的 PDF",
            "page_count": page_count,
            "mode": str(result.get("mode") or "dual"),
            "mode_label": mode_label,
            "mime": "application/pdf",
        })

    return {
        "title": title,
        "summary": summary,
        "sections": sections,
        # 【pdf-tools-history-persistence 新增】结构化产物信息,
        # 前端 ReportView 识别到 pdf_artifact 字段就显示"下载产物"按钮
        "pdf_artifact": pdf_artifact,
    }


def _persist_pdf_task(
    db: Session,
    *,
    tool: str,                   # to_word / to_excel / extract / compress
    pdf_content: bytes,
    original_name: str,
    result: dict,
    name_override: str | None = None,
) -> dict[str, Any]:
    """PDF 工具产物落库为历史任务。

    流程:
        1. 查内置模板(无 → 500)
        2. 建/获取锚点 Dataset(status=PDF工具产物)
        3. 组装 report_content.sections
        4. 新建 Task 并写库
        5. 返回 { task_id, task_url } 供前端弹底部"已保存"卡片

    Returns:
        { task_id, task_url, dataset_id }
    """
    tpl_name = _PDF_TPL_NAME.get(tool)
    if not tpl_name:
        raise HTTPException(500, f"未知的 PDF 工具类型:{tool}")
    tpl = db.query(ReportTemplate).filter(ReportTemplate.name == tpl_name).first()
    if not tpl:
        raise HTTPException(
            500,
            f"{tpl_name} 模板未初始化,请重启后端(seed.py 启动幂等会补齐)",
        )

    anchor = _ensure_pdf_anchor_dataset(
        db, pdf_content=pdf_content, original_name=original_name,
    )

    content = _build_pdf_report_content(
        tool=tool,
        anchor=anchor,
        original_name=original_name,
        result=result,
    )

    # name 入参优先;否则拼默认
    name = (name_override or "").strip()
    if not name:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        name = f"{Path(original_name).stem} · { _PDF_TOOL_LABEL[tool] }({ts})"
    name = name[:200]  # Task.name 上限

    # chart_option_json 留空(纯工具报告无图,ReportView 不会因空图报错)
    chart_option_json = "{}"
    # indicators_json:把元信息塞进去,便于 ReportView 摘要展示
    indicators = {
        "tool": tool,
        "original_name": original_name,
        "original_size": int(result.get("original_size", 0) or 0),
        "output_size": int(result.get("output_size", 0) or 0),
        "output_file": result.get("file_name"),
    }
    if tool == "extract":
        indicators["page_count"] = int(result.get("page_count", 0) or 0)
        indicators["char_count"] = int(result.get("char_count", 0) or 0)
        indicators["ai_fallback"] = bool(result.get("ai_fallback"))
    elif tool == "to_excel":
        indicators["table_count"] = int(result.get("table_count", 0) or 0)
        indicators["ai_used"] = bool(result.get("ai_used"))
    elif tool == "compress":
        indicators["compression_ratio"] = float(result.get("compression_ratio", 0) or 0)
    elif tool == "image_to_pdf":
        # 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 的统计指标
        indicators["page_count"] = int(result.get("page_count", 0) or 0)
        indicators["image_count"] = int(result.get("page_count", 0) or 0)
        indicators["ocr_success"] = int(result.get("ocr_success", 0) or 0)
        indicators["ocr_failed"] = int(result.get("ocr_failed", 0) or 0)
        indicators["mode"] = str(result.get("mode") or "dual")

    task = Task(
        name=name,
        dataset_id=anchor.id,
        template_id=tpl.id,
        status="success",
        indicators_json=json.dumps(indicators, ensure_ascii=False, default=str),
        report_content=json.dumps(content, ensure_ascii=False, default=str),
        chart_option_json=chart_option_json,
        annotations_json="[]",
        use_knowledge_base=False,
        kb_ids="",
        error_message="",
        parent_task_id=None,
        feedback_text="",
    )
    db.add(task)
    db.commit()
    db.refresh(task)

    return {
        "task_id": int(task.id),
        "task_url": f"/reports/{task.id}",
        "dataset_id": int(anchor.id),
    }


def _format_size(num_bytes: int) -> str:
    """字节数 → 人类可读,工具落库 report_content 内部使用。"""
    if not num_bytes:
        return "0 B"
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / 1024 / 1024:.2f} MB"


@router.post("/pdf/to_word")
async def pdf_to_word(
    file: UploadFile = File(...),
    save_to_history: bool = Form(False),
    name: str = Form(""),
    db: Session = Depends(get_db),
):
    """PDF → Word(.docx)。

    save_to_history: 是否把结果落库为历史任务(默认 False)
    """
    content = await file.read()
    _check_pdf_upload(file, content)
    try:
        result = pdf_service.pdf_to_word(content, file.filename or "input.pdf")
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    response: dict[str, Any] = {
        "file_name": result["file_name"],
        "original_size": int(result["original_size"]),
        "output_size": int(result["output_size"]),
        "download_url": _pdf_download_url(result["file_name"]),
    }
    if save_to_history:
        try:
            persisted = _persist_pdf_task(
                db,
                tool="to_word",
                pdf_content=content,
                original_name=file.filename or "input.pdf",
                result=result,
                name_override=name or None,
            )
            response.update(persisted)
        except HTTPException:
            raise
        except Exception as exc:
            logger.warning("PDF 转 Word 落库失败: %s", exc)
    return response


@router.post("/pdf/to_excel")
async def pdf_to_excel(
    file: UploadFile = File(...),
    ai_enabled: bool = Form(True),
    save_as_dataset: bool = Form(False),
    save_to_history: bool = Form(False),
    name: str = Form(""),
    db: Session = Depends(get_db),
):
    """PDF → Excel(.xlsx),每个表格一个 sheet。

    ai_enabled: 是否调用 LLM 辅助识别表头(默认 True,失败自动降级为纯库路径)
    save_as_dataset: 是否落为新数据集(默认 False)
    save_to_history: 是否把工具结果落库为历史任务(默认 False)
    """
    content = await file.read()
    _check_pdf_upload(file, content)
    # AI 增强:先抽表让 LLM 看一下
    ai_hints: list[dict] = []
    if ai_enabled:
        try:
            # 必须复用 pdf_service 的抽取函数:它带 lines → text 自适应降级,
            # 若这里另写一份默认 lines 逻辑,降级后表数量/顺序会与
            # pdf_to_excel 内部不一致,导致 ai_hints 的 table_index 指向错表。
            tables = pdf_service.extract_pdf_tables(content)
            ai_hints = ai_service.refine_pdf_tables(tables)
        except Exception:
            ai_hints = []
    try:
        result = pdf_service.pdf_to_excel(
            content,
            file.filename or "input.pdf",
            ai_hints=ai_hints,
        )
    except ValueError as exc:
        raise HTTPException(400, str(exc))

    response: dict[str, Any] = {
        "file_name": result["file_name"],
        "original_size": int(result["original_size"]),
        "output_size": int(result["output_size"]),
        "table_count": int(result["table_count"]),
        "sheets": result["sheets"],
        "ai_used": bool(result.get("ai_used")),
        "download_url": _pdf_download_url(result["file_name"]),
    }

    # 可选:落为新数据集(供"数据集"列表展示,与"历史任务"互不影响)
    if save_as_dataset:
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            stem = Path(file.filename or "pdf").stem
            ds_name = f"{stem}_pdf表格_{ts}"
            if db.query(Dataset).filter(Dataset.name == ds_name).first():
                n = 1
                while db.query(Dataset).filter(Dataset.name == f"{ds_name}_{n}").first():
                    n += 1
                ds_name = f"{ds_name}_{n}"
            new_ds = Dataset(
                name=ds_name,
                file_name=result["file_name"],
                file_path=result["file_path"],
                row_count=int(result.get("sheets", [{}])[0].get("rows", 0) or 0),
                status="已转换",
            )
            db.add(new_ds)
            db.commit()
            db.refresh(new_ds)
            response["dataset_id"] = int(new_ds.id)
            response["dataset_name"] = new_ds.name
        except Exception as exc:
            logger.warning("PDF 转 Excel 落库失败: %s", exc)

    # 可选:落为历史任务(save_to_history 独立于 save_as_dataset)
    if save_to_history:
        try:
            persisted = _persist_pdf_task(
                db,
                tool="to_excel",
                pdf_content=content,
                original_name=file.filename or "input.pdf",
                result=result,
                name_override=name or None,
            )
            response.update(persisted)
        except HTTPException:
            raise
        except Exception as exc:
            logger.warning("PDF 转 Excel 历史任务落库失败: %s", exc)

    return response


@router.post("/pdf/extract_text")
async def pdf_extract_text(
    file: UploadFile = File(...),
    ai_enabled: bool = Form(True),
    save_to_history: bool = Form(False),
    name: str = Form(""),
    db: Session = Depends(get_db),
):
    """PDF 提取文字,可选 AI 清洗排版。

    save_to_history: 是否把工具结果落库为历史任务(默认 False)
    """
    content = await file.read()
    _check_pdf_upload(file, content)
    # AI 增强:先抽 raw 文本(失败也不影响主流程)
    cleaned_text: str | None = None
    ai_fallback = False
    if ai_enabled:
        try:
            import pdfplumber, io as _io
            raw_parts: list[str] = []
            with pdfplumber.open(_io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text() or ""
                    if txt:
                        raw_parts.append(txt)
            raw = "\n\n".join(raw_parts)
            if raw.strip():
                cleaned_text, ai_fallback = ai_service.clean_pdf_text(raw)
        except Exception:
            cleaned_text = None
            ai_fallback = True
    try:
        result = pdf_service.pdf_extract_text(
            content,
            file.filename or "input.pdf",
            cleaned_text=cleaned_text,
            ai_fallback=ai_fallback,
        )
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    response: dict[str, Any] = {
        "text": result["text"],
        "char_count": int(result["char_count"]),
        "page_count": int(result["page_count"]),
        "ai_fallback": bool(result.get("ai_fallback")),
        "file_name": result["file_name"],
        "download_url": _pdf_download_url(result["file_name"]),
    }
    if save_to_history:
        try:
            persisted = _persist_pdf_task(
                db,
                tool="extract",
                pdf_content=content,
                original_name=file.filename or "input.pdf",
                result=result,
                name_override=name or None,
            )
            response.update(persisted)
        except HTTPException:
            raise
        except Exception as exc:
            logger.warning("PDF 提取文字落库失败: %s", exc)
    return response


@router.post("/pdf/compress")
async def pdf_compress(
    file: UploadFile = File(...),
    save_to_history: bool = Form(False),
    name: str = Form(""),
    db: Session = Depends(get_db),
):
    """PDF 压缩(pikepdf)。

    save_to_history: 是否把工具结果落库为历史任务(默认 False)
    """
    content = await file.read()
    _check_pdf_upload(file, content)
    try:
        result = pdf_service.pdf_compress(content, file.filename or "input.pdf")
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    response: dict[str, Any] = {
        "file_name": result["file_name"],
        "original_size": int(result["original_size"]),
        "output_size": int(result["output_size"]),
        "compression_ratio": float(result["compression_ratio"]),
        "download_url": _pdf_download_url(result["file_name"]),
    }
    if save_to_history:
        try:
            persisted = _persist_pdf_task(
                db,
                tool="compress",
                pdf_content=content,
                original_name=file.filename or "input.pdf",
                result=result,
                name_override=name or None,
            )
            response.update(persisted)
        except HTTPException:
            raise
        except Exception as exc:
            logger.warning("PDF 压缩落库失败: %s", exc)
    return response


# =============================================================================
# 【pdf-photo-to-pdf-ocr 新增】照片 → PDF(OCR 识别)
#
# 设计要点:
#   1. 与既有 4 个 PDF 端点平级,共用下载路由 / 锚点 Dataset / 落库函数。
#   2. OCR 走 ai_service.ocr_images(视觉大模型),逐图软降级:
#      单张识别失败只记 ocr_failed,**不中断整批**,PDF 照常产出。
#   3. ocr_enabled=false 时强制按 image 形态处理,完全跳过视觉模型(省配额)。
#   4. 锚点 Dataset 复用 _ensure_pdf_anchor_dataset,锚点文件即"产出的 PDF",
#      该 dataset 仍是 status="PDF工具产物",前端数据集列表过滤掉,不会污染。
# =============================================================================
@router.post("/pdf/image_to_pdf")
async def pdf_image_to_pdf(
    files: list[UploadFile] = File(...),
    mode: str = Form("dual"),
    ocr_enabled: bool = Form(True),
    save_to_history: bool = Form(False),
    name: str = Form(""),
    db: Session = Depends(get_db),
):
    """多张照片 → PDF(可选 OCR 文字识别)。

    Args:
        files: 1~20 张图片(png/jpg/jpeg/webp/bmp)
        mode: dual(双层,默认) / text(纯文字) / image(图片型)
        ocr_enabled: 是否做文字识别(默认 True);False 时强制按 image 形态
        save_to_history: 是否落库为历史任务(默认 False)
        name: 自定义任务名(可选)
    """
    # 形态校验放在最前:非法 mode 早失败,不必先把图片读进内存
    mode = (mode or "dual").strip().lower()
    if mode not in _IMG_PDF_MODES:
        raise HTTPException(
            400, "PDF 形态只能是 dual / text / image"
        )
    # 关闭 OCR 时不浪费视觉调用,直接按图片型合并
    effective_mode = mode if ocr_enabled else "image"

    _check_image_uploads(files)

    # 逐个读入内存:视觉模型需要 base64,PDF 生成也需要字节,先收齐再统一处理
    images: list[tuple[str, bytes]] = []
    for f in files:
        images.append((f.filename or "image.png", await f.read()))

    # OCR:逐图软降级,失败不抛错,结果与 images 等长同序
    #
    # 【2026-10】dual 形态改用 **ocr_layout**(结构化版面)而不是 ocr_images(纯文本):
    #   用户要求"文字扶正、正向横排、白底黑字、排版整齐,像 Word 导出的合同"。
    #   纯文本只能贴回原图(还是照片的歪斜),而版面 JSON 带着 title/role,
    #   pdf_service._reflow_layout_pdf 才能把文字**重新排版**成真正的文档。
    #   text 形态只需要纯文本,保持用 ocr_images(输出更短、更快)。
    ocr_texts: list[str | None] = []
    ocr_errors: list[str | None] = []
    ocr_layouts: list[dict | None] | None = None
    if ocr_enabled and effective_mode == "image":
        pass  # 图片型不需要识别
    elif ocr_enabled and effective_mode == "dual":
        ocr_layouts, ocr_errors = ai_service.ocr_layout(images)
        # 版面 blocks 里已经含全部文字,顺带拼一份纯文本供前端预览/落库展示
        ocr_texts = [
            "\n".join(
                [b.get("text", "") for b in (lay or {}).get("blocks", []) if b.get("text")]
            ) if isinstance(lay, dict) else None
            for lay in ocr_layouts
        ]
    elif ocr_enabled:
        ocr_texts, ocr_errors = ai_service.ocr_images(images)
    else:
        ocr_texts = [None] * len(images)
        ocr_errors = [None] * len(images)

    try:
        result = pdf_service.images_to_pdf(
            images,
            files[0].filename or "image.png",
            mode=effective_mode,
            ocr_texts=ocr_texts,
            ocr_errors=ocr_errors,
            ocr_layouts=ocr_layouts,
        )
    except ValueError as exc:
        raise HTTPException(400, str(exc))

    # 把识别文字拼成一份完整文本:供响应回传给前端预览 + 落库报告展示
    merged_text = "\n\n".join(
        t for t in (ocr_texts or []) if (t or "").strip()
    )
    result["ocr_text"] = merged_text

    response: dict[str, Any] = {
        "file_name": result["file_name"],
        "original_size": int(result["original_size"]),
        "output_size": int(result["output_size"]),
        "page_count": int(result["page_count"]),
        "mode": result["mode"],
        "mode_label": _IMG_PDF_MODES.get(result["mode"], result["mode"]),
        "image_count": len(images),
        "ocr_enabled": bool(ocr_enabled),
        "ocr_success": int(result["ocr_success"]),
        "ocr_failed": int(result["ocr_failed"]),
        "ocr_errors": result["ocr_errors"],
        "ocr_text": merged_text,
        "layout_used": bool(result.get("layout_used")),
        "download_url": _pdf_download_url(result["file_name"]),
    }
    if save_to_history:
        try:
            # 锚点用"产出的 PDF 字节",而不是原始图片:
            # Task 关联的是产物,后续在报告页下载/二次导出指向的也是这份 PDF。
            out_pdf_bytes = Path(result["file_path"]).read_bytes()
            persisted = _persist_pdf_task(
                db,
                tool="image_to_pdf",
                pdf_content=out_pdf_bytes,
                original_name=files[0].filename or "image.png",
                result=result,
                name_override=name or None,
            )
            response.update(persisted)
        except HTTPException:
            raise
        except Exception as exc:
            logger.warning("照片转 PDF 落库失败: %s", exc)
    return response


@router.get("/pdf/download/{file_name}")
def pdf_download(file_name: str):
    """下载 PDF 工具产出的文件。

    安全:仅允许以 pdf2word_/pdf2excel_/pdftext_/pdfcompress_ 开头的文件名,
    防止路径穿越与任意文件下载。
    """
    # 防路径穿越:file_name 已经是单段 path param(不含 /),再做一次 basename 兜底
    safe = Path(file_name).name
    if safe != file_name or not any(safe.startswith(p) for p in _PDF_DOWNLOAD_PREFIXES):
        raise HTTPException(400, "非法的文件名")
    file_path = settings.upload_dir / safe
    if not file_path.exists():
        raise HTTPException(404, "文件不存在")
    return FileResponse(
        path=str(file_path),
        filename=safe,
        media_type="application/octet-stream",
    )
