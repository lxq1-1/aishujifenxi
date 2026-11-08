"""任务 API：创建/查询/列表/重新分析/删除/意见反馈迭代。"""
import json
import logging

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException
from sqlalchemy.orm import Session

from app.core.database import get_db
from app.models.task import Task
from app.models.dataset import Dataset
from app.models.report_template import ReportTemplate
from app.services.analyzer import run_analysis
from app.services import ai_service

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/tasks", tags=["tasks"])


# ============================== 图表聚合工具 ==============================
# 历史上 chart_option 的存储形态经历过三次变化,且不同模板会同时落库:
#   1) 顶层 charts: [ {key,title,type,option}, ... ]               ← 早期
#   2) by_section:  { "节标题": [ {key,title,option}, ... ] }     ← per-section-charts
#   3) reports[*].charts: [ ... ]                                ← 兜底
# 此外用户在「图表编辑」里前置预选并提交 AI 的图表,
# 会先落到 selected_charts_json,运行后由 analyzer 按 title 去重合并到 chart_option_json
# (见 analyzer.py),但**老数据 / 失败重跑 / 兼容**等场景可能还留在 selected_charts_json。
#
# 因此接口侧必须把这四种来源都聚合起来,按 key 去重,保证前端"图表编辑"
# 弹窗「选择图表」下拉有完整候选清单 —— 之前只看顶层 charts,会把 by_section 形态
# 误判为「0 个图表」。

def _collect_available_charts(chart_option_json: str, selected_charts_json: str | None) -> list[dict]:
    """聚合一个任务的所有可被改图的图表(去重保序)。"""
    try:
        co = json.loads(chart_option_json or "{}")
    except Exception:
        co = {}
    try:
        sel = json.loads(selected_charts_json or "[]") or []
    except Exception:
        sel = []

    seen: set[str] = set()
    ordered: list[dict] = []

    def _ingest(charts_obj) -> None:
        if not charts_obj:
            return
        for c in charts_obj:
            if not isinstance(c, dict):
                continue
            key = c.get("key") or c.get("title") or ""
            if key and key in seen:
                continue
            if key:
                seen.add(key)
            ordered.append(c)

    # 用户预选(无论是否已合并进 chart_option_json,先纳入)
    _ingest(sel)
    # 形态 1:顶层 charts
    _ingest(co.get("charts"))
    # 形态 2:by_section:{section:[{key,title,option}]} 或 [{section,charts:[...]}]
    bs = co.get("by_section")
    if isinstance(bs, dict):
        for _, lst in bs.items():
            _ingest(lst)
    elif isinstance(bs, list):
        for entry in bs:
            if isinstance(entry, dict):
                _ingest(entry.get("charts"))
    # 形态 3:reports[*].charts
    for rep in (co.get("reports") or []):
        if isinstance(rep, dict):
            _ingest(rep.get("charts"))

    return ordered


# 注意:此前对 analysis_intent 和 feedback 都设了 500 字硬上限,
# 与"开放式反馈/不再用硬字数"理念冲突,现已放开:仅做 strip 与非空判断,
# 不做长度校验;Prompt 兜底截断由 prompt_builder 处按 token 上限处理(不影响 DB 存储)。


def _do_analyze(
    task_id: int,
    cleaning_rules: dict,
    analysis_intent: str = "",
    feedback: str = "",
) -> None:
    """后台线程：实际执行分析。

    Spec A:`analysis_intent` 透传到 run_analysis,用于注入 Prompt 第 6 段 + annotations info 标注。
    Spec B:`feedback` 透传到 run_analysis,用于注入 Prompt 第 7 段(不写 annotations,因为
            它已在 task.feedback_text 落库)。
    Spec C(generate-ai-presentation-deck):分析成功后**顺手预生成**一份 AI 演示稿,
            这样用户点「导出 PPT」时直接命中缓存、秒下,不必等 1~3 分钟现场生成。
            预生成失败**绝不影响**报告本身 —— 整段包 try/except,
            task.status 保持 success,只在 deck 表留一条 failed 记录。
    """
    from app.core.database import SessionLocal
    db = SessionLocal()
    try:
        task = db.get(Task, task_id)
        if task:
            run_analysis(
                db,
                task,
                cleaning_rules,
                analysis_intent=analysis_intent,
                feedback=feedback,
            )
            # ---- 预生成 AI 演示稿(失败不影响报告)----
            # 只在报告真正成功时才生成:failed 的任务没有可演示的内容。
            if getattr(task, "status", None) == "success":
                try:
                    from app.services import deck_service
                    deck_service.ensure_deck(db, task)
                except Exception as exc:  # pragma: no cover - 演示稿是附加能力
                    logger.warning(
                        "[deck] 预生成演示稿失败(task=%s): %s", task_id, exc
                    )
    finally:
        db.close()


# ============================== 版本 lineage 工具 ==============================

def _lineage_depth(db: Session, t: Task) -> int:
    """沿 parent_task_id 链向上追溯,返回祖先节点数(含 t 本身)。

    示例:
        基线 t(v1) → depth = 0
        v2(其父为基线) → depth = 1
        v3(其父为 v2) → depth = 2
    """
    depth = 0
    cur = t
    # 防御性循环:任何异常路径都断在 1000 步
    for _ in range(1000):
        if cur.parent_task_id is None:
            return depth
        parent = db.get(Task, cur.parent_task_id)
        if parent is None or parent.id == cur.id:
            return depth
        depth += 1
        cur = parent
    return depth


def _version_label_for(db: Session, t: Task) -> str:
    """把 lineage 深度转成展示用的版本标签:基线 → "", v2 → "v2" ..."""
    if t.parent_task_id is None:
        return ""
    return f"v{_lineage_depth(db, t) + 1}"


# ============================== 创建任务 ==============================

@router.post("")
def create_task(payload: dict, background: BackgroundTasks, db: Session = Depends(get_db)):
    """创建任务并立即在后台执行分析。

    请求体字段:
        dataset_id:         必填
        template_id:        必填
        cleaning_rules:     可选
        use_knowledge_base: 可选
        kb_ids:             可选
        name:               可选
        analysis_intent:    可选,工作台临时输入的分析重点(≤500 字)
        feedback:           可选,**预留**,当前 spec 不在创建时使用(留给将来扩展)
    """
    name = payload.get("name") or "未命名任务"
    dataset_id = payload.get("dataset_id")
    template_id = payload.get("template_id")
    cleaning_rules = payload.get("cleaning_rules") or {}
    use_kb = bool(payload.get("use_knowledge_base", False))
    kb_ids = payload.get("kb_ids") or ""
    # 仅做 strip,不限制长度(已放开 500 上限)
    intent_raw = (payload.get("analysis_intent") or "").strip()
    # feedback 在「创建任务」路径不被使用(spec 限定意见反馈走 /regenerate 接口),
    # 这里兼容读取;不传或传空均忽略。
    fb_raw = (payload.get("feedback") or "").strip()

    if not dataset_id or not template_id:
        raise HTTPException(400, "dataset_id 与 template_id 必填")

    if not db.get(Dataset, dataset_id):
        raise HTTPException(404, "数据集不存在")
    if not db.get(ReportTemplate, template_id):
        raise HTTPException(404, "模板不存在")

    task = Task(
        name=name,
        dataset_id=dataset_id,
        template_id=template_id,
        status="processing",
        use_knowledge_base=use_kb,
        kb_ids=kb_ids,
    )
    db.add(task)
    db.commit()
    db.refresh(task)

    background.add_task(_do_analyze, task.id, cleaning_rules, intent_raw)
    return {"id": task.id, "status": task.status}


@router.get("")
def list_tasks(db: Session = Depends(get_db)):
    import json
    items = db.query(Task).order_by(Task.create_time.desc()).all()
    # 【tools-results-history-persistence 新增】一次性拉出所有引用过的模板,
    # 缓存为 {template_id: category},避免在循环里多次 query。
    from app.models.report_template import ReportTemplate
    template_categories: dict[int, str] = {}
    template_ids = {t.template_id for t in items if t.template_id}
    if template_ids:
        rows = (
            db.query(ReportTemplate)
            .filter(ReportTemplate.id.in_(template_ids))
            .all()
        )
        template_categories = {r.id: (r.category or "") for r in rows}
    out = []
    for t in items:
        # 解析图表数量 + 第一个图表的标题,用于列表显示「图表编辑」入口
        # 委托给公共聚合函数(同时兼容 charts / by_section / reports[*].charts / selected_charts)
        ordered = _collect_available_charts(t.chart_option_json, t.selected_charts_json)
        chart_count = len(ordered)
        first_title = ordered[0].get("title") if ordered else ""
        # 【tools-results-history-persistence 新增】模板分类(供前端列表打角标)
        # 模板已被删除时填 "";前端不显示角标即可,行为安全。
        tpl_category = template_categories.get(t.template_id, "")
        out.append({
            "id": t.id,
            "name": t.name,
            "dataset_id": t.dataset_id,
            "template_id": t.template_id,
            "status": t.status,
            "chart_count": chart_count,
            "first_chart_title": first_title,
            "create_time": t.create_time.strftime("%Y-%m-%d %H:%M:%S"),
            "update_time": t.update_time.strftime("%Y-%m-%d %H:%M:%S"),
            "error": t.error_message,
            # Spec B:报告版本 lineage 字段
            "parent_task_id": t.parent_task_id,
            "version_label": _version_label_for(db, t),
            # 用于历史任务列表展示迭代脉络;截断避免列表过长
            "feedback_text": (t.feedback_text or "")[:120],
            # 【tools-results-history-persistence 新增】模板分类,前端按 category 显示角标
            # ("数据校验" / "数据问答" / "数据对比" 时打对应角标;其它不显示)
            "template_category": tpl_category,
        })
    return {"items": out}


@router.get("/{task_id}")
def get_task(task_id: int, db: Session = Depends(get_db)):
    t = db.get(Task, task_id)
    if not t:
        raise HTTPException(404, "任务不存在")
    import json
    # 解析用户预选的图表(留痕),空时给空列表
    try:
        selected_charts = json.loads(t.selected_charts_json or "[]")
        if not isinstance(selected_charts, list):
            selected_charts = []
    except Exception:
        selected_charts = []
    # 【图表编辑】候选图表清单(后端统一聚合,前端无需再判形态)
    available_charts = _collect_available_charts(t.chart_option_json, t.selected_charts_json)
    return {
        "id": t.id,
        "name": t.name,
        "dataset_id": t.dataset_id,
        "template_id": t.template_id,
        "status": t.status,
        "error": t.error_message,
        "indicators": json.loads(t.indicators_json or "{}"),
        "report": json.loads(t.report_content or "{}"),
        "chart_option": json.loads(t.chart_option_json or "{}"),
        "annotations": json.loads(t.annotations_json or "[]"),
        "create_time": t.create_time.strftime("%Y-%m-%d %H:%M:%S"),
        # Spec B:版本 lineage 字段(报告页做"上一版本"导航、回显迭代意见)
        "parent_task_id": t.parent_task_id,
        "version_label": _version_label_for(db, t),
        "feedback_text": t.feedback_text or "",
        # 【chart-edit-prebuilt-dashboard 新增】用户预选图表的"快照",
        # 前端报告页可显示"用户挑选的图表"标记;空时给空列表。
        "selected_charts": selected_charts,
        # 【图表编辑】可直接渲染的"图表清单" —— 兼容 charts / by_section / reports[*].charts,
        # 同时纳入 selected_charts 兜底;按 key 去重,首项即「首张图表」。
        "available_charts": available_charts,
        "available_chart_count": len(available_charts),
    }


@router.post("/{task_id}/reanalyze")
def reanalyze(
    task_id: int,
    background: BackgroundTasks,
    payload: dict | None = None,
    db: Session = Depends(get_db),
):
    """基于已有 task 重新分析(用相同 dataset/template)。可选 analysis_intent。

    与 regenerate 的区别:
        reanalyze = 同一个 task 把报告重写一遍(lineage 不变,id 不变);
        regenerate = 产生一个新 task(parent_task_id 指回原 task),保留原 task。
    """
    t = db.get(Task, task_id)
    if not t:
        raise HTTPException(404, "任务不存在")
    # 允许可选 analysis_intent(已放开 500 上限,仅 strip)
    intent_raw = ((payload or {}).get("analysis_intent") or "").strip()

    t.status = "processing"
    t.error_message = ""
    db.commit()
    background.add_task(_do_analyze, t.id, {}, intent_raw)
    return {"id": t.id, "status": "processing"}


@router.post("/{task_id}/summarize")
def summarize_task(task_id: int, db: Session = Depends(get_db)):
    """重新生成 AI 摘要,并写回 ``report_content.summary``。

    入参:无(``task_id`` 从 URL 取)。

    返回:
        最新 task 详情(同 ``GET /tasks/{task_id}``),
        方便前端调用后直接刷新报告页。

    失败模式:
        - 任务不存在 → 404
        - 任务尚未分析完成(status!=success)→ 409,前端可在按钮置灰时拦截
        - 其它异常 → 500 并保留原 summary
    """
    t = db.get(Task, task_id)
    if not t:
        raise HTTPException(404, "任务不存在")
    if t.status != "success":
        raise HTTPException(409, f"任务状态为 {t.status},请等待分析完成后再生成摘要")

    try:
        # 重新解析已存的 report / indicators
        report_obj = json.loads(t.report_content or "{}") or {}
        indicators = json.loads(t.indicators_json or "{}") or {}
        summary = ai_service.summarize_report(report_obj, indicators)
    except Exception as exc:
        # 摘要生成失败不应破坏已存报告 —— 仅日志,返回旧 summary
        logger.warning("summarize_task 失败,保留旧摘要: %s", exc)
        raise HTTPException(500, f"摘要生成失败:{exc}")

    # 把新摘要合并进 report_content(保留 title / sections)
    new_report = {
        **report_obj,
        "summary": summary,
    }
    t.report_content = json.dumps(new_report, ensure_ascii=False)
    db.commit()
    db.refresh(t)

    return {
        "id": t.id,
        "name": t.name,
        "status": t.status,
        "indicators": json.loads(t.indicators_json or "{}"),
        "report": json.loads(t.report_content or "{}"),
        "chart_option": json.loads(t.chart_option_json or "{}"),
        "annotations": json.loads(t.annotations_json or "[]"),
        "parent_task_id": t.parent_task_id,
        "version_label": _version_label_for(db, t),
        "feedback_text": t.feedback_text or "",
    }


@router.post("/{task_id}/regenerate")
def regenerate_task(
    task_id: int,
    payload: dict,
    background: BackgroundTasks,
    db: Session = Depends(get_db),
):
    """基于已有 task,按用户意见"重新生成一份"作为新版本任务。

    关键约束(Spec B):
        - 不修改原 task 的任何字段(status / indicators / report / chart / annotations 全部保留);
        - 新建一份 task,继承原 task 的 dataset_id/template_id/use_kb/kb_ids;
        - 新 task 的 parent_task_id = 原 task id;
        - 新 task name = "<原 name> - v{N}",N 由 _lineage_depth 计算;
        - 新 task 进入后台 _do_analyze,传 feedback;
        - 校验:
            * task 不存在 → 404
            * task.status != 'success' → 409(任务未完成不能重生成)
            * feedback 缺失或空字符串 → 400
            * feedback 长度 > 500 → 400

    请求体:`{feedback: str(1~500)}`

    返回:同 shape 的新 task 详情(GET /api/tasks/{id})。
    """
    feedback = (payload.get("feedback") or "").strip()
    if not feedback:
        raise HTTPException(400, "请填写修改意见")
    # 已放开 500 上限,不再校验长度

    src = db.get(Task, task_id)
    if not src:
        raise HTTPException(404, "任务不存在")
    if src.status != "success":
        raise HTTPException(
            409,
            f"任务状态为 {src.status},请等待完成后再生成新版",
        )

    # 计算新版本号 N
    depth = _lineage_depth(db, src)
    new_name = f"{src.name} - v{depth + 2}"

    new_task = Task(
        name=new_name,
        dataset_id=src.dataset_id,
        template_id=src.template_id,
        status="processing",
        use_knowledge_base=src.use_knowledge_base,
        kb_ids=src.kb_ids,
        parent_task_id=src.id,
        feedback_text=feedback,
    )
    db.add(new_task)
    db.commit()
    db.refresh(new_task)

    # 把 feedback 透传到后台;不传 analysis_intent(避免信息冗余)
    background.add_task(_do_analyze, new_task.id, {}, "", feedback)

    # 返回新 task 的详情(同 GET /api/tasks/{id} shape),前端可直接用
    return get_task(new_task.id, db)


@router.delete("/batch")
def batch_delete(payload: dict, db: Session = Depends(get_db)):
    ids = payload.get("ids") or []
    if not ids:
        return {"deleted": 0}
    for tid in ids:
        t = db.get(Task, tid)
        if t:
            db.delete(t)
    db.commit()
    return {"deleted": len(ids)}


@router.delete("/{task_id}")
def delete_task(task_id: int, db: Session = Depends(get_db)):
    t = db.get(Task, task_id)
    if not t:
        raise HTTPException(404, "任务不存在")
    db.delete(t)
    # 注意:子任务的 parent_task_id 通过 ondelete="SET NULL" 自动置空,
    # 保证 lineage 上其余版本仍可访问。
    db.commit()
    return {"ok": True}
