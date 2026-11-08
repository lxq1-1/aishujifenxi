"""报告模板 API。支持列表/详情/新建/更新/删除/复制,内置模板受保护。"""
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field
from sqlalchemy.orm import Session

from app.core.database import get_db
from app.models.report_template import ReportTemplate

router = APIRouter(prefix="/api/templates", tags=["templates"])


# ---------- Pydantic 请求体 ----------

class TemplateUpsert(BaseModel):
    """新建/更新模板的请求体。

    注意:不再接受 `prompt_structure` 字段。Prompt 模板(角色/输出 schema/占位符)
    由后端 `app.core.prompt_builder` 按分类统一制定,UI 不可编辑。
    """
    name: str = Field(..., min_length=1, max_length=100, description="模板名称(必填,1~100)")
    category: str = Field(..., min_length=1, max_length=50, description="分类(必填,1~50)")
    description: str = Field(default="", max_length=500)
    structure: str = Field(default="", description="报告章节结构(Markdown,由用户在前端填写)")


# ---------- 序列化 ----------

def _serialize(t: ReportTemplate) -> dict:
    """把 ORM 对象转 dict。

    模板不再返回 `prompt_structure`(后端统一管);详情接口与列表接口字段一致。
    """
    return {
        "id": t.id,
        "name": t.name,
        "category": t.category,
        "description": t.description or "",
        "structure": t.structure or "",
        "is_builtin": t.is_builtin or 0,
        "create_time": t.create_time.isoformat() if t.create_time else None,
    }


# ---------- 列表 / 详情 / 分类 ----------

@router.get("")
def list_templates(category: str | None = None, db: Session = Depends(get_db)):
    """模板列表,可按 category 筛选。"""
    q = db.query(ReportTemplate).order_by(ReportTemplate.is_builtin.desc(), ReportTemplate.id.asc())
    if category:
        q = q.filter(ReportTemplate.category == category)
    items = q.all()
    return {"items": [_serialize(t) for t in items]}


@router.get("/categories")
def list_categories(db: Session = Depends(get_db)):
    """返回所有去重后的分类列表,供前端下拉。"""
    rows = db.query(ReportTemplate.category).distinct().all()
    cats = sorted({(r[0] or "其他") for r in rows if r[0]})
    return {"items": cats}


@router.get("/{template_id}")
def get_template(template_id: int, db: Session = Depends(get_db)):
    """模板详情(不再包含 prompt_structure,该字段由后端按分类统一管理)。"""
    t = db.get(ReportTemplate, template_id)
    if not t:
        raise HTTPException(404, "模板不存在")
    return _serialize(t)


# ---------- CRUD ----------

@router.post("")
def create_template(payload: TemplateUpsert, db: Session = Depends(get_db)):
    """新建模板。用户自建模板,is_builtin=0。"""
    # 同名校验:避免与现有模板名冲突
    exists = db.query(ReportTemplate).filter(ReportTemplate.name == payload.name).first()
    if exists:
        raise HTTPException(400, f"已存在同名模板:{payload.name}")
    t = ReportTemplate(
        name=payload.name,
        category=payload.category,
        description=payload.description,
        structure=payload.structure,
        is_builtin=0,  # 用户自建
    )
    db.add(t)
    db.commit()
    db.refresh(t)
    return _serialize(t)


@router.put("/{template_id}")
def update_template(template_id: int, payload: TemplateUpsert, db: Session = Depends(get_db)):
    """更新模板字段。不允许改 is_builtin。"""
    t = db.get(ReportTemplate, template_id)
    if not t:
        raise HTTPException(404, "模板不存在")
    # 若改名,检查不与其他模板冲突
    if payload.name != t.name:
        conflict = db.query(ReportTemplate).filter(
            ReportTemplate.name == payload.name,
            ReportTemplate.id != template_id,
        ).first()
        if conflict:
            raise HTTPException(400, f"已存在同名模板:{payload.name}")
    t.name = payload.name
    t.category = payload.category
    t.description = payload.description
    t.structure = payload.structure
    db.commit()
    db.refresh(t)
    return _serialize(t)


@router.delete("/{template_id}")
def delete_template(template_id: int, db: Session = Depends(get_db)):
    """删除模板。内置模板拒绝并返回 400。"""
    t = db.get(ReportTemplate, template_id)
    if not t:
        raise HTTPException(404, "模板不存在")
    if t.is_builtin == 1:
        raise HTTPException(400, "内置模板不可删除,请先复制后修改")
    db.delete(t)
    db.commit()
    return {"deleted": template_id}


@router.post("/{template_id}/duplicate")
def duplicate_template(template_id: int, db: Session = Depends(get_db)):
    """复制模板:基于源模板创建"<源名> - 副本"的自建模板。"""
    src = db.get(ReportTemplate, template_id)
    if not src:
        raise HTTPException(404, "源模板不存在")
    new_name = f"{src.name} - 副本"
    # 避免多次重复时名字冲突:追加序号
    suffix = 1
    final_name = new_name
    while db.query(ReportTemplate).filter(ReportTemplate.name == final_name).first():
        suffix += 1
        final_name = f"{new_name} ({suffix})"
    dup = ReportTemplate(
        name=final_name,
        category=src.category,
        description=src.description,
        structure=src.structure,
        is_builtin=0,
    )
    db.add(dup)
    db.commit()
    db.refresh(dup)
    return _serialize(dup)
