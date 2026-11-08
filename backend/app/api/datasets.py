"""数据集 API。"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import pandas as pd
from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from app.core.config import settings, now_local
from app.core.database import get_db
from app.models.dataset import Dataset, INTERNAL_DATASET_STATUS
from app.services import pandas_engine as pe
from app.services.file_extractor import detect_file_type, extract_text, supported_suffixes

router = APIRouter(prefix="/api/datasets", tags=["datasets"])


def _empty_df() -> pd.DataFrame:
    """返回 0 列空 DataFrame,用于单 sheet 预览兜底分支。"""
    return pd.DataFrame()


@router.post("/upload")
def upload_dataset(file: UploadFile = File(...), db: Session = Depends(get_db)):
    """上传数据集,自动解析并落库。

    支持三类文件(2026-10 多格式扩展):
        - structured: .xlsx/.xls/.csv → Pandas 解析所有 sheet 纵向合并
        - document:   .docx/.doc/.pdf/.txt/.md → 抽取纯文本存 extracted_text
        - image:      .png/.jpg/.jpeg/.webp/.bmp → 多模态视觉识别为文本

    设计要点:
        - 非结构化文件的抽取失败**不阻断上传**:文件照常落库,
          extract_status='failed',用户可看到失败原因后重传;
          这样和"先上传后处理"的体验一致,不会因视觉 API 抖动丢文件。
        - 响应新增 file_type / extract_status,前端据此提示用户。
    """
    suffix = Path(file.filename or "").suffix.lower()
    file_type = detect_file_type(file.filename or "")
    if file_type is None:
        allow = ", ".join(supported_suffixes())
        raise HTTPException(400, f"不支持的文件类型: {suffix or '(无后缀)'}。支持: {allow}")

    save_path: Path = settings.upload_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
    save_path.parent.mkdir(parents=True, exist_ok=True)
    with open(save_path, "wb") as f:
        f.write(file.file.read())

    sheets: list[dict] = []
    row_count = 0
    extracted_text = ""
    extract_status = "success"
    status = "已清洗"

    try:
        if file_type == "structured":
            # 结构化:沿用"全部 sheet 合并"读取,确保多 sheet 数据全部进系统
            df, sheets = pe.read_dataframe_all_sheets(save_path)
            row_count = int(len(df))
        else:
            # 非结构化:抽取文本(内部已做异常降级,不会抛出)
            extracted_text, extract_status = extract_text(save_path, file_type)
            if extract_status == "failed":
                # 抽取失败仍保留文件,状态标记"待清洗",前端提示用户内容未能识别
                status = "待清洗"
    except Exception as exc:
        save_path.unlink(missing_ok=True)
        raise HTTPException(400, f"文件解析失败: {exc}")

    ds = Dataset(
        name=Path(file.filename).stem,
        file_name=file.filename,
        file_path=str(save_path),
        row_count=row_count,
        status=status,
        file_type=file_type,
        extracted_text=extracted_text,
        extract_status=extract_status,
    )
    db.add(ds)
    db.commit()
    db.refresh(ds)
    return {
        "id": ds.id,
        "name": ds.name,
        "row_count": ds.row_count,
        "sheets": sheets,
        "file_type": ds.file_type,
        "extract_status": ds.extract_status,
    }


@router.get("")
def list_datasets(db: Session = Depends(get_db)):
    """列表：含本月新增/已清洗/待处理统计。

    改版要点:
        - 每条 Dataset 额外返回 `active_sheet_count`(物理上有多少张工作表),
          前端 `getSheetCount` 据此在「工作表」列展示「📑 N 个工作表」标签;
          CSV 没有 sheet 概念,`list_sheet_names` 返回 `[<文件名>]`,length=1,
          前端 fallback 显示「单表」,与原行为一致。
        - 取 sheet 数失败(文件被磁盘删除/权限问题)时兜底 0,前端按 0 处理显示「单表」,
          不会让整条接口因单条脏数据而 500。

    【dataset-deleted-not-selectable 新增】过滤内部锚点数据集:
        - `status == "PDF工具产物"` 的行是 PDF 工具落库历史任务时的**载体占位**
          (见 models/dataset.py 的 INTERNAL_DATASET_STATUS),不是用户上传的数据。
        - 以前只有前端「数据集」页做了过滤,工作台的对比/合并/清洗下拉拿的是
          全量 items,于是用户在数据集页"看不到"的数据在工作台仍能选到 ——
          表现就是"数据集里已经删掉的工作台还能选"。
        - 现在过滤收敛到接口层:**接口返回什么,两个页面就显示什么**,
          口径唯一;统计(total/本月新增/已清洗/待处理)也只算可见数据集,
          避免"总数 64、列表只有 21 条"的错位。
        - 锚点仍可按 id 单独取用(db.get(Dataset, id)),历史任务/报告页不受影响。
    """
    # —— 关键:查询阶段就排除内部锚点,保证 items 与统计口径一致 ——
    items = (
        db.query(Dataset)
        .filter(Dataset.status != INTERNAL_DATASET_STATUS)
        .order_by(Dataset.create_time.desc())
        .all()
    )
    # 【fix-timezone-asia-shanghai 修复】原 datetime.utcnow 与库里"本地墙钟"不一致,
    # 导致本月统计跨日跨小时判错,这里统一用 now_local()(Asia/Shanghai)。
    now = now_local()
    month_count = sum(
        1 for d in items
        if d.create_time.year == now.year and d.create_time.month == now.month
    )
    cleaned = sum(1 for d in items if d.status == "已清洗")
    pending = sum(1 for d in items if d.status == "待清洗")

    # ——【方案 1】为每条 Dataset 现算 active_sheet_count ——
    # 单条失败不影响整体,容错返回 0
    out_items = []
    for d in items:
        sheet_count = 0
        # 只有结构化文件才有 sheet 概念,非结构化直接跳过读取,省 IO
        if getattr(d, "file_type", "structured") == "structured":
            try:
                sheet_count = len(pe.list_sheet_names(d.file_path))
            except Exception:
                sheet_count = 0
        out_items.append({
            "id": d.id,
            "name": d.name,
            "file_name": d.file_name,
            "row_count": d.row_count,
            "status": d.status,
            "create_time": d.create_time.strftime("%Y-%m-%d"),
            "update_time": d.update_time.strftime("%Y-%m-%d"),
            "active_sheet_count": sheet_count,   # ← 新增字段,前端 getSheetCount 直接读
            # —— 多格式文件支持(2026-10 新增) ——
            # file_type: structured/document/image,前端据此显示类型标签 + 对比弹窗分流
            # extract_status: 非结构化抽取失败时前端显示警示
            "file_type": getattr(d, "file_type", "structured"),
            "extract_status": getattr(d, "extract_status", "success"),
        })

    return {
        "total": len(items),
        "month_new": month_count,
        "cleaned": cleaned,
        "pending": pending,
        "items": out_items,
    }


@router.get("/{dataset_id}/preview")
def preview_dataset(
    dataset_id: int,
    sheet: str | None = None,
    db: Session = Depends(get_db),
):
    """预览数据集内容 + 字段类型 + sheet 明细。

    多格式分流(2026-10 新增):
        - structured: 与旧版一致,返回 DataFrame 前 10 行 + 列类型 + 质量报告;
          额外带 file_type 字段供前端判断。
        - document / image: 不再读 Pandas(会解析失败),改为返回抽取出的
          extracted_text 前 2000 字,columns/rows/quality 返回空结构,
          前端按"文本预览"渲染。

    多 sheet 分级预览(2026-10 新增):
        - 默认(不传 / sheet=""):返回**第一张 sheet** 的真实前 10 行 + 该 sheet 自身的
          missing_total(不会被跨表对齐的 NaN 干扰)。
        - sheet="<sheet 名>":只预览指定的那张 sheet;若该 sheet 不存在返回 400。
        - 输出始终带 `sheets` 列表(全表清单)和 `active_sheet`(本次命中的 sheet 名),
          前端据此渲染 sheet 切换 tab 和当前激活态。
        - **重要**:预览接口**不再**返回"全部 sheet 按列并集合并"的大表,该行为会被
          跨表结构差异污染成上万个 NaN,严重误导分析。AI 分析链路本身**没有动**,
          它们仍走 read_dataframe_all_sheets,与预览解耦。
    """
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")

    file_type = getattr(ds, "file_type", "structured")

    # —— 非结构化:返回文本预览 ——
    if file_type in ("document", "image"):
        text = (ds.extracted_text or "").strip()
        return {
            "file_type": file_type,
            "extract_status": getattr(ds, "extract_status", "success"),
            "extracted_text": text[:2000],
            "text_total_chars": len(text),
            "total": 0,
            "columns": [],
            "rows": [],
            "column_types": {},
            "quality": {},
            "sheets": [],
            "active_sheet": "",
            "active_sheet_count": 0,
        }

    # —— 结构化:先一次性取全表清单(sheets),保证任意路径都返回完整列表 —
    try:
        # sheets 仅用 uncolon-suffixed 文件名占位;try/except 避免单文件 IO 失败
        # 导致整接口 500。失败时给空列表,前端 sheet tab 不显示。
        try:
            sheet_names = pe.list_sheet_names(ds.file_path)
        except Exception:
            sheet_names = []
        sheets_info = [
            {"name": n, "rows": 0, "cols": 0} for n in sheet_names
        ]
        # 用 pe.read_dataframe_per_sheet 一次性拿到真实 rows/cols,
        # 顺便能在切到单 sheet 预览时复用其中的 df。
        per_sheets = pe.read_dataframe_per_sheet(ds.file_path)
        # 用 per_sheet 的真实行/列数覆写 sheet_info,保留 list_sheet_names 给出的顺序
        rows_map = {s["name"]: s["df"] for s in per_sheets}
        for s in sheets_info:
            df_tmp = rows_map.get(s["name"])
            if df_tmp is not None:
                s["rows"] = int(len(df_tmp))
                s["cols"] = int(len(df_tmp.columns))
        # 兼容老 CSV 路径:per_sheet 返回 [{name: 文件名, df}]
        if not sheets_info and per_sheets:
            sheets_info = [
                {"name": s["name"], "rows": int(len(s["df"])), "cols": int(len(s["df"].columns))}
                for s in per_sheets
            ]
    except Exception as exc:
        raise HTTPException(400, f"读取失败: {exc}")

    # —— 决定预览哪张表 ——
    # 默认行为改版(2026-10):
    #   - 不传 / sheet="" → 自动选**第一张 sheet** (sheets_info[0].name),
    #     而不是像之前那样返回"全部 sheet 按列并集合并"的大表。
    #   - 原因:合并大表里跨表对齐产生的 NaN(本表结构无此列)被算成 missing,
    #     5 sheet 文件动辄上万,严重误导分析;预览应看到真实数据。
    #   - AI 分析链路**没有动**:它们仍走 read_dataframe_all_sheets,语义不变。
    requested = (sheet or "").strip()
    target_sheet: str | None = None
    if requested:
        # 显式指定了 sheet 名 → 必须命中真实存在的 sheet,否则报错(避免静默返回错表)
        for s in sheets_info:
            if s["name"] == requested:
                target_sheet = requested
                break
        if target_sheet is None:
            raise HTTPException(400, f"工作表不存在: {requested}")
    elif sheets_info:
        # 默认 → 第一张 sheet
        target_sheet = sheets_info[0]["name"]
    # else: 没有 sheets_info(空文件/CSV 但没 sheet 名) → 后面用 _empty_df 兑底

    if target_sheet is None:
        return {
            **pe.dataframe_to_rows(_empty_df(), 10),
            "file_type": file_type,
            "extract_status": "success",
            "extracted_text": "",
            "text_total_chars": 0,
            "column_types": {},
            "quality": {"row_count": 0, "column_count": 0,
                        "missing_total": 0, "missing_by_column": {},
                        "duplicate_rows": 0},
            "sheets": sheets_info,
            "active_sheet": "",
            "active_sheet_count": len(sheets_info),
        }

    # 取该 sheet 的 DataFrame
    target_df = rows_map.get(target_sheet)
    if target_df is None:
        target_df = next(
            (s["df"] for s in per_sheets if s["name"] == target_sheet),
            None,
        )
    if target_df is None:
        return {
            **pe.dataframe_to_rows(_empty_df(), 10),
            "file_type": file_type,
            "extract_status": "success",
            "extracted_text": "",
            "text_total_chars": 0,
            "column_types": {},
            "quality": {"row_count": 0, "column_count": 0,
                        "missing_total": 0, "missing_by_column": {},
                        "duplicate_rows": 0},
            "sheets": sheets_info,
            "active_sheet": target_sheet,
            "active_sheet_count": len(sheets_info),
        }
    col_types = pe.detect_column_types(target_df)
    quality = pe.basic_quality(target_df)
    return {
        **pe.dataframe_to_rows(target_df, 10),
        "file_type": file_type,
        "extract_status": "success",
        "extracted_text": "",
        "text_total_chars": 0,
        "column_types": col_types,
        "quality": quality,
        "sheets": sheets_info,
        "active_sheet": target_sheet,
        "active_sheet_count": len(sheets_info),
    }


@router.post("/{dataset_id}/rename")
def rename_dataset(dataset_id: int, payload: dict, db: Session = Depends(get_db)):
    """重命名数据集。"""
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")
    new_name = (payload.get("name") or "").strip()
    if not new_name:
        raise HTTPException(400, "名称不能为空")
    ds.name = new_name
    db.commit()
    return {"id": ds.id, "name": ds.name}


@router.delete("/{dataset_id}")
def delete_dataset(dataset_id: int, db: Session = Depends(get_db)):
    """删除数据集：先删文件再删记录。"""
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")
    Path(ds.file_path).unlink(missing_ok=True)
    db.delete(ds)
    db.commit()
    return {"ok": True}


@router.get("/{dataset_id}/download")
def download_dataset(dataset_id: int, db: Session = Depends(get_db)):
    """下载数据集原始文件。

    说明:
        - 用 FileResponse 直接流回磁盘文件,FastAPI 自动按文件大小处理 Range/Content-Length。
        - 文件名优先用数据库里的 file_name(保留原始扩展名),
          兜底用 dataset.name + .xlsx,保证浏览器保存时有正确后缀。
        - 文件已被磁盘删除但 Dataset 行还在时,直接返回 410,前端可提示用户重新上传。
    """
    ds = db.get(Dataset, dataset_id)
    if not ds:
        raise HTTPException(404, "数据集不存在")
    file_path = Path(ds.file_path)
    if not file_path.exists():
        raise HTTPException(410, "源文件已被删除,请重新上传")
    # 仅允许下载已落盘的本地文件,防止路径穿越
    if not str(file_path.resolve()).startswith(str(settings.upload_dir.resolve())):
        raise HTTPException(400, "非法的文件路径")
    # 优先用 file_name 保留扩展名,兜底用 name
    download_name = ds.file_name or f"{ds.name}.xlsx"
    return FileResponse(
        path=str(file_path),
        filename=download_name,
        media_type="application/octet-stream",
    )
