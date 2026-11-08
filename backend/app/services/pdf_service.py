"""PDF 处理服务：转 Word / 转表格 / 提取文字 / 压缩。

依赖:
    - pdfplumber   文字与表格抽取
    - pdf2docx     PDF → Word 转换
    - pikepdf      PDF 压缩

设计要点:
    - 4 项能力各自独立函数,统一抛 ValueError(由路由转 400)
    - 文件输出统一放 settings.upload_dir,文件名带 pdf<action>_<ts>_<原名> 前缀
    - 表格 sheet 名表1/表2...,超 31 字截断 + 重名加 _2/_3 后缀
    - 压缩用 pikepdf + linearize + object_stream_mode=generate
"""
from __future__ import annotations

import io
import logging
import re
import time
from pathlib import Path
from typing import Any

import pdfplumber
from pdf2docx import Converter
import pikepdf

# —— 多张照片转 PDF(2026-10 新增)——
# Pillow 只做解码/方向校正/缩放,reportlab 只做排版与落盘,中文字体用
# reportlab 自带的 STSong-Light(CID 字体),不引入任何外部 ttf 文件。
from PIL import Image, ImageOps
# 照片 → 白底文档 的图像净化流水线专用(不新增任何第三方依赖)。
# opencv-python-headless 锁在 4.10.x 是为了兼容本项目锁定的 numpy 1.26.4,
# 理由见 requirements.txt 里的注释;这里用不到 GUI 相关能力,headless 足够。
import cv2
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    Frame,
    Image as RLImage,
    PageBreak,
    PageTemplate,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from xml.sax.saxutils import escape

from app.core.config import settings

logger = logging.getLogger(__name__)

# 输出文件名前缀(供 /download 路由做安全校验)
PREFIX_TO_WORD = "pdf2word_"
PREFIX_TO_EXCEL = "pdf2excel_"
PREFIX_EXTRACT = "pdftext_"
PREFIX_COMPRESS = "pdfcompress_"
PREFIX_IMG2PDF = "img2pdf_"
# 「照片 → Word 风格重排 PDF」的输出前缀(2026-10 新增)。
# 与其它 PREFIX_* 一样,仅供 /download 路由做安全校验与日志区分。
PREFIX_CLEANPDF = "cleanpdf_"

# Paragraph 文本里的 HTML 标签(<b>…</b> 等),anchor 匹配前先剥掉。
# 模块级预编译:_find_anchor_position 会对 story 里每个段落调用,反复编译不值。
_HTML_TAG_RE = re.compile(r"<[^>]+>")


def _ts() -> str:
    """统一时间戳,用于输出文件名。"""
    return time.strftime("%Y%m%d_%H%M%S")


def _safe_output_name(prefix: str, original_name: str, suffix: str) -> str:
    """拼接输出文件名: <prefix><ts>_<原 stem>.<suffix>。"""
    stem = Path(original_name or "file").stem
    # 去非法字符,保留中文 / 数字 / 字母
    stem = re.sub(r"[\\/:\*\?\"<>\|]", "_", stem).strip() or "file"
    return f"{prefix}{_ts()}_{stem}.{suffix}"


def _safe_sheet_name(raw: str, used: set[str]) -> str:
    """与 /tools/merge multi_sheet 同款 sheet 名清洗逻辑。"""
    cleaned = re.sub(r"[\\/\:\*\?\[\]]", "_", str(raw)).strip() or "sheet"
    cleaned = cleaned[:31]
    base = cleaned
    n = 2
    while cleaned in used:
        suffix = f"_{n}"
        cleaned = (base[: 31 - len(suffix)] + suffix)
        n += 1
    used.add(cleaned)
    return cleaned


def _table_data_rows(table: list) -> int:
    """统计一张表里"有文字的数据行"数量(第 1 行按表头处理,不计入)。"""
    n = 0
    for row in table[1:]:
        if any(cell is not None and str(cell).strip() for cell in row):
            n += 1
    return n


def _merge_close_positions(values: list[float], gap: float = 2.0) -> list[float]:
    """把靠得极近的坐标合并成一个。

    PDF 里一条表格线常被绘制成两条并排的细线(实测同一条竖线会出现
    x=53.5 与 x=54.5 两个值)。若不合并,列边界会被切成两条相邻线,
    导致中间多出一个空列、且最后一列被挤出表格。
    """
    merged: list[float] = []
    for v in sorted(values):
        if not merged or v - merged[-1] > gap:
            merged.append(v)
    return merged


def _derive_table_edges(page: Any) -> tuple[list[float], list[float]] | None:
    """从页面的矩形/直线中推导出表格的列边界与行边界。

    适用场景:表格**只画了表头行的框线**,数据行是纯文字(没有框)。
    此时 lines 策略只能切出表头那 1 行,而 text 策略又因为表头文字与
    数据文字的 x 起点不一致(表头居中、数据左对齐)把它们判成互不相关的
    行组,结果丢掉表头。

    这里改用"显式边界"策略:
        - 列边界:取所有矩形左右边的 x 坐标(合并近似重合的);
        - 行边界:取每一行文字的 top/bottom,各外扩 1pt 作为切分线。
    这样既保留了框线给出的准确列划分,又能覆盖到没有框线的下方数据区。
    """
    rects = list(getattr(page, "rects", []) or [])
    edges = list(getattr(page, "edges", []) or [])

    xs: list[float] = []
    for r in rects:
        xs.extend([float(r["x0"]), float(r["x1"])])
    for e in edges:
        if e.get("orientation") == "v":
            xs.append(float(e["x0"]))
    # 线条太少说明这页不是"表头带框"的表格,交给其它策略处理
    col_edges = _merge_close_positions(xs)
    if len(col_edges) < 2:
        return None

    # 行边界:按文字行的上下沿切
    rows: dict[float, list[dict]] = {}
    for w in page.extract_words():
        rows.setdefault(round(float(w["top"]), 1), []).append(w)
    if not rows:
        return None

    ys: list[float] = []
    for top in sorted(rows):
        group = rows[top]
        ys.append(min(float(w["top"]) for w in group) - 1.0)
        ys.append(max(float(w["bottom"]) for w in group) + 1.0)
    row_edges = _merge_close_positions(ys, gap=1.0)
    if len(row_edges) < 2:
        return None
    return col_edges, row_edges


def extract_pdf_tables(pdf_bytes: bytes) -> list[list[list[str]]]:
    """从 PDF 中抽取所有表格,采用多策略择优。

    背景 —— 单一策略都不够用：
        - ``lines``(pdfplumber 默认):只把**直线**当单元格边界。
          边框用矩形绘制、或数据行没有横线时,只能切出表头那一行,
          表现为"导出的表格只有表头、没有数据"。
        - ``text``:按文字位置切分,不依赖框线。但当表头文字居中、
          数据文字左对齐时,两者 x 起点不一致会被判成互不相关的行组,
          表现为"只有数据、没有表头"。

    策略(按优先级尝试,谁抽到的有效数据行多就采用谁):
        1) ``lines``     —— 规整表格最准;
        2) ``explicit``  —— 由矩形边框推列边界 + 由文字位置推行边界,
                           专门修复"只有表头带框线"的表格;
        3) ``text``      —— 兜底,适配完全无线条的排版。

    注意:``ai_hints`` 的 table_index 与这里的表序号一一对应,因此抽取逻辑
    必须集中在本函数内,避免 API 层与 Service 层各抽一遍导致索引错位。
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages = pdf.pages

        # 第一轮:默认 lines 策略
        lines_tables: list[list[list[str]]] = []
        for page in pages:
            lines_tables.extend(page.extract_tables() or [])
        if any(_table_data_rows(tb) > 0 for tb in lines_tables):
            return lines_tables

        # 第二轮:显式边界(列来自框线,行来自文字位置)
        explicit_tables: list[list[list[str]]] = []
        for page in pages:
            derived = _derive_table_edges(page)
            if not derived:
                continue
            col_edges, row_edges = derived
            try:
                explicit_tables.extend(
                    page.extract_tables(table_settings={
                        "vertical_strategy": "explicit",
                        "explicit_vertical_lines": col_edges,
                        "horizontal_strategy": "explicit",
                        "explicit_horizontal_lines": row_edges,
                        "snap_tolerance": 3,
                        "join_tolerance": 3,
                    }) or []
                )
            except Exception:
                continue
        if sum(_table_data_rows(tb) for tb in explicit_tables) > sum(
            _table_data_rows(tb) for tb in lines_tables
        ):
            return explicit_tables

        # 第三轮:text 兜底
        text_tables: list[list[list[str]]] = []
        for page in pages:
            text_tables.extend(page.extract_tables(table_settings={
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
            }) or [])

        best = max(
            (lines_tables, explicit_tables, text_tables),
            key=lambda group: sum(_table_data_rows(tb) for tb in group),
        )
        return best


def pdf_to_word(pdf_bytes: bytes, original_name: str) -> dict[str, Any]:
    """PDF → Word(.docx)。

    Returns:
        {file_path, file_name, original_size, output_size}
    """
    if not pdf_bytes:
        raise ValueError("PDF 内容为空")
    try:
        out_name = _safe_output_name(PREFIX_TO_WORD, original_name, "docx")
        out_path = settings.upload_dir / out_name
        settings.upload_dir.mkdir(parents=True, exist_ok=True)
        # pdf2docx 需要先落盘再转,先写临时 PDF
        tmp_pdf = settings.upload_dir / f"_tmp_{_ts()}_{Path(original_name or 'in').stem}.pdf"
        tmp_pdf.write_bytes(pdf_bytes)
        try:
            cv = Converter(str(tmp_pdf))
            cv.convert(str(out_path))
            cv.close()
        finally:
            try:
                tmp_pdf.unlink(missing_ok=True)
            except Exception:
                pass
        if not out_path.exists():
            raise ValueError("转换失败:未生成目标 docx 文件")
        return {
            "file_path": str(out_path),
            "file_name": out_name,
            "original_size": int(len(pdf_bytes)),
            "output_size": int(out_path.stat().st_size),
        }
    except ValueError:
        raise
    except Exception as exc:
        logger.exception("pdf_to_word 失败: %s", exc)
        raise ValueError(f"PDF 转 Word 失败:{exc}")


def pdf_to_excel(
    pdf_bytes: bytes,
    original_name: str,
    *,
    ai_hints: list[dict] | None = None,
) -> dict[str, Any]:
    """PDF → Excel(.xlsx),每个表格写一个 sheet。

    ai_hints 由 ai_service.refine_pdf_tables 提供:
        [{table_index, header_row_index, merged_cells_fix}],失败或 None 时为 []。
    """
    if not pdf_bytes:
        raise ValueError("PDF 内容为空")
    try:
        import pandas as pd
        out_name = _safe_output_name(PREFIX_TO_EXCEL, original_name, "xlsx")
        out_path = settings.upload_dir / out_name
        settings.upload_dir.mkdir(parents=True, exist_ok=True)

        # 用 pdfplumber 抽所有表格(内部自带 lines → text 的自适应降级)
        all_tables = extract_pdf_tables(pdf_bytes)

        if not all_tables:
            raise ValueError("未在 PDF 中检测到表格")

        hints_map: dict[int, dict] = {}
        for h in (ai_hints or []):
            try:
                idx = int(h.get("table_index"))
                if 0 <= idx < len(all_tables):
                    hints_map[idx] = h
            except Exception:
                continue

        # ===== 第一步：先在内存里把每张表整理成 DataFrame,确认非空后再开 writer =====
        # 为什么要两段式：
        #   pandas 的 ExcelWriter 在 with 块退出时才真正 save(),而 openpyxl 在
        #   workbook 里一张可见 sheet 都没有时会抛
        #   IndexError("At least one sheet must be visible")。
        #   如果边判断边写、所有表都被跳过,就会得到一个"空工作簿",报错发生在
        #   with 的 __exit__ 里,既不落到下面那句中文提示,又把英文异常直接暴露给前端。
        #   因此这里先全部构建好,再一次性写盘。
        def _cell_text(c: Any) -> str:
            """单元格统一转成去空白的字符串,None → ""。"""
            return "" if c is None else str(c).strip()

        used_names: set[str] = set()
        sheet_summaries: list[dict] = []
        pending: list[tuple[int, str, Any]] = []  # (表序号, sheet 名, DataFrame)

        for i, table in enumerate(all_tables):
            if not table:
                continue
            # LLM 提示的 header_row_index,缺省 0
            hint = hints_map.get(i) or {}
            try:
                header_idx = int(hint.get("header_row_index", 0))
            except (TypeError, ValueError):
                header_idx = 0
            header_idx = max(0, min(header_idx, max(0, len(table) - 1)))

            rows = [list(r) for r in table]
            if header_idx == 0 or len(rows) <= 1:
                header = rows[0] if rows else []
                body = rows[1:] if len(rows) > 1 else []
            else:
                header = rows[header_idx]
                body = rows[:header_idx] + rows[header_idx + 1:]

            # pdfplumber 对同一张表里每行返回的列数可能不一致(合并单元格尤其明显),
            # 这里按"表头行 / 数据行"里最宽的一列统一对齐,否则
            # pd.DataFrame(body, columns=header) 会抛
            # "N columns passed, passed data had M columns"。
            width = max([len(header)] + [len(r) for r in body]) if body else len(header)
            width = max(width, 1)

            # 真实内容判定:表头或表体任意一格有文字就算"有内容"。
            # 注意必须在回填 "列N" 占位名之前判断,否则空表会被误判为有内容,
            # 最终写出一个只有列名没有数据的空 sheet。
            has_content = any(_cell_text(c) for c in header) or any(
                any(_cell_text(c) for c in r) for r in body
            )
            if not has_content:
                continue

            # 表头清洗:None/空白 → "列N",并去重(Excel 不允许重名列)
            seen: dict[str, int] = {}
            header_clean: list[str] = []
            for j in range(width):
                name = _cell_text(header[j]) if j < len(header) else ""
                if not name:
                    name = f"列{j + 1}"
                if name in seen:  # 重名追加序号,避免 openpyxl 报 duplicate column
                    seen[name] += 1
                    name = f"{name}_{seen[name]}"
                else:
                    seen[name] = 1
                header_clean.append(name)

            body_clean = [
                [(_cell_text(r[j]) if j < len(r) else "") for j in range(width)]
                for r in body
            ]
            # text 降级策略会按文字位置推断行边界,行与行之间常被插入整行空白,
            # 直接写进 Excel 会出现大片空行,这里统一剔除
            body_clean = [r for r in body_clean if any(r)]

            df = pd.DataFrame(body_clean, columns=header_clean)
            # sheet 名:表N + 简短的 header 提示
            short = "".join(str(c) for c in header_clean[:2])[:8] or ""
            sheet_raw = f"表{i+1}" + (f"_{short}" if short else "")
            sheet = _safe_sheet_name(sheet_raw, used_names)
            pending.append((i, sheet, df))
            sheet_summaries.append({"index": i, "sheet": sheet, "rows": len(df), "cols": len(df.columns)})

        # 必须在开 writer 之前拦截,否则空工作簿会让 openpyxl 抛英文异常
        if not pending:
            raise ValueError("PDF 中表格均为空,无可写入内容")

        # ===== 第二步：确认至少有一张表可写,再统一落盘 =====
        with pd.ExcelWriter(str(out_path), engine="openpyxl") as writer:
            for _i, sheet, df in pending:
                df.to_excel(writer, sheet_name=sheet, index=False)

        return {
            "file_path": str(out_path),
            "file_name": out_name,
            "original_size": int(len(pdf_bytes)),
            "output_size": int(out_path.stat().st_size),
            "table_count": len(sheet_summaries),
            "sheets": sheet_summaries,
            "ai_used": bool(hints_map),
        }
    except ValueError:
        raise
    except Exception as exc:
        logger.exception("pdf_to_excel 失败: %s", exc)
        raise ValueError(f"PDF 转表格失败:{exc}")


def pdf_extract_text(
    pdf_bytes: bytes,
    original_name: str,
    *,
    cleaned_text: str | None = None,
    ai_fallback: bool = False,
) -> dict[str, Any]:
    """PDF 提取文字,可选返回 AI 清洗后的文本(cleaned_text)。

    cleaned_text 为 None 时直接用 pdfplumber 抽出的 raw 全文。
    """
    if not pdf_bytes:
        raise ValueError("PDF 内容为空")
    try:
        page_texts: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page_count = len(pdf.pages)
            for page in pdf.pages:
                txt = page.extract_text() or ""
                if txt:
                    page_texts.append(txt)
        raw_text = "\n\n".join(page_texts)
        # 优先用 AI 清洗结果
        final_text = (cleaned_text or "").strip() or raw_text
        char_count = len(final_text)

        out_name = _safe_output_name(PREFIX_EXTRACT, original_name, "txt")
        out_path = settings.upload_dir / out_name
        settings.upload_dir.mkdir(parents=True, exist_ok=True)
        out_path.write_text(final_text, encoding="utf-8")

        return {
            "text": final_text,
            "char_count": int(char_count),
            "page_count": int(page_count),
            "raw_char_count": int(len(raw_text)),
            "file_path": str(out_path),
            "file_name": out_name,
            "ai_fallback": bool(ai_fallback),
        }
    except ValueError:
        raise
    except Exception as exc:
        logger.exception("pdf_extract_text 失败: %s", exc)
        raise ValueError(f"PDF 提取文字失败:{exc}")


def pdf_compress(pdf_bytes: bytes, original_name: str) -> dict[str, Any]:
    """PDF 压缩(pikepdf + linearize + object_stream_mode=generate)。"""
    if not pdf_bytes:
        raise ValueError("PDF 内容为空")
    try:
        out_name = _safe_output_name(PREFIX_COMPRESS, original_name, "pdf")
        out_path = settings.upload_dir / out_name
        settings.upload_dir.mkdir(parents=True, exist_ok=True)
        with pikepdf.open(io.BytesIO(pdf_bytes)) as pdf:
            pdf.save(
                str(out_path),
                linearize=True,
                compress_streams=True,
                object_stream_mode=pikepdf.ObjectStreamMode.generate,
            )
        original_size = int(len(pdf_bytes))
        output_size = int(out_path.stat().st_size)
        # 压缩比 = (1 - output/original) * 100,允许负值(已优化过)
        ratio = round((1 - output_size / original_size) * 100, 2) if original_size else 0.0
        return {
            "file_path": str(out_path),
            "file_name": out_name,
            "original_size": original_size,
            "output_size": output_size,
            "compression_ratio": float(ratio),
        }
    except ValueError:
        raise
    except Exception as exc:
        logger.exception("pdf_compress 失败: %s", exc)
        raise ValueError(f"PDF 压缩失败:{exc}")


# ============================================================
# 多张照片 → PDF(带 OCR 文字层)
# ============================================================

# 单次上传图片数上限:再多会让单次请求体积与耗时不可控
_MAX_IMAGES = 20
# 单张图片字节上限,防止超大原图把内存打满
_MAX_IMAGE_BYTES = 20 * 1024 * 1024
# 图片长边像素上限:见 _prepare_image 的注释(视觉 token 成本 vs A4 可辨性平衡)
_IMG_MAX_EDGE = 1600
# A4 边距,与 exporter.export_pdf 保持一致
_A4_MARGIN = 40
# platypus Frame 默认每边 6pt 内边距,flowable 真正能放内容的区域 =
# A4 - 2*边距 - 2*该内边距。忽略它会让图片刚好超框 12pt,build() 直接抛
# LayoutError("Flowable ... too large on page"),这是 reportlab 的隐藏坑。
_FRAME_PADDING = 6

# 注册中文字体(ReportLab 自带 CID 字体,无需外部 ttf);与 exporter.py 同一套方案。
# 加一层 getRegisteredFontNames 判断,避免本模块被重复 import 时重复注册。
if "STSong-Light" not in pdfmetrics.getRegisteredFontNames():
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))


def _prepare_image(image_bytes: bytes) -> "Image.Image":
    """把上传的原始图片字节解码成可用于排版的 RGB PIL 对象。

    三步处理,每一步都对应一类真实踩坑:

    1. **EXIF 方向校正**(必须做)
       手机拍照只在 EXIF 里写 Orientation 标记,像素本身还是"躺着"的。
       如果不校正就直接画进 PDF,横拍的手机照片会在 PDF 里旋转 90°,
       而且 PIL 解码时不会替你转,这是照片转 PDF 最典型的翻车点。

    2. **等比缩放到长边 ≤1600px**
       视觉大模型按图片面积/像素数计费(可粗略理解为按 token 数计费),
       手机原图动辄 4000×3000,直接送 OCR 会让成本和耗时成倍上升;
       而本 PDF 最终打印在 A4 上,A4 可用宽仅约 515pt(≈ 1800px @254dpi),
       1600px 已经覆盖 A4 的有效分辨率,再大只是白白增加体积。

    3. **统一转 RGB**(必须做)
       PNG 常见 RGBA(带 alpha 透明通道)或调色板模式 P,
       reportlab 画图只认 RGB/RGBA/L/CMYK,其它模式会直接抛异常;
       且 RGBA 的 alpha 语义在 PDF 里没有对应物,直接画会出现黑底。
    """
    if not image_bytes:
        raise ValueError("图片无法解码,请确认文件是有效的图片")
    try:
        img = Image.open(io.BytesIO(image_bytes))
        # exif_transpose 会把 Orientation 信息"烧"进像素并清除该标记,
        # 等价于人工把照片摆正;返回 None 的情况(无 EXIF)极少,兜底用原图。
        img = ImageOps.exif_transpose(img) or img
        # thumbnail 是"原地等比缩放",只会缩小不会放大,不会像 resize 那样失真拉伸
        img.thumbnail((_IMG_MAX_EDGE, _IMG_MAX_EDGE), Image.LANCZOS)
        if img.mode not in ("RGB", "RGBA", "L", "CMYK"):
            img = img.convert("RGB")
        return img
    except ValueError:
        raise
    except Exception as exc:
        # 损坏文件 / 空字节都会走到这里:统一转成中文提示,交路由转 400
        logger.warning("图片预处理失败: %s", exc)
        raise ValueError("图片无法解码,请确认文件是有效的图片")


def _fit_box(img_w: int, img_h: int, max_w: float, max_h: float) -> tuple[float, float]:
    """按可用区域等比缩放并居中,返回 (宽, 高);不放大超过原始像素尺寸。"""
    scale = min(max_w / img_w, max_h / img_h, 1.0)
    return img_w * scale, img_h * scale


def _flowable_box() -> tuple[float, float]:
    """platypus 版式下 flowable 真正可用的 (宽, 高) = A4 - 2*边距 - 2*Frame 内边距。"""
    return (
        A4[0] - 2 * _A4_MARGIN - 2 * _FRAME_PADDING,
        A4[1] - 2 * _A4_MARGIN - 2 * _FRAME_PADDING,
    )


def _gen_image_pdf(out_path: Path, pil_images: list["Image.Image"]) -> None:
    """图片型 PDF:每张图一页,纯图像,不含任何文字层。"""
    max_w, max_h = _flowable_box()
    story: list = []
    for i, img in enumerate(pil_images):
        w, h = _fit_box(img.width, img.height, max_w, max_h)
        # platypus 的 RLImage 只接受 文件路径 / 文件对象 / 带 _renderPy 的绘图对象,
        # 直接传 PIL 对象会抛 TypeError(expected str, bytes or os.PathLike object)。
        # 因此这里先在内存里编码成 PNG bytes 再包成 BytesIO(不落临时文件)。
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        story.append(RLImage(buf, width=w, height=h))
        # PageBreak 只插在两张图之间:插在最后一张后面会多出一个空白页
        if i < len(pil_images) - 1:
            story.append(PageBreak())
    doc = SimpleDocTemplate(
        str(out_path), pagesize=A4,
        leftMargin=_A4_MARGIN, rightMargin=_A4_MARGIN,
        topMargin=_A4_MARGIN, bottomMargin=_A4_MARGIN,
    )
    doc.build(story)


def _gen_dual_pdf(
    out_path: Path,
    pil_images: list["Image.Image"],
    ocr_texts: list[str | None],
    ocr_layouts: list[dict | None] | None = None,
) -> None:
    """双层 PDF = **Word 风格重排**(优先) 或 白底净化图(降级兜底)。

    【2026-10 重写 · 关键认知】
        用户最初看到的是"照片版 PDF",要求"文字扶正、正向横排、白底黑字、
        排版整齐,像 Word 导出的合同"。上一版做的是"白底化 + 保留印章",
        结果仍然翻车 —— 因为照片里的文字是**印在像素上的既成事实**:
            - 拍摄时的轻微倾斜 → 文字跟着歪,图像处理只能做有限矫正;
            - 页面外的笔记本边缘、桌面、笔、订书钉 → 被当成"内容"二值化成黑块;
            - 原图的字号/间距本来就参差不齐 → 永远排不整齐。
        只有**彻底丢弃照片、用 OCR 文字重新排版**才能真正满足诉求:
        每个字的位置由排版引擎实时计算,天然正向横排、字号统一、间距规整。

    两条路径:
        1. **有版面数据(ocr_layouts 非空)** → 走 _reflow_layout_pdf 重排。
           文字全部由 Paragraph 引擎排出,并把印章抠出来回贴到签署区之后。
        2. **无版面数据 / 重排失败** → 退回"白底净化图 + 不可见文字层"
           (旧行为)。此时至少保证文字可搜索可复制,不会让整次请求失败。

    诚实的能力边界:
        1. 重排后的文字**不是图像**,是真正的 PDF 文本对象 ——
           可搜索、可复制、可在 WPS/Word 里直接编辑,这是相对旧方案的最大提升。
        2. 印章只能按**内容语义**盖在签署区之后,无法还原原图上的精确坐标
           (全文已重排,原坐标已不存在)。
        3. 表格、多栏排版会被模型降级成顺序文本流,本形态不还原表格框线。
    """
    # ---------- 路径 1:结构化版面重排(首选) ----------
    if ocr_layouts:
        try:
            pages = [lay for lay in ocr_layouts if isinstance(lay, dict) and lay.get("blocks")]
            if pages:
                # 逐页抠**全部**印章,并把模型给的 anchor(这枚章盖在哪行文字旁)
                # 分配给对应的那一枚。
                #
                # 分配依据:印章在页面上的**垂直位置**与 anchor 文字在
                # blocks 序列中的**相对顺序**。具体做法是
                # "按 cy 从上到下"排序章,再按"anchor 在原图文字序列中
                # 出现的位置"排序 anchor 描述,一一对应 ——
                # 一页上有两枚章(甲方处 + 落款处)时,这样能把章与它
                # 真正所属的文字对上,而不是全部堆到文末。
                seals: list[list[dict[str, Any]]] = []
                for i, page in enumerate(pages):
                    try:
                        page_seals = (
                            extract_seals_for_reflow(pil_images[i])
                            if i < len(pil_images) else []
                        )
                    except Exception:  # pragma: no cover
                        page_seals = []
                    _assign_seal_anchors(page_seals, page)
                    seals.append(page_seals)
                _reflow_layout_pdf(out_path, pages, seals)
                return
            logger.warning("版面数据为空,双层 PDF 退回白底图方案")
        except Exception as exc:
            # 重排失败不能拖垮整次请求,退回旧方案
            logger.warning("版面重排失败,双层 PDF 退回白底图方案: %s", exc)

    # ---------- 路径 2:白底净化图 + 不可见文字层(降级兜底) ----------
    page_w, page_h = A4
    max_w = page_w - 2 * _A4_MARGIN
    max_h = page_h - 2 * _A4_MARGIN
    # 不可见文字层的排版参数:10pt + 14pt 行距,保证 A4 下一页能放 ~50 行
    font_size, leading = 10, 14
    # 单行最大字符数(按中文全角 10pt 字号估算,A4 可用宽约 515pt ≈ 50 个全角字);
    # 超长行会被切断成多段,避免文字溢出页面右边界
    max_chars_per_line = 48

    # ---------- 先把每张"照片"净化成白底文档图 ----------
    # 印章 / 手写区域在此阶段被检出并保留彩色,其余部分二值化成黑字白底。
    # 单张净化失败不能拖垮整批:退回该图原样(退化成旧的照片版行为)。
    cleaned_images: list["Image.Image"] = []
    for img in pil_images:
        try:
            keep = _detect_seal_regions(img)
            keep += _extract_handwriting_regions(img)
            cleaned_images.append(_clean_document_background(img, keep))
        except Exception as exc:  # pragma: no cover - 防御性降级
            logger.warning("白底化失败,该页退回原图: %s", exc)
            cleaned_images.append(img)

    c = canvas.Canvas(str(out_path), pagesize=A4)
    for idx, img in enumerate(cleaned_images):
        # ---------- 底层:白底净化图,等比缩放并水平垂直居中 ----------
        w, h = _fit_box(img.width, img.height, max_w, max_h)
        # canvas.drawImage 通过 ImageReader 接受 PIL 对象(与 platypus 不同),
        # 省掉一次 PNG 编码
        c.drawImage(
            ImageReader(img),
            (page_w - w) / 2, (page_h - h) / 2,
            width=w, height=h,
        )

        # ---------- 上层:不可见文字层(OCR 识别结果) ----------
        text = (ocr_texts[idx] if idx < len(ocr_texts) else None) or ""
        # OCR 失败/空结果的图只放底图,不放文字层,避免写入无意义的空文本
        if text.strip():
            t = c.beginText()
            t.setTextRenderMode(3)  # 3 = invisible:仅用于检索/复制,不显示也不打印
            t.setFont("STSong-Light", font_size)
            t.setLeading(leading)
            # 从页面顶部往下排;y 低于下边距就停,防止文字画到页面外
            y = page_h - _A4_MARGIN - font_size
            t.setTextOrigin(_A4_MARGIN, y)
            for raw_line in text.splitlines():
                line = raw_line.strip()
                if not line:
                    # 空行也要保留:段落结构本身就是识别结果的一部分,
                    # 抹掉会让还原出来的文本黏成一坨
                    t.textLine(" ")
                    y -= leading
                    continue
                for seg in _split_long_line(line, max_chars_per_line):
                    t.textLine(seg)
                    y -= leading
                    if y < _A4_MARGIN:
                        break
                if y < _A4_MARGIN:
                    break
            c.drawText(t)
        c.showPage()
    c.save()


# ============================================================
# 照片 → Word 风格重排 PDF(2026-10 新增,替代"修补照片"路线)
# ============================================================
#
# 【为什么必须重排,而不是继续修补照片】
#   既有路线(_gen_dual_pdf / _gen_clean_pdf)是"在照片上做减法":
#   去底色 → 二值化 → 白底黑字 → 叠不可见文字层。它有三条**原理上**的死结:
#     1. 非文档物体无法被剔除。拍合同时的笔记本边缘、桌上的绿笔、书本线条,
#        在灰度空间里和文字一样"暗",二值化会原样留下 —— 用户截图里红框标出的
#        正是这些。它们不是"阈值调得不够狠",而是**本来就不是文字**,
#        图像处理无法区分"深色的桌面物体"和"深色的字"。
#     2. 文字的倾斜是几何事实,不是像素问题。照片里这一行字是斜的,
#        任何滤波/阈值都改不了"它斜着"这件事;整页透视重采样能扶正,
#        但扶正后仍然是一张"被拍过的照片",字号依然大小不一、间距依然不齐。
#     3. Word 的排版节奏(统一字号、行距、首行缩进)根本不存在于照片里,
#        任何图像处理都**无法凭空造出**它。
#   所以本路线改为:图像只用来"抠印章",文字全部由版面 JSON 喂给 reportlab
#   从零排版 —— 排版引擎画出来的文字天生正向横排、字号统一、自动换行,
#   背景杂物则被整体丢弃(不进 PDF)。
#
# 【能力边界(诚实说明)】
#   1. 印章位置是**推断**的,不是原坐标还原:版面 JSON 明确不输出坐标,
#      排版引擎又把全文重排了一遍,原图里"章压在第 3 段上"这种几何关系
#      已经不存在了。我们只能把章放在该页签署区之后 —— 这是内容语义的推断,
#      而非像素位置的还原。
#   2. 透视矫正只在"文档边界清晰"的图上有效。复杂背景(桌面花纹、
#      文档压在别的纸上、深色背景)下 Canny 找出的最大四边形可能根本不是文档,
#      此时返回 None,直接退回"不裁剪不矫正" —— 因为错裁比不裁更糟。
#   3. 表格、多栏排版本轮不支持(见 ai_service._OCR_LAYOUT_BLOCK_TYPES 的说明),
#      表格内容会被模型降级成 paragraph 顺序文本流。

# 重排专用边距(60pt ≈ 2.1cm):比 _A4_MARGIN(40)宽。
# 因为重排的目标是"像 Word 导出的合同",Word 默认页边距本身就比
# 这套"图塞满页面"的紧凑排版宽,直接用 40pt 会让文字贴边、看着不像文档。
_REFLOW_MARGIN = 60
# 标题字号 18pt + 标题下方 18pt 留白
_REFLOW_TITLE_SIZE = 18
_REFLOW_TITLE_GAP = 18
# 正文 11pt / 行距 18pt / 段后 8pt:11pt + 18pt 行距是中文合同类文书的
# 常见节奏,与 _gen_text_pdf 保持一致,避免同一套服务里两种 PDF 字号跳变。
_REFLOW_BODY_SIZE = 11
_REFLOW_BODY_LEADING = 18
_REFLOW_PARA_GAP = 8
# 印章回贴后的目标高度(pt),约 3.2cm:与真实公章在纸上的观感接近,
# 太大会在 A4 上喧宾夺主,太小则盖不住签署区的视觉分量。
_REFLOW_SEAL_HEIGHT = 90

# 正文排版用的字体族名。中文必须显式注册 font family 才能让 <b> 真的加粗:
# STSong-Light 是 reportlab 自带的 CID 字体,本身**只有单字重**,
# 不注册 family 时 <b> 会静默回退成"不加粗"(不报错,但用户看不到加粗)。
_REFLOW_FONT = "STSong-Light"
_REFLOW_FONT_BOLD = "CJKBoldReflow"
# macOS 自带的中文粗体候选(按顺序取第一个可用)。
# 为什么用系统字体而不是内嵌文件:reportlab 的 TTFont 需要一个**真实路径**。
# 注意 PingFang / Hiragino 是 CFF(PostScript)轮廓,TTFont 解析不了,
# 注册会直接抛 "TTFError: postscript outlines are not supported";
# STHeiti Medium.ttc 是 TrueType 轮廓且路径长期稳定,所以排第一。
# 全部候选都失败时(如部署到 Linux)退回"不注册粗体",
# 此时 <b> 静默不加粗 —— 版面依然正常,只是标题不够粗。
_REFLOW_BOLD_FONT_CANDIDATES = (
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
)


def _ensure_reflow_fonts() -> None:
    """注册重排用的"中文字体族"(STSong-Light 常规 + 一个真粗体),幂等。

    为什么必须有粗体:
        标题与 kv 标签要**加粗**才能像 Word 文档。但 STSong-Light 是单字重
        CID 字体,reportlab 在没有注册 font family 时遇到 <b> 会**静默**按
        常规字重渲染 —— 不报错,于是"加粗"看起来写上了却毫无效果。
        这里显式注册一个真粗体并挂到 family 上,加粗才真正生效。

    为什么整个过程要用 try/except 包住:
        字体路径是**运行环境相关**的,写死的路径在别的机器上必然不存在。
        注册失败只影响"加粗"这一个视觉细节,不能让整个 PDF 生成挂掉,
        所以这里吞掉异常,退回"标题不加粗"的降级版。
    """
    if _REFLOW_FONT_BOLD in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFontFamily(
            _REFLOW_FONT, normal=_REFLOW_FONT, bold=_REFLOW_FONT_BOLD,
            italic=_REFLOW_FONT, boldItalic=_REFLOW_FONT_BOLD,
        )
        return

    # 函数内 import:TTFont 只在这里用一次,避免模块加载期就因为
    # "系统字体不存在"导致整个 pdf_service 导入失败。
    from reportlab.pdfbase.ttfonts import TTFont

    for path in _REFLOW_BOLD_FONT_CANDIDATES:
        if not Path(path).is_file():
            continue
        try:
            # subfontIndex:TTC(字体集合)里第几个字体。
            # STHeiti Medium.ttc 里第 0 个就是 "Medium" 那一款;
            # 漏掉这个参数会报 "TTFError: subfontIndex required for TTC"。
            pdfmetrics.registerFont(TTFont(_REFLOW_FONT_BOLD, path, subfontIndex=0))
            pdfmetrics.registerFontFamily(
                _REFLOW_FONT, normal=_REFLOW_FONT, bold=_REFLOW_FONT_BOLD,
                italic=_REFLOW_FONT, boldItalic=_REFLOW_FONT_BOLD,
            )
            return
        except Exception as exc:  # pragma: no cover - 环境相关
            logger.warning("粗体字体注册失败(%s),标题将不加粗: %s", path, exc)


def _detect_page_corners(img: "Image.Image") -> "np.ndarray | None":
    """检测照片里**文档区域**的四角,返回 4x2 float32 数组 [[x,y] × 4];检不到返回 None。

    存在的意义:
        手机拍合同通常是"斜着拍"的 —— 文档在画面里是一个**梯形/四边形**。
        本函数的产物有两个用途:
            1. 让 _crop_document 用 getPerspectiveTransform + warpPerspective
               把这个四边形**拉正**成矩形(透视矫正),后续的印章抠图
               就建立在"正的文档"上,检测更稳;
            2. 输出**原图坐标系**下的四角,调用方据此把矫正图的坐标
               映射回原图。

    思路(Canny + 形态学闭运算 + 最大轮廓):
        1. 灰度 + 高斯模糊:模糊让 Canny 阈值稳定,减少噪点边;
        2. Canny 边缘:文档与桌面的交界本身就是最强的一条边;
        3. 闭运算 + 膨胀:文档轮廓常被文字、印章、断开的边框切成好几段,
           连不起来的话,最大的连通块可能只是"文档的半张脸";
        4. 找四边形:approxPolyDP 逼近到 4 点最理想;逼近不到 4 点时
           退化成 minAreaRect(外接矩形)凑一个可用的四边形。

    **为什么失败可以安全退回 None(而不是强行猜一个框)**:
        文档边界检测本质是启发式,在这些场景必然失败:
            - 纯扫描件 / 白底图:整幅图没有"边界"这条边,Canny 找不到闭合轮廓;
            - 复杂背景(桌面花纹、压在另一叠纸上):最大轮廓是**背景**不是文档;
            - 文档只露出半张、或者被手遮住一角:四角不全,拉正必然变形。
        此时硬猜一个框的后果是"把整张图裁掉一半"或"把文档拉伸变形",
        比不处理更糟。而重排路线的**主体是文字**(由版面 JSON 提供排版),
        图像只用来抠印章 —— 裁剪矫正只是锦上添花,不做也能得到正确的正文。
        所以这里一律"检测失败就返回 None,调用方退回不裁剪",
        把不确定性收敛成"最坏情况=没有优化",而不是"最坏情况=内容被裁坏"。
    """
    try:
        rgb = np.asarray(img.convert("RGB"))
        h, w = rgb.shape[:2]
        if h < 40 or w < 40:
            # 太小的图不可能有可用的四角轮廓,直接失败,
            # 避免后续 getPerspectiveTransform 拿到退化数据后产出废图。
            return None

        gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
        # 【踩过的坑·必须用小核模糊】模糊核写死 5,不能跟图像尺寸挂钩。
        #   实测:按短边 2% 取核(1600px 图上是 19px)会把文档的 1px 描边
        #   直接抹平 —— Canny 找不到"白纸 vs 深色桌面"这条分界,整图 0 轮廓。
        #   文档/桌面本就是**高对比硬边**,不需要强模糊,小核只用来压噪点。
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        # 【踩过的坑·Canny 阈值必须用 Otsu,不能用"均值 × 比例"】
        #   均值法的前提是"前景与背景各占一半",而实拍照片里白纸通常占 60%+
        #   (实测某张 900×1100 的样本均值 185),阈值被抬到 122~247,
        #   高于纸张与桌面的实际灰度差(248 vs 113),结果同样是 0 轮廓。
        #   Otsu 是按"直方图双峰"自动分界,不受前景/背景比例影响,
        #   白纸 248 / 桌面 113 这种双峰能稳定分开。
        otsu, _ = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        low = int(max(10, otsu * 0.3))
        high = int(min(255, max(otsu * 1.4, low + 20)))
        edges = cv2.Canny(blurred, low, high)

        # 闭运算 + 膨胀:把被文字/印章打断的文档外轮廓连成一个整体。
        # 核跟图像尺寸挂钩是对的 —— 连接"断开的边"需要跨越版面元素的量级。
        ck = max(3, int(min(h, w) * 0.015) | 1)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (ck, ck))
        edges = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, kernel, iterations=2)
        edges = cv2.dilate(edges, kernel, iterations=1)

        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return None
        # 只信"足够大"的轮廓:小于图面积 20% 的直接排除,
        # 否则桌面上一个深色物件的轮廓就可能被当作文档。
        min_area = h * w * 0.20
        quad: "np.ndarray | None" = None
        for cnt in sorted(contours, key=cv2.contourArea, reverse=True)[:5]:
            if cv2.contourArea(cnt) < min_area:
                continue
            # 【关键·先取凸包再自适应逼近】直接对原轮廓 approxPolyDP
            #   在"文档压在别的纸上/边缘被切断"的图上会逼近到 8~10 个点,
            #   直接判负就退化成 minAreaRect(正交外接矩形),
            #   把倾斜的文档"压"成轴对齐矩形 —— 透视矫正等于没做。
            #   先 convexHull 去掉因闭运算产生的锯齿与小凸起,
            #   再从 2% 起逐步放大 epsilon 找**第一个**逼近到 4 点的值,
            #   实测 4% 就能把样本的文档轮廓逼近到真实四角(误差 ~5px)。
            hull = cv2.convexHull(cnt)
            peri = cv2.arcLength(cnt, True)
            for eps_ratio in (0.02, 0.03, 0.04, 0.05, 0.06, 0.08):
                approx = cv2.approxPolyDP(hull, eps_ratio * peri, True)
                if len(approx) == 4 and cv2.isContourConvex(approx):
                    quad = approx.reshape(4, 2).astype(np.float32)
                    break
            if quad is None:
                # 退化路径:连凸包都逼近不到 4 点(文档被遮住一角、边缘毛糙)时,
                # 用最小外接矩形凑一个四边形 —— 至少保证有输出,
                # 后续 _crop_document 还有一道面积合理性检查兜底。
                quad = cv2.boxPoints(cv2.minAreaRect(cnt)).astype(np.float32)
            if quad is not None:
                break

        if quad is None:
            return None
        # 面积合理性检查:挡掉"四个角点全部落在同一条边上"这类畸形结果。
        if cv2.contourArea(quad.astype(np.int32)) < min_area:
            return None

        # 角点顺序必须固定为 左上→右上→右下→左下,否则 getPerspectiveTransform
        # 会把页面转 90° 或镜像。排序技巧:按 "x+y"(左上最小)、"x-y"
        # (右上最小)、"-(x+y)"(右下最大)、"y-x"(左下最大)四组极值取点,
        # 这是处理任意凸四边形角点顺序最稳的写法(依赖四边形大致为矩形,
        # 而拍文档时这一定成立)。
        pts = quad.reshape(4, 2).astype(np.float32)
        s = pts.sum(axis=1)
        d = pts[:, 0] - pts[:, 1]
        ordered = np.array(
            [
                pts[np.argmin(s)],   # 左上
                pts[np.argmin(d)],   # 右上
                pts[np.argmax(s)],   # 右下
                pts[np.argmax(d)],   # 左下
            ],
            dtype=np.float32,
        )
        return ordered
    except Exception as exc:  # pragma: no cover - 启发式链路长,纯防御
        # 边界检测是"锦上添花"的优化,任何异常都直接退回不裁剪,
        # 绝不让它影响重排主流程(正文由版面 JSON 提供,不依赖图像)。
        logger.warning("文档角点检测失败,退回不裁剪不矫正: %s", exc)
        return None


def _crop_document(
    img: "Image.Image",
) -> tuple["Image.Image", tuple[int, int, int, int] | None]:
    """把照片里的文档区域裁出来并**透视矫正**成矩形。

    与 :func:`_detect_page_corners` 的关系:
        本函数是它唯一的消费者,负责"检测到四角后真正执行变换",
        并把矫正后的图与**原图中的 bbox** 一起返回。

    为什么必须同时返回原图 bbox:
        矫正图与原图是**两个不同的坐标系**。调用方后续要在原图上抠印章,
        只能靠这个 bbox 把"矫正后的区域"对应回"原图的哪一块";
        只返回矫正图就丢掉了这个映射关系,要么无从下手,要么只能靠
        坐标比例硬猜(遇透视时必错)。

    Args:
        img: RGB 的 PIL Image(通常来自 _prepare_image)。

    Returns:
        ``(矫正后的图, 原图中的 bbox)``;检测失败时返回 ``(原图, None)``。
    """
    try:
        corners = _detect_page_corners(img)
        if corners is None:
            return (img, None)

        w, h = img.size
        # 输出尺寸用四角构成的边长估算。这里**刻意不追求 A4 比例** ——
        # 裁出来的图还要去抠印章,强行拉伸到 A4 比例会把印章一起拉变形。
        # 保持各边大致原长即可。
        out_w = int(max(
            float(np.linalg.norm(corners[1] - corners[0])),
            float(np.linalg.norm(corners[2] - corners[3])),
        ))
        out_h = int(max(
            float(np.linalg.norm(corners[3] - corners[0])),
            float(np.linalg.norm(corners[2] - corners[1])),
        ))
        # 太小/太大的输出尺寸会让 warpPerspective 产出退化图,直接放弃矫正。
        if out_w < 40 or out_h < 40:
            return (img, None)
        out_w = min(out_w, w * 2)
        out_h = min(out_h, h * 2)

        # 目标矩形的四角(左上→右上→右下→左下),顺序与 corners 严格对应。
        # getPerspectiveTransform 的 src/dst 必须**逐点一一对应**,
        # 顺序错一个点,出来的就是转置/镜像的废图。
        dst = np.array(
            [[0, 0], [out_w - 1, 0], [out_w - 1, out_h - 1], [0, out_h - 1]],
            dtype=np.float32,
        )
        warped = cv2.warpPerspective(
            np.asarray(img.convert("RGB")),
            cv2.getPerspectiveTransform(corners, dst),
            (out_w, out_h),
            flags=cv2.INTER_LINEAR,
        )
        # 原图 bbox:四角点的外接矩形,夹在图内。
        x1 = int(max(0, np.floor(corners[:, 0].min())))
        y1 = int(max(0, np.floor(corners[:, 1].min())))
        x2 = int(min(w, np.ceil(corners[:, 0].max())))
        y2 = int(min(h, np.ceil(corners[:, 1].max())))
        if x2 <= x1 or y2 <= y1:
            return (img, None)
        return (_np_to_pil(warped), (x1, y1, x2, y2))
    except Exception as exc:  # pragma: no cover - 纯防御
        logger.warning("文档裁剪/矫正失败,退回原图: %s", exc)
        return (img, None)


def _keep_blanks(text: str) -> str:
    """把文本里连续 2 个以上的空白**渲染成下划线**,保住原文的"留白栏"观感。

    【踩过的坑:U+00A0 / U+3000 在 reportlab 里根本画不出来】
        最初的想法是"把空格换成不换行空格( U+00A0 )或全角空格( U+3000 )",
        实测两者在 ``reportlab.platypus.Paragraph`` 里**都不产生任何可见宽度**:
        PDF 里画出的字符坐标显示「日期:」与「年」之间 x 间隔为 0,
        抽出的文本是「日期:年月日」—— 留白被完整吃掉了。
        原因在 STSong-Light 这个 CID 字体的字形映射上:
        它的 ToUnicode/CMap 不含空白类码位,reportlab 查不到字形就
        **静默跳过**(不报错),所以无论用哪个空格码位都一样没宽度。
        唯一能稳定产生可见宽度的是实心字符 —— 下划线 ``_``。

    所以最终方案:把"连续 ≥2 个的空白"折算成等长的下划线。
        - 「日期:      年    月    日」→ 「日期:______年____月____日」
        - 这正是中文合同里"日期:______年__月__日"的真实排版效果,
          比单纯的空白更贴近用户预期(用户填日期时有横线可写)。

    为什么这同时是"不能靠空格对齐"的又一次佐证:
        空格在 reportlab 的 CID 字体里**宽度不可控、甚至为 0**;
        想要"稳定的留白"就只能用实心字符。所以正文对齐一律走流式文本
        (见 _reflow_layout_pdf 的 kv 分支),只有"原文本来就有大段空白"的
        签署区/日期行才在这里被替换成下划线。

    Args:
        text: 已经 escape() 过的文本(本函数不再处理 XML 特殊字符)。

    Returns:
        把长空白替换为下划线后的文本。
    """
    # 先把制表符归一成普通空格,再统一处理"连续空白"这一种形态。
    normalized = text.replace("\t", " ")

    def _sub(match: "re.Match[str]") -> str:
        # 一个空白不处理(那是词内分隔,可能是英文单词之间的空格);
        # 连续 2 个以上才认定为"留白栏",折算成下划线。
        run = match.group(0)
        return run if len(run) < 2 else "_" * len(run)

    return re.sub(r" {2,}", _sub, normalized)


def _pil_to_rl_image(
    pil_img: "Image.Image", target_height: float,
) -> "RLImage | None":
    """把 PIL 图等比缩放到指定高度并包成 RLImage;失败返回 None。

    【踩过的坑:RLImage 不接受 PIL 对象】
        ``reportlab.platypus.Image`` 内部把参数当**路径**去打开,
        直接塞 PIL 对象会抛
        ``TypeError: expected str, bytes or os.PathLike object, got 'Image'``
        —— 项目里 _gen_image_pdf 已经踩过这个坑,改成了
        "先在内存里编码成 PNG bytes → 包 BytesIO → 再交给 RLImage"。
        这里沿用**同一套已验证的写法**,不落任何临时文件。

    等比缩放:
        只按高度定尺寸、宽度按原比例算(印章是正方的,按框拉伸会变形)。
        放大倍数设 3 倍上限:极小的印章(只检测到 20px)放大 4~5 倍
        只会得到一团模糊色块,不如不放。
    """
    try:
        if pil_img is None:
            return None
        w, h = pil_img.size
        if w <= 0 or h <= 0:
            return None
        scale = min(target_height / float(h), 3.0)
        buf = io.BytesIO()
        # convert("RGB"):印章图可能是 RGBA/P,RLImage 编码 PNG 时
        # 遇到 CMYK/带调色板的图会报 "cannot write mode ... as PNG"。
        pil_img.convert("RGB").save(buf, format="PNG")
        buf.seek(0)
        return RLImage(buf, width=w * scale, height=target_height)
    except Exception as exc:  # pragma: no cover - 纯防御
        logger.warning("印章图像转换失败,跳过该印章: %s", exc)
        return None


def _seal_to_pil(
    img: "Image.Image", bbox: tuple[int, int, int, int],
) -> "Image.Image | None":
    """按 bbox 从图上抠出印章区域;bbox 非法/过小返回 None。

    为什么不复用 _clean_document_background:
        那个函数的产物是"整页白底图",这里只需要**一块**印章。
        但单纯 crop 会把章周围的灰色阴影一起抠进来,输出上就是
        "灰色方块里嵌着一个红章",比不抠还难看 —— 所以这里额外做一次
        "除以局部背景"的提亮,把阴影压到接近白底,与 _clean_document_background
        第 4 步的做法保持一致。
    """
    try:
        x1, y1, x2, y2 = (int(v) for v in bbox)
        w, h = img.size
        x1 = max(0, min(x1, w))
        y1 = max(0, min(y1, h))
        x2 = max(0, min(x2, w))
        y2 = max(0, min(y2, h))
        if x2 - x1 < 8 or y2 - y1 < 8:
            # 太小的"印章"多半是噪点或误检,抠出来是一坨马赛克,
            # 丢弃比塞进 PDF 更干净。
            return None
        crop = img.crop((x1, y1, x2, y2)).convert("RGB")
        # 局部背景估计:闭运算把比核小的暗结构(章的笔画)填掉,
        # 剩下的近似纯背景;再按通道做除法归一化,阴影被压到白。
        arr = np.asarray(crop)
        ch, cw = arr.shape[:2]
        k_bg = max(3, (int(min(ch, cw) / 3) | 1))
        bg = cv2.morphologyEx(
            cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY),
            cv2.MORPH_CLOSE,
            cv2.getStructuringElement(cv2.MORPH_RECT, (k_bg, k_bg)),
        )
        bg = np.clip(bg, 1, 255).astype(np.uint8)
        normalized = cv2.divide(arr, cv2.cvtColor(bg, cv2.COLOR_GRAY2BGR), scale=255)

        # 【2026-10 修·去掉章外的灰底方框】
        # 踩过的坑:只做"除以背景"只能把阴影压亮,但章的 bbox 里
        # 那些**本来就不是章**的像素(比如章旁边的正文笔画、纸纹)
        # 依然保留着,贴进白底 PDF 后就是一个明显的灰色方块。
        # 修法:既然这枚章是**红色的**,就把"够红"的像素留下,
        # 其余(章外的灰底/黑字/纸纹)一律置白。
        # 判据用 HSV:R 明显大于 G/B 且饱和度不低 → 是红色印泥;
        # 边界处用形态学膨胀把抗锯齿的淡边一起纳入,避免章被削细。
        hsv = cv2.cvtColor(normalized, cv2.COLOR_RGB2HSV)
        hh, ss, vv = hsv[:, :, 0], hsv[:, :, 1], hsv[:, :, 2]
        rr, gg = normalized[:, :, 0].astype(np.int16), normalized[:, :, 1].astype(np.int16)
        # 红色:色相在 0~35 或 165~180(红色横跨色相环两端),饱和度 ≥ 45
        is_red = ((hh <= 35) | (hh >= 165)) & (ss >= 45) & (vv >= 40)
        # 强红(印泥本体)与弱红(抗锯齿淡边/边缘晕染):
        # 两者分开是为了区分"章的淡边"和"章旁边的灰色投影"——
        # 投影是中性灰(红通道不占优),淡边则是偏红的。
        strong_red = is_red & ((rr - gg) > 30)
        weak_red = is_red & ((rr - gg) > 8)
        # 章外全部置白
        out = normalized.copy()
        out[~weak_red] = 255
        # 闭运算:把章内部断开的笔画连上(公章文字常有空隙)
        strong_u8 = (strong_red.astype(np.uint8)) * 255
        k_cl = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
        strong_u8 = cv2.morphologyEx(strong_u8, cv2.MORPH_CLOSE, k_cl)
        # 只保留"紧挨着强红核心"的弱红像素:
        # 这样章的抗锯齿淡边被留下,而离章较远的灰色投影被剔除,
        # 章外不再出现一圈灰晕。
        k_near = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (9, 9))
        near_core = cv2.dilate(strong_u8, k_near, iterations=1)
        keep = (weak_red & (near_core > 0)) | (strong_u8 > 0)
        out[~keep] = 255
        return _np_to_pil(out)
    except Exception as exc:  # pragma: no cover
        logger.warning("印章抠图失败: %s", exc)
        return None


def _assign_seal_anchors(
    page_seals: list[dict[str, Any]],
    page: dict[str, Any],
) -> None:
    """把模型给的印章 anchor(这枚章盖在哪行文字旁)分配给对应的章。

    【为什么需要"分配"这一步】
        模型输出的 ``seals`` 是**按它自己观察到的顺序**列的
        (通常是从上到下),而 ``extract_seals_for_reflow`` 抠出来的章是
        **按面积从大到小**排序的。两个顺序不一致时,
        直接按下标配对就会把章安到错误的文字旁。

    分配算法:
        1. 把模型给的 seal 描述按 anchor 文字在 blocks 序列中的**出现位置**
           排序(anchor 找不到的排最后);
        2. 把抠出来的章按 **cy(垂直位置)从小到大**排序(即从上到下);
        3. 两者一一对应 —— 页面最上方的章配最上方的 anchor。

    诚实的能力边界:
        这仍然是**基于顺序的推断**,不是几何还原。模型若把 anchor 写错
        (比如两枚章的 anchor 写反了),章会贴到错误的行旁边。
        要做到精确还原必须让模型直接输出 bbox 坐标,
        而视觉模型给坐标的可靠性远不如给文字可靠。
    """
    if not page_seals:
        return
    descs = page.get("seals")
    if not isinstance(descs, list) or not descs:
        # 模型没给 anchor → 全部退回"贴到本页末尾"(由重排引擎处理)
        return

    # blocks 里所有文字按顺序拼起来,用于查 anchor 的位置
    blocks = page.get("blocks")
    blocks = blocks if isinstance(blocks, list) else []
    seq: list[str] = []
    for blk in blocks:
        if not isinstance(blk, dict):
            continue
        for key in ("text", "label", "value"):
            v = blk.get(key)
            if isinstance(v, str) and v.strip():
                seq.append(v.strip())
                break
    joined = "\n".join(seq)

    def anchor_pos(d: dict[str, Any]) -> int:
        a = str(d.get("anchor") or "").strip()
        if not a:
            return len(joined) + 1          # 没 anchor 的排最后
        idx = joined.find(a)
        if idx < 0:
            # 退而求其次:找最长公共前缀
            for cut in range(len(a) - 2, 3, -1):
                idx = joined.find(a[:cut])
                if idx >= 0:
                    break
        return idx if idx >= 0 else len(joined) + 2

    # 模型描述按 anchor 位置排序
    descs_sorted = sorted(
        (d for d in descs if isinstance(d, dict)),
        key=anchor_pos,
    )

    # 【2026-10 新增·数量对齐】
    # 抠出的章比模型报告的还多时,多出来的几乎都是碎片/误检
    # (油墨断连剥落的五角星、被误判的黄色便签等 —— 它们面积
    # 天然比完整章小)。模型报告的 seal 数是可信的语义信息
    # (它"看到"了几枚章),按面积保留前 N 枚,多余的不参与
    # 配对、也不会堆到页面末尾。
    # 反向情况(模型漏报)不截断:宁可多贴也不少贴。
    if 0 < len(descs_sorted) < len(page_seals):
        kept = sorted(
            page_seals, key=lambda d: d.get("area", 0.0), reverse=True,
        )[: len(descs_sorted)]
        logger.info(
            "印章数量对齐: 抠出 %d 枚 > 模型报告 %d 枚,按面积保留前 %d 枚",
            len(page_seals), len(descs_sorted), len(kept),
        )
        page_seals[:] = kept

    # 抠出来的章按 cy 从小到大(从上到下)
    seals_sorted = sorted(page_seals, key=lambda d: d.get("cy", 0.0))

    for i, seal in enumerate(seals_sorted):
        if i >= len(descs_sorted):
            break
        d = descs_sorted[i]
        seal["anchor"] = str(d.get("anchor") or "").strip()
        side = str(d.get("side") or "right").strip().lower()
        seal["side"] = "left" if side == "left" else "right"


def extract_seals_for_reflow(
    pil_image: "Image.Image",
) -> list[dict[str, Any]]:
    """从一张照片里抠出**全部**印章,返回带归一化坐标的列表。

    【2026-10 重写】为什么要从"只取最大的一枚"改成"全部":
        实测用户的合同照片上有**两枚章**(甲方处一枚 + 落款处一枚)。
        旧实现只取面积最大的那枚,于是甲方那枚直接丢失;
        而那一枚"最大"的章又因为统一贴到文末,跑到了跟它毫无关系的
        落款下面 —— 用户看到的就是"章的位置全乱了"。

    坐标为什么用**归一化**值(0~1 的相对比例)而不是像素:
        抠章用的是 _prepare_image 处理过的图(长边 ≤1600px),
        而版面 JSON 里的 anchor 是**文字内容**不是坐标。
        归一化坐标描述的是"这枚章在页面上的相对位置",
        重排后按这个比例去对应段落附近贴放,与图片实际像素尺寸解耦。

    Returns:
        ``[{"image": PIL.Image, "cx": float, "cy": float, "area": float}, ...]``
        按面积从大到小排序;检不到印章时返回空列表。
    """
    try:
        boxes = _detect_seal_regions(pil_image)
        if not boxes:
            return []
        w, h = pil_image.size
        if w <= 0 or h <= 0:
            return []
        out: list[dict[str, Any]] = []
        for box in boxes:
            seal = _seal_to_pil(pil_image, box)
            if seal is None:
                continue
            x1, y1, x2, y2 = box
            out.append({
                "image": seal,
                # 章的中心点(归一化)——重排时按它决定贴在哪一段附近
                "cx": (x1 + x2) / 2.0 / float(w),
                "cy": (y1 + y2) / 2.0 / float(h),
                # 面积占比(归一化)——用于按"章在页面上的相对大小"排序,
                # 让上方的章先于下方的章被贴出
                "area": ((x2 - x1) * (y2 - y1)) / float(w * h),
            })
        # 面积从大到小;同面积时按 cy 从小到大(上方的章先处理)
        out.sort(key=lambda d: (-d["area"], d["cy"]))
        for i, d in enumerate(out):
            d["id"] = i
            d.setdefault("anchor", "")
            d.setdefault("side", "right")
        return out
    except Exception as exc:  # pragma: no cover
        logger.warning("印章抠取失败: %s", exc)
        return []


def extract_seal_for_reflow(pil_image: "Image.Image") -> "Image.Image | None":
    """从一张照片里抠出**最大的那一枚印章**,供重排后回贴;检不到返回 None。

    为什么只取最大的一枚:
        一页合同上通常只盖一个主章。重排路线只保留一枚最显著的印章,理由是:
            1. 排版引擎重排后无法还原多枚章各自的原始位置,盲目回贴多枚
               会全部堆在签署区后面,变成一块看不出所以然的色块;
            2. 保留主章已足够表达"这份文书已签署"这一信息。

    Args:
        pil_image: RGB 的 PIL Image。**建议传 _crop_document 的输出** ——
            矫正后印章更方正,_detect_seal_regions 的"宽高比接近 1"硬门槛
            才不会被拍摄透视破坏,检出率明显更高。
    """
    try:
        boxes = _detect_seal_regions(pil_image)
        if not boxes:
            return None
        # 取面积最大的一枚:面积即视觉显著性,最大通常就是主章。
        best = max(boxes, key=lambda b: (b[2] - b[0]) * (b[3] - b[1]))
        return _seal_to_pil(pil_image, best)
    except Exception as exc:  # pragma: no cover
        logger.warning("印章抠取失败: %s", exc)
        return None


class _ReflowDocTemplate(SimpleDocTemplate):
    """重排专用模板:页眉标注"第几张原图" + 页脚总页码。

    为什么页眉要标"第几张照片":
        一份合同常常分多张照片拍,重排后会被合并进同一个 PDF。
        用户看到连续正文分页时,已经分不清"这是第 2 张照片的延续"。
        在页顶标一个**很淡**的"第 N 张照片"既保留了对照关系,
        又不像正式合同那样在正文里插一行多余说明文字。

    实现要点:
        - reportlab 的每页回调挂在 PageTemplate 的 onPage / onPageEnd 上;
        - "当前是第几张原图"无法从 canvas 直接读到,靠 afterFlowable 随
          flowable 流式记录(每个 flowable 携带 _reflow_src);
        - Frame 的内边距显式压到 0:默认 6pt 会让文字比设定边距再内缩一圈,
          页眉/页码的位置就会跟着偏移,对不齐。
    """

    def __init__(
        self, *args: Any, pages: list[dict[str, Any]] | None = None, **kwargs: Any,
    ) -> None:
        super().__init__(*args, **kwargs)
        self._pages = pages or []
        # 当前正在排第几张原图,由 afterFlowable 随 flowable 流式更新。
        self.current_src: int = 0

    def afterFlowable(self, flowable: Any) -> None:
        """flowable 落版后更新"当前原图序号"(flowable 上带着 _reflow_src)。"""
        src = getattr(flowable, "_reflow_src", None)
        if src is not None:
            self.current_src = int(src)

    def _draw_header(self, canv: Any, _doc: Any) -> None:
        """页眉/页码绘制:很淡的灰色,绝不与合同正文抢视觉权重。"""
        try:
            # 【踩过的坑:两个钩子都挂同一个函数 → 同一页画两遍】
            #   PageTemplate 的 onPage 与 onPageEnd 都会在每页触发,
            #   把 _draw_header 同时挂上去,页眉就被**画了两遍**,
            #   pdfplumber 抽出来是重叠的 '第第  11  张张照照片片'(chars 坐标里
            #   同一位置有两个 '第')。
            #
            # 【踩过的坑·为什么改用 handle_pageEnd 而不是 onPage/onPageEnd】
            #   实测两个 PageTemplate 钩子都不可靠:
            #     - onPage:在本页 flowable 落版**之前**触发,此时 current_src
            #       还是**上一页**的值,页眉会标错照片序号;
            #     - onPageEnd:最后一页不触发(最后一页走 _endBuild 的
            #       clean_hanging 路径,不会进入 handle_pageEnd),
            #       所以"最后一页没有页眉/页码"。
            #   handle_pageEnd 每页都会触发(含最后一页),
            #   且触发在 showPage() 之前、afterFlowable 之后,
            #   current_src 正是本页最后一个 flowable 的来源序号 —— 语义正确。
            canv.saveState()
            canv.setFont(_REFLOW_FONT, 8)
            canv.setFillColorRGB(0.72, 0.72, 0.72)
            # 只有一张原图时不标"第 1 张照片":单页文档加这行纯属噪音。
            if len(self._pages) > 1:
                canv.drawString(
                    _REFLOW_MARGIN, A4[1] - _REFLOW_MARGIN + 18,
                    f"第 {self.current_src + 1} 张照片",
                )
            canv.drawRightString(
                A4[0] - _REFLOW_MARGIN, _REFLOW_MARGIN - 24,
                f"第 {canv.getPageNumber()} 页",
            )
            canv.restoreState()
        except Exception as exc:  # pragma: no cover - 页眉失败不能影响正文
            logger.warning("重排页眉绘制失败: %s", exc)

    def handle_pageEnd(self) -> None:
        """页尾钩子:画页眉/页码,再交给父类完成收尾(详见 _draw_header 的注释)。"""
        self._draw_header(self.canv, self)
        # 必须调用父类:它负责 showPage() / 换页模板 / 无限循环检测,
        # 漏掉会导致整个 PDF 只剩一页(后续页内容被丢弃)。
        super().handle_pageEnd()

    def build(self, flowables: list[Any], *args: Any, **kwargs: Any) -> None:
        """build 前挂上显式 Frame 的 PageTemplate(控制内边距)。"""
        frame = Frame(
            _REFLOW_MARGIN, _REFLOW_MARGIN,
            A4[0] - 2 * _REFLOW_MARGIN, A4[1] - 2 * _REFLOW_MARGIN,
            id="reflow_body",
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        )
        # 不挂 onPage / onPageEnd:页眉统一由 handle_pageEnd 绘制(见上面的注释)。
        tpl = PageTemplate(id="reflow", frames=[frame])
        self.addPageTemplates([tpl])
        super().build(flowables, *args, **kwargs)


def _page_seals_for_index(seals: Any, page_index: int) -> list[dict[str, Any]]:
    """取第 page_index 页对应的印章列表,兼容旧的"每页一个印章"结构。

    为什么要有这层兼容:
        旧版 seals 是 ``list[Image.Image | None]``(每页一个 PIL 图);
        新版是 ``list[list[dict]]``(每页多枚,每枚带 anchor/坐标)。
        两种结构并存期间,这里按元素类型自动判断,避免调用方必须同步改造。
    """
    if not seals:
        return []
    # 新结构:list[list[dict]]
    if isinstance(seals[0], list):
        if page_index < len(seals):
            return [d for d in (seals[page_index] or []) if isinstance(d, dict)]
        return []
    # 旧结构:list[Image.Image | None] —— 包装成单元素列表
    if page_index < len(seals):
        img = seals[page_index]
        if img is not None:
            return [{"image": img, "anchor": "", "side": "right"}]
    return []


def _find_anchor_position(texts: list[str], anchor: str) -> int:
    """在"每行最终会渲染成的文字"列表里找最接近包含 anchor 的那一行下标。

    为什么入参是**纯文本列表**而不是 flowable 列表:
        排版改造后(2026-10 双栏并排),一行文字可能被装进 Table 里,
        flowable 不再是一条 Paragraph —— 再靠 ``getattr(fl, "text")`` 去读
        已经读不到内容。统一让调用方传"这一行最终会渲染成的文字",
        匹配逻辑就与内部结构解耦了。

    Returns:
        匹配到的行下标;找不到返回 -1。
    """
    if not anchor:
        return -1
    target = anchor.strip()
    if not target:
        return -1

    # 【2026-10 修·kv 段落永远匹配不上的 bug】
    # kv 段落在排版时被改写成 ``<b>甲方</b>:宁夏…`` —— 带着 HTML 标签、
    # 且用的是**英文冒号**;而模型 anchor 是原文照抄的「甲方:」(中文冒号)。
    # 直接 ``target in txt`` 必然失败;加上 anchor 往往只有 3~4 个字,
    # 下面的"缩短前缀"循环 (range(len-4, 3, -2)) 对短 anchor 是空序列,
    # 一步退路都没有 —— 最终章被 append 到 story 末尾,流落到最后一页。
    # 修法:匹配前先把文本剥标签;冒号**双向归一**成英文冒号
    # (kv 排版段用英文冒号,模型 anchor 常是中文冒号 U+FF1A,
    # 归一方向选英文是因为排版段那边只有英文冒号)。
    def _plain(txt: str) -> str:
        return _HTML_TAG_RE.sub("", txt).replace("：", ":")

    norm_target = _plain(target)

    # 1) 精确包含(最理想)
    for i, txt in enumerate(texts):
        if isinstance(txt, str) and norm_target in _plain(txt):
            return i
    # 2) 逐级缩短 anchor,取能匹配上的最长前缀
    for cut in range(len(norm_target) - 4, 3, -2):
        frag = norm_target[:cut]
        for i, txt in enumerate(texts):
            if isinstance(txt, str) and frag and frag in _plain(txt):
                return i
    return -1


# ============================================================
# 重排的"行"模型(2026-10 新增):让左右并排的两栏落在同一水平线
# ============================================================
#
# 要解决的问题:
#   旧实现把每个 block 直接 append 进 story,**一块一行、从上往下堆**。
#   而纸质合同里「甲方 | 乙方」「甲方盖章 | 乙方签字」「左日期 | 右日期」
#   本来就是**同一水平行的左右两栏**,按块堆叠后就变成上下两行,
#   用户看到的就是"位置全变了"(截图里甲方在左上、乙方却掉到下一行)。
#
# 解决思路:
#   排版单位从"块"改成"行":一行 = (左块, 右块 或 None)。
#   - 模型若已按 prompt 要求把并排两条合并成一个 block(right_* 字段),
#     直接展开成一行两栏;
#   - 模型没合并时做一次**保守兜底配对**(见 _looks_side_by_side):
#     只对语义上确定是并排的组合生效(signature/date 恒并排;
#     kv 仅当两侧都是缔约主体标签),避免把上下两行的 kv 误配成一行。
#
# 右栏宽度 45%:合同里左侧"甲方:"通常很短、右侧留空写乙方信息更长,
# 左窄右宽比 50/50 更接近原件观感。

_PAIRABLE_TYPES = frozenset({"kv", "signature", "date"})

# 缔约主体标签白名单:只有两侧都是这张表里的标签(且互不相同),
# 才认为它们是"甲方/乙方"这种同一行的左右并排。
# 为什么必须收紧:连续两个 kv 也可能是"甲方:"下面紧接"身份证号:"(上下行),
# 无条件配对会把两行硬压成一行,反而制造新的错位。
_PARTY_LABELS = frozenset({
    "甲方", "乙方", "丙方", "丁方", "买方", "卖方",
    "甲方单位", "乙方单位", "出租方", "承租方",
    "委托方", "受托方", "发包人", "承包人", "出卖人", "买受人",
})

# 模型对无法辨认的内容会写 `[无法辨认]` 这类占位。直接排进 PDF 很难看,
# 换成等宽下划线后与合同里"空栏待填"的观感一致。
_UNCLEAR_MARK_RE = re.compile(
    r"\[\s*(?:无法辨认|无法识别|看不清|辨认不清)\s*\]"
)


def _tidy_unclear(text: str) -> str:
    """把 `[无法辨认]` 一类的识别占位换成下划线空栏。"""
    return _UNCLEAR_MARK_RE.sub("______", text)


def _norm_reflow_block(blk: Any) -> dict[str, str]:
    """把一个 block 规范成排版用字典(缺字段补空串,占位符顺手换成下划线)。

    为什么要再规范化一层:
        _parse_layout_json 已经清洗过一遍,但 `_reflow_layout_pdf` 也会被
        内部单测/旧调用方直接喂原始 dict。统一在这里兜住"字段不是字符串"
        的情况,后面所有分支都只读字符串,不再到处写 isinstance 判断。
    """
    def _s(key: str) -> str:
        v = blk.get(key) if isinstance(blk, dict) else None
        return _tidy_unclear(v.strip()) if isinstance(v, str) else ""

    btype = blk.get("type") if isinstance(blk, dict) else None
    return {
        "type": btype if isinstance(btype, str) else "",
        "text": _s("text"),
        "label": _s("label"),
        "value": _s("value"),
        "right_text": _s("right_text"),
        "right_label": _s("right_label"),
    }


def _same_type_run(blocks: list[dict[str, str]], start: int) -> int:
    """从 start 起,连续同 type 的块有几个。"""
    t = blocks[start].get("type")
    j = start
    while j < len(blocks) and blocks[j].get("type") == t:
        j += 1
    return j - start


def _looks_side_by_side(a: dict[str, str], b: dict[str, str]) -> bool:
    """判断 a、b 两个块是否"同一水平行的左右两栏"(仅用于兜底配对)。"""
    if a.get("type") != b.get("type"):
        return False
    t = a.get("type")
    if t in ("signature", "date"):
        # 签署区(甲方盖章 | 乙方签字)与日期行(左日期 | 右日期)
        # 在中文合同里几乎总是左右并排,配对是安全的。
        return True
    if t == "kv":
        la, lb = a.get("label", ""), b.get("label", "")
        return bool(
            la and lb and la != lb
            and la in _PARTY_LABELS and lb in _PARTY_LABELS
        )
    return False


def _rows_from_blocks(
    blocks: list[Any],
) -> list[tuple[dict[str, str], dict[str, str] | None]]:
    """把线性的 blocks 折叠成"排版行":每行 ``(左块, 右块或 None)``。

    三条规则(按优先级):
        1. 块自带 right_text / right_label → 模型已标出并排,直接一行两栏;
        2. 连续**恰好两个**同类型块 + 语义上确定并排(见 _looks_side_by_side)
           → 兜底配对成一行两栏(模型没按 prompt 合并时仍能还原布局);
        3. 其余一律单栏。
    """
    norm = [_norm_reflow_block(b) for b in blocks if isinstance(b, dict)]
    rows: list[tuple[dict[str, str], dict[str, str] | None]] = []
    i = 0
    while i < len(norm):
        left = norm[i]

        # 规则 1:模型已给出右栏
        if left["right_text"] or left["right_label"]:
            right = {
                "type": left["type"], "text": left["right_text"],
                "label": left["right_label"], "value": left["right_text"],
                "right_text": "", "right_label": "",
            }
            rows.append((left, right))
            i += 1
            continue

        # 规则 2:保守兜底配对
        if (
            left["type"] in _PAIRABLE_TYPES
            and i + 1 < len(norm)
            and _same_type_run(norm, i) == 2
            and not norm[i + 1]["right_text"]
            and not norm[i + 1]["right_label"]
            and _looks_side_by_side(left, norm[i + 1])
        ):
            rows.append((left, norm[i + 1]))
            i += 2
            continue

        rows.append((left, None))
        i += 1
    return rows


def _block_has_content(b: dict[str, str]) -> bool:
    """该块是否有任何可排出的文字(空块会被跳过,避免凭空多出一个空行)。"""
    return bool(b.get("text") or b.get("label") or b.get("value"))


def _block_plain(b: dict[str, str]) -> str:
    """拼出该块的纯文本,供印章 anchor 做子串匹配(多拼字段没有副作用)。"""
    return " ".join(
        p for p in (
            b.get("label"), b.get("value"), b.get("text"),
            b.get("right_label"), b.get("right_text"),
        ) if p
    )


def _reflow_layout_pdf(
    out_path: Path,
    pages: list[dict[str, Any]],
    seals: Any,
) -> None:
    """把结构化版面 JSON **重排**成 Word 风格的 A4 PDF(文字全部正向横排)。

    【为什么这是唯一能满足"像 Word 导出合同"的实现】
        用户要的是"所有文字扶正、正向横排、白底黑字、排版整齐"。
        照片里的文字是**印在像素上的**,它的斜/字号不一/间距不齐都是既成事实;
        而这里排的每一个字都由 reportlab 的 Paragraph 引擎实时计算位置:
            - 字号统一(正文 11pt、标题 18pt);
            - 行距固定 18pt;
            - 首行缩进 2 字符、左对齐/两端对齐由引擎保证;
            - **正向横排**是排版引擎的默认行为,不需要任何图像旋转。
        也就是说,"扶正"不是靠修图实现的,而是靠"不再用图里的文字"实现的。

    排版参数与实现细节(逐条对应用户诉求):
        1. A4 纵向,边距 60pt(见 _REFLOW_MARGIN);
        2. 标题居中加粗 18pt + 下方 18pt 留白;
        3. paragraph:两端对齐 + 首行缩进 2 字符(用 ParagraphStyle 的
           ``firstLineIndent`` = 22pt 实现 —— STSong-Light 下全角空格宽度为 0,
           靠字符本身缩进是无效的,详见 para_style 处的注释);
        4. kv:标签加粗 + 「标签:值」自然流式文本。
           **为什么绝不能用空格/制表符对齐**:reportlab 里空格的宽度
           由字体决定 —— STSong-Light 的半角空格只有 2.28pt,而 U+3000/U+00A0
           在该 CID 字体下**绘制宽度为 0**(见 _keep_blanks 的实测);
           就算能用,同样"5 个空格"在不同字体/字号下宽度也不同,
           用空格凑出来的列必然**参差不齐**。正确做法是"甲方:万华商贸有限公司"
           这样的流式文本,让它随内容自然延伸;
        5. signature/date:左/右对齐,并把原文"日期:      年    月    日"里的
           连续空白折算成等长下划线(见 _keep_blanks);
        6. 文字全部正向横排(排版引擎保证,全程无任何图片旋转);
        7. 段后 8pt(_REFLOW_PARA_GAP);
        8. 所有进入 Paragraph 的文本都先 ``escape()`` —— 原文含 ``& < >``
           会直接抛 XMLSyntaxError;
        9. 页数超长自动分页(交给 SimpleDocTemplate.build);
       10. 每页顶部加一个很淡的"第 N 张照片"页眉(见 _ReflowDocTemplate)。

    印章回贴(**为什么只能"盖在末尾"而不是还原原坐标**):
        版面 JSON **不输出坐标**(理由见 ai_service._OCR_LAYOUT_USER_PROMPT),
        全文又被重排了一遍 —— 原图里"章压在第 3 段第 2 行"这种几何关系,
        在重排后的页面里**已经不存在了**,没有任何坐标可以拿来还原。
        所以只能按**内容语义**推断:把该页印章放在本页正文/签署区之后。
        这是推断而非还原,也是本形态的诚实边界。位置取右对齐
        (与"乙方盖章"这类右侧签署区一致),高度 _REFLOW_SEAL_HEIGHT。

    Args:
        out_path: 输出 PDF 路径。
        pages: ``[{"title": str, "blocks": [...]}, ...]``,每项对应一张原图。
        seals: 与 pages 等长的印章 PIL 列表,无印章的位置为 ``None``。
    """
    _ensure_reflow_fonts()

    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "reflow_body", parent=styles["BodyText"],
        fontName=_REFLOW_FONT, fontSize=_REFLOW_BODY_SIZE,
        leading=_REFLOW_BODY_LEADING, alignment=TA_JUSTIFY,
        # wordWrap="CJK":按**字符**断行而非按单词断行。
        # 中英混排时按单词断行会在行首/行尾留下大片空白,两端对齐又把整页
        # 拉成"疏密不匀"的锯齿;CJK 断行才符合中文文档观感。
        wordWrap="CJK", spaceAfter=_REFLOW_PARA_GAP,
    )
    # 【踩过的坑:首行缩进必须用 firstLineIndent,不能用全角空格】
    #   最初按"Word 缩进 2 字符"的字面意思,在段落文本前拼两个 U+3000。
    #   实测 STSong-Light 下 U+3000 的绘制宽度为 **0**(与 U+00A0 同病),
    #   PDF 里第一行的 x 起点与左边距完全相同 —— 缩进彻底没生效。
    #   reportlab 的 ParagraphStyle **有** firstLineIndent 属性(单位 pt),
    #   这里直接用它:2 个全角字符 = 2 × 字号 = 22pt。
    #   另起一个独立 style 而不是改 body:kv / signature / date 不该被缩进。
    para_style = ParagraphStyle(
        "reflow_para", parent=body,
        firstLineIndent=_REFLOW_BODY_SIZE * 2,
    )
    # 标题必须**显式加粗**:fontName 指向常规字重时,光改字号并不会变粗;
    # 而 ParagraphStyle 没有 fontWeight,reportlab 里加粗只有两条路:
    # 要么 fontName 直接用粗体字体的名字,要么在文本里写 <b>。</。
    # 这里两条都用到:fontName 用粗体名(保证"即使 <b> 失效也依然是粗体",
    # 规避 registerFontFamily 在部分 reportlab 版本上不生效的坑),
    # 再用 <b> 兜一层。
    title_style = ParagraphStyle(
        "reflow_title", parent=body,
        fontName=_REFLOW_FONT_BOLD,
        fontSize=_REFLOW_TITLE_SIZE, leading=_REFLOW_TITLE_SIZE + 8,
        alignment=TA_CENTER, spaceBefore=6, spaceAfter=_REFLOW_TITLE_GAP,
    )
    kv_style = ParagraphStyle("reflow_kv", parent=body, alignment=TA_LEFT)
    sig_style = ParagraphStyle(
        "reflow_sig", parent=body, alignment=TA_LEFT, spaceBefore=16,
    )
    date_style = ParagraphStyle(
        "reflow_date", parent=body, alignment=TA_RIGHT, spaceBefore=10,
    )

    story: list = []
    n = len(pages or [])
    # 可用正文宽度(A4 去掉左右边距),双栏/带章的表格按它分配列宽。
    avail_w = A4[0] - 2 * _REFLOW_MARGIN

    def _make_flow(b: dict[str, str]) -> Any:
        """单个 block → 一条 Paragraph(样式按 type 选)。"""
        btype = b.get("type")
        if btype == "kv":
            label = b.get("label") or ""
            value = b.get("value") or b.get("text") or ""
            # 标签加粗 + 英文冒号,与 Word 合同写法一致(理由见 docstring 第 4 条)
            seg = f"<b>{escape(label)}</b>:{escape(value)}" if label else escape(value)
            return Paragraph(seg, kv_style)
        if btype == "signature":
            return Paragraph(_keep_blanks(escape(b.get("text") or "")), sig_style)
        if btype == "date":
            return Paragraph(_keep_blanks(escape(b.get("text") or "")), date_style)
        if btype != "paragraph":
            # 未知 type 降级成正文,而不是丢弃 —— 丢弃等于让用户少一整段文字
            logger.info("重排:未知 block 类型 %r,按正文排出", btype)
        return Paragraph(escape(b.get("text") or ""), para_style)

    def _two_col(left_flow: Any, right_flow: Any) -> Any:
        """左/右两栏装进无内边距的 2 列表格 → 落在同一水平行上。"""
        t = Table(
            [[left_flow, right_flow]],
            colWidths=[avail_w * 0.45, avail_w * 0.55],
        )
        t.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            # Table 不继承 Paragraph 的 spaceAfter,段后间距只能在单元格底部补,
            # 否则并排的那一行会与下一行贴死。
            ("BOTTOMPADDING", (0, 0), (-1, -1), _REFLOW_PARA_GAP),
        ]))
        return t

    def _row_with_seals(
        left_flow: Any, right_flow: Any, seals: list[tuple[Any, bool]],
    ) -> Any:
        """把印章和本行文字装进同一个表格,让章与这行字处于同一水平线。

        为什么不再用"独立成行 + hAlign":
            那样章只能贴着页面左右边缘,而原件里章是压在某一行的**局部**位置。
            放进表格后,章占一列、文字占其余列,横向关系由列宽决定,
            既不会跑到页面另一头,也不会把整行文字遮住。
        """
        # 同一行可能不止一枚章(左右各一枚):按章数压缩单枚宽度,
        # 保证文字列至少还剩 25% 宽度,否则多枚章会把列宽加起来撑破页面。
        seal_count = max(1, len(seals))
        seal_w = min(
            avail_w * 0.24, _REFLOW_SEAL_HEIGHT * 1.3, avail_w * 0.75 / seal_count,
        )
        left_seals = [s for s, on_left in seals if on_left]
        right_seals = [s for s, on_left in seals if not on_left]

        # 文字单元格:原本就是两栏时再嵌套一层 2 列表格,保持左右并排
        text_cell = None
        if left_flow is not None and right_flow is not None:
            text_cell = _two_col(left_flow, right_flow)
        elif left_flow is not None:
            text_cell = left_flow
        elif right_flow is not None:
            text_cell = right_flow

        cols: list[Any] = list(left_seals)
        widths: list[float] = [seal_w] * len(left_seals)
        if text_cell is not None:
            cols.append(text_cell)
            # 剩余宽度给文字列;至少留 25%,避免章太大把文字挤成一条竖线
            widths.append(max(avail_w * 0.25, avail_w - seal_w * seal_count))
        cols.extend(right_seals)
        widths.extend([seal_w] * len(right_seals))

        if len(cols) == 1:
            # 只有章没有文字 → 独立成行,按原图的左右位置对齐
            flow = cols[0]
            if text_cell is None:
                flow.hAlign = "LEFT" if left_seals else "RIGHT"
            return flow

        t = Table([cols], colWidths=widths)
        t.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), _REFLOW_PARA_GAP),
        ]))
        return t

    for idx in range(n):
        page = pages[idx] if idx < len(pages) else None
        page = page if isinstance(page, dict) else {}
        blocks = page.get("blocks")
        blocks = blocks if isinstance(blocks, list) else []

        # 【2026-10】先按"行"重组:左右并排的两栏合成一行(见 _rows_from_blocks)
        rows = _rows_from_blocks(blocks)

        # 每行先只记录结构,印章分配完再统一渲染 ——
        # 这样"章是否与本行文字同排"可以按最终结构决定。
        entries: list[dict[str, Any]] = []

        raw_title = page.get("title")
        title = raw_title.strip() if isinstance(raw_title, str) else ""
        if title:
            entries.append({"kind": "title", "left": None, "right": None,
                            "text": title, "seals": []})

        for left, right in rows:
            has_any = _block_has_content(left) or (
                right is not None and _block_has_content(right)
            )
            if not has_any:
                continue
            entries.append({
                "kind": "row",
                "left": left,
                "right": right if (right is not None and _block_has_content(right)) else None,
                "text": _block_plain(left) + (
                    "\n" + _block_plain(right) if right is not None else ""
                ),
                "seals": [],
            })

        # ===== 印章分配:按 anchor 找到所属行,并记录原图上的左右位置 =====
        page_seals = _page_seals_for_index(seals, idx)
        row_texts = [e["text"] for e in entries]
        for seal in page_seals:
            img_obj = seal.get("image")
            if img_obj is None:
                continue
            seal_img = _pil_to_rl_image(img_obj, _REFLOW_SEAL_HEIGHT)
            if seal_img is None:
                continue
            # 【横向位置】按章在原图上的**归一化 cx** 判断左右,
            # 而不是只信模型给的 side —— 模型常把"压在左侧文字上的章"也标 right,
            # 旧实现据此一律右对齐,章就被顶到页面最右侧(用户截图里的错位)。
            on_left = False
            cx_raw = seal.get("cx")
            if cx_raw is not None:
                try:
                    on_left = float(cx_raw) < 0.5
                except (TypeError, ValueError):
                    on_left = str(seal.get("side") or "").lower() == "left"
            else:
                on_left = str(seal.get("side") or "").lower() == "left"
            seal_img.hAlign = "LEFT" if on_left else "RIGHT"

            anchor = str(seal.get("anchor") or "").strip()
            pos = _find_anchor_position(row_texts, anchor) if anchor else -1
            if pos >= 0:
                entries[pos]["seals"].append((seal_img, on_left))
            else:
                # 匹配不上 anchor → 退回本页末尾(独立成行),而不是硬塞进某一行
                entries.append({"kind": "row", "left": None, "right": None,
                                "text": "", "seals": [(seal_img, on_left)]})

        # ===== 渲染本页 =====
        has_text = False
        for e in entries:
            if e["kind"] == "title":
                story.append(_tag(
                    Paragraph(escape(_tidy_unclear(e["text"])), title_style), idx,
                ))
                continue

            left, right, seals_of_row = e["left"], e["right"], e["seals"]
            left_flow = _make_flow(left) if left is not None else None
            right_flow = _make_flow(right) if right is not None else None

            if seals_of_row:
                flow = _row_with_seals(left_flow, right_flow, seals_of_row)
            elif left_flow is not None and right_flow is not None:
                flow = _two_col(left_flow, right_flow)
            else:
                flow = left_flow if left_flow is not None else right_flow
            if flow is None:
                continue
            story.append(_tag(flow, idx))
            has_text = True

        if not has_text:
            # 既没有文字也没有印章:给出明确占位,而不是产出一个完全空白的
            # PDF 页(空白页在 Word 里读起来像"丢内容了")。
            story.append(_tag(Paragraph(escape("（未识别到文档内容）"), body), idx))

        # PageBreak 只插在两张图之间:插在最后一张后面会多出一个空白页。
        if idx < n - 1:
            story.append(PageBreak())

    # 极端兜底:pages 为空时也必须产出合法 PDF(否则下载 0 字节)。
    if not story:
        story = [Paragraph(escape("未识别到文字内容"), body)]

    doc = _ReflowDocTemplate(
        str(out_path), pagesize=A4,
        leftMargin=_REFLOW_MARGIN, rightMargin=_REFLOW_MARGIN,
        topMargin=_REFLOW_MARGIN, bottomMargin=_REFLOW_MARGIN,
        pages=pages,
    )
    doc.build(story)


def _tag(flowable: Any, src_index: int) -> Any:
    """给 flowable 打上"来自第几张原图"的标记,供 _ReflowDocTemplate 写页眉。

    为什么用属性标记而不是外部字典:
        platypus 的 afterFlowable 回调**只**把 flowable 本身传出来,
        没有序号上下文;挂在 flowable 上是唯一能把"来源序号"带进排版过程的
        办法,而且不需要给每个 flowable 建旁路字典(Paragraph/RLImage 都不可哈希
        也不便做键,挂在属性上最直接)。
    """
    setattr(flowable, "_reflow_src", int(src_index))
    return flowable


def _split_long_line(line: str, max_chars: int) -> list[str]:
    """把超长单行切成多段(不可见文字层不做自动换行,超宽会溢出页面右侧)。"""
    if len(line) <= max_chars:
        return [line]
    return [line[i:i + max_chars] for i in range(0, len(line), max_chars)]


def _gen_text_pdf(out_path: Path, texts: list[str | None]) -> None:
    """纯文字型 PDF:只排版 OCR 识别文字,不含任何图像对象。

    空行必须保留(段落结构是识别结果的一部分),超长单行交给 Paragraph
    原生换行。写入前统一 escape:OCR 结果里可能含 `&` `<` `>` 等 XML 敏感
    字符(例如"金额 & 利息"、"a < b"),不转义会直接让 reportlab 的
    paragraph 解析器抛 XMLSyntaxError。
    """
    styles = getSampleStyleSheet()
    body = ParagraphStyle("img2pdf_body", parent=styles["BodyText"],
                          fontName="STSong-Light", fontSize=11, leading=18)
    sep = ParagraphStyle("img2pdf_sep", parent=body, fontSize=12, leading=20,
                         spaceBefore=6, spaceAfter=6)

    story: list = []
    has_any_text = False
    for i, raw in enumerate(texts):
        text = (raw or "").strip()
        story.append(Paragraph(escape(f"—— 第 {i + 1} 张图片的识别结果 ——"), sep))
        if not text:
            story.append(Paragraph(escape("（未识别到文字内容）"), body))
            story.append(Spacer(1, 6))
            continue
        has_any_text = True
        for line in text.splitlines():
            line = line.rstrip()
            # 空行用 Spacer 占位,而不是塞空 Paragraph:空 Paragraph 会被
            # reportlab 当成"无内容可画",排版引擎可能直接忽略它,空行就丢了
            if line.strip():
                story.append(Paragraph(escape(line.strip()), body))
            else:
                story.append(Spacer(1, 6))

    # 全部图都没识别出文字时,也必须产出合法 PDF(不能是空文件,否则下载 0 字节)
    if not has_any_text:
        story = [Paragraph(escape("未识别到文字内容"), body)]

    doc = SimpleDocTemplate(
        str(out_path), pagesize=A4,
        leftMargin=_A4_MARGIN, rightMargin=_A4_MARGIN,
        topMargin=_A4_MARGIN, bottomMargin=_A4_MARGIN,
    )
    doc.build(story)


def _count_pdf_pages(pdf_path: Path) -> int:
    """数页数:统一走 pdfplumber,保证返回的是真实页数而不是推算值。"""
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            return len(pdf.pages)
    except Exception:
        logger.warning("统计 PDF 页数失败: %s", pdf_path)
        return 0


def images_to_pdf(
    images: list[tuple[str, bytes]],
    original_name: str,
    *,
    mode: str = "dual",
    ocr_texts: list[str | None] | None = None,
    ocr_errors: list[str | None] | None = None,
    ocr_layouts: list[dict | None] | None = None,
) -> dict[str, Any]:
    """多张照片 → PDF,支持 OCR 文字层。

    Args:
        images: [(原文件名, 图片字节), ...],按上传顺序。
        original_name: 用于拼输出文件名。
        mode: dual(双层,默认) / text(纯文字) / image(图片型)。
        ocr_texts: 视觉模型识别结果,与 images 等长同序;mode="image" 时可为 None。
        ocr_errors: 与 ocr_texts 同序的失败原因,用于回传前端提示。
        ocr_layouts: 【2026-10】结构化版面识别结果(由 ai_service.ocr_layout 产出),
            与 images 等长同序。**仅 mode="dual" 使用**:有它才会走
            "Word 风格重排"(文字正向横排、字号统一、白底黑字);
            没有它则退回"白底净化图 + 不可见文字层"。

    Returns:
        {file_path, file_name, original_size, output_size, page_count,
         mode, ocr_success, ocr_failed, ocr_errors, layout_used}
    """
    if not images:
        raise ValueError("请至少上传一张图片")
    if len(images) > _MAX_IMAGES:
        raise ValueError(f"单次最多支持 {_MAX_IMAGES} 张图片,当前 {len(images)} 张")
    for i, (_name, data) in enumerate(images):
        if data and len(data) > _MAX_IMAGE_BYTES:
            raise ValueError(f"第 {i + 1} 张图片超过 20MB 上限")
    if mode not in ("dual", "text", "image"):
        raise ValueError("PDF 形态只能是 dual / text / image")

    try:
        out_name = _safe_output_name(
            PREFIX_IMG2PDF, original_name or images[0][0], "pdf"
        )
        out_path = settings.upload_dir / out_name
        settings.upload_dir.mkdir(parents=True, exist_ok=True)

        # mode="text" 只排版识别文字,一个像素都用不上。
        # 若照样全量解码,20 张 1600px 图会把内存推到 ~100MB+,纯属浪费;
        # 因此只在 image / dual 两种形态下才做解码(EXIF 校正 + 缩放 + 转 RGB
        # 对这两种形态是必需的,否则图片会躺着、体积也虚高)。
        pil_images: list["Image.Image"] = []
        if mode in ("image", "dual"):
            pil_images = [_prepare_image(data) for _name, data in images]

        # 对齐 OCR 结果长度:入参可能缺项/为 None,统一补成与 images 等长同序
        n = len(images)
        given_texts = list(ocr_texts or [])[:n]
        given_errors = list(ocr_errors or [])[:n]
        texts: list[str | None] = given_texts + [None] * (n - len(given_texts))
        errors: list[str | None] = given_errors + [None] * (n - len(given_errors))

        # 【2026-10 重排】dual 形态优先用结构化版面数据把文字重新排版;
        # 没有版面数据时 _gen_dual_pdf 内部会自动退回"白底图 + 不可见文字层"。
        layouts: list[dict | None] = []
        if mode == "dual" and ocr_layouts:
            layouts = list(ocr_layouts)[:n] + [None] * max(0, n - len(ocr_layouts))

        if mode == "image":
            _gen_image_pdf(out_path, pil_images)
        elif mode == "text":
            _gen_text_pdf(out_path, texts)
        else:
            _gen_dual_pdf(out_path, pil_images, texts, layouts or None)

        if not out_path.exists():
            raise ValueError("转换失败:未生成目标 PDF 文件")

        # ocr_success/ocr_failed 统计的是"识别出非空文字"的图数,
        # 与 errors 里的失败原因一一对应,前端可据此提示"第 N 张识别失败"
        # 注意:重排模式下文字来自版面 blocks,不能只看纯文本 texts
        # (版面模式里 texts 可能为空,但 blocks 有内容),所以两者取并集。
        # 先无条件算出纯文本口径的基线,再按需用版面口径覆盖 ——
        # 踩过的坑:早先写成 if/else,if 分支里读了 ocr_success,
        # 而那个变量只在 else 分支才赋值 → UnboundLocalError。
        ocr_success = sum(1 for t in texts if (t or "").strip())
        if layouts:
            n_blocks = sum(
                1 for lay in layouts
                if isinstance(lay, dict) and (lay.get("blocks") or lay.get("title"))
            )
            ocr_success = max(ocr_success, n_blocks)
        ocr_failed = n - ocr_success

        return {
            "file_path": str(out_path),
            "file_name": out_name,
            "original_size": int(sum(len(data) for _name, data in images)),
            "output_size": int(out_path.stat().st_size),
            "page_count": int(_count_pdf_pages(out_path)),
            "mode": mode,
            "ocr_success": int(ocr_success),
            "ocr_failed": int(ocr_failed),
            "ocr_errors": errors,
            "layout_used": bool(layouts),
        }
    except ValueError:
        raise
    except Exception as exc:
        logger.exception("images_to_pdf 失败: %s", exc)
        raise ValueError(f"图片转 PDF 失败:{exc}")


# ============================================================
# 照片 → 白底可编辑 PDF 的图像净化流水线(2026-10 新增)
# ============================================================
#
# 要解决的问题:
#   现有 mode="dual" 生成的是"原图打底 + 不可见文字层",视觉上仍是一张照片
#   (有灰底、有阴影、有拍摄角度),和用户要的"像 Word 导出的白底黑字文档"是两回事。
#
# 整体思路(四步):
#   1. _detect_seal_regions          —— 找红色印章 bbox(印章是彩色的,白底化会把它抹掉)
#   2. _extract_handwriting_regions  —— 找手写签名/批注 bbox(同上,且手写要保原样)
#   3. _clean_document_background    —— 核心:背景估计除法 + 局部二值化 → 白底黑字图,
#                                      再把保留区**原样**贴回去
#   4. _enhance_for_ocr               —— 给 OCR 用的增强图(要更清楚的字,不是纯黑白)
#   最终由 _gen_clean_pdf 拼成"白底图 + 不可见文字层"。
#
# 颜色空间统一用 HSV 而非 RGB 判断彩色笔迹,原因见各函数注释。


def _np_to_pil(arr: "np.ndarray") -> "Image.Image":
    """numpy uint8 数组(RGB 通道顺序) → PIL Image。

    统一在这里兜一次 dtype:cv2.divide 在 scale 非 255 时会产出 float64,
    而 PIL.Image.fromarray 对 float64 会直接抛 "cannot handle this data type"。
    本项目所有调用点的结果都已是 uint8,这里 clip+astype 只是防御性的一层。
    """
    if arr.dtype != np.uint8:
        arr = np.clip(arr, 0, 255).astype(np.uint8)
    return Image.fromarray(arr)


def _detect_seal_regions(img: "Image.Image") -> list[tuple[int, int, int, int]]:
    """检测图中的**红色印章**区域,返回若干 bbox (x1, y1, x2, y2)。

    为什么用 HSV 而不是 RGB 判红:
        RGB 判红只能写成 `r > 150 and r > g*1.5 and r > b*1.5` 这类**同时**
        依赖三个通道绝对值的规则。而照片的实际光照是不可控的:
            - 暗处拍照 → r 从 255 掉到 120,阈值直接漏检;
            - 暖光/闪光灯 → 纸张偏黄,g 抬到 180 左右,盖章的红 g 也会跟着抬。
        HSV 把"颜色种类(Hue)"和"亮度(V)"拆成了独立通道,判红只看 Hue,
        明暗变化只影响 V —— 所以红章在暗处、在强光下、在阴影里都能被稳定识别。
        同时用 S(饱和度)排除灰白纸张:白纸的 S 接近 0,不会因为"偏红的光"
        就被当成印章。

    为什么红色要拆成**两个 Hue 区间**:
        OpenCV 的 Hue 是 0~179 的**环形量**(把 0~359 的色环对折映射过来),
        纯红对应 0°,而 359° 的"近红"被映射成 179 —— 于是红色被劈成了
        [0,10] 和 [170,180] 两头,只取其中一段会漏掉另一半红。
        另外公章油墨偏橙红(橙红 Hue≈15~30)非常常见,单独再放宽一段
        [10,35] 一起判,否则很多真实公章检不出来。

    形态学处理各自的作用:
        - MORPH_CLOSE 闭运算:把印章文字笔画之间的小空隙连上。公章印出来的
          "XX有限公司"是镂空字,笔画被断开,不闭运算 findContours 会返回
          几十个小碎片而不是一整枚章。
        - dilate 适度膨胀:公章最外圈那道细边框线(常只有 2~3px,且印刷时
          墨色不均)闭合后可能仍与主体脱开,膨胀把它兜进来,否则只保住了
          中间的字、边框丢了。

    轮廓筛选为什么不用"实心度"一刀切:
        真实公章都是**镂空**的(章面刻的字是空心/反白),加上油墨不均、
        拍照噪点,轮廓面积 / 外接矩形面积 的比值可能低到 0.5 以下;
        圆形财务章在边界处被裁到时也可能只有一半轮廓。实心度若当硬门槛,
        会把真实印章误杀。所以这里的策略是:
        **宽高比接近 1(方章/圆章的共同特征)是硬门槛**,
        实心度 + 最小外接圆只作为**对宽高比略微失真的轮廓的补充放行条件**,
        避免"实心度低"这一条把圆形章和镂空章一起过滤掉。

    Args:
        img: RGB 的 PIL Image(通常来自 _prepare_image,已缩放到 ≤1600px)。

    Returns:
        bbox 列表;检不到印章时返回空列表(不抛异常,调用方按"无印章"处理)。
    """
    rgb = np.asarray(img.convert("RGB"))
    h, w = rgb.shape[:2]
    if h == 0 or w == 0:
        return []

    # ---------- 1. 红色掩膜:Hue 分两段,再叠加 S / V 门槛 ----------
    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    h_ch, s_ch, v_ch = hsv[:, :, 0], hsv[:, :, 1], hsv[:, :, 2]
    # 区间 A:色相环两端的红(纯红 H≈0,近红 H≈175~179)
    # 区间 B:橙红(H≈10~35),很多公章/合同章的油墨落在这里
    red = (
        ((h_ch <= 10) | (h_ch >= 170))
        | ((h_ch >= 10) & (h_ch <= 35))
    )
    # S>=80:纸张/印刷黑字在 HSV 里饱和度接近 0,直接排除;
    # V>=50:排除暗红发黑(阴影里的暗红其实也该保留,这里只挡"几乎全黑"的像素,
    #        防止深色印刷物被当成红章)
    mask = (red & (s_ch >= 80) & (v_ch >= 50)).astype(np.uint8) * 255

    # ---------- 2. 形态学:先闭运算连笔画,再膨胀兜住细边框 ----------
    # kernel 尺寸跟图像长边挂钩:印章边框线宽与图像分辨率成正比,
    # 写死 5 在 1000px 图上够用、在 1600px 图上就偏细了。
    k = max(3, int(round(min(h, w) / 200)) | 1)  # |1 保证是奇数,闭运算要求
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (k, k))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    mask = cv2.dilate(mask, kernel, iterations=1)

    # ---------- 3. 找轮廓并按几何形状过滤 ----------
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    img_area = float(h * w)
    # 面积下限:图面积的 0.0008。1600×1200 的图约 1536px²,足以滤掉
    # 印刷红字(单个字约几百 px)和相机噪点,但一个公章至少几千 px。
    min_area = img_area * 0.0008

    boxes: list[tuple[int, int, int, int]] = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < min_area:
            continue
        x, y, cw, ch = cv2.boundingRect(cnt)
        if cw <= 0 or ch <= 0:
            continue
        aspect = cw / float(ch)
        # 几何形状判定(见 docstring):方章宽高比接近 1 且偏方,圆章同样接近 1。
        # 这里放宽到 [0.5, 2.0]:既能容下正方形章,也能容下略扁/略长的椭圆章;
        # 真正离谱的细长条(红色下划线、红色边框线)会被挡掉。
        is_squareish = 0.5 <= aspect <= 2.0

        # 补充放行条件:实心度 + 最小外接圆,专治"宽高比轻微失真"的轮廓。
        # 圆形章的最小外接圆直径天然 ≈ 外接框短边(圆的宽高比恒为 1),
        # 这一条能救回被轻微拉长/压扁的圆章,同时不会误放行细长条
        # (细长条的最小外接圆直径会远大于短边,条件不成立)。
        (_cx, _cy), radius = cv2.minEnclosingCircle(cnt)
        # 实心度:轮廓面积 / 外接框面积。方章高、镂空章低,这里只放宽不设卡
        solidity = area / float(cw * ch)
        # 【2026-10 修·堵细长实心条漏洞】
        # 踩过的坑:实心的黄色便签(45×120,aspect=0.38)会被旧条件放行 ——
        # 它的外接圆直径确实 ≥ 0.9×短边、solidity≈1.0,两条都满足。
        # 但章(圆/方)的宽高比不可能离谱到 0.38,所以补充条件必须
        # 先过一道"接近方正"的窄门槛,才轮得到外接圆/实心度说话。
        is_roundish = (
            0.6 <= aspect <= 1.7
            and (2 * radius) >= 0.9 * min(cw, ch)
            and solidity >= 0.25
        )
        if not (is_squareish or is_roundish):
            continue

        # 向外扩 2~3 像素:把印章外沿的抗锯齿淡边 / 墨迹晕开部分一起保住,
        # 否则裁出来的章会被削掉一圈,看起来"缺边"。
        pad = 3
        boxes.append((
            max(0, x - pad),
            max(0, y - pad),
            min(w, x + cw + pad),
            min(h, y + ch + pad),
        ))

    # ---------- 4. 合并"碎片" bbox:同一枚章只保留一个外框 ----------
    # 【2026-10 修·章碎成多块】
    # 踩过的坑:公章油墨不匀时,红色掩膜里五角星常与圆环**断连**,
    # findContours 就把它们返回成两个独立轮廓 —— 一枚章被抠出
    # "完整主体 + 孤立五角星"两张图。星形碎片随后还会被当作一枚
    # 独立的章参与 anchor 配对,贴到段落中间,PDF 里就出现一个
    # 突兀的大五角星。
    # 判据:小框 80% 以上落在已保留的大框内 → 视为同一枚章的碎片,
    # 直接吸收(五角星 bbox 通常**完全**落在整章 bbox 内)。
    # 两枚不同的章真实叠盖(骑缝章)极少见,且重叠率也很难到 80%,
    # 不会误合并。
    merged: list[tuple[int, int, int, int]] = []
    for box in sorted(boxes, key=lambda b: (b[2] - b[0]) * (b[3] - b[1]),
                      reverse=True):
        x1, y1, x2, y2 = box
        area_box = (x2 - x1) * (y2 - y1)
        absorbed = False
        for (mx1, my1, mx2, my2) in merged:
            ix1, iy1 = max(x1, mx1), max(y1, my1)
            ix2, iy2 = min(x2, mx2), min(y2, my2)
            inter = max(0, ix2 - ix1) * max(0, iy2 - iy1)
            if area_box > 0 and inter >= 0.8 * area_box:
                absorbed = True
                break
        if not absorbed:
            merged.append(box)
    boxes = merged

    return boxes


def _extract_handwriting_regions(img: "Image.Image") -> list[tuple[int, int, int, int]]:
    """检测图中的**手写签名 / 手写批注**区域(黑色或蓝色笔迹),返回 bbox 列表。

    策略:蓝色笔迹用 HSV 单独精确抓(蓝色在灰度里和黑字完全一样,只能靠颜色区分);
    黑色笔迹走灰度前景提取,但**必须极度保守地过滤**。

    为什么黑色手写检测要"宁可漏检、不可误判"(重要):
        白底化的逻辑是"把所有检出为保留的 bbox 原样贴回,其余全部二值化成
        黑白"。如果把一页规整的印刷正文误判成"手写"保留下来,那么整页正文
        都会以**照片原样**留在白底图上 —— 白底化直接失效,用户看到的还是一张
        带灰底的照片,比漏掉一个签名严重得多。
        所以这里的启发式门槛故意设得很高,宁可放过签名也不放过正文。

    具体的区分启发式:
        - 印刷字是**逐个分离的小连通域**,同一行内字高几乎相等(高度集中在
          同一水平带);手写是**连成一片**的笔画团,倾斜、连笔、字高不一。
        - 所以先用全部连通域高度的**中位数**当"常规印刷字高",
          再只保留"明显高于它(≥1.6 倍)"或"宽高比极扁(≤0.35)"的异常形状。
        - 蓝色笔迹不受印刷字干扰(合同正文极少用蓝色印刷体),
          所以蓝色走独立通道,门槛可以正常。

    已知的一处冗余(实测现象,非 bug):
        深红色印章在**灰度**里就是暗色块,所以黑色笔迹检测经常把红章
        自己也认成"笔迹"再返回一遍(实测:红章会同时出现在印章结果和
        手写结果里)。这不影响正确性 —— _clean_document_background 的保留区
        是把两类 bbox **合并**后一起原样贴回,重复保留同一块区域无副作用;
        但它会让调用方看到重复坐标。做数据统计时需自行去重。

    Args:
        img: RGB 的 PIL Image。

    Returns:
        bbox 列表;检不到时返回空列表。
    """
    rgb = np.asarray(img.convert("RGB"))
    h, w = rgb.shape[:2]
    if h == 0 or w == 0:
        return []

    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    h_ch, s_ch, v_ch = hsv[:, :, 0], hsv[:, :, 1], hsv[:, :, 2]
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    img_area = float(h * w)

    boxes: list[tuple[int, int, int, int]] = []

    # ---------- A. 蓝色笔迹:H in [100,130] 是蓝紫区间,S>=60 滤掉浅色印刷 ----------
    # Hue 100~130 覆盖正蓝(约 120)到蓝紫;S>=60 保证是"有颜色"的笔迹而非
    # 灰白纸张;V 门槛放到 30,避免把阴影里的深色块当蓝笔迹。
    blue_mask = (
        (h_ch >= 100) & (h_ch <= 130) & (s_ch >= 60) & (v_ch >= 30)
    ).astype(np.uint8) * 255

    # 形态学:闭运算把手写断开的笔画连上,膨胀把笔迹的浅色边缘裹进来
    k_blue = max(3, int(round(min(h, w) / 150)) | 1)
    k_blue = min(k_blue, 15)
    kb = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k_blue, k_blue))
    blue_mask = cv2.morphologyEx(blue_mask, cv2.MORPH_CLOSE, kb)
    blue_mask = cv2.dilate(blue_mask, kb, iterations=1)

    blue_contours, _ = cv2.findContours(
        blue_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    # 蓝色笔迹面积门槛比红色略宽松(签名笔画细、可能只有几笔),
    # 但仍要挡住蓝色印刷小字/图标
    min_blue_area = img_area * 0.0004
    for cnt in blue_contours:
        area = cv2.contourArea(cnt)
        if area < min_blue_area:
            continue
        x, y, cw, ch = cv2.boundingRect(cnt)
        if cw <= 0 or ch <= 0:
            continue
        pad = 3
        boxes.append((
            max(0, x - pad), max(0, y - pad),
            min(w, x + cw + pad), min(h, y + ch + pad),
        ))

    # ---------- B. 黑色手写:先估"常规印刷字高",再只保留明显异常的连通域 ----------
    # 用 Otsu 全局阈值拿前景。之所以不用自适应阈值作为第一步:
    #   自适应阈值对光照不均更鲁棒,但它会把纸张的渐变阴影也压出"文字状"噪声,
    #   噪声太多反而污染后面的形状统计。这里的目标是"看清有哪些连通域",
    #   前景稍粗一点没关系,后面还会按几何形状严格过滤。
    _, fg = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)

    # 连通域分析:拿到每个"笔画团"的 bbox 与面积,统计常规字高分布
    n_labels, _, stats, _ = cv2.connectedComponentsWithStats(fg, connectivity=8)
    comps: list[tuple[int, int, int, int, int]] = []  # (x, y, w, h, area)
    for i in range(1, n_labels):
        x, y, cw, ch, area = (
            int(stats[i, cv2.CC_STAT_LEFT]), int(stats[i, cv2.CC_STAT_TOP]),
            int(stats[i, cv2.CC_STAT_WIDTH]), int(stats[i, cv2.CC_STAT_HEIGHT]),
            int(stats[i, cv2.CC_STAT_AREA]),
        )
        # 极小的噪点(面积 < 几十 px)先丢掉,免得拉偏字高统计
        if cw == 0 or ch == 0 or area < 20:
            continue
        comps.append((x, y, cw, ch, area))

    if not comps:
        return boxes

    # 用所有连通域高度的中位数当"常规字高"。
    # 用中位数而不是平均数:少数几个异常大连通域(印章、标题、手写)会把平均数
    # 拉高,导致真正的印刷字反而显得"偏矮",误判方向正好反了。
    heights = sorted(c[3] for c in comps)
    median_char_h = heights[len(heights) // 2]
    # 常规字高至少 8px:图太小或文字极小时这个估计不可信,直接放弃黑色手写检测
    if median_char_h < 8:
        return boxes

    # 手写候选的高宽特征(满足其一即保留):
    #   (a) 明显高于常规字高(≥1.6 倍)—— 手写连笔/大字的连通域会明显更高;
    #       规整印刷的每个字高度都趋近中位数,全部落选。
    #   (b) 宽高比 ≤0.35(极扁)—— 手写长横笔/下划式签名;
    #       印刷正文的字宽高比基本都在 0.5~1.5。
    # 注意这里只判"明显异常",不做"正常范围"判断 —— 启发式越宽松,
    # 漏检手写的概率越高但误判正文的概率越低,这是刻意的不对称(见 docstring)。
    for (x, y, cw, ch, area) in comps:
        aspect = cw / float(ch)
        taller_than_text = ch >= median_char_h * 1.6
        too_flat = aspect <= 0.35
        if not (taller_than_text or too_flat):
            continue
        # 面积门槛:单个手写"字"/签名团通常比印刷单字大
        if area < img_area * 0.0006:
            continue
        # 长度门槛:整页表格横线、分隔线虽然"极扁"符合 (b),但宽度会横跨整页,
        # 必须挡掉。要求"宽度不超过图宽的一半",正文中的横线/下划线也一并排除。
        if cw > w * 0.5:
            continue
        pad = 3
        boxes.append((
            max(0, x - pad), max(0, y - pad),
            min(w, x + cw + pad), min(h, y + ch + pad),
        ))

    return boxes


def _clean_document_background(
    img: "Image.Image", keep_boxes: list[tuple[int, int, int, int]],
) -> "Image.Image":
    """把照片"净化"成白底黑字文档图,keep_boxes 内的区域原样保留。

    核心思路(按顺序四步):
        1. **背景估计 + 灰度除法** —— 照片的纸张亮度不均匀(暗角、阴影、
           拍摄角度导致的渐变),直接 Otsu 二值化会把阴影当成文字。
           经典解法:用大核形态学闭运算 / 高斯模糊估计出"只有背景、没有文字"
           的纸张底色图 background,再用 divide(gray, background, scale=255)
           做归一化 —— 数学上就是除掉底色(光照不均被抹平),
           剩下的就是"相对纸张暗多少"的字迹对比度。这是文档扫描增强的标准做法。
        2. **局部自适应阈值** —— 在已抹平光照的图上做 adaptiveThreshold
           (blockSize=31, C=10),得到黑白二值图:文字为 0(黑),纸张为 255(白)。
           用了全局 Otsu 之外还加自适应,是为了应对"同一页里一部分很暗、
           一部分很亮"的复杂光照。
        3. **形态学开运算去噪点** —— 抹掉二值图里残留的孤立黑点(传感器噪点、
           纸张纹理),避免输出图上出现"麻点"。
        4. **保留区贴回原图** —— 印章是彩色、手写是彩色笔迹,若一起二值化
           会全部变成黑色块,印章颜色丢失。所以最后把 keep_boxes 内外的**原始像素**
           直接拷回二值图上,得到"白底 + 黑字 + 原样彩色印章"的成品。

    诚实的能力边界:
        保留是"整块 bbox 原样贴回",不是逐像素语义分割。如果印章恰好压在
        一行正文上,那一小块**正文会作为图像被保留**,而不是变成可编辑文字;
        白底化只保证"绝大部分正文已变成干净黑字",压在被保留区里的少数字
        仍以图片形式存在。这是刻意的取舍 —— 为了保住印章的颜色与形状,
        宁可牺牲压在章下的几个字,也不能让红章被抹成黑色方块。

    Args:
        img: RGB 的 PIL Image。
        keep_boxes: 需要原样保留的 bbox 列表(印章 + 手写)。

    Returns:
        新的 RGB PIL Image(白底、黑字、彩色印章)。
    """
    rgb = np.asarray(img.convert("RGB"))
    h, w = rgb.shape[:2]
    if h == 0 or w == 0:
        return img.convert("RGB")

    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)

    # ---------- 1. 背景估计 + 灰度除法,抹平光照不均 ----------
    # 背景核要"大到能跨过文字笔画、又不会把版面结构当背景"。
    # 取短边的 ~1/12 且为奇数(自适应阈值的 blockSize 同样要求奇数)。
    k_bg = max(15, int(round(min(h, w) / 12)) | 1)
    kernel_bg = cv2.getStructuringElement(cv2.MORPH_RECT, (k_bg, k_bg))
    # MORPH_CLOSE 会把比核小的暗结构(文字)填掉,剩下的近似"纯背景"。
    # 用 divide 而非 subtract:subtract 只是减法,对"整体偏暗"和"整体偏亮"
    # 不对称;divide 是比值归一化,无论纸张偏亮还是偏暗都能压到同一个基准白。
    background = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, kernel_bg)
    # 分母可能为 0(纯黑区域)导致 divide 出 inf/nan;先 clip 下限防除零
    background = np.clip(background, 1, 255).astype(np.uint8)
    normalized = cv2.divide(gray, background, scale=255)

    # ---------- 2. 局部自适应阈值二值化 ----------
    # blockSize 必须是 **>1 的奇数**,否则 OpenCV 断言失败直接抛
    # error(-215:Assertion failed) blockSize % 2 == 1 && blockSize > 1。
    # 小图(短边 < 31)不能硬塞 31,也不能用 min(h,w)|1 —— 短边为 1 时
    # |1 得到 1,不满足 >1。这里统一夹到 [3, 31] 的奇数区间。
    block = max(3, min(31, min(h, w) | 1))
    binary = cv2.adaptiveThreshold(
        normalized, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,   # 文字比局部背景暗 -> 输出里为 0(黑),背景为 255(白)
        block, 12,
    )

    # ---------- 3. 开运算去孤立噪点 ----------
    # 核很小(3x3)只咬掉孤立黑点,不会伤到笔画本身
    kernel_np = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_np)

    # ---------- 4. 保留区贴回原图(彩色印章 / 手写) ----------
    keep_mask = np.zeros((h, w), dtype=np.uint8)
    for (x1, y1, x2, y2) in keep_boxes:
        x1 = max(0, min(int(x1), w))
        y1 = max(0, min(int(y1), h))
        x2 = max(0, min(int(x2), w))
        y2 = max(0, min(int(y2), h))
        if x2 <= x1 or y2 <= y1:
            continue
        keep_mask[y1:y2, x1:x2] = 255

    # 成品底图 = 三通道二值图(黑字白底)
    clean_rgb = cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)

    # 保留区贴回的必须是**去过阴影的彩色图**,而不是原始像素。
    # 踩过的坑:直接贴原图,印章 bbox 里连同周围的阴影一起被搬了过来,
    # 输出上就是一个"灰色方块里嵌着一个红章"——比不做白底化还难看。
    # 修法:对 RGB 三个通道各自除以同一张背景估计图(与前面抹平光照用的是
    # 同一张 background),阴影(约 180 的灰)被除到 255 白纸,
    # 而红色像素 (200,30,40) 除以 180 变成 (255,38,51),依旧是红的,
    # 色相/饱和度关系被保留,红章不会因为提亮而褪色。
    # 注:cv2.divide 要求两侧通道数一致,先把灰度背景扩成三通道。
    bg3 = cv2.cvtColor(background, cv2.COLOR_GRAY2BGR)
    normalized_rgb = cv2.divide(rgb, bg3, scale=255)

    # 再把保留区裁出来贴回 —— 这里用 bitwise_and + add,不能用 bitwise_or:
    #   保留区里若某个像素恰好是纯白(章与纸张之间的空白),or 之后仍然是白,
    #   看起来没问题;但保留区里的**黑色笔画**像素或上白底(255)会直接饱和成白,
    #   章上的字就被"洗掉"了一部分。正确做法是先把保留区清零(清成黑),
    #   再加回原始像素:add(0, 原像素) = 原像素,add(二值, 0) = 二值。
    #   用 add 而非"掩码选择"是为了保持三通道 uint8 的整块赋值,一行搞定。
    #   注意 bitwise_and 必须走 mask= 参数而不是把二维掩膜当第二操作数:
    #   OpenCV 的 binary op 要求两侧 shape 完全一致,3 通道数组和 2 通道掩膜
    #   直接做位运算会抛 error(-209:Sizes of input arguments do not match)。
    cleared = cv2.bitwise_and(clean_rgb, clean_rgb, mask=cv2.bitwise_not(keep_mask))
    out = cv2.add(cleared, normalized_rgb)

    return _np_to_pil(out)


def _enhance_for_ocr(img: "Image.Image") -> "Image.Image":
    """给 OCR 用的增强版图像(和白底化输出是**两个不同的东西**)。

    为什么要和白底化分开:
        - 白底化输出是给人看的成品(纯黑白,视觉上像 Word 文档),
          二值化反而丢掉抗锯齿细节;
        - OCR 输入需要的是"字形边缘更锐利、对比度更高、噪点更少"的**灰阶/彩色**
          图,不需要纯黑白,而且**必须保留红色印章的颜色** ——
          印章文字也要被 OCR 读出来,压成黑白灰度会让红章与黑字混在一起。

    因此这里的策略是"在原图上做局部增强",不做全局灰度化:
        1. 对灰度通道做 CLAHE(限制对比度自适应直方图均衡)—— 提对比、
           压局部过曝/欠曝;
        2. 用 filter2D 做轻量锐化(和原图 addWeighted 混合,避免锐过头出白边);
        3. 增强后的灰度只作用到**亮度方向**:保留原色相/饱和度,
           把红色印章"提亮但仍是红的",这样 OCR 既看得清也能分辨红章。

    Args:
        img: RGB 的 PIL Image。

    Returns:
        新的 RGB PIL Image(对比度增强 + 锐化,印章颜色保留)。
    """
    rgb = np.asarray(img.convert("RGB"))
    h, w = rgb.shape[:2]
    if h == 0 or w == 0:
        return img.convert("RGB")

    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)

    # ---------- 1. CLAHE 提升局部对比度 ----------
    # clipLimit=2.0:既能把阴影里的字提出来,又不会把白纸压到纯灰
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray_eq = clahe.apply(gray)

    # ---------- 2. 轻量锐化(拉普拉斯核),与原图混合避免过冲 ----------
    blur = cv2.GaussianBlur(gray_eq, (3, 3), 0)
    sharp = cv2.addWeighted(gray_eq, 1.5, blur, -0.5, 0)  # unsharp mask 思路

    # ---------- 3. 亮度增强但保留颜色:把增强结果作为 V 的"增益"而非直接替换 ----------
    # 为什么不能直接 V = sharp(踩过的坑):
    #   sharp 来自**灰度**,而灰度是亮度加权平均(0.299R+0.587G+0.114B)。
    #   对高饱和的红像素 (215,32,38),灰度只有 ≈87,而 V 是 215。
    #   若把 V 直接换成 87,红色会被压成暗红(甚至视觉上接近黑色),
    #   OCR 分不出"这是印章" —— 正好把"保留印章颜色"这个需求做反了。
    # 所以这里只取**比值增益** gain = 增强后亮度 / 原亮度:
    #   暗处(阴影)gain > 1 → 提亮;亮处 gain < 1 → 压一点,压缩光照差;
    #   而色相 H 与饱和度 S 一个字节都没动,红色依然是红色。
    denom = np.maximum(gray.astype(np.float32), 1.0)
    gain = sharp.astype(np.float32) / denom
    # gain 夹在 [0.6, 1.8]:防止纯黑区域的除法放大出极端值,
    # 也防止 CLAHE 在高对比边缘处产生过冲把纸面烧白。
    gain = np.clip(gain, 0.6, 1.8)

    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    v_new = np.clip(hsv[:, :, 2].astype(np.float32) * gain, 0, 255).astype(np.uint8)
    hsv[:, :, 2] = v_new  # H/S 保持原样,只改 V
    enhanced = cv2.cvtColor(hsv, cv2.COLOR_HSV2RGB)

    return _np_to_pil(enhanced)


def _gen_clean_pdf(
    out_path: Path,
    cleaned_images: list["Image.Image"],
    ocr_texts: list[str | None],
) -> None:
    """白底可编辑 PDF:每页 = 底层"净化后白底图" + 上层"不可见"文字层。

    形态上是 dual 的"净化版" —— 和 _gen_dual_pdf 的双层结构一致(原图打底 +
    invisible 文字层),区别只在于底层换成 _clean_document_background 处理过的
    白底图,所以视觉上是干净的黑字文档而非照片。

    为什么文字层仍用 setTextRenderMode(3)(不可见):
        页面上已经有清晰的黑色文字图形了。如果再叠一层**可见**文字,同一页
        会同时出现"图片里的字"和"排版引擎画的字",两者位置对不齐 → **重影**,
        反而比单一来源更难看。所以这里保持不可见:只用于让 PDF 里的文字
        **可搜索、可复制**,视觉外观完全由底层白底图决定。

    排版参数沿用 _gen_dual_pdf(10pt / 14pt 行距 / A4 边距),保证两类双层 PDF
    的文字层体验一致。

    诚实的能力边界(同 _gen_dual_pdf):
        OCR 只返回"文字行"的顺序、不返回坐标,故不可见文字层按从上到下固定
        行距排布 —— 可搜索/可复制完全可用,但高亮框是近似位置,不逐字对齐。
        加上净化本身的边界(压在被保留区内的少数字仍是图片),本形态提供的是
        "视觉白底可编辑 + 文字可检索",而非"每个字都是原生可编辑文字对象"。

    Args:
        out_path: 输出 PDF 路径。
        cleaned_images: 已经白底化的 RGB PIL 对象列表,每张一页。
        ocr_texts: OCR 识别结果,与 cleaned_images 等长同序。
    """
    page_w, page_h = A4
    max_w = page_w - 2 * _A4_MARGIN
    max_h = page_h - 2 * _A4_MARGIN
    # 与 _gen_dual_pdf 保持同一套排版参数
    font_size, leading = 10, 14
    max_chars_per_line = 48

    c = canvas.Canvas(str(out_path), pagesize=A4)
    for idx, img in enumerate(cleaned_images):
        # ---------- 底层:净化后的白底图,等比缩放居中 ----------
        w, h = _fit_box(img.width, img.height, max_w, max_h)
        c.drawImage(
            ImageReader(img),
            (page_w - w) / 2, (page_h - h) / 2,
            width=w, height=h,
        )

        # ---------- 上层:不可见文字层(同 dual,可搜索/可复制) ----------
        text = (ocr_texts[idx] if idx < len(ocr_texts) else None) or ""
        if text.strip():
            t = c.beginText()
            t.setTextRenderMode(3)  # invisible:仅检索/复制,不显示不打印
            t.setFont("STSong-Light", font_size)
            t.setLeading(leading)
            y = page_h - _A4_MARGIN - font_size
            t.setTextOrigin(_A4_MARGIN, y)
            for raw_line in text.splitlines():
                line = raw_line.strip()
                if not line:
                    t.textLine(" ")
                    y -= leading
                    continue
                for seg in _split_long_line(line, max_chars_per_line):
                    t.textLine(seg)
                    y -= leading
                    if y < _A4_MARGIN:
                        break
                if y < _A4_MARGIN:
                    break
            c.drawText(t)
        c.showPage()
    c.save()
