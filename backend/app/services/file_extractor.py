"""非结构化文件内容抽取器。

把 Word / PDF / 图片统一转成"纯文本",供下游:
    - 注入 LLM prompt 做数据集对比 / 分析
    - 在数据集详情页展示预览

设计原则:
    1. 抽取是"尽力而为",失败不阻断主流程,只标记 extract_status='failed'
    2. 所有出口都做长度截断(max_extract_chars),防止 prompt 爆炸
    3. PDF 先尝试文本层抽取(pdfplumber),命中失败再走视觉识别(pypdfium2 转图),
       这样电子版 PDF 不消耗视觉 API 配额
    4. 图片识别走多模态 LLM(llm_client.chat_vision),不引入本地 OCR 依赖
"""
from __future__ import annotations

import base64
import logging
from pathlib import Path
from typing import Literal

from app.core.config import settings
from app.services.llm_client import llm

logger = logging.getLogger(__name__)


# 文件类型字面量,与 Dataset.file_type 字段值保持一致
FileType = Literal["structured", "document", "image"]


# 各后缀 → 文件类型的映射,是 upload 路由和 extractor 的唯一事实来源
_SUFFIX_TO_TYPE: dict[str, FileType] = {
    # 结构化(走 Pandas)
    ".xlsx": "structured",
    ".xls": "structured",
    ".csv": "structured",
    # 文档(走文本抽取)
    ".docx": "document",
    ".doc": "document",     # 老版 Word,python-docx 不直接支持,会在抽取时报错降级
    ".pdf": "document",
    ".txt": "document",
    ".md": "document",
    # 图片(走视觉识别)
    ".png": "image",
    ".jpg": "image",
    ".jpeg": "image",
    ".webp": "image",
    ".bmp": "image",
}

# 图片后缀 → MIME 类型(组装 dataURL 时用)
_IMAGE_MIME: dict[str, str] = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".bmp": "image/bmp",
}


def detect_file_type(file_name: str) -> FileType | None:
    """根据文件名后缀判断文件类型。

    返回:
        "structured" / "document" / "image";无法识别时返回 None,由调用方决定
        是拒绝上传还是按 document 兜底。
    """
    suffix = Path(file_name or "").suffix.lower()
    return _SUFFIX_TO_TYPE.get(suffix)


def supported_suffixes() -> list[str]:
    """返回所有支持的上传后缀,用于 upload 接口的白名单校验和前端提示。"""
    return sorted(_SUFFIX_TO_TYPE.keys())


def _truncate(text: str) -> str:
    """统一截断到 max_extract_chars,截断处加省略标记,方便调试时辨认。"""
    if not text:
        return ""
    max_chars = max(1000, int(settings.max_extract_chars or 20000))
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + f"\n\n...[原文共 {len(text)} 字,已截断至 {max_chars} 字]"


# ============== Word 抽取 ==============

def _extract_docx(file_path: Path) -> str:
    """用 python-docx 抽取 Word 段落 + 表格文本。

    老版 .doc(OLE 复合文档)python-docx 不支持,抛异常让上层走通用降级。
    """
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError("缺少 python-docx 依赖,请 pip install python-docx") from exc

    if file_path.suffix.lower() == ".doc":
        raise RuntimeError("暂不支持老版 .doc 格式,请另存为 .docx 后重新上传")

    doc = Document(str(file_path))
    parts: list[str] = []

    # 段落
    for para in doc.paragraphs:
        t = (para.text or "").strip()
        if t:
            parts.append(t)

    # 表格(按 行 → 单元格 展平为 Markdown 风格)
    for idx, table in enumerate(doc.tables, 1):
        parts.append(f"\n[表格 {idx}]")
        for row in table.rows:
            cells = [(c.text or "").strip() for c in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))

    return "\n".join(parts)


# ============== PDF 抽取 ==============

def _extract_pdf_text_layer(file_path: Path) -> str:
    """用 pdfplumber 抽取 PDF 的文本层(仅适用于电子版 PDF)。

    扫描版 PDF 没有文本层,会返回空字符串,由上层决定是否走视觉识别。
    """
    try:
        import pdfplumber
    except ImportError as exc:
        raise RuntimeError("缺少 pdfplumber 依赖,请 pip install pdfplumber") from exc

    parts: list[str] = []
    with pdfplumber.open(str(file_path)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = (page.extract_text() or "").strip()
            if text:
                parts.append(f"--- 第 {i} 页 ---\n{text}")

            # 尝试抽表格(电子版 PDF 常见)
            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []
            for t_idx, table in enumerate(tables, 1):
                parts.append(f"[第 {i} 页 表格 {t_idx}]")
                for row in table:
                    cells = [(c or "").strip() for c in row]
                    if any(cells):
                        parts.append(" | ".join(cells))
    return "\n".join(parts)


def _pdf_to_images_b64(file_path: Path, max_pages: int) -> list[str]:
    """把 PDF 前 N 页转成 PNG 的 base64 列表。

    用 pypdfium2(纯 pip 安装,无需 poppler 系统依赖)渲染。
    每页渲染为 144 DPI(2x scale),在清晰度和 token 成本之间取平衡。
    """
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:
        raise RuntimeError("缺少 pypdfium2 依赖,请 pip install pypdfium2") from exc

    images: list[str] = []
    pdf = pdfium.PdfDocument(str(file_path))
    try:
        n_pages = min(len(pdf), max_pages)
        for i in range(n_pages):
            page = pdf[i]
            # scale=2 即 144 DPI,视觉模型识别精度足够
            pil_image = page.render(scale=2).to_pil()
            import io
            buf = io.BytesIO()
            pil_image.save(buf, format="PNG", optimize=True)
            images.append(base64.b64encode(buf.getvalue()).decode("ascii"))
            try:
                page.close()
            except Exception:
                pass
    finally:
        try:
            pdf.close()
        except Exception:
            pass
    return images


def _extract_pdf(file_path: Path) -> str:
    """PDF 抽取策略:先试文本层,失败走视觉识别。

    判定阈值:全文 < 50 字,视为扫描版(纯文本层抽不出来东西)。
    """
    text = _extract_pdf_text_layer(file_path).strip()

    if len(text) >= 50:
        return text

    # 文本层为空 → 扫描版 PDF,走视觉识别
    logger.info("[extract] PDF 无文本层,走视觉识别: %s", file_path.name)
    max_pages = max(1, int(settings.max_pdf_vision_pages or 20))
    images = _pdf_to_images_b64(file_path, max_pages=max_pages)
    if not images:
        return ""

    # 逐页识别再拼接(比一次性多图更稳,错误粒度更细)
    page_texts: list[str] = []
    for idx, img_b64 in enumerate(images, 1):
        try:
            page_text = llm.chat_vision(
                prompt=(
                    "请仔细识别这一页 PDF 中的所有文字内容,保持原有段落结构。"
                    "如果有表格,用 Markdown 表格语法输出;如果有图表,简要描述其含义。"
                    "直接输出识别结果,不要添加任何解释或评论。"
                ),
                images_b64=[img_b64],
                mime="image/png",
                max_tokens=2048,
                # 扫描件同样是复杂版面,M3.1-Flash 的 adaptive thinking 会把
                # 2048 的额度耗在思考上导致空输出;压低思考链,额度留给正文。
                reasoning_effort="low",
            )
            page_texts.append(f"--- 第 {idx} 页 ---\n{page_text.strip()}")
        except Exception as exc:
            logger.warning("[extract] PDF 第 %d 页视觉识别失败: %s", idx, exc)
            page_texts.append(f"--- 第 {idx} 页 ---\n[识别失败: {exc}]")

    return "\n\n".join(page_texts)


# ============== 图片抽取 ==============

def _extract_image(file_path: Path) -> str:
    """图片 → 视觉模型 → 文本。

    一次性把图喂给视觉模型,让它输出"图中有什么文字/数据/图表"。
    """
    mime = _IMAGE_MIME.get(file_path.suffix.lower(), "image/png")
    try:
        b64 = base64.b64encode(file_path.read_bytes()).decode("ascii")
    except Exception as exc:
        raise RuntimeError(f"图片读取失败: {exc}") from exc

    text = llm.chat_vision(
        prompt=(
            "请详细描述这张图片的内容,包括但不限于:\n"
            "1. 图中所有文字(逐字识别,不要遗漏)\n"
            "2. 如果是数据表格,用 Markdown 表格语法完整输出\n"
            "3. 如果是图表(折线图/柱状图/饼图等),说明图表类型、坐标轴含义、关键数据点\n"
            "4. 如果是截图(网页/软件界面),说明界面主题和关键信息\n"
            "请用清晰的中文输出,保持原有结构,不要添加评论。"
        ),
        images_b64=[b64],
        mime=mime,
        max_tokens=4096,
        # 同扫描版 PDF:压低思考链,避免 max_tokens 被思考耗光。
        reasoning_effort="low",
    )
    return text.strip()


# ============== 统一入口 ==============

def extract_text(file_path: str | Path, file_type: FileType) -> tuple[str, str]:
    """按文件类型抽取文本。

    入参:
        file_path: 文件绝对路径
        file_type: 上游 detect_file_type 判定的类型

    返回:
        (extracted_text, extract_status)
            - 成功: (text, "success")   — text 可能是空字符串(比如纯图 PDF 无文字)
            - 失败: ("", "failed")      — 不抛异常,由调用方决定如何提示用户
            - 跳过: ("", "success")    — structured 类型不需要抽取

    设计说明:
        所有异常都捕获并转为 (\"\", \"failed\"),让上传接口能继续把文件落库,
        用户可以先上传后重试抽取,而不是一次失败就丢失文件。
    """
    path = Path(file_path)
    if not path.exists():
        return "", "failed"

    # 结构化文件不需要抽取文本,直接返回空
    if file_type == "structured":
        return "", "success"

    try:
        if file_type == "document":
            suffix = path.suffix.lower()
            if suffix in (".docx", ".doc"):
                raw = _extract_docx(path)
            elif suffix == ".pdf":
                raw = _extract_pdf(path)
            elif suffix in (".txt", ".md"):
                raw = path.read_text(encoding="utf-8", errors="ignore")
            else:
                return "", "failed"
        elif file_type == "image":
            raw = _extract_image(path)
        else:
            return "", "failed"

        return _truncate(raw), "success"

    except Exception as exc:
        logger.exception("[extract] 文件内容抽取失败: %s (%s)", path.name, file_type)
        # 把错误信息存进 extracted_text,方便用户在前端预览时看到具体原因
        # 截断到 500 字防止超长堆栈污染数据库
        err_msg = f"[内容抽取失败: {type(exc).__name__}] {exc}"[:500]
        return err_msg, "failed"
