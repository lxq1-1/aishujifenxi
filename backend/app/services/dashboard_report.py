"""可视化报表内容生成 + Excel 渲染。

背景与目标：
    图表编辑弹窗目前只能产出"单张图"，用户无法一次性下载一份可以
    直接用于汇报/归档的完整报表。本模块补上"报表生成层"：

        1. ``build_visual_report`` —— 把指标、图表分组、可选明细表、
           可选 AI 结论组装成统一的 ``report_model``（纯数据，不含 HTTP）；
        2. ``render_visual_report_excel`` —— 把 ``report_model`` 渲染为
           ``.xlsx`` 字节流，包含 4 个板块：
             KPI 指标卡 / AI 文字结论 / 数据明细表 / 图表洞察注释；
        3. ``summarize_visual_report`` —— LLM 结论 + 数值溯源复核 +
           规则降级，保证"永远有结论且数字永远真实"；
        4. ``_adapt_chart_groups`` —— 兼容前端直传 / by_section / 孤儿图
           三种入参形态，统一成 ``chart_group`` 列表。

设计原则（为什么这么写）：
    - **异常安全**：AI 不可用、单图渲染失败、指标为空都不能让整份
      报表生成失败，调用方永远能拿到可下载的文件；
    - **不重复造轮子**：风格预设、兜底渲染、Sheet 名清洗、图表 option
      摘要、指标数值池、通知数字审计全部复用既有服务；
    - **零新增依赖**：openpyxl / Pillow 已在 requirements 中；
    - **数字可溯源**：AI 结论里的每个数字都必须在指标 JSON 找到原值，
      否则整条作废并降级为规则结论。

【重要】本模块不做 HTTP 接口，也不读写数据库，只负责内容生成。
"""
from __future__ import annotations

import base64
import io
import json
import logging
import math
import re
from typing import Any

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.core.config import now_local
from app.core.json_utils import json_default, strip_thinking
from app.services.ai_service import _normalize_summary_text
from app.services.analyzer import _collect_indicator_values, audit_section_numbers
from app.services.chart_renderer import (
    _clean_label,
    get_style_preset,
    render_chart_to_png,
)
from app.services.exporter import _chart_option_json, _sanitize_excel_sheet_title
from app.services.llm_client import llm
from app.services.pandas_engine import DETAIL_TABLE_MAX_ROWS

logger = logging.getLogger(__name__)


# ============================================================
# 1) 常量集中定义（A6.6：不散落魔法值）
# ============================================================

_MAX_KPI_CARDS = 4
_MAX_DETAIL_COLS = 16
_MAX_SUMMARY_PAYLOAD_CHARS = 4200
# 输出 token 上限：必须 >= 「思考 token + 答复 token」。
# 踩过的坑：最初设 700，模型把配额全用在思考过程上，strip_thinking 之后
# 拿到空串 → 每次都被动降级成规则结论，看起来像"AI 功能没生效"。
# 与既有 ai_service 的口径对齐（那里同样注释了"给 2000 留足余量"）。
_MAX_SUMMARY_OUTPUT_TOKENS = 2000
_SUMMARY_TEMPERATURE = 0.4

# —— 颜色（openpyxl 使用 ARGB；注释里保留设计稿的 #RRGGBB）——
_CARD_BG = "FFF8FAFC"        # #f8fafc 浅灰卡面
_BORDER_COLOR = "FFE2E8F0"   # #e2e8f0 细边框
_HEAD_BG = "FFF1F5F9"        # #f1f5f9 表头/段标题浅底
_INFO_BG = "FFEFF6FF"        # #eff6ff AI 结论浅蓝底
_INFO_BORDER = "FFBFDBFE"    # #bfdbfe AI 结论边框
_TEXT_PRIMARY = "FF1F2937"   # #1f2937 主文字
_TEXT_MUTED = "FF6B7280"     # #6b7280 次要文字
_TEXT_SUB = "FF64748B"       # #64748b 脚注文字
_WHITE = "FFFFFFFF"
_ACCENT_FALLBACK = "#2dd4bf"

# —— 字号 ——
_FONT_TITLE = 16
_FONT_SUBTITLE = 10
_FONT_SECTION = 11
_FONT_CARD_LABEL = 10
_FONT_CARD_VALUE = 18
_FONT_CARD_SUB = 9
_FONT_BODY = 10
_FONT_HEADER = 10
_FONT_NOTE = 9
_FONT_JSON = 8

# —— 尺寸 / 布局 ——
_CARD_COLS = 3                 # 每张 KPI 卡占 3 列
_CARD_ROWS = 3                 # 每张 KPI 卡占 3 行
_OVERVIEW_COLS = 12            # 概览页总列数（4 卡 × 3 列）
_CHART_IMAGE_WIDTH = 720
_CHART_IMAGE_HEIGHT = 480
_CHART_SEGMENT_ROWS = 14       # 一张图占 14 行
_CHART_ROW_HEIGHT = 26         # 14 行 ≈ 485px，与 480px 图片高度接近
_CHART_COL_WIDTH = 10
_INSIGHT_ROW_HEIGHT_MIN = 26
_INFO_CHARS_PER_LINE = 28      # 按每行约 28 字估算行高

# —— 看板网格（图表直接排在「报表概览」里，不再每组开一个 Sheet）——
# 设计参照用户给的《经营看板》版式：图表以「两列并排」的卡片形式排布，
# 每张卡片 = 标题条 + 图片 + 洞察注释，整体读起来是一张连续的下看。
#
# 为什么用 12 列 × 2 卡：一行放两张图，卡宽 6 列，与 KPI 卡区（4 卡 × 3 列 = 12 列）
# 共用同一套列宽，视觉对齐，不会出现"下面窄一截"的错位。
#
# 【尺寸换算】Excel 列宽单位 ≈ 字符宽，1 字符 ≈ 7px；列宽 12 ≈ 89px。
#   卡面宽 = 6 列 × 89px ≈ 534px  →  图片宽设 500px，两侧各留约 17px，
#   刚好露出卡片细边框，不会压到右边那张卡。
#   行高单位是磅(1pt = 1.33px)，行高 26pt ≈ 35px；图片区 10 行 ≈ 347px，
#   图片高 320px，底部留一点空白。
_DASH_COLS = 12                # 看板总列数
_DASH_CARD_COLS = 6            # 每张图表卡占 6 列（两卡并排）
_DASH_TITLE_ROWS = 1           # 卡内标题条行数
_DASH_INSIGHT_ROWS = 2         # 卡内洞察注释行数
_DASH_IMAGE_ROWS = 10          # 卡内图片区行数（10 × 35px ≈ 347px）
_DASH_ROW_HEIGHT = 26          # 看板统一行高（磅）
_DASH_IMAGE_WIDTH = 500        # 卡片图片宽（两卡并排，单卡不能超过半屏）
_DASH_IMAGE_HEIGHT = 320       # 卡片图片高（保持 500:320 ≈ 1.56:1）
_DASH_GAP_ROWS = 2             # 两行卡片之间的留白行数
_DASH_CHART_TITLE_FONT = 11
# 一张卡片占的总行数：标题 1 + 图片 10 + 洞察 2
_DASH_CARD_ROWS = _DASH_TITLE_ROWS + _DASH_IMAGE_ROWS + _DASH_INSIGHT_ROWS

# 比率类字段名关键词：用于判断"0.856 应该展示成 85.6%"。
_RATIO_HINTS = (
    "率", "占比", "比例", "比率", "达成", "同比", "环比",
    "增长", "增幅", "毛利率", "净利率", "转化", "留存",
    "percent", "ratio", "rate",
)

# 数字解析正则：兼容千分位、百分号、单位后缀空格。
_NUMBER_RE = re.compile(r"^[+-]?\d+(?:,\d{3})*(?:\.\d+)?$")


# ============================================================
# 2) 通用小工具：数值、格式化、Sheet 名、时间
# ============================================================

def _safe_float(value: Any) -> float | None:
    """把任意值尽力转成有限浮点数，失败返回 ``None``。

    为什么需要：
        指标 / 图表数据来自 pandas、numpy、LLM JSON，可能是
        ``np.float64``、``"1,234.5"``、``"85.6%"``、``None``、``NaN``。
        直接 ``float()`` 会抛异常或得到 NaN/Inf，后续格式化/排序都会炸。
    """
    if value is None or isinstance(value, bool):
        return None
    if hasattr(value, "item"):
        try:
            value = value.item()
        except (ValueError, AttributeError):
            pass
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace(" ", "")
        text = text.removesuffix("%")
        if not text:
            return None
        # 只接受"纯数字"，避免把 "销售额 100" 这类文本误当数值。
        if not _NUMBER_RE.match(text):
            return None
        value = text
    if not isinstance(value, (int, float, str)):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return number


def _is_ratio_name(name: Any) -> bool:
    """字段名是否属于"比率类"。

    用途：决定 0.856 是按 ``0.856`` 展示还是按 ``85.6%`` 展示。
    只按字段名判断，不按值域硬猜，避免把"销量 1.2"错当成 120%。
    """
    text = str(name or "").lower()
    return any(hint in text for hint in _RATIO_HINTS)


def _format_number(
    value: Any,
    *,
    percent: bool = False,
    unit: str = "",
) -> str:
    """数值格式化：千分位 + 小数位自适应。

    Args:
        value: 原始值（接受 str / numpy 标量）。
        percent: ``True`` 时把 0.856 转成 ``85.6%``（不做千分位）。
        unit: 非空且不是 ``%`` 时追加到结果尾部，如 ``" 万元"``。

    为什么这样取舍：
        - 整数不显示 ".00"，避免 4680 变成 "4,680.00" 的冗余；
        - 小于 1 的数保留最多 4 位小数（比率、单价常见）；
        - 其它保留 2 位并去掉尾随 0；
        - 始终加千分位，长金额在 Excel 里更易读。
    """
    number = _safe_float(value)
    if number is None:
        return "—"
    if percent:
        return f"{number * 100:.1f}%"
    if abs(number - round(number)) < 1e-9:
        text = f"{number:,.0f}"
    elif abs(number) < 1:
        text = f"{number:,.4f}".rstrip("0").rstrip(".")
    else:
        text = f"{number:,.2f}".rstrip("0").rstrip(".")
    unit_text = str(unit or "").strip()
    if unit_text and unit_text != "%":
        text = f"{text} {unit_text}"
    return text


def _format_metric(value: Any, name: Any, unit: str = "") -> str:
    """按指标名/单位自动选择"普通数值"或"百分数"展示。

    - 字段名含"率/占比/同比…"或单位是 ``%``：
        ``|v| <= 1.5`` 视为小数比率（0.856 → 85.6%）；
        否则视为已经是百分数值（85.6 → 85.6%）。
    - 其它：千分位 + 自适应小数位，并按需追加单位。
    """
    number = _safe_float(value)
    if number is None:
        return "—"
    ratio_like = _is_ratio_name(name) or str(unit or "").strip() == "%"
    if ratio_like:
        if abs(number) <= 1.5:
            return f"{number * 100:.1f}%"
        return f"{number:,.1f}%"
    return _format_number(number, unit=unit)


def _format_percent_value(value: Any) -> str:
    """格式化"已经是百分数"的值（如 93.6 → ``93.6%``）。"""
    number = _safe_float(value)
    if number is None:
        return "—"
    return f"{number:.1f}%"


def _now_label() -> str:
    """返回本地墙钟时间 ``YYYY-MM-DD HH:MM``。

    复用项目统一的 ``now_local``（Asia/Shanghai 口径），
    避免这里再写一份 ``datetime.now()`` 导致部署到 UTC 机器时
    副标题时间与任务列表时间相差 8 小时。
    """
    return now_local().strftime("%Y-%m-%d %H:%M")


def _accent_argb(color: Any) -> str:
    """把 ``#RRGGBB`` 转成 openpyxl 需要的 ARGB 形式。"""
    text = str(color or "").strip()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", text):
        return "FF" + text[1:].upper()
    return "FF" + _ACCENT_FALLBACK[1:].upper()


def _unique_sheet_title(workbook: Workbook, base: str) -> str:
    """生成工作簿内唯一的 Sheet 名。

    为什么需要：
        ``_sanitize_excel_sheet_title`` 只处理非法字符与 31 字符上限，
        但两个章节同名（或都被截断成同一前缀）时仍会冲突，
        openpyxl 会静默改名为 ``xxx1``，破坏"图 - 分组名"的对应关系。
    """
    title = _sanitize_excel_sheet_title(base)
    existing = {ws.title.lower() for ws in workbook.worksheets}
    if title.lower() not in existing:
        return title
    for index in range(2, 100):
        suffix = f"-{index}"
        candidate = f"{title[: max(1, 31 - len(suffix))]}{suffix}"
        if candidate.lower() not in existing:
            return candidate
    return title[:28] + "-x"


def _merge_and_fill(
    worksheet: Any,
    row: int,
    col_start: int,
    col_end: int,
    *,
    color: str | None = None,
) -> None:
    """先给区域内每个单元格上底色再合并。

    背景：openpyxl 合并单元格后，只有左上角单元格保留样式；
    若只给左上角上色，部分 Excel 渲染下会出现白边。先逐格填充
    再合并，视觉上最稳。
    """
    if col_end < col_start:
        return
    if color:
        fill = PatternFill(fill_type="solid", fgColor=color)
        for col in range(col_start, col_end + 1):
            worksheet.cell(row=row, column=col).fill = fill
    worksheet.merge_cells(
        start_row=row,
        start_column=col_start,
        end_row=row,
        end_column=col_end,
    )


def _estimate_wrapped_height(text: str, chars_per_line: int = _INFO_CHARS_PER_LINE) -> float:
    """按字数估算合并单元格的行高（每行约 28 个中文字符）。

    Excel 不会自动撑高"自动换行"的合并单元格，必须手工设置 row height，
    否则 AI 结论会被截断成一行。这里用保守的 15pt/行 + 10pt 内边距。
    """
    content = str(text or "")
    lines = max(1, math.ceil(len(content) / max(1, chars_per_line)))
    return float(min(409, max(30, lines * 15 + 10)))


def _excel_cell_value(value: Any) -> Any:
    """把报表单元格值转成 openpyxl 可写入的形态。

    - ``None`` / NaN → 空串（Excel 空单元格）；
    - numpy 标量 → Python 原生数；
    - 日期时间 → ISO 字符串；
    - list / dict → JSON 字符串（避免 openpyxl 写对象报错）；
    - 其余无法识别的对象 → ``str``，保证写入永不抛类型异常。
    """
    if value is None:
        return ""
    if hasattr(value, "item"):
        try:
            value = value.item()
        except (ValueError, AttributeError):
            pass
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return ""
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except (TypeError, ValueError):
            pass
    if isinstance(value, (list, dict)):
        try:
            return json.dumps(value, ensure_ascii=False, default=json_default)
        except (TypeError, ValueError):
            return str(value)
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _decode_png_base64(raw: Any) -> bytes | None:
    """解码前端直传的 ``png_base64``。

    兼容 ``data:image/png;base64,....`` 前缀与无前缀两种形态。
    任何异常都返回 ``None``，由调用方走兜底渲染，绝不让整份报表失败。
    """
    if not isinstance(raw, str) or not raw.strip():
        return None
    text = raw.strip()
    if text.startswith("data:") and "," in text:
        text = text.split(",", 1)[1]
    text = re.sub(r"\s+", "", text)
    try:
        data = base64.b64decode(text, validate=False)
    except (ValueError, TypeError, base64.binascii.Error):
        logger.warning("[dashboard_report] png_base64 解码失败")
        return None
    return data if len(data) > 8 else None


# ============================================================
# 3) 图表 option 反解：明细表 + 规则洞察
# ============================================================

def _unwrap_option(chart_or_option: Any) -> dict[str, Any]:
    """取出 ECharts option 本体。

    入参既可能是 ``{"option": {...}}``（前端图表项），也可能是
    ``{"series": [...]}``（裸 option）。统一在这里适配，避免下游重复判断。
    """
    if not isinstance(chart_or_option, dict):
        return {}
    inner = chart_or_option.get("option")
    if isinstance(inner, dict):
        return inner
    return chart_or_option


def _axis_obj(axis: Any) -> dict[str, Any]:
    """取坐标轴配置对象；ECharts 允许 xAxis/yAxis 是数组，取第一个。"""
    if isinstance(axis, list):
        for item in axis:
            if isinstance(item, dict):
                return item
        return {}
    return axis if isinstance(axis, dict) else {}


def _category_axis(option: dict[str, Any]) -> tuple[list[str], str]:
    """反解分类轴：返回 ``(分类名列表, 轴标题)``。

    关键兼容点：
        - 普通柱/折线图：分类在 ``xAxis.data``；
        - **横向条形图**：ECharts 约定 ``xAxis.type == "value"``，
          分类数据实际在 ``yAxis.data`` —— 这是最容易反解错的一类，
          若不处理，"横向条形图"的明细表会只剩一串数字。
    """
    x_axis = _axis_obj(option.get("xAxis"))
    y_axis = _axis_obj(option.get("yAxis"))
    x_type = str(x_axis.get("type") or "").lower()

    if x_type == "value":
        raw = y_axis.get("data")
        # 【只取分类轴的 name】横向条形图分类在 yAxis。
        # 绝不能回落成 xAxis.name —— 那是**数值轴**的名字（往往就是单位，如「件」），
        # 拿它当分类列标题会让表头出现「件 | 销售数量」，用户会以为单位是分类。
        axis_name = y_axis.get("name") or ""
    else:
        raw = x_axis.get("data")
        axis_name = x_axis.get("name") or ""

    # 兼容"未声明 type 但分类写在 yAxis"的脏数据。
    if not isinstance(raw, list) or not raw:
        alt = y_axis.get("data")
        if isinstance(alt, list) and alt:
            raw = alt
            # 同理：只有在确实改从 yAxis 取分类时，才改用 yAxis.name
            axis_name = y_axis.get("name") or axis_name

    categories: list[str] = []
    if isinstance(raw, list):
        for index, item in enumerate(raw):
            label = _clean_label(item, 64)
            categories.append(label or f"项{index + 1}")
    return categories, _clean_label(axis_name, 64)


def _series_values(series: dict[str, Any]) -> tuple[list[str], list[float | None]]:
    """把 ``series[i].data`` 归一成 ``(名称列表, 数值列表)``。

    兼容三种形态：
        - 裸数字：``[10, 20, 30]``
        - 对象：``[{"name": "A", "value": 10}, ...]``
        - 散点坐标：``[[x, y], ...]`` / ``{"value": [x, y]}``
    """
    data = series.get("data") if isinstance(series, dict) else None
    if not isinstance(data, list):
        return [], []
    names: list[str] = []
    values: list[float | None] = []
    for index, item in enumerate(data):
        if isinstance(item, dict):
            raw_name = item.get("name") or item.get("label") or ""
            raw_value = item.get("value")
            if isinstance(raw_value, list):
                raw_value = raw_value[-1] if raw_value else None
        elif isinstance(item, (list, tuple)):
            raw_name = item[0] if item else ""
            raw_value = item[-1] if item else None
        else:
            raw_name = ""
            raw_value = item
        name = _clean_label(raw_name, 64) or f"系列{index + 1}"
        names.append(name)
        values.append(_safe_float(raw_value))
    return names, values


def _number_cell(value: float | None) -> Any:
    """明细表数值单元格：整数值落成 int，避免 Excel 显示 10.0。"""
    if value is None:
        return None
    if abs(value - round(value)) < 1e-9:
        return round(value)
    return value


def _extract_chart_detail(option: dict[str, Any]) -> dict[str, Any]:
    """从单张 ECharts option 反解出明细表。

    Returns:
        ``{"columns": [str], "rows": [[cell]], "total_rows": int, "note": str}``

    为什么不做"尽力猜"：
        图表可能经过聚合/堆叠/百分比化，反解出的明细只是"图上可见数据"，
        因此这里只做确定性抽取；抽不到就返回空表 + 明确说明，
        避免给用户一张"看起来像明细、其实张冠李戴"的表。
    """
    opt = _unwrap_option(option)
    series_list = opt.get("series")
    if not isinstance(series_list, list) or not series_list:
        return {
            "columns": [],
            "rows": [],
            "total_rows": 0,
            "note": "该图无可反解的明细数据",
        }

    first = series_list[0] if isinstance(series_list[0], dict) else {}
    chart_type = str(first.get("type") or "bar").lower()

    # —— 饼图：名称/数值两列（名称来自 data[].name，不来自坐标轴）——
    if chart_type == "pie":
        names, values = _series_values(first)
        total_rows = len(names)
        rows: list[list[Any]] = []
        for index in range(min(total_rows, DETAIL_TABLE_MAX_ROWS)):
            rows.append([names[index] or f"系列{index + 1}", _number_cell(values[index])])
        note = ""
        if total_rows > len(rows):
            note = f"共 {total_rows} 行，仅展示前 {len(rows)} 行"
        return {
            "columns": ["名称", "数值"],
            "rows": rows,
            "total_rows": total_rows,
            "note": note,
        }

    # —— 笛卡尔坐标系：分类 + 每个系列一列数值 ——
    categories, axis_name = _category_axis(opt)
    series_columns: list[tuple[str, list[float | None]]] = []
    for index, series in enumerate(series_list):
        if not isinstance(series, dict):
            continue
        names, values = _series_values(series)
        column_name = _clean_label(series.get("name"), 64)
        if not column_name:
            column_name = "数值" if len(series_list) == 1 else f"数值{index + 1}"
        # 分类轴缺失时，退化为用 series data 里的 name 当分类（饼图式 bar 也会走到这里）。
        if not categories and names:
            categories = names
        series_columns.append((column_name, values))

    if not series_columns:
        return {
            "columns": [],
            "rows": [],
            "total_rows": 0,
            "note": "该图无可反解的明细数据",
        }

    lengths = [len(categories)] + [len(values) for _, values in series_columns]
    lengths = [length for length in lengths if length > 0]
    if not lengths:
        return {
            "columns": [],
            "rows": [],
            "total_rows": 0,
            "note": "该图无可反解的明细数据",
        }

    # 需求约定"长度不一致按短对齐"：只保留所有维度都有值的行，
    # 缺失单元格（例如某系列提前结束）补 None。
    aligned = min(lengths)
    total_rows = max(lengths)
    if not categories:
        categories = [f"系列{index + 1}" for index in range(aligned)]

    columns = [axis_name or "分类"] + [name for name, _ in series_columns]
    rows = []
    for row_index in range(min(aligned, DETAIL_TABLE_MAX_ROWS)):
        row: list[Any] = [categories[row_index] if row_index < len(categories) else None]
        for _, values in series_columns:
            row.append(
                _number_cell(values[row_index]) if row_index < len(values) else None
            )
        rows.append(row)

    note = "该图无可反解的明细数据" if not rows else ""
    if rows and total_rows > len(rows):
        note = f"共 {total_rows} 行，仅展示前 {len(rows)} 行"
    return {
        "columns": columns,
        "rows": rows,
        "total_rows": total_rows,
        "note": note,
    }


def _series_entries(
    option: dict[str, Any],
    series_index: int = 0,
) -> list[tuple[str, float]]:
    """把某个 series 与分类轴对齐成 ``[(名称, 数值), ...]``（过滤 None）。"""
    opt = _unwrap_option(option)
    series_list = opt.get("series")
    if not isinstance(series_list, list) or len(series_list) <= series_index:
        return []
    series = series_list[series_index]
    if not isinstance(series, dict):
        return []
    categories, _ = _category_axis(opt)
    names, values = _series_values(series)
    entries: list[tuple[str, float]] = []
    for index, value in enumerate(values):
        if value is None:
            continue
        if index < len(categories):
            name = categories[index]
        elif index < len(names):
            name = names[index]
        else:
            name = f"系列{index + 1}"
        entries.append((_clean_label(name, 64) or f"系列{index + 1}", value))
    return entries


def _chart_insight(option: dict[str, Any], title: str) -> str:
    """规则反解图表洞察（绝不调用 LLM）。

    为什么坚持"规则"：
        洞察要直接写进 Excel 单元格，必须与图上数字严格一致。
        LLM 即使被要求"只引用原值"，仍有概率换算单位/编造字段；
        规则反解虽然文风朴素，但**每个数字都可从 series 直接复算**。

    输出为可直接进单元格的中文短句；无法反解时返回说明句，不抛异常。
    """
    opt = _unwrap_option(option)
    series_list = opt.get("series")
    if not isinstance(series_list, list) or not series_list:
        return "该图暂无可解读的图表数据。"
    first = series_list[0] if isinstance(series_list[0], dict) else {}
    chart_type = str(first.get("type") or "bar").lower()
    title_text = _clean_label(title, 64)
    prefix = f"「{title_text}」：" if title_text else ""

    # —— 饼图：占比最高项 + 前三合计占比 ——
    if chart_type == "pie":
        names, values = _series_values(first)
        entries = [
            (names[index] or f"系列{index + 1}", value)
            for index, value in enumerate(values)
            if value is not None
        ]
        if not entries:
            return prefix + "该图暂无可解读的饼图数据。"
        ranked = sorted(entries, key=lambda item: item[1], reverse=True)
        total = math.fsum(value for _, value in entries)
        top_name, top_value = ranked[0]
        if abs(total) < 1e-12:
            return (
                prefix
                + f"占比最高的是「{top_name}」（数值 {_format_number(top_value)}），"
                + "其余项占比因合计为 0 无法计算。"
            )
        top_share = top_value / total * 100
        top3_share = math.fsum(value for _, value in ranked[:3]) / total * 100
        return (
            prefix
            + f"占比最高的是「{top_name}」（{top_share:.1f}%），"
            + f"前三项合计占比 {top3_share:.1f}%。"
        )

    # —— 散点图：点数 / 最大 / 最小 / 均值 ——
    if chart_type == "scatter":
        entries = _series_entries(opt)
        if not entries:
            return prefix + "该图暂无可解读的散点数据。"
        values = [value for _, value in entries]
        average = math.fsum(values) / len(values)
        return (
            prefix
            + f"共 {len(entries)} 个数据点，最大值 {_format_number(max(values))}，"
            + f"最小值 {_format_number(min(values))}，"
            + f"均值 {_format_number(average)}。"
        )

    # —— 柱状 / 折线：最高、最低、合计（折线额外首末变化 + 趋势 + 峰值）——
    entries = _series_entries(opt)
    if not entries:
        return prefix + "该图暂无可解读的图表数据。"
    highest = max(entries, key=lambda item: item[1])
    lowest = min(entries, key=lambda item: item[1])
    total = math.fsum(value for _, value in entries)
    base = (
        f"最高的是「{highest[0]}」（{_format_number(highest[1])}），"
        f"最低的是「{lowest[0]}」（{_format_number(lowest[1])}），"
        f"合计 {_format_number(total)}"
    )

    if chart_type == "line" and len(entries) >= 2:
        first_value = entries[0][1]
        last_value = entries[-1][1]
        peak_name, peak_value = highest
        if abs(first_value) < 1e-12:
            trend_text = "首末变化率因首期数值为 0 无法计算"
        else:
            change_rate = (last_value - first_value) / abs(first_value) * 100
            if change_rate > 3:
                direction = "整体上升"
            elif change_rate < -3:
                direction = "整体下降"
            else:
                direction = "基本持平"
            trend_text = f"首末变化率 {change_rate:.1f}%，{direction}"
        return (
            prefix
            + base
            + f"；{trend_text}；峰值出现在「{peak_name}」（{_format_number(peak_value)}）。"
        )

    return prefix + base + "。"


# ============================================================
# 4) 图表分组适配（前端直传 / by_section / 孤儿图）
# ============================================================

def _normalize_chart_item(raw: Any, index_hint: int) -> dict[str, Any] | None:
    """把任意形态的图项统一成报表内部结构。

    入参可能已经是 ``{"title","option","sheet","png_base64"}``，
    也可能是裸 ECharts option（顶层带 ``series``）。
    """
    if not isinstance(raw, dict):
        return None
    raw_title = raw.get("title")
    if isinstance(raw_title, dict):
        title_text = raw_title.get("text")
    else:
        title_text = raw_title

    option = raw.get("option") if isinstance(raw.get("option"), dict) else None
    if option is None:
        option = raw if raw.get("series") else {}

    if not title_text:
        option_title = option.get("title") if isinstance(option, dict) else None
        if isinstance(option_title, dict):
            title_text = option_title.get("text")
        elif isinstance(option_title, str):
            title_text = option_title
    title_text = _clean_label(title_text, 64) or f"图表 {index_hint}"

    return {
        "title": title_text,
        "option": option if isinstance(option, dict) else {},
        "png_base64": raw.get("png_base64") if isinstance(raw.get("png_base64"), str) else None,
        "key": None,
        "sheet": _clean_label(raw.get("sheet"), 64) or None,
    }


def _adapt_chart_groups(source: str, payload: Any) -> list[dict[str, Any]]:
    """把三种来源形态统一适配为 ``chart_group`` 列表。

    Args:
        source: ``direct`` / ``by_section`` / ``charts``。
        payload: 对应来源的原始数据。

    Returns:
        ``[{"group": str, "charts": [chart_item, ...]}, ...]``
        每张图补 ``key = f"{group}|{index}"``（组内序号从 1 开始）。

    三种形态：
        - ``direct``：前端直传 ``[{title, option, sheet, png_base64}]``，
          按 ``sheet`` 分组，无 sheet 归「默认」；
        - ``by_section``：LLM 报告里的 ``chart_option["by_section"]``，
          ``{章节名: [图...]}``，按章节名分组；
        - ``charts``：孤儿图列表，全部归入「数据图表」。
    """
    source_key = str(source or "").strip().lower()
    groups: dict[str, list[dict[str, Any]]] = {}
    order: list[str] = []

    def add_group(group_name: str, raw_charts: Any) -> None:
        """按组追加图项，并保持首次出现顺序。"""
        group = _clean_label(group_name, 64) or "默认"
        if not isinstance(raw_charts, list):
            return
        if group not in groups:
            groups[group] = []
            order.append(group)
        for raw in raw_charts:
            item = _normalize_chart_item(raw, len(groups[group]) + 1)
            if item is not None:
                groups[group].append(item)

    if source_key == "by_section":
        if isinstance(payload, dict):
            for section_name, charts in payload.items():
                add_group(str(section_name), charts)
    elif source_key == "charts":
        add_group("数据图表", payload if isinstance(payload, list) else [])
    elif source_key == "direct":
        if isinstance(payload, dict) and isinstance(payload.get("charts"), list):
            payload = payload["charts"]
        if isinstance(payload, list):
            # 按 sheet 二次归组：无 sheet 的图统一落到「默认」。
            for raw in payload:
                sheet = ""
                if isinstance(raw, dict):
                    sheet = _clean_label(raw.get("sheet"), 64)
                add_group(sheet or "默认", [raw])
    else:
        # 未知 source 时优先按列表直传处理，尽量不丢图。
        add_group("数据图表", payload if isinstance(payload, list) else [])

    result: list[dict[str, Any]] = []
    for group in order:
        charts = groups.get(group) or []
        if not charts:
            continue
        for index, chart in enumerate(charts, start=1):
            chart["key"] = f"{group}|{index}"
        result.append({"group": group, "charts": charts})
    return result


# ============================================================
# 5) KPI 指标卡选择
# ============================================================

def _sheet_names(indicators: dict[str, Any]) -> set[str]:
    """尽力收集数据集里的 sheet 名，用于拆分 ``<key>::<sheet>`` 形态。

    来源优先级：
        1. ``kpi_matrix_by_sheet`` 的 key（``"sheet::KPI"``）；
        2. ``indicators["sheets"]``（可能是 ``[{name: ...}]`` 或字符串列表）。
    """
    names: set[str] = set()
    by_sheet = indicators.get("kpi_matrix_by_sheet")
    if isinstance(by_sheet, dict):
        for key in by_sheet:
            head = str(key).split("::", 1)[0].strip()
            if head and head != "KPI":
                names.add(head)
    sheets = indicators.get("sheets")
    if isinstance(sheets, list):
        for item in sheets:
            if isinstance(item, dict):
                name = str(item.get("name") or "").strip()
            else:
                name = str(item or "").strip()
            if name:
                names.add(name)
    # 从 numeric_summary 自身推断：同一个前缀在多列上重复出现，
    # 基本只可能是 analyzer 的 "<sheet>::<列名>" 拼接前缀。
    numeric = indicators.get("numeric_summary")
    if isinstance(numeric, dict):
        prefix_counts: dict[str, int] = {}
        for key in numeric:
            text = str(key)
            if "::" not in text:
                continue
            prefix = text.split("::", 1)[0].strip()
            if prefix:
                prefix_counts[prefix] = prefix_counts.get(prefix, 0) + 1
        for prefix, count in prefix_counts.items():
            if count > 1:
                names.add(prefix)
    return names


def _split_sheet_key(key: Any, sheet_names: set[str]) -> tuple[str, str]:
    """拆分多 sheet 合并后的 ``numeric_summary`` key。

    现实中有两种拼接顺序（历史实现与需求描述不一致）：
        - analyzer 现状：``"<sheet>::<列名>"``
        - 需求描述：``"<列名>::<sheet>"``
    这里用已知 sheet 名做判定，两种都能拆对；都判断不出时按需求描述
    的 ``<列名>::<sheet>`` 兜底。
    """
    text = str(key or "")
    if "::" not in text:
        return text, ""
    parts = text.split("::")
    head, tail = parts[0].strip(), parts[-1].strip()
    if head in sheet_names and tail not in sheet_names:
        return tail, head
    if tail in sheet_names and head not in sheet_names:
        return head, tail
    return head, tail


def _select_kpi_cards(
    indicators: dict[str, Any],
    limit: int = _MAX_KPI_CARDS,
) -> list[dict[str, str]]:
    """挑选最多 ``limit`` 张 KPI 卡，按优先级取满即止。

    优先级（与需求一致）：
        1. KPI 宽表 ``records``（目标 / 实际 / 达成率 / 同比）；
        2. 数值列汇总 ``numeric_summary``（合计 / 均值 / 最大值 / 来源 sheet）；
        3. 数据规模 ``row_count`` / ``column_count``。

    "无数据不生成占位卡"：
        实际值缺失、汇总值缺失时直接跳过该指标，而不是写 "—" 占位，
        避免用户以为报表里有这个指标、只是数字没算出来。
    """
    if not isinstance(indicators, dict):
        indicators = {}
    cards: list[dict[str, str]] = []

    # ---- 1) KPI 宽表 ----
    kpi = indicators.get("kpi_matrix")
    if not (isinstance(kpi, dict) and kpi.get("records")):
        by_sheet = indicators.get("kpi_matrix_by_sheet")
        kpi = None
        if isinstance(by_sheet, dict):
            for candidate in by_sheet.values():
                if isinstance(candidate, dict) and candidate.get("records"):
                    kpi = candidate
                    break

    records = (kpi or {}).get("records") or [] if isinstance(kpi, dict) else []
    for record in records:
        if len(cards) >= limit:
            break
        if not isinstance(record, dict):
            continue
        metric = _clean_label(record.get("指标"), 64)
        actual = _safe_float(record.get("实际"))
        if not metric or actual is None:
            continue
        unit = str(record.get("单位") or "").strip()
        value_text = _format_metric(actual, metric, unit)

        # 达成率优先取已换算好的 "达成率%"；老数据可能只有 "达成率"(0.936)。
        rate_pct = _safe_float(record.get("达成率%"))
        if rate_pct is None:
            rate = _safe_float(record.get("达成率"))
            if rate is not None and abs(rate) <= 1.5:
                rate_pct = rate * 100
        sub = _format_percent_value(rate_pct) if rate_pct is not None else "—"

        note_parts: list[str] = []
        target = _safe_float(record.get("目标"))
        if target is not None:
            note_parts.append(f"目标 {_format_metric(target, metric, unit)}")
        yoy_pct = _safe_float(record.get("同比%"))
        if yoy_pct is None:
            yoy = _safe_float(record.get("同比"))
            if yoy is not None:
                yoy_pct = yoy * 100 if abs(yoy) <= 1.5 else yoy
        if yoy_pct is not None:
            note_parts.append(f"同比 {yoy_pct:+.1f}%")

        cards.append({
            "label": metric,
            "value": value_text,
            "sub": sub,
            "note": " · ".join(note_parts),
        })

    # ---- 2) 数值列汇总 ----
    sheet_names = _sheet_names(indicators)
    numeric = indicators.get("numeric_summary")
    if isinstance(numeric, dict):
        for key, aggregate in numeric.items():
            if len(cards) >= limit:
                break
            if not isinstance(aggregate, dict):
                continue
            column, sheet = _split_sheet_key(key, sheet_names)
            column = _clean_label(column, 64) or str(key)
            total = _safe_float(aggregate.get("sum"))
            if total is None:
                continue
            mean_value = _safe_float(aggregate.get("mean"))
            max_value = _safe_float(aggregate.get("max"))
            note = (
                f"最大值 {_format_metric(max_value, column)}"
                if max_value is not None
                else "最大值 —"
            )
            if sheet:
                note += f" · 来源 {sheet}"
            cards.append({
                "label": f"{column} 合计",
                "value": _format_metric(total, column),
                "sub": (
                    f"均值 {_format_metric(mean_value, column)}"
                    if mean_value is not None
                    else "—"
                ),
                "note": note,
            })

    # ---- 3) 数据规模 ----
    if len(cards) < limit:
        row_count = _safe_float(indicators.get("row_count"))
        if row_count is not None and row_count > 0:
            column_count = _safe_float(indicators.get("column_count"))
            cards.append({
                "label": "数据行数",
                "value": _format_number(row_count),
                "sub": "—",
                "note": (
                    f"字段数 {_format_number(column_count)}"
                    if column_count is not None
                    else ""
                ),
            })

    return cards[:limit]


# ============================================================
# 6) 明细段归一化 + AI 结论（LLM + 溯源 + 规则降级）
# ============================================================

def _normalize_detail_tables(detail_tables: Any) -> list[dict[str, Any]]:
    """把调用方传入的明细表归一成 ``detail_blocks``。

    兼容两种常见形态：
        - ``list[{title, columns, rows, note, total_rows}]``；
        - ``dict{章节名: {...}}``（``build_section_detail_tables`` 的返回值）。
    """
    if isinstance(detail_tables, dict):
        raw_items: list[Any] = []
        for name, value in detail_tables.items():
            if isinstance(value, dict):
                item = dict(value)
                item.setdefault("title", str(name))
                raw_items.append(item)
        source_items = raw_items
    elif isinstance(detail_tables, list):
        source_items = detail_tables
    else:
        return []

    blocks: list[dict[str, Any]] = []
    for index, item in enumerate(source_items, start=1):
        if not isinstance(item, dict):
            continue
        columns_raw = item.get("columns")
        columns = (
            [_clean_label(col, 64) or f"列{pos + 1}" for pos, col in enumerate(columns_raw)]
            if isinstance(columns_raw, list)
            else []
        )
        rows_raw = item.get("rows")
        rows: list[list[Any]] = []
        if isinstance(rows_raw, list):
            for row in rows_raw:
                if isinstance(row, (list, tuple)):
                    rows.append(list(row))
                else:
                    rows.append([row])
        title = _clean_label(item.get("title"), 64) or f"数据明细 {index}"
        note = str(item.get("note") or "").strip()
        total_rows = _safe_float(item.get("total_rows"))
        if (
            not note
            and total_rows is not None
            and total_rows > len(rows)
            and rows
        ):
            note = f"共 {int(total_rows)} 行，仅展示前 {len(rows)} 行"
        blocks.append({
            "title": title,
            "columns": columns,
            "rows": rows,
            "note": note,
        })
    return blocks


_SUMMARY_SYSTEM_PROMPT = (
    "你是一位严谨的管理层数据分析顾问。你的结论会直接印在可下载的报表上，"
    "因此数字准确性优先于文采。"
)


def _slim_indicators_for_summary(indicators: dict[str, Any]) -> dict[str, Any]:
    """裁剪指标 JSON，给 LLM 一份"可读体积"的上下文。

    裁剪原则：
        - KPI 记录信息密度最高，优先保留（最多 8 条）；
        - 数值汇总保留 6 列（sum/mean/max），跨量纲的 min/median 不喂；
        - by_group 最多 3 组、每组 5 项；
        - trend 只留 1 条，避免长序列挤爆 prompt。
    """
    if not isinstance(indicators, dict):
        return {}
    slim: dict[str, Any] = {}
    for key in ("row_count", "column_count"):
        value = indicators.get(key)
        if value is not None:
            slim[key] = value

    kpi = indicators.get("kpi_matrix")
    if not (isinstance(kpi, dict) and kpi.get("records")):
        by_sheet = indicators.get("kpi_matrix_by_sheet")
        kpi = None
        if isinstance(by_sheet, dict):
            for candidate in by_sheet.values():
                if isinstance(candidate, dict) and candidate.get("records"):
                    kpi = candidate
                    break
    records = (kpi or {}).get("records") or [] if isinstance(kpi, dict) else []
    if records:
        slim["kpi_matrix"] = {
            "records": [
                {
                    key: record.get(key)
                    for key in ("指标", "单位", "目标", "实际", "达成率%", "同比%")
                    if record.get(key) is not None
                }
                for record in records[:8]
                if isinstance(record, dict)
            ]
        }

    numeric = indicators.get("numeric_summary")
    if isinstance(numeric, dict):
        slim["numeric_summary"] = {
            str(key): {
                stat: aggregate.get(stat)
                for stat in ("sum", "mean", "max")
                if aggregate.get(stat) is not None
            }
            for key, aggregate in list(numeric.items())[:6]
            if isinstance(aggregate, dict)
        }

    groups = indicators.get("by_group")
    if isinstance(groups, dict):
        slim["by_group"] = {
            str(key): items[:5]
            for key, items in list(groups.items())[:3]
            if isinstance(items, list)
        }

    trend = indicators.get("trend")
    if isinstance(trend, dict) and trend:
        first_key = next(iter(trend))
        first_value = trend.get(first_key)
        if isinstance(first_value, dict):
            slim["trend"] = {
                str(first_key): {
                    "x": first_value.get("x"),
                    "y": first_value.get("y"),
                }
            }
    return slim


def _rule_conclusion(indicators: dict[str, Any]) -> str:
    """规则降级结论：只用真实指标拼确定性句子。

    为什么必须存在：
        LLM 可能超时、断网、输出编造数字。报表不能因此没有结论，
        规则结论保证"永远有结论，且每个数字都来自指标数据"。
    """
    if not isinstance(indicators, dict):
        indicators = {}
    parts: list[str] = []

    row_count = _safe_float(indicators.get("row_count"))
    column_count = _safe_float(indicators.get("column_count"))
    if row_count is not None and row_count > 0:
        scale = f"本期共 {_format_number(row_count)} 行数据"
        if column_count is not None:
            scale += f"、{_format_number(column_count)} 个字段"
        parts.append(scale + "。")

    # 数值列：只取第一列有 sum 的；比率列的 sum 无业务意义，
    # 因此比率列只报均值/最大值，不报"合计"。
    numeric = indicators.get("numeric_summary")
    sheet_names = _sheet_names(indicators)
    if isinstance(numeric, dict):
        for key, aggregate in numeric.items():
            if not isinstance(aggregate, dict):
                continue
            column, sheet = _split_sheet_key(key, sheet_names)
            column = _clean_label(column, 64) or str(key)
            source = f"（来源 {sheet}）" if sheet else ""
            if _is_ratio_name(column):
                mean_value = _safe_float(aggregate.get("mean"))
                max_value = _safe_float(aggregate.get("max"))
                if mean_value is None and max_value is None:
                    continue
                sentence = f"「{column}」"
                if mean_value is not None:
                    sentence += f"均值 {_format_metric(mean_value, column)}"
                if max_value is not None:
                    sentence += f"，最高 {_format_metric(max_value, column)}"
                parts.append(sentence + source + "。")
                break
            total = _safe_float(aggregate.get("sum"))
            if total is None:
                continue
            mean_value = _safe_float(aggregate.get("mean"))
            max_value = _safe_float(aggregate.get("max"))
            sentence = f"「{column}」合计 {_format_metric(total, column)}"
            if mean_value is not None:
                sentence += f"，均值 {_format_metric(mean_value, column)}"
            if max_value is not None:
                sentence += f"，最高 {_format_metric(max_value, column)}"
            parts.append(sentence + source + "。")
            break

    # 分组结构：最高分组 + 占比（占比由分组值确定性计算，可复算）。
    groups = indicators.get("by_group")
    if isinstance(groups, dict):
        for key, items in groups.items():
            if not isinstance(items, list):
                continue
            valid = [
                (str(item.get("name")), _safe_float(item.get("value")))
                for item in items
                if isinstance(item, dict) and _safe_float(item.get("value")) is not None
            ]
            if not valid:
                continue
            total = math.fsum(value for _, value in valid if value is not None)
            if abs(total) < 1e-12:
                continue
            top_name, top_value = max(valid, key=lambda pair: pair[1] or 0)
            share = (top_value or 0) / total * 100
            group_label = _humanize_group_key(key) or _clean_label(key, 64) or str(key)
            parts.append(
                f"按「{group_label}」分组，最高的是「{top_name}」"
                f"（{_format_number(top_value)}，占 {share:.1f}%）。"
            )
            break

    if not parts:
        return "暂无可量化的指标数据，建议结合数据明细表查看原始记录。"
    return "".join(parts)


def _humanize_group_key(key: Any) -> str:
    """把 by_group 的内部 key 转成「人话」维度名。

    为什么需要：analyzer 生成的聚合 key 形如 ``<sheet>::<维度>_by_<度量>``
    （例如 ``销售明细::渠道_by_数量``）。如果原样写进 AI 结论，用户会看到
    「按「销售明细::渠道_by_数量」分组」这种技术细节，既不专业也难读。
    这里剥掉 sheet 前缀与 ``_by_`` 后缀，只保留真正的维度名（渠道）。
    """
    text = str(key or "").strip()
    if "::" in text:
        # 取最右一段：兼容 "<sheet>::<维度>_by_<度量>" 与 "<维度>::<sheet>" 两种形态
        text = text.split("::")[-1].strip()
    if "_by_" in text:
        text = text.split("_by_", 1)[0].strip()
    return text


def _build_audit_pool(indicators: dict[str, Any]) -> list[float]:
    """构造结论溯源数值池。

    为什么要在此基础上补 row_count / column_count：
        既有 ``_collect_indicator_values`` 刻意只收录"业务指标"字段，
        不收录 ``row_count`` / ``column_count``（它们属于数据规模元信息）。
        但本报表的 LLM prompt 会把这两个值喂给模型，模型写「本期共 120 行」
        是完全正确的；若池里没有 120，就会把真实数字误判为编造，
        导致 LLM 路径几乎必然被整条作废。这里补上这两个真实值，
        其余数字仍全部来自 ``_collect_indicator_values``。
    """
    pool = _collect_indicator_values(indicators)
    for key in ("row_count", "column_count"):
        value = _safe_float(indicators.get(key))
        if value is not None:
            pool.append(value)

    # ---- 补充「年份」类数字 ----
    # 为什么必须补：指标里的日期是以**字符串**形态存放的（如 "2026-10-18 00:00:00"），
    # _collect_indicator_values 只抽数值型字段，抽不到这些字符串里的年份。
    # 于是模型只要写一句「2026 年 10 月销售额最高」，审计就会把 2026 判为
    # 「无出处数字」并把整条结论作废 —— 实测 LLM 路径因此几乎必然降级成规则文本。
    # 这里扫一遍指标文本，把真实出现过的 19xx/20xx 年份视为可溯源。
    # 只放行「年份」这一窄口径，其它 4 位数字不会因此被误放行。
    try:
        text = json.dumps(indicators, ensure_ascii=False, default=json_default)
    except (TypeError, ValueError):
        text = ""
    for year in set(re.findall(r"\b(?:19|20)\d{2}\b", text)):
        value = _safe_float(year)
        if value is not None:
            pool.append(value)
    return pool


def summarize_visual_report(
    indicators: dict[str, Any],
    chart_titles: list[str],
) -> dict[str, str]:
    """生成报表结论，返回 ``{"text": str, "source": "llm"|"rule"}``。

    LLM 约束（三层防线）：
        1. Prompt 硬性要求"每个数字必须能在指标 JSON 找到原值"；
        2. 输出后用 ``_collect_indicator_values`` 造数值池，
           交给 ``audit_section_numbers`` 做溯源；只要检出无出处数字，
           整条结论作废 —— 哪怕只错一个数字也不能上报表；
        3. 任何异常/超时/解析失败 → 规则降级，保证永远有结论。

    【并发风险】LLM 调用**串行单次**。项目实测并发请求会让单路耗时
    从 60s 拉到 600s，因此这里绝不并发、不重试风暴。
    """
    if not isinstance(indicators, dict):
        indicators = {}
    titles = [
        _clean_label(title, 64)
        for title in (chart_titles or [])
        if _clean_label(title, 64)
    ][:12]

    rule_result = {"text": _rule_conclusion(indicators), "source": "rule"}

    slim = _slim_indicators_for_summary(indicators)
    try:
        payload = json.dumps(slim, ensure_ascii=False, default=json_default)
    except (TypeError, ValueError):
        payload = "{}"
    if len(payload) > _MAX_SUMMARY_PAYLOAD_CHARS:
        payload = payload[:_MAX_SUMMARY_PAYLOAD_CHARS]

    user = (
        "请为一份可视化报表写 3~4 句中文结论（总字数不超过 180 字），"
        "覆盖整体规模、核心指标表现、结构或趋势要点。\n"
        "【硬性要求】\n"
        "  - 结论中每个数字必须能在下方指标 JSON 中找到原值；\n"
        "  - 严禁换算单位、严禁编造字段；\n"
        "  - 宁可只写定性描述，也不要写错数字；\n"
        "  - 不要输出 Markdown、JSON 或思考过程，仅返回结论文本。\n"
        f"指标JSON：{payload}\n"
        f"图表标题：{json.dumps(titles, ensure_ascii=False)}"
    )

    try:
        raw = llm.chat(
            _SUMMARY_SYSTEM_PROMPT,
            user,
            json_mode=False,
            temperature=_SUMMARY_TEMPERATURE,
            max_tokens=_MAX_SUMMARY_OUTPUT_TOKENS,
        )
        text = _normalize_summary_text(strip_thinking(str(raw or "")).strip()).strip()
    except Exception as exc:  # noqa: BLE001 - 任何 LLM 失败都必须降级，不能冒泡
        logger.warning("[dashboard_report] AI 结论生成失败，降级为规则结论: %s", exc)
        return rule_result

    if not text:
        logger.warning("[dashboard_report] AI 结论为空，降级为规则结论")
        return rule_result

    # ---- 溯源复核：无出处数字 → 整条作废 ----
    pool = _build_audit_pool(indicators)
    if pool:
        bad_numbers = audit_section_numbers("报表结论", text, pool)
        if bad_numbers:
            logger.warning(
                "[dashboard_report] AI 结论含无出处数字 %s，整条作废并降级",
                bad_numbers,
            )
            return rule_result
    else:
        # 指标池为空时任何数字都无法溯源，直接走规则结论（规则结论此时也不会写数字）。
        logger.warning("[dashboard_report] 指标数值池为空，AI 结论降级为规则结论")
        return rule_result

    return {"text": text, "source": "llm"}


# ============================================================
# 7) 报表模型组装（公共 API）
# ============================================================

def build_visual_report(
    title: str,
    style_key: str,
    indicators: dict[str, Any],
    chart_groups: list[dict[str, Any]],
    *,
    dataset_name: str = "",
    summary: str | None = None,
    detail_tables: list[Any] | None = None,
    include_ai: bool = True,
) -> dict[str, Any]:
    """组装 ``report_model``（不含 HTTP / 数据库副作用）。

    Args:
        title: 报表标题；空则用"可视化报表"。
        style_key: 图表风格 key，用于取主色与风格名。
        indicators: 指标 dict（可多 sheet 合并形态）。
        chart_groups: 已分组的图表列表（``_adapt_chart_groups`` 的输出）。
        dataset_name: 数据集名，用于副标题首段。
        summary: 已有报告摘要；非空时优先使用，不再调 LLM。
        detail_tables: 追加在"图表反解明细"之后的章节明细表。
        include_ai: ``False`` 时不调用 LLM（结论留空，由调用方决定）。

    Returns:
        严格符合 ``report_model`` 契约的 dict。
    """
    if not isinstance(indicators, dict):
        indicators = {}

    preset = get_style_preset(style_key)
    colors = preset.get("color") if isinstance(preset, dict) else None
    accent = (
        str(colors[0])
        if isinstance(colors, list) and colors
        else _ACCENT_FALLBACK
    )
    style_label = str((preset or {}).get("label") or "").strip()

    # —— 副标题：数据集名 · 生成时间 · 风格名（数据集名为空则省略首段）——
    subtitle_parts: list[str] = []
    dataset_text = str(dataset_name or "").strip()
    if dataset_text:
        subtitle_parts.append(dataset_text)
    subtitle_parts.append(f"生成于 {_now_label()}")
    if style_label:
        subtitle_parts.append(style_label)
    subtitle = " · ".join(subtitle_parts)

    # —— 图表分组：透传并补 key / insight ——
    normalized_groups: list[dict[str, Any]] = []
    chart_titles: list[str] = []
    raw_groups = chart_groups if isinstance(chart_groups, list) else []
    for raw_group in raw_groups:
        if not isinstance(raw_group, dict):
            continue
        group_name = _clean_label(raw_group.get("group"), 64) or "默认"
        raw_charts = raw_group.get("charts")
        if not isinstance(raw_charts, list):
            raw_charts = []
        charts_out: list[dict[str, Any]] = []
        for index, raw_chart in enumerate(raw_charts, start=1):
            item = _normalize_chart_item(raw_chart, index)
            if item is None:
                continue
            item["key"] = f"{group_name}|{index}"
            item["sheet"] = item.get("sheet") or group_name
            item["insight"] = _chart_insight(item.get("option") or {}, item["title"])
            charts_out.append(item)
            chart_titles.append(item["title"])
        if charts_out:
            normalized_groups.append({"group": group_name, "charts": charts_out})
    chart_groups = normalized_groups

    # —— KPI 卡 ——
    kpi_cards = _select_kpi_cards(indicators, limit=_MAX_KPI_CARDS)

    # —— 结论：报告摘要 > AI 结论 > 留空 ——
    if isinstance(summary, str) and summary.strip():
        conclusion = {"text": summary.strip(), "source": "report"}
    elif include_ai:
        try:
            conclusion = summarize_visual_report(indicators, chart_titles)
        except Exception as exc:  # noqa: BLE001 - 结论生成永远不能阻断报表
            logger.warning("[dashboard_report] 结论生成异常，降级规则结论: %s", exc)
            conclusion = {"text": _rule_conclusion(indicators), "source": "rule"}
    else:
        conclusion = {"text": "", "source": "rule"}

    # —— 明细段：图表反解（每图一段）在前，章节明细表在后 ——
    detail_blocks: list[dict[str, Any]] = []
    for group in chart_groups:
        for chart in group.get("charts") or []:
            detail = _extract_chart_detail(chart.get("option") or {})
            detail_blocks.append({
                "title": chart.get("title") or "图表明细",
                "columns": detail.get("columns") or [],
                "rows": detail.get("rows") or [],
                "note": detail.get("note") or "",
            })
    detail_blocks.extend(_normalize_detail_tables(detail_tables))

    return {
        "title": str(title or "").strip() or "可视化报表",
        "subtitle": subtitle,
        "accent": accent,
        "style_label": style_label,
        "kpi_cards": kpi_cards,
        "conclusion": conclusion,
        "detail_blocks": detail_blocks,
        "chart_groups": chart_groups,
    }


# ============================================================
# 8) Excel 渲染：公共样式 helpers
# ============================================================

def _thin_border() -> Border:
    """细边框（#e2e8f0），用于卡面 / 表头 / 数据行。"""
    side = Side(style="thin", color=_BORDER_COLOR)
    return Border(left=side, right=side, top=side, bottom=side)


def _card_border(is_top: bool, accent_argb: str) -> Border:
    """KPI 卡边框：顶行加 2pt accent 色条，其余细边框。"""
    thin = Side(style="thin", color=_BORDER_COLOR)
    return Border(
        left=thin,
        right=thin,
        bottom=thin,
        top=Side(style="thick", color=accent_argb) if is_top else thin,
    )


def _set_cell_style(
    cell: Any,
    *,
    font: Font | None = None,
    fill_color: str | None = None,
    border: Border | None = None,
    alignment: Alignment | None = None,
) -> None:
    """集中设置单元格样式，避免每个调用点重复写 4 行。"""
    if font is not None:
        cell.font = font
    if fill_color is not None:
        cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
    if border is not None:
        cell.border = border
    if alignment is not None:
        cell.alignment = alignment


def _build_overview_sheet(workbook: Workbook, report_model: dict[str, Any]) -> None:
    """渲染「报表概览」：标题 + 副标题 + KPI 卡 + AI 结论 + 脚注。"""
    ws = workbook.active
    ws.title = "报表概览"
    ws.sheet_view.showGridLines = False

    accent = _accent_argb(report_model.get("accent"))
    for col in range(1, _OVERVIEW_COLS + 1):
        ws.column_dimensions[get_column_letter(col)].width = 12

    # —— 标题（16pt 加粗）——
    ws.cell(row=1, column=1, value=str(report_model.get("title") or "可视化报表"))
    _merge_and_fill(ws, 1, 1, _OVERVIEW_COLS)
    ws.cell(row=1, column=1).font = Font(size=_FONT_TITLE, bold=True, color=_TEXT_PRIMARY)
    ws.cell(row=1, column=1).alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30

    # —— 副标题（灰色小字）——
    ws.cell(row=2, column=1, value=str(report_model.get("subtitle") or ""))
    _merge_and_fill(ws, 2, 1, _OVERVIEW_COLS)
    ws.cell(row=2, column=1).font = Font(size=_FONT_SUBTITLE, color=_TEXT_MUTED)
    ws.cell(row=2, column=1).alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 18

    # —— KPI 卡：4 张横排，每张 3 列 × 3 行 ——
    card_start_row = 4
    cards = report_model.get("kpi_cards")
    cards = cards if isinstance(cards, list) else []
    for index, card in enumerate(cards[:_MAX_KPI_CARDS]):
        if not isinstance(card, dict):
            continue
        col_start = 1 + index * _CARD_COLS
        col_end = col_start + _CARD_COLS - 1
        rows = range(card_start_row, card_start_row + _CARD_ROWS)
        for row in rows:
            for col in range(col_start, col_end + 1):
                _set_cell_style(
                    ws.cell(row=row, column=col),
                    fill_color=_CARD_BG,
                    border=_card_border(row == card_start_row, accent),
                )

        sub_text = str(card.get("sub") or "").strip()
        note_text = str(card.get("note") or "").strip()
        footer_parts = [
            part for part in (sub_text, note_text)
            if part and part not in ("—", "-")
        ]
        footer = " · ".join(footer_parts) or "—"

        payload = (
            (card_start_row, str(card.get("label") or ""), _FONT_CARD_LABEL,
             False, _TEXT_MUTED, "left"),
            (card_start_row + 1, str(card.get("value") or "—"), _FONT_CARD_VALUE,
             True, accent, "left"),
            (card_start_row + 2, footer, _FONT_CARD_SUB, False, _TEXT_MUTED, "left"),
        )
        for row, text, size, bold, color, align in payload:
            ws.cell(row=row, column=col_start, value=text)
            _merge_and_fill(ws, row, col_start, col_end, color=_CARD_BG)
            cell = ws.cell(row=row, column=col_start)
            cell.font = Font(size=size, bold=bold, color=color)
            cell.alignment = Alignment(horizontal=align, vertical="center")
            # 合并后再补一次边框，保证外轮廓完整。
            cell.border = _card_border(row == card_start_row, accent)
        ws.row_dimensions[card_start_row].height = 20
        ws.row_dimensions[card_start_row + 1].height = 32
        ws.row_dimensions[card_start_row + 2].height = 18

    # —— AI 结论区（浅蓝底提示框）——
    conclusion = report_model.get("conclusion")
    conclusion = conclusion if isinstance(conclusion, dict) else {}
    conclusion_text = str(conclusion.get("text") or "").strip()
    if not conclusion_text:
        conclusion_text = "暂无结论（可重新生成或补充指标数据）。"

    header_row = card_start_row + _CARD_ROWS + 2
    ws.cell(row=header_row, column=1, value="AI 结论")
    _merge_and_fill(ws, header_row, 1, _OVERVIEW_COLS)
    ws.cell(row=header_row, column=1).font = Font(
        size=_FONT_SECTION, bold=True, color=_TEXT_PRIMARY,
    )
    ws.row_dimensions[header_row].height = 22

    body_row = header_row + 1
    ws.cell(row=body_row, column=1, value=conclusion_text)
    _merge_and_fill(ws, body_row, 1, _OVERVIEW_COLS, color=_INFO_BG)
    body_cell = ws.cell(row=body_row, column=1)
    body_cell.font = Font(size=_FONT_BODY, color=_TEXT_PRIMARY)
    body_cell.alignment = Alignment(wrap_text=True, vertical="top")
    body_cell.border = Border(
        left=Side(style="thin", color=_INFO_BORDER),
        right=Side(style="thin", color=_INFO_BORDER),
        top=Side(style="thin", color=_INFO_BORDER),
        bottom=Side(style="thin", color=_INFO_BORDER),
    )
    ws.row_dimensions[body_row].height = _estimate_wrapped_height(conclusion_text)

    # —— 脚注：数据来源 / 生成时间 / 结论来源 ——
    source_label = {
        "llm": "AI 生成（已溯源校验）",
        "rule": "规则生成（基于真实指标）",
        "report": "报告摘要",
    }.get(str(conclusion.get("source") or ""), "—")
    footnote_rows = (
        f"数据来源：{report_model.get('subtitle') or '—'}",
        f"生成时间：{_now_label()}",
        f"结论来源：{source_label}",
    )
    footnote_start = body_row + 2
    for offset, text in enumerate(footnote_rows):
        row = footnote_start + offset
        ws.cell(row=row, column=1, value=text)
        _merge_and_fill(ws, row, 1, _OVERVIEW_COLS)
        ws.cell(row=row, column=1).font = Font(size=_FONT_NOTE, color=_TEXT_SUB)
        ws.row_dimensions[row].height = 15

    # ============================================================
    # 图表看板区：所有图表直接排在概览页里（不再每组开一个 Sheet）
    # ------------------------------------------------------------
    # 位置：脚注之后空一行作为「数据看板」板块小标题，再往下排图。
    # 版式对齐用户给的《经营看板》设计稿 —— 两列并排的图表卡片。
    # ============================================================
    charts = _flatten_charts(report_model.get("chart_groups"))
    if not charts:
        return

    dash_header_row = footnote_start + len(footnote_rows) + 1
    ws.cell(row=dash_header_row, column=1, value="数据看板")
    _merge_and_fill(ws, dash_header_row, 1, _OVERVIEW_COLS)
    ws.cell(row=dash_header_row, column=1).font = Font(
        size=_FONT_SECTION, bold=True, color=_TEXT_PRIMARY,
    )
    ws.row_dimensions[dash_header_row].height = 22

    _build_dashboard_charts(
        ws, charts, dash_header_row + 2, accent,
    )


def _build_detail_sheet(workbook: Workbook, report_model: dict[str, Any]) -> None:
    """渲染「数据明细」：按段输出 段标题 → 表头 → 数据行 → 说明。"""
    ws = workbook.create_sheet("数据明细")
    ws.sheet_view.showGridLines = False

    blocks = report_model.get("detail_blocks")
    blocks = blocks if isinstance(blocks, list) else []
    max_cols = 2
    for block in blocks:
        if isinstance(block, dict):
            max_cols = max(max_cols, len(block.get("columns") or []))
    max_cols = max(2, max_cols)

    ws.column_dimensions["A"].width = 20
    for col in range(2, max_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16

    row_cursor = 1
    for block in blocks:
        if not isinstance(block, dict):
            continue
        columns = block.get("columns") if isinstance(block.get("columns"), list) else []
        rows = block.get("rows") if isinstance(block.get("rows"), list) else []
        note = str(block.get("note") or "").strip()

        # 段标题：合并 + 加粗 + 浅底
        ws.cell(row=row_cursor, column=1, value=str(block.get("title") or "数据明细"))
        _merge_and_fill(ws, row_cursor, 1, max_cols, color=_HEAD_BG)
        ws.cell(row=row_cursor, column=1).font = Font(
            size=_FONT_SECTION, bold=True, color=_TEXT_PRIMARY,
        )
        ws.cell(row=row_cursor, column=1).alignment = Alignment(vertical="center")
        ws.row_dimensions[row_cursor].height = 22
        row_cursor += 1

        if not columns:
            ws.cell(row=row_cursor, column=1, value=note or "（暂无明细数据）")
            _merge_and_fill(ws, row_cursor, 1, max_cols)
            ws.cell(row=row_cursor, column=1).font = Font(
                size=_FONT_NOTE, color=_TEXT_MUTED,
            )
            ws.cell(row=row_cursor, column=1).alignment = Alignment(wrap_text=True)
            ws.row_dimensions[row_cursor].height = _estimate_wrapped_height(
                note or "（暂无明细数据）", chars_per_line=60,
            )
            row_cursor += 2
            continue

        # 表头行：浅底 + 加粗 + 细边框
        for col_index, column_name in enumerate(columns, start=1):
            cell = ws.cell(row=row_cursor, column=col_index, value=str(column_name))
            _set_cell_style(
                cell,
                font=Font(size=_FONT_HEADER, bold=True, color=_TEXT_PRIMARY),
                fill_color=_HEAD_BG,
                border=_thin_border(),
                alignment=Alignment(horizontal="center", vertical="center"),
            )
        ws.row_dimensions[row_cursor].height = 20
        row_cursor += 1

        # 数据行
        for row in rows:
            if not isinstance(row, (list, tuple)):
                row = [row]
            for col_index in range(1, len(columns) + 1):
                value = _excel_cell_value(row[col_index - 1]) if col_index <= len(row) else ""
                cell = ws.cell(row=row_cursor, column=col_index, value=value)
                alignment = (
                    Alignment(horizontal="right", vertical="center")
                    if isinstance(value, (int, float))
                    else Alignment(horizontal="left", vertical="center")
                )
                _set_cell_style(
                    cell,
                    font=Font(size=_FONT_BODY, color=_TEXT_PRIMARY),
                    border=_thin_border(),
                    alignment=alignment,
                )
            row_cursor += 1

        # 空段/截断说明
        if not rows:
            explanation = note or "（暂无明细数据）"
            ws.cell(row=row_cursor, column=1, value=explanation)
            _merge_and_fill(ws, row_cursor, 1, max_cols)
            ws.cell(row=row_cursor, column=1).font = Font(
                size=_FONT_NOTE, color=_TEXT_MUTED,
            )
            ws.cell(row=row_cursor, column=1).alignment = Alignment(wrap_text=True)
            ws.row_dimensions[row_cursor].height = _estimate_wrapped_height(
                explanation, chars_per_line=60,
            )
            row_cursor += 1
        elif note:
            ws.cell(row=row_cursor, column=1, value=f"说明：{note}")
            _merge_and_fill(ws, row_cursor, 1, max_cols)
            ws.cell(row=row_cursor, column=1).font = Font(
                size=_FONT_NOTE, color=_TEXT_MUTED,
            )
            ws.cell(row=row_cursor, column=1).alignment = Alignment(wrap_text=True)
            ws.row_dimensions[row_cursor].height = _estimate_wrapped_height(
                f"说明：{note}", chars_per_line=60,
            )
            row_cursor += 1

        # 段间空行
        row_cursor += 1

    if not blocks:
        ws.cell(row=1, column=1, value="（暂无明细数据）")
        _merge_and_fill(ws, 1, 1, max_cols)
        ws.cell(row=1, column=1).font = Font(size=_FONT_BODY, color=_TEXT_MUTED)


def _chart_png(chart: dict[str, Any]) -> bytes | None:
    """取得图表 PNG：前端 base64 优先，兜底 matplotlib 渲染。

    优先级的原因：
        前端 base64 是"用户编辑弹窗里看到的同一张图"，视觉最一致；
        没有 base64（或解码失败）时，用 ``render_chart_to_png`` 的 light
        主题重画，保证放进浅色 Excel 页面仍然清晰。

    【去标题】两条路径都会去掉图内标题：
        Excel 里每张图外面已有「卡片标题条」（分组名 · 图标题），
        PNG 内部再画一遍就是同一个标题显示两次。
        前端截图路径在 ``chartExport.stripOptionTitle`` 里做，
        后端兜底路径在这里做 —— 两边口径必须一致，
        否则"有前端截图"和"走兜底"的报表长得不一样。
    """
    png = _decode_png_base64(chart.get("png_base64"))
    if png and png.startswith(
        (b"\x89PNG\r\n\x1a\n", b"\xff\xd8", b"GIF8", b"RIFF")  # PNG/JPEG/GIF/WEBP
    ):
        return png
    if png:
        logger.warning("[dashboard_report] png_base64 不是可识别的图片，改走兜底渲染")
    try:
        return render_chart_to_png(
            _strip_chart_title(chart),
            width=8,
            height=5,
            dpi=150,
            theme="light",
        )
    except Exception as exc:  # noqa: BLE001 - 单图失败不能拖垮整份报表
        logger.warning("[dashboard_report] 图表兜底渲染失败: %s", exc)
        return None


def _strip_chart_title(chart: dict[str, Any]) -> dict[str, Any]:
    """返回一份**去掉了图内标题**的图表副本（不改动入参）。

    为什么去标题：Excel 报表里每张图外面已有一张卡片标题条，
    PNG 内部再画一遍标题会重复。详见前端 ``chartExport.stripOptionTitle``。
    """
    if not isinstance(chart, dict):
        return chart
    stripped = dict(chart)
    option = stripped.get("option")
    if isinstance(option, dict) and option.get("title"):
        new_option = dict(option)
        title_obj = new_option.get("title")
        if isinstance(title_obj, dict):
            new_option["title"] = {**title_obj, "text": ""}
        else:
            new_option["title"] = ""
        stripped["option"] = new_option
    return stripped


def _flatten_charts(chart_groups: Any) -> list[dict[str, Any]]:
    """把分组结构拍平成图表列表，并给每张图补上「分组名」与「全局序号」。

    为什么要拍平：现在的版式要求所有图表直接排在「报表概览」里，
    不再按分组各开一个 Sheet。但分组信息不能丢 —— 它要变成卡片上的
    「分组 · 标题」文字，让用户知道这张图来自哪个数据表。
    """
    flat: list[dict[str, Any]] = []
    groups = chart_groups if isinstance(chart_groups, list) else []
    for group in groups:
        if not isinstance(group, dict):
            continue
        group_name = _clean_label(group.get("group"), 64) or "默认"
        charts = group.get("charts")
        charts = charts if isinstance(charts, list) else []
        for chart in charts:
            if not isinstance(chart, dict):
                continue
            item = dict(chart)
            item["group_name"] = group_name
            flat.append(item)
    for index, item in enumerate(flat, start=1):
        item["seq"] = index
    return flat


def _dash_card_title(chart: dict[str, Any]) -> str:
    """卡片标题：「分组名 · 图标题」；没有标题时用「图表 N」。"""
    title = str(chart.get("title") or "").strip()
    if not title:
        title = f"图表 {chart.get('seq') or 1}"
    group_name = str(chart.get("group_name") or "").strip()
    if group_name and group_name != "默认":
        return f"{group_name} · {title}"
    return title


def _build_dashboard_charts(
    worksheet: Any,
    charts: list[dict[str, Any]],
    start_row: int,
    accent: str,
) -> int:
    """在概览页里以「两列并排卡片」排布全部图表，返回结束行号。

    版式（参照用户给的《经营看板》设计稿）::

        ┌─────────────┬─────────────┐
        │ 标题条      │ 标题条      │
        │ [图片卡面]  │ [图片卡面]  │
        │ 洞察注释    │ 洞察注释    │
        └─────────────┴─────────────┘
        ┌─────────────┬─────────────┐
        │ ...         │ ...         │
        └─────────────┴─────────────┘

    每张卡 = 1 行标题 + 15 行图片区 + 2 行洞察。图片区单元格铺浅灰底，
    透明底 PNG 叠上去就形成一个干净的"图表卡"。

    异常安全：单张图渲染失败只影响该卡（退化为 option JSON 摘要），
    绝不中断整份报表。
    """
    if not charts:
        return start_row

    row_cursor = start_row
    total = len(charts)
    for index, chart in enumerate(charts):
        # 位置计算：每行放 2 张卡，第二张从第 7 列开始
        slot = index % 2
        if slot == 0:
            if index > 0:
                row_cursor += _DASH_GAP_ROWS
            card_row = row_cursor
        else:
            card_row = row_cursor
        col_start = 1 + slot * _DASH_CARD_COLS
        col_end = col_start + _DASH_CARD_COLS - 1

        try:
            _build_dash_card(worksheet, chart, card_row, col_start, col_end, accent)
        except Exception as exc:  # noqa: BLE001 - 单卡失败不能拖垮整份报表
            logger.warning(
                "[dashboard_report] 图表卡「%s」渲染失败，已跳过: %s",
                chart.get("title"),
                exc,
            )

        # 一行放满两张后换行
        if slot == 1:
            row_cursor = card_row + _DASH_CARD_ROWS

    # 收尾：把最后一行（可能只有一张卡）的底部推进
    if total % 2 == 1:
        row_cursor += _DASH_CARD_ROWS
    return row_cursor


def _build_dash_card(
    worksheet: Any,
    chart: dict[str, Any],
    card_row: int,
    col_start: int,
    col_end: int,
    accent: str,
) -> None:
    """渲染单张图表卡：标题条 + 图片卡面 + 洞察注释。"""
    # —— 标题条 ——
    worksheet.cell(row=card_row, column=col_start, value=_dash_card_title(chart))
    _merge_and_fill(worksheet, card_row, col_start, col_end)
    title_cell = worksheet.cell(row=card_row, column=col_start)
    title_cell.font = Font(
        size=_DASH_CHART_TITLE_FONT, bold=True, color=_TEXT_PRIMARY,
    )
    title_cell.alignment = Alignment(vertical="center")
    worksheet.row_dimensions[card_row].height = _DASH_ROW_HEIGHT

    # —— 图片卡面：整块铺浅灰底，透明 PNG 叠上去形成干净卡片 ——
    image_start = card_row + _DASH_TITLE_ROWS
    image_end = image_start + _DASH_IMAGE_ROWS - 1
    for row in range(image_start, image_end + 1):
        worksheet.row_dimensions[row].height = _DASH_ROW_HEIGHT
        for col in range(col_start, col_end + 1):
            _set_cell_style(
                worksheet.cell(row=row, column=col),
                fill_color=_CARD_BG,
                border=_thin_border(),
            )

    png = _chart_png(chart)
    image_added = False
    if png:
        try:
            image = XLImage(io.BytesIO(png))
            image.width = _DASH_IMAGE_WIDTH
            image.height = _DASH_IMAGE_HEIGHT
            worksheet.add_image(image, f"{get_column_letter(col_start)}{image_start}")
            image_added = True
        except Exception as exc:  # noqa: BLE001
            logger.warning("[dashboard_report] 图片嵌入失败，退化为 JSON 摘要: %s", exc)

    if not image_added:
        # 两条图片来源都失败：写 option JSON 摘要，保证信息不丢。
        worksheet.merge_cells(
            start_row=image_start,
            start_column=col_start,
            end_row=image_end,
            end_column=col_end,
        )
        json_text = _chart_option_json(chart)
        worksheet.cell(row=image_start, column=col_start, value=json_text)
        fallback_cell = worksheet.cell(row=image_start, column=col_start)
        fallback_cell.font = Font(name="Menlo", size=_FONT_JSON, color=_TEXT_MUTED)
        fallback_cell.alignment = Alignment(wrap_text=True, vertical="top")

    # —— 洞察注释（2 行，合并换行）——
    insight_row = image_end + 1
    insight_text = str(chart.get("insight") or "（暂无图表洞察）")
    for offset in range(_DASH_INSIGHT_ROWS):
        worksheet.row_dimensions[insight_row + offset].height = _DASH_ROW_HEIGHT
    worksheet.merge_cells(
        start_row=insight_row,
        start_column=col_start,
        end_row=insight_row + _DASH_INSIGHT_ROWS - 1,
        end_column=col_end,
    )
    insight_cell = worksheet.cell(row=insight_row, column=col_start)
    insight_cell.value = insight_text
    insight_cell.font = Font(size=_FONT_NOTE, color=_TEXT_MUTED)
    insight_cell.alignment = Alignment(wrap_text=True, vertical="center")
    insight_cell.fill = PatternFill(fill_type="solid", fgColor=_WHITE)


# ============================================================
# 9) Excel 渲染入口（公共 API）
# ============================================================

def render_visual_report_excel(report_model: dict[str, Any]) -> bytes:
    """把 ``report_model`` 渲染为可下载的 ``.xlsx`` 字节流。

    Sheet 结构：
        ① 报表概览 —— 标题 / 副标题 / KPI 卡 / AI 结论 / 脚注
           **+ 数据看板**（全部图表以两列并排卡片直接排在概览页里）；
        ② 数据明细 —— 图表反解明细 + 章节明细表。

    为什么图表不再单独开 Sheet：
        用户要的是"一张能直接看的可视化报表"，而不是一堆散落的图页。
        图表排在概览页里，打开文件第一眼就是完整看板。

    异常安全：
        - 单张图失败只影响该卡（退化为 option JSON 摘要）；
        - 明细页失败也会保证文件可下载；
        - 任何情况下都返回可打开的 xlsx 字节流。
    """
    if not isinstance(report_model, dict):
        report_model = {}

    workbook = Workbook()
    _build_overview_sheet(workbook, report_model)

    try:
        _build_detail_sheet(workbook, report_model)
    except Exception as exc:  # noqa: BLE001 - 明细页失败也要保证文件可下载
        logger.warning("[dashboard_report] 数据明细页渲染失败: %s", exc)
        if "数据明细" not in {ws.title for ws in workbook.worksheets}:
            fallback = workbook.create_sheet("数据明细")
            fallback.cell(row=1, column=1, value="（数据明细渲染失败，请查看概览页看板区）")

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
