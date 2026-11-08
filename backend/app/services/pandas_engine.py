"""Pandas 数据处理引擎：解析、清洗、统计。"""
from __future__ import annotations

from pathlib import Path
from typing import Any
import logging
import re

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)


# ---------------------- 读取与基础信息 ----------------------

def list_sheet_names(file_path: str | Path) -> list[str]:
    """列出 Excel 文件中**所有**工作表名;CSV 返回单元素列表(用文件名占位)。

    说明:
        - 仅枚举 sheet 名,不读取数据,代价极低。
        - CSV 没有"工作表"概念,但为了上层 API 统一(返回 sheets 列表),
          这里返回 `["<文件名>"]`,让上层不需要分支判断。
        - 若文件不是 Excel/CSV,抛出 ValueError(上层应当校验后缀)。
    """
    p = Path(file_path)
    suffix = p.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        # pandas 1.x 起 ExcelFile.sheet_names 返回 list[str],稳定
        with pd.ExcelFile(p) as xls:
            return [str(n) for n in xls.sheet_names]
    if suffix == ".csv":
        return [p.name]
    raise ValueError(f"不支持的文件类型: {suffix}")


def read_dataframe(file_path: str | Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """根据扩展名读取 Excel/CSV 的**单张**工作表。

    说明:
        - 这是底层函数:Excel 默认读首页(sheet_name=0),CSV 整文件读。
        - **不要**直接用于"分析流水线":分析入口请用 `read_dataframe_all_sheets`,
          避免漏掉非首页数据(典型 bug —— 用户 Excel 含"销售明细/目标设定/
          渠道分析/产品分析"4 个 sheet,旧版只读首页,后 3 个 sheet 数据全丢)。
    """
    p = Path(file_path)
    suffix = p.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        return pd.read_excel(p, sheet_name=sheet_name)
    if suffix == ".csv":
        # 自动尝试编码
        for enc in ("utf-8", "gbk", "gb18030", "latin-1"):
            try:
                return pd.read_csv(p, encoding=enc)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(p, encoding="utf-8", encoding_errors="ignore")
    raise ValueError(f"不支持的文件类型: {suffix}")


# ============================================================
# 【chart-edit-prebuilt-dashboard 新增】智能表头 + 按 sheet 独立读取
# ============================================================
def _is_placeholder_row(row_vals: list[Any]) -> bool:
    """判断一行是不是"占位说明行"(典型特征:整行是中文长文本,无数字)。

    用于在 Excel 头部遇到 "销售明细（示例数据可替换）" 这种说明行时,
    自动把第一行当作 header 跳过。

    注意:
        - **不**把"整行全 nan/空"当作说明行(那种通常是 row0 空行 + row1 才是真表头),
          否则会把真正的列名行当作 placeholder 跳过,导致"Unnamed: ..."。
        - pandas 用 header=None 读时,会自动给空列起名 "Unnamed: 1" / "nan" 等,
          这种 cell 也是 placeholder 信号。
        - 浮点 NaN 用 pd.isna() 判断,不要用 str(v).strip()(因为 "nan".strip() 非空)。
    """
    if not row_vals:
        return False
    non_empty: list[Any] = []
    for v in row_vals:
        if v is None:
            continue
        try:
            if pd.isna(v):
                continue
        except (TypeError, ValueError):
            pass
        s = str(v).strip()
        if s == "":
            continue
        non_empty.append(v)
    if not non_empty:
        return False
    placeholder_count = 0
    real_count = 0
    num_count = 0
    total_str_len = 0
    for v in non_empty:
        s = str(v).strip()
        if (s.lower() == "nan"
            or s.startswith("Unnamed:")
            or s.startswith("unnamed:")):
            placeholder_count += 1
            continue
        real_count += 1
        total_str_len += len(s)
        for ch in s:
            if ch.isdigit():
                num_count += 1
                break
    if real_count == 0:
        return True
    return num_count == 0 and total_str_len >= 4


def _looks_like_data_row(values: list[Any]) -> bool:
    """判断一行"是不是数据行"(含数字 / 日期 / 大量非空短文本)。

    用于在 Excel 前几行里区分"占位说明行" vs "表头行" vs "数据行"。
    数据行特征:>=30% 的非空 cell 是数字或日期。
    """
    non_empty: list[Any] = []
    for v in values:
        if v is None:
            continue
        try:
            if pd.isna(v):
                continue
        except (TypeError, ValueError):
            pass
        s = str(v).strip()
        if s == "":
            continue
        non_empty.append(v)
    if not non_empty:
        return False
    import datetime as _dt
    num_cells = 0
    for v in non_empty:
        s = str(v).strip()
        try:
            float(s)
            num_cells += 1
            continue
        except ValueError:
            pass
        try:
            _dt.datetime.fromisoformat(s.replace("/", "-"))
            num_cells += 1
            continue
        except ValueError:
            pass
        if hasattr(v, "year") and hasattr(v, "month"):
            num_cells += 1
    return (num_cells / len(non_empty)) >= 0.3


def read_dataframe_smart_header(
    file_path: str | Path,
    sheet_name: str | int = 0,
) -> pd.DataFrame:
    """智能读取单张 sheet:自动识别"第一行是占位说明"并将其作为表头。"""
    p = Path(file_path)
    suffix = p.suffix.lower()
    # Excel / CSV 统一先用 header=None 读成"原始网格",再走同一套
    # 占位行识别 + 表头定位逻辑。CSV 旧实现直接 header=0 返回,
    # "导出说明（示例数据可替换）"这类占位首行同样会顶掉真表头。
    try:
        if suffix == ".csv":
            for enc in ("utf-8", "gbk", "gb18030", "latin-1"):
                try:
                    try:
                        raw = pd.read_csv(p, encoding=enc, header=None)
                    except pd.errors.ParserError:
                        # 锯齿 CSV(首行说明只占 1 列,真表头有 N 列)会让 C 引擎
                        # 报 "Expected N fields in line M"。改用 csv 模块逐行读,
                        # 短行右侧补 None 拼成矩形网格,再走统一表头识别。
                        import csv as _csv
                        with open(p, encoding=enc, newline="") as fh:
                            rows = [r for r in _csv.reader(fh) if r]
                        if not rows:
                            raise ValueError("CSV 文件为空")
                        width = max(len(r) for r in rows)
                        raw = pd.DataFrame(
                            [r + [None] * (width - len(r)) for r in rows]
                        )
                    break
                except UnicodeDecodeError:
                    continue
            else:
                raw = pd.read_csv(p, encoding="utf-8", encoding_errors="ignore", header=None)
        else:
            raw = pd.read_excel(p, sheet_name=sheet_name, header=None)
    except Exception:
        return read_dataframe(p, sheet_name=sheet_name)
    if raw.empty or len(raw) < 2:
        return read_dataframe(p, sheet_name=sheet_name)

    header_idx: int | None = None
    for i in range(min(5, len(raw))):
        row_vals = raw.iloc[i].tolist()
        if _is_placeholder_row(row_vals):
            continue
        # 全部 NaN 也跳过
        any_non_nan = False
        for v in row_vals:
            if v is None:
                continue
            try:
                if pd.isna(v):
                    continue
            except (TypeError, ValueError):
                pass
            any_non_nan = True
            break
        if not any_non_nan:
            continue
        if _looks_like_data_row(row_vals):
            # 数据行 → header 应该是它前面**最近的非占位非全空非数据行**;
            # 找不到则用第一行(保持原行为)
            for j in range(i - 1, -1, -1):
                prev = raw.iloc[j].tolist()
                if _is_placeholder_row(prev):
                    continue
                # 找 prev 是不是也有非空 cell 且不是数据行
                if _looks_like_data_row(prev):
                    continue
                # 确认 prev 至少有一个非 NaN cell
                has = False
                for v in prev:
                    if v is None:
                        continue
                    try:
                        if pd.isna(v):
                            continue
                    except (TypeError, ValueError):
                        pass
                    has = True
                    break
                if has:
                    header_idx = j
                    break
            if header_idx is None:
                header_idx = max(0, i - 1)
            break
        # 不是数据行 → 这就是 header
        header_idx = i
        break
    if header_idx is None:
        header_idx = 0
    # 把 header_idx 那行作为列名;header 里空白的列保持 `列N`,
    # 避免把"数据"误当"列名"(旧实现有过此类 bug)。
    header_row = raw.iloc[header_idx].tolist()
    new_header: list[str] = []
    for j, v in enumerate(header_row):
        if v is not None and str(v).strip():
            new_header.append(str(v).strip())
        else:
            new_header.append(f"列{j + 1}")
    body = raw.iloc[header_idx + 1:].copy()
    body.columns = new_header
    body.reset_index(drop=True, inplace=True)
    # 兜底:列名去重(防止两列同字符说明后撞名)
    seen: dict[str, int] = {}
    final_cols: list[str] = []
    for c in body.columns:
        if c in seen:
            seen[c] += 1
            final_cols.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            final_cols.append(c)
    body.columns = final_cols

    # ---- 清理 1:删掉列名是 "nan" / "nan.X" / "Unnamed: X" 的"占位列" ----
    # 这种列通常是 row1 也是空、只作为"占位"凑对齐的列,数据全空,
    # 留着只会稀释 _split_columns 的判断,生成的图也没业务意义。
    drop_cols = []
    for c in body.columns:
        cs = str(c).strip()
        if cs == "" or cs.lower() == "nan" or cs.startswith("Unnamed:"):
            # 再确认这一列真的全是 NaN(避免误删真有数据的列)
            try:
                if body[c].isna().all():
                    drop_cols.append(c)
            except Exception:
                pass
    if drop_cols:
        body = body.drop(columns=drop_cols)

    # ---- 清理 2:把"epoch 占位日期"列转回数值 + 删空列 ----
    # 用户 Excel 经常把数值列误用日期格式存,导致 pandas 读成 datetime 或字符串 "1970-01-01"。
    # 这些列应该被当数值用。但若转完后整列只有 1 个唯一值(全是 epoch),
    # 说明这列没填任何有意义的数据,直接 drop。
    import re as _re_epoch
    epoch_pattern = _re_epoch.compile(r"^(19|20)\d{2}-\d{1,2}-\d{1,2}$")
    epoch_drop = []
    for c in list(body.columns):
        try:
            series = body[c]
            # dtype 是 datetime64
            if series.dtype == "datetime64[ns]":
                excel_epoch = pd.Timestamp("1899-12-30")
                if (series.dt.year == 1970).mean() >= 0.5:
                    body[c] = (series - excel_epoch).dt.days.astype("Int64")
                continue
            # dtype 是 object:试 to_datetime + 1970 比例
            non_null = series.dropna().astype(str)
            if not len(non_null):
                epoch_drop.append(c)
                continue
            epoch_ratio = non_null.str.match(epoch_pattern).mean()
            if epoch_ratio < 0.5:
                continue
            # 全部转成 datetime
            coerced = pd.to_datetime(non_null, errors="coerce")
            if coerced.notna().mean() >= 0.9 and (coerced.dt.year == 1970).mean() >= 0.5:
                excel_epoch = pd.Timestamp("1899-12-30")
                body[c] = (coerced - excel_epoch).dt.days.astype("Int64")
                # 整列只有 1 个 unique 值(全是同一 epoch) → drop
                nn = body[c].dropna()
                if len(nn) and nn.nunique() <= 1:
                    epoch_drop.append(c)
        except Exception:
            pass
    if epoch_drop:
        body = body.drop(columns=epoch_drop)

    # ---- 清理 3:列名是 "nan" 的再删一次(转换后可能产生新 nan) ----
    body = body.dropna(axis=1, how="all")

    # ---- 清理 4:丑列名重命名 ----
    # 列名是 "Unnamed: N" / "nan" / "nan.N" / 空 → 重命名为 "列N"
    # 用户看 dashboardGroups 时,看到"Unnamed: 1"很迷惑
    new_col_names: list[str] = []
    for i, c in enumerate(body.columns):
        cs = str(c).strip()
        if (cs == ""
            or cs.lower() == "nan"
            or cs.startswith("Unnamed:")
            or (cs.startswith("nan.") and cs[4:].isdigit())):
            new_col_names.append(f"列{i + 1}")
        else:
            new_col_names.append(cs)
    body.columns = new_col_names

    # ---- 清理 5:全文本网格的数值化兜底 ----
    # 锯齿 CSV 走 csv 模块手动拼网格,读出的值全是字符串;这里对"全部非空值
    # 都能转成数字"的 object 列做数值化,与 pd.read_csv 的类型推断保持一致。
    # 只在 object 列上做、且要求 100% 可转,避免"合计"这类文本行被误转 NaN。
    for c in body.columns:
        if body[c].dtype != object:
            continue
        nn = body[c].dropna()
        if len(nn) and pd.to_numeric(nn, errors="coerce").notna().all():
            body[c] = pd.to_numeric(body[c], errors="coerce")
    return body


def read_dataframe_per_sheet(
    file_path: str | Path,
) -> list[dict[str, Any]]:
    """按 sheet 独立读取,返回 [{name, df}, ...]。

    与 read_dataframe_all_sheets 的差别:
        - 不再把所有 sheet 拼成一张大表,避免不同 sheet 列结构差异把列名搞乱
          (典型:Sheet1 有 5 列,Sheet2 有 8 列,concat 后出现大量 NaN 列,
           业务列被稀释,_split_columns 看到的字段严重不全)。
        - 每个 sheet 独立走 read_dataframe_smart_header,
          保证每个 sheet 内部的列名都正确。
    """
    p = Path(file_path)
    suffix = p.suffix.lower()
    out: list[dict[str, Any]] = []
    if suffix == ".csv":
        df = read_dataframe_smart_header(p)
        return [{"name": p.name, "df": df}]
    sheet_names = list_sheet_names(p)
    for name in sheet_names:
        try:
            df = read_dataframe_smart_header(p, sheet_name=name)
        except Exception:
            df = pd.DataFrame()
        out.append({"name": str(name), "df": df})
    return out


def read_dataframe_all_sheets(file_path: str | Path) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    """读取 Excel 的**所有**工作表,纵向合并为一张 DataFrame。

    设计要点:
        - **智能表头**:逐 sheet 用 read_dataframe_smart_header 读取(失败回退
          普通读取),自动跳过"示例数据可替换"这类占位标题行,避免标题行被
          当成列名、真表头沉进数据区。
        - **列名取并集**:不同 sheet 可能字段完全不一样(明细 vs 目标 vs 渠道 vs 产品),
          不能简单 concat,否则报错。这里先逐 sheet 读 → 收集列名并集 →
          每个 sheet 补齐缺失列(fill NaN) → 最后 concat。
        - **新增 `_sheet_name` 列**:每一行记录来源 sheet 名,LLM/前端据此识别
          "这行数据来自哪个表",便于跨 sheet 对比与解读。
        - **过滤全空 sheet**:若某 sheet 读出来 0 行有效数据,跳过合并,
          避免污染预览。
        - **CSV 兼容**:CSV 没有多 sheet 概念,直接读 + 返回 `sheets=[{name, rows, cols}]`,
          不引入 `_sheet_name` 列(避免给 CSV 路径引入歧义)。

    Args:
        file_path: 文件绝对路径(.xlsx/.xls/.csv)。

    Returns:
        (merged_df, sheet_info)
        - merged_df: 合并后的 DataFrame(列名并集 + 可能含 `_sheet_name`);
                     若文件无可读 sheet,返回空 df(仅含最少兜底列)。
        - sheet_info: `[{name: str, rows: int, cols: int}, ...]`,顺序与 Excel
                     工作表顺序一致;CSV 时为长度 1 的列表。
    """
    p = Path(file_path)
    suffix = p.suffix.lower()

    # CSV 走单独路径:无多 sheet 概念,直接读 + 返回 sheets=[单元素]
    if suffix == ".csv":
        df = read_dataframe(p)
        return df, [{"name": p.name, "rows": int(len(df)), "cols": int(len(df.columns))}]

    # Excel 路径:逐 sheet 读 → 收集列名并集 → 对齐后 concat
    sheet_names = list_sheet_names(p)
    if not sheet_names:
        return pd.DataFrame(), []

    frames: list[pd.DataFrame] = []
    sheet_info: list[dict[str, Any]] = []
    for name in sheet_names:
        try:
            # 优先走"智能表头"读取:Excel 模板类文件的第一行常是
            # "销售明细（示例数据可替换）"这类占位标题(甚至是合并单元格),
            # 直接 header=0 会把它当列名、把真正的表头行沉进数据区,
            # 产生 "Unnamed: N" 垃圾列 + 表头错位。smart_header 会:
            #   1) 跳过占位标题行、定位真正表头行;
            #   2) 删掉全空的占位列;
            #   3) 把残留的 "Unnamed: N"/"nan" 列重命名为 "列N"。
            # smart_header 自身异常时回退到原始读取,保证不阻塞主流程。
            try:
                sub_df = read_dataframe_smart_header(p, sheet_name=name)
            except Exception:
                sub_df = read_dataframe(p, sheet_name=name)
        except Exception:
            # 单个 sheet 读失败不阻塞整体,记录 0 行 0 列
            sheet_info.append({"name": name, "rows": 0, "cols": 0})
            continue
        # 把列名转 str,避免 Excel 里偶发的 datetime/int 列名导致后续对齐崩
        sub_df = sub_df.copy()
        sub_df.columns = [str(c) for c in sub_df.columns]
        sheet_info.append({
            "name": name,
            "rows": int(len(sub_df)),
            "cols": int(len(sub_df.columns)),
        })
        if len(sub_df) == 0:
            # 空 sheet 也保留列名信息(便于上层知道有这张表),
            # 但不参与 concat(concat 0 行 df 没必要且会带 _sheet_name 空值)
            continue
        sub_df["_sheet_name"] = str(name)
        frames.append(sub_df)

    if not frames:
        # 所有 sheet 都是空的,构造一个仅含 _sheet_name 的兜底 df
        # 这种情况极少(用户上传了一个全空 Excel),不影响主流程
        empty = pd.DataFrame(columns=["_sheet_name"])
        return empty, sheet_info

    # 列名并集
    all_cols: list[str] = []
    seen: set[str] = set()
    # 把 _sheet_name 放最末尾(且保证不被前面的业务列覆盖)
    for f in frames:
        for c in f.columns:
            if c == "_sheet_name":
                continue
            if c not in seen:
                seen.add(c)
                all_cols.append(c)
    all_cols.append("_sheet_name")

    # 对齐每个 df:缺失列补 NaN,统一列顺序
    aligned: list[pd.DataFrame] = []
    for f in frames:
        missing = [c for c in all_cols if c not in f.columns]
        for c in missing:
            f[c] = pd.NA
        aligned.append(f[all_cols])

    merged = pd.concat(aligned, ignore_index=True)
    # 把业务列的 dtype 强制 object 化(填充 NaN 时避免 FutureWarning)
    for c in all_cols[:-1]:
        if c in merged.columns:
            try:
                merged[c] = merged[c].astype(object)
            except Exception:
                pass
    return merged, sheet_info


def dataframe_to_rows(df: pd.DataFrame, limit: int = 10) -> dict[str, Any]:
    """转换为前端可消费的预览数据。

    返回结构：
        {
            "columns": ["列1", "列2", ...],
            "rows": [{"列1": 值, "列2": 值, ...}, ...],   # 对象数组，适配 Element Plus el-table
            "total": 总行数
        }

    说明：
        - 之前使用 ``head.values.tolist()`` 返回二维数组，导致前端 ``el-table`` 的
          ``:prop`` 无法匹配行内字段（表现为只渲染表头、无数据行）。
        - 现在改为 ``to_dict(orient="records")``，与 ``<el-table-column :prop="col" />`` 配合。
        - 保留 ``fillna("")`` 把 NaN 转成空串，避免前端把 NaN 显示成 "nan"。
        - 新版 pandas 在 object dtype 上调用 ``fillna`` 会触发 ``FutureWarning``（不再自动
          downcast），因此先 ``infer_objects`` 显式完成 downcast，再 ``fillna``。
    """
    # 显式触发 object dtype 的 downcast，避免 pandas FutureWarning：
    #   "Downcasting object dtype arrays on .fillna ... is deprecated"
    head = df.head(limit).infer_objects(copy=False).fillna("")
    columns = [str(c) for c in df.columns.tolist()]
    # orient="records" 会得到 [{colA: v1, colB: v2}, ...]，正是 el-table 期望的形状
    rows = head.astype(object).to_dict(orient="records")
    return {
        "columns": columns,
        "rows": rows,
        "total": int(len(df)),
    }


# ---------------------- 合计行/合计列剔除 ----------------------
#
# 背景(2026-10 修复"年度报告只有文字"根因之一):
#   业务台账 Excel 里,"合计/小计/总计/全年合计"这类行是**给人看的汇总**,
#   不是业务明细。把它们混进 groupby 求和,会导致:
#     - sum 翻倍(月度表 12 个月 4680 + 全年合计行 4680 = 9360,虚高一倍);
#     - 饼图里"合计"独占一块画布(渠道表 线上1950/门店1200/代理商900/分销630/
#       合计4680,"合计"成了最大的一块,占比 41.7%,完全失真);
#     - 分类模型把"合计"当成一个真实业务类别做归因,LLM 写出的分析全是错的。
#   所以在做任何聚合之前,必须先把这些汇总行/列从数据里摘掉。
#
# 判定采用"精确词表"而非模糊包含,避免把"小计以下明细"这类
# 真实业务名称误伤(例如产品名叫"小家电"不会被"小计"规则命中)。

# 行标签命中即视为汇总行(转小写后比较,兼容"总计/TOTAL")
_TOTAL_ROW_LABELS = frozenset({
    "合计", "小计", "总计", "总合计", "累计", "累积", "共计", "总共",
    "全年合计", "全年总计", "年度合计", "年度总计", "本期合计", "本月合计",
    "当日合计", "本日合计", "本年合计", "年合计", "月合计", "周合计", "季合计",
    "汇总", "汇总合计", "total", "sum", "subtotal", "grand total", "合计值",
})

# 列名命中即视为汇总列(如"合计金额""总计销量""total")
_TOTAL_COL_LABELS = frozenset({
    "合计", "小计", "总计", "总合计", "累计", "汇总", "total", "sum", "subtotal",
})


def _norm_label(v: Any) -> str:
    """把单元格/列名规整成"可比对的标签":去空白、全角转半角、转小写。"""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", "", str(v)).strip().lower()


def is_total_label(value: Any) -> bool:
    """判断一个单元格值是否是"合计/小计"类汇总标签。

    精确词表命中才返回 True。像"小家电""总量分析"这种包含"总"字的
    真实业务名不会被误判。
    """
    return _norm_label(value) in _TOTAL_ROW_LABELS


def _is_total_row(df: pd.DataFrame, label_cols: list[str]) -> bool:
    """判断某一行是不是汇总行:标签列中出现任一精确汇总词即命中。

    label_cols 取前几列(名称列通常在最左),避免在数值列上做无意义比对。
    """
    for c in label_cols:
        if c not in df.columns:
            continue
        try:
            if df[c].map(is_total_label).any():
                return True
        except Exception:
            continue
    return False


def find_total_rows(df: pd.DataFrame, label_cols: list[str] | None = None) -> list[int]:
    """找出所有汇总行的**位置索引**(不含表头,基于当前 df 的 index)。

    label_cols 未指定时,自动取前 3 列作为候选标签列。
    """
    if len(df) == 0:
        return []
    if label_cols is None:
        label_cols = list(df.columns[:3])
    try:
        mask = pd.Series(False, index=df.index)
        for c in label_cols:
            if c in df.columns:
                mask = mask | df[c].map(is_total_label).fillna(False)
        return [int(i) for i in df.index[mask]]
    except Exception:
        return []


def strip_total_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """剔除汇总行("合计/小计/全年合计"等),返回 (剔除后的 df, 被剔除的行标签)。

    背景见 `_TOTAL_ROW_LABELS` 上方注释:不做这一步会让 sum 翻倍、
    饼图出现"合计"伪类别、LLM 归因跑偏。

    注意:
        - 只在**文本列**上做判定,数值列不会误命中(数字不等于"合计")。
        - 全部行都是汇总行时不做剔除(返回原表),避免把表删空导致下游空指针。
    """
    if len(df) == 0:
        return df, []
    total_idx = set(find_total_rows(df))
    if not total_idx or len(total_idx) >= len(df):
        return df, []
    labels: list[str] = []
    for i in sorted(total_idx):
        row = df.loc[[i]]
        for c in row.columns:
            val = row.iloc[0][c]
            if is_total_label(val):
                labels.append(str(val))
                break
    out = df.drop(index=list(total_idx))
    return out, labels


# ---------------------- KPI 宽表(横向对照表)识别 ----------------------
#
# 背景(2026-10 修复"年度报告只有文字"根因之二):
#   业务年报里信息量最大的一张表通常长这样(年度核心KPI):
#       指标 | 单位 | 年度目标 | 年度实际 | 达成率 | 去年同期实际 | 同比
#       销售额| 万元 |    5000 |    4680 | 0.936 |        4200 | 0.114
#   这是一张**"每个指标一行、对比维度做成列"的宽表**。旧引擎把它当成
#   7 个互不相干的数值列做聚合,于是:
#     - sum=67052 —— 把 5000(万元)+60000(单)+850(元)+0.35(比率)跨量纲相加;
#     - mean/max/min 同样是混合量纲的垃圾值;
#     - LLM 拿到这种数字,只能写"收入低速增长、利润承压"这种无数字废话,
#       因为写具体数字必然写错。
#   识别出这类宽表后,应改为**逐指标**产出结构化记录:
#     {"指标":"销售额","单位":"万元","目标":5000,"实际":4680,
#      "达成率":0.936,"去年同期":4200,"同比":0.114}
#   这才是 LLM 能直接引用、图表能直接画的真实数据。

# 列名 → 语义角色的映射键(归一化后匹配)。覆盖各类常见命名。
_KPI_ROLE_KEYWORDS: dict[str, tuple[str, ...]] = {
    # 指标名称列本身
    "name": ("指标", "指标名称", "指标项", "kpi", "kpiname", "项目", "科目", "名称", "指标名"),
    # 目标
    "target": ("目标", "目标值", "年度目标", "预算", "预算值", "计划", "计划值",
               "target", "plan", "goal", "年度预算", "指标目标"),
    # 实际
    "actual": ("实际", "实际值", "年度实际", "本期", "本期值", "完成", "完成值",
               "实际完成", "actual", "real", "value", "本期实际", "本年实际"),
    # 达成率
    "rate": ("达成率", "完成率", "达成", "完成比例", "达成比例", "完成度",
             "达成情况", "rate", "attainment"),
    # 去年同期(同比基数)
    "last_year": ("去年同期", "同期", "上年同期", "上期", "同期值", "去年同期值",
                  "去年同期实际", "去年同期(万元)", "lastyear", "yoybase", "同期实际"),
    # 同比增长率
    "yoy": ("同比", "同比增长率", "同比增减", "较去年", "yoy", "yearoveryear"),
    # 环比增长率
    "mom": ("环比", "环比增长率", "较上月", "较上期", "mom", "monthovermonth"),
    # 单位(不参与数值计算,仅用于给数字加单位)
    "unit": ("单位", "计量单位", "unit"),
}

# 角色判定优先级:从"最具体"到"最宽泛",先匹配到先占位。
# 关键顺序说明:
#   1) last_year 必须在 yoy 之前 —— "去年同期实际" 含"去年",
#      若先被 yoy 的 "同比" 规则之外的宽泛规则抢走会误判;实际上两者词面不重叠,
#      但 "去年同期(万元)" 这类带括号的列名必须优先归到 last_year。
#   2) rate 在 actual 之前 —— "实际达成率" 应归 rate 而非 actual。
_ROLE_PRIORITY = ("name", "unit", "rate", "target", "last_year", "yoy", "mom", "actual")


def _match_kpi_role(col_name: Any) -> str | None:
    """把一个列名映射到 KPI 语义角色;无匹配返回 None。

    匹配策略:归一化列名(去空白/括号内容/单位后缀)后,按 _ROLE_PRIORITY
    顺序做"包含"匹配,命中即返回该角色。

    为什么要保留括号内容参与判断:
        "去年同期(万元)" 若先把括号剥掉会变成"去年同期"→ last_year,正确;
        但 "年度销售额(万元)" 剥掉括号变成"年度销售额"→ 会被 actual 的
        "年度"类词误吞。真实渠道表列名就是这种,所以**不能**无条件剥括号,
        只在括号内容是纯单位说明时才剥。
    """
    name = _norm_label(col_name)
    if not name:
        return None
    # 括号内容若是纯单位(万元/元/个/%/单),剥掉以提升匹配精度
    stripped = re.sub(r"[（(][^）)]{1,6}[）)]", "", name)
    for role in _ROLE_PRIORITY:
        for kw in _KPI_ROLE_KEYWORDS[role]:
            if kw in name or kw in stripped:
                return role
    return None


def detect_kpi_matrix(df: pd.DataFrame) -> dict[str, Any] | None:
    """识别 KPI 宽表并逐指标抽取结构化记录。

    识别条件(全部满足才认定):
        1) 表里有"指标名"列(role=name)——没有名称列就没法逐指标叙述;
        2) 至少再命中 1 个对比维度角色(target/actual/last_year/yoy/rate/mom);
        3) 行数 ≥ 2(单行不能构成"指标清单")。

    返回:
        {
          "metric_col": "指标",
          "unit_col": "单位",
          "role_cols": {"target": "年度目标", "actual": "年度实际", ...},
          "records": [
             {"指标": "销售额", "单位": "万元", "目标": 5000.0, "实际": 4680.0,
              "达成率": 0.936, "去年同期": 4200.0, "同比": 0.114},
             ...
          ],
          "dimensions": ["目标", "实际", "达成率", "去年同期", "同比"],
        }
        非 KPI 宽表时返回 None。
    """
    if len(df) < 2 or len(df.columns) < 2:
        return None

    role_cols: dict[str, str] = {}
    metric_col: str | None = None
    unit_col: str | None = None
    for c in df.columns:
        role = _match_kpi_role(c)
        if role is None:
            continue
        if role == "name" and metric_col is None:
            metric_col = c
            continue
        if role == "unit" and unit_col is None:
            unit_col = c
            continue
        # 同一角色只取第一列,避免"本期实际/本期目标"之类重复列互相覆盖
        role_cols.setdefault(role, c)

    # 条件 1:必须有指标名列
    if metric_col is None:
        return None
    # 条件 2:至少一个对比维度
    value_roles = ("target", "actual", "last_year", "yoy", "mom", "rate")
    if not any(r in role_cols for r in value_roles):
        return None

    # 抽取逐指标记录:先把角色列数值化,再逐行组装
    value_cols = {r: role_cols[r] for r in value_roles if r in role_cols}
    coerced: dict[str, pd.Series] = {}
    for role, col in value_cols.items():
        s = df[col]
        if pd.api.types.is_numeric_dtype(s):
            coerced[role] = s
        else:
            conv, _ = _coerce_numeric(s)
            coerced[role] = pd.to_numeric(conv, errors="coerce")

    records: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        name_val = row[metric_col]
        try:
            if pd.isna(name_val):
                continue
        except (TypeError, ValueError):
            pass
        name = str(name_val).strip()
        if not name or is_total_label(name):
            # 跳过空名与"合计"汇总行(合计行不是指标)
            continue
        rec: dict[str, Any] = {"指标": name}
        if unit_col is not None and unit_col in row.index:
            uv = row[unit_col]
            try:
                if not pd.isna(uv):
                    rec["单位"] = str(uv).strip()
            except (TypeError, ValueError):
                pass
        for role, s in coerced.items():
            try:
                val = s.loc[row.name]
            except Exception:
                val = None
            # 中文角色名(给 LLM 直读),值做 JSON 友好化
            label = {
                "target": "目标", "actual": "实际", "rate": "达成率",
                "last_year": "去年同期", "yoy": "同比", "mom": "环比",
            }[role]
            if val is None or pd.isna(val):
                rec[label] = None
            else:
                rec[label] = float(round(float(val), 6))
        # ---- 比率类指标额外产出"百分数形态"(2026-10 新增)----
        # 为什么需要:达成率/同比在源数据里是小数(0.936 / 0.114),
        # 但图表与业务语言习惯用百分数(93.6% / 11.4%)。
        # 若只在出图时做 ×100 换算,数值溯源校验(analyzer.verify_chart_data)
        # 会发现 "93.6 在数据里找不到"(数据里只有 0.936)而误判真实图为编造。
        # 所以在这里**同时**落一份百分数形态,让指标层与图表层量纲一致。
        # 判定"这是比率"的口径:值域落在 [0, 1.5] —— 金额/单量不可能落在这个区间。
        for pct_label, src_label in (
            ("达成率%", "达成率"), ("同比%", "同比"), ("环比%", "环比"),
        ):
            v = rec.get(src_label)
            if isinstance(v, (int, float)) and not isinstance(v, bool) \
                    and abs(float(v)) <= 1.5:
                rec[pct_label] = round(float(v) * 100, 2)
            else:
                rec[pct_label] = None
        records.append(rec)

    if len(records) < 2:
        # 只有 1 条记录,退化成普通表处理更稳妥(避免把 2 行的窄表当 KPI 表)
        return None

    return {
        "metric_col": str(metric_col),
        "unit_col": str(unit_col) if unit_col is not None else None,
        "role_cols": {k: str(v) for k, v in role_cols.items()},
        "dimensions": [
            {"target": "目标", "actual": "实际", "rate": "达成率",
             "last_year": "去年同期", "yoy": "同比", "mom": "环比"}[r]
            for r in ("target", "actual", "last_year", "yoy", "mom", "rate")
            if r in role_cols
        ],
        "records": records,
    }


def kpi_matrix_to_charts(kpi: dict[str, Any]) -> list[dict[str, Any]]:
    """把 KPI 宽表结构化记录转成可直接渲染的 ECharts option 列表。

    产出 3 类图(全部使用真实记录值,零编造):
        1. 「目标 vs 实际」分组柱状图 —— 逐指标对比,达成情况一目了然;
        2. 「达成率」条形图 —— 低于 100% 的指标一眼可见;
        3. 「同比变化」条形图 —— 正负增长方向清晰(若有同比列)。

    色板与 LLM 约束一致(见 prompt_builder 的安全色板),保证多来源图表风格统一。
    """
    records = kpi.get("records") or []
    if not records:
        return []

    labels = [str(r.get("指标", "")) for r in records]
    charts: list[dict[str, Any]] = []

    # ---- 图①:目标 vs 实际 ----
    has_tv = any(r.get("目标") is not None for r in records)
    has_av = any(r.get("实际") is not None for r in records)
    if has_tv or has_av:
        series: list[dict[str, Any]] = []
        if has_tv:
            series.append({
                "name": "目标", "type": "bar",
                "data": [r.get("目标") for r in records],
                "itemStyle": {"color": "#a78bfa"},
            })
        if has_av:
            series.append({
                "name": "实际", "type": "bar",
                "data": [r.get("实际") for r in records],
                "itemStyle": {"color": "#22d3ee"},
            })
        charts.append({
            "type": "bar",
            "title": "各指标目标 vs 实际",
            "option": {
                "title": {"text": "各指标目标 vs 实际", "textStyle": {"color": "#0e7490"}},
                "tooltip": {"trigger": "axis"},
                "legend": {"textStyle": {"color": "#475569"}},
                "grid": {"left": 60, "right": 30, "top": 60, "bottom": 80},
                "xAxis": {
                    "type": "category", "data": labels,
                    "axisLabel": {"rotate": 30, "interval": 0, "fontSize": 11,
                                  "color": "#475569"},
                },
                "yAxis": {"type": "value", "axisLabel": {"color": "#475569"}},
                "series": series,
            },
        })

    # ---- 图②:达成率 ----
    # 直接取 records 里已备好的「达成率%」字段(detect_kpi_matrix 已把
    # 小数口径折算成百分数并与源数据一起落库),不在这里二次 ×100,
    # 避免图表值与指标池量纲不一致而被溯源校验误判为编造。
    pct_vals = [r.get("达成率%") for r in records]
    if any(isinstance(v, (int, float)) for v in pct_vals):
        charts.append({
            "type": "bar",
            "title": "各指标达成率(%)",
            "option": {
                "title": {"text": "各指标达成率(%)", "textStyle": {"color": "#0e7490"}},
                "tooltip": {"trigger": "axis"},
                "grid": {"left": 60, "right": 30, "top": 60, "bottom": 80},
                "xAxis": {
                    "type": "category", "data": labels,
                    "axisLabel": {"rotate": 30, "interval": 0, "fontSize": 11,
                                  "color": "#475569"},
                },
                "yAxis": {"type": "value", "axisLabel": {"color": "#475569"}},
                "series": [{
                    "type": "bar", "data": pct_vals,
                    "itemStyle": {"color": "#4ade80"},
                    "label": {"show": True, "position": "top", "fontSize": 11,
                              "formatter": "{c}%"},
                }],
            },
        })

    # ---- 图③:同比变化(取「同比%」字段,口径同上) ----
    yoy_pct_vals = [r.get("同比%") for r in records]
    if any(isinstance(v, (int, float)) for v in yoy_pct_vals):
        charts.append({
            "type": "bar",
            "title": "各指标同比变化(%)",
            "option": {
                "title": {"text": "各指标同比变化(%)", "textStyle": {"color": "#0e7490"}},
                "tooltip": {"trigger": "axis"},
                "grid": {"left": 60, "right": 30, "top": 60, "bottom": 80},
                "xAxis": {
                    "type": "category", "data": labels,
                    "axisLabel": {"rotate": 30, "interval": 0, "fontSize": 11,
                                  "color": "#475569"},
                },
                "yAxis": {"type": "value", "axisLabel": {"color": "#475569"}},
                "series": [{
                    "type": "bar", "data": yoy_pct_vals,
                    "itemStyle": {"color": "#fbbf24"},
                    "label": {"show": True, "position": "top", "fontSize": 11,
                              "formatter": "{c}%"},
                }],
            },
        })

    return charts


# ---------------------- 周期列(月份/季度/周)识别与排序 ----------------------
#
# 背景(2026-10 修复"年度报告只有文字"根因之三):
#   月度经营数据表的"月份"列值是 "1月"/"2月"/.../"12月"/"全年合计",
#   pandas 的 pd.to_datetime("1月") 会解析失败或解析成 2001-01,
#   于是 detect_column_types 把它判为 text,后果是:
#     - trend 字段全空 —— LLM 手里没有任何"走势"数据,写不出环比分析;
#     - by_group 的 groupby 结果是**字符串排序**:1月,10月,11月,12月,2月,
#       3月... 折线图 X 轴顺序完全错乱;
#     - "全年合计"被当成 13 个月份之一混进趋势。
#   这里新增"周期列"概念:把月份/季度/周文本解析成有序的周期序号,
#   让趋势计算按真实时间顺序走,并把汇总行排除在外。

# 月份中文/数字写法 → 1~12
_MONTH_PATTERNS = (
    re.compile(r"^(\d{1,2})\s*月?$"),
    re.compile(r"^第\s*(\d{1,2})\s*月"),
)
# 季度写法 → 1~4
_QUARTER_PATTERNS = (
    re.compile(r"^(\d{1,2})\s*季度?$"),
    re.compile(r"^第\s*([1-4一二三四])\s*季度?"),
)
_CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6,
           "七": 7, "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12}

# 列名判定"这一列是周期轴"的关键词
_PERIOD_COL_KEYWORDS = ("月份", "季度", "月", "季", "week", "周", "month", "quarter", "period", "日期", "时间")


def parse_period_value(value: Any) -> tuple[str, int] | None:
    """把单个单元格解析为 (周期类型, 序号)。失败返回 None。

    返回示例:
        "1月"   -> ("month", 1)
        "Q3"    -> ("quarter", 3)
        "第三季度" -> ("quarter", 3)
        "全年合计" -> None  (汇总标签,不是周期点)
    """
    s = _norm_label(value)
    if not s or is_total_label(s):
        return None
    # 季度(Q3 / Q1 / 三季度)
    m = re.match(r"^q([1-4])$", s)
    if m:
        return ("quarter", int(m.group(1)))
    m = re.match(r"^第?([一二三四])季度?$", s)
    if m:
        return ("quarter", _CN_NUM[m.group(1)])
    m = re.match(r"^第?([1-4一二三四])季度?$", s)
    if m:
        g = m.group(1)
        return ("quarter", _CN_NUM.get(g) or int(g))
    # 月份
    m = re.match(r"^第?(\d{1,2})月?$", s)
    if m:
        v = int(m.group(1))
        if 1 <= v <= 12:
            return ("month", v)
    return None


def is_period_column(series: pd.Series, col_name: Any = "") -> bool:
    """判断一列是否是"周期轴"(月份/季度/周)。

    判定要求**同时**满足:
        1) 列名含周期关键词(月份/季度/月/季/周/日期/time/month/period...);
        2) 非空值中 ≥60% 能解析成周期序号。
    只看列名会把"月度说明"这类文本误判;只看值会把"1月/2月"这种
    恰好是数字的普通数据误判,两者结合才可靠。
    """
    name = _norm_label(col_name)
    if not any(k in name for k in _PERIOD_COL_KEYWORDS):
        return False
    nn = series.dropna()
    if len(nn) == 0:
        return False
    ok = sum(1 for v in nn if parse_period_value(v) is not None)
    return ok / len(nn) >= 0.6


def period_sort_key(value: Any) -> tuple[int, int]:
    """给周期值一个排序键,让 "1月" < "2月" < ... < "10月" < "11月" < "12月"。

    解析失败(汇总行等)排到最前并标记 type=99,便于调用方按需剔除。
    """
    parsed = parse_period_value(value)
    if parsed is None:
        return (99, 0)
    kind, idx = parsed
    # month 用 100+idx 保证排在 quarter(1~4)之后但组内有序
    return (10 if kind == "month" else 5, idx)


# ---------------------- 章节明细数据表（2026-10 新增）====================
# 背景:
#     用户反馈"年度报告不应该只有文字,内容应该还是数据吧"。修完正文数字密度后,
#     报告正文已经有数字了,但**管理层仍无法核对**——正文里的 4680 万元
#     是怎么来的?各渠道加总对不对?这些问题需要一张可读的表格来回答。
#     图表看趋势、明细表看构成,两者配合才是完整的数据呈现。
#
#     本函数把 indicators 里的结构化数据转成
#     {title, columns, rows, note} 形态,前端 ReportView 直接渲染成 <table>,
#     导出 PDF/Word 时也能落成真表格。

# 单表最多展示多少行(超出部分前端/导出提示"还有 N 条未展开")
DETAIL_TABLE_MAX_ROWS = 30


def _clean_cell(v: Any) -> Any:
    """单元格值转 JSON/表格友好形态:NaN/NaT → None,numpy 标量 → Python 原生。"""
    try:
        if v is None or pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(v, "item"):
        try:
            return v.item()
        except (ValueError, AttributeError):
            pass
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%Y-%m-%d")
        except (ValueError, AttributeError):
            pass
    return v


def build_section_detail_tables(
    indicators: dict[str, Any],
    section_names: list[str] | None = None,
) -> dict[str, dict[str, Any]]:
    """按章节语义产出对应的明细数据表。

    返回: {章节名: {"title": str, "columns": [str], "rows": [[Any]],
                    "note": str, "total_rows": int}}
        没有匹配到数据源的章节不出现在返回值里(前端不渲染空表)。

    章节 → 数据源映射(与 analyzer._build_section_digest 的分流逻辑保持一致,
    保证"正文摘要"和"明细表"讲的是同一份数据,不会出现前后矛盾):
        指标/概况/成本/盈利章 → KPI 明细表
        趋势/增长章           → 周期走势表(各月数值 + 环比)
        结构/渠道/产品章       → 分组占比表(含占比%)
        风险章                → 未达标 KPI 表
        附录/其它章           → 数值列汇总表
    """
    out: dict[str, dict[str, Any]] = {}
    if not isinstance(indicators, dict):
        return out

    # 取 KPI(多 sheet 取第一条有记录的)
    kpi = indicators.get("kpi_matrix")
    if not (isinstance(kpi, dict) and kpi.get("records")):
        for v in (indicators.get("kpi_matrix_by_sheet") or {}).values():
            if isinstance(v, dict) and v.get("records"):
                kpi = v
                break
    records = (kpi or {}).get("records") or [] if isinstance(kpi, dict) else []

    def kpi_table(recs: list[dict], title: str, note: str = "") -> dict[str, Any] | None:
        """KPI records → 明细表。比率类指标按百分数展示,避免"毛利率 0.338"。"""
        if not recs:
            return None
        cols = ["指标", "单位", "目标", "实际", "差值", "达成率", "同比", "去年同期"]
        rows: list[list[Any]] = []
        for rec in recs[:8]:
            unit = rec.get("单位")
            ratio_like = (unit == "%")

            def disp(key: str, is_ratio: bool = False) -> Any:
                v = rec.get(key)
                if not isinstance(v, (int, float)) or isinstance(v, bool):
                    return None
                return round(float(v) * 100, 2) if is_ratio else float(v)

            tgt, act = rec.get("目标"), rec.get("实际")
            gap = (act - tgt) if (
                isinstance(act, (int, float)) and isinstance(tgt, (int, float))
            ) else None
            rows.append([
                rec.get("指标"),
                unit or "—",
                disp("目标", ratio_like),
                disp("实际", ratio_like),
                round(gap * 100, 2) if (gap is not None and ratio_like) else gap,
                rec.get("达成率%"),
                rec.get("同比%"),
                disp("去年同期", ratio_like),
            ])
        return {
            "title": title,
            "columns": cols,
            "rows": rows[:DETAIL_TABLE_MAX_ROWS],
            "total_rows": len(rows),
            "note": note or "比率类指标(单位 %)按百分数展示;差值=实际-目标,负数表示未达标。",
        }

    # 章节主题 → 相关 KPI 指标关键词。
    # 与 analyzer._THEME_KPI_HINTS 同源:让"客户与市场维度"章只展示
    # 客单价/新客户数/留存率,而不是把 7 个指标全表贴一遍。
    _THEME_KPI_HINTS: list[tuple[tuple[str, ...], tuple[str, ...]]] = [
        (("客户", "用户", "市场", "客群"), ("客户", "客单", "留存", "复购", "新客")),
        (("成本", "毛利", "利润", "盈利", "费用"), ("毛利", "利润", "成本", "费用", "净利")),
        (("价格", "定价", "单价"), ("客单", "单价", "价格", "毛利")),
    ]

    def scoped_records(sec_name: str) -> list[dict]:
        """按章节主题筛出相关 KPI 记录;筛不到返回空列表。"""
        for themes, kpi_keys in _THEME_KPI_HINTS:
            if any(t in sec_name for t in themes):
                got = [r for r in records
                       if any(k in str(r.get("指标") or "") for k in kpi_keys)]
                if got:
                    return got
        return []

    for sec in (section_names or []):
        name = str(sec)
        is_metric = any(k in name for k in (
            "指标", "KPI", "业绩", "达成", "总体", "概况", "经营", "盈利", "成本", "毛利", "价格"))
        is_trend = any(k in name for k in ("趋势", "走势", "增长", "月度", "季度", "规模"))
        is_structure = any(k in name for k in (
            "结构", "占比", "渠道", "产品", "区域", "板块", "品类"))
        is_risk = any(k in name for k in (
            "风险", "问题", "归因", "异常", "不足", "挑战"))
        # 客户/用户维度:虽然含"市场"会被判成结构章,但其数据源应该是
        # KPI 里的客单价/新客户数/留存率,而非渠道销售额占比。
        is_customer = any(k in name for k in ("客户", "用户", "客群"))
        if is_customer:
            is_structure = False
            is_metric = True

        table: dict[str, Any] | None = None

        # ---- 风险章:只看未达标指标 ----
        if records and is_risk:
            missed = [r for r in records
                      if isinstance(r.get("达成率%"), (int, float)) and r["达成率%"] < 100]
            table = kpi_table(missed, "未达标指标明细",
                              note="仅列出达成率 < 100% 的指标;差值=实际-目标,负数表示未达标。")

        # ---- 指标/概况章:KPI 明细(主题明确时只取相关指标) ----
        if table is None and records and is_metric and not is_structure:
            scoped = scoped_records(name)
            if scoped:
                table = kpi_table(scoped, f"{'核心'} KPI 明细(目标 / 实际 / 达成率 / 同比)")
            else:
                table = kpi_table(records, "核心 KPI 明细(目标 / 实际 / 达成率 / 同比)")

        # ---- 周期走势表 ----
        if table is None and is_trend:
            for key, tr in (indicators.get("trend") or {}).items():
                if not isinstance(tr, dict) or not tr.get("x"):
                    continue
                xs, ys = tr.get("x") or [], tr.get("y") or []
                mom = tr.get("环比变化率%") or []
                peak = tr.get("峰值") or {}
                trough = tr.get("谷值") or {}
                rows = []
                for i in range(len(xs)):
                    rows.append([
                        xs[i],
                        ys[i] if i < len(ys) else None,
                        mom[i] if i < len(mom) else None,
                    ])
                total = tr.get("总计")
                note = (
                    f"合计 {total};" if isinstance(total, (int, float)) else ""
                ) + (
                    f"峰值 {peak.get('周期')}({peak.get('值')});"
                    f"谷值 {trough.get('周期')}({trough.get('值')})。"
                    if peak or trough else ""
                )
                table = {
                    "title": f"{key} · 周期走势明细",
                    "columns": ["周期", "数值", "环比变化率(%)"],
                    "rows": rows[:DETAIL_TABLE_MAX_ROWS],
                    "total_rows": len(rows),
                    "note": (note or "环比 = 与上一周期比较。").strip(),
                }
                break

        # ---- 分组占比表 ----
        if table is None and is_structure:
            for key, items in (indicators.get("by_group") or {}).items():
                if not isinstance(items, list) or not items:
                    continue
                vals = [float(it.get("value") or 0) for it in items if isinstance(it, dict)]
                total = sum(vals)
                rows = []
                for it in items:
                    if not isinstance(it, dict):
                        continue
                    v = float(it.get("value") or 0)
                    rows.append([
                        it.get("name"),
                        v,
                        round(v / total * 100, 2) if total else None,
                    ])
                table = {
                    "title": f"{key} · 结构占比明细",
                    "columns": ["分组", "数值", "占比(%)"],
                    "rows": rows[:DETAIL_TABLE_MAX_ROWS],
                    "total_rows": len(rows),
                    "note": "已剔除「合计/小计」汇总行,占比基于明细项重新计算。",
                }
                break

        # ---- 兜底:数值列汇总表 ----
        # 行动建议 / 摘要 / 下一步 这类章节没有专属数据源,
        # 但"数值列汇总"对它们几乎没有解释价值(字段名带 sheet 前缀、
        # 粒度是全表统计),所以优先回落 KPI 概览,而不是硬塞统计表。
        if table is None and records:
            table = kpi_table(records, "核心 KPI 明细(目标 / 实际 / 达成率 / 同比)")

        if table is None:
            ns = indicators.get("numeric_summary") or {}
            rows = [
                [col, agg.get("sum"), agg.get("mean"), agg.get("max"),
                 agg.get("min"), agg.get("median")]
                for col, agg in ns.items()
                if isinstance(agg, dict)
            ]
            if rows:
                table = {
                    "title": "数值列汇总明细",
                    "columns": ["字段", "合计", "均值", "最大值", "最小值", "中位数"],
                    "rows": rows[:DETAIL_TABLE_MAX_ROWS],
                    "total_rows": len(rows),
                    "note": "各数值列的描述性统计。",
                }

        if table:
            out[sec] = table

    return out


def _coerce_numeric(series: pd.Series) -> tuple[pd.Series, int]:
    """尝试把"看起来像数字"的文本列转为数值。

    常见被 Excel/CSV 当文本的数字:
        - 百分号: "12.5%"
        - 千分位: "1,234.56"
        - 货币符: "¥1,234.5" / "$123"
        - 单位: "10万" / "5k"
        - 中文前缀: "约 10"、"合计 100"

    返回:
        (转换后的 Series, 成功转换的非空个数)
        调用方根据"成功个数 / 非空总数 ≥ 阈值"判断是否真的算 number。
    """
    if series.dtype != object:
        return series, 0

    s = series.astype(str).str.strip()
    # 1. 去常见前缀:¥/$/€/约/共/大概/合计/总额
    s = s.str.replace(r"^[¥$€￡]|^(?:约|共|大概|合计|总额|约计)\s*", "", regex=True)
    # 2. 去常见后缀:万/亿/%/个/件/元/美元
    s = s.str.replace(r"(?:万|亿|%|个百分点|个|件|元|美元)$", "", regex=True)
    # 3. 千分位逗号 → 移除
    s = s.str.replace(",", "", regex=False)
    # 4. 英文 k/K 单位(只去后缀,避免破坏小数点)
    s = s.str.replace(r"(?<=\d)\s*[kK]\s*$", "", regex=True)

    converted = pd.to_numeric(s, errors="coerce")
    success = int(converted.notna().sum())
    return converted, success


# ---- ID/编号类列识别(防止被误判为 number 进 groupby) ----
#
# 背景:Excel 导出/业务系统里"单据编号"、"客户编号"、"订单号" 等列
# 通常是 12 位以上的数字(看起来是数值),但其实是"每个值唯一"的 ID。
# 如果把它当作 number,后续 compute_indicators 会按它做 groupby 求和,
# Y 轴就会变成 10^25 这种数量级的伪"数值"——典型的图表崩坏场景。
#
# 识别规则(命中任一即视为 ID-like text):
#   1. 列名包含常见 ID 关键词:单据编号/单号/订单号/编号/序号/ID/id/NO/No/编号
#   2. 全列去重数 / 非空数 ≥ 90%(高度唯一)
#   3. 全部由数字/短横线组成、长度 ≥ 8 位(看起来像规范编号)
_ID_HINT_KEYWORDS = (
    "编号", "单号", "订单号", "流水号", "序号", "编码", "id", "ID", "No", "NO",
    "uuid", "UUID", "key", "code", "Code",
)


def _looks_like_id_column(series: pd.Series, col_name: str) -> bool:
    """判断一列是否像"ID/编号"列(而不是可聚合的数值列)。

    入参:
        series:    pd.Series, dtype 不限(可能 object/int64)
        col_name:  原列名(字符串),用于关键词匹配

    返回:
        True  → 视为 ID 列,即便 coerce 后变成 number 也按 text 处理

    判定规则(命中任一即视为 ID):
        1) 列名包含 ID 类关键词(编号/单号/订单号/ID/No 等)
        2) 全列几乎不重复(唯一率 ≥ 0.9)且:
           a) 样本全是数字字符(可含 -/_)且平均长度 ≥ 8 位
           b) 样本全是数字字符且最大长度 ≥ 10 位(典型长数字 ID)
           c) 唯一率 ≥ 0.98 且样本字符长度差很大(混合长度,如 UUID/编号段)
        判定边界:
        - 必须是"已确认是数值"才会走到这条路径;
        - 短纯数值列(如 [100, 250, 89, 310])虽然唯一率=1,平均长度也才 3-6,
          不满足 ≥ 8 / ≥ 10,不会被误判为 ID
    """
    name_l = (str(col_name) or "").lower()
    if any(k.lower() in name_l for k in _ID_HINT_KEYWORDS):
        return True

    non_null = series.dropna()
    n = int(non_null.shape[0])
    if n < 3:  # 样本太少不做启发,放行
        return False

    sample = non_null.astype(str).str.strip().head(50)
    unique = int(sample.nunique())
    ratio = unique / len(sample)

    # 全数字字符(允许 -/_)
    sample_no_neg = sample.str.replace("-", "", regex=False)
    all_digit_only = sample.str.match(r"^[\d\-_A-Za-z]+$").all()
    # 是否主要为数字(去掉负号 / 横线后整串是数字)
    mostly_digit = sample_no_neg.str.match(r"^\d+$").mean() >= 0.9
    avg_len = float(sample.str.len().mean()) if len(sample) else 0
    max_len = int(sample.str.len().max()) if len(sample) else 0

    if ratio >= 0.9:
        # 规则 a:全数字字符 + 平均长度 ≥ 8 → ID
        if all_digit_only and mostly_digit and avg_len >= 8:
            return True
        # 规则 b:全数字 + 最大长度 ≥ 10 → ID(典型手机号/订单号 12-18 位)
        if mostly_digit and max_len >= 10:
            return True
        # 规则 c:唯一率 ≥ 0.99(几乎都不同) + 全字符型(非纯数字)→ 文本 ID,不再做数值聚合
        if ratio >= 0.99 and not mostly_digit and all_digit_only:
            return True

    return False


def detect_column_types(df: pd.DataFrame) -> dict[str, str]:
    """推断每列类型：date / number / text。

    改版要点:
        - 原本只走"Pandas dtype 推断 + 日期降级",遇到 Excel/CSV 里被存成
          字符串的数字(比如 "12.5%"、"1,234"、"¥100")会被错判为 text,
          进而导致 compute_indicators 的 numeric_cols 为空、AI 报告"没数据"。
        - 现在对每列先尝试"文本转数值"兜底:能成功 ≥ 70% 就视为 number。
        - 仍按 date → number → text 的优先级匹配。
        - **新增**:在识别 number 之前先做 ID/编号类识别,单据号/订单号这类列
          不再被误判为可聚合数值,避免后续图表出现 1e+25 的"伪数值"。
    """
    result: dict[str, str] = {}
    for col in df.columns:
        s = df[col]
        col_key = str(col)

        if pd.api.types.is_datetime64_any_dtype(s):
            result[col_key] = "date"
            continue
        if pd.api.types.is_numeric_dtype(s):
            # 已经 dtype 推断成 number 的列,也再过一遍 ID 检测,
            # 命中就降级为 text,防止单据号这种"数字型 ID"参与 groupby 求和
            if _looks_like_id_column(s, col_key):
                result[col_key] = "text"
            else:
                result[col_key] = "number"
            continue

        # 尝试解析为日期
        # 优先用 ``format="mixed"``（pandas 2.x 的快速路径）让 pandas 整体解析，
        # 失败再回退到 ``dateutil`` 逐元素解析（同时静默其 UserWarning）。
        # 解析成功率 ≥ 80% 才认定为日期列，避免文本误判。
        sample = s.dropna().head(20)
        if sample.shape[0] > 0:
            try:
                parsed = pd.to_datetime(sample, errors="raise", format="mixed")
                result[col_key] = "date"
                continue
            except (ValueError, TypeError):
                # 回退到 dateutil 逐元素解析；为避免刷屏的 UserWarning，
                # 用 warnings 上下文局部压制一次。
                import warnings as _warnings
                with _warnings.catch_warnings():
                    _warnings.simplefilter("ignore", UserWarning)
                    try:
                        parsed = pd.to_datetime(sample, errors="coerce")
                    except Exception:
                        parsed = None
                if parsed is not None and parsed.notna().mean() >= 0.8:
                    result[col_key] = "date"
                    continue

        # 文本列:尝试兜底转数值,≥ 70% 成功视为 number
        non_empty = int(s.dropna().shape[0])
        if non_empty > 0:
            _, success = _coerce_numeric(s)
            if success / non_empty >= 0.7:
                # 数值列但又像 ID → 降级 text,避免 groupby 出伪聚合
                if _looks_like_id_column(s, col_key):
                    result[col_key] = "text"
                else:
                    result[col_key] = "number"
                continue

        result[col_key] = "text"
    return result


def column_profiles(df: pd.DataFrame) -> list[dict[str, Any]]:
    """返回每列的结构画像,供「智能派生列」展示与 AI 推荐使用。

    - type 复用 detect_column_types(df),与工作台其它模块的类型判定保持一致
      (Excel 里存成文本的数字列也会被识别为 number);
    - non_null 只统计非空值,unique_count 只对非空值去重(不把 NaN 当成取值);
    - 列顺序严格保持 df.columns 的原顺序。

    返回: [{name, type, non_null, unique_count}, ...]
    """
    types = detect_column_types(df)
    profiles: list[dict[str, Any]] = []
    for col in df.columns:
        name = str(col)
        series = df[col]
        profiles.append({
            "name": name,
            "type": types.get(name, "text"),
            "non_null": int(series.notna().sum()),
            "unique_count": int(series.dropna().nunique()),
        })
    return profiles


def basic_quality(df: pd.DataFrame) -> dict[str, Any]:
    """基础质量指标。"""
    return {
        "row_count": int(len(df)),
        "column_count": int(len(df.columns)),
        "missing_total": int(df.isna().sum().sum()),
        "missing_by_column": {str(c): int(v) for c, v in df.isna().sum().items()},
        "duplicate_rows": int(df.duplicated().sum()),
    }


# ---------------------- 清洗算子 ----------------------

def clean(df: pd.DataFrame, rules: dict[str, bool]) -> pd.DataFrame:
    """按规则清洗：dropna / drop_duplicates / unify_date / mark_anomaly。"""
    out = df.copy()

    if rules.get("drop_empty_rows"):
        out = out.dropna(how="all")

    if rules.get("drop_duplicates"):
        out = out.drop_duplicates()

    # ——【新增于 workbench-cleansing-missing-data】缺失值填充 ——
    # 位置:drop_duplicates 之后、unify_date 之前
    # 理由:异常标记必须看到填充后的值,否则全 NaN 列 IQR 无法计算
    if rules.get("fill_missing"):
        # 缺省 policy 视为 {strategy: mean, scope: numeric}
        policy = rules.get("fill_missing_policy") or {"strategy": "mean", "scope": "numeric"}
        out = fill_missing(out, policy)

    if rules.get("unify_date"):
        for col in out.columns:
            s = out[col]
            try:
                # —— 守卫 1:数值列直接跳过 ——
                # pd.to_datetime 对"裸数字"的语义是"距 1970-01-01 的纳秒数":
                #     to_datetime(1791) → 1970-01-01 00:00:00.000001791
                # 再 strftime 就变成 "1970-01-01"。且这种转换永不报错、100% 成功,
                # 下面的"成功率 ≥50%"阈值根本拦不住——旧实现正是因此把
                # 销售数量/销售额/单价等所有数值列整列毁成 1970-01-01。
                if pd.api.types.is_numeric_dtype(s):
                    continue

                # —— 守卫 2:object 列里过半是"纯数字文本"也不是日期 ——
                # "120000"、"199" 这类文本会被 dateutil 当作"公元 199 年"等
                # 荒谬年份解析成功,同样会污染整列;先看数字占比,过高直接跳过。
                non_null = s.dropna()
                if non_null.empty:
                    continue
                as_str = non_null.astype(str).str.strip()
                if as_str.str.fullmatch(r"-?\d+(\.\d+)?").mean() >= 0.5:
                    continue

                # —— 守卫 3:真实日期列才转换 ——
                # 要求 (a) 非空值中 ≥50% 能解析成日期(沿用旧阈值);
                #       (b) 解析出的年份 ≥90% 落在 1975~2100 的合理区间,
                #           拦下"公元 199 年"这类通过解析但明显不是业务日期的值。
                converted = pd.to_datetime(s, errors="coerce")
                ok = converted.dropna()
                if len(ok) < max(1, int(len(out) * 0.5)):
                    continue
                if ok.dt.year.between(1975, 2100).mean() < 0.9:
                    continue
                out[col] = converted.dt.strftime("%Y-%m-%d")
            except Exception:
                continue

    if rules.get("mark_anomaly"):
        for col in out.select_dtypes(include=[np.number]).columns:
            out[f"{col}_is_outlier"] = _mark_iqr(out[col])

    return out


def _mark_iqr(series: pd.Series) -> pd.Series:
    """IQR 法标记异常值。"""
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return pd.Series([False] * len(series), index=series.index)
    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr
    return (series < lower) | (series > upper)


# ---------------------- 缺失值填充算子 ----------------------
# 新增于 workbench-cleansing-missing-data:为数据清洗弹窗提供第 5 项规则
# 的底层实现,6 种 strategy(mean/median/mode/zero/ffill/drop)
# × 2 种 scope(numeric/all),文本列仅允许 mode/ffill/drop,日期列仅允许
# ffill/drop,对文本列调用 mean/median/zero 直接 ValueError。

# strategy 白名单(小写字符串),用于 policy 校验
_STRATEGIES = ("mean", "median", "mode", "zero", "ffill", "drop")
# scope 白名单:numeric 默认只处理数值列;all 处理所有列
_SCOPES = ("numeric", "all")
# 文本/日期列允许的 strategy 子集,避免把文本均值算出来
# 这里合并了文本(mode/ffill/drop)与日期(ffill/drop)的允许集,
# 因为日期列在数值列前提下已经被 scope=numeric 过滤掉,
# 当 scope=all 时允许 mode/ffill/drop 已足够覆盖两类非数值列
_NON_NUMERIC_ALLOWED = {"mode", "ffill", "drop"}


def fill_missing(df: pd.DataFrame, policy: dict[str, str]) -> pd.DataFrame:
    """按策略填充缺失值(单元格级,与 drop_empty_rows 的"整行全空"不同)。

    Args:
        df:  原始 DataFrame(**不修改**,内部 .copy() 后操作)
        policy: {strategy, scope}
            - strategy ∈ {"mean", "median", "mode", "zero", "ffill", "drop"}
              mean/median 仅对数值列有效;zero 仅对数值列有效;
              mode/ffill/drop 对所有列有效
            - scope   ∈ {"numeric", "all"}
              numeric 默认,只处理数值列;all 处理所有列(含文本/日期)

    Returns:
        新的 DataFrame;空 df 直接返回

    Raises:
        ValueError: strategy/scope 非法,或对文本列调用 mean/median/zero
    """
    # 1. 解析并校验 policy:策略与作用域必须在白名单内,否则直接拒绝
    strategy = (policy or {}).get("strategy", "mean")
    scope = (policy or {}).get("scope", "numeric")
    if strategy not in _STRATEGIES:
        raise ValueError(
            f"不支持的缺失值填充策略: {strategy},可选 mean/median/mode/zero/ffill/drop"
        )
    if scope not in _SCOPES:
        raise ValueError(f"不支持的作用范围: {scope},可选 numeric/all")

    # 空 DataFrame 直接返回(避免空 series 上 .mean() 返回 NaN 后 fillna 出问题)
    if len(df) == 0 or len(df.columns) == 0:
        return df.copy()

    out = df.copy()

    # 2. 确定目标列集合:scope=numeric 时只盯数值列,scope=all 时覆盖所有列
    if scope == "numeric":
        target_cols = [str(c) for c in df.select_dtypes(include="number").columns]
    else:
        target_cols = [str(c) for c in df.columns]

    # 3. 校验:scope=all + 禁用 strategy 时,直接拒绝(给出具体 strategy 名称)
    # 例:文本列调 mean 会算出"字符串平均值",在数值语义上无意义;
    # 这里前置校验避免在循环里每列各抛一次。
    if scope == "all" and strategy not in _NON_NUMERIC_ALLOWED:
        label = {"mean": "均值", "median": "中位数", "zero": "置零"}.get(strategy, strategy)
        raise ValueError(
            f"文本列不支持{label}填充,请选择 mode/ffill/drop"
        )

    # 4. 按 strategy 执行各列填充
    for col in target_cols:
        series = out[col]
        # 无缺失值直接跳过(节省开销;常见:大部分列其实没 NaN)
        if not series.isna().any():
            continue

        if strategy == "mean":
            # 全 NaN 列 → mean() 返回 NaN → fillna(NaN) 等价不填,跳过避免虚假填充
            if series.dropna().empty:
                continue
            fill_value = round(float(series.mean()), 2)
            # mean() 对 NaN 内部不报错,但需要 isinstance 校验以防类型异常
            if pd.isna(fill_value):
                continue
            out[col] = series.fillna(fill_value)

        elif strategy == "median":
            if series.dropna().empty:
                continue
            fill_value = round(float(series.median()), 2)
            if pd.isna(fill_value):
                continue
            out[col] = series.fillna(fill_value)

        elif strategy == "mode":
            # mode() 返回 Series(可能多个众数),取首个非空众数;
            # 文本列众数也适用(取出现频率最高的非空字符串)
            mode_series = series.mode(dropna=True)
            if mode_series.empty:
                continue
            out[col] = series.fillna(mode_series.iloc[0])

        elif strategy == "zero":
            # 仅数值列;scope=numeric 时已天然过滤;scope=all 时前置校验已拒绝非数值列
            out[col] = series.fillna(0)

        elif strategy == "ffill":
            # 前向填充;首行 NaN 保持(不能凭空造);ffill 后若还有尾随 NaN 也不处理
            # 这里**不**用 bfill 兜底,严格按 spec "首行 NaN 保持不动"
            out[col] = series.fillna(method="ffill")

        elif strategy == "drop":
            # 仅标记,不立即删行(避免循环里 df 变化导致索引错乱)
            # 收集到待删索引,循环结束后统一 dropna
            pass  # 在 step 5 统一处理

    # 5. strategy="drop" 时,在循环外统一 dropna(subset=target_cols)
    #    注意:这里**只删"目标列中存在 NaN"的行**,不影响其它列;
    #    与 clean() 流水线里的 drop_empty_rows(dropna how="all")不重复
    if strategy == "drop" and target_cols:
        out = out.dropna(subset=target_cols).reset_index(drop=True)

    return out


# ---------------------- 派生列算子 ----------------------
# 【clean-ai-derived-columns-pivot 新增】
# 需求:数据清洗弹窗允许用户/AI 推荐派生列(如 销售额 = 销量 × 单价)。
# 安全边界(重要):
#   - 绝不使用 Python 的 eval / exec —— 用户输入直接进 Python 解释器等于
#     把服务器交给对方;
#   - 只允许「反引号包裹列名 + 数字 + + - * / ( ) . % + 空白」,
#     任何其它字符(函数调用、属性访问、比较运算等)一律拒绝;
#   - 真正的计算交给 DataFrame.eval:它只解析列引用与四则运算表达式,
#     不会执行任意代码。

_DERIVED_NAME_MAX_LEN = 30
# 反引号包裹的列名片段,例如 `销量`
_DERIVED_COL_REF_RE = re.compile(r"`([^`]+)`")
# 去掉列名引用后,公式里只允许:数字 + 四则运算 + 括号/小数点/百分号 + 空白
_DERIVED_ALLOWED_CHARS_RE = re.compile(r"^[0-9+\-*/().%\s]*$")


def validate_derived_specs(
    column_names: Any, specs: Any
) -> list[tuple[str, str]]:
    """静态校验派生列配置,返回 [(新列名, 公式), ...]。

    只做语法/安全校验,不执行计算。抽成公共函数的目的:
        - apply_derived_columns 执行前校验;
        - AI 推荐结果过滤时复用**完全同一套**规则,避免两边口径漂移。

    校验项:
        1. 名称 strip 后非空,且不超过 30 个字符;
        2. 同一批内名称不重复;
        3. 不与入参现有列同名;
        4. 公式不能为空;
        5. 反引号内引用的列必须真实存在;
        6. 去掉反引号片段后,剩余字符只能来自 0-9 + - * / ( ) . % 和空白。

    Raises:
        ValueError: 任一校验失败,消息为中文,由 API 层统一转 400。
    """
    if specs is None:
        return []
    if not isinstance(specs, list):
        raise ValueError("派生列配置必须是数组")

    existing = {str(c) for c in column_names}
    seen: set[str] = set()
    result: list[tuple[str, str]] = []

    for index, spec in enumerate(specs, start=1):
        if not isinstance(spec, dict):
            raise ValueError(f"第 {index} 条派生列配置不完整")
        name = str(spec.get("name") or "").strip()
        expression = str(spec.get("expression") or "").strip()

        if not name:
            raise ValueError(f"派生列名称不能为空:{index}")
        if len(name) > _DERIVED_NAME_MAX_LEN:
            raise ValueError(
                f"派生列名称不能超过 {_DERIVED_NAME_MAX_LEN} 个字符:{name}"
            )
        if name in seen:
            raise ValueError(f"派生列名称重复:{name}")
        if name in existing:
            raise ValueError(f"派生列名称与现有列重复:{name}")
        if not expression:
            raise ValueError(f"派生列公式不能为空:{name}")

        # 反引号内的列名必须存在;EXISTS 集合只含入参 df 的列,
        # 因此天然禁止「引用同批新列」这种依赖顺序的写法。
        for ref in _DERIVED_COL_REF_RE.findall(expression):
            if ref not in existing:
                raise ValueError(f"公式引用了不存在的列:{ref}")

        remainder = _DERIVED_COL_REF_RE.sub("", expression)
        if not _DERIVED_ALLOWED_CHARS_RE.fullmatch(remainder):
            raise ValueError(
                f"公式包含非法字符,列名需用反引号包裹:{expression}"
            )

        seen.add(name)
        result.append((name, expression))

    return result


def apply_derived_columns(
    df: pd.DataFrame, specs: list[dict[str, Any]]
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """按配置安全地追加派生列,返回 (新 DataFrame, 执行报告)。

    Args:
        df:    清洗后的 DataFrame(不修改入参)
        specs: [{name, expression}, ...];空列表/None 时返回副本与空报告

    Returns:
        (new_df, {"created": [新列名...], "warnings": [中文提示...]})

    执行规则:
        - 所有公式都基于**入参 df 的同一份副本**计算,算完再统一写回,
          保证同批派生列之间不能互相引用(避免 A 依赖同批 B 的隐式顺序);
        - 结果统一 pd.to_numeric(errors="coerce");
        - 有效值占比为 0 → ValueError(公式语法通过但算不出任何数字);
        - inf / -inf(除零、溢出)统一置为 NaN 并写入 warning;
        - 整列全 NaN 仍然保留该列,并写 warning 提示用户检查公式。
    """
    report: dict[str, Any] = {"created": [], "warnings": []}
    if not specs:
        return df.copy(), report

    parsed = validate_derived_specs([str(c) for c in df.columns], specs)
    out = df.copy()
    # 公式计算基准:与入参同内容的独立副本,不随 out 的新增列变化
    base = df.copy()

    for name, expression in parsed:
        try:
            result = base.eval(expression)
        except Exception as exc:
            # DataFrame.eval 的语法错误 / 列类型不支持等,统一转中文 ValueError
            raise ValueError(f"公式计算失败:{expression}({exc})") from exc

        # eval 正常返回 Series;若公式只含字面量(如 "1 + 2")会返回标量,
        # 极端情况下也可能返回多列 DataFrame,这里统一收敛为单列 Series。
        if isinstance(result, pd.DataFrame):
            if result.shape[1] != 1:
                raise ValueError(f"公式必须返回单列结果:{expression}")
            result = result.iloc[:, 0]
        if not isinstance(result, pd.Series):
            result = pd.Series(result, index=base.index)

        numeric = pd.to_numeric(result, errors="coerce")

        # 有效值占比为 0:说明公式没有产出任何可用数字(如整列文本相减)
        if len(numeric) > 0 and numeric.notna().mean() == 0:
            raise ValueError(f"公式未产生数值结果:{expression}")

        # 除零 / 溢出会产生 inf,统一转成空值;保留计数用于提示
        inf_count = int(((numeric == np.inf) | (numeric == -np.inf)).sum())
        series = numeric.replace([np.inf, -np.inf], np.nan)

        out[name] = series
        report["created"].append(name)
        if inf_count > 0:
            report["warnings"].append(
                f"派生列 {name} 有 {inf_count} 个值因除零或溢出被置为空值"
            )
        if len(series) > 0 and series.isna().all():
            report["warnings"].append(f"派生列 {name} 全部为空值,请检查公式")

    return out, report


# ---------------------- 派生列 → Excel 活公式注入 ----------------------
# 【clean-derived-excel-formula 新增】
# 背景:派生列默认以「静态数值」落盘,用户在 Excel/WPS 公式栏里看到的只是
# 一个数字,无法联动重算。这里的做法是 ExcelWriter 照常写值(保留缓存值,
# 本工具预览/二次分析不受影响),随后对 xlsx 的 sheet XML 做一次精准后处理,
# 在派生列单元格的 <v> 前插入 <f>,得到「活公式 + 缓存值」双写形态:
#   - Excel/WPS 打开:公式栏显示 =B2*C2,改源数据会联动重算;
#   - pandas 再读取(data_only):读到的仍是缓存值,不会变成 NaN。
#
# 安全前提(由 validate_derived_specs 保证):表达式只含 + - * / % ( ) 与
# 反引号列名,且不能引用同批新列,因此可以 1:1 翻译成 Excel 公式。

# 表达式 token 切分:反引号列名 / 数字(含小数) / 运算符与括号
_DERIVED_EXPR_TOKEN_RE = re.compile(r"\s*(`[^`]+`|\d+(?:\.\d+)?|[+\-*/%()])")


def _derived_expr_to_excel(expression: str, col_letter, row: int) -> str:
    """把派生列表达式翻译成指定行的 Excel 公式(不含开头的 =)。

    Args:
        expression: 派生列公式,如 `` `销量` * `单价` ``(已过 validate_derived_specs)
        col_letter: 回调,列名 → Excel 列字母(如 "B")
        row:        目标行的 Excel 行号(表头为 1,数据从 2 起)

    语法与白名单一致:仅 + - * / % 与括号;二元 % 在 Excel 中没有运算符
    形态,转写为 MOD(左, 右)。解析失败抛 ValueError,由调用方决定回退。
    """
    tokens = _DERIVED_EXPR_TOKEN_RE.findall(expression)
    if not tokens:
        raise ValueError(f"公式为空:{expression}")
    pos = 0

    def _peek() -> str | None:
        return tokens[pos] if pos < len(tokens) else None

    def _parse_expr() -> str:
        # expr := term (('+'|'-') term)*
        nonlocal pos
        left = _parse_term()
        while _peek() in ("+", "-"):
            op = tokens[pos]
            pos += 1
            left = f"{left}{op}{_parse_term()}"
        return left

    def _parse_term() -> str:
        # term := factor (('*'|'/'|'%') factor)*;二元 % → Excel MOD()
        nonlocal pos
        left = _parse_factor()
        while _peek() in ("*", "/", "%"):
            op = tokens[pos]
            pos += 1
            right = _parse_factor()
            left = f"MOD({left},{right})" if op == "%" else f"{left}{op}{right}"
        return left

    def _parse_factor() -> str:
        # factor := 数字 | `列名` | '(' expr ')' | 一元 +/- factor
        nonlocal pos
        tok = _peek()
        if tok is None:
            raise ValueError(f"公式不完整:{expression}")
        if tok in ("+", "-"):
            pos += 1
            return f"{tok}{_parse_factor()}"
        if tok == "(":
            pos += 1
            inner = _parse_expr()
            if _peek() != ")":
                raise ValueError(f"公式括号不配对:{expression}")
            pos += 1
            return f"({inner})"
        if tok.startswith("`"):
            pos += 1
            return f"{col_letter(tok[1:-1])}{row}"
        pos += 1
        return tok  # 数字字面量,Excel 与 Python 写法一致

    formula = _parse_expr()
    if pos != len(tokens):
        raise ValueError(f"公式存在无法解析的片段:{expression}")
    return formula


def inject_derived_formulas(
    path,
    written_sheet_names: dict[str, str],
    cleaned_sheets: dict[str, "pd.DataFrame"],
    derived_specs: dict[str, list[dict[str, str]]],
) -> int:
    """把导出 xlsx 中派生列的静态值升级成「活公式 + 缓存值」,返回注入的单元格数。

    Args:
        path:                已写出的 xlsx 路径(原地改写)
        written_sheet_names: 原 sheet 名 → 文件内实际 sheet 名(单表分支为 Sheet1)
        cleaned_sheets:      原 sheet 名 → 已落盘的 DataFrame(提供列序与空值掩码)
        derived_specs:       原 sheet 名 → [{name, expression}](仅已成功创建的列)

    行级掩码:仅对「计算结果非空」的行注入公式。空白行没有 <c> 元素,且
    Excel 中空单元格参与四则运算得 0 而非 pandas 的 NaN,跳过可保持
    重算前后语义一致。

    任何解析/结构异常都抛给调用方(tools.py 捕获后记 warning,文件保持静态值)。
    """
    import zipfile

    from openpyxl.utils import get_column_letter

    # 1) 逐 sheet 计算 {单元格 ref: Excel 公式} 注入计划
    sheet_plan: dict[str, dict[str, str]] = {}
    for orig_name, specs in (derived_specs or {}).items():
        df = cleaned_sheets.get(orig_name)
        if df is None or not specs:
            continue
        col_pos = {str(c): idx for idx, c in enumerate(df.columns)}

        def _letter(col_name: str, _pos=col_pos) -> str:
            if col_name not in _pos:
                raise ValueError(f"公式引用了不存在的列:{col_name}")
            return get_column_letter(_pos[col_name] + 1)

        file_sheet = written_sheet_names.get(orig_name, orig_name)
        plan = sheet_plan.setdefault(file_sheet, {})
        for spec in specs:
            name = str(spec.get("name") or "")
            expression = str(spec.get("expression") or "")
            if name not in col_pos or not expression:
                continue
            target_letter = get_column_letter(col_pos[name] + 1)
            series = pd.to_numeric(df[name], errors="coerce")
            for i, val in enumerate(series.tolist()):
                if pd.isna(val):
                    continue  # 空结果行保持空白,不注入公式(见 docstring)
                row_no = i + 2  # 表头占第 1 行
                plan[f"{target_letter}{row_no}"] = _derived_expr_to_excel(
                    expression, _letter, row_no
                )

    if not any(sheet_plan.values()):
        return 0

    # 2) sheet 名 → xl/worksheets/sheetN.xml 映射(workbook.xml + rels)
    with zipfile.ZipFile(path, "r") as zin:
        entries = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    workbook_xml = entries["xl/workbook.xml"].decode("utf-8")
    rels_xml = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    # 注意:openpyxl 写出的 Relationship 属性顺序是 Target 在前 Id 在后,
    # 不能用「Id 在前」的固定顺序正则,这里按标签逐个提取属性,与顺序无关。
    rid_to_target: dict[str, str] = {}
    for tag in re.findall(r"<Relationship\b[^>]*>", rels_xml):
        rid_m = re.search(r'\bId="([^"]+)"', tag)
        tgt_m = re.search(r'\bTarget="([^"]+)"', tag)
        if rid_m and tgt_m:
            rid_to_target[rid_m.group(1)] = tgt_m.group(1)
    sheet_file_by_name: dict[str, str] = {}
    for tag in re.findall(r"<sheet\b[^>]*>", workbook_xml):
        name_m = re.search(r'\bname="([^"]+)"', tag)
        rid_m = re.search(r'\br:id="([^"]+)"', tag)
        if not name_m or not rid_m:
            continue
        # Target 可能是相对(worksheets/sheet1.xml)或绝对(/xl/worksheets/...)
        target = rid_to_target.get(rid_m.group(1), "").lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        sheet_file_by_name[name_m.group(1)] = target

    # 3) 逐 sheet 在目标单元格的 <v> 前插入 <f>(仅匹配带缓存值的完整单元格;
    #    自闭合空单元格不在注入范围,行级掩码已排除)
    injected = 0
    for sheet_name, cell_formulas in sheet_plan.items():
        if not cell_formulas:
            continue
        xml_path = sheet_file_by_name.get(sheet_name)
        if not xml_path or xml_path not in entries:
            raise ValueError(f"导出文件中找不到工作表:{sheet_name}")
        xml = entries[xml_path].decode("utf-8")
        for ref, formula in cell_formulas.items():
            pattern = re.compile(
                r'(<c r="' + re.escape(ref) + r'"[^>]*>)(<v>.*?</v>)(</c>)'
            )
            xml, n = pattern.subn(
                lambda m: m.group(1) + "<f>" + formula + "</f>" + m.group(2) + m.group(3),
                xml,
                count=1,
            )
            injected += n
        entries[xml_path] = xml.encode("utf-8")

    if injected == 0:
        return 0

    # 4) 重写 zip:其余条目原样复制,保持字节不变
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for filename, data in entries.items():
            zout.writestr(filename, data)
    return injected


# ---------------------- 透视汇总算子 ----------------------
# 【clean-ai-derived-columns-pivot 新增】聚合方式 → 中文列名后缀
_PIVOT_AGG_CN: dict[str, str] = {
    "sum": "求和",
    "mean": "平均",
    "count": "计数",
    "max": "最大",
    "min": "最小",
}


def build_pivot_table(df: pd.DataFrame, config: dict[str, Any]) -> pd.DataFrame:
    """按配置生成一张**平面**透视汇总表。

    config:
        {row_dim, col_dim?, value, agg, include_margin}
        - row_dim: 行维度(必填,必须是 df 中存在的列)
        - col_dim: 列维度(可选;为空时只输出一列汇总值)
        - value:   汇总字段(必填,必须是数值列)
        - agg:     sum/mean/count/max/min(非法值直接 ValueError)
        - include_margin: 是否追加「总计」行/列

    输出列名:
        - col_dim 非空: <value>_<原列值>,例如 销售额_线上;
        - col_dim 为空: <value>_<聚合中文>,例如 销售额_求和;
        - 行维度列保持原名。

    返回 reset_index() 后的平面表,列名不含元组(避免导出 Excel 时表头异常)。
    """
    config = config or {}
    row_dim = str(config.get("row_dim") or "").strip()
    col_dim = str(config.get("col_dim") or "").strip()
    value = str(config.get("value") or "").strip()
    agg = str(config.get("agg") or "sum").strip().lower()
    include_margin = bool(config.get("include_margin", False))

    if agg not in _PIVOT_AGG_CN:
        raise ValueError(
            f"不支持的聚合方式:{agg},可选 sum/mean/count/max/min"
        )

    columns = [str(c) for c in df.columns]
    if not row_dim or row_dim not in columns:
        raise ValueError(f"透视表字段不存在:{row_dim}")
    if not value or value not in columns:
        raise ValueError(f"透视表字段不存在:{value}")
    if col_dim and col_dim not in columns:
        raise ValueError(f"透视表字段不存在:{col_dim}")

    # 汇总字段必须是数值列:先走 detect_column_types(兼容 Excel 文本型数字),
    # 再真正 to_numeric 一次,避免字符串求和把值拼接起来。
    if detect_column_types(df).get(value, "text") != "number":
        raise ValueError(f"汇总字段必须是数值列:{value}")

    work = df.copy()
    work[value] = pd.to_numeric(work[value], errors="coerce")
    if int(work[value].notna().sum()) == 0:
        raise ValueError(f"汇总字段没有可用数值:{value}")

    kwargs: dict[str, Any] = {
        "index": row_dim,
        "values": value,
        "aggfunc": agg,
        "margins": include_margin,
        "margins_name": "总计",
        "observed": False,
    }
    if col_dim:
        kwargs["columns"] = col_dim
    if agg in ("sum", "count"):
        # 求和/计数时空单元格补 0;平均/最大/最小保留 NaN(无数据就是无数据)
        kwargs["fill_value"] = 0

    pivot = pd.pivot_table(work, **kwargs).reset_index()

    # 展平列名。pivot_table 在单 value 时通常不是 MultiIndex,但带 margins
    # 或列维度较多时可能出现元组列名;这里统一取最后一级再拼前缀,
    # 保证导出 Excel 时表头是普通字符串,而不是 ("销售额", "线上") 元组。
    rename: dict[Any, str] = {}
    for col in pivot.columns:
        if str(col) == row_dim:
            continue
        label = col[-1] if isinstance(col, tuple) else col
        rename[col] = (
            f"{value}_{label}" if col_dim else f"{value}_{_PIVOT_AGG_CN[agg]}"
        )
    if rename:
        pivot = pivot.rename(columns=rename)
    return pivot


# ---------------------- 指标计算 ----------------------

def compute_indicators(df: pd.DataFrame) -> dict[str, Any]:
    """为 LLM 提供结构化指标 JSON。

    ===== 2026-10 三层增强(修复"年度报告只有文字,没有数据")=====

    旧实现把每张 sheet 当成"若干互不相干的数值列"做聚合,产出的是
    **跨量纲的垃圾数字**(实测年度复盘数据集):
        年度核心KPI::年度目标 = {sum: 67052, max: 60000, min: 0.35}
        月度经营数据::销售额 = {sum: 9360}   ← 真实全年 4680,汇总行重复计入
        渠道结构 by_group 里含 "合计: 4680"  ← 饼图最大的一块是伪类别
        trend = {}                          ← 月份列没被识别成时间轴
    LLM 拿到这种数字只能写"收入低速增长、利润承压"这种无数字废话
    (写具体数字必错),这就是报告"只有文字没有数据"的根因。

    本次加入三层前置处理,让喂给 LLM 的指标是**干净的真实业务数据**:

    第 1 层 · 合计行剔除(strip_total_rows)
        把"合计/小计/全年合计"等汇总行摘掉,避免 sum 翻倍、
        饼图出现伪类别。摘掉的行标签记入 total_rows_removed 字段留痕。

    第 2 层 · KPI 宽表识别(detect_kpi_matrix)
        识别"指标 × 目标/实际/达成率/同比"这种宽表,改为**逐指标**产出
        kpi_matrix.records:[{指标,单位,目标,实际,达成率,去年同期,同比}]。
        这类表**不再**进入 numeric_summary(跨量纲求和无意义),
        避免污染其它字段。

    第 3 层 · 周期列识别(is_period_column / period_sort_key)
        把"月份/季度"文本列识别成有序时间轴,让 trend 真正产出走势数据,
        且 groupby 按 1月→12月 排序而非字符串排序(否则顺序是
        1月,10月,11月,12月,2月...),汇总行不进趋势。

    其余逻辑(coerce 数值列、缺失率、topn、by_group、text_value_counts)
    保持原有行为,兼容既有前端/导出。
    """
    # ========== 第 1 层:剔除汇总行 ==========
    df_work, total_removed = strip_total_rows(df)

    types = detect_column_types(df_work)
    numeric_cols = [c for c, t in types.items() if t == "number"]
    date_cols = [c for c, t in types.items() if t == "date"]
    text_cols = [c for c, t in types.items() if t == "text"]

    # ========== 第 3 层:识别周期列(在 detect_column_types 之外) ==========
    # 周期列会被 detect_column_types 判成 text(因为 "1月" 不是合法日期),
    # 这里单独识别并从 text_cols 中摘掉,避免被当普通分类维度做 groupby。
    period_cols: list[str] = []
    for c in df_work.columns:
        if c in date_cols:
            continue
        try:
            if is_period_column(df_work[c], c):
                period_cols.append(c)
        except Exception:
            continue
    text_cols = [c for c in text_cols if c not in period_cols]

    # 把"被识别为 number 但 dtype 仍是 object"的列转成真正的数值列。
    # 这是 detect_column_types 增加"文本转数值兜底"后的必然后续处理,
    # 否则 sum/mean 会失败。优先用 coerce 后的 Series,不修改原 df。
    coerced_numeric: dict[str, pd.Series] = {}
    for col in numeric_cols:
        s = df_work[col]
        if pd.api.types.is_numeric_dtype(s):
            coerced_numeric[col] = s
            continue
        converted, success = _coerce_numeric(s)
        # 如果 coerce 后还有非空值,就用它;否则保留原 df 列(后续会被空过滤掉)
        coerced_numeric[col] = converted if success > 0 else s

    # 缺失率(百分比),不只给缺失计数,方便 LLM 写"X 列缺失 30% 以上"
    miss_counts = {c: int(v) for c, v in df_work.isna().sum().items() if v > 0}
    total_rows = max(1, int(len(df_work)))
    miss_ratio_top5 = {
        c: round(v / total_rows * 100, 1) for c, v in
        sorted(miss_counts.items(), key=lambda x: -x[1])[:5]
    }

    result: dict[str, Any] = {
        # row_count 报**剔除汇总行后**的真实明细行数,让 LLM 知道有多少条业务记录
        "row_count": int(len(df_work)),
        "column_count": int(len(df_work.columns)),
        "columns": [{"name": str(c), "type": types[c]} for c in df_work.columns],
        "numeric_summary": {},
        "topn": {},
        "by_group": {},
        "trend": {},
        "missing_top5": miss_counts,            # 兼容老字段:列 → 缺失数
        "missing_ratio_top5": miss_ratio_top5,  # 新字段:列 → 缺失率%
        "text_value_counts": {},                # 新字段:文本列频次 top5(无数值列时的兜底)
    }

    # 留痕:被剔除的汇总行,便于排查"为什么 sum 和 Excel 里不一样"
    if total_removed:
        result["total_rows_removed"] = total_removed
        result["raw_row_count"] = int(len(df))
    if period_cols:
        result["period_cols"] = [str(c) for c in period_cols]

    # ========== 第 2 层:KPI 宽表识别 ==========
    # 识别成功时:产出 kpi_matrix,并把 KPI 角色列从 numeric_cols 中排除,
    # 避免"年度目标/年度实际/达成率"被跨量纲求和(sum=67052 的元凶)。
    kpi = detect_kpi_matrix(df_work)
    if kpi:
        result["kpi_matrix"] = kpi
        kpi_value_cols = set(kpi.get("role_cols", {}).values())
        kpi_value_cols.add(kpi.get("metric_col"))
        if kpi.get("unit_col"):
            kpi_value_cols.add(kpi["unit_col"])
        numeric_cols = [c for c in numeric_cols if c not in kpi_value_cols]
        text_cols = [c for c in text_cols if c not in kpi_value_cols]
        # 已被 KPI 抽取的列,coerced_numeric 里的旧引用要同步剔除
        for c in list(coerced_numeric.keys()):
            if c in kpi_value_cols:
                coerced_numeric.pop(c, None)

    # 数值列汇总(用 coerce 后的 Series)
    for col in numeric_cols[:6]:
        s = coerced_numeric[col].dropna()
        if s.empty:
            continue
        result["numeric_summary"][col] = {
            "sum": float(round(s.sum(), 2)),
            "mean": float(round(s.mean(), 2)),
            "max": float(s.max()),
            "min": float(s.min()),
            "median": float(round(s.median(), 2)),
            "std": float(round(s.std(), 2)) if len(s) > 1 else 0.0,
        }
        # Top 5:把 coerce 后的 Series 临时合并,用于 nlargest,保持原 df 行顺序
        tmp_df = df_work.copy()
        tmp_df[col] = coerced_numeric[col]
        top = tmp_df.nlargest(5, col)[[c for c in df_work.columns]].head(5)
        result["topn"][col] = top.astype(object).infer_objects(copy=False).fillna("").values.tolist()

    # 按文本列聚合(原逻辑保持:有数值列时按第一列求和)
    for g in text_cols[:3]:
        if not numeric_cols:
            # 没有数值列时,改为统计"文本列自身频次",让 AI / 图表至少有一份
            vc = df_work[g].astype(str).fillna("(空值)").value_counts().head(8)
            result["text_value_counts"][g] = [
                {"name": str(k), "value": int(v)}
                for k, v in vc.items()
            ]
            continue
        agg_col = numeric_cols[0]
        # groupby 求和必须用 coerce 后的数值列
        grouped = (
            df_work.assign(**{agg_col: coerced_numeric[agg_col]})
                    .groupby(g)[agg_col].sum().sort_values(ascending=False).head(8)
        )
        result["by_group"][f"{g}_by_{agg_col}"] = [
            {"name": str(k), "value": float(round(v, 2))}
            for k, v in grouped.items()
        ]

    # ========== 趋势 ①:日期列 × 数值列(原有逻辑) ==========
    for d in date_cols[:1]:
        for n in numeric_cols[:2]:
            try:
                tmp = df_work[[d, n]].copy()
                tmp[n] = coerced_numeric[n]
                tmp = tmp.dropna()
                tmp[d] = pd.to_datetime(tmp[d], errors="coerce")
                tmp = tmp.dropna().sort_values(d)
                trend = tmp.groupby(d)[n].sum().reset_index()
                result["trend"][f"{n}_by_{d}"] = {
                    "x": [str(x.date()) if hasattr(x, "date") else str(x) for x in trend[d]],
                    "y": [float(round(v, 2)) for v in trend[n]],
                }
            except Exception:
                continue

    # ========== 趋势 ②:周期列(月份/季度)× 数值列(2026-10 新增) ==========
    # 这是让"月度经营数据"这类表真正产出走势数据的关键:
    #   旧实现 trend={} —— LLM 无任何走势素材,只能写定性废话;
    #   现在按 period_sort_key 排序(1月→12月,而非字符串顺序),
    #   并对每个数值列产出 {x:[月份], y:[值], 总计, 峰谷, 环比变化}。
    for p in period_cols[:1]:
        for n in numeric_cols[:3]:
            try:
                tmp = df_work[[p, n]].copy()
                tmp[n] = coerced_numeric[n]
                tmp = tmp.dropna(subset=[p])
                # 只保留能解析成周期的行(剔除"全年合计"等汇总行)
                tmp["__ord__"] = tmp[p].map(period_sort_key)
                tmp = tmp[tmp["__ord__"].apply(lambda t: t[0] != 99)]
                if tmp.empty:
                    continue
                # 去重:同一周期可能有重复行(如分渠道的多行),按周期求和。
                # 注意 groupby 会丢弃 "__ord__"(未参与聚合),所以聚合后必须
                # **重新计算**排序键,否则下面的 sort_values 会抛 KeyError
                # 被上一行的 except 吞掉,导致趋势静默丢失。
                agg = tmp.groupby(p, as_index=False)[n].sum()
                agg["__ord__"] = agg[p].map(period_sort_key)
                agg = agg.sort_values("__ord__")
                xs = [str(v) for v in agg[p].tolist()]
                ys = [float(round(float(v), 2)) for v in agg[n].tolist()]
                if not ys:
                    continue
                total_v = round(float(sum(ys)), 2)
                peak_i = max(range(len(ys)), key=lambda i: ys[i])
                trough_i = min(range(len(ys)), key=lambda i: ys[i])
                # 环比变化率(%):与上一周期比较,首项为 None
                mom: list[float | None] = [None]
                for i in range(1, len(ys)):
                    prev = ys[i - 1]
                    mom.append(
                        None if prev == 0 else round((ys[i] - prev) / abs(prev) * 100, 1)
                    )
                result["trend"][f"{n}_by_{p}"] = {
                    "x": xs,
                    "y": ys,
                    "总计": total_v,
                    "均值": round(total_v / len(ys), 2),
                    "峰值": {"周期": xs[peak_i], "值": ys[peak_i]},
                    "谷值": {"周期": xs[trough_i], "值": ys[trough_i]},
                    "环比变化率%": mom,
                }
            except Exception:
                continue

    return result


def indicators_to_charts(indicators: dict[str, Any]) -> list[dict[str, Any]]:
    """从指标生成默认 ECharts option。

    改版要点:
        - 原本只在 trend / by_group 非空时出图,导致"无图表数据"。
        - 现在增加文本列频次(text_value_counts)作为兜底饼图,
          即使没有任何数值列也能保证前端有一张图。

    ===== 2026-10 新增:KPI 宽表与周期趋势优先出图 =====
        修复"图表数据是编造的"问题 —— 之前这里只处理通用的
        trend / by_group,而真实业务数据里最有价值的两种结构
        (KPI 宽表、月度趋势)根本走不到这条路径,导致 LLM 自己
        编数据画图。现在:
          1) 有 kpi_matrix → 产出「目标vs实际」「达成率」「同比」三张真实图;
          2) 周期趋势(trend 里带 总计/峰值/环比变化率% 的那些)→ 画折线 + 环比柱,
             X 轴严格按 1月→12月 顺序。
    """
    charts: list[dict[str, Any]] = []

    # ---------- 【新增】KPI 宽表:目标vs实际 / 达成率 / 同比 ----------
    kpi = indicators.get("kpi_matrix")
    if isinstance(kpi, dict) and kpi.get("records"):
        charts.extend(kpi_matrix_to_charts(kpi))

    # 趋势线
    for k, v in (indicators.get("trend") or {}).items():
        if not isinstance(v, dict) or not v.get("x"):
            continue
        # 周期类趋势:额外画一张环比变化率柱状图(真实数据,非编造)
        is_period_trend = "环比变化率%" in v
        charts.append({
            "title": f"{k} 趋势",
            "type": "line",
            "option": {
                "title": {"text": k, "textStyle": {"color": "#5eead4"}},
                "tooltip": {"trigger": "axis"},
                "xAxis": {"type": "category", "data": v["x"], "axisLabel": {"color": "#94a3b8"}},
                "yAxis": {"type": "value", "axisLabel": {"color": "#94a3b8"}},
                "series": [{"type": "line", "data": v["y"], "smooth": True,
                            "lineStyle": {"color": "#22d3ee"},
                            "areaStyle": {"color": "rgba(34, 211, 238, 0.25)"}}],
            },
        })
        if is_period_trend:
            mom = v.get("环比变化率%") or []
            charts.append({
                "title": f"{k} 环比变化率",
                "type": "bar",
                "option": {
                    "title": {"text": f"{k} 环比变化率(%)",
                              "textStyle": {"color": "#5eead4"}},
                    "tooltip": {"trigger": "axis"},
                    "xAxis": {"type": "category", "data": v["x"],
                              "axisLabel": {"color": "#94a3b8"}},
                    "yAxis": {"type": "value", "name": "环比%", "axisLabel": {"color": "#94a3b8"}},
                    "series": [{
                        "type": "bar", "data": mom,
                        "itemStyle": {"color": "#fbbf24"},
                        "label": {"show": True, "position": "top",
                                  "color": "#cbd5e1", "fontSize": 10,
                                  "formatter": "{c}%"},
                    }],
                },
            })

    # 分组饼图
    for k, items in (indicators.get("by_group") or {}).items():
        charts.append({
            "title": f"{k} 占比",
            "type": "pie",
            "option": {
                "title": {"text": k, "textStyle": {"color": "#5eead4"}},
                "tooltip": {"trigger": "item"},
                "legend": {"textStyle": {"color": "#94a3b8"}},
                "series": [{
                    "type": "pie",
                    "radius": ["40%", "70%"],
                    "data": [{"name": it["name"], "value": it["value"]} for it in items],
                    "label": {"color": "#cbd5e1"},
                }],
            },
        })

    # 文本列频次饼图(兜底:无数值列时仍能出图)
    for k, items in (indicators.get("text_value_counts") or {}).items():
        if not items:
            continue
        charts.append({
            "title": f"{k} 频次分布",
            "type": "pie",
            "option": {
                "title": {"text": f"{k} 频次分布", "textStyle": {"color": "#5eead4"}},
                "tooltip": {"trigger": "item"},
                "legend": {"textStyle": {"color": "#94a3b8"}},
                "series": [{
                    "type": "pie",
                    "radius": ["40%", "70%"],
                    "data": [{"name": it["name"], "value": it["value"]} for it in items],
                    "label": {"color": "#cbd5e1"},
                }],
            },
        })

    return charts


# ---------------------- 数据集对比算子（dataset-compare-analyst-narration） ----------------------
# 四个 mode 共用的算子:输入两份 DataFrame + 用户指定的列,产出对齐后的 diff_details。
# 异常一律抛 ValueError(列不存在 / 非数值 / 关键列缺失),
# 让 /api/tools/compare 把 4xx 状态码 + 中文 message 直接回前端。

# 「左/右均缺失/无法解析」时的占位值,前端 el-table 里用 — 渲染。
_MISSING = -1.0


def _validate_cols(df_l: pd.DataFrame, df_r: pd.DataFrame, cols: list[str]) -> None:
    """校验一组列在两份 DataFrame 中都存在;不存在则抛 ValueError(中文 message)。

    为什么:前端 el-select 可能因为缓存/旧数据出现"列名已不在表里"的脏值,
    后端必须显式拦截,而不是让 pandas 抛 KeyError 让前端看到 500。
    """
    missing_l = [c for c in cols if c not in df_l.columns]
    missing_r = [c for c in cols if c not in df_r.columns]
    if missing_l or missing_r:
        sides = []
        if missing_l:
            sides.append(f"左数据集缺失: {', '.join(missing_l)}")
        if missing_r:
            sides.append(f"右数据集缺失: {', '.join(missing_r)}")
        raise ValueError("列校验失败:" + "; ".join(sides))


def _safe_numeric(series: pd.Series) -> pd.Series:
    """把 groupby 求和的 metric 列转成数值,无法解析的置 0(不影响聚合)。"""
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0)
    coerced, _ = _coerce_numeric(series)
    return pd.to_numeric(coerced, errors="coerce").fillna(0)


def _normalize_group_key(value: Any) -> str:
    """把分组键规整为字符串(用于内连接 / TopN 对齐)。
    NaN/None → 空串;数字/日期统一转 str,避免左右两表类型不一致导致 join miss。
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value)


def compare_by_group(
    df_l: pd.DataFrame,
    df_r: pd.DataFrame,
    group_col: str,
    metric_col: str,
    top_n: int | None = None,
) -> list[dict[str, Any]]:
    """业务对比:按 group_col 对 metric_col 求和,产出对齐后的 diff 行。

    返回:[
        { "group": "1月", "left_sum": 100.0, "right_sum": 150.0,
          "diff": 50.0, "growth_pct": 50.0 },
        ...
    ]

    规则:
        - 缺失侧用 _MISSING(-1.0) 标识,前端 el-table 渲染成 "—"
        - 增长率:左为 0 时记 None(避免除零)
        - 分组键:左右并集,保持左表出现顺序,新出现的追加在末尾

    【topn 并入 business 新增】top_n 参数:
        - None(默认):返回全部分组(旧行为不变)
        - 正整数:按「左右两侧合并后的指标值从大到小」排序,只保留前 N 个分组,
          并为每行附加**辅助字段** left_rank / right_rank(两侧各自在该分组全集里的名次)。
          排名只作参考信息,不参与 diff / growth_pct 的计算——因为名次变化可能
          只是「别人插进来」造成的相对位移,不代表经营指标本身涨跌
          (参见下方"排名误导"问题:右表新增一个爆款会让原 Top10 全体"排名下降",
           但它们各自的指标值其实一点没变)。
    """
    _validate_cols(df_l, df_r, [group_col, metric_col])

    # 把 metric 列转成数值(避免 object 列求和退化为字符串拼接)
    l_metric = _safe_numeric(df_l[metric_col])
    r_metric = _safe_numeric(df_r[metric_col])
    l_grp = df_l[group_col].map(_normalize_group_key)
    r_grp = df_r[group_col].map(_normalize_group_key)

    left_sum = pd.Series(l_metric.values, index=l_grp.values).groupby(level=0).sum()
    right_sum = pd.Series(r_metric.values, index=r_grp.values).groupby(level=0).sum()

    # 按左表出现顺序,再追加仅右表存在的分组
    ordered: list[str] = []
    seen: set[str] = set()
    for k in left_sum.index.tolist():
        if k not in seen:
            ordered.append(k)
            seen.add(k)
    for k in right_sum.index.tolist():
        if k not in seen:
            ordered.append(k)
            seen.add(k)

    rows: list[dict[str, Any]] = []
    for k in ordered:
        lv = float(left_sum.get(k, 0.0)) if k in left_sum.index else _MISSING
        rv = float(right_sum.get(k, 0.0)) if k in right_sum.index else _MISSING
        if lv == _MISSING or rv == _MISSING:
            diff: float | None = None
            growth: float | None = None
        else:
            diff = round(rv - lv, 2)
            growth = round((diff / lv * 100.0), 2) if lv != 0 else None
        rows.append({
            "group": k,
            "left_sum": lv,
            "right_sum": rv,
            "diff": diff,
            "growth_pct": growth,
        })

    # —— TopN 截断(可选) ——
    if top_n is not None and int(top_n) > 0:
        limit = int(top_n)

        # 两侧各自在"自己全集"里的名次(缺失侧不参与排名,记 0 表示不在榜)
        def _ranks(series: pd.Series) -> dict[str, int]:
            ordered_vals = series.sort_values(ascending=False)
            return {str(k): i + 1 for i, k in enumerate(ordered_vals.index.tolist())}

        l_rank = _ranks(left_sum)
        r_rank = _ranks(right_sum)

        # 排序依据:两侧都存在时取较大值(单侧存在时取该侧值),
        # 让「左表才有的老业务」和「右表才有的新业务」都能进榜
        def _sort_key(row: dict[str, Any]) -> float:
            lv = row["left_sum"]
            rv = row["right_sum"]
            vals = [v for v in (lv, rv) if v != _MISSING]
            return max(vals) if vals else float("-inf")

        rows.sort(key=_sort_key, reverse=True)
        rows = rows[:limit]

        for row in rows:
            g = row["group"]
            row["left_rank"] = l_rank.get(g, 0)
            row["right_rank"] = r_rank.get(g, 0)

    return rows


def row_level_diff(
    df_l: pd.DataFrame,
    df_r: pd.DataFrame,
    key_cols: list[str],
    modified_limit: int = 50,
) -> dict[str, Any]:
    """变更追踪:以 key_cols 为主键,产出 added/removed/modified 三类行级 diff。

    返回:{
        "added":   [{ "key": {...}, "row": {...} }, ...],
        "removed": [{ "key": {...}, "row": {...} }, ...],
        "modified":[{ "key": {...}, "field_changes": [{"field", "before", "after"}] }, ...],
        "modified_total": int,   # modified 真实条数(可能 > modified_limit)
    }

    规则:
        - key_cols 全部存在性校验;任一缺失抛 ValueError
        - 复合主键用 tuple 字符串作为 join key(避免列类型不一致)
        - modified:主键相同的行,逐字段对比,值不同的字段加入 field_changes
        - modified > modified_limit:只保留前 N 条,modified_total 仍记真实数
    """
    if not key_cols:
        raise ValueError("变更追踪至少选择 1 个主键列")
    _validate_cols(df_l, df_r, key_cols)

    # 复合主键转成单列字符串,便于 merge/索引
    def _key_col(df: pd.DataFrame) -> pd.Series:
        keys = df[key_cols].astype(object).fillna("")
        if len(key_cols) == 1:
            return keys[key_cols[0]].map(_normalize_group_key)
        return keys.apply(lambda r: "||".join(_normalize_group_key(v) for v in r), axis=1)

    l_keys = _key_col(df_l)
    r_keys = _key_col(df_r)
    df_l2 = df_l.copy()
    df_l2["__k__"] = l_keys.values
    df_r2 = df_r.copy()
    df_r2["__k__"] = r_keys.values

    # 排除全空主键(避免空串误进 added/removed)
    df_l2 = df_l2[df_l2["__k__"] != ""]
    df_r2 = df_r2[df_r2["__k__"] != ""]

    l_idx = set(df_l2["__k__"].tolist())
    r_idx = set(df_r2["__k__"].tolist())

    added_keys = sorted(r_idx - l_idx)
    removed_keys = sorted(l_idx - r_idx)
    common_keys = sorted(l_idx & r_idx)

    # 构建主键→整行的索引(取首条重复主键,简化展示)
    def _row_of(df: pd.DataFrame, k: str) -> dict[str, Any]:
        sub = df[df["__k__"] == k].head(1)
        if sub.empty:
            return {}
        rec = sub.iloc[0].to_dict()
        rec.pop("__k__", None)
        # 字段值规整(NaN → None,避免 JSON 报 NaN)
        return {
            k2: (None if pd.isna(v2) else v2)
            for k2, v2 in rec.items()
        }

    added = [
        {"key": k, "row": _row_of(df_r2, k)}
        for k in added_keys
    ]
    removed = [
        {"key": k, "row": _row_of(df_l2, k)}
        for k in removed_keys
    ]

    # modified:逐字段对比
    compare_fields: list[str] = []
    for c in df_l2.columns:
        if c == "__k__":
            continue
        if c in df_r2.columns:
            compare_fields.append(c)
    # 还要把"只在右表存在"的列也算上(等价于 "字段被删,值是 None")
    only_right = [c for c in df_r2.columns if c not in df_l2.columns and c != "__k__"]
    compare_fields.extend(only_right)

    modified: list[dict[str, Any]] = []
    for k in common_keys:
        row_l = df_l2[df_l2["__k__"] == k].iloc[0]
        row_r = df_r2[df_r2["__k__"] == k].iloc[0]
        changes: list[dict[str, Any]] = []
        for f in compare_fields:
            v_l = row_l.get(f) if f in row_l.index else None
            v_r = row_r.get(f) if f in row_r.index else None
            # NaN 归一
            try:
                l_is_na = pd.isna(v_l)
            except Exception:
                l_is_na = False
            try:
                r_is_na = pd.isna(v_r)
            except Exception:
                r_is_na = False
            if l_is_na and r_is_na:
                continue
            if (l_is_na != r_is_na) or (v_l != v_r):
                changes.append({
                    "field": f,
                    "before": (None if l_is_na else v_l),
                    "after": (None if r_is_na else v_r),
                })
        if changes:
            modified.append({"key": k, "field_changes": changes})

    modified_total = len(modified)
    truncated = modified_total > modified_limit
    if truncated:
        modified = modified[:modified_limit]

    return {
        "added": added,
        "removed": removed,
        "modified": modified,
        "modified_total": modified_total,
        "modified_truncated": truncated,
    }


# ============================================================================
# 【tools-results-history-persistence 新增】3 个工具结果落库用图生成函数
# 用途:把工作台工具(校验 / 对话 / 对比)的结果落库为历史任务时,生成对应的
# ECharts option。所有数据由对应算子直接计算,LLM 不参与,保证落库操作毫秒级。
# 输出形态:list[dict],每项为 {"type", "title", "option"},与 ReportView 的
# chart_option.by_section["数据图表"] 列表兼容(ReportView 用 v-chart 直接渲染)。
# ============================================================================


def build_validation_chart_option(
    quality: dict,
    column_types: dict,
    *,
    top_n_missing: int = 10,
) -> list[dict]:
    """数据校验·落库用图:① 缺失率 Top10 柱状图 ② 字段类型分布饼图。

    Args:
        quality: basic_quality() 输出,至少含 "missing_by_column"(字段名 → 缺失率)。
                 兼容 "row_count" / "missing_total" / "duplicate_rows" 等字段,本函数
                 不强依赖这些,但缺时不影响渲染。
        column_types: detect_column_types() 输出,字段名 → number / date / text。
        top_n_missing: 缺失率柱状图最多取几个字段(默认 10)。

    Returns:
        list[dict]: 至少 1 项,典型 2 项,格式:
            [{
                "type": "bar" | "pie",
                "title": "...",
                "option": {...完整的 ECharts option...},
            }, ...]

    设计要点:
        - 缺失率为 0 的字段**不**画进柱状图,标题写"无缺失"也能渲染。
        - 类型分布饼图用 "numeric" / "date" / "text" 3 类计数。
        - 类型字段不在预期集合(如 enum/array)时归入 "text" 兜底。
    """
    charts: list[dict] = []

    # ----- 图①:缺失率 Top10 柱状图 -----
    miss_by_col = quality.get("missing_by_column") or {}
    # 过滤出缺失数 > 0 的字段,按缺失数降序取前 N
    miss_items: list[tuple[str, int]] = [
        (str(k), int(v)) for k, v in miss_by_col.items() if int(v) > 0
    ]
    miss_items.sort(key=lambda x: -x[1])
    miss_items = miss_items[: max(0, int(top_n_missing))]

    if miss_items:
        bar_x = [k for k, _ in miss_items]
        bar_y = [int(v) for _, v in miss_items]
        title_txt = f"缺失率 Top{len(bar_x)}"
    else:
        bar_x = []
        bar_y = []
        title_txt = "缺失率(无缺失)"

    charts.append({
        "type": "bar",
        "title": title_txt,
        "option": {
            "title": {"text": title_txt, "left": "center"},
            "tooltip": {"trigger": "axis"},
            "grid": {"left": 60, "right": 30, "top": 50, "bottom": 60},
            "xAxis": {
                "type": "category",
                "data": bar_x,
                "axisLabel": {"rotate": 30, "interval": 0, "fontSize": 11},
            },
            "yAxis": {"type": "value", "name": "缺失数"},
            "series": [{
                "type": "bar",
                "data": bar_y,
                "itemStyle": {"color": "#fbbf24"},
                "label": {"show": True, "position": "top", "fontSize": 11},
            }],
        },
    })

    # ----- 图②:字段类型分布饼图 -----
    type_counter = {"numeric": 0, "date": 0, "text": 0}
    for _col, t in (column_types or {}).items():
        if t in type_counter:
            type_counter[t] += 1
        else:
            # 兜底:未知类型归入 text
            type_counter["text"] += 1
    pie_data = [
        {"name": k, "value": v} for k, v in type_counter.items() if v > 0
    ] or [{"name": "无字段", "value": 0}]

    charts.append({
        "type": "pie",
        "title": "字段类型分布",
        "option": {
            "title": {"text": "字段类型分布", "left": "center"},
            "tooltip": {"trigger": "item"},
            "legend": {"bottom": 0},
            "series": [{
                "type": "pie",
                "radius": ["35%", "60%"],
                "center": ["50%", "45%"],
                "data": pie_data,
                "label": {"formatter": "{b}: {c}"},
            }],
        },
    })

    return charts


def build_nl_query_chart_option(
    rows: list,
    columns: list | None = None,
    *,
    max_rows: int = 20,
) -> list[dict]:
    """数据对话·落库用图。

    Args:
        rows: NL 查询结果行。可能是:
              - 数值聚合结果,形如 [["字段A", 123], ["字段B", 456], ...];
              - 明细行,形如 [[col1_val, col2_val, ...], ...];
              - 或 `df.head(5)` 的二维 list。
        columns: 列名列表(可选)。明细行场景下应有与 rows 列数一致的列名;
                 聚合结果场景下 columns 可传 `["字段", "值"]`。
        max_rows: 数值柱状图最多取几个字段/行(默认 20)。

    Returns:
        list[dict]: 至少 1 项。
            - 当 rows 为数值聚合(每行恰好 2 列、第 2 列为数字):
              返回字段-值柱状图;
            - 否则:返回 1 张"结果明细"占位图(report 渲染时 series 为空,
              但保证至少 1 张图项落库,前端不会渲染空白)。

    判定"是否数值聚合"的规则:
        - len(rows) > 0
        - 每行是 list 且 len == 2
        - 每行第 2 个元素能转 float 且非 None
    """
    charts: list[dict] = []

    is_aggregate = False
    if rows and isinstance(rows, list):
        ok = True
        for r in rows[: max(1, min(len(rows), 5))]:  # 抽样前 5 行判定
            if not isinstance(r, list) or len(r) != 2:
                ok = False
                break
            try:
                float(r[1])
            except (TypeError, ValueError):
                ok = False
                break
        is_aggregate = ok and len(rows) > 0

    if is_aggregate:
        items = rows[: max(0, int(max_rows))]
        bar_x: list = []
        bar_y: list = []
        for r in items:
            label = r[0] if r[0] not in (None, "") else "—"
            try:
                val = float(r[1])
            except (TypeError, ValueError):
                continue
            bar_x.append(str(label))
            bar_y.append(val)
        charts.append({
            "type": "bar",
            "title": "数值结果",
            "option": {
                "title": {"text": "数值结果", "left": "center"},
                "tooltip": {"trigger": "axis"},
                "grid": {"left": 60, "right": 30, "top": 50, "bottom": 60},
                "xAxis": {
                    "type": "category",
                    "data": bar_x,
                    "axisLabel": {"rotate": 30, "interval": 0, "fontSize": 11},
                },
                "yAxis": {"type": "value"},
                "series": [{
                    "type": "bar",
                    "data": bar_y,
                    "itemStyle": {"color": "#22d3ee"},
                    "label": {"show": True, "position": "top", "fontSize": 11},
                }],
            },
        })
    else:
        # 占位图:明细/文本结果时,保证至少 1 张图项落库
        charts.append({
            "type": "table",
            "title": "结果明细",
            "option": {
                "title": {"text": "问答结果", "left": "center"},
                "tooltip": {"trigger": "item"},
                "series": [],   # series 为空时前端 v-chart 仍会渲染标题,不会报错
            },
        })

    return charts


def build_compare_chart_option(
    diff_mode: str,
    diff_details,
    left_dataset_name: str,
    right_dataset_name: str,
) -> list[dict]:
    """数据集对比·落库用图(沿用对比 spec 设计,这里独立实现以便校验)。

    Args:
        diff_mode: "business" | "row_diff"。
        diff_details: 对应 mode 的明细 dict / list(来自 ai_service.compare_datasets)。
        left_dataset_name: 左数据集名(用于标题/legend)。
        right_dataset_name: 右数据集名(用于标题/legend)。

    Returns:
        list[dict]: 至少 1 项;空 diff_details 时返回 1 张"对比摘要(空)"占位图。

    模式对应:
        - business: 左右双系列柱状图(group × left_sum/right_sum);
        - row_diff: ① 计数饼图(新增/删除/修改);② 修改字段频次柱状图。

    【topn 并入 business】原先独立的 "topn" 模式已下线:
        排行只是 business 聚合结果的另一种视角,拆成独立模式会导致
        "别人插进来 → 全体名次下降"这类相对位移被误读为经营下滑。
        TopN 现在是 business 的可选截断参数,图表仍是同一张左右对比柱状图。
    """
    charts: list[dict] = []

    # 防御性:diff_details 为 None 时返回占位
    if diff_details is None:
        diff_details = []

    if diff_mode == "business":
        details = diff_details if isinstance(diff_details, list) else []
        bar_x = [str(d.get("group", "—")) for d in details]
        bar_left = [
            d.get("left_sum") if d.get("left_sum") != -1 else None
            for d in details
        ]
        bar_right = [
            d.get("right_sum") if d.get("right_sum") != -1 else None
            for d in details
        ]
        title = f"业务对比 · {left_dataset_name} vs {right_dataset_name}"
        if not bar_x:
            charts.append({
                "type": "bar", "title": "对比摘要(空)", "option": {
                    "title": {"text": "对比结果为空", "left": "center"},
                    "series": [],
                },
            })
        else:
            charts.append({
                "type": "bar", "title": title, "option": {
                    "title": {"text": title, "left": "center"},
                    "tooltip": {"trigger": "axis"},
                    "legend": {"top": 25},
                    "grid": {"left": 70, "right": 30, "top": 70, "bottom": 60},
                    "xAxis": {
                        "type": "category", "data": bar_x,
                        "axisLabel": {"rotate": 30, "interval": 0, "fontSize": 11},
                    },
                    "yAxis": {"type": "value"},
                    "series": [
                        {"name": left_dataset_name, "type": "bar", "data": bar_left},
                        {"name": right_dataset_name, "type": "bar", "data": bar_right},
                    ],
                },
            })

    elif diff_mode == "row_diff":
        details = diff_details if isinstance(diff_details, dict) else {}
        added = int((details.get("added_count") if isinstance(details.get("added_count"), int) else 0) if details else 0)
        removed = int((details.get("removed_count") if isinstance(details.get("removed_count"), int) else 0) if details else 0)
        # row_diff 的 indicators 里通常有 added/removed/modified count;从 details 里直接拿也可
        try:
            added = max(added, len(details.get("added") or []))
        except Exception:
            pass
        try:
            removed = max(removed, len(details.get("removed") or []))
        except Exception:
            pass
        modified = len(details.get("modified") or []) if isinstance(details, dict) else 0

        # 计数饼图
        charts.append({
            "type": "pie", "title": "变更计数(新增/删除/修改)", "option": {
                "title": {"text": "变更计数", "left": "center"},
                "tooltip": {"trigger": "item"},
                "legend": {"bottom": 0},
                "series": [{
                    "type": "pie", "radius": ["35%", "60%"],
                    "center": ["50%", "45%"],
                    "data": [
                        {"name": "新增", "value": added},
                        {"name": "删除", "value": removed},
                        {"name": "修改", "value": modified},
                    ],
                    "label": {"formatter": "{b}: {c}"},
                }],
            },
        })

        # 修改字段频次柱状图(从 modified[*].changes[*].field 计数)
        field_counter: dict[str, int] = {}
        if isinstance(details, dict):
            for m in details.get("modified") or []:
                for ch in (m.get("changes") or []):
                    fld = str(ch.get("field") or "未知字段")
                    field_counter[fld] = field_counter.get(fld, 0) + 1
        if field_counter:
            items = sorted(field_counter.items(), key=lambda x: -x[1])[:10]
            charts.append({
                "type": "bar", "title": "修改字段频次 Top10", "option": {
                    "title": {"text": "修改字段频次 Top10", "left": "center"},
                    "tooltip": {"trigger": "axis"},
                    "grid": {"left": 60, "right": 30, "top": 50, "bottom": 60},
                    "xAxis": {
                        "type": "category",
                        "data": [k for k, _ in items],
                        "axisLabel": {"rotate": 30, "interval": 0, "fontSize": 11},
                    },
                    "yAxis": {"type": "value"},
                    "series": [{
                        "type": "bar",
                        "data": [v for _, v in items],
                        "itemStyle": {"color": "#fbbf24"},
                    }],
                },
            })

    else:
        # 未知 mode:兜底返回占位
        charts.append({
            "type": "bar", "title": "对比摘要(空)", "option": {
                "title": {"text": "对比结果为空", "left": "center"},
                "series": [],
            },
        })

    # 兜底:任何路径下若 charts 仍为空(极端),塞一张占位
    if not charts:
        charts.append({
            "type": "bar", "title": "对比摘要(空)", "option": {
                "title": {"text": "对比结果为空", "left": "center"},
                "series": [],
            },
        })

    return charts


# ----------------------------------------------------------------------
# 【clean-keep-formulas】公式保留回写
# ----------------------------------------------------------------------
def _anchor_eq(a, b) -> bool:
    """锚点宽松比较:数值按大小(120000.0 == 120000),文本按完全相等。

    用于"列N"空表头列的定位——df 侧值可能是 numpy 类型/被 unify_date
    格式化过,与 openpyxl 缓存值网格在类型上有差异,需宽松对齐。
    """
    try:
        if a is None or b is None:
            return False
        # 数值路径:两边都能转 float 则按数值比较
        fa, fb = float(a), float(b)
        return abs(fa - fb) < 1e-9
    except (TypeError, ValueError):
        pass
    # 文本路径:统一转字符串比较(去掉首尾空白)
    return str(a).strip() == str(b).strip()


def write_cleaned_keep_formulas(
    src_path: str | Path,
    out_path: str | Path,
    cleaned_sheets: dict[str, "pd.DataFrame"],
) -> dict:
    """把清洗后的值回写进"原工作簿副本",尽量保留公式 / 样式 / 多 sheet 布局。

    背景:清洗管线基于 pandas 数值运算,pd.read_excel 只能读到"上次保存时
    算好的缓存值",公式(如 =SUMIF(...))在读取阶段就已丢失,普通清洗落盘
    (to_excel)写出的自然是"死值"。本函数换一条落盘路径——用 openpyxl 以
    公式视图(data_only=False)打开原工作簿,只改数据区的值,其余一切不动,
    保存后 Excel/WPS 打开时会自动重算公式。

    规则:
      1) 用与 read_dataframe_smart_header 一致的占位行/数据行算法定位表头行;
      2) 从表头下一行起把清洗后的 DataFrame 逐格写回:
         - 源格是公式(openpyxl data_type == 'f')→ 保留公式不覆盖
           (其引用列的值更新后,公式重算结果自动跟着变);
         - 源格是普通值 → 覆盖为清洗值;
         - DataFrame 列多于原表(如 mark_anomaly 追加的 _is_outlier 列)
           → 在原表右侧追加新列(含表头行);
      3) 清洗后行数变少 → 尾部多余的原数据行整行清空(含公式),
         避免残留已在清洗中删除的行;
      4) 合并单元格(标题行等)不在写值范围内,天然不受影响。

    参数:
        src_path:      原始 Excel 路径(必须 .xlsx/.xlsm,CSV 无公式概念)
        out_path:      输出路径(建议同一目录新文件名,勿覆盖源文件)
        cleaned_sheets:{sheet 名: 清洗后的 DataFrame},列顺序须与原表一致
                       (read_dataframe_per_sheet + pe.clean 的产物即满足)

    返回:
        {"sheets": [{name, rows_written, cols_written}, ...],
         "preserved_formulas": int,   # 保留未覆盖的公式单元格数
         "cleared_cells": int}        # 尾部清空的单元格数
    """
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell

    p = Path(src_path)
    wb = load_workbook(p)        # 公式视图:用于写回(公式以 "=..." 保留)
    wb_vals = load_workbook(p, data_only=True)  # 缓存值视图:仅用于定位表头
    report: dict = {"sheets": [], "preserved_formulas": 0, "cleared_cells": 0}

    for name, df in cleaned_sheets.items():
        if name not in wb.sheetnames:
            # 清洗结果里的 sheet 在原工作簿不存在(理论不会发生),跳过
            continue
        ws = wb[name]

        # —— 1. 用"缓存值网格"定位表头 ——
        # 必须用 data_only=True 的网格:它与 pd.read_excel(header=None) 所见
        # 完全一致(公式格=计算值)。若用公式视图的网格,"=SUMIF(...)" 这类
        # 公式字符串会干扰 _looks_like_data_row 的数字判定,导致表头错位。
        grid: list[list] = [list(row) for row in wb_vals[name].iter_rows(values_only=True)]
        if not grid:
            continue
        max_col = max(len(r) for r in grid)  # 全表最大列数

        # —— 2. 定位表头行(与 read_dataframe_smart_header 同一套判定) ——
        header_idx: int | None = None
        for i, vals in enumerate(grid):
            if _is_placeholder_row(vals):
                continue
            if not any(v is not None and str(v) != "" for v in vals):
                continue  # 整行空,跳过
            if _looks_like_data_row(vals):
                # 当前是数据行 → 表头在它上方最近的"说明行"处
                for j in range(i - 1, -1, -1):
                    pv = grid[j]
                    if _is_placeholder_row(pv) or _looks_like_data_row(pv):
                        continue
                    if any(v is not None and str(v) != "" for v in pv):
                        header_idx = j
                        break
                if header_idx is None:
                    header_idx = max(0, i - 1)
                break
            else:
                # 第一个"非占位、非数据"的行本身就是表头
                header_idx = i
                break
        if header_idx is None:
            header_idx = 0
        data_start = header_idx + 1  # 数据区起始行(0-based)
        header_vals = grid[header_idx] if header_idx < len(grid) else []

        # —— 3. 建 df 列 → Excel 列号映射 ——
        # 智能表头读取时会删掉全空列、给空表头命名"列N",df 的列**位置**
        # 未必等于原表列位置,必须显式对齐,否则数据会整体左/右移一列。
        # 匹配优先级:① 表头行同名未用列;② 列名"列N" 且原表第 N 列表头为空
        # → 按位置 N;③ 都失败 → 作为新增列追加到原表右侧。
        df_cols = [str(c) for c in df.columns]
        used: set[int] = set()
        col_map: dict[int, int] = {}  # df 列索引 → 1-based Excel 列号
        pending_new: list[int] = []
        for ci, col_name in enumerate(df_cols):
            matched = None
            for xi, hv in enumerate(header_vals):
                xn = xi + 1
                if xn in used or hv is None:
                    continue
                if str(hv).strip() == col_name:
                    matched = xn
                    break
            if matched is not None:
                col_map[ci] = matched
                used.add(matched)
                continue
            # —— 锚点匹配:针对"列N"这类空表头列 ——
            # smart_header 给无表头列命名的"列N"中,N 是删空列后的 df 位置,
            # 不等于原表列位置,不能直接按 N 对齐。改为拿该列第一个非空值
            # 作锚点,在原表"表头为空的未用列"的数据区里搜索定位。
            anchor = None
            for v in df.iloc[:, ci]:
                if v is not None and not pd.isna(v):
                    anchor = v
                    break
            if anchor is not None:
                hit = None
                for cn in range(1, max_col + 1):
                    if cn in used or (cn - 1 < len(header_vals) and header_vals[cn - 1] is not None):
                        continue  # 已用 / 有表头名的列不参与
                    for r in range(data_start, min(len(grid), data_start + 50)):
                        gv = grid[r][cn - 1] if cn - 1 < len(grid[r]) else None
                        if gv is None or pd.isna(gv):
                            continue
                        if _anchor_eq(anchor, gv):
                            hit = cn
                            break
                    if hit is not None:
                        break
                if hit is not None:
                    col_map[ci] = hit
                    used.add(hit)
                    continue
            m = re.fullmatch(r"列(\d+)", col_name)
            if m:
                pos = int(m.group(1))
                hv = header_vals[pos - 1] if pos - 1 < len(header_vals) else None
                if 1 <= pos <= max_col and (hv is None or str(hv).strip() == ""):
                    col_map[ci] = pos
                    used.add(pos)
                    continue
            pending_new.append(ci)
        # 新增列依次排在原表最右侧
        next_col = max_col
        for ci in pending_new:
            next_col += 1
            col_map[ci] = next_col
            used.add(next_col)
            hcell = ws.cell(row=header_idx + 1, column=next_col)
            if not isinstance(hcell, MergedCell):
                hcell.value = df_cols[ci]

        # —— 4. 准备写回矩阵(numpy → Python 原生类型,openpyxl 才认) ——
        values = df.to_numpy(dtype=object, na_value=None)
        n_rows = values.shape[0]

        # —— 5. 按映射逐格写回;公式格保留 ——
        for r in range(n_rows):
            excel_row = data_start + r + 1  # openpyxl 行号从 1 起
            for ci in range(len(df_cols)):
                cell = ws.cell(row=excel_row, column=col_map[ci])
                if isinstance(cell, MergedCell):
                    continue
                if cell.data_type == "f":
                    # 公式格:不覆盖,保留公式(打开时自动重算)
                    report["preserved_formulas"] += 1
                    continue
                v = values[r][ci]
                if v is not None and hasattr(v, "item"):
                    try:
                        v = v.item()  # numpy.int64/float64 → Python 原生
                    except (ValueError, AttributeError):
                        pass
                cell.value = v

        # —— 6. 清洗后行数变少:清空尾部多余的原数据行(含公式) ——
        write_cols = sorted(used)
        for r in range(data_start + n_rows, len(grid)):
            for cn in write_cols:
                cell = ws.cell(row=r + 1, column=cn)
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                cell.value = None
                report["cleared_cells"] += 1

        report["sheets"].append({
            "name": name,
            "rows_written": int(n_rows),
            "cols_written": len(df_cols),
            "header_row": header_idx + 1,  # 1-based,便于前端/日志核对
            "col_map": {df_cols[ci]: cn for ci, cn in col_map.items()},
        })

    wb.save(out_path)
    return report


# ---------------------- 数据画像 & 受限计算执行器（数据对话两阶段管道） ----------------------
# 【nl-two-stage-pipeline 2026-10】
# 背景:「数据对话」旧实现是 df.head(20) 喂 LLM 直接猜答案,大表聚合(占比/求和)
#      严重不准 —— LLM 只看到前 20 行,却要对全表下结论。
# 新方案:ai_service 先用 build_sheet_profile 生成轻量画像喂给 LLM「做规划」,
#      LLM 只产出受限计算指令(ops 参数字典),再由 execute_computation_ops 用
#      Pandas 对**全量数据**真实执行,最后 LLM 基于真实计算结果写结构化叙述。
# 安全边界:ops 只是「参数字典」,执行器只有白名单算子,严禁 eval/exec/任意代码;
#      LLM 幻觉出的非法列名/算子会被跳过并记录,形成双重防线。

# filter 算子允许的比较符(白名单,防注入)
_NL_FILTER_OPERATORS = ("==", "!=", ">", ">=", "<", "<=", "contains")
# groupby_agg 算子允许的聚合函数
_NL_AGG_FUNCS = ("sum", "mean", "count", "max", "min")
# 数值列画像最多覆盖的列数(列很多时防 prompt 爆炸,多出的列只报类型/缺失率)
_NL_PROFILE_MAX_NUM_COLS = 20
# 分类列 top5 频次的基数上限(去重数超过视为高基数,只报去重数不报取值)
_NL_PROFILE_TOPN_MAX_CARDINALITY = 50
# 计算结果表最多返回的行数(控制接口响应与叙述 prompt 的体积)
_NL_RESULT_MAX_ROWS = 50


def _nl_plain_value(v: Any) -> Any:
    """把任意单元格值规整成可 JSON 序列化的纯 Python 类型。

    为什么单独抽出来:DataFrame.astype(object).values.tolist() 对 object 列里
    残留的 np.int64 / np.float64 / Timestamp 不会自动转原生类型,直接
    json.dumps 会抛 TypeError。逐值规整一次,保证结果表永远可序列化。
    """
    if v is None:
        return ""
    if isinstance(v, bool):  # bool 必须在 int 之前判断(bool 是 int 的子类)
        return v
    if isinstance(v, (int, float)):
        # float 特判 NaN/Inf:json 规范不允许这两个字面量,统一转空串
        if isinstance(v, float) and (v != v or v in (float("inf"), float("-inf"))):
            return ""
        return v
    if isinstance(v, str):
        return v
    # numpy 标量(np.int64/np.float64/np.bool_)→ .item() 拿原生数,再递归规整
    if hasattr(v, "item"):
        try:
            return _nl_plain_value(v.item())
        except Exception:
            pass
    # pandas Timestamp / datetime / date → ISO 字符串
    if hasattr(v, "isoformat"):
        try:
            return v.isoformat()
        except Exception:
            pass
    # 其余类型(Decimal 等)先试 pd.isna,是缺失值给空串,兜底转 str
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v)


def _nl_round_clean(v: float) -> float | int:
    """画像数值规整:整数化结果转 int(减小 prompt 体积),小数 round 4 位。"""
    if v != v:  # NaN 防御
        return 0
    if float(v).is_integer() and abs(v) < 1e15:
        return int(v)
    return round(float(v), 4)


def build_sheet_profile(df: pd.DataFrame) -> dict[str, Any]:
    """生成单 sheet 的轻量数据画像(全部字段为可 JSON 序列化的纯 Python 类型)。

    用途:「数据对话」两阶段管道的**阶段 1** —— 把画像(而不是 df.head)喂给 LLM,
    让模型知道每列的真实类型/缺失率/数值范围/分类取值,再去规划计算指令。

    为什么是"画像"而不是 head:
        - head(20) 只有前 20 行,LLM 据此答"全表占比"必然幻觉;
        - 画像里的 min/max/sum/去重值是**全量统计**,LLM 规划时不会选错列、编错值。

    体积控制:
        - 数值列统计只覆盖前 20 个数值列,超出的列只报类型与缺失率(stats_omitted);
        - 分类列 top5 频次仅对去重数 ≤ 50 的低基数列输出,高基数文本列只报去重数。
    """
    profile: dict[str, Any] = {
        "row_count": int(len(df)),
        "col_count": int(len(df.columns)),
        "columns": [],
    }
    # 空表直接返回骨架,避免后续统计对空 df 抛异常
    if len(df) == 0 or len(df.columns) == 0:
        return profile

    # 复用现有类型推断(date/number/text),保证画像与清洗/图表侧的类型口径一致
    col_types = detect_column_types(df)

    total_rows = int(len(df))
    num_profiled = 0  # 已输出统计的数值列计数(上限 _NL_PROFILE_MAX_NUM_COLS)
    for col in df.columns:
        col_key = str(col)
        ctype = col_types.get(col_key, "text")
        s = df[col]
        # 缺失率 = 缺失行数 / 总行数(0~1,保留 3 位小数)
        non_null = int(s.notna().sum())
        col_info: dict[str, Any] = {
            "name": col_key,
            "type": ctype,
            "missing_rate": round(1.0 - non_null / total_rows, 3) if total_rows else 0.0,
        }

        if ctype == "number":
            # 统一 to_numeric 兜底:文本里混入 "1,234"/"12%" 时 detect 已判 number,
            # 这里再 coerce 一次,保证个别脏值不会让统计抛异常
            num = pd.to_numeric(s, errors="coerce")
            valid = num.dropna()
            if len(valid) == 0:
                col_info["stats"] = None
            elif num_profiled < _NL_PROFILE_MAX_NUM_COLS:
                num_profiled += 1
                col_info["stats"] = {
                    "min": _nl_round_clean(float(valid.min())),
                    "max": _nl_round_clean(float(valid.max())),
                    "mean": _nl_round_clean(float(valid.mean())),
                    "sum": _nl_round_clean(float(valid.sum())),
                }
            else:
                # 超上限:只保留类型/缺失率,防止几十个数值列把 prompt 撑爆
                col_info["stats_omitted"] = True
        elif ctype == "date":
            # 日期列:解析后取 min/max 范围字符串,让 LLM 知道数据的时间跨度
            parsed = pd.to_datetime(s, errors="coerce")
            valid = parsed.dropna()
            if len(valid) > 0:
                col_info["range"] = {
                    "min": str(valid.min()),
                    "max": str(valid.max()),
                }
        else:
            # 分类/文本列:去重数必给;top5 频次仅低基数列输出(高基数列取值无叙述价值)
            distinct = int(s.nunique(dropna=True))
            col_info["distinct"] = distinct
            if 0 < distinct <= _NL_PROFILE_TOPN_MAX_CARDINALITY:
                vc = s.value_counts(dropna=True).head(5)
                col_info["top5"] = [
                    {"value": str(k)[:50], "count": int(v)}  # 值截断 50 字符,防超长文本
                    for k, v in vc.items()
                ]
        profile["columns"].append(col_info)
    return profile


def _nl_apply_operator(s: pd.Series, operator: str, val: Any) -> pd.Series:
    """按白名单比较符对 Series 做逐元素比较,返回布尔掩码。"""
    if operator == "==":
        return s == val
    if operator == "!=":
        return s != val
    if operator == ">":
        return s > val
    if operator == ">=":
        return s >= val
    if operator == "<":
        return s < val
    return s <= val


def _nl_build_compare_mask(s: pd.Series, operator: str, value: Any) -> pd.Series:
    """为 filter 算子构造布尔掩码(智能类型匹配)。

    设计要点:
        - 数值比较优先:LLM 传来的 value 常是字符串("100"),先尝试转 float;
          同时把目标列 to_numeric(转不动的变 NaN,比较结果为 False,天然排除);
        - 日期列:把 value 也 to_datetime 后再比较,支持 "2024-01" 这类写法;
        - 其余按字符串比较(astype(str),避免 object 列混型比较抛 TypeError)。
    """
    if pd.api.types.is_datetime64_any_dtype(s):
        val = pd.to_datetime(value, errors="coerce")
        if pd.isna(val):
            # 日期解析失败 → 全 False 掩码(条件不成立比误过滤更安全)
            return pd.Series(False, index=s.index)
        return _nl_apply_operator(s, operator, val)

    # 数值优先:value 可转数值 且 目标列至少有一部分能转数值 → 数值比较
    num = pd.to_numeric(s, errors="coerce")
    val = value
    if isinstance(val, str):
        try:
            val = float(val.replace(",", ""))  # "1,234" → 1234.0
        except ValueError:
            val = None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if num.notna().any():
            return _nl_apply_operator(num, operator, float(val))

    # 字符串比较兜底(==/!= 精确匹配,> < 走字典序)
    s_str = s.astype(str)
    val_str = "" if value is None else str(value)
    return _nl_apply_operator(s_str, operator, val_str)


def execute_computation_ops(df: pd.DataFrame, ops: list[dict]) -> dict[str, Any]:
    """白名单算子执行器:按顺序流水线执行 LLM 规划的计算指令(全量真实计算)。

    用途:「数据对话」两阶段管道的**阶段 3**。LLM 只产出 ops 参数字典,
    本函数用 Pandas 对全量 df 真实执行 —— 这就是"大表聚合不再靠猜"的根因:
    数字全部来自 Pandas 计算,LLM 只负责后续解读。

    安全边界:
        - 严禁 eval/exec/任意代码;只有 5 种白名单算子;
        - 未知算子忽略并 logger.warning;
        - 列名不存在 / 参数非法 → 该步跳过并记录,不中断流水线;
        - 任何单步异常被捕获记录到 error,继续尽量返回可序列化结果。

    ops 支持(数组顺序即流水线顺序:filter → groupby_agg → sort → topn):
        filter      : {"op":"filter","column":列名,"operator":"==|!=|>|>=|<|<=|contains","value":值}
        groupby_agg : {"op":"groupby_agg","group":[分组列],"agg":{数值列:"sum|mean|count|max|min"}}
        sort        : {"op":"sort","by":列名,"ascending":bool}
        topn        : {"op":"topn","by":数值列,"n":int,"ascending":bool(默认 False=最大前 N)}
        describe    : {"op":"describe"}   # 数值列统计摘要

    返回:{"steps":[每步执行说明与结果行数],
          "table":{"columns":[...], "rows":[[...]](最多 50 行)},
          "error":"成功为空串,否则为最后一条错误信息"}
    """
    steps: list[dict[str, Any]] = []
    error = ""
    work = df.copy()  # 拷贝执行,绝不污染调用方的原表

    ops_list = ops if isinstance(ops, list) else []
    for idx, op_item in enumerate(ops_list, start=1):
        # 非法项直接跳过(LLM 输出不完全可信,逐项防御)
        if not isinstance(op_item, dict):
            error = f"第 {idx} 步指令不是对象,已跳过"
            logger.warning(f"[execute_computation_ops] {error}")
            steps.append({"step": idx, "op": "", "desc": error, "rows": int(len(work))})
            continue
        op = str(op_item.get("op") or "").strip().lower()
        try:
            # —— 算子 1:filter 布尔过滤 ——
            if op == "filter":
                col = str(op_item.get("column") or "")
                operator = str(op_item.get("operator") or "").strip()
                value = op_item.get("value")
                if col not in work.columns:
                    # 列名是 LLM 编的 → 跳过该步(记录进 error 供叙述侧如实说明)
                    error = f"filter 跳过:列「{col}」不存在"
                    logger.warning(f"[execute_computation_ops] {error}")
                    steps.append({"step": idx, "op": op, "desc": error, "rows": int(len(work))})
                    continue
                if operator not in _NL_FILTER_OPERATORS:
                    error = f"filter 跳过:比较符「{operator}」不在白名单"
                    logger.warning(f"[execute_computation_ops] {error}")
                    steps.append({"step": idx, "op": op, "desc": error, "rows": int(len(work))})
                    continue
                s = work[col]
                if operator == "contains":
                    # contains 走字符串包含:关正则(把值当普通文本,防正则注入)
                    mask = s.astype(str).str.contains(str(value), case=False, regex=False, na=False)
                else:
                    mask = _nl_build_compare_mask(s, operator, value)
                before = int(len(work))
                work = work.loc[mask.fillna(False)].reset_index(drop=True)
                desc = f"filter:{col} {operator} {value!r},由 {before} 行过滤为 {len(work)} 行"
                steps.append({"step": idx, "op": op, "desc": desc, "rows": int(len(work))})

            # —— 算子 2:groupby_agg 分组聚合 ——
            elif op == "groupby_agg":
                group_raw = op_item.get("group")
                # group 兼容单字符串(LLM 可能传 "产品" 而不是 ["产品"])
                if isinstance(group_raw, str):
                    group_raw = [group_raw]
                group_cols = [str(g) for g in (group_raw or []) if str(g) in work.columns]
                agg_raw = op_item.get("agg") if isinstance(op_item.get("agg"), dict) else {}
                # 只保留「聚合列存在 + 聚合函数在白名单」的项,双重过滤防 LLM 幻觉
                agg_map: dict[str, str] = {}
                for c, f in agg_raw.items():
                    if str(c) in work.columns and str(f).lower() in _NL_AGG_FUNCS:
                        agg_map[str(c)] = str(f).lower()
                if not group_cols or not agg_map:
                    error = "groupby_agg 跳过:分组列或聚合列均不存在/不在白名单"
                    logger.warning(f"[execute_computation_ops] {error}")
                    steps.append({"step": idx, "op": op, "desc": error, "rows": int(len(work))})
                    continue
                # 聚合前把非数值 dtype 的聚合列 coerce 成数值(count 除外),
                # 防止 Excel 里存成文本的数字("1200")被 sum 漏加
                tmp = work.copy()
                for c in agg_map:
                    if agg_map[c] != "count" and not pd.api.types.is_numeric_dtype(tmp[c]):
                        tmp[c] = pd.to_numeric(tmp[c], errors="coerce")
                # 【2026-10 修复】聚合列与分组列同名时(group=["往来单位名称"],
                # agg={"往来单位名称":"count"} 是 LLM 极常见的写法),直接 reset_index()
                # 会抛 "cannot insert 往来单位名称, already exists" 导致该步白跑。
                # 处理:先把聚合结果里与分组列重名的输出列加函数后缀,再 reset_index。
                agg_result = tmp.groupby(group_cols, dropna=False).agg(agg_map)
                if not isinstance(agg_result.columns, pd.MultiIndex):
                    agg_result.columns = [
                        f"{c}_{agg_map.get(str(c), 'agg')}" if str(c) in group_cols else str(c)
                        for c in agg_result.columns
                    ]
                grouped = agg_result.reset_index()
                # 兜底拍平可能的 MultiIndex 列名("a_sum" 形式),保证 JSON 结构简单
                grouped.columns = [
                    "_".join(map(str, t)) if isinstance(t, tuple) else str(t)
                    for t in grouped.columns
                ]
                work = grouped
                desc = f"groupby_agg:按 {group_cols} 聚合 {agg_map} → {len(work)} 行"
                steps.append({"step": idx, "op": op, "desc": desc, "rows": int(len(work))})

            # —— 算子 3:sort 排序 ——
            elif op == "sort":
                by = str(op_item.get("by") or "")
                ascending = bool(op_item.get("ascending", True))
                if by not in work.columns:
                    error = f"sort 跳过:列「{by}」不存在"
                    logger.warning(f"[execute_computation_ops] {error}")
                    steps.append({"step": idx, "op": op, "desc": error, "rows": int(len(work))})
                    continue
                if pd.api.types.is_numeric_dtype(work[by]):
                    work = work.sort_values(
                        by=by, ascending=ascending, na_position="last"
                    ).reset_index(drop=True)
                else:
                    # 文本列尽量按数值排序:≥90% 可转数值时用临时键列排,更符合直觉
                    num = pd.to_numeric(work[by], errors="coerce")
                    if num.notna().mean() >= 0.9:
                        work = (
                            work.assign(_nl_sort_key=num)
                            .sort_values(by="_nl_sort_key", ascending=ascending, na_position="last")
                            .drop(columns=["_nl_sort_key"])
                            .reset_index(drop=True)
                        )
                    else:
                        work = work.sort_values(
                            by=by, ascending=ascending, na_position="last"
                        ).reset_index(drop=True)
                desc = f"sort:按 {by} {'升序' if ascending else '降序'}排列,共 {len(work)} 行"
                steps.append({"step": idx, "op": op, "desc": desc, "rows": int(len(work))})

            # —— 算子 4:topn 取前 N ——
            elif op == "topn":
                by = str(op_item.get("by") or "")
                try:
                    # n 夹在 [1, 50] 区间,防 LLM 传 99999 把结果表撑爆
                    n = max(1, min(int(op_item.get("n", 5)), _NL_RESULT_MAX_ROWS))
                except (TypeError, ValueError):
                    n = 5
                # 默认 ascending=False = 取最大前 N(常见"前几名"语义)
                ascending = bool(op_item.get("ascending", False))
                if by not in work.columns:
                    error = f"topn 跳过:列「{by}」不存在"
                    logger.warning(f"[execute_computation_ops] {error}")
                    steps.append({"step": idx, "op": op, "desc": error, "rows": int(len(work))})
                    continue
                # 按数值键排序取前 N 行(保留原行全部列);文本数字 coerce 后排
                key = (
                    work[by]
                    if pd.api.types.is_numeric_dtype(work[by])
                    else pd.to_numeric(work[by], errors="coerce")
                )
                order = key.sort_values(ascending=ascending, na_position="last").index[:n]
                work = work.loc[order].reset_index(drop=True)
                desc = f"topn:按 {by} 取{'最小' if ascending else '最大'}前 {n} 行 → {len(work)} 行"
                steps.append({"step": idx, "op": op, "desc": desc, "rows": int(len(work))})

            # —— 算子 5:describe 数值列统计摘要 ——
            elif op == "describe":
                desc_df = work.describe(numeric_only=True)
                if desc_df.empty:
                    error = "describe 跳过:无可用数值列"
                    logger.warning(f"[execute_computation_ops] {error}")
                    steps.append({"step": idx, "op": op, "desc": error, "rows": int(len(work))})
                    continue
                # 转置成「统计量 × 数值列」的表(每行一列的摘要),比原始 describe 好叙述
                t = desc_df.T.reset_index().rename(columns={"index": "column"})
                for c in t.columns:
                    if c != "column":
                        t[c] = pd.to_numeric(t[c], errors="coerce").round(4)
                work = t
                desc = f"describe:输出 {t.shape[1] - 1} 个数值列的 8 项统计摘要"
                steps.append({"step": idx, "op": op, "desc": desc, "rows": int(len(work))})

            # —— 未知算子:忽略并告警,不中断流水线 ——
            else:
                logger.warning(f"[execute_computation_ops] 第 {idx} 步未知算子「{op}」,已忽略")
                steps.append({
                    "step": idx, "op": op,
                    "desc": f"未知算子「{op}」,已忽略", "rows": int(len(work)),
                })
                continue
        except Exception as exc:
            # 单步异常不中断:记录错误,保留当前 work 继续后续算子
            error = f"第 {idx} 步({op})执行失败:{exc}"
            logger.warning(f"[execute_computation_ops] {error}")
            steps.append({"step": idx, "op": op, "desc": error, "rows": int(len(work))})

    # —— 结果表序列化:最多 50 行;NaN→""、numpy/Timestamp→原生类型 ——
    out = work.head(_NL_RESULT_MAX_ROWS)
    # NaN 处理思路(任务指定):astype(object).where(notna, "") 把缺失统一替换为空串
    out = out.astype(object).where(pd.notna(out), "")
    columns = [str(c) for c in out.columns]
    rows = [[_nl_plain_value(v) for v in row] for row in out.itertuples(index=False, name=None)]
    return {"steps": steps, "table": {"columns": columns, "rows": rows}, "error": error}
