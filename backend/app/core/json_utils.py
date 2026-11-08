"""跨模块复用的 JSON 序列化工具 + LLM 输出清洗工具。

为什么需要它：
    后端很多数据来源于 ``pandas``（``compute_indicators``、``read_dataframe``
    出来的 ``DataFrame``、``Timestamp``、``numpy.int64``、``numpy.ndarray`` 等），
    而标准库 ``json.dumps`` 不认识这些类型，会抛
    ``TypeError: Object of type Timestamp is not JSON serializable``。

    这个文件提供两类工具：
        1. ``json_default`` / ``dumps`` —— 兜底序列化器，统一处理 pandas/numpy 类型。
        2. ``strip_thinking`` / ``strip_english_reasoning`` / ``clean_llm_text``
           / ``clean_llm_obj`` —— 剥离 LLM 输出里的"思考过程"。

任何需要把含 pandas / numpy 类型的对象序列化为 JSON、或者把 LLM 输出文本
展示给最终用户的地方都应使用这里，避免在不同模块里重复实现但实现不全。
"""
from __future__ import annotations

import json
import re
from typing import Any


# ============================================================
# 1) pandas / numpy 兼容的 JSON 序列化
# ============================================================

def json_default(obj: Any) -> Any:
    """``json.dumps(..., default=json_default)`` 的兜底序列化器。

    优先级：
        1. 带 ``isoformat`` 的对象（``pd.Timestamp`` / ``datetime`` 等）→ ISO 字符串
        2. 带 ``item`` 的 numpy 标量（``np.int64`` / ``np.float64`` 等）→ Python 原生数
        3. 带 ``tolist`` 的 numpy 数组 → Python 列表
        4. 兜底 → ``str(obj)``
    """
    # 1) pandas Timestamp / datetime.datetime / datetime.date：转 ISO 字符串
    if hasattr(obj, "isoformat"):
        try:
            return obj.isoformat()
        except Exception:
            pass
    # 2) numpy 标量（np.int64 / np.float64 等）：.item() 拿到 Python 原生数
    if hasattr(obj, "item"):
        try:
            return obj.item()
        except Exception:
            pass
    # 3) numpy ndarray：递归转成 Python 列表
    if hasattr(obj, "tolist"):
        try:
            return obj.tolist()
        except Exception:
            pass
    # 4) 其余一律字符串化，避免再次抛错
    return str(obj)


def dumps(obj: Any, **kwargs: Any) -> str:
    """``json.dumps`` 的薄封装，自动注入 ``default=json_default``。"""
    kwargs.setdefault("ensure_ascii", False)
    kwargs.setdefault("default", json_default)
    return json.dumps(obj, **kwargs)


# ============================================================
# 1.5) LLM 输出宽容解析（解析 LLM 返回的"非完全合规 JSON"）
# ============================================================
# 现象：LLM 即使被要求输出 JSON,也会出现这些小毛病:
#   - 属性名没加双引号:`{type: "bar"}`（Python 风格）
#   - 用了单引号:`{'type': 'bar'}`（Python 风格）
#   - 末尾多了 `...` 或代码块 ```` ```json ```` ```` ``` ```` 包裹
#   - 多了尾巴如 `<|end|>`、`<|im_end|>` 等 thinking 截断标记
#   - 多了一层 reasoning 思考段
# 标准 json.loads 对这些全部抛错,体验差。
# 这个解析器按下列顺序尝试,每一层都尽量不破坏原意:
#   1) 抽第一个 `{...}` JSON 块;
#   2) 直接 json.loads;
#   3) 替换单引号为双引号(只在外层);
#   4) 把 `{ key:` 修复成 `{"key":`(给裸 key 加双引号);
#   5) 都失败 → 抛 JSONDecodeError。
def _safe_json_loads(text: str) -> Any:
    """宽容解析 LLM 返回的"看起来像 JSON"的字符串。

    Returns:
        解析后的 Python 对象(常见为 dict / list)。
    Raises:
        json.JSONDecodeError: 所有尝试都失败时。

    【2026-10 关键修复】不再用 ``re.search(r"\{[\s\S]*\}", text)`` 抽最外层块——
    旧实现会把 markdown 报告(含表格里的 `{...}`、正文里的中英文括号对)
    贪婪匹配成"超大 JSON",然后 ``json.loads`` 必然失败,最终落到
    ``chat_json`` 的 fallback ``{"title": "AI 分析", "sections": {"overview": raw[:500]}}``,
    报告页就只剩"本期数据未覆盖"占位文案。
    修法:用 ``json.JSONDecoder.raw_decode`` 从左到右扫描,逐个尝试,
    只接受**真正可解析**的最长 JSON 块;扫不到再降级到 markdown 包裹剥离 +
    修复层。优先尝试 markdown ```json``` ``` 包裹的内容。
    """
    if not text or not isinstance(text, str):
        raise json.JSONDecodeError("empty input", text or "", 0)

    decoder = json.JSONDecoder()

    # 第 0 步(优先):如果文本含 ```json ... ``` 代码块,直接取块内内容解析
    md_blocks: list[str] = []
    for m in re.finditer(r"```(?:json)?\s*([\s\S]*?)\s*```", text):
        md_blocks.append(m.group(1))
    for blk in md_blocks:
        for variant in (blk, blk.strip()):
            try:
                obj, _end = decoder.raw_decode(variant)
                if obj is not None:
                    return obj
            except json.JSONDecodeError:
                pass

    # 第 1 步:用 raw_decode 直接从头扫,只取第一个合法 JSON,自动忽略尾巴杂质。
    for cand in (text, text.strip(), _strip_markdown(text)):
        try:
            obj, _end = decoder.raw_decode(cand)
            if obj is not None:
                return obj
        except json.JSONDecodeError:
            pass

    # 第 2 步:从左到右扫描所有以 `{` 或 `[` 起始的位置,逐一尝试 raw_decode。
    # 这样即使 LLM 在前面写了 6KB 的中文 markdown,我们也能跳到**真正**的
    # JSON 起点(通常是尾部 chart_option 的 `{`)开始解析。
    best_obj: Any = None
    best_end = -1
    for m in re.finditer(r"[\{\[]", text):
        start = m.start()
        try:
            obj, end = decoder.raw_decode(text[start:])
            # 取解析位置最靠后的合法 JSON(即"最靠近文本末尾"的那个,
            # 通常是 LLM 把 chart_option 放在最后的形态)
            if end > best_end and obj is not None:
                best_obj = obj
                best_end = end
        except json.JSONDecodeError:
            continue
    if best_obj is not None:
        return best_obj

    # 第 3 步(最后兜底):对整段文本做层层修复,逐个尝试 raw_decode。
    # 已无 markdown 块、无任何合法 JSON 起点,说明 LLM 输出完全跑偏,
    # 这时直接 raise 让 chat_json 走 fallback。
    last_err: Exception | None = None
    for cand in (text, text.strip(), _strip_markdown(text)):
        for transformed in _iter_json_repairs(cand):
            try:
                obj, _end = decoder.raw_decode(transformed)
                if obj is not None:
                    return obj
            except json.JSONDecodeError as exc:
                last_err = exc
                continue

    if last_err:
        raise last_err
    raise json.JSONDecodeError("no parseable json", text, 0)


def _strip_markdown(text: str) -> str:
    """去掉 LLM 常见的 markdown 代码块包裹 / thinking 截断尾巴。

    这是一道"清洗层":只去除噪音,不动 JSON 内部。
    """
    s = text.strip()
    if s.startswith("```"):
        s = re.sub(r"^```(?:json)?\s*", "", s)
        s = re.sub(r"\s*```\s*$", "", s)
    # 去掉尾部 thinking 截断标记(<|im_end|> / <|end|> / ...)
    s = re.sub(r"<\|[^|>]*\|>\s*$", "", s).strip()
    return s


def _iter_json_repairs(text: str):
    """从最严格到最宽松,逐层产出"修复后"的字符串,给 _safe_json_loads 逐个尝试。

    Yields:
        字符串(后续 json.loads 失败就跳到下一个候选)。
    """
    s = text.strip()
    # 去掉 markdown 代码块包裹
    if s.startswith("```"):
        # 去掉首尾 ```...```
        s = re.sub(r"^```(?:json)?\s*", "", s)
        s = re.sub(r"\s*```\s*$", "", s)
    # 去掉尾部 thinking 截断标记(<|im_end|> / <|end|> / ...)
    s = re.sub(r"<\|[^|>]*\|>\s*$", "", s).strip()

    # 1) 原样(可能含未加引号属性名)
    yield s

    # 2) 给裸 key 加引号 ——
    #   - 匹配 `{` 或 `,` 或空白后,跟着 ASCII 字母/下划线/中文(裸 key),后跟 `:` 的位置
    #   - 用 `re.sub` 把裸 key 包成双引号
    #   - 仅这一段是最保守的修复:它不会影响已经加引号的字符串里的冒号
    fixed = re.sub(
        r'([\{\,\s])([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]*)\s*:',
        r'\1"\2":',
        s,
    )
    if fixed != s:
        yield fixed

    # 3) 单引号 → 双引号(仅一次)
    if "'" in s:
        yield s.replace("'", '"')

    # 4) 修复后 + 单引号双重
    if "'" in s:
        fixed2 = re.sub(
            r'([\{\,\s])([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]*)\s*:',
            r'\1"\2":',
            s.replace("'", '"'),
        )
        yield fixed2


# ============================================================
# 2) LLM 输出清洗（剥离思考过程，避免泄漏到最终展示）
# ============================================================
# 一些模型（MiniMax-M2.7 / 主流 reasoning 模型）会输出两类"不该让用户看到"的内容：
#
#   A. 包裹在标签内的思考块：``<think>...</think>`` / ``<thinks>...</thinks>`` 等
#      → 由 ``strip_thinking()`` 处理
#
#   B. 直接写在外层正文里的英文推导段（无标签包裹）。最常见的形式是
#      reasoning 模型先写一段"分析过程"再回到中文正文：
#      ```
#      Let me carefully analyze the data provided. Looking at the indicator data:
#      - 36 rows total
#      - 25 columns
#      - Customers: 福州朴朴 (13), 广州朴朴 (9), 成都朴朴 (7) ...
#      正式结论：本月共 36 条出货记录...
#      ```
#      → 由 ``strip_english_reasoning()`` 处理
#
# 这两类问题本质互补：清洗链路应同时调用两者（详见 ``clean_llm_text``）。

# ------------------------------------------------------------
# 2.1) ``<think>`` 类标签剥离
# ------------------------------------------------------------

# 思考块标签白名单（支持大小写、空格容忍、嵌套/同名变体）
_THINK_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"<think>[\s\S]*?</think>", re.IGNORECASE),
    re.compile(r"<thinks?>[\s\S]*?</thinks?>", re.IGNORECASE),
    re.compile(r"<thinking>[\s\S]*?</thinking>", re.IGNORECASE),
    re.compile(r"<reasoning>[\s\S]*?</reasoning>", re.IGNORECASE),
)


def strip_thinking(text: str) -> str:
    """剥离 LLM 返回文本里的思考块，返回剩余正文。

    处理流程：
        1. 依次应用上述正则，把所有 ``<think>...</think>`` 等标签段删除。
        2. 二次清洗：剥掉可能"成对标签不完整"导致的裸起止标签
           （常见于嵌套或被截断的输出，例如 ``</think> tail`` 残留）。
        3. 折叠 3 个及以上连续换行为最多 2 个，避免出现大段空白。
        4. 去掉首尾空白后返回。如果剥离后为空，返回原文（防止误伤
           真的只输出了 ``<think>`` 内容却没正文的极端模型）。

    幂等性：对已剥离的文本再次调用不会有副作用。
    """
    if not text:
        return text or ""

    cleaned = text
    for pat in _THINK_PATTERNS:
        cleaned = pat.sub("", cleaned)

    # 二次清洗：剥离残留的裸起止标签（嵌套/未配对场景）
    cleaned = re.sub(r"</?(?:think|thinks|thinking|reasoning)\s*/?>", "", cleaned, flags=re.IGNORECASE)

    # 折叠多余空行
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned).strip()

    # 兜底：剥离后若空了，保留原文（防止某些模型把所有内容都放在 <think> 里）
    return cleaned if cleaned else text.strip()


# ------------------------------------------------------------
# 2.2) 英文推导段剥离（处理"明文写在正文里"的思考过程外泄）
# ------------------------------------------------------------

# 段首指纹：段首前 60 字符内出现以下任意一条，即视作整段为模型思考泄漏
_REASONING_PREFIX_PATTERNS: tuple[re.Pattern[str], ...] = (
    # "Let me ..." 系列（最常见，放最前）
    re.compile(r"^\s*Let me\s+carefully\s+analy[sz]e\b", re.IGNORECASE),
    re.compile(r"^\s*Let me\s+(?:carefully\s+)?analy[sz]e\b", re.IGNORECASE),
    re.compile(r"^\s*Let me\s+(?:carefully\s+)?(?:check|think|examine|review|see|look|start|begin|break|walk|go)\b", re.IGNORECASE),
    re.compile(r"^\s*Let me\b", re.IGNORECASE),
    # "Looking at ..." 系列
    re.compile(r"^\s*Looking\s+at\b", re.IGNORECASE),
    re.compile(r"^\s*Looking\s+into\b", re.IGNORECASE),
    # OK / Okay / So / Now / Alright 等衔接词开头
    re.compile(r"^\s*(?:OK|Okay)[,.\s]+(?:so|let|now|I|let'?s)\b", re.IGNORECASE),
    re.compile(r"^\s*So[,.]?\s+(?:the|this|we|I|let|looking)\b", re.IGNORECASE),
    re.compile(r"^\s*Now[,.]?\s+(?:the|this|let|I|we|looking)\b", re.IGNORECASE),
    re.compile(r"^\s*Alright[,.\s]+\w", re.IGNORECASE),
    # First / Second / Third / Finally / Then 等 step-like 编号
    re.compile(r"^\s*(?:First|Second|Third|Fourth|Fifth)[,.]?\s+\w", re.IGNORECASE),
    re.compile(r"^\s*Finally[,.]?\s", re.IGNORECASE),
    re.compile(r"^\s*Then[,.]?\s+(?:the|this|I|we|looking)\b", re.IGNORECASE),
    # Note that / This is / Here's / Here is / Step N:
    re.compile(r"^\s*Note\s+that\b", re.IGNORECASE),
    re.compile(r"^\s*This\s+(?:is|means|suggests|implies|indicates)\b", re.IGNORECASE),
    re.compile(r"^\s*Here(?:'s| is)\b", re.IGNORECASE),
    re.compile(r"^\s*Step\s+\d+\s*[:.]", re.IGNORECASE),
)

# 行内指纹：段首 60 字符内出现以下模板，也算"被污染"。
# 覆盖"Let me analyze ... Looking at ..."连写在同一段内的场景。
_REASONING_INLINE_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"Let me\s+carefully\s+analy[sz]e", re.IGNORECASE),
    re.compile(r"Let me\s+(?:analyze|check|think|examine|review)", re.IGNORECASE),
    re.compile(r"Looking at (?:the |this )?(?:indicator|data|column|table|chart|summary|numeric)", re.IGNORECASE),
    # 极特征明显的"内部数据结构引用"，是 reasoning 模型外泄的强信号
    re.compile(r"\bnumeric_summary\b.*\bby_group\b", re.IGNORECASE),
    re.compile(r"\b(?:by_group|topn|trend)\b", re.IGNORECASE),
)

# 中文字符正则
_CN_RE = re.compile(r"[\u4e00-\u9fff]")

# 一行若完全没有中文
_NO_CN_LINE_RE = re.compile(r"^[^\u4e00-\u9fff]*$")

# "句末标点 + 间隔 + 中文段" —— 启发式定位"中文正文"的起始点
# 句末标点可以是中英文的 .。!?！？;；
_CN_SENTENCE_START_RE = re.compile(
    r"[。.!?！？;；]\s*"
    r"(?="                                # 向前预查,真正的中文起始位置在后面
    r"[\u4e00-\u9fff]"                    # 第一个字符必须是中文
    r"(?:[^.。!?！？;；\n]*[\u4e00-\u9fff]){2,}"  # 至少还有 2 个中文（保证是真正的中文段落）
    r")"
)

# 短段落豁免阈值：长度 < 该阈值的段落不做"逐行剥英文"等清洗，避免误伤
# 单段章节标题 / 标签（如 "overview" / "highlights" / "Q1 完成率 95%" 等）
_SHORT_PARAGRAPH_GUARD = 40


def _is_pure_english_paragraph(s: str) -> bool:
    """判断一段文本是不是"基本全是英文"。

    用作英文推导段清洗的第 3 条兜底：某些模型整段写英文思考，
    既不在 ``<think>`` 块内、也碰巧不含指纹短语，但仍是推导过程，
    我们一并视作污染。

    判定标准（同时满足以下所有条件即返回 True）：
        - 长度 >= 30
        - 中文字符数量 == 0
        - 字母占总字符比例 > 50%
    """
    if not s or len(s) < 30:
        return False
    chinese = sum(1 for c in s if '\u4e00' <= c <= '\u9fff')
    if chinese > 0:
        return False
    letters = sum(1 for c in s if c.isalpha())
    return letters >= 20 and letters / len(s) > 0.5


def _drop_english_only_lines(text: str) -> str:
    """逐行清理"完全没有中文"的英文行（含纯空白行原样保留分隔用途）。"""
    out: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            out.append(line)
            continue
        if _NO_CN_LINE_RE.match(stripped):
            continue
        out.append(line)
    return "\n".join(out)


def _slice_to_chinese_conclusion(s: str) -> str | None:
    """尝试从一段英文推导段里"挖出"中文正文。

    思路：
        段首是英文污染（命中指纹），整段剥光会把"中文正文"也丢。
        但中文正文往往是"独立句子"——前面有句末标点（.。!?；），
        后面是连续的 ≥3 个中文字符。
        所以在段中找"句末+中文段"的边界位置，截到那之后作为保留内容。
    """
    m = _CN_SENTENCE_START_RE.search(s)
    if not m:
        return None
    start = m.end()
    rest = s[start:].strip()
    # 跳过残留的列表标记（如 "：..."、"1、..." 等）
    rest = re.sub(
        r"^\s*(?:[-*•·+]|[\(（]\s*\d{1,3}\s*[\)）]?|[一二三四五六七八九十0-9]+[、.])"
        r"\s*[\.\)、:：\-—]*\s*",
        "", rest, count=1,
    )
    return rest.strip() or None


def strip_english_reasoning(text: str) -> str:
    """剥离 LLM 输出里的"英文思考段"。

    适用场景：
        ``<think>...</think>`` 之外的英文推导段落，常见于 reasoning 模型
        （DeepSeek-R1 / QwQ / MiniMax-M2.7 等）在 sections 正文中
        写入的"分析过程外泄"。

    清洗策略（多层处理，确保不破坏中文正文）：

        第 1 步 —— 段级指纹命中：
            按空行切段，对**段首 60 字内**匹配段首/行内指纹的段：
                - 尝试从段中切出"以中文为主体的、独立的句子"作为保留内容
                  （启发式：找"句末+中文段起始"的位置，从那里截断丢弃前段）。
                - **若切不出中文结论，整段直接丢弃，不保留原文**（用户的
                  期望是"不要让英文推导出现"，不是"宁可漏报也要原文"）。

        第 2 步 —— 逐行兜底：
            即使段级指纹未命中，reasoning 模型仍可能在某行中段插入英文
            思考（典型："本月销售 56 笔<br/>So this is a strange pattern ..."）
            此时按行扫描，对"整行无中文"的行一律剥离。
            **< 40 字符的短段落不走这一步**，避免误伤章节标题等短文本。

        第 3 步 —— 整段兜底：
            对"长度 >=30、无任何中文、英文占比 > 50%"的整段直接丢弃。
            **短段落同样豁免**。

        第 4 步 —— 折叠 3+ 空行为 2 空行。

    幂等性：已清洗的文本二次调用没有任何副作用（段首不再命中、行已清空）。

    注意：与 :func:`strip_thinking` 互补，建议**先 strip_thinking 再本函数**。
    """
    if not text:
        return text or ""

    kept: list[str] = []
    for para in re.split(r"\n\s*\n", text):
        p = para.strip()
        if not p:
            continue
        head = p[:60]
        # 第 1 步：段级指纹命中（不受短段落豁免保护 —— 这里就是模型思考的强信号）
        if (any(pat.search(head) for pat in _REASONING_PREFIX_PATTERNS)
                or any(pat.search(head) for pat in _REASONING_INLINE_PATTERNS)):
            rest = _slice_to_chinese_conclusion(p)
            if rest:
                kept.append(rest)
            # 找不到中文结论 → 整段直接丢弃（不保留原文 —— 用户的语义期望）
            continue
        # 第 3 步：整段纯英文（无指纹但仍是推导）。
        # 短段（< _SHORT_PARAGRAPH_GUARD 字符）豁免——通常是 "overview"
        # "highlights" 之类章节名,或被人为精简的句子,不应被剥空。
        if len(p) >= _SHORT_PARAGRAPH_GUARD and _is_pure_english_paragraph(p):
            continue
        # 第 2 步：未命中段级指纹，但段内可能混入英文行；同样做长度保护。
        if len(p) >= _SHORT_PARAGRAPH_GUARD:
            cleaned_p = _drop_english_only_lines(p).strip()
        else:
            cleaned_p = p
        if cleaned_p:
            kept.append(cleaned_p)

    joined = "\n\n".join(kept).strip()
    joined = re.sub(r"\n{3,}", "\n\n", joined)
    # 注意：本函数不像 strip_thinking 那样"剥空则保留原文"
    # —— 英文推导段的兜底就是不要给用户看，整段丢光也无所谓
    return joined


# ------------------------------------------------------------
# 2.3) 复合入口：一次完成两步清洗；递归清洗嵌套结构
# ------------------------------------------------------------

def clean_llm_text(text: str) -> str:
    """一次调用完成"``<think>`` 块剥离 + 英文推导段清洗"两步。

    任何对外展示给用户的 LLM 文本（报告章节、摘要、批注、解析文字等），
    在写入持久化字段或导出文件**之前**都应过这一函数，避免：
        1. 模型在 ``<think>`` 块内的思考过程泄漏到 Word/PDF/MD/Excel；
        2. reasoning 模型在正文里写的"Let me analyze ... Looking at ..."外泄；
        3. 连续多个空行造成排版难看。

    Args:
        text: 原始 LLM 输出文本。

    Returns:
        清洗后的用户可读正文。如果原文被判定为"全是思考过程"，可能为空串。
    """
    if not text:
        return text or ""
    return strip_english_reasoning(strip_thinking(text))


def clean_llm_obj(obj: Any) -> Any:
    """对 dict / list / str 递归调用 :func:`clean_llm_text`。

    用于把整个 report 字典（title + sections + chart_option + summary + ...）
    一次性清洗，任何嵌套的字符串值都会被剥 ``<think>`` + 剥英文推导段。
    非字符串类型（dict / list / 数字 / None）原样保留。
    """
    if isinstance(obj, dict):
        return {k: clean_llm_obj(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [clean_llm_obj(v) for v in obj]
    if isinstance(obj, str):
        return clean_llm_text(obj)
    return obj
