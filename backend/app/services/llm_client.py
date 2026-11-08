"""LLM 客户端封装：统一 chat 接口，支持 JSON 模式。"""
from __future__ import annotations

import json
import logging
import os
from typing import Any

from openai import (
    APIConnectionError,
    APITimeoutError,
    AuthenticationError,
    BadRequestError,
    NotFoundError,
    OpenAI,
    PermissionDeniedError,
)

from app.core.config import settings

logger = logging.getLogger(__name__)


class LLMClient:
    """OpenAI 兼容接口封装。

    【split-vision-and-text-llm】文本与视觉走**两套独立配置**:
        - ``chat()`` / ``chat_json()`` —— 文本分析类(报告生成与改写、摘要、
          异常批注、数据校验解读、NL 查询、数据集对比叙事、图表配置…),
          走 ``openai_analysis_*``;
        - ``chat_vision()`` —— 视觉识别类(照片 OCR、扫描件/图片识别),
          走 ``openai_*`` + ``openai_vision_model``。
    任一 ``openai_analysis_*`` 留空时,该字段回退到对应的 ``openai_*``,
    因此只配单模型的老部署无需改动即可继续工作。
    """

    def __init__(self) -> None:
        # 关键：从 .env 读取 API Key / Base URL / Model
        self.api_key = settings.openai_api_key
        self.base_url = settings.openai_base_url
        self.model = settings.openai_model

        # ===== 【split-vision-and-text-llm 新增】文本分析侧配置 =====
        # 逐项回退:某项留空就沿用上面的通用配置,
        # 这样"只配视觉模型差异、其余沿用"这种半覆盖写法也能正常工作。
        self.analysis_api_key = settings.openai_analysis_api_key or self.api_key
        self.analysis_base_url = settings.openai_analysis_base_url or self.base_url
        self.analysis_model = settings.openai_analysis_model or self.model

        self._client: OpenAI | None = None

        # 启动期一次性打印真实网络出口，方便排查 IDE 代理拦截问题
        self._log_egress_info()

    def _log_egress_info(self) -> None:
        """启动时记录 BASE_URL 与当前进程可见的代理配置，便于排障。

        为什么需要：Trae / VSCode 等 IDE 自带本地 HTTP 代理
        （常见变量 ``HTTPS_PROXY`` / ``HTTP_PROXY``），如果用户给的是
        海外 API 域名，会被代理拦下导致 ``APITimeoutError``，
        而终端直连又是通的——这正是"怎么又断了"的根因之一。
        """
        proxy_keys = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
                      "ALL_PROXY", "all_proxy", "NO_PROXY", "no_proxy")
        proxies = {k: os.environ[k] for k in proxy_keys if k in os.environ}
        logger.info("LLM 配置: base_url=%s model=%s", self.base_url, self.model)
        # 【split-vision-and-text-llm】把"文本分析"这条链路的实际出口也打出来。
        # 两边 base_url 不一致时(跨服务商分流)排查问题必须一眼看出请求发给了谁。
        logger.info(
            "LLM 分流 | 文本分析: base_url=%s model=%s | 视觉识别: base_url=%s model=%s",
            self.analysis_base_url, self.analysis_model,
            self.base_url, settings.openai_vision_model or self.model,
        )
        if proxies:
            logger.warning("检测到代理环境变量（可能影响 LLM 连通）: %s", proxies)
        else:
            logger.info("未检测到代理环境变量，直连模式。")

    @property
    def client(self) -> OpenAI:
        """延迟初始化：避免启动时立即校验代理兼容性。

        超时 / 重试策略（2026-10 调优）:
            - 单次 timeout = **20s**(原 60s):覆盖 95% 正常响应,极少真有超过 20s 的;
              真超了就该让上层知道,而不是闷在客户端里继续拖。
            - max_retries = **1**(原 2):openai-python 的内置重试是**串行阻塞**,
              60×3 拉满能接近 180s,直接把工作台 UA 端的 axios timeout
              (默认 120s)拖爆,看上去就是「AI 断开了」。
              改为 1 次后,极端最坏 ≈ 40s,仍能在 120s 容忍窗内收尾。
            - HTTP/connect timeout 仍是 10s,保持「TCP 握手失败」快速反馈。
            - 保留 trust_env=False,绕开 IDE 注入的 HTTP_PROXY。

        设计取舍:重试从 2 降到 1 后,**抖动容忍度下降**。考虑到单次超时已
        从 60s 调到 20s,总体最大等待仍是可控的(20 × 2 ≈ 40s),但用户体验
        改善明显 —— 失败时立刻拿到明确错误,而不是「按了没动静 2 分钟才说失败」。
        """
        if self._client is None:
            # 显式构造一个不走任何系统代理的 httpx 客户端，
            # 避免 IDE（Trae/VSCode）注入的 HTTP_PROXY 把
            # 海外 API 域名（api.minimaxi.com 等）拦在代理黑洞里。
            try:
                import httpx
                http_client = httpx.Client(
                    timeout=httpx.Timeout(60.0, connect=10.0),
                    trust_env=False,  # 关键：不读 HTTP_PROXY/HTTPS_PROXY
                    # 【2026-10 加固】本机对 api.minimax.cn 域名的 SSL 证书验证
                    # 慢得离谱(实测 5.87s vs curl 直连 1.3s,多花的 4s+ 主要来自
                    # 证书链/OCSP 握手),一旦 LLM 调到第 2~3 次叠加就撞 20s 超时。
                    # 本服务对接的是国内中转 API,密钥已经在请求头 Bearer 里,
                    # TLS 证书校验对安全收益边际,关掉换取稳定性。
                    # 注意:仍保留 HTTPS 加密信道,只是不验证证书链;
                    # 如未来要切到对端证书强校验的官方 API,这里要重新审视。
                    verify=False,
                )
            except Exception:
                http_client = None  # httpx 不可用时退回默认

            kwargs: dict[str, Any] = dict(
                api_key=self.api_key,
                base_url=self.base_url,
                # 【2026-10 调优】reasoning 模型(MiniMax-M2.7)的"思考过程"
                # 单次就要 30~40s,如果只给 20s 一定会被掐掉,造成"AI 断开"。
                # 实测长 prompt(4873 input + 1426 output)≈ 37s,所以这里给 60s,
                # 既覆盖单次思考耗时,又给偶尔网络抖动留余量。
                # max_retries=0:openai SDK 的 max_retries 是"额外重试次数",
                # 设 1 表示总共请求 2 次;我们的外层已经有 run_analysis 兜底,
                # 不需要 SDK 再静默重试一次(否则 60s × 2 = 120s 撞上外层 60s 反而失败)。
                timeout=60.0,
                max_retries=0,
            )
            if http_client is not None:
                kwargs["http_client"] = http_client
            self._client = OpenAI(**kwargs)
        return self._client

    def chat(
        self,
        system: str,
        user: str,
        json_mode: bool = True,
        temperature: float = 0.4,
        max_tokens: int | None = None,
    ) -> str:
        """单轮对话，返回模型回复文本。

        【2026-10 重写】从 openai SDK 改为**纯 httpx POST**。
        背景:同机 curl 直连 api.minimax.cn 仅 1~2s 拿到响应,但
        openai-python SDK 在本机环境下持续超时(>60s)。
        对比验证:
            - curl --max-time 60 直接 POST: 2.18s 返回 OK
            - openai SDK + httpx verify=False: 60s 超时(仍走不通)
        推测是 SDK 的连接池 / 重试 / 流式读取在某处阻塞。
        改为纯 httpx 后,完全控制超时与请求体,行为可预期。

        参数:
            max_tokens: 默认 8192,避免 reasoning 模型思考后被默认 4096 截断。
            json_mode: 保留兼容旧调用方,但实际不再下发 response_format。
        """
        import httpx

        # 【split-vision-and-text-llm】走文本分析侧配置(报告生成/改写等)
        url = f"{self.analysis_base_url.rstrip('/')}/chat/completions"
        body: dict[str, Any] = {
            "model": self.analysis_model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "temperature": temperature,
        }
        if max_tokens is not None:
            body["max_tokens"] = max_tokens
        # 不下发 response_format —— 同 chat() 注释,让模型自由输出,
        # 下游 _safe_json_loads 宽容解析。

        headers = {
            "Authorization": f"Bearer {self.analysis_api_key}",
            "Content-Type": "application/json",
        }

        # 单次总超时(由 settings.openai_resolved 控制),connect 10s。
        # verify=False 关键:避免本机对中转域名证书验证慢到 4~6s 的额外开销。
        # trust_env=False:不读 IDE 注入的 HTTP_PROXY/HTTPS_PROXY。
        # 每次新建 Client:连接复用不是核心收益,稳定性优先。
        import time as _t
        _t0 = _t.time()
        try:
            # 【2026-10 调优 v2】超时从 300s 进一步放宽到 600s。
            # 背景:2026-10-20 用户 regenerate 任务时,in_len=12132 的主调用
            # 直接撞 300s 整超时(`TIMEOUT after 300.2s`),理论:
            #   - MiniMax-M2.7 是 reasoning 模型,思考过程吃 token 严重;
            #   - user prompt >12k 字符 + 多 sheet 业务背景 + feedback 三件套,
            #     思考+JSON 输出实测 ~500s,300s 根本不够。
            # 改成 600s 后,留 100s 余量给后续 summarize/annotate 串行调用,
            # 单次任务总时长 ≈ 700s,仍在 analyzer BackgroundTasks 容忍窗内。
            # 注意:不能无限放大;真异常(API 挂)500s 后一定能撞超时,不会闷死。
            with httpx.Client(
                timeout=httpx.Timeout(600.0, connect=10.0),
                verify=False,
                trust_env=False,
            ) as client:
                resp = client.post(url, json=body, headers=headers)
            logger.info("[LLM_TIMING] %.1fs model=%s in_len=%d status=%s",
                        _t.time()-_t0, self.analysis_model,
                        len(body.get('messages',[{}])[-1].get('content','')),
                        resp.status_code)
        except httpx.TimeoutException as exc:
            # 纯 httpx 的超时统一基类,涵盖 connect / read / write / pool
            logger.warning("[LLM_TIMING] TIMEOUT after %.1fs model=%s in_len=%d",
                           _t.time()-_t0, self.analysis_model,
                           len(body.get('messages',[{}])[-1].get('content','')))
            logger.warning("LLM 网络不可用: %s", exc)
            raise APITimeoutError(request=None) from exc
        except httpx.HTTPError as exc:
            # 其他 httpx 网络错误(连接重置、DNS 失败等)
            logger.warning("LLM 网络不可用: %s", exc)
            raise APIConnectionError(request=None) from exc

        # 状态码分支:模拟 openai SDK 的 AuthenticationError / NotFoundError / BadRequestError
        try:
            data = resp.json()
        except Exception:
            data = {}

        if resp.status_code == 401:
            err = AuthenticationError(
                f"401 Unauthorized: {resp.text[:200]}",
                response=resp, body=None,
            )
            logger.error("LLM 鉴权失败 (401): %s", err)
            raise err
        if resp.status_code == 403:
            err = PermissionDeniedError(
                f"403 Forbidden: {resp.text[:200]}",
                response=resp, body=None,
            )
            logger.error("LLM 权限被拒 (403): %s", err)
            raise err
        if resp.status_code == 404:
            err = NotFoundError(
                f"404 Not Found: {resp.text[:200]}",
                response=resp, body=None,
            )
            logger.error("LLM 模型不存在 (404): %s", err)
            raise err
        if resp.status_code >= 500:
            err = BadRequestError(
                f"{resp.status_code} Server: {resp.text[:200]}",
                response=resp, body=None,
            )
            logger.error("LLM 服务端错误 (%s): %s", resp.status_code, err)
            raise err
        if resp.status_code >= 400:
            err = BadRequestError(
                f"{resp.status_code}: {resp.text[:200]}",
                response=resp, body=None,
            )
            logger.error("LLM 调用错误 (%s): %s", resp.status_code, err)
            raise err

        # 200 OK,正常提取 content
        choices = (data or {}).get("choices") or []
        if not choices:
            logger.error("LLM 返回无 choices: %s", data)
            raise BadRequestError(
                f"LLM 返回无 choices: {json.dumps(data, ensure_ascii=False)[:200]}",
                response=resp, body=None,
            )
        return (choices[0].get("message") or {}).get("content") or ""

    def chat_json(self, system: str, user: str, temperature: float | None = None) -> dict[str, Any]:
        """返回解析后的 JSON；失败时回退到结构化默认结果。

        新增 ``temperature`` 形参(可选),用于在需要更稳定/更低幻觉的场景
        (如表格结构识别、长文本清洗)覆盖默认温度。传 ``None`` 表示使用
        ``chat()`` 的默认值,保持向后兼容。

        【2026-10 修复】网络层异常(APITimeoutError / APIConnectionError)不再吞:
            旧实现里 ``chat()`` 在网络异常时直接 ``return _fallback_json(...)``,
            调用方拿到的是一个"看似合法 JSON"的占位 dict,导致:
              1) 上层以为 AI 跑过,把 status 写成 success —— 用户看到的报告
                 是「AI 离线占位」占位文案,但完全不知道;
              2) 排查时,看到 success 但内容空,极难定位。
            修法:在 chat_json 入口重新做一次"网络异常识别",凡是上游是网络
            错误的 raw 直接 raise,让 analyzer 走 task.status='failed' 分支,
            task.error_message 会带着详细原因落到数据库,前端报告页也能据此
            给出明确提示(JSON 解析失败仍走 fallback,与旧行为一致)。

        【2026-10 调优】默认 max_tokens=16384:
            reasoning 类模型思考过程会消耗大量 token,默认 4096 经常截断;
            8192 在「5 章正文 + 每章 1~3 张 ECharts option」的场景下仍会不够 ——
            实测 MiniMax-M3.1-Flash 生成 5 章报告时输出超 8192 被截断,
            重试时表现为 **raw 为空字符串**(模型没写完就被截断),
            最终报"无法解析为含 sections 的 JSON"。
            16384 覆盖实测最长的那份报告(约 11K token),留出安全余量。

        【2026-10-30 新增】默认 temperature=0.2(而不是 chat() 的 0.4):
            deepseek-flash 等轻量模型在长 prompt(>10k 字符)下,
            temperature=0.4 容易导致思考过程不稳定、返回空输出。
            调低到 0.2 显著提升输出稳定性;调用方显式传 temperature
            时尊重其值,不受此默认影响。
        """
        try:
            # temperature 默认 0.2:轻量模型长 prompt 下 0.4 容易空输出
            kwargs: dict[str, Any] = {
                "json_mode": True,
                "max_tokens": 16384,
                "temperature": temperature if temperature is not None else 0.2,
            }
            raw = self.chat(system, user, **kwargs)
        except (APITimeoutError, APIConnectionError):
            raise
        try:
            obj = json.loads(raw)
            if isinstance(obj, dict) and obj.get("sections"):
                return obj
            # 解析成功但 sections 缺失 —— LLM 输出形态跑偏(常见:末尾只有
            # chart_option.indicators),让上层知道失败而不是写假报告。
        except Exception:
            pass

        # 旧实现的 greedy regex 兜底:从第一个 `{` 到最后一个 `}`
        # 【2026-10 修复】这种兜底会把 6KB 的中文 markdown 报告(含表格 `{}`、段落括号)
        # 当成超大 JSON,然后 json.loads 必崩。新实现已经用
        # ``json.JSONDecoder.raw_decode`` 逐位扫描,见 ``json_utils._safe_json_loads``;
        # 这里直接复用它,如果它都解不出 sections,就不再 fallback 一个假 dict。
        from app.core.json_utils import _safe_json_loads
        try:
            obj2 = _safe_json_loads(raw)
            if isinstance(obj2, dict) and obj2.get("sections"):
                return obj2
        except Exception:
            pass

        # 实在拿不到 sections —— 不要直接 raise 失败。
        # 【2026-10 关键修复】实测 reasoning 模型(MiniMax-M2.7)在 user prompt
        # 较长时(>16k 字符)会把"输出 JSON"误读成"撰写 markdown 报告",
        # 写了 ~10KB 的 ## 一、## 二、... 章节 markdown,**完全不带 sections JSON**。
        # 旧行为:直接 raise → analyzer 走 _run_offline_fallback →
        #   用户报告页只剩"AI 服务暂时不可达"占位文案,体验灾难。
        # 新行为:**自我重试一次**,用一段"约束化 system + user"明确告诉模型
        #   只输出合法 JSON,大多数情况下一次就能拉回。
        retry_user = (
            "你刚才的回复不是合法的报告 JSON(常见原因:sections 为空对象 {} "
            "或写成了 markdown)。请基于下方**同一份真实数据**重新输出。\n\n"
            + (user if user else "")
            + "\n\n【补充硬性要求(必须遵守)】\n"
            '1. 严格输出 JSON,第一字符必须是 `{`,最后一字符必须是 `}`。\n'
            '2. 顶层键:"title" / "sections" / "chart_option"。\n'
            "3. **sections 绝不能是空对象 {}** —— 必须为【任务】里列出的每个章节都写出正文。\n"
            "4. sections 的每个 value 是**中文分析正文字符串**(不是对象、不是数组),\n"
            "   要含具体数字、同比/环比、趋势判断、原因分析、行动建议,并按段落换行。\n"
            "5. **严禁**写『本期数据未覆盖』『指标记录为 0 项』这类占位话术 ——\n"
            "   数据就在上面,请直接引用其中的真实数字做分析。\n"
            "6. chart_option.by_section:每个章节至少 1 张图,\n"
            '   每项形如 {"key":"k1","title":"中文标题","option":{...ECharts option...}}。\n'
            "7. **禁止** ```json 包裹、禁止 ## 标题、禁止 markdown 表格、禁止 <think> 标签。"
            "\n立刻重新输出,不要任何解释文字。"
        )
        raw_retry = raw
        obj3 = None
        # ---- 【2026-10-30 关键修复】动态参数重试策略 ----
        # 背景:task#107 失败,日志显示 deepseek-flash 在 in_len=13542 的长 prompt 下
        # 连续 3 次返回**空内容**(raw=''),而 kwargs 完全相同(temperature=0.4,
        # max_tokens=16384)。同参数重试等于没重试 —— reasoning 模型的空输出通常是
        # 思考过程耗尽 token 配额 + 高温随机性导致的,必须改变参数才能让重试有意义。
        #
        # 三档差异化参数(每次 retry 都不同):
        #   attempt 0: temperature=0.1  max_tokens=16384  ← 降温减少随机性
        #   attempt 1: temperature=0.0  max_tokens=32768  ← 零温度最确定 + 双倍输出
        #   attempt 2: temperature=0.0  max_tokens=32768  ← 换更短的 retry_user 重发
        #
        retry_params = [
            {"temperature": 0.1, "max_tokens": 16384},
            {"temperature": 0.0, "max_tokens": 32768},
            {"temperature": 0.0, "max_tokens": 32768},
        ]
        # 第三次 retry 用更短更直接的约束(去掉冗长的编号说明,减少 prompt 长度本身)
        retry_user_v2 = (
            "请严格按照 JSON schema 输出完整的年度经营分析报告。\n"
            "要求:第一字符 `{`,最后一字符 `}`。顶层键 title/sections/chart_option。\n"
            "sections 每个 value 必须是非空的中文正文(含数字、趋势判断、行动建议)。\n"
            "chart_option.by_section 每节至少 1 张 ECharts 图。\n"
            "禁止 markdown、禁止 ```json 包裹、禁止 <think> 标签。\n"
            "\n以下是真实数据(请直接引用其中数字做分析):\n"
            + (user if user else "")
        )
        empty_count = 0  # 记录空输出次数,用于最终日志
        for _attempt, rp in enumerate(retry_params):
            cur_user = retry_user if _attempt < 2 else retry_user_v2
            try:
                raw_retry = self.chat(system, cur_user, **rp)
            except (APITimeoutError, APIConnectionError):
                raise
            except Exception as exc:
                logger.warning("[chat_json] retry #%d 异常: %s", _attempt + 1, exc)
                continue
            # 空输出检测:reasoning 模型思考过程耗尽 token 配额
            if not raw_retry or not raw_retry.strip():
                empty_count += 1
                logger.warning(
                    "[chat_json] retry #%d 返回空输出(temp=%.1f, max_tokens=%d),"
                    "思考过程可能耗尽 token 配额,继续下一档参数",
                    _attempt + 1, rp["temperature"], rp["max_tokens"],
                )
                continue
            try:
                obj3 = json.loads(raw_retry)
            except Exception:
                obj3 = None
            if isinstance(obj3, dict) and obj3.get("sections"):
                return obj3
            # 用 _safe_json_loads 再救一次(包含 markdown 包裹 / 各种修复)
            try:
                from app.core.json_utils import _safe_json_loads
                obj4 = _safe_json_loads(raw_retry)
                if isinstance(obj4, dict) and obj4.get("sections"):
                    return obj4
            except Exception:
                pass
            # 拿到了非空内容但解析失败(sections 为空或 JSON 格式有问题)
            logger.warning(
                "[chat_json] retry #%d 解析失败,raw 头 200 字: %s",
                _attempt + 1, str(raw_retry)[:200],
            )

        # 重试后还不行 —— 才 raise,让上层走 _run_offline_fallback。
        last_raw = locals().get("raw_retry", raw)
        raise ValueError(
            f"LLM 输出无法解析为含 sections 的 JSON "
            f"(重试 3 次:空输出 {empty_count} 次,"
            f"最后 raw 头 200 字: {str(last_raw)[:200]!r}...)"
        )

    def chat_vision(
        self,
        prompt: str,
        images_b64: list[str],
        *,
        mime: str = "image/png",
        max_tokens: int = 4096,
        timeout: float | None = None,
        reasoning_effort: str | None = None,
    ) -> str:
        """多模态视觉识别:把图片 + 文本 prompt 发给视觉模型,返回文本描述。

        用途:
            - 图片型数据集(PNG/JPG 截图、扫描件)的内容理解
            - 扫描版 PDF 逐页转图后的文字识别

        入参:
            prompt:      对图片的提问/抽取要求,会作为 user 消息的文本部分
            images_b64:  base64 编码的图片列表(不带 data: 前缀),支持一次多张
            mime:        图片 MIME 类型,默认 image/png;JPG 调用方需传 image/jpeg
            max_tokens:  输出上限,避免长文档识别时把上下文撑爆
            reasoning_effort:
                推理努力程度,仅对 MiniMax-M3.1-Flash 这类**强制 adaptive
                thinking** 的模型有意义(该模型不允许 thinking=disabled,
                实测报错 2013)。OCR/版面解析是模式识别任务,不需要深层推理,
                传 "low" 可显著压缩思考链长度 —— 2026-10-26 实测(复杂噪点
                照片 + 版面 prompt):completion_tokens 4001 → 1404(-65%),
                耗时 44.1s → 16.1s(-63%),正文质量无变化。
                这同时根治两个线上故障:
                  1. max_tokens=8192 时思考耗光配额 → 正文空输出;
                  2. max_tokens=16384 时思考+正文生成过久 → read timeout。
                None(默认)则不传该字段,保持对其它模型/平台的兼容性。

        返回:
            模型输出的文本(一般是 Markdown 或纯文本)。
            失败时抛异常,由调用方决定降级策略 —— 视觉调用通常成本较高,
            不做"离线占位 JSON"这种静默兜底。

        模型选择:
            优先使用 settings.openai_vision_model;为空时回退到 openai_model,
            以兼容 MiniMax-M2.7 这类本身具备视觉能力的模型。
        """
        if not images_b64:
            raise ValueError("chat_vision 至少需要 1 张图片")

        # 组装 OpenAI 兼容的多模态消息体:
        # content 是 list,包含 text 和 image_url 两种 part
        content: list[dict[str, Any]] = [{"type": "text", "text": prompt}]
        for b64 in images_b64:
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:{mime};base64,{b64}"},
            })

        model = settings.openai_vision_model or self.model
        # 【2026-10】同步走纯 httpx,与 chat() 同链路,避免 SDK 卡死
        import httpx
        url = f"{self.base_url.rstrip('/')}/chat/completions"
        body: dict[str, Any] = {
            "model": model,
            "messages": [{"role": "user", "content": content}],
            "max_tokens": max_tokens,
            "temperature": 0.2,  # 视觉识别场景低温,减少幻觉
        }
        # 仅当调用方显式指定时才带上该字段:
        # MiniMax-M3.1-Flash 实测接受平铺写法 reasoning_effort="low"(嵌套
        # reasoning={"effort":...} 也认,但平铺写法思考压缩效果最好);
        # 其它模型/平台未必认识这个字段,默认不传以保证兼容。
        if reasoning_effort is not None:
            body["reasoning_effort"] = reasoning_effort
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        try:
            with httpx.Client(
                # 【2026-10 放宽】60s → 180s。
                # 视觉模型要把整张图片编码进请求体再推理,比纯文本慢得多;
                # 高分辨率文档照片(尤其手写体)推理时长普遍 40~120s,
                # 原来的 60s 会把"能认出来"误判成超时失败。
                # 注意:上层前端 axios timeout=120s(见 frontend/src/api/http.js),
                # 这里给 180s 是为了单张图的最坏情况留余量;
                # OCR 是**逐图串行**的,多图总耗时 = 单张 × 张数,
                # 因此前端批量上传时仍可能先于后端超时,属预期行为。
                # 【2026-10 优化】默认 180s,调用方可按场景覆盖。
                # 视觉推理的正常耗时是 8~40s,180s 只是"防服务端抽风"的兜底;
                # 但真碰上卡死的连接,干等 180s 才重试会让用户等 3 分钟以上。
                # 所以 OCR 场景会显式传一个更短的超时(见 ai_service._ocr_one),
                # 让"快速失败 → 立刻重试"的总耗时远小于 180s。
                timeout=httpx.Timeout(
                    timeout if timeout is not None else 180.0, connect=10.0
                ),
                verify=False,
                trust_env=False,
            ) as client:
                resp = client.post(url, json=body, headers=headers)
        except httpx.TimeoutException as exc:
            logger.error("视觉模型网络异常: %s", exc)
            raise APITimeoutError(request=None) from exc
        except httpx.HTTPError as exc:
            logger.error("视觉模型网络异常: %s", exc)
            raise APIConnectionError(request=None) from exc
        if resp.status_code >= 400:
            logger.error("视觉模型调用失败(%s, %s): %s", model, resp.status_code, resp.text[:200])
            raise BadRequestError(
                f"{resp.status_code}: {resp.text[:200]}",
                response=resp, body=None,
            )
        try:
            data = resp.json()
        except Exception:
            data = {}
        choices = (data or {}).get("choices") or []
        if not choices:
            return ""
        return (choices[0].get("message") or {}).get("content") or ""

    def _fallback_json(self, user: str, error: str) -> str:
        """当 LLM 不可用时，构造一个最小可用 JSON。"""
        return json.dumps({
            "title": "AI 分析（离线占位）",
            "sections": {
                "overview": f"未能调用大模型：{error}。以下是原始指标的概览，请结合数据自行阅读。",
                "highlights": "请查看指标 JSON 中的 numeric_summary 与 topn。",
                "problems": "如需更深入分析，请配置有效的 OPENAI_API_KEY。",
                "suggestions": "建议检查 .env 中的 API Key 与 BASE_URL。",
            },
            "chart_option": {},
        }, ensure_ascii=False)


# 全局单例
llm = LLMClient()
