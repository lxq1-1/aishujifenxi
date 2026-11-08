"""项目配置加载：读取 .env 中的 API Key 与路径。"""
from datetime import datetime, timezone as _tz_mod
from pathlib import Path
from zoneinfo import ZoneInfo

from pydantic_settings import BaseSettings, SettingsConfigDict


# ===== 【fix-timezone-asia-shanghai 新增】统一时区 =====
# 历史代码全部使用 datetime.utcnow(),等价于把北京时间 -8 小时后存进数据库,
# 任务列表 / 数据集列表 / 报告页里"创建时间"普遍比本地时间早 8 小时。
# 这里把所有"now()"统一收敛到 Asia/Shanghai,存的是"本地墙钟时间"。
# 改数据库里存的值与改模型 default 配合,详见 scripts/migrate_timestamps_to_local.py。
DEFAULT_TZ_NAME = "Asia/Shanghai"


def _resolve_tz() -> ZoneInfo:
    """从 .env(若有 TZ=...)或兜底取 Asia/Shanghai 的 ZoneInfo。

    兜底是 Asia/Shanghai,因为:
      1. 项目用户全在东八区,服务器(若部署在 UTC)仍希望显示东八区时间;
      2. .env 没配也不会让代码运行时再 fallback 到 UTC,行为稳定。

    实现:不读全局 settings,而是从环境变量 OS 层读 TZ,
    这样函数可以在 settings 实例化之前被调用,避免循环 import。
    """
    import os
    name = (os.environ.get("TZ") or "").strip() or DEFAULT_TZ_NAME
    try:
        return ZoneInfo(name)
    except Exception:
        return ZoneInfo(DEFAULT_TZ_NAME)


def now_local() -> datetime:
    """返回带本地时区的当前时间(naive,墙钟)。

    设计要点:
      - 返回 **naive** datetime(不带 tzinfo),与 SQLite + SQLAlchemy 旧 DateTime 列兼容,
        老库不需 ALTER 也能存得进去,strftime 出来的就是墙钟,前端直接显示。
      - 数据库存的是"北京时间字符串"(例如 22:00),不再存 UTC 偏移;
        旧 UTC 数据由迁移脚本 `+8 hours` 校正。
    """
    tz = _resolve_tz()
    # 取带时区的"现在",再去掉 tzinfo → 落到 naive datetime(墙钟)
    return datetime.now(tz).replace(tzinfo=None)


# 项目根目录（backend/ 的父目录）
BASE_DIR = Path(__file__).resolve().parent.parent.parent
BACKEND_DIR = BASE_DIR / "backend"


def _pick_env_file() -> str:
    """在多个候选位置中挑一个实际存在的 .env。

    优先级：
        1. ``backend/backend/.env``   — 与 ``run.sh`` 的 ``cd backend`` 对应
        2. ``backend/.env``           — 与直接在 backend/ 下开发对应
        3. 当前工作目录下的 ``.env``  — 兜底

    背景：pydantic-settings 只接受单个 ``env_file``，但本项目历史上
    .env 放在 ``backend/backend/.env``，后来又有人放到 ``backend/.env``，
    导致在不同启动方式下 ``OPENAI_*`` 静默走默认值，AI 配置"看似空"。
    这里用一个解析期函数兼容两种布局，避免再踩同一个坑。
    """
    candidates = [
        BACKEND_DIR / "backend" / ".env",
        BACKEND_DIR / ".env",
        Path.cwd() / ".env",
    ]
    for p in candidates:
        if p.is_file():
            return str(p)
    # 都找不到就返回第一个，与历史行为一致（pydantic 会静默忽略）
    return str(candidates[0])


class Settings(BaseSettings):
    """从 .env 读取运行期配置。"""
    # OpenAI 兼容接口
    openai_api_key: str = ""
    openai_base_url: str = "https://api.openai.com/v1"
    openai_model: str = "gpt-4o-mini"

    # —— 多模态视觉模型(2026-10 新增) ——
    # 用于识别图片/扫描版 PDF 的内容,输出结构化文本。
    # 需要服务商提供视觉能力的模型(如 gpt-4o / MiniMax-M3.1-Flash / qwen-vl-max 等),
    # 若留空则默认与 openai_model 相同(适用于本身支持视觉的多模态模型)。
    openai_vision_model: str = ""

    # ===== 【split-vision-and-text-llm 新增】文本分析模型(与视觉模型分流) =====
    # 背景:项目里 LLM 调用分两大类,诉求不同,需要能各自选模型:
    #   1) 文本分析类 —— 报告生成/改写、摘要、异常批注、数据校验解读、
    #      NL 查询、数据集对比叙事、图表配置等,交给**推理/长文本更强**的模型;
    #   2) 视觉识别类 —— 照片转 PDF 的 OCR、扫描件/图片内容识别,
    #      交给**多模态**模型(见上面的 openai_vision_model)。
    # 三项全部留空时回退到上面那组 openai_* 配置(单模型部署的向后兼容)。
    # 注意:当前 llm_client 里 **只有 chat/chat_json 走这组配置**,
    # chat_vision 仍走 openai_base_url/openai_api_key/openai_vision_model。
    openai_analysis_api_key: str = ""
    openai_analysis_base_url: str = ""
    openai_analysis_model: str = ""
    # 非结构化文件抽取文本的最大字符数,超出截断,防止 prompt 爆炸
    max_extract_chars: int = 20000
    # 扫描版 PDF 视觉识别页数上限,防止意外大文件导致成本失控
    max_pdf_vision_pages: int = 20

    # 运行期目录
    upload_dir: Path = BACKEND_DIR / "uploads"
    data_dir: Path = BACKEND_DIR / "data"
    report_dir: Path = BACKEND_DIR / "reports"
    db_path: Path = BACKEND_DIR / "data" / "app.db"

    # 跨域
    cors_origins: list[str] = ["*"]

    # ===== 【fix-timezone-asia-shanghai 新增】时区配置 =====
    # 历史代码一律用 UTC 写入数据库,任务/数据集/报告的"创建时间"普遍早 8 小时。
    # 这里允许通过 .env 的 TZ=<IANA Name> 切换(默认 Asia/Shanghai);
    # 详见 app.core.config.now_local() 与 scripts/migrate_timestamps_to_local.py。
    tz_name: str = "Asia/Shanghai"

    model_config = SettingsConfigDict(
        # 兼容两种常见 .env 布局
        env_file=_pick_env_file(),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    def ensure_dirs(self) -> None:
        """启动时确保所有目录存在。"""
        for d in (self.upload_dir, self.data_dir, self.report_dir):
            d.mkdir(parents=True, exist_ok=True)


settings = Settings()
settings.ensure_dirs()
