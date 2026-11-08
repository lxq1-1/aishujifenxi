"""FastAPI 应用入口。"""
import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pathlib import Path

from app.core.config import settings
from app.core.database import SessionLocal, init_db
from app.core.seed import seed_templates
from app.api import datasets, templates, tasks, workbench, tools, reports

logging.basicConfig(level=logging.INFO)

app = FastAPI(title="AI 数据处理工作台", version="1.0.0")

# 跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由：子 router 内部已带 /api 前缀，这里直接挂载即可
app.include_router(datasets.router)
app.include_router(templates.router)
app.include_router(tasks.router)
app.include_router(workbench.router)
app.include_router(tools.router)
app.include_router(reports.router)


@app.on_event("startup")
def on_startup():
    """启动：建表 + 写入种子模板。"""
    init_db()
    db = SessionLocal()
    try:
        seed_templates(db)
    finally:
        db.close()


@app.get("/api/health")
def health():
    return {"ok": True, "model": settings.openai_model}


@app.get("/api/ai/ping")
def ai_ping():
    """AI 连通性探针：用真实 key 打一发最小 chat 请求，返回明确状态。

    目的：取代 ``llm_client`` 对所有异常一律静默降级的隐患，
    让 ``curl /api/ai/ping`` 就能判断 AI 链路是否真正通畅。

    返回：
        - ``ok=True``  + ``reply``     ：鉴权/模型都通
        - ``ok=False`` + ``stage``     ：任何失败（auth/permission/model/network）
    """
    from openai import (
        APIConnectionError,
        APITimeoutError,
        AuthenticationError,
        BadRequestError,
        NotFoundError,
        OpenAI,
        PermissionDeniedError,
    )

    try:
        # 临时构造客户端，避免污染全局单例
        client = OpenAI(
            api_key=settings.openai_api_key,
            base_url=settings.openai_base_url,
            timeout=15.0,
        )
        r = client.chat.completions.create(
            model=settings.openai_model,
            messages=[{"role": "user", "content": "ping"}],
            temperature=0,
            max_tokens=5,
        )
        return {
            "ok": True,
            "model": settings.openai_model,
            "base_url": settings.openai_base_url,
            "reply": (r.choices[0].message.content or "").strip(),
        }
    except AuthenticationError as exc:
        return {"ok": False, "stage": "auth", "error": str(exc)[:300]}
    except PermissionDeniedError as exc:
        return {"ok": False, "stage": "permission", "error": str(exc)[:300]}
    except NotFoundError as exc:
        return {"ok": False, "stage": "model_not_found", "error": str(exc)[:300]}
    except BadRequestError as exc:
        return {"ok": False, "stage": "bad_request", "error": str(exc)[:300]}
    except (APITimeoutError, APIConnectionError) as exc:
        return {"ok": False, "stage": "network", "error": str(exc)[:300]}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "stage": "unknown", "error": f"{type(exc).__name__}: {exc}"[:300]}


# 前端静态资源（构建后由 vite 输出到 frontend/dist）
FRONTEND_DIST = Path(__file__).resolve().parent.parent.parent / "frontend" / "dist"
if FRONTEND_DIST.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIST), html=True), name="frontend")
