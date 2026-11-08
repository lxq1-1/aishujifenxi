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

# 注册路由
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


# 前端静态资源（构建后由 vite 输出到 frontend/dist）
FRONTEND_DIST = Path(__file__).resolve().parent.parent.parent / "frontend" / "dist"
if FRONTEND_DIST.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIST), html=True), name="frontend")
