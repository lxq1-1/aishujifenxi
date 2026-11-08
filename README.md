# AI 数据处理工作台

面向行政、运营、销售、人事等日常处理 Excel 台账的办公人员的一站式 Web 系统。
用户上传 Excel/CSV 台账后，系统自动完成**数据清洗 → 指标计算 → 图表生成 → AI 报告**，并支持导出 PDF / Word / Excel / Markdown。

## 技术栈
- **前端**：Vue 3 + Vite + Vue Router + Pinia + Element Plus + ECharts
- **后端**：FastAPI + SQLAlchemy + SQLite + Pandas + OpenAI 兼容接口
- **AI**：从 `.env` 读取 `OPENAI_API_KEY` / `OPENAI_BASE_URL` / `OPENAI_MODEL`，所有数值计算由 Pandas 完成、LLM 仅做文本生成

## 目录
```
.
├── backend/                # FastAPI 后端
│   ├── app/
│   │   ├── api/            # 路由：datasets/templates/tasks/workbench/tools/reports
│   │   ├── core/           # 配置、数据库、种子模板
│   │   ├── models/         # SQLAlchemy 模型（含 RAG 预留字段）
│   │   ├── services/       # pandas_engine / llm_client / ai_service / analyzer / exporter
│   │   └── main.py
│   ├── uploads/            # 上传的 Excel/CSV
│   ├── data/app.db         # SQLite 数据库（自动创建）
│   ├── reports/            # 导出文件临时目录
│   ├── requirements.txt
│   └── .env                # 从 .env.example 复制
├── frontend/               # Vue 3 前端
│   ├── src/
│   │   ├── views/          # 6 个页面：首页 / 工作台 / 数据集 / 模板 / 历史任务 / 报告预览
│   │   ├── components/
│   │   ├── router/
│   │   ├── store/          # Pinia：工作台草稿
│   │   ├── api/            # 统一 HTTP
│   │   └── assets/main.css # 深色科技风主题
│   ├── package.json
│   └── vite.config.js
└── .trae/specs/            # 开发规范文档
```

## 启动

### 1. 启动后端（端口 8000）
```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp ../.env.example .env   # 或手动从项目根目录的 .env.example 复制
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

### 2. 启动前端开发服务器（端口 5173）
```bash
cd frontend
npm install
npm run dev
```
浏览器访问 [http://localhost:5173](http://localhost:5173)

### 3. 生产构建（前端）
```bash
cd frontend && npm run build
```
构建后 `frontend/dist` 会被后端的 `app.main:app` 自动挂载（FastAPI 启动时检测到 dist 就接管静态资源 + SPA history fallback），访问 [http://localhost:8000](http://localhost:8000) 即可。

## 6 个页面
| 路径 | 说明 |
|------|------|
| `/` | 首页：产品介绍 + 3 卡片 + 4 步流程 + CTA |
| `/workbench` | 工作台：3 栏布局（数据集/表格/AI 结果）+ 5 个辅助工具弹窗 |
| `/data-datasets` | 数据集管理：上传/列表/预览/重命名/删除 |
| `/templates` | 模板管理：4 套内置模板 + 分类切换 + 「使用此模板」 |
| `/tasks` | 历史任务：状态筛选 + 批量删除 + 重新分析 |
| `/reports/:id` | 报告预览：正文 + 图表 + 指标 + 摘要 + 异常批注 + 4 种导出 |

## AI 能力
- 结构化报告生成（按 4 套模板）
- 自然语言数据对话（识别最大/最小/总和/平均等意图）
- 数据集对比解读（差值/增长率 + 文案）
- 报告摘要与异常批注
- 字段类型辅助识别（Pandas 推断 + 中文数字/百分号归一）

> LLM 不可达时系统自动降级：返回指标 JSON + Pandas 默认图表 + 占位文案，确保流水线稳定。

## 数据模型
| 表 | 关键字段 | RAG 预留 |
|----|----------|----------|
| `dataset` | id, name, file_path, row_count, status | — |
| `report_template` | id, name, category, prompt_structure | — |
| `task` | id, dataset_id, template_id, status, indicators_json, report_content, chart_option_json | `use_knowledge_base`(bool), `kb_ids`(str) |

## API 速查
- `GET  /api/health`
- 数据集：`GET /api/datasets`、`POST /api/datasets/upload`、`GET /api/datasets/{id}/preview`、`POST /api/datasets/{id}/rename`、`DELETE /api/datasets/{id}`
- 模板：`GET /api/templates`、`GET /api/templates/{id}`
- 任务：`GET /api/tasks`、`POST /api/tasks`、`GET /api/tasks/{id}`、`POST /api/tasks/{id}/reanalyze`、`DELETE /api/tasks/{id}`、`POST /api/tasks/batch`
- 工作台：`POST /api/workbench/analyze`（同步执行）
- 工具：`POST /api/tools/merge | validate | nl_query | compare`
- 导出：`GET /api/reports/{id}/export?format=pdf|docx|excel|md`
