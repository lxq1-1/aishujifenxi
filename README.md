# AI 数据处理工作台

面向行政、运营、销售、人事等日常处理台账的办公人员的一站式 Web 系统。
用户上传台账后，系统自动完成**文件解析 → 数据清洗 → 指标计算 → 图表生成 → AI 报告 → 多格式导出**。

## 支持的文件类型

| 类别 | 扩展名 | 处理方式 |
|---|---|---|
| 结构化 | `.xlsx` `.xls` `.csv` | Pandas 解析，多 sheet 纵向合并 |
| 文档 | `.docx` `.doc` `.pdf` `.txt` `.md` | 抽取纯文本存入 `extracted_text` |
| 图片 | `.png` `.jpg` `.jpeg` `.webp` `.bmp` | 多模态模型视觉识别为文本（OCR） |

## 技术栈

- **前端**：Vue 3 + Vite + Vue Router + Pinia + Element Plus + ECharts + vue-echarts + axios
- **后端**：FastAPI + SQLAlchemy + SQLite + Pandas + NumPy + OpenAI 兼容接口
- **报表/文档生成**：reportlab（PDF）、python-docx（Word）、python-pptx（PPT）、openpyxl（Excel）、matplotlib（图表 PNG 渲染）
- **文件解析**：pdfplumber、pdf2docx、pypdfium2、Pillow、opencv-python-headless
- **AI**：分**视觉识别侧**与**文本分析侧**两组独立配置，所有数值计算由 Pandas 完成、LLM 仅做文本生成

## AI 环境变量

从项目根目录 `.env.example` 复制为 `.env` 后填写（`.env` 已被 `.gitignore` 忽略，**严禁把真实密钥写进 `.env.example`**）。

| 变量 | 侧别 | 说明 |
|---|---|---|
| `OPENAI_API_KEY` | 视觉侧 | API 密钥 |
| `OPENAI_BASE_URL` | 视觉侧 | 接口地址 |
| `OPENAI_MODEL` | 视觉侧 | 模型名 |
| `OPENAI_VISION_MODEL` | 视觉侧 | 多模态模型，用于 OCR / 图片识别 |
| `OPENAI_ANALYSIS_API_KEY` | 文本侧 | 留空则回退到视觉侧 |
| `OPENAI_ANALYSIS_BASE_URL` | 文本侧 | 留空则回退到视觉侧 |
| `OPENAI_ANALYSIS_MODEL` | 文本侧 | 留空则回退到视觉侧 |
| `HOST` / `PORT` | 服务 | 后端监听地址与端口，供 `run.sh` 使用 |
| `CORS_ORIGINS` | 服务 | 允许的前端 origin，逗号分隔 |

## 目录

```
.
├── backend/                     # FastAPI 后端
│   ├── app/
│   │   ├── api/                 # 路由：datasets / templates / tasks / workbench / tools / reports
│   │   ├── core/                # config / database / seed / prompt_builder / deck_prompt / json_utils
│   │   ├── models/              # SQLAlchemy 模型：dataset / report_template / task / deck
│   │   ├── services/            # pandas_engine / llm_client / ai_service / analyzer
│   │   │                        # chart_renderer / dashboard_report / deck_service
│   │   │                        # exporter / file_extractor / pdf_service
│   │   ├── utils/
│   │   └── main.py              # 应用入口，挂载前端 dist + SPA fallback
│   ├── scripts/                 # 数据库迁移脚本
│   ├── uploads/                 # 上传台账
│   ├── data/app.db              # SQLite 数据库（自动创建）
│   ├── reports/                 # 导出文件临时目录
│   ├── requirements.txt
│   └── .env                     # 从 .env.example 复制
├── frontend/                    # Vue 3 前端
│   ├── src/
│   │   ├── views/               # 6 个页面：首页 / 工作台 / 数据集 / 模板 / 历史任务 / 报告预览
│   │   ├── router/
│   │   ├── store/               # Pinia：工作台草稿
│   │   ├── api/                 # 统一 HTTP 封装
│   │   ├── utils/               # chartExport.js（图表导出）
│   │   └── assets/main.css      # 深色科技风主题
│   ├── package.json
│   └── vite.config.js
└── .trae/specs/                 # 开发规范文档
```

## 启动

### 1. 启动后端（端口 8000）
```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp ../.env.example .env   # 再填入真实密钥
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
构建后 `frontend/dist` 会被 `app.main:app` 自动挂载（FastAPI 启动时检测到 dist 就接管静态资源 + SPA history fallback），访问 [http://localhost:8000](http://localhost:8000) 即可。

## 6 个页面

| 路径 | 说明 |
|------|------|
| `/` | 首页：产品介绍 + 3 卡片 + 4 步流程 + CTA |
| `/workbench` | 工作台：3 栏布局（数据集/表格/AI 结果）+ 辅助工具弹窗 |
| `/data-datasets` | 数据集管理：上传/列表/预览/下载/重命名/删除 |
| `/templates` | 模板管理：13 套内置模板 + 分类切换 + 新建/编辑/复制/删除 |
| `/tasks` | 历史任务：状态筛选 + 批量删除 + 重新分析 |
| `/reports/:id` | 报告预览：正文 + 图表 + 指标 + 摘要 + 异常批注 + 5 种导出 |

## AI 能力

- 结构化报告生成（内置 13 套模板）
- 自然语言数据对话（识别最大/最小/总和/平均等意图）
- 数据集对比解读（差值/增长率 + 文案）
- 报告摘要、异常批注与报告版本管理
- 字段类型辅助识别（Pandas 推断 + 中文数字/百分号归一）
- **AI 演示稿**：把报告正文重构为「一页一结论」的 PPT 大纲，支持讲稿导出
- **图表编辑**：预置仪表盘样式、在线改图、把自定义图表并入报告
- **数据清洗**：缺失值/异常值处理、字段合并、结果落库为任务
- **AI PDF 工具**：PDF 转 Word、PDF 转表格、提取文字、压缩、照片转 PDF（OCR）

> LLM 不可达时系统自动降级：返回指标 JSON + Pandas 默认图表 + 占位文案，确保流水线稳定。

## 导出格式

`GET /api/reports/{id}/export?format=...`

| format | 说明 |
|---|---|
| `pdf` / `docx`(或 `word`) / `excel`(或 `xlsx`) / `md`(或 `markdown`) | 报告正文导出 |
| `pptx`（或 `powerpoint`） | 支持 `theme=deep_cyan/midnight/light/business`，`source=report`（报告全文版）或 `ai`（AI 演示稿） |

## 数据模型

| 表 | 关键字段 | 说明 |
|----|----------|------|
| `dataset` | id, name, file_path, file_name, row_count, status, extracted_text | 上传的数据集 |
| `report_template` | id, name, category, description, structure, is_builtin | 报告模板，含 13 套内置 |
| `task` | id, name, dataset_id, template_id, status, indicators_json, report_content, chart_option_json, annotations_json, selected_charts_json, use_knowledge_base, kb_ids, error_message | 分析任务，`use_knowledge_base` / `kb_ids` 为 RAG 预留 |
| `deck` | id, task_id, status, deck_json, error_message | AI 演示稿大纲，与 task 为 1:N（保留历史版本） |

## API 速查

共 47 个模块化路由 + `/api/health`。

### 数据集
- `GET    /api/datasets`
- `POST   /api/datasets/upload`
- `GET    /api/datasets/{id}/preview`
- `GET    /api/datasets/{id}/download`
- `POST   /api/datasets/{id}/rename`
- `DELETE /api/datasets/{id}`

### 模板
- `GET    /api/templates`
- `GET    /api/templates/categories`
- `GET    /api/templates/{id}`
- `POST   /api/templates`
- `PUT    /api/templates/{id}`
- `DELETE /api/templates/{id}`
- `POST   /api/templates/{id}/duplicate`

### 任务
- `GET    /api/tasks`
- `POST   /api/tasks`
- `GET    /api/tasks/{id}`
- `POST   /api/tasks/{id}/reanalyze`
- `POST   /api/tasks/{id}/summarize`
- `POST   /api/tasks/{id}/regenerate`
- `DELETE /api/tasks/batch`
- `DELETE /api/tasks/{id}`

### 工作台
- `POST   /api/workbench/analyze`（同步执行）
- `GET    /api/workbench/charts/styles`
- `POST   /api/workbench/charts/preview`
- `POST   /api/workbench/charts/modify`
- `POST   /api/workbench/charts/report`

### 工具
- `POST   /api/tools/merge | validate | nl_query | compare | clean`
- `POST   /api/tools/chart_edit | chart_apply`
- `GET    /api/tools/schema`
- `POST   /api/tools/suggest_columns`
- `POST   /api/tools/pdf/to_word | pdf/to_excel | pdf/extract_text | pdf/compress | pdf/image_to_pdf`
- `GET    /api/tools/pdf/download/{file_name}`

### 报告与导出
- `GET    /api/reports/{id}/export?format=...`
- `POST   /api/reports/{id}/deck`
- `GET    /api/reports/{id}/deck`
- `GET    /api/reports/{id}/deck/script.md`
- `GET    /api/reports/{id}/preview_charts`
- `POST   /api/reports/{id}/visual-report`
