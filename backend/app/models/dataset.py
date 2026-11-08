"""Dataset 表：管理用户上传的数据文件。

支持的文件类型(见 `file_type` 字段):
    - structured: Excel/CSV,可走 Pandas 指标计算
    - document:   Word/PDF,抽取纯文本用于 LLM 分析
    - image:      PNG/JPG 等,经多模态视觉识别为文本后使用
"""
from datetime import datetime
from sqlalchemy import String, DateTime, Integer, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.core.database import Base
from app.core.config import now_local  # 【fix-timezone-asia-shanghai 新增】写入北京时间


# ===== 【dataset-deleted-not-selectable 新增】内部锚点数据集的状态标记 =====
# 背景:PDF 工具(转 Word / 转表格 / 提取文字 / 压缩 / 照片转 PDF)在落库为历史任务时,
# Task.dataset_id 是 NOT NULL 且必须指向一个 Dataset 行,于是需要一个"占位"数据集
# 充当载体(文件指向原始 PDF 字节)。这类行**只作载体**,不代表用户上传的数据文件。
#
# 统一口径(本常量是唯一事实来源,后端 api 与前端都不再各自硬编码字符串):
#   1. 数据集列表接口直接过滤掉 → 前端「数据集」页与工作台各类下拉看到的是同一批;
#   2. 统计口径(total / 本月新增 / 已清洗 / 待处理)同样只算可见数据集,
#      避免出现"总数 64、列表只有 21 条"的错位;
#   3. 锚点仍可按 id 单独取用(db.get(Dataset, id)),历史任务与报告页不受影响。
INTERNAL_DATASET_STATUS = "PDF工具产物"


class Dataset(Base):
    __tablename__ = "dataset"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(200), nullable=False, comment="用户展示名")
    file_name: Mapped[str] = mapped_column(String(255), nullable=False, comment="原始文件名")
    file_path: Mapped[str] = mapped_column(String(500), nullable=False, comment="磁盘绝对路径")
    row_count: Mapped[int] = mapped_column(Integer, default=0, comment="解析后的行数(非结构化文件为 0)")
    status: Mapped[str] = mapped_column(String(50), default="已清洗", comment="已清洗/待清洗")
    # —— 多格式文件支持(2026-10 新增) ——
    # file_type 是文件能力分流的唯一事实来源,下游 pipeline 按它走不同分支:
    #   structured → Pandas DataFrame 指标计算
    #   document   → extracted_text 注入 LLM prompt
    #   image      → extracted_text(多模态视觉识别结果) 注入 LLM prompt
    file_type: Mapped[str] = mapped_column(
        String(20), default="structured", comment="structured/document/image"
    )
    # 非结构化文件抽取出的纯文本(截断存),供对比/分析时注入 prompt
    # structured 类型恒为空字符串,避免误用
    extracted_text: Mapped[str] = mapped_column(
        Text, default="", comment="非结构化文件抽取的纯文本"
    )
    # 抽取状态独立记录,失败不阻断主流程,仅在前端提示
    extract_status: Mapped[str] = mapped_column(
        String(20), default="success", comment="success/pending/failed"
    )
    create_time: Mapped[datetime] = mapped_column(DateTime, default=now_local)  # 【fix-timezone-asia-shanghai 修复】原 datetime.utcnow
    update_time: Mapped[datetime] = mapped_column(
        DateTime, default=now_local, onupdate=now_local  # 【fix-timezone-asia-shanghai 修复】原 datetime.utcnow
    )
