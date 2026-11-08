"""ReportTemplate 表：报告模板定义。

注意:Prompt 结构不再落库,统一由后端 `app/core/prompt_builder.py` 按分类管理。
模板表只保存"用户可见"的字段(名称 / 分类 / 描述 / 章节结构)以及内置标记。
"""
from datetime import datetime
from sqlalchemy import String, DateTime, Integer, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.core.database import Base
from app.core.config import now_local  # 【fix-timezone-asia-shanghai 新增】写入北京时间


class ReportTemplate(Base):
    __tablename__ = "report_template"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(100), nullable=False)
    category: Mapped[str] = mapped_column(String(50), default="其他", comment="分类：年度/季度/述职/简报")
    description: Mapped[str] = mapped_column(String(500), default="")
    structure: Mapped[str] = mapped_column(Text, default="", comment="报告章节结构（Markdown,由用户在前端填写）")
    # 是否内置模板:1=内置(不可删除,仅可复制后修改);0=用户自建
    is_builtin: Mapped[int] = mapped_column(Integer, default=0, comment="1=内置,0=自建")
    create_time: Mapped[datetime] = mapped_column(DateTime, default=now_local)  # 【fix-timezone-asia-shanghai 修复】原 datetime.utcnow