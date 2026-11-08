"""Task 表：分析任务记录。"""
from datetime import datetime
from sqlalchemy import String, DateTime, Integer, Text, Boolean, ForeignKey
from sqlalchemy.orm import Mapped, mapped_column

from app.core.database import Base
from app.core.config import now_local  # 【fix-timezone-asia-shanghai 新增】写入北京时间


class Task(Base):
    __tablename__ = "task"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(200), default="未命名任务")
    dataset_id: Mapped[int] = mapped_column(Integer, ForeignKey("dataset.id"), nullable=False)
    template_id: Mapped[int] = mapped_column(Integer, ForeignKey("report_template.id"), nullable=False)
    status: Mapped[str] = mapped_column(String(50), default="processing",
                                        comment="processing/success/failed")
    indicators_json: Mapped[str] = mapped_column(Text, default="")
    report_content: Mapped[str] = mapped_column(Text, default="")
    chart_option_json: Mapped[str] = mapped_column(Text, default="")
    annotations_json: Mapped[str] = mapped_column(Text, default="")
    # ===== 【chart-edit-prebuilt-dashboard 新增】用户前置预选的图表 =====
    # 弹窗「图表编辑」里点过「+ 添加至报表」的图表,临时绑定当前任务;
    # 提交到后端后存到本字段,分析时与 AI 自动生成图表按 title 去重合并。
    # 形态:JSON 字符串,反序列化为 [{"title", "option", "style_key"}, ...]。
    # 留 nullable=True,老数据保持 NULL 不报错。
    selected_charts_json: Mapped[str | None] = mapped_column(
        Text, nullable=True, default=None,
        comment="用户前置预选的图表(JSON 列表,提交 AI 分析时合并输出)",
    )
    use_knowledge_base: Mapped[bool] = mapped_column(Boolean, default=False)
    kb_ids: Mapped[str] = mapped_column(String(500), default="")
    error_message: Mapped[str] = mapped_column(Text, default="")
    create_time: Mapped[datetime] = mapped_column(DateTime, default=now_local)  # 【fix-timezone-asia-shanghai 修复】原 datetime.utcnow
    update_time: Mapped[datetime] = mapped_column(
        DateTime, default=now_local, onupdate=now_local  # 【fix-timezone-asia-shanghai 修复】原 datetime.utcnow
    )

    # ===== 报告版本管理字段(Spec: report-feedback-and-versioning) =====
    # parent_task_id: 当前 task 是某原 task 的"重新生成"子任务时,指向原 task 的 id;
    # nullable=True,自引用外键,ondelete=SET NULL 保证原任务被删后子任务仍可访问;
    # 原任务(基线)/非 regenerate 产生的任务保持 None。
    parent_task_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("task.id", ondelete="SET NULL"),
        nullable=True, default=None,
        comment="父任务 id(用于版本 lineage),None 表示基线",
    )
    # feedback_text: 用户在报告页对原 task 提的修改意见(用于 regenerate 后审计追溯);
    # 非空即代表"此任务是基于父任务的意见重写得到的"。
    # 注意:工作台「开始 AI 分析」的 analysis_intent 不落库,只走 annotations。
    feedback_text: Mapped[str] = mapped_column(
        Text, default="",
        comment="用户修改意见(仅在 regenerate 链路里写入)",
    )
