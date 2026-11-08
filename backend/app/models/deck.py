"""Deck 表：保存由 AI 生成的「演示稿大纲」。

背景:
    报告正文是"阅读材料",一页几千字;演示稿是"汇报材料",一页一个结论 + 几条要点。
    两者形态差异巨大,不能靠"把正文切开放进幻灯片"来变。
    因此这里单独存一份 AI 重构后的演示稿大纲(纯 JSON),
    导出 PPT 时按大纲渲染,而不是按正文渲染。

与 task 表的关系:
    - 一个 task 可以有多条 deck 记录(重新生成会追加新版本),
      但对外只暴露"该 task 最新的一条 ready 记录"。
    - deck 与 task 是 1:N,不是 1:1 —— 因为重新生成演示稿不应该覆盖历史,
      便于排查"上一版演示稿长什么样"。

字段说明:
    status:  generating(生成中) / ready(可用) / failed(失败)
    deck_json: 归一化后的演示稿大纲原文(JSON 字符串),结构见 deck_service。
"""
from datetime import datetime

from sqlalchemy import String, DateTime, Integer, Text, ForeignKey
from sqlalchemy.orm import Mapped, mapped_column

from app.core.database import Base
from app.core.config import now_local  # 统一写北京时间(与其它表一致)


class Deck(Base):
    """演示稿大纲表。"""

    __tablename__ = "deck"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    # 关联的报告任务。ondelete=CASCADE:任务被删,它的演示稿一并消失,
    # 避免留下指向不存在任务的孤儿数据。
    task_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("task.id", ondelete="CASCADE"),
        nullable=False, index=True,
        comment="所属任务 id",
    )
    # generating / ready / failed
    status: Mapped[str] = mapped_column(
        String(20), default="generating",
        comment="generating/ready/failed",
    )
    # 大纲 JSON 原文:
    #   {"title": str, "subtitle": str,
    #    "slides": [{"kind","title","bullets":[],"chart_ref","notes"}]}
    deck_json: Mapped[str] = mapped_column(
        Text, default="",
        comment="演示稿大纲(JSON 字符串)",
    )
    # 生成失败时的原因,用于前端展示与排查
    error_message: Mapped[str] = mapped_column(Text, default="")
    create_time: Mapped[datetime] = mapped_column(DateTime, default=now_local)
    update_time: Mapped[datetime] = mapped_column(
        DateTime, default=now_local, onupdate=now_local
    )
