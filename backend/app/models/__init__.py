"""Models 包入口。

导入各模型模块以触发 ORM 元数据注册(Base.metadata)。
`init_db()` 里也会显式 import,这里再集中导一次,
是为了让 `from app.models import Deck` 这种用法直接可用。
"""
from app.models.dataset import Dataset
from app.models.report_template import ReportTemplate
from app.models.task import Task
from app.models.deck import Deck

__all__ = ["Dataset", "ReportTemplate", "Task", "Deck"]
