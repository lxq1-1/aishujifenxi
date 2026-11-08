"""SQLAlchemy 数据库引擎与会话工厂。"""
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker

from app.core.config import settings


# 单一 SQLite 文件
engine = create_engine(
    f"sqlite:///{settings.db_path}",
    connect_args={"check_same_thread": False},
    echo=False,
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class Base(DeclarativeBase):
    """所有 ORM 模型的基类。"""
    pass


def get_db():
    """FastAPI 依赖：每次请求提供一个 Session。"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db() -> None:
    """根据所有 import 的模型创建表。

    同时对已有表做轻量迁移:
      - report_template.is_builtin 列不存在则添加(默认 0)
      - 4 个内置模板按名称回填 is_builtin=1
      - 老库若存在 prompt_structure 列则按 12 步标准迁移重建表
      - task.parent_task_id / task.feedback_text 列不存在则添加(Spec B)
        SQLite + create_all 对"已有表 + 新增列"不会自动 ALTER,需手动补;
        兼容老数据:旧 task 的 parent_task_id=NULL, feedback_text=''.

    SQLAlchemy 对**全新环境**会自动 create_all 时把这 2 列加进 CREATE TABLE;
    对**已有 task 表**的环境(本仓库的 data/app.db)则走下面这段 PRAGMA + ALTER,
    保证升级无感。
    """
    # 导入模型以注册到 Base.metadata
    from app.models import dataset as _dataset  # noqa: F401
    from app.models import report_template as _template  # noqa: F401
    from app.models import task as _task  # noqa: F401
    # 【generate-ai-presentation-deck 新增】演示稿大纲表。
    # 这是**新建表**(不是给老表加列),所以 Base.metadata.create_all 直接建好,
    # 不需要像 task.parent_task_id 那样走 PRAGMA + ALTER。
    from app.models import deck as _deck  # noqa: F401

    Base.metadata.create_all(bind=engine)

    # 轻量迁移:给老库补列
    with engine.begin() as conn:
        # ---- report_template ----
        cols_tpl = [row[1] for row in conn.exec_driver_sql("PRAGMA table_info(report_template)").fetchall()]
        if "is_builtin" not in cols_tpl:
            conn.exec_driver_sql("ALTER TABLE report_template ADD COLUMN is_builtin INTEGER DEFAULT 0")
        # 兼容老数据:把当前 4 个内置模板按名称回填 is_builtin=1
        for builtin_name in ("年度复盘", "季度分析", "述职报告", "数据简报"):
            conn.exec_driver_sql(
                "UPDATE report_template SET is_builtin = 1 WHERE name = ? AND (is_builtin IS NULL OR is_builtin = 0)",
                (builtin_name,),
            )
        if "prompt_structure" in cols_tpl:
            _drop_prompt_structure_column(conn)

        # ---- task:补 parent_task_id + feedback_text(Spec B 升级兼容) ----
        cols_task = [row[1] for row in conn.exec_driver_sql("PRAGMA table_info(task)").fetchall()]
        if "parent_task_id" not in cols_task:
            # SQLite 不允许在 ALTER ADD CONSTRAINT 时直接声明外键 + ondelete,
            # 这里把列加上后,自引用外键通过 SQLAlchemy ORM 端维护完整性。
            conn.exec_driver_sql("ALTER TABLE task ADD COLUMN parent_task_id INTEGER DEFAULT NULL")
        if "feedback_text" not in cols_task:
            conn.exec_driver_sql("ALTER TABLE task ADD COLUMN feedback_text TEXT DEFAULT ''")

        # ---- dataset:多格式文件支持(2026-10 新增) ----
        # 给老库补 file_type / extracted_text / extract_status 三列。
        # 存量数据都是 Excel/CSV,统一回填 file_type='structured' + extract_status='success',
        # 让老数据集在新代码下行为完全等价于旧版(走 Pandas 链路)。
        cols_ds = [row[1] for row in conn.exec_driver_sql("PRAGMA table_info(dataset)").fetchall()]
        if "file_type" not in cols_ds:
            conn.exec_driver_sql("ALTER TABLE dataset ADD COLUMN file_type VARCHAR(20) DEFAULT 'structured'")
            conn.exec_driver_sql("UPDATE dataset SET file_type = 'structured' WHERE file_type IS NULL OR file_type = ''")
        if "extracted_text" not in cols_ds:
            conn.exec_driver_sql("ALTER TABLE dataset ADD COLUMN extracted_text TEXT DEFAULT ''")
        if "extract_status" not in cols_ds:
            conn.exec_driver_sql("ALTER TABLE dataset ADD COLUMN extract_status VARCHAR(20) DEFAULT 'success'")
            conn.exec_driver_sql("UPDATE dataset SET extract_status = 'success' WHERE extract_status IS NULL OR extract_status = ''")


def _drop_prompt_structure_column(conn) -> None:
    """迁移掉已废弃的 `prompt_structure` 列。

    SQLite 不支持 DROP COLUMN(<12.4),采用业界通用的"12 步标准迁移":
        1. 改原表名 → report_template__bak
        2. 按新 schema 重建 report_template
        3. 把数据从 bak 拷回(丢弃 prompt_structure)
        4. DROP bak
    """
    new_cols = ["id", "name", "category", "description", "structure", "is_builtin", "create_time"]
    col_list = ", ".join(new_cols)
    conn.exec_driver_sql("ALTER TABLE report_template RENAME TO report_template__bak")
    conn.exec_driver_sql(
        "CREATE TABLE report_template (" +
        "id INTEGER PRIMARY KEY AUTOINCREMENT," +
        "name VARCHAR(100) NOT NULL," +
        "category VARCHAR(50) DEFAULT '其他'," +
        "description VARCHAR(500) DEFAULT ''," +
        "structure TEXT DEFAULT ''," +
        "is_builtin INTEGER DEFAULT 0," +
        "create_time DATETIME" +
        ")"
    )
    conn.exec_driver_sql(
        f"INSERT INTO report_template ({col_list}) "
        f"SELECT {col_list} FROM report_template__bak"
    )
    conn.exec_driver_sql("DROP TABLE report_template__bak")
