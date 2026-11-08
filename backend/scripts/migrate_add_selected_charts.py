"""一次性幂等迁移脚本:给 task 表加 selected_charts_json 列。

特性
----
- 幂等:列已存在时跳过(SQLAlchemy create_all 会在老库静默跳过,本脚本兜底补 ALTER)
- 自动定位 backend/data/app.db(若不存在则跳过;通常 SQLAlchemy 已建好库)
- 用 PRAGMA table_info 探测列,避免"重复加列"异常
- 失败时打印 SQL 与异常,非零退出码,方便 CI 报错

用法
----
    python backend/scripts/migrate_add_selected_charts.py
"""
from __future__ import annotations

import os
import sqlite3
import sys

DB_CANDIDATES = [
    os.path.join(os.path.dirname(__file__), "..", "data", "app.db"),
    os.path.join(os.path.dirname(__file__), "..", "..", "data", "app.db"),
    # 历史工程里把 app.db 嵌在 backend/backend/data/,也兜底一下
    os.path.join(os.path.dirname(__file__), "..", "backend", "data", "app.db"),
]


def _find_db() -> str | None:
    """优先选存在且非空、含 task 表的 db 文件;否则选最大者。"""
    candidates = [os.path.abspath(p) for p in DB_CANDIDATES if os.path.exists(p)]
    if not candidates:
        return None
    # 过滤掉 0 字节空文件
    non_empty = [p for p in candidates if os.path.getsize(p) > 0]
    if not non_empty:
        return None
    # 进一步优先选"含 task 表"的;否则按文件大小降序,最大的最可能是真实库
    import sqlite3 as _sqlite
    valid: list[tuple[int, str]] = []
    for p in non_empty:
        try:
            con = _sqlite.connect(p)
            cur = con.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='task'")
            has_task = cur.fetchone() is not None
            con.close()
            valid.append((2 if has_task else 1, p))
        except Exception:
            continue
    if not valid:
        return None
    valid.sort(key=lambda x: (-x[0], -os.path.getsize(x[1])))
    return valid[0][1]


def main() -> int:
    db_path = _find_db()
    if not db_path:
        print("[migrate] 未找到 app.db,跳过(将由 SQLAlchemy create_all 自动建表)")
        return 0

    print(f"[migrate] 打开数据库: {db_path}")
    con = sqlite3.connect(db_path)
    try:
        cur = con.cursor()
        # 1) 确认 task 表存在
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='task'"
        )
        if cur.fetchone() is None:
            print("[migrate] task 表不存在,无需迁移")
            return 0

        # 2) 列是否存在
        cur.execute("PRAGMA table_info(task)")
        cols = [row[1] for row in cur.fetchall()]
        if "selected_charts_json" in cols:
            print("[migrate] selected_charts_json 已存在,无需迁移")
            return 0

        # 3) 加列(幂等已被上面的探测守住,这里直接执行 ALTER)
        print("[migrate] ALTER TABLE task ADD COLUMN selected_charts_json TEXT")
        cur.execute(
            "ALTER TABLE task ADD COLUMN selected_charts_json TEXT"
        )
        con.commit()
        print("[migrate] 完成 ✅")
        return 0
    except Exception as exc:
        print(f"[migrate] 失败: {exc}", file=sys.stderr)
        return 1
    finally:
        con.close()


if __name__ == "__main__":
    raise SystemExit(main())