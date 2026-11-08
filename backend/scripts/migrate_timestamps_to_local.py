"""【fix-timezone-asia-shanghai 新增】存量时间戳从 UTC 校正到 Asia/Shanghai。

背景:
    历史代码用 ``datetime.utcnow()`` 写库,等于把北京时间 -8 小时后存入 SQLite。
    本次修复后,模型 default 已切到 ``now_local()``(Asia/Shanghai 墙钟),
    新写入的时间就是本地时间;但**已有行**仍是 UTC,前端继续显示"早 8 小时"。

本脚本:
    1. 备份 ``data/app.db`` → ``data/app.db.bak.<时间戳>``
    2. 把以下列统一 +8 小时:
        task.create_time, task.update_time
        dataset.create_time, dataset.update_time
        report_template.create_time
    3. 输出样本前 5 行供人工复核。
    4. 幂等:维护一个 ``data/.tz_migrated`` 标记文件,再次执行不重复 +8h。

用法:
    cd backend
    python scripts/migrate_timestamps_to_local.py
    # 可选:--offset-hours=8  指定偏移量(默认 8 = 北京时间相对 UTC 的偏移)
    # 可选:--force          忽略已迁移标记,强制再跑一遍(例如重新部署时)
"""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

# 让 ``python scripts/migrate_timestamps_to_local.py`` 直接可跑,
# 不依赖从 backend/ 根目录启动。
BACKEND_DIR = Path(__file__).resolve().parent.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from sqlalchemy import text  # noqa: E402

from app.core.config import settings  # noqa: E402
from app.core.database import engine  # noqa: E402


MIGRATED_MARKER = BACKEND_DIR / "data" / ".tz_migrated"
TARGET_COLUMNS: dict[str, list[str]] = {
    "task":             ["create_time", "update_time"],
    "dataset":          ["create_time", "update_time"],
    "report_template":  ["create_time"],
}


def _backup_db() -> Path:
    """启动前先备份原 db;失败则中止,绝不直接覆盖。"""
    db_path = Path(settings.db_path)
    if not db_path.exists():
        raise FileNotFoundError(f"数据库不存在: {db_path}")
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = db_path.with_suffix(db_path.suffix + f".bak.tz-{ts}")
    shutil.copy2(db_path, bak)
    print(f"[backup] 已备份 {db_path.name} → {bak.name}")
    return bak


def _already_migrated() -> bool:
    return MIGRATED_MARKER.exists()


def _mark_migrated(offset_hours: int) -> None:
    MIGRATED_MARKER.parent.mkdir(parents=True, exist_ok=True)
    MIGRATED_MARKER.write_text(
        f"migrated_at={datetime.now().isoformat(timespec='seconds')}\n"
        f"offset_hours={offset_hours}\n",
        encoding="utf-8",
    )


def _shift_column(conn, table: str, column: str, offset: timedelta) -> int:
    """把 ``table.column`` 全部 ``+= offset``(SQLite 直接用 datetime 算)。

    SQLite 的 ``datetime(..., '+8 hours')`` 在 datetime 字符串上等价于 +8h,
    不会影响带小数秒的字符串。
    返回受影响行数。
    """
    # 注意:datetime() 函数需要 ISO 8601 字符串。我们存的也是 SQLite 默认格式,
    # 形如 '2026-10-19 14:00:00' / '2026-10-19 14:00:00.123456',datetime() 都能识别。
    sql = text(
        f"UPDATE {table} "
        f"SET {column} = datetime({column}, :sign :hours || ' hours') "
        f"WHERE {column} IS NOT NULL"
    )
    # SQLite 不支持占位符复用 sign/hours,直接拼字符串(简单且安全,数字来自 argparse)
    sql_str = (
        f"UPDATE {table} "
        f"SET {column} = datetime({column}, '+{int(offset.total_seconds() // 3600)} hours') "
        f"WHERE {column} IS NOT NULL"
    )
    result = conn.exec_driver_sql(sql_str)
    return result.rowcount if hasattr(result, "rowcount") else -1


def _sample(conn, table: str, column: str, limit: int = 5) -> list[tuple]:
    rows = conn.exec_driver_sql(
        f"SELECT id, {column} FROM {table} "
        f"WHERE {column} IS NOT NULL "
        f"ORDER BY {column} DESC LIMIT {limit}"
    ).fetchall()
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="把存量时间戳从 UTC 校正到本地时区")
    ap.add_argument("--offset-hours", type=int, default=8,
                    help="给存量时间戳加多少小时(默认 8 = UTC → Asia/Shanghai)")
    ap.add_argument("--force", action="store_true",
                    help="忽略已迁移标记,强制再跑一遍")
    args = ap.parse_args()

    if _already_migrated() and not args.force:
        print(f"[skip] 已检测到迁移标记 {MIGRATED_MARKER.name},跳过。"
              f"若需重跑请加 --force。")
        return 0

    _backup_db()
    offset = timedelta(hours=args.offset_hours)
    print(f"[shift] offset = +{args.offset_hours}h(Asia/Shanghai 相对 UTC 的偏移)")

    total = 0
    with engine.begin() as conn:
        for table, cols in TARGET_COLUMNS.items():
            # 检查表是否存在(老 db 可能还没有某些表)
            exists = conn.exec_driver_sql(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
                (table,),
            ).fetchone()
            if not exists:
                print(f"[skip] 表 {table} 不存在,跳过")
                continue
            for col in cols:
                # 检查列是否存在
                cols_info = conn.exec_driver_sql(
                    f"PRAGMA table_info({table})"
                ).fetchall()
                col_names = {row[1] for row in cols_info}
                if col not in col_names:
                    print(f"[skip] 列 {table}.{col} 不存在,跳过")
                    continue
                n = _shift_column(conn, table, col, offset)
                total += n
                print(f"[ok]   {table}.{col}:已校正 {n} 行")

        # 抽样:每张表显示校正后的最新 5 行
        print("\n[verify] 抽样(校正后,最新 5 行):")
        for table, cols in TARGET_COLUMNS.items():
            for col in cols:
                sample = _sample(conn, table, col, limit=5)
                if sample:
                    pretty = "\n".join(f"      id={r[0]:>4}  {col}={r[1]}" for r in sample)
                    print(f"  - {table}.{col}:\n{pretty}")

    _mark_migrated(args.offset_hours)
    print(f"\n[done] 总计校正 {total} 行,标记 {MIGRATED_MARKER} 已写入。"
          f"请重启后端服务,前端时间显示即为 Asia/Shanghai.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())