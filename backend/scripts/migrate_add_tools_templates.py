"""一次性幂等迁移脚本:补齐工作台工具结果用的 3 条 ReportTemplate。

新增模板(【tools-results-history-persistence】)
-----------------------------------------
- 数据校验报告 / category=数据校验 / structure=校验概览/问题清单/处理建议
- 数据问答报告 / category=数据问答 / structure=问题/AI 回答/结果明细
- 数据对比报告 / category=数据对比 / structure=关键发现/对比详情/风险/建议

特性
----
- 幂等:每个模板按 name 检查,已存在则跳过。
- 失败时打印错误并返回非零退出码,便于 CI。
- 直接走 ORM(SQLAlchemy),与 seed.py 行为一致;不直接写 SQL。

用法
----
    python backend/scripts/migrate_add_tools_templates.py
"""
from __future__ import annotations

import os
import sys
import traceback

# 允许独立运行:把 backend 目录加入 sys.path,确保 `app.*` 能被 import
BACKEND_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if BACKEND_DIR not in sys.path:
    sys.path.insert(0, BACKEND_DIR)


TEMPLATES_TO_ADD = [
    {
        "name": "数据校验报告",
        "category": "数据校验",
        "description": (
            "工作台「数据校验」工具落库的模板。展示数据集质量指标、"
            "AI 解读的校验概览、问题清单(按严重度)与处理建议。"
        ),
        "structure": (
            "一、校验概览\n"
            "二、问题清单\n"
            "三、处理建议"
        ),
    },
    {
        "name": "数据问答报告",
        "category": "数据问答",
        "description": (
            "工作台「数据对话」工具落库的模板。展示用户提问、AI 回答"
            "与结果明细(数据表 / 数值)。"
        ),
        "structure": (
            "一、问题\n"
            "二、AI 回答\n"
            "三、结果明细"
        ),
    },
    {
        "name": "数据对比报告",
        "category": "数据对比",
        "description": (
            "工作台「数据集对比」工具落库的模板。展示两份数据集对比的"
            "关键发现、对比详情、风险与建议。"
        ),
        "structure": (
            "一、关键发现\n"
            "二、对比详情\n"
            "三、风险\n"
            "四、建议"
        ),
    },
]


def main() -> int:
    """执行幂等 INSERT。返回 0 = 成功,1 = 失败。"""
    try:
        from app.core.database import SessionLocal
        from app.models.report_template import ReportTemplate
    except Exception as exc:
        print(f"[migrate] 导入失败: {exc}", file=sys.stderr)
        traceback.print_exc()
        return 1

    db = SessionLocal()
    try:
        added = 0
        skipped = 0
        for tpl in TEMPLATES_TO_ADD:
            exists = (
                db.query(ReportTemplate)
                .filter(ReportTemplate.name == tpl["name"])
                .first()
            )
            if exists:
                # 已存在:不覆盖(沿用 seed.py 策略;description / structure / category
                # 已在 seed.py 里刷新过,本脚本只兜底"缺哪条补哪条")
                print(f"[migrate] 已存在,跳过: {tpl['name']}")
                skipped += 1
                continue
            db.add(ReportTemplate(**tpl, is_builtin=1))
            added += 1
            print(f"[migrate] 新增: {tpl['name']}")
        db.commit()
        print(f"[migrate] 完成 ✅ 新增 {added} 条,跳过 {skipped} 条")
        return 0
    except Exception as exc:
        db.rollback()
        print(f"[migrate] 失败: {exc}", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        db.close()


if __name__ == "__main__":
    raise SystemExit(main())
