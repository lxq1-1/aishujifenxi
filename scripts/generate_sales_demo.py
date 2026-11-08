"""生成销售周报场景的演示 Excel(3 个月真实数字)。

为什么需要这份 demo:
    用户在生产报告页反馈"没有根据模板生成报告,也没有一个部分一张图或多个图" —
    根因是底层数据集「年度复盘_统计指标数据」全表缺失率 85%+,LLM 在贫数据下
    自创 overview 章节且把所有图堆到「数据图表」。修复后端归一化 + Prompt 后,
    需要一份含真实数字的销售明细来验证"按章节挂载图表"在生产报告页真的能看到。

本脚本生成 4 张 sheet:
    - 销售明细(本周/上月/上上月):按日聚合,200~500 条,字段:
        日期, 渠道(线上/线下/私域/直播), 产品SKU(20 个), 客户(50 个),
        销售额, 订单量, 客单价, 毛利率, 新客户数, 回款率, 退款金额
    - 周度KPI:周聚合,13 行
    - 渠道结构:4 行,占比/同比
    - 产品Top10:产品排名

输出路径:
    /Users/xiahua/trea/数据/preview/sales_demo_3months.xlsx
    校验脚本:/Users/xiahua/trea/数据/preview/preview_sales_demo.py(导出首列预览)
"""
from __future__ import annotations

import os
import random
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

random.seed(20261019)  # 固定随机种子,保证每次生成一致

OUT_PATH = Path("/Users/xiahua/trea/数据/preview/sales_demo_3months.xlsx")
OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

CHANNELS = ["线上", "线下", "私域", "直播"]
PRODUCTS = [
    ("SKU-001", "智能手表 Pro", 1899),
    ("SKU-002", "蓝牙耳机 X", 399),
    ("SKU-003", "运动手环", 199),
    ("SKU-004", "笔记本 Air", 5499),
    ("SKU-005", "机械键盘", 699),
    ("SKU-006", "无线鼠标", 159),
    ("SKU-007", "便携充电宝", 129),
    ("SKU-008", "智能音箱", 599),
    ("SKU-009", "降噪耳机", 1299),
    ("SKU-010", "4K 显示器", 2199),
    ("SKU-011", "护眼台灯", 269),
    ("SKU-012", "电动牙刷", 459),
    ("SKU-013", "电饭煲", 399),
    ("SKU-014", "空气净化器", 1899),
    ("SKU-015", "智能门锁", 1599),
    ("SKU-016", "无线吸尘器", 1499),
    ("SKU-017", "咖啡机", 2299),
    ("SKU-018", "电烤箱", 699),
    ("SKU-019", "智能马桶", 4999),
    ("SKU-020", "VR 一体机", 2999),
]
CUSTOMERS = [f"客户{i:03d}" for i in range(1, 51)]


def gen_sales_rows(start: date, days: int) -> list[list]:
    """生成指定天数的销售明细(每天 8~25 单,带跨产品/渠道波动)。"""
    rows: list[list] = []
    for d_off in range(days):
        d = start + timedelta(days=d_off)
        # 周末单量翻倍,渠道占比按周内波动
        weekend = d.weekday() >= 5
        order_count = random.randint(15, 28) if weekend else random.randint(8, 18)
        # 本月环比上行:每天再加一个微调,让趋势看得见
        day_growth = 1.0 + (d_off / days) * 0.15  # 0~15% 增长
        for _ in range(order_count):
            channel = random.choices(
                CHANNELS,
                weights=[35, 25, 20, 20] if weekend else [40, 30, 20, 10],
            )[0]
            sku_id, sku_name, base_price = random.choice(PRODUCTS)
            customer = random.choice(CUSTOMERS)
            # 客单价在基础价附近浮动(-10%~+15%)
            unit = round(base_price * random.uniform(0.9, 1.15) * day_growth, 2)
            qty = random.choice([1, 1, 1, 2, 2, 3])  # 多数买 1 件
            amount = round(unit * qty, 2)
            gross_rate = round(random.uniform(0.20, 0.45), 4)
            gross = round(amount * gross_rate, 2)
            new_cust = random.choice([0, 0, 0, 1, 1])  # 20%~40% 新客
            refund_rate = round(random.uniform(0.0, 0.05), 4)
            refund_amt = round(amount * refund_rate, 2)
            pay_back = round(random.uniform(0.85, 0.98), 4)
            rows.append([
                d.isoformat(), channel, sku_id, sku_name, customer,
                unit, qty, amount, gross_rate, gross, new_cust,
                refund_amt, pay_back,
            ])
    return rows


def main():
    wb = openpyxl.Workbook()

    # ========== Sheet 1: 销售明细 ==========
    ws = wb.active
    ws.title = "销售明细"
    headers = [
        "日期", "渠道", "产品SKU", "产品名称", "客户",
        "客单价(元)", "订单量", "销售额(元)", "毛利率", "毛利额",
        "新客户数", "退款金额(元)", "回款率",
    ]
    ws.append(headers)
    # 加粗表头
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="0E7490")
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")

    # 三个月数据:本月 + 上月 + 上上月,起始日期回推
    today = date(2026, 10, 19)
    this_month_start = date(2026, 10, 1)
    last_month_start = date(2026, 9, 1)
    last_last_start = date(2026, 8, 1)
    rows = []
    rows += gen_sales_rows(last_last_start, 31)   # 8 月 31 天
    rows += gen_sales_rows(last_month_start, 30)   # 9 月 30 天
    rows += gen_sales_rows(this_month_start, 19)   # 10 月 19 天(到目前为止)
    for r in rows:
        ws.append(r)

    # 列宽
    for col_letter, w in zip(
        ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"],
        [12, 10, 12, 22, 10, 12, 10, 14, 10, 12, 10, 14, 10],
    ):
        ws.column_dimensions[col_letter].width = w

    # ========== Sheet 2: 周度KPI ==========
    ws2 = wb.create_sheet("周度KPI")
    kpi_headers = [
        "周次", "起始日期", "结束日期",
        "销售额(元)", "订单量", "客单价(元)", "毛利率",
        "新客户数", "回款率", "环比",
    ]
    ws2.append(kpi_headers)
    for col in range(1, len(kpi_headers) + 1):
        c = ws2.cell(row=1, column=col)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")

    # 13 周数据(从 7 月底开始到 10 月底)
    week_starts = []
    cursor = date(2026, 7, 27)  # 周一开始
    while cursor <= today:
        week_starts.append(cursor)
        cursor += timedelta(days=7)
    prev_amount = None
    for idx, ws_start in enumerate(week_starts, 1):
        ws_end = ws_start + timedelta(days=6)
        # 从前面 rows 里筛本周的销售额/订单
        wk_rows = [r for r in rows if ws_start.isoformat() <= r[0] <= ws_end.isoformat()]
        if not wk_rows:
            continue
        amount = round(sum(r[7] for r in wk_rows), 2)
        orders = sum(r[6] for r in wk_rows)
        unit = round(amount / max(orders, 1), 2)
        gross_rate = round(sum(r[9] for r in wk_rows) / max(amount, 1), 4)
        new_cust = sum(r[10] for r in wk_rows)
        # 回款率:本周的订单平均回款率
        pay_back = round(sum(r[12] for r in wk_rows) / len(wk_rows), 4)
        ring = (
            round((amount - prev_amount) / prev_amount * 100, 2)
            if prev_amount else 0.0
        )
        ws2.append([
            f"W{idx}",
            ws_start.isoformat(),
            ws_end.isoformat(),
            amount, orders, unit, gross_rate, new_cust, pay_back, ring,
        ])
        prev_amount = amount

    for col_letter, w in zip(
        ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"],
        [10, 12, 12, 14, 10, 12, 10, 10, 10, 10],
    ):
        ws2.column_dimensions[col_letter].width = w

    # ========== Sheet 3: 渠道结构 ==========
    ws3 = wb.create_sheet("渠道结构")
    ws3.append(["渠道", "本周销售额(元)", "占比", "上周销售额(元)", "环比"])
    for col in range(1, 6):
        c = ws3.cell(row=1, column=col)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")

    # 本周(2026-10-13 ~ 2026-10-19)与上周(2026-10-06 ~ 2026-10-12)
    this_week_start = date(2026, 10, 13)
    this_week_end = date(2026, 10, 19)
    last_week_start = date(2026, 10, 6)
    last_week_end = date(2026, 10, 12)

    channel_data: dict[str, dict[str, float]] = {}
    for ch in CHANNELS:
        this_amt = sum(
            r[7] for r in rows
            if r[1] == ch and this_week_start.isoformat() <= r[0] <= this_week_end.isoformat()
        )
        last_amt = sum(
            r[7] for r in rows
            if r[1] == ch and last_week_start.isoformat() <= r[0] <= last_week_end.isoformat()
        )
        channel_data[ch] = {"this": this_amt, "last": last_amt}

    total_this = sum(d["this"] for d in channel_data.values())
    total_last = sum(d["last"] for d in channel_data.values())
    for ch, d in channel_data.items():
        ratio = round(d["this"] / total_this * 100, 2) if total_this else 0.0
        ring = round((d["this"] - d["last"]) / d["last"] * 100, 2) if d["last"] else 0.0
        ws3.append([
            ch, round(d["this"], 2), ratio, round(d["last"], 2), ring,
        ])
    # 合计行
    ws3.append([
        "合计", round(total_this, 2), 100.0,
        round(total_last, 2),
        round((total_this - total_last) / total_last * 100, 2) if total_last else 0.0,
    ])

    for col_letter, w in zip(["A", "B", "C", "D", "E"], [10, 18, 10, 18, 10]):
        ws3.column_dimensions[col_letter].width = w

    # ========== Sheet 4: 产品Top10 ==========
    ws4 = wb.create_sheet("产品Top10")
    ws4.append(["排名", "产品SKU", "产品名称", "本周销售额(元)", "占比", "环比"])
    for col in range(1, 7):
        c = ws4.cell(row=1, column=col)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")

    product_this: dict[str, float] = {}
    product_last: dict[str, float] = {}
    for r in rows:
        if this_week_start.isoformat() <= r[0] <= this_week_end.isoformat():
            product_this[r[2]] = product_this.get(r[2], 0) + r[7]
        elif last_week_start.isoformat() <= r[0] <= last_week_end.isoformat():
            product_last[r[2]] = product_last.get(r[2], 0) + r[7]
    sorted_products = sorted(product_this.items(), key=lambda x: -x[1])[:10]
    total_this_prod = sum(product_this.values())
    for rank, (sku, amt) in enumerate(sorted_products, 1):
        # 找产品名
        name = next((p[1] for p in PRODUCTS if p[0] == sku), sku)
        ratio = round(amt / total_this_prod * 100, 2) if total_this_prod else 0.0
        last_amt = product_last.get(sku, 0)
        ring = round((amt - last_amt) / last_amt * 100, 2) if last_amt else 0.0
        ws4.append([rank, sku, name, round(amt, 2), ratio, ring])

    for col_letter, w in zip(["A", "B", "C", "D", "E", "F"], [6, 12, 22, 16, 8, 10]):
        ws4.column_dimensions[col_letter].width = w

    # ========== Sheet 5: 客户Top20 ==========
    ws5 = wb.create_sheet("客户Top20")
    ws5.append(["排名", "客户", "本周销售额(元)", "订单量", "客单价(元)", "跟进阶段"])
    for col in range(1, 7):
        c = ws5.cell(row=1, column=col)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")
    cust_this: dict[str, dict] = {}
    for r in rows:
        if this_week_start.isoformat() <= r[0] <= this_week_end.isoformat():
            d = cust_this.setdefault(r[4], {"amt": 0, "qty": 0})
            d["amt"] += r[7]
            d["qty"] += r[6]
    sorted_cust = sorted(cust_this.items(), key=lambda x: -x[1]["amt"])[:20]
    stages = ["初步接触", "需求确认", "方案报价", "商务谈判", "已签约", "复购跟进"]
    for rank, (cust, d) in enumerate(sorted_cust, 1):
        unit = round(d["amt"] / max(d["qty"], 1), 2)
        ws5.append([rank, cust, round(d["amt"], 2), d["qty"], unit, random.choice(stages)])
    for col_letter, w in zip(["A", "B", "C", "D", "E", "F"], [6, 12, 16, 10, 12, 12]):
        ws5.column_dimensions[col_letter].width = w

    wb.save(OUT_PATH)
    print(f"✅ 已生成:{OUT_PATH}")
    print(f"   销售明细:{len(rows)} 行")
    print(f"   周度KPI:{ws2.max_row - 1} 行")
    print(f"   渠道结构:{ws3.max_row - 1} 行")
    print(f"   产品Top10:{ws4.max_row - 1} 行")
    print(f"   客户Top20:{ws5.max_row - 1} 行")


if __name__ == "__main__":
    main()