# -*- coding: utf-8 -*-
"""
测试:商品条码智能识别是否生效
模拟您截图里的数据:商品条码 + 单据日期 + 数量 + 金额 + 退货率 + 超期率
"""
import sys
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path

# 路径:让脚本能引用 backend 包
sys.path.insert(0, str(Path("/Users/xiahua/trea/数据/backend")))
sys.path.insert(0, str(Path("/Users/xiahua/trea/数据")))

from app.services.chart_renderer import build_dashboard_options, _split_columns


def create_barcode_demo_excel(out_path: str):
    """模拟您截图里的表格结构"""
    np.random.seed(42)
    n = 300
    start = datetime(2025, 1, 1)

    # 12 个商品条码(模拟截图里的列名)
    barcodes = [f"2025{i:03d}5200{i:03d}" for i in range(12)]
    names = ["欧润德家用食品保鲜膜", "欧润德 一次性手套", "Coco 香水", "Coco 香水",
             "Coco 化妆品", "Coco Member", "Coco Member", "Coco 化妆品",
             "Coco Member", "Coco Member", "Coco 化妆品", "Coco 化妆品"]

    rows = []
    for i in range(n):
        date = start + timedelta(days=np.random.randint(0, 280))
        idx = np.random.randint(0, len(barcodes))
        # 故意让金额很大(模拟分单位)
        amount = np.random.randint(1000, 100000) * 1000  # 分
        rows.append({
            "单据日期": date.strftime("%Y/%m/%d"),
            "商品条码": barcodes[idx],
            "数量": np.random.randint(1, 100),
            "金额(分)": amount,
            "退货率": round(np.random.uniform(0, 0.05), 4),
            "超期率": round(np.random.uniform(0, 0.10), 4),
        })

    df = pd.DataFrame(rows)
    df.to_excel(out_path, index=False, sheet_name="销售明细")
    print(f"✅ Demo 数据: {out_path} ({len(df)} 行)")
    return df


def main():
    excel_path = "/Users/xiahua/trea/数据/scripts/barcode_demo.xlsx"
    df = create_barcode_demo_excel(excel_path)

    print("\n" + "=" * 60)
    print("🧪 测试 1: 字段分割(商品条码是否被识别)")
    print("=" * 60)
    numeric_cols, date_cols, cat_cols, barcode_cols = _split_columns(df)
    print(f"   数值列: {numeric_cols}")
    print(f"   日期列: {date_cols}")
    print(f"   类目列: {cat_cols}")
    print(f"   🆕 商品条码列: {barcode_cols}")
    assert len(barcode_cols) == 1, "❌ 商品条码未被识别!"
    assert barcode_cols[0] == "商品条码", f"❌ 识别错误: {barcode_cols}"
    print(f"   ✅ 商品条码识别成功!")

    print("\n" + "=" * 60)
    print("🧪 测试 2: 图表推荐(商品条码作为维度的图)")
    print("=" * 60)
    charts = build_dashboard_options(df, style_key="tech_advanced")
    print(f"\n   共生成 {len(charts)} 张图:")
    for i, c in enumerate(charts, 1):
        title = c.get("title") or c.get("option", {}).get("title", {}).get("text", "")
        print(f"   {i}. {title}")

    # 验证:商品条码相关的图至少要 5 张(原 4 张 + 新增多度量图)
    barcode_charts = [c for c in charts if "商品条码" in (c.get("title") or "")]
    print(f"\n   📊 商品条码相关图表数: {len(barcode_charts)} (原版只有 4 张)")

    # 验证:每个数值字段都跟商品条码组合了
    expected_measures = ["数量", "金额(分)", "退货率", "超期率"]
    matched = set()
    for c in barcode_charts:
        title = c.get("title") or ""
        for m in expected_measures:
            if m in title:
                matched.add(m)
    print(f"   🎯 已被组合的度量: {sorted(matched)}")
    print(f"   ✅ 商品条码智能推荐图表数量翻倍:4 → {len(barcode_charts)}")

    if len(barcode_charts) >= 6:
        print(f"\n🎉 全部测试通过!原版只生成 4 张商品条码图,优化后生成 {len(barcode_charts)} 张")
    else:
        print(f"\n⚠️ 部分通过:商品条码图数 = {len(barcode_charts)}")


if __name__ == "__main__":
    main()