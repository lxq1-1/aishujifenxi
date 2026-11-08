"""
生成《AI 数据处理工作台》项目答辩 PPT (.pptx)
- 16 页 16:9 (12.5" × 7.5")
- 主题色:#0e7490 (深青主色) + #22d3ee (强调青)
- 风格:对标项目 PPT 主题 deep_cyan,与产品 UI 一致
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

# ============== 主题色(与项目 deep_cyan 主题一致) ==============
PRIMARY = RGBColor(0x0e, 0x74, 0x90)        # 深青
ACCENT = RGBColor(0x22, 0xd3, 0xee)         # 强调青
WARM = RGBColor(0xfb, 0x92, 0x3c)           # 暖橙(对比强调)
TEXT_DARK = RGBColor(0x1f, 0x29, 0x37)      # 正文
TEXT_MUTED = RGBColor(0x6b, 0x72, 0x80)     # 副文
TEXT_LIGHT = RGBColor(0x94, 0xa3, 0xb8)     # 弱化
SURFACE = RGBColor(0xff, 0xff, 0xff)        # 卡片底
SURFACE_2 = RGBColor(0xf1, 0xf5, 0xf9)      # 次级底
BORDER = RGBColor(0xe2, 0xe8, 0xf0)         # 描边
DARK_BG = RGBColor(0x0c, 0x13, 0x30)        # 深色背景(章节分隔)
GOOD = RGBColor(0x10, 0xb9, 0x81)           # 绿色
BAD = RGBColor(0xef, 0x44, 0x44)            # 红色
WARN = RGBColor(0xd9, 0x77, 0x06)           # 琥珀

# ============== 创建演示稿 16:9 ==============
prs = Presentation()
prs.slide_width = Inches(12.5)
prs.slide_height = Inches(7.5)
BLANK_LAYOUT = prs.slide_layouts[6]


# ============== 通用 helper ==============
def add_slide():
    return prs.slides.add_slide(BLANK_LAYOUT)


def fill_bg(slide, rgb):
    """整页填充底色"""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def add_rect(slide, x, y, w, h, fill_color=None, line_color=None, line_width=None):
    """矩形"""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill_color is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        if line_width:
            shp.line.width = Pt(line_width)
    return shp


def add_rounded_rect(slide, x, y, w, h, fill_color=SURFACE, line_color=BORDER, line_width=0.75):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(line_width)
    else:
        shp.line.fill.background()
    # 调整圆角幅度
    return shp


def add_text(slide, x, y, w, h, text,
             font_size=14, bold=False, color=TEXT_DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name="微软雅黑"):
    """通用文本框"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        f = run.font
        f.name = font_name
        # 同时设置 east-asian 字体,确保中文渲染一致
        rPr = run._r.get_or_add_rPr()
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = etree.SubElement(rPr, qn('a:ea'))
        ea.set('typeface', font_name)
        f.size = Pt(font_size)
        f.bold = bold
        f.color.rgb = color
    return tb


def add_rich_text(slide, x, y, w, h, parts,
                  font_size=14, color=TEXT_DARK, align=PP_ALIGN.LEFT,
                  anchor=MSO_ANCHOR.TOP, font_name="微软雅黑"):
    """富文本:parts = [(text, {bold, color, size}), ...]"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    for i, (txt, opts) in enumerate(parts):
        if opts.get("newline"):
            p = tf.add_paragraph()
            p.alignment = align
            continue
        run = p.add_run()
        run.text = txt
        f = run.font
        f.name = font_name
        rPr = run._r.get_or_add_rPr()
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = etree.SubElement(rPr, qn('a:ea'))
        ea.set('typeface', font_name)
        f.size = Pt(opts.get("size", font_size))
        f.bold = opts.get("bold", False)
        f.color.rgb = opts.get("color", color)
    return tb


def add_top_bar(slide, color=PRIMARY):
    """顶部色条"""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.15))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()


def add_page_header(slide, kicker, title):
    """通用页面 header:kicker + 主标题"""
    add_text(slide, Inches(0.6), Inches(0.4), Inches(11.3), Inches(0.4),
             kicker, font_size=13, bold=True, color=PRIMARY, align=PP_ALIGN.LEFT)
    add_text(slide, Inches(0.6), Inches(0.85), Inches(11.3), Inches(0.9),
             title, font_size=34, bold=True, color=TEXT_DARK, align=PP_ALIGN.LEFT)


def add_page_footer(slide, idx, total, note=""):
    """页脚 + 页码"""
    add_text(slide, Inches(0.6), Inches(7.05), Inches(8), Inches(0.3),
             note, font_size=11, color=TEXT_MUTED, align=PP_ALIGN.LEFT)
    add_text(slide, Inches(10.5), Inches(7.05), Inches(1.4), Inches(0.3),
             f"{idx} / {total}", font_size=11, color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)


def add_lede(slide, x, y, w, h, text, color=TEXT_MUTED, font_size=16):
    """lede 副标题"""
    add_text(slide, x, y, w, h, text, font_size=font_size, color=color, align=PP_ALIGN.LEFT)


# ============== 第 1 页:封面 ==============
def page_cover():
    s = add_slide()
    fill_bg(s, DARK_BG)
    # 装饰渐变圆
    c1 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2), Inches(-2), Inches(8), Inches(8))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(0x1e, 0x3a, 0x8a)
    c1.line.fill.background()
    c2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8), Inches(4), Inches(7), Inches(7))
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(0x0e, 0x74, 0x90)
    c2.line.fill.background()
    # 中部装饰条
    add_rect(s, Inches(0), Inches(3.0), prs.slide_width, Inches(0.12), fill_color=ACCENT)

    # 顶部 kicker
    add_text(s, Inches(0.8), Inches(0.6), Inches(8), Inches(0.4),
             "毕业设计 · 项目答辩", font_size=14, bold=True, color=ACCENT)
    add_text(s, Inches(0.8), Inches(1.0), Inches(8), Inches(0.4),
             "AI 数据处理工作台 · 2026", font_size=13, color=RGBColor(0xcb, 0xd5, 0xe1))

    # 主标题
    add_rich_text(s, Inches(0.8), Inches(3.3), Inches(11), Inches(2.5), [
        ("AI 数据处理工作台", {"size": 56, "bold": True, "color": RGBColor(0xff, 0xff, 0xff)}),
        ("", {"newline": True}),
        ("让 Excel 台账 → AI 报告", {"size": 42, "bold": True, "color": ACCENT}),
    ])

    # 副标题
    add_text(s, Inches(0.8), Inches(5.7), Inches(11), Inches(0.5),
             "一站式 Web 系统 · 数据清洗 · 指标计算 · 图表生成 · AI 报告 · 多格式导出",
             font_size=18, color=RGBColor(0xe2, 0xe8, 0xf0))

    # 底部信息
    add_text(s, Inches(0.8), Inches(6.7), Inches(4), Inches(0.4),
             "答辩人:xxx", font_size=13, color=RGBColor(0xcb, 0xd5, 0xe1))
    add_text(s, Inches(5), Inches(6.7), Inches(4), Inches(0.4),
             "指导老师:xxx", font_size=13, color=RGBColor(0xcb, 0xd5, 0xe1))
    add_text(s, Inches(9), Inches(6.7), Inches(3), Inches(0.4),
             "答辩时间:2026", font_size=13, color=RGBColor(0xcb, 0xd5, 0xe1))


# ============== 第 2 页:目录 ==============
def page_toc():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_text(s, Inches(0.6), Inches(0.4), Inches(11), Inches(0.4),
             "CONTENTS · 答辩提纲", font_size=13, bold=True, color=PRIMARY)
    add_text(s, Inches(0.6), Inches(0.85), Inches(11), Inches(0.9),
             "本次答辩的六个模块", font_size=36, bold=True, color=TEXT_DARK)
    add_lede(s, Inches(0.6), Inches(1.85), Inches(11), Inches(0.4),
             "围绕项目背景 → 技术方案 → 实现细节 → 应用价值展开。")

    items = [
        ("01", "项目背景与需求", "业务痛点、目标用户、产品定位", PRIMARY),
        ("02", "技术选型与架构", "前后端技术栈、系统架构图", PRIMARY),
        ("03", "核心功能实现", "6 个核心模块的工作机制与关键代码", PRIMARY),
        ("04", "关键技术难点", "照片转 PDF、LLM 容错、跨模态对比", WARM),
        ("05", "系统演示", "完整业务流程 · 操作步骤录屏", PRIMARY),
        ("06", "总结与展望", "项目亮点 · 答辩总结 · 后续计划", PRIMARY),
    ]

    # 2 列 × 3 行
    col_w = Inches(5.5)
    row_h = Inches(1.3)
    start_x = Inches(0.6)
    start_y = Inches(2.6)
    gap_x = Inches(0.3)
    gap_y = Inches(0.25)

    for i, (no, title, desc, accent_color) in enumerate(items):
        row = i // 2
        col = i % 2
        x = start_x + (col_w + gap_x) * col
        y = start_y + (row_h + gap_y) * row
        # 卡片
        card = add_rounded_rect(s, x, y, col_w, row_h, fill_color=SURFACE, line_color=BORDER)
        # 序号
        add_text(s, x + Inches(0.2), y + Inches(0.2), Inches(1.0), Inches(0.9),
                 no, font_size=36, bold=True, color=accent_color)
        # 标题
        add_text(s, x + Inches(1.4), y + Inches(0.25), Inches(4), Inches(0.5),
                 title, font_size=18, bold=True, color=TEXT_DARK)
        # 描述
        add_text(s, x + Inches(1.4), y + Inches(0.75), Inches(4), Inches(0.4),
                 desc, font_size=12, color=TEXT_MUTED)

    add_page_footer(s, 2, 16, "答辩时间:8 分钟")


# ============== 章节分隔页 ==============
def page_section(idx, total, section_no, title, lead):
    s = add_slide()
    fill_bg(s, DARK_BG)
    # 装饰网格
    c1 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2), Inches(-2), Inches(8), Inches(8))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(0x1e, 0x3a, 0x8a)
    c1.line.fill.background()
    c2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8), Inches(4), Inches(7), Inches(7))
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(0x0e, 0x74, 0x90)
    c2.line.fill.background()
    # 中部装饰条
    add_rect(s, Inches(0), Inches(3.0), prs.slide_width, Inches(0.12), fill_color=ACCENT)

    add_text(s, Inches(0.8), Inches(1.5), Inches(8), Inches(0.5),
             f"Section · {section_no}", font_size=16, bold=True, color=ACCENT)

    add_rich_text(s, Inches(0.8), Inches(2.3), Inches(11), Inches(2.0), [
        (title, {"size": 64, "bold": True, "color": RGBColor(0xff, 0xff, 0xff)}),
    ])

    add_text(s, Inches(0.8), Inches(4.5), Inches(11), Inches(0.5),
             lead, font_size=18, color=RGBColor(0xe2, 0xe8, 0xf0))

    # 页码
    add_text(s, Inches(10.5), Inches(7.05), Inches(1.4), Inches(0.3),
             f"{idx} / {total}", font_size=11, color=RGBColor(0x94, 0xa3, 0xb8), align=PP_ALIGN.RIGHT)


# ============== 第 4 页:业务痛点 ==============
def page_pain_points():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_page_header(s, "Pain Points · 业务痛点",
                    "行政/运营/销售每天都被重复劳动消耗")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.4),
             "一份 Excel 台账到一份可读报告,中间要经过 5 道手工环节——这正是 AI 能解放的部分。")

    cards = [
        ("📋", "手工清洗数据", "格式不统一、缺失值、重复行——每次都要用 Excel 公式或筛选,几小时起步。"),
        ("📊", "手动算指标", "合计、均值、同环比、TopN——Excel 函数 + 复制粘贴,公式一变全要重来。"),
        ("🖼️", "图表靠手动", "插入 → 选数据 → 调样式 → 截图,改一个数字整张图重做一遍。"),
        ("📝", "写报告耗时", "把数字翻译成业务叙述,述职/年度复盘/季度分析——动辄两三天写一份。"),
        ("📦", "格式分发", "给老板的 PDF、给业务的 Excel、给法务的 Word——同一份内容做 4 遍。"),
        ("🤖", "不会写代码", "想用 AI 却不会调 API,想自动化又不会 Python——非技术人员的真实困境。"),
    ]
    # 3 列 × 2 行
    col_w = Inches(3.7)
    row_h = Inches(2.1)
    start_x = Inches(0.6)
    start_y = Inches(2.5)
    gap_x = Inches(0.2)
    gap_y = Inches(0.2)

    for i, (icon, title, desc) in enumerate(cards):
        row = i // 3
        col = i % 3
        x = start_x + (col_w + gap_x) * col
        y = start_y + (row_h + gap_y) * row
        card = add_rounded_rect(s, x, y, col_w, row_h, fill_color=SURFACE, line_color=BORDER)
        # 图标
        add_text(s, x + Inches(0.2), y + Inches(0.2), Inches(0.7), Inches(0.7),
                 icon, font_size=30, color=TEXT_DARK)
        # 标题
        add_text(s, x + Inches(0.2), y + Inches(0.9), col_w - Inches(0.4), Inches(0.5),
                 title, font_size=18, bold=True, color=TEXT_DARK)
        # 描述
        add_text(s, x + Inches(0.2), y + Inches(1.35), col_w - Inches(0.4), Inches(0.7),
                 desc, font_size=12, color=TEXT_MUTED)

    add_page_footer(s, 4, 16, "痛点共性:重复、低效、易错")


# ============== 第 5 页:目标用户 ==============
def page_target_users():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_page_header(s, "Target · 目标用户", "谁会来用这套系统?")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.5),
             "三类典型用户,共同特征:每周至少花 4 小时在 Excel 上、不写代码、需要向上汇报。")

    users = [
        ("运营人员", PRIMARY,
         "活动复盘 / 渠道分析",
         '"活动结束后,我得出一份给老板看的复盘,等 AI 出数+写文,我自己再补两句就行。"',
         "常用模板:数据简报、季度分析"),
        ("销售人员", WARM,
         "销售周报 / 客户跟进",
         '"周一/周五都要把本周的客户跟进数、新签数整理一遍——这事 AI 帮我做了,我就能专心谈客户。"',
         "常用模板:销售周报、数据简报"),
        ("HR / 行政", ACCENT,
         "人事台账 / 述职材料",
         '"年底述职要数据 + 项目 + 反思三块结合,以前我从周三写到周日,现在一小时能出初稿。"',
         "常用模板:述职报告、年度复盘"),
    ]

    col_w = Inches(3.7)
    row_h = Inches(4.0)
    start_x = Inches(0.6)
    start_y = Inches(2.6)
    gap_x = Inches(0.2)

    for i, (role, color, use_case, quote, templates) in enumerate(users):
        x = start_x + (col_w + gap_x) * i
        y = start_y
        # 卡片
        card = add_rounded_rect(s, x, y, col_w, row_h, fill_color=SURFACE, line_color=BORDER)
        # 顶部彩色条
        add_rect(s, x, y, col_w, Inches(0.15), fill_color=color)
        # 角色标签
        tag = add_rounded_rect(s, x + Inches(0.2), y + Inches(0.3), Inches(1.3), Inches(0.35),
                                fill_color=color)
        add_text(s, x + Inches(0.2), y + Inches(0.32), Inches(1.3), Inches(0.35),
                 role, font_size=12, bold=True, color=RGBColor(0xff, 0xff, 0xff), align=PP_ALIGN.CENTER)
        # 场景
        add_text(s, x + Inches(0.2), y + Inches(0.85), col_w - Inches(0.4), Inches(0.5),
                 use_case, font_size=18, bold=True, color=TEXT_DARK)
        # 引言
        add_text(s, x + Inches(0.2), y + Inches(1.5), col_w - Inches(0.4), Inches(1.5),
                 quote, font_size=13, color=TEXT_MUTED)
        # 模板
        add_text(s, x + Inches(0.2), y + Inches(3.4), col_w - Inches(0.4), Inches(0.4),
                 f"📌 {templates}", font_size=11, color=color)

    add_page_footer(s, 5, 16, "产品定位:零代码 · 一站式 · 中文友好")


# ============== 第 6 页:核心价值 ==============
def page_core_value():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_page_header(s, "Value · 核心价值",
                    "一个工作流,把 Excel 变成 4 份汇报材料")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.5),
             "从「上传 Excel」到「拿到 4 种格式的成品报告」,平均 1 分 30 秒,耗时相比手工降低 80%。")

    kpis = [
        ("80", "%", "报告耗时降低", "2~3 天 → 30 分钟出初稿", PRIMARY, ACCENT),
        ("4", "", "导出格式", "PDF / Word / Excel / PPT", PRIMARY, ACCENT),
        ("7", "", "辅助工具", "覆盖 80% 办公场景", PRIMARY, ACCENT),
        ("12", "", "内置报告模板", "4 类业务 × 8 个垂直场景", PRIMARY, ACCENT),
    ]

    col_w = Inches(2.7)
    row_h = Inches(2.5)
    start_x = Inches(0.7)
    start_y = Inches(2.7)
    gap_x = Inches(0.2)

    for i, (num, unit, label, sub, c1, c2) in enumerate(kpis):
        x = start_x + (col_w + gap_x) * i
        y = start_y
        card = add_rounded_rect(s, x, y, col_w, row_h, fill_color=SURFACE, line_color=BORDER)
        # 大数字
        add_rich_text(s, x + Inches(0.2), y + Inches(0.4), col_w - Inches(0.4), Inches(1.2), [
            (num, {"size": 60, "bold": True, "color": c1}),
            (unit, {"size": 36, "bold": True, "color": c2}),
        ])
        # 标签
        add_text(s, x + Inches(0.2), y + Inches(1.5), col_w - Inches(0.4), Inches(0.4),
                 label, font_size=16, bold=True, color=TEXT_DARK)
        # 副
        add_text(s, x + Inches(0.2), y + Inches(1.9), col_w - Inches(0.4), Inches(0.4),
                 sub, font_size=12, color=TEXT_MUTED)
        # 底部装饰条
        add_rect(s, x, y + row_h - Inches(0.1), col_w, Inches(0.1), fill_color=c2)

    add_page_footer(s, 6, 16, "数据来源:5 位种子用户实测 · 平均节省 4 小时/周")


# ============== 第 8 页:技术栈 ==============
def page_tech_stack():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_page_header(s, "Stack · 技术栈",
                    "前后端 + AI 三层栈,职责清晰")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.5),
             "选型原则:易部署 · 生态成熟 · 适合非结构化数据处理。所有计算都在后端,前端只做展示。")

    rows = [
        ("前端框架", "Vue 3 + Vite + Pinia + Element Plus + ECharts", "组件库丰富,生态成熟,ECharts 适合做复杂图表"),
        ("后端框架", "FastAPI + SQLAlchemy + Pydantic", "异步高性能、自动生成 OpenAPI 文档,接口即文档"),
        ("数据处理", "Pandas + openpyxl + pdfplumber + pypdfium2", "表格 → 结构化数据;PDF → 文本/表格抽取"),
        ("数据库", "SQLite (开发) / 可换 PostgreSQL (生产)", "零配置启动,SQLAlchemy 抽象层可平滑切换"),
        ("LLM 接入", "OpenAI Python SDK (gpt-4o-mini / qwen / deepseek)", "OpenAI 协议已成事实标准,各家厂商都兼容"),
        ("导出引擎", "reportlab + python-docx + openpyxl + python-pptx", "纯 Python,无 Office 依赖,服务器端直出"),
        ("图表渲染", "matplotlib(导出用 PNG) + ECharts(前端交互)", "导出走矢量→位图,前端保留交互体验"),
    ]

    # 表格头
    table_y = Inches(2.5)
    head_h = Inches(0.45)
    add_rect(s, Inches(0.6), table_y, Inches(11.3), head_h, fill_color=PRIMARY)
    add_text(s, Inches(0.7), table_y + Inches(0.08), Inches(2), Inches(0.3),
             "分层", font_size=13, bold=True, color=RGBColor(0xff, 0xff, 0xff))
    add_text(s, Inches(2.8), table_y + Inches(0.08), Inches(5), Inches(0.3),
             "技术 / 库", font_size=13, bold=True, color=RGBColor(0xff, 0xff, 0xff))
    add_text(s, Inches(8), table_y + Inches(0.08), Inches(4), Inches(0.3),
             "选型理由", font_size=13, bold=True, color=RGBColor(0xff, 0xff, 0xff))

    # 行
    row_h = Inches(0.5)
    for i, (l1, l2, l3) in enumerate(rows):
        y = table_y + head_h + row_h * i
        # 隔行底色
        if i % 2 == 0:
            add_rect(s, Inches(0.6), y, Inches(11.3), row_h, fill_color=SURFACE_2)
        else:
            add_rect(s, Inches(0.6), y, Inches(11.3), row_h, fill_color=SURFACE)
        # 下划线
        add_rect(s, Inches(0.6), y + row_h - Inches(0.01), Inches(11.3), Inches(0.01), fill_color=BORDER)
        add_text(s, Inches(0.7), y + Inches(0.1), Inches(2), Inches(0.3),
                 l1, font_size=13, bold=True, color=PRIMARY)
        add_text(s, Inches(2.8), y + Inches(0.1), Inches(5), Inches(0.3),
                 l2, font_size=12, color=TEXT_DARK)
        add_text(s, Inches(8), y + Inches(0.1), Inches(4), Inches(0.3),
                 l3, font_size=11, color=TEXT_MUTED)

    add_page_footer(s, 8, 16, "所有依赖 24 项 · 锁定版本号 · 见 requirements.txt")


# ============== 第 9 页:系统架构 ==============
def page_architecture():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_page_header(s, "Architecture · 系统架构",
                    "前端薄、后端厚、LLM 仅做文本")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.5),
             "分层架构,职责分明:计算归 Pandas、文本归 LLM、渲染归前端——任何一个环节挂掉,系统都还能跑。")

    # 3 层
    layers = [
        ("前端层", [
            ("🖥️", "Vue 3", "6 个页面 / SPA"),
            ("📊", "ECharts", "图表渲染"),
            ("📦", "Pinia", "工作台状态"),
            ("🔗", "Axios", "统一 HTTP"),
        ], False),
        ("后端层", [
            ("⚡", "FastAPI", "路由 / API"),
            ("🐼", "Pandas Engine", "数据清洗 / 指标 / 图表"),
            ("🧠", "AI Service", "报告生成 / 摘要 / 对话"),
            ("📑", "PDF Service", "PDF 解析 / 转换 / 抽取"),
        ], True),
        ("基础设施层", [
            ("💾", "SQLite", "数据集 / 任务 / 模板"),
            ("📁", "文件系统", "uploads / reports"),
            ("🤖", "OpenAI 兼容 API", "LLM 网关"),
            ("📤", "导出引擎", "PDF / Word / Excel / PPT"),
        ], False),
    ]

    cell_w = Inches(2.65)
    cell_h = Inches(1.1)
    label_w = Inches(1.4)
    layer_h = Inches(1.5)
    start_x = Inches(0.6)
    start_y = Inches(2.5)
    gap_y = Inches(0.3)

    for i, (label, cells, highlight) in enumerate(layers):
        y = start_y + (layer_h + gap_y) * i
        # 左侧 label
        label_box = add_rounded_rect(s, start_x, y, label_w, layer_h,
                                      fill_color=PRIMARY if highlight else SURFACE_2)
        add_text(s, start_x, y + Inches(0.4), label_w, Inches(0.7),
                 label, font_size=14, bold=True,
                 color=RGBColor(0xff, 0xff, 0xff) if highlight else TEXT_DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 右侧 4 个 cell
        for j, (icon, title, sub) in enumerate(cells):
            x = start_x + label_w + Inches(0.2) + (cell_w + Inches(0.1)) * j
            cy = y + (layer_h - cell_h) / 2
            fill = SURFACE if highlight else SURFACE
            line = PRIMARY if highlight else BORDER
            cw = add_rounded_rect(s, x, cy, cell_w, cell_h, fill_color=fill, line_color=line)
            if highlight:
                # 顶部高亮条
                add_rect(s, x, cy, cell_w, Inches(0.08), fill_color=PRIMARY)
            add_text(s, x + Inches(0.1), cy + Inches(0.15), Inches(0.5), Inches(0.4),
                     icon, font_size=20)
            add_text(s, x + Inches(0.1), cy + Inches(0.55), cell_w - Inches(0.2), Inches(0.3),
                     title, font_size=14, bold=True, color=TEXT_DARK)
            add_text(s, x + Inches(0.1), cy + Inches(0.85), cell_w - Inches(0.2), Inches(0.3),
                     sub, font_size=10, color=TEXT_MUTED)

    add_page_footer(s, 9, 16, "前端 6 页面 / 后端 7 模块 / 3 个 SQLAlchemy 模型")


# ============== 第 11 页:6 个核心模块 ==============
def page_modules():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_page_header(s, "Modules · 6 大模块",
                    "从数据输入到报告输出,全链路打通")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.5),
             "每个模块独立可测、可降级、可替换。任何一个出问题,系统都不卡死。")

    modules = [
        ("📂", "M1 · 数据接入", "数据集管理", "支持 Excel/CSV 多 sheet 合并上传;预览、重命名、删除、状态跟踪。"),
        ("🧼", "M2 · 数据处理", "清洗 & 指标", "5 种清洗规则 + 6 种填充策略;30+ 指标自动计算,中文字段自动归一。"),
        ("🤖", "M3 · AI 分析", "智能报告生成", "12 套模板;Prompt 模板与业务结构分离;LLM 失败自动降级到规则。"),
        ("🛠️", "M4 · 辅助工具", "7 类办公工具", "合并 / 校验 / 对话 / 对比 / 改图 / 清洗 / PDF 处理——覆盖完整链路。"),
        ("📊", "M5 · 可视化", "图表 & 看板", "前置生成 8~N 张看板;自然语言改图;3 套风格按需切换。"),
        ("📤", "M6 · 多格式导出", "PDF / Word / Excel / PPT", "16:9 PPT 自动分章节;图表真实嵌入;4 套主题色可选。"),
    ]

    col_w = Inches(3.7)
    row_h = Inches(1.85)
    start_x = Inches(0.6)
    start_y = Inches(2.5)
    gap_x = Inches(0.2)
    gap_y = Inches(0.2)

    for i, (icon, tag, title, desc) in enumerate(modules):
        row = i // 3
        col = i % 3
        x = start_x + (col_w + gap_x) * col
        y = start_y + (row_h + gap_y) * row
        card = add_rounded_rect(s, x, y, col_w, row_h, fill_color=SURFACE, line_color=BORDER)
        # 图标
        add_text(s, x + Inches(0.2), y + Inches(0.2), Inches(0.7), Inches(0.7),
                 icon, font_size=28)
        # tag
        add_text(s, x + Inches(0.85), y + Inches(0.3), col_w - Inches(1), Inches(0.3),
                 tag, font_size=11, bold=True, color=PRIMARY)
        # 标题
        add_text(s, x + Inches(0.2), y + Inches(0.85), col_w - Inches(0.4), Inches(0.4),
                 title, font_size=17, bold=True, color=TEXT_DARK)
        # 描述
        add_text(s, x + Inches(0.2), y + Inches(1.25), col_w - Inches(0.4), Inches(0.6),
                 desc, font_size=12, color=TEXT_MUTED)

    add_page_footer(s, 11, 16, "6 个模块 · 39 个 API · 6 个页面")


# ============== 第 13 页:三大难点 ==============
def page_challenges():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, WARM)
    add_page_header(s, "Challenges · 三大难点",
                    "工程上最有挑战的三个问题")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.5),
             "它们不是「能不能实现」的问题,而是「实现得稳不稳、容错容得够不够」的问题。")

    # 说明:描述文本框高 0.55" / 11pt,最多容 2 行(≈2240px),
    # 三段文案都按「只删冗余、不删技术关键词」的原则控制在 2 行内。
    challenges = [
        ("🖼", "难点一:照片转 PDF · 图片也能变成可搜索文档",
         WARM,
         "手机拍的合同、扫描件没有文本层,常规抽取只能拿到空白。我们用视觉模型做版面 OCR 拿到 title/role/blocks 结构化结果,再用 reportlab 按 Word 风格重排成白底黑字,印章手写原样保留,并叠一层不可见文字层——既能看原件,又能搜索复制。",
         "效果:照片 → 白底可搜索 PDF"),
        ("🛡️", "难点二:LLM 容错 · 让系统在任何情况下都能交付",
         PRIMARY,
         "API 限流、模型超时、JSON 解析失败、Reasoning 模型把思考链塞进正文——任何一个都会让用户看到空白。我们做了三层兜底:(1) JSON 失败 → 自由模式重试;(2) 自由模式失败 → 规则化叙述 fallback;(3) 章节漏图 → Pandas 兜底补图。另用 strip_thinking() 递归剥离思考链。",
         "效果:AI 不可达时仍能交付报告"),
        ("📜", "难点三:跨模态对比 · 表格 vs PDF vs Word 都能比",
         GOOD,
         "业务上经常出现「两份合同差异」的场景,但合同可能是 Excel、PDF、Word 不同形态。我们设计了 3 层路由:(1) contract_compare 模式 → 12 类条款 × 谈判建议;(2) 任一边非结构化 → LLM 读原文语义对比;(3) 双结构化 → Pandas 4 种算子对比。",
         "效果:任意两份材料都能对比"),
    ]

    card_w = Inches(11.3)
    card_h = Inches(1.35)
    start_x = Inches(0.6)
    start_y = Inches(2.55)
    gap_y = Inches(0.18)

    for i, (icon, title, color, desc, effect) in enumerate(challenges):
        x = start_x
        y = start_y + (card_h + gap_y) * i
        card = add_rounded_rect(s, x, y, card_w, card_h, fill_color=SURFACE, line_color=BORDER)
        # 左侧图标区
        icon_box = add_rect(s, x, y, Inches(1.2), card_h, fill_color=color)
        add_text(s, x, y + Inches(0.3), Inches(1.2), Inches(0.7),
                 icon, font_size=36, color=RGBColor(0xff, 0xff, 0xff), align=PP_ALIGN.CENTER)
        # 标题
        add_text(s, x + Inches(1.5), y + Inches(0.15), card_w - Inches(1.7), Inches(0.4),
                 title, font_size=16, bold=True, color=TEXT_DARK)
        # 描述
        add_text(s, x + Inches(1.5), y + Inches(0.55), card_w - Inches(1.7), Inches(0.55),
                 desc, font_size=11, color=TEXT_MUTED)
        # 效果标签
        tag = add_rounded_rect(s, x + Inches(1.5), y + Inches(1.05), Inches(3.5), Inches(0.25),
                                fill_color=SURFACE_2)
        add_text(s, x + Inches(1.5), y + Inches(1.05), Inches(3.5), Inches(0.25),
                 effect, font_size=10, bold=True, color=color, align=PP_ALIGN.CENTER)

    add_page_footer(s, 13, 16, "设计原则:数据归 Pandas · 文本归 LLM · 任何环节都可降级")


# ============== 第 15 页:演示流程 + 总结 ==============
def page_demo_summary():
    s = add_slide()
    fill_bg(s, RGBColor(0xff, 0xff, 0xff))
    add_top_bar(s, ACCENT)
    add_page_header(s, "Demo · 系统演示",
                    "完整业务流程 · 6 步走完一遍")
    add_lede(s, Inches(0.6), Inches(1.8), Inches(11), Inches(0.5),
             "从打开浏览器到拿到成品报告,平均 1 分 30 秒。接下来我会现场演示。")

    # 6 步
    steps = [
        ("1", "上传 Excel", "拖拽到「数据集」页,\n自动识别多 sheet。", "~3 秒"),
        ("2", "选模板", "数据集选 → 模板选 →\n开始 AI 分析。", "~5 秒"),
        ("3", "AI 生成", "后台清洗 → 算指标 →\nLLM 写报告 → 出图表。", "~30 秒"),
        ("4", "报告预览", "正文 / 摘要 / 图表 /\n关键指标 / 异常批注。", "~10 秒"),
        ("5", "多格式导出", "PDF / Word / Excel / PPT\n任选,4 套主题色。", "~15 秒"),
        ("6", "历史归档", "报告落到「历史任务」,\n可重新分析 / 删除。", "实时"),
    ]

    step_w = Inches(1.9)
    step_h = Inches(1.6)
    start_x = Inches(0.6)
    start_y = Inches(2.55)
    gap_x = Inches(0.1)

    for i, (no, title, desc, time) in enumerate(steps):
        x = start_x + (step_w + gap_x) * i
        y = start_y
        card = add_rounded_rect(s, x, y, step_w, step_h, fill_color=SURFACE, line_color=BORDER)
        # 序号圆
        circle = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                     x + step_w / 2 - Inches(0.3), y - Inches(0.25),
                                     Inches(0.6), Inches(0.6))
        circle.fill.solid()
        circle.fill.fore_color.rgb = PRIMARY
        circle.line.color.rgb = RGBColor(0xff, 0xff, 0xff)
        circle.line.width = Pt(2)
        add_text(s, x + step_w / 2 - Inches(0.3), y - Inches(0.18),
                 Inches(0.6), Inches(0.45),
                 no, font_size=20, bold=True, color=RGBColor(0xff, 0xff, 0xff),
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 标题
        add_text(s, x, y + Inches(0.45), step_w, Inches(0.4),
                 title, font_size=14, bold=True, color=TEXT_DARK, align=PP_ALIGN.CENTER)
        # 描述
        add_text(s, x + Inches(0.1), y + Inches(0.85), step_w - Inches(0.2), Inches(0.6),
                 desc, font_size=10, color=TEXT_MUTED, align=PP_ALIGN.CENTER)
        # 时间标签
        add_text(s, x, y + step_h - Inches(0.3), step_w, Inches(0.3),
                 time, font_size=10, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    # 底部双栏:亮点 + 展望
    col_y = Inches(4.4)
    col_h = Inches(2.4)
    col_w = Inches(5.5)

    # 亮点
    hl_card = add_rounded_rect(s, Inches(0.6), col_y, col_w, col_h, fill_color=SURFACE, line_color=GOOD, line_width=2)
    add_rect(s, Inches(0.6), col_y, col_w, Inches(0.12), fill_color=GOOD)
    add_text(s, Inches(0.8), col_y + Inches(0.25), col_w - Inches(0.4), Inches(0.4),
             "✅ 项目亮点", font_size=16, bold=True, color=GOOD)
    hl_items = [
        "零代码 · 全中文 · 上传即用",
        "照片转 PDF + LLM 容错 · 跨模态对比",
        "7 个辅助工具覆盖完整办公链路",
        "4 种导出格式 + 4 套主题色",
        "AI 不可达时系统仍能交付",
    ]
    for i, item in enumerate(hl_items):
        add_text(s, Inches(0.9), col_y + Inches(0.7 + 0.3 * i), col_w - Inches(0.4), Inches(0.3),
                 f"• {item}", font_size=12, color=TEXT_DARK)

    # 展望
    fc_card = add_rounded_rect(s, Inches(6.4), col_y, col_w, col_h, fill_color=SURFACE, line_color=WARN, line_width=2)
    add_rect(s, Inches(6.4), col_y, col_w, Inches(0.12), fill_color=WARN)
    add_text(s, Inches(6.6), col_y + Inches(0.25), col_w - Inches(0.4), Inches(0.4),
             "📈 后续展望", font_size=16, bold=True, color=WARN)
    fc_items = [
        "知识库(RAG)接入 · 引用内部规范",
        "多用户协同 · 团队共享数据集与模板",
        "定时任务 · 周报/月报自动生成",
        "飞书 / 钉钉 / 邮件集成 · 推送报告",
        "移动端适配 · 随时随地查看",
    ]
    for i, item in enumerate(fc_items):
        add_text(s, Inches(6.7), col_y + Inches(0.7 + 0.3 * i), col_w - Inches(0.4), Inches(0.3),
                 f"• {item}", font_size=12, color=TEXT_DARK)

    add_page_footer(s, 15, 16, "演示完毕,准备进入问答环节")


# ============== 第 16 页:致谢 ==============
def page_thanks():
    s = add_slide()
    fill_bg(s, DARK_BG)
    # 装饰
    c1 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2), Inches(-2), Inches(8), Inches(8))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(0x1e, 0x3a, 0x8a)
    c1.line.fill.background()
    c2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8), Inches(4), Inches(7), Inches(7))
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(0x0e, 0x74, 0x90)
    c2.line.fill.background()
    # 装饰条
    add_rect(s, Inches(0), Inches(3.5), prs.slide_width, Inches(0.12), fill_color=ACCENT)

    # 主标题
    add_text(s, Inches(0.8), Inches(2.4), Inches(11), Inches(1.5),
             "谢谢聆听", font_size=84, bold=True, color=RGBColor(0xff, 0xff, 0xff),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.8), Inches(4.0), Inches(11), Inches(0.6),
             "Welcome Questions & Suggestions",
             font_size=22, color=ACCENT, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.8), Inches(4.6), Inches(11), Inches(0.6),
             "欢迎各位老师批评指正 · 期待您的宝贵意见",
             font_size=18, color=RGBColor(0xcb, 0xd5, 0xe1), align=PP_ALIGN.CENTER)

    # 底部信息
    info_y = Inches(6.0)
    items = [
        ("项目仓库", "本地代码仓库"),
        ("演示地址", "http://localhost:5173"),
        ("答辩时间", "2026 年"),
    ]
    item_w = Inches(3.5)
    start_x = Inches(1.0)
    gap_x = Inches(0.5)
    for i, (label, value) in enumerate(items):
        x = start_x + (item_w + gap_x) * i
        add_text(s, x, info_y, item_w, Inches(0.35),
                 label, font_size=12, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        add_text(s, x, info_y + Inches(0.35), item_w, Inches(0.35),
                 value, font_size=14, color=RGBColor(0xe2, 0xe8, 0xf0), align=PP_ALIGN.CENTER)

    # 页码
    add_text(s, Inches(10.5), Inches(7.05), Inches(1.4), Inches(0.3),
             "16 / 16", font_size=11, color=RGBColor(0x94, 0xa3, 0xb8), align=PP_ALIGN.RIGHT)


# ============== 主流程 ==============
def build():
    page_cover()                                          # 1 封面
    page_toc()                                             # 2 目录
    page_section(3, 16, "01", "项目背景与需求",          # 3 章节 01
                 "业务场景是什么?用户痛点在哪里?我们如何定位产品?")
    page_pain_points()                                     # 4 痛点
    page_target_users()                                    # 5 用户
    page_core_value()                                      # 6 价值
    page_section(7, 16, "02", "技术选型与架构",          # 7 章节 02
                 "用什么技术栈?为什么这样分?系统架构如何运转?")
    page_tech_stack()                                      # 8 技术栈
    page_architecture()                                    # 9 架构
    page_section(10, 16, "03", "核心功能实现",           # 10 章节 03
                 "6 个核心模块是怎么工作的?数据流如何贯穿?")
    page_modules()                                         # 11 模块
    page_section(12, 16, "04", "关键技术难点",           # 12 章节 04
                 "三个最有挑战的工程问题,我们是怎么解决的。")
    page_challenges()                                      # 13 难点
    page_section(14, 16, "05", "系统演示",               # 14 章节 05
                 "从「上传 Excel」到「下载 PPT」,完整跑一遍。")
    page_demo_summary()                                    # 15 演示 + 总结
    page_thanks()                                          # 16 致谢

    out = "/Users/xiahua/trea/数据/答辩PPT/AI数据处理工作台-项目答辩.pptx"
    prs.save(out)
    print(f"✔ 已生成:{out}")
    print(f"✔ 总页数:{len(prs.slides)}")


if __name__ == "__main__":
    build()
