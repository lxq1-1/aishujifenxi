"""初始化内置报告模板。"""
from sqlalchemy.orm import Session

from app.models.report_template import ReportTemplate


# 注意:此处只保存"用户可见"的模板字段(name / category / description / structure)。
# 真正的 Prompt 模板(role + 输出 schema + 占位符注入)由后端
# `app/core/prompt_builder.py` 统一管理,不允许用户在 UI 上编辑。
SEED_TEMPLATES = [
    {
        "name": "年度复盘",
        "category": "年度",
        "description": (
            "面向管理层与业务负责人的年度经营复盘报告，"
            "覆盖执行摘要、年度核心指标同比/环比、内外部归因、问题诊断与下年度行动建议。"
        ),
        "structure": (
            "一、执行摘要\n"
            "二、报告口径与数据来源\n"
            "三、年度核心指标(KPI 达成率/同比/环比)\n"
            "四、业务结构与亮点拆解\n"
            "五、主要问题与归因\n"
            "六、风险提示\n"
            "七、下年度行动建议与里程碑\n"
            "八、附录:关键术语与指标定义"
        ),
    },
    {
        "name": "季度分析",
        "category": "季度",
        "description": (
            "业务季度经营分析报告，聚焦季度环比与年度同比、异常波动归因、"
            "风险与机会盘点，并给出下季度可执行的工作计划与资源诉求。"
        ),
        "structure": (
            "一、执行摘要\n"
            "二、本季关键指标(环比/同比)\n"
            "三、关键业务变化与异常点\n"
            "四、归因分析(内/外部因素)\n"
            "五、风险与机会盘点\n"
            "六、下季度行动计划(Owner/时间/资源)\n"
            "七、附录:口径说明"
        ),
    },
    {
        "name": "述职报告",
        "category": "述职",
        "description": (
            "面向晋升/年度考核场景的个人述职报告，"
            "围绕岗位目标达成、关键项目贡献、能力成长、不足与改进、以及新周期承诺。"
        ),
        "structure": (
            "一、个人简介与岗位职责\n"
            "二、考核周期关键成果(量化指标 + 项目贡献)\n"
            "三、核心项目复盘(目标/过程/结果/复盘)\n"
            "四、能力成长与所获荣誉\n"
            "五、不足反思与改进计划\n"
            "六、下一周期工作承诺与目标设定\n"
            "七、附录:关键业绩证据清单"
        ),
    },
    {
        "name": "数据简报",
        "category": "简报",
        "description": (
            "面向高频阅读场景的一页式数据简报，"
            "用最少的数字呈现核心结论、趋势与三条以内可执行建议。"
        ),
        "structure": (
            "一、本期核心数字(亮点 + 风险各 1~2 个)\n"
            "二、同环比趋势(2~3 个关键指标)\n"
            "三、归因一句话说明\n"
            "四、行动建议(不超过 3 条,含 Owner)"
        ),
    },
    {
        "name": "销售周报",
        "category": "销售",
        "description": (
            "面向销售管理者与一线销售的周度业务复盘报告，"
            "聚焦 GMV/订单/客单价等核心指标周环比与累计达成，"
            "拆解 Top 客户/Top SKU、转化漏斗与退货风险，"
            "并给出下周重点客户跟进、促销活动与资源申请计划。"
        ),
        "structure": (
            "一、本周业绩摘要\n"
            "二、核心销售指标(达成率/周环比/同比)\n"
            "三、渠道与产品结构分析\n"
            "四、Top 客户与重点商机进展\n"
            "五、问题与风险(转化/退货/回款)\n"
            "六、下周行动清单(客户跟进/促销活动/资源申请)\n"
            "七、附录:本周关键数据明细"
        ),
    },
    # ===== 【tools-results-history-persistence 新增】辅助工具专用模板 =====
    # 这 3 条模板用于把工作台"数据校验 / 数据对话 / 数据集对比"的工具结果
    # 落库为历史任务。结构与对应工具的 narration 段落一一对应,LLM 不参与,
    # 由工具 API 直接渲染。Task 模板不应被前端误以为"AI 分析模板"而让用户去选。
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
    # ===== 【pdf-tools-history-persistence 新增】PDF 工具专用模板 =====
    # 4 条模板用于把工作台「AI PDF 处理」工具的产物落库为历史任务。
    # PDF 工具用户场景没有原始 dataset,后端会自动建一个
    # status="PDF工具产物" 的内部锚点 dataset 作为 Task.dataset_id 载体,
    # 该 dataset 不会出现在数据集管理列表里(前端有过滤)。
    {
        "name": "PDF 转 Word 报告",
        "category": "PDF处理",
        "description": (
            "工作台「AI PDF 处理 · PDF 转 Word」工具落库的模板。"
            "展示原始 PDF 与转换产物的元信息,并提供下载链接。"
        ),
        "structure": (
            "一、原文件信息\n"
            "二、转换结果\n"
            "三、下载产物"
        ),
    },
    {
        "name": "PDF 转表格 报告",
        "category": "PDF处理",
        "description": (
            "工作台「AI PDF 处理 · PDF 转表格」工具落库的模板。"
            "展示抽出的表格清单、每个表格的行列数与下载链接。"
        ),
        "structure": (
            "一、原文件信息\n"
            "二、识别到的表格\n"
            "三、AI 辅助说明\n"
            "四、下载产物"
        ),
    },
    {
        "name": "PDF 提取文字 报告",
        "category": "PDF处理",
        "description": (
            "工作台「AI PDF 处理 · PDF 提取文字」工具落库的模板。"
            "展示 PDF 页数、字符数、提取到的全文(可滚动查看)。"
        ),
        "structure": (
            "一、原文件信息\n"
            "二、提取概览\n"
            "三、提取内容"
        ),
    },
    {
        "name": "PDF 压缩 报告",
        "category": "PDF处理",
        "description": (
            "工作台「AI PDF 处理 · PDF 压缩」工具落库的模板。"
            "展示压缩前后体积、压缩比与下载链接。"
        ),
        "structure": (
            "一、原文件信息\n"
            "二、压缩结果\n"
            "三、下载产物"
        ),
    },
    # ===== 【pdf-photo-to-pdf-ocr 新增】照片转 PDF(OCR)模板 =====
    # 承接「AI PDF 处理 · 照片转 PDF」第 5 个 tab 的产物落库。
    # 与上面 4 条共用 status="PDF工具产物" 锚点 dataset 机制。
    {
        "name": "照片转 PDF 报告",
        "category": "PDF处理",
        "description": (
            "工作台「AI PDF 处理 · 照片转 PDF」工具落库的模板。"
            "展示上传照片清单、OCR 识别结果(打印版/手写版)与 PDF 产物下载链接。"
        ),
        "structure": (
            "一、原图信息\n"
            "二、识别结果\n"
            "三、产物下载"
        ),
    },
]


def seed_templates(db: Session) -> None:
    """首次启动写入/升级内置模板。

    策略(upsert):
        - 按 name 查找;不存在则新增(is_builtin=1)。
        - 已存在(通常是历史内置)则刷新 description / structure,
          但保留 id 与 is_builtin 不被覆盖,确保之前用户已有的引用稳定。
        - 内置模板标记 is_builtin=1,前端不渲染删除按钮,后端 DELETE 拒绝。
    """
    for tpl in SEED_TEMPLATES:
        exists = db.query(ReportTemplate).filter(ReportTemplate.name == tpl["name"]).first()
        if exists:
            # 升级路径:覆盖 description / structure / category,使内置更专业
            exists.description = tpl["description"]
            exists.structure = tpl["structure"]
            exists.category = tpl["category"]
            continue
        # 首次写入:内置模板
        db.add(ReportTemplate(**tpl, is_builtin=1))
    db.commit()