/**
 * 工作台草稿 store：暂存当前数据集、模板、清洗规则。
 * 跨页面（如模板页）选择的模板也写入此处。
 */
import { defineStore } from 'pinia'
import { ref } from 'vue'

export const useWorkbenchStore = defineStore('workbench', () => {
  const draft = ref({
    datasetId: null,
    datasetName: '',
    // ——【dataset-multisheet-merge 新增】——
    // 当前数据集的工作表明细:[{name, rows, cols}, ...],由预览接口拉取后写入,
    // 用于在工作台顶部展示"📑 包含 N 个工作表"标签与悬浮明细。
    // 老数据集(没拉过 sheets)保持空数组,UI 不显示工作表相关元素,行为兼容。
    datasetSheets: [],
    // ——【workbench-per-sheet-preview 新增】——
    // 当前预览激活的 sheet 名。'' 表示"全部 sheet 合并预览"(默认/AI 链路用);
    // 其它值表示"只看这张 sheet 的真实前 N 行"(工作台顶部 tab 切换)。
    // 切换数据集时会被重置为 '',切到具体 sheet 时由 onSheetChange 写入。
    activeSheet: '',
    templateId: null,
    templateName: '',
    cleaningRules: {
      drop_empty_rows: true,
      drop_duplicates: true,
      unify_date: true,
      // 异常值标记会追加 *_is_outlier TRUE/FALSE 列,默认关闭避免干扰
      mark_anomaly: false,
    },
    useKnowledgeBase: false,
    kbIds: [],
    // ——【workbench-cleansing-missing-data 新增】——
    // 第 5 项清洗规则开关(默认关闭),以及二级参数(策略 / 作用范围)
    fill_missing: false,
    fillMissingPolicy: { strategy: 'mean', scope: 'numeric' },
    // ——【clean-keep-formulas 新增】——
    // 「保留公式回写」开关:勾选后清洗结果用 openpyxl 写回原工作簿,
    // 保留未改动单元格的公式 / 样式 / 多 sheet 布局(默认关闭,走 pandas 纯值导出)
    keep_formulas: false,
    // 上一次清洗的差量统计预览,初始为 null(显示 —),跑过一次后由 onCleanSubmit 写入
    previewCleanedStats: null,
    // ——【clean-ai-derived-columns-pivot 新增】派生列与透视汇总草稿 ——
    // 清洗弹窗内所选数据集的字段结构:[{ name, row_count, columns:[{name,type,non_null,unique_count}] }]
    // 由 GET /tools/schema 拉取,用于派生列工作表下拉与透视字段候选
    cleanSchema: [],
    // AI 推荐的派生列,每条 { sheet_name, name, expression, description, selected };
    // selected 默认 true,允许用户取消勾选 / 改名 / 改公式
    derivedSuggestions: [],
    // 用户手动添加的派生列,每条 { sheet_name, name, expression }
    manualDerivedColumns: [],
    // 「AI 推荐派生列」按钮 loading(文案同时切到「AI 分析中…」)
    suggestColumnsLoading: false,
    // 透视汇总配置数组(支持多张并行)
    // 每条 { id, enabled, sheet_name, row_dim, col_dim, value, agg, include_margin, _origin_suggestion }
    // id 由 applyPivotSuggestion 生成("pivot_<递增序号>"),用于在 UI 上区分每条配置
    pivotConfigs: [],
    // AI 推荐的透视汇总方案,每条 { sheet_name, row_dim, col_dim, value, agg,
    // include_margin, description, value_is_derived };
    // 用户点"应用"时 push 一条到 pivotConfigs,多条可叠加
    pivotSuggestions: [],
    // ——【workbench-analysis-intent-input 新增】——
    // 用户在「开始 AI 分析」前临时输入的分析重点(自由文本 ≤500 字);
    // 空字符串表示本次未填重点(等同跳过)。
    analysisIntent: '',
    // 「本次跳过、直接生成报告」复选框状态;勾选时 textarea 被禁用 + 内容清空
    skipIntent: false,
    // ===== 【chart-edit-prebuilt-dashboard 新增】预选图表草稿 =====
    // - chartStyle: 当前弹窗选中的图表风格 key,默认 tech_advanced
    // - dashboardCharts: 兼容旧 API 返回的 flat 数组 [{title, option, style_key}, ...];
    //   新 API 返回 sheets 分组时也兼容。
    // - dashboardGroups: 按 sheet 分组的看板 [{sheet, charts: [...]}, ...]
    // - selectedChartKeys: 被用户勾选"加入报表"的图表 key 集合(Set,key 为 "<sheet>__<idx>")
    // - keepSelectionOnRegen: 重新生成看板时是否保留勾选状态,默认 false
    // - aiCurate: 【chart-render-smart-2026-10 新增】是否走"AI 智能筛选"。
    //   true → 走混合出图:本地骨架 + LLM 精选 Top3 + 标题润色(LLM 不可用时
    //   自动降级为本地精选),用于解决"图表过多且很多空"的痛点。
    //   false → 维持旧路径:本地规则出图 2~6 张。
    // - aiCurateTarget: aiCurate=true 时,每个 sheet 保留的图表数,默认 3。
    chartStyle: "tech_advanced",
    dashboardCharts: [],
    dashboardGroups: [],
    selectedChartKeys: new Set(),
    keepSelectionOnRegen: false,
    aiCurate: true,
    aiCurateTarget: 3,
  })

  function setDataset(d) {
    draft.value.datasetId = d.id
    draft.value.datasetName = d.name
    // 切换数据集时清掉旧 sheets,避免误显
    draft.value.datasetSheets = []
    // 【workbench-per-sheet-preview 新增】换数据集时重置 sheet 选择为"合并预览",
    // 避免旧的 activeSheet 不属于新数据集导致预览接口返回 4xx。
    draft.value.activeSheet = ''
    // 【chart-edit-prebuilt-dashboard 新增】换数据集时清空看板草稿,
    // 避免旧看板 / 旧勾选污染新数据集。
    draft.value.dashboardCharts = []
    draft.value.dashboardGroups = []
    draft.value.selectedChartKeys = new Set()
  }

  /**
   * 【dataset-deleted-not-selectable 新增】清空当前选中的数据集。
   *
   * 背景:store 是模块级单例,跨路由跳转存活。用户可能这样操作:
   *   在工作台选中数据集 A → 切到「数据集」页把 A 删掉 → 切回工作台。
   * 此时 draft.datasetId 仍指向已被删除的 A,工作台会去请求
   * /datasets/{A}/preview 拿 404,弹红色"数据集不存在"的报错,
   * 且"开始 AI 分析"仍会把一个已删除的 dataset_id 提交给后端。
   *
   * 这里提供"只清数据集、不动模板/看板/清洗规则"的最小粒度清理,
   * 由调用方(工作台 loadDatasets)在发现当前 id 已不在可见列表时调用。
   */
  function clearDataset() {
    draft.value.datasetId = null
    draft.value.datasetName = ''
    draft.value.datasetSheets = []
    draft.value.activeSheet = ''
  }

  /**
   * 【dataset-multisheet-merge 新增】写入当前数据集的工作表明细。
   * 调用时机:工作台顶部数据集卡 loadPreview() 拉到 preview.sheets 后调用。
   * 兼容老接口(没 sheets 字段)传空数组即可,UI 不显示工作表相关元素。
   */
  function setDatasetSheets(sheets) {
    draft.value.datasetSheets = Array.isArray(sheets) ? sheets : []
  }

  /**
   * 【workbench-per-sheet-preview 新增】切换当前预览的 sheet。
   * @param {string} name sheet 名;传 '' 表示回到"全部 sheet 合并预览"。
   */
  function setActiveSheet(name) {
    draft.value.activeSheet = name || ''
  }
  function setTemplate(t) {
    draft.value.templateId = t.id
    draft.value.templateName = t.name
  }

  /**
   * 【chart-edit-prebuilt-dashboard 新增】写入"前置生成"出来的看板。
   * @param {Object} payload {groups?: [{sheet, charts}], charts?: flat}
   *        - 优先用 groups(按 sheet 分组),fallback 到 charts(flat)
   * @param {boolean} keepSelection true 时保留旧勾选;false 时清空
   */
  function setDashboard(payload, keepSelection = false) {
    let groups = []
    if (payload && Array.isArray(payload.groups)) {
      groups = payload.groups.map((g) => ({
        sheet: g.sheet || "",
        charts: Array.isArray(g.charts) ? g.charts.map((c, idx) => ({
          ...c,
          _key: `${g.sheet || "_default"}__${idx}`,
        })) : [],
      }))
    } else if (payload && Array.isArray(payload.charts)) {
      // 兼容旧 flat 数组
      groups = [{
        sheet: "",
        charts: payload.charts.map((c, idx) => ({ ...c, _key: `__${idx}` })),
      }]
    }
    draft.value.dashboardGroups = groups
    draft.value.dashboardCharts = groups.flatMap((g) => g.charts)
    if (!keepSelection) {
      draft.value.selectedChartKeys = new Set()
    } else {
      const valid = new Set()
      groups.forEach((g) => g.charts.forEach((c) => {
        if (draft.value.selectedChartKeys.has(c._key)) valid.add(c._key)
      }))
      draft.value.selectedChartKeys = valid
    }
  }

  /**
   * 切换某张图表的"加入报表"勾选状态。
   * @param {string} key 全局唯一 key(<sheet>__<idx>)
   */
  function toggleChartSelected(key) {
    const next = new Set(draft.value.selectedChartKeys)
    if (next.has(key)) next.delete(key)
    else next.add(key)
    draft.value.selectedChartKeys = next
  }

  /**
   * 返回真正会被带去 AI 分析报告的预选图表(去重后的 chart 列表)。
   */
  function getSelectedCharts() {
    const out = []
    draft.value.dashboardGroups.forEach((g) => {
      g.charts.forEach((c) => {
        if (draft.value.selectedChartKeys.has(c._key)) {
          out.push({
            title: c.title,
            option: c.option,
            style_key: c.style_key || draft.value.chartStyle,
            sheet: c.sheet || g.sheet || "",
            type: (c.option && c.option.series && c.option.series[0] && c.option.series[0].type) || "bar",
          })
        }
      })
    })
    return out
  }

  function reset() {
    draft.value.datasetId = null
    draft.value.datasetName = ''
    // 【dataset-multisheet-merge 新增】重置时一并清掉工作表信息
    draft.value.datasetSheets = []
    // 【workbench-per-sheet-preview 新增】重置时清掉当前 sheet 选择
    draft.value.activeSheet = ''
    draft.value.templateId = null
    draft.value.templateName = ''
    // 【chart-edit-prebuilt-dashboard 新增】重置草稿看板
    draft.value.dashboardCharts = []
    draft.value.dashboardGroups = []
    draft.value.selectedChartKeys = new Set()
  }

  return {
    draft,
    setDataset,
    clearDataset,
    setDatasetSheets,
    setActiveSheet,
    setTemplate,
    setDashboard,
    toggleChartSelected,
    getSelectedCharts,
    reset,
  }
})
