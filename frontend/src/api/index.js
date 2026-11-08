/**
 * 后端 API 一站式调用
 */
import http from './http'

/**
 * 通用 Blob 响应下载工具（导出类接口共用）。
 *
 * 适用场景：后端直接返回文件流（xlsx / pdf / docx ...）的接口。
 * axios 响应拦截器在 responseType === 'blob' 时会**返回完整 resp**（含 headers），
 * 所以这里既能读到 resp.data 里的二进制，也能读到 resp.headers 里的 Content-Disposition。
 *
 * 处理流程：
 *   1. 取出 blob：axios 响应取 resp.data；调用方直接传 Blob 时取自身
 *   2. 解析文件名（优先级从高到低）：
 *      - filename*=UTF-8''xxx  → RFC 5987 编码，后端返回中文文件名时使用，需 decodeURIComponent
 *      - filename="xxx"         → 普通形式
 *      - 都没有 → 使用调用方传入的 fallbackName
 *   3. 通过 a[download] 触发浏览器保存，并延迟释放 ObjectURL 防止内存泄漏
 *
 * 失败时抛出中文错误信息，交给调用方用 ElMessage 统一提示。
 *
 * @param {Object|Blob} resp axios 响应对象，或已经是 Blob 的值
 * @param {string} fallbackName 文件名兜底值，默认 'download.xlsx'
 * @returns {Promise<string>} 实际使用的文件名
 */
export async function downloadBlobResponse(resp, fallbackName = 'download.xlsx') {
  // 1) 取 blob：axios 响应的二进制在 resp.data；直接传 Blob 时用自身
  const blob =
    resp?.data instanceof Blob ? resp.data : resp instanceof Blob ? resp : null
  if (!blob) throw new Error('导出失败：未获取到文件内容')

  // 2) 解析文件名：优先 filename*（含中文名时后端用这种编码形式）
  const cd = resp?.headers?.['content-disposition'] || ''
  let filename = ''
  const star = cd.match(/filename\*=UTF-8''([^;]+)/i)
  if (star) {
    try {
      filename = decodeURIComponent(star[1])
    } catch {
      // 解码失败时保留原始值，不因为文件名解析失败而中断下载
      filename = star[1]
    }
  }
  if (!filename) {
    const plain = cd.match(/filename="?([^";]+)"?/)
    // plain[1] 以 "*=" 开头说明命中的是 filename*= 空值等异常形态,视为未解析到文件名
    if (plain && !plain[1].startsWith('*=')) {
      try {
        filename = decodeURIComponent(plain[1])
      } catch {
        filename = plain[1]
      }
    }
  }
  if (!filename) filename = fallbackName

  // 3) a[download] 触发浏览器保存
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  // 释放 URL 资源，避免内存泄漏（延迟 1s，确保浏览器已接管这次下载）
  setTimeout(() => URL.revokeObjectURL(url), 1000)

  return filename
}

export const datasetsApi = {
  list: () => http.get('/datasets'),
  upload: (formData) =>
    http.post('/datasets/upload', formData, { headers: { 'Content-Type': 'multipart/form-data' } }),
  preview: (id, opts = {}) => {
    // 多 sheet 分级预览(2026-10 新增):
    //   - opts.sheet 传具体 sheet 名 → 后端按单 sheet 预览,缺失值/列类型按该 sheet 真实数据计算
    //   - 缺省 / opts.sheet='__all__' → 走旧的"全部 sheet 合并大表"预览(AI 分析链路默认行为)
    const params = {}
    if (opts.sheet) params.sheet = opts.sheet
    return http.get(`/datasets/${id}/preview`, { params })
  },
  rename: (id, name) => http.post(`/datasets/${id}/rename`, { name }),
  remove: (id) => http.delete(`/datasets/${id}`),
  /**
   * 下载数据集文件。
   * 调用后端 GET /datasets/{id}/download,后端返回文件流(FileResponse),
   * 这里通过 a 标签 + Blob 触发浏览器保存对话框。
   * 文件名优先从响应头 Content-Disposition 解析(后端给的 file_name 带正确后缀),
   * 兜底用数据集名,但只有"看起来是真后缀"(字母/数字结尾)时才直接用,否则补 .xlsx。
   * 失败时抛错,交给调用方统一 ElMessage 提示。
   */
  download: async (id, name) => {
    // blob 响应时拦截器返回完整 resp(含 headers)
    const resp = await http.get(`/datasets/${id}/download`, { responseType: 'blob' })
    const blob = resp?.data instanceof Blob ? resp.data : resp instanceof Blob ? resp : null
    if (!blob) throw new Error('下载内容为空')
    // 1) 优先解析 Content-Disposition(后端 FileResponse 对中文文件名会用 filename*=utf-8''编码格式)
    const cd = resp?.headers?.['content-disposition'] || ''
    let filename = ''
    const star = cd.match(/filename\*=(?:UTF-8|utf-8)''([^;]+)/)
    if (star) filename = decodeURIComponent(star[1])
    else {
      const plain = cd.match(/filename="?([^";]+)"?/)
      if (plain) filename = decodeURIComponent(plain[1])
    }
    // 2) 兜底:数据集名可能含时间戳"16.36.54"这类假后缀,必须校验结尾是 1~5 位字母数字才算真后缀
    if (!filename) {
      const okSuffix = name && /\.[A-Za-z0-9]{1,5}$/.test(name)
      filename = okSuffix ? name : `${name || 'dataset'}.xlsx`
    }
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    // 释放 URL 资源,避免内存泄漏
    setTimeout(() => URL.revokeObjectURL(url), 1000)
  },
}

export const templatesApi = {
  list: (category) => http.get('/templates', { params: { category } }),
  detail: (id) => http.get(`/templates/${id}`),
  // 新建模板(payload: { name, category, description, structure })
  // 注意:Prompt 结构不再由前端控制,统一由后端 prompt_builder 按 category 生成
  create: (payload) => http.post('/templates', payload),
  // 更新模板
  update: (id, payload) => http.put(`/templates/${id}`, payload),
  // 删除模板(内置模板后端会返回 400)
  remove: (id) => http.delete(`/templates/${id}`),
  // 复制模板:基于已有模板生成"<原名> - 副本"的自建模板
  duplicate: (id) => http.post(`/templates/${id}/duplicate`),
  // 已有分类列表(去重),供新建/编辑表单下拉
  categories: () => http.get('/templates/categories'),
}

export const tasksApi = {
  list: () => http.get('/tasks'),
  detail: (id) => http.get(`/tasks/${id}`),
  create: (payload) => http.post('/tasks', payload),
  reanalyze: (id) => http.post(`/tasks/${id}/reanalyze`),
  // 重新生成 AI 摘要:POST /tasks/{id}/summarize,响应是最新的 task 详情
  // 用于报告页「生成摘要」按钮 —— 之前该按钮只重新拉取,没有真正调模型,本接口修复这个 bug
  summarize: (id) => http.post(`/tasks/${id}/summarize`),
  // 基于用户的修改意见,新建一个 v2/v3 子任务并保持原任务不变。
  // 请求体:{feedback: str<=500}
  // 响应:新 task 的详情(同 tasksApi.detail 形状)。
  // 失败:
  //   404 任务不存在 / 409 任务未完成 / 400 feedback 缺失或超长
  regenerate: (id, feedback) => http.post(`/tasks/${id}/regenerate`, { feedback }),
  remove: (id) => http.delete(`/tasks/${id}`),
  // 批量删除:与单条删除保持一致用 DELETE,请求体里带 ids。
  // axios 的 delete 必须把 body 放到第二个参数的 data 字段里,直接传对象会被忽略。
  // 后端路由:@router.delete("/batch") —— 见 backend/app/api/tasks.py
  batchRemove: (ids) => http.delete('/tasks/batch', { data: { ids } }),
}

export const workbenchApi = {
  // 工作台同步分析入口。
  // 请求体:{ dataset_id, template_id, cleaning_rules?, name?, analysis_intent?: string<=500 }
  // analysis_intent 非空时,后端会把它注入到 Prompt 第 6 段,
  // 并在报告页底部「异常批注」追加一条 info 级"分析重点"标注;
  // 空字符串 / 不传 → 行为与旧实现完全一致,Prompt 无第 6 段。
  analyze: (payload) => http.post('/workbench/analyze', payload),
  // ===== 【chart-edit-prebuilt-dashboard 新增】4 个方法 =====
  // 3 套图表风格枚举(供前端 el-option 下拉)
  listChartStyles: () => http.get('/workbench/charts/styles'),
  // 前置生成可视化看板:dataset_id + style_key → 2~6 张 option(本地规则)
  // 【chart-render-smart-2026-10 新增】可选 ai_curate=true 走「混合出图」:
  //   本地骨架 + LLM 精选 TopN + 标题润色,LLM 不可用时自动降级为本地精选。
  //   target_count 控制 ai_curate=true 时每个 sheet 保留的图表数,默认 3。
  previewCharts: (dataset_id, style_key, opts = {}) => {
    const payload = { dataset_id, style_key }
    if (opts.aiCurate !== undefined) payload.ai_curate = !!opts.aiCurate
    if (opts.targetCount !== undefined) payload.target_count = Number(opts.targetCount) || 3
    return http.post('/workbench/charts/preview', payload)
  },
  // 自然语言 / 快捷指令改图:instruction + target_chart → 新 option
  // 命中关键词规则即返回;否则 LLM 兜底
  modifyChart: (payload) => http.post('/workbench/charts/modify', payload),
  // ===== 【visual-report-export 新增】可视化报表(Excel)导出 =====
  /**
   * 导出可视化报表(图表编辑弹窗入口)。
   * POST /workbench/charts/report
   * 请求体:{
   *   dataset_id: number,            // 必填
   *   style_key?: string,            // 可选,非法值后端回落 tech_advanced
   *   include_ai?: boolean,          // 可选,默认 true
   *   charts: [{ title, option, sheet?, png_base64 }]  // 必填非空
   * }
   * 说明:png_base64 为**不带 dataURL 前缀**的纯 base64 字符串。
   * 响应:xlsx 二进制流(Content-Disposition 带文件名),
   * 前端用 downloadBlobResponse(resp) 触发浏览器下载。
   */
  exportVisualReport: (payload) =>
    http.post('/workbench/charts/report', payload, { responseType: 'blob' }),
}

export const toolsApi = {
  merge: (payload) => http.post('/tools/merge', payload),
  validate: (payload) => http.post('/tools/validate', payload),
  // 自然语言数据问答:整体透传 payload(POST /tools/nl_query)
  // 请求体:{ dataset_id, question, persist?, name?, history?: [{role:'user'|'ai', content}] }
  // 响应除旧字段(explanation/rows/columns/sheets/sheet_results/plan/task_id/task_url)外,
  // 新增 structured 结构化结果:{ answer, analysis[], data_table{columns,rows}, suggestions[], followups[] },
  // LLM 失败时 structured 为 null(前端需降级渲染 explanation + rows/columns)
  nlQuery: (payload) => http.post('/tools/nl_query', payload),
  compare: (payload) => http.post('/tools/compare', payload),
  // 数据清洗:把当前数据集按 cleaning_rules 清洗后落盘为新数据集,不进 task 表也不调 LLM
  // 请求体:{ dataset_id, cleaning_rules, name?, keep_formulas?,
  //          derived_columns?: [{sheet_name,name,expression}],
  //          pivot?: {sheet_name,row_dim,col_dim,value,agg,include_margin} }
  // 响应:{ id, name, file_name, row_count, cleaned_stats }
  clean: (payload) => http.post('/tools/clean', payload),
  // ===== 【clean-ai-derived-columns-pivot 新增】2 个前端接口 =====
  // 数据集字段结构(供派生列 / 透视字段下拉使用)
  //   请求:query 参数 dataset_id
  //   响应:{ dataset_id, sheets: [{ name, row_count,
  //         columns: [{ name, type, non_null, unique_count }] }] }
  //   type 取值仅 number/text/date;多 sheet 时逐 sheet 返回各自列结构
  schema: (datasetId) => http.get('/tools/schema', { params: { dataset_id: datasetId } }),
  // AI 推荐派生列(该接口内部会调一次 LLM)
  //   请求体:{ dataset_id }
  //   响应:{ suggestions: [{ sheet_name, name, expression, description }], fallback: boolean }
  //   fallback=true 表示 LLM 不可用、后端走了规则兜底,此时列表可能为空
  suggestColumns: (payload) => http.post('/tools/suggest_columns', payload),
  // 图表编辑:自然语言指令 → 后端 LLM 生成新 ECharts option
  // 请求体:{ dataset_id, instruction, current_option?, current_type?, title? }
  // 响应:{ type, option, explanation, error, fallback }
  chartEdit: (payload) => http.post('/tools/chart_edit', payload),
  // 图表落库:用户在前端确认后,把候选图显式写回历史任务
  // 请求体:{ task_id, chart_index?, type, title, option }
  // 响应:{ task_id, chart_option }
  chartApply: (payload) => http.post('/tools/chart_apply', payload),

  // ===== 【workbench-ai-pdf-tools 新增】4 个 PDF 处理方法 =====
  // 4 个端点统一传 FormData(file + 可选 ai_enabled / save_as_dataset),
  // 后端响应里含 download_url,前端再调 downloadFile 触发浏览器保存。
  pdfToWord: (formData) =>
    http.post('/tools/pdf/to_word', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
    }),
  pdfToExcel: (formData) =>
    http.post('/tools/pdf/to_excel', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
    }),
  pdfExtractText: (formData) =>
    http.post('/tools/pdf/extract_text', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
    }),
  pdfCompress: (formData) =>
    http.post('/tools/pdf/compress', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
    }),
  // ===== 【pdf-photo-to-pdf-ocr 新增】照片转 PDF(OCR) =====
  // 与上面 4 个方法的唯一差别:FormData 里是**多个** files(同名 key 重复 append),
  // 另带 mode / ocr_enabled / save_to_history 三个 Form 字段。
  pdfImageToPdf: (formData) =>
    http.post('/tools/pdf/image_to_pdf', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
    }),
  // 拼出 PDF 工具产出文件的下载路径(后端 GET /api/tools/pdf/download/{file_name})
  // 必须是纯相对路径(不带 /api 前缀),由 axios baseURL='/api' 自动拼接,
  // 否则会变成 /api/api/... 双前缀导致 404(同 reportsApi.exportUrl 的注意事项)
  pdfDownloadUrl: (fileName) => `/tools/pdf/download/${fileName}`,
  // 通用文件下载:用 a 标签 + Blob 触发浏览器保存。
  // 优先解析 Content-Disposition 的 filename*;失败时回退到外部传入的 filename。
  downloadFile: async (url, filename) => {
    const resp = await http.get(url, { responseType: 'blob' })
    const blob = resp?.data instanceof Blob ? resp.data : resp instanceof Blob ? resp : null
    if (!blob) throw new Error('下载内容为空')
    const cd = resp?.headers?.['content-disposition'] || ''
    let finalName = filename || ''
    const star = cd.match(/filename\*=(?:UTF-8|utf-8)''([^;]+)/)
    if (star) finalName = decodeURIComponent(star[1])
    if (!finalName) finalName = 'download'
    const objUrl = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = objUrl
    a.download = finalName
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    setTimeout(() => URL.revokeObjectURL(objUrl), 1000)
  },
}

/**
 * 报告导出 / 预览 API。
 *
 * exportUrl(id, format, theme, source):
 *   - format === 'pptx' 时:会拼上 ?format=pptx&theme=<theme>(默认 deep_cyan)
 *   - 其它格式:仅 ?format=<format>,theme 参数忽略
 *   - source === 'ai' 时额外拼 &source=ai(AI 演示稿版 PPTX);
 *     不传或传 'report' 时保持老行为(报告全文版)
 *
 * previewCharts(id, theme):
 *   - 返回 {theme, slides: [{kind, title, png_b64?, body?}]}
 *   - 用于前端 PPT 弹窗的"逐页缩略图预览"(报告全文版)
 */
export const reportsApi = {
  // 注意:http 实例的 baseURL 已经是 /api,所以下面不要再拼 /api 前缀,
  // 否则会出现 /api/api/... 双前缀导致 404。
  exportUrl: (id, format, theme = "deep_cyan", source = "") => {
    let url = `/api/reports/${id}/export?format=${format}`
    if (format === "pptx" && theme && theme !== "deep_cyan") {
      url = `/api/reports/${id}/export?format=pptx&theme=${theme}`
    }
    // source=ai 只对 pptx 生效:让后端走「AI 演示稿」分支(命中缓存直接导出)
    if (format === "pptx" && source === "ai") {
      url += '&source=ai'
    }
    return url
  },
  previewCharts: (id, theme = "deep_cyan") =>
    http.get(`/reports/${id}/preview_charts`, { params: { theme } }),

  // ===== 【AI 演示稿】新增 3 个方法 =====
  /**
   * 生成(或命中缓存获取)AI 演示稿。
   * POST /reports/{id}/deck?force=true —— force=true 时忽略缓存强制重新生成。
   * 返回:{ status: 'ready'|'failed'|'generating', deck: {...}|null, error: string|null }
   * 该接口内部会调大模型,耗时 60~180 秒,前端必须给足超时(http 已是 600s)。
   */
  generateDeck: (id, force = false) =>
    // body 传 null:后端该接口没有 body 参数,force 走 query
    http.post(`/reports/${id}/deck`, null, { params: force ? { force: true } : {} }),

  /**
   * 取演示稿的逐页缩略图数据。
   * GET /reports/{id}/deck?theme=xxx
   * 返回:{ status: 'ready'|'none', deck: {...}|null, slides: [...] }
   * 没有演示稿时后端返回 {status:'none', slides:[]} 且不报错,前端据此显示"生成"入口。
   */
  deckPreview: (id, theme = "deep_cyan") =>
    http.get(`/reports/${id}/deck`, { params: { theme } }),

  /**
   * 演示稿讲义(Markdown)下载地址。
   * 这里**必须带 /api 前缀**:该地址是给 window.open / a[download] 用的,
   * 不经过 axios(axios baseURL 已是 /api),不带前缀会 404。
   */
  deckScriptUrl: (id) => `/api/reports/${id}/deck/script.md`,

  // ===== 【visual-report-export 新增】可视化报表(Excel)导出 =====
  /**
   * 导出可视化报表(报告页入口)。
   * POST /reports/{task_id}/visual-report
   * 请求体:{
   *   include_ai?: boolean,   // 可选,默认 true
   *   charts?: [{ key, png_base64 }]  // 可选
   * }
   * 说明:key 形如 "<章节名>|<序号>"(序号从 1 开始),用于把图片归位到对应章节。
   * 响应:xlsx 二进制流(Content-Disposition 带文件名),
   * 前端用 downloadBlobResponse(resp) 触发浏览器下载。
   */
  exportVisualReport: (taskId, payload = {}) =>
    http.post(`/reports/${taskId}/visual-report`, payload, { responseType: 'blob' }),
}
