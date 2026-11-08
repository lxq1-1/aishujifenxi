<!-- 报告预览：正文 + 图表 + 汇总 + 摘要 + 异常批注 + 导出 -->
<template>
  <div class="page">
    <div class="page-header">
      <h2 class="h1-title" style="font-size: 22px">{{ task?.name || '报告预览' }}</h2>
      <div>
        <el-button @click="onSummarize" :loading="summarizing">
          <el-icon><MagicStick /></el-icon> 生成摘要
        </el-button>
        <el-button @click="exportReport('pdf')">导出 PDF</el-button>
        <el-button @click="exportReport('docx')">导出 Word</el-button>
        <el-button @click="openPptxPreview">导出 PPT</el-button>
        <!-- 【export-visual-report 新增】导出 Excel 可视化报表
             与 PDF/Word/PPT 的区别：那三种是长文本报告，没有卡片化的指标概览；
             报表是结构化 xlsx（KPI 卡 + AI 结论 + 数据明细 + 图表洞察），可筛选可再加工。 -->
        <el-button type="primary" plain :loading="exportingVisualReport" @click="exportVisualReport">
          📊 导出 Excel 报表
        </el-button>

      </div>
    </div>

    <div v-if="loading" class="dim" style="padding: 40px; text-align: center">加载中...</div>
    <div v-else-if="task">
      <!-- 【report-failed-state 新增】任务失败时,显式展示原因 + 重试入口,
           不再像旧实现那样由本地 pandas 兜底拼接出"29 行 × 24 列"式的假报告。
           原因:`analyzer.run_analysis` 在 LLM 失败时已写 task.status='failed'
           + task.error_message,前端拿到这 2 个字段按 dashboard 风格展示。
           「重新分析」按钮复用 tasksApi.reanalyze(id),同 task 重跑一次。 -->
      <el-alert
        v-if="task.status === 'failed'"
        :title="failedTitle"
        type="error"
        :closable="false"
        show-icon
        style="margin-bottom: 14px"
      >
        <template #default>
          <div class="failed-alert-body">
            <div class="failed-error-text">{{ task.error_message || 'AI 生成失败,未提供具体原因。' }}</div>
            <div class="failed-actions">
              <el-button
                type="primary"
                size="small"
                :loading="retrying"
                @click="onRetryFailed"
              >
                重新分析
              </el-button>
              <el-button size="small" @click="$router.push('/workbench')">
                返回工作台
              </el-button>
            </div>
          </div>
        </template>
      </el-alert>
      <!-- 【report-feedback-and-versioning 新增】版本标识 + 上一版本导航 + 迭代意见提示行 -->
      <div class="version-bar" style="display: flex; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 14px">
        <el-tag v-if="displayVersionTag.kind === 'baseline'" type="info" effect="plain">
          {{ displayVersionTag.text }}
        </el-tag>
        <el-tag v-else type="success" effect="plain">
          {{ displayVersionTag.text }} (基于 #{{ task.parent_task_id }} 重写)
        </el-tag>
        <el-button
          v-if="task.parent_task_id"
          link
          type="primary"
          @click="goToParent"
        >
          ← 上一版本 (#{{ task.parent_task_id }})
        </el-button>
        <!-- 「下一版本」按钮 TODO:本次 spec 不要求实现,留作后续扩展 -->
      </div>
      <div v-if="task.feedback_text" class="feedback-hint dim small" style="margin-bottom: 12px">
        📌 本版本基于任务 #{{ task.parent_task_id }} 重新生成,意见:{{
          task.feedback_text.length > 100 ? task.feedback_text.slice(0, 100) + '…' : task.feedback_text
        }}
      </div>

      <article class="report-body">
        <h1>{{ cleanText(report.title) || task.name }}</h1>
        <div v-if="summary" class="summary">
          <strong>摘要：</strong>{{ summary }}
        </div>
        <!--
          【pdf-tools-history-persistence 新增】PDF 工具产物下载卡片:
          仅当 report.pdf_artifact 存在时显示,提供"下载 Word/Excel/压缩后PDF/提取TXT"按钮
          直接调 toolsApi.downloadFile 触发浏览器保存,不依赖 PPT 导出链路。
        -->
        <div v-if="pdfArtifact" class="pdf-artifact-card">
          <div class="pdf-artifact-info">
            <div class="pdf-artifact-title">
              <el-icon><Download /></el-icon>
              <span>处理产物</span>
            </div>
            <div class="pdf-artifact-meta dim small">
              <span v-if="pdfArtifact.original_size">
                原始 {{ formatArtifactSize(pdfArtifact.original_size) }}
              </span>
              <span v-if="pdfArtifact.output_size && pdfArtifact.output_size !== pdfArtifact.original_size">
                → 输出 {{ formatArtifactSize(pdfArtifact.output_size) }}
              </span>
              <span v-if="pdfArtifact.compression_ratio !== undefined">
                · 压缩比 {{ pdfArtifact.compression_ratio }}%
              </span>
              <span v-if="pdfArtifact.table_count !== undefined">
                · 识别 {{ pdfArtifact.table_count }} 个表格
              </span>
              <span v-if="pdfArtifact.char_count !== undefined">
                · {{ pdfArtifact.char_count }} 字
              </span>
              <span v-if="pdfArtifact.ai_fallback">
                · <span style="color:#e6a23c">AI 降级为原文</span>
              </span>
            </div>
          </div>
          <el-button
            type="primary"
            :loading="downloadingArtifact"
            @click="downloadPdfArtifact"
          >
            {{ pdfArtifact.label || '下载产物' }}
          </el-button>
        </div>
        <!--
          章节正文 + 该章节对应的图表:按 chart_option.by_section[key] 渲染。
          - 新报告(LLM 已按章节配图):每个有数据的章节正文下方紧贴其图表;
          - 旧报告(by_section 仅有 "数据图表"):章节正文下方不渲染图,统一走末尾兜底区。
        -->
        <div v-for="(content, key) in report.sections" :key="key" class="section-block">
          <h2>{{ cleanKey(key) }}</h2>
          <!--
            报告 sections 值的类型分支渲染(2026-10 重构):
              - str   : 通用报告,作为段落(white-space: pre-wrap 保证换行可见)
              - list  : 合同报告的「条款差异表/利弊分析/谈判建议」,按子类型分支:
                  * 每项是 dict 且 key === '条款差异表' → 条款卡片
                  * 每项是 dict 且 key === '利弊分析'   → 利弊卡片
                  * 每项是 str 且 key === '谈判建议'    → 建议列表
                  * 兜底 list → 通用列表
              - dict  : 数据集对比·变更追踪的「对比详情」,按 added/removed/modified 分块渲染表格
              - 其他  : 走 JSON.stringify,避免页面崩溃
            改这段同时给通用段落加 white-space: pre-wrap,
            避免历史报告中带 \n 的内容被 HTML 合并成单行。
          -->
          <!--
            章节正文按「段落」渲染,不再是一整块文字。
            splitParagraphs 会:尊重模型已有的空行分段 → 对超长段按句号切分
            → 再按「其一/其二/首先/其次」等语义标记隐式拆段。
            这样即使是模型写成的 500 字长块,页面也能呈现为 3~5 个可读段落。
          -->
          <div v-if="typeof content === 'string'" class="report-section-text">
            <p
              v-for="(para, pi) in splitParagraphs(content)"
              :key="'p-' + pi"
              class="report-paragraph"
            >{{ para }}</p>
          </div>
          <template v-else-if="Array.isArray(content)">
            <!-- 条款差异表:list[dict] -->
            <div
              v-if="key === '条款差异表'"
              class="contract-clause-list"
            >
              <div
                v-for="(item, i) in content"
                :key="'c' + i"
                class="contract-clause-card"
                :class="{
                  'clause-high': item.importance === 'high' && ['modified','added','removed'].includes(item.change_type),
                }"
              >
                <div class="clause-head">
                  <span class="clause-label">{{ item.clause_label || '—' }}</span>
                  <el-tag
                    :type="contractChangeTypeTag(item.change_type).type"
                    size="small"
                    effect="plain"
                  >
                    {{ contractChangeTypeTag(item.change_type).label }}
                  </el-tag>
                  <el-tag
                    v-if="item.importance === 'high'"
                    type="danger"
                    size="small"
                    effect="dark"
                  >高</el-tag>
                  <el-tag
                    v-else-if="item.importance === 'low'"
                    type="info"
                    size="small"
                    effect="plain"
                  >低</el-tag>
                  <el-tag v-else size="small" effect="plain">中</el-tag>
                </div>
                <div class="clause-row">
                  <div class="clause-side">
                    <div class="clause-side-label">左方</div>
                    <div class="clause-side-value">{{ item.left_value || '—' }}</div>
                  </div>
                  <div class="clause-arrow">→</div>
                  <div class="clause-side">
                    <div class="clause-side-label">右方</div>
                    <div class="clause-side-value">{{ item.right_value || '—' }}</div>
                  </div>
                </div>
                <div v-if="item.diff_note" class="clause-note">
                  <strong>差异:</strong>{{ item.diff_note }}
                </div>
                <div v-else class="clause-note dim">无明显差异</div>
              </div>
            </div>
            <!-- 利弊分析:list[dict] -->
            <div
              v-else-if="key === '利弊分析'"
              class="contract-pc-list"
            >
              <div
                v-for="(item, i) in content"
                :key="'pc' + i"
                class="contract-pc-card"
                :class="{
                  'pc-left':  item.side === 'left',
                  'pc-right': item.side === 'right',
                  'pc-both':  item.side === 'both' || !item.side,
                }"
              >
                <div class="pc-head">
                  <el-tag
                    :type="item.side === 'left' ? 'success' : (item.side === 'right' ? 'warning' : 'info')"
                    size="small"
                    effect="dark"
                  >
                    {{ item.side === 'left' ? '⬅️ 左方更优' : (item.side === 'right' ? '➡️ 右方更优' : '⚖️ 各有优劣') }}
                  </el-tag>
                  <span class="pc-title">{{ item.title }}</span>
                </div>
                <div v-if="item.detail" class="pc-detail">{{ item.detail }}</div>
              </div>
            </div>
            <!-- 谈判建议:list[str] -->
            <ul
              v-else-if="key === '谈判建议'"
              class="contract-tips"
            >
              <li v-for="(t, i) in content" :key="'tip' + i">
                <el-icon style="color:#22d3ee;margin-right:6px;vertical-align:middle"><Star /></el-icon>
                <span>{{ t }}</span>
              </li>
            </ul>
            <!-- 兜底 list:通用列表 -->
            <ul v-else class="report-generic-list">
              <li v-for="(item, i) in content" :key="'g' + i">
                <template v-if="typeof item === 'string'">{{ item }}</template>
                <template v-else>{{ formatListItem(item) }}</template>
              </li>
            </ul>
          </template>
          <!-- 【row-diff-detail-expansion 新增】对比详情·结构化明细
               值是 dict 且带 added/removed/modified 字段时,按"新增/删除/修改"三块渲染明细;
               旧报告(值是字符串)走上面 p 分支,行为不变。 -->
          <div
            v-else-if="key === '对比详情' && content && typeof content === 'object'
                       && Array.isArray(content.added) && Array.isArray(content.removed)"
            class="diff-detail"
          >
            <p class="report-section-text diff-summary">{{ cleanText(content.summary || '') }}</p>

            <!-- 新增行明细 -->
            <div class="diff-block" v-if="content.added.length">
              <div class="diff-block-head">
                <span class="diff-tag diff-tag-add">新增</span>
                <span class="diff-count">共 {{ content.added_total ?? content.added.length }} 条</span>
              </div>
              <table class="diff-table">
                <thead>
                  <tr>
                    <th
                      v-for="col in getDiffColumns(content.added)"
                      :key="'hav-add-' + col"
                    >{{ col }}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, i) in content.added.slice(0, 50)" :key="'hav-add-' + i">
                    <td
                      v-for="col in getDiffColumns(content.added)"
                      :key="'hav-add-' + col"
                    >{{ row[col] ?? '—' }}</td>
                  </tr>
                </tbody>
              </table>
              <div
                v-if="content.added.length > 50"
                class="dim diff-truncated"
              >仅展示前 50 条,还有 {{ content.added.length - 50 }} 条未展开</div>
            </div>

            <!-- 删除行明细 -->
            <div class="diff-block" v-if="content.removed.length">
              <div class="diff-block-head">
                <span class="diff-tag diff-tag-remove">删除</span>
                <span class="diff-count">共 {{ content.removed_total ?? content.removed.length }} 条</span>
              </div>
              <table class="diff-table">
                <thead>
                  <tr>
                    <th
                      v-for="col in getDiffColumns(content.removed)"
                      :key="'hav-rm-' + col"
                    >{{ col }}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, i) in content.removed.slice(0, 50)" :key="'hav-rm-' + i">
                    <td
                      v-for="col in getDiffColumns(content.removed)"
                      :key="'hav-rm-' + col"
                    >{{ row[col] ?? '—' }}</td>
                  </tr>
                </tbody>
              </table>
              <div
                v-if="content.removed.length > 50"
                class="dim diff-truncated"
              >仅展示前 50 条,还有 {{ content.removed.length - 50 }} 条未展开</div>
            </div>

            <!-- 修改明细:字段级 before / after -->
            <div class="diff-block" v-if="content.modified && content.modified.length">
              <div class="diff-block-head">
                <span class="diff-tag diff-tag-modify">修改</span>
                <span class="diff-count">
                  共 {{ content.modified_total ?? content.modified.length }} 条
                  <template v-if="content.truncated_modified">
                    · 仅展示前 {{ content.shown_modified ?? content.modified.length }} 条
                  </template>
                </span>
              </div>
              <div
                v-for="(m, i) in content.modified"
                :key="'hav-mod-' + i"
                class="diff-modify-card"
              >
                <div class="diff-modify-key">
                  <span class="dim">主键:</span>{{ m.key ?? '—' }}
                </div>
                <table class="diff-table diff-modify-fields" v-if="m.field_diffs && m.field_diffs.length">
                  <thead>
                    <tr><th>字段</th><th>变更前</th><th>变更后</th></tr>
                  </thead>
                  <tbody>
                    <tr v-for="(fd, j) in m.field_diffs" :key="'fd-' + i + '-' + j">
                      <td>{{ fd.field }}</td>
                      <td class="diff-before">{{ fd.before ?? '—' }}</td>
                      <td class="diff-after">{{ fd.after ?? '—' }}</td>
                    </tr>
                  </tbody>
                </table>
                <div v-else class="dim">无字段级差异</div>
              </div>
            </div>
          </div>
          <p v-else class="report-section-text">{{ formatObjectText(content) }}</p>
          <div v-if="sectionCharts[key] && sectionCharts[key].length" class="charts section-charts">
            <div v-for="(c, i) in sectionCharts[key]" :key="i" class="chart-box">
              <v-chart class="chart" :option="c.option" autoresize />
            </div>
          </div>
          <!--
            【2026-10 新增】章节明细数据表。
            正文里的数字需要能被核对,图表看趋势但看不出构成明细,
            所以每章节正文下方挂一张与该章节主题匹配的结构化表格
            (KPI 明细 / 周期走势 / 结构占比 / 未达标清单)。
            数据来自后端 chart_option.detail_tables,由 pandas_engine
            的 build_section_detail_tables 产出 —— 表格里的每一个数字
            都来自真实数据,不经过 LLM。
          -->
          <div v-if="getSectionDetailTable(key)" class="detail-table-block">
            <div class="detail-table-title">{{ getSectionDetailTable(key).title }}</div>
            <table class="detail-table">
              <thead>
                <tr>
                  <th v-for="col in getSectionDetailTable(key).columns" :key="col">{{ col }}</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(row, ri) in getSectionDetailTable(key).rows" :key="'dt-' + ri">
                  <td
                    v-for="(cell, ci) in row"
                    :key="'dc-' + ri + '-' + ci"
                    :class="{ 'dt-num': typeof cell === 'number' }"
                  >{{ formatCell(cell) }}</td>
                </tr>
              </tbody>
            </table>
            <div v-if="detailTableNote(key)" class="detail-table-note dim small">
              {{ detailTableNote(key) }}
            </div>
          </div>
          <!--
            【按章节挂载 · 补章]】
              若该章节没有对应的图表(LLM 漏配 + 后端 Pandas 兜底也未生成),
              在正文下方给一行温和提示,避免"读者以为章节坏掉了"。
              隐藏条件:
                - 章节正文里含「本期数据未覆盖」已与图表缺失同源,提示冗余
                  → 用精确包含「未覆盖」的字面量判断,不引入新的 i18n key;
                - 章节本身是「附录/说明」这类本质就无图章节(可通过
                  sectionCharts 长度 == 0 反推,但若全部都缺图提示会刷屏,
                  所以仅在「报告至少有孤儿图」时提示——见 totalCharts 计算)。
          -->
          <div
            v-if="(!sectionCharts[key] || !sectionCharts[key].length) && hasAnyCharts
                 && !(typeof content === 'string' && content.includes('本期数据未覆盖'))"
            class="section-no-chart dim small"
          >
            <el-icon style="vertical-align: middle; margin-right: 4px"><InfoFilled /></el-icon>
            本章节本期未生成图表,如需可视化请补充该章节相关数据后重新生成。
          </div>
        </div>

        <!-- 兜底:未挂载到任何章节的图表(典型为旧报告 {by_section={"数据图表":[...]}})统一汇总在末尾 -->
        <!-- 【dataset-contract-compare-report 新增】合同对比场景无结构化指标,这里整段隐藏 -->
        <template v-if="!isContractReport">
          <h2 v-if="orphanCharts.length">补充图表(未挂载到具体章节)</h2>
          <div v-if="orphanCharts.length" class="charts">
            <div v-for="(c, i) in orphanCharts" :key="'orphan-' + i" class="chart-box">
              <v-chart class="chart" :option="c.option" autoresize />
            </div>
          </div>
          <div v-else-if="!Object.values(sectionCharts).some(arr => arr && arr.length)" class="dim">无图表数据</div>

          <!-- 关键指标汇总 -->
          <h2>关键指标</h2>
          <table class="metric-table" v-if="hasNumeric">
          <thead>
            <tr><th>指标</th><th>合计</th><th>均值</th><th>最大</th><th>最小</th></tr>
          </thead>
          <tbody>
            <tr v-for="(v, k) in numeric" :key="k">
              <td>{{ k }}</td>
              <td>{{ v.sum }}</td>
              <td>{{ v.mean }}</td>
              <td>{{ v.max }}</td>
              <td>{{ v.min }}</td>
            </tr>
          </tbody>
        </table>

        <!-- 兜底展示:无数值列时,展示缺失率 / 文本列频次 -->
        <div v-else-if="hasFallback" class="metric-fallback">
          <div v-if="Object.keys(missingRatio).length">
            <h3 class="sub-h">缺失率 Top5</h3>
            <table class="metric-table">
              <thead><tr><th>字段</th><th>缺失率</th></tr></thead>
              <tbody>
                <tr v-for="(p, k) in missingRatio" :key="k">
                  <td>{{ k }}</td>
                  <td>{{ p }}%</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div v-for="(items, col) in textValueCounts" :key="col">
            <h3 class="sub-h">{{ col }} · 频次 Top8</h3>
            <table class="metric-table">
              <thead><tr><th>取值</th><th>出现次数</th></tr></thead>
              <tbody>
                <tr v-for="it in items" :key="it.name">
                  <td>{{ it.name }}</td>
                  <td>{{ it.value }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
        <div v-else class="dim">无关键指标</div>

        <!-- 异常批注 -->
        <h2>异常批注</h2>
        <div v-if="annotations.length">
          <div v-for="(a, i) in annotations" :key="i" class="anno" :class="a.level">
            <span class="anno-metric">{{ cleanText(a.metric) }}</span>
            <span class="anno-comment">{{ cleanText(a.comment) }}</span>
          </div>
        </div>
        <div v-else class="dim">无异常指标</div>
        </template>
        <!-- 【dataset-contract-compare-report 新增】template 闭合:上面这一段
             (数据图表/关键指标/异常批注)在合同对比场景下整段隐藏 -->
      </article>

      <!-- 【report-pptx-themes 新增】PPT 导出弹窗:主题选择 + 逐页缩略图 -->
      <el-dialog
        v-model="pptxPreviewOpen"
        title="导出 PPT · 选主题并预览"
        width="1080px"
        append-to-body
        top="4vh"
        :close-on-click-modal="false"
      >
        <div class="pptx-dialog-body">
          <!-- 左栏:主题选择 -->
          <div class="pptx-themes-col">
            <!-- 【AI 演示稿】版本选择:两种导出模式并行,默认「AI 演示稿」。
                 el-radio-button 在 Element Plus 2.6+ 用 :value 绑定选中值(不是旧的 :label) -->
            <div class="pptx-themes-label">版本</div>
            <el-radio-group v-model="pptxSource" class="pptx-themes-radio pptx-source-radio">
              <el-radio-button value="ai" class="pptx-theme-item">
                <div class="pptx-theme-name">
                  AI 演示稿
                  <span class="pptx-recommend-tag">推荐</span>
                </div>
                <div class="pptx-theme-desc">结论式演示,带配图与演讲备注</div>
              </el-radio-button>
              <el-radio-button value="report" class="pptx-theme-item">
                <div class="pptx-theme-name">报告全文版</div>
                <div class="pptx-theme-desc">报告正文按字数切块搬进幻灯片</div>
              </el-radio-button>
            </el-radio-group>

            <div class="pptx-themes-label" style="margin-top: 18px">选择主题</div>
            <el-radio-group v-model="pptxTheme" class="pptx-themes-radio">
              <el-radio-button
                v-for="t in PPTX_THEMES"
                :key="t.key"
                :value="t.key"
                class="pptx-theme-item"
              >
                <div class="pptx-theme-name">{{ t.name }}</div>
                <div class="pptx-theme-desc">{{ t.desc }}</div>
              </el-radio-button>
            </el-radio-group>
          </div>
          <!-- 右栏:逐页缩略图面板 -->
          <div class="pptx-slides-col">
            <!-- 【AI 演示稿】状态区:仅「AI 演示稿」模式显示。
                 把"演示稿是否已生成"这层状态显式暴露出来,避免用户
                 在没有演示稿时直接点导出(后端会现场生成,要等 1~3 分钟)。 -->
            <div v-if="pptxSource === 'ai'" class="pptx-deck-status" :data-status="deckStatus">
              <!-- 未生成 -->
              <template v-if="deckStatus === 'none'">
                <span class="pptx-deck-status-text">演示稿尚未生成</span>
                <el-button
                  type="primary"
                  size="small"
                  :loading="deckGenerating"
                  @click="onGenerateDeck(false)"
                >
                  让 AI 生成演示稿
                </el-button>
                <span class="pptx-deck-status-hint">AI 重构演示稿通常需要 1~3 分钟</span>
              </template>
              <!-- 生成中 -->
              <template v-else-if="deckStatus === 'generating'">
                <span class="pptx-deck-status-text">AI 正在重构演示稿…</span>
                <span class="pptx-deck-status-hint">通常需要 1~3 分钟，请勿关闭</span>
              </template>
              <!-- 已生成 -->
              <template v-else-if="deckStatus === 'ready'">
                <span class="pptx-deck-status-text">已生成 · 共 {{ deckSlides.length }} 页</span>
              </template>
              <!-- 生成失败 -->
              <template v-else-if="deckStatus === 'failed'">
                <span class="pptx-deck-status-text pptx-deck-status-error">
                  生成失败：{{ deckError || '未知原因' }}
                </span>
                <el-button
                  type="primary"
                  size="small"
                  :loading="deckGenerating"
                  @click="onGenerateDeck(true)"
                >
                  重新生成
                </el-button>
              </template>
            </div>

            <div class="pptx-themes-label">
              逐页预览 · 共 {{ pptxSource === 'ai' ? deckSlides.length : pptxSlides.length }} 页
            </div>
            <div
              class="pptx-slides-panel"
              v-loading="pptxSource === 'ai' ? deckLoading : pptxLoading"
              element-loading-text="生成预览中..."
            >
              <!-- ===== 报告全文版:沿用原有渲染逻辑,一行未改 ===== -->
              <template v-if="pptxSource === 'report'">
              <div
                v-for="(s, i) in pptxSlides"
                :key="i"
                class="pptx-slide-card"
                :data-kind="s.kind"
                :style="{
                  background: themeCss.cardBg,
                  borderColor: themeCss.cardBorder,
                  color: themeCss.text,
                }"
              >
                <!-- 顶条 accent:体现主题 accent 色 -->
                <div
                  class="pptx-slide-topbar"
                  :style="{ background: themeCss.accent }"
                ></div>
                <!-- 标题页/尾页 这两类用主题 title_bg 当整页色,模拟 PPT 实际效果 -->
                <div
                  v-if="s.kind === 'title' || s.kind === 'ending'"
                  class="pptx-slide-hero"
                  :style="{
                    background: themeCss.titleBg,
                    color: themeCss.titleText,
                  }"
                >
                  <div class="pptx-slide-hero-title">{{ s.title }}</div>
                  <div v-if="s.body" class="pptx-slide-hero-sub">{{ s.body }}</div>
                </div>
                <template v-else>
                  <div class="pptx-slide-card-no">
                    <span
                      class="pptx-slide-kind-chip"
                      :style="{
                        background: themeCss.pageKindBg,
                        color: themeCss.pageKindColor,
                      }"
                    >{{ kindLabel(s.kind) }}</span>
                    <span
                      class="pptx-slide-page-no"
                      :style="{ color: themeCss.muted }"
                    >第 {{ i + 1 }} 页</span>
                  </div>
                  <div
                    class="pptx-slide-title"
                    :style="{ color: themeCss.cardTitleColor }"
                  >{{ s.title }}</div>
                  <img
                    v-if="s.kind === 'chart' && s.png_b64"
                    :src="'data:image/png;base64,' + s.png_b64"
                    class="pptx-slide-img"
                    alt="chart preview"
                  />
                  <div v-else-if="s.kind === 'chart'" class="dim pptx-slide-img-fallback">
                    渲染失败,以下为摘要: <br />
                    <pre class="pptx-slide-pre">{{ s.body }}</pre>
                  </div>
                  <div
                    v-if="s.body && s.kind !== 'chart'"
                    class="pptx-slide-body"
                    :style="{ color: themeCss.text }"
                  >
                    {{ s.body }}
                  </div>
                </template>
              </div>
              </template>

              <!-- ===== AI 演示稿:一页一个结论 + bullet + 配图 + 讲稿 ===== -->
              <template v-else>
              <div
                v-for="(s, i) in deckSlides"
                :key="i"
                class="pptx-slide-card"
                :data-kind="s.kind"
                :style="{
                  background: themeCss.cardBg,
                  borderColor: themeCss.cardBorder,
                  color: themeCss.text,
                }"
              >
                <div
                  class="pptx-slide-topbar"
                  :style="{ background: themeCss.accent }"
                ></div>
                <!-- 封面页/结束页:与报告全文版的 title/ending 处理一致,整页用主题 title_bg -->
                <div
                  v-if="s.kind === 'cover' || s.kind === 'ending'"
                  class="pptx-slide-hero"
                  :style="{
                    background: themeCss.titleBg,
                    color: themeCss.titleText,
                  }"
                >
                  <div class="pptx-slide-hero-title">{{ s.title }}</div>
                  <div v-if="s.bullets && s.bullets.length" class="pptx-slide-hero-sub">
                    {{ s.bullets[0] }}
                  </div>
                </div>
                <template v-else>
                  <div class="pptx-slide-card-no">
                    <span
                      class="pptx-slide-kind-chip"
                      :style="{
                        background: themeCss.pageKindBg,
                        color: themeCss.pageKindColor,
                      }"
                    >{{ deckKindLabel(s.kind) }}</span>
                    <span
                      class="pptx-slide-page-no"
                      :style="{ color: themeCss.muted }"
                    >第 {{ i + 1 }} 页</span>
                  </div>
                  <div
                    class="pptx-slide-title"
                    :style="{ color: themeCss.cardTitleColor }"
                  >{{ s.title }}</div>
                  <!-- bullet 列表:一条一行,圆点前缀 -->
                  <ul
                    v-if="s.bullets && s.bullets.length"
                    class="pptx-slide-bullets"
                    :style="{ color: themeCss.text }"
                  >
                    <li v-for="(b, bi) in s.bullets" :key="bi">{{ b }}</li>
                  </ul>
                  <!-- 配图(chart_ref 命中图表目录时后端会渲染 base64 PNG) -->
                  <img
                    v-if="s.chart_png_b64"
                    :src="'data:image/png;base64,' + s.chart_png_b64"
                    class="pptx-slide-img"
                    alt="chart preview"
                  />
                  <!-- 演讲备注:提示这页怎么讲。
                       颜色取主题的 muted 色,而不是写死灰色 ——
                       暗夜主题下写死 #6b7280 在 #0f172a 上对比度只有 3.69:1,
                       小字号偏弱,跟着主题走才能都读得清。 -->
                  <div
                    v-if="s.notes"
                    class="pptx-slide-notes"
                    :style="{ color: themeCss.muted }"
                  >
                    讲稿：{{ s.notes }}
                  </div>
                </template>
              </div>
              </template>

              <!-- 空态提示:两种模式各判各的,避免切模式时残留上一种模式的数据判断 -->
              <div
                v-if="!pptxLoading && pptxSource === 'report' && pptxSlides.length === 0"
                class="dim pptx-empty"
              >
                暂无预览内容
              </div>
              <div
                v-if="!deckLoading && pptxSource === 'ai' && deckSlides.length === 0"
                class="dim pptx-empty"
              >
                暂无预览内容
              </div>
            </div>
          </div>
        </div>
        <template #footer>
          <el-button @click="pptxPreviewOpen = false">取消</el-button>
          <!-- 讲义导出:只有演示稿已生成才有意义 -->
          <el-button
            v-if="pptxSource === 'ai' && deckStatus === 'ready'"
            @click="downloadDeckScript"
          >
            导出讲义（Markdown）
          </el-button>
          <el-button
            type="primary"
            :loading="pptxSource === 'ai' ? deckLoading : pptxLoading"
            :disabled="pptxDownloadDisabled"
            @click="confirmDownloadPptx"
          >
            {{ downloadButtonText }}
          </el-button>
        </template>
      </el-dialog>

      <!-- 【report-feedback-and-versioning 新增】「让 AI 重新写一份」意见抽屉 -->
      <el-card class="feedback-card" style="margin-top: 24px">
        <template #header>
          <div style="display: flex; align-items: center; gap: 8px">
            <span style="font-weight: 600">📝 让 AI 重新写一份</span>
            <el-tag size="small" type="info" effect="plain">保留原版本</el-tag>
          </div>
        </template>
        <p class="dim small" style="margin: 0 0 10px">
          告诉 AI 你希望怎么调整这份报告,想说多少说多少。提交后 AI 会基于已有数据重新生成,旧版本保留在历史任务里。
        </p>
        <el-input
          v-model="feedback"
          type="textarea"
          :rows="3"
          placeholder="例如:把销售那段改成 3 个 bullet;问题这块按严重度排序;把多余的'综上'都删掉..."
        />
        <div style="margin-top: 10px; display: flex; justify-content: flex-end">
          <el-button
            type="primary"
            :loading="regenerating"
            :disabled="!feedback.trim()"
            @click="onRegenerate"
          >
            基于以上意见重新生成
          </el-button>
        </div>
      </el-card>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { MagicStick, Download, Star, InfoFilled } from '@element-plus/icons-vue'
import { use } from 'echarts/core'
import { CanvasRenderer } from 'echarts/renderers'
import { LineChart, BarChart, PieChart } from 'echarts/charts'
import { TitleComponent, TooltipComponent, LegendComponent, GridComponent } from 'echarts/components'
import VChart from 'vue-echarts'
import { tasksApi, reportsApi, toolsApi, downloadBlobResponse } from '@/api'
// 【export-visual-report 新增】离屏 ECharts 高保真截图通道（报表导出主路径）
import { captureChartsPng } from '@/utils/chartExport'

use([CanvasRenderer, LineChart, BarChart, PieChart, TitleComponent, TooltipComponent, LegendComponent, GridComponent])

// =====【report-pptx-themes 新增】PPT 导出主题多选 =====
// PPT 弹窗里 4 套主题的展示元数据;key 与后端 _PPT_THEMES 一一对应。
const PPTX_THEMES = [
  { key: 'deep_cyan', name: '深青主题', desc: '项目品牌色,默认' },
  { key: 'midnight',  name: '暗夜主题', desc: '暗色背景,适合全屏汇报' },
  { key: 'light',     name: '浅色主题', desc: '白底简版,印刷友好' },
  { key: 'business',  name: '商务主题', desc: '深蓝金黄,适合对外提案' },
]

const route = useRoute()
const router = useRouter()
const taskId = Number(route.params.id)
const task = ref(null)
const loading = ref(true)
const summary = ref('')
const summarizing = ref(false)

// =====【report-feedback-and-versioning 新增】=====
// 「让 AI 重新写一份」反馈抽屉的本地状态 + lineage 导航
const feedback = ref('')
const regenerating = ref(false)

// =====【report-pptx-themes 新增】PPT 弹窗状态 =====
const pptxPreviewOpen = ref(false)
const pptxTheme = ref('deep_cyan')
const pptxSlides = ref([])
const pptxLoading = ref(false)
const currentThemeName = computed(
  () => PPTX_THEMES.find(t => t.key === pptxTheme.value)?.name || pptxTheme.value
)

// =====【AI 演示稿 新增】导出模式与演示稿状态 =====
// pptxSource: 'ai' = AI 演示稿(结论式,默认) | 'report' = 报告全文版(老逻辑)
// deckStatus : 'none' 未生成 | 'generating' 生成中 | 'ready' 已生成 | 'failed' 失败
const pptxSource = ref('ai')
const deckStatus = ref('none')
const deckSlides = ref([])
const deckLoading = ref(false)
const deckGenerating = ref(false)
const deckError = ref('')

// 当前"模式 + 主题"的数据签名,以及已发起过加载的签名。
// 为什么要这个:openPptxPreview 里会主动拉一次数据,而 watch(pptxSource)/
// watch(pptxTheme) 在弹窗打开那一刻也会触发 —— 不做去重就会重复发请求。
// 签名一致说明这个组合已经拉过了,watch 里直接跳过。
const pptxDataKey = computed(() => `${pptxSource.value}:${pptxTheme.value}`)
const pptxLoadedKey = ref('')

// 请求序号:快速切换主题/版本时,后发起的请求先返回,先发起的响应就成了过期数据。
// 用自增序号做"谁最新谁生效"判断,过期响应直接丢弃,避免缩略图与当前主题对不上。
// 两种模式各用一个序号:切换模式时,旧模式的过期响应不能影响新模式的数据。
let deckReqSeq = 0
let reportReqSeq = 0

// 演示稿页型 → 中文标签(AI 演示稿专用,与报告全文版的 kindLabel 分开)
const DECK_KIND_LABELS = {
  cover: '封面',
  agenda: '目录',
  content: '内容',
  chart: '图表',
  metrics: '关键指标',
  conclusion: '结论',
  ending: '尾页',
}
function deckKindLabel(kind) {
  return DECK_KIND_LABELS[kind] || kind
}

// 下载按钮文案:让用户一眼知道下的是哪一版
const downloadButtonText = computed(() => {
  const modeName = pptxSource.value === 'ai' ? 'AI 演示稿' : '报告全文版'
  return `下载 PPT（${modeName} · ${currentThemeName.value}）`
})

// 下载按钮禁用条件:
//  - 报告全文版:预览为空时不可下载(与老逻辑一致)
//  - AI 演示稿:必须已生成。导出接口在无缓存时会现场生成,要等 1~3 分钟,
//    体验极差,所以前端强制先在弹窗里生成好再放行下载。
const pptxDownloadDisabled = computed(() => {
  if (pptxSource.value === 'ai') {
    return deckStatus.value !== 'ready' || deckLoading.value || deckGenerating.value
  }
  return pptxLoading.value || pptxSlides.value.length === 0
})

// 前端主题色表（与后端 _PPT_THEMES 字段对齐）——
// 让预览面板真的按当前主题上色,4 套主题能看出明显区别。
const THEME_CSS = {
  deep_cyan: {
    titleBg: '#0e7490', accent: '#22d3ee', text: '#1f2937',
    titleText: '#ffffff', background: '#ffffff',
    muted: '#6b7280', cardBg: '#ffffff', cardBorder: '#d1d5db',
    cardTitleColor: '#0e7490', pageKindBg: '#0e7490', pageKindColor: '#ffffff',
  },
  midnight: {
    titleBg: '#0f172a', accent: '#7dd3fc', text: '#e2e8f0',
    titleText: '#ffffff', background: '#0f172a',
    muted: '#94a3b8', cardBg: '#0f172a', cardBorder: '#1e293b',
    cardTitleColor: '#7dd3fc', pageKindBg: '#1e293b', pageKindColor: '#e2e8f0',
  },
  light: {
    titleBg: '#ffffff', accent: '#0e7490', text: '#1f2937',
    titleText: '#0e7490', background: '#ffffff',
    muted: '#6b7280', cardBg: '#ffffff', cardBorder: '#d1d5db',
    cardTitleColor: '#0e7490', pageKindBg: '#f3f4f6', pageKindColor: '#0e7490',
  },
  business: {
    titleBg: '#1e3a8a', accent: '#fbbf24', text: '#111827',
    titleText: '#ffffff', background: '#f5f1e6',
    muted: '#4b5563', cardBg: '#f5f1e6', cardBorder: '#d1cebd',
    cardTitleColor: '#1e3a8a', pageKindBg: '#1e3a8a', pageKindColor: '#fbbf24',
  },
}
const themeCss = computed(
  () => THEME_CSS[pptxTheme.value] || THEME_CSS.deep_cyan
)

async function loadPptxPreview() {
  // 报告全文版的预览逻辑
  pptxLoading.value = true
  // 请求序号守卫:与演示稿侧同一套机制。
  // 快速切主题/切版本时,先发起的请求可能后返回,如果不丢弃它的响应,
  // 它会把 pptxSlides 覆盖成空数组 —— 表现为「切到报告全文版后预览是空的」。
  const seq = ++reportReqSeq
  try {
    const slides = (await reportsApi.previewCharts(taskId, pptxTheme.value)).slides || []
    if (seq !== reportReqSeq) return
    pptxSlides.value = slides
  } catch (e) {
    // http.js 已经统一 ElMessage.error 提示
    if (seq !== reportReqSeq) return
    pptxSlides.value = []
  } finally {
    if (seq === reportReqSeq) pptxLoading.value = false
  }
}

// =====【AI 演示稿 新增】拉取演示稿缩略图 =====
/**
 * GET /reports/{id}/deck?theme=xxx
 * 后端在还没有演示稿时返回 {status:'none', deck:null, slides:[]} 且不报错,
 * 所以这里要把 status 一并映射到前端状态机:
 *   none  → 'none'(显示"生成演示稿"入口)
 *   ready → 'ready'(渲染缩略图)
 */
async function loadDeckPreview() {
  deckLoading.value = true
  // 先占位签名,避免与 watch 触发的加载重复请求(见 pptxLoadedKey 注释)
  pptxLoadedKey.value = pptxDataKey.value
  // 请求序号:连续切主题时旧响应可能后到,用它丢弃过期响应
  const seq = ++deckReqSeq
  try {
    const resp = await reportsApi.deckPreview(taskId, pptxTheme.value)
    if (seq !== deckReqSeq) return
    if (resp?.status === 'ready') {
      deckSlides.value = resp.slides || []
      deckStatus.value = 'ready'
      deckError.value = ''
    } else {
      // 后端明确说"没有已就绪的演示稿"。
      // 但如果本地正处于"生成中"(比如本窗口刚点过生成、或别的会话在跑),
      // 不能把状态打回"未生成",否则会误显示"生成"按钮、让用户重复发起。
      deckSlides.value = []
      if (deckStatus.value !== 'generating') {
        deckStatus.value = 'none'
        deckError.value = ''
      }
    }
  } catch (e) {
    // http.js 已经统一 ElMessage.error 提示
    if (seq !== deckReqSeq) return
    deckSlides.value = []
    // 请求失败(网络/服务端异常)不代表演示稿不存在:
    // 若本地已是"已生成"或"生成中",保持原状态,否则回落到"未生成"。
    // (典型场景:刚生成成功、紧接着拉缩略图失败 —— 演示稿其实已落库,
    //  若此时打回"未生成"会把下载按钮误禁用。)
    if (deckStatus.value !== 'generating' && deckStatus.value !== 'ready') {
      deckStatus.value = 'none'
    }
  } finally {
    if (seq === deckReqSeq) deckLoading.value = false
  }
}

/**
 * 生成 AI 演示稿。
 * force=true 时忽略后端缓存强制重新生成(用于"重新生成"按钮)。
 * 该接口会调大模型,耗时 60~180 秒,期间禁用按钮避免重复触发。
 */
async function onGenerateDeck(force = false) {
  if (deckGenerating.value) return
  deckGenerating.value = true
  // 先切到生成中,让用户立刻看到反馈(接口是长耗时同步请求)
  deckStatus.value = 'generating'
  deckError.value = ''
  try {
    const resp = await reportsApi.generateDeck(taskId, force)
    const st = resp?.status
    if (st === 'ready') {
      deckStatus.value = 'ready'
      // 生成成功后拉一次缩略图(带 base64 配图),把预览面板填满
      await loadDeckPreview()
    } else if (st === 'failed') {
      deckStatus.value = 'failed'
      deckError.value = resp?.error || '演示稿生成失败'
      ElMessage.error(deckError.value)
    } else if (st === 'generating') {
      // 后端说已经有人在生成(10 分钟内),保持"生成中"不重复发起
      deckStatus.value = 'generating'
    } else {
      // 未知状态:保守处理成失败,提示用户重试
      deckStatus.value = 'failed'
      deckError.value = '未知的生成状态'
    }
  } catch (e) {
    // http.js 已经统一 ElMessage.error 提示,这里只更新状态机
    deckStatus.value = 'failed'
    deckError.value = e?.response?.data?.detail || e?.message || '演示稿生成失败'
  } finally {
    deckGenerating.value = false
  }
}

// 导出讲义(Markdown):后端直接返回 text/markdown 附件,
// 用 a 标签触发下载(与 datasetsApi.download 同思路;这里后端已带
// Content-Disposition 文件名,window.open 即可,这里用 a[download] 更稳)
function downloadDeckScript() {
  const url = reportsApi.deckScriptUrl(taskId)
  const a = document.createElement('a')
  a.href = url
  a.download = `report_${taskId}_script.md`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
}

function openPptxPreview() {
  // 主题状态重置,避免误用上次选择
  pptxTheme.value = 'deep_cyan'
  // 默认走「AI 演示稿」模式
  pptxSource.value = 'ai'
  // 重置数据签名:本次打开需要重新拉数据
  pptxLoadedKey.value = ''
  pptxPreviewOpen.value = true
  // 统一走分发函数:'ai' 拉演示稿,'report' 拉报告全文版预览
  loadCurrentPptxData()
}

function confirmDownloadPptx() {
  // 按当前模式决定导出来源:AI 演示稿走 source=ai,报告全文版保持老行为(不传 source)
  const source = pptxSource.value === 'ai' ? 'ai' : 'report'
  pptxPreviewOpen.value = false
  window.open(reportsApi.exportUrl(taskId, 'pptx', pptxTheme.value, source), '_blank')
}

/**
 * 按当前模式加载对应的预览数据。
 * pptxLoadedKey 做去重:同一个"模式+主题"组合只拉一次,
 * 避免 openPptxPreview 的主动加载 与 watch 的自动加载 重复发请求。
 */
function loadCurrentPptxData() {
  if (pptxLoadedKey.value === pptxDataKey.value) return
  // 先记录签名再发起请求:并发 watch 触发时靠这个值挡掉重复请求
  pptxLoadedKey.value = pptxDataKey.value
  if (pptxSource.value === 'ai') {
    loadDeckPreview()
  } else {
    loadPptxPreview()
  }
}

// 切换主题立即重新拉预览
watch(pptxTheme, (val) => {
  if (!pptxPreviewOpen.value) return
  // 按当前模式走各自的数据源:'ai' 拉演示稿,'report' 拉报告全文版预览
  loadCurrentPptxData()
})

// =====【AI 演示稿 新增】切换导出版本 =====
// 切到 'ai' 拉演示稿缩略图;切到 'report' 拉报告全文版预览(老逻辑一行未改)
watch(pptxSource, (val) => {
  if (!pptxPreviewOpen.value) return
  loadCurrentPptxData()
})

// =====【report-pptx-themes 新增】页型 → 中文标签(用于缩略图卡片标题栏) =====
function kindLabel(kind) {
  return {
    title: '标题',
    summary: '摘要',
    section: '章节',
    chart: '图表',
    metrics: '关键指标',
    fallback: '字段明细',
    annotation: '异常批注',
    ending: '尾页',
  }[kind] || kind
}

// 兜底剥离 <think>...</think> 段,避免 reasoning 模型的"思考过程"展示给用户。
// 后端 ai_service 已做一次剥离,这里是前端二次保险(兼容历史脏数据)。
const THINK_RE = /<think>[\s\S]*?<\/think>|<thinks?>[\s\S]*?<\/thinks?>|<think>[\s\S]*?<\/thinks?>/gi
function cleanText(s) {
  if (typeof s !== 'string') return s || ''
  const out = s.replace(THINK_RE, '').replace(/\n{3,}/g, '\n\n').trim()
  return out || s.trim()
}

// 【report-sections-structured 新增】合同报告 sections 渲染辅助:
// - contractChangeTypeTag: 合同 change_type → 标签(与 WorkbenchView 一致)
// - formatListItem: 通用 list 兜底,处理 dict 类型的列表项
// - formatObjectText: 兜底 dict 类型的 section,JSON 化展示
const _CONTRACT_CHANGE_META = {
  modified:   { type: 'warning', label: '已变更' },
  added:      { type: 'success', label: '新增' },
  removed:    { type: 'danger',  label: '删除' },
  unchanged:  { type: 'info',    label: '未变' },
  unclear:    { type: '',        label: '不确定' },
}
function contractChangeTypeTag(t) {
  return _CONTRACT_CHANGE_META[t] || { type: '', label: t || '—' }
}
function formatListItem(item) {
  if (item == null) return ''
  if (typeof item === 'string') return item
  if (typeof item === 'object') {
    // 优先用 title/label/name,否则 JSON 化
    return item.title || item.label || item.name || JSON.stringify(item)
  }
  return String(item)
}
function formatObjectText(obj) {
  if (obj == null) return ''
  if (typeof obj === 'string') return obj
  try { return JSON.stringify(obj, null, 2) } catch (e) { return String(obj) }
}

// 【row-diff-detail-expansion 新增】提取一组行的列名(按出现顺序去重)
// - 数据集对比·变更追踪的 added/removed 是 list[dict],不同行可能字段不同,
//   这里把多行字段合并并保序,作为表头传给 <th v-for>。
function getDiffColumns(rows) {
  const seen = new Set()
  const cols = []
  for (const r of rows || []) {
    if (!r || typeof r !== 'object') continue
    for (const k of Object.keys(r)) {
      if (!seen.has(k)) {
        seen.add(k)
        cols.push(k)
      }
    }
  }
  return cols
}

// 章节标题清洗:后端 _parse_section_titles 已经把章节名里的"一、""1.""##" 等
// 前缀剥掉了,但用户可能填的是「一、执行摘要」整体传给 LLM,LLM 输出也是同款;
// 这里再做一次兜底:把"一、" "二、" "1." "1、" "(1)" 等编号前缀剥掉,
// 让显示更紧凑、可读性更好(不影响后端存储与匹配)。
const SECTION_NUM_PREFIX_RE = /^\s*(?:[一二三四五六七八九十百千]+|[0-9]+|[①-⑳]|[（(]\s*[一二三四五六七八九十0-9]+\s*[)）])\s*[、,.\)::]\s*/
function cleanKey(key) {
  if (typeof key !== 'string') return key || ''
  // 先剥离 <think> 等思考标签(同 cleanText)
  const stripped = key.replace(THINK_RE, '').trim()
  return stripped.replace(SECTION_NUM_PREFIX_RE, '').trim() || stripped
}

// 【章节缺图修复 · 前端兜底】把章节名宽松化。
// 与后端 ai_service._loosen_section_key 策略保持一致:

const report = computed(() => (task.value?.report?.sections ? { title: task.value.report.title, sections: task.value.report.sections } : { sections: {} }))
// 【dataset-contract-compare-report 新增】识别合同对比报告,
// 用于隐藏"数据图表/关键指标/异常批注"等对合同场景无意义且全空的固定板块。
// 优先读 report.mode(后端在 contract_compare 模式下显式写入);兜底匹配 task.name,
// 兼容历史 / 别名场景。
const isContractReport = computed(() => {
  const m = task.value?.report?.mode
  if (m === 'contract_compare') return true
  const name = task.value?.name || ''
  return name.includes('合同对比') || name.includes('·合同对比')
})

// 【pdf-tools-history-persistence 新增】PDF 工具产物元信息
// 后端在落库时把 file_name / download_url / size / kind / label 写到 report.pdf_artifact,
// ReportView 识别后渲染"下载产物"按钮,直接调 toolsApi.downloadFile 触发浏览器保存
const pdfArtifact = computed(() => task.value?.report?.pdf_artifact || null)
const downloadingArtifact = ref(false)
async function downloadPdfArtifact() {
  const art = pdfArtifact.value
  if (!art || !art.download_url) {
    ElMessage.warning('产物信息缺失')
    return
  }
  downloadingArtifact.value = true
  try {
    // download_url 是后端返回的纯相对路径 /tools/pdf/download/...,
    // toolsApi.pdfDownloadUrl 拼的是 /tools/pdf/download/<file>,
    // 这里直接用 art.download_url(已带 file_name),toolsApi.downloadFile 用 http.get 即可
    await toolsApi.downloadFile(art.download_url, art.file_name || 'download')
    ElMessage.success('下载已开始')
  } catch (e) {
    ElMessage.error('下载失败')
    console.error('pdf artifact download failed:', e)
  } finally {
    downloadingArtifact.value = false
  }
}
function formatArtifactSize(bytes) {
  if (!bytes && bytes !== 0) return '—'
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`
}
const indicators = computed(() => task.value?.indicators || {})
const numeric = computed(() => indicators.value.numeric_summary || {})
// 兜底数据:数值列为空时,展示缺失率 / 文本列频次,避免"无关键指标"空块
const missingRatio = computed(() => indicators.value.missing_ratio_top5 || {})
const textValueCounts = computed(() => indicators.value.text_value_counts || {})
const hasNumeric = computed(() => Object.keys(numeric.value).length > 0)
const hasFallback = computed(() => Object.keys(missingRatio.value).length > 0 || Object.keys(textValueCounts.value).length > 0)
// 图表数据拆分为「按章节挂载」+「兜底汇总」,覆盖新旧两种 chart_option:
//   - 新 shape:{by_section: {<章节名>: [...]}} → sectionCharts 按章节名挂载
//   - 旧 shape:{charts: [...]} → 视为 by_section["数据图表"],进入 orphanCharts
//   - LLM 直接给完整 option(有 series) → 进入 orphanCharts
//
// 【2026-10 章节缺图修复】sectionCharts / orphanCharts 的宽容匹配已下放给后端
// analyzer._normalize_by_section_keys 统一处理,前端只做"按字段取数"的简单职责,
// 不再做二次匹配 —— 这样后端落库时 by_section 已对齐 sections.key,前端无需
// 重复实现一遍宽窄策略(避免双端实现版本漂移)。
const chartOption = computed(() => task.value?.chart_option || {})
const sectionCharts = computed(() => {
  const opt = chartOption.value
  if (opt && typeof opt === 'object' && opt.by_section && typeof opt.by_section === 'object') {
    return opt.by_section || {}
  }
  return {}
})
const orphanCharts = computed(() => {
  const opt = chartOption.value
  if (!opt || typeof opt !== 'object') return []
  if (opt.by_section && typeof opt.by_section === 'object') {
    const sectionKeys = Object.keys(report.value.sections || {})
    const orphan = []
    for (const [k, arr] of Object.entries(opt.by_section)) {
      if (!Array.isArray(arr)) continue
      if (sectionKeys.includes(k)) continue
      for (const c of arr) orphan.push(c)
    }
    return orphan
  }
  if (Array.isArray(opt.charts)) return opt.charts
  if (opt.series) return [opt]
  return []
})

// =====【2026-10 新增】章节明细数据表 =====
// 后端 chart_option.detail_tables:{章节名: {title, columns, rows, note, total_rows}}。
// 章节 key 用宽松匹配回退(与 sectionCharts 同理):后端已按 sections.key
// 精确对齐,这里只做"取不到就按 cleaned key 再试一次"的兜底,
// 避免章节名带编号前缀时表格挂不上。
const detailTablesAll = computed(() => {
  const opt = chartOption.value
  if (opt && typeof opt === 'object' && opt.detail_tables && typeof opt.detail_tables === 'object') {
    return opt.detail_tables
  }
  return {}
})
const sectionDetailTables = computed(() => {
  const all = detailTablesAll.value || {}
  const exact = {}
  for (const [k, v] of Object.entries(all)) {
    exact[k] = v
  }
  // 宽松兜底:sections.key 与 detail_tables 的 key 未能精确对齐时,
  // 用 cleanKey 归一化后再匹配一次(例如"一、总体经营概况" vs "总体经营概况")
  const norm = {}
  for (const [k, v] of Object.entries(all)) {
    const nk = cleanKey(k)
    if (nk && !norm[nk]) norm[nk] = v
  }
  return { ...exact, ...norm }
})
// 取章节的明细表:优先精确 key,再试清洗后的 key
function getSectionDetailTable(key) {
  const exact = sectionDetailTables.value[key]
  if (exact) return exact
  const cleaned = cleanKey(key)
  return cleaned ? sectionDetailTables.value[cleaned] : null
}
function detailTableNote(key) {
  const t = getSectionDetailTable(key)
  if (!t) return ''
  const total = t.total_rows || (t.rows ? t.rows.length : 0)
  const shown = t.rows ? t.rows.length : 0
  const parts = []
  if (t.note) parts.push(t.note)
  if (total > shown) parts.push(`共 ${total} 行,仅展示前 ${shown} 行`)
  return parts.join(' ')
}
// 表格单元格展示:null → "—";数字去掉无意义小数尾巴(4680.0 → 4680)
// =====【2026-10 新增】章节正文分段 =====
// 需求:「章节要分段,不要全部都在一段上」。
// 三层策略(与后端 analyzer.split_paragraphs 同一套口径):
//   ① 尊重模型已有的空行分段;
//   ② 单段超过 maxLen 时按句号切,再合并成 ~maxLen 的段;
//   ③ 段内遇到「其一/其二/首先/其次」等语义标记时隐式拆段 ——
//      实测模型常写「结论。其一…其二…」这种流水句,
//      标记词是唯一的结构信号,不拆就会糊成一整块。
const PARA_MARKERS = [
  '其一', '其二', '其三', '其四',
  '第一', '第二', '第三', '第四',
  '首先', '其次', '再次', '最后',
  '一方面', '另一方面',
  '结论', '核心结论', '总体来看', '综上',
]
const PARA_MAX_LEN = 200

function splitByLength(blk, maxLen) {
  if (blk.length <= maxLen) return [blk]
  const sents = blk.split(/(?<=[。!?])/).map(s => s.trim()).filter(Boolean)
  const out = []
  let cur = ''
  for (const s of sents) {
    if (!cur) cur = s
    else if (cur.length + s.length <= maxLen) cur += s
    else { out.push(cur); cur = s }
  }
  if (cur) out.push(cur)
  return out
}

function splitParagraphs(text) {
  if (!text || typeof text !== 'string') return []
  const t = cleanText(text)
  if (!t) return []
  // ① 已有空行分段 → 尊重模型意图
  let blocks = t.split(/\n\s*\n/).map(b => b.trim()).filter(Boolean)
  if (blocks.length <= 1 && t.includes('\n')) {
    blocks = t.split('\n').map(b => b.trim()).filter(Boolean)
  }
  if (blocks.length <= 1) blocks = [t]
  // ② 超长段按句号切
  let out = []
  for (const b of blocks) out = out.concat(splitByLength(b, PARA_MAX_LEN))
  // ③ 按语义标记隐式拆段
  const finalParas = []
  for (const blk of out) {
    const sents = blk.split(/(?<=[。!?])/).map(s => s.trim()).filter(Boolean)
    let cur = ''
    for (const s of sents) {
      const isMarker = PARA_MARKERS.some(m => s.startsWith(m))
      if (cur && (isMarker || cur.length >= PARA_MAX_LEN)) {
        finalParas.push(cur)
        cur = s
      } else {
        cur += s
      }
    }
    if (cur) finalParas.push(cur)
  }
  return finalParas.length ? finalParas : [t]
}

function formatCell(cell) {
  if (cell === null || cell === undefined || cell === '') return '—'
  if (typeof cell === 'number') {
    if (!Number.isFinite(cell)) return '—'
    if (Number.isInteger(cell)) return String(cell)
    return String(parseFloat(cell.toFixed(4)))
  }
  return String(cell)
}
// 【按章节挂载 · 补章】本报告是否"至少有一张图":用于控制"章节缺图"占位提示的显示;
// 若整份报告都没图,占位提示刷屏会干扰阅读,故此场景下整体隐藏。
const hasAnyCharts = computed(() => {
  const secHas = Object.values(sectionCharts.value || {}).some(
    arr => Array.isArray(arr) && arr.length > 0
  )
  return secHas || orphanCharts.value.length > 0
})
const annotations = computed(() => task.value?.annotations || [])

async function load() {
  loading.value = true
  try {
    task.value = await tasksApi.detail(taskId)
    summary.value = cleanText(task.value.report.summary || '')
  } finally {
    loading.value = false
  }
}

async function onSummarize() {
  summarizing.value = true
  try {
    // 旧实现只 refresh,不调模型 → 报告页"生成摘要"按钮是个空操作
    // 现在真正调用后端 summarize 接口,触发一次 LLM 调用,并把最新 task 回填
    const detail = await tasksApi.summarize(taskId)
    task.value = detail
    summary.value = cleanText(detail.report?.summary || '')
    ElMessage.success('摘要已生成')
  } catch (e) {
    // http.js 已统一 ElMessage.error 提示
  } finally {
    summarizing.value = false
  }
}

// =====【report-feedback-and-versioning 新增】=====
// 后端会在详情响应里返回 version_label("" / "v2" / "v3"),这里转为纯数字方便模板展示
const versionLabelText = computed(() => {
  const v = task.value?.version_label || ''
  return v ? v.replace(/^v/, '') : ''
})
// 完整 lineage 标签(基线显示"v1 (原始版)",有父时显示后端给的 version_label 或派生值)
const displayVersionTag = computed(() => {
  const v = task.value?.version_label || ''
  if (!v) return { kind: 'baseline', text: 'v1 (原始版)' }
  return { kind: 'child', text: v }
})

// 意见反馈抽屉:用户输入反馈 → 后端 regenerate → 弹窗让用户选「查看新版本」或「留在当前」
async function onRegenerate() {
  const fb = feedback.value.trim()
  if (!fb) return
  regenerating.value = true
  try {
    const newTask = await tasksApi.regenerate(taskId, fb)
    // regenerate 返回的是新 task 的完整详情,显示成功信息并问用户是否跳转
    try {
      const action = await ElMessageBox.confirm(
        `已生成新版本:${newTask.id},是否立即查看?`,
        'AI 已重写',
        {
          confirmButtonText: '查看新版本',
          cancelButtonText: '留在当前',
          type: 'success',
        },
      )
      if (action === 'confirm') {
        router.push('/reports/' + newTask.id)
      } else {
        ElMessage.success(`已生成 v${versionLabelText.value || ''} 后的新版,请在历史任务查看 #${newTask.id}`)
      }
    } catch (e) {
      // 用户点 × 关掉了弹窗:同样给个友好提示
      ElMessage.success(`已生成新版本 #${newTask.id},可在历史任务查看`)
    }
    // 清空反馈框(因为已经提交,新版本是在另一页查看了)
    feedback.value = ''
  } catch (e) {
    // http.js 已经统一处理 ElMessage.error
  } finally {
    regenerating.value = false
  }
}

// 跳转到上一版本(parent_task_id)
function goToParent() {
  const pid = task.value?.parent_task_id
  if (pid) router.push('/reports/' + pid)
}

// 【report-failed-state 新增】失败提示标题 + 重试
const retrying = ref(false)
const failedTitle = computed(() => {
  const id = task.value?.id
  return id ? `任务 #${id} 生成失败` : '报告生成失败'
})
async function onRetryFailed() {
  const id = task.value?.id
  if (!id) return
  retrying.value = true
  try {
    await tasksApi.reanalyze(id)
    ElMessage.success('已重新提交,正在重跑 AI 分析')
    // 重新加载,让 user 看到 status=processing 的过渡态
    await load()
  } catch (e) {
    // http.js 已统一错误提示
  } finally {
    retrying.value = false
  }
}

/**
 * 触发浏览器下载报告。
 * format: 'pdf' | 'docx' | 'pptx'
 * theme:  仅 format === 'pptx' 时生效,默认 'deep_cyan';其它格式忽略
 */
function exportReport(format, theme = 'deep_cyan') {
  window.open(reportsApi.exportUrl(taskId, format, theme), '_blank')
}

// ============================================================
// 【export-visual-report 新增】导出 Excel 可视化报表（报告页入口）
// ------------------------------------------------------------
// 数据流：
//   1) 前端把页面上的图表（章节图 + 孤儿图）用**离屏 ECharts 实例**截成高保真 PNG；
//      截图走 chartExport.captureChartsPng()，**不触碰页面实例**，所以页面样式不变。
//   2) 只把 { key, png_base64 } 提交后端；指标 / 章节明细 / AI 摘要由后端直接读
//      task.indicators_json / chart_option_json / report_content，**不再调 LLM**。
//   3) key 约定（必须与后端 A8 的归一化完全一致）：
//        分组名 = 章节名（by_section 的 key）；孤儿图统一归入「数据图表」组；
//        单图 key = `${分组名}|${组内序号}`，序号从 1 开始。
//      后端按同一约定匹配截图，匹配不上的图自动走 matplotlib 兜底渲染。
// ============================================================
const exportingVisualReport = ref(false)

/**
 * 收集「页面当前展示的所有图表」并按后端 key 约定分组。
 * 为什么不用 sectionCharts + orphanCharts 直接拼：orphanCharts 已把多个
 * 非章节 key 的图**拍平混在一起**，丢掉了分组信息，无法反推 key。
 * 因此这里从原始 chart_option 重新按 by_section 遍历，与后端口径保持一致。
 */
function collectReportCharts() {
  const opt = chartOption.value
  const groups = [] // [{ group, charts: [{ title, option, key }] }]
  if (!opt || typeof opt !== 'object') return groups

  const sectionKeys = Object.keys(report.value.sections || {})
  const pushGroup = (groupName, arr) => {
    if (!Array.isArray(arr) || !arr.length) return
    groups.push({
      group: groupName,
      charts: arr.map((c, i) => ({
        title: (c && c.title) || '',
        option: (c && c.option) || c || {},
        // 序号从 1 开始，与后端 f"{group}|{index}" 约定一致
        key: `${groupName}|${i + 1}`,
      })),
    })
  }

  if (opt.by_section && typeof opt.by_section === 'object') {
    // 1) 章节内图表：key 就是章节名
    for (const [k, arr] of Object.entries(opt.by_section)) {
      if (!Array.isArray(arr)) continue
      if (sectionKeys.includes(k)) pushGroup(k, arr)
    }
    // 2) 孤儿图（不在章节列表里的 key）统一归入「数据图表」组，保持组内原顺序
    const orphan = []
    for (const [k, arr] of Object.entries(opt.by_section)) {
      if (!Array.isArray(arr)) continue
      if (sectionKeys.includes(k)) continue
      for (const c of arr) orphan.push(c)
    }
    pushGroup('数据图表', orphan)
    return groups
  }
  // 3) 旧形态：顶层 charts 数组 / 单图（opt 自带 series）
  if (Array.isArray(opt.charts)) {
    pushGroup('数据图表', opt.charts)
  } else if (opt.series) {
    pushGroup('数据图表', [opt])
  }
  return groups
}

async function exportVisualReport() {
  if (exportingVisualReport.value) return
  const groups = collectReportCharts()
  const total = groups.reduce((n, g) => n + g.charts.length, 0)
  if (!total) {
    ElMessage.warning('当前报告没有可导出的图表')
    return
  }
  exportingVisualReport.value = true
  try {
    // 拍平后一次性串行截图（单张失败只留 null，不阻断整批）
    const flat = groups.flatMap((g) => g.charts)
    const shots = await captureChartsPng(
      flat.map((c) => ({ title: c.title, option: c.option }))
    )
    // 截图失败的图**不传** png_base64，由后端兜底渲染
    const charts = flat.map((c, i) => ({
      key: c.key,
      png_base64: shots[i]?.png_base64 || null,
    }))
    const resp = await reportsApi.exportVisualReport(taskId, {
      include_ai: true,
      charts,
    })
    const filename = await downloadBlobResponse(resp, 'visual_report.xlsx')
    ElMessage.success(`已导出可视化报表（${total} 张图）· ${filename || ''}`)
  } catch (e) {
    ElMessage.error('导出报表失败：' + (e?.message || e))
  } finally {
    exportingVisualReport.value = false
  }
}

onMounted(load)
</script>

<style scoped>
.page-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 18px;
  flex-wrap: wrap;
  gap: 12px;
}

/* ===== 报告正文「白底黑字 + 深青标题」风格 =====
 * 设计意图：让前端预览尽可能贴近 Word / PDF 导出后的实际观感，
 * 避免用户看到深色网页后下载了白底文档产生「版式不一致」的困惑。
 * 仅作用于 .report-body 内部,系统其它深色 UI 不受影响。
 */
.report-body {
  background: #ffffff;
  color: #1f2937; /* 主正文:近黑 */
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  padding: 36px 44px;
  box-shadow: 0 6px 20px rgba(0, 0, 0, 0.25);
  font-size: 14px;
  line-height: 1.85;
  /* 报告整体最大宽度,模拟 A4 阅读区,避免太长一行读起来累 */
  max-width: 920px;
  margin: 0 auto;
}
.report-body h1 {
  color: #0e7490;            /* 主标题:深青,与 docx/PDF 章节色对齐 */
  font-size: 24px;
  font-weight: 700;
  border-bottom: 2px solid #0e7490;
  padding-bottom: 10px;
  margin: 0 0 22px;
  text-align: center;
}
.report-body h2 {
  color: #0e7490;            /* 章节标题:深青 */
  font-size: 18px;
  font-weight: 600;
  margin: 28px 0 12px;
  padding-left: 10px;
  border-left: 4px solid #0e7490;
  background: #f0f9fb;        /* 浅青底,与正文形成层次 */
}
.report-body p {
  color: #1f2937;
  margin: 8px 0 12px;
  text-indent: 0;             /* 报告体不强加首行缩进,避免 LLM 已写紧凑段落时被二次排版撑开 */
}
.report-body .dim {
  color: #6b7280;
}

/* ========== 【row-diff-detail-expansion 新增】数据集对比·变更追踪·明细样式 ==========
   报告 sections["对比详情"] 是 dict 时,新增 / 删除 / 修改 三个分块的视觉规范。
   - 颜色基调与合同报告一致(深色卡片)但更紧凑,适合"大量数据行"场景;
   - 表格横向可滚动,避免列多撑爆报告容器;
   - 字段级 diff 用红 / 绿 标记变更前 / 后,直观可读。*/
.diff-detail {
  margin-top: 8px;
}
.diff-detail .diff-summary {
  margin: 0 0 12px 0;
  padding: 8px 12px;
  background: #f3f4f6;
  border-left: 3px solid #0891b2;
  border-radius: 4px;
  color: #1f2937;
}
.diff-detail .diff-block {
  margin-bottom: 16px;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  overflow: hidden;
  background: #ffffff;
}
.diff-detail .diff-block-head {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 12px;
  background: #f9fafb;
  border-bottom: 1px solid #e5e7eb;
}
.diff-detail .diff-tag {
  display: inline-block;
  padding: 2px 10px;
  font-size: 12px;
  border-radius: 10px;
  color: #fff;
  font-weight: 600;
}
.diff-detail .diff-tag-add {
  background: #10b981;
}
.diff-detail .diff-tag-remove {
  background: #ef4444;
}
.diff-detail .diff-tag-modify {
  background: #f59e0b;
}
.diff-detail .diff-count {
  color: #6b7280;
  font-size: 13px;
}
.diff-detail .diff-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
  display: block;
  overflow-x: auto;
  white-space: nowrap;
}
.diff-detail .diff-table th,
.diff-detail .diff-table td {
  border-bottom: 1px solid #f1f5f9;
  border-right: 1px solid #f1f5f9;
  padding: 6px 10px;
  text-align: left;
  vertical-align: top;
}
.diff-detail .diff-table th {
  background: #f3f4f6;
  color: #374151;
  font-weight: 600;
}
.diff-detail .diff-table tr:last-child td {
  border-bottom: none;
}
.diff-detail .diff-truncated {
  padding: 8px 12px;
  font-size: 12px;
  background: #f9fafb;
}
.diff-detail .diff-modify-card {
  padding: 10px 12px;
  border-top: 1px solid #f1f5f9;
}
.diff-detail .diff-modify-card:first-of-type {
  border-top: none;
}
.diff-detail .diff-modify-key {
  font-size: 13px;
  color: #1f2937;
  margin-bottom: 6px;
}
.diff-detail .diff-modify-fields {
  margin-top: 4px;
}
.diff-detail .diff-before {
  background: #fef2f2;
  color: #991b1b;
}
.diff-detail .diff-after {
  background: #ecfdf5;
  color: #065f46;
}

.summary {
  background: #f0f9fb;
  border: 1px solid #cffafe;
  border-left: 3px solid #0e7490;
  padding: 10px 14px;
  border-radius: 4px;
  margin-bottom: 18px;
  color: #1f2937;
}
.summary strong {
  color: #0e7490;
}

/* 【pdf-tools-history-persistence 新增】PDF 产物下载卡片:与 summary 同款淡青风 */
.pdf-artifact-card {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  background: linear-gradient(90deg, rgba(34, 211, 238, 0.10), rgba(34, 211, 238, 0.04));
  border: 1px solid rgba(34, 211, 238, 0.35);
  border-left: 3px solid #0e7490;
  padding: 12px 16px;
  border-radius: 4px;
  margin-bottom: 18px;
}
.pdf-artifact-card .pdf-artifact-info {
  flex: 1;
  min-width: 0;
}
.pdf-artifact-card .pdf-artifact-title {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 14px;
  font-weight: 600;
  color: var(--accent-cyan, #0e7490);
  margin-bottom: 4px;
}
.pdf-artifact-card .pdf-artifact-meta {
  font-size: 12px;
  color: #64748b;
}
.pdf-artifact-card .pdf-artifact-meta span {
  margin-right: 2px;
}

.charts {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 16px;
  margin-bottom: 12px;
}
/* 章节块:正文 + 紧贴的图表组,版式紧凑、阅读路径连贯 */
.section-block {
  margin-bottom: 18px;
}
.section-block .section-charts {
  margin-top: 10px;
}
/* 【按章节挂载 · 补章】章节无图占位提示:浅灰底 + 左边细线,与正文段和谐相处 */
.section-no-chart {
  margin-top: 8px;
  padding: 6px 10px;
  background: #f9fafb;
  border-left: 2px solid #cbd5e1;
  border-radius: 4px;
  font-size: 12px;
  color: #64748b;
}
/* ========== 【2026-10 新增】章节明细数据表样式 ==========
   目的:让"文字 + 图表 + 数据表"三者共存于同一章节块内,
   图表看趋势、表格看构成,管理层可在同一屏内交叉核对数字。 */
.detail-table-block {
  margin-top: 12px;
  margin-bottom: 6px;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  overflow: hidden;
  background: #ffffff;
}
.detail-table-title {
  padding: 8px 12px;
  background: #f0f9fb;              /* 浅青底,与章节标题层级一致 */
  color: #0e7490;
  font-weight: 600;
  font-size: 13px;
  border-bottom: 1px solid #e5e7eb;
}
.detail-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
  display: block;
  overflow-x: auto;                 /* 列多时横向滚动,不撑破报告容器 */
  white-space: nowrap;
}
.detail-table th,
.detail-table td {
  padding: 6px 12px;
  border-bottom: 1px solid #f1f5f9;
  text-align: left;
}
.detail-table th {
  background: #f9fafb;
  color: #374151;
  font-weight: 600;
}
.detail-table tbody tr:last-child td {
  border-bottom: none;
}
.detail-table tbody tr:nth-child(even) td {
  background: #fcfdfe;              /* 极浅斑马纹,长表格更易读 */
}
/* 数字列右对齐,符合财务/数据表的阅读习惯 */
.detail-table td.dt-num {
  text-align: right;
  font-variant-numeric: tabular-nums;
}
.detail-table-note {
  padding: 6px 12px;
  background: #f9fafb;
  border-top: 1px solid #f1f5f9;
  color: #6b7280;
  font-size: 12px;
  line-height: 1.6;
}

/* 【2026-10 新增】章节段落样式。
   需求:「章节要分段,不要全部都在一段上」。
   段落之间用 margin-bottom + 无首行缩进(报告体是现代排版,
   中文公文式的首行缩进 2 字符反而让密集段落更难读)。 */
.report-paragraph {
  margin: 0 0 10px;
  line-height: 1.9;
  color: #1f2937;
  text-align: justify;
}
.report-paragraph:last-child {
  margin-bottom: 0;
}

/* 【report-sections-structured 新增】通用段落保留换行符,
   避免历史报告(如老合同报告的纯文本 sections)被 HTML 合并成单行 */
.report-section-text {
  white-space: pre-wrap;
  word-break: break-word;
  margin: 0;
}
/* 通用 list 兜底样式(非合同模式也兼容) */
.report-generic-list {
  list-style: disc inside;
  margin: 0;
  padding: 0;
}
.report-generic-list li {
  padding: 4px 0;
  line-height: 1.7;
  font-size: 14px;
  /* 报告正文卡片是白底,li 颜色必须强制黑色,不能继承项目深色主题变量 */
  color: #111111;
}

/* ========== 【report-sections-structured 新增】合同报告专属样式 ==========
   - 条款差异表:每条是一个卡片,左/右方用 → 串起来,高优变更加底色
   - 利弊分析:按 side 分为 左优/右优/各有 三种状态,左边色条对应
   - 谈判建议:加 ⭐ 图标,带边框卡片化 */
.contract-clause-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
  margin-top: 6px;
}
.contract-clause-card {
  background: rgba(10, 20, 48, 0.45);
  border: 1px solid var(--border-cyan-soft);
  border-left: 3px solid #94a3b8;
  border-radius: 6px;
  padding: 10px 14px;
  transition: background 0.15s ease;
}
.contract-clause-card.clause-high {
  background: rgba(248, 113, 113, 0.10);
  border-left-color: #f87171;
}
.clause-head {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 8px;
  flex-wrap: wrap;
}
.clause-label {
  font-weight: 600;
  font-size: 14px;
  color: var(--text-primary);
}
.clause-row {
  display: grid;
  grid-template-columns: 1fr 24px 1fr;
  align-items: center;
  gap: 8px;
  margin-bottom: 8px;
}
.clause-side-label {
  font-size: 11px;
  color: #94a3b8;
  margin-bottom: 2px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}
.clause-side-value {
  font-size: 13px;
  line-height: 1.6;
  color: var(--text-primary);
  word-break: break-word;
  white-space: pre-wrap;
}
.clause-arrow {
  text-align: center;
  color: var(--accent-cyan, #22d3ee);
  font-weight: 700;
  font-size: 16px;
}
.clause-note {
  font-size: 13px;
  line-height: 1.7;
  color: #cbd5e1;
  padding: 6px 10px;
  background: rgba(0, 0, 0, 0.2);
  border-radius: 4px;
  margin-top: 2px;
}

.contract-pc-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
  margin-top: 6px;
}
.contract-pc-card {
  background: rgba(10, 20, 48, 0.45);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 6px;
  border-left-width: 3px;
  padding: 10px 14px;
}
.contract-pc-card.pc-left  { border-left-color: #4ade80; }
.contract-pc-card.pc-right { border-left-color: #fbbf24; }
.contract-pc-card.pc-both  { border-left-color: #94a3b8; }
.pc-head {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 6px;
  flex-wrap: wrap;
}
.pc-title {
  font-size: 14px;
  font-weight: 600;
  color: var(--text-primary);
}
.pc-detail {
  font-size: 13px;
  line-height: 1.7;
  color: #cbd5e1;
  padding-left: 4px;
}

.contract-tips {
  list-style: none;
  margin: 0;
  padding: 0;
}
.contract-tips li {
  padding: 8px 12px;
  margin-bottom: 6px;
  background: rgba(10, 20, 48, 0.4);
  border: 1px solid rgba(34, 211, 238, 0.18);
  border-radius: 6px;
  font-size: 13px;
  line-height: 1.7;
  color: var(--text-primary);
  display: flex;
  align-items: flex-start;
}
.chart-box {
  background: #ffffff;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  padding: 8px;
}
.chart {
  height: 280px;
  width: 100%;
}
.metric-table {
  width: 100%;
  border-collapse: collapse;
  color: #1f2937;
}
.metric-table th,
.metric-table td {
  padding: 8px 12px;
  border: 1px solid #e5e7eb;
  text-align: left;
}
.metric-table th {
  background: #0e7490;        /* 表头深青底 */
  color: #ffffff;
  font-weight: 600;
}
.metric-table tbody tr:nth-child(even) td {
  background: #f9fafb;        /* 偶数行浅灰,与 docx Light Grid 一致 */
}
.anno {
  display: flex;
  gap: 12px;
  padding: 10px 14px;
  border-radius: 4px;
  margin: 6px 0;
  background: #f9fafb;
  border: 1px solid #e5e7eb;
  border-left: 3px solid #6b7280;
}
.anno.warn   { border-left-color: #fbbf24; }
.anno.danger { border-left-color: #f87171; }
.anno.info   { border-left-color: #0e7490; }
.anno-metric {
  color: #0e7490;
  font-weight: 600;
  min-width: 120px;
}
.anno-comment {
  color: #1f2937;
}
/* 兜底指标区:无数值列时,缺失率 / 文本频次也能呈现 */
.metric-fallback {
  display: flex;
  flex-direction: column;
  gap: 18px;
  margin-bottom: 12px;
}
.sub-h {
  font-size: 14px;
  color: #0e7490;
  margin: 4px 0 8px;
  font-weight: 600;
}

/* 「数据图表」孤儿图兜底区 h2 与上面 h2 视觉对齐 */
.report-body h2 + .charts {
  margin-top: 10px;
}

/* =====【report-pptx-themes 新增】PPT 弹窗样式 ===== */
.pptx-dialog-body {
  display: flex;
  gap: 18px;
  min-height: 480px;
  /* dialog body 本身白底,不被项目深色主题压住 */
  background: #ffffff;
}
.pptx-themes-col {
  width: 200px;
  flex-shrink: 0;
}
.pptx-themes-label {
  font-size: 13px;
  color: #1f2937;
  margin-bottom: 10px;
}
/* dialog 内的 title / 主题描述 / 列表等文字强制深色,白底上能读 */
.pptx-dialog-body :deep(.pptx-themes-label),
.pptx-dialog-body :deep(.pptx-theme-name),
.pptx-dialog-body :deep(.pptx-theme-desc),
.pptx-dialog-body :deep(.pptx-empty) {
  color: #1f2937;
}
.pptx-themes-radio {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.pptx-themes-radio :deep(.el-radio-button) {
  margin-right: 0;
  width: 100%;
}
.pptx-themes-radio :deep(.el-radio-button__inner) {
  width: 100%;
  padding: 10px 14px;
  border-radius: 6px !important;
  background: #f9fafb;
  border-color: #d1d5db !important;
  color: #1f2937 !important;
}
.pptx-themes-radio :deep(.el-radio-button.is-active .el-radio-button__inner) {
  /* 当前选中的主题 radio:用项目主色高亮,但缩略图里仍按 theme 自身染色 */
  background: var(--accent-cyan) !important;
  color: #ffffff !important;
  border-color: var(--accent-cyan) !important;
}
.pptx-theme-name {
  font-weight: 600;
  color: #1f2937;
  margin-bottom: 4px;
}
.pptx-theme-desc {
  font-size: 11px;
  color: #6b7280;
  line-height: 1.4;
}
/* Element Plus 的 .el-radio-button__inner 默认 white-space: nowrap,
   左栏只有 200px 宽,长描述文字会横向溢出。这里改成正常换行,
   让「版本」组的描述能完整折行显示。 */
.pptx-source-radio :deep(.el-radio-button__inner) {
  white-space: normal;
  text-align: left;
}
.pptx-slides-col {
  flex: 1;
  min-width: 0;
}
.pptx-slides-panel {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
  max-height: 60vh;
  overflow-y: auto;
  padding: 4px;
}
/* 窄屏(如 1366px 笔记本、侧边栏展开后只剩 1000px 左右)时
   3 列会把每张卡片压得过窄,标题和 bullet 都挤成两三行。
   这里在 1280px 断点下降到 2 列,保证卡片里的内容还能读。 */
@media (max-width: 1280px) {
  .pptx-slides-panel {
    grid-template-columns: repeat(2, 1fr);
  }
}
.pptx-slide-card {
  /* 白底:让 4 套主题色都能在预览里被看清;主题切换就是颜色切换,
  * 而不是被项目深蓝主题一直压住看不出来 */
  background: #ffffff;
  border: 1px solid #d1d5db;
  border-radius: 8px;
  padding: 0 0 10px 0;
  min-height: 130px;
  display: flex;
  flex-direction: column;
  gap: 6px;
  overflow: hidden;
}
/* 顶条 accent:每张卡顶部一条 4px 高的色条,color 由 inline style 注入 */
.pptx-slide-topbar {
  height: 4px;
  width: 100%;
}
/* 标题页/尾页 模拟 PPT 实际大色块效果 */
.pptx-slide-hero {
  flex: 1;
  margin: 6px 10px 0;
  border-radius: 6px;
  padding: 22px 14px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 8px;
  text-align: center;
}
.pptx-slide-hero-title {
  font-size: 16px;
  font-weight: 700;
}
.pptx-slide-hero-sub {
  font-size: 11px;
  opacity: 0.85;
}
.pptx-slide-card-no {
  font-size: 11px;
  color: #6b7280;
  padding: 8px 10px 0;
  display: flex;
  align-items: center;
  gap: 6px;
}
.pptx-slide-kind-chip {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 10px;
  font-size: 10px;
  font-weight: 600;
}
.pptx-slide-page-no {
  font-size: 10px;
  /* 颜色由模板内联 style 传 themeCss.muted 覆盖:
     写死灰色在暗夜主题下只有 3.69:1,小字号读不清。 */
}
.pptx-slide-title {
  font-weight: 600;
  color: #1f2937;
  font-size: 13px;
  line-height: 1.4;
  padding: 0 10px;
}
.pptx-slide-img {
  width: calc(100% - 20px);
  margin: 0 10px;
  border-radius: 4px;
  max-height: 120px;
  object-fit: contain;
  background: #fff;
}
.pptx-slide-img-fallback {
  font-size: 11px;
  color: #6b7280;
  background: #f3f4f6;
  padding: 6px 10px;
  border-radius: 4px;
  margin: 0 10px;
}
.pptx-slide-pre {
  font-size: 10px;
  max-height: 60px;
  overflow: hidden;
  white-space: pre-wrap;
  word-break: break-all;
  margin: 4px 0 0;
  color: #1f2937;
}
.pptx-slide-body {
  font-size: 12px;
  color: #374151;
  line-height: 1.6;
  max-height: 80px;
  overflow: hidden;
  word-break: break-all;
  padding: 0 10px;
}
.pptx-empty {
  grid-column: span 3;
  text-align: center;
  padding: 40px;
}

/* ========== 【report-failed-state 新增】任务失败提示样式 ==========
 * 失败时 el-alert 内放两行内容:
 *  - 第一行:后端给的 error_message(具体失败原因,鉴权/超时/解析都如实展示)
 *  - 第二行:重新分析 / 返回工作台 两个动作按钮
 * 与现有 dashboard 风格的 rgba 深青/灰底保持一致,白底报告卡片里也清晰可见。 */
.failed-alert-body {
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding-top: 4px;
}
.failed-error-text {
  color: #1f2937;
  font-size: 13px;
  line-height: 1.6;
  /* 错误文案可能较长,自动换行 */
  white-space: pre-wrap;
  word-break: break-word;
}
.failed-actions {
  display: flex;
  gap: 8px;
}

/* ========== 【AI 演示稿 新增】版本切换 / 状态区 / 演示稿缩略图样式 ==========
 * 风格与上面 .pptx-* 系列保持一致:白底卡片 + 灰边框 + 深色文字,
 * 这样在 dialog 白底里 4 套主题都能看清,不被项目深色主题压住。 */

/* 「推荐」小标签:贴在「AI 演示稿」标题右侧 */
.pptx-recommend-tag {
  display: inline-block;
  margin-left: 4px;
  padding: 1px 6px;
  border-radius: 8px;
  font-size: 10px;
  font-weight: 600;
  background: #ecfeff;
  color: #0e7490;
  border: 1px solid #a5f3fc;
  vertical-align: middle;
}
/* 选中态下标签要反色,否则白字底上的浅色标签看不清 */
.pptx-themes-radio :deep(.el-radio-button.is-active) .pptx-recommend-tag {
  background: rgba(255, 255, 255, 0.25);
  color: #ffffff;
  border-color: rgba(255, 255, 255, 0.6);
}
/* 选中态下描述文字也反色:否则深灰文字压在青色底上对比度不足 */
.pptx-source-radio :deep(.el-radio-button.is-active) .pptx-theme-desc {
  color: rgba(255, 255, 255, 0.85);
}

/* 演示稿状态区:一行三段(状态文案 + 按钮 + 提示),随状态换底色 */
.pptx-deck-status {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 10px;
  padding: 10px 12px;
  margin-bottom: 12px;
  border-radius: 8px;
  border: 1px solid #d1d5db;
  background: #f9fafb;
  font-size: 13px;
  color: #1f2937;
}
/* 未生成:中性灰,引导用户点生成 */
.pptx-deck-status[data-status='none'] {
  background: #f9fafb;
  border-color: #d1d5db;
}
/* 生成中:青色高亮,提示正在跑 LLM */
.pptx-deck-status[data-status='generating'] {
  background: #ecfeff;
  border-color: #a5f3fc;
}
/* 已生成:绿色,表示可以下载了 */
.pptx-deck-status[data-status='ready'] {
  background: #f0fdf4;
  border-color: #bbf7d0;
}
/* 失败:红色,展示失败原因 */
.pptx-deck-status[data-status='failed'] {
  background: #fef2f2;
  border-color: #fecaca;
}
.pptx-deck-status-text {
  font-weight: 600;
}
.pptx-deck-status-error {
  color: #b91c1c;
  /* 失败原因可能很长,允许换行并占满一行 */
  flex: 1 1 100%;
  word-break: break-word;
  white-space: pre-wrap;
}
.pptx-deck-status-hint {
  font-size: 12px;
  color: #6b7280;
}

/* 演示稿内容页的 bullet 列表:一条一行 + 圆点前缀 */
.pptx-slide-bullets {
  margin: 2px 10px 0;
  padding-left: 16px;
  font-size: 12px;
  line-height: 1.6;
  /* 缩略图高度有限,超出的部分裁掉。
     用渐隐遮罩收尾,而不是硬切 —— 硬切会从半行字中间断开,
     读者以为是内容本身断了,加了遮罩才知道"下面还有,被折叠了"。 */
  max-height: 96px;
  overflow: hidden;
  word-break: break-word;
  -webkit-mask-image: linear-gradient(to bottom, #000 78%, transparent 100%);
  mask-image: linear-gradient(to bottom, #000 78%, transparent 100%);
}
.pptx-slide-bullets li {
  list-style: disc;
  margin-bottom: 2px;
}

/* 演讲备注:主题 muted 色的小字,只做提示,不抢标题的视觉重心 */
.pptx-slide-notes {
  margin: 4px 10px 0;
  font-size: 11px;
  line-height: 1.5;
  max-height: 42px;
  overflow: hidden;
  word-break: break-word;
  -webkit-mask-image: linear-gradient(to bottom, #000 74%, transparent 100%);
  mask-image: linear-gradient(to bottom, #000 74%, transparent 100%);
}
</style>
