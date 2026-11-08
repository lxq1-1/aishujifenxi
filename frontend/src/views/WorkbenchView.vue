<!--
  工作台（核心页）：
  - 左：当前数据集 + 模板选择 + 6 个辅助工具按钮（含数据清洗）+ 开始 AI 分析
  - 中：3 个指标卡片 + 表格预览 + 解析时间线
  - 右：AI 分析结果（趋势图、占比、四段式文字）
  - 弹窗/抽屉：多表合并、数据校验、数据清洗、图表编辑、数据对话、数据集对比、PDF 处理
  - 清洗规则的 4 项开关已收敛到「数据清洗」弹窗(见 showClean)
-->
<template>
  <div class="page wb-page">
    <div class="wb-grid">
      <!-- 左侧 -->
      <aside class="wb-side">
        <div class="wb-section-title">当前数据集</div>
        <el-card v-if="wbStore.draft.datasetId" class="ds-card">
          <div class="cyan-title">{{ wbStore.draft.datasetName }}</div>
          <div class="dim small">ID: {{ wbStore.draft.datasetId }}</div>
          <!-- 【dataset-multisheet-merge 新增】工作表标签 + Tooltip 明细;
               仅当 ≥2 张 sheet 时显示;老数据未拉 sheets 时也无影响 -->
          <el-tooltip
            v-if="wbStore.draft.datasetSheets.length > 1"
            placement="top"
            :content="fileListMatrixTooltip"
          >
            <el-tag type="info" size="small" effect="plain" class="ds-sheet-tag">
              📑 包含 {{ wbStore.draft.datasetSheets.length }} 个工作表
            </el-tag>
          </el-tooltip>
          <!-- 【workbench-per-sheet-preview 新增】当前查看的 sheet 提示:
              单 sheet 文件无需显示(始终只有 1 张);
              多 sheet 文件显示当前选中的 sheet 名,默认是**第一张 sheet**。
              让用户随时知道当前看到的缺失值 / 列类型是来自哪张表。 -->
          <el-tag
            v-if="wbStore.draft.datasetSheets.length > 1 && wbStore.draft.activeSheet"
            size="small"
            effect="plain"
            type="success"
            class="ds-sheet-tag"
          >
            👁 当前查看:{{ wbStore.draft.activeSheet }}
          </el-tag>
          <el-button link size="small" @click="$router.push('/data-datasets')">更换数据集</el-button>
        </el-card>
        <el-button v-else class="wb-side-action" @click="$router.push('/data-datasets')">
          请先选择数据集
        </el-button>

        <div class="wb-section-title">报告模板</div>
        <el-card v-if="wbStore.draft.templateId" class="ds-card">
          <div class="cyan-title">{{ wbStore.draft.templateName }}</div>
          <el-button link size="small" @click="$router.push('/templates')">更换模板</el-button>
        </el-card>
        <el-button v-else class="wb-side-action" @click="$router.push('/templates')">
          请先选择模板
        </el-button>

        <div class="wb-section-title">辅助工具</div>
        <!-- 辅助工具顺序：多表合并 → 数据校验 → 数据清洗 → 图表编辑 → 数据对话 → 数据集对比 → PDF 处理 -->
        <el-button class="wb-side-action" @click="showMerge = true">
          <el-icon><Operation /></el-icon> 多表合并
        </el-button>
        <el-button class="wb-side-action" @click="showValidate = true">
          <el-icon><CircleCheck /></el-icon> 数据校验
        </el-button>
        <el-button class="wb-side-action" @click="showClean = true">
          <el-icon><Brush /></el-icon> 数据清洗
        </el-button>
        <el-button class="wb-side-action" @click="showChart = true">
          <el-icon><DataLine /></el-icon> 图表编辑
        </el-button>
        <el-button class="wb-side-action" @click="showNl = true">
          <el-icon><ChatDotRound /></el-icon> 数据对话
        </el-button>
        <el-button class="wb-side-action" @click="showCompare = true">
          <el-icon><Switch /></el-icon> 数据集对比
        </el-button>
        <!-- 【workbench-ai-pdf-tools 新增】第 7 个辅助工具:AI PDF 处理 -->
        <el-button class="wb-side-action" @click="showPdfTools = true">
          <el-icon><Document /></el-icon> PDF 处理
        </el-button>

        <el-button
          type="primary"
          size="large"
          class="wb-side-action wb-side-primary"
          :loading="analyzing"
          :disabled="!wbStore.draft.datasetId || !wbStore.draft.templateId"
          @click="onAnalyzeClick"
        >
          开始 AI 分析
        </el-button>
      </aside>

      <!-- 中部 -->
      <section class="wb-mid">
        <div class="metric-row">
          <div class="metric-card">
            <div class="metric-icon">📊</div>
            <div>
              <div class="metric-label">行数</div>
              <div class="metric-value">{{ metrics.row_count || 0 }}</div>
            </div>
          </div>
          <div class="metric-card">
            <div class="metric-icon">⚠️</div>
            <div>
              <div class="metric-label">缺失值</div>
              <div class="metric-value">{{ metrics.missing_total || 0 }} 处</div>
            </div>
          </div>
          <div class="metric-card">
            <div class="metric-icon">🗂️</div>
            <div>
              <div class="metric-label">重复值</div>
              <div class="metric-value">{{ metrics.duplicate_rows || 0 }} 条</div>
            </div>
          </div>
        </div>

        <h3 class="dim" style="margin: 16px 0 8px; font-size: 13px">数据预览（前 8 行）</h3>

        <!--
          【workbench-per-sheet-preview 新增】sheet 切换 tab:
          - 多 sheet 文件(≥2)且 data.sheets 列表非空时显示;
          - 默认选中**第一张 sheet**(由 loadPreview 完成后同步),
            不再展示"全部合并"——它的 11701 个跨表 NaN 缺失值严重误导分析。
          - 切换 tab → onPreviewSheetChange → store.setActiveSheet + 重拉预览。
          - 每个 tab 上用 el-tooltip 显示该 sheet 的行/列数,避免用户瞎猜。
        -->
        <el-tabs
          v-if="wbStore.draft.datasetSheets.length > 1"
          v-model="previewSheetTab"
          type="card"
          class="wb-preview-sheet-tabs"
          @tab-change="onPreviewSheetChange"
        >
          <el-tab-pane
            v-for="s in wbStore.draft.datasetSheets"
            :key="s.name"
            :name="s.name"
            :label="s.name"
          >
            <template #label>
              <el-tooltip
                placement="top"
                :content="`${s.name}: ${s.rows} 行 · ${s.cols} 列`"
              >
                <span>{{ s.name }}</span>
              </el-tooltip>
            </template>
          </el-tab-pane>
        </el-tabs>

        <el-table :data="previewRows" border size="small" v-loading="previewLoading">
          <el-table-column
            v-for="col in previewColumns"
            :key="col"
            :prop="col"
            :label="col"
            min-width="110"
          />
        </el-table>

        <div class="status-bar">
          <div class="timeline-row">
            <div class="timeline-dot done"></div>
            <span>解析完成</span>
          </div>
          <div class="timeline-row">
            <div class="timeline-dot" :class="{ done: analyzeStep >= 1, active: analyzeStep === 1 }"></div>
            <span>清洗完成</span>
          </div>
          <div class="timeline-row">
            <div class="timeline-dot" :class="{ done: analyzeStep >= 2, active: analyzeStep === 2 }"></div>
            <span>指标计算完成</span>
          </div>
          <div class="timeline-row">
            <div class="timeline-dot" :class="{ done: analyzeStep >= 3, active: analyzeStep === 3 }"></div>
            <span>AI 分析完成</span>
          </div>
        </div>
      </section>

      <!-- 右侧 AI 结果 -->
      <aside class="wb-result">
        <div v-if="!analyzeResult" class="dim" style="text-align: center; padding: 40px 0">
          请先选择数据集与模板<br />点击「开始 AI 分析」生成报告
        </div>
        <template v-else>
          <!-- 【workbench-failed-state 新增】任务失败时,右侧结果区不再渲染
               "29 行 × 24 列"式的本地兜底图,而是给一张错误卡片 + 两个动作按钮。
               旧实现下:LLM 失败 → 后端写入 report_content → 前端照样渲染各章节
               的占位段,用户看不到失败原因;现在统一把失败状态顶上来。 -->
          <div v-if="analyzeResult.error" class="wb-result-failed">
            <div class="wb-result-failed-title">
              ❌ 任务 #{{ analyzeResult.task_id }} 生成失败
            </div>
            <div class="wb-result-failed-body">
              {{ analyzeResult.error }}
            </div>
            <div class="wb-result-failed-actions">
              <el-button
                type="primary"
                size="small"
                :loading="analyzing"
                @click="onAnalyzeClick"
              >
                重新尝试
              </el-button>
              <el-button
                v-if="analyzeResult.task_id"
                size="small"
                @click="$router.push('/reports/' + analyzeResult.task_id)"
              >
                查看失败详情页
              </el-button>
            </div>
          </div>

          <div class="result-section">
            <h3 class="cyan-title">📈 趋势</h3>
            <v-chart v-if="trendOption" class="mini-chart" :option="trendOption" autoresize />
            <div v-else class="dim small">无趋势数据</div>
          </div>

          <div class="result-section">
            <h3 class="cyan-title">🥧 占比</h3>
            <v-chart v-if="pieOption" class="mini-chart" :option="pieOption" autoresize />
            <div v-else class="dim small">无占比数据</div>
          </div>

          <div class="result-section" v-if="analyzeResult.sections">
            <div v-for="(text, key) in analyzeResult.sections" :key="key" class="ai-section">
              <div class="cyan-title">{{ sectionLabel(key) }}</div>
              <div class="ai-text">{{ text }}</div>
            </div>
          </div>

          <el-button
            v-if="analyzeResult.task_id"
            type="primary"
            style="width: 100%; margin-top: 12px"
            @click="$router.push(`/reports/${analyzeResult.task_id}`)"
          >
            查看完整报告
          </el-button>
        </template>
      </aside>
    </div>

    <!-- 多表合并 -->
    <el-dialog v-model="showMerge" title="多表合并" width="640">
      <el-form label-width="100">
        <el-form-item label="新数据集名">
          <el-input v-model="mergeForm.name" placeholder="合并后的数据集名" />
        </el-form-item>
        <el-form-item label="合并方式">
          <!--
            两种模式:
              - flat        :把所选数据集纵向堆叠为 1 张大表(默认,旧行为)
              - multi_sheet :把每个数据集作为 1 个 sheet 合并到同一个 Excel
          -->
          <el-radio-group v-model="mergeForm.mode">
            <el-radio-button value="flat">合并为一张大表</el-radio-button>
            <el-radio-button value="multi_sheet">合并到 Excel 多工作表</el-radio-button>
          </el-radio-group>
          <div class="dim small" style="margin-top: 6px">
            {{
              mergeForm.mode === 'multi_sheet'
                ? '每个所选数据集对应输出 Excel 的 1 个 sheet,sheet 名=数据集名。'
                : '把所有所选数据集的数据纵向拼接为一张表(原行为)。'
            }}
          </div>
        </el-form-item>
        <el-form-item label="选择数据集">
          <el-select v-model="mergeForm.dataset_ids" multiple style="width: 100%">
            <el-option v-for="d in datasetOptions" :key="d.id" :label="d.name" :value="d.id" />
          </el-select>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showMerge = false">取消</el-button>
        <el-button type="primary" :loading="toolLoading" @click="onMerge">开始合并</el-button>
      </template>
    </el-dialog>

    <!-- 数据校验（带 AI 自然语言解读） -->
    <el-dialog v-model="showValidate" title="数据校验" width="720" top="6vh">
      <el-checkbox v-model="saveAsValidateTask" style="margin-bottom: 8px">
        💾 保存为历史任务
      </el-checkbox>
      <el-button type="primary" :loading="toolLoading" @click="onValidate">
        运行校验
      </el-button>
      <!-- 【tools-results-history-persistence 新增】已保存为历史任务提示卡 -->
      <div v-if="latestValidateTaskId" class="saved-task-card">
        📦 已保存为历史任务 #{{ latestValidateTaskId }} · 数据校验
        <el-button type="primary" size="small" @click="$router.push(`/reports/${latestValidateTaskId}`)">
          查看完整报告
        </el-button>
      </div>
      <el-button v-if="validateResult" link @click="showRawJson = !showRawJson">
        {{ showRawJson ? '收起原始数据' : '查看原始数据' }}
      </el-button>

      <!-- AI 解读区 -->
      <div v-if="validateResult?.ai_interpret" class="ai-interpret">
        <!-- 顶部状态卡:等级 + 一句话总览 -->
        <div class="ai-summary" :class="'level-' + validateResult.ai_interpret.level">
          <span class="level-badge">{{ levelLabel(validateResult.ai_interpret.level) }}</span>
          <span class="summary-text">{{ validateResult.ai_interpret.summary }}</span>
          <el-tooltip
            v-if="validateResult.ai_interpret.fallback"
            placement="top"
            :content="validateResult.ai_interpret.error || '未连接或调用失败'"
          >
            <el-tag type="warning" size="small" effect="plain">离线解读(LLM 未连接)</el-tag>
          </el-tooltip>
        </div>

        <!-- 亮点 -->
        <div v-if="validateResult.ai_interpret.highlights?.length" class="ai-block highlights">
          <div class="ai-h">✨ 数据亮点</div>
          <ul>
            <li v-for="(t, i) in validateResult.ai_interpret.highlights" :key="'h' + i">{{ t }}</li>
          </ul>
        </div>

        <!-- 问题清单 -->
        <div v-if="validateResult.ai_interpret.issues?.length" class="ai-block issues">
          <div class="ai-h">⚠️ 风险问题</div>
          <ul>
            <li v-for="(iss, i) in validateResult.ai_interpret.issues" :key="'i' + i">
              <span class="sev-badge" :class="'sev-' + iss.severity">{{ sevLabel(iss.severity) }}</span>
              <strong>{{ iss.field }}:</strong>
              <span>{{ iss.snippet }}</span>
            </li>
          </ul>
        </div>

        <!-- 建议 -->
        <div v-if="validateResult.ai_interpret.suggestions?.length" class="ai-block suggestions">
          <div class="ai-h">💡 处理建议</div>
          <ul>
            <li v-for="(t, i) in validateResult.ai_interpret.suggestions" :key="'s' + i">{{ t }}</li>
          </ul>
        </div>
      </div>

      <!-- 原始数据表格(折叠) -->
      <div v-if="validateResult && showRawJson" class="raw-preview">
        <div class="raw-preview-title">
          📋 数据集前 10 行预览(共 {{ rawPreviewTotal }} 行 · 共 {{ rawPreviewColumns.length }} 列)
        </div>
        <el-table
          :data="rawPreviewRows"
          border
          size="small"
          max-height="360"
          stripe
        >
          <el-table-column
            v-for="col in rawPreviewColumns"
            :key="col"
            :prop="col"
            :label="col"
            min-width="110"
          />
        </el-table>
      </div>
    </el-dialog>

    <!-- 数据对话:多轮气泡流(用户靠右主色底 / AI 靠左浅底,支持结构化渲染与追问 chips) -->
    <el-dialog v-model="showNl" title="数据对话" width="900" top="5vh" @closed="onNlDialogClosed">
      <!-- 消息列表:固定高 52vh 超出滚动,新消息入列后由 scrollNlChatBottom 自动滚到底 -->
      <div ref="nlChatListRef" class="nl-chat-list">
        <!-- 空状态:提示文案 + 示例问题(点击即发送) -->
        <div v-if="chatMessages.length === 0" class="nl-chat-empty">
          <p class="nl-chat-empty-title">💬 基于当前数据集向 AI 提问，支持多轮追问</p>
          <div class="nl-chat-examples">
            <span class="nl-chat-examples-label">示例问题：</span>
            <el-tag
              v-for="q in nlExampleQuestions"
              :key="q"
              type="info"
              effect="plain"
              class="nl-example-tag"
              @click="onNlSend(q)"
            >{{ q }}</el-tag>
          </div>
        </div>
        <!-- 气泡流:按角色左右分布 -->
        <div
          v-for="(msg, mi) in chatMessages"
          :key="mi"
          class="nl-chat-row"
          :class="msg.role === 'user' ? 'nl-row-user' : 'nl-row-ai'"
        >
          <div class="nl-bubble" :class="msg.role === 'user' ? 'nl-bubble-user' : 'nl-bubble-ai'">
            <!-- 用户气泡:纯问题文本 -->
            <template v-if="msg.role === 'user'">{{ msg.content }}</template>
            <!-- AI 气泡 -->
            <template v-else>
              <!-- 结构化渲染:LLM 正常返回 structured 时 -->
              <template v-if="msg.structured">
                <!-- 结论区:高亮卡(浅色背景 + 左边框强调条) -->
                <div class="nl-answer-card">{{ msg.structured.answer }}</div>
                <!-- 分析区:逐条圆点列表 -->
                <div v-if="msg.structured.analysis?.length" class="nl-section">
                  <div class="nl-section-title">分析</div>
                  <ul class="nl-points-list">
                    <li v-for="(a, ai) in msg.structured.analysis" :key="ai">{{ a }}</li>
                  </ul>
                </div>
                <!-- 支撑表:二维 rows 按列下标取值(与旧表格写法一致),最多显示前 20 行 -->
                <div v-if="msg.structured.data_table?.columns?.length" class="nl-section">
                  <div class="nl-section-title">支撑数据</div>
                  <el-table :data="nlTableRows(msg.structured.data_table)" size="small" border class="nl-data-table">
                    <el-table-column
                      v-for="(colName, ci) in msg.structured.data_table.columns"
                      :key="ci"
                      :label="colName"
                    >
                      <template #default="{ row }">
                        <span>{{ Array.isArray(row) ? row[ci] : row[colName] }}</span>
                      </template>
                    </el-table-column>
                  </el-table>
                </div>
                <!-- 建议区 -->
                <div v-if="msg.structured.suggestions?.length" class="nl-section">
                  <div class="nl-section-title">建议</div>
                  <ul class="nl-points-list">
                    <li v-for="(s, si) in msg.structured.suggestions" :key="si">{{ s }}</li>
                  </ul>
                </div>
                <!-- 追问 chips:点击即把该文本作为新问题直接发送 -->
                <div v-if="msg.structured.followups?.length" class="nl-followups">
                  <el-tag
                    v-for="(f, fi) in msg.structured.followups"
                    :key="fi"
                    type="info"
                    effect="plain"
                    class="nl-followup-tag"
                    @click="onNlSend(f)"
                  >{{ f }}</el-tag>
                </div>
              </template>
              <!-- 降级渲染:structured 为 null(LLM 失败/旧响应) → explanation 文本 + rows/columns 旧表格 -->
              <template v-else>
                <p class="nl-fallback-text">{{ msg.content }}</p>
                <el-table v-if="msg.rows && msg.rows.length" :data="msg.rows" size="small" border class="nl-data-table">
                  <el-table-column v-for="(colName, idx) in msg.columns || []" :key="idx" :label="colName">
                    <template #default="{ row }">
                      <span>{{ Array.isArray(row) ? row[idx] : row[colName] }}</span>
                    </template>
                  </el-table-column>
                </el-table>
              </template>
            </template>
          </div>
        </div>
      </div>
      <!-- 底部固定输入区:Enter 发送 / Shift+Enter 换行 -->
      <div class="nl-chat-input">
        <el-input
          v-model="nlForm.question"
          type="textarea"
          :rows="2"
          :disabled="toolLoading"
          placeholder="输入问题，Enter 发送，Shift+Enter 换行"
          @keydown.enter.exact.prevent="onNlEnterKey"
        />
        <div class="nl-chat-input-foot">
          <!-- 【tools-results-history-persistence 保留】保存为历史任务勾选 -->
          <el-checkbox v-model="saveAsNlTask">💾 保存为历史任务</el-checkbox>
          <el-button type="primary" :loading="toolLoading" :disabled="!nlForm.question.trim()" @click="onNlSend()">发送</el-button>
        </div>
      </div>
      <!-- 【tools-results-history-persistence 保留】已保存为历史任务提示卡 -->
      <div v-if="latestNlTaskId" class="saved-task-card">
        📦 已保存为历史任务 #{{ latestNlTaskId }} · 数据问答
        <el-button type="primary" size="small" @click="$router.push(`/reports/${latestNlTaskId}`)">
          查看完整报告
        </el-button>
      </div>
    </el-dialog>

    <!-- 数据集对比（4 种模式 + 跨模态对比） -->
    <el-dialog v-model="showCompare" title="数据集对比" width="780" top="6vh">
      <el-form label-width="100">
        <el-form-item label="左数据集">
          <el-select v-model="compareForm.left_dataset_id" style="width: 100%" placeholder="请选择左数据集" @change="onCompareDatasetChange">
            <el-option v-for="d in datasetOptions" :key="d.id" :label="d.name" :value="d.id">
              <span style="float:left">{{ d.name }}</span>
              <el-tag
                :type="d.file_type === 'image' ? 'warning' : (d.file_type === 'document' ? 'info' : 'success')"
                size="small"
                effect="plain"
                style="float:right;margin-top:2px"
              >
                {{ fileTypeLabel(d.file_type || 'structured') }}
              </el-tag>
            </el-option>
          </el-select>
        </el-form-item>
        <el-form-item label="右数据集">
          <el-select v-model="compareForm.right_dataset_id" style="width: 100%" placeholder="请选择右数据集" @change="onCompareDatasetChange">
            <el-option v-for="d in datasetOptions" :key="d.id" :label="d.name" :value="d.id">
              <span style="float:left">{{ d.name }}</span>
              <el-tag
                :type="d.file_type === 'image' ? 'warning' : (d.file_type === 'document' ? 'info' : 'success')"
                size="small"
                effect="plain"
                style="float:right;margin-top:2px"
              >
                {{ fileTypeLabel(d.file_type || 'structured') }}
              </el-tag>
            </el-option>
          </el-select>
        </el-form-item>

        <!-- 【dataset-multiformat-upload 新增】非结构化对比提示卡:
             选中任意非结构化数据集、且未显式选择合同对比时,
             自动走 LLM 跨模态对比,避免用户误解 -->
        <el-alert
          v-if="isDocumentCompare && compareForm.mode !== 'contract_compare'"
          type="info"
          :closable="false"
          show-icon
          style="margin-bottom: 14px"
        >
          <template #title>
            <span style="font-weight:600">📑 跨模态对比模式</span>
          </template>
          <div style="line-height:1.7">
            检测到本次对比涉及非结构化材料
            <el-tag
              v-if="compareTypePair.left !== 'structured'"
              size="small"
              effect="plain"
              :type="compareTypePair.left === 'image' ? 'warning' : 'info'"
            >左:{{ fileTypeLabel(compareTypePair.left) }}</el-tag>
            <el-tag
              v-if="compareTypePair.right !== 'structured'"
              size="small"
              effect="plain"
              :type="compareTypePair.right === 'image' ? 'warning' : 'info'"
              style="margin-left:4px"
            >右:{{ fileTypeLabel(compareTypePair.right) }}</el-tag>
            ,将由 AI 阅读两份原文并直接产出语义对比报告,无需选择分组/指标列。
          </div>
        </el-alert>

        <!-- 【dataset-contract-compare 新增】合同对比模式选项始终可见,
             即便两边是 structured 也能选(合同可能是 Excel 表格形式),
             选中后下方切换为 focus 输入框而非 Pandas 参数 -->
        <el-form-item label="对比模式">
          <el-radio-group v-model="compareForm.mode" @change="onCompareModeChange">
            <!-- Pandas 2 种:只在双 structured 时启用 -->
            <el-radio-button label="business" :disabled="isDocumentCompare">业务对比</el-radio-button>
            <el-radio-button label="row_diff" :disabled="isDocumentCompare">变更追踪</el-radio-button>
            <!-- 合同对比:任何类型都可用,因为合同可能也是 Excel/CSV -->
            <el-radio-button label="contract_compare">📑 合同对比</el-radio-button>
          </el-radio-group>
        </el-form-item>

        <!-- 合同对比专属参数:重点关注方向(可选) -->
        <template v-if="compareForm.mode === 'contract_compare'">
          <el-form-item label="重点关注">
            <el-input
              v-model="compareForm.focus"
              type="textarea"
              :rows="2"
              placeholder="可选,例如:重点对比付款账期、违约金比例、管辖地变更…留空则全面对比"
            />
          </el-form-item>
          <el-alert
            type="info"
            :closable="false"
            show-icon
            style="margin-bottom: 12px"
          >
            <template #title>
              <span style="font-weight:600">📑 合同条款级对比</span>
            </template>
            <div style="line-height:1.7">
              将逐条对比 12 类核心条款(标的/金额/期限/付款/违约/争议解决等),
              输出「条款差异表 + 利弊分析 + 谈判建议」,适合相邻年度合同对比场景。
            </div>
          </el-alert>
        </template>

        <!-- Pandas 对比模式:仅在非合同 + 双 structured 时显示参数 -->
        <template v-if="compareForm.mode !== 'contract_compare' && !isDocumentCompare">
          <template v-if="compareForm.mode === 'business'">
            <el-form-item label="分组列">
              <el-select v-model="compareForm.group_col" style="width: 100%" placeholder="按哪一列分组（如 月份/部门）">
                <el-option v-for="c in compareColumnsOptions" :key="c" :label="c" :value="c" />
              </el-select>
            </el-form-item>
            <el-form-item label="指标列">
              <el-select v-model="compareForm.metric_col" style="width: 100%" placeholder="对比的数值列（如 业绩）">
                <el-option v-for="c in compareColumnsOptions" :key="c" :label="c" :value="c" />
              </el-select>
            </el-form-item>
            <!-- 【topn 并入 business 新增】可选的 TopN 截断:
                 关闭时对比全部分组;开启后只保留指标值最高的 N 个,
                 并附带「左名次 / 右名次」作为参考列。
                 注意:名次变化只是相对位移(有更强的新业务挤入会让老业务名次下降),
                 判断经营增减请以「差值 / 变化幅度」为准。 -->
            <el-form-item label="只看 Top N">
              <el-switch v-model="compareForm.use_top_n" />
              <el-input-number
                v-if="compareForm.use_top_n"
                v-model="compareForm.top_n"
                :min="1"
                :max="100"
                style="margin-left: 12px; width: 140px"
              />
              <span class="muted" style="margin-left: 12px">
                只对比指标值最高的分组
              </span>
            </el-form-item>
          </template>

          <template v-else-if="compareForm.mode === 'row_diff'">
            <el-form-item label="主键列">
              <el-select v-model="compareForm.key_cols" multiple style="width: 100%" placeholder="选择 1~N 个主键（可多列组合）">
                <el-option v-for="c in compareColumnsOptions" :key="c" :label="c" :value="c" />
              </el-select>
            </el-form-item>
          </template>
        </template>
      </el-form>

      <el-checkbox v-model="saveAsCompareTask" style="margin-bottom: 8px; display: block">
        💾 保存为历史任务
      </el-checkbox>
      <el-button type="primary" :loading="toolLoading" @click="onCompare">开始对比</el-button>
      <el-button @click="compareResult = null" style="margin-left: 8px">清空结果</el-button>
      <!-- 【tools-results-history-persistence 新增】已保存为历史任务提示卡 -->
      <div v-if="latestCompareTaskId" class="saved-task-card">
        📦 已保存为历史任务 #{{ latestCompareTaskId }} · 数据对比
        <el-button type="primary" size="small" @click="$router.push(`/reports/${latestCompareTaskId}`)">
          查看完整报告
        </el-button>
      </div>

      <div v-if="compareResult" class="compare-result" style="margin-top: 12px">
        <el-alert
          v-if="compareResult.error"
          :title="compareResult.error"
          type="error"
          :closable="false"
          show-icon
        />
        <template v-else>
          <div class="compare-narration">
            <h4 class="cn-title">
              AI 差异解读
              <!-- 【dataset-multiformat-upload 新增】跨模态对比时明确告知用户,避免误以为是表格对比 -->
              <el-tag
                v-if="compareResult.diff_mode === 'document_compare'"
                type="info"
                size="small"
                effect="plain"
                style="margin-left:8px;vertical-align:middle"
              >
                📑 跨模态对比
              </el-tag>
            </h4>
            <p v-if="compareResult.narration?.summary" class="cn-summary">{{ compareResult.narration.summary }}</p>
            <p v-else class="cn-empty">AI 解读暂不可用,以下为对比数字。</p>
            <div class="cn-section">
              <div class="cn-h">关键发现</div>
              <ul v-if="compareResult.narration?.key_findings?.length">
                <li v-for="(t, i) in compareResult.narration.key_findings" :key="'k' + i">{{ t }}</li>
              </ul>
              <p v-else class="cn-empty-line">无</p>
            </div>
            <div class="cn-section">
              <div class="cn-h">风险</div>
              <ul v-if="compareResult.narration?.risks?.length">
                <li v-for="(t, i) in compareResult.narration.risks" :key="'r' + i">{{ t }}</li>
              </ul>
              <p v-else class="cn-empty-line">未识别到风险</p>
            </div>
            <div class="cn-section">
              <div class="cn-h">建议</div>
              <ul v-if="compareResult.narration?.suggestions?.length">
                <li v-for="(t, i) in compareResult.narration.suggestions" :key="'s' + i">{{ t }}</li>
              </ul>
              <p v-else class="cn-empty-line">暂无额外建议</p>
            </div>
          </div>

          <div v-if="compareResult.diff_mode === 'business'" class="result-block">
            <h4 class="rb-title">
              业务对比(按 {{ compareResult.indicators?.group_col }})
              <!-- 【topn 并入 business 新增】TopN 截断时提示口径:
                   名次只是相对位移,不代表经营涨跌 -->
              <el-tag
                v-if="compareResult.indicators?.top_n"
                type="info"
                size="small"
                effect="plain"
                style="margin-left:8px;vertical-align:middle"
              >
                仅对比 Top{{ compareResult.indicators.top_n }}
              </el-tag>
            </h4>
            <p v-if="compareResult.indicators?.top_n" class="muted" style="margin-bottom:6px">
              「名次」仅为参考:名次变化可能只是相对位移(有更强的新业务挤入),
              判断增减请以「差值 / 变化幅度」为准。
            </p>
            <el-table :data="compareResult.diff_details" stripe size="small" max-height="320">
              <el-table-column prop="group" label="分组" min-width="120" />
              <el-table-column label="左值" min-width="100">
                <template #default="{ row }">{{ row.left_sum === -1 ? '—' : row.left_sum }}</template>
              </el-table-column>
              <el-table-column label="右值" min-width="100">
                <template #default="{ row }">{{ row.right_sum === -1 ? '—' : row.right_sum }}</template>
              </el-table-column>
              <el-table-column label="差值" min-width="100">
                <template #default="{ row }">
                  <span :class="diffClass(row.diff)">{{ formatSigned(row.diff) }}</span>
                </template>
              </el-table-column>
              <el-table-column label="变化幅度" min-width="120">
                <template #default="{ row }">
                  <span v-if="row.growth_pct === null" class="muted">—</span>
                  <span v-else :class="diffClass(row.diff)">{{ formatSigned(row.growth_pct) }}%</span>
                </template>
              </el-table-column>
              <!-- 【topn 并入 business 新增】名次列为辅助参考,
                   仅在启用 TopN 截断时出现 -->
              <el-table-column
                v-if="compareResult.indicators?.top_n"
                label="左名次"
                width="80"
              >
                <template #default="{ row }">{{ row.left_rank || '—' }}</template>
              </el-table-column>
              <el-table-column
                v-if="compareResult.indicators?.top_n"
                label="右名次"
                width="80"
              >
                <template #default="{ row }">{{ row.right_rank || '—' }}</template>
              </el-table-column>
            </el-table>
          </div>

          <div v-else-if="compareResult.diff_mode === 'row_diff'" class="result-block">
            <h4 class="rb-title">变更追踪</h4>
            <div class="row-diff-tags">
              <div class="rd-tag rd-added">
                <div class="tag-label">新增</div>
                <div class="tag-num">{{ compareResult.indicators?.added_count || 0 }}</div>
              </div>
              <div class="rd-tag rd-removed">
                <div class="tag-label">删除</div>
                <div class="tag-num">{{ compareResult.indicators?.removed_count || 0 }}</div>
              </div>
              <div class="rd-tag rd-modified">
                <div class="tag-label">修改</div>
                <div class="tag-num">{{ compareResult.indicators?.modified_count || 0 }}</div>
              </div>
            </div>
            <p v-if="compareResult.diff_details?.truncated" class="muted" style="margin-bottom:6px">
              还有 {{ (compareResult.diff_details?.modified_total || 0) - (compareResult.diff_details?.modified?.length || 0) }} 条修改未展开
            </p>
            <el-collapse>
              <el-collapse-item :title="`新增明细（${compareResult.diff_details?.added?.length || 0} 行）`" name="a">
                <p v-if="!(compareResult.diff_details?.added?.length)" class="cn-empty-line">无</p>
              </el-collapse-item>
              <el-collapse-item :title="`删除明细（${compareResult.diff_details?.removed?.length || 0} 行）`" name="r">
                <p v-if="!(compareResult.diff_details?.removed?.length)" class="cn-empty-line">无</p>
              </el-collapse-item>
              <el-collapse-item :title="`修改明细（已展示 ${compareResult.diff_details?.modified?.length || 0} / 共 ${compareResult.diff_details?.modified_total || 0} 行）`" name="m">
                <div v-for="(m, idx) in compareResult.diff_details?.modified || []" :key="'m' + idx" class="modified-row">
                  <div class="modified-key">主键: {{ (m.key_values || []).map(v => v ?? '∅').join(' / ') }}</div>
                  <el-table :data="m.changes" stripe size="small" max-height="200" class="modified-table">
                    <el-table-column prop="field" label="字段" min-width="120" />
                    <el-table-column prop="before" label="修改前" min-width="120">
                      <template #default="{ row }"><span class="cell-before">{{ row.before ?? '∅' }}</span></template>
                    </el-table-column>
                    <el-table-column prop="after" label="修改后" min-width="120">
                      <template #default="{ row }"><span class="cell-after">{{ row.after ?? '∅' }}</span></template>
                    </el-table-column>
                  </el-table>
                </div>
                <p v-if="!(compareResult.diff_details?.modified?.length)" class="cn-empty-line">无</p>
              </el-collapse-item>
            </el-collapse>
          </div>

          <!-- 【dataset-contract-compare 新增】合同对比结果区:
               - 条款差异表(12 类核心条款的逐条对比)
               - 利弊分析(左右方各自优劣,业务视角)
               - 谈判建议(可操作的签约前提示) -->
          <div v-else-if="compareResult.diff_mode === 'contract_compare'" class="result-block">
            <h4 class="rb-title">
              📑 合同条款差异表
              <el-tag
                v-if="contractHighChangeCount > 0"
                type="danger"
                size="small"
                effect="dark"
                style="margin-left:8px;vertical-align:middle"
              >
                {{ contractHighChangeCount }} 项高优先级变更
              </el-tag>
            </h4>

            <!-- 条款差异表 -->
            <el-table
              :data="compareResult.diff_details || []"
              stripe
              size="small"
              max-height="380"
              :row-class-name="contractRowClass"
            >
              <el-table-column label="条款" min-width="130">
                <template #default="{ row }">
                  <span class="contract-clause-label">{{ row.clause_label }}</span>
                </template>
              </el-table-column>
              <el-table-column label="左方(旧)" min-width="160" show-overflow-tooltip>
                <template #default="{ row }">
                  <span v-if="row.left_value">{{ row.left_value }}</span>
                  <span v-else class="dim small">—</span>
                </template>
              </el-table-column>
              <el-table-column label="右方(新)" min-width="160" show-overflow-tooltip>
                <template #default="{ row }">
                  <span v-if="row.right_value">{{ row.right_value }}</span>
                  <span v-else class="dim small">—</span>
                </template>
              </el-table-column>
              <el-table-column label="差异说明" min-width="220">
                <template #default="{ row }">
                  <span v-if="row.diff_note">{{ row.diff_note }}</span>
                  <span v-else class="dim small">无明显差异</span>
                </template>
              </el-table-column>
              <el-table-column label="类型" width="100">
                <template #default="{ row }">
                  <el-tag :type="contractChangeTypeTag(row.change_type).type" size="small" effect="plain">
                    {{ contractChangeTypeTag(row.change_type).label }}
                  </el-tag>
                </template>
              </el-table-column>
              <el-table-column label="优先级" width="80">
                <template #default="{ row }">
                  <el-tag
                    v-if="row.importance === 'high'"
                    type="danger"
                    size="small"
                    effect="dark"
                  >高</el-tag>
                  <el-tag
                    v-else-if="row.importance === 'low'"
                    type="info"
                    size="small"
                    effect="plain"
                  >低</el-tag>
                  <el-tag v-else size="small" effect="plain">中</el-tag>
                </template>
              </el-table-column>
            </el-table>

            <!-- 利弊分析 -->
            <div v-if="compareResult.narration?.pros_cons?.length" class="contract-section">
              <h4 class="rb-title">⚖️ 利弊分析</h4>
              <div
                v-for="(pc, idx) in compareResult.narration.pros_cons"
                :key="'pc' + idx"
                class="pros-cons-card"
                :class="{
                  'pc-left':  pc.side === 'left',
                  'pc-right': pc.side === 'right',
                  'pc-both':  pc.side === 'both' || !pc.side,
                }"
              >
                <div class="pc-head">
                  <el-tag
                    :type="pc.side === 'left' ? 'success' : (pc.side === 'right' ? 'warning' : 'info')"
                    size="small"
                    effect="dark"
                  >
                    {{ pc.side === 'left' ? '左方更优' : (pc.side === 'right' ? '右方更优' : '各有优劣') }}
                  </el-tag>
                  <span class="pc-title">{{ pc.title }}</span>
                </div>
                <div class="pc-detail">{{ pc.detail }}</div>
              </div>
            </div>

            <!-- 谈判建议 -->
            <div v-if="compareResult.narration?.negotiation_tips?.length" class="contract-section">
              <h4 class="rb-title">💼 谈判建议</h4>
              <ul class="contract-tips">
                <li v-for="(t, i) in compareResult.narration.negotiation_tips" :key="'tip' + i">
                  <el-icon style="color:#22d3ee;margin-right:4px;vertical-align:middle"><Star /></el-icon>
                  {{ t }}
                </li>
              </ul>
            </div>
          </div>
        </template>
      </div>
    </el-dialog>

    <!-- 图表编辑(自然语言改图表 · 支持当前 + 历史)
         【chart-edit-prebuilt-dashboard 新增】
         - 「当前工作台」数据源下,前置生成 6~8 张可视化看板(无需先跑 AI 分析)
         - 3 套图表风格下拉
         - 每张卡:单图下载 PNG / + 添加至报表
         - 自然语言 / 快捷指令改图(命中规则即改,否则 LLM 兜底)
         - 整套重新生成 + 保留勾选
         - 下载整个数据可视化的看板(整页 PNG)
    -->
    <el-dialog v-model="showChart" title="图表编辑 · AI 改图" width="960" top="4vh" class="chart-edit-dialog">
      <p class="dim small" style="margin: 0 0 10px">
        用自然语言告诉 AI 怎么改图表。比如:<b>"换成折线图"</b>、<b>"配色换成蓝绿色"</b>、<b>"只显示 Top 5"</b>、<b>"加一条同比线"</b>。
        <br />
        <b>当前工作台模式</b>还支持"前置生成可视化看板":上传数据集后即可挑图加入报表,无需先跑 AI 分析。
        <br />
        <span style="color:#0e7490">💡</span> 工具栏勾选 <b>AI 智能筛选</b> 时,本地规则出骨架 + LLM 精选 TopN + 标题润色,有效避免图表过多或为空。
      </p>

      <!-- 数据源选择 -->
      <el-form label-width="80" style="margin-bottom: 10px">
        <el-form-item label="数据源">
          <el-radio-group v-model="chartSource">
            <el-radio-button label="workbench">当前工作台</el-radio-button>
            <el-radio-button label="history">历史任务</el-radio-button>
          </el-radio-group>
        </el-form-item>

        <el-form-item v-if="chartSource === 'history'" label="选择报告">
          <el-select
            v-model="chartHistoryTaskId"
            placeholder="选择一份历史报告"
            filterable
            style="width: 100%"
            @change="onHistoryTaskChange"
          >
            <el-option
              v-for="t in historyTaskOptions"
              :key="t.id"
              :label="`#${t.id} ${t.name} · ${t.chart_count || 0} 个图表 · ${t.status}`"
              :value="t.id"
            />
          </el-select>
          <div v-if="historyTaskOptions.length === 0" class="dim small" style="margin-top: 6px">
            暂无历史任务,请先在「开始 AI 分析」生成报告后再来。
          </div>
        </el-form-item>

        <el-form-item v-if="chartSource === 'history' && chartHistoryTaskId" label="选择图表">
          <el-select
            v-model="chartHistoryIndex"
            placeholder="选择图表"
            style="width: 100%"
          >
            <el-option
              v-for="(c, i) in historyChartOptions"
              :key="i"
              :label="`图表${i + 1}:${c.title || c.type}`"
              :value="i"
            />
          </el-select>
        </el-form-item>
      </el-form>

      <!-- ===== 【chart-edit-prebuilt-dashboard】前置看板顶部工具栏 ===== -->
      <div v-if="chartSource === 'workbench'" class="dashboard-toolbar">
        <el-form-item label="图表风格" class="tb-style">
          <el-select v-model="wbStore.draft.chartStyle" style="width: 180px">
            <el-option
              v-for="s in chartStyleOptions"
              :key="s.key"
              :label="s.label"
              :value="s.key"
            />
          </el-select>
        </el-form-item>
        <el-button
          type="primary"
          :loading="chartPreviewLoading"
          :disabled="!wbStore.draft.datasetId"
          @click="onRegenerateDashboard"
        >
          {{ wbStore.draft.dashboardCharts.length ? '生成新图表' : '生成新图表' }}
        </el-button>
        <!-- 【chart-render-smart-2026-10 新增】AI 智能筛选开关 -->
        <el-checkbox
          v-model="wbStore.draft.aiCurate"
          style="margin: 0 8px 0 12px"
          @change="onAiCurateToggle"
        >
          AI 智能筛选
        </el-checkbox>
        <el-select
          v-if="wbStore.draft.aiCurate"
          v-model="wbStore.draft.aiCurateTarget"
          size="default"
          style="width: 110px; margin-right: 8px"
        >
          <el-option
            v-for="n in [2, 3, 4, 5, 6]"
            :key="n"
            :label="`每 sheet 留 ${n} 张`"
            :value="n"
          />
        </el-select>
        <el-checkbox
          v-if="wbStore.draft.dashboardCharts.length"
          v-model="wbStore.draft.keepSelectionOnRegen"
          style="margin: 0 8px"
        >
          保留勾选
        </el-checkbox>
        <el-button
          v-if="wbStore.draft.dashboardCharts.length"
          type="success"
          :disabled="wbStore.getSelectedCharts().length === 0"
          @click="downloadWholeDashboard"
        >
          ⬇ 下载整个看板
        </el-button>
        <!-- 【export-visual-report 新增】导出 Excel 可视化报表
             与【下载整个看板】的区别：看板是纯位图长图（不可编辑/不可筛选/无 KPI 无结论），
             报表是结构化 xlsx（KPI 卡 + AI 结论 + 数据明细 + 图表洞察），可以带走直接用。 -->
        <el-tooltip
          :disabled="wbStore.getSelectedCharts().length > 0"
          content="请先勾选要加入报表的图表"
          placement="top"
        >
          <span style="margin-left: 8px">
            <el-button
              v-if="wbStore.draft.dashboardCharts.length"
              type="primary"
              plain
              :loading="exportingReport"
              :disabled="wbStore.getSelectedCharts().length === 0"
              @click="exportVisualReportFromDashboard"
            >
              📊 导出 Excel 报表
            </el-button>
          </span>
        </el-tooltip>
        <span v-if="wbStore.draft.dashboardCharts.length" class="dim small tb-stat">
          已选 <b style="color:#0e7490">{{ wbStore.getSelectedCharts().length }}</b> / {{ wbStore.draft.dashboardCharts.length }} 张加入报表
        </span>
      </div>

      <!-- ===== 当前工作台 · 前置看板(按 sheet 分组) ===== -->
      <div v-if="chartSource === 'workbench' && dashboardGroups.length" class="dashboard-groups">
        <div
          v-for="group in dashboardGroups"
          :key="group.sheet"
          class="dashboard-group"
        >
          <div class="dashboard-group-header">
            <span class="dashboard-group-title">📑 {{ group.sheet || '默认' }}</span>
            <span class="dim small">{{ group.charts.length }} 张图表</span>
          </div>
          <div class="dashboard-grid">
            <div
              v-for="c in group.charts"
              :key="c._key"
              :data-chart-key="c._key"
              class="chart-card"
              :class="{ selected: wbStore.draft.selectedChartKeys.has(c._key) }"
            >
              <div class="chart-card-header">
                <div class="chart-card-title">{{ c.title }}</div>
                <div class="chart-card-actions">
                  <el-button
                    size="small"
                    type="primary"
                    plain
                    @click.stop="downloadSingleChart(c, c._key)"
                  >
                    ↓PNG
                  </el-button>
                  <el-button
                    size="small"
                    :type="wbStore.draft.selectedChartKeys.has(c._key) ? 'success' : 'default'"
                    @click.stop="wbStore.toggleChartSelected(c._key)"
                  >
                    {{ wbStore.draft.selectedChartKeys.has(c._key) ? '✓ 已加入' : '+ 报表' }}
                  </el-button>
                </div>
              </div>
              <div class="chart-card-canvas">
                <v-chart :option="c.option" autoresize />
              </div>
            </div>
          </div>
        </div>
      </div>
      <div
        v-else-if="chartSource === 'workbench' && !dashboardGroups.length"
        class="dashboard-empty"
      >
        <div class="de-icon">📊</div>
        <div class="de-text">
          尚未生成看板。选择风格后点击「生成新图表」即可基于当前数据集的所有工作表预览可视化卡片。
        </div>
      </div>

      <!-- ===== 历史模式改图(原 chartEdit 路径,行为不变) ===== -->
      <div v-if="chartSource === 'history' && activeChart" class="chart-preview-block">
        <div class="chart-preview-title">📊 当前图表({{ activeChart.type }}):{{ activeChart.title }}</div>
        <v-chart class="chart-canvas" :option="activeChart.option" autoresize />
      </div>
      <div v-else-if="chartSource === 'history'" class="dim small" style="margin-bottom: 8px">
        该历史任务没有图表或正在加载。
      </div>

      <!-- 指令输入 + 提交 -->
      <div class="chart-edit-row">
        <el-input
          v-model="chartInstruction"
          placeholder="例如:换成折线图 + 配色换成蓝绿色"
          clearable
          @keyup.enter="onChartEdit"
        />
        <el-button
          type="primary"
          :loading="chartEditLoading"
          :disabled="!chartInstruction.trim() || !activeDatasetId"
          @click="onChartEdit"
        >
          生成新图表
        </el-button>
      </div>

      <!-- 快捷指令 -->
      <div class="quick-instructions">
        <span class="dim small">快捷指令:</span>
        <el-tag
          v-for="(t, i) in QUICK_CHART_INSTRUCTIONS"
          :key="i"
          class="quick-tag"
          effect="plain"
          @click="chartInstruction = t"
        >
          {{ t }}
        </el-tag>
      </div>

      <!-- 编辑结果 -->
      <div v-if="chartEditResult" class="chart-edit-result">
        <el-alert
          v-if="chartEditResult.fallback && !(chartEditResult.explanation || '').includes('本地降级')"
          type="warning"
          :title="chartEditResult.error || '本次编辑失败,已保留原图'"
          :closable="false"
          show-icon
        />
        <el-alert
          v-else-if="chartEditResult.fallback"
          type="info"
          :title="'AI 暂时不可用,已本地降级: ' + (chartEditResult.explanation || '已按指令改图')"
          :closable="false"
          show-icon
        />
        <div class="chart-preview-title">
          ✨ 新图表({{ chartEditResult.type }}):{{ chartEditResult.explanation }}
          <el-tag v-if="chartEditResult.saved_to_task_id" type="info" size="small" effect="plain">
            候选图,尚未写入任务
          </el-tag>
        </div>
        <v-chart
          v-if="chartEditResult.option && Object.keys(chartEditResult.option).length"
          class="chart-canvas"
          :option="chartEditResult.option"
          autoresize
        />
        <div v-else class="dim small">新图表为空,可能编辑失败。</div>

        <div class="chart-actions">
          <!-- 工作台模式:把新生成的候选图"替换"看板中选中的那张(或第一张) -->
          <template v-if="chartSource === 'workbench'">
            <el-button
              type="primary"
              :disabled="!chartEditResult.option || !Object.keys(chartEditResult.option).length"
              @click="applyChartToDashboard"
            >
              替换到看板
            </el-button>
          </template>
          <!-- 历史模式:候选图不直接落库,让用户显式确认 -->
          <template v-else>
            <el-button
              type="primary"
              :disabled="!chartEditResult.option || !Object.keys(chartEditResult.option).length"
              @click="applyChartToHistoryTask"
            >
              应用到任务
            </el-button>
            <el-button @click="chartEditResult = null">放弃</el-button>
          </template>
          <el-button @click="chartEditResult = null">清空</el-button>
        </div>
      </div>
    </el-dialog>

    <!-- 分析重点输入弹窗【workbench-analysis-intent-input】 -->
    <!-- 「开始 AI 分析」按钮被点击时弹出:用户可选输入本次临时想要 AI 关注的重点 -->
    <el-dialog v-model="showIntent" title="补充分析重点（可选）" width="540" :close-on-click-modal="false">
      <p class="dim small" style="margin: 0 0 10px">
        告诉 AI 这次想重点关注什么 / 不关注什么，例如：
        <b>「只看销售额下降的门店，不要分析库存」</b>
      </p>
      <el-input
        v-model="wbStore.draft.analysisIntent"
        type="textarea"
        :rows="4"
        :disabled="wbStore.draft.skipIntent"
        placeholder="本次分析重点"
      />
      <el-checkbox
        v-model="wbStore.draft.skipIntent"
        style="margin-top: 10px"
        @change="onSkipIntentChange"
      >
        本次跳过、直接生成报告
      </el-checkbox>
      <template #footer>
        <el-button @click="showIntent = false">取消</el-button>
        <el-button type="primary" :loading="analyzing" @click="confirmIntentAndAnalyze">
          生成报告
        </el-button>
      </template>
    </el-dialog>

    <!-- 数据清洗【workbench-standalone-cleansing + workbench-cleansing-missing-data】 -->
    <!-- 【clean-ai-derived-columns-pivot 调整】宽度 560 → 720,容纳派生列与透视汇总配置 -->
    <el-dialog v-model="showClean" title="数据清洗" width="720" :close-on-click-modal="false">
      <el-form label-width="100">
        <el-form-item label="选择数据集">
          <el-select v-model="cleanForm.dataset_id" style="width: 100%" placeholder="请选择数据集">
            <el-option
              v-for="d in datasetOptions"
              :key="d.id"
              :label="d.name"
              :value="d.id"
            />
          </el-select>
        </el-form-item>
        <el-form-item label="新数据集名称">
          <el-input v-model="cleanForm.name" placeholder="留空则默认 <原名>_cleaned" />
        </el-form-item>

        <!-- 4 项原有清洗规则(收敛到此处作为唯一 UI 入口) -->
        <el-form-item label="清洗规则">
          <el-checkbox v-model="wbStore.draft.cleaningRules.drop_empty_rows">清除空行</el-checkbox>
          <el-checkbox v-model="wbStore.draft.cleaningRules.drop_duplicates">去除重复</el-checkbox>
          <el-checkbox v-model="wbStore.draft.cleaningRules.unify_date">统一日期格式</el-checkbox>
          <el-checkbox v-model="wbStore.draft.cleaningRules.mark_anomaly">标记异常值</el-checkbox>
        </el-form-item>

        <!-- 【本 spec 新增】第 5 项规则:填充缺失值 + 二级策略配置 -->
        <el-form-item>
          <el-checkbox v-model="wbStore.draft.fill_missing">填充缺失值</el-checkbox>
        </el-form-item>
        <el-form-item v-if="wbStore.draft.fill_missing" label="填充策略">
          <el-radio-group v-model="wbStore.draft.fillMissingPolicy.strategy">
            <el-radio value="mean">均值</el-radio>
            <el-radio value="median">中位数</el-radio>
            <el-radio value="mode">众数</el-radio>
            <el-radio value="zero">置零</el-radio>
            <el-radio value="ffill">前向填充</el-radio>
            <el-radio value="drop">删除含缺失的行</el-radio>
          </el-radio-group>
        </el-form-item>
        <el-form-item v-if="wbStore.draft.fill_missing" label="作用范围">
          <el-radio-group v-model="wbStore.draft.fillMissingPolicy.scope">
            <el-radio value="numeric">仅数值列</el-radio>
            <el-radio value="all">所有列</el-radio>
          </el-radio-group>
        </el-form-item>

        <!-- 【clean-keep-formulas 新增】公式保留回写:勾选后清洗结果写回原工作簿,
             保留未改动单元格的公式 / 样式 / 多 sheet 布局(仅对 Excel 文件生效) -->
        <el-form-item>
          <el-checkbox v-model="wbStore.draft.keep_formulas">保留公式回写（仅 Excel，写回原表结构）</el-checkbox>
        </el-form-item>
      </el-form>

      <!-- ===== 【clean-ai-derived-columns-pivot 新增】智能派生列 ===== -->
      <div class="wb-clean-block">
        <div class="wb-clean-block-head">
          <div>
            <div class="wb-clean-block-title">智能派生列</div>
            <div class="wb-clean-block-tip">AI 会根据现有列推荐可计算的新列，例如销售额 = 销量 × 单价</div>
          </div>
          <el-button
            size="small"
            type="primary"
            :loading="wbStore.draft.suggestColumnsLoading"
            :disabled="!cleanForm.dataset_id"
            @click="onSuggestColumns"
          >
            {{ wbStore.draft.suggestColumnsLoading ? 'AI 分析中…' : 'AI 推荐派生列' }}
          </el-button>
        </div>

        <!-- AI 建议列表:复选框 + 可改列名 / 公式 + 只读说明 + 工作表标签;默认全选 -->
        <div v-if="wbStore.draft.derivedSuggestions.length" class="wb-derived-list">
          <div v-for="(row, i) in wbStore.draft.derivedSuggestions" :key="'sg' + i" class="wb-derived-item">
            <div class="wb-derived-line">
              <el-checkbox v-model="row.selected">启用</el-checkbox>
              <el-input v-model="row.name" size="small" placeholder="新列名" class="wb-derived-name" />
              <el-tag size="small" type="info" effect="plain">{{ row.sheet_name }}</el-tag>
              <span class="wb-derived-desc">{{ row.description || 'AI 推荐' }}</span>
            </div>
            <div class="wb-derived-line">
              <span class="wb-derived-label">公式</span>
              <el-input
                v-model="row.expression"
                size="small"
                placeholder="列名用反引号包裹,如 `销量` * `单价`"
              />
            </div>
          </div>
        </div>
        <div v-else class="wb-derived-empty">尚未生成推荐，可点击「AI 推荐派生列」或手动添加。</div>

        <!-- 手动添加的派生列:工作表 / 新列名 / 公式 / 删除 -->
        <div v-if="wbStore.draft.manualDerivedColumns.length" class="wb-derived-list">
          <div v-for="(row, i) in wbStore.draft.manualDerivedColumns" :key="'md' + i" class="wb-derived-item">
            <div class="wb-derived-line">
              <el-select v-model="row.sheet_name" size="small" placeholder="工作表" class="wb-derived-sheet">
                <el-option v-for="s in cleanSheetNames" :key="s" :label="s" :value="s" />
              </el-select>
              <el-input v-model="row.name" size="small" placeholder="新列名" class="wb-derived-name" />
              <el-button size="small" link type="danger" @click="removeManualDerived(i)">删除</el-button>
            </div>
            <div class="wb-derived-line">
              <span class="wb-derived-label">公式</span>
              <el-input
                v-model="row.expression"
                size="small"
                placeholder="列名用反引号包裹,如 `销量` * `单价`"
              />
            </div>
          </div>
        </div>
        <el-button size="small" class="wb-derived-add" @click="addManualDerived">手动添加派生列</el-button>
      </div>

      <!-- ===== 透视汇总区块:两区域分离式布局 ===== -->
      <div class="wb-clean-block wb-pivot-block">
        <!-- 区块标题(横向,无旋转) -->
        <div class="wb-pivot-block-title-row">
          <span class="wb-clean-block-title">透视汇总</span>
          <span class="wb-clean-block-tip">可同时配置多张透视表,每张独立选择工作表/维度/汇总字段</span>
        </div>

        <!-- ========== 区域 A:AI 推荐透视方案(可折叠卡片式列表) ========== -->
        <div
          v-if="wbStore.draft.pivotSuggestions.length"
          class="wb-pivot-recommend-section"
        >
          <div class="wb-pivot-recommend-head">
            <span class="wb-pivot-recommend-label">
              <span class="wb-pivot-recommend-icon">✦</span>
              AI 推荐了 {{ wbStore.draft.pivotSuggestions.length }} 条透视方案
            </span>
            <el-button
              link
              size="small"
              class="wb-pivot-recommend-toggle"
              @click="pivotRecommendCollapsed = !pivotRecommendCollapsed"
            >{{ pivotRecommendCollapsed ? '展开' : '收起' }}</el-button>
          </div>
          <div v-show="!pivotRecommendCollapsed" class="wb-pivot-recommend-list">
            <div
              v-for="(item, idx) in wbStore.draft.pivotSuggestions"
              :key="idx"
              class="wb-pivot-recommend-item"
            >
              <!-- 第一行:结构化信息 [工作表tag] 行维度 × 列维度 聚合 › 汇总值 -->
              <div class="wb-pivot-recommend-main">
                <el-tag size="small" class="wb-pivot-sheet-tag">{{ item.sheet_name }}</el-tag>
                <span class="wb-pivot-recommend-dim">
                  {{ item.row_dim }}
                  <template v-if="item.col_dim"> × {{ item.col_dim }}</template>
                </span>
                <span class="wb-pivot-recommend-op">
                  {{ { sum:'求和', mean:'平均', count:'计数', max:'最大', min:'最小' }[item.agg] || item.agg }}
                </span>
                <span class="wb-pivot-recommend-arrow">›</span>
                <span class="wb-pivot-recommend-val">
                  {{ item.value }}
                  <el-tag
                    v-if="item.value_is_derived"
                    size="small"
                    class="wb-pivot-derived-tag-soft"
                  >🔧派生列</el-tag>
                </span>
              </div>
              <!-- 第二行:中文描述 + 应用按钮 -->
              <div class="wb-pivot-recommend-footer">
                <span class="wb-pivot-recommend-desc">
                  {{ item.description || `按${item.row_dim}汇总${item.value}` }}
                </span>
                <el-button
                  size="small"
                  class="wb-pivot-recommend-apply"
                  @click="applyPivotSuggestion(item)"
                >应用</el-button>
              </div>
            </div>
          </div>
        </div>

        <!-- 分隔线:AI 推荐存在 且 已选配置也存在 时显示 -->
        <div
          v-if="wbStore.draft.pivotSuggestions.length && wbStore.draft.pivotConfigs?.length"
          class="wb-pivot-divider"
        />

        <!-- ========== 区域 B:已选透视汇总 ========== -->
        <div class="wb-pivot-selected-section">
          <div class="wb-pivot-selected-head">
            <span class="wb-pivot-selected-label">
              已选透视汇总
              <span v-if="wbStore.draft.pivotConfigs?.length" class="wb-pivot-selected-count">
                ({{ wbStore.draft.pivotConfigs.length }})
              </span>
            </span>
            <el-button
              link
              size="small"
              type="primary"
              class="wb-pivot-add-manual"
              @click="addManualPivot"
            >+ 手动添加透视</el-button>
          </div>

          <!-- 空状态提示 -->
          <div v-if="!wbStore.draft.pivotConfigs?.length" class="wb-pivot-empty">
            暂无已选配置 — 点击上方 AI 方案的「应用」或「+ 手动添加透视」开始
          </div>

          <!-- 已选配置卡片列表 -->
          <div v-else class="wb-pivot-config-list">
            <div
              v-for="(pivot, pIdx) in wbStore.draft.pivotConfigs"
              :key="pivot.id"
              class="wb-pivot-config-card"
            >
              <!-- 卡片头部:编号 + 工作表 + 删除按钮 -->
              <div class="wb-pivot-config-head">
                <div class="wb-pivot-config-title-row">
                  <span class="wb-pivot-config-badge">#{{ pIdx + 1 }}</span>
                  <span class="wb-pivot-config-sheet">{{ pivot.sheet_name || '未选择工作表' }}</span>
                </div>
                <el-button
                  link
                  size="small"
                  type="danger"
                  @click="removePivotConfig(pivot.id)"
                >删除</el-button>
              </div>

              <!-- 表单:两列 grid 布局,聚合/总计行占整行 -->
              <div class="wb-pivot-grid">
                <div class="wb-pivot-item">
                  <span class="wb-pivot-label">源工作表</span>
                  <el-select
                    v-model="pivot.sheet_name"
                    size="small"
                    placeholder="请选择"
                    class="wb-pivot-field"
                    @change="onPivotSheetChange(pivot.id)"
                  >
                    <el-option v-for="s in cleanSheetNames" :key="s" :label="s" :value="s" />
                  </el-select>
                </div>
                <div class="wb-pivot-item">
                  <span class="wb-pivot-label">行维度</span>
                  <el-select
                    v-model="pivot.row_dim"
                    size="small"
                    placeholder="请选择"
                    class="wb-pivot-field"
                  >
                    <el-option
                      v-for="c in pivotRowOptionsForSheet(pivot.sheet_name)"
                      :key="c.name"
                      :label="c.name"
                      :value="c.name"
                    />
                  </el-select>
                </div>
                <div class="wb-pivot-item">
                  <span class="wb-pivot-label">列维度</span>
                  <el-select
                    v-model="pivot.col_dim"
                    size="small"
                    placeholder="可选"
                    clearable
                    class="wb-pivot-field"
                  >
                    <el-option
                      v-for="c in pivotSheetColumnsForSheet(pivot.sheet_name)"
                      :key="c.name"
                      :label="c.name"
                      :value="c.name"
                    />
                  </el-select>
                </div>
                <div class="wb-pivot-item">
                  <span class="wb-pivot-label">汇总字段</span>
                  <el-select
                    v-model="pivot.value"
                    size="small"
                    placeholder="数值列"
                    class="wb-pivot-field"
                  >
                    <el-option
                      v-for="c in pivotValueOptionsForSheet(pivot.sheet_name)"
                      :key="c.name"
                      :label="c.name"
                      :value="c.name"
                    >
                      <span>{{ c.name }}</span>
                      <el-tag v-if="c.from_derived" size="small" class="wb-derived-option-tag">🔧</el-tag>
                    </el-option>
                  </el-select>
                </div>
                <div class="wb-pivot-item wb-pivot-full">
                  <span class="wb-pivot-label">聚合方式</span>
                  <el-radio-group v-model="pivot.agg" size="small">
                    <el-radio value="sum">求和</el-radio>
                    <el-radio value="mean">平均</el-radio>
                    <el-radio value="count">计数</el-radio>
                    <el-radio value="max">最大</el-radio>
                    <el-radio value="min">最小</el-radio>
                  </el-radio-group>
                </div>
                <div class="wb-pivot-item wb-pivot-full">
                  <span class="wb-pivot-label">总计行</span>
                  <el-checkbox v-model="pivot.include_margin">显示总计行</el-checkbox>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- 【本 spec 新增】清洗预览行:由 onCleanSubmit 成功后写入 store,初始 null 时显示「—」 -->
      <div class="wb-clean-preview">
        <template v-if="wbStore.draft.previewCleanedStats">
          <div>
            清洗预览:原 {{ wbStore.draft.previewCleanedStats.original_rows }} 行 →
            新 {{ wbStore.draft.previewCleanedStats.cleaned_rows }} 行,
            填充 {{ wbStore.draft.previewCleanedStats.missing_filled_count }} 个缺失值,
            删除 {{ wbStore.draft.previewCleanedStats.dropped_rows_by_missing }} 行，
            新增 {{ cleanPreviewDerivedCount }} 个派生列
            <!-- 【clean-derived-excel-formula 新增】派生列已写入活公式时给出提示,
                 让用户知道导出的 Excel 打开后可联动重算 -->
            <template v-if="wbStore.draft.previewCleanedStats.derived_formula_cells > 0">
              (已写入 {{ wbStore.draft.previewCleanedStats.derived_formula_cells }} 个 Excel 公式,导出文件中可联动计算)
            </template>
          </div>
          <!-- 【透视多张并行改造】遍历 pivot_sheets 数组显示每张透视结果 -->
          <div v-if="wbStore.draft.previewCleanedStats?.pivot_sheets?.length">
            透视汇总:共 {{ wbStore.draft.previewCleanedStats.pivot_sheets.length }} 张 —
            {{ wbStore.draft.previewCleanedStats.pivot_sheets.map(p => `${p.sheet_name}(${p.rows}行)`).join(', ') }}
          </div>
          <!-- 向后兼容旧单字段 -->
          <div v-else-if="wbStore.draft.previewCleanedStats?.pivot_sheet_name">
            透视汇总:{{ wbStore.draft.previewCleanedStats.pivot_sheet_name }} 共
            {{ wbStore.draft.previewCleanedStats.pivot_rows || 0 }} 行
          </div>
          <div v-if="cleanPreviewWarnings.length" class="wb-clean-warnings">
            {{ cleanPreviewWarnings.join('；') }}
          </div>
        </template>
        <template v-else>
          清洗预览:—
        </template>
      </div>

      <template #footer>
        <el-button @click="showClean = false">取消</el-button>
        <el-button
          type="primary"
          :loading="cleanLoading"
          :disabled="!cleanForm.dataset_id"
          @click="onCleanSubmit"
        >
          确认清洗并保存为新数据集
        </el-button>
      </template>
    </el-dialog>

    <!-- 【workbench-ai-pdf-tools 新增】AI PDF 处理弹窗 -->
    <el-dialog
      v-model="showPdfTools"
      title="PDF 处理"
      width="640"
      :close-on-click-modal="false"
      @close="onPdfDialogClosed"
    >
      <el-tabs v-model="pdfTab">
        <!-- Tab 1: PDF 转 Word -->
        <el-tab-pane label="PDF 转 Word" name="word">
          <el-upload
            drag
            :auto-upload="false"
            :show-file-list="true"
            :limit="1"
            :on-change="(f) => { if (f?.raw) beforePdfUpload(f.raw); else onPdfRemove() }"
            :before-upload="beforePdfUpload"
            :on-remove="onPdfRemove"
            accept=".pdf"
          >
            <el-icon class="el-icon--upload"><Document /></el-icon>
            <div class="el-upload__text">拖拽 PDF 到此 或 <em>点击选择</em></div>
            <template #tip>
              <div class="el-upload__tip">仅支持 .pdf 文件,大小 ≤ {{ PDF_MAX_SIZE_MB }}MB</div>
            </template>
          </el-upload>
        </el-tab-pane>

        <!-- Tab 2: PDF 转表格(AI 辅助) -->
        <el-tab-pane label="PDF 转表格" name="excel">
          <el-upload
            drag
            :auto-upload="false"
            :limit="1"
            :on-change="(f) => { if (f?.raw) beforePdfUpload(f.raw); else onPdfRemove() }"
            :before-upload="beforePdfUpload"
            :on-remove="onPdfRemove"
            accept=".pdf"
          >
            <el-icon class="el-icon--upload"><Document /></el-icon>
            <div class="el-upload__text">拖拽 PDF 到此 或 <em>点击选择</em></div>
            <template #tip>
              <div class="el-upload__tip">仅支持 .pdf 文件,大小 ≤ {{ PDF_MAX_SIZE_MB }}MB</div>
            </template>
          </el-upload>
          <el-form label-width="100" style="margin-top: 12px">
            <el-form-item label="AI 辅助识别">
              <el-switch v-model="pdfAiEnabled" />
              <span style="margin-left: 8px; color: #888; font-size: 12px">
                开启后 LLM 协助识别表头(失败自动降级)
              </span>
            </el-form-item>
            <el-form-item label="保存为数据集">
              <el-switch v-model="pdfSaveAsDataset" />
              <span style="margin-left: 8px; color: #888; font-size: 12px">
                转换结果落为新的数据集(出现在「数据集」列表)
              </span>
            </el-form-item>
          </el-form>
        </el-tab-pane>

        <!-- Tab 3: PDF 提取文字 -->
        <el-tab-pane label="PDF 提取文字" name="text">
          <el-upload
            drag
            :auto-upload="false"
            :limit="1"
            :on-change="(f) => { if (f?.raw) beforePdfUpload(f.raw); else onPdfRemove() }"
            :before-upload="beforePdfUpload"
            :on-remove="onPdfRemove"
            accept=".pdf"
          >
            <el-icon class="el-icon--upload"><Document /></el-icon>
            <div class="el-upload__text">拖拽 PDF 到此 或 <em>点击选择</em></div>
            <template #tip>
              <div class="el-upload__tip">仅支持 .pdf 文件,大小 ≤ {{ PDF_MAX_SIZE_MB }}MB</div>
            </template>
          </el-upload>
          <el-form label-width="100" style="margin-top: 12px">
            <el-form-item label="AI 清洗排版">
              <el-switch v-model="pdfAiEnabled" />
              <span style="margin-left: 8px; color: #888; font-size: 12px">
                开启后 LLM 合并多余换行 / 去除页眉页脚(失败自动降级)
              </span>
            </el-form-item>
          </el-form>
        </el-tab-pane>

        <!-- Tab 4: PDF 压缩 -->
        <el-tab-pane label="PDF 压缩" name="compress">
          <el-upload
            drag
            :auto-upload="false"
            :limit="1"
            :on-change="(f) => { if (f?.raw) beforePdfUpload(f.raw); else onPdfRemove() }"
            :before-upload="beforePdfUpload"
            :on-remove="onPdfRemove"
            accept=".pdf"
          >
            <el-icon class="el-icon--upload"><Document /></el-icon>
            <div class="el-upload__text">拖拽 PDF 到此 或 <em>点击选择</em></div>
            <template #tip>
              <div class="el-upload__tip">仅支持 .pdf 文件,大小 ≤ {{ PDF_MAX_SIZE_MB }}MB</div>
            </template>
          </el-upload>
        </el-tab-pane>

        <!-- 【pdf-photo-to-pdf-ocr 新增】Tab 5: 照片转 PDF(OCR 识别) -->
        <el-tab-pane label="照片转 PDF" name="image">
          <el-upload
            drag
            multiple
            :auto-upload="false"
            :show-file-list="true"
            :limit="PDF_MAX_IMG_COUNT"
            :on-change="onImgFileChange"
            :on-remove="onImgRemove"
            :file-list="pdfImgFileList"
            accept=".png,.jpg,.jpeg,.webp,.bmp"
          >
            <el-icon class="el-icon--upload"><Picture /></el-icon>
            <div class="el-upload__text">拖拽照片到此(可多选) 或 <em>点击选择</em></div>
            <template #tip>
              <div class="el-upload__tip">
                支持 .png/.jpg/.jpeg/.webp/.bmp,最多 {{ PDF_MAX_IMG_COUNT }} 张,单张 ≤ {{ PDF_MAX_IMG_SIZE_MB }}MB
              </div>
              <div class="el-upload__tip">
                按上传顺序合并为 PDF;勾选下方「识别图片文字」后,打印版与手写版均可转写为可搜索文字
              </div>
              <div class="el-upload__tip">
                「双层 PDF」会丢弃照片背景,仅用识别出的文字重排成 Word 风格文档(笔记本边缘、桌面、笔等杂物不会保留)
              </div>
            </template>
          </el-upload>

          <el-form label-width="100" style="margin-top: 12px">
            <el-form-item label="识别图片文字">
              <el-switch v-model="pdfImgOcr" />
              <span style="margin-left: 8px; color: #888; font-size: 12px">
                关闭后不做文字识别,速度更快(仅合并图片)
              </span>
            </el-form-item>
            <el-form-item label="PDF 形态">
              <el-radio-group v-model="pdfImgMode" :disabled="!pdfImgOcr">
                <el-radio-button label="dual">双层 PDF</el-radio-button>
                <el-radio-button label="text">纯文字 PDF</el-radio-button>
              </el-radio-group>
            </el-form-item>
            <el-form-item label="形态说明">
              <span style="color: #888; font-size: 12px">
                {{ pdfImgModeTip }}
              </span>
            </el-form-item>
          </el-form>
        </el-tab-pane>
      </el-tabs>

      <!-- 【pdf-tools-history-persistence 新增】保存为历史任务(5 tab 共享) -->
      <div class="pdf-save-row">
        <el-checkbox v-model="pdfSaveToHistory">
          💾 保存为历史任务(可在「历史任务」页查看,带完整报告与下载链接)
        </el-checkbox>
      </div>

      <!-- 结果区(4 个 tab 通用) -->
      <div v-if="pdfResult" class="pdf-result-box">
        <!-- 转 Word 结果 -->
        <template v-if="pdfTab === 'word'">
          <div class="pdf-result-line">✅ Word 文档生成成功</div>
          <div class="pdf-result-line">原始大小:{{ formatSize(pdfResult.original_size) }} → 输出大小:{{ formatSize(pdfResult.output_size) }}</div>
          <el-button type="primary" size="small" @click="onPdfDownload">下载 Word 文件</el-button>
        </template>
        <!-- 转表格结果 -->
        <template v-else-if="pdfTab === 'excel'">
          <div class="pdf-result-line">✅ 识别到 {{ pdfResult.table_count }} 个表格,已生成多 Sheet Excel</div>
          <div class="pdf-result-line" v-if="pdfResult.ai_used">已使用 AI 辅助识别表头</div>
          <div class="pdf-result-line" v-if="pdfResult.dataset_id">已保存为数据集:{{ pdfResult.dataset_name }} (#{{ pdfResult.dataset_id }})</div>
          <el-button type="primary" size="small" @click="onPdfDownload">下载 Excel 文件</el-button>
        </template>
        <!-- 提取文字结果 -->
        <template v-else-if="pdfTab === 'text'">
          <div class="pdf-result-line">
            ✅ 共 {{ pdfResult.page_count }} 页,提取 {{ pdfResult.char_count }} 字
            <span v-if="pdfResult.ai_fallback" style="color:#e6a23c"> (AI 降级为原文)</span>
          </div>
          <el-input
            type="textarea"
            :rows="8"
            :model-value="pdfResult.text"
            readonly
            style="margin-top: 8px"
          />
          <div style="margin-top: 8px">
            <el-button type="primary" size="small" @click="onPdfDownload">下载 TXT 文件</el-button>
          </div>
        </template>
        <!-- 压缩结果 -->
        <template v-else-if="pdfTab === 'compress'">
          <div class="pdf-result-line">✅ 压缩完成</div>
          <div class="pdf-result-line">
            原始大小:{{ formatSize(pdfResult.original_size) }} → 输出大小:{{ formatSize(pdfResult.output_size) }}
          </div>
          <div class="pdf-result-line">
            压缩比:
            <span :style="{ color: pdfResult.compression_ratio > 0 ? '#67c23a' : (pdfResult.compression_ratio < 0 ? '#e6a23c' : '#888') }">
              {{ pdfResult.compression_ratio }}%
            </span>
            <span v-if="pdfResult.compression_ratio <= 0" style="color:#888; margin-left:6px">(该 PDF 已优化,体积变化不大)</span>
          </div>
          <el-button type="primary" size="small" @click="onPdfDownload">下载压缩后的 PDF</el-button>
        </template>

        <!-- 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 结果 -->
        <template v-else-if="pdfTab === 'image'">
          <div class="pdf-result-line">
            ✅ 已将 {{ pdfResult.image_count }} 张照片合并为 {{ pdfResult.mode_label }}({{ pdfResult.page_count }} 页)
          </div>
          <div class="pdf-result-line">
            识别成功:{{ pdfResult.ocr_success }} 张
            <span v-if="pdfResult.ocr_failed > 0" style="color:#e6a23c">
              · 失败 {{ pdfResult.ocr_failed }} 张(已按原图合并,未影响生成)
            </span>
            <span v-else style="color:#67c23a"> · 全部成功</span>
          </div>
          <div class="pdf-result-line">
            产物大小:{{ formatSize(pdfResult.output_size) }}
            <span v-if="!pdfResult.ocr_enabled" style="color:#888">(未启用文字识别)</span>
          </div>

          <!-- 失败明细:让用户知道哪几张没识别出来,便于重试 -->
          <el-alert
            v-if="pdfImgErrorList.length"
            type="warning"
            :closable="false"
            show-icon
            style="margin-top: 8px"
          >
            <template #title>部分图片未识别成功</template>
            <div style="font-size: 12px; line-height: 1.7">
              <div v-for="e in pdfImgErrorList" :key="e.index">第 {{ e.index }} 张:{{ e.reason }}</div>
            </div>
          </el-alert>

          <!-- 识别文字预览 -->
          <el-input
            v-if="pdfResult.ocr_text"
            type="textarea"
            :rows="8"
            :model-value="pdfResult.ocr_text"
            readonly
            style="margin-top: 8px"
          />
          <el-button type="primary" size="small" style="margin-top: 8px" @click="onPdfDownload">
            下载 PDF
          </el-button>
        </template>

        <!-- 【pdf-tools-history-persistence 新增】已保存为历史任务卡片 -->
        <div v-if="pdfResult.task_id" class="pdf-saved-task-card">
          <div class="pdf-saved-task-text">
            📦 已保存为历史任务 <b>#{{ pdfResult.task_id }}</b> · {{ pdfToolLabel }}
          </div>
          <el-button
            type="primary"
            size="small"
            @click="$router.push('/reports/' + pdfResult.task_id)"
          >
            查看完整报告
          </el-button>
        </div>
      </div>

      <template #footer>
        <el-button @click="showPdfTools = false">关闭</el-button>
        <el-button
          type="primary"
          :loading="pdfLoading"
          :disabled="pdfTab === 'image' ? !pdfImgFileList.length : !pdfFile"
          @click="onPdfSubmit"
        >
          开始处理
        </el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, watch, reactive, nextTick } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import {
  Operation, CircleCheck, ChatDotRound, Switch, DataLine, Brush, Document,
  Star, Picture,
} from '@element-plus/icons-vue'
import { use } from 'echarts/core'
import { CanvasRenderer } from 'echarts/renderers'
// 【图表质感修复】原来只注册了 LineChart / PieChart，漏了 BarChart / ScatterChart。
// 后端规则引擎（chart_renderer.build_dashboard_options）主力产出的就是 bar 图，
// 也会产出 scatter —— 漏注册会让这些图在前置看板里渲染成**空白卡片**，
// 而前置看板正是用户挑选图表加入报表的入口，等于整个挑选环节不可用。
// ReportView.vue 早已注册 BarChart，这里补齐保持两处一致。
import { LineChart, PieChart, BarChart, ScatterChart } from 'echarts/charts'
import { TitleComponent, TooltipComponent, LegendComponent, GridComponent } from 'echarts/components'
import VChart from 'vue-echarts'

import { datasetsApi, workbenchApi, toolsApi, tasksApi, downloadBlobResponse } from '@/api'
// 【export-visual-report 新增】离屏 ECharts 高保真截图通道（报表导出主路径）
import { captureChartsPng } from '@/utils/chartExport'
import { useWorkbenchStore } from '@/store/workbench'

use([CanvasRenderer, LineChart, PieChart, BarChart, ScatterChart, TitleComponent, TooltipComponent, LegendComponent, GridComponent])

const wbStore = useWorkbenchStore()
const $router = useRouter()

const datasetOptions = ref([])
const previewColumns = ref([])
const previewRows = ref([])
const previewLoading = ref(false)
// 【workbench-per-sheet-preview 新增】预览区顶部 sheet tab 的双向绑定值;
// 与 store.draft.activeSheet 始终保持一致,初始化取自 store,切换 tab 时
// 由 onPreviewSheetChange 写回 store + 重新拉预览。
const previewSheetTab = ref(wbStore.draft.activeSheet || '')
const metrics = ref({})
const analyzeResult = ref(null)
const analyzing = ref(false)
const analyzeStep = ref(0)

const showMerge = ref(false)
const showValidate = ref(false)
const showNl = ref(false)
const showCompare = ref(false)
const showChart = ref(false)
const toolLoading = ref(false)
const validateResult = ref(null)
const showRawJson = ref(false)

// 「原始数据表格」用的 computed:从 validateResult.raw_preview 解出来,
// 兼容字段缺失/为 null 的旧缓存结果。
const rawPreviewRows = computed(() => validateResult.value?.raw_preview?.rows || [])
const rawPreviewColumns = computed(() => validateResult.value?.raw_preview?.columns || [])
const rawPreviewTotal = computed(() => validateResult.value?.raw_preview?.total || 0)

// 校验弹窗里 AI 解读的等级 / 严重度中文标签
const _LEVEL_LABELS = { good: '良好', warn: '需关注', danger: '较差' }
const _SEV_LABELS = { low: '轻微', medium: '中等', high: '严重' }
function levelLabel(l) { return _LEVEL_LABELS[l] || l || '需关注' }
function sevLabel(s) { return _SEV_LABELS[s] || s || '轻微' }
const nlForm = ref({ question: '' })
// ===== [tools-results-history-persistence] 3 个工具的[保存为历史任务]勾选与已保存卡片 =====
// 勾选默认开启;latest*TaskId 关闭弹窗不重置,下次打开仍能看到上次保存的 task
const saveAsValidateTask = ref(true)
const saveAsNlTask = ref(true)
const saveAsCompareTask = ref(true)
const latestValidateTaskId = ref(null)
const latestNlTaskId = ref(null)
const latestCompareTaskId = ref(null)
const nlResult = ref(null)
// ===== 【nl-chat-upgrade 新增】数据对话多轮气泡流状态 =====
// chatMessages:本次弹窗会话的消息流,每项 { role:'user'|'ai', content, structured?, rows?, columns? }
// - user 气泡只存问题文本;ai 气泡 structured 为后端结构化结果(为 null 时降级用 rows/columns 旧表格)
const chatMessages = ref([])
// 消息列表 DOM 引用:新消息入列后滚动到底部用
const nlChatListRef = ref(null)
// 空状态示例问题:点击即作为新问题发送
const nlExampleQuestions = [
  '今年哪个月业绩最低？',
  '销售额前 5 的产品是哪些？',
  '按地区统计订单量并从高到低排序',
  '找出数值异常偏高的记录',
]
// 多表合并表单:mode='flat' = 一张大表(旧行为);mode='multi_sheet' = 多 sheet 同一份 xlsx
const mergeForm = ref({ name: '合并数据集', mode: 'flat', dataset_ids: [] })
const compareForm = ref({
  left_dataset_id: null,
  right_dataset_id: null,
  mode: 'business',
  group_col: '',
  metric_col: '',
  key_cols: [],
  // 【topn 并入 business 新增】TopN 截断开关:
  // use_top_n 关闭 → 后端返回全部分组;开启 → 只保留指标值最高的 top_n 个分组
  use_top_n: false,
  top_n: 10,
  // 【dataset-contract-compare 新增】合同对比场景的"重点关注"软提示
  focus: '',
})
const compareResult = ref(null)

// 【dataset-contract-compare 新增】合同对比结果辅助:
// 1) 高优先级变更数(展示在标题徽标)
// 2) 变更类型 → 标签视觉(modified/added/removed/unchanged/unclear)
// 3) 表格行高亮:高优先级变更行加底色
const contractHighChangeCount = computed(() => {
  const details = compareResult.value?.diff_details
  if (!Array.isArray(details)) return 0
  return details.filter(
    (c) => c.importance === 'high' && ['modified', 'added', 'removed'].includes(c.change_type),
  ).length
})
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
function contractRowClass({ row }) {
  if (row?.importance === 'high' && ['modified', 'added', 'removed'].includes(row.change_type)) {
    return 'contract-row-high'
  }
  return ''
}
// mode 切换时清掉旧 result,避免不同模式的旧结果渲染混乱
function onCompareModeChange() {
  compareResult.value = null
}

// 图表编辑相关
const chartInstruction = ref('')
const chartEditLoading = ref(false)
const chartEditResult = ref(null)
const QUICK_CHART_INSTRUCTIONS = [
  '换成折线图',
  '换成饼图',
  '配色换成蓝绿色',
  '只显示 Top 5',
  '改成横向柱状图',
  '加一条同比增长折线',
]

// ===== 【chart-edit-prebuilt-dashboard 新增】前置看板相关 state =====
const chartStyleOptions = ref([])      // 后端 /workbench/charts/styles 返回的 3 套风格
const chartPreviewLoading = ref(false) // 「生成新图表」按钮 loading
// 当前工作台模式下,看板里被「指令改图」选中的图表索引(用于"替换到看板")
const chartEditTargetIdx = ref(-1)

// 数据源
const chartSource = ref('workbench')  // 'workbench' | 'history'
const chartHistoryTaskId = ref(null)
const chartHistoryIndex = ref(0)
const historyTaskOptions = ref([])
const historyChartOptions = ref([])
// 【修复】改成 reactive,让 activeDatasetId 等 computed 在赋值后能被精确触发;
//  原裸对象 {} 在 Vue 3 下也能在读写,但在某些边界场景下 computed 重算不可靠,
//  改 reactive 后响应式链路完整且零成本。后续所有 historyTaskCache[taskId] = xxx 都不变。
const historyTaskCache = reactive({})

// 当前工作台的图表(保留兼容旧 analyze 流程的右侧 AI 结果区)
const currentChart = computed(() => {
  const charts = analyzeResult.value?.chart_option?.charts || []
  return charts[0] || null
})

// 当前激活的图表(根据数据源决定取哪个)
const activeChart = computed(() => {
  if (chartSource.value === 'workbench') {
    // 工作台模式下,改图对象 = 看板里被点中的那张(否则为最后一张)
    const charts = wbStore.draft.dashboardCharts
    if (!charts.length) return null
    if (chartEditTargetIdx.value >= 0 && chartEditTargetIdx.value < charts.length) {
      const c = charts[chartEditTargetIdx.value]
      return { ...c, type: c.option?.series?.[0]?.type || 'bar' }
    }
    return charts[charts.length - 1]
      ? { ...charts[charts.length - 1], type: charts[charts.length - 1].option?.series?.[0]?.type || 'bar' }
      : null
  }
  if (chartHistoryIndex.value == null) return null
  return historyChartOptions.value[chartHistoryIndex.value] || null
})

// 当前激活的数据集 id(用于调用后端)
const activeDatasetId = computed(() => {
  if (chartSource.value === 'workbench') return wbStore.draft.datasetId
  // 【修复】history 模式三层兜底,确保按钮永远不"点不动":
  //   1) 优先从 detail.dataset_id 取(常规路径,后端 get_task 已保证返回)
  //   2) 兜底 #1:detail.id → 但后端返回的是 task id 不是 dataset id,这种数据不对
  //      所以改成:detail.dataset_id 严格为数字才采纳
  //   3) 兜底 #2:回落到 wbStore.draft.datasetId
  //      (用户在工作台已经选了数据集时,历史任务和该数据集通常对应同一源,
  //       退一步给个值,至少能让按钮响应,真不匹配也由后端 404 接住)
  const t = historyTaskCache[chartHistoryTaskId.value]
  if (t && (typeof t.dataset_id === 'number' || typeof t.dataset_id === 'string')) {
    return t.dataset_id
  }
  return wbStore.draft.datasetId || null
})

// 列缓存: { datasetId: [col1, col2, ...] } 用于弹窗内的下拉推断
const compareColumnsCache = ref({})
async function loadDatasetColumns(id) {
  if (!id || compareColumnsCache.value[id]) return
  try {
    const preview = await datasetsApi.preview(id)
    compareColumnsCache.value = {
      ...compareColumnsCache.value,
      [id]: preview?.columns || [],
    }
  } catch (e) {
    compareColumnsCache.value = { ...compareColumnsCache.value, [id]: [] }
  }
}
function onCompareDatasetChange() {
  compareResult.value = null
  const cols = compareColumnsOptions.value || []
  if (!compareForm.value.group_col && cols.length) {
    compareForm.value.group_col = cols[0]
  }
}
const compareColumnsOptions = computed(() => {
  const l = compareColumnsCache.value[compareForm.value.left_dataset_id] || []
  const r = compareColumnsCache.value[compareForm.value.right_dataset_id] || []
  return Array.from(new Set([...l, ...r]))
})

// 【dataset-multiformat-upload 新增】根据左右数据集的文件类型,判定走哪种对比链路
// - 双 structured → 走现有 Pandas 2 种 mode(business / row_diff)
// - 任一非 structured → 走 LLM 跨模态对比,Pandas 参数全部隐藏
// 用计算属性是因 datasetOptions 会随上传刷新,这里实时跟随
function getDatasetFileType(id) {
  if (!id) return 'structured'
  const d = (datasetOptions.value || []).find((x) => x.id === id)
  return d?.file_type || 'structured'
}
const isDocumentCompare = computed(() => {
  const lt = getDatasetFileType(compareForm.value.left_dataset_id)
  const rt = getDatasetFileType(compareForm.value.right_dataset_id)
  return lt !== 'structured' || rt !== 'structured'
})
// 文档对比时,两边的文件类型 + 类型标签(用于提示卡)
const compareTypePair = computed(() => {
  const lt = getDatasetFileType(compareForm.value.left_dataset_id)
  const rt = getDatasetFileType(compareForm.value.right_dataset_id)
  return { left: lt, right: rt }
})
// 文件类型 → 中文标签(沿用 DatasetsView 的视觉规则,保持一致)
const _FT_LABEL = { structured: '表格', document: '文档', image: '图片' }
function fileTypeLabel(t) { return _FT_LABEL[t] || '未知' }

// ===== 对比结果的数值格式化 =====
// 说明:这两个函数此前在模板中被调用但项目里没有定义,
// 导致业务对比表格的「差值 / 变化幅度」列既没有正负号也没有涨跌配色。
// 这里补齐:与 row_diff 明细里的 cell-before / cell-after 配色保持同一套语义
// (涨绿跌红,与变更追踪明细里的 cell-before / cell-after 同色系)。
function formatSigned(v) {
  if (v === null || v === undefined || v === '') return '—'
  const n = Number(v)
  if (!Number.isFinite(n)) return '—'
  // 正数显式带 "+",负数由 toFixed 自带负号
  return (n > 0 ? '+' : '') + n.toFixed(2)
}

// 差值 / 变化幅度的涨跌配色:涨=绿、跌=红、持平=灰
function diffClass(v) {
  if (v === null || v === undefined || v === '') return 'diff-flat'
  const n = Number(v)
  if (!Number.isFinite(n)) return 'diff-flat'
  if (n > 0) return 'diff-up'
  if (n < 0) return 'diff-down'
  return 'diff-flat'
}

// 数据清洗弹窗专用 ref
const showClean = ref(false)
const cleanLoading = ref(false)
const cleanForm = ref({ dataset_id: null, name: '' })

// ===== 【clean-ai-derived-columns-pivot 新增】派生列 / 透视汇总辅助逻辑 =====
// 清洗弹窗内所选数据集的所有工作表名(派生列工作表下拉、透视源工作表下拉共用)
const cleanSheetNames = computed(() =>
  (wbStore.draft.cleanSchema || []).map((s) => s.name).filter(Boolean),
)
// ===== 【透视多张并行改造】把原 pivotSheetColumns/pivotRowOptions/pivotValueOptions 改成函数,按 sheet_name 取候选 =====
// 取指定工作表的列结构
const pivotSheetColumnsForSheet = (sheetName) => {
  if (!sheetName) return []
  const sheet = (wbStore.draft.cleanSchema || []).find((s) => s.name === sheetName)
  return sheet?.columns || []
}
// 行维度候选:优先非数值列与低基数(unique_count<=50)列,数值 / 高基数列排后
const pivotRowOptionsForSheet = (sheetName) => {
  const cols = pivotSheetColumnsForSheet(sheetName)
  const score = (c) => {
    let s = 0
    if (c.type === 'number') s += 2
    if ((c.unique_count ?? 0) > 50) s += 1
    return s
  }
  return [...cols].sort((a, b) => score(a) - score(b))
}
// 汇总字段候选:数值列 + 本 sheet 已勾选/手动的派生列
// 派生列是在清洗阶段追加的,不在 cleanSheet 的原始列结构里,
// 但后端执行顺序是 clean → 派生列 → 透视,因此它们同样可作为汇总字段。
const pivotValueOptionsForSheet = (sheetName) => {
  if (!sheetName) return []
  const base = pivotSheetColumnsForSheet(sheetName).filter((c) => c.type === 'number')
  const derived = [
    ...wbStore.draft.derivedSuggestions.filter(
      (s) => s.selected && s.sheet_name === sheetName && (s.name || '').trim(),
    ),
    ...wbStore.draft.manualDerivedColumns.filter(
      (s) => s.sheet_name === sheetName && (s.name || '').trim(),
    ),
  ]
  const seen = new Set(base.map((c) => c.name))
  const extra = []
  for (const d of derived) {
    const name = d.name.trim()
    // 与原始数值列同名时以后端校验为准(会 400),这里不重复展示
    if (seen.has(name)) continue
    seen.add(name)
    extra.push({ name, type: 'number', from_derived: true })
  }
  return [...base, ...extra]
}
// ===== 透视汇总区块的折叠状态 =====
const pivotRecommendCollapsed = ref(false)

// 兼容旧模板中 pivotSheetColumns 引用(列维度下拉用)——用首张已应用配置的工作表,无则空
const pivotSheetColumns = computed(() => {
  const first = wbStore.draft.pivotConfigs?.[0]
  return first ? pivotSheetColumnsForSheet(first.sheet_name) : []
})
// 预览区兜底:旧响应没有 derived_columns_created 字段时按 0 处理
const cleanPreviewDerivedCount = computed(() => {
  const created = wbStore.draft.previewCleanedStats?.derived_columns_created
  return Array.isArray(created) ? created.length : 0
})
// 预览区兜底:旧响应没有 derived_columns_warnings 字段时按空数组处理
const cleanPreviewWarnings = computed(() => {
  const w = wbStore.draft.previewCleanedStats?.derived_columns_warnings
  return Array.isArray(w) ? w.filter(Boolean) : []
})

// ===== 【透视多张并行改造】去掉 enabled 字段,数组为空即表示不启用 =====
const _emptyPivotConfig = () => ({
  id: `pivot_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
  sheet_name: '',
  row_dim: '',
  col_dim: '',
  value: '',
  agg: 'sum',
  include_margin: true,
})

/**
 * 拉取所选数据集的字段结构并写入 store.cleanSchema。
 * 失败时退化为空数组(派生列 / 透视下拉为空),不阻塞弹窗其余操作。
 */
async function loadCleanSchema(datasetId) {
  if (!datasetId) {
    wbStore.draft.cleanSchema = []
    return
  }
  try {
    const r = await toolsApi.schema(datasetId)
    wbStore.draft.cleanSchema = Array.isArray(r?.sheets) ? r.sheets : []
  } catch (e) {
    wbStore.draft.cleanSchema = []
    console.error('load clean schema failed:', e)
  }
}

// 切换清洗数据集:重置派生列 / 透视配置 / 上次预览,再重新拉 schema
watch(
  () => cleanForm.value.dataset_id,
  (id) => {
    wbStore.draft.derivedSuggestions = []
    wbStore.draft.manualDerivedColumns = []
    wbStore.draft.pivotConfigs = []   // 【透视多张并行改造】清空数组
    wbStore.draft.previewCleanedStats = null
    loadCleanSchema(id)
  },
)

/**
 * 请求 AI 推荐派生列。成功写入 store(默认全选);
 * 兜底(fallback=true 且空列表)与异常都只提示、不阻塞弹窗。
 */
async function onSuggestColumns() {
  const datasetId = cleanForm.value.dataset_id
  if (!datasetId || wbStore.draft.suggestColumnsLoading) return
  wbStore.draft.suggestColumnsLoading = true
  try {
    const r = await toolsApi.suggestColumns({ dataset_id: datasetId })
    const list = Array.isArray(r?.suggestions) ? r.suggestions : []
    wbStore.draft.derivedSuggestions = list.map((s) => ({ ...s, selected: true }))
    // 同时拿到 AI 推荐的透视汇总方案,让用户"选一条应用"而不是自己从零配
    const pivotList = Array.isArray(r?.pivot_suggestions) ? r.pivot_suggestions : []
    wbStore.draft.pivotSuggestions = pivotList
    if (r?.fallback && !list.length && !pivotList.length) {
      ElMessage.warning('AI 推荐暂不可用,可手动添加派生列或手动配置透视')
    }
  } catch (e) {
    ElMessage.warning('AI 推荐暂不可用,可手动添加派生列或手动配置透视')
    console.error('suggestColumns failed:', e)
  } finally {
    wbStore.draft.suggestColumnsLoading = false
  }
}

/** 【透视多张并行改造】把一条 AI 推荐的透视方案 push 到 pivotConfigs 数组 */
function applyPivotSuggestion(item) {
  const newConfig = {
    ..._emptyPivotConfig(),
    sheet_name: item.sheet_name || '',
    row_dim: item.row_dim || '',
    col_dim: item.col_dim || '',
    value: item.value || '',
    agg: item.agg || 'sum',
    include_margin: item.include_margin !== false,
  }
  wbStore.draft.pivotConfigs.push(newConfig)
  // 应用后不隐藏推荐列表,允许用户接着应用下一条或微调已应用的
  ElementPlus.ElMessage.success(
    `已应用透视方案:${item.row_dim}${item.col_dim ? '×' + item.col_dim : ''} › ${item.value}`,
  )
}

/** 【透视多张并行改造】手动添加一张空白透视配置 */
function addManualPivot() {
  wbStore.draft.pivotConfigs.push(_emptyPivotConfig())
}

/** 【透视多张并行改造】删除指定透视配置 */
function removePivotConfig(id) {
  const idx = wbStore.draft.pivotConfigs.findIndex((p) => p.id === id)
  if (idx >= 0) wbStore.draft.pivotConfigs.splice(idx, 1)
}

/** 手动添加一行派生列:工作表默认取首张已应用透视的源表 → 首张表 → 空 */
function addManualDerived() {
  const sheet =
    wbStore.draft.pivotConfigs?.[0]?.sheet_name || cleanSheetNames.value[0] || ''
  wbStore.draft.manualDerivedColumns.push({ sheet_name: sheet, name: '', expression: '' })
}

/** 删除指定手动派生列行 */
function removeManualDerived(index) {
  wbStore.draft.manualDerivedColumns.splice(index, 1)
}

/** 【透视多张并行改造】透视源工作表切换:清空该行的行维度 / 列维度 / 汇总字段,避免残留旧表字段名 */
function onPivotSheetChange(id) {
  const p = wbStore.draft.pivotConfigs.find((c) => c.id === id)
  if (!p) return
  p.row_dim = ''
  p.col_dim = ''
  p.value = ''
}

// 【workbench-ai-pdf-tools 新增】AI PDF 处理弹窗专用 ref
// 5 个 tab 复用同一组状态,弹窗打开时重置。
const showPdfTools = ref(false)
const pdfTab = ref('word')                 // word | excel | text | compress | image
const pdfLoading = ref(false)
const pdfResult = ref(null)                // 后端响应对象
const pdfFile = ref(null)                  // 当前选中的 File 对象
const pdfAiEnabled = ref(true)             // 是否启用 AI 增强
const pdfSaveAsDataset = ref(false)        // 转表格时是否落为新数据集
const pdfSaveToHistory = ref(true)         // 【pdf-tools-history-persistence 新增】是否落为历史任务,默认勾选
// 单 tab 文件大小上限(MB),与后端 _MAX_PDF_SIZE 对齐
const PDF_MAX_SIZE_MB = 50

// ===== 【pdf-photo-to-pdf-ocr 新增】照片转 PDF(第 5 个 tab)专用状态 =====
const pdfImgFileList = ref([])             // el-upload 托管的文件列表(含 raw File 对象)
const pdfImgMode = ref('dual')             // dual(双层) | text(纯文字)
const pdfImgOcr = ref(true)                // 是否做文字识别,默认开启
const PDF_MAX_IMG_SIZE_MB = 20             // 单张上限,与后端 _MAX_IMAGE_SIZE 对齐
const PDF_MAX_IMG_COUNT = 20               // 张数上限,与后端 _MAX_IMAGE_COUNT 对齐

// 形态说明文案:把两种 PDF 的差异讲清楚,避免用户选错
const pdfImgModeTip = computed(() => {
  if (!pdfImgOcr.value) return '未开启识别:仅按顺序把照片合并为 PDF,速度最快,不消耗识别额度'
  return {
    dual: '双层 PDF:丢弃照片背景,用识别出的文字重新排版成 Word 风格文档 —— 纯白底黑字、文字正向横排、字号行距统一、首行缩进;甲方/乙方等左右并排的两栏会保持同一行,印章按它在原图上的左右位置贴到所属那一行旁边',
    text: '纯文字 PDF:只输出识别文字并重新排版,体积最小,不保留印章图像',
  }[pdfImgMode.value] || ''
})

// 识别失败明细:把后端返回的 ocr_errors 数组转成 { index, reason } 便于 v-for 展示
const pdfImgErrorList = computed(() => {
  const errs = pdfResult.value?.ocr_errors || []
  return errs
    .map((reason, i) => (reason ? { index: i + 1, reason } : null))
    .filter(Boolean)
})

// 当前 tab 对应的中文工具名(已保存卡片展示用)
const pdfToolLabel = computed(() => {
  return {
    word: 'PDF 转 Word',
    excel: 'PDF 转表格',
    text: 'PDF 提取文字',
    compress: 'PDF 压缩',
    image: '照片转 PDF',   // 【pdf-photo-to-pdf-ocr 新增】
  }[pdfTab.value] || 'PDF 处理'
})

// 图表编辑弹窗打开时加载历史任务
watch(showChart, (v) => {
  if (v) {
    loadHistoryTasks()  // 每次打开都拉,避免后端列表变化没刷到
    loadChartStyles()   // 【chart-edit-prebuilt-dashboard 新增】拉 3 套风格
  }
})

// 兜底:如果监听没生效(初始化时机问题),直接在 chartSource 切到 history 时也拉一次
watch(chartSource, (v) => {
  if (v === 'history' && historyTaskOptions.value.length === 0) loadHistoryTasks()
})

// ===== 【chart-edit-prebuilt-dashboard 新增】 =====
// 拉 3 套图表风格(供 el-option 下拉)
async function loadChartStyles() {
  if (chartStyleOptions.value.length) return  // 已有就跳过
  try {
    const r = await workbenchApi.listChartStyles()
    chartStyleOptions.value = r?.styles || []
  } catch (e) {
    // 兜底:后端挂了也能用本地 3 套
    chartStyleOptions.value = [
      { key: 'tech_advanced', label: '高级科技风格' },
      { key: 'business_pro', label: '商务职业风格' },
      { key: 'minimalist', label: '极简汇风格' },
    ]
  }
}
// 「生成新图表」整套重新生成(覆盖整个看板,按 sheet 分组)
// 【chart-render-smart-2026-10 新增】AI 智能筛选开关切换时:
  //   - 打开 → 提示用户"将走 LLM 精选,首次约 3~6 秒",让用户有预期
  //   - 关闭 → 清空旧看板,避免遗留的"AI 精选"结果与新模式不匹配
  function onAiCurateToggle(val) {
    if (val) {
      ElMessage.info('AI 智能筛选已开启:本地出骨架 + LLM 精选 TopN + 标题润色。首次约 3~6 秒。')
    } else {
      wbStore.setDashboard({ groups: [], charts: [] }, false)
    }
  }
async function onRegenerateDashboard() {
  if (!wbStore.draft.datasetId) {
    ElMessage.warning('请先选择数据集')
    return
  }
  chartPreviewLoading.value = true
  try {
    // 【chart-render-smart-2026-10 新增】根据 wbStore.draft.aiCurate 决定
    // 是走本地规则出图,还是走"混合出图"(本地骨架 + LLM 精选 + 标题润色)。
    const aiCurate = !!wbStore.draft.aiCurate
    const r = await workbenchApi.previewCharts(
      wbStore.draft.datasetId,
      wbStore.draft.chartStyle || 'tech_advanced',
      {
        aiCurate,
        targetCount: Number(wbStore.draft.aiCurateTarget) || 3,
      },
    )
    // 新 API 返回 {sheets: [{sheet, charts}], sheet_count, chart_count}
    // 旧 API 返回 {charts: [...]} — 都兼容
    if (Array.isArray(r?.sheets) && r.sheets.length) {
      wbStore.setDashboard({ groups: r.sheets }, wbStore.draft.keepSelectionOnRegen)
    } else if (Array.isArray(r?.charts) && r.charts.length) {
      wbStore.setDashboard({ charts: r.charts }, wbStore.draft.keepSelectionOnRegen)
    } else {
      ElMessage.warning('数据集字段不足,未生成图表')
      return
    }
    // 成功提示:勾选 AI 智能筛选时附上 LLM 精选标识,让用户知道走了哪条路径
    const tag = r.ai_curated ? '(AI 精选)' : '(本地规则)'
    ElMessage.success(`已生成 ${r.chart_count || 0} 张图表 ${tag} · 覆盖 ${r.sheet_count || 1} 个工作表`)
  } catch (e) {
    // axios 拦截器已经提示
  } finally {
    chartPreviewLoading.value = false
  }
}
// 拍平所有分组为单一数组(供下载、改图 activeChart 等使用)
const dashboardGroups = computed(() => wbStore.draft.dashboardGroups || [])
// 单图下载 PNG(ECharts getDataURL)
function downloadSingleChart(chart, key) {
  try {
    // 用 key 在 DOM 里定位卡片(每个 .chart-card 上的 data-chart-key 属性)
    const el = document.querySelector(`[data-chart-key="${CSS.escape(String(key))}"]`)
    if (!el) {
      ElMessage.warning('未找到对应图表元素')
      return
    }
    const canvas = el.querySelector('canvas')
    if (!canvas) {
      ElMessage.warning('图表尚未渲染完成,请稍后再试')
      return
    }
    const dataUrl = canvas.toDataURL('image/png')
    const a = document.createElement('a')
    a.href = dataUrl
    const safeTitle = (chart.title || `chart`).replace(/[\\/:*?"<>|]/g, '_')
    a.download = `${safeTitle}_${wbStore.draft.chartStyle}.png`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
  } catch (e) {
    ElMessage.error('下载失败:' + (e?.message || e))
  }
}
// 下载整个看板(按 sheet 分组纵向拼接)
async function downloadWholeDashboard() {
  try {
    // 收集所有 sheet 的卡片(每个 sheet 一个 section)
    const groups = Array.from(document.querySelectorAll('.dashboard-group'))
    if (!groups.length) {
      ElMessage.warning('看板未渲染')
      return
    }
    const blocks = []
    for (const g of groups) {
      const title = g.querySelector('.dashboard-group-title')?.textContent || ''
      const canvases = Array.from(g.querySelectorAll('.chart-card canvas'))
      if (!canvases.length) continue
      // 每张 sheet 内部 2 列网格排版
      const maxW = Math.max(...canvases.map((c) => c.width))
      const cellW = maxW / 2 - 8
      const cellH = Math.max(...canvases.map((c) => c.height)) * (cellW / maxW)
      const cols = 2
      const rows = Math.ceil(canvases.length / cols)
      const blockW = cellW * cols + 32
      const blockH = cellH * rows + 56 + 24
      const cnv = document.createElement('canvas')
      cnv.width = blockW
      cnv.height = blockH
      const ctx = cnv.getContext('2d')
      ctx.fillStyle = '#0a1a3a'
      ctx.fillRect(0, 0, blockW, blockH)
      ctx.fillStyle = '#2dd4bf'
      ctx.font = 'bold 16px "PingFang SC", sans-serif'
      ctx.fillText(title, 16, 28)
      // 把每个 canvas 内容画到 block
      await Promise.all(canvases.map(async (c, i) => {
        const img = await new Promise((resolve, reject) => {
          const im = new Image()
          im.onload = () => resolve(im)
          im.onerror = reject
          im.src = c.toDataURL('image/png')
        })
        const r = Math.floor(i / cols)
        const col = i % cols
        const x = 16 + col * (cellW + 16)
        const y = 56 + r * (cellH + 16)
        ctx.drawImage(img, x, y, cellW, cellH)
      }))
      blocks.push({ title, img: cnv })
    }
    if (!blocks.length) {
      ElMessage.warning('看板为空')
      return
    }
    // 把所有 block 拍成一张大图
    const maxW = Math.max(...blocks.map((b) => b.img.width))
    const totalH = blocks.reduce((s, b) => s + b.img.height + 16, 0) + 16
    const out = document.createElement('canvas')
    out.width = maxW
    out.height = totalH
    const ctx = out.getContext('2d')
    ctx.fillStyle = '#0a1a3a'
    ctx.fillRect(0, 0, out.width, out.height)
    let y = 16
    for (const b of blocks) {
      ctx.drawImage(b.img, (maxW - b.img.width) / 2, y)
      y += b.img.height + 16
    }
    const dataUrl = out.toDataURL('image/png')
    const a = document.createElement('a')
    a.href = dataUrl
    a.download = `dashboard_${wbStore.draft.chartStyle}_${Date.now()}.png`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
  } catch (e) {
    ElMessage.error('整看板下载失败:' + (e?.message || e))
  }
}

// ============================================================
// 【export-visual-report 新增】导出 Excel 可视化报表（弹窗入口）
// ------------------------------------------------------------
// 与「下载整个看板」的本质区别：
//   看板 = 纯位图长图，不可编辑、不可筛选、无 KPI、无结论、无明细数据；
//   报表 = 结构化 xlsx，含 KPI 指标卡 + AI 结论 + 数据明细 + 图表洞察注释。
//
// 数据流（为什么这么设计）：
//   1) 前端用**离屏 ECharts 实例**把已勾选的图表截成高保真 PNG（所见即所得），
//      走 chartExport.captureChartsPng()，**不改动页面上任何图表实例**，无副作用。
//   2) 只把 dataset_id / style_key / charts(含 option 与 png_base64) 提交给后端；
//      KPI 指标与明细数据由后端按 dataset_id **重新读取数据源实算**，
//      避免前端编造数字（后端有既有的防编造/溯源体系）。
//   3) 截图失败不会阻断导出：该项 png_base64 为 null，后端会用 matplotlib 兜底渲染。
// ============================================================
const exportingReport = ref(false)

async function exportVisualReportFromDashboard() {
  if (exportingReport.value) return
  const selected = wbStore.getSelectedCharts() // [{title, option, style_key, sheet, type}]
  if (!selected.length) {
    ElMessage.warning('请先勾选要加入报表的图表')
    return
  }
  exportingReport.value = true
  try {
    // 1) 逐张离屏截图（串行；单张失败不影响整批）
    const shots = await captureChartsPng(
      selected.map((c) => ({ title: c.title, option: c.option }))
    )
    // 2) 把截图结果按原数组下标回填到图表条目上
    const charts = selected.map((c, i) => ({
      title: c.title || '',
      option: c.option || {},
      sheet: c.sheet || '',
      // png_base64 为纯 base64（无 dataURL 前缀）；失败即 null → 后端兜底
      png_base64: shots[i]?.png_base64 || null,
    }))
    // 3) 提交后端生成 xlsx
    const resp = await workbenchApi.exportVisualReport({
      dataset_id: wbStore.draft.datasetId,
      style_key: wbStore.draft.chartStyle,
      include_ai: true,
      charts,
    })
    // 4) 用通用 blob 工具落盘（会读 Content-Disposition 里的真实文件名）
    const filename = await downloadBlobResponse(resp, 'visual_report.xlsx')
    ElMessage.success(`已导出可视化报表（${charts.length} 张图）· ${filename || ''}`)
  } catch (e) {
    ElMessage.error('导出报表失败：' + (e?.message || e))
  } finally {
    exportingReport.value = false
  }
}

const trendOption = computed(() => {
  const charts = analyzeResult.value?.chart_option?.charts || []
  const line = charts.find((c) => c.type === 'line')
  return line?.option
})
const pieOption = computed(() => {
  const charts = analyzeResult.value?.chart_option?.charts || []
  const pie = charts.find((c) => c.type === 'pie')
  return pie?.option
})

function sectionLabel(key) {
  return { overview: '概况', highlights: '亮点', problems: '问题', suggestions: '建议',
           本季概况: '本季概况', 关键变化: '关键变化', 风险与机会: '风险与机会', 下季计划: '下季计划',
           岗位与职责: '岗位与职责', 关键成果: '关键成果', 不足与反思: '不足与反思', 未来计划: '未来计划',
           核心数字: '核心数字', 关键趋势: '关键趋势', 行动建议: '行动建议' }[key] || key
}

async function loadDatasets() {
  const data = await datasetsApi.list()
  // 【dataset-deleted-not-selectable 修改】后端已在查询层过滤掉内部锚点 dataset,
  // 接口返回的 items 即"用户可见的数据集"全集 —— 与「数据集」页完全同源,
  // 因此工作台的对比/合并/清洗下拉不会再出现"数据集页里已经删掉还能选"的历史残留。
  datasetOptions.value = data.items || []

  // 【dataset-deleted-not-selectable 新增】清理"已被删除"的当前选中项:
  // store 是跨路由存活的单例,用户在「数据集」页删掉当前选中的数据集后,
  // draft.datasetId 仍指向那个已不存在的 id。这里用最新列表做一次校验,
  // 发现已失效就清空,避免后续预览 404 弹红字、或把死 id 提交给后端。
  // (清空后 watch(datasetId) 会自动触发 loadPreview 走 null 分支重置预览区)
  const curId = wbStore.draft.datasetId
  if (curId && !datasetOptions.value.some((d) => d.id === curId)) {
    wbStore.clearDataset()
    ElMessage.warning('当前选中的数据集已被删除，请重新选择数据集')
  }
}

async function loadPreview() {
  // 守卫:从首页/顶部"进入工作台"按钮直接跳过来时,store.draft.datasetId 是 null,
  // 此时调用 /datasets/{id}/preview 会触发"数据集不存在"的红色错误提示,
  // 这里直接 return,保持空工作台即可,避免误报。
  if (!wbStore.draft.datasetId) {
    previewColumns.value = []
    previewRows.value = []
    metrics.value = {}
    analyzeStep.value = 0
    return
  }
  previewLoading.value = true
  try {
    // 【workbench-per-sheet-preview 新增】按当前激活的 sheet 拉取预览。
    //  - activeSheet=<名> → 带 sheet=<名>,后端只返该 sheet 自身的前 10 行 + 真实缺失值;
    //  - activeSheet=''(默认) → 不带 sheet 参数,后端会自动返回第一张 sheet(改版后)。
    const data = await datasetsApi.preview(wbStore.draft.datasetId, {
      sheet: wbStore.draft.activeSheet || undefined,
    })
    previewColumns.value = data.columns
    previewRows.value = data.rows.slice(0, 8)
    metrics.value = data.quality
    // 【dataset-multisheet-merge 新增】把后端返回的 sheets 写入 store,
    // 让顶部数据集卡能展示"📑 包含 N 个工作表"标签。
    // 兼容老接口(没 sheets 字段)→ 传 undefined → store 兜底为 []
    wbStore.setDatasetSheets(data.sheets)
    // 【workbench-per-sheet-preview 新增】默认落到第一张 sheet:
    // 若 store.activeSheet 仍为空(初次进入或换数据集),用后端返回的 sheets 列表
    // 第一项兜底,避免 tab 留空。
    const active = wbStore.draft.activeSheet || data.active_sheet || (data.sheets && data.sheets[0] && data.sheets[0].name) || ''
    if (active !== wbStore.draft.activeSheet) {
      wbStore.setActiveSheet(active)
    }
    previewSheetTab.value = active
    analyzeStep.value = 0
  } catch (e) {
    // http.js 已经统一 ElMessage.error,这里只清空预览,不让页面卡死
    console.error('[workbench] loadPreview failed:', e)
    previewColumns.value = []
    previewRows.value = []
    metrics.value = {}
  } finally {
    previewLoading.value = false
  }
}

// 【workbench-per-sheet-preview 新增】切换预览 sheet:
//   - 点 tab 选具体 sheet → 写 store,重拉预览
//   - 同步本地 previewSheetTab,避免 store 与本地 ref 错位
async function onPreviewSheetChange(name) {
  const target = name || ''
  previewSheetTab.value = target
  wbStore.setActiveSheet(target)
  await loadPreview()
}

// 【workbench-per-sheet-preview 新增】watch store.activeSheet:
// 切换数据集(setDataset)会让 store.activeSheet 重置,这里同步本地 tab。
watch(() => wbStore.draft.activeSheet, (v) => {
  if (previewSheetTab.value !== (v || '')) {
    previewSheetTab.value = v || ''
  }
})

// 【dataset-multisheet-merge 新增】把 sheets 列表拼成 Tooltip 内容,
// 每行一张表:`name: 行数 行 · 列数 列`。空 sheets 时返回空串(配合 v-if 使用)
const fileListMatrixTooltip = computed(() => {
  const sheets = wbStore.draft.datasetSheets || []
  if (!sheets.length) return ''
  return sheets
    .map((s) => `${s.name}: ${s.rows} 行 · ${s.cols} 列`)
    .join('\n')
})

watch(() => wbStore.draft.datasetId, loadPreview)
watch(() => wbStore.draft.templateId, () => { /* 模板变更时清空结果 */ })

// 打开「数据清洗」弹窗时,默认回填左侧当前数据集(若已选);不覆盖用户已选值
watch(showClean, (v) => {
  if (v && !cleanForm.value.dataset_id && wbStore.draft.datasetId) {
    cleanForm.value.dataset_id = wbStore.draft.datasetId
  }
})

async function onAnalyze() {
  analyzing.value = true
  analyzeResult.value = null
  try {
    analyzeStep.value = 1
    await new Promise((r) => setTimeout(r, 200))
    analyzeStep.value = 2
    const resp = await workbenchApi.analyze({
      dataset_id: wbStore.draft.datasetId,
      template_id: wbStore.draft.templateId,
      cleaning_rules: wbStore.draft.cleaningRules,
      name: `${wbStore.draft.datasetName} - ${wbStore.draft.templateName}`,
    })
    // 异步路径:后端仅建 task(status=processing)+ 立即返回,
    // 真正的分析跑在 BackgroundTasks 里。这里直接走轮询。
    analyzeStep.value = 3
    await _pollUntilDone(resp.task_id)
  } catch (e) {
    ElMessage.error('分析异常:' + (e?.message || e))
  } finally {
    analyzing.value = false
  }
}

// =====【workbench-analysis-intent-input 新增】=====
// 「开始 AI 分析」按钮点击:不再直接发请求,而是先弹"分析重点"对话框,
// 让用户选填本次临时分析重点(可跳过)。确认后再走与原 onAnalyze 相同的执行路径。
const showIntent = ref(false)
function onAnalyzeClick() {
  // 弹窗打开时不需要重置内容:store.draft.analysisIntent 持久化用户上次输入,
  // 再次打开可继续编辑
  showIntent.value = true
}
function onSkipIntentChange(checked) {
  // 勾选「本次跳过」时,自动清空 textarea 内容,确保传到后端的 analysis_intent 是空串
  if (checked) {
    wbStore.draft.analysisIntent = ''
  }
}
async function confirmIntentAndAnalyze() {
  // 若仍勾选了"跳过",显式再清空一遍(覆盖"先勾再取消再勾"等场景)
  const intent = wbStore.draft.skipIntent ? '' : (wbStore.draft.analysisIntent || '').trim()
  showIntent.value = false
  // 调用真正的分析流程 —— 把 analysis_intent 一并透传到后端
  await runAnalyzeWithIntent(intent)
}
// 与原 onAnalyze 同语义,但多带 analysis_intent 参数
async function runAnalyzeWithIntent(analysisIntent) {
  analyzing.value = true
  analyzeResult.value = null
  try {
    analyzeStep.value = 1
    await new Promise((r) => setTimeout(r, 200))
    analyzeStep.value = 2
    // ===== 【chart-edit-prebuilt-dashboard 新增】 =====
    // 若用户在「图表编辑」弹窗里勾选了"加入报表"的图表,随请求一起带上,
    // 后端 analyzer 会按 title 去重合并入最终报告。
    const selectedCharts = wbStore.getSelectedCharts()
    const resp = await workbenchApi.analyze({
      dataset_id: wbStore.draft.datasetId,
      template_id: wbStore.draft.templateId,
      cleaning_rules: wbStore.draft.cleaningRules,
      name: `${wbStore.draft.datasetName} - ${wbStore.draft.templateName}`,
      analysis_intent: analysisIntent,
      selected_charts: selectedCharts,
    })
    // 异步路径:后端仅建 task(status=processing)+ 立即返回 task_id,
    // 真正的分析在 BackgroundTasks 里跑。这里改为「提交→轮询」。
    analyzeStep.value = 3
    await _pollUntilDone(resp.task_id)
  } catch (e) {
    ElMessage.error('分析异常:' + (e?.message || e))
  } finally {
    analyzing.value = false
  }
}

// ===================【workbench-async-polling 新增】===================
// 轮询 GET /api/tasks/{id},直到 status 变成 'success' / 'failed' 才停。
// 设计要点(2026-10 拉长窗口):
//   - 间隔 2000ms:首次请求到 + 预期 30~120s 分析时间,既能秒级感知失败,
//     也不会把后端打挂;从 1.5s 调到 2s 减少空轮询。
//   - 上限 600 次 ≈ 20 分钟:reasoning 模型(MiniMax-M2.7)+ 大 prompt 实测:
//       generate_report ≈ 117s + summarize ≈ 11s + annotate ≈ 34s ≈ 3 分钟。
//       旧 6 分钟上限在 user 报告/复杂数据集场景下偶发撞顶,任务被判「超时」但
//       后端实际上仍在跑 → 体验断档。现拉到 20 分钟,稳进 3~6 倍余量给网络抖动;
//       超过即判定「任务卡死」并给用户友好提示。
//   - 'failed' 时:把 task.error_message 提到 ElMessage 上,而不是只说"失败"。
//   - 'success' 时:从 detail.report.sections / detail.chart_option 拿结果,
//     写到 analyzeResult.value,与原行为一致。
//   - 异步成功 / 失败都保留 task_id,前端「查看报告」按钮可继续跳转;
//     失败时也保留(报告页能展示 failed 详情)。
const _POLL_INTERVAL_MS = 2000
const _POLL_MAX_TIMES = 600
async function _pollUntilDone(taskId) {
  const api = (await import('@/api')).tasksApi
  for (let i = 0; i < _POLL_MAX_TIMES; i++) {
    let detail
    try {
      detail = await api.detail(taskId)
    } catch (e) {
      // 轮询本身失败:不要立刻放弃,可能是临时网络抖动;提示并继续
      ElMessage.warning(`轮询任务状态失败,稍后重试: ${e?.message || e}`)
      await new Promise((r) => setTimeout(r, _POLL_INTERVAL_MS))
      continue
    }
    const status = detail?.status
    if (status === 'success') {
      analyzeResult.value = {
        task_id: taskId,
        sections: detail.report?.sections || {},
        chart_option: detail.chart_option || {},
      }
      ElMessage.success('分析完成')
      return
    }
    if (status === 'failed') {
      // 【workbench-failed-state 修复】不再用本地 pandas 兜底拼接出"29 行 × 24 列"
      // 这种假报告,而是把后端 error_message 直接顶到顶部卡片给用户。
      const msg = detail?.error || 'AI 生成失败,请稍后重试。'
      // ElMessage 仅作瞬时提醒;真正能让用户看清原因的是下面这条卡片。
      ElMessage.error(msg)
      analyzeResult.value = { task_id: taskId, error: msg }
      return
    }
    // processing / 其它:继续轮询
    await new Promise((r) => setTimeout(r, _POLL_INTERVAL_MS))
  }
  // 超时:任务超过 20 分钟还没结束,提示去任务列表看状态
  ElMessage.error(
    `分析任务超时(>${Math.round((_POLL_MAX_TIMES * _POLL_INTERVAL_MS) / 1000)} 秒),请到「任务列表」查看状态。`
  )
  analyzeResult.value = { task_id: taskId, error: 'timeout' }
}

// 工具：合并
async function onMerge() {
  if (mergeForm.value.dataset_ids.length < 2) {
    ElMessage.warning('至少选择 2 个数据集')
    return
  }
  toolLoading.value = true
  try {
    const r = await toolsApi.merge(mergeForm.value)
    // 根据合并模式分别提示:sheet 数或总行数
    if (r.mode === 'multi_sheet') {
      ElMessage.success(`已合并，共 ${r.sheet_count} 个工作表、${r.row_count} 行`)
    } else {
      ElMessage.success(`已合并，共 ${r.row_count} 行`)
    }
    showMerge.value = false
    await loadDatasets()
  } finally {
    toolLoading.value = false
  }
}

// 工具：校验
async function onValidate() {
  toolLoading.value = true
  try {
    // 直接使用当前工作台选中的 dataset_id,后端自动从数据库查 file_path
    if (!wbStore.draft.datasetId) {
      ElMessage.warning('请先选择数据集')
      return
    }
    validateResult.value = await toolsApi.validate({
      dataset_id: wbStore.draft.datasetId,
      // 【tools-results-history-persistence 新增】
      persist: saveAsValidateTask.value,
    })
    if (validateResult.value?.task_id) {
      latestValidateTaskId.value = validateResult.value.task_id
    }
  } catch (e) {
    // axios 拦截器已统一 ElMessage.error
    validateResult.value = null
  } finally {
    toolLoading.value = false
  }
}

// 工具：NL 查询（【nl-chat-upgrade】改造为多轮对话气泡流）
// 把 chatMessages 映射为请求 history:
// - user 气泡取问题原文;ai 气泡取 structured.answer(降级取 explanation)前 500 字
// - 只保留最近 5 轮(10 条);在用户问题入列前快照,即不含本轮(question 字段单独传,避免重复)
function buildNlHistory() {
  return chatMessages.value
    .map((m) => ({
      role: m.role,
      content:
        m.role === 'ai'
          ? String(m.structured?.answer || m.content || '').slice(0, 500)
          : String(m.content || ''),
    }))
    .slice(-10)
}

// 新消息入列后把消息列表滚到底部(等 DOM 更新完成再滚,保证 scrollHeight 是新值)
function scrollNlChatBottom() {
  nextTick(() => {
    const el = nlChatListRef.value
    if (el) el.scrollTop = el.scrollHeight
  })
}

// 发送一条问题:可由输入框(onNlSend())、空状态示例或追问 chips(onNlSend(text))触发
async function onNlSend(text) {
  // chip 触发时用传入文本;输入框触发时取输入框内容
  const question = String(text ?? nlForm.value.question ?? '').trim()
  if (!question) {
    ElMessage.warning('请输入问题')
    return
  }
  if (toolLoading.value) return // 上一轮尚未完成,防重复提交
  // 先快照历史(不含本轮),再把用户问题入列渲染并清空输入框
  const history = buildNlHistory()
  chatMessages.value.push({ role: 'user', content: question })
  nlForm.value.question = ''
  scrollNlChatBottom()
  toolLoading.value = true
  try {
    const resp = await toolsApi.nlQuery({
      dataset_id: wbStore.draft.datasetId,
      question,
      // 【tools-results-history-persistence 保留】
      persist: saveAsNlTask.value,
      // 【nl-chat-upgrade 新增】多轮上下文(最近 5 轮)
      history,
    })
    // 旧单结果状态保留赋值(UI 不再直接展示,兜底兼容)
    nlResult.value = resp
    // AI 回复入列:structured 优先;为 null(LLM 失败/旧响应)时降级为 explanation + rows/columns
    chatMessages.value.push({
      role: 'ai',
      content: resp?.structured?.answer || resp?.explanation || '（后端未返回说明文本）',
      structured: resp?.structured || null,
      rows: resp?.rows || [],
      columns: resp?.columns || [],
    })
    // 【tools-results-history-persistence 保留】每轮成功后更新最新保存的任务 id
    if (resp?.task_id) {
      latestNlTaskId.value = resp.task_id
    }
    scrollNlChatBottom()
  } catch (e) {
    // axios 拦截器已统一 ElMessage.error;本轮失败不入列 AI 气泡
  } finally {
    toolLoading.value = false
  }
}

// 输入框回车发送:排除中文输入法组合态(isComposing/keyCode 229 为选词确认);
// Shift+Enter 因 .exact 修饰符不命中,保持默认换行行为
function onNlEnterKey(e) {
  if (e.isComposing || e.keyCode === 229) return
  onNlSend()
}

// structured.data_table 行切片:最多渲染前 20 行(模板按二维数组 + 列下标取值,与旧表格写法一致)
function nlTableRows(dt) {
  return (dt?.rows || []).slice(0, 20)
}

// 弹窗关闭后清空会话:下次打开是新会话(latestNlTaskId 按原逻辑关闭不重置)
function onNlDialogClosed() {
  chatMessages.value = []
}

// 工具：对比
async function onCompare() {
  if (!compareForm.value.left_dataset_id || !compareForm.value.right_dataset_id) {
    ElMessage.warning('请选择两个数据集')
    return
  }
  // 【dataset-multiformat-upload 新增】非结构化对比无需任何 Pandas 参数,
  // 直接跳过 mode 校验,后端会自动走跨模态对比
  if (compareForm.value.mode === 'contract_compare') {
    // 合同对比:前后端都按 contract_compare 处理,无需额外校验
    return await doCompareRequest()
  }
  if (isDocumentCompare.value) {
    // 任一边是非结构化 → 前后端都按 document_compare 处理,
    // 这里的 mode 字段即使保留也不会生效,显式设为 document_compare 让请求体自描述
    compareForm.value.mode = 'document_compare'
  } else {
    // 按 mode 校验前端必填字段(后端也校验,前端给出更早的提示)
    const mode = compareForm.value.mode
    if (mode === 'business' && (!compareForm.value.group_col || !compareForm.value.metric_col)) {
      ElMessage.warning('业务对比需要选择 分组列 + 指标列')
      return
    }
    if (mode === 'row_diff' && (!compareForm.value.key_cols || compareForm.value.key_cols.length < 1)) {
      ElMessage.warning('变更追踪至少选择 1 个主键列')
      return
    }
  }
  toolLoading.value = true
  try {
    compareResult.value = await toolsApi.compare({
      ...compareForm.value,
      // 【topn 并入 business 新增】未启用 TopN 截断时传 null,
      // 后端据此返回全部分组(而不是默认截断成 Top10)
      top_n: compareForm.value.use_top_n ? compareForm.value.top_n : null,
      // 【tools-results-history-persistence 新增】
      persist: saveAsCompareTask.value,
    })
    if (compareResult.value?.task_id) {
      latestCompareTaskId.value = compareResult.value.task_id
    }
  } catch (e) {
    compareResult.value = null
  } finally {
    toolLoading.value = false
  }
}

// 【dataset-contract-compare 新增】合同对比走相同请求,仅复用此函数避免重复
async function doCompareRequest() {
  toolLoading.value = true
  try {
    compareResult.value = await toolsApi.compare({
      ...compareForm.value,
      persist: saveAsCompareTask.value,
    })
    if (compareResult.value?.task_id) {
      latestCompareTaskId.value = compareResult.value.task_id
    }
  } catch (e) {
    compareResult.value = null
  } finally {
    toolLoading.value = false
  }
}

// 监听左右数据集变化,加载列用于下拉推断
watch(
  () => [compareForm.value.left_dataset_id, compareForm.value.right_dataset_id],
  ([l, r]) => {
    if (l) loadDatasetColumns(l)
    if (r) loadDatasetColumns(r)
  },
)

// 工具：清洗(数据清洗弹窗提交)
// - cleaning_rules 是单一可信源,4 项原始开关在 store.draft.cleaningRules 里
// - 第 5 项 fill_missing 与二级参数 fillMissingPolicy 一并透传(默认 false)
// - 【clean-keep-formulas】keep_formulas 透传(默认 false,勾选后走公式保留回写)
// - 后端响应里的 cleaned_stats 写回 store.draft.previewCleanedStats,供下次打开弹窗预览
async function onCleanSubmit() {
  if (!cleanForm.value.dataset_id) {
    ElMessage.warning('请选择数据集')
    return
  }
  // 【clean-ai-derived-columns-pivot 新增】汇总待提交派生列:
  // 勾选的 AI 建议 + 手动添加行,过滤掉列名 / 公式为空的行,只保留后端契约字段
  const derivedColumns = [
    ...wbStore.draft.derivedSuggestions.filter((s) => s.selected),
    ...wbStore.draft.manualDerivedColumns,
  ]
    .filter((s) => (s.name || '').trim() && (s.expression || '').trim())
    .map((s) => ({
      sheet_name: s.sheet_name || '',
      name: s.name.trim(),
      expression: s.expression.trim(),
    }))
  // ===== 【透视多张并行改造】校验并收集有效的透视配置 =====
  const allPivots = wbStore.draft.pivotConfigs || []
  // 保留公式回写与派生列 / 透视互斥,前端先行拦截
  if (wbStore.draft.keep_formulas && (derivedColumns.length > 0 || allPivots.length > 0)) {
    ElMessage.warning('保留公式回写不支持派生列或透视汇总,请取消其中一项')
    return
  }
  // 每张透视的源工作表 / 行维度 / 汇总字段三项必填
  const incompletePivot = allPivots.find((p) => !p.sheet_name || !p.row_dim || !p.value)
  if (incompletePivot) {
    ElementPlus.ElMessage.warning('请完整选择每张透视表的源工作表、行维度和汇总字段')
    return
  }
  // 汇总字段可能是派生列:若用户随后取消勾选该派生列,候选列表会移除它,
  // 此时提交必然被后端判 400,前端提前拦截并提示重选。
  const invalidPivot = allPivots.find((p) => {
    const opts = pivotValueOptionsForSheet(p.sheet_name)
    return opts.length > 0 && !opts.some((c) => c.name === p.value)
  })
  if (invalidPivot) {
    ElementPlus.ElMessage.warning('某张透视的汇总字段已失效,请重新选择或删除该透视配置')
    return
  }
  cleanLoading.value = true
  try {
    const payload = {
      dataset_id: cleanForm.value.dataset_id,
      cleaning_rules: {
        ...wbStore.draft.cleaningRules,
        fill_missing: wbStore.draft.fill_missing,
        fill_missing_policy: { ...wbStore.draft.fillMissingPolicy },
      },
      keep_formulas: wbStore.draft.keep_formulas,
      derived_columns: derivedColumns,
      name: cleanForm.value.name,
    }
    if (allPivots.length > 0) {
      // 只透传后端契约字段(不含前端生成的 id)
      const pivots = allPivots.map((p) => ({
        sheet_name: p.sheet_name,
        row_dim: p.row_dim,
        col_dim: p.col_dim || null,
        value: p.value,
        agg: p.agg,
        include_margin: p.include_margin,
      }))
      payload.pivots = pivots
      // 向后兼容旧单字段(只在恰好 1 张时回写,便于老前端)
      if (pivots.length === 1) {
        payload.pivot = pivots[0]
      }
    }
    const r = await toolsApi.clean(payload)
    // 时长由 App.vue 的 el-config-provider 全局统一为 3 秒,此处无需单独指定
    ElMessage.success('已保存为新数据集:' + r.name)
    // 写入预览(供下次打开弹窗查看上次结果)
    wbStore.draft.previewCleanedStats = r.cleaned_stats
    showClean.value = false
    cleanForm.value.name = ''
    await loadDatasets()
  } catch (e) {
    // axios 拦截器已统一 ElMessage.error,这里只打日志
    console.error('clean failed:', e)
  } finally {
    cleanLoading.value = false
  }
}

// ===== 【workbench-ai-pdf-tools 新增】AI PDF 处理相关函数 =====

// el-upload before-upload 钩子:校验后缀与大小
// 返回 false 取消上传,返回 true 放行;这里采用"接收文件但暂存,不立刻提交"
function beforePdfUpload(file) {
  const name = (file.name || '').toLowerCase()
  if (!name.endsWith('.pdf')) {
    ElMessage.error('仅支持 PDF 文件')
    return false
  }
  if (file.size > PDF_MAX_SIZE_MB * 1024 * 1024) {
    ElMessage.error(`PDF 文件超过 ${PDF_MAX_SIZE_MB}MB 上限`)
    return false
  }
  pdfFile.value = file
  pdfResult.value = null
  return false  // 阻止 el-upload 自动上传,由"开始处理"按钮手动提交
}

// el-upload on-remove 钩子:用户点击删除文件
function onPdfRemove() {
  pdfFile.value = null
  pdfResult.value = null
}

// ===== 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 的上传处理 =====
// 图片类型白名单(与后端 _IMAGE_SUFFIXES 一致)
const PDF_IMG_EXTS = ['.png', '.jpg', '.jpeg', '.webp', '.bmp']

// el-upload on-change 钩子(多文件):逐个校验后缀/大小/总数,通过才入列
function onImgFileChange(uploadFile, uploadFiles) {
  const f = uploadFile?.raw
  if (!f) return
  const name = (f.name || '').toLowerCase()
  if (!PDF_IMG_EXTS.some((ext) => name.endsWith(ext))) {
    ElMessage.error(`「${f.name}」不是支持的图片格式`)
    // 用过滤后的列表覆盖,把这张从托管列表里剔除
    pdfImgFileList.value = uploadFiles.filter((x) => x.raw !== f)
    return
  }
  if (f.size > PDF_MAX_IMG_SIZE_MB * 1024 * 1024) {
    ElMessage.error(`「${f.name}」超过 ${PDF_MAX_IMG_SIZE_MB}MB 上限`)
    pdfImgFileList.value = uploadFiles.filter((x) => x.raw !== f)
    return
  }
  // 超量检查:以本次选择后的总数为准
  if (uploadFiles.length > PDF_MAX_IMG_COUNT) {
    ElMessage.error(`单次最多 ${PDF_MAX_IMG_COUNT} 张图片`)
    pdfImgFileList.value = uploadFiles.slice(0, PDF_MAX_IMG_COUNT)
    return
  }
  // 保持用户选择的原始顺序(合并 PDF 的页序 = 上传顺序)
  pdfImgFileList.value = uploadFiles.slice()
  pdfResult.value = null
}

// el-upload on-remove 钩子(多文件):移除单张
function onImgRemove(uploadFile, uploadFiles) {
  pdfImgFileList.value = (uploadFiles || []).slice()
  pdfResult.value = null
}

// 切换 tab 时清掉结果,避免串数据
// 【pdf-photo-to-pdf-ocr】同时清掉照片列表,让新 tab 从干净状态开始
watch(pdfTab, () => {
  pdfResult.value = null
  pdfImgFileList.value = []
})

// 弹窗关闭时重置所有状态
function onPdfDialogClosed() {
  pdfFile.value = null
  pdfResult.value = null
  pdfLoading.value = false
  pdfAiEnabled.value = true
  pdfSaveAsDataset.value = false
  pdfSaveToHistory.value = true  // 默认勾选,与初次进入一致
  // 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 状态一并重置
  pdfImgFileList.value = []
  pdfImgMode.value = 'dual'
  pdfImgOcr.value = true
}

// 提交:按当前 tab 路由到对应后端
async function onPdfSubmit() {
  // 【pdf-photo-to-pdf-ocr 新增】照片转 PDF 是多文件,校验逻辑与单文件的 4 个 tab 不同
  if (pdfTab.value === 'image') {
    if (!pdfImgFileList.value.length) {
      ElMessage.warning('请先选择图片文件')
      return
    }
    pdfLoading.value = true
    pdfResult.value = null
    try {
      const fd = new FormData()
      // 多文件必须用**同一个 key** 重复 append,后端才能收到 list[UploadFile]
      for (const item of pdfImgFileList.value) {
        if (item?.raw) fd.append('files', item.raw)
      }
      // 关闭识别时固定发 image,省掉视觉调用
      fd.append('mode', pdfImgOcr.value ? pdfImgMode.value : 'image')
      fd.append('ocr_enabled', String(pdfImgOcr.value))
      fd.append('save_to_history', String(pdfSaveToHistory.value))
      const resp = await toolsApi.pdfImageToPdf(fd)
      pdfResult.value = resp
      // 部分识别失败不算失败:PDF 已产出,只是部分没识别出文字
      if (resp.ocr_failed > 0) {
        ElMessage.warning(`处理完成,${resp.ocr_failed} 张未识别成功(已按原图合并)`)
      } else if (resp.task_id) {
        ElMessage.success(`处理完成,已保存为历史任务 #${resp.task_id}`)
      } else {
        ElMessage.success('处理完成')
      }
    } catch (e) {
      console.error('pdf image_to_pdf failed:', e)
    } finally {
      pdfLoading.value = false
    }
    return
  }

  if (!pdfFile.value) {
    ElMessage.warning('请先选择 PDF 文件')
    return
  }
  pdfLoading.value = true
  pdfResult.value = null
  try {
    const fd = new FormData()
    fd.append('file', pdfFile.value)
    // 4 个端点都接收 save_to_history 与 name
    fd.append('save_to_history', String(pdfSaveToHistory.value))
    let resp
    if (pdfTab.value === 'word') {
      resp = await toolsApi.pdfToWord(fd)
    } else if (pdfTab.value === 'excel') {
      fd.append('ai_enabled', String(pdfAiEnabled.value))
      fd.append('save_as_dataset', String(pdfSaveAsDataset.value))
      resp = await toolsApi.pdfToExcel(fd)
    } else if (pdfTab.value === 'text') {
      fd.append('ai_enabled', String(pdfAiEnabled.value))
      resp = await toolsApi.pdfExtractText(fd)
    } else if (pdfTab.value === 'compress') {
      resp = await toolsApi.pdfCompress(fd)
    } else {
      throw new Error('未知的 PDF 功能')
    }
    pdfResult.value = resp
    if (resp.task_id) {
      ElMessage.success(`处理完成,已保存为历史任务 #${resp.task_id}`)
    } else {
      ElMessage.success('处理完成')
    }
  } catch (e) {
    // axios 拦截器已统一提示
    console.error('pdf tool failed:', e)
  } finally {
    pdfLoading.value = false
  }
}

// 触发下载(由结果区"下载"按钮调用)
async function onPdfDownload() {
  const r = pdfResult.value
  if (!r || !r.download_url) return
  try {
    const url = toolsApi.pdfDownloadUrl(r.file_name)
    await toolsApi.downloadFile(url, r.file_name)
  } catch (e) {
    ElMessage.error('下载失败')
    console.error('pdf download failed:', e)
  }
}

// 字节数 → 人类可读
function formatSize(bytes) {
  if (!bytes && bytes !== 0) return '—'
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`
}

// 加载历史任务列表(图表编辑页用,含图表数量)
async function loadHistoryTasks() {
  try {
    const data = await tasksApi.list()
    const items = data.items || []
    console.log('[chart_edit] 历史任务原始数据:', items.length, items.slice(0, 2))
    // 后端返回 chart_count,如果没有这个字段(旧数据)就兜底为 0
    // —— 不过滤 chart_count=0,让用户能看到全部历史,选了再判断
    historyTaskOptions.value = items.map((t) => ({
      ...t,
      chart_count: t.chart_count ?? 0,
    }))
  } catch (e) {
    console.error('[chart_edit] 加载历史任务失败:', e)
    historyTaskOptions.value = []
  }
}

// 切换历史任务时:加载该任务的完整「可改图」图表清单
// 后端从 vN 起在详情里直接返回 available_charts(已合并 charts / by_section / reports[*].charts / selected_charts),
// 这里是单一可信源;只有当后端缺失该字段(例如老版本)时才做兜底展平。
function _flattenLegacyChartOption(co) {
  if (!co) return []
  // 1) 顶层 charts
  if (Array.isArray(co.charts) && co.charts.length) return co.charts
  // 2) by_section:{section:[chart]} 或 [{section, charts:[chart]}]
  const bs = co.by_section
  if (bs && typeof bs === 'object') {
    if (!Array.isArray(bs)) {
      const out = []
      for (const lst of Object.values(bs)) {
        if (Array.isArray(lst)) out.push(...lst)
      }
      if (out.length) return out
    } else {
      const out = []
      for (const entry of bs) {
        if (entry && Array.isArray(entry.charts)) out.push(...entry.charts)
      }
      if (out.length) return out
    }
  }
  // 3) reports[*].charts
  if (Array.isArray(co.reports)) {
    const out = []
    for (const rep of co.reports) {
      if (rep && Array.isArray(rep.charts)) out.push(...rep.charts)
    }
    if (out.length) return out
  }
  return []
}

async function onHistoryTaskChange(taskId) {
  if (!taskId) {
    historyChartOptions.value = []
    chartHistoryIndex.value = 0
    return
  }
  // 缓存里有就直接用,避免重复拉
  if (historyTaskCache[taskId]) {
    const d = historyTaskCache[taskId]
    historyChartOptions.value =
      (Array.isArray(d.available_charts) && d.available_charts.length)
        ? d.available_charts
        : _flattenLegacyChartOption(d.chart_option)
    chartHistoryIndex.value = 0
    return
  }
  try {
    const detail = await tasksApi.detail(taskId)
    historyTaskCache[taskId] = detail
    // 优先使用后端聚合后的 available_charts;缺失时再用兜底展平
    if (Array.isArray(detail?.available_charts) && detail.available_charts.length) {
      historyChartOptions.value = detail.available_charts
    } else {
      historyChartOptions.value = _flattenLegacyChartOption(detail?.chart_option)
    }
    chartHistoryIndex.value = 0
  } catch (e) {
    historyChartOptions.value = []
  }
}

// 图表编辑:调用后端 LLM 生成新 option（仅生成候选图,不直接落库）
async function onChartEdit() {
  const instr = chartInstruction.value.trim()
  if (!instr) {
    ElMessage.warning('请输入指令')
    return
  }
  if (!activeDatasetId.value) {
    ElMessage.warning('请先选择数据集')
    return
  }
  // 【修复】点击时打印关键上下文,方便排查"按钮点不动"问题:
  // 历史上 activeDatasetId 为 null 会让按钮被 disabled,导致用户以为"点击没反应",
  // 看起来页面没动静。其实是 :disabled 直接拦了 click 事件。
  // 这里在 click 真的触发后,记录一次成功进入函数体的日志;
  // 如果未来又出现"点了没反应",用户对照控制台即可定位到底是按钮没触发、还是接口报错。
  console.debug('[chart_edit] click', {
    source: chartSource.value,
    taskId: chartHistoryTaskId.value,
    chartIndex: chartHistoryIndex.value,
    activeChartType: activeChart.value?.type,
    datasetId: activeDatasetId.value,
    instruction: instr,
    currentOptionKeys: activeChart.value?.option ? Object.keys(activeChart.value.option) : null,
  })
  chartEditLoading.value = true
  try {
    // ===== 【修复】两条路径都走 workbenchApi.modifyChart(本地规则引擎 + LLM 兜底) =====
    // 理由:
    //   旧历史模式直接调 /api/tools/chart_edit,该接口把指令原样塞给 LLM,LLM 经常返回
    //   非标准 JSON(单引号、缺引号、未尾闭合),json.loads 失败,前端 fallback=true
    //   但"配色换成蓝绿色"这种明确指令,本来就该走 _apply_rule 关键词命中,
    //   根本不需要 LLM。新接口会先做规则匹配,命中即改,只有未命中才回退 LLM。
    //   行为对齐工作台模式,用户体验一致。
    const target = activeChart.value || {}
    chartEditResult.value = await workbenchApi.modifyChart({
      dataset_id: activeDatasetId.value,
      instruction: instr,
      style_key: wbStore.draft.chartStyle,
      target_chart: {
        title: target.title || '图表',
        type: target.type || 'bar',
        option: target.option || {},
        style_key: target.style_key || wbStore.draft.chartStyle,
      },
    })
    if (!chartEditResult.value?.fallback) {
      ElMessage.success(chartEditResult.value?.explanation || '已生成候选图表,请确认后再应用')
    } else {
      // 【修复】fallback 状态下的两种情况要区分对待:
      // 1) 后端完全失败(option 是原图/空):给红色 toast 告诉用户编辑失败;
      // 2) 后端走"本地降级"也给了可用 option(配色诉求本地规则处理):给 info 提示。
      const r = chartEditResult.value
      const has_new_option = r?.option && Object.keys(r.option).length
        && (r.explanation || '').includes('本地降级')
      if (has_new_option) {
        ElMessage.info(`AI 暂时不可用,已本地降级: ${r.explanation || '已按指令改图'}`)
      } else {
        const reason = chartEditResult.value?.error
        ElMessage.warning(
          reason
            ? `编辑失败,已保留原图(${reason.slice(0, 60)})`
            : '编辑失败,已保留原图'
        )
      }
    }
  } catch (e) {
    chartEditResult.value = null
  } finally {
    chartEditLoading.value = false
  }
}

// ===== 【chart-edit-prebuilt-dashboard 新增】 =====
// 把"指令改图"的候选图替换到当前工作台看板的某个位置:
// - 若已有勾选加入报表的图,优先替换第一个被勾选的
// - 否则替换"最后一张"(即 activeChart 指向的那张)
function applyChartToDashboard() {
  const r = chartEditResult.value
  if (!r?.option || !Object.keys(r.option).length) {
    ElMessage.warning('候选图表为空,无法应用')
    return
  }
  const groups = (wbStore.draft.dashboardGroups || []).slice()
  if (!groups.length || !groups.some((g) => g.charts.length)) {
    ElMessage.warning('请先生成看板')
    return
  }
  // 找目标 key(优先用 chartEditTargetIdx,否则第一个被勾选,否则最后一个)
  let targetKey = null
  if (chartEditTargetIdx.value) {
    targetKey = chartEditTargetIdx.value
  } else {
    for (const g of groups) {
      for (const c of g.charts) {
        if (wbStore.draft.selectedChartKeys.has(c._key)) {
          targetKey = c._key; break
        }
      }
      if (targetKey) break
    }
    if (!targetKey) {
      // 拿最后一个 chart
      const lastG = groups[groups.length - 1]
      const lastC = lastG.charts[lastG.charts.length - 1]
      targetKey = lastC._key
    }
  }
  // 在 groups 里定位并替换
  let replaced = false
  for (const g of groups) {
    for (let i = 0; i < g.charts.length; i++) {
      if (g.charts[i]._key === targetKey) {
        g.charts[i] = {
          ...g.charts[i],
          title: r.title || g.charts[i].title,
          option: r.option,
          style_key: r.style_key || wbStore.draft.chartStyle,
        }
        replaced = true
        break
      }
    }
    if (replaced) break
  }
  wbStore.setDashboard({ groups }, true)
  ElMessage.success('已替换看板图表')
  chartEditResult.value = null
}

// 历史模式:把候选图显式落库到指定任务的 chart_option.charts[index]
async function applyChartToHistoryTask() {
  const r = chartEditResult.value
  if (!r?.option || !Object.keys(r.option).length) {
    ElMessage.warning('候选图表为空,无法应用')
    return
  }
  if (chartSource.value !== 'history' || !chartHistoryTaskId.value) {
    ElMessage.warning('仅历史模式支持应用到任务')
    return
  }
  try {
    const updated = await toolsApi.chartApply({
      task_id: chartHistoryTaskId.value,
      chart_index: chartHistoryIndex.value,
      type: r.type,
      title: r.title,
      option: r.option,
    })
    // 同步本地缓存,避免重新拉
    const taskId = chartHistoryTaskId.value
    if (historyTaskCache[taskId]) {
      historyTaskCache[taskId] = {
        ...historyTaskCache[taskId],
        chart_option: updated.chart_option || historyTaskCache[taskId].chart_option,
      }
      // 顺手刷新弹窗里展示的历史图表
      const charts = updated.chart_option?.charts || historyTaskCache[taskId].chart_option?.charts || []
      historyChartOptions.value = charts
    }
    ElMessage.success('已应用到历史任务,可在报告页查看')
    chartEditResult.value = null
  } catch (e) {
    // http.js 已经统一处理错误提示
  }
}

// 把新图表写入 analyzeResult.chart_option.charts,工作台右侧立即刷新
function applyChartToResult() {
  if (!chartEditResult.value?.option || !Object.keys(chartEditResult.value.option).length) return
  if (!analyzeResult.value) {
    ElMessage.warning('请先点击「开始 AI 分析」生成初始报告')
    return
  }
  const charts = (analyzeResult.value.chart_option?.charts || []).slice()
  if (!charts.length) {
    // 工作台原本没有图表,直接 push 一个
    charts.push({
      type: chartEditResult.value.type,
      title: chartEditResult.value.title,
      option: chartEditResult.value.option,
    })
  } else {
    // 替换第一个图表,保留位置和标题
    charts[0] = {
      ...charts[0],
      type: chartEditResult.value.type,
      title: chartEditResult.value.title || charts[0].title,
      option: chartEditResult.value.option,
    }
  }
  analyzeResult.value = {
    ...analyzeResult.value,
    chart_option: { ...analyzeResult.value.chart_option, charts },
  }
  ElMessage.success('已应用到工作台图表')
  showChart.value = false
}

// 监听 toolDialog 切换时打开弹窗（占位保留）

onMounted(async () => {
  // 串行加载,任意一步失败不阻塞另一步(各自有 catch/守卫)
  try { await loadDatasets() } catch (e) { console.error('[workbench] loadDatasets failed:', e) }
  try { await loadPreview() } catch (e) { console.error('[workbench] loadPreview failed:', e) }
})
</script>

<style scoped>
.wb-page {
  padding: 16px 24px;
}
/* 左侧栏：所有交互元素统一左对齐，跟 wb-section-title 保持同一条边线 */
.wb-side :deep(.wb-side-action) {
  width: 100%;
  justify-content: flex-start;
  padding: 8px 12px;
  margin-bottom: 6px;
}
/* 关键：穿透到 Element Plus 内部的 .el-button__inner 与 el-icon */
.wb-side :deep(.wb-side-action .el-button__inner),
.wb-side :deep(.wb-side-action > span) {
  justify-content: flex-start !important;
  width: 100%;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-weight: 500;
}
.wb-side :deep(.wb-side-action .el-icon) {
  margin: 0 !important;
  display: inline-flex;
  font-size: 16px;
}
.wb-side :deep(.wb-side-action:last-of-type) {
  margin-bottom: 0;
}
.wb-side :deep(.wb-side-primary) {
  margin-top: 18px;
}
.wb-side .ds-card {
  margin-bottom: 8px;
  padding: 8px 12px;
}
/* 【workbench-per-sheet-preview 新增】预览区顶部 sheet tab 样式:
   - 与暗色主题适配,选中 tab 高亮为 cyan,未选中 tab 灰底;
   - 限制最大高度避免 5 个以上 sheet 把页面撑高。 */
.wb-preview-sheet-tabs {
  margin: 0 0 10px 0;
}
.wb-preview-sheet-tabs :deep(.el-tabs__header) {
  margin: 0;
}
.wb-preview-sheet-tabs :deep(.el-tabs__nav-wrap) {
  max-height: 36px;
}
.wb-preview-sheet-tabs :deep(.el-tabs__item) {
  font-size: 12px;
  padding: 0 12px !important;
  height: 30px;
  line-height: 30px;
}
.wb-preview-sheet-tabs :deep(.el-tabs__item.is-active) {
  color: #22d3ee !important;
  border-color: #22d3ee !important;
}

/* 【dataset-multisheet-merge 新增】工作表标签的微调:与数据集名/ID 文字层区分 */
.wb-side .ds-sheet-tag {
  margin-top: 6px;
  display: inline-block;
}
.wb-side .small {
  font-size: 11px;
}
.metric-row {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
}
.status-bar {
  margin-top: 18px;
  padding: 14px;
  background: rgba(10, 20, 48, 0.5);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
}
.result-section {
  margin-bottom: 18px;
}
.mini-chart {
  height: 180px;
  width: 100%;
}
.ai-section {
  margin-top: 12px;
  background: rgba(10, 20, 48, 0.5);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
  padding: 10px 12px;
}
.ai-text {
  color: var(--text-primary);
  font-size: 13px;
  line-height: 1.6;
  margin-top: 4px;
}
.json-block {
  background: rgba(10, 20, 48, 0.6);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 6px;
  padding: 10px;
  color: var(--accent-cyan-2);
  white-space: pre-wrap;
  font-size: 12px;
  max-height: 320px;
  overflow: auto;
}

/* ========== 【workbench-failed-state 新增】任务失败卡片样式 ==========
 * 旧实现:LLM 失败后,前端继续渲染"29 行 × 24 列"式本地拼接报告,
 * 用户看不出是 AI 失败,误以为是 AI 的"分析"。
 * 新实现:任务 failed 时把失败原因顶到右侧结果区顶部,给出重跑入口。
 * 配色沿用 dashboard 暗色主题,红底浅字 + 圆角,与其它 result-section 风格一致。 */
.wb-result-failed {
  margin-bottom: 16px;
  padding: 12px 14px;
  background: rgba(248, 113, 113, 0.08);
  border: 1px solid rgba(248, 113, 113, 0.45);
  border-left: 3px solid #f87171;
  border-radius: 8px;
  color: var(--text-primary);
}
.wb-result-failed-title {
  font-size: 14px;
  font-weight: 600;
  color: #f87171;
  margin-bottom: 6px;
}
.wb-result-failed-body {
  font-size: 13px;
  line-height: 1.7;
  color: var(--text-primary);
  white-space: pre-wrap;
  word-break: break-word;
  margin-bottom: 10px;
}
.wb-result-failed-actions {
  display: flex;
  gap: 8px;
}

/* ========== 数据校验·AI 解读卡片 ========== */
.ai-interpret {
  margin-top: 14px;
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.ai-summary {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 12px 14px;
  border-radius: 8px;
  border: 1px solid var(--border-cyan-soft);
  background: rgba(10, 20, 48, 0.55);
}
.ai-summary .level-badge {
  display: inline-block;
  padding: 2px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 600;
  letter-spacing: 1px;
}
.ai-summary .summary-text {
  font-size: 14px;
  color: var(--text-primary);
  flex: 1;
  line-height: 1.6;
}
.ai-summary.level-good {
  border-color: rgba(74, 222, 128, 0.45);
  background: linear-gradient(90deg, rgba(74, 222, 128, 0.10), rgba(10, 20, 48, 0.4));
}
.ai-summary.level-good .level-badge { background: rgba(74, 222, 128, 0.18); color: #4ade80; }
.ai-summary.level-warn {
  border-color: rgba(250, 204, 21, 0.45);
  background: linear-gradient(90deg, rgba(250, 204, 21, 0.10), rgba(10, 20, 48, 0.4));
}
.ai-summary.level-warn .level-badge { background: rgba(250, 204, 21, 0.18); color: #fbbf24; }
.ai-summary.level-danger {
  border-color: rgba(248, 113, 113, 0.45);
  background: linear-gradient(90deg, rgba(248, 113, 113, 0.10), rgba(10, 20, 48, 0.4));
}
.ai-summary.level-danger .level-badge { background: rgba(248, 113, 113, 0.18); color: #f87171; }

.ai-block {
  padding: 12px 14px;
  border-radius: 8px;
  background: rgba(10, 20, 48, 0.55);
  border: 1px solid var(--border-cyan-soft);
}
.ai-block .ai-h {
  font-size: 12px;
  color: var(--accent-cyan-2);
  font-weight: 600;
  letter-spacing: 0.5px;
  margin-bottom: 6px;
}
.ai-block ul {
  margin: 0;
  padding-left: 22px;
  color: var(--text-primary);
  font-size: 13px;
  line-height: 1.7;
}
.ai-block ul li { list-style: disc; margin-bottom: 2px; }
.ai-block.highlights ul li { list-style: '✓  '; }
.ai-block.suggestions ul li { list-style: '→  '; }
.ai-block.issues ul li { list-style: none; margin-left: -18px; display: flex; align-items: baseline; flex-wrap: wrap; gap: 6px; }
.ai-block.issues ul li .sev-badge {
  display: inline-block;
  font-size: 11px;
  padding: 1px 8px;
  border-radius: 999px;
  font-weight: 600;
}
.sev-badge.sev-low { background: rgba(96, 165, 250, 0.2); color: #60a5fa; }
.sev-badge.sev-medium { background: rgba(250, 204, 21, 0.2); color: #fbbf24; }
.sev-badge.sev-high { background: rgba(248, 113, 113, 0.2); color: #f87171; }

/* ========== 数据校验·原始数据表格 ========== */
.raw-preview {
  margin-top: 12px;
  padding: 12px 14px;
  background: rgba(10, 20, 48, 0.55);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
}
.raw-preview-title {
  font-size: 12px;
  color: var(--accent-cyan-2);
  font-weight: 600;
  letter-spacing: 0.5px;
  margin-bottom: 8px;
}

/* ========== 图表编辑弹窗(白底风格 · 与 Word/PPT 导出观感对齐) ==========
 * 设计意图:整个弹窗翻成白底,让用户在「改图」时看到的就是导出后图表的真实配色,
 * 不会被系统深色主题压住,改出来即可信。复用 ReportView 里白底报告的色板(#1f2937/#0e7490)。
 */
/* 弹窗整体反色:覆盖 Element Plus 默认深色样式 */
.chart-edit-dialog.el-dialog {
  background: #ffffff !important;
  border: 1px solid #e5e7eb !important;
  color: #1f2937 !important;
}
.chart-edit-dialog .el-dialog__title,
.chart-edit-dialog .el-dialog__body {
  color: #1f2937 !important;
}
/* 弹窗里的表单 label / 提示文字强制深色 */
.chart-edit-dialog .el-form-item__label,
.chart-edit-dialog .dim {
  color: #1f2937 !important;
}
/* 单选组(radio button)的白底适配,避免被压在深色上 */
.chart-edit-dialog .el-radio-button__inner {
  background: #ffffff !important;
  color: #1f2937 !important;
  border-color: #e5e7eb !important;
}
.chart-edit-dialog .el-radio-button.is-active .el-radio-button__inner {
  background: #0e7490 !important;
  color: #ffffff !important;
  border-color: #0e7490 !important;
}
/* select 下拉触发器也是白底 */
.chart-edit-dialog .el-select .el-input__wrapper {
  background: #ffffff !important;
  box-shadow: inset 0 0 0 1px #e5e7eb !important;
  color: #1f2937 !important;
}
/* 弹窗里的 input(指令输入框)同步白底 */
.chart-edit-dialog .chart-edit-row .el-input__wrapper {
  background: #ffffff !important;
  box-shadow: inset 0 0 0 1px #e5e7eb !important;
  color: #1f2937 !important;
}
/* 「快捷指令」标签:浅灰边 + 白底,hover 时转深青 */
.chart-edit-dialog .quick-tag {
  background: #ffffff !important;
  border-color: #e5e7eb !important;
  color: #1f2937 !important;
}
.chart-edit-dialog .quick-tag:hover {
  color: #0e7490 !important;
  border-color: #0e7490 !important;
}

.chart-preview-block {
  margin-top: 4px;
  padding: 10px 12px;
  background: #ffffff;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  margin-bottom: 12px;
}
.chart-preview-title {
  font-size: 12px;
  color: #0e7490;
  font-weight: 600;
  letter-spacing: 0.5px;
  margin: 8px 0 6px;
}
.chart-canvas {
  width: 100%;
  height: 280px;
}
.chart-edit-row {
  display: flex;
  gap: 8px;
  align-items: center;
  margin: 10px 0;
}
.chart-edit-row .el-input { flex: 1; }
.quick-instructions {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  align-items: center;
  margin-bottom: 10px;
}
.quick-tag { cursor: pointer; transition: all .15s; }
.quick-tag:hover {
  color: #0e7490;
  border-color: #0e7490;
}
.chart-edit-result {
  margin-top: 12px;
  padding: 12px 14px;
  background: #ffffff;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
}
.chart-actions {
  display: flex;
  gap: 8px;
  justify-content: flex-end;
  margin-top: 10px;
}

/* ========== 【chart-edit-prebuilt-dashboard 新增】前置看板样式(按 sheet 分组) ========== */
.dashboard-groups {
  max-height: 460px;
  overflow-y: auto;
  padding: 4px;
}
.dashboard-group {
  margin-bottom: 14px;
}
.dashboard-group-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
  padding: 6px 10px;
  background: rgba(45, 212, 191, 0.08);
  border-left: 3px solid #2dd4bf;
  border-radius: 4px;
}
.dashboard-group-title {
  font-size: 13px;
  font-weight: 600;
  color: #0e7490;
  letter-spacing: 0.5px;
}
.dashboard-toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px;
  padding: 10px 12px;
  margin-bottom: 10px;
  background: rgba(10, 20, 48, 0.4);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
}
.dashboard-toolbar .tb-style {
  margin-bottom: 0;
}
.dashboard-toolbar .tb-stat {
  margin-left: auto;
}
.dashboard-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 10px;
  padding: 4px;
}
.dashboard-empty {
  padding: 30px 20px;
  text-align: center;
  background: rgba(255, 255, 255, 0.55);
  border: 1px dashed #cbd5e1;
  border-radius: 8px;
  color: #1f2937;
}
.dashboard-empty .de-icon {
  font-size: 32px;
  margin-bottom: 6px;
}
.dashboard-empty .de-text {
  font-size: 13px;
  line-height: 1.6;
  color: #475569;
}
.chart-card {
  background: #ffffff;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  padding: 8px 10px;
  transition: border-color .15s, box-shadow .15s;
}
.chart-card.selected {
  border-color: #10b981;
  box-shadow: 0 0 0 2px rgba(16, 185, 129, 0.18);
}
.chart-card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin-bottom: 4px;
}
.chart-card-title {
  font-size: 12px;
  font-weight: 600;
  color: #1f2937;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  flex: 1;
  min-width: 0;
}
.chart-card-actions {
  display: flex;
  gap: 4px;
  flex-shrink: 0;
}
.chart-card-canvas {
  height: 200px;
  width: 100%;
}
.nl-result {
  margin-top: 12px;
}
.compare-result {
  margin-top: 12px;
}
/* 数据清洗弹窗底部预览行:灰底浅字,与其它提示文案风格一致 */
.wb-clean-preview {
  margin: 12px 0 0;
  padding: 10px 12px;
  background: rgba(10, 20, 48, 0.5);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 6px;
  color: #94a3b8;
  font-size: 12px;
  line-height: 1.6;
}

/* ===== 【clean-ai-derived-columns-pivot 新增】智能派生列 / 透视汇总区块 ===== */
/* 与 .wb-clean-preview 同款灰底浅字风格,不引入新图标 */
.wb-clean-block {
  margin: 12px 0 0;
  padding: 10px 12px;
  background: rgba(10, 20, 48, 0.5);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 6px;
  color: #94a3b8;
  font-size: 12px;
  line-height: 1.6;
}
.wb-clean-block-head {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 12px;
}
.wb-clean-block-title {
  color: #2dd4bf;
  font-size: 13px;
  font-weight: 600;
}
.wb-clean-block-tip {
  margin-top: 2px;
  color: #94a3b8;
}
/* 派生列列表:每行两段(勾选 / 列名 / 标签 / 说明 + 公式) */
.wb-derived-list {
  margin-top: 8px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.wb-derived-item {
  padding: 8px 10px;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 6px;
}
.wb-derived-line {
  display: flex;
  align-items: center;
  gap: 8px;
}
.wb-derived-line + .wb-derived-line {
  margin-top: 6px;
}
.wb-derived-line :deep(.el-checkbox__label) {
  color: #94a3b8;
  font-size: 12px;
}
.wb-derived-name,
.wb-derived-sheet {
  width: 150px;
  flex: none;
}
.wb-derived-label {
  flex: none;
  color: #94a3b8;
}
.wb-derived-desc {
  color: #64748b;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.wb-derived-empty {
  margin-top: 8px;
  color: #64748b;
}
.wb-derived-add {
  margin-top: 8px;
}
/* ===== 透视汇总区块:两区域分离式布局 ===== */
/* 区块标题行:横向,说明文字紧随其后 */
.wb-pivot-block-title-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 10px;
}
.wb-pivot-block-title-row .wb-clean-block-tip {
  color: #64748b;
  font-size: 11px;
}

/* ---- 区域 A:AI 推荐 ---- */
.wb-pivot-recommend-section {
  margin-bottom: 12px;
}
.wb-pivot-recommend-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
}
.wb-pivot-recommend-label {
  color: #cbd5e1;
  font-size: 12px;
  font-weight: 500;
}
.wb-pivot-recommend-icon {
  color: #67e8f9;
  margin-right: 4px;
}
.wb-pivot-recommend-toggle {
  color: #64748b;
  font-size: 11px;
  padding: 0;
}
.wb-pivot-recommend-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
/* 每条 AI 推荐方案:独立卡片,两行布局 */
.wb-pivot-recommend-item {
  padding: 8px 12px;
  background: rgba(20, 30, 60, 0.55);
  border: 1px solid rgba(100, 116, 139, 0.25);
  border-radius: 6px;
  transition: border-color 0.15s, background 0.15s;
}
.wb-pivot-recommend-item:hover {
  border-color: rgba(103, 232, 249, 0.4);
  background: rgba(20, 30, 60, 0.75);
}
/* 第一行:结构化信息 */
.wb-pivot-recommend-main {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #e2e8f0;
}
.wb-pivot-sheet-tag {
  background: rgba(103, 232, 249, 0.12);
  border: 1px solid rgba(103, 232, 249, 0.3);
  color: #67e8f9;
  font-size: 11px;
  font-family: 'SF Mono', 'Menlo', monospace;
  padding: 0 6px;
  height: 20px;
  line-height: 18px;
}
.wb-pivot-recommend-dim {
  color: #93c5fd;
  font-weight: 500;
}
.wb-pivot-recommend-op {
  color: #a78bfa;
  font-size: 12px;
}
.wb-pivot-recommend-arrow {
  color: #475569;
  font-size: 14px;
}
.wb-pivot-recommend-val {
  color: #34d399;
  font-weight: 500;
  display: inline-flex;
  align-items: center;
  gap: 4px;
}
/* 第二行:描述 + 应用按钮 */
.wb-pivot-recommend-footer {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-top: 6px;
}
.wb-pivot-recommend-desc {
  color: #64748b;
  font-size: 11px;
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  margin-right: 8px;
}
.wb-pivot-recommend-apply {
  flex-shrink: 0;
  background: rgba(103, 232, 249, 0.15);
  border: 1px solid rgba(103, 232, 249, 0.4);
  color: #67e8f9;
  padding: 0 12px;
  height: 24px;
}
.wb-pivot-recommend-apply:hover {
  background: rgba(103, 232, 249, 0.25);
  border-color: rgba(103, 232, 249, 0.6);
}
/* 派生列 tag:柔和蓝紫色,与 error 橙色区分 */
.wb-pivot-derived-tag-soft {
  background: rgba(167, 139, 250, 0.15);
  border: 1px solid rgba(167, 139, 250, 0.35);
  color: #c4b5fd;
  font-size: 10px;
  padding: 0 5px;
  height: 18px;
  line-height: 16px;
}

/* ---- 分隔线 ---- */
.wb-pivot-divider {
  height: 1px;
  background: linear-gradient(to right, transparent, rgba(100, 116, 139, 0.4), transparent);
  margin: 10px 0 14px;
}

/* ---- 区域 B:已选透视汇总 ---- */
.wb-pivot-selected-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
}
.wb-pivot-selected-label {
  color: #cbd5e1;
  font-size: 12px;
  font-weight: 500;
}
.wb-pivot-selected-count {
  color: #67e8f9;
  margin-left: 2px;
}
.wb-pivot-add-manual {
  font-size: 12px;
}
.wb-pivot-empty {
  padding: 14px;
  text-align: center;
  color: #475569;
  font-size: 12px;
  background: rgba(14, 19, 48, 0.4);
  border: 1px dashed rgba(100, 116, 139, 0.3);
  border-radius: 6px;
}

/* 已选配置卡片列表 */
.wb-pivot-config-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
/* 单张配置卡片 */
.wb-pivot-config-card {
  padding: 12px 14px;
  background: rgba(20, 30, 60, 0.65);
  border: 1px solid rgba(100, 116, 139, 0.35);
  border-radius: 6px;
}
.wb-pivot-config-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 10px;
  padding-bottom: 8px;
  border-bottom: 1px solid rgba(100, 116, 139, 0.2);
}
.wb-pivot-config-title-row {
  display: flex;
  align-items: center;
  gap: 8px;
}
.wb-pivot-config-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 22px;
  height: 22px;
  font-size: 11px;
  font-weight: 600;
  color: #0f172a;
  background: linear-gradient(135deg, #67e8f9, #34d399);
  border-radius: 50%;
}
.wb-pivot-config-sheet {
  color: #93c5fd;
  font-size: 12px;
  font-family: 'SF Mono', 'Menlo', monospace;
}

/* 卡片内表单 grid */
.wb-pivot-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 10px 14px;
}
.wb-pivot-item {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.wb-pivot-full {
  grid-column: 1 / -1;
}
.wb-pivot-label {
  color: #94a3b8;
  font-size: 11px;
}
.wb-pivot-field {
  width: 100%;
}
/* 汇总字段下拉中的派生列小标记 */
.wb-derived-option-tag {
  color: #c4b5fd;
  font-size: 11px;
  background: rgba(167, 139, 250, 0.12) !important;
  border: none !important;
  padding: 0 4px !important;
  height: 16px !important;
  line-height: 14px !important;
  margin-left: 4px;
}
.wb-clean-warnings {
  margin-top: 4px;
  color: #64748b;
  font-size: 12px;
}

/* 【workbench-ai-pdf-tools 新增】PDF 工具结果区:与清洗预览同款风格 */
.pdf-result-box {
  margin-top: 14px;
  padding: 10px 12px;
  background: rgba(10, 20, 48, 0.5);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 6px;
  color: #cbd5e1;
  font-size: 13px;
}
.pdf-result-box .pdf-result-line {
  line-height: 1.8;
}
.pdf-result-box .el-button {
  margin-top: 8px;
}
/* 【pdf-tools-history-persistence 新增】保存为历史任务勾选行 */
.pdf-save-row {
  margin: 8px 0 4px;
  padding: 4px 0;
  font-size: 13px;
  color: var(--text-primary);
}
/* 已保存为历史任务卡片(与 saved-task-card 同款) */
.pdf-saved-task-card {
  margin-top: 12px;
  padding: 10px 12px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: rgba(34, 211, 238, 0.08);
  border: 1px solid rgba(34, 211, 238, 0.35);
  border-radius: 6px;
}
.pdf-saved-task-card .pdf-saved-task-text {
  color: #cbd5e1;
  font-size: 13px;
}

/* ========== 【tools-results-history-persistence 新增】已保存为历史任务卡片 ========== */
.saved-task-card {
  margin-top: 12px;
  padding: 10px 12px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  background: rgba(34, 197, 94, 0.10);
  border: 1px solid rgba(34, 197, 94, 0.45);
  border-radius: 8px;
  color: #bbf7d0;
  font-size: 13px;
}

/* ========== 数据集对比:模式通用 ========== */
.result-block {
  margin-top: 14px;
  background: rgba(10, 20, 48, 0.55);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
  padding: 12px 14px;
}
.rb-title {
  margin: 0 0 10px;
  font-size: 13px;
  color: var(--accent-cyan-2);
  font-weight: 600;
  letter-spacing: 0.5px;
}
.muted { color: #94a3b8; font-size: 12px; }

/* ========== 【dataset-contract-compare 新增】合同对比结果样式 ==========
   - 条款差异表的列标签略加粗
   - 高优先级变更行加底色,提示重点关注
   - 利弊卡分 3 种状态:左优/右优/各有优劣,色板对应
   - 谈判建议加小图标 */
.contract-clause-label {
  font-weight: 600;
  color: var(--text-primary);
}
:deep(.contract-row-high) td {
  background: rgba(248, 113, 113, 0.08) !important;
}
:deep(.contract-row-high) td:first-child {
  border-left: 3px solid #f87171;
}
.contract-section {
  margin-top: 18px;
}
.pros-cons-card {
  margin-bottom: 10px;
  padding: 10px 14px;
  background: rgba(10, 20, 48, 0.45);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 6px;
  border-left-width: 3px;
}
.pros-cons-card.pc-left  { border-left-color: #4ade80; }
.pros-cons-card.pc-right { border-left-color: #fbbf24; }
.pros-cons-card.pc-both  { border-left-color: #94a3b8; }
.pc-head {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 6px;
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
}
.is-up { color: #4ade80; }
.is-down { color: #f87171; }

/* ========== AI 中文叙述 ========== */
.compare-narration {
  margin-top: 14px;
  background: rgba(10, 20, 48, 0.55);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
  padding: 14px 16px;
}
.cn-title {
  margin: 0 0 8px;
  font-size: 14px;
  color: var(--accent-cyan-2);
  font-weight: 600;
  letter-spacing: 1px;
}
.cn-summary {
  font-size: 14px;
  line-height: 1.7;
  color: var(--text-primary);
  margin: 0 0 12px;
  padding: 10px 12px;
  background: rgba(34, 211, 238, 0.06);
  border-left: 3px solid var(--accent-cyan-2);
  border-radius: 4px;
}
.cn-empty {
  font-size: 13px;
  color: #94a3b8;
  margin: 0 0 12px;
}
.cn-section { margin-top: 10px; }
.cn-section .cn-h {
  font-size: 12px;
  color: var(--accent-cyan-2);
  margin-bottom: 4px;
  font-weight: 600;
}
.cn-section ul {
  margin: 0;
  padding-left: 20px;
  /* 工作台卡片背景是 rgba(10,20,48,0.55) 半透明深蓝,文字用 --text-strong 提高可读性 */
  color: var(--text-strong, #f8fafc);
  font-size: 13px;
  line-height: 1.7;
}
.cn-section ul li {
  list-style: disc;
  /* 显式声明 li 颜色,在半透明深蓝背景上更亮一些,提升对比度 */
  color: var(--text-strong, #f8fafc);
}
.cn-empty-line {
  margin: 0;
  font-size: 12px;
  color: #94a3b8;
}

/* ========== row_diff 标签卡 ========== */
.row-diff-tags {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin-bottom: 12px;
}
.rd-tag {
  border-radius: 8px;
  padding: 14px;
  text-align: center;
  border: 1px solid var(--border-cyan-soft);
  background: rgba(10, 20, 48, 0.55);
}
.rd-tag .tag-label {
  font-size: 12px;
  color: #94a3b8;
  margin-bottom: 4px;
}
.rd-tag .tag-num {
  font-size: 26px;
  font-weight: 700;
  color: var(--text-primary);
}
.rd-added { border-color: rgba(34, 197, 94, 0.5); background: rgba(34, 197, 94, 0.08); }
.rd-added .tag-num { color: #4ade80; }
.rd-removed { border-color: rgba(248, 113, 113, 0.5); background: rgba(248, 113, 113, 0.08); }
.rd-removed .tag-num { color: #f87171; }
.rd-modified { border-color: rgba(250, 204, 21, 0.5); background: rgba(250, 204, 21, 0.08); }
.rd-modified .tag-num { color: #fbbf24; }

/* ========== row_diff 修改明细 ========== */
.modified-row {
  margin-bottom: 10px;
  padding: 8px 10px;
  border: 1px dashed var(--border-cyan-soft);
  border-radius: 6px;
}
.modified-key {
  font-size: 12px;
  color: var(--accent-cyan-2);
  margin-bottom: 6px;
}
.cell-before { color: #f87171; text-decoration: line-through; }
.cell-after { color: #4ade80; font-weight: 600; }

/* ========== 业务对比·差值 / 变化幅度 列上色 ==========
   【topn 并入 business 新增】原先的 .change-up / .change-down 等
   样式属于已下线的 topn 独立模式,这里改为与业务对比表格配套的
   .diff-up / .diff-down / .diff-flat,语义统一为「涨绿跌红持平灰」 */
.diff-up { color: #4ade80; font-weight: 600; }
.diff-down { color: #f87171; font-weight: 600; }
.diff-flat { color: #94a3b8; }

/* ========== 【nl-chat-upgrade 新增】数据对话气泡流样式 ========== */
/* 消息列表容器:固定高 52vh,纵向排列气泡行,超出滚动 */
.nl-chat-list {
  height: 52vh;
  overflow-y: auto;
  padding: 4px 6px;
  display: flex;
  flex-direction: column;
  gap: 14px;
}
/* 空状态:居中展示提示文案 + 示例问题 */
.nl-chat-empty {
  margin: auto;
  text-align: center;
}
.nl-chat-empty-title {
  margin: 0 0 14px;
  font-size: 13px;
  color: var(--text-secondary);
  line-height: 1.8;
}
.nl-chat-examples {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: center;
  gap: 8px;
}
.nl-chat-examples-label {
  font-size: 12px;
  color: #94a3b8;
}
.nl-example-tag {
  cursor: pointer;
}
/* 气泡行:按角色左右对齐 */
.nl-chat-row {
  display: flex;
}
.nl-chat-row.nl-row-user {
  justify-content: flex-end;
}
.nl-chat-row.nl-row-ai {
  justify-content: flex-start;
}
/* 气泡基础:限宽 82% + 圆角,文本统一左对齐保证长句可读 */
.nl-bubble {
  max-width: 82%;
  padding: 10px 14px;
  border-radius: 10px;
  font-size: 13px;
  line-height: 1.7;
  word-break: break-word;
  text-align: left;
}
/* 用户气泡:主色青渐变底(与全局主按钮 el-button--primary 同款渐变与文字色),靠右,右上角收小 */
.nl-bubble-user {
  background: linear-gradient(135deg, var(--accent-cyan), var(--accent-cyan-2));
  color: #062a2e;
  border-top-right-radius: 2px;
}
/* AI 气泡:浅底(卡片底色) + 青色柔边,靠左,左上角收小 */
.nl-bubble-ai {
  background: rgba(20, 36, 78, 0.65);
  border: 1px solid var(--border-cyan-soft);
  color: var(--text-primary);
  border-top-left-radius: 2px;
}
/* AI 气泡内:结论高亮卡(浅色背景 + 左边框强调条) */
.nl-answer-card {
  padding: 8px 12px;
  background: rgba(34, 211, 238, 0.10);
  border-left: 3px solid var(--accent-cyan);
  border-radius: 4px;
  color: var(--text-strong);
  font-weight: 500;
  white-space: pre-wrap;
}
/* AI 气泡内:小节标题(分析/支撑数据/建议)与圆点列表 */
.nl-section {
  margin-top: 10px;
}
.nl-section-title {
  font-size: 12px;
  color: var(--accent-cyan-2);
  font-weight: 600;
  letter-spacing: 0.5px;
  margin-bottom: 4px;
}
.nl-points-list {
  margin: 0;
  padding-left: 18px;
}
.nl-points-list li {
  margin-bottom: 4px;
}
/* 支撑数据表 */
.nl-data-table {
  margin-top: 6px;
  width: 100%;
}
/* 追问 chips:横向一排,可点击 */
.nl-followups {
  margin-top: 10px;
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}
.nl-followup-tag {
  cursor: pointer;
}
/* 降级渲染的 explanation 文本(保留换行) */
.nl-fallback-text {
  margin: 0;
  white-space: pre-wrap;
}
/* 底部固定输入区:textarea + 勾选/发送一行 */
.nl-chat-input {
  margin-top: 12px;
}
.nl-chat-input-foot {
  margin-top: 8px;
  display: flex;
  align-items: center;
  justify-content: space-between;
}
</style>
