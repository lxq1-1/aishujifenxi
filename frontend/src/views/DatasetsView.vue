<!-- 数据集管理：统计卡片 + 列表 + 上传/预览/重命名/删除 -->
<template>
  <div class="page">
    <div class="page-header">
      <h2 class="h1-title" style="font-size: 22px">数据集管理</h2>
      <el-button type="primary" @click="uploadVisible = true">
        <el-icon><Upload /></el-icon> 上传数据集
      </el-button>
    </div>

    <!-- 统计卡片 -->
    <div class="stat-row">
      <div class="metric-card">
        <div class="metric-icon">📦</div>
        <div>
          <div class="metric-label">数据集总数</div>
          <div class="metric-value">{{ stats.total }}</div>
        </div>
      </div>
      <div class="metric-card">
        <div class="metric-icon">📅</div>
        <div>
          <div class="metric-label">本月新增</div>
          <div class="metric-value">{{ stats.month_new }}</div>
        </div>
      </div>
      <div class="metric-card">
        <div class="metric-icon">✓</div>
        <div>
          <div class="metric-label">已清洗</div>
          <div class="metric-value">{{ stats.cleaned }}</div>
        </div>
      </div>
      <div class="metric-card">
        <div class="metric-icon">⏳</div>
        <div>
          <div class="metric-label">待处理</div>
          <div class="metric-value">{{ stats.pending }}</div>
        </div>
      </div>
    </div>

    <!-- 列表 -->
    <el-card style="margin-top: 18px">
      <el-table :data="list" stripe>
        <el-table-column type="selection" width="48" />
        <el-table-column prop="name" label="数据集名称" />
        <el-table-column prop="file_name" label="文件名" />
        <!-- 【dataset-multiformat-upload 新增】类型列:用彩色标签区分
             结构化(Excel/CSV) / 文档(Word/PDF) / 图片(PNG/JPG)三类 -->
        <el-table-column label="类型" width="110">
          <template #default="{ row }">
            <el-tag :type="fileTypeTag(row.file_type).type" size="small" effect="plain">
              {{ fileTypeTag(row.file_type).icon }} {{ fileTypeTag(row.file_type).label }}
            </el-tag>
          </template>
        </el-table-column>
        <el-table-column :width="rowCountLabel.width" :label="rowCountLabel.text">
          <template #default="{ row }">
            <span v-if="row.file_type === 'structured'">{{ row.row_count }}</span>
            <span v-else-if="row.file_type" class="dim small">—</span>
            <span v-else>{{ row.row_count }}</span>
          </template>
        </el-table-column>
        <!-- 【dataset-multisheet-merge 新增】工作表数列:
             后端只在响应里有 sheets 字段时返回;老数据(没 sheets 字段)按 1 处理,
             不显示标签。el-table-column 用 v-if 控制可见性,避免无意义空白列 -->
        <el-table-column label="工作表" width="110">
          <template #default="{ row }">
            <el-tag
              v-if="getSheetCount(row) > 1"
              type="info"
              size="small"
              effect="plain"
            >
              📑 {{ getSheetCount(row) }} 个工作表
            </el-tag>
            <span v-else class="dim small">单表</span>
          </template>
        </el-table-column>
        <el-table-column prop="update_time" label="最后更新" width="140" />
        <el-table-column label="状态" width="120">
          <template #default="{ row }">
            <el-tag :type="row.status === '已清洗' ? 'success' : 'warning'">{{ row.status }}</el-tag>
            <el-tooltip
              v-if="row.file_type && row.file_type !== 'structured' && row.extract_status === 'failed'"
              content="内容抽取失败,文件已保存,可下载原文核对"
              placement="top"
            >
              <el-icon style="color:#f87171;margin-left:4px;vertical-align:middle"><Warning /></el-icon>
            </el-tooltip>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="300">
          <template #default="{ row }">
            <el-button link type="primary" @click="onPreview(row)">预览</el-button>
            <el-button link @click="onRename(row)">重命名</el-button>
            <el-button link @click="onUse(row)">使用</el-button>
            <el-button link type="primary" :loading="downloadingId === row.id" @click="onDownload(row)">下载</el-button>
            <el-button link type="danger" @click="onDelete(row)">删除</el-button>
          </template>
        </el-table-column>
        <template #empty>
          <div class="dim" style="padding: 20px">暂无数据集，请点击右上角上传</div>
        </template>
      </el-table>
    </el-card>

    <!-- 上传对话框 -->
    <el-dialog v-model="uploadVisible" title="上传数据集" width="520">
      <el-upload
        drag
        multiple
        :auto-upload="false"
        :on-change="onFileChange"
        :on-remove="onFileRemove"
        :file-list="fileList"
        accept=".xlsx,.xls,.csv,.docx,.doc,.pdf,.txt,.md,.png,.jpg,.jpeg,.webp,.bmp"
      >
        <el-icon style="font-size: 48px; color: var(--accent-cyan)"><UploadFilled /></el-icon>
        <div style="font-size:15px;margin-bottom:6px">拖入或点击选择文件</div>
        <div class="dim small" style="line-height:1.7">
          <div>📊 <b>表格</b>:.xlsx / .xls / .csv(自动读取所有工作表)</div>
          <div>📄 <b>文档</b>:.docx / .pdf / .txt / .md(自动抽取文字)</div>
          <div>🖼️ <b>图片</b>:.png / .jpg / .jpeg / .webp / .bmp(多模态识别内容)</div>
        </div>
      </el-upload>
      <template #footer>
        <el-button @click="uploadVisible = false">取消</el-button>
        <el-button type="primary" :loading="uploading" @click="onUpload">开始上传</el-button>
      </template>
    </el-dialog>

    <!-- 预览抽屉 -->
    <el-drawer v-model="previewVisible" title="数据集预览" size="640">
      <div v-if="preview">
        <!-- 【dataset-multiformat-upload 新增】类型分流:
             structured → 表格预览(行/列/sheet/质量);
             document/image → 文本预览(抽取出的文字 + 字符数 + 抽取状态) -->
        <template v-if="(preview.file_type || 'structured') === 'structured'">
          <div class="preview-stats">
            <div class="metric-card">
              <div class="metric-icon">📊</div>
              <div>
                <div class="metric-label">总行数</div>
                <div class="metric-value">{{ preview.total }}</div>
              </div>
            </div>
            <div class="metric-card">
              <div class="metric-icon">📋</div>
              <div>
                <div class="metric-label">字段数</div>
                <div class="metric-value">{{ preview.columns.length }}</div>
              </div>
            </div>
            <div class="metric-card">
              <div class="metric-icon">⚠️</div>
              <div>
                <div class="metric-label">缺失值</div>
                <div class="metric-value">{{ preview.quality?.missing_total || 0 }}</div>
              </div>
            </div>
          </div>

          <!-- 【dataset-multisheet-merge 新增】工作表明细:
               当 sheets 存在且 length > 1 时,展示每张表的行/列数,
               让用户直观看到"这份文件里其实有 N 张表,数据全部进系统了"。
               仅显示不点击;预览表格按合并后 df 取前 10 行 -->
          <div v-if="preview.sheets && preview.sheets.length > 1" class="sheet-list">
            <div class="sheet-list-title">
              📑 本文件包含 {{ preview.sheets.length }} 个工作表(已全部纳入分析)
            </div>
            <div class="sheet-grid">
              <div
                v-for="(s, idx) in preview.sheets"
                :key="idx"
                class="sheet-chip"
              >
                <span class="sheet-name">{{ s.name }}</span>
                <span class="sheet-meta">{{ s.rows }} 行 · {{ s.cols }} 列</span>
              </div>
            </div>
          </div>

          <!--
            【workbench-per-sheet-preview 新增】预览抽屉里的 sheet 切换 tab。
            - 多 sheet 文件(≥2)显示此 tab,与工作台一致;
            - 默认选中**第一张 sheet**(由后端在不带 sheet 参数时自动返回);
            - 不再有"全部合并"tab——合并大表的 11701 个跨表 NaN 严重误导分析,
              用户应看到真实的单 sheet 数据;AI 分析链路另走 read_dataframe_all_sheets。
            - 点具体 sheet tab → onPreviewSheetChange → 重拉预览 → 表格换成该 sheet
              真实的前 10 行 + 真实缺失值。
          -->
          <el-tabs
            v-if="preview.sheets && preview.sheets.length > 1"
            v-model="previewActiveSheet"
            type="card"
            class="ds-preview-sheet-tabs"
            @tab-change="onPreviewSheetChange"
          >
            <el-tab-pane
              v-for="s in preview.sheets"
              :key="s.name"
              :name="s.name"
              :label="`${s.name} (${s.rows}行)`"
            />
          </el-tabs>

          <el-table :data="preview.rows" border size="small" style="margin-top: 14px">
            <el-table-column
              v-for="col in preview.columns"
              :key="col"
              :prop="col"
              :label="col"
              min-width="120"
            />
          </el-table>
          <div class="dim small" style="margin-top: 8px">仅显示前 10 行</div>
        </template>

        <!-- 非结构化文件预览:展示抽取出的文本 -->
        <template v-else>
          <div class="preview-stats">
            <div class="metric-card">
              <div class="metric-icon">{{ preview.file_type === 'image' ? '🖼️' : '📄' }}</div>
              <div>
                <div class="metric-label">文件类型</div>
                <div class="metric-value small">
                  {{ fileTypeTag(preview.file_type).label }}
                </div>
              </div>
            </div>
            <div class="metric-card">
              <div class="metric-icon">📝</div>
              <div>
                <div class="metric-label">文本字符数</div>
                <div class="metric-value small">
                  {{ preview.text_total_chars || 0 }}
                </div>
              </div>
            </div>
            <div class="metric-card">
              <div class="metric-icon">⚙️</div>
              <div>
                <div class="metric-label">抽取状态</div>
                <div class="metric-value small">
                  <el-tag :type="preview.extract_status === 'success' ? 'success' : 'danger'" size="small">
                    {{ preview.extract_status === 'success' ? '成功' : '失败' }}
                  </el-tag>
                </div>
              </div>
            </div>
          </div>

          <div class="text-preview-wrap">
            <div class="text-preview-title">📝 抽取内容预览</div>
            <pre v-if="preview.extracted_text" class="text-preview-body">{{ preview.extracted_text }}</pre>
            <div v-else class="dim" style="padding: 20px; text-align: center">
              {{ preview.extract_status === 'failed' ? '❌ 内容抽取失败' : '暂无可预览内容' }}
            </div>
            <div v-if="preview.text_total_chars > 2000" class="dim small" style="margin-top: 6px">
              仅显示前 2000 字,完整 {{ preview.text_total_chars }} 字
            </div>
          </div>
        </template>

        <div style="margin-top: 18px">
          <el-button type="primary" @click="onUseFromPreview">使用此数据集</el-button>
        </div>
      </div>
    </el-drawer>
  </div>
</template>

<script setup>
import { ref, onMounted, computed } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Upload, UploadFilled, Warning } from '@element-plus/icons-vue'
import { datasetsApi } from '@/api'
import { useWorkbenchStore } from '@/store/workbench'

const router = useRouter()
const wbStore = useWorkbenchStore()

const list = ref([])
const stats = ref({ total: 0, month_new: 0, cleaned: 0, pending: 0 })
const uploadVisible = ref(false)
const uploading = ref(false)
const fileList = ref([])
// 真正要上传的文件列表(el-upload 把 file 一次性给我们的是变化项,
// 全量要从 on-change 的第二个参数 fileList 里取出 raw 拿出来用)
const rawFiles = ref([])
const previewVisible = ref(false)
const preview = ref(null)
const currentPreviewId = ref(null)
// 【workbench-per-sheet-preview 新增】预览抽屉里"当前激活的 sheet";
// '' 表示走"全部 sheet 合并预览"(默认,与工作台语义一致)。
const previewActiveSheet = ref('')

async function refresh() {
  const data = await datasetsApi.list()
  // 【dataset-deleted-not-selectable 修改】内部锚点 dataset(status="PDF工具产物")
  // 已由后端 GET /api/datasets 在**查询层**过滤掉(见 backend/app/api/datasets.py),
  // 接口返回的就是用户可见的数据集全集,此处不再重复过滤,避免两端口径不一致。
  // stats 与列表同源,因此"数据集总数"与列表行数天然对齐。
  list.value = data.items || []
  stats.value = {
    total: data.total,
    month_new: data.month_new,
    cleaned: data.cleaned,
    pending: data.pending,
  }
}

// 【dataset-multisheet-merge 新增】从行里读 sheet 数;
// 列表接口(/datasets)目前不带 sheets 字段,默认 1;如未来扩展,可在此统一取
function getSheetCount(row) {
  if (!row) return 1
  if (Array.isArray(row.sheets) && row.sheets.length) return row.sheets.length
  if (typeof row.active_sheet_count === 'number' && row.active_sheet_count > 0) {
    return row.active_sheet_count
  }
  return 1
}

// 【dataset-multiformat-upload 新增】文件类型 → 标签视觉
// 三类配色与图标保持统一,模板多处复用(列表、预览抽屉、对比弹窗)
// 未识别类型(老数据/兜底)显示 "未知",避免空白
const _FILE_TYPE_META = {
  structured: { type: 'success', icon: '📊', label: '表格' },
  document:   { type: 'info',    icon: '📄', label: '文档' },
  image:      { type: 'warning', icon: '🖼️', label: '图片' },
}
function fileTypeTag(t) {
  return _FILE_TYPE_META[t] || { type: '', icon: '❓', label: '未知' }
}

// 【dataset-multiformat-upload 新增】行数列自适应表头:
// 只在存在非结构化文件时把表头改成 "规模",否则保持 "行数"
const rowCountLabel = computed(() => {
  const hasNonStructured = (list.value || []).some(
    (d) => d.file_type && d.file_type !== 'structured',
  )
  return hasNonStructured
    ? { text: '规模', width: 100 }
    : { text: '行数', width: 100 }
})

// el-upload 的 @change 事件每次文件列表变化都会触发一次,
// 入参 file 仅代表"本次变化的那一个文件",不是"全部已选文件"。
// 因此用 currentFile 单变量会被反复覆盖,导致多选时只剩最后一个被上传。
// 这里改为响应 file-list 的完整数组,存到 rawFiles,真正上传时逐个走接口。
function onFileChange(_file, fileListAll) {
  rawFiles.value = (fileListAll || []).map((f) => f.raw).filter(Boolean)
}

// 同步移除事件:用户在 el-upload 内点 × 删单个文件时,rawFiles 也得跟着剔除
function onFileRemove(file) {
  rawFiles.value = rawFiles.value.filter((r) => r !== file?.raw)
}

async function onUpload() {
  const queue = rawFiles.value.slice() // 拷贝,避免后面清空影响
  if (queue.length === 0) {
    ElMessage.warning('请选择文件')
    return
  }
  uploading.value = true
  try {
    // 串行上传:每发一份就拿到一份结果,集中提示;保证一张失败不影响其他
    const results = []
    for (let i = 0; i < queue.length; i++) {
      const f = queue[i]
      try {
        const fd = new FormData()
        fd.append('file', f)
        const resp = await datasetsApi.upload(fd)
        results.push({ ok: true, name: f.name, resp })
      } catch (e) {
        results.push({ ok: false, name: f.name, error: e?.message || '上传失败' })
      }
    }

    // 汇总提示:M 成功 / N 失败,失败的逐条点名
    const successItems = results.filter((r) => r.ok)
    const failItems = results.filter((r) => !r.ok)
    if (failItems.length === 0) {
      const merged = successItems.filter(
        (r) => r.resp && r.resp.file_type === 'structured' && r.resp.sheets && r.resp.sheets.length > 1,
      )
      if (merged.length > 0) {
        ElMessage.success(`已成功上传 ${successItems.length} 份(其中 ${merged.length} 份含多 sheet,已自动合并)`)
      } else {
        ElMessage.success(`已成功上传 ${successItems.length} 份`)
      }
    } else {
      ElMessage.warning(
        `上传完成:${successItems.length} 成功,${failItems.length} 失败。失败:${failItems.map((x) => x.name).join('、')}`,
      )
    }

    // 仅当全部成功才关闭弹窗,失败的话保留窗口让用户决定是否重试剩余文件
    if (failItems.length === 0) {
      uploadVisible.value = false
      fileList.value = []
      rawFiles.value = []
    } else {
      const failedNames = new Set(failItems.map((x) => x.name))
      fileList.value = fileList.value.filter((f) => !failedNames.has(f.name))
      rawFiles.value = rawFiles.value.filter((f) => !failedNames.has(f.name))
    }
    await refresh()
  } finally {
    uploading.value = false
  }
}

async function onPreview(row) {
  currentPreviewId.value = row.id
  previewActiveSheet.value = ''   // 【workbench-per-sheet-preview 新增】每次打开抽屉重置 sheet 选中,后端会默认返第一张 sheet
  preview.value = await datasetsApi.preview(row.id)
  // 【workbench-per-sheet-preview 新增】把后端这次返回的"激活 sheet"(默认第一张)
  // 同步到本地 v-model,让 tab 高亮落在第一张 sheet 上,而不是空白
  if (preview.value && preview.value.active_sheet) {
    previewActiveSheet.value = preview.value.active_sheet
  }
  previewVisible.value = true
}

// 【workbench-per-sheet-preview 新增】预览抽屉里的 sheet tab 切换:
//   1) 写 previewActiveSheet;
//   2) 重拉 /preview?id=...&sheet=... 把新表的内容填进抽屉。
async function onPreviewSheetChange(name) {
  const target = name || ''
  previewActiveSheet.value = target
  if (!currentPreviewId.value) return
  preview.value = await datasetsApi.preview(currentPreviewId.value, { sheet: target || undefined })
}

async function onRename(row) {
  const { value } = await ElMessageBox.prompt('请输入新名称', '重命名', {
    inputValue: row.name,
    confirmButtonText: '确定',
    cancelButtonText: '取消',
  })
  if (value && value !== row.name) {
    await datasetsApi.rename(row.id, value)
    ElMessage.success('已重命名')
    await refresh()
  }
}

async function onDelete(row) {
  await ElMessageBox.confirm(`确定删除数据集「${row.name}」？`, '提示', { type: 'warning' })
  await datasetsApi.remove(row.id)
  ElMessage.success('已删除')
  await refresh()
}

// 当前正在下载的数据集 id,用于在按钮上显示 loading 状态
const downloadingId = ref(null)
async function onDownload(row) {
  downloadingId.value = row.id
  try {
    // 浏览器自身会弹保存框,无需 ElMessage;成功时静默完成
    await datasetsApi.download(row.id, row.name)
  } catch (e) {
    ElMessage.error(e?.message || '下载失败')
  } finally {
    downloadingId.value = null
  }
}

function onUse(row) {
  wbStore.setDataset(row)
  ElMessage.success('已选择数据集，请前往工作台')
  router.push('/workbench')
}

function onUseFromPreview() {
  const row = list.value.find((d) => d.id === currentPreviewId.value)
  if (row) onUse(row)
  previewVisible.value = false
}

onMounted(refresh)
</script>

<style scoped>
.page-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 18px;
}
.stat-row {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 14px;
}
.preview-stats {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
}
.small {
  font-size: 12px;
}

/* ==========【dataset-multisheet-merge 新增】预览抽屉·工作表明细 ========== */
.sheet-list {
  margin-top: 14px;
  padding: 12px 14px;
  background: rgba(10, 20, 48, 0.5);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
}
.sheet-list-title {
  font-size: 13px;
  font-weight: 600;
  color: var(--accent-cyan-2);
  margin-bottom: 8px;
  letter-spacing: 0.5px;
}
.sheet-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
  gap: 8px;
}
.sheet-chip {
  display: flex;
  flex-direction: column;
  gap: 2px;
  padding: 8px 10px;
  background: rgba(34, 211, 238, 0.08);
  border: 1px solid rgba(34, 211, 238, 0.25);
  border-radius: 6px;
}
.sheet-name {
  font-size: 13px;
  color: var(--text-primary);
  font-weight: 500;
  word-break: break-all;
}
.sheet-meta {
  font-size: 11px;
  color: #94a3b8;
}

/* ==========【dataset-multiformat-upload 新增】非结构化文件文本预览 ========== */
.text-preview-wrap {
  margin-top: 14px;
  padding: 12px 14px;
  background: rgba(10, 20, 48, 0.5);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
}
.text-preview-title {
  font-size: 13px;
  font-weight: 600;
  color: var(--accent-cyan-2);
  margin-bottom: 8px;
  letter-spacing: 0.5px;
}
.text-preview-body {
  margin: 0;
  padding: 12px;
  background: rgba(0, 0, 0, 0.25);
  border-radius: 6px;
  color: var(--text-primary);
  font-family: ui-monospace, SFMono-Regular, 'SF Mono', Consolas, monospace;
  font-size: 12px;
  line-height: 1.7;
  max-height: 480px;
  overflow: auto;
  white-space: pre-wrap;
  word-break: break-word;
}

/* 【workbench-per-sheet-preview 新增】预览抽屉里的 sheet tab 样式;
   与工作台 wb-preview-sheet-tabs 同源,选中态用 cyan 高亮,未选中灰底。 */
.ds-preview-sheet-tabs {
  margin: 12px 0 8px 0;
}
.ds-preview-sheet-tabs :deep(.el-tabs__header) {
  margin: 0;
}
.ds-preview-sheet-tabs :deep(.el-tabs__nav-wrap) {
  max-height: 36px;
}
.ds-preview-sheet-tabs :deep(.el-tabs__item) {
  font-size: 12px;
  padding: 0 10px !important;
  height: 30px;
  line-height: 30px;
}
.ds-preview-sheet-tabs :deep(.el-tabs__item.is-active) {
  color: #22d3ee !important;
  border-color: #22d3ee !important;
}
</style>
