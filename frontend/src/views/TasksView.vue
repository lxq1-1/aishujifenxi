<!-- 历史任务：表格 + 状态筛选 + 批量删除 -->
<template>
  <div class="page">
    <div class="page-header">
      <h2 class="h1-title" style="font-size: 22px">历史任务</h2>
      <div>
        <el-select v-model="statusFilter" placeholder="状态筛选" clearable style="width: 160px; margin-right: 12px">
          <el-option label="处理中" value="processing" />
          <el-option label="已完成" value="success" />
          <el-option label="失败" value="failed" />
        </el-select>
        <el-button type="danger" :disabled="!selection.length" @click="onBatchDelete">
          批量删除 ({{ selection.length }})
        </el-button>
      </div>
    </div>

    <el-card>
      <!-- 【修复】row-key="id" + reserve-selection：
           轮询刷新会整体替换 list 数组（产生全新对象），Element Plus 默认按对象引用追踪勾选，
           引用一丢勾选就清空。改用 id 作为行唯一标识后，刷新也能保留勾选状态 -->
      <el-table
        :data="filtered"
        stripe
        row-key="id"
        @selection-change="(rows) => (selection = rows)"
      >
        <el-table-column type="selection" width="48" reserve-selection />
        <el-table-column prop="id" label="ID" width="64" />
        <el-table-column prop="name" label="任务名称">
          <!-- 【report-feedback-and-versioning 新增】任务名后面跟一个版本角标(v2/v3)
               + 父任务 id 提示(便于从列表一眼看出 lineage 关系) -->
          <template #default="{ row }">
            <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap">
              <span>{{ row.name }}</span>
              <el-tag
                v-if="row.version_label"
                size="small"
                type="success"
                effect="plain"
              >
                {{ row.version_label }}
              </el-tag>
              <el-tag
                v-if="row.parent_task_id"
                size="small"
                type="info"
                effect="plain"
              >
                基于 #{{ row.parent_task_id }}
              </el-tag>
              <!-- 【tools-results-history-persistence 新增】3 种工具角标 -->
              <el-tag
                v-if="row.template_category === '数据校验'"
                size="small"
                type="warning"
                effect="plain"
              >
                📊 数据校验
              </el-tag>
              <el-tag
                v-if="row.template_category === '数据问答'"
                size="small"
                type="info"
                effect="plain"
              >
                💬 数据问答
              </el-tag>
              <el-tag
                v-if="row.template_category === '数据对比'"
                size="small"
                type="warning"
                effect="plain"
              >
                📊 数据对比
              </el-tag>
            </div>
            <!-- 列出 feedback 摘要(截断 60 字),便于看到迭代脉络 -->
            <div
              v-if="row.feedback_text"
              class="dim small"
              style="margin-top: 4px; line-height: 1.4"
            >
              📝 {{ row.feedback_text.length > 60 ? row.feedback_text.slice(0, 60) + '…' : row.feedback_text }}
            </div>
          </template>
        </el-table-column>
        <el-table-column label="状态" width="120">
          <template #default="{ row }">
            <el-tag :type="tagType(row.status)">{{ statusLabel(row.status) }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="create_time" label="创建时间" width="180" />
        <el-table-column label="操作" width="240">
          <template #default="{ row }">
            <el-button link type="primary" :disabled="row.status !== 'success'" @click="$router.push(`/reports/${row.id}`)">
              查看报告
            </el-button>
            <el-button link @click="onReanalyze(row)">重新分析</el-button>
            <el-button link type="danger" @click="onDelete(row)">删除</el-button>
          </template>
        </el-table-column>
        <template #empty>
          <div class="dim" style="padding: 24px">暂无任务，请前往工作台执行分析</div>
        </template>
      </el-table>
    </el-card>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, onUnmounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { tasksApi } from '@/api'

const list = ref([])
const statusFilter = ref('')
const selection = ref([])
let timer = null

const filtered = computed(() =>
  statusFilter.value ? list.value.filter((t) => t.status === statusFilter.value) : list.value
)

async function refresh() {
  const data = await tasksApi.list()
  list.value = data.items
}

function statusLabel(s) {
  return { processing: '处理中', success: '已完成', failed: '失败' }[s] || s
}
function tagType(s) {
  return { processing: 'warning', success: 'success', failed: 'danger' }[s] || 'info'
}

async function onDelete(row) {
  await ElMessageBox.confirm(`确定删除任务「${row.name}」？`, '提示', { type: 'warning' })
  await tasksApi.remove(row.id)
  ElMessage.success('已删除')
  await refresh()
}

async function onBatchDelete() {
  await ElMessageBox.confirm(`确定删除选中的 ${selection.value.length} 个任务？`, '提示', { type: 'warning' })
  await tasksApi.batchRemove(selection.value.map((r) => r.id))
  ElMessage.success('已删除')
  selection.value = []
  await refresh()
}

async function onReanalyze(row) {
  await tasksApi.reanalyze(row.id)
  ElMessage.success('已重新加入分析队列')
  await refresh()
}

onMounted(() => {
  refresh()
  // 处理中时每 2s 轮询刷新
  timer = setInterval(() => {
    if (list.value.some((t) => t.status === 'processing')) refresh()
  }, 2000)
})
onUnmounted(() => clearInterval(timer))
</script>

<style scoped>
.page-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 18px;
}
</style>
