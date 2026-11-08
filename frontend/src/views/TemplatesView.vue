<!-- 模板管理：左侧分类 + 中部网格(支持新建/编辑/复制/删除) + 右侧预览 -->
<template>
  <div class="page tpl-page">
    <!-- 左侧分类 -->
    <aside class="tpl-side">
      <div class="tpl-section-title">模板分类</div>
      <div
        v-for="cat in categories"
        :key="cat.value"
        class="tpl-cat"
        :class="{ active: activeCat === cat.value }"
        @click="activeCat = cat.value"
      >
        {{ cat.label }}
      </div>
    </aside>

    <!-- 中部模板网格 -->
    <section class="tpl-main">
      <div class="tpl-header">
        <h2 class="h1-title" style="font-size: 22px">报告模板管理</h2>
        <!-- 顶部"新建模板"按钮可用,打开抽屉 -->
        <el-button type="primary" @click="openCreate">
          <el-icon><Plus /></el-icon> 新建模板
        </el-button>
      </div>

      <div class="tpl-grid">
        <div
          v-for="tpl in filtered"
          :key="tpl.id"
          class="cyan-card tpl-card"
          :class="{ active: selected?.id === tpl.id }"
          @click="select(tpl)"
        >
          <!-- 卡片右上角操作按钮(hover 显示) -->
          <div class="tpl-actions" @click.stop>
            <!-- 内置模板:仅显示 编辑 / 复制 -->
            <template v-if="tpl.is_builtin === 1">
              <el-button size="small" link @click="openEdit(tpl)">编辑</el-button>
              <el-button size="small" link @click="onDuplicate(tpl)">复制</el-button>
            </template>
            <!-- 自建模板:编辑 / 复制 / 删除 -->
            <template v-else>
              <el-button size="small" link @click="openEdit(tpl)">编辑</el-button>
              <el-button size="small" link @click="onDuplicate(tpl)">复制</el-button>
              <el-popconfirm
                :title="`确定删除模板「${tpl.name}」吗?`"
                confirm-button-text="删除"
                cancel-button-text="取消"
                @confirm="onRemove(tpl)"
              >
                <template #reference>
                  <el-button size="small" link type="danger" @click.stop>删除</el-button>
                </template>
              </el-popconfirm>
            </template>
          </div>

          <div class="cyan-title tpl-name">
            {{ tpl.name }}
            <!-- 内置模板标识 -->
            <el-tag v-if="tpl.is_builtin === 1" size="small" type="info" style="margin-left: 8px">内置</el-tag>
          </div>
          <div class="dim small tpl-desc">{{ tpl.description }}</div>
        </div>

        <!-- 空状态 -->
        <div v-if="filtered.length === 0" class="tpl-empty">
          当前分类暂无模板,点击右上角"新建模板"创建一个。
        </div>
      </div>
    </section>

    <!-- 右侧预览 -->
    <aside class="tpl-preview" v-if="selected">
      <div class="tpl-section-title">{{ selected.name }} 预览</div>
      <pre class="structure">{{ selected.structure }}</pre>
      <div class="tpl-section-title">结构</div>
      <div class="kv">
        <div class="kv-row">
          <span class="dim">输入：</span>
          <el-tag size="small">统计指标</el-tag>
        </div>
        <div class="kv-row">
          <span class="dim">输出：</span>
          <el-tag size="small">四段式分析报告</el-tag>
        </div>
      </div>
      <el-button type="primary" style="margin-top: 18px; width: 100%" @click="onUse">
        使用此模板
      </el-button>
    </aside>

    <!-- 新建/编辑抽屉 -->
    <el-drawer
      v-model="drawerVisible"
      :title="drawerMode === 'create' ? '新建模板' : '编辑模板'"
      direction="rtl"
      size="540px"
      destroy-on-close
    >
      <el-form
        ref="formRef"
        :model="form"
        :rules="rules"
        label-width="90px"
        style="padding-right: 12px"
      >
        <el-form-item label="名称" prop="name">
          <el-input
            v-model="form.name"
            placeholder="例如:销售周报"
            maxlength="100"
            show-word-limit
          />
        </el-form-item>

        <el-form-item label="分类" prop="category">
          <!-- 分类下拉:allow-create 支持输入新分类,filterable 支持搜索 -->
          <el-select
            v-model="form.category"
            placeholder="选择已有分类或输入新分类回车"
            allow-create
            filterable
            default-first-option
            style="width: 100%"
          >
            <el-option
              v-for="c in categoryOptions"
              :key="c"
              :label="c"
              :value="c"
            />
          </el-select>
        </el-form-item>

        <el-form-item label="描述">
          <el-input
            v-model="form.description"
            type="textarea"
            :rows="3"
            placeholder="一句话说明模板用途"
            maxlength="500"
            show-word-limit
          />
        </el-form-item>

        <el-form-item label="报告结构">
          <el-input
            v-model="form.structure"
            type="textarea"
            :rows="10"
            placeholder="Markdown,支持章节标题,例如:&#10;一、数据概况&#10;二、成果亮点&#10;三、问题诊断&#10;四、优化建议"
          />
          <div class="dim small" style="margin-top: 4px">
            支持 Markdown 章节标题。
          </div>
        </el-form-item>
      </el-form>

      <template #footer>
        <el-button @click="drawerVisible = false">取消</el-button>
        <el-button type="primary" :loading="saving" @click="onSave">
          {{ drawerMode === 'create' ? '保存' : '更新' }}
        </el-button>
      </template>
    </el-drawer>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, reactive } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { Plus } from '@element-plus/icons-vue'
import { templatesApi } from '@/api'
import { useWorkbenchStore } from '@/store/workbench'

const router = useRouter()
const wbStore = useWorkbenchStore()

// 左侧分类标签,默认提供常用分类作为快筛入口
// 实际展示时会与后端 /categories 返回的「已有分类」合并去重,实现动态扩展
const DEFAULT_CATEGORIES = [
  { label: '年度复盘', value: '年度' },
  { label: '季度分析', value: '季度' },
  { label: '销售周报', value: '销售' },
  { label: '述职报告', value: '述职' },
  { label: '数据简报', value: '简报' },
]
// 响应式分类列表(默认 + 后端动态合并)
const categories = ref([...DEFAULT_CATEGORIES])
const activeCat = ref('年度')

const list = ref([])
const selected = ref(null)

// 抽屉表单状态
const drawerVisible = ref(false)
const drawerMode = ref('create') // 'create' | 'edit'
const editingId = ref(null)
const saving = ref(false)
const formRef = ref(null)
const categoryOptions = ref([]) // 后端拉到的已有分类

const emptyForm = () => ({
  name: '',
  category: '',
  description: '',
  structure: '',
})
const form = reactive(emptyForm())

// 表单校验规则:名称 / 分类必填且长度限制
const rules = {
  name: [
    { required: true, message: '请输入模板名称', trigger: 'blur' },
    { max: 100, message: '名称不能超过 100 个字符', trigger: 'blur' },
  ],
  category: [
    { required: true, message: '请选择或输入分类', trigger: 'change' },
    { max: 50, message: '分类不能超过 50 个字符', trigger: 'change' },
  ],
}

const filtered = computed(() => list.value.filter((t) => t.category === activeCat.value))

async function refresh() {
  // 列表:不传 category,获取全部(分类筛选由前端过滤)
  const data = await templatesApi.list()
  list.value = data.items
  if (!selected.value && filtered.value.length) {
    selected.value = filtered.value[0]
  }
}

/**
 * 从后端加载所有已用分类,并合并到:
 * 1) categoryOptions:抽屉表单里「分类」下拉的可选项
 * 2) categories:左侧快筛分类栏,实现「自定义分类自动出现在左侧」
 *
 * 合并规则:默认 4 项始终保留在前;其后追加后端返回的「未在内置中」的分类项,
 * 沿用 value 作为展示 label,保证中文可读。
 */
async function loadCategories() {
  try {
    const data = await templatesApi.categories()
    const items = data.items || []
    const defaultValues = DEFAULT_CATEGORIES.map((c) => c.value)

    // 1. 抽屉下拉:默认 + 后端分类合并去重
    categoryOptions.value = Array.from(new Set([...defaultValues, ...items]))

    // 2. 左侧分类栏:默认在前,后端新分类追加在后面
    const existingValues = new Set(defaultValues)
    const newOnes = items
      .filter((v) => v && !existingValues.has(v))
      .map((v) => ({ label: v, value: v }))
    categories.value = [...DEFAULT_CATEGORIES, ...newOnes]

    // 3. 兜底:若当前 activeCat 在新分类中不存在,回退到第一个默认分类
    if (!categories.value.some((c) => c.value === activeCat.value)) {
      activeCat.value = DEFAULT_CATEGORIES[0].value
    }
  } catch (e) {
    // 失败兜底:用内置固定分类
    categoryOptions.value = DEFAULT_CATEGORIES.map((c) => c.value)
    categories.value = [...DEFAULT_CATEGORIES]
  }
}

function select(tpl) {
  selected.value = tpl
}

function onUse() {
  if (!selected.value) return
  wbStore.setTemplate(selected.value)
  ElMessage.success(`已选择模板：${selected.value.name}，请前往工作台执行分析`)
  router.push('/workbench')
}

// ---- 新建 / 编辑 / 删除 / 复制 ----

function resetForm() {
  Object.assign(form, emptyForm())
}

function openCreate() {
  drawerMode.value = 'create'
  editingId.value = null
  resetForm()
  drawerVisible.value = true
}

function openEdit(tpl) {
  drawerMode.value = 'edit'
  editingId.value = tpl.id
  // 编辑时用模板当前值预填
  Object.assign(form, {
    name: tpl.name || '',
    category: tpl.category || '',
    description: tpl.description || '',
    structure: tpl.structure || '',
  })
  drawerVisible.value = true
}

async function onSave() {
  if (!formRef.value) return
  await formRef.value.validate(async (valid) => {
    if (!valid) return
    saving.value = true
    try {
      const payload = {
        name: form.name.trim(),
        category: form.category.trim(),
        description: form.description,
        structure: form.structure,
      }
      if (drawerMode.value === 'create') {
        const created = await templatesApi.create(payload)
        ElMessage.success(`已创建模板:${created.name}`)
      } else {
        const updated = await templatesApi.update(editingId.value, payload)
        ElMessage.success(`已更新模板:${updated.name}`)
      }
      drawerVisible.value = false
      await refresh()
      await loadCategories()
    } catch (err) {
      // 后端 4xx 错误由 axios 拦截器统一提示,这里仅做兜底
      const msg = err?.response?.data?.detail || err?.message || '保存失败'
      ElMessage.error(msg)
    } finally {
      saving.value = false
    }
  })
}

async function onRemove(tpl) {
  try {
    await templatesApi.remove(tpl.id)
    ElMessage.success(`已删除模板:${tpl.name}`)
    // 如果被删除的是当前选中,清空 selected
    if (selected.value && selected.value.id === tpl.id) {
      selected.value = null
    }
    await refresh()
  } catch (err) {
    const msg = err?.response?.data?.detail || err?.message || '删除失败'
    ElMessage.error(msg)
  }
}

async function onDuplicate(tpl) {
  try {
    const dup = await templatesApi.duplicate(tpl.id)
    ElMessage.success(`已复制为新模板:${dup.name}`)
    await refresh()
    await loadCategories()
  } catch (err) {
    const msg = err?.response?.data?.detail || err?.message || '复制失败'
    ElMessage.error(msg)
  }
}

onMounted(async () => {
  await loadCategories()
  await refresh()
})
</script>

<style scoped>
.tpl-page {
  display: grid;
  grid-template-columns: 200px 1fr 320px;
  gap: 18px;
  align-items: start;
}
.tpl-side,
.tpl-preview {
  background: var(--bg-card);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 12px;
  padding: 16px;
}
.tpl-section-title {
  color: var(--text-secondary);
  font-size: 12px;
  letter-spacing: 1px;
  margin: 8px 0 10px;
}
.tpl-cat {
  padding: 10px 12px;
  border-radius: 8px;
  cursor: pointer;
  color: var(--text-primary);
  margin-bottom: 4px;
}
.tpl-cat:hover {
  background: rgba(34, 211, 238, 0.08);
}
.tpl-cat.active {
  background: linear-gradient(135deg, var(--accent-cyan), var(--accent-cyan-2));
  color: #062a2e;
  font-weight: 600;
}
.tpl-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 18px;
}
.tpl-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 16px;
}
.tpl-card {
  position: relative;
  min-height: 160px;
  cursor: pointer;
  transition: transform 0.2s;
}
.tpl-card:hover {
  transform: translateY(-3px);
}
.tpl-card.active {
  border-color: var(--accent-cyan);
  background: linear-gradient(145deg, rgba(34, 211, 238, 0.2), rgba(94, 234, 212, 0.1));
}
/* 卡片右上角操作按钮组,默认透明,hover 时显示 */
.tpl-actions {
  position: absolute;
  top: 8px;
  right: 8px;
  display: flex;
  gap: 4px;
  opacity: 0;
  transition: opacity 0.2s;
}
.tpl-card:hover .tpl-actions {
  opacity: 1;
}
.tpl-name {
  font-size: 18px;
  margin-bottom: 8px;
  padding-right: 100px; /* 给右上角操作留位置 */
}
.tpl-desc {
  line-height: 1.6;
}
.tpl-empty {
  grid-column: 1 / -1;
  text-align: center;
  padding: 40px 0;
  color: var(--text-secondary);
  font-size: 14px;
}
.small {
  font-size: 12px;
}
.structure {
  background: rgba(10, 20, 48, 0.6);
  border: 1px solid var(--border-cyan-soft);
  border-radius: 8px;
  padding: 12px;
  color: var(--accent-cyan-2);
  white-space: pre-wrap;
  font-family: inherit;
}
.kv-row {
  margin: 6px 0;
}
</style>
