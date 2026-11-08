<!-- 顶层布局：顶部导航 + 路由出口（首页不带左侧栏） -->
<template>
  <!-- el-config-provider 全局下发 ElMessage 配置:所有提示统一 3 秒自动关闭
       (element-plus 的 messageConfig 只在 ConfigProvider 挂载时写入,所以必须包在这里) -->
  <el-config-provider :message="{ duration: 3000 }">
    <div class="app-layout">
      <header class="top-bar">
        <div class="brand">
          <div class="logo">A</div>
          <span>AI数据处理工作台</span>
        </div>
        <nav class="nav-links">
          <router-link to="/">首页</router-link>
          <router-link to="/workbench">工作台</router-link>
          <router-link to="/data-datasets">数据集</router-link>
          <router-link to="/templates">模板</router-link>
          <router-link to="/tasks">历史任务</router-link>
        </nav>
        <div class="top-actions">
          <el-button type="primary" @click="$router.push('/workbench')">进入工作台</el-button>
        </div>
      </header>
      <main>
        <router-view v-slot="{ Component }">
          <component :is="Component" />
        </router-view>
      </main>
    </div>
  </el-config-provider>
</template>

<script setup>
import { useRoute } from 'vue-router'
const route = useRoute()
</script>
