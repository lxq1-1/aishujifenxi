/**
 * 路由配置
 * 6 个页面：首页、工作台、数据集、模板、历史任务、报告预览
 */
import { createRouter, createWebHistory } from 'vue-router'

const routes = [
  { path: '/', name: 'home', component: () => import('@/views/HomeView.vue') },
  { path: '/workbench', name: 'workbench', component: () => import('@/views/WorkbenchView.vue') },
  { path: '/data-datasets', name: 'datasets', component: () => import('@/views/DatasetsView.vue') },
  { path: '/templates', name: 'templates', component: () => import('@/views/TemplatesView.vue') },
  { path: '/tasks', name: 'tasks', component: () => import('@/views/TasksView.vue') },
  { path: '/reports/:id', name: 'report', component: () => import('@/views/ReportView.vue') },
]

export default createRouter({
  history: createWebHistory(),
  routes,
})
