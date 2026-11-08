/**
 * Axios 统一封装：baseURL、错误处理
 */
import axios from 'axios'
import { ElMessage } from 'element-plus'

const http = axios.create({
  baseURL: '/api',
  // 【2026-10 放宽】120s → 600s。
  // 原因:「照片转 PDF」的 OCR 是**逐图并发**调视觉模型(后端 max_workers=4),
  // 单张推理实测 3~150 秒(服务端负载波动极大),20 张图的批处理在极端情况下
  // 会超过 120s,前端会先于后端超时,表现为「按钮一直转圈然后失败」,
  // 但后端其实已经把 PDF 生成好了。
  // 真正的"卡住"由后端 180s 的单图超时兜底,这里给足余量避免前端提前放弃。
  timeout: 600000,
})

http.interceptors.response.use(
  (resp) => {
    // 文件下载(blob)响应保留完整 resp,调用方需要读 Content-Disposition 里的文件名;
    // 普通 JSON 响应仍只剥出 data,保持原调用习惯
    if (resp.config?.responseType === 'blob') return resp
    return resp.data
  },
  (err) => {
    const msg = err?.response?.data?.detail || err.message || '请求失败'
    ElMessage.error(msg)
    return Promise.reject(err)
  }
)

export default http
