/**
 * 图表导出工具（可视化报表导出 · 前端离屏截图通道）
 * ============================================================
 * 用途：
 *   「可视化报表导出（Excel）」主路径需要**所见即所得**的高保真图表截图。
 *   页面上的图表实例不能复用（直接 setOption 会闪屏并污染用户正在看的图），
 *   因此本模块在**独立离屏容器**里渲染一份「浅色化」后的 option，
 *   用 `getDataURL` 导出透明底 PNG，由后端解码后嵌入 Excel 卡面。
 *
 * 为什么必须浅色化：
 *   深色风格（tech_advanced）的 option 是「深底浅字」，直接落到 Excel 的
 *   白底 / 浅灰卡面上文字不可读。因此截图前把深色 option 转成浅色可读版。
 *   浅色化口径（WCAG 相对亮度判定，L > 0.6 视为浅色）与后端 matplotlib
 *   兜底渲染**保持一致**，保证两条路径产出的图视觉一致。
 *
 * 三个导出函数：
 *   - lightenOption(option)      深色 option → 浅色可读 option（返回深拷贝）
 *   - captureChartPng(option)    单张图离屏截图 → dataURL 字符串
 *   - captureChartsPng(list)     批量串行截图 → [{ index, title, png_base64 }]
 */

import * as echarts from 'echarts/core'
import { use } from 'echarts/core'
import {
  BarChart,
  LineChart,
  PieChart,
  ScatterChart,
  RadarChart,
  HeatmapChart,
  FunnelChart,
} from 'echarts/charts'
import {
  TitleComponent,
  TooltipComponent,
  LegendComponent,
  GridComponent,
  VisualMapComponent,
} from 'echarts/components'
import { CanvasRenderer } from 'echarts/renderers'

// ============================================================
// 【B1.1】完整图型集注册
// ------------------------------------------------------------
// 后端 chart_renderer 产出的图型就这 7 种（bar / line / pie / scatter /
// radar / heatmap / funnel），漏注册任何一个都会渲染成空白图。
// 这里注册的是**完整集**，页面上的图表实例只注册了自己用到的类型，
// 两者互不影响（use 是幂等的全局注册）。
//
// 【为什么额外注册 VisualMapComponent】
//   ECharts 的 heatmap **必须**配合 visualMap 才能工作：
//   - 缺 visualMap 时，heatmap 的每个格子都会退化成同一种颜色，
//     值→颜色的映射彻底失效（实测整张图只剩一种颜色）；
//   - 开发模式下（NODE_ENV !== production）ECharts 会直接抛
//     "Heatmap must use with visualMap" 报错，整张图截不出来。
//   而 LLM 生成的 heatmap option（ai_service 的图型白名单里含 heatmap）
//   必然带 visualMap，所以这里一并注册，保证 heatmap 能真正出图。
// ============================================================
use([
  BarChart,
  LineChart,
  PieChart,
  ScatterChart,
  RadarChart,
  HeatmapChart,
  FunnelChart,
  CanvasRenderer,
  TitleComponent,
  TooltipComponent,
  LegendComponent,
  GridComponent,
  VisualMapComponent,
])

// ============================================================
// 浅色化口径常量（与后端 matplotlib 兜底渲染共用同一套口径）
// ============================================================

/** 浅色判定阈值：WCAG 相对亮度 L > 0.6 视为「浅色」 */
const LIGHT_LUMINANCE_THRESHOLD = 0.6

/** 通用文字色（替换浅色文字） */
const TEXT_COLOR = '#334155'
/** 标题文字色（替换浅色标题） */
const TITLE_TEXT_COLOR = '#0f172a'
/** 网格区域背景色：保持透明，透出 Excel 的浅灰卡面 */
const GRID_BACKGROUND_COLOR = 'transparent'
/** 网格边框色 */
const GRID_BORDER_COLOR = '#e2e8f0'
/** 坐标轴轴线色 */
const AXIS_LINE_COLOR = '#cbd5e1'
/** 坐标轴刻度标签色 */
const AXIS_LABEL_COLOR = '#475569'
/** 图例文字色 */
const LEGEND_TEXT_COLOR = '#334155'

/**
 * 深拷贝（结构化克隆）
 * ------------------------------------------------------------
 * 浅色化必须返回深拷贝：**不能污染调用方传入的 option**，
 * 因为页面上的图表用的就是同一个 option 对象，改了就「闪屏 + 样式跑偏」。
 *
 * 注意：不能用 JSON.parse(JSON.stringify()) —— option 里可能含有
 * `echarts.graphic.LinearGradient` 这类**类实例**（渐变对象），
 * JSON 序列化会把原型链抹平，导致渐变失效。这里采用「保原型克隆」：
 * 普通对象/数组递归拷贝，类实例保留原型并浅拷贝自身属性。
 *
 * @param {*} value 待克隆的值
 * @returns {*} 克隆后的值
 */
function deepClone(value) {
  if (value === null || typeof value !== 'object') return value
  // 数组：逐项递归拷贝
  if (Array.isArray(value)) return value.map((item) => deepClone(item))
  // 普通对象：递归拷贝每个 key
  if (value.constructor === Object) {
    const out = {}
    for (const key of Object.keys(value)) out[key] = deepClone(value[key])
    return out
  }
  // 其它类实例（如 LinearGradient）：保留原型 + 浅拷贝自身属性，
  // 既不破坏渐变对象的类型信息，也不会与原对象共享可变状态。
  const proto = Object.getPrototypeOf(value)
  const out = Object.create(proto)
  for (const key of Object.keys(value)) out[key] = deepClone(value[key])
  return out
}

/**
 * 十六进制颜色 → RGB 三元组
 * ------------------------------------------------------------
 * 仅支持 `#rgb` / `#rrggbb` 两种写法；其它格式（rgb(...)、rgba(...)、
 * 具名色、渐变对象等）一律返回 null，由调用方安全跳过。
 *
 * @param {*} color 颜色值
 * @returns {{r:number,g:number,b:number}|null} RGB 三元组，解析失败返回 null
 */
function hexToRgb(color) {
  if (typeof color !== 'string') return null
  const hex = color.trim().replace(/^#/, '')
  if (!/^[0-9a-fA-F]+$/.test(hex)) return null
  // #rgb 形式补全成 #rrggbb
  const full =
    hex.length === 3
      ? hex
          .split('')
          .map((c) => c + c)
          .join('')
      : hex
  if (full.length !== 6) return null
  return {
    r: parseInt(full.slice(0, 2), 16),
    g: parseInt(full.slice(2, 4), 16),
    b: parseInt(full.slice(4, 6), 16),
  }
}

/**
 * 计算 WCAG 相对亮度
 * ------------------------------------------------------------
 * 流程：sRGB(0~255) → 归一化(0~1) → 线性化 → 加权求和
 * 线性化公式（WCAG 2.1）：C <= 0.03928 ? C/12.92 : ((C+0.055)/1.055)^2.4
 * 加权：L = 0.2126R + 0.7152G + 0.0722B
 *
 * @param {{r:number,g:number,b:number}} rgb RGB 三元组
 * @returns {number} 相对亮度（0 ~ 1）
 */
function relativeLuminance({ r, g, b }) {
  const linear = (c) => {
    const s = c / 255
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4)
  }
  return 0.2126 * linear(r) + 0.7152 * linear(g) + 0.0722 * linear(b)
}

/**
 * 判断某个颜色是否为「浅色」
 * ------------------------------------------------------------
 * 判定口径（前后端一致）：hex → 相对亮度 L > 0.6 即视为浅色。
 * 非 hex 颜色（rgb(...)/transparent/渐变对象）无法判定，返回 false（按深色处理，
 * 即「保持原值」），保证不会抛异常、也不会误改非 hex 值。
 *
 * @param {*} color 颜色值
 * @returns {boolean} 是否为浅色
 */
function isLightColor(color) {
  const rgb = hexToRgb(color)
  if (!rgb) return false
  return relativeLuminance(rgb) > LIGHT_LUMINANCE_THRESHOLD
}

/**
 * 替换一个颜色字段（无条件替换口径）
 * ------------------------------------------------------------
 * 两条硬约束：
 *   1. **字段不存在就跳过**，不凭空造字段。
 *   2. **当前值不是 hex 就跳过**（rgb(...)/transparent/渐变对象等），
 *      避免把渐变对象覆盖成纯色、或破坏其它颜色语义。
 *
 * @param {Object} holder 承载该字段的对象
 * @param {string} key 字段名
 * @param {string} newColor 目标颜色
 */
function replaceColorIfHex(holder, key, newColor) {
  if (!holder || typeof holder !== 'object') return
  if (!Object.prototype.hasOwnProperty.call(holder, key)) return
  // 非 hex 值安全跳过（渐变对象等原样保留）
  if (!hexToRgb(holder[key])) return
  holder[key] = newColor
}

/**
 * 替换一个「浅色才替换」的颜色字段
 * ------------------------------------------------------------
 * 与 replaceColorIfHex 的差别：只有当前值是**浅色**（L > 0.6）时才替换；
 * 深色（如 business_pro 的 #1e3a8a）保留原值。
 *
 * @param {Object} holder 承载该字段的对象
 * @param {string} key 字段名
 * @param {string} newColor 目标颜色
 */
function replaceColorIfLight(holder, key, newColor) {
  if (!holder || typeof holder !== 'object') return
  if (!Object.prototype.hasOwnProperty.call(holder, key)) return
  if (!isLightColor(holder[key])) return
  holder[key] = newColor
}

/**
 * 把「可能是对象也可能是数组」的字段统一成数组处理
 * ------------------------------------------------------------
 * ECharts 的 `xAxis` / `yAxis` / `grid` / `legend` 都可以写成单对象或数组，
 * 两种形态都要处理到。
 *
 * @param {*} value 字段值
 * @returns {Array} 规范化后的数组（无法识别时返回空数组）
 */
function toArray(value) {
  if (Array.isArray(value)) return value
  if (value && typeof value === 'object') return [value]
  return []
}

/**
 * 浅色化：把深色风格的 option 转成浅底可读的 option
 * ============================================================
 * 处理规则表（与后端 matplotlib 兜底渲染共用口径）：
 *
 * | option 字段                          | 处理                                          |
 * | ------------------------------------ | --------------------------------------------- |
 * | textStyle.color                      | 若为浅色 → #334155                            |
 * | title.textStyle.color                | 若为浅色 → #0f172a；本来就是深色则保留原值    |
 * | grid.backgroundColor                 | → transparent                                  |
 * | grid.borderColor                     | → #e2e8f0                                      |
 * | xAxis/yAxis 的 axisLine.lineStyle.color | → #cbd5e1                                   |
 * | xAxis/yAxis 的 axisLabel.color       | → #475569                                      |
 * | legend.textStyle.color               | → #334155                                      |
 * | 顶层 color 数组 / series[].color     | 一律保留（风格色盘是风格识别度的主要载体）    |
 *
 * 实现要点：
 *   1. 返回**深拷贝**，绝不改动调用方传入的原 option（页面上的图不能受影响）。
 *   2. 字段不存在就跳过，**不凭空造字段**（例如 minimalist 的 axisLine 没有
 *      lineStyle，就不去补一个）。
 *   3. 非 hex 颜色值安全跳过，不抛异常。
 *
 * @param {Object} option 原始 ECharts option
 * @returns {Object} 浅色化后的 option 深拷贝
 */
export function lightenOption(option) {
  if (!option || typeof option !== 'object') return {}

  // 深拷贝：后续所有修改都只落在副本上，原对象保持不变
  const next = deepClone(option)

  // ---- 1) 顶层 textStyle.color：浅色文字 → 深色文字 ----
  replaceColorIfLight(next.textStyle, 'color', TEXT_COLOR)

  // ---- 1.5) 顶层 backgroundColor：统一置 transparent ----
  // 【必须有】tech_advanced 风格的部分 option 带顶层 `backgroundColor`（深蓝 #0a1a3a）。
  // 若不处理，截出来是一整块深蓝底：落到 Excel 的白卡面上会变成刺眼的深色块，
  // 浅色文字叠在其上依旧难读 —— 等于浅色化白做。
  // 置成 transparent 后，由 Excel 侧填充 #f8fafc 浅灰卡面兜底（见 dashboard_report）。
  // 注意只替换 hex 值，渐变/具名色安静跳过。
  replaceColorIfHex(next, 'backgroundColor', GRID_BACKGROUND_COLOR)

  // ---- 2) title.textStyle.color：浅色标题 → #0f172a；深色标题保留原值 ----
  // title 可能是对象 / 数组，也可能是字符串（此时没有 textStyle，跳过）
  toArray(next.title).forEach((title) => {
    if (!title || typeof title !== 'object') return
    replaceColorIfLight(title.textStyle, 'color', TITLE_TEXT_COLOR)
  })

  // ---- 3) grid：背景透明 + 浅色边框 ----
  toArray(next.grid).forEach((grid) => {
    if (!grid || typeof grid !== 'object') return
    replaceColorIfHex(grid, 'backgroundColor', GRID_BACKGROUND_COLOR)
    replaceColorIfHex(grid, 'borderColor', GRID_BORDER_COLOR)
  })

  // ---- 4) 坐标轴：轴线色 + 刻度标签色 ----
  // xAxis / yAxis 可能是对象，也可能是数组，两种都处理
  ;['xAxis', 'yAxis'].forEach((axisKey) => {
    toArray(next[axisKey]).forEach((axis) => {
      if (!axis || typeof axis !== 'object') return
      // axisLine.lineStyle.color（没有 lineStyle 就跳过，不凭空造）
      const axisLine = axis.axisLine
      if (axisLine && typeof axisLine === 'object') {
        replaceColorIfHex(axisLine.lineStyle, 'color', AXIS_LINE_COLOR)
      }
      // axisLabel.color
      replaceColorIfHex(axis.axisLabel, 'color', AXIS_LABEL_COLOR)
    })
  })

  // ---- 5) legend.textStyle.color ----
  toArray(next.legend).forEach((legend) => {
    if (!legend || typeof legend !== 'object') return
    replaceColorIfHex(legend.textStyle, 'color', LEGEND_TEXT_COLOR)
  })

  // ---- 6) 顶层 color 数组 / series[].color 一律保留 ----
  // 风格色盘（青绿 / 霓虹 / 深蓝…）在浅底上依旧醒目，是风格识别度的主要载体，
  // 因此这里**不做任何处理**，保持原样。

  return next
}

/**
 * 去掉 option 里的图内标题
 * ============================================================
 * 为什么必须去掉：
 *   Excel 报表里每张图外面已经有一张「卡片标题条」（分组名 · 图标题），
 *   如果 PNG 内部再画一遍标题，就会出现**同一个标题显示两次**，
 *   视觉冗余、且图内标题是小字，卡片标题是大字，看起来像没对齐。
 *
 *   参考的《经营看板》版式就是"卡片标题条 + 无内部标题的纯图"。
 *
 * 为什么在截图环节做、而不是让调用方各自处理：
 *   截图与后端 matplotlib 兜底渲染是两条独立路径，必须**同步**去掉标题，
 *   否则"有前端截图"和"无前端截图（后端兜底）"两种报表长得不一样。
 *   后端 `_chart_png` 里有对应的去标题处理，两边口径一致。
 *
 * @param {Object} option 已深拷贝的 option（原地修改）
 * @returns {Object} 同一个对象（便于链式调用）
 */
function stripOptionTitle(option) {
  if (!option || typeof option !== 'object') return option
  // title 可能是对象 / 数组 / 字符串，三种形态都要清掉
  if (Array.isArray(option.title)) {
    option.title = []
  } else if (option.title && typeof option.title === 'object') {
    option.title = { ...option.title, text: '' }
  } else if (typeof option.title === 'string') {
    option.title = ''
  }
  return option
}

/**
 * 离屏截图：把一份 option 渲染成透明底 PNG
 * ============================================================
 * 实现要点：
 *   1. 离屏容器放在 `position:fixed; left:-10000px` 处，视觉上不可见，
 *      且**必须保留宽高**（否则 ECharts 画布是 0 宽高，截出来全白）。
 *   2. 用 `echarts/core` 的 `init()` 起一个**独立实例**，不复用页面上的实例，
 *      所以不会闪屏、不会污染用户视图。
 *   3. `finally` 里必须 `dispose()` + 移除 DOM 容器：防止内存泄漏，
 *      也防止反复导出时堆积实例把浏览器拖垮。
 *   4. 整体 try/catch，失败返回 null，**不把异常抛给调用方**
 *      （导出流程不该因为单张图失败而整体崩溃，后端还有 matplotlib 兜底）。
 *   5. 去掉图内标题（标题由 Excel 卡片标题条承担，见 stripOptionTitle）。
 *
 * @param {Object} option ECharts option（会被浅色化 + 去标题后再渲染）
 * @param {Object} [options] 渲染参数
 * @param {number} [options.width=900] 画布宽（px）
 * @param {number} [options.height=560] 画布高（px）
 * @param {number} [options.pixelRatio=2] 像素密度倍数（2 = 2x 高清图）
 * @returns {Promise<string|null>} dataURL 字符串（含 `data:image/png;base64,` 前缀），失败返回 null
 */
export async function captureChartPng(
  option,
  { width = 900, height = 560, pixelRatio = 2 } = {}
) {
  // 空 option / 非对象：直接返回 null，不做无谓的渲染
  if (!option || typeof option !== 'object') return null

  let container = null
  let chart = null
  try {
    // ---- 1) 创建离屏容器 ----
    container = document.createElement('div')
    container.setAttribute('aria-hidden', 'true')
    // 必须带尺寸：ECharts 依赖容器宽高初始化画布，缺尺寸会得到 0x0 空白画布
    container.style.cssText =
      `position:fixed;left:-10000px;top:0;width:${width}px;height:${height}px;` +
      `pointer-events:none;`
    document.body.appendChild(container)

    // ---- 2) 起独立实例并渲染浅色化 + 去标题后的 option ----
    chart = echarts.init(container, null, { width, height })
    // 【关键】必须关闭动画再截图。
    // ECharts 默认开启入场动画（柱子的高度、饼图扇区半径、折线的绘制都是从 0 长起来的），
    // setOption 之后**立即** getDataURL 会抓到动画第 0 帧 —— 图形尺寸还是 0，
    // 截出来的图只有标题和坐标轴，饼图甚至几乎全白（实测非白像素仅 0.14%）。
    // 关掉动画让 setOption 同步渲染出最终形态。
    const renderOption = stripOptionTitle(lightenOption(option))
    chart.setOption({ ...renderOption, animation: false }, { notMerge: true })

    // ---- 3) 等两帧再取图，确保 canvas 已完成一次真实绘制 ----
    // 为什么还要等：setOption 只是把渲染任务入队，真正的绘制发生在下一帧。
    // 双 requestAnimationFrame 保证「入队 → 绘制」至少发生一轮；
    // 若浏览器不支持 rAF（极老环境）则退化为 0ms 定时器，不至于卡死。
    await new Promise((resolve) => {
      if (typeof requestAnimationFrame === 'function') {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
      } else {
        setTimeout(resolve, 16)
      }
    })

    // ---- 4) 导出透明底 PNG（透明底可叠加在 Excel 的浅灰卡面上）----
    const dataUrl = chart.getDataURL({
      type: 'png',
      pixelRatio,
      backgroundColor: 'transparent',
    })
    return dataUrl || null
  } catch (e) {
    // 截图失败不抛给调用方：交由后端 matplotlib 兜底渲染
    console.warn('[chartExport] 图表截图失败：', e)
    return null
  } finally {
    // ---- 4) 无论如何都要清理：释放画布内存 + 移除 DOM 节点 ----
    try {
      if (chart) chart.dispose()
    } catch (e) {
      /* dispose 失败不影响主流程 */
    }
    try {
      if (container && container.parentNode) container.parentNode.removeChild(container)
    } catch (e) {
      /* 移除 DOM 失败不影响主流程 */
    }
  }
}

/**
 * 批量截图：逐张串行渲染
 * ============================================================
 * - **串行**执行：同时创建多个离屏实例会造成明显卡顿（甚至内存暴涨），
 *   所以严格一张接一张。
 * - **单张失败不中断整批**：失败的项 `png_base64` 为 null，
 *   由后端用 matplotlib 兜底渲染，保证整份报表仍能导出。
 *
 * @param {Array<{title:string, option:Object}>} list 待截图的图表列表
 * @returns {Promise<Array<{index:number,title:string,png_base64:string|null}>>}
 *   每项含原数组下标 index、标题 title、纯 base64（**不含 dataURL 前缀**）png_base64
 */
export async function captureChartsPng(list) {
  const results = []
  if (!Array.isArray(list)) return results

  // 串行：for...of + await 保证上一张渲染完再渲染下一张
  for (let index = 0; index < list.length; index += 1) {
    const item = list[index] || {}
    const title = item.title || ''
    let pngBase64 = null

    try {
      const dataUrl = await captureChartPng(item.option)
      if (dataUrl) {
        // 去掉 `data:image/png;base64,` 前缀，方便后端直接 base64 解码
        pngBase64 = dataUrl.replace(/^data:image\/png;base64,/, '')
      }
    } catch (e) {
      // 单张失败不影响整批：该项留空，后端兜底
      console.warn(`[chartExport] 第 ${index} 张图表截图失败：`, e)
      pngBase64 = null
    }

    results.push({ index, title, png_base64: pngBase64 })
  }

  return results
}

export default {
  lightenOption,
  captureChartPng,
  captureChartsPng,
}
