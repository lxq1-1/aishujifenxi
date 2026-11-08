<!-- 首页：产品介绍 + 3 卡片亮点 + 4 步流程 + CTA -->
<template>
  <div class="home">
    <!-- 主视觉 -->
    <section class="hero">
      <h1 class="hero-title">AI 数据处理工作台</h1>
      <p class="hero-sub">处理 Excel 台账，生成分析报告与述职材料</p>
      <p class="hero-desc">上传表格，自动清洗数据、计算指标、生成图表和分析文案</p>
      <el-button type="primary" size="large" round @click="$router.push('/workbench')">
        立即进入工作台
      </el-button>
    </section>

    <!-- 3 卡片 -->
    <section class="feature-row">
      <div class="cyan-card feature">
        <div class="feature-icon">📋</div>
        <div class="cyan-title">表格处理</div>
        <div class="dim small">清洗规整多份台账</div>
      </div>
      <div class="cyan-card feature">
        <div class="feature-icon">🧠</div>
        <div class="cyan-title">智能分析</div>
        <div class="dim small">AI 生成年度报告</div>
      </div>
      <div class="cyan-card feature">
        <div class="feature-icon">📤</div>
        <div class="cyan-title">报告输出</div>
        <div class="dim small">一键复制汇报材料</div>
      </div>
    </section>

    <!-- 4 步流程 -->
    <section class="flow">
      <h2 class="flow-title">从表格到报告</h2>
      <div class="flow-steps">
        <span class="flow-step">上传数据</span>
        <span class="flow-arrow">→</span>
        <span class="flow-step active">清洗计算</span>
        <span class="flow-arrow">→</span>
        <span class="flow-step">AI 生成报告</span>
        <span class="flow-arrow">→</span>
        <span class="flow-step">导出结果</span>
      </div>
    </section>
  </div>
</template>

<script setup></script>

<style scoped>
.home {
  padding: 40px 32px 80px;
  max-width: 1200px;
  margin: 0 auto;
}
.hero {
  text-align: center;
  padding: 80px 0 40px;
}
.hero-title {
  font-size: 56px;
  font-weight: 800;
  background: linear-gradient(135deg, var(--accent-cyan), var(--accent-cyan-2));
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  margin: 0 0 12px;
}
.hero-sub {
  font-size: 22px;
  color: var(--text-strong);
  margin: 0 0 8px;
}
.hero-desc {
  color: var(--text-secondary);
  font-size: 14px;
  margin: 0 0 28px;
}
.feature-row {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 18px;
  margin: 40px 0;
}
.feature {
  text-align: center;
  padding: 28px 18px;
}
.feature-icon {
  font-size: 32px;
  margin-bottom: 12px;
}
.feature .cyan-title {
  font-size: 18px;
  margin-bottom: 6px;
}
.small {
  font-size: 12px;
}
.flow {
  text-align: center;
  margin-top: 40px;
}
.flow-title {
  color: var(--text-strong);
  font-size: 22px;
  margin-bottom: 24px;
}
</style>
