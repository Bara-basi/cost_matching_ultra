<script setup>
import { computed, onMounted, onUnmounted, ref, watch } from 'vue'

const nav = [
  { id: 'home', icon: '◈', label: '工作台' },
  { id: 'lookup', icon: '⌕', label: '逐条核验' },
  { id: 'sync', icon: '⌁', label: '飞书区间' },
  { id: 'template', icon: '▤', label: '模板文件' },
  { id: 'jobs', icon: '◷', label: '任务记录' },
  { id: 'exceptions', icon: '◇', label: '异常记录' },
]
const screen = ref('home')
const theme = ref(localStorage.getItem('workbench-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'))
const motion = ref(localStorage.getItem('workbench-motion') !== 'off')
const operator = ref(localStorage.getItem('workbench-operator') || '')
const health = ref({})
const target = ref({})
const jobs = ref([])
const exceptions = ref([])
const exceptionVisible = ref(50)
const queueKind = ref('ready')
const submitting = ref('')
const progressView = ref(0)
const progressSettled = ref(false)
const flagReason = ref('')
const job = ref(null)
const rows = ref([])
const plan = ref(null)
const selection = ref(null)
const compareData = ref(null)
const evidence = ref([])
const evidenceComparison = ref(null)
const erpDetail = ref(null)
const audit = ref([])
const inputRows = ref([])
const onlyExceptions = ref(true)
const search = ref('')
const focused = ref(-1)
const toast = ref('')
const busy = ref('')
const modal = ref('')
const overrideAmount = ref('')
const overrideReason = ref('')
const editRows = ref([])
const editFocusIndex = ref(-1)
const showOtherInputs = ref(false)
const pasteText = ref('')
const lookupRows = ref([blankRow()])
const lookupWarnings = ref([])
const lookupFiles = ref([])
const templateFile = ref(null)
const templateCheck = ref(null)
const syncForm = ref({ start_date: '', end_date: '', declarations: '', contracts: '', limit: 50 })
const canvas = ref(null)
let timer, frame, resizeHandler, pointerHandler, progressFrame

function blankRow() { return { '报关单号': '', '合同号_1': '', '采购订单号': '', '商品序号': '', '报关品名': '', '海关编码': '', '报关金额': '', '币种': 'USD', '报关重量': '' } }
function money(value) { const n = Number(value); return Number.isFinite(n) ? n.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—' }
function shortDate(value) { const d = new Date(value); return Number.isNaN(d.getTime()) ? '近期' : `${d.getMonth()+1}月${d.getDate()}日` }
function rangeDate(value) { const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return match ? `${Number(match[2])}月${Number(match[3])}日` : '' }
function jobTitle(item) {
  const day = shortDate(item.createdAt)
  if (item.kind === 'sync') {
    const filters = item.selectionFilters || {}
    if (filters.start_date || filters.end_date) return `${day} · ${rangeDate(filters.start_date) || '起始'}至${rangeDate(filters.end_date) || '当前'}的飞书商品`
    const code = String(filters.declarations || '').split(/[,，\s]+/)[0]
    if (code) return `${day} · ${code}${String(filters.declarations).split(/[,，\s]+/).filter(Boolean).length > 1 ? '等' : ''}报关单`
    if (item.subject) return `${day} · ${item.subject}${item.declarationCount > 1 ? '等' : ''}报关单`
    return `${day} · 飞书未核算商品（${item.rowCount || 0}行）`
  }
  if (item.kind === 'files') return `${day} · 模板文件成本匹配`
  if (item.subject) return `${day} · ${item.subject}${item.declarationCount > 1 ? '等' : ''}报关单核验`
  return `${day} · ${item.rowCount || 0}行逐条核验`
}
function notify(message) { toast.value = message; setTimeout(() => { if (toast.value === message) toast.value = '' }, 4200) }
async function api(path, options = {}) {
  const response = await fetch(path, options)
  const content = response.headers.get('content-type') || ''
  const data = content.includes('json') ? await response.json() : await response.text()
  if (!response.ok) throw new Error(data?.detail || data?.message || `请求失败 (${response.status})`)
  return data
}
function form(data) { const body = new FormData(); Object.entries(data).forEach(([key, value]) => { if (value !== undefined && value !== null) body.append(key, value) }); return body }
async function action(name, run) { busy.value = name; try { return await run() } catch (error) { notify(error.message || '操作未完成') } finally { busy.value = '' } }
async function loadJobs() { const result = await api('/api/jobs'); jobs.value = result.jobs || [] }
async function loadExceptions() { exceptions.value = (await api('/api/exceptions')).rows || [] }
async function openQueue(kind) { queueKind.value = kind; exceptionVisible.value = 50; await loadJobs(); if (kind === 'exceptions') await loadExceptions(); screen.value = 'queue' }
const queueJobs = computed(() => jobs.value.filter(item => queueKind.value === 'ready'
  ? item.kind === 'sync' && item.writebackReady > 0 : item.state === 'complete'))
function startSubmission(kind) { submitting.value = kind; progressView.value = 0; progressSettled.value = false; screen.value = 'result'; job.value = null; rows.value = [] }
async function openJob(id, index = -1) {
  const wasSubmitting = Boolean(submitting.value)
  if (!wasSubmitting) { progressView.value = 100; progressSettled.value = true }
  screen.value = 'result'; focused.value = -1; onlyExceptions.value = true; search.value = ''
  // 切换任务时清空上一任务的预览态，避免写回预览 / 飞书对照 / 补齐标记串到本任务
  plan.value = null; compareData.value = null; evidenceComparison.value = null
  evidence.value = []; audit.value = []; inputRows.value = []; erpDetail.value = null
  job.value = await api(`/api/jobs/${id}`)
  submitting.value = ''
  await refreshJob()
  if (index >= 0) { focused.value = index; pulse.value = index; setTimeout(() => { pulse.value = -1 }, 1100) }
}
async function refreshJob() {
  if (!job.value) return
  job.value = await api(`/api/jobs/${job.value.id}`)
  if (job.value.state === 'complete') {
    rows.value = (await api(`/api/jobs/${job.value.id}/rows`)).rows || []
    if (job.value.kind === 'sync') plan.value = await api(`/api/sync/${job.value.id}/plan`).catch(() => null)
    evidence.value = (await api(`/api/jobs/${job.value.id}/evidence`).catch(() => ({ files: [] }))).files || []
    evidenceComparison.value = health.value.evidenceEnabled
      ? await api(`/api/jobs/${job.value.id}/evidence-comparison`).catch(() => null)
      : null
    audit.value = (await api(`/api/jobs/${job.value.id}/audit`).catch(() => ({ events: [] }))).events || []
    inputRows.value = (await api(`/api/jobs/${job.value.id}/input`).catch(() => ({ rows: [] }))).rows || []
    if (focused.value < 0) focusNext(true)
  }
}
function animateProgress() {
  if (submitting.value || (job.value && ['queued','running'].includes(job.value.state))) {
    const actual = Number(job.value?.progress || 0)
    const next = [8,20,45,70,92,99].find(value => value > actual) || 99
    const target = submitting.value ? 7.9 : Math.min(99, Math.max(actual, next - 0.35))
    const delta = target - progressView.value
    if (delta > 0.005) progressView.value = Math.min(target, progressView.value + Math.max(.01, delta * .025))
  } else if (job.value?.state === 'complete' && progressView.value < 100) {
    progressView.value = Math.min(100, progressView.value + Math.max(.08, (100 - progressView.value) * .09))
    if (progressView.value >= 99.99) { progressView.value = 100; progressSettled.value = true }
  }
  progressFrame = requestAnimationFrame(animateProgress)
}
function focusNext(initial = false) {
  const bad = rows.value.map((row, index) => ({ row, index })).filter(({ row }) => row['异常类型'] !== '正常')
  if (!bad.length) { focused.value = rows.value.length ? 0 : -1; return }
  const next = bad.find(({ index }) => index > focused.value) || bad[0]
  focused.value = next.index
  if (!initial || !sessionStorage.getItem(`focus-${job.value?.id}`)) {
    sessionStorage.setItem(`focus-${job.value?.id}`, 'yes')
    const el = document.getElementById(`result-${focused.value}`)
    setTimeout(() => el?.scrollIntoView({ behavior: motion.value ? 'smooth' : 'instant', block: 'center' }), 80)
    pulse.value = focused.value
    setTimeout(() => { pulse.value = -1 }, 1000)
  }
}
const pulse = ref(-1)
const filteredRows = computed(() => rows.value.map((row, index) => ({ row, index })).filter(({ row }) =>
  (!onlyExceptions.value || row['异常类型'] !== '正常') &&
  (!search.value || ['报关单号', '采购订单号', '供应商简称', '报关品名'].some(key => String(row[key] || '').includes(search.value)))))
const exceptionCount = computed(() => exceptions.value.length)
const readyCount = computed(() => jobs.value.filter(item => item.kind === 'sync' && item.writebackReady > 0).length)
const doneCount = computed(() => jobs.value.filter(item => item.state === 'complete').length)
const activeRow = computed(() => rows.value[focused.value] || null)
const activeSource = computed(() => plan.value?.groups?.find(group => group.children.some(child => child.index === focused.value)))
const contextIndexes = computed(() => new Set((plan.value?.groups || [])
  .filter(group => group.contextOnly).flatMap(group => group.children.map(child => child.index))))
const activePdf = computed(() => {
  const tokens = activeSource.value?.pdfTokens || []
  return evidence.value.find(item => tokens.some(token => item.name === `${token}.pdf`)) ||
    (evidence.value.length === 1 ? evidence.value[0] : null)
})

async function parsePaste() { await action('paste', async () => { const result = await api('/api/lookup/parse', { method: 'POST', body: form({ text: pasteText.value }) }); if (!result.rows?.length) return notify('没有识别到商品行，请手工填写'); if (result.rows.length > 50) throw new Error('识别超过 50 行，请分批粘贴'); lookupRows.value = result.rows.map(row => ({ ...blankRow(), ...row })); lookupWarnings.value = result.warnings || []; notify(`识别出 ${lookupRows.value.length} 条，请核对字段与补全来源`) }) }
async function enrichLookup() { await action('enrich', async () => { const result = await api('/api/lookup/enrich', { method: 'POST', body: form({ rows: JSON.stringify(lookupRows.value) }) }); lookupRows.value = result.rows.map(row => ({ ...blankRow(), ...row })); lookupWarnings.value = result.warnings || []; notify('已补全可唯一确认的字段，请核对提示') }) }
async function importLookup(event) { await action('import', async () => { const file = event.target.files?.[0]; if (!file) return; const result = await api('/api/lookup/import', { method: 'POST', body: form({ file }) }); if (result.errors?.length) return notify(`模板第 ${result.errors[0].row} 行：${result.errors[0].message}`); lookupRows.value = result.rows.map(row => ({ ...blankRow(), ...row })); notify(`已导入 ${lookupRows.value.length} 条，请核对后开始`) }) }
async function startLookup() { const valid = lookupRows.value.filter(row => row['报关单号'] || row['合同号_1']); if (!valid.length || valid.length > 50) return notify('请填写 1 至 50 条商品行'); startSubmission('lookup'); try { const body = form({ rows: JSON.stringify(valid) }); lookupFiles.value.forEach(file => body.append('files', file)); const result = await api('/api/lookup/with-evidence', { method: 'POST', body }); await openJob(result.id) } catch (error) { submitting.value = ''; screen.value = 'lookup'; notify(error.message || '提交未完成') } }
function chooseLookupFiles(event) { lookupFiles.value = [...event.target.files] }
function chooseTemplate(event) { templateFile.value = event.target.files?.[0] || null; templateCheck.value = null }
async function checkTemplate() { await action('check', async () => { if (!templateFile.value) throw new Error('请先选择模板文件'); const body = form({ file: templateFile.value }); templateCheck.value = await api('/api/template/check', { method: 'POST', body }); notify(templateCheck.value.errors.length ? '发现需要修正的单元格' : `检查通过：${templateCheck.value.rows} 行`) }) }
async function runTemplate() { await action('template', async () => { if (!templateFile.value) throw new Error('请先选择文件'); if (!templateCheck.value || templateCheck.value.errors.length) throw new Error('请先检查并修正模板'); const body = new FormData(); body.append('files', templateFile.value); const result = await api('/api/jobs', { method: 'POST', body }); await openJob(result.id) }) }
async function previewSync() { await action('preview', async () => { selection.value = await api('/api/sync/preview', { method: 'POST', body: form(syncForm.value) }) }) }
async function runSync() { if (!selection.value?.selection?.total) return; startSubmission('sync'); try { const result = await api('/api/sync/scan', { method: 'POST', body: form(syncForm.value) }); if (!result.id) { submitting.value = ''; screen.value = 'sync'; return notify(result.message || '没有可处理的记录') } await openJob(result.id) } catch (error) { submitting.value = ''; screen.value = 'sync'; notify(error.message || '读取未完成') } }
watch(syncForm, () => { selection.value = null }, { deep: true })
async function compare() { await action('compare', async () => { await api(`/api/jobs/${job.value.id}/feishu-comparison`, { method: 'POST' }); compareData.value = { state: 'running' }; const poll = async () => { const result = await api(`/api/jobs/${job.value.id}/feishu-comparison`); compareData.value = result; if (result.state === 'running') setTimeout(poll, 1800) }; poll() }) }
async function showErp() { await action('erp', async () => { erpDetail.value = await api(`/api/jobs/${job.value.id}/erp-evidence/${focused.value}`); modal.value = 'erp' }) }
async function markSample() { await action('sample', async () => { if (!operator.value.trim()) throw new Error('请先填写操作员称呼'); await api(`/api/jobs/${job.value.id}/reviewed`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ index: focused.value, operator: operator.value }) }); await refreshJob(); modal.value = ''; notify('已记录这条抽样复核') }) }
function compareStatus(row) {
  const candidates = compareData.value?.rows?.filter(item => item['报关单号'] === row['报关单号']) || []
  const contract = row['合同号（应收表格）'] || row['采购订单号'] || ''
  const same = candidates.find(item => item['合同号（应收表格）'] === contract) ||
    (candidates.length === 1 ? candidates[0] : null)
  if (!same || same['异常类型'] === '飞书没有这一行') return '暂无法比较'
  return same['核对状态'] === '正常' ? '结果一致' : '结果不一致'
}
function evidenceStatus(row) {
  const item = evidenceComparison.value?.rows?.find(item => item['报关单号'] === row['报关单号'])
  return item?.['状态'] || '暂无法比较'
}
async function saveOverride() { await action('override', async () => { if (!operator.value.trim()) throw new Error('请填写操作员称呼'); await api(`/api/jobs/${job.value.id}/override`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ index: focused.value, amount: overrideAmount.value, operator: operator.value, reason: overrideReason.value }) }); modal.value = ''; await refreshJob(); notify('调整已保存，写回前仍需通过整组校验') }) }
async function saveFlag() { await action('flag', async () => { await api(`/api/jobs/${job.value.id}/flag-exception`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ index: focused.value, operator: operator.value, reason: flagReason.value }) }); modal.value = ''; await refreshJob(); await loadJobs(); await loadExceptions(); pulse.value = focused.value; setTimeout(() => { pulse.value = -1 }, 1300); notify('已上报异常，相关记录已暂停写回') }) }
function openEdit() {
  if (!inputRows.value.length || !activeRow.value) return notify('此条没有可编辑的输入记录')
  const sourceId = activeSource.value?.sourceId
  const row = activeRow.value
  const matches = [
    item => sourceId && item.record_id === sourceId,
    item => item['报关单号'] === row['报关单号'] && item['合同号_1'] === row['合同号_1'] && item['报关品名'] === row['报关品名'],
    item => item['报关单号'] === row['报关单号'] && item['合同号_1'] === row['合同号_1'],
  ]
  const index = matches.map(match => inputRows.value.findIndex(match)).find(value => value >= 0) ?? -1
  if (index < 0) return notify('此条为补齐结果，没有对应的飞书输入行；请查看来源凭证')
  editRows.value = JSON.parse(JSON.stringify(inputRows.value))
  editFocusIndex.value = index
  showOtherInputs.value = false
  modal.value = 'edit'
}
async function rerun() { await action('rerun', async () => { const result = await api(`/api/jobs/${job.value.id}/rerun`, { method: 'POST', body: form({ rows: JSON.stringify(editRows.value.map(row => ({ ...row, '报关金额': row['总价'] || row['报关金额'] }))) }) }); modal.value = ''; await openJob(result.id) }) }
async function push() { await action('push', async () => { if (!operator.value.trim()) throw new Error('请填写操作员称呼'); if (!health.value.writeEnabled) throw new Error('部署尚未启用受保护写回'); const result = await api(`/api/sync/${job.value.id}/push`, { method: 'POST', body: form({ operator: operator.value }) }); notify(`写回完成 ${result.written} 组，失败 ${result.failed} 组`); await refreshJob() }) }
function switchTheme() { theme.value = theme.value === 'dark' ? 'light' : 'dark' }
watch(theme, value => { localStorage.setItem('workbench-theme', value); document.documentElement.dataset.theme = value }, { immediate: true })
watch(motion, value => { localStorage.setItem('workbench-motion', value ? 'on' : 'off'); document.documentElement.dataset.motion = value ? 'on' : 'off' }, { immediate: true })
watch(operator, value => localStorage.setItem('workbench-operator', value))

function startBackdrop() {
  const element = canvas.value; if (!element) return
  const ctx = element.getContext('2d'); const pointer = { x: 0.5, y: 0.5 }
  let width = 0, height = 0, tick = 0
  const particles = Array.from({ length: 35 }, (_, i) => ({ x: (i * .618033 + .17) % 1, y: (i * .4142 + .31) % 1, radius: 1 + (i % 3) * .5, speed: .00008 + (i % 5) * .000025 }))
  resizeHandler = () => { width = element.width = innerWidth * devicePixelRatio; height = element.height = innerHeight * devicePixelRatio }
  pointerHandler = event => { pointer.x = event.clientX / innerWidth; pointer.y = event.clientY / innerHeight }
  resizeHandler(); addEventListener('resize', resizeHandler); addEventListener('pointermove', pointerHandler)
  const draw = () => {
    ctx.clearRect(0, 0, width, height)
    const dark = theme.value === 'dark'; const rate = devicePixelRatio
    if (motion.value && !document.hidden && !matchMedia('(prefers-reduced-motion: reduce)').matches) tick++
    const shiftX = (pointer.x - .5) * 18 * rate, shiftY = (pointer.y - .5) * 12 * rate
    for (let i = 0; i < 20; i++) {
      const x = ((i * 149 + 83) % 1100) / 1100 * width + shiftX * .35
      const y = ((i * 241 + 13) % 900) / 900 * height + ((tick * (.35 + i % 3 * .13) * rate) % height)
      ctx.strokeStyle = dark ? 'rgba(125,196,200,.075)' : 'rgba(27,105,115,.065)'
      ctx.lineWidth = rate
      ctx.beginPath(); ctx.moveTo(x, y % height); ctx.lineTo(x - 7 * rate, (y + 42 * rate) % height); ctx.stroke()
    }
    for (const p of particles) {
      if (motion.value && !document.hidden) p.y = (p.y + p.speed) % 1
      const x = p.x * width + shiftX * (1.2 - p.radius * .2), y = p.y * height + shiftY * .6
      ctx.fillStyle = dark ? 'rgba(119,214,205,.25)' : 'rgba(25,112,118,.19)'
      ctx.beginPath(); ctx.arc(x, y, p.radius * rate, 0, Math.PI * 2); ctx.fill()
    }
    frame = requestAnimationFrame(draw)
  }
  draw()
}
onMounted(async () => { startBackdrop(); animateProgress(); await action('init', async () => { health.value = await api('/api/health'); target.value = await api('/api/sync/target').catch(() => ({})); await loadJobs(); await loadExceptions() }); let lastJobs = Date.now(); timer = setInterval(async () => { if (job.value?.state === 'queued' || job.value?.state === 'running') await refreshJob().catch(() => {}); if (Date.now() - lastJobs > 15000) { await loadJobs().catch(() => {}); await loadExceptions().catch(() => {}); lastJobs = Date.now() } }, 500) })
onUnmounted(() => { clearInterval(timer); cancelAnimationFrame(frame); cancelAnimationFrame(progressFrame); removeEventListener('resize', resizeHandler); removeEventListener('pointermove', pointerHandler) })
</script>

<template>
  <canvas ref="canvas" class="ambient" aria-hidden="true"></canvas>
  <div class="shell">
    <aside class="sidebar">
      <div class="brand"><div class="brand-mark">M<span>·</span></div><div><strong>成本匹配工作台</strong><small>FINANCE INTELLIGENCE</small></div></div>
      <div class="sidebar-label">工作区</div>
      <nav aria-label="主导航"><button v-for="item in nav" :key="item.id" class="nav-item" :class="{ active: screen === item.id }" @click="screen = item.id; if(item.id === 'jobs') loadJobs(); if(item.id === 'exceptions') loadExceptions()"><span class="nav-icon">{{ item.icon }}</span>{{ item.label }}<span v-if="item.id === 'jobs' && jobs.length" class="nav-count">{{ jobs.length }}</span></button></nav>
    </aside>
    <main class="main">
      <header class="topbar"><div class="crumb">财务运营 <span>/</span> {{ screen === 'result' ? '任务结果' : screen === 'queue' ? '任务概览' : nav.find(n => n.id === screen)?.label }}</div><div class="top-actions"><button class="icon-button" @click="switchTheme" :title="theme === 'dark' ? '切换白天模式' : '切换夜间模式'">{{ theme === 'dark' ? '☀' : '☾' }}</button></div></header>
      <div class="content">
        <section v-if="screen === 'home'" class="page-enter">
          <div class="hero"><div><div class="eyebrow">FINANCIAL OPERATIONS / 2026</div><h1>让每一笔成本，<br><em>都有据可循。</em></h1><p>报关原件、睿贝凭证与拆单结果，汇聚在一张清晰的财务工作台。</p><div class="hero-actions"><button class="btn primary" @click="screen='sync'">从飞书开始 <span>›</span></button><button class="btn ghost" @click="screen='lookup'">逐条核验</button></div></div><div class="hero-art"><div class="orb orb-one"></div><div class="orb orb-two"></div><div class="hero-grid"></div><div class="art-card"><span>核算流程</span><strong>报关 · 拆单 · 匹配</strong><div class="art-steps"><i></i><i></i><i></i><i></i></div><small>异常优先 · 凭证可追溯</small></div></div></div>
          <div class="stats"><button class="stat-card stat-link" @click="openQueue('exceptions')"><span>待复核记录</span><strong>{{ exceptionCount }}</strong><small>集中处理需要判断的结果 ›</small></button><button class="stat-card stat-link" @click="openQueue('ready')"><span>待写回任务</span><strong>{{ readyCount }}</strong><small>预览后确认写入副本 ›</small></button><button class="stat-card stat-link" @click="openQueue('done')"><span>已完成任务</span><strong>{{ doneCount }}</strong><small>最近任务保留完整结果 ›</small></button><div class="stat-card"><span>ERP 同步</span><strong class="stat-status">{{ health.erpSync?.state || '待检查' }}</strong><small>只读展示数据源状态</small></div></div>
          <div class="section-head"><div><div class="eyebrow">NEXT ACTION</div><h2>继续你的工作</h2></div><button class="text-button" @click="screen='jobs'">查看全部任务 ›</button></div>
          <div class="quick-grid"><button class="quick-card" @click="screen='lookup'"><span class="quick-icon">⌕</span><strong>核验一组记录</strong><small>手动输入、粘贴或附加原始凭证</small><b>进入核验 ›</b></button><button class="quick-card" @click="screen='sync'"><span class="quick-icon">◇</span><strong>处理飞书区间</strong><small>读取未核算商品行并预览写回</small><b>选择范围 ›</b></button><button class="quick-card" @click="screen='template'"><span class="quick-icon">▤</span><strong>匹配模板文件</strong><small>沿用飞书副本列名，批量生成结果附件</small><b>上传文件 ›</b></button></div>
          <div class="section-head"><div><div class="eyebrow">RECENT ACTIVITY</div><h2>最近任务</h2></div></div><div class="panel task-list"><button v-for="item in jobs.slice(0,5)" :key="item.id" class="task-row" @click="openJob(item.id)"><span class="task-badge" :class="item.state">{{ item.kind === 'sync' ? '飞书' : item.kind === 'lookup' ? '核验' : '文件' }}</span><span class="task-main"><strong>{{ jobTitle(item) }}</strong><small>{{ item.message }}</small></span><span class="task-meta">{{ item.summary?.resultRows ?? '—' }} 条结果</span><span class="arrow">›</span></button><div v-if="!jobs.length" class="empty">还没有任务。从上方选择一种方式开始。</div></div>
        </section>

        <section v-if="screen === 'lookup'" class="page-enter narrow"><div class="page-title"><div class="eyebrow">01 / SINGLE REVIEW</div><h1>逐条核验</h1><p>一次最多 50 条。先确认输入，再让系统处理拆单与成本。</p></div><div class="panel form-panel"><div class="panel-head"><div><h2>粘贴或填写报关商品</h2><p>识别结果会先进入下方表格，不会直接计算。</p></div><span class="step-number">01</span></div><textarea v-model="pasteText" rows="4" placeholder="粘贴报关单号、合同号、商品信息或表格文本…"></textarea><div class="button-row"><button class="btn secondary" @click="parsePaste" :disabled="busy==='paste'">解析粘贴</button><button class="btn subtle" @click="enrichLookup" :disabled="busy==='enrich'">从飞书与睿贝补全</button><label class="btn subtle import-label">上传记录<input type="file" accept=".xlsx" @change="importLookup" hidden></label><button class="btn subtle" @click="lookupRows.push(blankRow())" :disabled="lookupRows.length>=50">＋ 添加一行</button><span class="helper">{{ lookupRows.length }} / 50 条</span></div><div v-if="lookupWarnings.length" class="notice amber"><div v-for="(warning,index) in lookupWarnings.slice(0,8)" :key="index">{{ warning }}</div></div><div class="input-grid"><div v-for="(row, index) in lookupRows" :key="index" class="input-record"><div class="record-title"><strong>商品行 {{ String(index+1).padStart(2,'0') }}</strong><button class="text-button danger" @click="lookupRows.splice(index,1)" :disabled="lookupRows.length===1">删除</button></div><p v-if="row['_补全说明']" class="helper">{{ row['_补全说明'] }}</p><div class="fields"><label>报关单号<input v-model="row['报关单号']" placeholder="18 位报关单号"></label><label>合同号<input v-model="row['合同号_1']" placeholder="如 26MT-..."></label><label>采购订单号<input v-model="row['采购订单号']" placeholder="可用睿贝反查出运合同"></label><label>商品序号<input v-model="row['商品序号']" placeholder="可选"></label><label>报关品名<input v-model="row['报关品名']"></label><label>报关金额<input v-model="row['报关金额']" inputmode="decimal"></label><label>报关重量<input v-model="row['报关重量']" inputmode="decimal"></label></div></div></div></div><div class="panel form-panel"><div class="panel-head"><div><h2>补充凭证</h2><p>支持上传报关 PDF 供结果核验和预览。</p></div><span class="step-number">02</span></div><label class="upload-zone">＋ 添加 PDF 附件<input type="file" accept=".pdf" multiple @change="chooseLookupFiles" hidden></label><div class="file-chips"><span v-for="file in lookupFiles" :key="file.name">{{ file.name }}</span></div></div><div class="submit-bar"><span>开始后可随时查看处理进度与异常理由</span><button class="btn primary" @click="startLookup" :disabled="busy==='lookup'">{{ busy==='lookup' ? '正在提交…' : '开始核验 ›' }}</button></div></section>

        <section v-if="screen === 'sync'" class="page-enter narrow"><div class="page-title"><div class="eyebrow">02 / FEISHU RANGE</div><h1>从飞书选取</h1><p>按范围核对待处理商品，查看每笔成本的匹配依据。</p></div><div class="panel form-panel"><div class="panel-head"><div><h2>选择处理范围</h2><p>可按日期、报关单号或合同号选取；请先预览，再开始核验。</p></div><span class="step-number">01</span></div><div class="fields two"><label>起始日期<input type="date" v-model="syncForm.start_date"></label><label>结束日期<input type="date" v-model="syncForm.end_date"></label><label>报关单号<input v-model="syncForm.declarations" placeholder="多个单号可用逗号分隔"></label><label>合同号<input v-model="syncForm.contracts" placeholder="可选"></label><label>单批上限<input type="number" min="1" max="200" v-model.number="syncForm.limit"></label></div><div class="button-row"><button class="btn secondary" @click="previewSync" :disabled="busy==='preview'">{{ busy==='preview' ? '读取中…' : '预览范围' }}</button><button class="btn subtle" @click="syncForm={start_date:'',end_date:'',declarations:'',contracts:'',limit:50};selection=null">重置条件</button></div></div><div v-if="selection" class="panel preview-panel"><div class="panel-head"><div><h2>选择预览</h2><p>请核对本批外销商品与暂不处理的内销记录。</p></div><span class="step-number">02</span></div><div class="preview-stats"><div><strong>{{ selection.selection.total }}</strong><span>可处理行</span></div><div><strong>{{ selection.selection.declarations }}</strong><span>报关单</span></div><div><strong>{{ selection.selection.supplemental }}</strong><span>同合同补齐行</span></div></div><div class="trade-section-head">外销 · 待核验商品</div><div class="preview-list"><div v-for="item in selection.selected.slice(0,8)" :key="item.record_id"><span>{{ item['报关单号'] }}</span><span>{{ item['报关品名'] || '品名待核对' }}</span><b>{{ item.writebackBlocked ? '历史拆分 · 需复核' : '外销 · 待核验' }}</b></div></div><div v-if="selection.skipped.some(item=>item.tradeType==='内销')" class="trade-section-head">内销 · 本批暂不核算</div><div v-if="selection.skipped.some(item=>item.tradeType==='内销')" class="preview-list"><div v-for="item in selection.skipped.filter(item=>item.tradeType==='内销').slice(0,20)" :key="item.record_id"><span>{{ item['合同号_1'] || '合同号待核对' }}</span><span>内销</span><b>无需报关单号</b></div></div><div v-if="selection.skipped.some(item=>item.tradeType!=='内销')" class="notice amber">其余暂未纳入的记录请核对字段。</div><div v-if="selection.skipped.some(item=>item.tradeType!=='内销')" class="preview-list"><div v-for="item in selection.skipped.filter(item=>item.tradeType!=='内销').slice(0,20)" :key="item.record_id"><span>{{ item['报关单号'] || item['合同号_1'] || '—' }}</span><span>{{ item.reason }}</span></div></div><div v-if="selection.selection.merged" class="notice amber">{{ selection.selection.merged }} 条历史拆分行已合并金额用于计算，结果会阻止自动写回，请核对原行与旧子行。</div></div><div class="submit-bar"><span>{{ selection?.selection?.total ? `已选 ${selection.selection.total} 条未核算行；开始后可在任务页查看进度与异常` : '先预览范围，确认可处理行后开始核验' }}</span><button class="btn primary" @click="runSync" :disabled="!selection?.selection?.total||busy==='sync'">{{ busy==='sync' ? '正在提交…' : '开始核验 ›' }}</button></div></section>

        <section v-if="screen === 'template'" class="page-enter narrow"><div class="page-title"><div class="eyebrow">03 / TEMPLATE IMPORT</div><h1>模板文件匹配</h1><p>固定字段，批量处理，结果以 XLSX 附件返回。</p></div><div class="panel form-panel"><div class="panel-head"><div><h2>准备商品行</h2><p>直接上传飞书副本表导出的 XLSX，也可下载同列名模板填写；这条流程不会自动写回飞书。</p></div><span class="step-number">01</span></div><a class="btn secondary" href="/api/template" download>⌄ 下载飞书副本同列模板</a><label class="upload-zone large">{{ templateFile ? templateFile.name : '点击选择 .xlsx 文件' }}<input type="file" accept=".xlsx" @change="chooseTemplate" hidden></label><div class="button-row"><button class="btn secondary" @click="checkTemplate" :disabled="!templateFile||busy==='check'">检查文件</button><span v-if="templateCheck" class="helper">{{ templateCheck.errors.length ? `${templateCheck.errors.length} 处需修正` : `${templateCheck.rows} 行检查通过` }}</span></div><div v-if="templateCheck?.errors.length" class="validation"><div v-for="error in templateCheck.errors.slice(0,12)" :key="error.row">第 {{ error.row }} 行 · {{ error.message }}</div></div></div><div class="submit-bar"><span>检查通过后开始成本匹配</span><button class="btn primary" @click="runTemplate" :disabled="!templateCheck||templateCheck.errors.length||busy==='template'">开始匹配 ›</button></div></section>

        <section v-if="screen === 'queue'" class="page-enter"><div class="page-title"><div class="eyebrow">WORK QUEUE</div><h1>{{ queueKind === 'exceptions' ? '待复核记录' : queueKind === 'ready' ? '待写回任务' : '已完成任务' }}</h1><p>{{ queueKind === 'exceptions' ? '集中查看需要人工核对的商品行。' : queueKind === 'ready' ? '打开任务核对写回预览。' : '查看已完成的计算与凭证。' }}</p></div><div class="panel task-list"><template v-if="queueKind === 'exceptions'"><button v-for="item in exceptions.slice(0, exceptionVisible)" :key="item.jobId+'-'+item.index" class="task-row" @click="openJob(item.jobId,item.index)"><span class="task-badge failed">异常</span><span class="task-main"><strong>{{ item.declaration || item.contract || '待核对商品' }} · {{ item.product || '商品行' }}</strong><small>{{ shortDate(item.createdAt) }} · {{ item.type }} · {{ item.reason || '待补充原因' }}</small></span><span class="arrow">›</span></button><div v-if="!exceptions.length" class="empty">暂无待复核记录</div><div v-if="exceptions.length>exceptionVisible" class="list-more"><button class="btn subtle" @click="exceptionVisible+=50">查看更多（{{ exceptions.length-exceptionVisible }} 条）</button></div></template><template v-else><button v-for="item in queueJobs" :key="item.id" class="task-row" @click="openJob(item.id)"><span class="task-badge" :class="item.state">{{ item.kind === 'sync' ? '飞书' : item.kind === 'lookup' ? '核验' : '文件' }}</span><span class="task-main"><strong>{{ jobTitle(item) }}</strong><small>{{ item.message }}</small></span><span class="task-meta">{{ queueKind === 'ready' ? `${item.writebackReady} 组可写回` : `${item.summary?.resultRows ?? 0} 条结果` }}</span><span class="arrow">›</span></button><div v-if="!queueJobs.length" class="empty">暂无对应任务</div></template></div></section>

        <section v-if="screen === 'exceptions'" class="page-enter"><div class="page-title"><div class="eyebrow">EXCEPTION RECORDS</div><h1>全部异常记录</h1><p>系统异常与人工上报集中留存，点击可回到原任务核对。</p></div><div class="panel task-list"><button v-for="item in exceptions.slice(0, exceptionVisible)" :key="item.jobId+'-'+item.index" class="task-row" @click="openJob(item.jobId,item.index)"><span class="task-badge failed">异常</span><span class="task-main"><strong>{{ item.declaration || item.contract || '待核对商品' }} · {{ item.product || '商品行' }}</strong><small>{{ shortDate(item.createdAt) }} · {{ item.type }} · {{ item.reason || '待补充原因' }}</small></span><span class="arrow">›</span></button><div v-if="!exceptions.length" class="empty">暂无异常记录</div><div v-if="exceptions.length>exceptionVisible" class="list-more"><button class="btn subtle" @click="exceptionVisible+=50">查看更多（{{ exceptions.length-exceptionVisible }} 条）</button></div></div></section>

        <section v-if="screen === 'jobs'" class="page-enter"><div class="page-title"><div class="eyebrow">TASK HISTORY</div><h1>任务记录</h1><p>保留每次计算结果、异常与操作痕迹。</p></div><div class="panel task-list"><button v-for="item in jobs" :key="item.id" class="task-row" @click="openJob(item.id)"><span class="task-badge" :class="item.state">{{ item.kind === 'sync' ? '飞书' : item.kind === 'lookup' ? '核验' : '文件' }}</span><span class="task-main"><strong>{{ jobTitle(item) }}</strong><small>{{ item.message }}</small></span><span class="task-meta">{{ item.summary?.resultRows ?? '—' }} 条 · {{ item.summary?.exceptionRows ?? '—' }} 异常</span><span class="arrow">›</span></button><div v-if="!jobs.length" class="empty">暂无任务记录</div></div></section>

        <section v-if="screen === 'result' && (job || submitting)" class="page-enter"><div v-if="submitting && !job" class="result-header"><div><div class="eyebrow">正在准备任务</div><h1>读取并核验数据</h1><p>{{ submitting === 'sync' ? '正在读取所选飞书商品行…' : '正在接收核验记录…' }}</p></div><span class="status-pill running">处理中</span></div><div v-if="job" class="result-header"><div><button class="text-button back" @click="screen='jobs'">‹ 返回任务记录</button><div class="eyebrow">{{ jobTitle(job) }}</div><h1>{{ job.kind === 'sync' ? '飞书区间计算' : job.kind === 'lookup' ? '逐条核验结果' : '文件匹配结果' }}</h1><p>{{ job.message }}</p></div><span class="status-pill" :class="job.state">{{ job.state === 'complete' ? '处理完成' : job.state === 'failed' ? '处理失败' : '处理中' }}</span></div><div class="progress-rail"><div v-for="(step,index) in ['读取','解析','拆单','匹配','复核']" :key="step" :class="{ done: progressView >= [8,20,45,70,92][index] }"><i></i><span>{{ step }}</span></div></div><div v-if="submitting || (job && (job.state==='running'||job.state==='queued')) || (job?.state==='complete' && !progressSettled)" class="panel progress-panel"><div class="progress-track"><span :style="{width: progressView+'%'}"></span></div><strong>{{ progressView.toFixed(2) }}%</strong><p>{{ submitting && !job ? '正在读取并建立任务…' : job?.message }}</p><button v-if="job && job.state!=='complete'" class="btn subtle" @click="api(`/api/jobs/${job.id}/cancel`,{method:'POST'}).then(refreshJob)">取消任务</button></div><div v-else-if="job?.state==='failed'" class="notice red">{{ job.message }}</div><template v-else-if="job?.state==='complete'"><div class="stats result-stats"><div class="stat-card"><span>拆单结果</span><strong>{{ job.summary?.resultRows }}</strong><small>记录行</small></div><div class="stat-card"><span>正常</span><strong>{{ job.summary?.successRows }}</strong><small>写回资格以预览为准</small></div><div class="stat-card priority"><span>需要核对</span><strong>{{ job.summary?.exceptionRows }}</strong><small>异常先处理</small></div><div class="stat-card"><span>采购金额合计</span><strong class="amount-large">¥ {{ money(job.summary?.purchaseTotal) }}</strong><small>入库单实发口径</small></div></div><div v-if="job.skipped?.some(item=>item.tradeType==='内销')" class="notice amber skipped-result"><strong>内销记录 · 本批暂不核算</strong><div v-for="item in job.skipped.filter(item=>item.tradeType==='内销').slice(0,20)" :key="item.record_id">{{ item['合同号_1'] || '合同号待核对' }} · 无需报关单号</div></div><div v-if="job.skipped?.some(item=>item.tradeType!=='内销')" class="notice amber skipped-result"><strong>其它待核对记录</strong><div v-for="item in job.skipped.filter(item=>item.tradeType!=='内销').slice(0,20)" :key="item.record_id">{{ item['报关单号'] || item['合同号_1'] || '—' }} · {{ item.reason }}</div></div><div class="result-toolbar"><div class="segmented"><button :class="{selected:onlyExceptions}" @click="onlyExceptions=true">仅看异常</button><button :class="{selected:!onlyExceptions}" @click="onlyExceptions=false">查看全部</button></div><input v-model="search" class="search" placeholder="搜索报关单 / 采购单 / 供应商"><button class="btn secondary" @click="focusNext(false)">下一条异常 ⌄</button><button v-if="job.kind!=='sync'" class="btn subtle" @click="compare">对照飞书</button></div><div class="result-layout"><div class="panel result-list"><button v-for="{row,index} in filteredRows" :key="index" :id="`result-${index}`" class="result-row" :class="{ active:focused===index, alert:row['异常类型']!=='正常', pulse: pulse===index && motion && !row['_人工上报'], flagged: pulse===index && motion && row['_人工上报'] }" @click="focused=index"><span class="result-indicator"></span><span class="result-identity"><strong>{{ row['报关单号'] || row['合同号_1'] || '待核对' }} <em class="trade-tag">{{ row['报关单号'] ? '外销' : (row['_业务类型'] || '出运补齐') }}</em></strong><small>{{ row['报关品名'] }} · {{ row['供应商简称'] || '供应商待确认' }}</small></span><span class="result-money">¥ {{ money(row['采购金额']) }}</span><span class="row-status" :class="contextIndexes.has(index)?'context':row['异常类型']==='正常'?'ok':'warn'">{{ contextIndexes.has(index) ? '计算补齐 · 不写回' : row['异常类型']==='正常' ? '正常' : '需核对' }}</span><span class="arrow">›</span></button><div v-if="!filteredRows.length" class="empty">当前筛选下没有记录</div></div><div class="panel detail-panel" v-if="activeRow"><div class="detail-head"><span class="eyebrow">核验依据 / {{ String(focused+1).padStart(2,'0') }}</span><span class="row-status" :class="activeRow['异常类型']==='正常'?'ok':'warn'">{{ activeRow['异常类型'] }}</span></div><h2>{{ activeRow['报关品名'] || '商品详情' }}</h2><p class="detail-sub">{{ activeRow['报关单号'] ? '外销 · ' + activeRow['报关单号'] : (activeRow['_业务类型'] || '出运补齐') }} · {{ activeRow['合同号_1'] }}</p><div v-if="contextIndexes.has(focused)" class="notice amber">这条同合同记录仅参与本次成本计算，不属于所选未核算行，不会写回飞书。</div><div v-if="activeRow['异常类型']!=='正常'" class="notice red"><strong>需要核对</strong><br>{{ activeRow['异常明细'] || '请检查原始凭证与成本来源。' }}</div><div v-else-if="activeRow['异常明细']" class="notice amber">{{ activeRow['异常明细'] }}</div><div class="evidence-chain"><div><i>1</i><span>报关原件</span><strong>{{ activeRow['报关金额'] || '—' }}</strong></div><div><i>2</i><span>采购订单 / 供应商</span><strong>{{ activeRow['采购订单号'] || '—' }}</strong></div><div><i>3</i><span>成本分摊</span><strong>{{ activeRow['分摊依据'] || '待核对' }}</strong></div></div><div class="detail-values"><div><span>采购金额</span><strong>¥ {{ money(activeRow['采购金额']) }}</strong></div><div><span>报关金额</span><strong>{{ money(activeRow['报关金额']) }}</strong></div><div><span>报关重量</span><strong>{{ activeRow['报关重量'] || '—' }}</strong></div><div v-if="job.kind!=='sync'"><span>飞书对照</span><strong>{{ compareStatus(activeRow) }}</strong></div><div><span>附件对照</span><strong :title="evidenceComparison?.rows?.find(item => item['报关单号'] === activeRow['报关单号'])?.['说明'] || ''">{{ evidenceStatus(activeRow) }}</strong></div></div><div class="source-links"><button v-if="health.evidenceEnabled" class="source-button" @click="showErp">查看睿贝缓存凭证 ›</button><a v-if="activePdf" :href="activePdf.url" target="_blank">查看报关 PDF ›</a><a v-if="activeSource && target.appToken" :href="`https://my.feishu.cn/base/${target.appToken}?table=${target.tableId}`" target="_blank">打开飞书原表 · {{ activeRow['合同号_1'] || '查看原行' }} ›</a><a :href="activeRow['_睿贝采购单链接'] || 'https://erp.mtholdinggroup.com/purchase_goOutList?menuCode=80400'" target="_blank" rel="noopener noreferrer">在睿贝查采购单 · {{ activeRow['采购订单号'] }} ›</a><a :href="activeRow['_睿贝出运单链接'] || 'https://erp.mtholdinggroup.com/saleOrder?menuCode=80300'" target="_blank" rel="noopener noreferrer">在睿贝查出运单 · {{ activeRow['合同号_1'] }} ›</a></div><div class="button-row detail-actions"><button class="btn subtle" @click="modal='sample'" :disabled="activeRow['_已抽查']">{{ activeRow['_已抽查'] ? '已抽查 ✓' : '标记已抽查' }}</button><button class="btn secondary" @click="openEdit">修改输入并重算</button><button class="btn subtle" @click="overrideAmount=activeRow['采购金额'];overrideReason='';modal='override'">人工调整金额</button><button v-if="job.kind==='sync' && activeRow['异常类型']==='正常'" class="btn flag-button" @click="flagReason='';modal='flag'">上报异常</button></div></div></div><div class="bottom-actions"><div class="downloads"><a :href="job.downloads?.results" download>⌄ 完整结果 XLSX</a><a :href="job.downloads?.exceptions" download>⌄ 异常清单 XLSX</a></div><button v-if="job.kind==='sync'" class="btn primary" @click="modal='writeback'">查看写回预览 ›</button></div><div v-if="compareData" class="panel comparison"><h2>飞书对照</h2><p v-if="compareData.state==='running'">正在读取飞书对照记录…</p><p v-else-if="compareData.state==='failed'">{{ compareData.message }}</p><p v-else>{{ compareData.summary?.consistent }} 组结果一致 · {{ compareData.summary?.exceptions }} 组结果不一致或暂无法比较</p></div></template></section>
      </div>
    </main>
  </div>
  <nav class="mobile-nav" aria-label="移动导航"><button v-for="item in nav" :key="item.id" :class="{active:screen===item.id}" @click="screen=item.id"><span>{{ item.icon }}</span><small>{{ item.label }}</small></button></nav>
  <div v-if="toast" class="toast" role="status">{{ toast }}</div>
  <div v-if="modal" class="modal-backdrop" @click.self="modal=''" role="dialog" aria-modal="true"><div class="modal-card"><button class="modal-close" @click="modal=''" aria-label="关闭">×</button><template v-if="modal==='sample'"><div class="eyebrow">SAMPLE REVIEW</div><h2>记录抽样复核</h2><p>标记仅表示人工已查看此条，不改变系统金额和异常判定。</p><label>操作员称呼<input v-model="operator" placeholder="用于操作留痕，自填身份"></label><div class="button-row"><button class="btn primary" @click="markSample">确认已查看</button><button class="btn subtle" @click="modal=''">取消</button></div></template><template v-else-if="modal==='erp'"><div class="eyebrow">ERP SOURCE</div><h2>睿贝凭证详情</h2><p>来自本地同步缓存，请结合睿贝原系统核对。</p><div class="erp-source"><h3>采购订单</h3><div v-for="item in erpDetail?.purchase?.purchases || []" :key="item.purchase_id"><strong>{{ item.purchase_code }}</strong><span>{{ item.supplierName || '供应商待确认' }}</span><span>金额 {{ money(item.amount) }}</span></div><p v-if="!erpDetail?.purchase?.purchases?.length">缓存未找到对应采购单。</p><h3>出运单</h3><div v-for="item in erpDetail?.shipment?.shipments || []" :key="item.shipmentId"><strong>{{ item.invoiceCode }}</strong><span>出运日期 {{ item.shipDate || '—' }}</span><span>金额 {{ money(item.amount) }}</span></div><p v-if="!erpDetail?.shipment?.shipments?.length">缓存未找到对应出运单。</p></div><div class="button-row"><button class="btn secondary" @click="modal=''">关闭</button></div></template><template v-else-if="modal==='flag'"><div class="eyebrow">EXCEPTION REPORT</div><h2>上报这条记录的异常</h2><p>{{ activeRow?.['报关单号'] || activeRow?.['合同号_1'] }} · {{ activeRow?.['报关品名'] }}。上报后所在写回组将暂停，原因会进入全部异常记录。</p><label>异常原因<textarea v-model="flagReason" rows="4" placeholder="请说明核对发现的问题"></textarea></label><label>操作员称呼<input v-model="operator" placeholder="用于操作留痕"></label><div class="button-row"><button class="btn primary" @click="saveFlag" :disabled="busy==='flag'||!flagReason.trim()||!operator.trim()">确认上报</button><button class="btn subtle" @click="modal=''">取消</button></div></template><template v-else-if="modal==='override'"><div class="eyebrow">MANUAL ADJUSTMENT</div><h2>人工调整采购金额</h2><p>原计算值 {{ money(activeRow?.['采购金额']) }}。调整记录会保留，整组写回校验仍然生效。</p><label>调整后金额<input v-model="overrideAmount" inputmode="decimal"></label><label>调整依据（可选）<textarea v-model="overrideReason" rows="3"></textarea></label><label>操作员称呼<input v-model="operator" placeholder="用于操作留痕，自填身份"></label><div class="button-row"><button class="btn primary" @click="saveOverride">保存调整</button><button class="btn subtle" @click="modal=''">取消</button></div></template><template v-else-if="modal==='edit'"><div class="eyebrow">当前商品输入</div><h2>修改当前记录并重算</h2><p>先核对这条结果对应的输入行。保存后会带上同批其它商品行，重新完成整笔分摊。</p><div v-if="editFocusIndex>=0" class="edit-scroll"><div v-for="row in editRows.slice(editFocusIndex,editFocusIndex+1)" :key="editFocusIndex" class="fields two edit-row edit-primary"><label>报关单号<input v-model="row['报关单号']"></label><label>合同号<input v-model="row['合同号_1']"></label><label>品名<input v-model="row['报关品名']"></label><label>总价<input v-model="row['总价']"></label><label>重量<input v-model="row['报关重量']"></label></div><button v-if="editRows.length>1" class="text-button" @click="showOtherInputs=!showOtherInputs">{{ showOtherInputs ? '收起同批其它输入' : '查看同批其它输入（'+(editRows.length-1)+' 条）' }}</button><div v-if="showOtherInputs"><template v-for="(row,index) in editRows" :key="index"><div v-if="index!==editFocusIndex" class="fields two edit-row"><label>报关单号<input v-model="row['报关单号']"></label><label>合同号<input v-model="row['合同号_1']"></label><label>品名<input v-model="row['报关品名']"></label><label>总价<input v-model="row['总价']"></label><label>重量<input v-model="row['报关重量']"></label></div></template></div></div><div class="button-row"><button class="btn primary" @click="rerun">保存并重算</button><button class="btn subtle" @click="modal=''">取消</button></div></template><template v-else-if="modal==='writeback'"><div class="eyebrow">写回前核对</div><h2>飞书写回预览</h2><p>系统按报关单号、合同号、供应商和品名对应原商品行；无法唯一对应的组会暂停写回。</p><p>系统优先更新对应的原商品行，缺少对应行时新增；整组报关金额不守恒时暂停写回。</p><div class="preview-stats"><div><strong>{{ plan?.summary?.ready ?? 0 }}</strong><span>可写回组</span></div><div><strong>{{ plan?.summary?.blocked ?? 0 }}</strong><span>暂停组</span></div><div><strong>{{ plan?.summary?.unmatched ?? 0 }}</strong><span>待复核对应</span></div></div><div v-if="plan?.unmatched?.length" class="notice amber">有 {{ plan.unmatched.length }} 条结果无法唯一对应原商品行，已暂停相关写回；请核对合同号、供应商与品名。</div><div class="write-groups"><div v-for="group in plan?.groups || []" :key="group.sourceId" class="write-group"><div><strong>{{ group.declaration || group.contract }}</strong><span class="row-status" :class="group.status==='ready'?'ok':'warn'">{{ group.status==='ready' ? (group.manualAdjusted ? '人工调整，可写回' : '可写回') : '暂停' }}</span></div><small>原行 {{ money(group.originalAmount) }} 至拆分后 {{ money(group.resultAmount) }} · {{ group.children.length }} 条</small><p v-if="group.reasons.length">{{ group.reasons.join('；') }}</p><div v-for="(child,index) in group.children" :key="child.splitKey" class="write-child">{{ group.memberCount>1 ? '匹配原商品行' : index===0 ? '更新原行' : '新增子行' }} · {{ child['采购订单号'] }} · 报关 {{ money(child['报关金额']) }} · 采购 {{ money(child['采购金额']) }}</div></div></div><label>操作员称呼<input v-model="operator" placeholder="用于操作留痕，自填身份"></label><div v-if="!health.writeEnabled" class="notice amber">受保护写回尚未由部署方启用。可先核对预览并下载结果。</div><div class="button-row"><button class="btn primary" @click="push" :disabled="!health.writeEnabled||!plan?.summary?.ready||busy==='push'">确认写回</button><button class="btn subtle" @click="modal=''">返回复核</button></div></template></div></div>
</template>
