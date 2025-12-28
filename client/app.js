const API_BASE = () => window.API_BASE || 'http://127.0.0.1:8000'
const HWPX_VIEWER_URL = window.HWPX_VIEWER_URL || ''

const $ = (sel) => document.querySelector(sel)
const $$ = (sel) => Array.from(document.querySelectorAll(sel))

let __lastRedactedBlob = null
let __lastRedactedName = 'redacted.bin'

// html escape
const esc = (s) =>
  (s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

// 배지
const badge = (sel, n) => {
  const el = $(sel)
  if (el) el.textContent = String(n ?? 0)
}

// 아코디언
const setOpen = (name, open) => {
  const cont =
    name === 'pdf' ? $('#pdf-preview-block') : $(`#${name}-result-block`)
  const body = $(`#${name}-body`)
  const chev = document.querySelector(`[data-chevron="${name}"]`)
  cont && cont.classList.remove('hidden')
  body && body.classList.toggle('hidden', !open)
  chev && chev.classList.toggle('rotate-180', !open)
}
document.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-toggle]')
  if (!btn) return
  const name = btn.getAttribute('data-toggle')
  const body = document.getElementById(`${name}-body`)
  setOpen(name, body ? body.classList.contains('hidden') : true)
})

// 규칙 로드
async function loadRules() {
  try {
    const r = await fetch(`${API_BASE()}/text/rules`)
    if (!r.ok) throw 0
    const rules = await r.json()
    const box = $('#rules-container')
    if (!box) return
    box.innerHTML = ''
    for (const rule of rules) {
      const el = document.createElement('label')
      el.className = 'flex items-center gap-2'
      el.innerHTML = `<input type="checkbox" name="rule" value="${rule}" checked><span>${esc(
        rule
      )}</span>`
      box.appendChild(el)
    }
  } catch {
    console.warn('규칙 불러오기 실패')
  }
}
function selectedRuleNames() {
  return $$('input[name="rule"]:checked').map((el) => el.value)
}

// 드롭존
function setupDropZone() {
  const dz = $('#dropzone'),
    input = $('#file'),
    nameEl = $('#file-name'),
    statusEl = $('#status')
  if (!dz || !input) return

  let depth = 0
  const setActive = (on) => {
    dz.classList.toggle('ring-2', on)
    dz.classList.toggle('ring-blue-400', on)
    dz.classList.toggle('bg-blue-50', on)
  }
  const showName = (f) => {
    if (nameEl) nameEl.textContent = f ? `선택됨: ${f.name}` : ''
  }

  ;['dragover', 'drop'].forEach((ev) =>
    window.addEventListener(ev, (e) => e.preventDefault())
  )
  dz.addEventListener('dragenter', (e) => {
    e.preventDefault()
    depth++
    setActive(true)
    e.dataTransfer && (e.dataTransfer.dropEffect = 'copy')
  })
  dz.addEventListener('dragover', (e) => {
    e.preventDefault()
    e.dataTransfer && (e.dataTransfer.dropEffect = 'copy')
  })
  ;['dragleave', 'dragend'].forEach((ev) =>
    dz.addEventListener(ev, (e) => {
      e.preventDefault()
      depth = Math.max(0, depth - 1)
      if (!depth) setActive(false)
    })
  )
  dz.addEventListener('drop', (e) => {
    e.preventDefault()
    depth = 0
    setActive(false)
    const dt = e.dataTransfer
    let file = (dt.files && dt.files[0]) || null
    if (!file && dt.items) {
      for (const it of dt.items) {
        if (it.kind === 'file') {
          const f = it.getAsFile()
          if (f) {
            file = f
            break
          }
        }
      }
    }
    if (!file) {
      statusEl && (statusEl.textContent = '드래그한 항목이 파일이 아닙니다.')
      return
    }
    const repl = new DataTransfer()
    repl.items.add(file)
    input.files = repl.files
    input.dispatchEvent(new Event('change', { bubbles: true }))
    showName(file)
    statusEl &&
      (statusEl.textContent = '파일 선택 완료 — 스캔 실행을 눌러주세요.')
  })
  input.addEventListener('change', (e) => showName(e.target.files?.[0] || null))
}

// PDF 프리뷰(1페이지)
async function renderRedactedPdfPreview(blob) {
  const cv = $('#pdf-preview')
  if (!cv) return
  const g = cv.getContext('2d')
  if (!blob) return g.clearRect(0, 0, cv.width, cv.height)
  const pdf = await pdfjsLib.getDocument({ data: await blob.arrayBuffer() })
    .promise
  const page = await pdf.getPage(1)
  const vp = page.getViewport({ scale: 1.2 })
  cv.width = vp.width
  cv.height = vp.height
  await page.render({ canvasContext: g, viewport: vp }).promise
}

// 텍스트 하이라이트
const take = (s, n) => (s.length <= n ? s : s.slice(0, n) + '…')
function highlightFrag(ctx, val, pad = 60) {
  const i = (ctx || '').indexOf(val || '')
  if (i < 0) return esc(take(ctx || '', 140))
  const start = Math.max(0, i - pad)
  const end = Math.min((ctx || '').length, i + (val || '').length + pad)
  const pre = esc((ctx || '').slice(start, i))
  const mid = esc(val || '')
  const post = esc((ctx || '').slice(i + (val || '').length, end))
  return pre + `<mark class="bg-yellow-200 rounded px-1">${mid}</mark>` + post
}

let __segFilter = 'all' // all | ok | fail
function applySegmentFilter(root) {
  root.querySelectorAll('[data-valid]').forEach((el) => {
    const ok = el.getAttribute('data-valid') === '1'
    let show = true
    if (__segFilter === 'ok') show = ok
    else if (__segFilter === 'fail') show = !ok
    el.style.display = show ? '' : 'none'
  })
}
function wireSegmentButtons(root) {
  const setActive = (which) => {
    __segFilter = which
    ;['all', 'ok', 'fail'].forEach((k) => {
      const btn = $(`#seg-${k}`)
      if (!btn) return
      btn.classList.remove(
        'bg-gray-900',
        'text-white',
        'bg-emerald-600',
        'bg-rose-600'
      )
      if (k === 'all' && which === 'all')
        btn.classList.add('bg-gray-900', 'text-white')
      if (k === 'ok' && which === 'ok')
        btn.classList.add('bg-emerald-600', 'text-white')
      if (k === 'fail' && which === 'fail')
        btn.classList.add('bg-rose-600', 'text-white')
    })
    applySegmentFilter(root)
  }
  $('#seg-all')?.addEventListener('click', () => setActive('all'))
  $('#seg-ok')?.addEventListener('click', () => setActive('ok'))
  $('#seg-fail')?.addEventListener('click', () => setActive('fail'))
  setActive(__segFilter)
}
function renderRegexResults(res) {
  const items = Array.isArray(res?.items) ? res.items : []
  badge('#match-badge', items.length)

  const summary = $('#summary')
  if (summary) {
    const counts = res?.counts || {}
    summary.textContent = `검출: ${
      Object.keys(counts).length
        ? Object.entries(counts)
            .map(([k, v]) => `${k}=${v}`)
            .join(', ')
        : '없음'
    }`
  }

  const wrap = $('#match-groups')
  if (!wrap) return
  wrap.innerHTML = ''

  const groups = {}
  for (const it of items) (groups[it.rule || 'UNKNOWN'] ??= []).push(it)

  for (const [rule, arr] of Object.entries(groups).sort(
    (a, b) => b[1].length - a[1].length
  )) {
    const ok = arr.filter((x) => x.valid).length
    const fail = arr.length - ok

    const container = document.createElement('div')
    container.className = 'rounded-2xl border border-gray-200'
    container.innerHTML = `
      <button class="w-full flex items-center justify-between px-4 py-2.5 bg-gray-50 hover:bg-gray-100 rounded-t-2xl">
        <div class="flex items-center gap-2">
          <span class="text-sm font-semibold">${esc(rule)}</span>
          <span class="text-xs text-gray-500">총 ${arr.length}건</span>
          <span class="text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700">OK ${ok}</span>
          ${
            fail
              ? `<span class="text-[10px] px-1.5 py-0.5 rounded bg-rose-100 text-rose-700">FAIL ${fail}</span>`
              : ''
          }
        </div>
        <svg class="h-4 w-4 text-gray-500 transition-transform" viewBox="0 0 20 20" fill="currentColor">
          <path fill-rule="evenodd" d="M5.23 7.21a.75.75 0 011.06.02L10 10.94l3.71-3.7a.75.75 0 111.06 1.06l-4.24 4.24a.75.75 0 01-1.06 0L5.21 8.29a.75.75 0 01.02-1.08z" clip-rule="evenodd"/>
        </svg>
      </button>
      <div class="p-3 grid gap-2 rounded-b-2xl"></div>
    `
    const body = container.querySelector('.p-3')

    for (const r of arr) {
      const isOk = !!r.valid
      const ctx = r.context || ''
      const val = r.value || ''
      const card = document.createElement('div')
      card.dataset.valid = isOk ? '1' : '0'
      card.className =
        'border rounded-xl p-3 bg-white hover:shadow-sm transition ' +
        (isOk ? 'border-emerald-200' : 'border-rose-200')

      // 값은 평문 모노스페이스, 불필요한 배경/박스 제거
      card.innerHTML = `
        <div class="flex items-start justify-between gap-3">
          <div class="min-w-0">
            <div class="text-sm font-mono break-all">${esc(val)}</div>
            <div class="text-[12px] text-gray-600 mt-1 leading-relaxed break-words">
              ${highlightFrag(ctx, val)}
            </div>
          </div>
          <div class="shrink-0">
            <span class="inline-block text-[11px] px-1.5 py-0.5 rounded border ${
              isOk
                ? 'border-emerald-300 text-emerald-700'
                : 'border-rose-300 text-rose-700'
            }">${isOk ? 'OK' : 'FAIL'}</span>
          </div>
        </div>
      `
      body.appendChild(card)
    }

    let open = arr.length <= 10
    body.style.display = open ? '' : 'none'
    container.querySelector('button')?.addEventListener('click', () => {
      open = !open
      body.style.display = open ? '' : 'none'
      container.querySelector('svg')?.classList.toggle('rotate-180', !open)
    })

    wrap.appendChild(container)
  }

  wireSegmentButtons(wrap)

  $('#filter-search')?.addEventListener('input', (e) => {
    const q = (e.target.value || '').toLowerCase()
    wrap.querySelectorAll('[data-valid]').forEach((el) => {
      const txt = el.textContent.toLowerCase()
      const match = !q || txt.includes(q)
      el.style.display = match ? '' : 'none'
    })
    applySegmentFilter(wrap)
  })

  applySegmentFilter(wrap)
}

function normalizeNerItems(raw, srcText = '') {
  if (!raw) return { items: [] }
  let arr = []
  if (Array.isArray(raw.items)) arr = raw.items
  else if (Array.isArray(raw.entities)) arr = raw.entities
  else if (Array.isArray(raw.result?.entities)) arr = raw.result.entities
  else if (Array.isArray(raw)) arr = raw

  // /text/detect 폴백 형태: final_spans에서 source==='ner'
  if (!arr.length && Array.isArray(raw.final_spans)) {
    const spans = raw.final_spans.filter(
      (s) => (s.source || '').toLowerCase() === 'ner'
    )
    arr = spans.map((s) => {
      const start = Number(s.start ?? 0)
      const end = Number(s.end ?? 0)
      const text =
        start >= 0 && end > start && srcText
          ? srcText.slice(start, end)
          : s.text || ''
      return {
        label: s.label || '',
        text,
        score: s.score ?? s.prob,
        start,
        end,
      }
    })
  }

  // 필드 스펙 표준화
  const map = (e) => ({
    label: e.label ?? e.entity ?? e.tag ?? '',
    text: e.text ?? e.word ?? e.value ?? '',
    score:
      typeof e.score === 'number'
        ? e.score
        : typeof e.confidence === 'number'
        ? e.confidence
        : undefined,
    start: typeof e.start === 'number' ? e.start : 0,
    end:
      typeof e.end === 'number'
        ? e.end
        : typeof e.length === 'number'
        ? (e.start || 0) + e.length
        : 0,
  })
  return { items: arr.map(map) }
}
async function requestNerSmart(text) {
  // 1) /text/ner
  try {
    const r = await fetch(`${API_BASE()}/text/ner`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text }),
    })
    if (r.ok) {
      const j = await r.json()
      const n = normalizeNerItems(j)
      if (n.items.length) return n
    }
  } catch {}
  // 2) /text/detect (run_ner)
  const r2 = await fetch(`${API_BASE()}/text/detect`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      text,
      options: { run_regex: false, run_ner: true },
    }),
  })
  if (!r2.ok) return { items: [] }
  const j2 = await r2.json()
  return normalizeNerItems(j2, j2.text || text || '')
}

// NER 테이블
function renderNerTable(ner) {
  const rows = $('#ner-rows')
  if (!rows) return

  const items = state.detections.filter((d) => d.source === 'ner')

  const scores = items
    .map((d) => (typeof d.score === 'number' ? d.score : null))
    .filter((x) => x != null)
  const avg = scores.length
    ? scores.reduce((a, b) => a + b, 0) / scores.length
    : null
  const sum = $('#ner-summary')
  if (sum) sum.textContent = `총 ${items.length} · 평균 score ${Score(avg)}`

  rows.innerHTML = ''
  for (const d of items) {
    const tr = document.createElement('tr')
    tr.className = 'ner-row border-b hover:bg-gray-50 cursor-pointer'
    tr.dataset.id = d.id
    tr.innerHTML = `
      <td class="py-2 px-2 font-semibold">${escHtml(d.label || '')}</td>
      <td class="py-2 px-2">${escHtml(d.text)}</td>
      <td class="py-2 px-2 font-mono">${escHtml(Score(d.score))}</td>
      <td class="py-2 px-2 font-mono text-[12px] opacity-70">${escHtml(
        `${d.start ?? '-'}-${d.end ?? '-'}`
      )}</td>
    `
    tr.addEventListener('click', () => {
      state.selectedId = d.id
      setActiveResultItem(d.id)
      clearViewerSelection()
      applyViewerSelection(d.id)
    })
    rows.appendChild(tr)
  }
}

function setMatchTab(tab) {
  const t = tab === 'ner' ? 'ner' : 'regex'
  state.ui = state.ui || {}
  state.ui.matchTab = t

  const paneRegex = $('#match-pane-regex')
  const paneNer = $('#match-pane-ner')
  paneRegex && paneRegex.classList.toggle('hidden', t !== 'regex')
  paneNer && paneNer.classList.toggle('hidden', t !== 'ner')

  const label = $('#match-tab-label')
  if (label) label.textContent = t === 'regex' ? '정규식' : 'NER'

  const badge = $('#match-badge')
  if (badge) {
    const n =
      t === 'regex'
        ? state.detections.filter((d) => d.source === 'regex').length
        : state.detections.filter((d) => d.source === 'ner').length
    badge.textContent = String(n)
  }
}

function wireMatchTabs() {
  const prev = $('#btn-match-prev')
  const next = $('#btn-match-next')
  if (prev)
    prev.addEventListener('click', () =>
      setMatchTab((state.ui?.matchTab || 'regex') === 'regex' ? 'ner' : 'regex')
    )
  if (next)
    next.addEventListener('click', () =>
      setMatchTab((state.ui?.matchTab || 'regex') === 'regex' ? 'ner' : 'regex')
    )
}

function wireViewerClick() {
  const viewer = $('#doc-viewer')
  if (!viewer) return
  viewer.addEventListener('click', (e) => {
    const sp = e.target.closest('.pii-box')
    if (!sp) return
    const id = sp.getAttribute('data-id')
    if (!id) return
    state.selectedId = id
    clearViewerSelection()
    applyViewerSelection(id)
    setActiveResultItem(id)
    const d = state.detectionById?.get(id)
    if (d?.source === 'ner') setMatchTab('ner')
    else setMatchTab('regex')
  })
}

/** ---------- Stats & Report ---------- */
function pad2(n) {
  return String(n).padStart(2, '0')
}
function formatIsoToLocalKorean(iso) {
  if (!iso) return '-'
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return String(iso)
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(
    d.getDate()
  )} (${pad2(d.getHours())}:${pad2(d.getMinutes())})`
}
function Score(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return '-'
  return v.toFixed(2)
}

function computeScanStats({ matchData, nerItems, nerLabels, timings }) {
  const mdItems = Array.isArray(matchData?.items) ? matchData.items : []
  const ner = Array.isArray(nerItems) ? nerItems : []
  const allow = new Set((nerLabels || []).map((x) => String(x).toUpperCase()))

  let regex_ok = 0,
    regex_fail = 0
  const by_kind = {}
  for (const it of mdItems) {
    if (it?.valid) {
      regex_ok++
      const k = ruleToKind(it.rule)
      by_kind[k] = (by_kind[k] || 0) + 1
    } else regex_fail++
  }

  const by_label = {}
  const scores = []
  let nerAllowedCount = 0
  for (const it of ner) {
    const lab = String(it?.label || '').toUpperCase()
    if (!allow.has(lab)) continue
    nerAllowedCount++
    by_label[lab] = (by_label[lab] || 0) + 1
    if (typeof it?.score === 'number') scores.push(it.score)
  }

  // (기존 로직 유지) 점수는 지표일 뿐이니 대충: 정규식 OK 가중 + NER 허용 라벨 수 가중
  const risk = Math.min(100, Math.round(regex_ok * 10 + nerAllowedCount * 2))
  const nerAvg = scores.length
    ? scores.reduce((a, b) => a + b, 0) / scores.length
    : null

  return {
    risk_score: risk,
    total_raw: mdItems.length + nerAllowedCount,
    total_unique: Array.isArray(state.detections) ? state.detections.length : 0,
    regex_ok,
    regex_fail,
    by_kind,
    by_label,
    ner_avg: Score(nerAvg),
    timings: timings || {},
  }
}

function renderScanReport(stats) {
  if (!stats) return

  // stats 블록은 있어도 되고 없어도 됨(없으면 그냥 스킵)
  safeShow('stats-report-block', true)

  safeText('stats-risk-score', stats.risk_score)
  safeWidthPct('stats-risk-meter', `${stats.risk_score}%`)
  safeText('stats-total-unique', stats.total_unique)
  safeText('stats-total-raw', stats.total_raw)
  safeText('stats-regex-ok', stats.regex_ok)
  safeText('stats-regex-fail', stats.regex_fail)
  safeText('stats-ner-avg', stats.ner_avg)

  const kindBody = byId('stats-by-kind-rows')
  if (kindBody) {
    kindBody.innerHTML = Object.entries(stats.by_kind || {})
      .sort((a, b) => b[1] - a[1])
      .map(
        ([k, v]) =>
          `<tr>
            <td class="py-2 pl-4 pr-2 font-medium text-zinc-900">${escHtml(
              k
            )}</td>
            <td class="py-2 pr-5 text-right font-bold text-zinc-500">${escHtml(
              v
            )}</td>
          </tr>`
      )
      .join('')
  }

  const labelBody = byId('stats-by-label-rows')
  if (labelBody) {
    labelBody.innerHTML = Object.entries(stats.by_label || {})
      .sort((a, b) => b[1] - a[1])
      .map(
        ([k, v]) =>
          `<tr>
            <td class="py-2 pl-4 pr-2 font-medium text-zinc-900">${escHtml(
              k
            )}</td>
            <td class="py-2 pr-5 text-right font-bold text-zinc-500">${escHtml(
              v
            )}</td>
          </tr>`
      )
      .join('')
  }

  safeText('t-extract', Math.round(stats.timings.extract_ms || 0) + 'ms')
  safeText('t-match', Math.round(stats.timings.match_ms || 0) + 'ms')
  safeText('t-ner', Math.round(stats.timings.ner_ms || 0) + 'ms')
  safeText('t-redact', Math.round(stats.timings.redact_ms || 0) + 'ms')
  safeText('t-total', Math.round(stats.timings.total_ms || 0) + 'ms')

  // 정책(선택한 규칙/라벨 표시)
  const selectedRules = Array.isArray(state.rules) ? state.rules : []
  const selectedNer = Array.isArray(state.nerLabels) ? state.nerLabels : []

  const detectedRuleSet = new Set()
  const mdItems = Array.isArray(state.matchData?.items)
    ? state.matchData.items
    : []
  for (const it of mdItems) {
    if (it?.valid === false) continue
    if (!it?.rule) continue
    detectedRuleSet.add(String(it.rule))
  }

  const rulesBox = byId('stats-policy-rules')
  if (rulesBox) {
    rulesBox.innerHTML = selectedRules
      .map((r) => {
        const isHit = detectedRuleSet.has(String(r))
        const label = ruleToKind(r)
        const cls = isHit
          ? 'border-indigo-200 bg-indigo-50 text-indigo-700'
          : 'border-gray-200 bg-gray-50 text-gray-600'
        return `<span class="inline-flex items-center gap-1 px-2 py-1 rounded-full border text-[11px] ${cls}">${escHtml(
          label
        )}</span>`
      })
      .join('')
  }

  const nerBox = byId('stats-policy-nerlabels')
  if (nerBox) {
    nerBox.innerHTML = selectedNer
      .map(
        (lab) =>
          `<span class="inline-flex items-center px-2 py-1 rounded-full border border-gray-200 bg-gray-50 text-gray-600 text-[11px]">${escHtml(
            String(lab).toUpperCase()
          )}</span>`
      )
      .join('')
  }

  safeText('stats-json', JSON.stringify(stats, null, 2))
}

/** ---------- Main: Scan ---------- */
async function doScan() {
  const f = $('#file')?.files?.[0]
  if (!f) return alert('파일을 선택하세요.')

  state.file = f
  state.ext = (f.name.split('.').pop() || '').toLowerCase()
  state.rules = selectedRuleNames()
  state.nerLabels = selectedNerLabels()
  state.t0 = performance.now()

  setStatus('분석 시작...')
  lockInputs(true)

  try {
    const fd = new FormData()
    fd.append('file', f)

    const t1 = performance.now()
    const r1 = await fetch(`${API_BASE()}/text/extract`, {
      method: 'POST',
      body: fd,
    })
    if (!r1.ok) throw new Error('추출 실패')
    const extractData = await r1.json()
    state.timings = { extract_ms: performance.now() - t1 }

    const fullText = String(extractData.full_text || '')
    state.extractedText = fullText

    const md = extractData.markdown || fallbackMarkdownFromText(fullText)
    state.markdown = md
    setPages([md])

    const t2 = performance.now()
    setStatus('패턴 탐색...')
    const r2 = await fetch(`${API_BASE()}/text/match`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        text: fullText,
        rules: state.rules,
        normalize: true,
      }),
    })
    const rawMatchData = await r2.json()
    state.timings.match_ms = performance.now() - t2
    state.matchData = filterMatchByRules(rawMatchData, state.rules)

    const t3 = performance.now()
    setStatus('NER 탐지...')
    const nerResp = await requestNerSmart(
      fullText,
      buildExcludeSpansFromMatch(state.matchData),
      state.nerLabels
    )
    state.timings.ner_ms = performance.now() - t3
    state.nerItems = nerResp.items

    state.detections = buildDetections(
      state.matchData,
      state.nerItems,
      state.nerLabels
    )
    state.detectionById = new Map(state.detections.map((d) => [d.id, d]))

    safeClassRemove('doc-viewer-block', 'hidden')
    safeClassRemove('match-tabs-block', 'hidden')
    safeText(
      'doc-meta',
      `${f.name} · ${state.ext.toUpperCase()} · ${Math.round(f.size / 1024)}KB`
    )
    safeText('doc-detect-count', state.detections.length)

    // 오른쪽 badge(현재 탭은 setMatchTab에서 갱신)
    const mb = byId('match-badge')
    if (mb) {
      mb.textContent = String(
        state.detections.filter((d) => d.source === 'regex').length
      )
    }

    renderCurrentPage()
    wireViewerClick()
    renderMatchResults()
    renderNerResults()
    setMatchTab('regex')

    state.timings.total_ms = performance.now() - state.t0
    renderScanReport(
      computeScanStats({
        matchData: state.matchData,
        nerItems: state.nerItems,
        nerLabels: state.nerLabels,
        timings: state.timings,
      })
    )

    setStatus('완료')

    const btn = $('#btn-save-redacted')
    if (btn) {
      btn.classList.remove('hidden')
      btn.disabled = false
    }
  } catch (e) {
    console.error(e)
    setStatus('오류 발생')
  } finally {
    lockInputs(false)
  }
}

/** ---------- Redact + Download ---------- */
async function doRedactAndDownload() {
  const f = $('#file')?.files?.[0]
  if (!f) return alert('파일을 선택하세요.')

  if (!state.file || state.file !== f || !state.extractedText) {
    await doScan()
  }
  if (!state.file) return

  const btn = $('#btn-save-redacted')
  btn && (btn.disabled = true)

  setStatus('레닥션 실행 중...')
  lockInputs(true)

  const t0 = performance.now()
  try {
    const fd = new FormData()
    fd.append('file', state.file)

    const rulesJson = safeJson(state.rules || [])
    const labelsJson = safeJson(state.nerLabels || [])
    const entsJson = safeJson(state.nerItems || [])

    rulesJson && fd.append('rules_json', rulesJson)
    labelsJson && fd.append('ner_labels_json', labelsJson)
    entsJson && fd.append('ner_entities_json', entsJson)

    const r = await fetch(`${API_BASE()}/redact/file`, {
      method: 'POST',
      body: fd,
    })
    if (!r.ok) {
      const msg = await r.text().catch(() => '')
      throw new Error(msg || '레닥션 실패')
    }

    const blob = await r.blob()
    const cd = r.headers.get('Content-Disposition')
    const filename =
      parseContentDispositionFilename(cd) ||
      buildRedactedFallbackName(state.file?.name)

    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    a.remove()
    setTimeout(() => URL.revokeObjectURL(url), 1500)

    state.timings = state.timings || {}
    state.timings.redact_ms = performance.now() - t0
    safeText('t-redact', Math.round(state.timings.redact_ms) + 'ms')

    if (state.t0) {
      state.timings.total_ms = performance.now() - state.t0
      safeText('t-total', Math.round(state.timings.total_ms) + 'ms')
    }

    setStatus('레닥션 완료 · 다운로드 시작')
  } catch (e) {
    console.error(e)
    alert(`레닥션 실패: ${e?.message || e}`)
    setStatus('레닥션 오류')
  } finally {
    btn && (btn.disabled = false)
    lockInputs(false)
  }
}

function renderCurrentPage() {
  updatePageControls()
  renderMarkdownToViewer(state.markdown, state.detections)
  applyDocOrientationHint(state.markdown, $('#doc-viewer'))
}

/** ---------- Init ---------- */
document.addEventListener('DOMContentLoaded', () => {
  loadRules()
  setupDropZone()
})