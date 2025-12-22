const API_BASE = () => window.API_BASE || 'http://127.0.0.1:8000'
const HWPX_VIEWER_URL = window.HWPX_VIEWER_URL || ''

const $ = (sel) => document.querySelector(sel)
const $$ = (sel) => Array.from(document.querySelectorAll(sel))

let __lastRedactedBlob = null
let __lastRedactedName = 'redacted.bin'
let __lastNerEntities = []
let __lastScanStats = null

// 통계 갱신용 컨텍스트(체크박스 변경 시 재계산)
let __lastStatsContext = null // { file, ext, rules, matchData, timings, t0 }

const esc = (s) =>
  (s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

const badge = (sel, n) => {
  const el = $(sel)
  if (el) el.textContent = String(n ?? 0)
}

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

function selectedNerLabels() {
  const labels = []
  $('#ner-show-ps')?.checked !== false && labels.push('PS')
  $('#ner-show-lc')?.checked !== false && labels.push('LC')
  $('#ner-show-og')?.checked !== false && labels.push('OG')
  return labels
}

function setupDropZone() {
  const dz = $('#dropzone'),
    input = $('#file'),
    nameEl = $('#file-name'),
    statusEl = $('#status')
  if (!dz || !input) return

  let depth = 0
  const setActive = (on) => {
    dz.classList.toggle('ring-2', on)
    // Tailwind CDN이 외부 JS의 class string을 못 읽는 경우가 있어,
    // 색상은 inline style로 직접 토글한다.
    dz.style.setProperty('--tw-ring-color', on ? '#b7a3e3' : '')
    dz.style.backgroundColor = on ? '#f7f3ff' : ''
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

function highlightFrag(ctx, val) {
  const src = ctx || ''
  const needle = val || ''

  const i = src.indexOf(needle)
  if (i < 0) return esc(src)

  const pre = esc(src.slice(0, i))
  const mid = esc(needle)
  const post = esc(src.slice(i + needle.length))

  return (
    pre +
    `<mark class="rounded px-1" style="background-color:#e2d6fb">${mid}</mark>` +
    post
  )
}

let __segFilter = 'all'
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
      btn.classList.remove('text-white')
      btn.style.backgroundColor = ''
      btn.style.color = ''
      if (k === which) {
        btn.classList.add('text-white')
        btn.style.backgroundColor =
          which === 'all' ? '#533a8c' : which === 'ok' ? '#9e86d7' : '#8569cb'
      }
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
          <span class="text-[10px] px-1.5 py-0.5 rounded" style="background-color:#efe8ff;color:#533a8c">OK ${ok}</span>
          ${
            fail
              ? `<span class="text-[10px] px-1.5 py-0.5 rounded" style="background-color:#e2d6fb;color:#3a2a62">FAIL ${fail}</span>`
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
        'border rounded-xl p-3 bg-white hover:shadow-sm transition'
      card.style.borderColor = isOk ? '#e2d6fb' : '#d2c0f5'

      card.innerHTML = `
        <div class="flex items-start justify-between gap-3">
          <div class="min-w-0">
            <div class="text-sm font-mono break-all">${esc(val)}</div>
            <div class="text-[12px] text-gray-600 mt-1 leading-relaxed break-words">
              ${highlightFrag(ctx, val)}
            </div>
          </div>
          <div class="shrink-0">
            <span class="inline-block text-[11px] px-1.5 py-0.5 rounded border" style="border-color:${
              isOk ? '#d2c0f5' : '#b7a3e3'
            };color:${isOk ? '#533a8c' : '#3a2a62'}">${
        isOk ? 'OK' : 'FAIL'
      }</span>
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

function parseMarkdownTables(markdown) {
  const lines = (markdown || '').split(/\r?\n/)
  const tables = []
  let current = []

  for (const raw of lines) {
    const line = raw.trim()
    if (/^\s*\|.*\|\s*$/.test(line)) {
      current.push(line)
    } else {
      if (current.length >= 2) tables.push(current.slice())
      current = []
    }
  }
  if (current.length >= 2) tables.push(current.slice())
  return tables
}

function tableBlockToHtml(block) {
  if (!block || block.length < 2) return ''
  const headerLine = block[0]
  const headerCells = headerLine
    .split('|')
    .slice(1, -1)
    .map((s) => s.trim())
  const bodyLines = block.slice(2)

  let html =
    '<table class="min-w-full border border-gray-300 text-[11px] text-left border-collapse">'
  html += '<thead><tr>'
  for (const h of headerCells) {
    html += `<th class="border border-gray-300 px-2 py-1 bg-gray-50">${esc(
      h
    )}</th>`
  }
  html += '</tr></thead><tbody>'

  for (const line of bodyLines) {
    const cells = line
      .split('|')
      .slice(1, -1)
      .map((s) => s.trim())
    if (!cells.length) continue
    html += '<tr>'
    for (const c of cells) {
      html += `<td class="border border-gray-300 px-2 py-1 align-top">${esc(
        c
      )}</td>`
    }
    html += '</tr>'
  }
  html += '</tbody></table>'
  return html
}

function buildPlainTextFromMarkdown(markdown) {
  const lines = (markdown || '').split(/\r?\n/)
  const out = []
  let inTable = false

  for (const raw of lines) {
    const line = raw.trim()
    const isTableRow = /^\s*\|.*\|\s*$/.test(line)

    if (isTableRow) {
      inTable = true
      continue
    }

    if (!isTableRow && inTable) inTable = false
    if (!inTable) out.push(raw)
  }

  return out.join('\n').trim()
}

function renderTablePreview(markdown) {
  const wrap = $('#text-table-preview')
  if (!wrap) return 0
  wrap.innerHTML = ''

  const blocks = parseMarkdownTables(markdown)
  if (!blocks.length) return 0

  const parts = []
  parts.push(
    '<div class="text-[11px] text-gray-500 mb-1">표 구조 미리보기</div>'
  )
  blocks.forEach((block, idx) => {
    if (idx > 0) parts.push('<div class="h-2"></div>')
    parts.push(tableBlockToHtml(block))
  })
  wrap.innerHTML = parts.join('')
  return blocks.length
}

const JOIN_NEWLINE_RE = /([\w\uAC00-\uD7A3.%+\-/])\n([\w\uAC00-\uD7A3.%+\-/])/g
function joinBrokenLines(text) {
  if (!text) return ''
  let t = String(text).replace(/\r\n/g, '\n')
  let prev = null
  while (prev !== t) {
    prev = t
    t = t.replace(JOIN_NEWLINE_RE, '$1$2')
  }
  return t
}

function normalizeNerItems(raw) {
  if (!raw) return { items: [] }
  if (Array.isArray(raw.entities)) return { items: raw.entities }
  if (Array.isArray(raw.items)) return { items: raw.items }
  if (Array.isArray(raw)) return { items: raw }
  return { items: [] }
}
async function requestNerSmart(
  text,
  exclude_spans,
  debug = false,
  labels_override = null
) {
  const labels =
    Array.isArray(labels_override) && labels_override.length
      ? labels_override
      : selectedNerLabels()

  const bodyObj = {
    text: String(text || ''),
    labels: labels,
    exclude_spans: Array.isArray(exclude_spans) ? exclude_spans : [],
    debug: !!debug,
  }

  try {
    const r2 = await fetch(`${API_BASE()}/ner/predict`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyObj),
    })

    if (!r2.ok) {
      const txt = await r2.text()
      console.error('NER 요청 실패', r2.status, txt)
      setStatus(`NER 분석 실패 (${r2.status})`)
      return { items: [] }
    }

    const j2 = await r2.json()
    return normalizeNerItems(j2)
  } catch (e) {
    console.error('NER 요청 중 오류', e)
    setStatus(`NER 분석 중 오류: ${e.message || e}`)
    return { items: [] }
  }
}

function Score(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return '-'
  const t = Math.floor(v * 100) / 100
  return t.toFixed(2)
}

function renderNerTable(ner) {
  const rows = $('#ner-rows')
  const sum = $('#ner-summary')
  const allow = new Set()
  $('#ner-show-ps')?.checked !== false && allow.add('PS')
  $('#ner-show-lc')?.checked !== false && allow.add('LC')
  $('#ner-show-og')?.checked !== false && allow.add('OG')

  const items = (ner.items || []).filter((it) =>
    allow.has((it.label || '').toUpperCase())
  )

  if (rows) rows.innerHTML = ''
  for (const it of items) {
    const tr = document.createElement('tr')
    tr.className = 'border-b align-top'
    tr.innerHTML = `
      <td class="py-2 px-2 font-mono">${esc(it.label)}</td>
      <td class="py-2 px-2 font-mono">${esc(it.text)}</td>
      <td class="py-2 px-2 font-mono">${Score(it.score)}</td>
      <td class="py-2 px-2 font-mono">${it.start}-${it.end}</td>`
    rows?.appendChild(tr)
  }

  badge('#ner-badge', items.length)
  if (sum) {
    const counts = {}
    for (const it of items) counts[it.label] = (counts[it.label] || 0) + 1
    sum.textContent = `검출: ${
      Object.keys(counts).length
        ? Object.entries(counts)
            .map(([k, v]) => `${k}=${v}`)
            .join(', ')
        : '없음'
    }`
  }
}

function setStatus(msg) {
  const el = $('#status')
  if (el) el.textContent = msg || ''
}

function fmtMs(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return '-'
  return `${Math.round(v)}ms`
}

function pad2(n) {
  return String(n).padStart(2, '0')
}

// ISO 8601(…Z 포함)을 로컬 기준 YYYY-MM-DD (HH:MM)로 표시
function formatIsoToLocalKorean(iso) {
  if (!iso) return '-'
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return String(iso)
  const y = d.getFullYear()
  const m = pad2(d.getMonth() + 1)
  const day = pad2(d.getDate())
  const hh = pad2(d.getHours())
  const mm = pad2(d.getMinutes())
  return `${y}-${m}-${day} (${hh}:${mm})`
}

function nowIso() {
  const d = new Date()
  return d.toISOString()
}

function ruleToKind(rule) {
  const r = String(rule || '').toLowerCase()
  if (!r) return 'UNKNOWN'
  if (r.includes('rrn')) return '주민등록번호'
  if (r.includes('fgn')) return '외국인등록번호'
  if (r.includes('card')) return '카드번호'
  if (r.includes('email')) return '이메일'
  if (r.includes('passport')) return '여권번호'
  if (r.includes('driver')) return '운전면허번호'
  if (r.includes('account') || r.includes('bank')) return '계좌번호'
  if (r.includes('phone') || r.includes('tel') || r.includes('mobile'))
    return '전화번호'
  return String(rule)
}

function riskWeights() {
  return {
    주민등록번호: 30,
    외국인등록번호: 30,
    카드번호: 25,
    계좌번호: 20,
    운전면허번호: 18,
    여권번호: 18,
    전화번호: 10,
    이메일: 8,
    PS: 2,
    LC: 5,
    OG: 3,
  }
}

function onlyDigits(s) {
  return String(s || '').replace(/\D+/g, '')
}

// Luhn 체크(카드번호 등) — 체크디짓 검증 :contentReference[oaicite:0]{index=0}
function luhnValid(numStr) {
  const s = onlyDigits(numStr)
  if (s.length < 12) return false
  let sum = 0
  let alt = false
  for (let i = s.length - 1; i >= 0; i--) {
    let n = s.charCodeAt(i) - 48
    if (n < 0 || n > 9) return false
    if (alt) {
      n *= 2
      if (n > 9) n -= 9
    }
    sum += n
    alt = !alt
  }
  return sum % 10 === 0
}

// 이메일 형식(간단 검사): local-part@domain 기본 구조 :contentReference[oaicite:1]{index=1}
function emailLooksValid(v) {
  const s = String(v || '').trim()
  if (!s.includes('@')) return false
  const parts = s.split('@')
  if (parts.length !== 2) return false
  const [local, domain] = parts
  if (!local || !domain) return false
  if (domain.startsWith('.') || domain.endsWith('.')) return false
  if (!domain.includes('.')) return false
  return true
}

function rrnLooksValidFormat(v) {
  const d = onlyDigits(v)
  return d.length === 13
}

function phoneLooksValid(v) {
  const d = onlyDigits(v)
  // 한국 전화번호는 9~11자리 범위가 흔함(대표적으로 010 포함 11) (정확하지 않음)
  return d.length >= 9 && d.length <= 11
}

function inferFailReason(rule, value) {
  const r = String(rule || '').toLowerCase()
  const v = String(value || '')

  if (r.includes('card')) {
    if (!onlyDigits(v)) return '숫자 외 문자 포함'
    if (!luhnValid(v)) return ' Luhn 불일치'
    return '검증 로직 불일치'
  }

  if (r.includes('email')) {
    if (!emailLooksValid(v)) return '이메일 형식 불일치'
    return '검증 로직 불일치'
  }

  if (r.includes('rrn')) {
    if (!rrnLooksValidFormat(v)) return '자리수/형식 불일치'
    return '체크섬 불일치'
  }

  if (r.includes('fgn')) {
    const d = onlyDigits(v)
    if (d.length !== 13) return '자리수/형식 불일치'
    return '체크섬 불일치'
  }

  if (r.includes('phone') || r.includes('tel') || r.includes('mobile')) {
    if (!phoneLooksValid(v)) return '전화번호 자리수/형식 불일치'
    return '검증 로직 불일치'
  }

  if (r.includes('passport')) {
    return '여권번호 형식/국가규칙 불일치'
  }

  if (r.includes('driver')) {
    return '면허번호 형식/지역규칙 불일치'
  }

  return '검증 실패(원인 미분류)'
}

function scoreNote(score) {
  if (!(score >= 0)) return '-'
  if (score >= 80) return 'HIGH'
  if (score >= 45) return 'MEDIUM'
  return 'LOW'
}

function overlapRatio(a, b) {
  const a0 = Number(a?.start ?? -1)
  const a1 = Number(a?.end ?? -1)
  const b0 = Number(b?.start ?? -1)
  const b1 = Number(b?.end ?? -1)
  if (!(a1 > a0 && b1 > b0)) return 0
  const inter = Math.max(0, Math.min(a1, b1) - Math.max(a0, b0))
  const denom = Math.min(a1 - a0, b1 - b0)
  if (denom <= 0) return 0
  return inter / denom
}

function computeScanStats({
  file,
  ext,
  rules,
  nerLabels,
  matchData,
  nerItems,
  timings,
}) {
  const weights = riskWeights()
  const mdItems = Array.isArray(matchData?.items) ? matchData.items : []
  const ner = Array.isArray(nerItems) ? nerItems : []

  // 1) 정규식 OK/FAIL + FAIL 원인(추정)
  let regex_ok = 0
  let regex_fail = 0
  const fail_reason_counts = {}

  for (const it of mdItems) {
    if (it?.valid === true) {
      regex_ok++
    } else if (it?.valid === false) {
      regex_fail++
      const reason = inferFailReason(it?.rule, it?.value)
      fail_reason_counts[reason] = (fail_reason_counts[reason] || 0) + 1
    }
  }

  const fail_top3 =
    Object.entries(fail_reason_counts)
      .sort((a, b) => (b[1] || 0) - (a[1] || 0))
      .slice(0, 3)
      .map(([k, v]) => `${k} ${v}`)
      .join(' / ') || '-'

  // 2) by_kind (정규식 valid=true 기준)
  const by_kind = {}
  for (const it of mdItems) {
    if (it?.valid === false) continue
    const k = ruleToKind(it?.rule)
    by_kind[k] = (by_kind[k] || 0) + 1
  }

  // 3) by_label (선택된 라벨 기준)
  const allow = new Set((nerLabels || []).map((x) => String(x).toUpperCase()))
  const by_label = {}
  const scoreAgg = {}
  for (const it of ner) {
    const lab = String(it?.label || '').toUpperCase()
    if (!allow.has(lab)) continue
    by_label[lab] = (by_label[lab] || 0) + 1
    if (typeof it?.score === 'number' && Number.isFinite(it.score)) {
      ;(scoreAgg[lab] ??= []).push(it.score)
    }
  }

  // 4) NER 평균 score
  const nerAvg = {}
  for (const lab of ['PS', 'LC', 'OG']) {
    const arr = scoreAgg[lab] || []
    if (!arr.length) nerAvg[lab] = '-'
    else {
      const avg = arr.reduce((a, b) => a + b, 0) / arr.length
      nerAvg[lab] = Score(avg)
    }
  }

  // 5) Unique(중복 제거): regex(valid)+ner(selected) span을 겹침 기준으로 병합
  const spans = []
  for (const it of mdItems) {
    if (it?.valid === false) continue
    const s = Number(it?.start ?? -1)
    const e = Number(it?.end ?? -1)
    if (!(e > s)) continue
    spans.push({ start: s, end: e, source: 'regex' })
  }
  for (const it of ner) {
    const lab = String(it?.label || '').toUpperCase()
    if (!allow.has(lab)) continue
    const s = Number(it?.start ?? -1)
    const e = Number(it?.end ?? -1)
    if (!(e > s)) continue
    spans.push({ start: s, end: e, source: 'ner' })
  }
  spans.sort((a, b) => a.start - b.start || a.end - b.end)

  const clusters = []
  for (const sp of spans) {
    const last = clusters.length ? clusters[clusters.length - 1] : null
    if (!last) {
      clusters.push({
        start: sp.start,
        end: sp.end,
        has_regex: sp.source === 'regex',
        has_ner: sp.source === 'ner',
      })
      continue
    }
    const ratio = overlapRatio(last, sp)
    if (ratio >= 0.8 || sp.start <= last.end) {
      last.start = Math.min(last.start, sp.start)
      last.end = Math.max(last.end, sp.end)
      if (sp.source === 'regex') last.has_regex = true
      if (sp.source === 'ner') last.has_ner = true
    } else {
      clusters.push({
        start: sp.start,
        end: sp.end,
        has_regex: sp.source === 'regex',
        has_ner: sp.source === 'ner',
      })
    }
  }

  const total_raw = spans.length
  const total_unique = clusters.length
  const overlap_count = clusters.filter((c) => c.has_regex && c.has_ner).length
  const overlap_rate = total_unique ? overlap_count / total_unique : 0

  // 정규식과 NER 개수 분리
  const regex_unique = clusters.filter((c) => c.has_regex).length
  const ner_unique = clusters.filter((c) => c.has_ner).length

  let raw = 0
  for (const [k, v] of Object.entries(by_kind))
    raw += (weights[k] || 0) * (v || 0)
  for (const [lab, v] of Object.entries(by_label))
    raw += (weights[lab] || 0) * (v || 0)
  const risk_score_0_100 = Math.max(0, Math.min(100, Math.round(raw)))

  const policy = {
    rules: Array.isArray(rules) ? rules : [],
    ner_labels: Array.isArray(nerLabels) ? nerLabels : [],
  }

  // 탐지된 규칙과 라벨 수집
  const detectedRules = new Set()
  for (const it of mdItems) {
    if (it?.valid !== false && it?.rule) {
      detectedRules.add(String(it.rule).toLowerCase())
    }
  }
  const detectedLabels = new Set(
    Object.keys(by_label || {}).map((l) => l.toUpperCase())
  )

  return {
    version: 'scan-report-v1',
    created_at: nowIso(),
    document: {
      name: file?.name || '-',
      ext: String(ext || '').toLowerCase(),
      size_bytes: typeof file?.size === 'number' ? file.size : null,
    },
    policy,
    timings: { ...(timings || {}) },
    stats: {
      risk_score_0_100,
      total_raw,
      total_unique,
      regex_unique,
      ner_unique,
      overlap_count,
      overlap_rate,
      by_kind,
      by_label,
      ner_avg: nerAvg,

      // 추가된 필드
      regex_ok,
      regex_fail,
      fail_top3,
      fail_reason_counts,
      detected_rules: Array.from(detectedRules),
      detected_labels: Array.from(detectedLabels),
    },
  }
}

function renderChips(containerEl, list, detectedSet = null) {
  if (!containerEl) return
  containerEl.innerHTML = ''
  const arr = Array.isArray(list) ? list : []
  if (!arr.length) {
    containerEl.innerHTML = `<span class="text-[12px] text-gray-500">-</span>`
    return
  }
  for (const v of arr) {
    const span = document.createElement('span')
    // 정규식 규칙은 소문자로, NER 라벨은 대문자로 비교
    const vLower = String(v).toLowerCase()
    const vUpper = String(v).toUpperCase()
    const isDetected =
      detectedSet && (detectedSet.has(vLower) || detectedSet.has(vUpper))
    // 탐지된 항목은 보라색으로 강조
    if (isDetected) {
      span.className = 'text-[11px] px-2 py-1 rounded-full border font-semibold'
      span.style.borderColor = '#d2c0f5'
      span.style.backgroundColor = '#efe8ff'
      span.style.color = '#533a8c'
    } else {
      span.className =
        'text-[11px] px-2 py-1 rounded-full border border-gray-200 bg-gray-50 text-gray-800'
    }
    span.textContent = String(v)
    containerEl.appendChild(span)
  }
}

function renderScanReport(report) {
  if (!report) return
  __lastScanStats = report

  $('#stats-report-block')?.classList.remove('hidden')

  const doc = report.document || {}
  const s = report.stats || {}
  const t = report.timings || {}
  const pol = report.policy || {}

  const created = formatIsoToLocalKorean(report.created_at)
  const subtitle = `${doc.name || '-'} · ${String(
    doc.ext || ''
  ).toUpperCase()} · ${
    typeof doc.size_bytes === 'number' ? `${doc.size_bytes} bytes` : '-'
  } · ${created}`
  const subEl = $('#stats-subtitle')
  subEl && (subEl.textContent = subtitle)

  // Risk score
  const risk = Number(s.risk_score_0_100 ?? 0)
  $('#stats-risk-score') && ($('#stats-risk-score').textContent = String(risk))
  const meter = $('#stats-risk-meter')
  meter && (meter.style.width = `${Math.max(0, Math.min(100, risk))}%`)

  // Unique / Raw / Overlap
  $('#stats-total-unique') &&
    ($('#stats-total-unique').textContent = String(s.total_unique ?? 0))
  $('#stats-total-raw') &&
    ($('#stats-total-raw').textContent = String(s.total_raw ?? 0))
  $('#stats-overlap-rate') &&
    ($('#stats-overlap-rate').textContent = `${Math.round(
      (s.overlap_rate || 0) * 100
    )}%`)

  // 정규식과 NER 개수 분리 표시
  $('#stats-regex-unique') &&
    ($('#stats-regex-unique').textContent = String(s.regex_unique ?? 0))
  $('#stats-ner-unique') &&
    ($('#stats-ner-unique').textContent = String(s.ner_unique ?? 0))

  // NER 평균
  const nerAvg = s.ner_avg || {}
  $('#stats-ner-avg') &&
    ($('#stats-ner-avg').textContent = `PS ${nerAvg.PS || '-'} / LC ${
      nerAvg.LC || '-'
    } / OG ${nerAvg.OG || '-'}`)

  // 정규식 OK/FAIL + FAIL 원인 TOP3 (TOP3 카드 제거 후 이걸 사용)
  $('#stats-regex-ok') &&
    ($('#stats-regex-ok').textContent = String(s.regex_ok ?? 0))
  $('#stats-regex-fail') &&
    ($('#stats-regex-fail').textContent = String(s.regex_fail ?? 0))
  $('#stats-fail-top') &&
    ($('#stats-fail-top').textContent = s.fail_top3 || '-')

  // 정책(칩) - 탐지된 항목은 보라색으로 강조
  const detectedRules = new Set(
    (s.detected_rules || []).map((r) => String(r).toLowerCase())
  )
  const detectedLabels = new Set(
    (s.detected_labels || []).map((l) => String(l).toUpperCase())
  )

  renderChips($('#stats-policy-rules'), pol.rules, detectedRules)
  renderChips($('#stats-policy-nerlabels'), pol.ner_labels, detectedLabels)

  // 처리시간(표)
  $('#t-extract') && ($('#t-extract').textContent = fmtMs(t.extract_ms))
  $('#t-match') && ($('#t-match').textContent = fmtMs(t.match_ms))
  $('#t-ner') && ($('#t-ner').textContent = fmtMs(t.ner_ms))
  $('#t-redact') && ($('#t-redact').textContent = fmtMs(t.redact_ms))
  $('#t-total') && ($('#t-total').textContent = fmtMs(t.total_ms))

  // by_kind 표
  const kindRows = $('#stats-by-kind-rows')
  if (kindRows) {
    kindRows.innerHTML = ''
    const entries = Object.entries(s.by_kind || {}).sort(
      (a, b) => (b[1] || 0) - (a[1] || 0)
    )
    if (!entries.length) {
      const tr = document.createElement('tr')
      tr.innerHTML = `<td class="py-2 px-3 text-gray-500" colspan="2">없음</td>`
      kindRows.appendChild(tr)
    } else {
      for (const [k, v] of entries) {
        const tr = document.createElement('tr')
        tr.className = 'border-b'
        tr.innerHTML = `<td class="py-2 px-3">${esc(
          k
        )}</td><td class="py-2 px-3 text-right font-mono">${v}</td>`
        kindRows.appendChild(tr)
      }
    }
  }

  // by_label 표
  const labelRows = $('#stats-by-label-rows')
  if (labelRows) {
    labelRows.innerHTML = ''
    const entries = Object.entries(s.by_label || {}).sort(
      (a, b) => (b[1] || 0) - (a[1] || 0)
    )
    if (!entries.length) {
      const tr = document.createElement('tr')
      tr.innerHTML = `<td class="py-2 px-3 text-gray-500" colspan="2">없음</td>`
      labelRows.appendChild(tr)
    } else {
      for (const [k, v] of entries) {
        const tr = document.createElement('tr')
        tr.className = 'border-b'
        tr.innerHTML = `<td class="py-2 px-3">${esc(
          k
        )}</td><td class="py-2 px-3 text-right font-mono">${v}</td>`
        labelRows.appendChild(tr)
      }
    }
  }

  // JSON 프리뷰(열려 있을 때만 갱신)
  const jsonEl = $('#stats-json')
  if (jsonEl && !jsonEl.classList.contains('hidden')) {
    jsonEl.textContent = JSON.stringify(report, null, 2)
  }
}

function refreshStatsFromContext() {
  if (!__lastStatsContext) return
  const { file, ext, rules, matchData, timings, t0 } = __lastStatsContext
  const nextTimings = { ...(timings || {}) }
  if (typeof t0 === 'number') nextTimings.total_ms = performance.now() - t0

  const report = computeScanStats({
    file,
    ext,
    rules,
    nerLabels: selectedNerLabels(),
    matchData,
    nerItems: __lastNerEntities,
    timings: nextTimings,
  })
  renderScanReport(report)
}

function filterMatchByRules(matchData, rules) {
  const allow = new Set((rules || []).map((r) => String(r).toLowerCase()))
  const items = Array.isArray(matchData?.items) ? matchData.items : []
  const kept = allow.size
    ? items.filter((it) => allow.has(String(it.rule || '').toLowerCase()))
    : items
  const counts = {}
  for (const it of kept) {
    if (it?.valid) counts[it.rule] = (counts[it.rule] || 0) + 1
  }
  return { ...matchData, items: kept, counts }
}

function buildExcludeSpansFromMatch(matchData) {
  const items = Array.isArray(matchData?.items) ? matchData.items : []
  const spans = []
  for (const it of items) {
    if (it?.valid === false) continue
    const s = Number(it.start ?? -1)
    const e = Number(it.end ?? -1)
    if (!(e > s)) continue
    spans.push({ start: s, end: e })
  }
  return spans
}

$('#btn-scan')?.addEventListener('click', async () => {
  const f = $('#file')?.files?.[0]
  if (!f) return alert('파일을 선택하세요.')

  const ext = (f.name.split('.').pop() || '').toLowerCase()
  __lastRedactedName = f.name
    ? f.name.replace(/\.[^.]+$/, `_redacted.${ext}`)
    : `redacted.${ext}`

  const t0 = performance.now()
  const timings = {
    extract_ms: null,
    match_ms: null,
    ner_ms: null,
    redact_ms: null,
    total_ms: null,
  }

  setStatus('텍스트 추출 중...')
  const fd = new FormData()
  fd.append('file', f)

  $('#match-result-block')?.classList.remove('hidden')
  $('#ner-result-block')?.classList.remove('hidden')

  try {
    const tExtract = performance.now()
    const r1 = await fetch(`${API_BASE()}/text/extract`, {
      method: 'POST',
      body: fd,
    })
    if (!r1.ok)
      throw new Error(`텍스트 추출 실패 (${r1.status})\n${await r1.text()}`)
    const extractData = await r1.json()
    timings.extract_ms = performance.now() - tExtract

    const fullText = extractData.full_text || ''
    let analysisText = fullText || ''

    $('#text-preview-block')?.classList.remove('hidden')
    const ta = $('#txt-out')
    if (ta) {
      ta.classList.remove('hidden')
      ta.value = analysisText || '(본문 텍스트가 비어 있습니다.)'
    }

    const tablePreviewRoot = $('#text-table-preview')
    if (tablePreviewRoot) tablePreviewRoot.innerHTML = ''

    // PDF일 때 markdown은 UI용(표 미리보기)만 사용
    if (ext === 'pdf') {
      try {
        const fd2 = new FormData()
        fd2.append('file', f)
        const rMd = await fetch(`${API_BASE()}/text/markdown`, {
          method: 'POST',
          body: fd2,
        })
        if (rMd.ok) {
          const mdData = await rMd.json()
          let markdown = ''
          if (typeof mdData.markdown === 'string') markdown = mdData.markdown
          else if (Array.isArray(mdData.pages_md))
            markdown = mdData.pages_md.join('\n\n')
          else if (Array.isArray(mdData.pages))
            markdown = mdData.pages.map((p) => p.markdown || '').join('\n\n')

          const tableCount = renderTablePreview(markdown)

          // textarea는 보기용 plain preview
          const ta2 = $('#txt-out')
          if (ta2) {
            if (tableCount > 0) {
              const plainPreview = buildPlainTextFromMarkdown(markdown)
              if (plainPreview.trim()) {
                ta2.classList.remove('hidden')
                ta2.value = plainPreview
              } else {
                ta2.value = analysisText || ''
              }
            } else {
              ta2.classList.remove('hidden')
              ta2.value = analysisText || ''
            }
          }
        }
      } catch (e) {
        console.warn('markdown 추출 중 오류', e)
      }
    }

    // NER/정규식/레닥션 기준 텍스트는 full_text 기준
    const isPdf = ext === 'pdf'
    // 레이아웃(줄바꿈/탭)이 의미가 있는 문서들은 줄붙이기(joinBrokenLines)를 하지 않는다.
    // - pdf: 별도 처리(isPdf)
    // - hwpx/docx/pptx/xlsx: 서버 추출 단계에서 레이아웃을 최대한 살려서 반환
    // - hwp/doc/ppt/xls: 레거시 포맷도 줄바꿈을 유지해야 NER 컨텍스트가 깨지지 않음
    // - docm/xlsm/pptm: 매크로 확장자는 각각 docx/xlsx/pptx 계열로 취급
    const isLayoutPreserved = [
      'hwpx',
      'hwp',
      'docx',
      'doc',
      'pptx',
      'ppt',
      'xlsx',
      'xls',
      'docm',
      'pptm',
    ].includes(ext)
    const normalizedText =
      isPdf || isLayoutPreserved ? analysisText : joinBrokenLines(analysisText)

    setStatus('정규식 매칭 중...')
    const rules = selectedRuleNames()
    const tMatch = performance.now()
    const r2 = await fetch(`${API_BASE()}/text/match`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: normalizedText, rules, normalize: true }),
    })
    if (!r2.ok) throw new Error(`매칭 실패 (${r2.status})\n${await r2.text()}`)
    const rawMatchData = await r2.json()
    timings.match_ms = performance.now() - tMatch
    const matchData = filterMatchByRules(rawMatchData, rules)
    renderRegexResults(matchData)
    setOpen('match', true)

    setStatus('NER 분석 중...')
    const nerLabelsUsed = selectedNerLabels()
    const tNer = performance.now()
    let nerResp = { items: [] }

    if (!normalizedText.trim()) {
      nerResp = { items: [] }
      renderNerTable(nerResp)
      setOpen('ner', true)
      __lastNerEntities = []
    } else {
      const excludeSpans = buildExcludeSpansFromMatch(matchData)
      const ner = await requestNerSmart(
        normalizedText,
        excludeSpans,
        false,
        nerLabelsUsed
      )
      nerResp = ner
      __lastNerEntities = Array.isArray(ner?.items) ? ner.items : []
      renderNerTable(nerResp)
      setOpen('ner', true)
    }
    timings.ner_ms = performance.now() - tNer

    // 통계 컨텍스트 저장(체크박스 변경 시 재계산)
    timings.total_ms = performance.now() - t0
    __lastStatsContext = { file: f, ext, rules, matchData, timings, t0 }
    refreshStatsFromContext()

    setStatus(`스캔 완료 (${ext.toUpperCase()} 처리) — 레닥션 준비 중...`)

    setStatus('레닥션 파일 생성 중...')
    const tRedact = performance.now()
    const fdRedact = new FormData()
    fdRedact.append('file', f)

    const rulesForRedact = selectedRuleNames()
    if (rulesForRedact.length)
      fdRedact.append('rules_json', JSON.stringify(rulesForRedact))

    const nerLabelsForRedact = selectedNerLabels()
    if (nerLabelsForRedact.length)
      fdRedact.append('ner_labels_json', JSON.stringify(nerLabelsForRedact))

    if (Array.isArray(__lastNerEntities) && __lastNerEntities.length) {
      fdRedact.append('ner_entities_json', JSON.stringify(__lastNerEntities))
    }

    const r4 = await fetch(`${API_BASE()}/redact/file`, {
      method: 'POST',
      body: fdRedact,
    })
    if (!r4.ok) throw new Error(`레닥션 실패 (${r4.status})`)
    const blob = await r4.blob()
    timings.redact_ms = performance.now() - tRedact
    timings.total_ms = performance.now() - t0

    // 레닥션 이후 timings 반영해서 리포트 갱신
    __lastStatsContext = { file: f, ext, rules, matchData, timings, t0 }
    refreshStatsFromContext()

    const ctype = r4.headers.get('Content-Type') || 'application/octet-stream'
    __lastRedactedBlob = new Blob([blob], { type: ctype })

    if (ctype.includes('pdf')) {
      setOpen('pdf', true)
      await renderRedactedPdfPreview(__lastRedactedBlob)
    } else {
      setOpen('pdf', false)
    }

    const btn = $('#btn-save-redacted')
    if (btn) {
      btn.classList.remove('hidden')
      btn.disabled = false
    }
    setStatus('레닥션 완료 — 다운로드 가능')
  } catch (e) {
    console.error(e)
    setStatus(`오류: ${e.message || e}`)
  }
})

$('#btn-save-redacted')?.addEventListener('click', () => {
  if (!__lastRedactedBlob) return alert('레닥션된 파일이 없습니다.')
  const url = URL.createObjectURL(__lastRedactedBlob)
  const a = document.createElement('a')
  a.href = url
  a.download = __lastRedactedName || 'redacted_file'
  a.click()
  URL.revokeObjectURL(url)
})

document.addEventListener('DOMContentLoaded', () => {
  loadRules()
  setupDropZone()

  // NER 체크박스 변경 시: 리포트/NER 표 동기 갱신
  ;['#ner-show-ps', '#ner-show-lc', '#ner-show-og'].forEach((sel) => {
    $(sel)?.addEventListener('change', () => {
      // NER 표는 마지막 응답 기준으로 재렌더
      renderNerTable({ items: __lastNerEntities })
      refreshStatsFromContext()
    })
  })

  // 리포트 JSON 토글
  $('#btn-stats-json-toggle')?.addEventListener('click', () => {
    const pre = $('#stats-json')
    if (!pre) return
    const nextHidden = !pre.classList.contains('hidden')
    pre.classList.toggle('hidden', nextHidden)
    if (!nextHidden) {
      pre.textContent = __lastScanStats
        ? JSON.stringify(__lastScanStats, null, 2)
        : ''
    }
  })

  // 리포트 다운로드(JSON)
  $('#btn-stats-download')?.addEventListener('click', () => {
    if (!__lastScanStats)
      return alert('리포트가 없습니다. 먼저 스캔을 실행하세요.')
    const name0 = (__lastScanStats?.document?.name || 'report')
      .replace(/\.[^.]+$/, '')
      .slice(0, 120)
    const outName = `${name0}_risk_report.json`

    const blob = new Blob([JSON.stringify(__lastScanStats, null, 2)], {
      type: 'application/json',
    })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = outName
    a.click()
    URL.revokeObjectURL(url)
  })
})
