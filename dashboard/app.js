/* 엑스와이지 IP 통합 관리 대시보드
   데이터는 ../data/patents.json 한 파일. scripts/import-excel.js 가 엑셀을 이 형식으로 변환한다. */

// loungex-brand-dashboard 와 동일한 팔레트 — 두 대시보드의 차트 색을 맞춘다
const PALETTE = ['#4263eb', '#1f2329', '#7950f2', '#f59f00', '#2f9e44', '#e8590c', '#15aabf', '#e64980', '#5c7cfa', '#868e96'];

const STATUS_LABEL = {
  registered: '등록',
  pending: '출원중',
  rejected: '거절',
  abandoned: '포기',
  expired: '소멸',
};
const COUNTRY_LABEL = { KR: '한국', US: '미국', CN: '중국', JP: '일본', EP: '유럽', WO: 'PCT' };
const AXES = [
  { key: 'technology', label: '기술성' },
  { key: 'marketability', label: '시장성' },
  { key: 'rights', label: '권리성' },
  { key: 'utilization', label: '활용도' },
];
const VIEW_TITLE = {
  overview: 'IP 포트폴리오 개요',
  list: '특허 목록',
  family: '관련 특허군',
  inventors: '발명자 분석',
  evaluation: '기술평가',
};

let db = { meta: {}, patents: [] };
let currentView = 'overview';
const charts = {};
const sortState = {
  list: { key: 'applicationDate', dir: 'desc' },
  inventor: { key: 'total', dir: 'desc' },
  eval: { key: 'score', dir: 'desc' },
};
const filters = { q: '', country: '', status: '', category: '' };

const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));
const escapeHtml = (s) =>
  String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtDate = (d) => (d ? String(d).slice(0, 10) : '-');
const yearOf = (d) => (d ? Number(String(d).slice(0, 4)) : null);

/* ===== 평가 점수 ===== */
// 4개 축 중 입력된 것만 평균. 하나도 없으면 null(미평가)로 두고 등급도 NA.
function scoreOf(p) {
  const vals = AXES.map((a) => p.evaluation?.[a.key]).filter((v) => typeof v === 'number' && v > 0);
  if (!vals.length) return null;
  return vals.reduce((a, b) => a + b, 0) / vals.length;
}
function gradeOf(score) {
  if (score == null) return 'NA';
  if (score >= 4.5) return 'S';
  if (score >= 3.5) return 'A';
  if (score >= 2.5) return 'B';
  return 'C';
}

/* ===== 로드 & 정규화 ===== */
async function load() {
  const res = await fetch('../data/patents.json', { cache: 'no-store' });
  if (!res.ok) throw new Error(`patents.json 로드 실패 (${res.status})`);
  const raw = await res.json();
  db.meta = raw.meta || {};
  db.patents = (raw.patents || []).map(normalize);
  render();
}

function normalize(p, i) {
  const inventors = Array.isArray(p.inventors)
    ? p.inventors.filter(Boolean)
    : String(p.inventors || '').split(/[,;/·]/).map((s) => s.trim()).filter(Boolean);
  const ipc = Array.isArray(p.ipc) ? p.ipc : String(p.ipc || '').split(/[,;]/).map((s) => s.trim()).filter(Boolean);
  const productLine = Array.isArray(p.productLine)
    ? p.productLine
    : String(p.productLine || '').split(/[,;]/).map((s) => s.trim()).filter(Boolean);
  const score = scoreOf(p);
  return {
    ...p,
    id: p.id || `P${String(i + 1).padStart(3, '0')}`,
    title: p.title || '(제목 없음)',
    country: (p.country || 'KR').toUpperCase(),
    status: p.status || (p.registrationNumber ? 'registered' : 'pending'),
    inventors,
    ipc,
    productLine,
    family: p.family || '미분류',
    category: p.category || '미분류',
    inventorText: inventors.join(', '),
    score,
    grade: gradeOf(score),
  };
}

/* ===== 렌더 ===== */
function render() {
  $('#lastUpdated').textContent = db.meta.lastUpdated
    ? `업데이트 ${String(db.meta.lastUpdated).replace('T', ' ').slice(0, 16)}`
    : '업데이트 기록 없음';
  $('#pageEyebrow').textContent = db.meta.company || '주식회사 엑스와이지';

  const families = new Set(db.patents.map((p) => p.family));
  const inventors = new Set(db.patents.flatMap((p) => p.inventors));
  $('#badgeList').textContent = db.patents.length;
  $('#badgeFamily').textContent = families.size;
  $('#badgeInventors').textContent = inventors.size;

  buildFilterOptions();
  renderOverview();
  renderList();
  renderFamily();
  renderInventors();
  renderEvaluation();
  if (window.lucide) lucide.createIcons();
}

function kpi(label, value, sub, icon, accent) {
  return `
    <div class="kpi-card${accent ? ' accent' : ''}">
      <div class="kpi-icon"><i data-lucide="${icon}"></i></div>
      <div class="kpi-label">${escapeHtml(label)}</div>
      <div class="kpi-value">${value}</div>
      <div class="kpi-sub">${escapeHtml(sub)}</div>
    </div>`;
}

function countBy(arr, fn) {
  const m = new Map();
  arr.forEach((x) => {
    const k = fn(x);
    (Array.isArray(k) ? k : [k]).forEach((kk) => {
      if (kk == null || kk === '') return;
      m.set(kk, (m.get(kk) || 0) + 1);
    });
  });
  return [...m.entries()].sort((a, b) => b[1] - a[1]);
}

/* ===== 개요 ===== */
function renderOverview() {
  const ps = db.patents;
  const registered = ps.filter((p) => p.status === 'registered').length;
  const pending = ps.filter((p) => p.status === 'pending').length;
  const countries = new Set(ps.map((p) => p.country));
  const inventors = new Set(ps.flatMap((p) => p.inventors));
  const scored = ps.filter((p) => p.score != null);
  const avg = scored.length ? scored.reduce((a, p) => a + p.score, 0) / scored.length : null;

  $('#kpiRow').innerHTML = [
    kpi('전체 특허', ps.length, `등록 ${registered} · 출원중 ${pending}`, 'file-text', true),
    kpi('등록 특허', registered, `${countries.size}개국 · ${[...countries].join(' / ')}`, 'badge-check'),
    kpi('발명자', inventors.size, '공동발명 포함', 'users'),
    kpi(
      '평균 평가점수',
      avg == null ? '-' : `${avg.toFixed(2)}<small>/5</small>`,
      `평가 완료 ${scored.length} / ${ps.length}건`,
      'gauge'
    ),
  ].join('');

  drawDoughnut('countryChart', countBy(ps, (p) => p.country).map(([k, v]) => [COUNTRY_LABEL[k] || k, v]));
  drawDoughnut('categoryChart', countBy(ps, (p) => p.category));

  // 연도별 출원/등록 — 두 계열의 연도 축을 합쳐야 막대가 어긋나지 않는다
  const appYears = countBy(ps, (p) => yearOf(p.applicationDate));
  const regYears = countBy(ps, (p) => yearOf(p.registrationDate));
  const years = [...new Set([...appYears, ...regYears].map(([y]) => y))].filter(Boolean).sort();
  const appMap = new Map(appYears);
  const regMap = new Map(regYears);
  drawChart('yearlyChart', {
    type: 'bar',
    data: {
      labels: years,
      datasets: [
        { label: '출원', data: years.map((y) => appMap.get(y) || 0), backgroundColor: PALETTE[0], borderRadius: 4 },
        { label: '등록', data: years.map((y) => regMap.get(y) || 0), backgroundColor: PALETTE[4], borderRadius: 4 },
      ],
    },
    options: {
      ...baseOptions(),
      scales: { x: gridX(), y: { ...gridY(), ticks: { ...tickStyle(), precision: 0 } } },
    },
  });

  const top = [...ps].filter((p) => p.score != null).sort((a, b) => b.score - a.score).slice(0, 5);
  $('#topPatentsBody').innerHTML = top.length
    ? top
        .map(
          (p) => `<tr>
            <td><span class="grade ${p.grade}">${p.grade}</span></td>
            <td><span class="patent-link" data-id="${p.id}">${escapeHtml(p.title)}</span></td>
            <td><span class="tag country">${escapeHtml(p.country)}</span></td>
            <td><span class="tag ${p.status}">${STATUS_LABEL[p.status] || p.status}</span></td>
            <td>${scoreBar(p.score)}</td>
          </tr>`
        )
        .join('')
    : `<tr><td colspan="5" class="empty-msg">평가된 특허가 없습니다.</td></tr>`;
}

function scoreBar(score) {
  if (score == null) return '<span class="dim">-</span>';
  return `<div class="score">
    <div class="score-track"><div class="score-fill" style="width:${(score / 5) * 100}%"></div></div>
    <span class="score-value">${score.toFixed(1)}</span>
  </div>`;
}

/* ===== 특허 목록 ===== */
function buildFilterOptions() {
  const fill = (sel, entries, labelFn = (k) => k) => {
    const el = $(sel);
    const keep = el.querySelector('option').outerHTML;
    el.innerHTML = keep + entries.map(([k, v]) => `<option value="${escapeHtml(k)}">${escapeHtml(labelFn(k))} (${v})</option>`).join('');
  };
  fill('#filterCountry', countBy(db.patents, (p) => p.country), (k) => COUNTRY_LABEL[k] || k);
  fill('#filterStatus', countBy(db.patents, (p) => p.status), (k) => STATUS_LABEL[k] || k);
  fill('#filterCategory', countBy(db.patents, (p) => p.category));
}

function filtered() {
  const q = filters.q.trim().toLowerCase();
  return db.patents.filter((p) => {
    if (filters.country && p.country !== filters.country) return false;
    if (filters.status && p.status !== filters.status) return false;
    if (filters.category && p.category !== filters.category) return false;
    if (!q) return true;
    const hay = [p.title, p.applicationNumber, p.registrationNumber, p.inventorText, p.family, p.category, p.abstract, ...p.ipc]
      .join(' ')
      .toLowerCase();
    return hay.includes(q);
  });
}

function sortRows(rows, state) {
  const { key, dir } = state;
  const mul = dir === 'asc' ? 1 : -1;
  const GRADE_ORDER = { S: 4, A: 3, B: 2, C: 1, NA: 0 };
  return [...rows].sort((a, b) => {
    let x = a[key];
    let y = b[key];
    if (key === 'grade') { x = GRADE_ORDER[x]; y = GRADE_ORDER[y]; }
    // null 은 정렬 방향과 무관하게 항상 뒤로 — 미평가 건이 상위를 차지하면 표가 쓸모없어진다
    if (x == null && y == null) return 0;
    if (x == null) return 1;
    if (y == null) return -1;
    if (typeof x === 'number' && typeof y === 'number') return (x - y) * mul;
    return String(x).localeCompare(String(y), 'ko') * mul;
  });
}

function renderList() {
  const rows = sortRows(filtered(), sortState.list);
  $('#listHint').textContent = `${rows.length}건 표시 (전체 ${db.patents.length}건) · 컬럼 클릭 시 정렬`;
  $('#patentTable tbody').innerHTML = rows.length
    ? rows
        .map(
          (p) => `<tr>
            <td><span class="grade ${p.grade}">${p.grade === 'NA' ? '–' : p.grade}</span></td>
            <td><span class="patent-link" data-id="${p.id}">${escapeHtml(p.title)}</span></td>
            <td><span class="tag country">${escapeHtml(p.country)}</span></td>
            <td><span class="tag ${p.status}">${STATUS_LABEL[p.status] || p.status}</span></td>
            <td class="num dim">${escapeHtml(p.applicationNumber || '-')}</td>
            <td class="num">${fmtDate(p.applicationDate)}</td>
            <td class="num">${fmtDate(p.registrationDate)}</td>
            <td>${escapeHtml(p.inventorText || '-')}</td>
            <td>${scoreBar(p.score)}</td>
          </tr>`
        )
        .join('')
    : `<tr><td colspan="9" class="empty-msg">조건에 맞는 특허가 없습니다.</td></tr>`;
  markSorted('#patentTable', sortState.list);
}

function markSorted(tableSel, state) {
  $$(`${tableSel} thead th`).forEach((th) => {
    const on = th.dataset.sort === state.key;
    th.classList.toggle('sorted', on);
    th.textContent = th.textContent.replace(/ [▲▼]$/, '') + (on ? (state.dir === 'asc' ? ' ▲' : ' ▼') : '');
  });
}

/* ===== 관련 특허군 ===== */
function renderFamily() {
  const groups = new Map();
  db.patents.forEach((p) => {
    if (!groups.has(p.family)) groups.set(p.family, []);
    groups.get(p.family).push(p);
  });
  const sorted = [...groups.entries()].sort((a, b) => b[1].length - a[1].length);

  drawChart('familyChart', {
    type: 'bar',
    data: {
      labels: sorted.map(([k]) => k),
      datasets: [{ label: '보유 건수', data: sorted.map(([, v]) => v.length), backgroundColor: PALETTE[0], borderRadius: 4, maxBarThickness: 30 }],
    },
    options: {
      ...baseOptions(),
      indexAxis: 'y',
      plugins: { ...baseOptions().plugins, legend: { display: false } },
      scales: { x: { ...gridX(), ticks: { ...tickStyle(), precision: 0 } }, y: gridY(false) },
    },
  });

  $('#familyBody').innerHTML = sorted.length
    ? sorted
        .map(
          ([name, items]) => `
        <div class="family-group">
          <div class="family-name">
            <span class="family-dot"></span>${escapeHtml(name)}
            <span class="family-count">${items.length}건</span>
          </div>
          <div class="family-items">
            ${items
              .map(
                (p) => `<div class="family-item" data-id="${p.id}">
                  <div class="family-item-title">${escapeHtml(p.title)}</div>
                  <div class="family-item-meta">
                    <span class="tag country">${escapeHtml(p.country)}</span>
                    <span class="tag ${p.status}">${STATUS_LABEL[p.status] || p.status}</span>
                    <span class="grade ${p.grade}">${p.grade === 'NA' ? '–' : p.grade}</span>
                    <span class="dim num">${fmtDate(p.applicationDate)}</span>
                  </div>
                </div>`
              )
              .join('')}
          </div>
        </div>`
        )
        .join('')
    : `<div class="empty-msg">등록된 특허가 없습니다.</div>`;
}

/* ===== 발명자 ===== */
function inventorStats() {
  const m = new Map();
  db.patents.forEach((p) => {
    p.inventors.forEach((name) => {
      if (!m.has(name)) m.set(name, { name, total: 0, registered: 0, pending: 0, scores: [], categories: new Map() });
      const s = m.get(name);
      s.total++;
      if (p.status === 'registered') s.registered++;
      if (p.status === 'pending') s.pending++;
      if (p.score != null) s.scores.push(p.score);
      s.categories.set(p.category, (s.categories.get(p.category) || 0) + 1);
    });
  });
  return [...m.values()].map((s) => ({
    ...s,
    avgScore: s.scores.length ? s.scores.reduce((a, b) => a + b, 0) / s.scores.length : null,
    categoryText: [...s.categories.entries()].sort((a, b) => b[1] - a[1]).slice(0, 2).map(([k]) => k).join(', '),
  }));
}

function renderInventors() {
  const stats = inventorStats();
  const soleCount = db.patents.filter((p) => p.inventors.length === 1).length;
  const coCount = db.patents.filter((p) => p.inventors.length > 1).length;
  const avgPer = db.patents.length
    ? db.patents.reduce((a, p) => a + p.inventors.length, 0) / db.patents.length
    : 0;

  $('#inventorKpiRow').innerHTML = [
    kpi('총 발명자', stats.length, '중복 제거 기준', 'users', true),
    kpi('공동발명 건', coCount, `단독발명 ${soleCount}건`, 'user-plus'),
    kpi('건당 평균 발명자', avgPer.toFixed(1), '명', 'user-round'),
    kpi(
      '최다 참여',
      stats.length ? escapeHtml([...stats].sort((a, b) => b.total - a.total)[0].name) : '-',
      stats.length ? `${[...stats].sort((a, b) => b.total - a.total)[0].total}건 참여` : '',
      'crown'
    ),
  ].join('');

  const top = [...stats].sort((a, b) => b.total - a.total).slice(0, 15);
  drawChart('inventorChart', {
    type: 'bar',
    data: {
      labels: top.map((s) => s.name),
      // maxBarThickness — 발명자가 2~3명뿐일 때 막대가 캔버스 높이를 나눠 가져 우스꽝스럽게 두꺼워진다
      datasets: [
        { label: '등록', data: top.map((s) => s.registered), backgroundColor: PALETTE[0], borderRadius: 4, stack: 's', maxBarThickness: 30 },
        { label: '출원중', data: top.map((s) => s.pending), backgroundColor: PALETTE[3], borderRadius: 4, stack: 's', maxBarThickness: 30 },
        {
          label: '기타',
          data: top.map((s) => s.total - s.registered - s.pending),
          backgroundColor: PALETTE[9],
          borderRadius: 4,
          stack: 's',
          maxBarThickness: 30,
        },
      ],
    },
    options: {
      ...baseOptions(),
      indexAxis: 'y',
      scales: {
        x: { ...gridX(), stacked: true, ticks: { ...tickStyle(), precision: 0 } },
        y: { ...gridY(false), stacked: true },
      },
    },
  });

  const rows = sortRows(stats, sortState.inventor);
  $('#inventorTable tbody').innerHTML = rows.length
    ? rows
        .map(
          (s) => `<tr>
            <td><b style="color:var(--text)">${escapeHtml(s.name)}</b></td>
            <td class="num">${s.total}</td>
            <td class="num">${s.registered}</td>
            <td class="num">${s.pending}</td>
            <td>${scoreBar(s.avgScore)}</td>
            <td class="dim">${escapeHtml(s.categoryText || '-')}</td>
          </tr>`
        )
        .join('')
    : `<tr><td colspan="6" class="empty-msg">발명자 정보가 없습니다.</td></tr>`;
  markSorted('#inventorTable', sortState.inventor);
}

/* ===== 기술평가 ===== */
function renderEvaluation() {
  const scored = db.patents.filter((p) => p.score != null);
  const byGrade = (g) => db.patents.filter((p) => p.grade === g).length;
  const axisAvg = AXES.map((a) => {
    const vals = db.patents.map((p) => p.evaluation?.[a.key]).filter((v) => typeof v === 'number' && v > 0);
    return vals.length ? vals.reduce((x, y) => x + y, 0) / vals.length : 0;
  });

  $('#evalKpiRow').innerHTML = [
    kpi('평가 완료', `${scored.length}<small>/${db.patents.length}</small>`, `미평가 ${db.patents.length - scored.length}건`, 'clipboard-check', true),
    kpi('S등급', byGrade('S'), '종합 4.5점 이상 · 핵심 특허', 'star'),
    kpi('A등급', byGrade('A'), '종합 3.5점 이상', 'thumbs-up'),
    kpi('B·C등급', byGrade('B') + byGrade('C'), '보완 또는 정리 검토 대상', 'alert-circle'),
  ].join('');

  drawChart('radarChart', {
    type: 'radar',
    data: {
      labels: AXES.map((a) => a.label),
      datasets: [
        {
          label: '전체 평균',
          data: axisAvg,
          backgroundColor: 'rgba(66, 99, 235, 0.15)',
          borderColor: PALETTE[0],
          borderWidth: 2,
          pointBackgroundColor: PALETTE[0],
          pointRadius: 4,
        },
      ],
    },
    options: {
      ...baseOptions(),
      plugins: { ...baseOptions().plugins, legend: { display: false } },
      scales: {
        r: {
          min: 0,
          max: 5,
          ticks: { stepSize: 1, backdropColor: 'transparent', color: '#9a9fa8', font: { size: 10 } },
          grid: { color: '#ececf1' },
          angleLines: { color: '#ececf1' },
          pointLabels: { color: '#495057', font: { size: 12, weight: '600' } },
        },
      },
    },
  });

  // 등급별 색은 팔레트 순서가 아니라 등급에 고정 — S/A 를 같은 파란 계열로 두면 도넛에서 구분이 안 된다
  const GRADE_COLOR = { S: '#4263eb', A: '#7950f2', B: '#868e96', C: '#e8590c' };
  const gradeEntries = ['S', 'A', 'B', 'C'].map((g) => [g, byGrade(g)]).filter(([, v]) => v > 0);
  drawDoughnut(
    'gradeChart',
    gradeEntries.map(([g, v]) => [`${g}등급`, v]),
    gradeEntries.map(([g]) => GRADE_COLOR[g])
  );

  drawChart('matrixChart', {
    type: 'scatter',
    data: {
      datasets: [
        {
          label: '특허',
          // 같은 좌표에 여러 건이 겹치면 한 점으로 보인다 — 툴팁에 명칭을 실어 구분
          data: scored.map((p) => ({
            x: p.evaluation?.technology ?? 0,
            y: p.evaluation?.marketability ?? 0,
            id: p.id,
            title: p.title,
          })),
          backgroundColor: 'rgba(66, 99, 235, 0.65)',
          borderColor: PALETTE[0],
          pointRadius: 7,
          pointHoverRadius: 10,
        },
      ],
    },
    options: {
      ...baseOptions(),
      plugins: {
        ...baseOptions().plugins,
        legend: { display: false },
        tooltip: {
          ...baseOptions().plugins.tooltip,
          callbacks: {
            label: (ctx) => [`${ctx.raw.title}`, `기술성 ${ctx.raw.x} · 시장성 ${ctx.raw.y}`],
          },
        },
      },
      onClick: (evt, els) => {
        if (els.length) openDetail(charts.matrixChart.data.datasets[0].data[els[0].index].id);
      },
      // 점수는 1~5 뿐이라 0부터 그리면 캔버스 아래·왼쪽 절반이 늘 빈 채로 남는다.
      // 다만 min 을 0.5 로 두면 눈금이 0.5/1.5/... 로 찍혀서 정수로 고정한다
      scales: {
        x: { ...gridX(), min: 0.5, max: 5.5, title: { display: true, text: '기술성 →', color: '#9a9fa8', font: { size: 11 } }, ticks: tickStyle(), afterBuildTicks: (ax) => { ax.ticks = [1, 2, 3, 4, 5].map((v) => ({ value: v })); } },
        y: { ...gridY(), min: 0.5, max: 5.5, title: { display: true, text: '시장성 →', color: '#9a9fa8', font: { size: 11 } }, ticks: tickStyle(), afterBuildTicks: (ax) => { ax.ticks = [1, 2, 3, 4, 5].map((v) => ({ value: v })); } },
      },
    },
    plugins: [quadrantPlugin],
  });

  const rows = sortRows(
    db.patents.map((p) => ({
      ...p,
      technology: p.evaluation?.technology ?? null,
      marketability: p.evaluation?.marketability ?? null,
      rights: p.evaluation?.rights ?? null,
      utilization: p.evaluation?.utilization ?? null,
    })),
    sortState.eval
  );
  const cell = (v) => (v == null ? '<span class="dim">-</span>' : `<span class="num">${v}</span>`);
  $('#evalTable tbody').innerHTML = rows.length
    ? rows
        .map(
          (p) => `<tr>
            <td><span class="grade ${p.grade}">${p.grade === 'NA' ? '–' : p.grade}</span></td>
            <td><span class="patent-link" data-id="${p.id}">${escapeHtml(p.title)}</span></td>
            <td>${cell(p.technology)}</td>
            <td>${cell(p.marketability)}</td>
            <td>${cell(p.rights)}</td>
            <td>${cell(p.utilization)}</td>
            <td>${scoreBar(p.score)}</td>
          </tr>`
        )
        .join('')
    : `<tr><td colspan="7" class="empty-msg">등록된 특허가 없습니다.</td></tr>`;
  markSorted('#evalTable', sortState.eval);
}

/* ===== 상세 모달 ===== */
function openDetail(id) {
  const p = db.patents.find((x) => x.id === id);
  if (!p) return;
  $('#modalTitle').textContent = p.title;
  $('#modalSub').innerHTML = [
    `<span class="tag country">${escapeHtml(COUNTRY_LABEL[p.country] || p.country)}</span>`,
    `<span class="tag ${p.status}">${STATUS_LABEL[p.status] || p.status}</span>`,
    `<span class="tag">${escapeHtml(p.category)}</span>`,
    p.score != null ? `<span class="grade ${p.grade}">${p.grade}</span>` : '',
  ].join('');

  // 같은 특허군 = 관련 특허. 자기 자신은 제외
  const related = db.patents.filter((x) => x.family === p.family && x.id !== p.id);

  const row = (k, v) => `<div class="detail-key">${k}</div><div class="detail-val">${v}</div>`;
  $('#modalBody').innerHTML = `
    <div class="detail-section">
      <h4 class="detail-title">서지 정보</h4>
      <div class="detail-grid">
        ${row('출원번호', escapeHtml(p.applicationNumber || '-'))}
        ${row('출원일', fmtDate(p.applicationDate))}
        ${row('등록번호', escapeHtml(p.registrationNumber || '-'))}
        ${row('등록일', fmtDate(p.registrationDate))}
        ${row('출원인', escapeHtml(p.applicant || '-'))}
        ${row('발명자', escapeHtml(p.inventorText || '-'))}
        ${row('IPC', p.ipc.length ? `<div class="tag-row">${p.ipc.map((c) => `<span class="tag">${escapeHtml(c)}</span>`).join('')}</div>` : '-')}
        ${row('적용 제품', p.productLine.length ? `<div class="tag-row">${p.productLine.map((c) => `<span class="tag">${escapeHtml(c)}</span>`).join('')}</div>` : '-')}
        ${row('특허군', escapeHtml(p.family))}
      </div>
    </div>

    ${p.abstract ? `<div class="detail-section">
      <h4 class="detail-title">요약</h4>
      <div class="detail-text">${escapeHtml(p.abstract)}</div>
    </div>` : ''}

    <div class="detail-section">
      <h4 class="detail-title">기술평가</h4>
      ${p.score == null
        ? '<div class="detail-text dim">아직 평가되지 않았습니다.</div>'
        : `<div class="eval-bars">
            ${AXES.map((a) => {
              const v = p.evaluation?.[a.key];
              return `<div class="eval-bar">
                <span class="eval-bar-label">${a.label}</span>
                <div class="score-track"><div class="score-fill" style="width:${((v || 0) / 5) * 100}%"></div></div>
                <span class="score-value">${v ?? '-'}</span>
              </div>`;
            }).join('')}
            <div class="eval-bar" style="margin-top:4px;padding-top:10px;border-top:1px solid var(--border)">
              <span class="eval-bar-label"><b>종합</b></span>
              <div class="score-track"><div class="score-fill" style="width:${(p.score / 5) * 100}%"></div></div>
              <span class="score-value"><b>${p.score.toFixed(1)}</b></span>
            </div>
          </div>
          ${p.evaluation?.comment ? `<div class="eval-comment">${escapeHtml(p.evaluation.comment)}</div>` : ''}`}
    </div>

    <div class="detail-section">
      <h4 class="detail-title">관련 특허 (${related.length})</h4>
      ${related.length
        ? `<div class="related-list">${related
            .map(
              (r) => `<div class="related-item" data-id="${r.id}">
                <span class="tag country">${escapeHtml(r.country)}</span>
                <span class="related-item-title">${escapeHtml(r.title)}</span>
                <span class="tag ${r.status}">${STATUS_LABEL[r.status] || r.status}</span>
              </div>`
            )
            .join('')}</div>`
        : '<div class="detail-text dim">같은 특허군에 다른 특허가 없습니다.</div>'}
    </div>`;

  $('#modal').hidden = false;
  if (window.lucide) lucide.createIcons();
}

/* ===== 차트 헬퍼 ===== */

// 매트릭스를 3점 기준 4분면으로 나누고 우상단(고기술·고시장)을 옅게 칠한다.
// 이게 없으면 산점도가 그냥 점 무더기라 '우상단이 핵심'이라는 해석 기준이 화면에 안 보인다.
const quadrantPlugin = {
  id: 'quadrant',
  beforeDatasetsDraw(chart) {
    const { ctx, chartArea: a, scales: { x, y } } = chart;
    if (!a) return;
    const cx = x.getPixelForValue(3);
    const cy = y.getPixelForValue(3);
    ctx.save();
    ctx.fillStyle = 'rgba(66, 99, 235, 0.05)';
    ctx.fillRect(cx, a.top, a.right - cx, cy - a.top);
    ctx.strokeStyle = '#e0e1e8';
    ctx.setLineDash([4, 4]);
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(cx, a.top); ctx.lineTo(cx, a.bottom);
    ctx.moveTo(a.left, cy); ctx.lineTo(a.right, cy);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = '#c3c6cf';
    ctx.font = '600 11px Pretendard, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText('핵심 특허군', a.right - 10, a.top + 20);
    ctx.restore();
  },
};

function tickStyle() { return { color: '#9a9fa8', font: { size: 11 } }; }
function gridX() { return { grid: { display: false }, ticks: tickStyle() }; }
function gridY(grid = true) { return { grid: { color: '#ececf1', display: grid, drawBorder: false }, ticks: tickStyle() }; }
function baseOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'bottom', labels: { color: '#495057', usePointStyle: true, pointStyle: 'circle', boxWidth: 8, padding: 16, font: { size: 12 } } },
      tooltip: {
        backgroundColor: '#1f2329',
        padding: 12,
        cornerRadius: 8,
        titleFont: { size: 12 },
        bodyFont: { size: 12 },
        displayColors: false,
      },
    },
  };
}

function drawChart(id, config) {
  const el = document.getElementById(id);
  if (!el) return;
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(el, config);
}

function drawDoughnut(id, entries, colors) {
  drawChart(id, {
    type: 'doughnut',
    data: {
      labels: entries.map(([k]) => k),
      datasets: [
        {
          data: entries.map(([, v]) => v),
          backgroundColor: colors || PALETTE,
          borderWidth: 2,
          borderColor: '#fff',
        },
      ],
    },
    options: {
      ...baseOptions(),
      cutout: '58%',
      plugins: {
        ...baseOptions().plugins,
        tooltip: {
          ...baseOptions().plugins.tooltip,
          callbacks: {
            label: (ctx) => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total ? ((ctx.parsed / total) * 100).toFixed(1) : '0.0';
              return `${ctx.label}: ${ctx.parsed}건 (${pct}%)`;
            },
          },
        },
      },
    },
  });
}

/* ===== 뷰 전환 & 이벤트 ===== */
function switchView(view) {
  currentView = view;
  ['overview', 'list', 'family', 'inventors', 'evaluation'].forEach((v) => {
    $(`#${v}View`).hidden = v !== view;
  });
  $$('.nav-item').forEach((el) => el.classList.toggle('active', el.dataset.view === view));
  $('#pageTitle').textContent = VIEW_TITLE[view];
  // 숨겨진 canvas 는 크기가 0으로 잡혀 차트가 찌그러진다 — 표시 후 리사이즈
  Object.values(charts).forEach((c) => c.resize());
}

function bind() {
  $$('.nav-item').forEach((el) =>
    el.addEventListener('click', (e) => {
      e.preventDefault();
      switchView(el.dataset.view);
    })
  );

  // 상세 열기 — 목록/카드/모달 내 관련특허 모두 위임으로 처리
  document.addEventListener('click', (e) => {
    const t = e.target.closest('[data-id]');
    if (t) openDetail(t.dataset.id);
  });

  $$('[data-close]').forEach((el) => el.addEventListener('click', () => ($('#modal').hidden = true)));
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') $('#modal').hidden = true;
  });

  const bindSort = (tableSel, stateKey, rerender) => {
    $$(`${tableSel} thead th`).forEach((th) =>
      th.addEventListener('click', () => {
        const key = th.dataset.sort;
        if (!key) return;
        const s = sortState[stateKey];
        if (s.key === key) s.dir = s.dir === 'asc' ? 'desc' : 'asc';
        else { s.key = key; s.dir = 'desc'; }
        rerender();
      })
    );
  };
  bindSort('#patentTable', 'list', renderList);
  bindSort('#inventorTable', 'inventor', renderInventors);
  bindSort('#evalTable', 'eval', renderEvaluation);

  $('#searchInput').addEventListener('input', (e) => { filters.q = e.target.value; renderList(); });
  $('#filterCountry').addEventListener('change', (e) => { filters.country = e.target.value; renderList(); });
  $('#filterStatus').addEventListener('change', (e) => { filters.status = e.target.value; renderList(); });
  $('#filterCategory').addEventListener('change', (e) => { filters.category = e.target.value; renderList(); });
}

bind();
load().catch((err) => {
  console.error(err);
  const banner = document.createElement('div');
  banner.className = 'error-banner';
  banner.innerHTML = `데이터를 불러오지 못했습니다. <code>${escapeHtml(err.message)}</code><br />
    로컬에서 열었다면 <code>npm run dev</code> 로 정적 서버를 띄워 주세요 (file:// 은 fetch 가 막힙니다).`;
  document.querySelector('main').prepend(banner);
});
