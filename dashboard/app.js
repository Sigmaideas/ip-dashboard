/* 엑스와이지 IP 통합 관리 대시보드
   데이터는 data/rights.json 한 파일. scripts/import-excel.js 가 관리대장 엑셀을 이 형식으로 변환한다.
   특허 평가는 data/evaluations.json 에서 패밀리 단위로 붙는다. */

// loungex-brand-dashboard 와 동일한 팔레트 — 두 대시보드의 차트 색을 맞춘다
const PALETTE = ['#4263eb', '#1f2329', '#7950f2', '#f59f00', '#2f9e44', '#e8590c', '#15aabf', '#e64980', '#5c7cfa', '#868e96'];

const TYPE_LABEL = { patent: '특허', trademark: '상표', design: '디자인', copyright: '저작권', other: '기타' };
const TYPE_COLOR = { patent: '#4263eb', trademark: '#7950f2', design: '#f59f00', copyright: '#15aabf', other: '#868e96' };
const STATUS_LABEL = {
  registered: '등록', pending: '출원', examining: '심사중', pct: 'PCT',
  rejected: '거절', abandoned: '포기', transferred: '이관',
};
// 상태를 살아있는 권리 / 진행중 / 종료 세 갈래로 묶는다. KPI 와 상태 차트가 이 구분을 쓴다
const STATUS_GROUP = {
  registered: 'live', pending: 'progress', examining: 'progress', pct: 'progress',
  rejected: 'closed', abandoned: 'closed', transferred: 'closed',
};
const COUNTRY_LABEL = { KR: '한국', US: '미국', CN: '중국', JP: '일본', EP: '유럽', PCT: 'PCT', SG: '싱가포르', IN: '인도' };
const AXES = [
  { key: 'technology', label: '기술성' },
  { key: 'marketability', label: '시장성' },
  { key: 'rights', label: '권리성' },
  { key: 'utilization', label: '활용도' },
];
const GRADE_COLOR = { S: '#4263eb', A: '#7950f2', B: '#868e96', C: '#e8590c' };

/* 명칭에서 뽑는 도메인 개념어. 형태소 분석 없이 명칭을 토큰으로 쪼개면
   '이송로봇'/'이송 로봇', '로봇팔'/'로봇 팔' 이 서로 다른 낱말로 갈리고
   1건짜리 꼬리만 잔뜩 남아 분석이 되지 않는다. 개념 단위로 직접 묶는다.
   매칭은 공백을 지운 명칭에 대고 하므로 띄어쓰기 차이는 흡수된다.
   새 특허가 어느 키워드에도 안 걸리면 KPI 의 '미분류' 로 드러난다 — 그때 여기에 추가할 것 */
const KEYWORDS = [
  ['자율주행', /자율주행|자율이동|무인주행|자율/],
  ['이송·물류', /이송|물류|배송|전달|운반|수거/],
  ['로봇팔·핸드', /로봇팔|다관절|그리퍼|디스펜서|핸드/],
  ['무인매장·결제', /무인상점|무인결제|무인|결제|주문|매장/],
  ['식음료·조리', /커피|드립|아이스크림|식음료|음식|제조|인쇄/],
  ['의료·방역', /병원|수술|혈액|진단|열화상|방역/],
  ['안전·충돌방지', /충돌|안전|위험도|교통약자/],
  ['영상·인식', /영상|이미지|인식|판단/],
  ['플랫폼·모듈화', /플랫폼|모듈화|시스템용/],
  ['순찰·청소', /순찰|청소|해변|옥내외/],
];
// 히트맵용 순차 램프(밝음→어두움). 각 단계가 글자 대비 4.5:1 을 넘도록 고른 값이라
// 중간을 임의로 끼워 넣으면 셀 위 숫자가 읽히지 않는다
const HEAT_RAMP = ['#d7dffb', '#b6c4f8', '#8fa2f3', '#4263eb', '#26409e'];
const HEAT_INK = ['#1f2329', '#1f2329', '#1f2329', '#ffffff', '#ffffff'];
const VIEW_TITLE = {
  overview: 'IP 포트폴리오 개요',
  list: '권리 목록',
  family: '특허 패밀리',
  keyword: '키워드',
  inventors: '발명자 분석',
  evaluation: '기술평가',
};

let db = { meta: {}, rights: [] };
const charts = {};
const sortState = {
  list: { key: 'applicationDate', dir: 'desc' },
  inventor: { key: 'total', dir: 'desc' },
  eval: { key: 'score', dir: 'desc' },
};
const filters = { q: '', type: '', country: '', status: '' };

const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));
const escapeHtml = (s) =>
  String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtDate = (d) => (d ? String(d).slice(0, 10) : '-');
const yearOf = (d) => (d ? Number(String(d).slice(0, 4)) : null);
const patents = () => db.rights.filter((r) => r.type === 'patent');

/* ===== 평가 점수 ===== */
function scoreOf(ev) {
  const vals = AXES.map((a) => ev?.[a.key]).filter((v) => typeof v === 'number' && v > 0);
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

/* ===== 로드 ===== */
async function load() {
  const res = await fetch('data/rights.json', { cache: 'no-store' });
  if (!res.ok) throw new Error(`rights.json 로드 실패 (${res.status})`);
  const raw = await res.json();
  db.meta = raw.meta || {};
  db.rights = (raw.rights || []).map((r) => {
    const score = scoreOf(r.evaluation);
    return {
      ...r,
      inventors: r.inventors || [],
      inventorText: (r.inventors || []).join(', '),
      score,
      grade: gradeOf(score),
      statusGroup: STATUS_GROUP[r.status] || 'progress',
    };
  });
  render();
}

/* ===== 공통 ===== */
function kpi(label, value, sub, icon, accent) {
  return `<div class="kpi-card${accent ? ' accent' : ''}">
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
    if (k == null || k === '') return;
    m.set(k, (m.get(k) || 0) + 1);
  });
  return [...m.entries()].sort((a, b) => b[1] - a[1]);
}
function scoreBar(score) {
  if (score == null) return '<span class="dim">-</span>';
  return `<div class="score">
    <div class="score-track"><div class="score-fill" style="width:${(score / 5) * 100}%"></div></div>
    <span class="score-value">${score.toFixed(1)}</span>
  </div>`;
}
const typeTag = (t) => `<span class="tag type-${t}">${TYPE_LABEL[t] || t}</span>`;
const gradeTag = (g) => `<span class="grade ${g}">${g === 'NA' ? '–' : g}</span>`;

function render() {
  $('#lastUpdated').textContent = db.meta.lastUpdated
    ? `업데이트 ${String(db.meta.lastUpdated).replace('T', ' ').slice(0, 16)}`
    : '업데이트 기록 없음';
  $('#pageEyebrow').textContent = db.meta.company || '주식회사 엑스와이지';

  $('#badgeList').textContent = db.rights.length;
  $('#badgeKeyword').textContent = keywordStats().rows.length;
  $('#badgeFamily').textContent = new Set(patents().map((p) => p.family)).size;
  $('#badgeInventors').textContent = new Set(patents().flatMap((p) => p.inventors)).size;

  buildFilters();
  renderOverview();
  renderList();
  renderFamily();
  renderKeywords();
  renderInventors();
  renderEvaluation();
  if (window.lucide) lucide.createIcons();
}

/* ===== 개요 ===== */
function renderOverview() {
  const all = db.rights;
  const live = all.filter((r) => r.statusGroup === 'live');
  const pats = patents();
  const fams = new Set(pats.map((p) => p.family));
  const tms = all.filter((r) => r.type === 'trademark');
  const countries = new Set(all.map((r) => r.country));

  $('#kpiRow').innerHTML = [
    kpi('전체 IP', all.length, `등록 ${live.length}건 · ${countries.size}개국`, 'library', true),
    kpi('특허', pats.length, `${fams.size}개 발명 · 등록 ${pats.filter((p) => p.status === 'registered').length}`, 'file-text'),
    kpi('상표', tms.length, `등록 ${tms.filter((t) => t.status === 'registered').length}건`, 'badge-check'),
    kpi('디자인 · 저작권', all.filter((r) => r.type === 'design' || r.type === 'copyright').length, '디자인 + 저작권 합계', 'palette'),
  ].join('');

  drawDoughnut('typeChart',
    countBy(all, (r) => r.type).map(([k, v]) => [TYPE_LABEL[k] || k, v]),
    countBy(all, (r) => r.type).map(([k]) => TYPE_COLOR[k]));

  drawDoughnut('countryChart', countBy(all, (r) => r.country).map(([k, v]) => [COUNTRY_LABEL[k] || k, v]));

  // 연도별 출원 — 권리유형별로 쌓아 어느 해에 무엇을 집중 출원했는지 보이게 한다
  const years = [...new Set(all.map((r) => yearOf(r.applicationDate)).filter(Boolean))].sort();
  const types = ['patent', 'trademark', 'design', 'copyright'];
  drawChart('yearlyChart', {
    type: 'bar',
    data: {
      labels: years,
      datasets: types.map((t) => ({
        label: TYPE_LABEL[t],
        data: years.map((y) => all.filter((r) => r.type === t && yearOf(r.applicationDate) === y).length),
        backgroundColor: TYPE_COLOR[t],
        borderRadius: 3,
        stack: 'y',
      })),
    },
    options: {
      ...baseOptions(),
      scales: { x: { ...gridX(), stacked: true }, y: { ...gridY(), stacked: true, ticks: { ...tickStyle(), precision: 0 } } },
    },
  });

  const groups = [
    { key: 'live', label: '등록', color: PALETTE[4] },
    { key: 'progress', label: '진행중', color: PALETTE[3] },
    { key: 'closed', label: '종료(거절·포기·이관)', color: PALETTE[9] },
  ];
  drawChart('statusChart', {
    type: 'bar',
    data: {
      labels: types.map((t) => TYPE_LABEL[t]),
      datasets: groups.map((g) => ({
        label: g.label,
        data: types.map((t) => all.filter((r) => r.type === t && r.statusGroup === g.key).length),
        backgroundColor: g.color,
        borderRadius: 3,
        stack: 's',
        maxBarThickness: 40,
      })),
    },
    options: {
      ...baseOptions(),
      indexAxis: 'y',
      scales: { x: { ...gridX(), stacked: true, ticks: { ...tickStyle(), precision: 0 } }, y: { ...gridY(false), stacked: true } },
    },
  });

  // 개요의 TOP 5 는 개별 건이 아니라 발명(패밀리) 단위 — 같은 발명이 4번 나오면 표가 무의미해진다
  const top = familyStats().filter((f) => f.score != null).sort((a, b) => b.score - a.score).slice(0, 5);
  $('#topPatentsBody').innerHTML = top.length
    ? top.map((f) => `<tr>
        <td>${gradeTag(f.grade)}</td>
        <td><span class="patent-link" data-family="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span></td>
        <td><div class="tag-row">${[...f.countries].map((c) => `<span class="tag country">${escapeHtml(c)}</span>`).join('')}</div></td>
        <td>${scoreBar(f.score)}</td>
      </tr>`).join('')
    : `<tr><td colspan="4" class="empty-msg">평가된 특허가 없습니다.</td></tr>`;
}

/* ===== 권리 목록 ===== */
function buildFilters() {
  const tabs = [['', '전체'], ...countBy(db.rights, (r) => r.type).map(([k, v]) => [k, `${TYPE_LABEL[k]} ${v}`])];
  $('#typeTabs').innerHTML = tabs
    .map(([k, label]) => `<button class="scope-tab${filters.type === k ? ' active' : ''}" data-type="${k}">${escapeHtml(label)}</button>`)
    .join('');
  $$('#typeTabs .scope-tab').forEach((b) =>
    b.addEventListener('click', () => { filters.type = b.dataset.type; buildFilters(); renderList(); })
  );

  const fill = (sel, entries, labelFn) => {
    const el = $(sel);
    const keep = el.querySelector('option').outerHTML;
    el.innerHTML = keep + entries.map(([k, v]) => `<option value="${escapeHtml(k)}"${el.value === k ? ' selected' : ''}>${escapeHtml(labelFn(k))} (${v})</option>`).join('');
  };
  fill('#filterCountry', countBy(db.rights, (r) => r.country), (k) => COUNTRY_LABEL[k] || k);
  fill('#filterStatus', countBy(db.rights, (r) => r.status), (k) => STATUS_LABEL[k] || k);
}

function filtered() {
  const q = filters.q.trim().toLowerCase();
  return db.rights.filter((r) => {
    if (filters.type && r.type !== filters.type) return false;
    if (filters.country && r.country !== filters.country) return false;
    if (filters.status && r.status !== filters.status) return false;
    if (!q) return true;
    return [r.title, r.applicationNumber, r.registrationNumber, r.inventorText, r.family, r.mgmtNo, r.applicant]
      .join(' ').toLowerCase().includes(q);
  });
}

function sortRows(rows, state) {
  const { key, dir } = state;
  const mul = dir === 'asc' ? 1 : -1;
  const GRADE_ORDER = { S: 4, A: 3, B: 2, C: 1, NA: 0 };
  return [...rows].sort((a, b) => {
    let x = a[key], y = b[key];
    if (key === 'grade') { x = GRADE_ORDER[x]; y = GRADE_ORDER[y]; }
    // null 은 정렬 방향과 무관하게 항상 뒤로 — 빈 값이 상위를 차지하면 표가 쓸모없어진다
    if (x == null && y == null) return 0;
    if (x == null) return 1;
    if (y == null) return -1;
    if (typeof x === 'number' && typeof y === 'number') return (x - y) * mul;
    return String(x).localeCompare(String(y), 'ko') * mul;
  });
}

function markSorted(tableSel, state) {
  $$(`${tableSel} thead th`).forEach((th) => {
    const on = th.dataset.sort === state.key;
    th.classList.toggle('sorted', on);
    th.textContent = th.textContent.replace(/ [▲▼]$/, '') + (on ? (state.dir === 'asc' ? ' ▲' : ' ▼') : '');
  });
}

function renderList() {
  const rows = sortRows(filtered(), sortState.list);
  $('#listHint').textContent = `${rows.length}건 표시 (전체 ${db.rights.length}건) · 컬럼 클릭 시 정렬`;
  $('#rightsTable tbody').innerHTML = rows.length
    ? rows.map((r) => `<tr>
        <td>${typeTag(r.type)}</td>
        <td><span class="patent-link" data-id="${r.id}">${escapeHtml(r.title)}</span>${r.niceClass ? ` <span class="tag">${escapeHtml(r.niceClass)}류</span>` : ''}</td>
        <td><span class="tag country">${escapeHtml(r.country)}</span></td>
        <td><span class="tag ${r.status}">${STATUS_LABEL[r.status] || r.status}</span></td>
        <td class="num dim">${escapeHtml(r.applicationNumber || '-')}</td>
        <td class="num">${fmtDate(r.applicationDate)}</td>
        <td class="num dim">${escapeHtml(r.registrationNumber || '-')}</td>
        <td>${escapeHtml(r.inventorText || '-')}</td>
        <td>${r.type === 'patent' ? gradeTag(r.grade) : '<span class="dim">-</span>'}</td>
      </tr>`).join('')
    : `<tr><td colspan="9" class="empty-msg">조건에 맞는 권리가 없습니다.</td></tr>`;
  markSorted('#rightsTable', sortState.list);
}

/* ===== 특허 패밀리 ===== */
function familyStats() {
  const m = new Map();
  patents().forEach((p) => {
    if (!m.has(p.family)) {
      m.set(p.family, { name: p.family, items: [], countries: new Set(), inventors: new Set(), evaluation: p.evaluation });
    }
    const f = m.get(p.family);
    f.items.push(p);
    f.countries.add(p.country);
    p.inventors.forEach((i) => f.inventors.add(i));
    if (Object.keys(p.evaluation || {}).length) f.evaluation = p.evaluation;
  });
  return [...m.values()].map((f) => {
    const score = scoreOf(f.evaluation);
    return { ...f, count: f.items.length, score, grade: gradeOf(score), registered: f.items.filter((i) => i.status === 'registered').length };
  });
}

function renderFamily() {
  const fams = familyStats().sort((a, b) => b.count - a.count || (b.score ?? 0) - (a.score ?? 0));

  drawChart('familyChart', {
    type: 'bar',
    data: {
      labels: fams.map((f) => (f.name.length > 30 ? f.name.slice(0, 29) + '…' : f.name)),
      datasets: [{ label: '보유 건수', data: fams.map((f) => f.count), backgroundColor: PALETTE[0], borderRadius: 3, maxBarThickness: 22 }],
    },
    options: {
      ...baseOptions(),
      indexAxis: 'y',
      plugins: {
        ...baseOptions().plugins,
        legend: { display: false },
        // 라벨을 잘라 표시하므로 툴팁에서는 전체 명칭을 보여준다
        tooltip: { ...baseOptions().plugins.tooltip, callbacks: { title: (items) => fams[items[0].dataIndex].name } },
      },
      scales: { x: { ...gridX(), ticks: { ...tickStyle(), precision: 0 } }, y: { ...gridY(false), ticks: { ...tickStyle(), font: { size: 10 } } } },
    },
  });

  $('#familyBody').innerHTML = fams.length
    ? fams.map((f) => `
        <div class="family-group">
          <div class="family-name">
            <span class="family-dot"></span>${escapeHtml(f.name)}
            <span class="family-count">${f.count}건 · ${f.countries.size}개국</span>
            ${f.score != null ? gradeTag(f.grade) : ''}
          </div>
          <div class="family-items">
            ${f.items.map((p) => `<div class="family-item" data-id="${p.id}">
              <div class="family-item-meta">
                <span class="tag country">${escapeHtml(p.country)}</span>
                <span class="tag ${p.status}">${STATUS_LABEL[p.status] || p.status}</span>
                <span class="dim num">${fmtDate(p.applicationDate)}</span>
              </div>
              <div class="family-item-no num">${escapeHtml(p.registrationNumber || p.applicationNumber || '')}</div>
            </div>`).join('')}
          </div>
        </div>`).join('')
    : `<div class="empty-msg">등록된 특허가 없습니다.</div>`;
}

/* ===== 키워드 ===== */
function keywordStats() {
  // 같은 발명의 다국 출원이 중복으로 잡히지 않게 패밀리로 먼저 묶는다
  const fams = new Map();
  patents().forEach((p) => {
    if (!fams.has(p.family)) fams.set(p.family, []);
    fams.get(p.family).push(p);
  });

  const rows = KEYWORDS.map(([label, re]) => ({ label, re, families: [] }));
  let unmatched = 0;
  fams.forEach((items, name) => {
    const flat = name.replace(/\s+/g, '');
    const hit = rows.filter((r) => r.re.test(flat));
    if (!hit.length) { unmatched += 1; return; }
    hit.forEach((r) => r.families.push({ name, items }));
  });

  rows.forEach((r) => {
    r.familyCount = r.families.length;
    r.patentCount = r.families.reduce((n, f) => n + f.items.length, 0);
  });
  return {
    rows: rows.filter((r) => r.familyCount).sort((a, b) => b.familyCount - a.familyCount || a.label.localeCompare(b.label)),
    totalFamilies: fams.size,
    unmatched,
  };
}

function renderKeywords() {
  const { rows, totalFamilies, unmatched } = keywordStats();
  const top = rows[0];
  // 2개 이상 키워드에 걸친 패밀리 = 기술이 겹쳐 있는 정도
  const overlap = new Map();
  rows.forEach((r) => r.families.forEach((f) => overlap.set(f.name, (overlap.get(f.name) || 0) + 1)));
  const multi = [...overlap.values()].filter((n) => n > 1).length;

  $('#keywordKpiRow').innerHTML = [
    kpi('키워드', rows.length, `패밀리 ${totalFamilies - unmatched}개 분류`, 'tags', true),
    kpi('최다 키워드', top ? top.familyCount : 0, top ? `${top.label} · 특허 ${top.patentCount}건` : '해당 없음', 'trending-up'),
    kpi('복합 기술 패밀리', multi, '2개 이상 키워드에 걸친 건', 'layers'),
    kpi('미분류', unmatched, unmatched ? '키워드 사전 보완 필요' : '전 패밀리 분류됨', 'help-circle'),
  ].join('');

  drawChart('keywordChart', {
    type: 'bar',
    data: {
      labels: rows.map((r) => r.label),
      // 길이가 크기를 이미 나타내므로 색은 단일 색상으로 둔다
      datasets: [{ data: rows.map((r) => r.familyCount), backgroundColor: '#4263eb', borderRadius: 4, barThickness: 18 }],
    },
    options: {
      ...baseOptions(),
      indexAxis: 'y',
      plugins: {
        ...baseOptions().plugins,
        legend: { display: false },
        tooltip: {
          ...baseOptions().plugins.tooltip,
          callbacks: {
            label: (ctx) => `패밀리 ${ctx.parsed.x}개 · 특허 ${rows[ctx.dataIndex].patentCount}건`,
          },
        },
      },
      scales: { x: { ...gridY(), ticks: { ...tickStyle(), precision: 0 } }, y: gridX() },
      onClick: (_e, els) => { if (els.length) openKeyword(rows[els[0].index].label); },
    },
  });

  renderKeywordHeat(rows);
}

function renderKeywordHeat(rows) {
  const years = [...new Set(patents().map((p) => yearOf(p.applicationDate)).filter(Boolean))].sort();
  const at = (r, y) => r.families.reduce((n, f) => n + f.items.filter((p) => yearOf(p.applicationDate) === y).length, 0);
  const max = Math.max(1, ...rows.flatMap((r) => years.map((y) => at(r, y))));
  // 0 은 램프를 쓰지 않는다 — 빈 칸과 '조금 있음' 이 같은 색으로 보이면 안 된다
  const step = (v) => Math.min(HEAT_RAMP.length - 1, Math.ceil((v / max) * HEAT_RAMP.length) - 1);

  $('#keywordHeatmap').innerHTML = `
    <table class="heatmap">
      <thead>
        <tr><th>키워드</th>${years.map((y) => `<th class="num">${y}</th>`).join('')}<th class="num">합계</th></tr>
      </thead>
      <tbody>
        ${rows.map((r) => `<tr>
          <td class="heat-label"><span class="keyword-link" data-keyword="${escapeHtml(r.label)}">${escapeHtml(r.label)}</span></td>
          ${years.map((y) => {
            const v = at(r, y);
            if (!v) return '<td class="heat-cell heat-zero" title="0건">·</td>';
            const i = step(v);
            return `<td class="heat-cell" style="background:${HEAT_RAMP[i]};color:${HEAT_INK[i]}" title="${escapeHtml(r.label)} · ${y}년 ${v}건">${v}</td>`;
          }).join('')}
          <td class="num heat-total">${r.patentCount}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

function openKeyword(label) {
  const row = keywordStats().rows.find((r) => r.label === label);
  if (!row) return;
  $('#modalTitle').textContent = label;
  $('#modalSub').innerHTML = [
    `<span class="tag">패밀리 ${row.familyCount}개</span>`,
    `<span class="tag">특허 ${row.patentCount}건</span>`,
  ].join('');
  const fams = [...row.families].sort((a, b) => b.items.length - a.items.length);
  $('#modalBody').innerHTML = `
    <div class="detail-section">
      <h4 class="detail-title">해당 특허 패밀리 (${fams.length})</h4>
      <div class="related-list">
        ${fams.map((f) => `<div class="related-item" data-family="${escapeHtml(f.name)}">
          <span class="related-item-title">${escapeHtml(f.name)}</span>
          <span class="dim num">${f.items.length}건</span>
          <span class="tag country">${escapeHtml([...new Set(f.items.map((p) => p.country))].join(' '))}</span>
        </div>`).join('')}
      </div>
    </div>`;
  $('#modal').hidden = false;
  if (window.lucide) lucide.createIcons();
}

/* ===== 발명자 ===== */
function renderInventors() {
  const m = new Map();
  patents().forEach((p) => {
    p.inventors.forEach((name) => {
      if (!m.has(name)) m.set(name, { name, total: 0, registered: 0, scores: [], fams: new Map() });
      const s = m.get(name);
      s.total++;
      if (p.status === 'registered') s.registered++;
      if (p.score != null) s.scores.push(p.score);
      s.fams.set(p.family, (s.fams.get(p.family) || 0) + 1);
    });
  });
  const stats = [...m.values()].map((s) => ({
    ...s,
    families: s.fams.size,
    avgScore: s.scores.length ? s.scores.reduce((a, b) => a + b, 0) / s.scores.length : null,
    familyText: [...s.fams.entries()].sort((a, b) => b[1] - a[1]).slice(0, 2).map(([k]) => (k.length > 22 ? k.slice(0, 21) + '…' : k)).join(', '),
  }));

  const pats = patents();
  const co = pats.filter((p) => p.inventors.length > 1).length;
  const top1 = [...stats].sort((a, b) => b.total - a.total)[0];

  $('#inventorKpiRow').innerHTML = [
    kpi('총 발명자', stats.length, '중복 제거 기준', 'users', true),
    kpi('공동발명 건', co, `단독발명 ${pats.filter((p) => p.inventors.length === 1).length}건`, 'user-plus'),
    kpi('건당 평균 발명자', pats.length ? (pats.reduce((a, p) => a + p.inventors.length, 0) / pats.length).toFixed(1) : '0', '명', 'user-round'),
    kpi('최다 참여', top1 ? escapeHtml(top1.name) : '-', top1 ? `${top1.total}건 · 발명 ${top1.families}개` : '', 'crown'),
  ].join('');

  const top = [...stats].sort((a, b) => b.total - a.total).slice(0, 15);
  drawChart('inventorChart', {
    type: 'bar',
    data: {
      labels: top.map((s) => s.name),
      datasets: [
        { label: '등록', data: top.map((s) => s.registered), backgroundColor: PALETTE[0], borderRadius: 3, stack: 's', maxBarThickness: 30 },
        { label: '그 외', data: top.map((s) => s.total - s.registered), backgroundColor: PALETTE[3], borderRadius: 3, stack: 's', maxBarThickness: 30 },
      ],
    },
    options: {
      ...baseOptions(),
      indexAxis: 'y',
      scales: { x: { ...gridX(), stacked: true, ticks: { ...tickStyle(), precision: 0 } }, y: { ...gridY(false), stacked: true } },
    },
  });

  const rows = sortRows(stats, sortState.inventor);
  $('#inventorTable tbody').innerHTML = rows.length
    ? rows.map((s) => `<tr>
        <td><b style="color:var(--text)">${escapeHtml(s.name)}</b></td>
        <td class="num">${s.total}</td>
        <td class="num">${s.families}</td>
        <td class="num">${s.registered}</td>
        <td>${scoreBar(s.avgScore)}</td>
        <td class="dim">${escapeHtml(s.familyText || '-')}</td>
      </tr>`).join('')
    : `<tr><td colspan="6" class="empty-msg">발명자 정보가 없습니다.</td></tr>`;
  markSorted('#inventorTable', sortState.inventor);
}

/* ===== 기술평가 ===== */
function renderEvaluation() {
  const fams = familyStats();
  const scored = fams.filter((f) => f.score != null);
  const byGrade = (g) => fams.filter((f) => f.grade === g).length;
  const axisAvg = AXES.map((a) => {
    const vals = fams.map((f) => f.evaluation?.[a.key]).filter((v) => typeof v === 'number' && v > 0);
    return vals.length ? vals.reduce((x, y) => x + y, 0) / vals.length : 0;
  });

  $('#evalKpiRow').innerHTML = [
    kpi('평가 완료', `${scored.length}<small>/${fams.length}</small>`, '특허 패밀리 기준', 'clipboard-check', true),
    kpi('S등급', byGrade('S'), '종합 4.5↑ · 최우선 유지', 'star'),
    kpi('A등급', byGrade('A'), '종합 3.5↑ · 핵심 자산', 'thumbs-up'),
    kpi('B·C등급', byGrade('B') + byGrade('C'), '유지 실익 재검토 대상', 'alert-circle'),
  ].join('');

  drawChart('radarChart', {
    type: 'radar',
    data: {
      labels: AXES.map((a) => a.label),
      datasets: [{
        label: '패밀리 평균', data: axisAvg,
        backgroundColor: 'rgba(66, 99, 235, 0.15)', borderColor: PALETTE[0], borderWidth: 2,
        pointBackgroundColor: PALETTE[0], pointRadius: 4,
      }],
    },
    options: {
      ...baseOptions(),
      plugins: { ...baseOptions().plugins, legend: { display: false } },
      scales: {
        r: {
          min: 0, max: 5,
          ticks: { stepSize: 1, backdropColor: 'transparent', color: '#9a9fa8', font: { size: 10 } },
          grid: { color: '#ececf1' }, angleLines: { color: '#ececf1' },
          pointLabels: { color: '#495057', font: { size: 12, weight: '600' } },
        },
      },
    },
  });

  const gradeEntries = ['S', 'A', 'B', 'C'].map((g) => [g, byGrade(g)]).filter(([, v]) => v > 0);
  drawDoughnut('gradeChart', gradeEntries.map(([g, v]) => [`${g}등급`, v]), gradeEntries.map(([g]) => GRADE_COLOR[g]));

  drawChart('matrixChart', {
    type: 'bubble',
    data: {
      datasets: [{
        // 같은 좌표에 여러 발명이 겹치므로 점 크기로 보유 건수를, 툴팁으로 명칭을 구분한다
        data: scored.map((f) => ({
          x: f.evaluation.technology ?? 0,
          y: f.evaluation.marketability ?? 0,
          r: 6 + f.count * 2,
          name: f.name, count: f.count, grade: f.grade,
        })),
        backgroundColor: 'rgba(66, 99, 235, 0.5)',
        borderColor: PALETTE[0],
        borderWidth: 1.5,
      }],
    },
    options: {
      ...baseOptions(),
      plugins: {
        ...baseOptions().plugins,
        legend: { display: false },
        tooltip: {
          ...baseOptions().plugins.tooltip,
          callbacks: {
            title: (i) => i[0].raw.name,
            label: (ctx) => [`${ctx.raw.grade}등급 · ${ctx.raw.count}건 보유`, `기술성 ${ctx.raw.x} · 시장성 ${ctx.raw.y}`],
          },
        },
      },
      onClick: (evt, els) => {
        if (els.length) openFamily(charts.matrixChart.data.datasets[0].data[els[0].index].name);
      },
      // 점수는 1~5 뿐이라 0부터 그리면 캔버스 아래·왼쪽 절반이 늘 빈 채로 남는다.
      // min 을 0.5 로 두면 눈금이 0.5/1.5/... 로 찍혀서 정수로 고정한다
      scales: {
        x: { ...gridX(), min: 0.5, max: 5.5, title: { display: true, text: '기술성 →', color: '#9a9fa8', font: { size: 11 } }, ticks: tickStyle(), afterBuildTicks: (a) => { a.ticks = [1, 2, 3, 4, 5].map((v) => ({ value: v })); } },
        y: { ...gridY(), min: 0.5, max: 5.5, title: { display: true, text: '시장성 →', color: '#9a9fa8', font: { size: 11 } }, ticks: tickStyle(), afterBuildTicks: (a) => { a.ticks = [1, 2, 3, 4, 5].map((v) => ({ value: v })); } },
      },
    },
    plugins: [quadrantPlugin],
  });

  const rows = sortRows(fams.map((f) => ({
    ...f,
    technology: f.evaluation?.technology ?? null,
    marketability: f.evaluation?.marketability ?? null,
    rights: f.evaluation?.rights ?? null,
    utilization: f.evaluation?.utilization ?? null,
  })), sortState.eval);
  const cell = (v) => (v == null ? '<span class="dim">-</span>' : `<span class="num">${v}</span>`);
  $('#evalTable tbody').innerHTML = rows.length
    ? rows.map((f) => `<tr>
        <td>${gradeTag(f.grade)}</td>
        <td><span class="patent-link" data-family="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span></td>
        <td class="num dim">${f.count}</td>
        <td>${cell(f.technology)}</td><td>${cell(f.marketability)}</td>
        <td>${cell(f.rights)}</td><td>${cell(f.utilization)}</td>
        <td>${scoreBar(f.score)}</td>
      </tr>`).join('')
    : `<tr><td colspan="8" class="empty-msg">등록된 특허가 없습니다.</td></tr>`;
  markSorted('#evalTable', sortState.eval);
}

/* ===== 상세 모달 ===== */
const row = (k, v) => `<div class="detail-key">${k}</div><div class="detail-val">${v}</div>`;

function evalBlock(ev, score) {
  if (score == null) return '<div class="detail-text dim">아직 평가되지 않았습니다.</div>';
  return `<div class="eval-bars">
      ${AXES.map((a) => {
        const v = ev?.[a.key];
        return `<div class="eval-bar">
          <span class="eval-bar-label">${a.label}</span>
          <div class="score-track"><div class="score-fill" style="width:${((v || 0) / 5) * 100}%"></div></div>
          <span class="score-value">${v ?? '-'}</span>
        </div>`;
      }).join('')}
      <div class="eval-bar" style="margin-top:4px;padding-top:10px;border-top:1px solid var(--border)">
        <span class="eval-bar-label"><b>종합</b></span>
        <div class="score-track"><div class="score-fill" style="width:${(score / 5) * 100}%"></div></div>
        <span class="score-value"><b>${score.toFixed(1)}</b></span>
      </div>
    </div>
    ${ev?.comment ? `<div class="eval-comment">${escapeHtml(ev.comment)}</div>` : ''}`;
}

function openDetail(id) {
  const r = db.rights.find((x) => x.id === id);
  if (!r) return;
  $('#modalTitle').textContent = r.title;
  $('#modalSub').innerHTML = [
    typeTag(r.type),
    `<span class="tag country">${escapeHtml(COUNTRY_LABEL[r.country] || r.country)}</span>`,
    `<span class="tag ${r.status}">${STATUS_LABEL[r.status] || r.status}</span>`,
    r.niceClass ? `<span class="tag">${escapeHtml(r.niceClass)}류</span>` : '',
    r.type === 'patent' && r.score != null ? gradeTag(r.grade) : '',
  ].join('');

  const related = r.type === 'patent' ? db.rights.filter((x) => x.type === 'patent' && x.family === r.family && x.id !== r.id) : [];

  $('#modalBody').innerHTML = `
    <div class="detail-section">
      <h4 class="detail-title">서지 정보</h4>
      <div class="detail-grid">
        ${row('관리번호', escapeHtml(r.mgmtNo || '-'))}
        ${row('출원번호', escapeHtml(r.applicationNumber || '-'))}
        ${row('출원일', fmtDate(r.applicationDate))}
        ${row('등록번호', escapeHtml(r.registrationNumber || '-'))}
        ${row('등록일', fmtDate(r.registrationDate))}
        ${row('DUE DATE', fmtDate(r.dueDate))}
        ${row('출원인', escapeHtml(r.applicant || '-'))}
        ${r.inventorText ? row('발명자', escapeHtml(r.inventorText)) : ''}
        ${r.notes ? row('비고', escapeHtml(r.notes)) : ''}
        ${r.docRef ? row('증서', escapeHtml(r.docRef)) : ''}
      </div>
    </div>

    ${r.type === 'patent' ? `
    <div class="detail-section">
      <h4 class="detail-title">기술평가 <span class="dim" style="font-weight:400;text-transform:none;letter-spacing:0">— 패밀리 「${escapeHtml(r.family)}」 공통</span></h4>
      ${evalBlock(r.evaluation, r.score)}
    </div>

    <div class="detail-section">
      <h4 class="detail-title">같은 발명의 다른 국가 출원 (${related.length})</h4>
      ${related.length
        ? `<div class="related-list">${related.map((x) => `<div class="related-item" data-id="${x.id}">
            <span class="tag country">${escapeHtml(x.country)}</span>
            <span class="related-item-title">${escapeHtml(x.registrationNumber || x.applicationNumber || '-')}</span>
            <span class="tag ${x.status}">${STATUS_LABEL[x.status] || x.status}</span>
          </div>`).join('')}</div>`
        : '<div class="detail-text dim">이 발명은 해당 국가에만 출원되어 있습니다.</div>'}
    </div>` : ''}`;

  $('#modal').hidden = false;
  if (window.lucide) lucide.createIcons();
}

/** 평가 랭킹·TOP5 에서는 개별 건이 아니라 패밀리를 열어 구성 건을 한 번에 보여준다 */
function openFamily(name) {
  const f = familyStats().find((x) => x.name === name);
  if (!f) return;
  $('#modalTitle').textContent = f.name;
  $('#modalSub').innerHTML = [
    typeTag('patent'),
    `<span class="tag">${f.count}건 보유</span>`,
    `<span class="tag">${f.countries.size}개국</span>`,
    f.score != null ? gradeTag(f.grade) : '',
  ].join('');

  $('#modalBody').innerHTML = `
    <div class="detail-section">
      <h4 class="detail-title">기술평가</h4>
      ${evalBlock(f.evaluation, f.score)}
    </div>
    <div class="detail-section">
      <h4 class="detail-title">구성 특허 (${f.count})</h4>
      <div class="related-list">
        ${f.items.map((p) => `<div class="related-item" data-id="${p.id}">
          <span class="tag country">${escapeHtml(p.country)}</span>
          <span class="related-item-title">${escapeHtml(p.registrationNumber || p.applicationNumber || '-')}</span>
          <span class="dim num">${fmtDate(p.applicationDate)}</span>
          <span class="tag ${p.status}">${STATUS_LABEL[p.status] || p.status}</span>
        </div>`).join('')}
      </div>
    </div>
    <div class="detail-section">
      <h4 class="detail-title">발명자</h4>
      <div class="tag-row">${[...f.inventors].map((i) => `<span class="tag">${escapeHtml(i)}</span>`).join('') || '<span class="dim">-</span>'}</div>
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
    const cx = x.getPixelForValue(3), cy = y.getPixelForValue(3);
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
      tooltip: { backgroundColor: '#1f2329', padding: 12, cornerRadius: 8, titleFont: { size: 12 }, bodyFont: { size: 12 }, displayColors: false },
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
      datasets: [{ data: entries.map(([, v]) => v), backgroundColor: colors || PALETTE, borderWidth: 2, borderColor: '#fff' }],
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
              return `${ctx.label}: ${ctx.parsed}건 (${total ? ((ctx.parsed / total) * 100).toFixed(1) : '0.0'}%)`;
            },
          },
        },
      },
    },
  });
}

/* ===== 뷰 전환 & 이벤트 ===== */
function switchView(view) {
  ['overview', 'list', 'family', 'keyword', 'inventors', 'evaluation'].forEach((v) => {
    $(`#${v}View`).hidden = v !== view;
  });
  $$('.nav-item').forEach((el) => el.classList.toggle('active', el.dataset.view === view));
  $('#pageTitle').textContent = VIEW_TITLE[view];
  // 숨겨진 canvas 는 크기가 0으로 잡혀 차트가 찌그러진다 — 표시 후 리사이즈
  Object.values(charts).forEach((c) => c.resize());
}

function bind() {
  $$('.nav-item').forEach((el) =>
    el.addEventListener('click', (e) => { e.preventDefault(); switchView(el.dataset.view); })
  );

  // 상세 열기 — 목록/카드/모달 내 관련건 모두 위임으로 처리
  document.addEventListener('click', (e) => {
    const kw = e.target.closest('[data-keyword]');
    if (kw) return openKeyword(kw.dataset.keyword);
    const fam = e.target.closest('[data-family]');
    if (fam) return openFamily(fam.dataset.family);
    const t = e.target.closest('[data-id]');
    if (t) openDetail(t.dataset.id);
  });

  $$('[data-close]').forEach((el) => el.addEventListener('click', () => ($('#modal').hidden = true)));
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') $('#modal').hidden = true; });

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
  bindSort('#rightsTable', 'list', renderList);
  bindSort('#inventorTable', 'inventor', renderInventors);
  bindSort('#evalTable', 'eval', renderEvaluation);

  $('#searchInput').addEventListener('input', (e) => { filters.q = e.target.value; renderList(); });
  $('#filterCountry').addEventListener('change', (e) => { filters.country = e.target.value; renderList(); });
  $('#filterStatus').addEventListener('change', (e) => { filters.status = e.target.value; renderList(); });
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
