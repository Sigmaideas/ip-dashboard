#!/usr/bin/env node
/**
 * IP 관리대장 엑셀 → data/rights.json 변환기
 *
 *   node scripts/import-excel.js inbox/특허상표디자인_리스트_엑스와이지.xlsx
 *   node scripts/import-excel.js <파일> --sheet "통합 관리 시트" --dry
 *
 * 태창특허법률사무소와 공유하는 「통합 관리 시트」를 기준으로 맞췄고,
 * 컬럼명이 조금 달라져도 COLUMN_ALIASES 로 흡수한다.
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'data', 'rights.json');
const EVAL_FILE = path.join(ROOT, 'data', 'evaluations.json');
const DEFAULT_SHEET = '통합 관리 시트';

const COMPANY = '주식회사 엑스와이지';
// 표기 흔들림 — 시트에 띄어쓰기 유무·공동출원 표기가 섞여 있다
const COMPANY_ALIASES = ['주식회사 엑스와이지', '주식회사엑스와이지', '(주)엑스와이지', '엑스와이지', 'XYZ, Inc.', '주식회사 엑스와이지부산', '주식회사 코봇'];

const COLUMN_ALIASES = {
  type: ['구분', '권리', '권리구분'],
  mgmtNo: ['관리번호', 'refno', 'ref.no'],
  country: ['국가', '출원국'],
  gradeRaw: ['등급'],
  applicant: ['출원인/특허권자', '출원인(특허권자)', '출원인', '특허권자', '권리자'],
  inventors: ['발명자', '발명인'],
  title: ['명칭', '명칭(상표)', '발명의명칭', '발명의 명칭'],
  niceClass: ['상품류', '류'],
  applicationDate: ['출원일', '출원일자'],
  applicationNumber: ['출원번호'],
  registrationDate: ['등록일', '등록일자'],
  registrationNumber: ['등록번호'],
  status: ['현재상태', '상태', '권리상태'],
  dueDate: ['duedate', 'due date'],
  updatedAt: ['업데이트'],
  annuity: ['연차등록료 납부현황', '연차등록료', '연차료'],
  notes: ['비고', '메모'],
};

const TYPE_MAP = { 특허: 'patent', 상표: 'trademark', 디자인: 'design', 저작권: 'copyright', 실용신안: 'utility' };
const STATUS_MAP = {
  등록: 'registered',
  출원: 'pending',
  심사중: 'examining',
  PCT: 'pct',
  거절: 'rejected',
  포기: 'abandoned',
  이관: 'transferred',
};
const COUNTRY_MAP = { KR: 'KR', US: 'US', CN: 'CN', JP: 'JP', PCT: 'PCT', EP: 'EP', 싱가폴: 'SG', 싱가포르: 'SG', 인도: 'IN', 대만: 'TW', 베트남: 'VN' };

const norm = (s) => String(s ?? '').toLowerCase().replace(/[\s_()·.\-/]/g, '');

function buildHeaderMap(headers) {
  const map = {};
  const used = new Set();
  headers.forEach((h, i) => {
    const n = norm(h);
    if (!n) return;
    for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (map[field] !== undefined) continue;
      if (aliases.some((a) => norm(a) === n)) { map[field] = i; used.add(i); return; }
    }
  });
  // 완전일치로 못 찾은 것만 포함 관계로 완화 매칭 — '등록일'이 '등록일자'보다 먼저 걸리는 사고를 막는다
  headers.forEach((h, i) => {
    const n = norm(h);
    if (!n || used.has(i)) return;
    for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (map[field] !== undefined) continue;
      if (aliases.some((a) => norm(a).length >= 3 && n.includes(norm(a)))) { map[field] = i; used.add(i); return; }
    }
  });
  return { map, used };
}

/**
 * 엑셀 날짜.
 * cellDates:true 로 읽으면 xlsx 가 UTC 기준 Date 를 만들어 KST 기준 하루 전으로 밀린다
 * (2015-03-02 → 2015-03-01T14:59Z). 그래서 시리얼 넘버 그대로 받아 직접 변환한다.
 */
function toDate(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d || !d.y) return null;
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  if (v instanceof Date) {
    // 혹시 Date 로 들어오면 UTC 밀림을 되돌린다
    const t = new Date(v.getTime() + 9 * 3600 * 1000);
    return t.toISOString().slice(0, 10);
  }
  const s = String(v).trim().replace(/[.\/]/g, '-');
  const m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return null;
}

function toCountry(v) {
  const s = String(v ?? '').trim();
  if (COUNTRY_MAP[s]) return COUNTRY_MAP[s];
  const u = s.toUpperCase();
  return COUNTRY_MAP[u] || (/^[A-Z]{2}$/.test(u) ? u : 'KR');
}

function toStatus(v, regNo) {
  const s = String(v ?? '').trim();
  if (STATUS_MAP[s]) return STATUS_MAP[s];
  for (const [k, val] of Object.entries(STATUS_MAP)) if (s && norm(s).includes(norm(k))) return val;
  return regNo ? 'registered' : 'pending';
}

/**
 * 발명자.
 * 시트에 '민중후(정승교, 박혁, 김희철)' 처럼 대표 발명자 뒤 괄호로 공동발명자를 적은 행이 있어
 * 괄호를 벗겨 전부 개별 발명자로 편다. 상표·디자인은 'N/A' 라 빈 배열이 된다.
 */
function toInventors(v) {
  const s = String(v ?? '').trim();
  if (!s || /^n\/?a$/i.test(s)) return [];
  return [...new Set(
    s.replace(/[()（）]/g, ',')
      .split(/[,;、\n]/)
      .map((x) => x.trim())
      .filter((x) => x && !/^n\/?a$/i.test(x))
  )];
}

/** 패밀리 키 — 같은 발명이 KR/US/CN/PCT 로 나뉘어 있어 명칭을 정규화해 묶는다 */
const familyKey = (title) => String(title ?? '').replace(/\s+/g, ' ').trim();

// 명칭이 조금씩 달라 자동으로는 안 묶이지만 실질적으로 같은 계열인 건들.
// '로봇 팔 / 로봇 움직임 / 이동 로봇' 은 위험도 판단이라는 같은 안전제어 발명의 범위 차이다.
const FAMILY_MERGE = [
  { name: '위험도 판단 기반 로봇 안전제어', match: /위험도 판단/ },
  { name: '모듈화된 로봇 플랫폼', match: /^모듈화된 로봇 플랫폼/ },
  { name: 'AI 자율주행 스마트 물류', match: /AI기반의 자율주행/ },
  { name: '자율주행 차량 음식 주문·전달', match: /자율 주행 차량을 이용한 음식/ },
  { name: '방역 로봇', match: /방역/ },
];

function resolveFamily(title) {
  const t = familyKey(title);
  if (!t) return '미분류';
  const hit = FAMILY_MERGE.find((f) => f.match.test(t));
  return hit ? hit.name : t;
}

function main() {
  const args = process.argv.slice(2);
  const file = args.find((a) => !a.startsWith('--'));
  const dry = args.includes('--dry');
  const sheetArg = args.includes('--sheet') ? args[args.indexOf('--sheet') + 1] : null;

  if (!file) {
    console.error('사용법: node scripts/import-excel.js <엑셀파일> [--sheet 시트명] [--dry]');
    process.exit(1);
  }
  const abs = path.isAbsolute(file) ? file : path.join(ROOT, file);
  if (!fs.existsSync(abs)) {
    console.error(`파일을 찾을 수 없습니다: ${abs}`);
    process.exit(1);
  }

  const wb = XLSX.readFile(abs); // cellDates 끔 — toDate() 주석 참고
  const sheetName = sheetArg || (wb.SheetNames.includes(DEFAULT_SHEET) ? DEFAULT_SHEET : wb.SheetNames[0]);
  if (!wb.Sheets[sheetName]) {
    console.error(`시트 '${sheetName}' 없음. 사용 가능: ${wb.SheetNames.join(', ')}`);
    process.exit(1);
  }
  console.log(`파일: ${path.basename(abs)}`);
  console.log(`시트: ${sheetName} (전체: ${wb.SheetNames.join(', ')})`);

  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '', blankrows: false });
  if (!rows.length) { console.error('빈 시트입니다.'); process.exit(1); }

  let headerRow = 0, best = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const hits = buildHeaderMap(rows[i]).used.size;
    if (hits > best) { best = hits; headerRow = i; }
  }
  const headers = rows[headerRow];
  const { map, used } = buildHeaderMap(headers);
  console.log(`헤더 행: ${headerRow + 1}행 · 인식 컬럼 ${Object.keys(map).length}개`);
  const unmapped = headers.map((h, i) => (h && !used.has(i) ? h : null)).filter(Boolean);
  if (unmapped.length) console.log(`미매핑 → extra 보관: ${unmapped.join(', ')}`);

  const get = (row, f) => (map[f] === undefined ? '' : row[map[f]]);

  // 평가는 엑셀이 아니라 별도 파일에서 온다 (엑셀의 '등급' 열은 현재 전 행 공란).
  // 같은 발명의 KR/US/CN 건은 하나의 평가를 공유하므로 패밀리 단위로 붙인다.
  let evaluations = {};
  if (fs.existsSync(EVAL_FILE)) {
    evaluations = JSON.parse(fs.readFileSync(EVAL_FILE, 'utf8')).byFamily || {};
    console.log(`평가 데이터: ${Object.keys(evaluations).length}개 패밀리`);
  }

  const seq = {};
  const rights = rows
    .slice(headerRow + 1)
    .filter((row) => row.some((c) => String(c ?? '').trim() !== ''))
    .map((row) => {
      const type = TYPE_MAP[String(get(row, 'type')).trim()] || 'other';
      const title = String(get(row, 'title') ?? '').replace(/\s+/g, ' ').trim();
      const regNo = String(get(row, 'registrationNumber') ?? '').trim();
      const family = type === 'patent' ? resolveFamily(title) : familyKey(title) || '미분류';

      const prefix = { patent: 'PT', trademark: 'TM', design: 'DS', copyright: 'CR', other: 'OT' }[type];
      seq[prefix] = (seq[prefix] || 0) + 1;

      // 증서 파일명은 헤더가 비어 있는 꼬리 컬럼에 들어 있어 헤더 기준으로는 잡히지 않는다.
      // 행 전체에서 파일명/링크 패턴을 훑어 건진다
      let docRef = null;
      row.forEach((cell, ci) => {
        const s = String(cell ?? '').trim();
        if (!docRef && !used.has(ci) && /\.(pdf|jpg|png)$|^https?:\/\//i.test(s)) docRef = s;
      });

      const extra = {};
      headers.forEach((h, ci) => {
        if (h && !used.has(ci) && String(row[ci] ?? '').trim() !== '' && String(row[ci]).trim() !== docRef) {
          extra[String(h).trim()] = row[ci];
        }
      });

      return {
        id: `${prefix}${String(seq[prefix]).padStart(3, '0')}`,
        type,
        mgmtNo: String(get(row, 'mgmtNo') ?? '').trim() || null,
        title: title || '(명칭 미기재)',
        country: toCountry(get(row, 'country')),
        status: toStatus(get(row, 'status'), regNo),
        applicant: String(get(row, 'applicant') ?? '').replace(/\s+/g, ' ').trim() || COMPANY,
        inventors: toInventors(get(row, 'inventors')),
        niceClass: String(get(row, 'niceClass') ?? '').trim() || null,
        // 출원번호에 앞뒤 공백이 들어간 행이 있어 반드시 trim
        applicationNumber: String(get(row, 'applicationNumber') ?? '').trim() || null,
        applicationDate: toDate(get(row, 'applicationDate')),
        registrationNumber: regNo || null,
        registrationDate: toDate(get(row, 'registrationDate')),
        dueDate: toDate(get(row, 'dueDate')),
        annuity: String(get(row, 'annuity') ?? '').trim() || null,
        family,
        notes: String(get(row, 'notes') ?? '').trim() || null,
        docRef,
        evaluation: type === 'patent' ? evaluations[family] || {} : {},
        extra: Object.keys(extra).length ? extra : undefined,
      };
    });

  const out = {
    meta: {
      company: COMPANY,
      companyAliases: COMPANY_ALIASES,
      lastUpdated: new Date().toISOString(),
      source: path.basename(abs),
      sheet: sheetName,
    },
    rights,
  };

  const tally = (fn) => {
    const m = {};
    rights.forEach((r) => { const k = fn(r); m[k] = (m[k] || 0) + 1; });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${v}`).join(' · ');
  };
  const patents = rights.filter((r) => r.type === 'patent');
  console.log(`\n총 ${rights.length}건`);
  console.log(`  권리: ${tally((r) => r.type)}`);
  console.log(`  상태: ${tally((r) => r.status)}`);
  console.log(`  국가: ${tally((r) => r.country)}`);
  console.log(`  특허 패밀리: ${new Set(patents.map((r) => r.family)).size}개`);
  console.log(`  평가 적용: ${patents.filter((r) => Object.keys(r.evaluation).length).length}/${patents.length}건`);

  const noTitle = rights.filter((r) => r.title === '(명칭 미기재)');
  const noDate = rights.filter((r) => !r.applicationDate);
  if (noTitle.length || noDate.length) {
    console.log('\n확인 필요 (원본 엑셀 보완 대상):');
    noTitle.forEach((r) => console.log(`  · 명칭 없음 — ${r.id} ${r.country} ${r.applicationNumber || ''}`));
    noDate.forEach((r) => console.log(`  · 출원일 없음 — ${r.id} ${r.title.slice(0, 30)}`));
  }

  if (dry) {
    console.log('\n--dry — 파일 미저장. 첫 2건:');
    console.log(JSON.stringify(rights.slice(0, 2), null, 2));
    return;
  }
  fs.writeFileSync(OUT, JSON.stringify(out, null, 2) + '\n');
  console.log(`\n저장 완료 → ${path.relative(ROOT, OUT)}`);
}

main();
