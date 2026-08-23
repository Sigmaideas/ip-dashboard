#!/usr/bin/env node
/**
 * 특허 엑셀 → data/patents.json 변환기
 *
 *   node scripts/import-excel.js inbox/특허목록.xlsx
 *   node scripts/import-excel.js inbox/특허목록.xlsx --sheet "등록특허"
 *   node scripts/import-excel.js inbox/특허목록.xlsx --dry     # 파일 안 쓰고 매핑 결과만 출력
 *
 * 헤더 이름은 회사마다 제각각이라 COLUMN_ALIASES 로 흡수한다.
 * 매칭 안 된 컬럼은 버리지 않고 extra 에 담아 두므로, 나중에 별칭만 추가하면 재활용된다.
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'data', 'patents.json');

// 회사명 표기 흔들림 — 라운지랩에서 엑스와이지로 사명이 바뀌어 초기 출원은 옛 이름으로 남아 있다
const COMPANY = '주식회사 엑스와이지';
const COMPANY_ALIASES = ['주식회사 라운지랩', '라운지랩', '(주)라운지랩', 'LOUNGELAB', 'XYZ, Inc.', 'XYZ Inc', '엑스와이지', '(주)엑스와이지'];

const COLUMN_ALIASES = {
  title: ['발명의명칭', '발명명칭', '명칭', '특허명', '제목', 'title', '발명의 명칭'],
  titleEn: ['영문명칭', '영문 명칭', 'titleen', 'englishtitle'],
  country: ['국가', '출원국', '국가코드', 'country', '국별'],
  status: ['상태', '권리상태', '등록상태', '진행상태', 'status', '법적상태'],
  applicationNumber: ['출원번호', '출원 번호', 'applicationnumber', 'appno', '출원번호(국내)'],
  applicationDate: ['출원일', '출원일자', '출원 일자', 'applicationdate', 'filingdate'],
  registrationNumber: ['등록번호', '등록 번호', 'registrationnumber', 'patentnumber', '특허번호'],
  registrationDate: ['등록일', '등록일자', '등록 일자', 'registrationdate', 'grantdate'],
  publicationNumber: ['공개번호', 'publicationnumber'],
  publicationDate: ['공개일', '공개일자', 'publicationdate'],
  expiryDate: ['만료일', '존속기간만료일', '소멸일', 'expirydate'],
  applicant: ['출원인', '권리자', '특허권자', 'applicant', 'assignee', '출원인명'],
  inventors: ['발명자', '발명자명', '발명인', 'inventor', 'inventors'],
  ipc: ['ipc', 'ipc분류', '국제특허분류', 'cpc', '분류코드'],
  family: ['특허군', '패밀리', '패밀리명', 'family', '기술군', '관련특허군', '패밀리그룹'],
  category: ['기술분류', '분류', '카테고리', '기술분야', 'category', '기술구분'],
  productLine: ['적용제품', '제품', '제품군', '적용모델', 'product', 'productline'],
  abstract: ['요약', '초록', '개요', 'abstract', 'summary', '발명의요약'],
  claims: ['청구항수', '청구항', 'claims', '청구항 수'],
  notes: ['비고', '메모', 'note', 'notes', '특이사항'],
  // 기술평가 — 엑셀에 열이 있으면 그대로, 없으면 미평가로 남는다
  technology: ['기술성', '기술성점수', 'technology'],
  marketability: ['시장성', '시장성점수', 'marketability', '사업성'],
  rights: ['권리성', '권리성점수', 'rights', '권리범위'],
  utilization: ['활용도', '활용도점수', 'utilization', '실시여부'],
  evalComment: ['평가의견', '평가코멘트', '평가', 'evalcomment', '검토의견'],
};

const STATUS_MAP = {
  등록: 'registered', 등록결정: 'registered', registered: 'registered', granted: 'registered', 특허등록: 'registered',
  출원: 'pending', 출원중: 'pending', 심사중: 'pending', pending: 'pending', 공개: 'pending', 심사청구: 'pending',
  거절: 'rejected', 거절결정: 'rejected', rejected: 'rejected',
  포기: 'abandoned', 취하: 'abandoned', abandoned: 'abandoned', withdrawn: 'abandoned',
  소멸: 'expired', 만료: 'expired', expired: 'expired',
};

const norm = (s) => String(s ?? '').toLowerCase().replace(/[\s_()·.\-/]/g, '');

function buildHeaderMap(headers) {
  const map = {};
  const used = new Set();
  headers.forEach((h, i) => {
    const n = norm(h);
    if (!n) return;
    for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (map[field] !== undefined) continue;
      // 완전일치 우선, 없으면 포함 관계로 완화 매칭
      if (aliases.some((a) => norm(a) === n) || aliases.some((a) => n.includes(norm(a)) && norm(a).length >= 3)) {
        map[field] = i;
        used.add(i);
        return;
      }
    }
  });
  return { map, used };
}

/** 엑셀 날짜는 시리얼 넘버 / 'YYYY.MM.DD' / 'YYYY-MM-DD' 가 섞여 온다 */
function toDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const s = String(v).trim().replace(/[.\/]/g, '-').replace(/-+$/, '');
  const m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  const y = s.match(/^(\d{4})$/);
  return y ? `${y[1]}-01-01` : null;
}

function toCountry(v, appNo) {
  const s = String(v ?? '').trim().toUpperCase();
  if (/KR|한국|국내|대한민국/.test(s)) return 'KR';
  if (/US|미국|USA/.test(s)) return 'US';
  if (/CN|중국/.test(s)) return 'CN';
  if (/JP|일본/.test(s)) return 'JP';
  if (/EP|유럽|EU/.test(s)) return 'EP';
  if (/WO|PCT|국제/.test(s)) return 'WO';
  if (/^[A-Z]{2}$/.test(s)) return s;
  // 국가 열이 비었으면 출원번호 형태로 추정 — 10-2019-... 는 한국 양식
  const a = String(appNo ?? '');
  if (/^\d{2}-\d{4}-\d{6,7}/.test(a)) return 'KR';
  if (/^(US|\d{2}\/\d{6})/i.test(a)) return 'US';
  return 'KR';
}

function toStatus(v, regNo) {
  const n = norm(v);
  for (const [k, val] of Object.entries(STATUS_MAP)) if (n && n.includes(norm(k))) return val;
  return regNo ? 'registered' : 'pending';
}

/** IPC 는 'B25J 11/00' 처럼 슬래시가 코드의 일부라 분리자에서 빼야 한다 */
function toList(v, sep = /[,;/、·\n]/) {
  if (v == null || v === '') return [];
  if (Array.isArray(v)) return v.map(String).map((s) => s.trim()).filter(Boolean);
  return String(v).split(sep).map((s) => s.trim()).filter(Boolean);
}
const toCodeList = (v) => toList(v, /[,;、\n]/);

function toScore(v) {
  if (v == null || v === '') return undefined;
  const n = Number(String(v).replace(/[^\d.]/g, ''));
  if (!Number.isFinite(n) || n <= 0) return undefined;
  // 100점/10점 척도로 적어 온 경우를 5점 척도로 환산
  if (n > 10) return Math.max(1, Math.min(5, Math.round((n / 100) * 5)));
  if (n > 5) return Math.max(1, Math.min(5, Math.round((n / 10) * 5)));
  return Math.max(1, Math.min(5, Math.round(n)));
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

  const wb = XLSX.readFile(abs, { cellDates: true });
  const sheetName = sheetArg || wb.SheetNames[0];
  if (!wb.Sheets[sheetName]) {
    console.error(`시트 '${sheetName}' 없음. 사용 가능: ${wb.SheetNames.join(', ')}`);
    process.exit(1);
  }
  console.log(`시트: ${sheetName}${wb.SheetNames.length > 1 ? ` (전체: ${wb.SheetNames.join(', ')})` : ''}`);

  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '', blankrows: false });
  if (!rows.length) {
    console.error('빈 시트입니다.');
    process.exit(1);
  }

  // 제목 행이 1행이 아닐 수 있다 — '출원번호'/'명칭' 류가 가장 많이 걸리는 행을 헤더로 본다
  let headerRow = 0;
  let bestHits = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const hits = buildHeaderMap(rows[i]).used.size;
    if (hits > bestHits) { bestHits = hits; headerRow = i; }
  }
  const headers = rows[headerRow];
  const { map, used } = buildHeaderMap(headers);
  console.log(`헤더 행: ${headerRow + 1}행 · 인식된 컬럼 ${Object.keys(map).length}개`);
  Object.entries(map).forEach(([f, i]) => console.log(`  ${f.padEnd(20)} ← "${headers[i]}"`));
  const unmapped = headers.map((h, i) => (h && !used.has(i) ? h : null)).filter(Boolean);
  if (unmapped.length) console.log(`  (미매핑 → extra 보관: ${unmapped.join(', ')})`);

  const get = (row, field) => (map[field] === undefined ? '' : row[map[field]]);

  const patents = rows
    .slice(headerRow + 1)
    .filter((row) => row.some((c) => String(c ?? '').trim() !== ''))
    .map((row, i) => {
      const appNo = String(get(row, 'applicationNumber') ?? '').trim();
      const regNo = String(get(row, 'registrationNumber') ?? '').trim();
      const evaluation = {
        technology: toScore(get(row, 'technology')),
        marketability: toScore(get(row, 'marketability')),
        rights: toScore(get(row, 'rights')),
        utilization: toScore(get(row, 'utilization')),
        comment: String(get(row, 'evalComment') ?? '').trim() || undefined,
      };
      const extra = {};
      headers.forEach((h, ci) => {
        if (h && !used.has(ci) && String(row[ci] ?? '').trim() !== '') extra[String(h).trim()] = row[ci];
      });

      return {
        id: `P${String(i + 1).padStart(3, '0')}`,
        title: String(get(row, 'title') ?? '').trim() || '(제목 없음)',
        titleEn: String(get(row, 'titleEn') ?? '').trim() || undefined,
        country: toCountry(get(row, 'country'), appNo),
        status: toStatus(get(row, 'status'), regNo),
        applicationNumber: appNo || null,
        applicationDate: toDate(get(row, 'applicationDate')),
        registrationNumber: regNo || null,
        registrationDate: toDate(get(row, 'registrationDate')),
        publicationNumber: String(get(row, 'publicationNumber') ?? '').trim() || undefined,
        publicationDate: toDate(get(row, 'publicationDate')) || undefined,
        expiryDate: toDate(get(row, 'expiryDate')) || undefined,
        applicant: String(get(row, 'applicant') ?? '').trim() || COMPANY,
        inventors: toList(get(row, 'inventors')),
        ipc: toCodeList(get(row, 'ipc')),
        family: String(get(row, 'family') ?? '').trim() || '미분류',
        category: String(get(row, 'category') ?? '').trim() || '미분류',
        productLine: toList(get(row, 'productLine')),
        abstract: String(get(row, 'abstract') ?? '').trim() || undefined,
        claims: Number(get(row, 'claims')) || undefined,
        notes: String(get(row, 'notes') ?? '').trim() || undefined,
        evaluation: Object.values(evaluation).some((v) => v !== undefined) ? evaluation : {},
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
    patents,
  };

  const byStatus = patents.reduce((m, p) => ({ ...m, [p.status]: (m[p.status] || 0) + 1 }), {});
  const byCountry = patents.reduce((m, p) => ({ ...m, [p.country]: (m[p.country] || 0) + 1 }), {});
  const evaluated = patents.filter((p) => Object.values(p.evaluation).some((v) => typeof v === 'number')).length;
  console.log(`\n총 ${patents.length}건`);
  console.log(`  상태: ${Object.entries(byStatus).map(([k, v]) => `${k} ${v}`).join(' · ')}`);
  console.log(`  국가: ${Object.entries(byCountry).map(([k, v]) => `${k} ${v}`).join(' · ')}`);
  console.log(`  평가 입력됨: ${evaluated}건`);

  const noTitle = patents.filter((p) => p.title === '(제목 없음)').length;
  const noDate = patents.filter((p) => !p.applicationDate).length;
  const noInventor = patents.filter((p) => !p.inventors.length).length;
  if (noTitle || noDate || noInventor) {
    console.log('\n확인 필요:');
    if (noTitle) console.log(`  · 명칭 없음 ${noTitle}건`);
    if (noDate) console.log(`  · 출원일 없음 ${noDate}건`);
    if (noInventor) console.log(`  · 발명자 없음 ${noInventor}건`);
  }

  if (dry) {
    console.log('\n--dry 모드 — 파일을 쓰지 않았습니다. 첫 2건 미리보기:');
    console.log(JSON.stringify(patents.slice(0, 2), null, 2));
    return;
  }

  fs.writeFileSync(OUT, JSON.stringify(out, null, 2) + '\n');
  console.log(`\n저장 완료 → ${path.relative(ROOT, OUT)}`);
}

main();
