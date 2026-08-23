/* Cloudflare Pages 직접 업로드용 산출물 디렉터리를 만든다.
   프로젝트 루트를 그대로 올리면 node_modules 와 inbox/ 의 원본 엑셀까지 딸려가므로
   대시보드 구동에 필요한 파일만 dist/ 로 추린다. */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const DIST = path.join(ROOT, 'dist');
const INCLUDE = ['index.html', 'dashboard', 'data'];

fs.rmSync(DIST, { recursive: true, force: true });
fs.mkdirSync(DIST, { recursive: true });

for (const entry of INCLUDE) {
  const src = path.join(ROOT, entry);
  if (!fs.existsSync(src)) throw new Error(`빌드 대상이 없습니다: ${entry}`);
  fs.cpSync(src, path.join(DIST, entry), { recursive: true });
}

// data/*.json 은 매번 최신을 받아야 한다. Pages 기본 캐시를 끈다.
fs.writeFileSync(
  path.join(DIST, '_headers'),
  '/data/*\n  Cache-Control: no-store\n'
);

// 비밀번호 게이트. dist/_worker.js 가 있으면 Pages 는 모든 요청을 먼저 여기로 보낸다.
fs.copyFileSync(path.join(ROOT, 'src', 'gate.js'), path.join(DIST, '_worker.js'));

const count = (dir) =>
  fs.readdirSync(dir, { withFileTypes: true }).reduce(
    (n, d) => n + (d.isDirectory() ? count(path.join(dir, d.name)) : 1), 0);
console.log(`dist/ 생성 완료 — 파일 ${count(DIST)}개`);
