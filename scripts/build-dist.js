/* GitHub Pages 로 올릴 산출물 디렉터리를 만든다.
   리포 루트를 그대로 발행하면 package.json, server.js, scripts/ 까지 같이 서빙되므로
   대시보드 구동에 실제로 필요한 파일만 dist/ 로 추린다. */
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

// Pages 는 기본으로 Jekyll 을 태운다. 밑줄로 시작하는 파일이 조용히 빠지는 걸 막는다.
fs.writeFileSync(path.join(DIST, '.nojekyll'), '');

const count = (dir) =>
  fs.readdirSync(dir, { withFileTypes: true }).reduce(
    (n, d) => n + (d.isDirectory() ? count(path.join(dir, d.name)) : 1), 0);
console.log(`dist/ 생성 완료 — 파일 ${count(DIST)}개`);
