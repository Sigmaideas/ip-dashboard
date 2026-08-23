/* 발행 산출물 디렉터리를 만든다.
   사이트 루트가 곧 대시보드가 되도록 dashboard/ 의 내용물을 dist/ 바로 아래로 편다.
   리포 루트를 그대로 발행하면 package.json, server.js, scripts/ 까지 같이 서빙된다. */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const DIST = path.join(ROOT, 'dist');

fs.rmSync(DIST, { recursive: true, force: true });
fs.mkdirSync(DIST, { recursive: true });

// dashboard/* → dist/*  (index.html, app.js, style.css, logo.png)
fs.cpSync(path.join(ROOT, 'dashboard'), DIST, { recursive: true });
// data/ 는 그대로 한 단계 아래. app.js 가 'data/rights.json' 으로 읽는다
fs.cpSync(path.join(ROOT, 'data'), path.join(DIST, 'data'), { recursive: true });

// Pages 는 기본으로 Jekyll 을 태운다. 밑줄로 시작하는 파일이 조용히 빠지는 걸 막는다.
fs.writeFileSync(path.join(DIST, '.nojekyll'), '');

const count = (dir) =>
  fs.readdirSync(dir, { withFileTypes: true }).reduce(
    (n, d) => n + (d.isDirectory() ? count(path.join(dir, d.name)) : 1), 0);
console.log(`dist/ 생성 완료 — 파일 ${count(DIST)}개`);
