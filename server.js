/* 로컬 확인용 정적 서버.
   dist/ 를 서빙한다 — 소스 폴더를 그대로 열면 배포본과 경로가 어긋나
   로컬에선 되는데 배포하면 데이터가 안 뜨는 일이 생긴다.
   file:// 로 열면 fetch('data/rights.json') 이 CORS 로 막혀 빈 화면이 된다. */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const PORT = process.env.PORT || 5174;
const ROOT = path.join(__dirname, 'dist');
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

// 소스를 고치고 바로 띄우는 경우가 많으므로 매번 새로 빌드한다
execFileSync(process.execPath, [path.join(__dirname, 'scripts', 'build-dist.js')], { stdio: 'inherit' });

http
  .createServer((req, res) => {
    const urlPath = decodeURIComponent(req.url.split('?')[0]);
    let filePath = path.join(ROOT, urlPath);

    // 경로 탈출 차단 — ROOT 밖은 내보내지 않는다
    if (!path.resolve(filePath).startsWith(path.resolve(ROOT))) {
      res.writeHead(403).end('Forbidden');
      return;
    }
    if (fs.existsSync(filePath) && fs.statSync(filePath).isDirectory()) {
      filePath = path.join(filePath, 'index.html');
    }
    if (!fs.existsSync(filePath)) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('Not Found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(filePath)] || 'application/octet-stream',
      'Cache-Control': 'no-store',
    });
    fs.createReadStream(filePath).pipe(res);
  })
  .listen(PORT, () => console.log(`IP 대시보드 → http://localhost:${PORT}`));
