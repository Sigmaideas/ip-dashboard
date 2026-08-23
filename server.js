/* 로컬 확인용 정적 서버.
   file:// 로 열면 fetch('../data/patents.json') 이 CORS 로 막혀서 대시보드가 빈 화면이 된다. */
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 5174;
const ROOT = __dirname;
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

http
  .createServer((req, res) => {
    const urlPath = decodeURIComponent(req.url.split('?')[0]);

    // 루트에서 dashboard/index.html 을 그대로 서빙하면 style.css 같은 상대경로가
    // /style.css 로 풀려 404 가 난다. 경로를 옮겨 주고 브라우저가 다시 요청하게 한다
    if (urlPath === '/') {
      res.writeHead(302, { Location: '/dashboard/' }).end();
      return;
    }
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
