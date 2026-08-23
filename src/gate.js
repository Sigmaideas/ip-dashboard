/* Pages 배포본 앞단의 비밀번호 게이트.
   관리대장은 내부 문서라 pages.dev URL 을 아는 것만으로 열려서는 안 된다.
   HTTP Basic 인증으로 막고, 통과한 요청만 정적 자산으로 넘긴다. */

const REALM = '엑스와이지 IP 대시보드';
const USERNAME = 'xyz';

/* 비교 시간이 입력에 따라 달라지면 비밀번호를 한 글자씩 맞춰 볼 수 있다.
   양쪽을 SHA-256 으로 같은 길이로 만든 뒤 상수 시간 비교한다. */
async function matches(actual, expected) {
  const enc = new TextEncoder();
  const [a, b] = await Promise.all([
    crypto.subtle.digest('SHA-256', enc.encode(actual)),
    crypto.subtle.digest('SHA-256', enc.encode(expected)),
  ]);
  return crypto.subtle.timingSafeEqual(new Uint8Array(a), new Uint8Array(b));
}

function unauthorized() {
  return new Response('인증이 필요합니다.', {
    status: 401,
    headers: {
      // charset=UTF-8 이 없으면 일부 브라우저가 비ASCII 비밀번호를 latin-1 로 보낸다
      'WWW-Authenticate': `Basic realm="${REALM}", charset="UTF-8"`,
      'Content-Type': 'text/plain; charset=utf-8',
      'Cache-Control': 'no-store',
    },
  });
}

async function authorized(request, password) {
  const header = request.headers.get('Authorization') || '';
  const [scheme, encoded] = header.split(' ');
  if (scheme !== 'Basic' || !encoded) return false;

  let decoded;
  try {
    // atob 는 바이트만 돌려주므로 UTF-8 로 다시 해석해야 한글 비밀번호가 깨지지 않는다
    const bytes = Uint8Array.from(atob(encoded), (c) => c.charCodeAt(0));
    decoded = new TextDecoder().decode(bytes);
  } catch {
    return false;
  }

  // 비밀번호에 콜론이 들어갈 수 있으므로 첫 콜론에서만 자른다
  const sep = decoded.indexOf(':');
  if (sep === -1) return false;
  const user = decoded.slice(0, sep);
  const pass = decoded.slice(sep + 1);

  const [userOk, passOk] = await Promise.all([
    matches(user, USERNAME),
    matches(pass, password),
  ]);
  return userOk && passOk;
}

export default {
  async fetch(request, env) {
    // 비밀번호가 설정되지 않은 채 배포되면 무방비로 열린다. 열어 주느니 막는다.
    if (!env.SITE_PASSWORD) {
      return new Response('SITE_PASSWORD 가 설정되지 않아 접근을 차단했습니다.', {
        status: 503,
        headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' },
      });
    }

    if (!(await authorized(request, env.SITE_PASSWORD))) return unauthorized();

    const res = await env.ASSETS.fetch(request);
    // 인증 뒤의 응답이 공용 캐시에 남지 않도록 한다
    const headers = new Headers(res.headers);
    headers.set('Cache-Control', 'no-store');
    return new Response(res.body, { status: res.status, statusText: res.statusText, headers });
  },
};
