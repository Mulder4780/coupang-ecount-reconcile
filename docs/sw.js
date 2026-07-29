/* CSOS 고정 진입점 서비스 워커
 * ============================================================
 * 두 가지 일을 한다.
 *  1) 크롬이 [설치 및 바로가기 만들기]로 **진짜 앱 설치(WebAPK)** 를 해 주게 한다.
 *     fetch 핸들러가 있어야 설치를 제안한다 — 없으면 단순 북마크만 생긴다.
 *  2) 오프라인 앱과 잠긴 사본을 손에 쥐고 있는다. 지하주차장·엘리베이터처럼
 *     신호가 끊기는 곳에서도 열려야 하기 때문이다.
 */
/* ★ 이 이름을 바꾸면 activate 가 옛 캐시를 지운다. 페이지를 고칠 때마다 올릴 것 —
   안 올리면 옛 사본을 쥔 기기가 언제 새 것을 받을지 알 수 없다(2026-07-28 실사고:
   옛 app.html 이 창을 터널 주소로 옮겨 재부팅마다 죽은 주소가 떴다). */
const CACHE = 'csos-v12-oh-owner-scope-2026-only';
const HOLD = ['index.html', 'app.html', 'data.enc', 'manifest.json', 'icon.svg', 'icon-180.png'];

self.addEventListener('install', (e) => {
  // 미리 받아 둔다. 하나쯤 실패해도 설치는 계속한다(아이콘 하나 때문에 앱이 죽으면 안 된다).
  e.waitUntil(caches.open(CACHE)
    .then((c) => Promise.allSettled(HOLD.map((u) => c.add(u))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;

  // ★ endpoint.json 은 절대 캐시하지 않는다. 지금 살아 있는 PC 주소를 담고 있어서
  //   옛 값을 쥐고 있으면 폰이 죽은 터널로 계속 들어간다(실제로 겪은 증상).
  if (url.pathname.endsWith('endpoint.json')) {
    e.respondWith(fetch(e.request).catch(() => new Response('{}', {
      headers: { 'Content-Type': 'application/json' } })));
    return;
  }

  // 나머지는 **새 것 우선, 안 되면 손에 쥔 것**. 신호가 있으면 늘 최신 사본을 보고,
  // 끊기면 마지막으로 받아 둔 사본으로 계속 일한다.
  e.respondWith(
    fetch(e.request)
      .then((r) => {
        if (r && r.ok) {
          const copy = r.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
        }
        return r;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true })
        .then((hit) => hit || Response.error())));
});
