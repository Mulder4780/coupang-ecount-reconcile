/* CSOS 고정 진입점 서비스 워커
 * ============================================================
 * 크롬이 [설치 및 바로가기 만들기]로 **진짜 앱 설치(WebAPK)** 를 해 주려면
 * fetch 핸들러를 가진 서비스 워커가 있어야 한다. 이게 없으면 메뉴를 눌러도
 * 단순 북마크만 생기거나 아무 일도 안 일어난다.
 *
 * ★ 캐시는 하지 않는다. 이 페이지는 '지금 살아 있는 접속 주소'를 읽어 넘기는 곳이라
 *   옛 주소를 캐시해 두면 폰이 죽은 주소로 계속 들어가게 된다.
 */
self.addEventListener('install', (e) => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', (e) => {
  // 그대로 통과시킨다(설치 요건만 충족). 네트워크가 죽으면 브라우저 기본 오류가 뜬다.
  e.respondWith(fetch(e.request));
});
