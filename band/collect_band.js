/* ============================================================================
 * collect_band.js — 밴드 게시글 수집기 (브라우저 콘솔용)
 * ============================================================================
 * 밴드 공식 API 앱이 아직 심사 중이라, 이미 로그인된 브라우저에서 화면을 스크롤해
 * 게시글을 모은다(사람이 직접 스크롤해 읽는 것과 같은 동작).
 *
 * ▶ 사용법
 *   1) 크롬에서 밴드를 연다:  https://www.band.us/band/90610953
 *   2) F12 → [Console] 탭
 *   3) 아래 전체를 복사해 붙여넣고 Enter
 *   4) "목표 날짜까지 도달" 메시지가 뜨고 JSON 파일이 자동 다운로드된다
 *   5) 다운로드된 dump_*.json 을  ecount/band/cache/  폴더에 넣는다
 *   6) 다른 밴드(84789192)도 같은 방법으로 반복
 *
 * ▶ 목표 날짜를 바꾸려면 아래 UNTIL 값을 수정 (기본: 2026-01-01)
 * ========================================================================== */
(async () => {
  const UNTIL = '2026-01-01';          // 이 날짜 이전까지 거슬러 올라가면 멈춘다
  const MAX_SCROLL = 4000;             // 안전 상한(무한 스크롤 방지)
  const PAUSE = 700;                   // 스크롤 간 대기(ms) — 밴드 서버 예의

  const bandId = (location.pathname.match(/\/band\/(\d+)/) || [])[1] || 'unknown';
  const bandName = (document.querySelector('.bandName, ._bandName, h1') || {}).innerText || '';
  const ABS = /(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/;

  // 진행 상황 표시 박스
  const box = document.createElement('div');
  box.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:99999;background:#0E1B3F;' +
    'color:#fff;padding:14px 18px;border-radius:12px;font:13px/1.6 sans-serif;' +
    'box-shadow:0 10px 30px rgba(0,0,0,.4);min-width:240px';
  document.body.appendChild(box);
  const say = h => { box.innerHTML = h; console.log(box.innerText); };

  const posts = {};
  let oldest = '';

  /* 화면에 보이는 게시글 카드를 긁는다.
     밴드는 클래스명이 자주 바뀌므로 클래스에 의존하지 않고,
     "글쓴 시각 문자열을 품은 카드"를 기준으로 찾는다. */
  function harvest() {
    let added = 0;
    document.querySelectorAll('a[href*="/post/"]').forEach(a => {
      const m = a.getAttribute('href').match(/\/post\/(\d+)/);
      if (!m) return;
      const id = m[1];
      if (posts[id]) return;
      // 게시글 본문을 담은 조상 요소를 찾는다(시각 문자열이 들어 있는 가장 가까운 블록)
      let el = a, card = null;
      for (let i = 0; i < 8 && el; i++, el = el.parentElement) {
        const t = el.innerText || '';
        if (ABS.test(t) && t.length > 60) { card = el; }
      }
      if (!card) return;
      const text = card.innerText.trim();
      const tm = text.match(/(\d{4}년\s*\d{1,2}월\s*\d{1,2}일\s*(?:오전|오후)?\s*\d{1,2}:\d{2})/);
      const author = (card.querySelector('.userName, ._authorName, strong') || {}).innerText || '';
      // 사진 URL까지 담는다(거래명세서·계산서 사진을 OCR로 읽기 위해)
      const imgs = [...card.querySelectorAll('img')]
        .map(i => i.currentSrc || i.src || '')
        .filter(u => /^https?:/.test(u) && !/profile|emoticon|sticker|icon/i.test(u));
      const photo = imgs.length;
      const cmt = (text.match(/댓글\s*(\d+)/) || [0, 0])[1];
      posts[id] = {
        author: (author || '').trim(),
        timeText: tm ? tm[1].replace(/\s+/g, ' ') : '',
        content: text,
        photo_count: photo,
        comment_count: Number(cmt) || 0,
        images: imgs.slice(0, 12)
      };
      added++;
      const d = text.match(ABS);
      if (d) {
        const iso = `${d[1]}-${String(d[2]).padStart(2, '0')}-${String(d[3]).padStart(2, '0')}`;
        if (!oldest || iso < oldest) oldest = iso;
      }
    });
    return added;
  }

  let stall = 0, lastH = 0;
  for (let i = 0; i < MAX_SCROLL; i++) {
    harvest();
    say(`밴드 <b>${bandId}</b> 수집 중<br>모은 글 <b>${Object.keys(posts).length}</b>개` +
        `<br>가장 오래된 글 <b>${oldest || '-'}</b><br>목표 ${UNTIL}`);
    if (oldest && oldest <= UNTIL) { say(`✅ 목표 날짜까지 도달<br>총 <b>${Object.keys(posts).length}</b>개`); break; }
    window.scrollTo(0, document.body.scrollHeight);
    await new Promise(r => setTimeout(r, PAUSE));
    const h = document.body.scrollHeight;
    stall = (h === lastH) ? stall + 1 : 0;
    lastH = h;
    if (stall >= 8) { say(`⛔ 더 이상 불러오지 않음(끝)<br>총 <b>${Object.keys(posts).length}</b>개 · 최고참 ${oldest}`); break; }
  }
  harvest();

  // 명세서·계산서로 보이는 글의 사진은 **본문(base64)까지** 담아 바로 OCR에 넘긴다.
  // (밴드 이미지 URL은 로그인 쿠키가 있어야 열리므로, 여기서 받아 두는 게 확실하다)
  const DOC = /명세서|계산서|견적|청구|세금/;
  const targets = Object.entries(posts).filter(([,p]) => DOC.test(p.content || ''));
  say(`💾 문서 사진 내려받는 중… 대상 글 ${targets.length}개`);
  let got = 0;
  for (const [id, p] of targets) {
    for (const url of (p.images || []).slice(0, 4)) {
      try {
        const b = await (await fetch(url, { credentials: 'include' })).blob();
        if (b.size > 3 * 1024 * 1024) continue;                 // 3MB 초과는 건너뜀
        p.imageData = p.imageData || [];
        p.imageData.push(await new Promise(r => {
          const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b);
        }));
        got++;
      } catch (e) { /* 실패한 사진은 URL만 남긴다 */ }
    }
    say(`💾 문서 사진 ${got}장 확보 (${targets.length}개 글 중 진행)`);
  }

  const dump = { band: bandId, name: bandName.trim(), capturedAt: Date.now(), posts };
  const blob = new Blob([JSON.stringify(dump, null, 1)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `dump_${bandId}.json`;
  document.body.appendChild(a); a.click(); a.remove();
  say(`💾 저장됨 <b>dump_${bandId}.json</b><br>글 ${Object.keys(posts).length}개 · ${oldest} ~<br>` +
      `이 파일을 ecount/band/cache/ 에 넣어주세요`);
})();
