/* grab_posts.js — 밴드 게시글 "상세 페이지" 수집기 (브라우저 콘솔/자동화 주입용)
 *
 * 왜 이 방식인가 (2026-08-04 확정, AGENTS.md 현재상태 3차 참조)
 *  · 밴드 피드는 무한스크롤이 10글에서 멈춘다 — 점진 스크롤도 실패한다.
 *  · 상세 페이지는 '…더보기' 없이 전문이 나와 피드 긁기보다 품질이 좋다.
 *  → 숨은 iframe 으로 /band/<밴드>/post/<번호> 를 하나씩 열어 본문을 뜯는다.
 *
 * 쓰는 법 (밴드 탭에서 이 파일 내용을 붙여 넣은 뒤)
 *    __grabStart(90610953, __grabRange(5360, 5424))   // 시작(비동기, 즉시 반환)
 *    __grabStatus()                                   // 진행률 확인 (폴링)
 *    __grabSave()                                     // dump_<시각>_<밴드>.json 다운로드
 *  다운로드 파일은 download_intake.py --apply 가 Z: 밴드 덤프로 흡수하고
 *  band/convert_dump.py 가 캐시로 합친다(수정글 감지 포함). 손으로 옮기지 말 것.
 *
 * 진행 상태는 window.__GRAB 에 남는다 — 탭이 살아 있는 한 이어받을 수 있고,
 * 끊기면 recheck_plan.py 가 캐시를 보고 남은 번호를 다시 뽑아 준다.
 *
 * ★ 숨은 탭 타이머 throttling (2026-08-05 실측 — 15분에 0건)
 *   밴드 탭이 **뒤에 있으면** 크롬이 setTimeout 을 1초→1분 간격으로 늦춘다
 *   (5분 넘게 숨어 있으면 intensive throttling). 400ms 폴링 30회가 30분이 되어
 *   첫 글에서 멈춘 것처럼 보였다. 그래서 이 파일은 **페이지 타이머를 쓰지 않는다**:
 *     · 지연 → Web Worker 타이머(워커는 throttling 대상이 아니다)
 *     · 본문 등장 대기 → MutationObserver(이벤트는 늦춰지지 않는다)
 *   앞으로도 `setTimeout` 을 여기에 다시 넣지 말 것. 탭을 앞에 둘 필요도 없다.
 */
(function () {
  const S = (window.__GRAB = window.__GRAB || {
    band: null, posts: {}, done: [], missing: [], failed: [],
    // notime: 번호 → 그때 껍데기에 잡힌 본문 지문. '아직 없는 글' 판정의 재료다
    // (판정은 convert_dump 가 한다 — 브라우저는 사실만 적는다).
    notime: {},
    running: false, startedAt: null, total: 0, stop: false,
  });
  if (!S.notime) S.notime = {};        // 옛 탭에서 이어받을 때도 자리가 있게

  // 워커 타이머 — 숨은 탭에서도 제 시각에 온다.
  const W = new Worker(URL.createObjectURL(new Blob(
    ['onmessage=e=>{setTimeout(()=>postMessage(e.data.id),e.data.ms)}'],
    { type: 'text/javascript' })));
  const waiters = {};
  let wid = 0;
  W.onmessage = (e) => { const f = waiters[e.data]; if (f) { delete waiters[e.data]; f(); } };
  const sleep = (ms) => new Promise((r) => { const id = ++wid; waiters[id] = r; W.postMessage({ id, ms }); });
  // 밤새 이어 도는 주행기(grab_all.js)도 **같은 워커 타이머**를 써야 한다.
  // 페이지 타이머로 기다리면 탭이 뒤에 있을 때 1분까지 늦춰져 배치 사이가 멎는다.
  window.__grabSleep = sleep;

  window.__grabRange = (from, to) => {
    const out = [];
    for (let n = Number(from); n <= Number(to); n++) out.push(n);
    return out;
  };

  const txt = (root, sel) => {
    for (const s of sel.split(',')) {
      const e = root.querySelector(s.trim());
      if (e && (e.innerText || '').trim()) return e.innerText.trim();
    }
    return '';
  };

  // ── 댓글 본문 (2026-08-08) ────────────────────────────────────────────────
  // 접수 취소 통보는 대부분 **글이 아니라 댓글**로 온다("작동 원활함. 접수 취소 하세요").
  // 그동안 캐시에는 `comment_count` 숫자만 있어 cancel_watch 가 반쪽으로 돌았다.
  // 원칙은 본문과 같다 — **시각 없는 수확은 버린다.** 밴드는 아직 안 그려진 자리에도
  // 빈 껍데기를 두므로 시각이 없으면 직전 화면이 묻어 온 것일 수 있다.
  // 그리고 **몇 개를 봤는지 같이 적는다**: 접힌 댓글을 다 펴지 못했을 때
  // "댓글이 없다"와 "못 읽었다"를 읽는 쪽이 갈라야 한다(못 가르면 조용한 사고가 된다).
  const CMT_ITEM = 'ul.commentList > li, .commentList .commentItem, .cComment li, [class*="commentItem"]';
  const CMT_BODY = '.commentText, ._commentContent, .txt, [class*="commentText"]';
  const CMT_TIME = '.time, .date, [class*="time"]';
  const CMT_WHO = '.uName, .commentWriterInfo .text, .writeInfo .name, [class*="writerName"]';
  function readComments(root) {
    let items = [];
    for (const s of CMT_ITEM.split(',')) {
      items = [...root.querySelectorAll(s.trim())];
      if (items.length) break;
    }
    const out = [];
    for (const it of items) {
      // 둘 다 있어야 담는다 — 이 조건이 엉뚱한 <li> 를 댓글로 오해하는 것도 같이 막는다.
      const content = txt(it, CMT_BODY);
      const timeText = txt(it, CMT_TIME);
      if (!content || !timeText) continue;
      out.push({ author: txt(it, CMT_WHO), timeText, content: content.slice(0, 2000) });
    }
    return out;
  }

  // 상세 페이지 한 글을 iframe 으로 열어 본문·글쓴이·시각·사진/댓글 수를 뜯는다.
  // ★ 없는 글을 열면 밴드가 `alert('삭제되었거나 찾을 수 없습니다.')` 를 띄운다
  //   (2026-08-06 실측). 모달은 **탭 전체를 멈춘다** — 사람이 누를 때까지 수집이 선다.
  //   과거글 구간은 지운 글이 섞여 있어 수백 번 뜰 수 있다. iframe 은 같은 출처라
  //   그 안의 alert/confirm/prompt 를 조용한 함수로 갈아끼울 수 있다.
  function muteDialogs(f) {
    try {
      const w = f.contentWindow;
      if (!w) return;
      w.alert = () => {};
      w.confirm = () => true;
      w.prompt = () => null;
    } catch (e) { /* 출처가 다르면 손댈 수 없다 — 그때는 사람이 눌러야 한다 */ }
  }

  async function grabOne(band, no, waitMs, bodyMs) {
    const f = document.createElement('iframe');
    f.style.cssText = 'position:fixed;left:-9999px;top:0;width:1200px;height:900px';
    // src 를 넣기 전에 한 번, 로드된 뒤에 또 한 번 막는다. 문서가 바뀌면 window 의
    // alert 이 되살아나므로 한쪽만으로는 새는 경우가 있다.
    f.addEventListener('load', () => muteDialogs(f));
    f.src = `https://www.band.us/band/${band}/post/${no}`;
    document.body.appendChild(f);
    muteDialogs(f);
    try {
      await Promise.race([new Promise((r) => { f.onload = r; }), sleep(waitMs)]);
      muteDialogs(f);
      // 본문은 SPA 가 나중에 그린다 — 폴링 대신 '그려지는 순간'을 관찰한다.
      const d = f.contentDocument;
      if (!d) return { status: 'fail' };
      // ★ 주소부터 본다 (2026-08-07 3차). 없는 글(삭제됐거나 아직 안 생긴 번호)을 열면
      //   밴드는 오류 대신 **피드로 되돌린다.** 그 피드가 끝까지 그려지면 본문도 시각도
      //   멀쩡히 있어서(실측: "6시간 전"이 잡혔다) 시각 가드([130])까지 뚫고
      //   **날짜 달린 가짜**가 들어올 수 있다. 내가 연 주소에 내가 있는지가 유일한 확증이다.
      //   같은 출처(band.us)라 pathname 은 읽을 수 있다.
      let path = '';
      try { path = f.contentWindow.location.pathname || ''; } catch (e) { /* 출처 다름 — 아래 판정으로 */ }
      if (path && !path.endsWith('/post/' + no)) {
        const sig = txt(d, '.postText, .dPostTextView').replace(/\s+/g, ' ').slice(0, 200);
        // 묘비(missing)가 아니라 fail 이다 — '아직 안 생긴 번호'일 수 있고, 그 번호는
        // 내일 진짜로 생긴다. 없음의 증거는 지문 합의(convert_dump)가 맡는다.
        return { status: 'fail', reason: 'redirect', sig };
      }
      if (!d.querySelector('.postText')) {
        await new Promise((done) => {
          let mo = null, fin = false;
          const end = () => { if (!fin) { fin = true; if (mo) mo.disconnect(); done(); } };
          mo = new MutationObserver(() => { if (d.querySelector('.postText')) end(); });
          mo.observe(d.documentElement || d, { childList: true, subtree: true });
          sleep(bodyMs).then(end);            // 끝내 안 그려져도 반드시 빠져나온다
        });
      }
      // 삭제·권한 없는 글은 본문 자체가 없다 — '없는 글'로 구분해 남긴다.
      //
      // ★ 밴드는 지운 글을 열어도 '삭제됨'이라고 말해 주지 않는다(2026-08-05 실측).
      //   삭제 글과 정상 글의 DOM 을 나란히 떠 봤더니 **`.postText` 하나만 다르고**
      //   나머지(.postWrap·.postMain·body class)는 완전히 같았다. 그래서 안내문이나
      //   껍데기 표식으로는 구분할 수 없다. 구분 기준은 이것뿐이다:
      //     앱이 글 영역을 **다 그렸는데도**(.postMain/.postWrap 존재)
      //     대기 시간이 끝날 때까지 본문이 없다 → 없는 글.
      //   느려서 아직 안 그려진 경우와는 MutationObserver 로 bodyMs 를 꽉 채워 기다린
      //   뒤에 판정하므로 갈린다. 이 판정이 없으면 recheck_plan 이 같은 번호를
      //   영원히 다시 뽑는다(실제로 4건이 하루 6회차를 전부 실패했다).
      if (!d.querySelector('.postText')) {
        const body = (d.body && d.body.innerText) || '';
        if (/삭제|없는 게시글|권한|찾을 수 없/.test(body)) return { status: 'missing' };
        const rendered = !!d.querySelector('.postMain, .postWrap');
        return { status: rendered ? 'missing' : 'fail' };
      }
      const main = d.querySelector('.postMain') || d.querySelector('.postWrap') || d;
      // ★ 본문이 그려졌다고 **머리말까지 그려진 것은 아니다** (2026-08-07 실측).
      //   밴드는 본문(.postText)을 먼저 칠하고 글쓴이·작성시각(.time)을 조금 뒤에 채운다.
      //   `.postText` 를 보자마자 가져가면 날짜가 빈 채로 저장되고, 그 글은 **날짜가
      //   없어서 어떤 작업과도 대조되지 않는다** — 본문은 멀쩡한데 쓸 수가 없다.
      //   실제로 그렇게 모은 것이 621건이었다(90610953 523 · 84789192 98).
      //   그중엔 돌발유료 A/S 안내·쿠팡 PO 알림처럼 꼭 필요한 글이 섞여 있었다.
      //   본문 없음과 달리 여기서는 **짧게 더 기다린다** — 지운 글 판정과 무관하다.
      for (let i = 0; i < 12 && !txt(main, '.postListInfoWrap .time, .time'); i++) {
        await sleep(250);
      }
      const timeText = txt(main, '.postListInfoWrap .time, .time');
      // ★ 시각이 끝내 안 붙으면 **수확을 버린다** (2026-08-07 실사고 2차).
      //   밴드는 **아직 없는 글 번호**(최신 글보다 큰 번호)에도 200 을 주고 앱 껍데기를
      //   그린다. 그 화면에서 본문을 뜯으면 직전 화면(피드 맨 위 글)의 본문이 그대로
      //   잡히고 글쓴이·시각만 빈다. 이날 3539~3578 을 그렇게 모아 **40건이 전부 같은
      //   글**이었는데 status 가 ok 라 그대로 캐시에 들어갔다(3308→3348).
      //   본문이 있으니 아무도 실패인 줄 몰랐다 — 제일 나쁜 종류의 실패다.
      //   위(본문 없음)와 달리 'missing' 이 아니라 'fail' 로 돌린다:
      //   3539 는 **내일이면 진짜로 생긴다.** 묘비를 세우면 그때 영영 못 모은다.
      //   시각 없는 글은 어차피 어떤 작업과도 대조되지 않아 저장할 값이 없다.
      //   ★ 그러면서 **본문 지문은 남긴다** (2026-08-07 3차). 아직 없는 글이면 껍데기에
      //   직전 화면 본문이 그대로 잡히므로, 이런 번호끼리 지문이 **똑같다**. 그 사실이
      //   "이 번호들은 아직 없다"의 증거가 된다 — 판정은 여기서 하지 않고
      //   convert_dump 가 한다(브라우저는 사실만 적고, 판단은 시험할 수 있는 곳에서).
      if (!timeText) {
        const sig = txt(main, '.postText, .dPostTextView').replace(/\s+/g, ' ').slice(0, 200);
        return { status: 'fail', reason: 'no-time', sig };
      }
      const imgs = [...main.querySelectorAll('img')]
        .map((i) => i.src).filter((s) => /pstatic|phinf/.test(s));
      // ★ 개수는 **고정 선택자로 찾지 않는다** (2026-08-09 실사고).
      //   `._commentCount, .comment .count, .uComment .count` 로 찾던 것이 지금 밴드
      //   화면과 안 맞아 250건을 실패 0 으로 긁고도 **댓글이 한 건도 안 들어왔다.**
      //   개수를 0 으로 읽으니 댓글이 그려질 때까지 기다리지 않고 빈 배열을 담았다.
      //   선택자를 다시 맞춰 봐야 밴드가 화면을 고치면 또 같은 자리에서 깨진다.
      //   그래서 ① 선택자 → ② **글자 모양('댓글 12')** 순으로 찾고,
      //   ③ 그래도 모르면 **모른다고 적는다**(아래 countKnown).
      let cmt = (txt(main, '._commentCount, .comment .count, .uComment .count')
        .match(/\d+/) || [''])[0];
      let countKnown = cmt !== '';
      if (!countKnown) {
        for (const el of main.querySelectorAll('a, button, span, div')) {
          const t = (el.textContent || '').trim();
          if (t.length > 12) continue;                 // 본문을 훑지 않는다
          const m = t.match(/^댓글\s*([0-9,]+)$/);
          if (m) { cmt = m[1].replace(/,/g, ''); countKnown = true; break; }
        }
      }
      // 댓글은 본문보다 늦게 그려진다 — 있다고 적힌 만큼 보일 때까지 잠깐만 더 기다린다.
      // 무한정 기다리지 않는다(접힌 댓글은 끝내 안 펴질 수 있고, 그건 아래 comments_full 이 말한다).
      const want = parseInt(cmt || '0', 10) || 0;
      let cts = readComments(d);
      // 개수를 모를 때도 **한 번은 기다려 본다** — 0 으로 단정하고 지나가면 그 글은
      // '읽었는데 댓글이 없더라'로 굳는다.
      const tries = countKnown ? 8 : 4;
      for (let i = 0; i < tries && cts.length < (countKnown ? want : 1); i++) {
        await sleep(250);
        cts = readComments(d);
      }
      const post = {
        author: txt(main, '.postWriterInfoWrap .text, .postWriter .text, .uProfileText'),
        timeText,
        content: txt(main, '.postText, .dPostTextView'),
        photo_count: String(imgs.length),
        images: [...new Set(imgs)],
      };
      // ★ **확인 못 한 것은 적지 않는다.** 개수도 모르고 댓글도 못 봤으면 그것은
      //   '댓글이 없다'가 아니라 '못 읽었다'다. 그런데 `comments` 키를 달아 두면
      //   그 글은 사각지대 계기·수집 목록에서 **읽은 글로 세어져 영영 다시 안 뽑힌다.**
      //   못 읽은 글이 읽은 글로 둔갑하는 것이 이 사고의 본체였다.
      if (countKnown || cts.length) {
        post.comment_count = cmt || '0';
        post.comments = cts;
        // 적힌 개수만큼 실제로 읽었나. false 면 읽는 쪽은 '댓글 없음'으로 단정하지 않는다.
        post.comments_full = cts.length >= want;
      } else {
        post.comments_unverified = true;   // 다음 회차가 이 글을 다시 뽑는다
      }
      return { status: 'ok', post };
    } catch (e) {
      return { status: 'fail', error: String(e) };
    } finally {
      f.remove();
    }
  }

  // 시작하면 곧바로 반환한다(자동화 호출이 타임아웃으로 죽지 않게) — 진행은 __grabStatus 로 본다.
  // ★ 한 배치는 250건까지 (2026-08-06 실측). 250건 배치는 다섯 회차를 문제없이 돌았는데
  //   500건으로 올리자 **탭 렌더러가 얼어** CDP 호출이 45초 타임아웃으로 죽었다.
  //   그 순간 수집분은 탭 메모리에만 있어 새로고침하면 통째로 날아간다.
  //   많이 남았으면 250씩 나눠서 여러 번 돌린다 — 한 번에 밀어 넣지 않는다.
  const BATCH_MAX = 250;

  window.__grabStart = function (band, nos, opt) {
    opt = opt || {};
    if (nos.length > BATCH_MAX) {
      return `한 배치는 ${BATCH_MAX}건까지다(탭이 언다). ${nos.length}건을 나눠서 걸어라.`;
    }
    if (S.running) return '이미 실행 중 — __grabStatus() 로 보라';
    // ★ 탭이 안 보이면 시작하지 않는다 (2026-08-07 실사고).
    //   밴드는 SPA 라 본문을 requestAnimationFrame 으로 그린다. 그런데 rAF 는
    //   `document.hidden` 인 탭에서 **한 번도 안 불린다**(타이머와 달리 Worker 로도
    //   못 우회한다). 그래서 창을 뒤로 넘긴 채 돌리면 iframe 은 영원히
    //   "로딩 중입니다" 이고, 아래 판정은 그것을 **없는 글로 오해한다.**
    //   이날 503건을 그렇게 갈아 넣었다 — ok 9 · fail 61. 확실히 살아 있는 500번을
    //   같은 방법으로 열어 보고서야 알았다(그것도 "로딩 중입니다" 였다).
    //   조용히 실패하느니 시작을 거절하는 편이 낫다. 캐시에 가짜 묘비가 쌓이면
    //   recheck_plan 이 그 번호를 영영 다시 안 뽑는다.
    if (typeof document !== 'undefined' && document.hidden) {
      return '탭이 뒤에 있다 — 이 창을 **앞으로 꺼내 놓고** 다시 걸어라. '
           + '(밴드 본문은 보이는 탭에서만 그려진다. 지금 돌리면 전부 실패로 기록된다)';
    }
    Object.assign(S, {
      band: String(band), running: true, stop: false, startedAt: Date.now(),
      total: nos.length, posts: opt.keep === false ? {} : S.posts,
      done: [], missing: [], failed: [],
    });
    (async () => {
      for (const no of nos) {
        if (S.stop) break;                       // __grabStop() 으로 중간에 끊을 수 있다
        // 돌던 중에 창이 뒤로 넘어가면 **실패로 기록하지 말고 기다린다.**
        // 사람이 다른 창을 보다 돌아오는 일은 밤샘 수집에서 늘 생긴다.
        while (typeof document !== 'undefined' && document.hidden && !S.stop) {
          S.paused = true;
          await sleep(5000);
        }
        S.paused = false;
        if (S.stop) break;
        const r = await grabOne(band, no, opt.waitMs || 9000, opt.bodyMs || 12000);
        if (r.status === 'ok') { S.posts[no] = r.post; S.done.push(no); }
        else if (r.status === 'missing') S.missing.push(no);
        else {
          S.failed.push(no);
          if ((r.reason === 'no-time' || r.reason === 'redirect') && r.sig) S.notime[no] = r.sig;
        }
        await sleep(opt.gapMs || 300);
      }
      S.running = false;
      S.finishedAt = Date.now();
    })();
    return `시작: ${nos.length}건 (밴드 ${band})`;
  };

  // 잘못 건 배치를 탭 새로고침 없이 끊는다(새로고침하면 모은 것이 날아간다).
  window.__grabStop = () => { S.stop = true; return '다음 글에서 멈춘다 — __grabSave() 로 저장하라'; };

  window.__grabStatus = () => ({
    running: S.running, paused: !!S.paused, total: S.total,
    ok: S.done.length, missing: S.missing.length, failed: S.failed.length,
    posts: Object.keys(S.posts).length,
    sec: S.startedAt ? Math.round((Date.now() - S.startedAt) / 1000) : 0,
    last: S.done[S.done.length - 1] || null,
  });

  // 덤프 파일 이름은 **맨 뒤 숫자 덩어리가 밴드번호**여야 한다(convert_dump 규칙).
  window.__grabSave = function () {
    const d = new Date(), p = (n) => String(n).padStart(2, '0');
    const stamp = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}${p(d.getHours())}${p(d.getMinutes())}`;
    const doc = {
      band: S.band, name: document.title.split('|')[0].replace('게시글 :', '').trim(),
      capturedAt: Date.now(), posts: S.posts,
      missing: S.missing, failed: S.failed, notime: S.notime,
    };
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(doc)], { type: 'application/json' }));
    a.download = `dump_${stamp}_${S.band}.json`;
    document.body.appendChild(a); a.click(); a.remove();
    return `${Object.keys(S.posts).length}건 저장 → ${a.download} (download_intake 가 흡수한다)`;
  };

  return 'grab_posts.js 준비됨 — __grabStart(밴드, __grabRange(from,to))';
})();
