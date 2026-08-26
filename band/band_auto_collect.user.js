// ==UserScript==
// @name         쿠팡업무 — 밴드 크롬 전용 수집
// @namespace    coupang-ecount
// @version      2.0
// @description  로그인된 크롬 밴드 탭이 수집계획을 받아 스스로 긁는다. 돌았는지 안 돌았는지를 PC 에 되보고한다.
// @match        https://www.band.us/band/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==
//
// 왜 이게 있나 (2026-08-13 지시: "앞으로 크롬에서만 긁어오는 알고리즘 만들어서 적용해")
//
//   밴드는 조회 API 가 없다. 글·댓글은 **로그인된 브라우저 DOM 안에서만** 읽힌다.
//   그 DOM 에서 JS 를 돌릴 수 있는 길은 셋이었고 셋 다 사람이나 크레딧을 먹었다:
//     ① Claude 가 브라우저 도구로 주입 — 매번 크레딧을 쓴다
//     ② 붙여넣기 파일 — 사람이 콘솔에 붙여넣어야 한다
//     ③ 이 유저스크립트 — **한 번 설치하면 그 뒤로 손이 안 간다**
//   ③ 하나로 모으는 것이 이 파일이다. ②는 설치 전·실패 때의 폴백으로만 남는다.
//
// ★ v2 에서 고친 것 — 이 파일은 2026-08-09 에 만들어졌는데 2026-08-13 까지
//   **한 번도 안 돌았다.** Tampermonkey 가 안 깔려 있었기 때문인데, 그 사실을
//   말해 주는 화면이 어디에도 없었다. 나흘 동안 "자동 수집이 있다"고 적힌 채
//   아무 일도 안 일어났고 아무도 몰랐다 — 이 프로젝트가 반복해 당한
//   '실패가 성공처럼 보이는 자리'다.
//   그래서 v2 의 핵심은 긁는 기능이 아니라 **되보고**다:
//     · 긁었으면 긁었다고, 못 긁었으면 **왜** 못 긁었는지 PC 에 알린다.
//     · 소식이 끊기면 `band/userscript_watch.py` 가 인계 문서에 올린다.
//     · 그래서 "설치했는데 안 도는" 상태가 다시는 조용히 지나가지 않는다.
//
// 설치
//   1. 크롬에 Tampermonkey 설치 (한 번만)
//   2. 앱 [실행] 탭 → '크롬 전용 수집' 카드 → 설치 링크
//      (또는 http://127.0.0.1:8899/band_auto_collect.user.js 를 열면 Tampermonkey 가 받는다)
//   3. 로그인된 밴드 탭을 **보이는 창**에 열어 둔다
//
// ⚠ 창이 다른 창에 완전히 가려지면 크롬은 **모든 탭을 숨은 것으로** 표시한다.
//   밴드 본문은 숨은 탭에서 안 그려지므로 수집기가 스스로 멈춘다(가짜 실패 방지).
//   탭만 열어 두는 것으로는 부족하고 **창이 보여야** 한다 — 2026-08-13 실측.
//   그때는 조용히 넘어가지 않고 `hidden` 으로 되보고한다. 그래야 사람이
//   "왜 안 돌지"가 아니라 "창을 앞으로 꺼내면 되는구나"를 알 수 있다.

(function () {
  'use strict';

  var VER = '2.0';
  var APP_CANDIDATES = ['http://127.0.0.1:8899', 'http://localhost:8899'];
  // PC 가 꺼져 있으면 로컬 앱이 없다 — 그때는 게시 사본에서 계획·수집기를 받는다.
  var PAGES_BASE = 'https://mulder4780.github.io/coupang-ecount-reconcile/collect';
  // -- 간격: **할 일이 있으면 빨리, 없으면 쉰다** (2026-08-26 지시)
  //   형님 지시: "전부 긁고 다음부터 바로바로 긁어올 수 있게 알고리즘 구현해".
  //   * 예전에는 갈래를 안 보고 **언제나 3시간**이었다.  그래서 새 글이 올라와도
  //     최대 3시간 뒤에 들어왔고, 오염 602건은 한 배치(약 27분)씩 3시간마다라
  //     다 긁는 데 **여러 날**이 걸렸다.
  //   * 그렇다고 늘 짧게 두면 몰아 긁기가 된다 - 3시간은 그것을 막으려던 값이다.
  //     그래서 **남은 일감으로 정한다**: 할 일이 없으면 예전 그대로 3시간이고,
  //     일감은 유한하므로 다 긁으면 저절로 3시간으로 돌아간다.
  var GAP_MS = 3 * 60 * 60 * 1000;      // 할 일이 없을 때(예전 그대로)
  var GAP_NEW_MS = 5 * 60 * 1000;       // **새 글**이 있으면 - 바로바로
  var GAP_WORK_MS = 30 * 60 * 1000;     // 남은 일(재수집/댓글/오염)이 있으면
  var PROBE_GAP_MS = 5 * 60 * 1000;     // 계획을 물어보는 최소 간격(앱을 쪼지 않는다)
  var TICK_MS = 60 * 1000;              // 확인 주기 - 실제로 긁을지는 위 간격이 정한다
  var HEARTBEAT_MS = 30 * 60 * 1000;    // 긁을 게 없어도 이 간격마다 '살아 있다'를 알린다
  var POLL_MS = 4000;
  var MAX_WAIT_MS = 30 * 60 * 1000;

  // ★ 실릴 때마다 도장을 찍는다 — **'안 깔림' 과 '깔렸는데 조용함' 을 가르기 위해서다.**
  //   되보고는 30분 간격이라, 설치 직후에 확인하면 둘이 똑같이 조용하다([169]).
  //   이 한 줄이 없어서 2026-08-13 설치 직후에 "깔렸나?"를 답할 수 없었다.
  //   localStorage 는 band.us 것이라 페이지에서도 읽힌다 — 사람도 개발자도 바로 본다.
  try {
    localStorage.setItem('coupangAutoCollect.loaded',
      String(Date.now()) + '|' + VER + '|' + (typeof GM_info !== 'undefined' ? 'gm' : 'raw'));
  } catch (e) { }

  function bandNo() {
    var m = location.pathname.match(/^\/band\/(\d+)\b/);
    return m ? m[1] : '';
  }
  function keyLast(band) { return 'coupangAutoCollect.last.' + band; }
  function keyBeat(band) { return 'coupangAutoCollect.beat.' + band; }
  function keyProbe(band) { return 'coupangAutoCollect.probe.' + band; }

  // 남은 일감으로 이번 간격을 정한다.
  //   * **모르면 예전 그대로 3시간이다**([169]) - 계획을 못 읽었는데 짧은 간격을
  //     쓰면 앱이 죽은 밤에 1분마다 헛돈다.  모름을 '할 일 있음'으로 치지 않는다.
  //   * 새 글(미수집)이 하나라도 있으면 **그것이 이긴다** - 형님이 말씀하신
  //     "바로바로"가 이 한 줄이다.  오염이 아무리 많아도 새 글을 안 밀어낸다.
  function gapFor(counts) {
    if (!counts) return GAP_MS;
    if ((counts['미수집'] || 0) > 0) return GAP_NEW_MS;
    var work = (counts['재수집'] || 0) + (counts['댓글'] || 0) + (counts['오염'] || 0);
    return work > 0 ? GAP_WORK_MS : GAP_MS;
  }
  function getTs(k) { return parseInt(localStorage.getItem(k) || '0', 10) || 0; }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  // ── 되보고 ────────────────────────────────────────────────────────────────
  //   `text/plain` 으로 보낸다 — CORS 사전요청(OPTIONS)이 안 붙는 '단순 요청'이라
  //   서버가 헤더를 안 갖춰도 **요청 자체는 도착한다**.  응답은 안 읽는다.
  //   보고가 실패해도 수집은 계속한다 — 알리려다 일을 못 하면 본말전도다.
  var _appBase = null;
  function report(band, state, extra) {
    var body = JSON.stringify(Object.assign({
      band: band, state: state, at: new Date().toISOString(), ver: VER,
      hidden: document.hidden, url: location.pathname
    }, extra || {}));
    var base = _appBase || APP_CANDIDATES[0];
    var url = base + '/api/collect_report';
    try {
      return fetch(url, {
        method: 'POST', mode: 'cors', keepalive: true,
        headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: body
      }).catch(function () {
        // 창이 닫히는 중이어도 남도록 beacon 으로 한 번 더.
        try { navigator.sendBeacon(url, new Blob([body], { type: 'text/plain' })); } catch (e) { }
      });
    } catch (e) { return Promise.resolve(); }
  }

  function findApp() {
    var i = 0;
    function tryOne() {
      if (i >= APP_CANDIDATES.length) return Promise.resolve(null);
      var base = APP_CANDIDATES[i++];
      return fetch(base + '/api/ping', { mode: 'cors' })
        .then(function (r) { return r.ok ? base : tryOne(); })
        .catch(tryOne);
    }
    return tryOne().then(function (b) { _appBase = b; return b; });
  }

  // 수집기 정본은 **앱이 내려 주는 grab_posts.js 하나**다([162]).
  // 규칙이 바뀌면 앱만 고치면 되고 이 파일은 안 바꾼다.
  // (밴드에는 CSP 가 없다 — 2026-08-13 실측으로 헤더·메타 둘 다 없어 주입이 막히지 않는다.)
  function loadCollector(url) {
    return new Promise(function (res, rej) {
      if (typeof window.__grabStart === 'function') return res(true);
      var s = document.createElement('script');
      s.src = url + (url.indexOf('?') < 0 ? '?v=' : '&v=') + Date.now();
      s.onload = function () {
        if (typeof window.__grabStart === 'function') res(true);
        else rej(new Error('수집기를 받았는데 __grabStart 가 없다'));
      };
      s.onerror = function () { rej(new Error('수집기 로드 실패: ' + url)); };
      document.head.appendChild(s);
    });
  }

  // 어디서 계획·수집기를 받을지: 로컬 앱이 먼저, 없으면 게시 사본.
  function resolveSource(band) {
    return findApp().then(function (base) {
      if (base) {
        return fetch(base + '/api/collect_plan?band=' + band, { mode: 'cors' })
          .then(function (r) { return r.json(); })
          .then(function (plan) {
            return {
              collector: base + '/grab_posts.js',
              nos: (plan && plan.nos) || [],
              // * 남은 일감 - 이번 간격을 정하는 근거다.  대기열이 이미 세어
              //   내려 주므로 **여기서 다시 세지 않는다**([162]).
              건수: (plan && plan['대기열'] && plan['대기열']['건수']) || null,
              생성: (plan && plan.generated) || '',
              출처: 'app'
            };
          });
      }
      return fetch(PAGES_BASE + '/plan.json', { mode: 'cors' })
        .then(function (r) { return r.ok ? r.json() : null; })
        .then(function (doc) {
          if (!doc) return null;
          var b = (doc.bands || {})[String(band)] || {};
          return {
            collector: PAGES_BASE + '/grab_posts.js',
            nos: b.nos || [], 생성: doc.generated || '',
            건수: b['건수'] || null,      // 없으면 null -> 예전 간격(3시간)
            출처: 'pages'
          };
        })
        .catch(function () { return null; });
    });
  }

  // ── 회차 ──────────────────────────────────────────────────────────────────
  function run() {
    var band = bandNo();
    if (!band) return;

    // ★ 숨은 탭에서는 밴드 본문이 안 그려진다.  긁으면 시각 없는 수확이 되어
    //   전부 버려지므로([130]) 시작하지 않는다.  다만 **조용히 넘어가지 않는다** —
    //   사람이 창을 앞으로 꺼내면 되는 것을 모르고 "왜 안 돌지"로 시간을 쓴다.
    if (document.hidden) {
      if (Date.now() - getTs(keyBeat(band)) >= HEARTBEAT_MS) {
        localStorage.setItem(keyBeat(band), String(Date.now()));
        report(band, 'hidden', { why: '창이 가려져 있거나 탭이 뒤에 있다 — 보이면 자동으로 시작한다' });
      }
      return;
    }

    // * **계획을 물어보는 것에도 최소 간격을 둔다.**  틱이 1분이라 이것이 없으면
    //   앱 `/api/ping` 을 1분마다 두드린다.  `keyLast` 는 **긁을 때만** 찍히므로
    //   그것으로는 못 막는다 - 안 긁는 동안 계속 커지기 때문이다.
    if (Date.now() - getTs(keyProbe(band)) < PROBE_GAP_MS) return;
    localStorage.setItem(keyProbe(band), String(Date.now()));

    var since = Date.now() - getTs(keyLast(band));

    resolveSource(band).then(function (src) {
      if (!src) {
        return report(band, 'no-source', { why: '앱도 게시 사본도 못 찾았다' });
      }
      // * 간격 판정이 **계획을 받은 뒤**로 왔다 - 남은 일감을 알아야 정할 수 있다.
      var gap = gapFor(src['건수']);
      if (since < gap) {
        // 아직 간격 전이라도 **살아 있다는 것은 알린다** - 이 한 줄이 없으면
        // '설치 안 됨' 과 '설치했고 방금 긁어서 쉬는 중' 이 똑같이 조용하다([169]).
        if (Date.now() - getTs(keyBeat(band)) >= HEARTBEAT_MS) {
          localStorage.setItem(keyBeat(band), String(Date.now()));
          report(band, 'idle', { why: '간격 대기 중', 다음: getTs(keyLast(band)) + gap,
                                 간격분: Math.round(gap / 60000), 건수: src['건수'] || null });
        }
        return;
      }
      if (!src.nos.length) {
        // 긁을 게 없는 것과 계획이 낡은 것은 다르다 — 생성 시각을 같이 보낸다.
        // 감시자가 '계획이 굳었다'를 이걸로 판단한다(2026-08-13: 계획이 하루 넘게
        // 안 갱신돼 크롬이 아무리 돌아도 긁을 게 없던 일이 있었다).
        return report(band, 'no-plan', { 계획생성: src.생성, 출처: src.출처 });
      }
      // 시작 표시를 **긁기 직전에** 찍는다 - 실패해도 곧바로 또 긁지 않게.
      //   (예전에는 계획을 받기 전에 찍었다.  간격 판정이 계획 뒤로 오면서
      //    여기로 내려왔다 - 안 옮기면 안 긁고도 간격이 소모된다.)
      localStorage.setItem(keyLast(band), String(Date.now()));
      localStorage.setItem(keyBeat(band), String(Date.now()));
      return loadCollector(src.collector).then(function () {
        return startAndSave(band, src.nos, src);
      }).catch(function (e) {
        localStorage.removeItem(keyLast(band));   // 다음 기회에 다시
        return report(band, 'error', { why: String(e && e.message || e).slice(0, 200) });
      });
    }).catch(function (e) {
      localStorage.removeItem(keyLast(band));
      return report(band, 'error', { why: String(e && e.message || e).slice(0, 200) });
    });
  }

  function startAndSave(band, nos, src) {
    // ★ 밴드 가드 — 시작 직전에 탭이 아직 그 밴드인지 본다.  사람이 그 사이
    //   다른 밴드로 옮겼는데 그대로 긁으면 **남의 밴드 글이 이 밴드 캐시에 들어간다**
    //   (유령 밴드 사고와 같은 종류라 되돌리기가 어렵다).
    if (bandNo() !== band) {
      localStorage.removeItem(keyLast(band));
      return report(band, 'moved', { why: '시작 직전에 탭이 다른 밴드로 옮겨졌다', 지금: bandNo() });
    }
    var r;
    try { r = window.__grabStart(Number(band), nos, { keep: false }); }
    catch (e) {
      localStorage.removeItem(keyLast(band));
      return report(band, 'error', { why: '__grabStart 예외: ' + String(e && e.message || e).slice(0, 160) });
    }
    // 수집기는 시작을 거절할 수 있다(숨은 탭 등).  그때는 표시를 지워 다음에 다시.
    var st0 = window.__grabStatus ? window.__grabStatus() : { running: false };
    if (!st0.running) {
      localStorage.removeItem(keyLast(band));
      return report(band, 'refused', { why: String(r).slice(0, 160), 요청: nos.length });
    }
    report(band, 'start', { 요청: nos.length, 출처: src.출처, 계획생성: src.생성 });

    var t0 = Date.now();
    return (function poll() {
      return sleep(POLL_MS).then(function () {
        var st = window.__grabStatus ? window.__grabStatus() : { running: false };
        if (st.running && (Date.now() - t0) < MAX_WAIT_MS) return poll();
        var timedOut = st.running;
        var saved = false, why = '';
        if (window.__grabSave) {
          try { window.__grabSave(); saved = true; }
          catch (e) { why = '저장 실패: ' + String(e && e.message || e).slice(0, 120); }
        } else { why = '__grabSave 가 없다'; }
        // 성공을 거짓으로 적지 않는다 — 저장이 안 됐으면 안 됐다고 말한다.
        return report(band, saved ? (timedOut ? 'partial' : 'done') : 'save-failed', {
          요청: nos.length,
          수확: st.ok != null ? st.ok : null,
          // ★ `__grabStatus()` 가 내놓는 이름은 **`failed`** 다(`grab_posts.js`).
          //   여기가 `st.fail` 을 읽고 있어서 되보고의 `실패` 는 **늘 null** 이었다 —
          //   낱말이 어긋나면 한 건도 안 걸리면서 오류도 안 난다([165]).
          //   2026-08-26 에 87건이 전부 실패한 회차를 뜯어보고서야 드러났다.
          //   ⚠ 이 값으로 **헛돎을 판정하지 않는다**([440]) — 없는 번호만 든 배치도
          //     정당하게 `실패 == 요청` 이다.  가르는 것은 시간이다.  여기는 **사실을
          //     그대로 적는 자리**이고, 사람이 뒤져 볼 때 쓰인다.
          실패: st.failed != null ? st.failed : (st.fail != null ? st.fail : null),
          걸린초: Math.round((Date.now() - t0) / 1000),
          why: why || (timedOut ? '시간이 넘어 도중까지만 저장했다' : '')
        });
      });
    })();
  }

  // 보이는 탭에서만 뜻이 있다 — 지금 보이면 곧, 아니면 보일 때.
  if (document.visibilityState === 'visible') setTimeout(run, 4000);
  else setTimeout(run, 4000);   // 숨어 있어도 한 번은 돌아 'hidden' 을 알린다
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'visible') setTimeout(run, 2000);
  });
  // 긴 세션에서도 소식이 끊기지 않게 — 탭을 며칠 열어 두는 것이 이 설계의 전제다.
  //   * 틱은 1분이고 **실제로 긁을지는 `gapFor` 가 정한다.**  예전에는 30분마다만
  //     확인해서, 간격을 5분으로 줄여도 30분에 한 번밖에 안 봤다.
  //     되보고는 그대로 30분마다다(`HEARTBEAT_MS` 로 거른다).
  setInterval(run, TICK_MS);
})();
