// ==UserScript==
// @name         쿠팡업무 — 밴드 자동 댓글수집 (Claude Code 없이)
// @namespace    coupang-ecount
// @version      1.0
// @description  로그인된 밴드 탭에서 앱의 수집계획을 받아 스스로 댓글을 긁는다. Claude Code/Codex 크레딧을 안 쓴다.
// @match        https://www.band.us/band/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==
//
// 왜 이게 있나 (2026-08-09 지시: "클로드 코드 또는 코덱스 없이 앱 자체적으로
//   처리할 수 있는 AI 알고리즘 기능 넣어서 클로드코드 크래딧 최대한 아껴")
//
//   밴드는 조회 API 가 없어 **로그인된 브라우저 DOM** 안에서만 글·댓글을 읽을 수 있다.
//   그 DOM 안에서 JS 를 돌릴 수 있는 것은 둘뿐이었다 — Claude Code(브라우저 도구)냐,
//   사람 브라우저 안의 스크립트냐. 지금까지는 Claude Code 가 매번 수집기를 주입하느라
//   크레딧을 썼다. 이 유저스크립트가 그 자리를 대신한다: **한 번 설치하면** 로그인된
//   밴드 탭이 스스로 긁는다. 나머지(덤프 흡수·대조·취소감지)는 이미 스케줄러가 무인으로 돈다.
//
//   'AI 알고리즘'(무엇을 긁을지 고르는 우선순위)은 앱이 회차에서 미리 계산해 둔
//   `reports/밴드_수집계획.json` 을 `/api/collect_plan` 으로 내려 준다. 이 스크립트는
//   그 계획을 **그대로 실행**만 한다 — 무작정 최근순으로 긁지 않는다(CLAUDE.md
//   '무작정 긁지 않는다 — 수집도 알고리즘이다').
//
//   설치: Tampermonkey(또는 Violentmonkey)에 이 파일을 추가한다. 앱 [실행] 탭의
//   '밴드 자동수집' 카드에서 원클릭 설치 링크를 받을 수 있다.
//
//   안전
//     · 수집기 자체는 앱이 내려 주는 `grab_posts.js` **하나**다(정본). 규칙이 바뀌면
//       앱만 고치면 되고 이 스크립트는 안 바꾼다.
//     · 탭이 뒤에 있으면 밴드 본문이 안 그려지므로 grab_posts.js 가 알아서 멈춘다 —
//       보이는 동안에만 긁는다. 그래서 이 창을 열어 두기만 하면 된다.
//     · 회차마다 몰아서 긁지 않게 밴드당 간격(기본 3시간)을 둔다.
//     · 확인 못 한 것은 기록하지 않는다(grab_posts.js [178]) — 거짓 성공이 안 난다.

(function () {
  'use strict';

  var APP_CANDIDATES = ['http://localhost:8899', 'http://127.0.0.1:8899'];
  var GAP_MS = 3 * 60 * 60 * 1000;          // 밴드당 최소 간격 (몰아 긁기 방지)
  var POLL_MS = 4000;                       // 진행 폴링 간격
  var MAX_WAIT_MS = 30 * 60 * 1000;         // 한 배치 최대 대기(안 끝나면 저장하고 끝)

  function bandNo() {
    var m = location.pathname.match(/^\/band\/(\d+)\b/);
    return m ? m[1] : '';
  }

  function keyLast(band) { return 'coupangAutoCollect.last.' + band; }

  function dueNow(band) {
    var last = parseInt(localStorage.getItem(keyLast(band)) || '0', 10) || 0;
    return (Date.now() - last) >= GAP_MS;
  }

  // 앱을 찾는다(로컬 우선). 못 찾으면 조용히 아무것도 안 한다 — 앱이 꺼져 있을 뿐이다.
  function findApp() {
    var i = 0;
    function tryOne() {
      if (i >= APP_CANDIDATES.length) return Promise.resolve(null);
      var base = APP_CANDIDATES[i++];
      return fetch(base + '/api/ping', { mode: 'cors' })
        .then(function (r) { return r.ok ? base : tryOne(); })
        .catch(tryOne);
    }
    return tryOne();
  }

  function loadCollector(base) {
    return new Promise(function (res, rej) {
      if (typeof window.__grabStart === 'function') return res(true);
      var s = document.createElement('script');
      s.src = base + '/grab_posts.js?v=' + Date.now();   // 정본을 앱에서 받는다
      s.onload = function () { res(true); };
      s.onerror = function () { rej(new Error('grab_posts.js 로드 실패')); };
      document.head.appendChild(s);
    });
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function run() {
    var band = bandNo();
    if (!band) return;
    if (!dueNow(band)) return;               // 아직 간격이 안 지났다
    // 시작 표시를 **먼저** 찍는다 — 실패해도 다음 로드에서 곧바로 또 긁지 않게.
    localStorage.setItem(keyLast(band), String(Date.now()));

    findApp().then(function (base) {
      if (!base) return;                     // 앱이 꺼져 있다 — 사람이 여는 게 먼저다
      return fetch(base + '/api/collect_plan?band=' + band, { mode: 'cors' })
        .then(function (r) { return r.json(); })
        .then(function (plan) {
          var nos = (plan && plan.nos) || [];
          if (!nos.length) return;           // 이 밴드는 긁을 게 없다
          return loadCollector(base).then(function () { return startAndSave(band, nos); });
        });
    }).catch(function () { /* 조용히 — 앱/네트워크 문제는 다음 회차에 다시 시도 */ });
  }

  function startAndSave(band, nos) {
    var r = window.__grabStart(Number(band), nos);
    // grab_posts.js 는 탭이 뒤에 있으면 시작을 거절한다 — 그때는 보일 때 다시.
    if (typeof r === 'string' && r.indexOf('시작') !== 0) {
      // 다음 로드/포커스에서 재시도하도록 표시를 지운다.
      localStorage.removeItem(keyLast(band));
      return;
    }
    var t0 = Date.now();
    return (function poll() {
      return sleep(POLL_MS).then(function () {
        var st = window.__grabStatus ? window.__grabStatus() : { running: false };
        if (st.running && (Date.now() - t0) < MAX_WAIT_MS) return poll();
        // 끝났다(또는 시간 넘었다) — 모은 것을 내려받는다. download_intake 가 흡수한다.
        if (window.__grabSave) { try { window.__grabSave(); } catch (e) { } }
      });
    })();
  }

  // 보이는 탭에서만 뜻이 있다 — 지금 보이면 곧, 아니면 보일 때.
  if (document.visibilityState === 'visible') {
    setTimeout(run, 4000);
  }
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'visible') setTimeout(run, 2000);
  });
})();
