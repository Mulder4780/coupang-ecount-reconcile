/* grab_all.js — grab_posts.js 위에 얹는 **전량 주행기** (2026-08-06 지시)
 *
 * 사용자 지시: "밴드 데이터 및 ERP 데이터 전체 순차적으로 긁어와 어떻게든 다 긁어와서
 * 정리해 / 밤을 새서라도".
 *
 * 왜 따로 만드나 — grab_posts.js 는 **한 배치 250건**이 상한이다(그 이상은 탭 렌더러가
 * 얼어 모은 것을 통째로 잃는다). 남은 것이 7천 건이면 배치가 서른 번이다. 그때마다
 * 사람이나 AI 가 "다음 250건" 을 걸어 주어야 한다면 밤새 돌 수가 없다.
 * 그래서 **페이지 안에서** 배치를 이어 돌린다:
 *   ① 250건 수집 → ② 덤프 저장(다운로드) → ③ 탭 메모리 비우기 → ④ 다음 250건
 * 밖에서는 가끔 `__allStatus()` 만 보면 된다. 폴링 한 번에 한 줄이라 대화가 안 붇는다.
 *
 * ★ 여기서도 `setTimeout` 을 쓰지 않는다 — grab_posts.js 의 워커 타이머(`__grabSleep`)를
 *   빌려 쓴다. 탭이 뒤에 있으면 페이지 타이머는 1분까지 늦춰져 밤새 몇 건도 못 돈다.
 *
 * 쓰는 법 (grab_posts.js 를 먼저 주입한 뒤)
 *    __allStart(90610953, 1, 3590)     // 이 범위를 통째로. 이미 가진 번호는 밖에서 걸러 준다
 *    __allStatus()                     // {batch, done, ok, missing, failed, running}
 *    __allStop()                       // 지금 배치까지만 하고 멈춘다
 */
(function () {
  const A = (window.__ALL = window.__ALL || {
    band: null, queue: [], idx: 0, batch: 0, batches: 0,
    ok: 0, missing: 0, failed: 0, running: false, stop: false,
    startedAt: null, lastSave: null, note: '',
  });
  const CHUNK = 200;   // 250 이 상한이지만 밤새 도는 길이라 여유를 둔다

  const sleep = (ms) => new Promise((r) => {
    // grab_posts.js 의 워커 타이머를 쓴다. 없으면(주입 전) 페이지 타이머로 떨어진다.
    if (typeof window.__grabSleep === 'function') return window.__grabSleep(ms).then(r);
    setTimeout(r, ms);
  });

  window.__allStart = function (band, from, to, skip) {
    if (A.running) return '이미 돌고 있다 — __allStatus()';
    if (typeof window.__grabStart !== 'function') return 'grab_posts.js 를 먼저 주입하라';
    const have = new Set((skip || []).map(Number));
    const q = [];
    for (let n = Number(from); n <= Number(to); n++) if (!have.has(n)) q.push(n);
    Object.assign(A, {
      band: String(band), queue: q, idx: 0, batch: 0,
      batches: Math.ceil(q.length / CHUNK),
      ok: 0, missing: 0, failed: 0, running: true, stop: false,
      startedAt: Date.now(), note: '',
    });
    (async () => {
      while (A.idx < A.queue.length && !A.stop) {
        const nos = A.queue.slice(A.idx, A.idx + CHUNK);
        A.batch += 1;
        // keep:false — 앞 배치의 글을 메모리에서 버린다. 안 버리면 탭이 무거워져 언다.
        window.__grabStart(A.band, nos, { keep: false });
        // 배치가 끝날 때까지 기다린다. 상태만 보고 판단한다(타이머는 워커 것이다).
        for (;;) {
          await sleep(4000);
          const s = window.__grabStatus();
          if (!s.running) break;
        }
        const s = window.__grabStatus();
        A.ok += s.ok; A.missing += s.missing; A.failed += s.failed;
        A.idx += nos.length;
        try {
          A.note = window.__grabSave();          // 배치마다 파일로 떨군다 — 탭이 죽어도 남는다
          A.lastSave = Date.now();
        } catch (e) { A.note = '저장 실패: ' + e; }
        await sleep(2500);                        // 다운로드가 자리 잡을 틈
      }
      A.running = false;
      A.finishedAt = Date.now();
    })();
    return `전량 시작: ${q.length}건 · ${A.batches}배치 (밴드 ${band})`;
  };

  window.__allStop = () => {
    A.stop = true;
    if (typeof window.__grabStop === 'function') window.__grabStop();
    return '지금 배치까지만 하고 멈춘다';
  };

  window.__allStatus = () => {
    const s = (typeof window.__grabStatus === 'function') ? window.__grabStatus() : {};
    const sec = A.startedAt ? Math.round((Date.now() - A.startedAt) / 1000) : 0;
    const done = A.idx + (s.ok || 0) + (s.missing || 0) + (s.failed || 0) - 0;
    return {
      band: A.band, running: A.running,
      batch: `${A.batch}/${A.batches}`,
      진행: `${Math.min(done, A.queue.length)}/${A.queue.length}`,
      ok: A.ok + (s.ok || 0), missing: A.missing + (s.missing || 0),
      failed: A.failed + (s.failed || 0),
      분: Math.round(sec / 60),
      초당: sec ? +((A.ok + (s.ok || 0)) / sec).toFixed(2) : 0,
      마지막저장: A.lastSave ? new Date(A.lastSave).toTimeString().slice(0, 8) : null,
      메모: A.note,
    };
  };

  return 'grab_all.js 준비됨 — __allStart(밴드, from, to, [이미가진번호])';
})();
