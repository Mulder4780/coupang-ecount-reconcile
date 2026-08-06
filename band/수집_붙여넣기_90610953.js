/* ── 밴드 90610953 수집 — 이 파일 전체를 복사해 밴드 탭 콘솔(F12)에 붙여넣으세요 ──
 * 붙여넣는 즉시 시작합니다. 250건 · 글당 5초 안팎이라 약 20분.
 * 끝나면 dump 파일이 **자동으로 다운로드**됩니다. 그 파일은 손대지 마세요 —
 * download_intake 가 알아서 Z: 로 옮기고 캐시까지 합칩니다.
 * 진행 중에 탭을 새로고침하면 모은 것이 날아갑니다. 그냥 두면 됩니다(뒤에 있어도 돕니다).
 */
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
    running: false, startedAt: null, total: 0, stop: false,
  });

  // 워커 타이머 — 숨은 탭에서도 제 시각에 온다.
  const W = new Worker(URL.createObjectURL(new Blob(
    ['onmessage=e=>{setTimeout(()=>postMessage(e.data.id),e.data.ms)}'],
    { type: 'text/javascript' })));
  const waiters = {};
  let wid = 0;
  W.onmessage = (e) => { const f = waiters[e.data]; if (f) { delete waiters[e.data]; f(); } };
  const sleep = (ms) => new Promise((r) => { const id = ++wid; waiters[id] = r; W.postMessage({ id, ms }); });

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

  // 상세 페이지 한 글을 iframe 으로 열어 본문·글쓴이·시각·사진/댓글 수를 뜯는다.
  async function grabOne(band, no, waitMs, bodyMs) {
    const f = document.createElement('iframe');
    f.style.cssText = 'position:fixed;left:-9999px;top:0;width:1200px;height:900px';
    f.src = `https://www.band.us/band/${band}/post/${no}`;
    document.body.appendChild(f);
    try {
      await Promise.race([new Promise((r) => { f.onload = r; }), sleep(waitMs)]);
      // 본문은 SPA 가 나중에 그린다 — 폴링 대신 '그려지는 순간'을 관찰한다.
      const d = f.contentDocument;
      if (!d) return { status: 'fail' };
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
      const imgs = [...main.querySelectorAll('img')]
        .map((i) => i.src).filter((s) => /pstatic|phinf/.test(s));
      const cmt = (txt(main, '._commentCount, .comment .count, .uComment .count')
        .match(/\d+/) || [''])[0];
      return {
        status: 'ok',
        post: {
          author: txt(main, '.postWriterInfoWrap .text, .postWriter .text, .uProfileText'),
          timeText: txt(main, '.postListInfoWrap .time, .time'),
          content: txt(main, '.postText, .dPostTextView'),
          photo_count: String(imgs.length),
          comment_count: cmt || '0',
          images: [...new Set(imgs)],
        },
      };
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
    Object.assign(S, {
      band: String(band), running: true, stop: false, startedAt: Date.now(),
      total: nos.length, posts: opt.keep === false ? {} : S.posts,
      done: [], missing: [], failed: [],
    });
    (async () => {
      for (const no of nos) {
        if (S.stop) break;                       // __grabStop() 으로 중간에 끊을 수 있다
        const r = await grabOne(band, no, opt.waitMs || 9000, opt.bodyMs || 12000);
        if (r.status === 'ok') { S.posts[no] = r.post; S.done.push(no); }
        else if (r.status === 'missing') S.missing.push(no);
        else S.failed.push(no);
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
    running: S.running, total: S.total,
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
      missing: S.missing, failed: S.failed,
    };
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(doc)], { type: 'application/json' }));
    a.download = `dump_${stamp}_${S.band}.json`;
    document.body.appendChild(a); a.click(); a.remove();
    return `${Object.keys(S.posts).length}건 저장 → ${a.download} (download_intake 가 흡수한다)`;
  };

  return 'grab_posts.js 준비됨 — __grabStart(밴드, __grabRange(from,to))';
})();


(function () {
  const NOS = [5425, 5426, 5427, 5428, 5429, 5430, 5431, 5432, 5433, 5434, 5435, 5436, 5437, 5438, 5439, 5440, 5441, 5442, 5443, 5444, 5445, 5446, 5447, 5448, 5449, 5450, 5451, 5452, 5453, 5454, 5455, 5456, 5457, 5458, 5459, 5460, 5461, 5462, 5463, 5464, 3795, 3794, 3793, 3792, 3791, 3790, 3789, 3788, 3787, 3786, 3785, 3784, 3783, 3782, 3781, 3780, 3779, 3778, 3777, 3776, 3775, 3774, 3773, 3772, 3771, 3770, 3769, 3768, 3767, 3766, 3765, 3764, 3763, 3762, 3761, 3760, 3759, 3758, 3757, 3756, 3755, 3754, 3753, 3752, 3751, 3750, 3749, 3748, 3747, 3746, 3745, 3744, 3743, 3742, 3741, 3740, 3739, 3738, 3737, 3736, 3735, 3734, 3733, 3732, 3731, 3730, 3729, 3728, 3727, 3726, 3725, 3724, 3723, 3722, 3721, 3720, 3719, 3718, 3717, 3716, 3715, 3714, 3713, 3712, 3711, 3710, 3709, 3708, 3707, 3706, 3705, 3704, 3703, 3702, 3701, 3700, 3699, 3698, 3697, 3696, 3695, 3694, 3693, 3692, 3691, 3690, 3689, 3688, 3687, 3686, 3685, 3684, 3683, 3682, 3681, 3680, 3679, 3678, 3677, 3676, 3675, 3674, 3673, 3672, 3671, 3670, 3669, 3668, 3667, 3666, 3665, 3664, 3663, 3662, 3661, 3660, 3659, 3658, 3657, 3656, 3655, 3654, 3653, 3652, 3651, 3650, 3649, 3648, 3647, 3646, 3645, 3644, 3643, 3642, 3641, 3640, 3639, 3638, 3637, 3636, 3635, 3634, 3633, 3632, 3631, 3630, 3629, 3628, 3627, 3626, 3625, 3624, 3623, 3622, 3621, 3620, 3619, 3618, 3617, 3616, 3615, 3614, 3613, 3612, 3611, 3610, 3609, 3608, 3607, 3606, 3605, 3604, 3603, 3602, 3601, 3600, 3599, 3598, 3597, 3596, 3595, 3594, 3593, 3592, 3591, 3590, 3589, 3588, 3587, 3586];
  console.log(window.__grabStart(90610953, NOS));
  // 끝나는 순간 저장까지 자동으로 — 사람이 지켜보고 있을 필요가 없다.
  const t = setInterval(() => {
    const s = window.__grabStatus();
    console.log(`밴드 90610953 — ${s.ok}/${s.total} 수집 · 없는글 ${s.missing} · 실패 ${s.failed} · ${s.sec}초`);
    if (!s.running) {
      clearInterval(t);
      console.log('%c' + window.__grabSave(), 'color:#0b5cff;font-weight:700');
      console.log('%c끝났습니다. 다운로드된 dump 파일은 그대로 두세요.', 'color:#0b5cff;font-weight:700');
    }
  }, 5000);
})();
