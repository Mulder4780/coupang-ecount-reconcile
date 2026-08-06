/* ── 밴드 84789192 수집 — 이 파일 전체를 복사해 밴드 탭 콘솔(F12)에 붙여넣으세요 ──
 * 붙여넣는 즉시 시작하고 **끝까지 알아서 갑니다**. 붙여넣기는 이 한 번뿐입니다.
 *   367건 · 2회차(회차당 최대 250건) · 대략 30분.
 * 회차가 끝날 때마다 dump 파일이 **자동으로 다운로드**되고 탭 메모리는 비워집니다
 * (그래야 얼지 않습니다). 다운로드된 파일은 손대지 마세요 — download_intake 가
 * Z: 로 옮기고 convert_dump 가 캐시로 합칩니다.
 * 진행 중 탭을 새로고침하면 그 회차분이 날아갑니다. 그냥 두면 됩니다(뒤에 있어도 돕니다).
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
  const ROUNDS = [[3539, 3540, 3541, 3542, 3543, 3544, 3545, 3546, 3547, 3548, 3549, 3550, 3551, 3552, 3553, 3554, 3555, 3556, 3557, 3558, 3559, 3560, 3561, 3562, 3563, 3564, 3565, 3566, 3567, 3568, 3569, 3570, 3571, 3572, 3573, 3574, 3575, 3576, 3577, 3578, 3333, 3332, 3331, 3330, 3329, 3328, 3327, 3326, 3325, 3324, 3323, 3322, 3321, 3320, 3319, 3318, 3317, 3316, 3315, 3314, 3313, 3312, 3311, 3310, 3309, 3308, 3307, 3306, 3305, 3304, 3303, 3302, 3301, 3300, 3299, 3298, 3297, 3296, 3295, 3294, 3293, 3292, 3291, 3290, 3289, 3288, 3287, 3286, 3285, 3284, 3283, 3282, 3281, 3280, 3279, 3278, 3277, 3276, 3275, 3274, 3273, 3272, 3271, 3270, 3269, 3268, 3267, 3266, 3265, 3264, 3263, 3262, 3261, 3260, 3259, 3258, 3257, 3256, 3255, 3254, 3253, 3252, 3251, 3250, 3249, 3248, 3247, 3246, 3245, 3244, 3243, 3242, 3241, 3240, 3239, 3238, 3237, 3236, 3235, 3234, 3233, 3232, 3231, 3230, 3229, 3228, 3227, 3226, 3225, 3224, 3223, 3222, 3221, 3220, 3219, 3218, 3217, 3216, 3215, 3214, 3213, 3212, 3211, 3210, 3209, 3208, 3207, 3206, 3205, 3204, 3203, 3202, 3201, 3200, 3199, 3198, 3197, 3196, 3195, 3194, 3193, 3192, 3191, 3190, 3189, 3188, 3187, 3186, 3185, 3184, 3183, 3182, 3181, 3180, 3179, 3178, 3177, 3176, 3175, 3174, 3173, 3172, 3171, 3170, 3169, 3168, 3167, 3166, 3165, 3164, 3163, 3162, 3161, 3160, 3159, 3158, 3157, 3156, 3155, 3154, 3153, 3152, 3151, 3150, 3149, 3148, 3147, 3146, 3145, 3144, 3143, 3142, 3141, 3140, 3139, 3138, 3137, 3136, 3135, 3134, 3133, 3132, 3131, 3130, 3129, 3128, 3127, 3126, 3125, 3124], [3123, 3122, 3121, 3120, 3119, 3118, 3117, 3116, 3115, 3114, 3113, 3112, 3111, 3110, 3109, 3108, 3107, 3106, 3105, 3104, 3103, 3102, 3101, 3100, 3099, 3098, 3097, 3096, 3095, 3094, 3093, 3092, 3091, 3090, 3089, 3088, 3087, 3086, 3085, 3084, 3083, 3082, 3081, 3080, 3079, 3078, 3077, 3076, 3075, 3074, 3073, 3072, 3071, 3070, 3069, 3068, 3067, 3066, 3065, 3064, 3063, 3062, 3061, 3060, 3059, 3058, 3057, 3056, 3055, 3054, 3053, 3052, 3051, 3050, 3049, 3048, 3047, 3046, 3045, 3044, 3043, 3042, 3041, 3040, 3039, 3038, 3037, 3036, 3035, 3034, 3033, 3032, 3031, 3030, 3029, 3028, 3027, 3026, 3025, 3024, 3023, 3022, 3021, 3020, 3019, 3018, 3017, 3016, 3015, 3014, 3013, 3012, 3011, 3010, 3009, 3008, 3007]];
  const BAND = 84789192;
  const say = (m) => console.log('%c' + m, 'color:#0b5cff;font-weight:700');
  // 회차가 끝나기를 기다린다 — 폴링은 setInterval 이면 충분하다(수집 자체는
  // grab_posts.js 안의 Worker 타이머가 돌리므로 숨은 탭에서도 늦춰지지 않는다).
  const waitRound = (i) => new Promise((res) => {
    const t = setInterval(() => {
      const s = window.__grabStatus();
      console.log(`밴드 ${BAND} [${i + 1}/${ROUNDS.length}] ${s.ok}/${s.total} 수집 · 없는글 ${s.missing} · 실패 ${s.failed} · ${s.sec}초`);
      if (!s.running) { clearInterval(t); res(); }
    }, 5000);
  });
  (async () => {
    // 앞 회차가 아직 돌고 있으면 그것부터 끝내고 저장한다. 그냥 시작하면
    // __grabStart 가 '이미 실행 중'으로 거절해 첫 회차 번호가 통째로 빠진다.
    if (window.__grabStatus().running) {
      say('앞 배치가 아직 돌고 있습니다 — 그것부터 끝내고 이어 갑니다.');
      await waitRound(-1);
      say('앞 배치 ' + window.__grabSave());
    }
    for (let i = 0; i < ROUNDS.length; i++) {
      // keep:false → 지난 회차 글을 탭에서 비운다. 이미 저장했으니 잃는 것이 없다.
      console.log(window.__grabStart(BAND, ROUNDS[i], { keep: false }));
      await waitRound(i);
      say(`[${i + 1}/${ROUNDS.length}] ` + window.__grabSave());
    }
    say('밴드 ' + BAND + ' — 전체 ' + ROUNDS.length + '회차 끝. 다운로드 파일은 그대로 두세요.');
  })();
})();
