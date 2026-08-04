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
 */
(function () {
  const S = (window.__GRAB = window.__GRAB || {
    band: null, posts: {}, done: [], missing: [], failed: [],
    running: false, startedAt: null, total: 0,
  });

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
  async function grabOne(band, no, waitMs) {
    const f = document.createElement('iframe');
    f.style.cssText = 'position:fixed;left:-9999px;top:0;width:1200px;height:900px';
    f.src = `https://www.band.us/band/${band}/post/${no}`;
    document.body.appendChild(f);
    try {
      await new Promise((r) => { f.onload = r; setTimeout(r, waitMs); });
      let d = null;
      for (let i = 0; i < 30; i++) {              // 본문이 그려질 때까지 기다린다
        await new Promise((r) => setTimeout(r, 400));
        d = f.contentDocument;
        if (d && d.querySelector('.postText')) break;
      }
      if (!d) return { status: 'fail' };
      // 삭제·권한 없는 글은 본문 자체가 없다 — '없는 글'로 구분해 남긴다.
      if (!d.querySelector('.postText')) {
        const body = (d.body && d.body.innerText) || '';
        return { status: /삭제|없는 게시글|권한|찾을 수 없/.test(body) ? 'missing' : 'fail' };
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
  window.__grabStart = function (band, nos, opt) {
    opt = opt || {};
    if (S.running) return '이미 실행 중 — __grabStatus() 로 보라';
    Object.assign(S, {
      band: String(band), running: true, startedAt: Date.now(),
      total: nos.length, posts: opt.keep === false ? {} : S.posts,
      done: [], missing: [], failed: [],
    });
    (async () => {
      for (const no of nos) {
        const r = await grabOne(band, no, opt.waitMs || 9000);
        if (r.status === 'ok') { S.posts[no] = r.post; S.done.push(no); }
        else if (r.status === 'missing') S.missing.push(no);
        else S.failed.push(no);
        await new Promise((r2) => setTimeout(r2, opt.gapMs || 300));
      }
      S.running = false;
      S.finishedAt = Date.now();
    })();
    return `시작: ${nos.length}건 (밴드 ${band})`;
  };

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
