
// ERP **전 화면 몰이** — 로그인된 ec5 탭에서 한 번 넣으면 목록을 끝까지 돈다.
// 사용자 지시(2026-08-07): "긁어오라고 하면 모두 긁어와서".
//
// ★ 던져 놓고 폴링한다. 화면 하나에 25초가 넘는데 CDP `Runtime.evaluate` 는 45초에
//   끊긴다. 끝을 기다리며 실행하면 도중에 잘리고 **잘린 자리를 알 수 없다.**
//       window.__ERPALL   → {지금, 끝난것:[{키,결과,행,기간}], 남은것:[…], 완료}
// ★ 한 화면이 실패해도 **멈추지 않는다.** 멈추면 뒤의 멀쩡한 화면까지 못 받는다.
//   대신 무엇이 왜 실패했는지 남긴다 — 조용히 건너뛰면 '전부 받았다'로 읽힌다.
window.__ERPALL = {지금: null, 끝난것: [], 남은것: ["ledger", "tax", "slips", "taxinv", "taxleft", "taxhome", "slipview", "sales", "hometax"], 완료: false};
(async () => {
  const A = window.__ERPALL;
  const PLAN = [{"키": "ledger", "메뉴": "계정별원장", "모듈": "회계 I", "프리셋": "금월(~오늘)"}, {"키": "tax", "메뉴": "매출(세금)계산서현황", "모듈": "회계 I", "프리셋": "금월(~오늘)"}, {"키": "slips", "메뉴": "회계거래현황", "모듈": "회계 I", "프리셋": "금월(~오늘)"}, {"키": "taxinv", "메뉴": "매출(세금)계산서현황(재고)", "모듈": "재고 I", "프리셋": "금월(~오늘)"}, {"키": "taxleft", "메뉴": "(세금)계산서진행단계", "모듈": "", "프리셋": "금월(~오늘)"}, {"키": "taxhome", "메뉴": "매출(세금)계산서현황(세무)", "모듈": "", "프리셋": "금월(~오늘)"}, {"키": "slipview", "메뉴": "회계거래조회", "모듈": "", "프리셋": "금월(~오늘)"}, {"키": "sales", "메뉴": "매출(세금)계산서조회", "모듈": "", "프리셋": "금월(~오늘)"}, {"키": "hometax", "메뉴": "홈택스자료조회", "모듈": "", "프리셋": "금월(~오늘)"}];
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const kill = () => {                       // 경고 대화상자는 '취소'가 진행이다
    const d = [...document.querySelectorAll('.ui-dialog')].find(x => x.offsetParent !== null);
    if (!d) return false;
    const c = [...d.querySelectorAll('button')].find(b => (b.textContent||'').trim() === '취소');
    if (c) { c.click(); return true; }
    return false;
  };
  // ★ 화면마다 마크업이 **두 종류**다 (2026-08-08 실측).
  //   · 계정별원장류 → `data-cid="simpleSearch|searchGroup|outputExcel"` 가 있다.
  //   · (세금)계산서진행단계(E010849)·매출(세금)계산서현황(세무) → `data-cid` 가
  //     `year`·`month` 뿐이고 버튼은 **글자로만** 잡힌다. 격자는 `#grid-main`.
  //   그래서 cid 로 먼저 찾고, 없으면 글자로 다시 찾는다. cid 만 보면 이 두 화면은
  //   영영 "기간 프리셋 없음"으로 실패한다 — 버튼은 멀쩡히 있는데도.
  const shown = e => {                       // ★ `offsetParent` 로 재지 말 것.
    // position:fixed 안의 버튼은 offsetParent 가 null 이다. 실측에서 `검색(F8)` 이
    // '후보 1개, 보이는 것 0개'로 걸러졌다. 사각형 유무가 정확한 잣대다.
    try { return e.getClientRects().length > 0; } catch (_) { return true; }
  };
  // ★ 조회조건이 **접혀 있는 화면**이 있다 (매출(세금)계산서조회 E010727 등).
  //   버튼이 없는 게 아니라 접힌 칸 안에 있어 안 잡힌다. 그래서 프리셋을 못 찾으면
  //   여기서 한 번 펴고 다시 본다 — 펴기 전에 실패로 끝내면 "없다"고 잘못 적게 된다.
  //   ★ 누르는 것은 **여는 손잡이뿐**이다. 저장·전표 같은 단추는 글자가 달라 안 걸린다
  //     (E010301 실사고 — 입력 화면에서 단추를 누르면 진짜 전표가 만들어진다).
  //   ★ 전수 스캔 금지: 후보를 글자로 먼저 좁히고 그 몇 개에만 사각형을 묻는다.
  const expandSearch = async () => {
    const 손잡이 = ['조회조건', '검색조건', '상세조건', '조건'];
    const cand = [...document.querySelectorAll(
        'button,a,span[class*="btn"],div[class*="toggle"],[class*="fold"],[class*="collapse"]')]
      .filter(e => {
        const t = (e.textContent || '').trim();
        return t.length <= 8 && 손잡이.some(k => t === k || t.startsWith(k));
      });
    const h = cand.filter(shown)[0];
    if (!h) return false;
    h.click();
    await wait(1800);
    return true;
  };

  const pick = (cid, txt, exact) => {
    const hit = list => list.find(e => {
      const t = (e.textContent || '').trim();
      return exact ? t === txt : t.startsWith(txt);
    });
    if (cid) {
      const byCid = [...document.querySelectorAll('button[data-cid="' + cid + '"]')];
      const c = hit(byCid.filter(shown)) || hit(byCid);
      if (c) return c;
    }
    // 글자로 다시 — 후보만 모으고 그 몇 개에만 사각형을 묻는다(전수 스캔 금지).
    const byTxt = [...document.querySelectorAll('button,a,input[type="button"]')]
      .filter(e => {
        const t = (e.textContent || e.value || '').trim();
        return exact ? t === txt : t.startsWith(txt);
      });
    return byTxt.filter(shown)[0] || byTxt[0] || null;
  };

  const want = preset => {
    const t = new Date(), p2 = n => String(n).padStart(2, '0');
    const fmt = d => `${d.getFullYear()}/${p2(d.getMonth()+1)}/${p2(d.getDate())}`;
    const back = n => { const d = new Date(t); d.setDate(d.getDate() - n); return fmt(d); };
    const to = fmt(t);
    if (preset === '금일')        return {from: to, to};
    if (preset === '전일')        return {from: back(1), to: back(1)};
    if (preset === '금주(~오늘)') return {from: back(7), to};
    if (preset === '금월(~오늘)') return {from: `${t.getFullYear()}/${p2(t.getMonth()+1)}/01`, to};
    // ★ 기수(회계연도) 프리셋 — 이 회사는 기수 = 달력해다(E010809 요령에 '이번기수(=올해 전체)').
    //   이것이 없으면 2025년을 받으려 해도 want 가 '최근 45일'로 떨어져서, 격자에
    //   2025년 날짜만 있는 화면이 **'기간 밖 → 실패'** 로 버려진다. 조회는 제대로
    //   됐는데 Excel 을 안 누르고 실패로 적는다 — 화면만 보면 이유를 알 수 없다.
    if (preset === '이번기수') return {from: `${t.getFullYear()}/01/01`, to};
    // ★ 화면에 있는 낱말은 '직전기수' 다 (2026-08-08 실측으로 확인 — '전기수' 라고
    //   짐작했다가 '프리셋을 못 찾음' 으로 헛발질할 뻔했다). 둘 다 받아 둔다.
    if (preset === '직전기수' || preset === '전기수')
                              return {from: `${t.getFullYear()-1}/01/01`,
                                      to:   `${t.getFullYear()-1}/12/31`};
    return {from: back(45), to};
  };

  for (const step of PLAN) {
    A.지금 = step.키;
    A.남은것 = A.남은것.filter(k => k !== step.키);
    const done = r => A.끝난것.push(Object.assign({키: step.키, 메뉴: step.메뉴}, r));
    try {
      // ① 모듈이 다르면 먼저 바꾼다 — 사이트맵은 **지금 모듈 것만** 보여 준다.
      if (step.모듈) {
        const mod = [...document.querySelectorAll('a')]
          .find(a => (a.textContent||'').trim() === step.모듈);
        if (mod) { mod.click(); await wait(3500); }
      }
      // ② 사이트맵 → 메뉴
      const sm = [...document.querySelectorAll('a,button,span,div')]
        .find(e => (e.textContent||'').trim() === '사이트맵');
      if (sm) sm.click();
      // ★ 고정 대기로는 **빈 사이트맵**을 읽는다 (2026-08-08 실측 — 이것 때문에
      //   '메뉴 못 찾음'이 나왔고, 모듈이 다른 줄 알고 엉뚱한 데를 뒤졌다).
      //   링크 824개가 만들어지는 데 시간이 걸린다. 채워질 때까지 기다린다.
      let w = null;
      for (let i = 0; i < 16; i++) {              // 최대 8초
        w = document.querySelector('.wrapper-sitemap');
        if (w) { w.classList.add('visible'); if (w.querySelectorAll('a').length > 50) break; }
        await wait(500);
      }
      await wait(400);
      const 링크수 = w ? w.querySelectorAll('a').length : 0;
      const menu = w && [...w.querySelectorAll('a')]
        .find(a => (a.textContent||'').trim() === step.메뉴);
      if (w) w.classList.remove('visible');        // ★ 반드시 닫는다
      if (!menu) { done({결과: '실패', 링크수,
                         왜: 링크수 ? '메뉴를 못 찾음 — 모듈이 다를 수 있다'
                                    : '사이트맵이 안 열렸다(빈 채로 읽음)'}); continue; }
      menu.click();
      await wait(4500);
      // ③ 기간 프리셋 — 없으면 **조회조건이 접힌 것**이니 한 번 펴고 다시 본다
      let p = pick('simpleSearch', step.프리셋, true);
      if (!p && await expandSearch()) p = pick('simpleSearch', step.프리셋, true);
      if (!p) { done({결과: '실패', 왜: '기간 프리셋을 못 찾음(조회조건도 펴 봤다): '
                                       + step.프리셋}); continue; }
      p.click(); await wait(2500); kill(); await wait(4500);
      // ④ 조회
      const s = pick('searchGroup', '검색', false);
      if (!s) { done({결과: '실패', 왜: '검색 버튼을 못 찾음'}); continue; }
      s.click(); await wait(3000); kill(); await wait(9000);
      // ⑤ 조회가 **정말** 걸렸는지 — 행 수가 아니라 **격자에 찍힌 날짜**로 잰다.
      //   행 수로 재면 양쪽으로 다 틀린다(옛 결과를 새 기간으로 착각 / 5행→5행을 실패로 오판).
      const g = document.querySelector('[id^="grid-"]');
      const seen = [...(g ? g.querySelectorAll('tr') : [])]
        .map(t => ((t.textContent||'').match(/20\d\d\/\d\d\/\d\d/)||[])[0]).filter(Boolean);
      const rng = want(step.프리셋);
      const inR = seen.filter(d => d >= rng.from && d <= rng.to);
      if (!seen.length) { done({결과: '건너뜀', 왜: '격자에 날짜가 없다 — 0건이거나 조건이 더 필요한 화면',
                                기간: `${rng.from} ~ ${rng.to}`}); continue; }
      if (!inR.length)  { done({결과: '실패', 왜: '격자 날짜가 요청 기간 밖 — 조회가 안 걸렸다. Excel 안 누름',
                                기간: `${rng.from} ~ ${rng.to}`}); continue; }
      // ⑥ 엑셀
      // ★ 여기서만 **한 번** 누른다. 진행단계 화면처럼 Excel 단추가 둘인 곳에서
      //   후보를 모두 누르면 같은 파일이 두 벌 떨어진다(2026-08-08 실측 289KB ×2).
      const x = document.querySelector('[data-cid="outputExcel"]')
             || pick(null, 'Excel', true) || pick(null, '엑셀', true);
      if (!x) { done({결과: '실패', 왜: '엑셀 버튼이 없다 — 인쇄 미리보기 안에만 있는 화면일 수 있다'}); continue; }
      x.click(); await wait(5000);
      done({결과: '받음', 행: seen.length, 기간안: inR.length, 기간: `${rng.from} ~ ${rng.to}`});
    } catch (e) {
      done({결과: '실패', 왜: '예외: ' + (e && e.message)});
    }
  }
  A.지금 = null; A.완료 = true;
})();
window.__ERPALL;

