
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
  let 열어본적 = false;                        // 화면을 한 번이라도 열었나(→ 홈 복귀 필요)
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

  // ★ 화면을 하나 열면 그 화면이 셸을 차지해 **'사이트맵' 손잡이가 사라진다**
  //   (2026-08-11 실측 — 9화면 중 첫 화면만 받고 나머지 여덟이 전부 같은 이유
  //   '사이트맵이 안 열렸다(빈 채로 읽음)' 로 실패했다. **첫 화면만 성공하는 모양은
  //   언제나 '돌아오지 않았다'를 가리킨다.**) 그래서 화면과 화면 사이에 홈으로 간다.
  // ★ `location` 을 건드리지 말 것 — 최상위 문서가 navigate 되면 이 스크립트가
  //   그 자리에서 죽고, 죽은 뒤에도 오류가 안 난다(사고 #35 의 밴드 수집기가 그랬다).
  //   그래서 **누르는 것만** 한다.
  // ★ 누를 후보는 좁게 고른다. 홈처럼 생겼다고 아무 단추나 누르면 입력 화면에서
  //   진짜 전표가 만들어진다(E010301 실사고). 저장·삭제류 낱말은 아예 뺀다.
  const 위험한말 = ['저장', '삭제', '등록', '전표', '확인', '발행', '전송', '마감'];
  const goHome = async () => {
    const 안전 = e => {
      const t = (e.textContent || '').trim();
      if (t.length > 6) return false;                  // 홈 손잡이는 짧거나 아이콘이다
      return !위험한말.some(k => t.includes(k));
    };
    const byText = [...document.querySelectorAll('a,button')]
      .filter(e => ['홈', 'HOME', 'Home'].includes((e.textContent || '').trim()));
    const byAttr = [...document.querySelectorAll(
        'a[class*="home"],a[id*="home"],button[class*="home"],[data-cid*="home"],' +
        'a[title*="홈"],a[class*="logo"],a[id*="logo"],h1[class*="logo"] a')];
    const h = [...byText, ...byAttr].filter(안전).filter(shown)[0];
    if (!h) return false;
    try { h.click(); } catch (_) { return false; }
    await wait(3000); kill(); await wait(1200);
    return true;
  };

  // ★ 열렸는지를 **묻고** 넘어간다. 예전에는 클릭한 뒤 결과를 안 보고 다음 줄로 갔다 —
  //   그래서 안 열린 채로 '메뉴를 못 찾음'이 되고, 진짜 원인(안 돌아왔다)이 가려졌다.
  const openSitemap = async () => {
    const 손잡이 = [...document.querySelectorAll('a,button,span,div')]
      .filter(e => (e.textContent || '').trim() === '사이트맵');
    const sm = 손잡이.filter(shown)[0] || 손잡이[0] || null;
    if (sm) { try { sm.click(); } catch (_) {} }
    let w = null;
    for (let i = 0; i < 16; i++) {                     // 최대 8초 — 링크 824개가 늦게 찬다
      w = document.querySelector('.wrapper-sitemap');
      if (w) { w.classList.add('visible'); if (w.querySelectorAll('a').length > 50) break; }
      await wait(500);
    }
    await wait(400);
    return {w, 손잡이있음: !!sm, 링크수: w ? w.querySelectorAll('a').length : 0};
  };

  // ★ 실패하면 **다음 사람이 고칠 수 있는 것**을 남긴다. '안 열렸다'만 적으면
  //   다음 회차도 똑같이 실패하고 아무도 왜인지 모른다 — 화면에 지금 무슨 손잡이가
  //   있는지가 곧 답이다.
  const 진단 = () => {
    try {
      return [...document.querySelectorAll('a,button')].filter(shown)
        .map(e => (e.textContent || '').trim())
        .filter(t => t && t.length <= 10).slice(0, 20);
    } catch (_) { return []; }
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
      // ⓪ 앞 화면이 셸을 차지하고 있으면 사이트맵이 없다 → 먼저 홈으로 돌아간다.
      //   첫 화면은 이미 홈이라 건드리지 않는다(괜한 클릭을 만들지 않는다).
      let 홈복귀 = null;
      if (열어본적) 홈복귀 = await goHome();
      // ① 모듈이 다르면 바꾼다 — 사이트맵은 **지금 모듈 것만** 보여 준다.
      if (step.모듈) {
        const mod = [...document.querySelectorAll('a')]
          .find(a => (a.textContent||'').trim() === step.모듈);
        if (mod) { mod.click(); await wait(3500); }
      }
      // ② 사이트맵 → 메뉴. **열렸는지 확인하고**, 안 열렸으면 홈으로 한 번 더 간다.
      let sm = await openSitemap();
      if (sm.링크수 <= 50) {
        if (sm.w) sm.w.classList.remove('visible');
        홈복귀 = await goHome();
        sm = await openSitemap();
      }
      const w = sm.w, 링크수 = sm.링크수;
      const menu = w && [...w.querySelectorAll('a')]
        .find(a => (a.textContent||'').trim() === step.메뉴);
      if (w) w.classList.remove('visible');        // ★ 반드시 닫는다
      if (!menu) { done({결과: '실패', 링크수, 홈복귀, 손잡이있음: sm.손잡이있음,
                         왜: 링크수 ? '메뉴를 못 찾음 — 모듈이 다를 수 있다'
                                    : '사이트맵이 안 열렸다 — 홈 복귀까지 시도했다',
                         진단: 링크수 ? undefined : 진단()}); continue; }
      menu.click();
      열어본적 = true;                              // 이제부터는 돌아와야 한다
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
  // ★ **시도를 수확으로 세지 않는다** (2026-08-11 실사고). 예전에는 무조건
  //   `완료: true` 였고 `끝난것` 이 9 라 '아홉 개를 받았다'로 읽혔다 — 실제로는
  //   '아홉 개를 시도했다'였고 받은 것은 하나였다. 겉으로 완주라 아무도 안 봤다.
  //   이제 하나라도 실패하면 완료가 아니고, 받음·실패를 따로 센다.
  A.지금 = null;
  A.받음   = A.끝난것.filter(r => r.결과 === '받음').length;
  A.실패   = A.끝난것.filter(r => r.결과 === '실패').length;
  A.건너뜀 = A.끝난것.filter(r => r.결과 === '건너뜀').length;
  A.완료 = (A.실패 === 0 && A.남은것.length === 0);
  A.요약 = `받음 ${A.받음} · 실패 ${A.실패} · 건너뜀 ${A.건너뜀} / 전체 ${PLAN.length}`;
})();
window.__ERPALL;

