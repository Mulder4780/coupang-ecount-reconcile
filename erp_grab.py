# -*- coding: utf-8 -*-
"""
erp_grab.py — ERP 내보내기가 **종류별로** 얼마나 밀렸나 + 화면 몰이 요령 (2026-08-07)

왜 이 파일이 필요했나
  인계 문서의 신선도 표는 "ERP 내보내기 · 최신 2026-08-07 · 밀림 0" 이라고 말한다.
  그런데 그건 **건별 PDF 폴더 하나**를 본 값이다. 같은 날 실제로는 이랬다:

      ERP 거래명세서(건별 PDF)  8/06 23:34   ← 이것만 보고 "밀림 0"
      ERP 세금계산서(건별 PDF)  8/07 00:43
      ERP:stmt   (거래명세서현황) 8/05 13:16
      ERP:sales  (매출계산서조회) 8/05 11:59
      ERP:ledger (계정별원장)     8/05 08:34
      ERP:tax / taxinv / slips    8/05

  **Excel 내보내기는 전부 8/5 에 멈춰 있고 PDF 만 오늘까지 왔다.**
  8/5 는 이틀 전이라 ERP 한도(7일)를 넘지는 않는다 — 규칙 위반은 아니다.
  문제는 다른 데 있다: 한 덩어리로 세면 **그 안에서 무엇이 멈췄는지 안 보인다.**
  PDF 하나가 오늘 날짜라서 나머지 여섯 종류가 사흘째 안 들어와도 표는 "밀림 0"
  이라고 말한다. 한도를 넘는 순간까지 아무도 모르는 것이 아니라, 넘고 나서도
  모른다 — 세는 단위가 틀렸기 때문이다. 그래서 여기서는 **종류마다** 센다.

  python erp_grab.py              # 종류별 최신일·밀린 일수
  python erp_grab.py --limit 3    # 며칠부터 밀림으로 볼지 (기본 2)
  python erp_grab.py --js 계정별원장   # 브라우저에 넣을 수집 스크립트를 찍는다

────────────────────────────────────────────────────────────────────────────────
★ ec5 화면 몰이 — 이번 세션에서 **실제로 확인한 것만** 적는다
────────────────────────────────────────────────────────────────────────────────
★ 끝까지 통한 순서 (2026-08-07 실측 — 계정별원장 8/01~8/07 을 이 순서로 받았다)
  ① 세션 유지: 탭이 이미 `?w_flag=1&ec_req_sid=…` 로 살아 있으면 **그대로 쓴다.**
     세션 키 없는 주소로 새로 들어가면 끊긴다.
  ② 사이트맵: 텍스트가 정확히 '사이트맵' 인 요소를 `.click()`.
     안 열리면 `.wrapper-sitemap` 에 `visible` 클래스를 직접 붙인다.
     ★ **다 쓰면 반드시 `visible` 을 떼라.** 열어 둔 채로 두면 그 뒤 모든 클릭을 가로챈다.
  ③ 메뉴 이동: 메뉴명이 **정확히** 일치하는 `<a>` 를 `.click()` → URL 해시에 prgId 가 붙는다.
     확인: **계정별원장 = prgId E010807 · 격자 `#grid-EBZ057R`**
  ④ 기간: `simpleSearch` 프리셋 버튼을 쓰는 것이 가장 쉽다. 텍스트로 고른다 —
     `금일` `전일` `금주(~오늘)` `전주` `금월(~오늘)` `전월` `이번기수`.
     **`금월(~오늘)` 이 곧 이번 달 1일~오늘**이라 '밀린 며칠'을 받기에 딱 맞다.
  ⑤ 조회: `button[data-cid="searchGroup"]` 중 텍스트가 `검색(F8)` 인 것을 `.click()`.
  ⑥ 경고 대화상자 — "계정 검색조건을 지정하지 않고…재지정하겠습니까?" 가 뜬다.
     **'취소'** 를 눌러야 조회가 진행된다('확인'은 되돌아간다).
  ⑦ 엑셀: `[data-cid="outputExcel"]` 를 `.click()` → Downloads 에 떨어진다.
     그다음 `python download_intake.py --apply` 가 내용판별로 Z: 에 넣는다.

★★ 조회가 **정말** 걸렸는지 반드시 확인할 것 — 여기가 가장 위험한 자리다
  날짜만 바꾸고 조회를 안 걸면 화면의 날짜는 새것인데 **격자는 옛 결과 그대로**다.
  그 상태로 Excel 을 받으면 8/7 자료를 8/1~8/7 이라고 믿게 된다. 실제로 겪었다
  (날짜 입력칸에 Enter 를 쳤더니 날짜만 바뀌고 42행이 그대로였다).

  ★ 그런데 **행 수로 재면 안 된다.** 그렇게 만들었다가 반대쪽으로도 틀렸다 —
    매출(세금)계산서현황에서 제대로 조회했는데 우연히 5행→5행 이라 가드가 막았고,
    격자를 열어 보니 8/03 위더스물류·8/05 뮤토가 멀쩡히 찍혀 있었다.
    **진짜 자료를 거짓 음성으로 버릴 뻔했다.**
  → 결정적인 증거는 하나뿐이다: **격자에 찍힌 날짜가 요청한 기간 안에 있는가.**
    행 수는 참고로만 남긴다.
  ★ 그 기간을 **화면에서 읽으면 안 된다.** 날짜 위젯 cid 가 화면마다 다르다 —
    회계거래현황에는 `ddlSYear_DATE` 가 아예 없어 기간이 "// ~ //" 로 읽혔고,
    그러면 비교가 전부 거짓이 되어 또 멀쩡한 자료를 버린다. **프리셋에서 계산한다.**

날짜 위젯 속살 (프리셋으로 안 되는 기간을 잡아야 할 때)
  · ec5 에는 네이티브 `<select>` 가 **하나도 없다**(0개 확인).
  · `button[data-cid="ddlSYear_DATE"][data-index=N]` — N: 0=시작년 1=시작월 3=종료년 4=종료월
    `input [data-cid="ddlSYear_DATE"][data-index=N]`  — N: 2=시작일 5=종료일 (이건 보이는 칸)
  · 월/년은 짝인 **숨은 `input`** 에 값을 넣고 `input`·`change`·Enter 를 쏘면 라벨이 바뀐다.
  · ★ 함정: 간편검색 프리셋(`ddlSYear_SELECT`, 화면에 '전월'이라 보이는 것)은 부모가
    `class="control hidden"` 이라 **눌러도 아무 일도 안 난다.** 이것을 누르고
    "클릭이 안 먹는다"고 헤매지 말 것 — 진짜 날짜 위젯은 `ddlSYear_DATE` 쪽이다.
  · ★ 함정: `data-cid="simpleSearch"` 는 조회 버튼이 아니라 **기간 프리셋**이다.
    아무거나 누르면 날짜가 그 프리셋으로 **덮어써진다**(내가 맞춰 둔 8월이 '금일'로 날아갔다).

★ 사이트맵은 **모듈마다 다르다** (2026-08-07 실측)
  회계 화면에서 연 사이트맵에는 `매출(세금)계산서현황` 이 있고, 대시보드에서 연 것에는
  `매출(세금)계산서현황(재고)` 가 있다. **이름이 다르면 못 찾는다** — "메뉴를 못 찾음"이
  나오면 먼저 지금 사이트맵에 그 이름이 실제로 있는지 확인하라.

★ 스크립트는 **메뉴 전환을 못 넘긴다** (2026-08-07 실측)
  메뉴를 누르면 화면이 통째로 갈리면서 돌던 async 가 '기간 설정'에서 멈춰 있었다.
  그래서 이 스크립트는 **메뉴가 열린 뒤에 다시 넣어** 기간~엑셀만 몰아야 한다.
  (--js 가 찍는 것은 메뉴 이동까지 포함하니, 멈추면 그 지점부터 다시 넣으면 된다)

★ 사이트맵은 **지금 있는 모듈 것만** 보여 준다 — 다른 모듈 메뉴는 목록에 아예 없다
  회계 화면에서 사이트맵을 열면 링크가 158개뿐이고 `거래명세서`·`판매조회` 는 없다.
  그것들은 재고/판매 모듈에 있다. 그러므로 **메뉴를 못 찾으면 먼저 모듈을 바꿔야 한다** —
  대시보드 상단 모듈 막대(`재고 I` `재고 II` `회계 I` `회계 II` `관리` `세무` …)를 누른 뒤
  사이트맵을 다시 연다. 같은 화면이 모듈에 따라 이름이 달라지기도 한다
  (`매출(세금)계산서현황` ↔ `매출(세금)계산서현황(재고)`).

모듈 전환은 화면 위쪽 **모듈 막대의 `<a>`** 를 텍스트로 골라 누르면 된다:
  `재고 I` `재고 II` `회계 I` `회계 II` `관리` `세무` `그룹웨어` `데이터센터`
  (회계에서 267개짜리 재고 사이트맵으로 바뀌는 것을 확인)

확인된 화면 (2026-08-07 실측)
  [회계] 계정별원장             prgId E010807 · 8/01~8/07 118행 수령 ✓ → ledger
  [회계] 매출(세금)계산서현황    prgId E010845 · 8/01~8/07 수령 ✓      → tax
  [회계] 회계거래현황           prgId E010847 · 8월분 수령 ✓          → slips
                               (이 화면엔 ddlSYear_DATE 가 없다 — 기간은 프리셋에서 계산할 것)
  [재고] 매출(세금)계산서조회(재고) prgId E040218 · 8/01~8/07 수령 ✓   → taxinv
  [회계] 거래처별계정별원장      prgId E010809 · 격자 자체가 없다 — 거래처 지정이 더 필요

아직 못 받은 것
  거래명세서현황 — [재고] `거래명세서인쇄` 쪽이다. **100건 선택 제한**이 있어 25건씩
  끊어야 하고, 인쇄 미리보기가 **다 그려진 뒤에** Excel 을 눌러야 파일이 생긴다
  (길이가 두 번 연속 같으면 완료). 미리보기를 10초만 기다리고 닫으면 파일이 안 만들어진다.

★ 스크립트를 **로컬 서버로 날라 오지 말 것** — ERP 페이지에서는 막힌다 (2026-08-08 실측)
  밴드에서는 `fetch('http://localhost:8123/…')` 로 긴 스크립트를 받아 `eval` 하는 방법이
  잘 통했다(컨텍스트를 아낀다). ec5 에서는 **통하지 않는다.** https 페이지에서 사설망
  (127.0.0.1) 으로 나가는 요청을 크롬이 막는데, 거절이 아니라 **그냥 걸린다** —
  `Runtime.evaluate` 가 45초를 다 쓰고 시간초과로 죽는다(실측). 렌더러가 얼어붙은 것으로
  오해하기 딱 좋다. ec5 에서는 스크립트를 **직접 넣는다.**

★★ 반드시 지킬 것 — **DOM 전수 스캔을 하지 말 것**
  `document.querySelectorAll('ul,div')` 를 전부 돌며 `getBoundingClientRect()` 를
  부르면 **렌더러가 얼어붙는다**(CDP 45초 시간초과, 실측). 탐침은 항상 좁게:
  선택자를 구체적으로 쓰고 `.slice(0, N)` 으로 끊는다.

★ 무차별 API 탐침 금지(절대규칙 3) — 화면을 몰아서 받는다.
"""
import argparse
import collections
import json
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(ROOT, "reports", "원본색인.json")

# 종류 코드 → 사람이 아는 이름. 색인의 `kind` 값을 그대로 쓴다.
KIND_LABEL = {
    "ERP:ledger": "계정별원장",
    "ERP:stmt":   "거래명세서현황",
    "ERP:sales":  "매출계산서조회",
    "ERP:taxinv": "매출세금계산서",
    "ERP:tax":    "세금계산서현황",
    "ERP:taxstep": "계산서진행단계(잔량)",
    "ERP:quote": "견적서조회",
    "ERP:slips":  "회계전표·거래",
    "ERP:hometax": "홈택스",
    "ERP 거래명세서(건별 PDF)": "거래명세서 PDF",
    "ERP 세금계산서(건별 PDF)": "세금계산서 PDF",
    "ERP": "ERP 기타",
}
DEFAULT_LIMIT_DAYS = 2

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def survey(index_path=None):
    """색인에서 ERP 종류별 최신 mtime 을 뽑는다 → {kind: (최신, 건수)}."""
    path = index_path or INDEX
    try:
        doc = json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}
    latest, count = {}, collections.Counter()
    for r in doc.get("rows") or []:
        kind = r.get("kind") or ""
        if not kind.startswith("ERP"):
            continue
        count[kind] += 1
        m = str(r.get("mtime") or "")
        if m and m > latest.get(kind, ""):
            latest[kind] = m
    return {k: (latest.get(k, ""), count[k]) for k in count}


def stale(rows, limit_days=DEFAULT_LIMIT_DAYS, today=None):
    """밀린 종류만 → [(kind, 최신, 밀린일수)] · 오래 밀린 순."""
    day = str(today or datetime.now().strftime("%Y-%m-%d"))[:10]
    out = []
    for kind, (m, _n) in rows.items():
        if not m:
            continue
        try:
            age = (datetime.strptime(day, "%Y-%m-%d")
                   - datetime.strptime(m[:10], "%Y-%m-%d")).days
        except ValueError:
            continue
        if age > limit_days:
            out.append((kind, m, age))
    return sorted(out, key=lambda x: -x[2])


def run(limit_days=DEFAULT_LIMIT_DAYS):
    rows = survey()
    if not rows:
        print("✗ reports/원본색인.json 을 못 읽었다 — 09:35 원본정리가 돌았는지 확인할 것")
        return 1
    # ★ 이 표는 **색인이 만들어진 시점**의 사실이다. 방금 받은 파일은 색인이 다시
    #   만들어지기 전까지 여기 안 보인다 — 그걸 모르고 보면 "받았는데 왜 그대로냐"가 된다.
    #   (색인은 09:35 원본정리·워치독이 다시 만든다)
    try:
        built = (json.load(open(INDEX, encoding="utf-8")) or {}).get("built") or "?"
    except Exception:
        built = "?"
    print(f"ERP 내보내기 — 종류별 최신 (밀림 기준 {limit_days}일)")
    print(f"  ※ 근거는 원본색인 {built} 기준 — 그 뒤에 받은 파일은 아직 안 보인다")
    for kind, (m, n) in sorted(rows.items(), key=lambda x: x[1][0], reverse=True):
        label = KIND_LABEL.get(kind, kind)
        print(f"  {label:<18} {m[:16] or '-':<17} {n:>5}건")
    late = stale(rows, limit_days)
    if not late:
        print("\n밀린 종류 없음.")
        return 0
    print(f"\n★ 밀린 것 {len(late)}종 — 한 덩어리로 세면 안 보인다:")
    for kind, m, age in late:
        print(f"  · {KIND_LABEL.get(kind, kind)}: 최신 {m[:10]} · {age}일 밀림")
    print("\n받는 법: 로그인된 ERP 탭에서 화면을 몰아 Excel 을 받는다(이 파일 위쪽 요령).")
    return 0


GRAB_JS = r"""
// ERP 화면 몰이 — 로그인된 ec5 탭에서 그대로 실행한다 (erp_grab.py --js 가 찍어 준 것)
// 메뉴: %(menu)s · 기간 프리셋: %(preset)s
// ★ **던져 놓고 폴링한다.** 이 절차는 대기만 25초가 넘는데 CDP `Runtime.evaluate` 는
//   45초에 끊긴다. 끝을 기다리며 실행하면 도중에 잘리고, 잘린 자리를 알 수 없다.
//   그래서 결과를 `window.__ERPGRAB` 에 남기고 즉시 반환한다 — 진행은 따로 읽는다:
//       window.__ERPGRAB            → {단계, 조회전, 조회후, 완료, 오류}
window.__ERPGRAB = {단계: '시작', 조회전: null, 조회후: null, 완료: false, 오류: null};
(async () => {
  const G = window.__ERPGRAB;
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const say  = o => Object.assign(G, o);
  const kill = () => {                       // 경고 대화상자는 '취소'가 진행이다
    const d = [...document.querySelectorAll('.ui-dialog')].find(x => x.offsetParent !== null);
    if (!d) return false;
    const c = [...d.querySelectorAll('button')].find(b => (b.textContent||'').trim() === '취소');
    if (c) { c.click(); return true; }
    return false;
  };
  // 격자가 아예 없으면 0 이 나온다 — 그건 "결과 0건"이 아니라 "화면이 안 그려졌다"이므로
  // 아래에서 조회 전후가 **둘 다 0** 이면 성공으로 치지 않는다.
  const rows = () => { const g = document.querySelector('[id^="grid-"]');
                       return g ? g.querySelectorAll('tr').length : 0; };

  // ① 사이트맵 → 메뉴. 쓰고 나면 반드시 닫는다(열어 두면 모든 클릭을 가로챈다).
  const sm = [...document.querySelectorAll('a,button,span,div')]
    .find(e => (e.textContent||'').trim() === '사이트맵');
  if (sm) sm.click();
  await wait(1500);
  const w = document.querySelector('.wrapper-sitemap');
  if (w) w.classList.add('visible');
  await wait(600);
  const menu = [...document.querySelectorAll('a')]
    .find(a => (a.textContent||'').trim() === '%(menu)s');
  if (!menu) { say({오류:'메뉴를 못 찾음'}); return; }
  menu.click();
  if (w) w.classList.remove('visible');      // ★ 반드시 닫는다
  await wait(4000);

  // ② 기간 프리셋 — 없으면 조회조건이 접힌 것이니 한 번 펴고 다시 본다
  //    (매출(세금)계산서조회 E010727 이 그랬다. 펴기 전에 실패로 끝내면 "없다"고 잘못 적는다)
  const preset = () => [...document.querySelectorAll('button[data-cid="simpleSearch"]')]
    .find(b => (b.textContent||'').trim() === '%(preset)s');
  let p = preset();
  if (!p) {
    // 여는 손잡이만 누른다 — 저장·전표 단추는 글자가 달라 안 걸린다(E010301 실사고)
    const h = [...document.querySelectorAll('button,a,span[class*="btn"],[class*="fold"]')]
      .filter(e => { const t=(e.textContent||'').trim();
                     return t.length <= 8 && /^(조회조건|검색조건|상세조건|조건)/.test(t); })
      .filter(e => { try { return e.getClientRects().length > 0; } catch(_) { return true; } })[0];
    if (h) { h.click(); await wait(1800); p = preset(); }
  }
  if (!p) { say({오류:'기간 프리셋을 못 찾음(조회조건도 펴 봤다)'}); return; }
  p.click(); await wait(2500); kill(); await wait(5000);

  // ③ 조회 — 걸렸는지 **행 수 변화로** 확인한다(안 걸리면 옛 결과를 새 기간으로 착각한다)
  const before = rows();
  const s = [...document.querySelectorAll('button[data-cid="searchGroup"]')]
    .find(b => (b.textContent||'').trim().startsWith('검색'));
  if (!s) { say({오류:'검색 버튼을 못 찾음'}); return; }
  s.click(); await wait(3000); kill(); await wait(9000);
  const after = rows();
  // ★ 조회가 정말 걸렸는가 — **행 수로 재지 마라.** 처음엔 그렇게 했다가 두 번 다 틀렸다:
  //   · 날짜만 바꾸고 조회를 안 걸었는데 42행 그대로 → 옛 결과를 새 기간으로 착각할 뻔
  //   · 제대로 조회했는데 우연히 5행→5행 → **진짜 자료를 거짓 음성으로 버릴 뻔**
  //   결정적인 증거는 하나뿐이다: **격자에 찍힌 날짜가 내가 요청한 기간 안에 있는가.**
  // ★ 기간은 **화면에서 읽지 말고 프리셋에서 계산한다** (2026-08-07 실측).
  //   날짜 위젯의 cid 는 화면마다 다르다 — 회계거래현황에는 `ddlSYear_DATE` 가 아예 없어
  //   기간이 "// ~ //" 로 읽혔고, 그러면 비교가 전부 거짓이 되어 **멀쩡한 8월 자료를
  //   거짓 음성으로 버린다.** 프리셋 이름은 어느 화면에서나 같으므로 이쪽이 튼튼하다.
  const want = (() => {
    const t = new Date(), p2 = n => String(n).padStart(2, '0');
    const fmt = d => `${d.getFullYear()}/${p2(d.getMonth()+1)}/${p2(d.getDate())}`;
    const back = n => { const d = new Date(t); d.setDate(d.getDate() - n); return fmt(d); };
    const to = fmt(t);
    const preset = '%(preset)s';
    if (preset === '금일')        return {from: to,       to};
    if (preset === '전일')        return {from: back(1),  to: back(1)};
    if (preset === '금주(~오늘)') return {from: back(7),  to};
    if (preset === '최근30일')    return {from: back(30), to};
    if (preset === '금월(~오늘)') return {from: `${t.getFullYear()}/${p2(t.getMonth()+1)}/01`, to};
    // 기수(회계연도) — 이 회사는 기수 = 달력해다. 없으면 2025년 조회가 '기간 밖'으로 버려진다.
    if (preset === '이번기수') return {from: `${t.getFullYear()}/01/01`, to};
    // ★ 화면에 있는 낱말은 '직전기수' 다 (2026-08-08 실측으로 확인 — '전기수' 라고
    //   짐작했다가 '프리셋을 못 찾음' 으로 헛발질할 뻔했다). 둘 다 받아 둔다.
    if (preset === '직전기수' || preset === '전기수')
                              return {from: `${t.getFullYear()-1}/01/01`,
                                      to:   `${t.getFullYear()-1}/12/31`};
    return {from: back(45), to};              // 모르는 프리셋은 넉넉하게 본다
  })();
  const g = document.querySelector('[id^="grid-"]');
  const seen = [...(g ? g.querySelectorAll('tr') : [])]
    .map(t => ((t.textContent||'').match(/20\d\d\/\d\d\/\d\d/)||[])[0]).filter(Boolean);
  const inRange = seen.filter(d => d >= want.from && d <= want.to);
  say({단계:'조회함', 조회전: before, 조회후: after, 기간: `${want.from} ~ ${want.to}`,
       날짜찍힌행: seen.length, 기간안: inRange.length});
  if (!seen.length)   { say({오류:'격자에 날짜가 없다 — 결과 0건이거나 조건이 더 필요한 화면'}); return; }
  if (!inRange.length){ say({오류:'격자 날짜가 요청 기간 밖이다 — 조회가 안 걸렸다. Excel 을 누르지 않는다'}); return; }

  // ④ 엑셀
  const x = document.querySelector('[data-cid="outputExcel"]');
  if (!x) { say({오류:'엑셀 버튼을 못 찾음'}); return; }
  x.click(); await wait(4000);
  say({단계:'엑셀 받음', 완료:true});
})();
// 여기서 즉시 반환된다 — 진행은 window.__ERPGRAB 을 다시 읽어서 본다.
window.__ERPGRAB;
"""

# 메뉴명 → 기본 기간 프리셋. '금월(~오늘)' 이 이번 달 1일~오늘이라 밀린 며칠을 받기 좋다.
MENUS = {
    "계정별원장": "금월(~오늘)",
    "거래처별계정별원장": "금월(~오늘)",
    "매출(세금)계산서현황(재고)": "금월(~오늘)",
    "판매현황": "금월(~오늘)",
}

# ── 화면 등록부 ────────────────────────────────────────────────────────────────
# ★ **메뉴 이름을 코드에 박아 두고 추측하지 말 것** (2026-08-07 실측).
#   같은 화면이 모듈마다 이름이 다르다(`매출(세금)계산서현황` ↔ `…(재고)`).
#   그래서 등록부는 '확인된 것'과 '아직 이름을 모르는 것'을 **나눠서** 들고 있고,
#   모르는 것은 `--find` 가 **화면에서 찾아** 채운다. 못 찾으면 못 찾았다고 말한다 —
#   비슷한 이름을 골라 누르면 엉뚱한 화면의 Excel 을 받아 놓고 맞다고 믿게 된다.
#
# 사람이 고친 등록부는 config/erp_screens.json 이 이긴다(코드 판올림에 안 지워진다).
SCREENS_CFG = os.path.join(ROOT, "config", "erp_screens.json")

SCREENS = {
    # 키          메뉴명(정확히)                모듈       프리셋           색인 kind
    "ledger":   {"메뉴": "계정별원장",               "모듈": "회계 I", "프리셋": "금월(~오늘)", "kind": "ERP:ledger", "prgId": "E010807"},
    "tax":      {"메뉴": "매출(세금)계산서현황",      "모듈": "회계 I", "프리셋": "금월(~오늘)", "kind": "ERP:tax",    "prgId": "E010845"},
    "slips":    {"메뉴": "회계거래현황",             "모듈": "회계 I", "프리셋": "금월(~오늘)", "kind": "ERP:slips",  "prgId": "E010847"},
    "taxinv":   {"메뉴": "매출(세금)계산서현황(재고)", "모듈": "재고 I", "프리셋": "금월(~오늘)", "kind": "ERP:taxinv", "prgId": "E040218"},
}

# 아직 **정확한 메뉴 이름을 확인하지 못한** 화면. `--find` 로 찾아 등록부에 넣는다.
# 사용자 지시(2026-08-07): "세금계산서 ERP에서 잔량 다운로드 받을 수 있어,
#                          매출 전표랑 같이해서 찾아서 전부 다운로드 받아"
WANTED = {
    "taxleft":   {"찾을말": ["잔량", "미발행", "발행잔량"], "무엇": "세금계산서 잔량",
                  "kind": "ERP:taxleft", "프리셋": "금월(~오늘)"},
}

# ★ **누르면 안 되는 화면** — 이름이 맞는데 받을 것이 없는 곳 (2026-08-08 실측).
#   찾다가 여기에 닿으면 "아직 못 찾았다"고 착각해 다시 뒤지게 된다. 그래서 왜 아닌지를
#   남긴다. 값은 사람이 읽는 설명이다.
# ★ **사람이 한 번 골라 줘야 하는 화면** — 자동 몰이에 넣으면 조용히 0건이 된다
#   (2026-08-08 실측). 받을 것이 없는 게 아니라 조건이 덜 찬 것이라 `NOT_GRABBABLE` 과 다르다.
#   0건은 실패처럼 안 보인다 — "그날 거래가 없었나 보다"로 읽힌다. 그래서 따로 적어 둔다.
NEEDS_HUMAN = {
    "거래처별계정별원장": {
        "prgId": "E010809",
        "왜": "거래처가 필수다. 비우고 검색하면 격자가 **빈 채로** 돌아온다(오류도 안 난다).",
        "요령": "간편검색 '이번기수'(=올해 전체) → 라디오 '건별' + '개별거래처기준' → "
                "거래처 칸에 '쿠팡' 입력 후 Enter → **거래처검색 창에서 사람이 줄을 클릭** → 검색(F8) → Excel",
        "막힌 곳": "거래처검색 창의 줄은 합성 이벤트(click·dblclick)로 안 골라진다. "
                   "창에 확인·선택 단추도 없다 — 실제 마우스 클릭이 있어야 한다.",
        "여기까지는 된다": "★ 2026-08-08 실측 — 이 창은 **키보드로 만든 창**이다. 창 안에 "
                   "단축키 안내표가 있다: `검색/커서위치적용 Enter · 다음페이지 → · 5행 위 Shift+↑`. "
                   "CDP 로 보낸 진짜 키 입력(↓ 5번)이 먹어서 `2548801036 쿠팡로지스틱스서비스` 줄에 "
                   "`class=\"tr-odd active-kbd\"` 가 붙었다 — **줄 고르기는 키보드로 된다.** "
                   "다만 그 상태에서 Enter 를 눌러도 창이 안 닫혔다(그때 activeElement 가 BODY 였다 — "
                   "격자가 키 초점을 안 쥔 듯). 남은 것은 '적용' 한 번뿐이다. "
                   "★ 함정: `input[data-cid=\"__headerQuick\"]` 이 **두 개**다. 첫째는 숨은 미끼라 "
                   "`querySelector` 로 잡으면 focus 가 안 붙는다 — `offsetParent !== null` 인 것을 고를 것.",
        "왜 필요한가": "입금 귀속의 **유일한 근거**다. 은행 원본·정리표에는 거래처명과 금액뿐이라 "
                       "어느 청구건에 붙는지 알 수 없다(적요 10종이 전부 거래처명이었다). "
                       "이 원장에는 적요·전표번호가 있다. 지금 이 자료는 **한 번도 수집된 적이 없다**.",
    },
}

NOT_GRABBABLE = {
    "매출전표 I": "prgId E010301 — **입력 화면**이다. 저장(F8)·저장/전표(F7)뿐이고 "
                  "검색·격자·Excel 이 없다. 여기서 단추를 누르면 실제 전표가 만들어진다. "
                  "사용자가 말한 '매출 전표'의 **조회**는 `회계거래조회`(E010701) 다.",
}


def load_screens():
    """코드 기본값 위에 config/erp_screens.json 을 덮어 돌려준다."""
    out = {k: dict(v) for k, v in SCREENS.items()}
    try:
        doc = json.load(open(SCREENS_CFG, encoding="utf-8"))
        for k, v in (doc.get("screens") or {}).items():
            if isinstance(v, dict) and v.get("메뉴"):
                out[k] = {**out.get(k, {}), **v}
    except Exception:
        pass
    return out


def save_screen(key, rec):
    """`--find` 가 찾아낸 메뉴 이름을 등록부에 적는다(다음부터는 안 찾아도 된다)."""
    try:
        doc = json.load(open(SCREENS_CFG, encoding="utf-8"))
    except Exception:
        doc = {}
    doc.setdefault("screens", {})[key] = rec
    doc["갱신"] = datetime.now().isoformat(timespec="seconds")
    os.makedirs(os.path.dirname(SCREENS_CFG), exist_ok=True)
    tmp = SCREENS_CFG + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, SCREENS_CFG)


FIND_JS = r"""
// 사이트맵에서 **이름을 찾아본다** — 누르지 않는다. 읽기만 하는 탐침이다.
// 왜 필요한가: 같은 화면이 모듈마다 이름이 다르다(`매출(세금)계산서현황` ↔ `…(재고)`).
//   이름을 추측해 누르면 **엉뚱한 화면의 Excel** 을 받아 놓고 맞다고 믿게 된다.
// ★ 사이트맵은 **지금 있는 모듈 것만** 보여 준다. 못 찾으면 모듈 막대를 바꾼 뒤 다시 돌린다.
window.__ERPFIND = {단계: '시작', 모듈: null, 링크수: 0, 찾음: [], 전체: []};
(async () => {
  const F = window.__ERPFIND;
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const sm = [...document.querySelectorAll('a,button,span,div')]
    .find(e => (e.textContent||'').trim() === '사이트맵');
  if (sm) sm.click();
  await wait(1200);
  const w = document.querySelector('.wrapper-sitemap');
  if (w) w.classList.add('visible');
  await wait(500);
  // ★ 좁게 훑는다. `ul,div` 전수 스캔은 렌더러를 얼린다(CDP 45초 시간초과 실측).
  const links = [...(w || document).querySelectorAll('a')]
    .map(a => (a.textContent||'').trim()).filter(t => t && t.length < 40);
  if (w) w.classList.remove('visible');            // ★ 반드시 닫는다
  const words = %(words)s;
  const hit = [...new Set(links.filter(t => words.some(x => t.includes(x))))];
  F.모듈 = ([...document.querySelectorAll('.module a, .gnb a')]
             .find(a => a.className.includes('on')) || {}).textContent || '?';
  F.링크수 = links.length;
  F.찾음 = hit;
  F.전체 = [...new Set(links)].slice(0, 400);      // 못 찾았을 때 사람이 눈으로 고를 수 있게
  F.단계 = '끝';
})();
window.__ERPFIND;
"""


def emit_find(words=None):
    """사이트맵에서 메뉴 이름을 찾는 탐침 스크립트를 찍는다."""
    ws = list(words or [])
    if not ws:
        for v in WANTED.values():
            ws.extend(v["찾을말"])
    print(FIND_JS % {"words": json.dumps(sorted(set(ws)), ensure_ascii=False)})
    return 0


ALL_JS = r"""
// ERP **전 화면 몰이** — 로그인된 ec5 탭에서 한 번 넣으면 목록을 끝까지 돈다.
// 사용자 지시(2026-08-07): "긁어오라고 하면 모두 긁어와서".
//
// ★ 던져 놓고 폴링한다. 화면 하나에 25초가 넘는데 CDP `Runtime.evaluate` 는 45초에
//   끊긴다. 끝을 기다리며 실행하면 도중에 잘리고 **잘린 자리를 알 수 없다.**
//       window.__ERPALL   → {지금, 끝난것:[{키,결과,행,기간}], 남은것:[…], 완료}
// ★ 한 화면이 실패해도 **멈추지 않는다.** 멈추면 뒤의 멀쩡한 화면까지 못 받는다.
//   대신 무엇이 왜 실패했는지 남긴다 — 조용히 건너뛰면 '전부 받았다'로 읽힌다.
window.__ERPALL = {지금: null, 끝난것: [], 남은것: %(keys)s, 완료: false};
(async () => {
  const A = window.__ERPALL;
  const PLAN = %(plan)s;
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
"""


def emit_all(keys=None, preset=None):
    """등록부의 화면을 **한 번에** 도는 스크립트를 찍는다."""
    screens = load_screens()
    want = [k for k in (keys or screens) if k in screens]
    미확인 = [k for k in (keys or []) if k not in screens]
    if 미확인:
        print(f"i 등록부에 없는 화면: {', '.join(미확인)}")
        print("  먼저 `python erp_grab.py --find` 로 사이트맵에서 이름을 찾아 등록하라")
    if not want:
        print("✗ 돌 화면이 없다 — 등록부가 비었다")
        return 1
    plan = [{"키": k, "메뉴": screens[k]["메뉴"], "모듈": screens[k].get("모듈"),
             "프리셋": preset or screens[k].get("프리셋", "금월(~오늘)")} for k in want]
    print(ALL_JS % {"keys": json.dumps(want, ensure_ascii=False),
                    "plan": json.dumps(plan, ensure_ascii=False)})
    return 0


def emit_js(menu, preset=None):
    if menu not in MENUS and preset is None:
        print(f"i '{menu}' 은 기본 프리셋이 없다 — --preset 으로 지정하라")
        print(f"  아는 메뉴: {', '.join(MENUS)}")
        return 1
    print(GRAB_JS % {"menu": menu, "preset": preset or MENUS.get(menu, "금월(~오늘)")})
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="ERP 내보내기가 종류별로 얼마나 밀렸나")
    ap.add_argument("--limit", type=int, default=DEFAULT_LIMIT_DAYS)
    ap.add_argument("--js", metavar="메뉴명", help="브라우저에 넣을 수집 스크립트를 찍는다")
    ap.add_argument("--preset", help="기간 프리셋(금일·전일·금주(~오늘)·금월(~오늘)·전월…)")
    ap.add_argument("--find", nargs="*", metavar="낱말",
                    help="사이트맵에서 메뉴 이름을 찾는다(안 주면 아직 못 찾은 것 전부)")
    ap.add_argument("--all", nargs="*", metavar="화면키",
                    help="등록부 화면을 한 번에 몬다(안 주면 전부)")
    ap.add_argument("--register", nargs=3, metavar=("키", "메뉴명", "모듈"),
                    help="--find 로 찾은 메뉴 이름을 등록부에 적는다")
    ap.add_argument("--screens", action="store_true", help="등록부를 보여 준다")
    a = ap.parse_args(argv)
    if a.register:
        key, menu, mod = a.register
        want = WANTED.get(key) or {}
        save_screen(key, {"메뉴": menu, "모듈": mod,
                          "프리셋": want.get("프리셋", "금월(~오늘)"),
                          "kind": want.get("kind", f"ERP:{key}")})
        print(f"등록 — [{key}] {menu} ({mod}) → {SCREENS_CFG}")
        return 0
    if a.screens:
        cur = load_screens()
        print(f"등록된 화면 {len(cur)}종")
        for k, v in cur.items():
            print(f"  {k:<10} {v['메뉴']:<26} {v.get('모듈') or '-':<7} {v.get('프리셋')}")
        미 = [k for k in WANTED if k not in cur]
        if 미:
            print(f"\n★ 아직 메뉴 이름을 못 찾은 것 {len(미)}종 — `--find` 로 찾는다:")
            for k in 미:
                print(f"  · {k:<10} {WANTED[k]['무엇']} (찾을말: {', '.join(WANTED[k]['찾을말'])})")
        if NEEDS_HUMAN:
            print(f"\n★ 사람이 한 번 골라 줘야 하는 화면 {len(NEEDS_HUMAN)}종"
                  " — 자동 몰이에 넣으면 조용히 0건이 된다:")
            for name, info in NEEDS_HUMAN.items():
                print(f"  · {name} ({info['prgId']}) — {info['왜']}")
                print(f"      요령: {info['요령']}")
                print(f"      막힌 곳: {info['막힌 곳']}")
        return 0
    if a.find is not None:
        return emit_find(a.find)
    if a.all is not None:
        return emit_all(a.all or None, a.preset)
    if a.js:
        return emit_js(a.js, a.preset)
    return run(a.limit)


if __name__ == "__main__":
    sys.exit(main())
