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
  → **조회 전후로 격자 행 수를 세라.** 바뀌지 않으면 조회가 안 걸린 것이다.
    (이번 실측: 42행 → 56행 으로 바뀌는 것을 보고서야 Excel 을 눌렀다)

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
    print(f"ERP 내보내기 — 종류별 최신 (밀림 기준 {limit_days}일)")
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

  // ② 기간 프리셋
  const p = [...document.querySelectorAll('button[data-cid="simpleSearch"]')]
    .find(b => (b.textContent||'').trim() === '%(preset)s');
  if (!p) { say({오류:'기간 프리셋을 못 찾음'}); return; }
  p.click(); await wait(2500); kill(); await wait(5000);

  // ③ 조회 — 걸렸는지 **행 수 변화로** 확인한다(안 걸리면 옛 결과를 새 기간으로 착각한다)
  const before = rows();
  const s = [...document.querySelectorAll('button[data-cid="searchGroup"]')]
    .find(b => (b.textContent||'').trim().startsWith('검색'));
  if (!s) { say({오류:'검색 버튼을 못 찾음'}); return; }
  s.click(); await wait(3000); kill(); await wait(9000);
  const after = rows();
  say({단계:'조회함', 조회전: before, 조회후: after});
  // ★ 조회가 정말 걸렸는가. 행 수가 그대로면 **옛 결과를 새 기간으로 착각**하게 되므로
  //   Excel 을 누르지 않는다. 둘 다 0 이면 화면이 안 그려진 것이다(필터가 더 필요한 화면).
  if (after === 0)      { say({오류:'결과가 없다 — 이 화면은 거래처 등 조건이 더 필요하다'}); return; }
  if (after === before) { say({오류:'행 수가 그대로다 — 조회가 안 걸렸다. Excel 을 누르지 않는다'}); return; }

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
    a = ap.parse_args(argv)
    if a.js:
        return emit_js(a.js, a.preset)
    return run(a.limit)


if __name__ == "__main__":
    sys.exit(main())
