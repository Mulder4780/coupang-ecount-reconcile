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

────────────────────────────────────────────────────────────────────────────────
★ ec5 화면 몰이 — 이번 세션에서 **실제로 확인한 것만** 적는다
────────────────────────────────────────────────────────────────────────────────
되는 것 (실측)
  · 세션 유지: 탭이 이미 `?w_flag=1&ec_req_sid=…` 로 살아 있으면 그대로 쓴다.
    세션 키 없는 주소로 새로 들어가면 끊긴다(앞 세션 기록과 일치).
  · 사이트맵: 텍스트가 정확히 '사이트맵' 인 요소를 `.click()`.
    안 열리면 `.wrapper-sitemap` 에 `visible` 클래스를 직접 붙이면 열린다.
  · 메뉴 이동: 메뉴명이 **정확히** 일치하는 `<a>` 를 `.click()` → URL 해시에 prgId 가 붙는다.
    확인된 것: **계정별원장 = prgId E010807 · 격자 `#grid-EBZ057R`**
  · 엑셀 버튼은 `[data-cid="outputExcel"]` 로 잡힌다(화면에 1개 존재 확인).

아직 못 푼 것 — **다음 세션은 여기부터**
  · **조회 기간 위젯**. ec5 에는 네이티브 `<select>` 가 **하나도 없다**(0개 확인).
    날짜는 `button[data-role="select.selectbox"]` + 숨은 autocomplete `<input>` 조합이고,
    주소는 `data-cid` / `data-pcid` / `data-ecpath` 로 잡는다. 예:
        data-ecpath="EBZ057R_64892456398∫header∫all∫∫∫ddlSYear_SELECT∫ddlSYear"
  · ★ 함정: 간편검색 프리셋(`ddlSYear`, 화면에 '전월'이라 보이는 것)은 부모가
    `class="control hidden"` 이라 **눌러도 아무 일도 안 난다.** 이것을 누르고
    "클릭이 안 먹는다"고 헤매지 말 것 — 보이는 날짜 위젯은 따로 있다.

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


def main(argv=None):
    ap = argparse.ArgumentParser(description="ERP 내보내기가 종류별로 얼마나 밀렸나")
    ap.add_argument("--limit", type=int, default=DEFAULT_LIMIT_DAYS)
    a = ap.parse_args(argv)
    return run(a.limit)


if __name__ == "__main__":
    sys.exit(main())
