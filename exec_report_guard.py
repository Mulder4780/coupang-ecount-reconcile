# -*- coding: utf-8 -*-
"""대표보고에 **틀린 데이터가 들어가지 않게** 매번 검증한다 (2026-08-13 지시).

사용자 지시: **"대표 보고 화면 캡처 정말 중요함 / 잘못된 데이터 들어가지 않게
에이전트 만들어서 검증해 매번"**

시작은 형님의 한 마디였다 — *"잔여 미청구액 22000원 이거 오류인것 같은데"*.
집계기에 직접 물어보니 그 22,000원은 **합성 데모가 아니라 실제 행**이었다
(`UJ2601132 창원1MB(팔용동)`).  그런데 근거 열을 재 보니 진짜 문제가 그 뒤에 있었다.

실측 2026-08-13 · 06시트 750행:
  · **미청구액** 열 — 채워진 행 680, 그중 **0보다 큰 행은 1개**.  나머지 679행이 0 이다.
    그래서 화면이 "잔여 미청구액 22,000원"을 내놓는다.  이 회사가 못 받은 돈이
    2만 2천 원일 리 없다(세금계산서 미발행만 수백 건이다).
  · **미수금액** 열 — **채워진 행이 하나도 없다.**  그런데 화면은 `잔여 미수금액 0` 을
    **확언**한다.  이것이 `[169]` 다 — **없는 건가, 안 본 건가.**

즉 대표보고의 금액이 **뜻을 못 가진 채 사장님께 간다.**  오류도 안 나고 빈칸도
아니라서 아무도 안 봤다.  이 프로젝트가 반복해 당한 '조용한 사고'의 가장 비싼 판이다 —
잘못된 원장 값은 고치면 되지만, **잘못된 대표보고는 이미 사람의 판단을 바꾼 뒤다.**

## 무엇을 하나
매번(09:50 회차 · 캡처 직전) 대표보고를 **집계기에서 그대로 받아** 다섯 가지를 묻는다.
`truth_watch` 가 화면 전반에 대해 하는 일을 여기서는 **대표보고 한 장**에 대해 한다.

  ① **근거 열이 비어 있는데 숫자를 확언하나** — 채움률 0% 인 열에서 나온 `0` 은
     '0원'이 아니라 **'못 셈'**이다.
  ② **거의 다 0인 열에서 합계를 내나** — 680행 중 1행만 값이 있으면 그 합계는
     '잔액'이 아니라 '누가 한 칸 적었다'는 뜻이다.
  ③ **합성 데모 값이 실데이터에 섞였나** — 코드에 박힌 데모 상수·합성 낱말
     (`합성`·`데모`·`예시`)이 실보고에 나타나면 그 자리는 통째로 못 믿는다.
  ④ **금액이 말이 되나** — 음수 · 비정상 자릿수.
  ⑤ **근거가 얼마나 낡았나** — 캡처는 그 순간의 사실처럼 보인다.

## 지키는 것
- **읽기 전용이다.** 아무것도 안 고치고 큐에도 안 넣고 엑셀도 안 연다
  (`typo_watch`·`truth_watch` 와 같은 자리).  무엇이 맞는지는 사람만 안다.
- **판정을 새로 만들지 않는다**(`[162]`) — 화면이 실제로 쓰는 `app_server.get_exec_report`
  를 **그대로 불러** 그 결과를 잰다.  여기서 다시 집계하면 화면과 검증이 갈리고,
  **검증이 틀리면 아무도 그 사실을 모른다.**
- ⚠ **못 물어본 것을 '이상 없음'이라 하지 않는다**(`[169]`).  집계기를 못 부르면
  '통과'가 아니라 **'확인 못 함'** 이다.  이 구별이 없으면 감시자 자신이 눈먼 채
  "대표보고 정상"을 말한다.
- ⚠ **경보가 대부분이면 기준을 의심한다**(`[170]`).  그래서 '채움률 0%'와 '거의 0'은
  갈라서 세고, 낱말이 아니라 **숫자**로 말한다.

결과는 `reports/대표보고_검증.md`·`.json` → 인계 문서 '먼저 처리할 것'.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from typing import Any, Dict, List, Optional

# 무인 회차는 pythonw 라 `sys.stdout` 이 None 이고, 콘솔은 cp949 라 '—' 에서 죽는다([235]).
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_MD = os.path.join(ROOT, "reports", "대표보고_검증.md")
OUT_JSON = os.path.join(ROOT, "reports", "대표보고_검증.json")


#: 06시트에서 대표보고 금액이 근거로 삼는 열.  **여기 적힌 것만 잰다** —
#: 열 이름을 늘리려면 이 표를 고친다(부르는 쪽에 적으면 사본이 둘 된다, `[165]`).
#: ★ 2026-08-20 넷으로 넓혔다.  전에는 `잔여` 둘만 재서 **당일 지표 둘이 통째로
#:   사각지대**였다 — 실측 `입금일` 0/750 · `세금계산서합계` 0/750 이라
#:   `입금액 (당일)`·`세금계산서 발행액 (당일)` 은 **구조적으로 언제나 0** 인데
#:   화면은 매일 그 0 을 확언했고 감시자는 아무 말도 안 했다(`[169]`).
#: ★ 날짜 조건이 붙은 지표는 **날짜 열이 먼저 죽는다** — 금액 열이 멀쩡해도
#:   `입금일` 이 비면 한 건도 안 걸린다. 그래서 짝을 같이 적는다.
MONEY_COLUMNS = {
    "잔여 미청구액": "미청구액",
    "잔여 미수금액": "미수금액",
    "입금액 (당일)": "입금일",
    "세금계산서 발행액 (당일)": "세금계산서합계",
}

#: 채워진 행 중 값이 있는 비율이 이보다 낮으면 '거의 다 0' 이라 본다.
MOSTLY_ZERO_RATIO = 0.05


def _sheet06() -> Optional[List[Dict[str, Any]]]:
    """06시트를 읽는다.  못 읽으면 None — **빈 목록으로 돌려주지 않는다**(`[169]`)."""
    try:
        sys.path.insert(0, os.path.join(ROOT, "webapp"))
        sys.path.insert(0, ROOT)
        import warnings
        warnings.filterwarnings("ignore")
        import app_server as A
        from ecount_reconcile import resolve_master, load_config
        master = resolve_master(load_config()["reconcile"]["master_xlsx"])
        return A._sheet_records(A.master_book(master), "06_거래서류청구수금")
    except Exception:
        return None


def column_health(rows: Optional[List[Dict[str, Any]]], col: str) -> Dict[str, Any]:
    """그 열이 **근거가 될 만한가**.  채움률과 '값이 있는 행'을 따로 센다.

    ★ 둘을 가르는 것이 요점이다.  '채워짐'은 셀에 무엇이든 있다는 뜻이고, '값 있음'은
      0 이 아니라는 뜻이다.  680행이 채워졌는데 값 있는 행이 1개면, 그 합계는
      '잔액'이 아니라 **'누가 한 칸 적었다'** 는 뜻이다.
    """
    if rows is None:
        return {"열": col, "확인못함": "06시트를 못 읽었다"}
    total = len(rows)
    # ★ **없는 열과 빈 열은 다른 사실이다**(`[165]`).  예전에는 둘 다 '채워짐 0' 이라
    #   표에 오타 하나만 나도 감시자가 *"그 열이 비었다"* 고 **거짓 경보**를 내고,
    #   사람은 **없는 열을 채우러 간다**(`[172]`).  만들면서 그대로 밟았다 —
    #   `청구금액` 은 06시트에 없는 이름인데 '750행 중 0행'으로 나왔다.
    if total and col not in (rows[0] or {}):
        return {"열": col, "총행": total, "없는열": True,
                "확인못함": "06시트에 `%s` 라는 열이 없다" % col}
    filled = 0
    nonzero = 0
    for r in rows:
        raw = r.get(col)
        if str(raw or "").strip() not in ("", "None"):
            filled += 1
        try:
            if float(str(raw).replace(",", "")) > 0:
                nonzero += 1
        except (TypeError, ValueError):
            pass
    return {"열": col, "총행": total, "채워짐": filled, "값있음": nonzero,
            "채움률": round(filled / total, 4) if total else 0.0,
            "값비율": round(nonzero / total, 4) if total else 0.0}


def _amount_of(details: Dict[str, Any], label: str) -> Any:
    d = (details or {}).get(label) or {}
    return d.get("amount"), d.get("count"), (d.get("rows") or [])


def _exc_line(exc, mod=None, limit=300):
    """예외 -> **비지 않는** 한 줄. 갈래 이름을 절대 안 뺀다(`[292]`).

    ★ 왜 (2026-08-27 실측): 예전에는 `str(exc)[:120]` 뿐이라 **둘이 한꺼번에**
      무너졌다 — ① 갈래 이름(`FileNotFoundError`)이 없어 `autopilot._ERR_MARK`
      (`Error|Exception|Traceback|…`)가 오류 줄을 못 찾고 ② 120자에서 잘려
      **경로 끝의 `.xlsx` 가 사라진다**(실측: 그 경로는 `.xlsx` 까지 152자다).
      그러면 `autopilot.resource_back` 이 경로를 못 뽑아 영영 `None`(모름)이고,
      **지나간 자원 실패가 매일 아침 P1 로 인계 맨 위에 남는다**(`[461]` 이
      "안 고쳤다"고 적어 둔 자리다).

    ★ **둘 다 필요하다** — 실측으로 갈래 이름만 세우면 152자가 잘려 여전히
      `None` 이고, 자름만 늘리면 `_ERR_MARK` 가 못 찾아 여전히 `None` 이다.
      둘을 같이 하면 `True`(그 자원은 지금 살아 있다)가 나온다.
    ★ **만드는 자리는 하나다**(`[162]`) — `app_server.error_reason` 을 빌린다.
      그것은 갈래 이름과 **터진 자리**까지 세운다. 그 모듈을 못 부르는 갈래
      (import 자체가 실패한 자리)에서만 같은 계약으로 최소한을 만든다.
    """
    try:
        m = mod
        if m is None:
            import app_server as m  # noqa: PLC0415 - 오류 갈래에서만 늦게 부른다
        return m.error_reason(exc, limit)
    except Exception:  # noqa: BLE001 - 사유를 만들다 죽으면 사유가 통째로 없어진다
        kind = type(exc).__name__ or "Exception"
        why = str(exc).strip()
        line = ("%s: %s" % (kind, why)) if why else ("%s (사유 문구 없는 예외)" % kind)
        return line[:limit]


def build(day: Optional[str] = None) -> Dict[str, Any]:
    """대표보고를 **화면과 같은 길로** 받아 다섯 가지를 묻는다."""
    out: Dict[str, Any] = {"만든때": datetime.now().isoformat(timespec="seconds"),
                           "먼저볼것": [], "확인": [], "못물어봄": [], "잰것": {}}
    try:
        sys.path.insert(0, os.path.join(ROOT, "webapp"))
        sys.path.insert(0, ROOT)
        import warnings
        warnings.filterwarnings("ignore")
        import app_server as A
    except Exception as exc:
        out["못물어봄"].append("집계기를 못 불렀다: %s" % _exc_line(exc))
        return out

    if getattr(A, "DEMO", False):
        # ★ 데모 모드로 만든 보고가 캡처로 나가면 **지어낸 숫자가 사장님께 간다.**
        out["먼저볼것"].append("앱이 **합성 데모 모드**로 돌고 있다 — 이 보고의 숫자는 "
                             "전부 지어낸 값이다. 캡처로 내보내면 안 된다")
        return out

    try:
        rep = A.get_exec_report(day)
    except Exception as exc:
        out["못물어봄"].append("대표보고 집계가 실패했다: %s" % _exc_line(exc, A))
        return out

    details = rep.get("details") or {}
    meta = rep.get("meta") or {}
    out["기준일"] = meta.get("집계기준일") or meta.get("보고일") or (day or "")

    rows06 = _sheet06()
    if rows06 is None:
        out["못물어봄"].append("06시트를 못 읽어 근거 열의 채움률을 못 쟀다 — "
                             "이것은 '이상 없음'이 아니다")

    # ①② 근거 열이 비었나 · 거의 다 0인가
    dead: List[str] = []
    for label, col in MONEY_COLUMNS.items():
        amount, count, _rows = _amount_of(details, label)
        health = column_health(rows06, col)
        기준 = str((details.get(label) or {}).get("근거갈래") or "")
        out["잰것"][label] = {"보고값": amount, "건수": count,
                             "근거열": health, "근거갈래": 기준 or "06시트"}
        if 기준 == "ERP":
            # ★ 이 지표는 **06시트 열을 더 안 쓴다**([233] · 형님 2026-08-25 지시).
            #   죽은 열로 경보하면 거짓 경보가 되어 진짜 P1 을 덮는다([170]).
            #   그렇다고 조용히 빼지 않는다([169]) — 무엇으로 세는지 '잰것'에 남긴다.
            # ★ 표에 적어 두지 않고 **집계기가 실제로 붙인 표시**를 본다 — 나중에
            #   06시트로 되돌아가면 이 표시가 사라져 경보가 저절로 살아난다.
            continue
        if health.get("없는열"):
            # ★ 이것은 원장 문제가 아니라 **이 감시자의 표가 틀린 것**이다 —
            #   '먼저볼것'에 넣으면 사람이 없는 열을 채우러 간다(`[172]`).
            out["못물어봄"].append("`%s` — %s (MONEY_COLUMNS 를 고쳐야 한다)"
                                 % (label, health["확인못함"]))
            continue
        if health.get("확인못함"):
            continue
        if health["채워짐"] == 0:
            dead.append("%s 0/%d행" % (col, health["총행"]))
        elif health["값있음"] == 0:
            dead.append("%s 값 0/%d행" % (col, health["총행"]))
        elif health["값비율"] < MOSTLY_ZERO_RATIO:
            dead.append("%s 값 %d/%d행" % (col, health["값있음"], health["총행"]))
    if dead:
        # ★ **한 줄로 묶는다**(`[170]`).  원인이 하나인데(06시트 금액 칸은 사람 손
        #   입력이었고 2026-08-11 부터 그 입력이 끝났다) 네 줄로 울리면 진짜 P1 이
        #   묻힌다.  대신 **어느 열이 얼마나 비었는지 숫자로 전부 적는다**(`[169]`) —
        #   묶는 것과 감추는 것은 다르다.
        out["먼저볼것"].append(
            "대표보고 금액 지표 %d개의 근거 열이 06시트에서 **죽어 있다** — %s. "
            "이 숫자들은 '0원'이 아니라 **'못 셈'** 이다 — 화면이 확언하면 안 된다"
            % (len(dead), " · ".join(dead)))

    # ③ 합성 데모 섞임은 **여기서 안 본다** — 위 `A.DEMO` 한 줄이 이미 결정적이다.
    #    ★ 두 판 다 틀렸고 둘 다 **멀쩡한 행을 지목**했다:
    #      · 낱말(`합성`·`데모`·`예시`)로 훑기 → 사람이 적은 문제 본문에 그 낱말이 있다.
    #      · 식별자(`JS-2607-001`)로 훑기 → **그 번호는 실제 원장에도 실재한다**
    #        (`UJ2600975 송파5MB(감일동)`). 데모가 *진짜처럼 보이게* 지어졌기 때문이다.
    #    즉 **값으로는 데모와 실데이터를 구별할 수 없다.** 구별되는 것은 '어느 분기로
    #    만들었나' 뿐이고 그것은 `A.DEMO` 가 답한다. 못 잡는 것보다 **잘못 지목하는
    #    것이 나쁘다**(`[172]`) — 남겨 두면 매일 멀쩡한 행을 고치러 가게 만든다.

    # ④ 금액이 말이 되나 — ★ **뜻을 아는 칸만** 본다(`[172]`).
    #   첫 판이 모든 금액을 훑어 `작업금액 불일치 (현재)` 를 음수라고 지목했는데,
    #   그것은 **차액** 지표라 음수가 정상이다(실측 617건 · −458,839,843원).
    #   뜻을 모르는 칸에 규칙을 걸면 경보가 대부분이 되어 아무도 안 본다(`[170]`).
    for label in MONEY_COLUMNS:
        amt = (details.get(label) or {}).get("amount")
        if isinstance(amt, (int, float)) and amt < 0:
            out["먼저볼것"].append("**%s** 가 음수다(%s) — 잔액 칸에 음수가 나오면 "
                                 "부호나 열이 어긋난 것이다" % (label, amt))

    # ⑤ 근거가 얼마나 낡았나 — 캡처는 그 순간의 사실처럼 보인다
    base = str(out.get("기준일") or "")
    if base:
        try:
            d0 = datetime.fromisoformat(base[:10])
            age = (datetime.now().date() - d0.date()).days
            out["기준일나이"] = age
            if age >= 3:
                out["확인"].append("집계기준일이 **%d일 전(%s)** 이다 — 캡처는 오늘 것처럼 "
                                  "보이므로 날짜를 같이 보여 줘야 한다" % (age, base))
        except ValueError:
            out["못물어봄"].append("집계기준일을 못 읽었다: %r" % base)
    else:
        out["못물어봄"].append("집계기준일이 비어 있다")
    return out


def render(st: Dict[str, Any]) -> str:
    b: List[str] = ["# 대표보고 검증", "",
                    "만든때: %s · 집계기준일: %s" % (st.get("만든때"), st.get("기준일") or "?"), ""]
    if st.get("먼저볼것"):
        b.append("## 먼저 볼 것")
        b += ["- " + x for x in st["먼저볼것"]]
        b.append("")
    if st.get("확인"):
        b.append("## 확인")
        b += ["- " + x for x in st["확인"]]
        b.append("")
    if st.get("못물어봄"):
        # ★ 못 물어본 것을 '이상 없음'과 섞지 않는다 — 감시자가 눈먼 채 정상을 말한다.
        b.append("## 못 물어본 것 (이것은 '이상 없음'이 아니다)")
        b += ["- " + x for x in st["못물어봄"]]
        b.append("")
    if not (st.get("먼저볼것") or st.get("확인") or st.get("못물어봄")):
        b.append("먼저 볼 것 없음 — 잰 항목은 아래와 같다.")
        b.append("")
    b.append("## 잰 것")
    for label, v in (st.get("잰것") or {}).items():
        h = v.get("근거열") or {}
        if h.get("확인못함"):
            b.append("- **%s** = %s — 근거열 %s" % (label, v.get("보고값"), h["확인못함"]))
        else:
            b.append("- **%s** = %s (%s건) · 근거열 `%s`: %d행 중 채워짐 %d · 값있음 %d"
                     % (label, v.get("보고값"), v.get("건수"), h.get("열"),
                        h.get("총행", 0), h.get("채워짐", 0), h.get("값있음", 0)))
    return "\n".join(b) + "\n"


def lines(st: Optional[Dict[str, Any]] = None) -> List[str]:
    """인계 문서 '먼저 처리할 것' 에 올릴 줄.  깨끗하면 빈 목록이다."""
    s = st or build()
    out = ["대표보고 — " + x for x in (s.get("먼저볼것") or [])[:3]]
    out += ["대표보고 **확인 못 함** — " + x for x in (s.get("못물어봄") or [])[:2]]
    return out


def check(write: bool = True, day: Optional[str] = None) -> Dict[str, Any]:
    st = build(day)
    if write:
        os.makedirs(os.path.dirname(OUT_MD), exist_ok=True)
        for path, text in ((OUT_MD, render(st)),
                           (OUT_JSON, json.dumps(st, ensure_ascii=False, indent=1))):
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fh:
                fh.write(text)
            os.replace(tmp, path)
    return st


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="대표보고에 틀린 데이터가 들어가는지 매번 본다")
    ap.add_argument("--print", action="store_true", help="파일을 안 쓰고 화면에만")
    ap.add_argument("--day", default=None, help="집계기준일(기본: 앱이 정하는 날)")
    args = ap.parse_args(argv)
    st = check(write=not getattr(args, "print"), day=args.day)
    print("대표보고 검증 — 먼저볼것 %d · 확인 %d · 못물어봄 %d"
          % (len(st.get("먼저볼것") or []), len(st.get("확인") or []),
             len(st.get("못물어봄") or [])))
    for ln in (st.get("먼저볼것") or [])[:5]:
        print("  ! " + ln)
    for ln in (st.get("못물어봄") or [])[:3]:
        print("  ? " + ln)
    if not getattr(args, "print"):
        print("  상세: %s" % os.path.relpath(OUT_MD, ROOT))
    # 경보라고 exit 1 을 주지 않는다 — 회차 한 단계를 세우자고 감시자가 죽으면 안 된다.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
