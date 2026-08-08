# -*- coding: utf-8 -*-
"""
sales_slip_gap.py — 계산서는 나갔는데 **회계 전표가 안 보이는** 건을 짝지어 찾는다.

왜 (2026-08-08 사용자 지시 "매출 전표 현황" → "다 해")
  같은 달을 두 화면이 다르게 말했다. 2026-07 은 계산서 진행단계 **49건**인데
  회계거래조회에는 매출전표가 **5건**뿐이었다. 6월은 31 = 31 로 정확히 맞았다.
  이런 어긋남을 눈으로 보면 "7월 매출이 1,316만원으로 줄었다"는 **틀린 보고**가 나간다.

  두 화면은 **같은 전표번호**를 쓴다(`2026/07/01 -1`). 그래서 짝짓기는 추측이 아니라
  열쇠 맞추기다 — 이 도구가 그것을 한다.

무엇을 답하나
  ① 계산서만 있고 전표가 없는 건  → 전표를 안 끊었거나, 내보내기가 그 기간을 덜 담았다
  ② 전표만 있고 계산서가 없는 건  → 매출은 잡혔는데 계산서가 안 나갔다
  ③ 둘 다 있는데 **금액이 다른** 건 → 정정이 한쪽에만 반영됐다
  ④ 미발행으로 표시된 건          → 사람이 발행해야 하는 목록

  ★ 여기서 **발행하지 않는다.** 전표 실전송은 사람 몫이다(사용자 지시).
    이 도구는 목록을 만들 뿐이고, ERP 를 건드리지 않는다(DB 읽기 전용).

  python sales_slip_gap.py                 # 최근 6개월
  python sales_slip_gap.py --from 2026-01  # 기간 지정
  python sales_slip_gap.py --month 2026-07 # 한 달만
"""
import argparse
import json
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

WON = lambda v: format(int(v or 0), ",")


def _key(s):
    """`2026/07/01 -1` · `2026/07/01-1` → 같은 열쇠로. 공백만 다른 표기를 흡수한다."""
    return "".join(str(s or "").split())


def collect(con, since, until):
    """진행단계(계산서) · 회계거래조회(전표) 를 전표번호로 모은다."""
    inv, slip = {}, {}
    for r in con.execute(
            "SELECT biz_date,party,amount,status,payload FROM record"
            " WHERE kind='ERP:taxstep' AND biz_date>=? AND biz_date<=?", (since, until)):
        p = json.loads(r["payload"])
        k = _key(p.get("일자-No."))
        if k:
            inv[k] = {"일자": r["biz_date"], "거래처": r["party"], "금액": r["amount"],
                      "상태": r["status"] or "", "적요": p.get("적요명", ""),
                      "프로젝트": p.get("프로젝트명", ""), "승인번호": p.get("승인번호", "")}
    for r in con.execute(
            "SELECT biz_date,party,amount,status,payload FROM record"
            " WHERE kind='ERP:slips' AND biz_date>=? AND biz_date<=?", (since, until)):
        p = json.loads(r["payload"])
        # ★ 매출/매입은 **`status` 로** 가른다. 화면마다 그 칸의 이름이 다르다 —
        #   회계거래조회는 `입력메뉴`, 회계거래현황은 `거래유형`이고 값도
        #   `매출전표 I` · `매출전표 I(매출)` 로 다르다. payload 에서 특정 이름을
        #   찾으면 다른 모양이 통째로 빠진다: 실제로 7/09~8/16 의 **50건이 사라져**
        #   "7월 매출전표가 5건뿐"이라는 유령 구멍을 만들었다(2026-08-08).
        #   ERP_MAP 이 모양마다 그 칸을 `상태` 로 모아 주므로 여기서는 하나만 본다.
        if not str(r["status"] or "").startswith("매출"):
            continue                       # 매입전표는 이 물음의 대상이 아니다
        k = _key(p.get("전표번호"))
        if k:
            slip[k] = {"일자": r["biz_date"], "거래처": r["party"], "금액": r["amount"],
                       "적요": p.get("적요명", "")}
    return inv, slip


def diagnose(inv, slip):
    """네 갈래로 가른다. 짝은 **전표번호**로만 맞춘다(추측하지 않는다)."""
    only_inv = [dict(v, 전표번호=k) for k, v in sorted(inv.items()) if k not in slip]
    only_slip = [dict(v, 전표번호=k) for k, v in sorted(slip.items()) if k not in inv]
    diff = []
    for k in sorted(set(inv) & set(slip)):
        a, b = inv[k], slip[k]
        # 금액은 원 단위 정수로만 비교한다(엑셀이 실수로 담을 때가 있다)
        if int(a["금액"] or 0) != int(b["금액"] or 0):
            diff.append({"전표번호": k, "일자": a["일자"], "거래처": a["거래처"],
                         "계산서금액": a["금액"], "전표금액": b["금액"],
                         "적요": a["적요"]})
    unissued = [dict(v, 전표번호=k) for k, v in sorted(inv.items())
                if str(v["상태"]).startswith("미발행")]
    return {"계산서만": only_inv, "전표만": only_slip, "금액다름": diff, "미발행": unissued}


def month_table(inv, slip):
    """달마다 계산서·전표가 몇 건인지 — **어긋난 달을 눈에 띄게** 한다."""
    ms = {}
    for k, v in inv.items():
        ms.setdefault(v["일자"][:7], [0, 0])[0] += 1
    for k, v in slip.items():
        ms.setdefault(v["일자"][:7], [0, 0])[1] += 1
    return [(m, c[0], c[1]) for m, c in sorted(ms.items(), reverse=True)]


def report(g, tbl, since, until):
    L = ["# 매출 전표 ↔ 계산서 대조", "",
         "- 기간 **%s ~ %s** · 만든 시각 %s" % (since, until,
                                               datetime.now().strftime("%Y-%m-%d %H:%M")),
         "- 짝은 **전표번호**로만 맞춘다(`2026/07/01 -1`). 추측으로 잇지 않는다.",
         "- 이 문서는 목록일 뿐이다. **발행·전표 생성은 사람이 ERP 에서** 한다.", ""]
    L += ["## 달마다 몇 건인가", "", "| 월 | 계산서(진행단계) | 매출전표 | |",
          "|---|---:|---:|---|"]
    for m, a, b in tbl:
        mark = "" if a == b else ("★ %d건 어긋남" % abs(a - b))
        L.append("| %s | %d | %d | %s |" % (m, a, b, mark))
    L.append("")
    if g["미발행"]:
        L += ["## ① 미발행 — 사람이 발행해야 하는 것 (%d건)" % len(g["미발행"]), "",
              "| 일자 | 전표번호 | 거래처 | 금액 | 프로젝트 |", "|---|---|---|---:|---|"]
        for x in g["미발행"]:
            L.append("| %s | %s | %s | %s | %s |"
                     % (x["일자"], x["전표번호"], x["거래처"], WON(x["금액"]),
                        str(x["프로젝트"])[:40]))
        L.append("")
    if g["계산서만"]:
        L += ["## ② 계산서는 있는데 매출전표가 안 보인다 (%d건)" % len(g["계산서만"]), "",
              "> 전표를 안 끊었거나, 회계거래조회 내보내기가 그 기간을 덜 담은 것이다.",
              "> **어느 쪽인지는 ERP 화면을 열어야 갈린다** — 여기서는 못 정한다.", "",
              "| 일자 | 전표번호 | 거래처 | 금액 | 적요 |", "|---|---|---|---:|---|"]
        for x in g["계산서만"][:200]:
            L.append("| %s | %s | %s | %s | %s |"
                     % (x["일자"], x["전표번호"], x["거래처"], WON(x["금액"]),
                        str(x["적요"])[:36]))
        if len(g["계산서만"]) > 200:
            L.append("| … | | | | 그 밖 %d건 |" % (len(g["계산서만"]) - 200))
        L.append("")
    if g["전표만"]:
        L += ["## ③ 매출전표는 있는데 계산서가 안 보인다 (%d건)" % len(g["전표만"]), "",
              "| 일자 | 전표번호 | 거래처 | 금액 | 적요 |", "|---|---|---|---:|---|"]
        for x in g["전표만"][:200]:
            L.append("| %s | %s | %s | %s | %s |"
                     % (x["일자"], x["전표번호"], x["거래처"], WON(x["금액"]),
                        str(x["적요"])[:36]))
        L.append("")
    if g["금액다름"]:
        L += ["## ④ 둘 다 있는데 금액이 다르다 (%d건)" % len(g["금액다름"]), "",
              "| 일자 | 전표번호 | 거래처 | 계산서 | 전표 | 차이 |",
              "|---|---|---|---:|---:|---:|"]
        for x in g["금액다름"]:
            d = int(x["계산서금액"] or 0) - int(x["전표금액"] or 0)
            L.append("| %s | %s | %s | %s | %s | %s |"
                     % (x["일자"], x["전표번호"], x["거래처"],
                        WON(x["계산서금액"]), WON(x["전표금액"]), WON(d)))
        L.append("")
    if not any(g.values()):
        L.append("어긋난 것 없음 — 계산서와 매출전표가 모두 짝을 이룬다.")
    return "\n".join(L)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="since", help="시작 월/일 (예: 2026-01)")
    ap.add_argument("--until", help="끝 월/일")
    ap.add_argument("--month", help="한 달만 (예: 2026-07)")
    ap.add_argument("--out", help="리포트 파일 경로")
    a = ap.parse_args(argv)

    if a.month:
        since, until = a.month + "-01", a.month + "-31"
    else:
        since = (a.since or "2026-01") + ("-01" if len(a.since or "2026-01") == 7 else "")
        until = (a.until or datetime.now().strftime("%Y-%m")) + "-31"

    import datalake as D
    con = D.connect()
    try:
        inv, slip = collect(con, since, until)
        g = diagnose(inv, slip)
        tbl = month_table(inv, slip)
        md = report(g, tbl, since, until)
        D.log(con, "erp", "sales_slip_gap", ok=True,
              detail={"기간": "%s~%s" % (since, until), "계산서": len(inv), "전표": len(slip),
                      "계산서만": len(g["계산서만"]), "전표만": len(g["전표만"]),
                      "금액다름": len(g["금액다름"]), "미발행": len(g["미발행"])})
        con.commit()
    finally:
        con.close()

    out = a.out or os.path.join(ROOT, "reports", "매출전표_계산서_대조.md")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    open(out, "w", encoding="utf-8").write(md)

    print("매출 전표 ↔ 계산서 대조 (%s ~ %s)" % (since, until))
    print("  계산서 %d건 · 매출전표 %d건" % (len(inv), len(slip)))
    print("  ① 미발행 %d건 · ② 계산서만 %d건 · ③ 전표만 %d건 · ④ 금액다름 %d건"
          % (len(g["미발행"]), len(g["계산서만"]), len(g["전표만"]), len(g["금액다름"])))
    for m, x, y in tbl:
        if x != y:
            print("     ★ %s — 계산서 %d · 전표 %d" % (m, x, y))
    print("  → %s" % os.path.relpath(out, ROOT))
    return 0


if __name__ == "__main__":
    sys.exit(main())
