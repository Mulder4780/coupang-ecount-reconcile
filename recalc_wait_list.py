# -*- coding: utf-8 -*-
"""recalc_wait_list.py — **금액이 아직 확정되지 않은 정산 건** 목록을 뽑는다.

왜 이 파일이 따로 생겼나 (2026-08-08)
  인계 문서에 "금액 재계산 대기 37건 목록 뽑기"가 오래 남아 있었다. 그런데 그 말대로
  06시트 `청구상태` 를 세어 보면 **0건**이다 — 그 상태값은 원장에서 사라졌다.
  같은 날 ÷1.1 환산이 들어오면서 상태가 대부분 정리됐기 때문이다(quote_mismatch.py 주석).

  여기서 "0건이니 끝났다"고 적으면 **조용한 거짓말**이 된다. 상태 딱지가 없어졌을 뿐,
  `실제작업공급가액` 이 비어 있어 금액이 거래명세서에만 기대고 있는 건은 그대로 남아 있다.
  그래서 이 도구는 **딱지가 아니라 사실**로 고른다:

      실제작업공급가액이 비었다 AND 거래명세서합계가 있다 AND 객관완료로 입증되지 않았다

  `recalc_pending.py` 와 헷갈리지 말 것 — 그쪽은 "엑셀이 아직 수식을 계산 안 해서
  앱에 안 보이는 행"을 센다(화면 문제). 이쪽은 "금액 자체가 아직 안 정해진 건"이다(업무 문제).

읽기 전용이다. 원장에 아무것도 쓰지 않는다.

사용
  python recalc_wait_list.py            # reports/금액미확정_목록.md + .json 갱신
  python recalc_wait_list.py --print    # 사람이 읽는 한 줄
"""
import sys, os, io, json, collections
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

MD = os.path.join(ROOT, "reports", "금액미확정_목록.md")
JS = os.path.join(ROOT, "reports", "금액미확정_목록.json")


def collect():
    import ecount_reconcile as R
    recs = R.read_ledger(R.resolve_master(R.load_config()["reconcile"]["master_xlsx"]))
    try:
        import ledger_db
        res = ledger_db.resolutions() or {}
    except Exception:
        res = {}                       # DB 를 못 읽어도 목록은 나와야 한다(더 넓게 잡힐 뿐)
    out = []
    for sid, r in sorted(recs.items()):
        if r.get("원장_공급가액"):
            continue                   # 금액이 이미 확정됐다
        if not r.get("원장_거래명세서합계"):
            continue                   # 명세서도 없으면 '금액 대기'가 아니라 그냥 미착수다
        if str((res.get(sid) or {}).get("status") or "").startswith("완료("):
            continue                   # 다른 원천이 객관으로 입증한 건
        out.append({
            "정산ID": sid,
            "업무구분": r.get("업무구분") or "",
            "프로젝트NO": r.get("프로젝트NO") or "",
            "캠프명": r.get("캠프명") or "",
            "작업완료일": str(r.get("작업완료일") or "")[:10],
            "명세서합계": int(r.get("원장_거래명세서합계") or 0),
            "PO번호": r.get("원장_PO번호") or "",
        })
    return out


def to_md(rows):
    by_kind = collections.Counter(r["업무구분"] or "(빈칸)" for r in rows)
    by_month = collections.Counter((r["작업완료일"] or "")[:7] or "(날짜없음)" for r in rows)
    total = sum(r["명세서합계"] for r in rows)
    L = ["# 금액이 아직 확정되지 않은 정산 건",
         "",
         "- 만든 시각 %s · **%d건 · 명세서합계 %s원**" % (
             datetime.now().strftime("%Y-%m-%d %H:%M"), len(rows), format(total, ",")),
         "- 고른 기준: `실제작업공급가액`이 비어 있고 `거래명세서합계`는 있는 건"
         " (객관완료로 입증된 건은 뺀다).",
         "- **06시트 `청구상태`의 '금액 재계산 대기' 딱지로 고르지 않는다** — 그 값은"
         " 원장에서 사라졌고(0건), 딱지가 없어진 것과 금액이 정해진 것은 다른 말이다.",
         "- 이 문서는 목록일 뿐이다. 금액 확정은 사람이 한다.",
         "",
         "## 업무구분",
         "", "| 구분 | 건수 |", "|---|---:|"]
    L += ["| %s | %d |" % (k, n) for k, n in by_kind.most_common()]
    L += ["", "## 작업완료월", "", "| 월 | 건수 |", "|---|---:|"]
    L += ["| %s | %d |" % (k, n) for k, n in sorted(by_month.items(), reverse=True)]
    L += ["", "## 전체 목록", "",
          "| 정산ID | 구분 | 프로젝트NO | 캠프 | 완료일 | 명세서합계 | PO |",
          "|---|---|---|---|---|---:|---|"]
    for r in rows:
        L.append("| %s | %s | %s | %s | %s | %s | %s |" % (
            r["정산ID"], r["업무구분"], r["프로젝트NO"], r["캠프명"][:22],
            r["작업완료일"], format(r["명세서합계"], ","), r["PO번호"]))
    return "\n".join(L) + "\n"


def main():
    rows = collect()
    if "--print" in sys.argv:
        print("금액 미확정 %d건 · 명세서합계 %s원"
              % (len(rows), format(sum(r["명세서합계"] for r in rows), ",")))
        return 0
    os.makedirs(os.path.dirname(MD), exist_ok=True)
    io.open(MD, "w", encoding="utf-8").write(to_md(rows))
    json.dump({"만든시각": datetime.now().isoformat(timespec="seconds"),
               "건수": len(rows), "항목": rows},
              io.open(JS, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("금액 미확정 %d건 → %s" % (len(rows), os.path.relpath(MD, ROOT)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
