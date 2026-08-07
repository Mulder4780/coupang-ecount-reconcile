# -*- coding: utf-8 -*-
"""
pending_report.py — **아직 반영 안 된 것**을 한 장으로 모은다 (읽기 전용)

사용자 지시(2026-08-05): "지금 반영 안된 자료들 목록 정리 · 원본 데이터들 포함."

왜 도구로 만드나: 지금 이 답은 리포트 열 개(반영대기·명세서매칭·세금계산서경과·
견적명세·거래처코드·원본색인…)에 흩어져 있어, 물어볼 때마다 사람이 열 군데를 뒤져야 한다.
그리고 이 질문은 **또 나온다**. 그래서 한 번 세고 한 장으로 굳힌다.

무엇을 '미반영'으로 보나 — 네 갈래로 나눈다. 갈래가 곧 **누가 움직여야 하는가**다.
  A. 기다리면 자동으로 들어가는 것   → 아무도 안 해도 된다(다음 11:00·15:00 회차)
  B. 자료는 있는데 시스템이 못 잇는 것 → AI/사람이 근거를 더 대야 한다
  C. 아직 받지 못한 원본             → 받아야 한다(일부는 사람이 화면을 열어 줘야)
  D. 사람이 해야 반영되는 것          → 발행·확인처럼 시스템이 대신 못 하는 일

산출: reports/미반영_목록.md / .json
  python pending_report.py
  python pending_report.py --print     # 콘솔 요약만
"""
import json
import os
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
REPORTS = os.path.join(ROOT, "reports")
OUT_MD = os.path.join(REPORTS, "미반영_목록.md")
OUT_JSON = os.path.join(REPORTS, "미반영_목록.json")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def j(name, default=None):
    """리포트 하나를 읽는다. 없으면 조용히 넘긴다 — 한 갈래가 비어도 나머지는 보여야 한다."""
    try:
        with open(os.path.join(REPORTS, name), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default if default is not None else {}


def cnt(v):
    """리포트마다 같은 뜻을 숫자로도, 목록으로도 적어 둔다 — 어느 쪽이든 건수로 받는다."""
    if isinstance(v, (list, dict)):
        return len(v)
    try:
        return int(v or 0)
    except (TypeError, ValueError):
        return 0


def rows_of(d, *keys):
    for k in keys:
        v = d.get(k)
        if isinstance(v, list):
            return v
    return []


def collect():
    out = []

    def add(group, what, n, note="", amount=0, where="", who=""):
        n = cnt(n)
        if not n:
            return
        out.append({"group": group, "what": what, "n": n, "note": note,
                    "amount": amount, "where": where, "who": who})

    # ── A. 기다리면 들어간다 ────────────────────────────────────────────────
    q = j("반영대기.json")
    add("A. 기다리면 자동 반영", "입력 큐(DB에 있음)", q.get("대기", 0),
        f"다음 회차 {q.get('다음반영', '')}", where="db/ledger_queue.db", who="자동")
    add("A. 기다리면 자동 반영", "19시트 인수인계 예약", q.get("인수인계대기", 0),
        "같은 회차에 함께 들어간다", where="db/ledger_queue.db", who="자동")
    for slot in (q.get("밀린회차") or []):
        add("A. 기다리면 자동 반영", f"밀린 회차 {slot}", 1,
            "놓친 회차는 다음 허용 시각에 함께 처리된다", who="자동")

    # ── B. 자료는 있는데 못 잇는 것 ─────────────────────────────────────────
    m = j("명세서_프로젝트_매칭.json")
    un = rows_of(m, "unmatched")
    try:
        import re
        camp = [x for x in un if re.search(r"캠프|MB|Sub-?hub|Sub-?FC|허브|HUB|FC|센터|터미널",
                                           str(x.get("cust") or ""), re.I)]
    except Exception:
        camp = un
    add("B. 못 잇는 자료", "거래명세서 ↔ 프로젝트(UJ) 짝 없음(쿠팡 현장)", len(camp),
        "판매조회에 같은 금액·일자가 없다. 분할 청구·미등록 둘 다 가능",
        amount=sum(x.get("amount") or 0 for x in camp),
        where="reports/명세서_프로젝트_매칭.md", who="AI/사람")
    add("B. 못 잇는 자료", "거래명세서 짝 없음(쿠팡 아닌 거래처)", len(un) - len(camp),
        "UJ 프로젝트가 없는 거래 — 짝이 없는 것이 정상이다",
        amount=sum(x.get("amount") or 0 for x in un) -
               sum(x.get("amount") or 0 for x in camp),
        where="reports/명세서_프로젝트_매칭.md", who="참고")

    qm = j("견적명세_불일치.json")
    add("B. 못 잇는 자료", "견적 ↔ 명세 불일치", qm.get("total", 0),
        "명세합계가 다른 프로젝트 견적과 맞는 입력 밀림 의심",
        where="reports/견적명세_불일치.md", who="사람")

    ci = j("거래처코드_색인.json")
    add("B. 못 잇는 자료", "캠프 ↔ 거래처코드 미매칭", ci.get("unmatched", 0),
        "ERP 거래처등록에서 짝을 못 찾은 캠프", where="reports/거래처코드_색인.json",
        who="AI/사람")
    add("B. 못 잇는 자료", "캠프 ↔ 거래처코드 모호", ci.get("ambiguous", 0),
        "후보가 둘 이상 — 추측하지 않고 남겨 둔다", where="reports/거래처코드_색인.json",
        who="사람")
    # 캠프명 칸에 '0'·'...' 같은 표시가 들어온 건. 시스템이 짝을 찾을 수 없으니
    # '못 잇는 자료'가 아니라 **원장을 고쳐야 하는 일**로 따로 센다(2026-08-06).
    junk = ci.get("입력오류", 0)
    junk_n = sum(x.get("건수", 1) for x in junk) if isinstance(junk, list) else junk
    add("D. 사람이 해야 함", "원장 캠프명이 값 대신 표시(0·… 등)", junk_n,
        "캠프명 칸을 실제 이름으로 고쳐야 매출·서류가 붙는다",
        where="reports/거래처코드_색인.json", who="사람(원장 수정)")

    cm = j("캠프마스터.json")
    # 출처는 'ERP+원장+밴드' 처럼 더해서 적는다 — ERP 글자가 없으면 ERP 에 자리가 없는 캠프다.
    only = [r for r in rows_of(cm, "rows") if "ERP" not in str(r.get("출처") or r.get("source") or "")]
    add("B. 못 잇는 자료", "ERP에 없는 캠프(원장·밴드에만 있음)", len(only),
        "ERP 거래처등록에 자리가 없어 매출이 안 붙는다", where="reports/캠프마스터.json",
        who="사람")

    # ── C. 아직 못 받은 원본 ────────────────────────────────────────────────
    board = j("worksplit.json")
    for it in (board.get("items") or []):
        if it.get("state") in ("완료",):
            continue
        t = str(it.get("title") or "")
        if any(k in t for k in ("수집", "다운로드", "화면")):
            # 제목에 건수가 적혀 있으면 그것을 쓴다 — '1건'으로 세면 304건짜리 일이
            # 작은 일로 보인다.
            import re as _re
            mnum = _re.search(r"([\d,]{2,})\s*건", t)
            add("C. 못 받은 원본", t, int(mnum.group(1).replace(",", "")) if mnum else 1,
                str(it.get("detail") or "")[:90], where="reports/작업분담.md",
                who="사람" if it.get("state") == "사람대기" else "AI")

    idx = j("원본색인.json")
    kinds = {}
    for r in (idx.get("rows") or []):
        kinds[r.get("kind") or "?"] = kinds.get(r.get("kind") or "?", 0) + 1
    add("C. 못 받은 원본", "원본 중 '미분류'", kinds.get("미분류", 0),
        "종류를 못 가른 원본 — 내용 판별 규칙을 늘려야 한다",
        where="0. 원본 자료/9. 미분류", who="AI")
    add("C. 못 받은 원본", "투입함에 남은 것", kinds.get("투입 대기", 0),
        "다음 정리 회차가 정본 폴더로 옮긴다", where="0. 원본 자료/100. 업로드용 자료",
        who="자동")

    # ── D. 사람이 해야 반영되는 것 ──────────────────────────────────────────
    tw = j("세금계산서_미발행_경과.json")
    add("D. 사람이 해야 함", "세금계산서 미발행", tw.get("total", 0),
        "경과 구간별로 " + ", ".join(f"{k} {v}건" for k, v in
                                  (tw.get("buckets") or {}).items()),
        where="reports/세금계산서_미발행_경과.json", who="사람(발행)")

    fx = j("확인필요_집계.json")
    if not fx:
        # findings_export 가 남기는 집계 이름이 다를 수 있어, 없으면 건너뛴다.
        fx = {}
    for k, v in (fx.get("counts") or {}).items():
        add("D. 사람이 해야 함", f"확인 필요 — {k}", v,
            "B 갈래와 같은 건을 다른 눈으로 본 것일 수 있다(중복 주의)",
            # 2026-08-07 지시로 별도 엑셀을 없앴다 — 같은 내용이 관리대장 시트에 있다.
            where="관리대장 23_확인필요현황 시트", who="사람")

    return out


def main():
    items = collect()
    total = sum(x["n"] for x in items)
    groups = {}
    for x in items:
        groups.setdefault(x["group"], []).append(x)

    L = [f"# 아직 반영 안 된 것 (자동 집계 {time.strftime('%Y-%m-%d %H:%M')})", "",
         "갈래는 곧 **누가 움직여야 하는가**다.", "",
         "> ★ **갈래별로 읽고, 전부 더하지 마세요.** 같은 건이 갈래를 달리해 두 번 셀 수",
         "> 있습니다 — 예를 들어 '거래명세서 짝 없음'(B)과 '정산 조치필요'(D)는 같은 프로젝트를",
         "> 다른 눈으로 본 것일 수 있습니다. 각 줄의 근거 파일에서 실제 건을 확인하세요.",
         f"> (단순 합계는 {total}건이지만 이 숫자는 '서로 다른 자료 {total}개'라는 뜻이 아닙니다)",
         "",
         "| 갈래 | 뜻 |", "|---|---|",
         "| A | 기다리면 자동으로 들어간다 — 아무도 안 해도 된다 |",
         "| B | 자료는 있는데 시스템이 못 잇는다 — 근거를 더 대야 한다 |",
         "| C | 아직 받지 못한 원본 — 받아야 한다 |",
         "| D | 사람이 해야 반영된다 — 시스템이 대신 못 한다 |", ""]
    for g in sorted(groups):
        rows = sorted(groups[g], key=lambda x: -x["n"])
        n = sum(r["n"] for r in rows)
        L += [f"## {g} — {n}건", "",
              "| 무엇 | 건수 | 금액 | 누가 | 설명 | 어디서 보나 |",
              "|---|---:|---:|---|---|---|"]
        for r in rows:
            L.append(f"| {r['what']} | {r['n']} | "
                     f"{(format(r['amount'], ',') + '원') if r['amount'] else '-'} | "
                     f"{r['who']} | {r['note']} | `{r['where']}` |")
        L.append("")
    L += ["> 이 표는 **읽기 전용 집계**다. 숫자의 근거는 '어디서 보나' 열의 파일에 있다.",
          "> 갱신: `python ecount/pending_report.py` (daily_run 에도 들어 있다)"]

    os.makedirs(REPORTS, exist_ok=True)
    open(OUT_MD, "w", encoding="utf-8").write("\n".join(L) + "\n")
    json.dump({"at": time.strftime("%Y-%m-%d %H:%M"), "total": total, "items": items},
              open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print(f"미반영 집계 → reports/미반영_목록.md  (단순합 {total}건 — 갈래끼리 겹칠 수 있음)")
    for g in sorted(groups):
        print(f"  {g}: {sum(r['n'] for r in groups[g])}건")
    if "--print" in sys.argv:
        for x in sorted(items, key=lambda x: (x["group"], -x["n"])):
            print(f"    [{x['group'][:1]}] {x['what']:<40} {x['n']:>5}건  {x['who']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
