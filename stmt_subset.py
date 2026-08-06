# -*- coding: utf-8 -*-
"""
stmt_subset.py — 명세서 한 장이 **판매행 여러 개를 묶은 것**을 찾아 잇는다 (읽기 전용)

사용자 지시(2026-08-06): "매칭이 안된 데이터 다른 방법으로 매칭 방법 찾아서 해결해,
자료가 있는데 매칭이 안되는 경우가 많아."

왜 이 방법인가 — 원인부터 셌다(짝 없는 쿠팡 현장 명세서 141건):
    130건  명세서 1장 = 판매행 여러 개의 합   ← 92%
      5건  그 거래처 후보가 UJ 하나뿐(금액은 다름)
      3건  후보는 있으나 금액·조합 모두 안 맞음
      3건  금액 같은 행이 있는데 날짜가 멀다
`stmt_link` 는 "금액합계가 판매행 하나와 정확히 같을 때"만 잇는다. 그런데 현장은
프로젝트 두셋을 한 장에 묶어 청구하고 있었다. 그래서 **조합의 합**으로 잇는다.

  서초1Sub-hub 707,300 = UJ2600089(443,300) + UJ2600084(264,000)

무엇을 지키나 (이 프로젝트의 '추측 금지' 원칙)
  · 합이 맞는 조합이 **둘 이상이면 확정하지 않는다**(모호로 남긴다).
  · **이미 다른 명세서에 붙은 UJ 는 쓰지 않는다** — 한 프로젝트가 두 장에 이중으로
    붙는 것을 막는다. 실측: 제외 없이 133건 / 제외하면 124건, 모호는 양쪽 다 0건.
  · 거래처(캠프)가 같고 전표일 ±45일 안에서만 본다. 멀리서 우연히 맞는 조합을 줄인다.

어디서 도나 — **`stmt_link` 가 마지막 단계로 이 모듈의 `find()` 를 부른다.**
  그래서 daily_run 에 따로 단계를 넣지 않는다. 넣으면 판매조회 엑셀 수십 개를
  한 회차에 두 번 읽게 되고, 그 스캔이 이 파이프라인에서 가장 무거운 축이다.
  확정 결과와 근거(어느 UJ 를 얼마씩 더했는지)는 `명세서_프로젝트_매칭.json` 의
  각 항목 `parts` 에 그대로 들어간다.
  이 파일을 직접 돌리는 것은 **따로 들여다볼 때**를 위한 것이다.

산출: reports/명세서_부분합_매칭.md / .json   (엑셀에 쓰지 않는다 — DB·리포트가 정본)
  python stmt_subset.py
"""
import collections
import itertools
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT_JSON = os.path.join(ROOT, "reports", "명세서_부분합_매칭.json")
OUT_MD = os.path.join(ROOT, "reports", "명세서_부분합_매칭.md")

MAX_N = 4        # 명세서 한 장에 프로젝트 몇 개까지 묶이나 (실측 대부분 2~3)
POOL = 18        # 조합을 찾을 후보 판매행 수 상한 (C(18,4)=3060 — 넉넉하고 빠르다)
WINDOW = 45      # 전표일 ±며칠 안의 판매행만 본다
# 쿠팡 현장으로 보이는 거래처. 여기 안 걸리는 곳(물류사·제조사)은 UJ 프로젝트가 없어
# 짝이 없는 것이 정상이므로 아예 손대지 않는다.
CAMPLIKE = re.compile(r"캠프|MB|Sub-?hub|Sub-?FC|허브|HUB|FC|물류센터|센터|터미널", re.I)


def dnum(d):
    """'2026/01/26' → 20260126. 구분자를 믿지 않고 숫자만 뽑는다."""
    p = re.findall(r"\d+", str(d or ""))
    return int("".join(p)) if len(p) >= 3 else 0


def norm_cust(s):
    """거래처명 비교용 정규화 — **캠프명 정규화와 같은 규칙을 쓴다**(2026-08-06).

    예전엔 여기만 따로 괄호·공백만 지웠다. 그래서 붙여넣기로 들어온
    `&amp;`·`<-메모`·`?` 가 남아 같은 거래처가 둘로 갈렸고, 조합을 찾을 후보
    자체가 비어 부분합 매칭이 시작도 못 했다. 규칙이 두 벌이면 반드시 갈린다.
    """
    try:
        from customer_index import clean
        s = clean(s)
    except Exception:                    # 색인 모듈이 없어도 매칭은 계속돼야 한다
        pass
    return re.sub(r"[\s()（）\-_/?]", "", str(s or "")).lower()


def find(docs, sales, taken_uj, unmatched_slips):
    """짝 없는 명세서마다 조합을 찾는다. 확정·모호를 갈라 돌려준다."""
    by_cust = collections.defaultdict(list)
    for s in sales:
        if s.get("uj") and s.get("total"):
            by_cust[norm_cust(s.get("cust"))].append(s)

    linked, ambiguous, still = [], [], []
    for d in docs:
        if d["slip"] not in unmatched_slips:
            continue
        if not CAMPLIKE.search(d.get("cust") or ""):
            continue
        c = norm_cust(d.get("cust"))
        rows = by_cust.get(c)
        if not rows:                       # 이름이 조금 다를 수 있다 — 포함관계로 한 번 더
            rows = [s for k, v in by_cust.items() if c and (c in k or k in c) for s in v]
        sd = dnum(d["slip"].split("-")[0])
        pool = [s for s in rows
                if abs(dnum(s["date"]) - sd) <= WINDOW and s["uj"] not in taken_uj]
        pool.sort(key=lambda s: abs(dnum(s["date"]) - sd))     # 가까운 날짜 우선
        pool = pool[:POOL]

        found = {}
        for n in range(2, MAX_N + 1):
            for combo in itertools.combinations(pool, n):
                if sum(x["total"] for x in combo) == d["amount"]:
                    key = tuple(sorted({x["uj"] for x in combo}))
                    found.setdefault(key, combo)
        if len(found) == 1:
            key, combo = next(iter(found.items()))
            linked.append({
                "slip": d["slip"], "cust": d.get("cust", ""), "amount": d["amount"],
                "ujs": list(key), "n": len(key),
                "states": sorted({x.get("state", "") for x in combo}),
                "parts": [{"uj": x["uj"], "total": x["total"], "date": x["date"]}
                          for x in combo],
                "how": f"부분합 {len(key)}건",
            })
        elif found:
            ambiguous.append({"slip": d["slip"], "cust": d.get("cust", ""),
                              "amount": d["amount"],
                              "candidates": [list(k) for k in list(found)[:4]]})
        else:
            still.append({"slip": d["slip"], "cust": d.get("cust", ""),
                          "amount": d["amount"], "pool": len(pool)})
    return linked, ambiguous, still


def main():
    reports = os.path.join(ROOT, "reports")
    try:
        docs = json.load(open(os.path.join(reports, "거래명세서_상세.json"),
                              encoding="utf-8"))["docs"]
        mat = json.load(open(os.path.join(reports, "명세서_프로젝트_매칭.json"),
                             encoding="utf-8"))
    except Exception as exc:
        print(f"근거 파일을 읽지 못했습니다 — stmt_docs·stmt_link 를 먼저 돌리세요: {exc}")
        return 1

    unmatched = {x["slip"] for x in mat.get("unmatched", [])}
    taken = set()
    for x in mat.get("linked", []):
        for u in str(x.get("uj") or "").split(";"):
            if u:
                taken.add(u)

    import stmt_link
    sales, srcs = stmt_link.load_sales()
    if not sales:
        print("판매조회 원본을 찾지 못함 — 중단")
        return 1

    linked, ambiguous, still = find(docs, sales, taken, unmatched)
    amt = sum(x["amount"] for x in linked)

    os.makedirs(reports, exist_ok=True)
    json.dump({"sales_src": srcs, "linked": linked, "ambiguous": ambiguous,
               "still": still, "rule": {"max_n": MAX_N, "window": WINDOW,
                                        "pool": POOL, "exclude_taken": True}},
              open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)

    n_cnt = collections.Counter(x["n"] for x in linked)
    L = ["# 거래명세서 — 판매행 여러 개를 묶은 건(부분합) 매칭", "",
         f"- 대상: `stmt_link` 가 짝을 못 찾은 명세서 중 **쿠팡 현장** 건",
         f"- **확정 {len(linked)}건 / {amt:,}원** · 모호 {len(ambiguous)}건 · "
         f"여전히 못 찾음 {len(still)}건",
         f"- 규칙: 같은 거래처 · 전표일 ±{WINDOW}일 · 최대 {MAX_N}건 조합 · "
         "합이 맞는 조합이 **유일할 때만** 확정 · 이미 다른 명세서에 붙은 UJ 는 제외",
         "- 묶인 건수: " + ", ".join(f"{k}개 묶음 {v}건" for k, v in sorted(n_cnt.items())),
         "", "## 확정", "", "| 전표번호 | 거래처(캠프) | 금액 | 프로젝트 | 내역 |",
         "|---|---|---:|---|---|"]
    for x in sorted(linked, key=lambda r: -r["amount"]):
        parts = " + ".join(f"{p['uj']} {p['total']:,}" for p in x["parts"])
        L.append(f"| {x['slip']} | {x['cust'][:22]} | {x['amount']:,} | "
                 f"{'+'.join(x['ujs'])} | {parts} |")
    if ambiguous:
        L += ["", f"## 모호 — 조합이 둘 이상이라 확정하지 않음 ({len(ambiguous)}건)", "",
              "| 전표번호 | 거래처 | 금액 | 후보 |", "|---|---|---:|---|"]
        for x in ambiguous:
            L.append(f"| {x['slip']} | {x['cust'][:22]} | {x['amount']:,} | "
                     + " / ".join("+".join(c) for c in x["candidates"]) + " |")
    if still:
        L += ["", f"## 여전히 못 찾음 ({len(still)}건) — 판매조회에 근거가 없다", "",
              "| 전표번호 | 거래처 | 금액 | 후보 판매행 |", "|---|---|---:|---:|"]
        for x in sorted(still, key=lambda r: -r["amount"])[:30]:
            L.append(f"| {x['slip']} | {x['cust'][:22]} | {x['amount']:,} | {x['pool']} |")
    L += ["", "> 읽기 전용 근거다. 원장·완료 처리에 쓰려면 후속 도구가 이 파일을 인용한다."]
    open(OUT_MD, "w", encoding="utf-8").write("\n".join(L) + "\n")

    print(f"부분합 매칭: 확정 {len(linked)}건 {amt:,}원 · 모호 {len(ambiguous)} · "
          f"못 찾음 {len(still)} → reports/명세서_부분합_매칭.md")
    return 0


if __name__ == "__main__":
    sys.exit(main())
