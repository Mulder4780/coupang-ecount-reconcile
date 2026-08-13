# -*- coding: utf-8 -*-
"""
typo_watch.py — 오타·오기입 탐지 (읽기 전용)
=============================================
사용자 지시(2026-08-09): "오타 오기입 이런거 잡아낼 수 있는 알고리즘 구성해서 적용해"

원장에 사람이 손으로 적는 칸들이 있다 — PO번호·프로젝트NO·금액·날짜·캠프명.
손으로 적는 곳에는 반드시 오타가 생기고, **오타는 빈칸과 달리 눈에 띄지 않는다.**
`PO372139` 를 `PO372936` 이라 적어도 화면은 멀쩡히 번호를 보여 준다.

## 무엇을 근거로 잡나 — 지어내지 않는다
정답 목록이 **따로 있는 것만** 본다. 목록 없이 "이상해 보인다"로 지목하지 않는다.
  · PO번호      ← 쿠팡 PO목록 엑셀 (실재하는 번호의 전부)
  · 프로젝트NO   ← ERP 판매조회 색인
  · 캠프명      ← 원장에서 **자주 쓰인** 이름들(3회 이상 = 사람이 반복해 쓴 정본)
  · 금액        ← ERP·명세서 금액과의 **자릿수 관계**(10배·100배·1/10)
  · 날짜        ← 오늘·완료일과의 앞뒤 관계, 연도 한 자리 차이

## ★ 어려운 것은 잡는 게 아니라 잘못 지목하지 않는 것이다
오타라고 잘못 부르면 사람이 **멀쩡한 값을 고치러 간다** — 못 잡는 것보다 나쁘다.
그래서 문을 셋 건다:
  ① **후보가 유일할 때만** 지목한다. 편집거리 1인 실재 번호가 둘이면 어느 것인지
     모르는 것이다 — 그때는 '비슷한 것 여럿'으로만 적고 지목하지 않는다.
  ② **길이가 같을 때만** 번호 오타로 본다. 자리 수가 다르면 오타가 아니라 다른 체계다.
  ③ **절대 자동으로 고치지 않는다.** 큐에도 넣지 않는다. 원장 오타는 '무엇이 맞나'를
     사람만 안다 — 자동으로 고치면 그때 정말 뭐라고 적혀 있었는지를 잃는다.

## 왜 회차인가
오기입은 매일 새로 생긴다. 한 번 훑고 끝내면 그다음 날부터 낡는데,
**틀린 값은 비어 있지 않아서 어느 화면에도 티가 안 난다.**

실행:  python typo_watch.py            # 리포트만 (reports/오기입_확인.md)
       python typo_watch.py --json     # 집계 한 줄(JSON)
"""
import sys, os, re, json
from datetime import datetime, date
from collections import Counter, defaultdict

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
OUT_MD = os.path.join(REPORT_DIR, "오기입_확인.md")
OUT_JSON = os.path.join(REPORT_DIR, "오기입_확인.json")


def edit1(a, b):
    """편집거리가 1 이하인가 — **길이가 같을 때만** 본다(치환 1자리).

    자리 수가 다른 것은 오타가 아니라 다른 번호 체계일 때가 많다. 삽입·삭제까지
    받아 주면 `PO37213` 같은 잘린 값이 `PO372139` 의 오타로 잡혀 오탐이 는다.
    """
    if len(a) != len(b) or a == b:
        return False
    diff = 0
    for x, y in zip(a, b):
        if x != y:
            diff += 1
            if diff > 1:
                return False
    return diff == 1


def swap_typo(a, b):
    """이웃한 두 글자가 자리를 바꾼 것인가 — 실제로 가장 흔한 손오타다(3712 → 3721)."""
    if len(a) != len(b) or a == b:
        return False
    d = [i for i, (x, y) in enumerate(zip(a, b)) if x != y]
    return len(d) == 2 and d[1] == d[0] + 1 and a[d[0]] == b[d[1]] and a[d[1]] == b[d[0]]


def near(value, known):
    """`value` 와 한 글자/자리바꿈 차이인 실재 값들. 지목은 **유일할 때만** 한다."""
    return sorted(k for k in known if edit1(value, k) or swap_typo(value, k))


def digit_slip(a, b):
    """자릿수 실수인가 — 0 하나 더 치거나 덜 친 것. `배수` 를 돌려준다."""
    if not a or not b:
        return None
    for m in (10, 100, 1000):
        if a == b * m:
            return f"{m}배"
        if b == a * m:
            return f"1/{m}"
    return None


def scan(master=None):
    from ecount_reconcile import read_ledger, load_config
    cfg = load_config()
    master = master or cfg["reconcile"]["master_xlsx"]
    recs = read_ledger(master)

    # ── 정답 목록을 모은다(없으면 그 검사만 건너뛴다 — 지어내지 않는다) ──────────
    po_known = set()
    try:
        import po_reconcile as P
        from inbox_scan import pick
        for f in pick("po"):
            po_known |= {p["po"] for p in P.parse_po_export(f)}
    except Exception as e:
        print(f"  (쿠팡 PO목록 없음 — PO번호 검사 건너뜀: {e})")

    erp = {}
    try:
        import erp_sales_index
        erp = erp_sales_index.build()[0]
    except Exception as e:
        print(f"  (ERP 색인 없음 — 프로젝트NO·금액 검사 건너뜀: {e})")
    prj_known = set(erp)

    # 캠프명 정본은 원장 자신이 안다 — 3회 이상 쓰인 이름이 사람이 반복해 쓴 정본이다.
    camp_count = Counter(str(r.get("캠프명") or "").strip()
                         for r in recs.values() if r.get("캠프명"))
    camp_known = {c for c, n in camp_count.items() if n >= 3}

    today = date.today().isoformat()
    hits = defaultdict(list)

    for sid, r in sorted(recs.items()):
        prj = str(r.get("프로젝트NO") or "").strip()

        # ① PO번호 — 쿠팡 목록에 없는데 한 글자 차이로 실재하는 번호가 **하나** 있다
        po = str(r.get("원장_PO번호") or "").strip()
        # ★ 규칙은 한 곳에서 온다 — po_reconcile.norm_po (= PO_PAT). 여기에 비슷한 정규식을
        #   따로 적어 두면 언젠가 갈리고, 갈린 뒤에는 대조기와 오기입 감시가 같은 값을
        #   두고 서로 다르게 답한다(화면도 같은 규칙을 본다 — [252]).
        try:
            po = P.norm_po(po)
        except Exception:
            po = ""   # 목록을 못 읽어 P 가 없으면 이 검사만 건너뛴다(위에서 이미 알렸다)
        if po and po_known and po not in po_known:
            cand = near(po, po_known)
            if len(cand) == 1:
                hits["PO번호"].append({"정산ID": sid, "프로젝트NO": prj, "적힌값": po,
                                       "실재값": cand[0], "근거": "쿠팡 PO목록에 한 글자 차이로 있음"})
            elif cand:
                hits["PO번호(모호)"].append({"정산ID": sid, "적힌값": po,
                                           "비슷한값": ",".join(cand[:4]),
                                           "근거": "비슷한 번호가 여럿 — 어느 것인지 알 수 없음"})

        # ② 프로젝트NO — ERP 에 없는데 한 글자 차이로 실재하는 번호가 **하나**
        if prj and prj_known and prj not in prj_known:
            cand = near(prj, prj_known)
            if len(cand) == 1:
                hits["프로젝트NO"].append({"정산ID": sid, "적힌값": prj, "실재값": cand[0],
                                        "근거": "ERP 판매조회에 한 글자 차이로 있음"})

        # ③ 금액 자릿수 — ERP 와 **정확히** 10·100·1000배 관계면 0 개수 실수다
        led = r.get("원장_공급가액")
        ref = (erp.get(prj) or {}).get("supply")
        slip = digit_slip(int(led) if led else 0, int(ref) if ref else 0)
        if slip:
            hits["금액 자릿수"].append({"정산ID": sid, "프로젝트NO": prj,
                                    "원장": f"{int(led):,}", "ERP": f"{int(ref):,}",
                                    "근거": f"ERP 금액의 {slip} — 0 개수 실수로 보입니다"})

        # ④ 날짜 — 미래이거나, 완료보다 앞선 발행/입금이거나, 연도만 한 자리 다르다
        done = str(r.get("작업완료일") or "")[:10]
        for col, label in (("원장_거래명세서발행일", "명세서발행일"),
                           ("원장_세금계산서발행일", "계산서발행일"),
                           ("원장_입금일", "입금일")):
            v = str(r.get(col) or "")[:10]
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", v):
                continue
            if v > today:
                hits["날짜"].append({"정산ID": sid, "칸": label, "적힌값": v,
                                   "근거": f"오늘({today})보다 뒤 — 연도·월 오기입으로 보입니다"})
            elif done and v < done:
                # ★ 첫판에서 여기가 160건을 쏟아냈다(2026-08-09). 전부 명세서발행일이
                #   완료일보다 **1~3일** 앞선 것이었는데, 그건 오타가 아니라 **정상 업무**다
                #   — 거래명세서를 먼저 끊고 작업을 마치는 일이 흔하다.
                #   경보가 750행 중 160건이면 그 경보는 아무도 안 본다(오늘 세운 규칙).
                #   그래서 **며칠 앞선 것은 안 본다.** 진짜 오타는 자릿수로 튄다:
                #     · 연도가 한 자리 다르다(2025 ↔ 2026) — 해 바뀔 때 가장 흔하다
                #     · 1년 넘게 앞선다 — 며칠 차이로는 설명이 안 된다
                yr_slip = v[4:] >= done[4:] and edit1(v[:4], done[:4])
                gap = (date.fromisoformat(done) - date.fromisoformat(v)).days
                if yr_slip:
                    hits["날짜"].append({
                        "정산ID": sid, "칸": label, "적힌값": v, "완료일": done,
                        "근거": f"연도 한 자리 차이({v[:4]} vs {done[:4]}) — 해 바뀔 때 흔한 오타"})
                elif gap >= 365:
                    hits["날짜"].append({
                        "정산ID": sid, "칸": label, "적힌값": v, "완료일": done,
                        "근거": f"완료일보다 {gap}일 앞섭니다 — 며칠 차이로 설명되지 않습니다"})

        # ⑤ 캠프명 — 자주 쓰인 이름과 한 글자 차이인데 **한 번만** 쓰였다
        camp = str(r.get("캠프명") or "").strip()
        # 캠프명처럼 생긴 것만 본다 — `)` `0` 같은 부스러기는 오타가 아니라 **빈 칸 대신
        # 들어간 쓰레기**다. 그걸 "0을 )로 잘못 적었다"고 지목하면 사람이 헛일을 한다.
        if not re.search(r"[가-힣]{2,}", camp):
            camp = ""
        # ★ '1번만 쓰인 비슷한 이름'만으로는 약하다(2026-08-09 실측). `제주3캠프` 가
        #   `양주3캠프` 의 오타로 잡혔는데, 제주에는 제주1·2·3캠프가 **다 실재한다.**
        #   드물게 쓰인 것과 잘못 쓰인 것은 다르다.
        #   진짜 신호는 **괄호 안 지명이 같은데 앞 숫자만 다른 것**이다:
        #   `송파1MB(감일동)` 1번 ↔ `송파5MB(감일동)` 4번 — 감일동에 송파1MB 는 없다.
        #   괄호가 없는 이름(`제주3캠프`)은 이 규칙에 아예 안 걸린다 — 그래서 안전하다.
        if camp and camp_count.get(camp, 0) == 1:
            loc = re.search(r"\(([^)]+)\)$", camp)
            if loc:
                same_loc = [c for c in camp_known
                            if c.endswith("(" + loc.group(1) + ")") and edit1(camp, c)]
                if len(same_loc) == 1:
                    hits["캠프명"].append({
                        "정산ID": sid, "적힌값": camp, "실재값": same_loc[0],
                        "근거": f"같은 '{loc.group(1)}' 에 '{same_loc[0]}' 가 "
                                f"{camp_count[same_loc[0]]}번 · 이 이름은 1번뿐"})

    return hits, len(recs)


def write_report(hits, total):
    os.makedirs(REPORT_DIR, exist_ok=True)
    n = sum(len(v) for v in hits.values())
    lines = ["# 오기입 확인", "",
             f"- 생성 {datetime.now():%Y-%m-%d %H:%M} · 원장 {total}행 · 의심 **{n}건**",
             "- **이 도구는 아무것도 고치지 않습니다.** 무엇이 맞는지는 사람만 압니다.",
             "- 후보가 여럿이면 지목하지 않습니다 — 잘못 지목하면 멀쩡한 값을 고치러 갑니다.", ""]
    if not n:
        lines.append("의심되는 오기입이 없습니다.")
    for kind, rows in sorted(hits.items()):
        lines.append(f"## {kind} — {len(rows)}건")
        keys = list(rows[0].keys())
        lines.append("| " + " | ".join(keys) + " |")
        lines.append("|" + "---|" * len(keys))
        for r in rows[:200]:
            lines.append("| " + " | ".join(str(r.get(k, "")) for k in keys) + " |")
        if len(rows) > 200:
            lines.append(f"\n… 외 {len(rows) - 200}건 (전체는 오기입_확인.json)")
        lines.append("")
    with open(OUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({"생성": datetime.now().isoformat(timespec="seconds"),
                   "원장행": total, "의심": n,
                   "종류별": {k: len(v) for k, v in hits.items()},
                   "상세": hits}, f, ensure_ascii=False, indent=1)
    return n


def main():
    args = sys.argv[1:]
    master = args[args.index("--master") + 1] if "--master" in args else None
    hits, total = scan(master)
    n = write_report(hits, total)
    if "--json" in args:
        print(json.dumps({"의심": n, "종류별": {k: len(v) for k, v in hits.items()}},
                         ensure_ascii=False))
    else:
        print(f"오기입 의심 {n}건 / 원장 {total}행 — " +
              " · ".join(f"{k} {len(v)}" for k, v in sorted(hits.items())))
        print("리포트:", OUT_MD)
    return 0


if __name__ == "__main__":
    sys.exit(main())
