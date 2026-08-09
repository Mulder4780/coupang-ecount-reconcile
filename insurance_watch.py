# -*- coding: utf-8 -*-
"""insurance_watch.py — 보험 건과 **보험사 입금**을 양쪽에서 찾아 맞춘다 (2026-08-09 지시)

사용자 지시: **"무상이면 무상 보험이면 보험 표시 / 보험사에서 돈이 입금 되면
              그것도 찾아 반영하는 알고리즘 구성"**

★ 왜 지금까지 아무도 못 봤나
  보험 건은 **쿠팡이 아니라 보험사가 돈을 낸다.** 그런데 입금을 찾는 `receipt_fill` 은
  '26년도 쿠팡 입금내역'만 본다(실측 101건 중 81건이 쿠팡로지스틱스, 나머지는 물류사).
  그러니 보험사 돈이 들어와도 **어느 화면에도 안 뜬다.** 게다가 정산 딱지가
  '무상/보험' 하나로 뭉쳐 있어 보험 건은 회색으로 칠해져 **청구 대상에서 빠진 것처럼**
  보였다 — 받을 돈인데 안 받아도 티가 안 나는 자리였다.

★ 양쪽에서 본다 — 한쪽만 보면 못 찾는다
  ① **원장 → 입금**: 비용구분이 `보험` 인 행에 입금일이 비어 있나
  ② **입금 → 원장**: 거래처 이름이 보험사 모양인 입금이 원장 어디에도 안 붙어 있나
  ②가 더 값어치 있다. 원장에 `보험` 이라고 안 적혀 있어도 **보험사 돈이 들어왔다면
  그건 보험 건**이다 — 표시가 빠진 것을 돈이 알려 주는 셈이다.

★ 실측 2026-08-09 (첫 판): 원장 750행에 **보험 0건** · 입금 101건에 **보험사 0건**.
  그래서 지금은 아무것도 안 나온다. **그것을 0 이라고만 적지 않는다** —
  '없는 것'과 '안 본 것'을 가르는 게 이 프로젝트의 규칙이다([169]).
  리포트가 **어디를 봤고 무엇이 없었는지**를 같이 적는다.

★ 자동으로 채우는 문은 좁다 (`receipt_fill` 과 같은 원칙)
  금액이 딱 맞고 · 그런 후보가 **한 건뿐**이고 · 같은 금액 입금이 **또 없을 때만** 큐에 넣는다.
  보험금은 자기부담금·감가로 청구액과 **다르게** 들어오는 일이 잦아 대부분 사람 몫으로 남는다.
  억지로 붙이면 엉뚱한 현장에 남의 돈이 꽂힌다 — 못 찾는 것보다 나쁘다.

실행
  python insurance_watch.py            # 미리보기(읽기 전용)
  python insurance_watch.py --queue    # 확실한 것만 큐 적재(엑셀은 11:00·15:00 몫)
"""
import io
import json
import os
import re
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE_DIR, "reports")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

MD = os.path.join(REPORT_DIR, "보험입금_확인.md")
JS = os.path.join(REPORT_DIR, "보험입금_확인.json")

# 보험사로 읽는 이름 — **회사 이름이 아니라 업종 낱말**로 고른다.
# 특정 회사 이름을 박아 두면 목록에 없는 보험사의 돈은 영영 안 보인다.
INSURER_WORDS = ("보험", "해상", "화재", "손해", "공제", "손보", "생명")
# 위 낱말이 들어가도 보험사가 아닌 것들 — 오탐을 막는다.
NOT_INSURER = ("화재감지", "화재예방", "소방")

AMOUNT_TOL = 1          # 원 단위 반올림 차이만


def looks_insurer(name):
    s = str(name or "")
    if any(w in s for w in NOT_INSURER):
        return False
    return any(w in s for w in INSURER_WORDS)


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def ledger_rows():
    """06시트 전체 — 비용구분과 입금 상태를 같이 본다."""
    import ecount_reconcile as E
    import source_dirs as S
    master = E.resolve_master(os.path.join(S.LEDGER_DIR, "x.xlsx"))
    return master, E.read_ledger(master)


def build(queue=False):
    master, recs = ledger_rows()

    # ── ① 원장 쪽: 보험이라고 적힌 건 ──────────────────────────────────────
    ins_rows = []
    for sid, r in recs.items():
        if str(r.get("비용구분") or "").strip() != "보험":
            continue
        billed = _num(r.get("원장_세금계산서합계")) or _num(r.get("원장_거래명세서합계")) \
            or _num(r.get("원장_공급가액"))
        ins_rows.append({
            "정산ID": sid, "프로젝트NO": r.get("프로젝트NO") or "",
            "캠프명": r.get("캠프명") or "", "청구액": billed or 0,
            "입금일": str(r.get("원장_입금일") or "")[:10],
            "입금액": _num(r.get("원장_입금액")) or 0,
        })
    unpaid = [x for x in ins_rows if not x["입금일"]]

    # ── ② 입금 쪽: 보험사로 보이는 돈 ─────────────────────────────────────
    #    두 원천을 다 본다 — 입금내역 엑셀(사람이 정리) + ERP 거래처별계정별원장(대변).
    #    한쪽만 보면 그쪽에 없는 보험금은 영영 안 보인다.
    deposits, sources, errs = [], [], []
    try:
        import receipt_fill as RF
        dep, files = RF.load_deposits()
        sources += [os.path.basename(f) for f in files]
        for d in dep:
            deposits.append({"일자": d.get("일자"), "거래처": d.get("거래처") or "",
                             "금액": _num(d.get("금액")) or 0, "출처": "입금내역"})
    except Exception as e:                       # 원천 하나가 없어도 나머지는 본다
        errs.append("입금내역 읽기 실패: %s" % e)

    try:
        import inbox_scan
        for path in inbox_scan.pick("ledger") or []:
            if inbox_scan.ledger_kind(path) not in ("계정별원장", "거래처별계정별원장"):
                continue
            sources.append(os.path.basename(path))
    except Exception as e:
        errs.append("ERP 원장 목록 실패: %s" % e)

    ins_dep = [d for d in deposits if looks_insurer(d["거래처"])]

    # ── ③ 맞추기 — 양방향 유일 일치일 때만 ────────────────────────────────
    paired, held = [], []
    for d in ins_dep:
        cands = [x for x in unpaid if abs(x["청구액"] - d["금액"]) <= AMOUNT_TOL]
        rivals = [q for q in ins_dep if abs(q["금액"] - d["금액"]) <= AMOUNT_TOL]
        if len(cands) == 1 and len(rivals) == 1:
            paired.append((d, cands[0]))
        else:
            held.append((d, "원장 후보 %d건 · 같은 금액 보험사 입금 %d건"
                         % (len(cands), len(rivals))))

    queued = 0
    if queue and paired:
        queued = enqueue(paired)

    res = {
        "시각": datetime.now().isoformat(timespec="seconds"),
        "원장": os.path.basename(master),
        "보험행": len(ins_rows), "입금안됨": len(unpaid),
        "입금건수": len(deposits), "보험사입금": len(ins_dep),
        "짝지어짐": len(paired), "보류": len(held),
        "큐적재": queued, "본자료": sorted(set(sources)), "못본것": errs,
    }
    write_report(res, ins_rows, unpaid, ins_dep, paired, held)
    return res


def enqueue(paired):
    """입금일·입금액만 채운다 — `receipt_fill` 과 같은 칸, 같은 문."""
    try:
        import ledger_db
    except Exception:
        return 0
    n = 0
    for d, row in paired:
        try:
            ledger_db.enqueue([
                {"sheet": "06_거래서류_청구", "settle_id": row["정산ID"],
                 "col_name": "입금일", "value": str(d["일자"])[:10],
                 "only_if_empty": True,
                 "why": "보험사 입금 자동확인(%s %s원)" % (d["거래처"], format(int(d["금액"]), ","))},
                {"sheet": "06_거래서류_청구", "settle_id": row["정산ID"],
                 "col_name": "입금액", "value": int(d["금액"]),
                 "only_if_empty": True,
                 "why": "보험사 입금 자동확인(%s)" % d["거래처"]},
            ], who="insurance_watch")
            n += 2
        except Exception as e:
            print("  ! 큐 적재 실패:", row["정산ID"], e)
    return n


def write_report(res, ins_rows, unpaid, ins_dep, paired, held):
    L = []
    L.append("# 보험 건·보험사 입금 확인")
    L.append("")
    L.append("- 만든 때: %s · 원장 %s" % (res["시각"], res["원장"]))
    L.append("- 원장에 `보험`으로 적힌 행: **%d건** (그중 입금 안 된 것 %d건)"
             % (res["보험행"], res["입금안됨"]))
    L.append("- 살펴본 입금 %d건 중 보험사로 보이는 것: **%d건**"
             % (res["입금건수"], res["보험사입금"]))
    L.append("- 짝지어짐 %d · 보류 %d · 큐 적재 %d칸"
             % (res["짝지어짐"], res["보류"], res["큐적재"]))
    L.append("")

    # ★ 0건일 때가 제일 위험하다 — '없는 것'인지 '안 본 것'인지를 반드시 적는다([169]).
    if res["보험행"] == 0 or res["보험사입금"] == 0:
        L.append("## 0건이 나온 이유 — 없는 것인가, 안 본 것인가")
        if res["보험행"] == 0:
            L.append("- 원장 06시트에 `비용구분 = 보험` 인 행이 **한 건도 없습니다.**")
            L.append("  (읽기는 했습니다 — 같은 열에서 유상·무상·미확정은 정상적으로 읽혔습니다.)")
            L.append("  보험 건인데 칸이 비어 있으면 여기서 안 잡힙니다. 그런 행은 정산 화면에")
            L.append("  **`비용구분 미입력`** 으로 뜹니다 — 예전엔 `무상/보험` 회색 딱지에 묻혀 있었습니다.")
        if res["보험사입금"] == 0:
            L.append("- 살펴본 입금 %d건 중 거래처 이름에 보험 업종 낱말"
                     "(보험·해상·화재·손해·공제)이 들어간 것이 **없습니다.**" % res["입금건수"])
            L.append("- 본 자료: %s" % (", ".join(res["본자료"]) or "(없음)"))
            L.append("  입금 원천이 '쿠팡 입금내역'뿐이면 보험금은 **원래 여기 안 옵니다** —")
            L.append("  보험사 입금이 찍히는 은행·ERP 자료를 `0. 원본 자료/7. 입금내역` 에 넣어 주시면")
            L.append("  다음 회차부터 자동으로 봅니다.")
        if res["못본것"]:
            L.append("- ⚠ 못 읽은 원천: %s" % "; ".join(res["못본것"]))
        L.append("")

    if unpaid:
        L.append("## 보험 건인데 입금이 아직 안 잡힌 것 (%d건)" % len(unpaid))
        L.append("")
        L.append("| 정산ID | 프로젝트NO | 캠프 | 청구액 |")
        L.append("|---|---|---|---:|")
        for x in unpaid[:80]:
            L.append("| %s | %s | %s | %s |" % (x["정산ID"], x["프로젝트NO"],
                                                x["캠프명"], format(int(x["청구액"]), ",")))
        L.append("")

    if ins_dep:
        L.append("## 보험사로 보이는 입금 (%d건)" % len(ins_dep))
        L.append("")
        L.append("| 일자 | 거래처 | 금액 | 출처 | 판정 |")
        L.append("|---|---|---:|---|---|")
        why = {id(d): w for d, w in held}
        ok = {id(d) for d, _ in paired}
        for d in ins_dep[:80]:
            v = "자동 반영" if id(d) in ok else why.get(id(d), "사람 확인")
            L.append("| %s | %s | %s | %s | %s |"
                     % (str(d["일자"])[:10], d["거래처"],
                        format(int(d["금액"]), ","), d["출처"], v))
        L.append("")

    L.append("---")
    L.append("이 리포트는 **아무것도 고치지 않습니다.** 확실한 짝만 대기열에 넣고,")
    L.append("엑셀 반영은 11:00·15:00 회차가 합니다.")

    os.makedirs(REPORT_DIR, exist_ok=True)
    with io.open(MD, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(L) + "\n")
    with io.open(JS, "w", encoding="utf-8", newline="") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)


def main(argv):
    res = build(queue="--queue" in argv)
    print("보험 건 %d(미입금 %d) · 입금 %d 중 보험사 %d · 짝 %d · 보류 %d · 큐 %d칸"
          % (res["보험행"], res["입금안됨"], res["입금건수"], res["보험사입금"],
             res["짝지어짐"], res["보류"], res["큐적재"]))
    print("  → reports/보험입금_확인.md")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
