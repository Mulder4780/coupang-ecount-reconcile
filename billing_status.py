# -*- coding: utf-8 -*-
"""06시트 `청구상태`(AH)를 ERP 가 입증한 단계로 맞춘다.

사용자 지시(2026-08-08): **"ERP 기준으로 확정하고 객관적으로 입증되면 엑셀에 완료처리해"**

★ 왜 이 열은 엑셀에 쓰는가 — 프로젝트의 방향은 "엑셀에 저장하지 않고 DB로만"이다.
  그 예외는 둘뿐이고(① 사람이 엑셀에서 직접 읽는 값 ② 기존 수식이 참조하는 값),
  `청구상태` 는 ①이다. 데이터유효성 안내문이 **"[필수] 작업완료부터 입금완료까지 현재
  단계를 고릅니다"** 라고 적혀 있는, 사람이 보고 고르는 칸이다. 그래서 여기 있는 값이
  실제와 다르면 **사람이 틀린 단계를 보고 일한다** — DB 에만 적어 두면 뜻이 없다.

★ 무엇이 '객관적으로 입증'인가 — `settle_status(r) == '완료(ERP 수금확인)'` 하나다.
  그 프로젝트의 ERP 전표가 **전부 `7.수금완료`** 라는 뜻이다(섞여 있으면 `erp_progress`
  가 '혼재'로 돌려주므로 여기 안 걸린다). 원장 칸이 비었는지 여부와 무관하게, 돈이
  들어왔다고 ERP 가 말하는 것이 근거다 — **원장 빈 칸은 근거가 아니다**(그 착각이
  '세금계산서 미발행 190건'의 정체였다).

★ 낱말은 지어내지 않는다. 이 열에 실제로 쓰여 있는 값과 안내문이 사다리를 정한다:
      작업완료 → 거래명세서발행 → 세금계산서발행 → 청구 → 입금완료
  마지막 칸이 '완료'가 아니라 **'입금완료'** 인 이유가 이것이다. 새 낱말을 넣으면
  사람이 쓰던 사다리에 없는 값이 섞여, 정렬·필터가 조용히 어긋난다.

★ 사다리는 **한 방향으로만** 움직인다. 아래 단계('작업완료')가 적힌 칸은 올리고,
  모르는 낱말이 적힌 칸은 **건드리지 않고 사람에게 넘긴다.** 사람이 적은 값을 덮는
  것이 위험한 게 아니라, **무슨 뜻인지 모르는 값을 덮는 것**이 위험하다.

엑셀은 열지 않는다. `--queue` 로 넣으면 11:00·15:00 회차가 반영한다.
"""
from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SHEET = "06_거래서류청구수금"
COL = "청구상태"
DONE = "입금완료"

# 사다리 — 이 열에 실제로 쓰여 있는 낱말 + 안내문이 정한 순서다.
LADDER = ["작업완료", "거래명세서발행", "세금계산서발행", "청구", DONE]

REPORT = os.path.join(ROOT, "reports", "청구상태_반영.md")


def plan():
    """무엇을 어떻게 바꿀지 정한다(엑셀·DB 를 건드리지 않는다)."""
    import ecount_reconcile as R

    master = R.resolve_master(R.load_config()["reconcile"]["master_xlsx"])
    recs = R.read_ledger(master)
    fill, raise_, already, unknown, not_yet = [], [], [], [], {}

    for sid, r in sorted(recs.items()):
        state = R.settle_status(r)
        if state != "완료(ERP 수금확인)":
            not_yet[state] = not_yet.get(state, 0) + 1
            continue
        now = str(r.get("원장_청구상태") or "").strip()
        item = {
            "정산ID": sid,
            "프로젝트NO": r.get("프로젝트NO") or "",
            "캠프명": r.get("캠프명") or "",
            "업무구분": r.get("업무구분") or "",
            "지금": now or "(빈칸)",
        }
        if now == DONE:
            already.append(item)
        elif not now:
            fill.append(item)
        elif now in LADDER:
            # 사다리 위에 있는 낱말이면 위치를 비교해 **올릴 때만** 바꾼다.
            if LADDER.index(now) < LADDER.index(DONE):
                raise_.append(item)
            else:
                already.append(item)
        else:
            unknown.append(item)

    return {"master": master, "채움": fill, "올림": raise_,
            "이미맞음": already, "모르는값": unknown, "대상아님": not_yet}


def items_for_queue(p):
    """대기열에 넣을 셀 목록. 빈칸과 '올림'은 `only_if_empty` 가 다르다."""
    out = []
    for it in p["채움"]:
        out.append({
            "sheet": SHEET, "key_col": "정산ID", "key": it["정산ID"],
            "col": COL, "value": DONE, "vtype": "text",
            # 빈칸만 채운다 — 그 사이 사람이 무언가 적었으면 그 사람이 옳다.
            "only_if_empty": True,
            "evidence": "ERP 진행상태 7.수금완료 (프로젝트 %s 전표 전부) — 빈칸 채움"
                        % (it["프로젝트NO"] or "?"),
        })
    for it in p["올림"]:
        out.append({
            "sheet": SHEET, "key_col": "정산ID", "key": it["정산ID"],
            "col": COL, "value": DONE, "vtype": "text",
            # ★ 여기만 덮어쓴다. 근거는 '사다리를 거꾸로 가지 않는다'는 것뿐이다 —
            #   `작업완료` 라 적힌 칸에 ERP 가 수금완료라고 말하면 그 딱지는 **낡은
            #   것**이고, 낡은 단계를 그냥 두면 사람이 다 받은 돈을 다시 받으러 간다.
            "only_if_empty": False,
            "evidence": "ERP 진행상태 7.수금완료 (프로젝트 %s) — '%s'에서 올림"
                        % (it["프로젝트NO"] or "?", it["지금"]),
        })
    return out


def write_report(p):
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write("# 06시트 청구상태 — ERP 수금확인 반영\n\n")
        f.write("근거: ERP 판매조회 진행상태가 프로젝트의 모든 전표에서 `7.수금완료`.\n")
        f.write("원장의 빈 칸은 근거가 아닙니다(비었다는 것과 안 나갔다는 것은 다릅니다).\n\n")
        f.write("| 구분 | 건수 | 무엇을 하나 |\n|---|---|---|\n")
        f.write("| 빈칸 채움 | %d | `%s` 로 채웁니다 |\n" % (len(p["채움"]), DONE))
        f.write("| 낡은 단계 올림 | %d | 사다리 위로만 올립니다(되돌리지 않습니다) |\n"
                % len(p["올림"]))
        f.write("| 이미 맞음 | %d | 손대지 않습니다 |\n" % len(p["이미맞음"]))
        f.write("| 모르는 값 | %d | **사람 확인** — 뜻을 모르는 낱말은 덮지 않습니다 |\n"
                % len(p["모르는값"]))
        if p["올림"]:
            f.write("\n## 낡은 단계를 올리는 행 (사람이 적어 둔 칸입니다)\n\n")
            f.write("| 정산ID | 프로젝트NO | 캠프 | 업무 | 지금 |\n|---|---|---|---|---|\n")
            for it in p["올림"]:
                f.write("| %s | %s | %s | %s | %s |\n"
                        % (it["정산ID"], it["프로젝트NO"], it["캠프명"],
                           it["업무구분"], it["지금"]))
        if p["모르는값"]:
            f.write("\n## 뜻을 모르는 낱말 — 손대지 않았습니다\n\n")
            f.write("| 정산ID | 지금 적힌 값 |\n|---|---|\n")
            for it in p["모르는값"]:
                f.write("| %s | %s |\n" % (it["정산ID"], it["지금"]))
        f.write("\n## 아직 대상이 아닌 행\n\n")
        f.write("| 정산 상태 | 건수 |\n|---|---|\n")
        for k, n in sorted(p["대상아님"].items(), key=lambda x: -x[1]):
            f.write("| %s | %d |\n" % (k, n))
        f.write("\n※ 엑셀은 열지 않습니다. `--queue` 로 넣으면 11:00·15:00 회차가 반영합니다.\n")
    return REPORT


def main(argv=None):
    ap = argparse.ArgumentParser(description="06시트 청구상태를 ERP 수금확인으로 맞춘다")
    ap.add_argument("--queue", action="store_true", help="대기열에 넣는다(엑셀은 회차가 연다)")
    a = ap.parse_args(argv)

    p = plan()
    print("원장 기준 %s" % os.path.basename(p["master"]))
    print("  빈칸 채움 %d · 낡은 단계 올림 %d · 이미맞음 %d · 모르는값 %d"
          % (len(p["채움"]), len(p["올림"]), len(p["이미맞음"]), len(p["모르는값"])))
    print("  리포트: %s" % write_report(p))

    if not a.queue:
        print("  (넣지 않았습니다 — 넣으려면 --queue)")
        return 0

    items = items_for_queue(p)
    if not items:
        print("  대기열에 넣을 것이 없습니다.")
        return 0
    import ledger_db
    added = ledger_db.enqueue(items, source="billing_status", ingest_prefix="billstat")
    print("  대기열 %d건 (11:00·15:00 회차가 반영)" % added)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
