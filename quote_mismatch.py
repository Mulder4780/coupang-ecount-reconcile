# -*- coding: utf-8 -*-
"""금액 재계산 대기 건의 견적↔명세 교차·불일치 진단 (2026-08-03 지시, 상시).

실사례: JS-2604-494(UJ2600777)의 명세서합계 722,480원이 같은 PO의 **다른 프로젝트**
UJ2600783 견적과 정확히 일치했다 — 06시트 입력 때 행이 밀린 교차 오류다. 이런 건은
자동 완료하면 틀린 금액을 확정하므로, 대신 **어느 견적과 바뀌었는지 짝**을 찾아
사람이 바로잡을 목록을 만든다. 원장 값은 고치지 않는다(사람 판단 몫).

분류
  교차 의심  : 본인 견적과 다르고, 같은 PO 안 다른 프로젝트 견적과 유일 일치
  금액 불일치: 본인 견적은 있는데 금액이 다르고 교차 상대도 없음
  견적 없음  : 이 프로젝트 견적서를 아직 못 찾음(PO 원본 수집 필요)

daily_run 이 매일 실행하고 앱 '보고' 탭이 최신본을 보여 준다.
"""
import json
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import ledger_db
from ecount_reconcile import load_config, read_ledger, resolve_master, settle_status
from settlement_completion import _money, _po, dedup_quote_rows

REPORT_MD = os.path.join(ROOT, "reports", "견적명세_불일치.md")
REPORT_JSON = os.path.join(ROOT, "reports", "견적명세_불일치.json")


def diagnose(records, quotes, resolutions=None):
    resolutions = resolutions or {}
    by_po = {}
    by_project = {}
    for q in quotes or []:
        po = _po(q.get("PO번호"))
        prj = str(q.get("프로젝트NO") or "").strip().upper()
        total = _money(q.get("금액"))
        if po:
            by_po.setdefault(po, []).append((prj, total, q.get("파일") or ""))
        if prj and total > 0:
            by_project.setdefault(prj, []).append(total)

    rows = []
    for sid, r in sorted((records or {}).items()):
        if settle_status(r) != "금액 재계산 대기":
            continue
        if str((resolutions.get(sid) or {}).get("status") or "").startswith("완료("):
            continue
        prj = str(r.get("프로젝트NO") or "").strip().upper()
        po = _po(r.get("원장_PO번호"))
        stmt = _money(r.get("원장_거래명세서합계"))
        own = sorted(set(by_project.get(prj, [])))
        row = {"정산ID": sid, "프로젝트NO": prj, "PO": po, "명세합계": stmt,
               "본인견적": own}
        if own and stmt in own:
            continue                                   # 정상 일치 — 완료 알고리즘 몫
        cross = [(p, t, f) for p, t, f in by_po.get(po, []) if t == stmt and p != prj]
        if len(cross) == 1:
            row["유형"] = "교차 의심"
            row["교차상대"] = cross[0][0]
            row["근거"] = f"{cross[0][0]} 견적 {stmt:,}원과 정확 일치 ({os.path.basename(cross[0][2])[:40]})"
        elif own:
            row["유형"] = "금액 불일치"
            row["근거"] = f"본인 견적 {', '.join(f'{t:,}' for t in own)}원 ≠ 명세 {stmt:,}원"
        else:
            row["유형"] = "견적 없음"
            row["근거"] = "이 프로젝트 견적서를 캐시에서 못 찾음"
        rows.append(row)
    order = {"교차 의심": 0, "금액 불일치": 1, "견적 없음": 2}
    rows.sort(key=lambda x: (order.get(x["유형"], 9), x["정산ID"]))
    return rows


def write_report(rows):
    os.makedirs(os.path.dirname(REPORT_MD), exist_ok=True)
    counts = {}
    for r in rows:
        counts[r["유형"]] = counts.get(r["유형"], 0) + 1
    with open(REPORT_MD, "w", encoding="utf-8") as f:
        f.write(f"# 견적↔명세 불일치 진단 — {datetime.now():%Y-%m-%d %H:%M}\n\n")
        f.write("금액 재계산 대기 중 **명세서합계가 견적과 안 맞는** 건. '교차 의심'은 같은 PO의 "
                "다른 프로젝트 견적과 정확히 일치 — 06시트 입력 밀림일 가능성이 높다. "
                "명세서합계를 바로잡으면(사람) 다음 회차에 자동 완료된다.\n\n")
        f.write("| 유형 | 건수 |\n|---|---:|\n")
        for k, n in counts.items():
            f.write(f"| {k} | {n} |\n")
        f.write("\n| 정산ID | 프로젝트NO | PO | 명세합계 | 유형 | 근거 |\n|---|---|---|---:|---|---|\n")
        for r in rows:
            f.write(f"| {r['정산ID']} | {r['프로젝트NO']} | {r['PO']} | {r['명세합계']:,} "
                    f"| {r['유형']} | {r['근거']} |\n")
    payload = {"generated_at": datetime.now().isoformat(timespec="seconds"),
               "total": len(rows), "counts": counts, "rows": rows}
    tmp = REPORT_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    os.replace(tmp, REPORT_JSON)
    return payload


def main():
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    records = read_ledger(master)
    rows = diagnose(records, dedup_quote_rows(), ledger_db.resolutions())
    payload = write_report(rows)
    print("견적↔명세 진단:", " · ".join(f"{k} {n}" for k, n in payload["counts"].items()) or "문제 없음")
    print("리포트:", REPORT_MD)


if __name__ == "__main__":
    main()
