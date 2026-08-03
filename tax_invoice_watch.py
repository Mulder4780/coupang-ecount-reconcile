# -*- coding: utf-8 -*-
"""오래된 세금계산서 미발행 건을 경과일 순으로 잡아낸다 (2026-08-03 지시, 상시).

사용자 지시: "오래된 건 중 세금계산서 미발행 건이 너무 많은데 잡아내."

* 관리대장 06 정산 중 ``settle_status``가 ``세금계산서 미발행``이고 DB 완료
  정본(``resolution``)에도 완료가 없는 건을, 작업완료일 기준 경과일로 정렬해
  ``reports/세금계산서_미발행_경과.md``(+ ``.json``)로 쓴다.
* 발행일을 지어내거나 엑셀을 고치지 않는다 — 이 모듈은 읽기 전용 감시다.
  발행 조치는 사람이 하고, 발행 사실은 ERP 원본이 들어오면 객관완료가 잡는다.
* daily_run 이 매일 실행한다. 앱 '보고' 탭이 최신 리포트를 그대로 보여 준다.
"""
import json
import os
import sys
from datetime import date, datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import ledger_db
from ecount_reconcile import (erp_progress, load_config, read_ledger,
                              resolve_master, settle_status)

REPORT_MD = os.path.join(ROOT, "reports", "세금계산서_미발행_경과.md")
REPORT_JSON = os.path.join(ROOT, "reports", "세금계산서_미발행_경과.json")
BUCKETS = ((90, "90일 초과"), (60, "61~90일"), (30, "31~60일"), (0, "30일 이하"))


def _day(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "")[:10]
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def overdue_rows(records, resolutions, progress, today=None):
    today = today or date.today()
    rows = []
    for settle_id, record in sorted((records or {}).items()):
        if settle_status(record) != "세금계산서 미발행":
            continue
        resolved = str((resolutions.get(settle_id) or {}).get("status") or "")
        if resolved.startswith("완료("):
            continue
        done = _day(record.get("작업완료일")) or _day(record.get("원장_거래명세서발행일"))
        age = (today - done).days if done else None
        project = str(record.get("프로젝트NO") or "").strip()
        rows.append({
            "settle_id": settle_id,
            "project": project,
            "camp": str(record.get("캠프명") or "").strip(),
            "kind": str(record.get("업무구분") or "").strip(),
            "done": done.isoformat() if done else "",
            "age": age,
            "statement_total": record.get("원장_거래명세서합계") or 0,
            "erp_status": progress.get(project, ""),
        })
    rows.sort(key=lambda row: -(row["age"] if row["age"] is not None else -1))
    return rows


def bucket_counts(rows):
    counts = {label: 0 for _, label in BUCKETS}
    counts["완료일 불명"] = 0
    for row in rows:
        age = row["age"]
        if age is None:
            counts["완료일 불명"] += 1
            continue
        for limit, label in BUCKETS:
            if age > limit or limit == 0:
                counts[label] += 1
                break
    return counts


def write_reports(rows, today=None):
    today = today or date.today()
    counts = bucket_counts(rows)
    os.makedirs(os.path.dirname(REPORT_MD), exist_ok=True)
    with open(REPORT_MD, "w", encoding="utf-8") as handle:
        handle.write(f"# 세금계산서 미발행 경과 — {today.isoformat()}\n\n")
        handle.write(f"총 **{len(rows)}건**. 발행일을 지어내지 않으며, ERP 계산서 원본이 "
                     f"들어오면 객관완료가 자동으로 잡는다.\n\n")
        handle.write("| 경과 | 건수 |\n|---|---:|\n")
        for label in list(counts):
            handle.write(f"| {label} | {counts[label]} |\n")
        handle.write("\n## 오래된 순 (발행 조치 우선순위)\n\n")
        handle.write("| 정산ID | 프로젝트 | 캠프 | 구분 | 완료일 | 경과일 | 명세서합계 | ERP진행 |\n")
        handle.write("|---|---|---|---|---|---:|---:|---|\n")
        for row in rows:
            handle.write(
                f"| {row['settle_id']} | {row['project']} | {row['camp']} | {row['kind']} "
                f"| {row['done']} | {row['age'] if row['age'] is not None else '-'} "
                f"| {int(row['statement_total'] or 0):,} | {row['erp_status']} |\n"
            )
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "total": len(rows),
        "buckets": counts,
        "rows": rows,
    }
    tmp = REPORT_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=1)
    os.replace(tmp, REPORT_JSON)
    return payload


def main():
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    records = read_ledger(master)
    payload = write_reports(overdue_rows(records, ledger_db.resolutions(), erp_progress()))
    counts = payload["buckets"]
    print(f"세금계산서 미발행 {payload['total']}건 — "
          + " · ".join(f"{label} {counts[label]}" for label in counts if counts[label]))
    print("리포트:", REPORT_MD)


if __name__ == "__main__":
    main()
