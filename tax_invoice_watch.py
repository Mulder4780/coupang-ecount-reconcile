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
import re
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


def verified_erp_unissued(today=None):
    """ERP 계산서 진행단계의 **실제 미발행**을 회사/쿠팡/관리범위로 가른다.

    관리대장 발행일 공란은 '원장이 발행일을 모른다'는 뜻이지 미발행 증거가 아니다.
    실제 발행 조치 목록은 ERP ``taxstep``의 미발행 상태를 정본으로 삼고, 이 프로젝트의
    관리범위는 돌발AS·정기점검으로 제한한다. ERP가 없으면 0건이 아니라 ``available=False``로
    돌려 자료 부재와 실제 0건을 구분한다.
    """
    today = today or date.today()
    try:
        import datalake as D
        from sales_slip_gap import collect, diagnose, issued_index
        con = D.connect()
        try:
            since, until = f"{today.year}-01-01", today.isoformat()
            invoices, slips = collect(con, since, until)
            groups = diagnose(invoices, slips, issued_index(con, since, until))
        finally:
            con.close()
    except Exception as exc:
        return {"available": False, "error": f"{type(exc).__name__}: {str(exc)[:180]}",
                "all_company": 0, "coupang": 0, "rows": []}
    raw = groups.get("미발행") or []
    coupang = [row for row in raw if "쿠팡" in str(row.get("거래처") or "")]
    scope = [row for row in coupang
             if re.search(r"돌발\s*AS|정기\s*점검",
                          str(row.get("프로젝트") or "") + " " + str(row.get("적요") or ""), re.I)]
    return {"available": True, "all_company": len(raw), "coupang": len(coupang),
            "rows": scope, "since": f"{today.year}-01-01", "until": today.isoformat()}


def write_verified_reports(verified, ledger_blanks, today=None):
    """실제 미발행과 원장 빈칸 후보를 섞지 않은 보고서/앱 JSON을 쓴다."""
    today = today or date.today()
    rows = list((verified or {}).get("rows") or [])
    available = bool((verified or {}).get("available"))
    os.makedirs(os.path.dirname(REPORT_MD), exist_ok=True)
    with open(REPORT_MD, "w", encoding="utf-8") as handle:
        handle.write(f"# 세금계산서 미발행 교차확인 — {today.isoformat()}\n\n")
        if available:
            handle.write(
                f"ERP 진행단계에서 실제 `미발행`은 회사 전체 **{verified.get('all_company', 0)}건**, "
                f"쿠팡 **{verified.get('coupang', 0)}건**, 돌발AS·정기점검 관리범위 "
                f"**{len(rows)}건**입니다.\n\n")
        else:
            handle.write("ERP 진행단계 자료를 읽지 못했습니다. **0건으로 단정하지 않습니다.**\n\n")
        handle.write(
            f"> 관리대장 발행일 공란 **{len(ledger_blanks)}건**은 `미발행`이 아니라 "
            "**ERP 근거 연결 대기**입니다. 이 목록으로 재발행하면 이중발행 위험이 있으므로 "
            "실제 발행 조치 건수에 포함하지 않습니다.\n\n")
        handle.write("## 사람이 ERP에서 확인·발행할 관리범위\n\n")
        if rows:
            handle.write("| 일자 | 전표번호 | 거래처 | 금액 | 프로젝트 |\n|---|---|---|---:|---|\n")
            for row in rows:
                handle.write(f"| {row.get('일자','')} | {row.get('전표번호','')} | "
                             f"{row.get('거래처','')} | {int(row.get('금액') or 0):,} | "
                             f"{str(row.get('프로젝트') or '')[:80]} |\n")
        else:
            handle.write("현재 읽힌 ERP 근거에서 관리범위 미발행 확정 건이 없습니다.\n")
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "basis": "ERP 계산서 진행단계 미발행 상태 · 관리범위=돌발AS/정기점검",
        "available": available,
        "total": len(rows),
        "company_total": int((verified or {}).get("all_company") or 0),
        "coupang_total": int((verified or {}).get("coupang") or 0),
        "ledger_blank_total": len(ledger_blanks),
        "buckets": {"발행 승인 필요": len(rows), "원장 빈칸(발행대상 아님)": len(ledger_blanks)},
        "rows": rows,
        "error": (verified or {}).get("error", ""),
    }
    tmp = REPORT_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=1)
    os.replace(tmp, REPORT_JSON)
    return payload


def main():
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    records = read_ledger(master)
    ledger_blanks = overdue_rows(records, ledger_db.resolutions(), erp_progress())
    payload = write_verified_reports(verified_erp_unissued(), ledger_blanks)
    if payload["available"]:
        print(f"ERP 실제 미발행: 회사 전체 {payload['company_total']}건 · 쿠팡 "
              f"{payload['coupang_total']}건 · 돌발AS/정기점검 {payload['total']}건")
    else:
        print("ERP 미발행 원본을 읽지 못함 — 0건으로 처리하지 않음")
    print(f"관리대장 발행일 공란 {payload['ledger_blank_total']}건은 근거 연결 대기로 분리")
    print("리포트:", REPORT_MD)


if __name__ == "__main__":
    main()
