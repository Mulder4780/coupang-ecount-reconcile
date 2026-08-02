# -*- coding: utf-8 -*-
"""정산 대기 중 독립 원자료로 입증된 건을 DB 완료 정본에 기록한다.

완료 기준은 서로 다른 원천이 정확히 맞는 경우로 제한한다.

* 금액 재계산 대기: 관리대장 06의 프로젝트·PO·거래명세서 합계와 PO 원본
  견적서의 프로젝트·PO·부가세 포함 총액이 모두 같고, 일치 견적서가 한 장뿐인 건.
* 세금계산서 미발행: 관리대장 26_계산서구성에서 프로젝트가 연결되고 판정이
  ``확정``인 ERP 계산서 건.

완료 상태는 Excel 수식이나 발행일을 덮어쓰지 않고 ``ledger_db.resolution``에만
저장한다. 이 모듈은 daily_run에서 매일 재실행되며 upsert라 중복되지 않는다.
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import ledger_db
from ecount_reconcile import has_statement, load_config, read_ledger, resolve_master
from po_pdf import CACHE_FILE, PO_RE

REPORT = os.path.join(ROOT, "reports", "정산_객관완료.json")
AMOUNT_STATUS = "완료(견적·명세서 금액확인)"
INVOICE_STATUS = "완료(ERP 계산서 원본확인)"


def _money(value):
    try:
        return round(float(value or 0))
    except (TypeError, ValueError):
        return 0


def _po(value):
    match = re.search(r"PO\d+", str(value or ""), re.I)
    return match.group().upper() if match else ""


def _raw_amount_wait(record):
    return (
        record.get("비용구분") == "유상"
        and str(record.get("원천업무ID") or "").startswith("AS-")
        and not _money(record.get("원장_공급가액"))
        and bool(record.get("원장_거래명세서발행일"))
        and _money(record.get("원장_거래명세서합계")) > 0
    )


def dedup_quote_rows(cache_path=CACHE_FILE):
    """PO 원본 파서 캐시를 ``po_pdf.scan``과 같은 키로 중복 제거해 반환한다.

    원본이 정본 폴더와 과거 공유 폴더에 함께 있어 경로만 다르면 같은 견적서가 두
    장처럼 보인다. 네트워크 PDF 전체를 다시 열지 않고 이미 검증된 파서 캐시를 쓴다.
    """
    try:
        cache = json.load(open(cache_path, encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return []
    rows, seen = [], set()
    for cache_key, row in cache.items():
        if row.get("종류") != "견적서":
            continue
        try:
            path, size, _mtime = cache_key.rsplit("|", 2)
        except ValueError:
            continue
        joined = os.path.normpath(path).replace(" ", "")
        match = PO_RE.search(joined)
        parent = os.path.basename(os.path.dirname(path)).lower()
        key = (
            match.group(1) if match else parent,
            os.path.basename(path).lower(),
            size,
        )
        if key in seen:
            continue
        seen.add(key)
        rows.append(dict(row))
    return rows


def quote_index(rows):
    """(프로젝트, PO, 부가세포함 총액) → 중복 제거된 견적서 목록."""
    out = {}
    for row in rows or []:
        project = str(row.get("프로젝트NO") or "").strip().upper()
        po_no = _po(row.get("PO번호"))
        total = _money(row.get("금액"))
        if project and po_no and total > 0:
            out.setdefault((project, po_no, total), []).append(row)
    return out


def confirmed_invoice_index(master):
    """26_계산서구성의 확정 행을 프로젝트별로 읽는다."""
    import openpyxl

    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    try:
        ws = next((sheet for sheet in wb.worksheets if sheet.title.startswith("26_")), None)
        if ws is None:
            return {}
        header = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        index = {str(value or "").strip(): i for i, value in enumerate(header) if value}
        required = {"일자-No.", "계산서공급가액", "포함프로젝트NO", "판정"}
        if not required.issubset(index):
            return {}
        out = {}
        for row in ws.iter_rows(min_row=5, values_only=True):
            get = lambda col: row[index[col]] if index[col] < len(row) else None
            verdict = str(get("판정") or "").strip()
            amount = _money(get("계산서공급가액"))
            slip = str(get("일자-No.") or "").strip()
            if not verdict.startswith("확정") or not slip or amount <= 0:
                continue
            evidence = {"slip": slip, "amount": amount, "verdict": verdict}
            for project in str(get("포함프로젝트NO") or "").split(","):
                project = project.strip().upper()
                if project:
                    out.setdefault(project, []).append(evidence)
        return out
    finally:
        wb.close()


def objective_entries(records, quotes, invoices, existing=None):
    """현재 원장 레코드에서 객관완료 DB 항목을 만든다.

    이미 ERP 수금완료처럼 더 강한 완료 근거가 있으면 상태를 낮춰 쓰지 않는다. 이
    모듈이 전에 쓴 두 상태는 다시 반환해 ``last_seen``만 갱신한다.
    """
    existing = existing or {}
    qindex = quote_index(quotes)
    entries = []
    for settle_id, record in sorted((records or {}).items()):
        old = existing.get(settle_id) or {}
        old_status = str(old.get("status") or "")
        if old_status.startswith("완료(") and old_status not in (AMOUNT_STATUS, INVOICE_STATUS):
            continue
        project = str(record.get("프로젝트NO") or "").strip().upper()

        if _raw_amount_wait(record):
            po_no = _po(record.get("원장_PO번호"))
            total = _money(record.get("원장_거래명세서합계"))
            hits = qindex.get((project, po_no, total), [])
            if len(hits) == 1:
                quote = hits[0]
                issue_day = str(record.get("원장_거래명세서발행일") or "")[:10]
                entries.append({
                    "settle_id": settle_id,
                    "project": project,
                    "status": AMOUNT_STATUS,
                    "basis": (
                        f"PO 원본 견적서 {quote.get('파일') or '-'} · 프로젝트 {project} · "
                        f"{po_no} · 부가세포함 {total:,}원 = 관리대장 06 거래명세서 "
                        f"{issue_day} 합계 {total:,}원 (유일 일치)"
                    ),
                    "evidence_kind": "amount",
                })
                continue

        issued = record.get("원장_세금계산서실제발행일") or record.get("원장_세금계산서발행일")
        invoice_hits = invoices.get(project, [])
        if (
            record.get("비용구분") == "유상"
            and not _raw_amount_wait(record)
            and has_statement(record)
            and not issued
            and invoice_hits
        ):
            slips = ", ".join(str(item["slip"]) for item in invoice_hits)
            amounts = sum(_money(item["amount"]) for item in invoice_hits)
            verdicts = ", ".join(sorted({str(item["verdict"]) for item in invoice_hits}))
            entries.append({
                "settle_id": settle_id,
                "project": project,
                "status": INVOICE_STATUS,
                "basis": (
                    f"관리대장 25 ERP 매출(세금)계산서 원본 · 26 계산서구성 "
                    f"{verdicts} · 프로젝트 {project} · 전표 {slips} · 공급가액 {amounts:,}원"
                ),
                "evidence_kind": "invoice",
            })
    return entries


def write_report(master, entries, synced=0):
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "master": os.path.basename(master),
        "eligible": len(entries),
        "amount": sum(row.get("evidence_kind") == "amount" for row in entries),
        "invoice": sum(row.get("evidence_kind") == "invoice" for row in entries),
        "synced": synced,
        "entries": entries,
    }
    tmp = REPORT + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    os.replace(tmp, REPORT)
    return payload


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--sync", action="store_true", help="객관근거 완료를 resolution DB에 upsert")
    args = parser.parse_args(argv)
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    records = read_ledger(master)
    existing = ledger_db.resolutions()
    quotes = dedup_quote_rows()
    invoices = confirmed_invoice_index(master)
    entries = objective_entries(records, quotes, invoices, existing)
    synced = ledger_db.resolution_sync(entries) if args.sync else 0
    payload = write_report(master, entries, synced)
    print(
        f"정산 객관완료 후보 {payload['eligible']}건 "
        f"(금액 {payload['amount']} · 계산서 {payload['invoice']})"
        + (f" · DB 동기화 {synced}건" if args.sync else " · dry-run")
    )
    print("리포트:", REPORT)


if __name__ == "__main__":
    main()
