# -*- coding: utf-8 -*-
"""객관 자료가 확인된 업무를 담당자별 완료 정본에 계속 반영한다.

상태 문구는 사용자가 지정한 ``류지영 완료``·``오종현 완료``·``유현민 완료``로
고정한다. 사람의 체크나 추정 상태는 읽지 않고, 이미 별도 검증기가 확정한 DB 완료
근거와 PO 대조 근거만 받아 쓴다. Excel 상태 셀은 수정하지 않는다.
"""
from __future__ import annotations

import json
import os
import sys
from collections import Counter
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(ROOT, "reports")
PO_EVIDENCE = os.path.join(REPORT_DIR, "po_objective_evidence.json")
REPORT_PATH = os.path.join(REPORT_DIR, "담당자_객관완료.md")
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def objective_entries(po_evidence_path=PO_EVIDENCE):
    """검증된 정본만 담당자 완료 항목으로 변환한다."""
    import ledger_db

    entries = []
    with ledger_db.conn() as c:
        for kind, record_id, project, completed_on, basis in c.execute(
            "SELECT kind,record_id,project,completed_on,basis FROM work_resolution"
        ):
            entries.append({
                "owner": "류지영", "task_kind": f"field_{kind}",
                "record_id": record_id, "project": project,
                "completed_on": completed_on,
                "basis": "현장업무 객관완료 · " + basis,
            })
        for settle_id, project, status, basis, first_seen in c.execute(
            "SELECT settle_id,project,status,basis,first_seen FROM resolution"
        ):
            entries.append({
                "owner": "류지영", "task_kind": "settlement",
                "record_id": settle_id, "project": project,
                "completed_on": str(first_seen)[:10],
                "basis": f"정산 객관완료 · {status} · {basis}",
            })

    try:
        data = json.load(open(po_evidence_path, encoding="utf-8"))
        entries.extend(data.get("entries") or [])
    except FileNotFoundError:
        pass
    except (OSError, ValueError, TypeError) as exc:
        print(f"PO 객관근거 읽기 생략: {exc}")

    # 같은 담당자·업무종류·원천ID는 한 번만. 최신 대조 근거가 이전 근거를 갱신한다.
    deduped = {}
    for entry in entries:
        key = (str(entry.get("owner") or ""), str(entry.get("task_kind") or ""),
               str(entry.get("record_id") or ""))
        if all(key):
            deduped[key] = entry
    return list(deduped.values())


def objective_retractions(po_evidence_path=PO_EVIDENCE):
    """대조기가 명시적으로 비유일·충돌로 되돌린 자동 완료 키만 반환한다."""
    try:
        data = json.load(open(po_evidence_path, encoding="utf-8"))
        return list(data.get("retractions") or [])
    except FileNotFoundError:
        return []
    except (OSError, ValueError, TypeError) as exc:
        print(f"PO 완료회수 근거 읽기 생략: {exc}")
        return []


def write_report(rows, path=REPORT_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    counts = Counter(row["owner"] for row in rows)
    lines = [
        "# 담당자별 객관자료 완료",
        "",
        f"- 갱신: {datetime.now():%Y-%m-%d %H:%M}",
        "- 판정: 사람 체크가 아니라 완료일·원천 증빙·ERP·PO 대조로 확인된 건만 기록",
        "- Excel 상태 셀은 수정하지 않으며 DB가 완료 정본",
        "",
    ]
    for owner in ("류지영", "오종현", "유현민"):
        lines.append(f"## {owner} 완료 · {counts.get(owner, 0)}건")
        items = [row for row in rows if row["owner"] == owner][:30]
        if not items:
            lines.append("- 새로 입증된 완료 없음")
        for row in items:
            target = row.get("project") or row.get("record_id")
            lines.append(
                f"- {row['completed_on']} · {target} · {row['task_kind']} · {row['basis']}"
            )
        if counts.get(owner, 0) > len(items):
            lines.append(f"- 외 {counts[owner] - len(items)}건(DB 보관)")
        lines.append("")
    text = "\n".join(lines).rstrip() + "\n"
    tmp = path + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as out:
        out.write(text)
        out.flush()
        os.fsync(out.fileno())
    os.replace(tmp, path)
    return counts


def sync(po_evidence_path=PO_EVIDENCE, report_path=REPORT_PATH):
    import ledger_db

    entries = objective_entries(po_evidence_path)
    ledger_db.staff_resolution_sync(entries)
    ledger_db.staff_resolution_retract(objective_retractions(po_evidence_path))
    rows = ledger_db.staff_resolutions()
    counts = write_report(rows, report_path)
    return {owner: counts.get(owner, 0) for owner in ("류지영", "오종현", "유현민")}


def main():
    do_sync = "--sync" in sys.argv
    if do_sync:
        counts = sync()
    else:
        entries = objective_entries()
        counts = Counter(row["owner"] for row in entries)
        print("미리보기 — DB에 기록하려면 --sync")
    print(" · ".join(f"{owner} 완료 {counts.get(owner, 0)}건"
                     for owner in ("류지영", "오종현", "유현민")))


if __name__ == "__main__":
    main()
