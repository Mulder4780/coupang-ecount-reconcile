# -*- coding: utf-8 -*-
"""확인 필요의 객관근거를 안전하게 다시 대조하는 앱용 단일 회차.

실제 세금계산서 발행·외부 메시지·Excel 저장은 하지 않는다. 각 단계는 Windows에서도
끝나는 ``proc_guard``로 분리하고, 한 단계가 실패해도 나머지 진단은 계속 남긴다.
"""
from __future__ import annotations

import os
import sys

from proc_guard import run_tree

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

STEPS = (
    ("객관완료 DB 동기화", ["settlement_completion.py", "--sync"], 1500),
    ("ERP 원장 매칭키·프로젝트 재대조", ["erp_ledger_check.py"], 1500),
    ("세금계산서 실제 미발행 교차확인", ["tax_invoice_watch.py"], 900),
)


def main() -> int:
    failures = []
    for title, args, timeout in STEPS:
        print(f"\n===== {title} =====", flush=True)
        result = run_tree([PY, os.path.join(ROOT, args[0]), *args[1:]], cwd=ROOT,
                          timeout=timeout, drain_timeout=30, output_limit=120_000)
        if result.stdout:
            print(result.stdout.rstrip(), flush=True)
        if result.stderr:
            print(result.stderr.rstrip(), flush=True)
        if result.returncode or result.timed_out:
            why = "시간초과" if result.timed_out else f"종료코드 {result.returncode}"
            failures.append(f"{title}: {why}")
            print(f"! {why} — 다음 단계는 계속합니다", flush=True)
    if failures:
        print("\n일부 단계 실패: " + " · ".join(failures), flush=True)
        return 1
    print("\n근거 재대조 완료 — 실제 발행·외부 전송·Excel 저장 없음", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
