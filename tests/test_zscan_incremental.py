# -*- coding: utf-8 -*-
"""zscan --docs 증분 체크포인트 회귀검사 ([490]).

실 Z:·관리대장·ERP·밴드에는 닿지 않는다. 임시 폴더의 파일명만 읽는다.
"""
import json
import os
import sys
import tempfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import zscan
import autopilot


def _touch(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(b"pdf-name-only-test")


def main():
    with tempfile.TemporaryDirectory(prefix="zscan-490-") as td:
        root = os.path.join(td, "zroot")
        state = os.path.join(td, "progress.json")
        report = os.path.join(td, "Z폴더_서류대조.md")
        _touch(os.path.join(root, "a", "2026-06-19 A캠프 거래명세서.pdf"))
        _touch(os.path.join(root, "b", "2026.06.20 B캠프 세금계산서.PDF"))
        _touch(os.path.join(root, "b", "nested", "2026_06_21 C캠프 계산서.pdf"))
        _touch(os.path.join(root, "b", "날짜없는 거래명세서.pdf"))
        _touch(os.path.join(root, "b", "2026-13-01 잘못된날짜 거래명세서.pdf"))
        with open(report, "w", encoding="utf-8") as f:
            f.write("완전한 옛 리포트\n")

        calls = [0]

        def stop_after_root():
            calls[0] += 1
            return calls[0] > 1

        docs, complete, progress = zscan.doc_catalog_incremental(
            root, state_path=state, stop=stop_after_root)
        assert not complete and docs == [], (complete, docs)
        assert progress["scanned_dirs"] == 1 and progress["pending_dirs"], progress
        assert open(report, encoding="utf-8").read() == "완전한 옛 리포트\n", \
            "증분 중 기존 완전한 리포트를 덮었다"
        saved = json.load(open(state, encoding="utf-8"))
        assert saved["pending_dirs"] and saved["scanned_dirs"] == 1, saved

        docs, complete, progress = zscan.doc_catalog_incremental(
            root, state_path=state, stop=lambda: False)
        assert complete and not progress["pending_dirs"], progress
        assert len(docs) == 3, [(d["일자"], d["파일"]) for d in docs]
        assert len({(d["폴더"], d["파일"]) for d in docs}) == 3, "재개 뒤 서류가 중복됐다"

        zscan._write_docs_report(report, docs, 999, [], docs, [])
        zscan._clear_doc_progress(state)
        text = open(report, encoding="utf-8").read()
        assert "서류 PDF 3개 · 관리대장 v999" in text, text[:200]
        assert not os.path.exists(state), "완주 뒤 체크포인트가 남았다"
        assert not [p for p in os.listdir(td) if p.endswith(".tmp")], \
            "원자 교체 임시파일이 남았다"

        # main()의 미완주 계약은 0(완료)·1(실패)이 아니라 정확히 75다.
        original = zscan.doc_catalog_incremental
        old_argv = sys.argv[:]
        try:
            zscan.doc_catalog_incremental = lambda _root: (
                [], False, {"scanned_dirs": 7, "scanned_files": 80,
                            "pending_dirs": ["left"], "docs": []})
            sys.argv = ["zscan.py", "--docs", "--root", root]
            assert zscan.main() == zscan.INCREMENTAL_RETURN_CODE == 75
        finally:
            zscan.doc_catalog_incremental = original
            sys.argv = old_argv

    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    docs_cmd = [sys.executable, os.path.join(ROOT, "zscan.py"), "--docs"]
    scan_cmd = [sys.executable, os.path.join(ROOT, "zscan.py")]
    assert autopilot._child_budget_key(docs_cmd) == "ZSCAN_BUDGET_SEC", \
        "자율복구가 zscan --docs 예산을 전달하지 않는다"
    assert autopilot._child_budget_key(scan_cmd) is None, \
        "예산을 읽지 않는 일반 zscan까지 증분 명령으로 잘못 분류했다"
    assert '"ZSCAN_BUDGET_SEC": os.environ.get("ZSCAN_BUDGET_SEC") or "1500"' in daily, \
        "일일회차가 30분 부모 제한 전에 zscan을 돌려보내지 않는다"
    print("[490] zscan --docs 체크포인트 재개·반쪽 리포트 금지·75·예산 배선 OK")


if __name__ == "__main__":
    main()
