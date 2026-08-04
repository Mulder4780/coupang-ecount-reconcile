# -*- coding: utf-8 -*-
"""
collect_all.py — "긁어와" · "자료 취합해" 한 마디로 도는 수집·보관·대조 일괄 실행

사용자 지시(2026-08-05): "내가 '긁어와'나 '자료 취합해' 이런 명령 내리면 바로 실행."
그래서 사람도 AI 도 **이 파일 하나만 부르면** 되게 묶었다. 브라우저가 필요한 단계
(ERP 화면·밴드 상세 수집)는 여기 없다 — 그건 로그인된 크롬이 있어야 하므로 세션 AI 가
한다. 이 스크립트는 **이미 들어온 원본을 남김없이 흡수·보관·대조**하는 부분을 맡는다.

단계
  1. 투입함/다운로드 흡수      upload_intake --apply · download_intake --apply
  2. 밴드 덤프 → 캐시 병합      band/convert_dump.py
  3. 밴드 사진 전량 내려받기    band/fetch_images.py --all
  4. 밴드 게시글 보관(PDF·텍스트·사진)  band/archive_posts.py
  5. ERP 명세서 파싱·색인·건별 PDF     stmt_docs · stmt_link · stmt_archive
  6. ERP 엑셀 → PDF 사본        erp_pdf_export.py
  7. 대조                       erp_docs_check · band/band_reconcile · settlement_completion --sync
  8. 확정분 DB 적재             ledger_db --intake      (엑셀 반영은 11:00·15:00 규칙 그대로)

  python collect_all.py              # 전부
  python collect_all.py --quick      # 무거운 보관 단계(3·4·5) 건너뜀
  python collect_all.py --archive-limit 300
"""
import argparse
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def run(name, args, timeout=2400):
    t0 = time.time()
    try:
        r = subprocess.run([PY] + args, cwd=ROOT, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout)
        out = [ln for ln in (r.stdout or "").splitlines()
               if ln.strip() and "UserWarning" not in ln and "warn(" not in ln
               and "관리대장 최신본" not in ln]
        tail = out[-1] if out else ""
        ok = r.returncode == 0
    except subprocess.TimeoutExpired:
        ok, tail = False, f"시간초과({timeout}s)"
    except Exception as e:
        ok, tail = False, str(e)[:80]
    print(f"[{'OK ' if ok else 'FAIL'}] {name} ({int(time.time()-t0)}s) {tail[:110]}")
    return {"name": name, "ok": ok, "out": tail}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    ap.add_argument("--archive-limit", type=int, default=200)
    a = ap.parse_args()

    steps = []
    P = lambda *x: os.path.join(ROOT, *x)
    print(f"=== 자료 취합 시작 {time.strftime('%Y-%m-%d %H:%M')} ===")
    steps.append(run("투입함 원본 분류", [P("upload_intake.py"), "--apply"]))
    steps.append(run("다운로드 흡수", [P("download_intake.py"), "--apply"]))
    steps.append(run("밴드 덤프 → 캐시", [P("band", "convert_dump.py")]))
    if not a.quick:
        steps.append(run("밴드 사진 전량", [P("band", "fetch_images.py"), "--all"]))
        steps.append(run("밴드 게시글 보관(PDF·텍스트·사진)",
                         [P("band", "archive_posts.py"), "--limit", str(a.archive_limit)]))
    steps.append(run("ERP 명세서 파싱", [P("stmt_docs.py")]))
    steps.append(run("명세서 ↔ 프로젝트 색인", [P("stmt_link.py")]))
    if not a.quick:
        steps.append(run("명세서 건별 PDF",
                         [P("stmt_archive.py"), "--limit", str(a.archive_limit)]))
    steps.append(run("ERP 엑셀 PDF 사본", [P("erp_pdf_export.py")]))
    steps.append(run("ERP 매출서류 대조", [P("erp_docs_check.py")]))
    steps.append(run("밴드 대조", [P("band", "band_reconcile.py")]))
    steps.append(run("정산 객관완료 동기화", [P("settlement_completion.py"), "--sync"]))
    steps.append(run("확정분 DB 적재", [P("ledger_db.py"), "--intake"]))

    bad = [s for s in steps if not s["ok"]]
    print(f"=== 완료: {len(steps)-len(bad)}/{len(steps)} 성공 ===")
    for s in bad:
        print(f"  ! {s['name']} — {s['out'][:90]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
