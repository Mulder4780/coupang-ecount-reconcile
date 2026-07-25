# -*- coding: utf-8 -*-
"""
daily_run.py — 쿠팡 업무 자동대조 에이전트 (일일 오케스트레이터)
==================================================================
매일 1회(권장 07:50, 08:30 대표보고 전) 전체 파이프라인을 안전한 순서로 실행하고
reports/종합리포트_*.md 한 장으로 요약한다. Windows 작업 스케줄러에 daily_run.bat 등록 시 완전 자동.

원칙:
  - 0단계 합성검증(ALL GREEN) 실패 시 전체 중단 (사용자 상시 지시)
  - ERP 쓰기(--post)는 절대 자동 실행하지 않음 — 전송 대기 건수만 보고
  - 각 단계는 데이터가 없으면 조용히 건너뜀(스킵 사유 기록) — 있는 데이터만큼 검증
"""
import sys, os, glob, subprocess
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
REPORT_DIR = os.path.join(ROOT, "reports")
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}


def run(name, args, timeout=600):
    try:
        r = subprocess.run([PY] + args, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", cwd=ROOT, timeout=timeout, env=ENV)
        out = (r.stdout or "") + (("\n[stderr] " + r.stderr[:500]) if r.returncode != 0 and r.stderr else "")
        # 토큰 절약: 노이즈 제거 후 요약만 보존(상세는 각 모듈이 reports/에 파일로 남긴다)
        keep = [ln for ln in out.splitlines()
                if ln.strip() and not any(x in ln for x in
                    ("UserWarning", "warn(msg)", "  [예정]", "  [건너뜀]", "i 관리대장 최신본"))]
        return {"name": name, "ok": r.returncode == 0, "out": "\n".join(keep[-12:]).strip()}
    except subprocess.TimeoutExpired:
        return {"name": name, "ok": False, "out": f"시간초과({timeout}s)"}


def main():
    steps = []

    # 0. 합성검증 — 실패 시 전체 중단
    s = run("합성검증", [os.path.join(ROOT, "tests", "synthetic_check.py")])
    steps.append(s)
    if not s["ok"] or "ALL GREEN" not in s["out"]:
        finish(steps, aborted=True)
        sys.exit("합성검증 실패 — 전체 중단")

    # 1. 판매·세금계산서 inbox 대조 (inbox 없으면 원장 준비표)
    steps.append(run("판매·세금계산서 대조", [os.path.join(ROOT, "ecount_reconcile.py")]))

    # 2. ERP 계정별원장 4유형 대조 (inbox에 '원장' 파일 있을 때만)
    if [f for f in glob.glob(os.path.join(ROOT, "inbox", "*.xlsx"))
            if "원장" in os.path.basename(f) or "계정" in os.path.basename(f)]:
        steps.append(run("ERP원장 4유형 대조", [os.path.join(ROOT, "erp_ledger_check.py")]))
    else:
        steps.append({"name": "ERP원장 4유형 대조", "ok": None, "out": "스킵 — inbox/에 계정별원장 파일 없음"})

    # 2.5 쿠팡 PO 대조 (inbox에 'PO' 파일 있을 때만)
    if [f for f in glob.glob(os.path.join(ROOT, "inbox", "*.xlsx"))
            if "PO" in os.path.basename(f).upper() and "원장" not in os.path.basename(f)]:
        steps.append(run("쿠팡 PO 대조", [os.path.join(ROOT, "po_reconcile.py")]))
    else:
        steps.append({"name": "쿠팡 PO 대조", "ok": None, "out": "스킵 — inbox/에 쿠팡 PO 목록 파일 없음(파일명에 PO 포함)"})

    # 3. 밴드 수집·대조 — 공식 API 토큰이 있으면 수집+대조, 브라우저 수집 캐시만 있으면 대조만
    band_cache = [f for f in glob.glob(os.path.join(ROOT, "band", "cache", "*.json"))
                  if not os.path.basename(f).startswith(("raw_", "dump_"))]
    if os.path.exists(os.path.join(ROOT, "band", ".band_token.json")):
        steps.append(run("밴드 수집", [os.path.join(ROOT, "band", "band_sync.py")]))
        steps.append(run("밴드 대조", [os.path.join(ROOT, "band", "band_reconcile.py")]))
    elif band_cache:
        steps.append(run("밴드 대조(캐시)", [os.path.join(ROOT, "band", "band_reconcile.py")]))
    else:
        steps.append({"name": "밴드 수집·대조", "ok": None, "out": "스킵 — 밴드 미인증·캐시 없음(앱 심사 대기)"})

    # 4. 카톡 대조 (kakao/inbox에 txt 있을 때만)
    if glob.glob(os.path.join(ROOT, "kakao", "inbox", "*.txt")):
        steps.append(run("카톡 대조", [os.path.join(ROOT, "kakao", "kakao_reconcile.py")]))
    else:
        steps.append({"name": "카톡 대조", "ok": None, "out": "스킵 — kakao/inbox/에 대화 내보내기 txt 없음"})

    # 5. 확정 업데이트 자동 반영 — 빈 칸만·근거 보유·항상 새 버전(vN+1) 생성이라 안전
    # ERP 매출서류(계산서·명세서 현황) — 있으면 대조 + 25시트 반영
    try:
        sys.path.insert(0, ROOT)
        from inbox_scan import pick as _pick
        _has_tax = bool(_pick("tax"))
    except Exception:
        _has_tax = False
    if _has_tax:
        steps.append(run("ERP 매출서류 대조(25시트)", [os.path.join(ROOT, "erp_docs_check.py"), "--sheet"]))
    else:
        steps.append({"name": "ERP 매출서류 대조", "ok": None,
                      "out": "스킵 — inbox/에 매출(세금)계산서현황 없음"})

    # 밴드 문서 사진(거래명세서·세금계산서) OCR → 확실한 건은 빈칸 입력 큐에 적재
    _docs = os.path.join(ROOT, "band", "docs_inbox")
    if os.path.isdir(_docs) and any(f.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp"))
                                    for f in os.listdir(_docs)):
        steps.append(run("밴드 문서 이미지 대조·입력", [os.path.join(ROOT, "band", "doc_ocr.py"),
                                                       "--scan", "--apply"]))
    else:
        steps.append({"name": "밴드 문서 이미지 대조", "ok": None,
                      "out": "스킵 — band/docs_inbox/에 사진 없음"})

    steps.append(run("관리대장 자동입력(확정분)", [os.path.join(ROOT, "ledger_writer.py"), "--apply"]))
    # 입력 직후 무결성 확인 — 엑셀이 '복구' 대화상자를 띄우는 파일을 만들지 않기 위해
    steps.append(run("워크북 무결성 검사·복구", [os.path.join(ROOT, "fix_workbook.py"), "--apply"]))

    # 5.5 밴드 업무 추출 → 24_밴드업무추출 시트 (월별 백필 원천, 캐시 있을 때만)
    if band_cache:
        steps.append(run("밴드 업무추출(24시트)", [os.path.join(ROOT, "band_extract.py"), "--sheet"]))

    # 6. 확인필요현황 시트 갱신 — 관리대장 본체 23_확인필요현황 (변경 시에만 vN+1, 단일 엑셀 통합관리)
    steps.append(run("확인필요 시트 갱신(23시트)", [os.path.join(ROOT, "findings_sheet.py")]))

    # 7. 전표 전송 대기 현황 (dry-run만 — 실전송은 절대 자동화하지 않음)
    steps.append(run("전표 전송대기(dry-run)", [os.path.join(ROOT, "ecount_upload.py")]))

    finish(steps)


def finish(steps, aborted=False):
    os.makedirs(REPORT_DIR, exist_ok=True)
    # 동기화 백본: 앱(웹·워크벤치)이 읽는 기계 판독용 상태 파일 — 에이전트가 유일한 작성자
    import json
    json.dump({"time": datetime.now().isoformat(), "aborted": aborted,
               "steps": [{"n": s["name"], "s": ("ok" if s["ok"] else ("skip" if s["ok"] is None else "fail"))}
                          for s in steps]},
              open(os.path.join(REPORT_DIR, "agent_status.json"), "w", encoding="utf-8"), ensure_ascii=False)
    base = os.path.join(REPORT_DIR, f"종합리포트_{datetime.now():%Y%m%d_%H%M}.md")
    with open(base, "w", encoding="utf-8") as f:
        f.write(f"# 쿠팡 업무 자동대조 종합리포트 — {datetime.now():%Y-%m-%d %H:%M}\n\n")
        if aborted:
            f.write("**★ 합성검증 실패로 중단 — 아래 로그 확인 후 코드 수정 필요**\n\n")
        f.write("| 단계 | 결과 |\n|---|---|\n")
        for s in steps:
            mark = "✅" if s["ok"] else ("⏭ 스킵" if s["ok"] is None else "❌ 실패")
            f.write(f"| {s['name']} | {mark} |\n")
        f.write("\n---\n")
        for s in steps:
            f.write(f"\n## {s['name']}\n```\n{s['out']}\n```\n")
    print(f"\n종합리포트: {base}")
    for s in steps:
        mark = "OK " if s["ok"] else ("SKIP" if s["ok"] is None else "FAIL")
        print(f"  [{mark}] {s['name']}")


if __name__ == "__main__":
    main()
