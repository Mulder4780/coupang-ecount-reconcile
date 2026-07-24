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
        return {"name": name, "ok": r.returncode == 0, "out": out.strip()[-1500:]}
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

    # 3. 밴드 수집·대조 (토큰 있을 때만)
    if os.path.exists(os.path.join(ROOT, "band", ".band_token.json")):
        steps.append(run("밴드 수집", [os.path.join(ROOT, "band", "band_sync.py")]))
        steps.append(run("밴드 대조", [os.path.join(ROOT, "band", "band_reconcile.py")]))
    else:
        steps.append({"name": "밴드 수집·대조", "ok": None, "out": "스킵 — 밴드 미인증(앱 심사 대기)"})

    # 4. 카톡 대조 (kakao/inbox에 txt 있을 때만)
    if glob.glob(os.path.join(ROOT, "kakao", "inbox", "*.txt")):
        steps.append(run("카톡 대조", [os.path.join(ROOT, "kakao", "kakao_reconcile.py")]))
    else:
        steps.append({"name": "카톡 대조", "ok": None, "out": "스킵 — kakao/inbox/에 대화 내보내기 txt 없음"})

    # 5. 전표 전송 대기 현황 (dry-run만 — 실전송은 절대 자동화하지 않음)
    steps.append(run("전표 전송대기(dry-run)", [os.path.join(ROOT, "ecount_upload.py")]))

    finish(steps)


def finish(steps, aborted=False):
    os.makedirs(REPORT_DIR, exist_ok=True)
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
