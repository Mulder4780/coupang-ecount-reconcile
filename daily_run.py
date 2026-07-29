# -*- coding: utf-8 -*-
"""
daily_run.py — 쿠팡 업무 자동대조 에이전트 (일일 오케스트레이터)
==================================================================
매일 1회(09:50, 류지영 입력 종료 후) 전체 파이프라인을 안전한 순서로 실행하고
reports/종합리포트_*.md 한 장으로 요약한다. Windows 작업 스케줄러에 daily_run.bat 등록 시 완전 자동.

원칙:
  - 0단계 합성검증(ALL GREEN) 실패 시 전체 중단 (사용자 상시 지시)
  - ERP 쓰기(--post)는 절대 자동 실행하지 않음 — 전송 대기 건수만 보고
  - 각 단계는 데이터가 없으면 조용히 건너뜀(스킵 사유 기록) — 있는 데이터만큼 검증
"""
import sys, os, glob, subprocess
from datetime import datetime
from operation_window import input_window_label, is_input_window

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
    if is_input_window():
        print(f"입력 보호시간({input_window_label()}) — 일일 자동대조를 시작하지 않습니다.")
        return
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
    # PC가 꺼진 동안 휴대폰에서 예약한 코드를 영구 큐에서 안전하게 가져온다.
    # 반영 성공 확인 전에는 서버 항목을 지우지 않으며 입력 보호시간에는 이 단계도 멈춘다.
    steps.append(run("휴대폰 클라우드 예약 반영", [os.path.join(ROOT, "cloud_queue_sync.py")]))

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

    # 2.7 입금(수금) 자동입력 — 자료가 들어오면 사람 손 없이 채운다(사용자 지시 2026-07-28).
    #     파일이 없으면 receipt_fill 이 스스로 안내만 하고 조용히 끝나므로 조건 없이 돌린다.
    #     (계정별원장은 '0. 원본 자료' 에도 들어오므로 파일명 조건을 걸면 놓친다 — pick 이 내용으로 찾는다)
    steps.append(run("입금 대조·자동입력", [os.path.join(ROOT, "receipt_fill.py"), "--queue"]))

    # 2.8 카톡 신규 접수 등록 — 대화 내보내기가 들어오면 02·04 에 새 행으로 올린다.
    #     유형이 확정된 것(돌발·정기)만 올리고 철거·납품은 보류한다 — 대상 시트가 없다.
    if glob.glob(os.path.join(ROOT, "kakao", "inbox", "*.txt")):
        #     ★ --days 7 로 **최근분만** 올린다. 과거분(현재 미등록 82건)은 파싱 품질을 사람이
        #       한 번 보고 넣어야 해서 자동에 태우지 않는다: python kakao_extract.py --new
        steps.append(run("카톡 신규 접수 등록",
                         [os.path.join(ROOT, "kakao_extract.py"), "--new", "--days", "7", "--queue"]))
    else:
        steps.append({"name": "카톡 신규 접수 등록", "ok": None, "out": "스킵 — kakao/inbox/에 대화 내보내기 없음"})

    # 2.9 작업 내용 자동 기입 — 원문(카톡·밴드 **완료 글**)에 있으면 03시트에 채운다.
    #     사용자 지시(2026-07-29): "밴드나 카톡에서 확인되면 자동 기입. 모든 데이터가 마찬가지."
    #     ★ '신청내용'(요청)은 쓰지 않는다 — 그건 무엇을 했나가 아니다. 빈 양식도 거른다.
    steps.append(run("작업내용 자동기입(03시트)", [os.path.join(ROOT, "fill_work_detail.py"), "--apply"]))

    # 2.95 구글 캘린더 대조 — 사용자 지시(2026-07-29): "이 캘린더 추가하고 항상 대조해서
    #      엑셀과 앱에 반영해줘". 캘린더는 **예정**이므로 예정일 칸만 채우고 완료 칸은 안 건드린다.
    #      원천(비공개 iCal 주소·.ics 파일)이 없으면 gcal_sync 가 스스로 조용히 끝난다 —
    #      주소 하나 없다고 일일 파이프라인 전체가 실패로 물들면 안 된다.
    steps.append(run("구글 캘린더 대조", [os.path.join(ROOT, "gcal_sync.py"), "--queue"]))

    # 2.96 재계산 대기 세기 — 원장엔 올라왔는데 엑셀이 아직 계산 안 해 앱에 안 나오는 건.
    #      숫자가 틀린 게 아니라 대기 중이라는 걸 앱이 스스로 말하게 한다(사용자 오해 방지).
    steps.append(run("재계산 대기 확인", [os.path.join(ROOT, "recalc_pending.py")]))

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
        # 사진 1장에 OCR 1초 남짓. 첫 회는 1,458장에 25분이 걸려 600초 제한에 걸렸다(2026-07-26).
        # doc_ocr가 결과를 band/ocr_cache/에 저장하므로 두 번째 실행부터는 몇 초로 끝난다.
        steps.append(run("밴드 문서 이미지 대조·입력", [os.path.join(ROOT, "band", "doc_ocr.py"),
                                                       "--scan", "--apply"], timeout=3600))
    else:
        steps.append({"name": "밴드 문서 이미지 대조", "ok": None,
                      "out": "스킵 — band/docs_inbox/에 사진 없음"})

    # 완료보고서와 문서발행 표시는 프로젝트NO가 정확히 일치하는 근거만 빈칸에 큐잉한다.
    # 실제 ZIP 패치는 바로 아래 ledger_writer 한 번으로 합쳐 버전 난립과 충돌을 막는다.
    steps.append(run("완료보고서 확정분 큐", [os.path.join(ROOT, "confirm_fill.py"), "--queue"]))
    steps.append(run("ERP 판매전표·거래명세 확정분 큐",
                     [os.path.join(ROOT, "fill_erp_documents.py"), "--queue"]))

    steps.append(run("관리대장 자동입력(확정분)", [os.path.join(ROOT, "ledger_writer.py"), "--apply"]))
    # 입력 직후 무결성 확인 — 엑셀이 '복구' 대화상자를 띄우는 파일을 만들지 않기 위해
    steps.append(run("워크북 무결성 검사·복구", [os.path.join(ROOT, "fix_workbook.py"), "--apply"]))

    # 5.3 류지영 정기점검 스케줄 원본 → 27_정기점검원본일정.
    #     지정 폴더의 최신 xlsx를 매일 다시 읽으며, 내용이 같으면 새 버전을 만들지 않는다.
    #     원본에 UJ번호가 없으므로 임의 프로젝트를 만들지 않고 04시트와 근거가 있는 건만 연결한다.
    steps.append(run("정기점검 스케줄 원본 자동반영",
                     [os.path.join(ROOT, "pm_schedule_sync.py"), "--apply"]))

    # 5.5 밴드 업무 추출 → 24_밴드업무추출 시트 (월별 백필 원천, 캐시 있을 때만)
    if band_cache:
        steps.append(run("밴드 업무추출(24시트)", [os.path.join(ROOT, "band_extract.py"), "--sheet"]))

    # 6. 확인필요현황 시트 갱신 — 관리대장 본체 23_확인필요현황 (변경 시에만 vN+1, 단일 엑셀 통합관리)
    steps.append(run("확인필요 시트 갱신(23시트)", [os.path.join(ROOT, "findings_sheet.py")]))

    # 6.5 확인필요현황 **별도 엑셀** 갱신 — 23시트는 평면 목록이라 유형별 상세열
    #     (명세서번호·PO번호·매칭근거·확인방법)을 담지 못한다. 그 상세는 이 파일에만 있다.
    #     ★ 여기에 연결돼 있지 않아 2026-07-27 16:45 자로 하루 동안 멈춰 있었다 —
    #       읽는 사람은 멈춘 줄 모르고 어제 숫자를 오늘 숫자로 본다. 반드시 매일 같이 돈다.
    steps.append(run("확인필요현황 엑셀 갱신", [os.path.join(ROOT, "findings_export.py")]))

    # 6.7 자료현황 한 장 — "밴드에서 뭘 얼마나 가져왔나 / 지금 뭘 갖고 있나 / 원장이 얼마나 찼나".
    #     같은 질문을 매번 다시 세지 않으려고 만든다(사용자 지시 2026-07-29).
    #     느린 것(Z: 2만 개 순회)은 다시 돌지 않고 앞 단계가 남긴 리포트에서 숫자만 읽는다 —
    #     그래서 이 단계는 위 대조들이 **끝난 뒤에** 와야 한다.
    steps.append(run("자료현황 갱신", [os.path.join(ROOT, "data_status.py")]))

    # 7. 전표 전송 대기 현황 (dry-run만 — 실전송은 절대 자동화하지 않음)
    steps.append(run("전표 전송대기(dry-run)", [os.path.join(ROOT, "ecount_upload.py")]))

    # 8. 클라우드 사본 갱신 — PC가 꺼져도 폰에서 오늘 자료를 볼 수 있게.
    #    config/cloud.json 이 없으면 조용히 건너뛴다(설정 전에는 아무 일도 안 함).
    if os.path.exists(os.path.join(ROOT, "config", "cloud.json")):
        steps.append(run("클라우드 사본 올리기", [os.path.join(ROOT, "cloud_export.py"), "--upload"]))
    else:
        steps.append({"name": "클라우드 사본 올리기", "ok": None,
                      "out": "스킵 — config/cloud.json 미설정(CLOUD_SETUP.md 참고)"})

    # 9. 폰용 사본 — PC가 꺼져 있어도 열리는 HTML 한 장(서버·인터넷 불필요).
    #    이동이 잦아 PC를 켜 둘 수 없을 때 이게 유일하게 확실한 방법이다.
    steps.append(run("폰용 사본 만들기", [os.path.join(ROOT, "mobile_snapshot.py")]))

    # 9-b. 대표 보고용 내용 브리핑 — 대표 지시(2026-07-28): 숫자가 아니라 '무슨 일이 있었나'
    steps.append(run("대표 브리핑(내용)", [os.path.join(ROOT, "daily_brief.py"), "--md"]))

    # 10. 고정 주소 사본 — PC를 꺼도 폰·태블릿이 이걸로 조회·자동채움을 한다(잠가서 올린다)
    steps.append(run("고정 주소 사본 올리기", [os.path.join(ROOT, "cloud_publish.py"), "--push"]))

    # 11. 버전 파일 정리 — 최신본 하나만 작업 폴더에 두고 구버전은 사용자가 지정한
    #     OLD/ 한 곳으로 옮긴다. 같은 이름이 있어도 덮어쓰거나 삭제하지 않는다.
    steps.append(run("관리대장 버전 정리", [os.path.join(ROOT, "ledger_versions.py"), "--prune"]))

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
