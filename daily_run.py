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
import sys, os, glob, json, subprocess, time, uuid
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
RUN_LOCK = os.path.join(REPORT_DIR, ".daily_run.lock")


def _pid_alive(pid):
    """같은 PC의 프로세스가 살아 있는가. 죽은 잠금만 안전하게 회수한다.

    ★ 판정은 pid_alive.py 한 곳에서 한다 (2026-08-06 실사고 · 검증 [121]).
      여기 있던 옛 판정은 윈도우에서 **이미 끝난 프로세스도 살아 있다고** 했다
      (OpenProcess 는 종료된 프로세스에도 핸들을 준다). 그 탓에 잠금이 스스로 풀릴
      길이 없어져 daily_run 이 밤새 한 번도 못 돌았다.
      모르면(None) '살아 있다'로 본다 — 남의 회차를 밀어내는 쪽이 더 위험하다.
      다만 pid 자체가 없거나 망가진 잠금 파일은 회수 대상이다.
    """
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    try:
        import pid_alive
    except Exception:
        return True                      # 판정 못 하면 건드리지 않는다
    return pid_alive.alive(pid) is not False


def acquire_run_lock(path=RUN_LOCK):
    """프로세스 간 단발 잠금. 성공 시 소유 토큰, 중복 실행이면 None."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    token = f"{os.getpid()}:{time.time_ns()}:{uuid.uuid4().hex}"
    payload = {
        "pid": os.getpid(),
        "token": token,
        "started_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    for _attempt in range(2):
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                with open(path, encoding="utf-8") as f:
                    owner = json.load(f)
            except (OSError, ValueError, TypeError):
                owner = {}
            # O_EXCL로 파일을 만든 직후 JSON을 쓰는 아주 짧은 동안에는 다른 실행이
            # 빈/부분 파일을 볼 수 있다. 최근의 손상 파일은 선점 중으로 보고 건드리지 않는다.
            if not owner:
                try:
                    if time.time() - os.path.getmtime(path) < 60:
                        return None
                except OSError:
                    return None
            if _pid_alive(owner.get("pid")):
                return None
            try:
                os.unlink(path)
            except OSError:
                return None
            continue
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
            f.write("\n")
        return token
    return None


def release_run_lock(token, path=RUN_LOCK):
    """자기가 만든 잠금만 놓는다. 다른 실행의 잠금을 지우지 않는다."""
    try:
        with open(path, encoding="utf-8") as f:
            owner = json.load(f)
        if owner.get("token") == token:
            os.unlink(path)
    except (OSError, ValueError, TypeError):
        pass


def has_inbox_kind(kind):
    """무작위 파일명도 놓치지 않도록 엑셀 머리글 내용으로 원천을 찾는다."""
    try:
        from inbox_scan import pick
        return bool(pick(kind))
    except Exception:
        return False


def _source_files(dir_fn, extensions, name_prefixes=()):
    """source_dirs의 정본 폴더를 재귀 탐색한다. 투입함은 분류 후에만 읽는다."""
    try:
        import source_dirs
        folders = getattr(source_dirs, dir_fn)()
    except Exception:
        return []
    out, seen = [], set()
    for folder in folders:
        for base, _dirs, files in os.walk(folder):
            for name in files:
                low = name.lower()
                if extensions and not low.endswith(extensions):
                    continue
                if name_prefixes and not low.startswith(name_prefixes):
                    continue
                path = os.path.join(base, name)
                key = os.path.normcase(os.path.abspath(path))
                if key not in seen:
                    seen.add(key)
                    out.append(path)
    return out


# 되풀이하면 안 되는 단계 — 큐에 넣거나 파일을 옮기거나 원장을 건드리는 것들.
# ★ 호출부마다 retry=0 을 적게 하지 않는다 — 새 단계를 넣는 사람이 반드시 잊는다.
#   여기서 **인자를 보고** 자동으로 정한다.
NO_RETRY_FLAGS = ("--queue", "--apply", "--intake", "--force")
NO_RETRY_NAMES = ("cloud_queue_sync.py", "ledger_db.py", "workbook_patch.py",
                  "ledger_writer.py", "expand_rows.py",
                  # 게시는 git commit·push 다. 실패한 자리를 모르는 채 다시 하면
                  # 빈 커밋이나 반쯤 올라간 상태가 겹친다 — 다음 회차에 맡긴다.
                  "cloud_publish.py")


def _retryable(args):
    joined = " ".join(str(a) for a in args)
    if any(f in args for f in NO_RETRY_FLAGS):
        return False
    return not any(n in joined for n in NO_RETRY_NAMES)


def run(name, args, timeout=600, retry=None):
    """한 단계를 돌린다. 실패하면 **한 번만** 쉬었다 다시 해 본다.

    ★ 2026-08-05 — 14:02 회차에서 5단계가 실패했는데, 손으로 다시 돌리니 전부 정상이었다.
      원인은 결함이 아니라 **경합**이다(Z: 대량 보관 작업·엑셀 점유와 겹친 순간).
      이런 것은 한 번 쉬었다 하면 지나간다. 그런데 그동안은 그대로 '실패'로 남아,
      진짜 결함(오늘 잡은 stmt_link 회귀 같은 것)이 잡음에 묻혔다.
      ※ 되풀이해도 되는 것만 재시도한다 — 큐·반영·파일 이동 단계는 자동으로 제외된다.
    """
    if retry is None:
        retry = 1 if _retryable(args) else 0
    for attempt in range(retry + 1):
        got = _run_once(name, args, timeout)
        if got["ok"] or attempt >= retry:
            if not got["ok"] and attempt:
                got["out"] = (got["out"] + "\n[재시도 후에도 실패]").strip()
            return got
        time.sleep(20)          # 상대가 파일을 놓을 시간을 준다
    return got


def _kill_tree(p):
    """Windows 는 `kill()` 로 **자식의 자식까지 안 죽인다** — 나무째 끊는다."""
    if os.name == "nt":
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(p.pid)],
                           capture_output=True, timeout=30)
        except Exception:
            pass
    try:
        p.kill()
    except Exception:
        pass


def _tidy(out):
    # 토큰 절약: 노이즈 제거 후 요약만 보존(상세는 각 모듈이 reports/에 파일로 남긴다)
    keep = [ln for ln in out.splitlines()
            if ln.strip() and not any(x in ln for x in
                ("UserWarning", "warn(msg)", "  [예정]", "  [건너뜀]", "i 관리대장 최신본"))]
    return "\n".join(keep[-12:]).strip()


def _run_once(name, args, timeout):
    """★ `subprocess.run(timeout=)` 을 쓰면 안 된다 — **윈도우에서 영원히 멈춘다**
       (2026-08-08 실사고).

    CPython 의 `subprocess.run` 은 시간이 넘으면 `process.kill()` 을 부른 뒤
    윈도우에서만 **시간제한 없는** `process.communicate()` 를 한 번 더 부른다.
    자식이 SMB(Z:) 읽기처럼 **끊기지 않는 대기**에 걸려 있으면 TerminateProcess 가
    먹지 않고, 그 두 번째 드레인이 끝나지 않는다. 30분 제한을 걸어 둔 단계가
    **13시간 30분**을 매달려 있었다(부모 CPU 0.4초 · 자식 0.5초 — 둘 다 그냥 서 있었다).

    그리고 그동안 회차는 **락을 쥔 채**라 다음 회차가 "이미 실행 중"으로 조용히
    건너뛴다. 스케줄러는 '성공'이라 적는다. 즉 **하루치 대조가 통째로 안 돌면서
    아무 데도 티가 안 난다** — 09:50 이 하는 일(접수취소·객관완료·청구상태)이 전부 멈춘다.

    그래서: ① 나무째 죽이고 ② 드레인에도 **제한을 건다** ③ 그래도 안 죽으면
    포기하고 **다음 단계로 넘어간다**. 안 죽은 자식은 리포트에 pid 로 남긴다 —
    회차 하나를 살리자고 회차 전체를 세우지는 않는다.
    """
    p = subprocess.Popen([PY] + args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                         text=True, encoding="utf-8", errors="replace", cwd=ROOT, env=ENV)
    try:
        so, se = p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill_tree(p)
        try:
            p.communicate(timeout=60)
        except subprocess.TimeoutExpired:
            return {"name": name, "ok": False,
                    "out": f"시간초과({timeout}s) — 죽이려 해도 안 죽는다(pid {p.pid}). "
                           "Z: 대기에 걸린 것으로 보고 다음 단계로 넘어갔다"}
        return {"name": name, "ok": False, "out": f"시간초과({timeout}s)"}
    # 예외 이름과 원인은 traceback 끝에 있다. 앞 500자만 남기면 호출 위치만 보이고
    # 실제 TimeoutError/PermissionError가 잘려, 2026-08-01 큐 실패를 진단하지 못했다.
    out = (so or "") + (("\n[stderr] " + se[-2000:]) if p.returncode != 0 and se else "")
    return {"name": name, "ok": p.returncode == 0, "out": _tidy(out)}


def _run_pipeline():
    steps = []

    # 0. 합성검증 — 실패 시 전체 중단
    s = run("합성검증", [os.path.join(ROOT, "tests", "synthetic_check.py")])
    steps.append(s)
    if not s["ok"] or "ALL GREEN" not in s["out"]:
        finish(steps, aborted=True)
        sys.exit("합성검증 실패 — 전체 중단")

    # 합성검증 뒤, 모든 대조보다 먼저 단일 투입함을 정본 폴더로 분류한다.
    # 예전에는 다운로드 흡수가 파이프라인 끝이라 새 자료가 다음 날까지 반영되지 않았다.
    # ★ 보고일 자동 갱신(2026-08-06 지시) — 맨 앞에 둔다. 뒤 단계들이 이 날짜를
    #   기준으로 브리핑·캡처를 만들기 때문이다. 이미 오늘 것이면 큐를 늘리지 않는다.
    steps.append(run("보고일·집계기준일 자동 갱신",
                     [os.path.join(ROOT, "report_dates.py")], timeout=300))

    steps.append(run("업로드 투입함 원본 분류",
                     [os.path.join(ROOT, "upload_intake.py"), "--apply"]))

    # 1. 판매·세금계산서 inbox 대조 (inbox 없으면 원장 준비표)
    steps.append(run("판매·세금계산서 대조", [os.path.join(ROOT, "ecount_reconcile.py")]))

    # 2. ERP 계정별원장 4유형 대조 (inbox에 '원장' 파일 있을 때만)
    # PC가 꺼진 동안 휴대폰에서 예약한 코드를 영구 큐에서 안전하게 가져온다.
    # 반영 성공 확인 전에는 서버 항목을 지우지 않으며 입력 보호시간에는 이 단계도 멈춘다.
    steps.append(run("휴대폰 클라우드 예약 반영", [os.path.join(ROOT, "cloud_queue_sync.py")]))

    if has_inbox_kind("ledger"):
        steps.append(run("ERP원장 4유형 대조", [os.path.join(ROOT, "erp_ledger_check.py")]))
    else:
        steps.append({"name": "ERP원장 4유형 대조", "ok": None,
                      "out": "스킵 — 내용 판별 가능한 계정별원장 파일 없음"})

    # 2.5 쿠팡 PO 대조 (inbox에 'PO' 파일 있을 때만)
    if has_inbox_kind("po"):
        steps.append(run("쿠팡 PO 대조", [os.path.join(ROOT, "po_reconcile.py")]))
    else:
        steps.append({"name": "쿠팡 PO 대조", "ok": None, "out": "스킵 — inbox/에 쿠팡 PO 목록 파일 없음(파일명에 PO 포함)"})

    # 2.7 입금(수금) 자동입력 — 자료가 들어오면 사람 손 없이 채운다(사용자 지시 2026-07-28).
    #     파일이 없으면 receipt_fill 이 스스로 안내만 하고 조용히 끝나므로 조건 없이 돌린다.
    #     (계정별원장은 '0. 원본 자료' 에도 들어오므로 파일명 조건을 걸면 놓친다 — pick 이 내용으로 찾는다)
    steps.append(run("입금 대조·자동입력", [os.path.join(ROOT, "receipt_fill.py"), "--queue"]))

    # 2.8 카톡 신규 접수 등록 — 대화 내보내기가 들어오면 02·04 에 새 행으로 올린다.
    #     유형이 확정된 것(돌발·정기)만 올리고 철거·납품은 보류한다 — 대상 시트가 없다.
    kakao_sources = _source_files("kakao_dirs", (".txt",))
    if kakao_sources:
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

    # 2.955 청구(거래명세서) 근거 갱신 — ERP 판매조회의 새 매출을 06시트 빈 행에 올린다.
    #       ★ 같은 내용의 판매조회가 여러 벌 있으면 금액이 배수로 합산된다(2026-07-30 3배 사고).
    #         billing_fill 이 SHA256 으로 같은 파일을 한 번만 읽으므로 자동에 태워도 안전하다.
    steps.append(run("청구 근거 갱신(06시트)", [os.path.join(ROOT, "billing_fill.py"), "--queue"]))

    # 2.9555 청구상태(06시트 AH)를 ERP 수금확인으로 맞춘다 — 사용자 지시(2026-08-08)
    #        "ERP 기준으로 확정하고 객관적으로 입증되면 엑셀에 완료처리해".
    #        ★ 회차로 만드는 이유: 수금은 **매일 새로 생긴다.** 한 번 손으로 맞춰 두면
    #          그날 이후로는 다시 낡는데, 낡은 단계 딱지는 비어 있지 않아서
    #          **아무 화면에도 티가 안 난다** — 이 프로젝트가 계속 당해 온 종류다.
    steps.append(run("청구상태 ERP 수금확인 반영",
                     [os.path.join(ROOT, "billing_status.py"), "--queue"]))

    # 2.96 재계산 대기 세기 — 원장엔 올라왔는데 엑셀이 아직 계산 안 해 앱에 안 나오는 건.
    #      숫자가 틀린 게 아니라 대기 중이라는 걸 앱이 스스로 말하게 한다(사용자 오해 방지).
    steps.append(run("재계산 대기 확인", [os.path.join(ROOT, "recalc_pending.py")]))

    # 2.98 엑셀 재계산은 11:00·15:00 일괄반영 회차 안에서만 수행한다.
    #      09:50에는 승인·대기 상태만 읽어 보고, Excel COM을 열거나 저장하지 않는다.
    steps.append(run("엑셀 재계산 상태 확인", [os.path.join(ROOT, "excel_recalc.py")]))

    # 2.97 ERP 접속 IP 확인 — 공인 IP가 바뀌면 이카운트 OAPI가 통째로 막힌다.
    #      사용자 지시(2026-07-30): "IP가 변경되면 이 화면에서 등록해서 진행".
    #      자동 등록은 하지 않는다(회사 ERP 보안 설정) — 바뀐 사실과 넣을 값을 알린다.
    steps.append(run("ERP 접속 IP 확인", [os.path.join(ROOT, "erp_ip_guard.py")]))

    # 3. 밴드 수집·대조 — 공식 API 토큰이 있으면 수집+대조, 브라우저 수집 캐시만 있으면 대조만
    band_dumps = _source_files("band_dump_dirs", (".json",))
    if band_dumps:
        steps.append(run("밴드 업로드 원본 변환", [os.path.join(ROOT, "band", "ingest.py")]))

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
    if kakao_sources:
        steps.append(run("카톡 대조", [os.path.join(ROOT, "kakao", "kakao_reconcile.py")]))
    else:
        steps.append({"name": "카톡 대조", "ok": None, "out": "스킵 — kakao/inbox/에 대화 내보내기 txt 없음"})
    # 카톡 내보내기가 1일 넘게 오래되면 대시보드·아침 브리핑에 경고(2026-08-04 지시).
    steps.append(run("카톡 내보내기 경과 감시", [os.path.join(ROOT, "kakao", "export_watch.py")]))

    # 5. 확정 업데이트 자동 반영 — 빈 칸만·근거 보유·항상 새 버전(vN+1) 생성이라 안전
    # ERP 매출서류(계산서·명세서 현황) — 있으면 대조 + 25시트 반영
    _has_tax = has_inbox_kind("tax")
    if _has_tax:
        steps.append(run("ERP 매출서류 대조(반영 대기)", [os.path.join(ROOT, "erp_docs_check.py")]))
    else:
        steps.append({"name": "ERP 매출서류 대조", "ok": None,
                      "out": "스킵 — inbox/에 매출(세금)계산서현황 없음"})

    # 밴드 문서 사진(거래명세서·세금계산서) OCR → 확실한 건은 빈칸 입력 큐에 적재
    document_images = _source_files("doc_photo_dirs",
                                    (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".heic"))
    if document_images:
        # 사진 1장에 OCR 1초 남짓. 첫 회는 1,458장에 25분이 걸려 600초 제한에 걸렸다(2026-07-26).
        # doc_ocr가 결과를 band/ocr_cache/에 저장하므로 두 번째 실행부터는 몇 초로 끝난다.
        steps.append(run("밴드 문서 이미지 대조·입력", [os.path.join(ROOT, "band", "doc_ocr.py"),
                                                       "--scan", "--apply"], timeout=3600))
        # 그 위에 교차검증을 한 겹 더 얹는다(2026-08-06 지시). 엔진 하나의 답만 믿으면
        # 틀린 값이 조용히 원장에 들어간다 — 서로 다른 엔진 둘이 **같은 값**을 낸 항목만
        # 빈칸에 넣고, 갈리면 사람에게 넘긴다. 전량을 두 번 읽지는 않는다:
        # 금액 정합성이 깨졌거나 핵심 항목을 못 읽은 건, 그리고 원장에 쓰려는 건만 재검한다.
        steps.append(run("문서 OCR 교차검증", [os.path.join(ROOT, "band", "ocr_crosscheck.py"),
                                               "--scan", "--apply"], timeout=3600))
    else:
        steps.append({"name": "밴드 문서 이미지 대조", "ok": None,
                      "out": "스킵 — band/docs_inbox/에 사진 없음"})
        steps.append({"name": "문서 OCR 교차검증", "ok": None, "out": "스킵 — 사진 없음"})

    # 미수집 원본·사진·텍스트를 **매일 조금씩** 굳힌다 (2026-08-08 지시:
    # "미수집 데이터들 싹 다 긁어모아 … 두번 일 안하게"). 한 번에 다 하려 들면
    # 회차가 시간 제한에 걸려 통째로 실패한다 — 상한을 두고 매일 이어 간다.
    # 이미 있는 것은 파일을 열지도 않으므로 다 끝난 뒤에는 몇 초로 끝난다.
    # 로그인이 필요한 수집(밴드·ERP)은 여기에 없다 — 그건 사람 몫이다(절대규칙 3).
    steps.append(run("미수집 원본·사진·텍스트 보관", [os.path.join(ROOT, "collect_all.py"),
                                                     "--run", "--limit", "600"], timeout=3600))

    # 완료보고서와 문서발행 표시는 프로젝트NO가 정확히 일치하는 근거만 빈칸에 큐잉한다.
    # 실제 ZIP 패치는 바로 아래 ledger_writer 한 번으로 합쳐 버전 난립과 충돌을 막는다.
    steps.append(run("완료보고서 확정분 큐", [os.path.join(ROOT, "confirm_fill.py"), "--queue"]))
    # 밴드·카톡 완료보고 + ERP 판매조회 + 06 거래명세서 원장을 프로젝트NO로 함께 대조한다.
    # 02·03·04의 원인 열만 채우므로 검증결과 수식은 보존되고 Excel에서 정상/확인으로 재계산된다.
    steps.append(run("ERP·거래명세서·현장검증 확정분 큐",
                     [os.path.join(ROOT, "verification_sync.py"), "--queue"]))
    # 정기점검·돌발AS 일지(미실시건) — 완료근거는 원장 빈 칸만 보완 큐에 넣고,
    # 미실시 사유·취소·일자상이는 28_대조현황과 대표보고에 그대로 남긴다.
    steps.append(run("정기점검·돌발AS 일지 대조 안전입력 큐",
                     [os.path.join(ROOT, "work_log_sync.py"), "--queue"]))
    # 완료일·검증 정상 등 객관 요건을 충족한 행을 Excel 상태 셀 대신 DB 완료 정본에 기록한다.
    # confirm_fill의 밴드 완료글+날짜 판정과 함께 매일 돌아야 사람이 다시 지시하지 않는다.

    # ★ 2026-07-30 지시: 반영은 **DB에 모았다가 하루 두 번(11:00·15:00)만** 엑셀에 쓴다.
    #   전에는 여기서 바로 --apply 해서 하루에 관리대장 버전이 수십 개씩 늘었다(v311→v327).
    #   09:50 일일 대조는 적재까지만 하고, 별도 작업 스케줄러가 두 회차를 정확히 실행한다.
    # ★ ERP 등록 여부(02_돌발AS접수 ERP등록 · 04_정기점검 ERP판매전표)를 채운다.
    #   2026-08-07 에 알았다: 이 도구는 2026-07-28 에 만들어져 있었는데 **daily_run 에
    #   들어 있지 않았다.** 손으로 돌린 사람이 있을 때만 채워졌다는 뜻이다. 이날 그냥
    #   돌려 보니 채울 칸이 52개 남아 있었고, 그중 46건은 ERP 가 이미 완료를 입증한
    #   건이었다. 도구가 있는데 아무도 안 부르는 것이 가장 조용한 종류의 누락이다.
    #   판정 근거는 ERP **판매조회의 프로젝트코드(UJ번호)** 다 — 거래명세서번호가
    #   아니다(그 길은 2026-07-28 에 이미 막다른 길로 판명났다. 365건이 그렇게 남았었다).
    #   --queue 라 엑셀은 열지 않는다. 반영은 11:00·15:00 회차 그대로.
    #   (retry 는 적지 않는다 — `--queue` 라 NO_RETRY_FLAGS 가 알아서 막는다)
    steps.append(run("ERP 등록여부 판정(판매조회)",
                     [os.path.join(ROOT, "fill_erp_status.py"), "--queue"]))
    steps.append(run("입력 DB 적재", [os.path.join(ROOT, "ledger_db.py"), "--intake"]))
    # 방금 흡수한 신규행에도 실제 완료일+원천 근거가 있으면 Excel 회차를 기다리지 않고
    # SQLite 완료 정본에 즉시 기록한다. Excel 자체는 정해진 11:00·15:00 회차만 유지한다.
    steps.append(run("객관근거 완료 DB 동기화",
                     [os.path.join(ROOT, "complete_verified.py"), "--queue"]))
    # 금액은 PO 원본 견적서와 거래명세서가 프로젝트·PO·총액까지 유일 일치할 때만,
    # 계산서는 25/26시트 ERP 원본의 '확정' 프로젝트만 정산 완료 DB에 올린다.
    steps.append(run("정산 객관완료 DB 동기화",
                     [os.path.join(ROOT, "settlement_completion.py"), "--sync"]))
    # 세 사람의 담당 범위도 같은 객관 근거로 완료 이력을 남긴다. 상태 문구는
    # "류지영 완료"·"오종현 완료"·"유현민 완료"이며 사람 체크나 추정은 쓰지 않는다.
    steps.append(run("담당자별 객관완료 DB 동기화",
                     [os.path.join(ROOT, "staff_completion.py"), "--sync"]))
    # 오래된 세금계산서 미발행을 경과일 순으로 잡아낸다(2026-08-03 지시). 읽기 전용 감시 —
    # 발행일은 지어내지 않고, 발행 사실은 ERP 원본이 들어와야 객관완료가 잡는다.
    # ERP 거래명세서 **인쇄본**(ESD007E) 파싱·원장 대조. 현황표(erp_docs_check)와 레이아웃이
    # 달라 그 도구는 0건으로 봤다(2026-08-04 실측) — 별도 파서가 읽는다. 읽기 전용.
    steps.append(run("거래명세서 인쇄본 대조", [os.path.join(ROOT, "stmt_docs.py")],
                     timeout=1500))
    # ERP 판매조회 → 프로젝트 색인. **완료 판정·앱 금액의 정본**이라 명세서 색인보다 먼저.
    #   (2026-08-05: 이 색인이 없으면 settlement_completion 의 ERP 근거가 통째로 죽는다)
    steps.append(run("ERP 판매 프로젝트 색인", [os.path.join(ROOT, "erp_sales_index.py")],
                     timeout=900))
    # 명세서 ↔ 판매조회(UJ 프로젝트코드) 색인. 인쇄본에는 UJ 가 없어 금액·거래처·품목으로
    # 잇는다(2026-08-05). 이 색인이 건별 PDF 파일명의 프로젝트NO 가 된다.
    steps.append(run("명세서 ↔ 프로젝트 색인", [os.path.join(ROOT, "stmt_link.py")],
                     timeout=1500))
    # 거래처코드·캠프 마스터 — 앱이 CU코드를 표시하고, 캠프 공백(ERP에만/원장에만)을 드러낸다.
    steps.append(run("거래처코드 색인", [os.path.join(ROOT, "customer_index.py")], timeout=900))
    steps.append(run("캠프 마스터", [os.path.join(ROOT, "camp_master.py")], timeout=1200))
    # ★ 사용자 지시(2026-08-05) "각 건의 PDF·이미지를 번호로 알아보게 저장":
    #   밴드는 글 단위(PDF+텍스트+사진), ERP 명세서는 전표번호 단위 PDF 로 굳힌다.
    #   회차마다 상한을 둬 daily_run 이 길어지지 않게 한다 — 남은 건 다음 회차가 잇는다.
    steps.append(run("밴드 게시글 보관(PDF·텍스트·사진)",
                     [os.path.join(ROOT, "band", "archive_posts.py"), "--limit", "150"],
                     timeout=2400))
    steps.append(run("명세서 건별 PDF 보관",
                     [os.path.join(ROOT, "stmt_archive.py"), "--limit", "150"],
                     timeout=2400))
    steps.append(run("세금계산서 건별 PDF 보관",
                     [os.path.join(ROOT, "tax_archive.py"), "--limit", "150"],
                     timeout=2400))
    # ★ 원본이 늘면 분석 3종(미발행·불일치·확인필요현황)이 10분을 넘긴다 —
    #   2026-08-04 거래명세서 785건 흡수 후 기본 600초에 셋 다 타임아웃으로 FAIL했다.
    #   단독 재실행은 전부 성공(로직 문제 아님) → 시간만 넉넉히 준다.
    steps.append(run("세금계산서 미발행 경과 감시",
                     [os.path.join(ROOT, "tax_invoice_watch.py")], timeout=1500))
    # 금액 재계산 대기의 견적↔명세 교차·불일치 진단(2026-08-03 지시). 읽기 전용 —
    # 명세합계가 다른 프로젝트 견적과 일치하는 입력 밀림을 짝까지 찾아 보여 준다.
    steps.append(run("견적·명세 불일치 진단",
                     [os.path.join(ROOT, "quote_mismatch.py")], timeout=1500))
    # 09:50에는 읽기 전용 무결성 검사만 한다. 실제 복구도 11:00·15:00 회차 안에서만 한다.
    steps.append(run("워크북 무결성 검사", [os.path.join(ROOT, "fix_workbook.py")]))

    # 5.3 류지영 정기점검 스케줄 원본 → 27_정기점검원본일정.
    #     지정 폴더의 최신 xlsx를 매일 다시 읽으며, 내용이 같으면 새 버전을 만들지 않는다.
    #     원본에 UJ번호가 없으므로 임의 프로젝트를 만들지 않고 04시트와 근거가 있는 건만 연결한다.
    steps.append(run("정기점검 스케줄 원본 분석",
                     [os.path.join(ROOT, "pm_schedule_sync.py")]))

    # 신규 프로젝트 업무 흐름도는 관리대장·앱에 노출하지 않는다. 지정 원본 폴더의 최신본만
    # 내부 DB로 안전하게 교체해, 이후 신규 업무 기준을 추가해도 원본 구조를 그대로 보관한다.
    steps.append(run("신규 프로젝트 업무 흐름도 DB 동기화(앱 비표시)",
                     [os.path.join(ROOT, "new_project_flow_sync.py"), "--apply"]))

    # 5.5 밴드 업무 추출 보고. 24시트 반영은 11:00·15:00 회차에서만 한다.
    if band_cache:
        steps.append(run("밴드 업무추출(반영 대기)", [os.path.join(ROOT, "band_extract.py")]))
        # 5.6 접수했다가 취소된 건(2026-08-08 지시). 취소된 건이 원장에 접수 그대로
        #     남으면 'AS 미실시'가 영영 안 줄어든다 — 화면은 멀쩡한데 숫자가 거짓인
        #     조용한 사고다. 엑셀은 안 연다: 대기열에만 넣고 11:00·15:00 회차가 반영한다.
        steps.append(run("접수 취소 확인(반영 대기)",
                         [os.path.join(ROOT, "cancel_watch.py"), "--queue"], timeout=1500))
        # 5.7 위 단계가 **반쪽으로 도는 것**을 막는다 (2026-08-09). 취소는 대부분 댓글로
        #     오는데 실측 7,475글이 댓글을 한 번도 안 들여다봤다. 그 상태에서도
        #     cancel_watch 는 오류 없이 끝난다 — 사각지대는 오류가 아니라 **없는 자료**다.
        #     여기서는 **긁지 않는다**(사람 로그인이 필요하다 — 절대규칙 3).
        #     사람 손에 갈 붙여넣기 파일만 최근 것부터 다시 만들어 둔다.
        steps.append(run("댓글 사각지대 붙여넣기 파일 갱신",
                         [os.path.join(ROOT, "band", "comment_backfill.py"),
                          "--days", "90", "--write"], timeout=900))

    # 6. 확인필요현황은 별도 보고서만 갱신한다. 관리대장 23시트 쓰기는 11:00·15:00 회차.

    # 6.5 확인필요현황 집계 갱신.
    #     ★ 2026-08-07 지시로 **별도 엑셀을 더는 만들지 않는다** ("앞으로도 별도의
    #       엑셀 파일은 만들지 말고 관리대장으로만 관리해"). 예전 주석은 *"23시트는
    #       평면 목록이라 유형별 상세열을 담지 못한다"* 였는데, 담지 못한 게 아니라
    #       안 나눴을 뿐이었다 — 이제 프로젝트NO·명세서번호·PO번호·확인방법이
    #       시트에 **열로** 있다. 시트 반영은 11:00·15:00 회차(`ledger_db`)가 한다.
    #     이 단계가 여전히 매일 도는 이유는 **집계 JSON** 때문이다
    #     (reports/확인필요_집계.json — 앱·리포트가 이 숫자를 읽는다).
    #     ★ 여기에 연결돼 있지 않아 2026-07-27 16:45 자로 하루 동안 멈춰 있었다 —
    #       읽는 사람은 멈춘 줄 모르고 어제 숫자를 오늘 숫자로 본다. 반드시 매일 돈다.
    steps.append(run("확인필요현황 집계 갱신", [os.path.join(ROOT, "findings_export.py")],
                     timeout=1500))

    # 6.6 Z: 상시 공백 스캔 — 파일명 기반 전수조사와 서류↔원장 1:1 대조는 읽기 전용이다.
    #     2026-07-28 뒤 리포트가 멈춰 자료현황이 오래된 숫자를 계속 보여 준 문제를 막는다.
    steps.append(run("Z폴더 원장 누락·금액 공백 스캔", [os.path.join(ROOT, "zscan.py")], timeout=1800))
    steps.append(run("Z폴더 서류 1:1 대조", [os.path.join(ROOT, "zscan.py"), "--docs"], timeout=1800))

    # 6.7 자료현황 한 장 — "밴드에서 뭘 얼마나 가져왔나 / 지금 뭘 갖고 있나 / 원장이 얼마나 찼나".
    #     같은 질문을 매번 다시 세지 않으려고 만든다(사용자 지시 2026-07-29).
    #     느린 것(Z: 2만 개 순회)은 다시 돌지 않고 앞 단계가 남긴 리포트에서 숫자만 읽는다 —
    #     그래서 이 단계는 위 대조들이 **끝난 뒤에** 와야 한다.
    # ★ 원본 색인·폴더 정리(2026-08-05 지시 "누가 봐도 깔끔하게 항상 정리").
    #   색인이 있어야 앱 '원본 자료' 화면과 바로가기가 최신을 가리킨다.
    # 업무센터 UX 점검은 3일마다 12:00 별도 스케줄러가 돈다(쿠팡업무_업무센터UX점검).
    #   daily_run 에서도 가볍게 한 번 더 갱신해 리포트가 오래되지 않게 한다.
    steps.append(run("업무센터 UX 점검", [os.path.join(ROOT, "ux_review.py")], timeout=600))
    steps.append(run("원본 색인 갱신", [os.path.join(ROOT, "source_index.py")], timeout=2400))
    steps.append(run("원본 폴더 정리·바로가기", [os.path.join(ROOT, "source_tidy.py")], timeout=1800))
    # 캐시가 어긋나면 Z: 원본 전체를 다시 훑는다 — 실측 콜드 222초, 경합 아래서는 600초를
    # 넘겨 매 회차 FAIL 이었다(2026-08-07). `_보관` 제외로 줄였지만 여유를 둔다.
    steps.append(run("자료현황 갱신", [os.path.join(ROOT, "data_status.py")], timeout=1200))
    # 미반영 목록(2026-08-05 지시 "반영 안된 자료들 목록 정리"). 앞 단계들이 남긴 리포트를
    # 모아 세므로 **맨 뒤**여야 한다 — 먼저 돌면 어제 숫자를 오늘 것으로 보여 준다.
    steps.append(run("미반영 목록 갱신", [os.path.join(ROOT, "pending_report.py")], timeout=300))
    # 6.75 하루치 정산분 보고자료 — 다음 날 아침 대표 보고에 그대로 쓴다(2026-08-05 지시).
    #      카톡·밴드·ERP 를 각각 그 날짜로만 훑어 만든다. 앞 단계(색인·밴드 반영)가
    #      끝난 뒤여야 ERP 내보내기와 밴드 글이 최신으로 잡힌다.
    steps.append(run("정산분 보고자료(어제)", [os.path.join(ROOT, "settle_report.py")],
                     timeout=900))

    # 6.85 다운로드 흡수 — 실행 도중 새로 내려받은 파일을 다음 회차용 투입함에 보존한다.
    #      떨어진다. 손으로 옮기면 세션이 바뀌는 순간 잊힌다(2026-07-31 실제로 그랬다).
    #      내용 판별로 '0. 원본 자료'에 이동 — Z: 라서 옮겨지는 순간 모든 PC·세션이 본다.
    steps.append(run("다운로드 흡수(원본 자료로)", [os.path.join(ROOT, "download_intake.py"), "--apply"]))

    # 6.87 ERP 엑셀 → PDF 사본 (사용자 지시 2026-08-04 "PDF 또는 이미지로 저장해 전부 반영").
    #      엑셀은 열 때마다 서식이 흔들리고 재조회 시점에 숫자가 달라진다. PDF 는 "그때
    #      ERP 가 이렇게 보여 줬다"를 고정한다 — 사후 대조·감사에 쓰는 건 이 고정본이다.
    #      새 파일만 변환하므로 매 회차 비용이 거의 없다(대상 0개면 Excel 을 띄우지도 않는다).
    # ★ 예산이 설계값과 어긋나 있었다 (2026-08-07 실측). 이 도구는 안에서 Excel 서브프로세스에
    #   1800초를 주는데(erp_pdf_export.py:116) 여기서는 600초만 줬다. 그래서 변환할 게
    #   4개뿐인 날에도 `시간초과(600s)` 로 매 회차 FAIL 이었다(단독 실행은 315초 exit 0).
    steps.append(run("ERP PDF 사본 만들기", [os.path.join(ROOT, "erp_pdf_export.py")],
                     timeout=1800))

    # 6.9 폰 원격 준비 상태 — 막히는 건 조용히 막힌다. 절전 설정이 되살아나거나 미푸시가
    #     쌓이면 정작 폰에서 붙으려 할 때 알게 된다(사용자 지시 2026-07-31). 아무것도 바꾸지 않고
    #     상태만 남긴다 — 전원·SSH 는 관리자 권한이 필요한 시스템 설정이라 사람이 실행한다.
    steps.append(run("폰 원격 준비 점검", [os.path.join(ROOT, "remote_ready.py")]))

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
    # 공유 캘린더 꾸러미를 **먼저** 새로 잠근다 — 그래야 바로 아래 게시가 최신을 올린다.
    # (단톡방에 뿌린 링크는 이 파일 하나만 본다. 2026-08-06 지시)
    steps.append(run("공유 캘린더 갱신", [os.path.join(ROOT, "cal_share.py")], timeout=600))
    steps.append(run("고정 주소 사본 올리기", [os.path.join(ROOT, "cloud_publish.py"), "--push"]))

    # 11. 버전 파일 정리 — 최신본 하나만 작업 폴더에 두고 구버전은 사용자가 지정한
    #     OLD/ 한 곳으로 옮긴다. 같은 이름이 있어도 덮어쓰거나 삭제하지 않는다.
    # 9.9 복구용 보관 — 코드는 git bundle 한 파일로, 기록은 사실만. 서버(Z:)에 둔다.
    #     PC가 죽으면 PC 안의 백업은 같이 죽는다. 비밀키는 절대 담지 않는다(규칙 1).
    # ★ 재시도하지 않는다 (2026-08-07 실측). 이 단계는 Z: 로 1,721개·199.6MB 를 한 개씩
    #   복사하며 완주에 ~1,475초가 든다 — 산발적 경합이 아니라 **결정적으로** 한도를 넘는다.
    #   그런 단계를 재시도하면 실패에 시간을 두 배로 태울 뿐이다(세 단계 합쳐 회차당 ~81분).
    #   근본 해결(증분 복사·야간 이동)은 별도 과제로 남긴다 — reports/오류점검_20260807.md B1.
    steps.append(run("복구용 보관(서버)", [os.path.join(ROOT, "archive_keep.py")],
                     timeout=1800, retry=0))

    steps.append(run("관리대장 버전 정리", [os.path.join(ROOT, "ledger_versions.py"), "--prune"]))

    finish(steps)


def main():
    if is_input_window():
        print(f"입력 보호시간({input_window_label()}) — 일일 자동대조를 시작하지 않습니다.")
        return
    token = acquire_run_lock()
    if not token:
        print("다른 daily_run 프로세스가 이미 실행 중 — 중복 실행을 시작하지 않습니다.")
        return
    try:
        return _run_pipeline()
    finally:
        release_run_lock(token)


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
