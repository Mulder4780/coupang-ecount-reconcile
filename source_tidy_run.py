#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""원본 자료 정리 회차 — **어느 단계에서 멈췄는지 반드시 남긴다** (2026-08-19 실사고).

★ 실측: `쿠팡업무_원본자료자동정리` 가 09:35 에 떠서 **12:35 에 제한시간(PT3H)으로
  강제 종료**됐다. 그런데 `reports/source_organizer.log` 에는 **시각도 단계 표시도
  없이** 파일 이름만 쏟다가 11:37 에 멎어 있었다 — 즉 **58분을 아무 말 없이 서
  있었고**, 네 단계 중 어느 것이었는지 알 길이 없었다. 스케줄러가 아는 것은
  '끝내기 요청을 받고 멈췄다' 한 줄뿐이다(`[228]` — exit 코드는 왜인지를 말해 주지
  않는다). 그래서 같은 일이 며칠 반복돼도 아무도 못 고친다.

★ **짐작으로 제한시간부터 늘리지 않는다**(분담판 `[38]`). 먼저 **범인 단계를 대게**
  만드는 것이 순서다 — 그것이 이 파일이 하는 전부다.

★ **단계마다 제한을 건다.** 네 단계가 이어 달리는데 제한이 회차 전체(PT3H)에만 있으면
  **한 단계가 세 시간을 다 먹고 나머지 셋은 아예 안 돈다** — 그런데 그 셋이 안 돌았다는
  사실은 어느 화면에도 안 뜬다(`[169]`). 4×40분=160분이라 PT3H 안에 반드시 끝나고,
  끝나야 자국이 써지고 다음 회차가 돈다(`[180]`).

★ **한 단계가 죽어도 다음으로 간다**(`[175]`) — 단계 하나를 살리자고 회차 전체를
  세우지 않는다. 다만 **죽었다는 사실은 적는다**(조용히 넘어가면 '돌았는데 왜 결과가
  없나'가 된다).

★ `subprocess.run(timeout=)` 을 쓰지 않는다(`[175]`) — 윈도우에서 SMB(Z:) 대기에
  걸리면 `kill()` 뒤 무제한 `communicate()` 에 매달린다. `proc_guard.run_tree` 가
  나무째 끊고 드레인에도 제한을 건다.

★ **성공하면 자국을 지운다**(`[228]`) — 옛 자국이 남으면 이미 고쳐진 고장을 계속
  보고한다.

★ **갈래는 새로 만들지 않는다**(`[162]`) — `autopilot.classify_failure` 가 이미 가른다.
"""

import json
import os
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(ROOT, "reports")
LOG = os.path.join(REPORT_DIR, "source_organizer.log")
CRASH = os.path.join(REPORT_DIR, "원본정리_오류.json")
PROGRESS = os.path.join(REPORT_DIR, ".원본정리_진행.json")

# 단계 하나가 회차 전체를 먹지 못하게. 4×40분 = 160분 < PT3H(180분).
# ⚠ 이 값을 올릴 때는 **STEPS 개수 × 이 값 < 작업 제한시간** 을 지킨다 — 안 지키면
#   마지막 단계가 잘리면서 자국도 안 써진다(지금 고치려는 그 모양으로 되돌아간다).
STEP_TIMEOUT_S = int(os.environ.get("COUPANG_TIDY_STEP_S") or 2400)

# ★ 자식에게는 **더 짧은 예산**을 준다 — 그래야 스스로 멈추며 "이동 N개째 ·
#   훑은 파일 M개" 를 로그에 남긴다. 밖에서 죽이면(SIGKILL) 그 자국이 통째로
#   사라진다 — 실측 2026-08-22: `원본 폴더 정리` 가 313분을 먹고 죽었는데
#   로그에 한 줄도 안 남아 **무엇을 하다 죽었는지 알 길이 없었다**(`[169]`).
# ★ 그 자식의 기본 예산은 7200초(120분)라 단계 제한 2400초(40분)보다 길다 —
#   즉 **한 번도 안 쓰이는 규칙**이었다. 짧게 줘야 뜻이 생긴다.
# ⚠ 예산을 읽는 스크립트는 지금 `source_organizer` **하나뿐**이다. 안 읽는
#   자식에게 넣어 봐야 아무 일도 안 일어난다 — 없는 손잡이를 지어내지 않는다.
#   다른 자식이 예산을 읽게 되면 여기 표에 한 줄 더한다.
STEP_BUDGET_MARGIN_S = int(os.environ.get("COUPANG_TIDY_MARGIN_S") or 300)
CHILD_BUDGET_ENV = {"source_organizer.py": "SOURCE_ORGANIZER_BUDGET_SEC",
                    "collect_sources.py": "SOURCE_COLLECT_BUDGET_SEC"}

#: 자식이 **예산이 다 돼 스스로 멈췄다**는 뜻 — 실패가 아니다([427]).
#  값은 `child_budget` 한 곳에서 온다([162]).
import child_budget                                     # noqa: E402
INCREMENTAL_RETURN_CODE = child_budget.INCREMENTAL_RETURN_CODE


def _child_env(스크립트):
    """그 자식이 예산을 읽으면 단계 제한보다 짧게 줘서 스스로 멈추게 한다."""
    키 = CHILD_BUDGET_ENV.get(os.path.basename(스크립트))
    if not 키:
        return None                     # 안 읽는 자식은 건드리지 않는다
    env = dict(os.environ)
    env[키] = str(max(60, STEP_TIMEOUT_S - STEP_BUDGET_MARGIN_S))
    return env

STEPS = [
    ("업로드함 흡수", "upload_intake.py"),
    ("원본 모으기", "collect_sources.py"),
    ("원본 폴더 정리", "source_organizer.py"),
    ("신규 프로젝트 흐름", "new_project_flow_sync.py"),
]

_FIX = {
    "resource": "Z:(SMB)·파일을 못 읽었다 — 연결과 원본 폴더부터 본다. 코드가 아니다.",
    "auth": "로그인·인증이 풀렸다 — 사람이 한 번 로그인해야 이어진다.",
    "timeout": "그 단계가 제한시간을 넘겼다 — 무엇이 오래 걸리는지 그 단계 안을 본다"
               "(제한시간을 먼저 늘리지 말 것 · 분담판 [38]).",
    "code": "코드가 깨졌다 — 아래 자취를 그대로 읽는다.",
}


def _say(msg):
    """`pythonw.exe` 로 돌면 `sys.stdout` 이 **None** 이다(`[235]`)."""
    if getattr(sys, "stdout", None) is not None:
        try:
            print(msg)
        except Exception:
            pass


def _log(msg):
    """로그에 **시각과 함께** 적는다 — 이것이 없어서 58분의 침묵을 못 읽었다."""
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        with open(LOG, "a", encoding="utf-8", errors="replace", newline="") as fh:
            fh.write("[%s] %s\n" % (datetime.now().isoformat(timespec="seconds"), msg))
    except Exception:
        pass                      # 로그를 남기려다 회차를 세우지 않는다


def _note(state):
    """**죽어도 남는 자국.** 지금 어느 단계인지·얼마나 서 있는지를 계속 갱신한다."""
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        with open(PROGRESS, "w", encoding="utf-8") as fh:
            json.dump(state, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass


def _kind(text):
    try:
        import autopilot
        return autopilot.classify_failure(text or "") or ""
    except Exception:             # noqa: BLE001
        return ""                 # 모르면 '모름' — 지어내지 않는다(`[169]`)


def _worst(rows):
    """가장 오래 걸린 것을 고른다 — 회차를 실제로 잡아먹은 단계다.

    ★ 예전에는 `fails[0]`, 곧 **순서**로 골랐다. 실측 2026-08-22 회차에서
      40분 먹은 `원본 모으기` 가 범인이 되고 **313분(5.2시간) 먹은
      `원본 폴더 정리`** 는 '그 밖 1단계'로 묻혔다. 인계도 콘솔도 40분짜리를
      지목하니 사람은 **엉뚱한 단계를 고치러 간다**(`[172]`) — 실제로 그랬다.
    ★ 같은 회차를 놓고 `_note` 의 `가장오래`(시간 순)와 자국·콘솔(순서)이
      **서로 다른 답**을 하고 있었다. 판정은 한 곳이다(`[162]`).
    ⚠ 시간을 못 읽는 칸은 0 으로 본다 — 못 읽은 것을 '제일 오래'로 올리면
      그것이 또 하나의 틀린 지목이다(`[169]`).
    """
    if not rows:
        return None
    return sorted(rows, key=lambda d: -(d.get("분") or 0))[0]


def _leave_trace(fails, done):
    """왜 죽었는지 남긴다. `schedule_watch.traces()` 가 `*_오류.json` 을 글로브로 모은다."""
    첫 = _worst(fails) or fails[0]
    # ★ **아는 것을 짐작으로 덮지 않는다.** 제한시간에 끊긴 것은 `timed_out` 이 이미
    #   말해 준다 — 그런데 자취 글자를 분류기에 물으면 엉뚱한 갈래가 나온다(실측:
    #   시간초과인데 `code` 로 나와 조치가 *"코드가 깨졌다"* 였다). 조치는 갈래마다
    #   다르므로(`[289]`) 그 한 줄이 사람을 **멀쩡한 코드**로 보낸다(`[172]`).
    kind = "timeout" if 첫.get("시간초과") else _kind(첫.get("자취"))
    무엇 = "%s 단계가 %s (%.0f분)" % (첫["단계"], 첫["왜"], 첫["분"])
    if len(fails) > 1:
        # ★ 나머지를 숫자로만 세지 않는다 — 몇 분씩 먹었는지가 다음에 볼 자리다([169]).
        나머지 = [d for d in fails if d is not 첫]
        무엇 += " · 그 밖 %d단계도 실패(%s)" % (
            len(나머지),
            " · ".join("%s %.0f분" % (d.get("단계"), d.get("분") or 0)
                       for d in 나머지[:3]))
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        with open(CRASH, "w", encoding="utf-8") as fh:
            json.dump({"시각": datetime.now().isoformat(timespec="seconds"),
                       "명령": "source_tidy_run.py (원본 자료 정리 회차)",
                       "갈래": kind or "모름",
                       "무엇": 무엇,
                       "조치": _FIX.get(kind, "갈래를 못 가렸다 — 아래 자취를 그대로 읽는다."),
                       "단계별": done,
                       "자취": (첫.get("자취") or "")[-4000:]}, fh,
                      ensure_ascii=False, indent=1)
    except Exception:
        pass                      # 자국을 남기려다 종료를 막지 않는다


def _clear_trace():
    """다 됐으면 옛 자국을 지운다 — 안 지우면 고쳐진 고장을 계속 보고한다(`[228]`)."""
    try:
        if os.path.exists(CRASH):
            os.remove(CRASH)
    except Exception:
        pass


def main():
    import proc_guard

    시작 = time.time()
    done, fails = [], []
    for i, (이름, 스크립트) in enumerate(STEPS, 1):
        t0 = time.time()
        _note({"시각": datetime.now().isoformat(timespec="seconds"),
               "지금단계": 이름, "번호": "%d/%d" % (i, len(STEPS)),
               "상태": "도는중", "끝낸단계": [d["단계"] for d in done],
               "회차분": round((time.time() - 시작) / 60.0, 1)})
        _log("▶ %d/%d %s (%s) 시작" % (i, len(STEPS), 이름, 스크립트))
        r = proc_guard.run_tree([sys.executable, os.path.join(ROOT, 스크립트), "--apply"],
                                cwd=ROOT, timeout=STEP_TIMEOUT_S,
                                env=_child_env(스크립트))
        분 = (time.time() - t0) / 60.0
        이어감 = False
        out = (r.stdout or "") + (("\n" + r.stderr) if r.stderr else "")
        # ★ 출력을 버리지 않는다 — 창을 없앤 회차는 파일에만 남는다(`[248]`·`[289]`).
        if out.strip():
            _log("   ─ %s 출력 ─\n%s" % (이름, out.rstrip()))
        if r.timed_out:
            왜 = "제한시간(%d분)을 넘겨 끊겼다" % (STEP_TIMEOUT_S // 60)
            # ★ 죽였는데도 안 끝난 pid 를 버리지 않는다(`[175]` 가 시킨 그것).
            #   실측 2026-08-22: 제한이 40분인데 이 단계가 **313분**을 먹었다.
            #   `run_tree` 는 40분+50초에 반드시 돌아오게 돼 있으므로 그 273분이
            #   어디로 갔는지가 다음에 물어야 할 것인데, 그 근거를 여기서 버리고
            #   있었다 — 그러면 다음에도 "왜 313분인지 모른다"로 끝난다(`[169]`).
            if getattr(r, "stuck_pid", 0):
                왜 += " · 죽인 뒤에도 안 끝난 pid %s" % r.stuck_pid
        elif r.returncode == INCREMENTAL_RETURN_CODE:
            # ★ **실패가 아니다** — 예산이 다 돼 스스로 멈췄고 진도가 남는다
            #   ([427]).  실패로 세면 매일 가짜 경보가 되어 진짜를 덮는다([170]).
            # ★ 그렇다고 **조용히 완료라 적지도 않는다**([450]) — 로그와
            #   진행 자국이 '이어감'이라고 말한다.
            왜 = ""
            이어감 = True
        elif r.returncode != 0:
            왜 = "0 이 아닌 값으로 끝났다(코드 %s)" % r.returncode
        else:
            왜 = ""
        _log("◀ %d/%d %s %s · %.1f분"
             % (i, len(STEPS), 이름,
                왜 or ("예산이 다 돼 이어간다(다음 회차가 잇는다)" if 이어감
                       else "완료"), 분))
        칸 = {"단계": 이름, "분": round(분, 1), "왜": 왜, "코드": r.returncode,
             "이어감": bool(이어감),
             "시간초과": bool(r.timed_out),
             "안죽은pid": getattr(r, "stuck_pid", 0) or 0}
        done.append(칸)
        if 왜:
            칸["자취"] = out[-4000:]
            fails.append(칸)
            # ★ 다음 단계로 간다(`[175]`) — 하나를 살리자고 회차 전체를 세우지 않는다.

    느린 = [x for x in (_worst(done),) if x]     # 정렬 규칙은 한 곳이다([162])
    _note({"시각": datetime.now().isoformat(timespec="seconds"),
           "지금단계": "(회차 끝)", "상태": "실패" if fails else "완주",
           "끝낸단계": [d["단계"] for d in done], "단계별": done,
           "가장오래": 느린[0] if 느린 else None,
           "회차분": round((time.time() - 시작) / 60.0, 1)})
    if fails:
        _leave_trace(fails, done)
    else:
        _clear_trace()
    _log("== 회차 %s · %.1f분 · 가장 오래 걸린 단계: %s %s분" % (
        "실패" if fails else "완주", (time.time() - 시작) / 60.0,
        느린[0]["단계"] if 느린 else "?", 느린[0]["분"] if 느린 else "?"))
    _say("원본 자료 정리 %s — %d단계 · %.1f분%s" % (
        "실패" if fails else "완주", len(done), (time.time() - 시작) / 60.0,
        (" · 범인 %s" % (_worst(fails) or fails[0])["단계"]) if fails else ""))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
