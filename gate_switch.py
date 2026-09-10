# -*- coding: utf-8 -*-
"""합성검증(관문) 스위치 — **자동으로 도는 것만** 끈다 (2026-09-10 지시).

형님 지시: **"합성 검증 전체 중단시켜, 내가 지시할때만 진행해"**

★ **끄는 것은 '회차가 스스로 관문을 도는 것' 하나다.** 사람이 손으로 부르는 길은
  **한 글자도 안 막는다**([172]) — 형님이 "내가 지시할 때만 진행해" 라고 하셨으니
  그 길이 살아 있어야 지시가 뜻을 갖는다. 안 막는 길 셋(실측 2026-09-10):
    · `python tests/synthetic_check.py` (손으로)
    · `coupang_workbench.py` 의 메뉴 '합성검증만'
    · 앱 [실행] 화면의 `synthetic` 단추(`app_server.py`)

★ **막는 자리 넷**(실측 2026-09-10 · 전부 무인으로 도는 길):
    ① `daily_run` 의 **0단계** — 매일 09:50 대조의 첫 줄
    ② `noon_run.steps()` — 정오회차(12:00~13:00)
    ③ `overnight_run.green()` — 야간 회차의 관문
    ④ `handoff_review._synthetic_check()` — Terra→Sol 인수인계 검토

★★ **막았으면 회차는 그대로 돈다.** ①은 0단계라 예전에는 여기서 빨가면 그날 대조가
  통째로 안 돌았다(접수취소·객관완료·청구상태·오기입·사실대조·캠프 담당자).
  중단이면 **통과로 치고 다음 단계로 간다** — 안 그러면 끄는 것이 회차를 죽이는
  것이 되어 고치려던 것보다 나빠진다([172]).

★ **`daily_run._run_gate` 는 한 글자도 안 건드렸다**([211]). 그 함수를 부르는
  검사가 **열 곳이 넘는다**(`t414`·`t412`·`test_gate_proof`·`test_gate_split` …).
  거기에 문을 달면 그 검사들이 **이 표시 파일에 매여** 형님이 켜고 끌 때마다
  초록·빨강을 오간다. 그래서 문은 **부르는 자리**에 단다.

★ **되돌리기 두 길**(둘 다 되고, 환경변수가 더 세다):
    · `python gate_switch.py --resume`  (또는 `reports/합성검증_중단.json` 삭제)
    · 환경변수 `COUPANG_GATE_AUTO=1` ([126] 과 같은 보호장치)

⚠ **못 읽으면 '중단 아님'이다**([169] 를 이 자리에 맞게 정한 것).
  · 잘못 '중단 아님'으로 기울면 → 관문이 한 번 돈다(25분). **되돌릴 수 있다.**
  · 잘못 '중단'으로 기울면 → 형님이 다시 켜셨는데도 **영영 안 돌면서 조용하다.**
  물러나는 값은 회차 한 번이고, 잘못 막는 값은 **관문이 통째로 없어지는 것**이다.

⚠ **조용히 끄지 않는다**([169]). 인계 '먼저 처리할 것'에 알림 한 줄이 뜬다 —
  경보가 아니라 **알림**이다(형님이 스스로 끄신 것이지 고장이 아니다 · [170]).
"""
import io
import json
import os
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
MARK = os.path.join(ROOT, "reports", "합성검증_중단.json")
ENV = "COUPANG_GATE_AUTO"


def _mark_path():
    """표시 파일 자리 — 리포트 폴더를 옮기는 검증이 따라올 수 있게 부를 때 푼다([402])."""
    base = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
    if os.environ.get("COUPANG_REPORT_DIR"):
        return os.path.join(base, os.path.basename(MARK))
    return MARK


def stopped():
    """자동 관문이 중단인가 — (중단인가, 왜) 를 돌려준다.

    ★ 환경변수가 표시 파일보다 세다([126]) — 급할 때 한 줄로 되돌린다.
    ★ 못 읽으면 **중단 아님**(위 ⚠).
    """
    env = os.environ.get(ENV)
    if env is not None:
        if str(env).strip() in ("1", "true", "True", "on", "ON"):
            return False, "환경변수 %s=1 — 자동 관문을 켜 두었다" % ENV
        if str(env).strip() in ("0", "false", "False", "off", "OFF"):
            return True, "환경변수 %s=0 — 자동 관문을 꺼 두었다" % ENV
    try:
        with io.open(_mark_path(), encoding="utf-8") as fh:
            doc = json.load(fh)
    except (OSError, ValueError):
        return False, ""
    if not isinstance(doc, dict) or not doc.get("중단"):
        return False, ""
    why = str(doc.get("왜") or "형님 지시로 자동 관문을 중단했다")
    when = str(doc.get("때") or "")
    return True, ("%s (%s)" % (why, when) if when else why)


def skip_result(step_name="합성검증"):
    """중단일 때 회차에 돌려줄 결과 — `run()`/`_run_gate` 와 **같은 모양**이다.

    ★ `ok=True` 이고 `out` 에 `ALL GREEN` 이 **안 들어간다**. 부르는 쪽이 그 글자를
      확인하므로([400] — 초록은 두 숫자로 센다) 이것을 초록으로 오인하면 안 된다.
      대신 부르는 쪽이 '중단'을 먼저 보고 통과시킨다.
    """
    _, why = stopped()
    return {
        "name": step_name,
        "ok": True,
        "returncode": 0,
        "skipped_by_switch": True,
        "out": ("자동 관문 중단 — 건너뜀 (%s) · 돌리려면 "
                "`python tests/synthetic_check.py`" % (why or "형님 지시")),
    }


def stop(why="형님 지시 — 내가 지시할 때만 진행한다"):
    path = _mark_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = {
        "중단": True,
        "왜": why,
        "때": datetime.now().astimezone().isoformat(timespec="seconds"),
        "되돌리기": "python gate_switch.py --resume  (또는 %s=1)" % ENV,
        "안막는길": [
            "python tests/synthetic_check.py",
            "coupang_workbench.py 메뉴 '합성검증만'",
            "앱 [실행] 화면의 synthetic 단추",
        ],
    }
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)
    return doc


def resume():
    path = _mark_path()
    try:
        os.remove(path)
        return True
    except OSError:
        return False


def note():
    """인계·리포트에 실을 한 줄 — 중단이 아니면 **빈 문자열**([170])."""
    off, why = stopped()
    if not off:
        return ""
    return ("자동 합성검증(관문)이 중단돼 있다 — %s. "
            "회차는 관문을 건너뛰고 그대로 돈다. "
            "돌리려면 `python tests/synthetic_check.py`" % (why or "형님 지시"))


def main():
    import sys
    argv = sys.argv[1:]
    if "--stop" in argv:
        d = stop()
        print("자동 관문 중단: %s" % d["때"])
    elif "--resume" in argv:
        print("자동 관문 다시 켬" if resume() else "표시 파일이 없다 (이미 켜져 있다)")
    off, why = stopped()
    print("지금: %s%s" % ("중단" if off else "돎", (" — " + why) if why else ""))
    print("표시 파일: %s" % _mark_path())


if __name__ == "__main__":
    main()
