# -*- coding: utf-8 -*-
"""공용 운영시간 규칙 — 이제 기본은 **보호시간 없음**이다.

★ 2026-08-11 지시("이제 엑셀에 사람이 입력안하고 앱에서만 입력할거야")로
  아침 입력 보호시간(08:00~09:30)은 **퇴역했다.** '류지영 매니저가 아침에
  관리대장을 입력한다'는 전제 하나가 daily_run·게시·반영 등 아홉 진입점을
  매일 90분 멈추고 있었는데, 그 전제가 사라졌다.

  실제 파일 충돌 방지는 시각이 아니라 **증거**가 한다 — `ledger_db` 의
  ~$ 잠금 감지·연기가 그대로 살아 있다(열려 있는 파일은 못 갈아끼운다, [171]).
  시각 기반 정지는 '사람이 그 시간에 입력한다'가 참일 때만 뜻이 있었다.

  되돌리는 스위치는 환경변수 하나다(글꼴 보호장치 [126]과 같은 원칙 —
  원래 값을 지우지 않고 남긴다):
      COUPANG_INPUT_WINDOW=08:00-09:30
  형식은 HH:MM-HH:MM. 지우면 다시 보호시간 없음. 값이 망가져 있으면
  보호시간 없음으로 동작한다(망가진 설정이 자동화를 멈추면 안 된다).
"""
from __future__ import annotations

import os
import re
from datetime import datetime, time, timedelta, timezone

KST = timezone(timedelta(hours=9), name="KST")

# 퇴역 전 원래 값 — "되돌리라"는 말을 들었을 때 무엇으로 돌아가는지 잃지 않는다.
LEGACY_START = time(8, 0)
LEGACY_END = time(9, 30)

_ENV_KEY = "COUPANG_INPUT_WINDOW"
_ENV_RE = re.compile(r"^\s*(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})\s*$")


def _window() -> tuple[time, time] | None:
    """환경변수가 켜져 있을 때만 (시작, 끝). 기본은 None(보호시간 없음)."""
    raw = os.environ.get(_ENV_KEY, "")
    m = _ENV_RE.match(raw)
    if not m:
        return None
    try:
        start = time(int(m.group(1)), int(m.group(2)))
        end = time(int(m.group(3)), int(m.group(4)))
    except ValueError:
        return None
    if start >= end:
        return None
    return start, end


def korea_now() -> datetime:
    return datetime.now(KST)


def is_input_window(now: datetime | None = None) -> bool:
    """보호시간 안이면 True. 기본(스위치 꺼짐)은 언제나 False다.

    테스트에서 넘기는 naive datetime은 이미 한국 현지 시각으로 해석한다.
    """
    win = _window()
    if win is None:
        return False
    now = now or korea_now()
    if now.tzinfo is not None:
        now = now.astimezone(KST)
    current = now.time().replace(tzinfo=None)
    return win[0] <= current < win[1]


def input_window_label() -> str:
    win = _window()
    if win is None:
        return "없음(앱 전용 입력 — 2026-08-11)"
    return "%02d:%02d~%02d:%02d KST" % (win[0].hour, win[0].minute,
                                        win[1].hour, win[1].minute)


# ═══════════════════════════════════════════════════════════════
# ★ 2026-09-08 형님 지시: "일딜 대조 전부 끄고 내가 지시할 때만 긁어오는
#   구조로 변경해" - 자동으로 도는 대조 회차를 끈다.
#
#   ★ 가르는 근거는 **부르는 쪽이 스스로 밝히는 것** 하나다.
#     COUPANG_UNATTENDED 로 가르면 안 된다 - 앱 [전체 대조 실행] 도 그것을
#     심으므로(app_server:12356) 형님이 누르시는 길까지 막힌다(지시의 정반대).
#   ★ 되돌리기 한 줄: COUPANG_AUTO_DAILY=1 ([126] 과 같은 보호장치).
#   ★ 끄는 것은 **자동으로 도는 것**뿐이다 - 사람 길 셋(앱 · 워크벤치 ·
#     python daily_run.py)은 한 글자도 안 막는다([172] 좁히는 것도 고장이다).
# ═══════════════════════════════════════════════════════════════
AUTO_ROUND_KEY = "COUPANG_AUTO_ROUND"      # 자동 경로가 스스로 심는 표시
AUTO_DAILY_KEY = "COUPANG_AUTO_DAILY"      # 되돌리기 스위치
AUTO_OFF_WHY = ("자동 대조가 꺼져 있다 - 사람이 명령하거나 앱 [전체 대조 실행] 을 "
                "누를 때만 돈다(2026-09-08 지시)")


def auto_round_blocked(auto=None, env=None):
    """**자동으로** 대조를 돌려도 되나 - 자동 경로의 유일한 판정([162]).

    돌려주는 것은 왜 막는지를 적은 **말**이거나 None(돌아도 된다)이다.
    ★ 조용히 막지 않는다([169]) - 부르는 쪽이 그 말을 자국·화면에 적을 수
      있어야 "몇 달째 대조가 안 돌았다"를 아무도 모르는 일이 안 생긴다.

    auto : True 면 부르는 쪽이 "나는 자동 경로다"라고 밝힌 것이다(워치독).
           None 이면 환경변수 표시를 본다(스케줄러가 .bat 에서 심는다).
    env  : 검증이 갈아 끼우기 위한 자리 - 기본은 진짜 os.environ 이다.
    """
    e = os.environ if env is None else env
    if str(e.get(AUTO_DAILY_KEY, "0")) != "0":
        return None                      # 사람이 다시 켰다([126])
    if auto is None:
        auto = str(e.get(AUTO_ROUND_KEY, "")) == "1"
    if not auto:
        return None                      # 사람 길 - 한 글자도 안 막는다([172])
    return AUTO_OFF_WHY
