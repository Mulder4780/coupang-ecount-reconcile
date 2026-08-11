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
