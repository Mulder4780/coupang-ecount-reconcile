# -*- coding: utf-8 -*-
"""공용 운영시간 규칙.

류지영 매니저가 관리대장을 입력하는 08:00~09:30(KST)에는 자동 점검,
원장 읽기/쓰기, 게시 작업이 시작되지 않도록 모든 자동 진입점이 이 모듈을 쓴다.
"""
from __future__ import annotations

from datetime import datetime, time, timedelta, timezone

KST = timezone(timedelta(hours=9), name="KST")
INPUT_START = time(8, 0)
INPUT_END = time(9, 30)


def korea_now() -> datetime:
    return datetime.now(KST)


def is_input_window(now: datetime | None = None) -> bool:
    """08:00 이상 09:30 미만이면 True.

    테스트에서 넘기는 naive datetime은 이미 한국 현지 시각으로 해석한다.
    """
    now = now or korea_now()
    if now.tzinfo is not None:
        now = now.astimezone(KST)
    current = now.time().replace(tzinfo=None)
    return INPUT_START <= current < INPUT_END


def input_window_label() -> str:
    return "08:00~09:30 KST"
