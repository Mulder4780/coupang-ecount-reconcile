# -*- coding: utf-8 -*-
"""child_budget.py — 바깥이 준 **시간 예산**을 읽는 자리 하나 (2026-08-25)

★ 왜 이 파일이 있나 — 2026-08-25 실사고([427]).
  자율복구 대기표의 `밴드 게시글 보관` 은 제한 1800초를 꽉 채우다 **SIGKILL(-9)** 로
  끊겼다. 파이썬 stdout 은 파이프에 물리면 **블록 버퍼**라 그때까지 찍은 줄이
  **버퍼째 사라진다** — 그래서 27회 시도의 자국이 `returncode=-9` 다섯 글자뿐이었고,
  자율복구는 그것을 *"10회 넘게 재시도해도 안 풀린다"* 로 읽어 매일 가짜 경보를 세웠다
  ([170]). **일은 되고 있었다 — 잃은 것은 '얼마나 했나' 였다**([169]).

  고침은 예산 하나다: 다 되면 **새로 안 넣고 보고서를 쓰고** 증분(75)으로 돌아온다.
  그러면 자율복구가 `이어감` 으로 세고 실패로 안 센다.

★ **판정은 여기 한 곳이다**([162]). 쓰는 스크립트가 각자 제 손으로 적으면, 한 곳이
  "못 읽으면 예산 없음" 규칙을 빠뜨리는 날 **그 스크립트만 조용히 예전처럼 죽는다**
  — 오류도 안 나고 화면도 멀쩡하다([169]).

★ **딸린 것이 하나도 없다**(표준 모듈만). 자식 스크립트가 이것 하나 때문에 무거워지면
  안 된다.

쓰는 법 — 일을 시작하기 전에 한 번, 비싼 걸음마다 한 번:

    import child_budget
    budget = child_budget.start("STMT_ARCHIVE_BUDGET_SEC")
    ...
    if child_budget.over():      # 새로 안 넣는다(도는 것은 끝까지 둔다)
        cut = 1
        break
    ...
    if cut:
        print("  ★ 시간 예산(%d초)이 다 되어 …" % budget)
        return child_budget.INCREMENTAL_RETURN_CODE
"""
import os
import time

# 자율복구·회차가 아는 값 — `autopilot.INCREMENTAL_RETURN_CODE` 와 같아야 한다.
INCREMENTAL_RETURN_CODE = 75

_DEADLINE = None


def start(env_name):
    """바깥이 준 예산(초)을 읽어 마감을 세운다 — **안 주면 무제한**(예전 그대로)이다.

    ★ 못 읽으면 **'예산 없음'** 이지 0초가 아니다([169]). 0초로 읽으면 그 자식은
      아무 일도 못 하고 곧장 돌아와 **영원히 진도가 안 나간다**([199] 와 같은 모양).
    """
    global _DEADLINE
    try:
        sec = int(os.environ.get(env_name) or 0)
    except (TypeError, ValueError):
        sec = 0
    _DEADLINE = (time.monotonic() + sec) if sec > 0 else None
    return sec if sec > 0 else 0


def over():
    """예산이 다 됐나 — **예산이 없으면 언제나 거짓**이다."""
    return _DEADLINE is not None and time.monotonic() >= _DEADLINE


def clear():
    """마감을 지운다(검사가 되돌릴 때 쓴다 — 모듈 전역이라 프로세스 하나뿐이다 · [371])."""
    global _DEADLINE
    _DEADLINE = None
