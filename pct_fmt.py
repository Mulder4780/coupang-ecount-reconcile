# -*- coding: utf-8 -*-
"""
pct_fmt.py — 비율(%) 계산·표기의 단일 규칙 (2026-08-05 사용자 지시)

지시: "비율 표기는 소수점 1자리까지 표기, 1건이 안되도 100%로 보이는 문제 해결"

왜 한 곳에 모으나:
  화면·리포트·대표보고가 각자 `round(done/total*100)` 을 쓰다 보니
  · 999/1000(99.9%)가 반올림으로 **100%** 가 되어 "다 됐다"로 읽혔고,
  · 1/1000(0.1%)은 **0%** 가 되어 "하나도 안 했다"로 읽혔다.
  숫자 하나가 대표 보고의 결론을 바꾸므로, 계산과 글자 모양을 이 파일이 정한다.

규칙:
  · 소수점 **1자리** 고정 (예: 99.9% / 12.5% / 100.0%)
  · **완료가 아니면 절대 100.0% 로 보이지 않는다** — 반올림이 100에 닿으면 99.9 로 내린다.
  · **한 건이라도 했으면 0.0% 로 보이지 않는다** — 0에 닿으면 0.1 로 올린다.
  · 모수가 0이면 비율이 없다(None) — "대상 없음"으로 표기하고 0% 라고 하지 않는다.
"""

__all__ = ["pct", "pct_text"]


def pct(done, total):
    """비율을 소수점 1자리로. 모수가 없으면 None(=대상 없음)."""
    try:
        done, total = float(done or 0), float(total or 0)
    except (TypeError, ValueError):
        return None
    if total <= 0:
        return None
    v = round(done * 100.0 / total, 1)
    # 끝의 두 줄이 이 파일의 존재 이유다 — 반올림이 사실을 뒤집지 못하게 한다.
    if v >= 100.0 and done < total:
        v = 99.9
    if v <= 0.0 and done > 0:
        v = 0.1
    return v


def pct_text(done, total, none_text="대상 없음"):
    """화면·리포트에 그대로 쓰는 글자. 모수가 없으면 '대상 없음'."""
    v = pct(done, total)
    return none_text if v is None else "%.1f%%" % v
