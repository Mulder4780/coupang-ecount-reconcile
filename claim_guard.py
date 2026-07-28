# -*- coding: utf-8 -*-
"""
claim_guard.py — 원장을 고치는 도구가 **남의 점유를 밟지 못하게** 실제로 막는다
================================================================================
`ai_claim.py` 는 규칙을 적어 두는 곳이고, 이 파일은 그 규칙을 **강제**한다.

왜 필요한가
  문서에만 적어 두면 잊는다. AI든 사람이든 급하면 그냥 도구를 돌리고, 그 순간
  두 AI가 각자 vN+1을 만들어 **한쪽 작업이 통째로 묻힌다**(되돌릴 수 없다).
  그래서 원장을 쓰는 도구는 실행 직전에 여기를 거치게 한다.

동작
  · 다른 AI가 `ledger` 를 잡고 있으면 **멈추고 이유를 알려 준다**.
  · 아무도 안 잡고 있으면 조용히 자동으로 잡는다(사람이 매번 명령할 필요 없음).
  · 내가 잡은 것이면 그대로 진행한다.
  · `CSOS_AI=codex` 처럼 환경변수로 누구인지 알린다(기본값 unknown 은 막지 않는다 —
    사람이 직접 돌리는 경우를 방해하지 않기 위해서다).

  from claim_guard import require
  require("ledger", "confirm_fill 반영")     # 못 잡으면 SystemExit
"""
import os, sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)


def whoami():
    """이 프로세스가 누구인지. 환경변수로 알린다(CSOS_AI=claude|codex)."""
    return (os.environ.get("CSOS_AI") or "").strip().lower() or "unknown"


def require(lock="ledger", why="", who=None):
    """점유를 확보한다. 남이 잡고 있으면 실행을 멈춘다.

    사람이 직접 돌릴 때(CSOS_AI 미설정)는 막지 않는다 — 자동화가 사람 손을 묶으면 안 된다.
    대신 남이 잡고 있으면 **경고는 한다**.
    """
    me = (who or whoami())
    try:
        import ai_claim
    except Exception:
        return True                      # 조율 모듈이 없으면 예전처럼 그냥 진행

    try:
        cur = ai_claim.load().get(lock)
    except Exception:
        return True

    if cur and cur.get("who") not in (me, None):
        mins = 0
        try:
            import time
            mins = int((time.time() - cur.get("at", 0)) / 60)
        except Exception:
            pass
        msg = (f"★ '{lock}' 은 {cur.get('who')} 가 잡고 있습니다 "
               f"({mins}분 전 · {cur.get('why', '')}).")
        if me == "unknown":
            print(msg + "\n  (사람이 직접 실행한 것으로 보고 계속합니다 — 충돌에 주의하세요)")
            return True
        print(msg)
        print("  동시에 원장을 고치면 각자 vN+1을 만들어 한쪽이 통째로 묻힙니다.")
        print("  → 상대가 끝낼 때까지 기다리거나 다른 일을 하세요.")
        sys.exit(3)

    if me != "unknown":
        try:
            ai_claim.take(me, lock, why or os.path.basename(sys.argv[0]))
        except Exception:
            pass
    return True


def release(lock="ledger", who=None):
    me = (who or whoami())
    if me == "unknown":
        return
    try:
        import ai_claim
        ai_claim.free(me, lock)
    except Exception:
        pass
