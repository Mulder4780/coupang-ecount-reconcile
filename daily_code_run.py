# -*- coding: utf-8 -*-
"""daily_code_run.py — 코딩 세션 몫을 **하루 한 번 12~13시**에 무인으로 돈다

사용자 지시(2026-08-11): "하루한번 12시에서 13시 사이 다른 세션이랑 충돌 안나게
실행 알고리즘 코딩해(자동화 100%, 컴팩팅도 자동화, 내가 손 안대게 처리)"

★ 왜 필요한가 — 지금 있던 것은 **창이 열려 있어야만** 돌았다.
  코딩 세션 자동 루프(`/loop`)는 Claude 창 안에서 사는 것이라 창을 닫으면 같이 죽는다.
  결정론적 회차(08:00 재수집 · 09:50 대조 · 11:00·15:00 보관본)는 **업무 자료**를 보지
  코드 건강(합성검증 빨강 · 미커밋 · 안 밀린 커밋)은 아무도 안 본다.
  그래서 창이 닫힌 날은 그 축이 통째로 빈다 — 그리고 **어느 화면에도 티가 안 난다.**

무엇을 하나 (전부 읽기 전용이거나 되돌릴 수 있는 것뿐)
  ① 합성검증 — 빨강이면 그 자체가 제일 급한 일이다
  ② 컴팩팅 자동화가 살아 있나 — `autoCompactEnabled` · PreCompact 훅
  ③ 인계 문서 갱신 — 결과가 '먼저 처리할 것'으로 사람 눈에 닿는 유일한 길
  **수집은 절대 안 한다**(수집 세션 몫) · 엑셀도 안 연다 · 원장도 안 쓴다.

★ 충돌은 '피하는' 것이지 '이기는' 것이 아니다.
  남이 잡고 있으면 **오늘은 물러난다.** 기다리거나 빼앗지 않는다 —
  이 회차가 하는 일은 하루 미뤄도 잃는 것이 없고, 빼앗으면 남의 작업이 통째로 묻힌다.
  물러난 것은 **실패가 아니다**(그렇게 세면 매일 빨간불이 뜨고 아무도 안 본다).

검증 [220].
"""
import argparse
import json
import os
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE, "reports")
STATE = os.path.join(REPORT_DIR, "코딩회차_상태.json")

WINDOW_FROM, WINDOW_TO = 12, 13      # 12:00 이상 13:00 미만 (사용자 지시)
BUDGET_MIN = 40                      # 이 회차는 오래 끌 일이 없다. 넘으면 남기고 끝낸다

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def in_window(now=None):
    """지금이 12:00~13:00 인가. **창 밖이면 아무것도 안 한다.**

    ★ PC 가 꺼져 있어 늦게 떴으면 그날은 건너뛴다 — 창을 어긴 채 도는 것보다 낫다.
      이 회차는 Z: 를 훑는 다른 회차(09:50·11:00) 사이의 **빈 시간**에 놓여 있고,
      그 배치 자체가 충돌 회피의 첫 겹이다.
    """
    now = now or datetime.now()
    return WINDOW_FROM <= now.hour < WINDOW_TO


def _load():
    try:
        with open(STATE, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def _save(d):
    os.makedirs(REPORT_DIR, exist_ok=True)
    tmp = f"{STATE}.{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE)


def ran_today(now=None):
    """오늘 이미 **끝까지** 돌았나. 물러난 날은 '돌았다'로 치지 않는다 —
    그러면 남이 잠깐 잡고 있던 날의 몫이 영영 사라진다."""
    now = now or datetime.now()
    d = _load()
    return d.get("완주일") == now.strftime("%Y-%m-%d")


def busy_reason():
    """다른 세션·회차와 부딪히나. 부딪히면 **그 이유를 말하고 물러난다**.

    판단을 새로 만들지 않는다 — 이미 있는 눈을 그대로 빌린다([225] 와 같은 원칙):
      · 일일대조가 도는 중인가 → `session_handoff._daily_run_inflight`
      · `code` 를 남이 잡고 있나 → `ai_claim.take` 가 스스로 판정한다(아래 run 에서)
      · 사람이 지금 창에서 일하고 있나 → `session_wrapup._other_live_sessions`
    """
    try:
        import session_handoff as SH
        h = SH._daily_run_inflight()
        if h is not None:
            return f"일일자동대조가 {h}시간째 도는 중 (Z: 를 같이 훑으면 둘 다 느려진다)"
    except Exception:
        pass
    try:
        import session_wrapup as SW
        live = SW._other_live_sessions()
        if live:
            # ★ 사람이 창에서 일하는 중이면 그 창이 할 일이다. 여기서 같은 파일을
            #   건드리면 사고 #36(같은 일을 둘이) 이 그대로 재현된다.
            return f"다른 세션 {len(live)}개가 살아 있다 (그 창이 할 일이다)"
    except Exception:
        pass
    return None


def compact_health():
    """**컴팩팅 자동화가 살아 있나** — 사용자 지시의 '컴팩팅도 자동화' 몫.

    고치지는 않는다. `.claude/settings.json` 은 사람 설정 파일이라 기계가 손대면
    사람이 바꾼 값이 조용히 되돌아간다. **꺼져 있다는 사실을 말하는 것**이 여기 몫이다.
    """
    p = os.path.join(os.path.dirname(BASE), ".claude", "settings.json")
    try:
        with open(p, encoding="utf-8") as fh:
            cfg = json.load(fh)
    except (OSError, ValueError):
        return {"확인": False, "사유": "settings.json 을 못 읽었다"}
    on = bool(cfg.get("autoCompactEnabled"))
    hooks = json.dumps(cfg.get("hooks") or {}, ensure_ascii=False)
    pre = "PreCompact" in hooks and "session_wrapup" in hooks
    return {"확인": True, "자동요약": on, "인계훅": pre,
            "정상": bool(on and pre)}


def _step(name, args, budget_s):
    """한 단계 — `proc_guard` 로 **반드시 끝난다**([179] · `subprocess.run(timeout=)` 금지).

    ⚠ `run_tree` 는 dict 가 아니라 **`ProcessResult`**(returncode/stdout/stderr/
      timed_out/stuck_pid)를 돌려준다. dict 처럼 `.get()` 으로 읽으면 매번 예외가 나서
      **모든 단계가 조용히 실패**한다 — 회차는 멀쩡히 돌고 빨간불만 매일 뜬다.
    """
    try:
        import proc_guard
        r = proc_guard.run_tree([sys.executable] + args, cwd=BASE, timeout=budget_s)
        tail = (r.stdout or "").strip().splitlines()
        return {"단계": name, "ok": r.returncode == 0 and not r.timed_out,
                "코드": r.returncode, "시간초과": bool(r.timed_out),
                "끝": tail[-1][:200] if tail else ""}
    except Exception as e:
        return {"단계": name, "ok": False, "오류": str(e)[:200]}


def run(force=False, now=None):
    """회차 본체. 돌았으면 결과 dict, 물러났으면 `{"상태": "물러남", ...}`."""
    now = now or datetime.now()
    if not force and not in_window(now):
        return {"상태": "창밖", "사유": f"{WINDOW_FROM}:00~{WINDOW_TO}:00 에만 돈다",
                "지금": now.strftime("%H:%M")}
    if not force and ran_today(now):
        return {"상태": "오늘완주", "사유": "하루 한 번이다"}
    why = busy_reason()
    if why and not force:
        # ★ 물러난 것도 기록한다 — 며칠째 못 돌고 있다면 그게 사람이 알아야 할 사실이다.
        d = _load()
        d["마지막물러남"] = {"때": now.isoformat(timespec="seconds"), "사유": why}
        _save(d)
        return {"상태": "물러남", "사유": why}

    got_claim = False
    try:
        import ai_claim
        # take 는 남이 잡고 있으면 스스로 False 를 준다(이유도 화면에 찍는다).
        got_claim = ai_claim.take("claude", "code", "코딩 회차(12시) 자동 점검")
    except Exception:
        got_claim = False
    if not got_claim and not force:
        d = _load()
        d["마지막물러남"] = {"때": now.isoformat(timespec="seconds"),
                             "사유": "code 를 다른 세션이 잡고 있다"}
        _save(d)
        return {"상태": "물러남", "사유": "code 를 다른 세션이 잡고 있다"}

    steps, t0 = [], datetime.now()
    try:
        steps.append(_step("합성검증", [os.path.join("tests", "synthetic_check.py")],
                           BUDGET_MIN * 60))
        steps.append(_step("인계갱신", ["session_handoff.py", "--snapshot"], 300))
    finally:
        if got_claim:
            try:
                import ai_claim
                ai_claim.free("claude", "code")
            except Exception:
                pass

    comp = compact_health()
    bad = [s["단계"] for s in steps if not s.get("ok")]
    out = {"상태": "완주", "때": now.isoformat(timespec="seconds"),
           "걸린분": round((datetime.now() - t0).total_seconds() / 60, 1),
           "단계": steps, "실패단계": bad, "컴팩팅": comp}
    d = _load()
    d.update(out)
    d["완주일"] = now.strftime("%Y-%m-%d")      # 완주했을 때만 오늘을 적는다
    _save(d)
    return out


def banner():
    """인계 문서에 올릴 한 줄 — 문제가 있을 때만 말한다(없으면 None).

    ★ 매일 '정상'을 적으면 아무도 안 본다([196] 의 교훈). 빨강일 때만 올린다.
    """
    d = _load()
    if not d:
        return None
    bits = []
    if d.get("실패단계"):
        bits.append("코딩 회차 실패단계: " + ", ".join(d["실패단계"]))
    comp = d.get("컴팩팅") or {}
    if comp.get("확인") and not comp.get("정상"):
        miss = []
        if not comp.get("자동요약"):
            miss.append("autoCompactEnabled 꺼짐")
        if not comp.get("인계훅"):
            miss.append("PreCompact 인계훅 없음")
        bits.append("컴팩팅 자동화가 끊겼다 — " + " · ".join(miss))
    # 며칠째 못 돌았나 — 물러남이 이어지면 그것도 사실이다
    try:
        last = d.get("완주일")
        if last:
            days = (datetime.now().date() - datetime.strptime(last, "%Y-%m-%d").date()).days
            if days >= 3:
                bits.append(f"코딩 회차가 {days}일째 완주하지 않았다"
                            f" (마지막 물러남: {(d.get('마지막물러남') or {}).get('사유', '?')})")
    except (ValueError, TypeError):
        pass
    return " · ".join(bits) or None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", action="store_true", help="회차 실행(창·하루한번·충돌 규칙 적용)")
    ap.add_argument("--force", action="store_true", help="규칙 무시하고 지금 실행(사람 지시)")
    ap.add_argument("--status", action="store_true", help="지금 상태만 본다")
    a = ap.parse_args()
    if a.status or not (a.run or a.force):
        d = _load()
        print(json.dumps({"창안": in_window(), "오늘완주": ran_today(),
                          "막힘": busy_reason(), "컴팩팅": compact_health(),
                          "마지막": d.get("때"), "실패단계": d.get("실패단계")},
                         ensure_ascii=False, indent=1))
        return 0
    got = run(force=a.force)
    print(json.dumps({k: v for k, v in got.items() if k != "단계"},
                     ensure_ascii=False, indent=1))
    # ★ 물러남은 실패가 아니다 — 0 으로 끝내야 스케줄러가 매일 빨간불을 켜지 않는다.
    return 1 if got.get("실패단계") else 0


if __name__ == "__main__":
    raise SystemExit(main())
