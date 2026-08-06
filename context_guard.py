# -*- coding: utf-8 -*-
"""
context_guard.py — 컨텍스트가 다 차기 **전에** 알아서 정리시킨다 (2026-08-06 지시)

사용자 지시: "컨텍스트 윈도우 다 차기전 컴팩팅 자동으로 하는 알고리즘 추가"

무엇이 부족했나
  요약(compact) 자체는 이미 자동이다(`.claude/settings.json` 의 `autoCompactEnabled`).
  문제는 **그 직전까지 AI 가 새 작업을 벌인다**는 것이었다. PreCompact 훅이 인계를
  남겨 주기는 해도, 반쯤 하다 만 작업은 아무도 끝내 주지 않는다. 즉 빠진 것은
  "요약을 실행하는 기능"이 아니라 **"차 가고 있다는 사실을 제때 알려 주는 눈"** 이다.

이 파일이 그 눈이다
  · 매 사용자 입력마다(UserPromptSubmit 훅) 지금 세션이 한도의 몇 %인지 **실제 사용량**
    으로 잰다. 추정이 아니라 대화 기록에 남은 `usage`(입력+캐시읽기+캐시생성) 합이다.
  · 단계마다 대화에 한 줄을 끼워 넣는다 — AI 는 그 줄을 보고 스스로 행동을 바꾼다.
      70% 예고   : 큰 작업을 새로 벌이지 않는다
      85% 마무리 : 지금 것만 끝내고 커밋·인계로 전환한다 + **인계를 자동으로 남긴다**
      95% 즉시   : 새 작업 금지, 사람에게 `/compact` 를 권한다
  · 85% 를 처음 넘긴 순간 `session_wrapup.py` 를 **한 번** 돌린다(세션당 1회).
    그래서 PreCompact 훅이 판올림으로 사라져도 인계는 남는다 — 이중 안전장치다.

원칙
  · **절대 실패하지 않는다.** 어떤 예외에도 exit 0, 훅이 사람 입력을 막지 않는다.
  · **싸다.** 17MB 짜리 대화 기록을 통째로 읽지 않고 꼬리 2MB 만 훑는다.
  · 판정 결과는 `reports/컨텍스트_사용량.json` 에 남는다(워치독·인계 문서가 읽는다).

쓰는 법
  python ecount/context_guard.py --print          # 지금 몇 %인가
  python ecount/context_guard.py --hook           # 훅에서(표준입력 JSON)
  python ecount/context_guard.py --print --limit 450000
"""
import os
import re
import sys
import json
import glob
import argparse
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
STATE_PATH = os.path.join(REPORT_DIR, "컨텍스트_사용량.json")
SETTINGS = os.path.join(os.path.dirname(ROOT), ".claude", "settings.json")

# 한도를 못 읽었을 때의 보수적 기본값(설정 기본 500,000 의 90%)
DEFAULT_LIMIT = 450_000
TAIL_BYTES = 2_000_000        # 대화 기록 꼬리만 읽는다 — 앞쪽은 이미 지난 회차다

# 단계: (비율, 이름, AI 에게 끼워 넣을 지시)
STAGES = (
    (0.95, "즉시", "새 작업을 시작하지 말 것. 지금 손에 든 것만 저장·커밋하고, "
                   "사람에게 `/compact` 를 권한 뒤 마무리한다."),
    (0.85, "마무리", "새 작업을 벌이지 말 것. 하던 것만 끝내고 커밋·인계로 전환한다. "
                     "인계 자동 정리는 이미 돌았다(reports/세션마무리_기록.json)."),
    (0.70, "예고", "컨텍스트가 차 가고 있다. 큰 작업을 새로 벌이지 말고, "
                   "지금 것을 작게 끊어 커밋해 둔다."),
)


def _read_limit():
    """압축 시점(autoCompactWindow)을 설정에서 읽는다. 이 값이 곧 한도다."""
    try:
        with open(SETTINGS, encoding="utf-8") as fh:
            s = json.load(fh)
        if s.get("autoCompactEnabled") is False:
            return DEFAULT_LIMIT, False
        v = int(s.get("autoCompactWindow") or DEFAULT_LIMIT)
        return (v if 100_000 <= v <= 1_000_000 else DEFAULT_LIMIT), True
    except Exception:
        return DEFAULT_LIMIT, True


def _project_dir():
    """이 프로젝트의 대화 기록 폴더 — 경로를 슬러그로 바꾼 Claude Code 규칙 그대로.

    규칙은 **글자 하나당 대시 하나**다(묶지 않는다). `C:\\Users\\…` → `C--Users-…`
    — 여기서 `[^A-Za-z0-9]+` 로 묶었다가 폴더를 못 찾아 0% 로 나왔었다.
    """
    home = os.path.expanduser("~")
    base = os.path.join(home, ".claude", "projects")
    for proj in (os.path.dirname(ROOT), ROOT):
        d = os.path.join(base, re.sub(r"[^A-Za-z0-9]", "-", proj))
        if os.path.isdir(d):
            return d
    return ""


def _transcript(session_id="", given=""):
    """이 세션의 대화 기록 파일. 훅이 준 경로를 최우선으로 믿는다."""
    if given and os.path.exists(given):
        return given
    d = _project_dir()
    if not d:
        return ""
    if session_id:
        p = os.path.join(d, session_id + ".jsonl")
        if os.path.exists(p):
            return p
    files = glob.glob(os.path.join(d, "*.jsonl"))
    if not files:
        return ""
    return max(files, key=os.path.getmtime)


def _used_tokens(path):
    """마지막 assistant 응답의 usage 합 = 지금 프롬프트가 실제로 차지한 크기."""
    if not path or not os.path.exists(path):
        return 0
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as fh:
            if size > TAIL_BYTES:
                fh.seek(size - TAIL_BYTES)
                fh.readline()          # 잘린 첫 줄은 버린다
            raw = fh.read()
    except Exception:
        return 0
    used = 0
    for line in raw.decode("utf-8", "ignore").splitlines():
        if '"usage"' not in line:
            continue
        try:
            u = (json.loads(line).get("message") or {}).get("usage") or {}
        except Exception:
            continue
        n = (int(u.get("input_tokens") or 0)
             + int(u.get("cache_read_input_tokens") or 0)
             + int(u.get("cache_creation_input_tokens") or 0))
        if n:
            used = n                    # 꼬리로 갈수록 최신 — 마지막 값이 지금 크기다
    return used


def _stage(ratio):
    for cut, name, advice in STAGES:
        if ratio >= cut:
            return cut, name, advice
    return 0.0, "여유", ""


def _load_state():
    try:
        with open(STATE_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_state(st):
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        tmp = STATE_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(st, fh, ensure_ascii=False, indent=1)
        os.replace(tmp, STATE_PATH)
    except Exception:
        pass


def _run_wrapup(who, reason):
    """인계를 남긴다. 오래 걸려도 사람 입력을 막지 않게 조용히·짧게."""
    import subprocess
    try:
        subprocess.run(
            [sys.executable, os.path.join(ROOT, "session_wrapup.py"),
             "--who", who, "--reason", reason, "--quiet"],
            cwd=ROOT, timeout=600,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


def measure(session_id="", transcript="", limit=0, who="claude", act=True):
    """지금 사용량을 재고, 단계가 올라갔으면 필요한 조치를 한다."""
    lim, auto_on = _read_limit()
    lim = int(limit or lim)
    path = _transcript(session_id, transcript)
    used = _used_tokens(path)
    ratio = (used / lim) if lim else 0.0
    cut, name, advice = _stage(ratio)

    sid = session_id or os.environ.get("CLAUDE_CODE_SESSION_ID") or ""
    st = _load_state()
    same = st.get("session") == sid and sid
    prev = float(st.get("stage_cut") or 0) if same else 0.0
    wrapped = bool(st.get("wrapped")) if same else False

    fresh = cut > prev                       # 이번에 단계가 **올라갔다**
    did_wrap = False
    if act and fresh and cut >= 0.85 and not wrapped:
        did_wrap = _run_wrapup(who, "auto-context-guard")
        wrapped = wrapped or did_wrap

    res = {
        "session": sid, "transcript": os.path.basename(path or ""),
        "used": used, "limit": lim, "ratio": round(ratio, 4),
        "percent": round(ratio * 100, 1), "stage": name, "stage_cut": cut,
        "advice": advice, "fresh": fresh, "wrapped": wrapped,
        "wrapup_ran_now": did_wrap, "auto_compact": auto_on,
        "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    if act:
        _save_state(res)
    return res


def _message(r):
    """AI 대화에 끼워 넣을 한 문단. 단계가 올라간 순간에만 낸다(매번 떠들지 않는다)."""
    if not r["advice"] or not r["fresh"]:
        return ""
    lines = [
        "[컨텍스트 감시] 이 세션이 한도의 약 {p}% 를 썼다 "
        "({u:,} / {l:,} 토큰 · 단계: {s}).".format(
            p=r["percent"], u=r["used"], l=r["limit"], s=r["stage"]),
        r["advice"],
    ]
    if r["wrapup_ran_now"]:
        lines.append("인계 자동 정리를 방금 돌렸다 — 큐→DB·점유해제·커밋·인계기록 완료.")
    if not r["auto_compact"]:
        lines.append("※ 자동 요약이 꺼져 있다. 사람에게 `/compact` 를 권할 것.")
    return "\n".join(lines)


def hook():
    """UserPromptSubmit 훅 진입점. 무슨 일이 있어도 exit 0."""
    payload = {}
    try:
        raw = sys.stdin.read()
        if raw.strip():
            payload = json.loads(raw)
    except Exception:
        payload = {}
    try:
        r = measure(session_id=payload.get("session_id") or "",
                    transcript=payload.get("transcript_path") or "")
        msg = _message(r)
    except Exception:
        msg = ""
    if msg:
        print(json.dumps({
            "hookSpecificOutput": {
                "hookEventName": "UserPromptSubmit",
                "additionalContext": msg,
            }
        }, ensure_ascii=False))
    return 0


def main():
    ap = argparse.ArgumentParser(description="컨텍스트가 다 차기 전에 마무리로 전환시킨다")
    ap.add_argument("--hook", action="store_true", help="훅에서 호출(표준입력 JSON)")
    ap.add_argument("--print", dest="show", action="store_true", help="지금 사용량 한 줄")
    ap.add_argument("--json", action="store_true", help="JSON 으로")
    ap.add_argument("--limit", type=int, default=0, help="한도를 손으로 지정(토큰)")
    ap.add_argument("--who", default="claude", help="claude | codex")
    ap.add_argument("--dry", action="store_true", help="재기만 하고 인계·기록은 하지 않는다")
    a = ap.parse_args()

    if a.hook:
        return hook()

    r = measure(limit=a.limit, who=a.who, act=not a.dry)
    if a.json:
        print(json.dumps(r, ensure_ascii=False))
    else:
        print("컨텍스트 {p}%  ({u:,} / {l:,} 토큰) · 단계 {s}".format(
            p=r["percent"], u=r["used"], l=r["limit"], s=r["stage"]))
        if r["advice"]:
            print("  → " + r["advice"])
        if r["wrapup_ran_now"]:
            print("  → 인계 자동 정리 완료")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        sys.exit(0)          # 훅이 사람 입력을 막지 않는다
