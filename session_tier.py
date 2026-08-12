#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""이 **대화**의 모델·노력을 근거로 고른다 (2026-08-13 지시).

사용자 지시: "이 세션에서 크래딧 최대한 아끼고 최고의 퍼포먼스를 낼 수 있는 모델
자동 선택하고 추론 모드도 자동으로 선택해서 적용해서 진행하는 알고리즘 구현해"

★ 이미 있던 둘과 무엇이 다른가 — **비어 있던 자리가 제일 비싼 자리였다**
  · `ai_tier`([230])  = 기계가 부르는 CLI(agent_dispatch·autopilot)의 값을 고른다.
  · `loop_policy`([231]) = 루프 **틱**의 무게를 고른다.
  · 그런데 크레딧을 제일 많이 쓰는 것은 **사람과 지금 하는 이 대화**인데 거기는
    아무도 안 골랐다. "지금 몇 건이야" 한 줄이나 "구조를 새로 짜라"나 **같은 값**을
    치렀다. 자동화를 아무리 늘려도 이 자리를 안 고르면 절약이 안 된다.

★ **싸게 하는 것이 목적이 아니라 값에 맞게 하는 것이 목적이다**(ai_tier 와 같은 이유).
  이 프로젝트에서 잘못된 판단은 못 한 판단보다 나쁘다([172]·[196]). 그래서
  **올리는 문은 넓게, 내리는 문은 좁게** 판다. 내리는 근거는 지금 딱 하나뿐이다
  (아래 '컨텍스트' 항목).

무엇을 보는가 (전부 이미 있는 것 — 비싼 탐색을 새로 하지 않는다, [168])
  ① **요청문**            이번 턴에 무엇을 시켰나 (UserPromptSubmit 훅이 준다)
  ② **디스크 근거**        `loop_policy` 가 읽는 인계 '먼저 처리할 것' · 회차 경보
  ③ **컨텍스트 단계**      `reports/컨텍스트_사용량.json` 의 `stage`

어떻게 합치는가 — **셋의 역할이 다르다. 뭉뚱그리면 서로를 지운다.**
  · 갈래를 정하는 것은 **요청문**이다. 사람이 무엇을 시킬지는 사람이 정한다.
  · 디스크 근거는 **바닥만 올린다**(조회 → 재시도). 천장은 안 올린다 —
    세워 둔 일이 있다고 "지금 몇 건이야"까지 opus 로 물으면 이 알고리즘이 없느니만
    못하다. 디스크가 말하는 것은 '지금 조용하지 않다'이지 '이 질문이 어렵다'가 아니다.
  · 컨텍스트 단계는 **유일하게 내리는 근거**다(마무리·즉시). 그 국면의 규칙이 이미
    '새 작업 금지 · 확인과 커밋만'이라 값비싼 판단을 할 일 자체가 없다.
    ★ 단 **요청문이 원인·설계면 안 내린다.** 사람이 새 지시를 내렸는데 몰래 값을
      낮추면 그건 아끼는 것이 아니라 **오판을 싸게 사는 것**이다.

★ **남의 계기를 내 것으로 읽지 않는다.** `컨텍스트_사용량.json` 은 세션마다 덮어쓰는
  **한 장**이라, 옆 창이 방금 쓴 70% 를 내 108% 로 착각할 수 있다. 그래서 그 파일의
  `session` 이 내 세션과 다르면 **'모름'** 으로 두고 내리지 않는다([169]).

★ **값(모델·노력)을 여기 적지 않는다.** 낱말의 정본은 `ai_tier.TIERS` 한 곳이다([230]).
  여기가 정하는 것은 **이번 턴이 어느 갈래인가**뿐이다. 사본을 두면 언젠가 갈리고,
  갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다(loop_policy 가 같은 이유로 그렇게 한다).

★ **없는 손잡이를 지어내지 않는다.** 이 모듈은 **고르고 말할 뿐 적용하지 않는다** —
  돌고 있는 대화의 모델을 바꾸는 길은 확인된 것이 없다. 그래서 사람이 한 번 누를
  명령(`/model`·`/effort`)을 같이 적어 준다. 없는 깃발을 붙이면 조용히 안 된다([169]).

★ **바뀐 순간에만 말한다.** 매 턴 같은 줄을 찍으면 아무도 안 본다([170]).

쓰는 법
    python session_tier.py --print                  # 지금 무엇을 고르나
    python session_tier.py --prompt "왜 안돼"        # 이 요청이면 무엇을 고르나
    python session_tier.py --hook                   # 훅에서(표준입력 JSON)
검증 `[241]`.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):          # pythonw 에서는 None 이다([235])
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

try:                                            # 워크트리에서도 본체 상태를 본다([125])
    from worktree_state import shared
except Exception:                               # pragma: no cover - 단독 실행 보호
    def shared(*parts):
        return os.path.join(ROOT, *parts)

CTX = shared("reports", "컨텍스트_사용량.json")
OUT = shared("reports", "세션_모델선택.json")

#: 갈래의 세기. 뒤로 갈수록 세다. **낱말은 ai_tier.TIERS 의 열쇠와 같아야 한다.**
ORDER = ["조회", "재시도", "원인", "설계"]
DEFAULT = "재시도"                              # 모르는 모양은 낮추지 않는다
#: 컨텍스트가 이 단계면 마무리 국면이다(context_guard 가 쓰는 낱말 그대로).
WRAPUP_STAGES = ("마무리", "즉시")

#: 요청문 → 갈래. **여러 개가 맞으면 센 쪽이 이긴다**(한 문장에 '확인'과 '알고리즘'이
#: 같이 있으면 그것은 설계 지시다 — 확인은 그 안에 딸린 절차일 뿐이다).
INTENT_RULES = [
    ("설계", re.compile(r"알고리즘|코딩|구현|설계|구조|고도화|리팩|만들어|구성해|"
                        r"추가해|바꿔|고쳐|개선")),
    ("원인", re.compile(r"왜\b|왜\s|안\s*되|안돼|안됨|실패|오류|에러|이상해|"
                        r"틀렸|틀린|못\s*찾|멈춰|멈춤|죽었|죽는|누락|헛")),
    ("재시도", re.compile(r"다시|이어서|하던|계속|마저|커밋|푸시|재시도|진행해")),
    ("조회", re.compile(r"확인|알려|보여|무엇|뭐야|뭐지|몇\s*건|몇\s*개|상태|목록|"
                        r"현황|어디까지|정리해")),
]


def _read_json(path):
    """(읽었나, 값). **못 읽은 것과 빈 것을 가른다**([169])."""
    try:
        with io.open(path, encoding="utf-8") as f:
            return True, json.load(f)
    except (OSError, ValueError):
        return False, {}


def value_of(tier):
    """(모델, 노력, 왜) — 낱말의 정본은 ai_tier.TIERS 하나다([230])."""
    import ai_tier
    model, effort, why = ai_tier.TIERS[tier]
    return model, effort, why


def strongest(tiers):
    """센 쪽이 이긴다. 빈 목록이면 None."""
    got = [t for t in tiers if t in ORDER]
    return max(got, key=ORDER.index) if got else None


def intent_of(prompt):
    """요청문 → (갈래, 걸린 낱말들). 못 읽으면 (None, [])."""
    text = str(prompt or "")
    if not text.strip():
        return None, []
    hits, words = [], []
    for tier, rx in INTENT_RULES:
        m = rx.search(text)
        if m:
            hits.append(tier)
            words.append("%s:%s" % (tier, m.group(0).strip()))
    return strongest(hits), words


def disk_floor():
    """디스크 근거 → **바닥만** 올린다. (갈래 or None, 한 줄 근거).

    ★ 천장은 안 올린다 — 세워 둔 일이 있다는 것이 '이 질문이 어렵다'는 뜻은 아니다.
    ★ 못 읽으면 낮추지 않는다: loop_policy 자신이 '보통'으로 답한다([231]).
    """
    try:
        import loop_policy as lp
        ok_h, h = _read_text(lp.HANDOFF)
        ok_r, r = _read_text(lp.ROUNDS)
        d = lp.decide(lp.handoff_items(h) if ok_h else [], ok_h,
                      lp.round_alerts(r) if ok_r else 0)
    except Exception as e:                      # 근거를 못 세우면 바닥을 안 올린다
        return None, "디스크 근거 못 읽음(%s)" % type(e).__name__
    if d.get("갈래") == "가벼움":
        return None, d.get("근거") or ""
    return "재시도", d.get("근거") or ""        # 보통·무거움 둘 다 '조용하지 않다'


def _read_text(path):
    try:
        with io.open(path, encoding="utf-8") as f:
            return True, f.read()
    except OSError:
        return False, ""


def context_stage(session_id=""):
    """(단계, 퍼센트) — **내 세션의 계기일 때만** 값을 준다.

    그 파일은 세션마다 덮어쓰는 한 장이라, 옆 창이 방금 쓴 수치를 내 것으로 읽으면
    **남의 70% 를 보고 내 108% 를 모른다.** 다르면 '모름'이다([169]).
    """
    ok, d = _read_json(CTX)
    if not ok:
        return None, None
    mine = str(session_id or os.environ.get("CLAUDE_CODE_SESSION_ID") or "")
    if not mine or str(d.get("session") or "") != mine:
        return None, None                       # 남의 계기 — 모름
    return d.get("stage") or None, d.get("percent")


def decide(prompt="", session_id="", floor=None, floor_why="", stage=None, percent=None):
    """순수 함수 — 검증이 값을 넣어 부를 수 있어야 한다.

    `floor`·`stage` 를 주면 그것을 쓰고, 안 주면 부르는 쪽에서 채운다(pick 참고).
    """
    tier, words = intent_of(prompt)
    why = []
    if tier:
        why.append("요청문(%s)" % ", ".join(words[:3]))
    else:
        tier = DEFAULT
        why.append("요청문을 못 읽었다 — 모르는 모양은 낮추지 않는다")

    if floor and ORDER.index(floor) > ORDER.index(tier):
        tier = floor
        why.append("디스크 근거로 바닥을 올림: %s" % (floor_why or ""))
    elif floor_why:
        why.append(floor_why)

    내림 = None
    if stage in WRAPUP_STAGES:
        if tier in ("조회", "재시도"):
            tier = "조회"
            내림 = "컨텍스트 %s 국면 — 확인·커밋만 하는 자리다" % stage
        else:
            내림 = ("컨텍스트 %s 국면이지만 요청이 원인·설계라 값을 안 낮춘다 "
                    "— 새 지시를 싸게 판단하면 오판을 싸게 사는 것이다" % stage)
        why.append(내림)

    model, effort, tier_why = value_of(tier)
    return {"갈래": tier, "모델": model, "노력": effort,
            "근거": " · ".join(x for x in why if x),
            "값근거": tier_why, "컨텍스트단계": stage, "컨텍스트퍼센트": percent,
            "적용": "/model %s · /effort %s" % (model, effort)}


def pick(prompt="", session_id=""):
    """지금 이 턴의 선택. 디스크·컨텍스트까지 실제로 읽는다."""
    floor, floor_why = disk_floor()
    stage, percent = context_stage(session_id)
    return decide(prompt, session_id, floor, floor_why, stage, percent)


def _remember(got, session_id=""):
    """(바뀌었나, 이전갈래). **자국을 남기려다 판단을 막지 않는다.**

    갈래별 횟수를 같이 센다 — 나중에 "얼마나 아꼈나"를 물을 수 있어야 규칙이 자란다.
    """
    ok, st = _read_json(OUT)
    st = st if isinstance(st, dict) else {}
    prev = st.get("갈래") if str(st.get("session") or "") == str(session_id or "") else None
    tally = st.get("갈래별") if isinstance(st.get("갈래별"), dict) else {}
    tally[got["갈래"]] = int(tally.get(got["갈래"]) or 0) + 1
    st.update(got)
    st["session"] = str(session_id or "")
    st["갈래별"] = tally
    st["잰시각"] = datetime.now().astimezone().isoformat(timespec="seconds")
    try:
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with io.open(OUT, "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=1)
    except OSError:
        pass
    return (prev != got["갈래"]), prev


def message(got, prev=None):
    """사람과 AI 가 같이 읽는 한 줄. **바뀐 순간에만** 부르는 쪽이 쓴다."""
    head = "[모델 선택] 이번 갈래 **%s** → `%s` / 노력 `%s`" % (
        got["갈래"], got["모델"], got["노력"])
    if prev:
        head += "  (직전 %s)" % prev
    return ("%s\n  근거: %s\n  값근거: %s\n  적용(한 번만 누르면 된다): %s\n"
            "  ※ 고르기만 한다 — 돌고 있는 대화의 모델을 코드가 바꾸는 길은 없다." % (
                head, got["근거"], got["값근거"], got["적용"]))


def hook():
    """UserPromptSubmit 훅 진입점. **무슨 일이 있어도 exit 0** — 사람 입력을 막지 않는다."""
    payload = {}
    try:
        raw = sys.stdin.read()
        if raw.strip():
            payload = json.loads(raw)
    except Exception:
        payload = {}
    try:
        sid = payload.get("session_id") or ""
        got = pick(payload.get("prompt") or "", sid)
        changed, prev = _remember(got, sid)
        msg = message(got, prev) if changed else ""
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


def main(argv=None):
    ap = argparse.ArgumentParser(description="이 대화의 모델·노력을 근거로 고른다")
    ap.add_argument("--hook", action="store_true", help="훅에서 호출(표준입력 JSON)")
    ap.add_argument("--prompt", default="", help="이 요청문이면 무엇을 고르나")
    ap.add_argument("--print", dest="show", action="store_true", help="지금 선택 한 줄")
    ap.add_argument("--json", action="store_true", help="JSON 으로")
    a = ap.parse_args(argv)
    if a.hook:
        return hook()
    got = pick(a.prompt, os.environ.get("CLAUDE_CODE_SESSION_ID") or "")
    if a.json:
        print(json.dumps(got, ensure_ascii=False))
    else:
        print(message(got))
    return 0


if __name__ == "__main__":
    sys.exit(main())
