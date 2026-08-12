# -*- coding: utf-8 -*-
"""AI 를 부를 때 **모델과 노력을 스스로 고른다** (2026-08-12 지시).

사용자 지시: "자동으로 모델과 노력 강도 설정해서 진행하는 알고리즘 적용".

★ **지금까지는 고르지 않았다.** `agent_dispatch._agent_command` 가
  `claude -p <프롬프트>` 만 불렀으므로 **모든 인계가 기본 모델(=제일 비싼 것)** 로 돌았다.
  자율복구가 세 번 실패해 넘기는 재시도 한 장이나, 원인을 모르는 회차 고장이나
  같은 값을 치렀다.

★ **싸게 하는 것이 목적이 아니라 '값에 맞게' 하는 것이 목적이다.** 이 프로젝트에서
  잘못된 판단은 못 한 판단보다 나쁘다(`[172]`·`[196]`). 그래서 **판단이 필요한 일은
  값을 아끼지 않고**, 답이 이미 정해진 일에만 낮춘다. 가르는 근거는 셋뿐이다:
    ① **읽기만 하는가**(`--check`·`--print`·`--status`·`--plan` 같은 조회) → 싸게
    ② **같은 명령을 다시 돌리는 것인가**(자원·인증이 풀려 재시도) → 중간
    ③ **왜 실패했는지 모르는가**(코드 오류·시간 초과가 반복) → **제일 좋은 것으로**
  ③ 이 이 프로젝트에서 제일 비싼 실수가 나는 자리다 — 원인을 모르는 채 값싼 판단을
  받으면 그럴듯한 오진이 파일에 박히고, 그것이 다음 사람의 출발점이 된다.

★ **`haiku` 는 쓰지 않는다.** 싸지만 이 프로젝트는 지시문이 길고 판정이 미묘하다
  ("0건이 없는 건가 안 본 건가", "밀림인가 실패인가"). 감시·판정 층에 값싼 오판을
  넣으면 아끼는 것보다 잃는 것이 크다. 필요해지면 여기 한 줄로 넣는다.

★ **없는 손잡이를 지어내지 않는다.** `claude` CLI 에는 `--model` 이 있지만 '노력'을
  주는 깃발은 확인된 것이 없다. 그래서 **모델은 깃발로, 노력은 프롬프트 문장으로**
  넘기고 티켓에 둘 다 적어 둔다. 있지도 않은 깃발을 붙이면 CLI 가 통째로 안 뜨고,
  그러면 인계가 **조용히** 안 된다(`[169]`).

쓰는 법
    python ai_tier.py --kind code --title "회차가 죽는다"     # 무엇을 고르는지 본다
검증 `[230]`.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):              # 콘솔이 cp949 라도 안 깨진다
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

#: 갈래 → (모델, 노력, 왜). 낱말은 CLI 가 실제로 받는 것만 쓴다.
TIERS = {
    "조회":  ("sonnet", "low",
             "읽기만 하는 일이다 — 답이 파일에 이미 있고 판단할 것이 없다"),
    "재시도": ("sonnet", "medium",
             "같은 명령을 다시 돌리는 일이다 — 무엇을 할지는 이미 정해져 있다"),
    "원인":  ("opus", "high",
             "왜 실패했는지 모른다 — 값싼 오진이 파일에 박히면 다음 사람의 출발점이 된다"),
    "설계":  ("opus", "high",
             "구조를 정하는 일이다 — 되돌리기 어렵고 사본이 생기면 갈린다"),
}
DEFAULT = "재시도"

#: 읽기만 하는 명령의 표식. 이 목록에 없는 것을 조회라고 우기지 않는다.
READONLY = ("--check", "--print", "--status", "--plan", "--dry", "--stats", "--whoami")
#: 원인을 모르는 실패의 표식(autopilot 이 붙이는 갈래).
UNKNOWN_CAUSE = ("code", "timeout")


def classify(kind="", title="", args=None, attempts=0):
    """무슨 갈래인가 — **근거가 있을 때만** 위아래로 움직인다."""
    argv = [str(a) for a in (args or [])]
    text = "%s %s" % (title or "", " ".join(argv))
    if str(kind) in ("설계", "design"):
        return "설계"
    # ★ 원인을 모르는 반복 실패가 먼저다. 명령에 `--check` 가 섞여 있어도, 그것이
    #   세 번 연속 죽었다면 물어야 할 것은 '조회 결과'가 아니라 '왜 죽나'다.
    if str(kind) in UNKNOWN_CAUSE and int(attempts or 0) >= 2:
        return "원인"
    if any(f in argv for f in READONLY) and not re.search(r"--(apply|queue|send|post|upload|delete)\b", text):
        return "조회"
    return DEFAULT


def pick(kind="", title="", args=None, attempts=0):
    """(갈래, 모델, 노력, 왜) — 고른 이유를 같이 돌려준다.

    ★ **왜 고른지를 남기지 않으면 아무도 이 선택을 못 고친다.** 나중에 "왜 이건
      비싼 모델로 돌았나"를 물을 수 있어야 규칙이 자란다.
    """
    tier = classify(kind, title, args, attempts)
    model, effort, why = TIERS[tier]
    return {"갈래": tier, "모델": model, "노력": effort, "왜": why}


#: 깃발이 실재하는지 물어본 결과를 적어 두는 자리(CLI 판이 바뀌면 다시 묻는다).
FLAG_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "reports", ".ai_cli_flags.json")


def supports_flag(executable, flag):
    """그 CLI 가 **정말** 그 깃발을 받는가 — `--help` 에게 직접 묻는다.

    ★ 이 함수가 이 파일의 급소다. 없는 깃발을 하나 붙이면 CLI 가 **통째로 안 뜨고**,
      그러면 인계가 오류도 없이 **조용히** 안 된다(`[169]`). 그래서 규칙은 셋이다:
        ① 물어보고 **있다고 확인된 것만** 붙인다
        ② **못 물어봤으면 안 붙인다** — '아마 있겠지'는 근거가 아니다
        ③ CLI 가 판올림·다운그레이드되면 **다시 묻는다**(열쇠에 mtime·크기를 넣는다)
    ★ 2026-08-12 에는 `--effort` 가 확인되지 않아 일부러 문장으로 넘겼다. 2026-08-13
      실측(`claude 2.1.222 --help`)으로 `--effort <low|medium|high|xhigh|max>` 가
      실재함을 확인해 깃발로 옮겼다 — **확인 없이 옮기지 않았다는 것이 요점이다.**
    """
    if not executable or not flag:
        return False
    try:
        st = os.stat(executable)
        key = "%s|%d|%d|%s" % (executable, st.st_size, int(st.st_mtime), flag)
    except OSError:
        return False                                  # 실행파일을 못 보면 안 붙인다
    try:
        cache = json.load(open(FLAG_CACHE, encoding="utf-8"))
    except Exception:
        cache = {}
    if key in cache:
        return bool(cache[key])
    ok = False
    try:                                              # [175] — run 대신 나무째 관리
        from proc_guard import run_tree
        res = run_tree([executable, "--help"], timeout=30, drain_timeout=10,
                       output_limit=200_000)
        text = (res.stdout or "") + (res.stderr or "")
        ok = (not res.timed_out) and (flag in text)
    except Exception:
        ok = False                                    # 못 물어봤다 = 안 붙인다
    cache[key] = ok
    try:
        os.makedirs(os.path.dirname(FLAG_CACHE), exist_ok=True)
        tmp = FLAG_CACHE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as out:
            json.dump(cache, out, ensure_ascii=False)
        os.replace(tmp, FLAG_CACHE)
    except Exception:
        pass                                          # 기억 못 해도 판단은 옳다
    return ok


def flags(agent, chosen, executable=None):
    """CLI 에 실제로 붙일 깃발. **확인된 것만 붙인다.**

    ★ `claude` 는 `--model` 을 받는다(오래 확인된 것이라 그대로 붙인다).
      `--effort` 는 **그 실행파일에게 물어본 뒤에만** 붙인다 — `executable` 을 안 주면
      물어볼 수 없으므로 안 붙이고, 노력은 예전처럼 문장으로 간다.
    ★ `codex` 쪽은 확인된 것이 없어 아무것도 안 붙인다.
    """
    if agent != "claude":
        return []
    out = []
    if chosen.get("모델"):
        out += ["--model", str(chosen["모델"])]
    if chosen.get("노력") and supports_flag(executable, "--effort"):
        out += ["--effort", str(chosen["노력"])]
    return out


def effort_is_flagged(agent, chosen, executable=None):
    """노력이 **깃발로** 갔나 — 갔으면 문장이 같은 말을 반복하지 않는다."""
    return "--effort" in flags(agent, chosen, executable)


def prompt_line(chosen, effort_via_flag=False):
    """AI 에게 **어떻게 일할지**를 한 줄로 준다.

    ★ 노력 수치는 깃발이 받았으면 여기서 다시 말하지 않는다 — 같은 값을 두 곳에서
      말하면 언젠가 갈리고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다(`[162]`).
      깃발을 못 붙였을 때만 문장이 그 몫을 대신한다.
    """
    tag = "" if effort_via_flag else ("[노력 %s] " % chosen.get("노력", ""))
    if chosen.get("갈래") in ("원인", "설계"):
        return (tag + "원인을 단정하기 전에 근거를 파일에서 확인하세요. "
                "짐작으로 고치지 말고, 못 밝히면 못 밝혔다고 적으세요.")
    return (tag + "이미 정해진 일입니다. 새 판단을 만들지 말고 그대로 수행한 뒤 "
            "결과만 적으세요.")


# ───────────────────────────────── 사람이 친 말도 등급을 매긴다 (2026-08-13 지시)
#
# 사용자 지시: "이 세션에서 크래딧 최대한 아끼고 최고의 퍼포먼스를 낼 수 있는 모델
# 자동 선택하고 추론 모드도 자동으로 선택해서 적용해서 진행하는 알고리즘 구현해".
#
# ★ **정직하게 적어 두는 한계**: 대화창의 모델·추론 모드는 세션 **안에서 못 바꾼다**.
#   그건 앱이 정한다. 그래서 이 규칙이 실제로 적용되는 자리는 셋이다 —
#     ① `agent_dispatch` 가 AI CLI 를 부를 때(깃발로 자동 적용, `[230]`)
#     ② 이 세션이 **부하 에이전트**를 띄울 때(Agent 도구의 model·effort 로 자동 적용)
#     ③ 사람에게 **한 줄로 알려 주기**(UserPromptSubmit 훅)
#   ③ 을 '자동 적용'이라 적으면 거짓말이 된다 — 알려 주는 것까지가 여기 몫이다.
#
# ★ **아끼는 쪽으로 기울지 않는다.** 이 프로젝트에서 잘못된 판단은 못 한 판단보다
#   나쁘다(`[172]`). 그래서 갈리지 않으면 **모델은 안 내리고 노력만 낮춘다** —
#   싸게 틀리는 것이 비싸게 맞는 것보다 훨씬 비싸다.
PROMPT_TIERS = {
    "질문":  ("sonnet", "low",
             "답이 파일·화면에 이미 있는 조회다 — 새로 판단할 것이 없다"),
    "수행":  ("sonnet", "medium",
             "무엇을 할지 이미 정해진 일이다 — 실행하고 결과만 적으면 된다"),
    "설계":  ("opus", "high",
             "원인·구조를 정하는 일이다 — 값싼 오진이 파일에 박히면 다음 사람의 출발점이 된다"),
    "모호":  ("opus", "medium",
             "규칙이 못 갈랐다 — 판단이 섞일 수 있으니 모델은 안 내리고 노력만 낮춘다"),
}
#: 등급을 올리는 말. 하나라도 있으면 설계다 — **의심되면 위로 간다.**
_UP = ("왜", "원인", "문제", "고쳐", "고쳐줘", "안돼", "안 돼", "안되", "이상",
       "알고리즘", "구현", "설계", "구성", "고도화", "자동화", "분석", "최적화",
       "사고", "리팩", "성능", "왜그", "정리해서", "재발")
#: 답이 이미 있는 조회.
_ASK = ("확인", "상태", "보여", "알려", "뭐야", "무엇", "어디", "얼마", "몇",
        "목록", "조회", "언제", "남았", "됐어", "됐나", "현황")
#: 무엇을 할지 정해진 수행.
_DO = ("실행", "돌려", "커밋", "푸시", "반영", "복사", "지워", "삭제", "설치",
       "등록", "올려", "내려", "이어서", "진행해", "계속")


def pick_for_prompt(text):
    """사람이 친 말 한 줄의 등급. **못 가르면 위로 간다**(모델은 안 내린다)."""
    t = str(text or "").strip()
    if not t:
        return dict(zip(("모델", "노력", "왜"), PROMPT_TIERS["모호"]), 갈래="모호")
    hit = lambda ws: any(w in t for w in ws)                      # noqa: E731
    # ★ 순서가 뜻을 갖는다 — "왜 안 되는지 확인해줘" 는 조회가 아니라 원인 규명이다.
    tier = "설계" if hit(_UP) else "수행" if hit(_DO) else "질문" if hit(_ASK) else "모호"
    model, effort, why = PROMPT_TIERS[tier]
    return {"갈래": tier, "모델": model, "노력": effort, "왜": why}


LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "reports", "AI_티어_기록.json")


def _remember(tier):
    """직전 등급과 누적 횟수. **바뀐 순간에만 말하려고** 기억한다 —
    매 입력마다 같은 말을 하면 아무도 안 읽는다(`context_guard` 와 같은 규칙).
    누적은 나중에 규칙을 고칠 근거다 — 어느 갈래가 실제로 많이 오는지 기록이 말한다."""
    path, prev, cnt = LOG, "", {}
    try:
        with open(path, encoding="utf-8") as fh:
            d = json.load(fh)
        prev, cnt = str(d.get("직전") or ""), dict(d.get("누적") or {})
    except Exception:
        pass
    cnt[tier] = int(cnt.get(tier) or 0) + 1
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump({"직전": tier, "누적": cnt}, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception:
        pass
    return prev


def hook():
    """UserPromptSubmit 훅. **무슨 일이 있어도 exit 0** — 사람 입력을 막지 않는다."""
    payload = {}
    try:
        raw = sys.stdin.read()
        if raw.strip():
            payload = json.loads(raw)
    except Exception:
        payload = {}
    try:
        got = pick_for_prompt(payload.get("prompt") or "")
        prev = _remember(got["갈래"])
        if prev == got["갈래"]:
            return 0                              # 안 바뀌었으면 조용히 있는다
        msg = ("[모델 판정] 이번 요청은 **%s** — `%s` / 노력 `%s` 가 맞습니다 (%s).\n"
               "대화창 모델은 세션 안에서 못 바꿉니다(앱에서 고르세요). "
               "부하 에이전트를 띄울 때는 이 등급을 그대로 적용하세요."
               % (got["갈래"], got["모델"], got["노력"], got["왜"]))
        print(json.dumps({"hookSpecificOutput": {
            "hookEventName": "UserPromptSubmit", "additionalContext": msg}},
            ensure_ascii=False))
    except Exception:
        pass
    return 0


def main(argv=None):
    if "--hook" in (sys.argv[1:] if argv is None else argv):
        return hook()
    ap = argparse.ArgumentParser()
    ap.add_argument("--hook", action="store_true", help="훅에서 호출(표준입력 JSON)")
    ap.add_argument("--say", default="", help="사람 말 한 줄의 등급을 본다")
    ap.add_argument("--kind", default="")
    ap.add_argument("--title", default="")
    ap.add_argument("--attempts", type=int, default=0)
    ap.add_argument("--args", nargs="*", default=[])
    a = ap.parse_args(argv)
    got = pick_for_prompt(a.say) if a.say else pick(a.kind, a.title, a.args, a.attempts)
    print(json.dumps(got, ensure_ascii=False))
    print("claude 깃발:", " ".join(flags("claude", got)) or "(없음)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
