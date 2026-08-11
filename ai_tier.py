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


def flags(agent, chosen):
    """CLI 에 실제로 붙일 깃발. **확인된 것만 붙인다.**

    ★ `claude` 는 `--model` 을 받는다. `codex` 쪽은 확인된 것이 없어 안 붙인다 —
      틀린 깃발 하나면 CLI 가 통째로 안 뜨고 인계가 **조용히** 안 된다.
    """
    if agent == "claude" and chosen.get("모델"):
        return ["--model", str(chosen["모델"])]
    return []


def prompt_line(chosen):
    """노력은 깃발이 없으므로 **문장으로** 넘긴다 — 있지도 않은 깃발을 지어내지 않는다."""
    if chosen.get("갈래") in ("원인", "설계"):
        return ("[노력 %s] 원인을 단정하기 전에 근거를 파일에서 확인하세요. "
                "짐작으로 고치지 말고, 못 밝히면 못 밝혔다고 적으세요." % chosen["노력"])
    return ("[노력 %s] 이미 정해진 일입니다. 새 판단을 만들지 말고 그대로 수행한 뒤 "
            "결과만 적으세요." % chosen["노력"])


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--kind", default="")
    ap.add_argument("--title", default="")
    ap.add_argument("--attempts", type=int, default=0)
    ap.add_argument("--args", nargs="*", default=[])
    a = ap.parse_args(argv)
    got = pick(a.kind, a.title, a.args, a.attempts)
    print(json.dumps(got, ensure_ascii=False))
    print("claude 깃발:", " ".join(flags("claude", got)) or "(없음)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
