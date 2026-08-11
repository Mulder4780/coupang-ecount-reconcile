#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""루프 정책 — 이번 틱을 **어느 모델·어느 노력 강도**로 할지 근거로 정한다.

사용자 지시(2026-08-12): "루프는 하루에 한번 매일 12시에서 13시로 설정하고
자동으로 모델과 노력 강도 설정해서 진행하는 알고리즘 적용"

왜 파일인가
-----------
"이번엔 세게, 저번엔 가볍게" 를 대화에서 정하면 세션이 바뀌는 순간 사라진다.
그러면 다음 세션은 **아무것도 없는데 큰 모델로** 돌거나, **사고가 났는데 작은
모델로** 돈다. 둘 다 조용한 낭비다(전자는 크레딧, 후자는 놓친 사고).
그래서 판단을 파일 하나에 두고 회차·루프가 그것을 읽는다.

무엇을 보는가 (전부 이미 디스크에 있는 것 — 비싼 탐색을 새로 하지 않는다, [168])
  · reports/세션인계.md          '먼저 처리할 것' 항목들 (워치독이 30분마다 갱신)
  · reports/스케줄러_회차감시.md  회차 경보 수 ([228])

어떻게 가르는가 — **낱말이 아니라 '누가 할 일인가'로** 가른다.
  · 세워 둔 일([N]) · 회차 실패/강제종료  → 코드를 고쳐야 끝난다      → 무거움
  · 큐·미푸시·점유·문서                    → 확인하고 정리하면 끝난다  → 보통
  · 수집 밀림                              → 수집 세션 몫이라 내가 할 일이 아니다 → 가벼움
  · 아무것도 없음                          → 확인만 하고 끝            → 가벼움

★ **못 읽은 것을 '없음'으로 치지 않는다**([169]). 근거 파일을 못 읽으면 '모름'
  으로 두고 **보통**으로 간다 — 못 읽었다는 이유로 노력을 낮추면, 파일이 깨진
  날 사고가 가장 작은 모델을 만난다.

★ **모델 이름을 여기 적지 않는다.** 값은 `ai_tier.TIERS`([230]) 한 곳에서 온다.
  이 파일이 정하는 것은 **이번 틱이 어느 갈래인가**뿐이고, 그 갈래를 무슨 모델·
  노력으로 할지는 거기가 답한다. 같은 표를 두 곳에 적으면 언젠가 갈리고, 갈린
  뒤에는 어느 쪽이 맞는지 아무도 모른다. (첫 판에 그 표를 여기에도 적었다가
  `haiku` 를 넣었는데, ai_tier 는 **일부러 haiku 를 뺐다** — 값싼 오판이 아끼는
  것보다 크다는 이유였다. 사본을 두면 이런 이유가 조용히 사라진다.)
"""
from __future__ import annotations

import io
import json
import os
import re
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))

try:                                            # 워크트리에서도 본체 상태를 본다([125])
    from worktree_state import shared
except Exception:                               # pragma: no cover - 단독 실행 보호
    def shared(*parts):
        return os.path.join(ROOT, *parts)

HANDOFF = shared("reports", "세션인계.md")
ROUNDS = shared("reports", "스케줄러_회차감시.md")
OUT = shared("reports", "루프_정책.json")

# 틱 갈래 → ai_tier 갈래. **값(모델·노력)은 ai_tier 가 답한다.**
TIER_OF = {"가벼움": "조회", "보통": "재시도", "무거움": "원인"}
ORDER = ["가벼움", "보통", "무거움"]            # 센 쪽이 이긴다


def value_of(tier):
    """(모델, 노력, 왜) — 낱말의 정본은 ai_tier.TIERS 하나다([230])."""
    import ai_tier
    model, effort, why = ai_tier.TIERS[TIER_OF[tier]]
    return model, effort, why

# 항목 한 줄 → 갈래. 위에서부터 먼저 맞는 것.
RULES = [
    ("무거움", re.compile(r"세워 둔 일|회차 \[(?:실패|강제종료|멈춤)\]|합성검증|빨강|되돌아갔")),
    ("가벼움", re.compile(r"수집이 밀렸다|로그인")),          # 수집 세션·사람 몫
    ("보통",   re.compile(r"입력 큐|미푸시|점유|인계|정본|사본|tmp")),
]


def _read(path):
    """(읽었나, 내용). 못 읽은 것과 빈 것을 가른다."""
    try:
        with io.open(path, encoding="utf-8") as f:
            return True, f.read()
    except OSError:
        return False, ""


def handoff_items(text):
    """'먼저 처리할 것' 아래 굵은 줄들만 뽑는다."""
    if "먼저 처리할 것" not in text:
        return []
    body = text.split("먼저 처리할 것", 1)[1]
    body = re.split(r"\n## ", body, 1)[0]
    return [m.strip() for m in re.findall(r"^- \*\*(.+?)\*\*\s*$", body, re.M)]


def round_alerts(text):
    m = re.search(r"^## 경보 \((\d+)\)", text, re.M)
    return int(m.group(1)) if m else 0


def classify(item):
    for tier, rx in RULES:
        if rx.search(item):
            return tier
    return "보통"                                # 모르는 모양은 낮추지 않는다


def decide(items, read_ok, alerts=0):
    """순수 함수 — 검증이 값을 넣어 부를 수 있어야 한다."""
    if not read_ok:
        model, effort, _ = value_of("보통")
        return {"갈래": "보통", "모델": model, "노력": effort,
                "근거": "근거 파일을 못 읽었다 — 못 읽은 것을 '없음'으로 치지 않는다",
                "항목수": None, "경보수": None, "항목갈래": {}}

    per = {}
    tier = "가벼움"
    for it in items:
        t = classify(it)
        per[it] = t
        if ORDER.index(t) > ORDER.index(tier):
            tier = t
    if alerts and ORDER.index("무거움") > ORDER.index(tier):
        tier = "무거움"                          # 회차 경보는 그 자체로 코드 일이다

    if not items and not alerts:
        why = "먼저 처리할 것도 회차 경보도 없다 — 확인만 하고 끝낸다"
    else:
        heavy = [i for i, t in per.items() if t == "무거움"]
        why = (f"무거운 항목 {len(heavy)}건" if heavy else f"항목 {len(items)}건")
        if alerts:
            why += f" · 회차 경보 {alerts}건"
    model, effort, tier_why = value_of(tier)
    return {"갈래": tier, "모델": model, "노력": effort, "근거": why,
            "값근거": tier_why, "ai_tier갈래": TIER_OF[tier],
            "항목수": len(items), "경보수": alerts, "항목갈래": per}


def build():
    ok1, h = _read(HANDOFF)
    ok2, r = _read(ROUNDS)
    out = decide(handoff_items(h) if ok1 else [], ok1, round_alerts(r) if ok2 else 0)
    out["잰시각"] = datetime.now().astimezone().isoformat(timespec="seconds")
    out["근거파일"] = {"세션인계": ok1, "회차감시": ok2}
    try:
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with io.open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
    except OSError:
        pass                                     # 자국을 남기려다 판단을 막지 않는다
    return out


def main(argv):
    d = build()
    print(f"갈래 {d['갈래']} · 모델 {d['모델']} · 노력 {d['노력']}")
    print(f"  근거: {d['근거']}")
    if "--print" in argv or "-v" in argv:
        for it, t in list(d["항목갈래"].items())[:12]:
            print(f"  [{t}] {it[:70]}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
