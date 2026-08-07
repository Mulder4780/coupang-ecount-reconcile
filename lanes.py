# -*- coding: utf-8 -*-
"""
lanes.py — 세션마다 **차선**을 정해 둔다 (2026-08-07 지시)

사용자 지시: "이 세션은 자료 수집만 계속해서 'CSOS 앱 코딩(+엑셀)' 이 세션에서
코딩 작업 디자인 작업 등 관련 모든 작업 및 엑셀 보완 반영 작업을 계속 할 수 있게
하는 알고리즘 추가해서 계속 자료 수집해."

왜 `ai_claim` 만으로는 부족했나
  `ai_claim` 은 **자원 한 개**를 잡는다 — "지금 이 순간 관리대장을 누가 쓰나".
  짧고 사후적이다. 그런데 두 창을 몇 시간씩 나란히 굴리면 필요한 것은 다른 것이다:
  **"이 창은 무엇을 하는 창인가"** 라는 오래 가는 약속.

  약속이 없으면 이렇게 샌다(실측):
   · 수집하던 창이 "겸사겸사" 코드를 고치기 시작한다 → 앱 창과 같은 파일에서 만난다
   · 앱 창이 잠깐 비운 사이 수집 창이 `code` 를 집어 간다 → 앱 창이 되잡지 못한다
   · 사람이 "누가 뭘 하고 있지?" 를 매번 물어야 한다

  차선은 그 약속을 파일에 적어 둔 것이다. `ai_claim` 이 자물쇠라면 차선은 **분업표**다.
  층이 다르므로 서로를 대체하지 않는다:
      lanes      = 이 창은 무슨 일을 하는 창인가 (몇 시간)
      ai_claim   = 지금 이 자원을 누가 쥐고 있나 (몇 분)
      worksplit  = 어떤 할 일을 누가 맡았나 (건별)

무엇이 어느 차선인가
  collect(수집) — 밴드·ERP·카톡·다운로드 흡수·원본 정리. 원본을 **모으기만** 한다.
  build(앱·엑셀) — 앱 코딩·디자인·관리대장 반영·게시. 모인 것을 **쓰는** 쪽이다.
  두 차선은 건드리는 파일이 거의 겹치지 않아서 하루 종일 나란히 돌 수 있다.

★ 차선이 하나도 정해지지 않았으면 **아무것도 막지 않는다.**
  기존 세션·스케줄러가 그대로 돌아야 한다. 차선은 켠 사람에게만 적용된다.

★ 막는 것은 **내 차선 밖의 자원을 내가 잡으려 할 때**뿐이다.
  남의 차선을 대신 비워 주거나 남의 점유를 풀지 않는다 — 그건 `ai_claim` 의 몫이고,
  거기서 이미 "남의 것은 못 놓는다"가 지켜지고 있다.

사용
  python lanes.py                                  # 지금 누가 어느 차선인가
  python lanes.py --take collect --who claude      # 이 세션을 수집 차선에 넣는다
  python lanes.py --take build --who claude --why "CSOS 앱 코딩+엑셀"
  python lanes.py --free                           # 내 차선에서 빠진다
  python lanes.py --can code                       # 내 차선에서 code 를 잡아도 되나
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))

# 점유 파일과 **같은 자리**를 쓴다 — 워크트리에서 갈리면 분업표가 두 장이 된다.
try:
    from worktree_state import shared as _shared
    STATE_DIR = _shared("reports")
except Exception:
    STATE_DIR = os.path.join(ROOT, "reports")

LANES_FILE = os.path.join(STATE_DIR, "작업차선.json")

# 차선 → (이름, 이 차선에서 잡아도 되는 자원)
# `read`·`report` 는 배타가 아니라서 어느 차선에서도 열려 있다(조회·보고는 늘 허용).
LANES = {
    "collect": ("자료 수집", ["band", "read", "report"]),
    "build":   ("앱·엑셀·코드", ["code", "ledger", "publish", "read", "report"]),
}

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _claimlib():
    """세션 식별·생사 판정은 ai_claim 것을 그대로 쓴다 — 두 벌이면 어긋난다."""
    import ai_claim
    return ai_claim


def _path():
    """차선 파일은 **점유 파일 옆**에 산다 — 자리를 ai_claim 에게 물어본다.

    ★ 왜 상수가 아닌가: 합성검증은 점유를 임시 폴더로 돌려 놓고 `take` 를 시험한다.
      차선만 진짜 폴더를 보면, 수집 차선에 선 세션에서 그 시험이 통째로 실패한다
      (실측). 그러면 **"ALL GREEN 확인 후 실작업" 관문 자체를 통과할 수 없다** —
      차선을 켠 값으로 검증을 못 돌리는 것은 기능이 아니라 사고다.
      자리를 따라가게 하면 시험은 빈 차선을 보고, 운영에서는 둘 다 reports/ 다.
    """
    try:
        import ai_claim
        return os.path.join(os.path.dirname(ai_claim.CLAIMS), "작업차선.json")
    except Exception:
        return LANES_FILE


def _load():
    try:
        return json.load(open(_path(), encoding="utf-8")) or {}
    except Exception:
        return {}


def _save(d):
    p = _path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, p)                   # 원자적 교체(사고 24)


def _dead(rec):
    """주인 세션이 죽었나 — ai_claim 과 **같은 판정**을 쓴다(45분 기다리지 않는다)."""
    try:
        return _claimlib()._is_dead(rec)
    except Exception:
        return False


def _me():
    try:
        return _claimlib().session_id()
    except Exception:
        return ""


def owner(lane, d=None):
    """살아 있는 주인 기록만 돌려준다. 죽은 세션의 차선은 비어 있는 것으로 본다."""
    rec = (d if d is not None else _load()).get(lane)
    if not isinstance(rec, dict) or not rec.get("sid"):
        return None
    return None if _dead(rec) else rec


def my_lane(d=None):
    """이 세션이 서 있는 차선 이름 (없으면 None)."""
    d = d if d is not None else _load()
    me = _me()
    for lane in LANES:
        rec = owner(lane, d)
        if rec and rec.get("sid") == me:
            return lane
    return None


def can(what, d=None):
    """이 세션이 자원 `what` 을 잡아도 되나 → (허용여부, 이유).

    ★ 차선을 안 정했으면 **언제나 허용**이다. 켠 사람에게만 적용되는 규칙이다.
    """
    lane = my_lane(d)
    if lane is None:
        return True, ""
    allowed = LANES[lane][1]
    if what in allowed:
        return True, ""
    # 그 자원이 어느 차선 것인지 일러 주면 사람이 무엇을 해야 할지 안다.
    home = next((n for n, (_l, res) in LANES.items() if what in res and n != lane), None)
    why = (f"이 세션은 '{LANES[lane][0]}' 차선입니다 — '{what}' 은 "
           f"{'‘' + LANES[home][0] + '’ 차선' if home else '다른 차선'}의 자원입니다.")
    return False, why


def show(d=None):
    d = d if d is not None else _load()
    me = _me()
    print(f"작업 차선  (이 세션 = {me or '?'})")
    for lane, (label, res) in LANES.items():
        rec = owner(lane, d)
        if not rec:
            raw = d.get(lane)
            note = " (주인 세션 종료 — 비었음)" if isinstance(raw, dict) and raw.get("sid") else ""
            print(f"  [{lane:7}] {label:12} 비어 있음{note}")
            continue
        mins = int((time.time() - float(rec.get("at") or 0)) / 60)
        mark = " ← 내 것" if rec.get("sid") == me else ""
        print(f"  [{lane:7}] {label:12} {rec.get('who','?')}"
              f" [{rec.get('sid','?')}] {mins}분 전 · {rec.get('why','')[:32]}{mark}")
    lane = my_lane(d)
    if lane:
        print(f"\n이 세션이 잡을 수 있는 자원: {', '.join(LANES[lane][1])}")
    else:
        print("\n이 세션은 차선 밖입니다 — 아무것도 막지 않습니다.")
        print("잡기: python lanes.py --take <collect|build> --who claude")
    return 0


def take(lane, who, why=""):
    if lane not in LANES:
        print(f"✗ 모르는 차선 '{lane}' — {'/'.join(LANES)} 중 하나")
        return 1
    d = _load()
    me = _me()
    cur = owner(lane, d)
    if cur and cur.get("sid") != me:
        mins = int((time.time() - float(cur.get("at") or 0)) / 60)
        print(f"★ '{LANES[lane][0]}' 차선은 {cur.get('who','?')} 세션[{cur.get('sid')}] 이"
              f" 서 있습니다 ({mins}분 전 · {cur.get('why','')}).")
        print("  → 다른 차선을 잡으세요. 남의 차선을 빼앗으면 두 창이 같은 파일에서 만납니다.")
        return 2
    # 한 세션이 두 차선에 서지 않는다 — 그러면 분업표가 아니게 된다.
    prev = my_lane(d)
    if prev and prev != lane:
        d.pop(prev, None)
        print(f"i '{LANES[prev][0]}' 차선에서 빠집니다 (한 세션은 한 차선).")
    cl = _claimlib()
    d[lane] = {"who": who, "why": why, "at": time.time(),
               "when": datetime.now().strftime("%Y-%m-%d %H:%M"),
               "sid": me, "agent_pid": cl.agent_pid(), "host": cl.socket.gethostname()}
    _save(d)
    print(f"'{LANES[lane][0]}' 차선 — {who}[{me}]" + (f" · {why}" if why else ""))
    print(f"  잡을 수 있는 자원: {', '.join(LANES[lane][1])}")
    return 0


def free(who=None):
    """내 차선에서만 빠진다. 남의 차선은 건드리지 않는다."""
    d = _load()
    lane = my_lane(d)
    if not lane:
        print("이 세션은 차선 밖입니다 — 놓을 것이 없습니다.")
        return 0
    d.pop(lane, None)
    _save(d)
    print(f"'{LANES[lane][0]}' 차선에서 빠졌습니다.")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="세션의 작업 차선을 정한다")
    ap.add_argument("--take", metavar="LANE")
    ap.add_argument("--who", default="claude")
    ap.add_argument("--why", default="")
    ap.add_argument("--free", action="store_true")
    ap.add_argument("--can", metavar="RESOURCE")
    a = ap.parse_args(argv)
    if a.can:
        ok, why = can(a.can)
        print(("O " if ok else "X ") + (why or f"'{a.can}' 을 잡을 수 있습니다."))
        return 0 if ok else 3
    if a.free:
        return free(a.who)
    if a.take:
        return take(a.take, a.who, a.why)
    return show()


if __name__ == "__main__":
    sys.exit(main())
