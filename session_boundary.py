# -*- coding: utf-8 -*-
"""아무 때나 compact·clear 해도 이어진다 — **세션 경계** 한 곳 (2026-08-13 지시).

사용자 지시: **"아무때나 세션 컴팩팅이나 클리어 해도 문제 없이 돌아갈 수 있는
알고리즘 구현해"**

★ 빠져 있던 것은 '요약하는 기능'이 아니라 **요약 뒤에 남는 두 가지**였다.

① **`/clear` 에는 인계가 없었다.** `PreCompact` 훅은 compact 때만 온다. `/clear` 는
   대화만 비우고 **프로세스는 안 죽는다** — 그래서 `ai_claim._is_dead` 는 pid 를 보고
   '살아 있다'고 답하고, 그 sid 는 사라졌으니 `--free-all`(내 세션 것만, `[104]`)도
   닿지 않는다. **아무도 못 푸는 점유**가 남고, 그동안 관리대장 동시 쓰기 금지가
   조용히 무너진다. 화면에는 멀쩡히 '점유 중'이라고 뜬다.
② **요약이 담은 파일 사본은 낡는다.** 실측 2026-08-13: 컴팩션 요약에 실려 온
   `ai_tier.py` 가 옛 것(`flags` 2인자)이라, **이미 3인자로 고쳐진 코드**를 "둘이
   어긋나 있다"고 진단하고 도구를 세 번 헛돌렸다. 요약은 **대화**를 줄이는 것이고
   **파일을 다시 읽어 주지는 않는다.** 비어 있으면 알아채지만 낡은 사본은 안 띈다
   (`[165]` 와 같은 모양 — 안 읽은 열은 빈칸과 구별할 수 없다).

그래서 경계마다 이렇게 한다.
  · **끝날 때**(`SessionEnd` · `PreCompact`) → **이미 있는 `session_wrapup.py`** 가
    그대로 받는다. 여기에 마무리 사본을 만들지 않는다 — 같은 판단이 두 곳에 있으면
    언젠가 갈리고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다.
  · **시작할 때**(`SessionStart`) → 이 파일이 상태를 되찾아 **한 장으로 주입**한다:
    고아 점유 회수 · 내 점유·맡은 일 · 먼저 처리할 것 · 이번 세션 권장 등급 ·
    ★ **최근에 바뀐 파일**(요약 속 사본을 믿지 말라는 **근거**).

★ **훅 갈래를 문서에서 짐작하지 않는다.** `SessionStart` 의 `source`
  (startup·clear·compact·resume)와 `SessionEnd` 의 `reason` 은 여기 쌓인
  **실측**(`reports/세션경계_기록.json`)이 말해 준다. 짐작으로 분기를 만들면 그
  분기가 영영 안 타면서도 아무 오류를 안 낸다.

쓰는 법
    python session_boundary.py                 # 지금 무엇을 주입할지 사람이 본다
    python session_boundary.py --start         # SessionStart 훅(표준입력 JSON)
검증 `[241]`.
"""
from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):          # pythonw·cp949 에서도 안 죽는다(`[235]`)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:                        # 훅은 cwd 를 어디로 줄지 모른다
    sys.path.insert(0, ROOT)

try:                                            # 워크트리에서도 본체 상태를 본다(`[125]`)
    from worktree_state import shared
except Exception:                               # pragma: no cover - 단독 실행 보호
    def shared(*parts):
        return os.path.join(ROOT, *parts)

LOG = shared("reports", "세션경계_기록.json")
#: 이 분 안에 바뀐 파일은 "요약 속 사본이 낡았을 수 있다"의 근거가 된다.
FRESH_MIN = 45
#: 잡힌 지 이만큼도 안 된 점유는 고아로 보지 않는다 — 새 창은 트랜스크립트가 아직 없다.
MIN_AGE_MIN = 15
KEEP = 12                                       # 실측 기록은 최근 이만큼만
#: 훑을 자리. Z: 는 보지 않는다(`[168]`) — 비싼 탐색을 훅에 넣지 않는다.
WATCH_DIRS = ("", "band", "webapp", "tests")


def _read_json(path, default):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return default


def _write_json(path, value):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(value, fh, ensure_ascii=False, indent=1)
        os.replace(tmp, path)
    except Exception:                           # 기록을 남기려다 훅을 막지 않는다
        pass


def note(payload):
    """훅이 **실제로 받은 것**을 그대로 남긴다 — 짐작을 문서에 적지 않기 위해서다."""
    row = {
        "때": datetime.now().isoformat(timespec="seconds"),
        "이벤트": str(payload.get("hook_event_name") or ""),
        "갈래": str(payload.get("source") or payload.get("reason") or ""),
        "sid": str(payload.get("session_id") or "")[:8],
        "칸": sorted(str(k) for k in payload.keys())[:12],
    }
    d = _read_json(LOG, {})
    rows = [r for r in (d.get("기록") or []) if isinstance(r, dict)]
    rows.append(row)
    seen = dict(d.get("본갈래") or {})
    key = "%s/%s" % (row["이벤트"] or "?", row["갈래"] or "?")
    seen[key] = int(seen.get(key) or 0) + 1
    _write_json(LOG, {"기록": rows[-KEEP:], "본갈래": seen})
    return row


def orphan_claims(now=None):
    """**아무도 못 푸는 점유**를 고른다 — pid 는 살아 있는데 그 세션은 사라진 것.

    ★ **잘못 회수하면 살아 있는 옆 창의 점유를 빼앗는다** — 못 잡는 것보다 나쁘다.
      그래서 문을 넷 건다:
        ① 그 sid 가 **살아 있는 목록에 없다**(`session_wrapup.live_sids` — 판단을
           빌린다. 여기서 다시 만들면 두 눈이 갈린다)
        ② **내 sid 는 그 목록에 있다** — 목록 자체를 믿을 수 있다는 증거다.
           내가 거기 없으면 목록을 못 읽은 것이므로 **아무것도 회수하지 않는다**(`[169]`)
        ③ pid 가 살아 있다 — 죽은 것은 기존 규칙이 이미 즉시 넘겨받는다
        ④ 잡힌 지 %d분은 지났다 — 방금 잡은 창은 트랜스크립트가 아직 없을 수 있다
    """
    now = time.time() if now is None else now
    try:
        import ai_claim
        import session_wrapup
    except Exception:
        return []
    me = ""
    try:
        me = ai_claim.session_id() or ""
    except Exception:
        pass
    try:
        live = {str(s) for s in session_wrapup.live_sids()}
    except Exception:
        return []                               # 못 읽었으면 손대지 않는다
    if not live or (me and me not in live):
        return []                               # 목록을 못 믿는다 — 내가 거기 없다
    out = []
    for res, claim in (ai_claim.load() or {}).items():
        if not isinstance(claim, dict):
            continue
        sid = str(claim.get("sid") or "")
        if not sid or sid == me or sid in live:
            continue
        try:
            if ai_claim._is_dead(claim):
                continue                        # 기존 규칙 몫이다
        except Exception:
            continue
        age = (now - float(claim.get("at") or 0)) / 60.0
        if age < MIN_AGE_MIN:
            continue
        out.append({"자원": res, "sid": sid, "who": str(claim.get("who") or ""),
                    "왜": str(claim.get("why") or "")[:60], "나이분": int(age)})
    return out


orphan_claims.__doc__ = (orphan_claims.__doc__ or "") % MIN_AGE_MIN


def reclaim(rows):
    """고아 점유를 실제로 놓는다. **놓은 것은 반드시 적어 남긴다.**"""
    done = []
    try:
        import ai_claim
    except Exception:
        return done
    for r in rows:
        try:
            ai_claim.free(r.get("who") or "claude", r["자원"], force=True)
            done.append(r)
        except Exception:
            pass                                # 못 놓으면 그냥 보고만 한다
    return done


def recent_changes(minutes=FRESH_MIN, now=None):
    """최근에 바뀐 `.py` — **요약 속 사본을 믿지 말라는 근거**다.

    ★ 목록이 비면 "요약이 신선하다"가 아니라 **"바뀐 것을 못 봤다"** 일 수도 있다.
      그래서 훑은 자리 수를 같이 돌려준다(`[169]`).
    """
    now = time.time() if now is None else now
    cut = now - float(minutes) * 60
    hits, looked = [], 0
    for rel in WATCH_DIRS:
        base = os.path.join(ROOT, rel) if rel else ROOT
        try:
            with os.scandir(base) as it:        # 목록에 딸려 온 stat 을 쓴다(`[198]`)
                for e in it:
                    if not e.name.endswith(".py"):
                        continue
                    looked += 1
                    try:
                        if e.stat().st_mtime >= cut:
                            hits.append(os.path.join(rel, e.name) if rel else e.name)
                    except OSError:
                        pass
        except OSError:
            continue
    hits.sort()
    return {"목록": hits, "훑은수": looked, "창분": int(minutes)}


def _handoff_head(limit=2):
    """'먼저 처리할 것' 은 **이미 써 둔 인계 문서**에서 읽는다(`loop_policy` 와 같은 눈)."""
    try:
        import loop_policy
        text = loop_policy._read(loop_policy.HANDOFF)
        items = loop_policy.handoff_items(text) if text else []
        return {"수": len(items), "앞": [str(i)[:70] for i in items[:limit]],
                "읽음": bool(text)}
    except Exception:
        return {"수": 0, "앞": [], "읽음": False}


def _my_work():
    """내 점유와 내가 맡은 분담판 — `/clear` 뒤 제일 먼저 잃는 것이 이것이다."""
    mine, took = [], []
    try:
        import ai_claim
        me = ai_claim.session_id() or ""
        for res, c in (ai_claim.load() or {}).items():
            if isinstance(c, dict) and str(c.get("sid") or "") == me:
                mine.append("%s(%s)" % (res, str(c.get("why") or "")[:40]))
    except Exception:
        pass
    try:
        import worksplit
        for it in (worksplit.load() or {}).get("items", []):
            if str(it.get("owner") or "") == "claude" and \
                    str(it.get("state") or "") in ("진행", "맡음"):
                took.append("[%s] %s" % (it.get("no"), str(it.get("title") or "")[:50]))
    except Exception:
        pass
    return mine, took[:3]


def build(now=None, apply_reclaim=True):
    """새 대화가 받을 한 장을 만든다. 값싸다 — Z: 도 워크북도 열지 않는다."""
    orphans = orphan_claims(now=now)
    freed = reclaim(orphans) if (apply_reclaim and orphans) else []
    fresh = recent_changes(now=now)
    mine, took = _my_work()
    head = _handoff_head()
    tier = {}
    try:                                        # 값은 `ai_tier.TIERS` 한 곳에서 온다
        import loop_policy
        got = loop_policy.build()
        tier = {"무게": got.get("무게"), "모델": got.get("모델"),
                "노력": got.get("노력"), "왜": got.get("왜")}
    except Exception:
        pass
    return {"고아점유": orphans, "놓음": [r["자원"] for r in freed],
            "최근변경": fresh, "내점유": mine, "내가맡은일": took,
            "먼저처리할것": head, "등급": tier}


def message(st):
    """주입 문구 — **짧아야 한다.** 이것도 컨텍스트를 먹는다."""
    lines = ["[세션 이어받기] compact·clear 뒤에도 일은 그대로 남아 있습니다."]
    if st["놓음"]:
        lines.append("· 아무도 못 푸는 점유 %d개를 회수했습니다: %s (그 세션은 사라졌고 "
                     "pid 만 살아 있었습니다)" % (len(st["놓음"]), ", ".join(st["놓음"])))
    elif st["고아점유"]:
        lines.append("· 고아 점유 후보 %d개 — 근거가 약해 보고만 합니다: %s"
                     % (len(st["고아점유"]),
                        ", ".join(r["자원"] for r in st["고아점유"])))
    if st["내점유"]:
        lines.append("· 내 점유: " + " · ".join(st["내점유"]))
    if st["내가맡은일"]:
        lines.append("· 내가 맡은 일: " + " · ".join(st["내가맡은일"]))
    h = st["먼저처리할것"]
    if not h["읽음"]:
        lines.append("· 먼저 처리할 것: **못 읽었습니다**(인계 문서) — 없는 것이 아니라 못 본 것입니다")
    elif h["수"]:
        lines.append("· 먼저 처리할 것 %d건: %s" % (h["수"], " / ".join(h["앞"])))
    f = st["최근변경"]
    if f["목록"]:
        head = ", ".join(f["목록"][:5]) + (" 외 %d" % (len(f["목록"]) - 5) if len(f["목록"]) > 5 else "")
        lines.append("· ★ 최근 %d분 안에 바뀐 코드 %d개: %s"
                     % (f["창분"], len(f["목록"]), head))
        lines.append("  → **요약·기억 속의 그 파일 사본을 근거로 삼지 마세요.** 고치거나 "
                     "진단하기 전에 Read 로 지금 내용을 다시 읽으세요(2026-08-13 실측: "
                     "낡은 사본을 믿고 이미 고쳐진 코드를 '어긋났다'고 진단했습니다).")
    elif not f["훑은수"]:
        lines.append("· 최근 변경: **못 훑었습니다** — '바뀐 것 없음'이 아닙니다")
    t = st["등급"]
    if t.get("모델"):
        lines.append("· 이번 세션 권장 등급: %s → `%s` / 노력 `%s` (%s)"
                     % (t.get("무게"), t["모델"], t["노력"], str(t.get("왜") or "")[:60]))
    return "\n".join(lines)


def hook():
    """`SessionStart` 훅. **무슨 일이 있어도 exit 0** — 세션 시작을 막지 않는다."""
    payload = {}
    try:
        raw = sys.stdin.read()
        if raw.strip():
            payload = json.loads(raw)
    except Exception:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}
    try:
        note(payload)
    except Exception:
        pass
    try:
        msg = message(build())
    except Exception:
        msg = ""
    if msg:
        print(json.dumps({"hookSpecificOutput": {
            "hookEventName": "SessionStart", "additionalContext": msg}},
            ensure_ascii=False))
    return 0


def main(argv=None):
    argv = sys.argv[1:] if argv is None else list(argv)
    if "--start" in argv:
        return hook()
    st = build(apply_reclaim="--apply" in argv)
    if "--json" in argv:
        print(json.dumps(st, ensure_ascii=False))
    else:
        print(message(st))
        if st["고아점유"] and "--apply" not in argv:
            print("\n회수하려면: python session_boundary.py --apply")
    return 0


if __name__ == "__main__":
    sys.exit(main())
