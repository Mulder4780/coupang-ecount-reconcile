# -*- coding: utf-8 -*-
"""조직도가 바뀌면 붙잡아 두고, **따라가지 못한 것**을 말한다 (2026-08-13 지시).

사용자 지시: **"조직도 변경사항 감지되면 자동으로 반영해"**

★ **먼저 있는 것을 확인했다 — 반영하는 손은 이미 있다**(`[162]`):
  · 조직도 코드(`webapp/app_server.py`·`webapp/index.html`)가 바뀌면
    `watchdog.heal_stale_server` 가 **스스로 서버를 갈아 준다**(`[156]`).
    담당자가 최근 10분 안에 앱을 만졌으면 `restart_server.guard()` 가 미룬다(`[265]`).
  · 폰 사본은 **갱신할 것이 없다** — `cloud_publish` 에 조직도가 한 칸도 안 실리고
    `docs/app.html` 에는 조직도 마크업이 0건이다(실측). 없는 것을 갱신한다고 적으면
    그 문장 자체가 거짓이 된다(`[169]`).
  그래서 여기는 **반영하는 자리가 아니라 반영됐는지 보는 자리**다(읽기 전용).

★ **아무도 안 보던 자리는 여기다** — 흐름 정의는 **DB 표(`flow_step`)** 에서 온다.
  사람이 앱에서 흐름도를 고치면 **파일 mtime 이 한 톨도 안 움직인다.** 그러니
  `heal_stale_server` 는 "서버 정상"이라 말하고, 스케줄러도 아무 말이 없고,
  조직도 흐름 그림만 조용히 달라진다. 코드가 아닌 변경은 지금까지 어느 눈에도 안 걸렸다.

★ **지문에 사람 상태를 넣지 않는다**(`[170]` 재수집 `hash_on` 사고와 같은 모양).
  `state`·`msg`·`ago`·`badge` 와 **`ai` 구역**(세션·회차)은 회차마다 달라진다 —
  넣으면 **매번 '바뀜'** 이 되어 아무도 안 본다. 지문은 **누가 있고 무슨 일을 하나**뿐이다.

★ 근거는 `app_server.get_orgchart()` **하나**다(`[162]`). 로스터·흐름을 여기서 다시
  조립하면 화면과 감시자가 서로 다른 조직도를 놓고 답한다. 실측 0.26초(import 0.16 +
  호출 0.10)라 비싼 탐색도 아니다(`[168]`).

★ **못 읽으면 '이상 없음'이 아니라 '확인 못 함'** 이다(`[169]`).
"""
import hashlib
import json
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(ROOT, "reports", "조직도_반영.json")
KEEP = 30                       # 변경 이력은 최근 30건만 남긴다

# 사람 상태가 담기는 칸 — **지문에 넣지 않는다**. 넣으면 매 회차가 '바뀜'이다.
VOLATILE = ("state", "msg", "ago", "badge")
# 회차·세션이 들어오는 구역 — 자리 목록에서 뺀다(사람 자리가 아니다).
VOLATILE_ZONES = ("ai",)


def _snapshot():
    """지금 조직도가 **보여 주는 것**. 돌려주는 것: (스냅샷, 못읽은이유)."""
    try:
        wp = os.path.join(ROOT, "webapp")
        if wp not in sys.path:
            sys.path.insert(0, wp)
        import app_server
        d = app_server.get_orgchart()
    except Exception as exc:
        return None, "%s: %s" % (type(exc).__name__, str(exc)[:120])
    seats = []
    for z in (d.get("zones") or []):
        key = str(z.get("key") or "")
        if key in VOLATILE_ZONES:
            continue
        for p in (z.get("people") or []):
            seats.append({"구역": key, "이름": str(p.get("name") or ""),
                          "역할": str(p.get("role") or "")})
    f = d.get("flow") or {}
    if not f.get("ok"):
        # 흐름을 못 읽은 것은 '흐름이 없다'가 아니다.
        return {"자리": seats, "흐름": None,
                "흐름왜": str(f.get("why") or "")[:120]}, ""
    flow = {"접수": [{"단계": r.get("label", ""), "담당": list(r.get("who") or [])}
                     for r in (f.get("entries") or [])],
            "차례": [{"n": r.get("n"), "단계": r.get("label", ""),
                      "담당": list(r.get("who") or [])}
                     for r in (f.get("chain") or [])]}
    return {"자리": seats, "흐름": flow, "흐름왜": ""}, ""


def fingerprint(snap):
    """정의만 해싱한다 — 사람 상태는 위에서 이미 빠져 있다."""
    body = json.dumps({"자리": snap.get("자리"), "흐름": snap.get("흐름")},
                      ensure_ascii=False, sort_keys=True)
    return hashlib.sha1(body.encode("utf-8")).hexdigest()[:12]


def _diff(old, new):
    """무엇이 바뀌었는지 **사람 말로**. 지문만 적으면 사람이 또 파일을 뒤진다."""
    out = []
    o = {(s["구역"], s["이름"]): s.get("역할", "") for s in (old or {}).get("자리") or []}
    n = {(s["구역"], s["이름"]): s.get("역할", "") for s in (new or {}).get("자리") or []}
    for k in sorted(set(n) - set(o)):
        out.append("자리 생김: %s(%s)" % (k[1], k[0]))
    for k in sorted(set(o) - set(n)):
        out.append("자리 없어짐: %s(%s)" % (k[1], k[0]))
    for k in sorted(set(o) & set(n)):
        if o[k] != n[k]:
            out.append("역할 바뀜: %s — %s → %s"
                       % (k[1], o[k] or "(빈칸)", n[k] or "(빈칸)"))
    fo, fn = (old or {}).get("흐름"), (new or {}).get("흐름")
    if fo != fn:
        if fo is None:
            out.append("흐름 읽힘(전에는 못 읽었다)")
        elif fn is None:
            out.append("흐름 못 읽음(전에는 읽혔다)")
        else:
            for key in ("접수", "차례"):
                a = [r["단계"] for r in fo.get(key) or []]
                b = [r["단계"] for r in fn.get(key) or []]
                if a != b:
                    out.append("흐름 %s 바뀜: %d → %d단계" % (key, len(a), len(b)))
                elif ([r["담당"] for r in fo.get(key) or []]
                      != [r["담당"] for r in fn.get(key) or []]):
                    out.append("흐름 %s 담당 바뀜" % key)
    return out


def _seat_names(snap):
    return {s["이름"] for s in (snap.get("자리") or []) if s.get("이름")}


def follow(snap):
    """**무엇이 못 따라갔나**로 갈라 말한다. 돌려주는 것: (못따라감, 확인못함)."""
    gaps, unknown = [], []

    # ① 흐름도가 부르는 사람인데 조직도에 자리가 없다.
    #    ★ 자동으로 자리를 만들지 않는다 — 누가 어느 구역에 앉는지는 사람만 안다(`[172]`).
    if snap.get("흐름") is None:
        unknown.append("흐름 정의를 못 읽었다%s"
                       % (" — " + snap["흐름왜"] if snap.get("흐름왜") else ""))
    else:
        seats = _seat_names(snap)
        miss = []
        for key in ("접수", "차례"):
            for r in snap["흐름"].get(key) or []:
                for w in r.get("담당") or []:
                    if w and w not in seats and w not in miss:
                        miss.append(w)
        if miss:
            gaps.append({"갈래": "자리없는담당",
                         "말": "흐름도가 부르는데 조직도에 자리가 없는 사람 %d명: %s"
                               % (len(miss), ", ".join(miss)),
                         "조치": "webapp/app_server.py 의 STAFF_CENTERS·"
                                 "AS_TECH_CENTERS 에 넣을지 사람이 정합니다"
                                 "(자동으로 만들지 않습니다)"})

    # ② 서버가 아직 옛 코드다 — 정의는 바뀌었는데 화면에는 아직 안 나갔다는 뜻이다.
    #    ★ 여기서 재시작하지 않는다. 같은 회차의 `heal_stale_server` 가 이미 그 일을
    #      하고, 담당자가 쓰는 중이면 미룬다(`[265]`). 두 손이 같은 서버를 갈면
    #      폰이 502 를 두 번 받는다(실측 4~9초).
    try:
        wp = os.path.join(ROOT, "webapp")
        if wp not in sys.path:
            sys.path.insert(0, wp)
        import restart_server
        st = restart_server.stale()
        if st:
            newer = [f for f in st[2]
                     if f in ("webapp/app_server.py", "webapp/index.html")]
            if newer:
                gaps.append({"갈래": "서버옛코드",
                             "말": "조직도 코드가 바뀌었는데 서버는 아직 옛 코드다: "
                                   + ", ".join(newer),
                             "조치": "python webapp/restart_server.py"})
    except Exception as exc:
        unknown.append("서버 코드나이 확인 못 함: %s" % str(exc)[:80])
    return gaps, unknown


def _load():
    try:
        d = json.load(open(STATE, encoding="utf-8"))
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _save(d):
    try:
        os.makedirs(os.path.dirname(STATE), exist_ok=True)
        tmp = STATE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(d, fh, ensure_ascii=False, indent=1)
        os.replace(tmp, STATE)
    except Exception:
        pass                     # 자국을 못 남긴다고 회차를 세우지 않는다


def build(save=True):
    """지금 상태 한 장. `save=False` 면 아무것도 안 적는다(검증·조회용)."""
    prev = _load()
    snap, why = _snapshot()
    now = datetime.now().isoformat(timespec="seconds")
    if snap is None:
        # ★ 못 읽었다고 옛 지문을 지우지 않는다 — 지우면 다음 회차가 처음처럼 굴어
        #   '바뀜'을 한 번 더 만든다(없던 변경을 만들어 내는 자리다).
        out = dict(prev)
        out["때"] = now
        out["이번에바뀜"] = False
        out["못따라감"] = []
        out["확인못함"] = ["조직도를 못 읽었다 — %s" % why]
        if save:
            _save(out)
        return out

    fp = fingerprint(snap)
    changed = bool(prev.get("지문")) and prev["지문"] != fp
    hist = list(prev.get("바뀜이력") or [])
    if changed:
        hist.append({"때": now, "이전지문": prev.get("지문"), "지문": fp,
                     "무엇": _diff(prev.get("스냅샷"), snap)})
        hist = hist[-KEEP:]

    gaps, unknown = follow(snap)
    out = {"때": now, "지문": fp, "처음": not prev.get("지문"),
           "이번에바뀜": changed, "스냅샷": snap, "바뀜이력": hist,
           "못따라감": gaps, "확인못함": unknown}
    if save:
        _save(out)
    return out


def notices(d=None):
    """인계 '먼저 처리할 것'에 올릴 것.

    ★ **바뀐 것 자체는 알리지 않는다** — 바뀌고 잘 따라갔으면 그것은 정상이다.
      정상까지 경보하면 진짜 경보가 묻힌다(`[170]`).
    ★ **확인 못 한 것은 이상 없음이 아니다**(`[169]`) — 같이 올린다.
    """
    d = _load() if d is None else d
    out = []
    for g in (d.get("못따라감") or []):
        out.append(("[조직도] " + g.get("말", ""),
                    g.get("조치") or "python org_watch.py --print"))
    for u in (d.get("확인못함") or []):
        out.append(("[조직도] 확인 못 함 — " + str(u), "python org_watch.py --print"))
    return out


def _line(d):
    if d.get("확인못함") and not d.get("지문"):
        return "조직도 확인 못 함: %s" % str(d["확인못함"][0])[:70]
    bits = ["조직도 자리 %d" % len((d.get("스냅샷") or {}).get("자리") or [])]
    if d.get("처음"):
        bits.append("첫 지문 %s" % d.get("지문"))
    elif d.get("이번에바뀜"):
        h = (d.get("바뀜이력") or [{}])[-1]
        bits.append("바뀜: " + (", ".join(h.get("무엇") or []) or "지문만 다름")[:70])
    else:
        bits.append("그대로")
    if d.get("못따라감"):
        bits.append("못 따라간 것 %d건" % len(d["못따라감"]))
    if d.get("확인못함"):
        bits.append("확인 못 함 %d건" % len(d["확인못함"]))
    return " · ".join(bits)


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if hasattr(sys.stdout, "reconfigure"):      # 무인 회차는 stdout 이 None 이다(`[235]`)
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    d = build(save="--dry" not in argv)
    print(_line(d))
    if "--print" in argv:
        for t, a in notices(d):
            print("  - %s\n      %s" % (t, a))
        for h in (d.get("바뀜이력") or [])[-3:]:
            print("  · %s %s" % (h.get("때"), ", ".join(h.get("무엇") or [])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
