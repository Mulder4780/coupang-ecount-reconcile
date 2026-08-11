# -*- coding: utf-8 -*-
"""밴드 수집 생존확인의 계약 검사 (사고 #35 재발방지) — **아직 검증번호가 없다**

왜 별도 파일인가
--------------------------------------------------------------------------------
`tests/synthetic_check.py` 를 지금 옆 세션이 편집 중이라(실측 00:48:57 저장, `code`
점유 보유) 같은 파일 동시 편집 금지에 걸린다. 사고 #38 이 똑같은 상황을 겪고 쓴
방법을 그대로 따른다 — 몸통을 여기 두고 분담판에 올려, 그 세션이 커밋하면
`synthetic_check.py` 로 옮기며 검증번호를 단다.

★ **이 파일을 지우고 문자열 검사로 대신하지 말 것.** #38 이 남� 교훈이 그것이다 —
  그때 놓친 버그(`run()` 이 마지막 한 줄만 준다)는 소스에 그 문자열이 있는지로는
  안 잡히고 **실제로 돌려야** 드러났다. 여기서도 판정 함수를 진짜로 부른다.

  python ecount/tests/_pending_band_liveness_check.py
"""
import io
import json
import os
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
for p in (ROOT, os.path.join(ROOT, "band")):
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

FAILED = []


def check(name, got, want):
    if got != want:
        FAILED.append("%s — 얻음 %r · 바람 %r" % (name, got, want))
        print("  실패 %-52s 얻음 %r" % (name, got))
    else:
        print("  통과 %-52s" % name)


def read(p):
    with io.open(p, "r", encoding="utf-8") as f:
        return f.read()


def t_verdict_contract():
    """성공은 생존확인 **과** 수확증가가 다 있을 때만이다 — 이것이 #35 의 본체다."""
    import liveness as L

    def H(댓글=0, 영개=0, 미확인=0, n=1):
        return {"합계": {"댓글수": 댓글, "확인된0개": 영개, "미확인": 미확인,
                        "댓글담김": 1 if 댓글 else 0}, "덤프수": n, "덤프": []}

    A = {"판정": L.ALIVE, "상태": {"ok": 12, "total": 250}}
    D = {"판정": L.DEAD, "왜": "전역이 없다"}
    U = {"판정": L.UNKNOWN, "왜": "탐침 실패"}
    N = {"판정": L.NEVER, "왜": "심장 소리도 없다"}

    # ① 수확이 있어도 살아 있음을 확인 못 했으면 성공이 아니다.
    check("죽음+수확이라도 성공 아님", L.verdict([D], H(댓글=99))[0], 5)
    check("모름+수확이라도 성공 아님", L.verdict([U], H(댓글=99))[0], 4)
    check("시작안함+수확이라도 성공 아님", L.verdict([N], H(댓글=99))[0], 5)
    # ② 살아 있어도 수확이 안 늘면 성공이 아니다([162] — 살아서 0건을 담을 수 있다).
    check("생존+미확인만은 성공 아님", L.verdict([A], H(미확인=40))[0], 3)
    check("생존+덤프없음은 성공 아님", L.verdict([A], H(n=0))[0], 3)
    # ③ 둘 다 있으면 성공. 그리고 '확인된 0개'는 진척으로 센다([199]).
    check("생존+댓글증가만 성공", L.verdict([A], H(댓글=7))[0], 0)
    check("생존+확인된0개도 진척", L.verdict([A], H(영개=30))[0], 0)
    # ④ 0 을 뭉치지 않는다([169]) — 세 가지가 서로 다른 이름으로 나와야 한다.
    names = {L.verdict([A], H(n=0))[1], L.verdict([A], H(미확인=1))[1],
             L.verdict([A], H(영개=1))[1]}
    check("0건을 세 갈래로 가른다", len(names), 3)
    # ⑤ 설명이 근거를 담아야 한다 — '왜'가 없으면 사람이 또 없는 것을 찾아 나선다.
    check("죽음 설명이 이유를 말한다", "수집이 아니다" in L.verdict([D], H())[2], True)


def t_false_death():
    """싱싱한 심장 소리는 죽음이 아니다 — 거짓 죽음은 중복 수확을 부른다(#36)."""
    import liveness as L
    now = int(time.time() * 1000)
    cold = json.dumps({"verdict": "DIED_AFTER_START", "err": "NO __GRAB",
                       "beat": {"at": now - 3600 * 1000, "running": True}})
    warm = json.dumps({"verdict": "DIED_AFTER_START", "err": "NO __GRAB",
                       "beat": {"at": now - 4000, "running": True}})
    stopped = json.dumps({"verdict": "DIED_AFTER_START", "err": "NO __GRAB",
                          "beat": {"at": now - 4000, "running": False}})
    check("식은 심장 = 죽음", L.classify(cold)["판정"], L.DEAD)
    check("싱싱한 심장 = 다른탭 생존", L.classify(warm)["판정"], L.ALIVE_OTHER)
    check("다 끝난 심장 = 죽음(생존 아님)", L.classify(stopped)["판정"], L.DEAD)
    check("한 번도 안 함 = 죽음과 다른 이름",
          L.classify(json.dumps({"verdict": "NEVER_STARTED"}))["판정"], L.NEVER)
    check("탐침 무응답 = 모름", L.classify(None)["판정"], L.UNKNOWN)


def t_harvest_yardstick():
    """수확은 흡수기와 **같은 잣대**로 센다 — 시각 없는 댓글은 댓글이 아니다([130])."""
    import liveness as L
    with tempfile.TemporaryDirectory() as td:
        p = os.path.join(td, "dump_202608120101_90610953.json")
        with io.open(p, "w", encoding="utf-8") as f:
            json.dump({"band": "90610953", "posts": {
                "1": {"comments": [{"author": "가", "created_at": 1, "content": "접수취소"}]},
                "2": {"comments": [{"author": "나", "created_at": None, "content": ""}]},
                "3": {"comments": [], "comments_full": True},
                "4": {},
            }}, f, ensure_ascii=False)
        h = L.harvest_since(0, dirs=[td])
        check("시각 있는 댓글만 센다", h["합계"]["댓글수"], 1)
        check("껍데기는 안 센다", h["합계"]["껍데기"], 1)
        check("확인된 0개를 따로 센다", h["합계"]["확인된0개"], 1)
        check("미확인을 따로 센다", h["합계"]["미확인"], 2)
        check("회차 시작 뒤 덤프만 본다",
              L.harvest_since(time.time() + 60, dirs=[td])["덤프수"], 0)


def t_collector_defends_itself():
    """grab_posts.js 가 최상위 이동을 **막고 기록하는지**.

    2026-08-12 실측(로컬 재현): sandbox `allow-same-origin allow-scripts` 인 iframe 은
      · 프레임 깨기가 SecurityError 로 **막히고**(크롬: "The frame attempting
        navigation of the top-level window is sandboxed, but ... 'allow-top-navigation'
        ... is not set")
      · 그런데도 `contentDocument` 는 **읽힌다**(본문을 그대로 뜯었다).
    같은 자식을 sandbox 없이 실었을 때는 **최상위 문서가 실제로 넘어갔다** — #35 재현.
    """
    import re
    js = read(os.path.join(ROOT, "band", "grab_posts.js"))
    # ★ 파일 전체를 문자열로 훑지 않고 **실제로 실리는 속성값**을 뜯어 본다.
    #   첫 판은 `"allow-top-navigation" in js` 였는데, 그 플래그를 **왜 안 넣는지
    #   설명한 주석**에 걸려 빨강이 됐다 — 사고 #38 이 겪은 것과 같은 모양이다
    #   ('사고를 적는 행위가 사고 모양을 만들었다'). 설명은 남겨야 하고, 검사는
    #   값을 봐야 한다.
    m = re.search(r"setAttribute\(\s*'sandbox'\s*,\s*'([^']*)'\s*\)", js)
    toks = m.group(1).split() if m else []
    check("iframe 에 sandbox 를 건다", bool(m), True)
    check("같은 출처 유지(본문을 읽을 수 있어야 한다)", "allow-same-origin" in toks, True)
    check("스크립트 허용(SPA 가 그려야 한다)", "allow-scripts" in toks, True)
    # ★ 이 플래그가 실리면 막는 뜻이 사라진다 — 값에 절대 넣지 말 것.
    check("최상위 이동은 허용하지 않는다",
          [t for t in toks if t.startswith("allow-top-navigation")], [])
    # 죽음을 기록하는 손잡이 — 넘어간 뒤에는 전역이 없으니 localStorage 여야 한다.
    check("죽음 기록을 localStorage 에 남긴다", "__grabDeath" in js, True)
    check("심장 소리를 남긴다", "__grabBeat" in js, True)
    check("beforeunload 를 듣는다", "beforeunload" in js, True)
    check("pagehide 도 듣는다", "pagehide" in js, True)
    # 중간 저장이 없으면 죽는 순간 그때까지의 수확이 통째로 사라진다.
    check("중간 저장이 있다", "saveEvery" in js, True)
    # 상태 창구에 생존확인이 읽을 값이 있어야 한다(없으면 폴링이 눈멀다).
    for k in ("tried", "saves", "prevDeath", "prevBeat", "sandboxFellBack"):
        check("__grabStatus 가 %s 를 내놓는다" % k, k in js, True)


def t_probe_contract():
    """`NO __GRAB` 은 **파일 셋이 공유하는 계약**이다 — 이름이 갈리면 죽음을 못 알아본다."""
    st = read(os.path.join(ROOT, "band", "band_dump_state.js"))
    check("탐침이 NO __GRAB 을 그대로 말한다", "NO __GRAB" in st, True)
    check("탐침이 심장 소리를 읽는다", "__grabBeat" in st, True)
    check("탐침이 죽음 기록을 읽는다", "__grabDeath" in st, True)
    check("탐침이 시작안함과 죽음을 가른다",
          "NEVER_STARTED" in st and "DIED_AFTER_START" in st, True)
    # 죽음 표식을 찾는 쪽들이 여전히 같은 문자열을 본다.
    bc = read(os.path.join(ROOT, "band", "browser_chain.py"))
    check("browser_chain 이 같은 표식을 본다", "NO __GRAB" in bc, True)
    import liveness as L
    check("liveness 도 같은 표식을 본다", L.DEAD_MARK, "NO __GRAB")


def t_dump_name_is_not_a_ghost():
    """중간 저장 이름이 **유령 밴드**를 만들지 않는지 — 실제 규칙 함수로 확인한다.

    `dump_<12자리>s2_90610953.json` 의 후보는 `202608120050`·`90610953` 둘이다.
    앞엣것이 7~10자리 문에서 떨어져야 밴드가 옳게 잡힌다. 숫자를 그냥 이어 붙이면
    13자리 덩어리가 되어 없는 밴드가 생긴다(2026-08-08 두 차례 실사고와 같은 모양).
    """
    import convert_dump as cd
    for name in ("dump_202608120050_90610953.json",
                 "dump_202608120050s2_90610953.json",
                 "dump_202608120050s12_84789192.json"):
        want = "90610953" if "90610953" in name else "84789192"
        check("이름 규칙: %s" % name, cd.band_from_name(name, known=set()), want)


def t_yields_to_other_sessions():
    """밴드를 다른 세션이 잡고 있으면 **빼앗지 않고 물러난다**(사고 #27)."""
    import liveness as L
    src = read(os.path.join(ROOT, "band", "liveness.py"))
    check("점유 판정을 ai_claim 에서 빌린다", "_is_dead" in src and "_is_mine" in src, True)
    check("못 읽으면 '없음'으로 치지 않는다", "확인할 수 없다" in src, True)
    # subprocess.run(timeout=) 은 이 프로젝트에서 금지다([175]) — 회차가 영원히 멈춘다.
    check("subprocess.run(timeout= 을 쓰지 않는다", "subprocess.run(" in src, False)
    check("proc_guard.run_tree 를 쓴다", "proc_guard.run_tree" in src, True)


def main():
    print("밴드 수집 생존확인 계약 검사 (사고 #35)")
    for fn in (t_verdict_contract, t_false_death, t_harvest_yardstick,
               t_collector_defends_itself, t_probe_contract,
               t_dump_name_is_not_a_ghost, t_yields_to_other_sessions):
        print("\n[%s] %s" % (fn.__name__, (fn.__doc__ or "").splitlines()[0]))
        fn()
    print("")
    if FAILED:
        print("★ 실패 %d건" % len(FAILED))
        for f in FAILED:
            print("  - %s" % f)
        return 1
    print("ALL GREEN — 생존확인 없이는 성공이라 적히지 않는다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
