# -*- coding: utf-8 -*-
"""
collect_gate.py — 수집을 **누가** 하느냐가 아니라 **지금 자원이 비었나**로 정한다
                  (2026-08-19 지시)

사용자 지시: "CSOS 리서치 및 자료 수집 세션에서 하던 일을 지금 현재 이 세션에서도
작업할 수 있게 알고리즘 반영해"

무엇이 막고 있었나 — 실측부터
  2026-08-07 규칙은 "수집은 'CSOS 리서치 및 자료 수집' 세션이 맡는다"였다. 그런데
  그것은 **세션 이름**으로 쓴 규칙이라 기계가 지킬 수 없다 — `ai_claim` 은 창 제목을
  모르고 sid 만 안다. 2026-08-19 실측:
      작업 차선 collect · build **둘 다 비어 있음**
      `band` 점유 **비어 있음** (쥔 것은 code · publish 뿐)
  즉 **코드로 막는 것은 한 줄도 없었다.** 지킨 것은 사람의 약속뿐이었고, 그 약속은
  두 방향으로 샜다 — 옆 창이 마음대로 긁을 수도 있었고, 이 창은 **비어 있는 자원을
  두고도 일을 안 했다.**

그래서 이름을 자원으로 바꾼다
  규칙이 정말 막으려던 것은 '누구인가'가 아니라 **'같은 것을 같은 때에'** 다:
    · 사고 #27 — 두 세션이 같은 밴드를 긁어 캐시가 서로 덮였다 → `ai_claim` 의
      배타 자원 `band` 가 이미 막는다.
    · 사고 #29 — 수집이 Z:(SMB) I/O 를 독점해 앱·코딩까지 10시간 느려졌다 →
      Z: 를 훑는 회차들이 이미 락을 쥔다(`coordinate.LOCKS`).
  둘 다 **이미 기계가 아는 사실**이다. 그러므로 이 문은 새 판단을 만들지 않고
  세 판정을 **빌려서 합치기만** 한다(`[162]`).

갈래는 셋이다 — '모름'을 '가능'으로 치지 않는다 (`[169]`)
  가능 · 양보(주인 이름과 함께) · **모름**(못 읽었다). 못 읽었는데 긁으면 그때가
  바로 캐시가 덮이는 순간이다. 물러나는 값은 회차 한 번이고, 부딪히는 값은
  **되돌릴 수 없다.**

양보는 주장이므로 자국을 남긴다 (`[293]`)
  물러날 때 `coordinate.record_yield(주인=...)` 로 적는다 — 주인 이름은
  **지어내지 않고** `coordinate.running()`·점유판이 준 그대로 쓴다. 매일 양보만
  하는 문은 없는 문과 같으므로 굶주림 판정이 그것을 잡는다.

쓰는 법
  python collect_gate.py                          # 지금 수집해도 되나 (읽기 전용)
  python collect_gate.py --why "덤프 흡수" --run band/convert_dump.py
  python collect_gate.py --run band/recheck_plan.py --limit 3

★ `--run` **뒤는 전부 자식 명령**이다(REMAINDER). 그러므로 `--why`·`--who` 는
  반드시 **앞에** 온다. 뒤에 적으면 자식이 그 깃발을 받고 죽는다 — 실측으로
  그대로 당했다(`recheck_plan.py: error: unrecognized arguments: --why ...`).
  가로채서 고쳐 주지는 않는다 — 자식이 제 `--why` 를 가질 수 있다(`[172]`).
  종료코드: 0 완주 · 1 명령 실패 · 3 양보/모름(아무것도 안 했다)

★ `--force` 는 없다. 남이 쥔 자원을 뺏는 문이 아니라, **비어 있는 자원을 놀리지
  않는** 문이다.

★ 2026-08-19 지시로 **차선 밖은 '모름'** 이 됐다 — "수집은 'CSOS 리서치 및 자료
  수집_v2' 세션에서 한다, 너는 다 양보해". 세션 이름은 기계가 못 보므로 그 약속을
  적는 자리는 **차선표**뿐이다. 수집 창은 `python lanes.py --take collect` 로 한 번
  선언하면 그 뒤로 그대로 통과한다.
  ※ 위 인용은 2026-08-19 당시 세션 이름이다 — **지금은 v1**(`COLLECT_SESSION_NAME`).
    낡은 이름을 안내에 쓰면 사람이 없는 세션을 찾는다([172]).
"""
import argparse
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# 이 문이 지키는 자원. 밴드 수집·덤프 흡수·원본 흡수가 전부 여기로 들어온다.
RESOURCE = "band"
WORK = "수집"

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _lane_verdict():
    """내 차선이 이 자원을 허락하나 → (갈래, 왜, 주인). 못 읽으면 '모름'.

    ★ **차선 밖은 '가능'이 아니라 '모름'이다** (2026-08-19 지시: "리서치 자료 수집 및
      긁어오기는 'CSOS 리서치 및 자료 수집_v1' 이 세션에서 진행할거에 너는 이쪽으로
      다 양보해"). 세션 **이름**은 기계가 못 본다 — `ai_claim` 은 창 제목을 모르고
      sid 만 안다. 그러므로 "누가 수집 창인가"를 기계가 아는 길은 **차선표 하나**다.
      차선을 안 정한 창은 수집 창인지 코딩 창인지 알 수 없으므로 물러난다:
      물러나는 값은 회차 한 번이고, 부딪히는 값은 **되돌릴 수 없다**(사고 #27 —
      두 창이 같은 밴드를 긁어 캐시가 서로 덮였다).

    ★ `lanes` 자체의 "차선을 안 정했으면 아무것도 막지 않는다"는 **그대로 둔다.**
      거기를 뒤집으면 무인 회차와 `ai_claim --take` 까지 같이 막힌다. 좁히는 것은
      **사람이 손으로 부르는 이 문 안에서만**이다(실측: 이 문을 코드에서 부르는 곳 0곳).

    ★ 갈래를 '양보'가 아니라 '모름'이라 부르는 이유 — **주인을 못 대기 때문**이다.
      어느 세션이 수집 창인지 표가 비어 있으면 알 수 없다. 주인을 댈 수 있을 때만
      '양보'라 적는 것이 `[313]`-⑧ 의 규칙이고, 뒤집으면 사람이 할 일이 달라진다.
    """
    try:
        import lanes
    except Exception as exc:
        return "모름", "차선표를 못 읽었다: %s" % str(exc)[:60], ""
    try:
        lane = lanes.my_lane()
    except Exception as exc:
        return "모름", "차선 판정이 실패했다: %s" % str(exc)[:60], ""
    if lane is None:
        return ("모름",
                "이 창이 수집 창인지 아직 안 정해졌다 — 수집은 수집 차선 창이 한다. "
                "이 창이 그 창이면: python lanes.py --take collect --who claude",
                "")
    try:
        ok, why = lanes.can(RESOURCE)
    except Exception as exc:
        return "모름", "차선 판정이 실패했다: %s" % str(exc)[:60], ""
    if ok:
        return "가능", "", ""
    # 주인은 지어내지 않는다 — 그 자원이 사는 차선을 실제 표에서 읽는다.
    home = ""
    try:
        for name, (label, res) in lanes.LANES.items():
            if RESOURCE in res:
                owner = lanes.owner(name)
                home = "%s 차선" % label
                if owner:
                    home = "%s(%s)" % (home, owner)
                break
    except Exception:
        home = ""
    return "양보", why, home


def _claim_verdict():
    """점유판에서 이 자원을 남이 쥐고 있나 → (갈래, 왜, 주인)."""
    try:
        import ai_claim
    except Exception as exc:
        return "모름", "점유판을 못 읽었다: %s" % str(exc)[:60], ""
    try:
        board = ai_claim.load()
        me = ai_claim.session_id()
    except Exception as exc:
        return "모름", "점유판 읽기가 실패했다: %s" % str(exc)[:60], ""
    rec = (board or {}).get(RESOURCE)
    if not rec:
        return "가능", "", ""
    try:
        if ai_claim._is_dead(rec):
            return "가능", "", ""
    except Exception:
        # 생사를 못 재면 살아 있는 것으로 본다 — 의심스러우면 물러난다.
        pass
    sid = str(rec.get("sid") or rec.get("session") or "")
    if sid and sid == me:
        return "가능", "", ""
    who = str(rec.get("who") or "?")
    주인 = "%s[%s]" % (who, sid) if sid else who
    왜 = str(rec.get("why") or "").strip()
    return "양보", ("'%s' 점유를 %s 가 쥐고 있다%s"
                    % (RESOURCE, 주인, (" — " + 왜) if 왜 else "")), 주인


def _round_verdict():
    """Z: 를 훑는 회차가 지금 도나 → (갈래, 왜, 주인). 이름은 락 파일이 정한다."""
    try:
        import coordinate
    except Exception as exc:
        return "모름", "회차 락을 못 읽었다: %s" % str(exc)[:60], ""
    try:
        도는중 = coordinate.running()
    except Exception as exc:
        return "모름", "회차 락 판정이 실패했다: %s" % str(exc)[:60], ""
    if not 도는중:
        return "가능", "", ""
    주인 = str(도는중[0])
    return "양보", ("%s 이(가) 지금 돌고 있다 — 같이 Z: 를 긁지 않는다"
                    % " · ".join(str(x) for x in 도는중)), 주인


def check():
    """수집해도 되나 → {갈래, 왜, 주인, 근거}. **아무것도 안 고친다.**

    갈래: '가능' · '양보' · '모름'. 순서가 뜻을 갖는다 — 못 읽은 것(모름)이 하나라도
    있으면 나머지가 다 초록이어도 **가능이라 말하지 않는다**(`[169]`).
    """
    근거 = []
    for 이름, fn in (("차선", _lane_verdict), ("점유", _claim_verdict),
                     ("회차", _round_verdict)):
        갈래, 왜, 주인 = fn()
        근거.append({"무엇": 이름, "갈래": 갈래, "왜": 왜, "주인": 주인})
    양보 = [x for x in 근거 if x["갈래"] == "양보"]
    모름 = [x for x in 근거 if x["갈래"] == "모름"]
    if 양보:
        첫 = 양보[0]
        return {"갈래": "양보", "왜": 첫["왜"], "주인": 첫["주인"], "근거": 근거}
    if 모름:
        첫 = 모름[0]
        return {"갈래": "모름", "왜": 첫["왜"], "주인": "", "근거": 근거}
    return {"갈래": "가능", "왜": "", "주인": "", "근거": 근거}


def _note(갈래, 왜, 주인=""):
    """자국을 남긴다. **못 남겨도 일을 막지는 않는다** — 다만 조용히 넘기지 않는다."""
    try:
        import coordinate
        if 갈래 == "양보":
            coordinate.record_yield(WORK, 주인 or "모름", 왜)
        else:
            coordinate.record_run(WORK, 갈래, 왜)
        return True
    except Exception as exc:
        print("  ! 자국을 못 남겼습니다: %s" % str(exc)[:80])
        return False


COLLECT_SESSION_NAME = "CSOS 리서치 및 자료 수집_v1"


def is_unattended():
    """이 프로세스가 **무인 회차**인가 — 사람 창과 가르는 유일한 근거.

    ★ 왜 가르나 — 이 문을 수집 스크립트에 그냥 달면 **워치독·증분 파이프라인·
      09:50 회차까지 같이 막힌다**(실측: automation_pipeline·daily_run·watchdog 이
      convert_dump·band_sync·intake 를 자식으로 부른다). 그러면 형님이 시킨 것은
      "세션끼리 나눠 하라"인데 결과는 **자동 수집이 통째로 멈추는 것**이 된다 —
      게다가 그것은 조용하다(회차는 '성공'으로 적히고 수집만 0건이 된다 · [169]).
    ★ 근거는 **세션 신분 하나**다. 스케줄러가 띄운 프로세스에는 대화 세션 ID 가
      없다(실측 2026-08-22: 이 창은 CLAUDE_CODE_SESSION_ID 가 있고, 무인 회차는
      SID_ENV 가 전부 빈다). 목록은 `ai_claim.SID_ENV` **한 곳**에서 온다([162]) —
      여기 손으로 적으면 Codex 쪽 키가 늘어난 날 그 갈래만 조용히 샌다.
    """
    if os.environ.get("COUPANG_UNATTENDED") == "1":
        return True
    keys = ("CLAUDE_CODE_SESSION_ID", "CODEX_THREAD_ID")
    try:
        import ai_claim
        keys = ai_claim.SID_ENV
    except Exception:
        pass
    return not any(os.environ.get(k) for k in keys)


def guard(why="", 자원=None):
    """수집 스크립트 **맨 앞**에서 부른다 — 남의 차선 일을 하는 사람 창만 막는다.

    2026-08-22 형님 지시: "밴드 수집은 'CSOS 리서치 및 자료 수집_v1' 이 세션에서
    하는거야, 자료 긁어오는것도 전부 여기서 할거야 알고리즘 반영해".

    ★ **이 문이 없던 자리가 진짜 구멍이었다.** [313] 이 만든 `check()`/`run()` 은
      사람이 손으로 `collect_gate.py --run ...` 이라 쳐야만 돌았고, 실측으로 그것을
      **부르는 코드가 한 줄도 없었다**(문서 언급뿐). 그래서 코딩 창에서
      `python band/convert_dump.py` 를 그냥 치면 아무것도 안 막았다 — 규칙은 있는데
      기계가 지키는 자리가 없었다(2026-08-07 이름 규칙이 샜던 것과 같은 모양).

    ★ 막는 것은 **사람 창**뿐이다(`is_unattended()` 참조).
    ★ 대화 창인데 판정을 **못 하면 막는다**([169]) — 여기서 '모름'을 통과로 치면
      차선표가 깨진 날 두 창이 같은 밴드를 긁는다(사고 #27 · 캐시 오염은 되돌릴 수
      없다). 무인은 이미 위에서 빠졌으므로 막아도 자동화는 안 죽는다.
    ★ 막을 때 **자국을 남긴다**([293]) — 매일 막히기만 하는 문은 없는 문과 같고,
      굶주림 판정이 그것을 잡는다.
    ★ **안내에 세션 이름을 적는다.** 기계는 이름을 못 보지만([313]) 사람은 그것으로
      어느 창인지 안다. 적는 자리는 `COLLECT_SESSION_NAME` **하나**다.
    """
    if is_unattended():
        return True
    v = check()
    if v["갈래"] == "가능":
        return True
    _note("양보", v["왜"], v["주인"])
    이름 = 자원 or RESOURCE
    print("수집 문이 막았습니다 — 이 창은 수집 창이 아닙니다(자원 '%s')." % 이름)
    if why:
        print("  하려던 일: %s" % why)
    print("  왜: %s" % (v["왜"] or "판정을 못 했습니다"))
    print("  이 일은 '%s' 세션이 합니다(2026-08-22 형님 지시)." % COLLECT_SESSION_NAME)
    print("  · 이 창이 그 수집 창이면:  python lanes.py --take collect --who claude")
    print("  · 무인 회차(워치독·09:50·증분)는 그대로 돕니다 — 자동 수집은 안 멈춥니다.")
    raise SystemExit(3)


def run(command, why="", who="claude"):
    """문을 통과하면 명령을 돌린다. 점유는 **내가 잡은 것만** 놓는다(`[104]`)."""
    v = check()
    if v["갈래"] != "가능":
        _note("양보", v["왜"], v["주인"])
        print("수집 물러남 (%s) — %s" % (v["갈래"], v["왜"] or "이유 미상"))
        print("  아무것도 하지 않았습니다. 풀리면 다음에 그대로 이어집니다.")
        return 3
    import ai_claim
    if not ai_claim.take(who, RESOURCE, why or "수집(collect_gate)"):
        v2 = check()
        _note("양보", v2["왜"] or "점유를 잡지 못했다", v2["주인"])
        print("수집 물러남 — 점유 '%s' 를 잡지 못했습니다." % RESOURCE)
        return 3
    상태, 코드 = "실패", 1
    try:
        from proc_guard import run_tree      # `[175]` — subprocess.run(timeout=) 금지
        r = run_tree([sys.executable] + list(command), cwd=ROOT,
                     timeout=int(os.environ.get("COUPANG_COLLECT_TIMEOUT") or 1800))
        코드 = int(getattr(r, "returncode", 1) or 0)
        # ★ **stderr 를 버리지 않는다** (`[289]`). 실측으로 그대로 당했다 —
        #   없는 깃발을 준 argparse 오류가 stderr 로 나가는데 stdout 만 찍어서
        #   화면에는 '종료코드 2' 다섯 글자만 남았다. 실패를 말하기는 하는데
        #   **왜인지는 영영 알 수 없는** 자리다.
        out = (getattr(r, "stdout", "") or "")[-1500:]
        err = (getattr(r, "stderr", "") or "")[-1500:]
        if out:
            print(out)
        if err:
            print("  [stderr] " + err)
        if getattr(r, "timed_out", False):
            print("  ! 시간 초과로 나무째 끊었습니다 (stuck_pid=%s)"
                  % getattr(r, "stuck_pid", 0))
        상태 = "완주" if 코드 == 0 else "실패(종료코드 %d)" % 코드
    except Exception as exc:
        상태 = "실패(%s)" % str(exc)[:60]
        코드 = 1
    finally:
        try:
            ai_claim.free(who, RESOURCE)
        except Exception as exc:
            print("  ! 점유를 못 놓았습니다: %s" % str(exc)[:80])
    _note("완주" if 코드 == 0 else "실패", "%s :: %s" % (상태, " ".join(command))[:200])
    print("수집 %s" % 상태)
    return 코드


def show():
    v = check()
    print("수집 문 — %s" % v["갈래"])
    if v["왜"]:
        print("  %s" % v["왜"])
    for x in v["근거"]:
        꼬리 = (" · " + x["왜"]) if x["왜"] else ""
        print("  [%s] %s%s" % (x["무엇"], x["갈래"], 꼬리))
    if v["갈래"] == "가능":
        print("  → python collect_gate.py --run <스크립트> [인자...]")
    return 0 if v["갈래"] == "가능" else 3


def main(argv=None):
    ap = argparse.ArgumentParser(description="수집을 지금 해도 되나 (자원으로 판정)")
    ap.add_argument("--run", nargs=argparse.REMAINDER,
                    help="문을 통과하면 이 파이썬 스크립트를 돌린다 "
                         "(★ 뒤는 전부 자식 몫 — --why/--who 는 앞에 적는다)")
    ap.add_argument("--why", default="", help="무엇 때문에 수집하나 (--run 앞에)")
    ap.add_argument("--who", default="claude")
    a = ap.parse_args(argv)
    if a.run:
        # 조용히 고쳐 주지 않는다 — 자식이 제 `--why` 를 가질 수 있다(`[172]`).
        # 다만 말은 해 준다: 안 하면 자식이 죽고 사람은 이유를 못 찾는다(`[289]`).
        샌것 = [x for x in ("--why", "--who") if x in a.run]
        if 샌것:
            print("  ! %s 가 --run 뒤에 있어 **자식 명령으로 넘어갑니다**. "
                  "문에 주려면 --run 앞에 적으십시오." % " ".join(샌것))
        return run(a.run, a.why, a.who)
    return show()


if __name__ == "__main__":
    raise SystemExit(main())
