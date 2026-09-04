# -*- coding: utf-8 -*-
"""밴드 수집 스위치 — **긁는 것만** 끈다 (2026-09-01 지시).

형님 지시: **"밴드 자동 수집은 앞으로 하지마 이제 밴드에 이제 자료 안올라올거야"**

★ **끄는 것은 '밴드에 접속해 글을 가져오는 것' 하나다.**
  이미 받아 둔 덤프를 흡수하고(`convert_dump`) 대조하고 정리하는 길은
  **한 글자도 안 막는다** — 막으면 오늘 받은 85건이 반영이 안 되고 카톡·원장
  대조까지 같이 죽는다([172] — 좁히는 것도 고장이다).

★ **끄는 자리는 하나다**([162]). 실측 2026-09-01 로 자동으로 긁는 길이 셋이다:
    ① `watchdog.heal_band_bridge`(30분) — 사람 탭이 없으면 다리가 대신 긁는다([460])
    ② `browser_chain` 의 `band-*` 갈래(12:00) — 전면 크롬 몰이([269])
    ③ `app_server /api/collect_plan` — 브라우저 유저스크립트가 받아 가는 계획([182])
  각자 끄면 사본이 셋이 되고, **넷째가 생기는 날 그것만 조용히 샌다**([165]).

★ **공식 API(`band_sync`)는 원래 안 돈다** — 토큰 파일이 없다(실측). 그리고
  2026-08-24 에 폐기가 확인됐다([426]). 그래서 여기서 안 막는다([172]).

★ **사람이 손으로 부르는 길은 안 막는다.** `collect_gate --run` 으로 직접
  돌리는 것은 형님이 그때 판단해 시키신 것이다 — 막으면 나중에 밴드를 다시
  봐야 할 때 **길이 아예 없어진다**. 막는 것은 **자동으로 도는 것**뿐이다.

★ **되돌리기 두 길**(둘 다 되고, 환경변수가 더 세다):
    · `reports/밴드수집_중단.json` 을 지우거나 `{"중단": false}` 로 바꾼다
    · 환경변수 `COUPANG_BAND_COLLECT=1` ([126] 과 같은 보호장치)

⚠ **못 읽으면 '중단 아님'이다**([169] 를 이 자리에 맞게 정한 것). 파일이 깨졌다는
  이유로 수집을 막으면, 나중에 밴드를 다시 켜야 할 때 **왜 안 되는지 아무도 모른다.**
  물러나는 값은 헛 수집 한 번이고, 잘못 막는 값은 **자료를 영영 못 받는 것**이다.
"""
import io
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MARK = os.path.join(ROOT, "reports", "밴드수집_중단.json")
ENV = "COUPANG_BAND_COLLECT"


def _read():
    try:
        with io.open(MARK, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else None
    except Exception:
        return None


def stopped():
    """(중단인가, 왜) — 못 읽으면 (False, 왜못함)."""
    env = os.environ.get(ENV)
    if env is not None:
        if env == "0":
            return True, "환경변수 %s=0" % ENV
        return False, "환경변수 %s=%s (켬)" % (ENV, env)
    d = _read()
    if d is None:
        return False, ""
    if not d.get("중단"):
        return False, "중단 표시 없음"
    why = str(d.get("왜") or "").strip()
    when = str(d.get("언제") or "").strip()
    say = "밴드 자동 수집 중단"
    if when:
        say += "(%s 지시)" % when
    if why:
        say += " — " + why
    return True, say


def note():
    """사람·로그에 한 줄로 적을 말.  중단이 아니면 빈 문자열."""
    off, why = stopped()
    return why if off else ""


def warning_status():
    """경보도 수집과 같은 스위치를 본다. 못 읽으면 기존 경보를 유지한다([361])."""
    try:
        off, why = stopped()
    except Exception:
        off, why = False, "중단 설정 확인 못 함"
    return {"수집중단": bool(off), "왜": why,
            "재개": "python band/collect_switch.py --resume (또는 COUPANG_BAND_COLLECT=1)"}


def stop(why="", when="", write=True):
    """중단으로 적는다.  **사람이 명령할 때만** 부른다."""
    d = {"중단": True, "왜": why, "언제": when}
    if not write:
        return d
    os.makedirs(os.path.dirname(MARK), exist_ok=True)
    tmp = MARK + ".tmp"
    with io.open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(json.dumps(d, ensure_ascii=False, indent=1))
    os.replace(tmp, MARK)
    return d


def resume(write=True):
    """다시 켠다 — 표시를 지운다."""
    if not write:
        return True
    try:
        os.remove(MARK)
    except OSError:
        pass
    return True


def main():
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    a = sys.argv[1:]
    if a and a[0] == "--stop":
        why = a[1] if len(a) > 1 else ""
        when = a[2] if len(a) > 2 else ""
        stop(why, when)
        print("밴드 자동 수집을 중단으로 적었습니다 —", MARK)
    elif a and a[0] == "--resume":
        resume()
        print("밴드 자동 수집을 다시 켰습니다 — 표시를 지웠습니다")
    off, why = stopped()
    print("지금: %s%s" % ("중단" if off else "켬", (" · " + why) if why else ""))
    print("★ 끄는 것은 **긁는 것**뿐입니다 — 흡수·대조·정리는 그대로 돕니다.")
    print("★ 사람이 `collect_gate --run` 으로 직접 돌리는 길은 안 막습니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
