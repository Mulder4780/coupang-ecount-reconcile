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

# ★ 갈래마다 따로 끈다 (2026-09-08 형님 지시: "내가 자료 긁어오라고할 때만 긁어오고
#   수시로 긁어오는 건 다 멈춰").  기본값이 "band" 라 **옛 호출자는 한 글자도
#   안 바뀐다**([172]) — 넓히는 것이지 옛 동작을 바꾸는 것이 아니다.
#   ★ 표시 파일과 환경변수를 갈래마다 따로 두는 이유: 한 파일에 담으면 밴드를
#     다시 켜는 순간 ERP 까지 같이 켜진다.  끄고 켜는 값이 갈래마다 다르다.
KINDS = {
    "band": (MARK, ENV, "밴드"),
    "erp": (os.path.join(ROOT, "reports", "ERP수집_중단.json"),
            "COUPANG_ERP_COLLECT", "ERP"),
}


def _kind(kind):
    """갈래를 (표시파일, 환경변수, 이름) 으로 푼다.

    ★ 모르는 갈래면 **예외를 올린다**([165]) — 조용히 False 를 주면 오타 하나로
      그 갈래만 영영 안 막히면서 오류도 안 난다.  부르는 쪽은 어차피 try 로
      감싸 안전한 쪽(중단 아님)으로 떨어지므로 동작은 안 나빠지고, 검증이
      이 예외로 갈래 이름이 어긋난 것을 잡는다.
    """
    try:
        return KINDS[kind]
    except KeyError:
        raise ValueError("모르는 수집 갈래: %r (아는 것: %s)"
                         % (kind, ", ".join(sorted(KINDS))))


def _read(kind="band"):
    mark = _kind(kind)[0]
    try:
        with io.open(mark, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else None
    except Exception:
        return None


def stopped(kind="band"):
    """(중단인가, 왜) — 못 읽으면 (False, 왜못함).

    ★ 못 읽으면 **중단 아님**이다([169] 를 이 자리에 맞게 정한 것).  표시가
      깨졌다고 수집을 막으면 나중에 왜 안 되는지 아무도 모른다.  물러나는 값은
      헛 수집 한 번이고, 잘못 막는 값은 **자료를 영영 못 받는 것**이다.
    """
    mark, env_name, label = _kind(kind)
    env = os.environ.get(env_name)
    if env is not None:
        if env == "0":
            return True, "환경변수 %s=0" % env_name
        return False, "환경변수 %s=%s (켬)" % (env_name, env)
    d = _read(kind)
    if d is None:
        return False, ""
    if not d.get("중단"):
        return False, "중단 표시 없음"
    why = str(d.get("왜") or "").strip()
    when = str(d.get("언제") or "").strip()
    say = "%s 자동 수집 중단" % label
    if when:
        say += "(%s 지시)" % when
    if why:
        say += " — " + why
    return True, say


def note(kind="band"):
    """사람·로그에 한 줄로 적을 말.  중단이 아니면 빈 문자열."""
    off, why = stopped(kind)
    return why if off else ""


def warning_status(kind="band"):
    """경보도 수집과 같은 스위치를 본다. 못 읽으면 기존 경보를 유지한다([361])."""
    try:
        off, why = stopped(kind)
    except Exception:
        off, why = False, "중단 설정 확인 못 함"
    try:
        env_name = _kind(kind)[1]
    except ValueError:
        env_name = ENV
    tail = "" if kind == "band" else " --kind %s" % kind
    return {"수집중단": bool(off), "왜": why,
            "재개": "python band/collect_switch.py%s --resume (또는 %s=1)"
                    % (tail, env_name)}


def stop(why="", when="", write=True, kind="band"):
    """중단으로 적는다.  **사람이 명령할 때만** 부른다."""
    mark = _kind(kind)[0]
    d = {"중단": True, "왜": why, "언제": when}
    if not write:
        return d
    os.makedirs(os.path.dirname(mark), exist_ok=True)
    tmp = mark + ".tmp"
    with io.open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(json.dumps(d, ensure_ascii=False, indent=1))
    os.replace(tmp, mark)
    return d


def resume(write=True, kind="band"):
    """다시 켠다 — 표시를 지운다."""
    mark = _kind(kind)[0]
    if not write:
        return True
    try:
        os.remove(mark)
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
    # --kind 는 어디에 있어도 받는다 — 앞에 적든 뒤에 적든 같은 뜻이다.
    kind = "band"
    if "--kind" in a:
        i = a.index("--kind")
        if i + 1 < len(a):
            kind = a[i + 1]
        del a[i:i + 2]
    if "--all" in a:
        a.remove("--all")
        kinds = sorted(KINDS)
    else:
        kinds = [kind]
    try:
        for k in kinds:
            _kind(k)
    except ValueError as e:
        print(e)
        return 2
    for k in kinds:
        label = _kind(k)[2]
        if a and a[0] == "--stop":
            why = a[1] if len(a) > 1 else ""
            when = a[2] if len(a) > 2 else ""
            stop(why, when, kind=k)
            print("%s 자동 수집을 중단으로 적었습니다 — %s" % (label, _kind(k)[0]))
        elif a and a[0] == "--resume":
            resume(kind=k)
            print("%s 자동 수집을 다시 켰습니다 — 표시를 지웠습니다" % label)
        off, why = stopped(k)
        print("지금[%s]: %s%s" % (label, "중단" if off else "켬",
                                (" · " + why) if why else ""))
    print("★ 끄는 것은 **긁는 것**뿐입니다 — 흡수·대조·정리는 그대로 돕니다.")
    print("★ 사람이 직접 돌리는 길은 안 막습니다"
          " (collect_gate --run · browser_chain --manual erp · erp_api_collect --force).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
