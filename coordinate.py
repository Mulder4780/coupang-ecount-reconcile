# -*- coding: utf-8 -*-
"""겹치는 일을 조율한다 — 양보는 **주장**이고, 주장은 나중에 검사한다 (2026-08-17 지시).

사용자 지시: **"다른 앱의 작업과 겹치면 서로 조율해서 진행하는 알고리즘 구현해"**

## 무엇이 문제였나 — 겹침의 결과가 **세 얼굴**이었다

실측 2026-08-17. `CSOS_유수비_대표보고_자동준비` 와 `쿠팡업무_일일자동대조` 가
**같은 초(09:50:02)에 같은 명령**(`daily_run.bat`)으로 뜬다. 그런데 겹쳤을 때 남는
자국이 회차마다 달랐다:

  · `daily_run` — 락에 지면 한 줄 찍고 **exit 0 · 자국 없음** → 스케줄러는 '성공'
  · `automation_pipeline` — `already_running` 을 돌려주지만 부르는 쪽이 안 적으면 사라짐
  · `noon_run` — 제대로 '양보'로 적는다(그 회차 **안에서만**)

그래서 밖에서 보면 **겹쳐서 안 돈 것과 정말 다 한 것이 구별되지 않는다.** 이 프로젝트가
반복해 당한 그 모양이다(`[169]`). 실제로 나는 `대표보고 exit 1` 을 '겹침'으로 읽고
"작업을 지울까요"까지 갔는데, 열어 보니 원인은 **합성검증 실패**였다
(`sys.exit("…")` 는 문자열을 줘도 **exit 1** 이다).

## 여기서 새 정책을 만들지 않는다 (`[162]`)

무엇과 부딪히면 어떻게 할지는 **회차가 안다** — `noon_run` 은 제 단계와 창을 알고,
`daily_run` 은 제 락을 안다. 그 판정을 여기로 옮기면 사본이 둘 된다. 이 파일이 맡는
것은 **이름과 기록** 하나다: 양보를 성공으로도 실패로도 적지 않고 **제 이름**으로 적고,
그것이 거듭되는지 세고, **주장이 지켜졌는지 검사한다.**

## 양보는 주장이다 — 그래서 검사한다

양보는 "**내가 안 해도 저쪽이 그 일을 한다**"는 주장이다. 주인마저 못 끝내면 그 일은
**아무도 안 한 것**이 되는데, 오늘은 아무 화면에도 안 뜬다 — 양보는 실패가 아니라서
경보가 없고, 주인의 죽음은 제 이름으로 따로 적히기 때문이다. 그래서 `audit()` 이
"양보한 뒤 그 일이 실제로 됐나"를 되묻는다. 아무도 안 했으면 **헛양보**다.

★ **묻는 것은 "주인이 했나" 가 아니라 "그 일이 됐나" 다.** 주인이 못 끝냈어도 양보한
쪽이 나중에 스스로 돌면 그 일은 된 것이다 — `뒤에됨`(`done_later`). 그것까지 헛양보라
부르면 표는 시간이 갈수록 더 확신에 차서 틀린다. 2026-08-19 실측: 수집이 14:05 에
양보하고 20:15~21:44 에 여덟 번 완주했는데도 표는 "아무도 안 했다"고 확언했다
(`schedule_watch` 가 `[304]` 에서 배운 것과 같은 자리다).

## 잘못 양보하는 쪽이 더 위험하다 (`[172]`)

못 잡는 것보다 나쁜 것은 **없는 주인에게 양보하는 것**이다 — 그러면 일은 영영 안 되고
기록은 '양보(정상)'로 남는다. 그래서 `record_yield` 는 **주인 이름을 반드시 받는다.**
모르면 `주인=""` 로 적히고 `audit()` 이 그것을 **검사할 수 없는 양보**로 따로 센다
(조용히 정상으로 세지 않는다).

읽기 전용이다 — 회차를 다시 띄우지도, 스케줄러를 고치지도 않는다. 자국만 남긴다.

  python coordinate.py            # 지금 상태
  python coordinate.py --print    # 리포트 본문
"""
import io
import json
import os
import sys
import time
from datetime import datetime, timedelta

if hasattr(sys.stdout, "reconfigure"):        # 무인 회차는 pythonw = stdout 이 None [235]
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

MARK = os.path.join(ROOT, "reports", "작업_조율.json")
REPORT_MD = os.path.join(ROOT, "reports", "작업_조율.md")

KEEP = 40                 # 작업마다 남기는 자국 수
STARVE_N = 3              # 연속 양보가 이만큼이면 굶주림
STARVE_HOURS = 24         # 그리고 이 시간 넘게 한 번도 못 돌았으면
# ★ 판정에 **창**을 둔다. 첫 실행이 곧바로 거짓 경보를 냈다 — `정오회차` 가 11:27 에
#   `일일대조` 에게 양보했는데 그 회차는 **지금도 도는 중**이었다. '아직 하는 중'과
#   '끝내지 못했다'를 안 가르면 양보할 때마다 매번 뜬다(`[170]`·`[172]`).
#   그래서 ① 주인이 지금 도는 중이면 **판정하지 않고** ② 실측 회차 길이(292분)보다
#   넉넉히 지난 뒤에야 묻는다.
AUDIT_GRACE_HOURS = 8     # 이만큼 지나기 전에는 아직 판정하지 않는다
AUDIT_HOURS = 36          # 그리고 이 시간이 지나면 이미 지나간 일이라 안 묻는다

# 자국의 갈래 — 여기 없는 낱말을 지어내지 않는다(읽는 쪽이 갈래로 가른다).
KINDS = ("완주", "실패", "양보", "막힘")

# 락 파일 → 그 일의 이름. **이름은 여기 한 곳에서 온다.**
# ⚠ 양보한 쪽과 완주한 쪽이 다른 이름을 쓰면 `audit()` 이 둘을 못 잇고 **모든 양보가
#   영영 '헛양보'** 로 보인다 — 오류도 안 나고 경보만 매일 뜬다(`[165]` 모양).
#   실제로 그랬다: `noon_run` 은 `일일대조(09:50)` 로, `daily_run` 은 `일일대조` 로 적었다.
LOCKS = (
    (".daily_run.lock", "일일대조"),
    (".automation_pipeline.lock", "증분 파이프라인"),
    (".ledger_db_apply.lock", "보관본 생성"),
)


def running(report_dir=None):
    """지금 락을 쥐고 있는 일들 → `[이름]`. 주인이 **살아 있을 때만** 센다.

    판정은 `pid_alive` 것을 그대로 빌린다(`[162]`). 못 재면(`None`) **살아 있는 것으로
    본다** — 의심스러우면 양보하는 쪽이 안전하다(같이 Z: 를 긁는 것보다 낫다).
    """
    out = []
    d = report_dir or os.path.join(ROOT, "reports")
    try:
        import pid_alive as P
    except Exception:                     # noqa: BLE001 — 못 물어보면 아무 이름도 못 댄다
        return out
    for 파일, 이름 in LOCKS:
        try:
            with io.open(os.path.join(d, 파일), encoding="utf-8") as f:
                owner = json.load(f)
        except (OSError, ValueError):
            continue
        if not isinstance(owner, dict):
            continue
        if P.owner_alive(owner.get("pid"), owner.get("pid_started_at")) is not False:
            out.append(이름)
    return out


# ── 자국 ────────────────────────────────────────────────────────────────────
def _load(path=None):
    """→ (표, 못읽은이유). **못 읽은 것을 빈 표로 치지 않는다**(`[169]`).

    ★ 경로는 **부를 때** 읽는다 — `path=MARK` 는 기본인자라 **def 시각에 묶인다.**
      그래서 `coordinate.MARK` 를 임시 경로로 돌려도 **진짜 파일에 그대로 썼다.**
      2026-08-19 실측: 합성검증이 `daily_run.main()` 을 부르면서 `PROGRESS` 만
      돌리고 조율은 못 돌려, 일일대조 자국 **40줄 중 38줄이 검증이 쓴 가짜**였다.
      그림이 나쁜 자리는 여기다 — `audit()` 은 "양보했는데 주인이 끝냈나"를
      주인의 **완주 줄**로 묻는다. 가짜 완주가 **진짜 헛양보를 지운다**(`[293]` 이
      잡으려던 바로 그것). 오류도 안 나고 표도 그럴듯하다(`[169]`).
    """
    path = path or MARK
    if not os.path.exists(path):
        return {}, ""                     # 아직 한 번도 안 적었다 — 이것은 정상이다
    try:
        with io.open(path, encoding="utf-8") as f:
            d = json.load(f)
        return (d if isinstance(d, dict) else {}), ""
    except (OSError, ValueError) as exc:
        return {}, "%s 를 못 읽었다: %s" % (os.path.basename(path), exc)


def _save(d, path=None):
    """작게 쓰고 원자적으로 갈아끼운다. 읽는 쪽이 물고 있으면 물러서며 다시 건다.

    경로는 `_load` 와 같은 이유로 **부를 때** 읽는다(위 설명).
    """
    path = path or MARK
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    for wait in (0, 0.2, 0.5, 1.0):
        if wait:
            time.sleep(wait)
        try:
            os.replace(tmp, path)
            return True
        except OSError:
            continue
    return False                          # 못 적었으면 못 적었다고 말한다(아래 note 가 삼키지 않는다)


#: 주인이 **회차가 아니라 점유 세션**일 때 붙는 표시.  회차에는 `record_run` 이 있어
#: '끝냈나'를 물을 수 있지만 점유 세션에는 그런 자국이 없다 — 그러므로 이것은
#: **검사할 수 없는 양보**이되 **주인을 모르는 양보는 아니다.**  둘을 뭉치면 화면이
#: "주인을 모른다"고 말하는데 실은 알고 있어서, 사람이 없는 문제를 찾아 나선다(`[172]`).
CLAIM_OWNER = "점유:"


def note(작업, 갈래, 왜="", 주인="", **extra):
    """자국 한 줄. **회차를 죽이지 않는다** — 조율을 적으려다 일을 막지 않는다.

    ★ 그러나 **조용히 삼키지도 않는다**: 못 적으면 `False` 를 돌려주고, 그 사실은
      다음 `notices()` 가 '확인 못 함' 으로 말한다. 실패를 성공처럼 보이게 하는 것이
      이 프로젝트가 제일 자주 당한 사고다.
    """
    작업 = str(작업 or "").strip() or "(이름 없음)"
    갈래 = str(갈래 or "").strip()
    if 갈래 not in KINDS:                 # 낱말을 지어내지 않는다
        갈래 = "막힘"
    행 = {"때": datetime.now().isoformat(timespec="seconds"),
          "갈래": 갈래, "왜": str(왜 or "")[:300], "주인": str(주인 or "")[:80]}
    for k, v in (extra or {}).items():
        행[str(k)] = v if isinstance(v, (int, float, bool)) else str(v)[:200]
    try:
        d, _why = _load()
        칸 = d.setdefault(작업, [])
        if not isinstance(칸, list):
            칸 = []
        칸.append(행)
        d[작업] = 칸[-KEEP:]
        return bool(_save(d))
    except Exception:                     # noqa: BLE001 — 어떤 일이 있어도 부른 쪽을 안 죽인다
        return False


def record_yield(작업, 주인, 왜=""):
    """겹쳐서 물러났다. **주인 이름을 반드시 받는다** — 없는 주인에게 양보하면
    그 일은 영영 안 되고 기록만 '정상'으로 남는다(`[172]` 방향)."""
    return note(작업, "양보", 왜, 주인=주인)


def record_run(작업, 상태, 왜=""):
    """실제로 돌았다. `상태` 는 회차가 쓰는 말 그대로(완주·실패·중단…)를 받되
    자국 갈래는 완주/실패 둘로만 접는다 — 읽는 쪽이 갈래로 가르기 때문이다."""
    말 = str(상태 or "")
    갈래 = "완주" if 말.startswith("완주") else "실패"
    return note(작업, 갈래, 왜 or 말, 상태=말)


# ── 읽기 ────────────────────────────────────────────────────────────────────
def _때(행):
    try:
        return datetime.fromisoformat(str(행.get("때") or ""))
    except (TypeError, ValueError):
        return None


def history(작업, d=None):
    """한 작업의 지금 상태 → 연속양보·마지막완주·마지막양보."""
    if d is None:
        d, _ = _load()
    칸 = [x for x in (d.get(작업) or []) if isinstance(x, dict)]
    연속 = 0
    for 행 in reversed(칸):
        if 행.get("갈래") == "양보":
            연속 += 1
        else:
            break
    def 마지막(갈래들):
        for 행 in reversed(칸):
            if 행.get("갈래") in 갈래들:
                return 행
        return None
    return {"작업": 작업, "자국수": len(칸), "연속양보": 연속,
            "마지막돎": 마지막(("완주", "실패")), "마지막양보": 마지막(("양보",)),
            "한번도안돎": not any(x.get("갈래") in ("완주", "실패") for x in 칸)}


def done_later(d, 작업, t):
    """양보한 **그 작업 자신**이 그 뒤에 스스로 완주했나 — 그러면 그 일은 됐다.

    양보의 뜻은 "내가 안 해도 저쪽이 그 일을 한다" 이지만, 되물어야 할 것은
    "주인이 했나" 가 아니라 **"그 일이 됐나"** 다. 주인이 못 끝냈어도 이쪽이
    나중에 스스로 돌았으면 그 일은 된 것이다 — 그때도 "아무도 안 했다"고 적으면
    사람이 **없는 일**을 찾아 나서고(`[172]`), 거짓 경보는 진짜 경보를 덮는다(`[170]`).
    `schedule_watch` 가 `[304]` 에서 배운 `뒤에됨` 을 이 표로 옮긴 것이다.

    실측 2026-08-19: 수집이 14:05 에 `claude[b0b8fd43]` 에게 양보했고 그 주인은
    끝냄 자국을 안 남겼다 — 그런데 **수집 자신이 20:15~21:44 에 여덟 번 완주**했고
    그중 둘이 `band/convert_dump.py`(바로 그 양보가 미룬 밴드 흡수)였다. 그런데도
    표는 "그 일은 아무도 안 했다" 고 확언하고 있었다.

    ★ **완주만 받는다** — 주인 쪽 검사는 실패도 끝냄으로 치는데 물음이 다르다.
      거기는 "양보의 전제(저쪽이 지금 그 일을 하고 있다)가 지켜졌나" 이고 여기는
      "그 일이 결국 됐나" 다. **실패한 회차는 그 일을 못 했다.**
    ★ **양보 시각 뒤**의 완주만 본다 — 앞의 완주는 그 양보와 아무 상관이 없다.
    """
    for 뒤 in (d.get(작업) or []):
        if not isinstance(뒤, dict) or 뒤.get("갈래") != "완주":
            continue
        t2 = _때(뒤)
        if t2 and t2 > t:
            return 뒤
    return None

def _can_ask(주인, d):
    """이 주인에게 **'끝냈나'를 물을 자국이 있나**.

    ★ 접두(`점유:` 같은 약속)로 가르지 않는다 — 새 접두가 생기는 날 그 갈래만
      조용히 헛양보로 떨어지면서 **오류는 안 난다**(`[165]`). 실측 2026-08-22:
      `collect_gate` 가 차선 주인을 `자료 수집 차선({'who': ...})` 로 적어
      **살아 있는 옆 창에게 양보한 것이 "그 일은 아무도 안 했다"로 확언**됐다.

    근거는 **구조**다: `audit()` 이 완주를 찾는 곳은 이 표의 작업 칸뿐이므로,
    그 칸을 가질 수 있는 이름만 물을 수 있다. 세션·사람은 `record_run` 을
    안 남기므로 **물을 자국이 애초에 없다** — 그것을 헛양보라 부르면 거짓이고,
    거짓 경보는 진짜 헛양보를 덮는다(`[170]`).

    ★ 회차에게 양보한 것은 그대로 검사된다 — 문을 넓힌 것이 아니라 근거를 바로잡은
      것이다. `LOCKS` 는 아직 자국이 없는 회차까지 받는다(락 파일이 곧 그 일의 정체다).
    """
    if not 주인:
        return False
    if 주인.startswith(CLAIM_OWNER):
        return False                      # 옛 규약도 그대로 지킨다
    if 주인 in (d or {}):
        return True
    return any(주인 == 이름 for _f, 이름 in LOCKS)


def audit(d=None, now=None, 도는중=None):
    """**양보한 뒤 그 일이 정말 됐나.** 아무도 안 했으면 그것이 헛양보다.

    돌려주는 것 넷: `헛양보 · 못검사 · 점유 · 뒤에됨`. 뜻도 조치도 다르므로
    한 통에 담지 않는다 — 뭉치면 경보가 대부분이 되고 아무도 안 본다(`[170]`).

    근거는 이 표 안에만 있다 — 주인도 `record_run` 을 남기기 때문이다. 밖의 파일을
    다시 뒤지면 판정이 두 곳이 되고, 여기서 비싼 탐색을 하면 회차마다 값을 치른다(`[168]`).

    ★ **'아직 하는 중'을 '끝내지 못했다'로 읽지 않는다.** 주인이 지금 락을 쥐고 있으면
      판정 자체를 미룬다 — 안 그러면 양보할 때마다 경보가 뜬다(실측: 첫 실행이 곧바로
      거짓 경보를 냈다). `도는중` 을 넘기면 그것을 쓰고, 안 넘기면 락을 읽는다.
    ★ **주인만 보면 안 된다** — 양보한 쪽이 나중에 스스로 돌면 그 일은 된 것이다
      (`done_later`). 그것까지 헛양보라 부르면 표는 시간이 갈수록 더 확신에 차서
      틀린다 — 실측으로 그렇게 되고 있었다.
    """
    if d is None:
        d, _ = _load()
    now = now or datetime.now()
    if 도는중 is None:
        도는중 = running()
    도는중 = set(도는중 or ())
    헛, 못검사, 점유, 뒤에 = [], [], [], []
    for 작업, 칸 in sorted(d.items()):
        for 행 in (칸 or []):
            if not isinstance(행, dict) or 행.get("갈래") != "양보":
                continue
            t = _때(행)
            if not t or now - t > timedelta(hours=AUDIT_HOURS):
                continue                  # 오래된 것은 이미 지나간 일이다
            if now - t < timedelta(hours=AUDIT_GRACE_HOURS):
                continue                  # 아직 물을 때가 아니다(회차 실측 292분)
            if str(행.get("주인") or "").strip() in 도는중:
                continue                  # 주인이 지금도 하는 중이다 — 실패가 아니다
            주인 = str(행.get("주인") or "").strip()
            if 주인 and not _can_ask(주인, d):
                # 주인은 안다 — 다만 세션·사람이라 '끝냈나'를 물을 자국이 없다.
                # 이것을 경보로 올리면 매일 뜨고, 매일 뜨는 경보는 아무도 안 본다(`[170]`).
                # 정말 위험한 것(매일 양보만 하는 회차)은 `굶주림` 이 따로 잡는다.
                점유.append({"작업": 작업, "주인": 주인, "때": 행.get("때"), "왜": 행.get("왜")})
                continue
            if not 주인:
                # ★ 주인을 모르는 양보는 **검사할 수 없다.** 정상으로 세지 않는다(`[169]`).
                못검사.append({"작업": 작업, "때": 행.get("때"), "왜": 행.get("왜")})
                continue
            끝냈나 = False
            for 뒤 in (d.get(주인) or []):
                if not isinstance(뒤, dict) or 뒤.get("갈래") not in ("완주", "실패"):
                    continue
                t2 = _때(뒤)
                if t2 and t2 >= t:
                    끝냈나 = True
                    break
            if not 끝냈나:
                # ★ 주인이 안 끝냈어도 **그 일이 됐을 수 있다** — 양보한 쪽이 나중에
                #   스스로 돌면 그 일은 된 것이다(`done_later`).
                뒤것 = done_later(d, 작업, t)
                if 뒤것 is not None:
                    뒤에.append({"작업": 작업, "주인": 주인, "때": 행.get("때"),
                                "왜": 행.get("왜"), "된때": 뒤것.get("때"),
                                "된일": 뒤것.get("왜")})
                    continue
                헛.append({"작업": 작업, "주인": 주인, "때": 행.get("때"),
                          "왜": 행.get("왜")})
    return 헛, 못검사, 점유, 뒤에


def notices(d=None, why="", now=None, 도는중=None):
    """인계·감시가 읽을 줄. **경보는 좁게, 못 읽음은 숨기지 않고**(`[169]`·`[170]`)."""
    if d is None:
        d, why = _load()
    now = now or datetime.now()
    out = []
    if why:
        out.append({"갈래": "확인못함", "작업": "", "무엇": why,
                    "어떻게": "python coordinate.py --print"})
        return out                        # 근거를 못 읽었으면 아무 판정도 하지 않는다

    헛, 못검사, 점유, _뒤에 = audit(d, now=now, 도는중=도는중)
    for x in 헛:
        out.append({"갈래": "헛양보", "작업": x["작업"],
                    "무엇": "**%s** 이 '%s' 에게 양보했는데 그 주인도 끝내지 못했다 — "
                            "그 일은 아무도 안 했다 (%s)" % (x["작업"], x["주인"],
                                                       str(x.get("때"))[:16].replace("T", " ")),
                    "어떻게": "python coordinate.py --print"})
    if 못검사:
        out.append({"갈래": "확인못함", "작업": 못검사[0]["작업"],
                    "무엇": "주인을 모르는 양보 %d건 — 그 일이 됐는지 검사할 수 없다"
                            % len(못검사),
                    "어떻게": "python coordinate.py --print"})

    for 작업 in sorted(d):
        h = history(작업, d)
        if h["연속양보"] < STARVE_N:
            continue
        마지막 = _때(h["마지막돎"] or {})
        # ★ **매일 양보만 하는 회차는 없는 회차와 같다** (noon_run 이 제 안에서 배운 것을
        #   모든 작업에 넓힌다). 양보는 실패가 아니라서 아무 경보도 안 뜨는 자리다.
        if h["한번도안돎"]:
            out.append({"갈래": "굶주림", "작업": 작업,
                        "무엇": "**%s** 이 %d회 연속 양보만 했고 **한 번도 돈 적이 없다** — "
                                "겹치는 쪽을 옮기거나 이 작업이 필요한지 다시 본다"
                                % (작업, h["연속양보"]),
                        "어떻게": "python coordinate.py --print"})
        elif 마지막 and now - 마지막 > timedelta(hours=STARVE_HOURS):
            out.append({"갈래": "굶주림", "작업": 작업,
                        "무엇": "**%s** 이 %d회 연속 양보했고 마지막으로 돈 것은 %s 다"
                                % (작업, h["연속양보"],
                                   마지막.isoformat(timespec="minutes").replace("T", " ")),
                        "어떻게": "python coordinate.py --print"})
    return out


# ── 리포트 ──────────────────────────────────────────────────────────────────
def _md(d, why, now=None):
    now = now or datetime.now()
    L = ["# 겹치는 일 조율", "",
         "- 잰 때: %s · 작업 %d개" % (now.isoformat(timespec="seconds"), len(d)), ""]
    ns = notices(d, why, now=now)
    if ns:
        L += ["## 먼저 볼 것", ""] + ["- [%s] %s" % (n["갈래"], n["무엇"]) for n in ns] + [""]
    else:
        L += ["경보로 올릴 것 없음. — **자국이 없는 것과 겹침이 없는 것은 다르다**;"
              " 아래 표에 작업이 안 보이면 그 회차는 아직 조율을 안 적는다.", ""]
    L += ["## 작업마다", "", "| 작업 | 자국 | 연속양보 | 마지막 돎 | 마지막 양보 |",
          "|---|---:|---:|---|---|"]
    for 작업 in sorted(d):
        h = history(작업, d)
        돎 = h["마지막돎"] or {}
        양 = h["마지막양보"] or {}
        L.append("| %s | %d | %d | %s | %s |" % (
            작업, h["자국수"], h["연속양보"],
            (str(돎.get("때") or "")[:16].replace("T", " ") + " " + str(돎.get("갈래") or "")).strip() or "없음",
            (str(양.get("때") or "")[:16].replace("T", " ") + " " + str(양.get("주인") or "")).strip() or "없음"))
    L.append("")
    # ★ **뺀 것은 숫자로 말한다**([169]).  점유에 양보한 것은 경보로 안 올리지만
    #   조용히 없애지도 않는다 — 안 보이면 '겹친 적 없다'로 읽힌다.
    try:
        _헛, _못, _점, _뒤 = audit(d, now=now)
        if _점:
            L += ["", "## 세션·사람에게 양보한 것 (경보 아님)", "",
                  "- %d건 — 주인은 **안다**(점유·차선을 쥔 창). 다만 그 창에는 `record_run` 이"
                  " 없어 '끝냈나'를 물을 자국이 없다." % len(_점), ""]
            for x in _점[-5:]:
                L.append("  - %s · %s -> %s" % (str(x.get("때"))[:16].replace("T", " "),
                                                x.get("작업"), x.get("주인")))
            L.append("")
        # ★ **뒤에 스스로 된 것도 숫자로 말한다**(`[169]`). 경보에서는 내리지만
        #   조용히 없애면 "겹친 적 없다" 로 읽힌다 — 양보는 실제로 있었고 그 일은
        #   나중에 됐다는 것이 사실이다.
        if _뒤:
            L += ["", "## 뒤에 스스로 된 것 (경보 아님)", "",
                  "- %d건 — 주인은 끝냄 자국을 안 남겼지만 **그 작업 자신이 나중에**"
                  " **완주**했다. 그 일은 됐다 — 아무도 안 한 것이 아니다(`[304]`)."
                  % len(_뒤), ""]
            for x in _뒤[-5:]:
                L.append("  - %s %s 에게 양보 → %s %s" % (
                    str(x.get("때"))[:16].replace("T", " "), x.get("주인"),
                    str(x.get("된때"))[:16].replace("T", " "), str(x.get("된일") or "")[:60]))
            L.append("")
    except Exception:                      # 표를 못 그려도 리포트를 세우지 않는다
        pass
    return "\n".join(L)


def write_report():
    d, why = _load()
    os.makedirs(os.path.dirname(REPORT_MD), exist_ok=True)
    with io.open(REPORT_MD, "w", encoding="utf-8") as f:
        f.write(_md(d, why))
    return REPORT_MD


def main():
    d, why = _load()
    본문 = "--print" in sys.argv
    write_report()
    if 본문:
        print(_md(d, why))
        return
    ns = notices(d, why)
    if ns:
        for n in ns:
            print("[%s] %s" % (n["갈래"], n["무엇"]))
    else:
        print("겹침 경보 없음 — 작업 %d개 (상세 %s)"
              % (len(d), os.path.relpath(REPORT_MD, ROOT)))


if __name__ == "__main__":
    main()
