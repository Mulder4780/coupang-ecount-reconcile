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
"양보한 뒤 그 주인이 실제로 끝냈나"를 되묻는다. 못 끝냈으면 **헛양보**다.

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
def _load(path=MARK):
    """→ (표, 못읽은이유). **못 읽은 것을 빈 표로 치지 않는다**(`[169]`)."""
    if not os.path.exists(path):
        return {}, ""                     # 아직 한 번도 안 적었다 — 이것은 정상이다
    try:
        with io.open(path, encoding="utf-8") as f:
            d = json.load(f)
        return (d if isinstance(d, dict) else {}), ""
    except (OSError, ValueError) as exc:
        return {}, "%s 를 못 읽었다: %s" % (os.path.basename(path), exc)


def _save(d, path=MARK):
    """작게 쓰고 원자적으로 갈아끼운다. 읽는 쪽이 물고 있으면 물러서며 다시 건다."""
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


def audit(d=None, now=None, 도는중=None):
    """**양보한 뒤 그 주인이 정말 끝냈나.** 안 끝냈으면 그 일은 아무도 안 한 것이다.

    근거는 이 표 안에만 있다 — 주인도 `record_run` 을 남기기 때문이다. 밖의 파일을
    다시 뒤지면 판정이 두 곳이 되고, 여기서 비싼 탐색을 하면 회차마다 값을 치른다(`[168]`).

    ★ **'아직 하는 중'을 '끝내지 못했다'로 읽지 않는다.** 주인이 지금 락을 쥐고 있으면
      판정 자체를 미룬다 — 안 그러면 양보할 때마다 경보가 뜬다(실측: 첫 실행이 곧바로
      거짓 경보를 냈다). `도는중` 을 넘기면 그것을 쓰고, 안 넘기면 락을 읽는다.
    """
    if d is None:
        d, _ = _load()
    now = now or datetime.now()
    if 도는중 is None:
        도는중 = running()
    도는중 = set(도는중 or ())
    헛, 못검사 = [], []
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
                헛.append({"작업": 작업, "주인": 주인, "때": 행.get("때"),
                          "왜": 행.get("왜")})
    return 헛, 못검사


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

    헛, 못검사 = audit(d, now=now, 도는중=도는중)
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
