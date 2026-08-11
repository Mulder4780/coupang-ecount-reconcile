# -*- coding: utf-8 -*-
"""회차가 **정말 돌았나** — 스케줄러 결과를 읽는 눈 (2026-08-12, 분담판 [35]).

★ **자동화의 마지막 구멍이 여기였다.** 코드를 만들고 설치본을 만들고 지시문의
  '자동으로 도는 것' 목록에 한 줄을 적으면 자동이 된 것처럼 보인다. 그러나 실제로
  도는지를 아는 것은 그 목록이 아니라 **작업 스케줄러**다. 실측 2026-08-12:
  · `쿠팡업무_정오회차` — 코드·검증·설치본이 8/11 에 다 있었는데 **등록이 안 돼
    한 번도 안 돌았다.** 지시문에는 "12:00 에 돈다"고 적혀 있었다.
  · `쿠팡업무_일일자동대조`·`쿠팡업무_원본자료자동정리` — 마지막 결과 `0xC000013A`.
    제한시간(PT3H)에 걸려 **매일 강제 종료**되고 있었다(실측 회차 292.3분).
  · `쿠팡업무_밴드재수집` — exit 1.
  셋 다 **아무 화면에도 안 떴다.** 프로젝트 어디에도 `LastTaskResult` 를 읽는 코드가
  한 줄도 없었기 때문이다(실측 `grep -rl LastTaskResult` → 0건).

★ **읽기 전용이다.** 회차를 다시 띄우지도, 제한시간을 고치지도, 등록하지도 않는다.
  옳은 조치는 회차마다 다르고 — 지금 도는 중이면 **띄우면 안 된다**(잠금에 막혀
  조용히 건너뛰거나 같은 자료를 두 번 긁는다) — 그 판단은 사람과 기존 도구 몫이다.
  여기는 **보고 말하는 것까지**다(`typo_watch` 와 같은 자리).

★ **가르는 것이 세는 것보다 어렵다**(`[170]`). 회차가 열둘이면 '0 이 아닌 결과'는 늘
  몇 개씩 있다. 통째로 경보로 올리면 경보가 대부분이 되고, 그러면 아무도 안 본다:
    · **성공·도는중** — 정상이다. 5분 회차가 지금 도는 것은 사고가 아니다.
    · **밀림** — 앞 회차가 아직 돌아 새 회차가 거부됐다(`IgnoreNew`). 한 번은 정상이고
      **연속으로 이어질 때만** 경보다(그때는 앞 회차가 안 끝나고 있다는 뜻이다).
    · **죽음** — 강제 종료·비정상 종료. 한 번이라도 경보다.
    · **안 돎** — 예정 시각이 지났는데 그 뒤로 실행 기록이 없다. **등록 안 됨**도 여기다.

★ **못 본 것을 정상이라 하지 않는다**(`[169]`). 조회가 실패하면 '이상 없음'이 아니라
  **'확인 못 함'** 이다. 이 구별이 없으면 감시자 자신이 눈먼 채로 "모두 정상"을
  말하게 된다 — 이 파일이 막으려는 바로 그 사고다.

★ **컴팩팅 배선도 같이 본다.** 회차가 아니라 세션 쪽이지만 고장 모양이 똑같다:
  `.claude/settings.json` 에서 `autoCompactEnabled`·PreCompact 훅이 빠지면 세션이
  가득 찬 순간 **인계 없이 끊기는데** 그전까지 아무 화면에도 안 뜬다.

쓰는 법
    python schedule_watch.py            # 지금 상태를 보고 리포트를 쓴다
    python schedule_watch.py --print    # 마지막 결과만 (스케줄러를 다시 안 묻는다)
검증 `[228]`.
"""
from __future__ import annotations

import base64
import json
import os
import re
import sys
from datetime import datetime, timedelta

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
if hasattr(sys.stdout, "reconfigure"):              # 콘솔이 cp949 라도 판정문이 안 깨진다
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

STATE = os.path.join(ROOT, "reports", "스케줄러_회차감시.json")
REPORT = os.path.join(ROOT, "reports", "스케줄러_회차감시.md")

#: 이 프로젝트가 등록하는 작업 이름의 머리. 남의 작업(Microsoft 등)은 안 본다.
PREFIXES = ("쿠팡업무_", "CSOS")

#: 스케줄러가 돌려주는 결과 코드 → (갈래, 사람 말). 16진수 그대로 두면 아무도 안 읽는다.
RESULT = {
    0:          ("성공", "정상으로 끝났다"),
    267009:     ("도는중", "지금 돌고 있다"),                       # 0x00041301
    267010:     ("꺼짐", "트리거가 꺼져 있다"),                      # 0x00041302
    267011:     ("아직안돎", "등록된 뒤 한 번도 실행된 적이 없다"),      # 0x00041303
    267012:     ("예정없음", "남은 예정이 없다"),                     # 0x00041304
    267014:     ("중단됨", "끝내기 요청을 받고 멈췄다"),               # 0x00041306
    2147946720: ("밀림", "앞 회차가 아직 돌아 새 회차가 거부됐다(IgnoreNew)"),  # 0x800710E0
    3221225786: ("강제종료", "제한시간에 걸려 나무째 끊겼다"),          # 0xC000013A
    3221225477: ("비정상종료", "접근 위반으로 죽었다"),                # 0xC0000005
}

#: 한 번이라도 보이면 경보. **죽은 것은 다음 회차가 대신 해 주지 않는다.**
#: `확인못함` 이 여기 있는 이유 — 못 본 것을 정상이라 하지 않는다(`[169]`).
DEAD = ("강제종료", "비정상종료", "중단됨", "실패", "확인못함")
#: 반복될 때만 경보. 한 번의 밀림은 정상 운영이다.
REPEAT_ONLY = ("밀림",)
#: 연속 몇 번부터 밀림을 경보로 올리나 — 5분 회차가 세 번 연속 거부면 앞이 안 끝난 것이다.
REPEAT_LIMIT = 3


# ────────────────────────────────────────────────────────────── 스케줄러에 묻기
# TaskPath '\' 만 본다 — 우리 작업은 전부 뿌리에 있고, Microsoft 것 수백 개를 훑으면
# 느리다. 필터에 한글을 쓰지 않는 이유는 설치본들과 같다(명령줄 코드페이지).
#
# ★ 결과 코드는 **반드시 `[long]`** 이다. `[int]` 로 받으면 `0xC000013A`(3221225786)
#   ·`0x800710E0`(2147946720) 이 Int32 를 넘겨 변환이 터지고, 그 작업 하나가
#   **통째로 목록에서 빠진다.** 실측(이 파일 첫 실행): 12개 중 9개만 돌아왔고 빠진
#   셋이 하필 `일일자동대조`·`원본자료자동정리`·`AutomationPipeline` — **정확히 지금
#   실패하고 있는 회차들**이었다. 게다가 빠진 자리가 '등록 안 됨'으로 읽혀 멀쩡한
#   작업에 거짓 경보까지 났다. 감시자가 실패한 것만 골라 못 보는 자리다(`[169]`).
# ★ 그래서 작업마다 try/catch 로 감싼다 — 무엇이 잘못돼도 **행은 반드시 하나 나온다.**
#   조용히 빠지느니 '확인못함'이라고 적힌 행이 낫다.
_PS = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$out = @()
foreach ($t in (Get-ScheduledTask -TaskPath '\' -ErrorAction SilentlyContinue)) {
  try {
    $i = $t | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
    $trg = @()
    foreach ($g in $t.Triggers) {
      $trg += [pscustomobject]@{
        kind  = [string]$g.CimClass.CimClassName
        start = [string]$g.StartBoundary
        every = [string]$g.Repetition.Interval
        span  = [string]$g.Repetition.Duration
        days  = [string]$g.DaysInterval
        on    = [bool]$g.Enabled
      }
    }
    $out += [pscustomobject]@{
      name   = [string]$t.TaskName
      state  = [string]$t.State
      reg    = [string]$t.Date
      limit  = [string]$t.Settings.ExecutionTimeLimit
      multi  = [string]$t.Settings.MultipleInstances
      last   = [string]$(if ($i -and $i.LastRunTime) { $i.LastRunTime.ToString('s') } else { '' })
      next   = [string]$(if ($i -and $i.NextRunTime) { $i.NextRunTime.ToString('s') } else { '' })
      result = [long]$(if ($i) { $i.LastTaskResult } else { -1 })
      missed = [long]$(if ($i) { $i.NumberOfMissedRuns } else { 0 })
      err    = ''
      trig   = $trg
    }
  } catch {
    $out += [pscustomobject]@{
      name = [string]$t.TaskName; state = [string]$t.State; reg = ''; limit = '';
      multi = ''; last = ''; next = ''; result = -1; missed = 0;
      err = [string]$_.Exception.Message; trig = @()
    }
  }
}
ConvertTo-Json -InputObject @($out) -Depth 5 -Compress
"""


def query(timeout=90):
    """스케줄러에게 묻는다. **못 물으면 빈 목록이 아니라 예외다** — 조회 실패를
    '작업 없음'으로 돌려주면 감시자가 눈먼 채 "모두 정상"을 말한다(`[169]`)."""
    import proc_guard
    enc = base64.b64encode(_PS.encode("utf-16-le")).decode("ascii")
    res = proc_guard.run_tree(
        ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", enc],
        cwd=ROOT, timeout=timeout)
    if res.timed_out:
        raise RuntimeError("스케줄러 조회가 %d초를 넘겼다" % timeout)
    txt = (res.stdout or "").strip()
    if res.returncode != 0 and not txt:
        raise RuntimeError("스케줄러 조회 실패(rc=%s): %s"
                           % (res.returncode, (res.stderr or "")[:120]))
    if not txt:
        raise RuntimeError("스케줄러가 아무 답도 주지 않았다")
    data = json.loads(txt)
    if isinstance(data, dict):                      # 한 개면 ConvertTo-Json 이 객체를 준다
        data = [data]
    return [t for t in data
            if any(str(t.get("name", "")).startswith(p) for p in PREFIXES)]


# ─────────────────────────────────────────────────── 설치본이 선언한 '있어야 할 것'
_CHAR = re.compile(r"\[char\]\s*0x([0-9A-Fa-f]{1,6})")
_PLAIN = re.compile(r'^\s*\$TaskName\s*=\s*"([^"]+)"', re.M)
_JOIN = re.compile(r"\$TaskName\s*=\s*-join\s*@\((.*?)\)", re.S)


def declared():
    """`install_*.ps1` 이 **선언한** 작업 이름들. 목록을 손으로 적지 않는 이유는
    적는 순간 사본이 둘이 되어, 설치본만 늘고 여기는 안 늘면 새 회차가 등록 안 된 채
    조용히 빠지기 때문이다 — 그것이 정오회차 사고의 모양이다."""
    names = {}
    for fn in sorted(os.listdir(ROOT)):
        if not (fn.startswith("install_") and fn.endswith(".ps1")):
            continue
        try:
            src = open(os.path.join(ROOT, fn), encoding="utf-8", errors="replace").read()
        except OSError:
            continue
        m = _PLAIN.search(src)
        if m:
            names[m.group(1)] = fn
            continue
        m = _JOIN.search(src)
        if m:
            got = "".join(chr(int(h, 16)) for h in _CHAR.findall(m.group(1)))
            if got:
                names[got] = fn
    return names


# ─────────────────────────────────────────────────────────── 예정 시각 계산
def _dur(text):
    """ISO8601 기간(`PT3H`·`PT10M`·`P1D`) → timedelta. 못 읽으면 None."""
    if not text:
        return None
    m = re.match(r"^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$", str(text).strip())
    if not m:
        return None
    d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    out = timedelta(days=d, hours=h, minutes=mi, seconds=s)
    return out or None


def _dt(text):
    if not text:
        return None
    t = str(text).strip()
    m = re.match(r"^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2}:\d{2})", t)
    if not m:
        return None
    try:
        got = datetime.strptime(m.group(1) + "T" + m.group(2), "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None
    # 윈도우는 '한 번도 안 돎'을 1999-11-30 으로 적는다. 그 날짜를 실행 기록으로 읽으면
    # "26년 전에 돌았다"가 되어 밀림 판정이 엉킨다 — 없는 것은 없다고 한다.
    return got if got.year >= 2000 else None


def _trigger_due(trg, now):
    """이 트리거의 **마지막 예정 시각**. 알 수 없으면 None.

    ★ 부팅·로그온·등록 트리거는 예정 시각이 없다 — **모르면 모른다고 한다.**
      억지로 '지금쯤'이라 치면 멀쩡한 회차가 매번 밀림으로 나온다(`[172]` 의 문).
    """
    if not trg.get("on", True):
        return None
    start = _dt(trg.get("start"))
    if not start:
        return None
    kind = str(trg.get("kind") or "")
    every, span = _dur(trg.get("every")), _dur(trg.get("span"))
    if kind.endswith("DailyTrigger"):
        try:
            days = max(1, int(trg.get("days") or 1))
        except (TypeError, ValueError):
            days = 1
        base = datetime.combine(now.date(), start.time())
        if base > now:
            base -= timedelta(days=days)
        if base < start:
            return None                              # 아직 첫 회차가 오지 않았다
    elif kind.endswith("TimeTrigger"):
        base = start
        if base > now:
            return None
    else:
        return None
    if every:                                        # 창 안에서 반복하는 회차
        end = base + span if span else now
        t = min(now, end)
        if t >= base:
            base = base + every * int((t - base) // every)
    return base


def _due(task, now):
    """모든 트리거 중 **가장 최근** 예정 시각과 그 회차의 여유(grace)."""
    best, grace = None, timedelta(hours=3)
    for trg in (task.get("trig") or []):
        d = _trigger_due(trg, now)
        if d is None:
            continue
        if best is None or d > best:
            best = d
            every = _dur(trg.get("every"))
            # 반복 회차는 간격의 세 배까지 봐준다(한 번 놓친 것은 사고가 아니다).
            # 일 단위는 3시간 — `-StartWhenAvailable` 이라 PC 가 자고 있었으면 늦게 돈다.
            grace = max(every * 3, timedelta(minutes=30)) if every else timedelta(hours=3)
    return best, grace


# ─────────────────────────────────────────────────────────────────── 판정
def judge(task, now, before=None):
    """작업 하나를 갈래로 나눈다. `before` 는 지난 회차의 같은 작업 기록(연속 세기용)."""
    name = task.get("name", "")
    code = int(task.get("result", -1) or 0)
    state = str(task.get("state") or "")
    if task.get("err"):                              # 이 작업 하나를 못 읽었다 ≠ 정상이다
        return {"작업": name, "갈래": "확인못함", "말": str(task["err"])[:120],
                "코드": code, "상태": state, "마지막실행": "", "다음예정": "",
                "제한시간": "", "놓친횟수": 0, "예정": "", "연속": 1,
                "처음본때": now.strftime("%Y-%m-%dT%H:%M:%S")}
    kind, say = RESULT.get(code, ("실패", "0 이 아닌 값으로 끝났다"))
    if code not in RESULT:
        say = "%s (코드 %d · 0x%08X)" % (say, code, code & 0xFFFFFFFF)
    if state == "Running" and kind not in ("도는중",):
        kind, say = "도는중", "지금 돌고 있다(직전 결과는 %s)" % kind
    if state == "Disabled":
        kind, say = "꺼짐", "작업이 꺼져 있다 — 사람이 껐을 수 있다"

    last, reg = _dt(task.get("last")), _dt(task.get("reg"))
    due, grace = _due(task, now)
    late = None
    if kind not in ("도는중", "꺼짐") and due is not None:
        # 등록 이전의 예정은 없던 것이다 — 오늘 등록한 회차를 어제치로 나무라지 않는다.
        if reg is None or due >= reg:
            if last is None or last < due - grace:
                late = due
    if late is not None:
        kind = "안돎"
        say = ("예정 %s 가 지났는데 그 뒤로 실행 기록이 없다(마지막 실행 %s)"
               % (late.strftime("%m-%d %H:%M"),
                  last.strftime("%m-%d %H:%M") if last else "없음"))

    # ★ 연속은 **회차**를 센다 — 관찰이 아니다. 워치독이 30분마다 보므로 관찰을 세면
    #   하루 한 번 실패하는 회차가 "48회 연속"이 되고, 밀림 판정(3회)은 첫날 아침에
    #   터진다. 같은 회차를 다시 본 것인지는 **마지막 실행 시각**이 말해 준다.
    prev = (before or {}).get(name) or {}
    same, moved = prev.get("갈래") == kind, (prev.get("마지막실행") or "") != (task.get("last") or "")
    if not same:
        run = 1
    elif moved:
        run = int(prev.get("연속") or 0) + 1
    else:
        run = max(1, int(prev.get("연속") or 0))
    return {
        "작업": name, "갈래": kind, "말": say, "코드": code, "상태": state,
        "마지막실행": last.strftime("%Y-%m-%dT%H:%M:%S") if last else "",
        "다음예정": task.get("next") or "",
        "제한시간": task.get("limit") or "", "놓친횟수": int(task.get("missed") or 0),
        "예정": late.strftime("%Y-%m-%dT%H:%M:%S") if late else "",
        "연속": run, "처음본때": prev.get("처음본때") if run > 1 else now.strftime("%Y-%m-%dT%H:%M:%S"),
    }


def alarms(rows, missing):
    """경보로 올릴 것만 고른다 — 여기가 좁아야 사람이 읽는다(`[170]`)."""
    out = []
    for name, fn in sorted(missing.items()):
        out.append({"갈래": "등록안됨", "작업": name,
                    "무엇": "**%s** — 설치본 `%s` 은 있는데 스케줄러에 없다. "
                            "코드가 있는 것과 도는 것은 다른 말이다" % (name, fn),
                    "어떻게": "powershell -ExecutionPolicy Bypass -File ecount\\%s" % fn})
    for r in rows:
        why = None
        if r["갈래"] in DEAD:
            why = "**%s** — %s%s" % (r["작업"], r["말"],
                                     " · %d회 연속" % r["연속"] if r["연속"] > 1 else "")
        elif r["갈래"] == "안돎":
            why = "**%s** — %s" % (r["작업"], r["말"])
        elif r["갈래"] in REPEAT_ONLY and r["연속"] >= REPEAT_LIMIT:
            why = ("**%s** — %d회 연속 %s. 앞 회차가 안 끝나고 있다는 뜻이다"
                   % (r["작업"], r["연속"], r["말"]))
        if why:
            out.append({"갈래": r["갈래"], "작업": r["작업"], "무엇": why,
                        "어떻게": "python schedule_watch.py --print"})
    return out


# ──────────────────────────────────────────────────────── 컴팩팅 배선 점검
def compaction():
    """세션이 가득 찼을 때 **인계를 남기고 요약되는 배선**이 아직 있나.

    빠져도 그 순간에는 아무 일도 안 일어난다 — 세션이 가득 찬 날에야 인계 없이
    끊긴다. 회차가 조용히 죽는 것과 같은 모양이라 여기서 같이 본다.
    """
    path = os.path.join(ROOT, "..", ".claude", "settings.json")
    path = os.path.normpath(path)
    try:
        cfg = json.load(open(path, encoding="utf-8"))
    except Exception as exc:
        return {"확인": False, "왜": "settings.json 을 못 읽었다: %s" % str(exc)[:60],
                "빠진것": []}
    gone = []
    if cfg.get("autoCompactEnabled") is not True:
        gone.append("autoCompactEnabled 가 켜져 있지 않다")
    blob = json.dumps(cfg.get("hooks") or {}, ensure_ascii=False)
    if "session_wrapup" not in blob:
        gone.append("PreCompact 훅이 session_wrapup.py 를 안 부른다 — 요약 직전 인계가 없다")
    if "context_guard" not in blob:
        gone.append("context_guard.py 훅이 없다 — 70/85/95% 예고가 안 뜬다")
    return {"확인": True, "왜": "", "빠진것": gone}


# ─────────────────────────────────────────────────────────────── 회차 본체
def build(now=None):
    """스케줄러를 한 번 묻고 판정을 파일 두 장으로 남긴다."""
    now = now or datetime.now()
    before = {}
    old = _read()
    for r in (old.get("작업") or []):
        before[r.get("작업")] = r

    try:
        tasks = query()
        err = ""
    except Exception as exc:
        tasks, err = [], str(exc)[:160]

    rows = [judge(t, now, before) for t in tasks] if not err else []
    rows.sort(key=lambda r: (r["갈래"] in ("성공", "도는중"), r["작업"]))
    have = {t.get("name") for t in tasks}
    missing = {n: f for n, f in declared().items() if n not in have} if not err else {}
    comp = compaction()

    al = alarms(rows, missing) if not err else []
    for x in (comp["빠진것"] if comp["확인"] else [comp["왜"]]):
        al.append({"갈래": "컴팩팅배선", "작업": "세션 자동 마무리",
                   "무엇": "세션이 가득 찰 때 **인계 없이 끊긴다** — %s" % x,
                   "어떻게": ".claude\\settings.json 확인"})

    st = {
        "시각": now.strftime("%Y-%m-%dT%H:%M:%S"),
        "조회실패": err,
        "작업": rows,
        "등록안됨": missing,
        "컴팩팅": comp,
        "경보": al,
    }
    _write(st)
    _report(st)
    return st


def _read():
    try:
        return json.load(open(STATE, encoding="utf-8"))
    except Exception:
        return {}


def _write(st):
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    tmp = STATE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(st, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE)


def _report(st):
    L = ["# 스케줄러 회차 감시", "",
         "- 확인 시각: %s" % st["시각"],
         "- **읽기 전용이다** — 여기서 회차를 다시 띄우거나 등록하지 않는다.", ""]
    if st["조회실패"]:
        L += ["## ⚠ 확인 못 함", "",
              "- %s" % st["조회실패"],
              "- **이것은 '이상 없음'이 아니다.** 아래 표가 비어 있는 것은 회차가 없어서가",
              "  아니라 못 물어봐서다.", ""]
    else:
        al = st["경보"]
        L += ["## 경보 (%d)" % len(al), ""]
        if not al:
            L.append("걸린 것 없음 — 회차 %d개가 모두 정상이거나 도는 중이다." % len(st["작업"]))
        for a in al:
            L += ["- **[%s]** %s" % (a["갈래"], a["무엇"]), "  - `%s`" % a["어떻게"]]
        L += ["", "## 회차 %d개" % len(st["작업"]), "",
              "| 작업 | 갈래 | 마지막 실행 | 제한시간 | 무슨 일인가 |",
              "|---|---|---|---|---|"]
        for r in st["작업"]:
            # ×N 은 **경보가 되는 갈래에만** 붙인다 — 꺼 둔 작업에 "×2" 가 붙으면
            # 세는 것에 뜻이 있는 것처럼 읽힌다.
            L.append("| %s | %s%s | %s | %s | %s |"
                     % (r["작업"], r["갈래"],
                        " ×%d" % r["연속"] if r["연속"] > 1
                        and r["갈래"] in DEAD + REPEAT_ONLY + ("안돎",) else "",
                        (r["마지막실행"] or "없음").replace("T", " ")[:16],
                        r["제한시간"] or "-", r["말"]))
        if st["등록안됨"]:
            L += ["", "## 설치본은 있는데 등록이 안 된 것", ""]
            for n, f in sorted(st["등록안됨"].items()):
                L.append("- **%s** ← `%s`" % (n, f))
    c = st["컴팩팅"]
    L += ["", "## 컴팩팅 배선", ""]
    if not c["확인"]:
        L.append("- ⚠ %s" % c["왜"])
    elif c["빠진것"]:
        L += ["- ⚠ %s" % x for x in c["빠진것"]]
    else:
        L.append("- 정상 — 자동 요약 + 요약 직전 인계(session_wrapup) + 사용량 예고(context_guard).")
    L.append("")
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    with open(REPORT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))


def banner():
    """`session_handoff` 가 읽는 한 장. **여기서 스케줄러를 다시 묻지 않는다**(`[168]`) —
    인계 문서는 자주 만들어지고 조회는 비싸다. 회차(워치독 30분)가 써 둔 것을 읽는다."""
    st = _read()
    if not st:
        return None
    when = _dt(st.get("시각"))
    age = (datetime.now() - when).total_seconds() / 60.0 if when else None
    return {
        "시각": st.get("시각", ""), "나이분": round(age, 1) if age is not None else None,
        "조회실패": st.get("조회실패", ""),
        "경보": st.get("경보") or [],
        "회차수": len(st.get("작업") or []),
    }


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--print" in argv:
        st = _read()
        if not st:
            print("아직 확인한 적이 없다 — python schedule_watch.py")
            return 0
    else:
        st = build()
    if st.get("조회실패"):
        print("확인 못 함: %s" % st["조회실패"])
        return 1
    al = st.get("경보") or []
    print("회차 %d개 · 경보 %d건 (%s)" % (len(st.get("작업") or []), len(al), st.get("시각", "")))
    for a in al:
        print("  [%s] %s" % (a["갈래"], re.sub(r"\*\*", "", a["무엇"])))
    if not al:
        print("  걸린 것 없음")
    print("상세: reports/스케줄러_회차감시.md")
    return 0


if __name__ == "__main__":
    sys.exit(main())
