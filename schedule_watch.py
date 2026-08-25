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
    75:         ("이어감", "한 회차 몫을 마치고 다음 회차에 이어 간다"),
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
#: ★ 죽었는데 **그 코드가 마지막 실행 뒤에 바뀐** 갈래. 경보가 아니라 알림이다(`[288]`).
#:   실패한 코드는 이미 안 돈다 — 그렇다고 '고쳐졌다'고 말하지도 않는다.
#:   **다음 회차가 답한다.** DEAD 에 절대 넣지 말 것(넣으면 P0 로 되살아난다).
FIXWAIT = "고침대기"
#: 실패했는데 **그 일이 그 뒤에 실제로 끝났다** — 추측이 아니라 회차가 남긴 자국이다.
RANLATER = "뒤에됨"
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
    $act = @()
    foreach ($a in $t.Actions) {
      $act += [pscustomobject]@{
        exe  = [string]$a.Execute
        args = [string]$a.Arguments
      }
    }
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
      act    = $act
    }
  } catch {
    $out += [pscustomobject]@{
      name = [string]$t.TaskName; state = [string]$t.State; reg = ''; limit = '';
      multi = ''; last = ''; next = ''; result = -1; missed = 0;
      err = [string]$_.Exception.Message; trig = @(); act = @()
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
#: 설치본이 실제로 **등록에 넘기는 값**. 변수든 따옴표든 가리지 않는다.
_ARG = re.compile(r"-TaskName\s+(\$[A-Za-z_]\w*|'[^']*'|\"[^\"]*\")")


def _resolve(src, tok):
    """`-TaskName` 에 넘긴 토큰을 실제 이름으로 푼다. 못 풀면 None."""
    if tok[:1] in ("'", '"'):
        return tok[1:-1] or None
    var = re.escape(tok[1:])
    for pat in (r"\$%s\s*=\s*'([^']*)'" % var, r'\$%s\s*=\s*"([^"]*)"' % var):
        m = re.search(pat, src)
        if m:
            return m.group(1) or None
    m = re.search(r"\$%s\s*=\s*-join\s*@\((.*?)\)" % var, src, re.S)
    if m:
        return "".join(chr(int(h, 16)) for h in _CHAR.findall(m.group(1))) or None
    return None


def declared():
    """`install_*.ps1` 이 **등록하는** 작업 이름들 → `({이름: 파일}, 못읽은파일[])`.

    ★ 목록을 손으로 적지 않는다 — 적는 순간 사본이 둘이 되어, 설치본만 늘고 여기는
      안 늘면 새 회차가 등록 안 된 채 조용히 빠진다(정오회차 사고의 모양).
    ★ **이름 관례에 기대지 않는다.** 첫 판은 `$TaskName = "..."` 라는 모양을 찾았는데
      `install_browser_chain_schedule.ps1` 은 `$name = 'CSOS_BrowserChain'` 이었다 —
      변수 이름도 따옴표도 달라서 **그 회차만 목록에서 통째로 빠졌다.** 사라져도 아무
      경보가 안 뜬다는 뜻이고, 그것이 이 파일이 막으려는 바로 그 사고다.
      그래서 이제 **`-TaskName` 에 실제로 넘기는 값**을 읽는다. 관례는 어긋나지만
      등록에 넘기는 인자는 어긋날 수 없다 — 어긋나면 설치 자체가 안 된다.
    ★ **못 읽은 설치본은 조용히 넘기지 않는다**(`[169]`). 등록은 하는데 이름을 못 읽었다면
      그 회차는 감시 밖에 있다 — '이상 없음'이 아니라 '확인 못 함'이다.
    """
    names, unreadable = {}, []
    for fn in sorted(os.listdir(ROOT)):
        if not (fn.startswith("install_") and fn.endswith(".ps1")):
            continue
        try:
            src = open(os.path.join(ROOT, fn), encoding="utf-8", errors="replace").read()
        except OSError:
            continue
        if "Register-ScheduledTask" not in src:
            continue                                 # 등록하지 않는 도우미 스크립트
        got = None
        for tok in _ARG.findall(src):
            got = _resolve(src, tok)
            if got:
                break
        if got:
            names[got] = fn
        else:
            unreadable.append(fn)
    return names, unreadable


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
        # ★ 며칠마다인지를 **날짜에도** 적용한다. 예전엔 days 를 '오늘 시각이 아직
        #   안 왔을 때 한 칸 물러서는' 데만 써서, 예정을 언제나 **오늘**로 쳤다.
        #   그래서 3일마다 도는 회차(UX점검: 08-05 기준 08-08·08-11·08-14)가
        #   08-12·08-13 에도 '예정이 지났는데 안 돌았다'로 나왔다 — 실제로는
        #   08-11 에 rc=0 으로 멀쩡히 돈 회차다. **없는 예정을 지어내면 그 경보는
        #   매일 뜨고, 매일 뜨는 경보는 아무도 안 본다**(`[170]` 의 문).
        gap = (now.date() - start.date()).days
        if gap < 0:
            return None                              # 아직 첫 회차가 오지 않았다
        step = gap // days                           # 시작일부터 지난 주기 수
        base = datetime.combine(start.date() + timedelta(days=step * days), start.time())
        if base > now:                               # 오늘이 회차날인데 시각 전이면 직전 회차
            if step == 0:
                return None
            base = datetime.combine(
                start.date() + timedelta(days=(step - 1) * days), start.time())
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


# ────────────────────── 그 회차가 부르는 코드가 그 뒤에 바뀌었나 (`[288]`)
#: 실행기에서 **코드**로 볼 확장자. `.exe` 는 파이썬·wscript 같은 실행기라 안 본다.
_CODE_EXT = (".py", ".bat", ".ps1", ".vbs", ".cmd")
#: `.bat` 은 껍데기다 — 그 안이 부르는 `.py` 가 진짜 코드다. **한 겹만** 따라간다.
#:   더 깊이 가면 짐작이 되고, 한 겹은 bat 이 스스로 적어 둔 글자라 근거다.
_FOLLOW_DEPTH = 1
_TOKEN = re.compile(r"""["']([^"']+)["']|(\S+)""")


def _under_root(path):
    """이 프로젝트 안의 실제 파일이면 절대경로, 아니면 None."""
    p = path.strip().strip('"').strip("'")
    if not p:
        return None
    full = p if os.path.isabs(p) else os.path.join(ROOT, p)
    full = os.path.normpath(full)
    try:
        if os.path.commonpath([full, ROOT]) != ROOT or not os.path.isfile(full):
            return None
    except ValueError:                               # 드라이브가 다르면 남의 파일이다
        return None
    return full


def _slept_note(since, now):
    """그 예정 뒤로 이 PC 가 **잠들어 있던 분** — 있으면 한 마디, 없으면 빈 글자.

    ★ **'안 돌았다'와 '기계가 자고 있었다'는 다른 사실이다**([385] 가 워치독에서
      배운 것과 같은 자리인데, 회차 쪽은 그것을 모르고 있었다([300]).
      실측 2026-08-22: 정오회차가 **12:05:43** 에 떠서 `종료됨`(0x00041306) 으로
      끝났는데 Kernel-Power **506(절전 진입)이 12:05:40** 이다 — 3초 차이다.
      일일대조도 16:32 에 시작해 `(회차 끝)` 없이 사라졌고 **16:43:38 에 절전**
      진입이다. 그런데 화면은 '예정이 지났는데 안 돌았다'만 말했다 — 그 조치는
      코드를 뒤지는 것이라, 사람이 **멀쩡한 코드를 고치러 간다**([172]).

    ★ **갈래는 안 바꾼다.** 잠을 핑계로 '안 돌았다'를 지우면 안 된다 — 놓친 회차는
      `StartWhenAvailable` 이 꺼져 있으면 영영 안 돈다. 그것은 여전히 사실이고
      사람이 알아야 한다. 여기서 하는 것은 **사실 하나를 덧붙이는 것**까지다.

    ★ **확언하지 않는다**([169]) — 잠든 사이에 예정이 지나갔을 *수* 있다고만 적는다.
      Modern Standby 는 프로세스를 죽이지 않는 것이 보통이고, 깨어난 뒤 늦게
      실행하는지는 그 작업의 설정이 정한다. 우리가 아는 것은 **잤다**는 것뿐이다.
    ★ 못 재면 아무 말도 안 한다 — '안 잤다'로 치지 않는다.
    """
    try:
        분 = (now - since).total_seconds() / 60.0
        if 분 <= 0:
            return ''
        from system_audit import _sleep_minutes_since   # 늦게 — 순환 import 를 안 만든다
        잔분, _why = _sleep_minutes_since(분)
    except Exception:
        return ''
    if not 잔분 or 잔분 < 5:
        return ''
    return (' · 그 사이 이 PC 가 %d분 잤다'
            '(잠든 사이에 예정이 지나갔을 수 있다 — 코드가 아니라 전원 문제다)'
            % int(잔분))


def task_scripts(task, depth=_FOLLOW_DEPTH):
    """그 작업이 **실제로 돌리는** 이 프로젝트 파일들(절대경로). 없으면 빈 목록.

    ★ 실행기(`pythonw.exe`·`wscript.exe`)가 아니라 **인자에 적힌 코드**를 본다.
      `wscript run_hidden.vbs "…\\daily_run.bat"` 처럼 껍데기를 거치는 회차가 있어
      한 겹만 따라간다 — bat 이 부르는 `.py` 는 bat 이 제 손으로 적어 둔 글자다.
    """
    out, seen = [], set()

    def add(p):
        f = _under_root(p)
        if f and f.lower() not in seen:
            seen.add(f.lower())
            out.append(f)

    for a in (task.get("act") or []):
        add(str(a.get("exe") or ""))
        for m in _TOKEN.finditer(str(a.get("args") or "")):
            tok = m.group(1) or m.group(2) or ""
            if tok.lower().endswith(_CODE_EXT):
                add(tok)
    if depth > 0:
        for f in list(out):
            if not f.lower().endswith((".bat", ".cmd", ".vbs")):
                continue
            try:
                text = open(f, encoding="utf-8", errors="replace").read()
            except OSError:
                continue
            # ★ **배치 변수를 먼저 벗긴다** (2026-08-18 실측). `.bat` 은 제 폴더를
            #   `%~dp0` 로 적는데(`"%~dp0daily_run.py"`), 그대로 훑으면 정규식이
            #   `dp0daily_run.py` 를 잡아 **디스크에 없는 이름**이 되고 조용히 버려진다.
            #   그래서 `쿠팡업무_일일자동대조` 의 스크립트 목록이 bat 까지만이었고,
            #   `code_changed_after`(`[110]` 고침대기)가 이 회차에 대해 **눈이 멀어
            #   있었다** — 오류도 안 나고 목록도 그럴듯해서 아무도 몰랐다(`[165]`).
            text = re.sub(r"%~[a-zA-Z]*\d?", " ", text)      # %~dp0 · %~d0
            text = re.sub(r"%[^%\s]{1,40}%", " ", text)      # %LOCALAPPDATA%
            for m in re.finditer(r"[\w./\\-]+\.py\b", text):
                add(m.group(0))
    return [f for f in out if f.lower().endswith(_CODE_EXT)]


#: 회차가 **스스로 남기는 완주 자국**. daily_run 계열만 이 파일을 쓴다 —
#: 목록을 넓힐 때는 "그 회차가 정말 이 파일에 완주를 적는가"를 먼저 확인할 것.
RAN_TRACE = {"daily_run.py": os.path.join(ROOT, "reports", "agent_status.json")}


def ran_after(task, when):
    """실패한 뒤 **그 일이 실제로 끝났다는 자국**이 있나 -> `(시각, 단계수)`.

    ★ 스케줄러는 **작업**의 마지막 결과만 안다. 그런데 이 프로젝트에서 같은 *일*은
      여러 길로 불린다(예약 작업 · `automation_pipeline` · 워치독 · 사람). 실측
      2026-08-18: `쿠팡업무_일일자동대조` 가 09:50 에 exit 1 이었는데 **같은 회차가
      14:54 에 80단계를 완주**했다(`agent_status.json` `aborted:false`). 그런데도
      경보는 **[P0] 로 그대로 남아** 인계 맨 위에 올라왔다 — 경보가 대부분 가짜면
      나머지도 아무도 안 본다(`[170]`).
    ★ `[110]` 의 `고침대기` 보다 **센 근거**다: 저쪽은 '코드가 바뀌었으니 아마
      고쳐졌을 것'이라는 추측이고, 여기는 **그 일이 끝났다는 사실**이다. 그래서
      판정 순서가 이쪽이 먼저다(`[203]` — 좁고 센 근거가 앞에 온다).
    ★ **일반화하지 않는다.** 이 자국은 daily_run 계열에만 있다. 다른 회차에 이
      근거를 대면 '자국이 없다'가 곧 '안 됐다'로 읽혀 거짓이 된다.
    ★ **못 읽으면 그대로 실패로 둔다**(`[169]`) — '이상 없음'으로 바꾸지 않는다.
    """
    if when is None:
        return None, None
    names = {os.path.basename(f).lower() for f in task_scripts(task)}
    for key, path in RAN_TRACE.items():
        if key.lower() not in names:
            continue
        try:
            with open(path, encoding="utf-8") as fh:
                d = json.load(fh)
        except Exception:                            # noqa: BLE001
            continue                                 # 못 읽음 != 안 돌았다
        if d.get("aborted"):
            continue                                 # 중단으로 끝난 자국은 근거가 아니다
        t = _dt(d.get("time"))
        if t and t > when:
            return t, len(d.get("steps") or [])
    return None, None


def code_changed_after(task, when):
    """마지막 실행 **뒤에** 바뀐 코드 파일 → `(상대경로, 바뀐때)` · 없으면 `(None, None)`.

    ★ 커밋 시각이 아니라 **파일 mtime** 을 본다. 도는 것은 커밋이 아니라 디스크 위의
      그 파일이고, `git pull` 로 받아만 놓고 안 돈 경우까지 잡아야 하기 때문이다
      (`[156]` 이 앱 서버에서 쓴 것과 같은 근거).
    ★ 그래서 이 판정이 말할 수 있는 것은 **'그 뒤 코드가 바뀌었다'** 까지다 —
      '고쳐졌다'가 아니다. 무관한 이유로 파일을 건드렸을 수도 있다. 그 구별은
      **다음 회차**가 해 준다(`[169]`).
    """
    if when is None:
        return None, None
    newest, at = None, None
    for f in task_scripts(task):
        try:
            t = datetime.fromtimestamp(os.path.getmtime(f))
        except OSError:
            continue
        if t > when and (at is None or t > at):
            newest, at = os.path.relpath(f, ROOT), t
    return newest, at


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
        say += _slept_note(late, now)

    # ★ 죽었는데 **그 코드가 마지막 실행 뒤에 바뀌었다** — 실패한 코드는 이미 안 돈다.
    #   실측 2026-08-16: `CSOS_BrowserChain` 마지막 실행 08-15 12:00 exit 1 인데 그
    #   exit 1 을 없앤 고침은 08-16 02:04 커밋이었다(14시간 뒤). 그런데 감시자가
    #   고친 시각을 안 봐서 **매일 P0** 로 올라왔다 — 경보가 대부분 가짜면 나머지도
    #   아무도 안 본다(`[170]`). `error_book` 에서 고친 것과 같은 모양이다(`[288]`).
    #   ★ '고쳐졌다'고는 말하지 않는다. 사실은 **'고친 뒤 아직 안 돌았다'** 이고
    #     답은 다음 회차가 한다(`[169]`). '안돎'에는 걸지 않는다 — 그것은 코드가
    #     바뀌었든 말든 **돌아야 할 때 안 돈 것**이라 사실이 그대로 남는다.
    # ★ 죽었는데 **그 일이 그 뒤에 실제로 끝났다** — 자국이 그렇게 말한다.
    #   추측(고침대기)보다 센 근거라 **먼저** 묻는다.
    돈때, 단계수 = ran_after(task, last) if (kind in DEAD and late is None) else (None, None)
    if 돈때:
        say = ("%s — 그런데 **같은 일이 그 뒤 %s 에 완주했다**(%d단계). "
               "예약 회차는 실패했지만 그 일 자체는 됐다"
               % (say, 돈때.strftime("%m-%d %H:%M"), 단계수))
        kind = RANLATER

    바뀐것, 바뀐때 = (None, None)
    if kind in DEAD and late is None:
        바뀐것, 바뀐때 = code_changed_after(task, last)
        if 바뀐것:
            say = ("%s — 그런데 그 뒤 코드가 바뀌었다(%s %s). 실패한 코드는 이미 안 돈다"
                   " · 다음 예정 %s 가 답한다"
                   % (say, 바뀐것, 바뀐때.strftime("%m-%d %H:%M"),
                      (task.get("next") or "?")[:16].replace("T", " ")))
            kind = FIXWAIT

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
        "바뀐코드": 바뀐것 or "", "코드바뀐때": 바뀐때.isoformat(timespec="seconds") if 바뀐때 else "",
        # 실행기까지 남긴다 — 창이 뜨는 실행기로 등록돼 있는지는 여기서만 알 수 있다([107]).
        "실행기": [{"exe": str(a.get("exe") or ""), "args": str(a.get("args") or "")[:200]}
                 for a in (task.get("act") or []) if isinstance(a, dict)],
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


#: '일부러 꺼 뒀다' 고 사람이 적어 두는 곳 — 한 줄에 작업 이름 하나(`#` 는 주석).
OFF_OK = os.path.join(ROOT, "reports", "스케줄러_꺼둔회차.txt")


def _off_on_purpose():
    """사람이 '일부러 껐다'고 적어 둔 작업 이름들 → `(집합, 그 파일을 읽었나)`."""
    try:
        with open(OFF_OK, encoding="utf-8") as fh:
            names = set()
            for line in fh:
                line = line.split("#", 1)[0].strip()
                if line:
                    names.add(line)
        return names, True
    except FileNotFoundError:
        return set(), True                           # 없는 것은 정상이다(아무도 안 껐다)
    except OSError:
        return set(), False                          # 못 읽었다 ≠ 비었다(`[169]`)


def notices(rows, now=None):
    """경보는 아니지만 **알아 둬야 하는 것**(`[288]`).

    ★ 여기 있는 것을 `alarms()` 에 넣으면 안 된다 — `system_audit` 은 경보가 하나라도
      있으면 **무조건 P0** 를 만든다. 그러면 '고친 뒤 아직 안 돎'이 매일 P0 가 되어
      이 고침이 아무 뜻도 없어진다.
    ★ 그렇다고 조용히 빼지도 않는다(`[169]`) — 인계·리포트에 한 줄로 남는다.
    """
    now = now or datetime.now()
    out = []
    for r in rows:
        if r["갈래"] not in (FIXWAIT, RANLATER):
            continue
        out.append({"갈래": r["갈래"], "작업": r["작업"], "무엇": "**%s** — %s"
                    % (r["작업"], r["말"]),
                    "어떻게": "python schedule_watch.py --print"})

    # ★ **꺼진 회차는 실패하지 않는다** — 그래서 여태 아무 화면에도 안 떴다
    #   (2026-08-16 실측: 회차 5개가 08-15 부터 꺼져 있었고 그중 하나가 **09:50
    #   일일자동대조** 였다. 접수취소·객관완료·청구상태·대조·오기입·사실대조가
    #   하루 넘게 통째로 안 돌았는데 스케줄러는 아무 오류도 안 냈다).
    # ★ 그래도 **경보로 올리지 않는다** — 사람이 일부러 껐을 수 있고, 그것을 고장이라
    #   부르면 멀쩡한 결정을 고치러 간다(`[172]`). 되살리지도 않는다(읽기 전용).
    #   말하는 것까지가 이 파일 몫이다.
    꺼둠, 읽음 = _off_on_purpose()
    for r in rows:
        if r["갈래"] != "꺼짐" or r["작업"] in 꺼둠:
            continue
        last = _dt(r.get("마지막실행"))
        얼마 = ("마지막 실행 %s · %d일 전" % (r["마지막실행"][:16].replace("T", " "),
                                        (now - last).days)) if last else "실행 기록 없음"
        out.append({"갈래": "꺼짐", "작업": r["작업"],
                    "무엇": "**%s** — 꺼져 있다(%s). 사람이 껐다면 정상이다 — "
                            "일부러라면 `reports/스케줄러_꺼둔회차.txt` 에 이름을 적어 "
                            "이 줄을 내린다" % (r["작업"], 얼마),
                    "어떻게": "powershell Enable-ScheduledTask -TaskName '%s'" % r["작업"]})
    # ★ 살아 있는 작업이 **창 뜨는 실행기**로 등록돼 있나 (2026-08-16, 분담판 [107]).
    #   `[248]`·`[272]` 는 **설치본과 소스**를 본다 — 그런데 설치본을 고쳐도 이미 등록된
    #   작업은 그대로고, 살아 있는 작업만 고치면 기계를 새로 만들 때 되살아난다.
    #   여기는 **지금 등록돼 있는 것**을 본다. 고치지는 않는다(읽기 전용).
    try:
        from tools.window_audit import exe_verdict
    except Exception as exc:
        out.append({"갈래": "확인못함", "작업": "window_audit",
                    "무엇": "예약 작업이 창 뜨는 실행기로 등록됐는지 **못 봤다**(%s)"
                            % type(exc).__name__,
                    "어떻게": "python tools\\window_audit.py --live"})
    else:
        아는것 = [r for r in rows if r.get("실행기")]
        if rows and not 아는것:
            out.append({"갈래": "확인못함", "작업": "예약 작업 실행기",
                        "무엇": "실행기를 한 줄도 못 읽었다 — 창이 뜨는지 **0곳이라 "
                                "말하면 안 되는 자리**다(`[169]`)",
                        "어떻게": "python schedule_watch.py --print"})
        for r in 아는것:
            for a in r["실행기"]:
                판정 = exe_verdict(a.get("exe"), a.get("args"))
                if 판정 == "조용":
                    continue
                out.append({"갈래": "창뜸" if 판정 == "창뜸" else "확인못함", "작업": r["작업"],
                            "무엇": "**%s** — 실행기가 `%s` 다(%s). 이 회차가 돌 때마다 "
                                    "검은 창이 뜬다 — `pythonw.exe` 또는 "
                                    "`wscript.exe run_hidden.vbs` 로 등록한다"
                                    % (r["작업"], os.path.basename(a.get("exe") or "?"), 판정),
                            "어떻게": "python tools\\window_audit.py --live"})
    # ★ **로그인 자동실행도 본다** (2026-08-21 형님 지시). 위 `live()` 는 예약 작업만
    #   보는데, `[263]` 대로 예약 작업 등록이 막힌 기계에서는 설치기가 HKCU 로그인
    #   자동실행으로 **스스로 전환**한다. 그 자리는 지금껏 아무도 안 봤다 — 창 뜨는
    #   실행기가 들어가면 부팅할 때마다 창이 뜨는데 어느 화면에도 안 뜬다(`[169]`).
    #   읽기 전용이다 — 고치는 것은 사람이 정한다.
    try:
        from tools.window_audit import autorun as _autorun
    except Exception as exc:
        out.append({"갈래": "확인못함", "작업": "로그인 자동실행",
                    "무엇": "로그인 자동실행이 창 뜨는 실행기인지 **못 봤다**(%s)"
                            % type(exc).__name__,
                    "어떻게": "python tools" + chr(92) + "window_audit.py --autorun"})
    else:
        _rows, _why, _seen = _autorun()
        if _why:
            out.append({"갈래": "확인못함", "작업": "로그인 자동실행",
                        "무엇": "로그인 자동실행을 **확인 못 했다** — %s" % _why,
                        "어떻게": "python tools" + chr(92) + "window_audit.py --autorun"})
        for _name, _exe, _v in _rows:
            out.append({"갈래": "창뜸" if _v == "창뜸" else "확인못함",
                        "작업": "로그인 자동실행: %s" % _name,
                        "무엇": "**%s** — 로그인 자동실행이 `%s` 로 걸려 있다(%s). "
                                "부팅할 때마다 검은 창이 뜬다 — `pythonw.exe` 로 바꾼다"
                                % (_name, os.path.basename(_exe or "?"), _v),
                        "어떻게": "python tools" + chr(92) + "window_audit.py --autorun"})
    # ★ **세션 훅도 창을 띄운다 — 그리고 그 자리는 아무도 안 봤다.**
    #   `[272]` 는 소스를, `live()` 는 예약 작업을, `autorun()` 은 로그인 항목을 본다.
    #   그런데 훅은 `.claude/settings.json` 에 있어 셋 중 어디에도 안 잡혔다.
    #   실측 2026-08-21: 훅 여섯이 전부 `python`(콘솔)이었고 그중 `PostToolUse` 는
    #   **도구를 부를 때마다** 돈다 — 한 응답 동안 창이 수십 번 번쩍이는데 두 감사기는
    #   나란히 `0곳` 이라 말했다(`[169]`). 읽기 전용이다 — 고치는 것은 사람이 정한다.
    try:
        from tools.window_audit import hooks as _hooks
    except Exception as exc:
        out.append({"갈래": "확인못함",
                    "작업": "세션 훅",
                    "무엇": "세션 훅이 창 뜨는 실행기인지 **못 봤다**(%s)" % type(exc).__name__,
                    "어떻게": "python tools" + chr(92) + "window_audit.py --hooks"})
    else:
        _hrows, _hwhy, _hseen = _hooks()
        if _hwhy:
            out.append({"갈래": "확인못함",
                        "작업": "세션 훅",
                        "무엇": "세션 훅을 **확인 못 했다** — %s" % _hwhy,
                        "어떻게": "python tools" + chr(92) + "window_audit.py --hooks"})
        for _어디, _사건, _exe, _v in _hrows:
            out.append({"갈래": ("창뜸" if _v == "창뜸" else "확인못함"),
                        "작업": "세션 훅: %s %s" % (_어디, _사건),
                        "무엇": "**%s** 훅이 `%s` 로 걸려 있다(%s) — 그 훅이 돌 때마다 검은 창이 뜬다. `pythonw` 로 바꾼다(실측: pythonw 도 파이프 stdout 을 그대로 살려 훅이 내보내는 한 줄을 안 잃는다)." % (_사건, _exe, _v),
                        "어떻게": "python tools" + chr(92) + "window_audit.py --hooks"})
    # ★ **다섯째 구멍 — 시작 폴더.** (2026-08-21 형님 지시가 **두 번째로** 온 자리)
    #   네 축(소스·예약작업·로그인항목·훅)이 전부 초록인데도 같은 지시가 다시 왔다.
    #   실측: 시작 폴더에 스크립트 자동실행 셋이 있는데 **어느 축도 그 파일을 한 글자도
    #   안 봤다.** 지금은 셋 다 창 숨김(`Run …, 0`)이지만 그것을 **재는 계기가 없다는
    #   것**이 문제다 — 하나가 `1` 로 바뀌면 부팅할 때마다 창이 뜨는데 조용하다(`[169]`).
    #   그리고 `wscript` 를 실행기 이름만 보고 '조용'이라 하면 안 된다. 창이 뜨는지
    #   정하는 것은 실행기가 아니라 **vbs 안의 Run 두 번째 인자**다.
    try:
        from tools.window_audit import startup as _startup
    except Exception as exc:
        out.append({"갈래": "확인못함", "작업": "시작 폴더",
                    "무엇": "시작 폴더 자동실행이 창 뜨는지 **못 봤다**(%s)" % type(exc).__name__,
                    "어떻게": "python tools" + chr(92) + "window_audit.py --startup"})
    else:
        _srows, _swhy, _sseen = _startup()
        if _swhy:
            out.append({"갈래": "확인못함", "작업": "시작 폴더",
                        "무엇": "시작 폴더를 **확인 못 했다** — %s" % _swhy,
                        "어떻게": "python tools" + chr(92) + "window_audit.py --startup"})
        for _fn, _v, _why2 in _srows:
            out.append({"갈래": ("창뜸" if _v == "창뜸" else "확인못함"),
                        "작업": "시작 폴더: %s" % _fn,
                        "무엇": "**%s** — 로그인할 때 도는 자동실행이 창을 단다(%s · %s). "
                                "부팅할 때마다 검은 창이 떠 형님 화면을 가린다."
                                % (_fn, _v, _why2),
                        "어떻게": "python tools" + chr(92) + "window_audit.py --startup"})
    # ★ **여섯째 — 이 저장소 밖.** 형님 지시는 "**어떤 계정 어떤 세션에서 진행해도**" 다.
    #   그런데 이 감사기는 제 저장소 하나만 훑는다. 실측 2026-08-21: 이 PC 에서 자동으로
    #   도는 다른 프로젝트가 셋이고 그중 하나(UNI Cash Flow)에 **콘솔 exe 를 깃발 없이
    #   띄우는 자리 18곳**이 있었다 — 그 앱은 시작 폴더로 부팅마다 도는데 이 저장소의
    #   어느 계기도 그것을 못 셌다.
    #   ★ **경보가 아니라 알림이다**(`[170]`·`[172]`). 남의 저장소는 규칙이 다를 수 있고
    #     여기서 '위반'이라 부르면 거짓 경보가 된다. 고치는 것은 그 세션이 정한다 —
    #     여기는 **숫자로 말하는 것까지**다(`[169]` — 조용히 빼면 없는 것으로 읽힌다).
    # ★ 2026-08-25 형님 지시: **"csos 앱 관련 사항이 아니면 끊어"** — 그래서 이 축은
    #   **기본이 꺼짐**이다. 남의 저장소 소식은 이 앱의 일이 아니고, 매일 뜨면
    #   진짜 경보가 묻힌다([170]).
    #   ★ **재는 것과 말하는 것을 가른다**([169]) — 도구는 그대로 산다:
    #     `python tools/window_audit.py --neighbors` · `--root "<폴더>"` 는 여전히
    #     답한다. 끊은 것은 **이 회차가 인계에 올리는 길** 하나다.
    #   ★ 되돌리기는 한 줄이다([126] 의 보호장치):
    #     `COUPANG_NEIGHBOR_WINDOW_AUDIT=1`
    #   ⚠ **이 저장소 안의 여섯 축은 한 글자도 안 바뀐다**([172]) — 소스·예약작업·
    #     로그인항목·훅·시작폴더가 그대로 돈다. 끊은 것은 **이웃 저장소 하나**다.
    if os.environ.get("COUPANG_NEIGHBOR_WINDOW_AUDIT") == "1":
        try:
            from tools.window_audit import neighbors as _nb, split as _split
        except Exception:
            pass
        else:
            try:
                _dirs, _ = _nb()
            except Exception:
                _dirs = []
            for _d in _dirs:
                try:
                    _sure, _unk = _split(root=_d)
                except Exception as exc:
                    out.append({"갈래": "확인못함", "작업": "이웃 저장소: %s" % os.path.basename(_d),
                                  "무엇": "창 자리를 **못 셌다**(%s)" % type(exc).__name__,
                                  "어떻게": 'python tools' + chr(92) + 'window_audit.py --root "%s"' % _d})
                    continue
                if _sure or _unk:
                    out.append({"갈래": "이웃창", "작업": "이웃 저장소: %s" % os.path.basename(_d),
                                  "무엇": "이 PC 에서 자동으로 도는 다른 프로젝트 — 콘솔 exe 를 깃발 없이 "
                                          "띄우는 자리 **%d곳** · 무엇을 띄우는지 못 읽은 자리 **%d곳**. "
                                          "둘을 뭉치지 않는다(뒤는 지목이 아니라 못 봤다는 보고다 · [169]). "
                                          "그 세션이 고칠 자리다 — 여기서는 안 고친다." % (len(_sure), len(_unk)),
                                  "어떻게": 'python tools' + chr(92) + 'window_audit.py --root "%s"' % _d})
    # ★ **같은 회차를 두 작업이 부르면 한쪽은 늘 잠금에 막혀 실패한다.** 그런데 그 실패는
    #   '회차가 고장 났다'로 읽혀서 원인을 엉뚱한 데서 찾게 된다 — 2026-08-07 에
    #   `쿠팡업무_원장일괄반영_15시` 를 지운 것이 바로 이 모양이었고(매일 `결과: 1`),
    #   그때는 **사람이 손으로 찾아냈다.** 실측 2026-08-16: `쿠팡업무_일일자동대조` 와
    #   `CSOS_유수비_대표보고_자동준비` 가 같은 `daily_run.bat` 을 부른다.
    #   지우지는 않는다 — 어느 쪽을 남길지는 사람의 판단이다(읽기 전용).
    같은것 = {}
    for r in rows:
        for a in (r.get("실행기") or []):
            열쇠 = (str(a.get("exe") or "").strip().lower(),
                   str(a.get("args") or "").strip().lower())
            if not any(열쇠):
                continue
            같은것.setdefault(열쇠, set()).add(r["작업"])
    for (exe, _args), 이름들 in sorted(같은것.items(), key=lambda kv: sorted(kv[1])):
        if len(이름들) < 2:
            continue
        out.append({"갈래": "겹침", "작업": " · ".join(sorted(이름들)),
                    "무엇": "**%s** — 같은 것을 부른다(`%s`). 한쪽은 잠금에 막혀 늘 "
                            "실패하고, 그 실패가 '회차 고장'으로 읽힌다"
                            % (" · ".join(sorted(이름들)), os.path.basename(exe)),
                    "어떻게": "powershell Get-ScheduledTask -TaskName '%s' | "
                            "Select-Object -ExpandProperty Actions" % sorted(이름들)[0]})
    if not 읽음:
        out.append({"갈래": "확인못함", "작업": "스케줄러_꺼둔회차.txt",
                    "무엇": "'일부러 꺼 둔 회차' 목록을 **못 읽었다** — 꺼진 회차를 "
                            "일부러인지 아닌지 가리지 못한다",
                    "어떻게": "type reports\\스케줄러_꺼둔회차.txt"})

    # ★ **회차가 남긴 자국도 인계까지 간다** (2026-08-23 실측).
    #   여태 `traces()` 는 리포트 마크다운에만 실렸다. 그래서 스케줄러 결과가
    #   **성공(exit 0)** 인 회차가 남긴 자국은 **어느 화면에도 안 떴다** — 실측으로
    #   `브라우저수집_오류.json`(12일째 기회를 못 잡음)이 인계 18건 어디에도 없었다.
    #   exit 코드로 뜨는 것과 자국으로만 아는 것은 다른 사실이다([169]).
    # ★ 경보가 아니라 **알림**이다 — 경보에 넣으면 `system_audit` 이 무조건 P0 로
    #   올려 진짜 P0 를 덮는다([288]·[170]).
    # ★ 두 목소리로 울지 않는다([325]) — 그 작업이 이미 위에서 실렸으면 뺀다.
    이미 = set()
    for x in out:
        이미.add(str(x.get("작업") or ""))
    for r in rows:
        if r["갈래"] in DEAD + ("안돎",):
            이미.add(r["작업"])
    for t in traces():
        일 = str(t.get("작업") or "")
        if 일 and 일 in 이미:
            continue                       # 그 회차는 이미 제 이름으로 말했다
        out.append({"갈래": "자국", "작업": 일 or t["파일"],
                    "무엇": "`%s` %s — %s" % (t["파일"], t["시각"], t["무엇"]),
                    "어떻게": t.get("어떻게") or "type reports\\" + t["파일"]})
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
    decl, unreadable = declared()
    missing = {n: f for n, f in decl.items() if n not in have} if not err else {}
    comp = compaction()

    al = alarms(rows, missing) if not err else []
    for fn in unreadable:
        al.append({"갈래": "확인못함", "작업": fn,
                   "무엇": "설치본 `%s` 이 무슨 작업을 등록하는지 못 읽었다 — "
                           "그 회차는 감시 밖이라 사라져도 아무 경보가 안 뜬다" % fn,
                   "어떻게": "설치본에서 -TaskName 에 넘기는 값을 확인한다"})
    for x in (comp["빠진것"] if comp["확인"] else [comp["왜"]]):
        al.append({"갈래": "컴팩팅배선", "작업": "세션 자동 마무리",
                   "무엇": "세션이 가득 찰 때 **인계 없이 끊긴다** — %s" % x,
                   "어떻게": ".claude\\settings.json 확인"})

    st = {
        "시각": now.strftime("%Y-%m-%dT%H:%M:%S"),
        "조회실패": err,
        "작업": rows,
        "등록안됨": missing,
        "설치본못읽음": unreadable,
        "컴팩팅": comp,
        "경보": al,
        # 경보와 **같은 통에 담지 않는다** — system_audit 은 경보가 하나라도 있으면
        # 무조건 P0 를 만든다(`[288]`). 그러나 조용히 빼지도 않는다(`[169]`).
        "알림": notices(rows) if not err else [],
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
            # ⚠ "모두 정상"이라고 적지 않는다 — 꺼진 회차는 **경보가 아니라 알림**이라
            #   여기 안 들어온다. 경보 0 을 '이상 없음'으로 읽히게 두면 5개가 꺼져
            #   있어도 화면은 안심시킨다(2026-08-16 실측 · `[169]`).
            L.append("경보로 올릴 것 없음 — 회차 %d개%s." % (
                len(st["작업"]),
                "" if not st.get("알림") else
                " (다만 **알림 %d건**은 아래에 있다)" % len(st["알림"])))
        for a in al:
            L += ["- **[%s]** %s" % (a["갈래"], a["무엇"]), "  - `%s`" % a["어떻게"]]
        노 = st.get("알림") or []
        if 노:
            L += ["", "## 알림 — 경보는 아니지만 알아 둘 것 (%d)" % len(노), "",
                  "실패한 뒤 **그 코드가 바뀐** 회차다. 실패한 코드는 이미 안 돌지만",
                  "'고쳐졌다'고 단정하지 않는다 — **다음 회차가 답한다**(`[288]`).", ""]
            for a in 노:
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
    tr = traces()
    if tr:
        # ★ 스케줄러는 '어느 회차가 죽었나'까지만 안다. **왜**는 회차가 스스로 남겨야
        #   한다 — pythonw 로 도는 회차는 트레이스백이 어디에도 안 남기 때문이다.
        L += ["", "## 회차가 남긴 자국 (%d)" % len(tr), ""]
        for t in tr:
            L += ["- `%s` %s — %s" % (t["파일"], t["시각"], t["무엇"])]
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


def traces():
    """회차가 죽으며 남긴 자국들 — `reports/*_오류.json` (2026-08-12).

    ★ 스케줄러는 **'어느 회차가 죽었나'까지만** 안다. exit 1 은 그저 1 이다.
      **왜**인지는 회차가 스스로 남겨야 한다 — 이 프로젝트의 회차들은 `pythonw.exe`
      로 돌아 **창이 없고, 그래서 트레이스백이 어디에도 안 남는다.**
      실측: `쿠팡업무_밴드재수집` 이 매일 exit 1 이었는데 그 사실조차 이 파일을 만들고
      나서야 보였고, **왜인지는 그때도 알 길이 없었다.**
    ★ 목록을 손으로 적지 않는다 — 자국을 남기는 회차가 늘면 저절로 여기 나온다.
    """
    import glob as _glob
    out = []
    for path in sorted(_glob.glob(os.path.join(ROOT, "reports", "*_오류.json"))):
        try:
            d = json.load(open(path, encoding="utf-8"))
        except Exception:
            continue
        out.append({"파일": os.path.basename(path), "시각": d.get("시각", ""),
                    "무엇": str(d.get("무엇") or "")[:120],
                    # ★ 자국이 제 작업 이름을 적어 두면 위(`notices`)에서 두 목소리를
                    #   가를 수 있다. 안 적어도 된다 — 그때는 그대로 올린다([169]).
                    "작업": str(d.get("작업") or ""),
                    "어떻게": str(d.get("어떻게") or "")[:160]})
    return out


def due_state(script, done_at=None, now=None):
    """그 코드를 돌리는 예약 회차가 **아직 예정이 안 왔나** (2026-08-21 실사고).

    ★ 왜 필요한가 — `session_handoff.daily_run_health` 는 완주 기록의 **나이만** 보고
      20시간이 지나면 밀림이라 말한다. 그런데 이 회차는 하루 한 번 09:50 이라,
      어제 12:20 에 완주했으면 **오늘 08:20 부터 09:50 까지** 그 조건이 참이 된다 —
      회차가 늦은 것이 아니라 **아직 예정이 안 온 것**이다. 실측 2026-08-21 08:47:
      스케줄러는 `마지막실행 08-20 09:50 · 코드 0 · 다음예정 08-21 09:50` 이라 말하는데
      인계는 `[P1] 20시간째 완주하지 않았다` 를 올리고 조치로 `python daily_run.py` 를
      시켰다. 그대로 하면 **63분 뒤 예정된 회차를 잠금으로 막는** 150분짜리 Z: 회차를
      지금 띄운다 — 그러면 예정 회차는 조용히 건너뛰고 스케줄러는 '성공'이라 적는다.
      한 문서 안에서 두 줄이 어긋나면 사람은 없는 것을 찾아 나선다(`[172]`·`[325]`).
    ★ 여기서 스케줄러를 다시 묻지 않는다(`[168]`) — 회차(워치독 30분)가 써 둔 것을 읽는다.
      판정도 새로 만들지 않는다(`[162]`) — 어느 작업이 그 코드를 돌리는지는
      `task_scripts` 가 이미 안다. 작업 이름을 손으로 적으면 이름이 바뀐 날
      **한 건도 안 걸리면서 오류도 안 난다**(`[165]`·`[228]`).
    ★ **조용한 건너뜀을 놓치지 않는다.** 스케줄러의 마지막 실행이 완주 자국보다
      **나중**이면 그 회차는 돌았는데 완주를 못 한 것이다 — 그것이 바로 이 경보가
      원래 잡으려던 사고이므로(2026-08-07) 그때는 밀림 그대로 둔다.

    돌려주는 것 (`[169]` — 못 갈랐으면 '정상'이라 하지 않는다):
      `None`                    = 못 갈랐다(감시 파일 없음·그 코드를 도는 작업 없음·시각 못 읽음)
      `{"아직": True,  ...}`    = 마지막 예정 회차는 완주 자국이 설명하고 다음 예정은 미래다
      `{"아직": False, "왜": …}` = 예정이 지났는데 안 돌았거나, 돌았는데 완주 자국이 없다
    """
    st = _read()
    rows = st.get("작업") or []
    if not rows:
        return None
    now = now or datetime.now()
    want = str(script).lower()
    mine = []
    for r in rows:
        try:
            # 저장된 행은 실행기를 `실행기` 에 담는다 — `task_scripts` 는 `act` 로 받는다.
            files = task_scripts({"act": r.get("실행기") or []})
        except Exception:
            continue
        if any(os.path.basename(f).lower() == want for f in files):
            mine.append(r)
    if not mine:
        return None
    last, future = None, None
    for r in mine:
        t = _dt(r.get("마지막실행"))
        if t and (last is None or t > last):
            last = t
        # ★ 갈래가 '성공' 인 행의 다음 예정만 센다. 꺼져 있거나 죽은 회차의 예정은
        #   **오지 않는다** — 그것을 '곧 돈다'로 읽으면 진짜 밀림이 조용해진다.
        if r.get("갈래") == "성공":
            n = _dt(r.get("다음예정"))
            if n and n > now and (future is None or n < future):
                future = n
    if last is None:
        return None
    if done_at is not None and last > done_at:
        return {"아직": False, "마지막실행": last, "다음예정": future,
                "왜": "스케줄러는 %s 에 돌았다는데 완주 자국은 그보다 앞이다"
                      % last.strftime("%m-%d %H:%M")}
    if future is None:
        return {"아직": False, "마지막실행": last, "다음예정": None,
                "왜": "앞으로 예정된 정상 회차가 없다"}
    return {"아직": True, "마지막실행": last, "다음예정": future,
            "왜": "다음 예정 %s 가 아직 안 왔다" % future.strftime("%m-%d %H:%M")}


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
        # 경보와 갈라 싣는다 — 인계는 이것을 '먼저 처리할 것' 이 아니라 알림으로 적는다.
        "알림": st.get("알림") or [],
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
        print("  경보로 올릴 것 없음")
    # 알림도 콘솔에 보인다 — 파일에만 있으면 명령을 친 사람은 '이상 없음'으로 읽는다.
    for a in (st.get("알림") or [])[:6]:
        print("  (알림) [%s] %s" % (a["갈래"], re.sub(r"\*\*", "", a["무엇"])[:110]))
    print("상세: reports/스케줄러_회차감시.md")
    return 0


if __name__ == "__main__":
    sys.exit(main())
