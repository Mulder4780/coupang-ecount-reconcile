# -*- coding: utf-8 -*-
"""앱·Claude Code·Codex가 함께 쓰는 시스템·업무 진단 정본.

이 파일은 이미 각 자동화가 남긴 작은 상태 파일만 읽는다. 원장·엑셀·외부 사이트를
열지 않으므로 앱 서버가 주기적으로 실행해도 담당자 업무를 붙잡지 않는다. 결과는
``reports/시스템_업무진단.json``과 ``.md``에 원자적으로 교체한다.

사용법::

    python system_audit.py             # 보고서 갱신
    python system_audit.py --print     # 같은 판정을 콘솔에도 표시
    python system_audit.py --json      # 기계가 읽는 JSON 출력
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
REPORTS = ROOT / "reports"
OUT_JSON = REPORTS / "시스템_업무진단.json"
OUT_MD = REPORTS / "시스템_업무진단.md"
VERSION = 1


def _now() -> datetime:
    return datetime.now().astimezone()


def _sleep_minutes_since(minutes_ago: float) -> tuple[float | None, str]:
    """최근 `minutes_ago` 분 동안 이 PC 가 **Modern Standby 로 자고 있던 분**.

    ★ **'멈춤'과 '기계가 자고 있었다'는 다른 사실이다** (2026-08-22 실사고).
      워치독 판정이 **로그 나이 하나**뿐이라, 형님이 노트북을 덮어 둔 밤에도
      `[P0] 워치독 30분 회차가 멈춤` 을 확언했다. 실측: 08-21 22:04 ~ 08-22 01:19
      **3시간 15분** 잠(Kernel-Power 506 진입 · 507 이탈) — 공백이 정확히 그 안이다.
      그런데 조치는 *"확인창을 띄우지 말고…"* 라, 그대로 하면 사람이 **멀쩡한
      워치독 코드**를 고치러 간다([172] — 틀린 지목이 못 잡는 것보다 나쁘다).

    ★ **지어낼 것이 없다** — 이벤트 로그에 그대로 적혀 있다. 실측 조회 **0.41초**라
      비싼 탐색도 아니다([168]). 창은 안 띄운다(`proc_guard` · [272]).

    ★ **못 읽으면 `None` 이다 — 0 이 아니다**([169]). 0 을 주면 '안 잤다'가 되어
      **예전과 똑같이 P0** 를 확언하는데, 그것은 '확인했다'는 거짓이다.
      부르는 쪽이 그 구별을 그대로 사람에게 말한다.

    ⚠ **506/507 은 Modern Standby 뿐이다.** 옛 절전(S3)·최대절전·종료는 다른
      이벤트라 여기서 안 센다 — 모르는 갈래를 아는 것처럼 세지 않는다([169]).
    """
    if os.name != "nt":
        return None, "윈도우가 아니라 이벤트 로그를 못 본다"
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import proc_guard  # noqa: PLC0415  (늦게 들여온다 — 순환 방지)
    except Exception as exc:  # pragma: no cover
        return None, "%s: %s" % (type(exc).__name__, exc)
    hours = max(1, int(minutes_ago / 60) + 2)
    ps = ("$ErrorActionPreference='SilentlyContinue';"
          "Get-WinEvent -FilterHashtable @{LogName='System';"
          "ProviderName='Microsoft-Windows-Kernel-Power';Id=506,507;"
          "StartTime=(Get-Date).AddHours(-%d)} |"
          " ForEach-Object { '{0}|{1}' -f $_.Id, $_.TimeCreated.ToString('s') }" % hours)
    try:
        res = proc_guard.run_tree(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], timeout=40)
    except Exception as exc:
        return None, "%s: %s" % (type(exc).__name__, exc)
    if res.timed_out:
        return None, "이벤트 조회가 40초를 넘겼다"
    if res.returncode != 0:
        return None, "이벤트 조회 실패(rc=%s)" % res.returncode
    now = datetime.now()
    since = now - timedelta(minutes=minutes_ago)
    events = []
    for line in (res.stdout or "").splitlines():
        part = line.strip().split("|")
        if len(part) != 2:
            continue
        try:
            events.append((int(part[0]), datetime.fromisoformat(part[1])))
        except ValueError:
            continue
    # ★ 같은 초에 506·507 이 둘 다 오면 **시각만으로는 순서를 못 정한다**.
    #   실측 2026-08-22: `12:05:40` 에 507(이탈)·506(진입)이 같이 있고 그 잠의
    #   이탈은 `15:26:18` 이다. 시각만으로 정렬하면 파이썬 **안정 정렬**이 입력
    #   순서(내림차순 조회 결과)를 그대로 두어 `506 → 507` 로 짝지어지고
    #   **그 3시간 20분이 0분**이 된다. 그날 오후 실제 227분을 이 함수는 **0.0**
    #   이라 답했다 — 그러면 부르는 쪽은 '안 잤다'로 읽고 **예전과 똑같이 P0**
    #   를 확언한다([169] — 0 을 내는 계기는 아무도 의심하지 않는다).
    #   위 독스트링이 '못 읽으면 None 이지 0 이 아니다'를 경고해 뒀는데,
    #   **틀리게 세어 0 이 되는 갈래**는 막지 못했다.
    #   같은 초면 **이탈(507)이 먼저**다 — 깼다가 곧바로 다시 잔 것이다.
    events.sort(key=lambda x: (x[1], 0 if x[0] == 507 else 1))
    total = 0.0
    enter = None
    for ident, when in events:
        if ident == 506:
            enter = when
        elif ident == 507 and enter is not None:
            total += _overlap_minutes(enter, when, since, now)
            enter = None
    if enter is not None:          # 아직 안 깬 것으로 적혀 있으면 지금까지로 본다
        total += _overlap_minutes(enter, now, since, now)
    return total, ""


def _overlap_minutes(a0, a1, b0, b1) -> float:
    """두 구간이 겹치는 분. 잠이 **마지막 로그 이전**에도 있으므로 겹침만 센다."""
    lo = max(a0, b0)
    hi = min(a1, b1)
    if hi <= lo:
        return 0.0
    return (hi - lo).total_seconds() / 60.0


def _age_minutes(path: Path) -> float | None:
    try:
        return max(0.0, (_now().timestamp() - path.stat().st_mtime) / 60.0)
    except OSError:
        return None


def _repair_text(value: str) -> str:
    """옛 보고서에 섞인 UTF-8/CP949 이중 디코딩 흔적을 읽을 때만 복구한다."""
    choices = [value]
    for encoding in ("utf-8", "cp949"):
        try:
            choices.append(value.encode("latin1").decode(encoding))
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass

    def score(text: str) -> tuple[int, int]:
        korean = sum("가" <= ch <= "힣" for ch in text)
        noise = sum(ch in "ÃÂÀÁÈÉíìëê±°¶§" for ch in text)
        return korean - noise * 2, -len(text)

    return max(choices, key=score)


def _normalize(value: Any) -> Any:
    if isinstance(value, dict):
        return {_repair_text(str(k)): _normalize(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_normalize(v) for v in value]
    if isinstance(value, str):
        return _repair_text(value)
    return value


def _daily_step_reasons(names):
    """실패한 단계의 **사유**를 종합리포트에서 읽는다 — 못 읽으면 빈 표다(`[169]`).

    ★ 왜 필요한가 (2026-08-26 실사고): 진행 자국에는 실패한 **단계 이름**만 있고
      왜인지는 종합리포트의 `## <단계>` 블록에 있다. 이름만 실으면 사람은 파일을
      열어야 하고, 열지 않으면 **조치가 답을 안 준 것**이다(`[289]`).
    ★ **뽑는 법을 새로 만들지 않는다**(`[162]`) — `autopilot._why_line` 을 빌린다.
      그것은 트레이스백 꼬리가 아니라 **오류 줄을 앞세운다**(`[365]`).
    ★ 진행 자국과 종합리포트는 **같은 회차가 같은 순간에** 쓴다(`finish()` 와
      `note_progress("(회차 끝)")`) — 그래서 **가장 새 리포트**가 곧 이 자국의 짝이다.
    ★ 못 읽으면 **지어내지 않는다**(`[169]`) — 빈 표를 주면 부르는 쪽이 이름만 싣는다.
    """
    if not names:
        return {}
    try:
        picks = sorted(REPORTS.glob("종합리포트_*.md"))
        if not picks:
            return {}
        text = picks[-1].read_text(encoding="utf-8", errors="replace")
        from autopilot import _why_line
    except Exception:
        return {}
    out = {}
    for name in names:
        head = chr(10) + "## " + str(name) + chr(10)
        i = text.find(head)
        if i < 0:
            continue
        j = text.find(chr(10) + "## ", i + len(head))
        body = text[i + len(head): j if j > 0 else len(text)]
        why = _why_line(body.replace("```", " "))
        if why:
            out[str(name)] = why
    return out


def _steps_resource_recovered(reasons, names):
    """실패한 단계가 **전부 자원 탓**이고 그 자원이 **지금 살아 있나** — 모르면 None.

    ★ 왜 필요한가 (2026-08-27 실측): 어제 11:46 회차가 **끝까지 돌고** 단계 13개만
      실패했는데, 그 사유가 **13개 전부 같은 것**이었다 — `관리대장을 찾을 수 없음:
      Z:/…`. 곧 코드가 깨진 것이 아니라 **그 순간 공유폴더를 못 잡은 것**이고
      (핫스팟 회선 · `[443]`) 지금 재면 Z: 는 0.3초로 멀쩡하다. 그런데 판정이
      그것을 묻지 않아 **매일 아침 P0 가 인계 맨 위**를 차지했다 — 그 조치는
      사람을 멀쩡한 코드로 보내고(`[172]`) 진짜 경보를 덮는다(`[170]`).
      `[424]` 가 자율복구에서 배운 그 자리인데 **여기에는 안 와 있었다**(`[300]`).

    ★ **판정을 새로 만들지 않는다**(`[162]`) — 갈래는 `autopilot.classify_failure`,
      살아 있나는 `autopilot.resource_back` 을 그대로 빌린다.
    ★ **안전핀은 갈래다** — 하나라도 `resource` 가 아니면 **모름**이다.
      코드 고장이 섞였을 수 있고, 그것을 P2 로 내리면 못 잡는 것보다 나쁘다.
    ★ **사유를 못 읽은 단계가 있으면 모름**이다(`[169]`) — 안 본 것을 회복이라
      부르지 않는다.
    ★ **살아난 것이 하나도 없으면 모름**이다 — `resource_back` 은 문구에 오류
      표시가 없으면 `None` 을 준다(실측 넷 중 둘). 그 `None` 만으로는
      "자원이 돌아왔다" 고 말할 근거가 없다.

    돌려주는 값: True 회복 · False 아직 죽어 있음 · None 못 갈랐다.
    """
    if not names:
        return None
    try:
        from autopilot import classify_failure, resource_back
    except Exception:
        return None
    alive = False
    for name in names:
        why = reasons.get(name)
        if not why:
            return None
        try:
            if classify_failure(why) != "resource":
                return None
            back = resource_back(why)
        except Exception:
            return None
        if back is False:
            return False
        if back is True:
            alive = True
    return True if alive else None


def _read_json(path: Path) -> dict[str, Any] | None:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return _normalize(value) if isinstance(value, dict) else None
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return None


def _iso_age_minutes(value: Any) -> float | None:
    try:
        text = str(value or "").strip()
        if not text:
            return None
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.astimezone()
        return max(0.0, (_now() - parsed).total_seconds() / 60.0)
    except (TypeError, ValueError):
        return None


def _atomic_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix=path.stem + "_", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        os.replace(temp, path)
    finally:
        try:
            os.unlink(temp)
        except OSError:
            pass


#: 보호자가 **'아직 실패가 아니다 — 오탐 방지 재확인'** 이라 적는 상태
#: → (그 상태의 연속 실패 칸, 보호자가 쓰는 한도 이름).
#: ★ 낱말이 어긋나면 한 건도 안 걸리면서 오류도 안 난다(`[165]`) — 검증 `[307]` 이
#:   이 낱말들이 server_guard.py 에 실재하는지 매번 대 본다.
GUARD_RECHECK = {
    "degraded": ("consecutive_failures", "FAIL_LIMIT"),
    "staff-degraded": ("consecutive_staff_failures", "STAFF_FAIL_LIMIT"),
    "funnel-degraded": ("consecutive_funnel_failures", "FUNNEL_FAIL_LIMIT"),
}


def guard_limits() -> dict[str, int] | None:
    """판정 한도를 **보호자에게서 읽어 온다**(`[162]`).

    여기 3 이라 적어 두면 보호자가 한도를 바꾼 날 두 화면이 서로 다른 답을 한다.
    못 읽으면 **지어내지 않고** None 을 돌려준다 — 부르는 쪽이 '못 갈랐다'고 적는다.
    """
    try:
        webapp = str(ROOT / "webapp")
        if webapp not in sys.path:
            sys.path.insert(0, webapp)
        import server_guard
        return {name: int(getattr(server_guard, name))
                for _, name in GUARD_RECHECK.values()}
    except Exception:
        return None


def guard_verdict(guard: dict, limits: dict | None) -> dict[str, Any]:
    """보호자 상태 한 장을 감사 갈래로 옮긴다.

    ★ **판정을 새로 만들지 않는다**(`[162]`) — 보호자는 한 번 늦은 것을 죽음으로
      단정하지 않는다(`FAIL_LIMIT=3`). 그 한도를 그대로 빌려 쓴다.
    ★ 2026-08-18 실측: 그날 보호자 로그의 '지연 감지'가 **37회**인데 실제 재시작·
      실패는 **1회**였다(로그 누적 153회). 그런데 감사기가 **첫 blip 부터** P0 를
      올려, 앱 서버를 다시 띄우기만 해도(15:45:11 funnel-degraded → 15:46:13 정상
      복귀) 인계 맨 위가 빨개졌다. **경보가 대부분 가짜면 진짜 경보가 묻힌다**(`[170]`).
    ★ 내리는 것은 `*-degraded` **셋이 한도 아래일 때뿐**이다. `cooldown`(서버가
      죽은 채 재시작 과열 방지 대기) · `funnel-repairing`(이미 세 번 실패해 재등록
      중) 같은 상태는 **진짜 장애**라 그대로 P0 다 — 잘못 내리면 못 잡는 것보다
      나쁘다(`[172]`).

    아무 파일도 안 읽고 안 쓴다 — 검증이 합성 자료로 그대로 부른다(`[247]`).
    """
    state = str(guard.get("state") or "").strip().lower()
    msg = str(guard.get("message") or state or "상태 설명 없음")
    out: dict[str, Any] = {"상태": state, "재확인": None, "priority": None,
                           "id": "", "title": "", "evidence": msg}
    if state == "healthy":
        return out

    def 실패(why: str) -> dict[str, Any]:
        out.update(priority="P0", id="server-guard-failed",
                   title="앱 서버 보호자가 장애를 보고함", evidence=why)
        return out

    if state not in GUARD_RECHECK:
        return 실패(msg)

    count_key, limit_name = GUARD_RECHECK[state]
    count = guard.get(count_key)
    limit = (limits or {}).get(limit_name)
    out["재확인"] = {"칸": count_key, "횟수": count, "한도": limit}
    셀수있나 = isinstance(count, int) and not isinstance(count, bool)
    if not 셀수있나 or not isinstance(limit, int):
        # ★ 못 읽었으면 조용히 넘기지 않는다 — 오늘 하던 대로 올리되
        #   **왜 못 갈랐는지**를 적는다(`[169]`).
        return 실패("%s — 재확인 중인지 실패인지 **가르지 못했습니다**"
                    " (연속 횟수 %r · 보호자 한도 %r)." % (msg, count, limit))
    if count < limit:
        # 보호자가 **아직 실패로 안 봤다** — 여기서 먼저 실패라 부르지 않는다.
        return out
    return 실패("%s (연속 %d회 · 보호자 한도 %d회 도달)" % (msg, count, limit))


def _finding(identifier: str, priority: str, title: str, evidence: str,
             action: str, source: str) -> dict[str, str]:
    return {
        "id": identifier,
        "priority": priority,
        "title": title,
        "evidence": evidence,
        "action": action,
        "source": source,
    }


def _alert_text(item, cap=150):
    """회차 경보 한 줄 — **사전이면 그 안의 말**을 꺼낸다 (2026-08-19).

    실측: `schedule_watch` 의 `경보` 는 `{갈래·작업·무엇·어떻게}` **사전**인데 여기가
    `str(x)` 로 찍어 인계 문서에 **파이썬 사전이 그대로** 나왔다 —
    `{'갈래': '중단됨', '작업': '쿠팡업무_원본자료자동정리', '무엇': …}`.

    ★ 보기 나쁜 것보다 나쁜 것이 있다: `[:150]` 이 **사전 껍데기(약 40자)까지** 세므로
      말이 길면 **정작 사유가 잘린다.** 겉은 경보인데 왜인지는 못 읽는 자리다(`[169]`).
    ★ 만드는 쪽이 모양을 바꿔도 여기는 **오류가 안 난다**(`[165]`) — 그래서 문자열도
      그대로 받는다. 모양을 하나로 못 박으면 다음에 바뀔 때 또 조용해진다.
    """
    if isinstance(item, dict):
        text = str(item.get("무엇") or item.get("말") or "").strip()
        if not text:
            text = str(item.get("작업") or "").strip() or str(item)
        text = text.replace("**", "").strip()
        kind = str(item.get("갈래") or "").strip()
        if kind and not text.startswith("[%s]" % kind):
            text = "[%s] %s" % (kind, text)
    else:
        text = str(item)
    return text[:cap]


def _schedule_verdict(task_name):
    """스케줄러 감시자가 **이미 내린** 판정을 빌린다 — 여기서 다시 묻지 않는다.

    ★ **못 읽으면 None 이고, 그때는 예전대로 P0 다**([169]) — '확인 못 함'을
      '괜찮음'으로 치면 감사기 자신이 눈먼 채 조용해진다.
    """
    data = _read_json(REPORTS / "스케줄러_회차감시.json")
    if not isinstance(data, dict):
        return None
    for row in data.get("작업") or []:
        if isinstance(row, dict) and str(row.get("작업") or "") == task_name:
            return row
    return None


def build() -> dict[str, Any]:
    findings: list[dict[str, str]] = []
    sources: dict[str, Any] = {}

    def add(identifier: str, priority: str, title: str, evidence: str,
            action: str, source: str) -> None:
        if any(row["id"] == identifier for row in findings):
            return
        findings.append(_finding(identifier, priority, title, evidence, action, source))

    # 1) 앱 서버 보호자 — 담당자가 실제 업무를 볼 수 있는가.
    guard_path = REPORTS / "server_guard_status.json"
    guard = _read_json(guard_path)
    guard_age = _age_minutes(guard_path)
    sources["server_guard"] = {"age_minutes": guard_age, "read": guard is not None}
    if guard is None:
        add("server-guard-unreadable", "P0", "앱 서버 보호 상태를 읽지 못함",
            "server_guard_status.json이 없거나 손상되었습니다.",
            # ⚠ 옛 조치는 `python webapp/server_guard.py --once` 였는데 그 파일에는
            #   **인자 처리가 한 줄도 없다** — 깃발은 무시되고 `while True` 감시 루프가
            #   사람 창에서 영원히 돈다. 게다가 보호자가 이미 살아 있으면 singleton 에
            #   막혀 **조용히 exit 0** 이라 *"돌렸는데 아무 일도 안 났다"* 가 된다(`[169]`).
            #   되살리는 진짜 자리는 워치독의 `heal_server_guard` 다(`[263]`).
            "작업 스케줄러의 CSOS_AppServerGuard 와 reports/server_guard.log 를 확인합니다"
            " — 워치독 30분 회차가 스스로 다시 세웁니다.",
            "reports/server_guard_status.json")
    elif guard_age is None or guard_age > 5:
        add("server-guard-stale", "P0", "앱 서버 보호자 심박이 끊김",
            f"마지막 보호 기록이 {int(guard_age or 0)}분 전입니다(정상 한도 5분).",
            "작업 스케줄러의 CSOS_AppServerGuard와 server_guard.log를 확인합니다.",
            "reports/server_guard_status.json")
    else:
        verdict = guard_verdict(guard, guard_limits())
        sources["server_guard"]["state"] = verdict["상태"]
        sources["server_guard"]["recheck"] = verdict["재확인"]
        if verdict["priority"]:
            add(verdict["id"], verdict["priority"], verdict["title"],
                verdict["evidence"],
                # ★ 이 갈래에서 보호자는 **돌고 있다**(그래서 이 경보를 썼다) —
                #   다시 띄우면 singleton 에 막혀 아무 일도 안 한다. 볼 것은 로그다.
                "reports/server_guard.log 의 마지막 줄을 봅니다"
                " — 보호자는 돌고 있으므로 다시 띄우지 않습니다.",
                "reports/server_guard_status.json")

    # 2) 30분 워치독 — 입력 질문에 걸려도 로그가 오래 멈춘 것으로 잡는다.
    watchdog_path = REPORTS / "watchdog_log.txt"
    watchdog_age = _age_minutes(watchdog_path)
    sources["watchdog"] = {"age_minutes": watchdog_age, "read": watchdog_age is not None}
    watchdog_tail = ""
    try:
        with watchdog_path.open("rb") as fh:
            fh.seek(max(0, watchdog_path.stat().st_size - 12_000))
            watchdog_tail = fh.read().decode("utf-8", "replace")
    except OSError:
        pass
    if watchdog_age is None:
        add("watchdog-missing", "P0", "워치독 실행 기록이 없음",
            "30분 회차가 남겨야 할 watchdog_log.txt를 찾지 못했습니다.",
            "작업 스케줄러의 쿠팡업무_워치독을 확인합니다.", "reports/watchdog_log.txt")
    elif watchdog_age > 75:
        last_watchdog_line = next(
            (line for line in reversed(watchdog_tail.splitlines()) if line.strip()), "")
        waiting = "지금 내릴까요" in last_watchdog_line or "(y = 내림" in last_watchdog_line
        if waiting:
            # 입력 대기는 잠과 무관한 진짜 멈춤이다 — 한 글자도 안 건드린다([172]).
            add("watchdog-stale", "P0", "워치독이 입력 대기에서 멈춤",
                f"마지막 로그가 {int(watchdog_age)}분 전이며 끝부분에 사람 답변을"
                " 기다리는 문구가 있습니다.",
                "무인 회차에서는 확인창을 띄우지 말고 재시작을 보류 기록으로"
                " 남기도록 고칩니다.",
                "reports/watchdog_log.txt")
        else:
            slept, why = _sleep_minutes_since(watchdog_age)
            sources["watchdog"]["slept_minutes"] = slept
            sources["watchdog"]["sleep_note"] = why
            awake = None if slept is None else max(0.0, watchdog_age - slept)
            if awake is not None and awake <= 75:
                # 깨어 있던 시간으로는 안 밀렸다 — 경보가 아니라 **알림**이다([170]).
                # 그래도 조용히 빼지 않는다([169]): 몇 분을 잤는지 숫자로 말한다.
                add("watchdog-slept", "P2", "워치독 공백은 이 PC 가 잠들어 있던 시간",
                    f"마지막 로그가 {int(watchdog_age)}분 전이지만 그중"
                    f" {int(slept)}분은 이 PC 가 Modern Standby 로 자고 있었습니다"
                    f"(깨어 있던 공백 {int(awake)}분 · 한도 75분)."
                    " 워치독이 멈춘 것이 아니라 기계가 안 깨어 있었습니다.",
                    "그대로 두면 됩니다 — 깨어나면 다음 회차가 스스로 돕니다.",
                    "reports/watchdog_log.txt")
            else:
                extra = ("입니다." if slept is not None else
                         f"입니다. 잠든 시간은 갈라내지 못했습니다({why}) —"
                         " '자고 있었다'는 뜻이 아닙니다.")
                if slept:
                    extra = (f"이고 그중 {int(slept)}분은 잠이었습니다"
                             f"(깨어 있던 공백 {int(awake)}분 · 한도 75분).")
                add("watchdog-stale", "P0", "워치독 30분 회차가 멈춤",
                    f"마지막 로그가 {int(watchdog_age)}분 전" + extra,
                    "python watchdog.py   # 그 전에 tasklist 로 앞 회차가 도는지 본다",
                    "reports/watchdog_log.txt")

    # 3) 일일 대조 — 실패를 '오늘 실행됨'으로 세지 않는다.
    #
    # ★ 그런데 **실패한 코드가 이미 안 도는 경우**를 가른다 (2026-08-19).
    #   실측: 한 사건(09:50 회차 실패)이 인계 '먼저 처리할 것'에 **세 줄**로 떴는데
    #   그중 `schedule_watch` 만 "그 뒤 코드가 바뀌었다(daily_run.py 10:17)"를 알고
    #   있었고, 여기는 그것을 모른 채 **P0** 를 올렸다. 같은 파일을 보는 두 화면이
    #   서로 다르게 말하면 사람은 이미 고쳐진 것을 고치러 간다([172]).
    #   `[110]` 이 `schedule_watch` 에 세운 `고침대기` 를 여기서도 **빌린다**([162]) —
    #   여기서 스케줄러를 다시 묻지 않는다([168]). 회차가 써 둔 판정을 읽기만 한다.
    daily_path = REPORTS / ".daily_run.progress.json"
    daily = _read_json(daily_path)
    daily_age = _age_minutes(daily_path)
    sources["daily_run"] = {"age_minutes": daily_age, "read": daily is not None}
    if daily is None:
        add("daily-run-unreadable", "P1", "일일 대조 진행 자국을 읽지 못함",
            "마지막 완주·실패 단계를 판정할 근거가 없습니다.",
            # ⚠ `daily_run.py` 에는 **인자 처리가 한 줄도 없다**(실측 2026-08-26) —
            #   `--status` 를 붙여 부르면 그 깃발이 무시되고 **회차가 통째로 돈다**.
            #   임의 시각에 116분짜리 Z: 회차를 띄우면 예약 회차와 부딪혀 한쪽이
            #   잠금에 막혀 죽는다(2026-08-16 겹침 사고). 조치 칸에는 **붙여넣어
            #   도는 명령만** 넣는다(`[247]`) — 여기서는 읽기 전용 감시자를 준다.
            "python schedule_watch.py --print   # 회차가 실제로 돌았는지부터 본다",
            "reports/.daily_run.progress.json")
    else:
        state = str(daily.get("상태") or "")
        if state == "실패":
            # ★ **'중단'과 '완주했지만 단계가 실패'는 다른 사실이다** (2026-08-26 실사고).
            #   실측: 회차가 **116.1분을 끝까지 돌고**(`끝까지실행: True` · 단계
            #   `(회차 끝)`) 단계 13개만 실패했는데 화면은
            #   *"오늘 일일 대조가 **중단됨** · 실패 원인: **원인 없음**"* 이라 말하고
            #   조치로 *"실패한 합성검증을 고쳐라"* 를 줬다. **셋 다 틀렸다** —
            #   ① 중단이 아니다 ② 원인은 바로 그 파일의 `실패단계` 에 이름으로 다
            #   적혀 있다 ③ 합성검증은 **0단계**라 그것이 막히면 회차가 거기서
            #   끝난다(`[304]`) — 47단계를 지났다는 것이 곧 **관문은 통과했다**는
            #   증거다. 그러니 사람은 **멀쩡한 검증을 고치러 간다**(`[172]`).
            #   봉투에 답이 들어 있는데 판정 코드가 버리는 자리다(`[365]`·`[289]`).
            # ★ 예외로 죽은 갈래(`오류`·`오류유형` 이 있고 `끝까지실행` 이 없다)는
            #   **한 글자도 안 바꿨다**(`[172]` — 좁히는 것도 고장이다).
            steps_failed = [str(x) for x in (daily.get("실패단계") or []) if str(x)]
            ran_through = bool(daily.get("끝까지실행")) and bool(steps_failed)
            if ran_through:
                # ★ **같은 사유를 열세 번 늘어놓지 않는다**(`[170]` — 길면 안 읽히고,
                #   안 읽히면 없는 설명이다). 실측 2026-08-26 은 13개가 **전부 같은
                #   이유**(Z: 를 그 순간 못 잡았다)였다 — 묶어서 세면 짧아지면서
                #   *"한 가지 원인이 열세 단계를 죽였다"* 는 **더 참된 사실**을 말한다.
                # ★ **조용히 자르지 않는다**(`[273]`) — 못 실은 묶음은 숫자로 적는다.
                reasons = _daily_step_reasons(steps_failed)
                # ★ **지나간 자원 실패를 '굳었다'고 부르지 않는다**(`[424]`).
                #   갈래가 전부 `resource` 이고 그 자원이 지금 살아 있으면
                #   사람이 지금 할 일이 없다 — **다음 회차가 답한다**.
                recovered = _steps_resource_recovered(reasons, steps_failed)
                groups = {}
                for nm in steps_failed:
                    groups.setdefault((reasons.get(nm) or "")[:90], []).append(nm)
                order = sorted(groups.items(), key=lambda kv: -len(kv[1]))
                bits, left = [], 0
                for key, members in order:
                    if len(bits) >= 3:
                        left += len(members)
                        continue
                    names = ", ".join(members[:4])
                    if len(members) > 4:
                        names += " 외 %d개" % (len(members) - 4)
                    bits.append(("%s ← %d개(%s)" % (key, len(members), names))
                                if key else "사유 못 읽음 %d개(%s)" % (len(members), names))
                title = "일일 대조가 완주했지만 단계 %d개가 실패했다" % len(steps_failed)
                why = "%s · 회차는 %s분 만에 **끝까지 돌았다**(관문은 통과했다) · %s%s" % (
                    daily.get("시각") or "시각 없음",
                    daily.get("경과분") if daily.get("경과분") is not None else "?",
                    " / ".join(bits),
                    (" / 그 밖 %d개는 여기 못 실었다" % left) if left else "")
                action = ("python schedule_watch.py --print"
                          "   # 사유 전문은 reports/종합리포트_*.md 의 그 단계 블록에 있다")
                source = "reports/.daily_run.progress.json"
            else:
                recovered = None
                title = "오늘 일일 대조가 중단됨"
                why = "%s · 실패 원인: %s" % (
                    daily.get("시각") or "시각 없음",
                    daily.get("오류") or daily.get("오류유형") or "원인 없음")
                action = "실패한 합성검증을 고친 뒤 일일 대조를 다시 실행합니다."
                source = "reports/.daily_run.progress.json"
            verdict = _schedule_verdict("쿠팡업무_일일자동대조") or {}
            kind = str(verdict.get("갈래") or "")
            if kind in ("고침대기", "뒤에됨"):
                # ★ '고쳐졌다'고 말하지 않는다([110]) — 말할 수 있는 것은 '그 뒤
                #   코드가 바뀌었다' 까지다. 무관한 이유로 건드렸을 수도 있다.
                #   **다음 회차가 답한다.** 그러니 P0 가 아니라 지켜볼 것이다.
                add("daily-run-fix-pending", "P2",
                    "일일 대조가 실패했지만 그 뒤 코드가 바뀌었다",
                    "%s · %s" % (why, verdict.get("말") or "스케줄러 감시자가 그렇게 봤습니다"),
                    "python schedule_watch.py --print   # 다음 예정 회차가 답한다",
                    "reports/스케줄러_회차감시.json")
            elif recovered is True:
                # ★ **자원이 그때 끊긴 것이다** — 코드가 깨진 것이 아니다(`[424]`).
                #   조용히 빼지 않는다(`[169]`): 실패 사실과 사유는 그대로 싣고
                #   무게만 내린다. **"고쳐졌다"고 말하지 않는다**(`[322]`) —
                #   말할 수 있는 것은 "그 자원이 지금은 살아 있다" 까지다.
                add("daily-run-resource-back", "P2",
                    "일일 대조 단계 실패는 그때 공유폴더를 못 잡은 것이다",
                    "%s · 실패한 단계가 **모두 자원 탓**이고 그 자원은 **지금 살아 있다** — 코드가 깨진 것이 아니다. 다음 회차가 답한다." % why,
                    action, source)
            else:
                # ★ 무게는 **안 내린다**(`[172]`) — 단계 13개가 안 돈 것은
                #   그 자체로 P0 다. 고친 것은 **무엇이·왜·어디를 보라**뿐이다.
                add("daily-run-failed", "P0", title, why, action, source)
        elif daily_age is not None and daily_age > 20 * 60:
            # ★ **늦은 것과 아직 예정이 안 온 것은 다른 사실이다** (2026-08-21 실사고).
            #   이 회차는 하루 한 번 09:50 이라, 어제 12:20 에 끝났으면 오늘 08:20~09:50
            #   이 그대로 20시간을 넘는다 — 늦은 것이 아니라 순서가 안 온 것이다.
            #   같은 사건을 인계 문서와 여기가 **두 목소리로** 울리고 있었으므로 판정을
            #   한 곳에서 빌린다(`[162]`·`[322]`) — 스케줄러 사실은 회차가 이미 써 뒀다.
            #   **못 갈랐으면(None) 예전 그대로 P1** 이다(`[169]`).
            not_due = None
            try:
                import schedule_watch
                not_due = schedule_watch.due_state(
                    "daily_run.py",
                    done_at=datetime.fromtimestamp(daily_path.stat().st_mtime))
            except Exception:
                not_due = None
            if not (not_due or {}).get("아직"):
                add("daily-run-stale", "P1", "일일 대조 완주 기록이 하루 가까이 갱신되지 않음",
                    f"마지막 진행 기록이 {daily_age / 60:.1f}시간 전입니다.",
                    "python session_handoff.py --check", "reports/.daily_run.progress.json")

    # 4) 스케줄러 감시자가 스스로 낡았는지와 마지막 경보.
    schedule_path = REPORTS / "스케줄러_회차감시.json"
    schedule = _read_json(schedule_path)
    schedule_age = _age_minutes(schedule_path)
    sources["schedule_watch"] = {"age_minutes": schedule_age, "read": schedule is not None}
    if schedule is None or schedule_age is None:
        add("schedule-watch-unreadable", "P1", "스케줄러 회차 감시를 읽지 못함",
            "자동 회차가 실제로 돌았는지 확인할 보고서가 없습니다.",
            "python schedule_watch.py", "reports/스케줄러_회차감시.json")
    elif schedule_age > 90:
        add("schedule-watch-stale", "P1", "스케줄러 감시 보고서가 낡음",
            f"마지막 조회가 {int(schedule_age)}분 전이라 이후 실패를 반영하지 못합니다.",
            "python schedule_watch.py", "reports/스케줄러_회차감시.json")
    if schedule:
        alerts = schedule.get("경보") or []
        if alerts:
            sample = " · ".join(_alert_text(x) for x in alerts[:3])
            add("scheduled-round-alert", "P0", "자동 회차 실패 경보가 남아 있음",
                sample, "python schedule_watch.py --print", "reports/스케줄러_회차감시.json")

    # 5) 최근 밴드 글 수정·재수집. 긁기는 로그인 경계지만 누락 사실은 앱이 말한다.
    recollect_path = REPORTS / "밴드_재수집.json"
    recollect = _read_json(recollect_path)
    sources["band_recollect"] = {"age_minutes": _age_minutes(recollect_path),
                                  "read": recollect is not None}
    if recollect:
        pending = recollect.get("손볼것") or []
        changed = (recollect.get("최근변경") or {}).get("바뀐글") or recollect.get("바뀐글") or []
        acknowledged = bool(recollect.get("확인함"))
        if pending:
            add("band-recollect-pending", "P1", "최근 밴드 글 재수집이 덜 끝남",
                f"로그인된 탭에서 다시 받아야 할 묶음 {len(pending)}개가 남았습니다. "
                + " · ".join(str(x)[:120] for x in pending[:2]),
                "앱의 Band 로그인 상태를 확인하고 준비된 재수집 목록만 수집합니다.",
                "reports/밴드_재수집.json")
        if changed and not acknowledged:
            add("band-changes-unacknowledged", "P1", "수정된 밴드 글을 아직 확인하지 않음",
                f"최근 재수집에서 내용이 달라진 글 {len(changed)}건이 확인 대기입니다.",
                "python band/recollect.py --print  (확인 후 --ack)",
                "reports/밴드_재수집.json")

    # 6) 오류 사전 — 알려진 회귀와 아직 이름 없는 오류를 분리한다.
    error_path = REPORTS / "오류_사전.json"
    errors = _read_json(error_path)
    sources["error_book"] = {"age_minutes": _age_minutes(error_path), "read": errors is not None}
    if errors:
        regressions = errors.get("회귀") or []
        new_errors = errors.get("새오류") or []
        if regressions:
            total = sum(int(x.get("건수") or 0) for x in regressions if isinstance(x, dict))
            sample = " · ".join(str(x.get("무엇") or x.get("지문") or "")[:110]
                                for x in regressions[:3] if isinstance(x, dict))
            add("error-regression", "P1", "고쳤던 오류가 다시 발생함",
                f"회귀 {total}건 · {sample}", "python error_book.py --print",
                "reports/오류_사전.json")
        if new_errors:
            total = sum(int(x.get("건수") or 0) for x in new_errors if isinstance(x, dict))
            add("error-unclassified", "P2", "설명 규칙이 없는 새 오류가 있음",
                f"새 오류 {len(new_errors)}종류 · {total}건입니다.",
                "python error_book.py --print", "reports/오류_사전.json")

    # 7) 대표 보고 검증은 숫자를 못 센 것과 0을 가르는 마지막 문턱이다.
    exec_guard_path = REPORTS / "대표보고_검증.json"
    exec_guard = _read_json(exec_guard_path)
    sources["executive_guard"] = {"age_minutes": _age_minutes(exec_guard_path),
                                   "read": exec_guard is not None}
    if exec_guard is None:
        add("executive-guard-missing", "P1", "대표 보고 숫자 검증본이 없음",
            "일일 회차가 중단돼 잔여 미청구·미수금의 근거 열 충족도를 확인하지 못했습니다.",
            "python exec_report_guard.py", "reports/대표보고_검증.json")
    else:
        warnings = exec_guard.get("먼저볼것") or []
        unknown = exec_guard.get("못물어봄") or []
        if warnings:
            add("executive-guard-warning", "P1", "대표 보고에 근거가 약한 숫자가 있음",
                " · ".join(str(x)[:170] for x in warnings[:3]),
                "python exec_report_guard.py --print", "reports/대표보고_검증.json")
        if unknown:
            add("executive-guard-unknown", "P1", "대표 보고 검증이 일부 근거를 못 읽음",
                " · ".join(str(x)[:170] for x in unknown[:2]),
                "python exec_report_guard.py --print", "reports/대표보고_검증.json")

    # 8) 댓글 사각지대는 0건이 아니라 '아직 안 본 건'이다.
    cancel_path = REPORTS / "접수취소_확인.md"
    try:
        cancel_text = _repair_text(cancel_path.read_text(encoding="utf-8"))
    except OSError:
        cancel_text = ""
    match = re.search(r"놓쳤을 수 있는 글:\s*\*\*(\d[\d,]*)건\*\*\s*/\s*전체\s*(\d[\d,]*)건", cancel_text)
    if match:
        blind, total = (int(x.replace(",", "")) for x in match.groups())
        sources["cancel_comment_coverage"] = {"blind": blind, "total": total}
        if blind:
            add("cancel-comment-blind", "P1", "접수취소 댓글 사각지대가 남음",
                f"댓글을 다 읽지 못한 글 {blind:,}건 / 전체 {total:,}건입니다.",
                "우선순위 목록만 Band 댓글 재수집 대상으로 보냅니다.",
                "reports/접수취소_확인.md")

    # 9) AI 인계는 업무 실행 경로와 분리한다. 그래도 Claude 우선 경로가 죽으면 알려 준다.
    dispatch_path = REPORTS / "agent_dispatch_status.json"
    dispatch = _read_json(dispatch_path)
    sources["agent_dispatch"] = {"age_minutes": _age_minutes(dispatch_path),
                                  "read": dispatch is not None}
    if dispatch and dispatch.get("selected") != "claude":
        claude = (dispatch.get("agents") or {}).get("claude") or {}
        add("claude-fallback", "P2", "Claude Code 인계가 Codex 폴백 상태",
            str(claude.get("reason") or dispatch.get("note") or "Claude Code 사용 불가"),
            # ⚠ `agent_dispatch.py` 의 깃발은 `--run-ticket`·`--local-returncode`·
            #   `--supersede-queued`·`--status` 넷뿐이다 — `--route --force` 는
            #   argparse 가 `unrecognized arguments` 로 **바로 죽인다**(`[247]`).
            "Claude Code 설치·로그인을 복구한 뒤 python agent_dispatch.py --status",
            "reports/agent_dispatch_status.json")

    # 9-1) Claude 5시간 창과 Codex 사용 한도는 **서로 다른 계기**다. 하나만 보고
    # 'AI 사용 가능'이라 하면 실패표만 쌓인다([169]). 판정은 credit_window 자국을 읽는다.
    credit_path = REPORTS / "크레딧_창.json"
    credit = _read_json(credit_path)
    sources["credit_window"] = {"age_minutes": _age_minutes(credit_path),
                                "read": credit is not None}
    codex_credit = (((credit or {}).get("agents") or {}).get("codex") or {})
    if codex_credit.get("갈래") == "소진":
        try:
            reset = datetime.fromtimestamp(
                float(codex_credit.get("resetsAt") or 0)).astimezone().strftime("%m-%d %H:%M")
        except Exception:
            reset = "확인된 재개 시각"
        add("codex-credit-exhausted",
            "P1" if (dispatch or {}).get("selected") == "codex_pending" else "P2",
            "Codex 자동 인계 크래딧 대기",
            "Codex 실행표가 사용 재개 시각을 %s 로 명시했습니다." % reset,
            "새 실패표를 만들지 않고 기존 대기표를 보존합니다. 충전 뒤 30분 회차가 자동 재개합니다.",
            "reports/크레딧_창.json")

    # 10) ERP 공식 조회 API. IP 미등록은 키 오류가 아니며 브라우저 로그인 한 번이
    # 필요한 사람 경계다. 수집이 없다는 사실을 조용히 숨기지 않는다.
    erp_ip_path = REPORTS / "ERP_IP_등록필요.md"
    erp_api_path = REPORTS / "erp_api_latest.json"
    erp_api = _read_json(erp_api_path)
    sources["erp_api"] = {"age_minutes": _age_minutes(erp_api_path),
                          "read": erp_api is not None}
    if erp_ip_path.exists() and (_age_minutes(erp_ip_path) or 0) < 24 * 60:
        add("erp-api-ip-unregistered", "P1", "ERP API 허용 IP 등록 필요",
            "현재 공인 IP가 이카운트 API 허용 목록에 없어 호출 전에 안전하게 멈췄습니다.",
            "로그인된 이카운트의 API인증키발급 > IP등록에 보고서의 현재 IP를 저장합니다.",
            "reports/ERP_IP_등록필요.md")
    elif erp_api and not erp_api.get("ok"):
        # ★ 조치는 갈래마다 다르다 — 'IP 미등록' 에 `--force` 를 권하면 사람이
        #   같은 실패를 한 번 더 부르고 원인은 그대로 남는다. 수집기가 골라 둔 조치가
        #   있으면 그것을 쓰고, 없을 때만(모름) 예전 문구로 돌아간다.
        add("erp-api-failed", "P1", "ERP 공식 API 수집 실패",
            str(erp_api.get("error") or "원인 설명 없음")[:220],
            str(erp_api.get("조치") or "python erp_api_collect.py --force")[:200],
            "reports/erp_api_latest.json")

    order = {"P0": 0, "P1": 1, "P2": 2}
    findings.sort(key=lambda row: (order.get(row["priority"], 9), row["id"]))
    summary = {p: sum(1 for row in findings if row["priority"] == p) for p in order}
    summary["total"] = len(findings)
    state = "critical" if summary["P0"] else "warning" if summary["P1"] else "attention" if summary["P2"] else "healthy"
    return {
        "version": VERSION,
        "generated_at": _now().isoformat(timespec="seconds"),
        "state": state,
        "healthy": not findings,
        "summary": summary,
        "findings": findings,
        "sources": sources,
        "engine": "system_audit.py",
        "scope": "cached-state-only",
    }


def render_markdown(report: dict[str, Any]) -> str:
    summary = report.get("summary") or {}
    lines = [
        "# 시스템·업무 진단",
        "",
        f"- 만든 때: {report.get('generated_at')}",
        f"- 상태: **{report.get('state')}** · P0 {summary.get('P0', 0)} · "
        f"P1 {summary.get('P1', 0)} · P2 {summary.get('P2', 0)}",
        "- 판정 정본: `python system_audit.py --print` (앱·Claude Code·Codex 공용)",
        "",
    ]
    if not report.get("findings"):
        lines += ["지금 보고할 문제를 찾지 못했습니다.", ""]
    for priority in ("P0", "P1", "P2"):
        rows = [row for row in report.get("findings") or [] if row.get("priority") == priority]
        if not rows:
            continue
        lines += [f"## {priority}", ""]
        for row in rows:
            lines += [
                f"- **{row['title']}**",
                f"  - 근거: {row['evidence']}",
                f"  - 조치: {row['action']}",
                f"  - 출처: `{row['source']}`",
            ]
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def write_report(report: dict[str, Any] | None = None) -> dict[str, Any]:
    value = report or build()
    _atomic_text(OUT_JSON, json.dumps(value, ensure_ascii=False, indent=2) + "\n")
    _atomic_text(OUT_MD, render_markdown(value))
    return value


def read_cached() -> dict[str, Any]:
    value = _read_json(OUT_JSON)
    if value:
        value["report_age_minutes"] = _age_minutes(OUT_JSON)
        return value
    return {
        "version": VERSION,
        "generated_at": None,
        "state": "unknown",
        "healthy": False,
        "summary": {"P0": 0, "P1": 0, "P2": 0, "total": 0},
        "findings": [],
        "engine": "system_audit.py",
        "error": "진단 보고서가 아직 만들어지지 않았습니다.",
    }


def handoff_lines(limit: int = 5) -> list[str]:
    """세션 인계가 같은 캐시를 읽도록 하는 작고 부작용 없는 손잡이."""
    report = read_cached()
    return [f"[{row['priority']}] {row['title']} — {row['evidence']}"
            for row in (report.get("findings") or [])[:limit]]


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass
    parser = argparse.ArgumentParser(description="앱·Claude Code·Codex 공용 시스템·업무 진단")
    parser.add_argument("--print", action="store_true", help="보고서를 갱신하고 사람이 읽는 요약 출력")
    parser.add_argument("--json", action="store_true", help="보고서를 갱신하고 JSON 출력")
    args = parser.parse_args(argv)
    report = write_report()
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    elif args.print:
        print(render_markdown(report), end="")
    else:
        s = report["summary"]
        print(f"시스템·업무 진단 {report['state']} · P0 {s['P0']} · P1 {s['P1']} · P2 {s['P2']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
