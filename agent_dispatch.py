# -*- coding: utf-8 -*-
"""Claude Code 우선·Codex 폴백 작업 인계 큐.

실행 버튼이 실제 업무 스크립트를 한 번만 실행하는 동안, AI 검토·후속조치 요청은
별도 큐에 남긴다. AI CLI를 임의로 실행해 관리대장을 동시에 건드리지 않는다.
Claude Code가 사용 불가(크레딧/인증/명령 없음)이면 Codex 요청으로 자동 전환한다.

큐는 ``reports/agent_dispatch`` 아래에만 두며 비밀값·대화 원문은 저장하지 않는다.
"""
from __future__ import annotations

import json
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from proc_guard import background_popen_kwargs, run_tree


ROOT = Path(__file__).resolve().parent
REPORT_DIR = ROOT / "reports" / "agent_dispatch"
STATUS_PATH = ROOT / "reports" / "agent_dispatch_status.json"
PROBE_TIMEOUT_SECONDS = 4
ROUTE_CACHE_SECONDS = 60
AGENT_TIMEOUT_SECONDS = 60 * 60

#: ★ `not logged in` 이 빠져 있었다 (2026-08-13 실측). 로그아웃 상태의 CLI 는
#  "Not logged in · Please run /login" 을 주는데 그 문구가 이 목록 어디에도 안 맞아
#  **설계된 Codex 폴백이 한 번도 안 떴다** — 티켓은 그냥 failed 로 끝났고 codex 는
#  멀쩡히 깔린 채 standby 였다. 낱말 하나가 빠지면 폴백은 조용히 없는 것이 된다.
_UNAVAILABLE_RE = re.compile(
    r"credit|quota|usage|rate.?limit|billing|insufficient|not.?authenticated|"
    r"not.?logged.?in|/login|please log ?in|sign(ed)? ?in|로그인|"
    r"access is denied|permission denied|not recognized|not found|인증|크레딧|할당량|권한",
    re.I,
)

#: 로그인 여부를 **실행 없이** 묻는 길. claude 2.1.222 는 `auth status` 가 JSON 을 준다.
#  codex 쪽은 확인된 명령이 없어 넣지 않는다 — 없는 손잡이를 지어내면 멀쩡한 폴백이 죽는다.
AUTH_PROBE = {"claude": ("auth", "status")}

#: ★ **로그인 안 된 CLI 는 exit 0 을 준다** (2026-08-13 실측 · claude 2.1.222).
#:   `claude -p "..."` 가 `Not logged in · Please run /login` 한 줄만 찍고 **성공으로
#:   끝난다.** 그러면 아래 `status` 가 `done` 으로 적히고, 폴백 조건도
#:   `returncode != 0` 을 전제하므로 **Codex 로도 안 넘어간다** — 인계가 아무 일도
#:   안 하고 완료가 된다. **실패가 성공처럼 보이는 자리**다(`[171]` 과 같은 모양).
#: ★ 낱말이 아니라 **CLI 가 뱉는 정형 문구**로 가른다. 사람도 AI 도 본문에서 '로그인'을
#:   말할 수 있으므로(밴드·이카운트 로그인은 이 프로젝트의 일상이다) 짧은 응답일 때만
#:   본다 — 정상 답을 실패로 만드는 것이 못 잡는 것보다 나쁘다.
_NOT_LOGGED_RE = re.compile(r"please run\s*/login|^\s*not logged in\b", re.I | re.M)
_NOT_LOGGED_MAX = 400


def _looks_not_logged_in(text: str) -> bool:
    """정형 문구 **그리고** 짧은 응답일 때만 '로그인 안 됨'이라고 본다."""
    t = (text or "").strip()
    return bool(t) and len(t) < _NOT_LOGGED_MAX and bool(_NOT_LOGGED_RE.search(t))


def _atomic_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix=path.stem + "_", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(value, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def _clean_message(value: str, limit: int = 180) -> str:
    return " ".join((value or "").split())[:limit]


def resolve_agent_executable(name: str) -> str:
    """WindowsApps의 실행 불가 별칭을 피하고 실제 설치 실행 파일을 찾는다."""
    if name not in ("claude", "codex"):
        raise ValueError(f"unknown agent: {name}")
    candidates: list[str] = []
    direct = shutil.which(name)
    if direct:
        candidates.append(direct)
    local = os.environ.get("LOCALAPPDATA", "")
    roaming = os.environ.get("APPDATA", "")
    # Only known CLI install layouts belong here.  Desktop GUI executables can share
    # the same basename (Claude.exe/Codex.exe); starting one would raise the app and
    # interrupt the operator instead of consuming a background ticket.
    patterns = {
        "codex": [
            os.path.join(local, "OpenAI", "Codex", "bin", "*", "codex.exe"),
        ],
        "claude": [
            os.path.join(roaming, "Claude", "claude-code", "*", "claude.exe"),
        ],
    }[name]
    for pattern in patterns:
        candidates.extend(sorted(glob.glob(pattern, recursive=True), reverse=True))
    seen = set()
    for candidate in candidates:
        candidate = os.path.abspath(candidate)
        if candidate.lower() in seen or not os.path.isfile(candidate):
            continue
        seen.add(candidate.lower())
        # Microsoft Store 별칭은 존재해도 Access denied가 날 수 있으므로 실제 파일 우선.
        low = candidate.lower().replace("/", "\\")
        if "windowsapps" in low:
            continue
        if name == "claude" and "\\programs\\claude\\" in low and \
                "\\claude-code\\" not in low:
            continue
        if name == "codex" and "\\programs\\codex\\" in low and "\\bin\\" not in low:
            continue
        return candidate
    return ""


def auth_state(name: str, executable: str) -> tuple[str, str]:
    """로그인했나 — **셋으로** 답한다: `로그인` · `로그아웃` · `확인못함`.

    ★ `--version` 은 로그인을 안 봐도 0 을 준다. 그래서 2026-08-13 실측에서
      `route_status()` 가 `state: ready` 를 말하는 동안 실제 티켓은 전부
      "Not logged in · Please run /login" 으로 죽고 있었다. **계기가 초록이면
      아무도 안 본다** — 이 프로젝트가 반복해 당한 '실패가 성공처럼 보이는 자리'다.

    ★ **못 읽은 것을 로그아웃이라고 하지 않는다.** 옛 CLI 에는 `auth status` 가
      없을 수 있는데, 없다고 로그아웃으로 치면 멀쩡히 돌던 인계가 통째로 Codex 로
      샌다. 거꾸로 못 읽었다고 `ready` 라 하면 지금 이 고장이 그대로 남는다.
      그래서 모르면 **모른다고** 하고, 부르는 쪽이 그 사실을 그대로 적는다(`[169]`).

    비밀값은 읽지 않는다 — 이 명령은 토큰이 아니라 `loggedIn` 참/거짓만 돌려준다.
    """
    sub = AUTH_PROBE.get(name)
    if not sub:
        return "확인못함", "로그인을 묻는 명령이 확인되지 않았다"
    try:
        result = run_tree([executable, *sub], cwd=ROOT,
                          timeout=PROBE_TIMEOUT_SECONDS, drain_timeout=5,
                          output_limit=4000)
    except OSError as exc:
        return "확인못함", _clean_message(str(exc))
    if result.timed_out:
        return "확인못함", "로그인 확인 시간 초과"
    text = ((result.stdout or "") + " " + (result.stderr or "")).strip()
    # 종료 코드로 판정하지 않는다 — 로그아웃일 때도 1 이고 명령이 없을 때도 1 이다.
    try:
        body = json.loads(text[text.index("{"):text.rindex("}") + 1])
    except (ValueError, json.JSONDecodeError):
        return "확인못함", _clean_message(text) or "로그인 상태를 읽지 못했다"
    if not isinstance(body.get("loggedIn"), bool):
        return "확인못함", "loggedIn 값이 없다"
    if body["loggedIn"]:
        return "로그인", str(body.get("authMethod") or "")
    return "로그아웃", "로그인이 안 돼 있다 — 사람이 `claude auth` 로 한 번 로그인해야 한다"


def probe_agent(name: str) -> dict[str, str]:
    """Return a small, non-secret availability record without launching any task."""
    command = {"claude": "claude", "codex": "codex"}.get(name)
    if not command:
        raise ValueError(f"unknown agent: {name}")
    executable = resolve_agent_executable(name)
    if not executable:
        return {"agent": name, "state": "unavailable", "reason": "실행 명령을 찾지 못함"}
    try:
        result = run_tree([executable, "--version"], cwd=ROOT,
                          timeout=PROBE_TIMEOUT_SECONDS, drain_timeout=5,
                          output_limit=4000)
    except OSError as exc:
        return {"agent": name, "state": "unavailable", "reason": _clean_message(str(exc))}

    if result.timed_out:
        return {"agent": name, "state": "unavailable", "reason": "버전 확인 시간 초과"}

    message = _clean_message((result.stdout or "") + " " + (result.stderr or ""))
    if result.returncode == 0:
        # 버전이 뜬다고 일할 수 있는 것이 아니다 — 로그인을 한 번 더 묻는다.
        state, why = auth_state(name, executable)
        if state == "로그아웃":
            return {"agent": name, "state": "unavailable", "reason": why}
        if state == "확인못함":
            # 모르는 것을 아는 척하지 않는다. 일은 시켜 보되 **모른다고 적어 둔다** —
            # 그래야 실패했을 때 "로그인을 확인 못 한 채 보냈다"가 기록에 남는다.
            return {"agent": name, "state": "ready",
                    "reason": "%s · 로그인 확인못함(%s)" % (message or "사용 가능", why)}
        return {"agent": name, "state": "ready",
                "reason": "%s · 로그인 확인됨%s" % (message or "사용 가능",
                                                  (" (%s)" % why) if why else "")}
    if _UNAVAILABLE_RE.search(message):
        reason = "크레딧·인증·권한 또는 실행 환경을 확인해야 함"
    else:
        reason = message or f"버전 확인 실패(코드 {result.returncode})"
    return {"agent": name, "state": "unavailable", "reason": reason}


def _credit_block(name: str) -> tuple[bool, str]:
    """에이전트별 크래딧 문. 판정은 credit_window 한 곳에서 빌린다([162])."""
    try:
        import credit_window
        # Codex 근거는 **이 분담 큐**에 있다. 합성검증·격리 실행이 REPORT_DIR 을
        # 바꾸면 그 격리 큐만 읽어야 실측 실패가 시험 결과를 오염시키지 않는다([247]).
        st = (credit_window.codex_state(report_dir=str(REPORT_DIR))
              if name == "codex" else credit_window.state())
        blocked = st.get("갈래") == "소진"
        if blocked:
            when = datetime.fromtimestamp(st["resetsAt"]).strftime("%m-%d %H:%M")
            return True, "%s 크래딧 소진 — %s 이후 자동 재개" % (
                "Codex" if name == "codex" else "Claude", when)
    except Exception as exc:
        # 못 읽었다고 멀쩡한 에이전트를 막지는 않는다. 실제 실행 실패는 표에 남는다.
        return False, "크래딧 확인못함(%s)" % type(exc).__name__
    return False, ""


def _cached_route() -> dict[str, Any] | None:
    """Reuse a recent probe result so normal status refreshes stay instant."""
    try:
        value = json.loads(STATUS_PATH.read_text(encoding="utf-8"))
        checked_at = datetime.fromisoformat(str(value.get("checked_at", "")))
        if datetime.now() - checked_at > timedelta(seconds=ROUTE_CACHE_SECONDS):
            return None
        if value.get("selected") and isinstance(value.get("agents"), dict):
            return {key: value[key] for key in ("primary", "selected", "note", "agents", "checked_at")}
    except (OSError, ValueError, TypeError, json.JSONDecodeError, KeyError):
        pass
    return None


def route_status(*, force: bool = False) -> dict[str, Any]:
    """Claude Code first, then Codex.  Never executes either AI by itself."""
    if not force:
        cached = _cached_route()
        if cached:
            return cached
    claude = probe_agent("claude")
    c_blocked, c_why = _credit_block("claude")
    if c_blocked:
        claude = {"agent": "claude", "state": "unavailable", "reason": c_why}
    if claude["state"] == "ready":
        selected, note = "claude", "Claude Code 우선"
        x_blocked, x_why = _credit_block("codex")
        codex = {"agent": "codex", "state": "unavailable" if x_blocked else "standby",
                 "reason": x_why or "Claude Code 사용 가능 시 대기"}
    else:
        codex = probe_agent("codex")
        x_blocked, x_why = _credit_block("codex")
        if x_blocked:
            codex = {"agent": "codex", "state": "unavailable", "reason": x_why}
    if claude["state"] != "ready" and codex["state"] == "ready":
        selected, note = "codex", "Claude Code 사용 불가 → Codex 폴백"
    elif claude["state"] != "ready":
        selected, note = "codex_pending", "Claude Code 사용 불가 → Codex 작업 요청 대기"
    route = {
        "primary": "claude",
        "selected": selected,
        "note": note,
        "agents": {"claude": claude, "codex": codex},
        "checked_at": datetime.now().isoformat(timespec="seconds"),
    }
    _atomic_json(STATUS_PATH, route)
    return route


def enqueue(task_key: str, title: str, args: list[str], *, extra: str = "") -> dict[str, Any]:
    """Persist an AI review/follow-up request and return the selected route.

    The caller still runs the deterministic local script itself. This keeps one
    writer at a time and prevents an AI retry from duplicating ledger actions.
    """
    route = route_status(force=True)
    now = datetime.now()
    request_id = f"{now:%Y%m%d_%H%M%S}_{task_key}"
    record = {
        "id": request_id,
        "created_at": now.isoformat(timespec="seconds"),
        "task_key": task_key,
        "title": title,
        "local_command": [os.path.basename(os.environ.get("PYTHON", "python")), *args],
        "primary": route["primary"],
        "selected": route["selected"],
        "route_note": route["note"],
        "status": "queued",
        "extra": str(extra or ""),
        "safety": "AI 요청은 검토·실패 후속조치용이며, 실행 버튼의 업무 스크립트는 로컬에서 1회만 실행합니다.",
    }
    ticket_path = REPORT_DIR / f"{request_id}.json"
    _atomic_json(ticket_path, record)
    _atomic_json(STATUS_PATH, {**route, "last_request": record})
    return {**record, "_path": str(ticket_path)}


def _extra_block(record: dict[str, Any]) -> str:
    """표를 만든 쪽이 덧붙인 설명. 없으면 한 줄도 안 붙는다.

    ★ 실패한 명령이 아닌 표(예: 분담판에 세워 둔 일)에는 아래 1번 문장('로컬 업무
      스크립트는 이미 정확히 한 번 실행됐다')이 **사실이 아니다.** 그대로 보내면 AI 가
      없는 실행 로그를 찾아 헤맨다. 그래서 만든 쪽이 사실을 덧붙일 자리를 둔다 —
      프롬프트를 두 벌로 나누면 나중에 한쪽만 고쳐진다.
    """
    extra = str(record.get("extra") or "").strip()
    return ("\n덧붙임(이 표를 만든 쪽의 설명):\n" + extra + "\n") if extra else ""


def _ticket_prompt(record: dict[str, Any], local_returncode: int) -> str:
    return _extra_block(record) + f"""쿠팡 통합업무 자동화 프로젝트의 후속 검토 작업입니다.

작업명: {record.get('title', '')}
작업 키: {record.get('task_key', '')}
로컬 업무 스크립트 종료 코드: {local_returncode}

중요:
1. 로컬 업무 스크립트는 이미 정확히 한 번 실행됐으므로 다시 실행하지 마세요.
2. 먼저 AGENTS.md, session_handoff.py --check, 최신 19_AI작업인수인계 및 git 상태를 확인하세요.
3. 사용자의 현재 대화와 ai_claim 점유를 최우선으로 존중하세요. 다른 세션이 쓰는 파일은 건드리지 마세요.
4. 방금 실행 로그와 reports 결과를 검토하고, 실패·누락이 있으면 이 작업 범위 안에서만 안전하게 보완하세요.
5. ERP·밴드·카카오톡에 외부 메시지를 보내지 말고, 비밀키를 출력·커밋하지 마세요.
6. 관리대장은 openpyxl save 금지, 합성검증 및 프로젝트 규칙을 그대로 따르세요.
7. 완료 후 변경·검증 결과를 파일 기반 인수인계에 남기세요.
"""


# --------------------------------------------------------------------
# 보내기 전에 지시문 크기를 잰다 (2026-09-05 실사고).
#
# 실측: `claude -p` 를 cwd=ROOT(ecount)에서 부르면 CLAUDE.md 를 계층으로 읽어
#   Documents/CLAUDE.md(3,474) + PRJ/CLAUDE.md(1,107,757) + ecount/CLAUDE.md(1,107,757)
#   = 2,218,988 바이트를 싣는다. 그런데 뒤 둘은 **같은 파일의 사본**이다
#   (2026-07-31 규칙 - 세 파일은 내용이 같아야 한다).
#   그래서 요청이 1,079,788 토큰이 되어 한도(1,000,000)를 넘고 즉사한다.
#   대화 자체는 3,558 토큰뿐이다 - 곧 AI 인계가 구조적으로 100% 실패한다.
#   실측 2026-09: 실패 티켓 4건이 전부 이 사유다.
#
# * 막는 것이 아니라 헛시도를 없애는 것이다([67] - 자동 인계를 다시 조이지
#   않는다). 지시문이 줄면 이 관문은 저절로 안 걸린다.
# * 못 재면 보낸다([169]) - 물러나는 값은 거절 한 번이고, 잘못 막으면 AI
#   인계가 통째로 죽는다. 모를 때 기우는 방향은 보내는 쪽이다.
PROMPT_TOKEN_LIMIT = 1_000_000
_BYTES_PER_TOKEN = 2.2          # 실측 2,218,988바이트 <-> 1,079,788토큰(2.055)보다
                                # 크게 잡는다 = 토큰을 작게 어림 = 덜 막는다
_INSTRUCTION_NAME = {"claude": "CLAUDE.md", "codex": "AGENTS.md"}


def instruction_tokens(agent: str, cwd) -> int | None:
    """그 cwd 에서 CLI 가 계층으로 읽을 지시문 크기(토큰 어림). 못 재면 None."""
    name = _INSTRUCTION_NAME.get(agent)
    if not name:
        return None
    try:
        here = Path(cwd).resolve()
    except Exception:
        return None
    total = 0
    seen = False
    for d in [here, *here.parents]:
        try:
            p = d / name
            if p.is_file():
                total += p.stat().st_size
                seen = True
        except OSError:
            return None                 # 못 읽었으면 '작다'고 우기지 않는다
    return int(total / _BYTES_PER_TOKEN) if seen else None


def _too_long_trace(title: str, agent: str, est: int) -> None:
    """왜 못 보냈는지 자국을 남긴다. schedule_watch.traces() 가 reports/*_오류.json
    을 글로브로 모아 인계 '먼저 처리할 것' 에 저절로 싣는다([304]).
    """
    try:
        (ROOT / "reports").mkdir(exist_ok=True)
        _atomic_json(ROOT / "reports" / "AI인계_오류.json", {
            "작업": "AI 인계",
            "갈래": "프롬프트초과",
            "시각": datetime.now().isoformat(timespec="seconds"),
            "무엇": ("AI 인계를 못 보냈다 - %s 가 실을 지시문이 약 %d 토큰이라 "
                   "한도 %d 를 넘는다. 같은 지시문(CLAUDE.md 1.1MB)을 프로젝트 "
                   "뿌리와 ecount 에서 두 벌 싣기 때문이다. 표: %s"
                   % (agent, est, PROMPT_TOKEN_LIMIT, title[:60])),
            "조치": "python ecount/agent_dispatch.py --status",
        })
    except Exception:
        pass                            # 자국 하나로 인계를 죽이지 않는다


def _agent_command(agent: str, executable: str, prompt: str, last_message: Path,
                   chosen: dict[str, Any] | None = None) -> list[str]:
    """CLI 명령을 짓는다. `chosen` 이 있으면 **고른 모델**을 깃발로 붙인다.

    ★ 2026-08-12 전까지는 모델을 **한 번도 고르지 않았다** — `claude -p` 만 불러
      모든 인계가 기본 모델(제일 비싼 것)로 돌았다. 자율복구 재시도 한 장이나
      원인 모를 회차 고장이나 같은 값을 치렀다. 고르는 규칙은 `ai_tier` 한 곳이다.
    ★ 노력도 2026-08-13 부터 **깃발**로 간다(`--effort`). 다만 `ai_tier` 가 그 실행파일
      에게 `--help` 로 **직접 물어보고 있다고 확인된 때만** 붙인다 — 확인 안 된 깃발을
      붙이면 CLI 가 통째로 안 뜨고 인계가 조용히 안 된다. 못 물어봤으면 예전처럼
      프롬프트 문장이 그 몫을 한다(`ai_tier.prompt_line`).
    """
    if agent == "codex":
        return [executable, "exec", "-C", str(ROOT), "-s", "workspace-write",
                "--json", "-o", str(last_message), prompt]
    extra: list[str] = []
    if chosen:
        try:
            import ai_tier
            extra = ai_tier.flags(agent, chosen, executable)
        except Exception:                       # 고르기가 실패해도 인계는 나간다
            extra = []
    return [executable, *extra, "-p", prompt, "--output-format", "text"]


def _background_agent_env() -> dict[str, str]:
    """Environment contract for unattended Claude Code/Codex CLI work.

    No credentials are added here.  The markers only suppress colour/TUI behaviour
    and make it explicit to descendants that there is no interactive terminal.
    """
    env = os.environ.copy()
    env.update({
        "CI": "1",
        "NO_COLOR": "1",
        "TERM": "dumb",
        "CSOS_BACKGROUND_AGENT": "1",
    })
    return env


def run_ticket(ticket_path: str | Path, local_returncode: int = 0) -> dict[str, Any]:
    """한 티켓을 실제 AI CLI로 소비하고 상태를 원자적으로 남긴다."""
    path = Path(ticket_path)
    record = json.loads(path.read_text(encoding="utf-8"))
    if record.get("status") in ("running", "done"):
        return record
    route = route_status(force=True)
    agent = route.get("selected")
    if agent not in ("claude", "codex"):
        record.update({
            "status": "waiting",
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "error": route.get("note", "사용 가능한 AI CLI 없음"),
        })
        _atomic_json(path, record)
        return record
    executable = resolve_agent_executable(agent)
    if not executable:
        record.update({
            "status": "waiting",
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "error": f"{agent} 실제 실행 파일을 찾지 못함",
        })
        _atomic_json(path, record)
        return record

    last_message = path.with_suffix(".last.txt")
    log_path = path.with_suffix(".log")
    record.update({
        "selected": agent,
        "route_note": route.get("note", ""),
        "status": "running",
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "local_returncode": int(local_returncode),
        "log": log_path.name,
    })
    # ★ 모델·노력을 **여기서 고른다** (2026-08-12 지시 · `[230]`). 근거는 티켓이 이미
    #   갖고 있는 것뿐이다 — 갈래·제목·명령·시도 횟수. 지어내지 않는다.
    #   고른 것을 티켓에 적어 둔다: 나중에 "왜 이건 비싼 모델로 돌았나"를 물을 수
    #   있어야 규칙이 자란다.
    chosen: dict[str, Any] = {}
    prompt = _ticket_prompt(record, local_returncode)
    try:
        import ai_tier
        rc_i = int(local_returncode or 0)
        chosen = ai_tier.pick(
            kind=str(record.get("kind") or
                     ("timeout" if rc_i == 124 else "code" if rc_i else "")),
            title=str(record.get("title") or ""),
            args=list(record.get("args") or []),
            attempts=int(record.get("attempts") or (3 if rc_i else 0)))
        record["ai_tier"] = chosen
        prompt = prompt + "\n" + ai_tier.prompt_line(chosen)
    except Exception as exc:                    # 고르기가 실패해도 인계는 나간다
        record["ai_tier_error"] = str(exc)[:120]
    _atomic_json(path, record)
    # 보내기 전 관문: 지시문이 한도를 넘으면 안 보낸다(instruction_tokens 주석).
    _est = instruction_tokens(agent, ROOT)
    if _est is not None and _est > PROMPT_TOKEN_LIMIT:
        _why = ("AI 인계를 못 보냈다 - 지시문이 약 %d 토큰이라 한도 %d 를 넘는다"
                % (_est, PROMPT_TOKEN_LIMIT))
        record.update({"status": "failed", "error": _why, "not_sent": True,
                       "instruction_tokens": _est,
                       "completed_at": datetime.now().isoformat(timespec="seconds")})
        _atomic_json(path, record)
        _too_long_trace(str(record.get("title") or ""), agent, _est)
        return record

    command = _agent_command(agent, executable, prompt, last_message, chosen)
    try:
        result = run_tree(command, cwd=ROOT, env=_background_agent_env(),
                          timeout=AGENT_TIMEOUT_SECONDS,
                          drain_timeout=60, output_limit=200_000)
        combined = (result.stdout or "") + ("\n" + result.stderr if result.stderr else "")
        if result.timed_out:
            combined += ("\nAI 후속 검토 60분 시간 초과" +
                         (f" · 종료되지 않은 pid {result.stuck_pid}" if result.stuck_pid else ""))
        # `claude --version` can succeed even when the account has exhausted its
        # credits.  Detect that only after the real task starts and immediately
        # hand the same review ticket to Codex once.  The deterministic local
        # business script is not repeated; only the AI follow-up is retried.
        # ★ `returncode != 0` 만 보면 **로그인 안 된 CLI(exit 0)** 를 놓친다(위 실측).
        not_logged = _looks_not_logged_in(combined)
        if agent == "claude" and (result.returncode != 0 or not_logged) and (
                not_logged or _UNAVAILABLE_RE.search(combined)):
            codex_executable = resolve_agent_executable("codex")
            codex_blocked, codex_why = _credit_block("codex")
            if codex_executable and not codex_blocked:
                claude_output = combined
                agent = "codex"
                record.update({
                    "selected": "codex",
                    "fallback_from": "claude",
                    "route_note": "Claude Code 실제 실행 불가 → Codex 즉시 인계",
                    "fallback_at": datetime.now().isoformat(timespec="seconds"),
                })
                _atomic_json(path, record)
                command = _agent_command(
                    "codex", codex_executable,
                    _ticket_prompt(record, local_returncode), last_message,
                )
                result = run_tree(command, cwd=ROOT, env=_background_agent_env(),
                                  timeout=AGENT_TIMEOUT_SECONDS,
                                  drain_timeout=60, output_limit=200_000)
                codex_output = (result.stdout or "") + (
                    "\n" + result.stderr if result.stderr else ""
                )
                combined = (
                    "[Claude Code 실행 실패 — Codex로 자동 인계]\n"
                    + claude_output
                    + "\n\n[Codex 실행]\n"
                    + codex_output
                )
            elif codex_blocked:
                record.update({
                    "status": "waiting",
                    "updated_at": datetime.now().isoformat(timespec="seconds"),
                    "error": codex_why,
                    "route_note": "Claude Code 실제 실행 불가 · Codex 크래딧 대기",
                })
                log_path.write_text(combined[-200000:], encoding="utf-8")
                _atomic_json(path, record)
                _atomic_json(STATUS_PATH, {**route, "last_request": record})
                return record
        log_path.write_text(combined[-200000:], encoding="utf-8")
        record.update({
            # ★ exit 0 이어도 **로그인 안 됨**이면 성공이 아니다 — 아무 일도 안 했다.
            "status": "done" if (result.returncode == 0
                                 and not _looks_not_logged_in(combined)) else "failed",
            "agent_returncode": result.returncode,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
            "last_message": last_message.name if last_message.exists() else "",
            # 이유도 같은 판정을 쓴다 — 실패라 적고 이유를 비우면 아무도 못 고친다
            "error": "" if (result.returncode == 0
                            and not _looks_not_logged_in(combined)) else _clean_message(combined, 500),
        })
    except OSError as exc:
        record.update({
            "status": "failed",
            "completed_at": datetime.now().isoformat(timespec="seconds"),
            "error": _clean_message(str(exc), 500),
        })
    _atomic_json(path, record)
    _atomic_json(STATUS_PATH, {**route, "last_request": record})
    return record


def ai_paused() -> bool:
    """지금 AI 를 부르면 안 되나 — **크레딧 5시간 창이 막혔나**(2026-08-22 형님 지시).

    ★ **판정을 여기서 만들지 않는다**([162]) — `credit_window.blocked()` 를 빌린다.
      여기서 다시 세면 같은 순간에 두 답이 나온다.
    ★ **티켓을 만들기 전에** 물어야 한다. 만든 뒤에 막으면 그 티켓이 `queued` 로
      고아가 되고, 부르는 쪽들은 "표는 항목당 하나"라 **다시 안 만든다** — 충전이
      돼도 영영 안 도는 자리가 된다. 그래서 `_escalate`·`_hand_to_ai` 가 이 함수를
      **맨 앞에서** 부르고, 여기 `dispatch_async` 의 검사는 안전망이다.
    ★ **못 읽으면 안 막는다**([169] 의 방향을 자리에 맞게 정한 것) — 소진 중에 부르면
      실패 한 번이고 그것은 대기열이 되돌린다. 멀쩡한데 막으면 일이 안 된다.
    """
    try:
        import credit_window
        return credit_window.blocked()
    except Exception:
        return False


def dispatch_async(ticket: dict[str, Any], local_returncode: int = 0) -> bool:
    """앱 서버를 막지 않고 독립 워커가 티켓을 소비하게 한다."""
    # ★ 안전망 — 새 호출자가 생겨도 소진 중에는 워커를 안 띄운다. 티켓은 그대로
    #   남으므로 충전 뒤 다음 회차가 이어받는다.
    if ai_paused():
        return False
    if route_status(force=True).get("selected") not in ("claude", "codex"):
        return False
    ticket_path = ticket.get("_path")
    if not ticket_path:
        return False
    env = _background_agent_env()
    subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve()), "--run-ticket", str(ticket_path),
         "--local-returncode", str(int(local_returncode))],
        cwd=ROOT, env=env, stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        **background_popen_kwargs(),
    )
    return True


def resume_pending(limit: int = 1) -> list[str]:
    """크래딧·로그인 때문에 못 보낸 표를 다음 회차에서 다시 보낸다.

    새 표를 만들지 않고 기존 ``queued``/``waiting`` 표를 재사용한다. 그래서 크래딧이
    찬 동안 실패표가 쌓이지 않고, 충전 뒤 watchdog 회차가 가장 오래된 것부터 잇는다.
    """
    if route_status(force=True).get("selected") not in ("claude", "codex"):
        return []
    sent = []
    for path in sorted(REPORT_DIR.glob("*.json")):
        try:
            record = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError, json.JSONDecodeError):
            continue
        if record.get("status") not in ("queued", "waiting"):
            continue
        record["_path"] = str(path)
        if dispatch_async(record, int(record.get("local_returncode") or 0)):
            sent.append(str(record.get("id") or path.stem))
        if len(sent) >= max(1, int(limit)):
            break
    return sent


def supersede_queued(reason: str) -> int:
    """Close stale queued tickets after the same work was finished interactively."""
    count = 0
    for path in sorted(REPORT_DIR.glob("*.json")):
        try:
            record = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError, json.JSONDecodeError):
            continue
        if record.get("status") != "queued":
            continue
        record.update({
            "status": "superseded",
            "completed_at": datetime.now().isoformat(timespec="seconds"),
            "error": "",
            "superseded_reason": _clean_message(reason, 500),
        })
        _atomic_json(path, record)
        count += 1
    return count


def status() -> dict[str, Any]:
    """Latest availability and pending request summary for desktop/web status."""
    route = route_status()
    try:
        jobs = sorted(REPORT_DIR.glob("*.json"), reverse=True)
        latest = json.loads(jobs[0].read_text(encoding="utf-8")) if jobs else None
    except Exception:
        latest = None
    return {**route, "last_request": latest}


def route_label(info: dict[str, Any]) -> str:
    selected = (info or {}).get("selected", "codex_pending")
    if selected == "claude":
        return "Claude Code 우선"
    if selected == "codex":
        return "Claude Code 사용 불가 → Codex 폴백"
    return "Claude Code 사용 불가 → Codex 요청 대기"


def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--run-ticket")
    ap.add_argument("--local-returncode", type=int, default=0)
    ap.add_argument("--supersede-queued", metavar="REASON")
    ap.add_argument("--status", action="store_true")
    args = ap.parse_args(argv)
    if args.supersede_queued:
        count = supersede_queued(args.supersede_queued)
        print(json.dumps({"superseded": count}, ensure_ascii=False))
        return 0
    if args.run_ticket:
        result = run_ticket(args.run_ticket, args.local_returncode)
        print(json.dumps(result, ensure_ascii=False))
        return 0 if result.get("status") in ("done", "waiting") else 1
    print(json.dumps(status(), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
