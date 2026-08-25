# -*- coding: utf-8 -*-
"""쿠팡 업무 자율복구 제어기.

결정론적 업무 스크립트가 먼저 일하고, 공용 자원(Z:·관리대장) 장애는 실패 도미노로
만들지 않고 영속 대기열에 둔다. 워치독이 30분마다 안전한 항목만 재개하며 같은 코드·
시간초과가 세 번 반복될 때만 Claude Code→Codex 검토 큐로 넘긴다.

세금계산서 실발행·외부 메시지·로그인은 자동 실행하지 않는다. 인증 세션이 생기면
그 뒤 단계는 자동 재개하지만, 비밀번호나 법적 발행 승인을 대신하는 것은 자동화가
아니라 권한 탈취이기 때문이다.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from proc_guard import run_tree


ROOT = Path(__file__).resolve().parent
REPORTS = ROOT / "reports"
QUEUE_PATH = REPORTS / "autopilot_queue.json"
STATUS_PATH = REPORTS / "autopilot_status.json"
REPORT_PATH = REPORTS / "자율자동화_상태.md"
PY = sys.executable
MAX_ATTEMPTS_BEFORE_AI = 3
BASE_BACKOFF_MINUTES = 30
# 멱등 작업이 한 회차 몫을 정상 저장한 뒤 "아직 남음"을 알리는 반환값.
# 실패 횟수로 세지 않고 waiting으로 유지한다.
INCREMENTAL_RETURN_CODE = 75

# ★ **그 자식이 예산을 읽으면 바깥 제한보다 짧게 줘서 스스로 멈추게 한다**
#   (2026-08-25 실사고 · `source_tidy_run.CHILD_BUDGET_ENV` 와 같은 모양 · [324]).
#   예산이 없으면 `run_tree` 가 제한시간에 나무를 끊는데(SIGKILL), 그러면 자식의
#   stdout 버퍼가 통째로 사라져 **자국이 `returncode=-9` 다섯 글자뿐**이 된다.
#   실측: `밴드 게시글 보관` 27회 시도가 전부 그 상태였고, 그래서 일이 되고
#   있는데도 *"10회 넘게 재시도해도 안 풀린다"* 라는 가짜 경보가 매일 섰다([170]).
# ★ **표에 없는 자식은 안 건드린다**([324] · [169] 없는 손잡이를 지어내지 않는다).
#   그래서 `collect_all.py` 를 거쳐 도는 같은 스크립트는 한 톨도 안 바뀐다 —
#   그쪽은 이미 제 예산(7분)으로 스스로 멈춘다([172]).
CHILD_BUDGET_ENV = {
    "archive_posts.py": "ARCHIVE_POSTS_BUDGET_SEC",
    "stmt_archive.py": "STMT_ARCHIVE_BUDGET_SEC",
    # ⚠ 여기 이름을 더하려면 **그 스크립트가 그 열쇠를 실제로 읽어야** 한다.
    #    안 읽는 자식에게 넣어 봐야 아무 일도 안 일어난다([169] 없는 손잡이).
    # ⚠ 실측 2026-08-25 로 **일부러 안 넣은 것 둘** — 재 보고 안 맞아서다([172]):
    #    · `tax_archive.py` — `collect()` 가 PDF 한 장 만들기 **전에 193.7초**를 쓴다
    #      (Z: 재귀 glob + xlsx 80개를 `read_only=False` 로 연다). 예산이 60초면
    #      루프에 닿기도 전에 다 되어 **0건 만들고 75 로 돌아오는 것을 영원히**
    #      되풀이한다([199] 무한루프와 같은 모양). 그 194초를 먼저 줄여야 한다.
    #    · `zscan.py` — **스캔**이라 중간에 끊으면 리포트가 반쪽인데 겉모습은 완전하다
    #      ("서류 PDF N개" 의 N 만 줄어든다). 여기 예산을 주면 **조용히 틀린 리포트**를
    #      매 회차 만든다([169]) — 예산이 아니라 다른 고침이 필요하다.
}
CHILD_BUDGET_MARGIN_S = 300      # 보고서를 쓰고 돌아올 여유


def _child_env(args, timeout):
    """표에 있는 자식에게만 예산을 얹은 env 를 준다 — 아니면 `None`(그대로 물려줌)."""
    key = None
    for a in args:
        key = CHILD_BUDGET_ENV.get(os.path.basename(str(a)))
        if key:
            break
    if not key:
        return None
    env = dict(os.environ)
    env[key] = str(max(60, int(timeout) - CHILD_BUDGET_MARGIN_S))
    return env


def _configure_text_output() -> None:
    """윈도우 CP949 콘솔과 pythonw 모두에서 상태 출력을 안전하게 만든다."""
    for stream in (sys.stdout, sys.stderr):
        if stream is None or not hasattr(stream, "reconfigure"):
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            # 닫힌 파이프처럼 출력 통로 자체가 없는 경우 상태 판정은 계속한다.
            pass

RESOURCE_MARKERS = (
    "관리대장을 찾을 수 없음",
    "관리대장 v*.xlsx 를 찾을 수 없습니다",
    "네트워크 경로를 찾지 못했습니다",
    "폴더에 닿지 못했습니다",
    "정기점검 원본 폴더가 없습니다",
    "winerror 53",
    "z:\\",
    "z:/",
)
#: 수집 문(collect_gate.guard)이 거절한 표식. 이것은 고장이 아니라 **남의 차선 일**이다.
#: 낱말이 바뀌면 여기서 조용히 0건이 되므로 안내문과 같은 글자를 쓴다([165]).
LANE_MARKERS = ("수집 문이 막았습니다",)
AUTH_MARKERS = (
    "로그인이 필요", "인증 없음", "not authenticated", "login required",
    "밴드 미인증", "ecount 로그인",
)
IRREVERSIBLE_MARKERS = (
    "실전송", "세금계산서 발행", "외부 메시지", "카카오톡 전송", "밴드 게시",
)
NON_RETRY_FLAGS = {"--apply", "--queue", "--send", "--post", "--upload", "--delete"}


def _now() -> datetime:
    return datetime.now().astimezone()


def _atomic_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix=path.stem + "_", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as fh:
            json.dump(value, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def _load_queue() -> dict[str, Any]:
    try:
        doc = json.loads(QUEUE_PATH.read_text(encoding="utf-8"))
        if isinstance(doc, dict) and isinstance(doc.get("items"), list):
            return doc
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        pass
    return {"version": 1, "items": []}


def command_key(name: str, args: Iterable[str]) -> str:
    raw = json.dumps([name, *map(str, args)], ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:20]


def retry_safe(args: Iterable[str]) -> bool:
    """읽기·멱등 도구만 자동 재실행한다. 쓰기 가능 플래그는 사람이 아닌 정규 회차 몫."""
    values = {str(x).lower() for x in args}
    return not bool(values & NON_RETRY_FLAGS)


def classify_failure(output: str) -> str:
    text = " ".join(str(output or "").lower().split())
    if "증분 수집 계속 필요" in text:
        return "incremental"
    # ★ 수집 문 거절을 `code`(코드가 깨졌다)로 읽으면 조치가 사람을 **멀쩡한
    #   코드로 보낸다**([172]·[289]). 게다가 세 번 반복되면 `ai_tier` 가 '원인'
    #   등급(opus/high)을 매겨 **이미 아는 원인**에 비싼 모델을 부른다(실측).
    if any(x.lower() in text for x in LANE_MARKERS):
        return "lane"
    if any(x.lower() in text for x in AUTH_MARKERS):
        return "auth"
    if any(x.lower() in text for x in RESOURCE_MARKERS):
        return "resource"
    if "시간초과" in text or "timed out" in text or "timeout" in text:
        return "timeout"
    return "code"


def defer(name: str, args: list[str], timeout: int, output: str) -> dict[str, Any] | None:
    """안전한 실패를 중복 없이 대기열에 넣는다. 위험 작업은 절대 우회 재실행하지 않는다."""
    kind = classify_failure(output)
    # 합성검증은 회차의 안전문이다. 이것을 '나중에'로 돌리고 업무를 계속하면 안 된다.
    if name == "합성검증" or not retry_safe(args):
        return None
    now = _now()
    doc = _load_queue()
    key = command_key(name, args)
    item = next((x for x in doc["items"] if x.get("key") == key), None)
    if item is None:
        item = {
            "key": key,
            "name": name,
            "args": list(args),
            "timeout": int(timeout),
            "created_at": now.isoformat(timespec="seconds"),
            "attempts": 0,
            "status": "waiting",
            "ai_ticket": "",
        }
        doc["items"].append(item)
    item.update({
        "kind": kind,
        "status": "blocked" if kind == "auth" else "waiting",
        "last_error": " ".join(str(output or "").split())[-1200:],
        "updated_at": now.isoformat(timespec="seconds"),
        "next_attempt": (now + timedelta(minutes=BASE_BACKOFF_MINUTES)).isoformat(timespec="seconds"),
    })
    _atomic_json(QUEUE_PATH, doc)
    write_status(doc)
    return item


def resolve(name: str, args: list[str]) -> None:
    doc = _load_queue()
    key = command_key(name, args)
    changed = False
    for item in doc["items"]:
        if item.get("key") == key and item.get("status") != "done":
            item.update({"status": "done", "resolved_at": _now().isoformat(timespec="seconds")})
            changed = True
    if changed:
        _atomic_json(QUEUE_PATH, doc)
        write_status(doc)


def _due(item: dict[str, Any], now: datetime) -> bool:
    if item.get("status") not in ("waiting", "retry", "blocked"):
        return False
    try:
        return datetime.fromisoformat(str(item.get("next_attempt") or "")) <= now
    except (TypeError, ValueError):
        return True


def _escalate(item: dict[str, Any]) -> str:
    """같은 안전 작업이 세 번 실패한 경우에만 AI 한 장을 만든다(중복 금지)."""
    if (item.get("ai_ticket") or item.get("kind") in ("resource", "auth", "lane") or
            int(item.get("attempts") or 0) < MAX_ATTEMPTS_BEFORE_AI):
        return ""
    # ★ 크레딧 5시간 창이 막혔으면 **표를 만들지 않는다**(2026-08-22 형님 지시).
    #   여기서 만들어 두면 `ai_ticket` 이 채워져 **다시는 안 만든다** — 충전이 돼도
    #   영영 안 도는 자리가 된다. 안 만들면 다음 워치독 회차가 그대로 이어받는다.
    try:
        from agent_dispatch import ai_paused
        if ai_paused():
            return ""
    except Exception:
        pass
    try:
        from agent_dispatch import dispatch_async, enqueue
        ticket = enqueue(
            "autopilot-" + str(item["key"]),
            "자율복구 반복 실패: " + str(item.get("name") or ""),
            list(item.get("args") or []),
        )
        # 큐 파일이 만들어진 순간 인계는 내구성을 얻었다. 워커 기동이 실패해도 다음
        # 점검이 같은 티켓을 소비하게 두고, 새 티켓을 계속 만들지는 않는다.
        item["ai_ticket"] = str(ticket.get("id") or "queued")
        # 1을 고정으로 쓰면 실제 timeout도 "스크립트 종료 코드 1"로 인계돼 코드
        # 오류처럼 보인다. 워치독이 관측한 종료 사유를 그대로 넘긴다.
        local_rc = 124 if item.get("kind") == "timeout" else int(
            item.get("last_returncode") or 1)
        dispatch_async(ticket, local_returncode=local_rc)
        return item["ai_ticket"]
    except Exception as exc:
        item["ai_error"] = f"{type(exc).__name__}: {str(exc)[:180]}"
    return ""


def heal(*, limit: int = 2, budget_seconds: int = 600, dry: bool = False) -> dict[str, Any]:
    """워치독 회차에서 만기 항목을 조금씩 재개한다. 한 항목이 회차를 독점하지 않는다."""
    doc = _load_queue()
    now = _now()
    actions: list[dict[str, Any]] = []
    spent = 0
    for item in doc["items"]:
        if len(actions) >= max(0, int(limit)) or spent >= max(1, int(budget_seconds)):
            break
        if not _due(item, now):
            continue
        args = list(item.get("args") or [])
        if not retry_safe(args):
            item.update({"status": "manual", "last_error": "자동 재실행 금지 플래그 포함"})
            continue
        if dry:
            actions.append({"name": item.get("name"), "result": "dry"})
            continue
        timeout = min(int(item.get("timeout") or 600), max(30, budget_seconds - spent))
        started = _now()
        result = run_tree([PY, *args], cwd=ROOT, timeout=timeout, drain_timeout=30,
                          env=_child_env(args, timeout))
        spent += max(1, int((_now() - started).total_seconds()))
        combined = (result.stdout or "") + ("\n" + result.stderr if result.stderr else "")
        item["last_attempt"] = _now().isoformat(timespec="seconds")
        item["last_returncode"] = 124 if result.timed_out else int(result.returncode)
        if result.returncode == 0 and not result.timed_out:
            item.update({"status": "done", "resolved_at": item["last_attempt"], "last_error": ""})
            outcome = "done"
        elif result.returncode == INCREMENTAL_RETURN_CODE and not result.timed_out:
            # 정상적인 증분 진척은 실패도 완료도 아니다. 이전 실패 연속 횟수와 처리된
            # AI 티켓을 비워, 훗날의 실제 3회 연속 실패만 새로 인계되게 한다.
            item.pop("resolved_at", None)
            item.update({
                "status": "waiting",
                "attempts": 0,
                "continuations": int(item.get("continuations") or 0) + 1,
                "kind": "incremental",
                "ai_ticket": "",
                "last_error": "한 회차 몫을 저장했고 남은 안전 작업은 다음 회차가 이어서 한다",
                "next_attempt": (_now() + timedelta(minutes=BASE_BACKOFF_MINUTES)).isoformat(
                    timespec="seconds"),
            })
            outcome = "waiting"
        else:
            item["attempts"] = int(item.get("attempts") or 0) + 1
            kind = "timeout" if result.timed_out else classify_failure(combined)
            delay = min(12 * 60, BASE_BACKOFF_MINUTES * (2 ** min(item["attempts"], 4)))
            item.update({
                "status": "blocked" if kind == "auth" else "retry",
                "kind": kind,
                "last_error": (combined or f"returncode={result.returncode}")[-1200:],
                "next_attempt": (_now() + timedelta(minutes=delay)).isoformat(timespec="seconds"),
            })
            _escalate(item)
            outcome = item["status"]
        actions.append({"name": item.get("name"), "result": outcome})
    if dry:
        # --dry는 판단만 한다. 상태 파일의 시각까지 바꾸면 '실행했다'는 거짓 자국이다.
        return {**summary(doc), "actions": actions}
    doc["updated_at"] = _now().isoformat(timespec="seconds")
    _atomic_json(QUEUE_PATH, doc)
    status_value = write_status(doc, actions=actions)
    return {**status_value, "actions": actions}


def summary(doc: dict[str, Any] | None = None) -> dict[str, Any]:
    doc = doc or _load_queue()
    items = list(doc.get("items") or [])
    counts = {key: sum(1 for x in items if x.get("status") == key)
              for key in ("waiting", "retry", "blocked", "manual", "done")}
    active = [x for x in items if x.get("status") not in ("done", "superseded")]
    return {
        "time": _now().isoformat(timespec="seconds"),
        "mode": "deterministic-first-ai-on-exception",
        "active": len(active),
        "counts": counts,
        "human_gates": [
            "밴드·이카운트 최초 로그인/세션 만료 복구",
            "세금계산서 실발행·외부 메시지·취소 확정 같은 비가역 승인",
            "은행·PO 등 원천 자료가 아직 제공되지 않은 건",
        ],
        "items": active[:30],
    }


def write_status(doc: dict[str, Any] | None = None, *, actions: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    value = summary(doc)
    value["last_actions"] = actions or []
    _atomic_json(STATUS_PATH, value)
    REPORTS.mkdir(parents=True, exist_ok=True)
    lines = [
        "# 자율자동화 상태", "",
        f"- 갱신: {value['time']}",
        f"- 방식: 결정론적 자동처리 → 제한 재시도 → 코드·시간초과 3회 반복만 AI 자동 인계",
        f"- 자동복구 대기: {value['active']}건",
        "",
        "## 자동화가 넘지 않는 경계", "",
    ]
    lines.extend(f"- {x}" for x in value["human_gates"])
    lines += ["", "## 대기 항목", ""]
    if not value["items"]:
        lines.append("- 없음")
    else:
        for item in value["items"]:
            lines.append("- **%s** · %s · 시도 %s회 · %s" % (
                item.get("name"), item.get("status"), item.get("attempts", 0),
                item.get("kind", "")))
    REPORT_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return value


#: 출력에서 **오류처럼 보이는 줄**을 고르는 표시.
#: `실패` 는 넣지 않는다 — 정상 출력에 흔하다(`… 미분류 0건 · 실패 0건`).
#: 넓히면 멀쩡한 줄을 원인으로 지목하고, 그러면 사람이 엉뚱한 데를 고치러 간다([172]).
# ★ `시간 초과` 는 **공백이 있는 모양도 실재한다** — 실측으로 `collect_all.py` 와
#   `agent_dispatch.py` 가 그렇게 적고, `daily_run`·`erp_grab` 등은 붙여 적는다.
#   한 모양만 알면 다른 쪽 실패는 **표시를 못 찾아 꼬리로 떨어진다**(실측 '미수집 원본'
#   이 openpyxl 경고를 앞세웠다) — 낱말이 어긋나면 한 건도 안 걸리면서 오류도 안 난다(`[165]`).
_ERR_MARK = re.compile(
    r"(Error|Exception|Traceback|HTTP\s*\d{3}|returncode=|시간 ?초과|거부|Not Found)")


def _why_line(text: str) -> str:
    """실패 출력에서 **왜인지**를 뽑는다 — 못 찾으면 꼬리를 준다([169]).

    ★ 앞을 자르면 안 된다([365]) — 실측 2026-08-23 '고정 주소 사본 올리기' 는
      앞 160자가 전부 "사본 만드는 중… 관리대장 최신본 자동 탐지" 였다.
    ★ 그렇다고 꼬리만 실어도 안 된다 — 그 건은 **끝이 openpyxl 경고**라
      정작 원인인 `HTTP 404` 가 가운데 묻힌다. 그래서 **오류 줄을 먼저 세운다**.
    ★ 여러 개면 **마지막 것**이다(진짜 원인은 대개 마지막 오류).
    """
    lines = [" ".join(l.split()) for l in (text or "").splitlines()]
    lines = [l for l in lines if l]
    for l in reversed(lines):
        if _ERR_MARK.search(l):
            return l[:200]
    tail = " ".join(" ".join(lines).split())
    if len(tail) > 160:
        tail = "…(앞 %d자 줄임) " % (len(tail) - 160) + tail[-160:]
    return tail


#: 자원 실패 문구에서 **경로**를 뽑는 두 모양(실측 2026-08-25).
#:   · `[WinError 53] … : 'Z:\\'`               ← 작은따옴표 안(파이썬 repr)
#:   · `관리대장을 찾을 수 없음: Z:/… .xlsx (…)`  ← 확장자로 끝난다
#: ⚠ **오류 줄에서만** 뽑는다 — 트레이스백 프레임(File "C:/…/x.py")까지 세면
#:   후보가 넷이 되어 '유일할 때만' 문이 아무것도 안 걸린다(실측).
#: ⚠ 큰따옴표로 감싼 경로는 일부러 안 잡는다 — 못 잡으면 **모름**이라 경보가
#:   그대로 유지된다(틀려도 안전한 쪽으로만 틀린다).
_RES_PATH_RE = (
    re.compile(r"'([A-Za-z]:[\\/][^'\n]*)'"),
    re.compile(r"([A-Za-z]:[\\/][^\n']*?\.(?:xlsx|xlsm|xls|json|csv|txt|db|bundle))"),
)


def resource_back(text):
    """자원 실패가 가리키던 경로가 **지금** 살아 있나 — 모르면 None([169]).

    ★ 왜 필요한가 (2026-08-25 실측): `고정 주소 사본 올리기` 가 **31회 실패**로
      인계 맨 위에 서 있었는데, 그 마지막 시도(14:59:18)가 죽은 뒤
      **같은 일이 15:05 에 다른 길로 성공**했다(커밋 af683a5 · docs/data.enc 가
      13분마다 갱신된다). 오류는 둘 다 `Z:` 가 그 순간 끊긴 것이었다
      (`[WinError 53]` · `폴더에 v*.xlsx 없음`). 곧 **코드가 깨진 것이 아니라
      바쁜 시각에 공유폴더를 못 잡은 것**이고, 한가할 때는 저절로 된다.
      그런데 판정이 `attempts` 만 보아 "AI 인계까지 실패했다" 로 올렸다 —
      그 조치는 코드를 뒤지는 것이라 사람을 **멀쩡한 코드로 보낸다**([172]·[289]).
      그리고 가짜가 맨 위를 차지하면 진짜 경보가 묻힌다([170]).

    ★ **지어낼 것이 없다** — 경로는 오류 문구에 그대로 적혀 있다.
    ★ **후보가 유일할 때만** 답한다(이 저장소가 여러 곳에서 쓰는 그 문).
    ★ **틀려도 안전한 쪽으로만 틀린다** — 경로를 못 뽑거나 이스케이프가 어긋나면
      `isdir` 이 False 라 경보가 그대로 남는다.
    ★ **'고쳐졌다'고 말하지 않는다**([322]) — 말할 수 있는 것은
      "자원이 지금은 살아 있다" 까지이고, 답은 다음 재시도가 낸다.
    """
    lines = [" ".join(l.split()) for l in (text or "").splitlines()]
    errs = [l for l in lines if l and _ERR_MARK.search(l)]
    if not errs:
        return None
    cands = []
    for rx in _RES_PATH_RE:
        for m in rx.finditer(errs[-1]):
            p = m.group(1).strip()
            if p and p not in cands:
                cands.append(p)
    if len(cands) != 1:                # 없거나 여럿이면 **모름**이다([169])
        return None
    path = cands[0]
    probe = path if os.path.splitext(path)[1] == "" else os.path.dirname(path)
    if not probe:
        return None
    try:
        return bool(os.path.isdir(probe))
    except Exception:
        return None


#: 재시도를 이만큼 했는데도 안 풀리면 **사람이 봐야 한다**.
#: 3회 반복이면 이미 AI 에게 넘긴다([190]) — 그 갑절을 넘겼다는 것은
#: **AI 인계까지 실패했다**는 뜻이다. 낮추면 정상적으로 여러 회차 걸리는 일
#: (보관·대량 수집)이 매일 경보가 되어 아무도 안 본다([170]).
STUCK_TRIES = 10


def stuck(doc: dict[str, Any] | None = None) -> dict[str, Any]:
    """자율복구가 **오래 못 푸는 일**을 돌려준다 — 못 읽으면 그렇게 말한다([169]).

    ★ 왜 필요한가 (2026-08-23 실측): `reports/자율자동화_상태.md` 는 다섯 건을
      시도 횟수까지 정확히 적고 있었는데 **그 파일을 읽는 코드가 한 곳도 없었다**
      (`session_handoff`·`system_audit` 둘 다 0건 · [328]). 그래서 폰이 PC 없이
      보는 클라우드 사본이 **9일째 한 번도 안 올라갔는데** 인계 어디에도 안 떴다.
      `reports/cloud_continuity.json` 에는 `ok:false · HTTP 404` 가 그대로 있었다.

    ★ **경보 기준은 '대기 건수'가 아니라 '오래 굳었나'** 다. 대기 자체는 정상이고
      (양이 많아 여러 회차 걸리는 일이 있다) 매일 뜨면 아무도 안 본다([170]).

    ★ **왜인지는 지어내지 않는다**([169]·[289]) — 마지막 오류를 **그대로** 싣는다.
      갈래마다 조치가 다르므로(`code` 는 코드·설정, `timeout` 은 양, `auth` 는 인증)
      한 문장으로 뭉치면 사람이 엉뚱한 데를 고치러 간다([172]).
    """
    try:
        doc = doc or _load_queue()
        items = list(doc.get("items") or [])
    except Exception as exc:                      # 못 읽었다 ≠ 걸린 것 없다([169])
        return {"굳음": [], "자원회복": [], "못읽음": "%s: %s" % (type(exc).__name__, exc)}
    out, back = [], []
    for x in items:
        if x.get("status") not in ("retry", "blocked", "manual"):
            continue
        try:
            tries = int(x.get("attempts") or 0)
        except Exception:
            tries = 0
        if tries < STUCK_TRIES:
            continue
        # ★ **뒤에서** 싣는다([365]) — 진짜 원인은 출력의 **끝**에 있다.
        #   실측 2026-08-23: '고정 주소 사본 올리기' 는 앞 160자가 전부
        #   "사본 만드는 중… 관리대장 최신본 자동 탐지" 라 정작 원인인
        #   `HTTP 404` 가 안 실렸다 — 겉은 경보인데 왜인지는 못 읽는 자리다([169]).
        raw = str(x.get("last_error") or "")
        갈래 = str(x.get("kind") or "")
        rec = {"이름": str(x.get("name") or ""), "시도": tries,
               "갈래": 갈래, "왜": _why_line(raw)}
        # ★ 자원 실패는 **지나간 사고**일 수 있다 — 그 자원이 지금 살아 있으면
        #   경보가 아니라 알림이다(다음 재시도가 저절로 푼다). 조용히 빼지는
        #   않는다([169]) — `자원회복` 으로 세어 인계가 한 줄 적는다.
        if 갈래 == "resource" and resource_back(raw) is True:
            rec["다음시도"] = str(x.get("next_attempt") or "")
            back.append(rec)
            continue
        out.append(rec)
    out.sort(key=lambda r: -r["시도"])
    back.sort(key=lambda r: -r["시도"])
    return {"굳음": out, "자원회복": back, "못읽음": ""}


def status() -> dict[str, Any]:
    try:
        d = json.loads(STATUS_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError, json.JSONDecodeError):
        # 상태 조회/API GET은 읽기 전용이다. 첫 회차 전에는 메모리 요약만 돌려준다.
        d = summary()
    # ★ 사람이 읽는 자리에 **왜인지 한 줄**을 같이 준다.
    #   인계 '먼저 처리할 것' 이 조치로 `python autopilot.py --status` 를 주는데,
    #   `last_error` 는 1,200자 **꼬리**라 실측 2026-08-23 '고정 주소 사본 올리기'(26회)가
    #   맨 끝의 **openpyxl 경고**를 보여 줬다 — 진짜 원인(`HTTP 404`)은 가운데 묻혀 있었다.
    #   그러면 사람이 멀쩡한 데를 고치러 간다(`[172]`·`[289]` — 조치는 갈래마다 다르다).
    #   판정을 새로 만들지 않고 `_why_line` 을 빌린다(`[162]`), **원문은 그대로 둔다**
    #   (기계가 읽는다 · `[169]`) — 옆에 `왜` 를 더할 뿐이다.
    for _it in (d.get("items") or []):
        if isinstance(_it, dict) and _it.get("last_error"):
            _it["왜"] = _why_line(str(_it.get("last_error") or ""))
    return d


def selftest() -> None:
    assert classify_failure("FileNotFoundError: 관리대장을 찾을 수 없음: Z:/x") == "resource"
    assert classify_failure("로그인이 필요합니다") == "auth"
    assert classify_failure("시간초과(600s)") == "timeout"
    assert classify_failure("수집 문이 막았습니다 — 이 창은 수집 창이 아닙니다") == "lane"
    assert classify_failure("AssertionError") == "code"
    assert retry_safe(["read_only.py"])
    assert not retry_safe(["ledger_db.py", "--apply"])
    assert command_key("a", ["b"]) == command_key("a", ["b"])
    print("autopilot self-test: OK")


def main(argv: list[str] | None = None) -> int:
    _configure_text_output()
    ap = argparse.ArgumentParser()
    ap.add_argument("--heal", action="store_true")
    ap.add_argument("--dry", action="store_true")
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--limit", type=int, default=2)
    ap.add_argument("--budget", type=int, default=600)
    ns = ap.parse_args(argv)
    if ns.selftest:
        selftest()
        return 0
    value = heal(limit=ns.limit, budget_seconds=ns.budget, dry=ns.dry) if ns.heal else status()
    print(json.dumps(value, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
