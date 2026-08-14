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
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
REPORTS = ROOT / "reports"
OUT_JSON = REPORTS / "시스템_업무진단.json"
OUT_MD = REPORTS / "시스템_업무진단.md"
VERSION = 1


def _now() -> datetime:
    return datetime.now().astimezone()


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
            "python webapp/server_guard.py --once", "reports/server_guard_status.json")
    elif guard_age is None or guard_age > 5:
        add("server-guard-stale", "P0", "앱 서버 보호자 심박이 끊김",
            f"마지막 보호 기록이 {int(guard_age or 0)}분 전입니다(정상 한도 5분).",
            "작업 스케줄러의 CSOS_ServerGuard와 server_guard.log를 확인합니다.",
            "reports/server_guard_status.json")
    elif str(guard.get("state") or "").lower() != "healthy":
        add("server-guard-failed", "P0", "앱 서버 보호자가 장애를 보고함",
            str(guard.get("message") or guard.get("state") or "상태 설명 없음"),
            "python webapp/server_guard.py --once", "reports/server_guard_status.json")

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
        waiting = "지금 내릴까요" in watchdog_tail or "(y = 내림" in watchdog_tail
        add("watchdog-stale", "P0",
            "워치독이 입력 대기에서 멈춤" if waiting else "워치독 30분 회차가 멈춤",
            f"마지막 로그가 {int(watchdog_age)}분 전"
            + ("이며 끝부분에 사람 답변을 기다리는 문구가 있습니다." if waiting else "입니다."),
            "무인 회차에서는 확인창을 띄우지 말고 재시작을 보류 기록으로 남기도록 고칩니다.",
            "reports/watchdog_log.txt")

    # 3) 일일 대조 — 실패를 '오늘 실행됨'으로 세지 않는다.
    daily_path = REPORTS / ".daily_run.progress.json"
    daily = _read_json(daily_path)
    daily_age = _age_minutes(daily_path)
    sources["daily_run"] = {"age_minutes": daily_age, "read": daily is not None}
    if daily is None:
        add("daily-run-unreadable", "P1", "일일 대조 진행 자국을 읽지 못함",
            "마지막 완주·실패 단계를 판정할 근거가 없습니다.",
            "python daily_run.py --status", "reports/.daily_run.progress.json")
    else:
        state = str(daily.get("상태") or "")
        if state == "실패":
            add("daily-run-failed", "P0", "오늘 일일 대조가 중단됨",
                "%s · 실패 원인: %s" % (daily.get("시각") or "시각 없음",
                                        daily.get("오류") or daily.get("오류유형") or "원인 없음"),
                "실패한 합성검증을 고친 뒤 일일 대조를 다시 실행합니다.",
                "reports/.daily_run.progress.json")
        elif daily_age is not None and daily_age > 20 * 60:
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
            sample = " · ".join(str(x)[:150] for x in alerts[:3])
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
            "Claude Code 설치·로그인을 복구한 뒤 python agent_dispatch.py --route --force",
            "reports/agent_dispatch_status.json")

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
        add("erp-api-failed", "P1", "ERP 공식 API 수집 실패",
            str(erp_api.get("error") or "원인 설명 없음")[:220],
            "python erp_api_collect.py --force", "reports/erp_api_latest.json")

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
