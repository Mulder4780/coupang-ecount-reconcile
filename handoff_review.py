# -*- coding: utf-8 -*-
"""Terra -> Sol handoff review gate.

Terra writes a small, ignored marker once a meaningful work batch is committed.
Before Sol takes an exclusive write claim, it must review that batch from the
recorded base commit through the current HEAD.  The review is deliberately
file-backed so that a new session cannot lose the context held by another
model's chat window.

The report contains only file names and check results; it never stores a diff
or a possible credential value.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
from datetime import datetime
from typing import Any

from proc_guard import run_tree


BASE = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE, "reports")
MARKER = os.path.join(REPORT_DIR, "terra_handoff.json")
REVIEW_REPORT = os.path.join(REPORT_DIR, "sol_handoff_review.json")
IGNORED_DIRTY_PREFIXES = ("outputs/", "reports/")

# Only added, non-empty literal credential values are findings.  Names alone,
# empty values, environment-variable references, and documentation examples do
# not block the handoff.
SECRET_LITERAL = re.compile(
    r"(?ix)"
    r"(?:[\"']?)"
    r"(?:api[_-]?key|client[_-]?secret|access[_-]?token|refresh[_-]?token|"
    r"password|passwd)"
    r"(?:[\"']?)\s*[:=]\s*"
    r"[\"']([^\"'\r\n]{12,})[\"']"
)


def _git(*args: str, check: bool = False):
    result = run_tree(["git", *args], cwd=BASE, timeout=120, drain_timeout=10)
    if check and result.returncode:
        raise RuntimeError((result.stderr or result.stdout or "git 실패")[-1000:])
    return result


def _git_text(*args: str) -> str:
    try:
        return (_git(*args).stdout or "").strip()
    except Exception:
        return ""


def current_head() -> str:
    return _git_text("rev-parse", "HEAD")


def resolve_commit(value: str) -> str:
    out = _git_text("rev-parse", "--verify", f"{value}^{{commit}}")
    return out if re.fullmatch(r"[0-9a-f]{40}", out) else ""


def is_ancestor(base: str, head: str) -> bool:
    try:
        return _git("merge-base", "--is-ancestor", base, head).returncode == 0
    except Exception:
        return False


def _read_json(path: str) -> dict[str, Any] | None:
    try:
        with open(path, encoding="utf-8") as f:
            value = json.load(f)
        return value if isinstance(value, dict) else None
    except (OSError, ValueError, TypeError):
        return None


def _write_json_atomic(path: str, value: dict[str, Any]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".handoff-", suffix=".json", dir=os.path.dirname(path))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(value, f, ensure_ascii=False, indent=2)
            f.write("\n")
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def load_marker() -> dict[str, Any] | None:
    return _read_json(MARKER)


def load_review_report() -> dict[str, Any] | None:
    return _read_json(REVIEW_REPORT)


def sol_review_is_current(
    marker: dict[str, Any] | None = None,
    report: dict[str, Any] | None = None,
    head: str | None = None,
) -> bool:
    """True when no Terra handoff is pending, or Sol reviewed this exact state."""
    marker = load_marker() if marker is None else marker
    if not marker:
        return True
    report = load_review_report() if report is None else report
    head = current_head() if head is None else head
    return bool(
        report
        and report.get("passed") is True
        and report.get("base_commit") == marker.get("base_commit")
        and report.get("marker_head") == marker.get("head_commit")
        and report.get("reviewed_head") == head
    )


def review_state() -> dict[str, Any]:
    marker = load_marker()
    report = load_review_report()
    head = current_head()
    if not marker:
        return {"pending": False, "reason": "Terra handoff marker 없음", "head": head}
    current = sol_review_is_current(marker, report, head)
    return {
        "pending": not current,
        "reason": "검토 완료" if current else "Terra 변경분에 대한 Sol 검토 필요",
        "head": head,
        "base_commit": marker.get("base_commit"),
        "marker_head": marker.get("head_commit"),
    }


def mark_terra(base: str | None = None) -> dict[str, Any]:
    """Set or extend the review scope after Terra has committed its work."""
    head = current_head()
    if not head:
        raise RuntimeError("현재 Git HEAD를 확인할 수 없습니다.")
    marker = load_marker()
    report = load_review_report()
    if base:
        base_commit = resolve_commit(base)
    elif marker and not sol_review_is_current(marker, report, head):
        # Keep the original base while a handoff is still pending: no Terra
        # follow-up commit may escape the eventual Sol review.
        base_commit = str(marker.get("base_commit") or "")
    elif report and report.get("passed") and resolve_commit(str(report.get("reviewed_head") or "")):
        base_commit = str(report["reviewed_head"])
    else:
        base_commit = resolve_commit("HEAD~1") or head
    if not base_commit or not is_ancestor(base_commit, head):
        raise RuntimeError("검토 기준 커밋이 현재 HEAD의 조상이 아닙니다.")
    saved = {
        "schema_version": 1,
        "producer": "terra",
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "base_commit": base_commit,
        "head_commit": head,
    }
    _write_json_atomic(MARKER, saved)
    return saved


def blocking_dirty(lines: list[str]) -> list[str]:
    """Return dirty git paths that are meaningful source changes."""
    blocked: list[str] = []
    for line in lines:
        path = line[3:].strip() if len(line) >= 4 else line.strip()
        # Git's rename status is ``old -> new``; reviewing the new file is
        # sufficient for this gate.
        if " -> " in path:
            path = path.rsplit(" -> ", 1)[-1].strip()
        norm = path.replace("\\", "/")
        if norm and not norm.startswith(IGNORED_DIRTY_PREFIXES):
            blocked.append(norm)
    return blocked


def secret_findings(patch: str) -> list[str]:
    """Return affected files, never credential values, for added literals."""
    findings: list[str] = []
    current_file = "(unknown)"
    for line in patch.splitlines():
        if line.startswith("+++ b/"):
            current_file = line[6:]
        elif line.startswith("+") and not line.startswith("+++") and SECRET_LITERAL.search(line):
            if current_file not in findings:
                findings.append(current_file)
    return findings


def _changed_files(base: str, head: str) -> list[str]:
    return [p for p in _git_text("diff", "--name-only", f"{base}..{head}").splitlines() if p]


def _compile_python(files: list[str]) -> tuple[bool, str]:
    python_files = [os.path.join(BASE, p) for p in files if p.endswith(".py") and os.path.isfile(os.path.join(BASE, p))]
    if not python_files:
        return True, "변경된 Python 파일 없음"
    r = run_tree(
        [sys.executable, "-m", "py_compile", *python_files], cwd=BASE,
        timeout=120, drain_timeout=10,
    )
    return r.returncode == 0, "Python 문법 검사 통과" if r.returncode == 0 else "Python 문법 검사 실패"


def _synthetic_check() -> tuple[bool, str]:
    # Windows subprocess.run(timeout=...)은 SMB 대기에 걸린 자식의 communicate()에서
    # 다시 무기한 멈출 수 있다. 자식 트리와 출력 드레인까지 유한한 공용 실행기를 쓴다.
    # 합성 플래그와 보고서 경로도 자식 환경에서 강제해 호출자의 셸 설정에 좌우되지 않게 한다.
    with tempfile.TemporaryDirectory(prefix="csos-handoff-synthetic-") as sandbox:
        env = dict(os.environ)
        env["CSOS_SYNTHETIC"] = "1"
        env["COUPANG_REPORT_DIR"] = os.path.join(sandbox, "reports")
        env["COUPANG_UPDATES_DIR"] = os.path.join(sandbox, "updates")
        r = run_tree(
            [sys.executable, os.path.join(BASE, "tests", "synthetic_check.py")],
            cwd=BASE, env=env, timeout=600, drain_timeout=30,
        )
    output = (r.stdout or "") + ("\n" + r.stderr if r.stderr else "")
    if r.timed_out:
        tail = " (종료되지 않은 PID %s)" % r.stuck_pid if r.stuck_pid else ""
        return False, "합성 검증 시간초과%s" % tail
    return r.returncode == 0 and "ALL GREEN" in output, "합성 검증 ALL GREEN" if r.returncode == 0 and "ALL GREEN" in output else "합성 검증 실패"


def review_sol() -> tuple[bool, dict[str, Any] | None]:
    marker = load_marker()
    if not marker:
        print("대기 중인 Terra 인수인계가 없습니다. Sol 검토 관문 통과.")
        return True, None

    head = current_head()
    base = str(marker.get("base_commit") or "")
    errors: list[str] = []
    if not resolve_commit(base) or not head or not is_ancestor(base, head):
        errors.append("Terra 검토 기준 커밋을 현재 HEAD에서 확인할 수 없습니다.")
        files: list[str] = []
        patch = ""
    else:
        files = _changed_files(base, head)
        patch = _git_text("diff", "--no-ext-diff", f"{base}..{head}")
        whitespace = _git("diff", "--check", f"{base}..{head}")
        if whitespace.returncode != 0:
            errors.append("커밋 변경분에 공백 오류가 있습니다.")

    dirty = blocking_dirty([p for p in _git_text("status", "--porcelain").splitlines() if p])
    if dirty:
        errors.append("미커밋 변경이 있습니다: " + ", ".join(dirty[:10]))
    secret_files = secret_findings(patch)
    if secret_files:
        errors.append("추가된 변경분에서 비밀값 형태가 감지되었습니다: " + ", ".join(secret_files))

    compile_ok, compile_summary = _compile_python(files) if not errors else (False, "선행 오류로 Python 문법 검사를 건너뜀")
    if not compile_ok:
        errors.append(compile_summary)
    synthetic_ok, synthetic_summary = _synthetic_check() if not errors else (False, "선행 오류로 합성 검증을 건너뜀")
    if not synthetic_ok:
        errors.append(synthetic_summary)

    report = {
        "schema_version": 1,
        "reviewer": "sol",
        "reviewed_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "passed": not errors,
        "base_commit": base,
        "marker_head": marker.get("head_commit"),
        "reviewed_head": head,
        "changed_files": files,
        "changed_file_count": len(files),
        "dirty_paths": dirty,
        "secret_finding_files": secret_files,
        "python_compile": compile_summary,
        "synthetic_check": synthetic_summary,
        "errors": errors,
    }
    _write_json_atomic(REVIEW_REPORT, report)
    if errors:
        print("Sol 인수인계 검토 실패: " + " / ".join(errors))
        return False, report
    print("Sol 인수인계 검토 PASS: 변경 %d개 파일, %s, %s" % (len(files), compile_summary, synthetic_summary))
    return True, report


def main() -> int:
    ap = argparse.ArgumentParser(description="Terra 작업분을 Sol이 검토하는 파일 기반 관문")
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--mark-terra", action="store_true", help="Terra 완료 커밋을 Sol 검토 대상으로 표시")
    group.add_argument("--review-sol", action="store_true", help="Terra 변경분을 검토하고 합성 검증")
    group.add_argument("--status", action="store_true", help="검토 관문 상태만 읽기 전용으로 확인")
    ap.add_argument("--base", help="Terra 검토 범위의 시작 커밋(기본: 직전 검토/HEAD~1)")
    args = ap.parse_args()
    try:
        if args.mark_terra:
            marker = mark_terra(args.base)
            print("Terra 인수인계 표시 완료: %s..%s" % (marker["base_commit"][:12], marker["head_commit"][:12]))
            return 0
        if args.review_sol:
            return 0 if review_sol()[0] else 2
        state = review_state()
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 2 if state["pending"] else 0
    except Exception as exc:
        print("인수인계 검토 관문 오류: %s" % exc)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
