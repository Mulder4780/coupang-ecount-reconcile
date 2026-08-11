# -*- coding: utf-8 -*-
"""PC 가동 중 5분 주기 읽기 전용 문제 감시기.

원칙:
  * 08:00~09:30 KST에는 파일을 열거나 결과 파일을 갱신하지 않는다.
  * Excel 저장 잠금이 없고 2분 이상 바뀌지 않은 원장만 임시 사본으로 검사한다.
  * 관리대장·ERP·외부 메시지는 절대 수정하지 않는다.
  * 같은 이슈는 한 건으로 유지하고 신규/지속/해결 전이만 이력에 남긴다.
"""
from __future__ import annotations

import argparse
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from operation_window import KST, input_window_label, is_input_window, korea_now

ROOT = Path(__file__).resolve().parent
REPORT = ROOT / "reports" / "realtime_issues.json"
REPORT_MD = ROOT / "reports" / "realtime_issues.md"
HISTORY = ROOT / "reports" / "realtime_issues_history.jsonl"
CONFIG = ROOT / "config" / "ecount_config.json"
SETTLE_SECONDS = 120
VERSION_RE = re.compile(r"_v(\d+)\.xlsx$", re.IGNORECASE)
SEVERITY_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}


def _atomic_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + f".{os.getpid()}.tmp")
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


def _atomic_text(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + f".{os.getpid()}.tmp")
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        f.write(value)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


def _load_json(path: Path, default: Any = None) -> Any:
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _issue(issue_id: str, severity: str, title: str, evidence: str, action: str) -> dict[str, Any]:
    return {
        "id": issue_id,
        "severity": severity,
        "title": title,
        "evidence": evidence,
        "action": action,
    }


def _master_folder() -> Path:
    cfg = _load_json(CONFIG)
    if not isinstance(cfg, dict):
        raise RuntimeError("config/ecount_config.json을 읽을 수 없습니다")
    master = ((cfg.get("reconcile") or {}).get("master_xlsx"))
    if not master:
        raise RuntimeError("reconcile.master_xlsx 설정이 없습니다")
    return Path(master).expanduser().resolve().parent


def version_files(folder: Path) -> list[tuple[int, Path]]:
    found: list[tuple[int, Path]] = []
    for path in folder.glob("*_v*.xlsx"):
        if path.name.startswith("~$"):
            continue
        match = VERSION_RE.search(path.name)
        if match:
            found.append((int(match.group(1)), path))
    return sorted(found)


def fingerprint(path: Path) -> dict[str, Any]:
    stat = path.stat()
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(block)
    return {
        "path": str(path),
        "name": path.name,
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
        "sha256": digest.hexdigest(),
    }


HAND_EDIT_LOG = ROOT / "reports" / "엑셀_손입력_감지.json"


def hand_edit_verdict(prev: Any, cur: Any, machine_source: str) -> tuple[bool, str]:
    """직전 지문 대 현재 지문 — **기계 회차 근거 없이 바뀐 최신본은 손입력이다.**

    2026-08-11 지시(엑셀 손입력 종료)의 감지 반쪽. 역수입 금지라 그 값은 정본에
    안 들어가는데, 말없이 버리면 그 사람의 입력이 소리 없이 사라진다 — 그래서
    바뀐 것을 **알린다**(자동 반영은 하지 않는다).

    판정 불가(직전 지문 없음·해시 없음)는 경보가 아니다 — 모르면 함부로 말하지
    않는다([172]의 문). 순수 함수로 둔 이유는 검증 [212]가 워크북 없이 시험하기
    위해서다.
    """
    if not isinstance(prev, dict) or not isinstance(cur, dict):
        return False, ""
    if not prev.get("sha256") or not cur.get("sha256"):
        return False, ""
    if prev["sha256"] == cur["sha256"]:
        return False, ""
    if machine_source:
        return False, ""
    pv, cv = prev.get("version"), cur.get("version")
    if pv == cv and prev.get("name") == cur.get("name"):
        return True, f"같은 파일({cur.get('name')})의 내용이 기계 회차 없이 바뀜"
    return True, f"새 버전(v{pv}→v{cv})이 기계 회차 근거 없이 생김"


def _archive_change_source(mtime: float, window_min: int = 40) -> str:
    """보관본 생성기(archive_worker)가 그 시각 언저리에 내보냈는가 — batch 표와
    별개의 기계 경로라 따로 본다(경합을 손입력으로 오판하면 경보가 죽는다)."""
    try:
        st = _load_json(ROOT / "tmp" / "archive_spool" / "worker-status.json", {})
        if not isinstance(st, dict):
            return ""
        for stamp in (st.get("finished_at"),
                      (st.get("export") or {}).get("finished_at") if isinstance(st.get("export"), dict) else None):
            if not stamp:
                continue
            try:
                t = datetime.fromisoformat(str(stamp)).timestamp()
            except ValueError:
                continue
            if -60 <= (mtime - t) <= window_min * 60:
                return "보관본 생성(archive_worker)"
    except Exception:
        return ""
    return ""


def _note_hand_edit(entry: dict[str, Any]) -> None:
    """감지 기록 — session_handoff 가 해시 없이 읽는 싼 신호([168])."""
    if os.environ.get("CSOS_SYNTHETIC") == "1":
        return                      # 합성검증은 실기록을 오염시키지 않는다
    try:
        prev = _load_json(HAND_EDIT_LOG, [])
        if not isinstance(prev, list):
            prev = []
        prev.append(entry)
        _atomic_json(HAND_EDIT_LOG, prev[-100:])
    except Exception:
        pass


def workbook_ready(path: Path, now: datetime, settle_seconds: int = SETTLE_SECONDS) -> tuple[bool, str]:
    owner_files = list(path.parent.glob("~$*.xlsx"))
    if owner_files:
        return False, "Excel 임시 잠금 파일이 있음: " + ", ".join(p.name for p in owner_files[:3])
    try:
        age = now.timestamp() - path.stat().st_mtime
    except FileNotFoundError:
        return False, "원장 파일이 사라짐"
    if age < settle_seconds:
        return False, f"마지막 저장 후 {int(age)}초(안정화 기준 {settle_seconds}초)"
    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
        if bad:
            return False, f"ZIP 손상 항목: {bad}"
    except (OSError, zipfile.BadZipFile) as exc:
        return False, f"XLSX ZIP 검사 실패: {exc}"
    return True, "저장 안정화·ZIP 검사 통과"


def _snapshot(path: Path) -> tuple[Path | None, str]:
    before = path.stat()
    handle = tempfile.NamedTemporaryFile(prefix="realtime_audit_", suffix=".xlsx", delete=False)
    snapshot = Path(handle.name)
    handle.close()
    try:
        shutil.copy2(path, snapshot)
        after = path.stat()
        if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
            snapshot.unlink(missing_ok=True)
            return None, "사본 생성 중 원장이 변경됨"
        with zipfile.ZipFile(snapshot) as zf:
            if zf.testzip():
                snapshot.unlink(missing_ok=True)
                return None, "임시 사본 ZIP 검증 실패"
        return snapshot, "임시 사본 검증 통과"
    except Exception:
        snapshot.unlink(missing_ok=True)
        raise


def _run_audit(script: str, snapshot: Path, timeout: int = 180) -> tuple[int, str]:
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    proc = subprocess.run(
        [sys.executable, str(ROOT / script), "--file", str(snapshot)],
        cwd=ROOT,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )
    return proc.returncode, ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()


def audit_workbook(path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    snapshot, note = _snapshot(path)
    if snapshot is None:
        return [
            _issue("workbook_changed_during_snapshot", "P2", "원장 저장 안정화 대기",
                   note, "다음 5분 주기에 다시 검사")
        ], {"ok": False, "note": note}
    details: dict[str, Any] = {"ok": True, "note": note}
    try:
        rc1, out1 = _run_audit("fix_workbook.py", snapshot)
        structure_ok = rc1 == 0 and "구조 검사" in out1 and "이상 없음" in out1
        details["structure"] = {"returncode": rc1, "ok": structure_ok, "tail": out1[-1200:]}
        if not structure_ok:
            issues.append(_issue(
                "workbook_structure_audit", "P1", "관리대장 구조 검사 이상",
                out1[-500:] or f"fix_workbook 종료코드 {rc1}",
                "자동 수정하지 말고 원본·직전 정상본을 비교한 뒤 ledger 잠금 하에 복구",
            ))

        rc2, out2 = _run_audit("fix_formulas.py", snapshot)
        counts = [int(x) for x in re.findall(
            r"(?:깨진 수식 복구|누적번호 수식 정렬|범위 확장|빠진 수식 채움|"
            r"정적 상태 수식화|기사명 정규화)\s*(\d+)", out2
        )]
        formulas_ok = rc2 == 0 and len(counts) >= 6 and sum(counts) == 0
        details["formulas"] = {"returncode": rc2, "ok": formulas_ok, "counts": counts, "tail": out2[-1200:]}
        if not formulas_ok:
            issues.append(_issue(
                "workbook_formula_audit", "P1", "관리대장 수식 검사 이상",
                out2[-500:] or f"fix_formulas 종료코드 {rc2}",
                "자동 반영하지 말고 문제 셀과 수식 소유 열을 확인한 뒤 별도 보완",
            ))

        rc3, out3 = _run_audit("dashboard_clean.py", snapshot)
        dashboard_count = re.search(r"합계\s+(\d+)칸", out3)
        dashboard_ok = rc3 == 0 and dashboard_count is not None and int(dashboard_count.group(1)) == 0
        details["dashboard"] = {
            "returncode": rc3,
            "ok": dashboard_ok,
            "count": int(dashboard_count.group(1)) if dashboard_count else None,
            "tail": out3[-1200:],
        }
        if not dashboard_ok:
            issues.append(_issue(
                "dashboard_filldown_debris", "P1", "대시보드 채우기 내림 잔해",
                out3[-500:] or f"dashboard_clean 종료코드 {rc3}",
                "자동 반영하지 말고 문제 셀을 확인한 뒤 ledger 잠금 하에 전용 ZIP 패치로 정리",
            ))
        details["ok"] = not issues
    except subprocess.TimeoutExpired as exc:
        issues.append(_issue(
            "workbook_audit_timeout", "P1", "관리대장 검사 시간 초과",
            f"{Path(str(exc.cmd[1])).name} {exc.timeout}초 초과",
            "원장 크기·손상·Excel 프로세스 상태를 확인",
        ))
        details["ok"] = False
    finally:
        snapshot.unlink(missing_ok=True)
    return issues, details


def _queue_checks() -> list[dict[str, Any]]:
    path = ROOT / "updates" / "pending_updates.json"
    try:
        with open(path, encoding="utf-8") as f:
            queue = json.load(f)
    except Exception as exc:
        return [_issue(
            "pending_queue_invalid", "P1", "자동입력 대기열 JSON 손상",
            f"{path.name}: {type(exc).__name__}: {exc}",
            "파일을 빈 대기열로 간주하지 말고 백업과 JSON 구조를 먼저 확인",
        )]
    if not isinstance(queue, list):
        return [_issue(
            "pending_queue_schema", "P1", "자동입력 대기열 형식 이상",
            f"목록이어야 하나 {type(queue).__name__}",
            "pending_updates.json 스키마를 복구",
        )]
    if queue and (korea_now().timestamp() - path.stat().st_mtime) > 3600:
        return [_issue(
            "pending_queue_stale", "P2", "자동입력 대기 장기 미처리",
            f"{len(queue)}건, 마지막 변경 {datetime.fromtimestamp(path.stat().st_mtime):%m-%d %H:%M}",
            "근거와 대상 열을 검토한 뒤 승인된 원장 반영 절차 실행",
        )]
    return []


def _daily_checks(master: Path, now: datetime) -> list[dict[str, Any]]:
    path = ROOT / "reports" / "agent_status.json"
    data = _load_json(path)
    if not isinstance(data, dict):
        if now.time().replace(tzinfo=None) >= datetime.strptime("10:10", "%H:%M").time():
            return [_issue(
                "daily_status_missing", "P1", "일일 자동대조 상태 없음",
                "reports/agent_status.json을 읽을 수 없음",
                "합성검증 후 daily_run을 수동 점검",
            )]
        return []
    failed = [s for s in data.get("steps", []) if isinstance(s, dict) and s.get("ok") is False]
    if failed:
        names = ", ".join(str(s.get("name", "?")) for s in failed[:5])
        return [_issue(
            "daily_step_failed", "P1", "일일 자동대조 실패 단계",
            names, "해당 단계 로그를 확인하고 게시·원장 반영 여부를 검증",
        )]
    try:
        status_mtime = path.stat().st_mtime
        master_mtime = master.stat().st_mtime
        if master_mtime > status_mtime + 60:
            # ★ 이 알림은 **매일 반드시** 뜬다 — 11:00·15:00 반영이 오전 대조보다 나중이라
            #   원장이 항상 더 새 것이 된다(2026-08-05). 늘 뜨는 알림은 아무도 안 본다.
            #   그래서 "우리가 넣은 것"과 "누가 밖에서 만진 것"을 갈라서 말한다.
            who = _master_change_source(master_mtime)
            if who:
                return [_issue(
                    "daily_status_older_than_master", "P3",
                    "원장이 일일검증 뒤 반영 회차로 갱신됨(정상)",
                    f"일일상태 {datetime.fromtimestamp(status_mtime):%H:%M} → 원장 "
                    f"{datetime.fromtimestamp(master_mtime):%H:%M} · {who}",
                    "조치 불필요 — 다음 일일 대조가 교차검증한다",
                )]
            return [_issue(
                "daily_status_older_than_master", "P2",
                "원장이 **반영 회차 밖에서** 변경됨",
                f"일일상태 {datetime.fromtimestamp(status_mtime):%H:%M}, 원장 "
                f"{datetime.fromtimestamp(master_mtime):%H:%M} — 11:00·15:00 회차와 맞지 않음",
                "누가 열어 고쳤는지 확인하고, 다음 일일 대조에서 ERP·카톡·밴드 교차검증",
            )]
    except OSError:
        pass
    return []


def _master_change_source(mtime: float, window_min: int = 40) -> str:
    """원장이 바뀐 그 시각이 **우리 반영 회차** 때문인가. 맞으면 회차 이름을 돌려준다.

    엑셀 재계산까지 끝나는 데 시간이 걸려 파일 시각은 회차 시작보다 뒤로 밀린다 —
    그래서 넉넉한 창(기본 40분)으로 본다. 근거는 `ledger_db.batch` 표다(우리가 남긴 기록).
    """
    try:
        import sqlite3
        db = ROOT / "db" / "ledger_queue.db"
        if not db.exists():
            return ""
        con = sqlite3.connect(str(db))
        try:
            rows = con.execute(
                "select slot, started, finished from batch where ok=1 "
                "order by rowid desc limit 8").fetchall()
        finally:
            con.close()
        for slot, started, finished in rows:
            for stamp in (finished, started):
                if not stamp:
                    continue
                try:
                    t = datetime.fromisoformat(str(stamp)).timestamp()
                except ValueError:
                    continue
                if -60 <= (mtime - t) <= window_min * 60:
                    return f"{slot} 회차"
    except Exception:
        return ""
    return ""


def _claim_checks() -> list[dict[str, Any]]:
    path = ROOT / "reports" / "ai_claims.json"
    if not path.exists():
        return []
    try:
        data = json.load(open(path, encoding="utf-8"))
    except Exception as exc:
        return [_issue(
            "ai_claims_invalid", "P1", "AI 협업 점유 파일 손상",
            f"{type(exc).__name__}: {exc}", "원장 쓰기를 중단하고 점유 파일을 복구",
        )]
    if not isinstance(data, dict):
        return [_issue(
            "ai_claims_schema", "P1", "AI 협업 점유 형식 이상",
            type(data).__name__, "ai_claims.json 스키마를 확인",
        )]
    return []


def _kakao_checks() -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    roots = [
        ROOT / "kakao" / "inbox",
        Path.home() / "Desktop",
        Path.home() / "OneDrive" / "Desktop",
    ]
    files: list[Path] = []
    for root in roots:
        if root.is_dir():
            if root == ROOT / "kakao" / "inbox":
                files.extend(p for p in root.glob("*.txt") if p.name != ".gitkeep")
            else:
                files.extend(root.glob("KakaoTalk_*_group.txt"))
    unique = sorted({p.resolve() for p in files if p.is_file()}, key=lambda p: p.stat().st_mtime)
    sources = [
        {"path": str(p), "size": p.stat().st_size,
         "modified": datetime.fromtimestamp(p.stat().st_mtime, KST).isoformat(timespec="seconds")}
        for p in unique[-10:]
    ]
    outside = [p for p in unique if ROOT / "kakao" / "inbox" not in p.parents]
    inbox = [p for p in unique if ROOT / "kakao" / "inbox" in p.parents]
    issues: list[dict[str, Any]] = []
    if outside and (not inbox or outside[-1].stat().st_mtime > inbox[-1].stat().st_mtime + 1):
        issues.append(_issue(
            "kakao_export_not_ingested", "P2", "새 카카오톡 내보내기 미반영",
            f"{outside[-1].name} ({datetime.fromtimestamp(outside[-1].stat().st_mtime):%m-%d %H:%M})",
            "원본은 읽기 전용으로 두고 kakao_reconcile --file로 대조한 뒤 inbox 보관 여부 확인",
        ))
    return issues, sources


def reconcile_issues(current: list[dict[str, Any]], previous: dict[str, Any], now: datetime) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    prev = {
        x.get("id"): x for x in previous.get("issues", [])
        if isinstance(x, dict) and x.get("id") and x.get("status") != "resolved"
    }
    current_map = {x["id"]: x for x in current}
    result: list[dict[str, Any]] = []
    transitions: list[dict[str, Any]] = []
    stamp = now.isoformat(timespec="seconds")

    for issue_id, item in current_map.items():
        old = prev.get(issue_id)
        confirmations = int(old.get("confirmations", 0)) + 1 if old else 1
        if old:
            status = "new" if old.get("status") == "provisional" else "ongoing"
            first_seen = old.get("first_seen", stamp)
        else:
            status = "new" if item["severity"] in ("P0", "P1") else "provisional"
            first_seen = stamp
        merged = {
            **item,
            "status": status,
            "confirmations": confirmations,
            "first_seen": first_seen,
            "last_seen": stamp,
        }
        result.append(merged)
        if status == "new":
            transitions.append({**merged, "transition": "new", "at": stamp})

    for issue_id, old in prev.items():
        if issue_id in current_map or old.get("status") == "provisional":
            continue
        resolved = {**old, "status": "resolved", "last_seen": stamp}
        result.append(resolved)
        transitions.append({**resolved, "transition": "resolved", "at": stamp})

    result.sort(key=lambda x: (SEVERITY_ORDER.get(x.get("severity"), 9), x.get("id", "")))
    return result, transitions


def _render_md(report: dict[str, Any]) -> str:
    lines = [
        "# 실시간 문제 감시",
        "",
        f"- 점검 시각: {report['checked_at']}",
        f"- 감시 범위: PC 로그인·가동 중 5분 주기, 입력 제외 {input_window_label()}",
        f"- 원장: {report.get('source', {}).get('name', '확인 불가')}",
        f"- 활성 이슈: {report['summary']['active']}건 (신규 {report['summary']['new']} / 지속 {report['summary']['ongoing']} / 확인중 {report['summary']['provisional']})",
        "",
    ]
    active = [x for x in report["issues"] if x["status"] != "resolved"]
    if not active:
        lines += ["현재 확인된 문제가 없습니다.", ""]
    else:
        for item in active:
            lines += [
                f"## [{item['severity']}] {item['title']} — {item['status']}",
                "",
                f"- 근거: {item['evidence']}",
                f"- 조치: {item['action']}",
                "",
            ]
    resolved = [x for x in report["issues"] if x["status"] == "resolved"]
    if resolved:
        lines += ["## 이번 점검에서 해결됨", ""]
        lines += [f"- [{x['severity']}] {x['title']}" for x in resolved]
        lines.append("")
    return "\n".join(lines)


def run_once(now: datetime | None = None, settle_seconds: int = SETTLE_SECONDS) -> dict[str, Any] | None:
    now = now or korea_now()
    if is_input_window(now):
        print(f"입력 보호시간({input_window_label()}) — 파일 접근 없이 종료")
        return None

    previous = _load_json(REPORT, {})
    if not isinstance(previous, dict):
        previous = {}
    active: list[dict[str, Any]] = []
    audit: dict[str, Any] = {"ok": False, "note": "원장 미확인"}
    source: dict[str, Any] = {}
    master_deferred = False

    try:
        folder = _master_folder()
        versions = version_files(folder)
        if not versions:
            active.append(_issue(
                "master_missing", "P0", "관리대장 최신본 없음",
                str(folder), "관리대장 경로와 버전 파일을 즉시 확인",
            ))
        else:
            latest_version, master = versions[-1]
            source = fingerprint(master)
            source["version"] = latest_version
            # ★ 손입력 감지(2026-08-11 지시 — 앱 전용 입력): 직전 지문(source)은
            #   리포트에 이미 저장돼 있었는데 아무도 비교하지 않았다. 기계 회차
            #   (batch 표·보관본 생성) 근거 없이 바뀐 최신본은 사람 손이다.
            master_mtime = master.stat().st_mtime
            machine = (_master_change_source(master_mtime)
                       or _archive_change_source(master_mtime))
            hand, hand_why = hand_edit_verdict(previous.get("source"), source, machine)
            if hand:
                active.append(_issue(
                    "master_hand_edit", "P1", "관리대장 손입력 감지 — 앱 전용 입력 위반",
                    hand_why + " · 손으로 적은 값은 정본(DB)에 들어가지 않음",
                    "적은 사람을 찾아 앱으로 다시 입력하도록 안내 (자동 반영 금지 — 역수입 금지)",
                ))
                _note_hand_edit({
                    "시각": now.isoformat(timespec="seconds"),
                    "종류": "내용변경",
                    "파일": source.get("name"),
                    "이전sha": str((previous.get("source") or {}).get("sha256") or "")[:12],
                    "현재sha": str(source.get("sha256") or "")[:12],
                    "근거": hand_why,
                })
            forks = [
                p for version, p in versions[:-1]
                if p.stat().st_mtime > master.stat().st_mtime + 60
            ]
            if forks:
                active.append(_issue(
                    "master_fork_detected", "P1", "낮은 버전 원장이 최신본 뒤에 수정됨",
                    ", ".join(p.name for p in forks[-5:]),
                    "자동 아카이브하지 말고 파일별 변경 내용을 비교해 정본 결정",
                ))
            ready, why = workbook_ready(master, now, settle_seconds)
            if not ready:
                master_deferred = True
                active.append(_issue(
                    "master_not_stable", "P2", "관리대장 저장 안정화 대기",
                    why, "Excel 입력을 방해하지 않고 다음 주기에 재검사",
                ))
                audit = {"ok": False, "note": why}
            else:
                workbook_issues, audit = audit_workbook(master)
                active.extend(workbook_issues)
                active.extend(_daily_checks(master, now))
    except Exception as exc:
        active.append(_issue(
            "monitor_master_error", "P1", "관리대장 감시 오류",
            f"{type(exc).__name__}: {exc}", "설정·경로·권한을 확인",
        ))

    active.extend(_queue_checks())
    active.extend(_claim_checks())
    kakao_issues, kakao_sources = _kakao_checks()
    active.extend(kakao_issues)

    # 저장 중에는 구조·수식·일일상태를 재검사하지 못했을 뿐 해결된 것이 아니다.
    # 직전 활성 이슈를 유지해 "저장 대기 → 거짓 해결 → 재발" 알림 진동을 막는다.
    if master_deferred:
        current_ids = {x["id"] for x in active}
        carry_prefixes = ("workbook_", "daily_", "master_fork_")
        for old in previous.get("issues", []):
            if (
                isinstance(old, dict)
                and old.get("status") not in ("resolved", None)
                and str(old.get("id", "")).startswith(carry_prefixes)
                and old.get("id") not in current_ids
            ):
                active.append({
                    key: old[key] for key in ("id", "severity", "title", "evidence", "action")
                    if key in old
                })

    issues, transitions = reconcile_issues(active, previous, now)
    summary = {
        "active": sum(x["status"] != "resolved" for x in issues),
        "new": sum(x["status"] == "new" for x in issues),
        "ongoing": sum(x["status"] == "ongoing" for x in issues),
        "provisional": sum(x["status"] == "provisional" for x in issues),
        "resolved": sum(x["status"] == "resolved" for x in issues),
    }
    report = {
        "schema": 1,
        "checked_at": now.isoformat(timespec="seconds"),
        "quiet_window": input_window_label(),
        "mode": "read-only",
        "source": source,
        "audit": audit,
        "kakao_sources": kakao_sources,
        "summary": summary,
        "issues": issues,
    }
    _atomic_json(REPORT, report)
    _atomic_text(REPORT_MD, _render_md(report))
    if transitions:
        HISTORY.parent.mkdir(parents=True, exist_ok=True)
        with open(HISTORY, "a", encoding="utf-8", newline="\n") as f:
            for item in transitions:
                f.write(json.dumps(item, ensure_ascii=False, separators=(",", ":")) + "\n")
            f.flush()
            os.fsync(f.fileno())
    print(f"실시간 감시 완료: 활성 {summary['active']}건 · 신규 {summary['new']}건 · 해결 {summary['resolved']}건")
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--settle-seconds", type=int, default=SETTLE_SECONDS)
    args = parser.parse_args()
    run_once(settle_seconds=max(0, args.settle_seconds))


if __name__ == "__main__":
    main()
