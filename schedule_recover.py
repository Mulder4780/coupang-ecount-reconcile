# -*- coding: utf-8 -*-
"""놓친 일일 예약작업을 워치독 회차에서 한 건씩 다시 시작한다.

`schedule_watch` 는 읽기 전용 정본이다. 여기서는 그가 막 써 둔 판정 중
명시적으로 허용한 세 회차만 다룬다. 한 번에 하나만 시작해서 원본정리·일일대조가
서로 디스크와 Excel 을 놓고 싸우지 않게 하며, 작업 스케줄러의 ``IgnoreNew`` 가
두 번째 안전망이다.
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

from proc_guard import run_tree

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


ROOT = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(ROOT, "reports", "스케줄러_자동복구.json")
WATCH_STATE = os.path.join(ROOT, "reports", "스케줄러_회차감시.json")

# 원본을 먼저 모은 뒤 대조하고, 마지막에 값싼 후보검색을 한다.
RECOVER_ORDER = (
    "쿠팡업무_원본자료자동정리",
    "쿠팡업무_일일자동대조",
    "쿠팡업무_자동화후보검색",
)
RETRY_AFTER = timedelta(minutes=25)


def _read(path, default):
    try:
        with open(path, encoding="utf-8") as fh:
            value = json.load(fh)
        return value if isinstance(value, type(default)) else default
    except Exception:
        return default


def _write(path, value):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        json.dump(value, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _when(value):
    try:
        return datetime.fromisoformat(str(value or "")[:19])
    except Exception:
        return None


def pick(rows, state, now=None):
    """다시 시작할 한 행. 못 읽거나 이미 시도한 회차면 ``None``.

    완료 자국을 지어내지 않는다. 성공적으로 *시작 요청*을 보낸 뒤에도 다음
    ``schedule_watch`` 가 실제 LastRunTime/결과를 확인한다.
    """
    now = now or datetime.now()
    by_name = {str(r.get("작업") or ""): r for r in (rows or []) if isinstance(r, dict)}
    attempts = (state or {}).get("시도") or {}
    for name in RECOVER_ORDER:
        row = by_name.get(name)
        if not row or row.get("갈래") != "안돎" or row.get("상태") == "Running":
            continue
        due = str(row.get("예정") or "")
        due_at = _when(due)
        if not due_at or due_at > now:
            continue
        old = attempts.get(name) or {}
        if str(old.get("예정") or "") == due and old.get("요청성공"):
            # 시작 요청 성공은 완료가 아니다. schedule_watch가 이 행을 '안돎'에서
            # 내려 줄 때까지 뒤 회차를 시작하지 않는다.
            return None
        last_try = _when(old.get("시도시각"))
        if last_try and now - last_try < RETRY_AFTER:
            # 앞 회차 시작 자체가 실패했다면 뒤 회차가 낡은 원본으로 돌면 안 된다.
            return None
        return row
    return None


def _trigger(name):
    result = run_tree(
        ["schtasks.exe", "/Run", "/TN", name],
        cwd=ROOT, timeout=20, drain_timeout=5, output_limit=20_000,
    )
    ok = not result.timed_out and result.returncode == 0
    text = " ".join((result.stdout or "").split())[-300:]
    return ok, text or ("시간초과" if result.timed_out else "rc=%d" % result.returncode)


def _pipeline_busy():
    """A live/unknown automation owner blocks other heavy recovery jobs."""

    try:
        from automation_pipeline import pipeline_lock_status

        status = pipeline_lock_status(Path(ROOT) / "reports" / ".automation_pipeline.lock")
    except Exception:
        return None
    if not status.get("exists"):
        return False
    # Unknown ownership is not permission to start another disk-heavy pass.
    return status.get("alive") is not False


def run(dry=False, rows=None, now=None, trigger=None, pipeline_busy=None):
    now = now or datetime.now()
    report = _read(WATCH_STATE, {}) if rows is None else {"작업": rows}
    rows = report.get("작업") or []
    # 파일이 깨졌거나 옛 인코딩이면 키가 없다. 그것을 '할 일 없음'으로 확언하지 않는다.
    if rows and not any("작업" in r for r in rows if isinstance(r, dict)):
        return "놓친 회차 자동복구 확인 못 함 - 회차감시 파일 인코딩"
    state = _read(STATE, {})
    row = pick(rows, state, now)
    if row is None:
        return "놓친 일일회차 없음"
    busy = (pipeline_busy or _pipeline_busy)()
    if busy is not False:
        return "놓친 일일회차 대기 - 증분 자료 갱신이 끝난 뒤 한 건씩 시작"
    name, due = row["작업"], str(row.get("예정") or "")
    if dry:
        return "놓친 일일회차 예행 - %s" % name
    ok, detail = (trigger or _trigger)(name)
    state.setdefault("시도", {})[name] = {
        "예정": due,
        "시도시각": now.isoformat(timespec="seconds"),
        "요청성공": bool(ok),
        "결과": detail,
    }
    state["갱신"] = now.isoformat(timespec="seconds")
    _write(STATE, state)
    if ok:
        return "놓친 일일회차 시작 요청 - %s (실제 완료는 다음 감시가 확인)" % name
    return "놓친 일일회차 시작 실패 - %s · %s" % (name, detail[:100])


if __name__ == "__main__":
    print(run(dry="--dry" in os.sys.argv))
