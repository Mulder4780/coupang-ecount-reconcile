# -*- coding: utf-8 -*-
"""overnight_run.py — 밤새 스스로 다시 시도하는 전체 대조 (2026-08-06 지시)

사용자 지시: "밤을 새서라도 미처리된 건들 전부 대조해서 찾아서 해결하고 완료처리해".

무엇이 문제였나
  `daily_run.py` 는 맨 앞에서 합성검증을 돌리고, 하나라도 빨간 줄이 있으면 **전체를
  중단**한다(절대규칙: 실데이터 작업 전 합성검증 생략 금지). 옳은 설계다. 그런데
  이 프로젝트는 창을 여러 개 띄워 **동시에** 작업한다. 옆 세션이 화면을 고치는 중이면
  그 몇 분 동안 검증이 빨갛고, 그 순간에 걸린 대조는 통째로 날아간다. 밤 11시에 한 번
  걸어 두고 자면 아침에 "합성검증 실패 — 전체 중단" 한 줄만 남아 있는 것이다.

이 스크립트가 하는 일
  검증이 초록이 될 때까지 **기다렸다가** 대조를 돌린다. 검증을 건너뛰지 않는다 —
  건너뛰는 대신 **다시 시도**한다. 성공하면 그 회차로 끝낸다.

  · 검증이 빨가면 `--every` 분 뒤 다시 본다(기본 20분). 옆 세션이 커밋을 끝내면 초록이 된다.
  · `--hours` 안에 끝내 초록이 안 되면 그 사실을 기록으로 남기고 조용히 끝낸다.
  · 결과는 `reports/밤샘대조_기록.json` 에 매 시도마다 쌓인다.
  · 엑셀은 열지 않는다 — 반영은 11:00·15:00 회차 그대로다(daily_run 이 그 규칙을 지킨다).

  python ecount/overnight_run.py                  # 20분 간격 · 최대 12시간
  python ecount/overnight_run.py --every 10 --hours 8
"""
import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
LOG = os.path.join(ROOT, "reports", "밤샘대조_기록.json")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _run(args, timeout):
    try:
        p = subprocess.run([PY] + args, cwd=ROOT, timeout=timeout,
                           capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        return p.returncode, (p.stdout or "")[-4000:]
    except subprocess.TimeoutExpired:
        return 124, "시간 초과"
    except Exception as e:
        return 1, str(e)[:400]


def _note(rec):
    try:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        old = []
        if os.path.exists(LOG):
            try:
                old = json.load(open(LOG, encoding="utf-8"))
            except Exception:
                old = []
        old.append(rec)
        json.dump(old[-60:], open(LOG, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    except Exception:
        pass


def green():
    """합성검증이 초록인가. 건너뛰지 않는다 — 이것이 실데이터 작업의 관문이다."""
    rc, out = _run([os.path.join("tests", "synthetic_check.py")], 900)
    ok = (rc == 0) and ("ALL GREEN" in out)
    bad = ""
    if not ok:
        for line in out.splitlines():
            if "AssertionError" in line or "Error" in line:
                bad = line.strip()[:200]
    return ok, bad


def main():
    ap = argparse.ArgumentParser(description="검증이 초록이 될 때까지 기다렸다 전체 대조")
    ap.add_argument("--every", type=int, default=20, help="다시 볼 간격(분)")
    ap.add_argument("--hours", type=float, default=12.0, help="이 시간 안에서만 시도")
    ap.add_argument("--timeout", type=int, default=45, help="대조 한 회차의 제한(분)")
    a = ap.parse_args()

    deadline = time.time() + a.hours * 3600
    tries = 0
    while time.time() < deadline:
        tries += 1
        ok, bad = green()
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        if not ok:
            print(f"[{stamp}] 검증 빨강 — {a.every}분 뒤 다시. {bad}")
            _note({"at": stamp, "시도": tries, "결과": "검증대기", "메모": bad})
            time.sleep(a.every * 60)
            continue
        print(f"[{stamp}] 검증 초록 — 전체 대조를 돌린다")
        rc, out = _run(["daily_run.py"], a.timeout * 60)
        tail = "\n".join([l for l in out.splitlines() if l.strip()][-12:])
        _note({"at": stamp, "시도": tries,
               "결과": "대조완료" if rc == 0 else "대조실패(rc=%d)" % rc,
               "메모": tail[-1500:]})
        print(tail)
        if rc == 0:
            return 0
        # ★ 한 회차가 실패했다고 밤을 포기하지 않는다 (2026-08-06 실측: 첫 회차가
        #   90분 시간초과로 죽자 스크립트가 그대로 끝나 남은 밤이 통째로 비었다).
        #   막힌 원인이 잠깐의 잠금 다툼이면 다음 회차에 그냥 풀린다.
        print(f"[{stamp}] 대조 실패(rc={rc}) — {a.every}분 뒤 다시 시도한다")
        time.sleep(a.every * 60)

    _note({"at": datetime.now().strftime("%Y-%m-%d %H:%M"), "시도": tries,
           "결과": "시간초과", "메모": "정해진 시간 안에 검증이 초록이 되지 않았다"})
    print("정해진 시간 안에 검증이 초록이 되지 않았다 — 대조를 돌리지 않았다")
    return 2


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(1)
