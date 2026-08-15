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


def _run(args, timeout, tag="run"):
    """한 회차를 돌린다. 출력은 **파일로 흘려보낸다.**

    ★ 왜 capture_output 을 안 쓰나 (2026-08-07 실사고). 시간초과가 나면
      subprocess 는 모아 둔 출력을 **버린다**. 그래서 밤새 8회가 rc=124 로
      죽었는데 기록에 남은 것은 "시간 초과" 네 글자뿐이었고, 어느 단계에서
      느렸는지 알 길이 없어 원인 찾는 데 아침을 다 썼다. 파일로 흘리면
      죽은 자리까지의 로그가 그대로 남는다.
    """
    log = os.path.join(ROOT, "reports", f"밤샘_{tag}.log")
    try:
        os.makedirs(os.path.dirname(log), exist_ok=True)
    except Exception:
        pass
    try:
        with open(log, "w", encoding="utf-8", errors="replace") as fh:
            p = subprocess.run([PY] + args, cwd=ROOT, timeout=timeout,
                               stdout=fh, stderr=subprocess.STDOUT, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            rc = p.returncode
    except subprocess.TimeoutExpired:
        rc = 124
    except Exception as e:
        return 1, str(e)[:400]
    try:
        out = open(log, encoding="utf-8", errors="replace").read()
    except Exception:
        out = ""
    if rc == 124:
        tail = "\n".join([l for l in out.splitlines() if l.strip()][-15:])
        return 124, f"시간 초과 — 죽기 직전까지의 로그:\n{tail}"
    return rc, out[-4000:]


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
    rc, out = _run([os.path.join("tests", "synthetic_check.py")], 1800, tag="검증")
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
    # ★ 45분이 아니다 (2026-08-07 실측). 전체 대조 한 회차는 지금 **약 2시간** 걸린다
    #   (밴드 캐시가 2천 → 8천5백 글로 커지면서 색인·대조가 그만큼 무거워졌다).
    #   45분이던 시절 기준을 그대로 두는 바람에 밤새 8회가 전부 "거의 다 해 놓고" 잘렸다.
    #   넉넉히 잡는 편이 안전하다 — 어차피 끝나면 그 회차로 밤을 끝낸다.
    ap.add_argument("--timeout", type=int, default=240, help="대조 한 회차의 제한(분)")
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
        rc, out = _run(["daily_run.py"], a.timeout * 60, tag="대조")
        tail = "\n".join([l for l in out.splitlines() if l.strip()][-12:])
        # ★ "이미 실행 중" 은 **성공이 아니다** — daily_run 은 그때도 0 으로 끝난다.
        #   그대로 두면 이 스크립트가 "다 됐다"며 밤을 끝내 버린다(감시기를 두 개 띄운
        #   순간 실제로 그렇게 된다). 아무 일도 안 한 회차로 보고 다시 시도한다.
        if "이미 실행 중" in out:
            rc = 125
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
