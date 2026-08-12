# -*- coding: utf-8 -*-
"""dump_watch.py — 밴드 덤프가 떨어지면 **사람 손 없이** 흡수·대조·큐적재까지 이어 간다.

왜 (2026-08-06 지시: "내 손 안 가게")
  사람이 로그인된 탭에서 수집기를 돌리면 dump_*.json 이 다운로드 폴더에 떨어진다.
  거기서 멈추면 사람이 다시 명령을 내려야 한다 — 그러면 다음에도 또 시켜야 한다.
  이 감시기가 새 덤프를 보는 즉시 아래를 순서대로 돌린다.

    download_intake --apply   다운로드 → Z: 밴드 원본으로 이동(내용 판별)
    band/convert_dump.py      덤프 → band/cache 병합(수정글 감지 포함)
    band/band_reconcile.py    캐시 ↔ 관리대장 대조 → 입력 큐 적재
    ledger_db.py --intake     큐 → DB (엑셀 반영은 11:00·15:00 회차 규칙 그대로)

  ★ 엑셀은 열지 않는다. 반영 시각 규칙을 이 감시기가 깨면 안 된다.

  python band/dump_watch.py --minutes 90
"""
import argparse
import glob
import json
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable
if hasattr(sys.stdout, "reconfigure"):   # pythonw 는 sys.stdout 이 None 이다([43])
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WATCH = [os.path.join(os.path.expanduser("~"), "Downloads"),
         os.path.join(os.path.expanduser("~"), "Desktop"),
         os.path.join(os.path.expanduser("~"), "OneDrive", "바탕 화면")]
STATE = os.path.join(ROOT, "reports", "밴드덤프_감시.json")


def seen_dumps():
    out = set()
    for d in WATCH:
        if os.path.isdir(d):
            out |= {os.path.basename(p) for p in glob.glob(os.path.join(d, "dump_*.json"))}
    return out


def run(name, args, timeout=1800):
    try:
        r = subprocess.run([PY] + args, cwd=ROOT, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout)
        lines = [x for x in (r.stdout or "").splitlines()
                 if x.strip() and "UserWarning" not in x and "관리대장 최신본" not in x]
        tail = lines[-1] if lines else ""
        ok = r.returncode == 0
    except Exception as exc:
        ok, tail = False, str(exc)[:90]
    print(f"  [{'OK ' if ok else 'FAIL'}] {name} — {tail[:110]}", flush=True)
    return {"단계": name, "성공": ok, "끝줄": tail[:200]}


def ingest():
    print(f"새 덤프 감지 — 흡수 시작 {time.strftime('%H:%M:%S')}", flush=True)
    steps = [run("다운로드 흡수", [os.path.join(ROOT, "download_intake.py"), "--apply"]),
             run("덤프 → 캐시", [os.path.join(ROOT, "band", "convert_dump.py")]),
             run("밴드 대조", [os.path.join(ROOT, "band", "band_reconcile.py")]),
             run("큐 → DB", [os.path.join(ROOT, "ledger_db.py"), "--intake"])]
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    hist = []
    try:
        hist = json.load(open(STATE, encoding="utf-8"))
    except Exception:
        pass
    hist.append({"시각": time.strftime("%Y-%m-%d %H:%M:%S"), "단계": steps})
    json.dump(hist[-30:], open(STATE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return steps


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--minutes", type=int, default=90)
    ap.add_argument("--once", action="store_true", help="지금 있는 것만 흡수하고 끝")
    a = ap.parse_args()
    if a.once:
        ingest()
        return 0
    known = seen_dumps()
    print(f"감시 시작 — {a.minutes}분 · 이미 있는 덤프 {len(known)}개는 무시", flush=True)
    end = time.time() + a.minutes * 60
    while time.time() < end:
        now = seen_dumps()
        fresh = now - known
        if fresh:
            # 다운로드가 끝날 때까지 잠깐 기다린다(.crdownload 가 사라진 뒤가 안전하다)
            time.sleep(5)
            print("새 파일:", ", ".join(sorted(fresh)), flush=True)
            ingest()
            known = seen_dumps()
        time.sleep(10)
    print("감시 종료", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
