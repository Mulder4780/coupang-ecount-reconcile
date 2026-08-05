# -*- coding: utf-8 -*-
"""
call_notes.py — 통화·회의에서 정해진 것을 **파일과 원장으로** 남긴다
================================================================================
사용자 지시(2026-08-05): 통화 녹음 텍스트를 주며 "필요한 것만 반영, 욕은 반영하지 마."

왜 도구로 만드나: 대표는 현장·본사와 통화로 일을 정한다. 그런데 통화에서 정해진 것은
녹음 앱 안에만 남아 아무도 다시 보지 않는다 — 이 프로젝트의 상시 원칙("대화에 남긴 것은
사라진다. 파일에 넣은 것만 산다")이 그대로 적용되는 자리다. 다음에 또 통화 기록을 줄
것이므로 한 번 하고 끝내지 않고 도구로 만든다.

무엇을 남기고 무엇을 버리나 — **이것이 이 도구의 핵심이다**
  남긴다: 정해진 것 · 할 일(누가·언제) · 확인된 사실(재고 없음·절차 병목 등)
  버린다: 녹취 전문 · 욕설 · 연봉/인사/개인 평가 · 남 흉보는 대목
  판단은 사람(또는 AI)이 md 를 쓸 때 하고, 이 도구는 **쓴 것만** 옮긴다.
  녹취 원문을 이 폴더에 넣지 않는다 — 한 번 들어가면 지우기 어렵고, 공유 폴더다.

하는 일
  1. `## 할 일` 아래 `- [담당 · 기한] 내용` 줄을 뽑아 19시트 인수인계로 예약한다
     (`ledger_db.handoff_add` — 엑셀은 11:00·15:00 회차에만 열린다)
  2. 메모 자체를 `0. 원본 자료/10. 통화·회의 기록/` 에 날짜 이름으로 보관한다
  3. `reports/통화기록_색인.json` 에 색인을 남겨 앱·검색이 찾을 수 있게 한다

  python call_notes.py --file reports/통화_20260805_김준형.md --with 김준형 --on 2026-08-05
  python call_notes.py --list
"""
import json
import os
import re
import shutil
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
INDEX = os.path.join(ROOT, "reports", "통화기록_색인.json")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# `- [김준형 · 2026-08-06] 내용` / `- [AS팀 전원] 내용` 둘 다 받는다.
TODO_RE = re.compile(r"^\s*-\s*\[([^\]]+)\]\s*(.+?)\s*$")


def parse_todos(text):
    """`## 할 일` 절 안의 항목만 뽑는다. 다른 절의 목록은 건드리지 않는다."""
    out, inside = [], False
    for line in text.splitlines():
        if line.startswith("## "):
            inside = "할 일" in line
            continue
        if not inside:
            continue
        m = TODO_RE.match(line)
        if not m:
            continue
        head, body = m.group(1), m.group(2)
        parts = [p.strip() for p in head.split("·")]     # `담당 · 기한` 또는 `담당`
        out.append({"who": parts[0], "due": parts[1] if len(parts) > 1 else "", "what": body})
    return out


def note_dir():
    try:
        import source_dirs
        return source_dirs.CALL_NOTE_DIR
    except Exception:
        return os.path.join(ROOT, "reports", "통화기록")


def load_index():
    try:
        return json.load(open(INDEX, encoding="utf-8"))
    except Exception:
        return {"notes": []}


def add(path, whom="", on="", queue=True):
    if not os.path.exists(path):
        print(f"파일이 없습니다: {path}")
        return 2
    text = open(path, encoding="utf-8").read()
    todos = parse_todos(text)

    # 1) 할 일 → 19시트 인수인계 예약 (엑셀은 11:00·15:00 회차에만 열린다)
    queued = 0
    if queue and todos:
        try:
            import ledger_db
            for t in todos:
                title = f"[통화 {on or ''} {whom}] {t['what']}"[:120]
                detail = (f"담당 {t['who']}" + (f" · 기한 {t['due']}" if t["due"] else "")
                          + f" · 근거: 통화 기록 {os.path.basename(path)}")
                ledger_db.handoff_add(title, detail)
                queued += 1
        except Exception as exc:
            print(f"★ 인수인계 예약 실패(메모 보관은 계속합니다): {exc}")

    # 2) 원본 자료로 보관
    dst_dir = note_dir()
    saved = ""
    try:
        os.makedirs(dst_dir, exist_ok=True)
        saved = os.path.join(dst_dir, os.path.basename(path))
        shutil.copy2(path, saved)
    except OSError as exc:                      # Z: 가 안 붙어 있어도 예약은 남는다
        print(f"★ 원본 폴더 보관 실패(Z: 확인 필요): {exc}")

    # 3) 색인
    idx = load_index()
    idx["notes"] = [n for n in idx.get("notes", []) if n.get("file") != os.path.basename(path)]
    idx["notes"].append({"file": os.path.basename(path), "with": whom, "on": on,
                         "todos": todos, "saved": saved})
    idx["notes"].sort(key=lambda n: (n.get("on") or "", n.get("file") or ""))
    os.makedirs(os.path.dirname(INDEX), exist_ok=True)
    json.dump(idx, open(INDEX, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print(f"통화 기록 반영 — {whom or '상대 미상'} {on}: 할 일 {len(todos)}건"
          f"(인수인계 예약 {queued}건)" + (f" · 보관 {saved}" if saved else ""))
    for t in todos:
        print(f"  · [{t['who']}{(' · ' + t['due']) if t['due'] else ''}] {t['what']}")
    return 0


def show():
    idx = load_index()
    if not idx.get("notes"):
        print("보관된 통화 기록이 없습니다.")
        return 0
    for n in idx["notes"]:
        print(f"{n.get('on',''):<12} {n.get('with',''):<10} 할 일 {len(n.get('todos') or [])}건"
              f"  {n.get('file','')}")
    return 0


def main():
    a = sys.argv[1:]

    def get(f, d=""):
        return a[a.index(f) + 1] if f in a and len(a) > a.index(f) + 1 else d

    if "--list" in a or not a:
        return show()
    return add(get("--file"), get("--with"), get("--on"), "--no-queue" not in a)


if __name__ == "__main__":
    sys.exit(main())
