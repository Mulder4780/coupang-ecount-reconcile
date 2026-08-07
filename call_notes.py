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

★ 보관 자리가 바뀌었다 — DB 하나다 (2026-08-07 지시)
  사용자 지시: **"통화_MD는 원본 자료에서 안 보이게 처리하고 DB만 보관해(민감한 내용이
  포함되어있음)."** 예전에는 메모를 `0. 원본 자료/10. 통화·회의 기록/` 으로 복사했는데,
  그 폴더는 **공유 폴더(Z:)** 이고 앱의 '원본 자료' 목록에 그대로 떴다
  (2026-08-07 실측: `통화_20260805_김준형.md` 가 카드로 노출).
  이제 본문은 `ledger_db.call_note_save()` 로 **DB 에만** 들어간다.
  숨기는 것이 아니라 **파일을 두지 않는 것**이 조치다 — 없는 파일은 샐 수 없다.
  이미 Z: 에 있던 것은 `--migrate` 가 DB 로 옮기고 원본을 지운다.

하는 일
  1. `## 할 일` 아래 `- [담당 · 기한] 내용` 줄을 뽑아 19시트 인수인계로 예약한다
     (`ledger_db.handoff_add` — 엑셀은 11:00·15:00 회차에만 열린다)
  2. 메모 본문을 DB(`call_note` 표)에 보관한다. 같은 파일 이름이면 갱신한다.
  3. 원본 색인(source_index.py)은 통화 메모를 **절대 담지 않는다** — 검증 [129].

  python call_notes.py --file reports/통화_20260805_김준형.md --with 김준형 --on 2026-08-05
  python call_notes.py --list          # 목록(본문은 안 보여 준다)
  python call_notes.py --show 통화_20260805_김준형.md   # 본문까지
  python call_notes.py --migrate       # Z: 에 남은 메모를 DB 로 옮기고 원본 삭제
"""
import os
import re
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

    # 2) 본문은 DB 에만 (2026-08-07 지시). Z: 공유 폴더로 복사하지 않는다.
    try:
        import ledger_db
        ledger_db.call_note_save(os.path.basename(path), text, whom=whom, on=on, todos=todos)
        saved = "DB(call_note)"
    except Exception as exc:                    # DB 가 막혀도 예약은 남는다
        print(f"★ DB 보관 실패(메모 원본을 지우지 마세요): {exc}")
        return 1

    print(f"통화 기록 반영 — {whom or '상대 미상'} {on}: 할 일 {len(todos)}건"
          f"(인수인계 예약 {queued}건) · 보관 {saved}")
    for t in todos:
        print(f"  · [{t['who']}{(' · ' + t['due']) if t['due'] else ''}] {t['what']}")
    return 0


def migrate():
    """Z: 에 남아 있던 통화 메모를 DB 로 옮기고 **원본을 지운다**(2026-08-07 지시).

    지우기 전에 DB 에 같은 본문이 들어갔는지 반드시 확인한다 — 옮기다 만 상태로
    파일만 사라지는 것이 제일 나쁘다.
    """
    import ledger_db
    d = note_dir()
    if not os.path.isdir(d):
        print(f"통화 기록 폴더가 없습니다(옮길 것 없음): {d}")
        return 0
    names = [f for f in sorted(os.listdir(d)) if f.lower().endswith(".md")]
    if not names:
        print("Z: 에 남은 통화 메모가 없습니다.")
    moved, kept = 0, []
    for fn in names:
        src = os.path.join(d, fn)
        try:
            text = open(src, encoding="utf-8").read()
        except OSError as exc:
            kept.append(f"{fn} (읽기 실패: {exc})")
            continue
        # 파일 이름에서 날짜·상대를 추정한다: 통화_20260805_김준형.md
        m = re.match(r"^통화_(\d{4})(\d{2})(\d{2})_(.+)\.md$", fn)
        on = f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else ""
        whom = m.group(4) if m else ""
        ledger_db.call_note_save(fn, text, whom=whom, on=on, todos=parse_todos(text))
        back = ledger_db.call_note_get(fn)
        if not back or back.get("body") != text:      # 확인되기 전에는 절대 안 지운다
            kept.append(f"{fn} (DB 확인 실패 — 원본 유지)")
            continue
        try:
            os.remove(src)
            moved += 1
            print(f"  DB 로 옮김 · 원본 삭제: {fn}")
        except OSError as exc:
            kept.append(f"{fn} (삭제 실패: {exc})")

    # 예전 JSON 색인도 통화 내용을 담고 있었다 — DB 정본만 남긴다.
    if os.path.exists(INDEX):
        try:
            os.remove(INDEX)
            print(f"  예전 색인 삭제: {os.path.basename(INDEX)}")
        except OSError as exc:
            kept.append(f"{os.path.basename(INDEX)} (삭제 실패: {exc})")

    print(f"통화 메모 이관 — DB {moved}건, 남은 것 {len(kept)}건")
    for k in kept:
        print(f"  ★ {k}")
    return 1 if kept else 0


def show():
    """목록만 — 본문은 찍지 않는다(민감). 본문은 --show <파일이름> 으로 한 건씩."""
    import ledger_db
    notes = ledger_db.call_notes()
    if not notes:
        print("보관된 통화 기록이 없습니다.")
        return 0
    for n in notes:
        print(f"{n.get('on_date',''):<12} {n.get('whom',''):<10} 할 일 {len(n.get('todos') or [])}건"
              f"  {n.get('file','')}")
    return 0


def show_one(name):
    import ledger_db
    n = ledger_db.call_note_get(name)
    if not n:
        print(f"그런 통화 기록이 없습니다: {name}")
        return 2
    print(f"── {n.get('file')} · {n.get('whom','')} · {n.get('on_date','')} ──")
    print(n.get("body") or "")
    return 0


def main():
    a = sys.argv[1:]

    def get(f, d=""):
        return a[a.index(f) + 1] if f in a and len(a) > a.index(f) + 1 else d

    if "--migrate" in a:
        return migrate()
    if "--show" in a:
        return show_one(get("--show"))
    if "--list" in a or not a:
        return show()
    return add(get("--file"), get("--with"), get("--on"), "--no-queue" not in a)


if __name__ == "__main__":
    sys.exit(main())
