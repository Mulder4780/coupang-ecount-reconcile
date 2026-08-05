# -*- coding: utf-8 -*-
"""
worksplit.py — 여러 세션이 **동시에** 일할 때 무엇을 누가 맡는지 적어 두는 분담판
================================================================================
사용자 지시(2026-08-05): "지금 현재 열려있는 세션과 병렬 작업 가능한 구조로 정리하고."

`ai_claim.py` 와 무엇이 다른가 — 둘은 층이 다르다.
  · `ai_claim`  = **자원** 잠금. "관리대장 파일을 지금 내가 쓴다" (배타·짧게 잡고 놓는다)
  · `worksplit` = **할 일** 분담. "밴드 잔여 재수집은 codex 가 맡았다" (길게 간다)

왜 따로 필요한가: 지금까지 남은 일은 AGENTS.md 산문에만 있었다. 두 세션이 그 글을
각자 읽고 **같은 일을 동시에** 시작해도 아무도 못 막는다(자원 잠금은 파일을 쓸 때만
부딪히므로, 몇 시간 조사한 뒤에야 중복이 드러난다). 그래서 "일 단위"로 먼저 나눈다.

죽은 세션은 기다리지 않는다
  주인의 PID 가 사라졌으면 그 자리에서 **주인 없음**으로 보고 다른 세션이 가져갈 수 있다
  (session_handoff.pid_alive 와 같은 판정 — 45분을 기다릴 이유가 없다).

  python worksplit.py                                  # 분담판 보기
  python worksplit.py --add "제목" --detail "설명" --lock band [--human]
  python worksplit.py --who claude --take 3            # 내가 맡는다
  python worksplit.py --who claude --done 3 --note "결과"
  python worksplit.py --who claude --drop 3            # 놓는다(다른 세션이 가져갈 수 있게)
  python worksplit.py --mine --who claude              # 내가 맡은 것만
  python worksplit.py --free                           # 아무도 안 맡은 것만(다음에 뭐 할까)
"""
import json
import os
import socket
import sys
import time
from contextlib import contextmanager
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
BOARD = os.path.join(ROOT, "reports", "worksplit.json")
GUARD = os.path.join(ROOT, "reports", ".worksplit.guard")
OUT_MD = os.path.join(ROOT, "reports", "작업분담.md")

WAIT, DOING, DONE, HOLD = "대기", "진행", "완료", "사람대기"
# 이 일을 하려면 어떤 자원 잠금이 필요한가(ai_claim 의 이름과 같게 쓴다).
LOCK_LABEL = {"ledger": "관리대장", "band": "밴드", "publish": "게시", "code": "코드",
              "erp": "ERP 화면", "": "없음"}

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def pid_alive(pid):
    """주인 세션이 아직 살아 있나. 모르면 None(함부로 죽었다고 하지 않는다)."""
    if not pid:
        return None
    try:
        import ctypes
        h = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
        if not h:
            return False
        ctypes.windll.kernel32.CloseHandle(h)
        return True
    except Exception:
        return None


@contextmanager
def guard(timeout=10):
    """읽기-고치기-쓰기 전체를 원자적으로 감싼다(두 세션이 같은 순간 잡는 것을 막는다)."""
    os.makedirs(os.path.dirname(BOARD), exist_ok=True)
    deadline = time.time() + timeout
    while True:
        try:
            os.mkdir(GUARD)
            break
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(GUARD) > 30:
                    os.rmdir(GUARD)
                    continue
            except OSError:
                pass
            if time.time() >= deadline:
                raise TimeoutError("분담판 잠금을 10초 안에 얻지 못했습니다")
            time.sleep(0.05)
    try:
        yield
    finally:
        try:
            os.rmdir(GUARD)
        except OSError:
            pass


def _read():
    try:
        with open(BOARD, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    d.setdefault("items", [])
    d.setdefault("seq", 0)
    return d


def _write(d):
    tmp = BOARD + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, BOARD)


def load():
    with guard():
        return _read()


def _owner_state(it):
    """주인이 살아 있나 — (표시할 주인, 가져가도 되나)."""
    if it.get("state") != DOING:
        return it.get("who") or "", it.get("state") in (WAIT, HOLD)
    alive = pid_alive(it.get("pid"))
    if alive is False:
        return f"{it.get('who')}(세션 종료)", True
    return it.get("who") or "", False


def add(title, detail="", lock="", human=False, who=""):
    with guard():
        d = _read()
        d["seq"] += 1
        it = {"id": d["seq"], "title": title, "detail": detail, "lock": lock or "",
              "state": HOLD if human else WAIT, "who": "", "pid": None,
              "at": datetime.now().strftime("%Y-%m-%d %H:%M"), "note": "",
              "by": who or ""}
        d["items"].append(it)
        _write(d)
    print(f"[{it['id']}] 추가 — {title}" + (" (사람이 해야 함)" if human else ""))
    return it["id"]


def take(who, wid):
    with guard():
        d = _read()
        for it in d["items"]:
            if it["id"] != wid:
                continue
            if it["state"] == DONE:
                print(f"[{wid}] 은 이미 완료된 일입니다.")
                return False
            shown, free_ = _owner_state(it)
            if it["state"] == DOING and it.get("who") != who and not free_:
                print(f"★ [{wid}] {it['title']} 은 이미 {shown} 가 맡았습니다.")
                print("  → 다른 일을 고르세요:  python worksplit.py --free")
                return False
            it.update(state=DOING, who=who, pid=os.getpid(),
                      host=socket.gethostname(),
                      at=datetime.now().strftime("%Y-%m-%d %H:%M"))
            _write(d)
            print(f"[{wid}] {it['title']} — {who} 가 맡음")
            if it.get("lock"):
                _lock_hint(it["lock"], who)
            return True
    print(f"[{wid}] 그런 번호가 없습니다.")
    return False


def _lock_hint(lock, who):
    """이 일에 필요한 자원을 지금 남이 잡고 있는지 알려 준다(막지는 않는다)."""
    try:
        import ai_claim
        cur = ai_claim.load().get(lock)
    except Exception:
        return
    if cur and cur.get("who") != who:
        print(f"  ※ 자원 '{LOCK_LABEL.get(lock, lock)}' 은 지금 {cur.get('who')} 가 잡고 있습니다 — "
              f"쓰기 단계 전에 놓을 때까지 기다리세요.")
    elif not cur:
        print(f"  ※ 쓰기 전에 잡으세요:  python ai_claim.py --who {who} --take {lock} --why \"...\"")


def finish(who, wid, note="", state=DONE):
    with guard():
        d = _read()
        for it in d["items"]:
            if it["id"] != wid:
                continue
            if it.get("who") and it["who"] != who and it["state"] == DOING:
                shown, free_ = _owner_state(it)
                if not free_:
                    print(f"★ [{wid}] 은 {shown} 것이라 손대지 않습니다.")
                    return False
            it.update(state=state, note=note or it.get("note", ""),
                      at=datetime.now().strftime("%Y-%m-%d %H:%M"))
            if state != DOING:
                it["pid"] = None
            if state == WAIT:
                it["who"] = ""
            _write(d)
            print(f"[{wid}] {it['title']} → {state}")
            return True
    print(f"[{wid}] 그런 번호가 없습니다.")
    return False


def board(items=None, only=None, who=None):
    d = load()
    rows = items if items is not None else d["items"]
    if only == "mine":
        rows = [x for x in rows if x.get("who") == who and x["state"] == DOING]
    elif only == "free":
        rows = [x for x in rows if x["state"] == WAIT or _owner_state(x)[1] and x["state"] != DONE]
    else:
        rows = [x for x in rows if x["state"] != DONE]
    if not rows:
        print("분담판이 비어 있습니다." if only is None else "해당하는 일이 없습니다.")
        return rows
    order = {DOING: 0, WAIT: 1, HOLD: 2, DONE: 3}
    rows = sorted(rows, key=lambda x: (order.get(x["state"], 9), x["id"]))
    print("번호 상태     주인            자원      할 일")
    for x in rows:
        shown, _ = _owner_state(x)
        print(f"{x['id']:>3}  {x['state']:<7} {shown:<14} "
              f"{LOCK_LABEL.get(x.get('lock', ''), x.get('lock')):<8} {x['title']}")
    return rows


def render_md():
    """사람과 **다른 AI 세션**이 읽을 한 장. 분담판의 진실은 json 이고 이건 사본이다."""
    d = load()
    live = [x for x in d["items"] if x["state"] != DONE]
    done = [x for x in d["items"] if x["state"] == DONE][-10:]
    L = [f"# 작업 분담판 (갱신 {datetime.now().strftime('%Y-%m-%d %H:%M')})", "",
         "여러 세션(Claude·Codex)이 동시에 일할 때 **같은 일을 두 번 하지 않게** 나눠 적는 곳이다.",
         "일을 시작하기 전에 여기서 고르고 `--take` 로 이름을 적는다.", "",
         "```bash",
         "python ecount/worksplit.py --free                 # 아무도 안 맡은 일",
         "python ecount/worksplit.py --who claude --take 3  # 내가 맡는다",
         "python ecount/worksplit.py --who claude --done 3 --note \"결과\"",
         "```", "",
         f"## 지금 남은 일 ({len(live)})", "",
         "| 번호 | 상태 | 주인 | 자원 | 할 일 | 메모 |", "|---:|---|---|---|---|---|"]
    order = {DOING: 0, WAIT: 1, HOLD: 2}
    for x in sorted(live, key=lambda x: (order.get(x["state"], 9), x["id"])):
        shown, _ = _owner_state(x)
        L.append(f"| {x['id']} | {x['state']} | {shown or '—'} | "
                 f"{LOCK_LABEL.get(x.get('lock', ''), x.get('lock'))} | "
                 f"{x['title']}{(' — ' + x['detail']) if x.get('detail') else ''} | "
                 f"{x.get('note', '')} |")
    if done:
        L += ["", "## 최근 끝낸 일", ""]
        for x in done:
            L.append(f"- [{x['id']}] {x['title']} — {x.get('who', '')} ({x['at']})"
                     + (f" · {x['note']}" if x.get("note") else ""))
    os.makedirs(os.path.dirname(OUT_MD), exist_ok=True)
    open(OUT_MD, "w", encoding="utf-8").write("\n".join(L) + "\n")
    return OUT_MD


def summary():
    """다른 도구(session_handoff·대시보드)가 한 줄로 쓸 요약."""
    d = load()
    c = {WAIT: 0, DOING: 0, HOLD: 0, DONE: 0}
    orphan = 0
    for x in d["items"]:
        c[x["state"]] = c.get(x["state"], 0) + 1
        if x["state"] == DOING and _owner_state(x)[1]:
            orphan += 1
    return {"대기": c[WAIT], "진행": c[DOING], "사람대기": c[HOLD], "완료": c[DONE],
            "주인없음": orphan}


def main():
    a = sys.argv[1:]

    def get(f, d=None):
        return a[a.index(f) + 1] if f in a and len(a) > a.index(f) + 1 else d

    who = (get("--who") or os.environ.get("CSOS_AI") or "").strip().lower()
    try:
        if "--add" in a:
            add(get("--add"), get("--detail", ""), get("--lock", ""), "--human" in a, who)
        elif "--take" in a:
            if not who:
                print("--who 가 필요합니다 (claude|codex|...)")
                return 2
            if not take(who, int(get("--take"))):
                return 2
        elif "--done" in a:
            finish(who, int(get("--done")), get("--note", ""), DONE)
        elif "--drop" in a:
            finish(who, int(get("--drop")), get("--note", ""), WAIT)
        elif "--hold" in a:
            finish(who, int(get("--hold")), get("--note", ""), HOLD)
        elif "--mine" in a:
            board(only="mine", who=who)
        elif "--free" in a:
            board(only="free", who=who)
        elif "--summary" in a:
            print(json.dumps(summary(), ensure_ascii=False))
        else:
            board()
            print("\n분담판 문서: reports/작업분담.md")
    finally:
        try:
            render_md()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
