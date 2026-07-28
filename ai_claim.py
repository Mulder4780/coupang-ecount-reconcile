# -*- coding: utf-8 -*-
"""
ai_claim.py — 여러 AI가 **동시에** 일할 때 서로 밟지 않게 조율한다
================================================================================
사용자 지시(2026-07-28): "코덱스랑 동시 작업중일 때 둘이 대화해서 우선순위 정해서
협업해서 진행해."

AI끼리 직접 대화할 수단은 없다. 이 프로젝트의 진실의 원천은 **파일**이므로(AGENTS.md),
조율도 파일로 한다. 시작 전에 '내가 이걸 잡는다'를 적고, 끝나면 놓는다.

무엇을 막으려는 것인가
  · 관리대장은 **동시에 두 AI가 고치면 안 된다.** 각자 vN+1을 만들어 한쪽 작업이
    통째로 묻힌다(둘 다 v204를 읽고 각자 v205를 만들면 하나는 사라진다).
  · 같은 파일을 동시에 편집하면 뒤에 push 하는 쪽이 앞을 덮는다.

우선순위 (충돌 시 이 순서로 양보한다 — 사람이 매번 정하지 않아도 되게)
  1) 사용자가 **지금 대화 중인** AI가 우선. 상대는 읽기/분석으로 돌린다.
  2) 원장 쓰기(ledger_writer·workbook_patch·expand_rows·fix_*) 는 **한 번에 하나**.
  3) 나머지(리포트 생성·조회·문서화)는 동시에 해도 된다.
  4) 같은 순위면 **먼저 잡은 쪽**이 우선. 늦게 온 쪽은 다른 일을 한다.

  python ai_claim.py --who claude --take ledger --why "confirm_fill 반영"
  python ai_claim.py                      # 지금 누가 뭘 잡고 있나
  python ai_claim.py --who claude --free ledger
"""
import sys, os, json, time, socket
from contextlib import contextmanager
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
CLAIMS = os.path.join(ROOT, "reports", "ai_claims.json")
GUARD = os.path.join(ROOT, "reports", ".ai_claims.guard")
STALE = 45 * 60          # 45분 넘게 안 놓으면 죽은 것으로 본다(크레딧 소진·중단)

# 잡을 수 있는 것들. ledger/code/band/publish는 배타적이다.
LOCKS = {
    "ledger":  ("관리대장 쓰기", True),
    "band":    ("밴드 수집·반영", True),
    "publish": ("고정 주소·폰 사본 게시", True),
    "code":    ("코드·설정 변경", True),
    "report":  ("리포트·문서 작성", False),
    "read":    ("조회·분석만", False),
}

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


@contextmanager
def state_guard(timeout=10):
    """점유 JSON의 read-modify-write 전체를 원자적으로 감싼다."""
    os.makedirs(os.path.dirname(CLAIMS), exist_ok=True)
    deadline = time.time() + timeout
    while True:
        try:
            os.mkdir(GUARD)
            with open(os.path.join(GUARD, "owner.json"), "w", encoding="utf-8") as f:
                json.dump({"pid": os.getpid(), "host": socket.gethostname(), "at": time.time()}, f)
            break
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(GUARD) > 30:
                    owner = os.path.join(GUARD, "owner.json")
                    if os.path.exists(owner):
                        os.remove(owner)
                    os.rmdir(GUARD)
                    continue
            except OSError:
                pass
            if time.time() >= deadline:
                raise TimeoutError("AI 점유 파일 잠금을 10초 안에 얻지 못했습니다")
            time.sleep(0.05)
    try:
        yield
    finally:
        try:
            owner = os.path.join(GUARD, "owner.json")
            if os.path.exists(owner):
                os.remove(owner)
            os.rmdir(GUARD)
        except OSError:
            pass


def _load_unlocked():
    try:
        with open(CLAIMS, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    now = time.time()
    return {k: v for k, v in d.items() if now - v.get("at", 0) < STALE}


def load():
    with state_guard():
        return _load_unlocked()


def _save_unlocked(d):
    tmp = CLAIMS + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, CLAIMS)


def save(d):
    with state_guard():
        _save_unlocked(d)


def show(d=None):
    d = d if d is not None else load()
    if not d:
        print("지금 잡혀 있는 작업 없음 — 무엇이든 시작해도 됩니다.")
        return
    print("현재 점유 상황")
    for k, v in sorted(d.items()):
        label, excl = LOCKS.get(k, (k, False))
        mins = int((time.time() - v.get("at", 0)) / 60)
        print(f"  [{'배타' if excl else '공유'}] {label:<18} {v.get('who','?'):<8} "
              f"{mins}분 전 · {v.get('why','')[:40]}")


def take(who, what, why=""):
    label, excl = LOCKS.get(what, (what, True))
    with state_guard():
        d = _load_unlocked()
        cur = d.get(what)
        if cur and cur.get("who") != who and excl:
            mins = int((time.time() - cur.get("at", 0)) / 60)
            print(f"★ '{label}' 은 이미 {cur['who']} 가 잡고 있습니다 ({mins}분 전 · {cur.get('why','')}).")
            print("  배타 작업이라 동시에 하면 한쪽 결과가 통째로 묻힙니다.")
            print("  → 다른 일을 먼저 하거나, 상대가 끝낼 때까지 조회·분석만 하세요.")
            print(f"  (상대가 멈춘 것 같으면 {STALE // 60}분 뒤 자동 해제됩니다)")
            return False
        d[what] = {
            "who": who, "why": why, "at": time.time(),
            "when": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "pid": os.getpid(), "host": socket.gethostname(),
        }
        _save_unlocked(d)
    print(f"'{label}' 점유 — {who}" + (f" · {why}" if why else ""))
    return True


def free(who, what):
    with state_guard():
        d = _load_unlocked()
        if what in d and d[what].get("who") not in (who, None):
            print(f"★ '{what}' 은 {d[what]['who']} 것이라 놓을 수 없습니다.")
            return False
        d.pop(what, None)
        _save_unlocked(d)
    print(f"'{what}' 놓음 — {who}")
    return True


def main():
    a = sys.argv[1:]
    get = lambda f: (a[a.index(f) + 1] if f in a and len(a) > a.index(f) + 1 else None)
    who = get("--who") or "unknown"
    if "--take" in a:
        sys.exit(0 if take(who, get("--take"), get("--why") or "") else 2)
    if "--free" in a:
        free(who, get("--free"))
        return
    if "--free-all" in a:
        with state_guard():
            d = _load_unlocked()
            for k in [k for k, v in d.items() if v.get("who") == who]:
                d.pop(k)
            _save_unlocked(d)
        print(f"{who} 의 점유를 모두 놓았습니다.")
        return
    show()
    print("\n잡기:  python ai_claim.py --who claude --take ledger --why \"이유\"")
    print("놓기:  python ai_claim.py --who claude --free ledger")


if __name__ == "__main__":
    main()
