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

★ 2026-08-05 — **세션 단위**로 잡는다 (사용자 지시: "지금 현재 열려있는 세션과
   병렬 작업 가능한 구조로 정리")
  그동안 주인은 `--who claude` 한 단어였다. 그런데 같은 프로젝트 폴더에 **Claude 세션이
  둘 이상** 떠 있으면 둘 다 주인이 "claude" 라서
    · 뒤에 온 세션이 앞 세션의 배타 점유를 **말없이 빼앗고**(둘 다 vN+1 을 만든다)
    · `--free-all` 이 **남의 점유까지** 놓아 버렸다(PreCompact 자동 마무리가 특히 위험).
  이제 주인은 `who` 가 아니라 **세션 식별자(sid)** 다. sid 는 환경변수
  `CLAUDE_CODE_SESSION_ID` 에서 자동으로 온다 — 사람이 외우거나 넘길 것이 없다.
  같은 sid 면 다시 잡아도 되고(재진입), 다른 sid 면 `who` 가 같아도 못 빼앗는다.
"""
import sys, os, json, time, socket, hashlib, subprocess
from contextlib import contextmanager
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))

# ★ 점유 파일은 **본체 체크아웃 하나**에만 둔다 (2026-08-06).
#   git 워크트리(`.claude/worktrees/<이름>`)에서 돌면 ROOT 가 워크트리라서
#   점유 파일이 본체와 갈렸다. 두 세션이 동시에 `ledger` 를 잡아도 서로 안 보였다
#   — 관리대장 동시 쓰기 금지(CLAUDE.md)가 조용히 무너지는 자리였다.
#   본체에서는 `shared()` 가 ROOT 를 그대로 돌려주므로 동작이 하나도 안 바뀐다.
#   링크로 잇지 않는 이유: `_save_unlocked` 가 `os.replace` 로 갈아치우는데,
#   그 순간 하드링크가 끊겨 두 파일로 갈라진다.
try:
    from worktree_state import shared as _shared
    STATE_DIR = _shared("reports")
except Exception:
    STATE_DIR = os.path.join(ROOT, "reports")

CLAIMS = os.path.join(STATE_DIR, "ai_claims.json")
GUARD = os.path.join(STATE_DIR, ".ai_claims.guard")
WORKCENTER_ACTIVITY = os.path.join(STATE_DIR, "workcenter_activity.json")
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


# ── 세션 식별 (2026-08-05) ───────────────────────────────────────────────────
#  주인을 `who`(claude/codex) 가 아니라 **세션**으로 본다. 같은 폴더에 Claude 세션이
#  둘 떠 있어도 서로 못 빼앗게 하려는 것이다. 식별자는 환경에서 저절로 온다 —
#  사람이 외우거나 명령줄로 넘길 것이 없다(외우게 하면 언젠가 안 넘긴다).
SID_ENV = ("CLAUDE_CODE_SESSION_ID", "CLAUDE_CODE_HOST_SESSION_ID",
           "CODEX_SESSION_ID", "AI_SESSION_ID")


def session_id():
    """이 세션의 짧고 안정적인 식별자. 한 세션 안에서는 항상 같은 값이 나온다."""
    for key in SID_ENV:
        raw = (os.environ.get(key) or "").strip()
        if raw:
            return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]
    # 환경변수가 없는 곳(스케줄러·수동 실행)은 호스트+에이전트 PID 로 대신한다.
    raw = "%s/%s" % (socket.gethostname(), os.environ.get("CLAUDE_PID") or "manual")
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]


def agent_pid():
    """세션을 붙들고 있는 프로세스(있으면). 죽었으면 점유를 45분 기다리지 않고 푼다."""
    try:
        return int(os.environ.get("CLAUDE_PID") or 0)
    except ValueError:
        return 0


def _pid_alive(pid, host):
    """다른 PC 의 점유는 판정하지 않는다 — 모르면 살아 있다고 본다(안전한 쪽)."""
    if not pid or host != socket.gethostname():
        return True
    try:
        if os.name == "nt":
            out = subprocess.run(["tasklist", "/FI", "PID eq %d" % pid, "/NH"],
                                 capture_output=True, text=True, timeout=15).stdout or ""
            return str(pid) in out
        os.kill(pid, 0)
        return True
    except Exception:
        return True


def _is_mine(claim, who):
    """이 점유가 '지금 이 세션' 것인가.

    옛 형식(sid 없음)은 같은 who 면 내 것으로 본다 — 과도기 호환. 새로 잡는 순간
    sid 가 붙으므로 이 예외는 한 번만 쓰인다.
    """
    if not isinstance(claim, dict):
        return False
    sid = claim.get("sid")
    if sid:
        return sid == session_id()
    return claim.get("who") == who


def _is_dead(claim):
    """주인 세션이 이미 죽었나(크레딧 소진·창 닫힘). 죽었으면 즉시 넘겨받아도 된다."""
    if not isinstance(claim, dict):
        return True
    # 스케줄러 점유는 agent_pid 가 0 이다(에이전트가 아니라서). 그때는 프로세스
    # `pid` 가 증거다 — session_handoff 의 판정과 같은 폴백을 쓴다(2026-08-07 실사고:
    # 죽은 ledger_writer 점유를 --check 는 죽었다 하고 --adopt 는 살았다 해서 교착).
    pid = claim.get("agent_pid") or claim.get("pid")
    if not pid:
        return False                     # 알 수 없으면 살아 있다고 본다
    return not _pid_alive(int(pid), claim.get("host") or "")


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
    me = session_id()
    print(f"현재 점유 상황  (이 세션 = {me})")
    for k, v in sorted(d.items()):
        label, excl = LOCKS.get(k, (k, False))
        mins = int((time.time() - v.get("at", 0)) / 60)
        sid = v.get("sid") or "옛형식"
        mark = "← 내 것" if _is_mine(v, v.get("who")) and sid == me else ""
        dead = " · 세션 종료됨(넘겨받을 수 있음)" if _is_dead(v) else ""
        print(f"  [{'배타' if excl else '공유'}] {label:<18} {v.get('who','?'):<7}"
              f"[{sid}] {mins}분 전 · {v.get('why','')[:36]}{mark}{dead}")


def _sol_write_gate(who, what):
    """Require the Terra -> Sol review before Sol changes shared state.

    Read/report claims remain available so Sol can inspect the handoff and run
    the review itself. Any exclusive claim is fail-closed when a pending Terra
    marker exists.
    """
    _label, exclusive = LOCKS.get(what, (what, True))
    if "sol" not in (who or "").lower() or not exclusive:
        return True
    try:
        import handoff_review
        allowed = handoff_review.sol_review_is_current()
    except Exception as exc:
        print("Terra 인수인계 검토 상태를 읽을 수 없어 Sol 쓰기 작업을 차단합니다: %s" % exc)
        return False
    if allowed:
        return True
    print("Terra 인수인계 검토가 끝나지 않아 Sol 쓰기 작업을 차단합니다.")
    print("  먼저 실행: python handoff_review.py --review-sol")
    print("  합성 검증을 포함한 PASS 후에만 ledger/code/band/publish 점유가 가능합니다.")
    return False


def _workcenter_priority_gate(what):
    """담당자가 업무센터에서 입력 중이면 새 배타 작업을 잠시 미룬다."""
    _label, exclusive = LOCKS.get(what, (what, True))
    if not exclusive:
        return True
    try:
        activity = json.load(open(WORKCENTER_ACTIVITY, encoding="utf-8"))
        if time.time() < float(activity.get("active_until_ts") or 0):
            name = activity.get("name") or activity.get("slug") or "담당자"
            print(f"★ {name} 업무센터 입력이 진행 중이라 '{what}' 점유를 잠시 미룹니다.")
            print("  담당자 입력이 끝난 뒤 약 2분 동안 추가 heartbeat가 없으면 자동 해제됩니다.")
            return False
    except (FileNotFoundError, ValueError, TypeError, json.JSONDecodeError):
        pass
    return True


def take(who, what, why=""):
    label, excl = LOCKS.get(what, (what, True))
    if not _workcenter_priority_gate(what):
        return False
    if not _sol_write_gate(who, what):
        return False
    with state_guard():
        d = _load_unlocked()
        cur = d.get(what)
        # ★ 주인 판정은 who 가 아니라 **세션**이다. 같은 'claude' 라도 다른 창이면
        #   못 빼앗는다 — 그게 이 파일이 막으려던 바로 그 사고다(둘 다 vN+1 생성).
        if cur and excl and not _is_mine(cur, who):
            if _is_dead(cur):
                print(f"i '{label}' 을 잡고 있던 세션[{cur.get('sid','?')}]이 이미 종료되어 넘겨받습니다.")
            else:
                mins = int((time.time() - cur.get("at", 0)) / 60)
                print(f"★ '{label}' 은 이미 {cur.get('who','?')} 세션[{cur.get('sid','옛형식')}] 이"
                      f" 잡고 있습니다 ({mins}분 전 · {cur.get('why','')}).")
                print("  배타 작업이라 동시에 하면 한쪽 결과가 통째로 묻힙니다.")
                print("  → 다른 일을 먼저 하거나, 상대가 끝낼 때까지 조회·분석만 하세요.")
                print(f"  (상대 세션이 죽으면 즉시, 아니면 {STALE // 60}분 뒤 자동 해제됩니다)")
                return False
        d[what] = {
            "who": who, "why": why, "at": time.time(),
            "when": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "sid": session_id(), "agent_pid": agent_pid(),
            "pid": os.getpid(), "host": socket.gethostname(),
        }
        _save_unlocked(d)
    print(f"'{label}' 점유 — {who}[{session_id()}]" + (f" · {why}" if why else ""))
    return True


def free(who, what, force=False):
    """내 세션 것만 놓는다. 남의 것을 놓으면 그쪽이 원장을 쓰는 중에 문이 열린다."""
    with state_guard():
        d = _load_unlocked()
        cur = d.get(what)
        if cur and not _is_mine(cur, who) and not force:
            if _is_dead(cur):
                print(f"i '{what}' 주인 세션[{cur.get('sid','?')}]이 종료되어 정리합니다.")
            else:
                print(f"★ '{what}' 은 {cur.get('who','?')} 세션[{cur.get('sid','옛형식')}] 것이라"
                      " 놓을 수 없습니다. (정말 풀어야 하면 --force)")
                return False
        d.pop(what, None)
        _save_unlocked(d)
    print(f"'{what}' 놓음 — {who}[{session_id()}]")
    return True


def main():
    a = sys.argv[1:]
    get = lambda f: (a[a.index(f) + 1] if f in a and len(a) > a.index(f) + 1 else None)
    who = get("--who") or "unknown"
    if "--take" in a:
        sys.exit(0 if take(who, get("--take"), get("--why") or "") else 2)
    if "--whoami" in a:
        print(f"세션 {session_id()} · who={who} · agent_pid={agent_pid()} · {socket.gethostname()}")
        return
    if "--free" in a:
        free(who, get("--free"), force="--force" in a)
        return
    if "--free-all" in a:
        # ★ **내 세션 것만** 놓는다. 예전에는 who 만 봐서, 옆 창의 Claude 가 잡아 둔
        #   원장 점유까지 풀어 버렸다(PreCompact 자동 마무리가 특히 위험했다).
        force = "--force" in a
        with state_guard():
            d = _load_unlocked()
            mine = [k for k, v in d.items() if force or _is_mine(v, who) or _is_dead(v)]
            others = [k for k in d if k not in mine]
            for k in mine:
                d.pop(k)
            _save_unlocked(d)
        print(f"{who}[{session_id()}] 의 점유 {len(mine)}건을 놓았습니다."
              + (f" 다른 세션 것 {len(others)}건은 그대로 둡니다." if others else ""))
        return
    show()
    print("\n잡기:  python ai_claim.py --who claude --take ledger --why \"이유\"")
    print("놓기:  python ai_claim.py --who claude --free ledger")
    print("내 세션: python ai_claim.py --who claude --whoami")


if __name__ == "__main__":
    main()
