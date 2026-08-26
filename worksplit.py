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

주인은 **세션**이다 (사람 이름이 아니다)
  `ai_claim.session_id()` 를 그대로 쓴다 — 같은 폴더에 Claude 창이 둘 떠 있어도
  둘 다 "claude" 라서 서로 밟던 문제를 그쪽에서 이미 풀어 두었다.
  ★ 프로세스 PID 로 주인을 표시하지 않는다. 이 스크립트는 명령 한 번에 끝나는
    프로세스라, PID 를 적으면 **적자마자 죽은 주인**이 된다(2026-08-05 실제로 그랬다).
  주인이 오래(기본 8시간) 손대지 않은 '진행' 은 주인 없음으로 보고 다른 세션이 가져간다.

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
# ★ 분담판도 **본체 하나**를 함께 본다 (2026-08-06). ai_claim 과 같은 이유다 —
#   이건 "누가 어떤 일을 맡았나" 라는 **세션 사이의 약속**이라, 체크아웃마다 따로
#   있으면 서로의 분담을 못 본다. 워크트리에서 일을 맡아 적어도 본체 세션에는
#   안 보이고, 결국 둘이 같은 일을 몇 시간 조사한 뒤에야 겹친 걸 알게 된다
#   (worksplit 이 애초에 막으려던 바로 그 사고다).
#   본체에서는 `shared()` 가 제 폴더를 돌려주므로 동작이 하나도 바뀌지 않는다.
try:
    from worktree_state import shared as _shared
    _BOARD_DIR = _shared("reports")
except Exception:
    _BOARD_DIR = os.path.join(ROOT, "reports")
BOARD = os.path.join(_BOARD_DIR, "worksplit.json")
GUARD = os.path.join(_BOARD_DIR, ".worksplit.guard")
OUT_MD = os.path.join(_BOARD_DIR, "작업분담.md")

WAIT, DOING, DONE, HOLD = "대기", "진행", "완료", "사람대기"
# 이 일을 하려면 어떤 자원 잠금이 필요한가(ai_claim 의 이름과 같게 쓴다).
LOCK_LABEL = {"ledger": "관리대장", "band": "밴드", "publish": "게시", "code": "코드",
              "erp": "ERP 화면", "": "없음"}

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


ORPHAN_SEC = 8 * 3600      # 주인이 이만큼 손대지 않은 '진행' 은 주인 없음으로 본다


def session_id():
    """이 세션의 식별자. ai_claim 과 **같은 값**을 쓴다(둘이 다르면 조율이 어긋난다)."""
    try:
        import ai_claim
        return ai_claim.session_id()
    except Exception:
        return "manual"


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
    # 옛 항목에는 at_ts 가 없다. 없으면 나이를 못 재서 **영원히 주인 있는 일**이 된다 —
    # 표시용 at 문자열에서 되살린다.
    for it in d["items"]:
        if not it.get("at_ts"):
            try:
                it["at_ts"] = datetime.strptime(it.get("at", ""), "%Y-%m-%d %H:%M").timestamp()
            except (ValueError, TypeError):
                it["at_ts"] = 0
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


def _owner_state(it, me=""):
    """(표시할 주인, 내가 가져가도 되나).

    ★ sid 가 없는 **옛 기록** 처리(2026-08-05): sid 를 적기 전에 잡아 둔 항목은
      주인을 확인할 길이 없어, `_owner_state` 가 8시간이 지날 때까지 아무에게도
      내주지 않았다 — 잡은 본인조차 완료 처리를 못 했다(실제로 [2] 가 그랬다).
      ai_claim 과 같은 과도기 규칙을 쓴다: **sid 가 없으면 who 가 같은 쪽을 주인으로
      본다.** 새로 잡는 순간 sid 가 붙으므로 이 예외는 옛 항목에만 한 번 쓰인다.
    """
    who = it.get("who") or ""
    if it.get("state") != DOING:
        return who, it.get("state") in (WAIT, HOLD)
    if it.get("sid"):
        if it["sid"] == session_id():
            return f"{who}(나)", True              # 내 세션이 잡은 것 — 이어서 하면 된다
    elif me and who == me:
        return f"{who}(옛 기록)", True             # sid 없는 과도기 항목
    age = time.time() - float(it.get("at_ts") or 0)
    if it.get("at_ts") and age > ORPHAN_SEC:
        return f"{who}({int(age // 3600)}시간째 소식 없음)", True
    return who, False


def _next_id(d):
    """번호는 `seq` 와 **실제 쓰인 번호** 중 큰 것에서 나온다.

    ★ 2026-08-26 실사고: 옆 창이 판을 **직접 써서** `id 259` 를 만들고 `seq` 는
      258 로 두자, 여기서 `seq + 1` 이 또 259 가 되어 **같은 번호가 둘**이 됐다.
      그러면 `--done 259` 가 **먼저 걸린 남의 항목**을 완료로 찍는다(실제로 그랬고,
      하마터면 옆 창이 아직 하는 일을 끝난 것으로 남길 뻔했다 · [104]).
    ★ 판이 깨끗하면 예전과 **같은 번호**가 나온다 — 넓히는 것이 아니다([172]).
    """
    used = [int(x.get("id") or 0) for x in (d.get("items") or [])]
    return max([int(d.get("seq") or 0)] + used) + 1


def _dup(d, wid):
    """그 번호가 여럿이면 **아무것도 안 한다**([169]·[172]).

    어느 것인지 원본이 안 말해 주므로 고르면 남의 일을 건드린다 —
    **못 고치는 것보다 나쁘다.** 사람이 번호를 갈라 준 뒤에 한다.
    """
    n = sum(1 for x in (d.get("items") or []) if x.get("id") == wid)
    if n > 1:
        print(f"★ [{wid}] 이 {n}개입니다 — 어느 것인지 몰라 손대지 않습니다."
              " (판을 직접 쓴 창이 있으면 번호가 겹칩니다)")
        return True
    return False


def add(title, detail="", lock="", human=False, who=""):
    with guard():
        d = _read()
        d["seq"] = _next_id(d)
        it = {"id": d["seq"], "title": title, "detail": detail, "lock": lock or "",
              "state": HOLD if human else WAIT, "who": "", "sid": "",
              "at": datetime.now().strftime("%Y-%m-%d %H:%M"), "at_ts": time.time(),
              "note": "", "by": who or ""}
        d["items"].append(it)
        _write(d)
    print(f"[{it['id']}] 추가 — {title}" + (" (사람이 해야 함)" if human else ""))
    return it["id"]


def take(who, wid):
    with guard():
        d = _read()
        if _dup(d, wid):
            return False
        for it in d["items"]:
            if it["id"] != wid:
                continue
            if it["state"] == DONE:
                print(f"[{wid}] 은 이미 완료된 일입니다.")
                return False
            shown, free_ = _owner_state(it, who)
            if it["state"] == DOING and not free_:
                print(f"★ [{wid}] {it['title']} 은 이미 {shown} 가 맡았습니다.")
                print("  → 다른 일을 고르세요:  python worksplit.py --free")
                return False
            it.update(state=DOING, who=who, sid=session_id(),
                      host=socket.gethostname(),
                      at=datetime.now().strftime("%Y-%m-%d %H:%M"), at_ts=time.time())
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
        if _dup(d, wid):
            return False
        for it in d["items"]:
            if it["id"] != wid:
                continue
            if it["state"] == DOING:
                shown, free_ = _owner_state(it, who)
                if not free_:
                    print(f"★ [{wid}] 은 {shown} 것이라 손대지 않습니다.")
                    return False
            it.update(state=state, note=note or it.get("note", ""),
                      at=datetime.now().strftime("%Y-%m-%d %H:%M"), at_ts=time.time())
            if state == WAIT:
                it["who"] = ""
                it["sid"] = ""
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
        # 아무도 안 맡은 것 + 주인이 오래 소식 없는 것. 내가 이미 맡은 것은 빼고 보여 준다.
        rows = [x for x in rows
                if x["state"] == WAIT
                or (x["state"] == DOING and _owner_state(x)[1] and x.get("sid") != session_id())]
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
         "## ★ 커밋할 때 — `git add -A` 를 쓰지 않는다",
         "",
         "2026-08-05 실제로 겪은 일: 한쪽이 `git add -A` 로 올려 둔 사이 다른 쪽이 커밋해",
         "**남의 변경이 남의 커밋 메시지 아래로 들어갔다**(내용은 안 없어졌지만 이력이 섞였다).",
         "동시 작업 중에는 **내가 고친 파일만 이름으로** 올린다.",
         "",
         "```bash",
         "git add ecount/worksplit.py ecount/webapp/index.html   # 내가 고친 것만",
         "git commit -m \"...\"                                    # add -A · commit -a 금지",
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
