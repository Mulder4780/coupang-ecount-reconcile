# -*- coding: utf-8 -*-
"""세워 둔 일이 풀렸으면 스스로 알리고(없으면 AI 에게 넘기고), 보류된 푸시는 조용해지면 민다.

사용자 지시(2026-08-11): **"이 세션도 완전 자동화시켜"**

★ 빠져 있던 자리는 '일을 하는 쪽'이 아니라 **'막힘이 풀렸다는 사실을 아는 쪽'** 이었다.
  이날 실측 둘이 그것이다.
  · 분담판 `[34]`(사고 #38 의 검증 번호 달기)는 옆 세션이 `code` 를 잡고 있어 세워 뒀다.
    그 세션이 점유를 놓은 뒤에도 **어느 화면도 "이제 된다"고 말하지 않아서** 사람이
    "하던 작업 진행" 을 두 번 쳤다. 기계는 그때 이미 점유판을 보고 있었다.
  · 자동 마무리는 옆 세션이 살아 있으면 푸시를 **보류**한다(`[104]` — 그것은 옳다).
    그런데 그 세션이 사라진 뒤에는 **미는 사람이 없다.** 폰·웹(claude.ai/code)은
    **푸시된 것만** 보므로, 그동안 폰에서는 없는 코드다.
  둘 다 **판단 근거를 기계가 이미 다 갖고 있었다.** 그래서 새 판단을 만들지 않고
  이미 있는 판단을 제 시각에 **불러 주는 것**이 이 파일의 전부다.

지키는 것
- **판단을 두 곳에 만들지 않는다.** 살아 있는 세션은 `session_wrapup._other_live_sessions()`,
  점유 생사는 `ai_claim._is_dead`, 주인 판정은 `worksplit._owner_state`, 비밀값 형태는
  `handoff_review.secret_findings` 를 그대로 빌린다. 여기서 다시 세면 언젠가 갈리고,
  갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다.
- **아무도 없을 때만 AI 에게 넘긴다.** 사람과 대화 중인 세션이 있으면 그 세션이 할 일이다 —
  거기에 AI 를 하나 더 붙이면 같은 파일을 둘이 고친다(사고 #36 · `[104]`). 그때는
  인계 문서에 올리는 것으로 끝낸다 — 공짜고, 볼 사람이 있다.
- **한 항목에 표는 하나.** 회차마다 새 표를 만들면 크레딧이 새고 큐가 쓰레기가 된다.
- **강제 푸시는 없다.** 거절되면 그대로 적고 다음 회차로 넘긴다. 못 읽은 것은
  '괜찮다'로 치지 않는다 — 범위를 못 읽으면 밀지 않는다.
- **읽기 전용 판단 + 되돌릴 수 있는 행동만.** 엑셀·원장·밴드는 건드리지 않는다.

돌리는 법:  python worksplit_auto.py            # 지금 무엇이 풀렸나(안 함)
            python worksplit_auto.py --run      # 회차(워치독 30분이 이걸 부른다)
"""
import json
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import ai_claim                                                    # noqa: E402
import worksplit                                                   # noqa: E402

try:
    from worktree_state import shared as _shared
    _DIR = _shared("reports")
except Exception:
    _DIR = os.path.join(ROOT, "reports")

STATE = os.path.join(_DIR, "세션자동화_상태.json")
# ★ AI 에게 넘겨도 되는 자원은 **코드뿐**이다. 관리대장·밴드·게시는 사람 인증이나
#   비가역 승인이 걸려 있고(`[190]`), ERP 화면은 사람만 볼 수 있다. 그 갈래는 알리는
#   것까지가 기계 몫이다 — 넘길 수 없는 것을 넘기면 표만 쌓이고 아무 일도 안 된다.
AI_LOCKS = {"code"}


def _live_min():
    """'지금 일하는 중'의 기준 분. `session_wrapup` 이 정한 값을 그대로 쓴다."""
    try:
        import session_wrapup
        return int(session_wrapup.LIVE_TRANSCRIPT_MIN)
    except Exception:
        return 10


def _lock_en(lock):
    """분담판의 자원 이름을 `ai_claim` 열쇠(영문)로 맞춘다.

    ★ 한글이 그대로 저장된다 — `--lock 코드` 로 만들면 항목에 `"코드"` 가 들어간다
      (실측 `[34]`). 영문만 비교하면 그 항목은 **'자원 없음'이 되어 막힘을 못 보고
      늘 "가능"이라 답한다.** 오류는 안 난다 — `[165]` 와 같은 모양의 조용한 잘못이다.
    """
    s = (lock or "").strip()
    if not s or s in worksplit.LOCK_LABEL:
        return s
    for en, ko in worksplit.LOCK_LABEL.items():
        if ko == s:
            return en
    return s


def _claim_blocker(lock_en, claims=None):
    """그 자원을 **살아 있는 다른 세션**이 잡고 있나 — 잡고 있으면 그 기록을 돌려준다."""
    if not lock_en:
        return None
    c = claims if claims is not None else (ai_claim.load() or {})
    v = c.get(lock_en)
    if not v or ai_claim._is_dead(v):
        return None
    return v


def parked():
    """대기·유기된 일마다 **지금 되나 / 왜 안 되나**를 근거와 함께 돌려준다.

    주인 판정은 `worksplit._owner_state` 가 한다 — 죽은 세션이 '진행'으로 붙잡고 있는
    항목까지 그 함수가 이미 가려 준다. 여기서 다시 세면 8시간 규칙이 두 벌이 된다.
    """
    out = []
    try:
        d = worksplit.load()
    except Exception as exc:
        return [{"id": 0, "title": "분담판을 못 읽었다", "가능": False,
                 "사유": "%s: %s" % (type(exc).__name__, str(exc)[:80])}]
    claims = ai_claim.load() or {}
    for it in d.get("items") or []:
        state = it.get("state")
        if state in (worksplit.DONE, worksplit.HOLD):
            continue                      # 완료 · 사람대기(사람만 할 수 있는 일)는 여기 몫이 아니다
        try:
            _who, takeable = worksplit._owner_state(it)
        except Exception:
            takeable = state == worksplit.WAIT
        if not takeable:
            continue                      # 살아 있는 주인이 지금 하고 있다
        lock = _lock_en(it.get("lock"))
        row = {"id": it.get("id"), "title": it.get("title") or "", "자원": lock,
               "detail": (it.get("detail") or "")[:400]}
        held = _claim_blocker(lock, claims)
        if held:
            row.update({"가능": False, "사유": "자원 '%s' 을 %s[%s] 가 잡고 있다 — %s"
                                              % (worksplit.LOCK_LABEL.get(lock, lock),
                                                 held.get("who") or "?", held.get("sid") or "?",
                                                 (held.get("why") or "")[:60])})
        else:
            row.update({"가능": True, "사유": "막고 있던 것이 없다"})
        out.append(row)
    return out


def live_others():
    """지금 일하고 있는 **다른** 세션. 판단은 `session_wrapup` 한 곳에서 빌린다."""
    sids = []
    try:
        import session_wrapup
        for r in session_wrapup._other_live_sessions() or []:
            sids.append(str(r.get("sid") or "?")[:8])
        # ★ 기계 회차(워치독)에는 `CLAUDE_CODE_SESSION_ID` 가 없어 위 함수의 대화기록
        #   근거가 **조용히 빈손**으로 돌아온다 — 그러면 점유를 안 잡은 옆 창을 못 본다.
        #   그래서 대화기록을 한 번 더 직접 묻는다. 세션 안에서 부를 때는 **내 sid 를
        #   빼야** 한다(안 빼면 내가 나를 옆 세션으로 세어 영원히 보류다).
        me = (os.environ.get("CLAUDE_CODE_SESSION_ID") or "").strip()
        for sid in session_wrapup.live_transcripts(exclude=me):
            if sid not in sids:
                sids.append(sid)
    except Exception:
        pass
    return {"수": len(sids), "목록": sids, "기준분": _live_min()}


def _git(*a, **kw):
    try:
        from proc_guard import run_tree
        r = run_tree(["git", *a], cwd=ROOT, timeout=int(kw.get("timeout") or 180),
                     drain_timeout=10)
        return r.returncode == 0, (r.stdout or "")
    except Exception as exc:
        return False, "%s: %s" % (type(exc).__name__, str(exc)[:120])


def push_state():
    """보류된 푸시가 있나 · 지금 밀어도 되나."""
    ok, up = _git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{upstream}")
    up = (up or "").strip().splitlines()[-1].strip() if ok and (up or "").strip() else ""
    if not up:
        return {"미푸시": 0, "가능": False, "사유": "upstream 이 없다 — 브랜치를 아직 안 밀었다"}
    okg, gd = _git("rev-parse", "--git-dir")
    gitdir = os.path.join(ROOT, (gd or "").strip() or ".git")
    for mark in ("rebase-merge", "rebase-apply", "MERGE_HEAD", "CHERRY_PICK_HEAD"):
        if okg and os.path.exists(os.path.join(gitdir, mark)):
            return {"미푸시": 0, "가능": False,
                    "사유": "%s 가 진행 중이다 — 기계가 끼어들지 않는다" % mark}
    okc, n = _git("rev-list", "--count", up + "..HEAD")
    try:
        ahead = int((n or "0").strip() or 0)
    except ValueError:
        ahead = 0
    if not okc:
        return {"미푸시": 0, "가능": False, "사유": "미푸시 개수를 못 읽었다"}
    if ahead <= 0:
        return {"미푸시": 0, "가능": False, "사유": "밀 것이 없다", "upstream": up}
    # ★ 범위를 **못 읽으면 밀지 않는다.** '못 읽음'을 '깨끗함'으로 치는 것이
    #   이 프로젝트에서 제일 위험한 종류의 잘못이다(`[169]`).
    okd, patch = _git("diff", up + "..HEAD")
    if not okd:
        return {"미푸시": ahead, "가능": False, "사유": "미푸시 범위를 못 읽어 비밀값 검사를 못 했다",
                "upstream": up}
    try:
        import handoff_review
        bad = handoff_review.secret_findings(patch)
    except Exception as exc:
        return {"미푸시": ahead, "가능": False,
                "사유": "비밀값 검사기를 못 불렀다(%s)" % type(exc).__name__, "upstream": up}
    if bad:
        return {"미푸시": ahead, "가능": False, "upstream": up,
                "사유": "미푸시 범위에 비밀값 형태가 있다 — %s · 사람이 확인할 것"
                        % ", ".join(bad[:3])}
    return {"미푸시": ahead, "가능": True, "사유": "밀 수 있다", "upstream": up}


def _load_state():
    try:
        with open(STATE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save(doc):
    tmp = STATE + ".tmp"
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STATE)


def _hand_to_ai(row, tickets):
    """아무도 없을 때 한 항목을 AI 에게 넘긴다. 표는 항목당 하나뿐이다."""
    key = str(row.get("id"))
    if tickets.get(key):
        return ""
    from agent_dispatch import dispatch_async, enqueue
    ticket = enqueue(
        "worksplit-" + key,
        "세워 둔 일 [%s] 의 막힘이 풀렸다 — %s" % (key, row.get("title") or ""),
        ["worksplit.py", "--who", "claude", "--mine"],
        extra=("이 표는 실패한 명령이 아니라 **분담판에 세워 둔 일**입니다. "
               "로컬에서 실행된 업무 스크립트는 없습니다.\n"
               "상세: %s\n"
               "먼저 `python ai_claim.py --who claude --take %s --why \"...\"` 로 자원을 잡고, "
               "`python worksplit.py --take %s --who claude` 로 맡은 뒤 진행하세요. "
               "끝나면 `python worksplit.py --done %s --who claude --note \"결과\"`."
               % (row.get("detail") or "(상세 없음)", row.get("자원") or "code", key, key)))
    dispatch_async(ticket, local_returncode=0)
    tid = str(ticket.get("id") or "queued")
    tickets[key] = {"표": tid, "시각": datetime.now().strftime("%Y-%m-%d %H:%M")}
    return tid


def run(dry=False, ai=True):
    """회차 한 번. 워치독(30분)·daily_run 이 부른다. 요약 한 줄을 `한줄` 로 돌려준다."""
    st = _load_state()
    tickets = dict(st.get("표") or {})
    rows, others, ps = parked(), live_others(), push_state()
    ready = [r for r in rows if r.get("가능")]
    acts = []

    # ① 푸시 — 아무도 없을 때만. 보류는 옆 세션이 살아 있는 동안의 규칙이고,
    #    그 세션이 사라지면 보류를 이어 갈 이유가 없다(그때부터는 그냥 안 밀린 코드다).
    if ps.get("가능") and others["수"] == 0 and not dry:
        okp, out = _git("push")
        ps["밀었나"] = bool(okp)
        ps["결과"] = (out or "").strip()[-160:]
        acts.append("푸시 %d개 %s" % (ps["미푸시"], "성공" if okp else "실패"))
    elif ps.get("가능") and others["수"]:
        ps["사유"] = "다른 세션 %d개가 일하는 중이라 보류 — %s" % (others["수"],
                                                                ", ".join(others["목록"][:4]))

    # ② 세워 둔 일 — 아무도 없을 때만 AI 에게 넘긴다. 사람이 있으면 인계 문서로 족하다.
    ai_on = ai and os.environ.get("COUPANG_AUTO_AI", "1") != "0"
    if ready and ai_on and others["수"] == 0 and not dry:
        for row in ready:
            if row.get("자원") not in AI_LOCKS:
                continue
            try:
                tid = _hand_to_ai(row, tickets)
            except Exception as exc:
                row["표오류"] = "%s: %s" % (type(exc).__name__, str(exc)[:100])
                break
            if tid:
                acts.append("[%s] AI 인계 %s" % (row.get("id"), tid))
                break                     # 회차당 하나 — 표가 쌓이면 아무도 안 본다
    doc = {"시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "세워둔일": rows,
           "풀린것": [r["id"] for r in ready], "살아있는세션": others, "푸시": ps,
           "표": tickets, "한행동": acts}
    doc["한줄"] = ("세션자동화: 풀린 일 %d건%s · 미푸시 %d%s"
                   % (len(ready), (" " + ",".join("[%s]" % r["id"] for r in ready[:4])) if ready else "",
                      ps.get("미푸시") or 0,
                      (" · " + " · ".join(acts)) if acts else ""))
    if not dry:
        _save(doc)
    return doc


def banner():
    """인계 문서가 읽는 것만 — 판단은 여기서 하고, 문서는 담기만 한다."""
    doc = _load_state()
    ready = [r for r in (doc.get("세워둔일") or []) if r.get("가능")]
    ps = doc.get("푸시") or {}
    return {"시각": doc.get("시각") or "", "풀린일": [
        {"id": r.get("id"), "title": r.get("title"), "자원": r.get("자원")} for r in ready],
        "미푸시": int(ps.get("미푸시") or 0), "푸시사유": ps.get("사유") or "",
        "살아있는세션": (doc.get("살아있는세션") or {}).get("수") or 0}


def main(argv=None):
    a = list(argv if argv is not None else sys.argv[1:])
    if "--json" in a:
        print(json.dumps(run(dry=True), ensure_ascii=False)[:4000])
        return 0
    doc = run(dry="--run" not in a)
    print(doc["한줄"])
    for r in doc["세워둔일"]:
        print("  [%s] %s — %s" % (r.get("id"), "가능" if r.get("가능") else "막힘", r.get("사유")))
    ps = doc["푸시"]
    print("  푸시: 미푸시 %s · %s" % (ps.get("미푸시"), ps.get("사유")))
    return 0


if __name__ == "__main__":
    sys.exit(main())
