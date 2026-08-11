# -*- coding: utf-8 -*-
"""
session_wrapup.py — 세션이 끊기기 직전에 남길 것을 자동으로 남긴다 (2026-08-05 지시)

사용자 지시: "세션이 컨텍스트 윈도우가 90% 도달시 자동으로 정리하고 /compact 명령 실행"

왜 필요한가
  컨텍스트가 차면 대화는 요약(compact)되고 **그때까지 대화에만 있던 것은 사라진다.**
  지금까지는 사람이 "이제 인계해"라고 말해 줘야 했고, 말해 주기 전에 차 버리면
  큐·점유·미커밋이 그대로 남았다. 이제 compact 직전에 이 파일이 자동으로 돈다
  (`.claude/settings.json` 의 PreCompact 훅). 사람이 기억할 일이 아니다.

무엇을 하나 — CLAUDE.md "컨텍스트가 차 가면" 4단계를 그대로, 그 순서로
  1. 입력 큐를 DB 로 넘긴다        (엑셀 쓰기는 11:00·15:00 회차만 — 여기서 열지 않는다)
  2. 점유를 놓는다                 (죽은 세션이 원장·밴드를 붙들고 있지 않게)
  3. 커밋한다                      (푸시는 실패해도 커밋은 남는다)
  4. 19시트 인수인계 예약 + 세션인계 스냅샷

원칙
  · **절대 실패하지 않는다.** 어느 단계가 깨져도 다음 단계로 넘어가고 항상 exit 0 이다.
    인계를 남기려다 compact 자체를 막으면 더 나쁘다.
  · **엑셀을 열지 않는다.** 반영 회차(11:00·15:00)는 그대로 지킨다.
  · 여러 번 돌아도 안전하다(같은 인계 행은 ledger_db 가 중복을 막는다).

쓰는 법
  python ecount/session_wrapup.py                 # 사람이 직접
  python ecount/session_wrapup.py --who codex     # Codex 세션
  python ecount/session_wrapup.py --reason auto-compact --quiet
"""
import os
import sys
import json
import time
import glob
import argparse
import subprocess
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
LOG_PATH = os.path.join(REPORT_DIR, "세션마무리_기록.json")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def run(args, timeout=900):
    """자식 프로세스 한 번. 무슨 일이 있어도 예외를 밖으로 내보내지 않는다."""
    try:
        r = subprocess.run(args, cwd=ROOT, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout)
        tail = [x.strip() for x in ((r.stdout or "") + "\n" + (r.stderr or "")).splitlines()
                if x.strip()]
        return r.returncode == 0, (tail[-1] if tail else "")[:200]
    except Exception as exc:
        return False, ("%s: %s" % (type(exc).__name__, exc))[:200]


def git(*args):
    ok, out = run(["git"] + list(args), timeout=180)
    return ok, out


def git_lines(*args):
    """줄이 여러 개인 git 출력을 **다 받는다** — (ok, [줄...]).

    ⚠ `run()` 은 **마지막 한 줄만** 준다(기록용 요약이라 그렇게 만들었다). 목록을
      그것으로 읽으면 앞줄들이 **조용히 사라진다.** 2026-08-11 실측: 스테이징 목록을
      `git()` 으로 읽자 비밀값이 담긴 `leak.py` 가 빠지고 마지막 줄 `theirs.py` 만
      남아, 스캔이 그 파일을 **안 보고 통과**했다 — 막으려던 것을 새게 만드는 실수다.
      목록은 반드시 이 함수로 읽는다.
    """
    try:
        r = subprocess.run(["git"] + list(args), cwd=ROOT, capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=180)
    except Exception:
        return False, []
    return r.returncode == 0, [x.strip() for x in (r.stdout or "").splitlines() if x.strip()]


HUGE = 90 * 1024 * 1024          # GitHub 거절선은 100MB. 여유를 두고 90MB 에서 뺀다.

# 비밀값처럼 생긴 것 — `api_key = "..."` 꼴. 절대규칙 1(비밀키를 커밋에 넣지 않는다).
SECRET_SHAPE = (r"(api_?key|secret|password|passwd|token)[[:space:]]*[:=]"
                r"[[:space:]]*[\"'][A-Za-z0-9_-]{12,}")

LIVE_TRANSCRIPT_MIN = 10         # 옆 창의 대화기록이 이 분 안에 자랐으면 지금 일하는 중이다


def _unstage_huge():
    """스테이징에 올라온 거대 파일을 도로 뺀다. 뺀 경로 목록을 준다.

    2026-08-08 실사고: `db/source_index_cache.json` 이 106MB 로 자라 그대로
    커밋됐고, GitHub pre-receive 가 거절해 **저장소의 모든 푸시가 막혔다.**
    커밋 하나가 거절된 것이 아니라 그 커밋을 지나야 하는 뒤의 모든 푸시가
    같이 죽는다 — 폰에서 이어받기(푸시된 것만 보인다)도 그때 같이 죽었다.

    `add -A` 는 옆 세션 파일까지 담으므로 무엇이 올라올지 미리 알 수 없다.
    그래서 커밋 **전에** 크기로 한 번 거른다. 캐시·덤프는 다시 만들면 되고,
    정말 필요한 큰 파일이라면 사람이 LFS 로 넣을 일이지 자동 커밋이 밀 것이 아니다.
    """
    try:
        r = subprocess.run(["git", "diff", "--cached", "--name-only"], cwd=ROOT,
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=120)
    except Exception:
        return []
    dropped = []
    for rel in (r.stdout or "").splitlines():
        rel = rel.strip()
        if not rel:
            continue
        try:
            if os.path.getsize(os.path.join(ROOT, rel)) <= HUGE:
                continue
        except OSError:
            continue                      # 지워진 파일 — 삭제 스테이징은 그대로 둔다
        ok, _ = git("reset", "-q", "--", rel)
        if ok:
            dropped.append(rel)
    return dropped


def step_intake(steps):
    ok, note = run([sys.executable, os.path.join(ROOT, "ledger_db.py"), "--intake"])
    steps.append({"단계": "① 입력 큐 → DB", "성공": ok, "메모": note})


def step_free_claims(who, steps):
    ok, note = run([sys.executable, os.path.join(ROOT, "ai_claim.py"),
                    "--who", who, "--free-all"], timeout=120)
    steps.append({"단계": "② 점유 해제", "성공": ok, "메모": note})


def _live_sibling_sessions(minutes=LIVE_TRANSCRIPT_MIN):
    """살아 있는 **다른 Claude 창**을 증거로 센다. 세션 식별자 앞 8자리 목록을 준다.

    ★ 점유판만 보면 못 본다 (사고 #36 — 2026-08-11 하루에 세 번 실측).
      `ai_claim` 은 점유를 적은 **단발 pid**(`python ai_claim.py` 한 번)로 생사를
      판정하므로, 옆 세션이 멀쩡히 파일을 저장하는 중에도 "지금 잡혀 있는 작업 없음"
      이라 말한다. 점유판은 '무엇을 하겠다고 적었나'이고 여기는 '지금 무엇이
      벌어지나'다. 실측 21:07 — 점유판 0건 · 옆 창 셋이 1·15·52초 전에 자라는 중.

    ★ 미커밋 파일 mtime 으로 세지 않는 이유: **내가 방금 고친 것까지** 옆 세션으로
      읽혀 언제나 보류가 된다. 늘 걸리는 경보는 아무도 안 본다. 대화기록은 파일
      이름이 곧 세션 식별자라 **내 것을 이름으로 뺄 수 있다.**

    내가 누군지 모르면(환경변수 없음) 빈 목록이다 — 나를 옆 세션으로 세는 것보다
    점유판에 맡기는 편이 낫다. 그리고 **내 대화기록이 있는 폴더만** 본다(다른
    프로젝트의 창까지 세면 역시 늘 걸린다).
    """
    me = (os.environ.get("CLAUDE_CODE_SESSION_ID") or "").strip()
    if not me:
        return []
    home = os.path.join(os.path.expanduser("~"), ".claude", "projects")
    mine = [os.path.dirname(p) for p in glob.glob(os.path.join(home, "*", me + ".jsonl"))]
    if not mine:
        return []
    cut, live = time.time() - minutes * 60, []
    for name in os.listdir(mine[0]):
        if not name.endswith(".jsonl") or name[:-6] == me:
            continue
        try:
            if os.path.getmtime(os.path.join(mine[0], name)) >= cut:
                live.append(name[:8])
        except OSError:
            continue
    return live


def _other_live_sessions():
    """지금 살아 있는 **다른** 세션. 없으면 빈 목록.

    근거가 둘이고 층이 다르다 — **점유판**(무엇을 하겠다고 적었나) + **대화기록**
    (지금 무엇이 벌어지나). 하나라도 걸리면 살아 있다고 본다. 이 판정이 하는 일은
    푸시 보류뿐이라 헛짚어도 잃는 것이 없고, 못 보면 남의 반쯤 고친 코드가 원격으로 간다.
    """
    out = []
    try:
        import ai_claim
        out = [v for v in (ai_claim.load() or {}).values()
               if not ai_claim._is_mine(v, v.get("who")) and not ai_claim._is_dead(v)]
    except Exception:
        out = []
    out += [{"작업": "다른 창", "who": "claude", "sid": sid}
            for sid in _live_sibling_sessions()]
    return out


def step_commit(who, reason, steps):
    """미커밋이 있으면 커밋한다. 없으면 아무것도 만들지 않는다.

    비밀값 스캔은 커밋 **전에** 한다 — 절대규칙 1. 걸리면 커밋하지 않고 사람에게 넘긴다.
    단 훑는 범위는 **이번 커밋이 담는 경로**이고, 걸리면 **어디가 걸렸는지 적고**
    **인덱스를 담기 전 상태로 되돌린다**(2026-08-11 실사고 — 아래 주석).

    ★ 다른 세션이 살아 있으면 **푸시하지 않는다** (2026-08-05 실사고).
      작업 폴더는 세션끼리 공유하므로 `git add -A` 는 옆 세션이 편집 중인 파일까지
      담는다. 커밋은 해도 잃는 것이 없지만(파일은 디스크에 그대로다), 반쯤 고친 남의
      코드를 원격 master 로 밀면 그때부터는 남이 치운다. 그래서 로컬까지만 남긴다.
    """
    ok, dirty = git("status", "--porcelain")
    if not ok:
        steps.append({"단계": "③ 커밋", "성공": False, "메모": "git 상태를 읽지 못했다"})
        return
    if not dirty.strip():
        steps.append({"단계": "③ 커밋", "성공": True, "메모": "미커밋 없음 — 만들지 않았다"})
        return

    okp, pre_staged = git_lines("-c", "core.quotepath=false", "diff", "--cached", "--name-only")
    if not okp:
        pre_staged = ["(못 읽음)"]        # 못 읽었으면 '있다'고 본다 — 남의 것을 지우는 쪽으로 기울지 않는다

    def _rollback():
        """`add -A` 로 담은 것을 되돌린다 — 담기 전 인덱스가 **비어 있었을 때만**.

        ★ 멈추면서 스테이징을 남기면 **기계가 사고 #36 을 만든다** (2026-08-11 실사고).
          그날 이 단계가 14번 멈췄고, 그때마다 옆 세션이 편집 중인 파일이 공유 인덱스에
          담긴 채 남았다 — 다음 사람의 `git commit -m` 한 번이 그것을 통째로 커밋한다.
          파일은 디스크에 그대로라 아무 화면에도 티가 안 난다.
        원래 남의 스테이징이 있었으면 되돌리지 않는다 — 그것은 남이 적어 둔 뜻이다.
        """
        if pre_staged:
            return " · 원래 있던 스테이징 %d개라 인덱스는 그대로 뒀다(커밋 전 git status 를 볼 것)" \
                   % len(pre_staged)
        git("reset", "-q")
        return " · 인덱스는 담기 전(비어 있던) 상태로 되돌렸다"

    git("add", "-A")
    huge = _unstage_huge()               # 90MB 넘는 것은 커밋 전에 뺀다 — 푸시가 통째로 막힌다
    # 비밀값 스캔은 **이번 커밋이 담는 경로만** 본다 (2026-08-11 실사고).
    #   예전에는 인덱스 **전체**를 훑었다. 그래서 어제 커밋된 `canonical_sync.py` 의
    #   `completion_token = "canonical-…" + sha256_json(...)`(비밀값이 아니라
    #   멱등키를 만드는 계산식) 한 줄에 걸려 **그날 회차 14번이 전부** 여기서 멈췄다.
    #   커밋을 거부해도 이미 커밋된 줄은 사라지지 않으니 **영구히 잠긴 관문**이었고,
    #   자동 커밋이 하루 종일 한 건도 안 됐는데 화면에는 '4/5 단계 완료'만 떴다.
    #   막아야 하는 것은 '이 커밋이 새로 담는 것'이다. 새 파일은 전체가 담기므로 그대로 걸린다.
    okl, paths = git_lines("-c", "core.quotepath=false", "diff", "--cached",
                           "--name-only", "--diff-filter=ACMRT")
    # 목록을 못 읽으면 예전처럼 인덱스 전체를 훑는다 — 빠뜨리는 쪽으로 기울지 않는다.
    chunks = [paths[i:i + 120] for i in range(0, len(paths), 120)] if okl else [[]]
    for chunk in chunks:
        leaked, hits = git("grep", "--cached", "-nEI", SECRET_SHAPE,
                           *(["--"] + chunk if chunk else []))
        if leaked and hits.strip():      # grep 은 '찾으면' 0 을 준다 — 찾은 것이 사고다
            # 걸린 자리를 **적어 준다**. 어디인지 말해 주지 않으면 사람이 확인할 수 없어
            # 경보가 그대로 쌓인다(오늘 14번이 그랬다).
            steps.append({"단계": "③ 커밋", "성공": False,
                          "메모": "비밀값 형태가 이번 커밋 대상에 있어 커밋을 멈췄다 — %s · 사람이 확인할 것%s"
                                  % (hits.strip()[:110], _rollback())})
            return
    msg = ("세션 자동 마무리(%s) — %s\n\n"
           "컨텍스트가 차서 대화가 요약되기 전에 남긴 자동 커밋이다.\n"
           "내용 확인은 reports/세션인계.md 와 19_AI작업인수인계 시트를 볼 것."
           % (who, reason))
    others = _other_live_sessions()
    if others:
        # 무엇을 근거로 '살아 있다'고 봤는지 커밋에 적어 남긴다. 근거를 안 적으면
        # 나중에 이 커밋을 보는 사람이 왜 안 밀렸는지 알 수 없다.
        msg += ("\n\n다른 세션이 같은 작업 폴더에서 일하는 중이라 이 커밋에는 그쪽 파일이"
                "\n섞여 있을 수 있다. 그래서 **푸시하지 않았다** — 확인 후 사람이 밀 것."
                "\n살아 있다고 본 근거: %s"
                % ", ".join(str(o.get("sid") or o.get("작업") or "?")[:8] for o in others[:6]))
    if huge:
        msg += ("\n\n90MB 를 넘어 커밋에서 뺀 파일: %s"
                "\n(GitHub 100MB 한도 — 담았다면 이 저장소의 푸시가 전부 막힌다)"
                % ", ".join(huge))
    ok, note = git("commit", "-q", "-m", msg)
    tail = (" · 거대파일 %d개 제외(%s)" % (len(huge), ", ".join(huge))) if huge else ""
    if not ok:
        # 커밋이 깨진 자리도 스테이징을 남기지 않는다 — 남기면 다음 사람이 통째로 커밋한다.
        steps.append({"단계": "③ 커밋", "성공": False,
                      "메모": (note or "커밋 실패") + tail + _rollback()})
        return
    if others:
        steps.append({"단계": "③ 커밋", "성공": True,
                      "메모": ("커밋 완료 · 다른 세션 %d개가 작업 중이라 푸시는 보류"
                               % len(others)) + tail})
        return
    pushed, why = git("push", "-q")
    steps.append({"단계": "③ 커밋", "성공": True,
                  "메모": "커밋 완료"
                          + (" · 푸시 완료" if pushed
                             else " · 푸시 실패(커밋은 남음): %s" % (why or "이유 불명"))
                          + tail})


def step_handoff(who, reason, steps):
    """19시트 인수인계 예약 + 세션인계 스냅샷.

    엑셀은 여기서 열지 않는다 — 예약만 하고 11:00·15:00 회차가 기록한다.
    """
    ok, head = git("rev-parse", "--short", "HEAD")
    title = "세션 자동 마무리(%s)" % who
    detail = ("컨텍스트 한도로 대화가 요약되기 전 자동 인계. 계기=%s · 기준커밋 %s · %s. "
              "입력 큐는 DB 로 넘겼고 점유는 해제했다. 재개 지점은 reports/세션인계.md."
              % (reason, head or "?", datetime.now().strftime("%Y-%m-%d %H:%M")))
    # --supersede: 이 줄은 컨텍스트가 찰 때마다 다시 만들어진다. 상세에 시각·기준커밋이
    #   들어가 매번 다른 줄이 되므로 중복 인덱스가 못 걸렀고, 실측 하루 44줄이 쌓여
    #   19시트(사람이 읽는 원장)를 덮을 참이었다. 마지막 하나만 남긴다.
    ok1, note1 = run([sys.executable, os.path.join(ROOT, "ledger_db.py"),
                      "--handoff", "--supersede", "--b", title, "--c", detail])
    steps.append({"단계": "④ 19시트 인수인계 예약", "성공": ok1, "메모": note1})
    ok2, note2 = run([sys.executable, os.path.join(ROOT, "session_handoff.py"), "--snapshot"])
    steps.append({"단계": "④ 세션인계 스냅샷", "성공": ok2, "메모": note2})


def wrapup(who="claude", reason="manual"):
    steps = []
    step_intake(steps)
    step_free_claims(who, steps)
    step_commit(who, reason, steps)
    step_handoff(who, reason, steps)
    record = {"시각": datetime.now().isoformat(timespec="seconds"),
              "누가": who, "계기": reason, "단계": steps}
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        history = []
        if os.path.exists(LOG_PATH):
            try:
                history = json.load(open(LOG_PATH, encoding="utf-8")) or []
            except Exception:
                history = []
        history = ([record] + list(history))[:40]
        tmp = LOG_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(history, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
        os.replace(tmp, LOG_PATH)
    except Exception:
        pass
    return record


def main():
    ap = argparse.ArgumentParser(description="세션이 끊기기 전 인계를 자동으로 남긴다")
    ap.add_argument("--who", default="claude", help="claude | codex")
    ap.add_argument("--reason", default="manual", help="auto-compact · manual-compact · manual")
    ap.add_argument("--quiet", action="store_true", help="한 줄 요약만 (훅에서 쓴다)")
    a = ap.parse_args()

    # 훅은 stdin 으로 JSON 을 준다. 없어도 되고, 있으면 계기를 더 정확히 적는다.
    reason = a.reason
    if not sys.stdin.isatty():
        try:
            payload = json.loads(sys.stdin.read() or "{}")
            trigger = payload.get("trigger") or payload.get("matcher")
            if trigger:
                reason = "%s-compact" % trigger
        except Exception:
            pass

    record = wrapup(a.who, reason)
    good = sum(1 for s in record["단계"] if s["성공"])
    line = "세션 자동 마무리: %d/%d 단계 완료 (%s)" % (good, len(record["단계"]), reason)
    if a.quiet:
        # PreCompact 훅은 stdout 의 JSON 으로 사용자에게 한 줄을 띄운다.
        print(json.dumps({"systemMessage": line + " · 상세 reports/세션마무리_기록.json"},
                         ensure_ascii=False))
    else:
        print(line)
        for s in record["단계"]:
            print(("  OK " if s["성공"] else "  X  ") + s["단계"] +
                  (" | " + s["메모"] if s["메모"] else ""))
    return 0        # 인계를 남기려다 compact 를 막지 않는다 — 언제나 성공으로 끝낸다


if __name__ == "__main__":
    sys.exit(main())
