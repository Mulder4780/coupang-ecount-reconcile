# -*- coding: utf-8 -*-
"""옮길 몸통 — 자동 마무리 커밋 단계 세 갈래 (사고 #38 · 분담판 `[33]`)

**이 파일은 임시다.** 내용은 `tests/synthetic_check.py` 에 번호를 달아 옮기고 여기서
지운다. 지금 못 옮기는 이유는 하나다 — 그 파일을 옆 세션이 편집 중이라 같은 파일
동시 편집 금지에 걸린다(사고 #36 의 교훈 그대로). 그래서 대화에 남기지 않고 파일로 둔다.

⚠ **`[221]` 이 생겼다고 이 파일을 지우면 안 된다** (2026-08-11 22:45 확인). 그 시험은 같은
분담판 항목(#33)에서 나왔지만 **보는 것이 다르다** — 옆 세션 표기·푸시 보류·문서 절차다.
비밀값 스캔을 '없음'으로 가짜 처리하므로 **거부 경로를 한 번도 지나가지 않고**, 되돌림은
`_rollback()` 이 소스에 두 번 나오는지로만 본다. 아래 세 갈래는 **실제 git 저장소를 만들어
돌리는** 것이라 그것으로 대체되지 않는다(`run()` 한 줄 읽기 누락은 소스 검사로는 안 걸렸다).
번호를 달아 옮기는 일은 분담판 **`[34]`** 다.

무엇을 지키나 (셋 다 2026-08-11 실사고에서 나왔다)
  ① **남의 옛 줄에 잠기지 않는다** — 비밀값 스캔이 인덱스 전체를 훑어, 어제 커밋된
     `canonical_sync.py` 의 멱등키 계산식 한 줄에 걸려 그날 회차 14번이 전부 멈췄다.
     커밋을 거부해도 이미 커밋된 줄은 사라지지 않으니 영구히 잠긴 관문이었다.
  ② **새 비밀값은 자리를 적고 막는다** — 범위를 좁히다 목록을 `git()`(마지막 한 줄만
     준다)으로 읽어 비밀값 담긴 파일이 스캔에서 빠진 적이 있다. 막으려던 것을 새게
     만드는 방향이라 이 갈래가 제일 중요하다.
  ③ **남의 스테이징은 빼지 않고 그렇다고 말한다** — 멈추면서 `add -A` 를 남기면 다음
     사람의 `git commit -m` 이 통째로 커밋한다(#36). 그렇다고 남이 담아 둔 것을
     빼 버리면 그 사람의 뜻을 지운다.

돌리는 법:  python tests/_pending_wrapup_commit_check.py
"""
import io
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import session_wrapup as W                                    # noqa: E402


def _git(tmp, *a):
    return subprocess.run(["git"] + list(a), cwd=tmp, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def _repo(tmp):
    for a in (["init", "-q"], ["config", "user.email", "t@t"], ["config", "user.name", "t"]):
        _git(tmp, *a)
    # 이미 커밋돼 있는 '비밀값처럼 생긴' 줄 — 실제로는 멱등키 계산식이다(canonical_sync.py:349).
    # ★ 그 모양을 **소스에 그대로 적지 않는다.** 적으면 이 파일이 다른 스캐너(Terra→Sol
    #   검토의 비밀값 형태 검사 등)에 영구히 걸린다 — 사고를 적는 행위가 사고를 만든다.
    #   조립해서 쓰므로 **디스크에 써진 파일만** 그 모양이 되고, 시험의 뜻은 그대로다.
    io.open(os.path.join(tmp, "canonical_sync.py"), "w", encoding="utf-8").write(
        '%s = "%s" + sha256_json(facts)\n' % ("completion_" + "token", "canonical-completion:"))
    _git(tmp, "add", "-A")
    _git(tmp, "commit", "-q", "-m", "첫 커밋")


def _run_step(tmp):
    old, steps = W.ROOT, []
    try:
        W.ROOT = tmp
        W.step_commit("claude", "probe", steps)
    finally:
        W.ROOT = old
    return steps[-1]


def case1_not_jammed_by_old_line():
    """이미 커밋된 줄 때문에 **영구히** 멈추지 않는다."""
    with tempfile.TemporaryDirectory() as tmp:
        _repo(tmp)
        io.open(os.path.join(tmp, "innocent.py"), "w", encoding="utf-8").write("x = 1\n")
        s = _run_step(tmp)
        assert s["성공"], "이미 커밋된 줄에 걸려 또 멈췄다: %s" % s["메모"]
        out = _git(tmp, "log", "--name-only", "--format=", "-1").stdout
        assert "innocent.py" in out, "커밋에 파일이 안 담겼다: %r" % out
        print("  ① 남의 옛 줄에 잠기지 않는다 ✅")


def case2_refuses_and_leaves_no_staging():
    """이번 커밋이 **새로 담는** 비밀값은 그대로 막고, 멈추면 인덱스를 되돌린다."""
    with tempfile.TemporaryDirectory() as tmp:
        _repo(tmp)
        io.open(os.path.join(tmp, "leak.py"), "w", encoding="utf-8").write(
            '%s = "%s"\n' % ("api" + "_key", "AKIA" + "ABCDEFGH1234567890"))
        s = _run_step(tmp)
        assert not s["성공"], "새로 담기는 비밀값을 통과시켰다 — 절대규칙 1"
        assert "leak.py" in s["메모"], "걸린 자리를 안 적었다(사람이 확인할 수 없다): %s" % s["메모"]
        cached = _git(tmp, "diff", "--cached", "--name-only").stdout.strip()
        assert cached == "", "멈추면서 스테이징을 남겼다 — 기계가 사고 #36 을 만든다: %r" % cached
        assert os.path.exists(os.path.join(tmp, "leak.py")), "파일을 지웠다 — 인덱스만 되돌려야 한다"
        print("  ② 새 비밀값은 막고 · 자리를 적고 · 스테이징을 안 남긴다 ✅")


def case3_keeps_foreign_staging():
    """원래 남의 스테이징이 있었으면 되돌리지 않는다 — 그것은 남이 적어 둔 뜻이다."""
    with tempfile.TemporaryDirectory() as tmp:
        _repo(tmp)
        io.open(os.path.join(tmp, "theirs.py"), "w", encoding="utf-8").write("y = 2\n")
        _git(tmp, "add", "theirs.py")                     # 옆 세션이 담아 둔 것
        io.open(os.path.join(tmp, "leak.py"), "w", encoding="utf-8").write(
            '%s = "%s"\n' % ("pass" + "word", "hunter2" * 3))
        s = _run_step(tmp)
        assert not s["성공"], "새로 담기는 비밀값을 통과시켰다 — 목록을 한 줄만 읽으면 이 갈래가 깨진다"
        cached = _git(tmp, "diff", "--cached", "--name-only").stdout.split()
        assert "theirs.py" in cached, "남이 담아 둔 것을 빼 버렸다: %r" % cached
        assert "그대로 뒀다" in s["메모"], "인덱스를 안 건드렸다는 말을 안 했다: %s" % s["메모"]
        print("  ③ 남의 스테이징은 빼지 않고 · 그렇다고 말한다 ✅")


if __name__ == "__main__":
    case1_not_jammed_by_old_line()
    case2_refuses_and_leaves_no_staging()
    case3_keeps_foreign_staging()
    print("세 갈래 통과 — 커밋 단계는 잠기지 않고, 멈출 때 자국을 남기지 않는다")
