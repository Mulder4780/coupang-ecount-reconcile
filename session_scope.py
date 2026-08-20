# -*- coding: utf-8 -*-
"""session_scope.py — **이 지시가 이 세션 것인가**를 먼저 묻는다.

사용자 지시(2026-08-20): **"이 세션에 다른 세션 명령이 들어와도 잘 걸러내 앞으로"** ·
**"내가 잘못 올리는 경우가 종종 있어"**

★ 실제로 있었던 일이라 만든다. 그날 이 세션(쿠팡 통합업무)에 **ARTIS Control Hub**
  (Next.js 별도 저장소)의 '서식 문서' 화면 지시가 들어왔다 — 캡처만 보면 구별이
  안 된다. 그대로 시작했으면 **없는 파일을 찾아 헤매거나**, 더 나쁘게는 그 저장소
  워크트리를 **살아 있는 다른 세션과 동시에** 고칠 뻔했다(사고 #36).

지키는 것
  · ★ **짐작으로 거르지 않는다.** "이건 우리 게 아닌 것 같다"는 말은 근거가 아니다 —
    지시에 나온 **낱말을 실제로 찾아본 뒤**에 말한다. 잘못 걸러 내면 형님이 시킨
    일을 안 하고도 한 줄 보고로 넘어간다 — 못 하는 것보다 나쁘다([172]).
  · ★ **조용히 무시하지 않는다**([169]). 남의 것이면 **어느 앱·어느 파일인지**까지
    찾아 준다. 그래야 형님이 그 창에 그대로 넘길 수 있다.
  · ★ **'모름'을 '내 것'으로 치지 않는다.** 이 저장소에도 없고 이웃에도 없으면
    그것은 '새로 만들 일'일 수도 있고 '내가 못 찾은 것'일 수도 있다 — 갈라 말한다.
  · ★ **비싼 탐색은 뒤에 온다**([168]). 이 저장소를 먼저 보고, **0건일 때만**
    이웃 폴더를 훑는다. 이웃도 `node_modules`·`.next`·`.git` 은 건너뛴다.
  · ★ **남의 저장소에 살아 있는 세션이 있는지 같이 본다.** 있으면 그 파일을 고치는
    것 자체가 사고다 — 넘기라고 말한다.

쓰는 법
    python session_scope.py "현장점검표 캡처 엑셀 저장"
    python session_scope.py --file 캡처설명.txt

종료코드: 0 이 세션 것 · 3 다른 앱(경로를 찍는다) · 4 모름
"""
import os
import re
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:                                    # 무인 회차는 pythonw 라 stdout 이 None 이다([235])
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# 이웃 앱을 찾는 자리. **손으로 앱 이름을 적지 않는다** — 폴더를 훑어 찾는다.
NEIGHBOR_ROOT = os.path.dirname(os.path.dirname(ROOT))
SKIP_DIRS = {"node_modules", ".next", ".git", "__pycache__", ".venv", "venv",
             "dist", "build", ".vercel", "worktrees", ".claude"}
CODE_EXT = {".py", ".html", ".js", ".ts", ".tsx", ".jsx", ".md", ".json"}
MAX_FILES = 40_000          # 이웃 훑기 상한 — 여기서 몇 분을 쓰면 안 쓰느니만 못하다
# ★ **흔한 말로 판정하지 않는다** (첫 판이 그래서 틀렸다). `서식`·`캡처`·`엑셀로`
#   는 이 저장소 수십 파일에 있어 **어느 지시든 '내것'** 으로 만든다. 낱말이
#   이 저장소에서 이만큼 넘게 나오면 그것은 우리 어휘이지 그 화면의 이름이
#   아니다 — 근거에서 뺀다(빼되 **숫자로 말한다** · [169]).
GENERIC_MAX = 8

# 낱말 후보에서 걸러 낼 것 — 어디에나 있는 말은 근거가 못 된다
STOP = {
    "화면", "기능", "추가", "수정", "삭제", "저장", "구조", "정리", "적용", "변경",
    "문서", "이런", "이건", "여기", "모두", "전부", "그리고", "선택", "알고리즘",
    "코딩", "구현", "반영", "진행", "확인", "만들어", "넣어", "해줘", "고도화",
    "보이는", "가능하게", "업데이트", "버튼", "표시", "목록", "관리", "자동",
}


def words(text):
    """지시에서 **찾아볼 만한 낱말**만 뽑는다.

    한글 두 글자 이상 · 영문 세 글자 이상. 흔한 말(STOP)은 뺀다 —
    '저장'·'기능' 으로 찾으면 어느 저장소에서나 걸려 아무것도 안 가른다.
    """
    out, seen = [], set()
    for tok in re.findall(r"[가-힣]{2,}|[A-Za-z][A-Za-z0-9_]{2,}", str(text or "")):
        t = tok.strip()
        if t in STOP or t.lower() in STOP or t in seen:
            continue
        seen.add(t)
        out.append(t)
    return out[:12]                      # 열둘이면 충분하다 — 더 보면 느리기만 하다


def _walk(base, budget):
    """(경로, 이름) — 건너뛸 폴더는 안 들어간다."""
    n = 0
    for cur, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith(".")]
        for name in files:
            if os.path.splitext(name)[1].lower() not in CODE_EXT:
                continue
            n += 1
            if n > budget:
                return
            yield os.path.join(cur, name)


def hits_in(base, terms, budget=MAX_FILES, limit=6, cap=12):
    """그 낱말들이 실제로 **쓰인 파일**. 없으면 빈 목록이다(짐작 아님).

    ★ **자기 자신은 근거가 아니다.** 첫 판이 그래서 틀렸다 — 이 파일 설명에
      `현장점검표` 라고 적어 둔 것을 제가 찾아내고 "이 세션 것"이라 답했다
      (계기가 제 그림자를 보고 판정한 자리 · [301]⑨·[302] 와 같은 함정).
    ★ 낱말마다 **몇 파일에 있는지** 센다 — 흔한 말인지 그 화면 이름인지는
      그 숫자가 말해 준다."""
    found = {}
    me = os.path.abspath(__file__)
    for path in _walk(base, budget):
        if os.path.abspath(path) == me:
            continue                 # 내 그림자를 근거로 삼지 않는다
        try:
            with open(path, "rb") as fh:
                blob = fh.read(400_000)
        except Exception:
            continue                     # 못 읽은 파일은 '없다'가 아니라 그냥 못 본 것
        try:
            text = blob.decode("utf-8", "ignore")
        except Exception:
            continue
        for t in terms:
            if t in text:
                got = found.setdefault(t, [])
                if len(got) <= cap:
                    got.append(path)
        # 낱말마다 cap 까지만 세면 흔한 말인지 아닌지는 이미 갈린다
        if found and all(len(v) > cap for v in found.values()) and len(found) >= limit:
            break
    return found


def live_session_in(repo):
    """그 저장소에 **살아 있는 세션**이 있나 — 있으면 같이 고치면 안 된다(사고 #36).

    근거는 짐작이 아니라 둘이다: 최근 커밋 시각 · 미커밋 변경.
    못 읽으면 `None`(모름)이다 — '없다'고 하지 않는다([169]).
    """
    if not os.path.isdir(os.path.join(repo, ".git")):
        return None
    kw = dict(capture_output=True, text=True, encoding="utf-8", errors="replace",
              cwd=repo, timeout=20,
              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    try:
        last = subprocess.run(["git", "log", "-1", "--format=%ct %h %s"], **kw)
        dirty = subprocess.run(["git", "status", "--porcelain"], **kw)
    except Exception:
        return None
    if last.returncode != 0:
        return None
    try:
        when = int((last.stdout or "0").split(" ", 1)[0])
    except Exception:
        return None
    mins = (time.time() - when) / 60.0
    n_dirty = len([x for x in (dirty.stdout or "").splitlines() if x.strip()])
    return {"마지막커밋_분전": round(mins, 1),
            "제목": (last.stdout or "").strip()[:90],
            "미커밋": n_dirty,
            "살아있음": mins < 60 or n_dirty > 0}


def judge(text):
    """이 지시가 어디 것인가. 갈래 셋 — 내것 · 다른앱 · 모름."""
    terms = words(text)
    if not terms:
        return {"갈래": "모름", "왜": "찾아볼 낱말을 못 뽑았다", "낱말": []}

    mine = hits_in(ROOT, terms)
    # ★ 판정은 **흔하지 않은 낱말**로만 한다. 흔한 말은 버리지 않고 따로 적는다
    #   — 조용히 빼면 "아무것도 못 찾았다"로 읽힌다([169]).
    strong = {k: v for k, v in mine.items() if len(v) <= GENERIC_MAX}
    generic = sorted(k for k, v in mine.items() if len(v) > GENERIC_MAX)
    if strong:
        return {"갈래": "내것", "낱말": terms, "흔한말": generic,
                "근거": {k: v[:2] for k, v in strong.items()}}

    # 여기서만 이웃을 훑는다 — 이 저장소에 있으면 볼 필요가 없다([168])
    others = []
    try:
        names = sorted(os.listdir(NEIGHBOR_ROOT))
    except Exception:
        names = []
    for name in names:
        base = os.path.join(NEIGHBOR_ROOT, name)
        if not os.path.isdir(base) or os.path.abspath(base) == os.path.abspath(
                os.path.dirname(ROOT)):
            continue
        got = hits_in(base, terms, budget=12_000, limit=3)
        got = {k: v for k, v in got.items() if len(v) <= GENERIC_MAX}
        if got:
            files = [p for v in got.values() for p in v][:3]
            others.append({"앱": name, "파일": files, "세션": live_session_in(base)})
        if len(others) >= 3:
            break

    if others:
        return {"갈래": "다른앱", "낱말": terms, "흔한말": generic, "어디": others}
    return {"갈래": "모름", "낱말": terms, "흔한말": generic,
            "왜": "이 저장소에도 이웃 앱에도 그 낱말이 없다 — 새로 만들 일이거나 "
                  "내가 못 찾은 것이다. 어느 쪽인지는 사람이 안다"}


def report(v):
    g = v.get("갈래")
    gen = v.get("흔한말") or []
    if g == "내것":
        print("이 세션 것입니다 — 근거:")
        for k, paths in (v.get("근거") or {}).items():
            for p in paths:
                print("  %-14s %s" % (k, os.path.relpath(p, ROOT)))
        return 0
    if g == "다른앱":
        print("★ 이 세션 것이 아닙니다 — 이 저장소에 그 낱말이 한 건도 없습니다.")
        print("  찾아본 낱말: " + ", ".join(v.get("낱말") or []))
        if gen:
            print("  (흔한 말이라 근거에서 뺌: %s)" % ", ".join(gen))
        for o in v.get("어디") or []:
            print("  · %s" % o["앱"])
            for p in o["파일"]:
                print("      %s" % p)
            s = o.get("세션")
            if s is None:
                print("      세션: 확인 못 함 — 같이 고치기 전에 사람이 확인할 것")
            elif s.get("살아있음"):
                print("      ⚠ 살아 있는 세션 — 마지막 커밋 %s분 전 · 미커밋 %d개"
                      % (s["마지막커밋_분전"], s["미커밋"]))
                print("        같이 고치면 한쪽이 통째로 묻힌다(사고 #36) — 그 창에 넘긴다")
            else:
                print("      조용함 — 마지막 커밋 %s분 전 · 미커밋 %d개"
                      % (s["마지막커밋_분전"], s["미커밋"]))
        return 3
    print("모름 — %s" % v.get("왜", ""))
    print("  찾아본 낱말: " + ", ".join(v.get("낱말") or []))
    if gen:
        print("  (흔한 말이라 근거에서 뺌: %s — 이 저장소 %d파일 넘게 나온다)"
              % (", ".join(gen), GENERIC_MAX))
    return 4


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--file" in argv:
        i = argv.index("--file")
        if i + 1 >= len(argv):
            print("--file 뒤에 경로가 필요합니다")
            return 2
        try:
            text = open(argv[i + 1], encoding="utf-8", errors="replace").read()
        except Exception as e:
            print("못 읽었습니다: %s" % e)
            return 2
    else:
        text = " ".join(a for a in argv if not a.startswith("--"))
    if not text.strip():
        print(__doc__.strip().splitlines()[0])
        print("쓰는 법: python session_scope.py \"<지시 문구>\"")
        return 2
    return report(judge(text))


if __name__ == "__main__":
    raise SystemExit(main())
