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

종료코드: 0 이 세션 것 · 3 다른 앱(경로를 찍는다) · 4 모름·섞임
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
# ★ **설명 글은 근거가 아니다** (2026-08-20 실측으로 갈렸다). 이 저장소는
#   지시문·주석이 통째로 한국어 산문이라(정본 지시문 하나만 300KB) **어떤
#   한국어 지시든** 거기서 걸린다 — 그래서 ARTIS 지시가 '내것'으로 나왔다.
#   걷어낸 뒤 실측: `캡처하면` 2→0 · `멋진` 1→0 · `보고서로` 3→0 이고,
#   우리 것은 그대로 남는다(`리모컨` 57 · `캠프` 562).
ECHO_NAMES = {"claude.md", "agents.md", "incidents.md"}   # 지시를 **옮겨 적는** 파일
ECHO_DIRS = ("/reports/",)   # 회차가 만든 산출물
# 이 프로젝트의 워크트리·사본은 '남의 앱'이 아니다.
# ⚠ 사본은 파일이 **뿌리에** 있고 본체는 `ecount/` 밑에 있다(실측 2026-08-20) —
#   한 자리만 보면 우리 워크트리가 '다른 앱'으로 나간다([172]). 그리고 이름
#   하나는 우연히 겹칠 수 있어 **둘**을 본다.
SELF_MARK = ("session_handoff.py", "ai_claim.py")
SELF_DIRS = ("", "ecount")

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


def _is_echo(path):
    """지시를 **옮겨 적는** 파일인가 — 그러면 근거가 아니다."""
    low = path.replace(chr(92), "/").lower()
    if os.path.basename(low) in ECHO_NAMES:
        return True
    return any(d in low for d in ECHO_DIRS)


def code_only(text, ext):
    """설명 글(주석·삼중따옴표·마크다운)을 걷어낸 나머지.

    ★ 이 프로젝트가 다섯 번 밟은 자리다([301]⑨·[302]·[309]·[332]·[339] —
      *규칙을 세기 전에 주석을 걷어낸다*). 여기서는 그 사고가 한 겹 더 크다:
      **계기가 제 프로젝트의 지시문 사본을 읽고 '이 세션 것'이라 답했다.**
    """
    if ext == ".md":
        return ""                       # 마크다운은 통째로 설명 글이다
    out, inq = [], None
    for line in text.split(chr(10)):
        st = line.lstrip()
        if inq:
            if inq in line:
                inq = None
            continue
        if st.startswith("#") or st.startswith("//"):
            continue
        hit = None
        for q in ('"""', "'''"):
            if st.startswith(q) and line.count(q) == 1:
                hit = q
                break
        if hit:
            inq = hit                # 여는 줄부터 닫는 줄까지 통째로 설명 글
        else:
            out.append(line)
    return chr(10).join(out)


def hits_in(base, terms, budget=MAX_FILES, limit=6, cap=12, echo_out=None):
    """그 낱말들이 실제로 **코드에 쓰인 파일**. 없으면 빈 목록이다(짐작 아님).

    ★ **자기 자신은 근거가 아니다.** 첫 판이 그래서 틀렸다 — 이 파일 설명에
      적어 둔 낱말을 제가 찾아내고 "이 세션 것"이라 답했다.
    ★ **설명 글도 근거가 아니다** — `code_only` 로 걷어낸 뒤에 센다.
      산문에서만 걸린 낱말은 `echo_out` 에 담아 **숫자로 말한다**([169]);
      조용히 빼면 "아무것도 못 찾았다" 로 읽힌다.
    """
    found = {}
    me = os.path.abspath(__file__)
    for path in _walk(base, budget):
        if os.path.abspath(path) == me:
            continue                 # 내 그림자를 근거로 삼지 않는다
        try:
            with open(path, "rb") as fh:
                blob = fh.read(400_000)
        except Exception:
            continue                 # 못 읽은 파일은 '없다'가 아니라 그냥 못 본 것
        try:
            text = blob.decode("utf-8", "ignore")
        except Exception:
            continue
        code = None
        for t in terms:
            if t not in text:
                continue
            if not _is_echo(path):
                if code is None:
                    code = code_only(text, os.path.splitext(path)[1].lower())
                if t in code:
                    got = found.setdefault(t, [])
                    if len(got) <= cap:
                        got.append(path)
                    continue
            if echo_out is not None:   # 설명 글에서만 걸렸다 — 세어서 말한다
                echo_out[t] = echo_out.get(t, 0) + 1
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
    # ★ 창 없는 깃발은 **부르는 자리에** 적는다 — `**kw` 안에 숨기면 감사기가
    #   못 보고 [272] 가 '깃발 없이 띄운다'로 잡는다(실측 2026-08-20 관문).
    kw = dict(capture_output=True, text=True, encoding="utf-8", errors="replace",
              cwd=repo, timeout=20)
    try:
        last = subprocess.run(["git", "log", "-1", "--format=%ct %h %s"],
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), **kw)
        dirty = subprocess.run(["git", "status", "--porcelain"],
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), **kw)
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


def _is_our_copy(base):
    """이 프로젝트의 사본인가 — 뿌리·`ecount/` 두 자리를 다 본다."""
    for d in SELF_DIRS:
        here = os.path.join(base, d) if d else base
        if all(os.path.exists(os.path.join(here, f)) for f in SELF_MARK):
            return True
    return False


def _neighbors(terms):
    """이웃 앱에서 찾는다 — **우리 사본은 이웃이 아니다.**

    실측 2026-08-20: `Documents` 밑에 이 프로젝트의 워크트리 사본이 있다.
    그것을 이웃으로 세면 **우리 일이 '다른 앱 것'** 으로 나간다([172]).
    """
    out = []
    try:
        names = sorted(os.listdir(NEIGHBOR_ROOT))
    except Exception:
        return out               # 못 읽었으면 '없다'가 아니라 그냥 못 본 것
    here = os.path.abspath(os.path.dirname(ROOT))
    for name in names:
        base = os.path.join(NEIGHBOR_ROOT, name)
        if not os.path.isdir(base) or os.path.abspath(base) == here:
            continue
        if _is_our_copy(base):
            continue             # 이 프로젝트의 워크트리·사본이다
        got = hits_in(base, terms, budget=12_000, limit=3)
        if got:
            files = [os.path.relpath(p, base)
                     for v in got.values() for p in v][:3]
            out.append({"앱": name, "낱말": sorted(got),
                        "파일": files, "세션": live_session_in(base)})
        if len(out) >= 3:
            break
    return out


def judge(text):
    """이 지시가 어디 것인가. 갈래 넷 — 내것 · 다른앱 · 섞임 · 모름.

    ★ `GENERIC_MAX` 는 **빼는 문이 아니라 더 보라는 신호**다 (2026-08-20 실측).
      전에는 흔한 말을 근거에서 통째로 뺐는데, 산문을 걷어낸 뒤에는 그것이
      **가장 센 근거를 버리는 짓**이었다 — 실측으로 멀쩡한 지시
      (`돌발AS 미처리 사유 입력 캘린더 대표 캡처`)가 '못 찾음'으로 떨어졌다.
      잘못 걸러 내면 형님이 시킨 일을 안 하고도 한 줄 보고로 넘어간다([172]).
    """
    terms = words(text)
    if not terms:
        return {"갈래": "모름", "왜": "찾아볼 낱말을 못 뽑았다", "낱말": []}

    echo = {}
    mine = hits_in(ROOT, terms, echo_out=echo)
    prose_only = sorted(k for k in echo if k not in mine)
    base = {"낱말": terms, "산문만": prose_only,
            "셈": {k: len(v) for k, v in mine.items()}}
    strong = {k: v for k, v in mine.items() if len(v) <= GENERIC_MAX}

    def _out(**kw):
        d = dict(base)
        d.update(kw)
        return d

    if strong:
        # 드문 낱말이 코드에 있다 — 여기서 이웃까지 훑지 않는다([168])
        return _out(**{"갈래": "내것", "확신": "높",
                      "근거": {k: v[:2] for k, v in strong.items()}})

    others = _neighbors(terms)   # 흔한 말로만 걸렸거나 아예 없다 — 그때만 훑는다
    # ★ **있고 없고가 아니라 어느 쪽이 더 센가로 가른다** (2026-08-20 실측).
    #   `사유`·`입력`·`대표` 같은 말은 한국어 저장소면 어디에나 있어서, '있으면
    #   섞임'으로 두면 멀쩡한 쿠팡 지시가 매번 '사람이 고르라'로 떨어진다
    #   — 경보가 대부분 가짜면 진짜 경보가 묻힌다([170]).
    rival = max([len(o.get("낱말") or []) for o in others] or [0])
    if mine and rival >= len(mine):
        return _out(**{"갈래": "섞임", "어디": others,
                      "근거": {k: v[:2] for k, v in mine.items()},
                      "왜": "양쪽이 다 그 말을 쓴다 — 기계가 고르면 안 된다"})
    if mine:
        return _out(**{"갈래": "내것", "확신": "낮",
                      "근거": {k: v[:2] for k, v in mine.items()},
                      "어디": others,
                      "왜": ("흔한 말로만 걸렸다 — 그래도 이웃보다 이쪽에 더 많아 이쪽으로 본다"
                             " (이웃에도 조금 있으면 아래에 같이 적는다)")})
    if others:
        return _out(**{"갈래": "다른앱", "어디": others})
    return _out(**{"갈래": "모름",
                  "왜": "이 저장소 코드에도 이웃 앱에도 그 낱말이 없다 — 새로 만들 일이거나 "
                        "내가 못 찾은 것이다. 어느 쪽인지는 사람이 안다"})


def _where(v):
    """이웃 앱을 찍는다 — 살아 있는 세션이면 같이 고치지 말라고 말한다."""
    for o in v.get("어디") or []:
        print("  · %s   (%s)" % (o["앱"], ", ".join(o.get("낱말") or [])))
        for p in o["파일"]:
            print("      %s" % p)
        st = o.get("세션")
        if st is None:
            print("      세션: 확인 못 함 — 같이 고치기 전에 사람이 확인할 것")
        elif st.get("살아있음"):
            print("      ⚠ 살아 있는 세션 — 마지막 커밋 %s분 전 · 미커밋 %d개"
                  % (st["마지막커밋_분전"], st["미커밋"]))
            print("        같이 고치면 한쪽이 통째로 묻힌다(사고 #36) — 그 창에 넘긴다")
        else:
            print("      조용함 — 마지막 커밋 %s분 전 · 미커밋 %d개"
                  % (st["마지막커밋_분전"], st["미커밋"]))


def report(v):
    g = v.get("갈래")
    prose = v.get("산문만") or []
    cnt = v.get("셈") or {}

    def _ev():
        for k, paths in (v.get("근거") or {}).items():
            for p in paths:
                print("  %-14s %s  (코드 %d파일)"
                      % (k, os.path.relpath(p, ROOT), cnt.get(k, 0)))
        # 뺀 것은 **숫자로 말한다**([169]) — 조용히 빼면 '못 찾았다'로 읽힌다
        if prose:
            print("  · 설명 글(지시문·주석)에서만 걸린 말이라 근거에서 뺌: %s"
                  % ", ".join(prose))

    if g == "내것":
        print("이 세션 것입니다 (확신 %s) — 근거:" % v.get("확신", "?"))
        _ev()
        if v.get("왜"):
            print("  · %s" % v["왜"])
        _where(v)      # 이웃에도 조금 있으면 숨기지 않는다([169])
        return 0
    if g == "다른앱":
        print("★ 이 세션 것이 아닙니다 — 이 저장소 **코드**에 그 낱말이 한 건도 없습니다.")
        print("  찾아본 낱말: " + ", ".join(v.get("낱말") or []))
        _ev()
        _where(v)
        return 3
    if g == "섞임":
        print("양쪽에 다 있습니다 — 기계가 못 고릅니다(사람이 고른다).")
        _ev()
        _where(v)
        return 4
    print("모름 — %s" % v.get("왜", ""))
    print("  찾아본 낱말: " + ", ".join(v.get("낱말") or []))
    _ev()
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
