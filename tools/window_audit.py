# -*- coding: utf-8 -*-
"""자식 프로세스를 띄우면서 **창을 달 수 있는 자리**를 전부 센다 (2026-08-14).

★ 왜 필요한가 — 예약 작업 13개는 이미 `pythonw`·`run_hidden.vbs` 로 창이 없다(`[248]`).
  그런데 형님 화면에는 여전히 검은 창이 떴다. 남은 것은 **회차가 띄운 자식**이다:
  **콘솔 없는 부모가 콘솔 exe 를 깃발 없이 띄우면 윈도우가 새 콘솔을 할당한다.**
  즉 부모를 `pythonw` 로 바꾼 것이 오히려 자식의 창을 새로 만든 셈이다
  (2026-08-13 cloudflared 실사고와 같은 모양).

★ 두더지잡기를 하지 않는다. 한 자리를 고치면 다음 주에 다른 자리가 뜬다 —
  그래서 자리를 **세어** 놓고 검증이 그 수를 지킨다.

★ `DETACHED_PROCESS`(0x8)는 창을 없애는 깃발이 **아니다.** 부모 콘솔을 안 물려받는다는
  뜻일 뿐이라, 자식이 콘솔 exe 면 오히려 새 콘솔을 할당해 창이 뜬다.

쓰기:  python tools/window_audit.py          # 창 달릴 수 있는 자리 목록
      python tools/window_audit.py --count  # 건수만 (검증이 쓴다)
"""
from __future__ import annotations

import argparse
import json
import ast
import os
import re
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:                                    # pythonw 에서는 stdout 이 None 이다
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SPAWNERS = {"run", "Popen", "call", "check_call", "check_output"}

# 훑지 않는 곳 — 남의 체크아웃과 파생물
SKIP_DIRS = {".claude", "__pycache__", ".git", "node_modules", "updates", "_보관"}

# 콘솔 창이 실제로 뜨는 실행기. 여기 없는 것(예: pythonw)은 창이 없다.
CONSOLE_EXE = ("powershell", "cmd", "taskkill", "tasklist", "schtasks", "git",
               "cloudflared", "tailscale", "python.exe", "python\"", "wmic",
               "reg", "netsh", "curl", "npm", "node", "where", "chcp", "attrib",
               "robocopy", "xcopy", "sc.exe", "wscript", "cscript")


def _flag_var_ok(call: ast.Call, tree) -> bool:
    """`creationflags=flags` 처럼 **변수로 넘기는** 자리를 따라간다 (2026-08-14).

    이걸 안 하면 멀쩡히 깃발을 단 자리가 '못 읽음'으로 남아, 진짜 못 읽은 자리와
    섞여 아무도 안 본다([170]). 같은 파일 안에서 그 이름에 `CREATE_NO_WINDOW` 를
    넣는 대입이 있으면 통과시킨다 — 없으면 여전히 '모름'이다.
    """
    for kw in call.keywords:
        if kw.arg == "creationflags" and isinstance(kw.value, ast.Name):
            want = kw.value.id
            for node in ast.walk(tree):
                if isinstance(node, ast.Assign) and any(
                        isinstance(t, ast.Name) and t.id == want for t in node.targets):
                    if "CREATE_NO_WINDOW" in ast.unparse(node.value):
                        return True
    return False


def safe_helpers(trees):
    """`**헬퍼()` 로 펼쳐 쓰는 헬퍼 중 **본문에 `CREATE_NO_WINDOW` 가 있는** 것.

    ★ **이름으로 믿지 않는다 — 본문을 읽는다** (2026-08-21).  전에는
      `background_popen_kwargs` 라는 **이름 하나를 못 박아** 뒀다.  그래서 같은
      일을 다른 이름으로 하는 저장소(예: `no_window_kwargs`)는 멀쩡히 깃발을
      달고도 전부 '못 읽음'으로 쌓였고, 진짜 못 읽은 자리가 그 안에 묻혔다([170]).
      이름 목록을 늘리면 사본이 되고, 이름만 보고 믿으면 **창이 뜨는데 0곳이라
      말하는** 자리가 된다([169]) — 그래서 정의를 찾아 본문을 본다.
    ★ 정의를 못 찾으면(다른 폴더·건너뛴 폴더) 그 자리는 그대로 '모름'이다."""
    out = set()
    for _rel, tree in trees:
        for node in ast.walk(tree):
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            try:
                if "CREATE_NO_WINDOW" in ast.unparse(node):
                    out.add(node.name)
            except Exception:
                continue
    return out


def _has_no_window(call: ast.Call, helpers=()) -> bool:
    """이 호출이 `CREATE_NO_WINDOW` 를 확실히 달고 있나.

    ★ 이름만 본다 — 값을 계산하지 않는다. `flags` 같은 변수에 담아 넘기는 자리는
      '모름'으로 두고 사람이 본다. **모르는 것을 안다고 하지 않는다**([169]).
    """
    for kw in call.keywords:
        if kw.arg not in ("creationflags", "startupinfo"):
            continue
        if "CREATE_NO_WINDOW" in ast.unparse(kw.value):
            return True
        if kw.arg == "startupinfo":
            continue
        # proc_guard 헬퍼를 통째로 펼친 경우(**background_popen_kwargs())
    for kw in call.keywords:
        if kw.arg is not None:
            continue
        try:
            txt = ast.unparse(kw.value)
        except Exception:
            continue
        # 이 저장소의 정본 헬퍼 — `safe_helpers()` 가 본문을 읽어 저절로 잡지만,
        # 그 함수를 못 읽는 자리(다른 폴더에서 부를 때)를 위해 이름도 남겨 둔다.
        if "background_popen_kwargs" in txt:
            return True
        for name in helpers:
            if name + "(" in txt:
                return True
    return False


def _looks_console(call: ast.Call) -> str:
    """첫 인자가 콘솔 exe 를 가리키나. 가리키면 그 낱말을 돌려준다."""
    if not call.args:
        return "?"
    try:
        txt = ast.unparse(call.args[0])
    except Exception:
        return "?"
    low = txt.lower()
    for name in CONSOLE_EXE:
        if name in low:
            return name
    return "?"


# ──────────────────────────────── 살아 있는 예약 작업의 실행기 (2026-08-16, 분담판 [107])
#: 창이 **뜨는** 실행기. `pythonw`·`wscript` 는 창이 없다(`cscript` 는 콘솔이라 뜬다).
WINDOWED_EXE = ("python.exe", "cmd.exe", "powershell.exe", "pwsh.exe", "cscript.exe")
QUIET_EXE = ("pythonw.exe", "wscript.exe")


def exe_verdict(exe, args=""):
    """예약 작업 실행기 한 줄이 창을 다나 → `창뜸` · `조용` · `모름`.

    ★ `[248]` 은 **설치본**이 창 뜨는 실행기를 다시 등록하는지 본다. 그런데 설치본을
      고쳐도 **이미 등록돼 있는 작업은 그대로다** — 반대로 살아 있는 작업만 고치면
      기계를 새로 만들 때 되살아난다. 그래서 둘 다 봐야 한다.
    ★ 모르면 `모름` 이다. 못 읽은 것을 '조용'으로 세면 계기 자신이 눈이 먼다(`[169]`)."""
    name = os.path.basename(str(exe or "").strip().strip('"')).lower()
    if not name:
        return "모름"
    # 훅·예약작업은 실행기를 **확장자 없이** 적기도 한다(PATH 로 풀린다).
    # 표는 `.exe` 로 적혀 있으므로 꼬리를 떼고 대 본다 — 안 그러면 `pythonw` 가
    # `모름` 으로 새어 멀쩡한 설정에 거짓 경보가 난다(2026-08-21 실측 · `[165]`).
    stem = name[:-4] if name.endswith(".exe") else name
    if name in QUIET_EXE or stem in {x[:-4] for x in QUIET_EXE}:
        return "조용"
    if (name in WINDOWED_EXE or stem in {x[:-4] for x in WINDOWED_EXE}
            or name.endswith((".bat", ".cmd"))):
        return "창뜸"
    return "모름"


def live(state_path=None):
    """살아 있는 예약 작업 중 창 뜨는 실행기로 등록된 것 → `(목록, 왜못함)`.

    회차(`schedule_watch`)가 이미 물어서 써 둔 것을 읽는다 — 여기서 스케줄러를
    다시 묻지 않는다(`[168]`). 못 읽으면 **빈 목록이 아니라 이유**를 돌려준다."""
    import json
    if state_path is None:
        try:
            if ROOT not in sys.path:        # tools/ 에서 직접 부르면 루트가 안 잡힌다
                sys.path.insert(0, ROOT)
            import schedule_watch
            state_path = schedule_watch.STATE
        except Exception as exc:
            return [], "회차 감시를 못 불렀다(%s)" % type(exc).__name__
    try:
        with open(state_path, encoding="utf-8") as fh:
            rows = (json.load(fh) or {}).get("작업") or []
    except (OSError, ValueError):
        return [], "회차 감시 기록을 못 읽었다 — python schedule_watch.py 를 먼저 돌린다"
    if not rows:
        return [], "회차 기록이 비어 있다"
    if not any(r.get("실행기") for r in rows):
        # 0곳이 '없다'인지 '안 봤다'인지 가른다(`[169]`).
        return [], "실행기를 한 줄도 못 읽었다(옛 기록이거나 조회가 막혔다)"
    out = []
    for r in rows:
        for a in (r.get("실행기") or []):
            판정 = exe_verdict(a.get("exe"), a.get("args"))
            if 판정 != "조용":
                out.append((r.get("작업", ""), str(a.get("exe") or ""), 판정))
    return out, ""


def _is_ours(value):
    """이 프로젝트가 건 항목인가 — 남의 자동실행은 안 본다(`[172]`)."""
    v = str(value or "").lower()
    return os.path.dirname(ROOT).lower() in v or ROOT.lower() in v


def _first_token(cmd):
    """명령줄 첫 토막(실행기). 따옴표 없이 공백이 섞이면 틀리는데, 그때는
    `exe_verdict` 가 `모름` 을 주므로 **조용으로 새지 않는다**(`[169]`)."""
    s = str(cmd or "").strip()
    if s.startswith(chr(34)):
        end = s.find(chr(34), 1)
        return s[1:end] if end > 0 else s[1:]
    return s.split(" ", 1)[0]


def autorun():
    """로그인 자동실행(HKCU Run)에 **이 프로젝트 항목**이 창 뜨는 실행기로 걸렸나
    → `(목록, 왜못함, 본개수)`.

    ★ 왜 여기가 구멍이었나 (2026-08-21 형님 지시 "어떤 계정 어떤 세션에서 진행해도
      팝업은 백그라운드로") — `[248]` 은 **설치본**을, `[272]` 는 **소스**를,
      `live()` 는 **예약 작업**을 본다. 그런데 `[263]` 대로 예약 작업 등록이 막힌
      기계에서는 설치기가 **HKCU 로그인 자동실행으로 스스로 전환**한다. 그 자리는
      아무도 안 봐서, 창 뜨는 실행기가 들어가면 **부팅할 때마다** 창이 뜨는데
      어느 화면에도 안 뜬다(`[169]`).
    ★ 판정은 `exe_verdict` 를 빌린다(`[162]`) — 여기서 새로 만들면 예약 작업과 갈린다.
    ★ **`powershell` 로 읽지 않는다** — 그 자체가 콘솔 exe 라, 창을 막는 감사기가
      스스로 창을 띄우게 된다(`[272]` 가 제 코드를 잡는 자리다). `winreg` 는 창이 없다.
    ★ 보는 것은 HKCU Run 하나다 — 설치기가 쓰는 자리가 거기뿐이다. 넓히면 남의
      항목까지 판단하게 된다(`[172]`).
    ★ **'없다'와 '못 읽었다'를 가른다** — 본개수가 곧 그 구별이다."""
    try:
        import winreg
    except ImportError:
        return [], "윈도우가 아니다 — 로그인 자동실행이 없는 기계다", 0
    path = "Software/Microsoft/Windows/CurrentVersion/Run".replace("/", chr(92))
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
    except OSError as exc:
        return [], "로그인 자동실행을 못 읽었다(%s)" % type(exc).__name__, 0
    out, seen, i = [], 0, 0
    try:
        while True:
            try:
                name, val, _t = winreg.EnumValue(key, i)
            except OSError:
                break
            i += 1
            if not _is_ours(val):
                continue
            seen += 1
            exe = _first_token(val)
            verdict = exe_verdict(exe)
            if verdict != "조용":
                out.append((name, exe, verdict))
    finally:
        key.Close()
    return out, "", seen



# 시작 폴더에서 **우리가 판정할 수 있는 것**만 본다 — 스크립트 자동실행이다.
# `.lnk`·`.exe` 는 설치형 앱이 스스로 건 것이고 사람이 일부러 켠 것이라 안 본다(`[172]`).
STARTUP_SCRIPT_EXT = (".vbs", ".bat", ".cmd", ".ps1", ".js")


def _read_text_any(path):
    """인코딩을 모르는 파일을 읽는다 — 못 읽으면 None(=모름)."""
    for enc in ("utf-16", "utf-8-sig", "utf-8", "cp949"):
        try:
            return open(path, encoding=enc).read()
        except (UnicodeError, UnicodeDecodeError, OSError):
            continue
    return None


def _strip_vbs_comments(text):
    """VBScript 주석(줄 앞 ')을 걷어낸다.

    ★ **규칙을 세기 전에 설명을 걷어낸다.** 이 저장소가 여섯 번 밟은 자리다
      (`[301]`(9)·`[302]`·`[309]`·`[332]`·`[339]`·`[370]`). `run_hidden.vbs` 의
      머리 주석에는 '창이 뜬다'는 설명이 그대로 적혀 있어, 안 걷으면 설명문이
      위반으로 잡힌다."""
    out = []
    for line in (text or "").split("\n"):
        s = line.lstrip()
        if s.startswith("'"):
            continue
        out.append(line)
    return "\n".join(out)


def vbs_window_verdict(text):
    """`.vbs` 가 자식을 **창 없이** 띄우나 → `조용` · `창뜸` · `모름`.

    ★ 왜 이 함수가 필요한가 (2026-08-21 형님 지시가 두 번째로 온 자리).
      `exe_verdict` 는 `wscript.exe` 를 **무조건 조용**이라 판정한다. 그런데 창이
      뜨는지 정하는 것은 실행기가 아니라 **그 vbs 안의 `Run` 두 번째 인자**다 —
      `0` 이면 숨김, `1`(2·3·7…)이면 창이 뜬다. 즉 누가 `run_hidden.vbs` 의 `0` 을
      `1` 로 바꾸면 **그 실행기로 도는 회차가 전부 창을 다는데** 감사기는 나란히
      `0곳` 을 확언한다. 계기가 실제로 재지 않으면서 초록을 내는 자리다(`[169]`).
    ★ `Run(cmd, style, wait)` 에서 **뒤에서부터** 읽는다 — 명령 문자열 안에도
      콤마가 있어서(예: powershell 인자) 앞에서 세면 엉뚱한 숫자를 창모드로 읽는다.
    ★ 최소화(2·7)도 `창뜸` 이다. 작업표시줄에 뜨는 것도 형님 화면을 건드린다."""
    if text is None:
        return "모름", "파일을 못 읽었다"
    body = _strip_vbs_comments(text)
    if ".run" not in body.lower():
        return "모름", "Run 호출을 못 찾았다"
    styles = []
    # (1) 세 인자: Run cmd, style, wait     (2) 두 인자: Run cmd, style
    for m in re.finditer(r",\s*(\d+)\s*,\s*(?:True|False)\s*\)?", body, re.I):
        styles.append(int(m.group(1)))
    if not styles:
        for m in re.finditer(r"\.Run\b[^\n]*?,\s*(\d+)\s*(?:\)|$)", body, re.I | re.M):
            styles.append(int(m.group(1)))
    if not styles:
        return "모름", "Run 의 창모드 인자를 못 읽었다"
    bad = [s for s in styles if s != 0]
    if bad:
        return "창뜸", "Run 창모드 %s (0 이 아니면 창이 뜬다)" % ", ".join(str(x) for x in bad)
    return "조용", "Run 창모드 0 × %d곳" % len(styles)


def _startup_dirs():
    """사용자·공용 시작 폴더. **둘 다 본다** — 형님 지시가 '어떤 계정' 이다."""
    dirs = []
    app = os.environ.get("APPDATA")
    pro = os.environ.get("ProgramData")
    tail = os.path.join("Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    if app:
        dirs.append(os.path.join(app, tail))
    if pro:
        dirs.append(os.path.join(pro, tail))
    return dirs


def startup(dirs=None):
    """시작 폴더의 **스크립트 자동실행**이 창을 다나 → `(목록, 왜못함, 본개수)`.

    ★ 여기가 네 축(소스·예약작업·로그인항목·훅) 다음의 다섯째 구멍이었다.
      실측 2026-08-21: 시작 폴더에 자동실행 넷이 있었고 그중 셋이 `.vbs` 인데
      **어느 축도 그 파일을 한 글자도 안 봤다.** 지금은 셋 다 `0`(숨김)이지만,
      그것을 **재는 계기가 없다는 것**이 문제다 — 하나가 `1` 로 바뀌어도 조용하다.
    ★ 남의 앱은 안 본다(`.lnk`·`.exe`). 형님이 일부러 켠 것을 위반이라 부르면
      멀쩡한 것을 고치러 간다(`[172]`).
    ★ `.bat`·`.cmd`·`.ps1` 자동실행은 그 자체가 콘솔이라 `창뜸` 이다.
    ★ **못 읽은 것을 조용으로 세지 않는다**(`[169]`) — `모름` 으로 남긴다."""
    out, seen, bad = [], 0, []
    for d in (dirs if dirs is not None else _startup_dirs()):
        if not os.path.isdir(d):
            continue
        try:
            names = sorted(os.listdir(d))
        except OSError as exc:
            bad.append("%s(%s)" % (os.path.basename(d), type(exc).__name__))
            continue
        for fn in names:
            ext = os.path.splitext(fn)[1].lower()
            if ext not in STARTUP_SCRIPT_EXT:
                continue
            seen += 1
            path = os.path.join(d, fn)
            if ext == ".vbs":
                판정, 근거 = vbs_window_verdict(_read_text_any(path))
            elif ext == ".js":
                # wscript 로 열리면 조용하지만 무엇을 띄우는지는 못 읽는다.
                판정, 근거 = "모름", "JScript — 무엇을 띄우는지 못 읽었다"
            else:
                판정, 근거 = "창뜸", "%s 자동실행은 그 자체가 콘솔이다" % ext
            if 판정 != "조용":
                out.append((fn, 판정, 근거))
    why = ""
    if bad and not seen:
        why = "시작 폴더를 못 읽었다(%s)" % ", ".join(bad)
    elif bad:
        out.append((", ".join(bad), "확인못함", "시작 폴더를 못 읽었다"))
    return out, why, seen


# 세션이 최근에 연 저장소를 이웃으로 세는 창(일).
# 넓히면 몇 달 전에 한 번 열어 본 폴더까지 목록에 올라와 아무도 안 본다([170]).
NEIGHBOR_DAYS = 30


def _claude_projects_dir():
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return os.path.join(home, ".claude", "projects")


def _cwd_of(path, limit=65536):
    """대화기록 한 파일에서 그 세션이 연 폴더를 읽는다 — 앞부분만 본다([168]).

    파일이 수십 MB 라 통째로 읽으면 회차가 그만큼 느려진다. `cwd` 는 첫 줄들에 있다.
    """
    try:
        with open(path, "rb") as fh:
            head = fh.read(limit)
    except OSError:
        return None
    for line in head.decode("utf-8", "replace").split(chr(10)):
        if chr(34) + "cwd" + chr(34) not in line:
            continue
        try:
            obj = json.loads(line)
        except Exception:
            continue
        cwd = obj.get("cwd") if isinstance(obj, dict) else None
        if isinstance(cwd, str) and cwd.strip():
            return cwd.strip()
    return None


def _is_temp(p):
    """임시 폴더 안인가 — 세션 스크래치패드가 이웃 저장소로 세이는 것을 막는다."""
    low = os.path.normcase(p)
    for env in ("TEMP", "TMP"):
        t = os.environ.get(env)
        if t and low.startswith(os.path.normcase(t)):
            return True
    return os.path.normcase(os.path.join("appdata", "local", "temp")) in low

def _fold_worktree(p):
    """워크트리는 본체로 접는다 — 같은 저장소를 여러 번 세면 목록만 부푼다."""
    mark = os.path.normcase(os.path.join(".claude", "worktrees"))
    i = os.path.normcase(p).find(mark)
    if i > 0:
        return p[:i].rstrip(chr(92) + chr(47))
    return p


def _session_cwds(days=NEIGHBOR_DAYS, base=None):
    """이 PC 의 세션이 **실제로 연 폴더** — 지어내지 않고 대화기록에서 읽는다.

    ★ 왜 자동실행만으로는 모자란가 (2026-08-21 실측). `neighbors()` 의 옛 근거는
      자동실행·시작폴더가 가리키는 경로뿐이라 **자동으로 돌지는 않지만 세션이 매일
      고치는 저장소**가 통째로 빠졌다(하늘링고). 형님 지시는 *어떤 계정 **어떤
      세션**에서 진행해도* 였다 — 자동으로 도는 것만 세면 그 절반을 안 본 것이다.
    ★ 폴더 **이름**으로 되돌리지 않는다. 한글 경로는 대화기록 폴더 이름에서
      `C--Users-hueng-Documents------------` 처럼 전부 하이픈으로 뭉개져 복원이
      불가능하다(실측). 그래서 기록 **안**의 `cwd` 를 읽는다.
    ★ 폴더마다 **가장 최근 기록 하나**의 앞부분만 연다([168]).
    ★ 못 읽으면 빈 목록이 아니라 **이유**를 같이 돌려준다([169]) —
      '이웃이 없다' 와 '못 셌다' 는 다른 사실이다.
    """
    base = base or _claude_projects_dir()
    if not os.path.isdir(base):
        return [], "대화기록 폴더가 없다: %s" % base
    cut = time.time() - max(1, int(days)) * 86400
    out = {}
    try:
        names = sorted(os.listdir(base))
    except OSError as exc:
        return [], "대화기록 폴더를 못 읽었다(%s)" % type(exc).__name__
    for name in names:
        d = os.path.join(base, name)
        if not os.path.isdir(d):
            continue
        newest, newest_m = None, 0.0
        try:
            for fn in os.listdir(d):
                if not fn.endswith(".jsonl"):
                    continue
                q = os.path.join(d, fn)
                try:
                    m = os.path.getmtime(q)
                except OSError:
                    continue
                if m > newest_m:
                    newest, newest_m = q, m
        except OSError:
            continue
        if not newest or newest_m < cut:
            continue
        cwd = _cwd_of(newest)
        if not cwd:
            continue
        cwd = _fold_worktree(cwd)
        # 스크래치패드·임시폴더는 저장소가 아니다 — 훑어 봐야 남의 쓰레기다.
        if _is_temp(cwd):
            continue
        if os.path.isdir(cwd):
            out[os.path.normcase(cwd)] = cwd
    return sorted(out.values()), None

def neighbors(with_why=False):
    """이 PC 에서 **자동으로 도는 다른 프로젝트 폴더** — 지어내지 않고 읽어서 센다.

    ★ 왜 필요한가. 이 감사기는 `ROOT`(이 저장소) 하나만 훑는다. 그런데 형님 지시는
      "**어떤 계정 어떤 세션에서 진행해도**" 다 — 다른 세션이 고치는 저장소의 코드가
      창을 띄우면 여기서는 아무도 못 센다. 그 목록을 손으로 적으면 사본이 되어 늘
      뒤처지므로(`[162]`), **자동실행 항목이 실제로 가리키는 경로**에서 뽑는다.
    ★ 목록까지만 낸다 — 남의 저장소를 자동으로 훑어 경보를 내지 않는다(`[172]`).
      규칙이 다를 수 있고, 거짓 경보는 못 잡는 것보다 나쁘다. 볼 때는 사람이
      `python tools/window_audit.py --root <그 폴더>` 로 부른다."""
    BS, Q = chr(92), chr(34)
    found, seen = {}, 0
    vals = []
    try:
        import winreg
        path = "Software/Microsoft/Windows/CurrentVersion/Run".replace("/", chr(92))
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
        i = 0
        try:
            while True:
                try:
                    _n, v, _t = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                vals.append(str(v))
        finally:
            key.Close()
    except Exception:
        pass
    for d in _startup_dirs():
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if os.path.splitext(fn)[1].lower() in STARTUP_SCRIPT_EXT:
                t = _read_text_any(os.path.join(d, fn))
                if t:
                    vals.append(t)
    # ROOT 는 저장소 안(ecount)이라 그 부모가 곧 Documents 가 아니다 — 실측으로
    # 그렇게 잡았더니 실재하는 이웃 셋이 '0곳' 으로 나왔다. 경로에서 Documents
    # 세그먼트를 찾아 올라간다. 못 찾으면 사용자 폴더로 되돌린다.
    docs = ROOT
    while docs and os.path.basename(docs).lower() != 'documents':
        up = os.path.dirname(docs)
        if up == docs:
            docs = os.path.join(os.environ.get('USERPROFILE', ''), 'Documents')
            break
        docs = up
    parent = os.path.dirname(docs)
    for v in vals:
        seen += 1
        for m in re.finditer(re.escape(docs) + r"[" + re.escape(BS) + r"/]([^" + re.escape(BS) + r"/" + Q + r"]+)", v, re.I):
            folder = os.path.join(docs, m.group(1))
            mine = os.path.normcase(ROOT).startswith(os.path.normcase(folder))
            if os.path.isdir(folder) and not mine:
                found[os.path.normcase(folder)] = folder
    _ = parent
    why = dict((k, "자동으로 돈다(자동실행·시작폴더가 가리킨다)") for k in found)
    # ★ 두 번째 근거 — 세션이 최근에 연 저장소(2026-08-21).
    #   자동으로 돌지 않아도 세션이 매일 고치는 저장소면 거기서 창이 뜬다.
    sess, sess_why = _session_cwds()
    me = os.path.normcase(ROOT)
    for folder in sess:
        key = os.path.normcase(folder)
        # 내 저장소(그 위·아래 포함)는 뺀다 — ROOT 로 이미 센다.
        if me.startswith(key) or key.startswith(me):
            continue
        if key in found:
            why[key] = why[key] + " · 세션도 최근에 열었다"
            continue
        found[key] = folder
        why[key] = "세션이 최근 %d일 안에 열었다" % NEIGHBOR_DAYS
    dirs = sorted(found.values())
    if with_why:
        return [(d, why.get(os.path.normcase(d), "모름")) for d in dirs], seen, sess_why
    return dirs, seen


def _hook_files():
    """훅이 실릴 수 있는 자리 셋. **계정마다 다른 것은 마지막 하나**다.

    프로젝트 설정 둘은 폴더를 따라오므로 계정이 바뀌어도 그대로 물려지고,
    사용자 설정은 계정마다 다르다 — 그래서 셋을 다 본다."""
    base = os.path.dirname(ROOT)
    home = os.path.expanduser('~')
    return [
        (os.path.join(base, '.claude', 'settings.json'), '프로젝트'),
        (os.path.join(base, '.claude', 'settings.local.json'), '프로젝트(로컬)'),
        (os.path.join(home, '.claude', 'settings.json'), '사용자(계정)'),
    ]


def hooks(files=None):
    """세션 훅이 **창 뜨는 실행기**로 걸렸나 → `(목록, 왜못함, 본개수)`.

    ★ 왜 여기가 구멍이었나 (2026-08-21 형님 지시 "어떤 계정 어떤 세션에서 진행해도"
      팝업은 백그라운드로") — `[272]` 는 **소스**를, `live()` 는 **예약 작업**을,
      `autorun()` 은 **로그인 항목**을 본다. 그런데 훅은 **`.claude/settings.json`** 에
      있어 셋 중 어디에도 안 잡혔다. 실측 2026-08-21: 훅 여섯이 전부 `python`(콘솔)
      이었고 그중 `PostToolUse` 는 **도구를 부를 때마다** 돈다 — 사람이 한 응답을 받는
      동안 창이 수십 번 번쩍이는데 두 감사기는 나란히 `0곳` 이라 말했다(`[169]`).
    ★ 창을 없앤다고 기능을 죽이면 안 된다 — 실측으로 `pythonw` 는 **파이프 stdout 을
      그대로 살린다**(콘솔창핸들 0 · stdout 살아있음). 그래서 훅이 대화에 내보내는
      한 줄은 하나도 안 잃는다.
    ★ 판정은 `exe_verdict` 를 빌린다(`[162]`) — 여기서 새로 만들면 예약 작업·로그인
      항목과 갈리고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다.
    ★ **'훅이 없다'와 '설정을 못 읽었다'를 가른다**(`[169]`) — 본개수가 그 구별이다.
      깨진 설정 하나 때문에 나머지 파일을 안 보지도 않는다.
    ★ 읽기만 한다 — 고치는 것은 사람이 정한다(`typo_watch` 와 같은 자리)."""
    out, seen, bad = [], 0, []
    for path, 어디 in (files if files is not None else _hook_files()):
        if not os.path.exists(path):
            continue
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                conf = json.load(fh)
        except Exception as exc:
            bad.append('%s(%s)' % (os.path.basename(path), type(exc).__name__))
            continue
        for 사건, groups in (conf.get('hooks') or {}).items():
            for g in (groups or []):
                for h in ((g or {}).get('hooks') or []):
                    cmd = (h or {}).get('command')
                    if not cmd:
                        continue
                    seen += 1
                    exe = _first_token(cmd)
                    판정 = exe_verdict(exe)
                    if 판정 != '조용':
                        out.append((어디, 사건, os.path.basename(exe), 판정))
    why = ''
    if bad and not seen:
        why = '설정을 못 읽었다(%s)' % ', '.join(bad)
    elif bad:
        why = ''
        out.append(('설정', '읽기실패', ', '.join(bad), '확인못함'))
    return out, why, seen

def scan(root=None):
    """(파일, 줄, 무엇, 콘솔낱말) 목록 — 창이 달릴 수 있는 자리.

    ★ `root` 를 받는 이유 (2026-08-21): 형님 지시가 "어떤 계정 어떤 세션에서
      진행해도" 다. 이 감사기가 제 저장소에만 매여 있으면 다른 세션이 고치는
      프로젝트는 통째로 눈 밖이다. 기본값은 그대로 이 저장소라 동작이 안 바뀐다."""
    root = root or ROOT
    # ★ 두 번 훑는다: 먼저 **헬퍼 본문**을 읽어 무엇이 안전한지 정하고, 그다음 호출을
    #   본다.  한 번에 하면 파일 순서에 따라 **뒤에 정의된 헬퍼를 못 알아본다** —
    #   그러면 같은 코드가 훑는 순서에 따라 다른 답을 낸다.
    trees = []
    for base, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for fn in files:
            if not fn.endswith(".py"):
                continue
            p = os.path.join(base, fn)
            rel = os.path.relpath(p, root).replace(chr(92), "/")
            if rel.startswith("tools/window_audit"):
                continue
            try:
                trees.append((rel, ast.parse(open(p, encoding="utf-8").read())))
            except (OSError, SyntaxError, ValueError, RecursionError):
                continue
    helpers = safe_helpers(trees)
    bad = []
    for rel, tree in trees:
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            f_ = node.func
            if not isinstance(f_, ast.Attribute) or f_.attr not in SPAWNERS:
                continue
            if not (isinstance(f_.value, ast.Name) and f_.value.id == "subprocess"):
                continue
            if _has_no_window(node, helpers) or _flag_var_ok(node, tree):
                continue
            bad.append((rel, node.lineno, f_.attr, _looks_console(node)))
    bad.sort()
    return bad




def split(rows=None, root=None):
    """(확실히 콘솔, 모름) 으로 가른다 — 세는 쪽이 둘을 구별해 말할 수 있게."""
    rows = scan(root) if rows is None else rows
    sure = [r for r in rows if r[3] != "?"]
    unknown = [r for r in rows if r[3] == "?"]
    return sure, unknown


def main(argv=None):
    ap = argparse.ArgumentParser(description="창이 달릴 수 있는 자식 실행 자리")
    ap.add_argument("--count", action="store_true", help="건수만")
    ap.add_argument("--autorun", action="store_true",
                    help="로그인 자동실행(HKCU Run)에 걸린 이 프로젝트 항목")
    ap.add_argument("--hooks", action="store_true",
                    help="세션 훅(.claude/settings.json)이 창 뜨는 실행기인가")
    ap.add_argument("--live", action="store_true",
                    help="살아 있는 예약 작업의 실행기(회차가 써 둔 것을 읽는다)")
    ap.add_argument("--startup", action="store_true",
                    help="시작 폴더의 스크립트 자동실행이 창을 다나(.vbs 안까지 읽는다)")
    ap.add_argument("--neighbors", action="store_true",
                    help="이 PC 에서 자동으로 도는 다른 프로젝트 폴더 목록")
    ap.add_argument("--root", default=None,
                    help="소스를 훑을 폴더(기본: 이 저장소) — 다른 세션 저장소도 잰다")
    a = ap.parse_args(argv)
    if a.autorun:
        rows, why, seen = autorun()
        if why:
            print("로그인 자동실행을 **확인 못 했다** — %s" % why)
            return 0
        if not rows:
            print("로그인 자동실행 — 이 프로젝트 항목 %d개 · 창 뜨는 것 0곳" % seen)
            return 0
        print("★ 로그인 자동실행에 창 뜨는 실행기 %d곳" % len(rows))
        for name, exe, verdict in rows:
            print("  %s — %s (%s)" % (name, os.path.basename(exe), verdict))
        return 0
    if a.hooks:
        rows, why, seen = hooks()
        if why:
            print("세션 훅을 **확인 못 했다** — %s" % why)
            return 0
        if not rows:
            print("세션 훅 %d개 · 창 뜨는 것 0곳 (전부 pythonw·wscript)" % seen)
            return 0
        print("★ 창 뜨는 실행기로 걸린 세션 훅 %d곳 (전체 %d개)" % (len(rows), seen))
        for 어디, 사건, exe, 판정 in rows:
            print("  %-14s %-18s %-14s %s" % (어디, 사건, exe, 판정))
        return 0
    if a.startup:
        rows, why, seen = startup()
        if why:
            print("시작 폴더를 **확인 못 했다** — %s" % why)
            return 0
        if not rows:
            print("시작 폴더 스크립트 자동실행 %d개 · 창 뜨는 것 0곳" % seen)
            return 0
        print("★ 시작 폴더에 창 뜨는 자동실행 %d곳 (전체 %d개)" % (len(rows), seen))
        for fn, 판정, 근거 in rows:
            print("  %-34s %-8s %s" % (fn, 판정, 근거))
        return 0
    if a.neighbors:
        rows, _seen, why = neighbors(with_why=True)
        if why:
            print("세션이 연 저장소를 **확인 못 했다** — %s" % why)
        if not rows:
            print("이 PC 에서 볼 다른 프로젝트 0곳")
            return 0
        # ★ '자동으로 도는' 이라고만 적으면 틀린 말이 된다 — 근거가 둘이다([169]).
        print("이 PC 의 다른 프로젝트 %d곳 — 각각 --root 로 잰다" % len(rows))
        for d, w in rows:
            print("  %s" % d)
            print("      근거: %s" % w)
        return 0
    if a.live:
        rows, 왜못함 = live()
        if 왜못함:
            print("예약 작업 실행기를 **확인 못 했다** — %s" % 왜못함)
            return 0
        if not rows:
            print("창 뜨는 실행기로 등록된 예약 작업: 0곳 (전부 pythonw·wscript)")
            return 0
        print("★ 창 뜨는 실행기로 등록된 예약 작업 %d곳" % len(rows))
        for 작업, exe, 판정 in rows:
            print("  %-34s %-14s %s" % (작업, 판정, exe))
        return 0
    sure, unknown = split(root=a.root)
    if a.count:
        print(len(sure) + len(unknown))
        return 0
    if not sure and not unknown:
        print("창이 달릴 수 있는 자리: 0곳 (전부 깃발이 있다)")
        return 0
    if sure:
        print("★ 콘솔 exe 를 깃발 없이 띄우는 자리 %d곳" % len(sure))
        for rel, line, what, exe in sure:
            print("  %-44s L%-6d subprocess.%-12s %s" % (rel, line, what, exe))
    if unknown:
        # 지목이 아니라 보고다 — 무엇을 띄우는지 코드만 봐서는 모르는 자리.
        print("무엇을 띄우는지 못 읽은 자리 %d곳 — **0곳이라 말하면 안 되는 자리**"
              % len(unknown))
        for rel, line, what, _ in unknown:
            print("  %-44s L%-6d subprocess.%s" % (rel, line, what))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
