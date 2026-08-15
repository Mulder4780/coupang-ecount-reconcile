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
import ast
import os
import sys

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


def _has_no_window(call: ast.Call) -> bool:
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
        if kw.arg is None and "background_popen_kwargs" in ast.unparse(kw.value):
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


def scan():
    """(파일, 줄, 무엇, 콘솔낱말) 목록 — 창이 달릴 수 있는 자리."""
    bad = []
    for base, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for fn in files:
            if not fn.endswith(".py"):
                continue
            p = os.path.join(base, fn)
            rel = os.path.relpath(p, ROOT).replace("\\", "/")
            if rel.startswith("tools/window_audit"):
                continue
            try:
                tree = ast.parse(open(p, encoding="utf-8").read())
            except (OSError, SyntaxError):
                continue
            for node in ast.walk(tree):
                if not isinstance(node, ast.Call):
                    continue
                f = node.func
                if not isinstance(f, ast.Attribute) or f.attr not in SPAWNERS:
                    continue
                if not (isinstance(f.value, ast.Name) and f.value.id == "subprocess"):
                    continue
                if _has_no_window(node):
                    continue
                # ★ **모르는 것을 조용히 넘기지 않는다** (2026-08-14 실사고).
                #   전에는 무엇을 띄우는지 모르면 `continue` 로 건너뛰고 "0곳"이라
                #   말했다. 그런데 건너뛴 자리가 **30곳**이었고 그 안에
                #   `tailscale_serve.run()` 이 있었다 — `server_guard` 가 **60초마다**
                #   부르는 자리다. 즉 화면에는 1분마다 검은 창이 떴는데 계기는
                #   "0곳"을 확언했다. 0 이 '없다'인지 '안 봤다'인지 안 가르면
                #   계기 자신이 눈이 먼다([169]).
                #   그래서 이제 **모름도 돌려준다**. 지목이 아니라 '못 봤다'는 보고다.
                bad.append((rel, node.lineno, f.attr, _looks_console(node)))
    bad.sort()
    return bad


def split(rows=None):
    """(확실히 콘솔, 모름) 으로 가른다 — 세는 쪽이 둘을 구별해 말할 수 있게."""
    rows = scan() if rows is None else rows
    sure = [r for r in rows if r[3] != "?"]
    unknown = [r for r in rows if r[3] == "?"]
    return sure, unknown


def main(argv=None):
    ap = argparse.ArgumentParser(description="창이 달릴 수 있는 자식 실행 자리")
    ap.add_argument("--count", action="store_true", help="건수만")
    a = ap.parse_args(argv)
    sure, unknown = split()
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
