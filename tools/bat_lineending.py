# -*- coding: utf-8 -*-
"""회차 `.bat` 의 줄끝을 지킨다 — LF 로 다시 써지면 cmd 가 줄 중간에서 재개한다 (2026-08-20).

★ 무엇이 있었나 (분담판 [166] 실측)
  `쿠팡업무_원본자료자동정리` 가 09:35 에 **exit 9009**(0x2331 '명령을 찾을 수 없음')
  로 죽고 `reports/source_organizer.log` 에 **한 줄도 안 남았다** — python 이 아예
  안 떴다는 뜻이라 `[324]` 가 만든 자국도 당연히 없다. 원인은 그 전날 도구 편집이
  그 bat 을 **LF 전용**으로 다시 쓴 것이었다.

★ 가르는 실측 (진짜 회차는 안 돌리고 python 줄만 `cmd /c exit 7` 로 바꾼 사본으로 쟀다)
  ① ASCII bat = 7   ② **LF + 한글 rem = 49(엉뚱)**   ③ 같은 내용 CRLF = 7
  ④ LF 인데 한글 rem 제거 = 7   ⑤ LF 인데 rem 을 ASCII 로 = 7   ⑥ 고친 뒤 실제 파일 = 7
  즉 **LF 만으로도, 한글만으로도 안 깨진다 — 둘이 같이 있을 때만 깨진다.**
  cmd 는 배치파일을 바이트 오프셋으로 되읽는데, 멀티바이트(CP949) 해석 길이와
  어긋나 **줄 중간에서 재개**한다. 그래서 조각(예: `on_organizer.log 2>&1`)이
  명령으로 실행돼 9009 가 된다.

★ 위험한 것은 '늘 깨진다'가 아니라 **'운에 달렸다'** 는 점이다 — 그날 `daily_run.bat`
  도 같은 상태였는데 멀쩡히 돌았다. **주석 한 줄만 고쳐도 도는 회차가 9009 로 뒤집힌다.**
  그래서 갈래를 둘로 가르되 **둘 다 실패**로 본다:
  · `깨짐확인` — bare LF + 비ASCII. 실측으로 깨진 그 모양이다.
  · `위험`     — bare LF 인데 아직 ASCII 뿐. **오늘은 돈다.** 그러나 한글 주석 한 줄이면
                 `깨짐확인` 이 된다. 되돌리는 값은 CRLF 한 번이고 부딪히는 값은
                 **회차 하나가 통째로 안 도는 것**이다.
  갈래를 뭉치지 않는 이유는 조치가 아니라 **설명**이 다르기 때문이다(`[289]`) —
  '오늘 깨져 있다' 와 '내일 깨질 수 있다' 는 사람이 알아야 할 서로 다른 사실이다.

★ **못 읽으면 '안전' 이라 하지 않는다**(`[169]`) — `못읽음` 으로 세고 그것도 실패다.
  계기 자신이 눈먼 채 '이상 없음' 을 말하는 것이 이 프로젝트가 반복해 당한 모양이다.

★ 저장소는 원래 맞다 — git 은 이 파일들을 LF 로 갖고 `core.autocrlf=true` 가 체크아웃에서
  CRLF 를 만든다. 어긋나는 것은 **작업폴더뿐**이고, 어긋나게 만드는 것은 도구 편집이다.
  그러므로 `--fix` 는 **내용 바이트를 한 톨도 안 바꾸고 줄끝만** 되돌린다.

쓰기:  python tools/bat_lineending.py           # 지금 상태(읽기 전용) · 걸리면 exit 1
      python tools/bat_lineending.py --fix     # bare LF -> CRLF (사람이 부를 때만)
"""
from __future__ import annotations

import argparse
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:                                    # pythonw 에서는 stdout 이 None 이다
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BROKEN = "깨짐확인"      # bare LF + 비ASCII — 실측으로 cmd 가 줄 중간에서 재개한다
RISKY = "위험"           # bare LF 인데 ASCII 뿐 — 오늘은 돌지만 한글 한 줄이면 뒤집힌다
SAFE = "안전"
UNREAD = "못읽음"
BAD = (BROKEN, RISKY, UNREAD)                        # 셋 다 실패다


def judge(data):
    """바이트만 보고 갈래를 정한다. 파일 이름도 위치도 안 본다."""
    lf = data.count(b"\n")
    crlf = data.count(b"\r\n")
    bare = lf - crlf                                 # \r 없이 혼자 선 LF
    nonascii = any(b > 127 for b in data)
    if bare <= 0:
        kind = SAFE
    elif nonascii:
        kind = BROKEN
    else:
        kind = RISKY
    return kind, {"crlf": crlf, "bare_lf": bare, "비ASCII": nonascii}


def bat_files(root=None):
    """살아 있는 `.bat` 을 전부 모은다 — 스케줄러가 지금 부르는 것만 보지 않는다.

    ★ 오늘 안 불리는 bat 도 내일 불린다. 그리고 이 고장은 **조용하다** —
      로그가 한 줄도 안 남으므로 '안 불렸다' 와 '깨져서 안 떴다' 가 구별되지 않는다.
    """
    root = root or ROOT
    out = []
    for base, dirs, names in os.walk(root):
        dirs[:] = [d for d in dirs
                   if d not in (".git", "__pycache__", "node_modules") and not d.startswith(".")]
        for n in names:
            if n.lower().endswith(".bat"):
                out.append(os.path.join(base, n))
    return sorted(out)


def scan(root=None):
    rows = []
    for p in bat_files(root):
        try:
            data = open(p, "rb").read()
        except OSError as e:                          # 잠겨 있거나 권한이 없다
            rows.append({"path": p, "갈래": UNREAD, "왜": str(e)[:120]})
            continue
        kind, meta = judge(data)
        row = {"path": p, "갈래": kind}
        row.update(meta)
        rows.append(row)
    return rows


def fix(paths):
    """bare LF -> CRLF. **내용 바이트는 한 톨도 안 바꾼다.**

    원자적으로 바꾼다 — 반쯤 써진 bat 은 깨진 bat 보다 나쁘다.
    """
    done = []
    for p in paths:
        data = open(p, "rb").read()
        fixed = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
        if fixed == data:
            continue
        assert fixed.replace(b"\r\n", b"\n") == data.replace(b"\r\n", b"\n"), \
            "줄끝 말고 다른 것이 바뀌었다: " + p
        tmp = p + ".lineending.tmp"
        with open(tmp, "wb") as fh:
            fh.write(fixed)
        os.replace(tmp, p)
        done.append(p)
    return done


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fix", action="store_true", help="bare LF 를 CRLF 로 되돌린다")
    a = ap.parse_args()

    rows = scan()
    bad = [r for r in rows if r["갈래"] in BAD]

    if a.fix:
        targets = [r["path"] for r in bad if r["갈래"] in (BROKEN, RISKY)]
        done = fix(targets)
        print("고침 %d개 / 걸린 것 %d개" % (len(done), len(bad)))
        for p in done:
            print("  -", os.path.relpath(p, ROOT))
        rows = scan()
        bad = [r for r in rows if r["갈래"] in BAD]

    print("bat %d개 · 걸린 것 %d개" % (len(rows), len(bad)))
    for r in bad:
        print("  [%s] %s  (CRLF %s · 혼자선LF %s)"
              % (r["갈래"], os.path.relpath(r["path"], ROOT),
                 r.get("crlf", "?"), r.get("bare_lf", "?")))
    if not bad:
        print("  걸린 것 없음 — 모두 CRLF")
        return 0
    print("\n고치려면: python tools/bat_lineending.py --fix")
    return 1


if __name__ == "__main__":
    sys.exit(main())
