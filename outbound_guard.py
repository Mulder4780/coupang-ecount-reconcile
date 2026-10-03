# -*- coding: utf-8 -*-
"""
outbound_guard.py — 나가는 산출물에 **ARTIS 가 섞였는지** 본다
===============================================================================
형님 지시(2026-10-03): *"나가는 산출물에 'ARTIS·아티스·artis-elevator'가 들어가면
걸러내는 검사를 만들 수 있으면 만들어라."*

ARTIS 는 형님 **개인 사업자**이고 유니버셜리프트와 다른 사업이다. 회사 산출물
(보고서·엑셀·매뉴얼·메일·앱 화면·회사 NAS(Z:)·회사 문서)에 그 이름이 들어가면
**회사 자료에 개인 사업이 노출된다.** Z: 는 모든 PC·사람이 같이 보는 자리라
한 번 올라가면 지워도 본 사람이 남는다 — 되돌릴 수 없는 쪽이다.

## 세 가지만 지킨다

★ **지우지 않고 알린다**([169]). 자동으로 고치면 멀쩡한 글자가 조용히 사라진다 —
  영문 낱말 안에 우연히 든 글자, 사람 이름, 남의 회사 이름일 수 있다.
  **어디 몇 줄째에 걸렸는지 짚고 사람이 정한다.**

★ **이 지시문 자신과 기억·사고기록은 안 잰다**([170]). 그 글들은 *"ARTIS 를 넣지
  마라"* 를 적기 위해 그 낱말을 쓴다. 그것까지 걸면 경보가 상시가 되어 뜻을 잃는다 —
  `session_scope._is_echo` 가 지시문 사본을 근거에서 빼는 것과 같은 자리다.

★ **못 읽으면 '깨끗'이라 하지 않는다**([169]). 읽다 실패하면 `못읽음` 으로 적는다.
  '못 본 것'을 '이상 없음'으로 세면 그때가 바로 새는 순간이다.

## 사람이 쓰는 법
  python outbound_guard.py --check <파일…>        # 파일을 본다
  python outbound_guard.py --check <폴더> -r      # 폴더를 훑는다
  python outbound_guard.py --text "보낼 글"        # 글자를 바로 본다
종료코드: **0 깨끗 · 3 걸림 · 2 쓸 파일을 못 받음**

## 코드에서 부를 때
  import outbound_guard
  걸림 = outbound_guard.scan_text(글)       # [(줄번호, 걸린말, 그 줄), …]
  if 걸림: ...                           # 지우지 말고 사람에게 알린다
"""
import argparse
import io
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))

# ★ 금지어 목록은 **여기 한 곳**이다([162]). 늘릴 일이 생기면 여기만 고친다.
#   영문은 대소문자를 안 가린다. '아티스'는 한글이라 그대로 본다.
WORDS = ("ARTIS", "아티스", "artis-elevator", "artis_elevator", "artis.co", "artiselevator")

# ★ 재지 않는 글 — 이것들은 '넣지 마라'를 적느라 그 낱말을 쓴다([170]).
SKIP_NAMES = ("CLAUDE.md", "AGENTS.md", "INCIDENTS.md", "outbound_guard.py",
              "no-artis-in-universal-work.md", "MEMORY.md", "session_scope.py")
SKIP_DIRS = (".git", "__pycache__", "node_modules", ".claude")

# ★ 영문 낱말 한가운데 우연히 든 것은 안 센다 — 'cartridge' 가 'artis' 를 품지는
#   않지만, 같은 모양의 사고를 막으려고 경계를 둔다([172] 틀린 지목).
_PAT = re.compile("|".join(
    (r"(?<![A-Za-z0-9])" + re.escape(w) + r"(?![A-Za-z0-9])") if w.isascii() else re.escape(w)
    for w in WORDS), re.IGNORECASE)


def skip_path(path):
    """재지 않는 글인가 — 지시문 사본·기억·이 파일 자신."""
    n = os.path.basename(path)
    if n in SKIP_NAMES:
        return True
    low = path.replace("\\", "/").lower()
    if "/memory/" in low or "/.claude/" in low:
        return True
    return False


def scan_text(text):
    """걸린 자리 목록 — [(줄번호, 걸린말, 그 줄 일부), …]. 없으면 빈 목록."""
    out = []
    for i, line in enumerate(str(text or "").split("\n"), 1):
        for m in _PAT.finditer(line):
            out.append((i, m.group(0), line.strip()[:120]))
    return out


def scan_file(path):
    """돌려주는 것: ('깨끗'|'걸림'|'못읽음'|'건너뜀', 자세히)."""
    if skip_path(path):
        return "건너뜀", "지시문·기억 사본이라 안 잰다"
    try:
        with io.open(path, encoding="utf-8", errors="strict") as fh:
            t = fh.read()
    except UnicodeDecodeError:
        # 이진 파일(xlsx·docx·이미지)은 글자로 못 읽는다 — **'깨끗'이라 하지 않는다**.
        return "못읽음", "글자로 못 읽는 파일이다(이진) — 안을 못 봤다"
    except Exception as e:
        return "못읽음", str(e)[:120]
    걸림 = scan_text(t)
    return ("걸림", 걸림) if 걸림 else ("깨끗", "")


def walk(paths, recursive=False):
    for p in paths:
        if os.path.isdir(p):
            if not recursive:
                continue
            for dp, dn, fn in os.walk(p):
                dn[:] = [d for d in dn if d not in SKIP_DIRS]
                for f in fn:
                    yield os.path.join(dp, f)
        else:
            yield p


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="나가는 산출물에 ARTIS 가 섞였는지 본다(지우지 않고 알린다)")
    ap.add_argument("--check", nargs="*", default=[], metavar="경로")
    ap.add_argument("-r", "--recursive", action="store_true", help="폴더를 훑는다")
    ap.add_argument("--text", help="글자를 바로 본다")
    ap.add_argument("--quiet", action="store_true", help="걸린 것만 찍는다")
    a = ap.parse_args(argv)

    if a.text is not None:
        걸림 = scan_text(a.text)
        if not 걸림:
            print("깨끗 — 금지어 없음")
            return 0
        for n, w, line in 걸림:
            print("  걸림 %d줄 [%s] %s" % (n, w, line))
        print("★ 지우지 않았다 — 어디가 걸렸는지만 알린다. 뺄지는 사람이 정한다.")
        return 3

    if not a.check:
        print("쓸 파일이나 --text 를 주십시오. 보기:")
        print('  python outbound_guard.py --check reports/보고서.md')
        print('  python outbound_guard.py --text "보낼 글"')
        return 2

    셈 = {"깨끗": 0, "걸림": 0, "못읽음": 0, "건너뜀": 0}
    걸린파일, 못읽은파일 = [], []
    for p in walk(a.check, a.recursive):
        갈래, 자세히 = scan_file(p)
        셈[갈래] += 1
        if 갈래 == "걸림":
            걸린파일.append((p, 자세히))
        elif 갈래 == "못읽음":
            못읽은파일.append((p, 자세히))

    for p, 자세히 in 걸린파일:
        print("★ %s" % p)
        for n, w, line in 자세히[:8]:
            print("    %d줄 [%s] %s" % (n, w, line))
        if len(자세히) > 8:
            print("    … 그 밖 %d곳" % (len(자세히) - 8))
    if 못읽은파일 and not a.quiet:
        print()
        print("※ 안을 못 본 파일 %d개 — **'깨끗'이 아니라 '모름'이다**:" % len(못읽은파일))
        for p, why in 못읽은파일[:10]:
            print("    %s — %s" % (p, why))
        if len(못읽은파일) > 10:
            print("    … 그 밖 %d개" % (len(못읽은파일) - 10))

    print()
    print("깨끗 %d · **걸림 %d** · 못읽음 %d · 건너뜀 %d"
          % (셈["깨끗"], 셈["걸림"], 셈["못읽음"], 셈["건너뜀"]))
    if 걸린파일:
        print("★ 지우지 않았다 — 뺄지는 사람이 정한다([169]).")
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
