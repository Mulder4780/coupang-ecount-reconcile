# -*- coding: utf-8 -*-
"""font_switch.py — 앱 전체 글꼴을 한 줄로 바꾸고, 한 줄로 되돌린다

사용자 지시(2026-08-06): "폰트도 나눔고딕 말고 디자이너들과 사용자들이 선호하고
잘 보이는 폰트로 전체 변경해 / **나중에 내가 명령 내리면 다시 원래대로 할 수 있는
보호 장치도 마련해**."

왜 도구로 만드나
  글꼴은 네 파일에 흩어져 있다(업무센터 앱·폰 앱·공유 캘린더·폰 사본 생성기).
  "되돌려" 라는 말을 들었을 때 사람이나 AI 가 네 곳을 손으로 찾아 고치면
  **한 곳을 빠뜨린다** — 그 한 곳만 다른 글꼴로 남고, 그것도 화면을 열어 봐야 안다.
  그래서 되돌리기를 대화가 아니라 **명령 하나**로 만든다.

무엇을 건드리나
  각 파일의 `:root{ --font-ui-legacy: ...; --font-ui: ...; }` 두 줄뿐이다.
  `--font-ui-legacy` 는 **바꾸지 않는다** — 그게 원래 값을 기억하는 자리다.
  본문·캔버스·인쇄가 전부 `var(--font-ui)` 를 보므로 이 한 줄이면 전부 따라온다.

쓰기
  python webapp/font_switch.py             # 지금 어느 글꼴인가 (읽기만)
  python webapp/font_switch.py --legacy    # 예전 글꼴(나눔고딕)로 전부 되돌린다
  python webapp/font_switch.py --modern    # 다시 기본 글꼴로

검증 [126] 이 이 구조와 왕복(modern→legacy→modern)을 지킨다.
"""
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 글꼴을 정하는 파일들. 하나라도 빠지면 그 화면만 다른 글꼴로 남는다.
FILES = (
    os.path.join("webapp", "index.html"),      # 업무센터 앱(본체)
    os.path.join("docs", "app.html"),          # 폰 앱
    os.path.join("docs", "cal.html"),          # 공유 캘린더
    "mobile_snapshot.py",                      # 폰 사본 생성기(HTML 을 만들어 낸다)
)

# 기본값 — 각 기기가 제 UI 에 쓰는 글꼴을 먼저 쓴다(내려받을 것이 없다).
#
# ★ 순서가 곧 결과다 (2026-08-07 사용자 지적: "글씨가 잘 안보이는게 많아 /
#   아이폰이나 갤럭시처럼 눈에 잘 보이는 폰트로").
#   예전 순서는 …"Segoe UI","Malgun Gothic",Roboto,"Noto Sans KR"… 였다. 그런데
#   **Segoe UI 에는 한글이 없다.** 브라우저는 글자 하나하나마다 스택을 걸어 내려가며
#   그 글자가 있는 첫 글꼴을 쓰므로, 한글은 언제나 `맑은 고딕`에서 멈췄다 —
#   Noto Sans KR 은 뒤에 적혀 있어 영원히 차례가 오지 않았다. 실측(2026-08-07,
#   이 PC): Pretendard **미설치**, Noto Sans KR **설치됨**, 실제 렌더 = 맑은 고딕.
#   맑은 고딕은 12~13px 에서 획이 얇아 흐릿하다 — 그게 "잘 안 보인다"의 정체다.
#   그래서 **폰이 쓰는 한글 글꼴을 맑은 고딕 앞으로** 올린다:
#     아이폰 → Apple SD Gothic Neo · 갤럭시/안드로이드 → Roboto + Noto Sans KR(본고딕)
#     윈도우 → 라틴은 Segoe UI, 한글은 Noto Sans KR
#   맑은 고딕은 **지우지 않고 맨 뒤**에 남긴다(Noto 가 없는 PC 의 마지막 보루).
MODERN = ('-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Pretendard Variable",'
          'Pretendard,"Segoe UI Variable Text","Segoe UI",Roboto,'
          '"Noto Sans KR","Source Han Sans KR","본고딕","Malgun Gothic","맑은 고딕",'
          '"Helvetica Neue",Arial,sans-serif')

ACTIVE = re.compile(r"(--font-ui\s*:\s*)(.+?)(;)")
LEGACY = re.compile(r"--font-ui-legacy\s*:\s*(.+?);")


def _read(path):
    # 줄바꿈도 원본의 일부다. ``newline=None``(기본값)으로 읽으면 Windows의
    # CRLF가 LF로 정규화되어 글꼴 한 줄만 바꿔도 파일 전체가 달라진다.
    # 읽기·쓰기를 모두 ``newline=""``로 맞춰 왕복 시 바이트를 보존한다.
    with open(os.path.join(ROOT, path), encoding="utf-8", newline="") as f:
        return f.read()


def _write(path, text):
    with open(os.path.join(ROOT, path), "w", encoding="utf-8", newline="") as f:
        f.write(text)


def _active(text):
    """지금 쓰이는 글꼴 값. `--font-ui-legacy` 를 잡지 않도록 그 줄은 지우고 찾는다."""
    body = re.sub(r"--font-ui-legacy\s*:\s*.+?;", "", text)
    m = ACTIVE.search(body)
    return m.group(2).strip() if m else ""


def state():
    """파일마다 지금 무엇을 쓰나. 하나라도 어긋나면 그것이 곧 사고다."""
    out = []
    for rel in FILES:
        text = _read(rel)
        leg = LEGACY.search(text)
        act = _active(text)
        if not leg or not act:
            out.append({"파일": rel, "상태": "글꼴 변수를 못 찾음", "지금": "", "예전": ""})
            continue
        legv = leg.group(1).strip()
        out.append({"파일": rel,
                    "상태": "예전" if act == "var(--font-ui-legacy)" or act == legv else "기본",
                    "지금": act, "예전": legv})
    return out


def apply(mode):
    """mode: 'legacy' | 'modern'. 바뀐 파일 목록을 돌려준다."""
    if mode not in ("legacy", "modern"):
        raise ValueError(mode)
    changed = []
    for rel in FILES:
        text = _read(rel)
        leg = LEGACY.search(text)
        if not leg:
            raise RuntimeError("%s 에 --font-ui-legacy 가 없다 — 되돌릴 자리를 잃었다" % rel)
        want = "var(--font-ui-legacy)" if mode == "legacy" else MODERN
        # legacy 줄은 건드리지 않는다: 그 줄을 잠시 치워 두고 활성 줄만 바꾼다
        holder = "\x00LEGACY\x00"
        stashed = text.replace(leg.group(0), holder, 1)
        if _active(text) == want:
            continue
        new = ACTIVE.sub(lambda m: m.group(1) + want + m.group(3), stashed, count=1)
        new = new.replace(holder, leg.group(0), 1)
        if new != text:
            _write(rel, new)
            changed.append(rel)
    return changed


def _print():
    rows = state()
    for r in rows:
        mark = "예전(나눔고딕)" if r["상태"] == "예전" else r["상태"]
        print("  %-24s %s" % (r["파일"], mark))
        if r["지금"]:
            print("       지금: %s" % r["지금"][:88])
    kinds = {r["상태"] for r in rows}
    if len(kinds) > 1:
        print("\n★ 파일마다 글꼴이 다르다 — 한 화면만 다른 글꼴로 보인다.")
        print("  `python webapp/font_switch.py --modern` 또는 `--legacy` 로 맞춘다.")
    else:
        print("\n네 파일 모두 같은 글꼴을 쓴다.")


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else ""
    if arg in ("--legacy", "--modern"):
        mode = arg[2:]
        changed = apply(mode)
        label = "예전 글꼴(나눔고딕)" if mode == "legacy" else "기본 글꼴"
        print("%s 로 바꿨습니다 — 고친 파일 %d개%s"
              % (label, len(changed), (": " + ", ".join(changed)) if changed else " (이미 그 상태였다)"))
        print("폰 사본은 다음 게시 회차에 반영된다(mobile_snapshot.py 가 만들어 낸다).")
        return 0
    if arg in ("-h", "--help"):
        print(__doc__)
        return 0
    _print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
