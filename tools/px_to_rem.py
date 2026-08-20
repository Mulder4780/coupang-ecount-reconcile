# -*- coding: utf-8 -*-
"""글자 크기를 사람이 조절할 수 있게 — CSS 의 `font-size:Npx` 를 rem 으로 옮긴다.

2026-08-20 형님 지시(분담판 [72] · 선택지 넷 중 **px→rem 스크립트 변환**).

★ 왜 `body{font-size}` 하나로는 안 되나 — 실측(webapp/index.html):
    CSS 안 `font-size:Npx` **509곳** · 인라인 `style="…font-size"` **102곳** ·
    `rem`·`em`·`var()` **0곳**.
  px 는 `html` 에서 안 물려받으므로 `body` 를 키워도 **611곳은 그대로**다.
  글자 일부만 커지고 화면이 어긋난 채로 커진다 — 반쪽이 아니라 **깨진** 결과다.

★ 왜 `zoom` 이 아닌가 — 같은 파일에 `position:fixed` **14곳** · `sticky` **7곳** ·
  `@media max-width` **49개**가 있다. zoom 은 고정 요소와 부딪히고 미디어쿼리는
  다시 안 걸린다(사이드바·시트가 깨진다).

★ **배율 1 에서는 한 픽셀도 안 바뀐다.** `13px → 0.8125rem` 이고 뿌리가
  `calc(16px * var(--ui-scale,1))` 이므로 scale=1 이면 정확히 13px 이다.
  이 도구가 지키는 것이 그것이다 — 되돌려 곱해 **원래 px 가 안 나오면 안 바꾼다**.

★ **주석 안은 안 건드린다.** 이 파일의 주석에는 사고 기록이 잔뜩 들어 있고
  그중 `font-size:13px` 같은 글이 실제로 있다 — 바꾸면 기록이 뒤틀린다
  (이 프로젝트가 다섯 번 밟은 자리다 · [301]⑨·[302]·[309]·[332]·[339]).

★ **인라인 `style="…"` 과 JS 템플릿은 안 건드린다**(실측 102곳). 그것들은 화면
  코드 안에 있어 규칙이 다르고, 여기서 같이 바꾸면 무엇이 왜 바뀌었는지
  한 커밋에서 갈라 볼 수 없다. 남은 것은 **숫자로 말한다**([169]).

쓰는 법
    python tools/px_to_rem.py              # 무엇이 바뀌는지만 (아무것도 안 고친다)
    python tools/px_to_rem.py --apply      # 실제로 바꾼다
"""
import io
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):      # 무인 회차는 sys.stdout 이 None 이다([235])
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGETS = (os.path.join("webapp", "index.html"),)

BASE = 16.0                      # 뿌리 글자 크기(px). `html` 규칙과 **같은 값**이어야 한다.
FS = re.compile(r"(font-size\s*:\s*)([0-9]*\.?[0-9]+)px")
# ★ `font:` 축약형도 같이 본다(실측 35곳 · `font:800 12px inherit` 모양).
#   안 바꾸면 배율을 올렸을 때 **거기만 안 커진다** — 화면이 어긋난 채로 커진다.
#   축약형에서 첫 px 가 곧 글자 크기다(`12px/1.6` 의 1.6 은 줄높이이고 단위가 없다).
SHORT = re.compile(r"((?<!-)\bfont\s*:\s*[^;}]*?)([0-9]*\.?[0-9]+)px")
COMMENT = re.compile(r"/\*.*?\*/", re.S)


def style_spans(text):
    """`<style>` 덩어리들의 (시작, 끝) — 본문·JS 는 건드리지 않기 위해서다.

    ⚠ **여는 태그부터 세면 안 된다** — 이 파일에는 주석 안에도 `<style>` 이라는
      글자가 있다. 닫는 태그마다 **가장 가까운 여는 태그**를 짝짓는다([142] 실측).
    """
    out = []
    for m in re.finditer(r"</style>", text):
        i = text.rfind("<style", 0, m.start())
        if i < 0:
            continue
        j = text.find(">", i)
        if j < 0 or j > m.start():
            continue
        out.append((j + 1, m.start()))
    return out


def comment_spans(css):
    return [(m.start(), m.end()) for m in COMMENT.finditer(css)]


def to_rem(px):
    """px → rem. **되돌려 곱해 원래 값이 안 나오면 None** — 억지로 안 바꾼다."""
    rem = px / BASE
    s = ("%.6f" % rem).rstrip("0").rstrip(".")
    if not s:
        return None
    if abs(float(s) * BASE - px) > 1e-9:
        return None
    return s


def convert(text):
    """(새 본문, 바꾼 곳, 건너뛴 곳) — 건너뛴 것은 **왜인지 함께** 돌려준다."""
    changed, skipped = [], []
    pieces, last = [], 0
    for a, b in style_spans(text):
        css = text[a:b]
        holes = comment_spans(css)

        def in_comment(pos):
            return any(x <= pos < y for x, y in holes)

        out, prev = [], 0
        marks = sorted(list(FS.finditer(css)) + list(SHORT.finditer(css)),
                       key=lambda m: m.start())
        seen = set()
        for m in marks:
            if m.start(2) in seen:
                continue
            seen.add(m.start(2))
            if in_comment(m.start()):
                skipped.append((m.group(0), "주석 안"))
                continue
            px = float(m.group(2))
            rem = to_rem(px)
            if rem is None:
                skipped.append((m.group(0), "rem 으로 깨끗이 안 떨어진다"))
                continue
            out.append(css[prev:m.start()])
            out.append("%s%srem" % (m.group(1), rem))
            prev = m.end()
            changed.append((px, rem))
        out.append(css[prev:])
        pieces.append(text[last:a])
        pieces.append("".join(out))
        last = b
    pieces.append(text[last:])
    return "".join(pieces), changed, skipped


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    apply = "--apply" in argv
    bad = [a for a in argv if a.startswith("-") and a != "--apply"]
    if bad:
        # ★ 모르는 깃발을 조용히 무시하지 않는다([295] 가 배운 자리) — 무시하면
        #   `--aply` 오타에 "0곳 바꿨습니다" 라 답하고 사람은 됐다고 여긴다.
        print("모르는 깃발: %s (쓸 수 있는 것: --apply)" % ", ".join(bad))
        return 2

    total = 0
    for rel in TARGETS:
        p = os.path.join(ROOT, rel)
        raw = open(p, "rb").read()
        eol = chr(13) + chr(10) if raw.count(b"\r\n") else chr(10)
        text = io.open(p, encoding="utf-8", newline="").read()
        new, changed, skipped = convert(text)

        inline = len(re.findall(r'style="[^"]*font-size', text))
        print("== %s" % rel)
        print("   바꿀 곳 %d · 건너뜀 %d · **안 건드리는 인라인 %d**"
              % (len(changed), len(skipped), inline))
        if changed:
            uniq = sorted({c[0] for c in changed})
            print("   px 값 %d가지: %s%s"
                  % (len(uniq), ", ".join("%g" % v for v in uniq[:12]),
                     " …" if len(uniq) > 12 else ""))
        for g, why in skipped[:8]:
            print("   건너뜀: %-22s — %s" % (g, why))
        if len(skipped) > 8:
            print("   건너뜀 %d곳 더" % (len(skipped) - 8))

        if not apply:
            continue
        if new == text:
            print("   이미 같다 — 안 씀")
            continue
        if eol != chr(10):
            new = new.replace(chr(10), eol)
        open(p, "wb").write(new.encode("utf-8"))
        print("   고침 — %d bytes" % len(new.encode("utf-8")))
        total += len(changed)

    if not apply:
        print()
        print("아무것도 안 고쳤습니다 — 실제로 바꾸려면 --apply")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
