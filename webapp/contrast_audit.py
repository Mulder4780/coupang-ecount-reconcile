# -*- coding: utf-8 -*-
"""웹앱 글자 대비 전수 조사 — 읽기 전용.

WCAG 2.1: 본문 4.5:1, 큰 글씨(24px 이상 또는 18.66px 이상 굵게) 3:1.
셀렉터마다 '이 글자가 어떤 배경 위에 놓이나'를 조상 셀렉터에서 되짚어 정한다.
정확한 캐스케이드 계산은 아니지만, 실제 문제(밝은 회색 글씨)를 잡기엔 충분하다.
"""
import re
import sys
import os

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
PATH = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount\.claude\worktrees\unruffled-jang-78b561\webapp\index.html"
LIMIT = float(sys.argv[2]) if len(sys.argv) > 2 else 4.5

src = open(PATH, encoding="utf-8").read()
# ★ 실제 화면 CSS 는 <head> 안에만 있다. 본문 뒤쪽 JS 문자열에도 <style> 이 들어 있어
#   (이미지 뷰어 팝업 body{background:#111}) 그대로 읽으면 페이지 배경을 검정으로 잘못 잡는다.
head_end = src.find("</head>")
head = src[:head_end] if head_end > 0 else src
styles = re.findall(r"<style[^>]*>(.*?)</style>", head, re.S)
css = "\n".join(styles)
# 주석 제거
css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
# @media 등 중첩 블록의 바깥 껍데기를 벗겨 안쪽 규칙이 보이게 한다
css = re.sub(r"@media[^{]*\{", "", css)
css = re.sub(r"@supports[^{]*\{", "", css)


# CSS 변수(--brand 등)를 먼저 모은다. 이걸 안 풀면 `background:var(--brand)` 를 못 읽어
# 흰 글씨가 전부 '흰 배경 위'로 잘못 잡힌다(실측: 거짓 경보 25건).
VARS = {}
for m in re.finditer(r"(--[\w-]+)\s*:\s*([^;{}]+)", css):
    VARS.setdefault(m.group(1).strip(), m.group(2).strip())


def resolve(val, depth=0):
    if not val or depth > 6 or "var(" not in val:
        return val
    def sub(m):
        name, _, fb = m.group(1).partition(",")
        return VARS.get(name.strip(), fb.strip() or "")
    return resolve(re.sub(r"var\(\s*(--[^)]+)\)", sub, val), depth + 1)


def to_rgb(tok):
    tok = resolve(tok)
    tok = tok.strip().lower()
    m = re.match(r"#([0-9a-f]{3})\b", tok)
    if m and len(tok.split()[0]) == 4:
        h = m.group(1)
        return tuple(int(c * 2, 16) for c in h)
    m = re.match(r"#([0-9a-f]{6})\b", tok)
    if m:
        h = m.group(1)
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    m = re.match(r"rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:[,\s/]+([\d.]+))?", tok)
    if m:
        rgb = tuple(int(float(m.group(i))) for i in (1, 2, 3))
        return rgb + (float(m.group(4)),) if m.group(4) is not None else rgb
    named = {"white": (255, 255, 255), "black": (0, 0, 0), "#fff": (255, 255, 255)}
    return named.get(tok)


def over(fg, bg):
    """반투명 글자색을 배경 위에 **합성**한다.

    ★ 이걸 안 하면 실제 문제를 통째로 놓친다. 이 앱의 흐린 글씨는
      `--ink-2:rgba(60,60,67,.62)` · `--ink-3:rgba(60,60,67,.42)` 인데, 알파를 무시하면
      (60,60,67) 짙은 회색으로 보여 전부 '합격'으로 나온다. 실제로 흰 배경에 얹으면
      ink-3 은 약 (173,173,176) — 2.2:1 로, 사용자가 "잘 안 보인다"고 한 바로 그 글씨다.
    """
    if len(fg) == 4:
        a = fg[3]
        return tuple(int(round(fg[i] * a + bg[i] * (1 - a))) for i in range(3))
    return fg[:3]


def lum(rgb):
    def f(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def ratio(fg, bg):
    a, b = lum(fg), lum(bg)
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


def bg_of(raw):
    """배경 선언에서 '글자가 실제로 놓이는 색'을 뽑는다. 못 정하면 None(조상으로 넘김).

    · 반투명(rgba, 알파<0.6)은 배경이 아니라 **덧칠**이다 — 밑색이 진짜 배경이다.
      (.hero .hero-period 가 흰 글씨+흰 반투명이라 1.00:1 로 잘못 잡혔다)
    · 그라데이션은 첫 색을 대표로 쓴다.
    """
    if not raw:
        return None
    low = resolve(raw).strip().lower()
    if low.startswith("rgba"):
        m = re.match(r"rgba\(\s*[\d.]+[,\s]+[\d.]+[,\s]+[\d.]+[,\s]+([\d.]+)", low)
        if m and float(m.group(1)) < 0.6:
            return None
    if "gradient" in low:
        m = re.search(r"#[0-9a-f]{3,6}", low)
        return to_rgb(m.group(0)) if m else None
    if low.startswith("transparent") or low.startswith("none") or low.startswith("inherit"):
        return None
    return to_rgb(low.split()[0]) if low.split() else None


rules = []
for sel, body in re.findall(r"([^{}]+)\{([^{}]*)\}", css):
    sel = " ".join(sel.split())
    if not sel or sel.startswith("@"):
        continue
    decls = {}
    for d in body.split(";"):
        if ":" not in d:
            continue
        k, v = d.split(":", 1)
        decls[k.strip().lower()] = v.strip()
    for one in [s.strip() for s in sel.split(",") if s.strip()]:
        rules.append((one, decls))

# 배경색 지도: 셀렉터 → 배경 RGB
bgmap = {}
for sel, d in rules:
    c = bg_of(d.get("background-color") or d.get("background"))
    if c:
        bgmap[sel] = c

PAGE_BG = bgmap.get("body") or (255, 255, 255)


def guess_bg(sel, decls):
    """이 셀렉터의 글자가 놓이는 배경. 자기 자신 → 조상 → 페이지 순으로 찾는다."""
    c = bg_of(decls.get("background-color") or decls.get("background"))
    if c:
        return c, "자기"
    parts = sel.split()
    # 조상 조합을 길게부터 짧게 되짚는다: '.a .b .c' → '.a .b', '.a'
    for i in range(len(parts) - 1, 0, -1):
        anc = " ".join(parts[:i])
        if anc in bgmap:
            return bgmap[anc], anc
    # 같은 클래스에 배경만 따로 준 규칙이 있을 수 있다
    if sel in bgmap:
        return bgmap[sel], sel
    last = parts[-1] if parts else sel
    base = last.split(":")[0]
    if base in bgmap:
        return bgmap[base], base
    return PAGE_BG, "페이지"


def big_text(d):
    fs = d.get("font-size", "")
    m = re.match(r"([\d.]+)px", fs)
    size = float(m.group(1)) if m else 0
    fw = d.get("font-weight", "")
    bold = fw in ("bold", "700", "800", "900")
    return size >= 24 or (size >= 18.66 and bold)


# ── 확인 끝난 예외 ───────────────────────────────────────────────
# 이 감사기는 @media 를 펼쳐서 본다. 그래서 **같은 셀렉터가 화면 크기에 따라 다른 배경**
# 위에 놓이는 경우를 구분하지 못한다. 이 앱이 정확히 그렇다:
#   · 폰(<900px): .appbar = 남색, .tabbar = 흰색
#   · 데스크톱(≥900px): .appbar = 흰색, .tabbar = 남색 사이드바
# 아래 항목은 2026-08-06 에 CSS 를 직접 읽어 **문제 없음을 확인**한 것이다.
# 새로 추가할 때는 반드시 근거를 함께 적을 것 — 그냥 조용히 만들면 감사기가 무의미해진다.
OK_KNOWN = {
    ".gate p": "gate 안의 .box 가 흰색이다(gate 배경은 뒤쪽 그라데이션)",
    ".gate .err": "위와 같음 — 흰 상자 위",
    ".appbar #clock": "폰에서는 남색 헤더 위. 데스크톱(흰 헤더)은 --ink-2 로 따로 덮었다",
    ".appbar h1 small": "폰=남색 위. 데스크톱은 --ink-3 로 덮여 있다",
    ".appbar .sub": "폰=남색 위. 데스크톱은 --ink-3 로 덮여 있다",
    ".tabbar button": "폰에서 탭바는 흰색(--surface) → #6E6E72 는 5.08:1. 데스크톱은 #93A5D6 로 덮음",
    ".tabbar button.on": "위와 같음 — 데스크톱은 #fff 로 덮음",
    "body.ryu-mode .tabbar .brand .bt::after": "브랜드는 폰에서 display:none, 데스크톱 남색 사이드바에서만 보인다",
    ".workcenter-person .wc-icon": "40px 아이콘 타일의 흰 글리프 — 그림 요소 기준 3:1 을 넘는다(3.36)",
}

bad, seen, skipped = [], set(), []
for sel, d in rules:
    raw = d.get("color")
    if not raw:
        continue
    fg = to_rgb(raw)
    if not fg:
        continue
    bg, where = guess_bg(sel, d)
    bg = bg[:3]
    r = ratio(over(fg, bg), bg)
    need = 3.0 if big_text(d) else LIMIT
    key = (sel, raw.strip())
    eff0 = over(fg, bg)
    # 흰 글씨인데 배경이 밝게 잡혔고 CSS 에 배경 선언이 없다면, 배경을 JS·인라인이 정하는
    # 아이콘·배지다(예: .ios-avatar). 흰 글씨를 밝은 배경에 '실수로' 두는 일은 없다 —
    # 그러면 화면에서 대번에 안 보인다. 오탐이므로 따로 센다.
    if lum(eff0) > 0.7 and lum(bg) > 0.5 and where == "페이지":
        skipped.append(sel)
        continue
    if r < need and sel in OK_KNOWN:
        skipped.append(sel)
        continue
    if r < need and key not in seen:
        seen.add(key)
        eff = eff0
        shown_fg = resolve(raw.strip())
        if len(fg) == 4:
            shown_fg = "%s→#%02X%02X%02X" % (shown_fg.split("(")[0] + str(fg[3]), *eff)
        bad.append((r, need, sel, shown_fg[:26], "#%02X%02X%02X" % bg, where))

bad.sort()
print("검사한 색 규칙 %d개 · 기준 본문 %.1f:1 · 오탐 제외 %d건"
      % (sum(1 for s, d in rules if d.get("color")), LIMIT, len(skipped)))
print("기준 미달 %d건\n" % len(bad))

# ── 색깔별 요약: 이것이 고칠 목록이다(셀렉터 200줄이 아니라 색 몇 개가 원인이다) ──
from collections import defaultdict
by_color = defaultdict(list)
for r, need, sel, fg, bg, where in bad:
    by_color[fg].append((r, sel, bg))
print("== 고쳐야 할 색 (쓰임 많은 순) ==")
print("%-26s %-6s %-5s %s" % ("글자색", "최악대비", "곳수", "예시 셀렉터"))
for fg, uses in sorted(by_color.items(), key=lambda kv: -len(kv[1])):
    worst = min(u[0] for u in uses)
    ex = ", ".join(u[1][:30] for u in uses[:3])
    print("%-26s %5.2f  %4d  %s" % (fg, worst, len(uses), ex[:70]))

if "--all" in sys.argv:
    print("\n== 전체 ==")
    for r, need, sel, fg, bg, where in bad:
        print("%5.2f  %4.1f  %-26s %-9s %s   [%s]" % (r, need, fg, bg, sel[:52], where))
