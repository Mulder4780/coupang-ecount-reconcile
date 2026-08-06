# -*- coding: utf-8 -*-
"""contrast_audit.py — 앱의 **모든 글자색**을 배경 대비 명암비로 전수 검사한다.

왜 (2026-08-06 지시: "텍스트가 잘 안보이는 것들이 있어, 전체적으로 전수 조사해서
검토하고 잘 보이게 정리해")
  '연한 회색이 안 보인다'는 눈짐작으로 고치면 고친 곳만 좋아지고 나머지는 그대로다.
  실제로 `--ink-3: rgba(60,60,67,.42)` 는 흰 배경에서 **2.6:1** 이었다 — 본문 기준
  4.5:1 의 절반이다. 그런 값이 스타일시트 여기저기에 흩어져 있었다.
  그래서 색을 하나씩 보는 대신 **수치로 전부 세어** 미달만 골라낸다.

기준 (WCAG 2.1 AA — 사무용 화면의 사실상 표준)
  · 본문(작은 글자)      4.5:1
  · 큰 글자(18.66px+ 또는 14px+ 굵게)  3.0:1
  · 비활성(:disabled)·장식용 테두리는 대상이 아니다.

한계를 숨기지 않는다
  CSS 상속을 전부 풀지는 않는다. **같은 규칙 안에 배경이 있으면 그 배경**, 없으면
  페이지에서 가장 어두운 흔한 배경(--bg)을 쓴다 — 어두운 배경일수록 진한 글자의
  대비가 낮으니 **불리한 쪽으로** 잡는 셈이고, 그래서 '통과'는 믿을 수 있다.
  배경을 모르는 밝은 글자(흰 글자 등)는 실패로 세지 않고 '확인 필요'로 따로 낸다.

  python webapp/contrast_audit.py            미달 목록
  python webapp/contrast_audit.py --all      전부 (통과 포함)
"""
import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_HTML = os.path.join(HERE, "index.html")
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

AA_NORMAL, AA_LARGE = 4.5, 3.0
NAMED = {"white": (255, 255, 255), "black": (0, 0, 0), "red": (255, 0, 0),
         "gray": (128, 128, 128), "grey": (128, 128, 128)}
SKIP_VALUES = {"inherit", "currentcolor", "transparent", "unset", "initial", "none", "auto"}
# 대비를 따질 대상이 아닌 것들 — 비활성 표시·자리표시자는 일부러 흐리다.
SKIP_SELECTOR = re.compile(r":disabled|\[disabled\]|::placeholder|::-webkit|"
                           r"\.muted-ok|::selection|:focus-visible")


def _srgb(c):
    c /= 255.0
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def luminance(rgb):
    r, g, b = (_srgb(x) for x in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(fg, bg):
    a, b = luminance(fg), luminance(bg)
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


def over(fg, alpha, bg):
    """반투명 글자색을 배경 위에 합성한다 — rgba 는 이 계산을 해야 실제 색이 나온다."""
    return tuple(round(f * alpha + b * (1 - alpha)) for f, b in zip(fg, bg))


def parse_color(value, vars_, bg):
    """CSS 색 문자열 → (r,g,b). 못 읽으면 None. var()·rgba()·#RGB(A) 를 다룬다."""
    v = str(value or "").strip().lower()
    for _ in range(4):                      # var(--a, var(--b, #fff)) 중첩 풀기
        m = re.match(r"var\(\s*(--[\w-]+)\s*(?:,\s*(.+))?\)$", v)
        if not m:
            break
        v = (vars_.get(m.group(1)) or m.group(2) or "").strip().lower()
    if not v or v in SKIP_VALUES:
        return None
    if v in NAMED:
        return NAMED[v]
    m = re.match(r"#([0-9a-f]{3,8})$", v)
    if m:
        h = m.group(1)
        if len(h) in (3, 4):
            h = "".join(ch * 2 for ch in h)
        rgb = tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
        if len(h) == 8:                      # #rrggbbaa
            return over(rgb, int(h[6:8], 16) / 255.0, bg)
        return rgb
    m = re.match(r"rgba?\(([^)]+)\)$", v)
    if m:
        parts = [p.strip() for p in re.split(r"[,\s/]+", m.group(1)) if p.strip()]
        try:
            rgb = tuple(min(255, max(0, int(round(float(p.rstrip("%")))))) for p in parts[:3])
        except ValueError:
            return None
        if len(parts) >= 4:
            a = float(parts[3].rstrip("%"))
            if parts[3].endswith("%"):
                a /= 100.0
            return over(rgb, max(0.0, min(1.0, a)), bg)
        return rgb
    return None


def style_block(html):
    m = re.search(r"<style>(.*?)</style>", html, re.S)
    return re.sub(r"/\*.*?\*/", "", m.group(1), flags=re.S) if m else ""


def root_vars(css):
    """{테마이름: {변수: 값}} — 기본은 항상 있고, 다크모드가 있으면 따로 나온다.

    ★ `:root` 블록을 전부 한 사전에 합치면 안 된다 (2026-08-06 실측).
      docs/app.html 은 밝은 팔레트 뒤에 다크 팔레트가 또 나온다. 합치면 나중 값이
      이겨서 **밝은 화면을 어두운 값으로 재게 되고**, 멀쩡한 색이 미달로 나온다.
      테마마다 한 벌씩 재야 두 화면 모두를 실제로 지킬 수 있다.
    """
    base, over = {}, {}
    for m in re.finditer(r"(@media[^{]*\{)?\s*:root([^{\s]*)\s*\{([^}]*)\}", css, re.S):
        media, attr, body = (m.group(1) or ""), (m.group(2) or ""), m.group(3)
        tag = (media + attr).lower()
        name = "dark" if "dark" in tag else ("light" if "light" in tag else "")
        d = base if not name else over.setdefault(name, {})
        for k, val in re.findall(r"(--[\w-]+)\s*:\s*([^;]+)", body):
            d[k] = val.strip()
    out = {"기본": dict(base)}
    for name, d in over.items():
        merged = dict(base)
        merged.update(d)
        out["다크" if name == "dark" else "밝은"] = merged
    # 밝은 테마를 따로 선언한 파일에서는 '기본' 이 그것과 같다 — 두 번 세지 않는다.
    if out.get("밝은") == out["기본"]:
        out.pop("밝은")
    return out


def _decls(body):
    out = {}
    for k, v in re.findall(r"([\w-]+)\s*:\s*([^;]+)", body):
        out[k.strip().lower()] = v.strip()
    return out


def rules(css):
    """(선택자, 선언, @블록 문맥) 목록.

    ★ 정규식으로 `@media…{` 만 지워서는 안 된다 (2026-08-06 실측). 짝이 없어진
      닫는 중괄호 때문에 규칙 경계가 밀려, `.appbar h1 small`(실제 `#9FB4E8`)이
      엉뚱하게 `var(--ink-3)` 로 보고됐다. **없는 문제를 고치게 만든다.**
      그래서 중괄호 깊이를 세어 블록을 정확히 끊고, 어느 @블록 안인지도 남긴다
      (같은 선택자가 화면 폭에 따라 다른 색을 쓰는 곳이 있다).
    """
    out, stack, buf, i, n = [], [], [], 0, len(css)
    while i < n:
        ch = css[i]
        if ch == "{":
            head = " ".join("".join(buf).split())
            buf = []
            if head.startswith("@"):
                stack.append(head)
            else:
                depth, j = 1, i + 1
                while j < n and depth:
                    if css[j] == "{":
                        depth += 1
                    elif css[j] == "}":
                        depth -= 1
                    j += 1
                decl = _decls(css[i + 1:j - 1])
                if head and head != ":root" and decl:
                    out.append((head, decl, " ".join(stack)))
                i = j
                continue
        elif ch == "}":
            if stack:
                stack.pop()
            buf = []
        else:
            buf.append(ch)
        i += 1
    return out


def is_large(decl):
    """큰 글자 기준(18.66px 이상, 또는 14px 이상이면서 굵게)."""
    m = re.search(r"([\d.]+)px", decl.get("font-size", ""))
    size = float(m.group(1)) if m else 0.0
    w = decl.get("font-weight", "")
    bold = w in ("bold", "bolder") or (w.isdigit() and int(w) >= 600)
    return size >= 18.66 or (size >= 14 and bold)


def bg_value(decl):
    """규칙이 정한 배경색 문자열. 그라디언트면 **첫 색**을 쓴다(대개 가장 밝은 쪽)."""
    raw = decl.get("background-color") or decl.get("background") or ""
    if not raw:
        return ""
    m = re.search(r"(#[0-9a-fA-F]{3,8}|rgba?\([^)]*\)|var\(\s*--[\w-]+[^)]*\))", raw)
    return m.group(1) if m else raw.split()[0]


def opaque(value, vars_):
    """불투명한 배경인가.

    반투명 배경은 **그 밑에 무엇이 있느냐**로 실제 색이 달라진다. 예를 들어 남색
    탭바 위의 `rgba(255,255,255,.14)` 를 흰 페이지 위에 합성하면 거의 흰색이 되어,
    그 위의 흰 글자가 '명암비 1.0' 이라는 엉뚱한 결과가 나온다(실측). 그래서
    반투명이면 여기서 확정하지 않고 **더 위 조상**을 계속 찾는다.
    """
    v = str(value or "").strip().lower()
    for _ in range(4):
        m = re.match(r"var\(\s*(--[\w-]+)\s*(?:,\s*(.+))?\)$", v)
        if not m:
            break
        v = (vars_.get(m.group(1)) or m.group(2) or "").strip().lower()
    m = re.match(r"rgba\(([^)]+)\)$", v)
    if m:
        parts = [p.strip() for p in re.split(r"[,\s/]+", m.group(1)) if p.strip()]
        if len(parts) >= 4:
            try:
                a = float(parts[3].rstrip("%"))
                return (a / 100.0 if parts[3].endswith("%") else a) >= 0.9
            except ValueError:
                return False
    m = re.match(r"#([0-9a-f]{8})$", v)
    if m:
        return int(m.group(1)[6:8], 16) / 255.0 >= 0.9
    return True


def ancestor_bg(sel, bg_map, page_bg, ctx=""):
    """부모가 정한 배경을 따라간다.

    `.hero .hero-period` 처럼 **자기 규칙엔 배경이 없고 부모에 있는** 경우가 많다.
    이걸 안 보면 남색 바탕의 흰 글자가 전부 '미달'로 나온다(실측 거짓양성 다수).
    선택자를 뒤에서부터 한 마디씩 떼며 가장 가까운 조상의 배경을 찾는다.
    """
    parts = sel.split(",")[0].strip().split()
    # 같은 @블록 안의 배경을 먼저 본다 — 넓은 화면에서만 남색이 되는 사이드바처럼,
    # **같은 선택자가 화면 폭에 따라 다른 바탕**을 갖는 곳이 있다.
    for scope in ([ctx, ""] if ctx else [""]):
        for i in range(len(parts) - 1, -1, -1):
            for key in (" ".join(parts[:i + 1]), parts[i]):
                got = bg_map.get((scope, key))
                if got:
                    return got, True
    return page_bg, False


def inline_styles(html):
    """마크업의 `style="…"` 도 검사 대상이다.

    ★ 스타일시트만 보면 놓친다 (2026-08-06 실측): `.ui-card-ic` 아이콘이
      `style="background:#E3F0FF;color:#0A84FF"` 로 적혀 있어 3.0:1 이었는데
      <style> 블록에는 흔적이 없었다. **전수 조사라면 여기도 세야 한다.**
      배경과 글자색이 **함께 적힌 것만** 본다(배경이 없으면 무엇 위인지 알 수 없다).
    """
    body = re.sub(r"<style>.*?</style>", "", html, flags=re.S)
    out = []
    for i, raw in enumerate(re.findall(r'style="([^"]*)"', body)):
        decl = _decls(raw)
        if "color" in decl and bg_value(decl):
            out.append(("style=\"%s\"" % raw[:60], decl, ""))
    return out


def audit(path=DEFAULT_HTML):
    """(미달, 확인필요, 통과수) — 테마가 여러 개면 전부 재고 합쳐서 돌려준다."""
    html = open(path, encoding="utf-8").read()
    css = style_block(html)
    if not css:
        return [], [], 0
    bad, unknown, ok = [], [], 0
    for theme, v in root_vars(css).items():
        b, u, o = audit_theme(html, css, v, theme)
        bad += b
        unknown += u
        ok += o
    bad.sort(key=lambda r: r["ratio"])
    return bad, unknown, ok


def audit_theme(html, css, v, theme="기본"):
    """한 테마(변수 한 벌)로 전수 검사한다."""
    page_bg = parse_color(v.get("--bg", "#ffffff"), v, (255, 255, 255)) or (255, 255, 255)
    # 카드 배경의 이름은 파일마다 다르다(--surface / --card). 못 찾았다고 **흰색으로
    # 가정하면** 다크 테마를 흰 바탕에서 재게 되어 멀쩡한 색이 무더기로 미달로 나온다
    # (2026-08-06 실측: docs/app.html 다크 팔레트 9건이 전부 거짓 미달이었다).
    surface = None
    for key in ("--surface", "--card", "--bg"):
        if v.get(key):
            surface = parse_color(v[key], v, page_bg)
            if surface:
                break
    surface = surface or page_bg

    # 어떤 선택자가 어떤 배경을 깔았나 — 자식 글자의 실제 바탕을 알아내는 데 쓴다.
    parsed = rules(css) + inline_styles(html)
    bg_map = {}
    for sel, decl, ctx in parsed:
        raw = bg_value(decl)
        if not raw or not opaque(raw, v):
            continue
        got = parse_color(raw, v, page_bg)
        if got:
            for one in (s.strip() for s in sel.split(",")):
                if one:
                    bg_map.setdefault((ctx, one), got)

    bad, unknown, ok = [], [], 0
    for sel, decl, ctx in parsed:
        if "color" not in decl or SKIP_SELECTOR.search(sel):
            continue
        own = bg_value(decl)
        bg, known = page_bg, False
        if own and opaque(own, v):
            got = parse_color(own, v, page_bg)
            if got:
                bg, known = got, True
            else:
                # 배경을 **선언은 했는데 값을 못 읽는** 경우(사람마다 달라지는 아바타
                # 색처럼 실행 중에 정해지는 변수). 흰 바탕으로 가정하면 없는 문제를
                # 만든다 — 판정하지 않고 '확인 필요'로 넘긴다.
                unknown.append({"sel": sel, "color": decl["color"], "theme": theme})
                continue
        if not known:
            bg, known = ancestor_bg(sel, bg_map, page_bg, ctx)
        fg = parse_color(decl["color"], v, bg)
        if fg is None:
            continue
        # 바탕을 끝내 모르는 **밝은 글자**는 어두운 바탕 위일 가능성이 크다 — 실패로 세지 않는다.
        if not known and luminance(fg) > 0.35:
            unknown.append({"sel": sel, "color": decl["color"], "theme": theme})
            continue
        need = AA_LARGE if is_large(decl) else AA_NORMAL
        # 배경을 모르면 흔한 두 바탕 중 **불리한 쪽**으로 잡는다.
        ratio = min(contrast(fg, bg), contrast(fg, surface)) if not known \
            else contrast(fg, bg)
        if ratio + 1e-9 < need:
            bad.append({"sel": sel, "color": decl["color"], "rgb": fg,
                        "bg": bg, "ratio": round(ratio, 2), "need": need,
                        "배경확인": known, "ctx": ctx, "theme": theme})
        else:
            ok += 1
    return bad, unknown, ok


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=DEFAULT_HTML)
    ap.add_argument("--all", action="store_true")
    a = ap.parse_args()
    bad, unknown, ok = audit(a.file)
    print("글자색 명암비 전수검사 — %s" % os.path.basename(a.file))
    print("  통과 %d · 미달 %d · 배경 확인 필요 %d (기준 본문 %.1f:1 · 큰 글자 %.1f:1)"
          % (ok, len(bad), len(unknown), AA_NORMAL, AA_LARGE))
    for r in bad:
        print("  [%.2f:1 / %.1f 필요] %s  color:%s → rgb%s"
              % (r["ratio"], r["need"], r["sel"][:70], r["color"], r["rgb"]))
    if a.all:
        for r in unknown:
            print("  [배경모름] %s  color:%s" % (r["sel"][:70], r["color"]))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
