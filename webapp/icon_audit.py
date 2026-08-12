# -*- coding: utf-8 -*-
"""아이콘 전수 조사 — 스프라이트 심볼 · 사용처 · 렌더 크기 규칙. 읽기 전용."""
import collections
import re
import sys

if hasattr(sys.stdout, "reconfigure"):   # pythonw 는 sys.stdout 이 None 이다([43])
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
# 기본 경로는 **이 파일 옆의 index.html** 이다. 예전에는 "webapp/index.html" 이라
# 적어 두어, 정작 webapp/ 안에서 실행하면 FileNotFoundError 가 났다(2026-08-07).
# 어디서 부르든 같은 파일을 보게 한다 — 인자를 주면 그것이 우선이다.
import os
PATH = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "index.html")
h = open(PATH, encoding="utf-8").read()

syms = re.findall(r'<symbol[^>]*id="(i-[\w-]+)"', h)
uses = re.findall(r'href="#(i-[\w-]+)"', h)
cnt = collections.Counter(uses)

print("스프라이트 심볼 %d개 · 사용 %d곳 / 서로 다른 %d종" % (len(syms), len(uses), len(cnt)))
unused = [s for s in syms if s not in cnt]
missing = sorted(set(uses) - set(syms))
print("안 쓰는 심볼 %d개%s" % (len(unused), (": " + ", ".join(unused[:10])) if unused else ""))
print("스프라이트에 없는데 쓰는 것 %d개%s"
      % (len(missing), (": " + ", ".join(missing[:10])) if missing else " (없음)"))

# 아이콘 크기를 정하는 CSS 규칙을 모은다
print("\n== 아이콘 크기 규칙(css) ==")
head = h[:h.find("</head>")]
head = re.sub(r"/\*.*?\*/", "", head, flags=re.S)
rules = []
for sel, body in re.findall(r"([^{}]+)\{([^{}]*)\}", head):
    if "svg" not in sel and "icon" not in sel.lower():
        continue
    m = re.search(r"(?<!max-)(?<!min-)width\s*:\s*([\d.]+)px", body)
    n = re.search(r"(?<!max-)(?<!min-)height\s*:\s*([\d.]+)px", body)
    if m or n:
        w = float(m.group(1)) if m else None
        ht = float(n.group(1)) if n else None
        small = min([v for v in (w, ht) if v] or [99])
        rules.append((small, " ".join(sel.split()), w, ht))
rules.sort()
for small, sel, w, ht in rules[:26]:
    mark = "  ★작음" if small < 18 else ""
    print("  %-52s %sx%s%s" % (sel[:52], w or "-", ht or "-", mark))

print("\n== 무슨 아이콘인가 (사용 많은 순) ==")
NAME = {
    "search": "검색", "bell": "알림", "moon": "다크모드 켜기", "sun": "라이트모드 켜기",
    "printer": "인쇄", "image-down": "이미지로 저장", "clipboard-copy": "복사",
    "file-spreadsheet": "엑셀", "refresh-cw": "새로고침", "link-45deg": "링크",
    "key-fill": "PIN 잠금", "calendar": "달력", "check": "확인", "x": "닫기",
    "chevron-right": "더 보기", "arrow-left": "뒤로", "gear": "설정", "person": "담당자",
    "house": "홈", "clock": "시각", "download": "내려받기", "upload": "올려보내기",
    "trash": "삭제", "pencil": "고치기", "plus": "추가", "list": "목록",
}
for s, n in cnt.most_common(30):
    base = s[2:].replace("bootstrap-", "")
    ko = next((v for k, v in NAME.items() if k in base), "")
    print("  %-30s %-14s %d곳" % (s, ko or "?", n))
