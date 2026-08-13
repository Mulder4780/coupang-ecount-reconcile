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
  python webapp/font_switch.py                  # 지금 어느 글꼴인가 (읽기만)
  python webapp/font_switch.py --legacy         # 예전 글꼴(나눔고딕)로 전부 되돌린다
  python webapp/font_switch.py --modern         # 다시 기본 글꼴로
  python webapp/font_switch.py --list           # 고를 수 있는 글꼴을 다 보여 준다
  python webapp/font_switch.py --preset galaxy  # 모두의 기본을 갤럭시 글자체로
  python webapp/font_switch.py --sync           # 프리셋 CSS 블록만 표와 맞춘다

글꼴 프리셋 (2026-08-13 지시)
  `FONT_PRESETS` 표 하나가 정본이다. 이 도구는 두 가지를 만든다:
    ① `--font-ui` 한 줄 — **모두의 기본** (파일에 박힌다)
    ② `:root[data-font="<이름>"]` 블록 — **사람이 제 기기에서 고르는 것**
       (앱 [실행] 탭 '글꼴' 카드 · 브라우저가 기억 · 코드는 안 건드린다)
  ②가 없으면 앱에서 골라도 그 화면만 안 바뀐다 — 그래서 `apply()` 가 늘 같이 맞춘다.
  표에 프리셋을 하나 더하면 도구·화면·검증이 **저절로** 따라온다(이름을 손으로
  적는 자리가 없다). 화면 목록은 서버 `/api/font-presets` 가 이 표를 그대로 내려 준다.

검증 [126] 이 이 구조와 왕복(modern→legacy→modern)을 지키고,
검증 [246] 이 프리셋 표·블록·화면이 갈리지 않는지를 지킨다.
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

# ── 글꼴 프리셋 (2026-08-13 지시: "갤럭시 글자체, 아이폰 글자체, 스타일등을 적용해서
#    사용자가 설정에서 변경할 수 있게 다양한 스타일 적용 코딩해")
#
# ★ 표는 **여기 하나**다. 화면(index.html)도 서버도 이름을 손으로 적지 않는다 —
#   적는 순간 사본이 생기고, 프리셋을 하나 늘린 날 한쪽만 고쳐진다([162]).
#
# ★ **내려받는 글꼴을 넣지 않는다.** 전부 그 기기에 이미 깔려 있는 것만 쓴다.
#   웹폰트를 넣으면 폰이 처음 열 때 몇백 KB 를 받고, 못 받으면 **조용히 다른 글꼴로**
#   그려진다 — 화면을 열어 보기 전에는 아무도 모른다. 그래서 '갤럭시 글자체'는
#   삼성 기기에서만 진짜로 바뀌고 다른 기기에서는 그다음 후보가 쓰인다.
#   그 사실을 `실제로` 칸에 적어 두어 화면이 사람에게 그대로 말한다([169]).
#
# ★ 순서가 곧 결과다. 브라우저는 **글자 하나하나마다** 스택을 걸어 내려가며 그 글자가
#   있는 첫 글꼴을 쓴다. 그래서 한글이 없는 라틴 전용 글꼴(Segoe UI·SF Pro)을 앞에
#   두어도 한글은 뒤로 넘어간다 — 한글용을 **맑은 고딕보다 앞**에 두는 것이 요점이다
#   (2026-08-07 실사고: Noto Sans KR 이 맑은 고딕 뒤에 있어 영영 차례가 안 왔다).
FONT_PRESETS = {
    "basic": {
        "이름": "기본",
        "설명": "기기가 제 화면에 쓰는 글꼴 — 작은 글씨가 또렷합니다",
        "실제로": "어느 기기에서나 그 기기의 기본 글꼴을 씁니다",
        "값": MODERN,
    },
    "galaxy": {
        "이름": "갤럭시 글자체",
        "설명": "삼성 갤럭시가 제 화면에 쓰는 글꼴(One UI Sans·삼성고딕)",
        "실제로": "삼성 기기에서만 그대로 보입니다 — 그 밖에서는 본고딕/Roboto",
        "값": ('"One UI Sans KR VF","One UI Sans KR","SamsungOne","SamsungOneKorean",'
               '"Samsung Sharp Sans","삼성고딕","SamsungGothic",Roboto,'
               '"Noto Sans KR","Source Han Sans KR","본고딕",'
               '"Apple SD Gothic Neo","Malgun Gothic","맑은 고딕",sans-serif'),
    },
    "iphone": {
        "이름": "아이폰 글자체",
        "설명": "애플이 제 화면에 쓰는 글꼴(SF Pro·애플 SD 산돌고딕)",
        "실제로": "아이폰·아이패드·맥에서만 그대로 보입니다 — 그 밖에서는 본고딕",
        "값": ('-apple-system,BlinkMacSystemFont,"SF Pro Text","SF Pro Display",'
               '"Apple SD Gothic Neo","AppleSDGothicNeo-Regular","애플 SD 산돌고딕 Neo",'
               '"Noto Sans KR","Source Han Sans KR","본고딕",'
               '"Malgun Gothic","맑은 고딕",sans-serif'),
    },
    "windows": {
        "이름": "윈도우 글자체",
        "설명": "사무실 PC 화면과 같은 글꼴(Segoe UI·맑은 고딕)",
        "실제로": "윈도우에서만 그대로 보입니다 — 그 밖에서는 본고딕",
        "값": ('"Segoe UI Variable Text","Segoe UI","맑은 고딕","Malgun Gothic",'
               '"Noto Sans KR","Source Han Sans KR","본고딕",'
               '"Apple SD Gothic Neo",Roboto,sans-serif'),
    },
    "clear": {
        "이름": "또렷하게",
        "설명": "획을 굵게 그려 눈이 편합니다 — 글씨가 흐릴 때 고르세요",
        "실제로": "글꼴 굵기는 그대로이고 **그리는 방식**이 바뀝니다(크기는 안 커집니다)",
        # ★ Pretendard·본고딕을 맨 앞에 둔다(획이 고르다) + 안티에일리어싱을 끈다.
        #   `-webkit-font-smoothing:antialiased` 는 글자를 **얇게** 그린다 —
        #   그것이 "글씨가 잘 안 보인다"의 절반이다. `auto` 면 서브픽셀로 굵어진다.
        "값": ('"Pretendard Variable",Pretendard,"Noto Sans KR","Source Han Sans KR",'
               '"본고딕","Apple SD Gothic Neo",-apple-system,BlinkMacSystemFont,'
               'Roboto,"Segoe UI Variable Text","Segoe UI","Malgun Gothic","맑은 고딕",'
               'sans-serif'),
        "매끄럽게": "auto",
    },
    "legacy": {
        "이름": "예전(나눔고딕)",
        "설명": "2026-07-31까지 쓰던 글꼴로 되돌립니다",
        "실제로": "파일마다 적어 둔 --font-ui-legacy 값을 그대로 씁니다",
        "값": "var(--font-ui-legacy)",
    },
}

# 파일에 심는 프리셋 CSS 를 감싸는 표식. **이 사이는 도구가 만든다** —
# 사람이 고치면 다음 실행에 덮인다(그래서 블록 안에도 그렇게 적어 둔다).
PRESET_BEGIN = "/* ▼ 글꼴 프리셋 — font_switch.py 가 만든다. 손으로 고치지 말 것 (FONT-PRESETS:BEGIN) */"
PRESET_END = "/* ▲ (FONT-PRESETS:END) */"
# `legacy` 는 블록에 안 넣는다 — 그 한 줄은 2026-08-06 부터 네 파일에 이미 있고
# 되돌리기 경로 ③ 으로 문서에 적혀 있다. 옮기면 그 문서가 조용히 틀려진다.
BLOCK_KEYS = tuple(k for k in FONT_PRESETS if k != "legacy")

ACTIVE = re.compile(r"(--font-ui\s*:\s*)(.+?)(;)")
LEGACY = re.compile(r"--font-ui-legacy\s*:\s*(.+?);")
BLOCK = re.compile(re.escape(PRESET_BEGIN) + r".*?" + re.escape(PRESET_END), re.S)
# 블록을 심을 자리 — 되돌리기 스위치 바로 아래(같은 성격끼리 모아 둔다)
ANCHOR = re.compile(r'^:root\[data-font="legacy"\]\{[^\n]*\}$', re.M)

# ── 화면 단추도 같은 표에서 만든다 (앱 [실행] 탭 '글꼴' 카드)
#    ★ 이름·설명을 index.html 에 손으로 적으면 그 순간 사본이 둘이 된다 —
#      표에 프리셋을 하나 더한 날 CSS 는 늘고 단추는 안 늘면서 **오류가 안 난다**([162]).
#      그래서 단추 목록도 도구가 만든다. 서버 API 로 안 하는 이유는 **폰이 꺼진 PC 를
#      못 부르기 때문**이다 — 못 부르면 카드가 비고, 그러면 되돌릴 길까지 사라진다.
CARDS_BEGIN = "<!-- ▼ 글꼴 단추 — font_switch.py 가 만든다. 손으로 고치지 말 것 (FONT-CARDS:BEGIN) -->"
CARDS_END = "<!-- ▲ (FONT-CARDS:END) -->"
CARDS = re.compile(re.escape(CARDS_BEGIN) + r".*?" + re.escape(CARDS_END), re.S)
CARDS_FILE = os.path.join("webapp", "index.html")   # 고르는 화면은 여기 하나뿐이다


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
        # 어느 프리셋인지 **표에 대 본다** — 이름을 여기 손으로 적으면 사본이 된다.
        name = ""
        for k, p in FONT_PRESETS.items():
            if act == p["값"]:
                name = k
                break
        if not name and act == legv:
            name = "legacy"
        out.append({"파일": rel,
                    # 옛 낱말을 그대로 쓴다(검증·문서가 이 글자를 본다)
                    "상태": "예전" if name == "legacy" else ("기본" if name == "basic"
                            else (FONT_PRESETS[name]["이름"] if name else "목록 밖")),
                    "프리셋": name,
                    "지금": act, "예전": legv,
                    # ★ 블록이 낡으면 그 화면에서만 프리셋이 안 바뀐다 — 조용한 종류다.
                    "프리셋블록": "최신" if preset_css(_eol(text)) in text else
                                  ("낡음" if BLOCK.search(text) else "없음")})
    return out


def _eol(text):
    """줄바꿈은 원본의 일부다 — 심는 줄도 그 파일 방식을 따라간다."""
    return "\r\n" if "\r\n" in text else "\n"


def preset_css(eol="\n"):
    """프리셋 CSS 블록. **표 하나에서 만들어진다** — 파일마다 손으로 적지 않는다."""
    L = [PRESET_BEGIN,
         "/*   앱 [실행] 탭 '글꼴' 카드에서 고른다. 이 기기에서만 바뀐다(브라우저가 기억).",
         "     모두에게 바꾸는 것은 `python webapp/font_switch.py --preset <이름>` 이다. */"]
    for k in BLOCK_KEYS:
        p = FONT_PRESETS[k]
        L.append(':root[data-font="%s"]{ --font-ui:%s; }  /* %s */' % (k, p["값"], p["이름"]))
        if p.get("매끄럽게"):
            # ★ body 가 `-webkit-font-smoothing` 을 직접 적어 두었으므로 :root 만 바꾸면
            #   안 먹는다. body 까지 집어 주되 !important 는 안 쓴다 — 선택자가 더 세다.
            L.append(':root[data-font="%s"],:root[data-font="%s"] body{'
                     ' -webkit-font-smoothing:%s; -moz-osx-font-smoothing:auto; }'
                     % (k, k, p["매끄럽게"]))
    L.append(PRESET_END)
    return eol.join(L)


def _esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def preset_cards(eol="\n"):
    """앱 [실행] 탭 '글꼴' 카드의 단추들. **표 하나에서 만들어진다**."""
    L = [CARDS_BEGIN]
    for k, p in FONT_PRESETS.items():
        ic = "i-bootstrap-arrow-repeat" if k == "legacy" else "i-bootstrap-check-lg"
        L.append('        <button class="ui-card" type="button" data-fontkey="%s"'
                 ' onclick="setFontPreset(\'%s\')">' % (k, k))
        L.append('          <span class="ui-card-ic"><svg aria-hidden="true">'
                 '<use href="#%s"/></svg></span>' % ic)
        # ★ '실제로' 를 같이 적는다 — 갤럭시 글자체는 삼성 기기에서만 진짜로 바뀐다.
        #   그 말을 안 적으면 아이폰에서 고른 사람이 "안 바뀐다"고 여긴다([169]).
        L.append('          <span class="ui-card-tx"><b>%s</b><span>%s</span>'
                 '<span class="norm">%s</span></span></button>'
                 % (_esc(p["이름"]), _esc(p["설명"]), _esc(p["실제로"])))
    L.append("      " + CARDS_END)
    return eol.join(L)


def sync_cards():
    """화면 단추를 표와 맞춘다. 바뀌었으면 파일 이름을 담은 목록을 돌려준다."""
    text = _read(CARDS_FILE)
    eol = _eol(text)
    want = preset_cards(eol)
    if not CARDS.search(text):
        # ★ 자리를 못 찾으면 조용히 넘어가지 않는다 — 그러면 표만 늘고 화면은
        #   그대로인데 아무 오류도 안 난다.
        raise RuntimeError("%s 에 글꼴 단추 자리(FONT-CARDS)가 없다" % CARDS_FILE)
    new = CARDS.sub(lambda m: want, text, count=1)
    if new == text:
        return []
    _write(CARDS_FILE, new)
    return [CARDS_FILE]


def sync_presets():
    """네 파일의 프리셋 블록을 표와 맞춘다. 바뀐 파일 목록을 돌려준다.

    ★ 없으면 심고 있으면 갈아끼운다. **자리는 되돌리기 스위치 바로 아래** 하나뿐이라
      두 번 심기지 않는다. 자리를 못 찾으면 조용히 넘어가지 않고 예외를 올린다 —
      그 파일만 프리셋이 없는 채로 남으면 그 화면에서만 안 바뀌고, 화면을 열어
      보기 전에는 아무도 모른다([169]).
    """
    changed = []
    for rel in FILES:
        text = _read(rel)
        eol = _eol(text)
        want = preset_css(eol)
        if BLOCK.search(text):
            new = BLOCK.sub(lambda m: want, text, count=1)
        else:
            m = ANCHOR.search(text)
            if not m:
                raise RuntimeError(
                    '%s 에 :root[data-font="legacy"] 줄이 없다 — 프리셋을 심을 자리가 없다' % rel)
            new = text[:m.end()] + eol + want + text[m.end():]
        if new != text:
            _write(rel, new)
            changed.append(rel)
    return changed


def apply(mode):
    """모두의 기본 글꼴을 바꾼다. mode 는 `FONT_PRESETS` 의 열쇠.

    옛 이름도 그대로 받는다: 'modern' → 'basic'(2026-08-06 부터 쓰던 이름이라
    스크립트·문서·검증이 그 낱말로 부른다. 지우면 그것들이 조용히 죽는다).
    """
    mode = {"modern": "basic"}.get(mode, mode)
    if mode not in FONT_PRESETS:
        raise ValueError(mode)
    want = FONT_PRESETS[mode]["값"]
    # 프리셋 블록·화면 단추는 언제나 표와 같게 둔다 — 하나만 맞추면 갈린다
    changed = list(sync_presets())
    for r in sync_cards():
        if r not in changed:
            changed.append(r)
    for rel in FILES:
        text = _read(rel)
        leg = LEGACY.search(text)
        if not leg:
            raise RuntimeError("%s 에 --font-ui-legacy 가 없다 — 되돌릴 자리를 잃었다" % rel)
        # legacy 줄은 건드리지 않는다: 그 줄을 잠시 치워 두고 활성 줄만 바꾼다
        holder = "\x00LEGACY\x00"
        stashed = text.replace(leg.group(0), holder, 1)
        if _active(text) == want:
            continue
        new = ACTIVE.sub(lambda m: m.group(1) + want + m.group(3), stashed, count=1)
        new = new.replace(holder, leg.group(0), 1)
        if new != text:
            _write(rel, new)
            if rel not in changed:
                changed.append(rel)
    return changed


def _print():
    rows = state()
    for r in rows:
        mark = "예전(나눔고딕)" if r["상태"] == "예전" else r["상태"]
        print("  %-24s %s" % (r["파일"], mark))
        if r["지금"]:
            print("       지금: %s" % r["지금"][:88])
        if r.get("프리셋블록") != "최신":
            print("       ★ 프리셋 블록 %s — `--sync` 로 맞춘다" % r["프리셋블록"])
    kinds = {r["상태"] for r in rows}
    if len(kinds) > 1:
        print("\n★ 파일마다 글꼴이 다르다 — 한 화면만 다른 글꼴로 보인다.")
        print("  `python webapp/font_switch.py --modern` 또는 `--legacy` 로 맞춘다.")
    else:
        print("\n네 파일 모두 같은 글꼴을 쓴다.")
    if any(r.get("프리셋블록") != "최신" for r in rows):
        print("★ 프리셋 블록이 표와 다르다 — 앱에서 고른 글꼴이 그 화면에만 안 먹는다.")


def _list():
    print("고를 수 있는 글꼴 (표는 font_switch.FONT_PRESETS 하나다)\n")
    for k, p in FONT_PRESETS.items():
        print("  --preset %-8s %s" % (k, p["이름"]))
        print("      %s" % p["설명"])
        print("      실제로: %s" % p["실제로"])


def main():
    argv = sys.argv[1:]
    arg = argv[0] if argv else ""
    if arg in ("--list", "--presets"):
        _list()
        return 0
    if arg == "--sync":
        changed = sync_presets()
        for r in sync_cards():
            if r not in changed:
                changed.append(r)
        print("프리셋 블록·화면 단추를 표와 맞췄습니다 — 고친 파일 %d개%s"
              % (len(changed), (": " + ", ".join(changed)) if changed else " (이미 같았다)"))
        return 0
    mode = ""
    if arg in ("--legacy", "--modern"):
        mode = arg[2:]
    elif arg == "--preset":
        mode = argv[1] if len(argv) > 1 else ""
        if mode not in FONT_PRESETS:
            # ★ 모르는 이름을 조용히 무시하지 않는다 — 그러면 '바꿨습니다'가 찍히고
            #   아무것도 안 바뀐다(실패가 성공처럼 보이는 자리).
            print("모르는 글꼴 이름입니다: %r" % mode)
            print("고를 수 있는 것: %s" % ", ".join(FONT_PRESETS))
            return 2
    if mode:
        changed = apply(mode)
        key = {"modern": "basic"}.get(mode, mode)
        label = FONT_PRESETS[key]["이름"]
        print("모두의 기본 글꼴을 '%s' 로 바꿨습니다 — 고친 파일 %d개%s"
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
