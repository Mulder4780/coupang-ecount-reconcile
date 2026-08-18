# -*- coding: utf-8 -*-
"""
band_extract.py — 밴드 게시글 → 구조화 업무 레코드 추출 (월별 백필 원천)
==========================================================================
밴드 게시글은 아래 규격으로 작성되어 있어 기계 파싱이 가능하다.

    ☑️판매전표 +거래명세서 +견적서 = 메일발송 完 ⭕
    ♣ ［ 2026년 02분기 3개월 유료 A/S 완료 ]
    ● A/S 일자 : 2026.06.01 (월요일)
    ● A/S 담당 : 김필우
    ● 프로젝트NO : UJ2600931
    ● 캠프이름 : 양주1캠프

이를 파싱해 [프로젝트NO·업무유형·유상무상·작업일·담당기사·캠프명·진행상태·문서상태]로 만든다.
관리대장에 없는 과거 월(2026-06, 05 …) 백필의 1차 원천이며,
이미 원장에 있는 건은 '원장등록됨'으로 표시해 중복 입력을 막는다.

실행:
  python band_extract.py --month 2026-06            # 6월 추출 → 리포트
  python band_extract.py --month 2026-06 --sheet    # + 관리대장 24_밴드업무추출 시트 반영(vN+1)
  python band_extract.py --all                      # 전체 기간
"""
import sys, os, re, csv, json, glob
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import people_alias as _ALIAS

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(BASE_DIR, "band", "cache")
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE_DIR, "reports")

RE_PRJ = re.compile(r"프로젝트\s*NO\s*[:：]?\s*(UJ\d{6,})", re.I)
RE_DATE = re.compile(r"A/?S\s*일자\s*[:：]?\s*(\d{4})[.\-/](\d{1,2})[.,\-/](\d{1,2})")
RE_TECH = re.compile(r"A/?S\s*담당\s*[:：]?\s*([^\n●]*)")
RE_CAMP = re.compile(r"캠프\s*(?:이름|명)\s*[:：]?\s*([^\n●]*)")
RE_TITLE = re.compile(r"♣\s*[［\[]([^\]］]+)[\]］]")

# ★ 캠프 **담당자**(쿠팡 쪽 사람) — 우리 기사(`A/S 담당`)와 다른 사람이다.
#   접수 양식이 이렇게 생겼다(실측 2026-08-18 · 밴드 글 6,246건):
#       ● 캠프이름 : 울산2캠프
#       ● 캠프주소 : 울산광역시 …
#       ● 현장책임 : 제석화
#       ● 담당번호 : 010-7532-8543
#       ● 안전관리 : 이상협
#       ● 담당번호 : 010-2511-4947 / waynelee@coupang.com
#   ⚠ `담당번호` 가 **한 글에 두 번** 나온다. 순서로 고르면 언젠가 어긋나므로
#     **이름 줄 바로 다음 줄**의 번호만 그 사람 것으로 본다(짐작 금지 · [172]).
#   ⚠ 번호 칸에 이메일이 붙어 오므로 전화·메일을 갈라 담는다.
# ★ **메일이 오는 자리는 둘이다** (2026-08-18 실측 — 한쪽만 읽고 있었다).
#   ① `● 담당번호 : 010-2511-4947 / waynelee@coupang.com` — 번호 줄에 붙어 온다 (1,414줄)
#   ② `● E- MAIL : nominchul@coupangls.com` — **바로 다음 별도 줄**로 온다 (약 2,750줄)
#   ②를 안 읽어서 실측 현장책임 407명 중 메일이 **1명**뿐이었다 — 원본에는 있는데
#   화면이 비어 있었으니 '메일이 없는 캠프'로 보였다([165]: 안 읽은 칸은 빈칸과 같다).
#   라벨 표기가 제각각이라(`E- MAIL`·`E - MAIL`·`E-MAIL`·`E-mail`·`이메일`·`이- 메일`·
#   `email`) 공백·하이픈을 흘려 읽는다. **그 줄이 없으면 없는 대로 둔다** — 다음 사람의
#   메일을 끌어오면 대표 보고에 엉뚱한 주소가 박힌다(못 채우는 것보다 나쁘다 · [172]).
_MAIL_LINE = (r"(?:\n\s*●?\s*(?:E\s*[-–]?\s*MAIL|이\s*[-–]?\s*메일|이메일|메일\s*주소)"
              r"\s*[:：]?\s*([^\n●]*))?")
# ★ 라벨은 `현장책임`·`현장책임자`·`현장책임자명` 셋 다 온다 (2026-08-18 실측).
#   `자`·`자명` 을 안 먹으면 그것이 **이름 자리로 새어** `자 : One님`·`자명 : Turner님`
#   이 대표 보고에 그대로 실린다(실측 2건). 그렇다고 통째로 옵셔널하게 두면 이번엔
#   `현장책임 : 자명수` 같은 **멀쩡한 이름의 앞 글자를 라벨이 먹는다** — 못 읽는 것보다
#   나쁘다([172]). 그래서 뒤에 **콜론이 따라올 때만** 라벨로 친다(lookahead).
_MGR_SUFFIX = r"(?:\s*자\s*명?(?=\s*[:：]))?"
RE_SITE_MGR = re.compile(
    r"현장\s*책임" + _MGR_SUFFIX +
    r"\s*[:：]?\s*([^\n●]*)\n\s*●?\s*담당\s*(?:번호|자폰|자\s*폰)\s*[:：]?\s*([^\n●]*)"
    + _MAIL_LINE, re.IGNORECASE)
RE_SAFE_MGR = re.compile(
    r"안전\s*관리" + _MGR_SUFFIX +
    r"\s*[:：]?\s*([^\n●]*)\n\s*●?\s*담당\s*(?:번호|자폰|자\s*폰)\s*[:：]?\s*([^\n●]*)"
    + _MAIL_LINE, re.IGNORECASE)
# ★ **옛 양식은 한 사람이다** (2026-08-18 실측 — 안 읽고 있었다).
#   2023~2024 년 글은 현장책임·안전관리로 갈리기 전이라 이렇게 생겼다:
#       ● 캠프주소 : 경상남도 창원시 …
#       ● 담당자명 : 김재민
#       ● 담당번호 : 010-2003-3349
#   실측 캐시 7,813글 중 **옛 양식 2,714 · 새 양식 2,177**(둘 다 99). 즉 **옛 것이 더
#   많은데 위 두 정규식이 통째로 못 읽어** 그 캠프들이 전부 '전화 모름'으로 보였다 —
#   빈칸과 구별이 안 되는 종류의 잘못이다([165]).
# ⚠ **현장책임 칸에 넣지 않는다.** 원문은 이 사람이 현장책임인지 안전관리인지 **말한
#   적이 없다.** 넣으면 대표 보고에 근거 없는 직책이 박힌다 — 못 채우는 것보다 나쁘다
#   ([172]). 그래서 `담당자` 라는 제 칸으로 담고, 화면이 '직책 미상'으로 밝힌다.
# ⚠ `담당자명` 은 `담당번호`·`담당자폰` 과 겹치지 않는다(`명` 으로 끝난다).
#   `A/S 담당`(우리 기사)과도 다르다 — 그쪽은 RE_TECH 가 따로 읽는다.
RE_OLD_MGR = re.compile(
    r"담당\s*자\s*명\s*[:：]?\s*([^\n●]*)\n\s*●?\s*담당\s*(?:번호|자폰|자\s*폰)\s*[:：]?\s*([^\n●]*)"
    + _MAIL_LINE, re.IGNORECASE)
RE_CAMP_ADDR = re.compile(r"캠프\s*주소\s*[:：]?\s*([^\n●]*)")
RE_TEL = re.compile(r"01[016789][-. ]?\d{3,4}[-. ]?\d{4}")
RE_MAIL = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")


def _tel_fmt(raw):
    """번호 표기를 고른다. 자릿수가 아는 모양(10·11)일 때만 끊고, 아니면 원문 그대로
    둔다 — 모르는 모양을 억지로 끊으면 틀린 번호가 확정처럼 보인다."""
    d = re.sub(r"\D", "", raw)
    if len(d) == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    if len(d) == 10:
        return f"{d[:3]}-{d[3:6]}-{d[6:]}"
    return re.sub(r"[.\s]", "-", raw)


def _person(m):
    """(이름, 전화, 메일) — **셋 다 있는 것만 참으로 치지 않는다.**
    이름만 적힌 캠프도 있고 번호만 적힌 캠프도 있다. 없는 칸은 빈 문자열로 두고
    **지어내지 않는다** — 빈 칸은 사람이 채울 수 있지만 틀린 번호는 아무도 못 고친다."""
    if not m:
        return None
    name = (m.group(1) or "").strip(" :：-·").strip()
    rest = m.group(2) or ""
    tel = RE_TEL.search(rest)
    # 메일은 번호 줄(①)이 먼저고, 없으면 바로 다음 `E- MAIL` 줄(②)이다. 둘 다 없으면
    # **빈 칸으로 둔다** — 아래 사람의 메일을 끌어오지 않는다([172]).
    mail = RE_MAIL.search(rest)
    if not mail and m.lastindex and m.lastindex >= 3:
        mail = RE_MAIL.search(m.group(3) or "")
    # 표기만 고른다 — 숫자는 한 자도 안 건드린다. 원문에 `01038797505`·`010-71091794`
    # 처럼 하이픈이 없거나 어긋난 것이 실측 6건 있었고, 그대로 두면 대표 보고에서
    # 사람이 번호를 잘못 끊어 읽는다.
    tel = _tel_fmt(tel.group(0)) if tel else ""
    if not name and not tel:
        return None
    return {"이름": name, "전화": tel, "메일": mail.group(0) if mail else ""}
TECHS = ("김준형", "권오철", "김필우", "차동호", "김경원")
# 밴드 원문에서 실제 확인된 오탈자. 원문 캐시는 보존하되 구조화 결과에는
# 기준 기사명만 기록해 관리대장·기사별 집계로 오탈자가 전파되지 않게 한다.
TECH_ALIASES = {
    "권오절": "권오철",
    "권오처르": "권오철",
}


# 사람 이름처럼 생겼지만 사람이 아닌 것들. 알고리즘으로는 '하이테크'와 '엄진언'을
# 가를 수 없어 목록으로 둔다(2026-07-28 실데이터에서 확인).
VENDORS = ("하이테크", "대신택배")
# 이름에 붙여 쓴 직책 — '김승기기장' 처럼 띄어쓰기 없이 오는 경우가 있다
TITLES = ("기장", "기사", "과장", "차장", "부장", "대리", "팀장", "소장", "실장", "반장")
# 미배정을 뜻하는 자리표시자
PLACEHOLDER = re.compile(r"^(0{2,}|-+|자\)|미배정|미정|없음)$")


def _looks_like_name(t):
    return bool(re.fullmatch(r"[가-힣]{2,4}", t))


def normalize_tech(raw, when=""):
    """기사명만 남긴다 — 설명·직책·업체·자리표시자는 걸러낸다.

    ★ 2026-07-28 실사고: 이 함수가 '첫 조각'을 그대로 통과시켜
      `000 (캠프상태확인 및 스케쥴 세팅)`·`자) - 각캠프담당자 …` 같은 **작업 메모가
      담당기사 칸에 그대로 들어갔다.** 대표보고 'TOP 5' 에 `담당: 000 (…)` 로 노출됐고,
      기사별 집계도 오염됐다. 걸러 주는 clean_tech 는 리포트에만 쓰이고 있었다.
      버린 원문은 호출 쪽에서 비고에 남긴다(tech_note) — 정보를 없애지는 않는다."""
    cleaned = str(raw or "").strip()
    for wrong, right in TECH_ALIASES.items():
        cleaned = cleaned.replace(wrong, right)
    # ★ 그만둔 사람 이름이 양식 문구를 타고 담당 칸에 들어온다(2026-08-08 실측 35건).
    #   원문은 두고 **읽을 때만** 지금 담당자로 옮긴다 — people_alias 가 근거를 갖는다.
    for _h in _ALIAS.HANDOVERS:
        if _h["before"] in cleaned:
            cleaned = cleaned.replace(_h["before"],
                                      _ALIAS.resolve_person(_h["before"], when=when))
    tech = ", ".join(t for t in TECHS if t in cleaned)
    if tech:
        return tech
    # ★ 조각을 하나씩 보면 '스케쥴'·'체크' 같은 낱말도 이름처럼 생겨서 통과한다.
    #   그래서 **칸 전체를 먼저 판단한다** — 낱말이 넷 이상이면 이름이 아니라 문장이다.
    cleaned = re.sub(r"\(.*?\)", " ", cleaned)          # 괄호 설명은 통째로 뗀다
    cleaned = re.sub(r"\.{2,}.*$", " ", cleaned)
    tokens = [t for t in re.split(r"[,·+/\s]+", cleaned) if t]
    if not tokens or len(tokens) > 3:
        return ""
    names = []
    for part in tokens:
        for t in TITLES:                                # 김승기기장 → 김승기
            if part.endswith(t) and _looks_like_name(part[: -len(t)]):
                part = part[: -len(t)]
                break
        if PLACEHOLDER.match(part) or part in VENDORS:
            continue
        if _looks_like_name(part):
            names.append(part)
    return ", ".join(dict.fromkeys(names))


# ── 접수 취소 ────────────────────────────────────────────────────────────
# 사용자 지시(2026-08-08): **"접수 했다가 접수 취소하는 경우도 많은데 이것도
# 잡아내는 알고리즘 추가해"**
#
# 실제 사례(밴드 댓글): "통화 완료 했습니다 / 작동 원활함. 접수 취소 하세요"
#
# ★ 예전 규칙은 `"접수취소" in 본문` **한 줄**이었다. 두 군데서 새어 나갔다:
#   ① **띄어쓰기** — 사람은 '접수 취소'라고 쓴다. 붙여 쓴 것만 잡고 있었다.
#   ② **자리** — 취소는 본문이 아니라 **댓글**로 온다. 접수 글은 이미 올라간
#      뒤이므로 고칠 것이 댓글밖에 없다. 본문만 보면 영영 못 본다.
#   그래서 취소된 건이 '돌발AS 미처리'로 남아 AS 미실시 숫자를 계속 부풀렸다.
# ★★ '취소'라는 낱말만으로 판정하면 **멀쩡한 건이 죽는다.** 실측에서 그대로 나왔다:
#      "바디부분 아크릴판은 캠프담당 취소요청함" · "택배발송 취소요청하심"
#    둘 다 부품·택배 취소지 AS 접수 취소가 아니다. 취소로 처리하면 그 현장은
#    아무도 안 가는데 목록에서도 사라진다 — 미실시로 남는 것보다 나쁘다.
#    그래서 **'취소' 곁에 '접수'나 A/S 가 있을 때만** 접수 취소로 본다.
#    ★ '취소' 곁에 A/S 가 있으면 되게 했더니 **모든 글이 걸렸다** — 밴드 양식에
#      `● A/S 완료 :` 줄이 늘 따라붙기 때문이다(실측 2620: '택배발송 취소요청하심'
#      바로 뒤가 그 줄이었다). 그래서 근거는 **'접수'가 '취소'에 붙어 있는 것** 하나다.
_CANCEL = re.compile(
    r"(접\s*수\s*[^가-힣A-Za-z0-9\n]{0,3}(를|건|은|이)?\s*(요\s*청)?\s*취\s*소"  # 접수(를/건) 취소
    r"|취\s*소\s*[^가-힣A-Za-z0-9\n]{0,3}(된|할|하실)?\s*접\s*수"                # 취소 접수
    r"|오\s*접\s*수"                                   # 오접수
    r"|중\s*복\s*접\s*수"                              # 중복접수
    r"|접\s*수\s*(를\s*)?(철\s*회|반\s*려))")          # 접수 철회/반려
# 취소가 아닌데 '취소'가 들어간 말 — 걸러 내지 않으면 멀쩡한 건이 취소로 죽는다.
_NOT_CANCEL = re.compile(r"(접\s*수\s*취\s*소\s*(불\s*가|없|아\s*님|안\s*됨|보\s*류)"
                         r"|접\s*수(?:는|를)?\s*취\s*소\s*(하\s*지|안\s*함)"
                         r"|(?:보\s*험|택\s*배|부\s*품|예\s*약|주\s*문|발\s*송)\s*접\s*수\s*취\s*소"
                         r"|접\s*수\s*유\s*지|취\s*소\s*된\s*건\s*없)")

# 취소를 **되돌린** 말도 하나의 상태 사건이다. 예전처럼 댓글을 한 덩어리로
# 합친 뒤 `_NOT_CANCEL.search()`를 하면, 오래된 "접수 취소"와 나중의 "접수 유지"가
# 같은 문자열에 있다는 이유만으로 둘 다 사라졌다. 이제 문장별·댓글별 사건을 시간순으로
# 놓고 마지막 명시 상태가 이긴다(2026-08-11 엄격검토).
_ACTIVE = re.compile(
    r"(접\s*수\s*(?:는\s*)?(?:유\s*지|계\s*속|진\s*행|재\s*개)"
    r"|접\s*수(?:는|를)?\s*취\s*소\s*(?:하\s*지\s*않|안\s*함|안\s*하|아\s*님|불\s*가|안\s*됨|보\s*류)"
    r"|접\s*수\s*취\s*소\s*(?:철\s*회|해\s*제|번\s*복)"
    r"|취\s*소\s*(?:철\s*회|해\s*제|번\s*복)"
    r"|취\s*소\s*된\s*건\s*없)"
)
_NON_SERVICE_CANCEL = re.compile(
    r"(?:보\s*험|택\s*배|부\s*품|예\s*약|주\s*문|발\s*송)\s*접\s*수\s*취\s*소"
)


def cancel_state(text):
    """한 본문/댓글 안의 마지막 명시 상태: ``cancel``·``active``·빈 문자열.

    보험·택배·부품 접수 취소는 서비스 접수 상태가 아니므로 사건에서 제외한다.
    같은 댓글 안에 정정이 함께 적혀도 **뒤에 쓴 말**이 이긴다. 이 함수는 한 댓글만
    판단하고, 댓글 사이의 시간 순서는 :func:`latest_cancel_event`가 맡는다.
    """
    s = str(text or "")
    if not s:
        return ""
    events = []
    excluded = list(_NON_SERVICE_CANCEL.finditer(s))
    for match in _CANCEL.finditer(s):
        # "보험접수 취소"처럼 다른 종류의 접수 취소와 겹친 매치는 버린다.
        if any(not (match.end() <= bad.start() or match.start() >= bad.end())
               for bad in excluded):
            continue
        # 바로 뒤의 부정·보류는 취소가 아니라 active 사건으로 아래 _ACTIVE가 잡는다.
        tail = s[match.start():min(len(s), match.end() + 16)]
        if _NOT_CANCEL.search(tail):
            continue
        events.append((match.start(), match.end(), "cancel"))
    for match in _ACTIVE.finditer(s):
        events.append((match.start(), match.end(), "active"))
    if not events:
        return ""
    events.sort(key=lambda item: (item[0], item[1]))
    return events[-1][2]


def cancel_hit(text):
    """이 글/댓글이 **접수 취소**를 말하고 있나.

    부품·택배 취소와 갈라야 한다(위 주석). 근거는 '접수'가 '취소'에 붙어 있는 것,
    그리고 오접수·중복접수·접수철회다. 그냥 '취소'는 판정하지 않는다 —
    취소로 잘못 처리하면 그 현장은 아무도 안 가는데 목록에서도 사라진다.
    """
    return cancel_state(text) == "cancel"


def _event_epoch(value):
    """밴드 밀리초·초·ISO 시각을 정렬 가능한 epoch 초로 바꾼다."""
    if value in (None, ""):
        return None
    try:
        number = float(value)
        return number / 1000.0 if number > 10_000_000_000 else number
    except (TypeError, ValueError):
        pass
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).timestamp()
    except (TypeError, ValueError):
        return None


def latest_cancel_event(post):
    """본문과 댓글을 시간순으로 놓고 마지막 접수상태 사건을 돌려준다.

    ``{state, source, text, created_at, epoch, order}`` 모양이다. 시각 없는 댓글은
    수집 단계의 원칙과 같이 버린다. 본문은 게시시각, 댓글은 댓글시각을 사용하므로
    오래된 취소 댓글 뒤의 최신 "접수 유지"가 정상적으로 취소를 해제한다.
    """
    if not isinstance(post, dict):
        return None
    events = []
    body = str(post.get("content") or "")
    body_state = cancel_state(body)
    body_epoch = _event_epoch(post.get("created_at"))
    if body_state:
        events.append({"state": body_state, "source": "본문", "text": body,
                       "created_at": post.get("created_at"), "epoch": body_epoch,
                       "order": 0})
    for index, comment in enumerate(post.get("comments") or [], 1):
        if not isinstance(comment, dict):
            continue
        text = str(comment.get("content") or comment.get("body") or "")
        state = cancel_state(text)
        epoch = _event_epoch(comment.get("created_at"))
        if not state or epoch is None:  # 시각 없는 댓글로 순서를 지어내지 않는다.
            continue
        events.append({"state": state, "source": "댓글", "text": text,
                       "created_at": comment.get("created_at"), "epoch": epoch,
                       "order": index})
    if not events:
        return None
    # 본문 시각이 결측이면 알려진 댓글보다 앞선 것으로만 취급한다.
    events.sort(key=lambda event: (
        event.get("epoch") is not None,
        event.get("epoch") if event.get("epoch") is not None else float("-inf"),
        event.get("order", 0),
    ))
    return events[-1]


def comment_text(post):
    """글에 달린 댓글 본문을 한 덩어리로 — 캐시에 댓글이 없으면 빈 문자열.

    ★ 캐시 모양은 `comments: [{author, created_at, content}]` 하나다 (2026-08-08부터
      화면 긁기·API 양쪽이 같은 모양으로 담는다). 담는 쪽이 둘이라 **읽는 쪽은
      반드시 하나**여야 한다 — 갈리면 한쪽만 고쳐지고 다른 쪽은 조용히 옛것으로 남는다.
      수집 자체는 'CSOS 리서치 및 자료 수집' 세션 몫이다(CLAUDE.md).
      적힌 수만큼 못 읽은 글은 `cancel_blind_count()` 가 센다 —
      "댓글은 있는데 못 읽는다"를 조용히 넘기지 않기 위해서다.
    """
    if not isinstance(post, dict):
        return ""
    out = []
    for c in (post.get("comments") or []):
        if isinstance(c, dict):
            out.append(str(c.get("content") or c.get("body") or ""))
        else:
            out.append(str(c or ""))
    return "\n".join(out)


def cancel_blind_count(posts):
    """댓글이 달렸는데 **다 못 읽은** 글 수 — 취소를 놓칠 수 있는 사각지대의 크기.

    ★ '하나도 못 읽음'이 아니라 '적힌 수만큼 못 읽음'으로 센다 (2026-08-08).
      접힌 댓글을 한 개만 펴서 담은 글은 본문이 있으니 예전 기준으로는 안 걸렸는데,
      정작 취소 통보는 **못 편 그 댓글**일 수 있다. 반쯤 읽은 것을 다 읽은 것으로
      세면 사각지대가 0으로 보인다 — 제일 나쁜 종류의 안심이다.

    ★★ 그런데 그 세는 법마저 `comment_count` 에 기대고 있었다 (2026-08-08 저녁 실측).
      캐시 10,312글 중 `comment_count>0` 은 **6글**이고 댓글 본문은 **0글**이다.
      밴드에 댓글이 없어서가 아니라 **수집기가 그 숫자를 안 담아서**다. 그래서
      사각지대 계기가 "사각지대 0" 이라고 말했다 — **재는 도구가 같은 결측에 눈이
      멀어 있었다.** 없는 것과 안 본 것을 구별하지 못하면 계기는 늘 안심을 준다.
      이제 셋을 가른다:
        · `comments` 키가 **아예 없다**  → 한 번도 안 들여다봤다 → **사각지대**
        · `comments: []`                → 보긴 봤고 없었다 → 사각지대 아님
        · 적힌 수보다 적게 담겼다        → 반쯤 읽었다 → 사각지대(예전 기준)
    """
    n = 0
    for p in (posts or {}).values():
        if not isinstance(p, dict):
            continue
        if "comments" not in p:
            # 들여다본 적이 없다. comment_count 가 0 이어도 그 0 을 믿을 근거가 없다.
            n += 1
            continue
        try:                       # 캐시에 따라 숫자가 문자열로 들어 있다
            cnt = int(str(p.get("comment_count") or 0).strip() or 0)
        except ValueError:
            cnt = 0
        if cnt <= 0:
            continue
        got = len([c for c in (p.get("comments") or []) if isinstance(c, dict)])
        if got < cnt and not p.get("comments_full"):
            n += 1
    return n


def tech_note(raw, kept):
    """정규화하며 버린 부분 — 호출 쪽이 비고에 남겨 정보를 잃지 않게 한다."""
    raw = str(raw or "").strip()
    return "" if not raw or raw == kept else raw


def parse_post(no, p, band):
    c = p.get("content") or ""
    prj = RE_PRJ.search(c)
    title = (RE_TITLE.search(c).group(1).strip() if RE_TITLE.search(c) else "")
    if not prj and not title:
        return None                       # 업무 게시글이 아님(공지·자료 등)

    md = RE_DATE.search(c)
    work_date = ""
    if md:
        y, mo, d = int(md.group(1)), int(md.group(2)), int(md.group(3))
        work_date = f"{y:04d}-{mo:02d}-{d:02d}" if mo and d else ""   # 2026.00.00 = 미정

    prj_no = prj.group(1) if prj else ""
    if prj_no and set(prj_no[2:]) == {"0"}:      # UJ000000 = 양식 템플릿 게시글
        return None

    ts = p.get("created_at")
    posted = datetime.fromtimestamp(ts / 1000).strftime("%Y-%m-%d") if ts else ""

    tech_raw = (RE_TECH.search(c).group(1).strip() if RE_TECH.search(c) else "")
    # 게시일을 함께 넘긴다 — 인계 **전** 글은 그때 담당자 그대로 둬야 한다.
    tech = normalize_tech(tech_raw, when=posted)

    camp = (RE_CAMP.search(c).group(1).strip() if RE_CAMP.search(c) else "")
    camp = re.sub(r"\s*\.{3}더보기.*$", "", camp).strip()

    # 업무유형·유상무상·상태
    if "정기점검" in title or "3개월" in title or "분기" in title:
        kind = "정기점검"
    elif "돌발" in title:
        kind = "돌발AS"
    elif "설치" in title or "납품" in title:
        kind = "신규납품설치"
    else:
        kind = "기타"
    if "동시" in title or "동시진행" in c:
        kind += "(동시진행)"
    cost = "유상" if "유료" in title else ("무상" if "무료" in title else "")
    # ★ 본문과 댓글을 한 문자열로 합치지 않는다. 각 댓글의 시각을 읽어 마지막 명시
    #   상태가 이긴다. 단 완료 제목과 같은 시각의 본문 취소는 기존 안전규칙대로 완료에
    #   양보한다(실측 4979: 일부 작업만 취소하고 실제 작업은 완료).
    state_event = latest_cancel_event(p)
    if state_event and state_event["state"] == "cancel" and state_event["source"] == "댓글":
        status = "취소"
    elif "완료" in title:
        status = "작업완료"
    elif state_event and state_event["state"] == "cancel":
        status = "취소"
    elif "안내" in title:
        status = "접수·예정"
    else:
        status = ""

    docs = [d for d, kw in (("판매전표", "판매전표"), ("거래명세서", "거래명세서"),
                            ("견적서", "견적서"), ("메일발송", "메일발송")) if kw in c]
    # ★ 캠프 담당자는 **기존 칸을 하나도 안 건드리고** 옆에 더한다 — 이 반환값을 읽는
    #   곳이 여럿이다(cancel_watch · cross_signal · app_server 완료색인).
    addr = RE_CAMP_ADDR.search(c)
    return {"프로젝트NO": prj_no, "업무유형": kind, "비용구분": cost,
            "작업일": work_date, "담당기사": tech, "캠프명": camp, "진행상태": status,
            "문서상태": "+".join(docs), "사진": p.get("photo_count", 0),
            "게시일": posted, "밴드": band, "게시글": no,
            "캠프주소": (addr.group(1).strip() if addr else ""),
            "현장책임": _person(RE_SITE_MGR.search(c)),
            "안전관리": _person(RE_SAFE_MGR.search(c)),
            # 옛 양식(직책 미상) — 위 둘과 **다른 칸**이다. 기존 칸은 안 건드린다.
            "담당자": _person(RE_OLD_MGR.search(c))}


def load_records():
    out = []
    for f in glob.glob(os.path.join(CACHE_DIR, "*.json")):
        b = os.path.basename(f)
        if b.startswith(("raw_", "dump_")):
            continue
        d = json.load(open(f, encoding="utf-8"))
        band = d.get("band_name", b)
        for no, p in d.get("posts", {}).items():
            r = parse_post(no, p, band)
            if r:
                out.append(r)
    out += load_kakao_records()
    out.sort(key=lambda r: (r["작업일"] or r["게시일"], r["프로젝트NO"]))
    return out


def load_kakao_records():
    """카톡 내보내기(.txt)도 같은 양식(♣ ［…] ● 프로젝트NO / ● 캠프이름)을 쓴다.

    밴드에 안 올라오고 카톡에만 보고된 건이 있어(2026-07-27 기준 39건) 함께 읽는다.
    한 메시지에 여러 건이 담기므로 ♣ 로 덩어리를 나눠 게시글처럼 취급한다.
    """
    inbox = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kakao", "inbox")
    if not os.path.isdir(inbox):
        return []
    DAY = re.compile(r"-{3,}\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일")
    out, seen = [], set()
    for f in sorted(glob.glob(os.path.join(inbox, "*.txt"))):
        room = os.path.splitext(os.path.basename(f))[0]
        try:
            txt = open(f, encoding="utf-8", errors="replace").read()
        except OSError:
            continue
        day = ""
        for chunk in re.split(r"(?=-{3,}\s*\d{4}년)", txt):
            m = DAY.search(chunk)
            if m:
                day = "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
            for i, blk in enumerate(re.split(r"♣", chunk)[1:]):
                if "프로젝트NO" not in blk:
                    continue
                key = (day, blk[:200])
                if key in seen:
                    continue
                seen.add(key)
                r = parse_post(f"kakao-{room}-{day}-{i}",
                               {"content": "♣" + blk[:2000], "author": room,
                                "created_at": None, "photo_count": 0, "comment_count": 0},
                               f"카톡 {room}")
                if r:
                    if not r.get("게시일"):
                        r["게시일"] = day
                    out.append(r)
    return out


def ledger_projects(master):
    """원장에 이미 있는 프로젝트NO 집합 (02·04·05·06 시트)"""
    import openpyxl
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    seen = set()
    for sh in ("02_돌발AS접수", "04_정기점검", "05_신규납품설치", "06_거래서류청구수금"):
        if sh not in wb.sheetnames:
            continue
        ws = wb[sh]
        hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        try:
            j = [i for i, h in enumerate(hdr) if str(h).strip() == "프로젝트NO"][0]
        except IndexError:
            continue
        for row in ws.iter_rows(min_row=5, values_only=True):
            if j < len(row) and row[j]:
                seen.add(str(row[j]).strip())
    wb.close()
    return seen


HEADERS = ["프로젝트NO", "업무유형", "비용구분", "작업일", "담당기사", "캠프명",
           "진행상태", "문서상태", "사진", "원장등록", "게시일", "밴드"]
WIDTHS = [13, 17, 9, 12, 14, 22, 11, 26, 6, 10, 12, 24]


def main():
    args = sys.argv[1:]
    month = args[args.index("--month") + 1] if "--month" in args else None
    recs = load_records()
    if month:
        recs = [r for r in recs if (r["작업일"] or r["게시일"]).startswith(month)]

    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    known = ledger_projects(master)
    for r in recs:
        r["원장등록"] = "등록됨" if r["프로젝트NO"] in known else "미등록"

    new = [r for r in recs if r["원장등록"] == "미등록" and r["프로젝트NO"]]
    os.makedirs(REPORT_DIR, exist_ok=True)
    tag = month or "전체"
    base = os.path.join(REPORT_DIR, f"밴드업무추출_{tag}_{datetime.now():%Y%m%d_%H%M}")
    with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HEADERS); w.writeheader()
        w.writerows([{k: r.get(k, "") for k in HEADERS} for r in recs])

    from collections import Counter
    ck, cs = Counter(r["업무유형"] for r in recs), Counter(r["진행상태"] for r in recs)
    ct = Counter(r["담당기사"] for r in recs if r["담당기사"])
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write(f"# 밴드 업무 추출 — {tag}\n\n")
        f.write(f"- 생성 {datetime.now():%Y-%m-%d %H:%M} / 추출 {len(recs)}건 "
                f"(원장 등록됨 {len(recs)-len(new)} · **미등록 {len(new)}**)\n")
        f.write(f"- 업무유형: {dict(ck)}\n- 진행상태: {dict(cs)}\n- 담당기사: {dict(ct)}\n\n")
        f.write("## 원장 미등록 건 (백필 후보)\n\n")
        f.write("| 프로젝트NO | 유형 | 비용 | 작업일 | 기사 | 캠프 | 상태 | 문서 |\n|---|---|---|---|---|---|---|---|\n")
        for r in new:
            f.write(f"| {r['프로젝트NO']} | {r['업무유형']} | {r['비용구분']} | {r['작업일']} | "
                    f"{r['담당기사']} | {r['캠프명']} | {r['진행상태']} | {r['문서상태']} |\n")

    print(f"추출 {len(recs)}건 (등록됨 {len(recs)-len(new)} / 미등록 {len(new)})")
    print(f"유형 {dict(ck)}")
    print("리포트:", base + ".md")

    if "--sheet" in args:
        from findings_sheet import upsert, build_generic_sheet
        xml = build_generic_sheet(
            "24_밴드업무추출", HEADERS, WIDTHS,
            [[r.get(k, "") for k in HEADERS] for r in recs],
            f"[사용법] 밴드 게시글에서 자동 추출한 업무 원천({tag}). '원장등록=미등록' 행이 백필 후보입니다. "
            f"에이전트가 갱신하며 수기 입력은 하지 마세요.")
        dst, msg = upsert(master, xml, sheet_name="24_밴드업무추출", headers=HEADERS)
        print(f"24_밴드업무추출: {msg}")
        if dst:
            print("   ", dst)


if __name__ == "__main__":
    main()
