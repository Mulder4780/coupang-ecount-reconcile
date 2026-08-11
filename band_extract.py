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


def cancel_hit(text):
    """이 글/댓글이 **접수 취소**를 말하고 있나.

    부품·택배 취소와 갈라야 한다(위 주석). 근거는 '접수'가 '취소'에 붙어 있는 것,
    그리고 오접수·중복접수·접수철회다. 그냥 '취소'는 판정하지 않는다 —
    취소로 잘못 처리하면 그 현장은 아무도 안 가는데 목록에서도 사라진다.
    """
    s = str(text or "")
    if not s or _NOT_CANCEL.search(s):
        return False
    return bool(_CANCEL.search(s))


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
    # ★ 순서가 뜻을 가진다. 댓글은 글보다 **나중**에 달리므로 완료 글이라도 댓글의
    #   취소가 이긴다. 반대로 제목이 '완료'인 글의 본문에 나온 취소는 그 작업 얘기가
    #   아닐 때가 많다(실측 4979: '접수전 정기점검 취소되어 도어락만 교체진행됨' —
    #   작업은 실제로 했다). 그래서 본문 취소는 완료 제목에 양보한다.
    if cancel_hit(comment_text(p)):
        status = "취소"
    elif "완료" in title:
        status = "작업완료"
    elif cancel_hit(c):
        status = "취소"
    elif "안내" in title:
        status = "접수·예정"
    else:
        status = ""

    docs = [d for d, kw in (("판매전표", "판매전표"), ("거래명세서", "거래명세서"),
                            ("견적서", "견적서"), ("메일발송", "메일발송")) if kw in c]
    return {"프로젝트NO": prj_no, "업무유형": kind, "비용구분": cost,
            "작업일": work_date, "담당기사": tech, "캠프명": camp, "진행상태": status,
            "문서상태": "+".join(docs), "사진": p.get("photo_count", 0),
            "게시일": posted, "밴드": band, "게시글": no}


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
