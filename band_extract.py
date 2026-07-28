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


def normalize_tech(raw):
    """기사명 오탈자를 기준 이름으로 정규화하고 지원 인원 설명은 제외한다."""
    cleaned = str(raw or "").strip()
    for wrong, right in TECH_ALIASES.items():
        cleaned = cleaned.replace(wrong, right)
    tech = ", ".join(t for t in TECHS if t in cleaned)
    if tech:
        return tech
    # 사내 기사 목록 밖 이름도 증거로 보존하되 첫 이름 조각만 사용한다.
    tech = re.split(r"[,·]", cleaned)[0].strip()
    tech = re.sub(r"\.{2,}.*$", "", tech).strip()
    return "" if tech in (".", "…", "...") else tech


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

    tech_raw = (RE_TECH.search(c).group(1).strip() if RE_TECH.search(c) else "")
    tech = normalize_tech(tech_raw)

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
    if "접수취소" in c:
        status = "취소"
    elif "완료" in title:
        status = "작업완료"
    elif "안내" in title:
        status = "접수·예정"
    else:
        status = ""

    docs = [d for d, kw in (("판매전표", "판매전표"), ("거래명세서", "거래명세서"),
                            ("견적서", "견적서"), ("메일발송", "메일발송")) if kw in c]
    ts = p.get("created_at")
    posted = datetime.fromtimestamp(ts / 1000).strftime("%Y-%m-%d") if ts else ""
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
