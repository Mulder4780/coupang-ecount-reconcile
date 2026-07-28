# -*- coding: utf-8 -*-
"""
daily_brief.py — 대표 보고용 **내용 중심** 일일 브리핑을 원장에서 만든다
================================================================================
2026-07-28 대표 통화 지시(요지):

  "숫자를 나한테 보고하라는 게 아니야."
  · 돌발이 어제 몇 건 접수됐고 오늘 몇 건 처리할 건지
  · 처리한 건은 **무엇 때문에 갔는데 무슨 작업을 했는지**
  · **유상인데 무상 부분이 포함**됐거나, 돌발이지만 **무료로 해줄 수밖에 없었던** 사정
  · 갔더니 **뭐가 더 있어서 유상을 추가**한 건
  · 정기점검은 루틴대로 가고 있는지, 분기 몇 %인지, 마무리 전망
  · 정기점검 갔다가 **유료가 발생한** 건
  · 미처리 건이 있는지

즉 대표가 알고 싶은 것은 '건수'가 아니라 **판단이 필요한 사정**이다.
그 내용은 이미 원장에 있다(신청내용 100% · 유·무상 98% · 실제작업상세 86%).
여기서 그걸 문장으로 엮어 준다. 없는 건 지어내지 않고 '미기입'으로 남긴다.

  python daily_brief.py                # 어제(집계기준일) 기준 브리핑
  python daily_brief.py --date 2026-07-27
  python daily_brief.py --md           # reports/일일브리핑_YYYYMMDD.md
"""
import sys, os, re
from datetime import datetime, date, timedelta
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
FREE = ("무상", "보험")


def _d(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    m = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", str(v or ""))
    return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else ""


def _s(v):
    return str(v).strip() if v not in (None, "") else ""


def load(master=None):
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = master or resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)

    def rows(sh):
        if sh not in wb.sheetnames:
            return []
        ws = wb[sh]
        hdr = [_s(h) for h in next(ws.iter_rows(min_row=4, max_row=4, values_only=True))]
        ix = {h: i for i, h in enumerate(hdr) if h}
        out = []
        for r in ws.iter_rows(min_row=5, values_only=True):
            g = {h: (r[i] if i < len(r) else None) for h, i in ix.items()}
            if _s(g.get("프로젝트NO")):
                out.append(g)
        return out

    d = {"as": rows("02_돌발AS접수"), "pm": rows("04_정기점검"), "fw": rows("03_현장작업실적")}
    wb.close()
    return d, master


def brief(day=None, data=None):
    """하루치 브리핑. 반환값은 화면·이미지·문서 어디서나 같은 내용을 쓰도록 구조화한다."""
    data = data or load()[0]
    day = day or (date.today() - timedelta(days=1)).isoformat()
    A, P, F = data["as"], data["pm"], data["fw"]

    # 03시트(실제 작업 내용)를 프로젝트별로 붙여 둔다 — '무슨 작업을 했는지'가 여기 있다
    work = {}
    for r in F:
        k = _s(r.get("프로젝트NO"))
        if k and k not in work:
            work[k] = r

    def line(r, donec, kindc):
        """한 건을 '왜 갔고 무엇을 했고 유·무상은 어떻게'로 적는다."""
        prj = _s(r.get("프로젝트NO"))
        w = work.get(prj, {})
        why = _s(r.get("신청내용")) or _s(r.get("점검내용"))
        did = _s(w.get("실제작업상세")) or _s(w.get("실제작업항목"))
        cost = _s(r.get("유상·무상·보험")) or _s(w.get("비용구분"))
        extra = _s(r.get("최초접수외추가작업")) or _s(r.get("추가작업내용")) or _s(w.get("추가작업내용"))
        return {"프로젝트NO": prj, "캠프명": _s(r.get("캠프명")), "담당기사": _s(r.get("담당기사")),
                "왜": why, "무엇": did, "비용": cost, "추가작업": extra,
                "일자": _d(r.get(donec)), "구분": kindc,
                # ★ 사용자 지시(2026-07-28): "각 항목 옆에 날짜 붙여줘 — 금일이라고 하면
                #   어떤 날짜인지 헷갈림." 완료건은 '언제 접수돼 언제 끝났는지'가 같이 보여야
                #   대표가 "그거 언제 들어온 건데?" 를 되묻지 않는다.
                "접수일": _d(r.get("접수일자")) or _d(r.get("점검예정일")) or _d(r.get("요청일"))}

    # ── 돌발AS ──
    as_new = [r for r in A if _d(r.get("접수일자")) == day]
    as_done = [r for r in A if _d(r.get("작업완료일")) == day]
    # ★ '완료일이 없다'와 '아직 안 갔다'는 다르다. 오래된 건은 거의 다 **기록 누락**이다
    #   (2026-07-28 확인: 84건 중 46건이 5월 이전). 둘을 뭉뚱그려 '미처리 84건'이라고
    #   보고하면 대표가 놀라고, 반대로 '없다'고 하면 최근 건을 놓친다. 갈라서 말한다.
    _open = [r for r in A if not _d(r.get("작업완료일"))
             and _s(r.get("진행상태")) not in ("취소", "보류")]
    _cut = (datetime.strptime(day, "%Y-%m-%d").date() - timedelta(days=30)).isoformat()
    as_open = [r for r in _open if (_d(r.get("접수일자")) or "0000") >= _cut]   # 최근 = 진짜 미처리
    as_stale = [r for r in _open if (_d(r.get("접수일자")) or "0000") < _cut]   # 오래됨 = 완료일 미기입
    # ── 정기점검 ──
    pm_done = [r for r in P if _d(r.get("실제점검일")) == day]
    pm_plan = [r for r in P if _d(r.get("점검예정일")) == day]

    done = [line(r, "작업완료일", "돌발AS") for r in as_done] + \
           [line(r, "실제점검일", "정기점검") for r in pm_done]

    # 대표가 콕 집어 물은 것들
    free = [x for x in done if any(k in x["비용"] for k in FREE)]
    extra = [x for x in done if x["추가작업"]]
    pm_paid = [line(r, "실제점검일", "정기점검") for r in pm_done
               if _s(r.get("유상추가작업발생")) in ("Y", "예", "발생", "있음")]
    to_as = [line(r, "실제점검일", "정기점검") for r in pm_done
             if _s(r.get("돌발AS전환여부")) in ("전환", "Y", "예")]
    abnormal = [line(r, "실제점검일", "정기점검") for r in pm_done
                if _s(r.get("이상발견여부")) in ("있음", "Y", "예")]

    # 정기점검 진행률 — 분기 기준(대표가 "이 분기 몇 프로"를 물었다)
    y = int(day[:4]); q = (int(day[5:7]) - 1) // 3 + 1
    qs, qe = date(y, 3 * q - 2, 1), (date(y + (q == 4), (3 * q) % 12 + 1, 1) - timedelta(days=1))
    inq = [r for r in P if qs.isoformat() <= (_d(r.get("점검예정일")) or "9999") <= qe.isoformat()]
    inq_done = [r for r in inq if _d(r.get("실제점검일"))]

    # 내용이 비어 있으면 숨기지 않고 '미기입'으로 남긴다 — 채워야 할 칸이다
    blank = [x for x in done if not x["무엇"]]

    return {
        "기준일": day,
        "돌발AS": {"신규접수": len(as_new), "완료": len(as_done),
                    "미처리": len(as_open), "완료일미기입": len(as_stale)},
        "정기점검": {"예정": len(pm_plan), "완료": len(pm_done),
                     "분기": f"{y}년 {q}분기", "분기예정": len(inq), "분기완료": len(inq_done),
                     "분기진행률": round(len(inq_done) * 100 / len(inq)) if inq else 0},
        "완료내역": done, "무상건": free, "추가작업건": extra,
        "점검중유상": pm_paid, "AS전환": to_as, "이상발견": abnormal,
        "내용미기입": blank,
        "완료일미기입목록": [line(r, "접수일자", "돌발AS") for r in as_stale],
        "신규목록": [line(r, "접수일자", "돌발AS") for r in as_new],
    }


def md(s, base=""):
    """2026-07-27 → 07-27. 기준일과 해가 다르면 연도까지 적는다."""
    s = _s(s)
    if not s:
        return ""
    return s[5:] if base and s[:4] == base[:4] else s


def span(x, base=""):
    """완료건 꼬리표 — 접수 → 완료 경과일. 묵은 건이 눈에 띄게 한다."""
    a, b_ = x.get("접수일"), x.get("일자")
    if not (a and b_) or a == b_:
        return ""
    try:
        n = (datetime.strptime(b_, "%Y-%m-%d") - datetime.strptime(a, "%Y-%m-%d")).days
    except ValueError:
        return ""
    return f" (접수 {md(a, base)} · {n}일 만)" if n > 0 else ""


def tag(x, base=""):
    """항목 머리표 — '날짜 · 프로젝트NO · 캠프' 순서로 고정한다.

    ★ 사용자 지시(2026-07-28): "각 항목 옆에 날짜 붙여줘, 금일이라고 하면
      어떤 날짜인지 헷갈림." 어느 줄을 읽어도 날짜가 먼저 나오게 한다."""
    return f"{md(x.get('일자'), base) or '날짜미기입'} · {x.get('프로젝트NO') or '번호미기입'} · {x.get('캠프명') or '캠프미기입'}"


def text(b):
    """대표에게 그대로 읽어 드릴 수 있는 문장."""
    L = []
    d, a, p = b["기준일"], b["돌발AS"], b["정기점검"]
    wd = "월화수목금토일"[datetime.strptime(d, "%Y-%m-%d").weekday()]
    # ★ '금일·당일'이라는 말은 읽는 사람마다 다른 날을 떠올린다. 맨 위에 날짜를 못 박는다.
    L.append(f"[{d}({wd}) 실적 — 아래 날짜는 모두 실제 날짜입니다]")
    L.append(f"■ 돌발AS — 신규 접수 {a['신규접수']}건 · 완료 {a['완료']}건 · "
             f"미처리 {a['미처리']}건(최근 30일)")
    # ★ 신규와 완료가 같은 모양으로 나오면 어느 게 처리된 건지 안 보인다.
    #   대표가 묻는 건 "완료한 건 무슨 작업을 했느냐"이므로 완료는 따로, 더 자세히 적는다.
    if b["신규목록"]:
        L.append(f"  ▸ 새로 접수 {len(b['신규목록'])}건  (접수일 {md(d, d)})")
        for x in b["신규목록"][:6]:
            L.append(f"      {tag(x, d)}"
                     + (f" · {x['담당기사']}" if x["담당기사"] else " · 기사 미배정"))
            L.append(f"         내용 : {x['왜'][:52] or '접수내용 미기입'}")

    doneA = [y for y in b["완료내역"] if y["구분"] == "돌발AS"]
    if doneA:
        L.append(f"  ▸ 완료 {len(doneA)}건 — 무엇 때문에 갔고 무슨 작업을 했는지  (완료일 {md(d, d)})")
        for x in doneA[:8]:
            head = (f"      {tag(x, d)}{span(x, d)}"
                    f" · {x['담당기사'] or '기사 미기입'} · {x['비용'] or '비용 미기입'}")
            L.append(head)
            L.append(f"         왜   : {x['왜'][:52] or '접수내용 미기입'}")
            # ★ 작업내용이 비었으면 조용히 넘기지 않는다. 대표가 그 자리에서 물어볼 항목이라
            #   '없다'는 사실 자체를 알고 있어야 한다.
            L.append(f"         작업 : {x['무엇'][:52]}" if x["무엇"]
                     else "         작업 : ★ 미기입 — 기사에게 확인해 03_현장작업실적에 입력 필요")
            if x["추가작업"]:
                L.append(f"         추가 : {x['추가작업'][:48]}")

    L.append(f"\n■ 정기점검 — {md(d, d)} 완료 {p['완료']}건 (그날 예정 {p['예정']}건)")
    L.append(f"   {p['분기']} 진행률 {p['분기진행률']}% ({p['분기완료']}/{p['분기예정']}건)")
    if p["분기예정"]:
        L.append("   " + ("특별한 문제 없으면 분기 내 마무리 가능합니다."
                          if p["분기진행률"] >= 60 else "진행률이 낮아 일정 관리가 필요합니다."))

    if b["점검중유상"]:
        L.append(f"\n■ 정기점검 갔다가 유상 발생 {len(b['점검중유상'])}건")
        for x in b["점검중유상"][:5]:
            L.append(f"   {tag(x, d)} · {x['추가작업'][:36] or '내용 미기입'}")
    if b["무상건"]:
        L.append(f"\n■ 무상·보험 처리 {len(b['무상건'])}건 — 사유 확인 필요")
        for x in b["무상건"][:5]:
            L.append(f"   {tag(x, d)} · {x['왜'][:30]} [{x['비용']}]")
    if b["추가작업건"]:
        L.append(f"\n■ 접수 외 추가작업 {len(b['추가작업건'])}건")
        for x in b["추가작업건"][:5]:
            L.append(f"   {tag(x, d)} · {x['추가작업'][:40]}")
    if b["AS전환"]:
        L.append(f"\n■ 점검 중 돌발AS 전환 {len(b['AS전환'])}건")
    if a.get("완료일미기입"):
        # 접수일 범위를 적어 '언제부터 밀린 건지'를 한 줄로 보이게 한다
        ds = sorted(x["일자"] for x in b.get("완료일미기입목록", []) if x.get("일자"))
        rng = f" — 접수 {ds[0]} ~ {ds[-1]}" if ds else ""
        L.append(f"\n■ 완료일이 안 적힌 오래된 건 {a['완료일미기입']}건 (접수 후 30일 넘음){rng}")
        L.append("   실제로는 끝났을 가능성이 큽니다 — 완료일만 채우면 정리됩니다.")
    if b["내용미기입"]:
        L.append(f"\n■ 작업 내용이 안 적힌 완료건 {len(b['내용미기입'])}건 — 기사에게 확인 필요")
        for x in b["내용미기입"][:5]:
            L.append(f"   {tag(x, d)} · {x['담당기사'] or '기사 미기입'} (완료 처리됨)")
    return "\n".join(L)


def main():
    args = sys.argv[1:]
    day = args[args.index("--date") + 1] if "--date" in args else None
    data, master = load()
    b = brief(day, data)
    print(text(b))
    if "--md" in args:
        os.makedirs(REPORT_DIR, exist_ok=True)
        p = os.path.join(REPORT_DIR, f"일일브리핑_{b['기준일'].replace('-', '')}.md")
        open(p, "w", encoding="utf-8").write(
            f"# 일일 브리핑 {b['기준일']}\n\n원장: {os.path.basename(master)}\n\n```\n{text(b)}\n```\n")
        print("\n리포트:", p)


if __name__ == "__main__":
    main()
