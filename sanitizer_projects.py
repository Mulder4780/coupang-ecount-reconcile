"""쿠팡 소독기·세척기 — ERP·공유폴더 자료를 한 곳에 모아 앱에 싣는다 (2026-09-15 형님 지시).

형님 지시: "쿠팡 소독기 세척기 자료 ERP에서 자료 긁어서 카테고리 하나 만들어서 관리할 수
있게 정리해" · "이 폴더 자료 긁어와서 반영해 / 여기가 세척기 소독기 관련 폴더야" ·
"만들고 문제 되는 사항 정리해".

- **읽기만 한다.** 이카운트·공유폴더·관리대장에 한 글자도 안 쓴다. 쓰는 것은
  `reports/소독기세척기_현황.json`(앱이 읽는다)과 `reports/쿠팡_소독기세척기_관리.xlsx`
  (보고용 새 파일) 둘뿐이다.
- **ERP 종류 판정은 새로 만들지 않는다**([162]) — `inbox_scan.classify_rows` 를 빌린다.
- **못 읽은 것을 0 으로 세지 않는다**([169]) — 폴더를 못 열면 `못읽음` 에 이유를 적는다.
- 문제 판정 `find_issues` 는 **순수 함수**다 — 합성검증이 목 자료로 잰다([295]).

실행: python sanitizer_projects.py [--print]
"""
import datetime
import glob
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = (r"\\172.30.1.250\data\2. Cost\★★★쿠팡 업무 폴더★★★"
               r"\♣ 3. 쿠팡 신규 발주 프로젝트\2026.04.24 쿠팡 세척기(소독기)")
OUT_JSON = os.path.join(ROOT, "reports", "소독기세척기_현황.json")
OUT_XLSX = os.path.join(ROOT, "reports", "쿠팡_소독기세척기_관리.xlsx")
CACHE = os.path.join(ROOT, "reports", ".소독기세척기_스캔캐시.json")

# 소독기가 먼저다 — 'FRESH WASHING MACHINE (쿠팡소독기)' 처럼 둘이 섞인 이름이 있다.
STERILIZER = ("소독기", "sterilizer", "freshbox 소독")
WASHER = ("세척기", "washing machine")
ANY = STERILIZER + WASHER + ("freshbox",)
INVOICE_KINDS = {"tax", "taxstep", "taxinv", "hometax"}
STAGE = {"quote": "견적", "taxstep": "세금계산서 진행", "tax": "매출계산서 현황",
         "taxinv": "매출계산서 조회", "hometax": "홈택스 계산서", "stmt": "거래명세서",
         "slips": "회계전표"}


def is_target(text):
    t = (text or "").lower()
    return any(k in t for k in ANY)


def kind_of(text):
    """세척기와 소독기는 **같은 장비**다 (2026-09-15 형님 지시 · 오종현 확인:
    품목코드 55004 'Tote Freshbox 소독기' 하나). 이름이 둘이라 갈라 세면 같은 견적이
    두 번 뜨고 '세척기 전용 품목코드 없음' 같은 없는 문제가 생긴다 — 그래서 한 이름으로 모은다."""
    t = (text or "").lower()
    if any(k in t for k in STERILIZER + WASHER):
        return "소독기"
    return ""


ANSWERS_JSON = os.path.join(ROOT, "reports", "소독기_문제답변.json")


def attach_answers(issues, answers):
    """사람이 준 답을 같은 제목의 문제에 붙인다(순수 함수 · 판정은 안 바꾼다).
    답은 **확인 기록**이지 완료 처리가 아니다 — 문제는 그대로 남고 답이 옆에 붙는다.
    '추가' 답(문제 목록에 없는 확인 사항)은 참고 줄로 싣는다. 못 붙인 답은 돌려준다([169])."""
    out = [dict(i) for i in issues]
    left = []
    for a in answers or []:
        hit = [i for i in out if i["제목"] == a.get("제목")]
        for i in hit:
            i["답변"], i["답변자"], i["답변일"] = a.get("답변", ""), a.get("답변자", ""), a.get("답변일", "")
        if not hit:
            if a.get("추가"):
                out.append({"등급": a.get("등급", "참고"), "제목": a["제목"], "설명": a.get("설명", ""),
                            "근거": a.get("답변자", ""), "답변": a.get("답변", ""),
                            "답변자": a.get("답변자", ""), "답변일": a.get("답변일", "")})
            else:
                left.append(a)
    return out, left


def load_answers(path=ANSWERS_JSON):
    """못 읽으면 빈 목록과 이유(모름을 '답 없음'으로 뭉개지 않는다 · [169])."""
    if not os.path.exists(path):
        return [], ""
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f).get("답변", []), ""
    except (OSError, ValueError) as exc:
        return [], "답변 파일 못 읽음: %s" % type(exc).__name__


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _date(v):
    m = re.search(r"(20\d\d)[./-]?(\d\d)[./-]?(\d\d)", str(v or ""))
    return "%s-%s-%s" % m.groups() if m else ""


def _pick(row, *names):
    for n in names:
        for k, v in row.items():
            if k and n in k and v not in ("", None):
                return v
    return ""


# ── ERP 엑셀 훑기 ───────────────────────────────────────────────
def _load_cache():
    try:
        with open(CACHE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def scan_erp(roots):
    """ERP·쿠팡PO 엑셀에서 소독기·세척기 줄만 뽑는다. 파일 지문이 같으면 캐시를 쓴다([168])."""
    import openpyxl
    import inbox_scan
    cache, fresh, out, fails = _load_cache(), {}, [], []
    for root in roots:
        if not os.path.isdir(root):
            fails.append("폴더 없음: %s" % root)
            continue
        for dp, _dn, fns in os.walk(root):
            for fn in fns:
                if not fn.lower().endswith(".xlsx") or fn.startswith("~$"):
                    continue
                p = os.path.join(dp, fn)
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                key = "%s|%d|%d" % (fn, st.st_size, int(st.st_mtime))
                if key in fresh:          # 같은 파일이 날짜 폴더 여러 곳에 복사돼 있다
                    continue
                if key in cache:
                    fresh[key] = cache[key]
                    out.extend(cache[key])
                    continue
                rows = []
                try:
                    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
                    for ws in wb.worksheets:
                        head, first = None, []
                        for i, r in enumerate(ws.iter_rows(values_only=True)):
                            vals = ["" if v is None else str(v).strip() for v in r]
                            if i < 4:
                                first.append(vals)
                            if head is None and i < 4 and sum(1 for v in vals if v) >= 3:
                                head = vals
                                continue
                            if head and is_target(" ".join(vals)):
                                kind = inbox_scan.classify_rows(first) or ""
                                if "PO목록" in fn:
                                    kind = "po"
                                rows.append({"파일": fn, "종류": kind,
                                             "칸": dict(zip(head, vals))})
                    wb.close()
                except Exception as exc:  # 한 파일이 깨져도 나머지는 읽는다 — 이유는 남긴다
                    fails.append("%s: %s" % (fn, type(exc).__name__))
                fresh[key] = rows
                out.extend(rows)
    try:
        tmp = CACHE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(fresh, f, ensure_ascii=False)
        os.replace(tmp, CACHE)
    except OSError:
        pass
    return out, fails


def normalize(raw):
    """훑은 줄을 같은 모양의 거래 기록으로 맞추고, 여러 파일에 겹친 같은 줄을 하나로 줄인다."""
    recs, seen = [], set()
    for r in raw:
        c, kind = r["칸"], r["종류"]
        text = " ".join(str(v) for v in c.values())
        item = _pick(c, "품목명", "적요명", "프로젝트명") or ""
        spec = _pick(c, "규격")
        if spec and spec not in item:
            item = "%s [%s]" % (item, spec)
        stage = STAGE.get(kind, "")
        if kind == "po":
            stage = "쿠팡 PO"
        elif not stage:
            if _pick(c, "PO번호") and _pick(c, "금액합계", "판매금액합계"):
                stage = "판매"
            elif _pick(c, "발주일"):
                stage = "프로젝트 현황"
            else:
                stage = "전표(종류 미확인)"
        vals = list(c.values())
        date = _date(_pick(c, "일자", "발주일")) or next((_date(v) for v in vals if _date(v)), "")
        supply = _num(_pick(c, "공급가액", "견적공급가액합계", "판매공급가액합계", "공급가액합계"))
        if supply is None and kind == "po":
            supply = next((_num(v) for v in vals if (_num(v) or 0) > 100000), None)
        total = _num(_pick(c, "합계금액", "견적금액합계", "판매금액합계", "금액합계", "합계"))
        rec = {
            "일자": date, "단계": stage, "종류": kind, "구분": kind_of(text),
            "프로젝트": _pick(c, "프로젝트코드"), "프로젝트명": _pick(c, "프로젝트명"),
            "품목코드": _pick(c, "품목코드"), "품목": item,
            "수량": _num(_pick(c, "수량")), "단가": _num(_pick(c, "단가", "출고단가")),
            "공급가": supply, "합계": total,
            "거래처": _pick(c, "거래처명", "구매처명") or ("Coupang" if kind == "po" else ""),
            "진행": _pick(c, "진행상태", "전자(세금)계산서 진행단계"),
            "PO": _pick(c, "PO번호") or (vals[0] if kind == "po" else ""),
            "파일": r["파일"],
        }
        if not rec["구분"]:
            continue
        key = (rec["일자"], rec["단계"], rec["품목"], rec["거래처"], rec["공급가"])
        if key in seen:
            continue
        seen.add(key)
        recs.append(rec)
    recs.sort(key=lambda x: (x["일자"], x["단계"]))
    return recs


def item_master(path):
    try:
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError) as exc:
        return None, "품목 목록을 못 읽음: %s" % type(exc).__name__
    rows = d.get("rows") if isinstance(d, dict) else d
    out = []
    for i in rows or []:
        blob = json.dumps(i, ensure_ascii=False)
        if is_target(blob) or "세척솔" in blob:
            if "세척솔" in blob:
                continue  # 청소용 솔은 장비가 아니다
            out.append({"품목코드": i.get("PROD_CD", ""), "품목명": i.get("PROD_DES", ""),
                        "규격": i.get("SIZE_DES", ""), "판매단가": _num(i.get("OUT_PRICE")),
                        "적요": i.get("REMARKS_WIN", "")})
    return out, ""


# ── 공유폴더 ────────────────────────────────────────────────────
def _file_kind(rel):
    n = os.path.basename(rel)
    if "견적" in n:
        return "견적"
    if "계약" in rel or "게약" in n:   # 실측: '물품공급게약서.pdf' 오타 파일이 있다
        return "계약"
    if n.lower().endswith(".dwg") or "도면" in n:
        return "도면"
    if "매뉴얼" in rel:
        return "매뉴얼·카탈로그"
    return "제안·기타"


def parse_quote(path):
    """견적서 엑셀의 첫 장(견적서)에서 번호·품명·수량·단가·금액·특기사항을 읽는다."""
    import openpyxl
    q = {"파일": os.path.basename(path), "날짜": _date(os.path.basename(path)),
         "견적번호": "", "품명": "", "수량": None, "단가": None, "금액": None,
         "유효기간일": None, "특기_안전관리비": None, "내역_안전관리비": None}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for si, ws in enumerate(wb.worksheets):
            if si > 1:
                break
            for r in ws.iter_rows(values_only=True):
                vals = [str(v).strip() for v in r if v not in (None, "")]
                if not vals:
                    continue
                line = " ".join(vals)
                if si == 0:
                    if vals[0].startswith("견적번호") and len(vals) > 1:
                        q["견적번호"] = vals[1]
                    elif vals[0].replace(" ", "") == "1" and len(vals) >= 5 and q["단가"] is None:
                        nums = [_num(v) for v in vals[1:]]
                        q["품명"] = vals[1]
                        if len(nums) >= 3 and nums[-1] and nums[-2]:
                            q["금액"], q["단가"], q["수량"] = nums[-1], nums[-2], nums[-3]
                    m = re.search(r"견\s*적\s*유\s*효\s*기\s*간\D*(\d+)", line)
                    if m:
                        q["유효기간일"] = int(m.group(1))
                    m = re.search(r"안전관리비\s*([\d,]+)\s*원", line)
                    if m:
                        q["특기_안전관리비"] = _num(m.group(1))
                elif "표준안전관리비" in line and len(vals) >= 2:
                    q["내역_안전관리비"] = _num(vals[-1])
    finally:
        wb.close()
    q["구분"] = kind_of(q["파일"] + " " + q["품명"])
    return q


def scan_folder(folder):
    if not os.path.isdir(folder):
        return [], [], "폴더를 못 열었습니다: %s" % folder
    files, quotes = [], []
    for dp, _dn, fns in os.walk(folder):
        for fn in fns:
            if fn.lower() == "thumbs.db":
                continue
            p = os.path.join(dp, fn)
            rel = os.path.relpath(p, folder)
            try:
                st = os.stat(p)
            except OSError:
                continue
            files.append({"분류": _file_kind(rel), "경로": rel,
                          "수정일": datetime.date.fromtimestamp(st.st_mtime).isoformat(),
                          "KB": round(st.st_size / 1024)})
            if fn.lower().endswith(".xlsx") and "견적서" in fn and os.path.dirname(rel) == "":
                try:
                    quotes.append(parse_quote(p))
                except Exception as exc:
                    quotes.append({"파일": fn, "못읽음": type(exc).__name__})
    files.sort(key=lambda x: (x["분류"], x["경로"]))
    quotes.sort(key=lambda x: x.get("날짜") or "")
    return files, quotes, ""


# ── 문제 판정 (순수 함수) ───────────────────────────────────────
def _won(v):
    return "{:,.0f}원".format(v) if isinstance(v, (int, float)) else "-"


def find_issues(recs, items, quotes, files):
    """거래 기록·품목·견적·폴더 목록에서 사람이 확인할 것을 뽑는다. 아무것도 안 읽는다."""
    issues = []

    def add(grade, title, detail, basis):
        # 같은 줄이 여러 파일에 복사돼 있어 같은 문제가 두 번 뜬다 — 한 번만 싣는다([170])
        if any(i["제목"] == title and i["설명"] == detail for i in issues):
            return
        issues.append({"등급": grade, "제목": title, "설명": detail, "근거": basis})

    quotes = [q for q in quotes if not q.get("못읽음")]
    kinds = sorted({r["구분"] for r in recs if r.get("구분")})
    for k in kinds:
        mine = [r for r in recs if r["구분"] == k]
        inv = [r for r in mine if r["종류"] in INVOICE_KINDS or "세금계산서발행" in str(r.get("진행"))]
        last_inv = max((r["일자"] for r in inv), default="")
        orders = [r for r in mine if r["단계"] == "전표(종류 미확인)" and r["일자"] > last_inv]
        kq = [q for q in quotes if q.get("구분") == k and q.get("단가")]
        latest_q = kq[-1] if kq else None

        # ① ERP 주문 단가 ≠ 최종 견적 단가
        if latest_q:
            for o in orders:
                if o.get("단가") and abs(o["단가"] - latest_q["단가"]) >= 1:
                    add("높음", "%s ERP 단가가 최종 견적과 다름" % k,
                        "ERP %s(%s) · 견적 %s(%s) · 차이 %s/대" % (
                            _won(o["단가"]), o["일자"], _won(latest_q["단가"]),
                            latest_q["날짜"], _won(o["단가"] - latest_q["단가"])),
                        "%s · %s" % (o["파일"], latest_q["파일"]))
                    break
        # ② 주문은 있는데 그 뒤 세금계산서가 없음
        if orders:
            add("높음", "%s 주문 뒤 세금계산서 확인 안 됨" % k,
                "%d건 %s(최근 %s) — 받아 둔 ERP 자료에 발행 기록이 없다(발행 안 됐다는 뜻은 아님)" % (
                    len(orders), _won(sum(o.get("공급가") or 0 for o in orders)),
                    max(o["일자"] for o in orders)),
                ", ".join(sorted({o["파일"] for o in orders})))
        # ③ 프로젝트코드 없는 주문
        noproj = [o for o in orders if not o.get("프로젝트")]
        if noproj:
            add("보통", "%s 주문에 프로젝트코드 없음" % k,
                "%d건 — 원장·대조에서 이 건을 프로젝트로 못 묶는다" % len(noproj),
                ", ".join(sorted({o["파일"] for o in noproj})))
        # ④ 쿠팡 PO 없음
        if orders and not any(r["종류"] == "po" and r["일자"] >= (latest_q or {}).get("날짜", "") for r in mine):
            add("보통", "%s 쿠팡 PO 확인 안 됨" % k,
                "주문(%s) 에 맞는 쿠팡 PO 가 받아 둔 PO 목록에 없다" % max(o["일자"] for o in orders),
                "쿠팡PO목록")
        # ⑤ 견적 유효기간 지난 뒤 주문
        if latest_q and latest_q.get("유효기간일") and latest_q.get("날짜") and orders:
            end = (datetime.date.fromisoformat(latest_q["날짜"])
                   + datetime.timedelta(days=latest_q["유효기간일"])).isoformat()
            first = min(o["일자"] for o in orders)
            if first > end:
                add("보통", "%s 견적 유효기간이 지난 뒤 주문 등록" % k,
                    "견적 %s + %d일 = %s 까지인데 주문은 %s" % (
                        latest_q["날짜"], latest_q["유효기간일"], end, first), latest_q["파일"])
        # ⑥ 전용 품목코드 없음
        if items is not None and not any(kind_of(i["품목명"]) == k for i in items):
            codes = sorted({r["품목코드"] for r in mine if r.get("품목코드")})
            add("보통", "%s 전용 품목코드 없음" % k,
                "이카운트 품목에 %s 이름이 없고 범용 품목(%s)에 규격으로만 붙어 있다" % (
                    k, ", ".join(codes) or "-"), "이카운트 품목 목록")

    # ⑦ 같은 날·같은 금액인데 이름이 다른 견적
    qs = [r for r in recs if r["종류"] == "quote" and r.get("합계")]
    for i, a in enumerate(qs):
        for b in qs[i + 1:]:
            if a["일자"] == b["일자"] and a["합계"] == b["합계"] and a["구분"] != b["구분"]:
                add("보통", "같은 견적이 세척기·소독기 두 이름으로 등록",
                    "%s %s — '%s' 와 '%s'" % (a["일자"], _won(a["합계"]),
                                             a["프로젝트명"] or a["품목"], b["프로젝트명"] or b["품목"]),
                    "%s · %s" % (a["파일"], b["파일"]))
    # ⑧ 견적서 이름과 품명이 서로 다른 장비를 가리킴
    for q in quotes:
        if kind_of(q["파일"]) and kind_of(q["품명"]) and kind_of(q["파일"]) != kind_of(q["품명"]):
            add("참고", "견적서 파일명과 품명이 다른 장비 이름",
                "파일은 %s, 품명은 '%s'" % (kind_of(q["파일"]), q["품명"]), q["파일"])
    # ⑨ 특기사항 안전관리비 ≠ 내역 안전관리비
    for q in quotes:
        a, b = q.get("특기_안전관리비"), q.get("내역_안전관리비")
        if a and b and abs(a - b) >= 1:
            add("높음", "견적서 안전관리비 금액이 앞뒤로 다름",
                "특기사항 %s · 내역 %s" % (_won(a), _won(b)), q["파일"])
    # ⑩ 같은 견적번호·금액 파일 여러 개
    groups = {}
    for q in quotes:
        if q.get("견적번호") and q.get("금액"):
            groups.setdefault((q["견적번호"], q["금액"]), []).append(q["파일"])
    for (no, amt), fs in groups.items():
        if len(fs) > 1:
            add("참고", "같은 견적서가 여러 파일로 있음",
                "견적번호 %s · %s · %d개 — 어느 것이 최종인지 이름으로 안 갈린다" % (no, _won(amt), len(fs)),
                " · ".join(fs))
    # ⑪ 계약서가 우리 날인본뿐
    contracts = [f["경로"] for f in files if f["분류"] == "계약"]
    if contracts and not any(("쿠팡 날인" in c or "양사" in c or "최종" in c) for c in contracts):
        add("보통", "계약서 쿠팡 날인본 확인 안 됨",
            "폴더에 있는 계약서: %s" % " · ".join(contracts), "공유폴더")
    order = {"높음": 0, "보통": 1, "참고": 2}
    issues.sort(key=lambda x: order.get(x["등급"], 9))
    return issues


def summarize(recs):
    out = []
    for k in sorted({r["구분"] for r in recs}):
        mine = [r for r in recs if r["구분"] == k]
        inv = {}
        for r in mine:
            if (r["종류"] in INVOICE_KINDS or "세금계산서발행" in str(r.get("진행"))) and r.get("공급가"):
                inv[(r["일자"], r["공급가"])] = r
        last_inv = max((d for d, _ in inv), default="")
        orders = [r for r in mine if r["단계"] == "전표(종류 미확인)" and r["일자"] > last_inv]
        out.append({"구분": k,
                    "발행완료_공급가": sum(v for _, v in inv),
                    "발행완료_건": len(inv),
                    "주문등록_공급가": sum(o.get("공급가") or 0 for o in orders),
                    "주문등록_대수": sum(o.get("수량") or 0 for o in orders),
                    "최근": max((r["일자"] for r in mine), default="")})
    return out


# ── 저장 ────────────────────────────────────────────────────────
def write_xlsx(data, path=OUT_XLSX):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    head_font, head_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")

    def sheet(title, cols, rows):
        ws = wb.create_sheet(title)
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = head_font, head_fill
        for r in rows:
            ws.append([r.get(c, "") if isinstance(r, dict) else r for c in cols])
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, (int, float)) and abs(c.value) >= 1000:
                    c.number_format = "#,##0"
        ws.freeze_panes = "A2"
        return ws

    wb.remove(wb.active)
    sheet("1_요약", ["구분", "발행완료_건", "발행완료_공급가", "주문등록_대수", "주문등록_공급가", "최근"],
          data["요약"])
    sheet("2_문제사항", ["등급", "제목", "설명", "답변", "답변자", "답변일", "근거"], data["문제"])
    sheet("3_거래내역", ["일자", "단계", "구분", "프로젝트", "프로젝트명", "품목코드", "품목", "수량",
                     "단가", "공급가", "합계", "거래처", "진행", "PO", "파일"], data["거래"])
    sheet("4_견적이력", ["날짜", "구분", "견적번호", "품명", "수량", "단가", "금액", "유효기간일", "파일"],
          data["견적"])
    sheet("5_품목등록", ["품목코드", "품목명", "규격", "판매단가", "적요"], data["품목"] or [])
    ws = sheet("6_폴더자료", ["분류", "경로", "수정일", "KB"], data["폴더자료"])
    ws.append([])
    ws.append(["원본 폴더", PROJECT_DIR])
    tmp = path + ".tmp.xlsx"
    wb.save(tmp)
    os.replace(tmp, path)


def build():
    import source_dirs as sd
    raw, fails = scan_erp([sd.ERP_DIR, sd.COUPANG_DIR])
    recs = normalize(raw)
    items, item_err = item_master(os.path.join(ROOT, "inbox", "api", "ecount_items_latest.json"))
    files, quotes, folder_err = scan_folder(PROJECT_DIR)
    answers, ans_err = load_answers()
    issues, unmatched = attach_answers(find_issues(recs, items, quotes, files), answers)
    data = {
        "만든시각": datetime.datetime.now().isoformat(timespec="seconds"),
        "원본폴더": PROJECT_DIR,
        "요약": summarize(recs),
        "문제": issues,
        "답변_지난문제": unmatched,
        "거래": recs, "견적": quotes, "품목": items or [], "폴더자료": files,
        "못읽음": [m for m in [item_err, folder_err, ans_err] + fails[:20] if m],
    }
    tmp = OUT_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, OUT_JSON)
    try:
        write_xlsx(data)
    except Exception as exc:  # 보고용 엑셀이 열려 있어도 앱 자료는 이미 저장됐다 — 이유만 남긴다
        data["못읽음"].append("엑셀 저장 못함: %s" % type(exc).__name__)
    return data


if __name__ == "__main__":
    d = build()
    print("거래 %d · 견적 %d · 폴더 파일 %d · 문제 %d · 못읽음 %d" % (
        len(d["거래"]), len(d["견적"]), len(d["폴더자료"]), len(d["문제"]), len(d["못읽음"])))
    if "--print" in sys.argv:
        for s in d["요약"]:
            print("  %s" % s)
        for i in d["문제"]:
            print("  [%s] %s — %s" % (i["등급"], i["제목"], i["설명"]))
        for m in d["못읽음"]:
            print("  못읽음: %s" % m)
    sys.exit(0)
