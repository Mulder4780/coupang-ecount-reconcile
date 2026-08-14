# -*- coding: utf-8 -*-
"""
erp_api_parity.py — ERP API 값을 **정본으로 올리기 전에** 내보내기와 대조한다 (읽기 전용)

2026-08-14 형님 지시: **"API를 받아도 우선은 ERP 데이터가 우선이야 검증절차가 있어야돼"**

무엇을 막는가
-------------
OAPI 조회가 열리면 "이제 사람이 안 내려받아도 된다"로 곧장 갈아타기 쉽다. 그런데
이 프로젝트의 ERP 수치는 **돈**이다 — 공급가액·진행상태가 정산·계산서·수금 판정의
근거다. API 응답이 기간·집계·상태 낱말에서 조금만 달라도 화면은 **오류 없이**
다른 숫자를 보여 준다. 비어 있으면 사람이 알아채지만 **틀린 값은 안 띈다.**

그래서 순서를 못 박는다:
  ① 정본은 **사람이 ERP 화면에서 내보낸 엑셀**이다(reports/ERP판매_프로젝트색인.json).
  ② API 로 받은 것은 **별도 자리**에 둔다(reports/ERP판매_API색인.json).
  ③ 이 도구가 둘을 대조해 **관문**을 판정한다.
  ④ 관문을 통과하고 **사람이 승인**하기 전에는 아무도 API 색인을 안 읽는다.

계약 — API 쪽이 지킬 것
-----------------------
API 로 무엇을 어떻게 받든, 결과는 정본과 **똑같은 모양**으로 적는다:

    {"src": [...], "count": N, "as_of": "ISO시각",
     "index": {"UJ2600050": {"supply":433000,"vat":43300,"total":476300,
                             "state":"4.세금계산서발행대기","date":"2026/01/07",
                             "po":"","cust":"강서1MB(가양A)","rows":1}, ...}}

모양이 다르면 대조 자체가 못 선다. 칸 이름은 **영문**이다(state/supply/cust/po) —
한글로 물으면 모든 번호가 '상태 없음'으로 나오는데 오류는 안 난다.

지키는 것
---------
- **읽기 전용.** 색인을 고치지도, 큐에 넣지도, 엑셀을 열지도 않는다.
- **기간이 다른 것을 불일치라 부르지 않는다.** 두 색인의 날짜 범위가 겹치는
  구간 안에서만 비교한다. 안 그러면 경보가 대부분이 되고, 경보가 대부분이면
  아무도 안 본다.
- **겹침이 거의 없으면 A·B 이야기가 아니라 열쇠 이야기다.** 그때는 불일치를
  세지 않고 그렇게 말한다 — 짝이 안 지어진 것을 전부 '값이 다르다'로 내놓으면
  사람이 없는 오류를 찾아 나선다.
- **승인은 그 차이에 대해서만 유효하다.** 승인 지문에 불일치 목록이 들어가므로,
  새 차이가 생기면 승인이 저절로 풀린다. 한 번 승인이 영원한 통과가 되면 안 된다.
- **API 색인이 없으면 '이상 없음'이 아니라 '미실시'다.** 안 본 것을 본 것으로
  세지 않는다.

사용:
    python erp_api_parity.py            # 판정 요약
    python erp_api_parity.py --print    # 본문(불일치 목록)
    python erp_api_parity.py --ack      # 지금 차이를 사람이 설명·승인
"""
import hashlib
import json
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
if hasattr(sys.stdout, "reconfigure"):          # 무인 회차는 stdout 이 None 이다
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

REPORTS = os.path.join(ROOT, "reports")
CANON = os.path.join(REPORTS, "ERP판매_프로젝트색인.json")     # 정본(사람 내보내기)
API = os.path.join(REPORTS, "ERP판매_API색인.json")            # API 가 만든 것
OUT_JSON = os.path.join(REPORTS, "ERP_API대조.json")
OUT_MD = os.path.join(REPORTS, "ERP_API대조.md")
ACK = os.path.join(REPORTS, "ERP_API대조_승인.json")

# 값이 다르면 돈이 달라지는 칸. rows(행수)는 내보내기 방식에 따라 갈릴 수 있어 뺀다.
FIELDS = ("supply", "vat", "total", "state")
# 겹침이 이보다 적으면 열쇠가 어긋난 것으로 본다(값 비교가 뜻을 잃는다).
KEY_MIN_OVERLAP = 0.5


def _load(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
    except FileNotFoundError:
        return None, "없음"
    except Exception as e:
        return None, "못읽음(%s)" % type(e).__name__
    idx = d.get("index")
    if not isinstance(idx, dict):
        return None, "모양이 다름(index 없음)"
    return d, ""


def _day(v):
    """색인의 date 는 '2026/01/07' 모양이다. 비교용으로 자릿수만 맞춘다."""
    s = str(v or "").strip().replace("-", "/")
    return s[:10] if len(s) >= 10 else ""


def _span(idx):
    days = [d for d in (_day(v.get("date")) for v in idx.values()) if d]
    return (min(days), max(days)) if days else ("", "")


def compare(canon_idx, api_idx):
    """두 색인을 대조한다. **기간이 겹치는 구간 안에서만** 값을 비교한다."""
    c_lo, c_hi = _span(canon_idx)
    a_lo, a_hi = _span(api_idx)
    lo, hi = max(c_lo, a_lo), min(c_hi, a_hi)
    overlap_ok = bool(lo and hi and lo <= hi)

    def in_span(rec):
        d = _day(rec.get("date"))
        return bool(d) and lo <= d <= hi

    ck, ak = set(canon_idx), set(api_idx)
    both = ck & ak
    # 기간 밖은 '불일치'가 아니다 — 서로 다른 구간을 받아 온 것뿐이다.
    only_canon = sorted(u for u in (ck - ak) if not overlap_ok or in_span(canon_idx[u]))
    only_api = sorted(u for u in (ak - ck) if not overlap_ok or in_span(api_idx[u]))
    outside = len(ck - ak) - len(only_canon) + len(ak - ck) - len(only_api)

    diffs = []
    for uj in sorted(both):
        c, a = canon_idx[uj], api_idx[uj]
        if overlap_ok and not (in_span(c) or in_span(a)):
            continue
        bad = {f: [c.get(f), a.get(f)] for f in FIELDS
               if str(c.get(f, "")).strip() != str(a.get(f, "")).strip()}
        if bad:
            diffs.append({"프로젝트NO": uj, "칸": bad})

    smaller = min(len(ck), len(ak)) or 1
    return {
        "정본건수": len(ck), "API건수": len(ak), "겹친건수": len(both),
        "겹침비율": round(len(both) / smaller, 3),
        "정본기간": [c_lo, c_hi], "API기간": [a_lo, a_hi],
        "비교구간": [lo, hi] if overlap_ok else [],
        "기간밖제외": outside,
        "정본에만": only_canon, "API에만": only_api, "값다름": diffs,
    }


def fingerprint(cmp_):
    """승인 지문. **차이 자체**를 담는다 — 새 차이가 생기면 승인이 저절로 풀린다."""
    payload = json.dumps({"정본에만": cmp_["정본에만"], "API에만": cmp_["API에만"],
                          "값다름": cmp_["값다름"]}, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def _ack_ok(fp):
    try:
        with open(ACK, "r", encoding="utf-8") as f:
            return json.load(f).get("지문") == fp
    except Exception:
        return False


def judge():
    canon, c_why = _load(CANON)
    api, a_why = _load(API)
    if canon is None:
        return {"관문": "못함", "이유": "정본 색인을 못 읽었습니다(%s). 먼저 "
                                      "erp_sales_index.py 를 돌리세요." % c_why}
    if api is None:
        # ★ '이상 없음'이 아니라 '아직 안 봄'이다. 이 구별이 없으면 API 를 안 받은
        #   상태가 검증을 통과한 상태처럼 보인다.
        return {"관문": "미실시", "이유": "API 색인이 %s. 정본(사람 내보내기)만 씁니다." % a_why,
                "정본건수": len(canon["index"]), "정본원본": canon.get("src", [])}

    cmp_ = compare(canon["index"], api["index"])
    fp = fingerprint(cmp_)
    cmp_["지문"] = fp
    cmp_["API시각"] = api.get("as_of", "")

    if not cmp_["겹친건수"]:
        cmp_.update({"관문": "열쇠확인",
                     "이유": "겹치는 프로젝트NO 가 하나도 없습니다. 값이 다른 것이 아니라 "
                             "서로 다른 것을 보고 있습니다(기간·조회조건·번호 체계)."})
    elif cmp_["겹침비율"] < KEY_MIN_OVERLAP:
        cmp_.update({"관문": "열쇠확인",
                     "이유": "겹침이 %.0f%% 뿐입니다. 값 비교보다 조회조건을 먼저 맞추세요."
                             % (cmp_["겹침비율"] * 100)})
    elif not (cmp_["값다름"] or cmp_["정본에만"] or cmp_["API에만"]):
        cmp_.update({"관문": "통과", "이유": "비교 구간 안에서 차이가 없습니다."})
    elif _ack_ok(fp):
        cmp_.update({"관문": "통과(승인)",
                     "이유": "차이가 있으나 사람이 설명·승인했습니다(지문 %s)." % fp})
    else:
        cmp_.update({"관문": "미통과",
                     "이유": "값다름 %d · 정본에만 %d · API에만 %d. 전부 설명된 뒤에만 "
                             "넘어갑니다." % (len(cmp_["값다름"]), len(cmp_["정본에만"]),
                                             len(cmp_["API에만"]))})
    return cmp_


def render(r):
    L = ["# ERP API ↔ 내보내기 대조", ""]
    L.append("**관문: %s** — %s" % (r.get("관문"), r.get("이유", "")))
    L.append("")
    if r.get("관문") in ("미실시", "못함"):
        L.append("정본 색인 %s건 · 원본 %s개" % (r.get("정본건수", "?"),
                                              len(r.get("정본원본", []) or [])))
        L.append("")
        L.append("> API 값은 아직 어디에도 안 쓰입니다. 이 상태가 정상입니다 — "
                 "API 를 받은 뒤 이 대조를 통과해야 넘어갑니다.")
        return "\n".join(L) + "\n"
    L += ["| | |", "|---|---:|",
          "| 정본 건수 | %d |" % r["정본건수"],
          "| API 건수 | %d |" % r["API건수"],
          "| 겹침 | %d (%.0f%%) |" % (r["겹친건수"], r["겹침비율"] * 100),
          "| 정본 기간 | %s ~ %s |" % tuple(r["정본기간"]),
          "| API 기간 | %s ~ %s |" % tuple(r["API기간"]),
          "| 비교 구간 | %s |" % (" ~ ".join(r["비교구간"]) if r["비교구간"] else "(안 겹침)"),
          "| 기간 밖이라 뺀 건 | %d |" % r["기간밖제외"], ""]
    if r["값다름"]:
        L += ["## 값이 다른 건 (%d)" % len(r["값다름"]), "",
              "| 프로젝트NO | 칸 | 정본 | API |", "|---|---|---:|---:|"]
        for d in r["값다름"][:200]:
            for f, (cv, av) in d["칸"].items():
                L.append("| %s | %s | %s | %s |" % (d["프로젝트NO"], f, cv, av))
        if len(r["값다름"]) > 200:
            L.append("")
            L.append("> 앞 200건만 적었습니다. 전체는 ERP_API대조.json 에 있습니다.")
        L.append("")
    for key, title in (("정본에만", "정본에만 있는 건"), ("API에만", "API 에만 있는 건")):
        if r[key]:
            L += ["## %s (%d)" % (title, len(r[key])), "",
                  ", ".join(r[key][:100]) + (" …" if len(r[key]) > 100 else ""), ""]
    if r.get("관문") == "미통과":
        L += ["---", "",
              "차이를 하나씩 설명한 뒤 `python erp_api_parity.py --ack` 로 승인합니다.",
              "승인은 **지금 이 차이(지문 %s)** 에만 유효하고, 새 차이가 생기면 "
              "저절로 풀립니다." % r.get("지문", "")]
    return "\n".join(L) + "\n"


def main(argv):
    r = judge()
    os.makedirs(REPORTS, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(r, f, ensure_ascii=False, indent=1)
    md = render(r)
    with open(OUT_MD, "w", encoding="utf-8") as f:
        f.write(md)

    if "--ack" in argv:
        if r.get("관문") not in ("미통과",):
            print("승인할 차이가 없습니다 (관문: %s)." % r.get("관문"))
            return 1
        with open(ACK, "w", encoding="utf-8") as f:
            json.dump({"지문": r["지문"], "값다름": len(r["값다름"]),
                       "정본에만": len(r["정본에만"]), "API에만": len(r["API에만"])},
                      f, ensure_ascii=False, indent=1)
        print("승인했습니다. 지문 %s — 새 차이가 생기면 저절로 풀립니다." % r["지문"])
        return 0

    if "--print" in argv:
        print(md)
    else:
        print("관문: %s — %s" % (r.get("관문"), r.get("이유", "")))
        print("→ reports/ERP_API대조.md")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
