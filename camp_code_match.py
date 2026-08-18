# -*- coding: utf-8 -*-
"""캠프명 ↔ ERP 거래처코드 대조 — 유니웍스 기초등록용 (2026-08-18 노승용 매니저 요청).

요청 원문(업무협조): ① ERP 거래처코드(CU001)와 각자 엑셀의 캠프명이 **100% 일치**하는
것을 우선 유니웍스에 등록 ② 불일치·없는 것 찾아 등록·수정 — `US직구 허브넷` 인데
실제 캠프명은 `인천7캠프(허브넷)` 인 것, `00016 시흥1캠프(오류동)` 처럼 쿠팡 거래처코드
형식이 아닌 것, ERP 거래처코드로 검색이 안 되는 캠프.

★ **찾아 주기만 한다. 아무것도 안 고친다.** ERP·유니웍스·원장·큐 어디에도 안 쓴다.
  무엇이 맞는지는 사람만 안다 — 자동으로 짝지으면 **틀린 거래처코드가 유니웍스에
  박히고**, 그건 빈 칸보다 나쁘다(`[172]`). `typo_watch` 와 같은 자리다.

★ **원천마다 캠프 수가 다르다** — 그래서 "미확인 캠프"의 답은 기준에 따라 달라진다.
  실측 2026-08-18: ERP 거래처 마스터 268 · 밴드 접수 글 737 · 원장 색인 198.
  세 원천을 다 모으고 **어디서 온 캠프인지**를 행마다 적는다. 한 원천만 보면
  다른 원천에만 있는 캠프가 '없는 캠프'가 된다(`[169]`).

★ **판정을 새로 만들지 않는다**(`[162]`) — 캠프명 맞추기는 `camp_contacts._norm`,
  거래처코드 짝짓기는 `거래처코드_색인.json`(회차가 이미 만든 것)을 그대로 읽는다.
  여기서 다시 짝지으면 화면과 이 표가 서로 다른 답을 한다.

사람: `python camp_code_match.py`  ·  결과: `reports/캠프_거래처코드_대조.{md,xlsx}`
"""
import json
import os
import re
import sys
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):          # 무인 회차는 sys.stdout 이 None 이다(`[235]`)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT = os.path.join(ROOT, "reports")
OUT_MD = os.path.join(REPORT, "캠프_거래처코드_대조.md")
OUT_XLSX = os.path.join(REPORT, "캠프_거래처코드_대조.xlsx")

#: 쿠팡 거래처코드 정본 형식. 노승용 요청의 `00016 시흥1캠프(오류동)` 이 여기 안 맞는다.
CODE_OK = re.compile(r"^CU\d{3,4}$", re.I)

#: 갈래 — 순서가 곧 사람이 처리할 순서다(요청의 1번·2번 그대로).
KINDS = [
    ("A", "바로 등록 가능", "거래처코드가 있고 ERP 거래처명이 캠프명과 같다"),
    ("B", "이름 다름", "코드는 붙었는데 ERP 거래처명이 캠프명과 다르다 — 어느 쪽이 맞는지 사람이 정한다"),
    ("C", "코드 형식 다름", "쿠팡 거래처코드(CU+숫자) 형식이 아니다 — 유니웍스·ERP 둘 다 고쳐야 한다"),
    ("D", "후보 여럿", "거래처코드 후보가 둘 이상이다 — 자동으로 못 고른다"),
    ("E", "이름으로 못 찾음", "밴드에서만 보이는 캠프다 — ERP 마스터에도, ERP 거래처 원본"
     "(ESA001M) 에도 **댈 후보가 없다**. 없는 캠프인지 아직 못 본 것인지는 사람이 확인한다"),
    ("F", "코드만 비었음", "ERP·원장에도 있는 캠프인데 거래처코드 칸이 비었다 — 코드를 채우면 된다"),
    ("G", "이름이 다른 것으로 보임 ★요청 2번",
     "ERP 거래처 원본에서 **후보가 정확히 하나** 나왔다 — `US직구 허브넷` ↔ `인천7캠프(허브넷)` "
     "가 그 모양이다. **지목일 뿐 확정이 아니다**(아무것도 안 고친다) — 근거를 보고 사람이 정한다"),
    ("H", "후보 여럿 — 사람이 고른다",
     "후보가 둘 이상이라 자동으로 못 고른다. 하나를 짐작해 고르면 틀린 코드가 유니웍스에 "
     "박힌다 — **빈 칸보다 나쁘다.** 후보를 나란히 보여 준다"),
]

#: ★ **이 표로는 못 재는 것**(`[169]` — 못 잰 것을 '0건'이라 적지 않는다).
#:   요청 2번의 `US직구 허브넷` ↔ `인천7캠프(허브넷)` 같은 **이름 불일치**는 여기서
#:   B 로 안 걸린다. 근거인 `캠프마스터.json` 이 **`캠프명 == ERP거래처명` 인 것만**
#:   담기 때문이다 — 이름이 다르면 애초에 그 표에 안 들어온다. 그래서 B 가 구조상
#:   0 이고, 그 캠프들은 **E 로 떨어진다.** 제대로 재려면 ERP 거래처 원본
#:   (`ESA001M.xlsx` 전체 2,981건)을 주소·부분이름으로 대야 한다 — 분담판에 남겼다.
CANNOT_MEASURE = (
    "이름 불일치(요청 2번)는 **B 로 안 걸립니다.** 근거인 캠프마스터가 "
    "`캠프명 == ERP거래처명` 인 것만 담아, 이름이 다른 캠프는 애초에 그 표에 안 들어옵니다. "
    "그래서 ERP 거래처 원본(ESA001M) 전체를 따로 읽어 **G·H** 로 답합니다 — "
    "B 가 0 인 것은 '이름 다른 캠프가 없다'는 뜻이 아닙니다."
)

#: 지목의 근거 — 어느 것도 '확정'이 아니다. 사람이 보고 정한다.
#: ★ **잡는 것보다 잘못 지목하지 않는 것이 어렵다**(`[172]`). 실측으로 걸러 낸 오탐:
#:   · `용인1MB(구갈동)` → `용인2MB(구갈동)` — 괄호 지명이 같고 **앞 숫자만 다르다**.
#:     그 둘은 실재하는 **다른 캠프**다(`송파1MB`·`송파5MB` 와 같은 자리).
#:   · `MC12(삼척)` → `삼척파랑지게차` · `M_SNC1 (순천)` → `팔마중기(순천시지게차)` —
#:     괄호 안이 **도시 이름**이면 그 도시의 남의 회사가 전부 걸린다.
#:   · `남양주1MB(창동)` → `SR-ICH(**원**창동)` — 지명이 **남의 이름 안에 박힌 것**.
#: 그래서 문이 넷이다: 후보 풀을 캠프처럼 생긴 거래처로 좁히고 · 번호가 다르면 거절 ·
#: 지역이 다르면 거절 · 괄호 지명이 서로 어긋나면 거절. 그러고도 **유일할 때만** 지목한다.
CAMPISH = re.compile(r"캠프|모바일|미니|MB|Sub|허브|HUB|FC|OFC|직구", re.I)
TYPEWORD = re.compile(r"(SUBHUB|SUBFC|SUB|HUB|OFC|CAMP|MB|FC|캠프|모바일|미니|서브|허브|센터|물류|지점)", re.I)
MEMO = re.compile(r"[\[\uff3b\u3010][^\]\uff3d\u3011]*[\]\uff3d\u3011]")
PAREN = re.compile(r"[(\uff08]([^)\uff09]+)[)\uff09]")
ROAD = re.compile(r"([가-힣A-Za-z0-9]+(?:로|길))\s*([0-9]+(?:-[0-9]+)?)")
#: 괄호 안이 이것뿐이면 자리를 가리키는 말이 아니다 — 근거로 쓰지 않는다.
PAREN_GENERIC = {"주", "구", "동", "시", "군", "읍", "면", "리", "본사", "신규", "미정", "임시"}
ERP_CACHE = os.path.join(REPORT, ".거래처등록_캐시.json")


def _norm(name):
    """캠프명 맞추기 — `camp_contacts` 것을 **빌린다**(`[162]`). 못 부르면 같은 규칙을 쓴다."""
    try:
        import camp_contacts
        return camp_contacts._norm(name)
    except Exception:                            # noqa: BLE001
        return re.sub(r"[\s()\[\]_\-/·.]", "", str(name or "")).upper()


def _load(path):
    """→ `(내용, 못읽은이유)`. **못 읽음은 '없음'이 아니다**(`[169]`)."""
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh), ""
    except FileNotFoundError:
        return None, "파일이 없다(회차가 아직 안 돌았을 수 있다)"
    except Exception as exc:                     # noqa: BLE001
        return None, "%s: %s" % (type(exc).__name__, str(exc)[:80])


def _looks_like_camp(name):
    """캠프 이름 모양인가 — 파싱이 깨진 조각을 표에 싣지 않는다.

    실측: `거래처코드_색인.json` 의 `ambiguous` 에 `)` 한 글자짜리 키가 있었고 그 후보가
    `(주)대한민국문화재단` 같은 **남의 회사 2,981건**이었다. 그대로 실으면 표 맨 위가
    쓰레기로 덮인다. 다만 **조용히 빼지 않는다** — 몇 건을 뺐는지 리포트가 적는다(`[169]`).
    """
    s = str(name or "").strip()
    return len(s) >= 3 and _norm(s) != ""


def _erp_customers():
    """ERP 거래처 원본(ESA001M) — **캐시 검사가 먼저다**(`[168]`).

    `customer_index.load_customers()` 는 Z:(SMB) 를 재귀로 훑는다. 그 색인 단계가
    같은 회차에서 이미 돌아 캐시를 갱신하므로, 여기서는 **캐시를 그냥 읽는다**.
    캐시가 아예 없을 때만 원본을 읽고, **그 사실과 나이를 리포트가 적는다**(`[169]`) —
    낡은 근거로 '없다'고 말하는 것이 이 프로젝트에서 제일 위험하다.
    """
    try:
        c = json.load(open(ERP_CACHE, encoding="utf-8"))
        rows = c.get("rows") or []
        if rows:
            return rows, {"출처": c.get("src") or "ESA001M", "받은때": c.get("at") or "",
                          "길": "캐시", "건수": len(rows)}
    except (OSError, ValueError):
        pass
    try:
        from customer_index import load_customers
        rows, src = load_customers()
        return rows, {"출처": src or "ESA001M", "받은때": "", "길": "원본 다시 읽음",
                      "건수": len(rows)}
    except Exception as e:                       # noqa: BLE001
        return [], {"출처": "", "받은때": "", "길": "못 읽음: %s" % e, "건수": 0}


def _head(name):
    """괄호·대괄호 메모를 떼고 남은 앞부분 — `송파3Sub-FC [= …]` 의 그 메모다."""
    t = MEMO.sub(" ", str(name or ""))
    return re.sub(r"[\s\-_\u00b7,]", "", t.split("(")[0].split("\uff08")[0]).strip()


def _desig(name):
    """캠프 번호 지문 — 종류말(캠프·MB·모바일…)을 뗀 나머지의 숫자·낱글자."""
    return "".join(re.findall(r"[0-9]+|(?<![A-Z가-힣])[A-Z](?![A-Z])",
                              TYPEWORD.sub("", _head(name).upper())))


def _region(name):
    return "".join(re.findall(r"[가-힣]+", TYPEWORD.sub("", _head(name).upper())))


def _parens(name):
    out = []
    for m in PAREN.finditer(str(name or "")):
        t = re.sub(r"[\s\-_]", "", m.group(1).strip())
        if len(t) >= 2 and not t.isdigit() and t not in PAREN_GENERIC:
            out.append(t)
    return out


def _road_key(addr):
    m = ROAD.search(str(addr or ""))
    return (m.group(1) + m.group(2)) if m else ""


def _paren_clash(a, b):
    """괄호 지명이 양쪽 다 있는데 두 글자도 안 겹치면 **다른 자리**다."""
    pa, pb = _parens(a), _parens(b)
    if not pa or not pb:
        return False
    return not any(x[:2] == y[:2] or x in y or y in x for x in pa for y in pb)


def _reject(camp, erp_name):
    """거절 사유 — 없으면 None. 판정을 한 곳에 모아 둔다(`[162]`)."""
    da, db = _desig(camp), _desig(erp_name)
    if da and db and da != db:
        return "번호가 다르다 (%s ↔ %s)" % (da, db)
    # ★ 지역 판정은 **양쪽에 번호가 있을 때만** 건다. 번호 없는 이름(`US직구 허브넷`)은
    #   지역 추출 자체가 헛돈다 — 종류말 `허브` 를 뗀 `직구넷` 이 '지역'으로 잡혀
    #   **요청 2번의 본보기가 통째로 거절됐다**(만들면서 검증이 잡았다).
    #   남의 회사는 이미 후보 풀(캠프처럼 생긴 거래처)에서 걸러진다.
    ra, rb = _region(camp), _region(erp_name)
    if da and db and ra and rb and ra != rb and not (ra in rb or rb in ra):
        return "지역이 다르다 (%s ↔ %s)" % (ra, rb)
    if _paren_clash(camp, erp_name):
        return "괄호 지명이 다르다"
    return None


def guess_codes(camps, custs):
    """이름이 다른 캠프의 거래처코드 후보를 찾는다 — **아무것도 안 고친다**.

    근거 셋: 괄호 안 지명 · 도로명 주소 · 부분 이름. 셋 다 '있으면 후보'일 뿐이고
    `_reject` 를 통과한 것만 남는다. 돌려주는 것은 `{정규화이름: [후보…]}` 이다.
    """
    pool = [c for c in custs
            if str(c.get("code") or "").upper().startswith("CU")
            or CAMPISH.search(str(c.get("name") or ""))]
    by_road = {}
    for c in pool:
        k = _road_key(c.get("addr"))
        if k:
            by_road.setdefault(k, []).append(c)
    out = {}
    for row in camps:
        cam, cands, why = row["캠프명"], {}, {}

        def add(c, w):
            cands[str(c.get("code") or "")] = c
            why.setdefault(str(c.get("code") or ""), set()).add(w)

        for pz in _parens(cam):
            # ★ 앞 경계를 본다 — `창동` 이 `원창동` 안에 박힌 것을 짝이라 하지 않는다.
            pat = re.compile(r"(?<![가-힣A-Za-z])" + re.escape(pz))
            for c in pool:
                if pat.search(str(c.get("name") or "")):
                    add(c, "괄호 지명 `%s`" % pz)
        k = _road_key(row.get("주소"))
        if k:
            for c in by_road.get(k, []):
                add(c, "주소 `%s`" % k)
        h = _head(cam)
        if len(h) >= 4:
            for c in pool:
                n = re.sub(r"[\s\-_]", "", str(c.get("name") or ""))
                if h in n or (len(n) >= 4 and n in h):
                    add(c, "이름 일부 `%s`" % h)
        keep = []
        for code, c in cands.items():
            if not code or _reject(cam, str(c.get("name") or "")):
                continue
            keep.append({"code": code, "name": str(c.get("name") or ""),
                         "addr": str(c.get("addr") or "")[:44],
                         "근거": " · ".join(sorted(why[code]))})
        keep.sort(key=lambda x: (-len(x["근거"].split(" · ")), x["code"]))
        if keep:
            out[_norm(cam)] = keep
    return out


def build():
    master, e1 = _load(os.path.join(REPORT, "캠프마스터.json"))
    contacts, e2 = _load(os.path.join(REPORT, "캠프_담당자.json"))
    index, e3 = _load(os.path.join(REPORT, "거래처코드_색인.json"))
    못읽음 = [x for x in (("캠프마스터", e1), ("캠프_담당자", e2), ("거래처코드_색인", e3)) if x[1]]

    camps = {}                                   # 정규화이름 -> 행

    def slot(name, 출처):
        key = _norm(name)
        row = camps.setdefault(key, {
            "캠프명": name, "다른표기": set(), "거래처코드": "", "ERP거래처명": "",
            "정기점검": False, "건수": 0, "출처": set(), "후보": [], "주소": "",
        })
        if len(str(name)) < len(str(row["캠프명"])):
            row["다른표기"].add(row["캠프명"])   # 짧은 쪽을 대표로 — 긴 것은 메모가 붙은 표기다
            row["캠프명"] = name
        elif name != row["캠프명"]:
            row["다른표기"].add(name)
        row["출처"].add(출처)
        return row

    # ① ERP 거래처 마스터(ESA001M) — 코드의 정본
    for r in ((master or {}).get("rows") or []):
        if not _looks_like_camp(r.get("캠프명")):
            continue
        row = slot(r.get("캠프명"), "ERP")
        row["거래처코드"] = row["거래처코드"] or str(r.get("거래처코드") or "")
        row["ERP거래처명"] = row["ERP거래처명"] or str(r.get("ERP거래처명") or "")
        row["주소"] = row["주소"] or str(r.get("주소") or "")
        for a in str(r.get("별칭") or "").split(","):
            if a.strip():
                row["다른표기"].add(a.strip())
        row["건수"] = max(row["건수"], int(r.get("업무건수") or 0))

    # ② 밴드 접수 글 — 현장이 실제로 부르는 이름(가장 많다)
    for r in ((contacts or {}).get("rows") or []):
        if not _looks_like_camp(r.get("캠프명")):
            continue
        row = slot(r.get("캠프명"), "밴드")
        row["거래처코드"] = row["거래처코드"] or str(r.get("거래처코드") or "")
        row["주소"] = row["주소"] or str(r.get("캠프주소") or "")
        row["정기점검"] = row["정기점검"] or bool(r.get("정기점검"))
        row["건수"] = max(row["건수"], int(r.get("총건수") or 0))
        for a in (r.get("다른표기") or []):
            row["다른표기"].add(a)

    # ③ 원장 색인 — 이미 지어진 짝과 **못 지어진 짝**
    idx = index or {}
    for name, v in (idx.get("linked") or {}).items():
        if not _looks_like_camp(name):
            continue
        row = slot(name, "원장")
        row["거래처코드"] = row["거래처코드"] or str(v.get("code") or "")
        row["ERP거래처명"] = row["ERP거래처명"] or str(v.get("erp_name") or "")
        row["주소"] = row["주소"] or str(v.get("addr") or "")
    버린쓰레기 = 0
    for name, v in (idx.get("ambiguous") or {}).items():
        if not _looks_like_camp(name):
            버린쓰레기 += 1
            continue
        row = slot(name, "원장")
        row["후보"] = [{"code": c, "name": n} for c, n in
                       zip(v.get("codes") or [], v.get("names") or [])][:6]
    for u in (idx.get("unmatched") or []):
        name = u.get("camp") if isinstance(u, dict) else u
        if not _looks_like_camp(name):
            버린쓰레기 += 1
            continue
        slot(name, "원장")

    # ── 갈래 나누기 ────────────────────────────────────────────────────────
    rows = []
    for row in camps.values():
        code, erp = row["거래처코드"].strip(), row["ERP거래처명"].strip()
        if row["후보"]:
            kind, why = "D", "후보 %d개: %s" % (
                len(row["후보"]), ", ".join("%s(%s)" % (c["code"], c["name"][:18])
                                            for c in row["후보"]))
        elif not code:
            # ★ **'없다'와 '같은 이름으로 못 찾았다'는 다른 말이다**(`[169]`).
            #   실측 2026-08-18: 코드 없는 483건 중 **469건이 밴드에서만** 보인다 —
            #   ERP 마스터에 그 이름이 없을 뿐, 다른 이름으로 있을 수 있다.
            if "ERP" in row["출처"]:
                kind, why = "F", "ERP·원장에는 있는데 거래처코드 칸이 비었다"
            else:
                kind, why = "E", "밴드에서만 보인다 — ERP 마스터에 **같은 이름이** 없다"
        elif not CODE_OK.match(code):
            kind, why = "C", "코드 `%s` 가 CU+숫자 형식이 아니다" % code
        elif erp and _norm(erp) != _norm(row["캠프명"]):
            kind, why = "B", "ERP 거래처명 `%s` ↔ 캠프명 `%s`" % (erp, row["캠프명"])
        else:
            kind, why = "A", ""
        rows.append({
            "갈래": kind, "캠프명": row["캠프명"], "거래처코드": code,
            "ERP거래처명": erp, "정기점검": "O" if row["정기점검"] else "",
            "건수": row["건수"], "출처": "+".join(sorted(row["출처"])),
            "다른표기": " / ".join(sorted(row["다른표기"]))[:120],
            "주소": row["주소"][:60], "확인할 것": why,
            "추정코드": "", "추정근거": "",
        })
    # ── E 를 ERP 거래처 원본과 대 본다 (노승용 요청 2번) ────────────────
    #   ★ 여기서 처음으로 '이름이 다른 캠프'를 잰다. 위 갈래 나누기는 캠프마스터를
    #     근거로 하는데 그 표는 `캠프명 == ERP거래처명` 인 것만 담아 구조상 못 잰다.
    custs, erp원본 = _erp_customers()
    Es = [r for r in rows if r["갈래"] == "E"]
    guess = guess_codes(Es, custs) if custs else {}
    for r in Es:
        cs = guess.get(_norm(r["캠프명"])) or []
        if len(cs) == 1:
            c = cs[0]
            r["갈래"], r["추정코드"] = "G", c["code"]
            r["추정근거"] = "%s — %s" % (c["name"], c["근거"])
            r["확인할 것"] = ("ERP `%s` (%s) 로 **보인다** — 근거 %s. 지목일 뿐 확정이 "
                            "아니다(아무것도 안 고쳤다)" % (c["name"], c["code"], c["근거"]))
        elif cs:
            r["갈래"] = "H"
            r["추정근거"] = " / ".join("%s(%s)" % (c["code"], c["name"][:20]) for c in cs[:6])
            r["확인할 것"] = "후보 %d개 — %s" % (len(cs), r["추정근거"])
        elif custs:
            r["확인할 것"] = "ERP 거래처 원본에도 댈 후보가 없다"
    # ★ **못 읽은 것을 '후보 없음'으로 세지 않는다**(`[169]`).
    if not custs:
        못읽음.append(("ERP 거래처 원본(ESA001M)", erp원본["길"] or "읽지 못했다 — "
                    "G·H 는 **못 잰 것**이지 0건이 아니다"))

    rows.sort(key=lambda r: (r["갈래"], -r["건수"], r["캠프명"]))
    return {"만든때": datetime.now().isoformat(timespec="seconds"),
            "rows": rows, "못읽음": 못읽음, "버린쓰레기": 버린쓰레기,
            "ERP원본": erp원본,
            "원천": {"ERP 거래처 마스터": len((master or {}).get("rows") or []),
                     "밴드 접수 글": len((contacts or {}).get("rows") or []),
                     "원장 색인(linked)": len((idx.get("linked") or {})),
                     "ERP 거래처 원본(ESA001M)": erp원본["건수"]}}


def to_md(d):
    L = ["# 캠프명 ↔ ERP 거래처코드 대조", "",
         "- 만든 때: %s" % d["만든때"],
         "- **읽기 전용입니다.** 이 표는 아무것도 고치지 않습니다 — "
         "유니웍스·ERP 등록은 담당자가 확인 후 진행합니다.", ""]
    if d["못읽음"]:
        L += ["## ★ 못 읽은 원천 — 아래 숫자는 그만큼 모자랍니다", ""]
        L += ["- **%s** — %s" % (n, why) for n, why in d["못읽음"]] + [""]
    L += ["## 원천마다 캠프 수가 다릅니다", "",
          "| 원천 | 캠프 수 |", "|---|---:|"]
    L += ["| %s | %d |" % (k, v) for k, v in d["원천"].items()]
    L += ["", "세 원천을 모두 합쳐 표기만 다른 것을 묶었습니다"
          "(대소문자·공백·괄호·하이픈만 무시 — **숫자는 안 건드립니다**: "
          "`송파1MB(감일동)` 과 `송파5MB(감일동)` 은 다른 캠프입니다).", ""]
    cnt = {}
    for r in d["rows"]:
        cnt[r["갈래"]] = cnt.get(r["갈래"], 0) + 1
    e = d.get("ERP원본") or {}
    if e:
        L += ["> **G·H 의 근거**: ERP 거래처 원본 `%s` %d건 (%s%s). "
              "낡았으면 그 사이 등록된 캠프는 안 보입니다 — 근거의 나이를 같이 적습니다."
              % (e.get("출처") or "?", e.get("건수") or 0, e.get("길") or "?",
                 (" · 받은 때 " + e["받은때"]) if e.get("받은때") else ""), ""]
    L += ["## ★ 이 표가 못 재는 것", "", CANNOT_MEASURE, "",
          "## 갈래", "", "| 갈래 | 무엇 | 건수 | 뜻 |", "|---|---|---:|---|"]
    for k, title, desc in KINDS:
        L.append("| **%s** | %s | %d | %s |" % (k, title, cnt.get(k, 0), desc))
    L.append("")
    if d["버린쓰레기"]:
        L += ["> 캠프 이름 모양이 아닌 조각 **%d건**은 표에서 뺐습니다"
              "(색인 파싱이 깨진 자리 — 조용히 빼지 않고 여기 적습니다)." % d["버린쓰레기"], ""]
    for k, title, desc in KINDS:
        sel = [r for r in d["rows"] if r["갈래"] == k]
        if not sel:
            continue
        L += ["## %s. %s (%d건)" % (k, title, len(sel)), "", desc, ""]
        if k in ("G", "H"):
            L += ["| 캠프명 | 캠프 주소 | 정기점검 | ERP 후보(추정) | 근거 |",
                  "|---|---|:-:|---|---|"]
        else:
            L += ["| 캠프명 | 거래처코드 | ERP 거래처명 | 정기점검 | 건수 | 출처 | 확인할 것 |",
                  "|---|---|---|:-:|---:|---|---|"]
        for r in sel[:80 if k != "A" else 40]:
            if k == "G":
                L.append("| %s | %s | %s | **%s** %s | %s |" % (
                    r["캠프명"], r["주소"][:30], r["정기점검"], r["추정코드"],
                    r["추정근거"].split(" — ")[0], r["추정근거"].split(" — ")[-1]))
            elif k == "H":
                L.append("| %s | %s | %s | %s | 자동으로 못 고른다 |" % (
                    r["캠프명"], r["주소"][:30], r["정기점검"], r["추정근거"]))
            else:
                L.append("| %s | %s | %s | %s | %d | %s | %s |" % (
                    r["캠프명"], r["거래처코드"] or "—", r["ERP거래처명"] or "—",
                    r["정기점검"], r["건수"], r["출처"], r["확인할 것"]))
        if len(sel) > (80 if k != "A" else 40):
            L.append("| … | | | | | | **나머지 %d건은 엑셀에 있습니다** |"
                     % (len(sel) - (80 if k != "A" else 40)))
        L.append("")
    return "\n".join(L)


def to_xlsx(d, path):
    """담당자가 나눠 작업할 수 있게 **새 파일**로 만든다(관리대장은 안 건드린다)."""
    try:
        from openpyxl import Workbook
    except Exception:                            # noqa: BLE001
        return ""
    wb = Workbook()
    ws = wb.active
    ws.title = "캠프_거래처코드"
    cols = ["갈래", "캠프명", "거래처코드", "추정코드", "추정근거", "ERP거래처명",
            "정기점검", "건수", "출처", "다른표기", "주소", "확인할 것", "담당자", "처리결과"]
    ws.append(cols)
    for r in d["rows"]:
        ws.append([r.get(c, "") for c in cols])
    for i, w in enumerate((6, 26, 12, 12, 40, 26, 8, 8, 14, 30, 40, 46, 12, 14), 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "A2"
    wb.save(path)
    return path


def main():
    d = build()
    os.makedirs(REPORT, exist_ok=True)
    with open(OUT_MD, "w", encoding="utf-8") as fh:
        fh.write(to_md(d))
    with open(OUT_MD.replace(".md", ".json"), "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=1)
    xl = to_xlsx(d, OUT_XLSX)
    cnt = {}
    for r in d["rows"]:
        cnt[r["갈래"]] = cnt.get(r["갈래"], 0) + 1
    print("캠프 %d개 · %s" % (len(d["rows"]),
                            " · ".join("%s %d" % (k, cnt.get(k, 0)) for k, _t, _d in KINDS)))
    if d["못읽음"]:
        print("★ 못 읽은 원천 %d개 — 위 숫자는 그만큼 모자랍니다" % len(d["못읽음"]))
    print("→ %s%s" % (os.path.relpath(OUT_MD, ROOT),
                      " · " + os.path.relpath(xl, ROOT) if xl else " (엑셀 생략: openpyxl 없음)"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
