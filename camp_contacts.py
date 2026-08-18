# -*- coding: utf-8 -*-
"""
camp_contacts.py — **전국 쿠팡캠프 담당자 한 장**: 캠프명 · 담당자 이름 · 전화번호

유수비 대표 지시(2026-08-18 07:33 카톡): "쿠팡 전국정기점검 캠프와 담당자리스트
오늘중으로 주세요"

왜 새로 만드나 (실측 2026-08-18)
  · 캠프 목록은 이미 있다 — `reports/캠프마스터.json` 268개.
  · 그런데 그중 **연락처가 5개뿐**이었다. ERP 거래처등록(ESA001M)을 대 보니
    CU(캠프) 356개 중 전화번호가 **18개**다 — 열을 안 읽은 것이 아니라([165] 를
    먼저 의심했다) **ERP 에 안 적혀 있는 것**이다. 다른 거래처는 423개가 차 있다.
  · 진짜 원천은 **밴드 접수 글 본문**이다(9,557건 중 6,246건에 번호가 있다):
        ● 캠프이름 : 울산2캠프
        ● 캠프주소 : 울산광역시 …
        ● 현장책임 : 제석화
        ● 담당번호 : 010-7532-8543
        ● 안전관리 : 이상협
        ● 담당번호 : 010-2511-4947 / waynelee@coupang.com
    캠프마다 **현장책임·안전관리 두 사람**이 있고 셋(이름·전화·메일) 다 있다.

지키는 것
  · **읽는 자리는 하나다** — 뽑기는 `band_extract.parse_post` 가 한다([162]).
    여기서 본문을 다시 정규식으로 뜯지 않는다. 갈리면 한쪽만 고쳐진다.
  · **담당자는 바뀐다.** 캠프마다 **최신 게시일** 것을 채택하고 근거(게시일·밴드·
    글번호)를 같이 남긴다. 언제 자료인지 안 밝히면 낡은 번호를 확언하게 된다.
  · **못 뽑은 캠프를 '담당자 없음'이라 하지 않는다** — 밴드에 접수 글이 없으면
    그것은 **모르는 것**이다([169]). `근거` 칸이 비면 화면이 그렇게 적는다.
  · **짐작으로 채우지 않는다.** 캠프명이 비슷하다고 옆 캠프 번호를 붙이면 대표
    보고에 틀린 번호가 실린다 — 빈 칸보다 나쁘다([172]).
  · **Z: 를 안 훑는다** — 밴드 캐시만 읽는다([168]). 실측 3초.
  · 원문은 한 글자도 안 고친다.

산출: reports/캠프_담당자.json     (앱 `/api/camps` 가 이것만 읽는다)

  python camp_contacts.py            (사람이 볼 요약)
  python camp_contacts.py --write    (파일 생성 · 09:50 회차 단계)
"""
import json
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:  # 무인 회차는 pythonw 라 sys.stdout 이 None 이다 — [235]
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import band_extract as BE  # noqa: E402

REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
OUT = os.path.join(REPORT_DIR, "캠프_담당자.json")
MASTER = os.path.join(REPORT_DIR, "캠프마스터.json")

# 정기점검으로 세는 업무유형. `band_extract` 가 실제로 쓰는 낱말이며 여기서
# 새 낱말을 지어내지 않는다.
PM_KINDS = ("정기점검",)


# 캠프명 칸에 **사람이 메모를 적은** 글이 있다(실측: `서초1MB(양재동) = 서초1Sub-hub -
# 2RT 안전수칙은 …` 처럼 58자). 그런 이름은 **자동으로 자르지 않는다** — 잘라 붙이면
# 엉뚱한 캠프에 번호가 들어간다([172]). 대신 표시만 갈라 사람이 판단하게 둔다([169]).
NAME_MAX = 24


def _norm(name):
    """캠프명 맞추기 — **대소문자·공백·괄호·하이픈만** 없앤다.

    실측(2026-08-18): `울산1Sub-hub` / `울산1sub-hub`, `전주1 Sub Hub` /
    `전주1SUBHUB` / `전주1Sub-hub`, `M_익산1` / `M익산1` 처럼 **표기만 다른**
    묶음이 19개였다. 안 합치면 같은 캠프가 목록에 여러 줄로 뜬다.

    ★ 그래도 **숫자는 안 건드린다.** `송파1MB(감일동)` 과 `송파5MB(감일동)` 은
      실재하는 **다른 캠프**다([172] 의 실측 오탐) — 앞 숫자가 달라 안 합쳐진다.
    """
    import re as _re
    return _re.sub(r"[\s()\[\]_\-/·.]", "", str(name or "")).upper()


def build():
    recs = BE.load_records()
    camps = {}
    for r in recs:
        camp = (r.get("캠프명") or "").strip()
        if not camp:
            continue
        key = _norm(camp)
        posted = str(r.get("게시일") or "")
        c = camps.setdefault(key, {
            "캠프명": camp, "캠프주소": "", "현장책임": None, "안전관리": None,
            # 옛 양식(2023~2024)은 직책 없이 `● 담당자명` 한 사람이다 — 실측 2,714글.
            # 현장책임 칸에 합치지 않는다(원문이 직책을 말한 적 없다 · [172]).
            "담당자": None,
            "근거": {}, "정기점검건수": 0, "돌발AS건수": 0,
            "최근작업일": "", "총건수": 0,
        })
        # 표시 이름은 **가장 최근 글에 쓰인 그대로**를 쓴다(양식이 바뀌면 따라간다).
        c["총건수"] += 1
        if r.get("업무유형") in PM_KINDS:
            c["정기점검건수"] += 1
        elif r.get("업무유형") == "돌발AS":
            c["돌발AS건수"] += 1
        wd = str(r.get("작업일") or "")
        if wd > c["최근작업일"]:
            c["최근작업일"] = wd

        # ★ 사람 칸은 **각각 따로** 최신을 고른다. 한 글에 현장책임만 적히고
        #   안전관리는 빈 글이 흔하다 — 글 하나를 통째로 이기게 하면 애써 있는
        #   다른 칸이 빈 값에 덮인다.
        for slot in ("현장책임", "안전관리", "담당자"):
            p = r.get(slot)
            if not p:
                continue
            prev = c["근거"].get(slot, {})
            if posted >= str(prev.get("게시일") or ""):
                c[slot] = p
                c["근거"][slot] = {"게시일": posted, "밴드": r.get("밴드"),
                                   "글번호": r.get("게시글")}
        if r.get("캠프주소") and posted >= str(c["근거"].get("주소", {}).get("게시일") or ""):
            c["캠프주소"] = r["캠프주소"]
            c["근거"]["주소"] = {"게시일": posted}
        # ★ 표시 이름은 최신이 아니라 **가장 자주 쓰인 표기**다. 최신 글 하나가
        #   메모 섞인 이름이면(실측 있음) 그 캠프가 목록에서 이상해진다.
        c.setdefault("표기", {})
        c["표기"][camp] = c["표기"].get(camp, 0) + 1

    # ERP 거래처코드를 붙인다 — **이미 만들어 둔 캠프마스터를 읽기만** 한다.
    #   여기서 ERP 를 다시 훑으면 웹 요청 뒤에서 Z: 재귀 탐색이 된다([168]).
    master = {}
    try:
        with open(MASTER, encoding="utf-8") as f:
            for row in (json.load(f).get("rows") or []):
                master[_norm(row.get("캠프명"))] = row
    except Exception:
        master = {}   # 못 읽었으면 '없다'가 아니라 **안 붙인다**

    rows = []
    for key, c in camps.items():
        m = master.get(key) or {}
        # 같은 캠프의 여러 표기 중 **가장 많이 쓰인 것**. 같은 횟수면 짧은 쪽 —
        # 긴 쪽은 메모가 섞여 있을 확률이 높다.
        tally = c.get("표기") or {c["캠프명"]: 1}
        name = sorted(tally.items(), key=lambda kv: (-kv[1], len(kv[0])))[0][0]
        rows.append({
            "캠프명": name,
            # ★ 자동으로 자르지 않는다 — 화면이 '이름 확인 필요'로 갈라 적는다.
            "이름확인필요": len(name) > NAME_MAX,
            "다른표기": sorted(k for k in tally if k != name)[:4],
            "캠프주소": c["캠프주소"] or (m.get("주소") or ""),
            "거래처코드": m.get("거래처코드") or "",
            "정기점검": c["정기점검건수"] > 0,
            "정기점검건수": c["정기점검건수"],
            "돌발AS건수": c["돌발AS건수"],
            "총건수": c["총건수"],
            "최근작업일": c["최근작업일"],
            "현장책임": c["현장책임"] or {},
            "안전관리": c["안전관리"] or {},
            "담당자": c["담당자"] or {},
            "근거": c["근거"],
        })

    # 밴드에 글이 **한 번도 없는** 캠프도 목록에 남긴다 — 빼면 '없는 캠프'가 된다.
    #   담당자 칸은 비고 `근거` 도 비므로 화면이 '모름'이라 적는다([169]).
    for key, m in master.items():
        if key in camps:
            continue
        rows.append({
            "캠프명": m.get("캠프명") or "", "캠프주소": m.get("주소") or "",
            "이름확인필요": len(m.get("캠프명") or "") > NAME_MAX, "다른표기": [],
            "거래처코드": m.get("거래처코드") or "",
            "정기점검": False, "정기점검건수": 0, "돌발AS건수": 0, "총건수": 0,
            "최근작업일": m.get("최근작업일") or "",
            # ERP 에 적힌 담당자가 있으면 그것만은 싣는다(출처를 밝힌다).
            # ★ 여기도 **직책 미상**이다 — ERP 거래처등록의 담당자 칸은 현장책임인지
            #   안전관리인지 말하지 않는다. 그래서 `담당자` 칸으로 담는다([172]).
            "현장책임": {}, "안전관리": {},
            "담당자": ({"이름": m.get("담당자") or "", "전화": m.get("연락처") or "",
                        "메일": ""} if (m.get("담당자") or m.get("연락처")) else {}),
            "근거": ({"담당자": {"출처": "ERP 거래처등록"}}
                     if (m.get("담당자") or m.get("연락처")) else {}),
        })

    # 이름이 수상한 것은 **지우지 않고 맨 뒤로** 보낸다([169]).
    rows.sort(key=lambda r: (r["이름확인필요"], not r["정기점검"],
                             -r["정기점검건수"], r["캠프명"]))
    tel = sum(1 for r in rows
              if (r["현장책임"].get("전화") or r["안전관리"].get("전화")
                  or r["담당자"].get("전화")))
    return {
        "갱신": BE.datetime.now().isoformat(timespec="seconds")
        if hasattr(BE, "datetime") else "",
        "캠프수": len(rows),
        "정기점검캠프수": sum(1 for r in rows if r["정기점검"]),
        "전화있음": tel,
        "전화모름": len(rows) - tel,
        "이름확인필요": sum(1 for r in rows if r["이름확인필요"]),
        "출처": "밴드 접수 글 본문(band_extract) + reports/캠프마스터.json",
        "rows": rows,
    }


def main():
    import datetime as _dt
    d = build()
    d["갱신"] = _dt.datetime.now().isoformat(timespec="seconds")
    if "--write" in sys.argv:
        os.makedirs(REPORT_DIR, exist_ok=True)
        tmp = OUT + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        os.replace(tmp, OUT)
    print(f"캠프 {d['캠프수']}개 · 정기점검 {d['정기점검캠프수']}개 · "
          f"전화 있음 {d['전화있음']} · 모름 {d['전화모름']}"
          + (f" → {OUT}" if "--write" in sys.argv else ""))
    for r in d["rows"][:5]:
        # 화면과 같은 차례로 고른다 — 현장책임이 없으면 직책 미상 담당자가 ①이다.
        s = r["현장책임"] or r["담당자"] or {}
        print(f"  {r['캠프명']:<18} 정기{r['정기점검건수']:>3} "
              f"{s.get('이름','') or '-':<8} {s.get('전화','') or '-'}")


if __name__ == "__main__":
    main()
