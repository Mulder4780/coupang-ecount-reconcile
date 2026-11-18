# -*- coding: utf-8 -*-
"""밴드 글에서 뽑은 **원장 미등록** 건을 원장 입력 큐에 올린다.

왜 생겼나 (2026-11-18 형님 지시 "밴드 넣는 길 만들어")
------------------------------------------------------
그날 실측: 밴드 글에서 프로젝트번호 9,178개를 읽어 원장과 대 보니 **179개가
원장에 없었다.** 그런데 `band_extract` 는 **보고만 하고 넣는 길이 없었다** —
카톡에는 `--queue` 가 있는데 밴드에는 없다. 그래서 9월 말 정기점검 16건,
11월 정기점검 14건이 뭉텅이로 빠진 채 아무도 못 넣고 있었다. 원장에 없으면
**청구 근거가 없다.**

판단을 두 벌로 만들지 않는다([162])
-----------------------------------
행을 고르고 셀을 만드는 것은 **`kakao_extract.build_queue` 하나**가 한다 —
번호 없이 내용만 적힌 기존 행을 찾아 채우는 규칙(같은 작업이 두 행이 되는 것을
막는다)·여유행 판정·시트별 열 지도가 전부 거기 있다. 여기서 다시 만들면
언젠가 갈리고, 갈린 뒤엔 어느 쪽이 맞는지 아무도 모른다.
이 파일이 하는 일은 **밴드 기록을 그 함수가 읽는 모양으로 옮기는 것**뿐이다.

모르는 것은 안 채운다([169])
----------------------------
- **접수일자는 비워 둔다.** 밴드 글은 *작업을 마치고* 올리는 보고라 "언제
  접수됐나"가 적혀 있지 않다. 게시일을 접수일로 적으면 **없는 사실을 만든다.**
- `진행상태` 가 `작업완료` 면 작업일을 **완료일**로, `접수·예정` 이면 **예정일**로
  넣는다. 둘 다 아니면 **어느 칸에도 안 넣는다.**
- `업무유형` 이 `기타`·`동시진행` 인 건은 **보류**한다 — 어느 시트로 가야 할지
  글만 봐서는 모른다. 조용히 버리지 않고 몇 건인지 말한다([273]).
- `build_queue` 가 만드는 셀은 전부 `only_if_empty=True` 다 — 이미 값이 있는
  칸은 안 덮는다.

쓰는 법
-------
    python band_to_ledger.py                      # 무엇이 올라갈지 보기만 한다
    python band_to_ledger.py --since 2026-10-01   # 그날부터(작업일 기준)
    python band_to_ledger.py --since 2026-10-01 --queue   # 실제로 큐에 적재

★ 큐에 올리는 것까지다. 원장 반영은 `ledger_db --intake` 가, Excel 보관본은
  주 1회 회차가 한다([417]) — 이 도구는 엑셀을 열지도 쓰지도 않는다.
"""

import argparse
import sys


# 밴드의 '업무유형' → 원장 시트. 여기 없는 유형은 보류한다([166] — 낱말을 지어내지 않는다).
SHEET_OF = {
    "돌발AS": "02_돌발AS접수",
    "정기점검": "04_정기점검",
    "신규납품설치": "05_신규납품설치",
}

DONE = "작업완료"
OPEN = "접수·예정"


def _body_summary(text, limit=200):
    """글 본문에서 'A/S 내용' 아래를 꺼낸다 — 없으면 본문 앞부분."""
    s = str(text or "")
    for mark in ("A/S 내용", "AS 내용", "점검 내용", "작업 내용"):
        i = s.find(mark)
        if i >= 0:
            s = s[i + len(mark):]
            break
    s = s.lstrip(" :：\n")
    return " ".join(s.split())[:limit]


def to_rows(recs):
    """밴드 기록 → `kakao_extract.build_queue` 가 읽는 모양.

    돌려주는 것: (rows, held) — held 는 시트를 못 고른 건들(사람이 봐야 한다).
    """
    rows, held = [], []
    for r in recs:
        sheet = SHEET_OF.get(r.get("업무유형") or "")
        if not sheet:
            held.append(r)
            continue
        day = r.get("작업일") or ""
        state = r.get("진행상태") or ""
        body = _body_summary(r.get("본문"))
        rows.append({
            "시트": sheet,
            "프로젝트NO": r.get("프로젝트NO") or "",
            "캠프명": r.get("캠프명") or "",
            "담당기사": r.get("담당기사") or "",
            # ★ 접수일자는 모른다 — 비워 둔다(위 설명).
            "신청일자": "",
            "예정일": day if state == OPEN else "",
            "완료일": day if state == DONE else "",
            "신청내용": body,
            "작업내용": body,
            # 비고·근거에 실릴 출처
            "게시일": r.get("게시일") or "",
            "밴드": r.get("밴드") or "",
            "설비": "",
            "담당원문": r.get("담당기사") or "",
        })
    return rows, held


def pick(recs, since=None, until=None):
    """원장 미등록 + 작업일 구간으로 고른다. 프로젝트번호당 한 건만 남긴다."""
    seen, out = set(), []
    for r in recs:
        if r.get("원장등록") != "미등록":
            continue
        no = r.get("프로젝트NO") or ""
        if not no or no in seen:
            continue
        day = r.get("작업일") or ""
        if since and (not day or day < since):
            continue
        if until and (not day or day > until):
            continue
        seen.add(no)
        out.append(r)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="밴드 미등록 건을 원장 입력 큐에 올린다")
    ap.add_argument("--since", help="작업일 이 날짜부터 (예: 2026-10-01)")
    ap.add_argument("--until", help="작업일 이 날짜까지")
    ap.add_argument("--queue", action="store_true",
                    help="실제로 큐에 적재한다(없으면 보기만 한다)")
    args = ap.parse_args(argv)

    import band_extract
    from ecount_reconcile import load_config, resolve_master
    import kakao_extract

    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    recs = band_extract.load_records()
    known = band_extract.ledger_projects(master)
    for r in recs:
        r["원장등록"] = "등록됨" if r.get("프로젝트NO") in known else "미등록"

    picked = pick(recs, args.since, args.until)
    print("밴드 미등록 %d건 (작업일 %s ~ %s)"
          % (len(picked), args.since or "처음", args.until or "끝"))
    if not picked:
        return 0

    rows, held = to_rows(picked)
    for r in held:
        print("  [보류] %s %s — 업무유형 '%s' 는 어느 시트인지 모른다"
              % (r.get("프로젝트NO"), (r.get("캠프명") or "")[:20], r.get("업무유형")))

    queue, plan, blocked = kakao_extract.build_queue(
        rows, master, source="밴드", stamp_field="게시일", file_field="밴드")
    for b in blocked:
        print("  [보류]", b)

    print("\n등록 대상 %d건 / 셀 %d개 / 보류 %d건"
          % (len(plan), len(queue), len(held) + len(blocked)))
    for p in plan:
        print("   %s %s %s %s" % (p["시트"], p["프로젝트NO"],
                                  (p["캠프명"] or "")[:22], p["예정일"] or ""))

    if not args.queue:
        print("\n(보기만 했습니다 — 실제로 올리려면 --queue 를 붙이십시오)")
        return 0

    if not queue:
        print("올릴 셀이 없습니다.")
        return 0

    from ledger_writer import queue_add
    added = queue_add(queue)
    print("큐 적재: %d 개 셀 → ledger_db --intake 가 원장에 넣습니다" % added)
    return 0


if __name__ == "__main__":
    sys.exit(main())
