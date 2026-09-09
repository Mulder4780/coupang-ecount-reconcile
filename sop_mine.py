# -*- coding: utf-8 -*-
"""표준 업무 절차서 — DB 에서 캐고, 자료가 늘면 판을 올린다 (2026-09-10 형님 지시)

형님 지시: "절차서의 경우 DB에서 자료 확인해서 할 수 있는 작업 진행하고
자료가 DB가 추가되면 계속 수정 해서 버전 업하는 알고리즘 구현해."

★ **작업 이름을 손으로 적지 않는다**([165]).  표에 적어 두면 그 밖의 작업은
  영영 안 걸리면서 오류도 안 난다.  여기서는 DB 문장을 **갈라 세어** 뽑고,
  적게 나온 것도 숫자로 말한다([169]).

★ **부위를 지어내지 않는다**([169]).  `도어락 체크` 가 전장부인지 프레임부인지
  우리는 모른다.  비워 두고 사람이 채운다 — '기타'로 채우면 그 절차서는
  **찾을 수 없는 자리에 묻힌다**.

★ **순서를 지어내지 않는다**.  단계 차례는 DB 문장에 실제로 적힌 **평균 자리**
  순이고, 그 사실을 절차서에 그대로 적는다.

★ **사람이 손댄 절차서는 안 덮는다**([326]).  판정은 `sop_store.upsert_auto`
  한 곳이 한다([162]) — 여기서 다시 정하지 않는다.

★ **Z: 를 한 글자도 안 만진다**([168]).  읽는 것은 로컬 DB 둘뿐이라
  회차와 다투지 않는다(실측 1초 미만).
"""

import collections
import json
import os
import re
import sqlite3
import sys

if hasattr(sys.stdout, "reconfigure"):          # 무인 회차는 sys.stdout 이 None 이다([235])
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

BASE = os.path.dirname(os.path.abspath(__file__))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

import sop_store as S                            # noqa: E402

APP_DB = os.path.join(BASE, "db", "app_store.db")
PM_DB = os.path.join(BASE, "db", "pm_content.db")

# 한 절차서에 담을 단계 상한.  넘으면 **자른 만큼을 말한다**([273]).
단계상한 = 12
# 이만큼은 나와야 단계로 올린다.  한 번 나온 말은 그 현장만의 메모일 수 있다.
최소건수 = 3

# 조각을 가르는 자리.  ★ 실측 문장에서 온 것이다 —
#   "(유료) 3개월점검비 (유료) 그리스 주입완료 (유료) WD-40 도포 및 …"
#   "그리스 주입·WD-40 도포·도어락 체크·유압유 교체"
_자르기 = re.compile(r"\((?:유료|무료)\)|[·/\n]|,\s|\s{3,}")
# 값이 아니라 서류 흐름을 적은 문장.  절차서 재료가 아니다.
_서류 = re.compile(r"판매전표|거래명세서|견적서|메일발송|세금계산서|입금|청구")
# 조각 끝에 붙는 말 — 떼어야 같은 작업이 한 줄로 모인다.
_꼬리 = re.compile(r"(완료|요청|함|했음|하였음|실시|진행)\s*$")
_호기 = re.compile(r"[▒?]{2,}|\d+\s*호기|\d+\s*R/?T|매립형|장비\s*\d+대")


def _ro(path):
    return sqlite3.connect("file:%s?mode=ro" % path.replace(os.sep, "/"), uri=True)


def _norm(piece):
    """조각 하나를 낱말로 다듬는다.  못 다듬으면 빈 문자열(= 안 센다)."""
    t = (piece or "").strip().strip("-—:·.")
    t = _호기.sub(" ", t)
    t = re.sub(r"\s+", " ", t).strip()
    t = _꼬리.sub("", t).strip()
    if len(t) < 3 or len(t) > 40:
        return ""
    if _서류.search(t):
        return ""
    if not re.search(r"[가-힣A-Za-z]", t):        # 숫자·기호만이면 작업이 아니다
        return ""
    return t


def _문장들():
    """DB 에서 작업이 적힌 문장을 모은다 — (작업종류, 문장) 목록."""
    out = []
    못읽음 = []

    try:
        c = _ro(APP_DB)
        q = ("select field_key, value_json from work_field "
             "where field_key in ('점검내용','신청내용','실제조치') "
             "and value_json is not null and length(value_json) >= 18")
        for key, raw in c.execute(q):
            try:
                t = json.loads(raw)
            except Exception:
                t = raw
            if not isinstance(t, str):
                t = str(t)
            종류 = "정기점검" if key == "점검내용" else "돌발AS"
            out.append((종류, t))
        c.close()
    except Exception as exc:
        못읽음.append("app_store: %s" % exc)

    try:
        c = _ro(PM_DB)
        q = ("select content_excerpt from pm_content "
             "where content_excerpt is not null and length(content_excerpt) >= 20")
        for (t,) in c.execute(q):
            out.append(("정기점검", t))
        c.close()
    except Exception as exc:
        못읽음.append("pm_content: %s" % exc)

    return out, 못읽음


def survey():
    """무엇이 얼마나 있나 — **읽기만** 한다.  아무것도 안 담는다."""
    문장, 못읽음 = _문장들()
    본 = collections.Counter()
    셈 = collections.defaultdict(collections.Counter)
    자리 = collections.defaultdict(lambda: collections.defaultdict(list))
    for 종류, t in 문장:
        본[종류] += 1
        조각 = [x for x in _자르기.split(t) if x]
        낱말 = []
        for i, p in enumerate(조각):
            w = _norm(p)
            if w:
                낱말.append((i, w))
        for i, w in 낱말:
            셈[종류][w] += 1
            자리[종류][w].append(i)
    return {"문장수": dict(본), "낱말": {k: v for k, v in 셈.items()},
            "자리": {k: {w: (sum(v) / float(len(v))) for w, v in d.items()}
                     for k, d in 자리.items()},
            "못읽음": 못읽음}


def draft(종류):
    """그 작업 종류의 절차서 초안 한 건.  근거가 없으면 None(지어내지 않는다)."""
    d = survey()
    낱말 = d["낱말"].get(종류) or {}
    자리 = d["자리"].get(종류) or {}
    문장수 = (d["문장수"] or {}).get(종류, 0)
    쓸것 = [(w, n) for w, n in 낱말.items() if n >= 최소건수]
    if not 쓸것:
        return None, d
    # ★ 차례는 **DB 문장에 실제로 적힌 평균 자리** 순이다 — 우리가 정하지 않는다.
    쓸것.sort(key=lambda x: (자리.get(x[0], 99), -x[1]))
    잘림 = max(0, len(쓸것) - 단계상한)
    쓸것 = 쓸것[:단계상한]

    단계 = [{"내용": w, "주의": "DB %d건에서 확인" % n} for w, n in 쓸것]
    마무리 = ("이 초안은 앱 DB %d건의 %s 기록에서 캔 것입니다. "
              "차례는 DB 문장에 적힌 평균 자리 순이며 **현장 순서가 아닙니다** — "
              "차동호 팀장이 확인해 주셔야 합니다." % (문장수, 종류))
    if 잘림:
        마무리 += " (%d건 더 있었으나 %d단계까지만 실었습니다)" % (잘림, 단계상한)

    row = {
        "id": "SOP-AUTO-%s" % 종류,
        "제목": "%s 표준 작업(DB 초안)" % 종류,
        "부위": "",                               # ★ 지어내지 않는다 — 사람이 채운다
        "작업종류": 종류,
        "공구": [],                               # ★ DB 에 안 적혀 있다 — 비운다
        "단계": 단계,
        "마무리": 마무리,
        "주의사항": "자동 생성분입니다. 사람이 고치면 그 뒤로는 자동으로 안 덮습니다.",
        "출처": "앱 DB(work_field · pm_content) 자동 수집",
    }
    return row, d


def run(apply=False, path=None):
    """DB 를 훑어 초안을 담는다.  `apply` 가 아니면 **미리보기**다."""
    말 = []
    d = None
    for 종류 in ("정기점검", "돌발AS"):
        row, d = draft(종류)
        if row is None:
            말.append(("근거없음", "SOP-AUTO-%s" % 종류,
                       "%d건 이상 되풀이되는 작업이 없다 — 담지 않는다" % 최소건수))
            continue
        if not apply:
            말.append(("미리보기", row["id"], "단계 %d개" % len(row["단계"])))
            continue
        말.append(S.upsert_auto(row, path=path))
    return 말, (d or {})


def _main(argv):
    import argparse
    ap = argparse.ArgumentParser(description="DB 에서 절차서 초안을 캐고 판을 올린다")
    ap.add_argument("--apply", action="store_true", help="실제로 담는다(없으면 미리보기)")
    ap.add_argument("--survey", action="store_true", help="무엇이 얼마나 있나만 본다")
    a = ap.parse_args(argv)

    if a.survey:
        d = survey()
        print("문장:", d["문장수"])
        for 종류, c in sorted(d["낱말"].items()):
            print("[%s] 낱말 %d가지" % (종류, len(c)))
            for w, n in c.most_common(12):
                print("   %4d건  %s" % (n, w))
        if d["못읽음"]:
            # ★ 못 읽은 것을 조용히 넘기지 않는다([169]).
            print("못읽음:", " / ".join(d["못읽음"]))
        return 0

    말, d = run(apply=a.apply)
    for 갈래, sid, 설명 in 말:
        print("%-8s %-22s %s" % (갈래, sid, 설명))
    if d.get("못읽음"):
        print("못읽음:", " / ".join(d["못읽음"]))
    if not a.apply:
        print("(미리보기입니다 — 실제로 담으려면 --apply)")
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
