# -*- coding: utf-8 -*-
"""추천 등급을 **확정 칸에 반영한다** (2026-09-10 형님 지시).

형님 지시: **"미분류건 과 추천안 전부 반영해서 앱에 반영하고 캡처고 뭐고 싹 다 반영해"**

★★ 이것은 2026-09-09 지시("미분류로 남기고 사람이 정하기")를 **형님이 스스로
   바꾸신 것**이다.  그 전까지 `as_grade` 는 추천을 확정 칸에 절대 안 넣었고
   그 판단은 그때 옳았다 — 오늘 형님이 그 칸을 채우라고 정하셨다.

★ 저장 길을 새로 만들지 않는다([162]) — `app_server.save_staff_entry` 그대로다.
  그래야 **멱등키 · 낙관잠금 · 감사로그 · 등급 검사**(`as_grade.check`)가 공짜로 붙는다.
  A긴급은 사유가 없으면 그 검사가 막으므로 `추천근거` 를 사유로 같이 넣는다.

★ **사람이 이미 고른 값은 한 글자도 안 건드린다**([172]).  `_add_grade_hint` 가
  그런 행에는 추천을 아예 안 붙이므로 여기에 오지도 않는다.

★ **미리보기가 기본이다** — 적용은 `--apply`.  되돌리기는 앱에서 그 칸을 다시
  고르는 것이고, 누가 언제 무엇으로 바꿨는지는 감사로그에 남는다.

★ 진도를 남긴다([406]·[427]) — 중간에 죽어도 다음 회차가 이어받는다.
  이미 확정이 있는 행은 건너뛰므로 **여러 번 돌려도 안전하다**.

⚠ `추천작업뒤` 인 건(실측 12건)은 **작업이 끝난 뒤 적은 글**에서 나온 추천이라
  접수 시점 판단이 아니다.  기본으로 **같이 반영하되 근거에 그 사실을 적는다** —
  빼면 그 12건이 미분류로 남아 형님 지시("전부 반영")를 반쪽만 지킨다([169]).
  빼려면 `--skip-after-work`.

쓰는 법:
    python as_grade_apply.py                  # 미리보기 (아무것도 안 쓴다)
    python as_grade_apply.py --apply          # 반영
    python as_grade_apply.py --apply --limit 50
"""

import io
import json
import os
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
_WEB = os.path.join(ROOT, "webapp")
if _WEB not in sys.path:
    sys.path.insert(0, _WEB)

PROGRESS = os.path.join(ROOT, "reports", "돌발AS_등급반영.json")

# 돌발AS 칸이 열려 있는 업무센터 - 권한 검사가 이 slug 로 돌아간다.
# ⚠ `"관리자"` 는 `STAFF_CENTERS` 의 열쇠가 **아니다** - 실측으로 20건이
#   `등록되지 않은 업무센터입니다` 로 거절됐다(그랬도 **한 글자도 안 썼다**).
STAFF_SLUG = "ryu-jiyeong"
ACTOR = "claude:as_grade_apply(2026-09-10 형님 지시)"

# 한 회차 예산.  ★ 넘으면 **새로 안 넣고** 진도를 말하고 돌아온다([427]) —
# 다음 부름이 이어받는다(이미 확정이 있는 행은 건너뛴다).
BUDGET_SEC = float(os.environ.get("AS_GRADE_APPLY_BUDGET_SEC") or 0) or 0.0


def _now():
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _save(doc):
    """진도를 남긴다 — **죽어도 남는다. 그게 요점이다**([180])."""
    try:
        os.makedirs(os.path.dirname(PROGRESS), exist_ok=True)
        tmp = PROGRESS + ".tmp"
        with io.open(tmp, "w", encoding="utf-8", newline="\n") as f:
            f.write(json.dumps(doc, ensure_ascii=False, indent=2) + "\n")
        os.replace(tmp, PROGRESS)
    except Exception:                          # noqa: BLE001
        pass                                   # 자국 하나로 회차를 안 죽인다


def _row_key(r):
    """저장 열쇠를 만든다 — **한 곳**이다([162]).

    ★ 서버(`_staff_store_row`)는 `key_col`(접수ID) **또는 `_business_key`**
      (프로젝트NO)로 찾는다.  그런데 예전에는 여기서 `접수ID` 하나만 물어
      **답이 있는 46건이 조용히 빠졌다**(2026-09-10 실측 · 46건 전부
      프로젝트NO 로 찾히고 전부 유일했다).  오류도 안 나고 미리보기는
      "열쇠 없음"이라 말해서 아무도 몰랐다([165]·[300]).

    ★ **접수ID 가 먼저다**([172]) — 그것이 있으면 지금 동작이 한 톨도 안 바뀐다.
    ★ 둘 다 없으면 빈 문자열이다 — **지어내지 않는다**([169]).
    ★ 프로젝트NO 가 앱 DB에 여러 건이면 서버가 거절한다(`len(matches) != 1`).
      여기서 그 판정을 **베끼지 않는다**([162]) — 베끼면 사본이 둘 되어 갈린다.
      거절되면 실패 건수와 사유가 그대로 남는다.
    """
    return (str(r.get("접수ID") or "").strip()
            or str(r.get("프로젝트NO") or "").strip())


def pick(rows, skip_after_work=False):
    """반영할 행을 고른다.  돌려주는 것은 `(고른 것, 왜 뺐나)`."""
    import as_grade as G
    out, why = [], {"이미 확정": 0, "추천 없음": 0, "작업 뒤 글": 0, "열쇠 없음": 0}
    for r in (rows or {}).get("as") or []:
        if not isinstance(r, dict):
            continue
        if str(r.get("대응등급") or "").strip():
            why["이미 확정"] += 1
            continue
        g = str(r.get("추천등급") or "").strip()
        if not g or g == G.GRADE_UNSET:
            why["추천 없음"] += 1
            continue
        if skip_after_work and r.get("추천작업뒤"):
            why["작업 뒤 글"] += 1
            continue
        if not _row_key(r):
            # ★ 열쇠가 없으면 저장할 자리를 못 찾는다 — 지어내지 않는다([169]).
            why["열쇠 없음"] += 1
            continue
        out.append(r)
    return out, why


def _body(r):
    """`save_staff_entry` 가 받는 모양으로 만든다."""
    import as_grade as G
    grade = str(r.get("추천등급") or "").strip()
    type_ = str(r.get("추천유형") or "").strip()
    if type_ and type_ not in G.TYPES.get(grade, ()):
        type_ = ""                             # 어긋나면 등급만 — 지어내지 않는다
    # ★ A긴급은 사유가 없으면 `as_grade.check` 가 막는다.  추천근거가 곧 그 사유다.
    #   어느 칸에서 나온 추천인지 같이 적는다([169]) — `실제작업상세` 는 **작업이
    #   끝난 뒤** 적는 글이라 접수 시점 판단이 아니다.
    why = str(r.get("추천근거") or "").strip() or "증상 글에서 뽑은 추천"
    src = str(r.get("추천출처") or "").strip()
    tail = " · 근거: " + src if src else ""
    if r.get("추천작업뒤"):
        tail += " (" + G.AFTER_WORK_NOTE + ")"
    reason = "기계 추천 일괄 반영(2026-09-10 형님 지시) · " + why + tail
    key = _row_key(r)
    return {
        "category": "as",
        "key": key,
        "record_version": int(r.get("DB버전") or r.get("_record_version") or 0),
        "values": {"대응등급": grade, "대응유형": type_, "등급사유": reason[:900]},
        "reason": reason[:300],
        # ★ 멱등키 — 같은 건을 두 번 눌러도 한 번만 반영된다.
        "request_id": "asgrade-20260910-" + key,
    }


def run(apply=False, limit=0, skip_after_work=False):
    import app_server as A
    rows = A.get_works() or {}
    todo, why = pick(rows, skip_after_work=skip_after_work)
    if limit:
        todo = todo[:limit]

    grades = {}
    for r in todo:
        grades[r.get("추천등급")] = grades.get(r.get("추천등급"), 0) + 1
    print("반영 대상 %d건 %s" % (
        len(todo), " · ".join("%s %d" % (k, v) for k, v in sorted(grades.items()))))
    print("뺀 것: " + " · ".join("%s %d" % (k, v) for k, v in why.items() if v))
    if not apply:
        print("\n(미리보기입니다 — 아무것도 안 썼습니다. 반영하려면 --apply)")
        for r in todo[:6]:
            print("  ", r.get("접수ID"), "|", r.get("캠프명"), "|",
                  r.get("추천등급"), "/", r.get("추천유형"), "|",
                  str(r.get("추천근거"))[:52])
        return 0

    t0 = time.time()
    ok = fail = 0
    errs = {}
    doc = {"때": _now(), "대상": len(todo), "반영": 0, "실패": 0, "왜뺐나": why}
    for i, r in enumerate(todo):
        if BUDGET_SEC and (time.time() - t0) > BUDGET_SEC:
            doc["이어감"] = "예산(%.0f초)이 다 됐다 - %d/%d 까지 했다" % (
                BUDGET_SEC, i, len(todo))
            print("!", doc["이어감"], "· 다음 부름이 이어받는다")
            break
        try:
            # ★ slug 는 **권한**을 정하고(돌발AS 칸이 열려 있는 업무센터),
            #   `actor` 는 **누가 했나**를 적는다.  감사로그에 류지영이 한 것처럼
            #   남기면 그것은 **거짓**이다([169]) - 나중에 "누가 이걸 A로 했나"를
            #   물을 사람이 반드시 있다.
            A.save_staff_entry(STAFF_SLUG, _body(r), actor=ACTOR)
            ok += 1
        except Exception as e:                 # noqa: BLE001
            fail += 1
            k = type(e).__name__ + ": " + str(e)[:110]
            errs[k] = errs.get(k, 0) + 1
        if (i + 1) % 50 == 0:
            doc["반영"], doc["실패"] = ok, fail
            _save(doc)
            print("  ... %d/%d (반영 %d · 실패 %d)" % (i + 1, len(todo), ok, fail))
    doc["반영"], doc["실패"], doc["오류"] = ok, fail, errs
    doc["걸린초"] = round(time.time() - t0, 1)
    _save(doc)
    print("\n반영 %d건 · 실패 %d건 · %.1f초" % (ok, fail, doc["걸린초"]))
    # ★ 실패를 조용히 넘기지 않는다([169]) — 무엇이 왜 안 됐는지 적는다.
    for k, v in sorted(errs.items(), key=lambda x: -x[1])[:6]:
        print("  실패 %3d · %s" % (v, k))
    print("자국:", PROGRESS)
    return 0 if fail == 0 else 1


def main(argv=None):
    argv = list(argv if argv is not None else sys.argv[1:])
    known = {"--apply", "--limit", "--skip-after-work", "--help", "-h"}
    bad = [a for a in argv if a.startswith("-") and a.split("=")[0] not in known]
    if bad:
        # ★ 모르는 깃발을 조용히 무시하지 않는다([93]) — 오타 하나가 '0건 성공'이 된다.
        print("모르는 깃발:", " ".join(bad), "· 쓸 수 있는 것:", " ".join(sorted(known)))
        return 2
    if "--help" in argv or "-h" in argv:
        print(__doc__)
        return 0
    limit = 0
    if "--limit" in argv:
        i = argv.index("--limit")
        if i + 1 < len(argv):
            limit = int(argv[i + 1])
    return run(apply=("--apply" in argv), limit=limit,
               skip_after_work=("--skip-after-work" in argv))


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")   # 무인 회차는 stdout 이 None 일 수 있다([235])
    except Exception:
        pass
    raise SystemExit(main())
