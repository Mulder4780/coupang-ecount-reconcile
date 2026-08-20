# -*- coding: utf-8 -*-
"""
comment_backfill.py — **댓글을 한 번도 안 들여다본 글**을 골라 수집거리로 만든다

왜 따로 있나
  수집기가 챙기는 것은 늘 '없는 글을 채우는' 방향이었다 — `recheck_plan` 은 번호의
  구멍과 새 글을 고르고, `band_sync` 는 아는 글을 만나면 멈춘다. 그래서 **글은 다
  받았는데 댓글은 한 줄도 없는** 상태가 조용히 남는다. 접수취소는 대부분 댓글로
  오므로(CLAUDE.md '접수했다가 취소되는 건'), 그 상태에서 `cancel_watch` 는
  오류 없이 **반쪽으로** 돈다. 실측 2026-08-08: 8,259 / 8,561 글이 그랬다.

★ 무엇을 고르나 — **날짜가 아니라 쓸모로 고른다** (2026-08-09 지시:
  "무작정 자료 수집만 하지 말고 정확한 알고리즘을 만들어 수집하게 코딩해")

  처음엔 '최근 90일치 250건씩'이었다. 그건 결국 무작정이다 — 7,475건 × 5초 ≈ 10시간을
  긁으면서 그중 무엇이 무엇을 바꾸는지 아무도 모른다.

  기준은 **읽는 쪽의 판정을 그대로 쓴다.** `cancel_watch.build()` 는 취소를 발견해도
  그 프로젝트에 **아직 안 끝난 원장 행**이 있을 때만 대기열 행을 만든다
  (`open_ledger_rows()`). 그러므로 그 집합 밖의 글은 댓글을 다 읽어도
  **오늘 단 한 줄도 못 바꾼다.** 짐작이 아니라 소비자의 코드에서 나온 기준이다.

  실측 2026-08-09 — 댓글 미확인 7,475건의 정체:
      제외: 업무글이 아님(공지·자료·양식)      3,070   ← parse_post 가 None
      제외: 프로젝트NO 없음                   2,641
      [3] 원장에 없거나 이미 닫혔다            1,684
      [1] **열린 원장 행이 있다**                 80   ← 이것만 긁으면 된다
  10시간이 **7분**이 된다.

  갈래
    [1] 열린 원장 행이 있다 — 취소 댓글이 **오늘 숫자를 바꾼다**. **날짜 제한 없음**
        (반년 전 미실시가 그대로 얹혀 있는 것이 바로 이 사고의 본체다)
    [2] 업무글인데 원장에 없다 — 접수 누락일 수 있다. 최근 것부터
    [3] 나머지 업무글 — 결과가 이미 정해졌다. 남는 예산에만
    제외    업무글이 아니다 — 댓글이 바꿀 것이 없다. **긁지 않는다**

  공통 관문(어느 갈래든)
    · `comments` 키가 **아예 없는** 글만. `comments: []`(열어 봤고 없었다)와 가른다 —
      `band_extract.cancel_blind_count` 가 세는 것과 **같은 기준**이다. 세는 쪽과
      고르는 쪽이 갈리면 계기는 줄어드는데 목록은 안 줄고, 또는 그 반대가 된다.
    · 삭제·오염·유령·시각없음은 뺀다. 없는 글을 긁으면 캐시에 쓰레기가 들어간다
      (2026-08-07 사고 — CLAUDE.md "'★밀림'을 보고 없는 번호를 긁지 말 것").
    · 한 배치 250건 — 그 위로는 탭 렌더러가 언다(grab_posts.js 와 같은 값).

사용
  python band/comment_backfill.py                  # 무엇을 왜 긁을지만 (쓰기 없음)
  python band/comment_backfill.py --write          # 붙여넣기 파일까지 만든다
  python band/comment_backfill.py --tier 1         # 1순위만
"""
import argparse
import collections
import json
import io
import os
import sys
from datetime import date, timedelta

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
CACHE_DIR = os.path.join(HERE, "cache")
BATCH_MAX = 250          # grab_posts.js 와 같은 값 — 그 위로는 렌더러가 언다(실측)


def _dead(v):
    """삭제·오염·유령 표시 — 낱말 목록은 `recheck_plan.DEAD_FLAGS` 한 곳이 정한다.

    못 불러오면 **넓게 거른다**(모르면 안 긁는 쪽). 없는 번호를 긁는 것은 한 개당
    21초를 버리는 데 그치지 않고 캐시에 가짜 기록을 남긴다(2026-08-07 사고).
    """
    if not isinstance(v, dict):
        return False
    try:
        if HERE not in sys.path:
            sys.path.insert(0, HERE)
        import recheck_plan as _RP
        return _RP.is_dead(v)
    except Exception:
        return any(v.get(f) for f in
                   ("deleted", "contaminated", "absent", "tainted", "ghost", "dirty"))


def _day(v):
    """캐시의 created_at 은 밀리초 정수·초·ISO 가 섞여 있다 — 한 곳에서 흡수한다."""
    try:
        from datalake import band_day
        return band_day(v) or ""
    except Exception:
        return str(v or "")[:10]


TIER_WHY = {
    1: "열린 원장 행이 있다 — 취소 댓글이 오늘 숫자를 바꾼다",
    2: "업무글인데 원장에 없다 — 접수 누락일 수 있다",
    3: "업무글이지만 원장이 이미 닫혔다 — 남는 예산에만",
}


def open_projects():
    """아직 안 끝난 원장 행이 있는 프로젝트NO 집합.

    ★ **읽는 쪽의 판정을 그대로 빌린다.** `cancel_watch.build()` 는 취소를 발견해도
      이 집합 안에 있을 때만 대기열 행을 만든다. 여기서 따로 '열렸다'를 정의하면
      언젠가 둘이 갈리고, 그때 수집기는 열심히 긁는데 아무것도 안 나온다.
    """
    try:
        sys.path.insert(0, ROOT)
        import cancel_watch
        return set(cancel_watch.open_ledger_rows())
    except Exception as e:
        print("  ! 원장 열린 행을 못 읽었다(%s) — 갈래 없이 최근순으로만 고른다"
              % type(e).__name__)
        return None


def blind(band, days=None, opens=None):
    """[(갈래, 날짜, 글번호)] — 댓글을 한 번도 안 들여다본 글. 갈래 오름차순·최근순.

    `opens` 가 None 이면(원장을 못 읽으면) 전부 갈래 2로 두고 날짜로만 고른다 —
    **모르면서 1순위라고 우기지 않는다.**
    """
    path = os.path.join(CACHE_DIR, "%s.json" % band)
    try:
        d = json.load(io.open(path, encoding="utf-8"))
    except Exception:
        return []
    posts = d.get("posts") or d
    cut = (date.today() - timedelta(days=days)).isoformat() if days else ""
    try:
        sys.path.insert(0, ROOT)
        from band_extract import parse_post
    except Exception:
        parse_post = None
    # ★ 수확이 깨져 있으면 `comments` 키를 **믿지 않는다** (2026-08-09).
    #   250건을 실패 0 으로 긁고도 댓글이 한 건도 안 들어온 적이 있다. 그때 생긴
    #   `comments: []` 를 '읽었다'로 세면 그 글은 영영 다시 안 뽑힌다. 캐시는 고치지
    #   않는다 — 읽을 때만 무시한다. 진짜 댓글이 들어오는 순간 이 조건은 저절로 풀린다.
    #
    # ★ 그런데 distrust 만으로는 **영원히 안 줄어드는 밴드**가 있었다 (2026-08-11 실사고).
    #   90610953 은 열린 원장 1순위 95건을 두 번 재수집·흡수했는데도 1순위가 계속 95건.
    #   원인: 이 밴드에는 **댓글 있는 글이 실제로 하나도 없다**(캐시 5,255글 중 댓글
    #   담긴 글 0). 그러니 `harvest_looks_broken` 이 늘 참을 돌려주고, distrust 는
    #   `comments: []` 인 글을 죄다 '못 읽음'으로 되뽑는다. 아무리 긁어도 진짜 0 이라
    #   distrust 가 안 풀린다 → 같은 95건 무한루프.
    #
    #   가르는 근거는 수집기가 이미 달아 두는 `comments_full` 이다([182] '확인된 0개').
    #   입력창(`_commentInputRegion`)이 그려졌는데 목록이 0 이면 수집기가
    #   `comments_full=True` 로 닫는다 — 그건 못 읽은 게 아니라 **본 것**이다.
    #   반대로 [162] 사고의 깨진 수확은 `comment_count>0` 인데 목록이 0 이라
    #   `comments_full=False` 로 남는다. 그래서 **distrust 여도 comments_full 은 믿는다**:
    #   되뽑는 것은 '비었고 & comments_full 도 아닌' 글뿐이다. 실측 90610953 —
    #   확인된-0개 271건이 후보에서 빠지고, 진짜 못 읽은 4,352건만 남아 수렴한다.
    distrust = bool(harvest_looks_broken(band))
    out = []
    for k, v in posts.items():
        if not str(k).isdigit() or not isinstance(v, dict):
            continue
        # ★ 거를 낱말을 여기 손으로 적지 않는다 (2026-08-11). 캐시가 실제로 다는
        #   표시는 `contaminated` 인데 여기엔 `tainted` 라고 적혀 있어서 **오염을 한
        #   건도 안 걸렀다**(지금은 오염 글에 작성시각이 없어 아래 가드에 걸려 살았을
        #   뿐이다 — 근거가 우연이면 언젠가 무너진다). 낱말은 한 곳에서 온다.
        if _dead(v):
            continue
        if not v.get("created_at"):          # 시각 없는 수확은 믿지 않는다 (검증 [130])
            continue
        # ★ '본 것'의 근거는 **`comments_full`** 이지 `distrust` 가 아니다 (분담판 [39]).
        #   전에는 `distrust` 일 때만 빈 목록을 되뽑았다. 그런데 `harvest_looks_broken`
        #   은 '댓글 담긴 글이 하나라도 있으면' 거짓이라, 실측 두 밴드가 1건·7건으로
        #   거짓이 되어 **`comments: []` 인 5,829건이 통째로 '본 것'으로 넘어갔다.**
        #   계기는 "미확인 0건"이라 말했고 오류도 안 났다 — 검증 [169] 그 자리다.
        #   `comments_full` 없는 빈 목록은 "봤고 없었다"가 아니라 **목록이 다 그려진
        #   것을 확인 못 한 채 0 으로 적힌 것**이다([182] '확인된 0개' 조건 미충족).
        # ★ [199] 는 그대로 지켜진다 — 되뽑는 문을 넓히는 것이 아니라 **근거를
        #   `comments_full` 하나로 좁히는** 것이다. 진짜 0 은 수집기가 `comments_full`
        #   을 달아 닫으므로 다시 안 뽑힌다. 무한루프가 아니라 **한 번 훑으면 수렴**한다
        #   (`comments_full` 이 생기기 전에 받은 글이라 표시가 없을 뿐이다).
        seen = ("comments" in v) and (bool(v.get("comments")) or bool(v.get("comments_full")))
        if seen:
            continue                         # 댓글이 담겼거나, 확인된 0개다
        day = _day(v.get("created_at"))
        tier = 2
        if parse_post is not None:
            try:
                rec = parse_post(int(k), v, band)
            except Exception:
                rec = None
            if rec is None:
                continue                     # 업무글이 아니다 — 댓글이 바꿀 것이 없다
            prj = str(rec.get("프로젝트NO") or "").strip().upper()
            if not prj:
                continue                     # 어느 원장 행에도 못 닿는다
            if opens is not None:
                tier = 1 if prj in opens else 3
        # 1순위는 **날짜로 자르지 않는다** — 반년 전 미실시가 그대로 얹혀 있는 것이
        # 바로 이 사고의 본체다. 2·3순위만 최근 것으로 줄인다.
        if tier != 1 and cut and day < cut:
            continue
        out.append((tier, day, int(k)))
    out.sort(key=lambda t: (t[0], _neg_day(t[1])))
    return out


def _neg_day(day):
    """같은 갈래 안에서는 최근 것이 앞으로 오게 한다(문자열 날짜의 역순 정렬)."""
    return tuple(-ord(c) for c in day)


def harvest_looks_broken(band, floor=30):
    """수확이 '성공'했는데 **댓글이 하나도 없으면** 그건 없는 게 아니라 못 읽은 것이다.

    ★ 2026-08-09 실사고. 250건을 실패 0으로 긁었는데 캐시에 댓글이 **한 건도** 안
      들어왔다. 원인은 `grab_posts.js` 의 개수 선택자
      (`._commentCount, .comment .count, .uComment .count`)가 지금 밴드 화면과 안
      맞는 것이다 — 개수를 0으로 읽으니 댓글이 그려질 때까지 기다리지 않고, 그리기
      전에 읽어 **빈 배열**을 담는다.

      무서운 것은 그다음이다: `comments` 키가 **생기기 때문에** 그 글은 '들여다봤다'로
      세어지고, 사각지대 계기에서도 목록에서도 빠진다. 즉 **못 읽은 글이 읽은 글로
      둔갑해 영영 다시 안 뽑힌다.** 수집은 성공으로 끝나고 아무 데도 티가 안 난다.

      그래서 계기가 스스로 의심한다 — CLAUDE.md "무엇이든 0건이 나오면 묻는다:
      정말 없는 건가, 아니면 안 본 건가."
    """
    path = os.path.join(CACHE_DIR, "%s.json" % band)
    try:
        d = json.load(io.open(path, encoding="utf-8"))
    except Exception:
        return ""
    posts = d.get("posts") or d
    looked = [v for v in posts.values()
              if isinstance(v, dict) and "comments" in v]
    if len(looked) < floor:            # 표본이 적으면 아무 말도 하지 않는다
        return ""
    got = sum(1 for v in looked if v.get("comments"))
    if got:
        return ""
    # ★ '확인된 0개'(comments_full)와 '아직 확인 안 됨'을 갈라 말한다 (2026-08-11).
    #   전에는 뭉뚱그려 "그 글들은 목록에서 빠진다"고 겁을 줬는데, 이제 blind() 는
    #   comments_full 만 신뢰해 빼고 나머지는 재수집 대상으로 남긴다. 그러니 문구도
    #   그대로 말한다 — 겁주는 부분(빠진다)이 사실과 달라지면 사람이 고친 걸 안 고쳤다
    #   여긴다([169] '0건이 나오면 묻는다: 정말 없는 건가 안 본 건가').
    confirmed = sum(1 for v in looked if v.get("comments_full"))
    unconfirmed = len(looked) - confirmed
    return ("들여다봤다고 기록된 %d글 중 **댓글이 있는 글이 0건**입니다"
            "(확인된 0개 %d건 · 아직 확인 안 됨 %d건). "
            "확인된 0개(입력창까지 그려진 뒤 목록이 0)는 신뢰해 목록에서 빼지만, "
            "나머지 %d건은 **수집기가 못 읽은 것**일 수 있어 재수집 대상으로 남깁니다. "
            "이 밴드가 정말 댓글이 없는지, 수집기가 못 읽은 것인지는 그 나머지를 "
            "새 수집기로 다시 긁어야 갈립니다 — 그전엔 **성공으로 읽지 마십시오**."
            % (len(looked), confirmed, unconfirmed, unconfirmed))


def bands():
    return sorted(os.path.splitext(f)[0] for f in os.listdir(CACHE_DIR)
                  if f.endswith(".json") and os.path.splitext(f)[0].isdigit())


# ★ 여기 있던 HEAD/TAIL(제 붙여넣기 구동부)은 2026-08-11 에 **지웠다.**
#   make_oneclick 으로 합친 뒤에는 아무도 안 쓰는데, 남겨 두면 다음 사람이 그것을
#   고치고 파일은 안 바뀌는 일이 생긴다 — 안 읽히는 코드는 빈칸과 같다([165]).
#   내용(탭을 앞에 두라·덤프를 손대지 말라)은 make_oneclick 의 머리말에 이미 있다.


def write_paste(band, nos):
    """붙여넣기 파일을 만드는 자리는 **make_oneclick 하나다**.

    ★ 2026-08-11. 여기엔 제 구동부(위 HEAD/TAIL)가 따로 있었다. 그래서 이 갈래만
    **한 회차 · 250건**이었고, 실측 1,663건이면 사람이 붙여넣기를 **일곱 번** 해야
    했다 — 회차마다 저장하고 탭 메모리를 비우는 다회차 구동부가 make_oneclick 에
    이미 있는데도 그랬다. [162] 의 '담는 쪽이 둘이면 한쪽만 고쳐진다' 가 붙여넣기
    파일에서 그대로 반복된 자리다.
    · 250(`BATCH_MAX`)은 **한 회차** 한도지 한 파일 한도가 아니다. 그 위로 탭이
      어는 것은 맞지만, make_oneclick 은 250씩 끊어 회차 사이에 `__grabSave()` +
      `{keep:false}` 로 비운다 — 메모리는 250건어치로 유지되면서 건수만 이어 간다.
    · 덤으로 [217] 의 죽은 번호 거르기(`screen`)가 이 갈래에도 걸린다. 예전 구동부는
      안 걸러서 삭제·오염·유령 번호를 그대로 붙여넣었다(한 개당 21초에 수확 0).
    """
    # 늦게 부른다 — 모듈 두 개가 서로를 import 하며 도는 것을 막는다.
    try:
        import make_oneclick as mo          # `python band/comment_backfill.py`
    except ImportError:
        from band import make_oneclick as mo  # `from band import comment_backfill`
    js, note = mo.build(band, len(nos), nos=nos, why="댓글 채우기")
    if not js:
        return "(만들 것 없음 — %s)" % note
    p = os.path.join(HERE, "댓글채우기_붙여넣기_%s.js" % band)
    io.open(p, "w", encoding="utf-8").write(js)
    return "%s — %s" % (p, note)


# ── 앱이 읽는 수집 계획 (2026-08-09 지시: Claude Code 없이 앱이 스스로 수집) ────────
#   여기(스케줄된 회차)서 **미리** 골라 둔 것을 앱이 그대로 내려 준다. 브라우저 유저
#   스크립트가 이 계획을 받아 로그인된 밴드 탭에서 스스로 긁으므로 Claude Code 가
#   수집 루프에서 완전히 빠진다. 우선순위 알고리즘은 회차에서 한 번만 돌아(원장 읽기가
#   비싸다) 웹 요청마다 다시 계산하지 않는다.
PLAN_PATH = os.path.join(ROOT, "reports", "밴드_수집계획.json")

# ── PC 가 꺼져 있어도 받게 — 게시용 정적 사본 (2026-08-09 지시: "이 컴퓨터가
#   꺼져있어 연결되어있지 않더라도 앱 자체적으로 처리"). 로컬 앱(localhost)이 없으면
#   유저스크립트가 GitHub Pages 의 이 사본에서 계획·수집기를 받아 스스로 긁는다.
#   비밀 없음: 계획은 **글 번호**뿐, 수집기는 DOM 을 읽는 JS 뿐이다(data.enc 처럼
#   암호화할 것이 없다). 파일 이름은 **ASCII** 로 둔다 — 폰 fetch 가 한글 URL 인코딩에
#   안 걸리게(`plan.json`). 검증 [183].
DOCS_COLLECT = os.path.join(ROOT, "docs", "collect")


def publish_collect(plan, when=None):
    """게시 사본을 만든다: docs/collect/plan.json + grab_posts.js(정본 복사)."""
    import shutil
    os.makedirs(DOCS_COLLECT, exist_ok=True)
    doc = {"bands": plan}
    if when:
        doc["generated"] = when
    tmp = os.path.join(DOCS_COLLECT, "plan.json.tmp")
    io.open(tmp, "w", encoding="utf-8").write(json.dumps(doc, ensure_ascii=False))
    os.replace(tmp, os.path.join(DOCS_COLLECT, "plan.json"))
    # 정본 수집기를 그대로 복사한다 — 앱이 /grab_posts.js 로 내려 주는 것과 **같은 파일**.
    # 규칙이 바뀌면 이 파일만 바뀌고 유저스크립트는 안 바꾼다(붙여넣기 파일과 같은 원칙).
    shutil.copy(os.path.join(HERE, "grab_posts.js"),
                os.path.join(DOCS_COLLECT, "grab_posts.js"))
    return DOCS_COLLECT


def write_plan(plan, when=None):
    """plan = {band: {"nos":[...], "tiers":{"1":n,...}}}. 시각은 밖에서 받는다(테스트 가능)."""
    doc = {"bands": plan}
    if when:
        doc["generated"] = when
    tmp = PLAN_PATH + ".tmp"
    io.open(tmp, "w", encoding="utf-8").write(json.dumps(doc, ensure_ascii=False))
    os.replace(tmp, PLAN_PATH)
    # PC 가 꺼져도 폰/브라우저가 받게 게시 사본까지 — 실패해도 로컬 계획은 남긴다.
    try:
        publish_collect(plan, when)
    except Exception as e:
        print("  ! 게시 사본 실패(%s) — 로컬 계획은 정상" % type(e).__name__)
    return PLAN_PATH


def load_plan(band=None):
    """앱이 부른다. band 를 주면 그 밴드의 nos 만, 없으면 전체 문서."""
    try:
        doc = json.load(io.open(PLAN_PATH, encoding="utf-8"))
    except Exception:
        return ({} if band is None else {"band": band, "nos": []})
    if band is None:
        return doc
    b = (doc.get("bands") or {}).get(str(band)) or {}
    out = {"band": str(band), "nos": list(b.get("nos") or []),
           "tiers": dict(b.get("tiers") or {}), "generated": doc.get("generated")}
    # ★ 브라우저가 읽는 계획은 **한 곳**이어야 한다 (2026-08-20 지시).
    #   실측: 이 파일에는 90610953 의 2건만 있었는데 실제 남은 브라우저 일은 494건이었다
    #   (미수집·UI오염·재수집은 각자 제 파일로만 나갔다). 그래서 사람이 탭을 앞에 둬도
    #   2건만 긁고 끝났다 — 오류도 안 나고 화면도 멀쩡하다([169]).
    #   합치는 자리를 **읽는 쪽 하나**로 둔 이유: 쓰는 쪽을 합치면 09:50 회차의
    #   `--write` 와 대기열 회차가 같은 파일을 서로 덮는다([162]).
    try:
        import collect_queue as _CQ
        q = _CQ.load()
    except Exception:
        q = None
    if not q:
        # 못 읽은 것을 **조용히 넘기지 않는다**([169]) — 받는 쪽이 "이게 전부"로 읽으면 안 된다.
        out["대기열"] = {"상태": "못읽음",
                        "왜": "reports/밴드_수집대기열.json 을 못 읽었다 — 아래 목록은 댓글 갈래뿐이다"}
        return out
    qb = (q.get("bands") or {}).get(str(band)) or {}
    merged, seen = [], set()
    for n in list(qb.get("nos") or []) + out["nos"]:
        n = int(n)
        if n not in seen:
            seen.add(n)
            merged.append(n)
    # ★ 한 배치는 수집기가 정한 한도까지다 — 넘겨 보내면 수집기가 통째로 거절한다.
    #   나머지는 사라지는 것이 아니라 **다음 번 폴링**이 받아 간다(남은 수를 같이 적는다).
    out["nos"] = merged[:BATCH_MAX]
    out["tiers"] = dict(qb.get("tiers") or {}, **out["tiers"])
    out["대기열"] = {"상태": "정상", "전체": len(merged),
                    "남은": max(0, len(merged) - BATCH_MAX),
                    "만든때": q.get("generated"),
                    "건수": qb.get("건수") or {},
                    "갈래설명": q.get("갈래설명") or {},
                    "못읽음": qb.get("못읽음") or []}
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=90,
                    help="2·3순위를 며칠치로 줄일까 (1순위는 안 자른다, 0=전량)")
    ap.add_argument("--limit", type=int, default=BATCH_MAX,
                    help="앱 계획(유저스크립트)이 한 번에 받아 갈 글 수. "
                         "붙여넣기 파일은 남은 것 **전부**를 회차로 나눠 담는다")
    ap.add_argument("--tier", type=int, help="이 갈래까지만 (1=열린 원장 행만)")
    ap.add_argument("--band", help="한 밴드만")
    ap.add_argument("--write", action="store_true", help="붙여넣기 파일까지 만든다")
    a = ap.parse_args(argv)
    if a.limit > BATCH_MAX:
        print("한 배치는 %d건까지입니다(그 위로는 탭이 업니다)." % BATCH_MAX)
        a.limit = BATCH_MAX

    opens = open_projects()
    if opens is not None:
        print("원장에서 아직 안 끝난 프로젝트 %d개 — 이 안에 드는 글만 1순위입니다.\n"
              % len(opens))

    for b in ([a.band] if a.band else bands()):
        warn = harvest_looks_broken(b)
        if warn:
            print("⚠ 밴드 %s — %s\n" % (b, warn))

    grand = collections.Counter()
    plan = {}
    for b in ([a.band] if a.band else bands()):
        rows = blind(b, a.days or None, opens)
        if a.tier:
            rows = [r for r in rows if r[0] <= a.tier]
        by = collections.Counter(t for t, _d, _n in rows)
        for t, n in by.items():
            grand[t] += n
        if not rows:
            continue
        print("밴드 %s" % b)
        for t in sorted(by):
            print("   [%d] %-42s %4d건" % (t, TIER_WHY.get(t, ""), by[t]))
        nos_all = [n for _t, _d, n in rows]     # 남은 것 전부 — 붙여넣기 파일 몫
        nos = nos_all[:a.limit]                  # 배치 한도까지 — 앱 계획 몫
        head = collections.Counter(t for t, _d, _n in rows[:a.limit])
        print("   이번 배치 %d건 — %s"
              % (len(nos), " · ".join("%d순위 %d" % (t, head[t]) for t in sorted(head))))
        # 앱(유저스크립트)이 그대로 받아 쓸 계획 — 우선순위대로, 배치 한도까지.
        plan[str(b)] = {"nos": nos, "tiers": {str(t): head[t] for t in sorted(head)}}
        if a.write:
            # ★ 파일에는 **전부** 담는다(회차로 나뉜다). 배치 한도는 유저스크립트가
            #   한 번에 받아 갈 몫이지, 사람이 붙여넣을 파일의 한도가 아니다 —
            #   한도로 자르면 사람이 같은 일을 일곱 번 하게 된다(write_paste 주석).
            print("   →", write_paste(b, nos_all))

    if a.write:
        # 시각은 여기서 찍는다(모듈 함수는 시험 가능하게 밖에서 받는다).
        from datetime import datetime
        write_plan(plan, when=datetime.now().strftime("%Y-%m-%d %H:%M"))

    if grand:
        print()
        for t in sorted(grand):
            print("전체 [%d] %-42s %5d건" % (t, TIER_WHY.get(t, ""), grand[t]))
        if grand.get(1):
            print("\n★ 1순위 %d건만 긁으면 오늘 바뀔 수 있는 것은 다 봅니다"
                  " (약 %d분)." % (grand[1], max(1, round(grand[1] * 5 / 60))))
        print("업무글이 아닌 글(공지·자료·양식)과 프로젝트NO 없는 글은"
              " **애초에 목록에 없습니다** — 댓글이 바꿀 것이 없기 때문입니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
