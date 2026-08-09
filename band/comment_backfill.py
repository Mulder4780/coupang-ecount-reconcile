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
    distrust = bool(harvest_looks_broken(band))
    out = []
    for k, v in posts.items():
        if not str(k).isdigit() or not isinstance(v, dict):
            continue
        if v.get("deleted") or v.get("tainted") or v.get("absent"):
            continue
        if not v.get("created_at"):          # 시각 없는 수확은 믿지 않는다 (검증 [130])
            continue
        if "comments" in v and not (distrust and not v.get("comments")):
            continue                         # 열어 본 적이 있다 — 비었어도 본 것이다
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
    return ("들여다봤다고 기록된 %d글 중 **댓글이 있는 글이 0건**입니다. "
            "밴드에 댓글이 없어서가 아니라 **수집기가 못 읽은 것**일 수 있습니다"
            "(개수 선택자가 화면과 어긋나면 0으로 읽고 기다리지 않습니다). "
            "그 글들은 '읽음'으로 세어져 목록에서 빠지므로, 고치기 전에는 "
            "이 밴드 댓글 수집을 **성공으로 읽지 마십시오**." % len(looked))


def bands():
    return sorted(os.path.splitext(f)[0] for f in os.listdir(CACHE_DIR)
                  if f.endswith(".json") and os.path.splitext(f)[0].isdigit())


HEAD = """/* ── 밴드 %(band)s **댓글 채우기** — 이 파일 전체를 복사해 밴드 탭 콘솔(F12)에 붙여넣으세요
 *
 *  %(n)d건. 댓글을 한 번도 안 들여다본 글만 골랐습니다(최근 것부터).
 *
 *  ★ 이 탭을 **앞으로 꺼내 놓고** 두세요. 밴드는 보이는 탭에서만 본문을 그립니다 —
 *    뒤에 있으면 전부 실패로 기록됩니다. 중간에 다른 창을 봐도 됩니다(그동안 멈췄다
 *    돌아오면 이어서 갑니다).
 *
 *  끝나면 dump 파일이 자동으로 내려받아집니다. 손대지 마세요 —
 *  download_intake 가 Z: 로 옮기고 convert_dump 가 캐시에 합칩니다.
 */
"""

TAIL = """
(function () {
  var nos = %(nos)s;
  function go() {
    if (typeof window.__grabStart !== 'function') { setTimeout(go, 300); return; }
    if (document.hidden) { setTimeout(go, 1000); return; }   // 앞으로 나올 때까지 기다린다
    console.log(window.__grabStart(%(band)s, nos));
    var t = setInterval(function () {
      var s = window.__grabStatus();
      console.log('진행', s.ok + '/' + s.total, s.paused ? '(탭이 뒤에 있어 멈춤)' : '');
      if (!s.running) { clearInterval(t); console.log(window.__grabSave()); }
    }, 15000);
  }
  go();
})();
"""


def write_paste(band, nos):
    js = io.open(os.path.join(HERE, "grab_posts.js"), encoding="utf-8").read()
    body = (HEAD % {"band": band, "n": len(nos)}) + js + (
        TAIL % {"nos": json.dumps(nos), "band": band})
    p = os.path.join(HERE, "댓글채우기_붙여넣기_%s.js" % band)
    io.open(p, "w", encoding="utf-8").write(body)
    return p


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=90,
                    help="2·3순위를 며칠치로 줄일까 (1순위는 안 자른다, 0=전량)")
    ap.add_argument("--limit", type=int, default=BATCH_MAX, help="한 배치 글 수")
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
        nos = [n for _t, _d, n in rows][:a.limit]
        head = collections.Counter(t for t, _d, _n in rows[:a.limit])
        print("   이번 배치 %d건 — %s"
              % (len(nos), " · ".join("%d순위 %d" % (t, head[t]) for t in sorted(head))))
        if a.write:
            print("   →", write_paste(b, nos))

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
