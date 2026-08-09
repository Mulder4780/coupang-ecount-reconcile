# -*- coding: utf-8 -*-
"""
comment_backfill.py — **댓글을 한 번도 안 들여다본 글**을 골라 수집거리로 만든다

왜 따로 있나
  수집기가 챙기는 것은 늘 '없는 글을 채우는' 방향이었다 — `recheck_plan` 은 번호의
  구멍과 새 글을 고르고, `band_sync` 는 아는 글을 만나면 멈춘다. 그래서 **글은 다
  받았는데 댓글은 한 줄도 없는** 상태가 조용히 남는다. 접수취소는 대부분 댓글로
  오므로(CLAUDE.md '접수했다가 취소되는 건'), 그 상태에서 `cancel_watch` 는
  오류 없이 **반쪽으로** 돈다. 실측 2026-08-08: 8,259 / 8,561 글이 그랬다.

무엇을 고르나 (고르는 규칙 한 곳)
  · `comments` 키가 **아예 없는** 글 — 한 번도 안 열어 본 글이다.
    `comments: []`(열어 봤고 없었다)와는 가른다. `band_extract.cancel_blind_count`
    가 세는 것과 **같은 기준**이다 — 세는 쪽과 고르는 쪽이 갈리면 계기는 줄어드는데
    목록은 안 줄어든다(또는 그 반대).
  · 삭제·오염·유령·시각없음은 뺀다. 없는 글을 긁으면 캐시에 쓰레기가 들어간다
    (2026-08-07 사고 — CLAUDE.md "'★밀림'을 보고 없는 번호를 긁지 말 것").
  · **최근 것부터.** 취소가 뜻을 갖는 것은 아직 안 끝난 일이다. 오래된 글의 댓글은
    이미 결과가 원장에 반영돼 있어 지금 와서 바꿀 것이 없다.

  ★ 전량은 한 번에 못 한다 — 7,475건 × 5초 ≈ 10시간이다. `--days` 로 창을 좁히고
    `--limit`(한 배치 250건, 그 위로는 탭이 언다)로 끊는다. 남은 것은 다음 회차 몫이다.

사용
  python band/comment_backfill.py                      # 어디가 비었나만 (쓰기 없음)
  python band/comment_backfill.py --days 90 --write    # 붙여넣기 파일까지 만든다
"""
import argparse
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


def blind(band, days=None):
    """[(날짜, 글번호)] — 댓글을 한 번도 안 들여다본 글, 최근 것부터."""
    path = os.path.join(CACHE_DIR, "%s.json" % band)
    try:
        d = json.load(io.open(path, encoding="utf-8"))
    except Exception:
        return []
    posts = d.get("posts") or d
    cut = (date.today() - timedelta(days=days)).isoformat() if days else ""
    out = []
    for k, v in posts.items():
        if not str(k).isdigit() or not isinstance(v, dict):
            continue
        if v.get("deleted") or v.get("tainted") or v.get("absent"):
            continue
        if not v.get("created_at"):          # 시각 없는 수확은 믿지 않는다 (검증 [130])
            continue
        if "comments" in v:                  # 열어 본 적이 있다 — 비었어도 본 것이다
            continue
        day = _day(v.get("created_at"))
        if cut and day < cut:
            continue
        out.append((day, int(k)))
    out.sort(reverse=True)
    return out


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
    ap.add_argument("--days", type=int, default=90, help="며칠치까지 볼까 (0=전량)")
    ap.add_argument("--limit", type=int, default=BATCH_MAX, help="한 배치 글 수")
    ap.add_argument("--band", help="한 밴드만")
    ap.add_argument("--write", action="store_true", help="붙여넣기 파일까지 만든다")
    a = ap.parse_args(argv)
    if a.limit > BATCH_MAX:
        print("한 배치는 %d건까지입니다(그 위로는 탭이 업니다)." % BATCH_MAX)
        a.limit = BATCH_MAX

    total_all = 0
    for b in ([a.band] if a.band else bands()):
        allb = blind(b, None)
        win = blind(b, a.days or None)
        total_all += len(allb)
        print("밴드 %s — 댓글 미확인 전체 %d건 · 최근 %s일 %d건"
              % (b, len(allb), a.days or "전", len(win)))
        if not win:
            continue
        nos = [n for _d, n in win][:a.limit]
        print("   이번 배치 %d건 (%d~%d)" % (len(nos), max(nos), min(nos)))
        if a.write:
            print("   →", write_paste(b, nos))
    if total_all:
        print("\n전체 %d건이 아직 댓글을 한 번도 안 들여다봤습니다." % total_all)
        print("접수취소는 대부분 댓글로 옵니다 — 그동안 cancel_watch 는 반쪽으로 돕니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
