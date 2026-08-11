# -*- coding: utf-8 -*-
"""
comment_plan.py — 댓글을 **한 번도 안 본 글**을 찾아 훑을 계획을 세운다
======================================================================
사용자 지시(2026-08-09): "밴드도 댓글도 다 찾아 저장하는 알고리즘 구성"

## 왜 필요한가 — 고친 것과 채운 것은 다르다
2026-08-08 에 수집기가 댓글을 담도록 고쳤다(`grab_posts.js` → `convert_dump`,
검증 `[162]`). 그런데 **고친 뒤로 새로 긁은 글에만** 댓글이 들어온다.
실측 2026-08-09: 캐시 8,561글 중 **8,258글이 댓글을 한 번도 안 봤고**, 댓글 본문이
담긴 글은 **0건**이다. 그래서 `cancel_watch` 는 지금도 댓글 취소를 **하나도 못 잡는다**
— 코드는 멀쩡히 돌고 오류도 안 난다. 이것이 이 프로젝트가 말하는 조용한 사고다.

## 무엇을 고르나 — '없는 것'과 '보고 없던 것'을 가른다
  · `comments` 키가 **아예 없다** → 한 번도 안 들여다봤다 → **훑을 대상**
  · `comments: []`               → 보긴 봤고 없었다 → 대상 아님
검증 `[169]` 가 세운 구별을 그대로 쓴다. 이 둘을 섞으면 매 회차가 같은 글을 다시 훑는다.

## 순서 — 최근 글부터
취소·변경은 **최근 글**에 붙는다. 8,258글을 다 훑는 데 여러 밤이 걸리므로, 도중에
멈춰도 값어치가 남게 번호가 큰 것부터 간다. 오래된 글의 댓글은 이미 업무가 끝나
바뀔 일이 없다.

## 긁는 것은 여전히 사람 몫이다 (절대규칙 3 · 수집 세션)
이 도구는 **계획만** 세우고 붙여넣기 파일을 만든다. 실제 수집은
'CSOS 리서치 및 자료 수집' 세션이 로그인된 탭에서 한 번 붙여넣으면 이어진다.

실행:  python band/comment_plan.py              # 현황만
       python band/comment_plan.py --make       # 붙여넣기 파일까지 만든다
       python band/comment_plan.py --json
"""
import sys, os, json, glob, argparse
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CACHE = os.path.join(HERE, "cache")
# 거를 낱말·'없음 확인' 근거는 **한 곳**에서 온다 — 여기서 따로 적으면 낱말이 어긋나
# 한쪽만 거른다(이 파일이 2026-08-11 에 바로 그렇게 다쳤다).
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import recheck_plan as RP
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
OUT_JSON = os.path.join(REPORT_DIR, "밴드_댓글계획.json")

BATCH = 250          # 한 회차 상한(탭이 얼지 않는 선) — make_oneclick 과 같은 값


def bands():
    """캐시에 있는 **진짜 밴드**만.

    ★ 밴드번호는 **8자리**다. 날짜에서 온 이름(`260807` 6자리, `202608082047` 12자리)은
      유령이다 — 2026-08-08 에 `260807` 유령 하나가 두 밴드를 합쳐 5,453글을 만들었고,
      재수집 회차가 그 유령에 붙여넣기 파일을 만들고 나서야 드러났다.
      첫판에서 `len >= 7` 로 뒀더니 `202608082047` 이 그대로 들어와 헛 계획을 세웠다.
      **자릿수로 거른다** — 넓게 잡으면 유령이 늘 끼어든다.
    """
    return [b for b in (os.path.basename(f)[:-5]
                        for f in sorted(glob.glob(os.path.join(CACHE, "*.json"))))
            if b.isdigit() and len(b) == 8]


def load(band):
    with open(os.path.join(CACHE, f"{band}.json"), encoding="utf-8") as fh:
        return json.load(fh)


def pick(posts, band=None):
    """`(번호들, 제외사유)` — 댓글을 **한 번도 안 본** 글, 최근(번호 큰) 것부터.

    ★ `comments: []` 는 고르지 않는다. 보고 없었던 글까지 다시 훑으면 8,258 이
      영원히 안 줄고, 매 회차가 같은 자리를 맴돈다.

    ★ **없는 글을 목록에 넣지 않는다** (2026-08-11 실사고). 예전 조건은
      `deleted / ghost / dirty` 였는데 캐시가 실제로 다는 표시는 `contaminated`(오염)·
      `absent`(유령)다 — **낱말이 어긋나 한 건도 안 걸렀다.** 실측으로 남아 있던 계획
      84789192 **102건 전부** · 90610953 **522건 전부**가 그 표시가 달린 번호였다.
      그대로 긁으면 밴드가 200 과 앱 껍데기를 주므로 한 개당 21초를 채우고 시각이 없어
      버려진다 — 합쳐서 **약 3시간 반에 수확 0**이다. 이제 거르는 낱말은
      `recheck_plan.DEAD_FLAGS` **한 곳**에서 온다.

    ★ '없음 확인' 근거(`reports/밴드_확인시각.json`)가 **살아 있을 때만** 그 구간을 뺀다.
      낡은 근거로 빼면 그 사이 올라온 진짜 글을 영영 안 보게 된다 — 근거 없이
      '조용함'을 단정하지 않는다(검증 [131] 과 같은 문).

    ★ 작성시각이 없는 글은 **빼지 않는다.** 그건 '안 본 글'이 맞고([179]),
      그 수확을 다시 받는 일은 recheck_plan 몫이다 — 여기서 같이 빼면 두 회차가
      서로 미루다 아무도 안 본다.
    """
    nos, why = [], {}
    cut = RP.absent_line(band, posts)[0] if band else None
    for no, p in posts.items():
        if not isinstance(p, dict) or not str(no).isdigit():
            continue
        n = int(no)
        if RP.is_dead(p):
            why["삭제·오염·유령 표시"] = why.get("삭제·오염·유령 표시", 0) + 1
            continue                          # 삭제·유령·오염은 업무 기록이 아니다
        if cut is not None and n >= cut:
            why["없음 확인 구간"] = why.get("없음 확인 구간", 0) + 1
            continue
        if "comments" in p:
            continue                          # 이미 들여다봤다
        nos.append(n)
    return sorted(nos, reverse=True), why


def unlooked(posts, band=None):
    """번호들만 — 고르는 판단은 `pick()` 한 곳이다(옛 호출자·검증이 쓰는 모양)."""
    return pick(posts, band)[0]


def plan(make=False, limit=None):
    out, made = [], []
    for band in bands():
        d = load(band)
        posts = d.get("posts") or {}
        nos, skipped = pick(posts, band)
        looked = sum(1 for p in posts.values() if isinstance(p, dict) and "comments" in p)
        got = sum(1 for p in posts.values() if isinstance(p, dict) and p.get("comments"))
        item = {"밴드": band, "이름": d.get("band_name") or band,
                "글": len(posts), "댓글_안본": len(nos), "댓글_본": looked,
                "댓글_있음": got,
                # ★ 뺀 것을 **숫자로 남긴다.** 조용히 빼면 '0건'이 '다 봤다'로 읽힌다
                #   — 그건 안 본 것이다(검증 [169]). 무엇을 왜 뺐는지가 같이 있어야 한다.
                "제외": skipped,
                "회차수": (len(nos) + BATCH - 1) // BATCH,
                "다음배치_처음": nos[0] if nos else None,
                "다음배치_끝": nos[min(len(nos), BATCH) - 1] if nos else None}
        out.append(item)
        if make and nos:
            take = nos[:limit] if limit else nos
            try:
                import make_oneclick as MO
                js, note = MO.build(band, len(take), nos=take,
                                    why="댓글을 한 번도 안 본 글")
            except Exception as e:
                item["오류"] = f"{type(e).__name__}: {e}"
                continue
            if js:
                path = os.path.join(HERE, f"댓글_붙여넣기_{band}.js")
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(js)
                item["파일"] = os.path.basename(path)
                made.append(path)
    os.makedirs(REPORT_DIR, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as fh:
        json.dump({"생성": datetime.now().isoformat(timespec="seconds"),
                   "한회차": BATCH, "밴드별": out}, fh, ensure_ascii=False, indent=1)
    return out, made


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--make", action="store_true", help="붙여넣기 파일까지 만든다")
    ap.add_argument("--limit", type=int, help="밴드당 상한(시험용)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    rows, made = plan(a.make, a.limit)
    if a.json:
        print(json.dumps(rows, ensure_ascii=False))
        return 0
    tot = sum(r["댓글_안본"] for r in rows)
    for r in rows:
        print(f"밴드 {r['밴드']} — 글 {r['글']} · 댓글 안 본 글 **{r['댓글_안본']}** "
              f"({r['회차수']}회차) · 이미 본 {r['댓글_본']} · 댓글 있는 글 {r['댓글_있음']}"
              + (f" → {r['파일']}" if r.get("파일") else "")
              + (f"  ⚠ {r['오류']}" if r.get("오류") else ""))
        if r.get("제외"):
            # '안 본 글 0건'이 '다 봤다'로 읽히지 않게 — 무엇을 왜 뺐는지 같이 말한다.
            print("   · 목록에서 뺀 것: "
                  + " · ".join(f"{k} {v}건" for k, v in sorted(r["제외"].items())))
    print(f"합계 댓글 안 본 글 {tot}건 — 최근 글부터 훑습니다(도중에 멈춰도 값어치가 남게)")
    if made:
        print("긁는 것은 수집 세션 몫입니다. 로그인된 밴드 탭 콘솔에 한 번 붙여넣으면 이어집니다:")
        for p in made:
            print("  ", p)
    return 0


if __name__ == "__main__":
    sys.path.insert(0, HERE)
    sys.exit(main())
