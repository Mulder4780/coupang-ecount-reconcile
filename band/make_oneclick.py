# -*- coding: utf-8 -*-
"""make_oneclick.py — 사람이 로그인해 둔 밴드 탭에 **한 번 붙여넣으면 끝**나는 수집 파일 생성

왜 필요한가 (2026-08-06)
  밴드 로그인은 사람만 할 수 있다(절대규칙 3). 그런데 AI 가 조종하는 탭은 확장 프로그램이
  설치된 **크롬 프로필** 안에 있어서, 사람이 다른 프로필 창에서 로그인하면 AI 탭은 계속
  로그아웃으로 보인다. 실제로 이것 때문에 하루가 막혔다.
  → 프로필을 맞추게 하는 대신, **이미 로그인된 그 탭에서** 돌릴 수 있는 파일을 만든다.
    사람이 하는 일은 '열기 → 전체복사 → 콘솔에 붙여넣기' 셋뿐이고, 그다음은 전자동이다.

무엇이 들어가나
  grab_posts.js 전문 + 이번에 훑을 번호 배열 + 끝나면 자동 저장(__grabSave).
  번호는 recheck_plan 이 정한 순서 그대로다 — **새 글 먼저**, 그다음 최근 구멍부터.
  저장된 dump_*.json 은 download_intake --apply 가 Z: 로 흡수하고 convert_dump 가 캐시로 합친다.

  python band/make_oneclick.py                 # 밴드마다 한 파일씩
  python band/make_oneclick.py --band 90610953 --limit 250
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import recheck_plan as RP

# 한 배치 상한은 grab_posts.js 와 같아야 한다(250을 넘기면 탭 렌더러가 언다 — 실측).
BATCH_MAX = 250

# 없는 번호 하나가 무는 시간(초) — grab_posts.js 의 waitMs 9000 + bodyMs 12000.
# 여기 적어 두는 이유는 **사람에게 비용을 말해 주기 위해서**다. "40건"은 숫자일 뿐이지만
# "약 14분 쓰고 수확 0" 은 사람이 판단할 수 있는 말이다.
MISS_SEC = 21


def screen(band, nos, posts=None):
    """붙여넣기 파일에 **없는 번호가 들어가지 않게** 하는 마지막 문 (2026-08-11).

    왜 여기인가: 번호를 정하는 곳은 여럿이다(recheck_plan · comment_plan · recollect).
    거르는 자리를 각자 두면 한 곳만 고쳐지고, 안 고쳐진 쪽은 **없는 번호를 사람 손에
    들려 보낸다** — 밴드는 없는 번호에도 200 과 앱 껍데기를 주므로 수집기는 21초를
    꽉 채우고 no-time 으로 버린다(검증 [130]). 오류도 안 나고 '실패'로도 안 세인다.
    그래서 **파일로 나가는 길목 하나**에서 거른다.

    거르는 근거는 둘 다 이미 있는 것이다 — 새로 지어내지 않는다:
      · `reports/밴드_확인시각.json` 의 '없음 확인' 구간(살아 있는 근거일 때만)
      · 캐시가 달아 둔 삭제·오염·유령 표시(`recheck_plan.DEAD_FLAGS`)
    """
    if posts is None:
        posts = RP.load(band) or {}
    cut, why = RP.absent_line(band, posts)
    keep, dropped = [], {}
    for n in nos:
        n = int(n)
        if cut is not None and n >= cut:
            dropped.setdefault("없음 확인 구간(%s↑)" % cut, []).append(n)
            continue
        if RP.is_dead(posts.get(str(n))):
            dropped.setdefault("삭제·오염·유령 표시", []).append(n)
            continue
        keep.append(n)
    return keep, dropped, why


def build(band, limit, nos=None, why=""):
    """남은 것을 **250씩 여러 회차로** 끊어 한 파일에 담는다.

    `nos` 를 주면 recheck_plan 대신 **그 번호들을** 훑는다 — 재수집 회차
    (`band/recollect.py`)가 최근 30일 창을 넘길 때 쓴다. 붙여넣기 파일을 두 벌
    만들지 않기 위해 갈래를 여기 하나로 둔다(회차 JS·배치 상한·저장 규칙이
    갈리면 한쪽만 고쳐지고 다른 쪽은 조용히 옛날 것으로 남는다).

    왜 한 파일에 다 넣나 (2026-08-06 지시: "내 손 안 가게")
      250 은 탭이 얼지 않는 한 배치의 상한이지, 사람이 붙여넣어야 하는 횟수가 아니다.
      회차 사이에 `__grabSave()` 로 내보내고 `{keep:false}` 로 탭 메모리를 비우면
      메모리는 250건어치로 유지되면서 회차만 계속 이어 갈 수 있다.
      그래서 **붙여넣기는 밴드당 한 번**이면 끝난다(예전엔 250건마다 한 번이었다).
    """
    posts = RP.load(band)
    if posts is None:
        return None, "캐시 없음"
    head = ""
    if nos is None:
        sc = RP.scope()
        floor = int((sc.get("floor") or {}).get(str(band), 0) or 0)
        ahead = int(sc.get("ahead") or 40)
        p = RP.plan(str(band), posts, floor, ahead)
        if not p:
            # 계획을 못 세우는 밴드가 있다 — 유령 밴드의 **빈 캐시**가 그렇다.
            # 예전에는 여기서 TypeError 로 죽어 **뒤에 있는 진짜 밴드까지 전부**
            # 붙여넣기 파일을 못 만들었다(2026-08-08). 한 밴드가 나머지를 죽이면 안 된다.
            return None, "계획을 세울 수 없다(빈 캐시일 수 있다)"
        raw = p["new"] + sorted(p["gaps"], reverse=True) + sorted(p["stale"], reverse=True)
        # ★ 위쪽(아직 없을 수도 있는 번호)을 몇 개 넣었는지·왜 그런지 사람에게 말한다.
        #   "40건" 은 숫자일 뿐이고 "근거가 낡아 5건만 찔러 본다" 는 판단할 수 있는 말이다.
        head = ("위쪽 %d건(%s)" % (len(p["new"]), p.get("absent_why") or "")
                if p.get("new") else "위쪽 0건(%s)" % (p.get("absent_why") or ""))
    else:
        raw = [int(n) for n in nos]
    # ★ 어느 길로 왔든 **같은 문**을 지난다 — 번호를 정하는 곳은 셋인데 파일로 나가는
    #   길은 여기 하나다(recheck_plan 이 이미 걸러도 comment_plan·recollect 는 제 목록을 준다).
    todo, dropped, absent_why = screen(band, raw, posts)
    todo = todo[:limit]
    if not todo:
        return None, ("훑을 것이 없다" if not dropped else
                      "훑을 것이 없다(없는 번호 %d건 제외 — %s)"
                      % (sum(len(v) for v in dropped.values()),
                         " · ".join("%s %d건" % (k, len(v)) for k, v in dropped.items())))
    rounds = [todo[i:i + BATCH_MAX] for i in range(0, len(todo), BATCH_MAX)]
    body = open(os.path.join(HERE, "grab_posts.js"), encoding="utf-8").read()
    js = f"""/* ── 밴드 {band} {why or '수집'} — 이 파일 전체를 복사해 밴드 탭 콘솔(F12)에 붙여넣으세요 ──
 * 붙여넣는 즉시 시작하고 **끝까지 알아서 갑니다**. 붙여넣기는 이 한 번뿐입니다.
 *   {len(todo)}건 · {len(rounds)}회차(회차당 최대 {BATCH_MAX}건) · 대략 {len(todo) * 5 // 60}분.
 *   위쪽 근거: {absent_why}. 아직 없는 번호는 한 개당 {MISS_SEC}초를 무는데 수확이
 *   없으므로, 근거가 없을 때는 **존재 확인용으로 몇 개만** 넣습니다(있는 것이 확인되면
 *   다음 회차가 이어받습니다). 있는 글로 확인된 구간은 예전처럼 끝까지 훑습니다.
 * 회차가 끝날 때마다 dump 파일이 **자동으로 다운로드**되고 탭 메모리는 비워집니다
 * (그래야 얼지 않습니다). 다운로드된 파일은 손대지 마세요 — download_intake 가
 * Z: 로 옮기고 convert_dump 가 캐시로 합칩니다.
 * 진행 중 탭을 새로고침하면 그 회차분이 날아갑니다. 그냥 두면 됩니다(뒤에 있어도 돕니다).
 */
{body}

(function () {{
  const ROUNDS = {json.dumps(rounds)};
  const BAND = {band};
  const say = (m) => console.log('%c' + m, 'color:#0b5cff;font-weight:700');
  // 회차가 끝나기를 기다린다 — 폴링은 setInterval 이면 충분하다(수집 자체는
  // grab_posts.js 안의 Worker 타이머가 돌리므로 숨은 탭에서도 늦춰지지 않는다).
  const waitRound = (i) => new Promise((res) => {{
    const t = setInterval(() => {{
      const s = window.__grabStatus();
      console.log(`밴드 ${{BAND}} [${{i + 1}}/${{ROUNDS.length}}] ${{s.ok}}/${{s.total}} 수집 · 없는글 ${{s.missing}} · 실패 ${{s.failed}} · ${{s.sec}}초`);
      if (!s.running) {{ clearInterval(t); res(); }}
    }}, 5000);
  }});
  (async () => {{
    // 앞 회차가 아직 돌고 있으면 그것부터 끝내고 저장한다. 그냥 시작하면
    // __grabStart 가 '이미 실행 중'으로 거절해 첫 회차 번호가 통째로 빠진다.
    if (window.__grabStatus().running) {{
      say('앞 배치가 아직 돌고 있습니다 — 그것부터 끝내고 이어 갑니다.');
      await waitRound(-1);
      say('앞 배치 ' + window.__grabSave());
    }}
    for (let i = 0; i < ROUNDS.length; i++) {{
      // keep:false → 지난 회차 글을 탭에서 비운다. 이미 저장했으니 잃는 것이 없다.
      console.log(window.__grabStart(BAND, ROUNDS[i], {{ keep: false }}));
      await waitRound(i);
      say(`[${{i + 1}}/${{ROUNDS.length}}] ` + window.__grabSave());
    }}
    say('밴드 ' + BAND + ' — 전체 ' + ROUNDS.length + '회차 끝. 다운로드 파일은 그대로 두세요.');
  }})();
}})();
"""
    note = f"{len(todo)}건 · {len(rounds)}회차 ({todo[0]}~{todo[-1]})"
    if head:
        note += f" · {head}"
    if dropped:
        # ★ 거른 것을 **말없이 삼키지 않는다.** 조용히 빼면 "왜 그 번호가 안 들어왔나"를
        #   다음 사람이 다시 조사한다(그리고 대개 다시 넣는다).
        note += " · 제외 " + " · ".join("%s %d건" % (k, len(v)) for k, v in dropped.items())
    return js, note


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--band", help="특정 밴드만")
    ap.add_argument("--limit", type=int, default=10 ** 6,
                    help="총 상한(회차로 쪼갠다). 기본은 남은 것 전부")
    ap.add_argument("--out", default=HERE, help="파일을 만들 폴더")
    a = ap.parse_args()
    bands = [a.band] if a.band else sorted(
        f[:-5] for f in os.listdir(RP.CACHE)
        if f.endswith(".json") and f[:-5].isdigit())
    made = 0
    for band in bands:
        # ★ 한 밴드에서 무슨 일이 나든 **나머지는 계속 만든다.** 실측 2026-08-08:
        #   유령 밴드 하나가 TypeError 로 죽어 진짜 밴드 두 개의 파일이 안 만들어졌고,
        #   사람 손에 가는 것은 디스크에 있는 그 파일이라 **아무것도 못 긁었다.**
        try:
            js, note = build(band, a.limit)
        except Exception as e:
            print(f"밴드 {band}: 만들지 못했다 — {type(e).__name__}: {e}")
            continue
        if not js:
            print(f"밴드 {band}: {note}")
            continue
        path = os.path.join(a.out, f"수집_붙여넣기_{band}.js")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(js)
        made += 1
        print(f"밴드 {band}: {note} → {path}")
    print(f"\n만든 파일 {made}개 / 밴드 {len(bands)}개")
    print("쓰는 법: 파일 열기 → 전체 복사 → 로그인된 밴드 탭에서 F12 → Console 에 붙여넣기 → Enter")
    return 0


if __name__ == "__main__":
    sys.exit(main())
