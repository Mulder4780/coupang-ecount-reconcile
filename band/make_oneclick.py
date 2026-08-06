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


def build(band, limit):
    posts = RP.load(band)
    if posts is None:
        return None, "캐시 없음"
    sc = RP.scope()
    floor = int((sc.get("floor") or {}).get(str(band), 0) or 0)
    ahead = int(sc.get("ahead") or 40)
    p = RP.plan(str(band), posts, floor, ahead)
    todo = (p["new"] + sorted(p["gaps"], reverse=True)
            + sorted(p["stale"], reverse=True))[:min(limit, BATCH_MAX)]
    if not todo:
        return None, "훑을 것이 없다"
    body = open(os.path.join(HERE, "grab_posts.js"), encoding="utf-8").read()
    js = f"""/* ── 밴드 {band} 수집 — 이 파일 전체를 복사해 밴드 탭 콘솔(F12)에 붙여넣으세요 ──
 * 붙여넣는 즉시 시작합니다. {len(todo)}건 · 글당 5초 안팎이라 약 {len(todo) * 5 // 60}분.
 * 끝나면 dump 파일이 **자동으로 다운로드**됩니다. 그 파일은 손대지 마세요 —
 * download_intake 가 알아서 Z: 로 옮기고 캐시까지 합칩니다.
 * 진행 중에 탭을 새로고침하면 모은 것이 날아갑니다. 그냥 두면 됩니다(뒤에 있어도 돕니다).
 */
{body}

(function () {{
  const NOS = {json.dumps(todo)};
  console.log(window.__grabStart({band}, NOS));
  // 끝나는 순간 저장까지 자동으로 — 사람이 지켜보고 있을 필요가 없다.
  const t = setInterval(() => {{
    const s = window.__grabStatus();
    console.log(`밴드 {band} — ${{s.ok}}/${{s.total}} 수집 · 없는글 ${{s.missing}} · 실패 ${{s.failed}} · ${{s.sec}}초`);
    if (!s.running) {{
      clearInterval(t);
      console.log('%c' + window.__grabSave(), 'color:#0b5cff;font-weight:700');
      console.log('%c끝났습니다. 다운로드된 dump 파일은 그대로 두세요.', 'color:#0b5cff;font-weight:700');
    }}
  }}, 5000);
}})();
"""
    return js, f"{len(todo)}건 ({todo[0]}~{todo[-1]})"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--band", help="특정 밴드만")
    ap.add_argument("--limit", type=int, default=BATCH_MAX)
    ap.add_argument("--out", default=HERE, help="파일을 만들 폴더")
    a = ap.parse_args()
    bands = [a.band] if a.band else sorted(
        f[:-5] for f in os.listdir(RP.CACHE)
        if f.endswith(".json") and f[:-5].isdigit())
    for band in bands:
        js, note = build(band, a.limit)
        if not js:
            print(f"밴드 {band}: {note}")
            continue
        path = os.path.join(a.out, f"수집_붙여넣기_{band}.js")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(js)
        print(f"밴드 {band}: {note} → {path}")
    print("\n쓰는 법: 파일 열기 → 전체 복사 → 로그인된 밴드 탭에서 F12 → Console 에 붙여넣기 → Enter")
    return 0


if __name__ == "__main__":
    sys.exit(main())
