# -*- coding: utf-8 -*-
"""
ux_review.py — 담당자 업무센터 UI/UX 개선점을 **사용 기록에서** 찾아 낸다 (3일마다 12:00)

사용자 지시(2026-08-05): "담당자 업무센터 UI UX도 개선이 필요한 사항이 있으면 사용자가
편리하게 개선작업 진행해 (3일에 한번 12:00 에 실행)."

왜 도구로 만드나: 화면 개선을 감으로 하면 안 쓰는 곳만 고치게 된다. 이 프로젝트에는
이미 `ledger_db.ux_add()` 로 **실제 사용 기록**(화면 이동·클릭·검색 무결과·오류·느린 화면)이
쌓이고 있다. 그 기록에서 "사람이 막히는 지점"을 뽑아 개선 후보로 적어 둔다.
고치는 것은 사람/AI 가 판단한다 — 이 도구는 화면을 자동으로 바꾸지 않는다(읽기 전용).

무엇을 보나
  1. **오류가 반복되는 API/화면** — 가장 먼저 고칠 것(사용자는 원인을 모른 채 실패만 본다)
  2. **검색 무결과** — 사람이 찾는데 없는 말. 색인·라벨을 그 말에 맞춰야 한다
  3. **자주 여는 화면인데 클릭이 많은 곳** — 깊이 숨은 기능은 위로 올린다
  4. **거의 안 쓰는 화면** — 담당자 센터 기본 구성에서 내려도 되는 후보
  5. **느린 화면** — 응답이 느린 곳(사람은 '멈췄다'고 느낀다)

산출: reports/UX개선_제안.md  (+ 직전 회차와 비교해 새로 생긴 문제를 ★로 표시)
  python ux_review.py            # 분석·리포트
  python ux_review.py --print    # 콘솔 요약만
"""
import json
import os
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT_MD = os.path.join(ROOT, "reports", "UX개선_제안.md")
OUT_JSON = os.path.join(ROOT, "reports", "UX개선_제안.json")
# 담당자 업무센터에서 쓰는 화면 — 여기 것이 우선순위가 높다.
STAFF_VIEWS = {"center", "ryu", "pm", "as", "worklog", "settle", "check", "sources", "remote"}
VIEW_LABEL = {
    "center": "업무센터", "ryu": "류지영 기록", "pm": "정기점검", "as": "돌발AS",
    "worklog": "대표보고 일지", "settle": "정산", "check": "확인 필요",
    "sources": "원본 자료", "remote": "리모컨", "dash": "대시보드",
    "daily": "대표 보고", "run": "실행", "report": "기록", "calendar": "일정",
}


def label(v):
    return VIEW_LABEL.get(v, v)


def prev_report():
    try:
        return json.load(open(OUT_JSON, encoding="utf-8"))
    except Exception:
        return {}


def main():
    import ledger_db
    s = ledger_db.ux_summary()
    views = dict(s.get("화면별") or [])
    clicks = dict(s.get("많이누른것") or [])
    errors = [(a, b, c) for a, b, c in (s.get("오류") or [])]
    empty = dict((s.get("빈손검색") or []))
    # 느린화면은 (대상, 합계ms, 횟수) 3원소다 — 평균으로 바꿔 본다.
    slow = {}
    for row in (s.get("느린화면") or []):
        try:
            tgt, total_ms, cnt = row[0], float(row[1]), max(1, int(row[2]))
            slow[tgt] = int(total_ms / cnt)
        except Exception:
            continue

    prev = prev_report()
    prev_keys = set(prev.get("keys") or [])
    items, keys = [], []

    def add(kind, target, why, how, weight):
        key = f"{kind}|{target}"
        keys.append(key)
        items.append({"kind": kind, "target": target, "why": why, "how": how,
                      "weight": weight, "new": key not in prev_keys})

    # 1) 오류 — 사람이 원인을 모른 채 실패만 본다. 가장 먼저.
    for where, what, n in errors[:8]:
        if n < 3:
            continue
        add("오류", str(where)[:60], f"최근 7일 {n}회 실패",
            "실패해도 화면이 '왜 안 되는지'를 한 줄로 말하게 하고, 재시도 버튼을 둔다", n * 3)

    # 2) 검색 무결과 — 사람이 쓰는 말과 시스템의 말이 다른 지점
    for term, n in list(empty.items())[:8]:
        if n < 2:
            continue
        add("검색 무결과", str(term)[:40], f"{n}회 찾았는데 결과 없음",
            "이 말을 별칭으로 색인에 넣거나, 결과 없을 때 '이렇게 찾아 보세요'를 보여 준다", n * 2)

    # 3) 느린 화면
    for view, ms in list(slow.items())[:6]:
        try:
            ms = int(ms)
        except Exception:
            continue
        if ms < 1500:
            continue
        add("느림", label(view), f"평균 {ms}ms",
            "먼저 뼈대를 그리고 자료는 나중에 채운다(빈 화면에 '불러오는 중'만 두지 않는다)", ms // 100)

    # 4) 자주 여는데 클릭이 많은 화면 — 깊이 숨은 기능
    for view, opens in list(views.items())[:12]:
        if view not in STAFF_VIEWS or opens < 20:
            continue
        c = clicks.get(view, 0)
        if c > opens * 0.8:
            add("동선", label(view), f"{opens}번 열고 {c}번 누름 — 한 번에 못 끝낸다",
                "가장 많이 누르는 것을 카드 맨 위로 올리고, 두 번 이상 눌러야 닿는 기능은 바로가기를 만든다",
                c)

    # 5) 거의 안 쓰는 담당자 화면 — 기본 구성에서 내릴 후보
    for view in STAFF_VIEWS:
        n = views.get(view, 0)
        if n <= 3:
            add("미사용", label(view), f"최근 7일 {n}회 — 거의 안 열린다",
                "담당자 센터 기본 구성에서 내리고(숨김), 필요한 사람만 '내 화면 구성'에서 켜게 한다",
                5)

    items.sort(key=lambda x: (-x["weight"]))
    new_cnt = sum(1 for x in items if x["new"])

    L = [f"# 업무센터 UX 개선 제안 (자동 분석 {time.strftime('%Y-%m-%d %H:%M')})", "",
         f"- 기준: {s.get('기간', '최근 7일')} 사용 기록",
         f"- 제안 **{len(items)}건** (이번에 새로 생긴 것 ★ {new_cnt}건)",
         "- 이 문서는 **읽기 전용 진단**이다. 화면을 자동으로 바꾸지 않는다 —",
         "  고칠지는 사람이 정하고, 고친 뒤에는 다음 회차에서 사라졌는지로 확인한다.", ""]
    if not items:
        L += ["## 지금은 눈에 띄는 문제가 없다", "",
              "오류·검색 무결과·느린 화면이 기준치 아래다. 사용 기록이 더 쌓이면 다시 본다."]
    else:
        L += ["| | 종류 | 대상 | 무엇이 문제인가 | 어떻게 고칠까 |", "|---|---|---|---|---|"]
        for x in items[:25]:
            L.append(f"| {'★' if x['new'] else ''} | {x['kind']} | {x['target']} | "
                     f"{x['why']} | {x['how']} |")
    L += ["", "## 화면별 사용량(참고)", "", "| 화면 | 연 횟수 |", "|---|---:|"]
    for v, n in list(views.items())[:12]:
        L.append(f"| {label(v)} | {n} |")

    os.makedirs(os.path.dirname(OUT_MD), exist_ok=True)
    open(OUT_MD, "w", encoding="utf-8").write("\n".join(L))
    json.dump({"at": time.strftime("%Y-%m-%d %H:%M"), "keys": keys, "items": items},
              open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"UX 개선 제안 {len(items)}건 (새로 ★{new_cnt}) → reports/UX개선_제안.md")
    if "--print" in sys.argv:
        for x in items[:8]:
            print(f"  {'★' if x['new'] else ' '} [{x['kind']}] {x['target']} — {x['why']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
