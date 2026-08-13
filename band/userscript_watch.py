# -*- coding: utf-8 -*-
"""크롬 전용 수집이 **정말 돌고 있는지** 본다 (2026-08-13 지시).

사용자 지시: "앞으로 크롬에서만 긁어오는 알고리즘 만들어서 적용해"

빠져 있던 것은 긁는 기능이 아니었다.  `band/band_auto_collect.user.js` 는
2026-08-09 에 만들어졌고 검증 `[182]` 까지 붙어 있었는데 **2026-08-13 까지 한 번도
안 돌았다** — Tampermonkey 가 안 깔려 있었기 때문이다.  그런데 그 사실을 말해 주는
화면이 어디에도 없었다.  나흘 동안 '자동 수집이 있다'고 적힌 채 아무 일도 안
일어났고, 아무도 몰랐다.  이 프로젝트가 반복해 당한 **'실패가 성공처럼 보이는 자리'**
이며, `schedule_watch` 가 스케줄러 회차에 대해 하는 일을 여기서는 브라우저에 대해 한다.

가르는 것이 세는 것보다 어렵다(`[170]`).  갈래를 뭉치면 경보가 대부분이 되어 아무도
안 본다.  그래서 다섯으로 가른다:

  · **한 번도 안 옴**  — 설치가 안 됐거나 밴드 탭을 안 열었다.  **여기가 기본값이다.**
  · **소식 끊김**      — 전에는 왔는데 한도(기본 6시간)를 넘겼다.
  · **창이 가려짐**    — 살아 있는데 `hidden` 만 온다.  사람이 창만 앞으로 꺼내면 된다.
  · **계획이 굳음**    — 스크립트는 도는데 받아 갈 목록이 낡았다(회차 쪽 문제다).
  · **정상**           — 최근에 긁었거나, 긁을 게 없다는 보고가 신선하다.

읽기 전용이다.  아무것도 안 고치고 큐에도 안 넣는다 — 옳은 조치가 갈래마다 다르고
(설치·창 꺼내기·회차 고치기) 그 판단은 사람 몫이다(`typo_watch` 와 같은 자리).

⚠ **못 읽은 것을 정상이라 하지 않는다**(`[169]`).  보고 파일을 못 읽으면 '이상 없음'이
아니라 **'확인 못 함'** 이다.  이 구별이 없으면 감시자 자신이 눈먼 채 "정상"을 말한다.
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: 유저스크립트가 되보고한 것을 앱 서버가 여기 쌓는다(`POST /api/collect_report`).
REPORT = os.path.join(ROOT, "reports", "크롬수집_보고.json")
#: 사람이 읽는 정본.
OUT = os.path.join(ROOT, "reports", "크롬수집_상태.md")
#: 수집 계획 — 스크립트가 받아 가는 목록.  이것이 낡으면 크롬이 헛돈다.
PLAN = os.path.join(ROOT, "reports", "밴드_수집계획.json")

#: 이 시간을 넘겨 소식이 없으면 끊긴 것으로 본다.  유저스크립트는 30분마다
#: '살아 있다'를 보내므로(HEARTBEAT_MS) 여섯 시간은 열두 번을 놓친 것이다 —
#: 한두 번 놓친 것으로 경보를 올리면 아무도 안 본다(`[170]`).
SILENT_HOURS = 6.0
#: 계획이 이보다 오래되면 굳은 것으로 본다.  회차는 하루 한 번(09:50) 돌므로
#: 하루 반이면 한 번을 통째로 걸렀다는 뜻이다.
PLAN_STALE_HOURS = 36.0


def _now() -> datetime:
    return datetime.now()


def _parse(ts: Any) -> Optional[datetime]:
    """ISO·'YYYY-MM-DD HH:MM' 을 다 받는다.  못 읽으면 None — 지어내지 않는다."""
    s = str(ts or "").strip()
    if not s:
        return None
    s = s.replace("Z", "+00:00")
    for cut in (s, s[:19]):
        try:
            d = datetime.fromisoformat(cut)
            return d.replace(tzinfo=None) + (
                timedelta(hours=9) if d.tzinfo is not None else timedelta(0)
            )
        except ValueError:
            continue
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[: len(fmt) + 2].strip(), fmt)
        except ValueError:
            continue
    return None


def _age_hours(ts: Any, now: Optional[datetime] = None) -> Optional[float]:
    d = _parse(ts)
    if d is None:
        return None
    return ((now or _now()) - d).total_seconds() / 3600.0


def load_reports() -> Tuple[Optional[Dict[str, Any]], str]:
    """(보고, 못읽은이유).  **없는 것과 못 읽은 것을 가른다**(`[169]`)."""
    if not os.path.exists(REPORT):
        return None, "아직 한 건도 안 왔다"
    try:
        with open(REPORT, encoding="utf-8") as fh:
            return json.load(fh) or {}, ""
    except Exception as exc:  # 깨진 파일을 '없음'으로 치면 사고가 조용해진다
        return None, "보고 파일을 못 읽었다: %s" % str(exc)[:120]


def plan_state(now: Optional[datetime] = None) -> Dict[str, Any]:
    """수집 계획의 나이와 건수.  스크립트가 도는데 목록이 낡으면 헛돈다."""
    if not os.path.exists(PLAN):
        return {"있음": False, "왜": "계획 파일이 없다"}
    try:
        with open(PLAN, encoding="utf-8") as fh:
            doc = json.load(fh) or {}
    except Exception as exc:
        return {"있음": False, "왜": "계획을 못 읽었다: %s" % str(exc)[:100]}
    bands = doc.get("bands") or {}
    total = sum(len((v or {}).get("nos") or []) for v in bands.values())
    return {
        "있음": True,
        "생성": doc.get("generated") or "",
        "나이": _age_hours(doc.get("generated"), now),
        "밴드수": len(bands),
        "글수": total,
    }


def judge(doc: Optional[Dict[str, Any]], why: str,
          now: Optional[datetime] = None,
          plan: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """갈래를 고른다.  **모르면 '확인 못 함'** 이지 '정상'이 아니다.

    ★ 계획을 **인자로 받는다** — 여기서 파일을 직접 읽지 않는다.  판정 함수가
    제 손으로 근거를 읽으면 부르는 쪽이 보여 주는 것과 **서로 다른 한 장**을 놓고
    답하게 되고, 시험에서도 진짜 파일이 끼어들어 갈래를 못 갈라 본다
    (`recheck_plan.judge_absent` 가 같은 이유로 이 모양이다).
    """
    now = now or _now()
    plan = plan_state(now) if plan is None else plan

    if doc is None:
        # 파일이 아예 없는 것(설치 전)과 못 읽은 것(고장)을 가른다.
        kind = "안옴" if "안 왔다" in why else "확인못함"
        return {"갈래": kind, "왜": why, "밴드": {}, "계획": plan}

    bands = doc.get("밴드") or doc.get("bands") or {}
    if not bands:
        return {"갈래": "안옴", "왜": "보고 파일은 있는데 밴드가 하나도 없다",
                "밴드": {}, "계획": plan}

    rows: Dict[str, Any] = {}
    freshest: Optional[float] = None
    for band, rec in bands.items():
        rec = rec or {}
        age = _age_hours(rec.get("at") or rec.get("받은시각"), now)
        rows[str(band)] = {
            "상태": rec.get("state") or "?",
            "나이": age,
            "요청": rec.get("요청"),
            "수확": rec.get("수확"),
            "왜": rec.get("why") or "",
        }
        if age is not None and (freshest is None or age < freshest):
            freshest = age

    if freshest is None:
        return {"갈래": "확인못함", "왜": "보고에 시각이 없다", "밴드": rows, "계획": plan}
    if freshest > SILENT_HOURS:
        return {"갈래": "끊김",
                "왜": "가장 최근 보고가 %.1f시간 전이다(한도 %.0f시간)" % (freshest, SILENT_HOURS),
                "밴드": rows, "계획": plan}

    states = {r["상태"] for r in rows.values()}
    # 살아는 있는데 계속 숨어만 있다 — 사람이 창을 앞으로 꺼내면 끝나는 일이다.
    if states and states <= {"hidden"}:
        return {"갈래": "가려짐",
                "왜": "스크립트는 살아 있는데 창이 가려져 있다 — 밴드 탭이 보이면 스스로 시작한다",
                "밴드": rows, "계획": plan}
    # 도는데 받아 갈 목록이 낡았다 — 브라우저가 아니라 회차 쪽 문제다.
    pa = plan.get("나이")
    if "no-plan" in states and pa is not None and pa > PLAN_STALE_HOURS:
        return {"갈래": "계획굳음",
                "왜": "수집 계획이 %.1f시간째 그대로다 — 크롬은 도는데 긁을 목록이 안 온다" % pa,
                "밴드": rows, "계획": plan}
    return {"갈래": "정상", "왜": "가장 최근 보고 %.1f시간 전" % freshest,
            "밴드": rows, "계획": plan}


#: 갈래마다 **무엇을 하면 되는지**를 같이 준다.  '안 돈다'만 말하고 조치를 못 대면
#: 사람이 어디를 고쳐야 하는지 몰라 결국 아무도 안 본다.
FIX = {
    "안옴": "크롬에 Tampermonkey 를 설치하고 http://127.0.0.1:8899/band_auto_collect.user.js "
            "를 연 뒤, 로그인된 밴드 탭을 보이는 창에 열어 둔다",
    "끊김": "밴드 탭이 닫혔거나 유저스크립트가 꺼졌다 — 탭을 다시 열고 Tampermonkey 에서 켜져 있는지 본다",
    "가려짐": "크롬 창을 앞으로 꺼낸다(탭만 열려 있으면 부족하다 — 창이 가려지면 모든 탭이 숨은 것으로 잡힌다)",
    "계획굳음": "09:50 회차의 `band/comment_backfill.py --write` 가 도는지 본다(계획을 만드는 자리다)",
    "확인못함": "보고 파일을 사람이 본다 — 못 읽는 것을 '이상 없음'으로 치지 않는다",
    "정상": "",
}


def lines(state: Optional[Dict[str, Any]] = None) -> List[str]:
    """인계 문서 '먼저 처리할 것' 에 올릴 줄.  정상이면 빈 목록이다."""
    st = state or judge(*load_reports())
    kind = st.get("갈래")
    if kind == "정상":
        return []
    out = ["크롬 전용 수집 — %s · %s" % (kind, st.get("왜") or "")]
    fix = FIX.get(kind or "", "")
    if fix:
        out.append("  → " + fix)
    return out


def render(st: Dict[str, Any]) -> str:
    now = _now().strftime("%Y-%m-%d %H:%M")
    plan = st.get("계획") or {}
    buf = ["# 크롬 전용 수집 상태 (%s)" % now, ""]
    buf.append("- 판정: **%s** — %s" % (st.get("갈래"), st.get("왜") or ""))
    fix = FIX.get(st.get("갈래") or "", "")
    if fix:
        buf.append("- 할 일: %s" % fix)
    buf.append("")
    if plan.get("있음"):
        age = plan.get("나이")
        buf.append("## 수집 계획")
        buf.append("")
        buf.append("- 생성 %s%s · 밴드 %s개 · 글 %s건" % (
            plan.get("생성") or "?",
            (" (%.1f시간 전)" % age) if age is not None else "",
            plan.get("밴드수"), plan.get("글수")))
    else:
        buf.append("## 수집 계획")
        buf.append("")
        buf.append("- **못 읽음** — %s" % plan.get("왜"))
    buf.append("")
    rows = st.get("밴드") or {}
    buf.append("## 밴드별 마지막 보고")
    buf.append("")
    if not rows:
        # 0 을 '이상 없음'으로 읽히게 두지 않는다([169]).
        buf.append("_한 건도 없다 — 없는 것이 아니라 **아직 아무도 보고하지 않은 것**이다._")
    else:
        buf.append("| 밴드 | 상태 | 몇 시간 전 | 요청 | 수확 | 비고 |")
        buf.append("|---|---|---:|---:|---:|---|")
        for band, r in sorted(rows.items()):
            age = r.get("나이")
            buf.append("| %s | %s | %s | %s | %s | %s |" % (
                band, r.get("상태"),
                ("%.1f" % age) if age is not None else "?",
                r.get("요청") if r.get("요청") is not None else "-",
                r.get("수확") if r.get("수확") is not None else "-",
                (r.get("왜") or "")[:60]))
    buf.append("")
    buf.append("---")
    buf.append("")
    buf.append("이 문서는 **읽기만** 한다 — 수집을 다시 띄우지도, 계획을 고치지도 않는다.")
    buf.append("옳은 조치가 갈래마다 다르기 때문이다(설치 · 창 꺼내기 · 회차 고치기).")
    return "\n".join(buf) + "\n"


def check(write: bool = True) -> Dict[str, Any]:
    st = judge(*load_reports())
    if write:
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        tmp = OUT + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(render(st))
        os.replace(tmp, OUT)
    return st


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--print", action="store_true", help="파일을 안 쓰고 화면에만")
    args = ap.parse_args(argv)
    st = check(write=not getattr(args, "print"))
    print("크롬 전용 수집: %s — %s" % (st.get("갈래"), st.get("왜") or ""))
    for ln in lines(st):
        print("  " + ln)
    if not getattr(args, "print"):
        print("  상세: %s" % os.path.relpath(OUT, ROOT))
    # 경보라고 exit 1 을 주지 않는다 — 회차 한 단계를 세우자고 감시자가 죽으면 안 된다.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
