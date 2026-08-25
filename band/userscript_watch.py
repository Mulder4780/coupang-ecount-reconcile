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
import io
import json
import os
import re
import sys
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

# 무인 회차는 pythonw 라 `sys.stdout` 이 **None** 이고(`[235]`), 사람이 콘솔에서 부르면
# 윈도우 기본이 **cp949** 라 본문의 '—' 가 못 나가 통째로 죽는다(2026-08-13 실측 —
# 인계 문서가 알려 주는 `--print` 가 바로 그 명령이었다).  둘 다 여기서 막는다.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

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
#: `start` 되보고 뒤 이만큼이 지나면 **매달린 것**이다.  수집기는 한 회차를
#: `MAX_WAIT_MS` 까지만 기다리고 그 뒤 반드시 `done`/`partial`/`save-failed` 중
#: 하나를 보낸다 — 그러니 그 한도를 넘겨 `start` 로 남아 있으면 끊긴 것이다.
#: ★ 값을 손으로 적지 않는다(`[162]`) — 수집기가 정본이고 여기는 읽기만 한다.
#:   2026-08-22 실사고: 01:34 에 `start` 로 시작한 수집이 51분째 그대로였는데
#:   침묵 한도가 6시간이라 화면은 **정상**이라 말했다(수확 0건 · `[169]`).
#:   6시간은 30분 하트비트를 열두 번 놓친 값이라 옳지만, **매달린 회차는
#:   30분이면 판정할 수 있다.**
START_GRACE_MIN = 10.0   # 저장·업로드에 드는 여유


def _grab_wait_hours():
    """수집기의 회차 한도(시간)를 **수집기 파일에서 읽는다**.

    못 읽으면 None — 그때는 이 판정을 **아예 안 한다**(`[169]`).  모르는 것을
    근거로 '매달렸다'고 부르면 멀쩡히 도는 수집을 끊겼다고 말한다(`[172]`).
    """
    import re as _re
    try:
        js = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "band_auto_collect.user.js")
        src = io.open(js, encoding="utf-8", newline="").read()
        m = _re.search(r"MAX_WAIT_MS\s*=\s*([0-9*\s]+);", src)
        if not m:
            return None
        ms = 1
        for tok in m.group(1).split("*"):
            tok = tok.strip()
            if not tok.isdigit():
                return None
            ms *= int(tok)
        if ms <= 0:
            return None
        return ms / 3600000.0 + START_GRACE_MIN / 60.0
    except Exception:
        return None


#: 계획이 이보다 오래되면 굳은 것으로 본다.  회차는 하루 한 번(09:50) 돌므로
#: 하루 반이면 한 번을 통째로 걸렀다는 뜻이다.
PLAN_STALE_HOURS = 36.0
#: 사람이 크롬에서 여는 자리.  안내 문구가 여기저기 URL 을 손으로 적으면
#: 포트가 바뀐 날 한쪽만 고쳐진다.
def _port():
    """앱 서버 포트를 **app_server 에서 읽는다** — 손으로 적지 않는다(`[156]`).

    8765 라 적어 뒀다가 실제가 8899 라 "서버가 안 떴다"로 읽힌 전례가 있다.
    못 읽으면 8899 로 두되, 그때는 안내가 틀릴 수 있다는 뜻이다."""
    try:
        import sys as _s
        if ROOT not in _s.path:
            _s.path.insert(0, ROOT)
        from webapp import app_server as _a
        return int(getattr(_a, "PORT", 8899))
    except Exception:
        try:
            import re as _r
            with io.open(os.path.join(ROOT, "webapp", "app_server.py"),
                         encoding="utf-8", errors="replace") as _fh:
                m = _r.search(r"^PORT\s*=\s*(\d+)", _fh.read(), _r.M)
            return int(m.group(1)) if m else 8899
        except Exception:
            return 8899


PORT = _port()
USER_JS_URL = "http://127.0.0.1:%d/band_auto_collect.user.js" % PORT

#: ★ **확장·Tampermonkey 없이 되는 길** (2026-08-22 형님 지시 "계정이 바뀌어도 크롬
#: 잘 디버깅해서 잘 붙여서 할 수 있게").  Claude 확장 연결은 크롬이 아니라 **Claude
#: 계정**에 붙는다 — 계정을 바꾸면 `list_connected_browsers` 가 빈 목록이 되고 AI 는
#: 수집기를 **심을 수 없다**.  그때 Tampermonkey 를 지목하면 사람이 멀쩡한 확장을
#: 다시 깔러 간다([172] 틀린 지목 · [149] 로 이미 반증됐다).  이 두 줄은 계정과
#: 무관하게 언제나 된다 — 크롬은 localhost 를 신뢰 출처로 봐서 https 페이지에서도
#: http 로 가져온다([217] 에서 실측).
PASTE_HINT = ("로그인된 밴드 탭에서 F12 → Console 에 이 두 줄을 붙여넣는다: "
              "const s = await (await fetch('%s')).text(); (0,eval)(s);"
              "  — 그다음 그 크롬 창을 화면에 보이게 둔다" % USER_JS_URL)


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


def _read_side():
    """브라우저가 실제로 받는 목록을 읽는 그 함수를 빌린다([162]).

    합치는 자리는 `comment_backfill.load_plan` **하나**다([353]).  여기서 파일을
    따로 읽으면 감시자와 브라우저가 **서로 다른 한 장**을 놓고 답한다.
    """
    try:
        from band import comment_backfill as CB
        return CB, ""
    except Exception:
        pass
    try:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import comment_backfill as CB  # type: ignore
        return CB, ""
    except Exception as exc:
        return None, "읽는 쪽을 못 불렀다: %s" % str(exc)[:100]


def _split_tiers(건수, 전체):
    """대기열 건수를 **할 일**과 **되살아나기 어려운 것**으로 가른다.

    ★ 낱말을 여기 손으로 적지 않는다([162]) — 어느 갈래가 느린지는
      `collect_queue.SLOW_REVIVE` 하나가 정한다. 적어 두면 갈래가 늘어난 날
      이 화면만 옛 표를 보면서 **오류도 안 낸다**([165]).
    ★ **못 읽으면 0 이라 하지 않는다**([169]) — 건수 칸이 없으면 갈래를 모르는
      것이지 "느린 것이 없다"가 아니다. 그때는 None 을 돌려주고 화면이
      예전처럼 한 덩어리로만 말한다.
    """
    if not isinstance(건수, dict) or not 건수:
        return None
    느린이름 = ("오염",)
    try:
        import collect_queue as _CQ  # type: ignore
        느린이름 = tuple(getattr(_CQ, "SLOW_REVIVE", 느린이름) or 느린이름)
    except Exception:
        pass
    느림 = 0
    할일 = 0
    for k, v in 건수.items():
        try:
            n = int(v or 0)
        except (TypeError, ValueError):
            continue
        if str(k) in 느린이름:
            느림 += n
        else:
            할일 += n
    # ⚠ 건수 합과 `전체` 가 어긋날 수 있다(배치 상한·거른 것). **맞추려고
    #   숫자를 지어내지 않는다**([169]) — 있는 그대로 주고 화면이 그렇게 적는다.
    return {"할일": 할일, "느림": 느림, "합": 할일 + 느림, "느린갈래": list(느린이름)}


def plan_state(now: Optional[datetime] = None) -> Dict[str, Any]:
    """수집 계획의 나이와 건수.  스크립트가 도는데 목록이 낡으면 헛돈다.

    ★ **브라우저가 실제로 받는 목록**을 본다([162]).  전에는 `밴드_수집계획.json`
    (댓글 갈래 하나)만 읽었는데, 실측 2026-08-20 에 그 파일은 84789192 를 아예
    담고 있지 않아 이 화면이 그 밴드의 밀린 글을 **0건**이라 말했다 — 그 순간
    대기열에는 **289건**이 있었다([353]).  그 0 을 근거로 죽은 에이전트를 조용히
    넘기면 그것이 곧 [169] 다.  판정은 새로 안 만들고 읽는 쪽이 이미 붙여 주는
    `대기열` 칸(상태·나이·전체)을 그대로 쓴다.
    """
    CB, 왜 = _read_side()
    if CB is None:
        return {"있음": False, "왜": 왜}
    names = set()
    # 밴드 목록은 **두 원천의 합집합**이다 — 한쪽에만 있는 밴드를 빠뜨리면
    # 그 밴드는 이 화면에서 없는 밴드가 된다([169]).
    # ★ 대기열 경로를 여기 손으로 적지 않는다([162]) — 적으면 사본이 둘 되어
    #   대기열이 이사한 날 이 화면만 옛 자리를 본다(검증이 임시 경로로 재려다
    #   진짜 파일을 읽어 실제로 걸렸다).  정하는 자리는 collect_queue 하나다.
    큐경로 = None
    try:
        import collect_queue as _CQ  # type: ignore
        큐경로 = _CQ.QUEUE_PATH
    except Exception:
        pass
    for path in [p for p in (PLAN, 큐경로) if p]:
        try:
            with open(path, encoding="utf-8") as fh:
                names.update(str(b) for b in ((json.load(fh) or {}).get("bands") or {}))
        except Exception:
            continue
    if not names:
        return {"있음": False, "왜": "계획·대기열 어느 쪽도 못 읽었다"}
    per, total, 만든때, q상태, q나이 = {}, 0, "", "", None
    # ★ 갈래별로도 센다 — "남은 건수" 한 덩어리는 며칠이 지나도 안 줄어
    #   사람이 "왜 완료가 안 되나"로 읽는다(2026-08-25 형님 물음).
    #   실측: 782건 중 **오염이 661건(85%)** 이고 실제 할 일은 121건이었다.
    #   판정은 새로 안 만든다([162]) — `건수` 는 collect_queue 가 이미 적어 둔 것이다.
    갈래별 = {}
    for band in sorted(names):
        p = CB.load_plan(band) or {}
        q = p.get("대기열") or {}
        # 한 배치는 상한까지만 실린다 — 밀린 것은 `전체` 가 안다([353]).
        cnt = int(q.get("전체") or len(p.get("nos") or []))
        per[band] = cnt
        total += cnt
        갈래별[band] = _split_tiers(q.get("건수"), cnt)
        만든때 = q.get("만든때") or 만든때
        q상태 = q.get("상태") or q상태
        if q.get("나이시간") is not None:
            q나이 = q.get("나이시간")
    return {
        "있음": True,
        "생성": 만든때 or "",
        # 나이도 읽는 쪽이 잰 것을 쓴다 — 여기서 다시 재면 두 답이 생긴다([162]).
        "나이": q나이 if q나이 is not None else _age_hours(만든때, now),
        "대기열상태": q상태 or "모름",
        "밴드수": len(per),
        "글수": total,
        # 밴드마다 몇 건이 밀려 있나.  죽은 에이전트가 **일감을 두고** 죽었는지를
        # 가르는 근거다([186]) — 총계만으로는 어느 밴드가 굶는지 말할 수 없다.
        "밴드별": per,
        # 갈래를 갈라 둔다 — 받는 쪽이 "할 일"과 "되살아나기 어려운 것"을
        # 한 덩어리로 읽으면 며칠이 지나도 안 줄어 보인다([169]).
        "밴드별갈래": 갈래별,
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

    # ★ **밴드마다 따로 본다**([186]).  전에는 `freshest`(가장 최근 하나)로만
    #   갈라서, 한 밴드가 열 시간째 죽어 있어도 옆 밴드가 방금 보고했으면
    #   가장 최근 값이 작아 **가려짐 한 줄로 덮였다**.  2026-08-19 실측이 그것이다 —
    #   84789192 가 죽은 채 90610953 만 살아 있었는데 화면은 "창을 앞으로 꺼내라"
    #   고만 했고, 그 조치는 죽은 밴드를 한 건도 못 살린다([172] 틀린 지목).
    쉬는밴드 = (plan.get("밴드별") if plan.get("있음") else None)
    죽은 = []
    매달린 = []
    회차한도 = _grab_wait_hours()
    for band, r in rows.items():
        age = r.get("나이")
        멈춤 = (age is None) or (age > SILENT_HOURS)
        # ★ **`start` 로 매달린 것을 '정상'이라 하지 않는다** (2026-08-22 실사고).
        #   수집기는 회차 한도를 넘기면 반드시 끝 신호를 보낸다 — 안 보냈으면
        #   탭이 가려졌거나 닫혔거나 크롬이 타이머를 얼린 것이다.  침묵 한도를
        #   기다리는 동안 화면이 '정상'을 확언하면 그만큼 수확 0인 채로
        #   아무도 모른다(`[169]`).
        if (not 멈춤) and r.get("상태") == "start" \
           and age is not None and 회차한도 is not None and age > 회차한도:
            멈춤 = True
            매달린.append(band)
        r["갈래"] = "끊김" if 멈춤 else (r.get("상태") or "?")
        # 일감이 없는 밴드까지 매일 부르면 경보가 대부분이 되어 아무도 안 본다([170]).
        # 그러나 **모르는 것을 "일감 없음"으로 치지도 않는다**([169]) — 대기열을
        # 못 읽었으면 None 이고, 그때는 조용히 빼지 않고 그 사실을 같이 적는다.
        남은 = None if 쉬는밴드 is None else int(쉬는밴드.get(band) or 0)
        r["밀린글"] = 남은
        # 갈래도 같이 싣는다 — 표가 "782" 한 덩어리로 적으면 며칠이 지나도
        # 안 줄어 보인다(2026-08-25 형님 물음). 없으면 안 싣는다([169]).
        _g = (plan.get("밴드별갈래") or {}).get(band) if plan.get("있음") else None
        if _g:
            r["밀린갈래"] = _g
        if 멈춤 and (남은 is None or 남은 > 0):
            죽은.append(band)

    if freshest is None:
        return {"갈래": "확인못함", "왜": "보고에 시각이 없다", "밴드": rows, "계획": plan}
    if 죽은:
        # 전부 죽었나 몇만 죽었나를 갈라 말한다 — 조치가 다르다([289]).
        전부 = len(죽은) == len(rows)
        # ★ **침묵과 매달림을 갈라 적는다** — 뭉치면 오늘에 대해 틀린 말을
        #   확언한다(`[325]`).  매달린 밴드에 "6시간 넘게 끊겼다"를 붙이면
        #   실제로는 한 시간인데 여섯 시간이라 말하게 되고, 사람은 없는 것을
        #   찾아 나선다(`[172]`).  조치도 다르다(`[289]`) — 침묵은 '스크립트가
        #   안 도는가', 매달림은 '탭이 가려졌는가'다.
        침묵 = [b for b in 죽은 if b not in 매달린]
        조각 = []
        if 침묵:
            말 = ", ".join(sorted(침묵))
            if 전부 and not 매달린:
                조각.append("밴드 %s 되보고가 %.0f시간 넘게 끊겼다"
                            % (말, SILENT_HOURS))
            else:
                조각.append("**밴드 %s 만** 되보고가 끊겼다(한도 %.0f시간) —"
                            " 옆 밴드가 살아 있어 지금까지 한 줄로 덮여 있었다"
                            % (말, SILENT_HOURS))
        if 매달린:
            조각.append("밴드 %s 가 `start` 로 %.0f분 넘게 매달려 있다(수확 0건) —"
                        " 수집기는 회차 한도를 넘기면 끝 신호를 보내므로, 안 왔다는"
                        " 것은 탭이 가려졌거나 닫혔다는 뜻이다"
                        % (", ".join(sorted(매달린)), (회차한도 or 0) * 60))
        왜 = "  · ".join(조각)
        if 쉬는밴드 is None:
            왜 += " · 대기열을 못 읽어 밀린 글이 있는지는 모른다"
        return {"갈래": "끊김", "왜": 왜, "밴드": rows, "계획": plan,
                "끊긴밴드": sorted(죽은)}
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
    # ★ 이 문구는 **확장이 정말 없을 때**만 쓴다.  깔려 있는데도 이 말을 하면 사람이
    #   이미 한 일을 또 하러 간다(2026-08-18 형님 지적) — 갈래는 `fix_for()` 가 고른다.
    # ★ 붙여넣기를 **먼저** 준다 — 계정이 바뀌어도 되는 유일한 길이다.
    #   Tampermonkey 는 뒤에 선택지로만 둔다([149] 로 원인이 아니라고 반증됐다).
    "안옴": PASTE_HINT + "  · (한 번 설치해 두려면) Tampermonkey 에 "
            + USER_JS_URL + " 를 연다",
    "끊김": "밴드 탭이 닫혔거나 크롬 창·계정이 바뀌었다 — " + PASTE_HINT,
    # ★ 2026-08-25 실측 — **한 창에서는 밴드 하나만 긁힌다.** 두 밴드 탭이 같은 창에
    #   나란히 있으면 크롬은 **지금 보고 있는 탭 하나만** 보임으로 치고 나머지는
    #   `hidden` 이다(그날 90610953=보임 · 84789192=가려짐 이 나란히 찍혔다).
    #   "창을 앞으로 꺼내라"만 적어 두면 창은 앞에 있는데 왜 한쪽만 도는지 모른다([169]).
    "가려짐": "크롬 창을 앞으로 꺼낸다(탭만 열려 있으면 부족하다 — 창이 가려지면 모든 탭이 숨은 것으로 잡힌다). 밴드 탭이 여럿이면 한 창에서는 **앞 탭 하나만** 긁힌다 — 탭을 창 밖으로 끌어내 두 창을 나란히 두거나 한 밴드씩 차례로 한다",
    "계획굳음": "09:50 회차의 `band/comment_backfill.py --write` 가 도는지 본다(계획을 만드는 자리다)",
    "확인못함": "보고 파일을 사람이 본다 — 못 읽는 것을 '이상 없음'으로 치지 않는다",
    "정상": "",
}


#: Tampermonkey 로 알려진 확장 ID.  **이것으로 '없음'을 단정하지 않는다** — 아래
#: `chrome_side()` 는 먼저 **유저스크립트 지문**을 찾고, 그것이 안 보일 때만 이 목록을
#: 쓴다(다른 관리자를 쓰면 ID 가 다르다).
TM_IDS = (
    "dhdgffkkebhmkfjojejmpbldmpobfkfo",   # Tampermonkey (크롬 웹스토어)
    "gcalenpjmijncebpfijmoaglllgpjagf",   # Tampermonkey Beta
)

#: 확장 저장소에서 훑을 크기 상한.  워치독이 30분마다 부르므로 무한정 읽지 않는다.
SCAN_BYTES = 24 * 1024 * 1024


def _script_mark() -> Optional[str]:
    """유저스크립트를 알아볼 지문(`@namespace`)을 **그 파일에서 읽는다**.

    ★ 여기에 문자열을 손으로 적으면 사본이 둘이 된다 — `.user.js` 의 이름이 바뀐 날
      지문만 옛것으로 남아 **오류 없이 한 건도 안 걸리고**, 안내는 다시 "설치하세요"
      로 돌아간다(`[165]` 의 모양이자 이 항목이 고치려는 바로 그 고장이다).
    """
    try:
        with open(os.path.join(ROOT, "band", "band_auto_collect.user.js"),
                  encoding="utf-8") as fh:
            for ln in fh:
                if "@namespace" in ln:
                    v = ln.split("@namespace", 1)[1].strip()
                    return v or None
                if "==/UserScript==" in ln:
                    break
    except Exception:
        return None
    return None


def _profiles(user_data: str) -> List[str]:
    out = []
    try:
        for name in sorted(os.listdir(user_data)):
            p = os.path.join(user_data, name)
            if not os.path.isdir(p):
                continue
            if os.path.isdir(os.path.join(p, "Extensions")) or \
               os.path.isdir(os.path.join(p, "Local Extension Settings")):
                out.append(p)
    except Exception:
        return []
    return out


def _find_mark(store: str, mark: bytes) -> Optional[str]:
    """확장 저장소에서 지문을 찾고 **그 옆의 `enabled` 상태**를 돌려준다.

    돌려주는 값: `"켜짐"` · `"꺼짐"` · `"모름"`(지문은 있는데 상태를 못 읽음) · `None`(없음).

    ⚠ 가까운 `"enabled"` 를 쓴다 — 유저스크립트가 여럿이면 옆 스크립트의 값을 볼 수 있다.
      그래서 이 값은 **안내 문구를 고르는 데만** 쓰고 아무것도 안 고친다.  틀려도
      사람이 대시보드에서 바로 확인한다(쓰는 길이 아니라 읽는 길이다 · `[170]`).
    """
    budget = SCAN_BYTES
    seen = False
    try:
        names = sorted(os.listdir(store))
    except Exception:
        return None
    for name in names:
        if not (name.endswith(".log") or name.endswith(".ldb")):
            continue
        path = os.path.join(store, name)
        try:
            size = os.path.getsize(path)
            if size > budget:
                continue
            budget -= size
            with open(path, "rb") as fh:
                blob = fh.read()
        except Exception:
            continue          # 크롬이 물고 있을 수 있다 — 못 읽은 것은 없는 것이 아니다
        i = blob.find(mark)
        if i < 0:
            continue
        seen = True
        best, state = None, None
        for m in re.finditer(rb'"enabled"\s*:\s*(true|false)', blob):
            d = abs(m.start() - i)
            if d <= 4000 and (best is None or d < best):
                best, state = d, ("켜짐" if m.group(1) == b"true" else "꺼짐")
        return state or "모름"
    return "모름" if seen else None


def _dev_mode(profiles):
    """Chrome 의 **개발자 모드(= 사용자 스크립트 허용)** 가 켜져 있나.

    2026-08-19 실측으로 [149] 가 여기서 갈렸다.  Tampermonkey 는 설치돼 있고
    (`Local Extension Settings` 에 지문 실재) 스크립트도 `enabled:true` v2.0 이고
    `@match` 도 그 주소와 맞는데, 로그인된 밴드 탭의 `localStorage` 에
    `coupangAutoCollect.*` 가 **하나도 없었다**.  Chrome 151 은 MV3 라
    이 토글이 꺼져 있으면 Tampermonkey 가 스크립트를 **아예 못 넣는다** —
    '설치돼 있고 켜져 있는데 안 도는' 증상의 모양이 정확히 이것이다.

    ★ **못 읽으면 '모름'이다**(`[169]`).  꺼짐으로 단정하면 사람이 이미 켜 둔
      토글을 또 찾으러 간다(`[172]`).
    """
    for prof in profiles:
        pref = os.path.join(prof, "Preferences")
        if not os.path.isfile(pref):
            continue
        try:
            with io.open(pref, encoding="utf-8", newline="") as fh:
                d = json.load(fh)
        except Exception:
            continue
        ui = ((d.get("extensions") or {}).get("ui") or {})
        if "developer_mode" in ui:
            return ("켜짐" if ui.get("developer_mode") else "꺼짐"), os.path.basename(prof)
        # 키 자체가 없다 = 한 번도 켠 적이 없다(크롬 기본값 false).
        return "꺼짐", os.path.basename(prof)
    return "모름", None


def chrome_side(user_data: Optional[str] = None) -> Dict[str, Any]:
    """크롬 쪽 사실을 **로컬 증거로만** 잰다 (읽기 전용 · 크롬을 조종하지 않는다).

    2026-08-18 형님 지적: 안내가 *"크롬에 Tampermonkey 를 설치하고…"* 로 시작하는데
    **이미 깔려 있었다.**  실측하니 유저스크립트까지 8/13 에 들어가 켜져 있었다 —
    즉 안내가 **이미 한 일 둘을 또 시키면서** 정작 남은 하나를 안 말했다.

    갈래는 넷이고 뜻이 다 다르다(`[169]` — 못 읽은 것을 '없음'이라 하지 않는다):
      · `확장`   : 있음 · 없음 · 모름
      · `스크립트`: 켜짐 · 꺼짐 · 없음 · 모름

    ★ 여기서 재는 것은 **깔림**과 **켜짐**뿐이다.  그 스크립트가 실제로 **돌았는지**는
      되보고만이 안다(`load_reports`) — *깔림 · 켜짐 · 살아 있음 · 연결됨은 넷 다
      다른 말이다*(2026-08-12).
    """
    base = user_data or os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "User Data")
    out: Dict[str, Any] = {"확장": "모름", "스크립트": "모름", "판": None,
                           "프로필": None, "왜": "", "지문": None,
                           "개발자모드": "모름"}
    mark = _script_mark()
    out["지문"] = mark
    if not base or not os.path.isdir(base):
        out["왜"] = "크롬 사용자 폴더를 못 찾았다(%s)" % (base or "경로 없음")
        return out
    profs = _profiles(base)
    if profs:
        out["개발자모드"] = _dev_mode(profs)[0]
    if not profs:
        out["왜"] = "크롬 프로필을 못 읽었다"
        return out

    ext_dir = None
    if mark:
        needle = mark.encode("utf-8")
        for prof in profs:
            store_root = os.path.join(prof, "Local Extension Settings")
            if not os.path.isdir(store_root):
                continue
            try:
                ids = sorted(os.listdir(store_root))
            except Exception:
                continue
            for eid in ids:
                state = _find_mark(os.path.join(store_root, eid), needle)
                if state:
                    out["스크립트"] = state
                    out["확장"] = "있음"
                    out["프로필"] = os.path.basename(prof)
                    ext_dir = os.path.join(prof, "Extensions", eid)
                    break
            if ext_dir:
                break
    else:
        out["왜"] = "유저스크립트 파일에서 @namespace 를 못 읽었다"

    if not ext_dir:
        # 지문이 안 보였다 — 확장 자체는 있나?  이것이 '없음'과 '스크립트만 없음'을 가른다.
        for prof in profs:
            for tid in TM_IDS:
                p = os.path.join(prof, "Extensions", tid)
                if os.path.isdir(p):
                    out["확장"] = "있음"
                    out["프로필"] = os.path.basename(prof)
                    ext_dir = p
                    break
            if ext_dir:
                break
        if mark:
            out["스크립트"] = "없음" if ext_dir else "모름"
        if not ext_dir:
            out["확장"] = "없음"

    if ext_dir and os.path.isdir(ext_dir):
        try:
            vers = sorted(os.listdir(ext_dir))
            if vers:
                out["판"] = vers[-1].split("_")[0]
        except Exception:
            pass
    return out


RULED_OUT = os.path.join(ROOT, "reports", "크롬수집_후보제외.json")


def _dev_mode_ruled_out(path: Optional[str] = None) -> bool:
    """'개발자 모드' 후보가 **이미 반증됐나** (2026-08-19 실측).

    이 갈래는 원래 `Preferences` 의 `extensions.ui.developer_mode` 만 봤다.
    그런데 그 파일은 **실시간이 아니다** — 크롬은 설정을 메모리에 두고 뒤늦게
    디스크에 쓴다.  그래서 사람이 화면에서 방금 켜도 여기서는 한동안 `꺼짐` 으로
    읽히고, 감시자는 **이미 한 일을 또 시킨다**(`[172]`).

    그리고 2026-08-19 실측으로 그 지목 자체가 반증됐다 — 켜고 새로고침한 뒤에도
    밴드 탭 `localStorage` 에 `coupangAutoCollect.loaded` 가 **없었다**.
    유저스크립트는 숨은 탭에서도 로드되면 그 열쇠를 먼저 쓰므로(스크립트 59행),
    그 열쇠가 없다는 것은 **본문이 한 줄도 안 돌았다**는 뜻이고 탭이 가려진
    탓이 아니다.

    ★ **못 읽으면 False 다**(`[169]`) — '반증됐다'를 지어내지 않는다.
      파일이 없으면 예전대로 개발자 모드를 후보로 올린다.
    """
    try:
        with open(path or RULED_OUT, encoding="utf-8") as fh:
            d = json.load(fh)
        return bool(isinstance(d, dict) and d.get("개발자모드"))
    except Exception:
        return False


def fix_for(kind: str, side: Optional[Dict[str, Any]] = None) -> str:
    """갈래마다 **무엇을 하면 되는지**.

    `안옴` 만은 고정 문구가 될 수 없다 — 무엇이 남았는지가 크롬 쪽 상태에 달렸다.
    이미 한 일을 또 시키면 사람은 그 안내를 다시 안 읽는다(`[170]` 과 같은 결말이다).
    """
    if kind != "안옴":
        return FIX.get(kind or "", "")
    s = side if side is not None else chrome_side()
    ext, scr, ver = s.get("확장"), s.get("스크립트"), s.get("판")
    tag = "Tampermonkey" + (" v%s" % ver if ver else "")
    if ext == "없음":
        return FIX["안옴"]
    if ext == "모름" or scr == "모름":
        return ("크롬 쪽을 **확인 못 했다**(%s) — 설치돼 있을 수도 있다. "
                "%s 를 크롬에서 열어 [설치]인지 [이미 설치됨]인지 보고, "
                "로그인된 밴드 탭을 보이는 창에 열어 둔다" % (s.get("왜") or "근거 없음", USER_JS_URL))
    if scr == "없음":
        return ("%s 는 이미 깔려 있다 — **유저스크립트만 아직 안 들어갔다**. "
                "크롬에서 %s 를 열고 [설치] 를 누른다" % (tag, USER_JS_URL))
    if scr == "꺼짐":
        return ("%s 도 유저스크립트도 들어 있는데 **꺼져 있다** — "
                "Tampermonkey 대시보드에서 켠다" % tag)
    # 켜짐 — 둘 다 됐다.  여기서 **원인을 확언하면 안 된다** (2026-08-19 실측).
    #   그날 로그인된 밴드 탭을 그 주소로 실제로 열었다(피드 카드 10장 · 본문
    #   3,597자 · 그 페이지에서 `/api/ping` 200).  그런데 되보고는 **그대로 0건**
    #   이었고, 그 탭의 `localStorage` 에 이 스크립트가 쓰는
    #   `coupangAutoCollect.beat.<밴드>` 열쇠가 **하나도 없었다** — 본문이 한 줄도
    #   안 돈 것이다.  그러니 "한 번도 안 열렸다"는 **틀린 지목**이었고, 그 안내를
    #   따르면 사람은 **이미 한 일을 또 한다**(`[172]`·`[257]` 과 같은 자리).
    dev = s.get("개발자모드")
    if dev == "꺼짐" and not _dev_mode_ruled_out():
        # 2026-08-19 실측 — [149] 는 여기서 갈렸다.  설치·켜짐·@match 가 다 맞는데도
        # 밴드 탭 localStorage 에 `coupangAutoCollect.*` 가 0개였다.  Chrome 151 은
        # MV3 라 이 토글이 꺼져 있으면 Tampermonkey 가 스크립트를 **못 넣는다**.
        # ★ 확정이라 적지 않는다(`[172]`) — 가르는 법을 같이 준다.
        return ("%s 도 유저스크립트도 **이미 깔려 있고 켜져 있다**(v%s) — 다시 설치할 것 없다. "
                "남은 후보 하나는 **크롬 개발자 모드가 꺼져 있는 것**이다"
                "(이 프로필의 `extensions.ui.developer_mode` 가 꺼짐 · Chrome 은 MV3 라 "
                "이 토글이 없으면 Tampermonkey 가 스크립트를 아예 못 넣는다). "
                "조치 — `chrome://extensions` 를 열고 오른쪽 위 **개발자 모드**를 켠 뒤 "
                "로그인된 밴드 탭을 새로고침한다. "
                "확인 — 그 탭 콘솔(F12)에 "
                "`Object.keys(localStorage).filter(k=>k.indexOf('coupangAutoCollect')===0)` "
                "를 쳐서 **비어 있지 않으면** 그것이 원인이었다"
                % (tag, ver or "?"))
    return ("%s 도 유저스크립트도 **이미 깔려 있고 켜져 있다** — 다시 설치할 것 없다. "
            "그런데도 되보고가 0건이면 남은 후보는 **둘**이고 조치가 서로 다르다 — "
            "① 로그인된 밴드 탭이 아직 `www.band.us/band/<번호>` 주소로 안 열렸다"
            "(피드·다른 주소는 안 걸린다) · ② 열렸는데 그 탭에 스크립트가 안 들어갔다. "
            "가르는 법 — 그 탭 콘솔(F12)에 "
            "`Object.keys(localStorage).filter(k=>k.indexOf('coupangAutoCollect')===0)` "
            "를 쳐 본다. 비어 있으면 ②다(2026-08-19 실측은 ②였다). "
            "★ ②라면 **다음은 Tampermonkey 아이콘**이다 — 그 밴드 탭에서 아이콘을 눌러 "
            "*이 페이지에서 실행 중인 스크립트* 목록에 이 스크립트가 뜨는지 본다. "
            "안 뜨면 그 확장이 이 사이트에 주입을 못 하는 것이고, 뜨는데도 열쇠가 "
            "없으면 스크립트 본문이 죽은 것이다(콘솔 오류를 본다)" % tag)


def lines(state: Optional[Dict[str, Any]] = None) -> List[str]:
    """인계 문서 '먼저 처리할 것' 에 올릴 줄.  정상이면 빈 목록이다."""
    st = state or judge(*load_reports())
    kind = st.get("갈래")
    if kind == "정상":
        return []
    out = ["크롬 전용 수집 — %s · %s" % (kind, st.get("왜") or "")]
    fix = fix_for(kind or "", st.get("크롬쪽"))
    if fix:
        out.append("  → " + fix)
    return out


def render(st: Dict[str, Any]) -> str:
    now = _now().strftime("%Y-%m-%d %H:%M")
    plan = st.get("계획") or {}
    buf = ["# 크롬 전용 수집 상태 (%s)" % now, ""]
    buf.append("- 판정: **%s** — %s" % (st.get("갈래"), st.get("왜") or ""))
    fix = fix_for(st.get("갈래") or "", st.get("크롬쪽"))
    if fix:
        buf.append("- 할 일: %s" % fix)
    buf.append("")
    if plan.get("있음"):
        age = plan.get("나이")
        buf.append("## 수집 계획")
        buf.append("")
        _tot = {"할일": 0, "느림": 0}
        for _v in (plan.get("밴드별갈래") or {}).values():
            if _v:
                _tot["할일"] += int(_v.get("할일") or 0)
                _tot["느림"] += int(_v.get("느림") or 0)
        buf.append("- 생성 %s%s · 밴드 %s개 · 글 %s건%s" % (
            plan.get("생성") or "?",
            (" (%.1f시간 전)" % age) if age is not None else "",
            plan.get("밴드수"), plan.get("글수"),
            ("  ← **할 일 %s건** · 오염 %s건" % (_tot["할일"], _tot["느림"]))
            if _tot["느림"] else ""))
        if _tot["느림"]:
            # 뺀 것이 아니라 **뒤로 미룬 것**이라고 말한다([172]) — 조용히 빼면
            # 그 번호들은 어느 화면에도 안 뜬다(2026-08-24 실사고 609건).
            buf.append("")
            buf.append("  > 오염은 **맨 뒤에서 긁는다**([177] 순서가 곧 우선순위). "
                       "다시 긁어야 되살아나므로 목록에서 빼지 않는다 — 다만 "
                       "실측으로 되살아난 적이 드물어(10회 긁은 번호 4개가 그대로) "
                       "**이 숫자가 안 줄어드는 것은 고장이 아니다**.")
    else:
        buf.append("## 수집 계획")
        buf.append("")
        buf.append("- **못 읽음** — %s" % plan.get("왜"))
    buf.append("")
    side = st.get("크롬쪽") or {}
    if side:
        buf.append("## 크롬 쪽 — 깔렸나 · 켜졌나")
        buf.append("")
        buf.append("| 확장 | 유저스크립트 | 판 | 프로필 |")
        buf.append("|---|---|---|---|")
        buf.append("| %s | %s | %s | %s |" % (
            side.get("확장"), side.get("스크립트"),
            side.get("판") or "-", side.get("프로필") or "-"))
        if side.get("왜"):
            buf.append("")
            buf.append("- 못 잰 이유: %s" % side.get("왜"))
        buf.append("")
        buf.append("_여기서 재는 것은 **깔림**과 **켜짐**뿐이다 — 그 스크립트가 실제로")
        buf.append("**돌았는지**는 아래 되보고만이 안다(2026-08-12: 깔림·켜짐·살아 있음·")
        buf.append("연결됨은 넷 다 다른 말이다)._")
        buf.append("")
    rows = st.get("밴드") or {}
    buf.append("## 밴드별 마지막 보고")
    buf.append("")
    if not rows:
        # 0 을 '이상 없음'으로 읽히게 두지 않는다([169]).
        buf.append("_한 건도 없다 — 없는 것이 아니라 **아직 아무도 보고하지 않은 것**이다._")
    else:
        buf.append("| 밴드 | 판정 | 상태 | 몇 시간 전 | 밀린 글 | 요청 | 수확 | 비고 |")
        buf.append("|---|---|---|---:|---:|---:|---:|---|")
        for band, r in sorted(rows.items()):
            age = r.get("나이")
            남은 = r.get("밀린글")
            # ★ 한 덩어리로 적지 않는다 — 오염이 85%%인 날 그 숫자는 며칠이
            #   지나도 안 줄어 "완료가 안 된다"로 읽힌다(2026-08-25 형님 물음).
            _g = r.get("밀린갈래") or {}
            if _g and _g.get("느림"):
                남은글 = "%s (할일 %s · 오염 %s)" % (
                    남은 if 남은 is not None else "?", _g.get("할일"), _g.get("느림"))
            else:
                남은글 = 남은 if 남은 is not None else "모름"
            buf.append("| %s | %s | %s | %s | %s | %s | %s | %s |" % (
                band, r.get("갈래") or "?", r.get("상태"),
                ("%.1f" % age) if age is not None else "?",
                남은글,
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
    # 크롬 쪽은 **한 번만** 잰다 — 안내와 리포트가 서로 다른 측정을 보면 두 답이 생긴다.
    try:
        st["크롬쪽"] = chrome_side()
    except Exception as e:                       # 못 재도 회차를 세우지 않는다
        st["크롬쪽"] = {"확장": "모름", "스크립트": "모름", "판": None,
                        "프로필": None, "왜": "측정 실패: %s" % e, "지문": None}
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
