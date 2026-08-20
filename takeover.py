# -*- coding: utf-8 -*-
"""크레딧이 떨어져도 **다른 계정·다른 창이 이어받는다** (2026-08-17 지시).

사용자 지시: *"다른 계정 다른 세션에서도 진행할 수 있게 알고리즘 추가해서 관리해,
언제든지 이 세션의 크레딧이 떨어지면 다른 계정에서 작업할 수 있게 준비"*

★ 빠져 있던 것은 '이어받는 기능'이 아니다 — 그건 이미 있다
  (`session_handoff --adopt` · `session_boundary --start` · `ai_claim` · `worksplit`).
  빠진 것은 **끊겼다는 사실을 아는 것**이었다. 지금까지 세션이 끝나는 길은 셋뿐이었고
  (compact · `/clear` · 종료) 셋 다 훅이 받아 인계를 남긴다. 그런데 **크레딧 소진은
  훅이 없다** — 창은 그대로 떠 있고 pid 도 살아 있으며 대화기록만 그 자리에서 멈춘다.
  그러면:
    · `ai_claim._is_dead` 는 pid 를 보고 **'살아 있다'** 고 답한다 → 점유가 안 풀린다
    · 인계 문서는 그 창을 **일하는 창**으로 센다 → 다른 계정이 '남의 일'로 알고 비켜선다
    · 아무 화면도 "이 창은 멈췄다"고 말하지 않는다
  즉 **실패가 성공처럼 보이는 자리**다(`[169]`). 여기는 그 침묵을 재는 계기다.

★ **뺏지 않는다. 읽기 전용이다.** 자동 회수는 기존 규칙 몫이다
  (pid 사망 → `ai_claim` · 8시간 자국 없음 → `lanes` 자동회수 `[239]`).
  살아 있는 옆 창의 점유를 빼앗으면 두 창이 같은 파일에서 만난다(사고 #36) —
  **못 뺏는 것보다 나쁘다.** 여기는 보고 말하는 것까지다(`typo_watch` 와 같은 자리).

★ **새 판단을 만들지 않는다**(`[162]`): 살아 있음은 `session_wrapup.stem_ages`,
  이름공간은 `ai_claim.sid_of`, pid 생사는 `ai_claim._is_dead`, 맡은 일은 `worksplit`.
  같은 판단을 두 곳에서 하면 언젠가 갈리고, 갈린 뒤엔 어느 쪽이 맞는지 아무도 모른다.

    python takeover.py            # 지금 창들이 어떤 상태인가
    python takeover.py --write    # reports/이어받기.md 갱신(회차가 부른다)
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:                    # pythonw 에서는 sys.stdout 이 None 이다(`[235]`)
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
CARD = os.path.join(ROOT, "reports", "이어받기.md")

#: 이만큼 조용하면 '쉬는중'(사람이 잠깐 자리를 비웠을 수 있다)
WARM_MIN = int(os.environ.get("COUPANG_TAKEOVER_WARM_MIN", "45"))
#: 이만큼 조용하면 '끊긴듯' — 크레딧 소진·창 방치. 그래도 **뺏지는 않는다.**
COLD_MIN = int(os.environ.get("COUPANG_TAKEOVER_COLD_MIN", "180"))


def _git(*args):
    """git 한 줄. 못 물어보면 None — **'못 읽음'을 '깨끗함'으로 치지 않는다**(`[169]`)."""
    try:
        import proc_guard
        res = proc_guard.run_tree(["git", *args], cwd=ROOT, timeout=25)
    except Exception:
        return None
    if getattr(res, "timed_out", False) or res.returncode != 0:
        return None
    return (res.stdout or "").strip()


def repo_state():
    """다른 계정이 **볼 수 있는 상태인가.**

    폰·웹·다른 PC 는 **밀린 것만** 본다. 그래서 미푸시는 '나중에 밀면 되는 것'이 아니라
    **이어받기가 불가능한 상태**다. 미커밋은 그보다 더하다 — 그 창 안에만 있다.
    """
    out = {"미커밋": None, "미푸시": None, "왜못함": ""}
    dirty = _git("status", "--porcelain")
    if dirty is None:
        out["왜못함"] = "git 상태를 못 읽었다"
        return out
    out["미커밋"] = len([x for x in dirty.splitlines() if x.strip()])
    ahead = _git("rev-list", "--count", "@{u}..HEAD")
    out["미푸시"] = int(ahead) if (ahead or "").isdigit() else None
    if out["미푸시"] is None:
        out["왜못함"] = "upstream 이 없어 미푸시를 셀 수 없다"
    return out


def sessions(now=None):
    """창마다 `(sid, 조용한분, 갈래, 잡은자원[], 맡은일[])`. 못 읽으면 `(None, 이유)`."""
    now = now or time.time()
    try:
        import ai_claim
        import session_wrapup
    except Exception as exc:
        return None, "판단을 빌려올 모듈을 못 불렀다(%s)" % type(exc).__name__
    ages = session_wrapup.stem_ages("")
    if not ages:
        # 폴더를 못 찾은 것과 창이 없는 것은 다르다 — 기계 회차에는 자기 sid 가 없다.
        return None, "대화기록 폴더를 못 찾았다 — reports/.대화기록_폴더.txt 를 확인한다"

    claims = {}
    try:
        for 자원, c in (ai_claim._load_unlocked() or {}).items():
            claims.setdefault(str(c.get("sid") or ""), []).append((자원, c))
    except Exception:
        claims = {}
    맡은일 = {}
    try:
        import worksplit
        for it in (worksplit.load().get("items") or []):
            if str(it.get("state") or "") in ("진행", "대기") and it.get("who"):
                맡은일.setdefault(str(it.get("sid") or it.get("who") or ""), []).append(
                    "[%s] %s" % (it.get("id"), str(it.get("title") or "")[:48]))
    except Exception:
        맡은일 = {}

    rows = []
    for stem, t in ages.items():
        sid = ai_claim.sid_of(stem)
        조용 = (now - t) / 60.0
        내것 = claims.get(sid, [])
        일 = 맡은일.get(sid, [])
        # ★ **조용한 것만으로 '이어받아라'고 하면 경보가 전부가 된다**(2026-08-17 실측:
        #   첫 판이 25건을 쏟았고 대부분 몇 주 전에 **깨끗이 닫힌 창**이었다 — 대화기록
        #   파일은 영원히 남는다). 이어받을 것이 있다는 근거는 하나다: **뭔가를 남겼나.**
        #   자원을 쥐었거나 맡은 일이 진행 중이면 그것은 누가 이어야 하는 일이고,
        #   아무것도 안 남겼으면 이어받을 것이 없다(`[170]`·`[172]`).
        남김 = bool(내것 or 일)
        if 조용 < WARM_MIN:
            갈래 = "일하는중"
        elif 조용 < COLD_MIN:
            갈래 = "쉬는중"
        else:
            갈래 = "끊긴듯" if 남김 else "닫힘"
        # ★ pid 가 이미 죽었으면 그것은 기존 규칙이 회수한다 — 여기서 '끊긴듯'이라
        #   부르면 같은 일을 두 곳에서 판단하게 된다.
        if 내것 and all(ai_claim._is_dead(c) for _, c in 내것):
            갈래 = "죽음(회수됨)"
        rows.append({"sid": sid, "조용한분": round(조용, 1), "갈래": 갈래,
                     "자원": [자원 for 자원, _ in 내것], "맡은일": 일})
    rows.sort(key=lambda r: r["조용한분"])
    return rows, ""


def notices(rows=None, 왜못함="", repo=None):
    """인계 '먼저 처리할 것' 에 올릴 것 — **경보가 아니라 알림**이다.

    끊긴 창이 있다는 것은 고장이 아니다(크레딧은 언젠가 떨어진다). 알아야 할 것은
    **지금 이어받아도 되는가**뿐이다.
    """
    out = []
    if rows is None:
        rows, 왜못함 = sessions()
    if rows is None:
        out.append({"갈래": "확인못함", "무엇": "창 상태를 **못 봤다** — %s" % 왜못함,
                    "어떻게": "python takeover.py"})
        return out
    for r in rows:
        if r["갈래"] != "끊긴듯":
            continue
        가진것 = (" · 잡은 자원 %s" % ", ".join(r["자원"])) if r["자원"] else ""
        일 = (" · 맡은 일 %s" % " / ".join(r["맡은일"][:2])) if r["맡은일"] else ""
        out.append({"갈래": "끊긴듯", "무엇":
                    "창 `%s` 이 **%.0f분째 조용하다**(크레딧 소진일 수 있다)%s%s. "
                    "다른 계정에서 이어받을 수 있다 — 뺏지는 않았다"
                    % (r["sid"], r["조용한분"], 가진것, 일),
                    "어떻게": "python ecount/session_handoff.py --adopt"})
    repo = repo if repo is not None else repo_state()
    if repo.get("미푸시"):
        out.append({"갈래": "안밀림", "무엇":
                    "커밋 **%d개가 아직 안 밀렸다** — 다른 계정·폰·웹은 **밀린 것만** "
                    "본다. 지금 이어받으면 이 작업은 안 보인다" % repo["미푸시"],
                    "어떻게": "git push"})
    if repo.get("미커밋"):
        # ★ 남의 창이 반쯤 고쳐 둔 것일 수 있다 — 자동으로 커밋하지 않는다(`[104]`).
        out.append({"갈래": "미커밋", "무엇":
                    "고쳐 놓고 커밋 안 한 파일 **%d개**가 있다 — 그 창 안에만 있는 "
                    "상태다. 반쯤 고친 것일 수 있어 자동으로 커밋하지 않는다"
                    % repo["미커밋"],
                    "어떻게": "git status -s"})
    if repo.get("왜못함"):
        out.append({"갈래": "확인못함", "무엇": "이어받기 준비 상태를 못 읽었다 — %s"
                    % repo["왜못함"], "어떻게": "git status -sb"})
    return out


def browser_side(now=None):
    """브라우저 수집은 **계정을 안 따라간다** — 이어받는 쪽에 무엇이 없는지 적는다.

    Claude 확장 연결은 크롬이 아니라 **Claude 계정**에 붙는다(2026-08-12 실측 —
    크롬·확장·사람 로그인이 하나도 안 바뀌어도 계정이 바뀌면 목록이 빈다).
    그래서 새 계정은 밴드 탭에 수집기를 **다시 심을 수가 없다.**

    ★ 그러나 **이미 심어 둔 것은 페이지 안에 산다** — 계정과 무관하게 계속 돈다.
      탭을 새로고침하거나 닫거나 기계가 절전에 들어가면 그때 사라진다.
    ★ 그러므로 이어받는 쪽에 필요한 것은 "무엇이 막혔나"가 아니라 **확장 없이
      되는 길**이다 — 로그인된 밴드 탭 콘솔(F12)에 두 줄이면 된다([111]).

    읽기만 한다([168]) — 회차가 써 둔 파일 둘만 본다.  못 읽으면 **0 이 아니라
    "못 읽음"** 이다([169]): 대기 0 으로 보이면 이어받는 쪽이 할 일이 없는 줄 안다.
    """
    now = time.time() if now is None else now
    out = {"대기": None, "밴드": {}, "왜못함": ""}
    try:
        with open(os.path.join(ROOT, "reports", "밴드_수집대기열.json"),
                  encoding="utf-8") as fh:
            q = json.load(fh) or {}
        per = {str(b): len((v or {}).get("nos") or [])
               for b, v in (q.get("bands") or {}).items()}
        out["대기"] = sum(per.values())
        for b, cnt in per.items():
            out["밴드"][b] = {"대기": cnt, "상태": "?", "조용한분": None}
    except Exception as exc:
        out["왜못함"] = "대기열을 못 읽었다: %s" % str(exc)[:80]
        return out
    try:
        with open(os.path.join(ROOT, "reports", "크롬수집_보고.json"),
                  encoding="utf-8") as fh:
            rep = json.load(fh) or {}
        for b, rec in (rep.get("밴드") or {}).items():
            rec = rec or {}
            row = out["밴드"].setdefault(str(b), {"대기": 0, "상태": "?", "조용한분": None})
            row["상태"] = rec.get("state") or "?"
            ts = rec.get("at") or rec.get("받은시각")
            try:
                # ★ 시각대를 맞추지 않으면 몇 시간이 통째로 어긋난다([288]).
                #   `at` 는 UTC(+00:00)로 오고 `받은시각` 은 이 PC 시각이라
                #   tzinfo 를 그냥 떼면 9시간이 밀린다 — 실측으로 145분이
                #   561분으로 나왔다. aware 면 그대로, naive 면 이 PC 시각이다.
                t = datetime.fromisoformat(str(ts).replace("Z", "+00:00"))
                row["조용한분"] = max(0.0, (now - t.timestamp()) / 60.0)
            except Exception:
                row["조용한분"] = None
    except Exception as exc:
        out["왜못함"] = "되보고를 못 읽었다: %s" % str(exc)[:80]
    return out


def card(rows=None, 왜못함="", repo=None):
    """다른 계정이 첫 화면에서 볼 한 장."""
    if rows is None:
        rows, 왜못함 = sessions()
    repo = repo if repo is not None else repo_state()
    L = ["# 이어받기 — 다른 계정에서 지금 시작하려면", "",
         "- 기준: %s" % datetime.now().strftime("%Y-%m-%d %H:%M"),
         "- **첫 명령은 이 한 줄이다**: `python ecount/session_handoff.py --adopt`",
         "  (죽은 세션 점유 회수 · 큐 흡수 · 수집 밀림 판정 · 인계 갱신을 스스로 한다)",
         "- 그다음은 평소대로 시작 체크리스트 1~4.", ""]
    L += ["## 지금 열려 있는 창", ""]
    if rows is None:
        L += ["> **못 봤다** — %s" % 왜못함, "",
              "> 0개가 아니라 **확인 못 함**이다. 이 상태로 남의 점유를 건드리지 않는다.", ""]
    elif not rows:
        L += ["창이 하나도 안 잡힌다(이 회차에는 자기 자신도 없다).", ""]
    else:
        볼것 = [r for r in rows if r["갈래"] != "닫힘"]
        L += ["| 창 | 조용한 시간 | 갈래 | 잡은 자원 | 맡은 일 |", "|---|---:|---|---|---|"]
        for r in 볼것:
            L.append("| `%s` | %.0f분 | %s | %s | %s |" % (
                r["sid"], r["조용한분"], r["갈래"],
                ", ".join(r["자원"]) or "—", " / ".join(r["맡은일"][:2]) or "—"))
        접힘 = len(rows) - len(볼것)
        if 접힘:
            # 조용히 빼지 않는다 — 접었으면 몇 개를 접었는지 말한다(`[169]`).
            L.append("| … | | 닫힘 %d개 | 남긴 것 없음 | |" % 접힘)
        L.append("")
        L += ["> `끊긴듯` 은 **%d분 넘게 대화기록이 안 자란** 창이다. 크레딧이 떨어지면"
              " 창은 그대로 뜬 채 기록만 멈추므로 pid 로는 구별이 안 된다.",
              "> **점유는 뺏지 않았다** — 죽은 것은 `--adopt` 가, 8시간 넘은 차선은"
              " 워치독이 회수한다.", ""]
        L[-3] = L[-3] % COLD_MIN
    L += ["## 다른 계정이 볼 수 있는 상태인가", ""]
    L += ["- 미푸시 커밋: **%s**" % ("못 읽음" if repo["미푸시"] is None else "%d개" % repo["미푸시"]),
          "- 커밋 안 한 파일: **%s**" % ("못 읽음" if repo["미커밋"] is None else "%d개" % repo["미커밋"]),
          "", "> 폰·웹·다른 PC 는 **밀린 것만** 본다. 미푸시가 남아 있으면 이어받는 쪽은"
          " 그 작업을 아예 못 본다 — '나중에 밀면 되는 것'이 아니다.", ""]
    L += _browser_lines()
    return "\n".join(L)


def _browser_lines(b=None):
    """브라우저 수집 칸.  **할 일이 없으면 조용하다**([170])."""
    b = browser_side() if b is None else b
    if b.get("왜못함"):
        return ["## 브라우저 수집", "",
                "> **확인 못 함** — %s" % b["왜못함"], "",
                "> 0건이 아니라 못 읽은 것이다 — 할 일이 없다는 뜻이 아니다.", ""]
    if not b.get("대기"):
        return []
    L = ["## 브라우저 수집 — **계정을 안 따라간다**", "",
         "- 지금 대기: **%d건**" % b["대기"], ""]
    L += ["| 밴드 | 대기 | 마지막 되보고 | 상태 |", "|---|---:|---:|---|"]
    for band in sorted(b["밴드"]):
        r = b["밴드"][band]
        L.append("| %s | %d | %s | %s |" % (
            band, r.get("대기") or 0,
            ("%.0f분 전" % r["조용한분"]) if r.get("조용한분") is not None else "없음",
            r.get("상태") or "?"))
    L += ["",
          "> Claude 확장 연결은 크롬이 아니라 **Claude 계정**에 붙는다 — 계정이 바뀌면"
          " 새 계정은 수집기를 **다시 심을 수 없다**. 이미 심어 둔 것은 페이지 안에"
          " 살아 있어 계속 돈다(탭 새로고침·닫기·절전 전까지).", "",
          "**확장 없이 되는 길** — 로그인된 밴드 탭 콘솔(F12)에 이 두 줄:", "",
          "```js",
          "const s = await (await fetch('http://127.0.0.1:8899/band_auto_collect.user.js')).text();",
          "(0,eval)(s);",
          "```", "",
          "> 그다음은 **크롬 창을 화면에 보이게** 두기만 하면 된다. 숨은 탭에서는"
          " 밴드 본문이 안 그려져 수집기가 스스로 멈춘다 — 그 가드를 우회하면"
          " 실재하는 글에 **되돌릴 수 없는 묘비**가 박힌다(2026-08-19 실사고).", ""]
    return L


def write(text=None):
    text = card() if text is None else text
    os.makedirs(os.path.dirname(CARD), exist_ok=True)
    tmp = CARD + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    os.replace(tmp, CARD)
    return CARD


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    rows, 왜못함 = sessions()
    repo = repo_state()
    if "--write" in argv:
        write(card(rows, 왜못함, repo))
        print("이어받기 카드: reports/이어받기.md")
    if "--json" in argv:
        print(json.dumps({"창": rows, "왜못함": 왜못함, "저장소": repo,
                          "알림": notices(rows, 왜못함, repo)}, ensure_ascii=False))
        return 0
    print(card(rows, 왜못함, repo) if "--write" not in argv else "")
    for n in notices(rows, 왜못함, repo):
        print("  [%s] %s" % (n["갈래"], n["무엇"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
