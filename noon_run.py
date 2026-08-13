# -*- coding: utf-8 -*-
"""
noon_run.py — **하루 한 번 정오 회차** (2026-08-11 지시)
================================================================================
사용자 지시: "하루한번 12시에서 13시 사이 다른 세션이랑 충돌 안나게 실행 알고리즘
코딩해(자동화 100%, 컴팩팅도 자동화, 내가 손 안대게 처리)"

**왜 대화 루프가 아니라 스케줄러인가.** 세션 안의 `/loop` 는 창이 닫히면 같이 죽고,
깨우기 간격의 상한이 1시간이라 '하루 한 번'을 애초에 표현하지 못한다. 상시 원칙 그대로다 —
**파일과 작업 스케줄러에 넣은 것만 산다.** 그래서 이 회차는 사람도 대화도 없이 돈다.

## 창과 중복 — '하루 한 번'을 지키는 방법
- 창은 **12:00~13:00**(`COUPANG_NOON_WINDOW=12:00-13:00` 으로 조절). 창 밖에서는
  아무것도 하지 않는다. 스케줄러는 12:00 부터 **10분마다** 부르는데, 그것은 '여러 번
  돌라'는 뜻이 아니라 **양보한 뒤 다시 시도할 기회**를 주려는 것이다.
- 오늘 이미 완주했으면 즉시 물러난다(마커 `reports/.정오회차.json` 의 날짜).
  ★ 마커는 **완주했을 때만** 오늘 날짜로 찍는다. 양보하고 물러난 것을 '했다'로 적으면
    그날 회차는 **영영 안 돈다** — 아무 일도 안 일어났는데 안심시키는 결과다.

## 충돌 — 빼앗지 않고 **양보한다**
다른 세션(사람이 대화 중인 Claude·Codex·수집 세션)이 배타 자원을 잡고 있거나 09:50·
11:00 회차가 아직 돌고 있으면 **이번 부름은 그냥 물러난다.** 10분 뒤 다시 불리므로
창 안에서 최대 여섯 번의 기회가 있고, 끝까지 못 잡으면 그날은 건너뛰고 리포트에 남긴다.
★ 단, **도는 회차 때문에** 마지막 기회까지 밀렸으면 Z: 를 훑는 무거운 단계만 비켜 두고
돈다(`HEAVY_STEPS`). 증분 파이프라인은 5분마다 불리는데 새 자료를 만나면 12분을 넘겨
돌아서, 그대로 두면 **자료가 들어온 날은 하루도 못 돌면서 경보도 안 뜬다**(실측 2026-08-12).
`ledger`·`publish` 점유는 그래도 전부 양보다 — 그쪽은 vN+1 을 쓰는 중이다.
`--force` 로도 **남의 점유는 빼앗지 않는다** — 강제는 창·중복만 무시한다.

## 무엇을 하나 (전부 결정론적·재실행 안전·읽기 전용에 가깝다)
인계 갱신 · 합성검증 · 정기점검 내용 조사 · 캠프명 표준 대조(조사만) · 컴팩팅 자동설정
점검 · 미커밋 집계. **수집은 하지 않는다**(수집 세션 몫) · **엑셀을 열지 않는다**
(보관본은 11:00·15:00 회차 몫) · **커밋하지 않는다**(작업 폴더는 세션끼리 공유한다).

  python noon_run.py            # 스케줄러가 부르는 그대로 (창·중복·충돌을 본다)
  python noon_run.py --status   # 지금 상태만
  python noon_run.py --force    # 창·중복 무시 (남의 점유는 그래도 양보)
"""
import sys, os, json, subprocess
from datetime import datetime, date

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPORT_DIR = os.path.join(ROOT, "reports")
MARKER = os.path.join(REPORT_DIR, ".정오회차.json")
REPORT_MD = os.path.join(REPORT_DIR, "정오회차.md")
SETTINGS = os.path.join(ROOT, "..", ".claude", "settings.json")
BUDGET_MIN = int(os.environ.get("COUPANG_NOON_BUDGET_MIN", "25"))
EXCLUSIVE = ("ledger", "code", "band", "publish")
# ★ 창의 **마지막 기회**(끝 10분). 여기까지 양보만 했으면 무거운 단계만 비켜 두고 돈다.
LAST_CHANCE_MIN = 10
# Z:(SMB)를 함께 훑는 단계 — 이름은 `steps()` 와 **글자까지 같아야** 한다.
# 어긋나면 아무것도 안 건너뛰면서 건너뛴 줄 안다(`[165]` 모양). 검증이 지킨다.
HEAVY_STEPS = ("정기점검 내용 조사", "캠프명 표준 대조(조사만)")


def window():
    """(시작, 끝) 분 단위. 기본 12:00~13:00."""
    raw = os.environ.get("COUPANG_NOON_WINDOW", "12:00-13:00")
    try:
        a, b = raw.split("-")
        ah, am = (int(x) for x in a.split(":"))
        bh, bm = (int(x) for x in b.split(":"))
        return ah * 60 + am, bh * 60 + bm
    except Exception:
        return 12 * 60, 13 * 60


def load_marker():
    try:
        return json.load(open(MARKER, encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def decide(now, marker, claims, rounds, force=False):
    """돌 것인가 — **순수 함수**(검증이 그대로 주입한다).

    claims: [{"what","who","sid","alive"}]  · rounds: ["09:50 일일대조", ...] 도는 회차 이름
    돌려주는 것: {"go": bool, "why": str, "kind": 창밖|완주|양보|간다, "skip": [건너뛸 단계]}

    ★ **가를 수 있는 것을 '또는'으로 묶지 않는다.** 첫판은 배타 점유가 하나라도 살아
      있으면 통째로 물러났는데, 대화 중인 세션은 `code` 를 몇 시간씩 잡고 있다 —
      그러면 정오 회차는 **매일 한 번도 안 돌면서 아무 화면에도 티가 안 난다.**
      그래서 무엇과 부딪히는지를 단계별로 본다:
        · 회차가 도는 중 → **전부 양보**(Z: SMB 를 독점한다 — 진짜 충돌이다)
        · `ledger`·`publish` → **전부 양보**(vN+1 을 쓰는 중이다. 옆에서 읽지 않는다)
        · `code` → **합성검증만 건너뛴다.** 남이 반쯤 고쳐 둔 코드가 내는 빨강은
          내 회차의 사실이 아니다 — 그 경보는 사람을 엉뚱한 곳으로 보낸다.
        · `band` → 그대로 돈다(이 회차는 수집을 하지 않는다).
    """
    start, end = window()
    minute = now.hour * 60 + now.minute
    live = {c.get("what") for c in claims if c.get("alive")}
    # ★ 남의 점유는 force 로도 안 뺏는다 — 그것이 이 회차의 첫째 규칙이다.
    for what in ("ledger", "publish"):
        if what in live:
            c = next(x for x in claims if x.get("what") == what and x.get("alive"))
            return {"go": False, "kind": "양보", "skip": [],
                    "why": f"다른 세션이 '{what}' 를 잡고 있다({c.get('who','?')}"
                           f"[{str(c.get('sid') or '')[:8]}])"}
    if rounds:
        # ★ **매일 양보만 하는 회차는 없는 회차와 같다** (2026-08-12 실측). 증분 파이프라인은
        #   5분마다 불리는데 새 자료를 만나면 **12분을 넘겨** 돈다(ERP 대조 한 단계가 417초).
        #   그러니 자료가 들어온 날은 이 락이 창 내내 걸려 있고, 여섯 번의 기회가 **전부**
        #   양보로 끝난다 — 리포트에는 '양보'라 적히고 그것은 실패가 아니라서 아무 경보도
        #   안 뜬다(`[169]` 모양의 조용한 사고). 양보의 뜻은 **'Z: 를 같이 긁지 않는다'**
        #   이지 '아무것도 하지 않는다'가 아니다. 그래서 창의 마지막 기회에는 Z: 를 훑는
        #   단계만 비켜 두고 나머지(관문·인계)는 돈다.
        if start <= minute < end and minute >= end - LAST_CHANCE_MIN:
            return {"go": True, "kind": "간다(마지막 기회)", "skip": list(HEAVY_STEPS),
                    "why": "회차가 도는 중(%s) — 창이 끝나므로 Z: 를 훑는 단계(%s)만 "
                           "건너뛰고 돈다" % (", ".join(rounds), ", ".join(HEAVY_STEPS))}
        return {"go": False, "kind": "양보", "skip": [],
                "why": f"회차가 도는 중: {', '.join(rounds)}"}
    if not force:
        if not (start <= minute < end):
            return {"go": False, "kind": "창밖", "skip": [],
                    "why": f"창({start//60:02d}:{start%60:02d}~{end//60:02d}:{end%60:02d}) 밖이다"}
        if str(marker.get("done_date") or "") == now.date().isoformat():
            return {"go": False, "kind": "완주", "skip": [], "why": "오늘 회차는 이미 끝났다"}
    skip = ["합성검증"] if "code" in live else []
    why = "창 안 · 오늘 미완주 · 충돌 없음"
    if skip:
        why = "창 안 · 다른 세션이 코드를 고치는 중이라 합성검증만 건너뛴다"
    return {"go": True, "kind": "간다", "skip": skip, "why": why}


def live_claims():
    """살아 있는 배타 점유만 추린다(죽은 세션 것은 잡은 것이 아니다 — [210]·[213])."""
    out = []
    try:
        import ai_claim
        import pid_alive as P
        for what, c in (ai_claim.load() or {}).items():
            if not isinstance(c, dict):
                continue
            # 스케줄러가 잡은 점유는 `agent_pid` 가 0 이라 `pid` 로 봐야 한다(ai_claim 과 같은 규약).
            pid = c.get("agent_pid") or c.get("pid")
            alive = P.owner_alive(pid, c.get("agent_pid_started_at") if c.get("agent_pid")
                                  else c.get("pid_started_at"))
            out.append({"what": what, "who": c.get("who"), "sid": c.get("sid"),
                        # None(못 잼)은 **살아 있는 것으로 본다** — 의심스러우면 양보한다.
                        "alive": alive is not False})
    except Exception:
        pass                      # 점유를 못 읽으면 판단할 근거가 없다 → 아래 회차 검사로 간다
    return out


def running_rounds():
    """지금 도는 회차 이름들 — 잠금 주인이 실제로 살아 있을 때만 센다."""
    names = []
    try:
        import pid_alive as P
    except Exception:
        return names
    for path, label in ((os.path.join(REPORT_DIR, ".daily_run.lock"), "일일대조(09:50)"),
                        (os.path.join(REPORT_DIR, ".automation_pipeline.lock"), "증분 파이프라인"),
                        (os.path.join(REPORT_DIR, ".ledger_db_apply.lock"), "보관본 생성(11·15시)")):
        try:
            d = json.load(open(path, encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if P.owner_alive(d.get("pid"), d.get("pid_started_at")) is not False:
            names.append(label)
    return names


def compact_settings():
    """컴팩팅이 정말 자동인가 — 설정을 **읽어서** 말한다(짐작하지 않는다).

    사용자가 "컴팩팅도 자동화"라고 한 것은 이미 되어 있는 것을 **확인해 두라**는 뜻이다.
    훅이 빠지거나 자동요약이 꺼지면 세션이 가득 찬 순간 인계 없이 끊긴다 — 그런데
    그 사실은 **가득 차기 전까지 아무 화면에도 안 뜬다.**
    """
    try:
        d = json.load(open(os.path.abspath(SETTINGS), encoding="utf-8"))
    except (OSError, ValueError) as e:
        return {"ok": False, "why": f"settings.json 을 못 읽었다({type(e).__name__})"}
    auto = d.get("autoCompactEnabled")
    hooks = json.dumps(d.get("hooks") or {}, ensure_ascii=False)
    pre = "PreCompact" in hooks and "session_wrapup" in hooks
    ok = bool(auto is not False and pre)
    why = []
    if auto is False:
        why.append("autoCompactEnabled=false (자동 요약이 꺼져 있다)")
    if not pre:
        why.append("PreCompact 훅에 session_wrapup 이 안 걸려 있다 (요약 직전 인계가 안 된다)")
    return {"ok": ok, "why": " · ".join(why) or "자동 요약 + 요약 직전 인계 모두 걸려 있다"}


def steps():
    """도는 단계 — (이름, 인자, 제한초). 수집·엑셀쓰기·커밋은 여기 없다."""
    return [
        ("인계 갱신", [os.path.join(ROOT, "session_handoff.py"), "--check"], 600),
        ("합성검증", [os.path.join(ROOT, "tests", "synthetic_check.py")], 900),
        ("정기점검 내용 조사", [os.path.join(ROOT, "pm_content.py")], 600),
        ("캠프명 표준 대조(조사만)", [os.path.join(ROOT, "camp_standardize.py")], 900),
    ]


def run_steps(skip=()):
    """★ `subprocess.run(timeout=)` 을 쓰지 않는다 — 윈도우에서 영원히 안 끝날 수 있다([175])."""
    import proc_guard
    began = datetime.now()
    out = []
    for name, argv, limit in steps():
        if name in skip:
            out.append({"step": name, "state": "건너뜀",
                        "note": "다른 세션이 코드를 고치는 중 — 남의 빨강을 내 회차 사실로 적지 않는다"})
            continue
        left = BUDGET_MIN * 60 - (datetime.now() - began).total_seconds()
        if left <= 30:
            out.append({"step": name, "state": "건너뜀",
                        "note": "회차 예산을 넘었다 — 다음 회차가 이어서 합니다"})
            continue
        r = proc_guard.run_tree([sys.executable] + argv, cwd=ROOT,
                                timeout=min(limit, left), drain_timeout=10)
        tail = (r.stdout or r.stderr or "").strip().splitlines()
        out.append({"step": name, "state": "정상" if r.returncode == 0 else f"실패({r.returncode})",
                    "note": tail[-1][:160] if tail else ""})
    return out


def git_pending():
    """미커밋·미푸시는 **세어서 알리기만** 한다 — 작업 폴더는 세션끼리 공유한다([104])."""
    def g(args):
        try:
            p = subprocess.run(["git", "-C", ROOT] + args, capture_output=True,
                               text=True, encoding="utf-8", errors="replace",
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            return (p.stdout or "").strip()
        except Exception:
            return ""
    dirty = [x for x in g(["status", "--short"]).splitlines() if x.strip()]
    unpushed = g(["rev-list", "--count", "@{u}..HEAD"]) or "0"
    return {"미커밋": len(dirty), "미푸시": unpushed}


def write_report(verdict, results, compact, pending, now):
    os.makedirs(REPORT_DIR, exist_ok=True)
    bad = [r for r in results if r["state"].startswith("실패")]
    L = ["# 정오 회차 — 하루 한 번 (12:00~13:00 창)", "",
         f"- 실행: {now.isoformat(timespec='seconds')} · 판정: **{verdict['kind']}** — {verdict['why']}",
         f"- 컴팩팅 자동화: {'정상' if compact['ok'] else '★ 확인 필요'} — {compact['why']}",
         f"- 작업 폴더: 미커밋 {pending['미커밋']}개 · 미푸시 {pending['미푸시']}개 "
         f"(이 회차는 커밋하지 않는다 — 세션끼리 공유하는 폴더다)", ""]
    if results:
        L += ["| 단계 | 결과 | 비고 |", "|---|---|---|"]
        L += [f"| {r['step']} | {r['state']} | {r['note']} |" for r in results]
        L.append("")
    if bad:
        L += ["## ★ 실패한 단계 — 사람이 볼 것", ""]
        L += [f"- **{r['step']}** — {r['note']}" for r in bad]
        L.append("")
    if not verdict["go"]:
        L += ["이번 부름은 돌지 않았다. 스케줄러가 10분 뒤 다시 부른다"
              "(창 안에서 최대 여섯 번). 창을 넘기면 그날은 건너뛴다.", ""]
        # ★ 하루를 통째로 건너뛰는 것은 **조용히 넘어가면 안 되는 일**이다.
        #   양보는 옳지만, 매일 양보만 하다 한 번도 안 도는 것은 고장과 구별되지 않는다.
        start, end = window()
        if verdict["kind"] == "양보" and now.hour * 60 + now.minute >= end - 10:
            L.insert(1, f"> ★ **오늘({now.date().isoformat()}) 정오 회차를 건너뛴다** — "
                        f"창이 끝나도록 {verdict['why']}. 내일 창에서 다시 시도한다.")
    open(REPORT_MD, "w", encoding="utf-8").write("\n".join(L) + "\n")


def main():
    now = datetime.now()
    marker = load_marker()
    claims, rounds = live_claims(), running_rounds()
    verdict = decide(now, marker, claims, rounds, force="--force" in sys.argv)

    if "--status" in sys.argv:
        start, end = window()
        print(f"창 {start//60:02d}:{start%60:02d}~{end//60:02d}:{end%60:02d} · "
              f"오늘 완주 {marker.get('done_date') or '아직'} · 지금 판정 {verdict['kind']}: {verdict['why']}")
        print("컴팩팅 자동화:", compact_settings()["why"])
        return 0

    if not verdict["go"]:
        print(f"[{verdict['kind']}] {verdict['why']}")
        # ★ 양보·창밖은 마커에 **완주로 적지 않는다**(적으면 그날 회차가 영영 안 돈다).
        marker.setdefault("skips", [])
        marker["skips"] = (marker["skips"] + [{"at": now.isoformat(timespec="seconds"),
                                               "kind": verdict["kind"], "why": verdict["why"]}])[-12:]
        try:
            json.dump(marker, open(MARKER, "w", encoding="utf-8"), ensure_ascii=False)
        except OSError:
            pass
        write_report(verdict, [], compact_settings(), git_pending(), now)
        return 0

    results = run_steps(verdict.get("skip") or ())
    compact, pending = compact_settings(), git_pending()
    write_report(verdict, results, compact, pending, now)
    marker.update({"done_date": now.date().isoformat(),
                   "done_at": now.isoformat(timespec="seconds"),
                   "results": results, "compact_ok": compact["ok"]})
    try:
        json.dump(marker, open(MARKER, "w", encoding="utf-8"), ensure_ascii=False)
    except OSError:
        pass
    bad = [r["step"] for r in results if r["state"].startswith("실패")]
    print(f"정오 회차 완주 — 단계 {len(results)}개 · 실패 {len(bad)}"
          + (f" ({', '.join(bad)})" if bad else "")
          + f" · 컴팩팅 {'정상' if compact['ok'] else '★확인'}")
    print("리포트:", REPORT_MD)
    return 0


if __name__ == "__main__":
    sys.exit(main())
