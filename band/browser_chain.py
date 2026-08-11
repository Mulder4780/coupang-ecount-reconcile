# -*- coding: utf-8 -*-
"""브라우저가 있어야만 되는 수집을 **사람 손 없이** 이어 간다 (2026-08-11 지시).

사용자 지시: "내가 손 안대고 할 수 있는 방법 찾아 해결해"

빠져 있던 것은 '주입하는 기능'이 아니었다 — 그건 이미 있었다(`inject_here.ps1`).
빠진 것은 **다음 차례를 시작해 줄 사람**이다. 밴드 백필이 00:11 에 끝나도 그때
ERP 몰이를 이어서 할 사람이 없으면 아침까지 아무 일도 안 일어난다. 대화에 남긴
것은 사라지고 스케줄러에 넣은 것만 산다 — 그래서 이것은 도구가 아니라 **회차**다.

## 이 회차가 지키는 것

- **한 번에 한 가지만 시작하고 끝난다.** 지켜보지 않는다. 다음 tick 이 다시 본다.
  한 tick 이 몇 시간을 물고 있으면 그것이 바로 [175]·[180] 이 기록한 사고다.
- **먼저 '지금 뭔가 돌고 있나'를 값싸게 묻는다.** 덤프 파일 나이로 본다 — 주입
  탐침은 DevTools 를 토글하고 20초를 쓰므로 15분마다 하기엔 비싸고 시끄럽다.
  파일이 애매할 때만 탐침한다(비싼 확인은 값싼 확인 **뒤**에 온다, [168]).
- **성공을 거짓으로 적지 않는다.** 주입은 '붙여넣었다'까지만 증명한다. 살아 있는지
  물어 답을 받은 것만 '시작됨'이고, 수확은 언제나 캐시가 센다([162]·#35).
- **읽기 전용 수집만 무인으로 돌린다.** `--apply`·`--queue` 같은 쓰기 경로는 반쯤
  성공했을 수 있어 자동 재실행하지 않는다([190] 과 같은 이유).
- **자국을 남긴다.** 죽어도 남는 것이 요점이다([180]).

## 무엇을 이어 하나 (우선순위)

  ① 밴드 댓글 백필 — 남은 것이 있으면. 취소 댓글이 오늘 숫자를 바꾼다([177]).
  ② ERP 전화면 몰이 — 실패가 남아 있고 오늘 다시 안 해 봤으면.

쓰는 법:  python band/browser_chain.py            (한 tick)
          python band/browser_chain.py --status   (지금 상태만)
"""
import argparse
import glob
import io
import json
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

STATE = os.path.join(ROOT, "reports", "브라우저_체인.json")
LOCK = os.path.join(ROOT, "reports", ".browser_chain.lock")
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")

# 한 회차가 도는 중이라고 볼 시간. 밴드 한 회차는 250건×3.1초 ≈ 13분이고 회차
# 사이에 저장이 끼므로 20분을 넘겨 조용하면 그 수집은 끝났거나 죽은 것이다.
BUSY_MIN = 20
# 같은 일을 다시 시작하기 전에 두는 간격. 실패가 반복되면 15분마다 크롬 포커스를
# 빼앗아 사람이 쓰지도 못하게 된다 — 경보가 대부분이면 아무도 안 본다([170]).
RETRY_MIN = 90

# ⚠ ERP 는 더 길게 둔다. `erp_unfinished()` 가 읽는 상태 파일은 **스스로 안 바뀐다** —
#   몰이의 결과는 브라우저 안 `window.__ERPALL` 에만 있고, 주입 45초 뒤 탐침은 아직
#   안 끝난 회차를 본다(전역이 막 새로 만들어져 '실패 0' 으로 보이기도 한다). 그
#   숫자로 '다 됐다'를 판정하면 **끝나지도 않은 것을 끝났다고 적는** 자리가 된다.
#   그래서 지금은 판정하지 않고 **간격만 넓힌다** — 하룻밤에 한두 번이면 충분하고,
#   결과를 거두는 것(상태 파일 갱신)은 아직 사람/다음 세션 몫이다. 모르는 것을
#   아는 것처럼 적지 않는 편이 낫다.
ERP_RETRY_MIN = 240


def _now():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def load():
    try:
        with io.open(STATE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"단계": [], "마지막": {}}


def save(d):
    d["갱신"] = _now()
    tmp = STATE + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE)


def note(d, 무엇, 결과, 왜=""):
    """단계마다 자국을 남긴다 — **죽어도 남는 것이 요점이다**([180])."""
    d.setdefault("단계", []).append(
        {"때": _now(), "무엇": 무엇, "결과": 결과, "왜": 왜})
    d["단계"] = d["단계"][-40:]
    d.setdefault("마지막", {})[무엇] = {"때": _now(), "결과": 결과, "왜": 왜}
    save(d)
    print("%s · %s — %s %s" % (_now(), 무엇, 결과, 왜))


def lock_take():
    """앞 tick 이 아직 돌고 있으면 겹치지 않는다. 죽은 잠금은 즉시 회수한다."""
    try:
        with io.open(LOCK, "r", encoding="utf-8") as f:
            pid = int((f.read() or "0").strip() or 0)
        alive = False
        if pid:
            import proc_guard
            alive = proc_guard.pid_alive(pid) if hasattr(proc_guard, "pid_alive") \
                else _pid_alive(pid)
        if alive:
            return False
    except Exception:
        pass
    with io.open(LOCK, "w", encoding="utf-8") as f:
        f.write(str(os.getpid()))
    return True


def _pid_alive(pid):
    import subprocess
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", "PID eq %d" % pid], text=True, errors="replace")
        return str(pid) in out
    except Exception:
        return False


def lock_free():
    try:
        os.remove(LOCK)
    except Exception:
        pass


# ── ① 지금 뭔가 돌고 있나 — 값싼 확인이 먼저다 ────────────────────────────────
def newest_dump_age_min():
    """가장 최근 덤프가 몇 분 전 것인가. 없으면 None.

    Downloads 와 Z: 를 **둘 다** 본다 — 자동화 파이프라인(5분)이 Downloads 를
    비워 가므로 거기만 보면 '수확이 없었다'로 잘못 읽는다(실측 2026-08-11).
    """
    paths = glob.glob(os.path.join(DOWNLOADS, "dump_*.json"))
    try:
        import source_dirs
        paths += glob.glob(os.path.join(source_dirs.BAND_DIR, "**", "dump_*.json"),
                           recursive=True)
    except Exception:
        pass
    if not paths:
        return None
    newest = max(os.path.getmtime(p) for p in paths)
    return (time.time() - newest) / 60.0


def looks_busy():
    age = newest_dump_age_min()
    if age is None:
        return False, "덤프가 하나도 없다"
    if age < BUSY_MIN:
        return True, "가장 최근 덤프가 %.0f분 전 — 수집이 도는 중으로 본다" % age
    return False, "가장 최근 덤프가 %.0f분 전 — 조용하다" % age


# ── ② 무엇을 이어 할까 ────────────────────────────────────────────────────────
def band_remaining():
    """댓글 백필에 남은 건수. 못 세면 None(모르면 1순위라고 우기지 않는다, [177])."""
    try:
        sys.path.insert(0, HERE)
        import comment_backfill as cb
        opens = cb.open_projects()
        n = 0
        for b in cb.bands():
            n += len(cb.blind(b, None, opens))
        return n
    except Exception as e:
        print("   밴드 남은 건수를 못 셌다 — %s: %s" % (type(e).__name__, e))
        return None


def erp_unfinished():
    """ERP 몰이에 실패가 남아 있나. 근거는 마지막 실측 상태 파일이다."""
    p = os.path.join(ROOT, "reports", "ERP_전화면몰이_상태.json")
    try:
        with io.open(p, "r", encoding="utf-8") as f:
            d = json.load(f)
        done = d.get("끝난것") or []
        if isinstance(done, str):
            done = json.loads(done)
        return sum(1 for r in done if r.get("결과") == "실패")
    except Exception:
        return None


def too_soon(d, 무엇, 분=None):
    분 = RETRY_MIN if 분 is None else 분
    last = (d.get("마지막") or {}).get(무엇)
    if not last:
        return False
    try:
        t = time.mktime(time.strptime(last["때"], "%Y-%m-%d %H:%M:%S"))
    except Exception:
        return False
    return (time.time() - t) / 60.0 < 분


# ── ③ 주입하고 **살아 있는지 확인한다** ───────────────────────────────────────
def inject(js, host, probe, dead_marks):
    """붙여넣고 → 잠시 뒤 상태를 물어본다. 확인 못 한 주입은 수집이 아니다(#35).

    반환: (성공?, 설명)
    """
    import proc_guard
    ps = os.path.join(HERE, "inject_find_tab.ps1")
    if not os.path.exists(js):
        return False, "붙여넣기 파일이 없다: %s" % os.path.basename(js)
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", ps, "-Js", js, "-ExpectHost", host]
    r = proc_guard.run_tree(cmd, timeout=420, cwd=ROOT)
    out = (getattr(r, "stdout", "") or "") + (getattr(r, "stderr", "") or "")
    if "INJECTED on" not in out:
        tail = out.strip().splitlines()[-1:] or [""]
        return False, "붙여넣지 못했다 — %s" % tail[0][:120]

    # 시작할 틈을 준다. 여기서 곧장 물으면 아직 전역이 없어 '죽었다'로 읽는다.
    time.sleep(45)
    for f in glob.glob(os.path.join(DOWNLOADS, "__grabstate__*.json")) + \
             glob.glob(os.path.join(DOWNLOADS, "__erp__*.txt")):
        try:
            os.remove(f)
        except Exception:
            pass
    cmd2 = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", os.path.join(HERE, "inject_here.ps1"),
            "-Js", probe, "-ExpectHost", host]
    proc_guard.run_tree(cmd2, timeout=240, cwd=ROOT)

    state = None
    for _ in range(20):
        time.sleep(0.7)
        got = glob.glob(os.path.join(DOWNLOADS, "__grabstate__*.json")) + \
              glob.glob(os.path.join(DOWNLOADS, "__erp__*.txt"))
        if got:
            # ERP 탐침은 **파일 이름**에 상태를 싣는다(내용은 한 글자다).
            state = os.path.basename(got[0])
            try:
                if got[0].endswith(".json"):
                    state = io.open(got[0], encoding="utf-8").read()[:200]
                os.remove(got[0])
            except Exception:
                pass
            break
    if state is None:
        # ★ 모르는 것을 실패라고 부르지 않는다 — 탐침이 못 붙은 것과 스크립트가
        #   죽은 것은 다르다. 뭉치면 사람이 멀쩡한 수집을 다시 돌리러 간다(#36).
        return True, "붙여넣었으나 상태를 못 읽었다(탐침 실패) — 수확은 캐시로 확인할 것"
    for m in dead_marks:
        if m in state:
            return False, "붙여넣었지만 스크립트가 없다(%s) — 시작 못 했거나 화면이 넘어갔다" % m
    return True, "살아 있다 · %s" % state.replace("\n", " ")[:160]


# ── 회차 본체 ────────────────────────────────────────────────────────────────
def tick():
    d = load()
    busy, why = looks_busy()
    if busy:
        note(d, "점검", "건너뜀", why)
        return 0

    band_n = band_remaining()
    erp_n = erp_unfinished()
    print("남은 것 — 밴드 댓글 %s · ERP 실패 %s · %s"
          % ("모름" if band_n is None else band_n,
             "모름" if erp_n is None else erp_n, why))

    # ① 밴드가 먼저다 — 취소 댓글이 오늘 숫자를 바꾼다([177]).
    if band_n:
        for js in sorted(glob.glob(os.path.join(HERE, "댓글채우기_붙여넣기_*.js"))):
            무엇 = "밴드 댓글 " + os.path.basename(js).split("_")[-1][:-3]
            if too_soon(d, 무엇):
                continue
            ok, msg = inject(js, "band.us",
                             os.path.join(HERE, "band_dump_state.js"),
                             ["NO __GRAB"])
            note(d, 무엇, "시작됨" if ok else "실패", msg)
            return 0 if ok else 1

    # ② ERP — 실패가 남아 있으면 다시 몰아 본다.
    if erp_n:
        무엇 = "ERP 전화면 몰이"
        if not too_soon(d, 무엇, ERP_RETRY_MIN):
            ok, msg = inject(os.path.join(HERE, "ERP_전화면몰이_붙여넣기.js"),
                             "ecount", os.path.join(HERE, "erp_status_ping.js"),
                             ["NOSTATE"])
            note(d, 무엇, "시작됨" if ok else "실패", msg)
            return 0 if ok else 1

    note(d, "점검", "할 일 없음",
         "밴드 %s · ERP %s" % (band_n, erp_n))
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", action="store_true", help="지금 상태만 본다")
    a = ap.parse_args()
    if a.status:
        d = load()
        print("갱신", d.get("갱신"))
        for k, v in (d.get("마지막") or {}).items():
            print("  %-22s %s · %s · %s" % (k, v.get("때"), v.get("결과"), v.get("왜")))
        busy, why = looks_busy()
        print("지금:", "수집 중" if busy else "조용함", "—", why)
        return 0
    if not lock_take():
        print("앞 회차가 아직 돌고 있다 — 건너뛴다")
        return 0
    try:
        return tick()
    finally:
        lock_free()


if __name__ == "__main__":
    raise SystemExit(main())
