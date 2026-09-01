# -*- coding: utf-8 -*-
"""브라우저가 있어야만 되는 수집을 **사람 손 없이** 이어 간다 (2026-08-11 지시).

사용자 지시: "내가 손 안대고 할 수 있는 방법 찾아 해결해"

빠져 있던 것은 '주입하는 기능'이 아니었다 — 그건 이미 있었다(`inject_here.ps1`).
빠진 것은 **다음 차례를 시작해 줄 사람**이다. 밴드 백필이 00:11 에 끝나도 그때
ERP 몰이를 이어서 할 사람이 없으면 아침까지 아무 일도 안 일어난다. 대화에 남긴
것은 사라지고 스케줄러에 넣은 것만 산다 — 그래서 이것은 도구가 아니라 **회차**다.

## 이 회차가 지키는 것

- **한 번에 한 가지만 시작하고 끝난다.** 지켜보지 않는다. 다음 정오 회차나 형님의
  명시적 수동 명령이 다시 본다.
  한 tick 이 몇 시간을 물고 있으면 그것이 바로 [175]·[180] 이 기록한 사고다.
- **먼저 '지금 뭔가 돌고 있나'를 값싸게 묻는다.** 덤프 파일 나이로 본다 — 주입
  탐침은 DevTools 를 토글하고 시간을 쓰므로 불필요하게 반복하면 비싸고 시끄럽다.
  파일이 애매할 때만 탐침한다(비싼 확인은 값싼 확인 **뒤**에 온다, [168]).
- **성공을 거짓으로 적지 않는다.** 주입은 '붙여넣었다'까지만 증명한다. 살아 있는지
  물어 답을 받은 것만 '시작됨'이고, 수확은 언제나 캐시가 센다([162]·#35).
- **읽기 전용 수집만 무인으로 돌린다.** `--apply`·`--queue` 같은 쓰기 경로는 반쯤
  성공했을 수 있어 자동 재실행하지 않는다([190] 과 같은 이유).
- **자국을 남긴다.** 죽어도 남는 것이 요점이다([180]).

## 무엇을 이어 하나 (우선순위)

  ① 밴드 댓글 백필 — 남은 것이 있으면. 취소 댓글이 오늘 숫자를 바꾼다([177]).
  ② ERP 전화면 몰이 — 실패가 남아 있고 · **하루 한 번** · **12~13시 창 안**이고 ·
     다른 세션이 크롬(band/publish)을 잡고 있지 않을 때만 (2026-08-12 지시).
     지정 ERP 페이지가 전면이 아니면 그날 예약 회차는 조용히 끝낸다. 추가 실행은
     형님이 명령한 `--manual erp`만 허용한다.

쓰는 법:  python band/browser_chain.py            (12:00 예약 한 회차)
          python band/browser_chain.py --manual band-90610953
          python band/browser_chain.py --manual band-84789192
          python band/browser_chain.py --manual erp
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
# 같은 일을 다시 시작하기 전에 두는 간격. 수동 명령을 연달아 잘못 전달하거나 옛
# 호출기가 남더라도 반복 입력으로 번지지 않게 한다.
RETRY_MIN = 90

# ── ERP 몰이는 **하루 한 번, 12~13시 창 안에서만** (2026-08-12 지시) ──────────
#
# 사용자 지시: "하루한번 12시에서 13시 사이 다른 세션이랑 충돌 안나게 실행 알고리즘
#              코딩해(자동화 100%, 컴팩팅도 자동화, 내가 손 안대게 처리)"
#
# 전에는 15분 tick 이 간격(4시간)만 보고 **하루 중 아무 때나** 몰이를 시작했다.
# 사고 #35 가 난 것도 그 회차의 21:10 이다. 몰이는 크롬 창을 차지하고 키보드를
# 빼앗으므로 아무 때나 시작하면 셋과 겹친다:
#   · 사람이 그 크롬으로 일하는 중 — 포커스를 통째로 빼앗는다
#   · 'CSOS 리서치 및 자료 수집' 세션의 밴드·ERP 수집 — 같은 크롬·같은 Z: 다
#   · 이 회차 자신의 밴드 댓글 백필 — 한 tick 은 하나만 시작하지만 앞 tick 이
#     아직 긁는 중일 수 있다
# 그래서 **시각으로 자리를 하나 정해 준다.** 점심시간이 그 자리다.
#
# ⚠ 결과를 판정하지는 않는다. `erp_unfinished()` 가 읽는 상태 파일은 **스스로 안
#   바뀐다** — 몰이의 결과는 브라우저 안 `window.__ERPALL` 에만 있고, 주입 45초 뒤
#   탐침은 아직 안 끝난 회차를 본다(전역이 막 새로 만들어져 '실패 0' 으로 보이기도
#   한다). 그 숫자로 '다 됐다'를 판정하면 **끝나지도 않은 것을 끝났다고 적는**
#   자리가 된다. 하루 한 번이라는 문은 시각이 지키고, 수확은 캐시가 센다.
ERP_WINDOW = os.environ.get("COUPANG_ERP_WINDOW", "12:00-13:00")
# 예약 회차는 12:00 딱 한 번이다. 이 값은 같은 프로세스 안에서 중복 시도를 막는
# 방어선이며, 추가 실행은 형님의 명시적 `--manual` 명령만 허용한다.
ERP_WINDOW_RETRY_MIN = 12
# 몰이가 크롬을 독점하므로 **다른 세션이 잡고 있으면 안 한다.** 판단은 새로 만들지
# 않고 `ai_claim` 것을 그대로 빌린다([225] 와 같은 이유 — 같은 판단을 두 곳에서
# 하면 언젠가 갈리고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다).
ERP_CONFLICT_LOCKS = ("band", "publish")


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
    trace_missed(d)          # ★ 자국을 남기는 자리와 판정하는 자리를 붙여 둔다([328])
    print("%s · %s — %s %s" % (_now(), 무엇, 결과, 왜))


#: 기회를 며칠 못 잡으면 그때부터 자국을 남긴다. 하루 이틀은 정상이다 —
#: 정오에 크롬이 앞에 없는 것은 흔하고, 매일 뜨는 경보는 진짜 경보를 죽인다([170]).
MISS_DAYS = 3
def _miss_trace():
    """자국 자리 — **`STATE` 에서 파생시킨다**([247]).

    ★ 검증이 `BC.STATE` 만 임시 폴더로 돌리고 `settle_band()` 를 부르는 자리가
      있다(t269). 자국 경로를 따로 못 박아 두면 그 검사가 **진짜 증거 파일을
      지운다** — 실측으로 그렇게 한 번 사라졌다. 파생시키면 격리가 저절로 따라온다.
    """
    return os.path.join(os.path.dirname(STATE), "브라우저수집_오류.json")


def _miss_days(d, 무엇):
    """그 일이 **마지막으로 된 뒤** 며칠인가 — 못 재면 None([169])."""
    from datetime import datetime as _dt, date as _date   # 이 파일에 없다([355])
    rec = (d.get("마지막") or {}).get(무엇) or {}
    when = str(rec.get("때") or "")[:10]
    if not when:
        return None
    try:
        y, m, dd = (int(x) for x in when.split("-"))
        return (_dt.now().date() - _date(y, m, dd)).days
    except Exception:
        return None


def trace_missed(d):
    """브라우저 수집이 기회를 못 잡고 있으면 `*_오류.json` 으로 말한다.

    ★ 판정은 **'마지막으로 시작된 날'** 하나다. `건너뜀` 을 세지 않는 이유는
      건너뜀이 정상이기 때문이다([269] — 정오에 ERP 화면이 전면에 없는 것은 정상).
      물어야 할 것은 "몇 번 건너뛰었나"가 아니라 **"며칠째 못 했나"** 다.
    ★ **자동으로 다시 시도하지 않는다.** 이 회차는 사람 화면을 뺏으므로 [269] 가
      매일 12:00 한 번으로 정해 뒀다 — 여기서 되돌리면 그 지시가 조용히 사라진다.
      여기는 **보고 말하는 자리**다(`typo_watch` 와 같은 자리).
    ★ 못 재면 아무 말도 안 한다([169]) — 모르는 것으로 경보를 만들면 그것이
      매일 뜨는 가짜가 된다.
    """
    늦은것 = []
    for 무엇 in sorted((d.get("마지막") or {})):
        if 무엇 == "점검":                      # 할 일이 없던 날의 자국이다
            continue
        n = _miss_days(d, 무엇)
        rec = (d.get("마지막") or {}).get(무엇) or {}
        # ★ '시작됨' 도 샌다 — 시작만 하고 끝을 못 본 날이 실측으로 있었다
        #   (2026-08-23: 12:12 '시작됨' 인데 결과 파일은 12일째 옛것).
        if n is not None and n >= MISS_DAYS:
            늦은것.append("%s %d일째(마지막 %s · %s)"
                          % (무엇, n, str(rec.get("때") or "")[:10], rec.get("결과")))
    if not 늦은것:
        # ★ 잡았으면 지운다 — 옛 자국은 고쳐진 고장을 계속 보고한다([228]).
        try:
            os.remove(_miss_trace())
        except Exception:
            pass
        return
    try:
        with io.open(_miss_trace(), "w", encoding="utf-8") as fh:
            json.dump({
                "시각": _now(),
                "갈래": "기회놓침",
                "작업": "CSOS_BrowserChain",
                "무엇": ("브라우저 수집이 %d일 넘게 기회를 못 잡았다 — %s"
                         % (MISS_DAYS, " · ".join(늦은것[:4]))),
                "어떻게": ("허용된 크롬 페이지(밴드 /post 목록 또는 이카운트 ERP)를 "
                           "**전면**에 두고: python band/browser_chain.py --manual erp"),
                "늦은것": 늦은것,
            }, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass                                    # 자국 하나 때문에 회차를 세우지 않는다


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
            ["tasklist", "/FI", "PID eq %d" % pid], text=True, errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
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
def band_remaining(write=False):
    """댓글 백필에 남은 건수. 못 세면 None(모르면 1순위라고 우기지 않는다, [177]).

    ★ `write=True` 면 **붙여넣기 파일을 다시 만든다.** 이게 없으면 회차가 수렴하지
    않는다: 파일은 정적인데 캐시 흡수가 늦으면 다음 tick 이 **이미 긁은 1,663건을
    통째로 다시** 긁는다. [199] 가 기록한 무한루프와 같은 모양이고, 화면에는
    '수집 중'만 보여 아무도 이상하다 안 한다. 목록을 매번 새로 뽑으면 이미 확정된
    글은 저절로 빠진다.
    """
    try:
        if HERE not in sys.path:
            sys.path.insert(0, HERE)
        import comment_backfill as cb
        opens = cb.open_projects()
        n = 0
        for b in cb.bands():
            rows = cb.blind(b, None, opens)
            n += len(rows)
            if write and rows:
                cb.write_paste(b, [no for _t, _d, no in rows])
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


# ── ②-b 시각의 문 · 세션의 문 (2026-08-12 지시) ──────────────────────────────
def _hhmm(s):
    """'12:00' → 720분. 못 읽으면 None."""
    try:
        h, m = str(s).strip().split(":")
        h, m = int(h), int(m)
        if 0 <= h <= 23 and 0 <= m <= 59:
            return h * 60 + m
    except Exception:
        pass
    return None


def window_bounds(spec=None):
    """'12:00-13:00' → (720, 780). 못 읽으면 None — **없음이 아니라 모름**이다."""
    parts = str(ERP_WINDOW if spec is None else spec).split("-")
    if len(parts) != 2:
        return None
    a, b = _hhmm(parts[0]), _hhmm(parts[1])
    if a is None or b is None or b <= a:
        return None
    return (a, b)


def window_now(spec=None, now=None):
    """지금이 창 안인가. 반환 (열림?, 설명).

    ★ 창 설정을 **못 읽으면 열어 준다** — 그리고 그렇게 적는다. 오타 하나로 회차가
      영원히 조용히 멈추는 편이 더 나쁘다: 안 도는 회차는 아무 화면에도 안 뜬다([169]).
    """
    spec_s = ERP_WINDOW if spec is None else spec
    t = time.localtime() if now is None else now
    cur = t.tm_hour * 60 + t.tm_min
    w = window_bounds(spec_s)
    지금 = "지금 %02d:%02d" % (t.tm_hour, t.tm_min)
    if w is None:
        return True, "창 설정을 못 읽었다(%r) — 시각 제한 없이 진행 · %s" % (spec_s, 지금)
    if w[0] <= cur < w[1]:
        return True, "창 안 (%s · %s)" % (spec_s, 지금)
    return False, "창 밖 (%s · %s)" % (spec_s, 지금)


def erp_day(d):
    """오늘치 ERP 회차 기록. 날짜가 바뀌면 새로 시작한다.

    **하루 한 번의 근거는 이 날짜 도장이지 '스케줄러가 떴다'가 아니다.** 스케줄러는
    무슨 일이 있었든 '성공'이라 적는다 — 이 프로젝트가 여러 번 데인 자리다([180]).
    """
    today = time.strftime("%Y-%m-%d")
    rec = d.get("ERP회차") or {}
    if rec.get("날짜") != today:
        rec = {"날짜": today, "시작됨": False, "시도": 0, "마지막왜": "", "창닫힘기록": ""}
        d["ERP회차"] = rec
    return rec


def other_session_holds(locks=ERP_CONFLICT_LOCKS):
    """다른 세션이 겹치는 자원을 잡고 있나. 반환 (막힘?, 설명).

    ★ **못 읽으면 막힌 것으로 본다.** '못 읽음'을 '비었음'으로 치면 두 세션이 같은
      크롬을 동시에 몰게 된다 — 그러면 둘 다 깨지고, 깨진 쪽은 조용하다([169]).
    ★ 죽은 세션의 점유는 `ai_claim._is_dead` 가 가른다. 여기서 다시 판정하지 않는다.
    """
    try:
        if ROOT not in sys.path:
            sys.path.insert(0, ROOT)
        import ai_claim
        claims = ai_claim.load() or {}
    except Exception as e:
        return True, "점유를 못 읽었다(%s: %s) — 안전하게 건너뛴다" % (type(e).__name__, e)
    me = ""
    try:
        me = ai_claim.session_id()
    except Exception:
        pass
    for what in locks:
        c = claims.get(what)
        if not isinstance(c, dict):
            continue
        try:
            if ai_claim._is_dead(c):
                continue                      # 주인이 죽었다 — 겹칠 상대가 없다
        except Exception:
            pass                              # 판정이 안 되면 살아 있는 것으로 본다
        sid = c.get("sid") or "옛형식"
        if me and sid == me:
            continue                          # 내 것(스케줄러 실행에서는 거의 없다)
        return True, "다른 세션이 '%s' 를 잡고 있다 — %s[%s] · %s" % (
            what, c.get("who", "?"), sid, (c.get("why") or "")[:40])
    return False, "겹치는 점유 없음"


def _guard_refused(msg):
    """전면이 허용 Chrome 페이지가 아니어서 관문이 **안전 거절**한 것인가([269]).

    이것은 실패가 아니라 건너뜀이다 — 정오에 ERP 화면이 전면에 없는 것은 정상이다.
    진짜 오류(`ABORT: no payload`·시간초과·powershell 죽음)와는 가른다: 그것은
    exit 1 로 남아야 스케줄러가 잡는다. 표식은 inject_find_tab.ps1 이 낸다.
    """
    m = msg or ""
    return ("safe Chrome context required" in m) or \
           ("verified Chrome tab refused" in m)


def erp_step(d, 무엇="ERP 전화면 몰이"):
    """ERP 몰이 한 번. 문을 값싼 것부터 연다([168]).

    반환: 0/1 이면 이 tick 은 그것으로 끝난다. None 이면 아무것도 안 했다.
    """
    rec = erp_day(d)

    # ⓪ 오늘 이미 시작했다 — '하루 한 번'이 지시다.
    if rec.get("시작됨"):
        return None

    열림, 창설명 = window_now()
    # ① 창 밖이면 안 한다. 다만 **창이 닫힌 채 오늘을 못 했으면 그 사실을 적는다** —
    #    조용히 넘어가면 '돌긴 했는데 왜 결과가 없나'가 된다([180]). 하루 한 줄이다.
    if not 열림:
        w = window_bounds()
        t = time.localtime()
        지났나 = bool(w) and (t.tm_hour * 60 + t.tm_min) >= w[1]
        if 지났나 and rec.get("창닫힘기록") != rec["날짜"]:
            rec["창닫힘기록"] = rec["날짜"]
            note(d, 무엇, "오늘 못 함", "%s · 시도 %d회 · 마지막 이유: %s" % (
                창설명, rec.get("시도", 0),
                rec.get("마지막왜") or "창 안에서 한 번도 시작하지 못했다"))
        return None

    # ② 창 안이라도 방금 어긋났으면 잠깐 둔다 — 다음 tick 이 15분 뒤에 다시 온다.
    if too_soon(d, 무엇, ERP_WINDOW_RETRY_MIN):
        return None

    # ③ 다른 세션이 크롬을 쓰고 있으면 **양보한다.** 창 안에 남은 tick 이 다시 본다.
    막힘, 왜 = other_session_holds()
    if 막힘:
        rec["마지막왜"] = 왜
        note(d, 무엇, "양보", "%s · %s" % (창설명, 왜))
        return 0

    rec["시도"] = rec.get("시도", 0) + 1
    ok, msg = inject(os.path.join(HERE, "ERP_전화면몰이_붙여넣기.js"),
                     "erp", os.path.join(HERE, "erp_status_ping.js"),
                     ["NOSTATE"])
    rec["시작됨"] = bool(ok)
    rec["마지막왜"] = msg
    if ok:
        note(d, 무엇, "시작됨", "%s · %d번째 시도 · %s" % (창설명, rec["시도"], msg))
        return 0
    # ★ 전면이 허용 페이지가 아니면 관문이 안전 거절한다 — 실패가 아니라 건너뜀이다
    #   ([269]). 정오에 ERP 화면이 전면에 없는 것은 정상이고, 이걸 exit 1 로 내면
    #   스케줄러가 매일 P0 경보를 울린다(경보가 매일이면 아무도 안 본다, [170]).
    #   남은 tick 이 창 안에서 다시 본다. 진짜 오류만 exit 1 로 남긴다.
    if _guard_refused(msg):
        note(d, 무엇, "건너뜀", "%s · %d번째 시도 · %s" % (창설명, rec["시도"], msg))
        return 0
    note(d, 무엇, "실패", "%s · %d번째 시도 · %s" % (창설명, rec["시도"], msg))
    return 1


# ── ③ 주입하고 **살아 있는지 확인한다** ───────────────────────────────────────
def inject(js, site_key, probe, dead_marks):
    """붙여넣고 → 잠시 뒤 상태를 물어본다. 확인 못 한 주입은 수집이 아니다(#35).

    반환: (성공?, 설명)
    """
    import proc_guard
    ps = os.path.join(HERE, "inject_find_tab.ps1")
    if not os.path.exists(js):
        return False, "붙여넣기 파일이 없다: %s" % os.path.basename(js)
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", ps, "-Js", js, "-SiteKey", site_key]
    r = proc_guard.run_tree(cmd, timeout=420, cwd=ROOT)
    out = (getattr(r, "stdout", "") or "") + (getattr(r, "stderr", "") or "")
    if "INJECTED on" not in out:
        tail = out.strip().splitlines()[-3:] or [""]
        return False, "붙여넣지 못했다 — %s" % " | ".join(tail)[:240]

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
            "-Js", probe, "-SiteKey", site_key]
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


def _epoch(때):
    try:
        return time.mktime(time.strptime(때, "%Y-%m-%d %H:%M:%S"))
    except Exception:
        return None


def settle_band(d):
    """시작만 하고 **끝을 안 본** 밴드 수집을 주기적으로 물어 마무리한다 (분담판 [37]).

    ★ 지금까지는 붙여넣은 **직후 한 번만** 묻고 '시작됨'이라 적은 뒤 떠났다. 그런데
      250건 수집은 20분을 넘게 돈다 — 3분째에 죽어도(#35 의 그 죽음) 이 회차는 이미
      '시작됨'을 적어 뒀고, 다음 tick 은 `too_soon`(90분)에 걸려 건너뛴다.
      **성공도 실패도 아닌 채로 90분이 지나간다** — 화면에는 '시작됨'만 남아 있어
      아무도 이상하다 안 한다(`[169]` 와 같은 모양).
    ★ 판정은 `liveness.verdict` **한 곳**이 한다 — 여기서 다시 세면 두 판정이 갈리고,
      갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다. 이쪽 몫은 **'아직 이른가'** 뿐이다.
    ★ **막지 않는다.** tick 은 15분마다 오는데 여기서 `liveness.watch`(30분)를 돌면
      잠금을 쥔 채 다음 tick 을 굶긴다([175]). 한 tick 에 **한 번만** 묻고 넘어간다.
    ★ 밴드를 다른 세션이 잡고 있으면 물러난다(사고 #27) — 판정은 `liveness` 것을 빌린다.
    """
    pend = [(k, v) for k, v in (d.get("마지막") or {}).items()
            if k.startswith("밴드 댓글 ") and v.get("결과") in ("시작됨", "진행중")]
    if not pend:
        return
    try:
        if HERE not in sys.path:
            sys.path.insert(0, HERE)
        import liveness as L
    except Exception as exc:                       # 눈 하나 때문에 회차를 세우지 않는다
        print("생존확인을 못 불렀다: %s" % str(exc)[:60])
        return
    other = L.band_held_by_other()
    if other:
        print("밴드를 다른 세션이 잡고 있다 — 마무리는 다음 tick 으로: %s" % other)
        return
    for 무엇, v in pend:
        t0 = _epoch(v.get("때"))
        if t0 is None:
            continue
        지난분 = (time.time() - t0) / 60.0
        try:
            band_id = 무엇.rsplit(" ", 1)[-1]
            site_key = "band-%s" % band_id
            s = L.probe(host="band.us", find_tab=True, site_key=site_key)
            h = L.harvest_since(t0)
            code, 한마디, 왜 = L.verdict([s], h)
        except Exception as exc:
            print("%s 생존확인 실패: %s" % (무엇, str(exc)[:60]))
            continue
        if code == 3 and 지난분 < L.WATCH_MIN:
            # 살아 있는데 아직 수확이 없다 — 수집은 원래 오래 걸린다. 이것을 실패라
            # 부르면 멀쩡히 도는 수집을 사람이 다시 돌리러 간다(#36 의 거짓 죽음).
            note(d, 무엇, "진행중",
                 "살아 있고 아직 수확 전 — %.0f분째(한도 %d분)" % (지난분, L.WATCH_MIN))
        else:
            note(d, 무엇, 한마디, str(왜).replace("\n", " ")[:200])
        # ★ **재시도 시계를 되돌린다.** `note` 는 '마지막' 시각을 지금으로 갱신하는데,
        #   `too_soon` 이 그 시각을 보므로 마무리를 물을 때마다 90분이 다시 시작된다 —
        #   묻는 행위가 재시도를 영원히 미루는 자리다. 시계의 기준은 **붙여넣은 때**다.
        d["마지막"][무엇]["때"] = v["때"]
        save(d)


# ── 회차 본체 ────────────────────────────────────────────────────────────────
def tick():
    d = load()
    # ★ 새로 붙여넣기 **전에** 앞의 것을 마무리한다 — 순서가 뒤집히면 끝을 못 본
    #   수집 위에 또 하나를 얹는다.
    settle_band(d)
    busy, why = looks_busy()
    if busy:
        note(d, "점검", "건너뜀", why)
        return 0

    # 셀 때 파일도 같이 다시 만든다 — 목록과 파일이 갈리면 회차가 수렴하지 않는다.
    # ★ 밴드 수집 중단이면 **자동 갈래에서 밴드를 건너뛴다** (2026-09-01 지시).
    #   `manual_step` (형님 명령)은 안 막는다 — 사람이 그때 판단해 시키신 것이다([172]).
    try:
        from band import collect_switch as _CS
        _band_off, _band_why = _CS.stopped()
    except Exception:
        try:
            import collect_switch as _CS
            _band_off, _band_why = _CS.stopped()
        except Exception:
            _band_off, _band_why = False, ""
    band_n = band_remaining(write=True)
    erp_n = erp_unfinished()
    print("남은 것 — 밴드 댓글 %s · ERP 실패 %s · %s"
          % ("모름" if band_n is None else band_n,
             "모름" if erp_n is None else erp_n, why))

    # ① 밴드가 먼저다 — 취소 댓글이 오늘 숫자를 바꾼다([177]).
    if _band_off:
        band_n = 0                    # 조용히 빼지 않는다 — 아래 note 가 왜인지 적는다([169])
    if band_n:
        for js in sorted(glob.glob(os.path.join(HERE, "댓글채우기_붙여넣기_*.js"))):
            무엇 = "밴드 댓글 " + os.path.basename(js).split("_")[-1][:-3]
            if too_soon(d, 무엇):
                continue
            site_key = "band-" + os.path.basename(js).split("_")[-1][:-3]
            ok, msg = inject(js, site_key,
                             os.path.join(HERE, "band_dump_state.js"),
                             ["NO __GRAB"])
            # ★ 안전거절은 실패가 아니라 건너뜀이다([269]) — 판정은 erp_step 과
            #   같은 _guard_refused 하나를 쓴다(두 곳에서 가르면 언젠가 갈린다).
            #   '실패'로 적으면 이 자국을 읽는 화면·리포트가 "밴드 수집이 실패하고
            #   있다"고 말한다. 전면이 허용 Chrome 페이지가 아닌 것은 정상이고,
            #   매일 뜨는 가짜 경보는 진짜 경보까지 같이 죽인다([170]).
            note(d, 무엇,
                 "시작됨" if ok else ("건너뜀" if _guard_refused(msg) else "실패"),
                 msg)
            if ok:
                return 0
            # 틀린 허용 페이지가 전면이면 입력은 0회로 안전 거절된다. 하루 한 번
            # 회차에서는 나머지 두 허용 페이지도 읽기 전용으로 확인해 현재 페이지와
            # 맞는 한 곳을 찾는다. 탭 전환·포커스 이동은 절대 하지 않는다.
            continue

    # ② ERP — 하루 한 번, 정해진 창 안에서, 다른 세션과 겹치지 않을 때만.
    if erp_n:
        rc = erp_step(d)
        if rc is not None:
            return rc

    note(d, "점검", "할 일 없음",
         "밴드 %s · ERP %s · %s" % (band_n, erp_n, window_now()[1]))
    return 0


def manual_step(target):
    """형님 명령으로만 추가 실행. 시각은 우회해도 3페이지/Chrome 관문은 못 우회한다."""
    targets = {
        "band-90610953": (
            os.path.join(HERE, "댓글채우기_붙여넣기_90610953.js"),
            os.path.join(HERE, "band_dump_state.js"), ["NO __GRAB"]),
        "band-84789192": (
            os.path.join(HERE, "댓글채우기_붙여넣기_84789192.js"),
            os.path.join(HERE, "band_dump_state.js"), ["NO __GRAB"]),
        "erp": (
            os.path.join(HERE, "ERP_전화면몰이_붙여넣기.js"),
            os.path.join(HERE, "erp_status_ping.js"), ["NOSTATE"]),
    }
    js, probe, dead = targets[target]
    locks = ("band",) if target.startswith("band-") else ERP_CONFLICT_LOCKS
    blocked, why = other_session_holds(locks)
    if blocked:
        print("다른 세션이 사용 중이라 양보한다 — %s" % why)
        return 4
    ok, msg = inject(js, target, probe, dead)
    print("수동 실행 %s — %s" % ("시작됨" if ok else "실행 안 함", msg))
    if ok:
        return 0
    # ★ 안전거절은 실패가 아니다([269]) — 형님이 손으로 부른 자리에서 이것은
    #   고장이 아니라 "그 페이지를 앞으로 꺼내 달라"는 뜻이다. 1 로 내면 사람이
    #   멀쩡한 것을 고치러 간다([172]). 무엇을 하면 되는지까지 말한다.
    if _guard_refused(msg):
        print("  → 실패가 아니라 건너뜀이다. 허용된 Chrome 페이지를 **전면**에 두고"
              " 다시 부른다(밴드 /post 목록 또는 이카운트 ERP 화면).")
        return 0
    return 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", action="store_true", help="지금 상태만 본다")
    ap.add_argument("--manual", choices=("band-90610953", "band-84789192", "erp"),
                    help="사용자 명령으로 추가 실행(Chrome·정확한 페이지 관문은 유지)")
    a = ap.parse_args()
    if a.manual:
        return manual_step(a.manual)
    if a.status:
        d = load()
        print("갱신", d.get("갱신"))
        for k, v in (d.get("마지막") or {}).items():
            print("  %-22s %s · %s · %s" % (k, v.get("때"), v.get("결과"), v.get("왜")))
        busy, why = looks_busy()
        print("지금:", "수집 중" if busy else "조용함", "—", why)
        # ★ ERP 회차는 **왜 안 도는지**까지 말한다. '안 돎'만 보이면 사람이 손으로
        #   돌리러 가고, 그러면 창을 만든 뜻이 없다.
        rec = (d.get("ERP회차") or {})
        열림, 창설명 = window_now()
        막힘, 왜 = other_session_holds()
        print("ERP 몰이: %s · 오늘(%s) %s · 시도 %d회%s"
              % (창설명, rec.get("날짜") or "-",
                 "시작됨" if rec.get("시작됨") else "아직",
                 rec.get("시도", 0),
                 (" · 마지막 이유: " + rec["마지막왜"][:80]) if rec.get("마지막왜") else ""))
        print("        세션 충돌: %s" % 왜)
        if 열림 and not 막힘 and not rec.get("시작됨"):
            print("        → 지금 tick 이 돌면 시작한다")
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
