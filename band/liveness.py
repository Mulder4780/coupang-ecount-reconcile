# -*- coding: utf-8 -*-
"""주입한 수집기가 **아직 살아 있는지 물어보고**, 수확이 **정말 늘었는지** 센다.

사고 #35 (2026-08-11 실측)
--------------------------------------------------------------------------------
20:49:05 `collect_step.ps1` 이 `댓글채우기_붙여넣기_90610953.js` 를 www.band.us 에
주입했다. 주입 자체는 **핑으로 증명된 진짜 주입**이었다(#34 를 고친 뒤였다).
21:21 `band_dump_state.js` 로 상태를 물으니 **`NO __GRAB`** — 전역이 사라졌다.
창 제목은 `게시글 : (주)유니버셜리프트 쿠팡AS | 밴드` 였다. 최상위 문서가 글 상세
페이지로 넘어갔고, 그 순간 주입 스크립트는 그 자리에서 죽었다.

결과: `90610953` 글 5,255 · `comments` 키 4,733 · **댓글 본문 0건.** 그 사이 덤프가
하나 흡수됐지만 옛 덤프 재병합(5255→5255)이라 **새 수확이 아니었다.**

★ 그래서 여러 세션이 "댓글 아직 안 긁었다"고 결론 냈다. 실제 상태는
  **"긁기 시작했다가 죽었다"** 였다. 캐시만 보면 그 둘이 똑같이 생겼다.

이 파일이 없애는 것은 그 구별 불가능이다.

무엇이 성공인가 (이 모듈의 유일한 계약)
--------------------------------------------------------------------------------
**주입은 성공이 아니다.** 붙여넣었다는 것은 '붙여넣었다'만 증명한다(#34).
**살아 있음도 성공이 아니다.** 살아서 아무것도 안 담을 수 있다([162] — 선택자가
한 칸 아래를 가리켜 250건을 '실패 0'으로 긁고도 댓글이 0건이었다).

성공이라 적으려면 **둘 다** 있어야 한다:
  ① 생존 확인 **한 번 이상** — 페이지가 `__grabStatus()` 로 답한 것
  ② 이 회차에 나온 덤프에서 **쓸 수 있는 댓글 수가 실제로 늘었을 것**
     쓸 수 있는 댓글 = 작성시각이 있는 댓글. 시각 없는 수확은 흡수기가 버리므로
     ([130]) 여기서도 세지 않는다 — 안 그러면 계기가 거짓말을 한다.

그리고 **0 을 뭉치지 않는다**([169]). 댓글 0건은 세 가지 뜻이다:
  · `확인된0개`  — 입력창까지 그려진 뒤 목록이 비었다 = **진짜 댓글이 없는 글**([199])
  · `미확인`     — 들여다본 적이 없다 = **안 본 것**
  · 덤프 자체가 없다 — 저장까지 못 갔다
셋을 갈라 말한다. 뭉치면 "0건"이 "다 봤다"로 읽힌다.

거짓 죽음을 조심한다 (#36 의 반대편)
--------------------------------------------------------------------------------
탐침은 **지금 활성 탭**에 붙는다(`inject_here.ps1`). 사람이 다른 band.us 탭을 보고
있으면 그 탭에는 전역이 없으니 `NO __GRAB` 이 온다 — 죽음과 똑같이 생겼다.
멀쩡히 돌고 있는 것을 죽었다고 하면 **사람이 멀쩡한 수집을 다시 돌리러 간다**(중복 수확).

가르는 근거는 **심장 소리**다. `grab_posts.js` 가 localStorage 에 `__grabBeat` 을
쓰고, localStorage 는 **출처가 같으면 탭이 달라도 같은 것**이라 어느 band.us 탭에서도
읽힌다. 그래서 전역이 없어도 심장 소리가 **싱싱하면** '다른 탭에서 살아 있다'이고,
**식었으면** 죽음이다. 시각 하나가 그 둘을 가른다.

쓰는 법
--------------------------------------------------------------------------------
  python band/liveness.py --watch                 # 지금 도는 것을 지켜본다(주입 안 함)
  python band/liveness.py --watch --minutes 60    # 60분까지
  python band/liveness.py --run --js band/댓글채우기_붙여넣기_90610953.js
                                                  # 주입 + 생존확인 + 수확 판정(수집 세션 몫)
  python band/liveness.py --status                # 마지막 회차 기록만
  python band/liveness.py --selftest              # 크롬 없이 판정 논리만 검사

종료코드:  0 수확확인·확인된0개   3 살아있으나 수확 없음   4 모름(탐침 실패)   5 죽음
  ★ `--run` 은 밴드 점유를 **빼앗지 않는다** — 다른 세션이 잡고 있으면 물러난다.
  ★ 긁는 것 자체는 'CSOS 리서치 및 자료 수집' 세션 몫이다. 코딩 세션은 `--watch`·
    `--status` 로 **읽기만** 한다.
"""
import argparse
import glob
import io
import json
import os
import sys
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# ★ 윈도우 콘솔은 cp949 다 — '—' 한 글자에 UnicodeEncodeError 로 죽는다. 실측으로
#   표를 다 찍고 **마지막 판정 한 줄에서** 터진 적이 있다(그 판정이 제일 중요한 줄이다).
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import dump_comment_count as dcc          # 수확을 세는 잣대는 하나다([130])

DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
STATE = os.path.join(ROOT, "reports", "밴드_수집생존.json")
PROBE_JS = os.path.join(HERE, "band_dump_state.js")
INJECT_PS = os.path.join(HERE, "inject_here.ps1")
FIND_TAB_PS = os.path.join(HERE, "inject_find_tab.ps1")

# `NO __GRAB` 는 **계약**이다 — browser_chain.py 와 inject_and_verify.ps1 이 죽음의
# 표식으로 그 문자열을 찾는다. 이름을 바꾸면 그 둘이 죽음을 못 알아본다([169] 의 모양).
DEAD_MARK = "NO __GRAB"

POLL_SEC = int(os.environ.get("COUPANG_LIVENESS_POLL_SEC", "180"))
WATCH_MIN = int(os.environ.get("COUPANG_LIVENESS_WATCH_MIN", "30"))
# 진척 없는 폴이 이만큼 잇따르면 '멈춤'이라 부른다. 1회로 부르면 느린 글 하나에도
# 멈춤이라 하고, 그러면 경보가 대부분이 되어 아무도 안 본다.
STALL_POLLS = 3
# 심장 소리가 이보다 오래되면 식은 것으로 본다. 폴 간격의 두 배가 기준이다 —
# 한 폴을 건너뛴 것만으로 죽었다고 하지 않는다.
BEAT_FRESH_SEC = max(180, POLL_SEC * 2)

ALIVE, ALIVE_OTHER, DEAD, NEVER, UNKNOWN = (
    "살아있음", "살아있음(다른탭)", "죽음", "한번도안함", "모름")


def _now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _load_state():
    try:
        with io.open(STATE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_state(d):
    """죽어도 남는 것이 요점이다([180]) — 폴마다 쓴다."""
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    tmp = STATE + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE)


# ── 다른 세션과 부딪히지 않는다 ──────────────────────────────────────────────
def band_held_by_other():
    """다른 세션이 'band' 를 잡고 있으면 그 설명을 돌려준다(없으면 None).

    판정을 새로 만들지 않는다 — `ai_claim` 의 주인·생사 판정을 그대로 빌린다.
    같은 판단을 두 곳에서 하면 언젠가 갈리고, 갈린 뒤에는 어느 쪽이 맞는지 모른다.
    """
    try:
        import ai_claim
        cur = (ai_claim.load() or {}).get("band")
        if not cur:
            return None
        if ai_claim._is_mine(cur, cur.get("who", "claude")):
            return None
        if ai_claim._is_dead(cur):
            return None
        return "%s[%s] · %s" % (cur.get("who", "?"), cur.get("sid", "?"),
                               cur.get("why", ""))
    except Exception as e:                      # 못 읽으면 '없음'으로 치지 않는다
        return "점유를 확인할 수 없다(%s)" % type(e).__name__


# ── ① 살아 있는지 물어본다 ───────────────────────────────────────────────────
def _clear_probe_files():
    for f in glob.glob(os.path.join(DOWNLOADS, "__grabstate__*.json")):
        try:
            os.remove(f)
        except Exception:
            pass


def _read_probe_file(wait_sec=14):
    for _ in range(int(wait_sec / 0.7) + 1):
        time.sleep(0.7)
        got = sorted(glob.glob(os.path.join(DOWNLOADS, "__grabstate__*.json")))
        if got:
            raw = ""
            try:
                with io.open(got[0], "r", encoding="utf-8") as f:
                    raw = f.read()
            except Exception:
                pass
            try:
                os.remove(got[0])
            except Exception:
                pass
            return raw
    return None


def _beat_age(beat):
    """심장 소리가 몇 초 전 것인가. 모르면 None."""
    if not isinstance(beat, dict):
        return None
    at = beat.get("at")
    try:
        return max(0.0, time.time() - float(at) / 1000.0)
    except Exception:
        return None


def classify(raw):
    """탐침이 가져온 원문을 판정으로 바꾼다. 여기 하나가 판정하는 곳이다."""
    s = {"때": _now_iso(), "판정": UNKNOWN, "상태": None, "죽음기록": None,
         "심장": None, "심장나이초": None, "쪽": None, "제목": None, "왜": ""}
    if raw is None:
        s["왜"] = "탐침이 아무것도 안 돌려줬다(콘솔이 닫혔거나 붙지 못했다)"
        return s
    try:
        d = json.loads(raw)
    except Exception:
        # 옛 탐침이 디스크에 남아 있을 수 있다 — 문자열로라도 죽음은 알아본다.
        s["왜"] = "탐침 응답을 읽을 수 없다"
        if DEAD_MARK in (raw or ""):
            s["판정"] = DEAD
            s["왜"] = "%s (옛 탐침 응답 — 근거가 약하다)" % DEAD_MARK
        return s

    s["쪽"] = d.get("href")
    s["제목"] = d.get("title")
    s["심장"] = d.get("beat")
    s["죽음기록"] = d.get("death")
    s["심장나이초"] = _beat_age(d.get("beat"))
    v = d.get("verdict")

    if v == "ALIVE" or isinstance(d.get("status"), dict):
        s["판정"] = ALIVE
        s["상태"] = d.get("status")
        return s

    fresh = (s["심장나이초"] is not None and s["심장나이초"] <= BEAT_FRESH_SEC)
    beat_running = bool(isinstance(d.get("beat"), dict) and d["beat"].get("running"))

    if v == "NEVER_STARTED":
        s["판정"] = NEVER
        s["왜"] = "전역도 심장 소리도 없다 — 이 출처에서 한 번도 시작하지 않았다"
        return s

    if v == "DIED_AFTER_START" or DEAD_MARK in (raw or ""):
        # ★ 여기가 거짓 죽음을 막는 자리다. 전역이 없다는 것은 **이 탭에** 없다는
        #   뜻일 뿐이다. 심장 소리가 싱싱하고 아직 돌고 있다면 수집기는 다른 탭에서
        #   살아 있다 — 그것을 죽음이라 부르면 사람이 멀쩡한 수집을 다시 돌린다(#36).
        if fresh and beat_running:
            s["판정"] = ALIVE_OTHER
            s["상태"] = d.get("beat")
            s["왜"] = ("이 탭에는 전역이 없지만 심장 소리가 %d초 전 것이다"
                       " — 수집기는 다른 band.us 탭에 살아 있다" % s["심장나이초"])
            return s
        s["판정"] = DEAD
        age = "모름" if s["심장나이초"] is None else "%d초 전" % s["심장나이초"]
        s["왜"] = "전역이 없다 · 마지막 심장 소리 %s" % age
        if isinstance(d.get("death"), dict):
            dd = d["death"]
            s["왜"] += " · 죽은 자리 %s (%s)" % (dd.get("title") or dd.get("href"),
                                             dd.get("why"))
        return s

    if v == "PROBE_ERROR":
        s["왜"] = "탐침이 페이지에서 오류를 만났다: %s" % d.get("err")
        return s
    s["왜"] = "탐침 응답에 판정이 없다(옛 탐침일 수 있다)"
    return s


def probe(host="band.us", timeout=240, find_tab=False, site_key="band-90610953"):
    """탐침을 붙여 상태를 받아 온다. 판정까지 붙여 돌려준다."""
    import proc_guard
    if not os.path.exists(PROBE_JS):
        return dict(classify(None), 왜="탐침 파일이 없다: %s" % PROBE_JS)
    ps = FIND_TAB_PS if (find_tab and os.path.exists(FIND_TAB_PS)) else INJECT_PS
    _clear_probe_files()
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", ps, "-Js", PROBE_JS, "-SiteKey", site_key]
    r = proc_guard.run_tree(cmd, timeout=timeout, cwd=ROOT)
    out = (getattr(r, "stdout", "") or "") + (getattr(r, "stderr", "") or "")
    # ★ 종료코드만 믿지 않는다 — Write-Output 으로 끝나는 .ps1 은 $LASTEXITCODE 를
    #   건드리지 않아 성공이 실패로 읽힌 적이 있다(#36 (나)). 말한 것으로 판단한다.
    if "INJECTED on" not in out:
        tail = (out.strip().splitlines() or [""])[-1]
        s = classify(None)
        s["왜"] = "탐침을 붙이지 못했다 — %s" % tail[:140]
        return s
    s = classify(_read_probe_file())
    s["탐침"] = os.path.basename(ps)
    return s


# ── ② 수확이 정말 늘었나 ─────────────────────────────────────────────────────
def harvest_since(t0, dirs=None):
    """t0 이후에 나온 덤프에서 **쓸 수 있는 댓글**이 몇 건 늘었나.

    세는 잣대는 `dump_comment_count.count()` 하나다 — 흡수기와 같은 기준(시각이 있는
    댓글만)이라 계기와 캐시가 갈리지 않는다([130]).
    """
    dirs = list(dirs or [DOWNLOADS])
    tot = {"글": 0, "댓글담김": 0, "댓글수": 0, "확인된0개": 0, "미확인": 0,
           "없는글": 0, "실패": 0, "껍데기": 0, "시각없음": 0}
    files = []
    for d in dirs:
        for p in glob.glob(os.path.join(d, "dump_*.json")):
            try:
                if os.stat(p).st_mtime < t0:
                    continue
            except OSError:
                continue
            try:
                r = dcc.count(p)
            except Exception as e:
                files.append({"파일": os.path.basename(p),
                              "왜": "읽지 못했다(%s)" % type(e).__name__})
                continue
            for k in tot:
                tot[k] += int(r.get(k, 0) or 0)
            files.append(dict({"파일": os.path.basename(p)}, **r))
    return {"합계": tot, "덤프": files, "덤프수": len([f for f in files if "글" in f])}


# ── 판정: 성공이라 적을 수 있는가 ────────────────────────────────────────────
def verdict(samples, harvest):
    """(종료코드, 한마디, 설명) — 성공은 생존확인 **과** 수확증가가 다 있을 때만."""
    seen = [s.get("판정") for s in samples]
    alive = [s for s in samples if s.get("판정") in (ALIVE, ALIVE_OTHER)]
    tot = (harvest or {}).get("합계") or {}
    n_dump = (harvest or {}).get("덤프수") or 0
    댓글 = int(tot.get("댓글수", 0) or 0)
    영개 = int(tot.get("확인된0개", 0) or 0)
    미확인 = int(tot.get("미확인", 0) or 0)

    if not alive:
        if DEAD in seen:
            why = next((s.get("왜") for s in samples if s.get("판정") == DEAD), "")
            return 5, "죽음", (
                "수집기가 살아 있는 것을 **한 번도 확인하지 못했다** — %s.\n"
                "  이것은 수집이 아니다. 성공으로 적지 말 것(#35).\n"
                "  최상위 문서가 글 상세로 넘어가면 주입 스크립트는 그 자리에서 죽는다." % why)
        if NEVER in seen:
            return 5, "한번도안함", (
                "전역도 심장 소리도 없다 — 주입이 시작조차 못 했다.\n"
                "  '긁다가 죽은 것'과는 다르다. 붙여넣기부터 다시 보라.")
        why = next((s.get("왜") for s in samples if s.get("왜")), "")
        return 4, "모름", (
            "살아 있는지 **확인하지 못했다** — %s.\n"
            "  모르는 것을 실패라고도, 성공이라고도 적지 않는다(#36).\n"
            "  수확은 캐시로 다시 확인할 것." % why)

    last = alive[-1].get("상태") or {}
    한 = "생존확인 %d회" % len(alive)
    if isinstance(last, dict) and last.get("ok") is not None:
        한 += " · 마지막 %s건 처리(총 %s)" % (last.get("ok"), last.get("total"))

    if 댓글 > 0:
        return 0, "수확확인", (
            "%s · 이 회차 덤프 %d개에서 **쓸 수 있는 댓글 %d건**(글 %d개)이 새로 담겼다.\n"
            "  성공의 조건 둘을 다 지났다: 생존확인 + 수확증가." % (한, n_dump, 댓글, tot.get("댓글담김", 0)))
    if 영개 > 0 and 미확인 == 0:
        return 0, "확인된0개", (
            "%s · 새 댓글은 0건이지만 **확인된 0개가 %d글**이다 — 입력창까지 그려진 뒤\n"
            "  목록이 비었다는 뜻이라 **진짜 댓글이 없는 글**로 본다([199]).\n"
            "  이것은 진척이다(다음 회차가 같은 글을 다시 안 뽑는다)." % (한, 영개))
    if n_dump == 0:
        return 3, "살아있음_덤프없음", (
            "%s · 그런데 **이 회차에 덤프가 한 개도 안 나왔다.**\n"
            "  살아 있는 것과 수확한 것은 다르다 — 성공으로 적지 않는다.\n"
            "  중간 저장(saveEvery)이 도는지, 다운로드가 막혀 있지 않은지 보라." % 한)
    return 3, "살아있음_수확없음", (
        "%s · 덤프 %d개가 나왔지만 쓸 수 있는 댓글은 0건이고 **미확인이 %d글**이다.\n"
        "  '댓글이 없다'가 아니라 **'못 읽었다'** 쪽을 먼저 의심하라([162]·[169]) —\n"
        "  항목 선택자가 한 칸 아래를 가리키면 실패 0 으로 긁고도 댓글이 0건이 된다.\n"
        "  성공으로 적지 않는다." % (한, n_dump, 미확인))


# ── 회차 본체 ────────────────────────────────────────────────────────────────
def watch(minutes=WATCH_MIN, interval=POLL_SEC, host="band.us", t0=None,
          find_tab=False, label="watch", site_key="band-90610953"):
    """살아 있는지 **주기적으로** 물어본다. 폴마다 자국을 남긴다."""
    t0 = time.time() if t0 is None else t0
    deadline = time.time() + minutes * 60
    st = _load_state()
    run = {"시작": _now_iso(), "무엇": label, "간격초": interval,
           "한도분": minutes, "폴": []}
    st["마지막회차"] = run
    _save_state(st)

    stall = 0
    prev_tried = None
    while True:
        s = probe(host=host, find_tab=find_tab, site_key=site_key)
        stat = s.get("상태") or {}
        tried = stat.get("tried")
        if tried is None and isinstance(stat, dict):
            # 옛 수집기는 tried 를 안 내놓는다 — 세 개를 더해 같은 뜻을 만든다.
            got = [stat.get(k) for k in ("ok", "missing", "failed")]
            if all(isinstance(x, int) for x in got):
                tried = sum(got)
        s["진척"] = tried
        if s.get("판정") in (ALIVE, ALIVE_OTHER):
            if tried is not None and tried == prev_tried and not stat.get("paused"):
                stall += 1
            else:
                stall = 0
            prev_tried = tried
        run["폴"].append(s)
        run["끝"] = _now_iso()
        _save_state(st)

        mark = {ALIVE: "살아있음", ALIVE_OTHER: "살아있음(다른탭)",
                DEAD: "★죽음", NEVER: "★시작안함", UNKNOWN: "모름"}.get(s["판정"], s["판정"])
        print("  %s  %-16s %s" % (s["때"][-8:], mark,
                                  ("%s건 처리" % tried) if tried is not None else (s.get("왜") or "")[:70]))

        if s["판정"] in (DEAD, NEVER):
            break                                    # 더 물어볼 것이 없다
        if isinstance(stat, dict) and stat.get("running") is False:
            print("  수집기가 스스로 끝났다고 말한다 — 수확을 센다.")
            break
        if stall >= STALL_POLLS:
            s["왜"] = ("진척이 %d번 연속 그대로다(멈춤) — 죽은 것은 아니지만"
                       " 나아가지도 않는다" % stall)
            print("  ★ %s" % s["왜"])
            run["멈춤"] = True
            _save_state(st)
            break
        if time.time() >= deadline:
            print("  한도(%d분)에 닿았다 — 지켜보기를 끝낸다(수집기는 계속 돈다)." % minutes)
            break
        time.sleep(max(5, min(interval, max(5, deadline - time.time()))))

    h = harvest_since(t0)
    rc, one, why = verdict(run["폴"], h)
    run["수확"] = h["합계"]
    run["덤프"] = [f.get("파일") for f in h["덤프"]]
    run["판정"] = one
    run["설명"] = why
    run["종료코드"] = rc
    st["마지막판정"] = {"때": _now_iso(), "판정": one, "종료코드": rc, "무엇": label}
    _save_state(st)

    print("")
    print("판정: %s" % one)
    print(why)
    print("기록: %s" % os.path.relpath(STATE, ROOT))
    return rc


def run_step(js, host="band.us", minutes=WATCH_MIN, interval=POLL_SEC, wait=45,
             site_key="band-90610953"):
    """주입 → 생존확인 → 수확 판정. 한 걸음이 여기서 끝난다."""
    import proc_guard
    if not os.path.exists(js):
        print("붙여넣기 파일이 없다: %s" % js)
        return 5
    other = band_held_by_other()
    if other:
        # 빼앗지 않는다 — 같은 밴드를 두 세션이 긁으면 캐시가 서로 덮어써 오염된다(사고 #27).
        print("밴드를 다른 세션이 잡고 있다 — 물러난다: %s" % other)
        return 4
    t0 = time.time()
    ps = FIND_TAB_PS if os.path.exists(FIND_TAB_PS) else INJECT_PS
    print("주입: %s (%s)" % (os.path.basename(js), os.path.basename(ps)))
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", ps, "-Js", js, "-SiteKey", site_key]
    r = proc_guard.run_tree(cmd, timeout=420, cwd=ROOT)
    out = (getattr(r, "stdout", "") or "") + (getattr(r, "stderr", "") or "")
    for line in out.strip().splitlines()[-6:]:
        print("  | %s" % line)
    if "INJECTED on" not in out:
        print("붙여넣지 못했다 — 수집이 아니다. 성공으로 적지 말 것.")
        return 5
    print("붙여넣었다. %d초 뒤부터 %d초마다 살아 있는지 묻는다." % (wait, interval))
    # 곧장 물으면 아직 전역이 없어 '죽었다'로 읽는다 — 시작할 틈을 준다.
    time.sleep(wait)
    return watch(minutes=minutes, interval=interval, host=host, t0=t0,
                 find_tab=True, label=os.path.basename(js), site_key=site_key)


def show_status():
    st = _load_state()
    run = st.get("마지막회차") or {}
    if not run:
        print("아직 기록이 없다 — `--watch` 나 `--run` 을 한 번 돌린 뒤에 보라.")
        return 0
    print("마지막 회차: %s ~ %s · %s" % (run.get("시작"), run.get("끝"), run.get("무엇")))
    for s in (run.get("폴") or [])[-8:]:
        print("  %s  %-16s %s" % ((s.get("때") or "")[-8:], s.get("판정"),
                                  ("%s건" % s.get("진척")) if s.get("진척") is not None
                                  else (s.get("왜") or "")[:70]))
    if run.get("판정"):
        print("판정: %s (종료코드 %s)" % (run["판정"], run.get("종료코드")))
        print(run.get("설명") or "")
    return 0


# ── 크롬 없이 판정 논리만 검사 ───────────────────────────────────────────────
def selftest():
    """판정이 정말 갈리는지 본다. 계기 자체가 눈멀 수 있다([169])."""
    ok = True

    def eq(name, got, want):
        nonlocal ok
        if got != want:
            ok = False
            print("  실패 %-34s 얻음 %r · 바람 %r" % (name, got, want))
        else:
            print("  통과 %-34s %r" % (name, got))

    now_ms = int(time.time() * 1000)
    alive_raw = json.dumps({"verdict": "ALIVE",
                            "status": {"running": True, "ok": 12, "total": 250,
                                       "tried": 12, "paused": False}})
    eq("살아 있으면 살아있음", classify(alive_raw)["판정"], ALIVE)

    # 전역 없음 + 식은 심장 = 죽음
    cold = json.dumps({"verdict": "DIED_AFTER_START",
                       "err": "NO __GRAB - script is dead (page navigated?)",
                       "beat": {"at": now_ms - 3600 * 1000, "running": True, "ok": 9},
                       "death": {"why": "beforeunload", "title": "게시글 : 쿠팡AS | 밴드"}})
    eq("식은 심장은 죽음", classify(cold)["판정"], DEAD)

    # 전역 없음 + 싱싱한 심장 = 다른 탭에서 살아 있다(거짓 죽음 방지, #36)
    warm = json.dumps({"verdict": "DIED_AFTER_START",
                       "err": "NO __GRAB - script is dead (page navigated?)",
                       "beat": {"at": now_ms - 5000, "running": True, "ok": 9}})
    eq("싱싱한 심장은 다른탭 생존", classify(warm)["판정"], ALIVE_OTHER)

    eq("한 번도 안 함", classify(json.dumps({"verdict": "NEVER_STARTED"}))["판정"], NEVER)
    eq("탐침 무응답은 모름", classify(None)["판정"], UNKNOWN)
    # 옛 탐침(판정 없음)도 죽음은 알아본다
    eq("옛 탐침 문자열", classify('nope ' + DEAD_MARK)["판정"], DEAD)

    A = {"판정": ALIVE, "상태": {"ok": 12, "total": 250}}
    D = {"판정": DEAD, "왜": "전역이 없다"}
    U = {"판정": UNKNOWN, "왜": "탐침 실패"}

    def H(댓글=0, 영개=0, 미확인=0, n=1):
        return {"합계": {"댓글수": 댓글, "확인된0개": 영개, "미확인": 미확인,
                        "댓글담김": 1 if 댓글 else 0}, "덤프수": n, "덤프": []}

    eq("생존+수확 = 성공", verdict([A], H(댓글=7))[1], "수확확인")
    eq("생존+확인된0개 = 진척", verdict([A], H(영개=30))[1], "확인된0개")
    eq("생존+덤프없음 = 성공아님", verdict([A], H(n=0))[1], "살아있음_덤프없음")
    eq("생존+미확인만 = 성공아님", verdict([A], H(미확인=40))[1], "살아있음_수확없음")
    # ★ 이것이 이 파일의 계약이다 — 수확이 있어도 생존확인이 없으면 성공이 아니다.
    eq("죽음+수확 = 죽음", verdict([D], H(댓글=7))[1], "죽음")
    eq("모름 = 실패도 성공도 아님", verdict([U], H())[1], "모름")
    eq("죽음 종료코드", verdict([D], H())[0], 5)
    eq("모름 종료코드", verdict([U], H())[0], 4)
    eq("성공 종료코드", verdict([A], H(댓글=1))[0], 0)
    eq("성공 아님 종료코드", verdict([A], H(미확인=1))[0], 3)

    # 수확 세기: 시각 없는 댓글은 세지 않는다([130])
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        p = os.path.join(td, "dump_202608120101_90610953.json")
        with io.open(p, "w", encoding="utf-8") as f:
            json.dump({"band": "90610953", "posts": {
                "1": {"comments": [{"author": "가", "created_at": 1, "content": "취소요청"}]},
                "2": {"comments": [{"author": "나", "created_at": None, "content": ""}]},
                "3": {"comments": [], "comments_full": True},
                "4": {},
            }}, f, ensure_ascii=False)
        h = harvest_since(0, dirs=[td])
        eq("시각 있는 댓글만 센다", h["합계"]["댓글수"], 1)
        eq("확인된 0개는 따로", h["합계"]["확인된0개"], 1)
        eq("미확인은 따로", h["합계"]["미확인"], 2)
        eq("t0 뒤 덤프만", harvest_since(time.time() + 60, dirs=[td])["덤프수"], 0)

    print("\n%s" % ("ALL GREEN — 판정이 갈린다." if ok else "★ 실패가 있다 — 위를 보라."))
    return 0 if ok else 1


def main():
    ap = argparse.ArgumentParser(description="밴드 수집기 생존확인 · 수확 판정")
    ap.add_argument("--watch", action="store_true", help="지켜보기만(주입 안 함)")
    ap.add_argument("--run", action="store_true", help="주입 + 생존확인 + 수확 판정")
    ap.add_argument("--js", help="--run 이 붙여넣을 파일")
    ap.add_argument("--host", default="band.us")
    ap.add_argument("--site", choices=("band-90610953", "band-84789192"),
                    default="band-90610953",
                    help="현재 Chrome 전면 탭과 정확히 일치해야 하는 허용 페이지")
    ap.add_argument("--minutes", type=int, default=WATCH_MIN)
    ap.add_argument("--interval", type=int, default=POLL_SEC)
    ap.add_argument("--status", action="store_true", help="마지막 기록만")
    ap.add_argument("--selftest", action="store_true", help="크롬 없이 판정 논리 검사")
    a = ap.parse_args()

    if a.selftest:
        return selftest()
    if a.status:
        return show_status()
    if a.run:
        if not a.js:
            print("--run 에는 --js <붙여넣기 파일> 이 필요하다")
            return 2
        return run_step(a.js, host=a.host, minutes=a.minutes, interval=a.interval,
                        site_key=a.site)
    if a.watch:
        return watch(minutes=a.minutes, interval=a.interval, host=a.host,
                     site_key=a.site)
    return show_status()


if __name__ == "__main__":
    raise SystemExit(main())
