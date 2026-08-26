# -*- coding: utf-8 -*-
"""browser_bridge.py — **붙여넣기를 대신하는 다리** (2026-08-26 지시)

형님 지시: **"자동으로 붙이는 프로그램 만들어서 니가 알아서 밴드랑 ERP 데이터
긁어오게 코딩해"**

무엇이 막고 있었나
  수집기(`grab_posts.js`)를 로그인된 탭에 **넣는 것**이 사람 손이었다.  길이 셋인데
  셋 다 사람을 거쳤다 — ① Claude 확장 연결(계정에 매인다 · [362]) ② Tampermonkey
  설치(2026-08-19 부터 본문이 안 돈다 · [149]) ③ 콘솔 붙여넣기.
  그리고 붙여넣어도 **탭이 가려지면 멈춘다**(`document.hidden` 가드).

이 다리가 푸는 것
  크롬을 **디버깅 포트와 함께** 띄워 CDP(Chrome DevTools Protocol)로 붙는다.
  그러면 JS 를 **프로토콜로 직접** 넣는다 — 키보드를 흉내내지 않으므로 사람 화면을
  안 뺏는다([269] 의 정신은 그대로 지킨다).

★ `document.hidden` 가드를 **우회하지 않는다**
  2026-08-19 사고는 `defineProperty` 로 `hidden` 을 **거짓말**시킨 것이었고, 진짜
  범인은 **타이머 스로틀링**이었다(크롬은 백그라운드 탭의 타이머를 5분 뒤부터
  1분에 한 번으로 조인다).  그래서 본문 대기가 어긋나 **실재하는 글에 가짜 묘비**가
  박혔다 — 되돌릴 수 없는 쪽이다.

  여기서 쓰는 `Emulation.setFocusEmulationEnabled` 는 **크롬에게 그 탭을 실제로
  활성으로 다루라고 시키는 것**이다.  속이는 것이 아니라는 증거는 **타이머**다:

    | 방법 | hidden | 3초 타이머 |
    |---|---|---:|
    | 그냥(가려짐) | true | 3,376ms |
    | Page.bringToFront | true | 3,974ms |
    | **focusEmulation** | **false** | **3,011ms** |

  그리고 **8분을 재서** 5분 뒤 intensive throttling 이 안 걸리는 것을 확인했다
  (실측 5.3~7.8분 구간 3,004 / 3,009 / 3,013 / 3,008 / 3,450 / 3,014ms).
  ⚠ 그 사이 **10,783ms 가 한 번** 있었다 — 다른 무거운 작업과 겹친 것으로 **보이나
    확언하지 않는다**([169]).  수집기의 본문 대기가 12초라 이 정도는 견디지만,
    지연이 잦아지면 그때는 이 길을 다시 재야 한다.

★ **비밀번호를 대신 치지 않는다**
  로그인은 형님이 **한 번** 하신다.  전용 프로필은 계속 살아 있어 다시 하실 일이
  없다.  쿠키를 복제하거나 다른 프로필에서 훔쳐 오지 않는다(절대규칙).

  python band/browser_bridge.py --status         # 지금 상태만 (크롬 안 띄운다)
  python band/browser_bridge.py --up             # 크롬을 띄운다 (로그인하실 때)
  python band/browser_bridge.py --collect 90610953
  python band/browser_bridge.py --collect-all    # 대기열에 있는 밴드 전부
"""
import argparse
import json
import os
import subprocess
import sys
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

#: 디버깅 포트.  형님이 평소 쓰시는 크롬(포트 없음)과 **겹치지 않는다**.
PORT = int(os.environ.get("CSOS_BRIDGE_PORT", "9422"))

#: 전용 프로필 — **저장소 밖**에 둔다.
#:   `reports/` 에 두면 복구용 보관 회차가 수백 MB 를 매일 담는다([406]).
PROFILE = os.environ.get(
    "CSOS_BRIDGE_PROFILE",
    os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                 "CSOS", "chrome_bridge"))

#: 자국 — 이 다리가 무엇을 했나.  **죽어도 남는다**([180]).
STATE = os.path.join(ROOT, "reports", "브라우저다리_상태.json")

#: 한 밴드를 기다리는 최대 시간.  수집기의 `MAX_WAIT_MS`(30분)보다 넉넉해야
#: 도중에 끊고 '실패'라 적지 않는다([169]).
COLLECT_WAIT_S = int(os.environ.get("CSOS_BRIDGE_WAIT_S", "2400"))

CHROME_CANDIDATES = (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                 r"Google\Chrome\Application\chrome.exe"),
)


def chrome_exe():
    for p in CHROME_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return ""


def _get(path, timeout=5):
    return json.loads(urllib.request.urlopen(
        "http://127.0.0.1:%d%s" % (PORT, path), timeout=timeout).read().decode("utf-8"))


def alive():
    """CDP 가 살아 있나 → 브라우저 이름 또는 빈 문자열."""
    try:
        return _get("/json/version", timeout=2).get("Browser") or ""
    except Exception:
        return ""


def tabs():
    try:
        return [t for t in _get("/json") if t.get("type") == "page"]
    except Exception:
        return []


def up(wait_s=30):
    """전용 크롬을 띄운다.  이미 살아 있으면 **그대로 쓴다**(두 번 안 띄운다)."""
    b = alive()
    if b:
        return {"이미": True, "브라우저": b}
    exe = chrome_exe()
    if not exe:
        return {"오류": "크롬을 못 찾았다", "찾은자리": list(CHROME_CANDIDATES)}
    os.makedirs(PROFILE, exist_ok=True)
    args = [exe,
            "--remote-debugging-port=%d" % PORT,
            "--user-data-dir=" + PROFILE,
            "--no-first-run", "--no-default-browser-check",
            # 작게 화면 구석에 — 형님 화면을 덜 가린다.  가려져도 focusEmulation
            # 덕에 수집은 돈다(위 실측).
            "--window-size=560,420", "--window-position=1340,620",
            "about:blank"]
    # ★ 콘솔 창을 안 띄운다([272]).  크롬은 GUI 앱이라 창 자체는 필요하다.
    #   ⚠ 깃발을 `**kw` 로 넘기면 감사기가 **못 읽는다** — 그러면 멀쩡히 깃발을 단
    #     코드가 '무엇을 띄우는지 모름'으로 남아 관문이 막힌다(2026-08-26 실측:
    #     `t272` 가 여기서 죽었다). 감사기는 `creationflags=<이름>` 을 따라가므로
    #     **이름 그대로** 넘긴다 — 동작은 한 톨도 안 바뀐다.
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    subprocess.Popen(args, creationflags=flags)
    t0 = time.time()
    while time.time() - t0 < wait_s:
        time.sleep(1)
        b = alive()
        if b:
            return {"띄움": True, "브라우저": b, "초": round(time.time() - t0, 1)}
    return {"오류": "크롬이 %d초 안에 디버깅 포트를 안 열었다" % wait_s}


# ── CDP ───────────────────────────────────────────────────────────────────────

def _conn(tab, timeout=30):
    from websockets.sync.client import connect
    return connect(tab["webSocketDebuggerUrl"], max_size=None, open_timeout=timeout)


def _call(ws, method, params=None, _id=[0], timeout=60):
    _id[0] += 1
    mine = _id[0]
    ws.send(json.dumps({"id": mine, "method": method, "params": params or {}}))
    while True:
        m = json.loads(ws.recv(timeout=timeout))
        if m.get("id") == mine:
            if "error" in m:
                raise RuntimeError("%s: %s" % (method, str(m["error"])[:200]))
            return m.get("result") or {}


def _eval(ws, expr, timeout=60, await_promise=True):
    r = _call(ws, "Runtime.evaluate",
              {"expression": expr, "awaitPromise": await_promise,
               "returnByValue": True}, timeout=timeout)
    if "exceptionDetails" in r:
        raise RuntimeError("JS 예외: " + str(r["exceptionDetails"])[:300])
    return (r.get("result") or {}).get("value")


def activate(ws):
    """그 탭을 **크롬이 실제로 활성으로 다루게** 한다.

    ★ 속이는 것이 아니다 — 이 파일 머리의 실측 표를 볼 것.  `document.hidden` 을
      `defineProperty` 로 덮는 짓은 **하지 않는다**(2026-08-19 가짜 묘비).
    """
    _call(ws, "Emulation.setFocusEmulationEnabled", {"enabled": True})
    try:
        _call(ws, "Page.bringToFront")
    except Exception:
        pass          # 창이 최소화돼 있으면 실패할 수 있다 — focusEmulation 이 본체다


def open_tab(url):
    urllib.request.urlopen(
        "http://127.0.0.1:%d/json/new?%s" % (PORT, urllib.parse.quote(url, safe=":/?=&")),
        timeout=15).read()
    time.sleep(3)


def find_tab(prefix):
    for t in tabs():
        if (t.get("url") or "").startswith(prefix):
            return t
    return None


def ensure_tab(url, prefix=None, wait_s=25):
    """그 주소 탭이 있으면 쓰고, 없으면 연다."""
    t = find_tab(prefix or url)
    if t:
        return t
    open_tab(url)
    t0 = time.time()
    while time.time() - t0 < wait_s:
        t = find_tab(prefix or url)
        if t:
            return t
        time.sleep(1)
    return None


# ── 밴드 ──────────────────────────────────────────────────────────────────────

BAND_URL = "https://www.band.us/band/%s"


def band_login_state(ws):
    """로그인이 됐나 → ('ok'|'need-login'|'unknown', 설명).

    ★ **비밀번호를 대신 치지 않는다.**  안 됐으면 그 사실만 말하고 멈춘다.
    """
    try:
        v = _eval(ws, """(() => {
          const href = location.href;
          const feed = document.querySelectorAll('.cardMain, .postWrap, [data-viewname*="Post"]').length;
          const login = /\\/(login|signin)/i.test(href) ||
                        !!document.querySelector('a[href*="auth.band.us"], .loginArea');
          return { href: href.slice(0, 120), feed: feed, login: login };
        })()""", timeout=30)
    except Exception as e:
        return "unknown", "확인 못 함: %s" % str(e)[:120]
    if not isinstance(v, dict):
        return "unknown", "확인 못 함(응답 모양)"
    if v.get("login"):
        return "need-login", "로그인 화면이다 — 그 창에서 한 번 로그인하십시오"
    if (v.get("feed") or 0) > 0:
        return "ok", "글이 %d장 보인다" % v["feed"]
    return "unknown", "로그인 화면은 아닌데 글도 안 보인다 (%s)" % v.get("href", "")


def collect_band(band, wait_s=None, app_base="http://127.0.0.1:8899"):
    """한 밴드를 긁는다 → 결과 dict.

    수집기·계획은 **앱이 내려 주는 것을 그대로** 쓴다([162]) — 여기서 번호를
    새로 고르지 않는다.  붙여넣기가 하던 그 두 줄을 대신할 뿐이다.
    """
    wait_s = wait_s or COLLECT_WAIT_S
    out = {"밴드": str(band), "시작": time.strftime("%Y-%m-%dT%H:%M:%S")}
    t = ensure_tab(BAND_URL % band, prefix="https://www.band.us/band/%s" % band)
    if not t:
        out["결과"] = "실패"
        out["왜"] = "밴드 탭을 못 열었다"
        return out
    with _conn(t) as ws:
        activate(ws)
        state, why = band_login_state(ws)
        out["로그인"] = state
        if state != "ok":
            out["결과"] = "사람대기" if state == "need-login" else "모름"
            out["왜"] = why
            return out
        # 유저스크립트를 그대로 실은다 — 그 안에서 계획을 받고 수집기를 붙인다.
        src = _eval(ws, """(async () => {
          const r = await fetch('%s/band_auto_collect.user.js', { cache: 'no-store' });
          if (!r.ok) return { ok: false, why: 'HTTP ' + r.status };
          const s = await r.text();
          (0, eval)(s);
          return { ok: true, len: s.length };
        })()""" % app_base, timeout=90)
        if not isinstance(src, dict) or not src.get("ok"):
            out["결과"] = "실패"
            out["왜"] = "수집기를 못 실었다: %s" % (src or {}).get("why", "?")
            return out
        out["실은크기"] = src.get("len")
        # 유저스크립트는 스스로 간격을 본다([454]).  간격 전이면 안 긁으므로
        # **그 사실을 그대로 적는다** — '했다'고 꾸미지 않는다([169]).
        t0 = time.time()
        last = None
        while time.time() - t0 < wait_s:
            time.sleep(10)
            try:
                st = _eval(ws, "(window.__grabStatus ? window.__grabStatus() : null)",
                           timeout=30, await_promise=False)
            except Exception as e:
                out["결과"] = "실패"
                out["왜"] = "상태를 못 읽었다: %s" % str(e)[:120]
                return out
            if isinstance(st, dict):
                last = st
                if not st.get("running") and (st.get("ok") or st.get("failed")):
                    break
            elif last is None and time.time() - t0 > 120:
                out["결과"] = "안돎"
                out["왜"] = ("2분이 지나도 수집기가 시작을 안 했다 — 간격 전이거나"
                            " 계획이 비었을 수 있다(대기열을 본다)")
                return out
        out["걸린초"] = round(time.time() - t0, 1)
        out["상태"] = last
        if last and not last.get("running"):
            try:
                _eval(ws, "(window.__grabSave ? (window.__grabSave(), true) : false)",
                      timeout=60, await_promise=False)
                out["저장"] = True
            except Exception as e:
                out["저장"] = False
                out["왜"] = "저장 실패: %s" % str(e)[:120]
            out["결과"] = "완주" if out.get("저장") else "저장실패"
        else:
            out["결과"] = "시간초과"
            out["왜"] = "%d초 안에 안 끝났다 — 도중까지는 저장돼 있다" % wait_s
    return out


# ── 자국 ──────────────────────────────────────────────────────────────────────

def note(rec):
    """무엇을 했나 남긴다.  **못 적어도 일은 이미 됐다** — 조용히 넘어간다."""
    try:
        os.makedirs(os.path.dirname(STATE), exist_ok=True)
        old = []
        if os.path.exists(STATE):
            with open(STATE, encoding="utf-8") as f:
                old = json.load(f) or []
        if not isinstance(old, list):
            old = []
        old.insert(0, rec)
        tmp = STATE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(old[:60], f, ensure_ascii=False, indent=1)
        os.replace(tmp, STATE)
    except Exception:
        pass


def status():
    b = alive()
    out = {"포트": PORT, "프로필": PROFILE, "크롬": chrome_exe() or "(못 찾음)",
           "살아있음": bool(b), "브라우저": b}
    if b:
        out["탭"] = [{"url": (t.get("url") or "")[:80]} for t in tabs()]
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--up", action="store_true")
    ap.add_argument("--collect", metavar="밴드번호")
    ap.add_argument("--collect-all", action="store_true")
    ap.add_argument("--wait", type=int, default=COLLECT_WAIT_S)
    a = ap.parse_args()

    if a.status:
        print(json.dumps(status(), ensure_ascii=False, indent=1))
        return 0

    if a.up or a.collect or a.collect_all:
        r = up()
        print("크롬:", json.dumps(r, ensure_ascii=False))
        if r.get("오류"):
            return 1

    if a.up and not (a.collect or a.collect_all):
        t = ensure_tab("https://www.band.us/", prefix="https://www.band.us/")
        print("★ 그 창에서 밴드에 **한 번** 로그인하십시오 — 다음부터는 안 물어봅니다.")
        print("   (비밀번호는 제가 대신 치지 않습니다.)")
        if t:
            with _conn(t) as ws:
                activate(ws)
                st, why = band_login_state(ws)
                print("   지금:", st, "·", why)
        return 0

    bands = []
    if a.collect:
        bands = [a.collect]
    elif a.collect_all:
        try:
            import collect_queue as CQ
            with open(CQ.QUEUE_PATH, encoding="utf-8") as f:
                q = json.load(f)
            bands = sorted((q.get("bands") or {}).keys())
        except Exception as e:
            print("대기열을 못 읽었다:", type(e).__name__, str(e)[:120])
            return 1

    rc = 0
    for b in bands:
        r = collect_band(b, wait_s=a.wait)
        note(r)
        print(json.dumps(r, ensure_ascii=False))
        if r.get("결과") not in ("완주", "안돎"):
            rc = 1
    return rc


if __name__ == "__main__":
    # ★ 수집 문을 거친다([387]) — 두 창이 같은 밴드를 동시에 긁으면 캐시가
    #   오염된다(사고 #27).  무인 회차는 그대로 통과한다.
    try:
        # ★ 조회()는 문을 안 거친다([172]) — 상태를 보는 것까지 막으면
        #   무엇이 막혔는지도 못 본다.
        if not ({"--status"} & set(sys.argv[1:])):
            import collect_gate
            collect_gate.guard("브라우저 다리로 밴드 수집")
    except SystemExit:
        raise
    except Exception:
        pass          # 문을 못 읽는 것으로 수집을 통째로 막지는 않는다
    sys.exit(main())
