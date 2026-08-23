# -*- coding: utf-8 -*-
"""크레딧 5시간 창 — **소진과 충전을 기록에서 읽는다**(2026-08-22 형님 지시).

형님 지시: "5시간 크래딧 소진 시 멈췄다가 크래딧 충전되면 자동으로 시작하는
알고리즘 앱에 박아놔 · 다른 계정이나 다른 세션에서도 인식하고 자동으로 반영되게"

★ **지어낼 것이 없다 — 대화기록이 그대로 적어 준다.** 실측 2026-08-22:
  quotaLimits = {"status":"rejected", "rateLimitType":"five_hour",
  "resetsAt":1787326800, ...} 이 기록 파일에 **492회** 찍혀 있었고, 그 값이
  2026-08-21 21:40 소진 · 00:40 충전으로 형님이 겪으신 멈춤과 정확히 맞았다.
  그러므로 '몇 시간째 조용하다'로 **추정하지 않는다** — 그것은 크레딧 소진과
  창을 닫은 것과 자리를 비운 것을 구별하지 못한다([169]).

★ **갈래는 셋이고 '모름'을 '충전됨'으로 치지 않는다**([169]).
  · 소진   — 거절 기록이 있고 그 resetsAt 이 **아직 안 왔다**
  · 충전됨 — 읽었는데 미래 거절이 없다(한 번도 안 걸렸거나 이미 지났다)
  · 모름   — 기록 폴더를 못 찾았거나 읽기가 실패했다
  못 읽은 것을 '충전됨'이라 하면 **소진 중에 AI 를 계속 불러** 실패만 쌓인다.
  반대로 '소진'이라 하면 **멀쩡한데 일을 멈춘다** — 그래서 둘 다 아닌 이름을 준다.

★ **읽기 전용이다.** 아무것도 안 고치고 큐에도 안 넣는다. 이 파일이 하는 일은
  **사실을 말하는 것**뿐이고, 그 사실을 어떻게 쓸지는 부르는 쪽이 정한다([162]).

⚠ **어느 계정 것인지는 기록이 말해 주지 않는다.** 파일 어디에도 계정 표시가 없다.
  그래서 판정은 "이 기계의 Claude 기록에서 본 **가장 최근** five_hour 거절"이다 —
  계정을 바꾸면 새 창이 새 기록을 쓰므로 대개 그것이 지금 계정 것이다. 이 한계를
  숨기지 않는다: 근거 파일 이름과 그 기록의 시각을 늘 같이 돌려준다.

⚠ **비싼 읽기를 안 한다**([168]). 기록 파일은 수십 MB 가 되므로 **뒤 꼬리만**
  읽고(TAIL_BYTES), 최근에 자란 파일만 본다(FRESH_DAYS). 실측 **0.2초**.
  ★ 꼬리만 읽어도 되는 근거는 **소진되면 그 뒤로 기록이 거의 안 자란다**는 것이다 —
    창이 막히면 그 창은 멈추므로 거절 줄이 꼬리에 남는다. 충전 뒤 이어 가면 그 줄이
    앞으로 밀려나 안 보이게 되는데, **그때는 이미 지난 일이라 답이 같다**(충전됨).
    즉 이 전략은 **막고 있는 창을 놓치지 않는다** — 그것이 우리가 알아야 할 전부다.
  ⚠ 그러므로 이 파일로 **지난 소진의 역사를 세면 안 된다.** 그것은 못 보는 값이다.
"""
import glob
import io
import json
import os
import re
import sys
import time
from datetime import datetime

# ★ 무인 회차는 `pythonw` 로 돌아 `sys.stdout` 이 **None** 이고, 콘솔은 cp949 라
#   `—` 한 글자에 죽는다([235] · `userscript_watch` 가 같은 자리에서 당했다).
#   `hasattr(None, ...)` 는 False 이므로 이 한 줄이 둘 다 막는다.
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORTS = os.path.join(ROOT, "reports")
STATE = os.path.join(REPORTS, "크레딧_창.json")

#: 기록 파일 꼬리에서 이만큼만 읽는다 — 거절은 창이 막힌 순간에 찍혀 끝쪽에 있다.
TAIL_BYTES = 512 * 1024
#: 이보다 오래 안 자란 기록은 안 본다(지난 계정·지난 주 창).
FRESH_DAYS = 3
#: 이 낱말이 있는 줄만 파싱한다 — 수십만 줄을 통째로 json 으로 읽지 않는다.
MARK = "quotaLimits"
#: 우리가 아는 창 종류. 모르는 종류는 **소진이라 우기지 않는다**([169]).
KNOWN_TYPES = ("five_hour",)

#: Codex CLI 는 Claude 기록과 다른 곳에 한도 시각을 남긴다. 실제 실패 문구가
#: ``try again at Aug 28th, 2026 2:52 PM`` 모양이므로 그 **명시된 시각만** 읽는다.
#: 조용했던 시간으로 한도를 추정하지 않는다([169]).
CODEX_FRESH_DAYS = 7
CODEX_TAIL_BYTES = 64 * 1024
_CODEX_RESET_RE = re.compile(
    r"try again at\s+([A-Z][a-z]{2}\s+\d{1,2}(?:st|nd|rd|th),\s+"
    r"\d{4}\s+\d{1,2}:\d{2}\s+(?:AM|PM))",
    re.I,
)


def _dirs():
    """기록 폴더 후보. 이 프로젝트 것이 먼저이고, 그다음이 이 기계의 나머지다.

    ★ 다른 프로젝트 폴더까지 보는 이유는 **크레딧이 계정 단위**이기 때문이다 —
      형님이 옆 프로젝트 창에서 소진하셨어도 이 창은 같이 멈춘다. 한 폴더만 보면
      그 사실을 못 본다.
    """
    home = os.path.join(os.path.expanduser("~"), ".claude", "projects")
    if not os.path.isdir(home):
        return []
    mine = os.path.basename(ROOT)
    out = []
    try:
        for d in os.scandir(home):
            if d.is_dir():
                out.append(d.path)
    except OSError:
        return []
    out.sort(key=lambda p: (mine not in os.path.basename(p), os.path.basename(p)))
    return out


def _quota_of(line_text):
    """한 줄에서 quotaLimits 를 꺼낸다. 못 읽으면 None — 조용히 넘긴다."""
    try:
        d = json.loads(line_text)
    except Exception:
        return None, ""
    found = [None]

    def walk(o):
        if isinstance(o, dict):
            q = o.get(MARK)
            if isinstance(q, dict):
                found[0] = q
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for x in o:
                walk(x)

    walk(d)
    return found[0], str(d.get("timestamp") or "")


def scan(now=None, dirs=None):
    """기록을 훑어 **가장 최근 거절**을 찾는다. (기록, 왜못함).

    기록은 {"resetsAt", "종류", "언제", "파일"} 이고, 없으면 None 이다.
    """
    now = now or time.time()
    dirs = _dirs() if dirs is None else list(dirs)
    if not dirs:
        return None, "대화기록 폴더를 못 찾았습니다(~/.claude/projects)"
    best = None
    본파일 = 0
    실패 = 0
    cutoff = now - FRESH_DAYS * 86400
    for d in dirs:
        for f in glob.glob(os.path.join(d, "*.jsonl")):
            try:
                if os.path.getmtime(f) < cutoff:
                    continue
                size = os.path.getsize(f)
            except OSError:
                실패 += 1
                continue
            try:
                with io.open(f, "rb") as fh:
                    if size > TAIL_BYTES:
                        fh.seek(size - TAIL_BYTES)
                        fh.readline()          # 잘린 첫 줄은 버린다
                    raw = fh.read()
            except OSError:
                실패 += 1
                continue
            본파일 += 1
            for ln in raw.decode("utf-8", "replace").splitlines():
                if MARK not in ln:
                    continue
                q, ts = _quota_of(ln)
                if not isinstance(q, dict):
                    continue
                if str(q.get("status") or "") != "rejected":
                    continue
                종류 = str(q.get("rateLimitType") or "")
                try:
                    r = int(q.get("resetsAt") or 0)
                except (TypeError, ValueError):
                    continue
                if r <= 0:
                    continue
                rec = {"resetsAt": r, "종류": 종류, "언제": ts,
                       "파일": os.path.basename(f)}
                if best is None or r > best["resetsAt"]:
                    best = rec
    if 본파일 == 0:
        return None, "최근 %d일 안에 자란 기록이 없습니다(읽기 실패 %d)" % (FRESH_DAYS, 실패)
    return best, ""


def _codex_reset_at(text):
    """Codex 한도 문구에서 로컬 시각을 초로 바꾼다. 없거나 못 읽으면 0."""
    found = list(_CODEX_RESET_RE.finditer(text or ""))
    if not found:
        return 0
    raw = found[-1].group(1)
    raw = re.sub(r"(\d{1,2})(?:st|nd|rd|th)", r"\1", raw, flags=re.I)
    try:
        # CLI 문구에 시간대가 없으므로 그 문구를 만든 이 PC 의 로컬 시각으로 읽는다.
        return int(datetime.strptime(raw, "%b %d, %Y %I:%M %p").timestamp())
    except ValueError:
        return 0


def scan_codex(now=None, report_dir=None):
    """최근 Codex 실행표에서 **명시된 사용 재개 시각**을 찾는다.

    오래된 실패를 현재 고장으로 되살리지 않기 위해 최근 파일만 보고, JSON 의 잘린
    오류 대신 원문이 남는 ``.log`` 꼬리도 함께 본다. 결과는 가장 최근에 생긴
    한도 실패 하나다. 읽을 파일이 없으면 '충전됨'이 아니라 '모름'이다([169]).
    """
    now = now or time.time()
    report_dir = report_dir or os.path.join(REPORTS, "agent_dispatch")
    if not os.path.isdir(report_dir):
        return None, "Codex 실행표 폴더를 못 찾았습니다"
    cutoff = now - CODEX_FRESH_DAYS * 86400
    best = None
    본파일 = 0
    실패 = 0
    files = glob.glob(os.path.join(report_dir, "*.log"))
    files += glob.glob(os.path.join(report_dir, "*.json"))
    for path in files:
        try:
            mtime = os.path.getmtime(path)
            if mtime < cutoff:
                continue
            size = os.path.getsize(path)
            with io.open(path, "rb") as fh:
                if size > CODEX_TAIL_BYTES:
                    fh.seek(size - CODEX_TAIL_BYTES)
                    fh.readline()
                text = fh.read().decode("utf-8", "replace")
        except OSError:
            실패 += 1
            continue
        본파일 += 1
        reset = _codex_reset_at(text)
        if not reset:
            continue
        rec = {"resetsAt": reset, "종류": "codex_usage_limit",
               "언제": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime)),
               "파일": os.path.basename(path), "근거시각": mtime}
        if best is None or mtime > best["근거시각"]:
            best = rec
    if 본파일 == 0:
        return None, "최근 %d일 Codex 실행표가 없습니다(읽기 실패 %d)" % (
            CODEX_FRESH_DAYS, 실패)
    return best, ""


def _state_from_scan(rec, why, now):
    """두 수집 갈래가 함께 쓰는 세 상태 판정."""
    if why:
        return {"갈래": "모름", "resetsAt": None, "남은분": 0, "종류": "",
                "언제": "", "파일": "", "왜": why}
    if rec is None:
        return {"갈래": "충전됨", "resetsAt": None, "남은분": 0, "종류": "",
                "언제": "", "파일": "", "왜": "최근 기록에 한도 거절이 없습니다"}
    남은 = rec["resetsAt"] - now
    known = rec["종류"] in KNOWN_TYPES or rec["종류"] == "codex_usage_limit"
    갈래 = "소진" if 남은 > 0 and known else "모름" if 남은 > 0 else "충전됨"
    return {"갈래": 갈래, "resetsAt": rec["resetsAt"],
            "남은분": int(max(0, 남은) // 60), "종류": rec["종류"],
            "언제": rec.get("언제") or "", "파일": rec.get("파일") or "",
            "왜": "" if 갈래 != "모름" else
                  "모르는 창 종류입니다: " + (rec["종류"] or "(빈값)")}


def state(now=None, dirs=None):
    """지금 크레딧 창이 어떤가 — **판정은 여기 한 곳**이다([162]).

    돌려주는 것: {갈래, resetsAt, 남은분, 종류, 언제, 파일, 왜}.
    """
    now = now or time.time()
    rec, why = scan(now, dirs)
    return _state_from_scan(rec, why, now)


def codex_state(now=None, report_dir=None):
    """Codex 크래딧 상태. Claude 5시간 창과 절대 합쳐 세지 않는다."""
    now = now or time.time()
    rec, why = scan_codex(now, report_dir)
    return _state_from_scan(rec, why, now)


def combined_state(now=None, dirs=None, report_dir=None):
    """화면·공유 파일용 상태 — Claude와 Codex를 **갈라서** 싣는다."""
    now = now or time.time()
    agents = {"claude": state(now, dirs), "codex": codex_state(now, report_dir)}
    exhausted = [name for name, st in agents.items() if st.get("갈래") == "소진"]
    unknown = [name for name, st in agents.items() if st.get("갈래") == "모름"]
    if len(exhausted) == len(agents):
        갈래 = "소진"
    elif exhausted:
        갈래 = "제한"
    elif unknown:
        갈래 = "모름"
    else:
        갈래 = "충전됨"
    waits = [agents[n].get("남은분") or 0 for n in exhausted]
    resets = [agents[n].get("resetsAt") for n in exhausted if agents[n].get("resetsAt")]
    return {"갈래": 갈래, "대상": exhausted, "모름대상": unknown,
            "resetsAt": max(resets) if resets else None,
            "남은분": max(waits) if waits else 0, "종류": "multi_agent",
            "언제": "", "파일": "", "왜": "", "agents": agents}


def blocked(now=None, agent=None):
    """지금 AI 를 부르면 안 되나. **모름은 막지 않는다** — 멀쩡한데 멈추지 않기 위해서다.

    ★ 방향을 이렇게 정한 이유: 소진 중에 부르면 **실패 한 번**이고 그것은 되돌릴 수
      있다(대기열에 남아 다음에 다시 간다). 반대로 멀쩡한데 막으면 **일이 안 된다**.
      그러므로 확실할 때만 막는다.
    """
    if agent == "codex":
        return codex_state(now)["갈래"] == "소진"
    # 인자를 안 준 옛 호출자는 Claude 5시간 창 판정을 그대로 쓴다. 에이전트별
    # 길 선택은 agent_dispatch 가 ``blocked(agent=...)`` 로 따로 묻는다.
    return state(now)["갈래"] == "소진"


def note(now=None):
    """지금 상태를 파일에 적는다 — **다른 계정·세션이 읽는 자리**가 여기다.

    형님 지시의 "다른 계정이나 다른 세션에서도 인식"이 이 한 줄이다. 대화기록은
    창마다 따로지만 이 파일은 **저장소가 공유**하므로, 무인 회차가 30분마다 적어 두면
    어느 창이 열리든 같은 사실을 본다.

    ★ **못 적어도 판정은 돌려준다** — 자국 하나 때문에 답을 죽이지 않는다.
    """
    st = combined_state(now)
    st["적은때"] = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        os.makedirs(REPORTS, exist_ok=True)
        tmp = STATE + ".tmp"
        with io.open(tmp, "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=2)
        os.replace(tmp, STATE)
    except OSError as exc:
        st["적기실패"] = str(exc)[:80]
    return st


def load():
    """회차가 적어 둔 것을 읽는다(비싼 훑기 없이). 없으면 {}."""
    try:
        with io.open(STATE, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def line(st=None):
    """사람이 읽는 한 줄. 갈래마다 **조치가 다르다**([289])."""
    st = st or combined_state()
    agents = st.get("agents") or {}
    if agents:
        parts = []
        for name, label in (("claude", "Claude"), ("codex", "Codex")):
            one = agents.get(name) or {}
            if one.get("갈래") == "소진":
                when = time.strftime("%m-%d %H:%M", time.localtime(one["resetsAt"]))
                parts.append("%s 소진(%s 충전 · %d분)" %
                             (label, when, one.get("남은분") or 0))
            elif one.get("갈래") == "모름":
                parts.append("%s 확인못함" % label)
            else:
                parts.append("%s 사용 가능" % label)
        return "크래딧 — " + " · ".join(parts)
    갈 = st.get("갈래")
    if 갈 == "소진":
        when = time.strftime("%H:%M", time.localtime(st["resetsAt"]))
        return ("크레딧 5시간 창이 찼습니다 — %s 에 충전됩니다(%d분 남음). "
                "무인 회차는 그대로 돕니다(파이썬이라 크레딧과 무관합니다)."
                % (when, st.get("남은분") or 0))
    if 갈 == "모름":
        return "크레딧 창을 **확인 못 했습니다** — %s" % (st.get("왜") or "")
    return "크레딧 창 여유 있음"


def main(argv=None):
    import argparse
    p = argparse.ArgumentParser(description="크레딧 5시간 창 판정(읽기 전용)")
    p.add_argument("--print", action="store_true", help="사람이 읽는 한 줄")
    p.add_argument("--write", action="store_true", help="reports/크레딧_창.json 갱신")
    a = p.parse_args(argv)
    # 사람 출력도 공용 파일·앱과 같은 두 에이전트 판정을 쓴다. Claude만 보면
    # Codex가 막힌 동안에도 "여유 있음"으로 잘못 안내한다([203]).
    st = note() if a.write else combined_state()
    print(line(st))
    print(json.dumps({k: v for k, v in st.items() if k != "왜"}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
