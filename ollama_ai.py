# -*- coding: utf-8 -*-
"""이 PC 안에서 도는 동네 AI(Ollama) — 앱이 규칙으로 못 답할 때만 부른다.

2026-09-25 형님 지시: **"이 pc에 설치된 올라마 ai 무료로 활용 가능하게 연동할 수
있어? 클라우드 로그인은 안할거야"**

★ 무엇을 채우는가 — `local_ai.ask()` 는 규칙에 없는 질문을 만나면 *"이 모양의
  질문은 아직 규칙이 없습니다"* 로 끝났다(그다음은 사람이 클로드에게 붙여넣는
  것뿐이다). 그 빈자리에 **공짜이고 밖으로 한 글자도 안 나가는** 답을 하나 얹는다.

★ 밖으로 나가지 않는다 — 127.0.0.1:11434 한 곳만 부른다. 로그인도 열쇠도 없다.
  그래서 이 길은 **크레딧을 한 톨도 안 쓴다**(2026-09-03 지시와 같은 방향이다).

★ 지어낸 답을 사실로 적지 않는다([169]). 4B 짜리 모델은 우리 업무를 모른다 —
  그래서 ① 앱이 **이미 확인한 사실만** 재료로 주고 ② 답에 `동네AI` 라고 못박고
  ③ 클로드에게 넘기는 문구를 **그대로 같이** 돌려준다. 이 답은 **참고**지 근거가
  아니다.

★ 없으면 조용히 빠진다 — Ollama 가 안 떠 있거나 느리면 `ready()` 가 False 이고
  `ask()` 는 `None` 이다. 규칙 답변기는 예전과 한 글자도 다르지 않게 돈다([172]).

되돌리기 한 줄: `COUPANG_OLLAMA=0` (끄기) · 다시 켜기는 그 값을 지우면 된다.
모델 고르기: `COUPANG_OLLAMA_MODEL=gemma3:4b` 처럼 환경변수로 준다.

사람이 보는 자리:
  python ollama_ai.py                 # 떠 있나 · 어떤 모델이 있나
  python ollama_ai.py "질문"          # 한 번 물어본다
"""
from __future__ import annotations

import io
import json
import os
import sys
import urllib.error
import urllib.request

HOST = os.environ.get("COUPANG_OLLAMA_HOST") or "http://127.0.0.1:11434"

# 이 PC 에 있는 것 중 한국어를 가장 덜 흔드는 순서. 없으면 그다음으로 내려간다.
# ★ 목록을 손으로 고르는 이유: 4B 급은 모델마다 한국어 품질 차이가 크다.
#   여기 없는 모델만 깔려 있으면 **그것을 쓴다**(0건으로 끝내지 않는다).
PREFER = ("exaone3.5:2.4b", "qwen3:4b", "gemma3:4b")

TIMEOUT_TAGS = 2.0      # 떠 있나만 보는 값 — 길면 앱 화면이 그만큼 멈춘다
TIMEOUT_ASK = 60.0      # 4B 모델이 CPU 에서 답하는 데 실제로 걸리는 시간
ANSWER_MAX = 1200       # 화면에 싣는 길이 — 길면 아무도 안 읽는다([170])

MARK = "동네AI"         # 화면·기록이 이 답을 가려 볼 이름표


def enabled():
    """사람이 껐으면 False. 기본은 켬(깔려 있으면 쓴다)."""
    return (os.environ.get("COUPANG_OLLAMA") or "1").strip() not in ("0", "false", "off")


def _get(path, timeout):
    req = urllib.request.Request(HOST + path, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def models():
    """깔려 있는 모델 이름들. 못 물어보면 빈 목록이다(모름을 '없음'이라 적지 않는다)."""
    try:
        d = _get("/api/tags", TIMEOUT_TAGS)
    except Exception:
        return []
    return [str(m.get("name") or "") for m in (d.get("models") or []) if m.get("name")]


def pick_model():
    """쓸 모델 하나. 사람이 정한 것이 먼저고, 그다음이 PREFER 순서다."""
    want = (os.environ.get("COUPANG_OLLAMA_MODEL") or "").strip()
    have = models()
    if want:
        # 사람이 적은 이름이 실제로 있는지 본다 — 없는 이름을 조용히 갈아치우지 않는다
        if want in have:
            return want
        for m in have:
            if m.split(":")[0] == want.split(":")[0]:
                return m
        return None
    for m in PREFER:
        if m in have:
            return m
    return have[0] if have else None


def ready():
    """지금 물어봐도 되나 — 켜져 있고, 떠 있고, 쓸 모델이 있다."""
    if not enabled():
        return False
    return bool(pick_model())


def status():
    """사람과 화면이 읽을 한 줄짜리 상태."""
    if not enabled():
        return {"켬": False, "왜": "COUPANG_OLLAMA=0 으로 꺼 두셨습니다", "모델": []}
    have = models()
    if not have:
        return {"켬": True, "쓸수있나": False,
                "왜": "Ollama 가 %s 에서 응답하지 않습니다 (앱을 켜 두셨는지 보십시오)" % HOST,
                "모델": []}
    return {"켬": True, "쓸수있나": True, "모델": have, "고른모델": pick_model(),
            "왜": "이 PC 안에서만 돕니다 — 로그인·요금 없음"}


SYSTEM = (
    "너는 한국 승강기·물류리프트 회사의 업무 비서다. 아래 '앱이 확인한 사실'만 근거로 "
    "한국어 존댓말로 짧게 답한다. 사실에 없는 숫자·날짜·이름은 절대 지어내지 말고, "
    "모르면 '앱 자료로는 알 수 없습니다'라고 답한다. 답은 다섯 줄 안으로 쓴다."
)


def ask(question, facts="", model=None, timeout=TIMEOUT_ASK):
    """한 번 물어본다. 못 쓰면 **None** 을 돌려준다(예외를 밖으로 내보내지 않는다).

    facts 는 앱이 이미 확인한 사실만 넣는다 — 비밀설정·열쇠는 넣지 않는다.
    """
    q = (question or "").strip()
    if not q or not enabled():
        return None
    m = model or pick_model()
    if not m:
        return None
    prompt = SYSTEM + "\n\n[앱이 확인한 사실]\n" + (facts.strip() or "(없음)") + "\n\n[질문]\n" + q + "\n\n[답]\n"
    body = json.dumps({
        "model": m,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": 0.2, "num_predict": 400},
    }).encode("utf-8")
    req = urllib.request.Request(HOST + "/api/generate", data=body,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            d = json.loads(r.read().decode("utf-8", "replace"))
    except Exception as e:
        return {"모델": m, "답": "", "못함": "%s: %s" % (type(e).__name__, e)}
    txt = (d.get("response") or "").strip()
    # 생각 과정을 내보내는 모델(qwen3 등)은 그 부분을 뗀다 — 화면에 실을 글이 아니다
    if "</think>" in txt:
        txt = txt.split("</think>", 1)[1].strip()
    if not txt:
        return {"모델": m, "답": "", "못함": "빈 답"}
    if len(txt) > ANSWER_MAX:
        txt = txt[:ANSWER_MAX] + " …(줄임)"
    return {"모델": m, "답": txt, "못함": ""}


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    args = [a for a in sys.argv[1:] if a.strip()]
    st = status()
    if not args:
        print("Ollama 상태 — " + json.dumps(st, ensure_ascii=False))
        print("  주소: %s" % HOST)
        print("  끄기: COUPANG_OLLAMA=0 · 모델 고르기: COUPANG_OLLAMA_MODEL=<이름>")
        return 0
    if not st.get("쓸수있나"):
        print("못 물어봅니다 — " + str(st.get("왜")))
        return 3
    r = ask(" ".join(args))
    if not r or r.get("못함"):
        print("못 물어봅니다 — " + str((r or {}).get("못함") or "알 수 없음"))
        return 3
    print("[%s · %s] %s" % (MARK, r["모델"], r["답"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
