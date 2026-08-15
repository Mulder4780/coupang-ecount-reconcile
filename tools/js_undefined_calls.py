# -*- coding: utf-8 -*-
"""부르는데 **어디에도 만든 적 없는** 함수를 찾는다 (2026-08-15).

★ 왜 필요한가 — `fmtDateTime is not defined` 로 시스템 진단 카드가 통째로 안 그려지고
  있었다(실측 16건). 문법은 성하다(`node --check` 통과). **그 줄을 밟을 때만** 터지므로
  어느 검사에도 안 걸렸고, 터지면 그 함수만 실패하는 것이 아니라 **부르던 화면이 거기서
  통째로 멈춘다** — 카드가 비어 보일 뿐 오류창은 안 뜬다.

★ 잘못 지목하지 않는 것이 잡는 것보다 어렵다([172]). 그래서 문을 좁게 건다:
  · 파일 **어디에도** 선언이 없는 이름만 지목한다(함수·const·let·var·클래스·매개변수·
    객체 속성 전부 훑어 하나라도 있으면 안 건드린다).
  · 브라우저·표준 내장은 명단으로 뺀다.
  · 소문자로 시작하는 것만 본다(`Date(`·`Number(` 같은 생성자는 내장이다).
  후보가 애매하면 **안 부른다** — 없는 오타를 찾아 나서게 하는 것이 못 잡는 것보다 나쁘다.

쓰기:  python tools/js_undefined_calls.py          # 목록
      python tools/js_undefined_calls.py --count  # 건수만 (검증이 쓴다)
"""
from __future__ import annotations

import argparse
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGETS = ("webapp/index.html", "docs/app.html")

# 브라우저·표준이 주는 것. 여기 없는데 우리가 안 만들었으면 그것이 고장이다.
BUILTIN = {
    # 함수·전역
    "alert", "confirm", "prompt", "fetch", "setTimeout", "setInterval",
    "clearTimeout", "clearInterval", "requestAnimationFrame", "cancelAnimationFrame",
    "encodeURIComponent", "decodeURIComponent", "encodeURI", "decodeURI",
    "parseInt", "parseFloat", "isNaN", "isFinite", "eval", "structuredClone",
    "queueMicrotask", "btoa", "atob", "matchMedia", "getComputedStyle",
    "scrollTo", "scrollBy", "open", "close", "focus", "blur", "print",
    "addEventListener", "removeEventListener", "dispatchEvent",
    "postMessage", "reportError", "importScripts", "createImageBitmap",
    # 자주 쓰는 메서드 이름이 호출처럼 보이는 것들 (`.` 없이 잡히면 곤란하므로 함께 뺀다)
    "require", "define", "import",
}

DECL = [
    re.compile(r"\bfunction\s+([A-Za-z_$][\w$]*)"),
    re.compile(r"\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)"),
    re.compile(r"\bclass\s+([A-Za-z_$][\w$]*)"),
    # const {a, b} = ... / function f(a, b)  — 매개변수·구조분해까지 넉넉히 담는다
    re.compile(r"\b(?:const|let|var)\s*\{([^}]*)\}"),
    re.compile(r"\bfunction\s*[A-Za-z_$\w]*\s*\(([^)]*)\)"),
    re.compile(r"\(([^)]*)\)\s*=>"),
    re.compile(r"\b([A-Za-z_$][\w$]*)\s*:\s*(?:function|\()"),   # 객체 속성 형태
    re.compile(r"\b([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?(?:function|\()"),
]

CALL = re.compile(r"(?<![\w$.])([a-z][\w$]*)\s*\(")


def scripts(text: str):
    """inline <script> 본문만. 속성(onclick=…)은 따로 본다."""
    return [m.group(2) for m in
            re.finditer(r"<script([^>]*)>(.*?)</script>", text, re.S | re.I)
            if "src=" not in m.group(1)]


def handlers(text: str):
    """`onclick="foo(...)"` 같은 인라인 손잡이가 부르는 이름."""
    out = []
    for m in re.finditer(r"\bon[a-z]+\s*=\s*\"([^\"]*)\"", text, re.I):
        for c in CALL.finditer(m.group(1)):
            out.append(c.group(1))
    return out


def declared(text: str) -> set:
    names = set()
    for rx in DECL:
        for m in rx.finditer(text):
            for part in m.group(1).split(","):
                nm = part.strip().split("=")[0].strip().strip(".")
                nm = re.sub(r"^\.\.\.", "", nm)
                if re.fullmatch(r"[A-Za-z_$][\w$]*", nm or ""):
                    names.add(nm)
    return names


def scan():
    """(파일, 이름, 어디서 부르나) — 만든 적 없는 이름만."""
    bad = []
    for rel in TARGETS:
        p = os.path.join(ROOT, rel.replace("/", os.sep))
        if not os.path.isfile(p):
            continue
        text = open(p, encoding="utf-8", errors="replace").read()
        have = declared(text) | BUILTIN
        seen = {}
        for body in scripts(text):
            for m in CALL.finditer(body):
                seen.setdefault(m.group(1), "script")
        for nm in handlers(text):
            seen.setdefault(nm, "인라인 손잡이")
        for nm, where in sorted(seen.items()):
            if nm not in have:
                bad.append((rel, nm, where))
    return bad


def main(argv=None):
    ap = argparse.ArgumentParser(description="만든 적 없는데 부르는 함수")
    ap.add_argument("--count", action="store_true")
    a = ap.parse_args(argv)
    bad = scan()
    if a.count:
        print(len(bad))
        return 0
    if not bad:
        print("부르는데 만든 적 없는 함수: 0개")
        return 0
    print("부르는데 만든 적 없는 함수 %d개 — 그 줄을 밟으면 화면이 거기서 멈춘다" % len(bad))
    for rel, nm, where in bad:
        print("  %-22s %-28s (%s)" % (rel, nm, where))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
