# -*- coding: utf-8 -*-
"""
cloud_publish.py — PC가 꺼져 있어도 폰·태블릿이 쓰는 사본을 고정 주소에 올린다
================================================================================
지금까지의 폰 접속은 전부 **사무실 PC가 살아 있어야** 했다(터널·로컬 서버). 이동이 잦아
PC를 못 켜 두면 폰에는 크롬 기본 오류만 떴다. 그래서 데이터를 고정 주소에 얹는다.

  PC 켜져 있을 때 : 고정 주소 → 실시간 앱(입력·대조·엑셀 반영까지 전부)
  PC 꺼져 있을 때 : 고정 주소 → **오프라인 앱**(조회·프로젝트코드 자동채움·입력 예약)
                    예약한 입력은 PC가 켜지는 순간 자동으로 원장에 반영된다.

고정 주소 저장소는 공개다. 그래서 데이터는 **잠가서** 올린다(csos_crypto).
폰에서 PIN을 넣으면 브라우저가 그 자리에서 푼다 — 서버도 계정도 필요 없다.

  python cloud_publish.py            # 사본 생성(로컬 확인)
  python cloud_publish.py --push     # 생성 + git commit·push (고정 주소에 실제 반영)
"""
import sys, os, json, zlib, subprocess
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "webapp"))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DOCS = os.path.join(ROOT, "docs")
OUT = os.path.join(DOCS, "data.enc")
WEBCFG = os.path.join(ROOT, "config", "webapp.json")


def pin():
    """PIN은 config에서만 온다 — 코드·엑셀·커밋 어디에도 적지 않는다."""
    try:
        return str(json.load(open(WEBCFG, encoding="utf-8"))["pin"])
    except Exception:
        sys.exit("config/webapp.json 에 PIN이 없습니다 — 앱을 한 번 실행하면 만들어집니다.")


def payload():
    """폰이 필요로 하는 것만 담는다. 인증키·엑셀 원본·밴드 원문은 담지 않는다."""
    import mobile_snapshot as M
    import project_resolve as P

    d = M.collect()                       # 확인필요·업무·정산·계산서 (앱과 같은 소스)
    ev = P.evidence()

    codes = {}
    for c in sorted(set(ev["band"]) | set(ev["ledger"]) | set(ev["book"])):
        r = P.resolve(c, ev)
        rec = {k: r.get(k) for k in ("camp", "kind", "cost", "tech", "date", "status",
                                     "sheet", "row", "state") if r.get(k)}
        if r.get("ids"):
            rec["ids"] = r["ids"]
        if r.get("unknown"):
            rec["miss"] = r["unknown"]
        if not r.get("ok"):
            rec["why"] = r.get("reason", "")
        codes[c] = rec

    d["codes"] = codes
    d["tail"] = ev["tail"]
    d["cap"] = ev["cap"]
    d["gen"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return d


def main():
    import csos_crypto as C
    assert C.self_test(), "암호 자체검증 실패 — 올리면 폰에서 못 연다"

    print("사본 만드는 중…")
    d = payload()
    raw = json.dumps(d, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    # LTE로 받는 파일이라 크기가 곧 대기시간이다. 잠그기 **전에** 줄인다
    # (암호문은 무작위에 가까워서 나중에 줄이면 하나도 안 줄어든다).
    # 폰에서는 브라우저 내장 DecompressionStream('deflate')이 그대로 푼다.
    packed = zlib.compress(raw, 9)
    sealed = C.seal(packed, pin())
    sealed["zip"] = "deflate"

    os.makedirs(DOCS, exist_ok=True)
    json.dump(sealed, open(OUT, "w", encoding="utf-8"), separators=(",", ":"))
    kb, rkb = os.path.getsize(OUT) / 1024, len(raw) / 1024

    print(f"  프로젝트코드 {len(d['codes'])} · 확인필요 {len(d['issues'])} · "
          f"AS {len(d['as'])} · 점검 {len(d['pm'])} · 정산 {len(d['settle'])} · 계산서 {len(d['erp'])}")
    print(f"  {rkb:.0f}KB → 줄여서 {len(packed)/1024:.0f}KB → 잠근 뒤 {kb:.0f}KB  ({OUT})")

    if "--push" not in sys.argv:
        print("\n로컬 생성만 했습니다 — 고정 주소 반영: python cloud_publish.py --push")
        return

    # 잠근 파일만 올린다. 원본(raw)은 디스크에도 남기지 않는다.
    for cmd in (["git", "add", "docs/data.enc", "docs/app.html", "docs/index.html",
                 "docs/sw.js", "docs/resolve_index.json"],
                ["git", "commit", "-q", "-m",
                 f"폰 사본 갱신 {d['gen']} (프로젝트코드 {len(d['codes'])}·확인필요 {len(d['issues'])})"],
                ["git", "push", "-q"]):
        r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        if r.returncode and "nothing to commit" not in (r.stdout or ""):
            print("  git:", (r.stderr or r.stdout or "").strip()[:200])
            if cmd[1] == "push":
                sys.exit(1)
    print("\n고정 주소에 반영했습니다 — PC를 꺼도 폰에서 열립니다.")
    print("  https://mulder4780.github.io/coupang-ecount-reconcile/")


if __name__ == "__main__":
    main()
