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
import sys, os, re, json, zlib, subprocess
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

    # 기사별 방문 — '이 캠프 누가 갔었지'를 현장에서 바로 확인한다
    try:
        import tech_report as TR
        visits, pending, unknown, _m = TR.collect()
        by = TR.summary(visits)
        d["tech"] = [{"기사": t, "총": v["총"], "돌발AS": v["돌발AS"],
                      "정기점검": v["정기점검"], "캠프수": len(v["캠프"]), "최근": v["최근"]}
                     for t, v in sorted(by.items(), key=lambda x: -x[1]["총"])]
        # 전부 담으면 사본이 커진다 — 최근 것부터 필요한 만큼만
        d["visits"] = [{k: v[k] for k in ("기사", "방문일", "업무", "캠프명", "프로젝트NO")}
                       for v in sorted(visits, key=lambda x: x["방문일"], reverse=True)[:1200]]
        d["tech_gap"] = {"미방문": len(pending), "기사미기입": len(unknown)}
    except Exception as e:
        print("  ! 기사 정리 건너뜀:", e)

    d["codes"] = codes
    d["tail"] = ev["tail"]
    d["cap"] = ev["cap"]
    d["gen"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return d


def aging(path=None):
    """미청구 건의 **경과일만** 공개 파일로 낸다 → GitHub Actions가 매일 늙는 걸 본다.

    ★ 여기에는 금액·PO번호·캠프명을 절대 넣지 않는다. 저장소가 공개라서다.
      날짜만 있어도 '90일 넘은 게 3건' 은 계산되고, 그거면 알림으로 충분하다.
      상세는 잠긴 사본(data.enc)에 들어 있고 폰에서 PIN을 넣어야 보인다.

    이걸 PC가 아니라 Actions가 보는 이유: 출장으로 **PC를 며칠 못 켜도** 미청구는
    계속 늙기 때문이다. 경과일은 새 데이터 없이 날짜만으로 계산된다.
    """
    import csv as _csv, glob as _g
    path = path or os.path.join(DOCS, "aging.json")
    rep = sorted(_g.glob(os.path.join(ROOT, "reports", "PO대조_*.csv")))
    items = []
    if rep:
        with open(rep[-1], encoding="utf-8-sig") as f:
            for r in _csv.DictReader(f):
                if r.get("유형") != "A":            # A = 미청구(계산서 미발행)
                    continue
                d = (r.get("발행일") or "")[:10]
                if re.match(r"\d{4}-\d{2}-\d{2}", d):
                    items.append({"k": "po", "since": d})
    doc = {"generated": datetime.now().strftime("%Y-%m-%d"),
           "note": "경과일 계산용. 금액·번호·현장명은 담지 않는다(공개 저장소).",
           "warn_days": 90, "crit_days": 120, "items": sorted(items, key=lambda x: x["since"])}
    json.dump(doc, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return len(items)


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
    n_age = aging()          # 공개(날짜만) — Actions가 매일 경과일을 본다
    kb, rkb = os.path.getsize(OUT) / 1024, len(raw) / 1024

    print(f"  프로젝트코드 {len(d['codes'])} · 확인필요 {len(d['issues'])} · "
          f"AS {len(d['as'])} · 점검 {len(d['pm'])} · 정산 {len(d['settle'])} · 계산서 {len(d['erp'])}")
    print(f"  {rkb:.0f}KB → 줄여서 {len(packed)/1024:.0f}KB → 잠근 뒤 {kb:.0f}KB  ({OUT})")
    print(f"  경과일 공개파일 aging.json — 미청구 {n_age}건 (금액·번호는 담지 않음)")

    if "--push" not in sys.argv:
        print("\n로컬 생성만 했습니다 — 고정 주소 반영: python cloud_publish.py --push")
        return

    # 잠근 파일만 올린다. 원본(raw)은 디스크에도 남기지 않는다.
    for cmd in (["git", "add", "docs/data.enc", "docs/app.html", "docs/index.html",
                 "docs/sw.js", "docs/resolve_index.json", "docs/aging.json"],
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
