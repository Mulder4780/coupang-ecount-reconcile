# -*- coding: utf-8 -*-
"""
cloud_publish.py — PC가 꺼져 있어도 폰·태블릿이 쓰는 사본을 고정 주소에 올린다
================================================================================
지금까지의 폰 접속은 전부 **사무실 PC가 살아 있어야** 했다(터널·로컬 서버). 이동이 잦아
PC를 못 켜 두면 폰에는 크롬 기본 오류만 떴다. 그래서 데이터를 고정 주소에 얹는다.

  항상              : 고정 주소 → **PC 독립 앱**(조회·프로젝트코드 자동채움·입력 예약)
  PC 켜져 있을 때   : 독립 앱 안의 선택 버튼으로 실시간 앱 연결(입력·대조·엑셀 반영)
  PC 꺼져 있을 때   : 잠긴 최신 사본으로 계속 사용
                    예약은 PC가 켜진 뒤 같은 폰 앱을 열면 입력 큐로 전송된다.

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
QUEUECFG = os.path.join(ROOT, "config", "cloud_queue.local.json")
PUBLISH_FILES = (
    "docs/data.enc", "docs/app.html", "docs/index.html", "docs/manifest.json", "docs/sw.js"
)


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
    import app_server as A

    d = M.collect()                       # 확인필요·업무·정산·계산서 (앱과 같은 소스)
    ev = P.evidence()

    codes = {}
    for c in sorted(set(ev["band"]) | set(ev["ledger"]) | set(ev["book"])):
        if not re.fullmatch(r"UJ26\d{5}", str(c or ""), re.I):
            continue
        r = P.resolve(c, ev)
        if not A.app_project_result(c, r):
            continue
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
        visits = A.app_year_rows(visits, "visit")
        pending = A.app_year_rows(pending, "visit_pending")
        unknown = A.app_year_rows(unknown, "visit")
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

    # 대표 브리핑 — 폰에서 그대로 읽어 드릴 수 있게 담는다
    try:
        import daily_brief as DB
        _b = DB.brief(data=DB.load()[0])
        d["brief"] = {
            "기준일": _b["기준일"],
            "text": DB.text(_b),
            "돌발AS": _b["돌발AS"],
            "정기점검": _b["정기점검"],
            # 고정 주소의 PC 독립 앱에서도 숫자만 보이지 않고 원천 건을 열어야 한다.
            # 이 필드들은 data.enc 안에만 들어가며 공개 HTML에는 업무 내용이 남지 않는다.
            "당일처리목록": _b.get("당일처리목록", []),
            "완료일미기입목록": _b.get("완료일미기입목록", []),
            "내용미기입": _b.get("내용미기입", []),
        }
    except Exception as e:
        print("  ! 브리핑 건너뜀:", e)

    d["unbilled"] = unbilled()        # 미청구 — 앱이 열릴 때 경과일을 계산해 띄운다
    d["codes"] = codes
    d["tail"] = ev["tail"]
    d["cap"] = ev["cap"]
    d["app_year"] = "2026"
    # 휴대폰 입력 전용 토큰은 공개 소스가 아니라 PIN 암호화 사본 안에만 넣는다.
    # PC 작업자 토큰은 이 사본에 절대 포함하지 않는다.
    try:
        qcfg = json.load(open(QUEUECFG, encoding="utf-8"))
        if qcfg.get("api_base_url") and qcfg.get("enqueue_token"):
            d["cloud_queue"] = {
                "url": str(qcfg["api_base_url"]).rstrip("/"),
                "token": str(qcfg["enqueue_token"]),
            }
    except Exception:
        pass
    d["gen"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return d


def unbilled():
    """미청구 PO를 **잠긴 사본 안에** 담는다 — 금액·PO번호까지 그대로.

    ★ 예전에는 날짜만 담은 공개 파일(docs/aging.json)을 만들어 GitHub Actions가 매일 보고
      기준을 넘으면 메일을 보내게 했다. 사용자가 메일을 원하지 않아 그 경로는 걷어냈다.
      대신 여기 담아 두고 **앱을 열 때 화면 맨 위에 띄운다** — 경과일은 앱이 그 자리에서
      오늘 날짜로 계산하므로, 사본이 며칠 묵어도 '몇 일 지났는지'는 항상 정확하다.
    """
    import csv as _csv, glob as _g
    rep = sorted(_g.glob(os.path.join(ROOT, "reports", "PO대조_*.csv")))
    if not rep:
        return []
    out = []
    with open(rep[-1], encoding="utf-8-sig") as f:
        for r in _csv.DictReader(f):
            if r.get("유형") != "A":            # A = 미청구(계산서 미발행)
                continue
            d = (r.get("발행일") or "")[:10]
            if not re.match(r"2026-\d{2}-\d{2}", d):
                continue
            try:
                amt = int(float(r.get("쿠팡금액") or 0))
            except ValueError:
                amt = 0
            out.append({"PO": r.get("PO번호", ""), "발행일": d, "금액": amt,
                        "내용": (r.get("내용") or "")[:40],
                        "프로젝트NO": r.get("프로젝트NO", "")})
    return sorted(out, key=lambda x: x["발행일"])


def git_publish(message, runner=subprocess.run):
    """게시 파일을 커밋·푸시한다. 어느 단계든 실패하면 성공으로 보지 않는다."""
    commands = (
        ["git", "add", "--", *PUBLISH_FILES],
        ["git", "commit", "-q", "-m", message],
        ["git", "push", "-q"],
    )
    for cmd in commands:
        r = runner(cmd, cwd=ROOT, capture_output=True, text=True,
                   encoding="utf-8", errors="replace")
        detail = ((r.stderr or "") + "\n" + (r.stdout or "")).strip()
        if r.returncode:
            # 변경이 없는 재실행은 오류가 아니다. 이미 커밋된 미푸시분은 계속 push한다.
            if cmd[1] == "commit" and "nothing to commit" in detail.lower():
                continue
            return False, cmd[1], detail[:300]
    return True, "", ""


def _main():
    from operation_window import input_window_label, is_input_window
    if is_input_window():
        print(f"입력 보호시간({input_window_label()}) — 사본 생성·게시 생략")
        return
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
    tmp_out = OUT + ".tmp"
    with open(tmp_out, "w", encoding="utf-8") as f:
        json.dump(sealed, f, separators=(",", ":"))
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp_out, OUT)
    kb, rkb = os.path.getsize(OUT) / 1024, len(raw) / 1024

    print(f"  프로젝트코드 {len(d['codes'])} · 확인필요 {len(d['issues'])} · "
          f"AS {len(d['as'])} · 점검 {len(d['pm'])} · 정산 {len(d['settle'])} · 계산서 {len(d['erp'])}")
    print(f"  {rkb:.0f}KB → 줄여서 {len(packed)/1024:.0f}KB → 잠근 뒤 {kb:.0f}KB  ({OUT})")
    print(f"  미청구 {len(d.get('unbilled', []))}건 — 잠긴 사본 안에(앱 상단에 뜹니다)")

    if "--push" not in sys.argv:
        print("\n로컬 생성만 했습니다 — 고정 주소 반영: python cloud_publish.py --push")
        return

    # 잠근 파일만 올린다. 원본(raw)은 디스크에도 남기지 않는다.
    ok, stage, detail = git_publish(
        f"폰 사본 갱신 {d['gen']} (프로젝트코드 {len(d['codes'])}·확인필요 {len(d['issues'])})")
    if not ok:
        print(f"  git {stage} 실패:", detail)
        sys.exit(1)
    print("\n고정 주소에 반영했습니다 — PC를 꺼도 폰에서 열립니다.")
    print("  https://mulder4780.github.io/coupang-ecount-reconcile/")


def main():
    """게시 작업은 수동·앱서버·다른 AI 중 하나만 실행한다."""
    acquired_here = False
    if "--push" in sys.argv and (os.environ.get("CSOS_AI") or "").strip():
        import ai_claim
        from claim_guard import require
        me = (os.environ.get("CSOS_AI") or "").strip().lower()
        before = ai_claim.load().get("publish")
        require("publish", "폰 독립 사본 생성·Git 게시", who=me)
        acquired_here = not before
    try:
        return _main()
    finally:
        if acquired_here:
            from claim_guard import release
            release("publish", who=(os.environ.get("CSOS_AI") or "").strip().lower())


if __name__ == "__main__":
    main()
