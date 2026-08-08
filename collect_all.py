# -*- coding: utf-8 -*-
"""
collect_all.py — 미수집 자료를 **한 번에 다 긁어 보관**하고 보고서를 남긴다 (2026-08-08)

사용자 지시(2026-08-08): "미수집 데이터들 싹 다 긁어모아, 알고리즘 구성해서 두번 일
안하게 정리하고, 원본 데이터, 사진, 텍스트등 모두 가져와서 저장하고 보고서 작성해."

  python collect_all.py            # 무엇이 안 모였나만 본다(아무것도 안 건드림)
  python collect_all.py --run      # 사람 없이 되는 것을 전부 돌린다
  python collect_all.py --run --limit 500   # 이번 회차 상한(기본 400글)

★ '두 번 일 안 하게' 가 이 도구의 전부다
  같은 명령을 백 번 돌려도 이미 있는 것은 **파일을 열지도 않는다.** 그 판단을
  세 겹으로 둔다:
    ① 파일이 이미 있으면 건너뛴다 (archive_posts·fetch_images 가 원래 그렇다)
    ② 무엇을 언제 얼마나 했는지 `datalake.event` 에 남는다 — 다음 회차가 그걸 본다
    ③ 사람이 있어야 되는 것(로그인·브라우저)은 **돌리지 않고 보고서에 올린다**
  그래서 중간에 끊겨도 다시 부르면 **멈춘 자리에서** 이어진다. 처음부터 다시 하지 않는다.

★ 사람 없이 **안 되는 것**은 하지 않는다 (절대규칙 3)
  밴드·이카운트 수집은 사람 로그인이 먼저다. 그것을 자동으로 시도하면 로그인 화면을
  본문으로 착각해 캐시를 더럽힌다(2026-08-07 사고 — 마흔 건이 전부 같은 글이 됐다).
  그래서 여기서는 **이미 캐시에 든 것을 파일로 굳히는 일**만 한다.

무엇을 모으나
  · 밴드 글 → PDF + 텍스트 + 사진 한 세트  (`band/archive_posts.py`)
  · 밴드 문서 사진 → 명세서·계산서 사진 원본 (`band/fetch_images.py`)
  · 문서 사진 OCR → 금액·번호 읽기          (`band/doc_ocr.py`)
  · 위 전부를 보관소 DB 로                   (`datalake.py --scan`)
"""
import argparse
import glob
import json
import os
import subprocess
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT = os.path.join(ROOT, "reports", "미수집_수집보고서.md")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _py(*args, timeout=None):
    """자식 도구를 돌린다 → (성공, 마지막줄들). 하나가 죽어도 **멈추지 않는다.**"""
    cmd = [sys.executable] + list(args)
    try:
        r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout)
        tail = [l for l in (r.stdout or "").splitlines() if l.strip()][-6:]
        if r.returncode != 0:
            tail += [l for l in (r.stderr or "").splitlines() if l.strip()][-4:]
        return r.returncode == 0, tail
    except subprocess.TimeoutExpired:
        return False, [f"시간 초과({timeout}초) — 다음 회차가 이어서 한다"]
    except Exception as e:
        return False, [str(e)[:200]]


# ────────────────────────────────────────────────────────────────────
# 지금 무엇이 안 모였나 — 세는 것만 한다(파일을 만들지 않는다)
# ────────────────────────────────────────────────────────────────────
def survey(cache_dir=None, band_root=None):
    """캐시에 있는 것과 실제로 보관된 것을 맞대 본다 → 사람이 읽는 dict.

    두 자리를 인자로 받는 것은 **합성검증이 진짜 Z: 를 안 건드리게** 하려는 것이다.
    """
    import source_dirs as S

    posts, images = 0, 0
    for f in glob.glob(os.path.join(cache_dir or os.path.join(ROOT, "band", "cache"), "*.json")):
        if not os.path.basename(f)[:-5].isdigit():
            continue                      # raw_* 는 중간 산물이다 — 세지 않는다
        try:
            d = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        for v in (d.get("posts") or {}).values():
            # ★ 시각 없는 글은 세지 않는다. 밴드가 없는 번호에도 껍데기를 주기 때문에
            #   그것까지 세면 '영영 안 끝나는 미수집'이 생긴다(2026-08-07 사고).
            if isinstance(v, dict) and v.get("created_at"):
                posts += 1
                images += len(v.get("images") or [])

    have = {"pdf": 0, "txt": 0, "jpg": 0}
    band_root = band_root or getattr(S, "BAND_DIR", "")
    if band_root and os.path.isdir(band_root):
        for _r, _d, fs in os.walk(band_root):
            # ★ '날짜미상'은 **모수에 없는 글**이다 (2026-08-08 실측 873개).
            #   시각 없는 수확은 밴드가 없는 번호에 준 껍데기라 캐시 집계에서
            #   빠진다. 그런데 보관 수에는 들어가 있어서 '남음'이 실제보다 적게
            #   나왔다 — 다 됐다고 착각하게 만드는 종류의 오차다.
            if "날짜미상" in _r:
                continue
            for fn in fs:
                e = os.path.splitext(fn)[1].lower()
                if e == ".pdf":
                    have["pdf"] += 1
                elif e == ".txt":
                    have["txt"] += 1
                elif e in (".jpg", ".jpeg", ".png"):
                    have["jpg"] += 1

    return {
        "밴드글_캐시": posts,
        "밴드글_보관": have["pdf"],
        "밴드글_남음": max(0, posts - have["pdf"]),
        "사진_URL": images,
        "사진_보관": have["jpg"],
        "사진_남음": max(0, images - have["jpg"]),
        "텍스트_보관": have["txt"],
    }


def human_gaps():
    """**사람이 있어야** 메울 수 있는 구멍 → [(무엇, 왜, 어떻게)].

    자동으로 시도하지 않는다 — 로그인 화면을 본문으로 착각해 캐시를 더럽힌 전례가 있다.
    """
    out = []
    try:
        import erp_grab
        s = erp_grab.survey()
        limit = erp_grab.DEFAULT_LIMIT_DAYS
        today = datetime.now().date()
        for kind, (latest, cnt) in sorted(s.items()):
            if not latest:
                continue
            try:
                d = datetime.strptime(str(latest)[:10], "%Y-%m-%d").date()
            except ValueError:
                continue
            late = (today - d).days
            if late > limit:
                out.append((f"ERP {erp_grab.KIND_LABEL.get(kind, kind)}",
                            f"최신 {d} · {late}일 밀림",
                            "로그인된 ERP 탭에서 erp_grab.py --all"))
    except Exception as e:
        out.append(("ERP 밀림 판정", f"못 쟀다: {str(e)[:80]}", "erp_grab.py 를 직접 볼 것"))
    return out


# ────────────────────────────────────────────────────────────────────
# 실제로 긁는다 — 사람 없이 되는 것만
# ────────────────────────────────────────────────────────────────────
STEPS = [
    ("밴드글 보관(PDF·텍스트·사진)", ["band/archive_posts.py", "--limit", "{limit}"], 3600),
    ("밴드 문서사진 내려받기",        ["band/fetch_images.py"],                        1800),
    ("문서사진 OCR",                 ["band/doc_ocr.py"],                             1800),
    ("보관소 DB 흡수",               ["datalake.py", "--scan"],                       3600),
    # ★ 파일을 세는 것과 **그 안의 한 건씩을 아는 것**은 다르다(분담판 [24]).
    #   `--scan` 은 "ERP 엑셀 102개가 있다"까지다. 앱이 화면에 뿌리고 사람이 검색하려면
    #   전표 한 줄·계산서 한 건이 record 로 들어와 있어야 한다. 이미 뜯어 본 파일은
    #   건너뛰므로(크기·수정시각이 같으면) 매일 돌려도 새로 받은 것만 뜯는다.
    ("ERP 엑셀 → 건별 기록",         ["datalake.py", "--erp"],                        1800),
    ("밴드 글 → 건별 기록",          ["datalake.py", "--band"],                       1800),
]


def run(limit=400, only=None):
    """단계를 차례로 돌린다 → [{단계,결과,초,끝줄}].

    ★ **기록이 수집을 막지 않는다.** 보관소 DB 가 잠겨 있어도(다른 세션이 주사 중일
      수 있다) 긁는 일은 그대로 간다. 2026-08-08 실측: 로그 한 줄을 못 써서
      'database is locked' 로 회차 전체가 죽었다 — 꼬리가 몸통을 흔든 것이다.
    """
    con = None
    try:
        import datalake as D
        con = D.connect()
    except Exception as e:
        print(f"  i 보관소 DB 를 못 열었다({str(e)[:60]}) — 기록 없이 수집만 진행한다")

    def note(**kw):
        if con is None:
            return
        try:
            import datalake as D
            D.log(con, "collect", kw.pop("action"), **kw)
            con.commit()
        except Exception:
            pass          # 기록 실패는 수집 실패가 아니다

    done = []
    try:
        note(action="collect_all.start", detail={"limit": limit})
        for name, args, timeout in STEPS:
            if only and not any(o in name for o in only):
                continue
            argv = [a.format(limit=limit) for a in args]
            print(f"▶ {name} …")
            t0 = time.time()
            ok, tail = _py(*argv, timeout=timeout)
            초 = round(time.time() - t0, 1)
            for line in tail:
                print(f"    {line}")
            # ★ 실패도 반드시 남긴다. 남기지 않으면 다음 회차가 '다 됐다'고 읽는다.
            note(action="collect_all.step", ok=ok,
                 detail={"단계": name, "초": 초, "끝줄": tail[-2:]})
            done.append({"단계": name, "결과": "됨" if ok else "실패", "초": 초, "끝줄": tail})
        note(action="collect_all.end",
             ok=all(d["결과"] == "됨" for d in done), detail={"단계수": len(done)})
    finally:
        if con is not None:
            con.close()
    return done


def write_report(before, after, done, humans):
    """사람이 읽는 한 장. **무엇이 남았는지**가 맨 위에 온다."""
    L = []
    L.append("# 미수집 자료 수집 보고서")
    L.append("")
    L.append(f"- 기준: {datetime.now():%Y-%m-%d %H:%M}")
    L.append("- 이 문서는 `collect_all.py` 가 만든다. 같은 명령을 다시 돌려도 "
             "**이미 있는 것은 건드리지 않는다**(두 번 일 하지 않게).")
    L.append("")

    if humans:
        L.append("## ★ 사람이 있어야 되는 것 — 여기부터")
        L.append("")
        L.append("자동으로 시도하지 않았다. 로그인 화면을 본문으로 착각해 캐시를 "
                 "더럽힌 전례가 있어(2026-08-07), **사람 로그인 뒤에만** 긁는다.")
        L.append("")
        L.append("| 무엇 | 왜 | 어떻게 |")
        L.append("|---|---|---|")
        for a, b, c in humans:
            L.append(f"| {a} | {b} | `{c}` |")
        L.append("")

    L.append("## 얼마나 모였나")
    L.append("")
    L.append("| 항목 | 전 | 후 | 늘어난 것 |")
    L.append("|---|---:|---:|---:|")
    for k in ("밴드글_보관", "텍스트_보관", "사진_보관", "밴드글_남음", "사진_남음"):
        b, a = before.get(k, 0), after.get(k, 0)
        L.append(f"| {k.replace('_', ' ')} | {b:,} | {a:,} | {a - b:+,} |")
    L.append("")
    L.append(f"- 캐시에 든 밴드 글 **{after.get('밴드글_캐시', 0):,}건** · "
             f"사진 URL **{after.get('사진_URL', 0):,}개**가 모수다.")
    if after.get("밴드글_남음"):
        L.append(f"- **아직 {after['밴드글_남음']:,}글 · 사진 {after.get('사진_남음', 0):,}장이 남았다.** "
                 "한 회차 상한이 있어서다 — `--run` 을 다시 부르면 이어서 한다.")
    else:
        L.append("- 밴드 글은 **전부 보관됐다.**")
    L.append("")

    if done:
        L.append("## 이번 회차에 돌린 것")
        L.append("")
        L.append("| 단계 | 결과 | 걸린 시간 |")
        L.append("|---|---|---:|")
        for d in done:
            L.append(f"| {d['단계']} | {d['결과']} | {d['초']}초 |")
        L.append("")
        for d in done:
            if d["결과"] != "됨":
                L.append(f"> **{d['단계']} 실패** — {' / '.join(d['끝줄'][-2:])}")
        L.append("")

    L.append("## 다시 돌리려면")
    L.append("")
    L.append("```")
    L.append("python ecount/collect_all.py            # 무엇이 남았나만 본다")
    L.append("python ecount/collect_all.py --run      # 남은 것을 이어서 긁는다")
    L.append("```")
    L.append("")
    L.append("기록은 보관소 DB 에 남는다 — `python ecount/datalake.py --log area=collect`")

    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    open(REPORT, "w", encoding="utf-8").write("\n".join(L) + "\n")
    return REPORT


def main(argv=None):
    ap = argparse.ArgumentParser(description="미수집 자료를 한 번에 긁어 보관하고 보고서를 남긴다")
    ap.add_argument("--run", action="store_true", help="실제로 긁는다(없으면 세기만)")
    ap.add_argument("--limit", type=int, default=400, help="이번 회차 밴드 글 상한")
    ap.add_argument("--only", nargs="*", help="특정 단계만 (예: 사진 OCR)")
    a = ap.parse_args(argv)

    before = survey()
    humans = human_gaps()

    print("미수집 현황")
    print(f"  밴드 글  캐시 {before['밴드글_캐시']:,} · 보관 {before['밴드글_보관']:,}"
          f" · 남음 {before['밴드글_남음']:,}")
    print(f"  사진     URL  {before['사진_URL']:,} · 보관 {before['사진_보관']:,}"
          f" · 남음 {before['사진_남음']:,}")
    print(f"  텍스트   보관 {before['텍스트_보관']:,}")
    for x, y, _z in humans:
        print(f"  ★ 사람 필요: {x} — {y}")

    done = []
    if a.run:
        done = run(limit=a.limit, only=a.only)
    else:
        print("\n(세기만 했다 — 실제로 긁으려면 --run)")

    after = survey() if a.run else before
    path = write_report(before, after, done, humans)
    print(f"\n보고서: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
