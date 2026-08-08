# -*- coding: utf-8 -*-
"""접수했다가 취소된 건을 찾아낸다 (2026-08-08 지시).

사용자 지시: **"접수 했다가 접수 취소하는 경우도 많은데 이것도 잡아내는 알고리즘 추가해"**

실제 사례(밴드): 접수 글이 올라간 뒤 댓글로
    "통화 완료 했습니다 / 작동 원활함. 접수 취소 하세요"

**이것은 조용한 사고다.** 취소된 건은 아무도 안 가는데 원장에는 접수 그대로 남아
`AS 미실시`·`정기점검 대기`에 계속 얹힌다. 숫자가 비어 보이지 않으니 아무도 이상하다
하지 않고, 기사에게는 영영 안 없어지는 밀린 일로 보인다.

무엇이 막고 있었나 — 두 가지였다:
  ① 판정이 `"접수취소" in 본문` **한 줄**이라 사람이 쓰는 `접수 취소`(띄어쓰기)를 놓쳤다.
     실측 밴드 8,561글에서 48건만 잡히던 것이 60건이 됐다(+12, 전부 `✅ 접수 취소`).
  ② 취소는 **댓글**로 온다. 접수 글은 이미 올라간 뒤라 고칠 자리가 댓글밖에 없는데,
     예전 밴드 캐시는 `comment_count` 숫자만 담고 **본문이 없었다**.
     2026-08-08 에 담는 쪽 둘(화면 긁기 `grab_posts.js`→`convert_dump` · API
     `band_sync`)이 `comments: [{author, created_at, content}]` 로 담기 시작했다 —
     이 도구는 고칠 것이 없었다. 읽는 자리를 미리 하나로 두었기 때문이다(검증 [162]).
     그래도 **접힌 댓글은 다 안 펴질 수 있다.** 그래서 못 읽은 글이 몇 건인가를
     계속 함께 보고한다 — 사각지대를 0건이라 말하지 않기 위해서다.
     캐시를 채우는 수집 자체는 'CSOS 리서치 및 자료 수집' 세션 몫이다.
  ★★ 그 사각지대 계기마저 `comment_count` 에 기대고 있었다 (2026-08-08 저녁 실측).
     캐시 10,312글 중 `comment_count>0` 은 **6글**, 댓글 본문은 **0글**이다 —
     밴드에 댓글이 없어서가 아니라 **수집기가 그 숫자를 안 담아서**다. 그래서 계기가
     "사각지대 0건"이라 말했다. **재는 도구가 같은 결측에 눈이 멀어 있었다.**
     이제 `comments` 키가 아예 없는 글(= 한 번도 안 들여다본 글)도 사각지대로 센다.
     고친 뒤 실측 사각지대는 **8,259건 / 8,561건**이다. 그 값이 절반을 넘는 동안은
     이 리포트를 "취소가 이것뿐"이라는 뜻으로 읽으면 안 된다(리포트가 스스로 경고한다).

엑셀은 열지 않는다. 바꿀 것은 `ledger_db.enqueue()` 로 넣고 11:00·15:00 회차가 반영한다.

실행:
  python cancel_watch.py             # 무엇이 걸리는지 보기만(기본)
  python cancel_watch.py --queue     # 원장 반영 대기열에 넣기
"""
import sys, os, re, json, glob, argparse
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from band_extract import cancel_hit, comment_text, cancel_blind_count, RE_PRJ  # noqa: E402

CACHE_DIR = os.path.join(ROOT, "band", "cache")
REPORT = os.path.join(os.environ.get("COUPANG_REPORT_DIR")
                      or os.path.join(ROOT, "reports"), "접수취소_확인.md")

# 이미 끝났거나 이미 취소로 적힌 건은 건드릴 것이 없다.
_SETTLED = ("완료", "취소", "철회")


def _snippet(text, width=34):
    """취소라고 판단한 대목만 잘라 근거로 남긴다 — 사람이 한눈에 확인할 수 있게."""
    s = str(text or "")
    m = re.search(r".{0,%d}(접\s*수[^가-힣A-Za-z0-9\n]{0,3}(를|건|은|이)?\s*(요\s*청)?\s*취\s*소"
                  r"|오\s*접\s*수|중\s*복\s*접\s*수|접\s*수\s*(를\s*)?(철\s*회|반\s*려)).{0,%d}"
                  % (width, width), s)
    return re.sub(r"\s+", " ", (m.group(0) if m else s[:width * 2])).strip()


def scan_band(quiet=False):
    """밴드 캐시에서 취소로 읽히는 글을 모은다. **긁지 않는다 — 읽기만 한다.**"""
    hits, blind, total = {}, 0, 0
    for fp in sorted(glob.glob(os.path.join(CACHE_DIR, "*.json"))):
        base = os.path.basename(fp)
        if base.startswith("raw"):
            continue
        try:
            d = json.load(open(fp, encoding="utf-8"))
        except Exception:
            continue
        posts = d.get("posts") or {}
        blind += cancel_blind_count(posts)
        bname = str(d.get("band_name") or os.path.splitext(base)[0])
        for no, p in posts.items():
            if not isinstance(p, dict):
                continue
            total += 1
            body, cmt = p.get("content") or "", comment_text(p)
            where = "댓글" if cancel_hit(cmt) else ("본문" if cancel_hit(body) else "")
            if not where:
                continue
            m = RE_PRJ.search(body)
            if not m:
                continue                      # 프로젝트NO 가 없으면 원장에 붙일 수 없다
            ts = p.get("created_at")
            try:
                day = datetime.fromtimestamp(int(ts) / 1000).strftime("%Y-%m-%d") if ts else ""
            except Exception:
                day = ""
            prj = m.group(1).upper()
            # 같은 프로젝트가 여러 글에 나오면 **가장 나중 글**이 최신 뜻이다.
            cur = hits.get(prj)
            if cur and cur["게시일"] > day:
                continue
            hits[prj] = {"프로젝트NO": prj, "밴드": bname, "게시글": no, "게시일": day,
                         "자리": where, "근거": _snippet(cmt if where == "댓글" else body)}
    if not quiet:
        print(f"  밴드 {total}글 훑음 — 취소로 읽히는 건 {len(hits)}건"
              f" · 댓글을 못 읽어 놓쳤을 수 있는 글 {blind}건")
    return hits, blind, total


def open_ledger_rows():
    """원장에서 **아직 안 끝난** 돌발AS·정기점검을 프로젝트NO 로 뽑는다."""
    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    import app_server
    out = {}
    works = app_server.get_works()
    for kind, sheet, key_col in (("as", "02_돌발AS접수", "접수ID"),
                                 ("pm", "04_정기점검", "점검ID")):
        for r in works.get(kind) or []:
            prj = str(r.get("프로젝트NO") or "").strip().upper()
            if not prj:
                continue
            state = str(r.get("진행상태") or r.get("점검상태") or "")
            done = str(r.get("작업완료일") or r.get("실제점검일") or "").strip()
            if done or any(s in state for s in _SETTLED):
                continue
            out.setdefault(prj, []).append(
                {"sheet": sheet, "key_col": key_col,
                 "key": str(r.get(key_col) or "").strip(),
                 "업무": "돌발AS" if kind == "as" else "정기점검",
                 "캠프명": r.get("캠프명") or "", "상태": state})
    return out


def build(quiet=False):
    hits, blind, total = scan_band(quiet=quiet)
    ledger = open_ledger_rows()
    rows = []
    for prj, h in sorted(hits.items()):
        for t in ledger.get(prj) or []:
            if not t["key"]:
                continue
            rows.append({**h, **t})
    return rows, hits, blind, total


def write_report(rows, hits, blind, total=0):
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write(f"# 접수 취소 확인 — {now}\n\n")
        f.write("접수했다가 **취소된 것으로 읽히는** 건입니다. 원장에는 아직 접수 상태로\n"
                "남아 `AS 미실시`·`정기점검 대기`에 얹혀 있습니다.\n\n")
        f.write(f"- 밴드에서 취소로 읽힌 건: **{len(hits)}건**\n")
        f.write(f"- 그중 원장이 아직 안 끝낸 건: **{len(rows)}건** ← 정리 대상\n")
        f.write(f"- 댓글을 못 읽어 **놓쳤을 수 있는 글: {blind}건** / 전체 {total}건"
                "\n  (댓글을 한 번도 안 들여다본 글 + 달린 수만큼 본문이 없는 글."
                " 긁는 것은 수집 세션 몫입니다)\n")
        if total and blind * 2 > total:
            # ★ 이 숫자가 절반을 넘으면 '몇 건 놓쳤나'가 아니라 **이 리포트를 믿을 수
            #   있나**의 문제다. 위의 '취소로 읽힌 건'은 읽은 것 중에서만 센 값이다.
            f.write("\n> ⚠ 사각지대가 절반을 넘습니다 — 위 '취소로 읽힌 건'은 **읽은 글**"
                    " 안에서만 센 숫자입니다. 댓글 수집이 채워지기 전까지 이 리포트를"
                    " '취소가 이것뿐'이라는 뜻으로 읽으면 안 됩니다.\n")
        f.write("\n")
        if rows:
            f.write("| 프로젝트NO | 업무 | 캠프 | 지금 상태 | 취소 근거 | 자리 | 밴드/글 | 게시일 |\n")
            f.write("|---|---|---|---|---|---|---|---|\n")
            for r in rows:
                f.write(f"| {r['프로젝트NO']} | {r['업무']} | {r['캠프명']} | "
                        f"{r['상태'] or '-'} | {r['근거']} | {r['자리']} | "
                        f"{r['밴드']}/{r['게시글']} | {r['게시일']} |\n")
        else:
            f.write("정리할 건이 없습니다.\n")
        f.write("\n※ 엑셀은 열지 않습니다. `--queue` 로 넣으면 11:00·15:00 회차가 반영합니다.\n")
    return REPORT


def queue(rows):
    """진행상태를 '취소'로 바꾸도록 대기열에 넣는다(엑셀은 회차가 연다)."""
    import ledger_db
    items = []
    for r in rows:
        items.append({
            "sheet": r["sheet"], "key_col": r["key_col"], "key": r["key"],
            "col": "진행상태", "value": "취소", "vtype": "text",
            # ★ 이미 사람이 무언가 적어 둔 칸은 덮지 않는다.
            "only_if_empty": True,
            "evidence": f"밴드 {r['밴드']} {r['게시글']}글 {r['자리']} — {r['근거']}",
        })
    return ledger_db.enqueue(items, source="cancel_watch",
                             ingest_prefix="cancel") if items else 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="접수 취소 건 찾기")
    ap.add_argument("--queue", action="store_true", help="원장 반영 대기열에 넣는다")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)
    rows, hits, blind, total = build(quiet=a.quiet)
    path = write_report(rows, hits, blind, total)
    added = queue(rows) if a.queue else 0
    if not a.quiet:
        print(f"  취소로 읽힌 {len(hits)}건 중 원장이 아직 안 끝낸 {len(rows)}건"
              + (f" → 대기열 {added}건 추가" if a.queue else " (보기만 — 넣으려면 --queue)"))
        if blind:
            print(f"  ※ 댓글을 못 읽은 글 {blind}건 / 전체 {total}건"
                  " — 한 번도 안 들여다본 글까지 센다. 그 글의 취소는 아직 못 본다")
        print(f"  {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
