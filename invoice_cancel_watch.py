# -*- coding: utf-8 -*-
"""
invoice_cancel_watch.py : 쿠팡 세금계산서 취소(마이너스) 자동 검사
====================================================================
형님 지시(2026-11-18): "취소 계산서 자동 검사 만들어"

  python invoice_cancel_watch.py            # 검사하고 reports/계산서_취소검사.{json,md} 를 쓴다
  python invoice_cancel_watch.py --print    # 쓰지 않고 요약만

왜 : 2026-11-18 손으로 대 보니 8월 '미입금 14건'이 전부 발행→같은 날 취소 7쌍이었다.
     그리고 그 안에 **돈은 받았는데 살아 있는 계산서가 없는** 2건(공임 42,296,320 ·
     철거 7,694,830)이 숨어 있었다. 취소 짝을 걸러 내는 단계가 어디에도 없었다.

자료 : TF 세션이 내는 공유 DB(`_공유DB/자재수불/자재수불_공유_YYYYMMDD.db[.zip]`).
       읽기만 한다 · ERP·위하고에 쓰지 않는다 · 우리 정본에 되밀지 않는다(역수입 금지).

갈래(계산서 한 쌍 = 같은 거래처·같은 품명·금액 +A / -A)
  · 취소됐는데_입금됨   : 같은 금액 입금이 있고, 취소일 뒤 같은 금액 재발행이 없다  ← 가장 위험
  · 취소_재발행없음     : 입금도 재발행도 없다 (납품 취소인지 청구 누락인지 사람이 가른다)
  · 취소_재발행됨       : 취소일 뒤 같은 금액 계산서가 다시 있다 (정보)
  · 짝없는_마이너스     : 마이너스인데 같은 품명·금액 플러스가 없다

★ 금액으로만 맞춘다 : 우연히 같은 금액일 수 있다. 그래서 '확정'이라 적지 않고 '후보'라 적는다([169]).
★ 자료를 못 읽으면 0건이 아니라 '확인못함'이다([169]).
"""
import os, sys, json, glob, zipfile, sqlite3, shutil
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
SHARE = os.environ.get("ULH_SHARE_DB_DIR") or os.path.join(os.path.dirname(os.path.dirname(ROOT)), "_공유DB")
SRC_DIR = os.path.join(SHARE, "자재수불")
CACHE = os.path.join(ROOT, "reports", ".tf_share_cache")
OUT_JSON = os.path.join(ROOT, "reports", "계산서_취소검사.json")
OUT_MD = os.path.join(ROOT, "reports", "계산서_취소검사.md")
SINCE = "20260101"


def find_db(src_dir=None):
    """가장 최근 공유 DB 경로를 돌려준다. zip 이면 캐시에 풀어(새것일 때만) 그 경로를 준다."""
    d = src_dir or SRC_DIR
    cands = sorted(glob.glob(os.path.join(d, "자재수불_공유_*.db")) +
                   glob.glob(os.path.join(d, "자재수불_공유_*.db.zip")))
    if not cands:
        return None
    best = cands[-1]
    if best.endswith(".db"):
        return best
    os.makedirs(CACHE, exist_ok=True)
    target = os.path.join(CACHE, os.path.basename(best)[:-4])
    if os.path.exists(target) and os.path.getmtime(target) >= os.path.getmtime(best):
        return target
    with zipfile.ZipFile(best) as z:
        name = [n for n in z.namelist() if n.endswith(".db")][0]
        tmp = target + ".tmp"
        with z.open(name) as s, open(tmp, "wb") as t:
            shutil.copyfileobj(s, t)
    os.replace(tmp, target)
    for old in glob.glob(os.path.join(CACHE, "자재수불_공유_*.db")):
        if old != target:
            try:
                os.remove(old)
            except OSError:
                pass
    return target


def load(con, since=SINCE):
    inv = con.execute(
        "select 일자, cast(합계 as integer), 품명, 거래처명 from wehago_tax_invoice_out "
        "where 거래처명 like '%쿠팡%' and 줄종류='일반' and 일자>=? "
        "group by coalesce(sq_sbook, 원본파일 || 원본행번호) order by 일자", (since,)).fetchall()
    dep = con.execute(
        "select 일자, cast(입금 as integer) from wehago_bank where coalesce(겹침제외,0)=0 "
        "and (적요 like '%쿠팡%' or 거래처명 like '%쿠팡%') and coalesce(입금,0)>0").fetchall()
    inv_end = con.execute("select max(일자) from wehago_tax_invoice_out").fetchone()[0]
    bank_end = con.execute("select max(일자) from wehago_bank").fetchone()[0]
    return inv, dep, inv_end, bank_end


def judge(inv, dep):
    """inv: [(일자 YYYYMMDD, 합계, 품명, 거래처)], dep: [(일자 YYYY-MM-DD, 입금)]"""
    plus = [i for i in inv if i[1] > 0]
    minus = [i for i in inv if i[1] < 0]
    # 1) 취소 짝을 먼저 전부 묶는다 : 그래야 '다른 취소 건'을 재발행으로 잘못 세지 않는다
    used, pairs = set(), []
    for m in minus:
        k = next((k for k, p in enumerate(plus) if k not in used and p[1] == -m[1]
                  and p[2] == m[2] and p[3] == m[3]), None)
        if k is not None:
            used.add(k)
        pairs.append((m, k))
    live = [p for k, p in enumerate(plus) if k not in used]   # 취소되지 않은 계산서
    taken = set()
    out = []
    for m, k in pairs:
        amt = -m[1]
        if k is None:
            out.append({"갈래": "짝없는_마이너스", "취소일": m[0], "금액": amt, "품명": m[2]})
            continue
        # 2) 재발행 = 살아 있는 계산서 중 같은 금액 · 취소일 이후(같은 날 포함) · 아직 안 쓴 것
        #    살아 있는 것만 보지 않는다 : 당일 수정(+A -A +A) 뒤 다시 취소하고 재발행한 사슬은
        #    가운데 장도 이미 취소돼 있다(2026-11-18 소독기 송파1·북김해 실측, TF 판정).
        j = next((j for j, p in enumerate(plus) if j != k and j not in taken and p[1] == amt
                  and p[0] >= m[0] and (p[0], j) > (plus[k][0], k)
                  and (j not in used or (p[2] == m[2] and p[3] == m[3]))), None)  # 취소된 장은 같은 품명만
        reissue = []
        if j is not None:
            taken.add(j)
            reissue = [plus[j][0] + " " + plus[j][2]]
        paid = [d for d in dep if d[1] == amt and d[0].replace("-", "") >= plus[k][0]]
        row = {"발행일": plus[k][0], "취소일": m[0], "금액": amt, "품명": m[2],
               "재발행": reissue, "같은금액입금": [d[0] for d in paid]}
        if reissue:
            row["갈래"] = "취소_재발행됨"
        elif paid:
            row["갈래"] = "취소됐는데_입금됨"
        else:
            row["갈래"] = "취소_재발행없음"
        out.append(row)
    return out


def pair_live(inv):
    """취소 짝(같은 품명·거래처 +A/-A)을 빼고 살아 있는 계산서만 돌려준다."""
    plus = [i for i in inv if i[1] > 0]
    gone = set()
    for m in (i for i in inv if i[1] < 0):
        k = next((k for k, p in enumerate(plus) if k not in gone and p[1] == -m[1]
                  and p[2] == m[2] and p[3] == m[3]), None)
        if k is not None:
            gone.add(k)
    return [p for k, p in enumerate(plus) if k not in gone]


def reconcile(inv, dep, dep_since="2026-01-01"):
    """입금마다 같은 금액 · 입금일 이전 · 아직 안 쓴 살아 있는 계산서 중 가장 이른 것에 맞춘다.
    금액으로만 맞춘 후보다([169]). 돌려주는 것: (입금행들, 입금 없는 살아 있는 계산서)"""
    live = pair_live(inv)
    used, rows = set(), []
    for d, a in sorted(x for x in dep if x[0] >= dep_since):
        m = [k for k, i in enumerate(live) if k not in used and i[1] == a and i[0] <= d.replace("-", "")]
        if m:
            used.add(m[0])
            i = live[m[0]]
            rows.append({"입금일": d, "입금액": a, "계산서일": i[0], "품명": i[2], "후보수": len(m)})
        else:
            rows.append({"입금일": d, "입금액": a, "계산서일": "", "품명": "(맞는 계산서 없음)", "후보수": 0})
    unpaid = [i for k, i in enumerate(live) if k not in used and i[0] >= dep_since.replace("-", "")]
    return rows, unpaid


def stale_days(bank_end, today=None):
    """통장 자료 끝날이 오늘보다 며칠 뒤처졌나. 못 읽으면 None(모름)."""
    try:
        e = datetime.strptime(str(bank_end)[:10], "%Y-%m-%d").date()
    except Exception:
        return None
    return ((today or datetime.now().date()) - e).days


STALE_MAX = 3
OUT_DEP_MD = os.path.join(ROOT, "reports", "쿠팡_입금_대조.md")
OUT_DEP_CSV = os.path.join(ROOT, "reports", "쿠팡_입금_대조.csv")


def write_deposit(rows, unpaid, bank_end, src, note=""):
    import csv
    sd = stale_days(bank_end)
    warn = ("통장 자료 끝날을 못 읽음(모름)" if sd is None else
            ("통장 자료가 %d일째 멈춤(끝 %s) : 그 뒤 입금은 아직 안 보임" % (sd, bank_end)) if sd > STALE_MAX else "")
    ok = [r for r in rows if r["계산서일"]]
    lines = ["# 쿠팡 입금 ↔ 계산서 대조", "",
             "만든 때 %s · 자료 %s · 통장 끝 %s" % (datetime.now().isoformat(timespec="seconds"), src, bank_end), ""]
    if note:
        lines += ["**확인못함** : " + note, ""]
    if warn:
        lines += ["**경고** : " + warn, ""]
    lines += ["- 입금 %d건 %s원 · 맞음 %d · 안 맞음 %d · 후보 여럿 %d" % (
        len(rows), format(sum(r["입금액"] for r in rows), ","), len(ok), len(rows) - len(ok),
        sum(1 for r in ok if r["후보수"] > 1)),
        "- 입금 없는 살아 있는 계산서 %d건 %s원" % (len(unpaid), format(sum(i[1] for i in unpaid), ",")), ""]
    for i in unpaid:
        lines.append("  - %s · %s원 · %s" % (i[0], format(i[1], ","), i[2][:50]))
    lines += ["", "금액으로만 맞춘 후보입니다."]
    with open(OUT_DEP_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    with open(OUT_DEP_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["입금일", "입금액", "계산서일", "계산서품명", "같은금액후보수"])
        for r in rows:
            w.writerow([r["입금일"], r["입금액"], r["계산서일"], r["품명"], r["후보수"]])
        w.writerow([])
        w.writerow(["입금 없는 살아 있는 계산서(취소 짝 제외)"])
        for i in unpaid:
            w.writerow(["", i[1], i[0], i[2]])
    return warn


ORDER = ["취소됐는데_입금됨", "취소_재발행없음", "짝없는_마이너스", "취소_재발행됨"]
WORDS = {"취소됐는데_입금됨": "돈은 받았는데 살아 있는 계산서가 없음(재발행 확인 필요)",
         "취소_재발행없음": "취소 뒤 다시 낸 계산서도 입금도 없음(납품 취소인지 청구 누락인지 확인)",
         "짝없는_마이너스": "마이너스인데 짝이 되는 원 계산서를 못 찾음",
         "취소_재발행됨": "취소 뒤 같은 금액으로 다시 발행됨(정보)"}


def run(write=True, src_dir=None):
    res = {"만든때": datetime.now().isoformat(timespec="seconds"), "건": [], "확인못함": ""}
    try:
        p = find_db(src_dir)
        if not p:
            raise FileNotFoundError("공유 DB 없음: %s" % (src_dir or SRC_DIR))
        con = sqlite3.connect("file:%s?mode=ro" % p.replace("\\", "/"), uri=True)
        try:
            inv, dep, ie, be = load(con)
        finally:
            con.close()
        res.update({"자료": os.path.basename(p), "계산서끝": ie, "통장끝": be, "건": judge(inv, dep)})
        if write:
            con = sqlite3.connect("file:%s?mode=ro" % p.replace("\\", "/"), uri=True)
            try:
                inv_all = load(con, since="20250101")[0]
            finally:
                con.close()
            rows, unpaid = reconcile(inv_all, dep)
            res["입금대조"] = {"입금": len(rows), "안맞음": sum(1 for r in rows if not r["계산서일"]),
                            "미입금": len(unpaid), "경고": write_deposit(rows, unpaid, be, os.path.basename(p))}
    except Exception as e:  # 못 읽으면 0건이 아니라 확인못함
        res["확인못함"] = "%s: %s" % (type(e).__name__, e)
    cnt = {g: [r for r in res["건"] if r["갈래"] == g] for g in ORDER}
    res["집계"] = {g: len(v) for g, v in cnt.items()}
    lines = ["# 쿠팡 세금계산서 취소 검사", "",
             "만든 때 %s · 자료 %s · 계산서 끝 %s · 통장 끝 %s" % (
                 res["만든때"], res.get("자료", "-"), res.get("계산서끝", "-"), res.get("통장끝", "-")), ""]
    if res["확인못함"]:
        lines += ["**확인못함** : " + res["확인못함"], ""]
    for g in ORDER:
        lines.append("## %s : %d건 (%s)" % (g, len(cnt[g]), WORDS[g]))
        for r in cnt[g]:
            lines.append("- %s 취소 · %s원 · %s%s%s" % (
                r["취소일"], format(r["금액"], ","), r["품명"][:50],
                (" · 입금 " + ",".join(r.get("같은금액입금", []))) if r.get("같은금액입금") else "",
                (" · 재발행 " + "; ".join(r["재발행"])) if r.get("재발행") else ""))
        lines.append("")
    lines.append("금액으로만 맞춘 후보입니다. 같은 금액의 다른 돈일 수 있습니다.")
    if write:
        os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        with open(OUT_MD, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    print("계산서 취소 검사 : " + (("확인못함 " + res["확인못함"]) if res["확인못함"] else
          " · ".join("%s %d" % (g, n) for g, n in res["집계"].items())))
    return res


if __name__ == "__main__":
    run(write="--print" not in sys.argv)
