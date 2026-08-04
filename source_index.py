# -*- coding: utf-8 -*-
"""
source_index.py — 모든 원본 자료를 **한 색인**으로 모아 클릭 한 번에 찾고 열게 한다

사용자 지시(2026-08-05): "모든 원본 데이터 클릭 한번으로 찾고 열 수 있게 필터하고 구조화해."

지금 원본은 Z: 아래 여러 폴더(ERP 내보내기·밴드 원본·쿠팡 PO·서류·카톡·입금…)에 흩어져
있고, 이름이 무작위(0NSKITA3APTYVRL.xlsx)인 것도 많아 사람이 못 찾는다. 그래서
**파일 하나 = 한 줄**로 구조화한 색인을 만든다.

한 줄에 담기는 것
  path·name·kind(자료 종류)·프로젝트NO·전표/글번호·날짜·거래처(캠프)·크기·수정시각
  · kind 는 inbox_scan.classify(내용 판별)와 폴더 규칙을 함께 쓴다.
  · 프로젝트NO·전표번호는 **파일명에서** 뽑는다(어제 만든 건별 보관 파일명 규칙 덕에
    밴드 글·거래명세서는 이름만으로 무엇인지 알 수 있다).
  · 무거운 내용 판별은 xlsx 에만, 그것도 캐시(mtime+size)로 한 번만 한다.

산출
  reports/원본색인.json   — 앱·도구가 읽는 정본
  reports/원본색인.csv    — 엑셀에서 바로 필터할 사람용

  python source_index.py            # 새 파일만 판별(캐시 사용)
  python source_index.py --rescan   # 전부 다시 판별
"""
import argparse
import csv
import json
import os
import re
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT_JSON = os.path.join(ROOT, "reports", "원본색인.json")
OUT_CSV = os.path.join(ROOT, "reports", "원본색인.csv")
CACHE = os.path.join(ROOT, "db", "source_index_cache.json")
UJ = re.compile(r"UJ\d{7}")
SLIP = re.compile(r"\[(\d{4}-\d{2}-\d{2}-\d+)\]")     # 명세서 건별 PDF 이름
POST = re.compile(r"\[(\d{3,6})\]")                    # 밴드 글 보관 이름
DATE = re.compile(r"(20\d{2})[-/.]?(\d{2})[-/.]?(\d{2})")
PO = re.compile(r"PO\d{6,}", re.I)
SKIP_EXT = {".tmp", ".lnk", ".ini", ".db"}

# 폴더 이름으로 1차 분류 — 내용 판별보다 싸고, 사람이 이해하는 갈래와 같다.
FOLDER_KIND = [
    ("1. ERP 내보내기", "ERP"),
    ("거래명세서_건별", "ERP 거래명세서(건별 PDF)"),
    ("4. 밴드 원본", "밴드"),
    ("게시글보관", "밴드 게시글(보관)"),
    ("문서사진", "밴드 문서사진"),
    ("2. 쿠팡 PO", "쿠팡 PO"),
    ("3. 카카오톡", "카톡"),
    ("5. 입금", "입금"),
    ("6. 서류", "서류"),
    ("정기점검", "정기점검"),
    ("9. 미분류", "미분류"),
    ("100. 업로드용", "투입 대기"),
]


def load_cache():
    try:
        return json.load(open(CACHE, encoding="utf-8"))
    except Exception:
        return {}


def folder_kind(path):
    for key, label in FOLDER_KIND:
        if key in path:
            return label
    return ""


def guess_date(name, mtime):
    m = DATE.search(name)
    if m:
        y, mo, d = m.groups()
        if "2020" <= y <= "2030" and "01" <= mo <= "12":
            return f"{y}-{mo}-{d}"
    return time.strftime("%Y-%m-%d", time.localtime(mtime))


def scan(rescan=False):
    import source_dirs as S
    cache, out = load_cache(), []
    roots = []
    for attr in ("ERP_DIR", "BAND_DIR", "COUPANG_DIR", "KAKAO_DIR", "RECEIPT_DIR",
                 "DOC_DIR", "ORIGIN_ROOT"):
        p = getattr(S, attr, None)
        if p and os.path.isdir(p):
            roots.append(p)
    seen_root = []
    for r in sorted(set(roots), key=len):        # 상위 폴더가 하위를 포함하면 한 번만
        if not any(r.startswith(x + os.sep) for x in seen_root):
            seen_root.append(r)

    classify = None
    try:
        from inbox_scan import classify as _c
        classify = _c
    except Exception:
        pass

    for root in seen_root:
        for dirpath, _dirs, files in os.walk(root):
            for fn in files:
                ext = os.path.splitext(fn)[1].lower()
                if ext in SKIP_EXT or fn.startswith("~$"):
                    continue
                p = os.path.join(dirpath, fn)
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                key = f"{p}|{st.st_size}|{int(st.st_mtime)}"
                hit = None if rescan else cache.get(key)
                if hit is None:
                    kind = folder_kind(p)
                    # 내용 판별은 ERP 엑셀에만(그쪽만 이름이 무작위라 판별이 필요하다)
                    if classify and ext == ".xlsx" and kind in ("ERP", ""):
                        try:
                            k2 = classify(p)
                            if k2 and k2 != "unknown":
                                kind = f"ERP:{k2}"
                        except Exception:
                            pass
                    uj = (UJ.search(fn) or UJ.search(dirpath))
                    slip = SLIP.search(fn)
                    post = POST.search(fn) if not slip else None
                    po = PO.search(fn)
                    hit = {
                        "name": fn, "path": p, "kind": kind or "기타",
                        "uj": uj.group(0) if uj else "",
                        "slip": slip.group(1) if slip else "",
                        "post": post.group(1) if post else "",
                        "po": po.group(0).upper() if po else "",
                        "date": guess_date(fn, st.st_mtime),
                        "ext": ext.lstrip("."), "size": st.st_size,
                        "mtime": time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime)),
                    }
                    cache[key] = hit
                out.append(hit)
    os.makedirs(os.path.dirname(CACHE), exist_ok=True)
    json.dump(cache, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rescan", action="store_true")
    a = ap.parse_args()
    rows = scan(a.rescan)
    rows.sort(key=lambda r: (r["kind"], r["date"]), reverse=True)
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    json.dump({"count": len(rows), "built": time.strftime("%Y-%m-%d %H:%M"), "rows": rows},
              open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    with open(OUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["종류", "이름", "프로젝트NO", "전표번호", "글번호", "PO", "날짜",
                    "확장자", "크기", "수정", "경로"])
        for r in rows:
            w.writerow([r["kind"], r["name"], r["uj"], r["slip"], r["post"], r["po"],
                        r["date"], r["ext"], r["size"], r["mtime"], r["path"]])
    import collections
    c = collections.Counter(r["kind"] for r in rows)
    print(f"원본 색인 {len(rows)}개 파일 → reports/원본색인.json/csv")
    for k, v in c.most_common(10):
        print(f"  {k}: {v}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
