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


def _acquire_lock():
    """daily_run 과 같은 방식의 프로세스 간 단발 잠금. 겹치면 None."""
    try:
        sys.path.insert(0, ROOT)
        from daily_run import acquire_run_lock
        return acquire_run_lock(os.path.join(ROOT, "reports", ".source_index.lock"))
    except Exception:
        return "no-lock"          # 잠금을 못 쓰면 예전처럼 그냥 돈다(막지는 않는다)


def _release_lock(token):
    if token == "no-lock":
        return
    try:
        from daily_run import release_run_lock
        release_run_lock(token, os.path.join(ROOT, "reports", ".source_index.lock"))
    except Exception:
        pass
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
# ★ 우리가 만든 파생물은 색인에 넣지 않는다 (2026-08-05).
#   `_바로가기` 는 원본의 하드링크라 **같은 파일이 두 번** 잡히고,
#   `_보관` 은 날마다 쌓는 백업이라 어제·그제 사본이 검색 결과를 덮는다.
#   실측: 색인 17,368건 중 8,400여 건이 이 둘이었고 전부 '기타'로 뭉쳐 있었다 —
#   앱의 원본 자료 화면에서 같은 자료가 여러 번 나오는 원인이다.
SKIP_DIRS = {"_바로가기", "_보관", "__pycache__", ".git"}

# ★ 민감 자료는 색인에 **아예 담지 않는다** (2026-08-07 지시).
#   사용자 지시: "통화_MD는 원본 자료에서 안 보이게 처리하고 DB만 보관해(민감한 내용이
#   포함되어있음)." 이 색인은 앱 '원본 자료' 목록이면서 동시에 `/api/source-file` 이
#   내려보내도 되는 경로의 **화이트리스트**다. 그래서 여기서 한 번 빼면 목록에도 안 뜨고
#   내려받을 수도 없다 — 두 겹이 한 자리에서 닫힌다.
#   통화 메모의 정본은 DB(`ledger_db.call_note`)이고, 옮기기는 `call_notes.py --migrate`.
#   검증 [129].
PRIVATE_NAME = re.compile(r"^통화[_\-]")


def is_private(path, name=""):
    """색인에 넣으면 안 되는 민감 파일인가 — 이름 규칙 + 통화·회의 폴더 전체."""
    name = name or os.path.basename(path)
    if PRIVATE_NAME.match(name):
        return True
    try:
        import source_dirs
        d = getattr(source_dirs, "CALL_NOTE_DIR", "")
        if d:
            base = os.path.normcase(os.path.abspath(d)).rstrip("\\/")
            here = os.path.normcase(os.path.abspath(path))
            if here == base or here.startswith(base + os.sep):
                return True
    except Exception:
        pass
    return False

# 폴더 이름으로 1차 분류 — 내용 판별보다 싸고, 사람이 이해하는 갈래와 같다.
# ★ 순서가 곧 우선순위다. **좁은 것부터** 적는다 — 경로에는 상위 폴더 이름도 같이
#   들어 있기 때문이다. 예전에는 `4. 밴드 원본` 이 `게시글보관` 보다 앞이라,
#   `4. 밴드 원본/게시글보관/...` 이 전부 뭉뚱그려 '밴드' 가 됐고 게시글·문서사진
#   갈래가 **한 건도 생기지 않았다**(2026-08-05 실측: 밴드 4,990 / 게시글 0).
# 폴더 이름은 **source_dirs 가 정한다.** 여기 문자열을 따로 적어 두면 반드시 어긋난다 —
# 실제로 `2. 쿠팡 PO`·`5. 입금`·`6. 서류` 는 **없는 폴더**였고(진짜 이름은 `6. PO 원본`·
# `7. 입금내역`), 그래서 쿠팡 PO 576건이 전부 '기타' 로 빠져 앱에서 안 보였다(2026-08-05).
_SPECIFIC = [                      # 상위 폴더 안에 있는 좁은 갈래 — 반드시 먼저 본다
    ("거래명세서_건별", "ERP 거래명세서(건별 PDF)"),
    ("세금계산서_건별", "ERP 세금계산서(건별 PDF)"),
    ("게시글보관", "밴드 게시글(보관)"),
    ("문서사진", "밴드 문서사진"),
]
_BY_DIR = [("ERP_DIR", "ERP"), ("BAND_DIR", "밴드"), ("PO_DIR", "쿠팡 PO"),
           ("COUPANG_DIR", "쿠팡 목록"), ("KAKAO_DIR", "카톡"), ("RECEIPT_DIR", "입금"),
           ("DOC_DIR", "서류"), ("PM_SCHEDULE_DIR", "정기점검"),
           ("WORK_LOG_DIR", "업무일지"), ("LEGACY_WORK_LOG_DIR", "업무일지"),
           ("CALL_NOTE_DIR", "통화·회의"), ("NEW_PROJECT_FLOW_DIR", "업무 흐름도"),
           ("MISC_DIR", "미분류"), ("UPLOAD_DIR", "투입 대기")]


def _folder_rules():
    rules = list(_SPECIFIC)
    try:
        import source_dirs as S
    except Exception:
        return rules
    for attr, label in _BY_DIR:
        p = getattr(S, attr, None)
        if p:
            name = os.path.basename(str(p).rstrip("\\/"))
            if name and (name, label) not in rules:
                rules.append((name, label))
    return rules


FOLDER_KIND = _folder_rules()


# ★ 분류 규칙이 바뀌면 캐시를 통째로 버린다 (2026-08-05 실사고).
#   캐시 열쇠는 `경로|크기|수정시각` 이라 **파일이 안 바뀌면 옛 갈래를 그대로 돌려준다**.
#   그래서 FOLDER_KIND 순서를 고치고 다시 돌렸는데 결과가 하나도 안 바뀌었다
#   (밴드 5,209 그대로, 게시글 갈래 0). 규칙 지문을 같이 적어 두고 다르면 다시 판별한다 —
#   안 그러면 앞으로 규칙을 고칠 때마다 조용히 헛일이 된다.
RULES_KEY = "__rules__"


def rules_version():
    """★ 폴더 규칙만 해싱하면 **반쪽이다** (2026-08-08 실사고 — 같은 사고 두 번째).

    갈래를 정하는 것은 둘이다: 폴더 규칙(FOLDER_KIND)과 **내용 판별**(inbox_scan).
    그런데 지문에 앞엣것만 들어 있었다. 그래서 EBG006M(E010727)이 taxinv 로 새는 것을
    inbox_scan 에서 고치고 색인을 다시 돌렸는데 **결과가 하나도 안 바뀌었다** —
    파일이 안 바뀌었으니 캐시가 옛 갈래를 그대로 돌려줬다(ERP:sales 10건 그대로).
    2026-08-05 주석이 "앞으로 규칙을 고칠 때마다 조용히 헛일이 된다"고 경고한 그대로다.

    그래서 내용 판별의 **실제 코드**를 같이 지문에 넣는다. 주석만 고쳐도 다시 도는 것이
    아깝지만, 고친 규칙이 안 먹는 것보다 낫다 — 안 먹으면 아무도 모른다.
    """
    import hashlib
    parts = [repr(FOLDER_KIND)]
    try:
        import inspect, inbox_scan
        parts.append(repr(getattr(inbox_scan, "ERP_FILE_PREFIX", {})))
        parts.append(inspect.getsource(inbox_scan.classify_rows))
        parts.append(inspect.getsource(inbox_scan.classify))
    except Exception:
        parts.append("<inbox_scan 못 읽음>")   # 못 읽으면 예전대로 — 색인을 막지는 않는다
    return hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:10]


def load_cache():
    try:
        d = json.load(open(CACHE, encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(d, dict) or d.pop(RULES_KEY, None) != rules_version():
        return {}
    return d


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
            # 걸러낼 폴더는 **내려가기 전에** 지운다(os.walk 는 이 목록을 보고 내려간다).
            _dirs[:] = [d for d in _dirs if d not in SKIP_DIRS]
            for fn in files:
                ext = os.path.splitext(fn)[1].lower()
                if ext in SKIP_EXT or fn.startswith("~$"):
                    continue
                p = os.path.join(dirpath, fn)
                if is_private(p, fn):     # 통화 메모 등 — 색인 자체에 남기지 않는다
                    continue
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
    cache[RULES_KEY] = rules_version()      # 어떤 규칙으로 판별한 캐시인지 같이 적는다
    # ★ 통째 덮어쓰기(`open(CACHE,"w")`)는 **여는 순간 원본을 0바이트로 만든다.**
    #   이 파일은 수 MB 고 채우는 데 시간이 걸리는데, 그 사이에 세션이 끊기거나
    #   **다른 세션이 같은 색인을 돌면** 두 글이 섞여 읽을 수 없는 파일이 남는다.
    #   실측 2026-08-08: 색인 한 번이 Z: 에서 6시간 넘게 걸렸고 그동안 09:50 회차가
    #   같은 폴더를 훑고 있었다 — 겹칠 창이 여섯 시간이나 열려 있었다는 뜻이다.
    #   `load_cache` 가 깨진 파일을 {} 로 삼켜 주긴 하지만, 그 대가가 **또 6시간**이다.
    try:
        from ledger_writer import atomic_json_dump
        atomic_json_dump(cache, CACHE)      # (값, 경로) 순서다 — 뒤집으면 조용히 실패한다
    except Exception:
        # 헬퍼를 못 불러도 색인은 끝내야 한다 — 예전 방식으로라도 남긴다.
        json.dump(cache, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rescan", action="store_true")
    a = ap.parse_args()
    # ★ 한 번에 하나만 (2026-08-05 실사고). 세션이 둘 떠 있으면 양쪽 daily_run 이
    #   같은 시각에 Z: 11,000여 개 파일을 훑어, 느려지는 정도가 아니라 **색인 갱신·
    #   원본 정리·자료현황·폰 사본·보관 5단계가 통째로 실패**했다. 겹치면 이번 회차는
    #   건너뛴다 — 워치독이 30분마다 다시 만들므로 한 번 거르는 편이 낫다.
    token = _acquire_lock()
    if not token:
        print("다른 source_index 가 실행 중 — 이번 회차는 건너뜁니다(색인은 그대로 유효).")
        return 0
    try:
        return _build(a.rescan)
    finally:
        _release_lock(token)


def _build(rescan):
    rows = scan(rescan)
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
