# -*- coding: utf-8 -*-
"""
source_tidy.py — 원본 저장 폴더를 **누가 봐도 알아보게** 항상 정리한다 (상시 알고리즘)

사용자 지시(2026-08-05): "원본 저장폴더 누가 봐도 깔끔하게 확인할 수 있는 구조로
항상 정리하는 알고리즘 추가해."

무엇이 문제였나
  · ERP 내보내기 이름이 무작위다(`E91RXX1FJ7KAKFP.xlsx`) — 열어 보기 전엔 뭔지 모른다.
  · 폴더는 **받은 날짜**로 나뉘어 1~8월 자료가 전부 오늘 폴더에 쌓인다.
  · 그래서 사람이 "1~6월 자료가 안 보인다"고 느낀다(실제로는 있다).

무엇을 하나 (원본은 절대 지우거나 이동하지 않는다 — 읽기+링크만)
  1. **길잡이 README** 를 폴더마다 자동으로 쓴다. 무슨 자료가 몇 건 있고 어디를 보면
     되는지, 무작위 이름 파일이 실제로 무엇인지(판매조회·계산서현황…) 표로 적는다.
  2. **바로가기 폴더** `_바로가기/` 를 만든다 — 종류별·월별로 사람이 읽는 이름의
     **하드링크**(같은 파일, 공간 안 씀)를 걸어 둔다. 링크가 안 되면 README 만 쓴다.
       _바로가기/ERP_판매조회/2026-08-05_판매조회_1031행.xlsx
       _바로가기/월별/2026-01/… (건별 PDF 는 이미 월별이라 링크만 정리)
  3. 임시·중복 찌꺼기(`~$*`, `*.tmp.xlsx`, `* (1).xlsx`)를 목록으로 보고한다(삭제 안 함).

  python source_tidy.py            # 정리·링크·README 갱신
  python source_tidy.py --report   # 무엇을 할지만 보기(쓰기 없음)
"""
import argparse
import collections
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

KIND_LABEL = {
    "ERP:sales": "판매조회", "ERP:tax": "매출계산서현황", "ERP:taxinv": "계산서진행단계",
    "ERP:stmt": "거래명세서현황", "ERP:ledger": "거래처별원장", "ERP:slips": "회계거래(전표)",
    "ERP:hometax": "홈택스전자계산서", "ERP:po": "쿠팡PO", "ERP": "ERP 기타",
    "밴드": "밴드 원본", "밴드 게시글(보관)": "밴드 게시글", "밴드 문서사진": "밴드 문서사진",
    "카톡": "카톡 내보내기", "쿠팡 PO": "쿠팡 PO", "입금": "입금", "서류": "서류",
    "ERP 거래명세서(건별 PDF)": "거래명세서(건별)",
    "ERP 세금계산서(건별 PDF)": "세금계산서(건별)", "정기점검": "정기점검",
    "미분류": "미분류", "투입 대기": "투입 대기", "기타": "기타",
}
JUNK = (re.compile(r"^~\$"), re.compile(r"\.tmp\.xlsx$", re.I),
        re.compile(r" \(\d+\)\.(xlsx|pdf|txt)$", re.I))


def load_index():
    p = os.path.join(ROOT, "reports", "원본색인.json")
    with open(p, encoding="utf-8") as f:
        return json.load(f).get("rows") or []


def link(src, dst):
    """하드링크(같은 드라이브)로 걸고, 안 되면 건너뛴다. 복사는 하지 않는다 —
    원본이 몇 GB라 두 벌이 되면 안 된다."""
    if os.path.exists(dst):
        return "skip"
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        os.link(src, dst)
        return "ok"
    except Exception:
        return "fail"


def safe(s, n=60):
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]", " ", str(s or "")).strip()
    return re.sub(r"\s+", " ", s)[:n]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--report", action="store_true")
    a = ap.parse_args()

    import source_dirs as S
    rows = load_index()
    origin = S.ORIGIN_ROOT
    short = os.path.join(origin, "_바로가기")

    kinds = collections.Counter(r["kind"] for r in rows)
    junk = [r for r in rows if any(p.search(r["name"]) for p in JUNK)]

    made = fail = 0
    want = set()            # 이번에 있어야 할 바로가기 전체 경로
    if not a.report:
        # 1) 종류별 바로가기 — 무작위 이름을 사람이 읽는 이름으로 건다.
        for r in rows:
            kind = r.get("kind", "기타")
            if kind in ("밴드", "기타") or r.get("ext") not in ("xlsx", "pdf"):
                continue        # 밴드 원본(수천 장)·기타는 링크로 늘리지 않는다
            label = KIND_LABEL.get(kind, kind)
            base = r.get("slip") or r.get("uj") or os.path.splitext(r["name"])[0]
            nm = safe(f"{r.get('date','')}_{label}_{base}") + "." + r["ext"]
            dst = os.path.join(short, safe(label, 30), nm)
            want.add(os.path.normcase(dst))
            if link(r["path"], dst) == "ok":
                made += 1
            else:
                fail += 0
        # 2) 월별 바로가기 — 업무가 **일어난 달**로 묶는다.
        #    건별 PDF(전표번호)와 밴드 게시글 PDF(글번호+날짜)를 함께 넣어,
        #    "그 달에 무슨 일이 있었나"를 한 폴더에서 본다.
        for r in rows:
            if r.get("ext") != "pdf":
                continue
            slip = r.get("slip")
            ym = slip[:7] if slip else (r.get("date") or "")[:7]
            if not (r.get("post") or slip) or len(ym) != 7:
                continue
            dst = os.path.join(short, "월별", ym, r["name"])
            want.add(os.path.normcase(dst))
            if link(r["path"], dst) == "ok":
                made += 1

        # 3) 지난 회차의 찌꺼기를 **거둔다**. 색인이 한 번 오염되면(파생물까지 세면)
        #    그때 만든 링크가 영원히 남아 폴더가 계속 지저분해진다 —
        #    실제로 2026-08-05 에 색인 17,368건(절반이 파생물)으로 만든 링크가 남았다.
        #    ★ 안전장치: 하드링크는 원본과 같은 파일이다. 다른 이름이 **하나도 없는**
        #      것(st_nlink == 1)은 그 자체가 마지막 사본일 수 있으므로 지우지 않고 보고한다.
        removed, kept = 0, []
        for dirpath, _d, files in os.walk(short):
            for fn in files:
                p = os.path.join(dirpath, fn)
                if os.path.normcase(p) in want:
                    continue
                try:
                    if os.stat(p).st_nlink > 1:
                        os.remove(p)
                        removed += 1
                    else:
                        kept.append(p)
                except OSError:
                    kept.append(p)
        for dirpath, _d, files in os.walk(short, topdown=False):   # 빈 폴더 접기
            if not files and not os.listdir(dirpath) and dirpath != short:
                try:
                    os.rmdir(dirpath)
                except OSError:
                    pass

    # 3) 길잡이 README
    lines = [f"# 원본 자료 길잡이 (자동 생성 {time.strftime('%Y-%m-%d %H:%M')})", "",
             "이 폴더는 **원본 그대로** 둡니다. 파일을 옮기거나 지우지 마세요.",
             "찾기 편하도록 `_바로가기/` 안에 **같은 파일의 링크**를 종류별·월별로 걸어 둡니다",
             "(링크라서 용량을 두 배로 쓰지 않습니다).", "",
             "## 지금 갖고 있는 자료", "", "| 종류 | 건수 |", "|---|---:|"]
    for k, v in kinds.most_common():
        lines.append(f"| {KIND_LABEL.get(k, k)} | {v} |")
    lines += ["", "## 어디를 보면 되나", "",
              "- **거래명세서 한 건씩** → `1. ERP 내보내기/거래명세서_건별/2026/월/`",
              "  파일명이 `[전표번호]_거래처_프로젝트NO_금액.pdf` 라 열지 않아도 압니다.",
              "- **세금계산서 한 건씩** → `1. ERP 내보내기/세금계산서_건별/2026/월/`",
              "- **밴드 글 한 건씩** → `4. 밴드 원본/게시글보관/밴드이름/2026/월/`",
              "  `[글번호]_날짜_프로젝트NO_유형` 으로 PDF·본문txt·사진이 한 세트입니다.",
              "- **ERP 화면 엑셀** → `1. ERP 내보내기/2026/월/받은날짜/` (이름이 무작위입니다)",
              "  사람이 읽는 이름은 `_바로가기/` 에서 찾으세요.",
              "- **카톡·PO·입금·서류** → 각 번호 폴더", "",
              "## 무작위 이름 파일이 뭔지 알고 싶다면", "",
              "`ecount/reports/원본색인.csv` 를 엑셀로 열면 파일마다 종류·프로젝트NO·"
              "전표번호·날짜가 한 줄로 정리돼 있습니다. 앱의 **원본 자료** 화면에서는",
              "검색해서 클릭 한 번으로 열 수 있습니다.", ""]
    if junk:
        lines += [f"## 정리 후보 {len(junk)}건 (자동 삭제하지 않습니다)", ""]
        for r in junk[:20]:
            lines.append(f"- {r['name']}")
        lines.append("")
    readme = os.path.join(origin, "README_원본자료_길잡이.md")
    if not a.report:
        try:
            open(readme, "w", encoding="utf-8").write("\n".join(lines))
        except Exception as e:
            print("README 쓰기 실패:", str(e)[:60])

    if a.report:
        print(f"원본 정리: 색인 {len(rows)}건 · 정리후보 {len(junk)}건  (보고 전용)")
        return 0
    msg = (f"원본 정리: 색인 {len(rows)}건 · 바로가기 {made}개 생성 · "
           f"묵은 링크 {removed}개 거둠 · 정리후보 {len(junk)}건 → {short}")
    if kept:
        msg += f"\n  ※ 원본이 하나뿐이라 지우지 않은 링크 {len(kept)}개 — 사람이 확인 필요"
        for p in kept[:5]:
            msg += f"\n     {os.path.basename(p)}"
    print(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
