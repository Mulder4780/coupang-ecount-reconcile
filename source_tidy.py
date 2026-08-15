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
INCREMENTAL_RETURN_CODE = 75


def load_index():
    p = os.path.join(ROOT, "reports", "원본색인.json")
    with open(p, encoding="utf-8") as f:
        return json.load(f).get("rows") or []


def link(src, dst, existing=None, known_dirs=None):
    """하드링크(같은 드라이브)로 걸고, 안 되면 건너뛴다. 복사는 하지 않는다 —
    원본이 몇 GB라 두 벌이 되면 안 된다.

    ``existing`` 과 ``known_dirs`` 는 `_바로가기`를 한 번 훑어 만든 집합이다. Z:(SMB)에
    파일마다 ``exists``·``makedirs``를 다시 물으면 2만여 링크를 매 회차 처음부터
    왕복하느라 제한시간 안에 새 링크까지 도달하지 못한다.
    """
    key = os.path.normcase(dst)
    if existing is not None and key in existing:
        return "skip"
    if existing is None and os.path.exists(dst):
        return "skip"
    parent = os.path.dirname(dst)
    parent_key = os.path.normcase(parent)
    try:
        if known_dirs is None or parent_key not in known_dirs:
            os.makedirs(parent, exist_ok=True)
            if known_dirs is not None:
                known_dirs.add(parent_key)
        os.link(src, dst)
        if existing is not None:
            existing.add(key)
        return "ok"
    except FileExistsError:       # 다른 회차가 먼저 만들었어도 같은 완료다
        if existing is not None:
            existing.add(key)
        return "skip"
    except Exception:
        return "fail"


def safe(s, n=60):
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]", " ", str(s or "")).strip()
    return re.sub(r"\s+", " ", s)[:n]


# ── 시간 예산 (2026-08-12 · 분담판 [38]) ─────────────────────────────────────
# 이 회차는 작업 스케줄러 제한(PT3H)에 걸려 **매일 나무째 끊기고 있었다**
# (0xC000013A). 끊기면 스케줄러는 그 사실을 결과 코드로만 남기고 리포트는 한 줄도
# 안 써진다 — 그래서 반년을 돌면서도 아무 화면에 안 떴다([228] 이 잡아낸 것이 이것이다).
# 링크 하나하나가 Z:(SMB) 왕복이라 원본이 늘면 시간도 같이 는다. 그러니 "빠르게
# 만든다"로는 끝이 없고, **끝나는 시각을 정하는 것**이 답이다.
# ★ 워치독 한 회차의 전체 예산은 10분이다. 예전 기본 100분은 daily_run 30분 제한과
#   워치독 10분 제한보다 길어, 자체 이월 분기에 닿기 전에 매번 rc124로 잘렸다([287]).
#   7분 뒤 75로 돌아오면 워치독이 실패로 세지 않고 다음 회차에 이어 간다.
BUDGET_MIN = int(os.environ.get("COUPANG_TIDY_BUDGET_MIN", "7"))
_T0 = time.time()


def out_of_time():
    return (time.time() - _T0) / 60 >= BUDGET_MIN


def link_plan(rows, short):
    """이번 색인에서 있어야 할 링크 전체를 **로컬 메모리에서 먼저** 계산한다.

    계획이 완성돼야만 찌꺼기를 안전하게 가릴 수 있다. 링크를 만들면서 계획도 만들면
    시간예산에 걸린 순간 ``want``가 반쪽이라 멀쩡한 링크를 지울 수 있다.
    """
    plan, want = [], set()

    def add(src, dst):
        key = os.path.normcase(dst)
        want.add(key)
        plan.append((src, dst))

    # 1) 종류별 바로가기 — 무작위 이름을 사람이 읽는 이름으로 건다.
    for r in rows:
        kind = r.get("kind", "기타")
        if kind in ("밴드", "기타") or r.get("ext") not in ("xlsx", "pdf"):
            continue        # 밴드 원본(수천 장)·기타는 링크로 늘리지 않는다
        label = KIND_LABEL.get(kind, kind)
        base = r.get("slip") or r.get("uj") or os.path.splitext(r["name"])[0]
        nm = safe(f"{r.get('date','')}_{label}_{base}") + "." + r["ext"]
        add(r["path"], os.path.join(short, safe(label, 30), nm))

    # 2) 월별 바로가기 — 업무가 **일어난 달**로 묶는다.
    for r in rows:
        if r.get("ext") != "pdf":
            continue
        slip = r.get("slip")
        ym = slip[:7] if slip else (r.get("date") or "")[:7]
        if not (r.get("post") or slip) or len(ym) != 7:
            continue
        add(r["path"], os.path.join(short, "월별", ym, r["name"]))
    return plan, want


def link_snapshot(short):
    """기존 바로가기·폴더·크기를 SMB 한 번 순회로 함께 받는다."""
    from source_index import walk_stat
    existing, known_dirs, entries = set(), set(), []
    for dirpath, fn, st in walk_stat(short, skip_dirs=()):
        existing.add(os.path.normcase(os.path.join(dirpath, fn)))
        known_dirs.add(os.path.normcase(dirpath))
        entries.append((dirpath, fn, st))
    return existing, known_dirs, entries


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
    if not a.report:
        plan, want = link_plan(rows, short)
        print(f"[원본정리] 색인 {len(rows)}건 → 바로가기 계획 {len(want)}개 · "
              f"시간예산 {BUDGET_MIN}분", flush=True)
        existing, known_dirs, old_entries = link_snapshot(short)
        print(f"[원본정리] 기존 바로가기 {len(existing)}개 확인 · 새 링크 단계 시작",
              flush=True)

        # 기존 링크는 메모리 집합으로 건너뛴다. 새로 필요한 링크만 Z:에 쓴다.
        for pos, (src, dst) in enumerate(plan):
            if out_of_time():
                left = sum(1 for _src, p in plan[pos:]
                           if os.path.normcase(p) not in existing)
                print(f"[원본정리] 시간 예산 {BUDGET_MIN}분에 걸려 새 링크 {made}개까지 "
                      f"만들고 멈춥니다 — 남은 {left}개는 다음 회차가 이어서 합니다.\n"
                      "           찌꺼기 거두기는 계획을 다 반영한 뒤에만 합니다.",
                      flush=True)
                return INCREMENTAL_RETURN_CODE
            result = link(src, dst, existing, known_dirs)
            if result == "ok":
                made += 1
                if made % 250 == 0:
                    print(f"[원본정리] 새 링크 {made}개 생성 · 계획 {pos + 1}/{len(plan)}",
                          flush=True)
            elif result == "fail":
                fail += 1

        # 3) 지난 회차의 찌꺼기를 **거둔다**. 색인이 한 번 오염되면(파생물까지 세면)
        #    그때 만든 링크가 영원히 남아 폴더가 계속 지저분해진다 —
        #    실제로 2026-08-05 에 색인 17,368건(절반이 파생물)으로 만든 링크가 남았다.
        #    ★ 안전장치 — **st_nlink 로 판단하지 않는다.** Z: 는 네트워크 공유(SMB)라
        #      하드링크라도 링크 수가 늘 1 로 온다(실측 2026-08-06: 바로가기 6,008개
        #      **전부** nlink==1). 그것을 믿으면 아무것도 못 거두고(첫 실행 0개),
        #      무시하면 유일본을 지운다. 그래서 **원본이 색인에 남아 있는가**로 본다:
        #      크기가 같고, 원본 이름(확장자 뗀 것)이 바로가기 이름 안에 들어 있으면
        #      같은 파일로 본다(종류별 링크는 `날짜_라벨_원본stem.ext`, 월별은 원본 이름 그대로).
        by_size = collections.defaultdict(list)
        for r in rows:
            by_size[r.get("size") or 0].append(r)

        def has_original(fn, size):
            """이 바로가기의 원본이 색인에 아직 있나.

            이름을 세 가지로 견준다 — 바로가기 이름은 만든 방식이 둘이기 때문이다:
              · 월별 링크  : 원본 이름 그대로
              · 종류별 링크: `날짜_라벨_기준값.ext` 이고 기준값은 **전표·프로젝트·글번호**다.
            그래서 원본 파일 이름만 견주면 종류별 링크가 전부 '못 찾음'이 된다
            (실측: 6,008개 중 2,740개가 그랬다 — 전부 건별 PDF 였다).
            """
            low = fn.lower()
            for r in by_size.get(size, ()):
                nm = r["name"].lower()
                if nm == low or os.path.splitext(nm)[0] in low:
                    return True
                for key in (r.get("slip"), r.get("uj"), r.get("post")):
                    if key and str(key).lower() in low:
                        return True
            return False

        removed, kept = 0, []
        # ★ 크기는 **목록을 받을 때 이미 딸려 온다** — 파일마다 다시 묻지 않는다
        #   (2026-08-11, 검증 [198]). 예전엔 walk 로 이름만 받고 `os.stat(p).st_size` 를
        #   파일마다 불렀는데, Z:(SMB)에서는 그것이 **파일당 왕복 한 번**이라 실측
        #   135~155 ms 다(딸려 온 값은 0.04 ms). 이 단계가 13시간 30분 매달렸던
        #   그 단계다([175]는 죽이는 방법을 고쳤고, 여기는 애초에 왜 오래 걸렸나다).
        #   거를 폴더는 **없다**(빈 set) — 여기는 `_바로가기` 안을 훑는 자리라
        #   색인의 SKIP_DIRS 를 물려받으면 정리할 것을 통째로 안 보게 된다.
        print(f"[원본정리] 새 링크 단계 완료({made}개 생성) · 찌꺼기 확인 시작",
              flush=True)
        for pos, (dirpath, fn, st) in enumerate(old_entries):
            if out_of_time():
                print(f"[원본정리] 찌꺼기 {pos}/{len(old_entries)}개 확인 뒤 시간 예산 종료"
                      " — 다음 회차가 이어서 합니다.", flush=True)
                return INCREMENTAL_RETURN_CODE
            p = os.path.join(dirpath, fn)
            if os.path.normcase(p) in want:
                continue
            try:
                if has_original(fn, st.st_size):
                    os.remove(p)
                    removed += 1
                else:
                    kept.append(p)
            except OSError:
                kept.append(p)
        if out_of_time():
            print("[원본정리] 링크 정리는 끝났고 빈 폴더 접기는 다음 회차가 이어서 합니다.",
                  flush=True)
            return INCREMENTAL_RETURN_CODE
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
           f"묵은 링크 {removed}개 거둠 · 링크실패 {fail}개 · "
           f"정리후보 {len(junk)}건 → {short}")
    if kept:
        msg += (f"\n  ※ 원본을 색인에서 못 찾아 **지우지 않은** 링크 {len(kept)}개 — "
                f"유일본일 수 있으니 사람이 확인")
        for p in kept[:5]:
            msg += f"\n     {os.path.basename(p)}"
    print(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
