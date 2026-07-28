# -*- coding: utf-8 -*-
"""
collect_sources.py — 흩어져 있는 **원본 자료를 한 폴더로 모은다**
================================================================================
사용자 지시(2026-07-28): "0. 원본 자료 — 여기에 자료 모두 복사해서 붙여넣어주고
데이터 정리해줘. 구분해서 잘 보이게 깔끔하게 정리해줘."

지금까지 원본이 PC 여기저기(inbox·kakao/inbox·band/cache)에 흩어져 있었다.
PC가 꺼지거나 사람이 바뀌면 **무엇이 원본인지 아무도 모른다.** 그래서
관리대장 옆 '0. 원본 자료' 폴더에 종류별로 모아 둔다.

  0. 원본 자료/
      0. 수집안내.txt        ← 무엇이 언제 어디서 왔는지 (이 도구가 갱신)
      1. ERP 내보내기/       ← 이카운트에서 내려받은 xlsx
      2. 쿠팡 목록/          ← 쿠팡이 준 PO 목록 등
      3. 카카오톡 내보내기/   ← 대화방 txt
      4. 밴드 원본/          ← 밴드 API 원문 JSON
      26년도 PO 모음/        ← 쿠팡 PO 통지문·견적서 (오종현 수집, 이미 있음)

원칙
  · **복사만 한다. 원본을 지우거나 옮기지 않는다.** PC 쪽은 그대로 둔다.
  · 크기·수정시각이 같으면 건너뛴다(Z: 네트워크 드라이브가 느리다).
  · 덮어쓰기 전에 내용이 다르면 이름 뒤에 날짜를 붙여 **둘 다 남긴다.**

사용
  python collect_sources.py            # 무엇을 복사할지만 보여 준다 (안전)
  python collect_sources.py --apply    # 실제 복사
"""
import os
import sys
import glob
import shutil
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from source_dirs import ORIGIN_ROOT, ERP_DIR, COUPANG_DIR, KAKAO_DIR, BAND_DIR, PO_DIRS  # noqa: E402

GUIDE = os.path.join(ORIGIN_ROOT, "0. 수집안내.txt")


def _same(a, b):
    """이미 같은 파일인가 — 크기와 수정시각(초)으로 판단. 느린 드라이브에서 해시는 과하다."""
    try:
        sa, sb = os.stat(a), os.stat(b)
    except OSError:
        return False
    return sa.st_size == sb.st_size and int(sa.st_mtime) == int(sb.st_mtime)


# 이카운트에서 'Excel' 을 누르면 파일이 Downloads 로 떨어지고 이름이 무작위다
# (`8W1JR7MGB50PHOP.xlsx`). 사람이 옮기지 않아도 되게 여기서 직접 집어 온다.
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
DOWNLOAD_DAYS = 14      # 오래된 건 이미 반영됐거나 무관하다
# ★ Downloads 는 개인 폴더다. **아무 파일이나 퍼 오지 않는다** —
#   내용을 열어 아는 종류로 판별된 것만 가져온다.
KNOWN = ("ledger", "po", "sales", "tax", "stmt", "slips", "taxinv", "hometax")


def plan():
    """[(원본, 대상, 분류)] — 어디서 어디로 갈지. 판단은 여기서만 한다."""
    jobs = []

    # ERP 내보내기 / 쿠팡 목록 — 파일명이 무작위일 수 있어 **내용으로** 가른다
    try:
        from inbox_scan import classify, LABEL
    except Exception:
        classify, LABEL = None, {}

    def kind_of(path):
        try:
            return classify(path) if classify else "unknown"
        except Exception:
            return "unknown"

    for src in sorted(glob.glob(os.path.join(BASE, "inbox", "*.xls*"))):
        if os.path.basename(src).startswith("~$"):
            continue
        k = kind_of(src)
        jobs.append((src, COUPANG_DIR if k == "po" else ERP_DIR, LABEL.get(k, "엑셀")))

    # Downloads 에 떨어진 이카운트 내보내기 — 아는 종류만, 최근 것만
    cutoff = __import__("time").time() - DOWNLOAD_DAYS * 86400
    for src in sorted(glob.glob(os.path.join(DOWNLOADS, "*.xls*"))):
        if os.path.basename(src).startswith("~$"):
            continue
        try:
            if os.path.getmtime(src) < cutoff:
                continue
        except OSError:
            continue
        k = kind_of(src)
        if k not in KNOWN:
            continue
        jobs.append((src, COUPANG_DIR if k == "po" else ERP_DIR,
                     LABEL.get(k, "엑셀") + " (Downloads)"))

    # 카카오톡 내보내기
    for src in sorted(glob.glob(os.path.join(BASE, "kakao", "inbox", "*.txt"))):
        jobs.append((src, KAKAO_DIR, "카톡 대화 내보내기"))

    # 밴드 원문 — 우리가 API로 받아 온 원본 그대로. 가공본(캐시)도 같이 둔다.
    for src in sorted(glob.glob(os.path.join(BASE, "band", "cache", "*.json"))):
        jobs.append((src, BAND_DIR, "밴드 API 원문"))

    return jobs


def copy_one(src, dst_dir, apply=False):
    """반환: ('복사'|'동일'|'이름바꿈'|'실패', 대상경로)"""
    dst = os.path.join(dst_dir, os.path.basename(src))
    if os.path.exists(dst):
        if _same(src, dst):
            return "동일", dst
        # 내용이 다르면 덮어쓰지 않는다 — 어느 쪽이 맞는지 사람이 봐야 한다
        stamp = datetime.fromtimestamp(os.path.getmtime(src)).strftime("%y%m%d")
        root, ext = os.path.splitext(dst)
        dst = f"{root}_{stamp}{ext}"
        if os.path.exists(dst) and _same(src, dst):
            return "동일", dst
        state = "이름바꿈"
    else:
        state = "복사"
    if apply:
        try:
            os.makedirs(dst_dir, exist_ok=True)
            shutil.copy2(src, dst)          # copy2 = 수정시각까지 보존 (다음 실행에서 '동일' 판정)
        except OSError as e:
            return f"실패({e.strerror})", dst
    return state, dst


def count_rows(path):
    """엑셀은 행 수를, 텍스트는 줄 수를 센다 — '넣었는데 비어 있는' 파일을 잡기 위해."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext.startswith(".xls"):
            import openpyxl
            w = openpyxl.load_workbook(path, read_only=True, data_only=True)
            n = sum(1 for sn in w.sheetnames for r in w[sn].iter_rows(values_only=True)
                    if sum(1 for x in r if x not in (None, "")) >= 3)
            w.close()
            return n
        if ext == ".txt":
            with open(path, encoding="utf-8", errors="replace") as f:
                return sum(1 for _ in f)
    except Exception:
        pass
    return -1


def write_guide(rows):
    """폴더를 처음 여는 사람이 읽을 안내문. 파일이 아니라 **뜻**을 적는다."""
    L = [
        "이 폴더는 '쿠팡 통합업무 자동화'가 쓰는 원본 자료 보관소입니다.",
        f"마지막 정리: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "■ 폴더 구분",
        "  1. ERP 내보내기    이카운트에서 내려받은 엑셀 (거래명세서·계산서·회계거래)",
        "  2. 쿠팡 목록       쿠팡이 준 PO 목록 등",
        "  3. 카카오톡 내보내기  대화방 '대화 내보내기' txt",
        "  4. 밴드 원본       밴드에서 받아 온 게시글 원문(JSON) — 사람이 열 필요는 없습니다",
        "  26년도 PO 모음     쿠팡 PO 통지문·견적서 (PO번호별 하위 폴더)",
        "",
        "■ 자료를 새로 넣을 때",
        "  해당 폴더에 그냥 넣어 주시면 됩니다. 파일 이름은 아무래도 괜찮습니다 —",
        "  프로그램이 파일을 열어 내용으로 종류를 알아냅니다.",
        "  ★ 이카운트에서 'Excel' 로 내려받으면 Downloads 에 떨어지는데, 옮기지 않으셔도",
        "    됩니다. 최근 2주 안에 받은 것 중 아는 종류만 자동으로 가져옵니다.",
        "",
        "■ 이카운트에서 내보낼 화면 (매출 대조에 쓰는 것)",
        "  1) 재고 I > 영업관리 > 판매일괄회계반영 > 매출(세금)계산서조회(재고)",
        "  2) 재고 I > 영업관리 > 판매일괄회계반영 > 매출(세금)계산서현황(재고)  ← 품목·내역 단위",
        "  3) 회계 I > 전자(세금)계산서 > 홈택스자료조회 > 전자(세금)계산서",
        "     ('미반영'이 아니라 '전체' 를 고르고 기간을 올해로 잡아 주세요)",
        "  4) 회계 I > 전자(세금)계산서 > 이카운트 vs 홈택스 자료비교  ← 차이가 바로 보이는 화면",
        "",
        "■ 주의",
        "  · 이카운트에서 내보낼 때는 조회 조건을 넣어 **화면에 행이 보이는 상태**에서",
        "    내보내세요. 조건 없이 내보내면 회사명 한 줄만 든 빈 파일이 됩니다.",
        "  · 파일을 지우지 마세요. 옛 자료도 대조 근거로 계속 씁니다.",
        "",
        "■ 현재 보관 현황",
    ]
    for d, items in rows:
        L.append(f"  [{os.path.basename(d)}]  {len(items)}개")
        for name, n in items:
            mark = "  ★ 내용 없음" if n == 0 else (f"  {n:,}행" if n > 0 else "")
            L.append(f"      {name}{mark}")
    try:
        os.makedirs(ORIGIN_ROOT, exist_ok=True)
        with open(GUIDE, "w", encoding="utf-8") as f:
            f.write("\n".join(L) + "\n")
    except OSError as e:
        print(f"  안내문 저장 실패: {e}")


def main():
    apply = "--apply" in sys.argv
    if not os.path.isdir(ORIGIN_ROOT):
        print(f"원본 폴더에 접근할 수 없습니다(네트워크 드라이브 확인): {ORIGIN_ROOT}")
        return 2

    jobs = plan()
    print(f"{'복사 실행' if apply else '미리보기(복사 안 함)'} — 대상 {len(jobs)}개\n")
    tally = {}
    for src, dst_dir, kind in jobs:
        state, dst = copy_one(src, dst_dir, apply)
        tally[state] = tally.get(state, 0) + 1
        print(f"  [{state:6s}] {os.path.basename(dst_dir)}/{os.path.basename(dst)}   ({kind})")

    print("\n집계: " + ", ".join(f"{k} {v}개" for k, v in sorted(tally.items())))
    if not apply:
        print("\n실제로 복사하려면:  python collect_sources.py --apply")
        return 0

    # 정리 결과를 폴더별로 세어 안내문에 남긴다
    rows, empty = [], []
    for d in (ERP_DIR, COUPANG_DIR, KAKAO_DIR, BAND_DIR):
        if not os.path.isdir(d):
            continue
        items = []
        for name in sorted(os.listdir(d)):
            p = os.path.join(d, name)
            if os.path.isfile(p):
                n = count_rows(p)
                items.append((name, n))
                if n == 0:
                    empty.append(f"{os.path.basename(d)}/{name}")
        rows.append((d, items))
    for d in PO_DIRS:
        if os.path.isdir(d) and d.startswith(ORIGIN_ROOT):
            n = sum(len(f) for _b, _dd, f in os.walk(d))
            rows.append((d, [(f"(하위 폴더 포함 {n}개 파일)", -1)]))
    write_guide(rows)
    print(f"안내문 갱신: {os.path.basename(GUIDE)}")
    if empty:
        print(f"\n★ 내용이 비어 있는 파일 {len(empty)}개 — 다시 내보내야 합니다")
        for x in empty:
            print(f"    {x}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
