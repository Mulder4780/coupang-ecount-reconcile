# -*- coding: utf-8 -*-
"""
source_dirs.py — **원본 자료가 어디 있는지**를 한 곳에서만 정한다
================================================================================
사용자 지시(2026-07-28): "모든 원본 자료는 이 폴더에 취합해서 관리하게 정리해줘."

지금까지 원본 경로가 도구마다 흩어져 있었다(po_pdf 안에 하드코딩, inbox_scan 은 로컬만,
카톡은 kakao/inbox). 그래서 자료 위치가 바뀔 때마다 여기저기 고쳐야 했고, 실제로
PO 경로가 한 번 되돌아가 자료를 못 읽은 적이 있다(2026-07-28).

→ **경로는 이 파일에서만 정한다.** 다른 도구는 여기서 가져다 쓴다.

구조
  0. 원본 자료/                     ← 모든 원본이 모이는 곳(관리대장 폴더 아래)
      26년도 PO 모음/               ← 쿠팡 PO 통지문·견적서 (PO별 하위 폴더)
      (ERP 내보내기·카톡 내보내기 등도 여기 넣으면 자동으로 잡힌다)

로컬 inbox 도 계속 본다 — 사람이 급할 때 PC에 바로 떨어뜨리는 경우가 있어서다.
"""
import os

ROOT = os.path.dirname(os.path.abspath(__file__))

# 관리대장이 있는 폴더 아래 '0. 원본 자료'
LEDGER_DIR = (r"Z:\2. Cost\★★★쿠팡 업무 폴더★★★"
              r"\♣ 1000. 쿠팡 통합업무관리 전산화 프로젝트"
              r"\00. 대시보드 (프로젝트 일정, 담당자, 진행현황, 문제사항)"
              r"\00. 쿠팡 통합업무 일일보고 관리대장")
ORIGIN_ROOT = os.path.join(LEDGER_DIR, "0. 원본 자료")

# PO 원본 — 새 위치가 정본. 예전 공유 폴더도 계속 훑는다(오종현이 아직 거기 넣을 수 있다).
PO_DIRS = [
    os.path.join(ORIGIN_ROOT, "26년도 PO 모음"),
    r"Z:\16. Share\유현민\오종현\26년도 PO 모음",
]

# 엑셀 원본(ERP 내보내기·쿠팡 PO 목록 등)을 찾을 곳 — 원본 폴더 + 로컬 inbox
EXCEL_DIRS = [ORIGIN_ROOT, os.path.join(ROOT, "inbox")]

# 카톡 내보내기 txt
KAKAO_DIRS = [ORIGIN_ROOT, os.path.join(ROOT, "kakao", "inbox")]


def existing(paths):
    """실제로 있는 폴더만. 네트워크 드라이브가 끊겨도 죽지 않게."""
    out = []
    for p in paths:
        try:
            if os.path.isdir(p):
                out.append(p)
        except OSError:
            pass
    return out


def po_dirs():
    return existing(PO_DIRS)


def excel_dirs():
    return existing(EXCEL_DIRS)


def kakao_dirs():
    return existing(KAKAO_DIRS)


if __name__ == "__main__":
    import sys, glob
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print("원본 자료 루트:", ORIGIN_ROOT)
    print("  존재:", os.path.isdir(ORIGIN_ROOT))
    for name, fn in (("PO 원본", po_dirs), ("엑셀 원본", excel_dirs), ("카톡 원본", kakao_dirs)):
        print(f"\n{name}")
        for d in fn():
            n = sum(len(f) for _b, _d, f in os.walk(d))
            print(f"  {n:>4}개  {d}")
        if not fn():
            print("  (없음)")
