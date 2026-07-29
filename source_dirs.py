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

# 종류별 보관 폴더 — 사람이 열었을 때 무엇이 무엇인지 바로 보이게 번호를 붙인다
# (사용자 지시 2026-07-28: "구분해서 잘 보이게 깔끔하게 정리해줘")
ERP_DIR = os.path.join(ORIGIN_ROOT, "1. ERP 내보내기")
COUPANG_DIR = os.path.join(ORIGIN_ROOT, "2. 쿠팡 목록")
KAKAO_DIR = os.path.join(ORIGIN_ROOT, "3. 카카오톡 내보내기")
BAND_DIR = os.path.join(ORIGIN_ROOT, "4. 밴드 원본")
# 류지영 매니저가 정기점검 스케줄 원본을 계속 갱신하는 정본 폴더.
# 파일명은 바뀔 수 있으므로 동기화 도구가 이 폴더의 최신 정기점검 xlsx를 고른다.
PM_SCHEDULE_DIR = os.path.join(ORIGIN_ROOT, "5. 정기점검 스케쥴 원본")

# PO 원본 — 새 위치가 정본. 예전 공유 폴더도 계속 훑는다(오종현이 아직 거기 넣을 수 있다).
PO_DIRS = [
    os.path.join(ORIGIN_ROOT, "26년도 PO 모음"),
    r"Z:\16. Share\유현민\오종현\26년도 PO 모음",
]

# 입금(수금) 내역 — 오종현이 관리하는 공유 폴더가 정본이다(사용자 지시 2026-07-28:
# "쿠팡 입금 내역은 여기서 항상 확인해서 정리해줘"). PO 모음과 같은 자리에 있다.
# 이카운트 계정별원장으로는 캠프별 거래처 입금이 잡히지 않아 입금 0건으로 보였는데,
# 이 파일에는 실제로 들어와 있다 — 미수 금액을 말할 때 반드시 이걸 같이 본다.
RECEIPT_DIRS = [
    r"Z:\16. Share\유현민\오종현\26년도 쿠팡 입금내역",
    os.path.join(ORIGIN_ROOT, "5. 입금내역"),
]

# 엑셀 원본을 찾을 곳. ORIGIN_ROOT 자체도 남겨 둔다 — 사람이 하위 폴더를 안 거치고
# 루트에 바로 떨어뜨려도 잡히게 하기 위해서다(그게 제일 흔한 실수다).
EXCEL_DIRS = [ERP_DIR, COUPANG_DIR, ORIGIN_ROOT, os.path.join(ROOT, "inbox")]

# 카톡 내보내기 txt
KAKAO_DIRS = [KAKAO_DIR, ORIGIN_ROOT, os.path.join(ROOT, "kakao", "inbox")]


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


def receipt_dirs():
    return existing(RECEIPT_DIRS)


def pm_schedule_dirs():
    return existing([PM_SCHEDULE_DIR])


# 밴드 문서 사진(거래명세서·현장사진) — 1,459장 130MB. **원본이라 서버에 둔다**
# (사용자 지시 2026-07-28: 용량 큰 원본은 '0. 원본 자료' 로). PC 로컬 inbox 도 계속 본다 —
# 새로 받은 사진을 급히 떨어뜨리는 자리이고, 서버가 끊겨도 그건 읽을 수 있어야 한다.
DOC_PHOTO_DIRS = [
    os.path.join(BAND_DIR, "문서사진"),
    os.path.join(ROOT, "band", "docs_inbox"),
]


def doc_photo_dirs():
    return existing(DOC_PHOTO_DIRS)


if __name__ == "__main__":
    import sys, glob
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print("원본 자료 루트:", ORIGIN_ROOT)
    print("  존재:", os.path.isdir(ORIGIN_ROOT))
    for name, fn in (("PO 원본", po_dirs), ("엑셀 원본", excel_dirs), ("카톡 원본", kakao_dirs),
                     ("정기점검 스케줄 원본", pm_schedule_dirs)):
        print(f"\n{name}")
        for d in fn():
            n = sum(len(f) for _b, _d, f in os.walk(d))
            print(f"  {n:>4}개  {d}")
        if not fn():
            print("  (없음)")
