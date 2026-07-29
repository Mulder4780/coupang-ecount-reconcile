# -*- coding: utf-8 -*-
"""확인필요 항목의 내부 확인 담당자 확정 규칙.

사용자 확정(2026-07-29):
- 류지영: 캠프·일정, 카톡/밴드, 현장자료, 완료일 확인
- 변재선(회계): 거래명세서, 세금계산서, 입금, 금액 불일치 확인
- 오종현: PO 원본·견적서, 구매·입금 원천자료 취합 및 누락 확인
- 유현민: 기존 PO·ERP·원장 미등록·시스템 연결 확인

이 규칙은 현장 작업 담당자를 바꾸는 것이 아니라, ``23_확인필요현황``과 앱에서
"누가 확인하고 정리할 항목인가"를 표시하기 위한 내부 확인 책임 기준이다.
"""

RYU = "류지영"
BYUN = "변재선(회계)"
OH = "오종현"
YOO = "유현민"

CONFIRMED_SCOPES = (
    {
        "name": RYU,
        "scope": "캠프·일정, 카톡/밴드, 현장자료, 완료일 확인",
        "role": "운영 확인 담당",
    },
    {
        "name": BYUN,
        "scope": "거래명세서, 세금계산서, 입금, 금액 불일치 확인",
        "role": "회계 확인 담당",
    },
    {
        "name": OH,
        "scope": "PO 원본·견적서, 구매·입금 원천자료 취합·누락 확인",
        "role": "원천자료 취합 담당",
    },
    {
        "name": YOO,
        "scope": "PO, ERP·원장 미등록, 시스템 연결 확인",
        "role": "시스템 확인 담당",
    },
)

_SYSTEM_TERMS = ("원장 미등록", "시스템 연결")
_SOURCE_TERMS = (
    "PO 원본", "PO원본", "견적서", "구매자료", "구매 자료",
    "입금 원천자료", "원천자료", "원천 자료", "원본 수집", "원본 누락",
)
_ACCOUNTING_TERMS = (
    "거래명세서", "명세서", "세금계산서", "입금", "수금",
    "금액", "미청구", "청구", "회계",
)
_OPERATIONS_TERMS = (
    "캠프", "일정", "예정일", "접수일자", "카톡", "밴드",
    "현장자료", "현장 자료", "사진", "완료일", "완료보고",
    "작업일", "점검일", "날짜",
)


def confirmed_owner(issue="", category="", direct=""):
    """확정 분류에 해당하면 내부 담당자를, 아니면 기존 담당자를 반환한다."""
    issue = str(issue or "").strip()
    category = str(category or "").strip()
    direct = str(direct or "").strip()
    upper_category = category.upper()
    upper_issue = issue.upper()

    # 오종현은 PO·입금의 판정자가 아니라 원본·견적서·구매/입금 자료 취합 담당이다.
    # 따라서 "PO A" 같은 시스템 대조는 계속 유현민, 입금 금액 판정은 계속 변재선이 맡는다.
    if any(term in issue for term in _SOURCE_TERMS):
        return OH

    # PO·ERP는 설명에 금액·명세서가 함께 있어도 시스템 담당이 우선한다.
    if (upper_category in {"PO", "ERP"}
            or upper_issue == "PO"
            or upper_issue.startswith("PO ")
            or upper_issue == "ERP"
            or upper_issue.startswith("ERP ")
            or any(term in issue for term in _SYSTEM_TERMS)):
        return YOO

    # 정산 카테고리는 청구·입금·금액 확인 업무이므로 회계 담당으로 묶는다.
    if category == "정산" or any(term in issue for term in _ACCOUNTING_TERMS):
        return BYUN

    if any(term in issue for term in _OPERATIONS_TERMS):
        return RYU

    return direct


def assign_issue_row(row):
    """앱/엑셀 공용 확인필요 행에 확정 담당자를 채운 복사본을 반환한다."""
    out = dict(row or {})
    issue = (out.get("문제유형") or out.get("경고내용")
             or out.get("검증결과") or out.get("빈칸") or "")
    category = out.get("구분") or ""
    direct = out.get("담당자") or out.get("담당기사") or ""
    owner = confirmed_owner(issue, category, direct)
    if owner:
        out["담당자"] = owner
    return out
