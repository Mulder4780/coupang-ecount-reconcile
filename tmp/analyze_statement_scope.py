from collections import Counter

from webapp import app_server


rows = app_server.get_settlements().get("rows", [])
paid = [row for row in rows if "유상" in str(row.get("비용구분") or "")]
issued_states = {"있음", "발행", "발행완료", "완료", "반영완료"}


def issued_current(row):
    return bool(
        str(row.get("명세서번호") or "").strip()
        or str(row.get("명세서") or "").strip() in issued_states
    )


def issued_with_date(row):
    return bool(
        issued_current(row)
        or str(row.get("명세서발행일") or "").strip()
    )


current_issued = [row for row in paid if issued_current(row)]
dated = [row for row in paid if str(row.get("명세서발행일") or "").strip()]
revised_unissued = [row for row in paid if not issued_with_date(row)]

print(
    {
        "전체": len(rows),
        "유상": len(paid),
        "현발행": len(current_issued),
        "발행일근거": len(dated),
        "현미발행": len(paid) - len(current_issued),
        "수정미발행": len(revised_unissued),
    }
)
print(Counter(str(row.get("완료일") or "")[:7] for row in revised_unissued))
for row in revised_unissued[:25]:
    print(
        {
            key: row.get(key)
            for key in (
                "정산ID",
                "프로젝트NO",
                "캠프명",
                "명세서",
                "명세서번호",
                "명세서발행일",
                "완료일",
                "공급가액",
                "상태",
            )
        }
    )
