import sys
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from webapp.app_server import representative_summary


report = representative_summary(
    {"as": [], "pm": []},
    [
        {
            "정산ID": "JS-DATE",
            "프로젝트NO": "UJ2600001",
            "업무구분": "돌발AS",
            "캠프명": "날짜근거캠프",
            "완료일": "2026-07-06",
            "비용구분": "유상",
            "명세서": "없음",
            "명세서번호": "",
            "명세서발행일": "2026-07-07",
        },
        {
            "정산ID": "JS-MISSING",
            "프로젝트NO": "UJ2600002",
            "업무구분": "정기점검",
            "캠프명": "미발행캠프",
            "완료일": "2026-07-08",
            "비용구분": "유상",
            "명세서": "없음",
            "명세서번호": "",
            "명세서발행일": "",
        },
    ],
    "2026-07-29",
)
docs = {row["업무유형"]: row for row in report["거래명세서"]["업무유형별"]}
assert docs["돌발 AS"]["미발행"] == 0
assert docs["정기점검"]["미발행"] == 1
assert [row["ID"] for row in report["거래명세서"]["미발행목록"]] == ["JS-MISSING"]
print("거래명세서 발행일 근거 필터 검증 ✅")
