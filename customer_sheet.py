# -*- coding: utf-8 -*-
"""
customer_sheet.py — 거래처코드를 관리대장 안의 시트로 통합 (단일 엑셀 관리)
================================================================================
사용자 지시(2026-08-07): **"쿠팡_거래처코드_최신.xlsx 와 쿠팡_확인필요현황_최신.xlsx 를
관리대장에 통합해서 관리하고, 앞으로도 별도의 엑셀 파일은 만들지 말고 관리대장으로만
관리해. 그리고 필요없는 엑셀은 지워."**

무엇이 문제였나
  `customer_index.py` 가 캠프↔ERP거래처코드 연결 결과를 **별도 엑셀**로 만들고 있었다
  (2026-08-05 "앱과 엑셀에 추가해서 표기" 지시 때, 관리대장에 열을 늘릴 수 없어서
  택한 방법이다). 그런데 파일이 늘수록 "어느 게 최신인가"가 흐려지고, 사람이 열어 둔
  파일은 갱신도 막힌다. 관리대장 **안의 시트**면 그 문제가 통째로 사라진다.

어떻게
  『23_확인필요현황』을 넣는 `findings_sheet.py` 의 기계를 그대로 빌려 쓴다 —
  zip 4개 파트 수술로 시트를 넣고, 이후엔 시트 XML만 갈아끼운다.
  **관리대장을 openpyxl 로 열어 save() 하지 않는다**(차트·도형이 깨진다 — 절대규칙).

  · 내용이 이전과 같으면 새 버전을 만들지 않는다(멱등)
  · 바뀌었으면 vN+1 생성
  · 워크북 규약: 1행 제목, 2행 [사용법], 4행 머리글, 5행부터 데이터, 자동필터·틀고정

실행:  python customer_sheet.py [--master 경로]
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from ecount_reconcile import load_config, resolve_master  # noqa: E402
from findings_sheet import build_generic_sheet, upsert     # noqa: E402

SHEET_NAME = "27_거래처코드"
# 열 구성은 예전 별도 엑셀과 **같게** 둔다 — 사람이 보던 표가 자리만 옮기는 것이지
# 내용이 바뀌는 게 아니어야 한다. 옮기면서 열까지 달라지면 "그 파일 어디 갔냐"가 된다.
HEADERS = ["캠프명", "거래처코드", "ERP거래처명", "연결방식", "원장건수",
           "주소", "담당자", "연락처", "Email", "보유장비"]
WIDTHS = [26, 12, 26, 10, 9, 34, 12, 15, 24, 30]


def rows_from_index(linked, multi, none):
    """customer_index 의 세 갈래를 한 표로 편다. 정렬·순서는 예전 엑셀과 같다."""
    rows = []
    for camp, v in sorted(linked.items()):
        rows.append((camp, v.get("code", ""), v.get("erp_name", ""), v.get("how", ""),
                     v.get("건수", ""), v.get("addr", ""), v.get("manager", ""),
                     v.get("tel", ""), v.get("email", ""), v.get("equip", "")))
    for camp, v in sorted(multi.items()):
        rows.append((camp, "(후보 여럿)", " / ".join(v.get("names") or []), v.get("how", ""),
                     v.get("건수", ""), "", "", "", "", ""))
    for x in sorted(none, key=lambda r: -(r.get("건수") or 0)):
        rows.append((x.get("camp", ""), "(ERP에 없음)", "", "", x.get("건수", ""),
                     "", "", "", "", ""))
    return rows


def build_sheet_xml(rows):
    return build_generic_sheet(
        SHEET_NAME, HEADERS, WIDTHS, rows,
        "[사용법] 캠프명 ↔ ERP 거래처코드 연결 결과입니다. 에이전트가 자동 갱신하므로 "
        "수기 입력하지 마세요. '(후보 여럿)'은 이름이 비슷한 거래처가 둘 이상이라 "
        "사람이 골라야 하는 건, '(ERP에 없음)'은 ERP에 거래처 자체가 없는 건입니다.",
        empty_text="연결 결과 없음 — 원본 자료를 먼저 수집하세요")


def apply(linked, multi, none, master=None):
    """시트를 갱신한다. 세 갈래는 `customer_index` 가 계산해서 넘겨준다.

    ★ 여기서 customer_index 를 다시 부르지 않는다 — 그러면 Z: 를 두 번 훑는다
      (원본이 2만 개라 그것만으로 몇 분이다). 계산은 한 번, 표기는 이 파일이 맡는다.
    """
    master = master or resolve_master(load_config()["reconcile"]["master_xlsx"])
    rows = rows_from_index(linked, multi or {}, none or [])
    ok, msg = upsert(master, build_sheet_xml(rows), sheet_name=SHEET_NAME, headers=HEADERS)
    return rows, ok, msg


def main():
    # 이 파일은 표기만 맡는다. 계산은 customer_index 가 하고 끝에서 apply 를 부른다.
    print("이 도구는 customer_index.py 가 부른다 — "
          "`python customer_index.py --sheet` 로 실행하라.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
