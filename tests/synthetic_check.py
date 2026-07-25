# -*- coding: utf-8 -*-
"""
synthetic_check.py — 합성데이터 상시 검증 (실데이터·실서버 접촉 0)
====================================================================
사용자 상시 지시(2026-07-24): "항상 합성데이터 검증해서 작업 진행".
실제 시스템을 건드리기 전에 이 스크립트가 초록이어야 한다.

검증 항목:
  1. ERP원장 대조기(erp_ledger_check): 유형 A/B/C/D 판정이 설계대로 나오는가
  2. 전표 페이로드(ecount_upload.build_payload): 금액·부가세·일자·프로젝트 매핑
  3. 판매 inbox 매칭(ecount_reconcile.match_project): 프로젝트NO/금액 매칭 규칙

실행:  python tests/synthetic_check.py   (전부 통과 시 exit 0, 'ALL GREEN')
"""
import sys, os, re, tempfile, subprocess

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable


def make_ledger(path):
    wb = openpyxl.Workbook()
    ws6 = wb.active; ws6.title = "06_거래서류청구수금"
    h6 = ["정산ID", "업무구분", "원천업무ID", "프로젝트NO", "캠프명", "작업완료일", "비용구분",
          "실제작업공급가액", "실제작업부가세", "실제작업합계", "거래명세서번호", "거래명세서발행일",
          "거래명세서공급가액", "거래명세서합계", "세금계산서발행일", "세금계산서합계", "입금일", "입금액",
          "PO필요여부", "PO번호", "PO발행일"]
    for _ in range(3): ws6.append([])
    ws6.append(h6)
    # JS-1: ERP 일치 + 세금계산서 미발행 → C / PO111 쿠팡 일치
    ws6.append(["JS-1", "돌발AS", "AS-1", "UJ0001", "테스트캠프A", "2026-07-01", "유상",
                100000, 10000, 110000, "2026/07/01-4", "2026-07-01", 100000, 110000, None, 0, None, None,
                "필요", "PO111", "2026-07-03"])
    # JS-2: ERP 금액불일치 → D / PO333은 쿠팡 목록에 없음 → PO-B
    ws6.append(["JS-2", "정기점검", "PM-1", "UJ0002", "테스트캠프B", "2026-07-02", "유상",
                200000, 20000, 220000, "2026/07/02-1", "2026-07-02", 200000, 220000, "2026-07-03", 220000, None, None,
                "필요", "PO333", "2026-07-04"])
    # JS-3: 명세서번호 없음 → B / PO필요·번호없음 → PO444 유일매칭 → PO-D
    ws6.append(["JS-3", "돌발AS", "AS-2", "UJ0003", "테스트캠프C", "2026-07-05", "유상",
                300000, 30000, 330000, None, None, None, 0, None, 0, None, None,
                "필요", None, None])
    ws15 = wb.create_sheet("15_세금계산서관리")
    for _ in range(3): ws15.append([])
    ws15.append(["정산ID", "실제발행일", "발행금액", "승인번호"])
    ws15.append(["JS-2", "2026-07-03", 220000, "APPR-001"])
    ws16 = wb.create_sheet("16_입금수금관리")
    for _ in range(3): ws16.append([])
    ws16.append(["정산ID", "입금일", "입금액", "미수금액"])
    ws2 = wb.create_sheet("02_돌발AS접수")
    for _ in range(3): ws2.append([])
    ws2.append(["접수ID", "프로젝트NO", "캠프명", "담당기사", "작업완료일", "진행상태"])
    ws2.append(["AS-K1", "UJ0001", "테스트캠프A(감일동)", "김필우", "2026-07-01", "작업완료"])
    ws2.append(["AS-K2", "UJ0002", "다른캠프B(성수동)", "김준형", "2026-07-02", "작업완료"])
    wb.save(path)


def make_erp(path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "원장"
    ws.append(["거래처별계정별원장 (합성)"])
    ws.append(["일자-No.", "적요", "차변금액", "대변금액"])
    ws.append(["2026/07/01 -4", "테스트캠프A 돌발AS", None, 110000])        # JS-1 일치
    ws.append(["2026/07/02 -1", "테스트캠프B 정기점검", None, 999999])      # JS-2 금액불일치
    ws.append(["2026/07/05 -9", "인천8MB 상하차리프트(원장에 없음)", None, 500000])  # A형
    ws.append(["2026/07 계", "", None, 1609999])
    wb.save(path)


def t1_erp_check(tmp):
    ledger = os.path.join(tmp, "ledger.xlsx"); erp = os.path.join(tmp, "erp원장.xlsx")
    make_ledger(ledger); make_erp(erp)
    r = subprocess.run([PY, os.path.join(ROOT, "erp_ledger_check.py"),
                        "--file", erp, "--master", ledger],
                       capture_output=True, text=True, encoding="utf-8", cwd=ROOT, env={**os.environ, "COUPANG_REPORT_DIR": tmp, "COUPANG_UPDATES_DIR": tmp})
    out = r.stdout
    m = re.search(r"A\(ERP에만\) (\d+) / B\(원장에만\) (\d+) / C\(계산서미발행\) (\d+) / D\(금액불일치\) (\d+) / 정상 (\d+)", out)
    assert m, f"결과 라인 파싱 실패:\n{out}\n{r.stderr}"
    a, b, c, d, ok = map(int, m.groups())
    assert (a, b, c, d, ok) == (1, 1, 1, 1, 1), f"판정 불일치: A{a} B{b} C{c} D{d} OK{ok} (기대 1,1,1,1,1)"
    print("  [1] ERP원장 대조 A/B/C/D 판정 ✅")


def t2_payload():
    from ecount_upload import build_payload
    rec = {"정산ID": "JS-9", "캠프명": "합성캠프", "업무구분": "돌발AS",
           "프로젝트NO": "UJ9999", "작업완료일": "2026-07-20", "원장_공급가액": 1472500.0}
    up = {"CUST": "CUSTX", "CR_CODE": "4049", "TAX_GUBUN": "11"}
    p = build_payload(rec, up)["InvoiceAutoList"][0]["BulkDatas"]
    assert p["SUPPLY_AMT"] == "1472500" and p["VAT_AMT"] == "147250", p
    assert p["TRX_DATE"] == "20260720" and p["PJT_CD"] == "UJ9999" and p["CR_CODE"] == "4049", p
    assert "JS-9" in p["REMARKS"], p
    print("  [2] 전표 페이로드 매핑 ✅")


def t3_match():
    from ecount_reconcile import match_project, norm_ecount
    ec = norm_ecount([{"SUPPLY_AMT": "110000", "IO_DATE": "20260701", "REMARKS": "UJ0001 테스트캠프A", "CUST_DES": "쿠팡"},
                      {"SUPPLY_AMT": "500000", "IO_DATE": "20260705", "REMARKS": "무관 건", "CUST_DES": "쿠팡"}])
    rec = {"프로젝트NO": "UJ0001", "원장_공급가액": 110000.0}
    m, how = match_project(rec, ec, 0, "쿠팡")
    assert m and how == "프로젝트NO", (m, how)
    rec2 = {"프로젝트NO": "UJ_NOPE", "원장_공급가액": 500000.0}
    m2, how2 = match_project(rec2, ec, 0, "쿠팡")
    assert m2 and how2 == "금액", (m2, how2)
    print("  [3] inbox 판매 매칭 규칙 ✅")


def t4_kakao(tmp):
    ledger = os.path.join(tmp, "ledger_k.xlsx")
    make_ledger(ledger)
    txt = os.path.join(tmp, "쿠팡AS방.txt")
    with open(txt, "w", encoding="utf-8") as f:
        f.write("쿠팡AS방 님과 카카오톡 대화\n저장한 날짜 : 2026-07-24 18:00:00\n\n")
        f.write("--------------- 2026년 7월 1일 화요일 ---------------\n")
        f.write("[김필우] [오후 2:59] UJ0001 테스트캠프A 작업완료\n")
        f.write("사진 2장 첨부\n")                                   # 여러 줄 병합 검증
        f.write("2026년 7월 3일 오후 4:01, 유현민 : 일반 공지 메시지\n")  # 모바일 형식 검증
    r = subprocess.run([PY, os.path.join(ROOT, "kakao", "kakao_reconcile.py"),
                        "--file", txt, "--master", ledger],
                       capture_output=True, text=True, encoding="utf-8", cwd=ROOT, env={**os.environ, "COUPANG_REPORT_DIR": tmp, "COUPANG_UPDATES_DIR": tmp})
    m = re.search(r"확인 (\d+) / 미확인 (\d+)", r.stdout)
    assert m, f"카톡 결과 파싱 실패:\n{r.stdout}\n{r.stderr}"
    ok, miss = map(int, m.groups())
    assert (ok, miss) == (1, 1), f"카톡 판정 불일치: 확인{ok} 미확인{miss} (기대 1,1)"
    print("  [4] 카톡 내보내기 파싱·대조 (PC/모바일 형식·다중행 병합) ✅")


def t5_writer(tmp):
    import json
    src = os.path.join(tmp, "합성대장_v1.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "06_거래서류청구수금"
    for _ in range(3): ws.append([])
    ws.append(["정산ID", "거래명세서번호", "거래명세서발행일", "비고"])
    ws.append(["JS-W1", None, None, "빈칸 채움 대상"])
    ws.append(["JS-W2", "기존값-유지", None, "덮어쓰기 금지 검증"])
    h = wb.create_sheet("19_AI작업인수인계")
    h.append(["헤더"]); h.append(["기존 인수인계 행"])
    wb.save(src)
    q = os.path.join(tmp, "q.json")
    json.dump([
        {"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": "JS-W1", "col": "거래명세서번호",
         "value": "2026/07/24-9", "vtype": "text", "evidence": "합성", "only_if_empty": True},
        {"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": "JS-W1", "col": "거래명세서발행일",
         "value": "2026-07-24", "vtype": "date", "evidence": "합성", "only_if_empty": True},
        {"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": "JS-W2", "col": "거래명세서번호",
         "value": "덮어쓰기시도", "vtype": "text", "evidence": "합성", "only_if_empty": True},
    ], open(q, "w", encoding="utf-8"), ensure_ascii=False)
    r = subprocess.run([PY, os.path.join(ROOT, "ledger_writer.py"), "--apply",
                        "--master", src, "--queue", q],
                       capture_output=True, text=True, encoding="utf-8", cwd=ROOT, env={**os.environ, "COUPANG_REPORT_DIR": tmp, "COUPANG_UPDATES_DIR": tmp})
    assert "반영 완료" in r.stdout and "입력 2건 / 건너뜀 1건" in r.stdout, f"{r.stdout}\n{r.stderr}"
    dst = src.replace("_v1.xlsx", "_v2.xlsx")
    w2 = openpyxl.load_workbook(dst)
    ws2 = w2["06_거래서류청구수금"]
    assert ws2["B5"].value == "2026/07/24-9", ws2["B5"].value          # 빈칸 채움
    assert ws2["C5"].value is not None, "날짜 직렬값 기록 실패"          # date serial
    assert ws2["B6"].value == "기존값-유지", ws2["B6"].value            # 덮어쓰기 금지
    hand = w2["19_AI작업인수인계"]
    assert any("자동 입력" in str(c.value) for row in hand.iter_rows() for c in row if c.value), "인수인계 기록 없음"
    w2.close()
    # 회귀: 스타일만 있는 자기닫힘 빈 셀(<c s=.../>) 바로 뒤에 수식 셀 — 오매칭으로 '값 있음' 오판했던 실사고
    from ledger_writer import apply_to_xml
    xml = ('<worksheet><dimension ref="A1:C5"/><sheetData>'
           '<row r="5"><c r="A5" t="inlineStr"><is><t>K</t></is></c>'
           '<c r="B5" s="43"/><c r="C5" s="53"><f t="shared" si="1"/><v>0</v></c></row>'
           '</sheetData></worksheet>')
    ok, nx, why = apply_to_xml(xml, {"row": 5, "colL": "B", "value": "값", "vtype": "text", "only_if_empty": True})
    assert ok and '<c r="B5" s="43" t="inlineStr"><is><t>값</t></is></c>' in nx, (ok, why, nx)
    ok2, _, why2 = apply_to_xml(nx, {"row": 5, "colL": "C", "value": "X", "vtype": "text", "only_if_empty": True})
    assert not ok2 and "이미" in why2, (ok2, why2)     # 수식 셀은 여전히 보호
    print("  [5] 자동입력 엔진(빈칸만·날짜·덮어쓰기금지·자기닫힘셀 회귀·인수인계 기록) ✅")


def t7_po(tmp):
    ledger = os.path.join(tmp, "ledger_po.xlsx")
    make_ledger(ledger)
    pof = os.path.join(tmp, "쿠팡PO목록_테스트.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["쿠팡 PO 발행 목록"])
    ws.append(["발행일자", "PO번호", "금액"])
    ws.append(["2026-07-03", "PO111", 100000])   # JS-1 일치
    ws.append(["2026-07-05", "PO222", 999000])   # 원장 미등록 → A
    ws.append(["2026-07-06", "PO444", 300000])   # JS-3 유일매칭 → A+D
    wb.save(pof)
    r = subprocess.run([PY, os.path.join(ROOT, "po_reconcile.py"),
                        "--file", pof, "--master", ledger, "--no-queue"],
                       capture_output=True, text=True, encoding="utf-8", cwd=ROOT, env={**os.environ, "COUPANG_REPORT_DIR": tmp, "COUPANG_UPDATES_DIR": tmp})
    m = re.search(r"A\(원장미등록\) (\d+) / B\(쿠팡목록에없음\) (\d+) / C\(금액불일치\) (\d+) / D\(연결제안\) (\d+) / 정상 (\d+)", r.stdout)
    assert m, f"PO 결과 파싱 실패:\n{r.stdout}\n{r.stderr}"
    a, b, c, d, ok = map(int, m.groups())
    assert (a, b, c, d, ok) == (2, 1, 0, 1, 1), f"PO 판정 불일치: A{a} B{b} C{c} D{d} OK{ok} (기대 2,1,0,1,1)"
    import glob as _g
    assert "유일 매칭" in open(sorted(_g.glob(os.path.join(tmp, "PO대조_*.md")))[-1], encoding="utf-8").read()
    print("  [7] 쿠팡 PO 대조(원장미등록·오기입·금액·유일매칭 제안) ✅")


def t8_findings_sheet(tmp):
    ledger = os.path.join(tmp, "합성대장F_v1.xlsx")
    make_ledger(ledger)
    env = {**os.environ, "COUPANG_REPORT_DIR": tmp, "COUPANG_UPDATES_DIR": tmp}
    r = subprocess.run([PY, os.path.join(ROOT, "findings_sheet.py"), "--master", ledger],
                       capture_output=True, text=True, encoding="utf-8", cwd=ROOT, env=env)
    assert "시트 신규 추가" in r.stdout, f"{r.stdout}\n{r.stderr}"
    v2 = ledger.replace("_v1.xlsx", "_v2.xlsx")
    w = openpyxl.load_workbook(v2)
    ws = w["23_확인필요현황"]
    assert [c.value for c in ws[4][:3]] == ["구분", "ID", "문제유형"]
    vals = [ws.cell(row=i, column=1).value for i in range(5, ws.max_row + 1)]
    assert "정산" in vals, vals          # 합성 정산 조치필요가 들어와야 함
    w.close()
    # 멱등: 같은 내용으로 재실행 → v3 생성 안 됨
    r2 = subprocess.run([PY, os.path.join(ROOT, "findings_sheet.py"), "--master", v2],
                        capture_output=True, text=True, encoding="utf-8", cwd=ROOT, env=env)
    assert "변경 없음" in r2.stdout and not os.path.exists(ledger.replace("_v1", "_v3")), r2.stdout
    print("  [8] 확인필요 시트 통합(신규 추가·머리글·멱등) ✅")


def t10_band_extract():
    from band_extract import parse_post
    mk = lambda c: {"content": c, "created_at": 1780000000000, "photo_count": 4}
    r = parse_post("1", mk(
        "2026년 6월 1일 오전 8:08 게시글\n\n☑️판매전표 +거래명세서 +견적서 = 메일발송 完 ⭕\n"
        "♣ ［ 2026년 02분기 3개월 유료 A/S 완료 ]\n● A/S 일자 : 2026.06.01 (월요일)\n"
        "● A/S 담당 : 김필우\n● 프로젝트NO : UJ2600931\n● 캠프이름 : 양주1캠프\n\n...더보기"), "밴드A")
    assert r["프로젝트NO"] == "UJ2600931" and r["업무유형"] == "정기점검", r
    assert r["비용구분"] == "유상" and r["작업일"] == "2026-06-01" and r["진행상태"] == "작업완료", r
    assert r["담당기사"] == "김필우" and r["캠프명"] == "양주1캠프", r
    assert "판매전표" in r["문서상태"] and "메일발송" in r["문서상태"], r
    r2 = parse_post("2", mk("♣ ［ 돌발무료 A/S 안내 ]\n● A/S 일자 : 2026.00.00 (요일)\n"
                            "● 프로젝트NO : UJ2601999\n● 캠프이름 : 테스트캠프"), "밴드A")
    assert r2["업무유형"] == "돌발AS" and r2["비용구분"] == "무상", r2
    assert r2["작업일"] == "" and r2["진행상태"] == "접수·예정", r2      # 2026.00.00 = 미정
    assert parse_post("3", mk("● 프로젝트NO : UJ000000\n♣ ［ 돌발유료 A/S 안내 ]"), "밴드A") is None  # 템플릿 제외
    assert parse_post("4", mk("안녕하세요 일반 공지입니다"), "밴드A") is None                          # 비업무 제외
    print("  [10] 밴드 업무 추출(유형·유상무상·미정일자·템플릿/비업무 제외) ✅")


def t11_backfill():
    """백필 핵심 로직: 중복 제거·코드값 변환·용량 판단 (실데이터·실파일 접촉 없음)"""
    from backfill_rows import dedupe, enrich, SPEC
    base = {"프로젝트NO": "", "업무유형": "정기점검", "비용구분": "유상", "작업일": "", "담당기사": "김필우",
            "캠프명": "테스트캠프", "진행상태": "", "문서상태": "", "사진": 0, "게시일": "", "밴드": "B"}
    recs = [
        {**base, "프로젝트NO": "UJ1", "진행상태": "접수·예정", "게시일": "2026-06-01"},
        {**base, "프로젝트NO": "UJ1", "진행상태": "작업완료", "작업일": "2026-06-05", "게시일": "2026-06-05"},
        {**base, "프로젝트NO": "UJ2", "진행상태": "작업완료", "작업일": "2026-06-03", "게시일": "2026-06-03"},
    ]
    d = dedupe(recs)
    assert len(d) == 2, d                                   # UJ1 2건 → 1건
    u1 = [x for x in d if x["프로젝트NO"] == "UJ1"][0]
    assert u1["진행상태"] == "작업완료" and u1["_중복"] == 2, u1   # 완료본이 대표
    e = enrich(u1)
    assert e["_진행상태"] == "작업완료" and e["_완료일"] == "2026-06-05", e
    assert "게시 2건 통합" in e["_출처"], e
    e2 = enrich({**base, "프로젝트NO": "UJ3", "진행상태": "접수·예정", "작업일": "", "게시일": "2026-06-09"})
    assert e2["_진행상태"] == "접수" and e2["_완료일"] == "", e2   # 미완료는 완료일 비움
    # 매핑에 ID·자동계산 열이 없어야 함(수식 보호)
    for kind, spec in SPEC.items():
        assert not any(c.endswith("ID") for c in spec["map"]), spec["map"]
        assert "점검상태" not in spec["map"] and "기사배정상태" not in spec["map"], spec["map"]
    print("  [11] 백필 로직(중복통합·상태변환·완료일·수식열 제외) ✅")


def t12_dedupe(tmp):
    """중복 정리: 값만 비우고 수식 셀은 반드시 보존 (합성 워크북)"""
    import zipfile
    from dedupe_rows import blank_cells
    src = os.path.join(tmp, "합성중복_v1.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "02_돌발AS접수"
    for _ in range(3): ws.append([])
    ws.append(["접수ID", "프로젝트NO", "캠프명", "비고"])
    ws.append(["=IF(B5=\"\",\"\",\"AS-1\")", "UJ1", "캠프A", "유지"])
    ws.append(["=IF(B6=\"\",\"\",\"AS-2\")", "UJ1", "캠프A", "중복"])
    wb.save(src)
    dst, cleared, kept, _ = blank_cells(src, [("02_돌발AS접수", 6, ["A", "B", "C", "D"])])
    w = openpyxl.load_workbook(dst)
    ws2 = w["02_돌발AS접수"]
    assert ws2["B6"].value is None and ws2["D6"].value is None, (ws2["B6"].value, ws2["D6"].value)  # 값 제거
    assert str(ws2["A6"].value).startswith("="), ws2["A6"].value      # ★수식 셀은 보존
    assert ws2["B5"].value == "UJ1" and ws2["D5"].value == "유지"      # 유지 행 무손상
    w.close()
    assert cleared >= 2 and kept >= 1, (cleared, kept)
    assert zipfile.ZipFile(dst).testzip() is None
    print("  [12] 중복정리(값만 비움·수식 보존·유지행 무손상) ✅")


def t13_fix_ids(tmp):
    """ID 확정: 수식 규칙 재현·데이터 행만·빈 행 수식 보존"""
    from fix_ids import make_id, scan, apply_ids
    from datetime import date as _d
    assert make_id("AS", _d(2026, 6, 1), 66) == "AS-2606-062", make_id("AS", _d(2026, 6, 1), 66)
    assert make_id("PM", _d(2026, 7, 13), 45) == "PM-2607-041"
    src = os.path.join(tmp, "합성ID_v1.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "02_돌발AS접수"
    for _ in range(3): ws.append([])
    ws.append(["접수ID", "프로젝트NO", "캠프명", "접수일자"])
    ws.append(["AS-2607-001", "UJ1", "캠프A", _d(2026, 7, 1)])          # 기존 ID → 보존
    ws["A6"] = '=IF($B6="","","AS-"&TEXT($D6,"yymm")&"-"&TEXT(ROW()-4,"000"))'
    ws["B6"], ws["C6"], ws["D6"] = "UJ2", "캠프B", _d(2026, 6, 1)        # 데이터 있음 → 확정 대상
    ws["A7"] = '=IF($B7="","","AS-"&TEXT($D7,"yymm")&"-"&TEXT(ROW()-4,"000"))'  # 빈 행 → 수식 유지
    wb.save(src)
    plans, dups = scan(src)
    assert not dups and len(plans) == 1 and plans[0]["row"] == 6, (plans, dups)
    assert plans[0]["id"] == "AS-2606-002", plans[0]
    dst, done, _ = apply_ids(src, plans)
    w = openpyxl.load_workbook(dst)
    ws2 = w["02_돌발AS접수"]
    assert ws2["A5"].value == "AS-2607-001", ws2["A5"].value            # 기존 무손상
    assert ws2["A6"].value == "AS-2606-002", ws2["A6"].value            # 확정됨
    assert str(ws2["A7"].value).startswith("="), ws2["A7"].value        # ★빈 행 수식 보존
    w.close()
    print("  [13] ID 확정(수식규칙 재현·기존보존·빈행 수식유지) ✅")


def t9_watchdog():
    import time as _t
    from watchdog import pick_archive, pick_old_reports
    vers = [(19, "v19"), (25, "v25"), (21, "v21"), (24, "v24"), (23, "v23")]
    assert set(pick_archive(vers, keep=3)) == {"v19", "v21"}, pick_archive(vers, 3)   # 최신 3(25,24,23) 보존
    assert pick_archive(vers[:3], keep=3) == []                                        # 3개 이하면 이동 없음
    now = _t.time()
    files = [("old.md", now - 40*86400), ("new.md", now - 5*86400),
             ("agent_status.json", now - 90*86400), ("tunnel_url.txt", now - 90*86400)]
    dele = pick_old_reports(files, days=30)
    assert dele == ["old.md"], dele                                                    # 보호파일·최근파일 제외
    print("  [9] 워치독 판단 로직(버전 보존 3개·보호파일·30일 기준) ✅")


def t6_webapp():
    import time, json, urllib.request
    port = 18899
    p = subprocess.Popen([PY, os.path.join(ROOT, "webapp", "app_server.py"), "--demo", "--port", str(port)],
                         cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        base = f"http://127.0.0.1:{port}"
        for _ in range(30):                       # 기동 대기
            try:
                urllib.request.urlopen(base + "/api/ping", timeout=1); break
            except Exception:
                time.sleep(0.3)
        # 로그인(잘못된 PIN → 401, 데모 PIN 0000 → ok)
        req = urllib.request.Request(base + "/api/login", data=b'{"pin":"9999"}', method="POST")
        try:
            urllib.request.urlopen(req); assert False, "잘못된 PIN이 통과됨"
        except urllib.error.HTTPError as e:
            assert e.code == 401
        req = urllib.request.Request(base + "/api/login", data=b'{"pin":"0000"}', method="POST")
        assert json.loads(urllib.request.urlopen(req).read())["ok"]
        # 인증 상태·정산 API
        h = {"X-Pin": "0000"}
        st = json.loads(urllib.request.urlopen(urllib.request.Request(base + "/api/status", headers=h)).read())
        assert st.get("demo") and st["steps"], st
        se = json.loads(urllib.request.urlopen(urllib.request.Request(base + "/api/settlements", headers=h)).read())
        assert len(se["rows"]) >= 10 and se["rows"][0]["정산ID"].startswith("JS-"), len(se.get("rows", []))
        # 기준일 설정(데모: 시뮬레이션 응답) — 잠금 테스트 이전에 수행
        req = urllib.request.Request(base + "/api/set_dates", data='{"보고일":"2026-07-25"}'.encode("utf-8"),
                                     headers={"X-Pin": "0000"}, method="POST")
        d2 = json.loads(urllib.request.urlopen(req).read())
        assert d2.get("ok") and d2.get("demo"), d2
        # 미인증 접근 차단
        try:
            urllib.request.urlopen(base + "/api/settlements"); assert False, "무인증 접근 허용됨"
        except urllib.error.HTTPError as e:
            assert e.code == 401
        # 브루트포스 잠금: 5회 실패 후 429
        for _ in range(5):
            try:
                urllib.request.urlopen(urllib.request.Request(base + "/api/login", data=b'{"pin":"1111"}', method="POST"))
            except urllib.error.HTTPError:
                pass
        try:
            urllib.request.urlopen(urllib.request.Request(base + "/api/login", data=b'{"pin":"0000"}', method="POST"))
            assert False, "잠금 미작동"
        except urllib.error.HTTPError as e:
            assert e.code == 429, e.code
        # 메인 페이지 서빙
        html = urllib.request.urlopen(base + "/").read().decode("utf-8")
        assert "쿠팡 통합업무" in html and "tabbar" in html and "d_report" in html
        print("  [6] 웹앱 API(PIN 인증·상태·정산·페이지 서빙) ✅")
    finally:
        p.terminate()


if __name__ == "__main__":
    print("합성데이터 검증 시작 (실데이터·실서버 접촉 없음)")
    with tempfile.TemporaryDirectory() as tmp:
        t1_erp_check(tmp)
        t4_kakao(tmp)
        t5_writer(tmp)
        t7_po(tmp)
        t8_findings_sheet(tmp)
        t12_dedupe(tmp)
        t13_fix_ids(tmp)
    t2_payload()
    t3_match()
    t9_watchdog()
    t10_band_extract()
    t11_backfill()
    t6_webapp()
    print("ALL GREEN — 실작업 진행 가능")
