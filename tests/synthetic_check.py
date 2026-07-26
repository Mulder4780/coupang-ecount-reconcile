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


def t14_datesort():
    """날짜 정렬 규칙 — 과거가 위, 최근이 아래. 새 데이터도 자동으로 이 순서를 따라야 한다."""
    import sys as _s
    _s.path.insert(0, os.path.join(ROOT, "webapp"))
    from app_server import norm_date, row_date, sort_by_date
    assert norm_date("2026.6.3") == "2026-06-03", norm_date("2026.6.3")        # 형식 혼재 흡수
    assert norm_date("2026-06-03 00:00:00") == "2026-06-03"
    assert norm_date("") == "" and norm_date(None) == ""
    rows = [{"접수ID": "B", "접수일자": "2026-07-02"},
            {"접수ID": "C", "접수일자": ""},                                    # 날짜 없음 → 맨 뒤
            {"접수ID": "A", "접수일자": "2026.6.3"},
            {"접수ID": "D", "접수일자": "2025-12-31"}]
    got = [r["접수ID"] for r in sort_by_date(rows, "as", "접수ID")]
    assert got == ["D", "A", "B", "C"], got
    # 대표 날짜 열이 없는 시트(확인필요)는 행 안의 아무 날짜나 찾아 쓴다
    assert row_date({"메모": "발생 2026-03-05"}) == "2026-03-05"
    # 같은 날짜면 ID순 — 실행할 때마다 순서가 흔들리면 안 된다
    same = [{"접수ID": "Z", "접수일자": "2026-01-01"}, {"접수ID": "A", "접수일자": "2026-01-01"}]
    assert [r["접수ID"] for r in sort_by_date(same, "as", "접수ID")] == ["A", "Z"]
    print("  [14] 날짜 정렬(과거→최근·빈값 맨뒤·형식혼재·동률ID) ✅")


def t15_doc_ocr():
    """밴드 문서 이미지 파서 — 실제 Windows OCR이 뱉은 오인식 텍스트 그대로 검증.
    (이미지·OCR 엔진 없이 순수 파서만 돌리므로 빠르고 환경에 의존하지 않는다)"""
    import sys as _s
    _s.path.insert(0, os.path.join(ROOT, "band"))
    from doc_ocr import parse_doc, infer_amounts, projects, doc_type, build_updates
    명세서 = """거 래 명 세 서
작성 일자
공급받는자
표로젝트NO
2026-07-20
쿠팡로지스틱스서비스
U]2601138
명세서번호
드록번호
SL-2026-0712
123-81-45678
공급가액
1,850,000
150,000
2,000,000
합계금얙 OJAT 포함)
세액
185,000
15,000
200,000
2 200,000"""
    r = parse_doc(명세서, "s.png")
    assert r["유형"] == "거래명세서", r["유형"]                    # 자간 벌어진 제목
    assert r["프로젝트NO"] == "UJ2601138", r["프로젝트NO"]         # U] → UJ 보정
    assert r["명세서번호"] == "SL-2026-0712", r["명세서번호"]      # 옆칸 숫자와 안 붙어야
    assert (r["공급가액"], r["세액"], r["합계"]) == (2000000, 200000, 2200000), r
    assert r["발행일"] == "2026-07-20" and r["신뢰도"] == "높음", r

    계산서 = """전 자 세 금 계산 서
20260721-41000000-11223344
514-86-01234
U]2601138 인&2Sub
공급가액
2|000,000
작성일자
세액
200,000
2026-07-21
합계금액
2 200,000"""
    r2 = parse_doc(계산서, "t.png")
    assert r2["유형"] == "세금계산서" and r2["명세서번호"] == "", r2
    assert r2["승인번호"] == "20260721-41000000-11223344", r2["승인번호"]
    assert r2["공급가액"] == 2000000, r2                          # '2|000,000' 콤마 오인식 흡수

    # 금액 정합성(공급가+세액=합계)이 안 맞으면 채택하지 않는다 → OCR 자릿수 오류 차단
    assert infer_amounts([1234000, 99999, 5555555]) == (None, None, None)
    assert doc_type("아무 글") == "미상" and projects("UL2601140") == ["UJ2601140"]

    # 자동 입력 후보: 신뢰도 높음 + 판정 일치인 건만
    rows = [{"신뢰도": "높음", "판정": "일치", "정산ID": "JS-1", "파일": "a.png",
             "제안": {"거래명세서번호": "SL-2026-0712", "거래명세서발행일": "2026-07-20"}},
            {"신뢰도": "낮음", "판정": "일치", "정산ID": "JS-2", "파일": "b.png",
             "제안": {"거래명세서번호": "X"}},
            {"신뢰도": "높음", "판정": "금액 불일치 — 대장 1 / 문서 2", "정산ID": "JS-3",
             "파일": "c.png", "제안": {"공급가액": 2}}]
    up = build_updates(rows)
    assert [u["key"] for u in up] == ["JS-1", "JS-1"], up
    assert {u["vtype"] for u in up} == {"text", "date"}, up
    print("  [15] 밴드 문서 OCR 파서(오인식 보정·금액정합·자동입력 게이트) ✅")


def t17_expand():
    """행 확장 — 수식 이동·범위 확장·기존행 보존(합성 워크북)"""
    import sys as _s
    _s.path.insert(0, ROOT)
    from expand_rows import self_test
    self_test()
    print("  [17] 시트 행 확장(수식 재배치·표/검증 범위·기존행 보존) ✅")


def t18_erp_docs():
    """ERP 매출서류 분류·inbox 내용판별 — 파일명이 무작위여도 잡히는가"""
    import sys as _s
    _s.path.insert(0, ROOT)
    from erp_docs_check import work_kind, norm_slip
    from inbox_scan import classify_rows
    assert work_kind("돌발AS_울산2캠프 테이블리프트 작동 멈춤") == "돌발AS"
    assert work_kind("26년 2분기 정기점검 - 서초1MB(양재동B)") == "정기점검"
    assert work_kind("쿠팡신규납품_부산4캠프 2R/T Mobile-lift 2EA") == "신규납품"
    assert work_kind("쿠팡철거_창원1MB 매립형 1EA 철거") == "철거"
    assert work_kind("쿠팡_구리3캠프 전면 연장형 A형 계단 1EA") == "계단"
    assert norm_slip("2026/07/25 -11") == "2026/07/25-11"       # 공백 정규화
    # 내용 판별(파일명 무관)
    assert classify_rows([["회사명 : (주)유니버셜"], ["일자-No.", "거래처명", "담당자 이메일주소",
                          "프로젝트명", "공급가액", "매출부가세", "매출합계"]]) == "tax"
    assert classify_rows([["회사명"], ["일자 - 번호", "거래처명", "담당자 이메일주소",
                          "공급가액", "부가세", "합 계"]]) == "stmt"
    assert classify_rows([["회사명"], ["전표번호", "입력메뉴", "금액", "거래처명", "적요명"]]) == "slips"
    assert classify_rows([["아무 표"], ["가", "나"]]) == "unknown"
    assert classify_rows([["일자", "적요", "차변금액", "대변금액"]]) == "ledger"
    print("  [18] ERP 매출서류 유형분류·inbox 내용판별 ✅")


def t19_workbook_integrity(tmp):
    """엑셀 손상 판정 방지 — 수식 셀의 t 선언과 캐시값(<v>) 정합성.
    실사고: expand_rows가 만든 셀 1,573개가 t="str"인데 <v>가 없어
    엑셀이 열 때마다 '내용에 문제가 있습니다' 복구 대화상자를 띄웠다."""
    import sys as _s, zipfile as _z
    _s.path.insert(0, ROOT)
    from fix_workbook import add_missing_v, scan_sheet, iter_tags
    broken = ('<sheetData><row r="5">'
              '<c r="A5" s="1" t="str"><f>IF(1,2,3)</f></c>'          # ← 캐시값 없음(문제)
              '<c r="B5" s="1" t="str"><f>NOW()</f><v/></c>'          # 정상
              '<c r="C5" s="1"><f>SUM(A1:A2)</f></c>'                 # t 없음 → 대상 아님
              '</row></sheetData>')
    fixed, n = add_missing_v(broken)
    assert n == 1, n
    assert '<c r="A5" s="1" t="str"><f>IF(1,2,3)</f><v/></c>' in fixed, fixed
    assert fixed.count("<v/>") == 2, fixed          # 기존 정상 셀을 건드리지 않음
    _, again = add_missing_v(fixed)
    assert again == 0, "재실행 시 또 고치면 안 됨(멱등)"

    _, fixable = scan_sheet(broken, "t", 10)
    assert fixable == 1, fixable
    bad, _ = scan_sheet('<sheetData><row r="5"><c r="B5"/><c r="A5"/></row></sheetData>', "t", 10)
    assert any("열 순서" in b for b in bad), bad          # 열 역순 검출
    bad2, _ = scan_sheet('<sheetData><row r="7"><c r="A5"/></row></sheetData>', "t", 10)
    assert any("행 7 안에" in b for b in bad2), bad2      # 셀 참조/행 불일치 검출
    bad3, _ = scan_sheet('<sheetData><row r="5"><c r="A5" s="99"/></row></sheetData>', "t", 10)
    assert any("스타일" in b for b in bad3), bad3         # 스타일 인덱스 초과 검출
    # ★ 실사고 회귀: 속성이 붙은 <v>를 못 알아보고 <v/>를 하나 더 넣으면
    #   한 셀에 <v>가 2개가 되고, 엑셀이 그 시트를 통째로 비워 버린다(313셀 피해)
    from fix_workbook import remove_dup_v, count_dup_v, _V_RE
    keep = ('<sheetData><row r="5">'
            '<c r="A5" s="1" t="str"><f>X()</f><v xml:space="preserve">값 </v></c>'
            '</row></sheetData>')
    _, n2 = add_missing_v(keep)
    assert n2 == 0, "xml:space 붙은 <v>를 못 알아봄 — 중복 <v>를 또 만든다"
    assert _V_RE.search('<v xml:space="preserve">a</v>'), "정규식이 속성 있는 <v>를 놓침"
    dup = ('<sheetData><row r="5">'
           '<c r="A5" s="1" t="str"><f>X()</f><v xml:space="preserve">값 </v><v/></c>'
           '<c r="B5" s="1" t="str"><f>Y()</f><v/></c>'
           '</row></sheetData>')
    assert count_dup_v(dup) == 1, count_dup_v(dup)
    fixed3, n3 = remove_dup_v(dup)
    assert n3 == 1 and count_dup_v(fixed3) == 0, fixed3
    assert '<v xml:space="preserve">값 </v></c>' in fixed3, fixed3   # 원래 값은 보존
    assert '<c r="B5" s="1" t="str"><f>Y()</f><v/></c>' in fixed3, fixed3  # 정상 셀은 유지
    print("  [19] 워크북 무결성(수식 캐시값·<v> 중복·행열 순서·스타일 범위) ✅")


def t20_rep_no():
    """대표 프로젝트NO — 모든 건이 번호로 식별되어야 한다.
    실사고: 정규식의 \b가 파일에 **백스페이스 문자**로 저장돼 UJ 번호를 하나도 못 찾았다."""
    import sys as _s
    _s.path.insert(0, os.path.join(ROOT, "webapp"))
    from app_server import rep_no, _UJ_RE, apply_rep_no, build_prj_index
    assert "" not in _UJ_RE.pattern, "정규식에 백스페이스 문자가 섞임"
    assert _UJ_RE.search("명세서 2026/07/01-4 · PO PO367787 · UJ2600975").group() == "UJ2600975"
    assert not _UJ_RE.search("XUJ2600975"), "앞에 글자가 붙으면 잡으면 안 됨"
    # 1) 원본 우선
    assert rep_no({"프로젝트NO": "UJ2601138"}) == ("UJ2601138", "")
    # 2) 본문에서 복원
    assert rep_no({"내용·근거": "명세서 2026/07/01-4 · UJ2600975"}) == ("UJ2600975", "본문")
    # 3) 같은 캠프·같은 달의 실제 작업에서
    idx = build_prj_index({"as": [{"캠프명": "울산2캠프", "접수일자": "2026-05-03",
                                   "프로젝트NO": "UJ2600777"}], "pm": []})
    assert rep_no({"캠프명": "울산2캠프", "완료일": "2026-05-20"}, idx) == ("UJ2600777", "동일캠프·동월")
    # 4) 전표 기반(UJ처럼 보이지 않게)
    n, how = rep_no({"캠프명": "x"}, None, "2026/01/10-2")
    assert (n, how) == ("ERP-260110-2", "전표"), (n, how)
    assert not n.startswith("UJ"), "없는 UJ 번호를 지어내면 안 된다"
    # 5) 최후엔 자체 ID
    assert rep_no({"정산ID": "JS-2607-001"}) == ("JS-2607-001", "자체ID")
    # 6) 일괄 적용 시 빈 건이 남지 않는다
    rows = [{"프로젝트NO": "UJ2601138"}, {"내용·근거": "UJ2600975"}, {"ID": "JS-9"}]
    apply_rep_no(rows)
    assert all(str(r.get("프로젝트NO") or "").strip() for r in rows), rows
    print("  [20] 대표 프로젝트NO(원본·본문·동일캠프·전표·자체ID) ✅")


def t21_reorder():
    """행 재배치 — 과거→최근 정렬 + 수식 상대참조가 함께 이동하는가.
    실사고 회귀: 자기닫힘 <f .../> 를 게으른 정규식으로 잡아 다음 셀의 r=\"B5\"까지
    숫자를 바꿔 'B-3' 같은 잘못된 셀 참조가 만들어졌다."""
    import sys as _s
    _s.path.insert(0, ROOT)
    from reorder_rows import shift_rows, move_row, self_test
    self_test()
    assert shift_rows("$B5+E4", 3) == "$B8+E7"                 # 상대행만 이동
    assert shift_rows("$E$4:E9", 2) == "$E$4:E11"              # 절대행 고정
    row = ('<row r="9"><c r="A9" s="1" t="str"><f t="shared" si="3"/><v/></c>'
           '<c r="B9" s="1"><v>7</v></c>'
           '<c r="C9" s="1" t="str"><f>SUM($B$4:B8)</f><v/></c></row>')
    out = move_row(row, 5)
    assert '<c r="B5" s="1"><v>7</v></c>' in out, out           # ★ 셀 참조가 망가지면 안 됨
    assert '<c r="A5"' in out and '<c r="C5"' in out, out
    assert "SUM($B$4:B4)" in out, out                           # 직전행 참조가 따라 이동
    assert "B-" not in out, out
    print("  [21] 행 재배치(정렬·상대참조 이동·자기닫힘 f 안전) ✅")


def t22_bundle():
    """ERP 계산서 구성 추정 — 캠프 표기 흔들림·분기 제목·금액 정합 판정.
    추정을 확정으로 올려버리면 틀린 배분이 대장에 굳으므로 판정 경계를 못 박아 둔다."""
    import erp_bundle as B
    # 괄호 안 지명이 달라도 같은 캠프
    assert B.camp_key("송파1MB(감일동)") == B.camp_key("송파1MB"), B.camp_key("송파1MB(감일동)")
    assert B.camp_key("송파3Sub-hub") == B.camp_key("송파3subhub")
    assert B.camp_key("양주2캠프") != B.camp_key("양주1캠프")
    # 제목 → 유형
    assert B.kind_of("25년 4분기 정기점검 - 양주2캠프") == "정기점검"
    assert B.kind_of("돌발AS 인천8MB") == "돌발AS"
    assert B.kind_of("쿠팡신규_송파1MB-이동식상하차리프트 2RT") == "신규납품"
    # 제목에 분기가 적혀 있으면 발행월이 아니라 그 분기를 본다
    assert B.window({"month": "2026/01", "title": "25년 4분기 정기점검"}) == ("2025-10", "2025-12")
    assert B.window({"month": "2026/05", "title": "돌발AS"}) == ("2026-02", "2026-05")
    assert B.window({"month": "2026/02", "title": "돌발AS"}) == ("2025-11", "2026-02")  # 연 넘김
    W = [{"prj": "UJ2600001", "camp": "양주2캠프", "kind": "정기점검",
          "date": "2026-04-10", "amt": 600000, "src": "06"},
         {"prj": "UJ2600002", "camp": "양주2캠프", "kind": "정기점검",
          "date": "2026-05-11", "amt": 400000, "src": "06"}]
    doc = {"camp": "양주2캠프", "kind": "정기점검", "month": "2026/05",
           "title": "정기점검 - 양주2캠프", "amt": 1000000}
    prjs, v, tot = B.bundle(doc, W)
    assert prjs == ["UJ2600001", "UJ2600002"] and v == "확정" and tot == 1000000, (prjs, v, tot)
    prjs, v, _ = B.bundle({**doc, "amt": 1020000}, W)      # 2% 차이 → 유력
    assert v == "유력", v
    prjs, v, _ = B.bundle({**doc, "amt": 2000000}, W)      # 100% 차이 → 추정
    assert v == "추정", v
    # 후보를 못 찾으면 **이유**가 남아야 한다("미상" 한 마디는 조치를 못 한다)
    _, v, _ = B.bundle({**doc, "camp": "없는캠프"}, W)
    assert v.startswith("미상(") and "캠프" in v, v
    _, v, _ = B.bundle({**doc, "kind": "철거"}, W)
    assert v.startswith("미상(") and "철거" in v, v
    _, v, _ = B.bundle({**doc, "month": "2026/01"}, W)     # 기간 밖
    assert v.startswith("미상(") and "보유" in v, v
    print("  [22] ERP 계산서 구성 추정(캠프정규화·분기창·금액정합·미상사유) ✅")


def t23_formulas():
    """깨진 수식 복구 — 음수 행 참조·굳어버린 오류값·좁아진 범위.
    이 셋이 엑셀에서 #N/A 691개로 나타났다(2026-07-26). 다시 새지 않게 못 박는다."""
    import fix_formulas as F
    # 상대 행 오프셋을 유지한 채 다른 행에 재생산되는가
    tmpl = F.normalize('IF($B5="","",COUNT($AO$4:AO4)+1)', 5)
    assert F.instantiate(tmpl, 5) == 'IF($B5="","",COUNT($AO$4:AO4)+1)'
    assert F.instantiate(tmpl, 130) == 'IF($B130="","",COUNT($AO$4:AO129)+1)'
    # 깨짐 판정
    assert F.is_broken('COUNT($AO$4:AO-120)+1')      # 음수 행
    assert F.is_broken('#N/A')                       # 오류값이 수식으로 굳음
    assert F.is_broken('SUM(#REF!)')
    assert not F.is_broken('IF($A5="","",1)')
    # ★ 공유수식 참조 셀은 본문이 비어 있다. 이걸 본으로 뽑으면 열 전체가 못 고쳐진다.
    cells = [(5, 'Q', 'IF($A5="","",N($O5))', 0, 0),
             (6, 'Q', '', 0, 0), (7, 'Q', '', 0, 0), (8, 'Q', '', 0, 0)]
    tm = F.templates(cells)
    assert tm.get('Q'), '빈 수식이 최다 득표로 뽑히면 안 된다'
    assert F.instantiate(tm['Q'], 9) == 'IF($A9="","",N($O9))'
    # 범위 확장: 실제 마지막 행까지 늘어나되, 이미 충분하면 건드리지 않는다
    xml = "<f>COUNTIF('02_돌발AS접수'!$A$5:$A$154,1)</f>"
    y, n = F.widen(xml, {'02_돌발AS접수': 344})
    assert '$A$344' in y and n == 1, y
    y2, n2 = F.widen(y, {'02_돌발AS접수': 344})
    assert n2 == 0, '이미 맞는 범위를 또 건드리면 안 됨'
    # 셀 위치 계산: <row> 여는 태그 길이를 빼먹으면 엉뚱한 자리를 덮어쓴다
    sx = ('<sheetData><row r="5" spans="1:3"><c r="A5"><v>1</v></c>'
          '<c r="B5" t="e"><f>#N/A</f><v>#N/A</v></c></row>'
          '<row r="6"><c r="B6"><f>IF($A6="","",2)</f><v>2</v></c></row></sheetData>')
    cells = F.parse_cells(sx)
    b5 = [c for c in cells if c[1] == 'B' and c[0] == 5][0]
    assert sx[b5[3]:b5[4]].startswith('<c r="B5"'), sx[b5[3]:b5[4]][:30]
    fixed, k = F.fix_sheet(sx)
    assert k == 1 and 'IF($A5="","",2)' in fixed and '#N/A' not in fixed, fixed
    assert '<c r="A5"><v>1</v></c>' in fixed, '옆 셀이 망가지면 안 됨'
    print("  [23] 수식 복구(음수행·굳은오류·공유수식본·범위확장·셀위치) ✅")


def t16_status():
    """상태 수식 캐시 보정 — 새로 넣은 행은 상태열(수식)이 None이라 완료된 점검이
    전부 '미점검'으로 보였다. 원본 열로 직접 판정하는 로직 검증."""
    import sys as _s
    _s.path.insert(0, os.path.join(ROOT, "webapp"))
    from app_server import derive_status
    from datetime import date as _d, timedelta as _td
    past = (_d.today() - _td(days=10)).isoformat()
    future = (_d.today() + _td(days=10)).isoformat()
    r = {"점검상태": "", "실제점검일": "2026-05-27", "점검예정일": past}
    derive_status(r, "pm"); assert r["점검상태"] == "완료", r          # 완료일 있으면 완료
    r = {"점검상태": "", "실제점검일": "", "돌발AS전환여부": "Y", "점검예정일": past}
    derive_status(r, "pm"); assert r["점검상태"] == "AS전환", r
    r = {"점검상태": "", "실제점검일": "", "점검예정일": past}
    derive_status(r, "pm"); assert r["점검상태"] == "미점검", r        # 예정일 경과
    r = {"점검상태": "", "실제점검일": "", "점검예정일": future}
    derive_status(r, "pm"); assert r["점검상태"] == "예정", r          # 아직 안 지남
    r = {"점검상태": "완료", "실제점검일": ""}
    derive_status(r, "pm"); assert r["점검상태"] == "완료", r          # 기존 값은 안 건드림
    r = {"진행상태": "", "작업완료일": "2026-06-02"}
    derive_status(r, "as"); assert r["진행상태"] == "작업완료", r
    r = {"진행상태": "", "작업완료일": ""}
    derive_status(r, "as"); assert r["진행상태"] == "접수", r
    print("  [16] 상태 수식 캐시 보정(완료·AS전환·미점검·예정·기존값 보존) ✅")


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
        ds = [r.get("완료일", "") for r in se["rows"] if r.get("완료일")]
        assert ds == sorted(ds), "정산 목록이 과거→최근 순이 아님: " + str(ds[:5])
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
        assert "Coupang Service Operations System" in html and "tabbar" in html and "d_report" in html
        # 원천 검증은 밴드·카톡·ERP·쿠팡PO 4종이 **자료 유무와 무관하게** 항상 표시돼야 한다
        assert "4원천 검증" in html, "원천 검증 제목이 4원천이 아님"
        assert "쿠팡 PO" in html and "PO 목록 투입 시 자동 대조" in html, "PO 원천 행 누락"
        # 기간 필터: 프리셋 + 직접 지정 + 제외 건수 알림
        assert 'id="fperiod"' in html and 'id="fd1"' in html and 'id="fd2"' in html, "기간 선택 UI 누락"
        assert "최근 3개월" in html and "직접 지정" in html, "기간 프리셋 누락"
        assert "날짜 없는" in html, "기간 제외 건수 알림 누락"
        # 통계 탭 월별표는 폰에서도 보여야 한다(카드 대체본이 없어 숨기면 화면이 빈다)
        assert "table.grid.ptbl{display:table!important" in html.replace(" ", ""), "폰에서 월별표가 숨겨짐"
        assert "erpHtml(" in html and "api/erpdocs" in html, "ERP 매출 반영 누락"
        # 처리 안내: 확인 필요 건마다 '어디서 확인·어떻게 반영'이 붙어야 한다
        assert "이 건은 어떻게 처리하나요?" in html, "처리 안내 카드 누락"
        for k in ("금액 미입력", "세금계산서 미발행", "PO 미발행", "ERP 계산서(묶음)"):
            assert f"'{k}'" in html or f'"{k}"' in html, f"{k} 안내 누락"
        assert "band/docs_inbox" in html and "처리 방법 · 용어 설명" in html, "도움말 항목 누락"
        # 대표보고: 빈 절을 건너뛰고 요약이 '정리' 절 안으로 들어가야 한다(목차 중복 방지)
        assert "usedSum" in html and "empty(s) && !isSum" in html, "빈 절 건너뛰기 로직 누락"
        assert ".esec" in html and ".esum" in html and ".egroup" in html, "대표보고 스타일 누락"
        # 목록 카드·표에서 프로젝트NO가 맨 앞·굵게 나와야 한다(4개 탭 전부)
        assert html.count('class="prjno"') >= 6, "프로젝트NO 강조 표기 부족"
        assert '<b class="prjno">' in html, "표에서 프로젝트NO 굵게 표기 누락"
        assert 'class="sid"' in html, "보조 ID 표기 누락"
        # 구버전 감지: 폰에 켜 둔 앱이 예전 화면을 계속 쓰지 않도록
        assert "checkBuild" in html and "hardReload" in html, "구버전 자동 갱신 누락"
        pg = json.loads(urllib.request.urlopen(base + "/api/ping", timeout=5).read())
        assert pg.get("build"), "ping 응답에 build 값 없음"
        # 당일 업무 실적: 항목마다 어떤 건인지 상세가 붙고, 숫자가 어긋나면 알려야 한다
        assert "dayDetail(" in html and "dayRows(" in html, "당일 실적 상세 누락"
        assert "앱에서 찾은 건" in html, "보고 숫자와 불일치 알림 누락"
        for k in ("신규접수", "점검완료", "거래명세서발행", "입금건수"):
            assert f"'{k}'" in html, f"{k} 매핑 누락"
        # 문제 코드는 축약어가 아니라 '무엇이 비었고 무엇을 해야 하는지'로 풀어 써야 한다
        assert "관리자검증상태" in html and "중복판정(선택)" in html, "문제 코드 해설 누락"
        assert "이미 정리(중복 통합" in html, "삭제된 건 표시 누락"
        assert "codeHtml(" in html and "topLine(" in html, "코드 해설 함수 누락"
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
    t14_datesort()
    t15_doc_ocr()
    t16_status()
    t17_expand()
    t18_erp_docs()
    t19_workbook_integrity(None)
    t20_rep_no()
    t21_reorder()
    t22_bundle()
    t23_formulas()
    t6_webapp()
    print("ALL GREEN — 실작업 진행 가능")
