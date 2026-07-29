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
    dash = wb.create_sheet("00_대시보드"); dash["A2"] = "기존 사용법"
    m18 = wb.create_sheet("18_문서발행업무매뉴얼"); m18["A1"] = "기존 매뉴얼"
    m20 = wb.create_sheet("20_쿠팡통합업무상세매뉴얼"); m20.append([1, "기존 상세"])
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
    # 종료 인수인계와 상시지시(00 A2·18·20)를 한 버전에 함께 반영한다.
    import workbook_patch as WP
    v3 = src.replace("_v1.xlsx", "_v3.xlsx")
    WP.patch(dst, v3, "2026-07-28 #합성", "보완", "상세",
             "새 사용법", "18 갱신", "20 갱신")
    w3 = openpyxl.load_workbook(v3, read_only=True)
    assert w3["00_대시보드"]["A2"].value == "새 사용법"
    assert w3["18_문서발행업무매뉴얼"]["A2"].value == "18 갱신"
    assert w3["20_쿠팡통합업무상세매뉴얼"]["D2"].value == "20 갱신"
    w3.close()
    # 회귀: 스타일만 있는 자기닫힘 빈 셀(<c s=.../>) 바로 뒤에 수식 셀 — 오매칭으로 '값 있음' 오판했던 실사고
    import ledger_writer as LW
    from ledger_writer import apply_to_xml
    xml = ('<worksheet><dimension ref="A1:C5"/><sheetData>'
           '<row r="5"><c r="A5" t="inlineStr"><is><t>K</t></is></c>'
           '<c r="B5" s="43"/><c r="C5" s="53"><f t="shared" si="1"/><v>0</v></c></row>'
           '</sheetData></worksheet>')
    ok, nx, why = apply_to_xml(xml, {"row": 5, "colL": "B", "value": "값", "vtype": "text", "only_if_empty": True})
    assert ok and '<c r="B5" s="43" t="inlineStr"><is><t>값</t></is></c>' in nx, (ok, why, nx)
    ok2, _, why2 = apply_to_xml(nx, {"row": 5, "colL": "C", "value": "X", "vtype": "text", "only_if_empty": True})
    assert not ok2 and "이미" in why2, (ok2, why2)     # 수식 셀은 여전히 보호
    # 같은 기준일을 다시 저장하면 버전만 하나 더 생기면 안 된다.
    date_xml = ('<worksheet><sheetData><row r="3"><c r="B3" s="1"><v>46226</v></c>'
                '</row></sheetData></worksheet>')
    same, unchanged, why3 = apply_to_xml(
        date_xml, {"row": 3, "colL": "B", "value": "2026-07-23",
                   "vtype": "date", "only_if_empty": False})
    assert not same and unchanged == date_xml and "동일 값" in why3, (same, why3)
    # 큐 쓰기는 원자 교체 + 중복 방지여야 한다.
    old_pending = LW.PENDING
    try:
        LW.PENDING = os.path.join(tmp, "atomic_queue.json")
        item = {"sheet": "S", "key": "K", "col": "C"}
        assert LW.queue_add([item, item]) == 1
        assert LW.load_queue() == [item]
    finally:
        LW.PENDING = old_pending
    print("  [5] 자동입력 엔진(빈칸만·동일값 멱등·큐잠금/원자쓰기·자기닫힘셀·인수인계) ✅")


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
    # 판정 이름이 '원장미등록' → '미청구'로 바뀌었다(ERP 계산서 대조 후 진짜 미청구만 남김)
    m = re.search(r"A\((?:원장미등록|미청구)\) (\d+) / B\(쿠팡목록에없음\) (\d+) / C\(금액불일치\) (\d+) / D\(연결제안\) (\d+) / 정상 (\d+)", r.stdout)
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
    from band_extract import parse_post, normalize_tech
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
    assert normalize_tech("권오절") == "권오철"
    assert normalize_tech("권오처르 + 1명 지원") == "권오철"
    print("  [10] 밴드 업무 추출(유형·유상무상·기사명 오탈자 정규화·템플릿/비업무 제외) ✅")


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
    # ★ 2026-07-28: 워치독이 16:27~20:43 안 돌았고 그 사이 터널 주소가 죽어 폰 접속이 끊겼다.
    #   로그에는 '정상'만 남아 있어 아무도 몰랐다. 쉰 것보다 **쉰 걸 모르는 것**이 문제다.
    from watchdog import gap_note
    from datetime import datetime as _dt
    now = _dt(2026, 7, 28, 20, 43)
    assert "256분" in gap_note("[07-28 16:27] 서버 정상", now), gap_note("[07-28 16:27] x", now)
    assert gap_note("[07-28 20:13] 서버 정상", now) == ""          # 정상 주기는 조용히
    assert gap_note("[07-28 19:59] 서버 정상", now) == ""          # 44분은 여유 안(30+15)
    assert gap_note("", now) == "" and gap_note("깨진 줄", now) == ""
    # 해가 바뀌면 직전 기록은 작년이다 — 미래로 읽어 음수 공백이 나오면 안 된다
    assert gap_note("[12-31 23:50] x", _dt(2026, 1, 1, 0, 10)) == ""
    print("  [9] 워치독 판단 로직(버전 보존 3개·보호파일·30일 기준·실행 공백 감지) ✅")


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
    # 계단·철거·신규납품·기타는 AS/점검이 아니라 별도 공사다 — 02·04에 작업 행이 아예 없으므로
    # '미상'이 아니라 **대상외**로 못 박는다(그래야 영원히 미해결로 남지 않는다).
    _, v, _ = B.bundle({**doc, "kind": "철거"}, W)
    assert v.startswith("대상외(") and "철거" in v, v
    # 밴드에 사람이 적어 둔 계산서 목록·PO 발주글이 추정보다 우선한다
    binx = {("2026-05-11", 1000000): {"프로젝트": ["UJ2600009"], "캠프": "양주2캠프", "PO": ""}}
    prjs, v, _ = B.bundle({**doc, "slip": "2026/05/11-3", "amt": 1000000}, W, binx)
    assert v == "확정(밴드)" and prjs == ["UJ2600009"], (v, prjs)
    poinx = {777000: ["UJ2600077", "UJ2600078"]}
    prjs, v, _ = B.bundle({**doc, "amt": 777000}, W, None, poinx)
    assert v == "확정(밴드PO)" and len(prjs) == 2, (v, prjs)
    # PO 발주글의 총금액 표기는 제각각이다 — '원'만 보면 KRW로 적힌 글을 통째로 놓친다
    for t, want in (("★ 총금액 : 25,223,400원", "25,223,400"),
                    ("★ 총금액 :8,626,500원", "8,626,500"),
                    ("★ 총금액 : 13,866,500 KRW", "13,866,500"),
                    ("★ 총금액 :14,803,300 KRW", "14,803,300")):
        m = B._TOTAL_RE.search(t)
        assert m and m.group(1) == want, (t, m.group(1) if m else None)
    # 프로젝트NO가 안 적힌 PO는 번호·건수라도 알려 준다(추정으로 뭉개지 않는다)
    meta = {900000: {"PO": "PO364055", "품목": "정기점검 29건", "건수": 29, "유형": "정기점검"}}
    _, v2, _ = B.bundle({**doc, "amt": 900000}, [], None, None, meta)
    assert v2.startswith("PO확인(PO364055"), v2
    # 계단·철거는 PO를 알아도 '별도 공사'라는 사실이 유지돼야 한다
    _, v3, _ = B.bundle({**doc, "kind": "계단", "amt": 900000}, [], None, None,
                        {900000: {"PO": "PO111", "품목": "계단 1EA", "건수": 1, "유형": "계단"}})
    assert v3.startswith("대상외(계단") and "PO111" in v3, v3
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
    # 값이 보이는 상태열도 수식 전용 열이면 복구한다(첨부 화면의 '미배정' 고착 회귀).
    owned = ('<sheetData>'
             '<row r="5"><c r="B5" t="inlineStr"><is><t>UJ1</t></is></c>'
             '<c r="M5" t="inlineStr"><is><t>미배정</t></is></c></row>'
             '<row r="6"><c r="B6" t="inlineStr"><is><t>UJ2</t></is></c>'
             '<c r="M6"><f>IF($B6="","",IF($L6="","미배정","배정완료"))</f><v>미배정</v></c>'
             '</row></sheetData>')
    restored, nr = F.restore_owned_formulas(owned, ("M",))
    assert nr == 1 and '<c r="M5"><f>IF($B5="","",IF($L5="","미배정","배정완료"))</f><v/></c>' in restored, restored
    assert not F.direct_self_refs(restored), F.direct_self_refs(restored)
    direct = '<row r="5"><c r="M5"><f>IF(M5="","",1)</f><v/></c></row>'
    assert F.direct_self_refs(direct) == ["M5"]
    counter = ('<row r="130"><c r="AO130"><f>'
               'IF($B130="","",COUNT($AO$4:AO153)+1)</f><v>1</v></c></row>')
    counter_out, counter_n = F.fix_cumulative_counters(counter)
    assert counter_n == 1 and 'COUNT($AO$4:AO129)' in counter_out, counter_out
    # 일반 수식열에서도 사람이 확정한 inlineStr 값은 <v>가 없다는 이유로 덮으면 안 된다.
    mixed = ('<sheetData>'
             '<row r="5"><c r="H5"><f>IF($B5="","",1)</f><v>1</v></c>'
             '<c r="I5"><f>IF($B5="","",2)</f><v>2</v></c></row>'
             '<row r="6"><c r="H6"><f>IF($B6="","",1)</f><v>1</v></c>'
             '<c r="I6"><f>IF($B6="","",2)</f><v>2</v></c></row>'
             '<row r="7"><c r="H7" t="inlineStr"><is><t>엄진언</t></is></c></row>'
             '<row r="8"><c r="H8" t="inlineStr"><is><t></t></is></c>'
             '<c r="I8"><f>IF($B8="","",2)</f><v>2</v></c></row>'
             '</sheetData>')
    mixed_out, mixed_n = F.fill_missing(mixed)
    assert mixed_n == 2, mixed_n
    assert '<c r="H7" t="inlineStr"><is><t>엄진언</t></is></c>' in mixed_out, mixed_out
    assert '<c r="H8"><f>IF($B8="","",1)</f><v/></c>' in mixed_out, mixed_out
    assert '<c r="I7"><f>IF($B7="","",2)</f><v/></c>' in mixed_out, mixed_out
    assert "00_대시보드" in F.FILL_MISSING_EXCLUDE
    print("  [23] 수식 복구(수기값 보존·빈칸 판정·대시보드 제외·순환참조·범위확장) ✅")


def t24_reorder_safety():
    """행 재배치가 **엑셀이 못 여는 파일**을 만들지 않는지.
    2026-07-26 v101~v104가 이 두 가지로 열리지 않았다. 구조 검사로는 안 잡힌다 —
    XML은 멀쩡하고 '의미'만 어긋나기 때문에 여기서 못 박아 둔다."""
    import reorder_rows as R
    # (1) 배열수식 ref 는 자기 셀 주소다 — 행을 옮기면 함께 가야 한다
    row = ('<row r="69"><c r="Q69" s="53"><f t="array" ref="Q69" aca="1">SUM($A69:$B69)</f>'
           '<v>1</v></c></row>')
    moved = R.move_row(row, 5)
    assert 'ref="Q5"' in moved, moved
    assert '<row r="5"' in moved and '<c r="Q5"' in moved, moved
    assert 'SUM($A5:$B5)' in moved, moved
    # 범위형 ref 도 양끝이 따라간다
    row2 = '<row r="10"><c r="A10"><f t="array" ref="A10:B10">1</f><v>1</v></c></row>'
    assert 'ref="A12:B12"' in R.move_row(row2, 12), R.move_row(row2, 12)
    # (2) 공유수식은 '연속한 행' 전제라 섞기 전에 반드시 펼쳐야 한다
    xml = ('<sheetData>'
           '<row r="5"><c r="B5"><f t="shared" ref="B5:B7" si="3">IF($A5="","",1)</f><v>1</v></c></row>'
           '<row r="6"><c r="B6"><f t="shared" si="3"/><v>1</v></c></row>'
           '<row r="7"><c r="B7"><f t="shared" si="3"/><v>1</v></c></row>'
           '</sheetData>')
    out, n = R.unshare(xml)
    assert n == 3, n
    assert 't="shared"' not in out and 'si="3"' not in out, out
    assert 'IF($A6="","",1)' in out and 'IF($A7="","",1)' in out, out   # 각 행 수식이 생겼다
    assert '<c r="B6"' in out and '<v>1</v>' in out, out                # 셀·캐시는 그대로
    print("  [24] 재배치 안전성(배열수식 ref 이동·공유수식 펼침) ✅")


def t25_attachments():
    """행에 **붙어 있는 것들**이 함께 움직이는가 / 범위 확장이 시작행을 밀지 않는가.
    둘 다 2026-07-26 합성 점검에서 찾은 구멍이다."""
    import reorder_rows as R, expand_rows as E
    # (1) 하이퍼링크는 sheetData 밖에 적혀 있어 행만 정렬하면 제자리에 남는다 →
    #     9행이 다른 건으로 바뀌었는데 링크는 원래 건의 밴드 글을 가리킨다.
    post = '<hyperlinks><hyperlink ref="AM9" r:id="rId1"/><hyperlink ref="AM12" r:id="rId2"/></hyperlinks>'
    out, n = R.move_hyperlinks(post, {9: 300, 12: 7})
    assert n == 2 and 'ref="AM300"' in out and 'ref="AM7"' in out, out
    out2, n2 = R.move_hyperlinks(post, {})          # 옮길 행이 없으면 그대로
    assert n2 == 0 and out2 == post
    # (2) 범위 확장은 **끝행만** 늘려야 한다. 시작행까지 밀면 구간이 통째로 벗어난다.
    assert E.bump("A5:A5", 5, 8) == "A5:A8", E.bump("A5:A5", 5, 8)
    assert E.bump("V5:V31 V33:V474", 474, 500) == "V5:V31 V33:V500"
    assert E.bump("V31", 31, 60) == "V31", "단일 셀은 범위가 아니다"
    assert E.bump("A5:AG104", 104, 164) == "A5:AG164"
    print("  [25] 행 부속물(하이퍼링크 이동·범위 끝행만 확장) ✅")


def t26_mobile():
    """폰에서 여는 경로가 **끊기지 않는 구조인지** 합성으로 확인한다.

    실제로 반복해서 겪은 고장은 전부 '주소가 어딘가에 박혀 있어서'였다:
      · 폰 홈 화면 아이콘  → 추가할 때의 터널 주소가 매니페스트 start_url에 박힘
      · PC 앱(크롬 PWA)   → 설치할 때의 터널 주소가 app-id에 박힘
    터널 주소는 띄울 때마다 새로 받으므로 다음 날이면 전부 죽는다.
    그래서 **바깥으로 내보내는 주소는 언제나 고정 진입점**이어야 한다.
    """
    import json as _j, re as _re
    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    import app_server as A

    FIX = A.FIXED_ENTRY
    assert FIX.startswith("https://"), FIX
    assert "trycloudflare" not in FIX, "고정 진입점에 터널 주소를 넣으면 안 된다"

    # (1) 앱이 내주는 매니페스트: 홈 화면 아이콘이 고정 주소를 향하는가
    src = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    i = src.index('if p == "/manifest.json"')
    blk = src[i:i + 900]
    # ★ 매니페스트 start_url은 **그 페이지와 같은 출처**여야 한다.
    #   다른 도메인을 넣으면 크롬이 매니페스트를 무시해 [설치 및 바로가기 만들기]가 아예 안 된다
    #   (2026-07-27에 고정 주소를 넣었다가 폰에서 설치가 막혔다).
    assert '"start_url": "/"' in blk, "앱 매니페스트 start_url은 같은 출처(/)여야 설치가 된다"
    assert "FIXED_ENTRY" not in blk, "다른 도메인을 start_url에 넣으면 크롬이 설치를 거부한다"
    # 대신 앱 화면이 **고정 주소를 항상 보여줘야** 한다(닫히는 배너가 아니라 상시 표기).
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert FIX in idx, "앱이 고정 주소를 표기하지 않는다"
    assert "showFixedEntry" in idx and 'id="entrycard"' in idx, "고정 주소 상시 표기 카드가 없다"
    assert "tunnelNotice" not in idx, "닫히는 안내 배너는 상시 표기로 대체됐다"
    # 카드가 숨겨져 있으면 표기한 의미가 없다
    _card = idx[idx.index('id="entrycard"'):idx.index('id="entrycard"') + 60]
    assert "display:none" not in _card, "고정 주소 카드가 숨겨져 있다"

    # (1-b) 크롬 [설치 및 바로가기 만들기]는 fetch 핸들러를 가진 서비스 워커가 있어야 뜬다.
    #       고정 주소 쪽과 앱(터널) 쪽 **양쪽 출처 모두** 필요하다 — 설치는 출처 단위다.
    assert "serviceWorker" in idx and "/sw.js" in idx, "앱에 서비스 워커 등록이 없다 — 설치가 안 된다"
    assert 'p == "/sw.js"' in src, "앱 서버가 /sw.js 를 내주지 않는다"
    _swblk = src[src.index('p == "/sw.js"'):][:700]
    assert "addEventListener('fetch'" in _swblk, "fetch 핸들러가 없으면 크롬이 설치를 제안하지 않는다"
    _dsw = open(os.path.join(ROOT, "docs", "sw.js"), encoding="utf-8").read()
    assert "addEventListener('fetch'" in _dsw, "고정 주소 쪽 서비스 워커에 fetch 핸들러가 없다"

    # (2) 고정 진입점 페이지: 앱 시작 자체가 PC 터널과 무관한가
    doc = open(os.path.join(ROOT, "docs", "index.html"), encoding="utf-8").read()
    assert 'rel="manifest"' in doc, "고정 페이지에 매니페스트가 없으면 홈 화면 추가가 앱으로 안 붙는다"
    assert "serviceWorker" in doc and "sw.js" in doc, "고정 페이지에 서비스 워커 등록이 없다 — 설치가 안 된다"
    assert "location.replace('app.html')" in doc, "고정 주소가 PC 독립 앱으로 바로 들어가지 않는다"
    assert "endpoint.json" not in doc and "/api/ping" not in doc, \
        "앱 시작 전에 PC 터널을 확인하면 PC 종료 시 다시 막힌다"

    mf = _j.load(open(os.path.join(ROOT, "docs", "manifest.json"), encoding="utf-8"))
    assert mf["start_url"] == FIX + "app.html", mf
    assert mf.get("scope", FIX).startswith(FIX.rstrip("/")), mf
    assert "trycloudflare" not in _j.dumps(mf)

    # (3) 터널 주소는 어디에도 하드코딩돼 있으면 안 된다(리포트·설정 파일은 예외)
    hard = []
    for fn in ("webapp/app_server.py", "webapp/index.html", "docs/index.html",
               "phone_access.py", "publish_endpoint.py"):
        txt = open(os.path.join(ROOT, fn), encoding="utf-8").read()
        for m in _re.finditer(r"https://[a-z0-9-]+\.trycloudflare\.com", txt):
            hard.append(f"{fn}: {m.group()}")
    assert not hard, "터널 주소가 코드에 박혀 있다 — 주소가 바뀌면 그대로 죽는다: " + "; ".join(hard[:3])

    # (4) 자가복구가 붙어 있는가 — 죽은 주소를 붙들고 있으면 아무도 못 고친다
    import webapp.tunnel_run as T  # noqa: F401
    tr = open(os.path.join(ROOT, "webapp", "tunnel_run.py"), encoding="utf-8").read()
    assert "def watch(" in tr and "def publish(" in tr, "터널 자가점검·주소게시가 빠졌다"

    # ★ cloudflared 로그에는 api.trycloudflare.com(내부 API)도 나온다.
    #   그걸 터널 주소로 잘못 잡으면 고정 주소가 엉뚱한 곳을 가리켜 폰이 통째로 막힌다.
    import re as _re2
    m = _re2.search(r"m = re\.search\(r\"\((https[^\"]+)\)\"", tr)
    assert m, "터널 주소 정규식을 못 찾음"
    rx = _re2.compile(m.group(1))
    assert not rx.search("https://api.trycloudflare.com"), "api.trycloudflare.com을 걸러내지 못한다"
    assert rx.search("INF | https://mold-restored-flags-earthquake.trycloudflare.com |"), "정상 주소를 못 잡는다"

    # 주소를 너무 자주 새로 만들면 등록조차 안 되는 주소를 받는다 — 한도가 있어야 한다
    assert "_throttle" in tr and "MAX_PER_HOUR" in tr, "터널 재생성 한도가 없다"

    # 게시 직전에 형식·생존을 한 번 더 본다(죽은 주소를 올리면 폰이 막힌다)
    pe = open(os.path.join(ROOT, "publish_endpoint.py"), encoding="utf-8").read()
    assert "게시 취소" in pe and "/api/ping" in pe, "죽은 주소를 그대로 게시할 수 있다"
    wd = open(os.path.join(ROOT, "watchdog.py"), encoding="utf-8").read()
    assert "kill_stale_tunnel" in wd, "워치독이 좀비를 정리하지 않으면 재시작이 무효가 된다"
    # ★★ 터널 주소에서는 **설치가 되면 안 된다**. 터널 호스트는 매번 바뀌는데 거기서
    #   설치하면 그 임시 주소가 아이콘에 영구히 박혀, 주소가 바뀌는 순간 영영 안 열린다
    #   (2026-07-28: PC·폰의 설치된 앱이 둘 다 옛 터널 주소로 죽었다).
    i = src.index('if p in ("/", "/index.html")')
    blk2 = src[i:i + 1200]
    assert "trycloudflare.com" in blk2 and 'rel="manifest"' in blk2,         "터널 출처에서 매니페스트를 빼지 않는다 — 거기서 설치하면 아이콘이 곧 죽는다"
    assert "serviceWorker.register" in blk2, "터널 출처에서 서비스워커도 빼야 설치가 안 뜬다"
    print("  [26] 모바일 접속 경로(고정 진입점·터널출처 설치차단·자가복구) ✅")


def t27_po():
    """쿠팡 PO 대조 — **계산서가 끊긴 PO를 '누락'이라 부르지 않는다**.

    쿠팡 PO는 '정기점검 24건'처럼 여러 작업을 묶은 것이라 06시트 PO번호 칸에 안 적힌다.
    그걸 전부 '원장 미등록'으로 세면 96건 중 94건이 쏟아져 나와, 무엇이 진짜 미청구인지
    묻혀 버린다(2026-07-27 실제로 그랬다). ERP 계산서 금액과 맞춰 본 뒤,
    **계산서가 없는 PO만** 미청구로 올린다.
    """
    import po_reconcile as P, inspect
    src = inspect.getsource(P.main)
    assert "erp_amts" in src and "invoiced(" in src, "ERP 계산서 대조 없이 미등록으로 단정한다"
    assert "미청구 — 쿠팡 PO는 받았는데" in src, "판정 문구가 조치로 이어지지 않는다"
    assert "billed" in src, "계산서 발행된 PO를 따로 세지 않는다"
    # 밴드에 '세금계산서 발행 완료'라고 적힌 PO를 미청구로 보고하면 안 된다
    # (오종현 매니저 확인: PO 내역은 밴드 매출처업무에 1월부터 전부 기재한다)
    assert "band_po" in src and "발행완료" in src, "밴드에 적힌 PO 처리 상태를 안 본다"

    import po_band_status as PB
    body = "\n".join([
        "Coupang이(가) 새 구매 오더(PO344599)를 전송했습니다.",
        "⭐ 세금계산서 발행 2건 발행",
        "★오더번호 : PO344599",
        "2026.03.25 세금계산서 발행 완료",
        "★ 프로젝트 No. : UJ2600211",
    ])
    got = None
    for blk in __import__("re").split(r"(?=Coupang이\(가\) 새 구매 오더|⭐|✅)", body):
        if "PO344599" not in blk:
            continue
        st = [n for n, rx in PB.STATES if rx.search(blk)]
        if "발행완료" in st:
            got = st
    assert got and "발행완료" in got, ("발행 완료 문구를 못 읽는다", got)
    st2 = [n for n, rx in PB.STATES if rx.search("✅ 쿠팡오더처리 + 세금계산서 발행대기")]
    assert "발행대기" in st2, st2
    st3 = [n for n, rx in PB.STATES if rx.search("※쿠팡오더 금액 안맞아도 처리해도 된다고 확인 받음")]
    assert "금액이상" in st3, st3

    # PO 번호 정규화 — 표기가 흔들려도 같은 PO로 본다
    for t_, want in (("PO326234", "PO326234"), ("po 326234", "PO326234"),
                     ("PO-326234", "PO326234"), ("발주 PO326234 건", "PO326234")):
        assert P.norm_po(t_) == want, (t_, P.norm_po(t_))
    assert P.norm_po("PROJECT") == "" and P.norm_po("") == ""

    # 총금액이 부가세 포함으로 적힌 PO도 공급가 기준 계산서와 맞아야 한다
    erp = {1000000, 1100000}
    def invoiced(a):
        return int(a) in erp or round(int(a) / 1.1) in erp
    assert invoiced(1100000) and invoiced(1000000)
    assert not invoiced(999999)
    print("  [27] 쿠팡 PO 대조(계산서 발행분 제외·번호 정규화·부가세 표기) ✅")


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
        # 서비스 워커는 **진짜 자바스크립트**로 나가야 한다.
        # _send 가 str을 JSON으로 감싸는 바람에 "self.add..." 처럼 따옴표에 싸여 나가면
        # 브라우저가 등록을 거부하고 [설치 및 바로가기 만들기]가 그대로 먹통이 된다.
        _r = urllib.request.urlopen(base + "/sw.js", timeout=5)
        _sw = _r.read().decode("utf-8")
        assert "javascript" in _r.headers.get("Content-Type", ""), _r.headers.get("Content-Type")
        assert _sw.startswith("self.addEventListener"), "sw.js가 JSON으로 감싸여 나간다: " + _sw[:40]
        assert "addEventListener('fetch'" in _sw, "fetch 핸들러 없이는 설치가 제안되지 않는다"
        print("  [6] 웹앱 API(PIN 인증·상태·정산·페이지 서빙) ✅")
    finally:
        p.terminate()


def t28_resolve():
    """[28] 프로젝트 코드 하나로 나머지가 따라오는가.
    이 알고리즘의 목숨은 **채번이 시트 수식과 똑같은가**에 달렸다. 어긋나면 사람이 엑셀을
    연 순간 다시 계산돼 ID가 바뀌고, 03·06이 값으로 들고 있는 참조가 전부 끊긴다."""
    import project_resolve as P
    from datetime import datetime

    # (1) 코드 정규화 — 사람이 어떻게 쳐도 같은 코드로 모여야 한다
    for raw in ("UJ2601138", "uj2601138", " UJ 2601138 ", "2601138", "UJ-2601138"):
        assert P.norm(raw) == "UJ2601138", raw
    for bad in ("", None, "UJ260113", "UJ26011389", "AS-2607-001", "UJ260113X"):
        assert P.norm(bad) is None, bad

    # (2) 채번이 시트 수식과 같은가 — 수식: 접두어 & TEXT(날짜,"yymm") & TEXT(ROW()-4,"000")
    assert P.mint("AS", "2026-06-03", 531) == "AS-2606-527", P.mint("AS", "2026-06-03", 531)
    assert P.mint("PM", "2025-12-31", 9) == "PM-2512-005"
    assert P.mint("JS", "2026-07-01", 1004) == "JS-2607-1000", "네 자리도 잘려선 안 된다"
    assert P.mint("AS", "", 100) == "AS-%s-096" % datetime.now().strftime("%y%m")
    assert P.mint("AS", None, 100).endswith("-096")

    # (3) 실제 원장 수식과 대조 — 규칙을 코드에만 적어 두면 시트가 바뀌어도 모른다
    import openpyxl, re as _re
    from ecount_reconcile import load_config, resolve_master
    m = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(m, data_only=False)
    for sh, (idc, pfx, _dc, _kc) in P.SHEET_ID.items():
        if sh not in wb.sheetnames:
            continue
        ws = wb[sh]
        f = next((ws.cell(row=r, column=1).value for r in range(5, ws.max_row + 1)
                  if isinstance(ws.cell(row=r, column=1).value, str)
                  and str(ws.cell(row=r, column=1).value).startswith("=")), None)
        if not f:
            continue
        got = _re.search(r'"(\w{2})-"', f)
        assert got, "%s ID 수식에서 접두어를 못 찾음" % sh
        assert got.group(1) == pfx, (
            "%s: 시트 수식 접두어 '%s' != 코드 '%s' — 새 행 ID가 엉뚱한 이름으로 매겨진다"
            % (sh, got.group(1), pfx))
        assert "ROW()-4" in f.replace(" ", ""), "%s 채번이 ROW()-4가 아니다" % sh

        # ★ 행을 늘릴 때 새 행에도 같은 채번 수식이 심겨야 한다. expand_rows 의 표에서
        #   빠진 시트는 늘려도 ID가 영영 빈칸이다(03_현장작업실적이 실제로 그랬다 —
        #   120행을 늘려 놓고도 '여유 0'인 채였다, 2026-07-27).
        import expand_rows as E
        assert sh in E.ID_FORMULA, f"{sh} 가 expand_rows.ID_FORMULA 에 없다 — 늘려도 채번이 안 된다"
        _col, _name, _pfx, _dcol = E.ID_FORMULA[sh]
        assert _pfx == pfx, f"{sh} expand_rows 접두어 {_pfx} != 시트 {pfx}"
        assert _name == idc, f"{sh} expand_rows ID열 이름 {_name} != {idc}"
        # 날짜열도 실제 수식과 같아야 한다(02·04는 D, 03·05는 E, 06은 F로 서로 다르다)
        _m = _re.search(r'TEXT\(IF\(\$([A-Z]+)\d+="",TODAY\(\)', f.replace(" ", ""))
        assert _m and _m.group(1) == _dcol, (
            f"{sh}: expand_rows 날짜열 {_dcol} != 시트 수식 {_m.group(1) if _m else '?'}")
    wb.close()

    # (4) 리졸브 — 이미 등록된 코드는 새로 만들지 않고 그 행을 가리켜야 한다
    ev = P.evidence(m)
    known = next(iter(ev["ledger"]))
    r = P.resolve(known, ev)
    assert r["ok"] and r["state"] == "등록됨" and r.get("row"), r
    assert r["sheet"] in P.SHEET_ID, r["sheet"]

    # 신규 코드는 **빈 행**을 받아야 한다 — 기존 행을 덮으면 실데이터가 날아간다
    fake = "UJ9999999"
    ev2 = dict(ev)
    ev2["band"] = dict(ev["band"])
    ev2["band"][fake] = {"camp": "합성캠프", "kind": "돌발AS", "cost": "유상",
                         "tech": "홍길동", "date": "2026-06-03", "status": "작업완료",
                         "posted": "2026-06-03", "_score": 9}
    n = P.resolve(fake, ev2)
    assert n["ok"] and n["state"] == "신규", n
    assert n["row"] > ev["tail"]["02_돌발AS접수"], "새 행이 기존 데이터 위에 얹힌다"
    assert n["row"] <= ev["cap"]["02_돌발AS접수"], "수식 없는 행에 값만 쓰면 채번이 죽는다"
    assert n["ids"]["접수ID"] == P.mint("AS", "2026-06-03", n["row"])

    # 쓰기 항목은 좌표 지정 + '빈 칸만'이어야 한다(키 조회로는 아직 없는 행을 못 찾는다)
    items = P.row_items(n, ev2)
    assert items and all(i.get("only_if_empty") and i.get("cell") for i in items), items[:1]
    assert any(i["col"] == "프로젝트NO" and i["value"] == fake for i in items)
    assert all(str(i["cell"]).endswith(str(n["row"])) for i in items), "행 번호가 섞였다"
    # ID 열은 절대 쓰지 않는다 — 수식이 알아서 채운다
    assert not any(i["col"] in ("접수ID", "점검ID", "작업ID", "정산ID") for i in items)

    # 업무유형을 모르면 **찍지 말고 멈춰야** 한다
    ev2["band"][fake] = dict(ev2["band"][fake], kind="기타")
    assert not P.resolve(fake, ev2)["ok"], "시트를 모르면서 등록하면 엉뚱한 시트에 들어간다"
    print("  [28] 프로젝트코드 리졸브(정규화·채번=수식·빈행 배치·ID열 보호) ✅")


def t29_cloud():
    """[29] PC가 꺼져도 폰이 열 수 있는가 — 잠금·오프라인 폴백·예약 반영."""
    import csos_crypto as C, base64, hmac as _h, hashlib as _hl, zlib as _z, json as _j
    assert C.self_test(), "AES 공식 시험벡터 실패 — 폰이 절대 못 연다"

    raw = _j.dumps({"codes": {"UJ2601138": {"camp": "합성캠프"}}}, ensure_ascii=False).encode()
    packed = _z.compress(raw, 9)
    s = C.seal(packed, "0000", iters=1000)
    assert s["cipher"] == "AES-256-CBC" and s["kdf"] == "PBKDF2-SHA256", s
    k = C.derive("0000", base64.b64decode(s["salt"]), 1000)
    tag = _h.new(k[32:], base64.b64decode(s["iv"]) + base64.b64decode(s["ct"]), _hl.sha256).digest()
    assert _h.compare_digest(tag, base64.b64decode(s["tag"])), "무결성 태그 불일치"
    # 틀린 PIN은 **복호 전에** 걸러져야 한다(엉뚱한 데이터를 그리면 안 된다)
    kbad = C.derive("9999", base64.b64decode(s["salt"]), 1000)   # 일부러 다른 PIN
    bad = _h.new(kbad[32:], base64.b64decode(s["iv"]) + base64.b64decode(s["ct"]), _hl.sha256).digest()
    assert not _h.compare_digest(bad, base64.b64decode(s["tag"]))
    assert C.ITERS >= 300000, "반복이 낮으면 4자리 PIN이 순식간에 뚫린다"
    assert _z.decompress(packed) == raw

    # 올라간 사본이 **평문이 아니어야** 한다 — 공개 저장소다
    enc = os.path.join(ROOT, "docs", "data.enc")
    if os.path.exists(enc):
        d = _j.load(open(enc, encoding="utf-8"))
        assert set(("salt", "iv", "ct", "tag")) <= set(d), list(d)
        blob = open(enc, encoding="utf-8").read()
        for leak in ("캠프", "UJ25", "UJ26", "프로젝트NO"):
            assert leak not in blob, "사본에 평문 '%s' 이 그대로 있다" % leak

    # 고정 페이지: PC가 죽었을 때 **막다른 오류가 아니라** 오프라인 앱으로 가야 한다
    doc = open(os.path.join(ROOT, "docs", "index.html"), encoding="utf-8").read()
    assert "app.html" in doc and "location.replace('app.html')" in doc, \
        "고정 주소가 PC 독립 앱으로 바로 들어가지 않는다"
    app = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    for need in ("PBKDF2", "AES-CBC", "DecompressionStream", "csos_queue", "/api/queue"):
        assert need in app, "오프라인 앱에 %s 누락" % need
    assert "Authorization" in app and "Bearer " in app, \
        "클라우드 큐 인증 헤더가 없어 예약을 안전하게 보낼 수 없다"
    assert "클라우드에 보관되며 PC가 켜지면 자동 반영" in app, \
        "PC 독립 예약의 실제 동작 안내가 빠졌다"
    sw = open(os.path.join(ROOT, "docs", "sw.js"), encoding="utf-8").read()
    assert "endpoint.json" in sw and "data.enc" in sw, "사본을 쥐거나 주소를 캐시하는 규칙이 없다"
    assert "addEventListener('fetch'" in sw

    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert '"/api/enqueue"' in srv and "def enqueue_codes(" in srv, "폰 예약을 받을 곳이 없다"

    # git add부터 실패를 숨기면 실제 게시가 안 됐는데도 '반영했습니다'가 찍힌다.
    import cloud_publish as CP
    assert "docs/resolve_index.json" not in CP.PUBLISH_FILES, "존재하지 않는 파일을 git add 한다"
    assert "docs/manifest.json" in CP.PUBLISH_FILES, \
        "PC 독립 시작 주소를 바꿔도 매니페스트가 게시되지 않는다"
    class _R:
        def __init__(self, code=0, out="", err=""):
            self.returncode, self.stdout, self.stderr = code, out, err
    calls = []
    def fail_add(cmd, **_kwargs):
        calls.append(cmd)
        return _R(1, err="pathspec missing") if cmd[1] == "add" else _R()
    ok, stage, detail = CP.git_publish("합성", runner=fail_add)
    assert not ok and stage == "add" and len(calls) == 1 and "pathspec" in detail, (ok, stage, calls)
    calls.clear()
    def no_changes(cmd, **_kwargs):
        calls.append(cmd)
        return _R(1, out="nothing to commit") if cmd[1] == "commit" else _R()
    ok, stage, _ = CP.git_publish("합성", runner=no_changes)
    assert ok and stage == "" and [c[1] for c in calls] == ["add", "commit", "push"], calls

    # ★ 실 PIN이 소스 어딘가에 박히면 공개 저장소에 그대로 남아 사본 잠금이 무의미해진다.
    #   (실제로 자체검증 코드에 리터럴로 들어갔던 적이 있다 — 2026-07-27)
    try:
        real = str(_j.load(open(os.path.join(ROOT, "config", "webapp.json"),
                                encoding="utf-8"))["pin"])
    except Exception:
        real = ""
    if real:
        import subprocess as _sp
        r = _sp.run(["git", "grep", "-l", "--cached", real], cwd=ROOT,
                    capture_output=True, text=True, encoding="utf-8", errors="replace")
        hits = [ln for ln in (r.stdout or "").splitlines() if ln.strip()]
        assert not hits, "커밋된 파일에 실 PIN이 들어 있다: " + ", ".join(hits[:3])
    print("  [29] 폰 단독 사용(잠금·오프라인 폴백·git 게시 실패감지·PIN 비노출) ✅")


def t30_dns_and_versions():
    """[30] 오래 끌던 접속 불가의 진짜 원인과, 버전 파일이 쌓이는 문제.

    사내 DNS가 *.trycloudflare.com 을 안 풀어 준다. 그래서 PC에서 터널을 찔러 보면
    늘 실패했고, publish는 게시를 취소하고 watch는 멀쩡한 터널을 계속 갈아치웠다.
    폰에서는 살아 있던 주소인데 PC 혼자 못 보고 있었다."""
    import net_probe as N

    # 공개 DNS 우회가 실제로 동작하는가(회사 DNS가 막아도 이름을 얻어야 한다)
    ips = N.resolve_public("www.cloudflare.com")
    assert ips and all(re.fullmatch(r"[\d.]+", i) for i in ips), ips

    # 판정이 정직한가 — 없는 이름은 살아있다고 하면 안 된다
    ok, why = N.probe("https://this-name-does-not-exist-csos.trycloudflare.com/api/ping", timeout=6)
    assert not ok and "풀지 못함" in why, why

    # 게시·감시가 **회사 DNS에 속지 않는 경로**를 쓰는가
    pe = open(os.path.join(ROOT, "publish_endpoint.py"), encoding="utf-8").read()
    assert "from net_probe import probe" in pe, "게시가 여전히 회사 DNS로 판단한다"
    tr = open(os.path.join(ROOT, "webapp", "tunnel_run.py"), encoding="utf-8").read()
    assert tr.count("from net_probe import probe") >= 2, "감시 루프가 멀쩡한 터널을 죽었다고 본다"
    # 터널 대상은 127.0.0.1이어야 한다(localhost면 IPv6로 풀려 HTTP 530)
    assert 'f"http://127.0.0.1:{PORT}"' in tr, "localhost로 주면 IPv6(::1)로 풀려 앱에 못 닿는다"
    assert "localhost:{PORT}" not in tr, "localhost가 남아 있다"
    # 서비스가 죽어 있으면 생성 한도가 복구를 막으면 안 된다
    assert "_endpoint_alive" in tr and "HARD_MAX" in tr, "한도가 복구까지 막는다"

    # 버전 정리: 최신본은 어떤 경우에도 남아야 한다
    import ledger_versions as V
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    keep, move = V.plan(master)
    kept = {k["path"] for k in keep}
    assert master in kept, "최신본을 접으려 한다 — 모든 도구가 멈춘다"
    assert len(keep) >= min(V.KEEP_LATEST, len(keep) + len(move)), (len(keep), len(move))
    assert not (set(m["path"] for m in move) & kept), "같은 파일이 남김·접기 양쪽에 있다"
    # 지우지 않는다(옮기기만 한다)
    src = open(os.path.join(ROOT, "ledger_versions.py"), encoding="utf-8").read()
    assert "shutil.move" in src and "삭제하지 않는다" in src
    print("  [30] 사내 DNS 우회 판정 · 터널 대상 IPv4 · 버전 파일 정리(최신본 보호) ✅")


def t31_tech():
    """[31] 누가 어디를 다녀왔는가 — 세는 기준이 흔들리면 실적이 부풀거나 빠진다."""
    import tech_report as T

    # 동행 건은 양쪽 다 센다. 이름이 아닌 조각은 버린다.
    assert T.split_tech("김준형, 김필우") == ["김준형", "김필우"]
    assert T.split_tech("문상국. 최일파") == ["문상국", "최일파"]
    assert T.split_tech("000 (캠프상태확인 및 스케쥴 세팅)") == []
    assert T.split_tech("") == [] and T.split_tech(None) == []

    visits, pending, unknown, _m = T.collect()
    assert visits, "방문 기록이 하나도 안 잡힌다"
    # **완료된 것만** 방문으로 센다 — 예정만 있는 건은 pending 이어야 한다
    assert all(v["방문일"] for v in visits), "완료일 없는 건이 방문으로 세어졌다"
    assert all(not p["방문일"] for p in pending), "완료된 건이 미방문으로 빠졌다"
    assert all(u["방문일"] and not u["기사"] for u in unknown), unknown[:1]
    # 한 사람씩 펼쳐 담는다(집계할 때 다시 쪼갤 필요가 없게)
    assert all("," not in v["기사"] for v in visits), "동행 건이 안 펼쳐졌다"

    by = T.summary(visits)
    assert sum(d["총"] for d in by.values()) == len(visits)
    for t, d in by.items():
        assert d["돌발AS"] + d["정기점검"] == d["총"], (t, d)
        assert d["최근"] and re.fullmatch(r"\d{4}-\d{2}-\d{2}", d["최근"]), (t, d["최근"])
    print("  [31] 기사별 방문(동행 분리·완료분만·합계 정합) ✅")


def t32_band_sheet():
    """[32] 24_밴드업무추출이 통째로 잘리는 사고를 막는다.

    band_extract 의 --sheet 는 시트를 **덮어쓴다**. --month 와 같이 쓰면 그 달 것만 남는다.
    ingest 가 월별 루프로 이걸 반복해서 2026-07-27에 4,223행 → 506행으로 잘렸다
    (2025-12~2026-05가 통째로 사라졌다). 두 겹으로 막는다."""
    # ① ingest 는 --sheet 를 월별로 부르면 안 된다
    ing = open(os.path.join(ROOT, "band", "ingest.py"), encoding="utf-8").read()
    i = ing.index('if "--sheet" in args:')
    blk = ing[i:i + 700]
    assert '"--month", mo, "--sheet"' not in blk, (
        "ingest 가 24시트를 월별로 덮어쓴다 — 마지막 달만 남고 나머지가 사라진다")
    assert '"band_extract.py"), "--sheet"' in blk.replace("os.path.join(ROOT, ", ""), blk[:120]

    # ② 시트가 캐시보다 훨씬 적으면 잘린 것이다
    import json as _j, glob as _g
    posts = 0
    for f in _g.glob(os.path.join(ROOT, "band", "cache", "*.json")):
        if os.path.basename(f).startswith(("dump_", "raw_")):
            continue
        try:
            posts += len((_j.load(open(f, encoding="utf-8")).get("posts") or {}))
        except Exception:
            pass
    if posts:
        import openpyxl
        from ecount_reconcile import load_config, resolve_master
        wb = openpyxl.load_workbook(resolve_master(load_config()["reconcile"]["master_xlsx"]),
                                    read_only=True, data_only=True)
        ws = wb["24_밴드업무추출"]
        hdr = [str(h).strip() if h else "" for h in
               next(ws.iter_rows(min_row=4, max_row=4, values_only=True))]
        ig = hdr.index("게시일")
        days = sorted(str(r[ig])[:10] for r in ws.iter_rows(min_row=5, values_only=True)
                      if ig < len(r) and r[ig])
        wb.close()
        assert days, "24시트가 비어 있다"
        # 캐시는 1,200여 글인데 시트가 그 절반도 안 되면 잘린 것으로 본다
        assert len(days) > posts * 0.5, (
            f"24시트 {len(days)}행 vs 밴드 캐시 {posts}글 — 시트가 잘린 것으로 보인다")
        # 밴드 시작(2025-12)부터 있어야 한다 — 최근 몇 달만 남아 있으면 덮어써진 것이다
        assert days[0] <= "2026-01-31", (
            f"24시트가 {days[0]} 부터 시작한다 — 앞부분이 덮어써졌다")
    print("  [32] 24시트 전체 보존(월별 덮어쓰기 차단·행수 급감 감지) ✅")


def t33_unbilled_banner():
    """[33] 미청구가 눈에 띄는가 — 앱을 열면 화면 맨 위에 뜨고, 경과일은 오늘 기준이어야 한다.

    ★ 예전에는 GitHub Actions가 매일 세어 **메일**을 보냈다(공개 파일 docs/aging.json +
      워크플로 실패 트릭). 사용자가 메일을 원하지 않아 그 경로를 전부 걷어내고, 앱을 열 때
      화면에서 보이게 옮겼다. 지워진 것이 되살아나지 않는지도 함께 본다."""
    # (1) 메일 경로가 완전히 사라졌는가 — 남아 있으면 원치 않는 메일이 계속 간다
    assert not os.path.exists(os.path.join(ROOT, ".github", "workflows", "aging-alert.yml")), \
        "메일을 보내는 워크플로가 아직 있다"
    assert not os.path.exists(os.path.join(ROOT, "docs", "aging.json")), \
        "공개 aging.json 이 아직 있다(이제 쓰지 않는다)"
    assert not os.path.exists(os.path.join(ROOT, "tools", "aging_check.py")), \
        "메일용 판정 스크립트가 아직 있다"
    cp = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    # 설명 주석에 '예전에는 aging.json 을 썼다'는 이력은 남아도 된다 — **만드는 코드**만 본다
    assert "def aging(" not in cp, "cloud_publish 에 공개 aging.json 생성 함수가 남아 있다"
    _code = [ln for ln in cp.splitlines()
             if "aging.json" in ln and not ln.strip().startswith(("#", "*", "·"))
             and "예전" not in ln and "docs/aging.json)" not in ln]
    assert not _code, "cloud_publish 가 아직 aging.json 을 다룬다: " + str(_code[:2])

    # (2) 미청구가 **잠긴 사본 안에** 들어가는가(공개 파일이 아니라)
    assert "def unbilled(" in cp and 'd["unbilled"]' in cp, "미청구가 사본에 안 담긴다"
    import cloud_publish as CP
    rows = CP.unbilled()
    for r in rows:
        assert re.fullmatch(r"\d{4}-\d{2}-\d{2}", r["발행일"]), r
        assert isinstance(r["금액"], int), r
    assert rows == sorted(rows, key=lambda x: x["발행일"]), "발행일 순이 아니다"

    # (3) 앱이 **오늘 날짜로** 경과일을 계산하는가 — 사본이 묵어도 일수는 정확해야 한다
    app = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    assert "renderAlert" in app and 'id="alertbox"' in app, "미청구 배너가 없다"
    assert "renderAlert();" in app, "앱을 열 때 배너를 그리지 않는다"
    assert "Date.now()" in app.split("const DAYS")[1][:200], \
        "경과일을 사본 생성시각으로 계산하면 며칠 묵은 사본에서 숫자가 틀어진다"
    assert "D.unbilled" in app, "배너가 사본의 미청구를 읽지 않는다"
    # 90일/120일 단계 표시
    assert "r.d >= 120" in app and "r.d >= 90" in app, "경과 단계 표시가 없다"
    print("  [33] 미청구 배너(메일 경로 제거·사본 내 보관·오늘 기준 경과일) ✅")


def t34_capture_and_no_send():
    """[34] 확인 목록 캡처 기능 + **쿠팡에 자동 전송 경로가 없는지**.

    사용자 상시 지시(2026-07-27): 쿠팡 담당자에게는 어떤 메시지도 보내지 않는다.
    유니버셜 내부에서 처리한다. 캡처·공유는 '파일을 만들어 줄 뿐'이고 받는 사람은 사람이 고른다."""
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    # (1) 요약 캡처 + 건별 목록 캡처가 둘 다 있는가
    for fn in ("mineToPng", "captureMine", "shareMine",
               "mineDetailToPng", "captureMineDetail", "shareMineDetail"):
        assert "function " + fn in idx, f"{fn} 없음"
    assert 'onclick="captureMine()"' in idx and 'onclick="captureMineDetail()"' in idx, "버튼이 안 걸렸다"
    assert "window._mineIdx" in idx, "건별 캡처가 어느 항목인지 모른다"
    # 건수가 많을 때 조용히 자르면 받는 사람이 그게 전부인 줄 안다
    assert "MINE_CAP" in idx and "이 이미지에는" in idx, "잘린 사실을 이미지에 안 적는다"

    # (2) 자동 전송 경로가 없는가 — 공유는 navigator.share(사람이 수신자를 고름)만 허용
    for bad in ("mailto:", "smtplib", "sendmail", "api.telegram", "hooks.slack",
                "kakao.link", "openapi.kakao", "openapi.band", "band.us/api"):
        assert bad not in idx, f"앱에 자동 전송 경로가 있다: {bad}"
    # 밴드·카톡 모듈은 읽기 전용이어야 한다
    import glob as _g
    for f in _g.glob(os.path.join(ROOT, "band", "*.py")) + _g.glob(os.path.join(ROOT, "kakao", "*.py")):
        src = open(f, encoding="utf-8").read()
        for bad in ("create_post", "write_post", "post_comment", "mailto:", "smtplib"):
            assert bad not in src, f"{os.path.basename(f)} 에 글쓰기·발송 경로가 있다: {bad}"

    # (3) 메일 경로가 되살아나지 않았는가(2026-07-27 제거)
    assert not os.path.exists(os.path.join(ROOT, ".github", "workflows")), \
        "워크플로가 되살아났다 — 메일이 갈 수 있다"

    # (4) 규칙이 문서에 남아 있는가 — 다음 AI가 이어받아도 지키게
    ag = open(os.path.join(ROOT, "AGENTS.md"), encoding="utf-8").read()
    assert "쿠팡 담당자에게 어떤 메시지도 보내지 않는다" in ag, "절대규칙에 안 적혀 있다"

    # (5) '원천' 같은 속어를 화면 문구로 쓰지 않는다(사용자가 무슨 뜻인지 되물었다)
    _vis = [ln for ln in idx.splitlines()
            if "원천 없음" in ln and not ln.strip().startswith(("*", "/*", "//", "★"))]
    assert not _vis, "화면에 '원천 없음'이 남아 있다: " + str(_vis[:1])

    # (6) 자료 유무 판정이 상태 도착 뒤에 다시 그려지는가
    #     (안 그러면 밴드·카톡·PO가 멀쩡한데 '아직 없음'으로 4건이 거짓 표시된다)
    _ls = idx[idx.index("async function loadStatus()"):][:900]
    assert "srcStats = s.sources" in _ls and "renderBoard()" in _ls, \
        "상태를 받은 뒤 확인목록을 다시 그리지 않는다 — 자료가 있어도 '없음'으로 뜬다"
    # 앱서버 자동 게시와 사람/다른 AI 수동 게시가 겹치면 같은 Git 커밋이 두 번 생긴다.
    cp = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert 'require("publish"' in cp and "acquired_here" in cp, \
        "폰 사본 Git 게시에 publish 점유가 강제되지 않는다"
    assert '"CSOS_AI": "server"' in srv and "publish_env" in srv, \
        "앱서버 자동 게시가 점유 주체를 밝히지 않는다"
    print("  [34] 확인목록 캡처(요약·건별) · 자동 전송 경로 없음 · 자료판정 순서 ✅")


def t35_confirm_evidence():
    """[35] '확인 완료' 체크는 **근거가 완료라고 말할 때만** 찍혀야 한다.

    한 건에 글이 여러 개 올라온다(접수 안내 → 작업완료). band_map 이 **먼저 나온 것**만
    집던 탓에 접수 글이 잡혀 완료 글을 놓쳤다 — 전체로 '첫 기록이 완료' 382 vs
    '완료 기록이 하나라도 있음' 859. 그 상태로 확인 체크를 돌리면 근거가 '접수·예정'인
    행에 '일치·확인완료'가 찍힌다. 확인하지 않은 것을 확인했다고 남기는 것이라 가장 나쁘다.
    (2026-07-27: 대상 40건 중 37건이 그랬다)"""
    import confirm_fill as C

    # (1) 완료 글을 우선으로 고르는가 — 합성 레코드로 직접 확인
    fake = [{"프로젝트NO": "UJ9000001", "진행상태": "접수·예정", "작업일": "2026-06-01", "업무유형": "돌발AS"},
            {"프로젝트NO": "UJ9000001", "진행상태": "작업완료", "작업일": "2026-06-03", "업무유형": "돌발AS"},
            {"프로젝트NO": "UJ9000002", "진행상태": "접수·예정", "작업일": "2026-06-02", "업무유형": "돌발AS"}]
    import band_extract as B
    _orig = B.load_records
    try:
        B.load_records = lambda *a, **k: fake
        bm = C.band_map()
    finally:
        B.load_records = _orig
    assert bm["UJ9000001"]["status"] == "작업완료", "접수 글이 완료 글을 가린다"
    assert bm["UJ9000001"]["date"] == "2026-06-03", bm["UJ9000001"]
    assert bm["UJ9000002"]["status"] == "접수·예정", "없는 완료를 지어내면 안 된다"

    # (2) 실제 계획 — 확인 체크가 붙는 건은 **전부** 근거가 '작업완료'여야 한다
    real = C.band_map()
    _m, items, _stat = C.plan()
    VCOLS = ("관리자검증상태", "담당관리자", "최종확인일", "최종확인일(유현민 체크)")
    bad = [i["key"] for i in items if i["col"] in VCOLS
           and (real.get(i["key"]) or {}).get("status") != "작업완료"]
    assert not bad, ("근거가 완료가 아닌데 확인 체크를 찍으려 한다: " + ", ".join(sorted(set(bad))[:5]))

    # (3) 기존 값을 덮지 않는가 — 확인 기록을 덮어쓰면 되돌릴 수 없다
    assert all(i.get("only_if_empty") for i in items), "빈칸만 채우는 규칙이 빠졌다"
    # (4) 상태 열 자체는 건드리지 않는다(완료 여부는 사람이 정한다)
    assert not any(i["col"] in ("진행상태", "점검상태") for i in items), \
        "도구가 진행상태를 바꾸려 한다 — 완료 판정은 사람 몫이다"
    print("  [35] 확인 체크 근거 정합(완료 글 우선·불일치 0·빈칸만·상태 불변) ✅")


def t36_mobile_input():
    """[36] 폰에서 빈 항목을 편하게 채울 수 있는가 — 달력·드롭다운·확대 방지."""
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    # (1) 날짜는 손으로 치지 않는다 — 폰 달력이 뜨는 type="date"
    assert "function fieldInput" in idx, "입력칸 렌더 함수가 없다"
    _fi = idx[idx.index("function fieldInput"):][:1600]
    assert "type=\"date\"" in _fi, "날짜 칸이 달력으로 안 뜬다"
    assert 'inputmode="numeric"' in _fi, "숫자 칸에 숫자 키패드가 안 뜬다"
    assert "<select" in _fi, "선택지가 드롭다운이 아니다"
    # ★ 16px 미만이면 iOS 사파리가 입력할 때 화면을 확대해 버린다
    m = re.search(r"font-size:(\d+)px", _fi)
    assert m and int(m.group(1)) >= 16, f"입력 글자 크기 {m.group(1) if m else '?'}px — iOS에서 화면이 확대된다"

    # (2) 선택지는 **시트에서** 온다 — 화면에 박아 두면 사람이 바뀔 때 어긋난다
    assert "/api/codes" in idx and "loadCodes()" in idx, "코드 목록을 시트에서 안 읽는다"
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert "def get_codes(" in srv and "10_코드관리" in srv, "서버가 코드 시트를 안 읽는다"
    # 담당기사·관리자검증상태는 자유 입력이 아니어야 한다(오타·표기흔들림 방지)
    _spec = idx[idx.index("const INPUT_SPEC"):][:2000]
    assert "opts:'담당기사'" in _spec, "담당기사가 아직 자유 입력이다"
    assert "opts:'관리자검증상태'" in _spec, "관리자검증상태가 아직 자유 입력이다"

    # (3) 코드 목록을 못 받아도 입력 자체가 막히면 안 된다
    assert "자유 입력으로 떨어뜨린다" in _fi or "list.length" in _fi, "폴백이 없다"

    # (4) 기사 목록은 실제 배정 횟수 순 — 폰에서 자주 쓰는 사람이 위로
    assert "function byUsage" in idx, "기사 정렬이 없다"

    # (5) 실제 코드 시트에서 목록이 나오는가
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    wb = openpyxl.load_workbook(resolve_master(load_config()["reconcile"]["master_xlsx"]),
                                read_only=True, data_only=True)
    assert "10_코드관리" in wb.sheetnames, "코드 시트가 없다"
    rows = list(wb["10_코드관리"].iter_rows(min_row=4, values_only=True))
    wb.close()
    hdr = [str(h).strip() if h else "" for h in rows[0]]
    assert "담당기사" in hdr, "코드 시트에 담당기사 열이 없다"
    i = hdr.index("담당기사")
    names = [str(r[i]).strip() for r in rows[1:] if i < len(r) and r[i] not in (None, "")]
    for who in ("김준형", "권오철", "김필우", "차동호"):      # AGENTS.md 규칙 6의 기준 4인
        assert who in names, f"코드 시트 담당기사에 {who} 가 없다"
    # (6) 부가세는 **원장 값**을 쓴다 — 10%로 계산하면 반올림 때문에 서류와 어긋난다
    assert "function vatLine" in idx, "부가세 표기 함수가 없다"
    _v = idx[idx.index("function vatLine"):][:900]
    assert "r.부가세" in _v, "원장 부가세를 안 읽고 계산해 버린다"
    assert "합계−공급가액" in _v, "원장이 비었을 때 되짚은 값이라는 표시가 없다"
    assert "≠ 합계" in _v, "공급가+부가세가 합계와 안 맞을 때 알리지 않는다"
    er = open(os.path.join(ROOT, "ecount_reconcile.py"), encoding="utf-8").read()
    assert '"원장_부가세"' in er and "실제작업부가세" in er, "원장에서 부가세를 안 읽는다"

    # 실데이터 정합 — 공급가+부가세=합계 가 깨진 행이 있으면 원장이 틀린 것이다
    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    import app_server as _A
    bad = [r["정산ID"] for r in _A.real_settlements()
           if r.get("부가세") not in (None, "") and r.get("합계")
           and abs((r["공급가액"] or 0) + (r["부가세"] or 0) - (r["합계"] or 0)) > 1]
    assert not bad, "원장 공급가+부가세≠합계: " + ", ".join(bad[:5])
    # (7) 보고일·집계기준일을 **보고 화면에서 바로** 고칠 수 있는가
    #     예전에는 대시보드 맨 아래 카드에만 있어, 보고서를 보다 날짜가 틀린 걸 발견하면
    #     탭을 옮겨 찾아가야 했다.
    for fn in ("openRptDates", "saveRptDates", "rptDatesToday"):
        assert "function " + fn in idx, f"{fn} 없음"
    assert 'onclick="openRptDates()"' in idx, "보고 화면에 날짜 변경 버튼이 없다"
    _rd = idx[idx.index("function openRptDates"):][:1500]
    assert 'type="date"' in _rd, "보고 날짜가 달력으로 안 뜬다"
    _sv = idx[idx.index("function saveRptDates"):][:1200]
    assert "/api/set_dates" in _sv, "엑셀 00_대시보드에 안 쓴다"
    # 집계기준일이 보고일보다 뒤면 잘못 고른 것이다 — 그대로 쓰면 보고서가 어긋난다
    assert "집계기준일 > 보고일" in _sv.replace("집계기준일 &gt; 보고일", "집계기준일 > 보고일"),         "집계기준일이 보고일보다 뒤인 경우를 안 막는다"
    # 전 영업일 계산은 주말을 건너뛰어야 한다
    _td = idx[idx.index("function rptDatesToday"):][:600]
    assert "[0,6].includes" in _td.replace(" ", ""), "전 영업일 계산이 주말을 안 건너뛴다"
    print("  [36] 폰 입력(달력·드롭다운·16px·시트연동) · 부가세 원장값 · 보고일 즉시수정 ✅")


def t37_band_coverage():
    """[37] 밴드 수집이 안 닿는 기간을 '기사가 안 올렸다'로 몰지 않는가.

    2026-07-28: 류지영이 엑셀을 저장하면서 옛 행의 상태 수식이 계산됐고, 그 순간
    2025-12-08·09 작업 28건이 '밴드 게시 미확인'으로 떴다. 그런데 쿠팡AS 밴드 캐시는
    2025-12-16부터다 — 기사가 안 올린 게 아니라 우리가 그 이전을 못 긁어온 것이다.
    이대로 두면 없는 잘못으로 기사들을 추궁하게 된다."""
    src = open(os.path.join(ROOT, "band", "band_reconcile.py"), encoding="utf-8").read()
    assert "수집범위밖" in src, "수집이 안 닿는 기간을 따로 구분하지 않는다"
    # ★ 밴드별로 시작일을 보고 **가장 늦은 것**을 써야 한다. 합쳐서 최소값을 잡으면
    #   매출처 밴드가 쿠팡AS 밴드의 빈 구간을 덮어 버린다.
    assert "max(_first.values())" in src, "밴드별 시작일을 안 보고 합쳐서 판단한다"

    # 하위 도구는 '미확인'만 문제로 올려야 한다(수집범위밖은 사람이 할 일이 아니다)
    fe = open(os.path.join(ROOT, "findings_export.py"), encoding="utf-8").read()
    assert '밴드게시") == "미확인"' in fe, "확인필요 목록이 수집범위밖까지 문제로 올린다"
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert 'r.get("밴드게시") == "미확인"' in srv, "앱이 수집범위밖까지 문제로 센다"

    # 실제 리포트에 수집 시작일보다 앞선 건이 '미확인'으로 남아 있으면 안 된다
    import glob as _g, csv as _csv, json as _j
    from datetime import datetime as _dt
    rep = sorted(_g.glob(os.path.join(ROOT, "reports", "밴드대조_*.csv")))
    if rep:
        first = {}
        for f in _g.glob(os.path.join(ROOT, "band", "cache", "*.json")):
            if os.path.basename(f).startswith(("dump_", "raw_")):
                continue
            try:
                c = _j.load(open(f, encoding="utf-8"))
            except Exception:
                continue
            ds = [_dt.fromtimestamp(p["created_at"] / 1000).date()
                  for p in (c.get("posts") or {}).values() if p.get("created_at")]
            if ds:
                first[c.get("band_name", f)] = min(ds)
        if first:
            cut = max(first.values())
            bad = []
            for r in _csv.DictReader(open(rep[-1], encoding="utf-8-sig")):
                if r.get("밴드게시") != "미확인":
                    continue
                d = (r.get("완료일") or "")[:10]
                if d and _dt.strptime(d, "%Y-%m-%d").date() < cut:
                    bad.append(r.get("ID"))
            assert not bad, (f"수집 시작({cut}) 이전 건이 '미확인'으로 남아 있다: "
                             + ", ".join(x for x in bad[:5] if x))
    print("  [37] 밴드 수집범위 구분(밴드별 시작일·미확인만 문제로) ✅")


def t38_daily_brief():
    """[38] 대표 보고는 **숫자가 아니라 내용**이어야 한다(2026-07-28 통화 지시).

    "숫자를 나한테 보고하라는 게 아니야" — 왜 갔고 무슨 작업을 했고 유·무상은 어떻게
    됐는지, 추가작업이 생겼는지, 정기점검 분기 진행률이 어떤지가 보고 내용이다."""
    import daily_brief as D
    from datetime import date, timedelta

    data, _m = D.load()
    b = D.brief("2026-07-27", data)
    t = D.text(b)

    # (1) 대표가 물은 항목이 전부 구조에 있는가
    for k in ("돌발AS", "정기점검", "완료내역", "무상건", "추가작업건",
              "점검중유상", "AS전환", "내용미기입"):
        assert k in b, f"{k} 누락"
    assert "분기진행률" in b["정기점검"] and 0 <= b["정기점검"]["분기진행률"] <= 100

    # (2) ★ '완료일 없음'과 '아직 안 감'을 갈라야 한다.
    #     뭉치면 미처리 84건이라고 보고하게 되고(대표가 놀란다), 반대로 '없다'고 하면
    #     최근 건을 놓친다. 30일 기준으로 나눈다.
    assert "완료일미기입" in b["돌발AS"], "완료일 미기입을 미처리와 안 나눴다"
    assert b["돌발AS"]["미처리"] <= b["돌발AS"]["미처리"] + b["돌발AS"]["완료일미기입"]
    assert "최근 30일" in t, "미처리 숫자가 어느 범위인지 문장에 없다"

    # (3) 완료 건은 '왜 갔는지'가 있어야 보고가 된다
    for x in b["완료내역"]:
        assert set(("왜", "무엇", "비용", "추가작업")) <= set(x), x
    # 내용이 없으면 지어내지 말고 미기입으로 남겨야 한다
    assert all(x["무엇"] == "" for x in b["내용미기입"]), "미기입 판정이 틀렸다"

    # (4) 분기 계산이 맞는가 — 7월은 3분기
    assert b["정기점검"]["분기"].endswith("3분기"), b["정기점검"]["분기"]
    b1 = D.brief("2026-02-10", data)
    assert b1["정기점검"]["분기"].endswith("1분기"), b1["정기점검"]["분기"]

    # (5) 없는 날을 넣어도 죽지 않아야 한다(보고가 매일 돌아간다)
    empty = D.brief("2020-01-01", data)
    assert empty["돌발AS"]["완료"] == 0 and isinstance(D.text(empty), str)
    print("  [38] 대표 브리핑(내용 중심·미처리/미기입 분리·분기 진행률) ✅")


def t39_realtime_monitor():
    """[39] 입력 보호시간·이슈 상태전이·자동 진입점 방어."""
    from datetime import datetime
    from operation_window import is_input_window
    import realtime_monitor as R

    assert not is_input_window(datetime(2026, 7, 28, 7, 59, 59))
    assert is_input_window(datetime(2026, 7, 28, 8, 0, 0))
    assert is_input_window(datetime(2026, 7, 28, 9, 29, 59))
    assert not is_input_window(datetime(2026, 7, 28, 9, 30, 0))

    issue = R._issue("sample", "P2", "샘플", "근거", "조치")
    first, changes1 = R.reconcile_issues([issue], {}, datetime(2026, 7, 28, 10, 0))
    assert first[0]["status"] == "provisional" and not changes1
    second, changes2 = R.reconcile_issues(
        [issue], {"issues": first}, datetime(2026, 7, 28, 10, 5)
    )
    assert second[0]["status"] == "new" and changes2[0]["transition"] == "new"
    third, changes3 = R.reconcile_issues(
        [], {"issues": second}, datetime(2026, 7, 28, 10, 10)
    )
    assert third[0]["status"] == "resolved" and changes3[0]["transition"] == "resolved"

    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    watchdog = open(os.path.join(ROOT, "watchdog.py"), encoding="utf-8").read()
    endpoint = open(os.path.join(ROOT, "publish_endpoint.py"), encoding="utf-8").read()
    cloud = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    assert "if is_input_window()" in daily
    assert "if is_input_window()" in watchdog
    assert "from net_probe import probe" in watchdog
    main_body = watchdog[watchdog.index("def main():"):]
    assert "archive_versions(dry)" not in main_body
    assert "if is_input_window()" in endpoint and "if is_input_window()" in cloud
    formulas = open(os.path.join(ROOT, "fix_formulas.py"), encoding="utf-8").read()
    assert '"--file" in args' in formulas, "수식 검사가 임시 사본 대신 실원장을 열 수 있다"
    dashboard = open(os.path.join(ROOT, "dashboard_clean.py"), encoding="utf-8").read()
    monitor = open(os.path.join(ROOT, "realtime_monitor.py"), encoding="utf-8").read()
    assert '"--file" in args' in dashboard, "대시보드 검사가 임시 사본 대신 실원장을 열 수 있다"
    assert '"dashboard_clean.py"' in monitor and "dashboard_filldown_debris" in monitor

    claim = open(os.path.join(ROOT, "ai_claim.py"), encoding="utf-8").read()
    assert "os.mkdir(GUARD)" in claim and "os.replace(tmp, CLAIMS)" in claim
    print("  [39] 입력시간 완전정지·이슈 신규/지속/해결·점유 원자화 ✅")


def t41_dates_explicit():
    """[41] '금일'은 읽는 시점마다 다른 날이 된다 — 보고물에 남아 있으면 안 된다.

    사고 #0(2026-07-28): 타일 라벨이 '점검 예정 (금일)'인데 숫자는 집계기준일(어제) 것이라
    보고 자리에서 어긋났다. 화면·이미지·문장 세 곳 모두 **실제 날짜**로 말해야 한다.
    덤으로 사고 #16(서버 두 개)의 진짜 원인인 포트 재사용도 여기서 함께 막는다."""
    import daily_brief as DB

    # (1) 브리핑 문장 — 항목 줄이 전부 날짜로 시작하는가
    day = "2026-03-04"
    b = {
        "기준일": day,
        "돌발AS": {"신규접수": 1, "완료": 1, "미처리": 0, "완료일미기입": 0},
        "정기점검": {"예정": 0, "완료": 0, "분기": "2026년 1분기",
                     "분기예정": 0, "분기완료": 0, "분기진행률": 0},
        "완료내역": [], "무상건": [], "추가작업건": [],
        "점검중유상": [], "AS전환": [], "이상발견": [], "내용미기입": [],
        "완료일미기입목록": [],
        "신규목록": [{"프로젝트NO": "UJ2600001", "캠프명": "시험캠프", "담당기사": "홍길동",
                      "왜": "시험 접수내용", "무엇": "", "비용": "유상", "추가작업": "",
                      "일자": day, "구분": "돌발AS", "접수일": day}],
    }
    b["완료내역"] = [dict(b["신규목록"][0], 일자=day, 접수일="2026-02-25", 무엇="시험 작업")]
    out = DB.text(b)
    assert day in out.splitlines()[0], "머리줄에 기준일이 없다"
    assert "03-04 · UJ2600001 · 시험캠프" in out, "항목이 '날짜 · 프로젝트NO · 캠프'로 시작하지 않는다"
    assert "접수 02-25 · 7일 만" in out, "완료건에 접수일·경과일이 안 붙는다"
    assert "금일" not in out and "당일" not in out, "브리핑 문장에 '금일/당일'이 남아 있다"

    # (1-2) 브리핑 화면은 **문장 줄바꿈이 아니라 구조화된 데이터**로 그려야 한다.
    #       줄로 흘려 놓으면 어디서 어디까지가 한 건인지 안 보인다(사용자 지적 2026-07-28).
    _ix = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for need in ("function bCard(", "function briefBlock(", "BRIEF['신규목록']",
                 "BRIEF['완료내역']", 'class="bcard"', 'class="bmiss"'):
        assert need in _ix, f"브리핑 카드 렌더링에 {need} 가 없다"
    assert "BRIEF.text" in _ix, "복사 버튼이 서버 문장을 쓰지 않는다(대표 전달용 형식이 깨진다)"

    # (1-3) 구 버전 자동 접기 — 사용자 지시(2026-07-28) "말 안 해도 OLD 로".
    #       저장하는 도구가 11개라 **찾는 쪽**(resolve_master·latest_master)에 건다.
    import ledger_versions as LV
    assert hasattr(LV, "autoprune"), "autoprune 이 없다"
    _lv = open(os.path.join(ROOT, "ledger_versions.py"), encoding="utf-8").read()
    _auto = _lv[_lv.index("def autoprune("):_lv.index("def main(")]
    assert "os.remove" not in _auto and "shutil.rmtree" not in _auto, \
        "자동 정리가 파일을 지운다 — 접어 두기만 해야 한다"
    assert "shutil.move" in _auto, "자동 정리가 옮기지 않는다"
    for f in ("ecount_reconcile.py", "workbook_patch.py"):
        s = open(os.path.join(ROOT, f), encoding="utf-8").read()
        assert "autoprune" in s, f"{f} 가 구 버전을 자동으로 접지 않는다"
    assert LV.KEEP_LATEST >= 1 and LV.ARCHIVE == "OLD", "보관 정책이 바뀌었다"

    # ★ 자동으로 도는 정리는 아무도 안 보고 있다. **남의 파일을 옮기면 안 된다.**
    #   처음 구현이 `*_v*.xlsx` 를 통째로 잡아 합성검증용 임시 파일까지 옮겨 시험이 깨졌다.
    import tempfile as _tf, shutil as _sh
    _d = _tf.mkdtemp()
    try:
        for _n in ("합성대장F_v1.xlsx", "합성대장F_v2.xlsx", "기타자료_v3.xlsx"):
            open(os.path.join(_d, _n), "w").write("x")
        LV._AUTODONE = False
        assert LV.autoprune(os.path.join(_d, "합성대장F_v2.xlsx")) == 0, "남의 파일을 옮겼다"
        assert len(os.listdir(_d)) == 3, "남의 폴더를 건드렸다"

        for _v in (1, 2, 3):
            open(os.path.join(_d, f"쿠팡_통합업무_일일보고_관리대장_v{_v}.xlsx"), "w").write("x")
        open(os.path.join(_d, "쿠팡_통합업무_일일보고_관리대장_v9_보관.xlsx"), "w").write("x")
        LV._AUTODONE = False
        assert LV.autoprune(os.path.join(_d, "쿠팡_통합업무_일일보고_관리대장_v3.xlsx")) == 2, \
            "옛 버전 2개가 접히지 않았다"
        left = sorted(x for x in os.listdir(_d) if x.endswith(".xlsx"))
        assert "쿠팡_통합업무_일일보고_관리대장_v3.xlsx" in left, "최신본이 접혔다"
        assert "쿠팡_통합업무_일일보고_관리대장_v9_보관.xlsx" in left, "'보관' 표시본이 접혔다"
        assert "기타자료_v3.xlsx" in left and "합성대장F_v1.xlsx" in left, "남의 파일이 접혔다"
        assert sorted(os.listdir(os.path.join(_d, "OLD"))) == [
            "쿠팡_통합업무_일일보고_관리대장_v1.xlsx",
            "쿠팡_통합업무_일일보고_관리대장_v2.xlsx"], "접힌 목록이 다르다"
    finally:
        LV._AUTODONE = False
        _sh.rmtree(_d, ignore_errors=True)

    # (2) 화면·이미지 — '(금일)' 라벨을 실제 날짜로 바꾸는 코드가 살아 있는가
    src = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert "function dateLabel(" in src, "dateLabel 이 없다"
    assert src.count("dateLabel(l)") >= 3, "타일·셀·이미지 세 곳 모두 dateLabel 을 거쳐야 한다"
    assert 'class="dbanner"' in src, "기준일 안내 띠가 없다"
    # 도움말·목록 조회는 **원본 라벨**로 해야 한다(날짜를 박은 문자열로 찾으면 전부 어긋난다)
    assert "helpKey(l)" in src and "dayDetail(l,v)" in src, "라벨 가공본으로 조회하고 있다"

    # (3) 서버 두 개 방지 — 포트 재사용이 꺼져 있어야 새 서버가 조용히 묻히지 않는다
    ap = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert "allow_reuse_address = False" in ap, "포트 재사용이 켜져 있어 서버가 두 개 뜬다"

    # (4) 원본 자료 취합 — 종류별 폴더가 한 곳에서만 정의돼야 한다
    import source_dirs as SD
    for name in ("ERP_DIR", "COUPANG_DIR", "KAKAO_DIR", "BAND_DIR"):
        assert getattr(SD, name).startswith(SD.ORIGIN_ROOT), f"{name} 이 원본 폴더 밖을 가리킨다"
    cs = open(os.path.join(ROOT, "collect_sources.py"), encoding="utf-8").read()
    assert "shutil.move" not in cs and "os.remove" not in cs, "취합 도구가 원본을 지우거나 옮긴다"

    # (4-2) 이카운트 내보내기는 파일명이 무작위(8W1JR7MGB50PHOP.xlsx)라 **내용으로** 갈라야 한다.
    #       2026-07-28: 매출(세금)계산서조회(재고)가 '거래명세서 현황'으로 잘못 잡히고
    #       홈택스 전자(세금)계산서는 아예 판별 실패였다.
    from inbox_scan import classify_rows as _CR
    for name, rows, want in (
        ("매출(세금)계산서조회(재고)",
         [["일자 - 번호", "거래처명", "담당자 이메일주소", "공급가액", "부가세", "합 계", "내역보기"],
          ["2026/07/27 -2", "코리아종합물류", "", "352000", "35200", "387200", "내역보기 거래명세서"]],
         "taxinv"),
        ("홈택스 전자(세금)계산서",
         [["일자", "전송일자", "구분", "승인번호", "공급자사업자번호", "공급자상호",
           "공급받는자사업자번호", "공급받는자상호", "기준품목", "계산서분류"],
          ["2026/07/23", "2026/07/23", "매입", "2026072310", "2150351494", "하나산업",
           "8678100168", "유니버셜", "제빙기수리비", "세금계산서"]],
         "hometax"),
        # 기존 판별이 새 규칙에 밀리지 않는지 (겹치는 머리글이 있다)
        ("거래명세서 현황", [["일자-번호", "거래처명", "공급가액", "부가세", "합계"],
                             ["2026/07/01-1", "쿠팡", "100", "10", "110"]], "stmt"),
        ("매출(세금)계산서현황", [["일자-번호", "거래처명", "매출합계", "매출부가세"],
                                  ["2026/07/01-1", "쿠팡", "100", "10"]], "tax"),
        ("거래처별계정별원장", [["일자", "적요", "차변", "대변"],
                                ["2026-07-01", "매출", "100", "0"]], "ledger"),
    ):
        got = _CR(rows)
        assert got == want, f"{name} 판별 {got!r}, 기대 {want!r}"

    # Downloads 를 훑되 **아는 종류만** 가져와야 한다(개인 폴더를 통째로 퍼 오면 안 된다)
    assert "KNOWN" in cs and "kind_of" in cs, "Downloads 에서 종류를 안 가리고 가져온다"

    # (5) 대시보드 채우기 내림 잔해 — 관리대장에 닿을 때만 본다(네트워크가 끊겨도 검증은 돈다)
    #     빈 구역에 수식이 흘러들면 화면이 지저분해질 뿐 아니라 입력칸(H46)이 0 으로 덮여
    #     '당일 실적' 표가 통째로 죽는다(2026-07-28 실사고).
    try:
        import dashboard_clean as DC
        from workbook_patch import latest_master
        mm = latest_master()
        src = mm[0] if isinstance(mm, tuple) else mm
    except Exception:
        src = None
    # data_only=True에서는 캐시가 없는 수식이 None으로 보여 잔해를 놓친다.
    # XML을 직접 봐서 캐시 없는 수식은 잡고, 스타일만 남은 셀은 통과시킨다.
    _dash = ('<sheetData><row r="27">'
             '<c r="B27" s="1"><f>COUNTIF($A$1:$A$2,1)</f><v/></c>'
             '<c r="E27" s="1"/><c r="H27" s="1" t="inlineStr"><is><t>잔해</t></is></c>'
             '</row></sheetData>')
    _left = DC.survey_xml(_dash)
    assert [r for r, _v, _w in _left] == ["B27", "H27"], _left
    # (6) 담당기사 칸에 사람 아닌 값이 들어가면 대표보고 TOP 5 에 그대로 노출되고
    #     기사별 집계가 오염된다(2026-07-28 실사고 10건). 뽑는 단계에서 막는다.
    from band_extract import normalize_tech as NT
    for raw, want in (("000 (캠프상태확인 및 스케쥴 세팅)", ""),
                      ("자) - 각캠프담당자 캠프 컨디션상태 체크", ""),
                      ("하이테크 + 엄진언", "엄진언"),      # 업체는 빼고 사람만
                      ("김혜진 대신택배", "김혜진"),
                      ("김승기기장", "김승기"),             # 붙여 쓴 직책 분리
                      ("김필우 기사", "김필우"),
                      ("권오절", "권오철"),                 # 기존 오탈자 교정 유지
                      ("김준형, 권오철", "김준형, 권오철"),
                      ("미배정", ""), ("000", "")):
        got = NT(raw)
        assert got == want, f"normalize_tech({raw!r}) = {got!r}, 기대 {want!r}"

    # (7) 셀 교체가 **옆 칸을 먹어치우지 않는지**. 빈 셀은 자기닫힘(`<c .../>`)이라
    #     느슨한 정규식이 `/` 를 삼키고 다음 `</c>` 까지 지운다(2026-07-28 실사고 — AK33 소실).
    from workbook_patch import replace_inline_cell as _RC
    _x = '<c r="AJ33" s="5"/><c r="AK33" s="7" t="inlineStr"><is><t>지켜야함</t></is></c>'
    _o = _RC(_x, "AJ33", "메모")
    assert "지켜야함" in _o and 'r="AK33"' in _o, "빈 셀을 교체하며 옆 칸을 삼켰다"
    assert _o.count("<c ") == 2, f"셀 개수가 바뀌었다: {_o}"

    if src and os.path.exists(src):
        import openpyxl as _ox
        from project_resolve import clean_tech
        _w = _ox.load_workbook(src, read_only=True, data_only=True)
        junk = []
        # ★ 02시트만 보다가 03_현장작업실적에 같은 값이 남았다 — 기사 열이 있는 시트를 전부 본다
        for _sn in ("02_돌발AS접수", "03_현장작업실적", "04_정기점검", "05_신규납품설치"):
            if _sn not in _w.sheetnames:
                continue
            _s = _w[_sn]
            _h = [str(h or "").strip() for h in
                  next(_s.iter_rows(min_row=4, max_row=4, values_only=True))]
            _cs = [i for i, h in enumerate(_h) if h in ("담당기사", "담당자", "작업자")]
            for _row in _s.iter_rows(min_row=5, values_only=True):
                if not _row[0]:
                    continue
                for _i in _cs:
                    _v = str(_row[_i] or "").strip() if _i < len(_row) else ""
                    if _v and clean_tech(_v) != _v:
                        junk.append(f"{_sn}:{_v}")
        _w.close()
        assert not junk, ("담당기사 칸에 사람 아닌 값 %d건 — "
                          "python fix_tech_names.py 로 정정: %s" % (len(junk), junk[:3]))
        left = DC.survey(src)
        assert not left, ("대시보드에 채우기 내림 잔해 %d칸 — python dashboard_clean.py 로 정리: %s"
                          % (len(left), ", ".join(r for r, _v, _w in left[:8])))
        import openpyxl
        w = openpyxl.load_workbook(src, data_only=False)
        assert w["00_대시보드"]["H46"].value in (None, ""), \
            "H46(조회일 직접입력)이 비어 있지 않다 — 당일 실적 표가 죽는다"
        w.close()
        print("  [41] … 대시보드 잔해 0칸 · 입력칸 H46 정상")
    print("  [41] 날짜 명시(금일 금지·항목별 날짜)·서버 중복 방지·원본 취합 안전 ✅")


def t40_claim_enforced():
    """[40] 협업 규칙이 **문서로만** 있으면 잊는다 — 원장 도구가 실제로 막아야 한다.

    두 AI가 동시에 원장을 고치면 각자 vN+1을 만들어 한쪽이 통째로 묻힌다(되돌릴 수 없다).
    ai_claim 은 규칙을 적는 곳이고, claim_guard 가 그 규칙을 강제한다."""
    import claim_guard, ai_claim

    # (1) 원장을 쓰는 도구는 전부 가드를 거쳐야 한다
    for fn in ("ledger_writer.py", "workbook_patch.py", "expand_rows.py", "confirm_fill.py",
               "fix_formulas.py", "dashboard_clean.py", "reorder_rows.py"):
        src = open(os.path.join(ROOT, fn), encoding="utf-8").read()
        assert "claim_guard" in src and 'require("ledger"' in src, f"{fn} 이 점유를 확인하지 않는다"

    # 실제 협업 점유 파일을 시험이 비우면, 합성검증 도중 다른 AI가 원장 쓰기에 진입한다.
    # 독립 임시 점유 파일에서만 충돌·자동 점유를 검증하고 실제 점유는 그대로 보존한다.
    real_claims, real_guard = ai_claim.CLAIMS, ai_claim.GUARD
    with tempfile.TemporaryDirectory() as claim_tmp:
        ai_claim.CLAIMS = os.path.join(claim_tmp, "claims.json")
        ai_claim.GUARD = os.path.join(claim_tmp, ".guard")
        try:
            # (2) 남이 잡고 있으면 멈춰야 한다
            assert ai_claim.take("codex", "ledger", "합성검증"), "시험용 점유를 못 잡았다"
            try:
                os.environ["CSOS_AI"] = "claude"
                try:
                    claim_guard.require("ledger", "test")
                    raise AssertionError("남이 잡았는데 그냥 진행했다 — 동시 수정이 일어난다")
                except SystemExit as e:
                    assert e.code == 3, e.code
                # (3) 사람이 직접 실행할 때(CSOS_AI 없음)는 막지 않는다 — 손을 묶으면 안 된다
                os.environ.pop("CSOS_AI", None)
                assert claim_guard.require("ledger", "test") is True
            finally:
                os.environ.pop("CSOS_AI", None)
                ai_claim.free("codex", "ledger")

            # (4) 아무도 안 잡았으면 자동으로 잡고 진행한다
            os.environ["CSOS_AI"] = "claude"
            try:
                assert claim_guard.require("ledger", "test") is True
                assert (ai_claim.load().get("ledger") or {}).get("who") == "claude"
            finally:
                ai_claim.free("claude", "ledger")
                os.environ.pop("CSOS_AI", None)
        finally:
            ai_claim.CLAIMS, ai_claim.GUARD = real_claims, real_guard
            os.environ.pop("CSOS_AI", None)
    print("  [40] 점유 강제(원장 도구 차단·사람 실행 허용·자동 점유) ✅")


def t42_first_empty_row(tmp):
    """[42] '빈 행'을 프로젝트NO 한 열로 판정하면 **남의 행에 번호를 얹는다**.

    2026-07-28 실사고: 02시트 547행은 사람이 번호 없이 내용만 적어 둔 행
    (M_순천1·김필우·리모컨)이었는데, 프로젝트NO 가 비었다는 이유로 '빈 행'이 되어
    전혀 다른 건(UJ2601347 중구1·경광등)의 번호가 그 행에 들어갔다.
    나머지 칸은 '빈 칸만' 정책이 막아 줘서 겉으로는 조용했지만 한 행에 두 건이 섞였다.
    → 빈 행 판정은 **행 전체**로 한다. 번호만 없는 행은 '새 행'이 아니라
      내용으로 맞춰 **번호를 채워 줄 대상**이다(새 행을 만들면 같은 건이 두 행이 된다)."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import openpyxl
    from datetime import date as _date
    import kakao_extract as kx
    import backfill_rows as bf

    path = os.path.join(tmp, "합성_원장.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "02_돌발AS접수"
    hdr = ["접수ID", "프로젝트NO", "캠프명", "접수일자", "신청내용", "담당기사"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(row=4, column=i, value=h)
    for r, code in ((5, "UJ0000001"), (6, "UJ0000002")):
        ws.cell(row=r, column=2, value=code)
        ws.cell(row=r, column=3, value="캠프%d" % r)
        ws.cell(row=r, column=4, value=_date(2026, 7, 20))
    # 7행: 번호만 없는 기존 행(사람이 먼저 적어 둔 행) — 빈 행이 아니다
    ws.cell(row=7, column=3, value="M_순천1")
    ws.cell(row=7, column=4, value=_date(2026, 7, 27))
    ws.cell(row=7, column=5, value="리프트 리모콘 작동 안함")
    ws.cell(row=7, column=6, value="김필우")
    ws.cell(row=12, column=1, value="")          # 용량만 넓힌다(빈 행)
    wb.save(path)

    start, cols, cap, orphans = kx.sheet_state(path, "02_돌발AS접수")
    assert start == 8, "번호 없는 7행을 빈 행으로 봤다 — 남의 행에 번호를 얹는다 (start=%s)" % start
    assert [o["행"] for o in orphans] == [7], orphans
    assert orphans[0]["내용키"] == kx._key("리프트 리모콘 작동 안함"), orphans

    b_start, _, _ = bf.sheet_state(path, "02_돌발AS접수")
    assert b_start == 8, "backfill_rows 도 같은 함정에 빠진다 (start=%s)" % b_start

    # 공지 일자와 원장 일자는 며칠 어긋난다(공지가 다음 날 올라온다) — ±3일은 같은 건
    assert kx._near("2026-07-28", {_date(2026, 7, 27)}) is True
    assert kx._near("2026-07-28", {_date(2026, 7, 20)}) is False
    assert kx._near("", {_date(2026, 7, 27)}) is False

    # 두 자리 연도('26.07.28')를 못 읽으면 신청일자가 통째로 빈칸이 된다(실제 UJ2601345)
    assert kx.norm_date("26.07.28") == _date(2026, 7, 28)
    assert kx.norm_date("2026.00.00 (요일)") is None          # 템플릿은 채우지 않는다
    # 유형을 모르면 찍지 않는다 — 억지로 02·04 로 몰면 엉뚱한 시트에 등록된다
    assert kx.kind_of("♣ ［ 돌발유료 A/S 안내 ]") == "02_돌발AS접수"
    assert kx.kind_of("♣ ［ 2026년 03분기 3개월 유료 A/S 안내 ]") == "04_정기점검"
    assert kx.kind_of("♣ ［쿠팡 철거보관 안내］") not in ("02_돌발AS접수", "04_정기점검")
    assert kx.status_of("♣ ［ 돌발유료 A/S 완료 ]") == "완료"
    assert kx.status_of("✅ 접수취소 ♣ ［ 돌발유료 A/S 안내 ]") == "취소"
    # 담당기사 칸에 작업 메모가 들어가면 대표보고에 그대로 노출된다(2026-07-28 실사고)
    from band_extract import normalize_tech
    assert normalize_tech("000 (캠프상태확인 및 스케쥴 세팅)") == ""
    print("  [42] 빈 행 판정(행 전체)·번호없는 행 채움·카톡 파싱(2자리연도·유형·상태) ✅")


def t43_receipt_fill(tmp):
    """[43] 입금 자동입력 — 합계행을 입금으로 세면 가짜 수금이 생긴다.

    폰 앱 '빈 항목 입력'의 입금일·입금액만 사람이 손으로 넣고 있었다(사용자 지시 2026-07-28:
    자료 확인되면 자동화). 근거는 거래처별계정별원장의 **대변**이다. 함정 셋:
      · '월 계'·'누 계'·'전기이월' 행이 큰 금액을 갖고 있다 → 그대로 세면 가짜 입금 1건
      · 차변은 매출 발생이지 입금이 아니다
      · 쿠팡은 여러 건을 묶어 한 번에 넣는다 → 금액이 겹치면 엉뚱한 정산행에 붙는다"""
    import sys as _s
    _s.path.insert(0, ROOT)
    import openpyxl
    from datetime import date as _date
    import receipt_fill as rf

    path = os.path.join(tmp, "합성_계정별원장.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["회사명 : 합성"])
    ws.append(["일자", "일자-No.", "적요", "차변", "대변"])
    ws.append([_date(2026, 1, 10), "2026/01/10 -3", "제품매출", 3043150, None])   # 차변=매출
    ws.append([_date(2026, 2, 15), "2026/02/15 -1", "보통예금 입금", None, 3043150])
    ws.append([_date(2026, 2, 20), "2026/02/20 -2", "보통예금 입금", None, 500000])
    ws.append([None, None, "2026/02 월 계", 3043150, 3543150])                    # 합계행
    ws.append([None, None, "누 계", 3043150, 3543150])                            # 합계행
    wb.save(path)

    got, ok = rf.parse_receipts(path)
    assert ok, "적요+대변 머리글을 못 찾았다"
    assert len(got) == 2, "합계행·차변행이 섞였다: %s" % got     # 차변 1 + 합계 2 는 빠져야 한다
    assert {int(g["금액"]) for g in got} == {3043150, 500000}, got
    assert got[0]["일자"] == _date(2026, 2, 15), got[0]

    # 유일 일치만 자동입력 — 금액이 같은 후보가 둘이면 사람에게 넘긴다
    rows = [
        {"정산ID": "S1", "청구액": 3043150.0, "프로젝트NO": "UJ1", "캠프명": "양주2",
         "발행일": _date(2026, 1, 10)},
        {"정산ID": "S2", "청구액": 500000.0, "프로젝트NO": "UJ2", "캠프명": "부산1",
         "발행일": _date(2026, 1, 11)},
        {"정산ID": "S3", "청구액": 500000.0, "프로젝트NO": "UJ3", "캠프명": "창원1",
         "발행일": _date(2026, 1, 12)},
    ]
    paired, spare = rf.match(got, rows)
    assert [s["정산ID"] for _, s in paired] == ["S1"], paired      # 500000 은 후보 2건 → 보류
    assert len(spare) == 1, spare

    # 발행일보다 앞선 입금은 그 건의 입금이 아니다
    early = [{"일자": _date(2025, 12, 1), "금액": 3043150.0, "전표": "", "적요": "", "출처": ""}]
    assert rf.match(early, rows[:1])[0] == [], "청구 전에 들어온 돈을 그 건에 붙였다"

    # 청구일·지급예정일은 채우지 않는다 — 산정 규칙이 확정되기 전엔 만들어 넣으면 안 된다
    src = open(os.path.join(ROOT, "receipt_fill.py"), encoding="utf-8").read()
    body = src.split('if __name__')[0]
    assert '"col": "청구일"' not in body and '"col": "지급예정일"' not in body, \
        "규칙이 확정되지 않은 청구일·지급예정일을 채우고 있다"
    # 오종현 관리 '26년도 쿠팡 입금내역' — 사람이 만드는 표라 제목·빈 줄 수가 달라진다.
    # 머리글 행을 박아 두면 다음 달에 한 줄 밀리는 순간 조용히 0건이 된다.
    dep = os.path.join(tmp, "합성_입금내역.xlsx")
    wb2 = openpyxl.Workbook()
    w2 = wb2.active
    w2.append([])
    w2.append([None, "2026 쿠팡 입금 내역"])          # 제목 — 머리글이 아니다
    w2.append([])
    w2.append([None, "날짜", "거래처", "입금액"])       # 실제 머리글(4행)
    w2.append([None, _date(2026, 7, 27), "쿠팡로지스틱스", 715000])
    w2.append([None, _date(2026, 7, 27), "쿠팡로지스틱", 814000])      # 같은 거래처, 표기만 다름
    w2.append([None, _date(2026, 5, 6), "김진주（위더스 )", 387200])   # 전각 괄호
    w2.append([None, None, "합 계", 1916200])                          # 합계행
    wb2.save(dep)

    got2 = rf.parse_deposit_list(dep)
    assert len(got2) == 3, "합계행이 입금으로 세어졌다: %s" % got2
    assert sum(g["금액"] for g in got2) == 1916200, got2
    custs = {g["거래처"] for g in got2}
    assert "쿠팡로지스틱스" in custs and "쿠팡로지스틱" not in custs, \
        "같은 거래처가 표기 차이로 갈라졌다 — 금액이 둘로 쪼개진다: %s" % custs
    assert len(custs) == 2, custs                     # 쿠팡 + 김진주위더스
    assert rf.norm_cust("(주)모벤티스") == rf.norm_cust("주식회사 모벤티스"), "법인격 표기 미정규화"
    print("  [43] 입금 자동입력(합계행 제외·차변 제외·유일매칭·머리글 자동탐지·거래처 정규화) ✅")


def t44_zscan():
    """[44] 쿠팡 업무 폴더 2만 개 — **파일을 열지 않고** 파일명으로 고르고 대조한다.

    사용자 지시(2026-07-28): "관련있는 자료는 전부 긁어와서 비교 검토, 관련 없는 자료는
    db에 적용하지마." 지키는 방법은 쓰기 경로를 좁게 만드는 것이다 —
    캠프명+날짜(±7일)가 **둘 다** 맞는 1:1 확정 건만 쓰고, 후보가 여럿이면 사람에게 넘긴다.
    (같은 캠프에 한 달에 여러 건이 있고 같은 날 여러 캠프를 돈다 — 한쪽만으로는 못 잇는다)"""
    import sys as _s
    _s.path.insert(0, ROOT)
    import zscan

    # 안전·교육 폴더는 업무상 필요해도 관리대장에 들어갈 자리가 없다
    assert zscan.classify("♣ 6. 설치공사 안전보건대장", "허가서.pdf").startswith("무관")
    assert zscan.classify("♣ 7. 쿠팡 지게차 서류", "UJ2600136 지게차.pdf").startswith("무관")
    assert zscan.classify("♣ 2. 쿠팡 돌발AS", "철거 906,000원 UJ2600136.PDF").startswith("관련")
    assert zscan.classify("♣ 1. 쿠팡 정기점검", "2026-06-19 구로1MB 거래명세서.pdf").startswith("관련")
    assert zscan.classify("♣ 100. 쿠팡 기타자료", "현장사진.jpg").startswith("무관")

    # 캠프명 표기 흔들림 — 괄호 안 표기가 달라도 같은 캠프로 본다
    assert zscan.camp_key("구로1MB(독산동B)") == zscan.camp_key("구로1MB (독산동A)")
    assert zscan.camp_key("일산2MB(양평동4가B)") != zscan.camp_key("일산3MB(양평동6가B)")

    docs = [{"일자": "2026-06-19", "캠프키": zscan.camp_key("구로1MB(독산동B)"),
             "캠프원문": "구로1MB", "종류": "거래명세서", "파일": "a.pdf", "폴더": "x"},
            {"일자": "2026-06-19", "캠프키": zscan.camp_key("없는캠프"),
             "캠프원문": "없는캠프", "종류": "거래명세서", "파일": "b.pdf", "폴더": "x"}]
    rows = [{"시트": "04_정기점검", "ID": "PM-1", "프로젝트NO": "UJ1", "캠프명": "구로1MB(독산동B)",
             "캠프키": zscan.camp_key("구로1MB(독산동B)"), "일자": "2026-06-20"}]
    paired, orphan, amb = zscan.match_docs(docs, rows)
    assert len(paired) == 1 and paired[0][1]["프로젝트NO"] == "UJ1", paired   # 하루 차이는 같은 건
    assert len(orphan) == 1 and not amb, (orphan, amb)

    # 같은 캠프에 후보가 둘이면 **쓰지 않는다** — 엉뚱한 행에 발행완료가 찍힌다
    rows2 = rows + [{"시트": "04_정기점검", "ID": "PM-2", "프로젝트NO": "UJ2",
                     "캠프명": "구로1MB(독산동B)", "캠프키": zscan.camp_key("구로1MB(독산동B)"),
                     "일자": "2026-06-18"}]
    p2, _o2, a2 = zscan.match_docs(docs, rows2)
    assert not p2 and len(a2) == 1, (p2, a2)

    # 날짜가 멀면 다른 건이다
    far = [dict(docs[0], 일자="2026-05-01")]
    assert zscan.match_docs(far, rows)[0] == []
    print("  [44] 쿠팡 폴더 조사(무관 폴더 제외·캠프키·1:1 확정만 인정) ✅")


def t45_cloud_queue_and_erp_documents(tmp):
    """[45] PC 독립 큐와 ERP 문서 완료는 근거가 있는 빈칸만 건드린다."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import fill_erp_documents as fed

    book = os.path.join(tmp, "erp_docs.xlsx")
    wb = openpyxl.Workbook()
    a = wb.active
    a.title = "02_돌발AS접수"
    a.append([]); a.append([]); a.append([])
    a.append(["접수ID", "프로젝트NO", "진행상태", "ERP등록"])
    a.append(["AS-1", "UJ0000001", "작업완료", ""])
    a.append(["AS-2", "UJ0000002", "접수", ""])
    p = wb.create_sheet("04_정기점검")
    p.append([]); p.append([]); p.append([])
    p.append(["점검ID", "프로젝트NO", "점검상태", "ERP판매전표", "거래명세서"])
    p.append(["PM-1", "UJ0000001", "완료", "", ""])
    p.append(["PM-2", "UJ0000002", "완료", "", ""])
    wb.save(book)
    original = fed.document_evidence
    fed.document_evidence = lambda: {
        "UJ0000001": {"판매전표", "거래명세서"},
        "UJ0000002": {"판매전표"},
    }
    try:
        fills, counts, _ = fed.plan(book)
    finally:
        fed.document_evidence = original
    assert fills == {
        "02_돌발AS접수!D5": "완료",
        "04_정기점검!D5": "완료",
        "04_정기점검!E5": "발행완료",
        "04_정기점검!D6": "완료",
    }, fills
    assert not any("C5" in ref or "C6" in ref for ref in fills), "업무 상태 셀을 직접 덮었다"
    assert counts["04_정기점검:거래명세서"] == 1

    app = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    snap = open(os.path.join(ROOT, "mobile_snapshot.py"), encoding="utf-8").read()
    sync = open(os.path.join(ROOT, "cloud_queue_sync.py"), encoding="utf-8").read()
    for need in ("cloud_queue", "Authorization", "/api/queue", "accepted", "duplicate"):
        assert need in app, f"폰 클라우드 큐 연결 누락: {need}"
    for need in ("/api/queue/lease", "/api/queue/ack", "/api/queue/release",
                 "is_input_window", 'take("cloud-sync", "ledger"'):
        assert need in sync, f"로컬 큐 유실방지 누락: {need}"
    for need in ("완료보고서등록", "ERP판매전표", "거래명세서"):
        assert need in snap, f"폰 사본 상태 누락: {need}"
    print("  [45] PC 독립 영구큐·ERP 문서근거 완료·앱 상태 전파 ✅")


def t46_app_2026_only():
    """[46] 앱 표시 경계는 2025년을 숨기고 2026년만 통과시킨다."""
    import sys as _s
    _s.path.insert(0, os.path.join(ROOT, "webapp"))
    import app_server as app

    assert not app.app_year_record(
        {"프로젝트NO": "UJ2601234", "접수일자": "2025-12-31"}, "as"), \
        "프로젝트 번호가 26이어도 실제 접수일이 2025면 2025 업무다"
    assert app.app_year_record(
        {"프로젝트NO": "UJ2501234", "접수일자": "2026-01-02"}, "as"), \
        "2026 접수 건을 프로젝트 번호만 보고 숨겼다"
    assert app.app_year_record({"업무ID": "AS-2607-001"}, "as")
    assert not app.app_year_record({"업무ID": "PM-2512-099"}, "pm")
    assert app.app_year_record({"월": "2026/07", "전표": "2026/07/24 - 3"}, "erp")
    assert app.app_year_record({"전표": "26/07/24 - 3"}, "erp")
    assert not app.app_year_record({"월": "2025/12", "전표": "25/12/30 - 1"}, "erp")
    assert app.app_year_record({"기준일": "2026-07-28", "문제내용": "완료보고 확인"}, "issue")
    assert not app.app_year_record({"캠프명": "연도 확인 불가"}), "연도 미상 행이 2026 목록에 섞였다"
    assert not app.app_year_record(
        {"포함프로젝트": "UJ2500001, UJ2600001"}), "2025·2026 혼합 행이 통째로 표시됐다"
    assert not app.app_project_result(
        "UJ2600007", {"date": "2025-12-31", "ids": {"접수ID": "AS-2512-040"}}), \
        "UJ26 코드만 보고 연결된 2025 작업을 앱에 노출했다"
    assert not app.app_project_result(
        "UJ2600008", {"date": "2026-01-02", "ids": {"접수ID": "AS-2512-041"}}), \
        "2026 날짜와 2025 업무ID가 섞인 자동조회 결과를 통과시켰다"
    assert app.app_project_result(
        "UJ2600009", {"date": "2026-01-03", "ids": {"접수ID": "AS-2601-005"}})

    kept = app.app_year_rows([
        {"프로젝트NO": "UJ2500001"},
        {"프로젝트NO": "UJ2600001"},
        {"정산ID": "JS-2607-001"},
    ])
    assert len(kept) == 2 and all("25" not in str(r) for r in kept), kept

    pub = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    sw = open(os.path.join(ROOT, "docs", "sw.js"), encoding="utf-8").read()
    assert 're.fullmatch(r"UJ26\\d{5}"' in pub, "배포 프로젝트코드에 2025가 섞일 수 있다"
    assert "A.app_project_result(c, r)" in pub, "UJ26에 연결된 2025 작업을 사본에서 거르지 않는다"
    assert 're.match(r"2026-' in pub, "미청구 배포에 2025가 섞일 수 있다"
    assert "function keep2026" in phone and "D = keep2026(await open(pin))" in phone
    assert "codeIs2026(k,r)" in phone, "구 사본의 프로젝트 자동조회 2025 2차 방어가 없다"
    assert "2026-only" in sw, "구 서비스워커가 2025 포함 사본을 계속 쓸 수 있다"

    import daily_brief as db
    brief = db.brief("2026-07-28", {
        "as": [
            {"접수ID": "AS-2512-001", "프로젝트NO": "UJ2500001",
             "접수일자": "2025-12-01", "작업완료일": "", "진행상태": "접수"},
            {"접수ID": "AS-2607-001", "프로젝트NO": "UJ2600001",
             "접수일자": "2026-07-28", "작업완료일": "", "진행상태": "접수"},
        ],
        "pm": [], "fw": [],
    })
    assert brief["돌발AS"]["신규접수"] == 1
    assert brief["돌발AS"]["완료일미기입"] == 0, "2025 미완료가 2026 브리핑에 섞였다"

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for need in ("const APP_YEAR = '2026'", "py.innerHTML = `<option>${APP_YEAR}</option>`",
                 "return +d === FOCUS_YEAR", "const y = APP_YEAR"):
        assert need in live, f"라이브 앱 2026 고정 누락: {need}"
    print("  [46] 앱 전 화면 2026년 전용 필터·구 사본 2차 방어 ✅")


def t47_back_nav():
    """[47] 뒤로가기 — 바로 전 화면 → 없으면 대시보드 → 대시보드에서 종료(사용자 지시 2026-07-28).

    시트(상세창)와 탭을 **한 스택**으로 다뤄야 한다. 따로 두면 시트를 닫는 뒤로가기와
    탭을 되돌리는 뒤로가기가 어긋난다. 그리고 되돌린 뒤에는 반드시 여분 항목을 다시
    쌓아야(navGuard) 다음 뒤로가기가 앱 밖으로 새지 않는다."""
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    # 화면 전환과 기록을 분리했는가 — applyView 는 기록을 안 남기고, show 는 남긴다
    assert "function applyView(" in live and "function show(v){" in live, "applyView/show 분리 누락"
    assert "navPush({t:'tab', from:from})" in live, "탭 이동이 뒤로가기 기록에 안 쌓인다"
    assert "navPush({t:'sheet'})" in live, "시트가 뒤로가기 기록에 안 쌓인다"

    # 홈 판정과 종료
    assert "curView() !== 'dash'" in live and "exitApp()" in live, "홈(대시보드) 폴백·종료 누락"
    assert "function exitApp()" in live and "window.close()" in live, "종료 처리 누락"

    # ★ 여분 항목이 없으면 첫 뒤로가기가 앱 밖으로 나간다
    assert "function navGuard()" in live, "navGuard 누락"
    assert live.count("navGuard()") >= 4, "되돌린 뒤 여분 항목을 다시 쌓지 않는 경로가 있다"

    # ★ history.go(-n) 은 popstate 를 한 번만 낸다 — n 만큼 세면 다음 뒤로가기를 먹는다
    assert "_navSkip += 1" in live, "history.go(-n) 의 popstate 를 n 번으로 세고 있다"
    assert "_navSkip += back" not in live, "go(-n) popstate 를 n 번으로 세는 코드가 남아 있다"

    # 시트를 통째로 닫아도 탭 기록은 남아야 한다(그래야 뒤로가기가 이전 탭으로 간다)
    assert "while(_nav.length && _nav[_nav.length-1].t === 'sheet')" in live, \
        "closeSheetAll 이 탭 기록까지 걷어낸다"
    # 옛 분리 구현의 흔적이 남아 있으면 두 스택이 다시 어긋난다
    assert "_sheetHist" not in live, "옛 _sheetHist 카운터가 남아 있다(스택 이원화)"
    print("  [47] 뒤로가기(이전 화면→대시보드→종료)·시트/탭 단일 스택 ✅")


def t48_excel_2026_stats_and_verified_completion():
    """[48] Excel statistics are 2026-only; completion requires exact evidence."""
    import stats_2026 as S

    old = "=COUNTIF('02_돌발AS접수'!$Q$5:$Q$744,\"작업완료\")"
    fixed = S.yearize_formula(old)
    assert "COUNTIFS" in fixed
    assert "'02_돌발AS접수'!$D$5:$D$744" in fixed
    assert "DATE(2026,1,1)" in fixed and "DATE(2027,1,1)" in fixed

    risk = "=COUNTIF('07_불일치누락현황'!$F$5:$F$304,\"*미배정*\")"
    fixed_risk = S.yearize_formula(risk)
    assert "'07_불일치누락현황'!$Q$5:$Q$304,2026" in fixed_risk

    document = "=COUNTIF('17_문서대조현황'!$G$5:$G$154,\"미작성\")"
    fixed_document = S.yearize_formula(document)
    assert "'17_문서대조현황'!$A$5:$A$154,\"JS-26*\"" in fixed_document

    top = S.top5_formula(1)
    top_index = S._top_index(1)
    assert "$P28" in top
    assert "$Q$5:$Q$304=2026" in top_index
    assert "$V$5:$V$304=\"포함\"" in top_index
    assert "UJ25" not in top and "2025" not in top

    helper = S.issue_year_formula(5)
    assert "'02_돌발AS접수'!$D$5:$D$744" in helper
    assert "'04_정기점검'!$D$5:$D$624" in helper
    assert "'06_거래서류청구수금'!$L$5:$L$154" in helper

    import complete_verified as C

    assert C._year("2026-07-27") == 2026
    assert C._year("2025-12-31") == 2025
    ready = {
        "작업완료일": "2026-07-27",
        "관리자검증상태": "일치",
        "완료보고서등록": "등록",
        "사진등록": "등록",
        "동영상등록": "",
        "비용구분": "유상",
        "ERP등록": "등록",
        "재방문여부": "아니오",
        "최초접수외추가작업": "",
        "추가작업확인상태": "",
    }
    assert C._as_completion_ready(lambda name: ready.get(name))
    ready["ERP등록"] = ""
    assert not C._as_completion_ready(lambda name: ready.get(name))
    print("  [48] Excel 2026-only statistics and evidence-gated completion OK")


def t49_exec_metric_drilldown_and_sheet_scroll(tmp):
    """[49] 대표보고 3·4절은 숫자와 동일한 원천행을 열고, 긴 목록은 끝까지 스크롤된다."""
    path = os.path.join(tmp, "exec_metric.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def sheet(name, headers, rows):
        ws = wb.create_sheet(name)
        for _ in range(3):
            ws.append([])
        ws.append(headers)
        for row in rows:
            ws.append(row)

    sheet("06_거래서류청구수금", [
        "정산ID", "원천업무ID", "프로젝트NO", "캠프명", "작업완료일(자동)", "업무구분", "담당자",
        "거래명세서발행일", "거래명세서합계", "세금계산서발행일", "세금계산서합계",
        "입금일", "입금액", "미청구액", "미수금액", "작업대비거래명세서차액",
        "문제내용", "청구상태", "검증결과"
    ], [
        ["JS-2601", "AS-2601", "UJ260001", "합성1캠프", "2026-07-28", "돌발AS", "김기사",
         "2026-07-28", 11000, "2026-07-28", 11000, "2026-07-28", 11000,
         22000, 33000, -1000, "합성 문제", "청구중", "불일치"],
        ["JS-2602", "AS-2602", "UJ260002", "합성2캠프", "2026-07-28", "신규·납품·설치", "이기사",
         "", 0, "", 0, "", 0, 0, 0, 999, "", "", "불일치"],
        ["JS-2501", "AS-2501", "UJ250001", "과거캠프", "2025-07-28", "돌발AS", "구기사",
         "2026-07-28", 999999, "2026-07-28", 999999, "2026-07-28", 999999,
         999999, 999999, 999999, "2025 제외", "", "불일치"],
    ])
    sheet("02_돌발AS접수",
          ["접수ID", "프로젝트NO", "캠프명", "접수일자", "담당기사"],
          [["AS-2601", "UJ260001", "합성1캠프", "2026-07-27", "김기사"]])
    sheet("04_정기점검",
          ["점검ID", "프로젝트NO", "캠프명", "점검예정일", "담당기사"], [])
    sheet("07_불일치누락현황",
          ["업무기준연도(자동·숨김)", "최상위 업무키", "원천업무ID", "프로젝트NO", "문제상세", "조치상태"],
          [[2026, "UJ260001", "AS-2601", "UJ260001", "문제 A", "확인필요"],
           [2026, "UJ260001", "AS-2601", "UJ260001", "문제 B", "확인필요"],
           [2025, "UJ250001", "AS-2501", "UJ250001", "과거 문제", "확인필요"]])
    sheet("15_세금계산서관리",
          ["정산ID", "발행기한임박여부", "기한초과여부", "법정발행기한",
           "발행금액", "발행상태(자동)", "아리바청구상태"],
          [["JS-2601", "예", "예", "2026-07-28", 11000, "미발행", "등록대기"],
           ["JS-2501", "예", "예", "2025-07-28", 999999, "미발행", "등록대기"]])
    sheet("17_문서대조현황",
          ["정산ID", "경고내용", "우선순위", "PO상태", "거래명세서상태"],
          [["JS-2601", "문서 경고", "P1", "PO 발행대기", "미작성"],
           ["JS-2501", "과거 경고", "P1", "PO 발행대기", "미작성"]])
    wb.save(path)

    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    import app_server as A
    d = A.read_exec_details(path, "2026-07-28")
    assert d["청구액 (당일)"]["count"] == 1 and d["청구액 (당일)"]["amount"] == 11000
    assert d["세금계산서 발행액 (당일)"]["count"] == 1
    assert d["입금액 (당일)"]["count"] == 1
    assert d["잔여 미청구액"]["amount"] == 22000
    assert d["잔여 미수금액"]["amount"] == 33000
    assert d["작업금액 불일치 (현재)"]["count"] == 1, "신규납품·2025가 섞였다"
    assert d["문제 업무 건수(중복 제거)"]["count"] == 1
    assert d["문제 프로젝트 / 문제 행"]["count"] == 2
    assert d["세금계산서 기한 임박·초과"]["count"] == 2
    assert d["PO 미발행 · 확인필요"]["count"] == 1
    assert d["거래명세서 미작성"]["count"] == 1
    assert d["아리바 청구 미등록"]["count"] == 1
    row = d["잔여 미청구액"]["rows"][0]
    assert row["프로젝트NO"] == "UJ260001" and row["캠프명"] == "합성1캠프"
    assert row["종류"] == "settle" and row["레코드ID"] == "JS-2601"

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for token in ("function openExecMetric", "function metricToPng", "function metricOpenCall",
                  "function sheetSnapshot", "function setSheetContent", "sheetactions",
                  "scrollTop: $('sheetbody').scrollTop", "PageDown", "PageUp"):
        assert token in live, token + " 누락"
    assert "#sheetbody{flex:1 1 auto;min-height:0;overflow-y:auto" in live
    assert re.search(r"(?m)^\s*#slist\{display:none\}", live), "데스크톱 정산 목록 범위가 없다"
    assert not re.search(r"(?m)^\s*\.slist\{display:none\}", live), "상세 시트 목록까지 숨긴다"
    assert "const actions = body.querySelector('.actions.sticky')" in live
    assert "openRecord(${esc4(kind)},${esc4(id)},${esc4(prj)})" in live
    assert "window._board = Object.assign" not in live, "리스크 일부 목록이 전체 담당자 보드를 덮는다"
    assert "saveOrOpen(b, assigneeFileName())" in live, "공유 불가 시 이미지를 다시 렌더한다"
    print("  [49] 대표보고 숫자 원천행·정확 라우팅·캡처·모달 끝까지 스크롤 ✅")


def t50_stale_completion_drilldown_and_capture():
    """[50] 오래된 완료일 미기입 경고는 정확한 목록·원천행·캡처로 이어진다."""
    import daily_brief as D

    data = {
        "as": [{
            "접수ID": "AS-2606-001", "프로젝트NO": "UJ2606001", "캠프명": "합성캠프",
            "접수일자": "2026-06-01", "작업완료일": "", "진행상태": "접수",
            "담당기사": "김기사", "신청내용": "합성 완료일 확인",
        }],
        "pm": [], "fw": [],
    }
    b = D.brief("2026-07-29", data)
    assert b["돌발AS"]["완료일미기입"] == 1
    stale = b["완료일미기입목록"][0]
    assert stale["레코드ID"] == "AS-2606-001" and stale["레코드종류"] == "as"

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for token in ("function openStaleBrief(", "onclick=\"openStaleBrief()\"",
                  "window._briefMetric", "눌러서 ${stale.length}건 목록 보기",
                  "이미지 저장", "이미지로 전달", "r.담당자?' · '+r.담당자"):
        assert token in live, token + " 누락"
    assert ".bneed.actionable" in live and 'role="button"' in live

    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    pub = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    for token in ("function openStaleCloud(", "function staleToPng(", "function shareStaleCloud(",
                  "function saveStaleCloud(", 'id="sheetbody"', "프로젝트 번호를 누르면"):
        assert token in phone, "고정 주소 앱 " + token + " 누락"
    assert ".sheetbody{flex:1 1 auto;min-height:0;overflow-y:auto" in phone, \
        "고정 주소 목록을 끝까지 스크롤할 수 없다"
    assert '"완료일미기입목록": _b.get("완료일미기입목록", [])' in pub, \
        "암호화 사본에 완료일 누락 원천 목록이 없다"
    sw = open(os.path.join(ROOT, "docs", "sw.js"), encoding="utf-8").read()
    assert "csos-v8-brief-list-2026-only" in sw, \
        "설치형 휴대폰 앱이 이전 화면 캐시를 계속 쥘 수 있다"
    print("  [50] 오래된 완료일 미기입 목록·정확 라우팅·담당자 캡처 ✅")


def t51_manual_daily_activity():
    """[51] 프로젝트NO 없는 대표 접수와 택배 발송 처리도 당일 업무에서 빠지지 않는다."""
    import daily_brief as D

    data = {
        "as": [{
            "접수ID": "", "프로젝트NO": "", "캠프명": "GWJ1 M_순천1",
            "접수일자": "2026-07-27", "작업완료일": "", "진행상태": "접수",
            "담당기사": "김필우", "신청내용": "리프트 리모콘 작동 안함",
        }],
        "pm": [], "fw": [],
        "events": [{
            "날짜": "2026-07-28", "접수일": "2026-07-27", "캠프명": "GWJ1 M_순천1",
            "게시자": "유수비 대표", "처리자": "류지영 매니저",
            "신청내용": "리프트 리모콘 작동 안함", "처리내용": "리모컨 택배 발송 완료",
            "상태": "택배 발송 완료",
        }],
    }
    b = D.brief("2026-07-28", data)
    assert b["돌발AS"]["업무처리"] == 1
    assert b["당일처리목록"][0]["캠프명"] == "GWJ1 M_순천1"
    assert b["당일처리목록"][0]["게시자"] == "유수비 대표"
    assert b["당일처리목록"][0]["무엇"] == "리모컨 택배 발송 완료"
    text = D.text(b)
    assert "당일 업무 처리 1건" in text and "류지영 매니저" in text

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for token in ("BRIEF['당일처리목록']", "업무 처리 ${activityL.length}",
                  "bCard(x,'activity')", "현장 AS 완료와 별도"):
        assert token in live, token + " 누락"
    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    pub = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    assert "function renderBrief(" in phone and "D.brief || {}" in phone
    assert "현장 AS 완료와 별도" in phone
    assert '"당일처리목록": _b.get("당일처리목록", [])' in pub, \
        "암호화 사본에 유수비 게시·류지영 택배 처리 실적이 없다"
    print("  [51] 유수비 대표 접수·류지영 택배 발송을 당일 업무 처리로 포함 ✅")


def t52_data_status():
    """[52] 자료현황 한 장 — 같은 질문을 매번 다시 세지 않게(사용자 지시 2026-07-29).

    ★ 이 장이 **느리면 아무도 안 본다.** Z: 폴더 2만 개 순회 같은 무거운 일은 여기서
      다시 하지 않고, 앞 단계가 남긴 리포트에서 숫자만 읽는다. 그래서 daily_run 에서
      대조들이 **끝난 뒤에** 와야 한다.
    ★ 앱 [기록] 탭에 뜨지 않으면 만든 의미가 없다 — 파일명이 서버 목록과 맞아야 한다."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import data_status as D

    # 리포트에서 숫자만 긁는 함수가 실제로 동작하는가(없으면 조용히 빈 dict)
    with tempfile.TemporaryDirectory() as tmp:
        old = os.environ.get("COUPANG_REPORT_DIR")
        os.environ["COUPANG_REPORT_DIR"] = tmp
        try:
            import importlib
            importlib.reload(D)
            open(os.path.join(tmp, "입금현황.md"), "w", encoding="utf-8").write(
                "# 쿠팡 입금 현황\n\n- 입금 70건 · 합계 1,106,167,980원 · 2026-04-13 ~ 2026-07-27\n")
            got = D.from_report("입금현황.md", ("건수", r"입금\s*([\d,]+)건"),
                                ("합계", r"합계\s*([\d,]+)원"))
            assert got.get("건수") == 70 and got.get("합계") == 1106167980, got
            assert D.from_report("없는파일.md", ("x", r"(\d+)")) == {}, "없는 리포트에서 죽으면 안 된다"
        finally:
            if old is None:
                os.environ.pop("COUPANG_REPORT_DIR", None)
            else:
                os.environ["COUPANG_REPORT_DIR"] = old
            importlib.reload(D)

    # 무거운 재계산을 하지 않는다 — Z: 폴더를 직접 훑는 코드가 있으면 안 된다
    src = open(os.path.join(ROOT, "data_status.py"), encoding="utf-8").read()
    assert "os.walk" not in src, "자료현황이 폴더를 직접 순회한다 — 느려서 아무도 안 보게 된다"
    assert "from_report(" in src, "앞 단계 리포트를 재활용하지 않는다"

    # 앱 [기록] 탭에 뜨는가 (파일명이 서버 목록과 맞아야 한다)
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert '("자료현황.md", "자료현황")' in srv, "앱 리포트 목록에 자료현황이 없다"
    # daily_run 이 매일 갱신하는가
    dr = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "data_status.py" in dr, "daily_run 에 자료현황 갱신 단계가 없다"
    assert dr.index("findings_export.py") < dr.index("data_status.py"), \
        "자료현황이 대조·리포트보다 먼저 돌면 지난 숫자를 읽는다"
    print("  [52] 자료현황 한 장(리포트 재활용·앱 노출·daily_run 순서) ✅")


def t53_session_handoff():
    """[53] 세션이 갑자기 끊겨도 다음 세션이 이어받는다(사용자 지시 2026-07-29).

    ★ 종료 체크리스트는 **끝낼 시간이 있을 때만** 지켜진다. 컨텍스트가 차거나 크레딧이
      끊기면 그럴 기회가 없고, 그때 점유·큐·임시파일이 방치된다.
    ★ 죽은 점유는 **프로세스 생사**로 판정한다 — 시간(45분)만 보면 그동안 원장이 잠긴다.
      ai_claim 은 `at` 을 **에포크 초(float)** 로 적는다(ISO 문자열 아님) — 이걸 틀리면
      경과 시간이 '?' 로 나와 죽은 잠금을 못 가려낸다(실제로 처음에 그랬다).
    ★ 이 도구는 **아무것도 고치지 않는다.** 상대 AI 가 일하는 중일 수 있어 자동으로
      점유를 풀거나 큐를 반영하면 가로채게 된다."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import session_handoff as H

    # 에포크 초를 제대로 읽는가 + 죽은 프로세스면 시간과 무관하게 잔재로 보는가
    import time as _t
    fake = {"ledger": {"who": "codex", "why": "테스트", "at": _t.time() - 120, "pid": 999999}}
    orig = H.claims.__globals__.get("ai_claim")
    class _Stub:
        @staticmethod
        def load():
            return fake
    H.claims.__globals__["ai_claim"] = _Stub
    _s.modules["ai_claim"] = _Stub
    try:
        got = H.claims()
        assert got and got[0]["mins"] == 2, "에포크 초 파싱 실패(ISO 로 읽고 있다): %s" % got
        assert got[0]["stale"] is True, "죽은 프로세스인데 잔재로 안 본다 — 원장이 45분 잠긴다"
        bl = H.blockers({"큐잔량": 0, "임시파일": [], "점유": got, "미푸시": []})
        assert any("잔재" in w for w, _ in bl), bl
        assert any("--free ledger" in c for _, c in bl), "풀 명령을 알려주지 않는다"
        # 살아 있는 점유는 잔재가 아니다 — 상대가 일하는 중이다
        fake["ledger"]["pid"] = os.getpid()
        assert H.claims()[0]["stale"] is False, "살아 있는 점유를 잔재로 본다 — 상대 작업을 가로챈다"
    finally:
        if orig is not None:
            H.claims.__globals__["ai_claim"] = orig
        _s.modules.pop("ai_claim", None)

    # 큐·임시파일도 막힌 것으로 잡는가
    bl = H.blockers({"큐잔량": 7, "임시파일": ["x.tmp.xlsx"], "점유": [], "미푸시": ["a"]})
    kinds = " ".join(w for w, _ in bl)
    assert "입력 큐" in kinds and "임시파일" in kinds and "푸시" in kinds, kinds

    # 고치지 않는다 — 자동 해제·자동 반영 코드가 있으면 안 된다
    src = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    body = src.split("if __name__")[0]
    assert "ai_claim.free" not in body and "queue_add" not in body, \
        "세션인계가 스스로 고치고 있다 — 상대 AI 작업을 가로챈다"

    # 워치독이 30분마다 남기는가 · 시작 체크리스트 0번인가
    wd = open(os.path.join(ROOT, "watchdog.py"), encoding="utf-8").read()
    assert "snapshot_handoff" in wd and "session_handoff.py" in wd, "워치독이 스냅샷을 안 남긴다"
    cm = open(os.path.join(os.path.dirname(ROOT), "CLAUDE.md"), encoding="utf-8").read()
    assert "session_handoff.py --check" in cm, "시작 체크리스트에 없다 — 아무도 안 읽는다"
    assert cm.index("session_handoff.py --check") < cm.index("AGENTS.md` 전체 읽기"), \
        "세션인계가 체크리스트 맨 앞이 아니다"
    print("  [53] 세션 인계(죽은 점유 판정·큐/임시파일·워치독 스냅샷·체크리스트 0번) ✅")


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
        t42_first_empty_row(tmp)
        t43_receipt_fill(tmp)
        t45_cloud_queue_and_erp_documents(tmp)
        t49_exec_metric_drilldown_and_sheet_scroll(tmp)
    t50_stale_completion_drilldown_and_capture()
    t51_manual_daily_activity()
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
    t24_reorder_safety()
    t25_attachments()
    t26_mobile()
    t27_po()
    t28_resolve()
    t29_cloud()
    t30_dns_and_versions()
    t31_tech()
    t32_band_sheet()
    t33_unbilled_banner()
    t34_capture_and_no_send()
    t35_confirm_evidence()
    t36_mobile_input()
    t37_band_coverage()
    t38_daily_brief()
    t40_claim_enforced()
    t41_dates_explicit()
    t44_zscan()
    t46_app_2026_only()
    t47_back_nav()
    t52_data_status()
    t53_session_handoff()
    t48_excel_2026_stats_and_verified_completion()
    t39_realtime_monitor()
    t6_webapp()
    print("ALL GREEN — 실작업 진행 가능")
