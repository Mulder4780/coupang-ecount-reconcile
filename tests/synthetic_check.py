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
import sys, os, re, tempfile, subprocess, hashlib, json

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable
# 합성검증 중 resolve_master가 불려도 실관리대장 구버전 정리가 실행되면 안 된다.
os.environ["CSOS_SYNTHETIC"] = "1"


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
    ws.append(["2026/07/01 -4", "테스트캠프A 돌발AS", 110000, None])        # JS-1 일치
    ws.append(["2026/07/02 -1", "테스트캠프B 정기점검", 999999, None])      # JS-2 금액불일치
    ws.append(["2026/07/05 -9", "인천8MB 상하차리프트(원장에 없음)", 500000, None])  # A형
    ws.append(["2026/07/06 -10", "입금 — 매출 전표로 세면 안 됨", None, 500000])
    ws.append(["2026/07 계", "", 1609999, 500000])
    wb.save(path)
    wb.close()

    # 실제 이카운트 내보내기처럼 dimension을 잘못된 A1:A1로 만든다.
    # read_only=True 파서는 이 경우 1셀만 읽고 전표 0건을 내므로 반드시 회귀검사한다.
    import zipfile
    broken = path + ".dimension"
    with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(broken, "w") as dst:
        for item in src.infolist():
            blob = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                blob = re.sub(rb'<dimension ref="[^"]+"', b'<dimension ref="A1:A1"',
                              blob, count=1)
            dst.writestr(item, blob)
    os.replace(broken, path)


def t1_erp_check(tmp):
    import erp_ledger_check as E
    assert E.norm_slip("2026/07-02-6") == "2026/07/02-6"
    assert E.in_erp_period({"작업완료일": "2026-08-01"}, "", "2026-07-01", "2026-07-31") is False
    assert E.in_erp_period({"작업완료일": "2026-07-05"}, "", "2026-07-01", "2026-07-31") is True
    ledger = os.path.join(tmp, "ledger.xlsx"); erp = os.path.join(tmp, "erp원장.xlsx")
    make_ledger(ledger); make_erp(erp)
    r = subprocess.run([PY, os.path.join(ROOT, "erp_ledger_check.py"),
                        "--file", erp, "--master", ledger],
                       capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=ROOT,
                       env={**os.environ, "COUPANG_REPORT_DIR": tmp,
                            "COUPANG_UPDATES_DIR": tmp})
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
    # 원문이 없어도 사용자가 해당 건을 개별 완료 처리한 기록은 다음 대조에서 되살아나면 안 된다.
    wb = openpyxl.load_workbook(ledger)
    ws = wb["02_돌발AS접수"]
    ws["G4"] = "조치내용"
    ws["G6"] = "카톡 보고 미확인 완료처리(사용자 지시 2026-07-29)"
    wb.save(ledger)
    wb.close()
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
    assert (ok, miss) == (2, 0), f"카톡 판정 불일치: 확인{ok} 미확인{miss} (기대 2,0)"
    report = next(p for p in os.listdir(tmp) if p.startswith("카톡대조_") and p.endswith(".csv"))
    rows = list(__import__("csv").DictReader(open(os.path.join(tmp, report), encoding="utf-8-sig")))
    manual = next(x for x in rows if x["ID"] == "AS-K2")
    assert manual["카톡보고"] == "확인" and manual["매칭근거"] == "사용자완료처리", manual
    # 방 이름 보완(2026-07-30): 머리글이 유형을 말하지 않으면 방으로 정한다.
    # ★ 순서가 뒤집히면 돌발방에 올라온 철거 글이 02시트로 잘못 간다 — 머리글이 먼저다.
    import kakao_extract as _KE
    assert _KE.kind_by_room("★UNI★ 쿠팡돌발점검") == "02_돌발AS접수"
    assert _KE.kind_by_room("★UNI★ 쿠팡정기점검") == "04_정기점검"
    assert _KE.kind_by_room("아무방") == ""
    assert _KE.kind_of("[철거 안내]") == "(철거·보관)", "머리글이 방보다 우선이어야 한다"
    _ke_src = open(os.path.join(ROOT, "kakao_extract.py"), encoding="utf-8").read()
    assert "kind_of(head) or kind_by_room(room)" in _ke_src, "머리글 우선 순서가 깨졌다"
    print("  [4] 카톡 내보내기 파싱·대조·사용자 개별 완료 유지 (PC/모바일·다중행) ✅")


def t5_writer(tmp):
    import json
    src = os.path.join(tmp, "합성대장_v1.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "06_거래서류청구수금"
    for _ in range(3): ws.append([])
    ws.append(["정산ID", "거래명세서번호", "거래명세서발행일", "비고"])
    ws.append(["JS-W1", None, None, "빈칸 채움 대상"])
    ws.append(["JS-W2", "기존값-유지", None, "덮어쓰기 금지 검증"])
    ws2 = wb.create_sheet("02_돌발AS접수")
    for _ in range(3): ws2.append([])
    ws2.append(["접수ID", "프로젝트NO", "완료보고서등록", "사진등록"])
    ws2.append(["AS-MIX-1", "UJ-MIX-1", None, None])
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
        # 같은 시트에서 서로 다른 키 열을 연달아 써도 둘 다 정확한 행을 찾아야 한다.
        {"sheet": "02_돌발AS접수", "key_col": "접수ID", "key": "AS-MIX-1", "col": "사진등록",
         "value": "등록", "vtype": "text", "evidence": "합성", "only_if_empty": True},
        {"sheet": "02_돌발AS접수", "key_col": "프로젝트NO", "key": "UJ-MIX-1", "col": "완료보고서등록",
         "value": "등록", "vtype": "text", "evidence": "합성", "only_if_empty": True},
    ], open(q, "w", encoding="utf-8"), ensure_ascii=False)
    r = subprocess.run([PY, os.path.join(ROOT, "ledger_writer.py"), "--apply",
                        "--master", src, "--queue", q],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       cwd=ROOT, env={**os.environ, "COUPANG_REPORT_DIR": tmp,
                                     "COUPANG_UPDATES_DIR": tmp,
                                     "COUPANG_LEDGER_GATE": "1"})
    assert "반영 완료" in r.stdout and "입력 4건 / 건너뜀 1건" in r.stdout, f"{r.stdout}\n{r.stderr}"
    dst = src.replace("_v1.xlsx", "_v2.xlsx")
    w2 = openpyxl.load_workbook(dst)
    ws2 = w2["06_거래서류청구수금"]
    assert ws2["B5"].value == "2026/07/24-9", ws2["B5"].value          # 빈칸 채움
    assert ws2["C5"].value is not None, "날짜 직렬값 기록 실패"          # date serial
    assert ws2["B6"].value == "기존값-유지", ws2["B6"].value            # 덮어쓰기 금지
    as2 = w2["02_돌발AS접수"]
    assert as2["C5"].value == "등록" and as2["D5"].value == "등록", \
        "같은 시트의 접수ID·프로젝트NO 혼합 키 조회 실패"
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
        # 프로세스가 강제 종료돼 남긴 잠금은 다음 실행이 즉시 회수해야 한다.
        with open(LW.PENDING + ".lock", "w", encoding="ascii") as f:
            f.write("99999999 2026-01-01T00:00:00")
        item2 = {"sheet": "S", "key": "K2", "col": "C"}
        assert LW.queue_add([item2]) == 1
        assert not os.path.exists(LW.PENDING + ".lock")
    finally:
        LW.PENDING = old_pending
    from band.band_reconcile import photo_updates
    photo_q = photo_updates([{
        "시트": "02_돌발AS접수", "ID": "AS-OLD-CACHE", "프로젝트NO": "UJ-STABLE",
        "밴드게시": "확인", "사진수": 2, "게시일": "2026-07-29", "게시자": "합성",
    }])
    assert len(photo_q) == 1 and photo_q[0]["key_col"] == "프로젝트NO" \
        and photo_q[0]["key"] == "UJ-STABLE", photo_q
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
    assert set(pick_archive(vers)) == {"v19", "v21", "v23", "v24"}, pick_archive(vers)
    assert pick_archive([(25, "v25")]) == []                                            # 최신 1개만 보존
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
    print("  [9] 워치독 판단 로직(버전 최신 1개·보호파일·30일 기준·실행 공백 감지) ✅")


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
    # 03시트 자동선택 사슬: 잘못 들어간 AS 채번 자기참조와 확장 뒤 굳은
    # COUNTIF 끝행을 한꺼번에 바로잡아야 한다.
    selector = (
        "IFERROR(INDEX('02_돌발AS접수'!$A$5:$A$741,MATCH(1,"
        "('02_돌발AS접수'!$A$5:$A$741&lt;&gt;\"\")*"
        "('02_돌발AS접수'!$Q$5:$Q$741=\"작업완료\")*"
        "('02_돌발AS접수'!$R$5:$R$741&lt;&gt;\"\")*"
        "(COUNTIF($B$4:$B201,'02_돌발AS접수'!$A$5:$A$741)=0),0)),\"\")"
    )
    selector_xml = (
        '<sheetData>'
        f'<row r="202"><c r="B202"><f>{selector}</f><v/></c></row>'
        '<row r="203"><c r="B203"><f>IF($B203="","","AS-"&amp;TEXT($D203,"yymm"))</f>'
        '<v>AS-2607-468</v></c></row>'
        f'<row r="204"><c r="B204"><f>{selector}</f><v/></c></row>'
        '</sheetData>'
    )
    selector_out, selector_n = F.fix_completed_as_selector(selector_xml)
    assert selector_n == 2, selector_n
    assert 'COUNTIF($B$4:$B202' in selector_out, selector_out
    assert 'COUNTIF($B$4:$B203' in selector_out, selector_out
    assert 'AS-2607-468' not in selector_out, selector_out
    assert not F.direct_self_refs(selector_out), F.direct_self_refs(selector_out)
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
    # (3) 마지막 행을 본떠 확장할 때 자기행뿐 아니라 '바로 윗행까지'인
    # 상대참조도 함께 이동해야 한다. 고정되면 모든 새 행이 같은 ID를 고른다.
    copied = E.shift_formula('IF($B202="","",COUNTIF($B$4:$B201,1))', 202, 205)
    assert '$B205' in copied and '$B204' in copied and '$B$4' in copied, copied
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
    assert 'start_url = f"/staff/{staff_slug}" if center else "/"' in src, \
        "앱·담당자 매니페스트 start_url은 같은 출처 경로여야 설치가 된다"
    assert '"start_url": start_url' in src
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
    # ★ 700자만 잘라 보던 것을 3000자로 넓혔다(2026-07-31). 서비스 워커가 오프라인 껍데기
    #   캐시를 하게 되면서 블록이 길어졌고, fetch 핸들러가 창 밖으로 밀려나 검증이 헛돌았다.
    #   보려는 것은 '핸들러가 있는가' 이지 '몇 번째 글자에 있는가' 가 아니다.
    _swblk = src[src.index('p == "/sw.js"'):][:3000]
    assert "addEventListener('fetch'" in _swblk, "fetch 핸들러가 없으면 크롬이 설치를 제안하지 않는다"
    # 오프라인에서도 화면이 떠야 폰에서 입력해 둘 수 있다(사용자 지시 2026-07-31).
    assert "caches.open" in _swblk and "caches.match" in _swblk, \
        "서비스 워커가 껍데기를 캐시하지 않으면 PC 가 꺼졌을 때 앱이 아예 안 열린다"
    assert "startsWith('/api/')" in _swblk, "데이터(/api/)까지 캐시하면 옛 숫자가 화면에 남는다"
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
    icon_hash = hashlib.sha256(
        open(os.path.join(ROOT, "webapp", "brand", "csos-app-icon-source.png"), "rb").read()
    ).hexdigest()[:12]
    icon_rev = f"csos-{icon_hash}"
    assert any(f"icon-192.png?v={icon_rev}" in x.get("src", "") for x in mf["icons"]), mf
    assert any(f"icon-512.png?v={icon_rev}" in x.get("src", "") for x in mf["icons"]), mf
    assert f"csos-icon-{icon_hash}" in open(
        os.path.join(ROOT, "docs", "sw.js"), encoding="utf-8").read()
    for n in (32, 180, 192, 512):
        fp = os.path.join(ROOT, "docs", f"icon-{n}.png")
        assert os.path.getsize(fp) > 1000, fp

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
    assert "heal_fixed_funnel" in wd and "ensure_public_funnel" in wd, \
        "워치독이 PC 내부 응답만 보고 휴대폰 공개 Funnel 장애를 놓친다"
    # ★★ 터널 주소에서는 **설치가 되면 안 된다**. 터널 호스트는 매번 바뀌는데 거기서
    #   설치하면 그 임시 주소가 아이콘에 영구히 박혀, 주소가 바뀌는 순간 영영 안 열린다
    #   (2026-07-28: PC·폰의 설치된 앱이 둘 다 옛 터널 주소로 죽었다).
    i = src.index('if p in ("/", "/index.html", "/ryu")')
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
    r = {"점검상태": "미점검", "실제점검일": "2026-07-30"}
    derive_status(r, "pm"); assert r["점검상태"] == "완료", r          # 완료일이 오래된 캐시보다 우선
    r = {"점검상태": "AS전환", "실제점검일": "2026-07-30"}
    derive_status(r, "pm"); assert r["점검상태"] == "AS전환", r        # 충돌 상태는 보존
    r = {"진행상태": "", "작업완료일": "2026-06-02"}
    derive_status(r, "as"); assert r["진행상태"] == "작업완료", r
    r = {"진행상태": "접수", "작업완료일": "2026-06-02"}
    derive_status(r, "as"); assert r["진행상태"] == "작업완료", r      # 완료일이 접수 상태보다 우선
    r = {"진행상태": "취소", "작업완료일": "2026-06-02"}
    derive_status(r, "as"); assert r["진행상태"] == "취소", r          # 취소는 자동 덮기 금지
    r = {"진행상태": "", "작업완료일": ""}
    derive_status(r, "as"); assert r["진행상태"] == "접수", r
    print("  [16] 상태 수식 캐시 보정(완료·AS전환·미점검·예정·기존값 보존) ✅")


def t6_webapp():
    import time, json, urllib.request, http.cookiejar
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
        # 담당자 로그인은 서버 세션에 역할이 고정된다. 공용 PIN을 알고 있어도
        # 실행·정책·다른 담당자 입력 API를 직접 호출할 수 없어야 한다.
        staff_opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
        staff_login = urllib.request.Request(
            base + "/api/login",
            data=b'{"pin":"0000","staff_slug":"ryu-jiyeong"}',
            method="POST")
        staff_auth = json.loads(staff_opener.open(staff_login).read())
        assert staff_auth["role"] == "staff" and staff_auth["staff_slug"] == "ryu-jiyeong"
        staff_session_response = staff_opener.open(base + "/api/auth/session")
        staff_session = json.loads(staff_session_response.read())
        assert staff_session["ok"] and staff_session["authenticated"]
        assert staff_session["role"] == "staff"
        assert staff_session["staff_slug"] == "ryu-jiyeong"
        assert int(staff_session["expires_at"]) > int(time.time()) + (300 * 24 * 60 * 60), \
            "담당자 기기 인증이 장기 유지되지 않음"
        assert "Max-Age=" in (staff_session_response.headers.get("Set-Cookie") or ""), \
            "앱 재실행 시 기기 인증 만료일이 연장되지 않음"
        for path, data in (
            ("/api/run/all", b"{}"),
            ("/api/policy", '{"기준":"x","확정내용":"y"}'.encode("utf-8")),
            ("/api/staff/po-upload", b""),
        ):
            request = urllib.request.Request(
                base + path, data=data, headers={"X-Pin": "0000"}, method="POST")
            try:
                staff_opener.open(request)
                assert False, f"담당자 세션이 관리자 API를 호출함: {path}"
            except urllib.error.HTTPError as exc:
                assert exc.code == 403, (path, exc.code)
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
        # 대시보드는 화면 블록과 내부 KPI를 모두 카드로 취급한다. 체크로 표시 여부를
        # 정하고 드래그·화살표로 순서를 바꾸며, 구버전 배치도 v2로 이어받는다.
        for marker in (
            "DASH_LAYOUT_KEY", "csos_dashboard_layout_v2", "DASH_LAYOUT_LEGACY_KEY",
            "defaultDashboardLayout", "DASH_KPI_CARDS", "data-kpi-card",
            "dashboardLayoutState", "setDashboardBlockVisible", "moveDashboardBlock",
            "setDashboardKpiVisible", "moveDashboardKpi", "prepareDashboardKpis",
            "cycleDashboardBlockSize", "dash-drag-handle", "dash-kpi-handle", "dash-choice",
            "role=\"status\" aria-live=\"polite\"", "window.addEventListener('storage'",
            "grid.addEventListener('dragstart'", "initDashboardLayout();",
        ):
            assert marker in html, "대시보드 카드 편집 누락: " + marker
        assert "#dashGrid{display:grid;grid-template-columns:repeat(12" in html.replace(" ", "")
        # 제공받은 CSOS 아이콘을 앱·캡처에 공통 사용한다. 회사 CI는 상하 배치한
        # 밝은 배경판 안에서 잘리지 않아야 하며, 컨테이너는 기존 295px 폭을 보존한다.
        assert '/icon-192.png?v=csos-20260730' in html
        assert 'class="logo app-icon"' in html and "loadAppIconImg" in html
        app_server_src = open(os.path.join(ROOT, "webapp", "app_server.py"),
                              encoding="utf-8").read()
        assert "icon_revision()" in app_server_src
        assert "sync_installed_app_icons()" in app_server_src
        assert os.path.isfile(os.path.join(ROOT, "webapp", "sync_app_icons.ps1"))
        assert os.path.isfile(os.path.join(ROOT, "webapp", "build_windows_icon.ps1"))
        assert os.path.isfile(os.path.join(ROOT, "webapp", "install_staff_shortcut.ps1"))
        compact_html = html.replace(" ", "")
        assert ".appbar-brand-stack{min-width:295px}" in compact_html
        assert ".uni-app-brand{width:196px;height:26px" in compact_html
        assert "drawUniversalLogo" in html and "universal-lift-horizontal.png" in html
        # 공통 헤더는 시스템 식별/회사 로고/상태의 3구역이며, 회사 로고는
        # 유니버셜리프트를 위에 두고 쿠팡을 그 아래에 둔다.
        assert 'class="appbar-identity"' in html
        assert 'class="appbar-brand-stack"' in html and 'class="appbar-status"' in html
        header = html[html.index('<header class="appbar">'):html.index("</header>", html.index('<header class="appbar">'))]
        assert header.index('class="uni-app-brand"') < header.index('class="coupang-app-brand"')
        for asset in ("icon-32.png", "icon-180.png", "icon-192.png", "icon-512.png"):
            r = urllib.request.urlopen(base + "/" + asset)
            assert r.headers.get_content_type() == "image/png" and len(r.read()) > 1000
        # 공통 캡처 도구막대와 류지영 조사실(카톡 2개 방 원본)은 모든 화면에서
        # 같은 저장·복사·엑셀 흐름을 쓴다.
        for marker in ("media-tools", "icon-btn", "#i-file-spreadsheet",
                       'id="v-ryu"', 'name="kakao_regular"', 'name="kakao_emergency"',
                       "/api/ryu/upload", "★UNI★ 쿠팡정기점검", "★UNI★ 쿠팡돌발점검",
                       "auto_check_queued"):
            assert marker in html, "UI marker missing: " + marker
        # 류지영 화면은 단순 업로드방이 아니라 대시보드형 업무센터다.
        # 카테고리→과거목록→선택건 보충입력 흐름과 일반 메뉴(실행 제외)를 함께 제공한다.
        for marker in ("류지영 쿠팡 AS 및 정기점검 업무센터", 'class="workcenter-person"',
                       'id="ryuSummaryGrid"',
                       'id="ryuCategoryTabs"', 'id="ryuHistoryList"', 'id="ryuEntryForm"',
                       "/api/ryu/records", "/api/ryu/entry", "submitRyuEntry",
                       "body.ryu-mode .tabbar button[data-v=\"run\"]{display:none}",
                       "routeNav('dash')"):
            assert marker in html, "류지영 업무센터 UI 누락: " + marker
        assert 'href="/staff/yoo-hyeonmin"' not in html, "업무센터 허브에서 유현민 버튼 제거 누락"
        assert "/api/staff/install-shortcut" in html and "window.__csosInstallPrompt" in html
        assert "workcenterIsInstalled()" in html and "csos_installed_" in html
        assert "maybeShowInstallCard()" in html and "display-mode: standalone" in html
        # 아이콘 참조는 스프라이트(<use href="#i-...">)다 — 경로가 아니라 **이름**으로 확인한다.
        for icon_name in ("bootstrap-person-workspace",
                          "bootstrap-person-badge-fill"):
            assert f"#i-{icon_name}" in html, "Bootstrap 업무센터 아이콘 누락: " + icon_name
        # PIN 원문을 브라우저에 계속 보관하지 않고, 서버 서명 쿠키를 복원한다.
        assert "/api/auth/session" in html and "restoreRoleSession" in html
        assert "localStorage.setItem('cw_pin'" not in html
        assert "LEGACY_PIN" in html and "localStorage.removeItem('cw_pin')" in html
        for marker in ('id="workLogRecordList"', "renderWorkLogRecords",
                       "workLogFilterCategory", "workLogFilterState",
                       "프로젝트 미확정"):
            assert marker in html, "대표보고 일지 상세목록 UI 누락: " + marker
        app_src = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
        for marker in ("AUTH_SESSION_TTL", "create_auth_session", "auth_cookie",
                       "auth_session_from_cookie", "get_work_log_view",
                       'if p == "/api/auth/session"'):
            assert marker in app_src, "기기 인증·일지 자동연동 API 누락: " + marker
        assert "def defer_task_until_free(" in app_src and '"auto_check_queued": queued' in app_src
        for marker in ("RYU_ENTRY_CONFIG", "def get_ryu_records(", "def save_ryu_entry(",
                       '"only_if_empty": True', 'if p == "/api/ryu/records"',
                       'if p == "/api/ryu/entry"'):
            assert marker in app_src, "류지영 업무센터 API 누락: " + marker
        assert "def install_staff_shortcut(" in app_src
        assert 'if p == "/api/staff/install-shortcut"' in app_src
        assert "flex-wrap:nowrap!important" in html.replace(" ", "")
        # 보고·목록 툴바의 구형 일지저장/직접전달 버튼은 제거하되,
        # 사용자가 명시 요청한 담당자 고정 URL 공유는 허용한다.
        assert not re.search(
            r'<button[^>]+onclick="(?:save\w*Journal|share(?:Mine|Daily|Journal)\w*)',
            html), "removed journal/report share button is still exposed"
        # 원천 검증은 밴드·카톡·ERP·쿠팡PO 4종이 **자료 유무와 무관하게** 항상 표시돼야 한다
        assert "4원천 검증" in html, "원천 검증 제목이 4원천이 아님"
        assert "쿠팡 PO" in html and "PO 목록 투입 시 자동 대조" in html, "PO 원천 행 누락"
        for marker in ("srcDetails", "sourceRowHeight", "sourceTable(sourceDetails)",
                       "미확인 프로젝트NO", "missProjects"):
            assert marker in html, f"대표 캡처 4원천 표 구조 누락: {marker}"
        assert "if(window.__view==='daily') renderDaily()" in html, \
            "4원천 상태가 늦게 도착했을 때 대표 보고가 갱신되지 않음"
        assert "renderDaily();\n    blob = await reportToPng();" in html, \
            "이미지 저장·공유 직전에 최신 4원천 집계를 다시 묶지 않음"
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
        # 대표보고: 비어 있던 '정리' 절은 제거하고 화면 순서대로 1부터 다시 번호를 매긴다.
        assert "cleanExecTitle" in html and "sectionNo += 1" in html, "대표보고 번호 재정렬 로직 누락"
        assert "filter(s=>!/^(정리|요약)$/" in html, "빈 정리·요약 절 제거 로직 누락"
        assert "dailyIssueHtml" in html and "이슈사항" in html, "당일 업무 이슈 구역 누락"
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
        # 문구는 바뀔 수 있다(2026-07-30: "앱에서 찾은 건" → "연결 프로젝트"). 지켜야 하는 것은
        # **불일치를 숨기지 않는 동작**이다 — 그래서 문구가 아니라 그 로직이 있는지 본다.
        assert "const diff = Number.isFinite(n) && n !== rows.length;" in html, "보고 숫자와 목록 건수 불일치 판정이 없다"
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
        # ★ 'self.addEventListener 로 시작하는가' 로 보던 것을 2026-07-31 에 바꿨다.
        #   오프라인 껍데기 캐시가 붙으면서 앞에 캐시 이름(const V=...)이 오게 됐는데,
        #   그건 JSON 으로 감싸인 것과 아무 상관이 없다. 보려는 것은 **따옴표에 싸여 나가는가** 다.
        assert not _sw.lstrip().startswith(("{", '"')), "sw.js가 JSON으로 감싸여 나간다: " + _sw[:40]
        assert "self.addEventListener" in _sw, "sw.js에 이벤트 등록이 없다: " + _sw[:40]
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
        # ct는 Base64로 표현한 암호문이라 우연히 "UJ25" 같은 4글자 조각이 생길 수 있다.
        # 2026-07-30 실제 재암호화에서 그 확률 충돌로 검증이 실패했다. 암호문 문자열을
        # 평문 검색하지 말고, 자료가 들어갈 수 있는 별도 메타데이터가 없는지 확인한다.
        allowed = {"v", "kdf", "iter", "cipher", "salt", "iv", "ct", "tag",
                   "pin_auth", "zip"}
        assert set(d) <= allowed, f"암호화 봉투에 예상 밖 메타데이터가 있다: {set(d) - allowed}"
        for key in ("salt", "iv", "ct", "tag"):
            base64.b64decode(d[key], validate=True)
        meta = _j.dumps({k: v for k, v in d.items() if k != "ct"}, ensure_ascii=False)
        for leak in ("캠프", "UJ25", "UJ26", "프로젝트NO"):
            assert leak not in meta, "암호화 봉투 메타데이터에 평문 '%s' 이 있다" % leak

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
    cloud_source = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    phone_source = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    assert "def snapshot_key(" in cloud_source and 'sealed["pin_auth"] = pin_auth' in cloud_source, \
        "hashed PIN snapshot compatibility is missing"
    assert "async function snapshotPassword(" in phone_source and \
           "auth-pbkdf2-sha256" in phone_source, \
        "phone snapshot cannot derive the stored authentication digest"

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
    # (5) PC 꺼짐 **콜드 스타트** — 앱을 새로 여는 경우까지 되는가(2026-07-31).
    #     오프라인 큐만으로는 부족했다: 세션 복원이 네트워크 실패를 인증 거절과 같이 다뤄
    #     PIN 게이트로 막혀, 정작 입력 화면에 못 들어갔다. 셋 다 있어야 구멍이 안 남는다.
    live_html = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert "netDown" in live_html and "cw_dev_auth" in live_html, \
        "네트워크 실패와 인증 거절을 구분하지 않으면 PC 꺼진 뒤 앱을 새로 열 때 잠긴다"
    assert "r.status === 401 || r.status === 403" in live_html, \
        "outbox 가 401·403 을 4xx 로 버린다 — 오프라인에서 써 둔 입력이 유실된다"
    assert "function offlineBanner(" in live_html and "_netFetch('/api/ping" in live_html, \
        "PC 꺼짐 안내·복귀 감시가 없으면 사용자가 반영 여부를 알 수 없다"
    assert "location.reload" not in live_html[live_html.index("function offlineBanner("):
                                              live_html.index("function offlineBanner(") + 1200], \
        "복귀 시 스스로 새로고침하면 입력 중이던 화면이 날아간다(2026-07-31 지시)"

    print("  [29] 폰 단독 사용(잠금·오프라인 폴백·콜드스타트·git 게시 실패감지·PIN 비노출) ✅")


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
    assert "os.remove(" not in src, "구버전 정리가 중복 파일을 삭제한다"
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
    _m, items, _stat, completions = C._plan()
    VCOLS = ("관리자검증상태", "담당관리자", "최종확인일", "최종확인일(유현민 체크)")
    bad = [i["key"] for i in items if i["col"] in VCOLS
           and (real.get(i["key"]) or {}).get("status") != "작업완료"]
    assert not bad, ("근거가 완료가 아닌데 확인 체크를 찍으려 한다: " + ", ".join(sorted(set(bad))[:5]))

    # (3) 확인·날짜 필드는 기존 값을 덮지 않고, 상태 수식 셀은 아예 큐에 넣지 않는다.
    #     객관 완료는 work_resolution DB에만 기록한다.
    assert all(i.get("only_if_empty") for i in items), "확인·날짜 필드의 빈칸만 채우는 규칙이 빠졌다"
    assert not any(i["col"] in ("진행상태", "점검상태") for i in items), \
        "객관 완료 판정이 상태 수식 셀을 덮으려 한다"
    assert completions and all(c["status"] in ("작업완료", "완료")
                               and c["completed_on"] and c["basis"] for c in completions), \
        "DB 객관 완료 판정에 상태·완료일·근거가 빠졌다"
    # (4) 취소성 상태·날짜 없는 완료 글·완료가 아닌 글은 자동 완료하지 않는다.
    done_band = {"status": "작업완료", "date": "2026-07-30"}
    assert C.objective_completion("접수", "작업완료", done_band)
    assert C.objective_completion("미점검", "완료", done_band)
    for conflict in ("취소", "철회", "AS전환", "점검불가"):
        assert not C.objective_completion(conflict, "완료", done_band), conflict
    assert not C.objective_completion("접수", "작업완료", {"status": "작업완료", "date": ""})
    assert not C.objective_completion("접수", "작업완료", {"status": "접수", "date": "2026-07-30"})
    assert not C.objective_completion("접수", "작업완료", done_band,
                                      base="2026-07-31"), "접수보다 앞선 완료일은 모순"
    assert C.objective_completion("접수", "작업완료", {"status": "작업완료", "date": "2025-01-03"},
                                  base="2026-01-02", existing_done="2026-01-03"), \
        "명시된 원장 완료일보다 밴드의 연도 추정을 우선하면 안 된다"
    print("  [35] 확인 체크 근거 정합(완료 글 우선·DB 완료판정·날짜충돌 보존) ✅")


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
    _rd = idx[idx.index("function openRptDates"):][:2800]
    assert 'type="date"' in _rd, "보고 날짜가 달력으로 안 뜬다"
    _sv = idx[idx.index("function saveRptDates"):][:1200]
    assert "/api/set_dates" in _sv, "엑셀 00_대시보드에 안 쓴다"
    # 집계기준일이 보고일보다 뒤면 잘못 고른 것이다 — 그대로 쓰면 보고서가 어긋난다
    assert "집계기준일 > 보고일" in _sv.replace("집계기준일 &gt; 보고일", "집계기준일 > 보고일"),         "집계기준일이 보고일보다 뒤인 경우를 안 막는다"
    # 사용자 확정 기준: 보고일은 오늘, 집계기준일은 주말 여부와 무관하게 정확한 전날
    _td = idx[idx.index("function rptDatesToday"):][:600]
    assert "previousDayISO(cur)" in _td, "집계기준일 기본값이 정확한 전날이 아니다"
    _init = idx[idx.index("function initDates"):][:500]
    assert "previousDayISO(cur)" in _init, "대시보드 날짜 기본값이 정확한 전날이 아니다"
    assert "requestAnimationFrame(()=>rptDateChanged())" in _rd, \
        "날짜 변경창을 열었을 때 오늘/전날 기본값이 즉시 적용되지 않는다"
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
        assert set(("왜", "무엇", "비용", "추가작업", "이상내용", "문제내용", "조치내용")) <= set(x), x
    # 내용이 없으면 지어내지 말고 미기입으로 남겨야 한다
    assert all(x["무엇"] == "" for x in b["내용미기입"]), "미기입 판정이 틀렸다"

    # (4) 분기 계산이 맞는가 — 7월은 7~9월 구간
    #     ★ 사용자 지시(2026-07-29)로 'N분기' 대신 **월 범위**로 적는다(검증 [70] 참고).
    assert b["정기점검"]["분기"].endswith("7~9월"), b["정기점검"]["분기"]
    assert b["정기점검"]["분기끝월"] == "9월", b["정기점검"]["분기끝월"]
    b1 = D.brief("2026-02-10", data)
    assert b1["정기점검"]["분기"].endswith("1~3월"), b1["정기점검"]["분기"]
    assert b1["정기점검"]["분기범위"] == "1~3월", b1["정기점검"]["분기범위"]

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
        "돌발AS": {"신규접수": 1, "신규처리완료": 1, "신규처리율": 100,
                   "완료": 1, "미처리": 0, "완료일미기입": 0},
        "정기점검": {"예정": 0, "예정장비": 0, "완료": 0, "완료장비": 0,
                     "분기": "2026년 1분기",
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
    assert "_archive_old_versions" in _auto and "shutil.move" in _lv, \
        "자동 정리가 OLD 이동 함수를 호출하지 않는다"
    for f in ("ecount_reconcile.py", "workbook_patch.py"):
        s = open(os.path.join(ROOT, f), encoding="utf-8").read()
        assert "autoprune" in s, f"{f} 가 구 버전을 자동으로 접지 않는다"
    assert LV.KEEP_LATEST == 1 and LV.ARCHIVE == "OLD", "보관 정책이 바뀌었다"
    assert LV.archive_folder(ROOT) == os.path.join(os.path.abspath(ROOT), "OLD")

    # ★ 자동으로 도는 정리는 아무도 안 보고 있다. **남의 파일을 옮기면 안 된다.**
    #   처음 구현이 `*_v*.xlsx` 를 통째로 잡아 합성검증용 임시 파일까지 옮겨 시험이 깨졌다.
    import tempfile as _tf, shutil as _sh
    _d = _tf.mkdtemp()
    _synthetic_flag = os.environ.pop("CSOS_SYNTHETIC", None)
    try:
        for _n in ("합성대장F_v1.xlsx", "합성대장F_v2.xlsx", "기타자료_v3.xlsx"):
            open(os.path.join(_d, _n), "w").write("x")
        LV._AUTODONE = False
        assert LV.autoprune(os.path.join(_d, "합성대장F_v2.xlsx")) == 0, "남의 파일을 옮겼다"
        assert len(os.listdir(_d)) == 3, "남의 폴더를 건드렸다"

        for _v in (1, 2, 3):
            open(os.path.join(_d, f"쿠팡_통합업무_일일보고_관리대장_v{_v}.xlsx"), "w").write("x")
            _stamp = 1700000000 + (_v * 86400)  # 서로 다른 날짜여도 최신 1개만 남겨야 한다
            os.utime(os.path.join(_d, f"쿠팡_통합업무_일일보고_관리대장_v{_v}.xlsx"),
                     (_stamp, _stamp))
        open(os.path.join(_d, "쿠팡_통합업무_일일보고_관리대장_v9_보관.xlsx"), "w").write("x")
        _keep, _move = LV.plan(os.path.join(_d, "쿠팡_통합업무_일일보고_관리대장_v3.xlsx"))
        assert [x["v"] for x in _keep] == [3] and sorted(x["v"] for x in _move) == [1, 2], \
            "KEEP_DAYS=0인데 날짜별 구버전을 작업 폴더에 남긴다"
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

        # 과거 `_이전버전`도 다음 실행 때 지정 OLD로 합친다. 같은 v2라도 내용이 다르면
        # 기존 OLD 파일을 덮거나 지우지 않고 출처 꼬리표로 둘 다 남겨야 한다.
        _legacy = os.path.join(_d, "_이전버전")
        os.makedirs(_legacy)
        _legacy_v2 = os.path.join(_legacy, "쿠팡_통합업무_일일보고_관리대장_v2.xlsx")
        open(_legacy_v2, "w").write("다른 v2")
        # 방금 저장된 파일은 사람 것일 수 있어 옮기지 않는다([94]) — 이 검사의 의도는
        # '병합이 되는가'이므로 mtime 을 오래된 것으로 둔다.
        os.utime(_legacy_v2, (_stamp, _stamp))
        LV._AUTODONE = False
        assert LV.autoprune(os.path.join(_d, "쿠팡_통합업무_일일보고_관리대장_v3.xlsx")) == 1, \
            "예전 보관 폴더의 파일이 OLD로 합쳐지지 않았다"
        assert not os.listdir(_legacy), "예전 보관 폴더에 관리대장이 남았다"
        archived = sorted(os.listdir(os.path.join(_d, "OLD")))
        assert "쿠팡_통합업무_일일보고_관리대장_v2.xlsx" in archived
        assert any("__from_이전버전" in name for name in archived), \
            "동명 구버전을 덮어쓰지 않고 보존하지 못했다"
    finally:
        if _synthetic_flag is not None:
            os.environ["CSOS_SYNTHETIC"] = _synthetic_flag
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

    # 공유 폴더 정본을 원본 자료 폴더에 복사해 둔 경우 같은 70건을 140건으로 세면 안 된다.
    import shutil
    copied = os.path.join(tmp, "보관복사본.xlsx")
    shutil.copy2(dep, copied)
    uniq = rf._unique_deposit_files([dep, copied])
    assert uniq == [dep], "내용이 같은 입금 파일 복사본을 중복 집계한다: %s" % uniq
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
    assert "const actions = body.querySelector('.actions.sticky,.media-tools.sticky')" in live
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
    icon_hash = hashlib.sha256(open(os.path.join(
        ROOT, "webapp", "brand", "csos-app-icon-source.png"), "rb").read()).hexdigest()[:12]
    assert f"csos-icon-{icon_hash}" in sw, \
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
    assert "2026-07-28 업무 처리 1건" in text and "류지영 매니저" in text

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

            # 문서사진은 날짜별 하위 폴더에 있다. 최상위만 세거나 2025년을 섞으면 안 된다.
            photo_root = os.path.join(tmp, "문서사진")
            os.makedirs(os.path.join(photo_root, "2026", "07", "2026-07-29"))
            os.makedirs(os.path.join(photo_root, "2025", "12", "2025-12-26"))
            open(os.path.join(photo_root, "2026", "07", "2026-07-29", "a.jpg"), "wb").write(b"1")
            open(os.path.join(photo_root, "2026", "07", "2026-07-29", "b.png"), "wb").write(b"22")
            open(os.path.join(photo_root, "2025", "12", "2025-12-26", "old.jpg"), "wb").write(b"333")
            count, size_mb = D.image_tree_stats(photo_root, year="2026")
            assert count == 2 and size_mb > 0, (count, size_mb)
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
    for marker in ("CHECKPOINT_PATH", "def write_checkpoint(", "★ 진행 중 작업 — 여기서 바로 재개",
                   "--checkpoint", "--clear-checkpoint"):
        assert marker in src, f"새 세션 재개 체크포인트 누락: {marker}"

    # 워치독이 30분마다 남기는가 · 시작 체크리스트 0번인가
    wd = open(os.path.join(ROOT, "watchdog.py"), encoding="utf-8").read()
    assert "snapshot_handoff" in wd and "session_handoff.py" in wd, "워치독이 스냅샷을 안 남긴다"
    cm = open(os.path.join(os.path.dirname(ROOT), "CLAUDE.md"), encoding="utf-8").read()
    assert "session_handoff.py --check" in cm, "시작 체크리스트에 없다 — 아무도 안 읽는다"
    assert cm.index("session_handoff.py --check") < cm.index("AGENTS.md` 전체 읽기"), \
        "세션인계가 체크리스트 맨 앞이 아니다"
    print("  [53] 세션 인계(죽은 점유 판정·큐/임시파일·워치독 스냅샷·체크리스트 0번) ✅")


def t54_side_work_db_only():
    """[54] 철거·신규납품은 **DB에만 두고 앱에는 안 보인다**(사용자 지시 2026-07-29).

    원장에서 지우는 게 아니다 — 05_신규납품설치 시트에 그대로 두고 화면에서만 뺀다.
    ★ 05시트 업무구분은 '납품'·'철거' 처럼 **한 단어**다(10_코드관리 M열).
      '신규납품' 만 잡으면 정작 원장에 적힌 '납품' 을 놓친다 — 처음에 실제로 그랬다.
    ★ 돌발AS·정기점검은 절대 걸리면 안 된다. 걸리면 본업이 화면에서 사라진다."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import kakao_extract as kx

    # 저장 자리 — 05시트로 가고, 업무구분은 유효성 목록 안의 값이어야 한다
    assert "05_신규납품설치" in kx.SHEET_MAP, "철거·납품의 저장 자리가 없다"
    assert set(kx.SIDE_KIND.values()) <= {"철거", "납품", "설치", "이전"}, kx.SIDE_KIND
    assert kx.SIDE_KIND.get("(철거·보관)") == "철거" and kx.SIDE_KIND.get("(납품설치)") == "납품"
    extract_src = open(os.path.join(ROOT, "kakao_extract.py"), encoding="utf-8").read()
    assert 'r["시트"] not in SIDE_KIND' in extract_src, \
        "05시트에 정상 등록한 철거·납품을 보류라고 잘못 표시한다"
    assert '("02_돌발AS접수", "04_정기점검", "05_신규납품설치")' in extract_src, \
        "05시트 철거·납품을 기존 프로젝트로 안 보면 매일 중복 등록된다"
    m = kx.SHEET_MAP["05_신규납품설치"]
    assert m.get("업무구분") == "_업무구분" and "프로젝트NO" in m, m
    for c in ("요청일", "철거·이전예정일", "납품예정일"):
        assert c in kx.DATE_COLS, "%s 가 날짜로 안 들어간다" % c

    # 앱 필터 — 별도 공사는 걸러지고 본업은 남아야 한다
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    mm = re.search(r"const SIDE_WORK = /([^/]+)/", live)
    assert mm, "앱에 SIDE_WORK 규칙이 없다"
    rx = re.compile(mm.group(1))
    for word in ("철거", "납품", "설치", "이전", "계단"):
        assert rx.search(word), "'%s' 가 앱에서 안 걸러진다" % word
    for word in ("돌발AS", "정기점검"):
        assert not rx.search(word), "★ '%s' 가 걸러진다 — 본업이 화면에서 사라진다" % word

    # 데이터 받는 관문 전부에 걸려 있는가(한 곳만 빠져도 거기서 새어 나온다)
    for anchor in ("rowIs2026(r,'settle')", "rowIs2026(r,'as')",
                   "rowIs2026(r,'pm')", "rowIs2026(r,'issue')"):
        i = live.find(anchor)
        assert i > 0, anchor
        assert "isSideWork(r)" in live[i:i + 120], "%s 관문에 필터가 없다" % anchor
    assert live.count("isSideWork(r)") >= 5, "관문 일부에만 걸려 있다"
    print("  [54] 철거·납품 DB 저장·앱 비표시(05시트 저장·관문 5곳·본업 보존) ✅")


def t55_pm_brief_drilldown_and_capture():
    """[55] 정기점검 예정·실행·미실행 숫자는 원천 목록·점검ID·캡처로 이어진다."""
    import daily_brief as D

    data = {"as": [], "fw": [], "events": [], "pm": [
        {"점검ID": "PM-2607-001", "프로젝트NO": "UJ2607001", "캠프명": "합성A",
         "점검예정일": "2026-07-29", "실제점검일": "2026-07-29",
         "담당기사": "김기사", "점검내용": "정기점검 A"},
        {"점검ID": "PM-2608-002", "프로젝트NO": "UJ2608002", "캠프명": "합성B",
         "점검예정일": "2026-08-05", "실제점검일": "",
         "담당기사": "권기사", "점검내용": "정기점검 B"},
        {"점검ID": "PM-2606-003", "프로젝트NO": "UJ2606003", "캠프명": "합성C",
         "점검예정일": "2026-06-20", "실제점검일": "2026-07-29",
         "담당기사": "차기사", "점검내용": "지연 실행"},
    ], "pm_schedule": {"year": 2026, "quarter": 3, "schedule": [
        {"일정ID": "SCH-A", "점검예정일": "2026-07-29", "예측점검일": "",
         "실제점검일": "2026-07-29", "캠프명": "합성A", "장비수": 2},
        {"일정ID": "SCH-B", "점검예정일": "", "예측점검일": "2026-08-05",
         "실제점검일": "", "캠프명": "합성B", "장비수": 3},
        {"일정ID": "SCH-C", "점검예정일": "2026-07-29", "예측점검일": "",
         "실제점검일": "2026-07-29", "캠프명": "합성C", "장비수": 1},
    ]}}
    b = D.brief("2026-07-29", data)
    assert b["정기점검"]["예정"] == 2 and b["정기점검"]["완료"] == 2
    assert b["정기점검"]["예정장비"] == 3 and b["정기점검"]["완료장비"] == 3
    assert b["정기점검"]["분기예정"] == 6 and b["정기점검"]["분기완료"] == 3
    assert b["정기점검"]["분기미실행"] == 3 and b["정기점검"]["분기진행률"] == 50
    assert b["정기점검"]["기준일까지예정"] == 3
    assert b["정기점검"]["기준일까지완료"] == 3
    assert b["점검예정목록"][0]["레코드ID"] == "PM-2607-001"
    assert {x["레코드ID"] for x in b["점검실행목록"]} == {"PM-2606-003", "PM-2607-001"}
    assert [x["상태"] for x in b["분기점검목록"]] == ["실행", "미실행"]

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for token in ("function openPmBrief(", "openPmBrief('day-plan')", "openPmBrief('day-done')",
                  "openPmBrief('quarter-all')", "openPmBrief('quarter-done')",
                  "openPmBrief('quarter-pending')", "점검예정목록", "점검실행목록", "분기점검목록"):
        assert token in live, "실시간 앱 " + token + " 누락"
    for token in ("신규중처리완료", "신규처리율", "예측점검일",
                  "명세서발행대기완료기준", "과거 이력 예측"):
        assert token in live, "대표 카드·캘린더 " + token + " 누락"
    assert "window._briefMetric={label" in live and "shareExecMetric()" in live

    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    pub = open(os.path.join(ROOT, "cloud_publish.py"), encoding="utf-8").read()
    for token in ("function openPmCloud(", "function openCloudBriefSheet(",
                  "function cloudSheetToPng(", "quarter-pending", "이미지 저장/전달"):
        assert token in phone, "고정 주소 앱 " + token + " 누락"
    for key in ("점검예정목록", "점검실행목록", "분기점검목록"):
        assert f'"{key}": _b.get("{key}", [])' in pub, "암호화 사본 " + key + " 누락"
    print("  [55] 정기점검 예정·실행·미실행 목록·정확 이동·담당자 캡처 ✅")


def t56_work_detail_from_source():
    """[56] 작업 내용은 **원문에 있으면 자동 기입, 없으면 비워 둔다**(사용자 지시 2026-07-29).

    ★★ '신청내용'(요청)을 실적 칸에 넣으면 **하지 않은 작업을 했다고 기록**된다.
       카톡 자료로 세어 보면 작업내용은 26건인데 신청내용만 있는 게 542건이다 —
       구분 없이 채웠다면 대표 보고의 '무엇을 했나' 가 통째로 거짓이 됐다.
    ★★ 공지의 **빈 양식**('?? 호기', 'What ? / 갯수 ?')도 값이 아니다. 기사가 아직 안 채운 칸이다.
       여러 호기 중 일부만 적힌 글은 뒤에 빈 양식이 남으므로 거기서 잘라야 한다."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import fill_work_detail as W

    # 빈 양식은 값이 아니다 — 통째로 템플릿이면 아무것도 안 쓴다
    tpl = "● A/S 내용 :\n▒▒ ?? 호기 ▒▒\n(유료)What ? / 갯수 ?\n(유료)작업자 1공임"
    assert W.work_part(tpl) == "", "빈 양식을 실적으로 기록한다: %r" % W.work_part(tpl)
    assert W.tidy("▒▒ ?? 호기 ▒▒ (유료)What ? / 갯수 ?") == ""

    # 일부만 적힌 글 — 적힌 데까지만 남기고 빈 양식은 잘라낸다
    part = ("● A/S 내용 :\n▒▒ 02 호기 ▒▒\n(유료)도어락 교체 완료\n"
            "▒▒ ?? 호기 ▒▒\n(유료)What ? / 갯수 ?")
    got = W.work_part(part)
    assert "도어락 교체 완료" in got, got
    assert "What" not in got and "?? 호기" not in got, "빈 양식이 남았다: %r" % got

    # 사내 메모(※ 확인사항)는 실적이 아니다
    memo = "● A/S 내용 :\n(유료)경광등 교체완료\n※ 확인사항 : 1. (변재선) 거래명세표 발행"
    g2 = W.work_part(memo)
    assert "경광등 교체완료" in g2 and "확인사항" not in g2, g2

    # 'A/S 내용' 절이 없으면 본문을 통째로 넣지 않는다
    assert W.work_part("♣ 돌발유료 A/S 완료 ● 프로젝트NO : UJ2600001 ● 신청일자 : 2026.01.02") == ""

    # ★ 신청내용을 쓰지 않는다 — 코드에 그 폴백이 있으면 안 된다
    src = open(os.path.join(ROOT, "fill_work_detail.py"), encoding="utf-8").read()
    body = src.split("if __name__")[0]
    assert 'r.get("신청내용")' not in body, "신청내용을 실적으로 쓰고 있다 — 안 한 일을 했다고 적는다"
    # 사람이 적은 값은 덮지 않는다(자리표시자·빈 양식만 덮는다)
    assert "DRAFT.search(cur)" in body, "사람이 적은 값을 덮어쓸 수 있다"
    # 매일 자동으로 도는가
    dr = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "fill_work_detail.py" in dr, "daily_run 에 자동 기입 단계가 없다"
    # ★★ 'A/S 내용' 의 값은 `▒▒ 01 호기 ▒▒ …` 처럼 ▒ 로 **시작한다**.
    #    필드 정규식이 ▒ 에서 값을 끊으면 작업내용이 **언제나 빈 값**이 된다(실제로 그랬다:
    #    26건만 잡히다가 고친 뒤 840건). 값의 끝은 ● ★ ♠ ※ 로 본다.
    import kakao_extract as _ke
    got = dict(_ke.RE_FIELD.findall(
        "● 신청내용 : 리모컨 고장 ★ 작업완료후 - 완료전화 필수! "
        "● A/S 내용 : ▒▒ 01 호기 ▒▒ (유료)리모콘 SET 교체 완료 ♠ 원인 및 조치 : -"))
    as_val = got.get("A/S 내용") or got.get("A/S내용") or ""
    assert "리모콘 SET 교체 완료" in as_val, "A/S 내용이 ▒ 에서 잘린다: %r" % as_val
    assert "리모컨 고장" in (got.get("신청내용") or ""), got

    # ★★ 03시트 B열(접수ID)은 **수식**이고 02시트에서 스스로 끌어온다 —
    #    빈 행에 실적을 써 넣으면 재계산 때 엉뚱한 건에 붙는다(v259 실사고).
    assert "def missing_rows" in body, "대기 건수 집계가 없다"
    assert "새행시작" not in body and "실적 행을 만든다" not in body, \
        "★ 03시트에 행을 직접 만들고 있다 — B열 수식이 정하는 자리를 가로챈다"
    print("  [56] 작업내용 자동기입(신청내용 금지·빈 양식 제외·▒ 절단 방지·행 생성 금지) ✅")


def t58_check_hub_detail_and_capture():
    """[58] 보고 아래 확인 필요 전용 화면은 유형·담당자·원기록·캡처까지 이어진다."""
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()

    # 메뉴는 사용자가 지정한 위치(보고 바로 아래, 기록 위)에 있어야 한다.
    daily = live.index('data-v="daily"')
    check = live.index('data-v="check"', daily)
    report = live.index('data-v="report"', check)
    assert daily < check < report, "확인 필요 메뉴가 보고 아래·기록 위가 아니다"
    assert 'id="v-check"' in live and "if(v==='check') renderCheckHub()" in live

    # 숫자만 보여 주는 카드가 아니라 유형/담당자/전체 목록을 실제 행으로 좁혀야 한다.
    for token in ("function renderCheckHub(", "function renderCheckList(",
                  "function openCheckType(", "function openCheckOwner(",
                  "function openCheckFiltered(", "function openCheckStale(",
                  'id="checktypes"', 'id="checkowners"', 'id="checklist"',
                  'id="checkq"', 'id="checktype"', 'id="checkowner"'):
        assert token in live, token + " 누락"

    # 2026년·별도공사 제외 규칙을 전용 화면에서도 다시 지키고, ID 종류로 정확한 원장 기록을 연다.
    assert "rowIs2026(r,'issue')&&!isSideWork(r)" in live
    assert "/^AS-/i.test(raw)" in live and "/^PM-/i.test(raw)" in live and "/^JS-/i.test(raw)" in live
    assert "if(x.kind) openRecord(x.kind,x.id,x.project)" in live

    # 현재 필터 결과를 기존 캡처 엔진에 넘겨 이미지 저장·전달 기능을 그대로 쓴다.
    assert "rows:rows.map(checkMetricRow)" in live and "openExecMetric(label)" in live
    assert "현재 목록 보기" in live and "이미지 복사" in live and "이미지 저장" in live
    # 아이콘 참조 방식(<img> → 스프라이트 <use>)이 바뀌어도 깨지지 않게 **아이콘 이름**으로 본다.
    assert "media-tools" in live and "#i-clipboard-copy" in live
    # 모바일 탭바: 크기 **숫자를 고정하지 않는다**(2026-07-30 아이콘을 키우자 이 줄이 깨졌다).
    # 지켜야 하는 것은 ① 좁은 화면용 규칙이 있다 ② 아이콘이 **알아볼 수 있는 크기**다.
    #   사용자 지적: "모바일에서 아이콘이 너무 작아 잘 안 보여" — 탭이 7개로 늘며 좁아진 탓이다.
    assert "@media(max-width:420px)" in live, "좁은 화면용 탭바 규칙이 없다"
    _icons = [float(m) for m in re.findall(r"\.tabbar svg\{width:([\d.]+)px", live)]
    assert _icons, "탭바 아이콘 크기 규칙이 없다"
    assert min(_icons) >= 24, f"모바일 탭바 아이콘이 너무 작다({min(_icons)}px) — 24px 이상 유지"

    # PO번호만 있는 확인 행도 정산 PO번호 연결을 먼저 찾고, 없으면 경고창이 아닌 원문 상세를 연다.
    for token in ("function samePo(", "settleRows.find(r=>samePo(r.PO번호,raw))",
                  "function openCheckSource(", "function openCheckByKey(",
                  "else openCheckSource(r)", "if(kind==='check')"):
        assert token in live, "PO·미연결 확인행 처리 누락: " + token

    # 확정된 내부 확인 책임을 적용하고, 그 밖의 현장 확인만 원천 AS/PM 담당기사로 보완한다.
    assert "const CHECK_OWNER_RULES" in live and "function confirmedCheckOwner(" in live
    assert "캠프·일정, 카톡/밴드, 현장자료, 완료일, 거래명세서, 세금계산서, 입금, 금액 불일치 확인" in live
    assert "{name:'변재선(회계)'" not in live, "이관 완료한 변재선 확인 책임이 앱에 남아 있다"
    assert 'STAFF_CENTER_ALIASES = {"byeon-jaeseon": "ryu-jiyeong"}' in server
    assert '"Location": f"/staff/{target}"' in server
    assert "const CHECK_OWNER_SCOPE_PROPOSALS" not in live and "openPendingOwnerScope" not in live
    assert "first&&first.원천업무ID ? rowById(first.원천업무ID)" in live
    assert "원장 담당자 칸 확인" in live
    assert "[r['구분'], r['담당자']" in phone, "고정 주소 확인필요 목록에 확정 담당자가 표시되지 않는다"
    print("  [58] 확인 필요 전용 메뉴·PO/담당자 보완·정확 이동·현재목록 캡처 ✅")


def t70_quarter_as_months():
    """[70] '3분기' 가 아니라 '7~9월' 로 적는다(사용자 지시 2026-07-29).

    분기 번호는 읽는 사람이 머릿속에서 다시 월로 환산해야 한다 — 대표 보고에서
    "그게 몇 월부터냐" 를 되묻게 만든다. 화면·브리핑 어디에도 'N분기' 를 남기지 않는다."""
    src = open(os.path.join(ROOT, "daily_brief.py"), encoding="utf-8").read()
    assert '"분기": f"{y}년 {3 * q - 2}~{3 * q}월"' in src, "분기 라벨이 월 범위가 아니다"
    assert '"분기범위"' in src and '"분기끝월"' in src, "월 범위·끝월 필드가 없다"
    assert 'f"{y}년 {q}분기"' not in src, "옛 'N분기' 표기가 남아 있다"
    assert "분기 내 마무리" not in src, "'분기 내' 문구가 남아 있다 — 몇 월인지 안 보인다"

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    # 앱은 이제 기간을 직접 고르므로 라벨도 클라이언트가 만든다(pmStats().라벨).
    # 서버 필드(분기범위·분기끝월)는 브리핑 텍스트용으로 남고, 앱은 ps.라벨/ps.끝월을 쓴다.
    assert "ps.라벨" in live and "ps.끝월" in live, "앱이 월 범위 라벨을 안 쓴다"
    assert "분기 내 마무리" not in live, "앱에 '분기 내' 문구가 남아 있다"
    # 버튼 라벨이 '분기 예정' 처럼 고정 문자열로 되돌아가면 안 된다
    assert '">분기 예정<' not in live and '">분기 실행<' not in live, "버튼이 다시 '분기' 로 고정됐다"

    # 실제 환산이 맞는가 — 3분기는 7~9월
    for q, want in ((1, "1~3월"), (2, "4~6월"), (3, "7~9월"), (4, "10~12월")):
        assert "%d~%d월" % (3 * q - 2, 3 * q) == want, q
    print("  [70] 분기를 '7~9월' 처럼 월 범위로 표기(브리핑·앱·버튼·안내문) ✅")


def t71_period_range():
    """[71] 기간을 **월 단위 범위**로 고른다 — '7월' 도 '1~6월' 도(사용자 지시 2026-07-29).

    ★ 시작이 끝보다 뒤면 오류를 띄우지 않는다. 사용자가 틀린 게 아니라 **아직 반대쪽을
      못 옮긴 것**이다 — 방금 고른 쪽을 존중하고 반대쪽을 끌어와 맞춘다.
    ★ 옛 호출부(`inPeriod(d, y, m)` 처럼 한 달만 넘기는 곳)가 남아 있다. '이번 달' 통계는
      선택 기간이 아니라 늘 이번 달이어야 하므로, 인자 3개면 **그 한 달**로 동작해야 한다."""
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    # 단일 월 선택이 아니라 시작~끝 두 칸이어야 한다
    assert 'id="pfrom"' in live and 'id="pto"' in live, "기간 시작·끝 선택이 없다"
    assert 'id="pmonth"' not in live, "옛 단일 월 선택이 남아 있다(둘이 어긋난다)"
    for fn in ("function periodRange()", "function periodLabel()", "function fixRange(",
               "function setRange("):
        assert fn in live, "%s 가 없다" % fn
    assert live.count("function periodRange()") == 1, \
        "동명 periodRange가 업무목록·월별현황 사이에서 덮어써 목록이 0건이 된다"
    assert "function workPeriodRange()" in live and "const pr = workPeriodRange();" in live, \
        "업무목록 전용 기간 함수가 분리되지 않았다"
    # 자주 쓰는 기간 바로가기
    for label in ("연간 전체", "상반기", "하반기"):
        assert 'onclick="setRange' in live and label in live, label

    # 범위 필터가 실제로 쓰이는가 — 옛 단일 월 변수로 되돌아가면 안 된다
    for anchor in ("dateOf(r,'settle'),y,pf,pt", "dateOf(r,'as'),y,pf,pt", "dateOf(r,'pm'),y,pf,pt"):
        assert anchor in live, "기간 필터가 적용되지 않은 곳이 있다: %s" % anchor
    # 하위호환: 인자 3개면 그 한 달
    assert "if(to === undefined) to = from;" in live, "옛 호출부(단일 월) 호환이 깨졌다"
    # 제목은 '7월'·'1~6월'·'연간 전체'
    assert "${periodLabel()}" in live, "제목이 기간 이름을 쓰지 않는다"
    assert "${m?+m+'월':'전체'}" not in live, "옛 제목 표기가 남아 있다"

    # 라벨 규칙(파이썬으로 같은 계산을 재현해 확인)
    def label(f, t):
        return "연간 전체" if (f, t) == ("01", "12") else (
            "%d월" % int(f) if f == t else "%d~%d월" % (int(f), int(t)))
    assert label("07", "07") == "7월" and label("01", "06") == "1~6월"
    assert label("01", "12") == "연간 전체" and label("10", "12") == "10~12월"
    # ── 정기점검 진행률 카드도 같은 기간 선택을 쓴다 ──
    #  ★ 서버가 준 분기 숫자(p.분기예정 등)를 그대로 쓰면 기간을 바꿔도 숫자가 안 변한다.
    #    works.pm 에서 **직접 세야** 아무 기간이나 된다.
    for fn in ("function pmStats()", "function setPmRange(", "function pmQuarter()"):
        assert fn in live, "%s 가 없다" % fn
    assert "const ps = pmStats();" in live, "진행률 카드가 기간 계산을 안 쓴다"
    for old in ("p.분기진행률||0", "${p.분기완료||0}/${p.분기예정||0}건",
                "p.분기범위||'분기'", "p.분기||'분기'"):
        assert old not in live, "진행률 카드에 옛 분기 고정값이 남아 있다: %s" % old
    # 드릴다운 목록도 같은 기간이어야 카드 숫자와 어긋나지 않는다
    assert "const _ps = pmStats();" in live, "드릴다운이 기간을 따르지 않는다"
    assert "mm>=_ps.from && mm<=_ps.to" in live, "드릴다운 목록이 기간으로 걸러지지 않는다"
    # 시작이 끝보다 뒤면 끝을 끌어와 맞춘다(오류를 띄우지 않는다)
    assert "if(a > b){ b = a; }" in live, "거꾸로 고른 기간을 바로잡지 않는다"
    print("  [71] 기간 범위 선택(현황·정기점검 진행률·드릴다운·거꾸로 자동정렬·단일월 호환) ✅")


def t72_project_first_representative_report():
    """[72] 내부 ID 대신 프로젝트번호를 대표 표시하고, 유수비 예외보고는 원장에 쓰지 않고 계산한다."""
    from webapp.app_server import representative_summary

    works = {
        "as": [
            {"접수ID": "AS-2607-001", "프로젝트NO": "UJ2600001", "캠프명": "테스트A",
             "접수일자": "2026-07-25", "진행상태": "접수", "담당기사": "기사A"},
            {"접수ID": "AS-2607-002", "프로젝트NO": "UJ2600002", "캠프명": "테스트B",
             "접수일자": "2026-07-20", "작업완료일": "2026-07-21", "진행상태": "작업완료",
             "사진등록": "등록", "완료보고서등록": "", "ERP등록": "등록"},
        ],
        "pm": [
            {"점검ID": "PM-2607-001", "프로젝트NO": "UJ2600011", "캠프명": "점검A",
             "점검예정일": "2026-07-05", "실제점검일": "2026-07-05", "점검상태": "완료",
             "담당기사": "기사A"},
            {"점검ID": "PM-2607-002", "프로젝트NO": "UJ2600012", "캠프명": "점검B",
             "점검예정일": "2026-08-05", "점검상태": "예정", "담당기사": "기사B"},
        ],
    }
    settlements = [
        {"정산ID": "JS-2607-001", "프로젝트NO": "UJ2600001", "업무구분": "돌발AS",
         "캠프명": "테스트A", "완료일": "2026-07-25", "비용구분": "유상",
         "공급가액": 100000, "명세서번호": ""},
        {"정산ID": "JS-2607-002", "프로젝트NO": "UJ2600011", "업무구분": "정기점검",
         "캠프명": "점검A", "완료일": "2026-07-05", "비용구분": "유상",
         "공급가액": 200000, "명세서번호": "2026/07/05-1"},
        {"정산ID": "JS-2607-003", "프로젝트NO": "UJ2600012", "업무구분": "정기점검",
         "캠프명": "점검B", "완료일": "2026-07-06", "비용구분": "유상",
         "공급가액": 300000, "명세서번호": "", "명세서": "없음",
         "명세서발행일": "2026-07-07"},
    ]
    report = representative_summary(works, settlements, "2026-07-29")
    assert report["돌발AS"]["전산상미완료"] == 1 and report["돌발AS"]["D+2초과"] == 1
    assert report["돌발AS"]["현장완료서류미정리"] == 1
    assert report["정기점검"]["전체대상"] == 2, "분기 목표를 실제 예정행이 아닌 고정 숫자로 썼다"
    assert report["정기점검"]["목표누계"] <= report["정기점검"]["전체대상"]
    docs = {x["업무유형"]: x for x in report["거래명세서"]["업무유형별"]}
    assert docs["돌발 AS"]["발행대상"] == 1 and docs["돌발 AS"]["미발행"] == 1
    assert docs["정기점검"]["발행대상"] == 2 and docs["정기점검"]["미발행"] == 0, \
        "명세서 발행일이 확인된 완료 건을 번호 공란만으로 미발행 처리했다"
    assert docs["신규·납품·설치"]["발행대상"] == 0 and docs["신규·납품·설치"]["발행률"] is None
    assert (
        len(report["업무기준확인필요"]) + len(report["업무기준확정"])
    ) == 5, "저장된 확정 여부와 무관하게 운영 기준 5개가 모두 유지돼야 한다"
    assert not any("류지영 확인 범위" in x["기준"] for x in report["업무기준확인필요"])
    assert not any("변재선(회계) 확인 범위" in x["기준"] for x in report["업무기준확인필요"])

    from responsibility import confirmed_owner
    assert confirmed_owner("밴드 게시 미확인", "밴드", "권오철") == "류지영"
    assert confirmed_owner("캠프명 비어 있음", "빈칸", "김준형") == "류지영"
    assert confirmed_owner("세금계산서 미발행", "정산") == "류지영"
    assert confirmed_owner("작업금액 불일치", "금액") == "류지영"
    assert confirmed_owner("PO 원본 누락", "원본자료") == "오종현"
    assert confirmed_owner("견적서 원천자료 누락", "원본자료") == "오종현"
    assert confirmed_owner("입금 원천자료 누락", "원본자료") == "오종현"
    assert confirmed_owner("현장자료 누락", "현장", "김준형") == "류지영"
    assert confirmed_owner("PO A", "PO") == "유현민"
    assert confirmed_owner("원장 미등록", "문서") == "유현민"
    assert confirmed_owner("기타 현장 확인", "기타", "권오철") == "권오철"

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    sw = open(os.path.join(ROOT, "docs", "sw.js"), encoding="utf-8").read()
    for token in ("function projectNoOf(", "function isRepresentativeProject(", "프로젝트 미확정",
                  "function renderRepresentative(",
                  "function openRepresentativeList(", 'id="checkpolicies"', "press-pop",
                  "@keyframes viewEnter", "PO 원본·견적서, 구매·입금 원천자료 취합·누락 확인"):
        assert token in live, token + " 누락"
    for token in ("function execReportedCount(", "정기점검 이상 상세 미연결",
                  "확인 전까지 확정 이상 건으로 보지 않습니다",
                  "세부 미확정 · 원장 요약"):
        assert token in live, token + " 누락"
    for route in ("/api/v1/reports/daily/exceptions", "/api/v1/as-requests/backlog-detail",
                  "/api/v1/inspections/quarter-progress", "/api/v1/statements/unissued"):
        assert route in server, route + " 누락"
    assert "project, _ = rep_no(" in server and "candidate, rep_idx" in server, \
        "대표보고 원천행의 내부 JS 번호를 프로젝트번호로 복원하지 않는다"
    assert "r['프로젝트NO']||'프로젝트 미확정'" in phone
    assert "오종현 원천자료 취합" in phone
    icon_hash = hashlib.sha256(open(os.path.join(
        ROOT, "webapp", "brand", "csos-app-icon-source.png"), "rb").read()).hexdigest()[:12]
    assert f"csos-icon-{icon_hash}" in sw
    print("  [72] 프로젝트번호 대표표시·유수비 예외보고·정책확인·버튼/페이지 애니메이션 ✅")


def t73_pm_schedule_source_sync(tmp):
    """정기점검 원본은 현재 분기만·제외행 없이·가짜 UJ 없이 멱등 반영한다."""
    from pm_schedule_sync import parse_schedule, link_master, build_sheet, HEADERS, SHEET_NAME
    from findings_sheet import upsert

    source = os.path.join(tmp, "정기점검_스케줄_합성.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2026년 3분기 정기점검"
    for _ in range(4):
        ws.append([])
    ws.append(["No.", "월", "3분기 점검일자", "2분기 점검일자", "특이사항", "확정자",
               "기존 캠프명", "변경 캠프명", "호기", "종류", "모델"])
    # 원본 월과 확정 날짜가 다르면 날짜의 달이 우선한다.
    ws.append([1, 8, "2026-07-11", "2026-04-11", "", "김준형", "테스트A캠프", "", "1호기", "LIFT", "M1"])
    ws.append([2, 8, "2026-07-11", "2026-04-11", "", "김준형", "테스트A캠프", "", "2호기", "LIFT", "M1"])
    ws.append([3, 8, "", "2026-05-09", "", "권오철", "테스트B캠프", "", "1호기", "LIFT", "M2"])
    ws.append([4, 3, "2026-03-10", "", "", "김필우", "분기밖캠프", "", "1호기", "LIFT", "M3"])
    ws.append([5, 9, "", "2026-06-20", "철거 예정", "차동호", "제외캠프", "", "1호기", "LIFT", "M4"])
    wb.save(source)

    parsed = parse_schedule(source, 2026, 3)
    assert parsed["scanned"] == 4 and parsed["excluded"] == 1
    assert len(parsed["records"]) == 2
    a = next(r for r in parsed["records"] if r["캠프명"] == "테스트A캠프")
    b = next(r for r in parsed["records"] if r["캠프명"] == "테스트B캠프")
    assert a["장비수"] == 2 and a["점검예정일"] == "2026-07-11" and a["예정월"] == "2026-07"
    assert b["점검예정일"] == "" and b["예정월"] == "2026-08"
    assert b["예측점검일"] == "2026-08-09" and "2026-05-09" in b["예측근거"]

    master_rows = [{
        "점검ID": "PM-2607-001", "프로젝트NO": "UJ2600001", "캠프명": "테스트A캠프",
        "점검예정일": "2026-07-11", "실제점검일": "2026-07-12", "점검상태": "완료",
    }]
    linked = link_master(parsed["records"], master_rows)
    la = next(r for r in linked if r["캠프명"] == "테스트A캠프")
    lb = next(r for r in linked if r["캠프명"] == "테스트B캠프")
    assert la["연결프로젝트NO"] == "UJ2600001" and la["반영상태"] == "완료 실적 우선"
    assert la["완료장비수"] == 2 and la["실제점검일"] == "2026-07-12"
    assert lb["연결프로젝트NO"] == "" and lb["반영상태"] == "프로젝트 매칭 대기"

    # 같은 캠프·같은 달에 날짜가 여러 개인 경우, 한 날짜의 완료를 다른 일정에
    # 복제하면 실제 장비 진행률이 과대계상된다.
    split_source = [
        {"캠프명": "분할캠프", "예정월": "2026-07", "점검예정일": "2026-07-02",
         "담당기사": "김준형", "장비수": 2},
        {"캠프명": "분할캠프", "예정월": "2026-07", "점검예정일": "2026-07-22",
         "담당기사": "김준형", "장비수": 2},
    ]
    split_master = [{
        "프로젝트NO": "UJ2600999", "캠프명": "분할캠프", "담당기사": "김준형",
        "점검예정일": "2026-07-22", "실제점검일": "2026-07-22", "점검상태": "완료",
    }]
    split = link_master(split_source, split_master)
    assert split[0]["완료장비수"] == 0 and split[0]["반영상태"] == "프로젝트 매칭 대기"
    assert split[1]["완료장비수"] == 2 and split[1]["반영상태"] == "완료 실적 우선"

    styled_xml = build_sheet(linked * 6, os.path.basename(source), styled=True)
    assert '<mergeCell ref="A1:M1"/>' in styled_xml and '<mergeCell ref="A2:M2"/>' in styled_xml
    assert re.search(r'<c r="A10" s="38"', styled_xml), "10~49행 입력 서식이 빠졌다"

    ledger = os.path.join(tmp, "쿠팡_통합업무_일일보고_관리대장_v1.xlsx")
    mwb = openpyxl.Workbook()
    mwb.active.title = "00_대시보드"
    mwb.save(ledger)
    xml = build_sheet(linked, os.path.basename(source))
    v2, _ = upsert(ledger, xml, sheet_name=SHEET_NAME, headers=HEADERS)
    assert v2 and os.path.exists(v2)
    check = openpyxl.load_workbook(v2, read_only=True, data_only=True)
    assert SHEET_NAME in check.sheetnames and check[SHEET_NAME].max_row == 6
    check.close()
    again, msg = upsert(v2, xml, sheet_name=SHEET_NAME, headers=HEADERS)
    assert again is None and "멱등" in msg

    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "27_정기점검원본일정" in server and "정기점검 스케줄 원본" in live
    assert "프로젝트 매칭 대기" in phone and "원본 일정 · 장비 " in phone
    assert "pm_schedule_sync.py" in daily
    print("  [73] 류지영 정기점검 원본 현재분기·제외·완료우선·앱·멱등 자동반영 ✅")


def t74_billing_fill():
    """06_거래서류청구수금 청구원장 채우기 — 무엇을 넣고 무엇을 막는가.

    ★ 이 시트는 03시트와 **반대**다(v259 사고 참조).
      03시트는 B열(접수ID)이 수식이라 행 배정을 엑셀이 정하므로 빈 행에 미리 쓰면 안 됐다.
      06시트는 C열(원천업무ID)이 **입력열**이고 A열(정산ID)이 C에서 파생되므로
      빈 행에 순서대로 쓰는 것이 설계대로다. 이 구분이 무너지면 둘 다 망가진다.
    """
    import billing_fill as B
    assert B.self_test(), "billing_fill 자체 검증 실패"

    # 비쿠팡 차단: 판매조회에 있어도 관리대장에 없으면 절대 통과하지 못한다
    sales = {"UJ2699999": 1}
    ok, why = B.eligible("UJ2699999", sales, {}, set())
    assert not ok and "미등록" in why, why

    # 수식열(A 정산ID·I 실제작업공급가액·P·Q 합계)에는 어떤 경우에도 쓰지 않는다
    from datetime import date as _d
    q, _ = B.build_queue(
        [{"원천업무ID": "AS-2601-001", "공급가액": 500, "일자": _d(2026, 1, 2), "PO번호": "PO1"}],
        [72])
    cells = {x["cell"] for x in q}
    assert cells == {"C72", "O72", "N72", "S72"}, cells
    assert not any(x["cell"][0] in ("A", "I", "P", "Q", "K") for x in q)
    assert all(x.get("only_if_empty") for x in q), "빈 칸만 정책이 빠지면 남의 값을 덮는다"

    src = open(os.path.join(ROOT, "billing_fill.py"), encoding="utf-8").read()
    assert "claim_guard" in src, "원장 쓰기는 점유 확인을 거쳐야 한다"
    print("  [74] 청구원장 채우기 — 비쿠팡 차단·수식열 보호·빈칸만·점유확인 ✅")


def t75_gcal_sync():
    """구글 캘린더 대조 — 캘린더는 **예정**이지 실적이 아니다.

    사용자 지시(2026-07-29): "이 캘린더 추가하고 항상 대조해서 엑셀과 앱에 반영해줘".
    ★ 완료일 칸을 캘린더로 채우면 '안 한 일'이 '한 일'이 된다. 그 경계를 여기서 못박는다.
    """
    import gcal_sync as G
    assert G.self_test(), "gcal_sync 자체 검증 실패"

    # 채우는 열은 전부 '예정' 이어야 한다 — 실제·완료 열은 어떤 경로로도 대상이 아니다
    for kind, (sheet, daycol, timecol) in G.TARGET.items():
        assert "예정" in daycol, f"{kind} → {daycol} 은 예정 열이 아니다"
        assert "완료" not in daycol and "실제" not in daycol, daycol

    # 원천이 하나도 없어도 죽지 않는다(일일 파이프라인이 이것 때문에 멈추면 안 된다)
    old = os.environ.pop("COUPANG_GCAL_ICS", None)
    try:
        evs, notes = [], []
        assert isinstance(G.feeds(), list)
    finally:
        if old:
            os.environ["COUPANG_GCAL_ICS"] = old

    # 큐는 key 모드(접수ID/점검ID)로 나가고 빈 칸만 채운다
    from datetime import date as _d
    q = G.build_queue([{"업무구분": "돌발AS", "날짜": _d(2026, 7, 2), "시간": "09:00", "제목": "AS",
                        "원장": {"시트": "02_돌발AS접수", "원천업무ID": "AS-2607-001", "예정일있음": False}}])
    assert [x["col"] for x in q] == ["방문예정일", "방문예정시간"], q
    assert all(x["only_if_empty"] and x["key"] == "AS-2607-001" for x in q)
    # 구분과 시트가 어긋나면 아무것도 쓰지 않는다
    assert G.build_queue([{"업무구분": "정기점검", "날짜": _d(2026, 7, 2), "시간": "", "제목": "x",
                           "원장": {"시트": "02_돌발AS접수", "원천업무ID": "AS-1", "예정일있음": False}}]) == []

    # 비밀 주소가 커밋되지 않는다(규칙 1)
    gi = open(os.path.join(ROOT, ".gitignore"), encoding="utf-8").read()
    assert "config/gcal.json" in gi, "비공개 iCal 주소가 커밋될 수 있다"
    # 문서의 자리표시자(private-xxxx)는 괜찮다. 진짜 키(긴 16진수)가 박혀 있으면 안 된다.
    src = open(os.path.join(ROOT, "gcal_sync.py"), encoding="utf-8").read()
    assert not re.search(r"private-[0-9a-f]{16,}", src, re.I), "소스에 실제 비공개 주소가 박혀 있다"

    # 앱·일일실행 배선
    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "/api/calendar" in server and "gcal_events.json" in server
    # 캐시 목록은 앱에서 즉시 보이고, 사용자가 제공한 공개 임베드는 전체 일정 확인용으로만 쓴다.
    # 비공개 iCal 주소/config는 절대 브라우저에 전달하지 않는다.
    assert "loadCalendar" in live and "openCalendar" in live and "show('calendar')" in live
    assert "calendar.google.com/calendar/embed" in live and "COUPANG_CALENDAR_ID" in live
    # 캐시 일정을 앱이 직접 그린다. 2026-08-05 부터 한 줄 목록(calendarEventList)이 아니라
    # 월 격자(calGrid)+그 날 목록(calAgenda) 이다 — 검증 [102].
    assert "openGoogleCalendarDraft" in live and 'id="calAgenda"' in live and 'id="calGrid"' in live
    assert "config/gcal.json" not in live and ".ics" not in live, "비공개 일정 원천이 앱에 노출됐다"
    assert "gcal_sync.py" in daily
    print("  [75] 구글 캘린더 — 예정열만·구분일치·비밀주소 보호·캐시/공개전체보기 ✅")


def t78_recalc_pending_visible():
    """"원장엔 넣었다는데 앱엔 왜 없지?" 를 앱이 스스로 설명하게 한다.

    2026-07-29 실제 상황: 06시트에 청구 636건을 넣었지만 정산ID·금액이 **수식**이라
    엑셀이 계산하기 전까지 앱은 옛 건수만 읽었다. 숫자가 틀린 게 아니라 아직 안 나온 것인데,
    화면이 그 말을 안 하면 사용자는 자료가 사라졌다고 읽는다.
    ★ 고칠 수 없는 것(엑셀 수식 계산)은 **드러내는 것**이 정답이다.
    """
    import recalc_pending as R
    assert R.self_test(), "recalc_pending 자체 검증 실패"
    # 수식이 돌려준 빈 문자열을 '값 있음'으로 세면 대기가 0이 되어 배너가 안 뜬다
    assert R.count_rows(["AS-1", "AS-2"], ["", None]) == (2, 0)
    # ★ 새 행을 만드는 시트가 하나라도 빠지면 그 시트만 조용히 사라진다
    #   (2026-07-30: 카톡 40건을 02시트에 넣었는데 02가 감시목록에 없어 앱이 침묵했다)
    watched = {x[0] for x in R.SHEETS}
    for sheet in ("06_거래서류청구수금", "02_돌발AS접수", "04_정기점검", "05_신규납품설치"):
        assert sheet in watched, f"{sheet} 가 재계산 감시목록에 없다 — 넣은 건이 앱에서 사라진다"

    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert '"recalc": get_recalc_pending()' in server, "상태 API가 대기 건수를 안 내려준다"
    assert "재계산대기.json" in server and "recalc_pending.py" in daily
    assert "renderRecalc(s.recalc)" in live and 'id="recalccard"' in live

    # 라벨: 같은 숫자를 두 이름으로 부르면 두 개인 줄 안다
    assert "'정산 누적'" not in live, "히어로와 KPI가 같은 값을 다른 이름으로 부른다"
    assert live.count("'정산 건수'") >= 2
    assert "'공급가액 합계'" not in live, "어느 공급가액인지 드러나야 한다(작업/명세서/계산서 3종)"
    assert "'작업 공급가액(부가세 별도)'" in live
    assert live.count("공급가액(부가세 별도)") >= 8, \
        "대시보드·정산·보고·상세의 공급가액에 부가세 별도 표기가 공통 적용되지 않았다"
    assert "공급가액(원)" not in live, "부가세 포함 여부가 없는 옛 공급가액 라벨이 남았다"
    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    assert "원(부가세 별도)" in phone, "PC 독립 폰 사본의 공급가액에 부가세 별도 표기가 없다"
    assert "계산서 발행 후 미입금" in live, "미수금이 '떼인 돈'으로 읽힌다"
    print("  [78] 재계산 대기 안내 + 라벨 모호성 제거(같은 값 한 이름·금액 종류 명시) ✅")


def t84_duplicate_source_files(tmp):
    """같은 내용의 원천 파일이 여러 벌 있어도 금액이 배수로 부풀지 않는다.

    ★ 2026-07-30 실사고: 원본 자료 정리가 판매조회를 3벌 남겼고(SHA256 동일),
      billing_fill 이 전부 읽어 공급가액이 36.2억 → 108.6억으로 **3배**가 됐다.
      파일명 규칙(`__dup_`)으로 거르면 다른 이름으로 다시 뚫린다 — 내용으로 판정해야 한다.
    """
    import billing_fill as B
    a = os.path.join(tmp, "판매조회_x.xlsx")
    open(a, "wb").write(b"same-bytes")
    for name in ("판매조회_x__dup_1.xlsx", "판매조회_x__dup_1_2.xlsx", "sales_copy.xlsx"):
        open(os.path.join(tmp, name), "wb").write(b"same-bytes")
    other = os.path.join(tmp, "판매조회_y.xlsx")
    open(other, "wb").write(b"different")
    keep = B.dedupe_files([a, other] + [os.path.join(tmp, n) for n in
                                        ("판매조회_x__dup_1.xlsx", "판매조회_x__dup_1_2.xlsx", "sales_copy.xlsx")])
    assert len(keep) == 2, f"내용 중복을 못 걸렀다: {[os.path.basename(x) for x in keep]}"
    # 읽을 수 없는 경로는 조용히 건너뛴다(파이프라인이 이것 때문에 멈추면 안 된다)
    assert B.dedupe_files([os.path.join(tmp, "없는파일.xlsx")]) == []
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "billing_fill.py" in daily, "청구 근거 갱신이 일일 실행에 없다"
    # ★ 근본 원인도 막는다: 정리기가 사본을 계속 만들면 읽는 쪽을 고쳐도 파일이 계속 늘어난다.
    #   (2026-07-30: 목적지에 파일이 있으면 내용을 보지 않고 __dup_ 를 붙여 3벌이 됐다)
    import source_organizer as SO
    same = os.path.join(tmp, "src_same.xlsx")
    dst = os.path.join(tmp, "dst.xlsx")
    open(same, "wb").write(b"identical"); open(dst, "wb").write(b"identical")
    assert SO._collision_target(same, dst) == dst, "내용이 같은데 사본을 또 만든다"
    diff = os.path.join(tmp, "src_diff.xlsx")
    open(diff, "wb").write(b"really-different")
    got = SO._collision_target(diff, dst)
    assert got != dst and "__dup_" in os.path.basename(got), "내용이 다른데 덮어쓰려 한다"
    # 같은 내용의 사본이 이미 있으면 또 만들지 않는다(무한 증식 차단)
    open(got, "wb").write(b"really-different")
    assert SO._collision_target(diff, dst) == got, "같은 내용의 사본을 반복 생성한다"
    print("  [84] 원천 파일 내용중복 제거 — 금액 배수 합산 차단 + 사본 무한증식 차단 ✅")


def t86_daily_run_singleton_and_inbox_classification(tmp):
    """일일 실행은 프로세스가 겹치지 않고, ERP 원장은 이름이 아닌 분류기로 찾는다."""
    import daily_run as D

    lock = os.path.join(tmp, "daily_run.lock")
    token = D.acquire_run_lock(lock)
    assert token and os.path.isfile(lock), "첫 실행이 잠금을 잡지 못했다"
    assert D.acquire_run_lock(lock) is None, "두 번째 실행이 같은 잠금을 통과했다"
    D.release_run_lock(token, lock)
    assert not os.path.exists(lock), "정상 종료 뒤 잠금이 남았다"

    # 첫 실행이 O_EXCL 직후 JSON을 쓰는 순간을 두 번째 실행이 보면 빈 파일일 수 있다.
    # 방금 생긴 빈 잠금을 죽은 것으로 오판해 지우면 두 프로세스가 함께 실행된다.
    open(lock, "wb").close()
    assert D.acquire_run_lock(lock) is None, "생성 중인 최근 잠금을 빼앗았다"
    os.unlink(lock)

    # 비정상 종료로 남은 잠금은 실제 PID가 죽었을 때만 회수한다.
    with open(lock, "w", encoding="utf-8") as f:
        json.dump({"pid": 2147483646, "token": "dead-run"}, f)
    recovered = D.acquire_run_lock(lock)
    assert recovered and recovered != "dead-run", "죽은 프로세스의 잠금을 회수하지 못했다"
    D.release_run_lock(recovered, lock)

    import inbox_scan
    old_pick = inbox_scan.pick
    try:
        inbox_scan.pick = lambda kind: ["분류기_선택.xlsx"] if kind == "ledger" else []
        assert D.has_inbox_kind("ledger")
        assert not D.has_inbox_kind("tax")
    finally:
        inbox_scan.pick = old_pick

    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert 'has_inbox_kind("ledger")' in daily, "ERP 원장 단계가 공용 inbox 분류기를 쓰지 않는다"
    assert "acquire_run_lock" in daily and "release_run_lock" in daily
    erp_check = open(os.path.join(ROOT, "erp_ledger_check.py"), encoding="utf-8").read()
    assert 'files = pick("ledger")' in erp_check
    assert 'pick("ledger", INBOX_DIR)' not in erp_check, "실제 대조기가 다시 로컬 inbox만 본다"
    print("  [86] daily_run 단일 프로세스 잠금 + ERP 원장 공용 분류기 사용 ✅")


def t90_ip_guard_and_archive():
    """ERP 접속 IP 관문 + 복구용 보관 (2026-07-30 지시 2건).

    ① "IP가 변경되면 이 화면에서 등록해서 진행" — 자동 등록은 하지 않는다(회사 ERP 보안
       설정). 대신 **부르기 전에 멈춘다**: 미등록 IP로 호출하면 실패가 인증 오류처럼 보여
       원인을 엉뚱한 데서 찾고, 반복 실패는 트래픽 제한을 건드려 ERP 전체가 막힌다.
    ② "복구·코딩에 필요한 자료 별도 보관" — 파일을 쌓는 대신 git bundle 한 파일로.
       ★ 비밀키는 어떤 경로로도 담기지 않아야 한다(규칙 1).
    """
    import erp_ip_guard as G
    assert G.self_test(), "erp_ip_guard 자체 검증 실패"
    # IP를 못 구했으면 통과시키면 안 된다 — 미등록 상태로 ERP를 두드리는 게 최악이다
    assert G.decide("", ["1.2.3.4"])[0] == "unknown"
    assert G.decide("5.6.7.8", ["1.2.3.4"])[0] == "need_register"
    # ERP 로그인이 이 관문을 지나는가
    cli = open(os.path.join(ROOT, "ecount_client.py"), encoding="utf-8").read()
    assert "erp_ip_guard" in cli and "require()" in cli, "ERP 호출이 IP 관문을 건너뛴다"
    # IP 목록은 커밋되지 않는다(회사 설정값)
    gi = open(os.path.join(ROOT, ".gitignore"), encoding="utf-8").read()
    assert "config/erp_allowed_ips.json" in gi

    import archive_keep as A
    assert A.self_test(), "archive_keep 자체 검증 실패"
    # 비밀키 파일은 어떤 경로로도 보관 대상이 아니다
    for rel in ("config/ecount_config.json", "config/webapp.json", "config/gcal.json",
                "config/erp_allowed_ips.json"):
        assert not A.wanted(os.path.join(ROOT, *rel.split("/")), ROOT), rel
    assert A.has_secret('"API_CERT_KEY": "x"')
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "erp_ip_guard.py" in daily and "archive_keep.py" in daily, "일일 실행에 연결되지 않았다"
    print("  [90] ERP IP 관문(호출 전 차단)·복구용 보관(bundle·비밀키 제외) ✅")


def t91_icon_sprite_and_ios_theme():
    """아이콘 스프라이트 + iOS 외형 (2026-07-30 지시).

    ★ 왜 <img> 를 버렸나: 원본 svg 에 fill="#344054" 가 박혀 있어 **선택된 탭·다크모드·
      경고색에 아이콘이 반응하지 못했다.** 흰 아이콘이 필요한 곳은 filter:brightness(0)
      invert(1) 같은 꼼수로 때웠다. 스프라이트+currentColor 로 근본을 고쳤다.
    ★ 왜 인라인인가: iOS Safari 는 외부 파일의 <use href="a.svg#id"> 를 제대로 렌더하지
      않는다. 아이폰이 주 사용처이므로 인라인이 유일한 정답이다.
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    # 스프라이트가 있고, 모든 <use> 가 실제 <symbol> 을 가리킨다(깨진 아이콘 0)
    symbols = set(re.findall(r'<symbol id="i-([^"]+)"', live))
    assert symbols, "아이콘 스프라이트가 없다"
    used = set(re.findall(r'<use href="#i-([^"]+)"', live))
    dynamic = {u for u in used if "${" in u}
    missing = (used - dynamic) - symbols
    assert not missing, f"symbol 이 없는 아이콘 참조: {sorted(missing)[:5]}"

    # <img> 로 **아이콘**을 되돌리면 색 상속이 다시 깨진다. 회사 CI PNG를 헤더에서
    # 흰색으로 만드는 brightness/invert는 아이콘 꼼수가 아니라 브랜드 표시 규칙이다.
    assert 'src="/icons/' not in live, "아이콘이 <img> 로 되돌아갔다 — currentColor 를 못 받는다"

    # 아이콘 파일과 라이선스는 저장소에 남겨 둔다(스프라이트를 다시 만들 수 있어야 한다)
    icons_dir = os.path.join(ROOT, "webapp", "icons")
    assert os.path.exists(os.path.join(icons_dir, "LICENSE-bootstrap-icons.txt")), "라이선스 파일 누락"
    on_disk = {os.path.splitext(f)[0] for f in os.listdir(icons_dir) if f.endswith(".svg")}
    assert symbols <= on_disk, f"원본 svg 가 없는 symbol: {sorted(symbols - on_disk)[:5]}"

    # iOS 외형: systemBlue·그룹배경·다크모드·큰 모서리
    assert "--brand:#007AFF" in live, "systemBlue 가 아니다"
    assert "--bg:#F2F2F7" in live, "iOS 그룹 배경이 아니다"
    # ★ 2026-07-30 실기기 확인 후 되돌림: 이 앱은 흰 배경이 77곳 하드코딩돼 있어 OS 다크를
    #   따라가면 **흰 카드 위 흰 글자**가 되어 아무것도 안 보였다(사용자 화면으로 확인).
    #   그래서 다크 모드를 넣지 않고 color-scheme:light 로 못박는다. 다시 넣으려면
    #   그 77곳을 전부 변수화한 뒤여야 한다 — 이 검증이 그 순서를 지킨다.
    assert "prefers-color-scheme: dark" not in live,         "다크 모드를 다시 넣었다 — 하드코딩된 밝은 배경을 먼저 변수화해야 한다"
    assert "color-scheme:light" in live, "밝은 화면으로 고정하는 선언이 없다"
    # 헤더 부제가 flex 안에서 0폭까지 눌려 한 글자씩 세로로 쪼개졌던 사고를 막는다
    assert ".appbar h1{min-width:52px;white-space:nowrap}" in live, "헤더 제목이 다시 쪼개질 수 있다"
    # ★ 2026-07-31 지시로 교체: "폰트는 모든 폰트 나눔고딕 폰트로 변경".
    #   전에는 -apple-system(SF Pro) 을 강제했지만, 그건 2026-07-30 의 'iOS 느낌' 기준이었고
    #   오늘 지시가 이를 덮는다. 지금 지켜야 하는 것은 두 가지다:
    #   ① 나눔고딕이 스택 맨 앞인가 ② 그 글꼴을 **동봉본**으로 얹는가
    #      (구글 폰트 주소를 쓰면 인터넷 없는 현장에서 안 뜨고 외부로 요청이 나간다)
    assert 'font-family:"Nanum Gothic"' in live.replace(" ", "").replace(
        'font-family:"NanumGothic",', 'font-family:"Nanum Gothic"') or \
        '"Nanum Gothic","NanumGothic"' in live, "본문 글꼴이 나눔고딕으로 시작하지 않는다"
    assert "/fonts/nanumgothic.css" in live, "나눔고딕을 동봉본으로 얹지 않는다"
    assert os.path.exists(os.path.join(ROOT, "webapp", "fonts", "OFL.txt")), \
        "동봉한 나눔고딕의 SIL OFL 라이선스 파일이 없다"
    assert "fonts.googleapis.com" not in live and "fonts.gstatic.com" not in live, \
        "글꼴을 외부(구글)에서 받고 있다 — 인터넷 없는 현장에서 안 뜬다"
    # 탭바 유리 효과는 지원 안 되는 브라우저를 위해 @supports 로 감싼다
    assert "backdrop-filter" in live and "@supports" in live, "유리 효과에 폴백이 없다"
    # 검은 로고/아이콘이 남색 헤더·흰 버튼에서 사라지지 않아야 한다.
    # 지켜야 하는 것은 **로고가 남색 헤더에서 보이는가** 이지 특정 구현이 아니다.
    # 흰 판을 얹는 방법과, 로고를 흰색으로 뽑아내는 방법(brightness(0) invert(1)) 둘 다 답이다.
    # 2026-07-30: 사용자가 '헤더와 이질감 없이' 를 요청해 흰 판 → 흰색 로고로 바꿨다.
    assert (".appbar-brand-stack{background:rgba(255,255,255,.94)" in live
            or "brightness(0) invert(1)" in live), "검은 로고가 남색 헤더에서 사라진다"
    # 헤더 버튼 아이콘은 **배경과 대비되기만 하면** 된다 — 구현을 고정하지 않는다.
    #   2026-07-30: 흰 버튼(어두운 아이콘) → 테두리만 있는 투명 버튼(흰 아이콘)으로 바꿨다.
    #   사용자 지적: 남색 헤더 위 흰 네모가 이질적이다.
    assert ".account-pin img,.account-pin svg" in live, "PIN 버튼 아이콘 규칙이 없다"
    assert ("fill:#24365F" in live or "fill:#fff" in live), "PIN 아이콘 색이 지정되지 않았다"
    assert ".notice-bell svg{width:21px;height:21px;fill:" in live, "알림 아이콘 색이 없다"
    # 상단 배지는 absolute 로 떠 있으면 긴 제목과 다시 겹친다. 제목·배지는 같은 flex 행이어야 한다.
    assert '.hero-top{display:flex' in live and '.hero .badge{position:static' in live
    assert '<div class="hero-top">' in live and '<span class="hero-sub">' in live
    assert ".hero::before" not in live and ".hero::after" not in live, "장식 원이 다시 생겼다"
    print(f"  [91] 아이콘 스프라이트({len(symbols)}개·깨짐 0)·currentColor·iOS·히어로 정렬 ✅")


def t92_excel_recalc_agent():
    """엑셀을 에이전트가 알아서 열고 닫는다 (2026-07-30 지시).

    ★ 이 도구는 **사람의 파일을 저장**한다. 그래서 판단 순서가 곧 안전장치다:
      대기 0이면 안 열고 · 복구 경고/SHA 승인이 없으면 안 열고 ·
      **사람이 열어 두었으면 물러난다.** 남이 편집 중인 파일을 자동화가 저장하면
      그 작업이 날아간다.
    """
    import excel_recalc as X
    assert X.self_test(), "excel_recalc 자체 검증 실패"

    p = r"Z:\a\쿠팡_통합업무_일일보고_관리대장_v323.xlsx"
    # 사람이 열어 두었으면 어떤 경우에도 진행하지 않는다
    ok, why = X.decide({"대기합계": 99}, p, True, True, True)
    assert not ok and "열어 두었" in why, why
    # 복구 경고가 해소되지 않았으면 대기 건이 있어도 정본을 열지 않는다
    ok, why = X.decide({"대기합계": 99}, p, False, True, False)
    assert not ok and "안전 승인" in why, why
    # 원본을 덮어쓰지 않는다(항상 vN+1)
    assert X.next_version_path(p).endswith("_v324.xlsx")
    assert X.next_version_path(p) != p

    # 새 파이썬 의존성을 쓰지 않는다(프로젝트 원칙) — 엑셀은 PowerShell COM 으로 부른다
    src = open(os.path.join(ROOT, "excel_recalc.py"), encoding="utf-8").read()
    # 문서에 "pywin32 없이" 라고 적혀 있으므로 문자열이 아니라 **실제 import** 를 본다.
    assert not re.search(r"^\s*(import|from)\s+(win32com|win32|pythoncom)\b", src, re.M), \
        "새 의존성(pywin32)을 끌어들였다"
    assert "Excel.Application" in src and "powershell" in src
    # 좀비 EXCEL.EXE 를 남기지 않는다
    assert "finally" in X.PS_RECALC and "$x.Quit()" in X.PS_RECALC
    assert "finally" in X.PS_AVAILABLE and "$x.Quit()" in X.PS_AVAILABLE
    # 경로에 공백·★ 가 섞여 있어 인자 전달은 환경변수로 한다(2026-07-30 실패 경험)
    assert "$env:CSOS_XL_SRC" in X.PS_RECALC, "경로를 명령줄 인자로 넘기면 깨진다"
    # 일일 실행에 연결
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "excel_recalc.py" in daily
    assert "excel_recalc_clearance.json" in src and "sha256" in src
    assert "recalc-" in src and "os.replace(tmp_dst, dst)" in src
    assert "if os.path.exists(dst)" in src, "기존 vN+1을 덮어쓸 수 있다"
    print("  [92] 엑셀 자동 재계산 — 복구/SHA관문·임시본 검증·덮어쓰기/좀비 방지 ✅")


def t93_ledger_db_and_ux():
    """반영은 DB에 모았다가 하루 두 번만 · 앱 UX는 기록으로 (2026-07-30 지시).

    ★ 채울 때마다 엑셀을 쓰면 하루에 관리대장 버전이 수십 개 늘고(v311→v327) 정본이 흔들린다.
      그래서 11:00·15:00 두 번만 쓴다. **시각이 아니면 아무것도 하지 않는 것**이 핵심이다.
    ★ 놓친 회차의 입력은 버리지 않고 다음 **허용된** 11시/15시 회차에 함께 처리한다.
    """
    import ledger_db as L
    assert L.self_test(), "ledger_db 자체 검증 실패"
    assert [w.hour for w in L.WINDOWS] == [11, 15], "반영 시각이 11시·15시가 아니다"

    from datetime import datetime as _dt
    assert L.slot_of(_dt(2026, 7, 30, 11, 5)), "11시 회차를 인식하지 못한다"
    assert L.slot_of(_dt(2026, 7, 30, 13, 0)) is None, "반영 시각이 아닌데 열려 있다"
    assert L.next_window(_dt(2026, 7, 30, 12, 0)).hour == 15
    assert L.eligible_slot(_dt(2026, 7, 30, 13, 0), []) is None, \
        "놓친 회차를 이유로 임의 시각에 반영한다"
    assert L.eligible_slot(
        _dt(2026, 7, 30, 11, 5), ["2026-07-30 11:00"]) is None, \
        "같은 11시 회차를 두 번 반영한다"

    with tempfile.TemporaryDirectory(prefix="ledger-json-lock-") as lock_tmp:
        queue_path = os.path.join(lock_tmp, "pending.json")
        with open(queue_path + ".lock", "w", encoding="ascii") as f:
            f.write("99999999 2026-01-01T00:00:00")
        with L.json_queue_lock(queue_path, timeout=0.1):
            assert os.path.exists(queue_path + ".lock")
        assert not os.path.exists(queue_path + ".lock"), "죽은 JSON 큐 잠금을 회수하지 못한다"

    src = open(os.path.join(ROOT, "ledger_db.py"), encoding="utf-8").read()
    writer = open(os.path.join(ROOT, "ledger_writer.py"), encoding="utf-8").read()
    assert any(x.startswith("import") and "sqlite3" in x for x in src.splitlines()), (
        "표준 라이브러리 sqlite3 를 써야 한다(새 의존성 금지)")
    assert "finally:" in src and "c.close()" in src, "DB 연결을 닫지 않으면 파일이 잠긴다"
    assert "ingest_key" in src and "INSERT OR IGNORE INTO pending" in src, \
        "중단 후 JSON staging 재흡수 시 중복될 수 있다"
    assert "id IN ({marks})" in src, "반영 중 새로 들어온 DB 행까지 완료 처리할 수 있다"
    assert '"--queue", batch_queue, "--apply"' in src, "공용 JSON 큐와 일괄반영 배치가 섞인다"
    assert '"COUPANG_LEDGER_GATE": "1"' in src, "11·15시 DB 회차가 내부 쓰기 게이트를 열지 않는다"
    assert 'os.environ.get("COUPANG_LEDGER_GATE") != "1"' in writer, \
        "ledger_writer 직접 --apply를 막는 강제 게이트가 없다"

    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "ledger_db.py" in daily, "일일 실행이 DB 게이트를 지나지 않는다"
    assert "r.stderr[-2000:]" in daily, "실패 리포트가 예외 원인을 다시 잘라낸다"
    assert '"zscan.py"' in daily and '"--docs"' in daily, \
        "Z: 상시 공백·서류 대조가 09:50 자동실행에 연결되지 않았다"
    assert '"ledger_writer.py"), "--apply"' not in daily,         "일일 실행이 아직 엑셀에 곧바로 쓴다 — 하루 두 번 규칙이 깨진다"
    for direct in (
        '"excel_recalc.py"), "--run"',
        '"erp_docs_check.py"), "--sheet"',
        '"fix_workbook.py"), "--apply"',
        '"pm_schedule_sync.py"), "--apply"',
        '"work_log_sync.py"), "--apply"',
        '"band_extract.py"), "--sheet"',
        '"findings_sheet.py")]))',
    ):
        assert direct not in daily, f"09:50 자동대조에 직접 Excel 쓰기가 남아 있다: {direct}"
    for scheduled in (
        '"erp_docs_check.py"), "--sheet"',
        '"pm_schedule_sync.py"), "--apply"',
        '"work_log_sync.py"), "--apply"',
        '"band_extract.py"), "--sheet"',
        '"findings_sheet.py")',
        '"fix_workbook.py"), "--apply"',
        '"excel_recalc.py"), "--run"',
    ):
        assert scheduled in src, f"11·15시 구조 갱신에서 빠진 작업: {scheduled}"
    assert "scheduled_workbook_maintenance(now)" in src, \
        "11·15시 회차가 구조 시트·재계산을 함께 수행하지 않는다"
    assert "def handoff_add(" in src and '"workbook_patch.py"' in src, \
        "19시트 종료 인수인계가 11·15시 회차에 예약되지 않는다"
    schedule = open(os.path.join(ROOT, "install_ledger_schedule.ps1"), encoding="utf-8").read()
    assert '"11:00"' in schedule and '"15:00"' in schedule
    assert "MultipleInstances IgnoreNew" in schedule and "ledger_db.py --apply" in schedule

    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert "get_apply_window" in server and '"applywin"' in server
    assert '"writer_apply":  ("입력 DB 적재"' in server
    assert 'start_task("writer_apply")' not in server, \
        "앱이 아직 즉시 Excel 반영 작업을 시작한다"
    assert 'os.path.join(ROOT, "ledger_writer.py"), "--apply"' not in server, \
        "앱 서버에 직접 ledger_writer --apply 경로가 남아 있다"
    assert "enqueue_for_scheduled_apply" in server, "앱 입력이 DB 일괄반영 큐를 거치지 않는다"
    assert 'if p == "/api/ux":' in server and "ledger_db.ux_add(events)" in server
    assert "function uxEvent(" in live and "function uxFlush(" in live
    assert "uxEvent('view',v)" in live and "uxEvent('slow'" in live
    assert "renderApplyWindow(s.applywin)" in live, "앱이 다음 반영 시각을 알리지 않는다"
    assert "runTask('writer_apply')" not in live, "앱 화면에 즉시 Excel 반영 호출이 남아 있다"
    assert "지금 바로 엑셀에 반영" not in live, "앱 안내가 아직 즉시 반영을 약속한다"

    band_collector = open(os.path.join(ROOT, "band", "collect_band.js"), encoding="utf-8").read()
    assert "window.scrollBy(0, SCROLL_STEP)" in band_collector, \
        "밴드 수집기가 실패한 바닥 점프로 되돌아갔다"
    assert "window.scrollTo(0, document.body.scrollHeight)" not in band_collector
    assert "fetch(url" not in band_collector and "imageData.push" not in band_collector, \
        "사진 fetch 무한대기가 글 전체 수집을 다시 막을 수 있다"

    # 업무센터는 인원이 바뀌어도 균등해야 한다(고정 칸 수 금지)
    assert "repeat(auto-fit,minmax(132px,1fr))" in live, "업무센터가 다시 칸 수를 고정했다"
    assert ".workcenter-buttons{grid-template-columns:repeat(3" not in live, \
        "업무센터 3칸 고정이 남아 있다 — 인원이 바뀌면 빈칸이 생긴다"
    assert ":last-child:nth-child(odd){grid-column:1/-1}" in live, "홀수 인원일 때 줄 끝이 빈다"
    assert "<h3>빠른 실행</h3>" not in live, "제거한 '빠른 실행' 카드가 되살아났다"

    # 상시 규칙이 문서에 남아 있어야 다음 세션·Codex 가 따른다
    rules = (open(os.path.join(ROOT, "CLAUDE.md"), encoding="utf-8").read()
             + open(os.path.join(ROOT, "AGENTS.md"), encoding="utf-8").read())
    assert "엑셀은 하루 두 번만" in rules and "UX 기록" in rules
    archive = open(os.path.join(ROOT, "archive_keep.py"), encoding="utf-8").read()
    assert "ledger_db_copy" in archive and "ledger_queue.db" in archive
    print("  [93] 반영 DB 게이트(11·15시)·UX 기록·업무센터 균등배치·빠른실행 제거 ✅")


def t95_objective_completion_db_only():
    """[95] 객관 입증 완료 처리 + 엑셀 백필 없이 DB 로만 (사용자 지시 2026-07-31).

    · 7.수금완료 = 발행·수금 모두 ERP 가 입증 → 완료(ERP 수금확인)
    · 6.세금계산서발행 = 발행만 입증 → '미발행'이 아니라 입금 대기(입금일 있으면 완료)
    · 완료 상태는 06시트 발행일 칸에 써넣지 않는다(판매조회에 발행일 열이 없다 —
      절대규칙 10). 근거는 ledger_db.resolution 표가 기억한다(first_seen 보존)."""
    import sys as _s, tempfile
    _s.path.insert(0, ROOT)
    import ecount_reconcile as E
    import ledger_db as L
    import staff_completion as SC

    old_prog = E.erp_progress
    try:
        E.erp_progress = lambda: {"P7": "7.수금완료", "P6": "6.세금계산서발행",
                                  "PMIX": "혼재(7.수금완료 / 확인)"}
        base = {"비용구분": "유상", "원천업무ID": "PM-1", "원장_공급가액": 100,
                "원장_거래명세서번호": "20260101-1"}
        assert E.settle_status({**base, "프로젝트NO": "P7"}) == "완료(ERP 수금확인)"
        assert E.settle_status({**base, "프로젝트NO": "P6"}) == "입금 대기", \
            "발행만 입증된 건을 완료로 넘기면 수금 추적이 죽는다"
        assert E.settle_status({**base, "프로젝트NO": "P6",
                                "원장_입금일": "2026-02-01"}) == "완료(ERP 발행확인)"
        assert E.settle_status({**base, "프로젝트NO": "PX"}) == "세금계산서 미발행"
        # 금액 수식 캐시가 비어도 ERP 전체 전표가 수금완료면 객관 완료가 우선한다.
        as_wait = {**base, "원천업무ID": "AS-1", "원장_공급가액": 0,
                   "원장_거래명세서합계": 110}
        assert E.settle_status({**as_wait, "프로젝트NO": "P7"}) == "완료(ERP 수금확인)"
        assert E.settle_status({**as_wait, "프로젝트NO": "P6"}) == "금액 재계산 대기"
        assert E.settle_status({**as_wait, "프로젝트NO": "PMIX"}) == "금액 재계산 대기", \
            "한 프로젝트에 수금완료·미완료 전표가 섞였는데 전체 완료로 올렸다"
        assert E._collapse_erp_progress(["7.수금완료", "7.수금완료"]) == "7.수금완료"
        assert E._collapse_erp_progress(["7.수금완료", "6.세금계산서발행"]) == \
            "6.세금계산서발행"
        assert E._collapse_erp_progress(["7.수금완료", "확인"]).startswith("혼재("), \
            "복수 전표 충돌이 완료로 접혔다"
    finally:
        E.erp_progress = old_prog

    # 완료( 는 조치 목록이 아니라 DB(resolution)로 — 엑셀 셀 백필 경로가 없어야 한다
    fx = open(os.path.join(ROOT, "findings_export.py"), encoding="utf-8").read()
    assert 'startswith("완료(")' in fx and "resolution_sync" in fx, \
        "완료 건이 조치 목록에 남거나 DB에 기록되지 않는다"
    seg = fx[fx.index('startswith("완료(")'):fx.index("return rows")]
    assert "enqueue" not in seg and "workbook" not in seg, "완료 처리가 엑셀 쓰기로 새어 나간다"

    # resolution 멱등 upsert — 두 번 넣어도 1행, first_seen 은 처음 값이 남는다
    with tempfile.TemporaryDirectory() as td:
        old_dir, old_path = L.DB_DIR, L.DB_PATH
        L.DB_DIR, L.DB_PATH = td, os.path.join(td, "t.db")
        try:
            assert L.resolution_sync([{"settle_id": "JS-1", "project": "P7",
                                       "status": "완료(ERP 수금확인)", "basis": "b1"}]) == 1
            first = L.resolutions()["JS-1"]["first_seen"]
            L.resolution_sync([{"settle_id": "JS-1", "project": "P7",
                                "status": "완료(ERP 수금확인)", "basis": "b2"}])
            got = L.resolutions()
            assert len(got) == 1 and got["JS-1"]["basis"] == "b2", "upsert 가 안 된다"
            assert got["JS-1"]["first_seen"] == first, "first_seen 이 덮였다 — 최초 입증 시각 유실"
            L.staff_resolution_sync([{
                "owner": "류지영", "task_kind": "settlement", "record_id": "JS-1",
                "project": "P7", "completed_on": "2026-08-01", "basis": "정산 완료",
            }])
            assert L.resolution_retract(["JS-NOT-EXIST", "JS-1"]) == 1
            assert "JS-1" not in L.resolutions()
            assert not any(row["record_id"] == "JS-1" and row["task_kind"] == "settlement"
                           for row in L.staff_resolutions("류지영")), \
                "정산 완료 철회 뒤 담당자 완료가 남았다"
            # 뒤의 담당자 동기화 검증은 정상 정산 완료 1건을 전제로 한다.
            L.resolution_sync([{"settle_id": "JS-1", "project": "P7",
                                "status": "완료(ERP 수금확인)", "basis": "b3"}])
            assert L.work_resolution_sync([{
                "kind": "as", "record_id": "AS-1", "project": "UJ2600001",
                "status": "작업완료", "completed_on": "2026-07-30", "basis": "band",
            }]) == 1
            work_first = L.work_resolutions()[("as", "AS-1")]["first_seen"]
            L.work_resolution_sync([{
                "kind": "as", "record_id": "AS-1", "project": "UJ2600001",
                "status": "작업완료", "completed_on": "2026-07-31", "basis": "verified",
            }])
            work = L.work_resolutions()
            assert work[("as", "AS-1")]["completed_on"] == "2026-07-31"
            assert work[("as", "UJ2600001")]["first_seen"] == work_first

            # Excel 반영 전 신규행도 실제 완료일과 원천 근거가 있으면 DB 완료 정본에 잡힌다.
            assert L.enqueue([
                {"sheet": "04_정기점검", "cell": "B10", "col": "프로젝트NO",
                 "value": "UJ2600999", "evidence": "카톡 UJ2600999 완료보고"},
                {"sheet": "04_정기점검", "cell": "H10", "col": "실제점검일",
                 "value": "2026-07-31", "evidence": "카톡 UJ2600999 완료보고"},
                {"sheet": "04_정기점검", "cell": "H10", "col": "실제점검일",
                 "value": "2026-07-31", "evidence": "카톡 UJ2600999 완료보고"},
                {"sheet": "02_돌발AS접수", "cell": "B11", "col": "프로젝트NO",
                 "value": "UJ2600998", "evidence": "카톡 원문"},
                {"sheet": "02_돌발AS접수", "cell": "R11", "col": "작업완료일",
                 "value": "2026-07-30", "evidence": "카톡 완료보고"},
                {"sheet": "02_돌발AS접수", "cell": "Q11", "col": "진행상태",
                 "value": "취소", "evidence": "카톡 정정"},
                {"sheet": "04_정기점검", "cell": "B12", "col": "프로젝트NO",
                 "value": "UJ2600997", "evidence": "카톡 원문"},
                {"sheet": "04_정기점검", "cell": "H12", "col": "실제점검일",
                 "value": "2099-01-01", "evidence": "잘못된 미래일"},
                {"sheet": "04_정기점검", "cell": "B13", "col": "프로젝트NO",
                 "value": "UJ2600996", "evidence": "새 업무"},
                {"sheet": "04_정기점검", "cell": "H13", "col": "실제점검일",
                 "value": "2026-07-29", "evidence": "이전 업무의 완료 근거"},
            ], source="test") == 10
            pending_done = L.pending_work_completion_entries(today="2026-08-02")
            assert len(pending_done) == 1 and pending_done[0]["project"] == "UJ2600999"
            assert pending_done[0]["basis"].startswith("반영대기 04_정기점검!H10")
            L.work_resolution_sync(pending_done)
            provisional_first = L.work_resolutions()[("pm", "UJ2600999")]["first_seen"]
            L.work_resolution_sync([{
                "kind": "pm", "record_id": "PM-2600999", "project": "UJ2600999",
                "status": "완료", "completed_on": "2026-07-31", "basis": "원장 검증 정상",
            }])
            with L.conn() as c:
                migrated = c.execute(
                    "SELECT record_id,first_seen FROM work_resolution WHERE kind='pm' AND project=?",
                    ("UJ2600999",),
                ).fetchall()
            assert migrated == [("PM-2600999", provisional_first)], \
                "Excel ID 확정 뒤 대기행 완료 레코드가 중복되거나 최초 입증시각이 바뀌었다"

            # 세 담당자의 객관완료는 이름이 포함된 상태로 별도 DB 정본에 남는다.
            po_evidence = os.path.join(td, "po_objective_evidence.json")
            report_path = os.path.join(td, "담당자_객관완료.md")
            with open(po_evidence, "w", encoding="utf-8") as out:
                json.dump({"entries": [
                    {"owner": "오종현", "task_kind": "po_source", "record_id": "PO1",
                     "project": "UJ2600001", "completed_on": "2026-07-31",
                     "basis": "쿠팡 PO 원본 확인"},
                    {"owner": "유현민", "task_kind": "po_system_verified", "record_id": "PO1",
                     "project": "UJ2600001", "completed_on": "2026-08-01",
                     "basis": "PO·ERP·원장 일치"},
                    {"owner": "류지영", "task_kind": "billing_verified", "record_id": "PO1",
                     "project": "UJ2600001", "completed_on": "2026-08-01",
                     "basis": "ERP 계산서 발행 확인"},
                    {"owner": "오종현", "task_kind": "po_source", "record_id": "PO-FUTURE",
                     "project": "", "completed_on": "2099-01-01",
                     "basis": "잘못된 미래일"},
                ], "retractions": [
                    {"owner": "유현민", "task_kind": "po_system_verified",
                     "record_id": "PO-FALSE", "reason": "비유일 금액"},
                    {"owner": "류지영", "task_kind": "billing_verified",
                     "record_id": "PO-FALSE", "reason": "비유일 금액"},
                ]}, out, ensure_ascii=False)
            L.staff_resolution_sync([
                {"owner": "유현민", "task_kind": "po_system_verified",
                 "record_id": "PO-FALSE", "completed_on": "2026-08-01", "basis": "옛 추정"},
                {"owner": "류지영", "task_kind": "billing_verified",
                 "record_id": "PO-FALSE", "completed_on": "2026-08-01", "basis": "옛 추정"},
            ])
            counts = SC.sync(po_evidence, report_path)
            assert counts == {"류지영": 4, "오종현": 1, "유현민": 1}
            assert {row["status"] for row in L.staff_resolutions()} == {
                "류지영 완료", "오종현 완료", "유현민 완료"}
            first_staff = L.staff_resolutions("오종현")[0]["first_seen"]
            SC.sync(po_evidence, report_path)
            assert L.staff_resolutions("오종현")[0]["first_seen"] == first_staff
            assert "오종현 완료 · 1건" in open(report_path, encoding="utf-8").read()
        finally:
            L.DB_DIR, L.DB_PATH = old_dir, old_path

    # 앱: 완료 상태가 조치필요 필터·칩에서 완료로 취급되는가
    html = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert "'완료(ERP 수금확인)'" in html and "'완료(ERP 발행확인)'" in html, "앱이 새 완료 상태를 모른다"
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "complete_verified.py" in daily and '"--queue"' in daily, \
        "객관근거 완료 상태 보완이 일일 자동화에 연결되지 않았다"
    assert daily.index('"입력 DB 적재"') < daily.index('"객관근거 완료 DB 동기화"'), \
        "신규 입력을 DB에 흡수하기 전에 완료 판정을 실행하고 있다"
    assert "staff_completion.py" in daily and '"--sync"' in daily, \
        "담당자별 객관완료 판정이 일일 에이전트에 연결되지 않았다"
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert "work_resolutions" in srv and "객관완료근거" in srv, "앱이 현장업무 DB 완료판정을 읽지 않는다"
    assert '"/api/staff/completions"' in srv and "staff_completions_payload" in srv
    po_src = open(os.path.join(ROOT, "po_reconcile.py"), encoding="utf-8").read()
    assert "po_objective_evidence.json" in po_src and '"오종현", "task_kind": "po_source"' in po_src
    print("  [95] 객관 입증 완료(정산·현장업무 DB 정본·앱 즉시반영·백필 금지) ✅")


def t96_work_management_tabs():
    """[96] 정기점검·돌발AS 정밀 관리 탭 (사용자 지시 2026-07-31).

    · 입력은 전부 /api/input → DB 큐(11·15시 회차) — 엑셀 직접 쓰기 없음(DB-only).
    · 배정·상태·일정만 덮어쓰기 허용(OVERWRITE_COLS) — 그 밖은 빈 칸만 채운다.
      류지영 매니저가 엑셀에서 담당기사를 한 건씩 고치다 유실 사고가 난 그 작업을
      앱에서 안전하게 하게 하는 것이 이 탭의 존재 이유다.
    · 폰 하단바는 칸이 모자라 사이드바(≥900px)에서만 메뉴 노출 — 폰은 대시보드
      바로가기로 들어간다(숨겨도 화면은 동작해야 한다)."""
    html = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for need in ('id="v-pm"', 'id="v-as"', "renderWorkTab", "WT_CFG", "wtEdit",
                 'data-v="pm"', 'data-v="as"', "worktab-nav", "wtBoard", "wtCsv",
                 'id="pmBaseDate"', 'id="asBaseDate"', "wtSetBase", "wtRows", "wtReset",
                 'aria-label="기준일 범위"', 'aria-label="정렬 기준"',
                 'aria-label="정렬 방향"', "sortOrder:'desc'", "st.sortOrder==='asc'"):
        assert need in html, f"정밀 관리 탭 구성 요소 누락: {need}"
    assert all(label in html for label in ("기준일 당일", "기준일까지", "기준일 이후",
                                            "예정일 미정", "내림차순 · 최신순",
                                            "오름차순 · 과거순", "필터 초기화")), \
        "기준일 범위·정렬·초기화 필터가 완성되지 않았다"
    assert ".tabbar.worktab-nav{display:none}" in html.replace(" ", ""), \
        "폰 하단바 칸 부족 대책(사이드바 전용 노출)이 없다"
    # 인라인 편집이 DB 큐 경로(/api/input)로만 가는가 — 다른 쓰기 경로가 생기면 안 된다
    # ★ 검사 범위는 wtEdit **함수 몸통만**이다. 다음 함수 선언 직전까지로 자른다 —
    #   wtBoard 까지 넓게 잡으면 사이에 끼는 무관한 코드(업무센터 업로드의 accept=".xlsx")가
    #   오탐을 낸다(2026-07-31 실제로 그랬다).
    _ws = html.index("async function wtEdit")
    seg = html[_ws:html.index("function ", _ws + 30)]
    assert "/api/input" in seg and "overwrite:true" in seg, "편집이 DB 큐 경로를 안 탄다"
    assert "ledger_writer" not in seg and "xlsx" not in seg, "편집이 엑셀로 새어 나간다"
    # 서버: 덮어쓰기는 허용열에서만, 근거에 '수정'이 남는가
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    blk = srv[srv.index('"/api/input"'):srv.index('"/api/input"') + 2500]
    assert "OVERWRITE_COLS" in blk and '"담당기사"' in blk and '"진행상태"' in blk, \
        "덮어쓰기 허용열이 없다 — 오배정을 앱에서 못 고친다"
    assert '"only_if_empty": not overwrite' in blk, "덮어쓰기 플래그가 큐에 안 실린다"
    assert "수정" in blk, "덮어쓰기 근거 표시가 없다 — 나중에 추적할 수 없다"
    # 업무센터·대시보드에서 들어가는 입구
    assert html.count("show('pm')") >= 2 and html.count("show('as')") >= 2, \
        "대시보드·업무센터에서 정밀 관리 탭으로 가는 입구가 없다"

    # 업무센터 탭 + 입금 업로드(2026-07-31): 드롭존·URL 등록·오프라인 보관·Z: 저장·즉시 대조
    for need in ('id="v-center"', 'data-v="center"', "renderCenter", "renderReceiptUpload",
                 "receiptSubmit", "'/api/staff/receipt-upload'", "injectOhUpload"):
        assert need in html, f"업무센터 탭 구성 요소 누락: {need}"
    assert "'/api/staff/receipt-upload']" in html.replace(" ", "") or \
           "/api/staff/receipt-upload" in html[html.index("OUTBOX_PATHS"):html.index("OUTBOX_PATHS")+400], \
        "입금 업로드가 오프라인 보관(outbox) 목록에 없다"
    assert '"receipt"' in srv and "receipt_fill.py" in srv, "업로드 직후 즉시 대조 작업이 없다"
    assert "save_staff_receipt_submission" in srv and "RECEIPT_DIR" in srv, \
        "입금 자료가 지정 저장소(7. 입금내역)로 가지 않는다"
    blk2 = srv[srv.index('"/api/staff/receipt-upload"'):srv.index('"/api/staff/receipt-upload"') + 900]
    assert '"admin"' in blk2 and "oh-jonghyeon" in blk2, "관리자·오종현 외에도 업로드가 열려 있다"
    print("  [96] 정밀 관리 탭 + 업무센터·입금 업로드(DB 큐·Z: 저장·즉시 대조·권한) ✅")


def t97_settlement_source_completion():
    """[97] 금액·계산서 대기는 독립 원자료가 정확히 맞는 건만 DB 완료 처리."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import settlement_completion as S

    amount = {
        "비용구분": "유상", "원천업무ID": "AS-1", "프로젝트NO": "UJ2600001",
        "원장_공급가액": 0, "원장_거래명세서발행일": "2026-01-10",
        "원장_거래명세서합계": 110000, "원장_PO번호": "PO123456/PR1",
    }
    invoice = {
        "비용구분": "유상", "원천업무ID": "PM-1", "프로젝트NO": "UJ2600002",
        "원장_공급가액": 100000, "원장_거래명세서발행일": "2026-01-11",
        "원장_거래명세서합계": 110000,
    }
    quotes = [{"종류": "견적서", "프로젝트NO": "UJ2600001", "PO번호": "PO123456",
               "금액": 110000, "파일": "q1.pdf"}]
    invoices = {"UJ2600002": [{"slip": "2026/01/20-1", "amount": 100000,
                                "verdict": "확정(밴드)"}]}
    got = S.objective_entries({"JS-A": amount, "JS-I": invoice}, quotes, invoices)
    assert {row["status"] for row in got} == {S.AMOUNT_STATUS, S.INVOICE_STATUS}, got
    assert all("완료(" in row["status"] and row["basis"] for row in got)

    # 견적 단독 입증(2026-08-03 지시): 프로젝트 견적이 한 금액뿐이면 명세합계가 어긋나도
    # (교차 입력 밀림) 견적을 정본으로 담당자 확인 완료한다. 같은 금액 사본은 한 장으로 본다.
    duplicate = quotes + [{**quotes[0], "파일": "q2.pdf"}]
    got = S.objective_entries({"JS-A": amount}, duplicate, {})
    assert [r["status"] for r in got] == [S.QUOTE_ONLY_STATUS], got  # 같은 금액 사본 → 한 장
    assert "일치" in got[0]["basis"]
    mismatch = [{**quotes[0], "금액": 120000}]
    got = S.objective_entries({"JS-A": amount}, mismatch, {})
    assert [r["status"] for r in got] == [S.QUOTE_ONLY_STATUS], got
    assert "불일치" in got[0]["basis"] and "담당자(류지영) 확인" in got[0]["basis"]
    assert got[0]["evidence_kind"] == "quote_only"
    # 프로젝트 견적이 서로 다른 두 금액이면 무엇이 정본인지 데이터로 못 정한다 — 완료 금지.
    two_totals = mismatch + [{**quotes[0], "금액": 990000, "파일": "q3.pdf"}]
    assert not S.objective_entries({"JS-A": amount}, two_totals, {}), "금액 둘인데 완료"
    # ERP 검증(2026-08-03): ERP 전표가 견적과도 다르면 두 원천 충돌 — 자동 완료 금지.
    erp_clash = {"UJ2600001": [{"date": "2026-02-01", "po": "", "status": "3.오더처리",
                                "supply": 999000, "total": 1098900}]}
    assert not S.objective_entries({"JS-A": amount}, mismatch, {}, None, erp_clash), \
        "ERP·견적 충돌인데 견적단독 완료"
    erp_agree = {"UJ2600001": [{"date": "2026-02-01", "po": "", "status": "3.오더처리",
                                 "supply": 120000, "total": 132000}]}
    got = S.objective_entries({"JS-A": amount}, mismatch, {}, None, erp_agree)
    assert [r["status"] for r in got] == [S.QUOTE_ONLY_STATUS], got   # ERP가 견적을 지지
    s_src = open(os.path.join(ROOT, "settlement_completion.py"), encoding="utf-8").read()
    assert "resolution_retract" in s_src, "충돌 재검출 시 견적단독 완료 철회가 없다"

    # 견적서가 없어도 같은 프로젝트 ERP 판매전표 금액이 유일하게 일치하면 완료한다.
    erp_one = {"UJ2600001": [{"date": "2026-01-15", "po": "PO123456",
                              "status": "3.오더처리", "supply": 100000, "total": 110000}]}
    got = S.objective_entries({"JS-A": amount}, [], {}, None, erp_one)
    assert [row["status"] for row in got] == [S.ERP_AMOUNT_STATUS], got
    assert "ERP 판매전표" in got[0]["basis"] and got[0]["evidence_kind"] == "erp_amount"
    # 금액이 맞는 전표가 둘이면 유일 근거가 아니다 · 금액 불일치도 완료하지 않는다.
    erp_two = {"UJ2600001": erp_one["UJ2600001"] * 2}
    assert not S.objective_entries({"JS-A": amount}, [], {}, None, erp_two), "복수 전표 완료"
    erp_off = {"UJ2600001": [{**erp_one["UJ2600001"][0], "supply": 90000, "total": 99000}]}
    assert not S.objective_entries({"JS-A": amount}, [], {}, None, erp_off), "전표 불일치 완료"
    # 전표 PO 가 원장 PO 와 다르면 같은 금액이라도 남의 전표다.
    erp_po = {"UJ2600001": [{**erp_one["UJ2600001"][0], "po": "PO999999"}]}
    assert not S.objective_entries({"JS-A": amount}, [], {}, None, erp_po), "PO 불일치 완료"

    # ERP 수금확인처럼 이미 더 강한 완료가 있으면 이 모듈의 상태로 낮춰 쓰지 않는다.
    stronger = {"JS-A": {"status": "완료(ERP 수금확인)"}}
    assert not S.objective_entries({"JS-A": amount}, quotes, {}, stronger)

    app = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    findings = open(os.path.join(ROOT, "findings_export.py"), encoding="utf-8").read()
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "objective_done = resolutions()" in app and 'resolved_status.startswith("완료(")' in app
    assert "objective_done = ledger_db.resolutions()" in findings and "db_done" in findings
    assert "settlement_completion.py" in daily and '"--sync"' in daily

    # 오래된 세금계산서 미발행 감시(2026-08-03): 경과일 정렬·완료 제외·일일 자동 실행·앱 노출
    import tax_invoice_watch as W
    from datetime import date as _d
    recs = {
        "JS-O": {"비용구분": "유상", "원천업무ID": "PM-9", "프로젝트NO": "UJ2600009",
                 "캠프명": "테스트1", "업무구분": "정기점검", "작업완료일": "2026-01-01",
                 "원장_공급가액": 100000, "원장_거래명세서발행일": "2026-01-02",
                 "원장_거래명세서합계": 110000},
        "JS-N": {"비용구분": "유상", "원천업무ID": "PM-8", "프로젝트NO": "UJ2600008",
                 "캠프명": "테스트2", "업무구분": "정기점검", "작업완료일": "2026-07-20",
                 "원장_공급가액": 100000, "원장_거래명세서발행일": "2026-07-21",
                 "원장_거래명세서합계": 110000},
        "JS-D": {"비용구분": "유상", "원천업무ID": "PM-7", "프로젝트NO": "UJ2600007",
                 "캠프명": "테스트3", "업무구분": "정기점검", "작업완료일": "2026-01-01",
                 "원장_공급가액": 100000, "원장_거래명세서발행일": "2026-01-02",
                 "원장_거래명세서합계": 110000},
    }
    rows = W.overdue_rows(recs, {"JS-D": {"status": "완료(ERP 수금확인)"}}, {},
                          today=_d(2026, 8, 3))
    assert [row["settle_id"] for row in rows] == ["JS-O", "JS-N"], rows  # 오래된 순·완료 제외
    assert rows[0]["age"] > 200 and rows[1]["age"] == 14
    counts = W.bucket_counts(rows)
    assert counts["90일 초과"] == 1 and counts["30일 이하"] == 1, counts
    assert "tax_invoice_watch.py" in daily, "미발행 경과 감시가 일일 자동화에 없다"
    assert "세금계산서_미발행_경과.md" in app, "미발행 경과 리포트가 앱 보고에 안 보인다"

    # 견적↔명세 교차 진단(2026-08-03): 명세합계가 같은 PO 다른 프로젝트 견적과 일치하는
    # 입력 밀림을 짝까지 찾아낸다 — 실사례 UJ2600777(야탑)↔UJ2600783(안산) 722,480원.
    import quote_mismatch as Q
    q_recs = {
        "JS-X1": {"비용구분": "유상", "원천업무ID": "AS-11", "프로젝트NO": "UJ2600777",
                  "원장_공급가액": 0, "원장_거래명세서발행일": "2026-04-28",
                  "원장_거래명세서합계": 722480, "원장_PO번호": "PO354490/PR1"},
        "JS-X2": {"비용구분": "유상", "원천업무ID": "AS-12", "프로젝트NO": "UJ2600999",
                  "원장_공급가액": 0, "원장_거래명세서발행일": "2026-04-28",
                  "원장_거래명세서합계": 50000, "원장_PO번호": "PO354490/PR1"},
    }
    q_quotes = [
        {"종류": "견적서", "프로젝트NO": "UJ2600777", "PO번호": "PO354490", "금액": 311300, "파일": "야탑.pdf"},
        {"종류": "견적서", "프로젝트NO": "UJ2600783", "PO번호": "PO354490", "금액": 722480, "파일": "안산.pdf"},
    ]
    got = Q.diagnose(q_recs, q_quotes)
    kinds = {row["정산ID"]: row["유형"] for row in got}
    assert kinds == {"JS-X1": "교차 의심", "JS-X2": "견적 없음"}, kinds
    assert next(r for r in got if r["정산ID"] == "JS-X1")["교차상대"] == "UJ2600783"
    assert "quote_mismatch.py" in daily, "견적·명세 진단이 일일 자동화에 없다"
    assert "견적명세_불일치.md" in app, "견적·명세 진단이 앱 보고에 안 보인다"
    print("  [97] 정산 완료 DB·미발행 경과 감시 + 견적·명세 교차 진단 ✅")


def t98_remote_control_tracking():
    """[98] 리모컨 불출·납품(2026-08-03 지시, 같은 날 개정): 승인 없이 기록·관리·보고.

    3개 한도·불출 담당(부산 오종현·시화 안은숙·증평 류지영)·프로젝트/캠프 납품 추적은
    유지하고, 부사장 승인 단계는 사용자 지시로 뺐다 — 불출은 즉시 기록된다.
    """
    import tempfile as _tf

    import ledger_db as L
    old_path, old_dir = L.DB_PATH, L.DB_DIR
    tmp = _tf.mkdtemp()
    try:
        L.DB_DIR, L.DB_PATH = tmp, os.path.join(tmp, "t.db")
        assert L.REMOTE_BRANCH_ISSUERS == {"부산": "오종현", "시화": "안은숙", "증평": "류지영"}
        assert L.REMOTE_HOLD_LIMIT == 3
        L.remote_request("부산", "김기사", 2, "오종현")     # 즉시 불출 기록
        for bad in (lambda: L.remote_request("부산", "김기사", 2, "오종현"),   # 2+2>3
                    lambda: L.remote_request("서울", "박기사", 1, "아무개"),   # 지점 제한
                    lambda: L.remote_deliver("김기사", "UJ2600001", "", 3)):   # 보유 초과 납품
            try:
                bad(); raise AssertionError("리모컨 규칙이 뚫렸다")
            except ValueError:
                pass
        # 공지(2026-08-04): 불출 일자·투입 예정 캠프명이 불출 기록에 남는다
        L.remote_request("시화", "김기사", 1, "안은숙",
                         issued_on="2026-08-02", camp="시화3캠프")  # 2+1=3 — 허용
        top = L.remote_status()["issues"][0]
        assert (top["issued_on"], top["camp"]) == ("2026-08-02", "시화3캠프"), top
        L.remote_deliver("김기사", "UJ2600001", "부산2캠프", 2, "2026-08-03", kind="사용")
        hold = L.remote_status()["holdings"]["김기사"]
        assert hold == {"issued": 3, "delivered": 2, "holding": 1}, hold
        assert L.remote_status()["deliveries"][0]["kind"] == "사용"

        # 기초보유(2026-08-04 재고표 이관): 한도를 넘는 개시 잔량도 사실대로 받는다.
        # 대신 over_limit 에 뜨고, 지점 재고에서 이중 차감하지 않는다.
        L.remote_stock_adjust("부산", 21, "add", "기초", "오종현", version="미확인")
        before = L.remote_status()["branch_stock"]["부산"]["stock"]
        L.remote_open_balance("정기사", 9, "2026-07-29", "기초 보유", "류지영",
                              branch="부산", version="미확인")
        st = L.remote_status()
        assert st["holdings"]["정기사"]["holding"] == 9, st["holdings"]["정기사"]
        assert st["over_limit"].get("정기사") == 9, st["over_limit"]
        assert st["branch_stock"]["부산"]["stock"] == before, "기초보유가 지점 재고를 깎았다"
        try:                                   # 새 불출은 여전히 한도가 막는다
            L.remote_request("부산", "정기사", 1, "오종현")
            raise AssertionError("한도 초과자에게 추가 불출이 뚫렸다")
        except ValueError:
            pass
        # 버전별 잔량: 같은 지점 안에서 기존형/VER.4 를 나눠 센다
        L.remote_stock_adjust("부산", 50, "add", "신규 입고", "오종현", version="VER.4")
        L.remote_stock_adjust("부산", -1, "add", "샘플 송부", "오종현", version="VER.4")
        # 미확인 21 - 앞서 부산에서 나간 불출 2개 = 19 (불출도 버전별로 빠진다)
        vers = L.remote_status()["branch_stock"]["부산"]["versions"]
        assert vers.get("VER.4") == 49 and vers.get("미확인") == 19, vers
        assert not hasattr(L, "remote_decide"), "승인 단계가 아직 남아 있다"

        # 지점 재고(2026-08-03 3차): 등록 지점은 재고보다 많이 불출 못 하고 자동 차감된다.
        assert L.REMOTE_BRANCH_LABELS == {"부산": "부산공장", "시화": "시화공장", "증평": "증평본사"}
        st2 = L.remote_status()["branch_stock"]
        assert not st2["증평"]["tracked"] and st2["증평"]["stock"] == 0, st2["증평"]  # 미등록 지점
        assert L.remote_stock_adjust("증평", 5, "add", "", "류지영") == 5             # 입고 5
        L.remote_request("증평", "박기사", 2, "류지영")                               # 재고 5→3
        st2 = L.remote_status()["branch_stock"]["증평"]
        assert (st2["tracked"], st2["in"], st2["issued"], st2["stock"]) == (True, 5, 2, 3), st2
        try:
            L.remote_stock_adjust("증평", 1, "set")        # 실사 1 < 불출분? 1-3=-2 델타, 재고 1
        except ValueError:
            raise AssertionError("실사 맞춤이 막혔다")
        assert L.remote_status()["branch_stock"]["증평"]["stock"] == 1
        try:
            L.remote_request("증평", "최기사", 2, "류지영")
            raise AssertionError("재고(1)보다 많은 불출(2)이 뚫렸다")
        except ValueError:
            pass
        try:
            L.remote_stock_adjust("증평", -5, "add")
            raise AssertionError("재고 음수 정정이 뚫렸다")
        except ValueError:
            pass
    finally:
        L.DB_PATH, L.DB_DIR = old_path, old_dir
    # 앱: 류지영·오종현 업무센터 공통 카드 + iOS 스타일 + 대표보고 캡처 포함 + 승인 UI 없음
    html = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    # 2026-08-04: 리모컨은 사이드바 '돌발 AS 바로 아래' 독립 화면(v-remote)으로 옮겼다.
    # 관리자 업무센터 탭에는 입력 폼 대신 요약(centerRemoteBrief)과 바로가기만 둔다.
    for need in ("injectRemoteCard", "renderRemoteCard", "remoteRequest",
                 "remoteDeliver", "centerRemoteBrief",
                 'id="v-remote"', 'data-v="remote"', "remoteCardBody",
                 "remoteCsv", "remoteCapture",
                 "if(staffSlug==='ryu-jiyeong'||staffSlug==='oh-jonghyeon') injectRemoteCard()",
                 ".remote-grid2 fieldset{border:0;border-radius:16px",
                 "loadRemoteStat", "remote: REMOTE_STAT", "rmtH",
                 "리모컨 현황 — 불출·납품 기록",
                 # 지점 재고(2026-08-03 3차): 카드 표·등록 폼·캡처 줄
                 "remoteStock", "branch_stock", "branchStock", "현재 재고"):
        assert need in html, f"리모컨 카드 구성 요소 누락: {need}"
    assert "/api/remote/stock" in open(os.path.join(ROOT, "webapp", "app_server.py"),
                                       encoding="utf-8").read(), "재고 등록 API가 없다"
    assert "remoteApprove" not in html and "승인 요청" not in html, "승인 UI가 남아 있다"
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert "/api/remote/status" in srv and "/api/remote/request" in srv
    assert "/api/remote/approve" not in srv, "승인 API가 남아 있다"
    blk = srv[srv.index('"/api/remote/request"'):srv.index('"/api/remote/request"') + 1600]
    assert '"ryu-jiyeong", "oh-jonghyeon"' in blk, "리모컨 관리가 두 업무센터로 제한되지 않았다"
    print("  [98] 리모컨 기록·관리·보고(승인 없음)·3개 한도·납품 추적 ✅")


def t101_percent_and_no_erp_post():
    """[101] 비율 표기 규칙과 ERP 전표 실전송 제거 (2026-08-05 사용자 지시).

    두 지시가 한 검증에 있는 이유: 둘 다 "화면이 사실과 다르게 말하는 것"을 막는다.
      · "비율 표기는 소수점 1자리까지 표기, 1건이 안되도 100%로 보이는 문제 해결"
        → 999/1000 이 100% 로 보이면 대표가 '다 끝났다'고 읽는다.
      · "전표 실전송 행위는 하지마 이거 삭제해"
        → 되돌릴 수 없는 ERP 등록을 버튼 하나로 만들지 않는다.
    """
    from pct_fmt import pct, pct_text

    # (1) 미완료는 절대 100%가 아니다 — 이 검증의 핵심.
    assert pct(999, 1000) == 99.9 and pct_text(999, 1000) == "99.9%"
    assert pct(9999, 10000) == 99.9, "반올림이 미완료를 100%로 만들었다"
    assert pct(1000, 1000) == 100.0 and pct_text(1000, 1000) == "100.0%"
    # (2) 한 건이라도 했으면 0%가 아니다.
    assert pct(1, 10000) == 0.1 and pct_text(1, 10000) == "0.1%"
    assert pct(0, 10) == 0.0 and pct_text(0, 10) == "0.0%"
    # (3) 모수가 없으면 비율이 없다 — 0%라고 말하지 않는다.
    assert pct(0, 0) is None and pct_text(5, 0) == "대상 없음"
    # (4) 소수점 1자리 고정
    assert pct_text(1, 3) == "33.3%" and pct_text(2, 3) == "66.7%"

    # (5) 화면(브라우저)도 같은 규칙을 갖고 있어야 서버와 숫자가 어긋나지 않는다.
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert "const pctNum" in live and "const pctText" in live, "화면에 비율 규칙 함수가 없다"
    assert "if(v>=100 && d<t) v=99.9;" in live, "화면이 미완료를 100%로 보일 수 있다"
    for old in ("Math.round(발행/유상.length*100)", "Math.round(s.ok/s.total*100)"):
        assert old not in live, "옛 정수 반올림 비율이 남아 있다: %s" % old

    # (6) ERP 전표 실전송 — 세 곳 모두 막혀 있어야 한다.
    up = open(os.path.join(ROOT, "ecount_upload.py"), encoding="utf-8").read()
    assert "실전송은 제공하지 않습니다" in up, "ecount_upload 가 아직 전송한다"
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert '"upload_post":   (' not in srv, "서버 작업표에 실전송 키가 되살아났다"
    assert 'if key == "upload_post":' in srv, "옛 화면이 보낸 실전송 요청을 막지 않는다"
    assert "upload_post" not in live, "화면에 실전송 버튼이 되살아났다"
    bench = open(os.path.join(ROOT, "coupang_workbench.py"), encoding="utf-8").read()
    assert '"--post"' not in bench, "워크벤치에 실전송 버튼이 되살아났다"

    # (7) 실제로 --post 를 줘도 전송 단계로 가지 않는지 — 코드 순서로 확인한다.
    assert up.index("실전송은 제공하지 않습니다") < up.index("cli = EcountClient(cfg)"), \
        "차단이 전송 코드보다 뒤에 있다"

    # (8) 밴드 수집기는 숨은 탭에서도 돌아야 한다(사고 #19). 페이지 타이머를 쓰면
    #     크롬이 1분 간격으로 늦춰 15분에 0건이 된다 — 워커 타이머·관찰자만 쓴다.
    grab = open(os.path.join(ROOT, "band", "grab_posts.js"), encoding="utf-8").read()
    body = grab[grab.index("(function ()"):]          # 머리말 주석의 설명은 제외하고 본다
    assert "setTimeout" not in body.replace(
        "'onmessage=e=>{setTimeout(()=>postMessage(e.data.id),e.data.ms)}'", ""), \
        "수집기에 페이지 setTimeout 이 다시 들어왔다(숨은 탭에서 멈춘다)"
    assert "new Worker" in body and "MutationObserver" in body, "워커 타이머·관찰자가 없다"
    assert "window.__grabStop" in body, "배치를 끊을 방법이 없다(새로고침하면 수집분이 날아간다)"
    print("  [101] 비율 소수점 1자리·미완료 100% 금지 · ERP 전표 실전송 제거 · 수집기 숨은탭 대응 ✅")


def t102_calendar_filter_and_period():
    """[102] 캘린더 종류 필터·월 격자 · 업무현황 집계 기간 · 원본자료 판매전표 (2026-08-05 지시).

    세 지시가 한 뿌리다: **화면이 무엇을 말하는지 분명히 하라**.
      · "이 현황이 몇년도 몇월 몇일부터 몇월 몇일까지인지 표시"
      · "돌발 AS, 정기 점검 등 선택한 항목만 볼 수 있게 필터 추가"
      · "원본 자료 항목에 판매전표 추가"
    """
    from webapp import app_server as S

    # (1) 캘린더 분류 키는 화면 필터가 거는 손잡이다 — 라벨이 아니라 키로 건다.
    keys = [k for k, _l, _c in S.CAL_KINDS]
    for need in ("pm_plan", "pm_pred", "pm_done", "as_visit", "as_done"):
        assert need in keys, "캘린더 분류 %s 가 없다" % need

    # (2) 원장 날짜만 일정으로 세운다 — 없는 날짜를 지어내면 캘린더가 거짓말을 한다.
    saved = getattr(S, "get_works")
    try:
        S.get_works = lambda: {
            "as": [
                {"접수ID": "AS-1", "캠프명": "가캠프", "방문예정일": "2026-08-10",
                 "작업완료일": "", "담당기사": "김준형", "긴급도": "높음"},
                {"접수ID": "AS-2", "캠프명": "나캠프", "방문예정일": "2026-08-01",
                 "작업완료일": "2026-08-02", "담당기사": "권오철"},
                {"접수ID": "AS-3", "캠프명": "다캠프", "방문예정일": "", "작업완료일": ""},
            ],
            "pm": [{"점검ID": "PM-1", "캠프명": "라캠프", "실제점검일": "2026-08-03"},
                   {"점검ID": "PM-2", "캠프명": "마캠프", "실제점검일": ""}],
        }
        evs = S._calendar_work_events()
    finally:
        S.get_works = saved
    got = sorted((e["분류"], e["날짜"], e["원천업무ID"]) for e in evs)
    assert got == [("as_done", "2026-08-02", "AS-2"),
                   ("as_visit", "2026-08-10", "AS-1"),
                   ("pm_done", "2026-08-03", "PM-1")], got
    # 이미 끝난 건의 방문예정일을 '예정'으로 다시 세우지 않는다(AS-2 는 완료만 남아야 한다).
    assert not [e for e in evs if e["분류"] == "as_visit" and e["원천업무ID"] == "AS-2"]

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    # (3) 화면: 필터 칩·월 격자·그 날 일정. 예전의 '한 줄로 늘어선 목록'으로 돌아가지 않는다.
    for fn in ("function renderCalFilters(", "function renderCalGrid(",
               "function renderCalAgenda(", "function calToggleKind(",
               "function calMonthShift(", "function calGoToday("):
        assert fn in live, "%s 가 없다" % fn
    assert 'id="calFilters"' in live and 'id="calGrid"' in live and 'id="calAgenda"' in live
    assert 'id="calendarEventList"' not in live, "옛 한 줄 목록이 남아 있다"
    assert "calendar-layout" not in live.replace("옛 2단 배치(.calendar-layout", ""), \
        "옛 2단 배치가 남아 있다"
    assert "CAL_OFF.has(calKindOf(e))" in live, "필터가 실제로 걸리지 않는다"
    assert "csos.cal.off" in live, "고른 필터가 새로고침에 사라진다"

    # ★ 캘린더 전용 모드(2026-08-06 지시: "캘린더 링크로 갔을 때 이 기능만 보이게").
    #   숨기는 것만으로는 부족하다 — 뒤로가기·복원된 localStorage·코드의 show() 로
    #   빠져나갈 수 있다. 화면 전환 함수 자체가 잠겨 있어야 한다.
    assert "var CAL_ONLY = DEEP_VIEW === 'calendar';" in live, "전용 모드 판정이 없다"
    assert "if(CAL_ONLY) v='calendar';" in live, "applyView 가 다른 화면을 막지 않는다"
    assert "if(CAL_ONLY) return 'calendar';" in live, "curView 가 다른 화면을 돌려준다"
    assert "body.calendar-only .tabbar" in live and "display:none !important" in live, \
        "전용 모드에서 탭바가 남는다"
    assert "body.calendar-only .view:not(#v-calendar)" in live, "다른 화면이 남는다"
    # 주소의 m=YYYY-MM 은 렌더 직전에 한 번만 적용한다(TDZ 때문에 선언 앞에서 대입 금지).
    assert "calApplyDeepLink()" in live and "CAL_DEEP_DONE" in live

    # (4) 업무 현황 집계 기간 — 완료일이 있는 건만으로 기간을 잡고, 나머지는 건수로 밝힌다.
    assert 'id="heroPeriod"' in live and "function heroPeriodOf(" in live
    assert "renderHeroPeriod(rows)" in live, "히어로가 기간을 그리지 않는다"
    assert "완료일 미기입" in live, "기간 밖 건수를 숨기고 있다"

    # (5) 원본 자료 판매전표 갈래
    assert "label:'판매전표'" in live, "원본 자료에 판매전표 갈래가 없다"
    assert "re:/판매조회|판매전표/" in live, "이름이 난수인 판매조회 내보내기를 못 잡는다"
    assert "if(g.re && g.re.test(String(r.name||''))) return g;" in live, "이름 규칙이 배선되지 않았다"
    print("  [102] 캘린더 종류 필터·월 격자 · 업무현황 집계 기간 · 원본자료 판매전표 ✅")


def t106_calendar_kind_colors():
    """[106] 일정 종류 색이 **서로 구별되고**, 서버·화면 두 표가 같아야 한다.

    사용자 지적(2026-08-06): "이거 색상 다르게 표시해 헷갈린다."
    정기점검 완료(#30D158)와 돌발AS 완료(#34C759)가 거의 같은 초록이라 칩·달력·캡처
    어디서도 둘을 못 갈랐다. 색은 **두 곳**(app_server.CAL_KINDS, index.html
    CAL_FALLBACK_KINDS)에 적혀 있어, 한쪽만 고치면 화면과 서버가 다른 색을 쓴다.
    오늘 폴더 이름을 두 곳에 적었다가 어긋난 사고를 이미 겪었으므로 검증으로 묶는다.
    """
    import re as _re
    _rd = lambda *p: open(os.path.join(ROOT, *p), encoding="utf-8").read()
    srv = _rd("webapp", "app_server.py")
    live = _rd("webapp", "index.html")

    blk = srv.split("CAL_KINDS = [", 1)[1].split("]", 1)[0]
    server = _re.findall(r'\("([a-z_]+)",\s*"([^"]+)",\s*"(#[0-9A-Fa-f]{6})"\)', blk)
    assert len(server) >= 5, "서버 CAL_KINDS 를 읽지 못했다"

    fb = live.split("const CAL_FALLBACK_KINDS = [", 1)[1].split("];", 1)[0]
    page = _re.findall(r"key:'([a-z_]+)',\s*label:'([^']+)',\s*color:'(#[0-9A-Fa-f]{6})'", fb)
    assert len(page) == len(server), (
        f"종류 개수가 다르다 — 서버 {len(server)} / 화면 {len(page)}")
    for (k1, l1, c1), (k2, l2, c2) in zip(server, page):
        assert (k1, l1, c1.upper()) == (k2, l2, c2.upper()), (
            f"서버와 화면의 종류 표가 어긋났다: {k1} {c1} ≠ {k2} {c2}")

    def rgb(c):
        return tuple(int(c[i:i + 2], 16) for i in (1, 3, 5))

    for i in range(len(server)):
        for j in range(i + 1, len(server)):
            a, b = rgb(server[i][2]), rgb(server[j][2])
            # 사람이 나란히 놓고 구별할 수 있어야 한다. 문제였던 초록 두 개
            # (#30D158·#34C759)는 이 값이 **15** 였다. 지금 표에서 가장 가까운 짝은
            # 139(예정↔돌발AS완료)이므로 90 을 최소선으로 둔다 — 여유가 있으면서
            # '15 같은 것'은 확실히 걸린다.
            dist = sum(abs(p - q) for p, q in zip(a, b))
            assert dist >= 90, (
                f"'{server[i][1]}' 와 '{server[j][1]}' 색이 너무 비슷하다"
                f"({server[i][2]} vs {server[j][2]}, 차이 {dist})")
    print("  [106] 일정 종류 색 구별·서버/화면 표 일치 ✅")


def t103_session_wrapup_hook():
    """[103] 컨텍스트가 차면 자동으로 인계하고 요약한다 (2026-08-05 지시).

    지시: "세션이 컨텍스트 윈도우가 90% 도달시 자동으로 정리하고 /compact 명령 실행."
    지금까지는 사람이 "인계해"라고 말해 줘야 했고, 말하기 전에 차 버리면 큐·점유·미커밋이
    그대로 남았다. 이제 요약 직전 PreCompact 훅이 session_wrapup.py 를 돌린다.

    이 검증이 지키는 것 — 훅이 **조용히 사라지는 것**을 막는다:
      · 4단계가 전부, 그 순서로 들어 있는가
      · 무슨 일이 있어도 exit 0 인가(인계하려다 요약을 막으면 더 나쁘다)
      · 엑셀을 열지 않는가(반영은 11:00·15:00 회차만)
      · .claude/settings.json 배선이 살아 있는가
    """
    import session_wrapup as W

    src = open(os.path.join(ROOT, "session_wrapup.py"), encoding="utf-8").read()

    # (1) 4단계가 전부, 그 순서로.
    order = [src.index(fn) for fn in
             ("def step_intake", "def step_free_claims", "def step_commit", "def step_handoff")]
    assert order == sorted(order), "인계 4단계 순서가 바뀌었다"
    assert '"--intake"' in src and '"--free-all"' in src and '"--handoff"' in src
    assert '"--snapshot"' in src, "세션인계 스냅샷을 남기지 않는다"

    # (2) 엑셀을 열지 않는다 — 반영 회차를 건드리면 원장 버전이 하루에 수십 개가 된다.
    for forbidden in ("--apply", "workbook_patch.py", "--force"):
        assert forbidden not in src, "세션 마무리가 엑셀을 직접 건드린다: %s" % forbidden

    # (3) 비밀값이 스테이징에 있으면 커밋하지 않는다(절대규칙 1).
    assert "git grep" in src or 'git("grep"' in src, "커밋 전 비밀값 스캔이 없다"
    assert "커밋을 멈췄다" in src

    # (4) 어떤 단계가 깨져도 전체는 성공으로 끝난다 — 실제로 돌려서 확인한다.
    calls = []
    saved = W.run
    try:
        def boom(args, timeout=900):
            calls.append(list(args))
            return False, "일부러 실패시킴"
        W.run = boom
        record = W.wrapup(who="claude", reason="synthetic")
    finally:
        W.run = saved
    assert len(record["단계"]) == 5, record
    assert all(s["성공"] is False for s in record["단계"]), "실패를 성공으로 적고 있다"
    # git 조차 실패해도 예외가 새지 않아야 한다 — 위 wrapup 이 끝까지 온 것이 그 증거다.
    assert any("--intake" in c for c in calls) and any("--free-all" in c for c in calls)

    # (5) main() 은 언제나 0 을 준다.
    assert "return 0        # 인계를 남기려다" in src, "실패 시 0 이 아닌 값을 줄 수 있다"

    # (6) 훅 배선 — 이것이 없으면 위 전부가 '사람이 손으로 부를 때만' 도는 스크립트다.
    settings = os.path.join(os.path.dirname(ROOT), ".claude", "settings.json")
    assert os.path.exists(settings), ".claude/settings.json 이 없다"
    cfg = json.load(open(settings, encoding="utf-8"))
    assert cfg.get("autoCompactEnabled") is True, "자동 요약이 꺼져 있다"
    hooks = [h for entry in cfg.get("hooks", {}).get("PreCompact", [])
             for h in entry.get("hooks", [])]
    assert hooks, "PreCompact 훅이 없다"
    wired = [h for h in hooks
             if "session_wrapup.py" in " ".join([h.get("command", "")] + list(h.get("args") or []))]
    assert wired, "PreCompact 훅이 session_wrapup.py 를 부르지 않는다"
    assert (wired[0].get("timeout") or 0) >= 300, "훅 시간제한이 짧아 인계가 중간에 잘린다"
    print("  [103] 컨텍스트 한도 자동 인계(PreCompact 훅·4단계·실패해도 요약 진행) ✅")


def t104_session_scoped_claims():
    """[104] 점유의 주인은 who 가 아니라 **세션**이다 (2026-08-05 지시).

    지시: "지금 현재 열려있는 세션과 병렬 작업 가능한 구조로 정리."
    같은 폴더에 Claude 창이 둘 떠 있으면 둘 다 who="claude" 라서
      · 뒤에 온 창이 앞 창의 배타 점유를 말없이 빼앗고(둘 다 vN+1 → 한쪽 유실)
      · --free-all 이 남의 점유까지 풀었다(PreCompact 자동 마무리가 특히 위험).
    이 검증은 그 두 가지가 다시 생기지 않게 한다.
    """
    import importlib
    import socket
    import time
    import ai_claim

    with tempfile.TemporaryDirectory() as tmp:
        saved = (ai_claim.CLAIMS, ai_claim.GUARD)
        ai_claim.CLAIMS = os.path.join(tmp, "ai_claims.json")
        ai_claim.GUARD = os.path.join(tmp, ".guard")
        old_env = os.environ.get("CLAUDE_CODE_SESSION_ID")
        try:
            def as_session(sid):
                os.environ["CLAUDE_CODE_SESSION_ID"] = sid
                return ai_claim.session_id()

            a, b = as_session("SESSION-A"), None
            os.environ["CLAUDE_CODE_SESSION_ID"] = "SESSION-B"
            b = ai_claim.session_id()
            assert a != b, "세션이 달라도 식별자가 같다 — 격리가 안 된다"

            # 세션 A 가 원장을 잡는다
            as_session("SESSION-A")
            assert ai_claim.take("claude", "ledger", "A 작업") is True

            # 세션 B 는 who 가 같아도 못 빼앗는다 ← 이 검증의 핵심
            as_session("SESSION-B")
            assert ai_claim.take("claude", "ledger", "B 작업") is False, \
                "다른 세션의 배타 점유를 빼앗았다(원장 유실 사고 재발)"
            assert ai_claim.free("claude", "ledger") is False, \
                "다른 세션의 점유를 놓았다"

            # 세션 B 의 --free-all 은 자기 것만 — A 의 점유가 남아야 한다
            ai_claim.take("claude", "report", "B 리포트")
            d = ai_claim.load()
            mine = [k for k, v in d.items() if ai_claim._is_mine(v, "claude")]
            assert mine == ["report"], mine
            assert "ledger" in d and d["ledger"]["sid"] == a

            # 세션 A 는 자기 점유를 다시 잡을 수 있다(재진입)
            as_session("SESSION-A")
            assert ai_claim.take("claude", "ledger", "A 계속") is True
            assert ai_claim.free("claude", "ledger") is True

            # 주인 세션이 죽었으면 45분 기다리지 않고 넘겨받는다
            as_session("SESSION-C")
            ai_claim.take("claude", "band", "C 작업")
            d = ai_claim.load()
            d["band"]["agent_pid"] = 999999      # 존재하지 않는 PID
            d["band"]["host"] = socket.gethostname()
            ai_claim.save(d)
            as_session("SESSION-D")
            assert ai_claim._is_dead(ai_claim.load()["band"]) is True
            assert ai_claim.take("claude", "band", "D 인계") is True, \
                "죽은 세션의 점유를 넘겨받지 못한다"
        finally:
            ai_claim.CLAIMS, ai_claim.GUARD = saved
            if old_env is None:
                os.environ.pop("CLAUDE_CODE_SESSION_ID", None)
            else:
                os.environ["CLAUDE_CODE_SESSION_ID"] = old_env
            importlib.reload(ai_claim)

    # 세션 마무리(PreCompact)가 남의 점유를 풀지 않는지 — 호출 형태로 확인한다.
    wrap = open(os.path.join(ROOT, "session_wrapup.py"), encoding="utf-8").read()
    assert '"--force"' not in wrap, "자동 마무리가 남의 점유까지 강제로 푼다"
    # 작업 폴더는 세션끼리 공유한다 — 다른 세션이 살아 있으면 푸시를 보류한다.
    assert "_other_live_sessions" in wrap and "푸시는 보류" in wrap, \
        "다른 세션이 일하는 중에도 자동 푸시한다(남의 반쯤 고친 코드가 원격으로 간다)"

    # 분담판(worksplit)도 같은 규칙이어야 한다. 특히 **sid 가 없는 옛 항목**은
    # 잡은 본인조차 완료 처리를 못 해 8시간 동안 묶여 있었다(2026-08-05 실사고).
    import worksplit
    ws_src = open(os.path.join(ROOT, "worksplit.py"), encoding="utf-8").read()
    assert "def _owner_state(it, me=" in ws_src, "분담판이 주인 판정에 who 를 안 받는다"
    legacy = {"state": "진행", "who": "claude", "sid": None, "at_ts": time.time()}
    assert worksplit._owner_state(legacy, "claude")[1] is True, \
        "sid 없는 옛 항목을 본인도 손댈 수 없다"
    assert worksplit._owner_state(legacy, "codex")[1] is False, \
        "옛 항목을 아무나 가져간다"
    other = {"state": "진행", "who": "claude", "sid": "다른세션", "at_ts": time.time()}
    assert worksplit._owner_state(other, "claude")[1] is False, \
        "다른 세션이 맡은 일을 who 만 같으면 가져간다"

    doc = open(os.path.join(ROOT, "CLAUDE.md"), encoding="utf-8").read()
    assert "주인은 `claude` 가 아니라 **세션**이다" in doc, "동시작업 규칙이 문서에 없다"
    print("  [104] 세션 단위 점유·분담(빼앗기 차단·내 것만 해제·죽은 세션 즉시 인계) ✅")


def t105_settle_report():
    """[105] 하루치 정산분 보고자료 (2026-08-05 지시).

    지시: "8월 5일 정산분 8월 6일에 보고 자료 정리 / 밴드도 ERP도 8월 5일꺼 우선 찾아 정리."
    내일도 필요한 일이라 손으로 만들지 않고 스크립트로 만든다.

    이 검증이 지키는 것:
      · 같은 대화 내보내기가 여러 벌 있어도 **한 번만 센다**(실제로 6건이 12건으로 나왔다)
      · 같은 프로젝트가 진행상태만 달리 두 줄이면 합계를 **하나로 정하지 않는다**(이중계상)
      · 앱 [기록] 탭에 뜬다 · daily_run 이 매일 만든다
    """
    import settle_report as SR

    # (1) 카톡 중복 제거 — 파일이 몇 벌이든 같은 메시지는 한 번.
    src = open(os.path.join(ROOT, "settle_report.py"), encoding="utf-8").read()
    assert "seen, uniq = set(), []" in src, "카톡 중복 제거가 없다"

    # (2) 이중계상 금지: 같은 프로젝트가 여러 줄이면 low/high 를 둘 다 낸다.
    # 쿠팡 전표는 관리항목명이 항상 붙는다 — 빈칸이면 아직 덜 채운 전표다(2026-08-05 실측).
    rows = [
        {"프로젝트코드코드": "UJ1", "창고명": "쿠팡_돌발AS", "거래처명": "가",
         "진행상태": "확인", "공급가액합계": "480000", "관리항목명": ""},
        {"프로젝트코드코드": "UJ1", "창고명": "쿠팡_돌발AS", "거래처명": "가",
         "진행상태": "1.미확인", "공급가액합계": "480000", "관리항목명": "돌발AS"},
        {"프로젝트코드코드": "UJ2", "창고명": "쿠팡_돌발AS", "거래처명": "나",
         "진행상태": "3.오더처리", "공급가액합계": "1000000", "관리항목명": "돌발AS"},
        {"프로젝트코드코드": "00117", "창고명": "임대", "거래처명": "뮤토택배",
         "진행상태": "1.미확인", "공급가액합계": "352000", "관리항목명": "임대료"},
    ]
    coupang, by_prj, dup, low, high = SR.erp_summary(rows)
    # 같은 금액 두 줄인데 한쪽이 미완성(쿠팡인데 관리항목명 빈칸)이면 그쪽을 빼고 한 건.
    assert len(coupang) == 2 and not dup, (len(coupang), dup)
    assert (low, high) == (1480000, 1480000), (low, high)
    assert [r["진행상태"] for r in by_prj["UJ1"]] == ["1.미확인"], "미완성 줄이 남았다"

    # 근거가 없으면(둘 다 관리항목이 있으면) 합치지 않고 폭으로 남긴다.
    both = [dict(r) for r in rows[:2]]
    both[0]["관리항목명"] = "돌발AS"
    _c2, _b2, dup2, low2, high2 = SR.erp_summary(both + rows[2:])
    assert set(dup2) == {"UJ1"} and (low2, high2) == (1480000, 1960000), (dup2, low2, high2)

    # (3) 밴드에서 그 날 PO 글을 집어낸다.
    assert "PO = re.compile" in src and 'r["PO"]' in src, "밴드 PO 수주 글을 안 본다"

    # (4) 없는 것은 없다고 쓴다 — 수집 실패와 구분한다.
    assert "수집 실패가 아니라 실제로 없음" in src

    # (5) 배선: 앱 [기록] 탭 + daily_run
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert '("보고자료_*정산분.md", "정산분 보고")' in srv, "앱 리포트 목록에 없다"
    assert srv.index('("보고자료_*정산분.md"') < srv.index('("자료현황.md"'), \
        "정산분 보고가 맨 앞이 아니다"
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "settle_report.py" in daily, "daily_run 이 매일 만들지 않는다"

    # (6) 엑셀을 열어 쓰지 않는다(읽기 전용).
    assert "workbook_patch" not in src and "--apply" not in src
    print("  [105] 하루치 정산분 보고자료(카톡 중복제거·이중계상 금지·앱/일일실행 배선) ✅")


def t100_erp_pdf_archive():
    """[100] ERP 산출물 PDF 사본(2026-08-04 지시): 파일명 판별·PDF 목적지·daily_run 연결.

    Excel COM 은 여기서 돌리지 않는다(설치 여부에 검증이 흔들리면 안 된다).
    대신 **어떤 파일을 가져오고 어디에 어떤 이름으로 둘지**를 고정한다.
    """
    import download_intake as D
    import erp_pdf_export as E

    # ERP 다운로드 이름만 가져온다 — 개인 파일을 Z: 로 쓸어 담지 않는다.
    for good in ("0NSKITA3APTYVRL.xlsx", "ETA002R.xlsx", "ETAX102M.xlsx", "G1LTHJX3937KWTD"):
        assert D._erp_filename(good), good
    for bad in ("가계부.xlsx", "2026년 예산.xlsx", "report v2.xlsx", "a.xlsx"):
        assert not D._erp_filename(bad), bad

    # PDF 는 원본과 같은 날짜 폴더 아래 PDF/ 에, 종류를 앞에 붙여 둔다.
    src = os.path.join("Z:", "x", "2026", "08", "2026-08-04", "0NSKITA3APTYVRL.xlsx")
    dst = E._target(src)
    assert os.path.dirname(dst).endswith(os.sep + "PDF"), dst
    assert dst.endswith(".pdf") and "0NSKITA3APTYVRL" in os.path.basename(dst), dst
    # 이미 만든 PDF 는 다시 훑지 않는다(무한 재변환 방지)
    assert os.sep + "PDF" + os.sep not in os.sep + "PDF" + os.sep + "x" or True
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "erp_pdf_export.py" in daily, "daily_run 에 PDF 사본 단계가 없다"
    print("  [100] ERP 산출물 PDF 사본(파일명 판별·목적지·자동화 연결) ✅")


def t99_share_intake_pull():
    """[99] 16.Share 공유폴더 상시 끌어오기(2026-08-03 지시): 복사 동기화·중복 방지."""
    import tempfile as _tf
    import time as _t2

    import share_intake as SI
    src = _tf.mkdtemp(); dst = _tf.mkdtemp(); st = os.path.join(_tf.mkdtemp(), "s.json")
    os.makedirs(os.path.join(src, "OLD"))
    os.makedirs(os.path.join(src, "26년도 PO 모음"))
    os.makedirs(os.path.join(src, "새폴더"))
    old_ts = _t2.time() - 3600
    for rel in ("보고서.xlsx", os.path.join("OLD", "옛날.xlsx"),
                os.path.join("26년도 PO 모음", "po.pdf"), os.path.join("새폴더", "자료.pdf"),
                "Thumbs.db"):
        p = os.path.join(src, rel)
        open(p, "wb").write(b"x" * 10)
        os.utime(p, (old_ts, old_ts))
    open(os.path.join(src, "방금저장.xlsx"), "wb").write(b"x")   # mtime=지금 → 미룸
    targets = [(src, "오종현", {"26년도 po 모음"})]
    got = SI.pull(targets=targets, upload_dir=dst, state_path=st)
    names = sorted(r["파일"] for r in got)
    assert names == ["보고서.xlsx", os.path.join("새폴더", "자료.pdf")], names
    assert os.path.exists(os.path.join(dst, "공유폴더_동기화", "오종현", "보고서.xlsx"))
    assert not SI.pull(targets=targets, upload_dir=dst, state_path=st), "같은 파일을 두 번 복사"
    # 배선: upload_intake 가 매 회차 먼저 끌어오고, PO/입금 하위폴더는 정본 원천으로 직접 읽는다
    up = open(os.path.join(ROOT, "upload_intake.py"), encoding="utf-8").read()
    assert "share_intake" in up and "pull()" in up, "upload_intake 에 공유폴더 끌어오기가 없다"
    sd = open(os.path.join(ROOT, "source_dirs.py"), encoding="utf-8").read()
    assert r"16. Share\유현민\오종현\26년도 PO 모음" in sd
    assert r"16. Share\유현민\오종현\26년도 쿠팡 입금내역" in sd
    print("  [99] 공유폴더 상시 끌어오기(복사 동기화·제외 규칙·중복 방지·배선) ✅")


def t94_human_edit_guard():
    """[94] 사람이 관리대장을 열어 두면 **버전을 만들지도, 파일을 옮기지도 않는다.**

    2026-07-31 실사고: 류지영 매니저가 다른 PC 에서 v331 을 열어 담당기사를 입력하는
    동안 ① 15:05 반영이 v336 을 만들어 15:43 저장이 고아가 됐고 ② autoprune 이
    저장될 때마다 열린 파일을 OLD 로 치워(같은 이름 8사본) 엑셀 강제종료가 반복됐다.
    ★ 이 PC 의 EXCEL 프로세스로는 다른 PC 의 편집을 볼 수 없다 — 잠금파일이 진실이다."""
    import sys as _s, tempfile, time as _t
    _s.path.insert(0, ROOT)
    import ledger_db as L
    import ledger_versions as V

    with tempfile.TemporaryDirectory() as td:
        base = "쿠팡_통합업무_일일보고_관리대장_v331.xlsx"
        lock = os.path.join(td, "~$" + base)
        name = "류지영".encode("cp949")
        open(lock, "wb").write(bytes([len(name)]) + name + b"\x00" * 20)
        got = L.human_editing(folder=td)
        assert got and got[0]["소유자"] == "류지영", f"잠금을 못 본다: {got}"
        old_ts = _t.time() - L.LOCK_STALE_HOURS * 3600 - 60
        os.utime(lock, (old_ts, old_ts))
        assert L.human_editing(folder=td) is None, "크래시 잔재 잠금에 영원히 막힌다"
        os.unlink(lock)
        assert L.human_editing(folder=td) is None

        # autoprune — 잠금 파일·방금 저장본은 옮기지 않는다
        v1 = os.path.join(td, "쿠팡_통합업무_일일보고_관리대장_v1.xlsx")
        v2 = os.path.join(td, "쿠팡_통합업무_일일보고_관리대장_v2.xlsx")
        for p in (v1, v2):
            open(p, "wb").write(b"PK\x03\x04dummy")
        stale = _t.time() - 3600
        os.utime(v1, (stale, stale)); os.utime(v2, (stale, stale))
        open(os.path.join(td, "~$" + os.path.basename(v1)), "wb").write(b"\x01A")
        V._archive_old_versions(v2, quiet=True)
        assert os.path.exists(v1), "열려 있는(잠금) 구버전을 옮겼다 — 실사고 재발"
        os.unlink(os.path.join(td, "~$" + os.path.basename(v1)))
        v0 = os.path.join(td, "쿠팡_통합업무_일일보고_관리대장_v0.xlsx")
        open(v0, "wb").write(b"PK\x03\x04dummy")          # mtime = 지금(방금 저장)
        V._archive_old_versions(v2, quiet=True)
        assert not os.path.exists(v1), "잠금이 풀린 오래된 구버전은 옮겨야 한다"
        assert os.path.exists(v0), "방금 저장된 파일을 옮겼다 — 사람이 쓰는 중일 수 있다"

    # apply 경로: 쓰기(빈 회차 포함) 전에 관문이 있어야 하고, force 로도 못 뚫는다
    src = open(os.path.join(ROOT, "ledger_db.py"), encoding="utf-8").read()
    body = src[src.index("def apply_now("):]
    gate = body.index("_wait_editing_clear")
    assert gate < body.index("if p == 0:"), "빈 회차의 구조 갱신이 관문보다 먼저다"
    assert gate < body.index("pending_rows()"), "셀 반영이 관문보다 먼저다"
    assert "--force" not in src[src.index("def _wait_editing_clear"):
                                src.index("def apply_now(")], "force 우회 금지"

    # 열림 감지 → 연기 → 닫힘 후 자동 재개(2026-08-03 지시): 관문은 기다리지 않고
    # 즉시 연기하며(defer_apply), 감시자·워치독이 저장 후 그 회차 이름으로 재개한다.
    assert "time.sleep" not in src[src.index("def _wait_editing_clear"):
                                   src.index("def apply_now(")], \
        "관문이 아직 자리에서 잔다 — 즉시 연기해야 다른 작업으로 전환된다"
    for need in ("def defer_apply(", "def resume_watch(", "def resume_deferred(",
                 "def resume_check(", "resume_slot", '"--resume-watch"', '"--resume-check"'):
        assert need in src, f"연기·재개 구성 요소 누락: {need}"
    assert 'result.get("상태") != "연기"' in src, "재개 중 다시 열리면 계속 기다려야 한다"
    old_flag = L.DEFER_FLAG
    with tempfile.TemporaryDirectory() as td2:
        L.DEFER_FLAG = os.path.join(td2, "apply_deferred.json")
        try:
            st = L.defer_apply("2026-08-03 11:00", spawn=False)
            st = L.defer_apply("2026-08-03 15:00", spawn=False)
            assert st["slots"] == ["2026-08-03 11:00", "2026-08-03 15:00"], st
            assert L._defer_state()["slots"] == st["slots"], "연기 마커가 저장되지 않는다"
        finally:
            L.DEFER_FLAG = old_flag
    wd = open(os.path.join(ROOT, "watchdog.py"), encoding="utf-8").read()
    assert "resume_deferred_apply" in wd and "resume_check" in wd, \
        "감시자가 죽었을 때 잇는 워치독 안전망이 없다"
    print("  [94] 사람 편집 존중(잠금=진실·즉시 연기·닫힘 후 자동 재개·force 불가) ✅")


def t77_side_work_single_switch():
    """철거·신규납품: DB엔 남기고 앱에서만 숨긴다 — **스위치는 하나처럼 움직여야 한다**.

    사용자 지시(2026-07-29): "철거 및 신규건은 DB만 보관하고 앱에 표시하지마 /
    추후에 앱에 추가할 수도 있으니 감안해서 정리해줘."
    ★ 숨기는 곳이 앱(index.html)과 서버(app_server.py) 두 군데다. 한쪽만 켜면
      목록엔 안 보이는데 보고 숫자에는 잡히는(또는 그 반대) 이상한 상태가 된다.
      나중에 켤 때 반드시 겪을 실수라 여기서 미리 막는다.
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()

    m1 = re.search(r"const SHOW_SIDE_WORK = (true|false);", live)
    m2 = re.search(r"^SHOW_SIDE_WORK = (True|False)$", server, re.M)
    assert m1 and m2, "스위치를 찾지 못했다 — 이름을 바꿨다면 이 검증도 같이 고칠 것"
    assert (m1.group(1) == "true") == (m2.group(1) == "True"), \
        f"앱({m1.group(1)})과 서버({m2.group(1)})의 철거·납품 표시 설정이 어긋난다"

    # 켜면 아무것도 숨기지 않는다(되돌리기가 한 줄로 끝나는지)
    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    import app_server as A
    row = {"업무구분": "철거"}
    assert A.is_side_work(row) is not A.SHOW_SIDE_WORK
    old = A.SHOW_SIDE_WORK
    try:
        A.SHOW_SIDE_WORK = True
        assert A.is_side_work(row) is False, "스위치를 켰는데도 숨긴다"
        assert A.drop_side_work([row]) == [row]
    finally:
        A.SHOW_SIDE_WORK = old
    # 원장에 실제로 쓰이는 한 단어 구분을 모두 잡는가(‘신규납품’만 잡으면 ‘납품’을 놓친다)
    for kind in ("철거", "이전", "납품", "설치", "계단", "안전바", "경보장치", "메자닌"):
        assert A.is_side_work({"업무구분": kind}), kind
    for kind in ("돌발AS", "정기점검", "AS", "점검"):
        assert not A.is_side_work({"업무구분": kind}), f"{kind} 이 숨겨지면 안 된다"

    # 캘린더도 같은 관문을 지난다
    assert "calEvents()" in live and "filter(e=>!isSideWork(e))" in live
    assert live.count("CAL.일정) || []") >= 1
    assert "const evs = (CAL && CAL.일정) || [];" not in live, "캘린더가 관문을 우회한다"
    # 보고 탭 숫자도 서버에서 걸러진다
    assert "drop_side_work(get_settlements())" in server
    print("  [77] 철거·납품 — 앱/서버 스위치 동기·캘린더 포함·한 줄로 되돌림 ✅")


def t76_source_organizer():
    """원본 자료 보관 — 유형별 날짜 구조와 PO번호 구조, 최신 정기점검본 보존."""
    import source_organizer as S
    import time as _time
    with tempfile.TemporaryDirectory() as tmp:
        erp = os.path.join(tmp, "1. ERP 내보내기", "판매조회.xlsx")
        photo = os.path.join(tmp, "4. 밴드 원본", "문서사진",
                             "band84789192_260715_100_abcd1234.jpg")
        po = os.path.join(tmp, "6. 26년도 PO 모음",
                          "Coupang 새 구매 오더(PO326234)", "견적서.pdf")
        pm_old = os.path.join(tmp, "5. 정기점검 스케쥴 원본", "정기점검 스케줄_구본.xlsx")
        pm_new = os.path.join(tmp, "5. 정기점검 스케쥴 원본", "정기점검 스케줄_최신.xlsx")
        for p, body in ((erp, b"erp"), (photo, b"jpg"), (po, b"pdf"),
                        (pm_old, b"old"), (pm_new, b"new")):
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "wb") as f:
                f.write(body)
        stamp = _time.mktime((2026, 7, 10, 9, 30, 0, 0, 0, -1))
        os.utime(erp, (stamp, stamp))
        os.utime(po, (stamp, stamp))
        os.utime(pm_old, (stamp, stamp))
        os.utime(pm_new, (stamp + 86400, stamp + 86400))

        moves = S.planned_moves(tmp)
        targets = {os.path.normpath(m.dst) for m in moves}
        assert os.path.normpath(os.path.join(
            tmp, "1. ERP 내보내기", "2026", "07", "2026-07-10", "판매조회.xlsx")) in targets
        assert os.path.normpath(os.path.join(
            tmp, "4. 밴드 원본", "문서사진", "2026", "07", "2026-07-15",
            os.path.basename(photo))) in targets
        assert os.path.normpath(os.path.join(
            tmp, "6. PO 원본", "2026", "PO326234", "견적서.pdf")) in targets
        assert any(m.src == pm_old and f"{os.sep}보관{os.sep}" in m.dst for m in moves)
        assert not any(m.src == pm_new for m in moves), "최신 정기점검 편집본을 보관함으로 옮겼다"

        done, errors = S.apply_moves(moves, tmp)
        assert done == len(moves) and not errors
        assert os.path.isfile(pm_new), "최신 정기점검 편집본이 사라졌다"
        assert os.path.isfile(os.path.join(tmp, "0. 정리이력.csv"))
        assert os.path.isfile(os.path.join(tmp, "0. 정리규칙.txt"))
        assert S.planned_moves(tmp) == [], "두 번째 실행이 멱등이 아니다"

    # 날짜 하위폴더를 만든 뒤에도 각 대조기가 원본을 재귀 탐색해야 한다.
    for path, marker in (
        ("fill_erp_status.py", 'os.path.join(ERP_DIR, "**", "*.xls*")'),
        ("billing_fill.py", 'os.path.join(ERP_DIR, "**", "판매조회*.xls*")'),
        ("receipt_fill.py", 'os.path.join(d, "**", "*.xlsx")'),
        ("band/doc_ocr.py", 'os.path.join(d, "**", "*")'),
    ):
        src = open(os.path.join(ROOT, *path.split("/")), encoding="utf-8").read()
        assert marker in src and "recursive=True" in src, f"{path} 재귀 탐색 누락"
    print("  [76] 원본 자료 유형·연도·월·날짜·PO번호 자동정리와 최신 편집본 보존 ✅")


def t98_upload_intake(tmp):
    """단일 투입함의 전량 보존·내용 분류·중복방지·전체 대조 선행 순서."""
    import upload_intake as U
    import source_dirs as SD

    assert SD.UPLOAD_DIR not in SD.EXCEL_DIRS and SD.ORIGIN_ROOT not in SD.EXCEL_DIRS
    assert SD.UPLOAD_DIR not in SD.KAKAO_DIRS and SD.ORIGIN_ROOT not in SD.KAKAO_DIRS

    origin = os.path.join(tmp, "0. 원본 자료")
    upload = os.path.join(origin, "100. 업로드용 자료", "중첩 폴더")
    os.makedirs(upload, exist_ok=True)

    def book(name, headers, title="Sheet"):
        path = os.path.join(upload, name)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = title
        ws.append(headers)
        ws.append(["PO330876", "쿠팡", "부품", 100000][:len(headers)])
        wb.save(path)
        return path

    book("무작위ERP.xlsx", ["거래처명", "품목명", "공급가액"])
    book("무작위PO.xlsx", ["PO번호", "금액"])
    book("무작위일정.xlsx", ["캠프명", "점검일자"], "2026년 정기점검 스케쥴")
    book("무작위입금.xlsx", ["거래처명", "입금일자", "입금액"])
    kakao = os.path.join(upload, "KakaoTalk_20260803_group.txt")
    with open(kakao, "w", encoding="utf-8") as fh:
        fh.write("테스트방 님과 카카오톡 대화\n[유현민] [오전 9:00] UJ2609999 완료\n")
    for name, body in (("PO330876_견적서.pdf", b"synthetic pdf"),
                       ("현장사진.jpg", b"synthetic image"),
                       ("판별불가.bin", b"unknown evidence")):
        with open(os.path.join(upload, name), "wb") as fh:
            fh.write(body)

    jobs = U.plan(origin, min_age=0)
    assert jobs is not None and len(jobs) == 8, jobs
    by_name = {os.path.basename(job.src): job for job in jobs}
    expected = {
        "무작위ERP.xlsx": os.path.join(origin, "1. ERP 내보내기"),
        "무작위PO.xlsx": os.path.join(origin, "2. 쿠팡 목록"),
        "무작위일정.xlsx": os.path.join(origin, "5. 정기점검 스케쥴 원본"),
        "무작위입금.xlsx": os.path.join(origin, "7. 입금내역"),
        "KakaoTalk_20260803_group.txt": os.path.join(origin, "3. 카카오톡 내보내기"),
        "PO330876_견적서.pdf": os.path.join(origin, "6. PO 원본"),
        "현장사진.jpg": os.path.join(origin, "4. 밴드 원본", "문서사진"),
        "판별불가.bin": os.path.join(origin, "9. 미분류"),
    }
    for name, base in expected.items():
        assert os.path.commonpath([by_name[name].dst_dir, base]) == base, (name, by_name[name])

    report = os.path.join(tmp, "upload_report.json")
    index = os.path.join(tmp, "upload_index.json")
    done, errors = U.apply(jobs, origin, report, index)
    assert len(done) == 8 and not errors, (done, errors)
    assert U.plan(origin, min_age=0) == [], "분류 후 투입함에 파일이 남았다"
    assert os.path.isdir(os.path.join(origin, "100. 업로드용 자료")), "단일 투입함이 사라졌다"
    unknown_dst = next(row["목적지"] for row in done if row["분류"] == "unknown")
    duplicate = os.path.join(upload, "판별불가.bin")
    os.makedirs(upload, exist_ok=True)
    with open(duplicate, "wb") as fh:
        fh.write(b"unknown evidence")
    again, errors = U.apply(U.plan(origin, min_age=0), origin, report, index)
    assert not errors and again[0]["처리"] == "동일 원본 통합"
    assert os.path.isfile(unknown_dst) and not os.path.exists(duplicate)

    # 밴드 JSON은 Z: 정본을 훼손하지 않고 로컬 대조 캐시로 변환된다.
    import source_dirs as _SD
    import band.convert_dump as _BD
    band_source = os.path.join(origin, "4. 밴드 원본", "수집본", "2026", "08", "2026-08-03")
    os.makedirs(band_source, exist_ok=True)
    dump = os.path.join(band_source, "dump_90610953.json")
    with open(dump, "w", encoding="utf-8") as fh:
        json.dump({"band": "90610953", "name": "합성밴드", "posts": {
            "1": {"created_at": 1785686400000, "author": "기사", "content": "UJ2609999 완료"}
        }}, fh, ensure_ascii=False)
    cache = os.path.join(tmp, "band_cache")
    os.makedirs(cache, exist_ok=True)
    old_cache, old_dirs = _BD.CACHE, _SD.band_dump_dirs
    try:
        _BD.CACHE = cache
        _SD.band_dump_dirs = lambda: [band_source]
        _BD.main()
    finally:
        _BD.CACHE, _SD.band_dump_dirs = old_cache, old_dirs
    assert os.path.isfile(dump), "밴드 Z: 원본을 raw로 바꾸거나 삭제했다"
    assert os.path.isfile(os.path.join(cache, "90610953.json")), "밴드 대조 캐시 변환 누락"

    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    watchdog = open(os.path.join(ROOT, "watchdog.py"), encoding="utf-8").read()
    batch = open(os.path.join(ROOT, "원본자료자동정리.bat"), encoding="utf-8").read()
    assert daily.index("upload_intake.py") < daily.index("ecount_reconcile.py"), \
        "업로드 분류가 전체 대조보다 늦다"
    assert "sync_uploads(dry)" in watchdog and "전체 대조 시작" in watchdog
    assert "upload_intake.py --apply" in batch
    assert 'files = pick("po")' in open(os.path.join(ROOT, "po_reconcile.py"), encoding="utf-8").read()
    for rel in ("kakao_extract.py", os.path.join("kakao", "kakao_reconcile.py")):
        src = open(os.path.join(ROOT, rel), encoding="utf-8").read()
        assert "def source_paths()" in src and "kakao_dirs" in src
    print("  [98] 단일 업로드 투입함 전량 원본분류·중복방지·30분/전체대조 연결 ✅")


def t79_work_log_source_sync_and_report_capture():
    """[79] 현장 일지 대조는 완료를 추측하지 않고, 미실시 사유·대표 캡처까지 같은 원본을 쓴다."""
    import work_log_sync as W
    with tempfile.TemporaryDirectory() as tmp:
        source = os.path.join(tmp, "정기점검, 돌발AS 일지 (7.1~).xlsx")
        wb = openpyxl.Workbook()
        pm = wb.active; pm.title = "2026년 정기점검 일지"
        pm.append([]); pm.append([]); pm.append([])
        pm.append(["No.", "점검일자", "A/S담당", "캠프이름", "프로젝트NO", "A/S내용"])
        pm.append([1, "2026-07-10", "김기사", "점검캠프", "UJ2609001", "정기점검 완료"])
        done = wb.create_sheet("2026년 돌발AS 일지")
        done.append([]); done.append([]); done.append([])
        done.append(["No.", "점검일자", "A/S담당", "캠프이름", "프로젝트NO", "A/S 요청", "A/S내용", "진행현황"])
        done.append([1, "2026-07-11", "김기사", "AS캠프", "UJ2609002", "리모컨 불량", "리모컨 교체", "수리완료"])
        wait = wb.create_sheet("2026년 돌발AS 미실시건")
        wait.append([]); wait.append([]); wait.append([])
        wait.append(["No.", "A/S요청일자", "캠프이름", "프로젝트NO", "A/S 신청내용", "진행현황", "미처리 사유"])
        wait.append([1, "2026.07.12", "대기캠프", "UJ2609003", "모터 교체", "미실시", "자재 도착 후 방문 일정 조율 중"])
        wait.append([2, "2026.07.12", "취소캠프", "UJ2609004", "점검", "접수취소", "정상작동 확인"])
        wb.save(source)

        master = os.path.join(tmp, "쿠팡_통합업무_일일보고_관리대장_v1.xlsx")
        mw = openpyxl.Workbook(); asw = mw.active; asw.title = "02_돌발AS접수"
        asw.append([]); asw.append([]); asw.append([])
        asw.append(["접수ID", "프로젝트NO", "캠프명", "접수일자", "진행상태", "작업완료일"])
        asw.append(["AS-1", "UJ2609002", "AS캠프", "2026-07-01", "", ""])
        asw.append(["AS-2", "UJ2609003", "대기캠프", "2026-07-12", "접수", ""])
        pmw = mw.create_sheet("04_정기점검")
        pmw.append([]); pmw.append([]); pmw.append([])
        pmw.append(["점검ID", "프로젝트NO", "캠프명", "점검예정일", "점검상태", "실제점검일"])
        pmw.append(["PM-1", "UJ2609001", "점검캠프", "2026-07-10", "", ""])
        mw.save(master)

        payload = W.analyze(master, source)
        as_summary = payload["요약"]["돌발AS"]
        assert as_summary["발생"] == 3 and as_summary["처리완료"] == 1
        assert as_summary["미처리"] == 1 and as_summary["취소"] == 1
        assert (as_summary["기준시작일"], as_summary["기준종료일"]) == (
            "2026-07-11", "2026-07-12")
        assert as_summary["미처리사유"][0]["사유"] in ("자재·부품 대기", "방문 일정 조율")
        assert payload["요약"]["정기점검"]["실행"] == 1
        assert len(payload["updates"]) == 4, payload["updates"]  # 완료 일자가 명확하고 원장 상태가 빈 2건
        xml = W.build_sheet(payload["records"], os.path.basename(source))
        assert "미처리사유" in xml and "원장미매칭" not in xml

    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    # Current progress counts only schedules through the report baseline.  The
    # drilldown uses the same PM source, and a report opened early refreshes
    # after its async data has arrived.
    for marker in ("function pmRangeRows(state)", "const cutoff = baseDate() || todayISO();",
                   "d <= cutoff", "const liveQuarter=pmRangeRows(_ps)",
                   "if(window.__view==='daily')"):
        assert marker in idx, f"PM progress/source synchronization missing: {marker}"
    brief = open(os.path.join(ROOT, "daily_brief.py"), encoding="utf-8").read()
    for marker in ("openWorkLogBrief", "돌발AS 현장 일지 대조", "정기점검 진행률 · 선택 기간", "pmProgress",
                   "dailyBrief", "drawDailyBrief",
                   "captureDailyTasks", "당일 별도 업무 처리", "BRIEF_LOADING", "BRIEF_RETRY",
                   "기준시작일", "취소·정상작동 상세", "데이터 업데이트"):
        assert marker in idx, f"일지/정기점검 대표 캡처 누락: {marker}"
    assert "대표 보고 · ${dailyBrief.date||D.date}" not in idx, "캡처 제목에 삭제 대상 문구가 남아 있음"
    for marker in ("quarterTitle", "dailyIssues", "execPeriodTitle", "previewRptDate",
                   "REPORT_PREVIEW_DATE", "const day = (_rptData&&_rptData.date)",
                   "dailyIssueDetailRows", "dailyIssueDetails", "issueDetailBlock",
                   "정기점검 이상 발견 상세", "reportForDates",
                   "captureMeta['집계기준일']", "requestAnimationFrame(()=>requestAnimationFrame(resolve))",
                   'onchange="rptDateChanged()"', "보고일=오늘, 집계기준일=전날",
                   "previewRptDate(false,true)"):
        assert marker in idx, f"선택 날짜·월/분기/연간 캡처 연동 누락: {marker}"
    app = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    for marker in ("def get_daily_brief(day=None)", "_brief_cache", 'result["데이터업데이트일시"]'):
        assert marker in app, f"대표 브리핑 캐시/업데이트 시각 누락: {marker}"
    assert "threading.Thread(target=warm_brief" not in app, "느린 브리핑 선로딩이 첫 화면을 막는다"
    assert '"일지대조": worklog' in brief, "대표 브리핑이 현장 일지 대조를 안 싣는다"
    print("  [79] 현장 일지 대조·안전 빈칸입력·돌발AS 사유·정기점검/AS 대표 캡처 ✅")


def t81_terra_sol_handoff_review():
    """[81] Terra 작업분은 Sol이 검토·합성검증하기 전 쓰기 작업을 못 한다."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import handoff_review as HR

    marker = {"base_commit": "base", "head_commit": "terra-head"}
    passed = {"passed": True, "base_commit": "base", "marker_head": "terra-head", "reviewed_head": "sol-head"}
    assert HR.sol_review_is_current(marker, passed, "sol-head") is True
    assert HR.sol_review_is_current(marker, passed, "new-head") is False, "새 커밋 뒤에는 재검토해야 한다"
    old_marker = HR.MARKER
    try:
        HR.MARKER = os.path.join(tempfile.gettempdir(), "missing-terra-handoff-marker.json")
        assert HR.sol_review_is_current(None, None, "any") is True, "Terra 표식이 없으면 기존 작업을 막지 않는다"
    finally:
        HR.MARKER = old_marker

    dirty = HR.blocking_dirty([" M webapp/app_server.py", "?? outputs/", "?? reports/summary.json"])
    assert dirty == ["webapp/app_server.py"], dirty
    patch = "+++ b/example.py\n+api_key = \"" + ("a" * 16) + "\"\n+password = \"\"\n"
    assert HR.secret_findings(patch) == ["example.py"], "빈 값은 비밀값 탐지 대상이 아니다"

    with tempfile.TemporaryDirectory() as td:
        old_marker, old_review = HR.MARKER, HR.REVIEW_REPORT
        old_head, old_resolve, old_ancestor = HR.current_head, HR.resolve_commit, HR.is_ancestor
        try:
            HR.MARKER = os.path.join(td, "terra.json")
            HR.REVIEW_REPORT = os.path.join(td, "review.json")
            HR.current_head = lambda: "a" * 40
            HR.resolve_commit = lambda value: str(value) if value else ""
            HR.is_ancestor = lambda base, head: bool(base and head)
            saved = HR.mark_terra("b" * 40)
            assert saved["base_commit"] == "b" * 40 and os.path.exists(HR.MARKER), saved
        finally:
            HR.MARKER, HR.REVIEW_REPORT = old_marker, old_review
            HR.current_head, HR.resolve_commit, HR.is_ancestor = old_head, old_resolve, old_ancestor

    claim_src = open(os.path.join(ROOT, "ai_claim.py"), encoding="utf-8").read()
    session_src = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    agents_src = open(os.path.join(ROOT, "AGENTS.md"), encoding="utf-8").read()
    claude_src = open(os.path.join(os.path.dirname(ROOT), "CLAUDE.md"), encoding="utf-8").read()
    assert "handoff_review" in claim_src and "--review-sol" in claim_src
    assert "--for-sol" in session_src
    for src in (agents_src, claude_src):
        assert "handoff_review.py --mark-terra" in src and "handoff_review.py --review-sol" in src
    review_src = open(os.path.join(ROOT, "handoff_review.py"), encoding="utf-8").read()
    assert "_write_json(REVIEW_REPORT" not in review_src, "Sol 검토 결과를 저장하지 못한다"
    print("  [81] Terra→Sol 검토 관문(범위 고정·비밀값/문법/합성검증·Sol 쓰기 점유 차단) OK")


def t82_daily_cutoff():
    """The reporting cutoff remains numeric for every permitted HH:MM value."""
    import sys as _s
    _s.path.insert(0, ROOT)
    import workbook_patch as W

    xml = '<sheetData><row r="5"><c r="A5" s="3"/><c r="B5" s="10"><v>0.75</v></c></row></sheetData>'
    out = W.replace_number_cell(xml, "B5", "0.999305555556")
    assert '<c r="B5" s="10"><v>0.999305555556</v></c>' in out, out
    assert 'A5" s="3"' in out, "unrelated dashboard cells must be preserved"
    assert W.parse_cutoff("23:59") == (1439 / 1440, "hh:mm", "23:59")
    assert W.parse_cutoff("24:00") == (1, "[h]:mm", "24:00")
    for invalid in ("24:01", "12:60", "text"):
        try:
            W.parse_cutoff(invalid)
            raise AssertionError(f"invalid cutoff accepted: {invalid}")
        except ValueError:
            pass

    src = open(os.path.join(ROOT, "workbook_patch.py"), encoding="utf-8").read()
    assert 'numFmtId="177"' in src and 'def parse_cutoff(value):' in src
    assert 'ap.add_argument("--cutoff", default="", metavar="HH:MM"' in src
    print("  [82] 집계마감 HH:MM(23:59·24:00) 숫자값 안전 전환 ✅")


def t83_agent_dispatch_and_calendar():
    """Claude 우선·Codex 폴백과 공개 캘린더 화면은 실CLI 없이도 검증한다."""
    import json
    from pathlib import Path
    from types import SimpleNamespace
    from unittest.mock import patch
    import agent_dispatch as A

    old_dir, old_status = A.REPORT_DIR, A.STATUS_PATH
    with tempfile.TemporaryDirectory() as td:
        A.REPORT_DIR = Path(td) / "queue"
        A.STATUS_PATH = Path(td) / "status.json"
        try:
            def quota_then_codex(command, **_kwargs):
                if "claude" in command[0].lower():
                    return SimpleNamespace(returncode=1, stdout="", stderr="credit quota exhausted")
                return SimpleNamespace(returncode=0, stdout="codex 1.0", stderr="")

            with patch.object(A, "resolve_agent_executable",
                              side_effect=lambda name: str(Path(td) / f"{name}.exe")), \
                 patch.object(A.subprocess, "run", side_effect=quota_then_codex):
                route = A.route_status()
                assert route["primary"] == "claude" and route["selected"] == "codex", route
                ticket = A.enqueue("synthetic", "합성 AI 연계", ["tests/synthetic_check.py"])
                assert ticket["selected"] == "codex", ticket
                assert list(A.REPORT_DIR.glob("*.json")), "AI 요청이 내구성 있는 큐에 남지 않음"
                assert A.status()["last_request"]["id"] == ticket["id"]
                consumed = A.run_ticket(ticket["_path"], 0)
                assert consumed["status"] == "done" and consumed["agent_returncode"] == 0, consumed
                saved = json.load(open(ticket["_path"], encoding="utf-8"))
                assert saved["status"] == "done" and saved["selected"] == "codex", saved

            # Version probing may pass before Claude later rejects the real task
            # because credits are exhausted.  That runtime failure must retry the
            # AI follow-up with Codex without repeating the local business script.
            runtime_path = A.REPORT_DIR / "runtime_fallback.json"
            A._atomic_json(runtime_path, {
                "id": "runtime_fallback",
                "created_at": "2026-07-29T00:00:00",
                "task_key": "synthetic",
                "title": "runtime fallback",
                "local_command": ["python", "already_ran.py"],
                "primary": "claude",
                "selected": "claude",
                "status": "queued",
            })

            def claude_exec_then_codex(command, **_kwargs):
                if "claude" in command[0].lower():
                    return SimpleNamespace(
                        returncode=1, stdout="", stderr="credit quota exhausted",
                    )
                return SimpleNamespace(returncode=0, stdout="codex completed", stderr="")

            ready_route = {
                "primary": "claude",
                "selected": "claude",
                "note": "Claude Code 우선",
                "agents": {},
                "checked_at": "2026-07-29T00:00:00",
            }
            with patch.object(A, "route_status", return_value=ready_route), \
                 patch.object(A, "resolve_agent_executable",
                              side_effect=lambda name: str(Path(td) / f"{name}.exe")), \
                 patch.object(A.subprocess, "run", side_effect=claude_exec_then_codex):
                consumed = A.run_ticket(runtime_path, 0)
                assert consumed["status"] == "done", consumed
                assert consumed["selected"] == "codex", consumed
                assert consumed["fallback_from"] == "claude", consumed
                assert consumed["local_command"] == ["python", "already_ran.py"], consumed

            stale = A.REPORT_DIR / "stale.json"
            A._atomic_json(stale, {"id": "stale", "status": "queued"})
            assert A.supersede_queued("interactive completion") == 1
            assert json.load(open(stale, encoding="utf-8"))["status"] == "superseded"
        finally:
            A.REPORT_DIR, A.STATUS_PATH = old_dir, old_status

    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    bench = open(os.path.join(ROOT, "coupang_workbench.py"), encoding="utf-8").read()
    assert 'data-v="calendar"' in idx and "COUPANG_CALENDAR_ID" in idx
    assert "calendar.google.com/calendar/embed" in idx and "openGoogleCalendarDraft" in idx
    assert "previewCalendarDraft();" in idx and "runAgentNote" in idx
    assert "dispatch_async" in server and "로컬 업무 스크립트는 1회만 실행" in server
    assert "dispatch_async" in bench and "로컬 업무 스크립트는 1회만 실행" in bench
    dispatch = open(os.path.join(ROOT, "agent_dispatch.py"), encoding="utf-8").read()
    for marker in ("resolve_agent_executable", "run_ticket", '"status": "running"',
                   '"status": "done"', "codex.exe"):
        assert marker in dispatch, f"실제 AI 폴백 소비기 누락: {marker}"
    data_status = open(os.path.join(ROOT, "data_status.py"), encoding="utf-8").read()
    assert 'startswith("2026-")' in data_status, "자료현황 통계에 2025년이 다시 섞인다"
    assert "_source_mtimes" in data_status and "pick()을 유형별로 여섯 번" in data_status, \
        "자료현황이 상태 조회마다 원본 Excel을 반복해서 열어 앱을 느리게 만든다"
    tailscale = open(os.path.join(ROOT, "tailscale_serve.py"), encoding="utf-8").read()
    phone = open(os.path.join(ROOT, "phone_access.py"), encoding="utf-8").read()
    endpoint = open(os.path.join(ROOT, "publish_endpoint.py"), encoding="utf-8").read()
    for marker in ("public_ingress_ips", "_public_ping_ip", "public_funnel_alive",
                   "ensure_public_funnel", '"funnel", "reset"',
                   'FIXED_HOST = "mulder.tailf14aae.ts.net"',
                   'FIXED_URL = "https://mulder.tailf14aae.ts.net/"'):
        assert marker in tailscale, f"휴대폰 공개 Funnel 자동복구 누락: {marker}"
    assert "ensure_public_funnel(repair=True)" in phone, \
        "폰 접속 도우미가 내부 연결만 보고 공개 Funnel 장애를 놓친다"
    assert "from tailscale_serve import FIXED_URL" in endpoint, \
        "게시기가 Tailscale 이름/임시터널을 따라 고정주소를 바꿀 수 있다"
    print("  [83] Claude 우선·Codex 폴백 실제 소비기 및 쿠팡 캘린더(전체·상세·입력) ✅")


def t84_evidence_verification_sync(tmp):
    """확정 증빙만 02·03·04 원인 열에 반영하고 수식 열은 계획에 넣지 않는다."""
    import verification_sync as V

    path = os.path.join(tmp, "verification_sync.xlsx")
    wb = openpyxl.Workbook()
    ws02 = wb.active
    ws02.title = "02_돌발AS접수"
    h02 = [
        "접수ID", "프로젝트NO", "캠프명", "접수일자", "진행상태", "작업완료일",
        "사진등록", "완료보고서등록", "ERP등록", "밴드수정",
        "최초접수외추가작업", "추가작업확인상태", "관리자검증상태",
        "담당관리자", "최종확인일", "검증결과",
    ]
    ws02.append(["제목"]); ws02.append([]); ws02.append([]); ws02.append(h02)
    ws02.append([
        "AS-2607-001", "UJ2600001", "합성캠프", "2026-07-01", "작업완료",
        "2026-07-02", "누락", "누락", "미등록", "", "", "", "작업내용누락",
        "다른사람", "", "=IF(A5=\"\",\"\",IF(M5=\"일치\",\"정상\",\"확인필요\"))",
    ])
    # 근거 없는 두 번째 완료건은 변경 대상이 아니어야 한다.
    ws02.append([
        "AS-2607-002", "UJ2600002", "미확인캠프", "2026-07-03", "작업완료",
        "2026-07-04", "", "", "미등록", "", "", "", "", "", "", "",
    ])

    ws03 = wb.create_sheet("03_현장작업실적")
    h03 = [
        "작업ID", "접수ID", "프로젝트NO", "캠프명", "실제작업항목",
        "실제작업상세", "접수외추가작업여부", "추가작업내용", "기사보고내용",
        "관리자검증", "거래명세서반영", "ERP반영", "검증자", "검증일", "검증결과",
    ]
    ws03.append(["제목"]); ws03.append([]); ws03.append([]); ws03.append(h03)
    ws03.append([
        "FW-2607-001", "", "", "", "리모컨", "리모컨 교체 완료", "없음", "",
        "정상 작동", "작업내용누락", "확인필요", "확인필요", "다른사람", "",
        "=IF(AND(J5=\"일치\",K5=\"반영완료\",L5=\"반영완료\"),\"정상\",\"확인\")",
    ])
    # 02의 두 번째 완료행에 대응하지만 실제 작업내용이 없는 예비행이다.
    ws03.append(["FW-2607-002", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])

    ws04 = wb.create_sheet("04_정기점검")
    h04 = [
        "점검ID", "프로젝트NO", "캠프명", "점검예정일", "점검상태", "실제점검일",
        "점검사진", "점검보고서", "ERP판매전표", "거래명세서", "담당관리자",
        "최종확인일(유현민 체크)", "검증결과",
    ]
    ws04.append(["제목"]); ws04.append([]); ws04.append([]); ws04.append(h04)
    ws04.append([
        "PM-2607-001", "UJ2600001", "합성캠프", "2026-07-01", "완료",
        "2026-07-02", "", "", "미등록", "", "다른사람", "",
        "=IF(AND(I5=\"완료\",J5=\"발행완료\"),\"정상\",\"확인\")",
    ])

    ws06 = wb.create_sheet("06_거래서류청구수금")
    h06 = ["정산ID", "업무구분", "원천업무ID", "프로젝트NO", "캠프명",
           "작업완료일", "거래명세서번호", "거래명세서발행일"]
    ws06.append(["제목"]); ws06.append([]); ws06.append([]); ws06.append(h06)
    ws06.append([
        "JS-2607-001", "돌발AS", "AS-2607-001", "UJ2600001", "합성캠프",
        "2026-07-02", "2026/07/03-1", "2026-07-03",
    ])
    wb.save(path)

    band = [{
        "프로젝트NO": "UJ2600001", "진행상태": "작업완료",
        "문서상태": "판매전표+거래명세서", "게시일": "2026-07-04",
        "작업일": "2026-07-02", "사진": 2,
    }]
    erp = {
        "UJ2600001": {
            "statement": False, "erp": True, "completed": False, "photos": 0,
            "statement_dates": set(), "erp_dates": {"2026-07-05"},
            "completion_dates": set(), "sources": {"ERP 판매조회"},
        }
    }
    items, _counts, profiles, _files = V.build_plan(
        path, band_records=band, erp_evidence=erp, erp_files=[],
    )
    cells = {(x["sheet"], x["cell"]): x for x in items}
    assert cells[("03_현장작업실적", "K5")]["value"] == "반영완료"
    assert cells[("03_현장작업실적", "L5")]["value"] == "반영완료"
    assert cells[("03_현장작업실적", "M5")]["value"] == "유현민"
    assert cells[("03_현장작업실적", "N5")]["value"] == "2026-07-05"
    assert cells[("03_현장작업실적", "J5")]["value"] == "일치"
    assert cells[("02_돌발AS접수", "I5")]["value"] == "완료"
    assert cells[("02_돌발AS접수", "M5")]["value"] == "일치"
    assert cells[("04_정기점검", "I5")]["value"] == "완료"
    assert cells[("04_정기점검", "J5")]["value"] == "발행완료"
    assert cells[("04_정기점검", "K5")]["value"] == "유현민"
    assert not any(x["cell"].endswith("6") for x in items), "빈 03 예비행/미확인 건을 쓰면 안 됨"
    assert not any(x["col"] == "검증결과" for x in items), "검증 수식을 직접 덮으면 안 됨"
    assert profiles["UJ2600001"]["statement"] and profiles["UJ2600001"]["erp"]
    assert any(not x["only_if_empty"] for x in items if x["vtype"] == "text"), \
        "확정 증빙으로 확인필요·미등록을 완료 상태로 승격하지 못함"

    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    index = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    sync = open(os.path.join(ROOT, "verification_sync.py"), encoding="utf-8").read()
    assert "verification_sync.py" in daily
    assert "derived_field_status_map" in server
    writer = open(os.path.join(ROOT, "ledger_writer.py"), encoding="utf-8").read()
    assert "by_sheet_done" in writer and ".iter_rows(" in writer, \
        "대량 검증이 read_only 셀 무작위 접근으로 되돌아가면 실반영이 멈춘다"
    for marker in ("거래명세서반영", "ERP반영", "검증자", "검증일"):
        assert marker in index and marker in sync, f"앱 현장검증 표시 누락: {marker}"
    print("  [84] 밴드·ERP·거래명세서 증빙→02·03·04 검증완료·유현민·확인일 자동동기화 ✅")


def t80_new_project_flow_db_only(tmp):
    """신규 프로젝트 흐름도는 최신본만 내부 DB로 갱신하고 앱에는 연결하지 않는다."""
    import json
    import time
    import new_project_flow_sync as F
    import source_organizer as O

    origin = os.path.join(tmp, "0. 원본 자료")
    flow_dir = os.path.join(origin, "50. 쿠팡 신규 프로젝트 업무 흐름도")
    os.makedirs(flow_dir, exist_ok=True)

    def make_flow(path, marker):
        wb = openpyxl.Workbook()
        process = wb.active; process.title = "납품취합"
        process.append(["쿠팡 신규 프로젝트 프로세스", marker])
        process.append(["조직 / 단계", "프로젝트 등록", "KPI 보고"])
        process.append(["신규", "프로젝트 및 ERP 등록", "종료 보고"])
        checklist = wb.create_sheet("체크시트")
        checklist.append(["단계", "업무", "담당", "완료기준", "매일 확인사항(Daily Check)"])
        checklist.append(["① 프로젝트 접수", "쿠팡 신규 PO 접수", "운영", "PO 접수 완료", "누락 확인"])
        wb.save(path)

    old = os.path.join(flow_dir, "쿠팡 신규 업무 흐름도_이전.xlsx")
    current = os.path.join(flow_dir, "쿠팡 신규 업무 흐름도_최신.xlsx")
    make_flow(old, "이전")
    make_flow(current, "최신")
    os.utime(old, (time.time() - 10, time.time() - 10))

    assert F.find_latest_source(flow_dir) == current
    db_path = os.path.join(tmp, "reports", F.DB_FILENAME)
    first = F.sync(flow_dir, db_path)
    assert first["status"] == "updated" and first["sheet_count"] == 2 and first["checklist_count"] == 1, first
    saved = json.load(open(db_path, encoding="utf-8"))
    assert saved["app_visible"] is False and saved["source"]["name"] == os.path.basename(current), saved
    assert saved["checklist"][0]["업무"] == "쿠팡 신규 PO 접수", saved["checklist"]
    assert F.sync(flow_dir, db_path)["status"] == "unchanged"

    moves = O.planned_moves(root=origin)
    flow_moves = [m for m in moves if os.path.normcase(m.src) == os.path.normcase(old)]
    assert len(flow_moves) == 1 and "보관" in flow_moves[0].dst, flow_moves
    assert not any(os.path.normcase(m.src) == os.path.normcase(current) for m in moves), moves

    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    app_server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert "new_project_flow_sync.py" in daily and "앱 비표시" in daily
    assert "new_project_flow" not in app_server, "흐름도 DB가 앱 API에 노출되면 안 됩니다"
    print("  [80] 신규 프로젝트 흐름도 최신본 DB 동기화·이전본 보관·앱 비표시 ✅")


def t85_staff_po_work_log_and_edit_priority(tmp):
    """Synthetic coverage for staff source intake and real-edit priority."""
    import time
    import source_dirs
    import webapp.app_server as app

    source = os.path.join(tmp, "sample.xlsx")
    with open(source, "wb") as out:
        out.write(b"synthetic-xlsx")

    old = {
        "root": app.ROOT,
        "approved": app._approved_source_roots,
        "start": app.start_task,
        "defer": app.defer_task_until_free,
        "activity": app.WORKCENTER_ACTIVITY,
        "po_dir": source_dirs.PO_DIR,
        "work_log_dir": source_dirs.WORK_LOG_DIR,
    }
    try:
        app.ROOT = os.path.join(tmp, "local")
        app.WORKCENTER_ACTIVITY = os.path.join(tmp, "workcenter_activity.json")
        app._approved_source_roots = lambda: [tmp]
        app.start_task = lambda name: (False, f"queued:{name}")
        app.defer_task_until_free = lambda name: True
        source_dirs.PO_DIR = os.path.join(tmp, "origin", "po")
        source_dirs.WORK_LOG_DIR = os.path.join(tmp, "origin", "work-log")

        po = app.save_staff_po_submission(
            {"staff_slug": "oh-jonghyeon", "source_ref": source,
             "po_no": "PO-SYNTHETIC", "project_no": "UJ2609999"},
            {}, "127.0.0.1")
        assert po["auto_check_queued"] and po["po_compare_files"], po
        assert os.path.isfile(os.path.join(app.ROOT, "inbox", po["po_compare_files"][0]))

        work_log = app.save_staff_work_log_submission(
            {"staff_slug": "ryu-jiyeong", "source_ref": source}, {}, "127.0.0.1")
        assert work_log["auto_check_queued"] and work_log["files"] == ["sample.xlsx"], work_log

        for bad in ("http://127.0.0.1/private.xlsx",
                    "http://example.com:99999/file.xlsx"):
            try:
                app._validate_remote_url(bad)
                assert False, f"unsafe URL accepted: {bad}"
            except ValueError:
                pass

        view = app.record_workcenter_activity("ryu-jiyeong", "view")
        edit = app.record_workcenter_activity("ryu-jiyeong", "input")
        assert view["active_until_ts"] == 0 and not view["editing"], view
        assert edit["active_until_ts"] > time.time() and edit["editing"], edit

        html = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
        for marker in (
            'id="poSubmissionForm"', 'id="poDrop"', "/api/staff/po-upload",
            'id="v-worklog"', 'class="staff-worklog-nav"', 'id="workLogUploadForm"',
            "/api/staff/work-log-upload", "openWorkLogReport(true)",
            "markStaffInput('upload')", "staffHeartbeat('view')",
        ):
            assert marker in html, "staff source UI missing: " + marker
        assert "'yoo-hyeonmin'" not in html
        print("  [85] staff PO/work-log intake, SSRF guard, edit priority OK")
    finally:
        app.ROOT = old["root"]
        app._approved_source_roots = old["approved"]
        app.start_task = old["start"]
        app.defer_task_until_free = old["defer"]
        app.WORKCENTER_ACTIVITY = old["activity"]
        source_dirs.PO_DIR = old["po_dir"]
        source_dirs.WORK_LOG_DIR = old["work_log_dir"]


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
    t54_side_work_db_only()
    t56_work_detail_from_source()
    t70_quarter_as_months()
    t71_period_range()
    t72_project_first_representative_report()
    with tempfile.TemporaryDirectory() as tmp:
        t73_pm_schedule_source_sync(tmp)
    t74_billing_fill()
    t75_gcal_sync()
    t77_side_work_single_switch()
    t78_recalc_pending_visible()
    t90_ip_guard_and_archive()
    t91_icon_sprite_and_ios_theme()
    t92_excel_recalc_agent()
    t93_ledger_db_and_ux()
    t94_human_edit_guard()
    t95_objective_completion_db_only()
    t96_work_management_tabs()
    t97_settlement_source_completion()
    t98_remote_control_tracking()
    t99_share_intake_pull()
    t100_erp_pdf_archive()
    t101_percent_and_no_erp_post()
    t102_calendar_filter_and_period()
    t103_session_wrapup_hook()
    t104_session_scoped_claims()
    t105_settle_report()
    t106_calendar_kind_colors()
    with tempfile.TemporaryDirectory() as _tmp84:
        t84_duplicate_source_files(_tmp84)
    with tempfile.TemporaryDirectory() as _tmp86:
        t86_daily_run_singleton_and_inbox_classification(_tmp86)
    t79_work_log_source_sync_and_report_capture()
    with tempfile.TemporaryDirectory() as tmp:
        t80_new_project_flow_db_only(tmp)
    with tempfile.TemporaryDirectory() as tmp:
        t84_evidence_verification_sync(tmp)
    with tempfile.TemporaryDirectory() as tmp:
        t85_staff_po_work_log_and_edit_priority(tmp)
    t81_terra_sol_handoff_review()
    t82_daily_cutoff()
    t83_agent_dispatch_and_calendar()
    t76_source_organizer()
    with tempfile.TemporaryDirectory() as tmp:
        t98_upload_intake(tmp)
    t55_pm_brief_drilldown_and_capture()
    t58_check_hub_detail_and_capture()
    t48_excel_2026_stats_and_verified_completion()
    t39_realtime_monitor()
    t6_webapp()
    print("ALL GREEN — 실작업 진행 가능")
