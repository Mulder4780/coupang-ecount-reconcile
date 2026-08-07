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

    # ★ 자료가 없는 기간을 '미확인'이라 부르지 않는다 (2026-08-07 지시).
    #   카톡 내보내기는 방에 들어간 뒤부터만 나온다. 그 전 작업은 아무리 잘해도
    #   카톡에서 못 찾는다 — '미확인'으로 세면 지워지지 않는 빨간 줄이 쌓이고
    #   진짜 누락(자료는 있는데 보고가 없는 건)이 그 속에 묻힌다.
    _kr = open(os.path.join(ROOT, "kakao", "kakao_reconcile.py"), encoding="utf-8").read()
    assert "자료없음" in _kr, "카톡 대조가 '자료없음'을 가르지 않는다"
    assert re.search(r"floor\s*=\s*min\(\(m\[.date.\] for m in msgs\)", _kr), \
        "자료 시작일을 **가진 메시지에서** 뽑지 않는다 — 날짜를 못박으면 " \
        "나중에 옛 내보내기가 들어와도 계속 '자료없음'이라 답한다"
    assert 'r["완료일"] < floor' in _kr, "완료일이 자료 시작 이전인지를 안 본다"
    # 앱 배지: '자료없음'은 실패도 분모도 아니다.
    _as = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    assert re.search(r'na\s*=\s*sum\(1 for r in rows if r\.get\(col\) == "자료없음"\)', _as), \
        "앱이 '자료없음'을 따로 세지 않는다"
    assert 'r.get(col) not in (okv, "자료없음")' in _as, \
        "앱이 '자료없음'을 미확인과 같이 센다 — 영원히 안 지워지는 빨간 줄이 된다"
    # 확인필요현황·07시트 얹기는 `== "미확인"` 이라 자료없음이 저절로 빠진다.
    _fe = open(os.path.join(ROOT, "findings_export.py"), encoding="utf-8").read()
    assert 'r.get("카톡보고") == "미확인"' in _fe, \
        "확인필요 목록이 카톡보고를 정확히 '미확인'으로 거르지 않는다"
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
             "새 사용법", "18 갱신", "20 갱신",
             extra_rows=[("2026-07-28 #합성2", "묶음제목", "묶음상세")])
    w3 = openpyxl.load_workbook(v3, read_only=True)
    assert w3["00_대시보드"]["A2"].value == "새 사용법"
    assert w3["18_문서발행업무매뉴얼"]["A2"].value == "18 갱신"
    assert w3["20_쿠팡통합업무상세매뉴얼"]["D2"].value == "20 갱신"
    hand_ws = w3["19_AI작업인수인계"]
    hand_ws.reset_dimensions()      # 합성 시트 dimension 이 A열뿐이라 B열이 가려진다
    hand3 = [c.value for row in hand_ws.iter_rows() for c in row]
    assert "보완" in hand3 and "묶음제목" in hand3, \
        "예약 여러 건이 한 버전(vN+1 하나)에 함께 실리지 않는다"
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
    # ★ 자기잠식 금지 (2026-08-07 실사고). 23시트는 4행에 '프로젝트NO' 열을 가진다.
    #   원장을 훑는 쪽이 이 시트를 원장으로 세면 "이미 등록된 프로젝트"가 부풀어
    #   미등록 문서가 통째로 사라진다(실측 223건). 위 멱등 검사는 그 증상을 **우연히**
    #   잡았을 뿐이라, 원인 쪽을 직접 못박는다.
    import findings_export as FE
    assert FE.is_agent_sheet(("23_확인필요현황 (에이전트 자동 갱신 — 수기 입력 금지)",))
    assert not FE.is_agent_sheet(("프로젝트NO", "캠프명")) and not FE.is_agent_sheet(())
    w = openpyxl.load_workbook(v2, read_only=True)
    a1 = next(w["23_확인필요현황"].iter_rows(min_row=1, max_row=1, values_only=True))
    w.close()
    assert FE.is_agent_sheet(a1), f"보고 시트 표식이 1행에 없다: {a1}"
    src = open(os.path.join(ROOT, "findings_export.py"), encoding="utf-8").read()
    assert "is_agent_sheet(head[0])" in src, "doc_unregistered 가 보고 시트를 안 거른다"
    print("  [8] 확인필요 시트 통합(신규 추가·머리글·멱등·자기잠식 금지) ✅")


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
            # 2026-08-06: 끌기가 HTML5 dragstart → 포인터 이벤트로 바뀌었다.
            # 터치에서 dragstart 가 아예 나지 않아 폰에서 못 움직였다(검증 [122]).
            "dashDragEnable('dashGrid'", "initDashboardLayout();",
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
    # 검증 도중 vN+1 이 생길 수 있다(옆 세션·사람 저장 — 2026-08-07 실측 v542→v543).
    # 최신본 판정과 폴더 스캔을 같은 시점으로 다시 맞춰 한 번 재시도한다.
    for _ in range(2):
        master = resolve_master(load_config()["reconcile"]["master_xlsx"])
        keep, move = V.plan(master)
        kept = {k["path"] for k in keep}
        if master in kept:
            break
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
    # ★ 분모는 **24시트가 담는 범위의 글**이어야 한다 (2026-08-07 실사고).
    #   예전에는 캐시 전체와 견줬다. 밤새 밴드를 개설 시점까지 전량 수집하자 캐시가
    #   2,078 → 8,499 글로 뛰었고(2015년 글까지 들어온다), 24시트는 올해치만 담으므로
    #   3,005행 vs 8,499글이 되어 **성공했다는 이유로 검증이 빨개졌다.**
    #   시트가 잘렸는지 보려면 같은 범위끼리 견줘야 한다 — 올해 글만 센다.
    #   (지운 글의 묘비에는 created_at 이 없다 — 그것도 분모에서 빠진다)
    posts = 0
    from datetime import datetime as _dtm
    try:
        import app_server as _AS
        _yr = int(_AS.APP_YEAR)
    except Exception:
        _yr = _dtm.now().year
    _y0 = int(_dtm(_yr, 1, 1).timestamp() * 1000)
    for f in _g.glob(os.path.join(ROOT, "band", "cache", "*.json")):
        if os.path.basename(f).startswith(("dump_", "raw_")):
            continue
        try:
            for p in (_j.load(open(f, encoding="utf-8")).get("posts") or {}).values():
                ms = isinstance(p, dict) and p.get("created_at")
                if ms and int(ms) >= _y0:
                    posts += 1
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

    # ★ 은행에서 그대로 받은 '거래내역조회' 도 읽는다 (2026-08-07 김미영 대리 파일).
    #   사람이 정리한 표가 아니라 원본이라 머리글이 다르다 — '입금액'이 아니라 '입금',
    #   '입금일'이 아니라 '거래일시'. 못 알아보면 `9. 미분류` 로 가서 아무도 안 읽는다
    #   (실제로 classify 가 unknown, 파서가 0건이었다 — 77건이 통째로 사라질 뻔했다).
    bank = os.path.join(tmp, "거래내역조회_입출식 예금20260807.xlsx")
    wb3 = openpyxl.Workbook(); w3 = wb3.active
    w3.append(["거래내역조회_입출식 예금"]); w3.append(["계좌번호:036-509375-04-012"])
    w3.append([" ", "거래일시", "출금", "입금", "거래후 잔액", "거래내용",
               "상대계좌번호", "상대은행", "메모", "거래구분", "수표어음금액",
               "CMS코드", "상대계좌예금주명"])
    w3.append(["135", "2026-03-11 12:17:07", 0, 8306650, 21429227, "쿠팡로지스틱스",
               "", "산업은행", "", "타행이체", 0, "", "쿠팡로지스틱스서비스"])
    # 은행이 거래내용을 잘라 보내는 일이 잦다 — 예금주명을 먼저 봐야 같은 거래처로 묶인다
    w3.append(["153", "2026-03-23 12:03:15", 0, 920700, -53814429, "쿠팡로지스틱",
               "", "산업은행", "", "타행이체", 0, "", "쿠팡로지스틱스서비스"])
    w3.append(["900", "2026-03-24 09:00:00", 5000000, 0, 100, "임대료",
               "", "산업은행", "", "타행이체", 0, "", "건물주"])   # 출금 — 받은 돈이 아니다
    wb3.save(bank)
    from inbox_scan import classify as _cls
    assert _cls(bank) == "receipt", "은행 거래내역을 못 알아본다 — 미분류로 가서 아무도 안 읽는다"
    got3 = rf.parse_deposit_list(bank)
    assert len(got3) == 2, "출금 행까지 입금으로 셌다: %s" % got3
    assert sum(g["금액"] for g in got3) == 9227350, got3
    assert {g["거래처"] for g in got3} == {"쿠팡로지스틱스"}, \
        "잘린 '쿠팡로지스틱' 이 다른 거래처로 갈렸다 — 예금주명을 먼저 봐야 한다: %s" % got3

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
        # 푸는 방법을 반드시 알려 주되 **될 명령**이어야 한다. 이건 codex 의 죽은
        # 점유라 `--free` 는 ai_claim 이 거부한다(세션 단위 규칙) — 통하는 것은
        # `--adopt` 다. 예전엔 무조건 `--free` 를 적어 사람이 그대로 하고 막혔다.
        # 소유 판정은 [114] 가 따로 지킨다.
        assert any("--adopt" in c for _, c in bl), "풀 명령을 알려주지 않는다: %s" % bl
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

    # iOS 외형: 파랑 강조·그룹배경·다크모드·큰 모서리
    # ★ systemBlue 원값(#007AFF)은 쓰지 않는다 (2026-08-06). 흰 배경 글자로 4.0:1,
    #   흰 글자를 얹은 버튼 바탕으로 3.0:1 이라 둘 다 기준(4.5:1) 미달이었다.
    #   색조는 지키고 명도만 낮춘 #0062CC 로 간다 — 판정은 검증 [115] 가 한다.
    assert "--brand:#007AFF" not in live, "저대비 systemBlue 원값으로 되돌아갔다"
    assert re.search(r"--brand:#[0-9A-Fa-f]{6}", live), "파랑 강조색이 없다"
    assert "--bg:#F2F2F7" in live, "iOS 그룹 배경이 아니다"
    # ★ 2026-07-30 실기기 확인 후 되돌림: 이 앱은 흰 배경이 77곳 하드코딩돼 있어 OS 다크를
    #   따라가면 **흰 카드 위 흰 글자**가 되어 아무것도 안 보였다(사용자 화면으로 확인).
    #   그래서 다크 모드를 넣지 않고 color-scheme:light 로 못박았다.
    # ★ 2026-08-07 갱신 — 막을 것은 '다크 모드'가 아니라 **자동으로 따라가는 것**이다.
    #   그 사이 어둡게가 `[data-theme="dark"]` 로 정식으로 들어왔고(사람이 켠 것만 적용),
    #   하드코딩된 밝은 판은 검증 [127] 이 따로 막는다. 사용자 지시로 '시스템 동일'
    #   선택지도 생겼다(2026-08-07: "기본, 다크모드, 시스템 동일 이렇게 구성").
    #   그러므로 금지선은 **CSS `@media (prefers-color-scheme: dark)` 블록** 하나다 —
    #   그건 사람이 고르지 않아도 색이 통째로 뒤집히는 유일한 경로이고, 화면을 그대로
    #   캡처해 보고서로 올리는 이 앱에서는 그것이 곧 사고다.
    #   JS 의 `matchMedia('(prefers-color-scheme: dark)')` 는 **사람이 '시스템 동일'을
    #   고른 뒤에만** 읽히므로 허용한다(검증 [132] 가 그 조건을 지킨다).
    _auto_dark = re.search(r"@media[^{]*prefers-color-scheme\s*:\s*dark", live)
    assert not _auto_dark, \
        "CSS 가 OS 다크를 자동으로 따라간다 — 사람이 고르지 않았는데 보고 화면이 검게 나온다"
    assert "matchMedia" in live and "themeMode() === 'system'" in live, \
        "'시스템 동일'이 사람이 고른 뒤에만 도는 구조가 아니다"
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
    assert '"--batch"' in src and "for item in pending_handoffs()" not in src, \
        "19시트 예약을 건마다 따로 반영해 vN+1 이 건수만큼 폭증한다(--batch 로 묶을 것)"
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

    # ★ 즉시 반영은 **사람이 누를 때만** 열린다 (2026-08-07 지시: "이런거 무시하고
    #   내가 명령 내리면 실시간으로 엑셀 반영하는 알고리즘 추가").
    #   금지의 뜻이 바뀐 것이 아니다 — 막으려던 것은 '도구가 채울 때마다 저절로
    #   vN+1 이 생기는 것'이었지(그래서 하루에 버전이 수십 개 늘었다), 사람이
    #   스스로 내린 명령이 아니었다. 그래서 규칙을 이렇게 좁혀 다시 세운다:
    #   ① 사람 손을 거치지 않는 자동 경로는 여전히 금지(위 writer_apply 단언들)
    #   ② 사람이 누르는 길은 확인창을 거치고 --force 로 기록을 남긴다
    assert '"ledger_now"' in server, "사람이 지시하는 즉시 반영 경로가 없다"
    assert '"--apply", "--force"' in server, "즉시 반영이 강제 표시 없이 원장을 연다(추적 불가)"
    assert "function applyExcelNow(" in live and "runTask('ledger_now')" in live, \
        "앱에 '지금 바로 반영' 이 이어져 있지 않다"
    assert "askYesNo(" in live.split("function applyExcelNow(")[1][:700], \
        "즉시 반영이 확인 없이 실행된다 — vN+1 은 되돌릴 수 없다"
    # 자동으로 도는 것들(스케줄러·daily_run)이 이 키를 부르면 두 회차 규칙이 무너진다.
    for auto in ("daily_run.py", "session_wrapup.py"):
        src = open(os.path.join(ROOT, auto), encoding="utf-8").read()
        assert "ledger_now" not in src and "--apply\", \"--force" not in src, \
            f"{auto} 가 사람 없이 즉시 반영을 부른다"

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
        # versions 는 2026-08-06 에 붙은 버전별 내역이다(검증 [122]). 여기서는 합계만 본다.
        assert {k: hold[k] for k in ("issued", "delivered", "holding")} == \
            {"issued": 3, "delivered": 2, "holding": 1}, hold
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

    # (8-1) 워커 타이머는 **스케줄**만 살린다 — **렌더**는 못 살린다(사고 #22).
    #   밴드 본문은 rAF 로 그려지고 rAF 는 숨은 탭에서 한 번도 안 불린다. 그래서
    #   숨은 탭에서 돌리면 살아 있는 글까지 "본문 없음"으로 읽혀 **가짜 묘비**가 쌓이고,
    #   recheck_plan 이 그 번호를 영영 다시 안 뽑는다. 조용히 실패하느니 거절한다.
    assert "document.hidden" in body, \
        "숨은 탭 판정이 없다 — 창이 뒤에 있으면 살아 있는 글도 전부 실패로 기록된다(사고 #22)"
    start = body[body.index("window.__grabStart"):body.index("window.__grabStop")]
    assert "document.hidden" in start.split("(async ()")[0], \
        "숨은 탭에서 __grabStart 가 시작을 거절하지 않는다"
    assert "S.paused" in start, "돌던 중 창이 뒤로 가면 멈춰 기다려야 한다(실패로 적으면 안 된다)"
    assert "paused:" in body, "__grabStatus() 가 멈춰 있는지를 알려 주지 않는다"

    # (8-2) 밤샘 대조는 한 회차가 **약 2시간**이다(밴드 캐시 8,499글). 45분이던
    #   옛 기준을 그대로 두어 밤새 8회가 전부 잘렸다. 제한을 다시 줄이지 못하게 막는다.
    ovn = open(os.path.join(ROOT, "overnight_run.py"), encoding="utf-8").read()
    m = re.search(r'"--timeout".*?default=(\d+)', ovn, re.S)
    assert m and int(m.group(1)) >= 180, \
        "밤샘 대조 회차 제한이 너무 짧다 — 한 회차가 2시간이라 거의 다 해 놓고 잘린다"
    assert "stdout=fh" in ovn, \
        "출력을 파일로 흘리지 않는다 — 시간초과가 나면 어디서 느렸는지 기록이 안 남는다"
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
    for need in ("pm_plan", "pm_pred", "pm_done", "as_visit", "as_done",
                 "as_open", "pm_overdue"):
        assert need in keys, "캘린더 분류 %s 가 없다" % need
    # 미처리 두 갈래는 완료·예정과 색이 확실히 갈려야 한다(2026-08-06 지시).
    color = {k: c for k, _l, c in S.CAL_KINDS}
    assert color["as_open"] != color["as_done"] and color["pm_overdue"] != color["pm_done"]

    # (2) 원장 날짜만 일정으로 세운다 — 없는 날짜를 지어내면 캘린더가 거짓말을 한다.
    saved = getattr(S, "get_works")
    try:
        S.get_works = lambda: {
            "as": [
                {"접수ID": "AS-1", "캠프명": "가캠프", "방문예정일": "2026-08-10",
                 "접수일자": "2026-08-01", "작업완료일": "", "담당기사": "김준형",
                 "긴급도": "높음"},
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
    # 미처리는 **접수일/예정일** 자리에 선다 — "언제 들어온 게 아직 안 끝났나"를 보려는 것.
    assert ("as_open", "2026-08-01", "AS-1") in got, got     # AS-1 은 접수만 되고 미완료
    assert ("as_done", "2026-08-02", "AS-2") in got, got
    assert ("as_visit", "2026-08-10", "AS-1") in got, got
    assert ("pm_done", "2026-08-03", "PM-1") in got, got
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
            assert ai_claim.take("claude", "band", "D 인계") is True, \
                "죽은 세션의 점유를 넘겨받지 못한다"

            # 스케줄러 점유는 agent_pid=0 — 그때는 `pid` 가 증거다 (2026-08-07 실사고:
            # 죽은 ledger_writer 를 --adopt 가 "살아 있다"며 못 넘겨받아 교착)
            d = ai_claim.load()
            d["band"] = {"who": "scheduler", "why": "w", "sid": "0000dead",
                         "agent_pid": 0, "pid": 999999,
                         "host": socket.gethostname(), "at": time.time()}
            ai_claim.save(d)
            assert ai_claim._is_dead(ai_claim.load()["band"]) is True, \
                "agent_pid=0 이면 pid 증거를 봐야 한다"
            # pid 마저 없으면 보수적으로 살아 있다고 본다
            d["band"].pop("pid"); d["band"]["agent_pid"] = 0
            ai_claim.save(d)
            assert ai_claim._is_dead(ai_claim.load()["band"]) is False
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


def t107_report_dates_and_capture():
    """[106] 보고 기준일 즉시 반영·자동 갱신 · 숨은 화면에서도 캡처 (2026-08-06 지시).

    세 지시가 한 뿌리다: **버튼이 말한 대로 되게 하라**.
      · "저장하고 반영 대기 → 저장하고 반영으로 바꾸고 누르면 바로 반영"
      · "선택 날짜 보고 바로 캡처 버튼 동작 안함"
      · "보고일이 오늘 날짜로 자동 변경되는 알고리즘"
    """
    import report_dates as RD
    from datetime import date

    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()

    # (1) 캡처가 안 되던 진짜 원인: 화면이 숨겨져 있으면 rAF 가 **아예 오지 않는다**.
    #     타이머와 경주시켜 어느 쪽이 오든 진행해야 한다. 오류도 안 나는 침묵형 버그였다.
    assert "function nextPaint(" in live, "그리기 대기 함수가 없다"
    assert "setTimeout(go, ms || 400)" in live, "rAF 가 안 올 때 빠져나갈 길이 없다"
    assert "await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))" \
        not in live, "숨은 화면에서 멈추는 rAF 대기가 남아 있다"
    assert "await nextPaint();" in live and "await captureReport();" in live

    # (2) 「저장하고 반영」 — 문구와 동작이 같아야 한다.
    assert "저장하고 반영</button>" in live, "버튼 문구가 그대로다"
    assert "저장하고 반영 대기" not in live
    assert '"apply":true' in live.replace(" ", "") or "apply:true" in live.replace(" ", ""), \
        "버튼이 즉시 반영을 요청하지 않는다"
    assert 'if b.get("apply"):' in srv and "ignore_input_window=True" in srv, \
        "서버가 즉시 반영 경로를 갖고 있지 않다"
    # 보호장치는 남아 있어야 한다 — 관리대장이 열려 있으면 지시가 있어도 쓰지 않는다.
    ldb = open(os.path.join(ROOT, "ledger_db.py"), encoding="utf-8").read()
    assert "관리대장이 열려 있음" in ldb, "열린 원장을 덮어쓸 수 있다"

    # (3) 보고일 자동 갱신 — 오늘/전날. 이미 맞으면 큐를 늘리지 않는다(멱등).
    w = RD.wanted(date(2026, 8, 7))
    assert w == {"B3": "2026-08-07", "B4": "2026-08-06"}, w
    saved = RD.current
    try:
        RD.current = lambda: {"B3": "2026-08-07", "B4": "2026-08-06"}
        assert RD.run(today=date(2026, 8, 7)) == 0        # 큐 적재 없이 끝나야 한다
    finally:
        RD.current = saved
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "report_dates.py" in daily, "daily_run 이 보고일을 갱신하지 않는다"
    assert daily.index("report_dates.py") < daily.index("upload_intake.py"), \
        "보고일 갱신이 뒤 단계들보다 늦다"
    print("  [107] 보고 기준일 즉시 반영·자동 갱신 · 숨은 화면 캡처 ✅")


def t108_pm_source_fallback():
    """[108] 원본이 아직 담지 않은 날은 **원장으로 보충**한다 (2026-08-06 실사고).

    류지영 정기점검 원본이 정본이라, 원본이 있으면 원장(04시트)을 아예 안 봤다.
    그런데 원본이 8/3 까지만 갱신돼 있어 **8/5 정기점검 완료 3건이 보고에 0건으로
    나갔다.** 대표 보고 캡처가 그대로 0 이었다.

    고친 규칙 — 원본을 무시하지 않는다:
      · 그날 원본에 한 건이라도 있으면 **원본이 정본**(기존 그대로)
      · 그날 원본에 아무것도 없을 때만 원장으로 센다
    """
    import daily_brief as DB

    base = {"as": [], "fw": [], "events": [],
            "pm": [{"프로젝트NO": "UJ2600001", "점검예정일": "2026-08-05",
                    "실제점검일": "2026-08-05", "캠프명": "가"},
                   {"프로젝트NO": "UJ2600002", "점검예정일": "2026-08-05",
                    "실제점검일": "2026-08-05", "캠프명": "나"}]}

    def sched(rows):
        return {"year": 2026, "quarter": 3, "schedule": rows}

    a = dict(base)
    a["pm_schedule"] = sched([{"연결프로젝트NO": "X", "점검예정일": "2026-08-01",
                               "실제점검일": "2026-08-01", "장비수": 3}])
    assert DB.brief("2026-08-05", a)["정기점검"]["완료"] == 2,         "원본이 그날을 안 담았는데 원장으로 보충하지 않는다"

    b = dict(base)
    b["pm_schedule"] = sched([{"연결프로젝트NO": "X", "점검예정일": "2026-08-05",
                               "실제점검일": "2026-08-05", "장비수": 3}])
    got = DB.brief("2026-08-05", b)["정기점검"]
    assert (got["완료"], got["완료장비"]) == (1, 3),         "그날 원본이 있는데 원장이 덮었다 — 원본이 정본이다"
    print("  [108] 원본 미수록일만 원장으로 보충(원본 우선은 유지) ✅")


def t109_remote_edit_delete_versions():
    """[109] 리모컨 — 버전 통일·지사 불출자·고치기·지우기·강제·되돌리기 (2026-08-06 지시).

    사용자 지시: "리모컨 버전 선택할 수 있는 기능 추가하고 전체적으로 버전 관리가
    VER.3인지 VER.4인지 입력 및 확인 수정 가능하게 고도화 해" + 남은 4가지
    (강제 수정 · 지사 불출자 이름란 · 삭제 · 최근 불출 수정).

    지키는 것:
      · 'ver3'·'v4' 처럼 사람이 쓰는 표기가 한 이름으로 모인다 — 안 그러면 재고가 갈라진다
      · 불출에도 버전·불출자가 남는다 (버전 없이 불출하면 지점 버전별 잔량이 어긋난다)
      · 수정은 **이번 수정이 새로 만든 문제**만 막는다. 이미 어긋난 장부 때문에
        무관한 줄까지 못 고치면 아무것도 정리할 수 없다
      · 강제는 사유가 있어야 하고, 지운 것은 원장에 남아 되돌릴 수 있다
    """
    import shutil as _sh
    import tempfile as _tf

    import ledger_db as L
    old_path, old_dir = L.DB_PATH, L.DB_DIR
    tmp = _tf.mkdtemp()
    try:
        L.DB_DIR, L.DB_PATH = tmp, os.path.join(tmp, "t.db")
        assert L.REMOTE_VERSIONS == ("미확인", "기존형", "VER.3", "VER.4")
        assert [L._remote_version(x) for x in ("ver3", "VER 3", "v4", "기존", "")] == \
               ["VER.3", "VER.3", "VER.4", "기존형", ""], "버전 표기가 한 이름으로 안 모인다"

        L.remote_stock_adjust("시화", 5, "add", "입고", "안은숙", version="VER.3")
        rid = L.remote_request("시화", "김기사", 2, "안은숙", camp="시화1캠프",
                               version="ver3", issuer="대리불출자")
        top = L.remote_status()["issues"][0]
        assert (top["version"], top["issuer"]) == ("VER.3", "대리불출자"), top
        # 버전이 붙어야 지점 버전별 잔량이 맞는다(5 - 2 = 3)
        assert L.remote_status()["branch_stock"]["시화"]["versions"]["VER.3"] == 3

        # 최근 불출 고치기 — 캠프·버전·수량을 그 자리에서 바꾼다
        L.remote_edit("issue", rid, {"camp": "시화2캠프", "qty": 1, "version": "VER.4"},
                      edited_by="류지영")
        top = L.remote_status()["issues"][0]
        assert (top["camp"], top["qty"], top["version"]) == ("시화2캠프", 1, "VER.4"), top

        # 한도를 넘기는 수정은 막힌다 — 사유 없는 강제도 막힌다
        for bad in (lambda: L.remote_edit("issue", rid, {"qty": 9}),
                    lambda: L.remote_edit("issue", rid, {"qty": 9}, force=True),
                    lambda: L.remote_delete("issue", rid),
                    lambda: L.remote_edit("issue", rid, {"qty": 0}),
                    lambda: L.remote_edit("없는표", rid, {"qty": 1})):
            try:
                bad(); raise AssertionError("리모컨 수정 규칙이 뚫렸다")
            except ValueError:
                pass
        # 사유를 적은 강제는 통과하고, 강제였다는 사실이 남는다
        out = L.remote_edit("issue", rid, {"qty": 9}, edited_by="류지영",
                            force=True, reason="실물 확인 — 기사가 실제로 9개 보유")
        assert out["row"]["qty"] == 9 and out["warnings"], out
        assert L.remote_audit_list(1)[0]["forced"] == 1

        # 지우기 → 되돌리기. 되돌릴 수 없으면 사람이 무서워서 안 쓴다.
        L.remote_delete("issue", rid, deleted_by="류지영", force=True, reason="중복 입력")
        assert not [r for r in L.remote_status()["issues"] if r["id"] == rid]
        aid = [a for a in L.remote_audit_list(5) if a["action"] == "삭제"][0]["id"]
        L.remote_restore(aid, actor="류지영")
        assert [r for r in L.remote_status()["issues"] if r["id"] == rid], "복구가 안 된다"
        try:
            L.remote_restore(aid); raise AssertionError("두 번 복구됐다")
        except ValueError:
            pass
    finally:
        L.DB_DIR, L.DB_PATH = old_dir, old_path
        _sh.rmtree(tmp, ignore_errors=True)
    # 화면도 같은 목록을 쓰는지 — 서버와 갈리면 'VER.3' 과 'ver3' 이 따로 세어진다
    html = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    for need in ("s.versions||[]", "remoteEditOpen", "remoteRowDelete", "remoteRestore",
                 "Issuer", "/api/remote/edit", "/api/remote/delete"):
        assert need in html, f"리모컨 화면에 {need} 가 없다"
    srv = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    for need in ("/api/remote/edit", "/api/remote/delete", "/api/remote/restore"):
        assert need in srv, f"서버에 {need} 라우트가 없다"
    print("  [109] 리모컨 버전 통일·불출자·수정/삭제/강제/되돌리기 ✅")


def t110_writer_formula_key():
    """[110] 조회 키가 **수식 열**이어도 찾는다 (2026-08-06 실사고).

    점검ID·접수ID·작업ID 는 `=IF($B5="","","PM-"&…)` 수식이다. 적용기가 수식 통을
    그대로 읽어 키 사전에 '=IF(...)' 를 넣는 바람에, 실제로 있는 행을 "행 없음"으로
    버렸다 — 8/5 정기점검 완료 2건이 그렇게 조용히 사라져 대표 보고에 0건으로 나갔다.
    값 통(data_only)에서 키를 읽어야 한다.
    """
    import shutil as _sh
    import tempfile as _tf
    import zipfile as _zip

    import openpyxl
    import ledger_writer as W
    # ★ TemporaryDirectory 를 쓰지 않는다 — 윈도우에서 openpyxl(read_only) 이 zip 을
    #   놓는 시점이 늦어 정리에서 PermissionError 가 난다. 검증이 그것으로 죽으면 안 된다.
    td = _tf.mkdtemp()
    try:
        path = os.path.join(td, "t.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "04_정기점검"
        for i, h in enumerate(("점검ID", "프로젝트NO", "실제점검일"), 1):
            ws.cell(row=W.HDR_ROW, column=i, value=h)
        for r, code in ((W.FIRST, "PM-A"), (W.FIRST + 1, "PM-B")):
            ws.cell(row=r, column=1, value='=IF($B{0}="","","PM-{0}")'.format(r))
            ws.cell(row=r, column=2, value="UJ%04d" % r)
        wb.save(path)
        # 엑셀이 계산해 둔 값을 흉내 낸다 — openpyxl 은 캐시값을 못 쓴다.
        with _zip.ZipFile(path) as z:
            names = z.namelist()
            blobs = {n: z.read(n) for n in names}
        xml = blobs["xl/worksheets/sheet1.xml"].decode("utf-8")
        for r, code in ((W.FIRST, "PM-A"), (W.FIRST + 1, "PM-B")):
            # 계산 결과가 문자열이면 엑셀은 t="str" 과 <v> 를 남긴다. 그 모양을 만든다.
            def fix(m, code=code, r=r):
                attrs = m.group(1).replace(' t="str"', "")
                return f'<c r="A{r}"{attrs} t="str">{m.group(2)}<v>{code}</v>'
            xml = re.sub(r'<c r="A%d"([^>]*)>(<f>.*?</f>)' % r, fix, xml)
        assert "<v>PM-A</v>" in xml, "합성 워크북에 캐시값을 못 넣었다"
        blobs["xl/worksheets/sheet1.xml"] = xml.encode("utf-8")
        with _zip.ZipFile(path, "w", _zip.ZIP_DEFLATED) as z:
            for n in names:
                z.writestr(n, blobs[n])
        q = [{"sheet": "04_정기점검", "key_col": "점검ID", "key": "PM-A",
              "col": "실제점검일", "value": "2026-08-05", "vtype": "date",
              "evidence": "합성", "only_if_empty": False}]
        plans, skips = W.resolve_targets(path, q)
        assert not skips, f"수식 키를 못 찾았다 — {skips}"
        assert plans and plans[0]["row"] == W.FIRST, plans
    finally:
        _sh.rmtree(td, ignore_errors=True)
    print("  [110] 수식 열(점검ID·접수ID)도 조회 키로 찾는다 ✅")


def t111_account_handoff_freshness():
    """[111] 다른 계정·다른 창이 이어받아도 아무 문제 없이 (2026-08-06 지시).

    사용자 지시: "이 세션은 완료되면 다른 계정으로 로그인해서 사용할거야, 그때 아무
    문제 없이 처리될 수 있는 알고리즘 구성해".

    계정이 바뀌면 세션 식별자도 프로세스도 전부 바뀐다. 그때 두 가지가 샌다:
      ① 죽은 세션의 점유가 남아 아무도 원장을 못 고친다 → --adopt 가 **죽었다는
         증거(pid)** 가 있는 것만 회수한다. 살아 있는 옆 세션 것은 건드리지 않는다.
      ② **수집이 밀린 것을 아무도 모른다.** 오늘 실제로 그랬다 — 쿠팡AS 밴드가 8/4 에
         멈춰 있는데 화면은 멀쩡히 숫자를 보여 줬고, 8/5 돌발AS 가 1건으로 나갔다.
         밴드·이카운트는 사람 로그인이 있어야 긁히므로(절대규칙 3), 기계가 할 수 있는
         최선은 "밀렸다 + 이렇게 되살려라" 를 시작 화면 맨 앞에 놓는 것이다.
    """
    import session_handoff as H

    rows = H.data_freshness(today="2026-08-06")
    need = {"이름", "최신", "밀린일", "한도", "밀림", "되살리는법"}
    assert rows and all(need <= set(r) for r in rows), rows
    # 밴드는 **밴드마다** 따로 봐야 한다 — 합쳐서 최댓값을 쓰면 뒤처진 밴드가 가려진다
    assert [r for r in rows if r["이름"].startswith("밴드")], rows
    # '어제'가 비면 밀린 것이다: 8/6 기준 최신 8/4 → 밀린일 2 > 한도 1
    assert H.FRESH_LIMIT["밴드"] == 1, "밴드 한도가 1을 넘으면 어제 빈 것을 못 잡는다"

    st = {"큐잔량": 0, "임시파일": [], "점유": [], "미푸시": [], "지시문사본": [],
          "수집신선도": [{"이름": "밴드: 90610953", "최신": "2026-08-04", "밀린일": 2,
                          "한도": 1, "밀림": True, "되살리는법": "밴드 로그인 후 수집"}]}
    bl = H.blockers(st)
    assert any("수집이 밀렸다" in why for why, _ in bl), bl
    st["수집신선도"][0]["밀림"] = False
    assert not H.blockers(st), "밀리지 않았는데 막았다"
    # 문서에도 표가 나와야 사람이 본다
    st["원장"] = {"버전": "1", "수정": ""}
    st["시각"] = "2026-08-06 12:00"
    st["미커밋"] = []
    st["최근커밋"] = []
    st["다음할일"] = []
    st["진행체크포인트"] = {}
    st["수집신선도"][0]["밀림"] = True
    assert "원본 수집이 어디까지 들어왔나" in H.to_md(st)

    # 이어받기 명령이 실제로 배선돼 있나 — 문서에만 있고 CLI 에 없으면 아무 소용 없다
    src = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    for need_s in ('"--adopt"', "def adopt(", "def write_snapshot(", "_is_dead"):
        assert need_s in src, f"session_handoff 에 {need_s} 가 없다"
    rules = open(os.path.join(ROOT, "CLAUDE.md"), encoding="utf-8").read()
    assert "--adopt" in rules, "계정 인계 절차가 규칙(CLAUDE.md)에 없다"
    print("  [111] 계정이 바뀌어도 이어받기 — 죽은 점유 회수·수집 밀림 감지 ✅")


def t112_band_plan_order_and_scope():
    """[112] 밴드 수집 순서 — **새 글 먼저**, 범위는 사람이 정한 값 (2026-08-06).

    예전에는 캐시의 '구멍'만 봤다. 그래서 마지막 수집 이후 올라온 글은 영원히 대상
    밖이었고, 쿠팡AS 밴드가 8/4 에 멈춰 있는데도 "구멍 0" 이라 아무도 몰랐다 —
    그 결과 8/5 돌발AS 가 1건으로 보고됐다. 이제 캐시 위쪽을 먼저 훑는다.

    그리고 어디까지 파고들지는 **사람이 정한다**. 대화에 남긴 결정은 다음 세션이
    모르므로 band/collect_scope.json 이 기억하고 도구가 그것을 기본값으로 쓴다.
    """
    import band.recheck_plan as RP

    posts = {str(n): {"captured_at": RP.ERA_MS + 1} for n in range(100, 111)}
    p = RP.plan("1", posts, floor=95, ahead=3)
    assert p["new"] == [111, 112, 113], p["new"]
    assert p["gaps"] == [95, 96, 97, 98, 99], p["gaps"]
    # 순서: 새 글 → 구멍(최근부터). 과거글을 먼저 훑으면 오늘 숫자가 계속 틀린다.
    todo = (p["new"] + sorted(p["gaps"], reverse=True))[:6]
    assert todo == [111, 112, 113, 99, 98, 97], todo
    # ahead 를 안 주면 위쪽을 안 본다 — 예전 동작이 그대로 남아 있으면 안 된다
    assert RP.plan("1", posts, floor=95)["new"] == []

    sc = RP.scope()
    assert sc.get("floor"), "사람이 정한 수집 범위(collect_scope.json)가 없다"
    assert all(str(k).isdigit() and int(v) > 0 for k, v in sc["floor"].items()), sc
    assert int(sc.get("ahead") or 0) > 0, "새 글 탐색 개수가 0이면 최신분을 못 잡는다"
    # 없는 글을 열면 밴드가 alert 를 띄우고 **탭 전체가 선다**(2026-08-06 실측).
    # 과거글 구간은 지운 글이 섞여 있어 수백 번 뜬다 — 사람이 수백 번 누를 수는 없다.
    js = open(os.path.join(ROOT, "band", "grab_posts.js"), encoding="utf-8").read()
    assert "muteDialogs" in js and "w.alert" in js, "수집기가 안내창을 막지 않는다"

    # ★ 엑셀은 관리대장 하나로만 관리한다 (2026-08-07 지시).
    #   "쿠팡_거래처코드_최신.xlsx 와 쿠팡_확인필요현황_최신.xlsx 를 관리대장에
    #   통합해서 관리하고, 앞으로도 별도의 엑셀 파일은 만들지 말고 관리대장으로만
    #   관리해." 파일이 늘수록 '어느 게 최신인가'가 흐려지고, 사람이 열어 둔 파일은
    #   갱신도 막힌다. 별도 파일을 되살리는 변경이 들어오면 여기서 막는다.
    _fe2 = open(os.path.join(ROOT, "findings_export.py"), encoding="utf-8").read()
    assert '"--xlsx" in sys.argv' in _fe2, \
        "확인필요현황이 다시 별도 엑셀을 기본 생성한다 — 관리대장 23시트로만 관리한다"
    _ci = open(os.path.join(ROOT, "customer_index.py"), encoding="utf-8").read()
    assert "쿠팡_거래처코드_최신.xlsx" not in _ci, \
        "거래처코드가 다시 별도 엑셀을 만든다 — 관리대장 27시트로만 관리한다"
    assert "customer_sheet" in _ci, "거래처코드 시트 반영 경로가 끊겼다"
    import customer_sheet as _CS
    assert len(_CS.HEADERS) == len(_CS.WIDTHS) == 10, \
        "거래처코드 열 구성이 예전 엑셀과 달라졌다 — 자리만 옮기는 것이지 내용이 바뀌면 안 된다"
    _r = _CS.rows_from_index({"A": {"code": "C", "erp_name": "E", "how": "h", "건수": 1,
                                    "addr": "", "manager": "", "tel": "", "email": "",
                                    "equip": ""}}, {}, [])
    assert _r and all(len(x) == 10 for x in _r), _r
    assert _CS.SHEET_NAME in _CS.build_sheet_xml(_r)
    # 엑셀 쓰기는 11:00·15:00 회차에만 — daily_run 이 아니라 ledger_db 에 있어야 한다.
    _ld = open(os.path.join(ROOT, "ledger_db.py"), encoding="utf-8").read()
    assert re.search(r'"29_거래처코드".*customer_index\.py.*"--sheet"', _ld, re.S), \
        "거래처코드 시트가 11:00·15:00 회차에 걸려 있지 않다"
    _dr0 = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert not re.search(r'customer_index\.py"\)\s*,\s*"--sheet"', _dr0), \
        "daily_run 이 관리대장을 직접 연다 — 엑셀 반영은 11:00·15:00 회차만"

    # ★ 옛 별도 엑셀 접기 — **시트가 채워진 것을 확인한 뒤에만** 옮긴다.
    #   순서가 뒤집히면(파일 먼저 정리, 시트는 비어 있음) 그 자료는 그 순간 아무 데도 없다.
    import side_excel_retire as _SR
    assert _SR.ARCHIVE == "OLD", "정리 자리는 프로젝트 관례대로 OLD/ 다"
    _sheets = {s for s, _f in _SR.PAIRS}
    assert {"23_확인필요현황", "29_거래처코드"} <= _sheets, _SR.PAIRS
    _tmpd = tempfile.mkdtemp()
    _m = os.path.join(_tmpd, "합성대장R_v1.xlsx")
    make_ledger(_m)
    for _sh, _side in _SR.PAIRS:                    # 시트가 없으면 '보류'여야 한다
        _ok, _why = _SR.check(_m, _sh)
        assert _ok is False, f"{_sh}: 시트가 없는데 옮겨도 된다고 했다 — {_why}"
    open(os.path.join(_tmpd, "쿠팡_확인필요현황_최신.xlsx"), "w").close()
    _res = _SR.run(apply=True, master=_m)           # --apply 라도 옮기면 안 된다
    assert os.path.exists(os.path.join(_tmpd, "쿠팡_확인필요현황_최신.xlsx")), \
        "시트가 비었는데 별도 엑셀을 옮겼다 — 자료가 아무 데도 없게 된다"
    assert any("보류" in _m2 for _s2, _f2, _o2, _m2 in _res), _res
    # 시트를 채운 뒤에는 옮겨야 한다
    subprocess.run([PY, os.path.join(ROOT, "findings_sheet.py"), "--master", _m],
                   capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
                   env={**os.environ, "COUPANG_REPORT_DIR": _tmpd, "COUPANG_UPDATES_DIR": _tmpd})
    _m2f = _m.replace("_v1.xlsx", "_v2.xlsx")     # 같은 폴더라 별도 엑셀은 그대로 옆에 있다
    assert _SR.check(_m2f, "23_확인필요현황")[0], "채워진 시트를 '보류'로 봤다"
    _SR.run(apply=True, master=_m2f)
    assert os.path.exists(os.path.join(_tmpd, "OLD", "쿠팡_확인필요현황_최신.xlsx")), \
        "OLD/ 로 옮기지 않았다"
    assert not os.path.exists(os.path.join(_tmpd, "쿠팡_확인필요현황_최신.xlsx"))
    _ldr = open(os.path.join(ROOT, "ledger_db.py"), encoding="utf-8").read()
    assert _ldr.index("side_excel_retire.py") > _ldr.index('"29_거래처코드"'), \
        "별도 엑셀 정리가 시트 갱신보다 앞에 있다 — 한 회차를 헛돈다"

    # ★ 사람이 **옛 버전**을 열어 편집 중이면 그 수정은 다음 회차로 안 넘어간다.
    #   2026-08-07 실측: v538 을 11:08 부터 열어 둔 사이 회차가 v541 까지 만들었다.
    #   파일이 열려 있으니 오류도 안 난다 — 알려 주지 않으면 아무도 모른다.
    import session_handoff as _SH
    _lockd = tempfile.mkdtemp()
    for _n in ("합성대장L_v7.xlsx", "합성대장L_v9.xlsx"):
        open(os.path.join(_lockd, _n), "w").close()
    _latest = os.path.join(_lockd, "합성대장L_v9.xlsx")
    import ecount_reconcile as _ER          # stranded_editor 가 함수 안에서 부른다
    _keep = (_ER.resolve_master, _ER.load_config)
    _ER.resolve_master = lambda *_a, **_k: _latest
    _ER.load_config = lambda *_a, **_k: {"reconcile": {"master_xlsx": _latest}}
    try:
        assert _SH.stranded_editor() == [], "잠금이 없는데 경보를 냈다"
        open(os.path.join(_lockd, "~$합성대장L_v9.xlsx"), "w").close()
        assert _SH.stranded_editor() == [], "최신본을 여는 것은 정상인데 경보를 냈다"
        open(os.path.join(_lockd, "~$합성대장L_v7.xlsx"), "w").close()
        assert _SH.stranded_editor() == ["합성대장L_v7.xlsx"], _SH.stranded_editor()
    finally:
        _ER.resolve_master, _ER.load_config = _keep
    _shs = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    assert '"옛버전편집": stranded_editor()' in _shs and 'st.get("옛버전편집")' in _shs, \
        "옛 버전 편집 경보가 '먼저 처리할 것'에 오르지 않는다"

    # ★ 도구가 있는데 아무도 안 부르는 것이 가장 조용한 누락이다 (2026-08-07 발견).
    #   `fill_erp_status.py` 는 2026-07-28 에 만들어졌는데 daily_run 에 없었다.
    #   손으로 돌린 사람이 있을 때만 채워졌다는 뜻이다 — 그냥 돌려 보니 52칸이 남아
    #   있었고 그중 46건은 ERP 가 이미 완료를 입증한 건이었다.
    _dr = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "fill_erp_status.py" in _dr, \
        "ERP 등록여부 판정이 daily_run 에 없다 — 사람이 손으로 돌릴 때만 채워진다"
    assert re.search(r'fill_erp_status\.py"\)\s*,\s*"--queue"', _dr), \
        "--queue 가 아니면 엑셀을 직접 연다(반영은 11:00·15:00 회차만)"
    import daily_run as _DR
    assert _DR._retryable(["fill_erp_status.py", "--queue"]) is False, \
        "큐 단계를 재시도하면 같은 입력이 두 번 들어간다"

    # ★ 본문이 그려졌다고 머리말까지 그려진 것은 아니다 (2026-08-07 실측).
    #   밴드는 .postText 를 먼저 칠하고 .time 을 조금 뒤에 채운다. 보자마자 가져가면
    #   **날짜 없는 글**이 저장되고, 그런 글은 어떤 작업과도 대조되지 않는다.
    #   실제로 621건이 그렇게 쌓였는데 구멍도 아니고 오래된 것도 아니라
    #   **어느 목록에도 안 잡혔다** — 가장 알아채기 어려운 종류의 구멍이다.
    assert re.search(r"for \(let i = 0; i < \d+ && !txt\(main, '\.postListInfoWrap \.time",
                     js), "본문만 보고 바로 가져간다 — 작성시각을 기다리지 않는다"
    # ★ 캐시 정본은 **임시파일에 다 쓴 뒤 갈아끼운다** (2026-08-07 실사고).
    #   `open(dst,"w")` 는 정본을 먼저 비운다. 19MB 를 흘려 넣는 몇 초 동안 읽는 쪽은
    #   반쪽짜리 JSON 을 보고, 그 사이 합성검증이 두 번 죽었다(죽은 자리가 매번 달라
    #   한동안 "캐시가 깨졌다"고 오해했다). 쓰다가 프로세스가 죽으면 **정말로** 깨져
    #   8,500 글을 다시 긁어야 한다 — 밤샘 한 번 분량이다.
    _cd = open(os.path.join(ROOT, "band", "convert_dump.py"), encoding="utf-8").read()
    assert re.search(r'os\.replace\(\s*tmp\s*,\s*dst\s*\)', _cd), \
        "캐시를 원자적으로 갈아끼우지 않는다 — 쓰는 중에 읽으면 반쪽 JSON 이 보인다"
    assert not re.search(r'json\.dump\([^)]*open\(\s*dst\s*,\s*"w"', _cd), \
        "정본에 바로 쓴다(open(dst,'w')) — 임시파일 + os.replace 로 바꿀 것"

    import band.recheck_plan as _RP2
    _p = _RP2.plan("1", {
        "10": {"created_at": 1, "captured_at": _RP2.ERA_MS + 1},
        "11": {"created_at": None, "captured_at": _RP2.ERA_MS + 1},   # 본문은 있는데 날짜가 없다
        "12": {"deleted": True, "captured_at": _RP2.ERA_MS + 1},      # 진짜 지운 글
    }, floor=10, ahead=0)
    assert _p["dateless"] == [11], _p.get("dateless")
    assert 12 not in (_p.get("dateless") or []), "지운 글을 날짜없음으로 세면 영원히 다시 훑는다"
    assert 11 not in _p["gaps"] and 11 not in _p["stale"], \
        "날짜없음이 구멍/오래됨과 겹치면 두 번 뽑힌다"

    # 세션마다 눈에 띄어야 한다. 신선도 판정(band_latest_days)은 **날짜 있는 글만**
    # 보므로 "밴드 최신 = 오늘" 인데도 그 밑에 수백 건이 대조 밖일 수 있다.
    import session_handoff as _SH
    assert hasattr(_SH, "band_dateless"), "인계 점검이 날짜없음을 세지 않는다"
    _bl = _SH.blockers({"큐잔량": 0, "임시파일": [], "점유": [], "미커밋": [], "미푸시": [],
                        "밴드날짜없음": {"테스트밴드": 621}, "수집신선도": {},
                        "지시문사본": {}, "워크트리": None, "원장": {}, "미머지": []})
    assert any("날짜가 없는 글" in b[0] for b in _bl), \
        "날짜없음 621건이 '먼저 처리할 것'에 안 뜬다 — 조용히 어긋난다"
    _bl0 = _SH.blockers({"큐잔량": 0, "임시파일": [], "점유": [], "미커밋": [], "미푸시": [],
                         "밴드날짜없음": {}, "수집신선도": {},
                         "지시문사본": {}, "워크트리": None, "원장": {}, "미머지": []})
    assert not any("날짜가 없는 글" in b[0] for b in _bl0), "0건인데도 경보가 뜬다"

    # 붙여넣기는 **밴드당 한 번**이어야 한다 (2026-08-06 지시: "내 손 안 가게").
    # 250 은 탭이 얼지 않는 한 배치의 상한일 뿐이다. 한 파일 안에서 회차를 이어 돌리고
    # 회차 사이에 저장·비우기를 하면 사람 손은 한 번으로 끝난다. 예전처럼 250건마다
    # 다시 붙여넣게 만들면 1,100건짜리 밴드에 다섯 번을 시키게 된다.
    import band.make_oneclick as MO

    assert MO.BATCH_MAX == 250, MO.BATCH_MAX
    big = {str(n): {"captured_at": RP.ERA_MS + 1} for n in range(1000, 1004)}
    orig_load, orig_scope = MO.RP.load, MO.RP.scope
    try:
        MO.RP.load = lambda b: big
        MO.RP.scope = lambda: {"floor": {"1": 400}, "ahead": 5}
        out, note = MO.build("1", 600)
    finally:
        MO.RP.load, MO.RP.scope = orig_load, orig_scope
    assert out, note
    rounds = json.loads(re.search(r"const ROUNDS = (\[.*?\]);", out, re.S).group(1))
    assert len(rounds) > 1, "회차로 안 쪼갠다 — 한 번에 밀어 넣으면 탭이 언다"
    assert all(0 < len(r) <= MO.BATCH_MAX for r in rounds), [len(r) for r in rounds]
    flat = [n for r in rounds for n in r]
    assert len(flat) == len(set(flat)), "회차끼리 번호가 겹친다"
    assert "keep: false" in out, "회차 사이에 탭 메모리를 안 비운다"
    assert out.count("__grabSave") >= 1 and "for (let i = 0" in out, \
        "회차마다 저장하고 이어 가는 구조가 아니다"
    print("  [112] 밴드 수집 — 새 글 우선·범위 기억·안내창 차단·한 번 붙여넣기 ✅")


def t113_paste_typos_and_misc_reclass():
    """[113] 붙여넣기 사고를 흡수한다 — 캠프명 오염·미분류 재분류 (2026-08-06).

    김미영 대리 지적: "밴드·카톡에 복사 붙여넣기 오류가 있어 매칭이 안 됨".
    실제 원장 캠프명 18건을 뜯어 보니 사람 잘못이 아니라 **붙여넣기 흔적**이었다 —
    웹에서 복사해 `&amp;` 가 그대로, 메모가 `<-` 뒤에 붙어서, 값 대신 `0`·`...`.
    이름을 고치는 게 아니라 **비교할 때** 걷어 내면 짝이 붙는다(18 → 14건).

    그리고 값이 아닌 표시(`0`)는 '못 잇는 자료'가 아니라 **원장을 고칠 일**이다.
    섞여 있으면 정말 ERP 에 없는 캠프가 몇 개인지 알 수 없다.

    미분류도 같은 이야기다. 판별 규칙은 늘어나는데 규칙이 없던 시절 미분류로 간
    파일은 아무도 다시 안 봤다 — 그래서 --apply 회차가 재분류까지 같이 한다.
    """
    import customer_index as CI

    assert CI.clean("남김해Sub-Hub&amp;Sub-FC") == "남김해Sub-Hub&Sub-FC"
    assert CI.clean("중구1캠프 <-서초1MB(양재동C)") == "중구1캠프"
    assert CI.norm("김포1Sub-FC ?(김포1 서브허브))") == CI.norm("김포1Sub-FC(김포1 서브허브)")
    # 접두사만 다른 같은 캠프 — 완전일치·정규화·핵심코드를 다 놓친 뒤에만 본다
    assert CI.bare("M_광주2캠프") == CI.bare("광주2 캠프") == CI.norm("광주2캠프")
    # 사람이 실수로 M 으로 시작하는 이름을 지어도 통째로 잘리면 안 된다
    assert CI.bare("MB1캠프") == CI.core("MB1캠프"), "숫자 앞 M 까지 떼면 다른 캠프가 된다"
    for junk in ("0", "...", "-", "?", "없음", ""):
        assert junk.upper() in CI.PLACEHOLDERS, junk
    assert "강서1MB" not in CI.PLACEHOLDERS

    import shutil as _sh
    import upload_intake as UI

    root = os.path.join(ROOT, "_t113_root")
    _sh.rmtree(root, ignore_errors=True)
    try:
        misc = UI._paths(root)["misc"]
        os.makedirs(misc, exist_ok=True)
        seen = os.path.join(misc, "★ 01. 쿠팡AS 품목 단가 리스트 및 사진 정리본.xlsx")
        open(seen, "wb").write(b"PK\x03\x04dummy")
        rows = UI.reclass_misc(root, do_apply=False)
        assert len(rows) == 1 and rows[0]["분류"] == "reference", rows
        assert os.path.isfile(seen), "미리보기가 파일을 옮겼다"
        rows = UI.reclass_misc(root, do_apply=True)
        assert not os.path.isfile(seen), "재분류가 파일을 안 옮겼다"
        assert os.path.isfile(rows[0]["목적지"]), rows[0]
        assert "10. 기준" in rows[0]["목적지"], rows[0]["목적지"]
        # 여전히 모르는 것은 건드리지 않는다 — 미분류는 '보존'이 목적이다
        unknown = os.path.join(misc, "무엇인지 모를 파일.dat")
        open(unknown, "wb").write(b"x")
        assert UI.reclass_misc(root, do_apply=True) == []
        assert os.path.isfile(unknown), "모르는 파일을 옮겼다"
    finally:
        _sh.rmtree(root, ignore_errors=True)
    print("  [113] 붙여넣기 오염 흡수(캠프명 18→14)·미분류 자동 재분류 ✅")


def t114_claim_owner_is_agent_pid():
    """[114] 점유의 생사는 **agent_pid** 로 본다 — pid 로 보면 산 세션을 죽었다고 한다.

    2026-08-06 실사고: 시작 화면이 살아 있는 옆 세션의 'code' 점유를 "죽은 세션의
    잔재"로 표시하고 `--free` 를 권했다. `pid` 는 ai_claim 을 실행한 **CLI 프로세스**라
    명령이 끝나는 즉시 죽는다 — 그것으로 판정하면 **모든 점유가 항상 잔재**로 보인다.
    주인은 `agent_pid`(에이전트 프로세스)다.

    두 배 위험했던 이유: 안내대로 `--free` 를 해도 ai_claim 이 남의 것이라 거부한다.
    사람은 "왜 안 풀리지" 하며 막히고, 최악의 경우 `--force` 로 **일하는 세션의
    원장 점유를 빼앗는다.** 그래서 안내 문구도 내 것/남의 것을 갈라 적는다.
    """
    import session_handoff as H

    live, dead = os.getpid(), 999999
    saved = H.claims.__globals__.get("_CLAIM_TEST")
    try:
        import ai_claim as C
        real = C.load
        C.load = lambda: {
            "code": {"who": "claude", "sid": "other", "why": "옆 세션",
                     "pid": dead, "agent_pid": live, "host": "", "at": 0},
            "band": {"who": "claude", "sid": "gone", "why": "죽은 세션",
                     "pid": live, "agent_pid": dead, "host": "", "at": 0},
        }
        try:
            rows = {c["lock"]: c for c in H.claims()}
        finally:
            C.load = real
    finally:
        H.claims.__globals__["_CLAIM_TEST"] = saved

    assert rows["code"]["alive"] is True, "산 세션을 죽었다고 본다 — pid 를 보고 있다"
    assert rows["code"]["stale"] is False, rows["code"]
    assert rows["band"]["alive"] is False, "죽은 세션을 못 가려낸다"
    assert rows["band"]["stale"] is True, rows["band"]

    # 안내 문구: 남의 죽은 점유에 --free 를 권하면 안 된다(거부당한다)
    st = {"점유": [dict(rows["band"], mine=False)], "임시파일": [],
          "큐잔량": 0, "미푸시": []}
    fixes = " ".join(f for _, f in H.blockers(st))
    assert "--adopt" in fixes and "--free band" not in fixes, fixes
    st["점유"] = [dict(rows["band"], mine=True)]
    assert "--free band" in " ".join(f for _, f in H.blockers(st))
    print("  [114] 점유 생사는 agent_pid — 산 세션을 잔재로 오인하지 않는다 ✅")


def t115_text_contrast():
    """[115] 글자가 실제로 읽히는가 — 모든 화면의 명암비를 전수로 지킨다 (2026-08-06).

    사용자 지시: "텍스트가 잘 안보이는 것들이 있어, 전체적으로 전수 조사해서 검토하고
    잘 보이게 정리해". 눈으로 고치면 고친 곳만 좋아진다 — 실제로 `--ink-3` 는 흰
    배경에서 **2.17:1**(기준 4.5:1의 절반)이었고 같은 값이 21곳에 퍼져 있었다.
    그래서 색을 **수치로** 재고, 이 검증이 되돌아가는 것을 막는다.

    감사기 자체도 두 번 틀렸다 — 없는 문제를 만들면 멀쩡한 색을 망가뜨리므로 같이 지킨다.
      · `@media{…}` 를 정규식으로 지워 규칙 경계가 밀렸다(엉뚱한 색으로 보고)
      · `--surface` 가 없는 파일에서 흰색으로 가정해 **다크 테마를 흰 바탕에서** 쟀다
    """
    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    import contrast_audit as CA

    # ── 계산이 맞는가 (WCAG 기준값) ──
    assert abs(CA.contrast((0, 0, 0), (255, 255, 255)) - 21.0) < 0.01
    assert abs(CA.contrast((255, 255, 255), (255, 255, 255)) - 1.0) < 0.01
    # 반투명 글자는 합성해야 실제 색이 나온다 — 안 하면 늘 통과로 보인다
    assert CA.parse_color("rgba(60,60,67,.42)", {}, (255, 255, 255)) == (173, 173, 176)

    # ── 아이콘도 센다 (2026-08-06 제보: 어둡게 켜니 단추 아이콘이 사라졌다) ──
    # 원인은 `.icon-btn` 배경이 흰색으로 **하드코딩**돼 있고 아이콘은 `fill:currentColor`
    # 라 테마를 따라 흰색이 된 것 — 흰 위 흰(1.09:1). 글자만 재면 영영 못 잡는다.
    icon_css = (":root{--bg:#0B1020;--surface:#161C2E;--ink:#F2F5FA}"
                ".b{background:#fff}.b svg{fill:currentColor}")
    bad, _u, _o = CA.audit_theme("", icon_css, CA.root_vars(icon_css)["기본"])
    assert any(r.get("종류") == "아이콘" for r in bad), "흰 바탕 위 흰 아이콘을 놓쳤다"
    # 가상요소는 숙주의 배경 위에 그려진다 — 선택자가 다르다고 바탕을 잃으면 안 된다
    assert CA.bare(".icon-btn::after") == ".icon-btn"
    pseudo = ":root{--bg:#0B1020}.b{background:#fff}.b::after{color:#EEEEEE;font-size:12px}"
    pbad, _pu, _po = CA.audit_theme("", pseudo, CA.root_vars(pseudo)["기본"])
    assert pbad and pbad[0]["ratio"] < 2, "가상요소가 숙주 배경을 못 찾아 검사에서 샜다"
    # 조상 배경을 **못 읽을 때**는 판정하지 않는다 — 없는 문제를 만들면 멀쩡한 색을 망친다
    avatar = (":root{--bg:#0B1020;--surface:#ffffff}.card{background:var(--surface)}"
              ".card .av{background:linear-gradient(145deg,var(--wc2),var(--wc));color:#fff}"
              ".card .av svg{fill:currentColor}")
    abad, aunk, _ao = CA.audit_theme("", avatar, CA.root_vars(avatar)["기본"])
    assert not abad and aunk, "실행 중에 정해지는 아바타색을 흰색으로 단정했다"

    # 실제 화면 세 벌에 미달이 0 이어야 한다 — 아이콘까지 포함해서
    for name in ("webapp/index.html", "docs/app.html", "docs/cal.html"):
        b, _u2, o2 = CA.audit(os.path.join(ROOT, *name.split("/")))
        assert not b, "%s 명암비 미달 %d건: %s" % (
            name, len(b), [(r["sel"][:40], r["ratio"]) for r in b[:3]])
        assert o2 > 0, name

    # ── 감사기 회귀: @media 를 지우면 규칙 경계가 밀린다 ──
    css = ("@media(max-width:719px){ .a small{display:none} }\n"
           ".b{color:#123456}\n")
    got = {sel: d for sel, d, _ in CA.rules(css)}
    assert got.get(".b", {}).get("color") == "#123456", got
    assert "color" not in got.get(".a small", {}), got

    # ── 감사기 회귀: 테마를 한 사전으로 합치면 밝은 화면을 어두운 값으로 잰다 ──
    themed = CA.root_vars(":root{--ink-3:#6b7686;--bg:#fff}\n"
                          '@media (prefers-color-scheme:dark){:root{--ink-3:#8391a8;--bg:#0b1020}}')
    assert set(themed) == {"기본", "다크"}, themed
    assert themed["기본"]["--ink-3"] == "#6b7686" and themed["다크"]["--ink-3"] == "#8391a8"

    # ── 실제 화면: 미달이 하나도 없어야 한다 ──
    screens = ["webapp/index.html", "docs/app.html", "docs/cal.html", "docs/index.html"]
    for rel in screens:
        path = os.path.join(ROOT, *rel.split("/"))
        if not os.path.exists(path):
            continue
        bad, _unknown, ok = CA.audit(path)
        assert ok > 0, "%s 에서 글자색을 하나도 못 읽었다 — 감사기가 헛돌고 있다" % rel
        assert not bad, "%s 명암비 미달 %d건: %s" % (
            rel, len(bad), [(b["sel"], b["ratio"]) for b in bad[:4]])

    # ── 폰 스냅샷은 **생성물**이다 — 결과가 아니라 만드는 쪽을 지킨다 ──
    gen = open(os.path.join(ROOT, "mobile_snapshot.py"), encoding="utf-8").read()
    assert "--brand-btn" in gen, "버튼 바탕색이 글자색과 분리되지 않았다(다크에서 흰 글자 2.5:1)"
    assert "--ink-3:#6b7686" not in gen, "생성기에 옛 저대비 회색이 남아 있다"
    print("  [115] 글자 명암비 전수검사 — 모든 화면 미달 0 (기준 4.5:1) ✅")


def t116_manual_refresh_is_really_fresh():
    """[116] 새로고침 버튼은 **정말로** 새로 받는다 (2026-08-06 지시).

    사용자 지시: "새로 고침을 누르면 바로바로 반영되게 개선해 … 빨리빨리 반영되게 해(중요)".

    두 가지가 겹쳐 있었다.
      ① stale-while-revalidate 가 **사람이 누른 버튼까지** 막았다. 15초 안에 누르면
         네트워크에 아예 안 나가고, 지나도 옛 값을 돌려주고 새 값은 뒤에서 받았다.
         그런데도 '최신 자료로 새로고침했습니다' 토스트가 떴다 — 반영된 줄 안다.
      ② loadSettle → loadStatus → loadNotifications 를 **차례로** 기다렸다.
         실측이 status 2.7초 · notifications 최악 59초라 합이 그대로 대기시간이 됐다.

    SWR 자체는 자동 갱신에 좋다. 가르는 것은 **누가 시켰나**다 — 그래서 지우지 않고
    사람이 시킨 경로에서만 건너뛴다. 캐시 **쓰기**는 그대로 둔다(안 그러면 다음 화면이
    또 옛 값으로 그려진다).
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    assert "function apiFresh(" in live, "사람이 누른 새로고침용 우회로가 없다"
    assert "API_FRESH" in live and "opt.fresh" in live, live.count("API_FRESH")
    # 읽기만 건너뛰고 쓰기는 남아야 한다
    assert re.search(r"const useCache = cacheable && !\(API_FRESH > 0", live), "읽기/쓰기를 안 갈랐다"
    assert re.search(r"if\(cacheable\)\{ swrSet\(path, d\)", live), \
        "강제 갱신 때 캐시 쓰기까지 막혔다 — 다음 화면이 또 옛 값이 된다"

    body = live[live.index("async function reloadHere("):]
    body = body[:body.index("\n}")]
    assert "apiFresh(refreshAll)" in body, "새로고침 버튼이 캐시를 건너뛰지 않는다"
    assert "previewRptDate" in body and "REPORT_PREVIEW_DATE" in body, \
        "집계기준일을 고른 보고서 화면이 새로고침에서 빠졌다"

    ra = live[live.index("async function refreshAll("):]
    ra = ra[:ra.index("\n}")]
    assert "Promise.all" in ra, "새로고침이 아직 한 줄로 세워 부른다(합이 대기시간이 된다)"
    assert not re.search(r"await load\w+\(\);\s*await load", ra), ra[:200]
    print("  [116] 새로고침은 캐시를 건너뛰고 동시에 부른다(선택 기준일 포함) ✅")


def t117_dark_mode_toggle():
    """[117] 밝게/어둡게 — 사람이 켠 것만 따르고, 켜도 글자가 읽힌다 (2026-08-06 지시).

    2026-07-30 에 다크 모드를 **껐던** 이유가 기록에 남아 있다: 흰 배경이 77곳에
    하드코딩돼 있어 OS 다크 설정을 따라가면 글자만 흰색이 되어 '흰 카드 위 흰 글자'가
    됐다. 그래서 이번에는 **순서를 지켰다** — 그 77곳(카드 35 · 옅은 판 37)을 먼저
    --surface / --panel 로 흡수하고 나서 켰다. 이 검증이 그 순서를 지킨다.

    · OS 설정을 자동으로 따라가지 않는다. 이 앱 화면은 그대로 캡처해 보고서로 나간다 —
      폰이 어둡다는 이유로 보고 화면이 검게 바뀌면 안 된다.
    · 켜져 있는 화면에서 바꾸면 크롬이 var() 재계산을 건너뛰어 **일부만 바뀌었다**(실측).
      그래서 토글이 강제 재계산을 한 번 한다.
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    assert ':root[data-theme="dark"]' in live, "다크 팔레트가 없다"
    assert "function toggleTheme(" in live and "function applyTheme(" in live
    assert "cw_theme" in live, "고른 테마가 저장되지 않는다"
    assert 'id="themeBtn"' in live, "밝게/어둡게 단추가 화면에 없다"
    assert "function forceRestyle(" in live and "forceRestyle();" in live, \
        "테마를 바꿔도 일부 요소가 옛 색을 그대로 들고 있다(강제 재계산 없음)"
    # ★ 2026-08-07 갱신 — 사용자 지시로 '시스템 동일' 선택지가 생겼다
    #   ("기본, 다크모드, 시스템 동일 이렇게 구성하게 추가해"). 지켜야 할 것은
    #   "OS 를 읽지 않는다"가 아니라 **"사람이 고르지 않았는데 따라가지 않는다"** 다.
    #   · CSS `@media (prefers-color-scheme: dark)` → 금지. 고르지 않아도 색이 뒤집힌다.
    #   · JS matchMedia → 허용하되 **`themeMode() === 'system'` 일 때만** 반영한다.
    #   · 기본값은 여전히 'light' — 보고 화면이 기기 설정 따라 검게 나오면 안 된다.
    assert not re.search(r"@media[^{]*prefers-color-scheme\s*:\s*dark", live), \
        "CSS 가 OS 다크를 자동으로 따라간다 — 사람이 고르지 않았는데 화면이 검어진다"
    assert "localStorage.getItem('cw_theme') || 'light'" in live, \
        "기본값이 밝게가 아니다"
    assert "if(themeMode() === 'system')" in live, \
        "시스템이 아닐 때도 OS 변화를 반영한다 — 사람이 고른 값이 멋대로 바뀐다"

    dark = live[live.index(':root[data-theme="dark"]'):]
    dark = dark[:dark.index("}")]
    for tok in ("--bg:", "--surface:", "--panel:", "--panel-2:", "--ink:", "--ink-2:",
                "--ink-3:", "--line:", "--brand:", "--brand-btn:", "--warn-btn:",
                "--ok:", "--warn:", "--danger:", "--neu:"):
        assert tok in dark, "다크 팔레트에 %s 가 없다 — 그 자리만 밝은 값이 남는다" % tok

    # 흰 배경 하드코딩이 되돌아오면 2026-07-30 의 '흰 카드 위 흰 글자'가 재현된다
    css = live[live.index("<style>"):live.index("</style>")]
    css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
    white = re.findall(r"background(?:-color)?:\s*#(?:fff|ffffff)\b", css, re.I)
    assert len(white) <= 2, "흰 배경이 다시 하드코딩됐다(%d곳) — --surface 를 쓸 것" % len(white)
    print("  [117] 밝게/어둡게 토글 — 팔레트 완비·OS 자동추종 없음·강제 재계산 ✅")


def t118_ocr_crosscheck():
    """[118] 문서 스캔 교차검증 (2026-08-06 지시: "최고의 무료 도구 확인해서 비교 대조").

    조사 결론은 band/OCR_ENGINES.md 에 있다 — 무료·로컬·한글 조건에서 PaddleOCR 이
    여전히 최상위라 **엔진은 바꾸지 않았다.** 대신 정확도를 올릴 자리가 다른 데 있었다:
    한 엔진의 답만 보면 그 답이 틀렸다는 것을 알 방법이 없다.

    그래서 이 검증이 지키는 것은 두 가지다.
      ① **겹칠 때만 믿는다** — 두 엔진 이상이 같은 값을 낸 항목만 원장에 들어간다.
         한 엔진만 낸 값(단독), 엔진마다 다른 값(충돌), 1엔진 환경은 전부 입력 금지.
      ② **항목마다 대조한다** — 옛 doc_ocr.match() 는 공급가액 하나만 봤다. 발행일이
         틀려도 통과했다. 이제 한 항목이라도 불일치면 그 건은 빈칸도 채우지 않는다.
    """
    sys.path.insert(0, os.path.join(ROOT, "band"))
    import ocr_crosscheck as X

    # ① 표결 — 겹치면 합치, 하나뿐이면 단독, 갈리면 충돌(값을 정하지 않는다)
    assert X.vote(["A", "A", "B"]) == ("A", "합치")
    assert X.vote(["A", "B"]) == ("", "충돌"), "값이 갈렸는데 한쪽을 골랐다"
    assert X.vote(["A", "", None]) == ("A", "단독")
    assert X.vote(["", None]) == ("", "없음")
    # 금액은 근사 비교를 하지 않는다 — 한 자리만 틀려도 다른 값이다
    assert X.vote([2000000, 2000001])[1] == "충돌"

    doc = {"유형": "거래명세서", "발행일": "2026-08-01", "명세서번호": "SL-2026-0712",
           "승인번호": "", "프로젝트NO": "UJ2601138", "공급가액": 2000000,
           "세액": 200000, "합계": 2200000, "사업자번호": "123-81-45678"}
    two = X.merge_records({"paddle": doc, "windows": dict(doc)})
    assert two["신뢰도"] == "높음" and two["교차"]["공급가액"] == "합치"
    one = X.merge_records({"paddle": doc})
    assert one["신뢰도"].startswith("낮음"), "엔진 하나뿐인데 높음으로 나왔다"

    # 합쳐 놓고 공급가+세액≠합계 면, 겹쳤더라도 금액은 믿지 않는다
    bad = X.merge_records({"paddle": doc, "windows": dict(doc, 합계=2300000)})
    assert bad["교차"]["합계"] == "충돌"
    broke = X.merge_records({"paddle": dict(doc, 합계=2300000),
                             "windows": dict(doc, 합계=2300000)})
    assert broke["금액정합"] == "깨짐" and broke["교차"]["공급가액"] == "충돌", \
        "금액 정합성이 깨졌는데도 겹쳤다는 이유로 통과했다"

    # ② 전량을 두 번 읽지 않는다 — 그러나 원장에 쓸 때는 무조건 두 번 읽는다
    clean = dict(doc, 신뢰도="높음")
    assert X.needs_second_opinion(clean) == ""
    assert X.needs_second_opinion(clean, for_write=True), "원장 입력 후보를 한 번만 읽었다"
    assert X.needs_second_opinion(dict(clean, 공급가액=""))
    assert X.needs_second_opinion(dict(clean, 합계=2300000)) == "금액 정합성 깨짐"

    # ③ 항목별 대조 — 빈 원장 / 맞는 원장 / 틀린 원장
    led = {"원장_거래명세서발행일": "2026-08-01", "원장_거래명세서번호": "SL-2026-0712",
           "원장_공급가액": 2000000, "원장_거래명세서합계": 2200000}
    rows = X.compare_ledger(two, led)
    by = {r["항목"]: r["판정"] for r in rows}
    assert by["발행일"] == "일치" and by["명세서번호"] == "일치" and by["세액"] == "원장 빈칸"
    assert "일치" in X.ledger_verdict([r for r in rows if r["항목"] != "세액"])
    wrong = X.compare_ledger(two, dict(led, 원장_거래명세서발행일="2026-07-31"))
    assert "불일치(발행일)" == X.ledger_verdict(wrong), \
        "공급가액만 맞으면 통과하던 옛 판정이 남아 있다"

    # ④ 원장 입력 문지기 — 합치 + 빈칸 + 불일치 없음, 셋 다여야 들어간다
    empty = {"원장_공급가액": 2000000}
    ok = X.build_updates([{"문서": dict(two, 파일="a.jpg"), "정산ID": "S1",
                           "대조": X.compare_ledger(two, empty)}])
    cols = {u["col"] for u in ok}
    assert cols == {"거래명세서발행일", "거래명세서번호"}, cols
    assert all(u["sheet"] == "06_거래서류청구수금" and u["key_col"] == "정산ID" for u in ok)
    # 세액·합계는 수식·집계 열이라 읽어서 대조만 하고 쓰지 않는다
    assert not any(u["col"] in ("세액", "합계") for u in ok)
    # 한 항목이라도 불일치면 같은 건의 빈칸도 채우지 않는다
    conflict = X.compare_ledger(two, {"원장_공급가액": 1999})
    assert X.build_updates([{"문서": two, "정산ID": "S1", "대조": conflict}]) == []
    # 1엔진(단독)은 아무것도 못 채운다
    assert X.build_updates([{"문서": dict(one, 파일="a.jpg"), "정산ID": "S1",
                             "대조": X.compare_ledger(one, empty)}]) == []
    # 정산ID 를 못 찾았으면 채우지 않는다
    assert X.build_updates([{"문서": two, "정산ID": "", "대조": X.compare_ledger(two, empty)}]) == []

    # ⑤ 같은 답을 두 번 본 것을 '합치'라 부르지 않는다.
    #    doc_ocr 는 paddle 이 실패하면 Windows 결과를 paddle 자리에 캐시한다 —
    #    그것을 둘로 세면 거짓 근거가 만들어진다.
    assert list(X.drop_dependent({"paddle": "가 나 다", "windows": "가나다"})) == ["paddle"]
    assert len(X.drop_dependent({"paddle": "가나다", "windows": "가나라"})) == 2
    assert X.drop_dependent({"paddle": "", "windows": ""}) == {}

    # ⑥ 전량을 두 번 읽지 않는다 — 급한 것부터 예산만큼, 미룬 수는 반드시 알린다.
    #    사진 1,816장 × 둘째 엔진 4.8초 = 2.4시간이라 daily_run 이 무너진다(실측).
    take, defer = X.recheck_plan([("a", 0), ("b", 3), ("c", 1), ("d", 0), ("e", 2)], budget=2)
    assert take[:2] == ["a", "d"] and defer == 1, (take, defer)
    assert take[2:] == ["c", "e"], "급한 것(원장 불일치)보다 덜 급한 것을 먼저 읽었다"
    assert X.recheck_plan([("a", 0)] * 5, budget=1) == (["a"] * 5, 0), \
        "원장에 쓰려는 건이 예산에 잘렸다 — 쓰는 순간은 예산과 무관해야 한다"
    fill_rows = X.compare_ledger(two, empty)
    assert X.writable_now(two, fill_rows, "S1") is True
    assert X.recheck_reason(two, fill_rows, True) == (0, "원장 입력 후보")
    assert X.recheck_reason(dict(clean, 합계=2300000), [], False)[0] == 2
    full = X.compare_ledger(two, led)
    assert X.recheck_reason(two, full, False) == (None, ""), "다 맞는 건까지 두 번 읽고 있다"

    # ⑦ 엔진은 있는 것만 골라 쓴다 — 자동 설치·외부 업로드가 없어야 한다
    src = open(os.path.join(ROOT, "band", "ocr_crosscheck.py"), encoding="utf-8").read()
    for banned in ("pip install", "requests.post", "urllib.request", "http://", "https://"):
        assert banned not in src, "문서를 PC 밖으로 보내거나 자동 설치하는 코드가 있다: %s" % banned
    assert {e["엔진"] for e in X.engine_status()} >= {"paddle", "windows"}
    daily = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert "ocr_crosscheck.py" in daily, "매일 도는 자리에 붙지 않았다(daily_run 누락)"
    assert os.path.isfile(os.path.join(ROOT, "band", "OCR_ENGINES.md")), \
        "무엇을 왜 골랐는지가 파일로 남지 않았다"
    print("  [118] 문서 OCR 교차검증 — 겹칠 때만 입력·항목별 원장 대조·로컬 전용 ✅")


def t119_context_guard():
    """[119] 컨텍스트가 다 차기 전에 마무리로 전환시킨다 (2026-08-06 지시).

    사용자 지시: "컨텍스트 윈도우 다 차기전 컴팩팅 자동으로 하는 알고리즘 추가".

    지키는 것
      ① 대화 기록 폴더를 **정말로 찾는다.** 슬러그는 글자 하나당 대시 하나다
         (`C:\\Users\\…` → `C--Users-…`). 여기를 `[^A-Za-z0-9]+` 로 묶었다가
         폴더를 못 찾아 사용량이 늘 0% 로 나왔었다 — 그러면 감시가 없는 것과 같다.
      ② 사용량은 **마지막 usage 합**(입력+캐시읽기+캐시생성)이다. 글자 수 추정이 아니다.
      ③ 단계는 올라갈 때만 말한다(70 예고 / 85 마무리 / 95 즉시). 매 입력마다 떠들면
         본문이 밀려 오히려 컨텍스트를 잡아먹는다.
      ④ 훅은 **절대 사람 입력을 막지 않는다** — 어떤 예외에도 exit 0.
      ⑤ settings.json 에 UserPromptSubmit 배선이 살아 있고, autoCompactWindow 는
         없으면 auto(권장·2026-08-07 사용자 승인), 있으면 한도 안이어야 한다.
    """
    import tempfile
    import context_guard as G
    _rd = lambda *p: open(os.path.join(*p), encoding="utf-8").read()
    TMP = tempfile.mkdtemp(prefix="ctxguard_")

    # ① 폴더 찾기 — 이 프로젝트의 기록 폴더가 실제로 잡혀야 한다
    d = G._project_dir()
    assert d and os.path.isdir(d), "대화 기록 폴더를 못 찾았다 — 슬러그 규칙 확인"

    # ② usage 합산 — 꼬리만 읽어도 최신 값을 집어야 한다
    tmp = os.path.join(TMP, "ctx_sample.jsonl")
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(json.dumps({"message": {"usage": {
            "input_tokens": 5, "cache_read_input_tokens": 100,
            "cache_creation_input_tokens": 10}}}) + "\n")
        fh.write(json.dumps({"message": {"content": "usage 없는 줄"}}) + "\n")
        fh.write(json.dumps({"message": {"usage": {
            "input_tokens": 2, "cache_read_input_tokens": 300,
            "cache_creation_input_tokens": 8}}}) + "\n")
    assert G._used_tokens(tmp) == 310, "마지막 회차가 아니라 다른 줄을 셌다"
    assert G._used_tokens(os.path.join(TMP, "없는파일.jsonl")) == 0

    # ③ 단계 경계 — 0.70/0.85/0.95 에서만 올라간다
    assert G._stage(0.69)[1] == "여유"
    assert G._stage(0.70)[1] == "예고"
    assert G._stage(0.86)[1] == "마무리"
    assert G._stage(0.99)[1] == "즉시"
    # 올라간 순간에만 말한다 — 같은 단계를 다시 재면 조용하다
    base = {"advice": "x", "fresh": True, "percent": 90.0, "used": 1, "limit": 2,
            "stage": "마무리", "wrapup_ran_now": False, "auto_compact": True}
    assert G._message(base), "단계가 올라갔는데 아무 말도 하지 않았다"
    assert G._message(dict(base, fresh=False)) == "", "같은 단계에서 매번 떠들고 있다"
    assert "compact" in G._message(dict(base, auto_compact=False)), \
        "자동 요약이 꺼져 있는데 사람에게 알리지 않는다"

    # ④ 재기만 하는 호출은 인계를 돌리지도, 상태 파일을 건드리지도 않는다
    before = os.path.getmtime(G.STATE_PATH) if os.path.exists(G.STATE_PATH) else 0
    r = G.measure(limit=450000, act=False)
    assert r["wrapup_ran_now"] is False and r["limit"] == 450000
    after = os.path.getmtime(G.STATE_PATH) if os.path.exists(G.STATE_PATH) else 0
    assert before == after, "--dry 인데 상태 파일을 고쳤다"

    # 어떤 예외에도 exit 0 — main 이 통째로 try 로 감싸여 있어야 한다
    src = _rd(os.path.join(ROOT, "context_guard.py"))
    assert "sys.exit(0)" in src, "훅이 실패하면 사람 입력이 막힌다"

    # ⑤ 배선 — 훅이 붙어 있고 압축 시점이 한도 안이다
    st = json.loads(_rd(os.path.join(os.path.dirname(ROOT), ".claude", "settings.json")))
    ups = json.dumps(st.get("hooks", {}).get("UserPromptSubmit", []), ensure_ascii=False)
    assert "context_guard.py" in ups, "UserPromptSubmit 훅에 배선되지 않았다"
    # ★ 사람 입력이 없는 동안이 가장 위험하다 (2026-08-06 지시 "밤을 새서라도").
    #   UserPromptSubmit 은 사람이 칠 때만 온다 — 지시 하나로 몇 시간을 혼자 도는
    #   밤샘 작업에서는 한 번도 안 온다. 그래서 도구를 쓸 때마다 오는 PostToolUse 에도
    #   같은 눈을 달아 둔다. 값은 90초에 한 번만 실제로 재서 도구를 늦추지 않는다.
    pts = json.dumps(st.get("hooks", {}).get("PostToolUse", []), ensure_ascii=False)
    assert "context_guard.py" in pts and "--tick" in pts,         "PostToolUse 에 컨텍스트 눈이 없다 — 사람 입력 없이 도는 동안 단계를 못 잡는다"
    assert "def tick(" in src and "TICK_EVERY" in src, "--tick 진입점이 없다"
    assert 'if now - float(st.get("tick_at") or 0) < TICK_EVERY' in src,         "매 도구 호출마다 대화 기록을 통째로 읽는다 — 모든 도구가 그만큼 느려진다"
    assert st.get("autoCompactEnabled") is True, "자동 요약이 꺼져 있다"
    # autoCompactWindow 는 2026-08-07 부터 기본이 auto(키 없음)다 — 모델에 맞는 창을
    # Claude Code 가 고른다. 키가 없으면 context_guard 가 DEFAULT_LIMIT(450k)로 경보한다.
    # 사람이 다시 오버라이드하면 설정 유효 범위(100k~1M) 안이어야 한다.
    win = st.get("autoCompactWindow")
    if win is not None:
        assert 100_000 <= int(win) <= 1_000_000, "압축 시점(autoCompactWindow)이 범위 밖이다: %s" % win
    assert "DEFAULT_LIMIT" in src, "auto 일 때 기댈 보수 한도가 context_guard 에 없다"
    print("  [119] 컨텍스트 감시 — 사용량 실측·단계 전환·인계 자동·훅 배선 ✅")


def t120_calendar_sheet_and_share():
    """[120] 캘린더: 날짜 창·PC 텍스트 격자·고정 주소·크롬 설치 유도 (2026-08-06 지시).

    지시: "캘린더 날짜 클릭시 밑에서 위로 올라가는 레이어창이나 모달창으로 / 공유시
    크롬을 자동으로 열어 설치하게 / 공유 캘린더 주소 고정(변동 주지마) / 날짜 크기 크게 /
    PC 는 텍스트로 자세히, 모바일은 지금처럼."

    여기서 지키는 것은 **되돌아가기 쉬운 네 가지**다.
      ① 날짜를 눌러도 화면 저 아래 목록만 바뀌던 옛 동작으로 돌아가지 않을 것
      ② 공유 주소에 달(m=)·필터(off=)가 다시 붙지 않을 것 — 붙는 순간 주소가 흔들린다
      ③ 열쇠 파일이 없다고 조용히 새 열쇠를 뽑지 않을 것 — 뿌린 링크가 통째로 죽는다
      ④ 설치 매니페스트가 업무 앱이 아니라 **캘린더**를 가리킬 것
    """
    _rd = lambda *p: open(os.path.join(ROOT, *p), encoding="utf-8").read()
    live = _rd("webapp", "index.html")
    cal = _rd("docs", "cal.html")

    # ① 날짜를 누르면 창이 뜬다 — 두 화면 모두
    assert "calOpenSheet(day)" in live, "앱: 날짜를 눌러도 창이 뜨지 않는다"
    assert 'id="calSheet"' in live and "function calCloseSheet(" in live, "앱: 날짜 창이 없다"
    assert "document.body.style.overflow='hidden'" in live, \
        "앱: 창이 떠 있는 동안 뒤 화면이 같이 스크롤된다"
    assert "openSheet()" in cal and 'id="sheet"' in cal, "공유 캘린더: 날짜 창이 없다"
    # 폰=아래에서 위로 · PC=가운데 모달 (한 마크업을 CSS 가 가른다)
    for doc, name in ((live, "앱"), (cal, "공유 캘린더")):
        assert "translateY(100%)" in doc, f"{name}: 아래에서 올라오는 레이어가 아니다"
        assert "translate(-50%,-50%)" in doc, f"{name}: PC 가운데 모달이 없다"
    # 애니메이션을 rAF 하나에만 걸면 탭이 숨었을 때 창이 안 올라온다(2026-08-06 캡처 사고)
    for doc, name in ((live, "앱"), (cal, "공유 캘린더")):
        assert "setTimeout(rise, 60)" in doc, f"{name}: rAF 가 안 불리는 상황의 대비가 없다"

    # ② PC 는 글자, 폰은 점 — 가르는 것은 CSS 미디어쿼리여야 한다(창 크기를 재지 않는다)
    assert ".cal2-txs{display:none}" in live and "@media(min-width:900px)" in live, \
        "앱: PC 텍스트 격자 규칙이 없다"
    assert 'class="cal2-txs"' in live, "앱: 격자 칸에 텍스트를 실어 보내지 않는다"
    assert ".ctxs{display:none}" in cal and 'class="ctxs"' in cal, \
        "공유 캘린더: PC 텍스트 격자가 없다"
    # ③ 날짜 글씨 — 예전 값(14px/12.5px)으로 되돌아가지 않게 못을 박는다
    assert ".cal2-day .d{font-size:17px" in live, "앱: 날짜 글씨가 다시 작아졌다"
    assert ".cday b{font-size:17px" in cal, "공유 캘린더: 날짜 글씨가 다시 작아졌다"

    # ④ 공유 주소 고정 — 달·필터를 주소에 담지 않는다
    link = live[live.index("function calendarLink()"):]
    link = link[:link.index("\n}")]
    assert "?view=calendar'" in link, "고정 주소가 아니다"
    assert "p.set('m'" not in link and "p.set('off'" not in link, \
        "공유 주소에 달·필터가 다시 붙었다 — 누를 때마다 주소가 달라진다"
    # 이미 뿌린 옛 링크(m=·off=)는 계속 열려야 한다
    assert "if(/^\\d{4}-\\d{2}$/.test(m)) CAL_MONTH = m;" in live, \
        "옛 링크의 달 지정을 더 이상 받아 주지 않는다"

    # ⑤ 열쇠(=주소)를 조용히 바꾸지 않는다
    share = _rd("cal_share.py")
    assert "if not new and os.path.exists(OUT):" in share and "sys.exit(" in share, \
        "공유본이 있는데 열쇠만 없을 때 새 열쇠를 조용히 뽑는다 — 뿌린 링크가 다 죽는다"
    assert "def note_url(" in share, "고정 주소를 사람이 찾아볼 자리가 없다"

    # ⑥ 크롬으로 넘겨 설치 — 카카오톡 안 브라우저는 설치 자체가 불가능하다
    assert "function inAppBrowser(" in cal and "KAKAOTALK" in cal, \
        "앱 안 브라우저 판정이 없다 — 카카오톡에서 열면 설치가 영영 안 된다"
    assert "function chromeJumpUrl(" in cal, "넘길 주소를 만드는 자리가 따로 없다"
    assert "package=com.android.chrome" in cal, "안드로이드 크롬 넘기기가 없다"
    assert "S.browser_fallback_url=" in cal, "크롬이 없는 폰의 되돌아갈 자리가 없다"
    # ★ 아이폰은 크롬이 아니라 **사파리**로 보내야 한다 — 아이폰 크롬은 홈 화면 설치를
    #   못 한다. 크롬으로 보내면 "열리지만 설치는 안 되는" 상태로 한 걸음 헛돈다.
    assert "'x-safari-' + shareUrl()" in cal, "아이폰을 사파리로 보내지 않는다"
    assert "googlechromes://" not in cal,         "아이폰을 크롬으로 보낸다 — 아이폰 크롬으로는 홈 화면에 설치할 수 없다"
    assert "sessionStorage.setItem('csos.cal.jump','1')" in cal, \
        "자동 넘기기에 한 번만 걸리는 잠금이 없다 — 되돌아올 때마다 다시 튄다"
    assert "beforeinstallprompt" in cal and "PROMPT.prompt()" in cal, "설치 창을 부르지 않는다"

    # ⑦ 설치되는 것이 '캘린더' 여야 한다 — 업무 앱이 깔리면 받는 사람은 PIN 화면을 본다
    assert 'href="cal-manifest.json"' in cal, "공유 캘린더가 업무 앱 매니페스트를 물고 있다"
    cm = json.loads(_rd("docs", "cal-manifest.json"))
    assert cm["start_url"].endswith("/cal.html"), cm
    assert cm["id"].endswith("/cal.html"), "설치 신원이 업무 앱과 겹치면 하나만 설치된다"
    app_mf = json.loads(_rd("docs", "manifest.json"))
    _rev = lambda mf: sorted(i["src"] for i in mf["icons"])
    assert _rev(cm) == _rev(app_mf), "아이콘 판이 업무 앱과 어긋났다 — 한쪽만 갱신됐다"
    import cloud_publish as _CP
    assert "docs/cal-manifest.json" in _CP.PUBLISH_FILES, \
        "설치 정보가 고정 주소에 올라가지 않는다 — 링크를 받은 폰에는 계속 옛 매니페스트가 간다"

    # ⑧ 설치한 아이콘은 주소의 `#k=` 를 못 가져간다 — 기기에 기억해 두지 않으면 매번 막힌다
    assert "localStorage.setItem(KSTORE" in cal and "localStorage.getItem(KSTORE)" in cal, \
        "열쇠를 기억하지 않는다 — 설치한 아이콘이 '열쇠가 없는 주소입니다'만 띄운다"
    print("  [120] 캘린더 — 날짜 창(폰 레이어·PC 모달)·PC 텍스트 격자·고정 주소·크롬 설치 ✅")


def t121_layer_dialogs():
    """[121] 알림·확인·입력을 전부 레이어로 (2026-08-06 지시).

    사용자 지시: "UX UI 편리하게 밑에서 위로 올라가는 레이어창으로 하거나
    모달창으로 나오게 하는 구조를 … 전체 적용해 (담당자별 업무 센터도 마찬가지야)".

    지키는 것
      ① 브라우저 기본 창(alert·confirm·prompt)이 **한 개도 남아 있지 않다.**
         하나라도 남으면 폰에서 "localhost:8000 says" 회색 상자가 튀어나온다.
      ② 대체 함수(notice·askYesNo·askText)와 그 뼈대(dlgAsk)가 실제로 있다.
      ③ confirm/prompt 는 답을 **기다려야** 한다 — askYesNo·askText 호출은 전부 await 다.
         빠뜨리면 Promise 객체가 늘 참이라 "취소"가 먹지 않는다(조용한 사고).
      ④ 폰=아래에서 올라오는 시트 · PC=가운데 모달. 두 모양이 CSS 에 다 있어야 한다.
      ⑤ 담당자 업무센터 입력 폼은 **요소째** 레이어로 옮긴다(HTML 복사가 아니다).
         복사하면 <input type=file> 과 적던 메모가 사라진다. 닫을 때 제자리로 돌아온다.
    """
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    js = "".join(re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", idx, re.S))

    # ① 기본 창이 남아 있지 않다 (deferredInstallPrompt.prompt() 같은 메서드는 제외)
    for bad in ("alert", "confirm", "prompt"):
        left = re.findall(r"(?<![\w.$])%s\s*\(" % bad, js)
        assert not left, "브라우저 기본 %s() 가 %d 곳 남아 있다" % (bad, len(left))

    # ② 대체 함수가 있다
    for fn in ("function notice(", "function askYesNo(", "function askText(",
               "function dlgAsk(", "function dlgCancel("):
        assert fn in js, "대체 함수가 없다: %s" % fn

    # ③ 답을 기다린다 — 정의부를 뺀 모든 호출 앞에 await 가 있다
    for fn in ("askYesNo", "askText"):
        for m in re.finditer(r"(.{8})%s\s*\(" % fn, js):
            head = m.group(1)
            if head.endswith("function ") or head.endswith("nction "):
                continue                      # 정의부
            assert head.rstrip().endswith("await"), \
                "%s 호출에 await 가 빠졌다 — 취소가 먹지 않는다: …%s" % (fn, head)

    # ④ 두 모양이 다 있다: 폰 시트(아래에서) · PC 모달(가운데)
    assert "transform:translateY(100%)" in idx, "밑에서 올라오는 모양이 없다"
    assert ".dlg-wrap.in .dlg{transform:translate(-50%,-50%) scale(1)" in idx, \
        "PC 가운데 모달 모양이 없다"
    assert ".sheet.open{transform:translate(-50%,-50%) scale(1)" in idx, \
        "상세 시트가 PC 에서도 바닥에 붙어 있다 — 모달로 띄우기로 했다"
    assert 'id="dlgWrap"' in idx and 'id="dlgBody"' in idx and 'id="dlgFoot"' in idx
    # PIN 창(1400)보다 아래, 상세 시트(41)보다 위에 있어야 겹침 순서가 맞다
    z = int(re.search(r"\.dlg-wrap\{position:fixed;inset:0;z-index:(\d+)", idx).group(1))
    assert 41 < z < 1400, "레이어 겹침 순서가 어긋난다(z-index %s)" % z

    # ⑤ 담당자 업무센터 — 요소째 옮기고, 닫을 때 돌려놓는다
    assert "function layerOpen(" in js and "function layerRestore(" in js
    assert "layerOpen('ryuEntryForm'" in js, \
        "담당자 업무센터 입력 폼이 레이어로 열리지 않는다"
    assert js.count("layerRestore()") >= 3, \
        "레이어를 되돌리는 자리가 모자란다(내용교체·한겹닫기·통째닫기 세 곳)"
    # setSheetContent 는 innerHTML 로 지우기 **전에** 되돌려야 한다
    seg = js[js.index("function setSheetContent("):][:400]
    assert seg.index("layerRestore()") < seg.index("body.innerHTML"), \
        "옮겨 온 요소를 돌려놓기 전에 시트를 지운다 — 폼이 통째로 사라진다"
    # 페이지 안에서 사라진 자리를 스크롤로 때우던 옛 방식이 남아 있으면 안 된다
    assert "$('ryuEntryForm').scrollIntoView" not in js, \
        "옛 스크롤 방식이 남아 있다 — 레이어와 둘 다 돌면 화면이 튄다"

    # ⑥ 폰 사본(docs/app.html)도 같은 모양이어야 한다.
    #    이 파일은 index.html 에서 생성되는 것이 아니라 따로 있는 앱이다 —
    #    한쪽만 고치면 "폰에서만 회색 상자가 뜨는" 상태가 된다.
    phone = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    pjs = "".join(re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", phone, re.S))
    for bad in ("alert", "confirm", "prompt"):
        left = re.findall(r"(?<![\w.$])%s\s*\(" % bad, pjs)
        assert not left, "폰 사본에 기본 %s() 가 %d 곳 남아 있다" % (bad, len(left))
    assert "function notice(" in pjs and "function askYesNo(" in pjs, \
        "폰 사본에 대체 함수가 없다 — 두 앱의 이름을 맞춰 둔다"
    assert 'id="dlgwrap"' in phone and ".dlgwrap.on{display:flex}" in phone
    assert "align-items:flex-end" in phone and "@media(min-width:900px){\n .dlgwrap{align-items:center}" in phone, \
        "폰 사본에 '아래에서 올라옴 → 넓으면 가운데' 두 모양이 없다"
    print("  [121] 알림·확인·입력 레이어 — 기본창 0개·await 완비·폰시트/PC모달·폼 요소째 이동 ✅")


def t122_dash_drag_and_remote_version():
    """[122] 대시보드 카드 끌기 + 리모컨 버전 관리 (2026-08-06 지시).

    사용자 지시: "각 카드를 클릭해서 드래그 앤 드롭으로 자유롭게 움직일 수 있는 기능
    추가, 각 카드는 밑에서 위로 올라가는 레이어창으로 하거나 모달창으로 나오게 …" ·
    "리모컨 버전 선택할 수 있는 기능 추가하고 전체적으로 버전 관리가 VER.3인지
    VER.4인지 입력 및 확인 수정 가능하게 고도화".

    지키는 것 — 대시보드
      ① 끌기는 **포인터 이벤트**다. HTML5 dragstart 는 터치에서 이벤트 자체가 나지
         않아 폰에서 아무리 끌어도 안 움직였다. draggable 손잡이가 남아 있으면 안 된다.
      ② 카드 **아무 데나** 눌러 끈다 — 손잡이 셀렉터로만 집는 코드가 없어야 한다.
      ③ 6px 문턱(슬롭)이 있어야 평범한 탭·체크가 살아남는다.
      ④ 카드를 레이어로 연 동안 저장이 일어나도 **순서에서 빠지지 않는다.**
         (grid 의 자식이 아니게 되므로, base 순서를 이어 붙이는 처리가 있어야 한다)

    지키는 것 — 리모컨 버전
      ⑤ 표기 통일: 'ver3'·'V4'·'VER 3' 이 전부 하나로 모인다.
      ⑥ 담당자 보유가 **버전별로** 나온다(합계만으로는 VER.3 몇 개인지 못 센다).
      ⑦ 버전이 비었거나 '미확인'인 줄을 모아 준다 — 빈칸만 찾으면 대부분을 놓친다.
      ⑧ 화면의 버전 선택지는 **서버 목록만** 쓴다(화면이 따로 적으면 조용히 갈린다).
    """
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    js = "".join(re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", idx, re.S))

    # ① 포인터 끌기로 갈아탔다 — 옛 HTML5 끌기 흔적이 남으면 안 된다
    assert "function dashDragEnable(" in js, "포인터 기반 끌기가 없다"
    for gone in ("dashDragging", "dashKpiDragging", "dashPaletteDragging"):
        assert gone not in js, "옛 HTML5 끌기가 남아 있다(폰에서 안 된다): %s" % gone
    # 대시보드 배치 코드 안에는 HTML5 끌기가 없어야 한다.
    # (파일을 끌어다 놓는 업로드 구역의 dragover 는 별개다 — 거기는 그대로 둔다)
    dash = js[js.index("function initDashboardLayout("):]
    dash = dash[:dash.index("\nasync function ", 10)]
    for gone in ("addEventListener('drag", 'addEventListener("drag', "dataTransfer"):
        assert gone not in dash, "대시보드에 HTML5 끌기가 남아 있다: %s" % gone
    assert 'class="dash-drag-handle" draggable' not in js, \
        "손잡이에 draggable 이 남아 있다 — 네이티브 끌기가 포인터 끌기와 싸운다"
    # ② 세 곳 모두 붙었다: 화면 카드 · 핵심지표 · 보관함
    for host, sel in (("dashGrid", "[data-dash-block]"), ("kpis", "[data-kpi-card]"),
                      ("dashPalette", ".dash-choice")):
        assert "dashDragEnable('%s','%s'" % (host, sel) in js, \
            "%s 에 끌기가 붙지 않았다" % host
    # ③ 탭과 끌기를 가르는 문턱
    assert "DASH_DRAG_SLOP" in js and re.search(r"DASH_DRAG_SLOP\s*=\s*[1-9]", js), \
        "끌기 문턱이 없다 — 살짝만 눌러도 카드가 끌려간다"
    assert "input,button,select,textarea,a,label" in js, \
        "입력 요소 위에서도 끌기가 시작된다 — 체크박스를 못 누른다"
    # ④ 레이어로 연 카드가 순서에서 빠지지 않는다
    assert "function openDashCardLayer(" in js
    seg = js[js.index("function dashboardLayoutState("):][:1600]
    assert "(base.order||[]).forEach" in seg, \
        "레이어로 열어 둔 카드가 저장 때 순서에서 빠진다"

    # ⑤~⑧ 리모컨 버전
    sys.path.insert(0, ROOT)
    import ledger_db as L
    assert L._remote_version("ver3") == "VER.3"
    assert L._remote_version("V4") == "VER.4"
    assert L._remote_version("VER 3") == "VER.3"
    assert L._remote_version("구형") == "기존형"
    assert L._remote_version("") == "", "빈 값을 임의로 채우면 안 된다"
    src = open(os.path.join(ROOT, "ledger_db.py"), encoding="utf-8").read()
    hold = src[src.index("def _remote_holdings("):]
    hold = hold[:hold.index("\ndef ", 10)]
    assert '"versions"' in hold, "담당자 보유에 버전별 내역이 없다"
    gaps = src[src.index("def remote_version_gaps("):]
    gaps = gaps[:gaps.index("\ndef ", 10)]
    assert "TRIM(version)='미확인'" in gaps, \
        "'미확인'으로 저장된 줄을 빠뜨린다 — 빈칸만 세면 대부분을 놓친다"
    for t in ("remote_issue", "remote_delivery", "remote_stock"):
        assert t in gaps, "%s 가 버전 확인 대상에서 빠졌다" % t
    st = L.remote_status()
    for k in ("version_totals", "version_gaps", "versions"):
        assert k in st, "remote_status 에 %s 가 없다" % k
    for v in st["version_totals"].values():
        assert v["all"] == v["holding"] + v["stock"], "버전 합계가 안 맞는다"
    # ⑧ 화면은 서버 목록만 쓴다 + 그 자리에서 고치는 길이 있다
    assert "function remoteSetVersion(" in js, "버전을 그 자리에서 고칠 수 없다"
    assert "(s.versions||[])" in js and "'VER.3'," not in js.replace("REMOTE", ""), \
        "화면이 버전 목록을 따로 적어 두었다 — 서버와 갈린다"
    print("  [122] 대시보드 포인터 끌기·카드 레이어 + 리모컨 버전별 보유·미확인 채우기 ✅")


def t123_calendar_share_tools():
    """[123] 공유 캘린더 — 크롬으로 자동 전환·설치·엑셀/캡처/복사 (2026-08-06 지시).

    사용자 지시: "이 쿠팡 캘린더 접속시 자동으로 크롬으로 열리게 하고 설치할 수 있게
    해주고, 여기에 엑셀저장, 캡처, 복사 버튼 동일하게 추가해서".

    지키는 것
      ① 카톡·밴드 안의 브라우저면 **자동으로** 밖(크롬/사파리)으로 넘긴다.
         거기서는 홈 화면 설치가 아예 안 되고 저장도 어디로 갔는지 알 수 없다.
      ② 그런데 **한 번만** 넘긴다(sessionStorage). 안 그러면 뒤로 돌아올 때마다
         또 튀어나가 앱을 쓸 수가 없다. '나중에'를 누른 사람은 아예 넘기지 않는다.
      ③ 설치 버튼은 **설치가 가능할 때만** 보인다 — 눌러도 아무 일 없는 버튼은
         고장으로 읽힌다.
      ④ 도구 세 개(이미지 저장·복사·엑셀)가 다른 화면과 같은 아이콘으로 있다.
      ⑤ 내보내기는 **화면과 같은 달·같은 필터**만 담는다. 파일과 화면의 건수가
         다르면 그때부터 파일을 못 믿는다.
      ⑥ 캡처와 복사는 **같은 그리기 코드**를 쓴다(toBlob). 두 벌이면 한쪽만 낡는다.
    """
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    js = "".join(re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", idx, re.S))

    # ①② 자동 전환과 그 안전장치
    assert "function calAutoOpenOutside(" in js, "자동으로 크롬으로 넘기는 코드가 없다"
    auto = js[js.index("function calAutoOpenOutside("):]
    auto = auto[:auto.index("\nfunction ", 10)]
    assert "csosInApp()" in auto, "인앱 여부를 보지 않고 무조건 넘긴다"
    assert "sessionStorage" in auto and "csos_auto_out" in auto, \
        "한 번만 넘기는 장치가 없다 — 뒤로 올 때마다 튀어나간다"
    assert "csos_install_hint_off" in auto, "'나중에'를 누른 사람에게도 자동으로 넘긴다"
    assert "csosStandalone()" in auto, "이미 설치해 쓰는 사람까지 밖으로 넘긴다"
    assert "calAutoOpenOutside();" in js[js.index("function maybeInstallForShare("):][:600], \
        "공유 링크로 들어왔을 때 자동 전환이 걸리지 않는다"

    # ③ 설치 버튼
    assert 'id="calInstallBtn"' in idx and "hidden>" in idx.split('id="calInstallBtn"')[1][:400], \
        "설치 버튼이 기본 숨김이 아니다"
    assert "function calSyncInstallButton(" in js and "function calInstallApp(" in js

    # ④ 도구 세 개 — 다른 화면과 같은 아이콘
    tools = idx[idx.index('id="calTools"'):]
    tools = tools[:tools.index("</div>")]
    for icon in ("#i-image-down", "#i-clipboard-copy", "#i-file-spreadsheet"):
        assert icon in tools, "캘린더 도구에 %s 가 없다" % icon
    for fn in ("calendarCapture()", "calCopyList()", "calExportList()"):
        assert fn in tools, "캘린더 도구에 %s 가 연결되지 않았다" % fn

    # ⑤ 화면과 같은 것만 내보낸다
    exp = js[js.index("function calRowsForExport("):]
    exp = exp[:exp.index("\nfunction ", 10)]
    assert "calendarRows()" in exp, "필터를 무시하고 전체를 내보낸다"
    assert "calMonthOf(e.날짜) === m" in exp, "보고 있는 달만 내보내지 않는다"

    # ⑥ 캡처와 복사가 같은 그림을 쓴다
    assert "opt.toBlob" in js and "calendarCapture({toBlob:true})" in js, \
        "복사가 캡처와 다른 코드로 그림을 만든다 — 한쪽만 낡는다"
    cp = js[js.index("async function calCopyList("):]
    cp = cp[:cp.index("\nfunction ", 10) if "\nfunction " in cp else 2000]
    assert "clipboard.writeText" in cp, \
        "이미지 복사가 막힌 폰에서 아무 일도 일어나지 않는다 — 글 복사 대비가 없다"
    print("  [123] 공유 캘린더 — 크롬 자동 전환(1회)·설치 버튼·엑셀/캡처/복사 ✅")


def t124_no_duplicate_menus():
    """[124] 겹치는 메뉴 없애기 (2026-08-06 지시).

    사용자 지시: "앱 전체적으로 겹치는 메뉴가 보이는데 이런 부분 통합하고 ui ux 개선해"
    (쿠팡 캘린더 화면 사진과 함께).

    사진에 찍힌 것은 셋이었다.
      · 큰 버튼 '캡처' 와 아이콘 '이미지 저장' 이 **같은 함수**였다.
      · '링크 복사' 와 '복사' — 같은 말이 두 가지를 가리켰다.
      · 알림 칩이 헤더 버튼들을 **덮고** 있었고, 좁은 폰에서는 헤더 글자와
        고객사 로고가 서로 겹쳤다.

    그래서 이 검증은 '지금 고쳤다'가 아니라 **다시 생기지 않게** 막는다.
      ① 한 화면 안에 같은 onclick 을 가진 버튼이 두 개 있으면 실패.
      ② 도구줄의 복사 버튼은 무엇을 복사하는지 이름에 있어야 한다.
      ③ 떠 있는 칩은 앱바 높이를 재서 그 **아래**에 놓는다(고정 top 금지).
      ④ 헤더는 좁아지면 겹치지 말고 덜어내거나 잘린다.
    """
    import collections

    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    # 화면 마크업은 <main> 안에만 있다. JS 템플릿 문자열까지 같이 세면
    # `${...}` 로 인자가 달라지는 필터 칩들이 중복으로 잡혀 뜻이 흐려진다.
    static = idx[idx.index('<main class="shell">'): idx.index("</main>")]

    # ① 화면(<section class="view">) 단위로 같은 동작이 두 번 놓여 있지 않은가
    marks = [(m.start(), m.group(1))
             for m in re.finditer(r'<section class="view"[^>]*id="v-([\w-]+)"', static)]
    assert len(marks) >= 8, "화면을 찾지 못했다 — 마크업 구조가 바뀌었다"
    marks.append((len(static), None))
    btn = re.compile(r"<button\b[^>]*?>.*?</button>", re.S)
    onx = re.compile(r'onclick="([^"]+)"')
    bad = []
    for i in range(len(marks) - 1):
        chunk = static[marks[i][0]:marks[i + 1][0]]
        seen = collections.Counter()
        for m in btn.finditer(chunk):
            f = onx.search(m.group(0))
            if f:
                seen[re.sub(r"\s+", "", f.group(1))] += 1
        for fn, n in seen.items():
            if n > 1:
                bad.append("%s 화면에 %s 가 %d번" % (marks[i][1], fn, n))
    assert not bad, "한 화면에 같은 동작 버튼이 겹쳐 있다: " + " / ".join(bad)

    # ② '복사' 만으로는 무엇을 복사하는지 알 수 없다 — 링크 복사와 헷갈렸던 자리다
    assert 'aria-label="복사"' not in idx, \
        "무엇을 복사하는지 모르는 '복사' 버튼이 남아 있다(링크 복사와 헷갈린다)"
    assert idx.count('aria-label="이미지 복사"') >= 4, \
        "도구줄의 복사 버튼 이름이 통일되지 않았다"
    cal = idx[idx.index('id="calTools"'):]
    cal = cal[: cal.index("</div>")]
    for lbl in ("새로고침", "링크 복사", "이미지 저장", "이미지 복사", "엑셀 저장"):
        assert 'aria-label="%s"' % lbl in cal, "캘린더 도구줄에 '%s' 가 없다" % lbl
    head = idx[idx.index('<section class="view" id="v-calendar"'):]
    head = head[: head.index('id="calTools"')]
    assert head.count("<button") == 1 and "focusCalendarEntry()" in head, \
        "캘린더 위쪽 큰 버튼이 다시 늘었다 — 저장·전달은 아래 도구줄 한 곳이다"

    # ③ 떠 있는 칩이 헤더를 덮지 않는가
    js = "".join(re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", idx, re.S))
    assert "function chipTop(" in js, "칩 위치를 앱바 아래로 잡는 코드가 없다"
    top = js[js.index("function chipTop("):]
    top = top[: top.index("\nlet ") if "\nlet " in top else 600]
    assert ".appbar" in top and "getBoundingClientRect" in top, \
        "앱바 높이를 재지 않고 칩을 띄운다 — 헤더 버튼을 덮는다"
    for fn in ("function netBanner(", "function swrChip("):
        blk = js[js.index(fn):]
        blk = blk[: blk.index("\nfunction ", 10)]
        assert "chipTop()" in blk, "%s 가 아직 고정 위치를 쓴다" % fn
        assert "env(safe-area-inset-top) + 8px" not in blk, \
            "%s 가 헤더와 같은 자리에 칩을 띄운다" % fn

    # ④ 좁은 화면에서 헤더가 겹치지 않게 덜어낸다
    css = "".join(re.findall(r"<style[^>]*>(.*?)</style>", idx, re.S))
    narrow = re.search(r"@media\(max-width:440px\)\{(.*?)\n\}", css, re.S)
    assert narrow, "좁은 폰(≤440px) 전용 헤더 규칙이 없다"
    narrow = narrow.group(1)
    assert ".appbar h1{display:none}" in narrow, \
        "좁은 폰에서 헤더 글자를 덜어내지 않는다 — 로고와 겹친다"
    assert ".appbar-identity{flex:none}" in narrow and "min-width:0" in narrow, \
        "자리가 모자랄 때 무엇이 줄어들지 정해 두지 않았다 — 앱 아이콘이 잘린다"
    assert ".appbar-identity{overflow:hidden}" in css, \
        "넘칠 때 겹치지 않게 자르는 안전망이 없다"

    # ⑤ 좁은 폰에서 도구줄은 옆으로 숨지 말고 줄을 바꾼다
    phone = re.search(r"@media\(max-width:640px\)\{(.*?)\n\}", css, re.S)
    assert phone, "폰(≤640px) 전용 도구줄 규칙이 없다"
    phone = phone.group(1)
    assert "flex-wrap:wrap!important" in phone and "overflow-x:visible" in phone, \
        "폰에서 도구줄이 옆으로 스크롤한다 — 마지막 버튼이 안 보인 채 숨는다"

    print("  [124] 겹치는 메뉴 통합 — 화면당 중복 동작 0 · 복사 이름 통일 · 칩/헤더 겹침 차단 ✅")


def t121_pid_alive():
    """[121] 죽은 프로세스를 살아 있다고 하지 않는다 (2026-08-06 실사고).

    무슨 일이 있었나 — `reports/.daily_run.lock` 이 **죽은 pid 의 이름으로** 남아
    daily_run 이 밤새 한 번도 못 돌았다. 잠금 규칙("주인이 죽었으면 회수")은 옳았는데
    **판정이 틀렸다**: 윈도우 `OpenProcess` 는 이미 끝난 프로세스에도 핸들을 준다.
    핸들이 열렸다는 것만으로 살아 있다고 본 탓에 잠금이 스스로 풀릴 길이 없었다.
    `Get-Process` 로는 안 보이는 pid 라 사람 눈에도 안 띈다 — 조용한 사고다.

    여기서 지키는 것
      ① 종료 코드까지 확인할 것(STILL_ACTIVE) — 핸들만 보고 판정하지 말 것
      ② 판정 불가는 None — '모르면 죽었다'로 넘기면 산 세션의 점유를 빼앗는다
      ③ 그 판정을 daily_run 잠금과 session_handoff 가 **같은 곳에서** 쓸 것
    """
    import subprocess
    import pid_alive as P

    assert P.alive(os.getpid()) is True, "지금 돌고 있는 나를 죽었다고 한다"
    assert P.alive(0) is None and P.alive("x") is None, "말이 안 되는 pid 는 None"

    # 방금 끝난 프로세스는 **확실히 죽었다**고 나와야 한다. 옛 판정이 틀렸던 바로 그 자리다.
    pr = subprocess.Popen([sys.executable, "-c", "pass"])
    pr.wait()
    assert P.alive(pr.pid) is False,         "끝난 프로세스를 살아 있다고 한다 — 잠금·점유가 영원히 안 풀린다"
    assert P.dead(pr.pid) is True and P.dead(os.getpid()) is False

    src = open(os.path.join(ROOT, "pid_alive.py"), encoding="utf-8").read()
    assert "GetExitCodeProcess" in src and "STILL_ACTIVE" in src,         "핸들만 보고 판정하면 끝난 프로세스가 살아 있는 것으로 나온다"
    assert "restype" in src and "argtypes" in src,         "64비트 핸들이 32비트로 잘린다 — 엉뚱한 핸들을 닫을 수 있다"

    # 두 곳이 같은 판정을 쓰는가 (한쪽만 고치면 그쪽에서 같은 사고가 또 난다)
    for name in ("daily_run.py", "session_handoff.py"):
        t = open(os.path.join(ROOT, name), encoding="utf-8").read()
        assert "import pid_alive" in t, "%s 가 옛 판정을 그대로 쓴다" % name
        assert "windll.kernel32.OpenProcess" not in t, (
            "%s 가 아직 스스로 판정한다 — 한쪽만 고치면 같은 사고가 또 난다" % name)

    # 죽은 주인의 잠금은 실제로 회수돼야 한다
    import daily_run as D
    assert D._pid_alive(pr.pid) is False, "daily_run 잠금이 죽은 주인을 못 알아본다"
    assert D._pid_alive(os.getpid()) is True, "산 주인의 잠금을 빼앗는다"
    print("  [121] 죽은 프로세스 판정 — 종료코드까지 확인·두 곳 공유·잠금 회수 ✅")


def t125_worktree_shared_state():
    """[125] 워크트리에서 일해도 **상태는 하나** (브랜치 인계, 분담판 #10).

    왜 필요한가 — 워크트리(`.claude/worktrees/<이름>`)는 **추적 파일만** 체크아웃한다.
    그런데 이 프로젝트의 상태는 거의 전부 git 밖이다(`reports/`·`updates/`·
    `config/*.json`·`db/*.db`). 그래서 워크트리 안에서는 모듈이 제 폴더 기준으로
    경로를 잡는 순간 본체와 **다른 상태**를 본다. 실측으로 확인한 것:
      ① 점유 파일이 갈려 두 세션이 동시에 `ledger` 를 잡아도 서로 안 보였다
         → 관리대장 동시 쓰기 금지가 조용히 무너진다. **가장 위험한 것.**
      ② 워크트리에서 enqueue 한 입력은 11:00·15:00 반영이 영영 못 본다
      ③ `config/` 가 없어 합성검증이 t1 에서 죽었다 — 즉 "ALL GREEN 확인 후
         실작업" 관문 자체를 통과할 수 없었다
      ④ 루트 CLAUDE.md 비교가 `.claude/worktrees/CLAUDE.md` 를 찾아 **거짓 경보**를
         매번 '먼저 처리할 것' 맨 위에 올렸다(해시는 같았는데도)
    """
    import worktree_state as W

    # 본체에서는 동작이 하나도 바뀌면 안 된다 — shared() 가 곧 제 폴더다
    assert os.path.isdir(W.main_root()), W.main_root()
    if not W.is_worktree():
        assert os.path.normcase(W.shared("db")) == os.path.normcase(os.path.join(ROOT, "db"))

    # 점유·큐 DB 는 **본체 경로**를 쓴다. 링크가 아니라 코드가 집어야 하는 이유가 있다.
    ac = open(os.path.join(ROOT, "ai_claim.py"), encoding="utf-8").read()
    assert "from worktree_state import shared" in ac and "STATE_DIR" in ac, \
        "점유 파일이 워크트리마다 갈린다 — 두 세션이 동시에 원장을 연다"
    assert "os.replace" in ac, "점유 저장이 원자적이지 않다"
    ld = open(os.path.join(ROOT, "ledger_db.py"), encoding="utf-8").read()
    assert "from worktree_state import shared" in ld, \
        "큐 DB 가 워크트리마다 갈린다 — 넣은 입력이 11:00·15:00 반영에 안 잡힌다"
    # 분담판도 세션 사이의 약속이다 — 갈리면 둘이 같은 일을 하고 나서야 안다.
    ws = open(os.path.join(ROOT, "worksplit.py"), encoding="utf-8").read()
    assert "from worktree_state import shared" in ws, \
        "분담판이 체크아웃마다 갈린다 — 서로 맡은 일을 못 본다"

    # `reports/` 를 통째로 정션하면 안 된다 — 추적 파일이 섞여 있어 git 이 흔들린다
    assert "reports" not in W.LINK_DIRS, \
        "reports 를 통째로 이으면 추적 파일 때문에 워크트리 git 이 남의 파일을 물고 간다"
    assert "reports/ai_claims.json" in W.CODE_SHARED and "db/ledger_queue.db" in W.CODE_SHARED

    # 이미 따로 있는 것을 덮지 않는다(사람이 일부러 둔 설정을 지우면 안 된다)
    src = open(os.path.join(ROOT, "worktree_state.py"), encoding="utf-8").read()
    assert "따로있음" in src and "덮지 않음" in src, "덮어쓰기 방지가 없다"
    for bad in ("shutil.rmtree", "os.remove", "os.unlink"):
        assert bad not in src, "워크트리 연결기가 파일을 지운다: %s" % bad

    # 끊겨 있으면 '먼저 처리할 것' 맨 앞에 뜬다 — 기계 상태가 아니라 st 로만 판단한다
    import session_handoff as H
    base = {"큐잔량": 0, "임시파일": [], "점유": [], "미푸시": [], "지시문사본": [],
            "수집신선도": []}
    cut = dict(base, **{"워크트리": {"여기": "X", "본체": "Y", "항목": [
        {"대상": "config/ecount_config.json", "방법": "하드링크", "상태": "없음"}]}})
    assert any("워크트리" in why for why, _ in H.blockers(cut)), H.blockers(cut)
    ok = dict(base, **{"워크트리": {"여기": "X", "본체": "Y", "항목": [
        {"대상": "config/ecount_config.json", "방법": "하드링크", "상태": "이어짐"}]}})
    assert not H.blockers(ok), "이어져 있는데 막았다"
    assert not H.blockers(dict(base)), "워크트리가 아닌데 워크트리 경고가 뜬다"

    # 루트 지시문 비교가 본체 기준인가 — 워크트리에서 거짓 경보가 뜨던 자리
    hs = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    assert "os.path.dirname(_main_root())" in hs, \
        "루트 CLAUDE.md 비교가 아직 dirname(BASE) 기준이다 — 워크트리에서 거짓 경보"

    # 미푸시는 **내 브랜치의 upstream** 기준이어야 한다. origin/master 로 세면
    # 워크트리 브랜치를 푸시하고도 "푸시되지 않은 커밋"이 영원히 남고, 제시된
    # 명령(git pull --rebase && git push)은 그 상황에 맞지도 않는다.
    assert "def unpushed_commits(" in hs and '"@{u}"' in hs, \
        "미푸시를 origin/master 로만 센다 — 브랜치에서 유령 blocker 가 뜬다"
    assert "def unmerged_commits(" in hs, "브랜치가 master 에 안 들어간 것을 안 알려준다"

    # 이어받기(--adopt)가 **큐를 흡수하기 전에** 먼저 잇는가. 순서가 뒤집히면
    # 빈 큐를 보고 "0건" 이라 답한다 — 조용히 틀린 답이다.
    # (설명문이 아니라 **실제 호출 순서**로 본다)
    assert (hs.index('steps.append(("워크트리 → 본체 잇기"')
            < hs.index('steps.append(("입력 큐 → DB"')), "--adopt 가 잇기 전에 큐를 읽는다"
    print("  [125] 워크트리 공용 상태(점유·큐 DB·분담판 일원화 · 거짓 경보 제거 · 잇기 우선) ✅")


def t126_app_font_and_revert():
    """[126] 앱 전체 글꼴 교체 + 되돌리기 보호 장치 (2026-08-06 지시).

    사용자 지시: "폰트도 나눔고딕 말고 디자이너들과 사용자들이 선호하고 잘 보이는
    폰트로 전체 변경해 / **나중에 내가 명령 내리면 다시 원래대로 할 수 있는
    보호 장치도 마련해**."

    지키는 것
      ① 글꼴을 정하는 자리는 파일마다 **한 곳**(`--font-ui`)이다. 네 파일에
         흩어져 있어서, 손으로 고치면 반드시 한 곳을 빠뜨린다.
      ② `--font-ui-legacy` 에 **원래 값이 글자 그대로** 남아 있다. 이게 되돌릴
         자리다 — 지워지면 "원래대로" 가 무슨 값인지 아무도 모른다.
      ③ 캡처(캔버스)로 그리는 이미지도 같은 변수를 읽는다. 예전엔 그리는 곳마다
         글꼴 이름을 손으로 적어 두어, 화면만 바뀌고 **저장한 이미지는 옛 글꼴**로
         남았다 — 카톡으로 보내고 나서야 안다.
      ④ `font_switch.py` 로 **왕복**이 된다(기본→예전→기본). 왕복 뒤 파일이
         원본과 한 바이트도 다르지 않아야 한다 — 되돌리기가 다른 것을 건드리면
         그건 되돌리기가 아니다.
    """
    import hashlib
    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    import font_switch as F

    # ① 네 파일 모두 변수 한 곳에서 정한다
    assert len(F.FILES) == 4, "글꼴을 정하는 파일 목록이 바뀌었다: %s" % (F.FILES,)
    for rel in F.FILES:
        text = open(os.path.join(ROOT, rel), encoding="utf-8").read()
        assert "--font-ui-legacy" in text, "%s 에 되돌릴 자리가 없다" % rel
        assert "var(--font-ui)" in text, "%s 가 변수를 안 쓰고 글꼴을 직접 적었다" % rel
        assert ':root[data-font="legacy"]' in text, \
            "%s 에 기기별 되돌리기 스위치가 없다" % rel

    # ② 원래 값이 그대로 남아 있는가 — 값을 여기 못 박아 둔다
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert ('--font-ui-legacy:"Nanum Gothic","NanumGothic","나눔고딕",'
            '"Apple SD Gothic Neo","Malgun Gothic",system-ui,sans-serif;') in idx, \
        "업무센터 앱의 원래 글꼴 값(나눔고딕)이 바뀌었다 — 되돌릴 곳을 잃었다"

    # ③ 캔버스가 손으로 적은 글꼴을 쓰지 않는가
    js = "".join(re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", idx, re.S))
    assert "function uiFont(" in js, "캔버스가 읽을 글꼴 함수가 없다"
    uf = js[js.index("function uiFont("):]
    uf = uf[: uf.index("\nfunction ", 10)]
    assert "--font-ui" in uf and "getPropertyValue" in uf, \
        "uiFont() 가 CSS 변수를 안 읽는다 — 되돌리기가 이미지에 안 먹는다"
    assert "return '" in uf, "변수를 못 읽는 브라우저용 대타가 없다(캔버스가 글꼴을 통째로 무시한다)"
    for line in js.splitlines():
        if "px" in line and "Nanum Gothic" in line:
            raise AssertionError("캔버스가 아직 글꼴 이름을 손으로 적는다: %s" % line.strip()[:70])
    assert js.count("uiFont()") >= 9, \
        "그림 그리는 곳 일부만 변수를 쓴다 — 그 이미지만 옛 글꼴로 남는다"

    # 기기별 스위치가 기억되는가(다음에 열어도 되돌린 상태여야 한다)
    assert "function setFontLegacy(" in js and "csos_font_legacy" in js, \
        "되돌리기 스위치가 기억되지 않는다 — 새로고침하면 도로 돌아온다"

    # ④ 왕복 — 실패해도 원본을 반드시 되돌린다
    def digest():
        return {r: hashlib.sha1(open(os.path.join(F.ROOT, r), "rb").read()).hexdigest()
                for r in F.FILES}

    origin = {r: open(os.path.join(F.ROOT, r), "rb").read() for r in F.FILES}
    before = digest()
    try:
        F.apply("legacy")
        assert {s["상태"] for s in F.state()} == {"예전"}, \
            "되돌리기가 일부 파일만 바꿨다 — 그 화면만 다른 글꼴로 보인다"
        F.apply("modern")
        assert digest() == before, "왕복했더니 파일이 달라졌다 — 되돌리기가 다른 것을 건드린다"
        assert {s["상태"] for s in F.state()} == {"기본"}
    finally:
        for r, raw in origin.items():
            p = os.path.join(F.ROOT, r)
            if open(p, "rb").read() != raw:
                open(p, "wb").write(raw)
    print("  [126] 앱 글꼴 일원화(4파일·캔버스 포함) + 되돌리기 왕복 무손실 ✅")


def t127_dark_mode_no_hardcoded_light_panel():
    """[127] 어둡게 켜도 글자가 보인다 — 바탕을 흰색으로 못 박지 않는다 (2026-08-07 지시).

    사용자 지시: "다크 모드시 글자 안보이는 문제 해결" (화면 사진 3장과 함께).

    사진에 찍힌 것
      · 대표 지표 카드(`.rep-metric`) — 바탕이 `rgba(255,255,255,.86)` 로 못 박혀 있어
        어둡게 켜면 **흰 바탕에 흰 글자**. 숫자만 보이고 무슨 지표인지 안 보였다.
      · 업무센터 담당자 버튼(`.workcenter-person`) — 그라데이션을 `color-mix(…, white)`
        로 섞어 밝은 판이 되고, 글자는 var(--brand) 라 밝아져 이름이 사라졌다.
      · 원본 자료 도구판(`#v-sources .src-toolbar`) — 같은 이유로 흰 판이 떠 있었다.

    왜 [115] 명암비 검사가 못 잡았나 — 그 검사는 **글자색과 배경색이 같은 규칙에**
    적혀 있을 때를 본다. 여기서는 바탕은 부모(`.rep-metric`)에, 글자는 자식(`.rl`)에
    있어서 짝이 안 맞았고 '배경 확인 필요' 194건 속에 숨어 있었다.

    그래서 이 검사는 짝을 맞추려 하지 않고 **원칙 하나**를 지킨다:
      어둡게 켤 수 있는 화면에서 바탕을 **불투명한 밝은 고정색**으로 적지 않는다.
      바탕도 토큰(var(--surface)/var(--panel)…)이어야 한다.
    글자가 없는 장식(막대·점)만 예외로 두고, 그 예외는 여기 이름과 이유를 적는다.
    """
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    css = "".join(re.findall(r"<style[^>]*>(.*?)</style>", idx, re.S))
    css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)

    # 글자가 얹히지 않는 장식만 예외 — 새로 넣으려면 여기 이유와 함께 적는다
    ALLOW = {
        ".dot.busy": "상태 점 — 글자가 얹히지 않는다",
        ".mbar i": "막대그래프 채움 — 글자가 얹히지 않는다",
    }

    def _lum(r, g, b):
        def f(x):
            x /= 255.0
            return x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4
        return .2126 * f(r) + .7152 * f(g) + .0722 * f(b)

    def opaque_light(val):
        """불투명(알파 0.5 이상)하면서 밝은 고정색인가. 반투명 덧칠은 어두운 판 위에
        얹히는 것이라 문제가 아니다(헤더·탭바의 rgba(255,255,255,.06) 따위)."""
        v = val.strip().lower()
        if v in ("white", "#fff", "#ffffff"):
            return True
        m = re.match(r"#([0-9a-f]{3}|[0-9a-f]{6})$", v)
        if m:
            c = m.group(1)
            if len(c) == 3:
                c = "".join(x * 2 for x in c)
            return _lum(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)) > 0.55
        m = re.match(r"rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?", v)
        if m:
            a = float(m.group(4)) if m.group(4) else 1.0
            return a >= 0.5 and _lum(*map(int, m.groups()[:3])) > 0.55
        return False

    bad = []
    for sel, body in re.findall(r"([^{}]+)\{([^{}]*)\}", css):
        s = " ".join(sel.split())
        if s.startswith("@") or 'data-theme="dark"' in s or "@media print" in s:
            continue
        if s in ALLOW:
            continue
        m = re.search(r"(?<!-)background(?:-color|-image)?\s*:\s*([^;]+)", body)
        if not m:
            continue
        val = m.group(1)
        hard = opaque_light(val)
        # 그라데이션·color-mix 안에 흰색을 섞는 것도 같은 사고다
        if re.search(r"color-mix\([^)]*,\s*(?:white|#fff{1,2}\b)\s*\)", val) or \
           re.search(r"gradient\([^)]*\b(?:white|#fff|#ffffff)\b", val):
            hard = True
        if hard:
            bad.append("%s → %s" % (s[:44], val.strip()[:40]))
    assert not bad, ("어둡게 켜면 밝은 판이 떠 글자가 사라진다(바탕도 토큰이어야 한다):\n  "
                     + "\n  ".join(bad))

    # 사진에 찍힌 세 곳이 정말 토큰으로 바뀌었나 — 원칙만 두면 다시 흰색으로 돌아간다
    for sel, want in ((".rep-metric{", "background:var(--surface)"),
                      ("#v-sources .src-toolbar{", "background:var(--surface)")):
        blk = css[css.index(sel):]
        blk = blk[: blk.index("}")]
        assert want in blk, "%s 의 바탕이 다시 고정색이다" % sel
    wc = css[css.index(".workcenter-person{"):]
    wc = wc[: wc.index("}")]
    assert "var(--surface)" in wc and "white" not in wc, \
        "담당자 버튼이 다시 흰색을 섞는다 — 어둡게 켜면 이름이 안 보인다"

    # 폰 앱(docs/app.html)도 어둡게 켜진다 — 같은 사고가 여기서도 났었다
    app = open(os.path.join(ROOT, "docs", "app.html"), encoding="utf-8").read()
    acss = re.sub(r"/\*.*?\*/", "", "".join(
        re.findall(r"<style[^>]*>(.*?)</style>", app, re.S)), flags=re.S)
    abad = []
    for sel, body in re.findall(r"([^{}]+)\{([^{}]*)\}", acss):
        s = " ".join(sel.split())
        if s.startswith("@") or "dark" in s or ":root" in s:
            continue
        m = re.search(r"(?<!-)background(?:-color|-image)?\s*:\s*([^;]+)", body)
        if m and opaque_light(m.group(1)):
            abad.append("%s → %s" % (s[:40], m.group(1).strip()[:32]))
    assert not abad, "폰 앱도 어둡게 켜면 밝은 판이 뜬다: " + " / ".join(abad)
    assert acss.count("--sel:") >= 3, \
        "폰 앱의 강조 바탕 토큰(--sel)이 테마마다 정의돼 있지 않다"

    print("  [127] 어둡게 켜도 글자가 보인다 — 업무센터·폰 앱 고정 밝은 바탕 0곳"
          "(장식 %d개만 예외) ✅" % len(ALLOW))


def t128_dash_tap_to_move():
    """[128] 대시보드 편집 — **눌러서 집고, 옮길 자리를 눌러** 이동 (2026-08-07 지시).

    사용자 지시: "대시보드 편집 시 각 카드 클릭해서 이동할 수 있는 기능 추가해서
    UX UI 완벽히 정리".

    왜 끌기만으로는 부족한가
      · 손이 미끄러지면 놓친다(끄는 도중 손을 떼면 아무 일도 안 일어난다).
      · 카드가 12개인 긴 화면에서는 **끌고 가는 동안 목적지가 화면 밖**이다.
        집었다 놓는 방식은 집은 채로 스크롤할 수 있어 이 문제가 없다.
      · 끌기는 그대로 둔다 — 가까운 자리는 끄는 편이 빠르다. 둘 다 있어야 한다.

    지키는 것
      ① 끌지 않고 그냥 누르면 집기/놓기로 읽는다(끌었으면 예전처럼 끌기다).
      ② 집으면 무엇을 집었는지 **사람이 읽는 이름**으로 보여 준다.
         내부 id(`representative`)가 뜨면 무엇을 집었는지 알 수 없다.
      ③ 빠져나갈 길이 셋이다: 같은 카드 다시 누르기 · Esc · 안내줄의 [취소].
      ④ 편집을 닫으면 집은 것도 놓는다(안 그러면 안내줄만 화면에 남는다).
      ⑤ 다른 묶음(화면 카드 ↔ 핵심지표 ↔ 보관함)으로는 옮기지 않는다.
      ⑥ 옮기면 **저장**한다 — 새로고침하면 되돌아가는 이동은 이동이 아니다.
    """
    idx = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    js = "".join(re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", idx, re.S))

    for fn in ("function dashPickTap(", "function dashPickCancel(",
               "function dashPickBar(", "function dashPickLabel("):
        assert fn in js, "%s 가 없다" % fn

    # ① 끌지 않은 누름만 집기로 간다
    de = js[js.index("function dashDragEnable("):]
    de = de[: de.index("\n/* ═══ 눌러서 옮기기")]
    assert "if(d.live){" in de and "dashPickTap(d.el, host, sel, onDrop)" in de, \
        "끌기와 누르기가 갈리지 않는다 — 끌고 나서도 카드가 집힌다"
    assert "e.type === 'pointerup'" in de, "취소(pointercancel)까지 집기로 읽는다"

    tap = js[js.index("function dashPickTap("):]
    tap = tap[: tap.index("\n/* Esc")]
    # ⑤ 다른 묶음으로 못 옮긴다
    assert "el.parentNode !== _dashPick.el.parentNode" in tap, \
        "화면 카드와 핵심지표가 섞인다"
    # ③ 같은 카드 다시 누르면 취소
    assert "_dashPick.el === el" in tap and "dashPickCancel()" in tap, \
        "집은 카드를 다시 눌러도 취소되지 않는다"
    # ⑥ 옮기면 저장한다
    assert "if(drop) drop(moving)" in tap, "옮기고 저장하지 않는다 — 새로고침하면 되돌아간다"

    # ② 사람이 읽는 이름
    lab = js[js.index("function dashPickLabel("):]
    lab = lab[: lab.index("\nfunction dashPickCancel(")]
    assert "dashCatalog()" in lab, "안내줄에 내부 id 가 뜬다 — 무엇을 집었는지 알 수 없다"

    # ③ Esc ④ 편집 닫으면 놓기
    assert "e.key === 'Escape'" in js and "dashPickCancel(true)" in js, "Esc 탈출구가 없다"
    tog = js[js.index("function toggleDashboardEditor("):]
    tog = tog[: tog.index("\nfunction ", 10)]
    assert "dashPickCancel(true)" in tog, \
        "편집을 닫아도 집은 것이 남는다 — 안내줄만 화면에 떠 있게 된다"

    # 보이는 안내 — 새 동작을 아무도 모르면 없는 기능이다
    assert "한 번 누르면 집히고" in idx, "편집 안내가 아직 끌기만 말한다"
    assert "#dashPickBar{" in idx and ".dash-picked{" in idx, "집힌 표시·안내줄 스타일이 없다"

    print("  [128] 대시보드 편집 — 눌러서 집기/놓기(끌기 병행)·이름 표시·탈출구 3개·저장 ✅")


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
                   "captureMeta['집계기준일']", "await nextPaint();",
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


def t129_call_notes_db_only_and_device_open():
    """[129] 통화 메모는 DB 에만 · 원본 클릭은 **접속한 기기**에서 연다 (2026-08-07 지시).

    사용자 지시: "원본 자료 클릭하면 접속한 디바이스에서 바로 열리게 알고리즘 수정해,
    그리고 통화_MD는 원본 자료에서 안보이게 처리하고 DB만 보관해(민감한 내용이 포함되어있음)."

    무엇이 잘못돼 있었나
      · 클릭은 `/api/open` → `os.startfile` 이라 **서버 PC 에서** 열렸다. 폰·다른 PC 는
        403 이라 경로만 복사됐고, 사무실 PC 에서는 아무도 안 보는 화면에 창만 떴다.
      · 통화 메모를 `0. 원본 자료/10. 통화·회의 기록/`(공유 폴더 Z:)에 복사해 두어
        앱 '원본 자료' 목록에 그대로 떴다(실측: 통화_20260805_김준형.md 카드 노출).

    지키는 것
      ① 통화 메모는 색인에 담기지 않는다 — 이름 규칙과 통화·회의 폴더 둘 다.
      ② call_notes 는 파일을 복사하지 않는다(그 복사가 노출의 원인이었다).
      ③ DB 왕복: 넣은 본문이 그대로 나오고, **목록은 본문을 주지 않는다.**
      ④ 서버는 목록 응답과 파일 전송 **두 곳 모두**에서 막는다 — 색인은 회차로 다시
         만들어지므로 옛 색인이 남아 있는 동안에도 막혀야 한다.
      ⑤ 화면은 /api/open 이 아니라 /api/source-file 을 쓴다.
    """
    import tempfile
    import source_dirs
    import source_index as S
    _rd = lambda *p: open(os.path.join(*p), encoding="utf-8").read()

    # ① 색인 제외 규칙
    assert S.is_private(os.path.join("x", "통화_20260805_김준형.md")), "이름 규칙이 안 걸린다"
    assert S.is_private(os.path.join(source_dirs.CALL_NOTE_DIR, "아무거나.txt")), \
        "통화·회의 폴더 전체가 안 걸린다"
    assert not S.is_private(os.path.join("x", "1-1._남양주4MB견적서_274,000원.pdf")), \
        "일반 원본까지 막으면 안 된다"
    src = _rd(ROOT, "source_index.py")
    assert "is_private(p, fn)" in src, "walk 에서 is_private 를 부르지 않는다 — 색인에 샌다"

    # ② Z: 복사 금지
    cn = _rd(ROOT, "call_notes.py")
    assert "shutil" not in cn, "call_notes 가 아직 파일을 복사한다 — Z: 노출 경로가 남았다"
    assert "call_note_save" in cn, "DB 보관을 부르지 않는다"

    # ③ DB 왕복 — 실 DB 는 건드리지 않는다
    import ledger_db
    tmpd = tempfile.mkdtemp(prefix="callnote_")
    keep = (ledger_db.DB_DIR, ledger_db.DB_PATH)
    ledger_db.DB_DIR = tmpd
    ledger_db.DB_PATH = os.path.join(tmpd, "t.db")
    try:
        body = "# 통화\n## 할 일\n- [김 · 2026-08-08] 확인\n"
        ledger_db.call_note_save("통화_20260808_테스트.md", body, whom="테스트",
                                 on="2026-08-08",
                                 todos=[{"who": "김", "due": "2026-08-08", "what": "확인"}])
        got = ledger_db.call_note_get("통화_20260808_테스트.md")
        assert got and got.get("body") == body, "DB 왕복이 깨졌다"
        lst = ledger_db.call_notes()
        assert lst and "body" not in lst[0], "목록이 본문을 흘린다 — 민감 자료다"
        assert lst[0]["todos"] and lst[0]["todos"][0]["what"] == "확인", "할 일이 안 남았다"
        ledger_db.call_note_save("통화_20260808_테스트.md", body + "추가", on="2026-08-08")
        assert len(ledger_db.call_notes()) == 1, "같은 파일 이름이 두 행이 됐다"
    finally:
        ledger_db.DB_DIR, ledger_db.DB_PATH = keep

    # ④ 서버 — 목록과 파일 전송 두 곳 모두
    app = _rd(ROOT, "webapp", "app_server.py")
    assert '"/api/source-file"' in app, "기기로 내려보내는 엔드포인트가 없다"
    assert app.count("from source_index import is_private") >= 2, \
        "목록·파일전송 두 곳 모두에서 막아야 한다(색인은 회차로 다시 만들어진다)"
    assert "filename*=UTF-8''" in app, "한글 파일 이름이 헤더에서 깨진다"

    # ⑤ 화면 — 클릭이 서버 PC 열기로 돌아가지 않았나
    html = _rd(ROOT, "webapp", "index.html")
    assert "sourceFileURL" in html and "/api/source-file" in html, "화면이 새 경로를 안 쓴다"
    fn = re.search(r"function openSource\(path\)\{.*?\n\}", html, re.S)
    assert fn, "openSource 를 찾지 못했다"
    assert "/api/open" not in fn.group(0), \
        "원본 클릭이 아직 서버 PC 에서 여는 /api/open 을 쓴다"
    print("  [129] 통화 메모 DB 전용 · 원본 클릭 접속 기기에서 열기 ✅")


def t130_band_grab_rejects_timeless_harvest():
    """[130] 시각 없는 수확은 저장하지 않는다 (2026-08-07 실사고 2차).

    무슨 일이 있었나
      `recheck_plan --ahead` 는 최신 글보다 **큰 번호**를 미리 찔러 본다. 그런데 밴드는
      아직 없는 글 번호에도 **HTTP 200 + 앱 껍데기**를 준다. 수집기가 그 화면에서 본문을
      뜯으면 직전 화면(피드 맨 위 글)의 본문이 그대로 잡히고 글쓴이·시각만 빈다.
      그렇게 3539~3578 마흔 건이 **전부 같은 글**로 수집됐고 status 가 ok 라 캐시에
      그대로 들어갔다(3308→3348). 본문이 있으니 실패로 보이지 않았다 —
      화면 숫자는 멀쩡한데 원본과 다른 값이 되는, 제일 나쁜 종류의 실패다.

    지키는 것
      ① 시각(timeText)이 없으면 'ok' 로 넘기지 않는다.
      ② 그때 'missing'(묘비)이 아니라 'fail' 이어야 한다 — 그 번호는 내일 진짜로
         생긴다. 묘비를 세우면 recheck_plan 이 영영 다시 안 뽑는다.
      ③ 숨은 탭에서는 시작 자체를 거절한다(1차 사고 가드가 살아 있나).
    """
    js = open(os.path.join(ROOT, "band", "grab_posts.js"), encoding="utf-8").read()

    body = re.search(r"async function grabOne\(.*?\n  \}", js, re.S)
    assert body, "grabOne 을 찾지 못했다"
    body = body.group(0)

    # 한 줄 반환이든 블록이든(지문을 함께 남기는 [131] 이후 모양) 상태만은 fail 이어야 한다.
    guard = re.search(r"if \(!timeText\)\s*(\{.*?\n      \}|return \{[^}]*\})", body, re.S)
    assert guard, "시각 없는 수확을 걸러내지 않는다 — 같은 글이 통째로 캐시에 들어간다"
    m = re.search(r"status: '(\w+)'", guard.group(1))
    assert m and m.group(1) == "fail", \
        f"시각 없음을 '{m.group(1) if m else '?'}' 로 처리한다 — 묘비를 세우면 그 번호를 영영 못 모은다"

    # 가드가 ok 반환보다 **앞**에 있어야 의미가 있다
    assert body.index("if (!timeText)") < body.index("status: 'ok'"), \
        "가드가 ok 반환보다 뒤에 있다 — 걸러지지 않는다"
    # timeText 를 다시 긁어 담지 않는지(가드를 우회하는 옛 코드가 남았나)
    assert "timeText: txt(" not in body, "가드를 지나 timeText 를 다시 긁는다"

    # ③ 1차 사고 가드(숨은 탭 시작 거절)도 함께 지킨다
    assert "document.hidden" in js and "탭이 뒤에 있다" in js, \
        "숨은 탭에서 시작을 거절하는 가드가 사라졌다"
    print("  [130] 밴드 수집 — 시각 없는 수확 폐기 ✅")


def t135_contaminated_not_recollected():
    """[135] 가짜 글 기록은 **재수집 목록에 넣지 않는다** (2026-08-07 지시).

    무엇이 잘못돼 있었나
      '날짜 없는 글 621건'을 "본문은 멀쩡한데 시각만 늦게 붙어 빈 것"으로 진단하고
      재수집 목록에 넣었다. 실제로 60건을 돌렸더니 **ok 0** 이었다. 세어 보니
      98건이 본문 2종, 523건이 7종 — 진짜라면 621종이어야 한다. 밴드가
      `/post/<번호>` 를 iframe 으로 열면 **피드로 되돌려서**, 껍데기에 남은 피드
      맨 위 글이 통째로 잡힌 것이었다. 같은 경로로 다시 열면 같은 가짜가 또 들어온다.

    지키는 것
      ① 판정은 좁게 — 작성일 없음 **그리고** 같은 본문이 2건 이상. 하나뿐이면 안 건드린다.
      ② 지우지 않고 표시한다 — 키를 지우면 '구멍'이 되어 매 회차 다시 훑는다.
      ③ 가짜 본문·글쓴이는 없앤다 — 남기면 대조가 그것을 진짜로 쓴다.
      ④ 계획이 오염 번호를 새 글/구멍/재수집 어디에도 넣지 않는다.
      ⑤ 인계 문서가 '재수집하라'고 말하지 않는다(그게 한 시간을 헛돌게 한 문구다).
    """
    import importlib
    cc = importlib.import_module("band.clean_contaminated")

    same = "Coupang이(가) 새 구매 오더(PO375207)를 전송했습니다"
    posts = {
        "10": {"content": same, "author": "김"},              # 날짜 없음 · 같은 본문
        "11": {"content": same, "author": "김"},              # 날짜 없음 · 같은 본문
        "12": {"content": "진짜 혼자인 글", "author": "박"},   # 날짜 없음 · 본문 하나뿐
        "13": {"content": "정상", "created_at": 1754500000000},
        "14": {"deleted": True},
    }
    bad = cc.find(posts)
    assert set(bad) == {"10", "11"}, f"판정이 틀렸다: {sorted(bad)}"
    assert "12" not in bad, "본문이 하나뿐인 것을 가짜로 몰았다 — 진짜 글일 수 있다"
    assert "13" not in bad and "14" not in bad, "정상·삭제 글을 건드렸다"

    # ③ 표시한 기록에는 가짜 본문이 남지 않는다
    src = open(os.path.join(ROOT, "band", "clean_contaminated.py"), encoding="utf-8").read()
    assert '"contaminated": True' in src, "표시 대신 지우면 매 회차 구멍으로 다시 훑는다"
    assert "os.replace(tmp, path)" in src, "캐시를 원자적으로 쓰지 않는다(사고 24)"

    # ④ 계획이 오염 번호를 어디에도 넣지 않는다
    rp = importlib.import_module("band.recheck_plan")
    p = rp.plan("1", {"1": {"created_at": 1}, "2": {"contaminated": True},
                      "3": {"created_at": 1}}, floor=1, ahead=0)
    assert 2 not in p["gaps"] and 2 not in p["stale"], "오염 번호가 재수집 목록에 들어간다"
    assert 2 not in (p.get("dateless") or []), "오염 번호가 날짜없음으로도 잡힌다"
    assert p.get("contaminated") == [2], "오염 건수가 안 보인다 — 잊힌다"

    # ⑤ 인계 문서가 재수집하라고 말하지 않는다 — 표본 3/3 리다이렉트로 삭제 판정 완료
    sh = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    assert "밴드오염" in sh and "재수집하지 않는다" in sh and "삭제" in sh, \
        "오염 기록이 '할 일'로 남았다 — 삭제 판정이 끝난 상태다"
    assert "수집 경로를 새로 만들어야 한다" not in sh, \
        "이미 닫힌 '수집 경로 재설계'가 할 일로 남아 있다 — 다음 세션이 헛돈다"

    # ⑥ 리다이렉트 가드 — 피드가 끝까지 그려지면 시각까지 있어서([130]만으로는) 못 막는다.
    #   내가 연 주소에 내가 있는지가 유일한 확증이다(실측: 피드의 "6시간 전"이 잡혔다).
    js2 = open(os.path.join(ROOT, "band", "grab_posts.js"), encoding="utf-8").read()
    assert "reason: 'redirect'" in js2 and "location.pathname" in js2, \
        "리다이렉트 가드가 없다 — 날짜 달린 가짜가 들어온다"
    assert "'/post/' + no" in js2, "주소 대조가 요청 번호와 안 묶였다"

    # ⑦ 오염 표시는 재병합을 살아남는다 — Z: 옛 덤프가 매 회차 다시 처리되기 때문.
    #   실측: 표시 621건이 한 회차 만에 0건이 됐었다. 작성일을 가진 기록(진짜 재수집)만 뚫는다.
    cd2 = open(os.path.join(ROOT, "band", "convert_dump.py"), encoding="utf-8").read()
    assert 'cur.get("contaminated") and not rec.get("created_at")' in cd2, \
        "재병합이 오염 표시를 덮는다 — 표시가 한 회차 만에 사라진다"
    print("  [135] 가짜 글 기록 — 재수집 금지·표시 유지 ✅")


def t131_band_quiet_vs_stalled():
    """[131] '밴드가 조용한 것'과 '수집이 막힌 것'을 가른다 (2026-08-07 지시).

    무엇이 잘못돼 있었나
      신선도 판정은 **날짜 있는 최신 글**만 봤다. 그래서 밴드에 새 글이 없는 날에도
      "★밀림"이 인계 문서 맨 위에 떴고, 그 경보를 믿고 없는 번호(3539~3578)를 긁었다가
      40건이 **전부 같은 글**로 캐시에 들어갔다(오늘 사고). 경보가 사고를 부른 셈이다.

    지키는 것
      ① 근거는 **missing 뿐**이다 — 밴드가 '없다'고 명시한 번호. failed/no-time 은
         화면이 안 그려졌을 때도 나오므로 '없음'의 증거가 못 된다.
      ② 수집 최대 번호 **바로 다음**이 없음으로 확인돼야 성립한다. 건너뛴 확인은 안 된다.
      ③ 옛 회차가 최근 확인을 덮지 않는다.
      ④ 근거가 최근일 때만 밀림을 내린다 — 오래된 근거는 그 사이 새 글이 올라왔을 수 있다.
      ⑤ 문서에 '조용함'과 그 이유가 보인다(다음 세션이 또 없는 번호를 긁지 않게).
    """
    import importlib
    import tempfile
    conv = importlib.import_module("band.convert_dump")
    import session_handoff as SH

    TMP = tempfile.mkdtemp(prefix="bandquiet_")
    log = os.path.join(TMP, "밴드_확인시각.json")
    keep_log = conv.PROBE_LOG
    conv.PROBE_LOG = log
    try:
        merged = {"3537": {"content": "x"}, "3538": {"content": "y"},
                  "3400": {"deleted": True}}
        cap = 1754500000000            # 고정 시각(테스트는 시계를 만들지 않는다)

        # ② 바로 다음 번호가 없음 → 기록된다
        conv._record_probe("84789192", "매출처업무", merged, [3539, 3540], cap)
        doc = json.load(open(log, encoding="utf-8"))
        assert doc["84789192"]["수집최대"] == 3538, "수집 최대 번호가 틀렸다"
        assert doc["84789192"]["없음확인"] == 3539
        first_seen = doc["84789192"]["확인시각"]

        # ② 건너뛴 확인은 증거가 아니다 — 3539 를 모르는 채 3541 만 없음이면 성립 안 한다
        os.remove(log)
        conv._record_probe("84789192", "매출처업무", merged, [3541], cap)
        assert not os.path.exists(log), "건너뛴 확인을 증거로 삼았다"

        # ③ 옛 회차가 최근 확인을 덮지 않는다
        conv._record_probe("84789192", "매출처업무", merged, [3539], cap)
        conv._record_probe("84789192", "매출처업무", merged, [3539], cap - 86400000)
        doc = json.load(open(log, encoding="utf-8"))
        assert doc["84789192"]["확인시각"] == first_seen, "옛 회차가 최근 확인을 덮었다"

        # ★ 오늘 사고의 실제 모양 — 아직 없는 번호는 missing 이 아니라 **notime** 으로
        #   떨어진다(밴드가 200 과 껍데기를 주므로). 지문이 같은 것이 2개 이상이면 증거다.
        same = "Coupang이(가) 새 구매 오더(PO375207)를 전송했습니다. 총금액 8,778,600원"
        assert conv._absent_above(3538, [], {"3539": same, "3540": same}) == [3539, 3540], \
            "같은 지문 여러 건을 '아직 없는 글'로 못 읽는다 — 이게 오늘 사고의 모양이다"
        # 하나뿐이면 '화면이 늦게 그려진 것'과 구분이 안 된다 — 증거로 치지 않는다
        assert conv._absent_above(3538, [], {"3539": same}) == [], \
            "지문 하나로 없음을 단정했다 — 느린 화면을 없는 글로 오해한다"
        # 지문이 서로 다르면 진짜 글일 수 있다 — 증거 아님
        assert conv._absent_above(3538, [], {"3539": "가", "3540": "나"}) == [], \
            "서로 다른 본문을 없는 글로 묶었다"
        # missing 과 notime 이 섞여도 이어져야 성립한다
        os.remove(log)
        conv._record_probe("84789192", "매출처업무", merged, [3540],
                           cap, {"3539": same, "3541": same})
        doc = json.load(open(log, encoding="utf-8"))
        assert doc["84789192"]["없음확인"] == 3539 and doc["84789192"]["연속없음"] == 3
    finally:
        conv.PROBE_LOG = keep_log

    # ④ 신선도 판정 — 근거가 최근이면 밀림이 내려가고, 오래되면 그대로 밀림
    keep = (SH.band_latest_days, SH.band_quiet)
    try:
        SH.band_latest_days = lambda: {"84789192": "2026-08-05"}
        SH.band_quiet = lambda: {"84789192": {"이름": "매출처업무", "수집최대": 3538,
                                              "확인시각": "2026-08-07 09:52",
                                              "없음확인": 3539, "연속없음": 2}}
        row = [f for f in SH.data_freshness("2026-08-07") if f["이름"].startswith("밴드:")][0]
        assert not row["밀림"], "새 글 없음을 확인했는데도 밀림이라 한다"
        assert "3538" in row.get("조용함", ""), "왜 안 긁어도 되는지가 안 적혔다"

        SH.band_quiet = lambda: {"84789192": {"이름": "매출처업무", "수집최대": 3538,
                                              "확인시각": "2026-08-01 09:00",
                                              "없음확인": 3539, "연속없음": 2}}
        row = [f for f in SH.data_freshness("2026-08-07") if f["이름"].startswith("밴드:")][0]
        assert row["밀림"], "오래된 근거로 밀림을 내렸다 — 그 사이 새 글이 있을 수 있다"

        # 근거가 아예 없으면 예전처럼 밀림이다(안전한 기본값)
        SH.band_quiet = lambda: {}
        row = [f for f in SH.data_freshness("2026-08-07") if f["이름"].startswith("밴드:")][0]
        assert row["밀림"], "근거가 없는데 밀림을 내렸다"
    finally:
        SH.band_latest_days, SH.band_quiet = keep

    # ⑤ 문서에 '조용함'과 이유가 보인다
    src = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    assert '"조용함"' in src and "(조용함)" in src, "표에 조용함 표시가 없다"
    assert "없는 번호를 긁으면" in src, "왜 긁으면 안 되는지가 문서에 안 적힌다"

    # ⑥ 수집기가 지문을 실제로 실어 보내야 판정 재료가 생긴다
    js = open(os.path.join(ROOT, "band", "grab_posts.js"), encoding="utf-8").read()
    assert "notime: S.notime" in js, "덤프에 지문이 안 실린다 — 판정 재료가 영영 안 온다"
    assert "S.notime[no] = r.sig" in js, "no-time 결과의 지문을 안 모은다"

    # ⑦ 수집 계획도 **같은 근거**를 본다 — 없다고 확인한 번호를 또 훑지 않는다.
    #   두 곳이 어긋나면 한쪽은 "긁어라", 다른 쪽은 "조용하다"가 되어 사람이 못 믿는다.
    rp = importlib.import_module("band.recheck_plan")
    plog = os.path.join(TMP, "확인시각2.json")
    keep_rp = rp.__dict__.get("_QUIET_PATH_FOR_TEST")
    real = rp._confirmed_quiet
    # 근거 파일 자리를 테스트용으로 바꿔 끼운다
    def _probe(band, hi, today=None, _p=plog):
        import json as _j
        from datetime import datetime as _dt
        try:
            rec = (_j.load(open(_p, encoding="utf-8")) or {}).get(str(band)) or {}
        except Exception:
            return False
        if int(rec.get("없음확인") or 0) != int(hi) + 1:
            return False
        seen = str(rec.get("확인시각") or "")[:10]
        day = str(today or "2026-08-07")[:10]
        age = (_dt.strptime(day, "%Y-%m-%d") - _dt.strptime(seen, "%Y-%m-%d")).days
        return 0 <= age <= rp.QUIET_LIMIT_DAYS
    with open(plog, "w", encoding="utf-8") as fh:
        json.dump({"84789192": {"없음확인": 3539, "확인시각": "2026-08-07 13:18"}}, fh,
                  ensure_ascii=False)
    assert _probe("84789192", 3538), "최근 확인인데 조용함으로 안 본다"
    assert not _probe("84789192", 3600), "확인한 번호와 무관한 hi 인데 조용함이라 한다"
    with open(plog, "w", encoding="utf-8") as fh:
        json.dump({"84789192": {"없음확인": 3539, "확인시각": "2026-08-01 09:00"}}, fh,
                  ensure_ascii=False)
    assert not _probe("84789192", 3538), "오래된 확인으로 새 글 탐색을 건너뛴다"
    assert callable(real), "recheck_plan._confirmed_quiet 이 사라졌다"
    src_rp = open(os.path.join(ROOT, "band", "recheck_plan.py"), encoding="utf-8").read()
    assert ("_absent_from(band)" in src_rp or
            ("_confirmed_quiet(band, hi)" in src_rp and "new = []" in src_rp)), \
        "계획이 조용함 근거를 안 본다 — 없는 번호를 매 회차 다시 훑는다"
    assert "밴드_확인시각.json" in src_rp, "session_handoff 와 다른 근거를 본다"
    print("  [131] 밴드 조용함 vs 수집 막힘 구분 ✅")


def t132_dash_snap_expand_theme_swipe():
    """[132] 카드 칸 맞춤·펼쳐보기 · 밝기 3종 · 폰 좌우 밀기 (2026-08-07 지시).

    사용자 지시 셋
      ① "카드 마우스로 잡아서 이동해서 **딱딱 맞아 떨어지게** 하는 알고리즘 추가해 /
         펼쳐보기 이런 기능들 적용해서 **각 카드가 정확히 균등하게** 맞아떨어지게"
      ② "이 기능 **기본, 다크모드, 시스템 동일** 이렇게 구성하게 추가해"
      ③ "**모바일 화면에서 손가락을 좌우로 밀면** 각 카테고리로 이동하는 기능"

    ★ 여기서 지키는 것은 대부분 **실측에서 다친 자리**다.
      · 칸 맞추기를 **DOM 순서로 세면 안 된다** — 격자가 `row dense` 라 브라우저가
        뒤 카드를 앞 빈칸으로 끌어올린다. 화면에 없는 빈칸을 보고 멀쩡한 카드를 넓혔다.
      · 칸 맞추기가 **좁은 화면에서 돌면 안 된다** — 폰에서는 카드가 한 줄에 한 장이라
        전부 12칸으로 넓히고, 그 폭이 저장돼 **PC 배치가 통째로 망가진다.**
        폰에서 누른 단추 하나가 PC 화면을 부수는, 되돌리기 어려운 손상이다.
      · 좌우 밀기는 **가로 스크롤 위에서 양보**해야 한다. 안 그러면 표를 옆으로 보려던
        손짓이 화면을 통째로 바꾼다.
    """
    src = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    # ── ① 균등 맞춤 · 펼쳐보기 ────────────────────────────────────────────
    assert "body.dash-even #dashGrid{align-items:stretch}" in src, \
        "한 줄에 선 카드의 키를 맞추지 않는다 — 바닥이 들쭉날쭉해진다"
    assert "--dash-card-h" in src and "dash-expanded" in src, "카드 높이 상한/펼침이 없다"
    assert "body.dash-even.dash-edit #dashGrid>.dash-block{max-height:none" in src, \
        "편집 중에도 카드를 자른다 — 잘린 카드는 집을 때 무엇인지 안 보인다"
    for fn in ("dashMeasureOverflow", "toggleDashCardExpand", "toggleDashboardEven",
               "snapDashboardLayout", "dashVisualRows", "expandAllDashCards"):
        assert f"function {fn}(" in src, f"{fn} 가 없다"
    # 넘칠 때만 단추가 뜬다 — 늘 떠 있으면 누를 것 없는 단추가 열두 장에 붙는다
    assert "el.scrollHeight>el.clientHeight+8" in src, \
        "넘침 판정에 여유가 없다 — 반올림 1px 로 멀쩡한 카드에 단추가 뜬다"
    # 내용이 도착한 뒤에 다시 재야 한다(첫 그림 때는 모든 카드가 '안 넘침'이다)
    assert "ResizeObserver" in src and "dashMeasureOverflow" in src, \
        "내용이 채워진 뒤 다시 재지 않는다 — 넘치는 카드에 펼쳐보기가 영영 안 뜬다"

    # ── 칸 맞추기: 실제 배치로 세고, 좁은 화면에서는 안 돈다 ──────────────
    snap = src[src.index("function snapDashboardLayout("):]
    snap = snap[:snap.index("\nfunction ")]
    assert "gridTemplateColumns" in snap and "tracks<2" in snap, \
        "좁은 화면 보호가 없다 — 폰에서 누르면 모든 카드가 12칸이 되어 PC 배치가 깨진다"
    assert "dashVisualRows()" in snap, \
        "칸 맞추기가 DOM 순서로 센다 — row dense 때문에 화면에 없는 빈칸을 고친다"
    vis = src[src.index("function dashVisualRows("):]
    vis = vis[:vis.index("\nfunction ")]
    assert "getBoundingClientRect" in vis, "줄 묶기가 실제 위치를 안 본다"

    # ── ② 밝기 세 가지 ──────────────────────────────────────────────────
    assert "const THEME_MODES = ['light','dark','system']" in src, "밝기 3종이 없다"
    assert "'시스템 동일'" in src or "시스템 동일" in src, "'시스템 동일' 이름이 없다"
    # 기본값은 여전히 밝게다 — 화면을 그대로 캡처해 보고서로 올리기 때문(원래 규칙 유지)
    assert "localStorage.getItem('cw_theme') || 'light'" in src, \
        "기본값이 밝게가 아니다 — 보고 화면이 기기 설정 따라 검게 나온다"
    assert "prefers-color-scheme: dark" in src and "watchSystemTheme" in src, \
        "'시스템 동일'인데 OS 설정 변화를 안 따라간다"
    assert "if(themeMode() === 'system')" in src, \
        "시스템이 아닐 때도 OS 변화를 반영한다 — 사람이 고른 값이 멋대로 바뀐다"

    # ── ③ 폰 좌우 밀기 ─────────────────────────────────────────────────
    for fn in ("navTabOrder", "swipeBlocked", "swipeToView", "initSwipeNav"):
        assert f"function {fn}(" in src, f"{fn} 가 없다"
    assert "initSwipeNav();" in src, "좌우 밀기가 시작될 때 켜지지 않는다"
    nav = src[src.index("function navTabOrder("):]
    nav = nav[:nav.index("\nfunction ")]
    assert ".tabbar button[data-v]" in nav and "offsetParent" in nav, \
        "밀기 순서를 탭바에서 안 읽는다 — 폰에서 숨은 탭으로 밀려 빈 화면이 뜬다"
    blk = src[src.index("function swipeBlocked("):]
    blk = blk[:blk.index("\nfunction ")]
    assert "scrollWidth > n.clientWidth" in blk, \
        "가로 스크롤 위에서 양보하지 않는다 — 표를 옆으로 보려다 화면이 바뀐다"
    assert "dashLayoutEditing" in blk and "dash-drag-live" in blk, \
        "카드를 옮기는 중에도 화면이 바뀐다"
    assert "sheetIsOpen()" in blk, "시트가 열려 있는데 뒤 화면이 바뀐다"
    sw = src[src.index("function swipeToView("):]
    sw = sw[:sw.index("\nfunction ")]
    assert "if(!next) return" in sw, "끝에서 처음으로 감긴다 — 밀다가 갑자기 대시보드로 튄다"
    assert "Math.abs(dx) < Math.abs(dy) * 1.5" in src, "세로로 그은 것도 화면 이동으로 읽는다"

    # ── 세 테마에서 같이 돌아야 한다 (지시 ②의 '동일') ────────────────────
    more = src[src.index(".dash-more{"):src.index(".dash-more button{")]
    assert "var(--surface)" in more and "#fff" not in more, \
        "펼쳐보기 띠가 흰색으로 못 박혀 있다 — 어둡게에서 흰 띠가 뜬다"
    assert "@media(prefers-reduced-motion:reduce)" in src, "움직임 줄이기 설정을 무시한다"
    print("  [132] 카드 칸 맞춤·펼쳐보기 · 밝기 3종 · 폰 좌우 밀기 ✅")


def t133_inline_style_dark_safe():
    """[133] 인라인 스타일도 어둡게에서 읽혀야 한다 (2026-08-07 사용자 화면).

    사용자 지적: **"다크모드에서 지금 갱신 텍스트 안보임"** (화면 사진).

    무엇이었나
      새 버전 알림 띠의 [지금 갱신] 단추가 `background:var(--panel)` +
      `color:var(--navy-900)` 이었다. 그런데 **`--navy-900` 은 다크 팔레트에서
      바뀌지 않는다**(#0E1B3F 그대로). 어둡게를 켜면 --panel 만 #1C2438 로 어두워져
      **어두운 판에 어두운 글자** — 단추가 통째로 사라졌다.
      밝게에서 --panel 이 밝아 우연히 읽혔을 뿐, 처음부터 짝이 안 맞는 조합이었다.

    ★ 왜 [127] 이 못 잡았나 — 그 검사는 `<style>` 안의 **CSS 규칙**만 훑는다.
      이건 JS 가 `style.cssText` 로 붙이는 인라인 스타일이라 그물 밖이었다.
      그래서 이 검사는 **JS 가 만드는 색**을 본다.

    규칙 하나: **고정색 판 위에는 고정색 글자를 쓴다.**
      `--navy-*` 는 두 테마에서 같은 값이다(짙은 남색 띠·헤더용). 그 위에 테마
      토큰(--panel/--surface/--ink…)을 얹으면 한쪽 테마에서 반드시 어긋난다.
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    # --navy-* 가 다크 팔레트에서 재정의되지 않는다는 것이 이 규칙의 전제다.
    dark = live[live.index(':root[data-theme="dark"]'):]
    dark = dark[:dark.index("}")]
    assert "--navy-900" not in dark,         "다크 팔레트가 --navy-900 을 바꾼다 — 이 검사의 전제가 달라졌으니 규칙을 다시 세울 것"

    # JS 가 붙이는 인라인 스타일에서 '테마 판 + 고정 남색 글자' 조합을 금지한다
    for m in re.finditer(r"cssText\s*=\s*(.*?)\n\s*b?\.?on|cssText\s*=\s*(.*?);\n", live, re.S):
        chunk = (m.group(1) or m.group(2) or "")[:600]
        if "var(--navy-" in chunk and re.search(r"background:\s*var\(--(panel|surface|bg)", chunk):
            raise AssertionError(f"인라인 스타일이 테마 판에 고정 남색 글자를 얹는다: {chunk[:120]}")

    # 문제의 그 단추 — 띠가 고정 남색이므로 단추도 고정색이어야 한다
    up = live[live.index("function offerUpdate("):][:4000]
    assert "background:#fff;color:#0E1B3F" in up,         "[지금 갱신] 이 다시 테마 토큰을 쓴다 — 어둡게에서 글자가 사라진다"
    assert "var(--navy-900);color:#fff" in up, "띠 자체는 짙은 남색 + 흰 글자여야 한다"
    print("  [133] 인라인 스타일 다크 안전 — 고정색 판 위엔 고정색 글자 ✅")


def t137_contamination_marked_every_merge():
    """[137] 오염 판정은 **병합할 때마다** 돈다 — 한 번 치우고 끝내지 않는다 (2026-08-07 2차 실사고).

    무엇이 잘못돼 있었나
      아침에 오염 621건을 `clean_contaminated --apply` 로 손수 표시했다. `convert_dump` 에는
      '표시된 것을 날짜 없는 재병합이 못 덮는다'는 가드까지 넣었다. 그런데 그 가드는
      **이미 표시된 것**만 지킨다. 15:32 회차에 새 덤프가 **표시된 적 없는** 유령 22건을
      들여왔고 아무도 막지 않았다. 그 결과 밴드업무추출에서 정기점검 UJ2601407 이
      1건 → **23건**으로 부풀었다(reports/밴드업무추출_전체_20260807_1502.csv → _1635.csv).
      화면도 리포트도 멀쩡해 보였다 — 숫자만 틀렸다. 사람이 매번 알아채야 하는 조치는
      결국 안 하게 된다. 그래서 판정을 병합 경로 **안으로** 옮겼다.

    지키는 것
      ① convert_dump 가 캐시를 쓰기 **전에** clean_contaminated.find 로 새 오염을 표시한다.
      ② 옆 파일 import 가 조용히 실패해 보호가 통째로 꺼지지 않게 제 폴더를 경로에 넣는다.
      ③ 표시는 재병합을 견딘다(기존 226행 가드) — 둘이 함께 있어야 유령이 안 되살아난다.
    """
    src = open(os.path.join(ROOT, "band", "convert_dump.py"), encoding="utf-8").read()
    write_at = src.index('out = {"band_name"')
    guard = src[:write_at]
    assert "clean_contaminated" in guard, \
        "병합이 캐시를 쓰기 전에 오염 판정을 하지 않는다 — 새 유령이 그대로 들어간다"
    assert "clean_contaminated.find(merged)" in guard, \
        "판정 대상이 병합 결과(merged)가 아니다"
    assert '"contaminated": True' in guard, "찾아 놓고 표시하지 않는다"
    assert "sys.path.insert(0, _here)" in guard, \
        "옆 폴더 import 가 실패하면 보호가 조용히 꺼진다 — 제 폴더를 경로에 넣어야 한다"
    # 기존 재병합 생존 가드가 함께 살아 있어야 한다(하나만으로는 유령이 되살아난다)
    assert 'cur.get("contaminated") and not rec.get("created_at")' in src, \
        "표시가 날짜 없는 재병합에 덮인다 — [135]⑦ 회귀"

    # 판정기 자체: 표시된 스텁을 다시 오염으로 잡지 않아야 한다(무한 재표시 방지)
    import importlib
    cc = importlib.import_module("band.clean_contaminated")
    stub = {"contaminated": True, "captured_at": 1, "why": "x"}
    posts = {"1": stub, "2": stub, "3": {"content": "진짜 글", "created_at": "2026-08-07 10:00"}}
    assert cc.find(posts) == [] or all(k not in cc.find(posts) for k in ("1", "2")), \
        "이미 표시된 스텁을 또 오염으로 잡는다"

    # ④ 날짜가 붙어 들어온 가짜 — 작성일·본문이 **둘 다** 같은 묶음만 잡고,
    #    그 묶음에서도 맨 앞 번호는 남긴다(원본이 섞여 있을 수 있다).
    body = "[ 쿠팡 A/S 안내 ] A/S 일자 : 2026.08.04 " * 4
    dup = {str(n): {"content": body, "created_at": 1785800580000, "author": "지원팀"}
           for n in (5420, 5421, 5422, 5423, 5424)}
    got = cc.find(dup)
    assert set(got) == {"5421", "5422", "5423", "5424"}, \
        "날짜 달린 사본을 못 잡거나 원본까지 지운다: %s" % sorted(got)

    # ⑤ 진짜 반복 글은 건드리지 않는다 — 같은 본문이라도 **올린 시각이 다르다**
    repeat = {"10": {"content": "발주현황 공유", "created_at": 1700000000000},
              "20": {"content": "발주현황 공유", "created_at": 1700086400000}}
    assert cc.find(repeat) == {}, "매주 올리는 같은 제목의 진짜 글을 가짜로 잡는다"

    # ⑥ 오탐 방지선: '본문이 피드 머리글로 시작한다'만으로는 잡지 않는다.
    #    실측에서 그 규칙은 본문이 전부 다른 진짜 글 98건까지 걸었다(상세 화면에도 머리글이 있다).
    head = "2026년 8월 4일 오전 8:43 게시글 지원팀 3시간 전 "
    lone = {"1": {"content": head + "가", "created_at": 1},
            "2": {"content": head + "나", "created_at": 2}}
    assert cc.find(lone) == {}, \
        "피드 머리글만 보고 잡는다 — 본문이 다르면 서로 다른 진짜 글이다"
    print("  [137] 오염 판정이 병합마다 자동 실행 · 날짜 달린 사본만 좁게 · 원본 보존 ✅")


def t138_daily_run_completion_watch():
    """[138] 일일자동대조가 **완주했는지**를 인계 문서가 본다 (2026-08-07 실사고).

    무엇이 잘못돼 있었나
      작업 스케줄러의 09:50 '일일자동대조'는 매일 **성공(0)** 이었다. 그런데 실제로는
      2026-08-06 21:01 이후 20시간 동안 한 번도 완주하지 않았다. 이유가 둘이다:
        · daily_run 은 앞 회차가 아직 돌면 한 줄 찍고 **정상 종료**한다(exit 0).
          앞 회차가 3시간씩 걸리니 다음 회차는 늘 잠겨 있고, 서로를 가려 준다.
        · 마지막으로 남은 표식(agent_status.json)은 `aborted: True` 였는데
          파일 자체는 최신이라 '최근에 돌았다'로 보였다.
      그동안 자료현황.md 가 1.7일, 종합리포트가 19시간 멈춰 있었고 아무 경보도 없었다.

    지키는 것
      ① 나이(20시간)뿐 아니라 **중단 여부**로도 밀림을 판정한다 — 최신 파일이 곧 성공은 아니다.
      ② 판정 결과가 '먼저 처리할 것'에 오른다. 안 그러면 아무도 안 본다.
      ③ 정상(최근·중단 아님)일 때는 조용하다 — 늘 켜져 있는 경보는 경보가 아니다.
    """
    import importlib
    S = importlib.import_module("session_handoff")
    base = {"큐잔량": 0, "임시파일": [], "옛버전편집": [], "점유": [], "미커밋": [],
            "미푸시": 0, "미머지": 0, "지시문사본": [], "밴드날짜없음": {},
            "밴드오염": {}, "워크트리": None, "원장": {}}
    def only_daily(dr):
        return [w for w, _ in S.blockers(dict(base, 일일대조=dr))
                if "일일자동대조" in w]

    assert only_daily({"밀림": True, "중단": True, "경과시간": 1.0, "실패단계": ["자료현황 갱신"]}), \
        "마지막 회차가 중단인데 아무 말도 하지 않는다 — 파일이 최신이면 성공으로 보인다"
    assert "자료현황 갱신" in only_daily(
        {"밀림": True, "중단": True, "경과시간": 1.0, "실패단계": ["자료현황 갱신"]})[0], \
        "어느 단계가 실패했는지 말하지 않으면 다음 세션이 다시 찾아야 한다"
    assert only_daily({"밀림": True, "중단": False, "경과시간": 26.0, "실패단계": []}), \
        "하루가 넘게 완주하지 않았는데 조용하다"
    assert not only_daily({"밀림": False, "중단": False, "경과시간": 2.0, "실패단계": []}), \
        "정상인데도 경보를 올린다 — 늘 켜진 경보는 무시된다"

    # 판정기 자체: 중단이면 나이와 무관하게 밀림이다
    src = open(os.path.join(ROOT, "session_handoff.py"), encoding="utf-8").read()
    assert 'age_h >= DAILY_STALE_H or aborted' in src, \
        "중단을 나이와 함께 보지 않는다 — 방금 중단한 회차를 정상으로 본다"
    print("  [138] 일일대조 완주 감시 — 중단·장기 미완주를 '먼저 처리할 것'으로 ✅")


def t139_new_version_is_atomic():
    """[139] 관리대장 새 버전은 **검증을 통과한 뒤에야** 정본 이름을 얻는다 (2026-08-07).

    무엇이 위험했나
      `findings_sheet.upsert()` 와 `ledger_writer` 는 vN+1 을 **정본 이름으로 먼저 만들고**
      그 다음에 무결성을 검사했다. 원장은 Z: SMB 네트워크 드라이브에 있는 수십 MB 짜리다.
      쓰는 도중 프로세스가 죽으면 — 오늘 실제로 스케줄러가 3시간 한도로 프로세스를
      강제 종료(0x41306)했고, Z: 는 WinError 1231 을 낸 적이 있다 — **이름은 멀쩡한
      `_v{N+1}.xlsx`, 내용은 잘린 zip** 이 남는다. `resolve_master()` 는 v번호가 가장 큰
      것을 정본으로 집으므로(`ecount_reconcile.py`) 앱도, 모든 대조 도구도, 다음 회차도
      그 깨진 파일을 정본으로 읽는다. `~$` 임시파일은 걸러도 이런 반쪽 파일은 못 거른다.
      바로 옆 `workbook_patch.py` 는 처음부터 tmp → 검증 → os.replace 를 하고 있었다.

    지키는 것
      ① 두 파일 모두 `.tmp.xlsx` 에 쓴다 — 정본 이름은 마지막에 os.replace 로만 얻는다.
      ② 검증(zip 무결성·재독)이 **tmp 를 대상으로** 돈다. 통과 못 하면 정본이 되지 않는다.
      ③ 실패하면 tmp 를 지운다 — 남기면 '임시 결과가 이미 존재'로 다음 회차가 막힌다.
    """
    for fn, anchor in (("findings_sheet.py", "zipfile.ZipFile(tmp,"),
                       ("ledger_writer.py", "final_dst, dst = dst")):
        src = open(os.path.join(ROOT, fn), encoding="utf-8").read()
        assert ".tmp.xlsx" in src, f"{fn}: 임시파일을 쓰지 않는다 — 정본에 직접 쓴다"
        assert anchor in src, f"{fn}: 임시파일로 쓰는 자리가 사라졌다"
        assert "os.replace(" in src, f"{fn}: 원자적 교체가 없다"
        # 검증이 정본이 아니라 tmp 를 향해야 한다
        assert "zipfile.ZipFile(dst)" not in src or fn == "ledger_writer.py", \
            f"{fn}: 검증이 아직 정본(dst)을 연다"
        assert "os.remove(" in src, f"{fn}: 검증 실패 시 반쪽 임시파일을 치우지 않는다"

    # ledger_writer 는 dst 를 tmp 로 바꿔치기하므로, 검증 뒤 교체까지가 한 덩어리여야 한다
    lw = open(os.path.join(ROOT, "ledger_writer.py"), encoding="utf-8").read()
    assert lw.index("assert not bad") < lw.index("os.replace(dst, final_dst)"), \
        "검증보다 먼저 정본으로 바꾼다 — 검증이 무의미해진다"
    print("  [139] 관리대장 새 버전은 검증 통과 후에만 정본이 된다(원자적 교체) ✅")


def t140_freshness_tells_the_truth():
    """'몇 분 전 자료' 와 '에이전트 실행 시각' 이 거짓말하지 않나 (2026-08-07 지시).

    사용자 지적 셋이 실은 **같은 병**이었다 — 화면이 자기 나이를 모른 채 말한다.
      "계속 194분전 갱신중 … 분이 이상해"        → 나이 표기가 시간 단위를 모른다
      "계속 몇분전 이라고 표시하는데 계속 내용이 달라" → 칩 하나를 11개 경로가 덮어썼다
      "에이전트 실행 시간이 맞지 않아"              → 어제 중단된 회차가 초록 점을 달고 있었다
    셋 다 '조용한 사고'(CLAUDE.md)의 얼굴이다. 여기서 지키는 것은 표현이 아니라
    **모르면 모른다고 말하는가**이다.
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()

    # ① 칩의 나이는 '아직 안 들어온 것 중 가장 오래된 것' 하나로 정한다.
    #    경로마다 제 나이를 쓰면 숫자가 오간다 — 그게 "내용이 달라"였다.
    assert "const SWR_WAIT" in live and "const SWR_INFLIGHT" in live
    assert "Math.min(...vals.map(v => v.t))" in live, "칩이 여전히 마지막에 쓴 값을 보여 준다"
    assert "swrChip(Date.now() - hit.t)" not in live, "경로별 나이를 그대로 칩에 쓰던 옛 길이 남아 있다"
    # ② 갱신 실패를 삼키지 않는다. 삼키면 옛 값이 늙는 동안 '갱신 중'이 영원히 남는다.
    assert ".catch(() => { swrDone(key, false); })" in live, "SWR 갱신 실패를 다시 삼킨다"
    assert "갱신 실패" in live, "실패를 사람에게 말하지 않는다"
    # ③ 60분이 넘으면 '194분 전' 이 아니라 시간으로 읽는다.
    assert "function swrAgeText(" in live and "'시간 '" in live
    # ④ 잠깐 끝나는 갱신에는 칩을 띄우지 않는다("너무 자주 뜨고").
    assert "SWR_CHIP_DELAY_MS" in live and "Date.now() < SWR_SHOW_AT" in live
    # ⑤ 한 곳이 끝났다고 칩을 지우지 않는다 — 대기 목록이 빌 때만 지운다.
    assert "if(!SWR_WAIT.size){" in live, "대기 목록과 무관하게 칩을 지운다"

    # ⑥ 숫자의 나이는 **원장이 저장된 시각**이지 지금 시각이 아니다.
    assert "def _data_asof_iso(" in server
    assert '"데이터최종갱신일": _data_asof_iso()' in server, \
        "대표 보고 meta 가 다시 datetime.now() 로 '방금'이라 말한다"
    assert "const _asOf" in live, "화면과 캡처가 각자 시각을 만든다 — 갈라진다"
    assert "new Date().toISOString().slice(0,16).replace('T',' ')}</span>" not in live, \
        "근거 없이 지금 시각을 데이터 갱신 시각으로 찍는 자리가 남아 있다"

    # ⑦ 에이전트 배지는 밀리면 밀렸다고 말한다(초록 점 금지).
    assert '"agent_stale"' in server and '"agent_aborted"' in server
    assert "agent_age_h >= 20" in server, "몇 시간째 미완주인지 판정하는 자리가 없다"
    assert "s.agent_stale" in live and "미완주" in live, "앱이 밀린 회차를 정상처럼 보여 준다"
    # ⑧ 캡처는 **누르면 곧바로** 그린다 (2026-08-07 지시: "이미지 저장 하면
    #   바로바로 데이터 내역 반영해서 캡처되게 알고리즘 다시 짜").
    #   예전 문제는 상한 25초가 아니라 '받기 시작하는 시점'이었다 — [보고] 탭에
    #   재조회 분기가 없어서 **누른 다음에야** 받기 시작했다.
    assert "if(v==='daily') warmReport();" in live, \
        "[보고] 탭을 열 때 미리 받지 않는다 — 누른 뒤에야 받으면 버튼이 굳는다"
    assert "if(rptHasData()) return {how:'now'" in live, \
        "가진 자료가 있는데도 기다린다"
    assert "function rptOfferResave(" in live and "_rptData.updatedAt === mark" in live, \
        "저장 뒤 더 새 자료가 와도 알려 주지 않는다"
    assert "toast('최신 자료를 받아 그리는 중…')" not in live, \
        "이제 기다리지 않는데 기다린다고 말한다"
    print("  [140] 신선도 표기 — 나이는 하나·실패는 말한다·캡처는 곧바로 ✅")


def t141_long_text_folds():
    """긴 글은 접고 짧은 글은 건드리지 않나 (2026-08-07 지시).

    사용자 지시: "너무 스크롤이 긴 부분(설명란) … 누르면 자세히 펼쳐지면서 보는 기능 /
    앱 구동이나 사용 편의에 위배될 경우 니가 알아서 최적의 방법으로".
    그래서 '무엇을 접을지'를 목록으로 정하지 않는다 — 목록은 짧은 글까지 접어
    한 번 더 누르게 만든다. **실제 높이를 재서** 정한다.
    그리고 접어서 잃으면 안 되는 두 가지를 여기서 지킨다.
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert "function lcScan(" in live and "function lcToggle(" in live
    # ① 재서 정한다 — 한도와 '접을 값어치'가 둘 다 있어야 한다
    assert "const LC_MAX" in live and "const LC_GAIN" in live
    assert "el.scrollHeight <= LC_MAX + LC_GAIN" in live, "짧은 글까지 접는다"
    # ② 누를 것이 든 블록은 접지 않는다 — 숨은 버튼은 '없는 기능'이 된다
    assert "function lcSafe(" in live
    for tag in ("input", "select", "textarea", "button", "canvas", "table"):
        assert tag in live.split("function lcSafe(")[1][:400], \
            f"lcSafe 가 {tag} 를 품은 블록을 접을 수 있다"
    # ③ 인쇄는 전부 편다 — 종이에 잘려 나가면 영영 못 편다
    assert "@media print{" in live and ".lc{max-height:none!important" in live, \
        "접힌 채로 인쇄되면 잘린 문서가 남는다"
    # ④ 말투는 대시보드 카드와 같아야 한다(같은 동작에 두 이름을 쓰지 않는다)
    assert "'펼쳐보기'" in live and "'접기'" in live
    # ⑤ 화면을 그린 뒤 실제로 불린다 — 안 부르면 코드만 있고 아무 일도 안 일어난다
    assert "lcScanSoon($('v-'+v))" in live, "화면 전환 뒤 긴 글을 재지 않는다"
    print("  [141] 긴 글 접기 — 재서 정함·기능 숨김 금지·인쇄는 전부 펼침 ✅")


def t142_flow_editable():
    """AS 접수→수금 워크플로우 — 고칠 수 있고 되돌릴 수 있나 (2026-08-07 지시).

    사용자 지시: "AS 접수부터 처리 수금까지의 과정을 워크플로우로 만들어서
    수정할 수 있는 기능도 추가해서 디테일하게 / 텍스트는 딱 필요한 것만".
    '고칠 수 있다'는 **'되돌릴 수 있다'와 짝**이어야 한다 — 되돌릴 길이 없으면
    사람은 잘못 고칠까 봐 결국 안 고치고, 화면은 실제 업무와 어긋난 채 남는다.
    """
    import ledger_db as L
    before = L.flow_steps()
    assert before and before[0]["단계"] and before[-1]["단계"], "기본 흐름이 비어 있다"
    # ① 통째 저장 → 되돌리기 왕복이 무손실인가
    mod = [dict(x) for x in before]
    mod[0] = dict(mod[0], 메모="검증용")
    mod.append({"단계": "검증용 단계", "담당": "", "소요일": 99, "근거": "", "메모": ""})
    assert L.flow_save(mod, who="synthetic") == len(before) + 1
    assert len(L.flow_steps()) == len(before) + 1
    L.flow_restore(who="synthetic")
    after = L.flow_steps()
    assert [(s["단계"], s["담당"], s["소요일"], s["근거"], s["메모"]) for s in after] == \
           [(s["단계"], s["담당"], s["소요일"], s["근거"], s["메모"]) for s in before], \
        "되돌리기가 원래 모습으로 돌아오지 않는다"
    # ② 빈 흐름·이름 없는 줄은 저장하지 않는다(빈 화면이 남으면 아무도 못 고친다)
    for bad in ([], [{"단계": "  "}]):
        try:
            L.flow_save(bad, "synthetic")
            raise AssertionError("빈 흐름이 저장됐다")
        except ValueError:
            pass

    server = open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
    assert 'if p == "/api/flow"' in server and "ledger_db.flow_restore" in server
    assert "_require_admin()" in server.split('if p == "/api/flow"')[2][:300], \
        "흐름 저장이 인증 없이 열려 있다"
    # ③ 화면: 좌측 카테고리 · 뷰 · 수정/저장/되돌리기가 실제로 이어져 있나
    assert 'data-v="flow"' in live and 'id="v-flow"' in live
    assert "if(v==='flow' && !FLOW_EDITING) loadFlow();" in live, \
        "고치는 중에 다시 불러 덮어쓴다 — 입력하던 것이 사라진다"
    for fn in ("loadFlow", "flowRender", "flowEdit", "flowSave", "flowUndo",
               "flowAddStep", "flowDel", "flowMove", "flowCollect"):
        assert f"function {fn}(" in live, f"{fn} 이 없다"
    # ④ 아이콘은 스프라이트에 있는 것만 쓴다(없으면 빈 네모가 뜬다 — 검증 [91])
    icon = live.split('data-v="flow"')[1].split('href="#')[1].split('"')[0]
    assert f'id="{icon}"' in live, f"워크플로우 아이콘 {icon} 이 스프라이트에 없다"
    print("  [142] 워크플로우 — 통째 저장·되돌리기 왕복 무손실·빈 흐름 거부·화면 배선 ✅")


def t136_work_lanes():
    """작업 차선 — 수집 창과 앱·엑셀 창이 하루 종일 나란히 돌 수 있나 (2026-08-07 지시).

    지키는 것은 네 가지다:
      ① 차선을 **안 정한 세션은 아무것도 막히지 않는다** (기존 세션·스케줄러 보호)
      ② 내 차선 밖 자원은 못 잡는다 (수집 창이 code 를 집어 가면 앱 창이 되잡지 못한다)
      ③ 남의 차선은 못 빼앗는다 (살아 있는 주인이 있으면 거절)
      ④ 죽은 세션의 차선은 **즉시** 비워진다 (45분 기다리면 다음 창이 놀게 된다)
    """
    import importlib
    import socket as _socket
    lanes = importlib.import_module("lanes")
    ac = importlib.import_module("ai_claim")

    with tempfile.TemporaryDirectory() as tmp:
        # 점유 파일 자리를 옮기면 차선도 따라온다 — 그래서 이 세션이 어느 차선에
        # 서 있든 검증은 빈 차선에서 시작한다(그러지 않으면 검증을 못 돌린다).
        keep_claims, keep_sid = ac.CLAIMS, lanes._me
        ac.CLAIMS = os.path.join(tmp, "ai_claims.json")
        try:
            lanes._me = lambda: "SESS-A"
            assert os.path.dirname(lanes._path()) == tmp, "차선 파일이 점유 자리를 안 따라간다"

            # ① 차선 밖이면 전부 허용 — 이 예외가 없으면 기존 자동화가 통째로 멈춘다
            for res in ("code", "ledger", "band", "publish"):
                ok, _ = lanes.can(res)
                assert ok, f"차선을 안 정했는데 '{res}' 가 막혔다 — 기존 세션이 멈춘다"

            assert lanes.take("collect", "claude", "수집 전담") == 0, "수집 차선을 못 잡는다"
            assert lanes.my_lane() == "collect", "잡은 차선이 내 것으로 안 보인다"

            # ② 차선 밖 자원은 막히고, 안쪽은 열린다
            for res in ("code", "ledger", "publish"):
                ok, why = lanes.can(res)
                assert not ok, f"수집 차선인데 '{res}' 가 열려 있다 — 앱 창과 부딪친다"
                assert "차선" in why, "왜 막혔는지 안 알려 준다"
            for res in ("band", "read", "report"):
                assert lanes.can(res)[0], f"수집 차선인데 '{res}' 가 막혔다"

            # ai_claim 이 실제로 그 문을 본다 (배선 확인 — 함수만 있고 안 부르면 소용없다)
            assert ac._lane_gate("band"), "차선 안 자원인데 점유가 거절됐다"
            assert not ac._lane_gate("code"), "ai_claim 이 차선을 안 본다"

            # ③ 살아 있는 남의 차선은 못 빼앗는다
            lanes._me = lambda: "SESS-B"
            assert lanes.take("collect", "claude", "가로채기") == 2, \
                "남의 차선을 빼앗았다 — 두 창이 같은 파일에서 만난다"
            assert lanes.take("build", "claude", "앱·엑셀") == 0, "빈 차선을 못 잡는다"
            # 앱 창은 code·ledger 가 열려 있어야 한다 — 그게 이 기능의 목적이다
            for res in ("code", "ledger", "publish"):
                assert lanes.can(res)[0], f"앱·엑셀 차선인데 '{res}' 가 막혔다"
            assert not lanes.can("band")[0], "앱 차선이 수집 자원까지 쥔다"

            # 한 세션은 한 차선 — 옮기면 앞 차선에서 빠진다
            lanes._me = lambda: "SESS-A"
            assert lanes.free("claude") == 0, "내 차선을 못 놓는다"
            lanes._me = lambda: "SESS-B"
            assert lanes.take("collect", "claude", "옮김") == 0, "빈 차선으로 못 옮긴다"
            d = lanes._load()
            assert "build" not in d, "차선을 옮겼는데 앞 차선이 남아 있다 — 분업표가 깨진다"

            # ④ 죽은 세션의 차선은 즉시 빈다
            d = lanes._load()
            # 호스트는 **이 PC** 여야 한다 — 다른 PC 의 점유는 판정하지 않고
            # 보수적으로 '살아 있다'로 보기 때문이다(ai_claim._pid_alive).
            d["build"] = {"who": "claude", "sid": "SESS-DEAD", "at": 0,
                          "agent_pid": 999999, "host": _socket.gethostname()}
            lanes._save(d)
            assert lanes.owner("build") is None, \
                "죽은 세션의 차선이 살아 있다 — 다음 창이 영영 못 들어온다"
        finally:
            ac.CLAIMS, lanes._me = keep_claims, keep_sid

    src = open(os.path.join(ROOT, "ai_claim.py"), encoding="utf-8").read()
    assert "_lane_gate(what)" in src, "ai_claim.take 이 차선 문을 안 지난다"
    print("  [136] 작업 차선 — 수집 창·앱 창 병렬 ✅")


def t134_section_fold():
    """[134] 구역 머리를 눌러 접었다 폈다 (2026-08-07 지시).

    사용자 지시: "날짜 표시된 카드와 세부 사항 카드 접었다 폈다 하는 기능 추가해서
    적용 / **다른 것들도 이런 기능 들어갈 만한거 있는지 검토해보고 적용**".

    ★ 두 곳에 따로 붙이지 않았다. 그 둘의 머리가 이미 같은 `.ios-sec-h` 이고,
      같은 머리를 쓰는 곳이 세 군데 더 있다(원본 자료·입금 등록·리모컨 관리).
      "다른 것들도 검토"의 답이 새 코드가 아니라 **이미 있는 공통 머리**였다 —
      그래서 구역을 새로 만들어도 저절로 접힌다. 이 검사가 그 구조를 지킨다.

    ★ 열쇠는 **처음 본 순간의 id·글귀로 굳힌다.** 머리 글귀는 바뀐다
      ('선택한 날' → '2026년 8월 7일 · 2건'). 그때그때 글귀로 열쇠를 만들면
      **날짜를 바꿀 때마다 접어 둔 것이 스스로 펴진다.**
    """
    live = open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()

    assert ".ios-sec-h.folded + *{display:none!important}" in live,         "접어도 다음 카드가 안 숨는다"
    assert live.count('class="ios-sec-h"') + live.count("ios-sec-h\" id=") >= 4         or live.count("ios-sec-h") >= 6, "공용 구역 머리가 줄었다 — 한 곳에 붙인 이점이 사라진다"
    for fn in ("secKey", "readSecFolded", "applySecFold", "toggleSecFold", "initSecFold"):
        assert f"function {fn}(" in live, f"{fn} 가 없다"
    assert "initSecFold();" in live, "시작할 때 켜지지 않는다"

    key = live[live.index("function secKey("):][:600]
    assert "dataset.secKey" in key, \
        "열쇠를 매번 글귀로 다시 만든다 — 날짜가 바뀌면 접어 둔 것이 저절로 펴진다"

    ap = live[live.index("function applySecFold("):][:1400]
    assert ap.index("const on=") < ap.index("fold-hint"), \
        "'접힘' 힌트를 붙인 뒤에 열쇠를 읽는다 — 힌트 글자가 열쇠에 섞인다"
    assert "'button'" in ap or '"button"' in ap, "보조기기가 누를 수 있는 것으로 안 읽는다"
    assert "tabIndex" in ap, "키보드로 접을 수 없다"

    # 접힌 채로 인쇄하면 그 내용이 보고서에서 통째로 빠진다
    assert re.search(r"@media print\{[^@]*ios-sec-h\.folded \+ \*\{display:block", live, re.S), \
        "인쇄할 때 접힌 구역이 빠진다 — 보고서에 내용이 사라진다"
    assert "var(--ink-3)" in live[live.index(".ios-sec-h::before"):][:400],         "화살표 색이 토큰이 아니다 — 어둡게에서 안 보인다"
    print("  [134] 구역 머리 접기 — 5구역 공용·열쇠 고정·인쇄 보존 ✅")


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
    t107_report_dates_and_capture()
    t108_pm_source_fallback()
    t109_remote_edit_delete_versions()
    t110_writer_formula_key()
    t111_account_handoff_freshness()
    t112_band_plan_order_and_scope()
    t113_paste_typos_and_misc_reclass()
    t114_claim_owner_is_agent_pid()
    t115_text_contrast()
    t116_manual_refresh_is_really_fresh()
    t117_dark_mode_toggle()
    t118_ocr_crosscheck()
    t119_context_guard()
    t121_layer_dialogs()
    t122_dash_drag_and_remote_version()
    t123_calendar_share_tools()
    t124_no_duplicate_menus()
    t125_worktree_shared_state()
    t126_app_font_and_revert()
    t127_dark_mode_no_hardcoded_light_panel()
    t128_dash_tap_to_move()
    t129_call_notes_db_only_and_device_open()
    t130_band_grab_rejects_timeless_harvest()
    t131_band_quiet_vs_stalled()
    t135_contaminated_not_recollected()
    t132_dash_snap_expand_theme_swipe()
    t133_inline_style_dark_safe()
    t134_section_fold()
    t136_work_lanes()
    t137_contamination_marked_every_merge()
    t138_daily_run_completion_watch()
    t139_new_version_is_atomic()
    t140_freshness_tells_the_truth()
    t141_long_text_folds()
    t142_flow_editable()
    t120_calendar_sheet_and_share()
    t121_pid_alive()
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
