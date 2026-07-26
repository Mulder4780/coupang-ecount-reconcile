# -*- coding: utf-8 -*-
"""
probe_check.py — **문제를 찾으러 다니는** 합성 점검 (synthetic_check의 반대 방향)
================================================================================
synthetic_check.py 는 "이미 아는 규칙이 안 깨졌는지" 지킨다.
이 스크립트는 반대로 **아직 모르는 구멍을 찾는다** — 도구마다 합성 입력을 밀어 넣어
결과가 상식과 어긋나는 곳을 보고한다. 실패해도 exit 0(찾은 것을 나열하는 게 목적).

  python tests/probe_check.py
"""
import sys, os, re, json, tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

FOUND, NOTE = [], []


def bad(area, what, detail=""):
    FOUND.append((area, what, detail))


def note(area, what, detail=""):
    """고쳐 뒀거나 다른 도구가 처리하는 것 — 알고는 있어야 하는 항목"""
    NOTE.append((area, what, detail))


def ok(area, what):
    print(f"  · {area}: {what} 정상")


# ── 1. 행 재배치: 셀에 붙은 부속물이 따라가는가 ──────────────────────────
def p_reorder_attachments():
    import reorder_rows as R
    row = ('<row r="9"><c r="A9" t="inlineStr"><is><t>AS-1</t></is></c>'
           '<c r="AM9" t="inlineStr"><is><t>링크</t></is></c></row>')
    moved = R.move_row(row, 300)
    if 'r="AM300"' not in moved:
        bad("reorder", "셀 참조가 안 따라감", moved[:80])
    # 꼬리(hyperlink/조건부서식)는 move_row 밖이라 여기선 시트 단위로 본다
    xml = ('<worksheet><sheetData>'
           '<row r="5"><c r="A5" t="inlineStr"><is><t>2026-01-02</t></is></c></row>'
           '<row r="6"><c r="A6" t="inlineStr"><is><t>2026-01-01</t></is></c></row>'
           '</sheetData>'
           '<hyperlinks><hyperlink ref="AM5" r:id="rId1"/></hyperlinks>'
           '<conditionalFormatting sqref="V5"><cfRule type="expression"/></conditionalFormatting>'
           '</worksheet>')
    head, ordered, moved_n, _ = R.plan(xml, ("일자",), "ID", None) if False else (None, None, None, None)
    # plan은 머리글이 필요해 여기선 rebuild 경로만 확인한다
    if "hyperlink" in xml:
        # 실제 워크북 검사에서 잡는다(아래 p_real_workbook)
        pass


# ── 2. 행 확장: 검증·조건부서식 범위가 새 행까지 늘어나는가 ─────────────
def p_expand_ranges():
    import expand_rows as E
    src = ('<worksheet><sheetData>'
           '<row r="4"><c r="A4" t="inlineStr"><is><t>ID</t></is></c></row>'
           '<row r="5"><c r="A5"><f>IF(1,2,3)</f><v>2</v></c></row>'
           '</sheetData>'
           '<dataValidation type="list" sqref="A5:A5"><formula1>"a,b"</formula1></dataValidation>'
           '<conditionalFormatting sqref="A5:A5"><cfRule type="expression" priority="1"/></conditionalFormatting>'
           '<autoFilter ref="A4:A5"/></worksheet>')
    if not hasattr(E, "expand_sheet_xml"):
        bad("expand", "expand_sheet 함수를 못 찾음 — 점검 불가", ",".join(
            n for n in dir(E) if not n.startswith("_"))[:120])
        return
    out, _lo, _hi = E.expand_sheet_xml(src, 3)
    for tag, pat in (("dataValidation", r'sqref="A5:A(\d+)"'),
                     ("conditionalFormatting", r'sqref="A5:A(\d+)"'),
                     ("autoFilter", r'ref="A4:A(\d+)"')):
        m = re.search(pat, out[out.index("</sheetData>"):]) if "</sheetData>" in out else None
        if not m or int(m.group(1)) < 8:
            bad("expand", f"{tag} 범위가 새 행까지 안 늘어남", (m.group(0) if m else "없음"))


# ── 3. 날짜 값: 문자열로 들어가면 정렬·집계가 어긋난다 ──────────────────
def p_date_types():
    import ledger_writer as L
    for vtype, want in (("date", "serial"), ("number", "num"), ("text", "str")):
        xml = L.cell_xml("D", 5, 3, "2026-07-26" if vtype == "date" else 1000, vtype)
        if vtype == "date" and ('t="' in xml):
            bad("writer", "날짜인데 t 속성이 붙음(문자열로 저장될 수 있음)", xml)
        if vtype == "number" and 't="s"' in xml:
            bad("writer", "숫자인데 공유문자열로 저장", xml)
    # 잘못된 키 이름을 넘겼을 때 조용히 text가 되지 않는지
    import inspect
    src = inspect.getsource(L.resolve_targets)
    if 'get("vtype"' in inspect.getsource(L) and '"type"' not in src:
        ok("writer", "vtype 키만 사용")


# ── 4. 수식 범위 확장이 꼬리(검증·조건부서식)를 망가뜨리지 않는가 ────────
def p_widen_tail():
    import fix_formulas as F
    xml = ('<sheetData><row r="5"><c r="A5"><f>SUM(\'02_돌발AS접수\'!$A$5:$A$154)</f><v>1</v></c></row></sheetData>'
           '<dataValidation sqref="B5:B154"><formula1>\'02_돌발AS접수\'!$A$5:$A$154</formula1></dataValidation>')
    out, n = F.widen(xml, {"02_돌발AS접수": 594})
    if "$A$594" not in out:
        bad("widen", "본문 범위가 안 늘어남", out[:90])
    if out.count("$A$594") < 2:
        bad("widen", "검증 목록(formula1)의 범위는 그대로 — 새 행이 목록에서 빠짐",
            re.search(r"<formula1>[^<]*</formula1>", out).group(0) if "formula1" in out else "")
    if 'sqref="B5:B154"' in out:
        note("widen", "dataValidation sqref는 widen이 안 건드림",
             "행 확장(expand_rows)이 담당한다 — 실제 워크북은 현재 행수까지 맞아 있음")


# ── 5. 날짜 정렬: 형식이 섞여도 시간순인가 ─────────────────────────────
def p_sort():
    sys.path.insert(0, os.path.join(ROOT, "webapp"))
    from app_server import sort_by_date, norm_date
    rows = [{"접수일자": "2026-1-5"}, {"접수일자": "2026-01-04 00:00:00"},
            {"접수일자": ""}, {"접수일자": "2025-12-31"}, {"접수일자": "2026/01/03"}]
    got = [norm_date(r["접수일자"]) for r in sort_by_date(rows, "as")]
    if got != ["2025-12-31", "2026-01-03", "2026-01-04", "2026-01-05", ""]:
        bad("sort", "형식 혼재 시 시간순이 깨짐", str(got))


# ── 6. 계산서 구성: 연말·부가세 표기 흔들림 ────────────────────────────
def p_bundle_edges():
    import erp_bundle as B
    # 총금액이 부가세 포함일 때도 공급가로 찾아지는가
    poinx = {1100000: ["UJ2600001"], 1000000: ["UJ2600001"]}
    _, v, _ = B.bundle({"slip": "2026/05/11-1", "amt": 1000000, "camp": "x",
                        "kind": "돌발AS", "month": "2026/05", "title": ""}, [], None, poinx)
    if v != "확정(밴드PO)":
        bad("bundle", "PO 총금액이 부가세 포함일 때 공급가로 못 찾음", v)
    # 12월 건을 1월에 올린 글 → 연도가 작년이어야 한다
    src = B.band_invoice_index.__doc__ or ""
    _, v2, _ = B.bundle({"slip": "2026/01/05-1", "amt": 100, "camp": "양주2캠프",
                         "kind": "정기점검", "month": "2026/01",
                         "title": "25년 4분기 정기점검"}, [])
    lo, hi = B.window({"month": "2026/01", "title": "25년 4분기 정기점검"})
    if (lo, hi) != ("2025-10", "2025-12"):
        bad("bundle", "분기 표기 해석 오류", f"{lo}~{hi}")


# ── 7. 캠프 추출: 목록글 패턴이 엉뚱한 줄을 잡지 않는가 ─────────────────
def p_camp():
    import camp_fill as C
    good = "3. 송파1MB(감일동)(1/10) : 2R/T 19,780,000원 / PO326259 UJ2501950"
    m = C.LIST_RE.search(good)
    if not m or m.group(1) != "송파1MB(감일동)" or m.group(2) != "UJ2501950":
        bad("camp", "목록글에서 캠프·번호를 잘못 뽑음", str(m.groups() if m else None))
    # 캠프가 아닌 묶음 라벨이 캠프로 들어가지 않는지
    grp = "1. A/S_1(2/10) : 12,000,000원 / PO111 UJ2600001"
    m2 = C.LIST_RE.search(grp)
    if m2 and m2.group(1) == "A/S_1":
        if getattr(C, "LABEL_RE", None) and C.LABEL_RE.match("A/S_1"):
            note("camp", "묶음 라벨은 LABEL_RE로 걸러짐", "A/S_1·정기점검_2 등")
        else:
            bad("camp", "묶음 라벨(A/S_1)이 캠프명으로 입력될 수 있음", "")


# ── 8. 실제 워크북: 셀에 붙은 부속물이 재배치 뒤에도 맞는가 ─────────────
def p_real_workbook():
    import zipfile
    import fix_formulas as F
    from ecount_reconcile import load_config, resolve_master
    z = zipfile.ZipFile(resolve_master(load_config()["reconcile"]["master_xlsx"]))
    for name, p in F.sheet_map(z).items():
        if name not in ("02_돌발AS접수", "03_현장작업실적", "04_정기점검", "06_거래서류청구수금"):
            continue
        x = z.read(p).decode("utf-8")
        tail = x[x.index("</sheetData>"):]
        import reorder_rows as R
        has_fix = hasattr(R, "move_hyperlinks")
        for m in re.finditer(r'<hyperlink\b[^>]*ref="([A-Z]+)(\d+)"', tail):
            if int(m.group(2)) >= 5:
                (note if has_fix else bad)(
                    "workbook", f"{name}: 셀 하이퍼링크 {m.group(1)}{m.group(2)}",
                    "재배치 때 함께 이동함(move_hyperlinks)" if has_fix
                    else "행이 옮겨지면 다른 건의 링크가 된다")
        for m in re.finditer(r'<(conditionalFormatting|dataValidation)\b[^>]*(?:sqref)="([^"]+)"', tail):
            for part in m.group(2).split():
                rr = [int(r) for r in re.findall(r"[A-Z]+(\d+)", part)]
                if rr and min(rr) >= 5 and (max(rr) - min(rr)) < 50:
                    note("workbook", f"{name}: 특정 행 구간에만 걸린 {m.group(1)}",
                        f"{part} — 행이 옮겨지면 엉뚱한 건에 적용된다")
    z.close()


def main():
    print("■ 합성 점검 — 문제 찾기")
    for fn in (p_reorder_attachments, p_expand_ranges, p_date_types, p_widen_tail,
               p_sort, p_bundle_edges, p_camp, p_real_workbook):
        try:
            fn()
        except Exception as e:
            bad(fn.__name__, f"점검 중 예외: {type(e).__name__} {e}", "")
    print()
    if NOTE:
        print(f"■ 주의(알고 있으면 되는 것) {len(NOTE)}건")
        seen2 = set()
        for area, what, detail in NOTE:
            if (area, what) in seen2:
                continue
            seen2.add((area, what))
            print(f"  [{area}] {what}" + (f" — {detail[:90]}" if detail else ""))
        print()
    if not FOUND:
        print("문제 없음 ✅")
        return
    print(f"■ 발견 {len(FOUND)}건")
    seen = set()
    for area, what, detail in FOUND:
        key = (area, what)
        if key in seen:
            continue
        seen.add(key)
        n = sum(1 for a, w, _ in FOUND if (a, w) == key)
        print(f"  [{area}] {what}" + (f"  (같은 유형 {n}건)" if n > 1 else ""))
        if detail:
            print(f"       → {detail[:150]}")


if __name__ == "__main__":
    main()
