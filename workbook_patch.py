# -*- coding: utf-8 -*-
"""
workbook_patch.py — 관리대장 안전 버전업 도구 (vN → vN+1)
===========================================================
관리대장의 19_AI작업인수인계 시트에 인수인계 행을 추가하며 다음 버전 파일을 만든다.

★ 왜 이 방식인가: openpyxl로 열어 save()하면 차트·도형버튼·x14 검증이 파괴된다(19시트 경고).
  이 도구는 zip 안의 해당 시트 XML 1개만 교체하므로 나머지 전체가 바이트 단위로 보존된다.
  매 실행 후 3중 검증(zip 무결성 / 새 행 판독 / 타 파트 동일성)을 자동 수행한다.

사용:
    python workbook_patch.py --b "작업 제목" --c "상세 내용"
    (--a 생략 시 오늘 날짜 + 자동 순번, 원본은 보존되고 vN+1 파일이 새로 생성됨)
"""
import sys, os, re, json, glob, zipfile, argparse
from datetime import date, time as clock_time, timedelta

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config", "ecount_config.json")
SHEET_NAME = "19_AI작업인수인계"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def latest_master():
    cfg = json.load(open(CONFIG_PATH, encoding="utf-8"))
    folder = os.path.dirname(cfg["reconcile"]["master_xlsx"])
    cands = glob.glob(os.path.join(folder, "쿠팡_통합업무_일일보고_관리대장_v*.xlsx"))
    def ver(p):
        m = re.search(r"_v(\d+)\.xlsx$", p)
        return int(m.group(1)) if m else -1
    cands = [c for c in cands if ver(c) >= 0 and "~$" not in c]
    if not cands:
        sys.exit("관리대장 v*.xlsx 를 찾을 수 없습니다.")
    best = max(cands, key=ver)
    # 구 버전은 말 안 해도 OLD 로 접는다(사용자 지시 2026-07-28). 최신본은 손대지 않는다.
    try:
        from ledger_versions import autoprune
        autoprune(best)
    except Exception:
        pass
    return best, ver(best)


def sheet_xml_path(z, sheet_name=SHEET_NAME):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = None
    for tag in re.findall(r"<sheet\b[^>]*/?>", wb):
        mn = re.search(r'name="([^"]+)"', tag)
        mr = re.search(r'r:id="([^"]+)"', tag)
        if mn and mr and mn.group(1) == sheet_name:
            rid = mr.group(1)
            break
    if not rid:
        sys.exit(f"시트 '{sheet_name}' 를 찾을 수 없습니다.")
    t = None
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        mi = re.search(r'Id="([^"]+)"', tag)
        mt = re.search(r'Target="([^"]+)"', tag)
        if mi and mt and mi.group(1) == rid:
            t = mt.group(1)
            break
    if not t:
        sys.exit(f"시트 관계 '{sheet_name}' 를 찾을 수 없습니다.")
    return t[1:] if t.startswith("/") else (t if t.startswith("xl/") else "xl/" + t)


def append_row(xml, values):
    """시트 마지막에 값 행 1개를 추가한다. values=[(열, 값, 숫자여부)]."""
    rows = [int(r) for r in re.findall(r'<row r="(\d+)"', xml)]
    nr = max(rows) + 1
    cells = []
    for col, value, number in values:
        styles = re.findall(r'<c r="%s\d+"[^>]*\ss="(\d+)"' % col, xml)
        s_attr = f' s="{styles[-1]}"' if styles else ""
        if number:
            cells.append(f'<c r="{col}{nr}"{s_attr}><v>{value}</v></c>')
        else:
            cells.append(f'<c r="{col}{nr}"{s_attr} t="inlineStr"><is><t>{esc(str(value))}</t></is></c>')
    xml = xml.replace("</sheetData>", f'<row r="{nr}">{"".join(cells)}</row></sheetData>', 1)
    xml = re.sub(r'(<dimension ref="[^"]*:[A-Z]{1,3})\d+("/>)',
                 lambda m: f"{m.group(1)}{nr}{m.group(2)}", xml, count=1)
    return xml, nr


def replace_inline_cell(xml, ref, value):
    """기존 셀의 스타일을 보존하며 문자열 값을 교체한다.

    ★ 2026-07-28 실사고: 예전 패턴은 `<c r="AJ33" s="5"/>` 처럼 **빈 셀(자기닫힘)** 에서
      `[^>]*` 가 `/` 까지 삼킨 뒤 `>.*?</c>` 가 **다음 셀을 통째로 먹어치웠다.**
      (03_현장작업실적 33행에서 비고를 채우려다 옆 칸 AK33 이 사라졌다 — 검증이 잡았다.)
      속성은 `이름="값"` 형태만 받도록 좁혀 `/` 를 삼키지 못하게 한다."""
    m = re.search(r'<c r="%s"(?:\s+[a-zA-Z:]+="[^"]*")*\s*(?:/>|>.*?</c>)' % re.escape(ref),
                  xml, re.S)
    if not m:
        raise AssertionError(f"{ref} 셀 XML 없음")
    sm = re.search(r'\ss="(\d+)"', m.group(0))
    s_attr = f' s="{sm.group(1)}"' if sm else ""
    new = f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{esc(value)}</t></is></c>'
    return xml[:m.start()] + new + xml[m.end():]


def replace_number_cell(xml, ref, value):
    """Keep a cell's style while replacing it with a numeric value."""
    m = re.search(r'<c r="%s"(?:\s+[a-zA-Z:]+="[^"]*")*\s*(?:/>|>.*?</c>)' % re.escape(ref),
                  xml, re.S)
    if not m:
        raise AssertionError(f"{ref} cell XML not found")
    sm = re.search(r'\ss="(\d+)"', m.group(0))
    s_attr = f' s="{sm.group(1)}"' if sm else ""
    new = f'<c r="{ref}"{s_attr}><v>{value}</v></c>'
    return xml[:m.start()] + new + xml[m.end():]


def parse_cutoff(value):
    """Return the Excel duration, display format, and normalized cutoff text."""
    m = re.fullmatch(r"(\d{1,2}):(\d{2})", str(value or "").strip())
    if not m:
        raise ValueError("집계마감시간은 HH:MM 형식이어야 합니다 (예: 23:59)")
    hour, minute = int(m.group(1)), int(m.group(2))
    if minute > 59 or hour > 24 or (hour == 24 and minute != 0):
        raise ValueError("집계마감시간은 00:00~24:00 범위여야 합니다")
    normalized = f"{hour:02d}:{minute:02d}"
    serial = (hour * 60 + minute) / (24 * 60)
    # Excel must use a bracketed format only for the one valid 24:00 duration.
    return serial, ("[h]:mm" if hour == 24 else "hh:mm"), normalized


def set_dashboard_cutoff(zin, changed, verify, cutoff):
    """Set dashboard B5 to a real, calculation-safe aggregation cutoff."""
    serial, number_format, display = parse_cutoff(cutoff)
    dash_name = "00_대시보드"
    dash_path = sheet_xml_path(zin, dash_name)
    dash_xml = zin.read(dash_path).decode("utf-8")
    dash_xml = replace_number_cell(dash_xml, "B5", f"{serial:.12g}")
    dash_xml = replace_inline_cell(
        dash_xml, "C3",
        f"※ 집계 원칙: 당일 {display}(B5)까지 등록된 건은 집계기준일에 반영, 이후 등록건은 다음 업무일로 자동 이월(토·일·10_코드관리 공휴일 제외).",
    )
    changed[dash_path] = dash_xml.encode("utf-8")

    # Format 177 is dedicated to dashboard B5, verified by regression tests.
    styles = zin.read("xl/styles.xml").decode("utf-8")
    updated, n = re.subn(
        r'<numFmt numFmtId="177" formatCode="[^"]*"/>',
        f'<numFmt numFmtId="177" formatCode="{number_format}"/>',
        styles,
        count=1,
    )
    if n != 1:
        raise AssertionError("dashboard cutoff number format 177 not found")
    changed["xl/styles.xml"] = updated.encode("utf-8")
    verify[dash_name] = ("cutoff", (display, number_format))


def patch(src, dst, a, b, c, a2="", manual18="", manual20="",
          manual20_title="", manual20_area="", cutoff=""):
    zin = zipfile.ZipFile(src)
    target = sheet_xml_path(zin, SHEET_NAME)
    xml, nr = append_row(zin.read(target).decode("utf-8"),
                         [("A", a, False), ("B", b, False), ("C", c, False)])
    changed = {target: xml.encode("utf-8")}
    verify = {SHEET_NAME: (nr, [a, b, c])}

    if a2:
        path = sheet_xml_path(zin, "00_대시보드")
        x = zin.read(path).decode("utf-8")
        changed[path] = replace_inline_cell(x, "A2", a2).encode("utf-8")
    if cutoff:
        set_dashboard_cutoff(zin, changed, verify, cutoff)
    if manual18:
        name = "18_문서발행업무매뉴얼"
        path = sheet_xml_path(zin, name)
        x, rown = append_row(zin.read(path).decode("utf-8"), [("A", manual18, False)])
        changed[path] = x.encode("utf-8")
        verify[name] = (rown, [manual18])
    if manual20:
        name = "20_쿠팡통합업무상세매뉴얼"
        path = sheet_xml_path(zin, name)
        x0 = zin.read(path).decode("utf-8")
        rown = max(int(r) for r in re.findall(r'<row r="(\d+)"', x0)) + 1
        title = manual20_title or "22-5. 수식·경고 대응(2026-07-28)"
        area = manual20_area or "담당기사·자동상태"
        x, rown = append_row(x0, [
            ("A", rown, True),
            ("B", title, False),
            ("C", area, False),
            ("D", manual20, False),
        ])
        changed[path] = x.encode("utf-8")
        verify[name] = (rown, [rown, title, area, manual20])

    tmp = dst[:-5] + ".tmp.xlsx"
    if os.path.exists(tmp):
        raise FileExistsError(f"임시 결과가 이미 존재: {tmp}")
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        d = changed.get(it.filename, zin.read(it.filename))
        # 주의: 원본 ZipInfo를 그대로 넘기면 writestr가 오프셋을 변조하므로 이름만 전달
        zout.writestr(it.filename, d)
    zout.close()
    zin.close()

    # 3중 검증 (원본·결과 모두 새로 열어서 — ZipInfo 오염 방지)
    z = zipfile.ZipFile(tmp)
    assert z.testzip() is None, "zip 무결성 실패"
    import openpyxl
    w = openpyxl.load_workbook(tmp, read_only=True)
    for name, (rown, expected) in verify.items():
        if rown == "cutoff":
            dash = w[name]
            display, number_format = expected
            expected_value = timedelta(days=1) if display == "24:00" else clock_time(*map(int, display.split(":")))
            assert dash["B5"].value == expected_value, f"dashboard B5 is not {display}"
            assert dash["B5"].number_format == number_format, "dashboard B5 number format mismatch"
            assert dash["C3"].value and display in str(dash["C3"].value), "dashboard cutoff guide missing"
            continue
        vals = [cell.value for cell in next(
            w[name].iter_rows(min_row=rown, max_row=rown, max_col=len(expected)))]
        assert vals == expected, f"{name} 행 재독 실패: {vals}"
    if a2:
        assert w["00_대시보드"]["A2"].value == a2, "00_대시보드 A2 재독 실패"
    w.close()
    zsrc = zipfile.ZipFile(src)
    diff = [n for n in zsrc.namelist() if n not in changed and zsrc.read(n) != z.read(n)]
    assert not diff, f"의도치 않은 파트 변경: {diff}"
    zsrc.close(); z.close()
    os.replace(tmp, dst)
    return nr


def main():
    # 다른 AI가 원장을 잡고 있으면 여기서 멈춘다(동시 수정 시 한쪽이 통째로 묻힌다)
    from claim_guard import require
    require("ledger", "workbook_patch")
    ap = argparse.ArgumentParser()
    ap.add_argument("--a", default="", help="A열(날짜 #순번). 생략 시 자동")
    ap.add_argument("--b", required=True, help="B열(작업 제목)")
    ap.add_argument("--c", required=True, help="C열(상세 내용)")
    ap.add_argument("--a2", default="", help="00_대시보드 A2 사용법 전체 문구 교체(선택)")
    ap.add_argument("--manual18", default="", help="18_문서발행업무매뉴얼 끝행 추가 문구(선택)")
    ap.add_argument("--manual20", default="", help="20_쿠팡통합업무상세매뉴얼 끝행 추가 설명(선택)")
    ap.add_argument("--manual20-title", default="", help="20시트 추가 행 B열 제목(선택)")
    ap.add_argument("--manual20-area", default="", help="20시트 추가 행 C열 분류(선택)")
    ap.add_argument("--cutoff", default="", metavar="HH:MM",
                    help="00_대시보드 집계마감시간 설정(예: 23:59)")
    ap.add_argument("--cutoff24", action="store_true",
                    help="호환용 별칭: --cutoff 24:00")
    args = ap.parse_args()

    src, v = latest_master()
    dst = src.replace(f"_v{v}.xlsx", f"_v{v+1}.xlsx")
    if os.path.exists(dst):
        sys.exit(f"{os.path.basename(dst)} 이미 존재 — 중복 실행 방지를 위해 중단")
    a = args.a or f"{date.today().isoformat()} #auto"
    if args.cutoff and args.cutoff24:
        ap.error("--cutoff과 --cutoff24는 함께 사용할 수 없습니다")
    cutoff = args.cutoff or ("24:00" if args.cutoff24 else "")
    nr = patch(src, dst, a, args.b, args.c, args.a2, args.manual18, args.manual20,
               args.manual20_title, args.manual20_area, cutoff)
    print(f"OK: v{v} → v{v+1} 생성, {SHEET_NAME} {nr}행 추가, 검증 3종 통과")
    print("   ", dst)


if __name__ == "__main__":
    main()
