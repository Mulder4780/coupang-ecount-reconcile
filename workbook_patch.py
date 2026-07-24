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
from datetime import date

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
    return best, ver(best)


def sheet_xml_path(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(r'<sheet name="%s"[^>]*r:id="(rId\d+)"' % SHEET_NAME, wb)
    if not m:
        sys.exit(f"시트 '{SHEET_NAME}' 를 찾을 수 없습니다.")
    m2 = re.search(r'<Relationship Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels)
    t = m2.group(1)
    return "xl/" + t if not t.startswith("/") else t[1:]


def patch(src, dst, a, b, c):
    zin = zipfile.ZipFile(src)
    target = sheet_xml_path(zin)
    xml = zin.read(target).decode("utf-8")
    rows = [int(r) for r in re.findall(r'<row r="(\d+)"', xml)]
    nr = max(rows) + 1
    assert f'<row r="{nr}"' not in xml
    row = (f'<row r="{nr}" spans="1:3">'
           f'<c r="A{nr}" s="93" t="inlineStr"><is><t>{esc(a)}</t></is></c>'
           f'<c r="B{nr}" s="93" t="inlineStr"><is><t>{esc(b)}</t></is></c>'
           f'<c r="C{nr}" s="93" t="inlineStr"><is><t>{esc(c)}</t></is></c></row>')
    xml = xml.replace("</sheetData>", row + "</sheetData>", 1)
    xml = re.sub(r'<dimension ref="A1:C\d+"/>', f'<dimension ref="A1:C{nr}"/>', xml, count=1)

    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        d = zin.read(it.filename)
        if it.filename == target:
            d = xml.encode("utf-8")
        # 주의: 원본 ZipInfo를 그대로 넘기면 writestr가 오프셋을 변조하므로 이름만 전달
        zout.writestr(it.filename, d)
    zout.close()
    zin.close()

    # 3중 검증 (원본·결과 모두 새로 열어서 — ZipInfo 오염 방지)
    z = zipfile.ZipFile(dst)
    assert z.testzip() is None, "zip 무결성 실패"
    import openpyxl
    w = openpyxl.load_workbook(dst, read_only=True)
    vals = [cell.value for cell in next(w[SHEET_NAME].iter_rows(min_row=nr, max_row=nr, max_col=3))]
    assert vals[0] == a and vals[1] == b, "새 행 판독 실패"
    w.close()
    zsrc = zipfile.ZipFile(src)
    diff = [n for n in zsrc.namelist() if n != target and zsrc.read(n) != z.read(n)]
    assert not diff, f"의도치 않은 파트 변경: {diff}"
    zsrc.close(); z.close()
    return nr


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--a", default="", help="A열(날짜 #순번). 생략 시 자동")
    ap.add_argument("--b", required=True, help="B열(작업 제목)")
    ap.add_argument("--c", required=True, help="C열(상세 내용)")
    args = ap.parse_args()

    src, v = latest_master()
    dst = src.replace(f"_v{v}.xlsx", f"_v{v+1}.xlsx")
    if os.path.exists(dst):
        sys.exit(f"{os.path.basename(dst)} 이미 존재 — 중복 실행 방지를 위해 중단")
    a = args.a or f"{date.today().isoformat()} #auto"
    nr = patch(src, dst, a, args.b, args.c)
    print(f"OK: v{v} → v{v+1} 생성, {SHEET_NAME} {nr}행 추가, 검증 3종 통과")
    print("   ", dst)


if __name__ == "__main__":
    main()
