# -*- coding: utf-8 -*-
"""
patch_formula.py — 시트의 **특정 셀 수식 한 개**를 안전하게 바꾼다(zip 패치)
================================================================================
ledger_writer는 값만 넣는다. 집계 수식 자체를 고쳐야 할 때 쓴다.
openpyxl save 금지 규칙을 지키려면 이렇게 zip 파트만 갈아끼워야 한다.

  python patch_formula.py "00_대시보드" H21 "SUMPRODUCT(...)"        # 미리보기
  python patch_formula.py "00_대시보드" H21 "SUMPRODUCT(...)" --apply

수식은 **= 없이** 본문만 넘긴다. 캐시값은 지워 두고 엑셀이 다시 계산하게 한다.
"""
import sys, os, re, zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from fix_formulas import sheet_map, parse_cells, esc


def main():
    args = [a for a in sys.argv[1:] if a != "--apply"]
    if len(args) < 3:
        sys.exit(__doc__)
    sheet, cell, formula = args[0], args[1].upper(), args[2].lstrip("=")
    col = re.match(r"([A-Z]+)(\d+)", cell)
    if not col:
        sys.exit(f"셀 주소 오류: {cell}")
    colL, rown = col.group(1), int(col.group(2))

    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    z = zipfile.ZipFile(master)
    smap = sheet_map(z)
    if sheet not in smap:
        sys.exit(f"시트 없음: {sheet}")
    xml = z.read(smap[sheet]).decode("utf-8")
    z.close()

    hit = [c for c in parse_cells(xml) if c[0] == rown and c[1] == colL]
    if not hit:
        sys.exit(f"{sheet} {cell} 에 수식이 없습니다(값 셀이면 ledger_writer를 쓰세요)")
    _, _, old, s, e = hit[0]
    print(f"{sheet} {cell}")
    print(f"  전: ={old[:150]}")
    print(f"  후: ={formula[:150]}")

    chunk = xml[s:e]
    new = re.sub(r"(<f(?:\s[^>]*)?>)(.*?)(</f>)",
                 lambda m: m.group(1) + esc(formula) + m.group(3), chunk, count=1, flags=re.S)
    new = re.sub(r"<v[^>]*>.*?</v>|<v\s*/>", "<v/>", new, count=1, flags=re.S)
    new = re.sub(r'\st="e"', "", new, count=1)
    out = xml[:s] + new + xml[e:]

    if "--apply" not in sys.argv:
        print("\n반영하려면 끝에 --apply")
        return
    mv = re.search(r"_v(\d+)\.xlsx$", master)
    dst = re.sub(r"_v\d+\.xlsx$", f"_v{int(mv.group(1))+1}.xlsx", master)
    if os.path.exists(dst):
        sys.exit(f"{os.path.basename(dst)} 이미 존재 — 중단")
    zin = zipfile.ZipFile(master)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        zout.writestr(it.filename,
                      out.encode("utf-8") if it.filename == smap[sheet] else zin.read(it.filename))
    zout.close(); zin.close()
    zc = zipfile.ZipFile(dst)
    assert zc.testzip() is None, "zip 무결성 실패"
    zc.close()
    print(f"\n생성: {os.path.basename(dst)}")


if __name__ == "__main__":
    main()
