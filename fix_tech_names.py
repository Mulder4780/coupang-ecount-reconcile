# -*- coding: utf-8 -*-
"""
fix_tech_names.py — 담당기사 칸에 들어간 **사람 아닌 값**을 바로잡는다
================================================================================
사용자 확인(2026-07-28)으로 값이 확정된 10건을 고친다.

왜 생겼나
  밴드/카톡 글에서 기사명을 뽑는 `band_extract.normalize_tech` 가 '첫 조각'을
  그대로 통과시켰다. 그래서 작업 메모가 담당기사 칸에 들어갔고, 대표보고
  'TOP 5' 에 `담당: 000 (캠프상태확인 및 스케쥴 세팅)` 로 노출됐다.
  (걸러 주는 함수는 있었지만 리포트에서만 쓰이고 원장에 쓸 때는 안 거쳤다.)

원칙
  · **원문을 없애지 않는다.** 담당기사 칸에서 뺀 값은 비고(AJ)에 그대로 남긴다.
    나중에 "왜 이름이 바뀌었지?" 를 되짚을 수 있어야 한다.
  · 행 번호를 박지 않고 **값으로 찾는다** — 행이 밀려도 맞는 곳을 고친다.
  · 관리대장은 zip 패치로만 만진다(차트·도형 보호).

사용
  python fix_tech_names.py            # 무엇을 어떻게 고칠지만 보여 준다
  python fix_tech_names.py --apply    # vN+1 생성
"""
import os
import re
import sys
import zipfile

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from workbook_patch import latest_master, sheet_xml_path, replace_inline_cell, esc  # noqa: E402

SHEET = "02_돌발AS접수"
HDR_ROW = 4
TECH_COL = "담당기사"
NOTE_COL = "비고"

# 현재 값 → 고칠 값. 사용자 확인 결과(2026-07-28):
#   · 김혜진  = 유니버셜 퇴사자, 이전 담당  → 이름이 맞다
#   · 하이테크 = 협력업체, 엄진언 = 사람     → 담당기사는 엄진언
#   · 기장    = 직책                      → 김승기
#   · 000 / 자) = 미배정 자리표시자 + 작업 메모 → 담당기사는 비운다
FIX = {
    "000 (캠프상태확인 및 스케쥴 세팅)": "",
    "자) - 각캠프담당자 캠프 컨디션상태 체크": "",
    "하이테크 + 엄진언": "엄진언",
    "김혜진 대신택배": "김혜진",
    "김승기기장": "김승기",
}


def _cols(ws):
    hdr = [str(h or "").strip() for h in
           next(ws.iter_rows(min_row=HDR_ROW, max_row=HDR_ROW, values_only=True))]
    from openpyxl.utils import get_column_letter as GL
    ix = {h: GL(i + 1) for i, h in enumerate(hdr) if h}
    for need in (TECH_COL, NOTE_COL):
        if need not in ix:
            raise RuntimeError(f"{SHEET} 에 '{need}' 열이 없다")
    return ix[TECH_COL], ix[NOTE_COL]


def survey(path):
    """[(행, 프로젝트NO, 현재기사, 새기사, 현재비고, 새비고)]"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[SHEET]
    tc, nc = _cols(ws)
    out = []
    for r in range(HDR_ROW + 1, ws.max_row + 1):
        cur = str(ws[f"{tc}{r}"].value or "").strip()
        if cur not in FIX:
            continue
        new = FIX[cur]
        note = str(ws[f"{nc}{r}"].value or "").strip()
        keep = f"담당기사 원본: {cur}"
        newnote = note if keep in note else (f"{note} | {keep}" if note else keep)
        out.append((r, str(ws[f"B{r}"].value or ""), cur, new, note, newnote))
    wb.close()
    return out, tc, nc


def blank_cell(xml, ref):
    """스타일만 남기고 완전히 비운다(빈 문자열이 아니라 진짜 빈 칸)."""
    pat = re.compile(r'<c r="%s"((?:\s+[a-zA-Z:]+="[^"]*")*)\s*(?:/>|>.*?</c>)' % ref, re.S)
    m = pat.search(xml)
    if not m:
        raise AssertionError(f"{ref} 셀 XML 없음")
    attrs = re.sub(r'\s+t="[^"]*"', "", m.group(1) or "")
    return xml[:m.start()] + f'<c r="{ref}"{attrs}/>' + xml[m.end():]


def apply(src, dst, rows, tc, nc):
    zin = zipfile.ZipFile(src)
    sp = sheet_xml_path(zin, SHEET)
    xml = zin.read(sp).decode("utf-8")
    # 02시트에는 공유수식이 있다(A·F·M·AK·AL·AN·AP·AR — ID 채번·자동판정 열).
    # 시트에 있다는 이유로 막으면 아무것도 못 고친다. **내가 건드릴 칸이 거기 속하는지**만 본다.
    # 공유수식 칸을 갈아엎으면 같은 그룹의 다른 칸이 통째로 깨진다.
    targets = [f"{c}{r}" for r, *_ in rows for c in (tc, nc)]
    for ref in targets:
        m = re.search(r'<c r="%s"[^>]*>\s*<f([^>]*)>' % ref, xml)
        if m and ('t="shared"' in m.group(1) or 't="array"' in m.group(1)):
            raise RuntimeError(f"{ref} 이 공유·배열 수식이라 손대면 위험합니다 — 중단")

    for r, _prj, _cur, new, _note, newnote in rows:
        xml = (blank_cell(xml, f"{tc}{r}") if not new
               else replace_inline_cell(xml, f"{tc}{r}", new))
        xml = replace_inline_cell(xml, f"{nc}{r}", newnote)

    # 담당기사를 비우면 기사배정상태(M)·대표보고 '담당자 미배정' 수가 따라 바뀌어야 한다.
    # 저장된 계산 결과는 옛 값 그대로이므로 열 때 다시 계산하도록 표시해 둔다.
    from dashboard_clean import force_recalc
    wbp = "xl/workbook.xml"
    changed = {sp: xml.encode("utf-8"),
               wbp: force_recalc(zin.read(wbp).decode("utf-8")).encode("utf-8")}

    tmp = dst[:-5] + ".tmp.xlsx"
    if os.path.exists(tmp):
        raise FileExistsError(f"임시 결과가 이미 존재: {tmp}")
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        zout.writestr(it.filename, changed.get(it.filename, zin.read(it.filename)))
    zout.close()
    zin.close()

    verify(src, tmp, rows, tc, nc, sp)
    os.replace(tmp, dst)


def verify(src, out, rows, tc, nc, sp):
    import openpyxl
    z = zipfile.ZipFile(out)
    assert z.testzip() is None, "zip 무결성 실패"
    wa = openpyxl.load_workbook(src, data_only=True)
    wb_ = openpyxl.load_workbook(out, data_only=True)
    sa, sb = wa[SHEET], wb_[SHEET]

    touched = set()
    for r, prj, _cur, new, _note, newnote in rows:
        got = str(sb[f"{tc}{r}"].value or "").strip()
        assert got == new, f"{r}행 담당기사 {got!r} ≠ {new!r}"
        assert str(sb[f"{nc}{r}"].value or "").strip() == newnote, f"{r}행 비고 반영 실패"
        assert str(sb[f"B{r}"].value or "") == prj, f"{r}행이 밀렸다"
        touched |= {f"{tc}{r}", f"{nc}{r}"}

    # ★ 고친 칸 말고는 **한 칸도** 안 바뀌어야 한다
    diff = []
    for r in range(1, max(sa.max_row, sb.max_row) + 1):
        for c in range(1, max(sa.max_column, sb.max_column) + 1):
            ca, cb = sa.cell(r, c), sb.cell(r, c)
            if ca.coordinate in touched:
                continue
            if ca.value != cb.value:
                diff.append(ca.coordinate)
                if len(diff) > 5:
                    break
    assert not diff, f"의도치 않은 셀 변경: {diff}"
    assert wa.sheetnames == wb_.sheetnames, "시트 구성이 바뀌었다"
    wa.close(); wb_.close()

    zs = zipfile.ZipFile(src)
    allowed = {sp, "xl/workbook.xml"}
    other = [n for n in zs.namelist() if n not in allowed and zs.read(n) != z.read(n)]
    assert not other, f"의도치 않은 파트 변경: {other}"
    zs.close(); z.close()
    print(f"  검증 통과 — {len(rows)}행 수정 · 다른 셀·파트 변경 없음")


def main():
    do = "--apply" in sys.argv
    m = latest_master()
    src = m[0] if isinstance(m, tuple) else m
    print(f"원본: {os.path.basename(src)}\n")

    rows, tc, nc = survey(src)
    if not rows:
        print("고칠 행이 없습니다 — 이미 정리됐습니다.")
        return 0
    for r, prj, cur, new, _note, _nn in rows:
        print(f"  {r:>4}행 {prj:11s} {cur!r}")
        shown = repr(new) if new else "(비움 — 미배정)"
        print(f"        → 담당기사 {shown} · 비고에 '담당기사 원본: {cur}' 추가")
    print(f"\n합계 {len(rows)}행")
    if not do:
        print("\n실제로 고치려면:  python fix_tech_names.py --apply")
        return 0

    mm = re.search(r"_v(\d+)\.xlsx$", src)
    dst = re.sub(r"_v\d+\.xlsx$", f"_v{int(mm.group(1)) + 1}.xlsx", src)
    print(f"\n결과: {os.path.basename(dst)}")
    apply(src, dst, rows, tc, nc)
    return 0


if __name__ == "__main__":
    sys.exit(main())
