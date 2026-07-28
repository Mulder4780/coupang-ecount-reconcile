# -*- coding: utf-8 -*-
"""관리대장 보고 통계를 2026년 업무로 한정한다.

원본 2025년 행은 보존한다. 이 도구는 00_대시보드·01_대표보고의 집계 수식과
07_불일치누락현황의 숨김 연도 보조열(Q)만 ZIP 수준에서 패치한다.
openpyxl save()는 사용하지 않는다.

실행:
  python stats_2026.py            # 미리보기
  python stats_2026.py --apply    # ledger 점유 후 vN+1 생성
"""
from __future__ import annotations

import os
import re
import sys
import zipfile
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from fix_formulas import esc, force_recalc_part, parse_cells
from fix_workbook import col_num
from workbook_patch import latest_master, sheet_xml_path

YEAR = 2026
DASH = "00_대시보드"
REPORT = "01_대표보고"
ISSUES = "07_불일치누락현황"

# 각 원천 행의 업무 연도를 판정하는 기준 열. 앱의 2026년 판정과 같은 우선 기준이다.
YEAR_SOURCE = {
    "02_돌발AS접수": ("D", "date"),
    "03_현장작업실적": ("E", "date"),
    "04_정기점검": ("D", "date"),
    "05_신규납품설치": ("E", "date"),
    "06_거래서류청구수금": ("L", "date"),
    ISSUES: ("Q", "year"),
    "15_세금계산서관리": ("B", "settlement_id"),
    "16_입금수금관리": ("B", "settlement_id"),
    "17_문서대조현황": ("A", "settlement_id"),
}


def _split_args(body: str) -> list[str]:
    """Excel 함수 인수를 문자열·괄호를 보존한 채 최상위 쉼표로 나눈다."""
    out, start, depth, quoted = [], 0, 0, False
    i = 0
    while i < len(body):
        ch = body[i]
        if ch == '"':
            if quoted and i + 1 < len(body) and body[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                out.append(body[start:i])
                start = i + 1
        i += 1
    out.append(body[start:])
    return out


def _call_spans(formula: str):
    """집계 함수 호출 범위를 바깥 수식 안에서 찾는다."""
    rx = re.compile(r"\b(COUNTIFS|COUNTIF|SUMIFS|SUMIF|COUNTA|COUNT|SUM)\(", re.I)
    for match in rx.finditer(formula):
        depth, quoted, i = 1, False, match.end()
        while i < len(formula) and depth:
            ch = formula[i]
            if ch == '"':
                if quoted and i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 2
                    continue
                quoted = not quoted
            elif not quoted:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
            i += 1
        if depth == 0:
            yield match.start(), i, match.group(1).upper(), formula[match.end():i - 1]


def _range_sheet(expr: str):
    match = re.search(
        r"'([^']+)'!\$([A-Z]{1,3})\$(\d+):\$([A-Z]{1,3})\$(\d+)",
        expr,
    )
    if not match or match.group(1) not in YEAR_SOURCE:
        return None
    return match.group(1), int(match.group(3)), int(match.group(5))


def _year_args(sheet: str, first: int, last: int) -> list[str]:
    col, kind = YEAR_SOURCE[sheet]
    rng = f"'{sheet}'!${col}${first}:${col}${last}"
    if kind == "date":
        return [
            rng,
            f'">="&DATE({YEAR},1,1)',
            rng,
            f'"<"&DATE({YEAR + 1},1,1)',
        ]
    if kind == "year":
        return [rng, str(YEAR)]
    return [rng, f'"JS-{str(YEAR)[-2:]}*"']


def _yearize_call(name: str, body: str) -> str:
    args = _split_args(body)
    if not args:
        return f"{name}({body})"

    criteria_pos = {
        "COUNTIF": 0,
        "COUNTIFS": 0,
        "SUMIF": 0,
        "SUMIFS": 1,
        "COUNT": 0,
        "COUNTA": 0,
        "SUM": 0,
    }[name]
    if criteria_pos >= len(args):
        return f"{name}({body})"
    found = _range_sheet(args[criteria_pos])
    if not found:
        return f"{name}({body})"
    sheet, first, last = found
    year_args = _year_args(sheet, first, last)
    kind = YEAR_SOURCE[sheet][1]
    if (
        f"DATE({YEAR},1,1)" in body
        or (kind == "year" and re.search(rf"\b{YEAR}\b", body))
        or (kind == "settlement_id" and f"JS-{str(YEAR)[-2:]}*" in body)
    ):
        return f"{name}({body})"

    if name == "COUNTIF" and len(args) >= 2:
        return "COUNTIFS(" + ",".join(year_args + args[:2]) + ")"
    if name == "COUNTIFS":
        return "COUNTIFS(" + ",".join(year_args + args) + ")"
    if name == "SUMIF" and len(args) >= 2:
        sum_range = args[2] if len(args) >= 3 else args[0]
        return "SUMIFS(" + ",".join([sum_range] + year_args + args[:2]) + ")"
    if name == "SUMIFS" and len(args) >= 1:
        return "SUMIFS(" + ",".join([args[0]] + year_args + args[1:]) + ")"
    if name == "COUNT":
        return "COUNTIFS(" + ",".join(year_args + [args[0], '">=-1E+307"']) + ")"
    if name == "COUNTA":
        return "COUNTIFS(" + ",".join(year_args + [args[0], '"<>"']) + ")"
    if name == "SUM":
        return "SUMIFS(" + ",".join([args[0]] + year_args) + ")"
    return f"{name}({body})"


def yearize_formula(formula: str) -> str:
    """대시보드/대표보고 집계 함수에 2026년 기준 범위를 삽입한다."""
    out = formula
    for start, end, name, body in reversed(list(_call_spans(formula))):
        out = out[:start] + _yearize_call(name, body) + out[end:]
    return out


def issue_year_formula(row: int) -> str:
    """07 숨김 Q열: 원천 유형/행을 따라가 실제 업무 기준연도를 계산한다."""
    return (
        f'=IF($S{row}="","",'
        f'IF($R{row}=6,'
        f'IFERROR(2000+VALUE(MID(INDEX(\'17_문서대조현황\'!$A$5:$A$154,$S{row}),4,2)),""),'
        f'IFERROR(YEAR(CHOOSE($R{row},'
        f'INDEX(\'02_돌발AS접수\'!$D$5:$D$744,$S{row}),'
        f'INDEX(\'03_현장작업실적\'!$E$5:$E$322,$S{row}),'
        f'INDEX(\'04_정기점검\'!$D$5:$D$624,$S{row}),'
        f'INDEX(\'05_신규납품설치\'!$E$5:$E$54,$S{row}),'
        f'INDEX(\'06_거래서류청구수금\'!$L$5:$L$154,$S{row}))),"")))'
    )


def _top_index(rank: int) -> str:
    """n번째 2026년 고위험 행의 상대 위치.

    AGGREGATE 배열식은 이 통합문서의 호환 계산 모드에서 셀에 저장되면 #NAME?이
    발생하므로, INDEX로 배열을 강제한 MATCH와 이전 보조행 위치를 사용한다.
    """
    base = (
        "('07_불일치누락현황'!$Q$5:$Q$304=2026)"
        "*('07_불일치누락현황'!$V$5:$V$304=\"포함\")"
    )
    if rank == 1:
        return f"MATCH(1,INDEX({base},0),0)"
    positions = (
        "(ROW('07_불일치누락현황'!$Q$5:$Q$304)"
        "-ROW('07_불일치누락현황'!$Q$5)+1"
        f">$P{26 + rank})"
    )
    return f'IFERROR(MATCH(1,INDEX({base}*{positions},0),0),"")'


def top5_formula(rank: int) -> str:
    """2026년의 높음/치명적 이슈 중 앞선 항목을 대표보고 문장으로 만든다."""
    idx = f"$P{27 + rank}"
    ref = lambda col: f"INDEX('07_불일치누락현황'!${col}$5:${col}$304,{idx})"
    return (
        '=IFERROR("'
        + str(rank)
        + '. ["&'
        + ref("A")
        + '&"] "&'
        + ref("D")
        + '&" / "&'
        + ref("B")
        + '&" · "&'
        + ref("C")
        + '&"  →  "&'
        + ref("F")
        + '&"   (담당: "&IF('
        + "OR("
        + ref("L")
        + '="",'
        + ref("L")
        + '=0),"미배정",'
        + ref("L")
        + ')&")","")'
    )


def explicit_formula_updates() -> dict[str, str]:
    """SUMPRODUCT·고유개수·TOP5처럼 일반 변환만으로 부족한 수식."""
    s06 = "'06_거래서류청구수금'"
    s07 = "'07_불일치누락현황'"
    out = {
        "H17": (
            f"=SUMPRODUCT(({s06}!$L$5:$L$154>=DATE(2026,1,1))*"
            f"({s06}!$L$5:$L$154<DATE(2027,1,1))*"
            f"({s06}!$A$5:$A$154<>\"\")*({s06}!$AC$5:$AC$154>0))"
        ),
        "H19": (
            f"=SUMPRODUCT(({s06}!$L$5:$L$154>=DATE(2026,1,1))*"
            f"({s06}!$L$5:$L$154<DATE(2027,1,1))*"
            f"({s06}!$A$5:$A$154<>\"\")*({s06}!$AD$5:$AD$154>0))"
        ),
        "H21": (
            f"=SUMPRODUCT(({s06}!$L$5:$L$154>=DATE(2026,1,1))*"
            f"({s06}!$L$5:$L$154<DATE(2027,1,1))*"
            f"({s06}!$A$5:$A$154<>\"\")*({s06}!$AE$5:$AE$154<>0)*"
            f"({s06}!$Q$5:$Q$154<>0)*({s06}!$B$5:$B$154<>\"신규·납품·설치\"))"
        ),
        "K15": (
            f'=SUMPRODUCT(({s07}!$Q$5:$Q$304=2026)*'
            f'((ISNUMBER(SEARCH("누락",{s07}!$F$5:$F$304))+'
            f'ISNUMBER(SEARCH("미첨부",{s07}!$F$5:$F$304)))>0))'
        ),
        "K19": (
            f'=SUMPRODUCT(({s07}!$Q$5:$Q$304=2026)*'
            f'({s07}!$B$5:$B$304<>"")*({s07}!$L$5:$L$304=""))'
        ),
        "B62": (
            f'=SUMPRODUCT(({s07}!$Q$5:$Q$304=2026)*({s07}!$T$5:$T$304<>"")/'
            f'COUNTIFS({s07}!$T$5:$T$304,{s07}!$T$5:$T$304&"",'
            f'{s07}!$Q$5:$Q$304,{s07}!$Q$5:$Q$304))'
        ),
        "B63": (
            f'=SUMPRODUCT(({s07}!$Q$5:$Q$304=2026)*({s07}!$C$5:$C$304<>"")/'
            f'COUNTIFS({s07}!$C$5:$C$304,{s07}!$C$5:$C$304&"",'
            f'{s07}!$Q$5:$Q$304,{s07}!$Q$5:$Q$304))'
        ),
    }
    for rank, row in enumerate(range(28, 33), 1):
        out[f"A{row}"] = top5_formula(rank)
        out[f"P{row}"] = "=" + _top_index(rank)
    for row, offset in zip(range(78, 84), range(-5, 1)):
        out[f"A{row}"] = (
            f'=IF(YEAR(EDATE($B$4,{offset}))<>2026,"",'
            f'TEXT(EDATE($B$4,{offset}),"yyyy-mm"))'
        )
    return out


def _formula_changes(xml: str) -> dict[str, tuple[str, str]]:
    changes = {}
    for row, col, formula, _start, _end in parse_cells(xml):
        new = yearize_formula(formula)
        if new != formula:
            changes[f"{col}{row}"] = ("formula", "=" + new.lstrip("="))
    return changes


def _cell_style(cell_xml: str | None, fallback: str | None = None) -> str:
    match = re.search(r'\ss="(\d+)"', cell_xml or "")
    return match.group(1) if match else (fallback or "")


def _new_cell(ref: str, kind: str, value: str, style: str = "") -> str:
    style_attr = f' s="{style}"' if style else ""
    if kind == "formula":
        return f'<c r="{ref}"{style_attr}><f>{esc(value.lstrip("="))}</f><v/></c>'
    return (
        f'<c r="{ref}"{style_attr} t="inlineStr"><is><t xml:space="preserve">'
        f"{esc(value)}</t></is></c>"
    )


def patch_cells(xml: str, changes: dict[str, tuple[str, str]]) -> str:
    """한 시트의 여러 셀을 행 단위로 안전하게 교체/삽입한다."""
    by_row = defaultdict(dict)
    for ref, change in changes.items():
        match = re.fullmatch(r"([A-Z]{1,3})(\d+)", ref)
        if not match:
            raise ValueError(f"잘못된 셀 주소: {ref}")
        by_row[int(match.group(2))][match.group(1)] = change

    fallback_style = {}
    for match in re.finditer(r'<c r="([A-Z]{1,3})\d+"[^>]*>', xml):
        col = match.group(1)
        if col not in fallback_style:
            fallback_style[col] = _cell_style(match.group(0))

    # self-closing spacer row(<row .../>)를 여는 행으로 오인하면 다음 행까지 삼킨다.
    row_rx = re.compile(r'(<row r="(\d+)"[^>]*(?<!/)>)(.*?)(</row>)', re.S)
    seen_rows = set()

    def repl_row(match):
        row = int(match.group(2))
        if row not in by_row:
            return match.group(0)
        seen_rows.add(row)
        head, body, tail = match.group(1), match.group(3), match.group(4)
        # `<c .../>` 분기와 `<c ...>...</c>` 분기를 분리한다. 한 패턴으로 쓰면
        # 첫 self-closing 셀이 뒤의 일반 셀까지 삼켜 중복 셀을 만들 수 있다.
        cell_rx = re.compile(
            r'<c r="([A-Z]{1,3})' + str(row) + r'"[^>]*?/>'
            r'|<c r="([A-Z]{1,3})' + str(row) + r'"[^>]*(?<!/)>.*?</c>',
            re.S,
        )
        cells = {(m.group(1) or m.group(2)): m.group(0) for m in cell_rx.finditer(body)}
        for col, (kind, value) in by_row[row].items():
            old = cells.get(col)
            style = _cell_style(old, fallback_style.get(col))
            cells[col] = _new_cell(f"{col}{row}", kind, value, style)
        return head + "".join(cells[c] for c in sorted(cells, key=col_num)) + tail

    out = row_rx.sub(repl_row, xml)
    missing = set(by_row) - seen_rows
    if missing:
        raise AssertionError(f"행 XML 없음: {sorted(missing)[:10]}")
    return out


def ensure_issue_year_hidden(xml: str) -> str:
    if re.search(r'<col min="17" max="17"[^>]*hidden="1"', xml):
        return xml
    marker = '<col min="18" max="19"'
    hidden = '<col min="17" max="17" width="12" hidden="1" customWidth="1"/>'
    if marker in xml:
        return xml.replace(marker, hidden + marker, 1)
    return xml.replace("</cols>", hidden + "</cols>", 1)


def ensure_dashboard_helper_hidden(xml: str) -> str:
    """TOP5 행번호 보조열 P를 숨기고 사용 범위를 P열까지 넓힌다."""
    xml = re.sub(
        r'<dimension ref="A1:O(\d+)"\s*/>',
        r'<dimension ref="A1:P\1"/>',
        xml,
        count=1,
    )
    if not re.search(r'<col min="16" max="16"[^>]*hidden="1"', xml):
        hidden = '<col min="16" max="16" width="3" hidden="1" customWidth="1"/>'
        xml = xml.replace("</cols>", hidden + "</cols>", 1)
    return xml


def build_changes(path: str):
    archive = zipfile.ZipFile(path)
    dash_part = sheet_xml_path(archive, DASH)
    report_part = sheet_xml_path(archive, REPORT)
    issue_part = sheet_xml_path(archive, ISSUES)
    dash_xml = archive.read(dash_part).decode("utf-8")
    report_xml = archive.read(report_part).decode("utf-8")
    issue_xml = archive.read(issue_part).decode("utf-8")

    dash_changes = _formula_changes(dash_xml)
    dash_changes.update({k: ("formula", v) for k, v in explicit_formula_updates().items()})
    dash_changes.update({
        "A2": (
            "text",
            "목적: 2026년 쿠팡 업무(돌발AS·정기점검·신규납품설치)의 처리·서류·청구·수금 "
            "상태만 집계합니다. 2025년 원본 행은 보존되지만 모든 통계·TOP5에서는 제외됩니다. "
            "노란색 입력칸(B3·B4·B5·B6)만 입력하고 나머지는 자동 계산합니다.",
        ),
        "D47": ("text", "2026 누적"),
        "G47": ("text", "리스크 · 잔여 (2026 관리)"),
        "A25": ("text", "▶ 문제 행 수 (2026 · 중복 포함)"),
        "A64": ("text", "③ 문제 행 수"),
        "A73": ("text", "▶ 2026 월별 집계 현황"),
        "P27": ("text", "TOP5 원천행(자동·숨김)"),
    })
    report_changes = _formula_changes(report_xml)
    report_changes.update({
        "A2": (
            "text",
            "2026년 업무의 집계기준일 하루 실적과 현재 리스크·잔여만 보고합니다. "
            "2025년 원본 행은 통계와 TOP5에서 제외됩니다.",
        ),
        "A25": ("text", "문제 프로젝트 / 문제 행"),
    })

    issue_changes = {"Q4": ("text", "업무기준연도(자동·숨김)")}
    issue_changes.update({
        f"Q{row}": ("formula", issue_year_formula(row))
        for row in range(5, 305)
    })

    changed = {
        dash_part: ensure_dashboard_helper_hidden(
            patch_cells(dash_xml, dash_changes)
        ).encode("utf-8"),
        report_part: patch_cells(report_xml, report_changes).encode("utf-8"),
        issue_part: ensure_issue_year_hidden(
            patch_cells(issue_xml, issue_changes)
        ).encode("utf-8"),
    }
    archive.close()
    return changed, {
        "dashboard_formulas": sum(1 for kind, _ in dash_changes.values() if kind == "formula"),
        "report_formulas": sum(1 for kind, _ in report_changes.values() if kind == "formula"),
        "issue_year_formulas": 300,
        "text_labels": sum(1 for kind, _ in dash_changes.values() if kind == "text") + 1,
    }


def apply(path: str, changed: dict[str, bytes]) -> str:
    match = re.search(r"_v(\d+)\.xlsx$", path)
    if not match:
        raise ValueError(f"버전 파일명이 아님: {path}")
    out = re.sub(r"_v\d+\.xlsx$", f"_v{int(match.group(1)) + 1}.xlsx", path)
    if os.path.exists(out):
        raise FileExistsError(os.path.basename(out))
    temp = out[:-5] + ".tmp.xlsx"

    source = zipfile.ZipFile(path)
    target = zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED)
    for item in source.infolist():
        data = changed.get(item.filename, source.read(item.filename))
        recalc = force_recalc_part(item.filename, data)
        if recalc is not None:
            target.writestr(item.filename, recalc)
    target.close()
    source.close()

    check = zipfile.ZipFile(temp)
    assert check.testzip() is None, "ZIP 무결성 실패"
    check.close()

    import openpyxl

    wb = openpyxl.load_workbook(temp, read_only=True, data_only=False)
    assert "2026" in str(wb[DASH]["A2"].value)
    assert str(wb[ISSUES]["Q5"].value or "").startswith("=")
    for cell in ("A28", "A29", "A30", "A31", "A32"):
        formula = str(wb[DASH][cell].value or "")
        assert "$P" in formula and "UJ25" not in formula
    for cell in ("P28", "P29", "P30", "P31", "P32"):
        formula = str(wb[DASH][cell].value or "")
        assert "$Q$5:$Q$304=2026" in formula and "$V$5:$V$304=\"포함\"" in formula
    assert "DATE(2026,1,1)" in str(wb[DASH]["H17"].value)
    assert "2026" in str(wb[DASH]["B62"].value)
    wb.close()
    os.replace(temp, out)
    return out


def main():
    path = latest_master()[0]
    changed, counts = build_changes(path)
    print(f"원본: {os.path.basename(path)}")
    print(
        "반영 계획: 대시보드 수식 {dashboard_formulas} · 대표보고 수식 {report_formulas} · "
        "이슈 연도 보조수식 {issue_year_formulas} · 안내/라벨 {text_labels}".format(**counts)
    )
    print("원본 2025년 행은 삭제·변경하지 않습니다.")
    if "--apply" not in sys.argv:
        print("반영하려면: python stats_2026.py --apply")
        return
    from claim_guard import require

    require("ledger", "stats_2026")
    out = apply(path, changed)
    print(f"생성: {os.path.basename(out)}")


if __name__ == "__main__":
    main()
