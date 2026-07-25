# -*- coding: utf-8 -*-
"""
fix_workbook.py — 관리대장 내부 무결성 검사·복구 (엑셀 "내용에 문제가 있습니다" 방지)
================================================================================
엑셀이 파일을 열 때 복구 대화상자를 띄우는 원인을 찾아 고친다.

실제로 겪은 사고
  수식 셀에 `t="str"`(캐시값이 문자열이다)만 써 놓고 `<v>`를 넣지 않았다.
  엑셀 자신이 쓴 셀은 예외 없이 `<f>` 다음에 `<v>`가 있다. 선언과 내용이 어긋나면
  엑셀은 파일을 손상으로 판단한다. (expand_rows가 만든 행 1,573개가 여기 해당)

검사 항목
  1. zip 무결성 · 모든 XML 파트 파싱
  2. 수식 셀의 t 속성 ↔ 캐시값(<v>) 정합성      ← 복구 대화상자 주원인
  3. 행/셀 순서·중복, 셀 참조와 행 번호 불일치
  4. 시트 ↔ rels ↔ Content_Types 연결
  5. 스타일 인덱스 범위 초과
  6. calcChain 잔재(파트는 지웠는데 참조가 남은 경우)

사용
  python fix_workbook.py                 # 검사만 (최신본)
  python fix_workbook.py --apply         # 문제 수정 → vN+1 생성
  python fix_workbook.py --file 경로.xlsx # 특정 파일 검사
"""
import sys, os, re, zipfile
import xml.etree.ElementTree as ET

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
# ★ 캐시값 존재 판정은 반드시 이 정규식으로. '<v>' 문자열만 찾으면
#   <v xml:space="preserve">…</v> 를 놓쳐 <v>를 하나 더 넣게 되고,
#   한 셀에 <v>가 2개면 엑셀이 그 시트를 통째로 버린다(2026-07-26 실사고, 313셀).
_V_RE = re.compile(r"<v[\s/>]")


def iter_tags(xml, name):
    """자기닫힘 태그를 게으른 정규식으로 잡으면 다음 태그까지 삼킨다 → 2단계 파싱"""
    i, open_t = 0, "<" + name
    while True:
        s = xml.find(open_t, i)
        if s < 0:
            return
        nxt = xml[s + len(open_t):s + len(open_t) + 1]
        if nxt not in (" ", ">", "/"):
            i = s + len(open_t)
            continue
        e = xml.find(">", s)
        if e < 0:
            return
        tag = xml[s:e + 1]
        if tag.endswith("/>"):
            yield s, e + 1, tag, None
            i = e + 1
        else:
            c = xml.find("</" + name + ">", e)
            if c < 0:
                return
            yield s, c + len(name) + 3, tag, xml[e + 1:c]
            i = c + len(name) + 3


def col_num(ref):
    m = re.match(r"([A-Z]+)", ref)
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n


# ── 검사 ─────────────────────────────────────────────────────────────
def scan_sheet(xml, label, style_max):
    """(문제목록, 고칠_수식셀수)"""
    bad, fixable = [], 0
    prev_r, rows_seen = 0, set()
    for _, _, rtag, rinner in iter_tags(xml, "row"):
        if rinner is None:
            continue
        rm = re.search(r'r="(\d+)"', rtag)
        if not rm:
            continue
        r = int(rm.group(1))
        if r in rows_seen:
            bad.append(f"{label}: 행 {r} 중복")
        rows_seen.add(r)
        if r <= prev_r:
            bad.append(f"{label}: 행 순서 역전 {prev_r}→{r}")
        prev_r = r
        pc, cells_seen = 0, set()
        for _, _, ctag, cinner in iter_tags(rinner, "c"):
            ref = (re.search(r'r="([A-Z]+\d+)"', ctag) or [None])
            ref = ref.group(1) if hasattr(ref, "group") else None
            if not ref:
                bad.append(f"{label}: 셀 참조 없음")
                continue
            if int(re.match(r"[A-Z]+(\d+)", ref).group(1)) != r:
                bad.append(f"{label}: {ref} 가 행 {r} 안에 있음")
            if ref in cells_seen:
                bad.append(f"{label}: 셀 {ref} 중복")
            cells_seen.add(ref)
            cn = col_num(ref)
            if cn <= pc:
                bad.append(f"{label}: 열 순서 역전 {ref}")
            pc = cn
            sm = re.search(r'\ss="(\d+)"', ctag)
            if sm and style_max is not None and int(sm.group(1)) >= style_max:
                bad.append(f"{label}: {ref} 스타일 인덱스 {sm.group(1)} 범위 초과")
            if cinner and "<f" in cinner:
                has_v = bool(_V_RE.search(cinner))
                t = re.search(r'\st="([^"]+)"', ctag)
                if t and not has_v:
                    fixable += 1     # ★ 복구 대화상자 주원인
    return bad, fixable


def add_missing_v(xml):
    """수식 셀에 캐시값이 없으면 <v/>를 넣어 t 선언과 맞춘다(엑셀 자체 출력과 동일한 형태)."""
    out, last, n = [], 0, 0
    for s, e, ctag, cinner in iter_tags(xml, "c"):
        if cinner is None or "<f" not in cinner:
            continue
        if _V_RE.search(cinner):        # <v>, <v/>, <v xml:space="preserve"> 전부 포함
            continue
        if not re.search(r'\st="[^"]+"', ctag):
            continue
        out.append(xml[last:s])
        out.append(ctag + cinner + "<v/></c>")
        last = e
        n += 1
    out.append(xml[last:])
    return "".join(out), n


def remove_dup_v(xml):
    """한 셀에 <v>가 2개 이상이면 뒤에 붙은 빈 <v/>를 걷어낸다(잘못 넣은 것 되돌리기)."""
    out, last, n = [], 0, 0
    for s, e, ctag, cinner in iter_tags(xml, "c"):
        if cinner is None or len(_V_RE.findall(cinner)) < 2:
            continue
        fixed = re.sub(r"<v\s*/>(?=(?:<[^v]|$))", "", cinner)   # 값 없는 <v/>만 제거
        if fixed == cinner:
            fixed = re.sub(r"<v\s*/>", "", cinner, count=1)
        out.append(xml[last:s])
        out.append(ctag + fixed + "</c>")
        last = e
        n += 1
    out.append(xml[last:])
    return "".join(out), n


def count_dup_v(xml):
    return sum(1 for s, e, ct, ci in iter_tags(xml, "c")
               if ci and len(_V_RE.findall(ci)) > 1)


def check(path, verbose=True):
    problems, fix_total, dup_total = [], 0, [0]
    z = zipfile.ZipFile(path)
    if z.testzip() is not None:
        problems.append("zip 무결성 실패")
    names = set(z.namelist())
    ct = z.read("[Content_Types].xml").decode("utf-8")
    if "calcChain" in ct and "xl/calcChain.xml" not in names:
        problems.append("calcChain 참조는 남았는데 파트가 없음")
    style_max = None
    if "xl/styles.xml" in names:
        m = re.search(r'<cellXfs count="(\d+)"', z.read("xl/styles.xml").decode("utf-8"))
        style_max = int(m.group(1)) if m else None
    for n in sorted(names):
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:
                problems.append(f"{n}: XML 파싱 실패 {str(e)[:50]}")
    for n in sorted(names):
        if not re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            continue
        x = z.read(n).decode("utf-8")
        b, f = scan_sheet(x, n.split("/")[-1], style_max)
        problems += b
        fix_total += f
        d = count_dup_v(x)
        if d:
            problems.append(f"{n.split('/')[-1]}: 한 셀에 <v>가 2개인 셀 {d}개 "
                            f"— 엑셀이 시트를 통째로 버린다")
            dup_total[0] += d
    z.close()
    if verbose:
        print(f"검사: {os.path.basename(path)}")
        print(f"  캐시값 없는 수식 셀(t 선언만 있음): {fix_total}개"
              + ("   ★ 엑셀 복구 대화상자의 원인" if fix_total else "  ✅"))
        if problems:
            for p in problems[:10]:
                print(f"  ★ {p}")
            if len(problems) > 10:
                print(f"  … 외 {len(problems)-10}건")
        else:
            print("  구조 검사(행·셀 순서·중복·스타일·연결) 이상 없음 ✅")
    return problems, fix_total


def repair(src, dst):
    zin = zipfile.ZipFile(src)
    patched, total = {}, 0
    for n in zin.namelist():
        if not re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            continue
        x = zin.read(n).decode("utf-8")
        y, d = remove_dup_v(x)          # 잘못 넣은 <v/> 먼저 걷어내고
        y, cnt = add_missing_v(y)       # 진짜 빠진 곳만 채운다
        if cnt or d:
            patched[n] = y.encode("utf-8")
            total += cnt + d
    if not total:
        zin.close()
        return 0, 0
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = patched.get(it.filename) or zin.read(it.filename)
        zout.writestr(it.filename, data)
    zout.close()
    zin.close()
    return total, len(patched)


def main():
    args = sys.argv[1:]
    if "--file" in args:
        master = args[args.index("--file") + 1]
    else:
        from ecount_reconcile import load_config, resolve_master
        master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    problems, fixable = check(master)
    if "--apply" not in args:
        if fixable or problems:
            print("\n수정하려면: python fix_workbook.py --apply")
        return
    if not fixable:
        print("고칠 항목 없음 — 새 버전 미생성")
        return
    m = re.search(r"_v(\d+)\.xlsx$", master)
    dst = re.sub(r"_v\d+\.xlsx$", f"_v{int(m.group(1))+1}.xlsx", master)
    if os.path.exists(dst):
        sys.exit(f"{os.path.basename(dst)} 이미 존재 — 중단")
    n, parts = repair(master, dst)
    print(f"\n복구: 수식 셀 {n}개에 캐시값 추가 (시트 {parts}개) → {os.path.basename(dst)}")
    p2, f2 = check(dst)
    if f2 == 0 and not p2:
        print("재검사 통과 ✅ — 엑셀에서 복구 대화상자 없이 열립니다")
    else:
        print("★ 재검사에서 문제 남음 — 확인 필요")


if __name__ == "__main__":
    main()
