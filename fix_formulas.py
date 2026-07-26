# -*- coding: utf-8 -*-
"""
fix_formulas.py — 관리대장 안의 **깨진 수식·좁아진 범위**를 고친다
================================================================================
엑셀에서 계속 보이던 오류의 실체(2026-07-26 조사):

  1) 음수 행 참조 213개   `COUNT($AO$4:AO-120)`  → #N/A
     행을 옮기면서 상대 행번호가 0 밑으로 내려갔다. 행 재배치·행 확장의 부작용.
  2) 오류가 수식으로 굳음 273개  `=#N/A`
     원래 수식이 통째로 사라지고 오류값만 남았다. 그 열의 정상 수식으로 되살린다.
  3) 집계 범위가 옛날 크기 그대로
     02시트는 344행까지 늘렸는데 대시보드·대표보고 수식은 아직 `$A$5:$A$154`만 본다
     → **최근에 넣은 건이 통계에서 통째로 빠진다**(오류 표시는 없지만 숫자가 틀린다).

고치는 방법
  · 같은 열의 **정상 수식 하나**를 본으로 삼아, 행 오프셋(0=같은 행, -1=윗행)을 유지한 채
    깨진 행에 다시 써 넣는다. 사람이 손으로 짠 수식 구조를 그대로 보존한다.
  · 시트별 실제 마지막 행을 구해 모든 참조 범위의 끝행을 거기에 맞춘다.
  · 전부 zip 패치 — openpyxl save 금지(차트·도형 파괴).

실행
  python fix_formulas.py             # 무엇을 고칠지 미리보기
  python fix_formulas.py --apply     # vN+1 생성
"""
import sys, os, re, zipfile, collections

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from fix_workbook import iter_tags, col_num          # 자기닫힘 태그 안전 파서

NEG_RE = re.compile(r"[A-Z]{1,3}-\d")
DEAD_RE = re.compile(r"^#[A-Z/]+[!?]?$")
# 상대 행번호(=$가 안 붙은 행). 열에 $가 붙어도 행에 없으면 상대다.
REL_RE = re.compile(r"(\$?[A-Z]{1,3})(\d+)(?![\d(])")
XML_ESC = {"&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'"}


def unesc(s):
    for k, v in XML_ESC.items():
        s = s.replace(k, v)
    return s


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def sheet_map(z):
    """시트이름 → 시트 XML 경로"""
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    out = {}
    for m in re.finditer(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
        t = rid2t.get(m.group(2), "")
        if t.startswith("worksheets/"):
            out[unesc(m.group(1))] = "xl/" + t
    return out


def parse_cells(xml):
    """[(row, colLetters, formula, 셀시작, 셀끝)] — 수식이 있는 셀만"""
    out = []
    for rs, re_, rtag, rinner in iter_tags(xml, "row"):
        if rinner is None:
            continue
        rm = re.search(r'r="(\d+)"', rtag)
        if not rm:
            continue
        rn = int(rm.group(1))
        # rinner 는 <row ...> 여는 태그 **다음**부터다. 이 길이를 안 더하면
        # 셀 위치가 통째로 어긋나 엉뚱한 자리를 덮어쓴다.
        base = rs + len(rtag)
        for cs, ce, ctag, cinner in iter_tags(rinner, "c"):
            if not cinner or "<f" not in cinner:
                continue
            ref = re.search(r'r="([A-Z]{1,3})(\d+)"', ctag)
            if not ref:
                continue
            fm = None
            for _, _, ftag, finner in iter_tags(cinner, "f"):
                fm = finner or ""
                break
            if fm is None:
                continue
            out.append((rn, ref.group(1), unesc(fm), base + cs, base + ce))
    return out


def normalize(formula, row):
    """수식 안의 상대 행번호를 '이 행 기준 오프셋'으로 바꾼다 → 다른 행에 재사용 가능"""
    def rep(m):
        col, n = m.group(1), int(m.group(2))
        if col.endswith("$"):            # 행에 $ 없으니 여기 올 일은 없다
            return m.group(0)
        return f"{col}\x00{n - row}\x00"
    return REL_RE.sub(rep, formula)


def instantiate(tmpl, row):
    return re.sub(r"\x00(-?\d+)\x00", lambda m: str(row + int(m.group(1))), tmpl)


def is_broken(f):
    return bool(NEG_RE.search(f)) or bool(DEAD_RE.match(f.strip())) or "#REF!" in f


def templates(cells):
    """열 → 그 열에서 가장 흔한 정상 수식 본"""
    per = collections.defaultdict(collections.Counter)
    for rn, col, f, _, _ in cells:
        # 공유수식 참조 셀(<f t="shared" si="3"/>)은 본문이 비어 있다.
        # 이걸 세면 빈 문자열이 최다 득표로 뽑혀 본이 통째로 비어버린다.
        if not f.strip() or is_broken(f):
            continue
        t = normalize(f, rn)
        # 오프셋이 -1..0 을 벗어나면 그 행만의 특수 수식일 가능성 → 본으로 쓰지 않는다
        if any(abs(int(o)) > 1 for o in re.findall(r"\x00(-?\d+)\x00", t)):
            continue
        per[col][t] += 1
    return {c: cnt.most_common(1)[0][0] for c, cnt in per.items() if cnt}


def sheet_last_row(xml):
    rows = [int(m) for m in re.findall(r'<row r="(\d+)"', xml)]
    return max(rows) if rows else 4


def fix_sheet(xml):
    """(새 XML, 고친 수식 수)"""
    cells = parse_cells(xml)
    tmpl = templates(cells)
    edits = []
    for rn, col, f, s, e in cells:
        if not is_broken(f):
            continue
        t = tmpl.get(col)
        if not t or not t.strip():
            continue
        edits.append((s, e, instantiate(t, rn)))
    if not edits:
        return xml, 0
    out, last = [], 0
    for s, e, newf in sorted(edits):
        chunk = xml[s:e]
        # <f ...>본문</f> 의 본문만 갈아끼운다(자기닫힘 <f/>는 대상이 아니다)
        fixed = re.sub(r"(<f(?:\s[^>]*)?>)(.*?)(</f>)",
                       lambda m: m.group(1) + esc(newf) + m.group(3), chunk, count=1, flags=re.S)
        # 캐시값은 지운다 — 옛 오류값이 남아 있으면 엑셀이 그대로 보여준다
        fixed = re.sub(r"<v[^>]*>.*?</v>|<v\s*/>", "<v/>", fixed, count=1, flags=re.S)
        fixed = re.sub(r'\st="e"', "", fixed, count=1)
        out.append(xml[last:s]); out.append(fixed); last = e
    out.append(xml[last:])
    return "".join(out), len(edits)


def widen(xml, ends, self_name=None, self_end=None):
    """참조 범위의 끝행을 시트의 실제 마지막 행까지 넓힌다"""
    n = [0]

    def cross(m):
        sh, a, b, old = m.group(1), m.group(2), m.group(3), int(m.group(4))
        new = ends.get(unesc(sh))
        if not new or new <= old:
            return m.group(0)
        n[0] += 1
        return f"'{sh}'!{a}:{b}${new}"

    xml = re.sub(r"'([^']+)'!(\$[A-Z]{1,3}\$\d+):(\$[A-Z]{1,3})\$(\d+)", cross, xml)

    if self_name and self_end:
        old_end = ends.get(self_name)          # 같은 시트 안에서 쓰는 옛 끝행
        if old_end:
            def same(m):
                a, b, old = m.group(1), m.group(2), int(m.group(3))
                if old != self_end or old >= old_end:
                    return m.group(0)
                n[0] += 1
                return f"{a}:{b}${old_end}"
            xml = re.sub(r"(?<!!)(\$[A-Z]{1,3}\$\d+):(\$[A-Z]{1,3})\$(\d+)", same, xml)
    return xml, n[0]


def main():
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    z = zipfile.ZipFile(master)
    smap = sheet_map(z)
    raw = {name: z.read(p).decode("utf-8") for name, p in smap.items()}
    ends = {name: sheet_last_row(x) for name, x in raw.items()}

    # 지금 수식들이 쓰고 있는 '옛 끝행'을 시트별로 알아낸다(가장 많이 나오는 값)
    old_ends = collections.defaultdict(collections.Counter)
    for x in raw.values():
        for m in re.finditer(r"'([^']+)'!\$[A-Z]{1,3}\$5:\$[A-Z]{1,3}\$(\d+)", x):
            old_ends[unesc(m.group(1))][int(m.group(2))] += 1
    old_ends = {k: v.most_common(1)[0][0] for k, v in old_ends.items()}

    print("시트별 실제 마지막 행 / 수식이 보고 있는 끝행:")
    for name in sorted(ends):
        oe = old_ends.get(name)
        if oe and oe < ends[name]:
            print(f"  {name}: 실제 {ends[name]} / 수식 {oe}  ← {ends[name]-oe}행이 집계에서 빠짐")

    patched, nfix, nwide = {}, 0, 0
    for name, path in smap.items():
        x = raw[name]
        y, k = fix_sheet(x)
        y, w = widen(y, ends, name, old_ends.get(name))
        if k or w:
            patched[path] = y.encode("utf-8")
            nfix += k
            nwide += w
    z.close()
    print(f"\n깨진 수식 복구 {nfix}개 · 범위 확장 {nwide}곳 (시트 {len(patched)}개)")

    if "--apply" not in sys.argv:
        print("반영하려면: python fix_formulas.py --apply")
        return
    if not patched:
        print("고칠 것 없음 — 새 버전 미생성")
        return
    mv = re.search(r"_v(\d+)\.xlsx$", master)
    dst = re.sub(r"_v\d+\.xlsx$", f"_v{int(mv.group(1))+1}.xlsx", master)
    if os.path.exists(dst):
        sys.exit(f"{os.path.basename(dst)} 이미 존재 — 중단")
    zin = zipfile.ZipFile(master)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        zout.writestr(it.filename, patched.get(it.filename) or zin.read(it.filename))
    zout.close(); zin.close()

    # 검증: zip 무결성 + 음수행이 남아 있지 않은지
    zc = zipfile.ZipFile(dst)
    assert zc.testzip() is None, "zip 무결성 실패"
    # ★ XML 전체를 훑으면 'AS-2607-001' 같은 **접수ID 문자열**까지 세어 버린다.
    #   반드시 수식 본문만 본다(2026-07-26에 1,084개로 잘못 세었던 자리).
    left = 0
    for p in smap.values():
        for _, _, f, _, _ in parse_cells(zc.read(p).decode("utf-8")):
            if is_broken(f):
                left += 1
    zc.close()
    print(f"생성: {os.path.basename(dst)} · 남은 음수행 참조 {left}개")


if __name__ == "__main__":
    main()
