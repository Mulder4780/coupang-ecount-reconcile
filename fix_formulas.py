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

from fix_workbook import iter_tags, col_num, _V_RE   # 자기닫힘 태그 안전 파서 · <v> 판정

NEG_RE = re.compile(r"[A-Z]{1,3}-\d")
DEAD_RE = re.compile(r"^#[A-Z/]+[!?]?$")
# 상대 행번호(=$가 안 붙은 행). 열에 $가 붙어도 행에 없으면 상대다.
REL_RE = re.compile(r"(\$?[A-Z]{1,3})(\d+)(?![\d(])")
XML_ESC = {"&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'"}
# 사람이 입력하는 열과 달리 아래 열은 값이 보이더라도 수식이 진실의 원천이다.
# ID(A열)는 행 재배치 뒤 참조 보호를 위해 확정값일 수 있으므로 절대 포함하지 않는다.
FORMULA_OWNED = {
    "02_돌발AS접수": ("F", "G", "M", "AK", "AL", "AN"),
}
KNOWN_TEXT_FIXES = {
    "권오절": "권오철",
    "권오처르": "권오철",
}


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


def blank_value(ctag, cinner, sst):
    """수식 없이 **빈 문자열로 굳어 있는** 셀인가.

    엑셀에서 수식이 통째로 사라지고 빈 값만 남으면 화면에는 그냥 빈칸으로 보인다.
    사람이 기사 이름을 넣어도 상태가 안 바뀌는데 원인을 알 길이 없다
    (2026-07-27 02시트 기사배정상태 89칸이 이 상태였다).
    '배정완료' 같은 **의미 있는 값은 손대지 않는다** — 사람이 고쳐 둔 것일 수 있다.
    """
    if not cinner or "<f" in cinner:
        return False
    if 't="inlineStr"' in ctag or "<is>" in cinner:
        return not "".join(re.findall(r"<t[^>]*>(.*?)</t>", cinner, re.S)).strip()
    m = re.search(r"<v[^>]*>(.*?)</v>", cinner, re.S)
    if not m:
        return True
    raw = m.group(1).strip()
    if 't="s"' in ctag and raw.isdigit():
        return sst.get(int(raw), "") == ""
    return raw == ""


def shared_strings(z):
    """공유문자열 인덱스 → 값 (빈 문자열 판정에 쓴다)"""
    if "xl/sharedStrings.xml" not in z.namelist():
        return {}
    x = z.read("xl/sharedStrings.xml").decode("utf-8")
    out = {}
    for i, si in enumerate(re.findall(r"<si>(.*?)</si>", x, re.S)):
        out[i] = "".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)).strip()
    return out


def fill_missing(xml, min_share=0.5, sst=None):
    sst = sst or {}
    """**수식 열인데 빈칸인 셀**에 그 열의 수식을 다시 넣는다.

    왜 필요한가 — 행을 늘리거나 백필하면서 어떤 열은 수식이 안 따라붙는다.
    그 행은 사람이 값을 입력해도 상태가 안 바뀐다("기사 넣었는데 배정완료로 안 변함",
    2026-07-27 류지영 매니저 보고. 02시트 기사배정상태 89행이 빈칸이었다).
    화면에는 그냥 비어 보여서 아무도 원인을 모른다.

    판단 기준: 그 열의 데이터 행 중 **절반 이상이 수식**이면 수식 열로 본다.
    사람이 직접 적는 열(방문일정상태처럼 수식이 아예 없는 열)은 건드리지 않는다.
    이미 값이 들어 있는 칸도 그대로 둔다 — 사람이 손으로 고쳐 둔 것일 수 있다.
    """
    cells = parse_cells(xml)
    tmpl = templates(cells)
    if not tmpl:
        return xml, 0
    has_f = collections.Counter(c[1] for c in cells)
    style = {}                                   # 열 → 수식 셀의 스타일(서식 유지용)
    for rs, _e, rtag, rinner in iter_tags(xml, "row"):
        if rinner is None:
            continue
        for _cs, _ce, ctag, cinner in iter_tags(rinner, "c"):
            ref = re.search(r'r="([A-Z]{1,3})\d+"', ctag)
            if ref and cinner and "<f" in cinner and ref.group(1) not in style:
                sm = re.search(r'\ss="(\d+)"', ctag)
                style[ref.group(1)] = sm.group(1) if sm else None

    n_rows = collections.Counter()               # 열별 데이터 행 수(존재하는 셀 기준)
    for rs, _e, rtag, rinner in iter_tags(xml, "row"):
        if rinner is None:
            continue
        rm = re.search(r'r="(\d+)"', rtag)
        if not rm or int(rm.group(1)) <= 4:
            continue
        for _cs, _ce, ctag, _ci in iter_tags(rinner, "c"):
            ref = re.search(r'r="([A-Z]{1,3})\d+"', ctag)
            if ref:
                n_rows[ref.group(1)] += 1

    want = {c for c in tmpl
            if tmpl[c].strip() and n_rows[c] and has_f[c] / n_rows[c] >= min_share}
    if not want:
        return xml, 0

    def new_cell(col, rn):
        s = style.get(col)
        return (f'<c r="{col}{rn}"' + (f' s="{s}"' if s else "") + ">"
                + "<f>" + esc(instantiate(tmpl[col], rn)) + "</f><v/></c>")

    out, last, n = [], 0, 0
    for rs, re_, rtag, rinner in iter_tags(xml, "row"):
        if rinner is None:
            continue
        rm = re.search(r'r="(\d+)"', rtag)
        if not rm:
            continue
        rn = int(rm.group(1))
        if rn <= 4:
            continue
        pieces, here, changed = [], {}, False
        for cs, ce, ctag, cinner in iter_tags(rinner, "c"):
            ref = re.search(r'r="([A-Z]{1,3})\d+"', ctag)
            if ref:
                here[ref.group(1)] = (cs, ce, ctag, cinner)
        # ★ 엑셀은 빈 셀을 아예 안 적는다. 그래서 '비어 있는 칸'은 <c>가 없는 경우가 대부분이고,
        #   기존 <c>만 훑으면 89칸 중 8칸밖에 못 찾는다. 없는 셀은 만들어 끼워 넣는다.
        missing = [c for c in want
                   if c not in here or (here[c][3] is None)
                   or ("<f" not in (here[c][3] or "") and not _V_RE.search(here[c][3] or ""))
                   or blank_value(here[c][2], here[c][3], sst)]
        if not missing:
            continue
        body, cur = [], 0
        allcols = sorted(set(list(here) + missing), key=col_num)
        miss_set = set(missing)
        for col in allcols:
            # ★ 채울 대상(missing)으로 이미 판정된 칸을 여기서 '값이 있으니 그대로 둔다'고
            #   되살리면 아무것도 안 고쳐진다(빈 문자열로 굳은 칸이 딱 이 경우다).
            if col not in miss_set and col in here and here[col][3] is not None and (
                    "<f" in here[col][3] or _V_RE.search(here[col][3])):
                cs, ce, _t, _i = here[col]
                body.append(rinner[cs:ce])
            elif col in want:
                body.append(new_cell(col, rn)); n += 1
            elif col in here:
                cs, ce, _t, _i = here[col]
                body.append(rinner[cs:ce])
        out.append(xml[last:rs])
        out.append(rtag + "".join(body) + "</row>")
        last = re_
    out.append(xml[last:])
    return ("".join(out), n) if n else (xml, 0)


def restore_owned_formulas(xml, columns, key_col="B", sst=None):
    """수식 전용 열에 굳어버린 정적 값을 정상 수식으로 되돌린다.

    2026-07-28 실제 사례: 02시트 L523에 김준형을 골랐지만 M523이 수식이 아닌
    문자열 '미배정'으로 저장돼 상태가 바뀌지 않았다. 빈칸 복구만으로는 못 잡으므로
    명시한 수식 전용 열만 복구한다.
    """
    sst = sst or {}
    tmpl = templates(parse_cells(xml))
    wanted = {c for c in columns if c in tmpl and tmpl[c].strip()}
    if not wanted:
        return xml, 0
    edits = []
    for rs, _re, rtag, rinner in iter_tags(xml, "row"):
        if rinner is None:
            continue
        rm = re.search(r'r="(\d+)"', rtag)
        if not rm or int(rm.group(1)) <= 4:
            continue
        rn = int(rm.group(1))
        base = rs + len(rtag)
        here = {}
        for cs, ce, ctag, cinner in iter_tags(rinner, "c"):
            ref = re.search(r'r="([A-Z]{1,3})\d+"', ctag)
            if ref:
                here[ref.group(1)] = (cs, ce, ctag, cinner)
        key = here.get(key_col)
        if not key or blank_value(key[2], key[3], sst):
            continue
        for col in wanted:
            cur = here.get(col)
            # 없는 셀·빈 셀은 기존 fill_missing이 스타일 본까지 고려해 만든다.
            if not cur or cur[3] is None or "<f" in cur[3]:
                continue
            sm = re.search(r'\ss="(\d+)"', cur[2])
            style = f' s="{sm.group(1)}"' if sm else ""
            new = (f'<c r="{col}{rn}"{style}><f>{esc(instantiate(tmpl[col], rn))}</f>'
                   f'<v/></c>')
            edits.append((base + cur[0], base + cur[1], new))
    if not edits:
        return xml, 0
    out, last = [], 0
    for start, end, new in sorted(edits):
        out.append(xml[last:start])
        out.append(new)
        last = end
    out.append(xml[last:])
    return "".join(out), len(edits)


def normalize_known_text(xml):
    """구조화 결과에 남은 확정 기사명 오탈자를 정확한 단일 텍스트 셀만 고친다."""
    count = 0
    for wrong, right in KNOWN_TEXT_FIXES.items():
        rx = re.compile(r"(<t(?:\s[^>]*)?>)" + re.escape(wrong) + r"(</t>)")
        xml, n = rx.subn(lambda m: m.group(1) + right + m.group(2), xml)
        count += n
    return xml, count


def direct_self_refs(xml):
    """A10 수식이 A10을 직접 참조하는 명백한 순환참조 목록."""
    bad = []
    for rn, col, formula, _s, _e in parse_cells(xml):
        # 다른 시트의 같은 주소(A5 등)는 순환이 아니다. 범위의 뒤쪽 주소까지
        # 통째로 제거해야 'Sheet'!$A$5:$A$154의 A154를 자기참조로 오인하지 않는다.
        local = re.sub(
            r"(?:'[^']+'|[A-Za-z0-9_가-힣]+)!"
            r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?",
            "", formula)
        if re.search(rf"(?<![A-Z0-9_])\$?{re.escape(col)}\$?{rn}(?!\d)", local, re.I):
            bad.append(f"{col}{rn}")
    return bad


def force_recalc_part(name, data):
    """calcChain을 제거하고 Excel이 열 때 자동 전체 재계산하도록 한다."""
    if name == "xl/calcChain.xml":
        return None
    text = None
    if name == "[Content_Types].xml":
        text = re.sub(r'<Override[^>]*calcChain[^>]*/>', "", data.decode("utf-8"))
    elif name == "xl/_rels/workbook.xml.rels":
        text = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", data.decode("utf-8"))
    elif name == "xl/workbook.xml":
        text = data.decode("utf-8")
        attrs = 'calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"'
        if "<calcPr" in text:
            def repl(m):
                tag = re.sub(r'\s(?:calcMode|fullCalcOnLoad|forceFullCalc)="[^"]*"', "", m.group(0))
                return tag.replace("<calcPr", "<calcPr " + attrs, 1)
            text = re.sub(r"<calcPr\b[^>]*?/?>", repl, text, count=1)
        else:
            text = text.replace("</workbook>", f"<calcPr {attrs}/></workbook>", 1)
    return text.encode("utf-8") if text is not None else data


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
    sst = shared_strings(z)
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

    patched, nfix, nwide, nmiss, nowned, ntext = {}, 0, 0, 0, 0, 0
    miss_by = {}
    for name, path in smap.items():
        x = raw[name]
        y, k = fix_sheet(x)
        y, w = widen(y, ends, name, old_ends.get(name))
        y, owned = restore_owned_formulas(y, FORMULA_OWNED.get(name, ()), sst=sst)
        y, mm = fill_missing(y, sst=sst)
        y, nt = normalize_known_text(y)
        if k or w or mm or owned or nt:
            patched[path] = y.encode("utf-8")
            nfix += k
            nwide += w
            nmiss += mm
            nowned += owned
            ntext += nt
            if mm:
                miss_by[name] = mm
    if "xl/sharedStrings.xml" in z.namelist():
        sx, nt = normalize_known_text(z.read("xl/sharedStrings.xml").decode("utf-8"))
        if nt:
            patched["xl/sharedStrings.xml"] = sx.encode("utf-8")
            ntext += nt
    z.close()
    print(f"\n깨진 수식 복구 {nfix}개 · 범위 확장 {nwide}곳 · 빠진 수식 채움 {nmiss}개 · "
          f"정적 상태 수식화 {nowned}개 · 기사명 정규화 {ntext}개 "
          f"(시트 {len(patched)}개)")
    for k, v in sorted(miss_by.items(), key=lambda x: -x[1]):
        print(f"   빠진 수식 {k}: {v}개")

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
        data = patched.get(it.filename)
        if data is None:
            data = zin.read(it.filename)
        data = force_recalc_part(it.filename, data)
        if data is not None:
            zout.writestr(it.filename, data)
    zout.close(); zin.close()

    # 검증: zip 무결성 + 음수행이 남아 있지 않은지
    zc = zipfile.ZipFile(dst)
    assert zc.testzip() is None, "zip 무결성 실패"
    # ★ XML 전체를 훑으면 'AS-2607-001' 같은 **접수ID 문자열**까지 세어 버린다.
    #   반드시 수식 본문만 본다(2026-07-26에 1,084개로 잘못 세었던 자리).
    left = 0
    self_refs = []
    for p in smap.values():
        xml = zc.read(p).decode("utf-8")
        self_refs.extend(direct_self_refs(xml))
        for _, _, f, _, _ in parse_cells(xml):
            if is_broken(f):
                left += 1
    assert not self_refs, "직접 순환참조 발견: " + ", ".join(self_refs[:10])
    zc.close()
    print(f"생성: {os.path.basename(dst)} · 남은 음수행 참조 {left}개 · 직접 순환참조 0개")


if __name__ == "__main__":
    main()
