# -*- coding: utf-8 -*-
"""
tax_archive.py — 세금계산서를 **건별 PDF** 로 보관한다 (일자-No. 가 곧 파일명)

사용자 지시(2026-08-05): "1월부터 지금까지 세금계산서 pdf 건별 전체 다운로드 …
원본 저장". ERP `(세금)계산서진행단계` 엑셀은 246건이 한 파일에 들어 있어 사람이
특정 계산서를 찾을 수 없다. 한 건 = 한 PDF 로 굳혀 월별 폴더에 넣는다.

    세금계산서_건별/2026/01/[2026-01-15-3]_쿠팡로지스틱스_UJ2600123_1,650,000원_발행완료.pdf
      └ 대괄호=일자-No.(ERP 키) · 거래처 · 프로젝트NO · 금액 · 진행단계

프로젝트NO 는 프로젝트명에서 UJ 패턴이 보일 때만 넣는다(지어내지 않는다).
승인번호가 있으면 PDF 본문에 남긴다 — 홈택스 대조의 정본 키다.

  python tax_archive.py             # 새 건만
  python tax_archive.py --limit 300
  python tax_archive.py --force
"""
import argparse
import glob
import json
import os
import re
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

UJ = re.compile(r"UJ\d{7}")
OUT_JSON = os.path.join(ROOT, "reports", "세금계산서_상세.json")


def out_root():
    import source_dirs as S
    return os.path.join(S.ERP_DIR, "세금계산서_건별")


def safe(s, n=24):
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]", " ", str(s or "")).strip()
    return re.sub(r"\s+", " ", s)[:n]


def norm_slip(raw):
    """`2026/01/15 -3` → `2026-01-15-3`."""
    parts = re.findall(r"\d+", str(raw or ""))
    if len(parts) < 4:
        return ""
    return f"{parts[0]}-{parts[1]}-{parts[2]}-{int(parts[3])}"


def collect():
    """계산서진행단계 엑셀(내용 판별 taxinv)에서 건별 행을 모은다."""
    import warnings
    warnings.filterwarnings("ignore")
    import openpyxl
    import source_dirs as S
    cands = [p for p in glob.glob(os.path.join(S.ERP_DIR, "**", "*.xlsx"), recursive=True)
             if not os.path.basename(p).startswith(("~$", "ESD007E"))]
    cands.sort(key=os.path.getmtime, reverse=True)
    rows, seen, src = [], set(), []
    for p in cands[:80]:
        try:
            wb = openpyxl.load_workbook(p, read_only=False, data_only=True)
            ws = wb.active
            head = [str(c or "") for c in next(ws.iter_rows(min_row=2, max_row=2, values_only=True))]
            j = "|".join(head)
            if "진행단계" not in j or "승인번호" not in j:
                wb.close()
                continue
            idx = {h: i for i, h in enumerate(head)}
            for r in ws.iter_rows(min_row=3, values_only=True):
                r = ["" if c is None else str(c).strip() for c in r]
                if not r or len(r) < 5:
                    continue
                slip = norm_slip(r[0])
                if not slip or slip in seen:
                    continue
                seen.add(slip)
                g = lambda k, d="": (r[idx[k]] if k in idx and len(r) > idx[k] else d)
                amt = (g("공급가액") or "").replace(",", "")
                rows.append({
                    "slip": slip, "project": g("프로젝트명"), "cust": g("거래처명"),
                    "issued": g("발행일자"), "supply": int(float(amt)) if re.fullmatch(r"-?\d+(\.\d+)?", amt) else 0,
                    "vat": g("부가세"), "total": g("합계금액"), "type": g("종류"),
                    "state": g("전자(세금)계산서 진행단계"), "approve": g("승인번호"),
                    "memo": g("적요명"), "src": os.path.basename(p),
                })
            src.append(os.path.basename(p))
            wb.close()
        except Exception:
            continue
    return rows, src


def render(row, path):
    from archive_render import html_to_pdf, esc
    uj = (UJ.search(row.get("project") or "") or [None])
    uj = uj.group(0) if hasattr(uj, "group") else ""
    html = f"""
<h1>세금계산서 {esc(row['slip'])}</h1>
<div class="meta">
 <b>거래처</b> {esc(row.get('cust'))} &nbsp;|&nbsp; <b>프로젝트</b> {esc(row.get('project'))}
 {f'({esc(uj)})' if uj else ''}<br>
 <b>발행일자</b> {esc(row.get('issued') or '(미발행)')} &nbsp;|&nbsp;
 <b>진행단계</b> {esc(row.get('state'))} &nbsp;|&nbsp; <b>종류</b> {esc(row.get('type'))}<br>
 <b>승인번호</b> {esc(row.get('approve') or '-')}
</div>
<table>
 <tr><th>공급가액</th><th>부가세</th><th>합계금액</th></tr>
 <tr><td class="num">{row.get('supply', 0):,}</td><td class="num">{esc(row.get('vat'))}</td>
     <td class="num">{esc(row.get('total'))}</td></tr>
</table>
<p><b>적요</b> {esc(row.get('memo'))}</p>
<div class="foot">원본: ERP (세금)계산서진행단계 {esc(row.get('src'))} ·
 보관 생성 {time.strftime('%Y-%m-%d %H:%M')} · 자동 생성본(원본 불변)</div>
"""
    return html_to_pdf(html, path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=400)
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()

    rows, src = collect()
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    json.dump({"count": len(rows), "src": src, "rows": rows},
              open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    root, made, skip, fail = out_root(), 0, 0, 0
    for row in rows:
        if made >= a.limit:
            break
        y, m = row["slip"][:4], row["slip"][5:7]
        uj = (UJ.search(row.get("project") or "") or [None])
        uj = uj.group(0) if hasattr(uj, "group") else ""
        state = "발행완료" if str(row.get("state", "")).startswith(("발행", "전송")) else \
                safe(row.get("state") or "미발행", 10)
        name = (f"[{row['slip']}]_{safe(row.get('cust'), 18)}"
                f"{'_' + uj if uj else ''}_{row.get('supply', 0):,}원_{state}.pdf")
        dst = os.path.join(root, y, m, name)
        if not a.force and os.path.exists(dst):
            skip += 1
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if render(row, dst):
            made += 1
        else:
            fail += 1
    print(f"세금계산서 건별 PDF: 새로 {made}건 · 건너뜀 {skip} · 실패 {fail} "
          f"(원본 {len(rows)}건) → {root}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
