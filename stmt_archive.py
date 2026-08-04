# -*- coding: utf-8 -*-
"""
stmt_archive.py — 거래명세서를 **건별 PDF** 로 보관한다 (전표번호가 곧 파일명)

사용자 지시(2026-08-05): "엑셀만 받지 말고 각 건의 PDF 및 이미지 모두 번호 및 알아볼 수
있게 저장해서 관리."

ERP 인쇄본 엑셀(ESD007E)은 25건이 한 파일에 묶여 있어 사람이 특정 명세서를 찾을 수 없다.
`stmt_docs.py` 가 뜯어 둔 건별 데이터를 **한 건 = 한 PDF** 로 굳힌다.

    거래명세서/2026/07/[2026-07-14-3]_송파5캠프_UJ2600895_1,650,000원.pdf
      └ 대괄호=전표번호(원장 06시트 거래명세서번호와 같은 키) · 캠프 · 프로젝트NO · 금액

프로젝트NO 는 `stmt_link.py` 가 판매조회와 맞춰 확정한 것만 넣는다(추측 금지).

  python stmt_archive.py            # 새 건만
  python stmt_archive.py --limit 100
  python stmt_archive.py --force
"""
import argparse
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

DETAIL = os.path.join(ROOT, "reports", "거래명세서_상세.json")
LINK = os.path.join(ROOT, "reports", "명세서_프로젝트_매칭.json")


def out_root():
    import source_dirs as S
    return os.path.join(S.ERP_DIR, "거래명세서_건별")


def safe(s, n=28):
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]", " ", str(s or "")).strip()
    return re.sub(r"\s+", " ", s)[:n]


def render(doc, uj, state, path):
    from archive_render import html_to_pdf, esc
    rows = "".join(
        f"<tr><td>{esc(i.get('date'))}</td><td>{esc(i.get('name'))}</td>"
        f"<td class='num'>{i.get('qty') or ''}</td>"
        f"<td class='num'>{(i.get('supply') or 0):,}</td>"
        f"<td class='num'>{(i.get('vat') or 0):,}</td>"
        f"<td>{esc(i.get('note'))}</td></tr>" for i in doc.get("items", []))
    html = f"""
<h1>거래명세서 {esc(doc['slip'])}</h1>
<div class="meta">
 <b>거래처</b> {esc(doc.get('cust'))} &nbsp;|&nbsp; <b>출고창고</b> {esc(doc.get('warehouse'))}
 &nbsp;|&nbsp; <b>프로젝트NO</b> {esc(uj or '(미확정)')}
 &nbsp;|&nbsp; <b>ERP 진행상태</b> {esc(state or '-')}
 <br><b>금액</b> {doc.get('amount', 0):,}원
</div>
<table>
 <tr><th>일자</th><th>품목[규격]</th><th>수량</th><th>공급가액</th><th>부가세</th><th>적요</th></tr>
 {rows}
</table>
<div class="foot">원본: ERP 거래명세서 인쇄본 {esc(doc.get('src'))} ·
 보관 생성 {time.strftime('%Y-%m-%d %H:%M')} · 자동 생성본(원본 불변)</div>
"""
    return html_to_pdf(html, path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=300)
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()

    docs = json.load(open(DETAIL, encoding="utf-8"))["docs"]
    link = {}
    if os.path.exists(LINK):
        for x in json.load(open(LINK, encoding="utf-8")).get("linked", []):
            link[x["slip"]] = x
    root, made, skip, fail = out_root(), 0, 0, 0
    for d in docs:
        if made >= a.limit:
            break
        slip = d["slip"]                       # 2026/07/14-3
        ym = slip[:7].split("/")
        info = link.get(slip) or {}
        uj = info.get("uj", "")
        name = (f"[{slip.replace('/', '-')}]_{safe(d.get('cust'), 20)}"
                f"{'_' + uj if uj else ''}_{d.get('amount', 0):,}원.pdf")
        dst = os.path.join(root, ym[0], ym[1], name)
        if not a.force and os.path.exists(dst):
            skip += 1
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if render(d, uj, info.get("state"), dst):
            made += 1
        else:
            fail += 1
    print(f"거래명세서 건별 PDF: 새로 {made}건 · 건너뜀 {skip} · 실패 {fail} → {root}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
