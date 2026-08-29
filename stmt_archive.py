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
import child_budget
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DETAIL = os.path.join(ROOT, "reports", "거래명세서_상세.json")
LINK = os.path.join(ROOT, "reports", "명세서_프로젝트_매칭.json")

# ── 시간 예산 — **바깥에서 죽기 전에 스스로 멈춘다** (2026-08-25 · [427] 과 같은 자리)
# ★ 자율복구 대기표가 이 파일을 `--limit 150` 으로 부르는데 예산이 없어 제한시간에
#   **SIGKILL(-9)** 로 끊겼다(실측 시도 4회 · 자국이 `returncode=-9` 다섯 글자뿐).
#   파이썬 stdout 은 파이프에 물리면 **블록 버퍼**라 그때까지 찍은 줄이 버퍼째 사라진다.
# ★ **일은 되고 있었다** — PDF 는 한 건씩 `os.replace` 로 굳고 다음 회차가 그대로
#   이어받는다(`os.path.exists` 로 건너뛴다). 잃은 것은 '얼마나 했나' 였다([169]).
# ★ 실측 2026-08-25: 원본 812건 · 이미 보관 211 → **남음 601건**.
BUDGET_ENV = "STMT_ARCHIVE_BUDGET_SEC"
INCREMENTAL_RETURN_CODE = child_budget.INCREMENTAL_RETURN_CODE


def out_root():
    import source_dirs as S
    return os.path.join(S.ERP_DIR, "거래명세서_건별")


def safe(s, n=28):
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]", " ", str(s or "")).strip()
    return re.sub(r"\s+", " ", s)[:n]


def path_key(path):
    return os.path.normcase(os.path.abspath(path))


def existing_pdf_paths(root):
    """기존 PDF를 디렉터리 열람 한 번으로 읽는다.

    `os.path.exists(dst)`를 문서마다 부르면 Z: 왕복이 원본 수만큼 생긴다. 보관이 거의
    끝난 뒤에는 새로 만들 것은 0건인데도 812개를 하나씩 stat 하느라 300초 바깥 제한에
    죽었다. 스캔 오류는 무시하지 않는다 — 일부만 읽고 없는 것으로 오인해 덮어쓰는 것보다
    실패로 남기는 편이 안전하다.
    """
    if not os.path.isdir(root):
        return set()

    def raise_walk_error(exc):
        raise exc

    found = set()
    for base, _dirs, files in os.walk(root, onerror=raise_walk_error):
        for name in files:
            if name.lower().endswith(".pdf"):
                found.add(path_key(os.path.join(base, name)))
    return found


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


def render_atomic(doc, uj, state, path):
    """완성된 PDF만 정식 이름으로 보인다.

    회차 시간 제한이 크롬을 중간에 끊어도 `.part-*`만 남고, 다음 회차가 정식 파일을
    다시 만든다. 예전에는 반쪽 PDF도 `exists()`에 걸려 영구 완료로 오인될 수 있었다.
    """
    tmp = path + f".part-{os.getpid()}"
    try:
        made = render(doc, uj, state, tmp)
        if made and os.path.exists(tmp):
            os.replace(tmp, path)
            return path
        return None
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


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
    budget = child_budget.start(BUDGET_ENV)
    root, made, skip, fail, attempted = out_root(), 0, 0, 0, 0
    existing = set() if a.force else existing_pdf_paths(root)
    cut_reason, seen = "", 0
    for d in docs:
        if attempted >= a.limit:
            # ★ `--limit`도 **잔량 이월**이다. 예전에는 여기서 조용히 빠진 뒤 0을
            #   돌려, 상한 뒤에 아직 문서가 있어도 자율복구가 큐를 `done`으로 닫았다.
            #   다음 항목이 실제로 새 작업인지까지 훑지 않았으므로 한 번 더 확인하는
            #   것이 안전하다. 다음 회차가 전부 exists면 그때 0으로 끝난다.
            cut_reason = "limit"
            break
        seen += 1
        slip = d["slip"]                       # 2026/07/14-3
        ym = slip[:7].split("/")
        info = link.get(slip) or {}
        uj = info.get("uj", "")
        name = (f"[{slip.replace('/', '-')}]_{safe(d.get('cust'), 20)}"
                f"{'_' + uj if uj else ''}_{d.get('amount', 0):,}원.pdf")
        dst = os.path.join(root, ym[0], ym[1], name)
        if not a.force and path_key(dst) in existing:
            skip += 1
            continue
        if child_budget.over():   # ★ 예산이 다 되면 **새로 안 만든다**
            cut_reason = "budget"
            break
        attempted += 1
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if render_atomic(d, uj, info.get("state"), dst):
            made += 1
            existing.add(path_key(dst))
        else:
            fail += 1
    print(f"거래명세서 건별 PDF: 새로 {made}건 · 건너뜀 {skip} · 실패 {fail} → {root}")
    if fail:
        # 렌더가 False를 돌린 문서는 정식 PDF가 생기지 않았다. 실패가 있는데 0으로
        # 닫으면 다음 회차가 영영 오지 않을 수 있으므로 코드 실패로 분명히 남긴다.
        print(f"  ★ PDF {fail}건을 만들지 못했다 — 완료로 닫지 않고 실패로 남긴다.")
        return 1
    if cut_reason == "budget":
        # ★ **멈춘 사실과 숫자를 말한다.** 조용히 돌아가면 부르는 쪽은 실패인지
        #   완료인지 구별할 수 없다([169]). 건너뜀 숫자는 여기까지 본 것뿐이다.
        print(f"  ★ 시간 예산({budget}초)이 다 되어 여기까지 하고 돌아온다 — "
              f"원본 {len(docs):,}건 중 앞에서부터 {seen:,}건까지 봤다. "
              "남은 것은 다음 회차가 이어서 한다(PDF 는 한 건씩 저장돼 있다).")
        return INCREMENTAL_RETURN_CODE
    if cut_reason == "limit":
        print(f"  ★ 건수 상한({a.limit:,}건)에 닿아 여기까지 하고 돌아온다 — "
              f"원본 {len(docs):,}건 중 앞에서부터 {seen:,}건까지 봤다. "
              "잔량 확인은 다음 회차가 이어서 한다(PDF 는 한 건씩 저장돼 있다).")
        return INCREMENTAL_RETURN_CODE
    return 0


if __name__ == "__main__":
    sys.exit(main())
