# -*- coding: utf-8 -*-
"""
mobile_snapshot.py — **PC가 꺼져 있어도 폰에서 열리는 사본**을 만든다
================================================================================
터널·클라우드는 둘 다 PC(또는 서버)가 살아 있어야 한다. 이동이 잦아 PC를 못 켜 두는
상황에서는 그 어느 쪽도 답이 아니다.

  → 화면에 필요한 값을 **HTML 한 파일 안에 넣어** 버린다.
    서버도, 네트워크도, 엑셀도 필요 없다. 파일만 열리면 끝.

담는 것 : 확인필요 목록(처리 방법 포함)·핵심 숫자·돌발AS/정기점검 현황·계산서 구성
담지 않는 것 : 인증키·엑셀 원본·밴드 원문

실행
  python mobile_snapshot.py            # reports/mobile_snapshot.html 생성
"""
import sys, os, json, html
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "webapp"))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT = os.path.join(ROOT, "reports", "mobile_snapshot.html")


def pick(d, keys):
    """필요한 열만 남긴다 — 폰에서 열 파일이라 용량이 곧 대기시간이다"""
    return {k: d.get(k) for k in keys if d.get(k) not in (None, "")}


def collect():
    import app_server as A
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])

    def safe(fn, dflt):
        try:
            return fn()
        except Exception as e:
            print(f"  ! {e}")
            return dflt

    works = safe(A.get_works, {"as": [], "pm": []})
    issues = safe(A.get_issues, {"rows": [], "cols": []})
    settle = safe(A.get_settlements, [])
    erp = safe(A.get_erpdocs, {"rows": []})

    W_AS = ["프로젝트NO", "캠프명", "접수일자", "작업완료일", "담당기사", "진행상태",
            "유상·무상·보험", "완료보고서등록", "사진등록", "ERP등록",
            "검증결과", "검증문제코드", "관리자검증상태"]
    W_PM = ["점검ID", "프로젝트NO", "캠프명", "점검예정일", "실제점검일", "담당기사", "점검상태",
            "ERP판매전표", "거래명세서", "검증결과", "검증문제코드",
            "장비수", "장비내역", "반영상태", "원본행", "출처"]
    S_K = ["정산ID", "프로젝트NO", "캠프명", "업무구분", "완료일", "공급가액",
           "상태", "명세서번호", "계산서발행일", "입금일"]
    E_K = ["전표", "월", "유형", "공급가액", "프로젝트명", "판정", "포함프로젝트"]

    return {
        "기준": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "원본": os.path.basename(master),
        "as": [pick(r, W_AS) for r in A.app_year_rows(works.get("as", []), "as")],
        "pm": [pick(r, W_PM) for r in A.app_year_rows(works.get("pm", []), "pm")],
        "issues": A.app_year_rows(issues.get("rows", []), "issue"),
        "settle": [pick(r, S_K) for r in A.app_year_rows(settle, "settle")],
        "erp": [pick(r, E_K) for r in A.app_year_rows(erp.get("rows", []), "erp")],
    }


PAGE = """<style>
:root{
  --ink-1:#111826; --ink-2:#414b60; --ink-3:#5f6a7d; --line:#dfe4ec;
  --bg:#f5f7fa; --card:#ffffff; --navy:#0e1b3f; --brand:#1c4fd8; --brand-btn:#1c4fd8;
  --ok:#0f7a3d; --warn:#a85a06; --bad:#b3202c;
  --okbg:#e6f4ec; --warnbg:#fdf0e0; --badbg:#fbe9eb;
}
@media (prefers-color-scheme:dark){
  :root{ --ink-1:#e9edf5; --ink-2:#b3bccd; --ink-3:#8391a8; --line:#28324a;
    --bg:#0b1020; --card:#131a2c; --navy:#070d1c; --brand:#7ba2ff; --brand-btn:#2b57c9;
    --ok:#5fd08a; --warn:#f0ad4e; --bad:#ff8189;
    --okbg:#12321f; --warnbg:#33240d; --badbg:#3a1519; }
}
:root[data-theme="dark"]{ --ink-1:#e9edf5; --ink-2:#b3bccd; --ink-3:#8391a8; --line:#28324a;
  --bg:#0b1020; --card:#131a2c; --navy:#070d1c; --brand:#7ba2ff; --brand-btn:#2b57c9;
  --ok:#5fd08a; --warn:#f0ad4e; --bad:#ff8189;
  --okbg:#12321f; --warnbg:#33240d; --badbg:#3a1519; }
:root[data-theme="light"]{ --ink-1:#111826; --ink-2:#414b60; --ink-3:#5f6a7d; --line:#dfe4ec;
  --bg:#f5f7fa; --card:#ffffff; --navy:#0e1b3f; --brand:#1c4fd8; --brand-btn:#1c4fd8;
  --ok:#0f7a3d; --warn:#a85a06; --bad:#b3202c;
  --okbg:#e6f4ec; --warnbg:#fdf0e0; --badbg:#fbe9eb; }

/* 앱 전체 글꼴 — 여기 한 곳에서만 정한다. 되돌리기: `python webapp/font_switch.py --legacy`
   또는 --font-ui 값을 --font-ui-legacy 로 바꾼다. legacy 값은 원래 값 그대로 남긴다. */
:root{
  --font-ui-legacy:"Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",system-ui,sans-serif;
  --font-ui:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Pretendard Variable",Pretendard,"Segoe UI Variable Text","Segoe UI",Roboto,"Noto Sans KR","Source Han Sans KR","본고딕","Malgun Gothic","맑은 고딕","Helvetica Neue",Arial,sans-serif;
}
:root[data-font="legacy"]{ --font-ui:var(--font-ui-legacy); }
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink-1);
  font-family:var(--font-ui);
  font-size:15px;line-height:1.55;-webkit-text-size-adjust:100%}
.wrap{max-width:820px;margin:0 auto;padding:0 14px 64px}
header{background:var(--navy);color:#fff;padding:14px 0 12px;margin-bottom:14px}
header .wrap{padding-bottom:0}
h1{margin:0;font-size:17px;font-weight:800;letter-spacing:-.3px}
.sub{font-size:11.5px;color:#9fb0d4;margin-top:3px;letter-spacing:.02em}
.offline{display:inline-block;margin-top:8px;font-size:11px;background:#1b2c56;color:#c9d8ff;
  border-radius:999px;padding:3px 10px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:9px;margin:0 0 16px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:11px 12px}
.kpi .l{font-size:11.5px;color:var(--ink-3)}
.kpi .v{font-size:22px;font-weight:800;letter-spacing:-.5px;font-variant-numeric:tabular-nums;margin-top:2px}
.kpi.bad .v{color:var(--bad)} .kpi.warn .v{color:var(--warn)} .kpi.ok .v{color:var(--ok)}
nav{display:flex;gap:6px;overflow-x:auto;padding-bottom:8px;margin-bottom:10px;
  border-bottom:1px solid var(--line)}
nav button{flex:0 0 auto;background:transparent;border:1px solid var(--line);color:var(--ink-2);
  border-radius:999px;padding:7px 14px;font-size:13px;font-weight:700;cursor:pointer;
  font-family:inherit}
nav button[aria-selected="true"]{background:var(--brand-btn);border-color:var(--brand-btn);color:#fff}
nav button:focus-visible{outline:2px solid var(--brand);outline-offset:2px}
.tools{display:flex;gap:8px;margin-bottom:12px}
.tools input,.tools select{flex:1;min-width:0;background:var(--card);color:var(--ink-1);
  border:1px solid var(--line);border-radius:10px;padding:9px 11px;font-size:14px;font-family:inherit}
.tools input:focus,.tools select:focus{outline:2px solid var(--brand);outline-offset:1px}
.count{font-size:12px;color:var(--ink-3);margin:0 0 8px}
ul.list{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:8px}
li.row{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:11px 12px;
  border-left:3px solid var(--line)}
li.row.sev-bad{border-left-color:var(--bad)} li.row.sev-warn{border-left-color:var(--warn)}
li.row.sev-ok{border-left-color:var(--ok)}
.r1{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap}
.prj{font-weight:800;font-size:15.5px;font-variant-numeric:tabular-nums;letter-spacing:-.2px}
.camp{font-weight:700;font-size:13.5px;color:var(--ink-2)}
.camp.none{color:var(--bad);font-weight:600;font-size:12.5px}
.chip{font-size:10.5px;font-weight:800;border-radius:999px;padding:2px 8px;white-space:nowrap}
.chip.ok{background:var(--okbg);color:var(--ok)} .chip.warn{background:var(--warnbg);color:var(--warn)}
.chip.bad{background:var(--badbg);color:var(--bad)}
.meta{font-size:12px;color:var(--ink-3);margin-top:5px;display:flex;gap:10px;flex-wrap:wrap;
  font-variant-numeric:tabular-nums}
.how{font-size:12.5px;color:var(--ink-2);margin-top:7px;padding-top:7px;border-top:1px dashed var(--line)}
.empty{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:22px;
  text-align:center;color:var(--ink-3);font-size:13.5px}
footer{margin-top:22px;font-size:11.5px;color:var(--ink-3);line-height:1.7}
@media (prefers-reduced-motion:no-preference){li.row{transition:border-color .15s}}
</style>

<header><div class="wrap">
  <h1>쿠팡 통합업무 — 이동 중 확인용</h1>
  <div class="sub" id="stamp"></div>
  <span class="offline">인터넷·PC 없이 열립니다</span>
</div></header>

<div class="wrap">
  <div class="kpis" id="kpis"></div>
  <nav id="tabs" role="tablist"></nav>
  <div class="tools">
    <input id="q" type="search" placeholder="프로젝트NO·캠프·기사 검색" autocomplete="off">
    <select id="f"></select>
  </div>
  <p class="count" id="count"></p>
  <ul class="list" id="list"></ul>
  <footer id="foot"></footer>
</div>

<script>
const D = window.__CSOS__;
const $ = i => document.getElementById(i);
const esc = s => String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmt = n => (+n||0).toLocaleString('ko-KR');

const TABS = [
  {k:'issues', n:'확인필요'},
  {k:'as',     n:'돌발AS'},
  {k:'pm',     n:'정기점검'},
  {k:'settle', n:'정산'},
  {k:'erp',    n:'계산서'},
];
let tab = 'issues';

/* 심각도 — 무엇부터 봐야 하는지가 색으로 먼저 읽히게 */
function sev(r){
  if(tab==='issues') return r['구분']==='정산' ? 'bad' : 'warn';
  const v = r['검증결과']||'';
  if(v==='정상') return 'ok';
  if(v==='누락·지연') return 'bad';
  if(tab==='settle') return (r['상태']==='정상') ? 'ok' : 'warn';
  return v ? 'warn' : '';
}
function chip(r){
  if(tab==='issues') return `<span class="chip warn">${esc(r['구분'])}</span>`;
  const st = r['진행상태']||r['점검상태']||r['상태']||r['판정']||'';
  const cls = /완료|정상|입금/.test(st) ? 'ok' : /취소|미점검|누락/.test(st) ? 'bad' : 'warn';
  return st ? `<span class="chip ${cls}">${esc(st)}</span>` : '';
}
function line(r){
  const prj = r['프로젝트NO']||r['ID']||r['전표']||'-';
  const camp = r['캠프명'];
  const campHtml = camp ? `<span class="camp">${esc(camp)}</span>`
                        : (tab==='issues'||tab==='erp' ? '' : `<span class="camp none">캠프 미상</span>`);
  const meta = [];
  if(tab==='issues'){ if(r['문제유형']) meta.push(esc(r['문제유형'])); if(r['담당자']) meta.push(esc(r['담당자'])); }
  if(tab==='as'){ meta.push('접수 '+esc(r['접수일자']||'-'));
                  if(r['작업완료일']) meta.push('완료 '+esc(r['작업완료일']));
                  if(r['담당기사']) meta.push(esc(r['담당기사'])); }
  if(tab==='pm'){ meta.push('예정 '+esc(r['점검예정일']||'-'));
                  if(r['실제점검일']) meta.push('실제 '+esc(r['실제점검일']));
                  if(r['담당기사']) meta.push(esc(r['담당기사'])); }
  if(tab==='settle'){ if(r['완료일']) meta.push('완료 '+esc(r['완료일']));
                      if(r['공급가액']) meta.push(fmt(r['공급가액'])+'원');
                      if(r['정산ID']) meta.push(esc(r['정산ID'])); }
  if(tab==='erp'){ if(r['월']) meta.push(esc(r['월'])); if(r['유형']) meta.push(esc(r['유형']));
                   if(r['공급가액']) meta.push(fmt(r['공급가액'])+'원'); }
  const how = tab==='issues' ? (r['내용·근거']||'')
            : tab==='erp'    ? (r['포함프로젝트']||'')
            : (r['검증문제코드']||'');
  return `<li class="row sev-${sev(r)}">
    <div class="r1"><span class="prj">${esc(prj)}</span>${campHtml}${chip(r)}</div>
    ${meta.length?`<div class="meta">${meta.map(m=>`<span>${m}</span>`).join('')}</div>`:''}
    ${how?`<div class="how">${esc(how)}</div>`:''}</li>`;
}

function rows(){ return D[tab]||[]; }
function filterVals(){
  const key = tab==='issues' ? '구분' : tab==='erp' ? '판정'
            : tab==='settle' ? '상태' : (tab==='as'?'진행상태':'점검상태');
  const s = new Set(rows().map(r=>String(r[key]||'').split('(')[0]).filter(Boolean));
  return [key, [...s].sort()];
}
function render(){
  const [key, vals] = filterVals();
  const q = $('q').value.trim().toLowerCase();
  const fv = $('f').value;
  let list = rows();
  if(fv) list = list.filter(r=>String(r[key]||'').split('(')[0]===fv);
  if(q) list = list.filter(r=>Object.values(r).some(v=>String(v==null?'':v).toLowerCase().includes(q)));
  $('count').textContent = `${list.length}건` + (list.length!==rows().length ? ` (전체 ${rows().length})` : '');
  $('list').innerHTML = list.length ? list.map(line).join('')
    : `<div class="empty">조건에 맞는 건이 없습니다</div>`;
}
function syncFilter(){
  const [, vals] = filterVals();
  $('f').innerHTML = `<option value="">전체</option>` + vals.map(v=>`<option>${esc(v)}</option>`).join('');
}
function setTab(k){
  tab = k;
  [...$('tabs').children].forEach(b=>b.setAttribute('aria-selected', String(b.dataset.k===k)));
  syncFilter(); render();
}

$('stamp').textContent = `${D['기준']} 기준 · ${D['원본']}`;
$('tabs').innerHTML = TABS.map(t=>
  `<button role="tab" data-k="${t.k}" aria-selected="${t.k===tab}">${t.n} ${ (D[t.k]||[]).length }</button>`).join('');
$('tabs').addEventListener('click', e=>{ const b=e.target.closest('button'); if(b) setTab(b.dataset.k); });
$('q').addEventListener('input', render);
$('f').addEventListener('change', render);

const iss = D.issues||[], asx = D.as||[], pm = D.pm||[];
const cnt = k => iss.filter(r=>r['구분']===k).length;
$('kpis').innerHTML = [
  ['확인필요', iss.length, 'bad'],
  ['정산 조치필요', cnt('정산'), 'warn'],
  ['돌발AS', asx.length, ''],
  ['정기점검', pm.length, ''],
].map(([l,v,c])=>`<div class="kpi ${c}"><div class="l">${l}</div><div class="v">${fmt(v)}</div></div>`).join('');
$('foot').innerHTML = `이 페이지는 만든 시점의 <b>사본</b>입니다 — 인터넷이 끊겨도 열리지만 숫자는 갱신되지 않습니다.<br>
  입력·대조·엑셀 반영은 사무실 PC에서 합니다. 최신본이 필요하면 PC에서 다시 만들어 주세요.`;

setTab('issues');
</script>
"""


def main():
    print("폰용 사본 만드는 중…")
    d = collect()
    body = ('<script>window.__CSOS__=' +
            json.dumps(d, ensure_ascii=False).replace("</", "<\\/") + ';</script>\n' + PAGE)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    open(OUT, "w", encoding="utf-8").write(body)
    kb = os.path.getsize(OUT) / 1024
    print(f"확인필요 {len(d['issues'])} · AS {len(d['as'])} · 점검 {len(d['pm'])} · "
          f"정산 {len(d['settle'])} · 계산서 {len(d['erp'])}")
    print(f"→ {OUT}  ({kb:.0f} KB)")


if __name__ == "__main__":
    main()
