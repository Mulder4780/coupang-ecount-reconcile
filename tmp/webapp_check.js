
/* 크롬의 [설치 및 바로가기 만들기]는 fetch 핸들러를 가진 서비스 워커가 있어야 동작한다.
   매니페스트만으로는 단순 북마크가 되거나 메뉴가 아예 반응하지 않는다. */
if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(() => {});

let PIN = localStorage.getItem('cw_pin') || '';
const APP_YEAR = '2026';
let settleRows = [], reports = [], curReport = 0, polling = null;
const STAFF_CENTERS = {
  'ryu-jiyeong':{name:'류지영',title:'류지영 쿠팡 AS 및 정기점검 업무센터',assignee:'류지영',
    checklist:['신규 돌발AS 접수와 처리상태 확인','정기점검 예정·실행·미실시 사유 입력','카카오톡 정기점검방·돌발점검방 원본 업로드','택배 발송·현장 조치 완료일과 근거 첨부','거래명세서·세금계산서 발행 확인','입금일·입금액·잔여 미수금 확인','ERP 금액 불일치와 청구 미등록 보완']},
  'oh-jonghyeon':{name:'오종현',title:'오종현 업무센터',assignee:'오종현',
    checklist:['PO 원본·견적서 수신 여부 확인','구매·입금 원천자료 누락 확인','프로젝트번호·캠프·금액 불일치 보완']},
};
const staffSlug=(()=>{
  const m=location.pathname.match(/^\/staff\/([a-z0-9-]+)\/?$/);
  return m&&STAFF_CENTERS[m[1]]?m[1]:(location.pathname.replace(/\/+$/,'')==='/ryu'?'ryu-jiyeong':'');
})();
let deferredInstallPrompt=window.__csosInstallPrompt||null;
let notificationItems=[];

const $ = id => document.getElementById(id);
const fmt = n => {
  if(n==null || n==='') return '-';
  const v = Number(n);
  return Number.isFinite(v)
    ? Math.round(v).toLocaleString('ko-KR',{maximumFractionDigits:0})
    : '-';
};
async function api(path, opt={}){
  opt.headers = Object.assign({'X-Pin': PIN}, opt.headers||{});
  if(staffSlug) opt.headers['X-Staff-Slug']=staffSlug;
  const r = await fetch(path, opt);
  if(r.status===401){
    PIN=''; localStorage.removeItem('cw_pin');   // 구 PIN 즉시 폐기 → 자동 폴링 중단(자기 잠금 방지)
    $('gate').style.display='flex'; throw 'PIN';
  }
  return r.json();
}
function escNotice(s){return esc2(String(s||''));}
async function loadNotifications(){
  if(!PIN) return;
  try{
    const d=await api('/api/notifications?t='+Date.now());
    notificationItems=d.items||[];
    const badge=$('noticeBadge');
    if(badge){
      badge.textContent=String(d.count||0);
      badge.style.display=(d.count||0)?'flex':'none';
    }
    const panel=$('noticePanel');
    if(panel) panel.innerHTML=notificationItems.length
      ? notificationItems.map(x=>`<div class="notice-item ${escNotice(x.severity)}">
          <b>${escNotice(x.title)}</b><span>${escNotice(x.detail)}</span></div>`).join('')
      : '<div class="notice-item info"><b>새 알림이 없습니다.</b><span>입력과 자동 반영 상태가 정상입니다.</span></div>';
  }catch(e){}
}
function toggleNotifications(e){
  if(e) e.stopPropagation();
  const p=$('noticePanel'); if(!p) return;
  p.classList.toggle('on');
  if(p.classList.contains('on')) loadNotifications();
}
document.addEventListener('click',e=>{
  const p=$('noticePanel');
  if(p&&p.classList.contains('on')&&!e.target.closest('.notice-wrap')) p.classList.remove('on');
});
window.addEventListener('beforeinstallprompt',e=>{
  if(!staffSlug) return;
  e.preventDefault(); deferredInstallPrompt=e;
  maybeShowInstallCard();
});
window.addEventListener('csos-install-ready',()=>{
  if(!staffSlug) return;
  deferredInstallPrompt=window.__csosInstallPrompt||deferredInstallPrompt;
  maybeShowInstallCard();
});
window.addEventListener('appinstalled',()=>{
  deferredInstallPrompt=null;
  window.__csosInstallPrompt=null;
  if(staffSlug) localStorage.setItem('csos_installed_'+staffSlug,'1');
  if($('installCard')) $('installCard').classList.remove('on');
  toast('업무센터 앱 설치가 완료되었습니다');
});
function workcenterRunsAsApp(){
  return window.matchMedia('(display-mode: standalone)').matches||
    window.matchMedia('(display-mode: fullscreen)').matches||
    window.navigator.standalone===true;
}
async function workcenterIsInstalled(){
  if(!staffSlug) return false;
  if(workcenterRunsAsApp()||localStorage.getItem('csos_installed_'+staffSlug)==='1') return true;
  if(typeof navigator.getInstalledRelatedApps==='function'){
    try{
      const apps=await navigator.getInstalledRelatedApps();
      if(Array.isArray(apps)&&apps.length){
        localStorage.setItem('csos_installed_'+staffSlug,'1');
        return true;
      }
    }catch(e){}
  }
  return false;
}
async function maybeShowInstallCard(){
  if(!staffSlug||localStorage.getItem('csos_install_dismissed_'+staffSlug)) return;
  if(await workcenterIsInstalled()){
    if($('installCard')) $('installCard').classList.remove('on');
    return;
  }
  showInstallCard(true);
}
async function showInstallCard(skipInstalledCheck=false){
  const center=STAFF_CENTERS[staffSlug];
  if(!center) return;
  if(!skipInstalledCheck&&await workcenterIsInstalled()){
    $('installCard').classList.remove('on');
    toast('이미 설치된 업무센터입니다');
    return;
  }
  $('installTitle').textContent=`${center.title}를 앱으로 설치`;
  const ios=/iPhone|iPad|iPod/i.test(navigator.userAgent);
  $('installGuide').textContent=ios
    ? 'Safari 공유 버튼 → “홈 화면에 추가”를 누르면 고정 아이콘으로 설치됩니다.'
    : 'PC 바탕화면이나 모바일 홈 화면에 전용 아이콘을 만들 수 있습니다.';
  $('installCard').classList.add('on');
}
async function installWorkcenter(){
  if(deferredInstallPrompt){
    deferredInstallPrompt.prompt();
    const choice=await deferredInstallPrompt.userChoice.catch(()=>({outcome:'dismissed'}));
    if(choice.outcome==='accepted'){
      deferredInstallPrompt=null; window.__csosInstallPrompt=null;
      $('installCard').classList.remove('on');
    }
    return;
  }
  const ua=navigator.userAgent;
  if(/Windows/i.test(ua)&&staffSlug){
    $('installGuide').textContent='PC 바탕화면에 업무센터를 설치하고 있습니다…';
    try{
      const result=await api('/api/staff/install-shortcut',{
        method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({staff_slug:staffSlug})
      });
      if(!result.ok) throw new Error(result.error||'설치 실패');
      localStorage.setItem('csos_installed_'+staffSlug,'1');
      $('installGuide').textContent='설치 완료! 바탕화면의 CSOS 업무센터 아이콘을 눌러 실행하세요.';
      setTimeout(()=>$('installCard').classList.remove('on'),2200);
      return;
    }catch(e){
      $('installGuide').textContent='자동 설치를 완료하지 못했습니다. Chrome/Edge 메뉴 → 앱 → 이 사이트를 앱으로 설치를 선택해 주세요.';
      return;
    }
  }
  $('installGuide').textContent=/iPhone|iPad|iPod/i.test(ua)
    ? 'Safari 공유 버튼 → “홈 화면에 추가”를 선택해 주세요.'
    : 'Chrome 메뉴(⋮) → “홈 화면에 추가” 또는 “앱 설치”를 선택해 주세요.';
}
function dismissInstall(){
  if(staffSlug) localStorage.setItem('csos_install_dismissed_'+staffSlug,'1');
  $('installCard').classList.remove('on');
}
function fixedStaffUrl(){
  return staffSlug?`https://mulder.tailf14aae.ts.net/staff/${staffSlug}`:location.href;
}
async function copyStaffCenterUrl(){
  try{await navigator.clipboard.writeText(fixedStaffUrl());toast('고정 업무센터 주소를 복사했습니다')}
  catch(e){prompt('아래 주소를 복사해 주세요',fixedStaffUrl())}
}
async function shareStaffCenter(){
  const center=STAFF_CENTERS[staffSlug]||{title:'CSOS 업무센터'};
  const data={title:center.title,text:`${center.title} 고정 접속 주소`,url:fixedStaffUrl()};
  if(navigator.share){try{await navigator.share(data);return}catch(e){if(e.name==='AbortError')return}}
  copyStaffCenterUrl();
}
function initStaffCenter(){
  if(!staffSlug) return;
  const center=STAFF_CENTERS[staffSlug];
  document.body.classList.add('staff-mode','ryu-mode');
  document.body.classList.toggle('staff-ryu',staffSlug==='ryu-jiyeong');
  document.title=center.title+' · CSOS';
  if($('staffCenterTitle')) $('staffCenterTitle').textContent=center.title;
  if($('staffCenterEyebrow')) $('staffCenterEyebrow').textContent=`${center.name} OPERATIONS CENTER`;
  if($('staffListToolbar')) $('staffListToolbar').setAttribute('aria-label',`${center.name} 업무 목록 도구`);
  const list=$('staffChecklistItems');
  if(list) list.innerHTML=center.checklist.map((x,i)=>{
    const key=`csos_check_${staffSlug}_${todayISO()}_${i}`;
    return `<label><input type="checkbox" ${localStorage.getItem(key)==='1'?'checked':''}
      onchange="localStorage.setItem('${key}',this.checked?'1':'0')"> ${esc2(x)}</label>`;
  }).join('');
  if($('improvementStaffSlug')) $('improvementStaffSlug').value=staffSlug;
  if($('poSubmissionForm')) $('poSubmissionForm').style.display=staffSlug==='oh-jonghyeon'?'block':'none';
  if($('ryuOnlySubmission')) $('ryuOnlySubmission').style.display=staffSlug==='ryu-jiyeong'?'block':'none';
  if($('workLogReportDate')) $('workLogReportDate').value=todayISO();
  document.querySelectorAll('#newWorkForm [name="submitter"]').forEach(x=>x.value=center.name);
  const brand=document.querySelector('.tabbar .brand .bt');
  if(brand) brand.dataset.staff=center.title;
  if(staffSlug==='ryu-jiyeong') loadWorkLogStatus();
  setTimeout(()=>maybeShowInstallCard(),1800);
}
async function staffHeartbeat(event='view'){
  if(!staffSlug||!PIN||document.hidden) return;
  try{await api('/api/staff/activity',{method:'POST',body:JSON.stringify({staff_slug:staffSlug,event})});}catch(e){}
}
let staffInputPulse=0;
function markStaffInput(event='input'){
  if(!staffSlug||!PIN) return;
  const now=Date.now();
  if(now-staffInputPulse<1200) return;
  staffInputPulse=now;
  staffHeartbeat(event);
}
let RYU=null, ryuCategory='as', ryuSelected=null, ryuVisible=[];
async function loadRyuRecords(force=false){
  if(RYU&&!force){ renderRyuCenter(); return; }
  try{
    RYU=await api('/api/ryu/records?ts='+Date.now());
    renderRyuCenter();
  }catch(e){
    const list=$('ryuHistoryList');
    if(list) list.innerHTML='<div class="ryu-result err" style="display:block">업무 목록을 불러오지 못했습니다.</div>';
  }
}
function ryuCategoryInfo(key){
  return ((RYU&&RYU.categories)||[]).find(x=>x.key===key)||{key,label:key,count:0,attention:0};
}
function renderRyuCenter(){
  if(!RYU) return;
  const summary=$('ryuSummaryGrid'), tabs=$('ryuCategoryTabs');
  const mainCats=(RYU.categories||[]).filter(x=>x.key!=='upload');
  if(summary) summary.innerHTML=mainCats.map(x=>`
    <button class="ryu-summary-card" type="button" onclick="selectRyuCategory('${x.key}')">
      <small>${esc2(x.label)}</small><strong>${fmt(x.count)}건</strong>
      <span class="attention">${x.attention?`확인·진행 ${fmt(x.attention)}건`:'현재 특이사항 없음'}</span>
    </button>`).join('');
  if(tabs) tabs.innerHTML=(RYU.categories||[]).map(x=>`
    <button class="ryu-category-tab ${x.key===ryuCategory?'on':''}" type="button"
      onclick="selectRyuCategory('${x.key}')"><span>${esc2(x.label)}</span><em>${fmt(x.count)}건</em></button>`).join('');
  if($('ryuUpdatedAt')) $('ryuUpdatedAt').textContent=`· 데이터 ${String(RYU.updated_at||'').replace('T',' ')}`;
  selectRyuCategory(ryuCategory,true);
}
function selectRyuCategory(key,keep=false){
  if(!RYU) return;
  ryuCategory=key;
  document.querySelectorAll('.ryu-category-tab').forEach(b=>b.classList.toggle('on',
    (b.getAttribute('onclick')||'').includes(`'${key}'`)));
  const upload=key==='upload';
  if($('ryuWorkspace')) $('ryuWorkspace').style.display=upload?'none':'grid';
  if($('ryuUploadPanel')) $('ryuUploadPanel').classList.toggle('on',upload);
  if(upload) return;
  const info=ryuCategoryInfo(key);
  if($('ryuHistoryTitle')) $('ryuHistoryTitle').textContent=`${info.label} 과거 목록`;
  const statuses=[...new Set(((RYU.rows||{})[key]||[]).map(x=>x.status).filter(Boolean))].sort();
  const status=$('ryuHistoryStatus');
  if(status){
    const old=keep?status.value:'';
    status.innerHTML='<option value="">상태 전체</option>'+statuses.map(x=>`<option>${esc2(x)}</option>`).join('');
    if(statuses.includes(old)) status.value=old;
  }
  if(!keep){ ryuSelected=null; if($('ryuHistoryQuery')) $('ryuHistoryQuery').value=''; }
  renderRyuHistory();
  renderRyuEditor();
}
function ryuPeriodPass(day,period){
  if(!period) return true;
  if(!day) return false;
  const now=new Date(), d=new Date(day+'T00:00:00');
  if(Number.isNaN(d.getTime())) return false;
  const monthStart=new Date(now.getFullYear(),now.getMonth(),1);
  const recentStart=new Date(now.getFullYear(),now.getMonth()-2,1);
  if(period==='month') return d>=monthStart;
  if(period==='recent3') return d>=recentStart;
  if(period==='older') return d<recentStart;
  return true;
}
function ryuIsAttention(row){
  return !['정상','완료','작업완료','발행완료','원본 저장'].includes(String(row.status||'').trim());
}
function renderRyuHistory(){
  if(!RYU||ryuCategory==='upload') return;
  const q=(($('ryuHistoryQuery')||{}).value||'').trim().toLowerCase();
  const status=(($('ryuHistoryStatus')||{}).value||'');
  const period=(($('ryuHistoryPeriod')||{}).value||'');
  const allRows=((RYU.rows||{})[ryuCategory]||[]);
  const center=STAFF_CENTERS[staffSlug]||null;
  const rows=(center&&staffSlug!=='ryu-jiyeong')
    ? allRows.filter(r=>String(r.assignee||'').replace(/\s+/g,'').includes(String(center.assignee||'').replace(/\s+/g,'')))
    : allRows;
  ryuVisible=rows.filter(r=>{
    if(status&&r.status!==status) return false;
    if(!ryuPeriodPass(r.date,period)) return false;
    if(!q) return true;
    return JSON.stringify(r).toLowerCase().includes(q);
  });
  if($('ryuHistoryCount')) $('ryuHistoryCount').textContent=`표시 ${fmt(ryuVisible.length)}건 / 전체 ${fmt(rows.length)}건`;
  const box=$('ryuHistoryList');
  if(!box) return;
  if(!ryuVisible.length){
    box.innerHTML='<div class="ryu-editor-empty">조건에 맞는 업무가 없습니다.</div>';return;
  }
  box.innerHTML=ryuVisible.map((r,i)=>`
    <button type="button" class="ryu-history-row ${ryuSelected&&ryuSelected.key===r.key?'on':''}"
      onclick="selectRyuRecord(${i})">
      <span class="main">
        <span class="title">${esc2(r.project_no||r.key||'번호 없음')} · ${esc2(r.camp||'캠프 미입력')}</span>
        <span class="sub">${esc2(r.key||'')} ${r.assignee?'· '+esc2(r.assignee):''}<br>${esc2(r.summary||'입력 내용 없음')}</span>
        <span class="state ${ryuIsAttention(r)?'warn':''}">${esc2(r.status||'상태 미입력')}</span>
      </span>
      <span class="date">${esc2(r.date||'날짜 미입력')}</span>
    </button>`).join('');
}
function ryuRowsForExport(){
  const label=ryuCategoryInfo(ryuCategory).label||ryuCategory;
  return ryuVisible.map(r=>({
    프로젝트NO:r.project_no||'',업무ID:r.key||'',캠프명:r.camp||'',구분:label,
    확인사항:r.summary||'',현재상태:r.status||'',기준일자:r.date||'',담당자:r.assignee||''
  }));
}
function ryuListFileName(ext='png'){
  const label=String(ryuCategoryInfo(ryuCategory).label||'업무목록').replace(/[\\/:*?"<>|\s·]+/g,'_');
  const who=(STAFF_CENTERS[staffSlug]||{name:'류지영'}).name.replace(/[()]/g,'');
  return `CSOS_${who}_${label}_${todayISO()}.${ext}`;
}
async function ryuListToPng(){
  const all=ryuVisible||[], cap=40, rows=all.slice(Math.max(0,all.length-cap));
  const S=2,W=820,rowH=52,headH=104,footH=62,H=headH+42+Math.max(1,rows.length)*rowH+footH;
  const cv=document.createElement('canvas');cv.width=W*S;cv.height=H*S;
  const ctx=cv.getContext('2d');ctx.scale(S,S);
  const F=(s,w)=>(w||400)+' '+s+'px "Nanum Gothic","Malgun Gothic",sans-serif';
  const txt=(t,x,y,f,c,a)=>{ctx.font=f;ctx.fillStyle=c;ctx.textAlign=a||'left';
    ctx.fillText(String(t==null?'':t),x,y);ctx.textAlign='left';};
  const clip=(t,max)=>{let s=String(t==null?'':t);if(ctx.measureText(s).width<=max)return s;
    while(s.length>1&&ctx.measureText(s+'…').width>max)s=s.slice(0,-1);return s+'…';};
  ctx.fillStyle='#fff';ctx.fillRect(0,0,W,H);
  const g=ctx.createLinearGradient(0,0,W,headH);g.addColorStop(0,'#101B47');g.addColorStop(1,'#4959E6');
  ctx.fillStyle=g;ctx.fillRect(0,0,W,headH);await drawLogo(ctx,26,25,38,true);
  const label=ryuCategoryInfo(ryuCategory).label||'업무 목록';
  txt(((STAFF_CENTERS[staffSlug]||{title:'류지영 쿠팡 AS 및 정기점검 업무센터'}).title)+' · '+label,82,46,F(20,900),'#fff');
  txt(`2026년 · 현재 필터 ${all.length}건 · 과거 → 최근`,82,70,F(11.5,700),'#CAD5FF');
  txt(todayISO(),W-26,46,F(12,800),'#fff','right');
  let y=headH+28;
  txt('프로젝트NO / 업무ID · 캠프 · 내용',26,y,F(10.5,800),'#667085');
  txt('상태',W-150,y,F(10.5,800),'#667085','right');txt('일자',W-26,y,F(10.5,800),'#667085','right');
  y+=14;ctx.strokeStyle='#D4DCEB';ctx.beginPath();ctx.moveTo(26,y);ctx.lineTo(W-26,y);ctx.stroke();
  const list=rows.length?rows:[{project_no:'—',key:'',camp:'해당 건 없음',summary:'',status:'',date:'-'}];
  list.forEach((r,i)=>{
    const top=y+i*rowH;
    if(i%2){ctx.fillStyle='#F6F8FC';ctx.fillRect(26,top+3,W-52,rowH-6);}
    txt(clip(`${r.project_no||r.key||'번호 없음'} · ${r.camp||'캠프 미입력'}`,430),
      30,top+22,F(12.5,850),'#101828');
    txt(clip(`${r.key||''} ${r.assignee?'· '+r.assignee:''} ${r.summary?'· '+r.summary:''}`,430),
      30,top+41,F(10.5),'#667085');
    txt(clip(r.status||'상태 미입력',150),W-150,top+29,F(11.5,800),
      ryuIsAttention(r)?'#B42318':'#16794A','right');
    txt(r.date||'-',W-26,top+29,F(10.5),'#667085','right');
  });
  y+=list.length*rowH+24;
  if(all.length>cap) txt(`※ 최근 ${cap}건 표시 · 전체 ${all.length}건은 엑셀 저장에서 확인`,26,y,F(10.5,800),'#B54708');
  txt('데이터 업데이트 '+String((RYU&&RYU.updated_at)||'').replace('T',' '),W-26,y+20,F(10),'#98A2B3','right');
  return new Promise(res=>cv.toBlob(res,'image/png'));
}
async function ryuCaptureList(){
  try{const b=await ryuListToPng();if(!b)throw new Error('빈 이미지');saveOrOpen(b,ryuListFileName());}
  catch(e){alert('이미지 저장 실패: '+e);}
}
async function ryuCopyList(){
  try{const b=await ryuListToPng();if(!b)throw new Error('빈 이미지');await copyPngBlob(b);}
  catch(e){alert('이미지 복사 실패: '+e);}
}
function ryuExportList(){
  exportRowsXlsx(`류지영_${ryuCategoryInfo(ryuCategory).label||'업무목록'}`,ryuRowsForExport());
}
function selectRyuRecord(index){
  ryuSelected=ryuVisible[index]||null;
  renderRyuHistory();
  renderRyuEditor();
  if(window.innerWidth<900&&$('ryuEntryForm')) $('ryuEntryForm').scrollIntoView({behavior:'smooth',block:'start'});
}
function ryuInputHtml(field,existing){
  const name=esc2(field.name), label=esc2(field.label||field.name);
  const has=existing!==undefined&&existing!==null&&String(existing).trim()!=='';
  const disabled=has?' disabled':'';
  const note=has?' · 기존 입력 완료':'';
  if(field.type==='select'){
    const opts=['',...(field.options||[])];
    return `<label>${label}${note}<select name="${name}"${disabled}>${opts.map(x=>
      `<option value="${esc2(x)}"${has&&String(existing)===String(x)?' selected':''}>${esc2(x||'선택 안 함')}</option>`).join('')}</select></label>`;
  }
  if(field.type==='textarea'){
    return `<label class="wide">${label}${note}<textarea name="${name}"${disabled}
      placeholder="${has?'':label+' 입력'}">${has?esc2(existing):''}</textarea></label>`;
  }
  return `<label>${label}${note}<input name="${name}" type="${field.type==='date'?'date':field.type==='number'?'number':'text'}"
    ${field.type==='number'?'step="any" ':''}value="${has?esc2(existing):''}"${disabled} placeholder="${has?'':label+' 입력'}"></label>`;
}
function renderRyuEditor(){
  const empty=$('ryuEditorEmpty'), body=$('ryuEditorBody');
  if(!empty||!body) return;
  if(!ryuSelected){
    empty.style.display='block';body.style.display='none';return;
  }
  empty.style.display='none';body.style.display='block';
  const r=ryuSelected, info=ryuCategoryInfo(ryuCategory);
  $('ryuEntryCategory').value=ryuCategory;$('ryuEntryKey').value=r.key||'';
  $('ryuTargetCategory').value=r.target_category||'';$('ryuTargetKey').value=r.target_key||'';
  $('ryuSelectedSummary').innerHTML=`<div class="key">${esc2(r.project_no||r.key||'번호 없음')}</div>
    <div class="camp">${esc2(r.camp||'캠프 미입력')} · ${esc2(r.date||'날짜 미입력')} · ${esc2(r.status||'상태 미입력')}</div>
    <div class="muted" style="margin-top:6px">${esc2(r.summary||'')}</div>`;
  const detail=r.detail||{};
  $('ryuSelectedDetail').innerHTML=Object.entries(detail).map(([k,v])=>`
    <div class="ryu-detail-item"><small>${esc2(k)}</small><span title="${esc2(v)}">${esc2(v)}</span></div>`).join('')
    ||'<div class="muted">기존 입력값이 없습니다.</div>';
  const schema=((RYU.schema||{})[ryuCategory]||{}).fields||[];
  const fields=$('ryuEntryFields'), save=$('ryuEntrySaveBtn');
  if(!r.editable){
    fields.innerHTML='<div class="ryu-editor-empty" style="grid-column:1/-1">이 건은 ERP 보완·원본 일정·자동 집계 건이라 이 화면에서 직접 수정하지 않습니다. 목록 확인용으로만 표시합니다.</div>';
    save.style.display='none';
  }else{
    fields.innerHTML=schema.map(f=>ryuInputHtml(f,detail[f.name])).join('');
    save.style.display='flex';
  }
  const result=$('ryuEntryResult');result.className='ryu-result';result.textContent='';
}
async function submitRyuEntry(ev){
  ev.preventDefault();
  markStaffInput('save');
  const form=$('ryuEntryForm'),btn=$('ryuEntrySaveBtn'),result=$('ryuEntryResult');
  if(!ryuSelected||!ryuSelected.editable) return;
  btn.disabled=true;btn.classList.add('busy');result.className='ryu-result';result.textContent='';
  try{
    const r=await fetch('/api/ryu/entry',{method:'POST',cache:'no-store',headers:{'X-Pin':PIN},body:new FormData(form)});
    if(r.status===401){PIN='';localStorage.removeItem('cw_pin');$('gate').style.display='flex';throw new Error('PIN을 다시 입력해 주세요');}
    const d=await r.json();if(!r.ok||!d.ok) throw new Error(d.error||`HTTP ${r.status}`);
    result.className='ryu-result ok';
    result.textContent=d.queued?`${d.queued}개 항목을 안전 입력 대기열에 넣었습니다. 관리대장 반영이 이어집니다.`:'근거 파일과 조사 기록을 저장했습니다.';
    toast('류지영 업무센터 입력을 접수했습니다');
  }catch(e){
    result.className='ryu-result err';result.textContent='저장 실패 · '+String(e.message||e);
  }finally{btn.disabled=false;btn.classList.remove('busy');}
}
async function submitRyuUpload(ev){
  ev.preventDefault();
  markStaffInput('upload');
  const form=$('ryuUploadForm'), btn=$('ryuUploadBtn'), result=$('ryuUploadResult');
  if(!form||!btn||!result) return;
  const data=new FormData(form);
  btn.disabled=true;btn.classList.add('busy');
  result.className='ryu-result';result.textContent='';
  try{
    const r=await fetch('/api/ryu/upload',{method:'POST',cache:'no-store',
      headers:{'X-Pin':PIN},body:data});
    if(r.status===401){
      PIN='';localStorage.removeItem('cw_pin');$('gate').style.display='flex';
      throw new Error('PIN을 다시 입력해 주세요');
    }
    const d=await r.json();
    if(!r.ok||!d.ok) throw new Error(d.error||`HTTP ${r.status}`);
    const files=(d.saved&&d.saved['파일'])||[];
    result.className='ryu-result ok';
    result.innerHTML=`저장 완료 · ${files.map(x=>`${esc2(x['방'])} ${fmt(x['메시지'])}개 메시지`).join(' · ')}
      <br>${d.auto_check_started?'자동 카톡 대조를 시작했습니다.':
        (d.auto_check_queued?'현재 작업이 끝나는 즉시 자동 카톡 대조를 시작합니다.':
         '자동 대조 예약 상태를 확인해 주세요.')}`;
    [...form.querySelectorAll('input[type=file]')].forEach(x=>x.value='');
    toast('류지영 업무센터 원본 저장 및 자동 대조 접수 완료');
  }catch(e){
    result.className='ryu-result err';result.textContent='업로드 실패 · '+String(e.message||e);
  }finally{
    btn.disabled=false;btn.classList.remove('busy');
  }
}
async function submitNewWork(ev){
  ev.preventDefault();
  markStaffInput('save');
  const form=ev.currentTarget, btn=$('newWorkBtn'), result=$('newWorkResult');
  btn.disabled=true;btn.classList.add('busy');
  result.className='ryu-result';result.textContent='';
  try{
    const r=await fetch('/api/staff/new-job',{method:'POST',headers:{'X-Pin':PIN},body:new FormData(form)});
    const d=await r.json();
    if(!r.ok||!d.ok) throw new Error(d.error||'신규 업무 저장 실패');
    result.className='ryu-result ok';
    result.innerHTML=`저장 완료 · ${esc2(d.manifest['업무ID'])} · 자동 반영 ${fmt(d.queued)}개 항목
      <br>${d.applying?'관리대장 반영을 시작했습니다.':'현재 작업이 끝나는 즉시 자동 반영합니다.'}`;
    const keepDate=form.querySelector('[name="work_date"]').value;
    form.reset();form.querySelector('[name="work_date"]').value=keepDate;
    await Promise.all([loadRyuRecords(true),loadNotifications()]);
    toast('신규 업무 저장·자동 반영 접수 완료');
  }catch(e){
    result.className='ryu-result err';result.textContent='저장 실패 · '+String(e.message||e);
    loadNotifications();
  }finally{btn.disabled=false;btn.classList.remove('busy')}
}
function setSourceFile(inputId,nameId,file){
  if(!file) return;
  const input=$(inputId), label=$(nameId);
  if(!input) return;
  const dt=new DataTransfer();dt.items.add(file);input.files=dt.files;
  if(label) label.textContent=`선택됨 · ${file.name}`;
}
async function submitPoSubmission(ev){
  ev.preventDefault();
  markStaffInput('upload');
  const form=ev.currentTarget,btn=$('poSubmissionBtn'),result=$('poSubmissionResult');
  btn.disabled=true;btn.classList.add('busy');result.className='ryu-result';result.textContent='';
  try{
    const r=await fetch('/api/staff/po-upload',{method:'POST',cache:'no-store',
      headers:{'X-Pin':PIN},body:new FormData(form)});
    const d=await r.json();
    if(!r.ok||!d.ok) throw new Error(d.error||'PO 원본 등록 실패');
    result.className='ryu-result ok';
    result.innerHTML=`원본 ${fmt((d.files||[]).length)}개 저장 완료
      ${d.po_compare_files&&d.po_compare_files.length
        ?`· PO 자동대조 ${d.auto_check_started?'시작':'대기열 등록'}`
        :'· Excel(.xlsx)이 없어 원본 보관만 완료'}`;
    form.querySelector('[name="po_file"]').value='';
    if($('poFileName')) $('poFileName').textContent='선택된 파일 없음';
    await Promise.all([loadNotifications(),loadRyuRecords(true)]);
    toast('오종현 PO 원본 저장을 완료했습니다');
  }catch(e){
    result.className='ryu-result err';result.textContent='등록 실패 · '+String(e.message||e);
  }finally{btn.disabled=false;btn.classList.remove('busy')}
}
async function loadWorkLogStatus(){
  const box=$('workLogStatus');if(!box||!PIN) return;
  try{
    const d=await api('/api/staff/work-log-status?t='+Date.now());
    if(!d.ok) throw new Error(d.error||'상태 확인 실패');
    const as=(d.summary||{})['돌발AS']||{}, pm=(d.summary||{})['정기점검']||{};
    box.innerHTML=`
      <div><small>현재 원본</small><strong title="${esc2(d.source_name)}">${esc2(d.source_name||'-')}</strong></div>
      <div><small>마지막 검증</small><strong>${esc2(String(d.verified_at||'-').replace('T',' '))}</strong></div>
      <div><small>돌발AS 미처리</small><strong>${fmt(as['미처리']||0)}건</strong></div>
      <div><small>정기점검 실행</small><strong>${fmt(pm['실행']||0)}건</strong></div>`;
  }catch(e){
    box.innerHTML=`<div style="grid-column:1/-1"><small>상태</small><strong>확인 실패 · ${esc2(String(e.message||e))}</strong></div>`;
  }
}
async function postWorkLog(data){
  const r=await fetch('/api/staff/work-log-upload',{method:'POST',cache:'no-store',
    headers:{'X-Pin':PIN},body:data});
  const d=await r.json();
  if(!r.ok||!d.ok) throw new Error(d.error||'대표보고 일지 검증 실패');
  return d;
}
async function submitWorkLog(ev){
  ev.preventDefault();
  markStaffInput('upload');
  const form=ev.currentTarget,btn=$('workLogUploadBtn'),result=$('workLogUploadResult');
  btn.disabled=true;btn.classList.add('busy');result.className='ryu-result';result.textContent='';
  try{
    const d=await postWorkLog(new FormData(form));
    result.className='ryu-result ok';
    result.textContent=`원본 ${fmt((d.files||[]).length)}개 저장 · 전체 검증 ${d.auto_check_started?'시작':'대기열 등록'} 완료`;
    form.querySelector('[name="work_log_file"]').value='';
    if($('workLogFileName')) $('workLogFileName').textContent='선택된 파일 없음';
    await Promise.all([loadWorkLogStatus(),loadNotifications()]);
    toast('대표보고 일지 원본 저장·검증을 접수했습니다');
  }catch(e){
    result.className='ryu-result err';result.textContent='등록 실패 · '+String(e.message||e);
  }finally{btn.disabled=false;btn.classList.remove('busy')}
}
async function syncCurrentWorkLog(){
  markStaffInput('submit');
  const result=$('workLogUploadResult');
  try{
    const data=new FormData();
    data.set('staff_slug','ryu-jiyeong');data.set('use_current','1');
    const d=await postWorkLog(data);
    if(result){result.className='ryu-result ok';
      result.textContent=`현재 고정 원본 검증 ${d.auto_check_started?'시작':'대기열 등록'} 완료`;}
    await Promise.all([loadWorkLogStatus(),loadNotifications()]);
    toast('현재 대표보고 일지 재검증을 접수했습니다');
  }catch(e){
    if(result){result.className='ryu-result err';result.textContent='검증 실패 · '+String(e.message||e);}
  }
}
async function openWorkLogReport(andCapture){
  const day=($('workLogReportDate')&&$('workLogReportDate').value)||todayISO();
  if(!/^2026-\d{2}-\d{2}$/.test(day)){alert('2026년 보고 기준일을 선택해 주세요');return;}
  try{
    const [rep,b]=await Promise.all([
      api('/api/exec_report?date='+encodeURIComponent(day)),
      api('/api/brief?date='+encodeURIComponent(day))
    ]);
    execRep=filterExec(rep||{});BRIEF=b&&b.ok!==false?b:null;REPORT_PREVIEW_DATE=day;
    show('daily');renderDaily();
    if(andCapture){await new Promise(resolve=>requestAnimationFrame(resolve));await captureReport();}
  }catch(e){alert('대표 보고를 만들지 못했습니다: '+String(e.message||e));}
}
async function submitImprovement(ev){
  ev.preventDefault();
  const form=ev.currentTarget,btn=$('improvementBtn'),result=$('improvementResult');
  btn.disabled=true;btn.classList.add('busy');
  try{
    const r=await fetch('/api/staff/improvement',{method:'POST',headers:{'X-Pin':PIN},body:new FormData(form)});
    const d=await r.json();
    if(!r.ok||!d.ok) throw new Error(d.error||'개선 요청 등록 실패');
    result.className='ryu-result ok';
    result.textContent=`등록 완료 · ${d.ticket.id} · ${d.ticket.route}`;
    form.querySelector('[name="title"]').value='';
    form.querySelector('[name="description"]').value='';
    $('improvementAttachment').value='';$('improvementFileName').textContent='파일을 선택하거나 캡처 후 Ctrl+V';
    toast('개선 요청을 AI 검토 대기열에 등록했습니다');
  }catch(e){
    result.className='ryu-result err';result.textContent='등록 실패 · '+String(e.message||e);
  }finally{btn.disabled=false;btn.classList.remove('busy')}
}
function setImprovementFile(file){
  if(!file) return;
  const dt=new DataTransfer();dt.items.add(file);
  $('improvementAttachment').files=dt.files;
  $('improvementFileName').textContent=`첨부됨 · ${file.name}`;
}
document.addEventListener('paste',e=>{
  if(!staffSlug) return;
  const files=[...(e.clipboardData&&e.clipboardData.files||[])];
  const file=files[0];
  if(e.target.closest&&e.target.closest('#poDrop')&&file){
    e.preventDefault();setSourceFile('poFile','poFileName',file);toast('PO 파일을 붙였습니다');return;
  }
  if(e.target.closest&&e.target.closest('#workLogDrop')&&file){
    e.preventDefault();setSourceFile('workLogFile','workLogFileName',file);toast('대표보고 일지를 붙였습니다');return;
  }
  const image=files.find(x=>x.type.startsWith('image/'));
  if(image&&$('improvementForm')){setImprovementFile(image);toast('캡처 이미지를 개선 요청에 붙였습니다')}
});
document.addEventListener('dragover',e=>{
  if(!staffSlug||!e.dataTransfer||![...e.dataTransfer.types].includes('Files')) return;
  e.preventDefault();
  const target=e.target.closest&&e.target.closest('#poDrop,#workLogDrop,#improvementDrop');
  if(target) target.classList.add('drag');
});
document.addEventListener('dragleave',e=>{
  const target=e.target.closest&&e.target.closest('#poDrop,#workLogDrop,#improvementDrop');
  if(target) target.classList.remove('drag');
});
document.addEventListener('drop',e=>{
  if(!staffSlug||!e.dataTransfer||!e.dataTransfer.files.length) return;
  const target=e.target.closest&&e.target.closest('#poDrop,#workLogDrop,#improvementDrop');
  if(!target) return;
  e.preventDefault();target.classList.remove('drag');
  const file=e.dataTransfer.files[0];
  if(target.id==='poDrop') setSourceFile('poFile','poFileName',file);
  else if(target.id==='workLogDrop') setSourceFile('workLogFile','workLogFileName',file);
  else setImprovementFile(file);
});
async function login(){
  const pin = $('pin').value.trim();
  const r = await fetch('/api/login',{method:'POST',body:JSON.stringify({pin,staff_slug:staffSlug||''})});
  if(r.ok){ PIN=pin; localStorage.setItem('cw_pin',pin); $('gate').style.display='none';
            show(curView()); refreshAll(); }
  else if(r.status===429) $('pinerr').textContent = '시도 초과로 잠금 — 10분 후 다시 시도하세요';
  else $('pinerr').textContent = 'PIN이 올바르지 않습니다';
}
$('pin').addEventListener('keyup',e=>{ if(e.key==='Enter'||$('pin').value.length===4) login(); });

async function restoreRoleSession(){
  if(!PIN) return false;
  try{
    const r=await fetch('/api/login',{
      method:'POST',
      body:JSON.stringify({pin:PIN,staff_slug:staffSlug||''})
    });
    if(!r.ok) throw new Error('session');
    const d=await r.json();
    return d.role===(staffSlug?'staff':'admin') &&
           (!staffSlug||d.staff_slug===staffSlug);
  }catch(e){
    PIN=''; localStorage.removeItem('cw_pin');
    $('gate').style.display='flex';
    return false;
  }
}

/* 화면만 바꾼다 — 뒤로가기 기록은 건드리지 않는다(복귀·popstate 처리에서 쓴다). */
function applyView(v){
  document.querySelectorAll('.view').forEach(x=>x.classList.remove('active'));
  $('v-'+v).classList.add('active');
  const navView=v==='ryu'?'dash':v;
  document.querySelectorAll('.tabbar button').forEach(b=>b.classList.toggle('on', b.dataset.v===navView));
  if(v==='report' && !reports.length) loadReports();
  if(v==='ryu' && PIN && !RYU) loadRyuRecords();
  if(v==='worklog' && PIN) loadWorkLogStatus();
  if(v==='dash' && !CAL) loadCalendar();
  if(v==='calendar'){
    if(!CAL) loadCalendar().then(renderCalendarPage);
    else renderCalendarPage();
  }
  if(v==='check') renderCheckHub();
  try{ localStorage.setItem('cw_view', v); }catch(e){}   // 새로고침 후 같은 화면으로 복귀
  window.__view = v;
}
/* 사용자가 화면을 옮긴다 — 어디서 왔는지 남겨 뒤로가기가 그리로 돌아가게 한다. */
function show(v){
  const from = curView();
  if(v !== from) navPush({t:'tab', from:from});
  applyView(v);
}
function routeNav(v){
  if(document.body.classList.contains('ryu-mode')&&v==='dash'){show('ryu');return;}
  show(v);
  if(v==='daily') renderDaily();
}
function curView(){
  if(staffSlug) return window.__view || 'ryu';
  const v=window.__view || localStorage.getItem('cw_view') || 'dash';
  return v==='ryu' ? 'dash' : v;
}
/* ── 구버전 감지 ─────────────────────────────────────────────
   폰에서 앱을 켜 두면 화면은 그대로라 서버가 바뀌어도 예전 HTML·JS가 계속 돈다.
   (데이터만 새로 받는 새로고침으로는 화면 구조가 안 바뀐다)
   서버의 build 값이 달라지면 문서 자체를 다시 읽는다. */
let BUILD = '';
async function checkBuild(auto){
  try{
    const r = await fetch('/api/ping?t=' + Date.now(), {cache:'no-store'});
    const d = await r.json();
    if(!d.build) return false;
    if(!BUILD){ BUILD = d.build; return false; }
    if(d.build === BUILD) return false;
    if(auto){
      if(polling || $('sheetbg').style.display === 'block') return true;  // 작업 중엔 미룸
      toast('새 버전을 받았습니다 · 화면을 갱신합니다');
      setTimeout(hardReload, 700);
    }
    return true;
  }catch(e){ return false; }
}
/* 캐시를 확실히 비우고 문서를 다시 읽는다 */
function hardReload(){
  const u = new URL(location.href);
  u.searchParams.set('v', Date.now());
  location.replace(u.toString());
}
/* 새로고침: 데이터를 다시 읽고 **보고 있던 화면 그대로** 다시 그린다 */
async function reloadHere(btn){
  const v = curView();
  if(btn){ btn.disabled = true; btn.classList.add('busy'); }
  // 서버가 새 버전이면 데이터만 받아선 화면이 안 바뀐다 → 문서를 다시 읽는다
  if(await checkBuild(false)){ hardReload(); return; }
  try{ await refreshAll(); }catch(e){}
  try{ renderDaily(); }catch(e){}
  try{ renderBoard(); }catch(e){}
  try{ if(typeof renderPeriod==='function' && $('pyear').value) renderPeriod(); }catch(e){}
  renderSettle();
  show(v);
  if(btn){ btn.disabled = false; btn.classList.remove('busy'); }
  toast('최신 자료로 새로고침했습니다');
}

/* ── 대시보드 ── */
async function loadStatus(){
  const s = await api('/api/status');
  srcStats = s.sources || {};
  const route = s.agent_dispatch || {};
  const routeEl = $('runAgentNote');
  if(routeEl){
    const selected = route.selected || 'codex_pending';
    const label = selected === 'claude' ? 'Claude Code 우선' :
      selected === 'codex' ? 'Claude Code 사용 불가 · Codex 대행' :
      'Claude Code 사용 불가 · Codex 요청 대기';
    routeEl.innerHTML = `<b>AI 실행 연계</b> ${esc2(label)} · 실행 버튼은 기존 로컬 업무 스크립트를 한 번만 실행합니다.`;
  }
  /* ★ '내가 확인할 사항'은 loadSettle 에서 먼저 그려지는데, 그때는 srcStats 가 아직 비어 있다.
     그대로 두면 밴드·카톡·PO 자료가 멀쩡히 있어도 **'아직 없음' 4건이 거짓으로** 뜬다
  (2026-07-27에 실제로 그렇게 보였다). 상태가 도착했으니 다시 그린다. */
  try{ renderBoard(); }catch(e){ console.warn('board', e); }
  // 보고 화면을 먼저 연 직후 상태 API가 도착하면 4원천 표도 즉시 실제 집계로 다시 그린다.
  // 이 재렌더가 없으면 화면·캡처에 잠시 '자료 없음' 안내가 고정될 수 있다.
  try{ if(window.__view==='daily') renderDaily(); }catch(e){ console.warn('daily sources',e); }
  $('demobadge').textContent = s.demo ? '· 데모(합성데이터)' : '';
  const rows = settleRows.length ? settleRows : [];
  const 유상 = rows.filter(r=>r.비용구분==='유상');
  const 발행 = 유상.filter(r=>r.계산서==='발행').length;
  const 미수 = rows.reduce((a,r)=>a+(+r.미수금||0),0);
  const issues = rows.filter(r=>needAction(r));
  const asOpen = (works.as||[]).filter(r=>r.진행상태 && r.진행상태!=='작업완료').length;
  const pmWait = (works.pm||[]).filter(r=>r.점검상태 && r.점검상태!=='완료').length;
  const tile=(cls,ico,label,val,sub,click)=>`<div class="kpi ${cls}" onclick="kTap(this);${click||''}">
    <div class="head"><div class="ico">${ico}</div><div class="l">${label}</div></div>
    <div class="v">${val}</div>${sub?`<div class="s">${sub}</div>`:''}</div>`;
  $('kpis').innerHTML =
    tile('accent','📋','정산 건수', rows.length||'-', `유상 ${유상.length}건 · 클릭=목록`, "goList('settle','')") +
    tile('','💰','작업 공급가액', fmt(rows.reduce((a,r)=>a+(+r.공급가액||0),0)), '원 · 실제 작업분 · 클릭=내역', "goList('settle','')") +
    tile(미수?'warn':'ok','⏳','미수금', fmt(미수), '원 · 계산서 발행 후 미입금 · 클릭=목록', "goList('settle','입금 대기')") +
    tile(발행===유상.length&&유상.length?'ok':'','🧾','계산서 발행율',
      (유상.length?Math.round(발행/유상.length*100):0)+'%',
      `${발행}/${유상.length}건 <span class="bar"><i style="width:${유상.length?발행/유상.length*100:0}%"></i></span>`,
      "goList('settle','세금계산서 미발행')") +
    tile(issues.length?'danger':'ok','⚠️','조치 필요', issues.length, issues.length?'클릭해 아래 유형 확인':'모두 정상',
      "flashTop()") +
    tile(asOpen?'warn':'','🔧','돌발AS 진행중', asOpen, `전체 ${(works.as||[]).length}건 · 클릭=목록`, "goList('as','')") +
    tile(pmWait?'warn':'','🗓','정기점검 대기', pmWait, `전체 ${(works.pm||[]).length}건 · 클릭=목록`, "goList('pm','')") +
    (erpDocs.total ? tile('','🧾', APP_YEAR+'년 ERP 매출',
      fmt(erpDocs.total), `계산서 ${erpDocs.rows.length}장 · 클릭=보고`,
      "show('daily');renderDaily()") : '');
  // 상태 분포 바 (2px 간격 마크)
  const SCOLOR = {'정상':'#12813F','입금 대기':'#B54708','세금계산서 미발행':'#C0212E',
    '미청구(전표 없음)':'#8B1E68','금액 미입력':'#5B6B82','무상/보험':'#B9C3D3'};
  const dist = {};
  rows.forEach(r=>{ dist[r.상태]=(dist[r.상태]||0)+1; });
  const entries = Object.entries(dist).sort((a,b)=>b[1]-a[1]);
  $('distbar').innerHTML = entries.map(([k,v])=>
    `<i style="flex:${v};background:${SCOLOR[k]||'#5B6B82'}" title="${k} ${v}건"></i>`).join('');
  $('distlegend').innerHTML = entries.map(([k,v])=>
    `<span><b style="background:${SCOLOR[k]||'#5B6B82'}"></b>${k} ${v}</span>`).join('');
  // 조치 필요 TOP: 유형별 집계 + 상위 건
  const byType = {};
  issues.forEach(r=>{ (byType[r.상태]=byType[r.상태]||[]).push(r); });
  $('toplist').innerHTML = Object.entries(byType).sort((a,b)=>b[1].length-a[1].length).map(([t,list])=>{
    const amt = list.reduce((a,r)=>a+(+r.공급가액||0),0);
    const chips = list.slice(0,4).map(r=>{
      const prj=projectNoOf(r), key=prj||recordIdOf(r);
      return `<span class="prj" onclick="event.stopPropagation();openByPrj(${esc4(key)})"
        title="${esc2(recordIdOf(r))} · ${esc2(r.캠프명||'')} — 클릭해 상세">${esc2(prj||'프로젝트 미확정')}</span>`;
    }).join('');
    return `<div class="topitem" onclick="jumpIssue('${t}')">
      <span style="min-width:0"><b>${t}</b><div style="margin-top:5px;display:flex;flex-wrap:wrap;gap:4px">${chips}${
        list.length>4?`<span style="font-size:11px;color:var(--ink-3);align-self:center">외 ${list.length-4}건</span>`:''}</div></span>
      <span style="text-align:right;flex:none"><b>${list.length}건</b><br><span style="font-size:11.5px;color:var(--ink-3)">${fmt(amt)}원</span></span></div>`;
  }).join('') || '<div style="color:var(--ink-3);font-size:13px">조치 필요 건이 없습니다 🎉</div>';
  $('agentTime').textContent = '— ' + (s.agent_last||'');
  // 히어로 배너
  const now = new Date();
  $('heroDate').textContent = now.toLocaleDateString('ko-KR',{year:'numeric',month:'long',day:'numeric',weekday:'long'});
  $('heroAgent').textContent = '에이전트 ' + (s.agent_last ? s.agent_last.slice(5,16) + ' 실행' : '대기');
  const hs = (l,v)=>`<div class="hstat">${l}<b>${v}</b></div>`;
  $('heroStats').innerHTML =
    hs('정산 건수', (rows.length||0)+'건') +
    hs('작업 공급가액', fmt(rows.reduce((a,r)=>a+(+r.공급가액||0),0))+'원') +
    hs('조치 필요', issues.length+'건') +
    (srcStats.band ? hs('밴드 사진 확인', `${srcStats.band.total}건 중 ${srcStats.band.ok}건`) : '') +
    (s.tunnel ? hs('외부접속', '연결됨') : '');
  window._agentSteps = s.steps || [];          // 대시보드 캡처용
  const icon = {ok:'✅',skip:'⏭',fail:'❌'};
  $('steps').innerHTML = (s.steps||[]).map(t=>`<div class="step"><span class="ic">${icon[t.s]||'·'}</span>${t.n}</div>`).join('') || '<span style="color:var(--muted)">아직 실행 기록 없음</span>';
  renderRecalc(s.recalc);
  if(s.fork && s.fork.length){ $('forkcard').style.display='block';
    $('forkmsg').textContent = `구버전 ${s.fork.join(', ')} 이 최신본보다 나중에 수정되었습니다. 잘못된 파일에 입력 중일 수 있으니 확인하세요.`; }
  else $('forkcard').style.display='none';
  if(s.tunnel){ $('tunnelrow').style.display='block';
    $('tunnelurl').textContent = s.tunnel; $('tunnelurl').href = s.tunnel; }
  else $('tunnelrow').style.display='none';
}

/* ── 정산·업무 ── */
let works = {as:[], pm:[]}, mode = 'settle', issuesData = {rows:[],cols:[]}, checks = {};
function rowIs2026(r, kind){
  if(!r || typeof r!=='object') return false;
  const dateKeys = {
    settle:['완료일'], as:['접수일자'], pm:['점검예정일'], erp:['월','전표'],
    issue:['기준일','접수일자','점검예정일','완료일','발행일','일자']
  }[kind] || [];
  for(const k of dateKeys){
    const s=String(r[k]||''), m=s.match(/(^|[^\d])(20\d{2})[-./]/);
    if(m) return m[2]===APP_YEAR;
    const q=s.match(/(^|[^\d])(\d{2})\/\d{2}\/\d{2}\s*-\s*\d+/);
    if(q) return q[2]===APP_YEAR.slice(2);
  }
  const ids = [r.업무ID,r.접수ID,r.점검ID,r.정산ID,r.원천업무ID,r.ID].map(x=>String(x||'')).join(' ');
  const im = ids.match(/\b(?:AS|PM|JS)-(\d{2})\d{2}(?:-|$)/i);
  if(im) return im[1]===APP_YEAR.slice(2);
  const ps = [r.프로젝트NO,r.포함프로젝트,r.프로젝트명].map(x=>String(x||'')).join(' ');
  const ys = [...ps.matchAll(/\bUJ(\d{2})\d{5}\b/gi)].map(m=>m[1]);
  return ys.length>0 && new Set(ys).size===1 && ys[0]===APP_YEAR.slice(2);
}
/* ★ 사용자 지시(2026-07-29): "철거·신규납품 건은 DB만 저장해놓고 앱에 표기하지마."
   원장에서 지우지 않는다 — 05_신규납품설치 시트에 그대로 있고, **화면에서만** 뺀다.
   05시트는 앱이 읽지 않으므로 보통은 저절로 안 보이지만, 정산(06)·확인필요(23) 로
   흘러들어오면 보인다. 그래서 데이터를 받는 관문에서 한 번 더 거른다.
   ★ 사용자 지시(2026-07-29 추가): "추후에 앱에 추가할 수도 있으니 감안해서 정리해줘."
   그래서 **스위치 한 개**로 모았다. 나중에 보여 달라고 하면 아래 한 줄만 true 로 바꾸면
   화면 전체(정산·업무·이슈·확인필요·일정)가 한 번에 다시 나온다. 자료는 계속 쌓이고 있으므로
   그 즉시 과거분까지 전부 보인다 — 다시 모으거나 되돌릴 작업이 없다.
   숨기는 곳을 여기저기 흩어 놓으면 켤 때 한두 군데를 반드시 빠뜨린다. */
const SHOW_SIDE_WORK = false;   // ← 철거·신규납품을 앱에 보이려면 true 한 글자만 바꾼다
/* 05시트 업무구분은 '납품'·'설치'·'철거'·'이전' 처럼 **한 단어**다(10_코드관리 M열).
   '신규납품' 만 잡으면 정작 원장에 적힌 '납품' 을 놓친다 — 실제로 처음에 그랬다. */
const SIDE_WORK = /철거|이전|납품|설치|계단|안전바|경보장치|메자닌/;
function isSideWork(r){
  if(SHOW_SIDE_WORK) return false;          // 스위치가 켜지면 아무것도 숨기지 않는다
  if(!r || typeof r!=='object') return false;
  for(const k of ['업무구분','업무유형','구분','종류','품목']){
    if(r[k] && SIDE_WORK.test(String(r[k]))) return true;
  }
  return false;
}
function hide2025(r){
  const old=/\bUJ25\d{5}\b|\b(?:AS|PM|JS)-25\d{2}-\d{3}\b|(?<!\d)2025[-./]\d{1,2}(?:[-./]\d{1,2})?|2025년|(?<!\d)25(?:년도|년)/g;
  return Object.fromEntries(Object.entries(r||{}).map(([k,v])=>
    [k,typeof v==='string'?v.replace(old,'').trim():v]));
}
function cleanExec2026(d){
  const old = x => /\b(?:AS|PM|JS)-25\d{2}-\d{3}\b|\bUJ25\d{5}\b|(^|[^\d])2025[-./]|2025년|(^|[^\d])25(?:년도|년)/.test(JSON.stringify(x||''));
  if(old((d.meta||{})['보고일']) || old((d.meta||{})['집계기준일'])) return {};
  d.summary = (d.summary||[]).filter(x=>!old(x));
  d.sections = (d.sections||[]).map(s=>({...s,
    items:(s.items||[]).filter(x=>!old(x)),
    lines:(s.lines||[]).filter(x=>!old(x)),
    groups:(s.groups||[]).map(g=>({...g,items:(g.items||[]).filter(x=>!old(x))}))
  })).filter(s=>!old(s.title));
  d.details = Object.fromEntries(Object.entries(d.details||{}).map(([k,v])=>[
    k,{...v,rows:(v.rows||[]).filter(r=>rowIs2026(r,'issue') && !isSideWork(r)).map(hide2025)}
  ]).filter(([,v])=>!old(v)));
  return d;
}
function cleanErp2026(d){
  const rows=(d.rows||[]).filter(r=>rowIs2026(r,'erp')).map(hide2025), months={}, kinds={};
  let total=0;
  rows.forEach(r=>{
    const mo=String(r.월||''), sup=+r.공급가액||0, kind=String(r.유형||'');
    const m=months[mo] ||= {합계:0,건수:0}; m.합계+=sup; m.건수++; m[kind]=(m[kind]||0)+sup;
    kinds[kind]=(kinds[kind]||0)+sup; total+=sup;
  });
  return {...d,rows,months,kinds,total};
}
/* ★ 이 주소는 **어떤 일이 있어도 바뀌지 않는다**(2026-07-27 확정).
   경유하는 터널 주소는 띄울 때마다 새로 받지만, 폰이 아는 주소는 늘 이것 하나다.
   배너로 띄워 닫게 하지 않고, 대시보드에 **항상 보이게** 박아 둔다. */
const FIXED_ENTRY = "https://mulder4780.github.io/coupang-ecount-reconcile/";
function showFixedEntry(){
  const a = $('fixedurl'); if(!a) return;
  a.textContent = FIXED_ENTRY.replace(/^https:\/\//,'').replace(/\/$/,'');
  a.href = FIXED_ENTRY;
  const b = $('copyfixed');
  b.onclick = async () => {
    try{ await navigator.clipboard.writeText(FIXED_ENTRY); }
    catch(e){ /* http 출처에서는 clipboard API가 막힌다 — 옛 방식으로 되돌린다 */
      const t = document.createElement('textarea');
      t.value = FIXED_ENTRY; document.body.appendChild(t); t.select();
      try{ document.execCommand('copy'); }catch(_){}
      t.remove();
    }
    b.textContent = '복사됨'; setTimeout(()=>{ b.textContent = '복사'; }, 1600);
  };
}
async function loadSettle(){
  const d = await api('/api/settlements');
  settleRows = (d.rows||[]).filter(r=>rowIs2026(r,'settle') && !isSideWork(r)).map(hide2025);
  try{ works = await api('/api/works'); }catch(e){}
  try{ issuesData = await api('/api/issues'); }catch(e){}
  try{ checks = await api('/api/checks'); }catch(e){}
  try{ execRep = await api('/api/exec_report'); }catch(e){}
  try{ repData = await api('/api/v1/reports/daily/exceptions'); }catch(e){ repData={}; }
  try{ erpDocs = await api('/api/erpdocs'); }catch(e){ erpDocs = {rows:[],months:{},kinds:{},total:0}; }
  works.as = (works.as||[]).filter(r=>rowIs2026(r,'as') && !isSideWork(r)).map(hide2025);
  works.pm = (works.pm||[]).filter(r=>rowIs2026(r,'pm') && !isSideWork(r)).map(hide2025);
  issuesData.rows = (issuesData.rows||[]).filter(r=>rowIs2026(r,'issue') && !isSideWork(r)).map(hide2025);
  checks = Object.fromEntries(Object.entries(checks||{}).filter(([k])=>rowIs2026({ID:k},'issue')));
  execRep = cleanExec2026(execRep||{});
  erpDocs = cleanErp2026(erpDocs||{});
  try{ brandLogo = (await api('/api/brand')).logo || ''; }catch(e){ brandLogo=''; }
  applyBrand();
  syncSortUI(); fillStatOptions(); renderSettle(); periodInit();
  try{ renderBoard(); }catch(e){ console.warn('board', e); }
  try{ renderRepresentative(); }catch(e){ console.warn('representative report', e); }
  try{ renderCheckHub(); }catch(e){ console.warn('check hub', e); }
  try{ $('helpcard').innerHTML = helpAll(); }catch(e){ console.warn('help', e); }
  // API 응답이 큰 경우 [보고]를 먼저 연 사용자가 0/0을 보지 않도록, 원천 업무가
  // 모두 도착한 뒤 현재 열린 일일보고를 한 번 더 그린다.
  if(window.__view==='daily') try{ renderDaily(); }catch(e){ console.warn('daily refresh', e); }
}
function setMode(m){
  mode = m;
  document.querySelectorAll('#seg button').forEach(b=>b.classList.toggle('on', b.dataset.m===m));
  resetSort();                       // 탭을 바꾸면 항상 기본 순서(과거→최근)로 복귀
  fillStatOptions(); renderSettle();
}
function kTap(el){ el.classList.remove('tapped'); void el.offsetWidth; el.classList.add('tapped'); }
document.addEventListener('pointerdown',e=>{
  const el=e.target.closest('button,.topitem,.srow,.check-kpi,.kpi,.rep-metric,.ecell.metric,.bneed.actionable,.bpill.actionable');
  if(!el||el.disabled) return;
  el.classList.remove('press-pop'); void el.offsetWidth; el.classList.add('press-pop');
  el.addEventListener('animationend',()=>el.classList.remove('press-pop'),{once:true});
},{passive:true});
/* 종류·ID를 알고 있는 목록은 그 힌트를 버리지 않는다.
   같은 프로젝트NO가 AS·점검·정산에 함께 있는 실제 데이터가 많아서 정산부터 찾으면 오탐이 난다. */
function openRecord(kind, id, project){
  const list = kind==='settle' ? settleRows : kind==='as' ? (works.as||[]) : kind==='pm' ? (works.pm||[]) : [];
  const rid = String(id||'').trim(), prj = String(project||'').trim();
  let idx = list.findIndex(r=>[r.정산ID,r.접수ID,r.점검ID].some(x=>String(x||'')===rid));
  if(idx<0 && prj) idx = list.findIndex(r=>String(r.프로젝트NO||'')===prj);
  if(idx<0){ openByPrj(rid||prj); return; }
  show('settle'); setMode(kind); window._lastRows = list;
  const r = list[idx];
  openSheet(kind, idx, r.정산ID||r.접수ID||r.점검ID||prj);
}

/* 프로젝트NO(또는 ID)로 상세 열기.
   ID 정확일치는 바로 열고, 프로젝트NO가 여러 화면에 있으면 사용자가 고르게 한다. */
function openByPrj(key){
  const k = String(key||'').trim();
  const all = [
    ...settleRows.map((r,i)=>({kind:'settle',i,r,id:r.정산ID||''})),
    ...(works.as||[]).map((r,i)=>({kind:'as',i,r,id:r.접수ID||''})),
    ...(works.pm||[]).map((r,i)=>({kind:'pm',i,r,id:r.점검ID||''}))
  ];
  const exact = all.filter(x=>String(x.id)===k);
  const found = exact.length ? exact : all.filter(x=>projectNoOf(x.r)===k);
  if(found.length===1){ const x=found[0]; openRecord(x.kind,x.id,projectNoOf(x.r)); return; }
  if(found.length>1){
    const kindName={settle:'정산·청구',as:'돌발AS',pm:'정기점검'};
    const heading=exact.length ? projectLabel(found[0].r) : k;
    showSheet(`<h2>${esc2(heading)} <span class="chip c-warn">${found.length}건</span></h2>
      <div class="sub">같은 프로젝트 번호에 연결된 기록이 여러 개입니다. 확인할 기록을 선택하세요.</div>
      <div class="slist">${found.map(x=>{
        const d=x.r.완료일||x.r.접수일자||x.r.점검예정일||x.r.작업완료일||'';
        const prj=projectLabel(x.r);
        return `<div class="srow" onclick="openRecord(${esc4(x.kind)},${esc4(x.id)},${esc4(projectNoOf(x.r))})">
          <div class="top"><span class="prjno">${esc2(prj)}</span><span class="chip">${kindName[x.kind]}</span></div>
          <div class="top"><span class="camp">${esc2(x.r.캠프명||'캠프 미상')}</span><span class="camp">${esc2(d||'-')}</span></div>
          ${x.id?`<div class="meta"><span class="sid">${esc2(x.id)}</span></div>`:''}
        </div>`;}).join('')}</div>`);
    return;
  }
  alert(`${k} 를 찾지 못했습니다 (다른 월 데이터이거나 아직 미등록)`);
}
function goList(mode, filter){
  setTimeout(()=>{
    show('settle'); setMode(mode);
    if(filter){ $('fstat').value = filter; renderSettle(); }
  }, 140);   // 탭 애니메이션 보이고 전환
}
function flashTop(){
  const card = $('toplist').closest('.card');
  card.scrollIntoView({behavior:'smooth', block:'center'});
  card.classList.remove('flash'); void card.offsetWidth; card.classList.add('flash');
}
function jumpIssue(t){
  show('settle'); mode='settle';
  document.querySelectorAll('#seg button').forEach(b=>b.classList.toggle('on', b.dataset.m==='settle'));
  fillStatOptions(); $('fstat').value = t; renderSettle();
}
function curRows(){
  if(mode==='as') return works.as||[];
  if(mode==='pm') return works.pm||[];
  if(mode==='check') return issuesData.rows||[];
  return settleRows;
}
function statOf(r){ return mode==='as' ? r.진행상태 : mode==='pm' ? r.점검상태 :
  mode==='check' ? (r.문제유형||r.검증결과||r.우선순위||'확인필요') : r.상태; }
/* 프로젝트NO 바로 옆에 붙는 캠프 이름.
   어느 현장인지가 번호만큼 중요해서 카드 첫 줄로 올렸다(2026-07-26 요청).
   캠프를 모르는 건은 빈칸으로 두지 않고 '캠프 미상'이라고 적어 조치 대상임을 보이게 한다. */
/* 관리자 확인이 끝난 건에 붙는 체크 표시 — 무엇이 남았는지 한눈에 보이게 한다. */
function okTag(state, when){
  const st = String(state||'').trim();
  if(!st) return '';
  const good = st==='일치';
  return `<span class="oktag ${good?'':'warn'}">${good?'✓ 확인완료':st}${when?` ${String(when).slice(5)}`:''}</span>`;
}
/* 엑셀 검증결과(수식)를 그대로 띄운다 — 앱이 따로 계산하지 않고 대장을 그대로 보여주는 것이
   숫자가 어긋나지 않는 유일한 방법이다. 문제코드는 눌러서 설명을 볼 수 있게 넘긴다. */
function verTag(res, codes){
  const r = String(res||'').trim();
  if(!r) return '';
  if(r==='정상') return `<span class="oktag">✓ 검증 정상</span>`;
  const n = String(codes||'').split(/[,·]/).map(x=>x.trim()).filter(Boolean).length;
  return `<span class="oktag ${r==='누락·지연'?'bad':'warn'}">${r}${n?` ${n}`:''}</span>`;
}
function campTag(v){
  const s = String(v||'').trim();
  return s ? `<span class="camptag">${s}</span>`
           : `<span class="camptag none">캠프 미상</span>`;
}
function vbadge(id){
  const c = checks[id]; if(!c) return '';
  const one = (label,v)=>{ if(!v) return '';
    const good = v.startsWith('확인')||v==='일치';
    return `<span class="chip ${good?'c-ok':'c-danger'}" style="font-size:10px">${label} ${good?'✓':'✗'}</span>`; };
  return one('카톡',c.kakao)+one('밴드',c.band)+(c.erp?`<span class="chip c-warn" style="font-size:10px">ERP ${c.erp.slice(0,14)}</span>`:'')+(c.po?`<span class="chip c-warn" style="font-size:10px">PO ${c.po.slice(0,14)}</span>`:'');
}
/* ── 기간 필터 ────────────────────────────────────────────────
   탭마다 기준 날짜가 다르므로(정산=완료일, AS=접수일, 점검=예정일)
   정렬과 같은 dateOf()를 그대로 쓴다 — 화면에 보이는 날짜와 필터가 어긋나지 않게. */
function periodPick(){
  const v = $('fperiod').value;
  $('fdates').classList.toggle('on', v === 'custom');
  if(v === 'custom' && !$('fd1').value){
    const ds = curRows().map(r=>dateOf(r, mode)).filter(Boolean).sort();
    if(ds.length){ $('fd1').value = ds[0]; $('fd2').value = ds[ds.length-1]; }
  }
  renderSettle();
}
function workPeriodRange(){
  const v = ($('fperiod')||{}).value || '';
  if(!v) return null;
  if(v === 'custom') return [$('fd1').value || '0000-01-01', $('fd2').value || '9999-12-31'];
  const t = new Date();
  const ym = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  if(v === 'm0') return [ym(t)+'-01', ym(t)+'-31'];
  if(v === 'm1'){ const d = new Date(t.getFullYear(), t.getMonth()-1, 1); return [ym(d)+'-01', ym(d)+'-31']; }
  if(v === 'm3'){ const d = new Date(t.getFullYear(), t.getMonth()-2, 1); return [ym(d)+'-01', ym(t)+'-31']; }
  if(v === 'y0') return [t.getFullYear()+'-01-01', t.getFullYear()+'-12-31'];
  return null;
}
function showPeriodNote(pr, shown, undated){
  const el = $('fnote'); if(!el) return;
  if(!pr){ el.innerHTML = ''; return; }
  const label = {m0:'이번 달', m1:'지난 달', m3:'최근 3개월', y0:'올해',
                 custom:'직접 지정'}[$('fperiod').value] || '';
  // 날짜 없는 건이 조용히 사라지지 않게 몇 건이 빠졌는지 반드시 알린다
  el.innerHTML = `기간 <b>${label} (${pr[0]} ~ ${pr[1]})</b> · ${shown}건 표시` +
    (undated ? ` · <b>날짜 없는 ${undated}건은 제외</b>` : '');
}
function fillStatOptions(){
  const stats = [...new Set(curRows().map(statOf).filter(Boolean))];
  $('fstat').innerHTML = '<option value="">상태 전체</option>' + stats.map(s=>`<option>${s}</option>`).join('');
}
/* 조치가 필요한 건인가 — ERP 계산서(묶음)는 이미 발행된 서류라 조치 대상이 아니다 */
function needAction(r){
  return !['정상','무상/보험','ERP 계산서(묶음)'].includes(r.상태);
}
function chip(st){
  const ok=['정상','작업완료','완료','없음'], warn=['입금 대기','방문예정','작업중','예정','접수'];
  const skip=['무상/보험','ERP 계산서(묶음)'];
  const cls = ok.includes(st)?'c-ok': skip.includes(st)?'c-skip': warn.includes(st)?'c-warn':'c-danger';
  return `<span class="chip ${cls}">${st||'-'}</span>`;
}
/* ── 정렬 ──────────────────────────────────────────────────────
   기본은 언제나 "과거 → 최근": 오래된 날짜가 맨 위, 최근 날짜가 맨 아래.
   서버도 같은 규칙으로 내려주므로 새로 추가되는 건도 자동으로 이 순서를 따른다.
   기준(날짜/ID/금액/상태)과 방향은 툴바에서 바꿀 수 있고, 표 헤더 클릭으로도 된다. */
const SORTKEYS = ['__date','__id','__amt','__stat'];
window._sk = '__date';   // 정렬 기준 (특수키 또는 실제 열 이름)
window._sd = 1;          // 1 = 과거→최근(오름차순), -1 = 최근→과거

function normDate(v){
  const m = String(v==null?'':v).match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}` : '';
}
function anyDate(r){                       // 날짜 열 이름을 모르는 시트(확인필요)용
  for(const v of Object.values(r||{})){ const d = normDate(v); if(d) return d; }
  return '';
}
function recordIdOf(r){
  return String((r&&(
    r.정산ID||r.접수ID||r.점검ID||r.레코드ID||r.원천업무ID||r.업무ID||r.ID
  ))||'').trim();
}
function isInternalNo(v){ return /^(?:AS|PM|JS)-/i.test(String(v||'').trim()); }
function isRepresentativeProject(v){
  const s=String(v||'').trim();
  return /^UJ26\d{4,}$/i.test(s)||/^ERP(?:[-_\s]|$)/i.test(s);
}
/* 화면의 대표 식별자는 언제나 프로젝트 번호다.
   내부 AS/PM/JS 번호만 들어온 예외·KPI 행은 원장 3종을 역참조하고, 그래도 못 찾으면
   내부번호를 프로젝트처럼 가장하지 않고 '프로젝트 미확정'으로 명시한다. */
function projectNoOf(r){
  if(!r||typeof r!=='object') return '';
  const direct=String(r.프로젝트NO||'').trim();
  if(isRepresentativeProject(direct)) return direct;
  for(const v of Object.values(r)){
    const hit=String(v==null?'':v).match(/\bUJ26\d{4,}\b/i);
    if(hit) return hit[0].toUpperCase();
  }
  const ids=[r.정산ID,r.접수ID,r.점검ID,r.레코드ID,r.원천업무ID,r.업무ID,r.ID]
    .map(v=>String(v||'').trim()).filter(Boolean);
  if(!ids.length) return '';
  const all=[...(settleRows||[]),...((works&&works.as)||[]),...((works&&works.pm)||[])];
  const hit=all.find(x=>{
    const xid=[x.정산ID,x.접수ID,x.점검ID,x.원천업무ID,x.업무ID,x.ID]
      .map(v=>String(v||'').trim()).filter(Boolean);
    return ids.some(id=>xid.includes(id));
  });
  const linked=String((hit&&hit.프로젝트NO)||'').trim();
  return isRepresentativeProject(linked)?linked:'';
}
function projectLabel(r){ return projectNoOf(r)||'프로젝트 미확정'; }
function idOf(r){ return projectNoOf(r)||recordIdOf(r); }
function sortVal(r, k){
  if(k==='__date') return (mode==='check' ? anyDate(r) : normDate(dateOf(r, mode))) || anyDate(r);
  if(k==='__id')   return idOf(r);
  if(k==='__stat') return statOf(r)||'';
  if(k==='__amt')  return +r.공급가액||0;
  return r[k];
}
function setSort(k){                       // 표 헤더 클릭
  if(window._sk===k) window._sd*=-1; else { window._sk=k; window._sd=1; }
  syncSortUI(); renderSettle();
}
function sortPick(){ window._sk = $('fsort').value; window._sd = 1; syncSortUI(); renderSettle(); }
function toggleDir(){ window._sd*=-1; syncSortUI(); renderSettle(); }
function resetSort(){ window._sk='__date'; window._sd=1; syncSortUI(); }
function syncSortUI(){
  const sel = $('fsort'); if(!sel) return;
  const std = SORTKEYS.includes(window._sk);
  sel.querySelector('option[value="__col"]').hidden = std;
  sel.value = std ? window._sk : '__col';
  const asc = window._sd>0;
  $('fdir').textContent = asc ? '과거 → 최근' : '최근 → 과거';
  $('fdir').title = asc ? '오래된 항목이 맨 위, 최근 항목이 맨 아래(기본)' : '최근 항목이 맨 위';
  $('fdir').classList.toggle('desc', !asc);
}
function thx(label, key, num){
  const on = window._sk===key;
  return `<th data-sk="1" ${num?'class="num"':''} onclick="setSort('${key}')">${label}${on?(window._sd>0?' ▲':' ▼'):''}</th>`;
}
function sortRows(rows){
  const k = window._sk || '__date';
  const blank = window._sd>0 ? 1 : -1;      // 값이 빈 행은 방향과 무관하게 항상 맨 뒤
  return [...rows].sort((a,b)=>{
    const av = sortVal(a,k), bv = sortVal(b,k);
    const ae = av===''||av==null, be = bv===''||bv==null;
    if(ae&&be) return String(idOf(a)).localeCompare(String(idOf(b)),'ko');
    if(ae) return blank;
    if(be) return -blank;
    let r;
    if(typeof av==='number' && typeof bv==='number') r = av-bv;
    else if(normDate(av) && normDate(bv)) r = normDate(av).localeCompare(normDate(bv));
    else {
      const pure = s => /^[\d,.\s-]+$/.test(String(s));
      r = (pure(av)&&pure(bv))
        ? parseFloat(String(av).replace(/[^0-9.-]/g,'')) - parseFloat(String(bv).replace(/[^0-9.-]/g,''))
        : String(av).localeCompare(String(bv),'ko');
    }
    if(!r) r = String(idOf(a)).localeCompare(String(idOf(b)),'ko');
    return r*window._sd;
  });
}
function openListKpi(label, rule){
  const all=window._visibleListRows||[];
  let rows=all;
  if(rule==='unpaid') rows=all.filter(r=>(+r.미수금||0)>0);
  else if(rule==='action') rows=all.filter(r=>needAction(r));
  else if(rule==='as-done') rows=all.filter(r=>r.진행상태==='작업완료');
  else if(rule==='as-open') rows=all.filter(r=>r.진행상태!=='작업완료');
  else if(rule==='pm-done') rows=all.filter(r=>r.점검상태==='완료');
  else if(rule==='pm-open') rows=all.filter(r=>r.점검상태!=='완료');
  window._briefMetric={label,data:{rows:rows.map(r=>({...r,종류:mode,레코드ID:recordIdOf(r),
    프로젝트NO:projectNoOf(r),일자:dateOf(r,mode),담당자:r.담당자||r.담당기사||'',
    문제:rule==='unpaid'?`미수금 ${fmt(+r.미수금||0)}원`:
      rule==='action'?String(r.상태||r.검증결과||'조치 필요'):'',
    상태:statOf(r),금액:+r.공급가액||0})),
    count:rows.length,kind:'list-filter',
    basis:`현재 ${mode==='settle'?'정산·청구':mode==='as'?'돌발AS':mode==='pm'?'정기점검':'확인필요'} 화면의 검색·상태·기간 필터 결과`}};
  openExecMetric(label);
}
function renderSettle(){
  const q = ($('q').value||'').toLowerCase(), f = $('fstat').value;
  const pr = workPeriodRange();
  let undated = 0;
  const base = curRows().filter(r => {
    if(f && statOf(r)!==f) return false;
    if(q && !Object.values(r).join('').toLowerCase().includes(q)) return false;
    if(pr){
      const d = dateOf(r, mode);
      if(!d){ undated++; return false; }     // 날짜 없는 건은 기간으로 판단 불가 → 제외하고 아래에 알림
      if(d < pr[0] || d > pr[1]) return false;
    }
    return true;
  });
  const rows = sortRows(base);
  window._visibleListRows=rows;
  showPeriodNote(pr, rows.length, undated);
  if(mode==='settle'){
    const need = rows.filter(r=>needAction(r));
    const 미수 = rows.reduce((a,r)=>a+(+r.미수금||0),0);
    $('skpis').innerHTML = `
      <div class="kpi accent" role="button" tabindex="0" onclick="openListKpi('공급가액 목록','all')"><div class="v">${fmt(rows.reduce((a,r)=>a+(+r.공급가액||0),0))}</div><div class="l">공급가액(원)</div></div>
      <div class="kpi ${미수?'warn':''}" role="button" tabindex="0" onclick="openListKpi('미수금 목록','unpaid')"><div class="v">${fmt(미수)}</div><div class="l">미수금(원)</div></div>
      <div class="kpi ${need.length?'danger':''}" role="button" tabindex="0" onclick="openListKpi('조치 필요 목록','action')"><div class="v">${need.length}/${rows.length}</div><div class="l">조치 필요</div></div>`;
    $('slist').innerHTML = rows.map((r,i)=>`
      <div class="srow" onclick="openSheet('settle',${i},'${r.정산ID}')">
        <div class="top"><span class="prjno">${esc2(projectLabel(r))}</span>${campTag(r.캠프명)}${chip(r.상태)}</div>
        <div class="top"><span class="camp">${r.업무구분||''}</span><span class="amt">${fmt(r.공급가액)}원</span></div>
        <div class="meta"><span class="sid">${r.정산ID}</span><span>완료 ${r.완료일||'-'}</span>
          <span>명세서 ${r.명세서번호||r.명세서}</span><span>계산서 ${r.계산서}</span>
          ${r.입금일?`<span>입금 ${r.입금일}</span>`:''} ${vbadge(r.정산ID)}
          ${r.적요?`<span style="flex:1 1 100%;color:var(--ink-3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.적요}</span>`:''}</div>
      </div>`).join('') || '<div class="card">조건에 맞는 건이 없습니다</div>';
    $('sgrid').innerHTML = `<tr>${thx('프로젝트NO','프로젝트NO')}${thx('정산ID','정산ID')}${thx('구분','업무구분')}${thx('캠프','캠프명')}
      ${thx('공급가액','공급가액',1)}${thx('명세서번호','명세서번호')}${thx('계산서','계산서발행일')}${thx('PO','PO번호')}${thx('입금일','입금일')}${thx('미수금','미수금',1)}${thx('완료일','완료일')}${thx('상태','상태')}</tr>` +
      rows.map((r,i)=>`<tr onclick="openSheet('settle',${i},'${r.정산ID}')" style="cursor:pointer">
        <td><b class="prjno">${esc2(projectLabel(r))}</b></td><td>${r.정산ID}</td><td>${r.업무구분}</td>
        <td>${r.캠프명}</td><td class="num">${fmt(r.공급가액)}</td><td>${r.명세서번호||'-'}</td>
        <td>${r.계산서}${r.계산서발행일?' '+r.계산서발행일:''}</td><td>${r.PO번호||'-'}</td><td>${r.입금일||'-'}</td>
        <td class="num">${r.미수금===''?'-':fmt(r.미수금)}</td><td>${r.완료일||'-'}</td><td>${chip(r.상태)}</td></tr>`).join('') +
      `<tr class="tot"><td colspan="4">합계 ${rows.length}건</td><td class="num">${fmt(rows.reduce((a,r)=>a+(+r.공급가액||0),0))}</td>
       <td colspan="4"></td><td class="num">${fmt(rows.reduce((a,r)=>a+(+r.미수금||0),0))}</td><td colspan="2"></td></tr>`;
  } else if(mode==='as'){
    const done = rows.filter(r=>r.진행상태==='작업완료').length;
    $('skpis').innerHTML = `
      <div class="kpi accent" role="button" tabindex="0" onclick="openListKpi('돌발AS 접수 목록','all')"><div class="v">${rows.length}</div><div class="l">접수 건수</div></div>
      <div class="kpi" role="button" tabindex="0" onclick="openListKpi('돌발AS 작업완료','as-done')"><div class="v">${done}</div><div class="l">작업완료</div></div>
      <div class="kpi ${rows.length-done?'warn':''}" role="button" tabindex="0" onclick="openListKpi('돌발AS 진행중','as-open')"><div class="v">${rows.length-done}</div><div class="l">진행중</div></div>`;
    $('slist').innerHTML = rows.map((r,i)=>`
      <div class="srow" onclick="openSheet('as',${i},'${r.접수ID}')">
        <div class="top"><span class="prjno">${esc2(projectLabel(r))}</span>${campTag(r.캠프명)}${chip(r.진행상태)}</div>
        <div class="top"><span class="camp">${r.담당기사||'-'}</span></div>
        <div class="meta"><span class="sid">${r.접수ID||''}</span><span>접수 ${r.접수일자||'-'}</span>${r.작업완료일?`<span>완료 ${r.작업완료일}</span>`:''}<span>${r['유상·무상·보험']||''}</span>${okTag(r.관리자검증상태, r.최종확인일)}${verTag(r.검증결과, r.검증문제코드)}
          ${r.완료보고서등록?`<span>완료보고서 ${r.완료보고서등록}</span>`:''}${r.ERP등록?`<span>ERP ${r.ERP등록}</span>`:''}
          ${r.거래명세서반영?`<span>명세서 ${r.거래명세서반영}</span>`:''}${r.ERP반영?`<span>현장ERP ${r.ERP반영}</span>`:''}
          ${r.검증자?`<span>검증 ${r.검증자}${r.검증일?' · '+r.검증일:''}</span>`:''}
          ${r.긴급도==='긴급'?'<span style="color:var(--danger);font-weight:700">긴급</span>':''}
          <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.신청내용||''}</span></div>
      </div>`).join('') || '<div class="card">조건에 맞는 건이 없습니다</div>';
    $('sgrid').innerHTML = `<tr>${thx('프로젝트NO','프로젝트NO')}${thx('접수ID','접수ID')}${thx('캠프','캠프명')}${thx('접수일','접수일자')}${thx('기사','담당기사')}${thx('유·무상','유상·무상·보험')}<th>내용</th>${thx('완료일','작업완료일')}<th>완료보고서</th><th>명세서반영</th><th>ERP반영</th><th>검증자·일</th>${thx('상태','진행상태')}</tr>` +
      rows.map((r,i)=>`<tr onclick="openSheet('as',${i},'${r.접수ID}')" style="cursor:pointer">
        <td><b class="prjno">${esc2(projectLabel(r))}</b></td><td>${r.접수ID||'-'}</td><td>${r.캠프명}</td><td>${r.접수일자||'-'}</td><td>${r.담당기사||'-'}</td>
        <td>${r['유상·무상·보험']||'-'}</td><td style="max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.신청내용||''}</td>
        <td>${r.작업완료일||'-'}</td><td>${r.완료보고서등록||'-'}</td><td>${r.거래명세서반영||'-'}</td><td>${r.ERP반영||r.ERP등록||'-'}</td><td>${r.검증자?(r.검증자+(r.검증일?' · '+r.검증일:'')):'-'}</td><td>${chip(r.진행상태)}</td></tr>`).join('');
  } else if(mode==='check'){
    // 엑셀 07_불일치누락현황 — 검증 안 된·확인해야 할 항목 그대로
    const types = {};
    rows.forEach(r=>{ const t=statOf(r); types[t]=(types[t]||0)+1; });
    $('skpis').innerHTML = `
      <div class="kpi danger" role="button" tabindex="0" onclick="openListKpi('확인필요 전체','all')"><div class="v">${rows.length}</div><div class="l">확인필요 건수</div></div>
      <div class="kpi" role="button" tabindex="0" onclick="openCheckFiltered()"><div class="v">${Object.keys(types).length}</div><div class="l">문제 유형 수</div></div>
      <div class="kpi" role="button" tabindex="0" onclick="openCheckFiltered()"><div class="v">${(issuesData.cols||[]).length}</div><div class="l">검증 항목(열)</div></div>`;
    const keyOf = r => projectNoOf(r)||recordIdOf(r)||Object.values(r)[0]||'';
    const descOf = r => r.문제내용||r.경고내용||r.검증문제코드||r.누락서류||r['내용·근거']||'';
    $('slist').innerHTML = rows.map((r,i)=>`
      <div class="srow" onclick="openByPrj('${keyOf(r)}')">
        <div class="top"><span class="prjno">${esc2(projectLabel(r))}</span>${campTag(r.캠프명)}${chip(statOf(r))}</div>
        <div class="top"><span class="camp">${r.담당자||r.담당기사||''}</span></div>
        <div class="meta"><span class="sid">${keyOf(r)}</span><span style="flex:1">${descOf(r)}</span></div>
      </div>`).join('') || '<div class="card">확인필요 항목이 없습니다 🎉</div>';
    const cols = (issuesData.cols||[]).filter(c=>rows.some(r=>r[c])).slice(0,9);
    $('sgrid').innerHTML = `<tr>${cols.map(c=>`<th>${c}</th>`).join('')}</tr>` +
      rows.map((r,i)=>`<tr onclick="openByPrj('${keyOf(r)}')" style="cursor:pointer">${
        cols.map(c=>`<td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r[c]||'-'}</td>`).join('')}</tr>`).join('');
  } else {
    const done = rows.filter(r=>r.점검상태==='완료').length;
    $('skpis').innerHTML = `
      <div class="kpi accent" role="button" tabindex="0" onclick="openListKpi('정기점검 전체','all')"><div class="v">${rows.length}</div><div class="l">점검 건수</div></div>
      <div class="kpi" role="button" tabindex="0" onclick="openListKpi('정기점검 완료','pm-done')"><div class="v">${done}</div><div class="l">완료</div></div>
      <div class="kpi ${rows.length-done?'warn':''}" role="button" tabindex="0" onclick="openListKpi('정기점검 예정·미점검','pm-open')"><div class="v">${rows.length-done}</div><div class="l">예정·미점검</div></div>`;
    $('slist').innerHTML = rows.map((r,i)=>`
      <div class="srow" onclick="openSheet('pm',${i},'${r.점검ID}')">
        <div class="top"><span class="prjno">${esc2(projectLabel(r))}</span>${campTag(r.캠프명)}${chip(r.점검상태)}</div>
        <div class="top"><span class="camp">${r.담당기사||'-'}</span></div>
        <div class="meta"><span class="sid">${r.점검ID||''}</span><span>예정 ${r.점검예정일||'-'}</span><span>실제 ${r.실제점검일||'-'}</span>${okTag(r['최종확인일(유현민 체크)']?'일치':'', r['최종확인일(유현민 체크)'])}${verTag(r.검증결과, r.검증문제코드)}
          ${r.ERP판매전표?`<span>판매전표 ${r.ERP판매전표}</span>`:''}${r.거래명세서?`<span>거래명세서 ${r.거래명세서}</span>`:''}
          ${r.검증자?`<span>검증 ${r.검증자}${r.검증일?' · '+r.검증일:''}</span>`:''}
          ${r.이상발견여부==='있음'?'<span style="color:var(--danger);font-weight:700">이상발견</span>':''}</div>
      </div>`).join('') || '<div class="card">조건에 맞는 건이 없습니다</div>';
    $('sgrid').innerHTML = `<tr>${thx('프로젝트NO','프로젝트NO')}${thx('점검ID','점검ID')}${thx('캠프','캠프명')}${thx('예정일','점검예정일')}${thx('실제점검일','실제점검일')}${thx('기사','담당기사')}${thx('이상','이상발견여부')}${thx('AS전환','돌발AS전환여부')}<th>판매전표</th><th>거래명세서</th>${thx('상태','점검상태')}</tr>` +
      rows.map((r,i)=>`<tr onclick="openSheet('pm',${i},'${r.점검ID}')" style="cursor:pointer">
        <td><b class="prjno">${esc2(projectLabel(r))}</b></td><td>${r.점검ID||'-'}</td><td>${r.캠프명}</td><td>${r.점검예정일||'-'}</td><td>${r.실제점검일||'-'}</td>
        <td>${r.담당기사||'-'}</td><td>${r.이상발견여부||'-'}</td><td>${r.돌발AS전환여부||'-'}</td><td>${r.ERP판매전표||'-'}</td><td>${r.거래명세서||'-'}</td><td>${chip(r.점검상태)}</td></tr>`).join('');
  }
  window._lastRows = rows;
}

/* ── 상세 시트 ── */
function openSheet(kind, idx, id){
  const r = (window._lastRows||[])[idx];
  if(!r) return;
  const displayProject=projectLabel(r), internalId=recordIdOf(r)||String(id||'');
  let html = '';
  if(kind==='pm' && r.출처==='정기점검 스케줄 원본'){
    html = `<h2>${esc2(displayProject)} ${chip(r.점검상태)}</h2>
      <div class="sub">${esc2(r.캠프명||'캠프 미상')} · 류지영 정기점검 스케줄 원본</div>
      <div class="card" style="margin:12px 0">
        <b>원본 일정 안내</b>
        <div class="sub" style="margin-top:5px">프로젝트번호가 확정되기 전의 캠프 일정입니다.
        날짜가 월까지만 보이면 원본에도 정확한 점검일이 아직 없습니다.</div>
      </div>
      <div class="dl">
        <dt>일정ID</dt><dd>${esc2(r.점검ID||'-')}</dd>
        <dt>캠프명</dt><dd>${esc2(r.캠프명||'-')}</dd>
        <dt>예정일·월</dt><dd>${esc2(r.점검예정일||'-')}</dd>
        <dt>담당기사</dt><dd>${esc2(r.담당기사||'미배정')}</dd>
        <dt>장비수</dt><dd>${esc2(r.장비수||'-')}</dd>
        <dt>장비내역</dt><dd>${esc2(r.장비내역||'-')}</dd>
        <dt>반영상태</dt><dd>${esc2(r.반영상태||'-')}</dd>
        <dt>원본 위치</dt><dd>${esc2((r.원본파일||'')+' '+(r.원본행?`(${r.원본행}행)`:'')||'-')}</dd>
      </div>`;
    showSheet(html);
    return;
  }
  if(kind==='settle'){
    html = `<h2>${esc2(displayProject)} ${chip(r.상태)}</h2>
    <div class="sub">${esc2(r.캠프명||'캠프 미상')} · ${esc2(r.업무구분||'')} · 내부기록 ${esc2(internalId||'-')}</div>
    <div class="dl">
      <div class="sec">작업</div>
      <dt>프로젝트NO</dt><dd>${esc2(displayProject)}</dd>
      <dt>원천업무</dt><dd>${r.원천업무ID||'-'}</dd>
      <dt>작업완료일</dt><dd>${r.완료일||'-'}</dd>
      <dt>비용구분</dt><dd>${r.비용구분||'-'}</dd>
      <dt>공급가액</dt><dd>${fmt(r.공급가액)}원</dd>
      <dt>부가세</dt><dd>${vatLine(r)}</dd>
      <dt>합계(VAT포함)</dt><dd>${fmt(r.합계)}원</dd>
      <div class="sec">쿠팡 PO</div>
      <dt>PO 필요여부</dt><dd>${r.PO필요||'-'}</dd>
      <dt>PO 번호</dt><dd>${r.PO번호||'없음'}</dd>
      <dt>PO 발행일</dt><dd>${r.PO발행일||'-'}</dd>
      <div class="sec">거래명세서 / ERP</div>
      <dt>명세서번호</dt><dd>${r.명세서번호||'없음'}</dd>
      <dt>발행일</dt><dd>${r.명세서발행일||'-'}</dd>
      <div class="sec">세금계산서</div>
      <dt>발행여부</dt><dd>${r.계산서}</dd>
      <dt>발행일</dt><dd>${r.계산서발행일||'-'}</dd>
      <dt>승인번호</dt><dd>${r.승인번호||'-'}</dd>
      <div class="sec">수금</div>
      <dt>입금일</dt><dd>${r.입금일||'-'}</dd>
      <dt>입금액</dt><dd>${r.입금액?fmt(r.입금액)+'원':'-'}</dd>
      <dt>미수금</dt><dd>${r.미수금===''?'-':fmt(r.미수금)+'원'}</dd>
      <div class="sec">4원천 검증 (카톡·밴드·ERP·쿠팡PO)</div>
      <dt>대조결과</dt><dd>${vbadge(r.정산ID)||'대조 데이터 없음 — inbox에 파일을 넣고 대조를 실행하세요'}</dd>
    </div>
    ${helpBox(r.상태)}
    ${inputForm('settle', r, r.정산ID)}`;
  } else {
    // 엑셀 검증결과를 맨 위에 놓고, 문제코드마다 "무엇이 왜 문제이고 어떻게 처리하는지"를 편다.
    const vres = r.검증결과||'', vcodes = r.검증문제코드||'';
    const vbox = vres ? `<div class="card" style="margin:0 0 10px">
        <div style="font-weight:800;margin-bottom:4px">엑셀 검증결과 — ${vres}
          ${r.관리자검증상태==='일치'?'<span class="oktag">✓ 관리자 확인완료</span>':''}</div>
        ${vcodes?`<div style="font-size:12px;color:var(--ink-2);margin-bottom:6px">${vcodes}</div>${codeHtml(vcodes)}`
                :'<div style="font-size:12px;color:var(--ink-3)">문제 없음</div>'}
      </div>` : '';
    html = `<h2>${esc2(displayProject)} ${chip(statOf(r))}</h2>
      <div class="sub">${esc2(r.캠프명||'캠프 미상')} · 내부기록 ${esc2(internalId||'-')}</div>` + vbox +
      '<div class="dl">' +
      Object.entries(r).map(([k,v])=>`<dt>${k}</dt><dd>${v||'-'}</dd>`).join('') + '</div>' +
      inputForm(kind, r, id);
  }
  showSheet(html);
}

/* ── 시트 열기·닫기 한 곳으로 ───────────────────────────────────────────────
   폰에서 두 가지가 불편했다(사용자 지적 2026-07-28):
     ① 시트를 띄운 채 **뒤로가기를 누르면 앱이 통째로 뒤로 간다** — 시트만 닫혀야 한다.
     ② 시트 안에서 스크롤하면 **뒤 화면이 같이 움직인다**(스크롤 체이닝).
   목록 → 상세로 들어간 경우에는 뒤로가기가 **목록으로 돌아가야** 한다.
   그래서 시트 내용을 스택으로 들고 있다가 한 겹씩 되돌린다.                     */
const _sheetStack = [];      // 시트 안에서 다시 연 화면의 '이전 내용'
let _bgScroll = null;        // 배경을 잠글 때 저장한 스크롤 위치

/* ── 뒤로가기 한 곳으로 (사용자 지시 2026-07-28) ─────────────────────────────
   "뒤로 가면 바로 전 화면. 전 화면이 없으면 대시보드. 대시보드가 홈이고,
    대시보드에서 다시 뒤로 가면 앱 종료."

   시트(상세창)와 탭(대시보드·정산·실행·보고·기록)을 **한 스택**으로 다룬다.
   따로 두면 시트를 닫는 뒤로가기와 탭을 되돌리는 뒤로가기가 서로 어긋난다.

   ★ 되돌린 뒤에는 항상 **여분 항목을 다시 쌓는다**(navGuard). 안 그러면 다음
     뒤로가기가 우리 손을 떠나 브라우저가 앱 밖으로 나가 버린다 — 홈에서
     종료할지 물어볼 기회조차 없어진다.                                        */
const _nav = [];             // 우리가 history 에 쌓아 둔 것들(뒤로가기 1회 = 1개 되돌리기)
let _navSkip = 0;            // history.go(-n) 이 부를 popstate 를 그만큼 무시한다

function navPush(entry){
  _nav.push(entry);
  history.pushState({csosNav:_nav.length}, '');
}
function navGuard(){ history.pushState({csosNav:'guard'}, ''); }

function exitApp(){
  if(!confirm('앱을 종료할까요?')){ navGuard(); return; }
  // 설치형 앱 창은 스크립트로 닫힌다. 일반 탭이면 브라우저가 막으므로 안내만 남긴다.
  window.close();
  setTimeout(()=>{ toast('창을 직접 닫아 주세요 — 브라우저가 자동 종료를 막았습니다'); navGuard(); }, 400);
}

window.addEventListener('popstate', ()=>{
  // closeSheetAll 이 history.go(-n) 으로 되돌린 몫은 이미 처리했으니 건너뛴다.
  if(_navSkip > 0){ _navSkip--; return; }
  const e = _nav.pop();
  if(e && e.t === 'sheet'){ closeSheet(true); navGuard(); return; }
  if(e && e.t === 'tab'){ applyView(e.from); navGuard(); return; }
  // 되돌릴 게 없다 — 홈(대시보드)이 아니면 홈으로, 홈이면 종료를 묻는다.
  if(curView() !== 'dash'){ applyView('dash'); navGuard(); return; }
  exitApp();
});

function sheetIsOpen(){ return $('sheetbg').style.display === 'block'; }

/* 배경 잠금 — body{overflow:hidden} 만으로는 iOS 사파리가 계속 움직인다.
   위치를 고정하고 스크롤 값을 기억했다가 닫을 때 그대로 돌려놓는다. */
function lockBackground(){
  if(_bgScroll !== null) return;
  _bgScroll = window.scrollY || document.documentElement.scrollTop || 0;
  const b = document.body;
  b.style.position = 'fixed'; b.style.top = (-_bgScroll) + 'px';
  b.style.left = '0'; b.style.right = '0'; b.style.width = '100%';
}
function unlockBackground(){
  if(_bgScroll === null) return;
  const b = document.body;
  b.style.position = ''; b.style.top = ''; b.style.left = ''; b.style.right = ''; b.style.width = '';
  window.scrollTo(0, _bgScroll);
  _bgScroll = null;
}

function sheetSnapshot(){
  return {
    body: $('sheetbody').innerHTML,
    actions: $('sheetactions').innerHTML,
    scrollTop: $('sheetbody').scrollTop
  };
}
function setSheetContent(html, actionsHtml){
  const body = $('sheetbody'), foot = $('sheetactions');
  body.innerHTML = html || '';
  foot.innerHTML = actionsHtml || '';
  if(!actionsHtml){
    const actions = body.querySelector('.actions.sticky,.media-tools.sticky');
    if(actions){ actions.remove(); foot.appendChild(actions); }
  }
  foot.classList.toggle('open', !!foot.children.length);
}
function clearSheetActions(){
  const foot = $('sheetactions');
  foot.classList.remove('open');
  foot.innerHTML = '';
}
function showSheet(html){
  const already = sheetIsOpen();
  if(already) _sheetStack.push(sheetSnapshot());             // 내용·버튼·스크롤 위치를 함께 보관
  setSheetContent(html);
  $('sheetbody').scrollTop = 0;                              // 새 화면은 항상 맨 위부터
  if(!already){ $('sheetbg').style.display = 'block'; lockBackground(); }
  // rAF는 화면이 가려진 탭에서 실행되지 않아 시트가 안 열릴 수 있다 → 타이머 사용
  setTimeout(()=>{
    $('sheet').classList.add('open');
    try{ $('sheetbody').focus({preventScroll:true}); }catch(_){}
  }, 10);
  navPush({t:'sheet'});                                      // 뒤로가기 한 번 = 이 화면 하나
}

/* fromPop=true 는 popstate 가 부른 것. 사용자가 X·배경을 눌러 닫을 때는
   history 를 먼저 되돌려야 뒤로가기 횟수가 어긋나지 않는다. */
function closeSheet(fromPop){
  if(!fromPop && _nav.length && _nav[_nav.length-1].t === 'sheet'){ history.back(); return; }
  if(_sheetStack.length){                                     // 목록 → 상세였다면 목록으로
    const prev = _sheetStack.pop();
    setSheetContent(prev.body, prev.actions);
    $('sheetbody').scrollTop = prev.scrollTop || 0;
    try{ $('sheetbody').focus({preventScroll:true}); }catch(_){}
    return;
  }
  $('sheet').classList.remove('open');
  clearSheetActions();
  setTimeout(()=>{ $('sheetbg').style.display = 'none'; unlockBackground(); }, 220);
}

/* 시트를 통째로 닫는다(저장 후처럼 목록으로 돌아갈 필요가 없을 때). */
function closeSheetAll(){
  // 시트로 쌓인 것만 걷어낸다. 탭 이동 기록은 남겨야 뒤로가기가 이전 탭으로 간다.
  let back = 0;
  while(_nav.length && _nav[_nav.length-1].t === 'sheet'){ _nav.pop(); back++; }
  _sheetStack.length = 0;
  $('sheet').classList.remove('open');
  clearSheetActions();
  setTimeout(()=>{ $('sheetbg').style.display = 'none'; unlockBackground(); }, 220);
  // ★ history.go(-n) 은 몇 칸을 건너뛰든 popstate 를 **한 번만** 발생시킨다.
  //   n 만큼 세면 그 차이가 남아 **다음 뒤로가기를 통째로 먹는다**(실제로 겪음).
  if(back > 0){ _navSkip += 1; history.go(-back); }
}

/* 버튼 위에서 휠을 돌려도 바로 위 목록이 움직이게 한다. PC와 키보드 접근도 같은 스크롤포트를 쓴다. */
$('sheetactions').addEventListener('wheel', e=>{
  if(!e.deltaY) return;
  $('sheetbody').scrollTop += e.deltaY;
  e.preventDefault();
}, {passive:false});
$('sheetbody').addEventListener('keydown', e=>{
  const body=$('sheetbody'), page=Math.max(120, body.clientHeight*.82);
  if(e.key==='PageDown'){ body.scrollTop+=page; e.preventDefault(); }
  else if(e.key==='PageUp'){ body.scrollTop-=page; e.preventDefault(); }
  else if(e.key==='Home'){ body.scrollTop=0; e.preventDefault(); }
  else if(e.key==='End'){ body.scrollTop=body.scrollHeight; e.preventDefault(); }
});

/* ── 앱 ↔ 엑셀 양방향 입력 (빈 칸만·근거 기록·즉시 또는 09:50 반영) ── */
const INPUT_SPEC = {
  settle: { sheet:'06_거래서류청구수금', key_col:'정산ID', fields:[
    {label:'PO번호', col:'PO번호', vtype:'text', cur:r=>r.PO번호, ph:'예: PO367787'},
    {label:'PO발행일', col:'PO발행일', vtype:'date', cur:r=>r.PO발행일, ph:'YYYY-MM-DD'},
    {label:'명세서번호', col:'거래명세서번호', vtype:'text',  cur:r=>r.명세서번호, ph:'예: 2026/07/24-1'},
    {label:'명세서발행일', col:'거래명세서발행일', vtype:'date', cur:r=>r.명세서발행일, ph:'YYYY-MM-DD'},
    {label:'세금계산서발행일', col:'세금계산서발행일', vtype:'date', cur:r=>r.계산서발행일, ph:'YYYY-MM-DD'},
    {label:'청구일', col:'청구일', vtype:'date', cur:r=>r.청구일, ph:'YYYY-MM-DD'},
    {label:'지급예정일', col:'지급예정일', vtype:'date', cur:r=>r.지급예정일, ph:'YYYY-MM-DD'},
    {label:'입금일', col:'입금일', vtype:'date', cur:r=>r.입금일, ph:'YYYY-MM-DD'},
    {label:'입금액', col:'입금액', vtype:'number', cur:r=>r.입금액, ph:'숫자만'},
  ]},
  as: { sheet:'02_돌발AS접수', key_col:'접수ID', fields:[
    {label:'담당기사', col:'담당기사', vtype:'text', cur:r=>r.담당기사, opts:'담당기사'},
    {label:'방문예정일', col:'방문예정일', vtype:'date', cur:r=>r.방문예정일, ph:'YYYY-MM-DD'},
    {label:'작업완료일', col:'작업완료일', vtype:'date', cur:r=>r.작업완료일, ph:'YYYY-MM-DD'},
    {label:'관리자검증상태', col:'관리자검증상태', vtype:'text', cur:r=>r.관리자검증상태,
     opts:'관리자검증상태'},
    {label:'최종확인일', col:'최종확인일', vtype:'date', cur:r=>r.최종확인일, ph:'YYYY-MM-DD'},
  ]},
  pm: { sheet:'04_정기점검', key_col:'점검ID', fields:[
    {label:'실제점검일', col:'실제점검일', vtype:'date', cur:r=>r.실제점검일, ph:'YYYY-MM-DD'},
    {label:'최종확인일(유현민 체크)', col:'최종확인일(유현민 체크)', vtype:'date',
     cur:r=>r['최종확인일(유현민 체크)'], ph:'YYYY-MM-DD'},
  ]},
};
/* 코드 목록(드롭다운 선택지)은 10_코드관리 시트에서 온다 — 화면에 박아 두면 시트와 어긋난다 */
let CODES = {};
async function loadCodes(){
  try{ CODES = await api('/api/codes'); }catch(e){ CODES = {}; }
}

/* 폰에서 치기 편해야 한다.
   · 날짜  → type="date"  : 폰이 **달력**을 띄운다(YYYY-MM-DD를 손으로 칠 일이 없다)
   · 선택지 → <select>    : 오타·표기흔들림이 사라진다(시트 드롭다운과 같은 값만 들어간다)
   · 숫자  → inputmode="numeric" : 숫자 키패드가 뜬다
   ★ 글자 크기를 16px 미만으로 두면 iOS 사파리가 입력할 때 화면을 확대해 버린다 — 16px 유지. */
function fieldInput(f, i){
  const base = 'flex:1;min-width:0;padding:12px 11px;border:1.5px solid var(--line);' +
               'border-radius:10px;font-size:16px;background:var(--card);color:var(--ink-1);' +
               'font-family:inherit';
  const attr = `id="inp_${i}" data-col="${f.col}" data-vtype="${f.vtype}"`;
  if(f.opts){
    let list = CODES[f.opts] || [];
    /* 시트 순서 그대로 두면 폰에서 **관리자 이름 6개를 지나야** 실제 기사가 나온다.
       실제로 배정된 횟수가 많은 사람을 위로 올린다 — 시트는 그대로 두고 보여주는 순서만
       바꾸는 것이라, 사람이 바뀌면 다음 날 자동으로 따라간다. */
    if(f.opts === '담당기사') list = byUsage(list);
    if(list.length){
      return `<select ${attr} style="${base}">` +
             `<option value="">— 선택 —</option>` +
             list.map(v=>`<option value="${esc2(v)}">${esc2(v)}</option>`).join('') +
             `</select>`;
    }
    // 코드 목록을 못 받았으면 자유 입력으로 떨어뜨린다(입력 자체를 막으면 안 된다)
  }
  if(f.vtype === 'date')
    return `<input ${attr} type="date" style="${base}">`;
  if(f.vtype === 'number')
    return `<input ${attr} type="text" inputmode="numeric" placeholder="${f.ph||'숫자만'}" style="${base}">`;
  return `<input ${attr} placeholder="${f.ph||''}" style="${base}">`;
}
/* 원장에 실제로 배정된 횟수 순으로 정렬(한 칸에 두 명인 건도 각각 센다) */
function byUsage(list){
  const cnt = {};
  ((works.as||[]).concat(works.pm||[])).forEach(r=>{
    String(r.담당기사||'').split(/[,./·]| 및 /).forEach(t=>{
      t=t.trim(); if(t) cnt[t]=(cnt[t]||0)+1;
    });
  });
  const used = list.filter(v=>cnt[v]).sort((a,b)=>cnt[b]-cnt[a]);
  const rest = list.filter(v=>!cnt[v]);          // 한 번도 안 나온 이름은 뒤로
  return used.concat(rest);
}

/* 부가세 표기.
   ★ 원장에 적힌 값을 그대로 쓴다. 공급가액×10%로 계산해 버리면 반올림 때문에 서류와
     1원씩 어긋날 수 있고, 화면 숫자가 거래명세서와 다르면 그 순간 앱을 못 믿게 된다.
   원장이 비어 있을 때만 합계−공급가액으로 되짚고, 그 사실을 화면에 밝힌다.
   셋이 안 맞으면(공급가+부가세≠합계) 숨기지 않고 빨갛게 알린다 — 원장이 틀린 것이다. */
function vatLine(r){
  const sup = +r.공급가액 || 0, tot = +r.합계 || 0;
  const has = r.부가세 !== null && r.부가세 !== undefined && r.부가세 !== '';
  const vat = has ? (+r.부가세 || 0) : (tot ? tot - sup : 0);
  if(!has && !tot) return '-';
  let out = fmt(vat) + '원';
  if(!has) out += ' <span class="norm">(합계−공급가액)</span>';
  if(tot && Math.abs(sup + vat - tot) > 1)
    out += ` <span style="color:var(--danger);font-weight:800">⚠ 공급가+부가세 ${fmt(sup+vat)}원 ≠ 합계</span>`;
  return out;
}

function esc2(s){ return String(s==null?'':s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

function inputForm(kind, r, id){
  const spec = INPUT_SPEC[kind];
  if(!spec) return '';
  const empty = spec.fields.filter(f=>!f.cur(r));
  if(!empty.length) return '';
  return `<div class="sec" style="border-top:1px solid var(--line);margin-top:14px;padding-top:12px">빈 항목 입력 → 엑셀 반영 (빈 칸만·기존 값 보호·새 버전 생성)</div>
    <div style="display:flex;flex-direction:column;gap:8px;margin-top:8px">` +
    empty.map((f,i)=>`<div style="display:flex;gap:8px;align-items:center">
      <span style="width:110px;font-size:12.5px;color:var(--ink-3);font-weight:600">${f.label}</span>
      ${fieldInput(f, i)}
    </div>`).join('') +
    `</div><button onclick="saveInputs('${kind}','${id}',${empty.length})" style="margin-top:10px;width:100%;
      padding:12px;background:var(--brand);color:#fff;border-radius:11px;font-weight:700;font-size:14px">💾 입력 저장</button>
    <div style="font-size:11px;color:var(--ink-3);margin-top:6px">저장 후 "지금 반영"을 선택하면 몇 초 안에 엑셀 새 버전에 기록됩니다. (아니면 09:50 에이전트가 자동 반영)</div>`;
}
async function saveInputs(kind, id, n){
  const spec = INPUT_SPEC[kind];
  let queued = 0;
  for(let i=0;i<n;i++){
    const el = $('inp_'+i);
    if(!el || !el.value.trim()) continue;
    const r = await api('/api/input',{method:'POST',body:JSON.stringify(
      {sheet:spec.sheet, key_col:spec.key_col, key:id,
       col:el.dataset.col, value:el.value.trim(), vtype:el.dataset.vtype})});
    if(r.ok) queued++;
  }
  if(!queued){ alert('입력된 값이 없습니다'); return; }
  closeSheetAll();      // 저장이 끝났으면 목록으로 되돌아갈 게 아니라 통째로 닫는다
  if(confirm(`${queued}건 저장됨.\n\n지금 바로 엑셀에 반영할까요? (새 버전 vN+1 생성)\n[취소]하면 09:50 에이전트가 자동 반영합니다.`)){
    runTask('writer_apply');
  } else {
    loadStatus();
  }
}

/* ── 월별·연도별 현황 ── */
function dateOf(r, kind){
  // 표기 형식이 섞여 있어도(2026.6.3 / 2026-06-03) 항상 YYYY-MM-DD로 맞춘다 —
  // 그래야 문자열 비교만으로 정확한 시간순 정렬·월별 집계가 된다.
  const v = kind==='settle' ? (r.완료일||'')
          : kind==='as'     ? (r.작업완료일||r.접수일자||'')
          : kind==='check'  ? anyDate(r)
          : (r.실제점검일||r.점검예정일||'');
  return normDate(v) || String(v||'').slice(0,10);
}
function periodInit(){
  const py = $('pyear'), pf = $('pfrom'), pt = $('pto');
  const keepF = pf.value, keepT = pt.value;
  py.innerHTML = `<option>${APP_YEAR}</option>`;
  const opts = Array.from({length:12},(_,i)=>
    `<option value="${String(i+1).padStart(2,'0')}">${i+1}월</option>`).join('');
  pf.innerHTML = opts; pt.innerHTML = opts;
  py.value = APP_YEAR;
  if(!keepF && !window.__pinit){
    // 최초: 데이터가 있는 최신 월 한 달만 — 처음 열었을 때 가장 궁금한 게 이번 달이다
    const mm = settleRows.map(r=>dateOf(r,'settle')).filter(d=>d.startsWith(py.value))
                         .map(d=>d.slice(5,7)).sort().pop() || String(new Date().getMonth()+1).padStart(2,'0');
    pf.value = mm; pt.value = mm;
    window.__pinit = true;
  } else { pf.value = keepF || '01'; pt.value = keepT || '12'; }
  renderPeriod();
}
function pickYear(){ renderPeriod(); }
/* 시작이 끝보다 뒤면 사용자가 틀린 게 아니라 **다른 쪽을 아직 못 옮긴 것**이다.
   방금 고른 쪽을 존중하고 반대쪽을 끌어와 맞춘다 — 오류 메시지를 띄우지 않는다. */
function fixRange(which){
  const pf = $('pfrom'), pt = $('pto');
  if(pf.value > pt.value){ if(which==='from') pt.value = pf.value; else pf.value = pt.value; }
  renderPeriod();
}
function setRange(a, b){
  $('pfrom').value = String(a).padStart(2,'0');
  $('pto').value   = String(b).padStart(2,'0');
  renderPeriod();
}
/* 기간 = [시작월, 끝월]. 한 달만 보려면 둘을 같게 둔다. */
function periodRange(){
  const f = ($('pfrom')||{}).value || '01', t = ($('pto')||{}).value || '12';
  return f <= t ? [f, t] : [t, f];
}
function inPeriod(d, y, from, to){
  if(!d || !d.startsWith(y)) return false;
  if(from === undefined){ const r = periodRange(); from = r[0]; to = r[1]; }
  if(to === undefined) to = from;              // 한 달만 넘긴 경우(옛 호출부 호환)
  const mm = d.slice(5,7);
  return mm >= from && mm <= to;
}
/* 화면에 적을 기간 이름 — '7월' 또는 '1~6월' 또는 '연간 전체' */
function periodLabel(){
  const [f, t] = periodRange();
  if(f === '01' && t === '12') return '연간 전체';
  return f === t ? `${+f}월` : `${+f}~${+t}월`;
}
function renderPeriod(){
  const y = $('pyear').value;
  if(!y) return;
  const [pf, pt] = periodRange();          // 선택한 기간(같으면 한 달)
  const m = (pf === pt) ? pf : '';         // 막대 강조는 한 달일 때만
  // 12개월 막대 — 관리대장 정산(파랑) + ERP 실매출(회색)을 함께 본다.
  // 06시트에 1~6월 정산 행이 없어 파랑만 보면 그 달이 '실적 0'처럼 보이기 때문.
  const byM = Array.from({length:12},()=>({amt:0,cnt:0,erp:0,ecnt:0}));
  settleRows.forEach(r=>{ const d=dateOf(r,'settle');
    if(d.startsWith(y)){ const i=+d.slice(5,7)-1; if(i>=0&&i<12){ byM[i].amt+=(+r.공급가액||0); byM[i].cnt++; } } });
  Object.entries(erpDocs.months||{}).forEach(([k,v])=>{
    if(k.slice(0,4)!==y) return;
    const i = +k.slice(5,7)-1;
    if(i>=0&&i<12){ byM[i].erp = v.합계||0; byM[i].ecnt = v.건수||0; }
  });
  const max = Math.max(1, ...byM.map(x=>Math.max(x.amt, x.erp)));
  $('mbars').innerHTML = byM.map((x,i)=>{
    const mm = String(i+1).padStart(2,'0');
    const sel = m===mm;
    return `<div class="mbar ${sel?'on':''}" onclick="setRange(${sel?1:+mm},${sel?12:+mm})"
      title="${i+1}월 — 대장 정산 ${fmt(x.amt)}원(${x.cnt}건) / ERP 계산서 ${fmt(x.erp)}원(${x.ecnt}장)">
      ${sel&&(x.amt||x.erp)?`<span style="font-size:9px;font-weight:800;color:var(--brand);white-space:nowrap">${Math.round((x.amt||x.erp)/10000).toLocaleString()}만</span>`:''}
      <span class="bstack">
        <i class="erp" style="height:${Math.round(x.erp/max*76)+(x.erp?2:0)}px"></i>
        <i style="height:${Math.round(x.amt/max*76)+(x.amt?2:0)}px"></i>
      </span><b>${i+1}월</b></div>`;
  }).join('');
  $('blegend').innerHTML = byM.some(x=>x.erp)
    ? '<span><i class="k1"></i>대장 정산</span><span><i class="k2"></i>ERP 계산서(발행 기준)</span>' : '';
  // 연간 전체 선택 시: 월별 통계표
  if(!m){
    const mt = byM.map((x,i)=>{
      const mm2 = String(i+1).padStart(2,'0');
      const Si = settleRows.filter(r=>inPeriod(dateOf(r,'settle'),y,mm2));
      return {월:(i+1)+'월', 건수:x.cnt, 공급가액:x.amt, ERP:x.erp, ERP장:x.ecnt,
        입금:Si.reduce((a,r)=>a+(+r.입금액||0),0),
        미수:Si.reduce((a,r)=>a+(+r.미수금||0),0),
        AS:(works.as||[]).filter(r=>inPeriod(dateOf(r,'as'),y,mm2)).length,
        점검:(works.pm||[]).filter(r=>inPeriod(dateOf(r,'pm'),y,mm2)).length};
    }).filter(r=>r.건수||r.AS||r.점검||r.ERP);
    $('ptable').innerHTML = mt.length ? `<table class="grid ptbl" style="display:table">
      <tr><th>월</th><th class="num">정산</th><th class="num">공급가액</th><th class="num">ERP 계산서</th><th class="num">입금</th><th class="num">미수</th><th class="num">AS</th><th class="num">점검</th></tr>` +
      mt.map(r=>`<tr onclick="setRange(+'${r.월.replace('월','')}',+'${r.월.replace('월','')}')" style="cursor:pointer">
        <td><b>${r.월}</b></td><td class="num">${r.건수}</td><td class="num">${fmt(r.공급가액)}</td>
        <td class="num">${r.ERP?fmt(r.ERP):'-'}</td><td class="num">${fmt(r.입금)}</td><td class="num">${fmt(r.미수)}</td><td class="num">${r.AS}</td><td class="num">${r.점검}</td></tr>`).join('') +
      `<tr class="tot"><td>합계</td><td class="num">${mt.reduce((a,r)=>a+r.건수,0)}</td>
       <td class="num">${fmt(mt.reduce((a,r)=>a+r.공급가액,0))}</td>
       <td class="num">${fmt(mt.reduce((a,r)=>a+r.ERP,0))}</td><td class="num">${fmt(mt.reduce((a,r)=>a+r.입금,0))}</td>
       <td class="num">${fmt(mt.reduce((a,r)=>a+r.미수,0))}</td><td class="num">${mt.reduce((a,r)=>a+r.AS,0)}</td>
       <td class="num">${mt.reduce((a,r)=>a+r.점검,0)}</td></tr></table>` : '';
  } else $('ptable').innerHTML = '';
  // 선택 기간 KPI
  const S = settleRows.filter(r=>inPeriod(dateOf(r,'settle'),y,pf,pt));
  const A = (works.as||[]).filter(r=>inPeriod(dateOf(r,'as'),y,pf,pt));
  const P = (works.pm||[]).filter(r=>inPeriod(dateOf(r,'pm'),y,pf,pt));
  const 공급 = S.reduce((a,r)=>a+(+r.공급가액||0),0);
  const 입금 = S.reduce((a,r)=>a+(+r.입금액||0),0);
  const 미수 = S.reduce((a,r)=>a+(+r.미수금||0),0);
  const 문제 = S.filter(r=>needAction(r)).length;
  $('ptitle').textContent = ` — ${y}년 ${periodLabel()}`;
  $('pkpis').innerHTML = [
    ['정산',`${S.length}건`],['공급가액',fmt(공급)],['입금',fmt(입금)],
    ['미수금',fmt(미수)],['조치필요',`${문제}건`],['AS·점검',`${A.length}·${P.length}건`]
  ].map(([l,v])=>`<div class="pk"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('');
}

/* ── 실행 ── */
function confirmRun(key,msg){ if(confirm(msg+'\n\n진행할까요?')) runTask(key); }
/* 워크벤치 대체: 로컬 파일·폴더 열기 (서버가 도는 PC에서만 허용) */
async function openLocal(what){
  try{
    const r = await api('/api/open',{method:'POST',body:JSON.stringify({what})});
    if(r.ok) toast('열었습니다 · ' + (r.opened||what));
    else alert(r.error || '열지 못했습니다');
  }catch(e){ alert('사무실 PC에서만 사용할 수 있는 기능입니다.'); }
}
async function runTask(key){
  show('run');
  const r = await api('/api/run/'+key,{method:'POST'});
  if(!r.ok && r.msg!=='demo'){ alert(r.msg); return; }
  if(!polling) polling = setInterval(pollLog, 1200);
  pollLog();
}
async function pollLog(){
  const d = await api('/api/tasklog');
  $('livedot').className = 'dot' + (d.busy?' busy':'');
  document.body.classList.toggle('agent-running', !!d.busy);
  if(d.busy){
    $('heroAgent').textContent = `에이전트 실행 중${d.task?' · '+d.task:''}`;
    if(!polling) polling = setInterval(pollLog, 1200);
  }
  if(d.log.length) { $('logbox').textContent = d.log.join('\n'); $('logbox').scrollTop = 1e9; }
  if(!d.busy && polling){ clearInterval(polling); polling=null; loadStatus(); loadSettle(); }
}

/* ── 리포트 ── */
function mdRender(t){
  const esc = s=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;');
  let html='', intbl=false;
  for(const ln of t.split('\n')){
    if(/^\|/.test(ln)){
      if(/^\|[\s\-|]+\|$/.test(ln)) continue;
      const cells = ln.split('|').slice(1,-1).map(c=>esc(c.trim()));
      if(!intbl){ html+='<table><tr>'+cells.map(c=>`<th>${c}</th>`).join('')+'</tr>'; intbl=true; }
      else html+='<tr>'+cells.map(c=>`<td>${c}</td>`).join('')+'</tr>';
      continue;
    }
    if(intbl){ html+='</table>'; intbl=false; }
    if(/^# /.test(ln)) html+=`<h1>${esc(ln.slice(2))}</h1>`;
    else if(/^## /.test(ln)) html+=`<h2>${esc(ln.slice(3))}</h2>`;
    else if(/^- /.test(ln)) html+=`<li>${esc(ln.slice(2))}</li>`;
    else if(ln.trim()) html+=`<p>${esc(ln)}</p>`;
  }
  if(intbl) html+='</table>';
  return html;
}
/* ── 구글 캘린더(COUPANG 설치+납품+AS) ──────────────────────────
   서버가 gcal_sync 캐시를 그대로 준다. 앱은 구글을 직접 부르지 않는다 —
   폰에서 열 때 외부를 기다리면 화면이 멈춘다. */
/* ★ "원장엔 넣었다는데 앱엔 왜 없지?" 를 없앤다.
   06시트 정산ID·금액은 **수식**이라 엑셀이 한 번 계산해야 값이 생긴다. 그때까지 앱은
   옛 건수만 읽는다 — 숫자가 틀린 게 아니라 아직 안 나온 것이다. 그 사실을 화면이 말한다.
   (도구로 대신 계산할 수 없다: 엑셀 수식은 엑셀만 계산하고, 빈 행에 미리 써넣으면
    v259 처럼 엉뚱한 건에 붙는다) */
function renderRecalc(rc){
  const card = $('recalccard');
  if(!card) return;
  const n = (rc && rc.대기합계) || 0;
  if(!n){ card.style.display='none'; return; }
  const 목록 = (rc.항목||[]).map(x=>`<b>${esc2(x.이름)} ${x.대기.toLocaleString()}건</b>`).join(' · ');
  card.style.display='block';
  $('recalcmsg').innerHTML =
    `관리대장에는 이미 올라와 있지만, 엑셀이 아직 계산하지 않아 <b>이 앱에 안 나오는 건이 ${n.toLocaleString()}건</b> 있습니다.<br>
     ${목록}<br>
     <span style="color:var(--muted)">숫자가 틀린 게 아니라 <b>대기 중</b>입니다.
     관리대장을 엑셀에서 한 번 열었다 닫으면 계산이 끝나고 바로 반영됩니다.</span>`;
}

var CAL = null;   // let 이면 applyView 가 먼저 불릴 때 TDZ 로 죽는다
/* 캘린더도 같은 관문을 지난다 — 철거·납품 일정은 원장(05시트)에 계속 쌓이지만 화면에선 뺀다.
   SHOW_SIDE_WORK 를 켜면 여기도 자동으로 같이 나온다. 목록·카운트·시트 전부 이 함수만 쓴다. */
function calEvents(){
  return ((CAL && CAL.일정) || []).filter(e=>!isSideWork(e));
}
function calHiddenCount(){
  return ((CAL && CAL.일정) || []).length - calEvents().length;
}
async function loadCalendar(){
  try{ CAL = await api('/api/calendar'); }catch(e){ CAL = null; }
  const box = $('callist'), sub = $('calsub');
  if(!box){ renderCalendarPage(); return; }
  const evs = calEvents(), 숨김 = calHiddenCount();
  if(!evs.length){
    /* 전부 철거·납품이라 걸러진 경우와 아예 안 들어온 경우는 다르다.
       "일정이 없다"고만 하면 수집이 고장난 줄 안다. */
    sub.textContent = 숨김 ? `· 철거·납품 ${숨김}건 보관 중` : '';
    box.innerHTML = '<div class="muted">' + esc2(숨김
      ? `표시할 일정이 없습니다 (철거·납품 ${숨김}건은 관리대장에만 보관)`
      : ((CAL && CAL.안내) || '아직 수집된 일정이 없습니다')) + '</div>';
    renderCalendarPage();
    return;
  }
  const today = new Date(); today.setHours(0,0,0,0);
  const soon = evs.filter(e=>{ const d=new Date(e.날짜+'T00:00:00'); return d>=today; }).slice(0,5);
  const 미연결 = evs.filter(e=>!e.원천업무ID).length;
  sub.textContent = `· 2026년 ${evs.length}건` + (미연결 ? ` · 원장 미연결 ${미연결}건` : '')
                  + (숨김 ? ` · 철거·납품 ${숨김}건 보관` : '');
  box.innerHTML = soon.length
    ? soon.map(calLine).join('')
    : '<div class="muted">앞으로 예정된 일정이 없습니다</div>';
  renderCalendarPage();
}
function calLine(e){
  const tag = e.원천업무ID
    ? `<span class="chip c-ok">${esc2(e.원천업무ID)}</span>`
    : '<span class="chip c-warn">원장 미연결</span>';
  return `<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;
      padding:7px 0;border-bottom:1px solid var(--line)">
    <b style="min-width:112px">${esc2(e.날짜)}${e.시간 ? ' ' + esc2(e.시간) : ''}</b>
    <span style="flex:1;min-width:120px">${esc2(e.업무구분||'미상')} · ${esc2(e.제목)}</span>${tag}</div>`;
}
let calendarSelected = 0;
const COUPANG_CALENDAR_ID = '053a21a7e074aac2e58c27f305247c1cd6481d50841c1dfc9b83cdad2d3e876b@group.calendar.google.com';

function calendarRows(){
  return calEvents().slice().sort((a,b)=>{
    const aa=`${a.날짜||''} ${a.시간||''}`, bb=`${b.날짜||''} ${b.시간||''}`;
    return aa.localeCompare(bb, 'ko');
  });
}
function calendarWhen(e){
  return `${e.날짜||'-'}${e.시간 ? ` ${e.시간}` : ' 종일'}`;
}
function renderCalendarPage(){
  const list=$('calendarEventList'), meta=$('calendarMeta'), detail=$('calendarDetail');
  if(!list || !meta || !detail) return;
  const rows=calendarRows(), hidden=calHiddenCount(), unmatched=rows.filter(e=>!e.원천업무ID).length;
  meta.textContent = `갱신 ${(CAL&&CAL.갱신)||'-'} · 2026년 ${rows.length}건 · 원장 미연결 ${unmatched}건`+
    (hidden ? ` · 철거·납품 ${hidden}건은 관리대장 보관` : '');
  if(!rows.length){
    list.innerHTML=`<div class="muted">${esc2((CAL&&CAL.안내)||'아직 수집된 원장 대조 일정이 없습니다.')}</div>`;
    detail.textContent='Google Calendar의 전체 일정은 왼쪽 월간 화면에서 확인할 수 있습니다.';
    return;
  }
  calendarSelected=Math.max(0, Math.min(calendarSelected, rows.length-1));
  list.innerHTML=rows.map((e,i)=>`<button type="button" class="calendar-event ${i===calendarSelected?'on':''}" onclick="selectCalendarEvent(${i})">
    <span class="when">${esc2(calendarWhen(e))}</span><span><span class="what">${esc2(e.제목||'제목 없음')}</span>
    <span class="sub">${esc2(e.프로젝트NO||'프로젝트 미확정')} · ${esc2(e.캠프명||e.장소||e.업무구분||'캠프 미상')}</span></span></button>`).join('');
  const e=rows[calendarSelected];
  detail.innerHTML=`<div class="detail-title"><b>${esc2(e.제목||'제목 없음')}</b></div><dl>
    <dt>일시</dt><dd>${esc2(calendarWhen(e))}</dd>
    <dt>프로젝트</dt><dd>${esc2(e.프로젝트NO||'미확정')}</dd>
    <dt>캠프·장소</dt><dd>${esc2(e.캠프명||e.장소||'미입력')}</dd>
    <dt>업무구분</dt><dd>${esc2(e.업무구분||'미상')}</dd>
    <dt>원장 연결</dt><dd>${esc2(e.원천업무ID||'미연결 — 확인 필요')}</dd>
    <dt>연결 근거</dt><dd>${esc2(e.연결근거||'없음')}</dd></dl>`;
}
function selectCalendarEvent(index){ calendarSelected=index; renderCalendarPage(); }
async function refreshCalendarPage(){ await loadCalendar(); toast('캘린더 대조 목록을 새로고침했습니다'); }
function openCalendar(){ show('calendar'); }
function focusCalendarEntry(){
  if(curView()!=='calendar') show('calendar');
  setTimeout(()=>$('calendarTitle') && $('calendarTitle').focus(), 80);
}
function calendarDraft(){
  const value=id=>($(''+id)&&$(''+id).value.trim())||'';
  return {title:value('calendarTitle'), date:value('calendarDate'), start:value('calendarStart'),
          end:value('calendarEnd'), project:value('calendarProject'), camp:value('calendarCamp'), note:value('calendarNote')};
}
function previewCalendarDraft(){
  const d=calendarDraft(), out=$('calendarDraftPreview'); if(!out) return;
  const lines=[d.title||'일정 제목 미입력', `${d.date||'날짜 미입력'} ${d.start||'시작 시간 미입력'}${d.end?' ~ '+d.end:''}`,
    d.project?`프로젝트: ${d.project}`:'', d.camp?`캠프: ${d.camp}`:'', d.note?`내용: ${d.note}`:''].filter(Boolean);
  out.textContent=lines.join('\n');
}
function clearCalendarDraft(){
  ['calendarTitle','calendarDate','calendarProject','calendarCamp','calendarNote'].forEach(id=>{ if($(id)) $(id).value=''; });
  if($('calendarStart')) $('calendarStart').value='09:00';
  if($('calendarEnd')) $('calendarEnd').value='10:00';
  previewCalendarDraft();
}
function openGoogleCalendarDraft(){
  const d=calendarDraft();
  if(!d.title || !d.date){ alert('일정 제목과 날짜를 먼저 입력해 주세요.'); return; }
  const date=d.date.replace(/-/g,''), start=(d.start||'09:00').replace(':',''), end=(d.end||'10:00').replace(':','');
  const details=[d.project&&`프로젝트NO: ${d.project}`,d.camp&&`캠프명: ${d.camp}`,d.note].filter(Boolean).join('\n');
  const u=new URL('https://calendar.google.com/calendar/render');
  u.searchParams.set('action','TEMPLATE'); u.searchParams.set('src',COUPANG_CALENDAR_ID);
  u.searchParams.set('text',d.title); u.searchParams.set('dates',`${date}T${start}00/${date}T${end}00`);
  if(details) u.searchParams.set('details',details); if(d.camp) u.searchParams.set('location',d.camp);
  window.open(u.toString(),'_blank','noopener');
  toast('Google Calendar에서 세부 내용을 확인한 뒤 저장해 주세요.');
}

async function loadReports(){
  const d = await api('/api/reports');
  reports = d.reports||[];
  $('rtabs').innerHTML = reports.map((r,i)=>`<button class="${i===curReport?'on':''}" onclick="pickReport(${i})">${r.kind}</button>`).join('') || '';
  pickReport(Math.min(curReport, Math.max(0,reports.length-1)));
}
function pickReport(i){
  curReport=i;
  document.querySelectorAll('#rtabs button').forEach((b,j)=>b.classList.toggle('on',j===i));
  $('mdbox').innerHTML = reports[i] ? mdRender(reports[i].text) : '리포트가 아직 없습니다';
}

/* ── 보고 기준일 → 엑셀 00_대시보드 B3·B4 ── */
function initDates(){
  const t = new Date();
  const iso = d => d.toISOString().slice(0,10);
  $('d_report').value = iso(t);
  const prev = new Date(t);                       // 집계기준일 = 전 영업일
  do { prev.setDate(prev.getDate()-1); } while([0,6].includes(prev.getDay()));
  $('d_base').value = iso(prev);
}
async function setDates(){
  const 보고일 = $('d_report').value, 집계기준일 = $('d_base').value;
  if(!보고일 && !집계기준일){ alert('날짜를 선택하세요'); return; }
  if(!confirm(`엑셀 00_대시보드에 반영합니다 (새 버전 생성):\n보고일 ${보고일||'-'} · 집계기준일 ${집계기준일||'-'}\n\n진행할까요?`)) return;
  const r = await api('/api/set_dates',{method:'POST',body:JSON.stringify({보고일, 집계기준일})});
  if(r.ok){ show('run'); if(!polling) polling = setInterval(pollLog, 1200); pollLog(); }
  else alert(r.error||'실패');
}

/* 보고일·집계기준일을 **보고 화면에서 바로** 고친다.
   기존에는 대시보드 맨 아래 카드에만 있어서, 보고서를 보다가 날짜가 틀린 걸 발견하면
   탭을 옮겨 찾아가야 했다. 값은 엑셀 00_대시보드 B3·B4가 진실이므로 거기에 쓴다. */
function openRptDates(){
  const meta = (execRep && execRep.meta) || {};
  const cur = normDate(meta['보고일']) || todayISO();
  const base = normDate(meta['집계기준일']) || '';
  openPane(`<h2>보고 기준일 변경</h2>
    <div class="sub">엑셀 00_대시보드 B3·B4에 기록됩니다 (새 버전 생성)</div>
    <div style="display:flex;flex-direction:column;gap:12px;margin-top:16px">
      <label style="font-size:12.5px;color:var(--ink-3);font-weight:700">보고일
        <input type="date" id="rd_report" value="${cur}" style="width:100%;padding:12px;margin-top:5px;
          border:1.5px solid var(--line);border-radius:10px;font-size:16px;font-family:inherit;
          background:var(--card);color:var(--ink-1)"></label>
      <label style="font-size:12.5px;color:var(--ink-3);font-weight:700">집계기준일
        <input type="date" id="rd_base" value="${base}" style="width:100%;padding:12px;margin-top:5px;
          border:1.5px solid var(--line);border-radius:10px;font-size:16px;font-family:inherit;
          background:var(--card);color:var(--ink-1)"></label>
      <button onclick="saveRptDates()" style="width:100%;padding:13px;background:var(--brand);color:#fff;
        border:0;border-radius:11px;font-weight:800;font-size:15px;font-family:inherit;cursor:pointer">
        저장하고 엑셀에 반영</button>
      <button onclick="previewRptDate(false)" style="width:100%;padding:12px;background:#EEF2FF;
        color:#1B41BC;border:1.5px solid #C7D2FE;border-radius:11px;font-weight:800;
        font-size:14px;font-family:inherit;cursor:pointer">선택 날짜 보고 미리보기</button>
      <button onclick="previewRptDate(true)" style="width:100%;padding:12px;background:#177245;
        color:#fff;border:0;border-radius:11px;font-weight:800;font-size:14px;
        font-family:inherit;cursor:pointer">선택 날짜 보고 바로 캡처</button>
      <button onclick="rptDatesToday()" style="width:100%;padding:11px;background:transparent;
        color:var(--brand);border:1.5px solid var(--line);border-radius:11px;font-weight:700;
        font-size:13.5px;font-family:inherit;cursor:pointer">오늘 · 전 영업일로 맞추기</button>
    </div>`);
}

/* 가장 흔한 입력을 한 번에 — 보고일=오늘, 집계기준일=직전 영업일(주말 건너뜀) */
function rptDatesToday(){
  const t = new Date(), iso = d => d.toISOString().slice(0,10);
  const prev = new Date(t);
  do { prev.setDate(prev.getDate()-1); } while([0,6].includes(prev.getDay()));
  $('rd_report').value = iso(t);
  $('rd_base').value = iso(prev);
}

async function previewRptDate(andCapture){
  const day=($('rd_base')&&$('rd_base').value)||($('rd_report')&&$('rd_report').value)||'';
  if(!/^2026-\d{2}-\d{2}$/.test(day)){alert('2026년 날짜를 선택하세요');return;}
  try{
    const [rep,b]=await Promise.all([
      api('/api/exec_report?date='+encodeURIComponent(day)),
      api('/api/brief?date='+encodeURIComponent(day))
    ]);
    execRep=filterExec(rep||{});
    BRIEF=b&&b.ok!==false?b:null;
    REPORT_PREVIEW_DATE=day;
    closeSheetAll();
    renderDaily();
    toast(`${day} 기준 보고를 불러왔습니다`);
    if(andCapture) await captureReport();
  }catch(e){alert('선택 날짜 보고를 만들지 못했습니다: '+e);}
}

async function saveRptDates(){
  const 보고일 = $('rd_report').value, 집계기준일 = $('rd_base').value;
  if(!보고일 && !집계기준일){ alert('날짜를 선택하세요'); return; }
  // 집계기준일이 보고일보다 뒤면 십중팔구 잘못 고른 것이다 — 그대로 쓰면 보고서가 어긋난다
  if(보고일 && 집계기준일 && 집계기준일 > 보고일 &&
     !confirm(`집계기준일(${집계기준일})이 보고일(${보고일})보다 뒤입니다.
그대로 진행할까요?`)) return;
  if(!confirm(`엑셀 00_대시보드에 반영합니다 (새 버전 생성):
보고일 ${보고일||'-'} · 집계기준일 ${집계기준일||'-'}

진행할까요?`)) return;
  const r = await api('/api/set_dates',{method:'POST',body:JSON.stringify({보고일, 집계기준일})});
  if(!r.ok){ alert(r.error||'실패'); return; }
  closeSheetAll();      // 곧 [실행] 탭으로 넘어가므로 시트를 남겨 두면 안 된다
  // 대시보드 카드도 같은 값으로 맞춰 둔다(두 곳이 다르면 헷갈린다)
  try{ if($('d_report')) $('d_report').value = 보고일; if($('d_base')) $('d_base').value = 집계기준일; }catch(e){}
  toast('반영 중 — 엑셀 새 버전이 만들어집니다');
  show('run'); if(!polling) polling = setInterval(pollLog, 1200); pollLog();
}

/* ── 대표 보고서: 생성 → 캡처(PNG) → 공유 ── */
let brandLogo = '';      // webapp/brand/ 에 넣은 고객사 CI 파일명(없으면 기본 마크)
let erpDocs = {rows:[],months:{},kinds:{},total:0};   // 25_ERP매출서류 (이카운트 계산서 원본)
let srcStats = {};                     // /api/status 의 sources (4원천 실집계)
let _rptData = null;                   // 보고서 데이터(HTML·Canvas 공용)
let execRep = {};                      // 01_대표보고 시트 원본(엑셀 수식 집계 결과)
let repData = {};                      // 유수비 대표 통화 기준 예외·진도·거래서류 요약

function representativeRows(kind){
  const as=(repData&&repData['돌발AS'])||{}, pm=(repData&&repData['정기점검'])||{};
  const docs=(repData&&repData['거래명세서'])||{};
  if(kind==='as-backlog') return as['미완료목록']||[];
  if(kind==='as-d2') return (as['미완료목록']||[]).filter(r=>(+r['경과일']||0)>2);
  if(kind==='paperwork') return as['서류미정리목록']||[];
  if(kind==='pm-gap') return (pm['목록']||[]).filter(r=>!/완료/.test(String(r.상태||'')));
  if(kind==='statement-unissued') return docs['미발행목록']||[];
  return [];
}
function openRepresentativeList(kind){
  if(kind==='policy'){
    const policies=(repData&&repData['업무기준확인필요'])||[];
    openPane(`<h2>업무기준 확인 필요 <span class="chip c-warn">${policies.length}건</span></h2>
      <div class="sub">항목을 누르고 확정 기준을 입력한 뒤 저장하면 보고·확인필요 화면 전체에 즉시 반영됩니다.</div>
      <div class="slist">${policies.map((p,i)=>`<div class="srow" role="button" tabindex="0"
        onclick="openPolicyEdit(${i})"
        onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();openPolicyEdit(${i})}">
        <div class="top"><span class="prjno">기준 ${i+1}</span>${chip(p.상태||'확인 필요')}</div>
        <div class="metric-issue">${esc2(p.기준||'')}</div>
        <div class="meta"><span>눌러 기준 입력·저장</span></div></div>`).join('')||
        '<div class="card">확인 대기 기준이 없습니다.</div>'}</div>`);
    return;
  }
  const names={
    'as-backlog':'돌발 AS 전산상 미완료','as-d2':'돌발 AS D+2 초과',
    paperwork:'현장완료 · 서류 미정리', 'pm-gap':'분기 정기점검 미실행',
    'statement-unissued':'거래명세서 미발행'
  };
  const basis={
    'as-backlog':'접수일이 있는 2026년 돌발 AS 중 완료일·완료상태가 확인되지 않은 건',
    'as-d2':'집계기준일 기준 전산상 미완료가 2일을 초과한 돌발 AS',
    paperwork:'현장 완료는 확인됐으나 사진·완료보고서·ERP 중 하나 이상이 비어 있는 건',
    'pm-gap':'현재 분기 실제 점검대상 중 완료일·완료상태가 확인되지 않은 건',
    'statement-unissued':'비용구분이 유상으로 확인된 정산 중 거래명세서 번호·발행상태가 없는 건'
  };
  const rows=representativeRows(kind), label=names[kind]||'대표 예외 목록';
  window._briefMetric={label,data:{rows,count:rows.length,kind:'representative',basis:basis[kind]||''}};
  openExecMetric(label);
}
function openPolicyEdit(index){
  const p=((repData&&repData['업무기준확인필요'])||[])[index]; if(!p)return;
  window._policyEditing=p;
  openPane(`<h2>업무기준 입력·저장</h2>
    <div class="sub">저장 후 이 기준은 확인 대기에서 제외되고 시스템 공통 확정 기준으로 사용됩니다.</div>
    <div class="metric-basis"><b>확인 항목</b> ${esc2(p.기준||'')}</div>
    <label style="display:block;margin-top:14px;font-size:12px;font-weight:800;color:var(--ink-2)">
      확정 기준 내용</label>
    <textarea id="policyValue" rows="7" placeholder="예: 거래명세서는 같은 캠프·같은 월 완료분을 월말 기준으로 묶는다."
      style="width:100%;margin-top:7px;border:1px solid var(--line);border-radius:12px;padding:13px;
      font:inherit;font-size:16px;line-height:1.6;resize:vertical">${esc2(p.확정내용||'')}</textarea>
    <label style="display:block;margin-top:12px;font-size:12px;font-weight:800;color:var(--ink-2)">
      확인자</label>
    <input id="policyOwner" value="유현민" style="width:100%;margin-top:7px;border:1px solid var(--line);
      border-radius:11px;padding:12px;font:inherit;font-size:16px">
    <div class="actions sticky">
      <button class="abtn primary" onclick="savePolicy()"><span class="em">💾</span>저장·전체 반영</button>
      <button class="abtn" onclick="openRepresentativeList('policy')">취소</button>
    </div>`);
}
async function savePolicy(){
  const p=window._policyEditing, value=String(($('policyValue')&&$('policyValue').value)||'').trim();
  const owner=String(($('policyOwner')&&$('policyOwner').value)||'유현민').trim();
  if(!p||!value){alert('확정 기준 내용을 입력하세요');return;}
  try{
    await api('/api/policy',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({기준:p.기준,확정내용:value,저장자:owner})});
    repData=await api('/api/v1/reports/daily/exceptions');
    renderRepresentative();renderCheckHub();closeSheetAll();
    toast('업무기준을 저장하고 전체 화면에 반영했습니다');
  }catch(e){alert('업무기준 저장 실패: '+e);}
}
function renderRepresentative(){
  if(!$('repKpis')) return;
  const as=(repData&&repData['돌발AS'])||{}, pm=(repData&&repData['정기점검'])||{};
  const docs=(repData&&repData['거래명세서'])||{}, policies=(repData&&repData['업무기준확인필요'])||[];
  if(!repData||!Object.keys(repData).length){
    $('repSignal').className='signal 황색'; $('repSignal').textContent='새로고침 필요';
    $('repSummary').textContent='대표 예외보고 API를 불러오지 못했습니다. 서버가 갱신된 뒤 새로고침하면 자동 표시됩니다.';
    $('repKpis').innerHTML=''; $('repDocs').innerHTML=''; return;
  }
  const signal=String(pm['신호']||'녹색');
  $('repSignal').className='signal '+signal; $('repSignal').textContent=`${signal} 신호`;
  $('repSummary').textContent=repData['한줄종합보고']||'현재 원장 기준으로 예외를 집계했습니다.';
  const metric=(kind,label,value,sub,cls)=>`<button class="rep-metric ${cls||''}" onclick="openRepresentativeList(${esc4(kind)})">
    <div class="rv">${esc2(value)}</div><div class="rl">${esc2(label)}</div><div class="rs">${esc2(sub)}</div></button>`;
  const gap=Number(pm['계획대비']||0);
  $('repKpis').innerHTML=
    metric('as-backlog','AS 전산상 미완료',as['전산상미완료']||0,'접수 후 완료 미확인',as['전산상미완료']?'danger':'ok')+
    metric('as-d2','AS D+2 초과',as['D+2초과']||0,'대표 지속보고 대상',as['D+2초과']?'danger':'ok')+
    metric('paperwork','현장완료·서류미정리',as['현장완료서류미정리']||0,'사진·보고서·ERP',as['현장완료서류미정리']?'warn':'ok')+
    metric('pm-gap','정기점검 계획대비',(gap>=0?'+':'')+gap+'건',
      `목표누계 ${pm['목표누계']||0} · 실제 ${pm['실제완료']||0}`,gap<0?'danger':'ok')+
    metric('statement-unissued','거래명세서 미발행',(docs['미발행목록']||[]).length,
      '유상 발행대상 기준',(docs['미발행목록']||[]).length?'warn':'ok')+
    metric('policy','업무기준 확인',policies.length,'승인 전 자동화 제외',policies.length?'warn':'ok');
  const docRows=docs['업무유형별']||[];
  $('repDocs').innerHTML=docRows.map(d=>`<div class="rep-doc">
    <b>${esc2(d.업무유형||'')}</b><span>대상 ${d.발행대상||0}</span><span>발행 ${d.발행완료||0}</span>
    <span>미발행 ${d.미발행||0}</span><span>${d.발행률==null?'대상 없음':d.발행률+'%'}</span></div>`).join('');
}

/* 01_대표보고 섹션을 보고서 HTML로 — 엑셀 대표보고와 동일한 항목·수치 */


/* ══ 문제 코드 정확히 풀어쓰기 ═════════════════════════════════
   관리대장 수식이 만드는 축약 문구를 '무엇이 비었고 무엇을 해야 하는지'로 바꾼다.
   (02_돌발AS접수 AN열 수식에서 나오는 코드들) */
const CODE = {
  '관리자 확인 필요': ['작업완료인데 <b>관리자검증상태</b>가 비어 있음',
                       '관리자가 일치 / 추가작업발생 / 작업내용누락 / 확인필요 중 하나를 지정해야 합니다'],
  '중복 가능성':      ['같은 <b>프로젝트NO·캠프·접수경로</b>로 2건 이상 등록됨',
                       '중복판정(선택) 열에서 중복 아님 / 재접수 / 추가작업 / 다른 호기 등으로 확정하세요'],
  '담당자 미배정':    ['담당기사가 지정되지 않음', '기사를 배정해야 작업이 진행됩니다'],
  '방문 일정 미확정':  ['방문예정일이 비어 있음', '기사와 일정을 잡아 입력하세요'],
  '작업완료일 누락':  ['상태는 작업완료인데 완료일이 비어 있음',
                       '완료일이 없으면 월별 실적·정산에서 빠집니다'],
  '사진 미첨부':      ['작업완료인데 사진등록이 안 됨', '밴드 게시글의 사진을 확인하세요'],
  '동영상 미첨부':    ['작업완료인데 동영상등록이 안 됨', '필요한 건인지 확인 후 등록'],
  '완료보고서 미첨부': ['작업완료인데 완료보고서가 없음', '보고서를 받아 등록하세요'],
  '비용 구분 미확정':  ['유상·무상·보험 구분이 정해지지 않음', '구분이 없으면 청구 여부를 판단할 수 없습니다'],
  'ERP 미등록':       ['작업완료인데 이카운트에 전표가 등록되지 않음', '전표 등록 후 상태를 갱신하세요'],
  '재방문 여부 미확정': ['재방문 필요 여부가 비어 있음', '한 번에 끝났는지 확인해 표시하세요'],
  '접수 외 추가 작업 미반영': ['현장에서 추가 작업이 있었는데 반영되지 않음', '추가작업 내용을 등록하세요'],
  '실제 작업내용 확인 필요': ['관리자검증상태가 확인필요로 지정됨', '실제 작업 내용을 확인해 주세요'],
  '작업내용 누락':    ['작업 내용이 기록되지 않음', '기사에게 작업 내용을 받아 입력하세요'],
  '완료예정일 경과':  ['완료예정일이 지났는데 아직 작업완료가 아님', '일정을 다시 잡거나 상태를 갱신하세요'],
  '접수일자 비어 있음': ['언제 접수된 건인지 날짜가 없음 — 밴드·카톡 어디에도 근거를 못 찾음',
                       '밴드 게시글이나 카톡 보고에서 실제 날짜를 찾아 <b>02_돌발AS접수</b>의 접수일자 칸에 입력하세요. 날짜가 없으면 월별 실적에서 빠집니다'],
  '점검예정일 비어 있음': ['언제 점검할 건인지 날짜가 없음',
                       '밴드 게시글이나 점검 일정표에서 날짜를 찾아 <b>04_정기점검</b>의 점검예정일 칸에 입력하세요'],
  '캠프명 비어 있음': ['어느 현장 작업인지 캠프명이 비어 있음 — 앱 카드에 <b>캠프 미상</b>으로 표시됩니다',
                       '밴드에서 그 프로젝트NO로 검색해 어느 캠프인지 확인한 뒤 <b>02_돌발AS접수</b>·<b>04_정기점검</b>의 캠프명 칸에 입력하세요'],
  '원장 미등록':      ['밴드에 올라온 명세서·계산서 사진에는 있는데 관리대장 어디에도 없는 프로젝트NO',
                       '견적·판매전표처럼 AS·점검이 아닌 업무일 수 있습니다. 밴드 해당 게시글을 열어 캠프·업무유형·금액을 확인한 뒤 알맞은 시트(02·04·13_PO발주관리)에 등록하세요'],
};

/* 문제 문자열(쉼표로 이어진 코드들) → 코드별 설명 목록 */
function codeHtml(text){
  const t = String(text||'');
  const hit = Object.keys(CODE).filter(k=>t.includes(k));
  if(!hit.length) return '';
  return `<ul class="codes">${hit.map(k=>{
    const [what, todo] = CODE[k];
    return `<li><b>${k}</b> — ${what}<div class="ct">${todo}</div></li>`;
  }).join('')}</ul>`;
}

/* TOP5 줄의 업무ID가 지금도 실제로 존재하는지 확인.
   엑셀이 재계산되지 않으면 이미 정리된 건이 계속 남아 보인다(실제로 겪음). */
function liveIds(){
  const s = new Set();
  (works.as||[]).forEach(r=>{ if(r.접수ID) s.add(String(r.접수ID).trim()); });
  (works.pm||[]).forEach(r=>{ if(r.점검ID) s.add(String(r.점검ID).trim()); });
  (settleRows||[]).forEach(r=>{ if(r.정산ID) s.add(String(r.정산ID).trim()); });
  return s;
}
/* ID로 원장 행을 찾아 **날짜와 대표 프로젝트NO**를 붙인다.
   보고서 줄에는 사내 ID(AS-2512-005)만 있어서, 대표가 "언제 건이냐 / 어느 프로젝트냐"를
   물으면 바로 답을 못 했다. 프로젝트NO는 apply_rep_no 가 채운 **대표번호**를 쓴다. */
function rowById(id){
  const all = (works.as||[]).concat(works.pm||[], settleRows||[]);
  return all.find(r => [r.접수ID, r.점검ID, r.정산ID].some(x => String(x||'') === id));
}
function idMeta(id){
  const r = rowById(id);
  if(!r) return '';
  const d = r.접수일자 || r.점검예정일 || r.완료일 || r.작업완료일 || r.실제점검일 || '';
  const prj = projectNoOf(r);
  const bits = [];
  if(d)   bits.push(String(d).slice(0,10));
  if(prj) bits.push(prj);
  return bits.length
    ? '<span style="color:var(--ink-3);font-size:11.5px;margin-left:6px">· ' + bits.join(' · ') + '</span>'
    : '';
}

/* ★ 사용자 지시(2026-07-28): "2025년도 건은 빼고, 기준일 ±10일 것만."
   TOP5는 07_불일치누락현황의 맨 위 5행을 그대로 가져오는데, 그 시트가 옛 건부터 쌓여
   2025-12 건이 먼저 나온다. 지금 손쓸 수 있는 건이 아니라 화면만 어지럽힌다.
   ※ 엑셀 시트를 고치지 않고 화면에서 거른다 — 07시트는 자동수집 수식이라 건드리면
     다른 집계가 같이 흔들린다. */
const TOP_DAYS = 10;
function topInRange(id){
  const r = typeof rowById === 'function' ? rowById(id) : null;
  if(!r) return false;                      // 2026년 행으로 확인되지 않으면 표시하지 않는다
  const d = String(r.접수일자||r.점검예정일||r.완료일||r.작업완료일||r.실제점검일||'').slice(0,10);
  if(!/^\d{4}-\d{2}-\d{2}$/.test(d)) return true;
  const base = baseDate();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(String(base||''))) return true;
  const diff = Math.abs(Date.parse(d+'T00:00:00') - Date.parse(base+'T00:00:00')) / 864e5;
  return diff <= TOP_DAYS;
}
function topLine(t){
  const txt = String(t||'');
  const m = txt.match(/\b((?:AS|PM|JS)-\d{4}-\d{3})\b/);
  if(/\b(?:AS|PM|JS)-25\d{2}-\d{3}\b|\bUJ25\d{5}\b|(^|[^\d])2025[-./]|2025년|(^|[^\d])25(?:년도|년)/.test(txt)) return '';
  if(m && !topInRange(m[1])) return '';     // 기준일에서 멀리 떨어진 건은 안 싣는다
  const live = liveIds();
  const gone = m && live.size && !live.has(m[1]);
  const body = gone
    ? `<span class="gone">${txt}</span>
       <div class="ct warn">이 건은 지금 관리대장에 없습니다 — 이미 정리(중복 통합·삭제)된 건이
          엑셀 재계산 전이라 보고서에 남아 있는 것입니다. 엑셀을 한 번 열었다 저장하면 사라집니다.</div>`
    : txt;
  return `<li>${body}${m?idMeta(m[1]):''}${gone?'':codeHtml(txt)}</li>`;
}

/* ── 대표에게 읽어 드릴 브리핑 문장 ──
   출처는 daily_brief 하나뿐이다(/api/brief). 화면·PC 리포트·폰 사본이 같은 문장을 써야
   "화면은 이런데 보고는 저렇다"가 안 생긴다. */
let BRIEF = null, BRIEF_LOADING = false, BRIEF_RETRY = null, REPORT_PREVIEW_DATE = '';
async function loadBrief(day=''){
  if(BRIEF_LOADING) return;
  BRIEF_LOADING = true;
  try{
    const b = await api('/api/brief'+(day?`?date=${encodeURIComponent(day)}`:''));
    BRIEF = b && b.ok!==false ? b : null;
    REPORT_PREVIEW_DATE = day || '';
    if(BRIEF_RETRY){ clearTimeout(BRIEF_RETRY); BRIEF_RETRY = null; }
  }catch(e){
    BRIEF = null;
    // 첫 구동 때 Z: 원장/현장일지 사전 준비가 터널 제한시간보다 길어져도, 서버 캐시가
    // 완성된 뒤 자동 재요청하여 저장 이미지가 빈 상태로 굳지 않게 한다.
    if(!BRIEF_RETRY) BRIEF_RETRY = setTimeout(()=>{ BRIEF_RETRY=null; loadBrief(); }, 5000);
  }finally{
    BRIEF_LOADING = false;
  }
  try{ renderDaily(); }catch(e){}
}

/* ── 읽어 드릴 내용 ──────────────────────────────────────────────────────────
   예전에는 서버가 만든 문장을 줄 단위로 그대로 흘려 놨더니 **어디서 어디까지가
   한 건인지** 안 보였다(사용자 지적 2026-07-28). 같은 데이터를 카드로 끊어
   그린다 — 날짜·번호·캠프는 한 줄에, 내용은 그 아래, 빠진 항목은 붉게.
   복사 버튼은 그대로 서버 문장을 쓴다(대표께 그대로 보내는 용도라 형식이 중요). */
const _e = s => String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;');
const _md = s => { s=String(s||''); return s.length>=10 ? s.slice(5) : s; };

function bTag(x){
  const paid = /유상/.test(x.비용||'') ? '<span class="btag paid">유상</span>'
             : (x.비용 ? `<span class="btag">${_e(x.비용)}</span>` : '');
  const who = x.담당기사 ? `<span class="bm">${_e(x.담당기사)}</span>`
                         : '<span class="bm none">기사 미배정</span>';
  return `<div class="bc1"><span class="bd">${_e(_md(x.일자))}</span>`
       + `<span class="bp">${_e(projectNoOf(x)||'프로젝트 미확정')}</span>`
       + `<b>${_e(x.캠프명||'캠프 미기입')}</b>${who}${paid}</div>`;
}
function bSpan(x){
  const a=x.접수일, b=x.일자;
  if(!a||!b||a===b) return '';
  const n=Math.round((Date.parse(b+'T00:00:00')-Date.parse(a+'T00:00:00'))/864e5);
  return n>0 ? `<span class="bage">접수 ${_e(_md(a))} · ${n}일 만</span>` : '';
}
function bRow(lab, val, missNote){
  return val ? `<div class="bc2"><i>${lab}</i><span>${_e(val)}</span></div>`
             : `<div class="bc2"><i>${lab}</i><span class="bmiss">★ ${missNote}</span></div>`;
}
function bCard(x, kind){
  if(kind==='new') return `<div class="bcard">${bTag(x)}${bRow('내용', x.왜, '접수내용 미기입')}</div>`;
  if(kind==='done') return `<div class="bcard">${bTag(x)}${bSpan(x)}`
      + bRow('왜', x.왜, '접수내용 미기입')
      + bRow('작업', x.무엇, '미기입 — 기사에게 확인해 03_현장작업실적에 입력 필요')
      + (x.추가작업 ? bRow('추가', x.추가작업, '') : '') + `</div>`;
  if(kind==='activity') return `<div class="bcard">${bTag(x)}${bSpan(x)}`
      + bRow('요청', x.왜, '신청내용 미기입')
      + bRow('처리', x.무엇, '처리내용 미기입')
      + (x.게시자 ? bRow('게시', x.게시자, '') : '') + `</div>`;
  if(kind==='pm-plan') return `<div class="bcard">${bTag(x)}`
      + bRow('상태', x.상태, '상태 미기입')
      + bRow('점검', x.왜, '점검내용 미기입')
      + (x.실행일 ? bRow('실행', x.실행일, '') : '') + `</div>`;
  if(kind==='pm-done') return `<div class="bcard">${bTag(x)}`
      + (x.예정일 ? bRow('예정', x.예정일, '') : '')
      + bRow('점검', x.왜, '점검내용 미기입')
      + (x.무엇 ? bRow('작업', x.무엇, '') : '') + `</div>`;
  return `<div class="bcard slim">${bTag(x)}</div>`;
}

function briefBlock(){
  if(!BRIEF || !BRIEF.ok) return '';
  const a = BRIEF['돌발AS']||{}, p = BRIEF['정기점검']||{};
  const wl = BRIEF['일지대조']||{}, wla = wl['돌발AS']||{};
  const newL = BRIEF['신규목록']||[], doneL = (BRIEF['완료내역']||[]).filter(x=>x.구분==='돌발AS');
  const activityL = BRIEF['당일처리목록']||[];
  const pmPlanL = BRIEF['점검예정목록']||[], pmDoneL = BRIEF['점검실행목록']||[];
  const blank = BRIEF['내용미기입']||[], stale = BRIEF['완료일미기입목록']||[];
  const d = BRIEF['기준일']||'';
  let h = '';

  /* 돌발 AS */
  h += `<div class="bsec"><div class="bst"><b>돌발 AS</b>
    <span class="bpill new">신규 ${a.신규접수||0}</span>
    <span class="bpill done">완료 ${a.완료||0}</span>
    ${activityL.length?`<span class="bpill done">업무 처리 ${activityL.length}</span>`:''}
    <span class="bpill warn">미처리 ${a.미처리||0}<em>최근 30일</em></span></div>`;
  if(Object.keys(wla).length){
    const reasons=wla.미처리사유||[];
    h += `<div class="bglab">현장 일지 대조 · 발생 ${wla.발생||0}건 · 처리완료 ${wla.처리완료||0}건 · 미처리 ${wla.미처리||0}건 · 취소 ${wla.취소||0}건</div>
      <div class="pmstats">
        <button class="done" onclick="openWorkLogBrief('done')">처리완료<b>${wla.처리완료||0}건</b></button>
        <button class="warn" onclick="openWorkLogBrief('open')">미처리<b>${wla.미처리||0}건</b></button>
        <button onclick="openWorkLogBrief('all')">전체 발생<b>${wla.발생||0}건</b></button>
      </div>`;
    if(reasons.length) h += `<div class="bneed actionable" role="button" tabindex="0" onclick="openWorkLogBrief('open')"
      onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();openWorkLogBrief('open')}">
      <div class="bnh">미처리 사유 <b>${reasons.map(x=>`${_e(x.사유)} ${x.건수}건`).join(' · ')}</b></div>
      <div class="bnd">눌러서 프로젝트번호·캠프·원문 사유 확인 · 이미지 저장·복사</div></div>`;
    if(wla.처리완료일확인) h += `<div class="bnd">※ 처리완료 ${wla.처리완료일확인}건은 원본에 완료일이 없어 임의 입력하지 않았습니다.</div>`;
  }
  if(newL.length) h += `<div class="bglab new">새로 접수 ${newL.length}건 · 접수일 ${_md(d)}</div>`
                     + newL.slice(0,6).map(x=>bCard(x,'new')).join('');
  if(doneL.length) h += `<div class="bglab done">완료 ${doneL.length}건 · 무엇 때문에 갔고 무슨 작업을 했는지</div>`
                     + doneL.slice(0,8).map(x=>bCard(x,'done')).join('');
  if(activityL.length) h += `<div class="bglab done">당일 업무 처리 ${activityL.length}건 · 현장 AS 완료와 별도</div>`
                         + activityL.map(x=>bCard(x,'activity')).join('');
  if(!newL.length && !doneL.length && !activityL.length)
    h += `<div class="bnone">그날 새로 접수되거나 완료·처리된 건이 없습니다</div>`;
  h += `</div>`;

  /* 정기점검 — 진행률은 막대로 보여 준다(숫자만 있으면 감이 안 온다).
     ★ 사용자 지시(2026-07-29): 기간을 월 단위로 골라 볼 수 있어야 한다(7월만, 1~6월 등).
       분기는 **처음 값**일 뿐이고, 고르면 그 기간으로 다시 센다(pmStats). */
  const ps = pmStats();
  const pct = ps.진행률;
  h += `<div class="bsec"><div class="bst"><b>정기점검</b>
    <button class="bpill actionable" onclick="openPmBrief('day-plan')">예정 ${p.예정||0}</button>
    <button class="bpill done actionable" onclick="openPmBrief('day-done')">실행 ${p.완료||0}</button></div>
    ${pmPlanL.length?`<div class="bglab">기준일 예정 ${pmPlanL.length}건 · 눌러서 전체 목록 확인</div>`
      +pmPlanL.slice(0,4).map(x=>bCard(x,'pm-plan')).join(''):''}
    ${pmDoneL.length?`<div class="bglab done">기준일 실행 ${pmDoneL.length}건</div>`
      +pmDoneL.slice(0,4).map(x=>bCard(x,'pm-done')).join(''):''}
    ${!pmPlanL.length&&!pmDoneL.length?`<div class="bnone">기준일에 예정되거나 실행된 정기점검이 없습니다</div>`:''}
    <div class="bbar"><div class="bbt"><span>${_e(APP_YEAR)}년 ${_e(ps.라벨)} 진행률</span>
      <b>${pct}% <em>${ps.실행}/${ps.예정}건</em></b></div>
      <div class="btrack"><div class="bfill${pct>=60?'':' low'}" style="width:${Math.min(pct,100)}%"></div></div>
      <div class="pmrange">
        <select onchange="setPmRange(this.value, $('pmTo').value)" id="pmFrom">${pmMonthOpts(ps.from)}</select>
        <span>~</span>
        <select onchange="setPmRange($('pmFrom').value, this.value)" id="pmTo">${pmMonthOpts(ps.to)}</select>
        ${[['이번 분기', ps.q0, ps.q1], ['상반기','01','06'], ['하반기','07','12'], ['연간','01','12']]
          .map(([t,a,b])=>`<button class="pbtn${ps.from===a&&ps.to===b?' on':''}"
             onclick="setPmRange('${a}','${b}')">${t}</button>`).join('')}
      </div>
      <div class="bnote">${pct>=60?`특별한 문제 없으면 ${ps.끝월}까지 마무리 가능합니다.`
                                 :'진행률이 낮아 일정 관리가 필요합니다.'}</div>
      <div class="pmstats">
        <button onclick="openPmBrief('quarter-all')">${_e(ps.라벨)} 예정<b>${ps.예정}건</b></button>
        <button class="done" onclick="openPmBrief('quarter-done')">${_e(ps.라벨)} 실행<b>${ps.실행}건</b></button>
        <button class="warn" onclick="openPmBrief('quarter-pending')">${_e(ps.라벨)} 미실행<b>${ps.미실행}건</b></button>
      </div></div></div>`;

  /* 확인이 필요한 것 — 대표가 물어볼 항목만 모아 둔다 */
  const staleN = a.완료일미기입||0;
  if(staleN || blank.length){
    h += `<div class="bsec need"><div class="bst"><b>확인이 필요합니다</b></div>`;
    if(staleN){
      const ds = stale.map(x=>x.일자).filter(Boolean).sort();
      h += `<div class="bneed actionable" role="button" tabindex="0" onclick="openStaleBrief()"
        onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();openStaleBrief()}">
        <div class="bnh">완료일이 안 적힌 오래된 건 <b>${staleN}건</b></div>
        <div class="bnd">접수 후 30일 넘음${ds.length?` · 접수 ${ds[0]} ~ ${ds[ds.length-1]}`:''}</div>
        <div class="bnd">실제로는 끝났을 가능성이 큽니다 — 완료일만 채우면 정리됩니다.</div>
        <div class="bnd actionhint">눌러서 ${stale.length}건 목록 보기 · 이미지 저장·복사</div></div>`;
    }
    if(blank.length){
      h += `<div class="bneed"><div class="bnh">작업 내용이 안 적힌 완료건 <b>${blank.length}건</b></div>
        <div class="bnd">기사에게 확인해 03_현장작업실적에 입력해야 합니다</div>`
        + blank.slice(0,5).map(x=>bCard(x,'slim')).join('') + `</div>`;
    }
    h += `</div>`;
  }

  return `<div class="ebrief">
    <div class="bhead"><span>읽어 드릴 내용</span>
      <em>숫자가 아니라 무슨 일이 있었는지</em>
      <button onclick="copyBrief()">📋 문장 복사</button></div>
    ${h}</div>`;
}

async function copyBrief(){
  const t = (BRIEF && BRIEF.text) || '';
  try{ await navigator.clipboard.writeText(t); }
  catch(e){
    const a=document.createElement('textarea'); a.value=t; document.body.appendChild(a);
    a.select(); try{ document.execCommand('copy'); }catch(_){} a.remove();
  }
  toast('보고 문장을 복사했습니다');
}

/* 현장 일지 원본(정기점검·돌발AS 일지)의 처리/미처리 목록.
   일지 행에는 원장 ID가 없을 수 있어 프로젝트NO로 상세를 연다. */
function openWorkLogBrief(scope){
  const wla=((BRIEF&&BRIEF['일지대조'])||{})['돌발AS']||{};
  const open=wla.미처리목록||[], done=wla.처리완료목록||[], cancel=wla.취소목록||[];
  const source=scope==='open'?open:scope==='done'?done:[...done,...open,...cancel];
  const label=scope==='open'?'돌발AS 일지 미처리':scope==='done'?'돌발AS 일지 처리완료':'돌발AS 일지 전체 발생';
  const rows=source.map(x=>({
    프로젝트NO:x.프로젝트NO||'', 캠프명:x.캠프명||'', 일자:x.일자||'', 담당자:x.담당자||'담당 미기입',
    문제:[x.요청내용?`요청: ${x.요청내용}`:'',x.미처리사유?`사유: ${x.미처리사유}`:'',
          x.실제조치?`조치: ${x.실제조치}`:''].filter(Boolean).join(' · '),
    상태:x.상태||x.원본상태||'', 종류:'', ID:x.프로젝트NO||''
  }));
  window._briefMetric={label,data:{rows,count:rows.length,kind:'work-log',
    basis:'정기점검·돌발AS 일지(미실시건) 원본을 프로젝트번호로 원장과 대조한 결과'}};
  openExecMetric(label);
}

/* ══ 당일 업무 실적 — 항목별 상세 ══════════════════════════════
   '신규 접수 3'이 어떤 건인지 바로 아래에 보여 준다.
   대표보고 엑셀에는 숫자만 있어서, 같은 기준일로 앱 데이터에서 직접 찾아 붙인다. */
function baseDate(){
  const m = (execRep.meta||{});
  return normDate(m['집계기준일'] || m['보고일']) || todayISO();
}
const _k = s => String(s||'').replace(/[\s·()]/g,'');

/* '(금일)'은 읽는 사람마다 다른 날을 떠올린다 — 타일은 전부 집계기준일로 세므로 그 날짜를 박는다.
   사용자 지시(2026-07-28): "금일이라고 하면 어떤 날짜인지 헷갈림." */
function dateLabel(l){
  const d = baseDate();
  if(!d) return l;
  return String(l||'').replace(/\((금일|당일|오늘)\)/g, `(${d.slice(5)})`);
}
const WD = '월화수목금토일';
function wdOf(iso){
  const t = Date.parse(iso+'T00:00:00');
  return Number.isFinite(t) ? WD[(new Date(t).getDay()+6)%7] : '';
}

/* 라벨 → 그 숫자를 만든 실제 건 목록 */
function dayRows(label){
  const d = baseDate();
  const as = works.as||[], pm = works.pm||[], st = settleRows||[];
  const on = (v)=> normDate(v) === d;
  const paid = r => String(r['유상·무상·보험']||r.비용구분||'').includes('유상');
  const M = {
    '신규접수':      ()=>[as.filter(r=>on(r.접수일자)), 'as'],
    '작업완료':      ()=>[as.filter(r=>on(r.작업완료일)), 'as'],
    '현장작업':      ()=>[[], 'as'],
    '유상발생':      ()=>[as.filter(r=>on(r.접수일자)&&paid(r)), 'as'],
    '재방문예정금일': ()=>[as.filter(r=>on(r.방문예정일)), 'as'],
    '점검완료':      ()=>[pm.filter(r=>on(r.실제점검일)), 'pm'],
    '점검예정금일':   ()=>[pm.filter(r=>on(r.점검예정일)), 'pm'],
    '이상발견':      ()=>[pm.filter(r=>on(r.실제점검일)&&String(r.이상발견여부||'').trim()&&r.이상발견여부!=='없음'), 'pm'],
    '돌발AS전환':    ()=>[pm.filter(r=>on(r.실제점검일)&&String(r.돌발AS전환여부||'').trim()&&r.돌발AS전환여부!=='없음'), 'pm'],
    '유상점검':      ()=>[pm.filter(r=>on(r.실제점검일)&&paid(r)), 'pm'],
    '거래명세서발행':  ()=>[st.filter(r=>on(r.명세서발행일)), 'settle'],
    '세금계산서발행':  ()=>[st.filter(r=>on(r.계산서발행일)), 'settle'],
    'ERP등록작업완료기준': ()=>[st.filter(r=>on(r.완료일)&&r.명세서번호), 'settle'],
    '청구진행':      ()=>[st.filter(r=>on(r.청구일)), 'settle'],
    '입금건수':      ()=>[st.filter(r=>on(r.입금일)), 'settle'],
  };
  const f = M[_k(label)];
  return f ? f() : [null, ''];
}

/* 0건일 때 '무슨 뜻인지'를 대신 보여 준다 */
const DAYNOTE = {
  '신규접수': '그날 새로 접수된 돌발AS 건',
  '작업완료': '그날 작업이 끝난 것으로 기록된 건',
  '현장작업': '03_현장작업실적 시트 기준 — 앱에는 아직 연결 안 됨',
  '유상발생': '그날 접수 중 유상(청구 대상)인 건',
  '재방문예정금일': '그날 재방문이 예정된 건',
  '점검완료': '그날 실제점검일이 기록된 정기점검',
  '점검예정금일': '그날 점검 예정으로 잡힌 건',
  '이상발견': '점검에서 이상이 발견돼 조치가 필요한 건',
  '돌발AS전환': '점검 중 발견돼 돌발AS로 넘어간 건',
  '유상점검': '그날 점검 중 유상(청구 대상)인 건',
  '거래명세서발행': '그날 거래명세서를 발행한 건',
  '세금계산서발행': '그날 세금계산서를 발행한 건',
  'ERP등록작업완료기준': '작업완료일이 그날이고 명세서번호가 있는 건',
  '청구진행': '그날 청구가 진행된 건',
  '입금건수': '그날 입금이 확인된 건',
};

function dayDetail(label, val){
  const [rows, kind] = dayRows(label);
  const note = DAYNOTE[_k(label)] || '';
  if(!rows) return note ? `<div class="gd note">${note}</div>` : '';
  // 대표보고 숫자는 엑셀에서, 목록은 앱 데이터에서 온다. 두 값이 다르면 숨기지 않고 알린다.
  const n = parseInt(String(val).replace(/[^0-9]/g,''), 10);
  const diff = Number.isFinite(n) && n !== rows.length;
  const warn = diff ? `<span class="dw">앱에서 찾은 건 ${rows.length}건 (보고 숫자 ${n})</span>` : '';
  if(!rows.length) return `<div class="gd note">${note}${note?' — ':''}해당 없음${warn}</div>`;
  const chips = rows.slice(0,3).map(r=>{
    const no = r.프로젝트NO || r.정산ID || r.접수ID || r.점검ID || '';
    const camp = String(r.캠프명||'').slice(0,12);
    return `<span class="prj" onclick="event.stopPropagation();openByPrj(${esc4(no)})">${no}</span>` +
           (camp?`<span class="dc">${camp}</span>`:'');
  }).join('');
  const more = rows.length>3 ? `<span class="more">외 ${rows.length-3}건</span>` : '';
  return `<div class="gd"><div class="chips">${chips}${more}</div>${warn}</div>`;
}

function cleanExecTitle(title){
  return String(title||'').replace(/^\s*\d+\.\s*/,'').trim();
}
function execPeriodTitle(title){
  const t=cleanExecTitle(title), bd=baseDate();
  if(/당일 금액|잔여 현황/.test(t)) return `당일 금액 · 잔여 현황 (${bd})`;
  if(/^리스크/.test(t)) return `리스크 (${APP_YEAR}-01-01 ~ ${bd})`;
  return t;
}
function dailyIssueRows(){
  const a=(BRIEF&&BRIEF['돌발AS'])||{}, wl=((BRIEF&&BRIEF['일지대조'])||{})['돌발AS']||{};
  const blank=(BRIEF&&BRIEF['내용미기입'])||[];
  return [
    {label:'돌발 AS 미처리', count:+wl.미처리||+a.미처리||0, detail:'현장 일지 기준 미실시·미완료', call:"openWorkLogBrief('open')"},
    {label:'완료일 확인 필요', count:+a.완료일미기입||0, detail:'접수 후 30일 초과·완료일 미기입', call:'openStaleBrief()'},
    {label:'작업내용 확인 필요', count:blank.length, detail:'완료됐으나 실제 작업내용 미기입', call:'openBlankWorkBrief()'}
  ];
}
function dailyIssueHtml(){
  const rows=dailyIssueRows(), total=rows.reduce((n,x)=>n+x.count,0);
  return `<div class="egroup issue${total?'':' ok'}">
    <div class="gh">이슈사항 · ${total?`${total}건 확인 필요`:'특이사항 없음'}</div>
    <div class="gb">${rows.map(x=>`<div class="gitem" role="button" tabindex="0"
      onclick="${x.call}" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();${x.call}}">
      <div class="gi"><span>${x.label}</span><b style="color:${x.count?'#B42318':'#177245'}">${x.count}건</b></div>
      <div class="gd note">${x.detail}</div></div>`).join('')}</div></div>`;
}
function openBlankWorkBrief(){
  const source=(BRIEF&&BRIEF['내용미기입'])||[];
  const rows=source.map(x=>({
    프로젝트NO:x.프로젝트NO||'', 캠프명:x.캠프명||'', 일자:x.일자||'',
    담당자:x.담당기사||'담당 미기입', 문제:'작업완료이나 실제 작업내용 미기입',
    상태:'확인 필요', 종류:x.레코드종류||'as', ID:x.레코드ID||x.프로젝트NO||''
  }));
  const label='작업내용 확인 필요';
  window._briefMetric={label,data:{rows,count:rows.length,kind:'brief-blank',
    basis:'완료 처리된 2026년 업무 중 실제 작업내용이 비어 있는 건'}};
  openExecMetric(label);
}

function execHtml(){
  const secs = (execRep.sections||[]).filter(s=>!/^(정리|요약)$/.test(cleanExecTitle(s.title)));
  const empty = x => !(x.items||[]).length && !(x.groups||[]).length && !(x.lines||[]).length;
  if(!secs.some(s=>/당일 업무 실적/.test(cleanExecTitle(s.title)))){
    secs.unshift({title:'당일 업무 실적',items:[],groups:[],lines:[],briefOnly:true});
  }
  let usedBrief = false, h = '', sectionNo=0;

  secs.forEach(s=>{
    if(empty(s) && !s.briefOnly && !/당일 업무 실적/.test(cleanExecTitle(s.title))) return;
    const raw=cleanExecTitle(s.title), tt=execPeriodTitle(raw);
    sectionNo += 1;
    h += `<div class="esec"><div class="eh"><i>${sectionNo}</i><span>${tt}</span></div>`;
    // ★ 대표 지시(2026-07-28): "숫자를 나한테 보고하라는 게 아니야."
    //   숫자 타일 위에 **읽어서 그대로 전달할 수 있는 문장**을 먼저 놓는다.
    if(/당일 업무 실적/.test(tt)){
      usedBrief = true;
      // ★ 어느 날 숫자인지부터 못 박는다. 아래 타일은 전부 '집계기준일' 하루치다.
      const bd = baseDate(), rd = normDate((execRep.meta||{})['보고일']);
      h += `<div class="dbanner">
        <b>${bd}${wdOf(bd)?`(${wdOf(bd)})`:''}</b> 하루치 실적입니다
        <span>· 아래 모든 숫자·목록이 이 날짜 기준</span>
        ${rd&&rd!==bd?`<span class="rd">보고일 ${rd}${wdOf(rd)?`(${wdOf(rd)})`:''}</span>`:''}
      </div>`;
      h += briefBlock();
    }
    if((s.groups||[]).length){
      h += `<div class="egrid">` + s.groups.map(g=>`<div class="egroup">
        <div class="gh">${g.name}</div>
        <div class="gb">${g.items.map(([l,v])=>{
          const k = helpKey(l);
          return `<div class="gitem"><div class="gi"><span${k?` class="has-help" onclick="openHelp('${k}')"`:''}>${dateLabel(l)}${k?'<i class="qm">?</i>':''}</span><b>${v||'-'}</b></div>${dayDetail(l,v)}</div>`;
        }).join('')}</div></div>`).join('');
      if(/당일 업무 실적/.test(raw)) h += dailyIssueHtml();
      h += `</div>`;
    }
    if((s.items||[]).length){
      // 3·4절 숫자는 모두 원천 행 목록으로 들어간다. 물음표는 집계 설명만 따로 연다.
      h += `<div class="ecells">` + s.items.map(([l,v])=>{
        const k = helpKey(l), metric = !!((execRep.details||{})[l]);
        const click = metric ? `openExecMetric(${esc4(l)})` : k ? `openHelp(${esc4(k)})` : '';
        return `<div class="ecell${metric?' metric':k?' has-help':''}"${click
          ?` role="button" tabindex="0" onclick="${click}" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();${click}}"`
          :''}>
          <div class="l">${dateLabel(l)}${k
            ?`<i class="qm" onclick="event.stopPropagation();openHelp(${esc4(k)})">?</i>`:''}</div>
          <div class="v">${v||'-'}</div></div>`;
      }).join('') + `</div>`;
    }
    if((s.lines||[]).length){
      const li = s.lines.map(t=>topLine(String(t).replace(/^\s*\d+\.\s*/,''))).filter(Boolean);
      // 전부 걸러졌으면 빈 목록을 남기지 않고 **왜 비었는지**를 적는다.
      h += li.length
        ? `<ol class="elines">${li.join('')}</ol>`
        : `<div class="bnone">기준일 ${baseDate()} 앞뒤 ${TOP_DAYS}일 안에 해당하는 건이 없습니다`
          + ` <span style="color:var(--ink-3)">(오래된 건은 07_불일치누락현황에서 보세요)</span></div>`;
    }
    h += `</div>`;
  });

  if(!usedBrief){
    h = `<div class="esec"><div class="eh"><i>1</i><span>당일 업무 실적</span></div>
      ${briefBlock()}<div class="egrid">${dailyIssueHtml()}</div></div>` + h;
  }
  return h;
}
function erpHtml(y){
  const M = erpDocs.months || {}, K = erpDocs.kinds || {};
  const mos = Object.keys(M).filter(k=>k.startsWith(y)).sort();
  if(!mos.length) return '';
  const tot = mos.reduce((a,k)=>a+M[k].합계, 0);
  const cnt = mos.reduce((a,k)=>a+M[k].건수, 0);
  const kinds = Object.entries(K).sort((a,b)=>b[1]-a[1]);
  const max = Math.max(1, ...mos.map(k=>M[k].합계));
  return `<h4>${y}년 ERP 매출 (세금계산서 발행 기준 · 계산서 ${cnt}장 / ${fmt(tot)}원)</h4>
    <table class="stack">
      <colgroup><col style="width:14%"><col style="width:22%"><col style="width:10%"><col style="width:54%"></colgroup>
      <tr><th>월</th><th style="text-align:right">공급가액(원)</th><th style="text-align:right">장수</th><th>비중</th></tr>
      ${mos.map(k=>`<tr><td class="hd">${k.slice(5)}월</td>
        <td data-l="공급가액" style="text-align:right">${fmt(M[k].합계)}</td>
        <td data-l="장수" style="text-align:right">${M[k].건수}</td>
        <td data-l="비중"><span style="display:inline-block;height:8px;border-radius:4px;background:var(--brand);
          width:${Math.round(M[k].합계/max*58)}%;min-width:3px;vertical-align:middle"></span>
          <span style="font-size:10.5px;color:var(--ink-3);margin-left:6px">${Math.round(M[k].합계/tot*100)}%</span></td></tr>`).join('')}
    </table>
    <div style="font-size:11px;color:var(--ink-3);margin-top:6px">
      유형별: ${kinds.map(([k,v])=>`${k} ${fmt(v)}`).join(' · ')}<br>
      ※ ERP는 여러 작업을 한 장으로 묶어 발행합니다 — 장수는 계산서 매수이지 작업 건수가 아닙니다.</div>
    ${bundleHtml(y)}`;
}
/* 계산서 1장에 어떤 프로젝트가 묶였는지 (26_계산서구성 추정) */
function bundleHtml(y){
  const rows = (erpDocs.rows||[]).filter(r=>String(r.월||'').startsWith(y));
  if(!rows.length) return '';
  const kind = r => String(r.판정||'').split('(')[0];
  const ok  = rows.filter(r=>['확정','유력'].includes(kind(r))).length;
  const est = rows.filter(r=>kind(r)==='추정').length;
  const out = rows.filter(r=>kind(r)==='대상외').length;
  const poc = rows.filter(r=>kind(r)==='PO확인').length;
  const unk = rows.length - ok - est - out - poc;
  return `<h4 style="margin-top:14px">계산서별 포함 프로젝트
      <span class="norm">— 확정 ${ok}${poc?` · PO확인 ${poc}`:''}${est?` · 추정 ${est}`:''} · 대상외 ${out}${unk?` · 미상 ${unk}`:''} (번호를 누르면 그 건으로 이동)</span></h4>
    <div style="font-size:11px;color:var(--ink-3);margin:-4px 0 6px">
      확정 = 밴드에 적힌 구성 그대로 · PO확인 = 쿠팡 PO 번호까지 확인(품목은 아리바에서 조회)<br>
      대상외 = 계단·철거·신규납품처럼 AS·점검이 아닌 별도 공사(원장에 작업 행이 없습니다)</div>
    <table class="stack">
      <colgroup><col style="width:15%"><col style="width:11%"><col style="width:19%"><col style="width:55%"></colgroup>
      <tr><th>전표</th><th>판정</th><th style="text-align:right">공급가액(원)</th><th>포함 프로젝트NO</th></tr>
      ${rows.map(r=>{
        const v = String(r.판정||'미상'), k = v.split('(')[0];
        const col = k==='확정'?'#12813F' : k==='유력'?'#0B6BCB' : k==='추정'?'#B25E09'
                  : k==='대상외'?'#667085' : (k==='밴드확인'||k==='PO확인')?'#0B6BCB' : '#98A2B3';
        const list = String(r.포함프로젝트||'').split(',').map(s=>s.trim()).filter(Boolean);
        return `<tr><td class="hd">${r.전표}</td>
          <td data-l="판정"><span style="color:${col};font-weight:700">${k}</span>${
            v.includes('(') && k==='확정' ? `<div style="font-size:10px;color:var(--ink-3)">${v.slice(v.indexOf('(')+1,-1)}</div>` : ''}</td>
          <td data-l="공급가액" style="text-align:right">${fmt(r.공급가액)}</td>
          <td class="prjs" data-l="포함 프로젝트NO"><div class="chips">${
            list.map(k=>`<span class="prj" onclick="openByPrj('${k}')">${k}</span>`).join('')
            || `<span style="color:var(--ink-3);font-size:11px">${v.includes('(')?v.slice(v.indexOf('(')+1,-1):'후보 없음'}</span>`}
            </div>${r.후보합계 && r.후보합계!==r.공급가액
              ? `<div style="font-size:10.5px;color:var(--ink-3);margin-top:2px">후보합계 ${fmt(r.후보합계)}원 — 계산서와 ${fmt(Math.abs(r.공급가액-r.후보합계))}원 차이</div>` : ''}
          </td></tr>`;}).join('')}
    </table>
    <div style="font-size:11px;color:var(--ink-3);margin-top:6px">
      ※ 캠프·유형·기간으로 좁힌 <b>추정</b>입니다. 금액 합이 계산서와 맞으면 '확정'.
      품목 단위 판매현황이 들어오면 건별로 정확히 배분됩니다.</div>`;
}
function srcRow(label, s, emptyMsg){
  if(!s || !s.total) return `<tr><td class="hd">${label}</td><td data-l="결과" colspan="2" style="color:#98A2B3">${emptyMsg}</td></tr>`;
  const pct = Math.round(s.ok/s.total*100);
  const res = s.ok ? `대상 ${s.total}건 중 <b>${s.ok}건 확인(${pct}%)</b> · 미확인 ${s.miss}건`
                   : `검출 ${s.total}건 — 확인필요`;
  return `<tr><td class="hd">${label}</td><td data-l="결과">${res}</td>
    <td class="prjs" data-l="미확인 프로젝트NO"><div class="chips">${
      (s.miss_prj||[]).filter(Boolean).map(k=>
        `<span class="prj" onclick="openByPrj('${k}')">${k}</span>`).join('') || '-'}</div></td></tr>`;
}
function renderDaily(){
  if(!BRIEF && !BRIEF_LOADING) loadBrief();
  const S_ = srcStats || {};
  const selectedDay = baseDate() || todayISO();
  const today = new Date(selectedDay+'T00:00:00');
  const y = APP_YEAR, m = String(today.getMonth()+1).padStart(2,'0');
  // 선택한 보고 기준일을 월·분기·연간 통계와 캡처가 공통으로 사용한다.
  const ytdFrom = y + '-01-01', ytdTo = selectedDay;
  const inYTD = r => { const d = dateOf(r,'settle'); return d && d >= ytdFrom && d <= ytdTo; };
  const S = settleRows.filter(inYTD);
  const outYTD = settleRows.length - S.length;
  const Sm = S.filter(r=>inPeriod(r.완료일||'', y, m));
  const 유상 = S.filter(r=>r.비용구분==='유상');
  const 발행 = 유상.filter(r=>r.계산서==='발행').length;
  const 미수 = S.reduce((a,r)=>a+(+r.미수금||0),0);
  const issues = S.filter(r=>needAction(r));
  const asOpen = (works.as||[]).filter(r=>r.진행상태 && r.진행상태!=='작업완료');
  const pmWait = (works.pm||[]).filter(r=>r.점검상태 && r.점검상태!=='완료');
  const byType = {};
  issues.forEach(r=>{ (byType[r.상태]=byType[r.상태]||[]).push(r); });
  // KPI를 누르면 그 숫자를 만든 실제 건들을 볼 수 있게 집합을 등록해 둔다
  const KS = {}; window._kpiSets = KS;
  const reg = (k, title, rows, kind) => { KS[k] = {title, rows: rows||[], kind: kind||'settle'}; return k; };
  const AS_m = (works.as||[]).filter(r=>inPeriod(dateOf(r,'as'),y,m));
  const PM_m = (works.pm||[]).filter(r=>inPeriod(dateOf(r,'pm'),y,m));
  const qNo=Math.floor((+m-1)/3)+1, qFrom=String(qNo*3-2).padStart(2,'0'),
        qTo=String(qNo*3).padStart(2,'0');
  const inQuarter=d=>{const s=String(d||'');return s.startsWith(y+'-')&&s.slice(5,7)>=qFrom&&s.slice(5,7)<=qTo;};
  const Sq=S.filter(r=>inQuarter(dateOf(r,'settle')));
  const AS_q=(works.as||[]).filter(r=>inQuarter(dateOf(r,'as')));
  const PM_q=(works.pm||[]).filter(r=>inQuarter(dateOf(r,'pm')));
  reg('k_all',  `정산 총계 (${y}년 누적)`, S);
  reg('k_amt',  `공급가액 (${y}년 누적)`, [...S].sort((a,b)=>(+b.공급가액||0)-(+a.공급가액||0)));
  reg('k_due',  '미수금 보유 건', S.filter(r=>(+r.미수금||0)>0));
  reg('k_tax',  '계산서 미발행 (유상 기준)', 유상.filter(r=>r.계산서!=='발행'));
  reg('k_iss',  '조치 필요', issues);
  reg('k_as',   '돌발AS 진행중', asOpen, 'as');
  reg('k_pm',   '정기점검 대기', pmWait, 'pm');
  reg('m_set',  `${+m}월 정산`, Sm);
  reg('m_amt',  `${+m}월 공급가액`, [...Sm].sort((a,b)=>(+b.공급가액||0)-(+a.공급가액||0)));
  reg('m_pay',  `${+m}월 입금`, Sm.filter(r=>(+r.입금액||0)>0));
  reg('m_as',   `${+m}월 돌발AS`, AS_m, 'as');
  reg('m_pm',   `${+m}월 정기점검`, PM_m, 'pm');
  reg('m_iss',  `${+m}월 조치필요`, Sm.filter(r=>needAction(r)));
  reg('q_set',  `${qNo}분기 정산`, Sq);
  reg('q_as',   `${qNo}분기 돌발AS`, AS_q, 'as');
  reg('q_pm',   `${qNo}분기 정기점검`, PM_q, 'pm');
  const cell = (v,l,k)=>{
    const on = k && KS[k] && KS[k].rows.length;
    return `<div class="rcell${on?' clk':''}"${on?` onclick="openKpi('${k}')"`:''}>` +
      `<div class="v">${v}</div><div class="l">${l}${on?' <span class="go">▸</span>':''}</div></div>`;
  };
  const rptLogo = `<div class="logo app-icon"><img src="/icon-192.png?v=csos-20260730" alt="CSOS"></div>`;
  $('rpt').innerHTML = `
    <div class="rhead">
      ${rptLogo}
      <img class="report-uni-brand" src="/brand/universal-lift-horizontal.png" alt="UNIVERSAL LIFT &amp; HITACHI KOREA">
      <h2 style="font-size:17px">Coupang Service Operations System<small>일일보고 · (주)유니버셜리프트앤히타치코리아</small></h2>
      <div class="rdate">보고일<b>${(execRep.meta&&execRep.meta['보고일'])||`${y}-${m}-${String(today.getDate()).padStart(2,'0')}`}</b>
        <span style="font-size:10.5px;font-weight:600;color:var(--ink-3)">집계기준일 ${
          (execRep.meta&&execRep.meta['집계기준일'])||'-'} · 보고자 ${(execRep.meta&&execRep.meta['보고자'])||'-'}</span>
        <span style="font-size:10.5px;font-weight:700;color:#2452E6">데이터 업데이트 ${
          (BRIEF&&BRIEF['데이터업데이트일시'])||new Date().toISOString().slice(0,16).replace('T',' ')}</span>
        <button onclick="openRptDates()" style="margin-top:5px;background:#EEF2FF;color:#1C3FA8;border:0;
          border-radius:8px;padding:5px 10px;font-size:11px;font-weight:800;font-family:inherit;
          cursor:pointer">📅 날짜 변경</button></div>
    </div>
    <h4>${+m}월 실적 <span class="norm">— 선택 기준일 ${selectedDay}</span></h4>
    <div class="rgrid">
      ${cell(Sm.length+'건','정산','m_set')}
      ${cell(fmt(Sm.reduce((a,r)=>a+(+r.공급가액||0),0)),'공급가액(원)','m_amt')}
      ${cell(fmt(Sm.reduce((a,r)=>a+(+r.입금액||0),0)),'입금(원)','m_pay')}
      ${cell(AS_m.length+'건','돌발AS','m_as')}
      ${cell(PM_m.length+'건','정기점검','m_pm')}
      ${cell(Sm.filter(r=>needAction(r)).length+'건','당월 조치필요','m_iss')}
    </div>
    <h4>${y}년 ${qNo}분기 통계 <span class="norm">— ${+qFrom}~${+qTo}월</span></h4>
    <div class="rgrid">
      ${cell(Sq.length+'건','분기 정산','q_set')}
      ${cell(fmt(Sq.reduce((a,r)=>a+(+r.공급가액||0),0)),'분기 공급가액(원)','q_set')}
      ${cell(AS_q.length+'건','분기 돌발AS','q_as')}
      ${cell(PM_q.length+'건','분기 정기점검','q_pm')}
      ${cell(Sq.filter(r=>needAction(r)).length+'건','분기 조치필요','q_set')}
    </div>
    <h4>${y}년 누적 <span class="norm">— ${ytdFrom} ~ ${ytdTo}</span></h4>
    <div class="rgrid k7">
      ${cell(S.length+'건','정산 총계','k_all')}
      ${cell(fmt(S.reduce((a,r)=>a+(+r.공급가액||0),0)),'공급가액(원)','k_amt')}
      ${cell(fmt(미수),'미수금(원)','k_due')}
      ${cell((유상.length?Math.round(발행/유상.length*100):0)+'%','계산서 발행율','k_tax')}
      ${cell(issues.length+'건','조치 필요','k_iss')}
      ${cell(asOpen.length+'건','돌발AS 진행중','k_as')}
      ${cell(pmWait.length+'건','정기점검 대기','k_pm')}
    </div>
    ${outYTD?`<div style="font-size:11px;color:var(--ink-3);margin-top:6px">
      ※ 완료일이 없거나 ${y}년·선택 기준일 이후인 ${outYTD}건은 누적에서 제외</div>`:''}
    ${erpHtml(y)}
    ${execHtml()}
    ${(()=>{ window._issueSets = {}; Object.entries(byType).forEach(([t,l])=>window._issueSets[t]=l); return ''; })()}
    <h4>조치 필요 — 유형별 (돌발AS / 정기점검 구분) <span class="norm">— 유형을 누르면 담당자별 전달</span></h4>
    <table class="stack">
      <colgroup><col style="width:23%"><col style="width:9%"><col style="width:17%"><col style="width:16%"><col style="width:35%"></colgroup>
      <tr><th>유형</th><th style="text-align:right">건수</th><th>돌발AS · 정기점검</th><th style="text-align:right">금액(원)</th><th>프로젝트NO</th></tr>
      ${Object.entries(byType).sort((a,b)=>b[1].length-a[1].length).map(([t,l])=>{
        const as=l.filter(r=>String(r.업무구분||'').includes('돌발')).length;
        const pm=l.filter(r=>String(r.업무구분||'').includes('점검')).length;
        return `<tr style="cursor:pointer" onclick="openIssueType(${esc4(t)})">
         <td class="hd"><b>${t}</b> <span class="go">▸</span></td><td data-l="건수" style="text-align:right">${l.length}</td>
         <td data-l="구분">AS ${as} · 점검 ${pm}${l.length-as-pm?` · 기타 ${l.length-as-pm}`:''}</td>
         <td data-l="금액(원)" style="text-align:right">${fmt(l.reduce((a,r)=>a+(+r.공급가액||0),0))}</td>
         <td class="prjs" data-l="프로젝트NO"><div class="chips">${l.slice(0,5).map(r=>{
           const k = projectNoOf(r)||recordIdOf(r);
           return `<span class="prj" onclick="event.stopPropagation();openByPrj(${esc4(k)})" title="${esc2(recordIdOf(r))} ${esc2(r.캠프명||'')} — 클릭해 상세">${esc2(projectNoOf(r)||'프로젝트 미확정')}</span>`;
         }).join('')}${l.length>5?`<span class="more">외 ${l.length-5}건</span>`:''}</div></td></tr>`;}).join('')
      || '<tr><td colspan="5">조치 필요 없음 ✅</td></tr>'}
    </table>
    <h4>4원천 검증 (밴드·카톡·ERP·쿠팡PO)</h4>
    <table class="stack">
      <colgroup><col style="width:16%"><col style="width:40%"><col style="width:44%"></colgroup>
      <tr><th>원천</th><th>결과</th><th>미확인 프로젝트NO</th></tr>
      ${srcRow('밴드 게시', S_.band, '게시글 수집 필요')}
      ${srcRow('카톡 보고', S_.kakao, '대화 내보내기 투입 시 자동 대조')}
      ${srcRow('ERP 원장', S_.erp, '계정별원장 투입 시 자동 대조(4유형)')}
      ${srcRow('쿠팡 PO', S_.po, 'PO 목록 투입 시 자동 대조')}
    </table>
    <div class="rfoot"><span>자동 생성 · Coupang Service Operations System Agent</span><span>${new Date().toLocaleString('ko-KR')}</span></div>`;

  // 캡처(Canvas)용 동일 데이터 — HTML과 이미지가 항상 같은 내용을 보이도록 한 곳에서 구성
  const srcRows = [], srcDetails = [];
  const addSrc = (label,s,empty)=>{
    const hasData = !!(s&&s.total);
    const result = !hasData ? empty : s.ok
      ? `대상 ${s.total}건 중 ${s.ok}건 확인(${Math.round(s.ok/s.total*100)}%) · 미확인 ${s.miss}건`
      : `검출 ${s.total}건 — 확인필요`;
    const missProjects = hasData ? (s.miss_prj||[]).filter(Boolean) : [];
    srcRows.push([label, result, missProjects.join(', ')||'-']);
    srcDetails.push({label, result, missProjects, empty:!hasData});
  };
  addSrc('밴드 게시', S_.band, '게시글 수집 필요');
  addSrc('카톡 보고', S_.kakao, '대화 내보내기 투입 시 자동');
  addSrc('ERP 원장', S_.erp, '계정별원장 투입 시 자동');
  addSrc('쿠팡 PO', S_.po, 'PO 목록 투입 시 자동');
  // 대표 캡처도 화면과 같은 현장 일지·정기점검 기간을 쓴다.
  // 수치를 다시 계산하거나 분기 기본값으로 되돌리면 화면과 이미지가 달라진다.
  const capturePm = pmStats();
  const captureWorkLog = ((BRIEF&&BRIEF['일지대조'])||{});
  const captureBriefAs = ((BRIEF&&BRIEF['돌발AS'])||{});
  const captureDailyTasks = ((BRIEF&&BRIEF['당일처리목록'])||[]).map(x=>({
    project: x['프로젝트NO'] || '프로젝트 미확정',
    camp: x['캠프명'] || '-',
    date: x['일자'] || '',
    requestDate: x['접수일'] || '',
    poster: x['게시자'] || '',
    handler: x['담당기사'] || '',
    request: x['왜'] || '',
    action: x['무엇'] || x['상태'] || ''
  }));
  _rptData = {
    date: (execRep.meta&&execRep.meta['보고일']) || `${y}-${m}-${String(today.getDate()).padStart(2,'0')}`,
    meta: execRep.meta || {},
    summary: execRep.summary || [],
    exec: (execRep.sections||[]).map(s=>({title:s.title, items:s.items||[],
                                          groups:s.groups||[], lines:s.lines||[]})),
    kpiTitle: `${y}년 누적 (${ytdFrom} ~ ${ytdTo})`,
    kpiNote: outYTD ? `※ 완료일이 없거나 ${y}년·선택 기준일 이후인 ${outYTD}건은 누적에서 제외` : '',
    erpTitle: `${y}년 ERP 매출 (정기점검·돌발 AS(AS) · 세금계산서 발행 기준)`,
    erpTotal: Object.entries(erpDocs.months||{}).filter(([k])=>k.startsWith(y))
                .reduce((a,[,v])=>a+v.합계,0),
    erpMonths: Object.entries(erpDocs.months||{}).filter(([k])=>k.startsWith(y)).sort()
                .map(([k,v])=>[k.slice(5)+'월', v.합계, v.건수+'장']),
    pmProgress: [
      [`${capturePm.예정}건`, `${capturePm.라벨} 예정`],
      [`${capturePm.실행}건`, `${capturePm.라벨} 실행`, '#12813F'],
      [`${capturePm.미실행}건`, `${capturePm.라벨} 미실행`, capturePm.미실행?'#B54708':'#12813F'],
      [`${capturePm.진행률}%`, `${capturePm.라벨} 진행률`, capturePm.진행률>=60?'#12813F':'#B54708'],
    ],
    workLog: captureWorkLog,
    updatedAt: (BRIEF&&BRIEF['데이터업데이트일시']) ||
      new Date().toISOString().slice(0,16).replace('T',' '),
    // 유수비 대표가 올리고 류지영 매니저가 처리한 택배·부품발송처럼
    // 현장 AS 완료와 별개인 당일 업무도 저장 이미지에서 빠지지 않게 싣는다.
    dailyBrief: BRIEF ? {
      date: BRIEF['기준일'] || '',
      metrics: [
        [`${captureBriefAs['신규접수']||0}건`, '돌발AS 신규 접수'],
        [`${captureBriefAs['완료']||0}건`, '돌발AS 당일 완료', '#12813F'],
        [`${captureBriefAs['업무처리']||captureDailyTasks.length}건`, '당일 별도 업무 처리',
          captureDailyTasks.length?'#1B41BC':'#12813F']
      ],
      activities: captureDailyTasks
    } : null,
    kpi: [[S.length+'건','정산 총계','#2452E6'], [fmt(S.reduce((a,r)=>a+(+r.공급가액||0),0)),'공급가액(원)'],
          [fmt(미수),'미수금(원)', 미수?'#B54708':'#12813F'],
          [(유상.length?Math.round(발행/유상.length*100):0)+'%','계산서 발행율'],
          [issues.length+'건','조치 필요', issues.length?'#C0212E':'#12813F'],
          [asOpen.length+'건','돌발AS 진행중'], [pmWait.length+'건','정기점검 대기']],
    monthTitle: `${+m}월 실적`,
    month: [[Sm.length+'건','정산'], [fmt(Sm.reduce((a,r)=>a+(+r.공급가액||0),0)),'공급가액(원)'],
            [fmt(Sm.reduce((a,r)=>a+(+r.입금액||0),0)),'입금(원)'],
            [(works.as||[]).filter(r=>inPeriod(dateOf(r,'as'),y,m)).length+'건','돌발AS'],
            [(works.pm||[]).filter(r=>inPeriod(dateOf(r,'pm'),y,m)).length+'건','정기점검'],
            [Sm.filter(r=>needAction(r)).length+'건','당월 조치필요']],
    quarterTitle: `${y}년 ${qNo}분기 통계 (${+qFrom}~${+qTo}월)`,
    quarter: [[Sq.length+'건','분기 정산'],
              [fmt(Sq.reduce((a,r)=>a+(+r.공급가액||0),0)),'분기 공급가액(원)'],
              [AS_q.length+'건','분기 돌발 AS(AS)'],
              [PM_q.length+'건','분기 정기점검'],
              [Sq.filter(r=>needAction(r)).length+'건','분기 조치필요']],
    dailyIssues: dailyIssueRows().map(x=>[`${x.count}건`,x.label,x.count?'#B42318':'#177245',x.detail]),
    issues: Object.entries(byType).sort((a,b)=>b[1].length-a[1].length).map(([t,l])=>{
      const as=l.filter(r=>String(r.업무구분||'').includes('돌발')).length;
      const pm=l.filter(r=>String(r.업무구분||'').includes('점검')).length;
      return [t, `${l.length}건 (AS ${as}·점검 ${pm})`, fmt(l.reduce((a,r)=>a+(+r.공급가액||0),0)),
        l.slice(0,5).map(r=>projectNoOf(r)||'프로젝트 미확정').join(', ')+(l.length>5?` 외 ${l.length-5}건`:'')];}),
    srcs: srcRows,
    srcDetails
  };
}



/* ══ 처리 안내 ══════════════════════════════════════════════════
   "이건 어디서 확인해서 어떻게 넣어야 프로그램·엑셀에 반영되나"를
   건마다 붙여 준다. 상태·문제유형 이름을 키로 쓴다. */
const HELP = {
  '잔여 미청구액': {
    what: '작업은 끝났는데 <b>거래명세서로 아직 청구하지 못한 금액</b>입니다.<br>'
        + '계산식: <code>실제작업합계 − 거래명세서합계</code> 중 <b>양수만</b> 더한 값 '
        + '(06_거래서류청구수금 시트 「미청구액」 열).<br>'
        + '작업금액보다 명세서 금액이 작으면 그만큼 덜 청구한 것이므로 여기 잡힙니다. '
        + '<b>「작업금액 불일치」와 같은 건을 금액으로 본 것</b>입니다 — 건수로 보면 불일치, 돈으로 보면 미청구액.',
    where: ['06_거래서류청구수금 시트에서 <b>미청구액</b> 열이 0보다 큰 행',
            '앱 [확인필요] 탭의 <b>금액 불일치</b> 항목',
            '밴드 그 작업 게시글의 거래명세서 사진 / 이카운트 [판매 → 거래명세서]'],
    how: ['작업금액과 명세서 금액 중 <b>어느 쪽이 맞는지</b> 확인합니다',
          '<b>명세서가 맞으면</b>: 실제작업 공급가액을 명세서 금액에 맞춰 고칩니다',
          '<b>작업금액이 맞으면</b>: 차액만큼 <b>추가 청구</b>하고 명세서를 다시 끊습니다',
          '고치면 이 숫자는 자동으로 0이 됩니다(수기 조정 불필요)'],
    who: '유현민 (금액 확인 후 정정 또는 추가 청구)'
  },
  '작업금액 불일치': {
    what: '실제 작업금액과 거래명세서 금액이 <b>서로 다른 건수</b>입니다.<br>'
        + '금액이 아직 안 채워진 건(미청구)과 신규납품·계단 같은 <b>별도 공사는 세지 않습니다</b> — '
        + '그건 각각 따로 관리합니다.',
    where: ['06_거래서류청구수금 시트 「작업대비거래명세서차액」 열이 0이 아닌 행',
            '앱 [확인필요] 탭 → 구분 <b>금액</b>'],
    how: ['「잔여 미청구액」과 같은 건입니다 — 그쪽 설명대로 처리하면 함께 사라집니다'],
    who: '유현민'
  },
  '잔여 미수금액': {
    what: '세금계산서까지 발행했는데 <b>아직 입금되지 않은 금액</b>입니다.<br>'
        + '계산식: <code>세금계산서합계 − 입금액</code>. 청구는 끝났고 돈만 안 들어온 상태입니다.',
    where: ['06_거래서류청구수금 시트 「세금계산서대비입금차액」 열',
            '통장 입금 내역 / 쿠팡 아리바 지급 예정일'],
    how: ['입금이 확인되면 <b>입금일·입금액</b>을 앱이나 대장에 입력하면 자동으로 줄어듭니다',
          '지급예정일이 지났는데 미입금이면 쿠팡 담당자에게 확인'],
    who: '유현민'
  },
  '청구액 (당일)': {
    what: '<b>오늘 날짜로 발행된 거래명세서</b>의 합계입니다. 오늘 청구한 금액이 없으면 0입니다.',
    where: ['06_거래서류청구수금 시트에서 거래명세서 발행일이 오늘인 행'],
    how: ['오늘 발행한 명세서가 있는데 0으로 나오면 <b>발행일이 대장에 안 적힌 것</b>입니다 — 발행일을 입력하세요'],
    who: '류지영 매니저 / 유현민'
  },
  '문제 업무 건수(중복 제거)': {
    what: '검증에서 걸린 업무를 <b>같은 건은 한 번만</b> 세어 정리한 수입니다.<br>'
        + '한 건에 문제 코드가 여러 개 붙어 있어도(사진 미첨부 + ERP 미등록 …) 1건으로 셉니다. '
        + '아래 「문서 경고」와 겹쳐 보이는 이유가 이것입니다.',
    where: ['앱 [확인필요] 탭', '23_확인필요현황 시트'],
    how: ['건별로 열어 문제 코드마다 안내된 대로 처리하면 줄어듭니다'],
    who: '유현민 / 담당기사'
  },
  '세금계산서 기한 임박·초과': {
    what: '세금계산서는 <b>공급일이 속한 달의 다음 달 10일까지</b> 발행해야 합니다. '
        + '그 기한이 가까워졌거나 이미 지난 건수입니다. 넘기면 가산세가 붙습니다.',
    where: ['06_거래서류청구수금 시트에서 작업완료일 대비 세금계산서발행일이 빈 행'],
    how: ['이카운트에서 즉시 발행 → [매출(세금)계산서현황] 엑셀을 <code>inbox/</code>에 넣으면 자동 반영'],
    who: '유현민 (발행)'
  },
  'PO 미발행 · 확인필요': {
    what: '쿠팡 발주번호(PO)가 있어야 청구가 되는 건인데 <b>PO 번호가 비어 있거나 확인이 안 된</b> 건수입니다.',
    where: ['쿠팡 아리바(Ariba) 발주 화면', '밴드 「구매 오더 전송」 게시글의 PO 번호'],
    how: ['PO 목록 엑셀을 <code>inbox/</code>에 넣으면 자동으로 대조·연결됩니다',
          'PO가 아직 안 나왔으면 쿠팡 담당자에게 발주 요청'],
    who: '유현민'
  },
  '거래명세서 미작성': {
    what: '작업은 끝났는데 <b>거래명세서를 아직 안 끊은</b> 건수입니다. 청구의 첫 단계가 안 된 상태입니다.',
    where: ['06_거래서류청구수금 시트에서 거래명세서번호가 빈 행'],
    how: ['이카운트에서 명세서 발행 후 번호·발행일을 앱이나 대장에 입력'],
    who: '류지영 매니저 / 유현민'
  },
  '아리바 청구 미등록': {
    what: '세금계산서는 발행했는데 <b>쿠팡 아리바에 청구 등록을 안 한</b> 건수입니다. '
        + '등록하지 않으면 지급 절차가 시작되지 않습니다.',
    where: ['쿠팡 아리바 → 송장(Invoice) 등록 화면'],
    how: ['아리바에 청구 등록 후 대장의 아리바승인일을 입력'],
    who: '유현민'
  },

  '금액 미입력': {
    what: '작업은 끝났는데 공급가액이 비어 있습니다. 금액이 없으면 거래명세서·세금계산서를 만들 수 없어 청구가 멈춥니다.',
    where: ['밴드 해당 작업 게시글의 <b>견적서·거래명세서 사진</b>',
            '이카운트 → 판매 → <b>거래명세서</b> 화면에서 해당 현장 조회',
            '담당기사에게 작업 내역·금액 확인'],
    how: ['<b>앱에서 바로</b>: 이 화면 아래 입력칸에 공급가액을 넣고 저장 → 09:50 자동 실행 때 관리대장 빈칸에 기록',
          '<b>사진으로</b>: 명세서 사진을 <code>band/docs_inbox/</code>에 넣으면 OCR이 금액을 읽어 자동 입력(금액 검산 통과분만)'],
    who: '유현민 (금액 확인 후 입력)'
  },
  '세금계산서 미발행': {
    what: '거래명세서는 있는데 세금계산서가 아직 발행되지 않았습니다. 매출 인식과 수금이 지연됩니다.',
    where: ['이카운트 → 판매 → <b>매출(세금)계산서현황</b>에서 해당 명세서번호 조회',
            '실제로 발행했는데 대장에만 안 적힌 경우가 많습니다'],
    how: ['<b>이미 발행했다면</b>: 이카운트에서 [매출(세금)계산서현황]을 엑셀로 받아 <code>inbox/</code>에 넣기 → 자동 대조되어 발행일이 채워집니다',
          '<b>아직 안 했다면</b>: 이카운트에서 발행 후 위와 동일하게 파일 투입'],
    who: '유현민 (발행 처리)'
  },
  '미청구(전표 없음)': {
    what: '작업은 끝났는데 거래명세서(전표)가 만들어지지 않았습니다. 청구 자체가 시작되지 않은 상태입니다.',
    where: ['이카운트 → 판매 → <b>거래명세서</b>에서 해당 현장·일자로 조회',
            '밴드 게시글에 명세서 사진이 있는지 확인'],
    how: ['이카운트에서 거래명세서를 발행한 뒤 [거래명세서 현황]을 엑셀로 받아 <code>inbox/</code>에 투입',
          '또는 명세서 사진을 <code>band/docs_inbox/</code>에 넣으면 번호·일자가 자동 입력됩니다'],
    who: '유현민 (명세서 발행)'
  },
  '입금 대기': {
    what: '세금계산서는 발행됐는데 입금이 확인되지 않았습니다.',
    where: ['이카운트 → 회계 → <b>거래처별계정별원장</b>(쿠팡로지스틱스)에서 입금 확인',
            '통장·은행 거래내역'],
    how: ['[거래처별계정별원장] 엑셀을 <code>inbox/</code>에 넣으면 입금일·입금액이 자동 대조됩니다'],
    who: '유현민 (입금 확인)'
  },
  'PO 미발행': {
    what: '쿠팡 PO(발주번호)가 필요한 건인데 번호가 없습니다. PO가 없으면 쿠팡 시스템에서 청구가 막힐 수 있습니다.',
    where: ['쿠팡 담당자에게 PO 발행 요청',
            '이미 받았다면 쿠팡에서 받은 <b>PO 목록</b> 확인'],
    how: ['PO 목록 엑셀을 <code>inbox/</code>에 넣으면 프로젝트NO·금액으로 자동 매칭됩니다(파일명 자유)',
          '개별 건은 이 화면 아래 입력칸에 PO번호를 직접 넣어도 됩니다'],
    who: '유현민 → 쿠팡 담당자'
  },
  'ERP 계산서(묶음)': {
    what: '이 행은 <b>작업 1건이 아니라 세금계산서 1장</b>입니다. 이카운트는 여러 작업(최대 15건)을 한 장으로 묶어 발행하는데, 관리대장에 그 달 작업 기록이 없어서 계산서로 대신 채워 넣은 것입니다.',
    where: ['이카운트 → 판매 → 해당 계산서의 <b>[내역보기]</b>를 열면 묶인 작업 목록이 나옵니다'],
    how: ['작업 건별로 나누려면 <b>품목 단위 판매현황</b> 엑셀이 필요합니다. 그걸 <code>inbox/</code>에 넣으면 건별 금액까지 분해됩니다',
          '밴드 1~3월 게시글을 수집하면 작업 건별 기록이 생겨 이 묶음 행이 실제 작업 행으로 대체됩니다'],
    who: '유현민 (자료 투입)'
  },
  '밴드 게시 누락': {
    what: '작업은 완료로 기록됐는데 밴드에 게시글이 없습니다. 증빙이 비어 있어 나중에 확인이 어렵습니다.',
    where: ['해당 담당기사에게 밴드 게시 여부 확인'],
    how: ['기사가 밴드에 올리면 다음 자동 수집 때 반영됩니다',
          '앱 [대시보드] → 담당자별 확인 필요에서 해당 기사를 눌러 <b>[이미지로 전달]</b>로 요청하세요'],
    who: '담당기사'
  },
  '점검 지연': {
    what: '정기점검 예정일이 지났는데 완료 기록이 없습니다.',
    where: ['담당기사에게 실제 점검 여부 확인'],
    how: ['점검을 했다면 밴드 게시 → 자동 반영. 아직이면 일정 재조정',
          '앱 [대시보드] → 담당자별 확인 필요에서 목록을 이미지로 만들어 기사에게 전달'],
    who: '담당기사'
  },
  '미점검': {
    what: '정기점검 예정일이 지났는데 점검 기록이 없습니다.',
    where: ['담당기사에게 실제 점검 여부 확인'],
    how: ['점검 후 밴드 게시 → 자동 반영'],
    who: '담당기사'
  },
  '완료일 누락': {
    what: '상태는 작업완료인데 완료일이 비어 있습니다. 완료일이 없으면 월별 실적·정산 대상에서 빠집니다.',
    where: ['밴드 게시글의 작업 일자', '담당기사 확인'],
    how: ['밴드 게시글이 있으면 자동 수집 때 채워집니다. 없으면 이 건을 열어 완료일을 직접 입력'],
    who: '유현민 또는 담당기사'
  },
  '담당자 미지정': {
    what: '접수는 됐는데 담당기사가 정해지지 않았습니다.',
    where: ['배정 계획 확인 (김준형·권오철·김필우·차동호)'],
    how: ['이 건을 열어 담당기사를 입력하면 관리대장에 반영됩니다'],
    who: '유현민 (배정)'
  },
  '접수': {
    what: '접수만 되고 아직 작업이 끝나지 않은 건입니다.',
    where: ['담당기사에게 진행 상황 확인'],
    how: ['작업 완료 후 밴드 게시 → 자동으로 완료 처리됩니다'],
    who: '담당기사'
  }
};

/* 상태·문제유형 → 안내 카드 HTML (없으면 빈 문자열) */
function helpBox(key){
  const h = HELP[String(key||'').trim()];
  if(!h) return '';
  const li = a => (a||[]).map(x=>`<li>${x}</li>`).join('');
  return `<div class="helpbox">
    <div class="hh">이 건은 어떻게 처리하나요?</div>
    <div class="hw">${h.what}</div>
    <div class="hs"><b>① 어디서 확인</b><ul>${li(h.where)}</ul></div>
    <div class="hs"><b>② 어떻게 반영</b><ul>${li(h.how)}</ul></div>
    ${h.who?`<div class="hwho">담당: ${h.who}</div>`:''}
  </div>`;
}

/* 용어·처리방법 전체 목록 (실행 탭 도움말) */
/* 보고 화면 라벨 → HELP 열쇠. 괄호·공백 표기가 흔들려도 찾아낸다. */
function helpKey(label){
  const raw = String(label||'').trim();
  if(HELP[raw]) return raw;
  const norm = t => String(t).replace(/[\s()（）·]/g,'').toLowerCase();
  const n = norm(raw);
  let k = Object.keys(HELP).find(x => norm(x) === n);
  if(k) return k;
  // '작업금액 불일치 (현재)'·'청구액 (당일)'처럼 뒤에 (현재)·(당일)이 붙은 표기도 같은 항목이다
  const base = norm(raw.replace(/\((현재|당일|누적|금월|전체)\)\s*$/,''));
  k = Object.keys(HELP).find(x => norm(x) === base || norm(x).startsWith(base) && base.length >= 4);
  return k || '';
}
function helpAll(){
  const terms = [
    ['대표 프로젝트NO', '모든 건이 번호로 식별되게 붙인 번호입니다. 원본 번호가 없으면 ①본문에서 찾은 UJ번호 ②같은 캠프·같은 달의 실제 작업 번호 ③ERP 전표번호(<code>ERP-260110-2</code> 형태) ④자체 ID 순으로 정합니다. <b>없는 UJ 번호를 지어내지는 않습니다.</b>'],
    ['ERP 계산서(묶음)', '이카운트가 여러 작업을 한 장으로 묶어 발행한 세금계산서입니다. 관리대장에 그 달 작업 기록이 없을 때만 대신 표시되며, 작업 기록이 들어오면 실제 건으로 대체됩니다.'],
    ['원천 데이터', '대조의 근거가 되는 자료입니다. 밴드(작업 증빙) · 카톡(보고) · ERP(회계) · 쿠팡 PO(발주). 넣는 곳은 각각 자동 수집 / <code>kakao/inbox/</code> / <code>inbox/</code> / <code>inbox/</code> 입니다.'],
    ['빈칸만 입력', '에이전트는 <b>비어 있는 칸에만</b> 값을 씁니다. 이미 사람이 넣은 값은 절대 덮어쓰지 않습니다.'],
    ['vN 파일', '관리대장은 수정할 때마다 v47 → v48처럼 새 파일로 저장됩니다. 원본은 그대로 두어 문제가 생기면 되돌릴 수 있습니다. 항상 <b>가장 큰 번호</b>가 최신입니다.'],
    ['09:50 자동 실행', '매일 오전 9시 50분에 전체 대조 → 확정된 값 자동 입력 → 리포트 생성까지 자동으로 돕니다. 류지영 매니저의 오전 입력(08:00~09:30) 이후에 돌도록 맞춰져 있습니다.']
  ];
  return `<div class="card"><h3>처리 방법 · 용어 설명</h3>
    <div class="toplist">${Object.keys(HELP).map(k=>`
      <div class="topitem" onclick="openHelp(${esc4(k)})">
        <span style="min-width:0"><b>${k}</b>
          <div style="font-size:11.5px;color:var(--ink-3);margin-top:3px">${HELP[k].what.replace(/<[^>]+>/g,'').slice(0,46)}…</div></span>
        <span style="flex:none;color:var(--brand);font-weight:800;font-size:11px">처리법 ▸</span></div>`).join('')}</div>
    <div style="margin-top:14px"><h3 style="font-size:13px">용어</h3>
      <div class="dl" style="margin-top:8px">${terms.map(([t,d])=>
        `<dt>${t}</dt><dd style="line-height:1.65">${d}</dd>`).join('')}</div></div>
  </div>`;
}
function openHelp(key){
  const h = HELP[key]; if(!h) return;
  openPane(`<h2>${key}</h2><div class="sub">확인 방법과 반영 절차</div>${helpBox(key)}`);
}

/* ══ 담당자별 확인 필요 / 내가 확인할 사항 ══════════════════════════
   현장 기사에게 물어봐야 할 것과, 내가(관리) 처리해야 할 것을 갈라서 보여준다.
   담당자 항목은 눌러 내역을 보고, 그대로 이미지로 만들어 전달할 수 있다. */
const _today = () => new Date().toISOString().slice(0,10);
/* 로컬 기준 오늘(YYYY-MM-DD). toISOString은 UTC라 한국 시간대에서 하루 밀릴 수 있다 */
function todayISO(){
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,'0')}-${String(t.getDate()).padStart(2,'0')}`;
}
/* 끝난 건 = 더 물어볼 것 없음. '취소'도 종결로 본다(기사에게 확인 요청할 대상이 아니다) */
const _fin = s => ['작업완료','완료','정상','취소','종결','철회','AS전환'].includes(String(s||'').trim());
/* 담당기사 칸에 '김준형, 김필우'처럼 여러 명이 들어있으면 각자에게 나눠 붙인다 */
function owners(v){
  const l = String(v||'').split(/[,·/]|\s{2,}/).map(x=>x.trim()).filter(Boolean);
  return l.length ? l : ['담당 미지정'];
}

/* ★ 사용자 지시(2026-07-28): "2025년도 자료는 보관만 하고, 지금 목표는 2026년 정리다."
   원장에서 지우지 않는다 — **화면에서만** 뺀다. 나중에 정리할 때 FOCUS_YEAR 를 낮추거나
   0 으로 두면 예전 건이 그대로 돌아온다. 자료를 잃지 않으면서 지금 볼 것만 보는 방법이다. */
const FOCUS_YEAR = Number(APP_YEAR);
function inFocus(it){
  if(!FOCUS_YEAR) return true;
  const d = String((it&&it.date) || '').slice(0,4);
  if(!/^\d{4}$/.test(d)) return false;
  return +d === FOCUS_YEAR;
}

function buildBoard(){
  const by = {}, add = (whos,it) => { if(!inFocus(it)) return;
    owners(whos).forEach(w=>{ (by[w] = by[w]||[]).push(it); }); };
  (works.as||[]).forEach(r=>{
    const id = r.접수ID || r.프로젝트NO;
    const base = {id, prj:r.프로젝트NO, camp:r.캠프명, kind:'돌발AS', band:r['밴드 바로가기']||''};
    if(!_fin(r.진행상태))
      add(r.담당기사, Object.assign({}, base, {issue:r.진행상태||'접수', date:r.접수일자||'', desc:(r.신청내용||'').slice(0,60)}));
    else if(String(r.진행상태||'').trim()==='작업완료' && !r.작업완료일)
      add(r.담당기사, Object.assign({}, base, {issue:'완료일 누락', date:r.접수일자||'', desc:'작업완료인데 완료일이 비어 있음'}));
    const c = checks[id];
    if(c && c.band === '미확인')
      add(r.담당기사, Object.assign({}, base, {issue:'밴드 게시 누락', date:r.작업완료일||r.접수일자||'', desc:'밴드에 작업 게시가 없음'}));
  });
  (works.pm||[]).forEach(r=>{
    const id = r.점검ID || r.프로젝트NO, base = {id, prj:r.프로젝트NO, camp:r.캠프명, kind:'정기점검'};
    if(!_fin(r.점검상태)){
      // 아직 예정일이 안 지난 점검은 '확인할 것'이 아니다 — 지났거나 미점검인 것만 올린다
      const late = r.점검예정일 && String(r.점검예정일) < _today();
      const miss = String(r.점검상태||'').trim() === '미점검';
      if(!late && !miss) return;
      add(r.담당기사, Object.assign({}, base, {issue: late?'점검 지연':'미점검', date:r.점검예정일||'',
                desc: late?('예정일 '+r.점검예정일+' 경과'):''}));
    } else if(String(r.점검상태||'').trim()==='완료' && !r.실제점검일)
      add(r.담당기사, Object.assign({}, base, {issue:'점검일 누락', date:r.점검예정일||'', desc:'완료인데 실제점검일이 비어 있음'}));
  });
  Object.values(by).forEach(l=>l.sort((a,b)=>String(a.date||'9999').localeCompare(String(b.date||'9999'))));
  return by;
}

function buildMine(){
  const out = [];
  // 2025년 건은 원장에 그대로 두고 화면에서만 뺀다(FOCUS_YEAR — 위 설명 참고).
  const rowYear = r => String(r.접수일자||r.점검예정일||r.완료일||r.작업완료일||
                             r.실제점검일||r.요청일||'').slice(0,4);
  const focus = r => { const y = rowYear(r); return /^\d{4}$/.test(y) && +y === FOCUS_YEAR; };
  const push = (cat,todo,lv,items,go)=>{ items = (items||[]).filter(focus);
    if(items.length) out.push({cat:cat,todo:todo,lv:lv,items:items,go:go,
    n:items.length, amt:items.reduce((a,r)=>a+(+r.공급가액||0),0)}); };
  [['세금계산서 미발행','발행 처리 필요','danger'],
   ['미청구(전표 없음)','거래명세서 발행 필요','danger'],
   ['금액 미입력','금액 확인 후 입력','warn'],
   ['입금 대기','입금 확인 필요','warn']].forEach(function(x){
     push(x[0], x[1], x[2], settleRows.filter(r=>r.상태===x[0]), ()=>goList('settle',x[0])); });
  push('PO 미발행','쿠팡에 PO 요청','warn',
       settleRows.filter(r=>String(r.PO필요||'').includes('필요') && !r.PO번호), ()=>goList('settle',''));
  // 아직 진행 중인 건만 배정이 필요하다. 이미 끝난 건·ERP 보완 행은 제외.
  push('담당자 미지정','기사 배정 필요','warn',
       (works.as||[]).concat(works.pm||[]).filter(r=>
         !(r.담당기사||'').trim() && r.출처!=='ERP' && !_fin(r.진행상태||r.점검상태)),
       ()=>goList('as',''));
  /* 대조에 쓸 자료가 아직 안 들어온 것도 '내가 할 일'이다.
     ★ 예전 문구는 'OO 원천 없음'이었는데 '원천'이 무슨 말인지 되물으셨다(2026-07-27).
       화면 문구는 **그걸 처음 보는 사람이 바로 알아듣는 말**이어야 한다. */
  const src = srcStats || {};
  [['band','밴드 게시글','밴드에서 글을 아직 못 가져왔습니다 — [실행] 탭에서 밴드 수집'],
   ['kakao','카톡 대화','카톡 대화 내보내기(.txt)를 kakao/inbox 폴더에 넣어주세요'],
   ['erp','ERP 거래처별계정별원장','이카운트에서 원장을 엑셀로 받아 inbox 폴더에 넣어주세요'],
   ['po','쿠팡 PO 목록','쿠팡 PO 목록 엑셀을 inbox 폴더에 넣어주세요']].forEach(function(x){
     const S = src[x[0]] || {};
     if(S.total) return;                         // 대조가 돌았으면 할 일 없음
     /* ★ '파일을 안 넣었다'와 '넣었는데 내용이 비어 있다'는 완전히 다른 일이다.
        2026-07-27에 ERP 파일 3개가 회사명 한 줄만 있는 빈 내보내기였는데, 화면에는
        그냥 '없음'으로만 떠서 아무도 몰랐다. 넣으신 걸 못 봤다고 하면 안 된다. */
     const inb = S.inbox;
     if(inb && inb.empty && inb.empty.length){
       out.push({cat:x[1]+' — 넣으신 파일이 비어 있습니다', lv:'warn', n:0, amt:0, items:[],
         todo:'다시 내보내 주세요: '+inb.empty.slice(0,2).join(', ')+
              (inb.empty.length>2 ? ' 외 '+(inb.empty.length-2)+'개' : ''),
         go:()=>show('run')});
     } else if(inb && inb.files){
       out.push({cat:x[1]+' — 파일은 있는데 대조를 안 돌렸습니다', lv:'info', n:0, amt:0, items:[],
         todo:'[실행] 탭에서 대조를 한 번 돌리면 됩니다 (파일 '+inb.files+'개)',
         go:()=>show('run')});
     } else {
       out.push({cat:x[1]+' 아직 없음', todo:x[2], lv:'info', n:0, amt:0, items:[], go:()=>show('run')});
     }
   });
  return out;
}

function renderBoard(){
  const by = buildBoard(); window._board = by;
  const names = Object.keys(by).sort((a,b)=>by[b].length-by[a].length);
  $('assignlist').innerHTML = names.map(function(n){
    const l = by[n], t = {};
    l.forEach(i=>t[i.issue]=(t[i.issue]||0)+1);
    const chips = Object.entries(t).sort((a,b)=>b[1]-a[1]).slice(0,3)
      .map(e=>'<span class="chip c-warn" style="font-size:10.5px">'+e[0]+' '+e[1]+'</span>').join('');
    return '<div class="topitem" onclick="openAssignee('+esc4(n)+')">'+
      '<span style="min-width:0"><b>'+n+'</b>'+
      '<div style="margin-top:5px;display:flex;flex-wrap:wrap;gap:4px">'+chips+'</div></span>'+
      '<span style="text-align:right;flex:none"><b>'+l.length+'건</b><br>'+
      '<span style="font-size:11px;color:var(--brand);font-weight:800">내역·전달 ▸</span></span></div>';
  }).join('') || '<div style="color:var(--ink-3);font-size:13px">담당자에게 확인할 사항이 없습니다 🎉</div>';

  const mine = buildMine(); window._mine = mine;
  $('mylist').innerHTML = mine.map(function(m,i){
    const col = m.lv==='danger' ? 'var(--danger)' : m.lv==='warn' ? 'var(--warn)' : 'var(--ink-3)';
    return '<div class="topitem" onclick="openMine('+i+')">'+
      '<span style="min-width:0"><b>'+m.cat+'</b>'+
      '<div style="font-size:11.5px;color:var(--ink-3);margin-top:3px">'+m.todo+'</div></span>'+
      '<span style="text-align:right;flex:none"><b style="color:'+col+'">'+(m.n?m.n+'건':'투입 대기')+'</b>'+
      (m.amt?'<br><span style="font-size:11px;color:var(--ink-3)">'+fmt(m.amt)+'원</span>':'')+'</span></div>';
  }).join('') || '<div style="color:var(--ink-3);font-size:13px">확인할 사항이 없습니다 🎉</div>';
}

/* ══ 확인 필요 전용 화면 ═════════════════════════════════════════════════════
   23_확인필요현황을 원본으로 삼되, 단순 표가 아니라
   요약 → 유형/담당자 → 검색 목록 → 정확한 업무기록 → 전달 이미지로 이어 준다. */
function checkTypeOf(r){ return String(r.문제유형||r.경고내용||r.검증결과||'기타 확인').trim(); }
const CHECK_OWNER_RULES=[
  {name:'류지영',scope:'캠프·일정, 카톡/밴드, 현장자료, 완료일, 거래명세서, 세금계산서, 입금, 금액 불일치 확인',role:'운영·회계 확인 담당'},
  {name:'오종현',scope:'PO 원본·견적서, 구매·입금 원천자료 취합·누락 확인',role:'원천자료 취합 담당'},
  {name:'유현민',scope:'PO, ERP·원장 미등록, 시스템 연결 확인',role:'시스템 확인 담당'}
];
function confirmedCheckOwner(r){
  const category=String(r.구분||'').trim(), type=checkTypeOf(r);
  const cat=category.toUpperCase(), upper=type.toUpperCase();
  if(/PO\s*원본|견적서|구매\s*자료|입금\s*원천\s*자료|원천\s*자료|원본\s*(수집|누락)/i.test(type))
    return '오종현';
  if(cat==='PO'||cat==='ERP'||upper==='PO'||upper.startsWith('PO ')||
     upper==='ERP'||upper.startsWith('ERP ')||/원장\s*미등록|시스템\s*연결/.test(type)) return '유현민';
  if(category==='정산'||/거래명세서|명세서|세금계산서|입금|수금|금액|미청구|청구|회계/.test(type))
    return '류지영';
  if(/캠프|일정|예정일|접수일자|카톡|밴드|현장\s*자료|사진|완료일|완료보고|작업일|점검일|날짜/.test(type))
    return '류지영';
  return '';
}
function openConfirmedOwnerScope(name){
  const p=CHECK_OWNER_RULES.find(x=>x.name===name); if(!p) return;
  const list=(window._checkRows||[]).filter(r=>checkOwnerOf(r)===name);
  openPane(`<h2>${esc2(p.name)} ${chip('확정 기준')}</h2>
    <div class="sub">${esc2(p.role)} · 현재 ${list.length}건이 이 기준으로 분류되었습니다.</div>
    <div class="metric-basis"><b>적용 범위</b> ${esc2(p.scope)}</div>
    ${list.length?`<button class="abtn primary" onclick="closePane();openCheckOwner(${esc4(name)})">현재 ${list.length}건 목록 보기</button>`:
      '<div class="card">현재 확인할 항목이 없습니다.</div>'}`);
}
function checkOwnerOf(r){
  const confirmed=confirmedCheckOwner(r);
  if(confirmed) return confirmed;
  const direct=String(r.담당자||r.담당기사||'').trim();
  if(direct) return direct;
  const raw=checkIdOf(r), first=typeof rowById==='function' ? rowById(raw) : null;
  const source=first&&first.원천업무ID ? rowById(first.원천업무ID) : first;
  if(source&&String(source.담당기사||source.담당자||'').trim())
    return String(source.담당기사||source.담당자).trim();
  const project=checkProjectOf(r);
  const linked=(works.as||[]).concat(works.pm||[]).filter(x=>
    project&&String(x.프로젝트NO||'').trim()===project&&String(x.담당기사||'').trim());
  const names=[...new Set(linked.flatMap(x=>owners(x.담당기사)).filter(x=>x!=='담당 미지정'))];
  return names.length ? names.join(', ') : '담당 표기 없음';
}
function checkDateOf(r){ return String(r.일자||r.기준일||r.접수일자||r.점검예정일||r.완료일||'').slice(0,10); }
function checkIdOf(r){ return String(r.ID||r.업무ID||r.정산ID||r.접수ID||r.점검ID||'').trim(); }
function checkProjectOf(r){ return projectNoOf(r); }
function checkDescOf(r){ return String(r['내용·근거']||r.문제내용||r.경고내용||r.검증문제코드||'').trim(); }
function checkDaysOld(r){
  const d=normDate(checkDateOf(r)); if(!d) return -1;
  const a=new Date(d+'T00:00:00'), b=new Date(todayISO()+'T00:00:00');
  return Math.floor((b-a)/86400000);
}
function samePo(a,b){ return String(a||'').replace(/[^A-Z0-9]/gi,'').toUpperCase()===
  String(b||'').replace(/[^A-Z0-9]/gi,'').toUpperCase(); }
function checkLink(r){
  const raw=checkIdOf(r);
  let hit=typeof rowById==='function' ? rowById(raw) : null;
  if(!hit&&/^PO/i.test(raw)) hit=settleRows.find(r=>samePo(r.PO번호,raw))||null;
  const textProject=(checkDescOf(r).match(/\bUJ26\d{4,}\b/i)||[])[0]||'';
  let kind='';
  if(/^AS-/i.test(raw) || (hit&&hit.접수ID)) kind='as';
  else if(/^PM-/i.test(raw) || (hit&&hit.점검ID)) kind='pm';
  else if(/^JS-/i.test(raw) || (hit&&hit.정산ID)) kind='settle';
  const targetId=kind==='as'&&hit ? hit.접수ID : kind==='pm'&&hit ? hit.점검ID
    : kind==='settle'&&hit ? hit.정산ID : raw;
  return {
    kind, id:String(targetId||raw).trim(), issueId:raw,
    project:String((hit&&projectNoOf(hit))||textProject||checkProjectOf(r)||'').trim(),
    camp:String((hit&&hit.캠프명)||r.캠프명||'').trim(),
    title:String((hit&&hit.프로젝트명)||r.프로젝트명||'').trim()
  };
}
function checkMetricRow(r){
  const x=checkLink(r), amount=Number(String(r.금액||'').replace(/[^0-9.-]/g,''))||0;
  return {
    프로젝트NO:x.project, 프로젝트명:x.title, 캠프명:x.camp,
    일자:checkDateOf(r), 담당자:checkOwnerOf(r),
    문제:checkDescOf(r)||checkTypeOf(r), 상태:checkTypeOf(r), 금액:amount,
    종류:x.kind||'check', 레코드ID:x.kind?x.id:x.issueId, ID:x.id||x.project
  };
}
function checkRows(){
  return (issuesData.rows||[]).filter(r=>rowIs2026(r,'issue')&&!isSideWork(r));
}
function groupCheck(rows, getter){
  const out={};
  rows.forEach(r=>{ const k=getter(r); (out[k]=out[k]||[]).push(r); });
  return out;
}
function checkSortGroups(g){
  return Object.keys(g).sort((a,b)=>g[b].length-g[a].length||a.localeCompare(b,'ko'));
}
function openCheckRows(title, rows, basis){
  rows=Array.isArray(rows)?rows:[];
  const label=title||'확인 필요 목록';
  window._briefMetric={label,data:{
    rows:rows.map(checkMetricRow), count:rows.length, kind:'check',
    basis:basis||`${APP_YEAR}년 23_확인필요현황 · 현재 미해결 ${rows.length}건`
  }};
  openExecMetric(label);
}
function openCheckType(type){
  const rows=checkRows().filter(r=>checkTypeOf(r)===type);
  openCheckRows(type,rows,`${APP_YEAR}년 확인필요 원장 중 문제유형이 “${type}”인 건`);
}
function openCheckOwner(owner){
  const rows=checkRows().filter(r=>checkOwnerOf(r)===owner);
  openCheckRows(`${owner} 확인 요청`,rows,
    `${APP_YEAR}년 확인필요 원장에서 담당자가 “${owner}”인 건 · 프로젝트 번호를 누르면 원기록으로 이동`);
}
function openCheckFiltered(){
  const rows=window._checkVisible||checkRows();
  const t=$('checktype')&&$('checktype').value, o=$('checkowner')&&$('checkowner').value;
  const bits=[t,o].filter(Boolean), title=bits.length?bits.join(' · '):'확인 필요 전체';
  openCheckRows(title,rows,`${APP_YEAR}년 확인필요 원장 · 현재 화면의 검색·필터 결과 ${rows.length}건`);
}
function openCheckStale(){
  const rows=checkRows().filter(r=>checkDaysOld(r)>30);
  openCheckRows('30일 초과 확인',rows,`${todayISO()} 기준, 확인필요 원장 일자가 30일을 넘긴 건`);
}
function checkTargetExists(project){
  const p=String(project||'').trim(); if(!p) return false;
  return settleRows.some(r=>String(r.프로젝트NO||'')===p) ||
    (works.as||[]).some(r=>String(r.프로젝트NO||'')===p) ||
    (works.pm||[]).some(r=>String(r.프로젝트NO||'')===p);
}
function openCheckSource(r){
  const x=checkLink(r);
  window._singleCheckRow=r;
  const fields=Object.entries(r).filter(([,v])=>String(v==null?'':v).trim());
  openPane(`<h2>${esc2(checkTypeOf(r))} ${chip('원장 연결 확인')}</h2>
    <div class="sub">${esc2(x.project||x.issueId||'확인 항목')} · 연결된 AS·점검·정산 원장 행이 없어 확인 원문을 엽니다.</div>
    <div class="metric-basis"><b>다음 조치</b> 프로젝트 번호·PO 번호로 원장을 확인한 뒤,
      해당 업무가 등록되면 이 화면에서 자동으로 원기록과 연결됩니다.</div>
    <div class="dl">${fields.map(([k,v])=>`<dt>${esc2(k)}</dt><dd>${esc2(v)}</dd>`).join('')}</div>
    ${helpBox(checkTypeOf(r))}
    <div class="actions sticky">
      <button class="abtn primary" onclick="openSingleCheckCapture()"><span class="em">📋</span>이 항목 목록 열기</button>
      <button class="abtn" onclick="showCheckRaw()"><span class="em">📋</span>원본표 보기</button></div>`);
}
function openSingleCheckCapture(){
  const r=window._singleCheckRow; if(r) openCheckRows(checkTypeOf(r),[r],'원장 연결을 확인할 단일 항목');
}
function openCheckByKey(id,project){
  const rows=checkRows();
  const r=rows.find(x=>checkIdOf(x)===String(id||'')) ||
          rows.find(x=>checkProjectOf(x)===String(project||'')); if(!r) return;
  const link=checkLink(r);
  if(link.kind) openRecord(link.kind,link.id,link.project);
  else if(checkTargetExists(link.project)) openByPrj(link.project);
  else openCheckSource(r);
}
function openCheckItem(i){
  const r=(window._checkVisible||[])[i]; if(!r) return;
  const x=checkLink(r);
  if(x.kind) openRecord(x.kind,x.id,x.project);
  else if(checkTargetExists(x.project)) openByPrj(x.project);
  else openCheckSource(r);
}
function showCheckRaw(){
  show('settle'); setMode('check');
}
function resetCheckFilters(){
  if($('checkq')) $('checkq').value='';
  if($('checktype')) $('checktype').value='';
  if($('checkowner')) $('checkowner').value='';
  renderCheckList();
}
function renderCheckList(){
  if(!$('checklist')) return;
  const q=String(($('checkq')&&$('checkq').value)||'').trim().toLowerCase();
  const type=String(($('checktype')&&$('checktype').value)||'');
  const owner=String(($('checkowner')&&$('checkowner').value)||'');
  const all=checkRows();
  const rows=all.filter(r=>{
    if(type&&checkTypeOf(r)!==type) return false;
    if(owner&&checkOwnerOf(r)!==owner) return false;
    return !q || Object.values(r).join(' ').toLowerCase().includes(q);
  });
  window._checkVisible=rows;
  $('checksummary').innerHTML=`전체 ${all.length}건 중 <b>${rows.length}건</b> 표시 · 오래된 건부터 최근 순`;
  $('checklist').innerHTML=rows.map((r,i)=>{
    const x=checkLink(r), id=x.id||'', prj=x.project||'프로젝트 미확정';
    const amount=Number(String(r.금액||'').replace(/[^0-9.-]/g,''))||0;
    return `<div class="srow check-row metric-row" role="button" tabindex="0" onclick="openCheckItem(${i})"
      onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();openCheckItem(${i})}">
      <div class="top"><span class="project-line"><span class="prjno">${esc2(prj)}</span>
        ${id!==prj?`<span class="sid">· ${esc2(id)}</span>`:''}
        <span class="camp-name">${esc2(x.camp||'캠프 미상')}</span></span>
        ${amount?`<span class="check-money">${fmt(Math.abs(amount))}원</span>`:chip(checkTypeOf(r))}</div>
      <div class="meta"><span class="metric-date">${esc2(checkDateOf(r)||'-')}</span>
        <span>${esc2(checkOwnerOf(r))}</span><span>${esc2(r.구분||'')}</span>
        ${amount?chip(checkTypeOf(r)):''}</div>
      ${checkDescOf(r)?`<div class="check-desc">${esc2(checkDescOf(r))}</div>`:''}
    </div>`;
  }).join('')||'<div class="card">조건에 맞는 확인 항목이 없습니다 🎉</div>';
}
function renderCheckHub(){
  if(!$('checkkpis')) return;
  const rows=checkRows(), byType=groupCheck(rows,checkTypeOf), byOwner=groupCheck(rows,checkOwnerOf);
  const stale=rows.filter(r=>checkDaysOld(r)>30), unassigned=byOwner['담당 표기 없음']||[];
  const ras=(repData&&repData['돌발AS'])||{}, rpm=(repData&&repData['정기점검'])||{};
  const rdocs=(repData&&repData['거래명세서'])||{}, policies=(repData&&repData['업무기준확인필요'])||[];
  const statementMissing=(rdocs['미발행목록']||[]).length, pmGap=Number(rpm['계획대비']||0);
  $('checkkpis').innerHTML=`
    <div class="check-kpi danger" onclick="openCheckRows('확인 필요 전체',checkRows())">
      <div class="v">${rows.length}</div><div class="l">전체 확인 필요</div><div class="s">누르면 전체 목록·캡처</div></div>
    <div class="check-kpi" onclick="document.getElementById('checktypes').scrollIntoView({behavior:'smooth'})">
      <div class="v">${Object.keys(byType).length}</div><div class="l">문제유형</div><div class="s">유형별로 바로 구분</div></div>
    <div class="check-kpi ${unassigned.length?'warn':''}" onclick="openCheckOwner('담당 표기 없음')">
      <div class="v">${unassigned.length}</div><div class="l">담당 표기 없음</div><div class="s">원장 담당자 칸 확인</div></div>
    <div class="check-kpi ${stale.length?'warn':''}" onclick="openCheckStale()">
      <div class="v">${stale.length}</div><div class="l">30일 초과</div><div class="s">오래된 확인 항목</div></div>
    <div class="check-kpi ${ras['D+2초과']?'danger':''}" onclick="openRepresentativeList('as-d2')">
      <div class="v">${ras['D+2초과']||0}</div><div class="l">AS D+2 초과</div><div class="s">대표 지속보고 대상</div></div>
    <div class="check-kpi ${pmGap<0?'warn':''}" onclick="openRepresentativeList('pm-gap')">
      <div class="v">${pmGap>=0?'+':''}${pmGap}</div><div class="l">점검 계획대비</div><div class="s">목표누계 대비 실제</div></div>
    <div class="check-kpi ${statementMissing?'warn':''}" onclick="openRepresentativeList('statement-unissued')">
      <div class="v">${statementMissing}</div><div class="l">명세서 미발행</div><div class="s">유상 발행대상 기준</div></div>`;
  $('checktypes').innerHTML=checkSortGroups(byType).map(t=>{
    const list=byType[t], ownersN=new Set(list.map(checkOwnerOf)).size;
    return `<button class="check-group" onclick="openCheckType(${esc4(t)})">
      <span><span class="name">${esc2(t)}</span><span class="hint">담당 ${ownersN}명 · 목록/전달 이미지</span></span>
      <span class="count">${list.length}건</span></button>`;
  }).join('')||'<div style="color:var(--ink-3)">확인할 항목이 없습니다 🎉</div>';
  const confirmedNames=new Set(CHECK_OWNER_RULES.map(p=>p.name));
  const confirmedOwnerCards=CHECK_OWNER_RULES.map(p=>{
    const list=byOwner[p.name]||[], typesN=new Set(list.map(checkTypeOf)).size;
    return `<div class="topitem" onclick="${list.length?`openCheckOwner(${esc4(p.name)})`:`openConfirmedOwnerScope(${esc4(p.name)})`}">
      <span style="min-width:0"><b>${esc2(p.name)}</b>
        <div style="font-size:11px;color:var(--ink-3);margin-top:3px">${esc2(p.scope)}</div></span>
      <span style="text-align:right;flex:none"><b>${list.length}건</b><br>
        <span style="font-size:11px;color:var(--brand);font-weight:800">${list.length?`${typesN}개 유형 · 내역·전달`:'확정 범위'} ▸</span></span></div>`;
  }).join('');
  $('checkowners').innerHTML=confirmedOwnerCards+checkSortGroups(byOwner).filter(o=>!confirmedNames.has(o)).map(o=>{
    const list=byOwner[o], typesN=new Set(list.map(checkTypeOf)).size;
    return `<div class="topitem" onclick="openCheckOwner(${esc4(o)})">
      <span style="min-width:0"><b>${esc2(o)}</b>
        <div style="font-size:11px;color:var(--ink-3);margin-top:3px">${typesN}개 유형 · ${
          list.slice(0,2).map(r=>esc2(checkTypeOf(r))).join(', ')}${list.length>2?' 외':''}</div></span>
      <span style="text-align:right;flex:none"><b>${list.length}건</b><br>
        <span style="font-size:11px;color:var(--brand);font-weight:800">내역·전달 ▸</span></span></div>`;
  }).join('');

  const mine=buildMine(); window._mine=mine;
  $('checkmine').innerHTML=mine.map((m,i)=>{
    const col=m.lv==='danger'?'var(--danger)':m.lv==='warn'?'var(--warn)':'var(--ink-3)';
    return `<div class="topitem" onclick="openMine(${i})">
      <span style="min-width:0"><b>${esc2(m.cat)}</b>
        <div style="font-size:11.5px;color:var(--ink-3);margin-top:3px">${esc2(m.todo)}</div></span>
      <span style="text-align:right;flex:none"><b style="color:${col}">${m.n?m.n+'건':'투입 대기'}</b>
        ${m.amt?`<br><span style="font-size:11px;color:var(--ink-3)">${fmt(m.amt)}원</span>`:''}</span></div>`;
  }).join('')||'<div style="color:var(--ink-3)">관리·청구 확인 사항이 없습니다 🎉</div>';
  if($('checkpolicies')) $('checkpolicies').innerHTML=policies.map((p,i)=>`
    <div class="topitem" onclick="openRepresentativeList('policy')">
      <span style="min-width:0"><b>${esc2(p.기준||'업무기준')}</b>
        <div style="font-size:11px;color:var(--ink-3);margin-top:3px">업무 처리 기준 추가 확인 필요</div></span>
      <span class="chip c-warn">${esc2(p.상태||'확인 필요')}</span></div>`).join('')||
    '<div style="color:var(--ink-3)">확인 대기 중인 업무기준이 없습니다.</div>';

  const typeEl=$('checktype'), ownerEl=$('checkowner');
  const keepT=typeEl.value, keepO=ownerEl.value;
  typeEl.innerHTML='<option value="">문제유형 전체</option>'+
    checkSortGroups(byType).map(t=>`<option value="${esc2(t)}">${esc2(t)} (${byType[t].length})</option>`).join('');
  ownerEl.innerHTML='<option value="">담당자 전체</option>'+
    checkSortGroups(byOwner).map(o=>`<option value="${esc2(o)}">${esc2(o)} (${byOwner[o].length})</option>`).join('');
  if(byType[keepT]) typeEl.value=keepT;
  if(byOwner[keepO]) ownerEl.value=keepO;
  renderCheckList();
}

/* onclick 속성에 안전하게 문자열 넣기 */
function esc4(s){ return "'" + String(s==null?'':s).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'&quot;') + "'"; }

/* 제공받은 유니버셜리프트 CI는 앱바에서 은은하게 노출하고 보고서 이미지에도 원본 비율로 넣는다.
   CSOS 상태 아이콘은 실행 중 애니메이션의 기준이므로 CI로 교체하지 않는다. */
function brandImg(cls){
  return brandLogo
    ? `<img src="/brand/${encodeURIComponent(brandLogo)}" alt="쿠팡" class="${cls||''}">`
    : '';
}
function applyBrand(){
  const el=document.querySelector('.uni-app-brand');
  if(el) el.src='/brand/universal-lift-horizontal.png';
}
function openPane(html){
  showSheet(html);        // 시트와 같은 창을 쓴다 — 뒤로가기·배경잠금도 그대로 적용된다
}

/* 대표보고 3·4절 숫자 → 그 숫자를 만든 정확한 원천 행.
   프로젝트NO 옆에 프로젝트명·캠프, 다음 줄에 날짜를 두어 캡처만 봐도 식별할 수 있게 한다. */
const METRIC_CAP = 60;
function metricDetail(label){
  const adhoc=window._briefMetric;
  if(adhoc && adhoc.label===label) return adhoc.data;
  return (execRep.details||{})[label] || {rows:[],basis:'집계 근거 없음',kind:'risk',count:0,amount:0};
}
function openStaleBrief(){
  const all=(BRIEF&&BRIEF['완료일미기입목록'])||[];
  const label='완료일이 안 적힌 오래된 건';
  const rows=all.map(x=>({
    프로젝트NO:x.프로젝트NO||'',
    캠프명:x.캠프명||'',
    일자:x.일자||'',
    담당자:x.담당기사||'기사 미배정',
    문제:'접수 후 30일 초과 · 작업완료일 미기입',
    상태:'확인필요',
    종류:x.레코드종류||'as',
    레코드ID:x.레코드ID||'',
    ID:x.레코드ID||x.프로젝트NO||''
  }));
  window._briefMetric={label,data:{
    rows, count:rows.length, kind:'brief-stale',
    basis:`${(BRIEF&&BRIEF['기준일'])||baseDate()} 기준, 접수 후 30일이 넘었지만 작업완료일이 빈 돌발AS`
  }};
  openExecMetric(label);
}
/* ── 정기점검 진행률 기간 (사용자 지시 2026-07-29: 월 단위로 골라 본다) ──────────
   처음 값은 **이번 분기**(대표가 늘 보던 기준)이고, 고르면 그 기간으로 다시 센다.
   ★ 서버가 준 분기 숫자를 쓰지 않고 `works.pm` 에서 직접 센다 — 그래야 아무 기간이나 된다.
   판정은 daily_brief 와 같다: 예정 = 점검예정일이 기간 안 · 실행 = 그중 실제점검일이 있음. */
let PM_RANGE = null;
function pmQuarter(){
  const m = +(baseDate()||todayISO()).slice(5,7) || (new Date().getMonth()+1);
  const q = Math.floor((m-1)/3)+1;
  return [String(3*q-2).padStart(2,'0'), String(3*q).padStart(2,'0')];
}
function pmMonthOpts(sel){
  return Array.from({length:12},(_,i)=>{ const v=String(i+1).padStart(2,'0');
    return `<option value="${v}"${v===sel?' selected':''}>${i+1}월</option>`; }).join('');
}
function setPmRange(a, b){
  if(a > b){ b = a; }                    // 시작이 뒤면 끝을 끌어와 맞춘다(오류를 띄우지 않는다)
  PM_RANGE = [a, b];
  renderDaily();
}
function pmPlanDate(r){
  const v = r['\uC810\uAC80\uC608\uC815\uC77C'];
  return normDate(v) || String(v||'').slice(0,10);
}
function pmSourceRows(state){
  const cutoff = baseDate() || todayISO();
  return (works.pm||[]).filter(r=>{
    const d = pmPlanDate(r);
    return d.startsWith(APP_YEAR) && d.slice(5,7) >= state.from && d.slice(5,7) <= state.to
      // 향후 일정은 보관하되, 기준일 현재 진행률의 분모에는 넣지 않는다.
      && d <= cutoff;
  });
}
function pmRangeRows(state){
  // 대표 진행률은 프로젝트 단위 업무의 예정/실행을 비교한다. UJ 번호가 없는
  // 류지영 원본 일정은 실행 탭의 원본 일정으로 보존하지만, 프로젝트 진행률을
  // 부풀리지 않도록 별도 매칭 대상으로 둔다.
  return pmSourceRows(state).filter(r=>String(r['\uD504\uB85C\uC81D\uD2B8NO']||'').trim());
}
function pmStats(){
  const [q0, q1] = pmQuarter();
  const [f, t] = PM_RANGE || [q0, q1];
  const y = APP_YEAR;
  const inR = d => { const s=String(d||''); if(!s.startsWith(y)) return false;
                     const mm=s.slice(5,7); return mm>=f && mm<=t; };
  const rows = pmRangeRows({from:f, to:t});
  const done = rows.filter(r=>/\d{4}-\d{2}-\d{2}/.test(String(r.실제점검일||''))).length;
  const label = (f==='01'&&t==='12') ? '연간' : (f===t ? `${+f}월` : `${+f}~${+t}월`);
  return {from:f, to:t, q0:q0, q1:q1, 라벨:label, 끝월:`${+t}월`,
          예정:rows.length, 실행:done, 미실행:rows.length-done,
          진행률: rows.length ? Math.round(done*100/rows.length) : 0};
}
function openPmBrief(scope){
  const p=(BRIEF&&BRIEF['정기점검'])||{}, d=(BRIEF&&BRIEF['기준일'])||baseDate();
  const plan=(BRIEF&&BRIEF['점검예정목록'])||[];
  const done=(BRIEF&&BRIEF['점검실행목록'])||[];
  // ★ 기간을 고를 수 있으므로(사용자 지시 2026-07-29) 분기 고정 목록 대신
  //   지금 선택된 기간으로 다시 고른다. 카드 숫자와 목록이 어긋나면 안 된다.
  const _ps = pmStats();
  const quarter=((BRIEF&&BRIEF['분기점검목록'])||[]).filter(x=>{
    const s=String(x.예정일||x.일자||''); if(!s.startsWith(APP_YEAR)) return false;
    const mm=s.slice(5,7); return mm>=_ps.from && mm<=_ps.to; });
  // 범위 버튼의 숫자와 상세 목록은 같은 원천(점검예정일·기준일까지)에서 만든다.
  // BRIEF의 기본 분기는 그대로 두되, 월/반기/연간 선택 때 목록이 달라지는 문제를 막는다.
  const liveQuarter=pmRangeRows(_ps).map(r=>{
    const plan=pmPlanDate(r), actual=normDate(r['\uC2E4\uC81C\uC810\uAC80\uC77C'])||String(r['\uC2E4\uC81C\uC810\uAC80\uC77C']||'').slice(0,10);
    return {
      ['\uD504\uB85C\uC81D\uD2B8NO']:r['\uD504\uB85C\uC81D\uD2B8NO']||'',
      ['\uCEA0\uD504\uBA85']:r['\uCEA0\uD504\uBA85']||'',
      ['\uC77C\uC790']:plan,
      ['\uC608\uC815\uC77C']:plan,
      ['\uC2E4\uD589\uC77C']:actual,
      ['\uB2F4\uB2F9\uAE30\uC0AC']:r['\uB2F4\uB2F9\uAE30\uC0AC']||'',
      ['\uB0B4\uC6A9']:r['\uC810\uAC80\uB0B4\uC6A9']||'',
      ['\uC0C1\uD0DC']:actual?'\uC2E4\uD589':'\uBBF8\uC2E4\uD589',
      ['\uB808\uCF54\uB4DCID']:r['\uC810\uAC80ID']||''
    };
  });
  let source=[], label='', basis='';
  if(scope==='day-plan'){
    source=plan;label=`${d} 정기점검 예정`;
    basis=`${d} 점검예정일이 기록된 ${plan.length}건`;
  }else if(scope==='day-done'){
    source=done;label=`${d} 정기점검 실행`;
    basis=`${d} 실제점검일이 기록된 ${done.length}건`;
  }else{
    source=scope==='quarter-done' ? liveQuarter.filter(x=>x.상태==='실행')
          : scope==='quarter-pending' ? liveQuarter.filter(x=>x.상태!=='실행') : liveQuarter;
    label=`${APP_YEAR}년 ${_ps.라벨} 정기점검 ${
      scope==='quarter-done'?'실행':scope==='quarter-pending'?'미실행':'전체 예정'}`;
    basis=`${APP_YEAR}년 ${_ps.라벨} 예정 ${_ps.예정}건 · 실행 ${_ps.실행}건 · 미실행 ${_ps.미실행}건`;
  }
  const rows=source.map(x=>({
    프로젝트NO:x.프로젝트NO||'', 캠프명:x.캠프명||'',
    일자:x.일자||x.예정일||x.실행일||'', 담당자:x.담당기사||'기사 미배정',
    문제:[x.예정일?`예정 ${x.예정일}`:'',x.실행일?`실행 ${x.실행일}`:'',
          x.왜||'점검내용 미기입'].filter(Boolean).join(' · '),
    상태:x.상태||'', 종류:'pm', 레코드ID:x.레코드ID||'',
    ID:x.레코드ID||x.프로젝트NO||''
  }));
  window._briefMetric={label,data:{rows,count:rows.length,kind:'brief-pm',basis}};
  openExecMetric(label);
}
function metricOpenCall(r){
  const kind=String(r.종류||''), id=recordIdOf(r), prj=projectNoOf(r);
  if(kind==='check') return `openCheckByKey(${esc4(id)},${esc4(prj)})`;
  return kind ? `openRecord(${esc4(kind)},${esc4(id)},${esc4(prj)})`
              : `openByPrj(${esc4(prj||id)})`;
}
function openExecMetric(label){
  const d=metricDetail(label), rows=d.rows||[];
  window._curMetricLabel=label;
  const amount=(+d.amount||0), projects=d.project_count;
  const summary = projects!=null ? `${projects}개 프로젝트 · ${rows.length}개 문제 행`
                : `${rows.length}건${amount?` · ${fmt(Math.abs(amount))}원`:''}`;
  const body = rows.map(r=>{
    const prj=esc2(projectLabel(r)), id=esc2(recordIdOf(r)), title=esc2(r.프로젝트명||'');
    const camp=esc2(r.캠프명||'캠프 미상'), when=esc2(r.일자||'-');
    const issue=esc2(r.문제||''), state=esc2(r.상태||''), owner=esc2(r.담당자||'');
    const amount=+r.금액||0;
    return `<div class="srow metric-row" role="button" tabindex="0"
      onclick="${metricOpenCall(r)}"
      onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();${metricOpenCall(r)}}">
      <div class="top"><span class="project-line"><span class="prjno">${prj}</span>
        ${title?`<span class="sid">· ${title}</span>`:''}<span class="camp-name">${camp}</span></span>
        ${amount?`<span class="amt">${fmt(Math.abs(amount))}원</span>`:state?chip(state):''}</div>
      <div class="meta"><span class="metric-date">${when}</span>
        ${id&&id!==prj?`<span class="sid">${id}</span>`:''}${owner?`<span>${owner}</span>`:''}</div>
      ${issue?`<div class="metric-issue">${issue}</div>`:''}
    </div>`;
  }).join('');
  openPane(`<h2>${esc2(label)} <span class="chip c-warn">${summary}</span></h2>
    <div class="sub">프로젝트 번호를 누르면 이 목록의 정확한 업무 기록으로 이동합니다.</div>
    <div class="metric-basis"><b>집계 기준</b> ${esc2(d.basis||'')}</div>
    <div class="lhead"><b>목록</b><span>${rows.length}건</span></div>
    <div class="slist">${body||'<div class="card">해당 건이 없습니다.</div>'}</div>
	    <div class="media-tools sticky" role="toolbar" aria-label="목록 도구">
	      <button class="icon-btn" onclick="printExecMetric()" title="인쇄" aria-label="인쇄"><img src="/icons/printer.svg" alt=""></button>
	      <button class="icon-btn primary" onclick="saveExecMetric()" title="이미지 저장" aria-label="이미지 저장"><img src="/icons/image-down.svg" alt=""></button>
	      <button class="icon-btn" onclick="copyExecMetric()" title="복사" aria-label="복사"><img src="/icons/clipboard-copy.svg" alt=""></button>
	      <button class="icon-btn excel" onclick="exportMetricXlsx()" title="엑셀 저장" aria-label="엑셀 저장"><img src="/icons/file-spreadsheet.svg" alt=""></button>
	    </div>`);
}

/* KPI 숫자 → 그 숫자를 만든 실제 건 목록 */
/* 조치 필요 유형(세금계산서 미발행 등) → 담당자별로 묶어 이미지 전달 */
function openIssueType(t){
  const list = (window._issueSets||{})[t] || [];
  const by = {};
  list.forEach(r=>{
    (owners(r.담당기사 || r.담당자).forEach(w=>{
      (by[w] = by[w] || []).push({
        id: r.정산ID || r.접수ID || r.점검ID || '', prj: r.프로젝트NO || '',
        camp: r.캠프명 || '', kind: r.업무구분 || '', issue: t,
        date: r.완료일 || r.접수일자 || r.점검예정일 || '',
        desc: r.공급가액 ? fmt(r.공급가액) + '원' : ''
      });
    }));
  });
  const names = Object.keys(by).sort((a,b)=>by[b].length-by[a].length);
  window._issueBoard = by;                                // 전체 담당자 보드는 덮어쓰지 않는다
  openPane(`<h2>${t} <span class="chip c-danger">${list.length}건</span></h2>
    <div class="sub">담당자별 목록을 열어 이미지로 저장하거나 복사할 수 있습니다</div>
    ${helpBox(t)}
    <div class="toplist" style="margin-top:12px">${names.map(n=>`
      <div class="topitem">
        <span style="min-width:0"><b>${n}</b>
          <div style="font-size:11.5px;color:var(--ink-3);margin-top:3px">${by[n].length}건 · ${
            by[n].slice(0,2).map(x=>x.prj||x.id).join(', ')}${by[n].length>2?' 외':''}</div></span>
        <span style="flex:none;display:flex;gap:6px">
          <button class="abtn primary" style="padding:7px 10px;font-size:12px"
            onclick="event.stopPropagation();openIssueAssignee(${esc4(n)})">목록 보기</button>
        </span></div>`).join('') ||
      '<div class="card">담당자 정보가 없는 건입니다</div>'}</div>`);
}
function openKpi(key){
  const set = (window._kpiSets||{})[key]; if(!set) return;
  const kind = set.kind, rows = set.rows;
  const sum = rows.reduce((a,r)=>a+(+r.공급가액||0),0);
  const line = r => {
    if(kind==='as')  return {id:r.접수ID||'', camp:r.캠프명, who:r.담당기사||'',
                             date:r.접수일자||'', st:r.진행상태||'', amt:0, prj:projectNoOf(r)};
    if(kind==='pm')  return {id:r.점검ID||'', camp:r.캠프명, who:r.담당기사||'',
                             date:r.점검예정일||'', st:r.점검상태||'', amt:0, prj:projectNoOf(r)};
    return {id:r.정산ID, camp:`${r.캠프명||''} · ${r.업무구분||''}`, who:'',
            date:r.완료일||'', st:r.상태||'', amt:+r.공급가액||0, prj:projectNoOf(r)};
  };
  openPane(`<h2>${set.title} <span class="chip c-warn">${rows.length}건</span></h2>` +
    `<div class="sub">${sum?fmt(sum)+'원 · ':''}누르면 해당 건 상세로 이동합니다</div>` +
    `<div class="slist" style="margin-top:12px">` + (rows.map(r=>{
      const v = line(r);
      return `<div class="srow" onclick="openRecord(${esc4(kind)},${esc4(v.id)},${esc4(v.prj)})">
        <div class="top"><span class="prjno">${esc2(v.prj||'프로젝트 미확정')}</span>${v.amt?`<span class="amt">${fmt(v.amt)}원</span>`:chip(v.st)}</div>
        <div class="top"><span class="camp">${v.camp||''}</span>
          <span class="camp">${v.who||''} ${v.date||''}</span></div>
        <div class="meta">${v.id?`<span class="sid">${esc2(v.id)}</span>`:''}
          ${v.amt?`<span>${chip(v.st)}</span><span>${r.명세서번호||''}</span>`:''}</div>
      </div>`;
    }).join('') || '<div class="card">해당 건이 없습니다</div>') + `</div>`);
}
function selectIssueAssignee(name){
  window._curAssignee=name;
  window._curAssigneeList=((window._issueBoard||{})[name]||[]);
}
function openIssueAssignee(name){
  openAssignee(name, ((window._issueBoard||{})[name]||[]));
}
function openAssignee(name, source){
  const l = Array.isArray(source) ? source : ((window._board||{})[name]||[]);
  window._curAssignee = name;
  window._curAssigneeList = l;
  // ★ 예전에는 도움말 카드와 버튼 아래에 목록을 놓아 **목록이 있는 줄도 몰랐다**
  //   (사용자 지적 2026-07-28). 목록이 주인공이므로 위로 올리고, 버튼은 아래에
  //   고정해 50건을 다 내려도 항상 손이 닿게 한다.
  openPane('<h2>'+name+' <span class="chip c-danger">'+l.length+'건</span></h2>'+
    '<div class="sub">확인·조치가 필요한 항목입니다. 아래 목록에서 <b>누르면 그 건으로 이동</b>하고,'+
    ' <b>밴드</b>를 누르면 원문 글이 열립니다.</div>'+
    '<div class="lhead"><b>목록</b><span>'+l.length+'건</span></div>'+
    '<div class="slist">' + (l.map(function(i){
      const key = i.prj||i.id;
      const band = i.band
        ? '<a class="blink" href="'+i.band+'" target="_blank" rel="noopener"'+
          ' onclick="event.stopPropagation()">밴드 ↗</a>' : '';
      return '<div class="srow" onclick="openByPrj('+esc4(key)+')">'+
        '<div class="top"><span class="id">'+key+'</span>'+
        '<span style="display:flex;gap:5px;align-items:center">'+band+
        '<span class="chip c-warn">'+i.issue+'</span></span></div>'+
        '<div class="top"><span class="camp">'+(i.camp||'')+' · '+i.kind+'</span>'+
        '<span class="camp">'+(i.date||'-')+'</span></div>'+
        (i.desc?'<div class="meta"><span style="flex:1">'+i.desc+'</span></div>':'')+'</div>';
    }).join('') || '<div class="card">항목 없음</div>') + '</div>'+
    (l.length?helpBox(l[0].issue):'')+
	    '<div class="media-tools sticky" role="toolbar" aria-label="담당자 목록 도구">'+
	      '<button class="icon-btn" onclick="printAssignee()" title="인쇄" aria-label="인쇄"><img src="/icons/printer.svg" alt=""></button>'+
	      '<button class="icon-btn primary" onclick="captureAssignee()" title="이미지 저장" aria-label="이미지 저장"><img src="/icons/image-down.svg" alt=""></button>'+
	      '<button class="icon-btn" onclick="copyAssigneeImage()" title="복사" aria-label="복사"><img src="/icons/clipboard-copy.svg" alt=""></button>'+
	      '<button class="icon-btn excel" onclick="exportAssigneeXlsx()" title="엑셀 저장" aria-label="엑셀 저장"><img src="/icons/file-spreadsheet.svg" alt=""></button></div>');
}

function openMine(idx){
  const m = (window._mine||[])[idx]; if(!m) return;
  const rows = (m.items||[]).slice(0, 200);
  const body = rows.length
    ? '<div class="slist" style="margin-top:12px">' + rows.map(function(r){
        const key = projectNoOf(r)||recordIdOf(r);
        return '<div class="srow" onclick="openByPrj('+esc4(key)+')">'+
          '<div class="top"><span class="prjno">'+esc2(projectNoOf(r)||'프로젝트 미확정')+'</span>'+
          (r.공급가액?'<span class="amt">'+fmt(r.공급가액)+'원</span>':'')+'</div>'+
          '<div class="top"><span class="camp">'+(r.캠프명||'')+'</span>'+
          '<span class="camp">'+(r.완료일||r.접수일자||r.점검예정일||'')+'</span></div>'+
          (recordIdOf(r)?'<div class="meta"><span class="sid">'+esc2(recordIdOf(r))+'</span></div>':'')+'</div>';
      }).join('') + '</div>'
    : '<div class="card" style="margin-top:12px">이 자료가 있어야 대조를 할 수 있습니다.<br><b>'+
      m.todo+'</b><br><br>[실행] 탭의 <b>파일·폴더 열기</b>에서 해당 폴더를 열 수 있습니다.</div>';
	  window._mineIdx = idx;                      // 캡처가 어느 항목인지 알아야 한다
	  const cap = rows.length
	    ? '<div class="media-tools" role="toolbar" aria-label="내 확인 상세 목록 도구" style="margin-top:12px">'+
	      '<button class="icon-btn" onclick="printMineDetail()" title="인쇄" aria-label="인쇄"><img src="/icons/printer.svg" alt=""></button>'+
	      '<button class="icon-btn primary" onclick="captureMineDetail()" title="이미지 저장" aria-label="이미지 저장"><img src="/icons/image-down.svg" alt=""></button>'+
	      '<button class="icon-btn" onclick="copyMineDetailImage()" title="복사" aria-label="복사"><img src="/icons/clipboard-copy.svg" alt=""></button>'+
	      '<button class="icon-btn excel" onclick="exportMineDetailXlsx()" title="엑셀 저장" aria-label="엑셀 저장"><img src="/icons/file-spreadsheet.svg" alt=""></button></div>'
	    : '';
  openPane('<h2>'+m.cat+' <span class="chip '+(m.lv==='danger'?'c-danger':'c-warn')+'">'+
    (m.n?m.n+'건':'대기')+'</span></h2><div class="sub">'+m.todo+
    (m.amt?' · '+fmt(m.amt)+'원':'')+'</div>'+helpBox(m.cat)+cap+body);
}

/* 담당자 전달용 이미지 — 캔버스 직접 렌더(외부 리소스 0, 폰·PC 동일) */
const ASSIGNEE_CAP = 60;
async function assigneeToPng(name){
  const all = (window._curAssignee===name && Array.isArray(window._curAssigneeList))
    ? window._curAssigneeList : ((window._board||{})[name]||[]);
  const l = all.slice(0, ASSIGNEE_CAP), cut=all.length>l.length;
  const S=2, W=760, rowH=52, headH=100;
  const H = headH + 40 + Math.max(1,l.length)*rowH + (cut?24:0) + 52;
  const cv=document.createElement('canvas'); cv.width=W*S; cv.height=H*S;
  const ctx=cv.getContext('2d'); ctx.scale(S,S);
  const F=(s,w)=>(w||400)+' '+s+'px "Nanum Gothic","Malgun Gothic",sans-serif';
  const txt=(t,x,y,f,c,a)=>{ctx.font=f;ctx.fillStyle=c;ctx.textAlign=a||'left';
    ctx.fillText(String(t==null?'':t),x,y);ctx.textAlign='left';};
  const box=(x,y,w,h,fill,r)=>{ctx.beginPath();ctx.roundRect(x,y,w,h,r||8);ctx.fillStyle=fill;ctx.fill();};
  const clip=(t,max)=>{let s=String(t==null?'':t);
    if(ctx.measureText(s).width<=max)return s;
    while(s.length>1&&ctx.measureText(s+'…').width>max)s=s.slice(0,-1);return s+'…';};
  ctx.fillStyle='#fff'; ctx.fillRect(0,0,W,H);
  const g=ctx.createLinearGradient(0,0,W,headH); g.addColorStop(0,'#0E1B3F'); g.addColorStop(1,'#2452E6');
  ctx.fillStyle=g; ctx.fillRect(0,0,W,headH);
  await drawLogo(ctx,26,24,36,true);
  txt('확인 요청 · '+name, 86, 44, F(20,800), '#fff');
  txt('Coupang Service Operations System · 확인 필요 '+all.length+'건', 86, 68, F(11.5), '#9FB4E8');
  txt(new Date().toLocaleDateString('ko-KR'), W-26, 44, F(12.5,700), '#fff','right');
  let y = headH + 30;
  txt('프로젝트NO',26,y,F(10.5,800),'#667085'); txt('구분 · 캠프',176,y,F(10.5,800),'#667085');
  txt('확인 사항',432,y,F(10.5,800),'#667085'); txt('일자',W-26,y,F(10.5,800),'#667085','right');
  y+=10; ctx.strokeStyle='#D0D8E6'; ctx.lineWidth=1;
  ctx.beginPath(); ctx.moveTo(26,y); ctx.lineTo(W-26,y); ctx.stroke();
  const list = l.length?l:[{prj:'—',camp:'확인할 항목이 없습니다',kind:'',issue:'',date:''}];
  list.forEach(function(i,n){
    const t = y + n*rowH;
    if(n%2) box(26,t+4,W-52,rowH-8,'#F6F9FD',8);
    txt(i.prj||i.id||'', 26, t+26, F(13,800), '#101828');
    txt(clip((i.kind||'')+' · '+(i.camp||''), 244), 176, t+26, F(12), '#475467');
    txt(clip(i.issue||'', 236), 432, t+26, F(12,800), '#B42318');
    txt(i.date||'-', W-26, t+26, F(11.5), '#667085','right');
    if(i.desc) txt(clip(i.desc, 380), 176, t+42, F(10.5), '#98A2B3');
  });
  y += list.length*rowH + 22;
  if(cut){
    txt('※ 이미지에는 '+l.length+'건만 표시했습니다 (전체 '+all.length+'건은 앱 목록에서 확인).',
        26, y, F(11,800), '#B42318');
    y += 22;
  }
  txt('※ 처리 후 밴드에 게시하거나 회신해 주세요.', 26, y, F(11), '#98A2B3');
  txt('자동 생성 · CSOS Agent', W-26, y, F(10), '#98A2B3','right');
  return new Promise(res=>cv.toBlob(res,'image/png'));
}
/* ── '내가 확인할 사항' 캡처 ────────────────────────────────────────────
   담당자별 카드와 같은 방식(Canvas 직접 렌더)이라 웹폰트 오염이 없고 폰·PC 동일하게 나온다.
   ★ 보내는 건 **사람이 직접** 고른다 — 자동 전송은 하지 않는다.
     쿠팡 담당자에게는 어떤 메시지도 보내지 않는다(유니버셜 내부에서 처리). */
async function mineToPng(){
  const l = window._mine || [];
  const S=2, W=760, rowH=56, headH=100;
  const H = headH + 40 + Math.max(1,l.length)*rowH + 56;
  const cv=document.createElement('canvas'); cv.width=W*S; cv.height=H*S;
  const ctx=cv.getContext('2d'); ctx.scale(S,S);
  const F=(s,w)=>(w||400)+' '+s+'px "Nanum Gothic","Malgun Gothic",sans-serif';
  const txt=(t,x,y,f,c,a)=>{ctx.font=f;ctx.fillStyle=c;ctx.textAlign=a||'left';
    ctx.fillText(String(t==null?'':t),x,y);ctx.textAlign='left';};
  const box=(x,y,w,h,fill,r)=>{ctx.beginPath();ctx.roundRect(x,y,w,h,r||8);ctx.fillStyle=fill;ctx.fill();};
  const clip=(t,max)=>{let s=String(t==null?'':t);
    if(ctx.measureText(s).width<=max)return s;
    while(s.length>1&&ctx.measureText(s+'…').width>max)s=s.slice(0,-1);return s+'…';};

  ctx.fillStyle='#fff'; ctx.fillRect(0,0,W,H);
  const g=ctx.createLinearGradient(0,0,W,headH); g.addColorStop(0,'#0E1B3F'); g.addColorStop(1,'#2452E6');
  ctx.fillStyle=g; ctx.fillRect(0,0,W,headH);
  await drawLogo(ctx,26,24,36,true);
  const totN = l.reduce((a,m)=>a+(m.n||0),0), totA = l.reduce((a,m)=>a+(m.amt||0),0);
  txt('내가 확인할 사항', 86, 44, F(20,800), '#fff');
  txt('CSOS · 유현민 (관리·청구·원천데이터) · 총 '+totN+'건', 86, 68, F(11.5), '#9FB4E8');
  txt(new Date().toLocaleDateString('ko-KR'), W-26, 44, F(12.5,700), '#fff','right');

  let y = headH + 30;
  txt('항목',26,y,F(10.5,800),'#667085');
  txt('건수',W-150,y,F(10.5,800),'#667085','right');
  txt('금액',W-26,y,F(10.5,800),'#667085','right');
  y+=10; ctx.strokeStyle='#D0D8E6'; ctx.lineWidth=1;
  ctx.beginPath(); ctx.moveTo(26,y); ctx.lineTo(W-26,y); ctx.stroke();

  const list = l.length?l:[{cat:'확인할 사항이 없습니다', todo:'', n:0, amt:0, lv:''}];
  list.forEach(function(m,n){
    const t = y + n*rowH;
    if(n%2) box(26,t+4,W-52,rowH-8,'#F6F9FD',8);
    const col = m.lv==='danger' ? '#B42318' : m.lv==='warn' ? '#A85A06' : '#667085';
    txt(clip(m.cat, 480), 26, t+26, F(13.5,800), '#101828');
    if(m.todo) txt(clip(m.todo, 500), 26, t+43, F(10.5), '#98A2B3');
    txt(m.n ? m.n+'건' : '투입 대기', W-150, t+30, F(13,800), col, 'right');
    if(m.amt) txt(fmt(m.amt)+'원', W-26, t+30, F(12,700), '#475467','right');
  });
  y += list.length*rowH + 24;
  if(totA) txt('금액 합계 '+fmt(totA)+'원', 26, y, F(11.5,800), '#475467');
  txt('유니버셜 내부 확인용 · 자동 생성 CSOS Agent', W-26, y, F(10), '#98A2B3','right');
  return new Promise(res=>cv.toBlob(res,'image/png'));
}

/* ── 항목 하나의 **건별 목록** 캡처 ──
   요약 캡처(mineToPng)는 '어떤 항목이 몇 건'만 보여준다. 실제로 처리하려면
   '어느 프로젝트·어느 캠프·얼마'가 필요해서 건별 목록을 따로 그린다.
   건수가 많으면 다 그리지 않고 **잘랐다고 이미지에 적는다** — 조용히 자르면
   받는 사람이 그게 전부인 줄 안다. */
const MINE_CAP = 60;

async function mineDetailToPng(idx){
  const m = (window._mine||[])[idx];
  if(!m) throw new Error('항목 없음');
  const all = m.items || [];
  const l = all.slice(0, MINE_CAP);
  const S=2, W=760, rowH=46, headH=100;
  const cut = all.length > l.length;
  const H = headH + 40 + Math.max(1,l.length)*rowH + (cut?26:0) + 56;
  const cv=document.createElement('canvas'); cv.width=W*S; cv.height=H*S;
  const ctx=cv.getContext('2d'); ctx.scale(S,S);
  const F=(s,w)=>(w||400)+' '+s+'px "Nanum Gothic","Malgun Gothic",sans-serif';
  const txt=(t,x,y,f,c,a)=>{ctx.font=f;ctx.fillStyle=c;ctx.textAlign=a||'left';
    ctx.fillText(String(t==null?'':t),x,y);ctx.textAlign='left';};
  const box=(x,y,w,h,fill,r)=>{ctx.beginPath();ctx.roundRect(x,y,w,h,r||8);ctx.fillStyle=fill;ctx.fill();};
  const clip=(t,max)=>{let s=String(t==null?'':t);
    if(ctx.measureText(s).width<=max)return s;
    while(s.length>1&&ctx.measureText(s+'…').width>max)s=s.slice(0,-1);return s+'…';};

  ctx.fillStyle='#fff'; ctx.fillRect(0,0,W,H);
  const g=ctx.createLinearGradient(0,0,W,headH); g.addColorStop(0,'#0E1B3F'); g.addColorStop(1,'#2452E6');
  ctx.fillStyle=g; ctx.fillRect(0,0,W,headH);
  await drawLogo(ctx,26,24,36,true);
  txt(clip(m.cat, 520), 86, 44, F(20,800), '#fff');
  txt((m.todo||'') + ' · ' + (m.n||all.length) + '건' + (m.amt? ' · '+fmt(m.amt)+'원' : ''),
      86, 68, F(11.5), '#9FB4E8');
  txt(new Date().toLocaleDateString('ko-KR'), W-26, 44, F(12.5,700), '#fff','right');

  let y = headH + 30;
  txt('프로젝트NO / ID',26,y,F(10.5,800),'#667085');
  txt('캠프',232,y,F(10.5,800),'#667085');
  txt('금액',W-140,y,F(10.5,800),'#667085','right');
  txt('일자',W-26,y,F(10.5,800),'#667085','right');
  y+=10; ctx.strokeStyle='#D0D8E6'; ctx.lineWidth=1;
  ctx.beginPath(); ctx.moveTo(26,y); ctx.lineTo(W-26,y); ctx.stroke();

  const rows = l.length?l:[{}];
  rows.forEach(function(r,n){
    const t = y + n*rowH;
    if(n%2) box(26,t+4,W-52,rowH-8,'#F6F9FD',8);
    const key = projectNoOf(r)||'프로젝트 미확정';
    const dt  = r.완료일||r.접수일자||r.점검예정일||r.작업완료일||'-';
    txt(clip(key,190), 26, t+28, F(12.5,800), '#101828');
    txt(clip(r.캠프명||'', 300), 232, t+28, F(12), '#475467');
    if(r.공급가액) txt(fmt(r.공급가액)+'원', W-140, t+28, F(12,700), '#101828','right');
    txt(dt, W-26, t+28, F(11), '#667085','right');
  });
  y += rows.length*rowH + 18;
  if(cut){
    txt('※ 이 이미지에는 '+l.length+'건만 담겼습니다 (전체 '+all.length+'건) — 나머지는 앱에서 확인',
        26, y, F(11,800), '#B42318');
    y += 22;
  }
  txt('유니버셜 내부 확인용 · 자동 생성 CSOS Agent', W-26, y, F(10), '#98A2B3','right');
  return new Promise(res=>cv.toBlob(res,'image/png'));
}

function mineDetailFileName(){
  const m = (window._mine||[])[window._mineIdx] || {};
  return 'CSOS_'+String(m.cat||'목록').replace(/[\/:*?"<>|\s]/g,'') +
         '_'+new Date().toISOString().slice(0,10)+'.png';
}
async function captureMineDetail(){
  let b;
  try{ b = await mineDetailToPng(window._mineIdx); if(!b) throw new Error('빈 이미지'); }
  catch(e){ alert('이미지 생성 실패: '+e); return; }
  saveOrOpen(b, mineDetailFileName());
}
async function printMineDetail(){
  try{printPngBlob(await mineDetailToPng(window._mineIdx),mineDetailFileName());}
  catch(e){alert('인쇄 이미지 생성 실패: '+e);}
}
async function copyMineDetailImage(){
  try{await copyPngBlob(await mineDetailToPng(window._mineIdx));}
  catch(e){alert('이미지 복사 실패: '+e);}
}
function mineDetailRows(){
  const m=(window._mine||[])[window._mineIdx]||{};
  return (m.items||[]).map(r=>({...r,업무구분:r.업무구분||m.cat,문제:r.문제||m.todo}));
}
function exportMineDetailXlsx(){
  const m=(window._mine||[])[window._mineIdx]||{};
  exportRowsXlsx(m.cat||'내 확인 상세',mineDetailRows());
}
async function shareMineDetail(){
  let b;
  try{ b = await mineDetailToPng(window._mineIdx); if(!b) throw new Error('empty'); }
  catch(e){ alert('이미지 생성 실패: '+e); return; }
  const m = (window._mine||[])[window._mineIdx] || {};
  const f = new File([b], mineDetailFileName(), {type:'image/png'});
  try{
    // 받는 사람은 **사람이 직접** 고른다(자동 전송 아님)
    if(navigator.canShare && navigator.canShare({files:[f]})){
      await navigator.share({files:[f], title:m.cat||'확인 목록',
                             text:'[CSOS] '+(m.cat||'')+' '+(m.n||'')+'건'});
      return;
    }
  }catch(e){ if(e.name==='AbortError') return; }
  saveOrOpen(b, mineDetailFileName());
}

function mineFileName(){
  return 'CSOS_내가확인할사항_'+new Date().toISOString().slice(0,10)+'.png';
}
async function captureMine(){
  let b;
  try{ b = await mineToPng(); if(!b) throw new Error('빈 이미지'); }
  catch(e){ alert('이미지 생성 실패: '+e); return; }
  saveOrOpen(b, mineFileName());
}
async function printMine(){
  try{printPngBlob(await mineToPng(),mineFileName());}
  catch(e){alert('인쇄 이미지 생성 실패: '+e);}
}
async function copyMineImage(){
  try{await copyPngBlob(await mineToPng());}
  catch(e){alert('이미지 복사 실패: '+e);}
}
function mineRows(){
  return (window._mine||[]).flatMap(m=>(m.items||[]).map(r=>(
    {...r,업무구분:r.업무구분||m.cat,문제:r.문제||m.todo}
  )));
}
function exportMineXlsx(){exportRowsXlsx('내가 확인할 사항',mineRows());}
async function shareMine(){
  let b;
  try{ b = await mineToPng(); if(!b) throw new Error('empty'); }
  catch(e){ alert('이미지 생성 실패: '+e); return; }
  const f = new File([b], mineFileName(), {type:'image/png'});
  const n = (window._mine||[]).reduce((a,m)=>a+(m.n||0),0);
  try{
    // 받는 사람은 **이 창에서 사람이 직접** 고른다(자동 전송 아님)
    if(navigator.canShare && navigator.canShare({files:[f]})){
      await navigator.share({files:[f], title:'내가 확인할 사항',
                             text:'[CSOS] 내가 확인할 사항 '+n+'건'});
      return;
    }
  }catch(e){ if(e.name==='AbortError') return; }
  saveOrOpen(b, mineFileName());     // 공유를 못 쓰는 브라우저는 저장으로
}

function assigneeFileName(){
  return 'CSOS_확인요청_'+(window._curAssignee||'담당자')+'_'+new Date().toISOString().slice(0,10)+'.png';
}
async function metricToPng(label){
  const d=metricDetail(label), all=d.rows||[], rows=all.slice(0,METRIC_CAP), cut=all.length>rows.length;
  const S=2, W=760, rowH=52, headH=100;
  const H=headH+40+Math.max(1,rows.length)*rowH+(cut?24:0)+74;
  const cv=document.createElement('canvas'); cv.width=W*S; cv.height=H*S;
  const ctx=cv.getContext('2d'); ctx.scale(S,S);
  const F=(s,w)=>(w||400)+' '+s+'px "Nanum Gothic","Malgun Gothic",sans-serif';
  const txt=(t,x,y,f,c,a)=>{ctx.font=f;ctx.fillStyle=c;ctx.textAlign=a||'left';
    ctx.fillText(String(t==null?'':t),x,y);ctx.textAlign='left';};
  const box=(x,y,w,h,fill,r)=>{ctx.beginPath();ctx.roundRect(x,y,w,h,r||8);ctx.fillStyle=fill;ctx.fill();};
  const clip=(t,max)=>{let s=String(t==null?'':t);if(ctx.measureText(s).width<=max)return s;
    while(s.length>1&&ctx.measureText(s+'…').width>max)s=s.slice(0,-1);return s+'…';};
  ctx.fillStyle='#fff';ctx.fillRect(0,0,W,H);
  const g=ctx.createLinearGradient(0,0,W,headH);g.addColorStop(0,'#0E1B3F');g.addColorStop(1,'#2452E6');
  ctx.fillStyle=g;ctx.fillRect(0,0,W,headH);
  await drawLogo(ctx,26,24,36,true);
  txt(clip(label,520),86,44,F(20,800),'#fff');
  const pc=d.project_count!=null?d.project_count+'개 프로젝트 · ':'';
  txt('2026년 기준 · '+pc+all.length+'건'+((+d.amount||0)?' · '+fmt(Math.abs(+d.amount))+'원':''),
      86,68,F(11.5),'#9FB4E8');
  txt(new Date().toLocaleDateString('ko-KR'),W-26,44,F(12.5,700),'#fff','right');
  let y=headH+30;
  txt('프로젝트NO · 프로젝트명 / 캠프 · 담당자',26,y,F(10.5,800),'#667085');
  txt('금액·상태',W-150,y,F(10.5,800),'#667085','right');
  txt('일자',W-26,y,F(10.5,800),'#667085','right');
  y+=10;ctx.strokeStyle='#D0D8E6';ctx.beginPath();ctx.moveTo(26,y);ctx.lineTo(W-26,y);ctx.stroke();
  const list=rows.length?rows:[{프로젝트NO:'—',프로젝트명:'해당 건 없음',캠프명:'',일자:'-'}];
  list.forEach((r,n)=>{
    const t=y+n*rowH;if(n%2)box(26,t+4,W-52,rowH-8,'#F6F9FD',8);
    txt(clip(projectLabel(r)+' · '+(r.프로젝트명||''),340),26,t+23,F(12.5,800),'#101828');
    txt(clip((r.캠프명||'')+(r.담당자?' · '+r.담당자:''),340),26,t+41,F(11),'#667085');
    const value=(+r.금액||0)?fmt(Math.abs(+r.금액))+'원':(r.상태||r.출처||'');
    txt(clip(value,170),W-150,t+29,F(11.5,700),'#475467','right');
    txt(r.일자||'-',W-26,t+29,F(10.5),'#667085','right');
  });
  y+=list.length*rowH+18;
  if(cut){txt('※ 이미지에는 '+rows.length+'건만 표시 (전체 '+all.length+'건은 앱에서 확인).',
      26,y,F(11,800),'#B42318');y+=22;}
  txt(clip('집계 기준: '+(d.basis||''),680),26,y,F(10.5),'#667085');
  txt('자동 생성 · CSOS Agent',W-26,y+22,F(10),'#98A2B3','right');
  return new Promise(res=>cv.toBlob(res,'image/png'));
}
function metricFileName(){
  return 'CSOS_'+String(window._curMetricLabel||'목록').replace(/[\/:*?"<>|\s·()]/g,'')+
    '_'+new Date().toISOString().slice(0,10)+'.png';
}
async function saveExecMetric(){
  let b;try{b=await metricToPng(window._curMetricLabel);if(!b)throw new Error('빈 이미지');}
  catch(e){alert('이미지 생성 실패: '+e);return;}
  saveOrOpen(b,metricFileName());
}
async function shareExecMetric(){
  let b;try{b=await metricToPng(window._curMetricLabel);if(!b)throw new Error('빈 이미지');}
  catch(e){alert('이미지 생성 실패: '+e);return;}
  const label=window._curMetricLabel||'확인 목록';
  const f=new File([b],metricFileName(),{type:'image/png'});
  try{
    if(navigator.canShare&&navigator.canShare({files:[f]})){
      await navigator.share({files:[f],title:label,text:'[CSOS] '+label+' 확인 목록'});return;
    }
  }catch(e){if(e.name==='AbortError')return;}
  saveOrOpen(b,metricFileName());
}
async function captureAssignee(){
  let b;
  try{ b = await assigneeToPng(window._curAssignee); if(!b) throw new Error('빈 이미지'); }
  catch(e){ alert('이미지 생성 실패: '+e); return; }
  saveOrOpen(b, assigneeFileName());
}
async function shareAssignee(){
  let b;
  try{ b = await assigneeToPng(window._curAssignee); if(!b) throw new Error('empty'); }
  catch(e){ alert('이미지 생성 실패: '+e); return; }
  const f = new File([b], assigneeFileName(), {type:'image/png'});
  const name = window._curAssignee||'';
  const n = (window._curAssigneeList||[]).length;
  try{
    if(navigator.canShare && navigator.canShare({files:[f]})){
      await navigator.share({files:[f], title:'확인 요청 · '+name,
                             text:'[CSOS] '+name+'님 확인 요청 '+n+'건'});
      return;
    }
  }catch(e){ if(e.name==='AbortError') return; }
  saveOrOpen(b, assigneeFileName());
}

/* 캡처 이미지용 CSOS 마크 — 앱 아이콘과 같은 도형을 캔버스로 직접 그린다(외부 리소스 0) */
/* 캡처 이미지에 고객사 CI를 그린다(있을 때). 같은 출처 이미지라 canvas 오염 없음.
   실패하면 기본 마크로 자동 대체 — 보고 이미지가 비는 일이 없게. */
let _brandImg = null;
function loadBrandImg(){
  if(!brandLogo) return Promise.resolve(null);
  if(_brandImg) return Promise.resolve(_brandImg);
  return new Promise(res=>{
    const im = new Image();
    im.onload = ()=>{ _brandImg = im; res(im); };
    im.onerror = ()=>res(null);
    im.src = '/brand/' + encodeURIComponent(brandLogo);
  });
}
let _appIconImg = null;
function loadAppIconImg(){
  if(_appIconImg) return Promise.resolve(_appIconImg);
  return new Promise(res=>{
    const im = new Image();
    im.onload = ()=>{ _appIconImg = im; res(im); };
    im.onerror = ()=>res(null);
    im.src = '/icon-512.png?v=csos-20260730';
  });
}
async function drawLogo(ctx, x, y, size, onDark){
  const im = await loadAppIconImg();
  if(!im){ drawMark(ctx, x, y, size); return; }
  ctx.save();
  ctx.drawImage(im, x, y, size, size);
  ctx.restore();
}
async function drawUniversalLogo(ctx, x, y, maxW, maxH){
  const im = await loadBrandImg();
  if(!im) return;
  const scale = Math.min(maxW / im.width, maxH / im.height);
  const w = im.width * scale, h = im.height * scale;
  ctx.save();
  ctx.globalAlpha = .9;
  ctx.drawImage(im, x, y + (maxH-h)/2, w, h);
  ctx.restore();
}
function drawMark(ctx,x,y,size){
  const p = v => v*size/512;
  ctx.save(); ctx.translate(x,y);
  ctx.strokeStyle='#fff'; ctx.lineWidth=p(56); ctx.lineCap='round'; ctx.lineJoin='round';
  ctx.beginPath(); ctx.moveTo(p(122),p(262)); ctx.lineTo(p(214),p(352)); ctx.lineTo(p(396),p(116)); ctx.stroke();
  ctx.fillStyle='#fff';
  ctx.beginPath(); ctx.roundRect(p(136),p(398),p(240),p(28),p(14)); ctx.fill();
  ctx.globalAlpha=.42;
  ctx.beginPath(); ctx.roundRect(p(182),p(444),p(148),p(18),p(9)); ctx.fill();
  ctx.restore();
}
/* Canvas 직접 렌더 — 외부 리소스를 쓰지 않아 taint(오염)가 없고 PC·태블릿·폰에서 동일하게 동작.
   (이전 SVG foreignObject 방식은 웹폰트 때문에 canvas가 오염되어 toBlob이 null을 반환했다) */
async function reportToPng(){
  const D = _rptData || {};
  const S = 2, W = 860;                       // 논리 폭 860, 2배 해상도
  const F = (sz,w) => `${w||400} ${sz}px "Nanum Gothic","Malgun Gothic",sans-serif`;
  const cv = document.createElement('canvas');
  const ctx = cv.getContext('2d');
  // ── 높이 선계산 (대표보고 섹션 포함)
  const tiles = r => Math.ceil(r/3)*74;
  const execSections=(D.exec||[]).filter(s=>!/^(정리|요약)$/.test(cleanExecTitle(s.title)));
  const dailyIssueList=D.dailyIssues||[];
  const execSecH = execSections.reduce((a,s)=>{
    const gh = (s.groups||[]).length
      ? 38 + Math.max(...s.groups.map(g=>g.items.length))*24 + 18 : 0;
    const issueH=/당일 업무 실적/.test(cleanExecTitle(s.title)) && dailyIssueList.length
      ? 42 + dailyIssueList.length*30 : 0;
    return a + 34 + gh + (s.items.length?tiles(s.items.length):0)
             + issueH + (s.lines.length? s.lines.length*20 + 6 : 0);
  }, 0);
  // ★ 타일 개수를 하드코딩(6)하면 지표가 7개일 때 마지막 타일이 캔버스 밖으로 잘린다
  const erpM = D.erpMonths || [];
  const erpH = erpM.length ? 34 + (erpM.length+1)*28 + 20 : 0;
  const pmProgress = D.pmProgress || [];
  const wlAs = ((D.workLog||{})['돌발AS']) || {};
  const wlReasons = wlAs.미처리사유 || [];
  const wlOpenRows = wlAs.미처리목록 || [];
  const wlCancelRows = wlAs.취소목록 || [];
  const wlUnknownRows = wlAs.처리완료일확인목록 || [];
  const dailyBrief = D.dailyBrief || null;
  const dailyActivities = (dailyBrief&&dailyBrief.activities) || [];
  const pmProgressH = pmProgress.length ? 34 + tiles(pmProgress.length) + 20 : 0;
  const workDetailH = rows => rows.length ? 26 + rows.length*50 + 6 : 0;
  const workLogH = Object.keys(wlAs).length
    ? 34 + tiles(4) + (wlReasons.length ? 18 + wlReasons.length*20 : 0)
      + workDetailH(wlOpenRows) + workDetailH(wlCancelRows)
      + workDetailH(wlUnknownRows) + 16 : 0;
  const dailyBriefH = dailyBrief
    ? tiles((dailyBrief.metrics||[]).length) + (dailyActivities.length ? dailyActivities.length*90 : 28) + 10
    : 0;
  const quarterH=(D.quarter||[]).length ? 34+tiles(D.quarter.length) : 0;
  // 화면의 4원천 표와 같은 구조로 캡처한다. 미확인 번호는 한 줄 문자열이 아니라
  // 3열 칩으로 배치하므로 번호 수에 따라 행 높이도 함께 늘어나야 잘리지 않는다.
  const sourceDetails = (D.srcDetails||[]).length ? D.srcDetails : (D.srcs||[]).map(r=>({
    label:r[0], result:r[1],
    missProjects:String(r[2]||'').split(/\s*,\s*/).filter(v=>v&&v!=='-'),
    empty:false
  }));
  const sourceRowHeight = r => Math.max(44, 20 + Math.ceil(Math.max(1,(r.missProjects||[]).length)/3)*32);
  const sourceTableH = 32 + sourceDetails.reduce((n,r)=>n+sourceRowHeight(r),0);
  let H = 108 + 34 + tiles((D.month||[]).length) + quarterH
          + 34 + tiles((D.kpi||[]).length) + (D.kpiNote ? 18 : 0)
          + execSecH
          + dailyBriefH + pmProgressH + workLogH
          + erpH
          + 34 + (D.issues.length+1)*30 + 34 + sourceTableH + 54;
  cv.width = W*S; cv.height = H*S;
  ctx.scale(S,S);
  ctx.textBaseline = 'middle';
  const txt = (s,x,y,font,color,align)=>{ ctx.font=font; ctx.fillStyle=color||'#101828';
    ctx.textAlign=align||'left'; ctx.fillText(s==null?'':String(s),x,y); ctx.textAlign='left'; };
  const box = (x,y,w,h,fill,stroke,r=10)=>{ ctx.beginPath();
    if(ctx.roundRect) ctx.roundRect(x,y,w,h,r); else ctx.rect(x,y,w,h);
    if(fill){ctx.fillStyle=fill;ctx.fill();} if(stroke){ctx.strokeStyle=stroke;ctx.lineWidth=1;ctx.stroke();} };
  const clip = (s,max)=>{ s=String(s==null?'':s); ctx.font=F(12);
    if(ctx.measureText(s).width<=max) return s;
    while(s.length>1 && ctx.measureText(s+'…').width>max) s=s.slice(0,-1);
    return s+'…'; };

  ctx.fillStyle='#fff'; ctx.fillRect(0,0,W,H);
  // ── 헤더
  const g = ctx.createLinearGradient(0,0,W,96);
  g.addColorStop(0,'#0E1B3F'); g.addColorStop(1,'#2452E6');
  ctx.fillStyle=g; ctx.fillRect(0,0,W,96);
  if(!brandLogo) box(30,26,44,44,'rgba(255,255,255,.16)',null,12);   // CI가 있으면 배경 사각형 없이
  await drawLogo(ctx,30,30,36,true);
  await drawUniversalLogo(ctx,500,20,170,28);
  txt('Coupang Service Operations System',92,36,F(17,800),'#fff');
  txt('일일보고 · (주)유니버셜리프트앤히타치코리아',92,60,F(11.5),'#9FB4E8');
  txt('보고일',W-30,32,F(11),'#9FB4E8','right');
  txt(D.date,W-30,52,F(17,800),'#fff','right');
  txt(`집계기준일 ${(D.meta&&D.meta['집계기준일'])||'-'} · 보고자 ${(D.meta&&D.meta['보고자'])||'-'}`,
      W-30,72,F(10),'#9FB4E8','right');
  txt(`데이터 업데이트 ${D.updatedAt||'-'}`,92,82,F(9.5,700),'#C9D7FA');

  let y = 96;
  const section = (t,no=null,tone='#1B41BC') => {
    y += 38;
    if(no!=null){
      box(30,y-25,26,26,tone,null,8);
      txt(no,43,y-12,F(12,900),'#FFF','center');
      txt(t,66,y-12,F(14,900),'#12214B');
      ctx.strokeStyle='#D9E1EF';ctx.lineWidth=2;ctx.beginPath();
      ctx.moveTo(66+ctx.measureText(t).width+14,y-12);ctx.lineTo(W-30,y-12);ctx.stroke();
    }else{
      box(30,y-25,7,26,tone,null,4);
      txt(t,48,y-12,F(13.5,900),tone);
      ctx.strokeStyle='#E4E9F0';ctx.lineWidth=1;ctx.beginPath();
      ctx.moveTo(48+ctx.measureText(t).width+14,y-12);ctx.lineTo(W-30,y-12);ctx.stroke();
    }
  };
  const grid = arr => { const cw=(W-60-16)/3;
    arr.forEach((t,i)=>{ const cx=30+(i%3)*(cw+8), cy=y+Math.floor(i/3)*74;
      const accents=['#4960DF','#2B9560','#D08A2E'], accent=accents[i%3];
      box(cx,cy,cw,64,'#FAFBFE','#DFE6F2');
      box(cx,cy,cw,4,accent,null,4);
      txt(t[1],cx+14,cy+22,F(11,800),'#667085');
      txt(t[0],cx+14,cy+44,F(19,800),t[2]||'#101828'); });
    y += tiles(arr.length); };
  const table = (cols,rows,widths)=>{
    const xs=[]; let cx=30; widths.forEach(w=>{xs.push(cx); cx+=w;});
    box(30,y,W-60,28,'#F1F5FB','#E4E9F0',8);
    cols.forEach((c,i)=>txt(c,xs[i]+10,y+14,F(11,800),'#475467'));
    y+=28;
    rows.forEach((r,ri)=>{ if(ri%2) { ctx.fillStyle='#FAFCFF'; ctx.fillRect(30,y,W-60,30); }
      r.forEach((c,i)=>txt(clip(c,widths[i]-20),xs[i]+10,y+15,F(12,i===0?700:400),i===0?'#101828':'#475467'));
      ctx.strokeStyle='#EEF2F7'; ctx.beginPath(); ctx.moveTo(30,y+30); ctx.lineTo(W-30,y+30); ctx.stroke();
      y+=30; });
  };
  const sourceTable = rows=>{
    const widths=[130,330,340], xs=[30,160,490], headH=32;
    const rule=(x1,y1,x2,y2,color='#DCE3EF')=>{
      ctx.strokeStyle=color;ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
    };
    box(30,y,W-60,headH,'#F1F5FB','#DCE3EF',0);
    ['원천','결과','미확인 프로젝트NO'].forEach((c,i)=>
      txt(c,xs[i]+12,y+headH/2,F(11.5,800),'#344054'));
    rule(xs[1],y,xs[1],y+headH); rule(xs[2],y,xs[2],y+headH);
    y+=headH;
    rows.forEach((r,ri)=>{
      const rh=sourceRowHeight(r), mid=y+rh/2;
      ctx.fillStyle=ri%2?'#FCFDFF':'#FFF';ctx.fillRect(30,y,W-60,rh);
      rule(30,y,W-30,y); rule(30,y+rh,W-30,y+rh);
      rule(30,y,30,y+rh); rule(W-30,y,W-30,y+rh); rule(xs[1],y,xs[1],y+rh);
      txt(r.label,xs[0]+12,mid,F(12,500),'#101828');
      if(r.empty){
        // 자료가 없을 때는 첨부 화면처럼 결과·번호 칸을 합쳐 안내를 보여 준다.
        txt(clip(r.result,W-xs[1]-54),xs[1]+12,mid,F(11.5),'#98A2B3');
      }else{
        rule(xs[2],y,xs[2],y+rh);
        const m=String(r.result||'').match(/^(대상 \d+건 중 )(\d+건 확인\(\d+%\))(.*)$/);
        if(m){
          let tx=xs[1]+12;
          [[m[1],400,'#101828'],[m[2],800,'#101828'],[m[3],400,'#101828']].forEach(p=>{
            ctx.font=F(11.5,p[1]);txt(p[0],tx,mid,ctx.font,p[2]);tx+=ctx.measureText(p[0]).width;
          });
        }else txt(clip(r.result,widths[1]-24),xs[1]+12,mid,F(11.5),'#101828');
        const ps=(r.missProjects||[]).filter(Boolean);
        if(!ps.length){
          txt('-',xs[2]+12,mid,F(11.5),'#475467');
        }else{
          const gap=6, chipW=(widths[2]-24-gap*2)/3;
          ps.forEach((p,i)=>{
            const cx=xs[2]+12+(i%3)*(chipW+gap);
            const cy=y+10+Math.floor(i/3)*32;
            box(cx,cy,chipW,26,'#EEF2FF','#D5DDFC',8);
            txt(clip(p,chipW-16),cx+chipW/2,cy+13,F(11.5,800),'#3046C5','center');
          });
        }
      }
      y+=rh;
    });
  };

  // ── 통계는 현재 월 → 이번 분기 → 해당 연도 순서로 고정한다.
  const lines = arr => { arr.forEach(t=>{ txt(clip(t, W-70), 34, y+8, F(12), '#475467'); y += 20; }); y += 6; };
  section(D.monthTitle||'현재 월 통계',null,'#4960DF'); grid(D.month||[]);
  if((D.quarter||[]).length){ section(D.quarterTitle||'이번 분기 통계',null,'#2B9560'); grid(D.quarter); }
  section(D.kpiTitle||`${APP_YEAR}년 누적`,null,'#9A5A12'); grid(D.kpi||[]);
  if(D.kpiNote){ txt(D.kpiNote,30,y+8,F(10),'#98A2B3');y+=18; }

  // ── 번호가 붙는 대표 보고 섹션. 빈 '정리'는 버리고 화면 순서대로 1부터 다시 센다.
  const groupCols = gs => {
    const n = Math.min(gs.length, 3), gw = (W - 60 - (n-1)*8) / n;
    const rowsMax = Math.max(...gs.map(g=>g.items.length));
    const bh = 38 + rowsMax*24 + 12;
    const pal=[
      {head:'#EEF2FF',line:'#CAD4FF',text:'#3448BE',accent:'#4960DF'},
      {head:'#ECF9F1',line:'#BCE6CD',text:'#177245',accent:'#2B9560'},
      {head:'#FFF6E8',line:'#F2D19B',text:'#9A5A12',accent:'#D08A2E'}
    ];
    gs.slice(0,n).forEach((g,i)=>{
      const gx=30+i*(gw+8), p=pal[i%pal.length];
      box(gx,y,gw,bh,'#FFF',p.line,12);
      box(gx,y,gw,34,p.head,p.line,12);
      box(gx+10,y+9,5,16,p.accent,null,3);
      txt(g.name,gx+23,y+17,F(11.5,900),p.text);
      g.items.forEach(([l,v],j)=>{
        const ry=y+34+14+j*24;
        if(j%2){ctx.fillStyle='#FAFBFD';ctx.fillRect(gx+7,ry-10,gw-14,21);}
        txt(clip(dateLabel(l),gw-75),gx+12,ry,F(11.2),'#475467');
        txt(v||'-',gx+gw-12,ry,F(12.2,900),'#101828','right');
      });
    });
    y += bh + 6;
  };
  const issueBlock=items=>{
    if(!items.length)return;
    const total=items.reduce((n,x)=>n+(parseInt(x[0])||0),0), h=42+items.length*30;
    box(30,y,W-60,h,total?'#FFF8F8':'#F3FBF6',total?'#F1C2C2':'#BCE6CD',12);
    box(30,y,W-60,36,total?'#FFF1F1':'#ECF9F1',total?'#F1C2C2':'#BCE6CD',12);
    box(42,y+10,5,16,total?'#B42318':'#177245',null,3);
    txt(`이슈사항 · ${total?`${total}건 확인 필요`:'특이사항 없음'}`,55,y+18,F(11.8,900),
        total?'#A63737':'#177245');
    items.forEach((x,i)=>{
      const ry=y+50+i*30;
      if(i){ctx.strokeStyle='#F0E5E5';ctx.beginPath();ctx.moveTo(42,ry-15);ctx.lineTo(W-42,ry-15);ctx.stroke();}
      txt(x[1],44,ry,F(11.5,700),'#475467');
      txt(x[3]||'',280,ry,F(10.5),'#98A2B3');
      txt(x[0],W-44,ry,F(12,900),x[2]||'#101828','right');
    });
    y+=h+8;
  };
  const drawDailyBrief=()=>{
    if(!dailyBrief)return;
    if((dailyBrief.metrics||[]).length)grid(dailyBrief.metrics);
    if(dailyActivities.length){
      dailyActivities.forEach(a=>{
        box(30,y,W-60,82,'#F8FAFD','#DCE5F5',10);
        box(30,y,6,82,'#4960DF',null,4);
        txt(clip(`${a.project} · ${a.camp}`,560),46,y+17,F(12.5,800),'#101828');
        txt(a.date||'-',W-44,y+17,F(11.5,700),'#475467','right');
        txt(clip(`${a.poster||'게시자 미기입'} → ${a.handler||'처리자 미기입'}`,720),
            46,y+37,F(11.5,700),'#1B41BC');
        const when=a.requestDate?`접수 ${a.requestDate}`:'접수일 미기입';
        txt(clip(`요청 · ${a.request||'신청내용 미기입'} (${when})`,370),46,y+61,F(10.8),'#475467');
        txt(clip(`처리 · ${a.action||'처리내용 미기입'}`,350),430,y+61,F(10.8,700),'#12813F');
        y+=90;
      });
    }
    y+=10;
  };
  const drawExec=[...execSections];
  if(!drawExec.some(s=>/당일 업무 실적/.test(cleanExecTitle(s.title))))
    drawExec.unshift({title:'당일 업무 실적',items:[],groups:[],lines:[]});
  let reportNo=0;
  drawExec.forEach(s=>{
    const raw=cleanExecTitle(s.title), title=execPeriodTitle(raw);
    section(title,++reportNo);
    if((s.groups||[]).length) groupCols(s.groups);
    if(/당일 업무 실적/.test(raw)){
      issueBlock(dailyIssueList);
      drawDailyBrief();
    }
    if(s.items.length) grid(s.items.map(([l,v])=>[v||'-', l]));
    if(s.lines.length) lines(s.lines);
  });

  // 화면에서 고른 기간(7~9월/상반기/연간)을 그대로 넣는다. 대표 캡처만 옛 분기
  // 고정 숫자를 쓰면 화면과 다른 보고가 되므로, 위 _rptData의 pmStats 결과만 사용한다.
  if(pmProgress.length){
    section('정기점검 진행률 · 선택 기간');
    grid(pmProgress);
    txt('화면에서 선택한 기간 기준 · 예정/실행/미실행 모두 프로젝트별 상세 목록으로 확인 가능',
        30, y+8, F(10), '#98A2B3');
    y += 20;
  }

  // 유수비 대표 요청: 돌발AS는 단순 진행중 숫자가 아니라 원본 일지 기준으로
  // 발생·처리·미처리·취소와 미처리 사유까지 한 화면/한 이미지에서 확인한다.
  if(Object.keys(wlAs).length){
    const rangeStart = wlAs.기준시작일 || '-';
    const rangeEnd = wlAs.기준종료일 || '-';
    section(`돌발AS 현장 일지 대조 · ${rangeStart} ~ ${rangeEnd}`);
    grid([
      [String(wlAs.발생||0)+'건', '발생'],
      [String(wlAs.처리완료||0)+'건', '처리완료', '#12813F'],
      [String(wlAs.미처리||0)+'건', '미처리', (wlAs.미처리||0)?'#B54708':'#12813F'],
      [String(wlAs.취소||0)+'건', '취소·정상작동'],
    ]);
    if(wlReasons.length){
      txt('미처리 사유', 30, y+8, F(11,800), '#475467'); y += 18;
      wlReasons.forEach(x=>{ txt(`${x.사유} ${x.건수}건`, 40, y+8, F(11.5), '#475467'); y += 20; });
    }
    const detailRows = (title, rows, accent) => {
      if(!rows.length) return;
      txt(`${title} · ${rows.length}건`, 30, y+10, F(11.5,800), accent); y += 26;
      rows.forEach(r=>{
        box(30,y,W-60,44,'#F8FAFD','#E4E9F0',8);
        txt(clip(`${r.프로젝트NO||'프로젝트 미확정'} · ${r.캠프명||'-'}`, 500),
            42,y+14,F(11.5,800),'#101828');
        txt(`${r.일자||'-'} · ${r.원본상태||r.상태||'-'}`,
            W-42,y+14,F(10.5,700),'#667085','right');
        const why = r.미처리사유 || r.실제조치 || r.요청내용 || r.비고 || '세부 내용 미기입';
        const owner = r.담당자 ? `담당 ${r.담당자} · ` : '';
        txt(clip(`${owner}${why}`, W-92),42,y+33,F(10.5),'#475467');
        y += 50;
      });
      y += 6;
    };
    detailRows('미실시 상세', wlOpenRows, '#B54708');
    detailRows('취소·정상작동 상세', wlCancelRows, '#8B5E16');
    detailRows('처리완료일 확인 상세', wlUnknownRows, '#1B41BC');
    if(wlAs.처리완료일확인){
      txt(`※ 위 ${wlAs.처리완료일확인}건은 원본 완료일 확인 전까지 임의 완료 처리하지 않습니다.`,
          30, y+8, F(10), '#98A2B3'); y += 16;
    }
  }

  if(erpM.length){
    section(D.erpTitle || 'ERP 매출');
    const cw2 = [90, 200, 80, 300];
    const hx = [30, 120, 320, 400];
    box(30,y-2,W-60,26,'#F1F5FB',null,6);
    txt('월', hx[0]+10, y+11, F(10.5,800), '#475467');
    txt('공급가액(원)', hx[1]+cw2[1]-10, y+11, F(10.5,800), '#475467', 'right');
    txt('장수', hx[2]+cw2[2]-10, y+11, F(10.5,800), '#475467', 'right');
    txt('비중', hx[3]+10, y+11, F(10.5,800), '#475467');
    y += 26;
    const mx = Math.max(1, ...erpM.map(r=>r[1]));
    erpM.forEach((r,i)=>{
      if(i%2) box(30,y,W-60,28,'#F8FAFD',null,0);
      txt(r[0], hx[0]+10, y+15, F(12,800), '#101828');
      txt(fmt(r[1]), hx[1]+cw2[1]-10, y+15, F(12), '#101828', 'right');
      txt(r[2], hx[2]+cw2[2]-10, y+15, F(12), '#475467', 'right');
      const bw = Math.max(2, Math.round(r[1]/mx*230));
      box(hx[3]+10, y+10, bw, 8, '#2452E6', null, 4);
      txt(Math.round(r[1]/(D.erpTotal||1)*100)+'%', hx[3]+bw+18, y+15, F(10.5), '#98A2B3');
      y += 28;
    });
    txt('※ ERP는 여러 작업을 한 장으로 묶어 발행합니다 — 장수는 계산서 매수입니다.',
        30, y+12, F(10), '#98A2B3');
    y += 20;
  }
  section('조치 필요 — 유형별 (프로젝트NO)');
  table(['유형','건수','금액(원)','프로젝트NO'],
        D.issues.length?D.issues:[['조치 필요 없음','-','-','-']], [200,70,130,440-100]);
  section('4원천 검증 (밴드·카톡·ERP·쿠팡PO)');
  sourceTable(sourceDetails);

  y += 20;
  txt('자동 생성 · Coupang Service Operations System Agent',30,y,F(10),'#98A2B3');
  txt(new Date().toLocaleString('ko-KR'),W-30,y,F(10),'#98A2B3','right');

  return new Promise(res=>cv.toBlob(res,'image/png'));
}
/* 대시보드 전체를 이미지로 — 보고서와 동일한 Canvas 직접 렌더(외부 리소스 0) */
async function dashToPng(){
  const S = 2, W = 900;
  const F = (sz,w) => `${w||400} ${sz}px "Nanum Gothic","Malgun Gothic",sans-serif`;
  const rows = settleRows, 유상 = rows.filter(r=>r.비용구분==='유상');
  const 발행 = 유상.filter(r=>r.계산서==='발행').length;
  const 미수 = rows.reduce((a,r)=>a+(+r.미수금||0),0);
  const issues = rows.filter(r=>needAction(r));
  const asOpen = (works.as||[]).filter(r=>r.진행상태 && r.진행상태!=='작업완료').length;
  const pmWait = (works.pm||[]).filter(r=>r.점검상태 && r.점검상태!=='완료').length;
  const y = APP_YEAR;
  const byM = Array.from({length:12},()=>0);
  rows.forEach(r=>{ const d=r.완료일||''; if(d.startsWith(y)){ const i=+d.slice(5,7)-1; if(i>=0&&i<12) byM[i]+=(+r.공급가액||0); } });
  const dist = {}; rows.forEach(r=>{ dist[r.상태]=(dist[r.상태]||0)+1; });
  const distE = Object.entries(dist).sort((a,b)=>b[1]-a[1]);
  const SCOLOR = {'정상':'#12813F','입금 대기':'#B54708','세금계산서 미발행':'#C0212E',
    '미청구(전표 없음)':'#8B1E68','금액 미입력':'#5B6B82','무상/보험':'#B9C3D3'};
  const byType = {}; issues.forEach(r=>{ (byType[r.상태]=byType[r.상태]||[]).push(r); });
  const typeE = Object.entries(byType).sort((a,b)=>b[1].length-a[1].length);
  const steps = (window._agentSteps||[]);

  const H = 104 + 34 + 148 + 34 + 190 + 34 + 76 + 34 + (typeE.length+1)*30 + 34 + (steps.length*22+10) + 50;
  const cv = document.createElement('canvas'); cv.width=W*S; cv.height=H*S;
  const ctx = cv.getContext('2d'); ctx.scale(S,S); ctx.textBaseline='middle';
  const txt=(s,x,yy,f,c,a)=>{ctx.font=f;ctx.fillStyle=c||'#101828';ctx.textAlign=a||'left';
    ctx.fillText(s==null?'':String(s),x,yy);ctx.textAlign='left';};
  const box=(x,yy,w,h,fill,stroke,r=10)=>{ctx.beginPath();
    if(ctx.roundRect)ctx.roundRect(x,yy,w,h,r);else ctx.rect(x,yy,w,h);
    if(fill){ctx.fillStyle=fill;ctx.fill();}if(stroke){ctx.strokeStyle=stroke;ctx.lineWidth=1;ctx.stroke();}};
  const clip=(s,max)=>{s=String(s==null?'':s);ctx.font=F(12);
    if(ctx.measureText(s).width<=max)return s;
    while(s.length>1&&ctx.measureText(s+'…').width>max)s=s.slice(0,-1);return s+'…';};

  ctx.fillStyle='#fff'; ctx.fillRect(0,0,W,H);
  const g=ctx.createLinearGradient(0,0,W,96); g.addColorStop(0,'#0E1B3F'); g.addColorStop(1,'#2452E6');
  ctx.fillStyle=g; ctx.fillRect(0,0,W,96);
  if(!brandLogo) box(30,26,44,44,'rgba(255,255,255,.16)',null,12);   // CI가 있으면 배경 사각형 없이
  await drawLogo(ctx,30,30,36,true);
  txt('Coupang Service Operations System',92,36,F(17,800),'#fff');
  txt('통합 현황 · UNIVERSAL LIFT',92,58,F(11.5),'#9FB4E8');
  txt(new Date().toLocaleDateString('ko-KR',{year:'numeric',month:'long',day:'numeric',weekday:'long'}),
      92,77,F(11),'#9FB4E8');
  txt('에이전트 '+($('heroAgent').textContent||'').replace('에이전트 ',''),W-30,40,F(11),'#9FB4E8','right');
  txt(rows.length+'건 · '+fmt(rows.reduce((a,r)=>a+(+r.공급가액||0),0))+'원',W-30,62,F(14,800),'#fff','right');

  let yy=96;
  const section=t=>{yy+=34;txt(t,30,yy-8,F(13,800),'#1B41BC');
    ctx.strokeStyle='#E4E9F0';ctx.beginPath();ctx.moveTo(30+ctx.measureText(t).width+12,yy-8);
    ctx.lineTo(W-30,yy-8);ctx.stroke();};

  section('핵심 지표');
  const kpi=[[rows.length+'건','정산 건수','#2452E6'],[fmt(rows.reduce((a,r)=>a+(+r.공급가액||0),0)),'공급가액(원)'],
    [fmt(미수),'미수금(원)',미수?'#B54708':'#12813F'],
    [(유상.length?Math.round(발행/유상.length*100):0)+'%','계산서 발행율'],
    [issues.length+'건','조치 필요',issues.length?'#C0212E':'#12813F'],
    [asOpen+' · '+pmWait,'AS진행 · 점검대기']];
  const cw=(W-60-16)/3;
  kpi.forEach((t,i)=>{const cx=30+(i%3)*(cw+8), cy=yy+Math.floor(i/3)*74;
    box(cx,cy,cw,64,'#F8FAFD','#EEF2F7');
    txt(t[1],cx+14,cy+22,F(11,700),'#98A2B3'); txt(t[0],cx+14,cy+44,F(19,800),t[2]||'#101828');});
  yy+=148;

  section(y+'년 월별 공급가액');
  const maxV=Math.max(1,...byM), bw=(W-60)/12-8;
  byM.forEach((v,i)=>{const bx=30+i*((W-60)/12), bh=Math.round(v/maxV*120)+2;
    box(bx,yy+130-bh,bw,bh,v?'#2452E6':'#E8EDF4',null,5);
    txt((i+1)+'월',bx+bw/2,yy+146,F(10,600),'#98A2B3','center');
    if(v) txt(Math.round(v/10000).toLocaleString()+'만',bx+bw/2,yy+130-bh-9,F(9,700),'#475467','center');});
  yy+=190;

  section('정산 상태 분포');
  let dx=30; const tot=distE.reduce((a,e)=>a+e[1],0)||1;
  distE.forEach(([k,v])=>{const w=(W-60)*v/tot;
    box(dx,yy,Math.max(w-2,4),14,SCOLOR[k]||'#5B6B82',null,4); dx+=w;});
  let lx=30, ly=yy+30;
  distE.forEach(([k,v])=>{const label=`${k} ${v}`; ctx.font=F(11,600);
    const wd=ctx.measureText(label).width+20;
    if(lx+wd>W-30){lx=30;ly+=20;}
    box(lx,ly-4,9,9,SCOLOR[k]||'#5B6B82',null,2);
    txt(label,lx+14,ly,F(11,600),'#475467'); lx+=wd+10;});
  yy+=76;

  section('조치 필요 — 유형별 (프로젝트NO)');
  const xs=[30,260,340,470]; const wds=[230,80,130,W-30-470];
  box(30,yy,W-60,28,'#F1F5FB','#E4E9F0',8);
  ['유형','건수','금액(원)','프로젝트NO'].forEach((c,i)=>txt(c,xs[i]+10,yy+14,F(11,800),'#475467'));
  yy+=28;
  (typeE.length?typeE:[['조치 필요 없음',[]]]).forEach(([t,l],ri)=>{
    if(ri%2){ctx.fillStyle='#FAFCFF';ctx.fillRect(30,yy,W-60,30);}
    const cells=[t,l.length?l.length+'건':'-',l.length?fmt(l.reduce((a,r)=>a+(+r.공급가액||0),0)):'-',
      l.slice(0,4).map(r=>projectNoOf(r)||'프로젝트 미확정').join(', ')+(l.length>4?` 외 ${l.length-4}`:'')];
    cells.forEach((c,i)=>txt(clip(c,wds[i]-20),xs[i]+10,yy+15,F(12,i===0?700:400),i===0?'#101828':'#475467'));
    ctx.strokeStyle='#EEF2F7';ctx.beginPath();ctx.moveTo(30,yy+30);ctx.lineTo(W-30,yy+30);ctx.stroke();
    yy+=30;});

  section('에이전트 최근 실행');
  steps.forEach(s=>{const ic=s.s==='ok'?'✔':(s.s==='skip'?'−':'✕');
    txt(ic,34,yy+4,F(12,800),s.s==='ok'?'#12813F':(s.s==='skip'?'#98A2B3':'#C0212E'));
    txt(s.n,52,yy+4,F(12),'#475467'); yy+=22;});

  yy+=18;
  txt('자동 생성 · Coupang Service Operations System Agent',30,yy,F(10),'#98A2B3');
  txt(new Date().toLocaleString('ko-KR'),W-30,yy,F(10),'#98A2B3','right');
  return new Promise(res=>cv.toBlob(res,'image/png'));
}
function dashFileName(){ return `CSOS_대시보드_${new Date().toISOString().slice(0,10)}.png`; }
async function captureDash(){
  let blob;
  try{ blob = await dashToPng(); if(!blob) throw new Error('빈 이미지'); }
  catch(e){ alert('이미지 생성 실패: '+e); return; }
  saveOrOpen(blob, dashFileName());
}
async function printDash(){
  try{printPngBlob(await dashToPng(),dashFileName());}
  catch(e){alert('인쇄 이미지 생성 실패: '+e);}
}
async function copyDashImage(){
  try{await copyPngBlob(await dashToPng());}
  catch(e){alert('이미지 복사 실패: '+e);}
}
function dashboardExcelRows(){
  const tagged=(rows,kind)=>(rows||[]).map(r=>({...r,업무구분:r.업무구분||kind}));
  return [
    ...tagged(settleRows,'정산·청구'),
    ...tagged((works&&works.as)||[],'돌발AS'),
    ...tagged((works&&works.pm)||[],'정기점검'),
    ...tagged((issuesData&&issuesData.rows)||[],'확인 필요')
  ];
}
function exportDashXlsx(){exportRowsXlsx('대시보드 전체 목록',dashboardExcelRows());}
/* PC=다운로드 / iOS·미지원=새 창 표시 (공통) */
function saveOrOpen(blob, filename){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  if(('download' in a) && !isIOS()){
    a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url), 60000);
    toast('이미지를 저장했습니다 · '+Math.round(blob.size/1024)+'KB');
  } else {
    const w = window.open();
    if(w){ w.document.write(`<html><head><meta name="viewport" content="width=device-width,initial-scale=1">
      <title>${filename}</title><style>body{margin:0;background:#111;font-family:sans-serif}
      p{color:#fff;text-align:center;padding:12px;font-size:14px;margin:0}img{width:100%;display:block}</style></head>
      <body><p>이미지를 길게 눌러 <b>사진에 저장</b>하세요</p><img src="${url}"></body></html>`);
      w.document.close();
    } else alert('팝업이 차단됐습니다. 팝업을 허용해 주세요.');
  }
}
function downloadBlobFile(blob, filename, message){
  const url=URL.createObjectURL(blob), a=document.createElement('a');
  a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),60000);
  toast(message||`${filename} 저장 완료`);
}
async function copyPngBlob(blob){
  try{
    if(!navigator.clipboard||typeof ClipboardItem==='undefined') throw new Error('clipboard unsupported');
    await navigator.clipboard.write([new ClipboardItem({'image/png':blob})]);
    toast('이미지를 클립보드에 복사했습니다 · 카톡에 바로 붙여넣을 수 있습니다');
    return true;
  }catch(e){
    alert('이 브라우저에서는 이미지 복사가 제한됩니다. 이미지 저장 기능을 이용해 주세요.');
    return false;
  }
}
function printPngBlob(blob,title){
  const url=URL.createObjectURL(blob), w=window.open('','_blank');
  if(!w){URL.revokeObjectURL(url);alert('팝업을 허용하면 인쇄할 수 있습니다.');return;}
  w.document.write(`<html><head><title>${esc2(title||'CSOS')}</title>
    <style>@page{margin:8mm}body{margin:0}img{display:block;max-width:100%;margin:auto}</style></head>
    <body><img src="${url}" onload="setTimeout(()=>window.print(),180)"></body></html>`);
  w.document.close();setTimeout(()=>URL.revokeObjectURL(url),120000);
}
function excelRows(rows){
  return (rows||[]).map(r=>({
    프로젝트NO:projectNoOf(r)||r.prj||r.프로젝트||'',
    업무ID:recordIdOf(r)||r.id||r.레코드ID||r.ID||'',
    캠프명:r.캠프명||r.camp||'',
    구분:r.업무구분||r.kind||r.종류||r.구분||'',
    확인사항:r.문제||r.issue||r.desc||r.확인사항||'',
    현재상태:r.상태||r.진행상태||r.점검상태||'',
    기준일자:r.일자||r.date||r.완료일||r.접수일자||r.점검예정일||'',
    담당자:r.담당자||r.담당기사||''
  }));
}
async function exportRowsXlsx(title,rows){
  try{
    const r=await fetch('/api/export_xlsx',{method:'POST',cache:'no-store',
      headers:{'X-Pin':PIN,'Content-Type':'application/json'},
      body:JSON.stringify({title,rows:excelRows(rows)})});
    if(r.status===401){PIN='';localStorage.removeItem('cw_pin');$('gate').style.display='flex';throw new Error('PIN');}
    if(!r.ok){let d={};try{d=await r.json();}catch(_){}
      throw new Error(d.error||`HTTP ${r.status}`);}
    const b=await r.blob(), safe=String(title||'확인목록').replace(/[\\/:*?"<>|\s]+/g,'_');
    downloadBlobFile(b,`CSOS_${safe}_${todayISO()}.xlsx`,'담당자 회신용 엑셀을 저장했습니다');
  }catch(e){alert('엑셀 저장 실패: '+e);}
}
function reportExcelRows(){
  const D=_rptData||{}, out=[];
  const add=(group,label,value,detail='')=>out.push({
    프로젝트NO:'',업무ID:'',캠프명:'',구분:group,확인사항:label,
    현재상태:value,기준일자:D.date||baseDate(),담당자:'',desc:detail
  });
  (D.month||[]).forEach(x=>add(D.monthTitle||'현재 월',x[1],x[0]));
  (D.quarter||[]).forEach(x=>add(D.quarterTitle||'이번 분기',x[1],x[0]));
  (D.kpi||[]).forEach(x=>add(D.kpiTitle||'연간 누적',x[1],x[0]));
  (D.dailyIssues||[]).forEach(x=>add('이슈사항',x[1],x[0],x[3]||''));
  (D.exec||[]).forEach(s=>{
    (s.items||[]).forEach(x=>add(cleanExecTitle(s.title),x[0],x[1]));
    (s.groups||[]).forEach(g=>(g.items||[]).forEach(x=>add(g.name,x[0],x[1])));
  });
  return out;
}
async function saveReportJournal(){
  try{renderDaily();const b=await reportToPng();downloadBlobFile(b,'CSOS_업무일지_'+
    ((_rptData&&_rptData.date)||baseDate())+'.png','업무일지를 저장했습니다');}
  catch(e){alert('일지 저장 실패: '+e);}
}
async function printReport(){try{renderDaily();printPngBlob(await reportToPng(),rptFileName());}
  catch(e){alert('인쇄 이미지 생성 실패: '+e);}}
async function copyReportImage(){try{renderDaily();await copyPngBlob(await reportToPng());}
  catch(e){alert('이미지 복사 실패: '+e);}}
function exportReportXlsx(){exportRowsXlsx(`일일보고_${(_rptData&&_rptData.date)||baseDate()}`,reportExcelRows());}
async function saveMetricJournal(){try{const b=await metricToPng(window._curMetricLabel);
  downloadBlobFile(b,'CSOS_업무일지_'+metricFileName());}catch(e){alert('일지 저장 실패: '+e);}}
async function printExecMetric(){try{printPngBlob(await metricToPng(window._curMetricLabel),metricFileName());}
  catch(e){alert('인쇄 이미지 생성 실패: '+e);}}
async function copyExecMetric(){try{await copyPngBlob(await metricToPng(window._curMetricLabel));}
  catch(e){alert('이미지 복사 실패: '+e);}}
function exportMetricXlsx(){const d=metricDetail(window._curMetricLabel);
  exportRowsXlsx(window._curMetricLabel||'확인목록',d.rows||[]);}
async function saveAssigneeJournal(){try{const b=await assigneeToPng(window._curAssignee);
  downloadBlobFile(b,'CSOS_업무일지_'+assigneeFileName());}catch(e){alert('일지 저장 실패: '+e);}}
async function printAssignee(){try{printPngBlob(await assigneeToPng(window._curAssignee),assigneeFileName());}
  catch(e){alert('인쇄 이미지 생성 실패: '+e);}}
async function copyAssigneeImage(){try{await copyPngBlob(await assigneeToPng(window._curAssignee));}
  catch(e){alert('이미지 복사 실패: '+e);}}
function exportAssigneeXlsx(){
  exportRowsXlsx(`확인요청_${window._curAssignee||'담당자'}`,window._curAssigneeList||[]);
}
function rptFileName(){
  const day = (_rptData&&_rptData.date) || REPORT_PREVIEW_DATE || baseDate() ||
              new Date().toISOString().slice(0,10);
  return `CSOS_일일보고_${day}.png`;
}
function isIOS(){ return /iPad|iPhone|iPod/.test(navigator.userAgent) ||
  (navigator.platform==='MacIntel' && navigator.maxTouchPoints>1); }

/* 이미지 저장 — PC는 즉시 다운로드, iOS/미지원 기기는 새 창에 띄워 길게 눌러 저장 */
async function captureReport(){
  let blob;
  try{
    // 클릭 시점의 최신 4원천 집계를 다시 묶어 화면과 저장 이미지가 항상 같게 한다.
    renderDaily();
    blob = await reportToPng();
    if(!blob) throw new Error('이미지 생성 결과가 비어 있습니다');
  }catch(e){ alert('이미지 생성 실패: '+e+'\n화면 캡처(Win+Shift+S / 전원+음량↓)를 사용해 주세요.'); return; }

  saveOrOpen(blob, rptFileName());
}
async function shareReport(){
  let blob;
  try{
    renderDaily();
    blob = await reportToPng();
    if(!blob) throw new Error('empty');
  }catch(e){ alert('이미지 생성 실패: '+e); return; }
  const file = new File([blob], rptFileName(), {type:'image/png'});
  try{
    if(navigator.canShare && navigator.canShare({files:[file]})){
      await navigator.share({files:[file], title:'Coupang Service Operations System — 일일보고',
                             text:'Coupang Service Operations System — 일일보고'});
      return;
    }
  }catch(e){ if(e.name==='AbortError') return; }
  captureReport();     // 공유 미지원 → 저장/새창 경로로 폴백
}
function toast(msg){
  let t = document.getElementById('_toast');
  if(!t){ t = document.createElement('div'); t.id='_toast';
    t.style.cssText='position:fixed;left:50%;bottom:80px;transform:translateX(-50%);z-index:60;'+
      'background:#101828;color:#fff;padding:11px 18px;border-radius:11px;font-size:13.5px;'+
      'font-weight:700;box-shadow:0 8px 24px rgba(0,0,0,.3);opacity:0;transition:opacity .2s';
    document.body.appendChild(t); }
  t.textContent = msg; t.style.opacity='1';
  clearTimeout(t._h); t._h = setTimeout(()=>t.style.opacity='0', 2600);
}

/* ── 부트 ── */
function tick(){ $('clock').textContent = new Date().toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'}); }
async function refreshAll(){
  try{ await loadSettle(); await loadStatus(); await loadNotifications(); if(staffSlug) await staffHeartbeat('view'); }catch(e){}
}
(async function(){
  tick(); setInterval(tick, 30000); initDates(); previewCalendarDraft();
  initStaffCenter();
  const surveyDate=document.querySelector('#ryuUploadForm [name="survey_date"]');
  if(surveyDate&&!surveyDate.value) surveyDate.value=todayISO();
  const newWorkDate=document.querySelector('#newWorkForm [name="work_date"]');
  if(newWorkDate&&!newWorkDate.value) newWorkDate.value=todayISO();
  showFixedEntry();          // 고정 주소는 로그인·데이터 로딩과 무관하게 **항상** 보인다
  loadCodes();               // 드롭다운 선택지(10_코드관리)
  applyView(curView());                  // 새로고침 전에 보던 화면으로 복귀(뒤로가기 기록은 안 쌓는다)
  // ★ 여분 항목 하나를 미리 쌓아 둔다. 이게 없으면 **첫 뒤로가기가 앱 밖으로 나가** 버려서
  //   홈에서 종료를 묻는 것도, 이전 화면으로 돌아가는 것도 할 수 없다.
  navGuard();
  if(PIN && await restoreRoleSession()){
    $('gate').style.display='none';
    refreshAll();
    pollLog().catch(()=>{});
  }
  if(staffSlug){
    setInterval(()=>staffHeartbeat('view'),30000);
    document.addEventListener('input',()=>markStaffInput('input'),true);
    document.addEventListener('change',()=>markStaffInput('change'),true);
    document.addEventListener('paste',()=>markStaffInput('paste'),true);
    document.addEventListener('drop',()=>markStaffInput('drop'),true);
  }
  setInterval(()=>{if(PIN) loadNotifications()},60000);
  setInterval(()=>{ if(PIN && !runnerBusy()) loadStatus().catch(()=>{}); }, 30000);
  checkBuild(false);                                  // 최초 기준값 기록
  setInterval(()=>checkBuild(true), 20000);           // 새 버전 나오면 자동 갱신
  document.addEventListener('visibilitychange', ()=>{ if(!document.hidden) checkBuild(true); });
})();
function runnerBusy(){ return !!polling; }
