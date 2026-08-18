const $=()=>({checked:true}),uiFont=()=>"sans",campRows=()=>[],campRoles=()=>[],CAMP_ROLES=[[0,"a"],[0,"b"],[0,"c"]],notice=()=>0,toast=()=>0,uxEvent=()=>0,saveOrOpen=()=>0,todayISO=()=>"d";
async function campsCapture(){
  const rows = campRows();
  if(!rows.length){ notice('캡처할 캠프가 없습니다.'); return; }
  /* 여백은 네 변이 같다(2026-08-18 지시) — 좌·우·위·아래 전부 M.
     열 폭도 여기서 정해 **마지막 열이 정확히 W-M 에 닿는다.** 안 맞추면 표가 캔버스
     밖으로 넘어가 오른쪽 끝 글자가 말없이 사라진다 — 실측으로 010-6445-1001 이
     010-6445-100 으로 나갔는데 '줄인 값' 각주도 안 붙었다(넘친 것은 자른 것이
     아니라서 세지도 못한다). 번호는 한 글자만 틀려도 전화가 안 걸리는데
     잘린 티가 안 난다. 열을 늘리면 반드시 이 계산을 다시 본다. */
  const S=2, W=1740, M=28, RH=24, BAR=66;
  const F=(z,w)=>`${w||400} ${z}px ${uiFont()}`;
  const mc = document.createElement('canvas').getContext('2d');

  /* 열 자리 — [시작x, 폭]. 사람마다 이름·전화·이메일 셋이고 역할이 셋이라 아홉 칸이다.
     이메일만 남는 폭을 나눠 갖고, 마지막 열은 남은 것을 그대로 받는다. */
  const GAP=6, BLOCK=20, CAMP=252, NAME=116, TEL=124, LAST=104;
  const MAIL = Math.floor((W - M*2 - (CAMP + BLOCK*4 + 3*(NAME+GAP+TEL+GAP) + LAST))/3);
  const cols=[]; let cx=M;
  cols.push([cx,CAMP]); cx += CAMP + BLOCK;
  for(let i=0;i<3;i++){
    cols.push([cx,NAME]); cx += NAME+GAP;
    cols.push([cx,TEL ]); cx += TEL +GAP;
    cols.push([cx,MAIL]); cx += MAIL+BLOCK;
  }
  cols.push([cx, W-M-cx]);

  /* 칸에 안 들어가면 **그 칸만** 글자를 줄인다. 그래도 안 되면 그때만 자르고 센다.
     그리지 않고 '어떻게 들어가는지'만 답하므로 미리 셀 수 있다 — 각주가 몇 줄인지
     정해져야 아래 여백을 위와 같게 맞출 수 있다. */
  const shrink=(g,t,w,weight)=>{
    t = String(t==null?'':t);
    for(const z of [12,11,10,9,8]){ g.font=F(z,weight);
      if(g.measureText(t).width<=w) return {s:t,z,cut:false}; }
    g.font=F(8,weight);
    let s=t; while(s.length>1 && g.measureText(s+'…').width>w) s=s.slice(0,-1);
    return {s:(s===t?s:s+'…'), z:8, cut:s!==t};
  };
  const valuesOf=r=>{ const v=[r.캠프명];
    campRoles(r).forEach(rl=>v.push(rl.p.이름||'모름', rl.p.전화||'모름', rl.p.메일||'모름'));
    v.push(r.최근작업일||''); return v; };

  const roleY = BAR + M + 12, colY = roleY + 22, ruleY = colY + 6, row0 = ruleY + 18;
  const NGAP=20, NLH=15, DESC=4;
  const notes=[
    "'모름'은 담당자가 없는 것이 아니라 밴드 접수 글에서 아직 찾지 못한 것입니다.",
    "현장책임 이메일이 대부분 비어 있는 것은 자료가 빠진 것이 아니라 접수 양식이 그 칸을 적지 않기 때문입니다 — 이메일은 안전관리 쪽에 적혀 옵니다.",
    "담당자(직책 미상)는 옛 접수 양식(2023~2024) 글의 한 사람으로, 원문에 현장책임인지 안전관리인지 적혀 있지 않습니다."
  ];
  /* 한도는 캔버스 안전 면적에서 나온다 — iOS 는 약 1,678만 픽셀에서 저장이 실패한다.
     각주가 최대 두 줄 더 붙을 수 있으므로 그만큼을 미리 빼 둔다. */
  const foot = k => NGAP + (k-1)*NLH + DESC + M;
  const LIM = Math.max(20, Math.floor((16777216/(W*S*S) - row0 - foot(notes.length+2))/RH));
  const show = rows.slice(0, LIM);

  let clipped = 0;
  show.forEach(r=>valuesOf(r).forEach((t,i)=>{ if(shrink(mc,t,cols[i][1],i===0?600:400).cut) clipped++; }));
  if(rows.length>show.length) notes.push(`※ 이미지에는 ${show.length}개만 실었습니다 — 나머지 ${rows.length-show.length}개는 엑셀 저장으로 받으십시오.`);
  if(clipped) notes.push(`※ 칸에 넣지 못해 줄인 값 ${clipped}개가 있습니다 — 전체 값은 엑셀 저장으로 확인하십시오.`);

  const H = row0 + show.length*RH + foot(notes.length);
  const cv=document.createElement('canvas'); cv.width=W*S; cv.height=H*S;
  const x=cv.getContext('2d'); x.scale(S,S); x.textAlign='left';
  const put=(t,px,py,w,weight)=>{ const r=shrink(x,t,w,weight); x.font=F(r.z,weight); x.fillText(r.s,px,py); };

  x.fillStyle='#fff'; x.fillRect(0,0,W,H);
  x.fillStyle='#0E1B3F'; x.fillRect(0,0,W,BAR);
  x.fillStyle='#fff'; x.font=F(20,700);
  x.fillText($('campPmOnly').checked ? '전국 정기점검 캠프 · 담당자' : '전국 쿠팡캠프 · 담당자',M,34);
  x.font=F(12); x.fillStyle='#C9D5FA';
  x.fillText(`총 ${rows.length}개 · 이 이미지에 ${show.length}개 · 기준 ${new Date().toLocaleString('ko-KR')}`,M,54);

  /* 머리글 두 줄: 위는 역할, 아래는 그 역할의 세 칸. 받는 사람이 ①②를 해석할 일이 없다. */
  x.fillStyle='#0E1B3F';
  CAMP_ROLES.forEach((rl,i)=>{ const a=cols[1+i*3][0], b=cols[3+i*3][0]+cols[3+i*3][1];
    x.font=F(12,700); x.fillText(rl[1], a, roleY);
    x.strokeStyle='#C7D2FE'; x.beginPath(); x.moveTo(a,roleY+5); x.lineTo(b,roleY+5); x.stroke(); });
  x.fillStyle='#374151'; x.font=F(11,700);
  ['캠프명','이름','전화','이메일','이름','전화','이메일','이름','전화','이메일','최근작업']
    .forEach((h,i)=>x.fillText(h,cols[i][0],colY));
  x.strokeStyle='#9CA3AF'; x.beginPath(); x.moveTo(M,ruleY); x.lineTo(W-M,ruleY); x.stroke();

  let y=row0;
  show.forEach((r,ri)=>{
    if(ri%2){ x.fillStyle='#F8FAFC'; x.fillRect(M,y-14,W-M*2,RH); }
    valuesOf(r).forEach((t,i)=>{ x.fillStyle=(t==='모름')?'#9CA3AF':'#111827';
      put(t, cols[i][0], y, cols[i][1], i===0?600:400); });
    y+=RH;
  });
  y+=NGAP; x.font=F(11);
  notes.forEach((n,i)=>{ x.fillStyle = n.startsWith('※') ? '#B45309' : '#6B7280';
    x.fillText(n, M, y+i*NLH); });

  const blob=await new Promise(r=>cv.toBlob(r,'image/png'));
  if(!blob){ notice('이미지 생성 실패'); return; }
  saveOrOpen(blob,`전국쿠팡캠프_담당자_${todayISO()}.png`);
  toast('캠프 담당자 목록 이미지를 저장했습니다'); uxEvent('tap','campsCapture');
}
