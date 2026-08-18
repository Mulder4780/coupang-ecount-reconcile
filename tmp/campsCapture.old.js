async function campsCapture(){
  const rows = campRows();
  if(!rows.length){ notice('캡처할 캠프가 없습니다.'); return; }
  const S=2, W=1740, F=(z,w)=>`${w||400} ${z}px ${uiFont()}`;
  /* 한도는 캔버스 안전 면적에서 나온다 — iOS 는 약 1,678만 픽셀에서 저장이 실패한다.
     폭 W*S 를 빼면 세로가 정해지고, 머리글·각주 자리를 뺀 나머지가 실을 수 있는 줄이다. */
  const RH=24, HEAD=112, FOOT=92;
  const LIM = Math.max(20, Math.floor((16777216/(W*S)/S - HEAD - FOOT)/RH));
  const show = rows.slice(0, LIM);
  const H = HEAD + show.length*RH + FOOT;
  const cv=document.createElement('canvas'); cv.width=W*S; cv.height=H*S;
  const x=cv.getContext('2d'); x.scale(S,S);
  let clipped = 0;
  /* 칸에 안 들어가면 **그 칸만** 글자를 줄인다. 그래도 안 되면 그때만 자르고 센다. */
  const fit=(t,px,py,w,weight)=>{
    t = String(t==null?'':t);
    for(const z of [12,11,10,9,8]){
      x.font=F(z,weight);
      if(x.measureText(t).width <= w){ x.fillText(t,px,py); x.font=F(12,weight); return; }
    }
    x.font=F(8,weight);
    let s=t; while(s.length>1 && x.measureText(s+'…').width>w) s=s.slice(0,-1);
    if(s!==t) clipped++;
    x.fillText(s===t?s:s+'…',px,py); x.font=F(12,weight);
  };
  x.fillStyle='#fff'; x.fillRect(0,0,W,H);
  x.fillStyle='#0E1B3F'; x.fillRect(0,0,W,66);
  x.fillStyle='#fff'; x.font=F(20,700);
  x.fillText($('campPmOnly').checked ? '전국 정기점검 캠프 · 담당자' : '전국 쿠팡캠프 · 담당자',24,34);
  x.font=F(12); x.fillStyle='#C9D5FA';
  x.fillText(`총 ${rows.length}개 · 이 이미지에 ${show.length}개 · 기준 ${new Date().toLocaleString('ko-KR')}`,24,54);
  /* 열 자리 — [시작x, 폭]. 사람마다 이름·전화·이메일 셋이고 역할이 셋이라 아홉 칸이다. */
  const NAME=110, TEL=112, MAIL=206, GAP=4;
  const cols=[[24,250]]; let cx=282;
  for(let i=0;i<3;i++){ cols.push([cx,NAME]); cols.push([cx+NAME+GAP,TEL]);
    cols.push([cx+NAME+GAP+TEL+GAP,MAIL]); cx += NAME+TEL+MAIL+GAP*3+16; }
  cols.push([cx,100]);
  /* 머리글 두 줄: 위는 역할, 아래는 그 역할의 세 칸. 받는 사람이 ①②를 해석할 일이 없다. */
  let y=84; x.fillStyle='#0E1B3F'; x.textAlign='left';
  CAMP_ROLES.forEach((rl,i)=>{ const a=cols[1+i*3][0], b=cols[3+i*3][0]+cols[3+i*3][1];
    x.font=F(12,700); x.fillText(rl[1], a, y);
    x.strokeStyle='#C7D2FE'; x.beginPath(); x.moveTo(a,y+5); x.lineTo(b,y+5); x.stroke(); });
  y+=22; x.fillStyle='#374151'; x.font=F(11,700);
  ['캠프명','이름','전화','이메일','이름','전화','이메일','이름','전화','이메일','최근작업']
    .forEach((h,i)=>x.fillText(h,cols[i][0],y));
  y+=6; x.strokeStyle='#9CA3AF'; x.beginPath(); x.moveTo(24,y); x.lineTo(W-24,y); x.stroke();
  y+=18;
  show.forEach((r,ri)=>{
    if(ri%2){ x.fillStyle='#F8FAFC'; x.fillRect(24,y-14,W-48,RH); }
    const v=[r.캠프명];
    campRoles(r).forEach(rl=>v.push(rl.p.이름||'모름', rl.p.전화||'모름', rl.p.메일||'모름'));
    v.push(r.최근작업일||'');
    v.forEach((t,i)=>{ x.fillStyle=(t==='모름')?'#9CA3AF':'#111827';
      fit(t, cols[i][0], y, cols[i][1], i===0?600:400); });
    y+=RH;
  });
  y+=20; x.fillStyle='#6B7280'; x.font=F(11);
  x.fillText("'모름'은 담당자가 없는 것이 아니라 밴드 접수 글에서 아직 찾지 못한 것입니다.",24,y); y+=15;
  x.fillText("현장책임 이메일이 대부분 비어 있는 것은 자료가 빠진 것이 아니라 접수 양식이 그 칸을 적지 않기 때문입니다 — 이메일은 안전관리 쪽에 적혀 옵니다.",24,y); y+=15;
  x.fillText("담당자(직책 미상)는 옛 접수 양식(2023~2024) 글의 한 사람으로, 원문에 현장책임인지 안전관리인지 적혀 있지 않습니다.",24,y);
  if(rows.length>show.length){ y+=16; x.fillStyle='#B45309';
    x.fillText(`※ 이미지에는 ${show.length}개만 실었습니다 — 나머지 ${rows.length-show.length}개는 엑셀 저장으로 받으십시오.`,24,y); }
  if(clipped){ y+=16; x.fillStyle='#B45309';
    x.fillText(`※ 칸에 넣지 못해 줄인 값 ${clipped}개가 있습니다 — 전체 값은 엑셀 저장으로 확인하십시오.`,24,y); }
  const blob=await new Promise(r=>cv.toBlob(r,'image/png'));
  if(!blob){ notice('이미지 생성 실패'); return; }
  saveOrOpen(blob,`전국쿠팡캠프_담당자_${todayISO()}.png`);
  toast('캠프 담당자 목록 이미지를 저장했습니다'); uxEvent('tap','campsCapture');
}