# -*- coding: utf-8 -*-
import io, os, sys
P = r"C:/Users/hueng/Documents/COUPANG_INTEGRATED_WORK_AGENT/ecount/webapp/index.html"
SP = sys.argv[1]
src = io.open(P, encoding="utf-8", newline="").read()

# 재려는 코드는 실제 소스에서 뽑는다 — 스텁으로 때우지 않는다([366])
i = src.index("\nlet LOGIN_TRY = null;")
j = src.index("\n$('pin').addEventListener('keyup'", i)
code = src[i+1:j]
k = src.index("\nconst RETRY_WAIT_MS = [")
ladder = src[k+1:src.index("\n", k+1)]

HARNESS = """
// --- 환경 스텁(재려는 것이 아니다) ---
let PIN='', LEGACY_PIN='', staffSlug='ryu-jiyeong';
const LIVE_SYNC={booting:false};
const _els={pin:{value:'1234'},pinerr:{innerHTML:'',textContent:''},gate:{style:{}}};
const $ = id => _els[id];
const localStorage={ _d:{}, getItem(k){return this._d[k]||null;}, setItem(k,v){this._d[k]=v;}, removeItem(k){delete this._d[k];} };
let booted=0;
function show(){} function curView(){return 'dash';} function bootstrapAuthenticatedData(){booted++;}
// _wait 은 목이다 — 진짜로 기다리면 검사가 2분 걸린다
let WAITS=[];
let CANCEL_AT=0;
const _wait = ms => { WAITS.push(ms);
  /* 목이 즉시 끝나므로 setTimeout 취소는 루프가 다 돈 뒤에야 온다.
     재려는 것은 '기다리는 중에 취소가 걸리면 멈추는가' 이므로
     그 순간을 목이 직접 만든다([272] — 재료가 사고를 재현해야 한다). */
  if(CANCEL_AT && WAITS.length===CANCEL_AT) loginCancel();
  return Promise.resolve(); };
__LADDER__
__CODE__

// --- 재기 ---
let FETCHES=0, MODE='throw';
globalThis.fetch = async () => {
  FETCHES++;
  if(MODE==='throw') throw new Error('net');
  if(MODE==='throw-then-ok' && FETCHES<3) throw new Error('net');
  if(MODE==='429') return {ok:false, status:429};
  if(MODE==='401') return {ok:false, status:401};
  return {ok:true, status:200};
};
const out=[];
function reset(m){ MODE=m; FETCHES=0; WAITS=[]; booted=0; LOGIN_TRY=null;
                   _els.pinerr.innerHTML=''; _els.pinerr.textContent=''; }

(async () => {
  // (1) 계속 끊기면 사다리 끝까지 다시 건다
  reset('throw'); await login();
  out.push(['1 재시도 횟수', FETCHES, RETRY_WAIT_MS.length+1]);
  out.push(['1 기다린 초', JSON.stringify(WAITS), JSON.stringify(RETRY_WAIT_MS)]);
  out.push(['1 마지막 문구가 오프라인 안내', /오프라인 입력 모드/.test(_els.pinerr.innerHTML), true]);

  // (2) 중간에 돌아오면 그때 들어간다
  reset('throw-then-ok'); await login();
  out.push(['2 세 번째에 성공', FETCHES, 3]);
  out.push(['2 부팅됨', booted, 1]);

  // (3) 429(잠금)는 다시 걸지 않는다 — 걸수록 나빠진다
  reset('429'); await login();
  out.push(['3 429 는 한 번만', FETCHES, 1]);
  out.push(['3 429 문구', _els.pinerr.textContent, '시도 초과로 잠금 — 10분 후 다시 시도하세요']);

  // (4) PIN 오류도 한 번만
  reset('401'); await login();
  out.push(['4 401 은 한 번만', FETCHES, 1]);
  out.push(['4 401 문구', _els.pinerr.textContent, 'PIN이 올바르지 않습니다']);

  // (5) 성공하면 한 번 그대로 — 정상 경로가 안 느려진다
  reset('ok'); await login();
  out.push(['5 성공은 한 번', FETCHES, 1]);
  out.push(['5 기다림 없음', WAITS.length, 0]);

  // (6) 취소하면 멈춘다
  reset('throw'); CANCEL_AT=1; await login(); CANCEL_AT=0;
  out.push(['6 취소하면 끝까지 안 감', FETCHES < RETRY_WAIT_MS.length+1, true]);

  // (7) 두 번 눌러도 한 번만 돈다
  reset('throw');
  const a = login(), b = login();
  await Promise.all([a,b]);
  out.push(['7 두 번 눌러도 한 벌', FETCHES, RETRY_WAIT_MS.length+1]);

  let bad=0;
  for(const [name,got,want] of out){
    const ok = JSON.stringify(got)===JSON.stringify(want);
    if(!ok) bad++;
    console.log((ok?'OK  ':'FAIL')+'  '+name+'  got='+JSON.stringify(got)+(ok?'':'  want='+JSON.stringify(want)));
  }
  console.log(bad? ('실패 '+bad+'건') : '모두 통과');
})();
"""
js = HARNESS.replace("__LADDER__", ladder).replace("__CODE__", code)
io.open(os.path.join(SP, "t390.mjs"), "w", encoding="utf-8", newline="").write(js)
print("하네스 기록 · 코드 %d자" % len(code))
