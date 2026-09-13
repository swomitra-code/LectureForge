// Headless Edge acceptance for our local Studio only; no provider transport.
const {spawn}=require('node:child_process');
const fs=require('node:fs'),path=require('node:path');
const root=path.resolve(process.argv[2]||path.join(process.env.LOCALAPPDATA,'LectureForgeD2'));
const id=process.argv[3]||'709e1df2cb7a',url='http://127.0.0.1:'+(process.argv[4]||8778)+'/',debug=9238;
const checks=[],profile=path.join(root,'browser-'+Date.now());fs.mkdirSync(profile);
const edge=spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',['--headless=new','--no-first-run','--disable-background-networking','--disable-sync','--remote-debugging-port='+debug,'--user-data-dir='+profile,'about:blank'],{windowsHide:true,stdio:'ignore'});
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
let ws,seq=0;const pending=new Map();
function command(method,params={}){return new Promise((resolve,reject)=>{const n=++seq;const timer=setTimeout(()=>{pending.delete(n);reject(Error('CDP timeout: '+method))},30000);pending.set(n,{resolve:v=>{clearTimeout(timer);resolve(v)},reject});ws.send(JSON.stringify({id:n,method,params}));});}
async function evaluate(expression){const r=await command('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.text);return r.result.value;}
async function until(expression){for(let i=0;i<60;i++){if(await evaluate(expression))return;await sleep(500)}throw Error('UI condition timeout: '+expression);}
function check(name,ok){if(!ok)throw Error('FAIL: '+name);checks.push({name,passed:true});console.log('PASS: '+name);}
(async()=>{
 try{
  let target;for(let i=0;i<60;i++){try{const r=await fetch('http://127.0.0.1:'+debug+'/json');target=(await r.json()).find(t=>t.type==='page');if(target)break}catch{}await sleep(250)}
  if(!target)throw Error('Isolated Edge did not start');
  ws=new WebSocket(target.webSocketDebuggerUrl);await new Promise((r,j)=>{ws.onopen=r;ws.onerror=j});
  ws.onmessage=e=>{const r=JSON.parse(e.data);if(r.id&&pending.has(r.id)){const p=pending.get(r.id);pending.delete(r.id);r.error?p.reject(Error(r.error.message)):p.resolve(r.result)}};
  await command('Page.enable');await command('Page.navigate',{url});await until("document.getElementById('open')!==null");
  await evaluate(`localStorage.setItem('lectureforge-project','${id}')`);await command('Page.reload');
  await until("document.getElementById('avatar-summary').textContent.includes('13 Avatar Ready')");
  check('Browser loads durable 13-ready dashboard',await evaluate("!document.getElementById('avatar-dashboard').hidden"));
  check('Browser displays Slide 8 verified six-second pause',await evaluate("document.getElementById('avatar-jobs').textContent.includes('6.000 s pause verified')"));
  await evaluate("document.getElementById('pause-avatars').click()");await until("document.getElementById('pause-avatars').textContent==='Resume Queue'");
  check('Pause Queue works through actual browser button',true);
  await command('Page.reload');await until("document.getElementById('avatar-summary').textContent.includes('13 Avatar Ready')&&document.getElementById('pause-avatars').textContent==='Resume Queue'");
  check('Browser refresh restores project, paused queue and completed avatars',true);
  await evaluate("document.getElementById('pause-avatars').click()");await until("document.getElementById('pause-avatars').textContent==='Pause Queue'");
  await evaluate("document.querySelector('#avatar-jobs summary').click()");
  check('View Details opens without mutating provider state',await evaluate("document.querySelector('#avatar-jobs details').open"));
  const state=await evaluate(`fetch('/api/avatar-status?id=${id}').then(r=>r.json())`);
  check('Browser review, refresh and queue controls make zero provider calls',state.ready===13&&state.provider_calls===0&&!state.paused);
  await command('Emulation.setDeviceMetricsOverride',{width:1280,height:900,deviceScaleFactor:1,mobile:false});
  await evaluate("document.getElementById('avatar-dashboard').scrollIntoView()");
  const shot=await command('Page.captureScreenshot',{format:'png'});fs.writeFileSync(path.join(root,'dashboard.png'),Buffer.from(shot.data,'base64'));
  fs.writeFileSync(path.join(root,'browser-results.json'),JSON.stringify({checks,elevenlabs_calls:0,heygen_calls:0},null,2));
 }finally{if(ws)ws.close();edge.kill();}
})().catch(e=>{console.error(e.message);process.exitCode=1});
