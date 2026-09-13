// Offline F browser acceptance. Fixture import is test setup, never a Studio feature.
const {spawn}=require('node:child_process');
const fs=require('node:fs'), path=require('node:path');
const repo=path.resolve(__dirname,'../..');
const root=path.resolve(process.argv[2]||path.join(process.env.LOCALAPPDATA,'LectureForgeF'));
const port=Number(process.argv[3]||8797), url=`http://127.0.0.1:${port}/`, debug=9241;
const checks=[], pending=new Map();let ws,edge,seq=0,id;
const resume=process.argv.includes('--resume');
if(resume){const old=readPrevious();id=old.id;checks.push(...old.checks);}
function readPrevious(){return JSON.parse(fs.readFileSync(path.join(root,'browser-results.json'),'utf8'));}
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const read=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
async function runtime(){for(let i=0;i<20;i++){try{return read(path.join(root,'runtime.json'))}catch(e){if(i===19)throw e;await sleep(50)}}}
async function run(exe,args){return new Promise((resolve,reject)=>{
 const child=spawn(exe,args,{cwd:repo,windowsVerbatimArguments:exe==='cmd.exe',stdio:['ignore','pipe','pipe']});let output='';
 child.stdout.on('data',d=>{output+=d;process.stdout.write(d)});
 child.stderr.on('data',d=>{output+=d;process.stderr.write(d)});
 child.on('error',reject);child.on('exit',code=>{child.stdout.destroy();child.stderr.destroy();code===0?resolve(output):reject(Error(`${exe} exited ${code}`))});
});}
function ps(mode){return run('powershell.exe',['-NoProfile','-ExecutionPolicy','Bypass','-File','tests/v2/Prepare-DressRehearsal.ps1','-Mode',mode,'-RuntimeRoot',root,...(id?['-Id',id]:[])]);}
function launch(){return run('cmd.exe',['/d','/c',`""${path.join(repo,'LectureForge.cmd')}" -RuntimeRoot "${root}" -Port ${port}"`]);}
function command(method,params={}){return new Promise((resolve,reject)=>{const n=++seq,t=setTimeout(()=>{pending.delete(n);reject(Error('CDP timeout '+method))},120000);pending.set(n,{timer:t,resolve:r=>{clearTimeout(t);resolve(r)},reject:e=>{clearTimeout(t);reject(e)}});ws.send(JSON.stringify({id:n,method,params}));});}
async function evaluate(expression){const r=await command('Runtime.evaluate',{expression:`eval(${JSON.stringify(expression)})`,userGesture:true,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw Error(JSON.stringify(r.exceptionDetails));return r.result.value;}
async function until(expression,minutes=3){const end=Date.now()+minutes*60000;while(Date.now()<end){try{if(await evaluate(expression))return}catch(e){if(!/ReferenceError|Cannot read properties of null|Cannot find context|context was destroyed/.test(e.message))throw e}await sleep(1500)}throw Error('UI timeout '+expression);}
async function click(id){
 // Dispatch the real button event, then await its existing asynchronous handler.
 // This prevents the test from clicking the next screen before navigation finishes.
 await evaluate(`(async()=>{const b=document.getElementById(${JSON.stringify(id)});if(b.disabled||!b.checkVisibility())throw Error('Control is not actionable: '+b.id);const handler=b.onclick;if(!handler){b.click();return;}let done;b.onclick=function(e){done=Promise.resolve(handler.call(this,e));return done;};try{b.click();await done;}finally{b.onclick=handler;}})()`);
}
function check(name,ok){if(!ok)throw Error('FAIL: '+name);if(!checks.some(c=>c.name===name))checks.push({name,passed:true});console.log('PASS: '+name);fs.writeFileSync(path.join(root,'browser-results.json'),JSON.stringify({checks,id,url,elevenlabs_calls:0,heygen_calls:0},null,2));}
async function stopFixture(){await click('stop');for(let i=0;i<120&&(await runtime()).status!=='stopped';i++)await sleep(500);if((await runtime()).status!=='stopped')throw Error('Fixture did not stop cleanly');}
async function upload(selector,file){const d=await command('DOM.getDocument');const n=await command('DOM.querySelector',{nodeId:d.root.nodeId,selector});await command('DOM.setFileInputFiles',{nodeId:n.nodeId,files:[file]});}
(async()=>{try{
 if(!fs.existsSync(path.join(root,'input.pptx')))await ps('prepare');
 else if(fs.existsSync(path.join(root,'Projects'))&&!resume)throw Error('Use --resume to preserve existing fixture evidence.');
 const started=await launch();check('Simple CMD launcher detects prerequisites and opens Studio',started.includes('LectureForge ready')&&started.includes('ElevenLabs API key found')&&started.includes('HeyGen API key found'));
 for(const key of [process.env.ELEVENLABS_API_KEY,process.env.HEYGEN_API_KEY])if(key&&started.includes(key))throw Error('Secret exposed in startup');
 check('Startup reports key presence without values',true);
 const instance=read(path.join(root,'runtime.json')).instance;await launch();check('Double launch reuses the existing instance',read(path.join(root,'runtime.json')).instance===instance);
 edge=spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',['--no-first-run','--disable-background-networking','--disable-sync','--disable-backgrounding-occluded-windows','--disable-features=CalculateNativeWinOcclusion','--remote-debugging-port='+debug,'--user-data-dir='+path.join(root,'browser'),url],{stdio:'ignore'});
 let target;for(let i=0;i<120;i++){try{target=(await(await fetch(`http://127.0.0.1:${debug}/json`)).json()).find(t=>t.type==='page');if(target)break}catch{}await sleep(500)}
 if(!target)throw Error('Visible test browser unavailable');ws=new WebSocket(target.webSocketDebuggerUrl);await new Promise((r,j)=>{ws.onopen=r;ws.onerror=j});ws.onmessage=e=>{const r=JSON.parse(e.data),p=pending.get(r.id);if(p){pending.delete(r.id);r.error?p.reject(Error(r.error.message)):p.resolve(r.result)}else if(r.method?.startsWith('Media.'))fs.appendFileSync(path.join(root,'media-diagnostics.jsonl'),JSON.stringify(r)+'\n')};
 await command('Media.enable');
 await command('Page.enable');await until("document.getElementById('new')!==null");
 if(!resume){await click('new');
 await evaluate("document.getElementById('name').value='Solar Thermal V2 Dress Rehearsal'");await upload('#pptx',path.join(root,'input.pptx'));
 await until("document.getElementById('preset').options.length>0");check('Renewable Energy preset available in New Lecture',await evaluate("document.getElementById('preset').selectedOptions[0].textContent.includes('Renewable')"));
 await evaluate("document.querySelector('#new-form button').click()");await until("!document.getElementById('project').hidden");id=await evaluate("localStorage.getItem('lectureforge-project')");
 check('Fresh project created through source upload UI',!!id&&await evaluate("document.getElementById('source').textContent.includes('14 slides')"));
 await evaluate("document.getElementById('script-file').closest('details').open=true");await upload('#script-file',path.join(root,'scripts.md'));await until("document.getElementById('script-import').value.includes('## Slide 14')");await click('import');await until("document.getElementById('message').textContent==='Scripts imported and saved.'");
 await evaluate("const p=document.querySelector('[data-number=\"8\"] .pause');p.value='6';p.dispatchEvent(new Event('input',{bubbles:true}))");await click('save');await until("document.getElementById('saved').textContent==='Saved'");
 check('All scripts and Slide 8 prepared pause saved via Studio',await evaluate("[...document.querySelectorAll('#slides .script')].every(s=>s.value.length>0)&&document.querySelector('[data-number=\"8\"] .pause').value==='6'"));
 await ps('import');}
 else{await evaluate(`localStorage.setItem('lectureforge-project','${id}')`);}
 await command('Page.reload');await until("document.getElementById('progress').textContent.includes('14 / 14 slides ready')");
 if(process.argv.includes('--reload-worker')){
  await until('avatarState?.authorized===14');const before=await evaluate('avatarState.jobs.map(j=>j.id).sort().join()');
  await click('stop');for(let i=0;i<120&&(await runtime()).status!=='stopped';i++)await sleep(500);
  if((await runtime()).status!=='stopped')throw Error('Worker did not stop cleanly');
  await launch();await command('Page.reload');await until('avatarState?.authorized===14');
  check('Mid-production restart preserves existing authorized job identities',before===await evaluate('avatarState.jobs.map(j=>j.id).sort().join()'));
 }
 if(!checks.some(c=>c.name==='All 14 instructor fixture selections made through review controls')){
 await click('generate');await until("document.getElementById('message').textContent.includes('No new takes needed')");check('Generate 3 Takes safely reuses all 42 verified fixture takes',true);
 await click('review');await until("document.querySelectorAll('#takes audio').length===3");
 const reviewWindow=await command('Browser.getWindowForTarget');
 await command('Browser.setWindowBounds',{windowId:reviewWindow.windowId,bounds:{windowState:'normal'}});
 await command('Page.bringToFront');
 // Audition each player with a user gesture; background tabs may defer preload.
 for(let take=0;take<3;take++){
  await evaluate(`document.querySelectorAll('#takes audio')[${take}].play()`);
  await until(`document.querySelectorAll('#takes audio')[${take}].currentTime>0`,1);
  await evaluate(`document.querySelectorAll('#takes audio')[${take}].pause()`);
 }
 check('Three browser audio players load durations',await evaluate("[...document.querySelectorAll('#takes audio')].every(a=>a.duration>0)"));
 await command('Page.bringToFront');
 await evaluate("window.rehearsalAudio=document.querySelector('#takes audio');window.rehearsalAudio.play()");
 await until('window.rehearsalAudio.currentTime>0',1);
 console.log('AUDIO',await evaluate("JSON.stringify({same:window.rehearsalAudio===document.querySelector('#takes audio'),time:window.rehearsalAudio.currentTime,paused:window.rehearsalAudio.paused,error:window.rehearsalAudio.error?.message,ready:window.rehearsalAudio.readyState})"));
 check('Narration plays in browser',await evaluate("window.rehearsalAudio.currentTime>0"));
 await evaluate("const a=document.querySelector('#takes audio');a.pause();a.currentTime=a.duration/2");await sleep(1500);check('Browser audio seeking works',await evaluate("const a=document.querySelector('#takes audio');Math.abs(a.currentTime-a.duration/2)<1"));
 await evaluate("document.querySelectorAll('#takes button')[1].click()");await until("document.getElementById('selection').textContent==='Selected: Take 2'");await click('change-selection');await until("document.getElementById('selection').textContent==='No take selected'");check('Dedicated Change Selection clears intent and preserves three takes',await evaluate("document.querySelectorAll('#takes audio').length===3"));
 await evaluate("document.querySelectorAll('#takes button')[0].click()");await until("document.getElementById('selection').textContent==='Selected: Take 1'");check('Different take can be selected freely',true);
 const fixture=read(path.join(root,'fixture.json'));
 for(let slide=1;slide<=14;slide++){
  await evaluate(`const s=document.getElementById('review-slide');s.value='${slide}';s.dispatchEvent(new Event('change'))`);
  await until(`document.getElementById('review-title').textContent.startsWith('Slide ${slide} -')&&document.querySelectorAll('#takes button').length===3`);
  await evaluate(`document.querySelectorAll('#takes button')[${fixture.expected[slide]-1}].click()`);await until(`document.getElementById('selection').textContent==='Selected: Take ${fixture.expected[slide]}'`);
 }
 check('All 14 instructor fixture selections made through review controls',true);}
 if(!checks.some(c=>c.name==='Explicit confirmation opens background dashboard')){
 await click('review-selections');await until("!document.getElementById('selection-summary').hidden&&document.querySelectorAll('#selection-rows tr').length===14&&document.getElementById('selection-cost').textContent.includes('14 reusable')");
 check('Summary shows zero new jobs and 14 reusable assets',await evaluate("document.getElementById('selection-cost').textContent.includes('0 NEW HeyGen jobs')&&document.getElementById('selection-cost').textContent.includes('14 reusable')"));
 check('Slide 4 Take 2 and Slide 8 6.000-second pause visible',await evaluate("document.querySelectorAll('#selection-rows tr')[3].textContent.includes('Take 2')&&document.querySelectorAll('#selection-rows tr')[7].textContent.includes('6.000 s prepared pause')"));
 check('Selection alone created no authorization',await evaluate("document.getElementById('authorization-status').textContent.includes('not authorized')"));
 await evaluate("document.querySelector('[aria-label=\"Change Slide 4 narration\"]').click()");
 await until("document.getElementById('review-title').textContent.startsWith('Slide 4 -')&&!document.getElementById('review-panel').hidden");
 check('Summary CHANGE returns to intended slide',true);
 await click('review-selections');
 await until("!document.getElementById('selection-summary').hidden&&!document.getElementById('authorize-preview').disabled");
 await click('authorize-preview');
 await until("!document.getElementById('authorization-confirmation').hidden&&!document.getElementById('confirm-avatars').disabled");
 check('Final confirmation shows exact 14 mappings and zero new jobs',await evaluate("document.querySelectorAll('#confirmation-mapping li').length===14&&document.getElementById('confirmation-cost').textContent.includes('0 NEW HeyGen jobs')"));
 await until("!document.getElementById('confirm-avatars').disabled");await click('confirm-avatars');await until("!document.getElementById('avatar-dashboard').hidden");check('Explicit confirmation opens background dashboard',true);
 }
 await until("!document.getElementById('avatar-dashboard').hidden");
 if(process.argv.includes('--recovery-checks')&&!checks.some(c=>c.name==='Transient lock exceptions recover through local-only UI retry')){
  await until('avatarState?.ready===14',25);
  const before=await evaluate('avatarState.jobs.map(j=>j.id).sort().join()');
  await stopFixture();await ps('recover-local');await launch();await command('Page.reload');
  await until('avatarState?.authorized===14&&avatarState.exceptions===1');
  check('Mid-production restart preserves existing authorized job identities',before===await evaluate('avatarState.jobs.map(j=>j.id).sort().join()'));
 }
 const productionDeadline=Date.now()+25*60000;let retries=0;
 while(Date.now()<productionDeadline){
  const state=await evaluate('avatarState');
  if(state?.ready===14)break;
  if(state?.exceptions&&!state.jobs.some(j=>!['Exception','Avatar Ready'].includes(j.stage))){
   const job=state.jobs.find(j=>j.stage==='Exception');
   if((!job.output_path&&!job.row.reusable_avatar)||!job.error.includes('Narration state busy')||retries>=14)throw Error('Unexpected local exception: '+job.error);
   await evaluate(`const row=[...document.querySelectorAll('#avatar-jobs article')].find(r=>r.querySelector('details').dataset.job===${JSON.stringify(job.id)});const b=[...row.querySelectorAll('button')].find(b=>b.textContent==='Retry local validation');b.id='rehearsal-local-retry'`);
   await click('rehearsal-local-retry');retries++;
   await until(`avatarState.jobs.find(j=>j.id===${JSON.stringify(job.id)}).stage!=='Exception'`);
   console.log('Retried local validation through UI for Slide '+job.slide);
  }
  await sleep(3000);
 }
 check('Real worker reuses and validates all 14 existing avatars',await evaluate('avatarState.ready===14'));
 if(retries)check('Transient lock exceptions recover through local-only UI retry',true);
 if(process.argv.includes('--recovery-checks')&&!checks.some(c=>c.name==='Interrupted local assembly recovers through Retry Assembly UI')){
  await stopFixture();await ps('recover-assembly');await launch();await command('Page.reload');
  await until('avatarState?.ready===14&&assemblyState?.current?.stage==="Exception"');
 }
 await click('check-assembly');await until("!document.getElementById('create-recording').disabled&&assemblyState!==null");await evaluate(`document.getElementById('output-folder').value=${JSON.stringify(path.join(root,'Output'))}`);
 const assemblyRetry=await evaluate("!document.getElementById('retry-assembly').hidden");
 await click(assemblyRetry?'retry-assembly':'create-recording');
 await until("document.getElementById('assembly-progress').textContent==='RECORDING POWERPOINT READY'",25);check('One-click assembly finishes through live Studio',true);
 if(assemblyRetry)check('Interrupted local assembly recovers through Retry Assembly UI',true);
 await click('open-recording');await click('open-output');await sleep(2000);check('Open Recording and Output Folder controls succeed',await evaluate("!document.getElementById('message').classList.contains('error')"));
 await ps('capture');await stopFixture();
 for(let i=0;i<120&&(await runtime()).status!=='stopped';i++)await sleep(500);check('Studio Stop cleanly stops server and worker',(await runtime()).status==='stopped');
 await launch();await command('Page.reload');await until("document.getElementById('assembly-progress').textContent==='RECORDING POWERPOINT READY'");check('Simple launcher restart restores same project and Recording',await evaluate(`localStorage.getItem('lectureforge-project')==='${id}'`));
 await click('back');await click('open');await until("document.querySelectorAll('#project-list button').length===1");await evaluate("document.querySelector('#project-list button').click()");await until("document.getElementById('assembly-progress').textContent==='RECORDING POWERPOINT READY'");check('Open Existing Lecture recovers completed workflow',true);
 await ps('verify');check('Exact selections, authorization, ready jobs and Recording hashes survive restart',true);check('Source and all paid fixture assets unchanged; provider calls zero',true);
 fs.writeFileSync(path.join(root,'results.json'),JSON.stringify({project_id:id,project_root:path.join(root,'Projects',id),url,checks,elevenlabs_calls:0,heygen_calls:0},null,2));
 const shot=await command('Page.captureScreenshot',{format:'png'});fs.writeFileSync(path.join(root,'dress-rehearsal.png'),Buffer.from(shot.data,'base64'));
 }finally{for(const p of pending.values())clearTimeout(p.timer);pending.clear();if(ws)ws.close();if(edge)edge.kill();}
})().catch(e=>{console.error(e.message);process.exitCode=1});
