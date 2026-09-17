// Production UI handlers and HTTP routes, with no worker or live provider calls.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const {spawn}=require('node:child_process'),net=require('node:net');
const fixture=JSON.parse(fs.readFileSync(process.argv[2],'utf8').replace(/^\uFEFF/,''));
const repo=path.resolve(__dirname,'../..'),sleep=ms=>new Promise(r=>setTimeout(r,ms));
class Element {
 reportValidity(){return true;} pause(){} removeAttribute(key){delete this[key];} getAttribute(key){return this[key];}
 constructor(){this.children=[];this.dataset={};this.classList={toggle(){}};this.textContent='';}
 querySelector(selector){return this.children.find(c=>c.className===selector.slice(1))||this.children.map(c=>c.querySelector?.(selector)).find(Boolean);}
 append(...c){this.children.push(...c);} replaceChildren(...c){this.children=c;} addEventListener(){}
}
(async()=>{
 const probe=net.createServer();await new Promise(r=>probe.listen(0,'127.0.0.1',r));const port=probe.address().port;await new Promise(r=>probe.close(r));
 const server=spawn('powershell.exe',['-NoProfile','-ExecutionPolicy','Bypass','-File',path.join(repo,'studio/Server.ps1'),'-RuntimeRoot',fixture.root,'-Port',String(port),'-Instance','take-regression','-WorkerPid',String(process.pid)],{windowsHide:true,env:{...process.env,ELEVENLABS_API_KEY:'',HEYGEN_API_KEY:''},stdio:'pipe'});
 let diagnostic='';server.stderr.on('data',d=>diagnostic+=d);server.stdout.resume();
 const base=`http://127.0.0.1:${port}`;
 try{
  let boot;for(let i=0;i<120;i++){try{boot=await(await fetch(base+'/api/bootstrap')).json();break;}catch{await sleep(500);}}
  assert.ok(boot,'Server startup: '+diagnostic);
  const elements=new Map(),get=id=>{if(!elements.has(id))elements.set(id,new Element());return elements.get(id);};

  const posts=[];
  async function transport(url,options){if(options?.method==='POST')posts.push({url,body:JSON.parse(options.body)});return fetch(new URL(url,base),options);}
  const context=vm.createContext({document:{getElementById:get,createElement:()=>new Element(),createTextNode:text=>({textContent:text}),querySelectorAll:()=>[]},window:{addEventListener(){}},localStorage:{getItem:()=>null,setItem(){},removeItem(){}},fetch:transport,setInterval(){},URLSearchParams,confirm:()=>true});
  const run=code=>vm.runInContext(code,context);
  vm.runInContext(fs.readFileSync(path.join(repo,'studio/web/studio.js'),'utf8'),context);
  await run('refresh()');context.fixtureProject=await(await fetch(base+'/api/project?id='+fixture.id)).json();
  run('renderProject(fixtureProject)');await run('loadNarration()');
  assert.equal(Number(get('ns-stability').value),.34);assert.equal(Number(get('ns-takes_per_slide').value),3);
  get('ns-style').value=.9;get('narration-defaults').onclick();assert.equal(Number(get('ns-style').value),.5);assert.equal(run('dirty'),true);
  get('narration-preset').value='natural';await get('narration-preset').onchange();assert.equal(Number(get('ns-stability').value),.47);assert.equal(Number(get('ns-style').value),.38);
  await run('save()');const saved=await(await fetch(base+'/api/project?id='+fixture.id)).json();assert.equal(saved.preset.narration.stability,.47);
  get('ns-stability').value=.61;get('ns-takes_per_slide').value=5;
  const row=get('slides').children[0];row.querySelector('.script').value='Unsaved [REVEAL]HTTP preview.';
  const projectHash=fs.readFileSync(path.join(fixture.root,'Projects',fixture.id,'project.json'),'utf8');
  const pending=get('narration-test').onclick();assert.equal(get('narration-test').disabled,true);await get('narration-test').onclick();await pending;
  const tests=posts.filter(p=>p.url.startsWith('/api/narration-test'));assert.equal(tests.length,1);assert.equal(tests[0].body.settings.stability,.61);assert.equal(tests[0].body.slide,1);
  assert.equal(fs.readFileSync(path.join(fixture.root,'Projects',fixture.id,'project.json'),'utf8'),projectHash);
  const history=JSON.parse(fs.readFileSync(path.join(fixture.root,'Projects',fixture.id,'narration.json'),'utf8'));assert.equal(history.revisions.at(-1).takes.length,1);assert.equal(history.revisions.at(-1).script,'Unsaved [REVEAL]HTTP preview.');assert.ok(get('test-status').textContent.includes('queued'));
  // Display the previously completed real-executor preview without replacing DOM every poll.
  run('narration.previews.pop();renderTestTake()');assert.equal(get('test-audio').hidden,false);assert.ok(get('test-audio').src.includes(fixture.preview));
  const audio=await fetch(base+get('test-audio').src,{headers:{Range:'bytes=0-99'}});assert.equal(audio.status,206);assert.equal((await audio.arrayBuffer()).byteLength,100);
  const invalid=await transport('/api/narration-test?id='+fixture.id,{method:'POST',headers:{'X-LF-Token':boot.token,'Content-Type':'application/json'},body:JSON.stringify({...tests[0].body,settings:{...tests[0].body.settings,speed:5}})});assert.equal(invalid.status,400);
  assert.ok(posts.every(p=>!p.url.includes('avatar')));
  console.log('PASS: settings UI, restore/preset/save, unsaved preview, duplicate click guard, status/playback, real HTTP validation, no avatar routes');
 }finally{try{const b=await(await fetch(base+'/api/bootstrap')).json();await fetch(base+'/api/stop',{method:'POST',headers:{'X-LF-Token':b.token,'Content-Type':'application/json'},body:'{}'});}catch{}await sleep(300);if(server.exitCode===null)server.kill();}
})().catch(e=>{console.error(e);process.exitCode=1;});
