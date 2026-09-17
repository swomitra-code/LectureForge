// Production UI handlers and HTTP routes, with no worker or live provider calls.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const {spawn}=require('node:child_process'),net=require('node:net');
const fixture=JSON.parse(fs.readFileSync(process.argv[2],'utf8').replace(/^\uFEFF/,''));
const repo=path.resolve(__dirname,'../..'),sleep=ms=>new Promise(r=>setTimeout(r,ms));
class Element {
 reportValidity(){return true;} pause(){} removeAttribute(key){delete this[key];} getAttribute(key){return this[key];}
 constructor(){this.children=[];this.dataset={};this.classList={toggle(){}};this.textContent='';}
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
  const posts=[];let release;const gate=new Promise(r=>release=r);
  const allowed=new Set(['/api/bootstrap','/api/narration','/api/narration-regenerate']);
  async function transport(url,options){
   const pathname=new URL(url,base).pathname;assert.ok(allowed.has(pathname),'Unexpected endpoint: '+url);
   if(options?.method==='POST'){posts.push({url,body:JSON.parse(options.body)});await gate;}
   return fetch(new URL(url,base),options);
  }
  const p=await(await fetch(base+'/api/project?id='+fixture.id)).json();
  const context=vm.createContext({document:{getElementById:get,createElement:()=>new Element(),querySelectorAll:()=>[]},window:{addEventListener(){}},localStorage:{getItem:()=>null,setItem(){},removeItem(){}},fetch:transport,setInterval(){},URLSearchParams,confirm(){throw Error('Unexpected confirmation');}});
  vm.runInContext(fs.readFileSync(path.join(repo,'studio/web/studio.js'),'utf8'),context);
  const run=code=>vm.runInContext(code,context);context.fixtureProject=p;
  run(`token=${JSON.stringify(boot.token)};project=fixtureProject`);await run('loadNarration()');
  const cards=()=>get('takes').children,button=i=>cards()[i].children.find(x=>x.textContent==='Regenerate');
  assert.equal(cards().length,3);assert.ok([0,1,2].every(i=>button(i)&&!button(i).disabled));
  const oldSrc=cards()[0].children.find(x=>x.controls).src;
  assert.ok(oldSrc.includes('&asset='),'Playback URL is versioned by audio asset path');
  const b=button(0),pending=b.onclick();assert.equal(b.disabled,true);await b.onclick();assert.equal(posts.length,1);
  release();await pending;
  assert.equal(posts[0].body.take,1);assert.ok(posts[0].body.attempt);assert.equal(button(0).disabled,true);
  assert.equal(button(1).disabled,false);assert.equal(button(2).disabled,false);
  assert.ok(cards()[0].children.some(x=>x.textContent.includes('queued')));
  const response=await fetch(base+posts[0].url,{method:'POST',headers:{'X-LF-Token':boot.token,'Content-Type':'application/json'},body:JSON.stringify(posts[0].body)});
  assert.equal(response.ok,false,'Duplicate HTTP request rejected');
  run("narration.slides[0].takes[0].state='generating';renderReview()");assert.ok(cards()[0].children.some(x=>x.textContent.includes('Generating replacement')));
  run("narration.slides[0].takes[0].state='ready';narration.slides[0].takes[0].asset.path='replacement.mp3';narration.slides[0].takes[0].attempts.at(-1).previous_asset={};renderReview()");
  assert.notEqual(cards()[0].children.find(x=>x.controls).src,oldSrc);assert.ok(cards()[0].children.some(x=>x.textContent.includes('Regenerated successfully')));
  run("narration.slides[0].takes[0].state='uncertain';narration.slides[0].takes[0].attempts.at(-1).error='Fixture failure';renderReview()");
  assert.equal(button(0).disabled,true);assert.ok(cards()[0].children.some(x=>x.textContent==='Restore previous audio'));
  console.log('PASS: production UI/HTTP regeneration; immediate duplicate guard; queued/generating/success/error; cache refresh; recovery control; no HeyGen endpoint.');
 }finally{server.kill();await new Promise(r=>server.once('exit',r));}
})().catch(e=>{console.error(e);process.exitCode=1;});
