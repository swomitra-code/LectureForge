// Real HTTP routes + production UI handlers in a minimal DOM, without a worker.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {spawn} = require('node:child_process');
const net = require('node:net');
const fixture = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const repo = path.resolve(__dirname, '../..');
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
class Element {
 reportValidity(){return true;} pause(){} removeAttribute(key){delete this[key];} getAttribute(key){return this[key];}
  constructor() { this.children=[]; this.dataset={}; this.value=''; this.textContent=''; this.classList={toggle(){}}; }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children=children; }
  addEventListener() {}
  querySelector(selector) { return this.children.find(c=>c.className===selector.slice(1)) || this.children.map(c=>c.querySelector?.(selector)).find(Boolean); }
}
(async()=>{
 const probe=net.createServer(); await new Promise(resolve=>probe.listen(0,'127.0.0.1',resolve));
 const port=probe.address().port; await new Promise(resolve=>probe.close(resolve));
 const server=spawn('powershell.exe',['-NoProfile','-ExecutionPolicy','Bypass','-File',path.join(repo,'studio/Server.ps1'),'-RuntimeRoot',fixture.root,'-Port',String(port),'-Instance','script-import-test','-WorkerPid',String(process.pid)],{windowsHide:true,env:{...process.env,ELEVENLABS_API_KEY:'',HEYGEN_API_KEY:''},stdio:'pipe'});
 let diagnostic=''; server.stderr.on('data',d=>diagnostic+=d); server.stdout.resume();
 const base=`http://127.0.0.1:${port}`;
 try {
  let boot;
  for(let i=0;i<80;i++){try{boot=await (await fetch(base+'/api/bootstrap')).json();break;}catch{await sleep(100);}}
  assert.ok(boot,'Server started: '+diagnostic);
  const allowed=new Set(['/api/bootstrap','/api/new','/api/save','/api/narration','/api/project']);
  const calls=[];
  async function transport(url,options){
   assert.ok(allowed.has(new URL(url,base).pathname),'Unexpected or paid endpoint: '+url);
   calls.push(url); return fetch(new URL(url,base),options);
  }
  const created=await transport('/api/new?name=HTTP-import&filename=input.pptx&preset=renewable-energy',{method:'POST',headers:{'X-LF-Token':boot.token,'Content-Type':'application/octet-stream'},body:fs.readFileSync(fixture.source)});
  assert.equal(created.status,200); const project=await created.json();
  const elements=new Map(); const get=id=>{if(!elements.has(id))elements.set(id,new Element());return elements.get(id);};
  const context=vm.createContext({document:{getElementById:get,createElement:()=>new Element(),createTextNode:text=>({textContent:text}),querySelectorAll:()=>[]},window:{addEventListener(){}},localStorage:{getItem:()=>null,setItem(){},removeItem(){}},fetch:transport,setInterval(){},URLSearchParams,confirm(){throw Error('Import must not ask for generation');}});
  vm.runInContext(fs.readFileSync(path.join(repo,'studio/web/studio.js'),'utf8'),context);
  const run=code=>vm.runInContext(code,context);
  context.inputProject=project;
  run(`token=${JSON.stringify(boot.token)};renderProject(inputProject)`);
  async function settled(){
   for(let i=0;i<80;i++){if(run('narration!==null'))break;await sleep(25);}
   assert.equal(get('message').textContent.includes('not a function'),false);
   assert.ok(run('narration!==null'),'Narration loaded: '+get('message').textContent);
  }
  await settled();
  const minimal="Slide 1:\nHere is the question for this lesson.\n\nSlide 2:\nNow let's make the puzzle harder.";
  async function importText(text){get('script-import').value=text;await get('import').onclick();await settled();assert.equal(get('message').textContent,'Scripts imported and saved.');}
  await importText(minimal);
  assert.equal(run('project.slides[0].script'),'Here is the question for this lesson.');
  assert.ok(run('narration.slides.every(s=>Array.isArray(s.takes)&&s.takes.length===0)'));
  assert.equal(get('takes').children.length,3);
  console.log('PASS: new lecture -> Import click -> HTTP save -> narration JSON -> progress and review render');
  const text=Array.from({length:20},(_,i)=>`Slide ${i+1}:\nNarration ${i+1}.`).join('\n\n');
  await importText(text); const before=run('JSON.stringify(project.slides)');await importText(text);
  assert.equal(run('JSON.stringify(project.slides)'),before);
  assert.ok(run('project.slides.every(s=>s.script===`Narration ${s.number}.`)'));
  console.log('PASS: twenty scripts and repeated HTTP/UI import preserve all slide bindings');
  context.inputProject=await (await transport('/api/project?id='+fixture.id)).json();run('renderProject(inputProject)');await settled();
  const historyPath=path.join(fixture.root,'Projects',fixture.id,'narration.json'); const history=fs.readFileSync(historyPath);
  await importText(minimal);assert.deepEqual(fs.readFileSync(historyPath),history);
  assert.equal(run('narration.slides[0].takes[0].number'),2);
  console.log('PASS: existing lecture import retains take 2 and renders its audio card');
  for(const value of [null,{},'legacy',{number:2,state:'failed',attempts:[]},{2:{number:2,state:'failed',attempts:[]}},[null,'legacy',{number:1,state:'failed',attempts:[]}]] ){
   context.legacyValue=value;
   run('narration.slides[0].takes=normalizeTakes(legacyValue);reviewSignature="";renderNarration()');
   assert.equal(get('takes').children.length,3);
  }
  console.log('PASS: legacy response values cannot crash progress or review rendering');
  assert.ok(calls.every(url=>allowed.has(new URL(url,base).pathname)));
  const files=fs.readdirSync(path.join(fixture.root,'Projects',project.id)).sort();
  assert.deepEqual(files,['project.json','source.pptx']);
  console.log('PASS: import calls only local project/save/narration routes; no provider requests or jobs');
 } finally {
  // Terminate only the isolated server; no worker was launched or provider enabled.
  server.kill(); await new Promise(resolve=>server.exitCode!==null?resolve():server.once('exit',resolve));
 }
})().catch(error=>{console.error(error);process.exitCode=1;});
