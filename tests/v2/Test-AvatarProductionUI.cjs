// Exercise shipped UI handlers without a browser, worker, or provider transport.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const web=path.resolve(__dirname,'../../studio/web');
class Element {
 constructor(){this.children=[];this.dataset={};this.style={};this.value='';this.hidden=false;}
 append(...items){this.children.push(...items);} replaceChildren(...items){this.children=items;}
 setAttribute(){} scrollIntoView(){} querySelectorAll(){return [];}
 get options(){return this.children;}
}
const elements=new Map(),get=id=>{if(!elements.has(id))elements.set(id,new Element());return elements.get(id);};
const preset=JSON.parse(fs.readFileSync(path.join(web,'../presets/renewable-energy.json'),'utf8'));
const slides=[1,2].map(number=>({number,title:'Slide '+number,enabled:true,selected_take:2,status:'narration selected',narration:{duration_seconds:10},new_jobs:1}));
const snapshot={slides,can_authorize:true,expected_new_provider_jobs:2,reusable_avatars:0,excluded_slides:[3],avatar_preset:{avatar:preset.avatar}};
let review={current:{snapshot,binding_sha256:'binding'},saved:null,review:null,authorization_status:'not authorized',review_stale:false};
let avatars={authorized:0,ready:0,exceptions:0,jobs:[],queued:0,provider_processing:0};
const calls=[],placement={version:1,overrides:[]};let assembly={current:null,last_success:null},assemblyFailure=null;
async function request(url,body){
 calls.push({url,body});
 if(url.startsWith('/api/selections'))return review;
 if(url.startsWith('/api/selection-review')){review.saved={phase:body.phase};review.review={id:'binding',snapshot:structuredClone(snapshot)};return review;}
 if(url.startsWith('/api/avatar-authorize')){assert.equal(body.confirm,true);assert.equal(review.saved.phase,'confirmation');review.authorization_status='authorized';return {};}
 if(url.startsWith('/api/avatar-status'))return avatars;
 if(url.startsWith('/api/avatar-replacement')&&body){assert.equal(body.confirm,true);assert.equal(body.job,'2');avatars.jobs[1].stage='Queued';avatars.jobs[1].provider_status=null;avatars.exceptions=0;avatars.queued=1;return {job:'replacement'};}
 if(url.startsWith('/api/avatar-replacement'))return {job:'2',slide:2,title:'Slide 2',selected_take:2,avatar_preset:{avatar:preset.avatar},expected_new_provider_jobs:1,reusable_avatars:0,paid_replacement:true,selection_binding:'binding'};
 if(url.startsWith('/api/assembly-ready'))return {ready:avatars.ready===2&&avatars.exceptions===0,errors:['Avatars not ready'],binding:'assembly-binding',snapshot:{slides},placement,default_placement:preset.placement};
 if(url.startsWith('/api/assembly-create')){if(assemblyFailure)throw Error(assemblyFailure);await new Promise(r=>setTimeout(r,10));if(!assembly.current)assembly.current={id:'attempt-1',binding:body.expected,stage:'Queued',slide:0};return assembly;}
 if(url.startsWith('/api/assembly-status'))return assembly;
 if(url.startsWith('/api/assembly-placement')){placement.overrides=body.placement?[{slide:body.slide,placement:body.placement}]:[];return {};}
 throw Error('Unexpected request '+url);
}
const context=vm.createContext({$,document:{createElement:()=>new Element()},project:{id:'legacy',slides,source:{width_emu:1600,height_emu:900}},stopped:false,dirty:false,reviewNumber:1,request,save:async()=>{},action:fn=>fn(),message(){},renderReview(){},setInterval(){},clearInterval(){},setTimeout});
function $(id){return get(id);}
const run=code=>vm.runInContext(code,context);
for(const file of ['selections.js','avatars.js','assembly.js'])run(fs.readFileSync(path.join(web,file),'utf8'));
(async()=>{
 const html=fs.readFileSync(path.join(web,'index.html'),'utf8');
 assert.ok(html.indexOf('id="selection-summary"')<html.indexOf('id="assembly-panel"'));
 assert.ok(!html.includes('id="selection-summary" hidden'));
 await run('loadSelectionReview(true)');
 assert.equal(get('selection-summary').hidden,false);assert.equal(get('authorize-preview').disabled,false);
 assert.match(get('avatar-readiness').textContent,/2 \/ 2/);assert.match(get('avatar-configuration').textContent,/SKM-BLUE-7.*255, 255, 255/);
 assert.match(get('selection-cost').textContent,/2 NEW HeyGen/);assert.equal(get('selection-rows').children.length,2);
 assert.equal(get('authorization-confirmation').hidden,true);assert.equal(get('confirm-avatars').disabled,true);
 await run('loadAvatarStatus()');
 assert.equal(run('assemblyProject'),null,'Avatar polling must leave initial placement loading to assembly');
 assert.ok(calls.every(c=>c.body===undefined));
 snapshot.can_authorize=false;slides[0].status='missing narration selection';
 review.saved={phase:'narration',slide:1};await run('loadSelectionReview(true)');
 assert.equal(get('selection-summary').hidden,false);assert.equal(get('authorize-preview').disabled,true);assert.match(get('avatar-readiness').textContent,/1 \/ 2/);
 snapshot.can_authorize=true;slides[0].status='narration selected';await run('loadSelectionReview()');
 await get('check-assembly').onclick();assert.equal(get('create-recording').disabled,true);
 assert.ok(calls.every(c=>c.body===undefined),'Selection/readiness must be read-only');
 await get('authorize-preview').onclick();
 assert.equal(get('authorization-confirmation').hidden,false);assert.equal(get('confirmation-mapping').children.length,2);
 assert.match(get('confirmation-excluded').textContent,/3/);assert.equal(get('confirm-avatars').disabled,false);
 assert.ok(!calls.some(c=>c.url.includes('avatar-authorize')),'Summary appears before authorization');
 await get('confirm-avatars').onclick();
 assert.equal(calls.filter(c=>c.url.includes('avatar-authorize')).length,1);
 avatars={...avatars,authorized:2,authorization:'binding',queued:2,jobs:slides.map(s=>({id:String(s.number),slide:s.number,stage:'Queued',row:{narration:{}},history:[]}))};
 await run('loadAvatarStatus()');assert.equal(get('avatar-dashboard').hidden,false);assert.match(get('avatar-summary').textContent,/2 queued/);
 avatars.provider_processing=1;avatars.jobs[0].stage='HeyGen Processing';await run('loadAvatarStatus()');assert.match(get('avatar-summary').textContent,/1 provider processing/);
 avatars.ready=1;avatars.jobs[0].stage='Avatar Ready';await run('loadAvatarStatus()');assert.equal(get('create-recording').disabled,true);
 avatars.exceptions=1;avatars.jobs[1].stage='Exception';avatars.jobs[1].provider_status='failed';avatars.jobs[1].error='HeyGen job failed';await run('loadAvatarStatus()');assert.equal(get('create-recording').disabled,true);
 const failedRow=get('avatar-jobs').children[1],retry=failedRow.children.find(x=>x.textContent==='Retry / Re-authorize');assert.ok(retry,'Failed provider job exposes retry authorization');
 await retry.onclick();assert.equal(get('replacement-confirmation').hidden,false);assert.match(get('replacement-summary').textContent,/Slide 2.*Take 2.*1 NEW HeyGen job.*0 reused.*PAID/);assert.equal(calls.filter(c=>c.url.includes('avatar-replacement')&&c.body).length,0,'Summary makes no paid submission');
 await get('confirm-replacement').onclick();assert.equal(calls.filter(c=>c.url.includes('avatar-replacement')&&c.body).length,1,'Explicit confirmation queues one replacement');
 assert.equal(get('create-recording').disabled,true,'Recording remains gated while replacement is queued');
 avatars.exceptions=0;avatars.ready=2;avatars.jobs[1].stage='Avatar Ready';await run('loadAvatarStatus()');assert.equal(get('create-recording').disabled,false);
 get('placement-slide').value='1';get('placement-left').value='80';get('placement-top').value='75';get('placement-width').value='10';
 await get('placement-save').onclick();assert.equal(placement.overrides[0].placement.left,.8);assert.equal(placement.overrides[0].placement.width,.1);
 await get('placement-reset').onclick();assert.equal(placement.overrides.length,0);
 assert.equal(calls.filter(c=>c.url.includes('avatar-authorize')).length,1,'Placement/status must not authorize');
 const p1=get('create-recording').onclick(),p2=get('create-recording').onclick();await Promise.all([p1,p2]);
 assert.equal(calls.filter(c=>c.url.includes('/api/assembly-create')).length,1,'Repeated click submits one assembly request');
 assert.equal(assembly.current.id,'attempt-1');assert.match(get('assembly-progress').textContent,/Queued/);
 assembly.current=null;assemblyFailure='fixture backend exception';await get('create-recording').onclick().catch(()=>{});
 assert.match(get('assembly-progress').textContent,/failed.*fixture backend exception/i,'Backend exception is visible instead of an endless spinner');
 assert.equal(calls.filter(c=>c.url.includes('elevenlabs')||c.url.includes('heygen')).length,0,'Assembly makes no provider calls');
 console.log('PASS: persistent production UI, legacy navigation, missing/complete narration, preset/summary, explicit confirmation, read-only readiness, queue/partial/error/completed status, unchanged placement; no provider transport');
})().catch(e=>{console.error(e);process.exitCode=1;});

