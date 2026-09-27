const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
class Element{
 constructor(){this.children=[];this.dataset={};this.style={};this.value='';}
 append(...items){this.children.push(...items);} replaceChildren(){this.children=[];}
 querySelectorAll(){return [];} scrollIntoView(){} contains(el){return this===el||this.children.some(c=>c.contains?.(el));}
}
const elements=new Map(),$=id=>{if(!elements.has(id))elements.set(id,new Element());return elements.get(id);};
const calls=[],messages=[],job={id:'blocked',slide:16,take:2,stage:'Exception',exception_kind:'existing submission',row:{narration:{sha256:'hash'}},history:['Authorized','Queued']};
const document={activeElement:null,createElement:tag=>Object.assign(new Element(),{tagName:tag.toUpperCase()})};
const context=vm.createContext({$,document,project:{id:'fixture'},stopped:false,setInterval(){},action:fn=>fn(),message:m=>messages.push(m),request:async(url,body)=>{calls.push({url,body});if(url.startsWith('/api/avatar-status'))return {authorized:1,jobs:[job]};assert.equal(url,'/api/avatar-recover?id=fixture');job.stage='Downloading Avatar';return {stage:job.stage};}});
vm.runInContext(fs.readFileSync(path.join(__dirname,'../../studio/web/avatars.js'),'utf8'),context);
const run=()=>vm.runInContext('loadAvatarStatus()',context);
function all(el){return [el,...el.children.flatMap(all)];}
(async()=>{
 await run();let nodes=all($('avatar-jobs'));
 const toggle=nodes.find(n=>n.textContent==='Recover Existing HeyGen Video');assert.ok(toggle);toggle.onclick();
 let input=nodes.find(n=>n.type==='url');input.value='https://app.heygen.com/videos/'+'a'.repeat(32);input.oninput();
 document.activeElement=input;await run();assert.ok(all($('avatar-jobs')).includes(input),'Polling preserves input focus');
 document.activeElement=null;await run();nodes=all($('avatar-jobs'));input=nodes.find(n=>n.type==='url');assert.ok(input.value.endsWith('a'.repeat(32)),'Draft survives refresh');
 await nodes.find(n=>n.textContent==='Recover & Continue').onclick();
 const mutations=calls.filter(c=>c.body);assert.equal(mutations.length,1);assert.equal(mutations[0].body.job,'blocked');assert.equal(mutations[0].body.url,input.value);
 assert.match(messages.at(-1),/Existing HeyGen video recovered/);assert.equal(job.stage,'Downloading Avatar');
 console.log('PASS: recovery UI, draft/focus preservation, recovery-only request and success feedback');
})().catch(e=>{console.error(e);process.exitCode=1;});
