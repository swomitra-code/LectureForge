'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert/strict'),path=require('path');
class Element{constructor(){this.disabled=false;this.hidden=true;this.textContent='';this.value='';}}
const elements=new Map(),get=id=>{if(!elements.has(id))elements.set(id,new Element());return elements.get(id);};
const calls=[];let answer={},selection={state:'selected',path:'C:\\Slides\\source.pptx'},error=null,saves=0;
const ctx=vm.createContext({$:get,console,setTimeout:fn=>fn(),project:{id:'fixture',source:{external_path:'C:\\Slides\\source.pptx',sha256:'old'}},dirty:false,reviewNumber:1,
 action:async fn=>{try{return await fn();}catch(e){ctx.message(e.message,true);}},message:(text,isError)=>{get('message').textContent=text;get('message').error=isError;},save:async()=>{saves++;},
 renderProject:p=>ctx.project=p,checkAssembly:async()=>{},loadAssembly:async()=>{},
 request:async(url,body)=>{calls.push({url,body});if(error)throw Error(error);if(url==='/api/source-choose')return {choice:'token'};if(url.startsWith('/api/source-choice'))return selection;
  if(url.startsWith('/api/source-compare'))return answer;
  if(url.startsWith('/api/source-reload'))return {project:{id:'fixture',source:{external_path:answer.path,sha256:'new'}},message:'Source reloaded. Assets preserved.'};
  if(url==='/api/new-source')return {id:'imported',source:{external_path:selection.path}};
  if(url.startsWith('/api/source-folder'))return {message:'Opened Windows Explorer: C:\\Slides\\source.pptx'};
  throw Error('Unexpected endpoint '+url);
 }});
const repo=path.resolve(__dirname,'../..');vm.runInContext(fs.readFileSync(path.join(repo,'studio/web/source.js'),'utf8'),ctx);
const run=code=>vm.runInContext(code,ctx);
function comparison(compatible=true){return {path:selection.path,old:{sha256:'OLD_SHA',slide_count:34,width_emu:100,height_emu:50},new:{sha256:'NEW_SHA',slide_count:34,width_emu:100,height_emu:50},ordered_ids_match:compatible,dimensions_match:true,compatible,problems:compatible?[]:['Ordered slide identities changed.'],changed:true,expected:'binding'};}
(async()=>{
 answer={...comparison(),changed:false};answer.new.sha256=answer.old.sha256;
 ctx.dirty=true;
 await get('reload-source').onclick();
 assert.equal(get('message').textContent,'Source PowerPoint is already up to date.');
 assert.equal(get('source-comparison').hidden,true,'Unchanged reload exits without a confirmation');
 await get('source-confirm').onclick();
 assert.equal(calls.filter(c=>c.url.startsWith('/api/source-reload')).length,0,'Unchanged reload never calls the mutation endpoint');
 assert.equal(get('reload-source').disabled,false,'No-op releases the UI busy state');
 assert.equal(saves,0,'No-op does not implicitly save unrelated narration edits');
 assert.equal(ctx.dirty,true,'Unsaved edits remain available');ctx.dirty=false;
 answer=comparison();await get('reload-source').onclick();
 assert.equal(calls.filter(c=>c.url==='/api/source-choose').length,0,'Known path reload uses no picker');
 assert.equal(calls.filter(c=>c.url.startsWith('/api/source-reload')).length,0,'Comparison never applies automatically');
 assert.equal(get('source-old-sha').textContent,'OLD_SHA');assert.equal(get('source-new-sha').textContent,'NEW_SHA');
 assert.match(get('source-counts').textContent,/34 before.*34 now/);assert.equal(get('source-order').textContent,'Match');assert.match(get('source-dimensions').textContent,/match/);
 assert.match(get('source-confirmation').textContent,/Existing narration and avatar assets will be preserved/);
 get('source-cancel').onclick();await get('source-confirm').onclick();assert.equal(calls.filter(c=>c.url.startsWith('/api/source-reload')).length,0);
 answer=comparison(false);await get('reload-source').onclick();assert.equal(get('source-confirm').hidden,true);await get('source-confirm').onclick();assert.equal(calls.filter(c=>c.url.startsWith('/api/source-reload')).length,0);
 answer=comparison();await get('reload-source').onclick();await get('source-confirm').onclick();
 const apply=calls.find(c=>c.url.startsWith('/api/source-reload'));assert.equal(apply.body.confirm,true);assert.equal(apply.body.expected,'binding');assert.equal(ctx.project.source.sha256,'new');
 ctx.project={id:'fixture',source:{}};await get('reload-source').onclick();assert.ok(calls.some(c=>c.url==='/api/source-choose'));
 const compared=calls.filter(c=>c.url.startsWith('/api/source-compare')).at(-1);assert.equal(compared.body.path,selection.path);
 const count=calls.length;selection={state:'cancelled'};await get('choose-source').onclick();assert.equal(calls.slice(count).some(c=>c.url.startsWith('/api/source-compare')),false);
 error='Source missing. Use Choose Different Source PowerPoint.';await get('source-folder').onclick();assert.match(get('message').textContent,/Choose Different/);error=null;
 selection={state:'selected',path:'C:\\Slides\\moved.pptx'};await get('choose-import-source').onclick();get('name').value='New lecture';get('preset').value='renewable-energy';await run('createSourceLecture()');
 assert.equal(calls.at(-1).url,'/api/new-source');assert.equal(calls.at(-1).body.choice,'token');assert.equal(ctx.project.source.external_path,selection.path);
 assert.equal(calls.some(c=>/avatar-authorize|narration-generate|avatar-retry|source-remap/.test(c.url)),false);
 console.log('PASS: reload comparison/confirmation/cancel, incompatibility, legacy picker, native import, folder feedback; zero paid actions');
})().catch(e=>{console.error(e);process.exitCode=1;});
