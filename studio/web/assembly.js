'use strict';
let assemblyReady=null,assemblyProject=null,assemblyBusy=false,assemblySubmitting=false,assemblyState=null,choiceTimer=null;
function placementForm(){if(!assemblyReady)return;const n=Number($('placement-slide').value);const p=assemblyReady.placement.overrides.find(x=>x.slide===n)?.placement||assemblyReady.default_placement;for(const k of ['left','top','width'])$('placement-'+k).value=(p[k]*100).toFixed(4);}
async function checkAssembly(refreshPlacement=true){
 if(!project||stopped)return;
 const id=project.id,d=await request('/api/assembly-ready?id='+id);
 if(project?.id!==id)return;
 assemblyReady=d;assemblyProject=id;
 $('assembly-gate').textContent=d.ready?`${d.snapshot.slides.length} slides ready for assembly`:d.errors.join('\n');
 const active=assemblyState?.current&&!['Complete','Exception'].includes(assemblyState.current.stage);
 $('create-recording').disabled=dirty||active||!d.ready;
 if(refreshPlacement){
  const selected=$('placement-slide').value;
  $('placement-slide').replaceChildren();
  for(const s of project.slides.filter(s=>s.enabled)){const o=document.createElement('option');o.value=s.number;o.textContent='Slide '+s.number;$('placement-slide').append(o);}
  if([...$('placement-slide').options].some(o=>o.value===selected))$('placement-slide').value=selected;
  placementForm();
 }
}
async function loadAssembly(){
 if(!project||stopped||assemblyBusy)return;
 assemblyBusy=true;const id=project.id;
 try{
  const changedProject=assemblyProject!==id;
  if(changedProject){assemblyReady=null;assemblyState=null;assemblyProject=id;await checkAssembly();}
  const d=await request('/api/assembly-status?id='+id);
  if(project?.id!==id)return;
  assemblyState=d;const c=d.current,active=c&&!['Complete','Exception'].includes(c.stage);
  const stale=c?.stage==='Complete'&&assemblyReady&&c.binding!==assemblyReady.binding;
  $('assembly-progress').textContent=stale?'Previous recording available. Inputs changed; create a new recording.':c?(c.stage==='Complete'?'RECORDING POWERPOINT READY':c.stage+(c.slide?' - Slide '+c.slide:'')+(c.error?' - '+c.error:'')):'';
  $('create-recording').disabled=dirty||active||!assemblyReady?.ready;
  $('retry-assembly').hidden=c?.stage!=='Exception';
  $('recording-path').textContent=d.last_success?.output||'';
  $('open-recording').hidden=$('open-output').hidden=!d.last_success;
  if(changedProject)$('output-folder').value=c?.folder||'';
 }catch(e){message(e.message,true);}finally{assemblyBusy=false;}
}
async function createRecording(){
 if(assemblySubmitting)return;
 let submitted=false;
 assemblySubmitting=true;$('create-recording').disabled=true;$('retry-assembly').disabled=true;
 $('assembly-progress').textContent='Validating readiness...';message('Validating Recording PowerPoint inputs.');
 try{
  const inputsChanged=dirty||assemblyProject!==project.id||!assemblyReady;await save();if(inputsChanged)await checkAssembly();
  if(!assemblyReady.ready)throw Error(assemblyReady.errors.join('; ')||'Recording PowerPoint inputs are not ready.');
  $('assembly-progress').textContent='Readiness passed. Submitting assembly request...';
  const accepted=await request('/api/assembly-create?id='+project.id,{expected:assemblyReady.binding,folder:$('output-folder').value});
  submitted=true;
  assemblyState=accepted;$('assembly-progress').textContent=accepted.current?.stage||'Assembly request accepted';
  message('Assembly request accepted. Progress will update here.');await loadAssembly();
 }catch(e){$('assembly-progress').textContent='Assembly request failed - '+e.message;throw e;}
 finally{assemblySubmitting=false;$('retry-assembly').disabled=false;if(submitted)await loadAssembly();else $('create-recording').disabled=dirty||!assemblyReady?.ready;}
}
$('check-assembly').onclick=()=>action(checkAssembly);$('create-recording').onclick=$('retry-assembly').onclick=()=>action(createRecording);$('placement-slide').onchange=placementForm;
async function savePlacement(reset){if(!assemblyReady)await checkAssembly();const width=Number($('placement-width').value)/100;const p=reset?null:{left:Number($('placement-left').value)/100,top:Number($('placement-top').value)/100,width,height:width*project.source.width_emu/project.source.height_emu*560/640};await request('/api/assembly-placement?id='+project.id,{slide:Number($('placement-slide').value),placement:p,version:assemblyReady.placement.version});await checkAssembly();message('Placement saved. Narration and avatar media are unchanged.');}
$('placement-save').onclick=()=>action(()=>savePlacement(false));$('placement-reset').onclick=()=>action(()=>savePlacement(true));
$('open-recording').onclick=()=>action(()=>request('/api/recording-open?id='+project.id,{kind:'deck'}));$('open-output').onclick=()=>action(()=>request('/api/recording-open?id='+project.id,{kind:'folder'}));
$('choose-output').onclick=()=>action(async()=>{const id=project.id;await request('/api/output-choose?id='+id,{});clearInterval(choiceTimer);let count=0;choiceTimer=setInterval(async()=>{try{const d=await request('/api/output-choice?id='+id);if(++count>120||['selected','cancelled'].includes(d.state)){clearInterval(choiceTimer);if(project?.id===id&&d.state==='selected')$('output-folder').value=d.path;}}catch{clearInterval(choiceTimer);}},1000);});
setInterval(loadAssembly,2000);
