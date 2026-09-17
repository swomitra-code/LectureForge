'use strict';
const regeneratingTakes=new Set();
let token='',project=null,dirty=false,projects=[],narration=null,reviewNumber=1,reviewSignature='',polling=false,stopped=false,narrationDefaults=null,testSubmitting=false;
const $=id=>document.getElementById(id);
function message(text,error=false){$('message').textContent=text;$('message').classList.toggle('error',error);}
async function request(path,body,raw=false){const options=body===undefined?{}:{method:'POST',headers:{'X-LF-Token':token,'Content-Type':raw?'application/octet-stream':'application/json'},body:raw?body:JSON.stringify(body)};const r=await fetch(path,options);const d=await r.json();if(!r.ok)throw Error(d.error||'Request failed');return d;}
function changed(){dirty=true;$('saved').textContent='Unsaved changes';}
function renderProject(p){project=p;dirty=false;localStorage.setItem('lectureforge-project',p.id);$('home').hidden=true;$('project').hidden=false;$('project-name').textContent=p.name;$('source').textContent=`${p.source.original_name} · ${p.source.slide_count} slides · source SHA-256 ${p.source.sha256}`;$('preset-info').textContent=p.preset.name;$('saved').textContent='Saved';$('slides').replaceChildren();
 for(const s of p.slides){const row=document.createElement('article');row.className='slide'+(s.enabled?'':' disabled');row.dataset.number=s.number;const heading=document.createElement('h3');const label=document.createElement('label');const enabled=document.createElement('input');enabled.type='checkbox';enabled.checked=s.enabled;enabled.className='enabled';label.append(enabled,document.createTextNode(`Slide ${s.number} — ${s.title}`));heading.append(label);const scriptLabel=document.createElement('label');scriptLabel.textContent='Narration script';const script=document.createElement('textarea');script.rows=5;script.value=s.script;script.className='script';scriptLabel.append(script);const pauseLabel=document.createElement('label');pauseLabel.textContent='Silence after speech (seconds)';const pause=document.createElement('input');pause.type='number';pause.min='0';pause.max='120';pause.step='0.001';pause.value=s.post_speech_silence_seconds;pause.className='pause';pauseLabel.append(pause);row.append(heading,scriptLabel,pauseLabel);row.addEventListener('input',()=>{row.classList.toggle('disabled',!enabled.checked);changed();});$('slides').append(row);} renderNarrationSettings(); reviewSignature=''; narration=null; action(loadNarration);if(typeof loadSelectionReview==='function')action(()=>loadSelectionReview(true));}
async function openProject(id){renderProject(await request('/api/project?id='+encodeURIComponent(id)));message('Project loaded.');await request('/api/thumbnails?id='+id,{});}
async function refresh(){const d=await request('/api/bootstrap');token=d.token;projects=d.projects;narrationDefaults=d.presets[0].narration;$('root').textContent='Projects: '+d.root;$('preset').replaceChildren();for(const p of d.presets){const o=document.createElement('option');o.value=p.id;o.textContent=p.name;$('preset').append(o);}$('project-list').replaceChildren();for(const p of projects){const b=document.createElement('button');b.className='secondary';b.textContent=`${p.name} · ${p.slides} slides`;b.onclick=()=>action(()=>openProject(p.id));$('project-list').append(b);}if(!projects.length)$('project-list').textContent='No saved lectures yet.';}
async function action(fn){try{await fn();}catch(e){message(e.message,true);}}
async function save(){if(!project||!dirty)return;const slides=[...$('slides').children].map(r=>({number:Number(r.dataset.number),enabled:r.querySelector('.enabled').checked,script:r.querySelector('.script').value,post_speech_silence_seconds:Number(r.querySelector('.pause').value)}));renderProject(await request('/api/save?id='+project.id,{version:project.version,slides,narration_settings:readNarrationSettings()}));message('Changes saved.');}
$('new').onclick=()=>{$('new-form').hidden=false;$('projects').hidden=true;$('name').focus();};$('open').onclick=()=>action(async()=>{await refresh();$('projects').hidden=false;$('new-form').hidden=true;});
$('new-form').onsubmit=e=>{e.preventDefault();action(async()=>{const file=$('pptx').files[0];if(!file||!file.name.toLowerCase().endsWith('.pptx'))throw Error('Choose a .pptx file.');const button=e.target.querySelector('button');button.disabled=true;message('Copying and inspecting PowerPoint…');try{const q=new URLSearchParams({name:$('name').value,filename:file.name,preset:$('preset').value});renderProject(await request('/api/new?'+q,await file.arrayBuffer(),true));message('Lecture created. Import or edit its slide scripts.');}finally{button.disabled=false;}});};
$('save').onclick=()=>action(save);$('back').onclick=()=>action(async()=>{await save();project=null;localStorage.removeItem('lectureforge-project');$('project').hidden=true;$('home').hidden=false;await refresh();});
$('script-file').onchange=()=>action(async()=>{$('script-import').value=await $('script-file').files[0].text();});
$('import').onclick=()=>action(async()=>{await save();renderProject(await request('/api/save?id='+project.id,{version:project.version,script_import:$('script-import').value}));message('Scripts imported and saved.');});
for(const [id,value] of [['enable-all',true],['disable-all',false]])$(id).onclick=()=>{for(const row of $('slides').children){row.querySelector('.enabled').checked=value;row.classList.toggle('disabled',!value);}changed();};
$('folder').onclick=()=>action(()=>request('/api/folder?id='+project.id,{}));
$('stop').onclick=()=>action(async()=>{await save();await request('/api/stop',{});dirty=false;stopped=true;document.querySelectorAll('button,input,textarea,select').forEach(x=>x.disabled=true);message('LectureForge stopped. Your projects are saved. Double-click LectureForge.cmd to return.');});
window.addEventListener('beforeunload',e=>{if(dirty){e.preventDefault();e.returnValue='';}});
action(async()=>{await refresh();const id=localStorage.getItem('lectureforge-project');if(id&&projects.some(p=>p.id===id))await openProject(id);});

// Normalize at the response boundary for older servers/saved narration shapes.
function normalizeTakes(value){
 const items=Array.isArray(value)?value:value&&typeof value==='object'?(Object.hasOwn(value,'number')?[value]:Object.values(value)):[];
 return items.filter(t=>t&&typeof t==='object'&&[1,2,3,4,5].includes(t.number)&&typeof t.state==='string');
}
async function loadNarration(){if(!project||stopped)return;const id=project.id;const d=await request('/api/narration?id='+id);if(project?.id!==id)return;if(d.project_version!==project.version){if(dirty){message('Project changed in another window. Reload before reviewing or saving.',true);return;}renderProject(await request('/api/project?id='+id));return;}narration={...d,slides:d.slides.map(s=>({...s,takes:normalizeTakes(s.takes)}))};renderNarration();}
function renderNarration(){
 const enabled=narration.slides.filter(s=>s.enabled);const ready=enabled.filter(s=>s.takes.length>0&&s.takes.every(t=>t.state==='ready')).length;
 $('progress').textContent=`${ready} / ${enabled.length} slides ready | ${narration.provider_calls} ElevenLabs requests recorded`;
 $('slide-progress').replaceChildren();for(const s of enabled){const line=document.createElement('button');line.className='secondary';const active=s.takes.find(t=>t.state!=='ready');line.textContent=`Slide ${s.number} - ${!s.takes.length?'Awaiting generation':active?'Take '+active.number+' '+active.state:'Ready'}${s.selected_take?' | Selected Take '+s.selected_take:''}`;line.onclick=()=>{reviewNumber=s.number;$('review-panel').hidden=false;renderReview();};$('slide-progress').append(line);}
 renderReview();
}
function renderReview(){
 if(!project||!narration)return;
 const slides=project.slides.filter(s=>s.enabled);if(!slides.length)return;
 if(!slides.some(s=>s.number===reviewNumber))reviewNumber=slides[0].number;
 const s=slides.find(s=>s.number===reviewNumber),n=narration.slides.find(s=>s.number===reviewNumber);
 $('test-slide').value=String(reviewNumber);renderTestTake();
 $('review-slide').replaceChildren();for(const s of slides){const o=document.createElement('option');o.value=s.number;o.textContent='Slide '+s.number;o.selected=s.number===reviewNumber;$('review-slide').append(o);}
 $('previous').disabled=slides[0].number===reviewNumber;$('next').disabled=slides.at(-1).number===reviewNumber;
 $('review-title').textContent=`Slide ${s.number} - ${s.title}`;$('review-script').textContent=s.script;
 $('selection').textContent=n.selected_take?`Selected: Take ${n.selected_take}`:'No take selected';$('change-selection').disabled=!n.selected_take;
 const signature=JSON.stringify([project.id,s.number,n.revision,n.takes,n.selected_take,[...regeneratingTakes]]);
 if(signature===reviewSignature)return;reviewSignature=signature;
 $('thumbnail-status').textContent='';$('thumbnail').hidden=false;$('thumbnail').src=`/api/thumbnail?id=${project.id}&slide=${s.number}`;
 $('thumbnail').onerror=()=>{$('thumbnail').hidden=true;$('thumbnail-status').textContent='Slide preview is preparing or unavailable.';};
 $('takes').replaceChildren();
 for(let i=1;i<=(Math.max(project.preset.narration.takes_per_slide,...n.takes.map(t=>t.number)));i++){
  const t=n.takes.find(t=>t.number===i),card=document.createElement('article'),title=document.createElement('h4');title.textContent='Take '+i;card.append(title);
  if(t?.state==='ready'){
   const duration=document.createElement('p');duration.textContent=t.asset.duration_seconds.toFixed(3)+' seconds'+(t.asset.post_speech_silence_seconds?` | includes ${t.asset.post_speech_silence_seconds.toFixed(3)} s prepared silence`:'');
   const audio=document.createElement('audio');audio.controls=true;audio.preload='metadata';audio.src=`/api/audio?id=${project.id}&revision=${n.revision}&take=${i}&asset=${encodeURIComponent(t.asset.path)}`;audio.onplay=()=>document.querySelectorAll('audio').forEach(a=>{if(a!==audio)a.pause();});
   const select=document.createElement('button');select.textContent='Select Take '+i;select.onclick=()=>action(()=>selectTake(i));card.append(duration,audio,select);
  }else{const status=document.createElement('p');status.textContent=t?.state||'Not generated';card.append(status);if(t?.attempts?.at(-1)?.error){const error=document.createElement('pre');error.textContent=t.attempts.at(-1).error;card.append(error);}if(t?.state==='failed'){const retry=document.createElement('button');retry.textContent='Retry only Take '+i;retry.onclick=()=>action(async()=>{if(confirm(`Retry Slide ${s.number} Take ${i}? One ElevenLabs generation. Other takes are preserved.`)){await request('/api/narration-retry?id='+project.id,{revision:n.revision,take:i,confirm:true});await loadNarration();}});card.append(retry);}}
  if(t?.asset){
   const id=project.id,key=`${id}/${n.revision}/${i}`,attempt=t.attempts?.at(-1)||{};
   const busy=regeneratingTakes.has(key)||['queued','dispatched','generating'].includes(t.state);
   const regenerate=document.createElement('button');regenerate.textContent='Regenerate';
   regenerate.disabled=!attempt.id||busy||t.state!=='ready'||attempt.state==='uncertain';
   regenerate.onclick=()=>action(async()=>{
    if(regeneratingTakes.has(key))return;
    regeneratingTakes.add(key);regenerate.disabled=true;
    try{
     await request('/api/narration-regenerate?id='+id,{revision:n.revision,take:i,attempt:attempt.id});
     message(`Take ${i} regeneration queued using the saved narration. One ElevenLabs request.`);
    }finally{regeneratingTakes.delete(key);reviewSignature='';await loadNarration();}
   });
   const status=document.createElement('p');status.role='status';
   status.textContent=busy?(t.state==='generating'?'Generating replacement…':'Replacement queued…'):
    attempt.error?'Regeneration error. Previous audio is preserved.':
    attempt.previous_asset?'Regenerated successfully. This is the current audio for this slot.':'Uses the current saved narration.';
   if(n.selected_take===i)status.textContent+=' This slot remains selected.';
   card.append(regenerate,status);
   if(t.state==='ready'&&attempt.error){const error=document.createElement('pre');error.textContent=attempt.error;card.append(error);}
   if(['failed','uncertain','local exception'].includes(t.state)){
    const restore=document.createElement('button');restore.textContent='Restore previous audio';
    restore.onclick=()=>action(async()=>{restore.disabled=true;try{await request('/api/narration-restore?id='+id,{revision:n.revision,take:i});await loadNarration();}catch(e){restore.disabled=false;throw e;}});
    card.append(restore);
   }
  }
  $('takes').append(card);
 }
}
async function selectTake(take){await save();await loadNarration();const n=narration.slides.find(s=>s.number===reviewNumber);await request('/api/narration-select?id='+project.id,{slide:reviewNumber,revision:n.revision,take});await loadNarration();if(typeof loadSelectionReview==='function')await loadSelectionReview(false);message(take?'Selection saved. You can change it at any time.':'Selection cleared. Choose any ready take.');}
$('generate').onclick=()=>action(async()=>{await save();const plan=await request('/api/narration-plan?id='+project.id);if(!plan.generations){message('No new takes needed. Existing attempts are preserved; retry confirmed failures individually.');return;}if(confirm(`${plan.slides} slides\n${plan.generations} narration takes\nGenerate?`)){await request('/api/narration-generate?id='+project.id,{version:plan.version,confirm:true});await loadNarration();message('Narration queued. You can continue using Studio.');}});
$('review').onclick=()=>action(async()=>{await save();await loadNarration();await request('/api/thumbnails?id='+project.id,{});$('review-panel').hidden=false;$('selection-summary').hidden=false;renderReview();});
$('review-slide').onchange=()=>{reviewNumber=Number($('review-slide').value);renderReview();};
for(const [id,step] of [['previous',-1],['next',1]])$(id).onclick=()=>{const slides=project.slides.filter(s=>s.enabled);const index=slides.findIndex(s=>s.number===reviewNumber);reviewNumber=slides[Math.max(0,Math.min(slides.length-1,index+step))].number;renderReview();};
$('change-selection').onclick=()=>action(()=>selectTake(0));
setInterval(async()=>{if(polling||!project||stopped)return;polling=true;try{await loadNarration();if($('thumbnail').hidden&& !$('review-panel').hidden){$('thumbnail').src=`/api/thumbnail?id=${project.id}&slide=${reviewNumber}`;$('thumbnail').onload=()=>{$('thumbnail').hidden=false;$('thumbnail-status').textContent='';};}}catch(e){message('Studio unavailable. Restart LectureForge, then reload to recover saved narration.',true);}finally{polling=false;}},2000);

const narrationKeys=['voice_id','model_id','stability','similarity_boost','style','speed','use_speaker_boost','takes_per_slide'];
function showNarrationSettings(settings){
 for(const key of narrationKeys){const input=$('ns-'+key);if(key==='use_speaker_boost')input.checked=settings[key];else input.value=settings[key];}
 $('generate').textContent='GENERATE '+settings.takes_per_slide+' TAKES';
}
function renderNarrationSettings(){
 const settings={...narrationDefaults,...project.preset.narration};
 const model=$('ns-model_id');model.replaceChildren();
 for(const value of new Set(['eleven_multilingual_v2',settings.model_id])){const o=document.createElement('option');o.value=value;o.textContent=value;model.append(o);}
 showNarrationSettings(settings);$('narration-preset').value='';
 $('test-slide').replaceChildren();const slides=project.slides.filter(s=>s.enabled);
 if(!slides.some(s=>s.number===reviewNumber))reviewNumber=slides[0]?.number||1;
 for(const s of slides){const o=document.createElement('option');o.value=s.number;o.textContent='Slide '+s.number;o.selected=s.number===reviewNumber;$('test-slide').append(o);}
 $('test-status').textContent='';$('test-audio').hidden=true;$('test-audio').pause?.();$('test-audio').removeAttribute?.('src');
 $('narration-test').disabled=!slides.length;
}
function readNarrationSettings(){
 const result={};for(const key of narrationKeys){const input=$('ns-'+key);if(!input.reportValidity())throw Error('Check narration settings.');result[key]=key==='use_speaker_boost'?input.checked:['voice_id','model_id'].includes(key)?input.value:Number(input.value);}
 return result;
}
for(const key of narrationKeys)$('ns-'+key).addEventListener('input',()=>{changed();$('narration-preset').value='';$('generate').textContent='GENERATE '+$('ns-takes_per_slide').value+' TAKES';});
function applyNarrationPreset(natural=false){
 const settings=natural?{...readNarrationSettings(),stability:.47,similarity_boost:.80,style:.38,speed:1,use_speaker_boost:true}:{...narrationDefaults};
 showNarrationSettings(settings);changed();
}
$('narration-defaults').onclick=()=>{applyNarrationPreset();$('narration-preset').value='default';};
$('narration-preset').onchange=()=>action(()=>{if($('narration-preset').value)applyNarrationPreset($('narration-preset').value==='natural');});
$('test-slide').onchange=()=>{reviewNumber=Number($('test-slide').value);renderReview();};
function renderTestTake(){
 const preview=narration?.previews?.filter(p=>p.slide_number===reviewNumber).at(-1),take=preview?.takes?.[0];
 const busy=testSubmitting||['queued','dispatched','generating'].includes(take?.state);
 $('narration-test').disabled=busy||!project.slides.some(s=>s.number===reviewNumber&&s.enabled);
 $('test-status').textContent=testSubmitting?'Submitting test take...':take?'Slide '+reviewNumber+' test take: '+take.state+(take.attempts?.at(-1)?.error?' ? '+take.attempts.at(-1).error:''):'Test preview only. One ElevenLabs request; no HeyGen generation.';
 const audio=$('test-audio');
 if(take?.state==='ready'){
  const src='/api/audio?id='+project.id+'&revision='+preview.id+'&take=1&asset='+encodeURIComponent(take.asset.path);
  if(audio.getAttribute('src')!==src)audio.src=src;audio.hidden=false;
 }else{audio.hidden=true;audio.pause?.();audio.removeAttribute?.('src');}
}
$('test-audio').onplay=()=>document.querySelectorAll('audio').forEach(a=>{if(a!==$('test-audio'))a.pause();});
$('narration-test').onclick=()=>action(async()=>{
 if(testSubmitting)return;
 const id=project.id,slide=reviewNumber,settings=readNarrationSettings();
 const row=[...$('slides').children].find(r=>Number(r.dataset.number)===slide);
 const text=row?.querySelector('.script').value;if(!text?.trim())throw Error('Current slide needs narration text.');
 testSubmitting=true;renderTestTake();
 try{await request('/api/narration-test?id='+id,{slide,version:project.version,settings,text});}
 catch(e){$('test-status').textContent='Test take could not be queued: '+e.message;throw e;}
 finally{testSubmitting=false;if(project?.id===id){await loadNarration();}}
});
