'use strict';
let selectionReview=null,selectionProjectId=null,selectionBusy=false,selectionPolling=false;
async function loadSelectionReview(restore=false){
 if(!project||stopped)return;
 const id=project.id,d=await request('/api/selections?id='+id);
 if(project?.id!==id)return;
 selectionReview=d;selectionProjectId=id;
 if(restore){
  $('selection-summary').hidden=false;
  $('review-panel').hidden=!!d.saved&&d.saved.phase!=='narration';
  if(d.saved?.phase==='narration'&&d.saved.slide){reviewNumber=d.saved.slide;$('review-panel').hidden=false;renderReview();}
 }
 renderSelectionReview();
}
function renderSelectionReview(){
 const d=selectionReview;if(!d||!project||selectionProjectId!==project.id)return;
 const s=d.current.snapshot;
 $('selection-summary').hidden=false;
 const selected=s.slides.filter(r=>r.status==='narration selected').length;
 $('avatar-readiness').textContent=selected+' / '+s.slides.length+' required narration takes selected and valid';
 const avatar=s.avatar_preset?.avatar;
 $('avatar-configuration').textContent=avatar?'Avatar: '+avatar.name+' ('+avatar.id+') | Background RGB: '+avatar.background_rgb.join(', ')+' | Crop: '+avatar.crop.width+' x '+avatar.crop.height+' at '+avatar.crop.x+', '+avatar.crop.y+' | Selected narration is the final audio.':'Avatar configuration unavailable; review the project preset.';
 $('review-production').disabled=selectionBusy;
 $('authorization-status').textContent=d.authorization_status==='authorized'?'Avatar generation authorized — waiting for production worker':d.authorization_status==='stale'?'Avatar authorization stale — review and authorize again':d.authorization_status==='consumed'?'Avatar authorization already consumed — resubmission blocked':s.can_authorize?'NARRATION SELECTED — avatar generation not authorized':'Narration selections incomplete — avatar generation not authorized';
 $('summary-warning').textContent=d.review_stale?'The saved summary is stale. Click REVIEW SELECTIONS to review current inputs.':!s.can_authorize?'Resolve every missing, stale, or invalid selection before authorization.':'';
 $('selection-rows').replaceChildren();
 for(const r of s.slides){
  const tr=document.createElement('tr');
  const cells=[`Slide ${r.number} — ${r.title}`,r.selected_take?`Take ${r.selected_take}`:'Missing',r.narration?`${r.narration.duration_seconds.toFixed(3)} s`:'—',r.preparation?.post_speech_silence_seconds?`${r.preparation.post_speech_silence_seconds.toFixed(3)} s prepared pause`:'None',r.status,r.reusable_avatar?'Reusable approved asset · 0 new jobs':r.narration?'1 new job':'Blocked'];
  for(const text of cells){const td=document.createElement('td');td.textContent=text;tr.append(td);}
  const td=document.createElement('td'),change=document.createElement('button');change.className='secondary';change.textContent='CHANGE';change.setAttribute('aria-label',`Change Slide ${r.number} narration`);change.onclick=()=>action(async()=>{
   await save();await request('/api/selection-change?id='+project.id,{slide:r.number});await loadNarration();reviewNumber=r.number;$('selection-summary').hidden=false;$('review-panel').hidden=false;renderReview();$('review-panel').scrollIntoView({block:'start'});
  });td.append(change);tr.append(td);$('selection-rows').append(tr);
 }
 $('selection-cost').textContent=`${s.slides.length} enabled slides · ${s.expected_new_provider_jobs} NEW HeyGen jobs expected · ${s.reusable_avatars} reusable approved avatars`;
 $('selection-excluded').textContent='Excluded slides: '+(s.excluded_slides.join(', ')||'none');
 const already=d.authorization_status==='authorized'||d.authorization_status==='consumed';
 $('authorize-preview').disabled=selectionBusy||!s.can_authorize||d.review_stale||already;
 $('authorization-confirmation').hidden=d.saved?.phase!=='confirmation';
 const reviewed=d.review?.snapshot;
 if(reviewed){
  $('confirmation-cost').textContent=`${reviewed.slides.length} slides · ${reviewed.expected_new_provider_jobs} NEW HeyGen jobs · ${reviewed.reusable_avatars} reusable approved avatars`;
  $('confirmation-mapping').replaceChildren();for(const r of reviewed.slides){const li=document.createElement('li');li.textContent=`Slide ${r.number} → Take ${r.selected_take}${r.preparation?.post_speech_silence_seconds?` · ${r.preparation.post_speech_silence_seconds.toFixed(3)} s prepared pause`:''}`;$('confirmation-mapping').append(li);}
  $('confirmation-excluded').textContent='Excluded slides: '+(reviewed.excluded_slides.join(', ')||'none');
 }
 $('confirm-avatars').disabled=selectionBusy||d.review_stale||!s.can_authorize||already||!reviewed;
}
async function reviewSelections(){
 await save();await loadSelectionReview();
 await request('/api/selection-review?id='+project.id,{expected:selectionReview.current.binding_sha256,phase:'review'});
 await loadSelectionReview();$('review-panel').hidden=true;$('selection-summary').hidden=false;$('selection-summary').scrollIntoView({block:'start'});
}
$('review-selections').onclick=$('review-production').onclick=()=>action(reviewSelections);
$('authorize-preview').onclick=()=>action(async()=>{
 if(selectionBusy)return;selectionBusy=true;renderSelectionReview();
 try{await save();await loadSelectionReview();await request('/api/selection-review?id='+project.id,{expected:selectionReview.current.binding_sha256,phase:'confirmation'});await loadSelectionReview();$('authorization-confirmation').scrollIntoView({block:'start'});}finally{selectionBusy=false;renderSelectionReview();}
});
$('confirm-avatars').onclick=()=>action(async()=>{
 if(selectionBusy)return;selectionBusy=true;renderSelectionReview();
 const id=project.id,reviewId=selectionReview.review.id;
 try{await request('/api/avatar-authorize?id='+id,{confirm:true,review_id:reviewId,expected:reviewId});await loadSelectionReview();if(typeof loadAvatarStatus==='function')await loadAvatarStatus();message('Avatar generation authorized — waiting for production worker. Background production will now begin.');}finally{selectionBusy=false;renderSelectionReview();}
});
$('cancel-confirmation').onclick=()=>action(reviewSelections);
setInterval(async()=>{
 if(!project||stopped||$('selection-summary').hidden||selectionPolling||selectionBusy)return;
 selectionPolling=true;try{await loadSelectionReview();}catch(e){$('summary-warning').textContent=e.message;$('authorize-preview').disabled=true;$('confirm-avatars').disabled=true;}finally{selectionPolling=false;}
},5000);
