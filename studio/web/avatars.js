'use strict';
let avatarBusy=false,avatarState=null,avatarShown=null;
async function loadAvatarStatus(){
 if(!project||stopped||avatarBusy)return;avatarBusy=true;const id=project.id;
 try{
  const d=await request('/api/avatar-status?id='+id);if(project?.id!==id)return;avatarState=d;
  $('avatar-dashboard').hidden=!d.authorized;if(!d.authorized)return;if(avatarShown!==d.authorization){avatarShown=d.authorization;$('selection-summary').hidden=true;$('review-panel').hidden=true;$('avatar-dashboard').scrollIntoView({block:'start'});}
  $('avatar-summary').textContent=`${d.authorized} authorized · ${d.ready} Avatar Ready · ${d.provider_processing} provider processing · ${d.queued} queued · ${d.exceptions} exceptions${d.paused?' · Queue paused':''}`;
  $('pause-avatars').textContent=d.paused?'Resume Queue':'Pause Queue';
  const open=new Set([...$('avatar-jobs').querySelectorAll('details[open]')].map(x=>x.dataset.job));
  $('avatar-jobs').replaceChildren();
  for(const j of d.jobs){
   const row=document.createElement('article'),title=document.createElement('h3');title.textContent=`Slide ${j.slide} — ${j.stage}${j.stage==='Avatar Ready'&&j.validation?.prepared_pause_seconds?` — ${j.validation.prepared_pause_seconds.toFixed(3)} s pause verified`:''}`;row.append(title);
   const details=document.createElement('details');details.dataset.job=j.id;details.open=open.has(j.id);const label=document.createElement('summary');label.textContent='View Details';const pre=document.createElement('pre');pre.style.whiteSpace='pre-wrap';pre.textContent=JSON.stringify({take:j.take,authorization:j.authorization,job_id:j.video_id,narration_sha256:j.row.narration.sha256,output:j.output_path,placement:j.placement,validation:j.validation,error:j.error,history:j.history},null,2);details.append(label,pre);row.append(details);
   if(j.stage==='Exception'){
    const error=document.createElement('p');error.textContent=j.error;row.append(error);
    let actions=[];
    if(j.video_id&&j.provider_status!=='failed')actions.push(['resume','Resume known HeyGen job']);
    if(j.video_id&&j.provider_status==='completed')actions.push(['download','Retry avatar download']);
    if(j.raw_path)actions.push(['convert','Retry white-avatar conversion']);
    if(j.output_path)actions.push(['validate','Retry local validation']);
    if(j.exception_kind==='submission uncertain'&&!j.video_id)actions.push(['reconcile','Reconcile existing job ID']);
    for(const [kind,text] of actions){const b=document.createElement('button');b.className='secondary';b.textContent=text;b.onclick=()=>action(async()=>{const body={job:j.id,action:kind};if(kind==='reconcile'){body.video_id=prompt('Existing HeyGen job ID from the provider account. This does not create a replacement.');if(!body.video_id)return;}await request('/api/avatar-retry?id='+project.id,body);await loadAvatarStatus();});row.append(b);}
   }
   $('avatar-jobs').append(row);
  }
 }catch(e){message(e.message,true);}finally{avatarBusy=false;}
}
$('pause-avatars').onclick=()=>action(async()=>{if(!avatarState)return;await request('/api/avatar-pause?id='+project.id,{paused:!avatarState.paused});await loadAvatarStatus();});
setInterval(loadAvatarStatus,2000);
