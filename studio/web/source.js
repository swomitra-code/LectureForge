'use strict';
let sourceChoice=null,sourceComparison=null,sourceBusy=false;
function clearSourceComparison(){sourceComparison=null;$('source-comparison').hidden=true;}
function sourceButtons(busy){
 sourceBusy=busy;
 for(const id of ['reload-source','choose-source','source-confirm','source-cancel','choose-import-source','create-lecture'])$(id).disabled=busy;
}
async function chooseSourceFile(){
 const started=await request('/api/source-choose',{});
 for(let n=0;n<1200;n++){
  const result=await request('/api/source-choice?choice='+started.choice);
  if(result.state==='selected')return {...result,choice:started.choice};
  if(result.state==='cancelled')return null;
  if(result.state==='error')throw Error(result.error||'Could not open the PowerPoint picker.');
  await new Promise(resolve=>setTimeout(resolve,500));
 }
 throw Error('Source selection timed out. Choose the file again.');
}
async function chooseImportSource(){
 if(sourceBusy)return;
 sourceButtons(true);
 try{const selected=await chooseSourceFile();if(selected){sourceChoice=selected;$('pptx').value=selected.path;}}
 finally{sourceButtons(false);}
}
async function createSourceLecture(){
 if(sourceBusy)return;
 if(!sourceChoice)throw Error('Choose a source PowerPoint first.');
 sourceButtons(true);message('Copying and inspecting PowerPoint...');
 try{
  const p=await request('/api/new-source',{name:$('name').value,preset:$('preset').value,choice:sourceChoice.choice});
  renderProject(p);message('Lecture created. The external source path is saved. Import or edit its slide scripts.');
 }finally{sourceButtons(false);}
}
function showSourceComparison(c,id){
 sourceComparison={...c,projectId:id};$('source-comparison').hidden=false;
 $('source-comparison-path').textContent=c.path;
 $('source-old-sha').textContent=c.old.sha256;$('source-new-sha').textContent=c.new.sha256;
 $('source-counts').textContent=`${c.old.slide_count} before / ${c.new.slide_count} now`;
 $('source-order').textContent=c.ordered_ids_match?'Match':'Different';
 $('source-dimensions').textContent=`${c.old.width_emu} x ${c.old.height_emu} before / ${c.new.width_emu} x ${c.new.height_emu} now (${c.dimensions_match?'match':'different'})`;
 $('source-confirm').hidden=!c.compatible;
 $('source-confirmation').textContent=c.compatible
  ?`${c.changed?'Source PowerPoint changed.':'Source PowerPoint is unchanged; its location can be saved.'} ${c.old.slide_count} slides before, ${c.new.slide_count} slides now. Existing narration and avatar assets will be preserved. Reload source?`
  :`${c.problems.join(' ')} Reload stopped. Slide production state will not be remapped.`;
}
async function compareSource(chooseDifferent=false){
 if(sourceBusy||!project)return;
 sourceButtons(true);$('source-comparison').hidden=true;sourceComparison=null;
 const id=project.id;
 try{
  let path='';
  if(chooseDifferent||!project.source.external_path){const selected=await chooseSourceFile();if(!selected)return;path=selected.path;}
  const c=await request('/api/source-compare?id='+id,{path});
  if(project?.id!==id)return;
  if(c.needs_selection)throw Error('Choose Different Source PowerPoint to locate the source.');
  if(!c.changed&&c.path===project.source.external_path){message('Source PowerPoint is already up to date.');return;}
  showSourceComparison(c,id);
 }finally{sourceButtons(false);}
}
async function confirmSourceReload(){
 if(sourceBusy||!sourceComparison?.compatible||project?.id!==sourceComparison.projectId)return;
 if(dirty)throw Error('Project edits changed after comparison. Save them and compare the source again.');
 const c=sourceComparison;sourceButtons(true);
 message('Backing up the project and reloading the source. PowerPoint previews may take a few minutes...');
 try{
  const result=await request('/api/source-reload?id='+c.projectId,{path:c.path,expected:c.expected,confirm:true});
  sourceComparison=null;$('source-comparison').hidden=true;
  if(project?.id===c.projectId){
   renderProject(result.project);
   if(typeof checkAssembly==='function')await checkAssembly();
   if(typeof loadAssembly==='function')await loadAssembly();
   if(!$('review-panel').hidden)$('thumbnail').src=`/api/thumbnail?id=${project.id}&slide=${reviewNumber}&source=${project.source.sha256}`;
  }
  message(result.message);
 }finally{sourceButtons(false);}
}
$('choose-import-source').onclick=()=>action(chooseImportSource);
$('reload-source').onclick=()=>action(()=>compareSource(false));
$('choose-source').onclick=()=>action(()=>compareSource(true));
$('source-confirm').onclick=()=>action(confirmSourceReload);
$('source-cancel').onclick=()=>{sourceComparison=null;$('source-comparison').hidden=true;};
$('source-folder').onclick=()=>action(async()=>{const result=await request('/api/source-folder?id='+project.id,{});message(result.message);});
