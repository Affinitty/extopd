import {parseWorkbook,inspectZip,imageKey,filteredStudents,matchPhotos,formRows} from './model.js';
const $=id=>document.getElementById(id);
const state={excel:null,zip:null,students:[],images:[],warnings:[],title:'',selected:null,busy:false};
const make=(tag,cls,text)=>{const e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;};
function message(text,error=false){for(const id of ['message','workspace-message']){$(id).textContent=text;$(id).classList.toggle('error',error);}}
function showView(view){const workspace=view==='workspace';$('import-view').hidden=workspace;$('workspace-view').hidden=!workspace;$('return-workspace').hidden=!state.students.length;window.scrollTo(0,0);requestAnimationFrame(syncStudentScrollbar);}
function setBusy(busy){state.busy=busy;document.querySelectorAll('button,input,select').forEach(e=>e.disabled=busy);$('import').disabled=busy||!state.excel||!state.zip;updateDownloads();}
function choose(file,kind){if(!file)return;const valid=kind==='excel'?/\.xlsx?$/i.test(file.name):/\.zip$/i.test(file.name);const limit=kind==='excel'?10:50;if(!valid||file.size>limit*1024*1024){message(`Choose a ${kind==='excel'?'valid Excel':'ZIP'} file smaller than ${limit} MB.`,true);return;}state[kind]=file;$(kind==='excel'?'excel-name':'zip-name').textContent=file.name;$('import').disabled=!state.excel||!state.zip;message('');}
$('excel').onchange=e=>choose(e.target.files[0],'excel');$('photos').onchange=e=>choose(e.target.files[0],'zip');
for(const [id,kind] of [['excel-drop','excel'],['zip-drop','zip']]){const el=$(id);el.ondragover=e=>{e.preventDefault();el.classList.add('dragover');};el.ondragleave=()=>el.classList.remove('dragover');el.ondrop=e=>{e.preventDefault();el.classList.remove('dragover');if(!state.busy)choose(e.dataTransfer.files[0],kind);};}
async function loadImages(buffer){
 inspectZip(buffer);const zip=await JSZip.loadAsync(buffer);const entries=Object.values(zip.files).filter(e=>!e.dir&&!e.name.startsWith('__MACOSX/')&&/\.(jpe?g|png|webp)$/i.test(e.name));
 if(entries.length>1000)throw Error('Maximum 1,000 photos per import.');
 const images=[],warnings=[];
 for(const entry of entries){let url;try{const bytes=await entry.async('uint8array');if(bytes.length>10*1024*1024){warnings.push(`${entry.name}: skipped (photo exceeds 10 MB).`);continue;}const type=/\.png$/i.test(entry.name)?'image/png':/\.webp$/i.test(entry.name)?'image/webp':'image/jpeg';url=URL.createObjectURL(new Blob([bytes],{type}));const img=new Image();img.src=url;await img.decode();if(img.naturalWidth*img.naturalHeight>25000000)throw Error('image dimensions too large');const canvas=document.createElement('canvas');canvas.width=600;canvas.height=756;const ratio=Math.min(canvas.width/img.naturalWidth,canvas.height/img.naturalHeight);const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height);const dw=img.naturalWidth*ratio,dh=img.naturalHeight*ratio;ctx.drawImage(img,(canvas.width-dw)/2,(canvas.height-dh)/2,dw,dh);images.push({name:entry.name,key:imageKey(entry.name),url:canvas.toDataURL('image/jpeg',.9)});}catch{warnings.push(`${entry.name}: skipped (unreadable or oversized image).`);}finally{if(url)URL.revokeObjectURL(url);}}
 return {images,warnings};
}
$('import').onclick=async()=>{setBusy(true);message('Reading spreadsheet and matching photos…');try{
 const excelBuffer=await state.excel.arrayBuffer();if(/\.xlsx$/i.test(state.excel.name))inspectZip(excelBuffer,40*1024*1024);
 const parsed=parseWorkbook(excelBuffer,XLSX);const photos=await loadImages(await state.zip.arrayBuffer());
 parsed.warnings.push(...matchPhotos(parsed.students,photos.images));
 state.students=parsed.students;state.images=photos.images;state.warnings=[...parsed.warnings,...photos.warnings];state.title=parsed.title;state.selected=state.students[0].id;
 $('search').value='';$('photo-filter').value='all';$('attribute').replaceChildren(new Option('All attributes','all'),...parsed.fields.map(([key,label])=>new Option(label,key)));
 $('student-list').scrollTop=0;$('workspace').hidden=false;$('empty').hidden=true;render();showView('workspace');message(`Imported ${state.students.length} student${state.students.length===1?'':'s'} and ${state.images.length} photo${state.images.length===1?'':'s'}.`);
 }catch(e){message(e.message||'Unable to import these files. Please check the workbook and ZIP.',true);}finally{setBusy(false);}};
function current(){return state.students.find(s=>s.id===state.selected);}
function visible(){return filteredStudents(state.students,$('search').value,$('attribute').value,$('photo-filter').value);}
function updateDownloads(){const s=current();const list=visible();$('download').disabled=state.busy||!s?.photo||!list.includes(s);$('download-all').disabled=state.busy||!state.students.length||state.students.some(s=>!s.photo);$('download-filtered').hidden=!$('search').value&&$('photo-filter').value==='all';$('download-filtered').disabled=state.busy||!list.length||list.some(s=>!s.photo);document.querySelectorAll('.student-download').forEach(b=>{b.disabled=state.busy||!state.students.find(s=>s.id===Number(b.dataset.student))?.photo;});}
function renderList(resetScroll=false){
 const scroll=$('student-list');const previousScroll=resetScroll?0:scroll.scrollTop;
 const list=visible();$('results').textContent=`${list.length} of ${state.students.length} students`;scroll.replaceChildren();
 for(const s of list){
  const card=make('div',`student-card${s.id===state.selected?' active':''}`);const b=make('button','student');b.setAttribute('aria-pressed',s.id===state.selected);
  let avatar;if(s.photo){avatar=make('img','student-avatar');avatar.src=s.photo.url;avatar.alt='';}else avatar=make('span','student-avatar',s.roll);
  const body=make('span');body.append(make('strong','',s.name),make('small','',`Roll ${s.roll}`),make('span',`badge${s.photo?'':' missing'}`,s.photo?(s.photoMatch?.startsWith('automatic')?'Photo auto-matched':'Photo assigned'):'Photo needed'));b.append(avatar,body);
  b.onclick=()=>{state.selected=s.id;renderList();renderPreview();};const download=make('button','student-download','Download PDF');download.dataset.student=s.id;download.setAttribute('aria-label',`Download PDF for roll ${s.roll}, ${s.name}`);download.onclick=()=>runDownload([s]);card.append(b,download);$('student-list').append(card);
 }
 if(!list.length)scroll.append(make('p','muted','No matching students. Try another search.'));scroll.scrollTop=previousScroll;updateDownloads();requestAnimationFrame(syncStudentScrollbar);
}
function render(){const matched=state.students.filter(s=>s.photo).length;const used=new Set(state.students.map(s=>s.photo?.name));const unmatched=state.images.filter(i=>!used.has(i.name)).length;$('count').textContent=state.students.length;$('summary').textContent=`${matched} photos assigned · ${state.students.length-matched} students need a photo · ${unmatched} unassigned photos`;
 const notes=[...state.warnings];if(unmatched)notes.push(`${unmatched} photos are unassigned. Select a student and choose a photo below the preview heading.`);if(matched<state.students.length)notes.push('Assign a photo to each student before downloading their PDF.');$('issues').hidden=!notes.length;$('issues-title').textContent=`Import notes (${notes.length})`;$('issues-list').replaceChildren(...notes.map(t=>make('li','',t)));renderList();renderPreview();}
function createPaper(student){
 const paper=make('article','form-paper');
 const letterhead=make('header','school-header');const banner=make('img','school-banner');banner.src='assets/school-letterhead.jpeg';banner.alt='મહીસાગર માધ્યમિક શાળા સાલાવાડા, તા: લુણાવાડા, જિ: મહીસાગર';
 const codes=make('div','school-codes');for(const line of [['SSC INDEX NO. 85.0114','UDISE: 24310325802','COM. CODE: 427'],['HSC INDEX NO. 34.0038','','MO. NO.: 9737674751']]){const row=make('div','school-code-row');for(const text of line)row.append(make('span','',text));codes.append(row);}letterhead.append(banner,codes);paper.append(letterhead);
 const header=make('div','form-header');const heading=make('div','form-heading');heading.append(make('p','form-student-name',student.name),make('p','form-roll',`રોલ નંબર / Roll No. ${student.roll}`));header.append(heading);
 if(student.photo){const img=make('img','form-photo');img.src=student.photo.url;img.alt=`Photo for roll ${student.roll}`;header.append(img);}else header.append(make('div','photo-placeholder','Student photo\nnot assigned'));paper.append(header);
 const table=make('table','form-table');const cols=make('colgroup');for(const width of ['20%','30%','20%','30%']){const col=make('col');col.style.width=width;cols.append(col);}table.append(cols);
 const tbody=make('tbody');for(const fields of formRows(student.values)){const row=make('tr');for(const field of fields){const th=make('th','',field.label);const td=make('td','',field.value||'—');if(fields.length===1)td.colSpan=3;row.append(th,td);}tbody.append(row);}table.append(tbody);paper.append(table);
 const foot=make('div','form-foot');paper.append(foot);return paper;
}
function renderPreview(){const student=current();if(!student)return;$('preview-heading').textContent=`Roll ${student.roll} · Form`;$('preview').replaceChildren(createPaper(student));const used=new Set(state.students.filter(s=>s.id!==student.id).map(s=>s.photo?.name));$('assign').replaceChildren(new Option('Select a photo from the ZIP',''),...state.images.map((im,i)=>{const o=new Option(im.name+(used.has(im.name)?' (assigned to another student)':''),String(i));o.disabled=used.has(im.name);return o;}));$('assign').value=student.photo?String(state.images.indexOf(student.photo)):'';document.querySelector('.photo-assignment').hidden=!!student.photo&&student.photoMatch?.startsWith('automatic');$('photo-note').textContent=student.photo?(student.photoMatch==='automatic-remainder'?`The only remaining photo, ${student.photo.name}, was selected automatically.`:student.photoMatch==='automatic'?`Automatically matched to roll ${student.roll}: ${student.photo.name}`:'This photo will appear in the downloaded form.'):'No matching filename found. Choose a photo, or name it by roll number in the ZIP.';updateDownloads();}
$('assign').onchange=()=>{const s=current();s.photo=$('assign').value===''?null:state.images[Number($('assign').value)];s.photoMatch=s.photo?'manual':null;render();};
for(const id of ['search','attribute','photo-filter'])$(id).addEventListener(id==='search'?'input':'change',()=>{const list=visible();if(list.length&&!list.some(s=>s.id===state.selected))state.selected=list[0].id;renderList(true);renderPreview();});
const studentList=$('student-list'),scrollbar=$('student-scrollbar'),thumb=scrollbar.querySelector('.scrollbar-thumb');
function syncStudentScrollbar(){
 const max=Math.max(0,studentList.scrollHeight-studentList.clientHeight),track=scrollbar.parentElement.clientHeight;
 scrollbar.hidden=!track||max<2;
 if(scrollbar.hidden)return;
 const thumbHeight=Math.max(32,Math.min(track,track*studentList.clientHeight/studentList.scrollHeight));
 const top=(track-thumbHeight)*studentList.scrollTop/max;
 thumb.style.setProperty('--thumb-height',`${thumbHeight}px`);thumb.style.setProperty('--thumb-y',`${top}px`);
 scrollbar.setAttribute('aria-valuenow',String(Math.round(studentList.scrollTop)));
 scrollbar.setAttribute('aria-valuemax',String(Math.round(max)));
}
studentList.addEventListener('scroll',syncStudentScrollbar,{passive:true});
new ResizeObserver(syncStudentScrollbar).observe($('directory')||document.querySelector('.directory'));
window.addEventListener('resize',syncStudentScrollbar);
document.querySelector('.directory').addEventListener('wheel',event=>{
 if(event.ctrlKey||!matchMedia('(min-width:701px)').matches)return;
 const factor=event.deltaMode===1?24:event.deltaMode===2?studentList.clientHeight:1;
 studentList.scrollTop+=event.deltaY*factor;
 event.preventDefault();
},{passive:false});
let drag=null;
scrollbar.addEventListener('pointerdown',event=>{
 if(event.button!==0)return;
 const track=scrollbar.clientHeight,thumbHeight=thumb.getBoundingClientRect().height;
 if(event.target!==thumb){const y=event.clientY-scrollbar.getBoundingClientRect().top-thumbHeight/2;studentList.scrollTop=y/Math.max(1,track-thumbHeight)*(studentList.scrollHeight-studentList.clientHeight);}
 drag={y:event.clientY,scroll:studentList.scrollTop,track:Math.max(1,track-thumbHeight),max:Math.max(0,studentList.scrollHeight-studentList.clientHeight)};
 scrollbar.setPointerCapture(event.pointerId);event.preventDefault();
});
scrollbar.addEventListener('pointermove',event=>{if(drag){studentList.scrollTop=drag.scroll+(event.clientY-drag.y)/drag.track*drag.max;event.preventDefault();}});
for(const type of ['pointerup','pointercancel'])scrollbar.addEventListener(type,()=>{drag=null;});
scrollbar.addEventListener('keydown',event=>{
 const steps={ArrowDown:64,ArrowUp:-64,PageDown:studentList.clientHeight*.8,PageUp:-studentList.clientHeight*.8};
 if(event.key==='Home')studentList.scrollTop=0;else if(event.key==='End')studentList.scrollTop=studentList.scrollHeight;else if(steps[event.key])studentList.scrollTop+=steps[event.key];else return;
 event.preventDefault();
});
$('change-files').onclick=()=>showView('import');$('return-workspace').onclick=()=>showView('workspace');
$('clear').onclick=()=>{state.students=[];state.images=[];state.warnings=[];state.excel=null;state.zip=null;state.selected=null;$('excel').value='';$('photos').value='';$('excel-name').textContent='Choose or drop an Excel file';$('zip-name').textContent='Choose or drop a photo ZIP';$('workspace').hidden=true;$('empty').hidden=false;$('preview').replaceChildren();$('student-list').replaceChildren();$('assign').replaceChildren();$('import').disabled=true;showView('import');message('Session cleared.');};
function saveBlob(blob,name){const url=URL.createObjectURL(blob);const a=make('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);}
function fileName(s){return `Roll-${s.roll.replace(/[^\p{L}\p{N}_-]/gu,'_')}-${s.name.replace(/[^\p{L}\p{N}_-]+/gu,'_').slice(0,65)}.pdf`;}
async function generatePdf(student){
 if(!student.photo)throw Error(`Assign a photo for roll ${student.roll} first.`);
 await document.fonts.load('500 14px Gujarati');await document.fonts.ready;
 const root=$('export-root');root.replaceChildren();const source=createPaper(student);root.append(source);await Promise.all([...source.querySelectorAll('img')].map(i=>i.decode()));
 // Paginate at table-row boundaries, preserving every field and repeating the student header.
 const rows=[...source.querySelectorAll('tbody tr')];const tableTop=source.querySelector('table').getBoundingClientRect().top-source.getBoundingClientRect().top;const available=1123-tableTop-78;const groups=[];let group=[],height=0;
 for(const row of rows){const rh=row.getBoundingClientRect().height;if(rh>available)throw Error(`A field for roll ${student.roll} is too long for an A4 page. Shorten that cell and import again.`);if(group.length&&height+rh>available){groups.push(group);group=[];height=0;}group.push(row.cloneNode(true));height+=rh;}if(group.length)groups.push(group);
 const pdf=new jspdf.jsPDF({unit:'mm',format:'a4',compress:true});pdf.setProperties({title:`Student biodata - Roll ${student.roll}`,subject:'Student biodata form',creator:'Student Forms'});
 for(let page=0;page<groups.length;page++){const paper=createPaper(student);paper.querySelector('tbody').replaceChildren(...groups[page]);paper.querySelector('.form-foot').append(make('span','',`${page+1} / ${groups.length}`));root.replaceChildren(paper);await Promise.all([...paper.querySelectorAll('img')].map(i=>i.decode()));const canvas=await html2canvas(paper,{scale:2,backgroundColor:'#ffffff',logging:false,windowWidth:1200,windowHeight:1200,scrollX:0,scrollY:0});if(page)pdf.addPage();pdf.addImage(canvas.toDataURL('image/jpeg',.96),'JPEG',0,0,210,297,undefined,'FAST');canvas.width=0;canvas.height=0;}
 root.replaceChildren();return pdf.output('blob');
}
async function runDownload(records,bulk=false){if(state.busy||!records.length)return;setBusy(true);try{const zip=bulk?new JSZip():null;for(let i=0;i<records.length;i++){message(`Creating form ${i+1} of ${records.length}…`);await new Promise(r=>setTimeout(r,0));const blob=await generatePdf(records[i]);if(zip)zip.file(fileName(records[i]),blob);else saveBlob(blob,fileName(records[i]));}if(zip){message('Preparing PDF ZIP…');saveBlob(await zip.generateAsync({type:'blob',compression:'STORE'}),'Student-Forms.zip');}message(`${records.length} form${records.length===1?'':'s'} ready. Download started.`);}catch(e){message(e.message||'Could not create the PDF. Please try again.',true);}finally{$('export-root').replaceChildren();setBusy(false);}}
$('download').onclick=()=>runDownload([current()]);$('download-all').onclick=()=>runDownload([...state.students],true);$('download-filtered').onclick=()=>runDownload(visible(),true);
window.addEventListener('beforeunload',e=>{if(state.students.length){e.preventDefault();e.returnValue='';}});



