export const normalize = value => String(value ?? '').normalize('NFKC').replace(/[૦-૯]/g,c=>String(c.charCodeAt(0)-0x0ae6)).trim().toLowerCase().replace(/[\s.:_()\-]/g,'');
export const rollKey = value => { const key=normalize(value); return /^\d+$/.test(key) ? key.replace(/^0+(?=\d)/,'') : key; };
export const isRoll = value => /^(roll(no|number|num)?|રોલ(નંબર|નં|ન))$/.test(normalize(value));
export const imageKey = path => rollKey(path.replace(/\\/g,'/').split('/').pop().replace(/\.[^.]+$/,'').replace(/^(?:roll[\s._-]*(?:number|no)?|photo|image|student)[\s._-]*/i,''));
export function matchPhotos(students, images) {
 const byRoll=new Map();for(const image of images){const matches=byRoll.get(image.key)||[];matches.push(image);byRoll.set(image.key,matches);}
 const warnings=[];
 for(const student of students){const matches=byRoll.get(student.key)||[];student.photo=matches.length===1?matches[0]:null;student.photoMatch=matches.length===1?'automatic':null;if(matches.length>1)warnings.push(`Roll ${student.roll}: multiple photos match. Select the correct photo manually.`);}
 const unmatched=students.filter(student=>!student.photo);
 const used=new Set(students.map(student=>student.photo).filter(Boolean));
 const remaining=images.filter(image=>!used.has(image));
 if(unmatched.length===1&&remaining.length===1){unmatched[0].photo=remaining[0];unmatched[0].photoMatch='automatic-remainder';}
 return warnings;
}
// All columns remain in their original order; wide values receive a full row.
export function formRows(fields) {
 const rows=[];let pair=[];const flush=()=>{if(pair.length){rows.push(pair);pair=[];}};
 for(const field of fields){const wide=field.value.length>45||/address|સરનામ|fullname|studentname|આખુંનામ/.test(field.key);if(wide){flush();rows.push([field]);}else{pair.push(field);if(pair.length===2)flush();}}
 flush();return rows;
}
export function inspectZip(buffer, limit=150*1024*1024) {
 const v=new DataView(buffer); let end=-1;
 for(let i=v.byteLength-22;i>=Math.max(0,v.byteLength-65557);i--) if(v.getUint32(i,true)===0x06054b50){end=i;break;}
 if(end<0) throw Error('This is not a readable ZIP archive.');
 const count=v.getUint16(end+10,true); let offset=v.getUint32(end+16,true),total=0;
 if(count>5000) throw Error('The archive contains too many files (maximum 5,000).');
 for(let n=0;n<count;n++) {
  if(offset+46>v.byteLength || v.getUint32(offset,true)!==0x02014b50) throw Error('The ZIP directory is invalid or uses unsupported ZIP64 format.');
  if(v.getUint16(offset+8,true)&1) throw Error('Password-protected archives are not supported.');
  const size=v.getUint32(offset+24,true); total+=size;
  if(size===0xffffffff || total>limit) throw Error('The expanded archive is too large. Please split it into smaller batches.');
  offset+=46+v.getUint16(offset+28,true)+v.getUint16(offset+30,true)+v.getUint16(offset+32,true);
 }
 return total;
}
export function parseWorkbook(buffer, XLSX) {
 const workbook=XLSX.read(buffer,{type:'array',cellNF:true,cellText:true,sheetRows:1100});
 const students=[], warnings=[], seen=new Set(); let title='વિદ્યાર્થી નો બાયોડેટા'; const fields=new Map();
 for(const sheetName of workbook.SheetNames){
  const sheet=workbook.Sheets[sheetName]; if(!sheet['!ref'])continue;
  const range=XLSX.utils.decode_range(sheet['!ref']);
  if(range.e.c>199) throw Error('Maximum 200 columns per worksheet.');
  if(sheet['!fullref'] && XLSX.utils.decode_range(sheet['!fullref']).e.r>1099) throw Error('Maximum 1,000 students per import. Please split the workbook.');
  let hr=-1, rc=-1;
  for(let r=range.s.r;r<=Math.min(range.e.r,29);r++) for(let c=range.s.c;c<=range.e.c;c++) if(isRoll(sheet[XLSX.utils.encode_cell({r,c})]?.v)){hr=r;rc=c;break;}
  if(hr<0){warnings.push(`Skipped worksheet “${sheetName}”: no Roll number / રોલ નંબર header found.`);continue;}
  if(!students.length && hr>0) {const cell=sheet[XLSX.utils.encode_cell({r:hr-1,c:range.s.c})];if(cell?.v)title=String(cell.v).trim();}
  const columns=[], counts=new Map();
  for(let c=range.s.c;c<=range.e.c;c++) {
   let label=String(sheet[XLSX.utils.encode_cell({r:hr,c})]?.v??'').trim();
   if(!label){let used=false;for(let r=hr+1;r<=range.e.r;r++)if(sheet[XLSX.utils.encode_cell({r,c})]?.v!=null){used=true;break;}if(!used)continue;label=`Column ${XLSX.utils.encode_col(c)}`;}
   const n=(counts.get(label)||0)+1;counts.set(label,n);if(n>1)label+=` (${n})`;
   const key=normalize(label);columns.push({c,label,key});fields.set(key,label);
  }
  for(let r=hr+1;r<=range.e.r;r++){
   if(!columns.some(({c})=>String(sheet[XLSX.utils.encode_cell({r,c})]?.v??'').trim()))continue;
   const rawRoll=sheet[XLSX.utils.encode_cell({r,c:rc})];const roll=String(rawRoll?.w??rawRoll?.v??'').trim();
   if(!roll){warnings.push(`${sheetName}, row ${r+1}: skipped because roll number is missing.`);continue;}
   const key=rollKey(roll); if(seen.has(key))throw Error(`Duplicate roll number “${roll}”. Roll numbers must be unique across the uploaded workbook to match photos safely.`);seen.add(key);
   const values=columns.map(col=>{
    const cell=sheet[XLSX.utils.encode_cell({r,c:col.c})]; let value='';
    if(cell?.v!=null){
     if(cell.t==='n' && cell.z && XLSX.SSF.is_date(cell.z)) value=XLSX.SSF.format('dd/mm/yyyy',cell.v);
     else if(cell.t==='n' && Number.isInteger(cell.v) && Math.abs(cell.v)>=1e10){value=String(cell.v);if(!Number.isSafeInteger(cell.v))warnings.push(`Roll ${roll}: ${col.label} is a long numeric ID. Store it as text in Excel to avoid rounded digits.`);}
     else value=String(cell.w??cell.v).trim();
    }
    return {...col,value};
   });
   const name=values.find(f=>/અંગ્રેજી|english/.test(f.label.toLowerCase())&&/નામ|name/.test(f.label.toLowerCase()))?.value || values.find(f=>/^(fullname|studentname|name)$/.test(f.key))?.value || values.find(f=>/ગુજરાતી/.test(f.label))?.value || `Student ${roll}`;
   students.push({id:students.length,roll,key,name,values,sheet:sheetName,photo:null});
   if(students.length>1000)throw Error('Maximum 1,000 students per import. Please split the workbook.');
  }
 }
 if(!students.length) throw Error('No student records found. Include a “Roll number” or “રોલ નંબર” column and at least one student row.');
 return {students,warnings,title,fields:[...fields]};
}
export function filteredStudents(students,query,attribute,status){
 const q=String(query).normalize('NFKC').trim().toLocaleLowerCase();
 return students.filter(s=>(status==='all'||(status==='ready'?!!s.photo:!s.photo)) && (!q||(attribute==='all'?s.values:s.values.filter(f=>f.key===attribute)).some(f=>f.value.normalize('NFKC').toLocaleLowerCase().includes(q))));
}
