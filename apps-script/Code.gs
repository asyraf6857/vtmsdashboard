/** PTP VTMS VOR — Google Apps Script. Bind this script to the master spreadsheet.
 * Enable Advanced Google service: Drive API (v3).
 * Only aggregated data is served to the public dashboard; do not publish vessel-level rows.
 */
const VOR = {sheet:'MONTHLY_SUMMARY',reason:'VISIT_REASON',loa:'LOA_ANALYSIS',log:'UPLOAD_LOG'};
const CATS = [['Mother',8],['Feeder',9],['Container Others',10],['ATB',14],['Malakoff',15],['Bunker',16],['Ship to Ship',17],['Water',18],['Other',19]];
const BINS = ['≤100 m','100–150 m','150–200 m','200–250 m','250–300 m','300–350 m','350–400 m','>400 m'];
function onOpen(){SpreadsheetApp.getUi().createMenu('VTMS VOR').addItem('Import Excel from Drive','importPrompt').addToUi();}
function importPrompt(){const ui=SpreadsheetApp.getUi();const p=ui.prompt('Import Excel','Paste the Google Drive FILE ID of an uploaded .xlsx VOR file (not the folder ID).',ui.ButtonSet.OK_CANCEL);if(p.getSelectedButton()!==ui.Button.OK)return;try{const msg=importExcel(p.getResponseText().trim());ui.alert(msg);}catch(e){ui.alert('Import failed: '+e.message);}}
function importExcel(fileId){
 const lock=LockService.getDocumentLock();lock.waitLock(30000);
 let tempId;
 try {
  const file=DriveApp.getFileById(fileId);
  if(!/\.xlsx?$/i.test(file.getName()))throw Error('Expected an Excel .xlsx file');
  const converted=Drive.Files.create({name:'VTMS_TEMP_'+Date.now(),mimeType:'application/vnd.google-apps.spreadsheet'},file.getBlob(),{fields:'id'});
  tempId=converted.id;
  const source=SpreadsheetApp.openById(tempId);
  const yearMatch=file.getName().match(/20\d{2}/);
  if(!yearMatch)throw Error('Excel filename must include a year, e.g. VOR 2026.xlsx');
  const year=Number(yearMatch[0]);
  const tab=source.getSheetByName(String(year));
  if(!tab)throw Error('Year worksheet '+year+' not found in Excel');
  const v=tab.getRange('C8:N44').getValues();
  const cell=(r,m)=>{const x=v[r-8][m];return x===''||x===null?null:Number(x)};
  const months=[];for(let m=0;m<12;m++){const t=cell(21,m);if(t!==null&&t>0)months.push(m);}
  if(!months.length)throw Error('No populated monthly totals found in row 23');
  const summary=[],reasons=[],loa=[];
  months.forEach(m=>{
   const c=cell(11,m),n=cell(20,m),total=cell(21,m);
   if(c===null||n===null||total===null||!Number.isFinite(c)||!Number.isFinite(n)||!Number.isFinite(total)||c+n!==total)throw Error('Summary mismatch '+year+' month '+(m+1)+': Container (row 11)='+c+', Non-Container (row 20)='+n+', Total (row 21)='+total+'. Expected Container + Non-Container = Total. Check source layout. No data was replaced.');
   summary.push([year,m+1,total,c,n,'','','','',new Date()]);
   CATS.forEach(([name,row])=>{const count=cell(row,m);reasons.push([year,m+1,name,count===null?'':count,count===null?'':count/total]);});
   let loaSum=0;
   for(let b=0;b<8;b++){const lc=cell(27+b*2,m)||0,ln=cell(28+b*2,m)||0;loa.push([year,m+1,BINS[b],lc,ln,lc+ln]);loaSum+=lc+ln;}
   if(loaSum!==total)throw Error('LOA mismatch '+year+' month '+(m+1));
  });
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  function replaceYear(name,newRows){
   const sh=ss.getSheetByName(name),last=sh.getLastRow(),cols=sh.getLastColumn();
   const old=last>1?sh.getRange(2,1,last-1,cols).getValues():[];
   const keep=old.filter(row=>Number(row[0])!==year && row[0]!=='' );
   const all=keep.concat(newRows);
   if(last>1)sh.getRange(2,1,last-1,cols).clearContent();
   if(all.length)sh.getRange(2,1,all.length,cols).setValues(all);
  }
  replaceYear(VOR.sheet,summary);replaceYear(VOR.reason,reasons);replaceYear(VOR.loa,loa);
  ss.getSheetByName(VOR.log).appendRow([Utilities.getUuid(),file.getName(),new Date(),summary.length,0,0,Session.getActiveUser().getEmail(),'SUCCESS','Replaced '+year+' monthly aggregates']);
  return 'Imported '+months.length+' months for '+year+'. Dashboard API data refreshed.';
 }catch(e){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOR.log).appendRow([Utilities.getUuid(),fileId,new Date(),0,0,0,'','FAILED',String(e.message)]);throw e;}
 finally{if(tempId)try{DriveApp.getFileById(tempId).setTrashed(true);}catch(e){}lock.releaseLock();}
}
function doGet(e){
 if(e && e.parameter && e.parameter.page === 'movement-upload'){try{assertUploadAccess_();return HtmlService.createHtmlOutputFromFile('UploadMovement').setTitle('VTMS Movement Planning Upload');}catch(err){return HtmlService.createHtmlOutput('<h2>Access denied</h2><p>Authorized Google account required.</p>');}}
 if(e && e.parameter && e.parameter.page === 'upload'){
   try { assertUploadAccess_(); return HtmlService.createHtmlOutputFromFile('Upload').setTitle('VTMS VOR — Secure Upload'); }
   catch(err){ return HtmlService.createHtmlOutput('<h2>Access denied</h2><p>Upload requires an authorized Google account. Contact the VTMS administrator.</p>').setTitle('Access denied'); }
 }

 const ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(VOR.reason);
 const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,4).getValues():[];
 const names={'Mother':'Mother','Cont. Mother':'Mother','Feeder':'Feeder','Cont. Feeder':'Feeder','Container Others':'Container Others','Cont. Others':'Container Others','ATB':'ATB','Malakoff':'Malakoff','Bunker':'Bunker','Ship to Ship':'Ship to Ship','Water':'Water','Other':'Other'};
 const result={};
 rows.forEach(r=>{const y=String(r[0]),m=Number(r[1])-1,k=names[r[2]];if(!k||m<0||m>11||!y.match(/^20\d\d$/))return;if(!result[y])result[y]={};if(!result[y][k])result[y][k]=[];result[y][k][m]=r[3]===''?0:Number(r[3])||0;});
 Object.values(result).forEach(groups=>{const length=Math.max(0,...Object.values(groups).map(a=>a.length));[...new Set(Object.values(names))].forEach(k=>{groups[k]=Array.from({length},(_,i)=>groups[k]?.[i]||0);});});
 // Public dashboard receives only monthly aggregated LOA counts, never vessel-level records.
 const loaSheet=ss.getSheetByName(VOR.loa);
 const loaRows=loaSheet && loaSheet.getLastRow()>1?loaSheet.getRange(2,1,loaSheet.getLastRow()-1,6).getValues():[];
 result.__loa=loaRows.filter(r=>/^20\d{2}$/.test(String(r[0]))&&Number(r[1])>=1&&Number(r[1])<=12).map(r=>[Number(r[0]),Number(r[1]),String(r[2]),Number(r[3])||0,Number(r[4])||0]);
 const movementSheet=ss.getSheetByName('MOVEMENT_SUMMARY');
 result.__movement=movementSheet&&movementSheet.getLastRow()>1?movementSheet.getRange(2,1,movementSheet.getLastRow()-1,10).getValues().filter(r=>Number(r[0])>=2020&&Number(r[1])>=1&&Number(r[1])<=12).map(r=>r.map(v=>Number(v)||0)):[];
 const dailySheet=ss.getSheetByName('MOVEMENT_DAILY');
 result.__movementDaily=dailySheet&&dailySheet.getLastRow()>1?dailySheet.getRange(2,1,dailySheet.getLastRow()-1,11).getDisplayValues().filter(r=>/^20\d{2}-\d{2}-\d{2}$/.test(String(r[0]))).map(r=>[String(r[0]),...r.slice(1).map(v=>Number(v)||0)]):[];
 const callback=(typeof e!=='undefined'&&e&&e.parameter&&e.parameter.callback)||''; if(callback && /^vtmsCallback_[a-zA-Z0-9_]+$/.test(callback)) return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');').setMimeType(ContentService.MimeType.JAVASCRIPT); return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/** Upload permission: script property VTMS_UPLOAD_ALLOWLIST is comma-separated emails.
 * Fail closed when account identity cannot be verified. Never put credentials in GitHub Pages.
 * Deploy web app "Execute as: User accessing the web app"; restrict access to your organization where available.
 */
function assertUploadAccess_(){
 const email=(Session.getActiveUser().getEmail()||'').trim().toLowerCase();
 const allow=(PropertiesService.getScriptProperties().getProperty('VTMS_UPLOAD_ALLOWLIST')||'').split(',').map(s=>s.trim().toLowerCase()).filter(Boolean);
 if(!email || !allow.includes(email))throw Error('Not authorized for VTMS VOR upload');
 return email;
}
function importUploadedForm(form){
 const user=assertUploadAccess_();
 if(!form || !form.vorFile)throw Error('Please select a VOR Excel file');
 const blob=form.vorFile, name=String(form.originalFileName||blob.getName()||'').trim();
 if(!name.toLowerCase().endsWith('.xlsx'))throw Error('Only .xlsx Excel files are accepted (received: '+name+')');
 if(blob.getBytes().length>10*1024*1024)throw Error('Maximum upload size is 10 MB');
 const file=DriveApp.createFile(blob).setName(name);
 try { return importExcel(file.getId())+' Uploaded by '+user; }
 finally { try { file.setTrashed(true); } catch(e){} }
}

/** Import movement planning workbook into MOVEMENT_SUMMARY.
 * Counts rows with a nonblank Assign Pilot; month determined by Move SRT.
 * Replace only years represented in valid input data.
 */
function importMovementUploadedForm(form, previewOnly){
 const user=assertUploadAccess_();
 if(!form||!form.movementFile)throw Error('Select Movement Planning .xlsx');
 const name=String(form.originalFileName||form.movementFile.getName()||'');
 if(!/\.xlsx$/i.test(name))throw Error('Only .xlsx files are supported');
 if(form.movementFile.getBytes().length>12*1024*1024)throw Error('Maximum upload size: 12 MB');
 const lock=LockService.getDocumentLock();lock.waitLock(30000);
 let uploaded=null,converted=null;
 try{
  uploaded=DriveApp.createFile(form.movementFile).setName(name);
  converted=Drive.Files.create({name:'VTMS_MOVEMENT_TEMP_'+Date.now(),mimeType:'application/vnd.google-apps.spreadsheet'},uploaded.getBlob(),{fields:'id'}).id;
  const source=SpreadsheetApp.openById(converted);
  const candidate=source.getSheets().find(sh=>{const a=sh.getRange(1,1,1,Math.min(sh.getLastColumn(),30)).getDisplayValues()[0].map(x=>String(x).trim().toLowerCase());return a.includes('assign pilot')&&a.includes('move srt')&&a.includes('move type')});
  if(!candidate)throw Error('Cannot find a movement worksheet with Assign Pilot, Move SRT and Move Type');
  const cols=candidate.getRange(1,1,1,candidate.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim().toLowerCase());
  const pilot=cols.indexOf('assign pilot'),date=cols.indexOf('move srt'),type=cols.indexOf('move type'),reason=cols.indexOf('visit reason');
  if(reason<0)throw Error('Visit Reason column missing');
  const incoming=new Map(),years=new Set(),tz='Asia/Kuala_Lumpur';
  let assigned=0,skipped=0;
  function dateKey_(value){
   if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,tz,'yyyy-MM-dd');
   if(typeof value==='number'&&Number.isFinite(value))return Utilities.formatDate(new Date(Date.UTC(1899,11,30)+Math.floor(value)*86400000),'UTC','yyyy-MM-dd');
   const text=String(value||'').trim(),iso=text.match(/^(20\d{2})-(\d{1,2})-(\d{1,2})(?:\s|T|$)/);
   if(iso){const y=+iso[1],m=+iso[2],d=+iso[3],test=new Date(Date.UTC(y,m-1,d));return test.getUTCFullYear()===y&&test.getUTCMonth()===m-1&&test.getUTCDate()===d?y+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0'):null}
   const dm=text.match(/^(\d{1,2})\/(\d{1,2})\/(20\d{2})(?:\s|$)/);
   if(dm){const y=+dm[3],m=+dm[2],d=+dm[1],test=new Date(Date.UTC(y,m-1,d));return test.getUTCFullYear()===y&&test.getUTCMonth()===m-1&&test.getUTCDate()===d?y+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0'):null}
   return null;
  }
  const last=candidate.getLastRow(),width=Math.max(pilot,date,type,reason)+1;
  for(let start=2;start<=last;start+=3000){
   candidate.getRange(start,1,Math.min(3000,last-start+1),width).getValues().forEach(row=>{
    if(!String(row[pilot]||'').trim()){skipped++;return}
    const day=dateKey_(row[date]);if(!day){skipped++;return}
    const year=Number(day.slice(0,4)),month=Number(day.slice(5,7));
    if(year<2020||year>2100){skipped++;return}
    const pos={ARRIVAL:0,DEPARTURE:2,MOVE:4,SHIFT:6}[String(row[type]||'').trim().toUpperCase()];
    if(pos===undefined){skipped++;return}
    const group=String(row[reason]||'').trim().toLowerCase()==='container'?0:1;
    if(!incoming.has(day))incoming.set(day,[day,year,month,0,0,0,0,0,0,0,0]);
    incoming.get(day)[3+pos+group]++;
    years.add(year);assigned++;
   });
  }
  if(!assigned)throw Error('No valid assigned-pilot movements found');
  if(years.size!==1)throw Error('File contains multiple movement years; upload one year at a time');
  const ss=SpreadsheetApp.getActiveSpreadsheet(),daily=ss.getSheetByName('MOVEMENT_DAILY'),monthly=ss.getSheetByName('MOVEMENT_SUMMARY');
  if(!daily||!monthly)throw Error('MOVEMENT_DAILY or MOVEMENT_SUMMARY tab missing');
  const oldDaily=daily.getLastRow()>1?daily.getRange(2,1,daily.getLastRow()-1,11).getValues():[];
  const merged=new Map();
  oldDaily.forEach(row=>{
   if(row[0]===''||row[0]===null)return;
   const key=dateKey_(row[0]);if(!key)throw Error('Invalid existing daily date: '+row[0]+'. No data changed.');
   if(merged.has(key))throw Error('Duplicate existing daily date '+key+'. Resolve before uploading.');
   merged.set(key,[key,Number(key.slice(0,4)),Number(key.slice(5,7)),...row.slice(3,11).map(v=>Number(v)||0)]);
  });
  // Replace only dates included in the uploaded file; preserve all other dates.
  incoming.forEach((row,key)=>merged.set(key,row));
  if(previewOnly){
   const details=[...incoming.entries()].sort((a,b)=>a[0].localeCompare(b[0])).map(([day,row])=>{
    const before=oldDaily.find(r=>dateKey_(r[0])===day);
    const previous=before?before.slice(3,11).map(v=>Number(v)||0):[0,0,0,0,0,0,0,0];
    const current=row.slice(3,11),oldTotal=previous.reduce((a,b)=>a+b,0),newTotal=current.reduce((a,b)=>a+b,0);
    return {date:day,previous:oldTotal,uploaded:newTotal,difference:newTotal-oldTotal,categories:current,previousCategories:previous};
   });
   return {year:[...years][0],assigned,skipped,dates:details,changed:details.filter(d=>d.difference!==0||d.categories.some((v,i)=>v!==d.previousCategories[i])).length};
  }
  const ordered=[...merged.values()].sort((a,b)=>a[0].localeCompare(b[0]));
  const aggregates=new Map();
  ordered.forEach(row=>{
   const key=row[1]+'-'+row[2];
   if(!aggregates.has(key))aggregates.set(key,[row[1],row[2],0,0,0,0,0,0,0,0]);
   for(let j=0;j<8;j++)aggregates.get(key)[2+j]+=row[3+j];
  });
  const monthRows=[...aggregates.values()].sort((a,b)=>a[0]-b[0]||a[1]-b[1]);
  if(daily.getMaxRows()<ordered.length+1)daily.insertRowsAfter(daily.getMaxRows(),ordered.length+1-daily.getMaxRows());
  if(monthly.getMaxRows()<monthRows.length+1)monthly.insertRowsAfter(monthly.getMaxRows(),monthRows.length+1-monthly.getMaxRows());
  const previousDaily=daily.getLastRow(),previousMonthly=monthly.getLastRow();
  // Write validated results; daily dates are stored as real date cells.
  const dailyValues=ordered.map(r=>[new Date(Number(r[1]),Number(r[2])-1,Number(r[0].slice(8,10)),12,0,0),...r.slice(1)]);
  daily.getRange(2,1,dailyValues.length,11).setValues(dailyValues);
  daily.getRange(2,1,dailyValues.length,1).setNumberFormat('yyyy-mm-dd');
  daily.getRange(2,12,dailyValues.length,1).setFormulas(dailyValues.map((r,i)=>['=SUM(D'+(i+2)+':K'+(i+2)+')']));
  if(previousDaily>dailyValues.length+1)daily.getRange(dailyValues.length+2,1,previousDaily-dailyValues.length-1,12).clearContent();
  monthly.getRange(2,1,monthRows.length,10).setValues(monthRows);
  monthly.getRange(2,11,monthRows.length,1).setFormulas(monthRows.map((r,i)=>['=SUM(C'+(i+2)+':J'+(i+2)+')']));
  if(previousMonthly>monthRows.length+1)monthly.getRange(monthRows.length+2,1,previousMonthly-monthRows.length-1,11).clearContent();
  return 'Updated '+incoming.size+' dates ('+assigned+' movements). Preserved '+(ordered.length-incoming.size)+' other dates. Recalculated '+monthRows.length+' monthly summaries. Skipped '+skipped+' rows. Uploaded by '+user+'.';
 }finally{
  if(converted)try{DriveApp.getFileById(converted).setTrashed(true)}catch(e){}
  if(uploaded)try{uploaded.setTrashed(true)}catch(e){}
  lock.releaseLock();
 }
}

/** Two-step upload: preview is read-only; confirmation requires a short-lived token. */
function previewMovementUploadedForm(form){
 const email=assertUploadAccess_();
 if(!form||!form.movementFile)throw Error('Select an Excel file');
 const result=importMovementUploadedForm(form,true);
 const blob=form.movementFile;
 const token=Utilities.getUuid().replace(/-/g,'');
 const saved=DriveApp.createFile(blob).setName('VTMS_PENDING_MOVEMENT_'+token+'.xlsx');
 try{
  PropertiesService.getScriptProperties().setProperty('VTMS_PREVIEW_'+token,JSON.stringify({email,fileId:saved.getId(),name:String(form.originalFileName||blob.getName()),expires:Date.now()+20*60*1000}));
 }catch(e){saved.setTrashed(true);throw e}
 result.token=token;
 return result;
}
function confirmMovementPreview(token){
 const email=assertUploadAccess_();
 if(!/^[a-f0-9]{32}$/.test(String(token||'')))throw Error('Invalid preview token');
 const props=PropertiesService.getScriptProperties(),key='VTMS_PREVIEW_'+token,raw=props.getProperty(key);
 if(!raw)throw Error('Preview expired or already confirmed. Upload again.');
 const record=JSON.parse(raw);
 if(record.email!==email||Date.now()>record.expires)throw Error('Preview expired or unauthorized. Upload again.');
 // Delete token before writing to prevent double confirmation.
 props.deleteProperty(key);
 const file=DriveApp.getFileById(record.fileId);
 try{return importMovementUploadedForm({movementFile:file.getBlob().setName(record.name),originalFileName:record.name},false)}
 finally{try{file.setTrashed(true)}catch(e){}}
}
