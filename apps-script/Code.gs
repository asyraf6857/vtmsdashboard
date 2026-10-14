/** PTP VTMS VOR — Google Apps Script. Bind this script to the master spreadsheet.
 * Enable Advanced Google service: Drive API (v3).
 * Only aggregated data is served to the public dashboard; do not publish vessel-level rows.
 */
const VOR = {sheet:'MONTHLY_SUMMARY',reason:'VISIT_REASON',loa:'LOA_ANALYSIS',log:'UPLOAD_LOG'};
const CATS = [['Mother',10],['Feeder',11],['Container Others',12],['ATB',16],['Malakoff',17],['Bunker',18],['Ship to Ship',19],['Water',20],['Other',21]];
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
  const v=tab.getRange('C10:N44').getValues();
  const cell=(r,m)=>{const x=v[r-10][m];return x===''||x===null?null:Number(x)};
  const months=[];for(let m=0;m<12;m++){const t=cell(23,m);if(t!==null&&t>0)months.push(m);}
  if(!months.length)throw Error('No populated monthly totals found in row 23');
  const summary=[],reasons=[],loa=[];
  months.forEach(m=>{
   const c=cell(13,m),n=cell(22,m),total=cell(23,m);
   if(c===null||n===null||c+n!==total)throw Error('Summary mismatch '+year+' month '+(m+1));
   summary.push([year,m+1,total,c,n,'','','','',new Date()]);
   CATS.forEach(([name,row])=>{const count=cell(row,m);reasons.push([year,m+1,name,count===null?'':count,count===null?'':count/total]);});
   let loaSum=0;
   for(let b=0;b<8;b++){const lc=cell(29+b*2,m)||0,ln=cell(30+b*2,m)||0;loa.push([year,m+1,BINS[b],lc,ln,lc+ln]);loaSum+=lc+ln;}
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
 const ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(VOR.reason);
 const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,4).getValues():[];
 const names={'Mother':'Mother','Feeder':'Feeder','Container Others':'Container Others','ATB':'ATB','Malakoff':'Malakoff','Bunker':'Bunker','Ship to Ship':'Ship to Ship','Water':'Water','Other':'Other'};
 const result={};
 rows.forEach(r=>{const y=String(r[0]),m=Number(r[1])-1,k=names[r[2]];if(!k||m<0||m>11||!y.match(/^20\d\d$/))return;if(!result[y])result[y]={};if(!result[y][k])result[y][k]=[];result[y][k][m]=r[3]===''?0:Number(r[3])||0;});
 Object.values(result).forEach(groups=>{const length=Math.max(0,...Object.values(groups).map(a=>a.length));Object.keys(names).forEach(k=>{groups[k]=Array.from({length},(_,i)=>groups[k]?.[i]||0);});});
 const callback=(typeof e!=='undefined'&&e&&e.parameter&&e.parameter.callback)||''; if(callback && /^vtmsCallback_[a-zA-Z0-9_]+$/.test(callback)) return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');').setMimeType(ContentService.MimeType.JAVASCRIPT); return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
