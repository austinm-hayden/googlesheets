/***************************************************************************************
 * PM SERVICE PLANNER – RAW UPLOAD → BRANCH SPLITTER + FORMAT CLONER + ARCHIVER
 * Version: 1.2.0  (2025-10-16)
 * Author Credit:  Created by Austin Monson (with ChatGPT assistance)
 * License: MIT
 *
 * -------------------------------------------------------------------------------------
 * CHANGELOG
 * 1.0.0 – Base release (upload → split → archive)
 * 1.1.0 – Added doPost() upload reliability
 * 1.2.0 – Restored in-sheet upload dialog; fixed hanging; case-insensitive HTML loading
 ***************************************************************************************/

const CONFIG = {
  TEMPLATE_SHEET: 'Pending Service - Springfield',
  BRANCH_COLUMN: 'Branch',
  BRANCHES: [
    { key: 'Springfield', tabName: 'Pending Service - Springfield' },
    { key: 'West Plains', tabName: 'Pending Service - West Plains' },
    { key: 'Villa Ridge', tabName: 'Pending Service - Villa Ridge' },
  ],
  HEADER_ORDER: [
    'Stock #','Description','Type','Overdue',
    'Serial Number','Manufacturer','Model','Due','Notes'
  ],
  DUE_FILTERS: ['Corrected','Service not Needed','Removed'],
  HEADER_ROW: 1
};

/* =============================================================================
   MENU
============================================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PM Planner')
    .addItem('Upload Raw Data (XLSX → Split + Style + Archive)','showUploadDialog')
    .addSeparator()
    .addItem('Restore Archived Sheet…','showRestoreDialog')
    .addToUi();
}

/* =============================================================================
   DIALOG LAUNCHERS (case-insensitive lookup)
============================================================================= */
function _getHtmlFileCaseInsensitive(base) {
  const variants=[base,base.toLowerCase(),base.toUpperCase(),
    base.charAt(0).toUpperCase()+base.slice(1).toLowerCase()];
  for (const n of variants){
    try{return HtmlService.createHtmlOutputFromFile(n);}catch(e){}
  }
  throw new Error('HTML file not found for '+base);
}

/* Template injection so ScriptApp URL resolves correctly inside sheet */
function showUploadDialog(){
  const t=HtmlService.createTemplateFromFile('Dialog');
  t.webAppUrl=ScriptApp.getService().getUrl();
  const html=t.evaluate().setWidth(520).setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html,'Upload Raw PM Data');
}
function showRestoreDialog(){
  const html=_getHtmlFileCaseInsensitive('Restore')
    .setWidth(540).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html,'Restore Archived Branch Sheet');
}

/* =============================================================================
   doPost() – Receives XLSX blob from form
============================================================================= */
function doPost(e){
  try{
    const blob=e.files?.file;
    if(!blob)throw new Error('No file uploaded.');
    const temp=DriveApp.createFile(blob);
    temp.setName('PM_Raw_Upload_'+_stampForFilename()+'.xlsx');
    _logProgress('Uploaded file → '+temp.getName());
    const res=_processUploadedRawFile(temp);
    return HtmlService.createHtmlOutput('Upload completed: '+JSON.stringify(res));
  }catch(err){
    return HtmlService.createHtmlOutput('Error: '+err.message);
  }
}

/* =============================================================================
   CORE PROCESSOR – converts, splits, archives, carries Due/Notes
============================================================================= */
function _processUploadedRawFile(tempFile){
  const ss=SpreadsheetApp.getActive();
  const template=ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if(!template)throw new Error('Template sheet missing.');

  _logProgress('Converting to Google Sheet…');
  const meta={title:tempFile.getName(),mimeType:'application/vnd.google-apps.spreadsheet'};
  const converted=Drive.Files.copy(meta,tempFile.getId());
  const raw=SpreadsheetApp.openById(converted.id);
  const rawSheet=raw.getSheets()[0];
  const rows=_getSheetDataAsObjects(rawSheet);
  Drive.Files.remove(converted.id);
  tempFile.setTrashed(true);

  if(!rows.length)throw new Error('Raw file is empty.');
  if(!(CONFIG.BRANCH_COLUMN in rows[0]))
    throw new Error('Missing branch column "'+CONFIG.BRANCH_COLUMN+'"');

  _logProgress('Splitting rows by branch…');
  const groups={}; CONFIG.BRANCHES.forEach(b=>groups[b.key]=[]);
  rows.forEach(r=>{
    const b=(r[CONFIG.BRANCH_COLUMN]||'').toString().trim();
    if(groups[b])groups[b].push(r);
  });

  CONFIG.BRANCHES.forEach(branch=>{
    _logProgress('Processing '+branch.key+'…');
    const carry=_readCarryoverMap(ss,branch.tabName);
    _archiveSheetIfExists(ss,branch.tabName);
    const sheet=template.copyTo(ss).setName(branch.tabName);
    sheet.showSheet(); _clearDataBelowHeader(sheet,CONFIG.HEADER_ROW);
    const staged=(groups[branch.key]||[]).map(r=>{
      const key=(r['Stock #']||'').toString().trim();
      const prev=carry.get(key)||{};
      const due=prev.Due??r['Due']??'';
      const notes=prev.Notes??r['Notes']??'';
      return CONFIG.HEADER_ORDER.map(h=>{
        if(h==='Due')return due;
        if(h==='Notes')return notes;
        return r[h]??'';
      });
    }).filter(a=>{
      const d=(a[CONFIG.HEADER_ORDER.indexOf('Due')]||'').toString().trim();
      return !CONFIG.DUE_FILTERS.includes(d);
    });
    if(staged.length)
      sheet.getRange(CONFIG.HEADER_ROW+1,1,staged.length,CONFIG.HEADER_ORDER.length).setValues(staged);
    _writeHeaderLabels(sheet,CONFIG.HEADER_ROW,CONFIG.HEADER_ORDER);
    _logProgress(branch.key+' done ('+staged.length+' rows)');
  });

  _logProgress('All branches complete.');
  return{ok:true};
}

/* =============================================================================
   ARCHIVE / RESTORE
============================================================================= */
function listArchives(){
  const ss=SpreadsheetApp.getActive();
  return ss.getSheets().filter(s=>s.getName().startsWith('_Archive_'))
    .map(s=>({name:s.getName(),hidden:s.isSheetHidden()}));
}
function restoreArchive(name){
  if(!name)throw new Error('No archive selected.');
  const ss=SpreadsheetApp.getActive();
  const arch=ss.getSheetByName(name);
  if(!arch)throw new Error('Archive not found.');
  const branchKey=name.replace(/^_Archive_/,'').replace(/_\d{4}-\d{2}-\d{2}_\d{4}$/,'');
  const branch=CONFIG.BRANCHES.find(b=>b.key===branchKey);
  if(!branch)throw new Error('Branch not recognized.');
  _archiveSheetIfExists(ss,branch.tabName);
  const copy=arch.copyTo(ss).setName(branch.tabName);
  copy.showSheet(); ss.setActiveSheet(copy);
  return{ok:true};
}

/* =============================================================================
   HELPERS
============================================================================= */
function _getSheetDataAsObjects(sheet){
  const v=sheet.getDataRange().getValues(); if(!v.length)return[];
  const h=v[0].map(x=>String(x||'').trim()); const out=[];
  for(let i=1;i<v.length;i++){
    const o={}; for(let j=0;j<h.length;j++)o[h[j]]=v[i][j];
    if(Object.values(o).some(x=>x!==''&&x!=null))out.push(o);
  } return out;
}
function _readCarryoverMap(ss,tab){
  const m=new Map(); const sh=ss.getSheetByName(tab); if(!sh)return m;
  _getSheetDataAsObjects(sh).forEach(o=>{
    const k=(o['Stock #']||'').toString().trim(); if(!k)return;
    m.set(k,{Due:o['Due']??'',Notes:o['Notes']??''});
  }); return m;
}
function _archiveSheetIfExists(ss,name){
  const sh=ss.getSheetByName(name); if(!sh)return;
  const key=CONFIG.BRANCHES.find(b=>b.tabName===name)?.key||name;
  const archName=`_Archive_${key}_${_stampForName()}`;
  const arch=sh.copyTo(ss).setName(archName); arch.hideSheet(); ss.deleteSheet(sh);
}
function _clearDataBelowHeader(sh,row){
  const maxR=sh.getMaxRows(),maxC=sh.getMaxColumns();
  if(maxR>row)sh.getRange(row+1,1,maxR-row,maxC).clearContent();
}
function _writeHeaderLabels(sh,row,headers){
  sh.getRange(row,1,1,headers.length).setValues([headers]);
}
function _stampForName(){
  const d=new Date(),y=d.getFullYear(),
  m=String(d.getMonth()+1).padStart(2,'0'),
  dd=String(d.getDate()).padStart(2,'0'),
  h=String(d.getHours()).padStart(2,'0'),
  mi=String(d.getMinutes()).padStart(2,'0');
  return`${y}-${m}-${dd}_${h}${mi}`;
}
function _stampForFilename(){
  const d=new Date();
  return`${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}_${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}${String(d.getSeconds()).padStart(2,'0')}`;
}
function _logProgress(msg){CacheService.getScriptCache().put('pm_progress',msg,120);}
function getLastProgress(){return CacheService.getScriptCache().get('pm_progress')||'';}
