/**
 * V16: One vehicle = one complete row in LOADING_REPORT. No report formulas.
 * Deploy this new version, then run setupLoadingDirectReport once.
 * Existing Code.gs and other stage reports stay in place.
 */
const LOADING_LOG_WEB_TABLES = {
  report: { sheet: 'LOADING_REPORT', fields: ['DATE','LOADING INCHARGE','CUSTOMER','LOADING START','LOADING END','TOTAL_HRS','VEHICLE NO','VEHICLE FEET','REMARKS','SHIFT','HELPER COUNT','ENTRY_ID','SHIFT_ID','CHECK','UPDATED_AT'] },
  header: { sheet: 'LOADING_HEADER', fields: ['LOADING_ID','DATE','SHIFT','LOADING INCHARGE','HELPER COUNT'] },
  item: { sheet: 'LOADING_ITEMS', fields: ['ITEM_ID','LOADING_ID','Customer','LOADING START','LOADING END','TOTAL_HRS','VEHICLE NO','VEHICL FEET','Remarks'] },
  stage: { sheet: 'STAGE_TIME', fields: ['Date','Customer Name','Item Name','Item Code','Invoice Number','OQC End','Loading End','Invoice Time','Gate Out Time','GAP-1 (Production and Loading)','GAP-2 (Loading and Billing)','GAP-3 (Billing and Gate Out)'] }
};

function doGet(e) {
  const callback = String(e && e.parameter && e.parameter.callback || '');
  if (!/^[A-Za-z_$][\w.$]{0,100}$/.test(callback)) return loadingWebText_('Invalid callback.');
  let result;
  try { result = loadingWebLoad_(); } catch (err) { result = {ok:false,error:String(err && err.message || err)}; }
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents || '{}');
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      loadingWebReady_();
      if (body.action === 'save' && body.header && Array.isArray(body.items)) loadingWebSave_(body);
      else if (body.action === 'deleteItem' && body.itemId) loadingWebDeleteItem_(String(body.itemId));
      else throw new Error('Invalid save or delete request.');
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
    return loadingWebText_(JSON.stringify({ok:true,id:String(body.header && body.header.id || body.itemId || '')}));
  } catch (err) { return loadingWebText_(JSON.stringify({ok:false,error:String(err && err.message || err)})); }
}
function loadingWebReady_() {
  if (PropertiesService.getScriptProperties().getProperty('LOADING_DIRECT_READY') !== 'v16') throw new Error('One-table setup pending: run setupLoadingDirectReport in Apps Script.');
}
function loadingWebSS_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Bind this script to the LOADING TIMESTAMP Google Sheet.');
  return ss;
}
function loadingWebDeletedItemIds_() {
  try { const ids=JSON.parse(PropertiesService.getScriptProperties().getProperty('LOADING_WEB_DELETED_ITEM_IDS') || '[]'); return Array.isArray(ids) ? [...new Set(ids.map(String))] : []; } catch (_) { return []; }
}
function loadingWebLoad_() {
  loadingWebReady_();
  const ss=loadingWebSS_(), tz=ss.getSpreadsheetTimeZone();
  const sheet=loadingWebRequireTable_(ss,LOADING_LOG_WEB_TABLES.report);
  const deletedItemIds=loadingWebDeletedItemIds_(), deleted=new Set(deletedItemIds);
  const rows=loadingWebReadRows_(sheet,LOADING_LOG_WEB_TABLES.report,tz).filter(r=>r.ENTRY_ID && !deleted.has(r.ENTRY_ID));
  const headers=new Map(), items=[];
  rows.forEach(r=>{
    // Per-row metadata is authoritative, even after a direct Sheet correction.
    // Split a group in the app when one row's date/shift/incharge/helpers differ.
    const id=loadingWebGroupId_(r);
    if (!headers.has(id)) headers.set(id,{id,date:r.DATE,shift:r.SHIFT,incharge:r['LOADING INCHARGE'],helperCount:Number(r['HELPER COUNT'])||0,createdAt:r.DATE});
    items.push({id:r.ENTRY_ID,loadingId:id,customer:r.CUSTOMER,start:r['LOADING START'],end:r['LOADING END'],totalHours:Number(r.TOTAL_HRS)||0,vehicleNo:r['VEHICLE NO'],vehicleFeet:r['VEHICLE FEET'],remarks:r.REMARKS,check:r.CHECK});
  });
  const stageSheet=ss.getSheetByName('STAGE_TIME');
  const stageRows=stageSheet ? loadingWebReadRows_(stageSheet,LOADING_LOG_WEB_TABLES.stage,tz).map(r=>({date:r.Date,customerName:r['Customer Name'],itemName:r['Item Name'],itemCode:r['Item Code'],invoiceNumber:r['Invoice Number'],oqcEnd:r['OQC End'],loadingEnd:r['Loading End'],invoiceTime:r['Invoice Time'],gateOutTime:r['Gate Out Time'],gap1:r['GAP-1 (Production and Loading)'],gap2:r['GAP-2 (Loading and Billing)'],gap3:r['GAP-3 (Billing and Gate Out)']})).filter(r=>r.invoiceNumber||r.itemCode||r.itemName) : [];
  const customerSheet=ss.getSheetByName('CUSTOMER_MASTER');
  const customerNames=customerSheet && customerSheet.getLastRow()>1 ? [...new Set(customerSheet.getRange(2,1,customerSheet.getLastRow()-1,1).getDisplayValues().flat().map(v=>String(v).trim()).filter(Boolean))] : [];
  return {ok:true,storageMode:'direct-report-v16',headers:[...headers.values()],items,stageRows,customerNames,deletedItemIds};
}
function loadingWebGroupId_(r) {
  return String(r.SHIFT_ID || ('entry:' + r.ENTRY_ID));
}
function loadingWebHours_(start,end) {
  return Math.round(((loadingWebTimeFraction_(end)-loadingWebTimeFraction_(start)+1)%1)*2400)/100;
}
function loadingWebCheck_(h,item) {
  const issues=[];
  if (!h.date || !h.shift || !h.incharge) issues.push('MISSING SHIFT DETAILS');
  if (!item.customer || !item.vehicleNo) issues.push('MISSING VEHICLE DETAILS');
  try {
    const start=loadingWebTimeFraction_(item.start)*24;
    if (loadingWebHours_(item.start,item.end)>8) issues.push('CHECK AM/PM / DURATION > 8 HRS');
    if (String(h.shift).toUpperCase()==='DAY' && (start<6 || start>=18)) issues.push('CHECK DATE / DAY SHIFT TIME');
    if (String(h.shift).toUpperCase()==='NIGHT' && start>=8 && start<16) issues.push('CHECK DATE / NIGHT SHIFT TIME');
  } catch (_) { issues.push('INVALID / MISSING TIME'); }
  return issues.join('; ') || 'OK';
}
function loadingWebRow_(h,item,stamp) {
  let start='', end='', hours='';
  try { start=loadingWebTimeFraction_(item.start); } catch (_) { start=String(item.start||''); }
  try { end=loadingWebTimeFraction_(item.end); } catch (_) { end=String(item.end||''); }
  if (typeof start==='number' && typeof end==='number') hours=loadingWebHours_(item.start,item.end);
  return [loadingWebParseDate_(h.date),String(h.incharge).trim().toUpperCase(),String(item.customer).trim(),start,end,hours,String(item.vehicleNo).trim().toUpperCase(),String(item.vehicleFeet||''),String(item.remarks||''),String(h.shift).toUpperCase(),Number(h.helperCount)||0,String(item.id),String(h.id),loadingWebCheck_(h,item),stamp];
}
function loadingWebValidate_(h,item) {
  if (!h.id || !h.date || !['DAY','NIGHT'].includes(String(h.shift).toUpperCase()) || !String(h.incharge||'').trim() || !Number.isFinite(Number(h.helperCount)) || Number(h.helperCount)<0) throw new Error('Invalid shift details.');
  loadingWebParseDate_(h.date);
  if (item) {
    if (!item.id || !String(item.customer||'').trim() || !String(item.vehicleNo||'').trim()) throw new Error('Required vehicle details are missing.');
    loadingWebTimeFraction_(item.start); loadingWebTimeFraction_(item.end);
  }
}
function loadingWebSave_(body) {
  const ss=loadingWebSS_(), sheet=loadingWebRequireTable_(ss,LOADING_LOG_WEB_TABLES.report), h=body.header;
  loadingWebValidate_(h); body.items.forEach(item=>loadingWebValidate_(h,item)); // Validate whole request before mutation.
  const deleted=new Set(loadingWebDeletedItemIds_()), stamp=Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd HH:mm:ss');
  const last=sheet.getLastRow(), raw=last>1 ? sheet.getRange(2,1,last-1,15).getValues() : [];
  const normalized=loadingWebReadRows_(sheet,LOADING_LOG_WEB_TABLES.report,ss.getSpreadsheetTimeZone());
  const positions=new Map(raw.map((r,i)=>[String(r[11]),i+2]));
  if (positions.size!==raw.length) throw new Error('Duplicate ENTRY_ID in LOADING_REPORT. Correct it before saving.');
  if (!body.items.length) {
    normalized.forEach((r,i)=>{ if (loadingWebGroupId_(r)===String(h.id) || r.SHIFT_ID===String(h.id)) {
      const item={id:r.ENTRY_ID,customer:r.CUSTOMER,vehicleNo:r['VEHICLE NO'],vehicleFeet:r['VEHICLE FEET'],remarks:r.REMARKS,start:r['LOADING START'],end:r['LOADING END']};
      sheet.getRange(i+2,1,1,15).setValues([loadingWebRow_(h,item,stamp)]);
    }});
  } else body.items.forEach(item=>{
    if (deleted.has(String(item.id))) return; // A stale phone must not recreate a deleted vehicle.
    const row=positions.get(String(item.id)) || sheet.getLastRow()+1;
    if (row>sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(),Math.max(100,row-sheet.getMaxRows()));
    sheet.getRange(row,1,1,15).setValues([loadingWebRow_(h,item,stamp)]);
    positions.set(String(item.id),row);
  });
  loadingWebFormat_(sheet);
}
function loadingWebDeleteItem_(itemId) {
  const sheet=loadingWebRequireTable_(loadingWebSS_(),LOADING_LOG_WEB_TABLES.report);
  const ids=sheet.getLastRow()>1 ? sheet.getRange(2,12,sheet.getLastRow()-1,1).getDisplayValues().flat() : [];
  for (let i=ids.length-1;i>=0;i--) if (String(ids[i])===itemId) sheet.deleteRow(i+2);
  const idsDeleted=loadingWebDeletedItemIds_(); if (!idsDeleted.includes(itemId)) idsDeleted.push(itemId);
  PropertiesService.getScriptProperties().setProperty('LOADING_WEB_DELETED_ITEM_IDS',JSON.stringify(idsDeleted));
}

/** One-time cutover; rerunning never overwrites direct records from archives. */
function setupLoadingDirectReport() {
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss=loadingWebSS_(), props=PropertiesService.getScriptProperties();
    if (props.getProperty('LOADING_DIRECT_READY')==='v16') return 'Already active: one-table report.';
    const tz=ss.getSpreadsheetTimeZone();
    const headerSheet=loadingWebRequireTable_(ss,LOADING_LOG_WEB_TABLES.header), itemSheet=loadingWebRequireTable_(ss,LOADING_LOG_WEB_TABLES.item);
    const headers=new Map(loadingWebReadRows_(headerSheet,LOADING_LOG_WEB_TABLES.header,tz).map(r=>[r.LOADING_ID,r]));
    const legacy=loadingWebReadRows_(itemSheet,LOADING_LOG_WEB_TABLES.item,tz).filter(r=>r.ITEM_ID);
    const legacyIds=new Set();
    legacy.forEach(r=>{if (legacyIds.has(r.ITEM_ID)) throw new Error('Duplicate legacy ITEM_ID: '+r.ITEM_ID); legacyIds.add(r.ITEM_ID); if (!headers.has(r.LOADING_ID)) throw new Error('Missing legacy header for '+r.ITEM_ID);});
    let sheet=ss.getSheetByName('LOADING_REPORT');
    const stamp=Utilities.formatDate(new Date(),tz,'yyyy-MM-dd HH:mm:ss');
    const deleted=new Set(loadingWebDeletedItemIds_());
    const migrated=legacy.filter(r=>!deleted.has(r.ITEM_ID)).map(r=>{
      const x=headers.get(r.LOADING_ID), h={id:r.LOADING_ID,date:x.DATE,shift:x.SHIFT,incharge:x['LOADING INCHARGE'],helperCount:Number(x['HELPER COUNT'])||0};
      const item={id:r.ITEM_ID,customer:r.Customer,start:r['LOADING START'],end:r['LOADING END'],vehicleNo:r['VEHICLE NO'],vehicleFeet:r['VEHICL FEET'],remarks:r.Remarks};
      return loadingWebRow_(h,item,stamp);
    });
    if (sheet && !ss.getSheetByName('LOADING_REPORT_FORMULA_BACKUP')) {
      const backup=sheet.copyTo(ss).setName('LOADING_REPORT_FORMULA_BACKUP');
      backup.getDataRange().copyTo(backup.getDataRange(),{contentsOnly:true}); backup.hideSheet();
    }
    if (!sheet) sheet=ss.insertSheet('LOADING_REPORT');
    // Preserve new direct-only entries staged before cutover, but import latest legacy edits.
    const first=sheet.getRange(1,1,1,15).getDisplayValues()[0];
    const staged=first[11]==='ENTRY_ID' && sheet.getLastRow()>1 ? sheet.getRange(2,1,sheet.getLastRow()-1,15).getValues().filter(r=>r[11] && !legacyIds.has(String(r[11])) && !deleted.has(String(r[11]))) : [];
    const all=migrated.concat(staged).sort((a,b)=>Number(a[0])-Number(b[0]) || Number(a[3])-Number(b[3]));
    if (all.length+1>sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(),all.length+1-sheet.getMaxRows());
    if (sheet.getMaxColumns()<15) sheet.insertColumnsAfter(sheet.getMaxColumns(),15-sheet.getMaxColumns());
    sheet.getRange(1,1,sheet.getMaxRows(),15).clearContent();
    sheet.getRange(1,1,all.length+1,15).setValues([LOADING_LOG_WEB_TABLES.report.fields,...all]);
    loadingWebFormat_(sheet); SpreadsheetApp.flush();
    props.setProperty('LOADING_DIRECT_READY','v16');
    sheet.showSheet(); headerSheet.hideSheet(); itemSheet.hideSheet();
    // Keep derived totals/checks current if a supervisor corrects a row in Sheet.
    if (!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='loadingDirectOnEdit')) ScriptApp.newTrigger('loadingDirectOnEdit').forSpreadsheet(ss).onEdit().create();
    return all.length+' vehicle rows migrated. App now uses only LOADING_REPORT.';
  } finally { lock.releaseLock(); }
}
function loadingDirectOnEdit(e) {
  if (!e || !e.range || e.range.getSheet().getName()!=='LOADING_REPORT' || e.range.getRow()+e.range.getNumRows()-1<2) return;
  loadingWebReady_();
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss=loadingWebSS_(), sheet=e.range.getSheet(), start=Math.max(e.range.getRow(),2), end=Math.min(e.range.getRow()+e.range.getNumRows()-1,sheet.getLastRow());
    if (end<start) return;
    const rows=sheet.getRange(start,1,end-start+1,15).getValues(), shown=sheet.getRange(start,1,end-start+1,15).getDisplayValues();
    const ids=sheet.getRange(2,12,Math.max(sheet.getLastRow()-1,1),1).getDisplayValues().flat();
    rows.forEach((r,i)=>{
      if (!r.some(v=>v!=='')) return;
      if (!r[11]) r[11]=Utilities.getUuid();
      const firstCol=e.range.getColumn(), lastCol=firstCol+e.range.getNumColumns()-1;
      if (!r[12] || [1,2,10,11].some(c=>c>=firstCol && c<=lastCol)) r[12]=Utilities.getUuid();
      const h={id:String(r[12]),date:loadingWebNormalizeDate_(r[0],shown[i][0],ss.getSpreadsheetTimeZone()),shift:r[9],incharge:r[1],helperCount:r[10]};
      const item={id:r[11],customer:r[2],start:loadingWebNormalizeTime_(r[3],shown[i][3],ss.getSpreadsheetTimeZone()),end:loadingWebNormalizeTime_(r[4],shown[i][4],ss.getSpreadsheetTimeZone()),vehicleNo:r[6],vehicleFeet:r[7],remarks:r[8]};
      try { loadingWebValidate_(h,item); r[0]=loadingWebParseDate_(h.date); r[3]=loadingWebTimeFraction_(item.start); r[4]=loadingWebTimeFraction_(item.end); r[5]=loadingWebHours_(item.start,item.end); r[13]=loadingWebCheck_(h,item); }
      catch(err) { r[5]=''; r[13]='CHECK: '+err.message; }
      if (ids.filter(id=>String(id)===String(r[11])).length>1) r[13]='DUPLICATE ENTRY_ID';
      r[14]=Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(start+i,1,1,15).setValues([r]);
    });
  } finally {lock.releaseLock();}
}
function loadingWebFormat_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,15).setFontWeight('bold').setWrap(true);
  const count=Math.max(sheet.getMaxRows()-1,1);
  sheet.getRange(2,1,count,1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2,4,count,2).setNumberFormat('HH:mm');
  sheet.getRange(2,6,count,1).setNumberFormat('0.00');
  sheet.getRange(2,11,count,1).setNumberFormat('0');
  sheet.getRange(2,10,count,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['DAY','NIGHT'],true).setAllowInvalid(false).build());
  sheet.getRange(2,14,count,1).setWrap(true);
  const filter=sheet.getFilter();
  if (!filter) sheet.getRange(1,1,sheet.getMaxRows(),15).createFilter();
  else if (filter.getRange().getNumColumns()!==15 || filter.getRange().getNumRows()<sheet.getMaxRows()) {
    const criteria=[]; for(let col=1;col<=Math.min(15,filter.getRange().getNumColumns());col++) criteria[col]=filter.getColumnFilterCriteria(col);
    filter.remove(); const expanded=sheet.getRange(1,1,sheet.getMaxRows(),15).createFilter();
    criteria.forEach((value,col)=>{if(value) expanded.setColumnFilterCriteria(col,value);});
  }
  const rule=SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($L2<>"",$N2<>"OK")').setBackground('#fff1cc').setRanges([sheet.getRange(2,1,count,15)]).build();
  const unrelated=sheet.getConditionalFormatRules().filter(r=>!r.getRanges().some(x=>x.getColumn()===1 && x.getNumColumns()===15));
  sheet.setConditionalFormatRules(unrelated.concat([rule]));
}

function loadingWebRequireTable_(ss, definition) {
  const sheet = ss.getSheetByName(definition.sheet);
  if (!sheet) throw new Error('Required sheet tab not found: ' + definition.sheet);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const names = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v).trim().toUpperCase());
  const missing = definition.fields.filter(field => names.indexOf(field.toUpperCase()) < 0);
  if (missing.length) throw new Error('Missing columns in ' + definition.sheet + ': ' + missing.join(', '));
  return sheet;
}

function loadingWebReadRows_(sheet, definition, timezone) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || !lastCol) return [];
  const raw = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const names = raw[0].map(v => String(v).trim().toUpperCase());
  return raw.slice(1).map((row, rowIndex) => {
    const item = {};
    definition.fields.forEach(field => {
      const col = names.indexOf(field.toUpperCase());
      const value = row[col];
      if (field.toUpperCase() === 'DATE') item[field] = loadingWebNormalizeDate_(value, display[rowIndex + 1][col], timezone);
      else if (['LOADING_ITEMS','LOADING_REPORT'].includes(definition.sheet) && (field === 'LOADING START' || field === 'LOADING END')) item[field] = loadingWebNormalizeTime_(value, display[rowIndex + 1][col], timezone);
      else if (['OQC END', 'LOADING END', 'INVOICE TIME', 'GATE OUT TIME'].indexOf(field.toUpperCase()) >= 0) item[field] = loadingWebNormalizeTimestamp_(value, display[rowIndex + 1][col], timezone);
      else if (field === 'LOADING START' || field === 'LOADING END') item[field] = String(display[rowIndex + 1][col] || '').trim().slice(0, 5);
      else item[field] = value instanceof Date ? Utilities.formatDate(value, timezone, 'HH:mm:ss') : String(display[rowIndex + 1][col] || '').trim();
    });
    return item;
  });
}

function loadingWebParseDate_(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error('Date must use YYYY-MM-DD.');
  const y = Number(match[1]), m = Number(match[2]), d = Number(match[3]);
  const check = new Date(Date.UTC(y, m - 1, d));
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== m - 1 || check.getUTCDate() !== d) throw new Error('Invalid date.');
  // Numeric spreadsheet date avoids script/spreadsheet timezone conversion.
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

function loadingWebNormalizeDate_(value, shown, timezone) {
  if (value instanceof Date) return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  const text = String(shown || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const dmy = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmy) return dmy[3] + '-' + dmy[2].padStart(2, '0') + '-' + dmy[1].padStart(2, '0');
  return text;
}

function loadingWebNormalizeTimestamp_(value, shown, timezone) {
  if (value instanceof Date) return Utilities.formatDate(value, timezone, 'yyyy-MM-dd HH:mm:ss');
  const text = String(shown || '').trim();
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}:\d{2})(?::(\d{2}))?$/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3] + ' ' + iso[4] + ':' + (iso[5] || '00');
  const dmy = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\s+(\d{1,2}:\d{2})(?::(\d{2}))?$/);
  if (dmy) return dmy[3] + '-' + dmy[2].padStart(2, '0') + '-' + dmy[1].padStart(2, '0') + ' ' + dmy[4].padStart(5, '0') + ':' + (dmy[5] || '00');
  return text;
}

function loadingWebTimeFraction_(value) {
  const match = String(value).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) throw new Error('Time must use HH:MM.');
  const hours = Number(match[1]), minutes = Number(match[2]), seconds = Number(match[3] || 0);
  if (hours > 23 || minutes > 59 || seconds > 59) throw new Error('Invalid time value.');
  return (hours * 3600 + minutes * 60 + seconds) / 86400;
}

function loadingWebText_(value) {
  return ContentService.createTextOutput(value).setMimeType(ContentService.MimeType.TEXT);
}


function loadingWebNormalizeTime_(value, shown, timezone) {
  if (typeof value === 'number') {
    const minutes = Math.round(((value % 1 + 1) % 1) * 1440) % 1440;
    return String(Math.floor(minutes / 60)).padStart(2, '0') + ':' + String(minutes % 60).padStart(2, '0');
  }
  // Time-only Dates use an 1899 base date; displayed clock avoids historical timezone offsets.
  const text = String(shown || '').trim();
  const match = text.match(/(?:^|[ T])(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (match) {
    let hour = Number(match[1]);
    if (match[3]) hour = hour % 12 + (match[3].toUpperCase() === 'PM' ? 12 : 0);
    if (hour < 24 && Number(match[2]) < 60) return String(hour).padStart(2, '0') + ':' + match[2];
  }
  if (value instanceof Date) return Utilities.formatDate(value, timezone, 'HH:mm');
  return text;
}

