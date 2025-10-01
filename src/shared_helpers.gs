/* =========================
   SHARED HELPERS
========================= */
function KNB_allGids_(){
  const all = [
    KNB_CFG.GID.REQUESTED,
    KNB_CFG.GID.INPROGRESS,
    KNB_CFG.GID.FORAPPROVAL,
    KNB_CFG.GID.DONE,
    ...Object.values(KNB_CFG.ROUTE)
  ].map(Number);
  return Array.from(new Set(all)).filter(Boolean);
}

// micro-cache for fast sheet lookup
const __KNB_SHEETS_CACHE = { id: null, byGid: {} };
function KNB_sheetById_(gid){
  const ss = SpreadsheetApp.getActive();
  if (__KNB_SHEETS_CACHE.id !== ss.getId()){
    __KNB_SHEETS_CACHE.id = ss.getId();
    __KNB_SHEETS_CACHE.byGid = {};
    ss.getSheets().forEach(s => { __KNB_SHEETS_CACHE.byGid[s.getSheetId()] = s; });
  }
  return __KNB_SHEETS_CACHE.byGid[Number(gid)] || null;
}

function KNB_headerMap_(sh){
  const n = Math.max(1, sh.getLastColumn());
  const row1 = sh.getRange(1,1,1,n).getDisplayValues()[0].map(v=>String(v).trim());
  const m = {}; row1.forEach((h,i)=>{ if(h) m[h]=i+1; }); return m;
}
function KNB_headerIndex_(sh){ return KNB_headerMap_(sh); }
function KNB_eachRow(ss, gid, fn){
  const sh = KNB_sheetById_(gid); if (!sh) return;
  const idx = KNB_headerIndex_(sh); if (!idx[KNB_CFG.COL.STATUS]) return;
  for (let r = sh.getLastRow(); r >= 2; r--) fn(sh, r, idx);
}
function KNB_colToA1_(n){ let s='',t=n; while(t>0){ let m=(t-1)%26; s=String.fromCharCode(65+m)+s; t=(t-m-1)/26|0;} return s; }
function KNB_hexToRgb_(hex){ const h=hex.replace('#',''); const n=parseInt(h.length===3?h.split('').map(c=>c+c).join(''):h,16); return {r:(n>>16)&255,g:(n>>8)&255,b:n&255}; }
function KNB_pickTextColor_(hex){
  const {r,g,b}=KNB_hexToRgb_(hex);
  const c=[r,g,b].map(v=>{v/=255;return v<=.03928?v/12.92:Math.pow((v+.055)/1.055,2.4)});
  const L=.2126*c[0]+.7152*c[1]+.0722*c[2];
  return L>0.45?'#000':'#fff';
}
function KNB_sanitize_(s){ return String(s||'').replace(/[\[\]\(\)\n\r]/g,'').slice(0,180); }
function KNB_cellHasValidUrl_(cell){
  try{
    const f = String(cell.getFormula()||'');
    if (/^=HYPERLINK\(/i.test(f)){
      if (/^=HYPERLINK\(\s*"https?:\/\/[^"]+"/i.test(f)) return true;
      if (/^=HYPERLINK\(\s*https?:\/\//i.test(f)) return true;
    }
    const rtv = cell.getRichTextValue && cell.getRichTextValue();
    if (rtv && rtv.getLinkUrl && rtv.getLinkUrl()){
      return /^https?:\/\//i.test(rtv.getLinkUrl());
    }
    const text = String(cell.getDisplayValue()||'').trim();
    return /^https?:\/\/\S+$/i.test(text);
  }catch(_){ return false; }
}

// Rate limit & suppression (per user)
function KNB_userKey_(){ const email=(Session.getActiveUser&&Session.getActiveUser().getEmail())||''; return email?email.toLowerCase():'anon'; }
function KNB_isRateOk_(){ const key='KNB_LAST_'+KNB_userKey_(); const p=PropertiesService.getDocumentProperties(); const last=Number(p.getProperty(key)||0); return Date.now()-last>=KNB_CFG.RATE_LIMIT_MS; }
function KNB_noteRate_(){ const key='KNB_LAST_'+KNB_userKey_(); PropertiesService.getDocumentProperties().setProperty(key, String(Date.now())); }
function KNB_isSuppressed_(){ const p=PropertiesService.getDocumentProperties(); const until=Number(p.getProperty('KNB_SUPPRESS_UNTIL')||0); return Date.now()<until; }
function KNB_suppressEdits_(ms){ const p=PropertiesService.getDocumentProperties(); p.setProperty('KNB_SUPPRESS_UNTIL', String(Date.now()+(ms||500))); }

// ===== Task Details storage (no tooltip; uses hidden sheet) =====
const KNB_DETAILS_SHEET = '_KB_DETAILS';

function KNB_detailsSheet_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(KNB_DETAILS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(KNB_DETAILS_SHEET);
    sh.getRange(1,1,1,2).setValues([['RowID','HTML']]);
    try { sh.hideSheet(); } catch(_){}
  }
  return sh;
}

// Ensure a stable Row ID in the data row; auto-create if empty
function KNB_rowIdFor_(sh, row, idx){
  const c = idx[KNB_CFG.COL.ROWID];
  if (!c) throw new Error('Missing "Row ID" column. Add it to headers or update KNB_CFG.COL.ROWID.');
  let v = String(sh.getRange(row, c).getValue()||'').trim();
  if (!v) { v = Utilities.getUuid(); sh.getRange(row, c).setValue(v); }
  return v;
}

function KNB_detailsRead_(rowId){
  if (!rowId) return '';
  const sh = KNB_detailsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return '';
  const range = sh.getRange(2,1,last-1,2).getValues();
  for (let i=0;i<range.length;i++){
    if (range[i][0] === rowId) return String(range[i][1]||'');
  }
  return '';
}

function KNB_detailsWrite_(rowId, html){
  if (!rowId) return;
  const sh = KNB_detailsSheet_();
  const last = sh.getLastRow();
  if (last >= 2){
    const finder = sh.getRange(2,1,last-1,1).createTextFinder(rowId).matchEntireCell(true).findNext();
    if (finder){ sh.getRange(finder.getRow(), 2).setValue(html); return; }
  }
  sh.appendRow([rowId, html]);
}

function KNB_firstEmptyDataRow_(sh) {
  const idx = KNB_headerIndex_(sh);
  // Choose key columns that indicate a real task row
  const keys = [idx[KNB_CFG.COL.TASK], idx[KNB_CFG.COL.STATUS]];
  const last = Math.max(2, sh.getLastRow());
  const width = Math.max(1, sh.getLastColumn());
  const vals = sh.getRange(2, 1, last - 1, width).getValues();

  // Find last row that has *any* key populated, then append after it
  for (let i = vals.length - 1; i >= 0; i--) {
    const hasKey = keys.some(c => c && String(vals[i][c - 1] || '').trim() !== '');
    if (hasKey) return i + 2 + 1; // next row after last real row
  }
  return 2; // table is empty
}

// Hide "Task Details (HTML)" and "Task Details (Draft)" on ALL boards
function KNB_hideTaskDetailsStorage_AllBoards(){
  const ss = SpreadsheetApp.getActive();
  (KNB_allGids_() || []).forEach(gid => {
    const sh = KNB_sheetById_(gid);
    if (!sh) return;
    const hCol = KNB_findHeaderColumn_(sh, 'Task Details (HTML)');
    const dCol = KNB_findHeaderColumn_(sh, 'Task Details (Draft)');
    if (hCol) try { sh.hideColumn(sh.getRange(1, hCol)); } catch(_){}
    if (dCol) try { sh.hideColumn(sh.getRange(1, dCol)); } catch(_){}
  });
  SpreadsheetApp.getActive().toast('Hidden storage columns on all boards.', '📝', 4);
}

// Helper: case-insensitive header lookup
function KNB_findHeaderColumn_(sh, headerName){
  const n = Math.max(1, sh.getLastColumn());
  const row1 = sh.getRange(1,1,1,n).getDisplayValues()[0];
  const want = String(headerName||'').trim().toLowerCase();
  for (let i=0;i<row1.length;i++){
    if (String(row1[i]||'').trim().toLowerCase() === want) return i+1;
  }
  return 0;
}

// Busy flag (blocks rapid re-entry during moves)
function KNB_setBusy_(ms){ 
  const key = 'KNB_BUSY_'+KNB_userKey_();
  PropertiesService.getDocumentProperties().setProperty(key, String(Date.now() + (ms||1500)));
}
function KNB_isBusy_(){
  const key = 'KNB_BUSY_'+KNB_userKey_();
  const until = Number(PropertiesService.getDocumentProperties().getProperty(key) || 0);
  return Date.now() < until;
}

// ===== UI: blocking "please wait" spinner (auto-close) =====
function KNB_UI_showBusy_(msg, autoMs){
  var ms = Number(autoMs || 2500); // auto-close after this many ms
  var html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"></head>' +
    '<body style="margin:0;font:14px/1.6 system-ui,-apple-system,Segoe UI,Roboto">' +
      '<div style="padding:16px;min-width:280px;display:flex;gap:10px;align-items:center">' +
        '<div style="width:16px;height:16px;border:2px solid #cbd5e1;border-top-color:#111827;border-radius:50%;animation:spin 1s linear infinite"></div>' +
        '<div>'+ (msg || 'Working…') +'</div>' +
      '</div>' +
      '<style>@keyframes spin{to{transform:rotate(360deg)}}</style>' +
      '<script>' +
        '(function(){ var ms='+ ms +'; if(ms>0){ setTimeout(function(){ try{google.script.host.close()}catch(e){} }, ms); } })();' +
      '</script>' +
    '</body></html>'
  ).setWidth(320).setHeight(90);
  SpreadsheetApp.getUi().showModalDialog(html, 'Please wait…');
}

// Optional: explicit closer (kept for manual close if you still call it)
function KNB_UI_closeBusy_(){
  var html = HtmlService.createHtmlOutput('<script>google.script.host.close()</script>')
    .setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, '');
}

// Fast header index cache (per sheetId)
var __KNB_HDR_CACHE = { ssId:null, map:{} };

function KNB_headerIndex_CACHED_(sh){
  var ss = SpreadsheetApp.getActive();
  if (__KNB_HDR_CACHE.ssId !== ss.getId()){
    __KNB_HDR_CACHE = { ssId:ss.getId(), map:{} };
  }
  var id = sh.getSheetId();
  if (__KNB_HDR_CACHE.map[id]) return __KNB_HDR_CACHE.map[id];

  var n = Math.max(1, sh.getLastColumn());
  var row1 = sh.getRange(1,1,1,n).getDisplayValues()[0].map(function(v){ return String(v).trim(); });
  var m = {};
  row1.forEach(function(h,i){ if(h) m[h]=i+1; });
  __KNB_HDR_CACHE.map[id] = m;
  return m;
}

