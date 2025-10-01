/* =========================
   PRIVATE VIEWS
========================= */
// Build a header→index map (1-based) from an array of header texts
function KNB_indexFromArray_(arr){
  const map = {};
  (arr || []).forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i + 1; 
  });
  return map;
}

function KNB_PVX_plainText_(html){
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(div|p|li|h[1-6])>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .trim();
}

function KNB_PVX_openMyPrivate(){
  const who = KNB_PVX_promptName_('Open My Private'); if(!who) return;
  KNB_PVX_wait_('Building your private view…');
  try{
    const url = KNB_PVX_upsertPrivateForName_(who);
    KNB_PVX_linkDialog_('Your private tasks are ready:', url);
  } finally { KNB_PVX_closeWait_(); }
}
function KNB_PVX_refreshMyPrivate(){
  const who = KNB_PVX_promptName_('Refresh My Private'); if(!who) return;
  KNB_PVX_wait_('Refreshing your private view…');
  try{
    KNB_PVX_upsertPrivateForName_(who);
    SpreadsheetApp.getActive().toast('Private view refreshed for '+who, 'Private Views', 4);
  } finally { KNB_PVX_closeWait_(); }
}
function KNB_PVX_publishAll(){
  KNB_PVX_wait_('Publishing all private views…');
  try{
    Object.keys(KNB_PVX_DEST_IDS).forEach(name => KNB_PVX_upsertPrivateForName_(name));
    SpreadsheetApp.getActive().toast('Published '+Object.keys(KNB_PVX_DEST_IDS).length+' private file(s).','Private Views',5);
  } finally { KNB_PVX_closeWait_(); }
}
function KNB_PVX_installHourlyRefresh(){
  ScriptApp.getProjectTriggers().forEach(t=>{ if (t.getHandlerFunction && t.getHandlerFunction()==='KNB_PVX_hourlyRefresh') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('KNB_PVX_hourlyRefresh').timeBased().everyHours(1).create();
  SpreadsheetApp.getActive().toast('Hourly refresh installed.', 'Private Views', 3);
}
function KNB_PVX_removeHourlyRefresh(){
  ScriptApp.getProjectTriggers().forEach(t=>{ if (t.getHandlerFunction && t.getHandlerFunction()==='KNB_PVX_hourlyRefresh') ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getActive().toast('Hourly refresh removed.', 'Private Views', 3);
}
function KNB_PVX_hourlyRefresh(){
  const lock = LockService.getScriptLock(); if(!lock.tryLock(2000)) return;
  try{ Object.keys(KNB_PVX_DEST_IDS).forEach(name => KNB_PVX_upsertPrivateForName_(name)); }
  finally { lock.releaseLock(); }
}
function KNB_PVX_upsertPrivateForName_(name){
  const sections = [
    { gid: KNB_CFG.GID.REQUESTED,   status:'Requested'   },
    { gid: KNB_CFG.GID.INPROGRESS,  status:'In Progress' },
    { gid: KNB_CFG.GID.FORAPPROVAL, status:'For Approval' }
  ];

  const cols = KNB_PVX_COLS.slice();      // output column order
  const out  = [ cols ];                  // header row
  const meta = [];                        // per-section layout + notes to attach later

  sections.forEach(sec=>{
    const sh = KNB_sheetById_(sec.gid); if (!sh) return;

    // Map source headers
    const hdr = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn()))
                  .getDisplayValues()[0].map(String);
    const map = KNB_indexFromArray_(hdr);
    if (!map[KNB_CFG.COL.ASSIGNEE]) return;

    // Column indexes we might read from (1-based)
    const cAss    = map[KNB_CFG.COL.ASSIGNEE] || 0;
    const cRowId  = map[KNB_CFG.COL.ROWID]    || 0;
    const cDetVis = map[KNB_CFG.COL.DETAILS]  || 0; // visible "Task Details"
    // common hidden column names you’ve used in this project
    const cDetHtml =
      map['Task Details (HTML)'] ||
      map['Task Details HTML']   ||
      0;

    const last   = sh.getLastRow();
    const group  = KNB_PVX_GROUPS[name] || [name];
    const bucket = [];
    const detailNotes = [];               // plain-text notes aligned to exported rows

    if (last >= 2){
      const width = Math.max(1, sh.getLastColumn());

      // Get values row-wise
      const vals   = sh.getRange(2,1,last-1,width).getDisplayValues();

      // If we might need cell notes, fetch them in one shot to avoid per-cell calls
      const detNotesVis = cDetVis ? sh.getRange(2, cDetVis, last-1, 1).getNotes() : null;

      for (let i=0;i<vals.length;i++){
        const rowVals = vals[i];

        // Skip non-assignee rows
        const ass = cAss ? String(rowVals[cAss-1]||'').trim() : '';
        if (!ass || group.indexOf(ass) === -1) continue;

        // 1) Row ID → storage
        let html = '';
        if (cRowId){
          const rowId = String(rowVals[cRowId-1]||'').trim();
          if (rowId) html = KNB_detailsRead_(rowId);
        }

        // 2) Fallback: hidden "Task Details (HTML)" column in the board
        if (!html && cDetHtml) html = String(rowVals[cDetHtml-1]||'').trim();

        // 3) Fallback: visible Task Details NOTE
        if (!html && detNotesVis && detNotesVis[i] && detNotesVis[i][0]){
          // If somebody stored raw HTML in a note, keep it. If it’s already plain text, we’ll just show it.
          html = detNotesVis[i][0];
        }

        const notePlain = KNB_PVX_plainText_(html);
        // Build the outgoing row following KNB_PVX_COLS order
        const picked = cols.map(h=>{
          if (h === KNB_CFG.COL.DETAILS){
            return notePlain ? '📝' : ''; // marker in the private sheet
          }
          const c = map[h]; return c ? rowVals[c-1] : '';
        });

        bucket.push(picked);
        detailNotes.push(notePlain);
      }
    }

    // section header divider row
    out.push(['— '+sec.status+' —'].concat(Array(cols.length-1).fill('')));
    const headerRow = out.length;
    const dataStart = headerRow + 1;
    bucket.forEach(r=>out.push(r));
    const dataEnd = dataStart + bucket.length - 1;

    meta.push({ headerRow, dataStart, dataEnd, detailNotes });
  });

  // Open/create destination (and apply sharing per your helper if present)
  const ssDest = KNB_PVX_openOrCreateDest_ ? KNB_PVX_openOrCreateDest_(name)
                                           : SpreadsheetApp.openById(KNB_PVX_DEST_IDS[name]);
  const shOut = ssDest.getSheetByName('Overview') || ssDest.getSheets()[0] || ssDest.insertSheet('Overview');
  shOut.setName('Overview');

  // Clean slate
  try { shOut.getRange(1,1,shOut.getMaxRows(),shOut.getMaxColumns()).breakApart(); } catch(_){}
  if (shOut.getFilter()) try { shOut.getFilter().remove(); } catch(_){}
  shOut.clear();

  // Ensure at least 1 body row
  if (out.length === 1) out.push(Array(cols.length).fill(''));

  // Write table
  shOut.getRange(1,1,out.length, cols.length).setValues(out);

  // Header style
  shOut.setFrozenRows(1);
  shOut.getRange(1,1,1,cols.length)
    .setBackground('#111827').setFontColor('#ffffff').setFontWeight('bold').setVerticalAlignment('middle');
  try { shOut.setRowHeight(1, 34); } catch(_){}

  // Section shading + zebra
  const rules = [];
  meta.forEach(m=>{
    shOut.getRange(m.headerRow,1,1,cols.length).setBackground('#f1f5f9');
    try { shOut.setRowHeight(m.headerRow, 26); } catch(_){}
    if (m.dataEnd >= m.dataStart){
      const rng = shOut.getRange(m.dataStart,1, m.dataEnd-m.dataStart+1, cols.length);
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISEVEN(ROW()-'+(m.dataStart-1)+')')
        .setBackground('#fafbfc').setRanges([rng]).build());
      try { rng.setBorder(false,false,true,false,false,false,'#e5e7eb', SpreadsheetApp.BorderStyle.SOLID); } catch(_){}
    }
  });
  shOut.setConditionalFormatRules(rules);

  // Attach Task Details notes (plain text)
  const colMap = KNB_indexFromArray_(cols);
  const detailsCol = colMap[KNB_CFG.COL.DETAILS] || 0;
  if (detailsCol){
    meta.forEach(m=>{
      if (m.dataEnd >= m.dataStart && m.detailNotes && m.detailNotes.length){
        const noteMatrix = m.detailNotes.map(n=>[n || '']);
        shOut.getRange(m.dataStart, detailsCol, noteMatrix.length, 1).setNotes(noteMatrix);
      }
    });
  }

  // Date formats + filter
  const lastRow = shOut.getLastRow();
  ['Start Date','Due Date','End Date'].forEach(h=>{
    const c = colMap[h]; if (c) shOut.getRange(2,c, Math.max(0,lastRow-1),1)
      .setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('right');
  });

  const widths = [110,120,180,260,300,130,110,110,110,110];
  for (let i=1;i<=cols.length;i++){ try{ shOut.setColumnWidth(i, widths[i-1]||120); }catch(_){ } }
  shOut.getRange(1,1, Math.max(1,lastRow), cols.length).createFilter();

  return ssDest.getUrl();
}


function KNB_PVX_promptName_(title){
  const names = Object.keys(KNB_PVX_DEST_IDS);
  const txt = 'Enter your name exactly as listed:\n'+names.join(', ');
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(title, txt, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return '';
  const name = (r.getResponseText()||'').trim();
  if (!KNB_PVX_DEST_IDS[name]) { ui.alert('Not found. Valid options:\n'+names.join(', ')); return ''; }
  return name;
}
function KNB_PVX_wait_(msg){
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px/1.6 Arial,sans-serif;padding:14px;display:flex;gap:10px;align-items:center">'+
    '<div style="width:16px;height:16px;border:2px solid #ccc;border-top-color:#111;border-radius:50%;animation:spin 1s linear infinite"></div>'+
    '<div>'+ (msg||'Working…') +'</div>'+
    '<style>@keyframes spin{to{transform:rotate(360deg)}}</style>'+
    '</div>'
  ).setWidth(320).setHeight(90);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Please wait…');
}
function KNB_PVX_closeWait_(){
  const html = HtmlService.createHtmlOutput('<script>google.script.host.close()</script>').setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Closing…');
}
function KNB_PVX_linkDialog_(title, url){
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px/1.6 Arial,sans-serif;padding:14px">'+
      '<p>'+title+'</p>'+
      '<p><a target="_blank" href="'+url+'">Open your private spreadsheet</a></p>'+
      '<p>You can bookmark this link.</p>'+
    '</div>'
  ).setWidth(420).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, 'Private View Ready');
}

// ---- CONFIG knob (put this in config.gs if you want) ----
const KNB_PVX_SHARE = {
  mode: 'ANYONE',   // 'ANYONE' | 'DOMAIN' | 'PRIVATE'
  role: 'EDITOR'    // 'VIEWER' | 'EDITOR'
};

// Map role -> DriveApp
function KNB__pvxRole_(role){
  return (String(role||'').toUpperCase()==='EDITOR')
    ? DriveApp.Permission.EDIT
    : DriveApp.Permission.VIEW;
}
function KNB__pvxApplySharing_(file){
  const role = KNB__pvxRole_(KNB_PVX_SHARE.role);
  const mode = String(KNB_PVX_SHARE.mode||'').toUpperCase();
  if (mode === 'ANYONE'){
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, role);
  } else if (mode === 'DOMAIN'){
    // Works inside Google Workspace domains
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, role);
  } else {
    // PRIVATE (do nothing). You must explicitly share to users elsewhere.
  }
}

// Open, or create if missing; also auto-share per KNB_PVX_SHARE
function KNB_PVX_openOrCreateDest_(name){
  const id = KNB_PVX_DEST_IDS[name];
  if (id) {
    try {
      const ss = SpreadsheetApp.openById(id);
      try { KNB__pvxApplySharing_(DriveApp.getFileById(ss.getId())); } catch(_){}
      return ss;
    } catch (e) {
      // fall-through to create
    }
  }
  // Create new file owned by the script’s owner
  const ss = SpreadsheetApp.create('Private - ' + name);
  try { KNB__pvxApplySharing_(DriveApp.getFileById(ss.getId())); } catch(_){}
  // Tell you to update config for next time
  const cfgLine = `Add/update KNB_PVX_DEST_IDS["${name}"] = "${ss.getId()}";`;
  SpreadsheetApp.getUi().alert(
    'Created private spreadsheet for "'+name+'".\n\nID:\n' + ss.getId() +
    '\n\n' + cfgLine
  );
  return ss;
}
