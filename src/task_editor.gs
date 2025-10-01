// 📝 icon shown in the visible "Task Details" cell
const KNB_MTE_ICON = '📝';

// Ensure hidden storage columns exist & hide them
function KNB_MTE_ensureColumns_(){
  const sh = SpreadsheetApp.getActiveSheet();
  const idx0 = KNB_headerIndex_(sh);

  // Insert "Task Details (HTML)" after "Task Details"
  if (!idx0['Task Details']) throw new Error('Missing header "Task Details".');
  if (!idx0['Task Details (HTML)']) {
    sh.insertColumnAfter(idx0['Task Details']);
    sh.getRange(1, idx0['Task Details'] + 1).setValue('Task Details (HTML)');
  }

  // Recompute because columns shifted
  const idx1 = KNB_headerIndex_(sh);
  if (!idx1['Task Details (Draft)']) {
    sh.insertColumnAfter(idx1['Task Details (HTML)']);
    sh.getRange(1, idx1['Task Details (HTML)'] + 1).setValue('Task Details (Draft)');
  }

  // Hide the storage columns
  const idx2 = KNB_headerIndex_(sh);
  try { sh.hideColumn(sh.getRange(1, idx2['Task Details (HTML)'])); } catch(_){}
  try { sh.hideColumn(sh.getRange(1, idx2['Task Details (Draft)'])); } catch(_){}
}

function KNB_MTE_openEditorForActiveRow(){
  const sh=SpreadsheetApp.getActiveSheet();
  const rg=sh.getActiveRange();
  if (!rg) return SpreadsheetApp.getUi().alert('Select a data row first.');
  const row = rg.getRow(); if (row<2) return SpreadsheetApp.getUi().alert('Select a data row (≥2).');
  KNB_MTE_openEditorForRow_(row);
}

function KNB_MTE_onSelectionChange_(e){
  try{
    if(!e||!e.range) return;
    const sh=e.range.getSheet(); const row=e.range.getRow(); const col=e.range.getColumn();
    if (row<2) return;
    const idx = KNB_headerIndex_(sh);
    if (!idx['Task Details'] || col !== idx['Task Details']) return;
    const v = String(sh.getRange(row, col).getDisplayValue()||'').trim();
    if (v === KNB_MTE_ICON || v === '') KNB_MTE_openEditorForRow_(row);
  }catch(_){}
}

function KNB_MTE_onEditOpen_(e){
  try{
    const sh=e.range.getSheet(); const row=e.range.getRow(); const col=e.range.getColumn();
    if (row<2) return;
    const idx = KNB_headerIndex_(sh);
    if (!idx['Task Details'] || col !== idx['Task Details']) return;
    const oldVal = (e.oldValue||'').toString().trim();
    const newVal = (e.value||'').toString().trim();
    const shouldOpen = (oldVal === KNB_MTE_ICON || oldVal === '') || (newVal === KNB_MTE_ICON);
    if (!shouldOpen) return;
    sh.getRange(row, idx['Task Details']).setValue(oldVal); SpreadsheetApp.flush();
    KNB_MTE_openEditorForRow_(row);
  }catch(_){}
}

function KNB_MTE_openEditorForRow_(row){
  const sh  = SpreadsheetApp.getActiveSheet();
  KNB_MTE_ensureColumns_();                 // make sure storage columns exist
  const idx = KNB_headerIndex_(sh);

  // Read fields (OK if blank)
  const get = h => (idx[h] ? sh.getRange(row, idx[h]).getDisplayValue() : '');
  const html = idx['Task Details (HTML)'] ? String(sh.getRange(row, idx['Task Details (HTML)']).getValue()||'') : '';
  const draft= idx['Task Details (Draft)']? String(sh.getRange(row, idx['Task Details (Draft)']).getValue()||'') : '';
  const initial = draft || html || '';

  const tpl = HtmlService.createTemplateFromFile('views/TaskEditor'); // your taskeditor.html
  tpl.row   = row;
  tpl.tName = get('Task Name') || '';
  tpl.ass   = get('Assignee') || '';
  tpl.cl    = get('Client Name') || '';
  tpl.st    = get('Status') || '';
  tpl.html  = initial;
  SpreadsheetApp.getUi()
    .showModalDialog(tpl.evaluate().setWidth(860).setHeight(600), 'Task Details — Editor');
}

function KNB_MTE_save(row, html){
  const sh  = SpreadsheetApp.getActiveSheet();
  KNB_MTE_ensureColumns_();
  const idx = KNB_headerIndex_(sh);
  const v = String(html||'').trim();

  const NOTE_LIMIT = 48000;
  if (v.length > NOTE_LIMIT)
    throw new Error(`Content is too large (${v.length} chars). Keep under ${NOTE_LIMIT.toLocaleString()} characters.`);

  // Store HTML in hidden column; clear any tooltip
  if (!idx['Task Details (HTML)'] || !idx['Task Details'])
    throw new Error('Columns missing—run Ensure/Repair Columns.');
  sh.getRange(row, idx['Task Details (HTML)']).setValue(v);
  sh.getRange(row, idx['Task Details']).setValue(v ? KNB_MTE_ICON : '');
  try { sh.getRange(row, idx['Task Details']).setNote(''); } catch(_){}

  // Optional: wipe draft on final save
  if (idx['Task Details (Draft)']) sh.getRange(row, idx['Task Details (Draft)']).setValue('');
  return true;
}

function KNB_MTE_backfillIcons(){
  const sh=SpreadsheetApp.getActiveSheet();
  KNB_MTE_ensureColumns_();
  const idx=KNB_headerIndex_(sh);
  const last = sh.getLastRow(); if (last<2) return;
  const html = sh.getRange(2, idx['Task Details (HTML)'], last-1, 1).getValues();
  const td   = sh.getRange(2, idx['Task Details'],      last-1, 1).getDisplayValues();
  let upd=0;
  for(let i=0;i<html.length;i++){
    const has = String(html[i][0]||'').trim();
    const icon= String(td[i][0]||'').trim()===KNB_MTE_ICON;
    if (has && !icon){ sh.getRange(i+2, idx['Task Details']).setValue(KNB_MTE_ICON); upd++; }
  }
  SpreadsheetApp.getActive().toast(`Backfilled 📝 on ${upd} row(s).`, '📝', 4);
}
