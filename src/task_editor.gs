/***** RICH TEXT FIELDS CONFIG *****/
const KNB_RTE = {
  details: {
    key: 'details',
    label: 'Task Details',
    html:  'Task Details (HTML)',
    draft: 'Task Details (Draft)',
    icon:  '📝',
    dialogTitle: 'Task Details — Editor'
  },
  revision: {
    key: 'revision',
    label: 'Revision Notes',
    html:  'Revision Notes (HTML)',
    draft: 'Revision Notes (Draft)',
    icon:  '🧾', // feel free to change icon
    dialogTitle: 'Revision Notes — Editor'
  }
};

// =========================
// CORE (generic helpers)
// =========================
function KNB_RTE_ensureColumns_(field){
  const sh = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);

  if (!idx[field.label]) throw new Error(`Missing header "${field.label}".`);
  if (!idx[field.html] || !idx[field.draft]) {
    throw new Error(
      `Missing "${field.html}" and/or "${field.draft}". ` +
      `Ask an editor to add them immediately after "${field.label}" and keep them hidden.`
    );
  }
  // Keep hidden
  try { sh.hideColumn(sh.getRange(1, idx[field.html])); } catch(_){}
  try { sh.hideColumn(sh.getRange(1, idx[field.draft])); } catch(_){}
}


function KNB_RTE_openEditorForActiveRow(fieldKey){
  const field = KNB_RTE[fieldKey];
  const sh=SpreadsheetApp.getActiveSheet();
  const rg=sh.getActiveRange();
  if (!rg) return SpreadsheetApp.getUi().alert('Select a data row first.');
  const row = rg.getRow(); if (row<2) return SpreadsheetApp.getUi().alert('Select a data row (≥2).');
  KNB_RTE_openEditorForRow_(row, field);
}

// in taskeditor.gs
function KNB_RTE_openEditorForRow_(row, field){
  const sh  = SpreadsheetApp.getActiveSheet();
  KNB_RTE_ensureColumns_(field);
  const idx = KNB_headerIndex_(sh);

  const get = h => (idx[h] ? sh.getRange(row, idx[h]).getDisplayValue() : '');
  const html = idx[field.html]  ? String(sh.getRange(row, idx[field.html]).getValue()||'')   : '';
  const draft= idx[field.draft] ? String(sh.getRange(row, idx[field.draft]).getValue()||'')  : '';
  const initial = draft || html || '';

  // ⬇️ Fix: use the actual file name (no "views/")
  const tpl = HtmlService.createTemplateFromFile('views/TaskEditor');
  tpl.row   = row;
  tpl.tName = get('Task Name') || '';
  tpl.ass   = get('Assignee') || '';
  tpl.cl    = get('Client Name') || '';
  tpl.st    = get('Status') || '';
  tpl.html  = initial;

  // ⬇️ Pass field context
  tpl.dialogTitle = field.dialogTitle;
  tpl.fieldKey    = field.key;

  SpreadsheetApp.getUi()
    .showModalDialog(tpl.evaluate().setWidth(860).setHeight(600), field.dialogTitle);
}


function KNB_RTE_save(row, html, fieldKey){
  const field = KNB_RTE[fieldKey];
  const sh  = SpreadsheetApp.getActiveSheet();
  KNB_RTE_ensureColumns_(field);
  const idx = KNB_headerIndex_(sh);
  const v = String(html||'').trim();

  const NOTE_LIMIT = 48000;
  if (v.length > NOTE_LIMIT)
    throw new Error(`Content is too large (${v.length} chars). Keep under ${NOTE_LIMIT.toLocaleString()} characters.`);

  if (!idx[field.html] || !idx[field.label])
    throw new Error('Columns missing—run Ensure/Repair Columns.');

  // Persist and show icon (or blank)
  sh.getRange(row, idx[field.html]).setValue(v);
  sh.getRange(row, idx[field.label]).setValue(v ? field.icon : '');
  try { sh.getRange(row, idx[field.label]).setNote(''); } catch(_){}

  if (idx[field.draft]) sh.getRange(row, idx[field.draft]).setValue('');
  return true;
}

function KNB_RTE_backfillIcons(fieldKey){
  const field = KNB_RTE[fieldKey];
  const sh=SpreadsheetApp.getActiveSheet();
  KNB_RTE_ensureColumns_(field);
  const idx=KNB_headerIndex_(sh);
  const last = sh.getLastRow(); if (last<2) return;
  const html = sh.getRange(2, idx[field.html], last-1, 1).getValues();
  const vis  = sh.getRange(2, idx[field.label], last-1, 1).getDisplayValues();
  let upd=0;
  for(let i=0;i<html.length;i++){
    const has = String(html[i][0]||'').trim();
    const icon= String(vis[i][0]||'').trim()===field.icon;
    if (has && !icon){ sh.getRange(i+2, idx[field.label]).setValue(field.icon); upd++; }
  }
  SpreadsheetApp.getActive().toast(`Backfilled ${field.icon} on ${upd} row(s).`, field.label, 4);
}

// =========================
// LEGACY-COMPAT SHIMS (keep your old function names working)
// =========================
const KNB_MTE_ICON = KNB_RTE.details.icon; // keeps any external references safe

function KNB_MTE_openEditorForActiveRow(){ KNB_RTE_openEditorForActiveRow('details'); }
function KNB_MTE_openEditorForRow_(row){ KNB_RTE_openEditorForRow_(row, KNB_RTE.details); }
function KNB_MTE_save(row, html){ return KNB_RTE_save(row, html, 'details'); }
// function KNB_MTE_backfillIcons(){ KNB_RTE_backfillIcons('details'); }

// =========================
// EVENTS: open editor when user clicks into the icon cell
// =========================
function KNB_RTE_onSelectionChange_(e){
  try{
    if(!e||!e.range) return;
    const sh=e.range.getSheet(); const row=e.range.getRow(); const col=e.range.getColumn();
    if (row<2) return;
    const idx = KNB_headerIndex_(sh);

    for (const field of Object.values(KNB_RTE)){
      if (!idx[field.label] || col !== idx[field.label]) continue;
      const v = String(sh.getRange(row, col).getDisplayValue()||'').trim();
      if (v === field.icon || v === '') KNB_RTE_openEditorForRow_(row, field);
    }
  }catch(_){}
}

function KNB_RTE_onEditOpen_(e){
  try{
    const sh=e.range.getSheet(); const row=e.range.getRow(); const col=e.range.getColumn();
    if (row<2) return;
    const idx = KNB_headerIndex_(sh);

    for (const field of Object.values(KNB_RTE)){
      if (!idx[field.label] || col !== idx[field.label]) continue;
      const oldVal = (e.oldValue||'').toString().trim();
      const newVal = (e.value||'').toString().trim();
      const shouldOpen = (oldVal === field.icon || oldVal === '') || (newVal === field.icon);
      if (!shouldOpen) continue;
      sh.getRange(row, idx[field.label]).setValue(oldVal); SpreadsheetApp.flush();
      KNB_RTE_openEditorForRow_(row, field);
    }
  }catch(_){}
}
