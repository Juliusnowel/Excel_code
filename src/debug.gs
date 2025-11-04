/* =========================
   DEBUG
========================= */
function KNB_DIAG_whoAmI(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheets()[0];
  return {
    user: (Session.getActiveUser() && Session.getActiveUser().getEmail()) || '(blank)',
    file: ss.getName(),
    onSharedDrive: !!ss.getOwner() ? false : true,
    sheet: sh.getName(),
    gid: sh.getSheetId(),
    editors: ss.getEditors().map(e=>e.getEmail())
  };
}

function KNB_DIAG_probeSimpleWrite(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheets().find(x => x.getSheetId() === Number(KNB_CFG.GID.REQUESTED))
         || ss.getSheets().find(x => x.getName().trim().toLowerCase() === 'requested');
  if (!sh) throw new Error('No Requested sheet found (by name or gid).');

  const row = Math.max(2, sh.getLastRow()+1);
  const col = 1; // Column A only
  const rg  = sh.getRange(row, col);
  const v   = rg.getValue();        // read
  rg.setValue(v || 'probe');        // minimal write
  return {sheet: sh.getName(), row, col, wrote: rg.getValue()};
}

function KNB_TASK_diagnoseEnvironment_() {
  const s = SpreadsheetApp.getActive();
  const reqGid = Number(KNB_CFG.GID.REQUESTED);
  const byGid  = s.getSheets().find(x => x.getSheetId() === reqGid);
  const byName = s.getSheets().find(x => String(x.getName()).trim().toLowerCase() === 'requested');
  const sh = byGid || byName || s.getActiveSheet();

  const where = (byGid)
    ? `OK: Using Requested gid ${reqGid} on sheet "${sh.getName()}".`
    : byName ? `Using sheet named "Requested" (gid ${sh.getSheetId()}).`
             : `Using ACTIVE sheet "${sh.getName()}" (gid ${sh.getSheetId()}).`;

  const idx = KNB_headerIndex_(sh);
  const needCols = [
    KNB_CFG.COL.DEPARTMENT,
    KNB_CFG.COL.ASSIGNEE,
    KNB_CFG.COL.CLIENT,
    KNB_CFG.COL.TASK,
    KNB_CFG.COL.PRIORITY,
    KNB_CFG.COL.CREATED,
    KNB_CFG.COL.STATUS,
    KNB_CFG.COL.DETAILS
  ];
  const missing = needCols.filter(h => !idx[h]);
  const hdrs = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0];

  const msg = [
    where,
    missing.length ? ('MISSING headers: ' + missing.join(', ')) : 'All required headers present ✅',
    'Detected headers: ' + hdrs.join(' | ')
  ].join('\n\n');

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

function KNB_debugTriggers(){
  const list = ScriptApp.getProjectTriggers().map(t => (t.getHandlerFunction?t.getHandlerFunction():'?')+' — '+t.getEventType()).join('\n');
  Logger.log(list||'(no triggers)');
  SpreadsheetApp.getUi().alert('See Apps Script → View → Logs for trigger list.');
}

function KNB_cleanupNotesHere(){
  const sh = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const c = idx['Task Details']; if (!c) return;
  const last = sh.getLastRow();
  if (last < 2) return;
  const rg = sh.getRange(2, c, last-1, 1);
  const notes = rg.getNotes();
  let dirty = false;
  for (let i=0;i<notes.length;i++){
    if (String(notes[i][0]||'').length){
      notes[i][0] = ''; dirty = true;
    }
  }
  if (dirty) rg.setNotes(notes);
  SpreadsheetApp.getActive().toast('Cleared Task Details tooltips on this sheet.', 'Cleanup', 4);
}

// Ensure Task Details storage columns exist (right after "Task Details") and hide them
function KNB_TASK_ensureDetailsColumns_Here(sh){
  if (!sh) return;
  const idx = KNB_headerIndex_(sh);
  const td = idx['Task Details'];
  if (!td) return; // sheet doesn't have Task Details; nothing to do

  // Make/locate "Task Details (HTML)"
  let htmlCol = idx['Task Details (HTML)'];
  if (!htmlCol){
    sh.insertColumnAfter(td);
    htmlCol = td + 1;
    sh.getRange(1, htmlCol).setValue('Task Details (HTML)');
  }

  // Re-read to get fresh indices
  const idx2 = KNB_headerIndex_(sh);

  // Make/locate "Task Details (Draft)"
  let draftCol = idx2['Task Details (Draft)'];
  if (!draftCol){
    sh.insertColumnAfter(htmlCol);
    draftCol = htmlCol + 1;
    sh.getRange(1, draftCol).setValue('Task Details (Draft)');
  }

  // Hide both storage columns
  try { sh.hideColumn(sh.getRange(1, htmlCol)); } catch(_){}
  try { sh.hideColumn(sh.getRange(1, draftCol)); } catch(_){}
}

// Run across all boards
// function KNB_TASK_ensureDetailsColumns_AllBoards(){
//   const ss = SpreadsheetApp.getActive();
//   (KNB_allGids_() || []).forEach(gid => {
//     const sh = KNB_sheetById_(gid);
//     if (sh) KNB_TASK_ensureDetailsColumns_Here(sh);
//   });
//   SpreadsheetApp.getActive().toast('Ensured + hid Task Details storage on all boards', '📝', 4);
// }

/* =========================
   DROPDOWNS / VALIDATION
========================= */

function KNB_applyAssigneeDropdownHere(){
  const sh  = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const c   = idx[KNB_CFG.COL.ASSIGNEE]; 
  if (!c) return SpreadsheetApp.getUi().alert('No "Assignee" column found on this sheet.');
  const rows = Math.max(1, sh.getMaxRows() - 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KNB_CFG.ASSIGNEES || [], true) // show dropdown
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, c, rows, 1).setDataValidation(rule);
  SpreadsheetApp.getActive().toast('Assignee dropdown applied.', 'Tasks', 3);
}

function KNB_applyClientDropdownHere(){
  const sh  = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const c   = idx[KNB_CFG.COL.CLIENT]; 
  if (!c) return SpreadsheetApp.getUi().alert('No "Client Name" column found on this sheet.');
  const values = KNB_CFG.CLIENTS || [];
  if (!values.length) return SpreadsheetApp.getUi().alert('KNB_CFG.CLIENTS is empty.');

  const rows = Math.max(1, sh.getMaxRows() - 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  sh.getRange(2, c, rows, 1).setDataValidation(rule);
  SpreadsheetApp.getActive().toast('Client Name dropdown applied.', 'Tasks', 3);
}

function KNB_applyPriorityDropdownHere(){
  const sh  = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const c   = idx[KNB_CFG.COL.PRIORITY];
  if (!c) return SpreadsheetApp.getUi().alert('No "Task Priority" column found on this sheet.');
  const rows = Math.max(1, sh.getMaxRows() - 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KNB_CFG.PRIORITIES || [], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, c, rows, 1).setDataValidation(rule);
  SpreadsheetApp.getActive().toast('Task Priority dropdown applied.', 'Tasks', 3);
}

// All boards (Assignee + Priority)
function KNB_applyAssigneeAndPriority_AllBoards(){
  (KNB_allGids_() || []).forEach(gid => {
    const sh = KNB_sheetById_(gid); if (!sh) return;
    const idx = KNB_headerIndex_(sh);
    const rows = Math.max(1, sh.getMaxRows() - 1);

    // Assignee
    if (idx[KNB_CFG.COL.ASSIGNEE]) {
      const r1 = SpreadsheetApp.newDataValidation()
        .requireValueInList(KNB_CFG.ASSIGNEES || [], true).setAllowInvalid(false).build();
      sh.getRange(2, idx[KNB_CFG.COL.ASSIGNEE], rows, 1).setDataValidation(r1);
    }
    // Priority
    if (idx[KNB_CFG.COL.PRIORITY]) {
      const r2 = SpreadsheetApp.newDataValidation()
        .requireValueInList(KNB_CFG.PRIORITIES || [], true).setAllowInvalid(false).build();
      sh.getRange(2, idx[KNB_CFG.COL.PRIORITY], rows, 1).setDataValidation(r2);
    }
  });
  SpreadsheetApp.getActive().toast('Assignee & Priority dropdowns applied on all boards.', 'Tasks', 4);
}

// helpers
// function KNB_installOwnerColumnAll(){
//   KNB_allGids_().forEach(gid=>{
//     const sh = KNB_sheetById_(gid); if(!sh) return;
//     const idx = KNB_headerIndex_(sh);
//     if (idx[KNB_CFG.COL.OWNER]) return;              
//     const cAss = idx[KNB_CFG.COL.ASSIGNEE] || 0;
//     const cDept= idx[KNB_CFG.COL.DEPARTMENT] || 0;
//     const insertAt = cAss ? cAss : (cDept ? cDept+1 : 2);
//     sh.insertColumnBefore(insertAt);
//     sh.getRange(1, insertAt).setValue(KNB_CFG.COL.OWNER);
//   });
//   SpreadsheetApp.getActive().toast('Owner column installed portfolio-wide.', 'Setup', 4);
// }

// function KNB_applyOwnerDropdown_AllBoards(){
//   (KNB_allGids_() || []).forEach(gid=>{
//     const sh = KNB_sheetById_(gid); if(!sh) return;
//     SpreadsheetApp.setActiveSheet(sh);
//     try { KNB_applyOwnerDropdownHere(); } catch(_) {}
//   });
//   SpreadsheetApp.getActive().toast('Owner dropdown applied on all boards.', 'Tasks', 4);
// }

// Refresh Priority DV + chips on ALL boards
// function KNB_refreshPriority_AllBoards(){
//   (KNB_allGids_() || []).forEach(gid => {
//     const sh = KNB_sheetById_(gid); if (!sh) return;
//     const idx = KNB_headerIndex_(sh);
//     const cPrio = idx[KNB_CFG.COL.PRIORITY]; if (!cPrio) return;

//     // 1) Re-apply dropdown (now includes 'Urgent')
//     const rows = Math.max(0, sh.getLastRow() - 1);
//     if (rows > 0){
//       const dv = SpreadsheetApp.newDataValidation()
//         .requireValueInList(KNB_CFG.PRIORITIES || [], true)
//         .setAllowInvalid(false)
//         .setHelpText('Choose a Priority')
//         .build();
//       sh.getRange(2, cPrio, rows, 1).setDataValidation(dv);
//     }

//     // 2) Rebuild ONLY the priority chip rules, keep all other CF rules
//     const rules = sh.getConditionalFormatRules();
//     const keep  = rules.filter(r => !r.getRanges().some(rg => rg.getColumn() === cPrio && rg.getNumColumns() === 1));

//     const a1 = KNB_colToA1_(cPrio);
//     const range = sh.getRange(2, cPrio, Math.max(1, sh.getLastRow()-1), 1);
//     const prioRules = Object.entries(KNB_CFG.PRIORITY_COLORS || {}).map(([label, bg]) =>
//       SpreadsheetApp.newConditionalFormatRule()
//         .whenFormulaSatisfied(KNB_cfMatch_(a1, label))
//         .setBackground(bg)
//         .setFontColor(KNB_pickTextColor_(bg))
//         .setRanges([range])
//         .build()
//     );

//     sh.setConditionalFormatRules([...keep, ...prioRules]);
//   });

//   SpreadsheetApp.getActive().toast('Priority dropdowns + chips refreshed (all boards).', 'Tasks', 4);
// }

function KNB_applyPriorityDropdownOnSheet_(sh){
  const idx = KNB_headerIndex_(sh);
  const c   = idx[KNB_CFG.COL.PRIORITY]; 
  if (!c) return;
  const rows = Math.max(1, sh.getMaxRows() - 1);
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(KNB_CFG.PRIORITIES || [], true)
    .setAllowInvalid(false)
    .setHelpText('Choose a Priority')
    .build();
  sh.getRange(2, c, rows, 1).setDataValidation(dv);
}

// function KNB_removeLateValidationHere(){
//   const sh = SpreadsheetApp.getActiveSheet();
//   const idx = KNB_headerIndex_(sh);
//   if (!idx || !idx['Late or not?']) {
//     SpreadsheetApp.getActive().toast('Header "Late or not?" not found on this sheet.', 'Late Fix', 4);
//     return;
//   }
//   const c = idx['Late or not?'];
//   sh.getRange(2, c, sh.getMaxRows() - 1, 1).clearDataValidations();
//   try { sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('0'); } catch(_){}
//   SpreadsheetApp.getActive().toast('Removed Late/Not? validation on "' + sh.getName() + '"', 'Late Fix', 3);
// }

function KNB_TRIG_reset(){
  // Purge all onEdit handlers in this project
  ScriptApp.getProjectTriggers().forEach(t=>{
    const fn = t.getHandlerFunction && t.getHandlerFunction();
    if (fn && /onedit/i.test(fn)) ScriptApp.deleteTrigger(t);
  });
  // Reinstall our single entry point
  ScriptApp.newTrigger('KNB_onEdit_')
    .forSpreadsheet(SpreadsheetApp.getActive().getId())
    .onEdit()
    .create();
  SpreadsheetApp.getActive().toast('onEdit trigger reset → KNB_onEdit_.', 'Triggers', 4);
}

function KNB_AUDIT_backfillFreeze(){
  (KNB_allGids_() || []).forEach(gid=>{
    const sh = KNB_sheetById_(gid); if (!sh) return;
    const idx = KNB_headerIndex_(sh);
    const cSts = idx[KNB_CFG.COL.STATUS]; if (!cSts) return;
    const cFrz = KNB_ensureFreezeColumn_(sh, idx);

    const last = Math.max(2, sh.getLastRow());
    if (last < 2) return;

    const sts = sh.getRange(2, cSts, last-1, 1).getDisplayValues();
    const frz = sh.getRange(2, cFrz, last-1, 1).getValues();

    let dirty = false;
    for (let i=0;i<sts.length;i++){
      if (String(sts[i][0]).trim() === 'For Approval' && !frz[i][0]) {
        frz[i][0] = new Date();            // stamp once
        dirty = true;
      }
    }
    if (dirty){
      sh.getRange(2, cFrz, frz.length, 1).setValues(frz);
      try { sh.getRange(2, cFrz, frz.length, 1).setNumberFormat('yyyy-mm-dd'); } catch(_){}
    }
  });
  SpreadsheetApp.getActive().toast('Freeze dates backfilled for For Approval rows.', 'SLA', 4);
}

function KNB_SMOKE_freezeOnStatus(){
  const sh = SpreadsheetApp.getActiveSheet();
  const idx = KNB_headerIndex_(sh);
  const row = SpreadsheetApp.getActiveRange().getRow();
  const col = idx[KNB_CFG.COL.STATUS];
  if (!col || row < 2) throw new Error('Select a data row with a Status cell.');
  const range = sh.getRange(row, col);
  const oldValue = range.getDisplayValue();
  range.setValue('For Approval');                
  KNB_onEdit_({range, value:'For Approval', oldValue}); 
}
